import os
import re
import json
import html
import time
import math
import logging
from dataclasses import dataclass, asdict
from typing import Dict, Iterable, List, Optional, Set, Tuple
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_URL = os.getenv("BASE_URL", "https://cartechbg.com/").rstrip("/") + "/"
OUTPUT_FILE = os.getenv("OUTPUT_FILE", "cartechbg_all_products.xlsx")
REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "30"))
REQUEST_SLEEP = float(os.getenv("REQUEST_SLEEP", "0.4"))
MAX_PRODUCTS = int(os.getenv("MAX_PRODUCTS", "0"))  # 0 = no limit
MAX_SHOP_PAGES = int(os.getenv("MAX_SHOP_PAGES", "200"))

HEADERS = {
    "User-Agent": os.getenv(
        "USER_AGENT",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/125.0 Safari/537.36 CarTechBGScraper/1.0",
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.8",
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
log = logging.getLogger("cartechbg-scraper")

session = requests.Session()
session.headers.update(HEADERS)


@dataclass
class Product:
    product_level_code: str = ""
    variation_level_code: str = ""
    product_name: str = ""
    category: str = ""
    subcategory: str = ""
    category_path: str = ""
    price_eur: str = ""
    all_image_urls: str = ""
    description: str = ""
    product_url: str = ""


def fetch(url: str, accept_xml: bool = False) -> Optional[str]:
    """Fetch URL with retries. Returns text or None."""
    for attempt in range(1, 4):
        try:
            headers = dict(HEADERS)
            if accept_xml:
                headers["Accept"] = "application/xml,text/xml,*/*;q=0.8"
            r = session.get(url, timeout=REQUEST_TIMEOUT, headers=headers)
            if r.status_code in (403, 429, 500, 502, 503, 504):
                raise requests.HTTPError(f"HTTP {r.status_code}")
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r.text
        except Exception as exc:
            wait = attempt * 2
            log.warning("Fetch failed (%s/%s): %s | %s", attempt, 3, url, exc)
            if attempt < 3:
                time.sleep(wait)
    return None


def normalize_url(url: str) -> str:
    url = url.strip()
    if not url:
        return ""
    return urljoin(BASE_URL, url).split("#")[0]


def is_product_url(url: str) -> bool:
    parsed = urlparse(url)
    if not parsed.netloc.endswith("cartechbg.com"):
        return False
    path = parsed.path.rstrip("/") + "/"
    return "/product/" in path and "/product-category/" not in path


def parse_xml_locs(xml_text: str) -> List[str]:
    soup = BeautifulSoup(xml_text, "xml")
    return [loc.get_text(strip=True) for loc in soup.find_all("loc") if loc.get_text(strip=True)]


def collect_sitemap_product_urls() -> List[str]:
    """Collect product URLs from common WordPress/WooCommerce sitemap locations."""
    candidates = [
        urljoin(BASE_URL, "sitemap.xml"),
        urljoin(BASE_URL, "sitemap_index.xml"),
        urljoin(BASE_URL, "product-sitemap.xml"),
        urljoin(BASE_URL, "product-sitemap1.xml"),
        urljoin(BASE_URL, "product-sitemap2.xml"),
        urljoin(BASE_URL, "wp-sitemap-posts-product-1.xml"),
    ]

    sitemap_urls: Set[str] = set()
    product_urls: Set[str] = set()

    for sitemap in candidates:
        text = fetch(sitemap, accept_xml=True)
        if not text:
            continue
        locs = parse_xml_locs(text)
        for loc in locs:
            l = loc.lower()
            if is_product_url(loc):
                product_urls.add(normalize_url(loc))
            elif any(token in l for token in ["product-sitemap", "wp-sitemap-posts-product"]):
                sitemap_urls.add(normalize_url(loc))

    # Also probe product-sitemapN.xml until several misses in a row.
    misses = 0
    for i in range(1, 21):
        suffix = "" if i == 1 else str(i)
        probe = urljoin(BASE_URL, f"product-sitemap{suffix}.xml")
        if probe in sitemap_urls:
            continue
        text = fetch(probe, accept_xml=True)
        if not text:
            misses += 1
            if misses >= 3:
                break
            continue
        misses = 0
        locs = parse_xml_locs(text)
        found_products = [normalize_url(loc) for loc in locs if is_product_url(loc)]
        if found_products:
            product_urls.update(found_products)
        else:
            sitemap_urls.add(probe)

    for sitemap in sorted(sitemap_urls):
        text = fetch(sitemap, accept_xml=True)
        if not text:
            continue
        locs = parse_xml_locs(text)
        for loc in locs:
            if is_product_url(loc):
                product_urls.add(normalize_url(loc))

    urls = sorted(product_urls)
    log.info("Product URLs from sitemap: %s", len(urls))
    return urls


def collect_shop_product_urls() -> List[str]:
    """Fallback crawler through shop pagination."""
    product_urls: Set[str] = set()
    stop_after_empty = 0

    patterns = [
        urljoin(BASE_URL, "shop/"),
        urljoin(BASE_URL, "shop/?product-page={page}"),
        urljoin(BASE_URL, "shop/page/{page}/"),
    ]

    for page in range(1, MAX_SHOP_PAGES + 1):
        page_urls = []
        if page == 1:
            page_urls.append(patterns[0])
        page_urls.append(patterns[1].format(page=page))
        page_urls.append(patterns[2].format(page=page))

        before = len(product_urls)
        for page_url in page_urls:
            html = fetch(page_url)
            if not html:
                continue
            soup = BeautifulSoup(html, "lxml")
            for a in soup.select("a[href]"):
                href = normalize_url(a.get("href", ""))
                if is_product_url(href):
                    product_urls.add(href)
        gained = len(product_urls) - before
        log.info("Shop page %s gained %s product URLs; total %s", page, gained, len(product_urls))
        if gained == 0:
            stop_after_empty += 1
            if stop_after_empty >= 5:
                break
        else:
            stop_after_empty = 0
        time.sleep(REQUEST_SLEEP)

    urls = sorted(product_urls)
    log.info("Product URLs from shop fallback: %s", len(urls))
    return urls


def clean_text(text: str) -> str:
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def parse_price_eur(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    # Prefer values followed by euro sign or EUR.
    m = re.search(r"(\d+[\d\s,.]*)\s*(?:€|EUR|евро)", text, flags=re.I)
    if not m:
        return ""
    raw = m.group(1).replace(" ", "")
    # Convert European decimal comma to dot when needed.
    if raw.count(",") == 1 and raw.count(".") == 0:
        raw = raw.replace(",", ".")
    # Remove thousands separators safely.
    if raw.count(".") > 1:
        parts = raw.split(".")
        raw = "".join(parts[:-1]) + "." + parts[-1]
    return raw


def extract_json_ld(soup: BeautifulSoup) -> List[dict]:
    items = []
    for script in soup.find_all("script", type=lambda v: v and "ld+json" in v):
        raw = script.string or script.get_text() or ""
        raw = raw.strip()
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except Exception:
            continue
        stack = data if isinstance(data, list) else [data]
        while stack:
            obj = stack.pop(0)
            if isinstance(obj, dict):
                items.append(obj)
                graph = obj.get("@graph")
                if isinstance(graph, list):
                    stack.extend(graph)
            elif isinstance(obj, list):
                stack.extend(obj)
    return items


def first_product_json_ld(soup: BeautifulSoup) -> dict:
    for item in extract_json_ld(soup):
        typ = item.get("@type")
        types = typ if isinstance(typ, list) else [typ]
        if any(str(t).lower() == "product" for t in types):
            return item
    return {}


def unique_keep_order(values: Iterable[str]) -> List[str]:
    seen = set()
    out = []
    for v in values:
        v = normalize_url(str(v))
        if not v or v in seen:
            continue
        seen.add(v)
        out.append(v)
    return out


def unique_text_keep_order(values: Iterable[str]) -> List[str]:
    """Deduplicate plain text values without treating them as URLs."""
    seen = set()
    out = []
    for v in values:
        v = clean_text(str(v))
        if not v:
            continue
        key = v.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(v)
    return out


def parse_category_parts(soup: BeautifulSoup, product_json: dict) -> Tuple[str, str, str]:
    """Return main category, subcategory and full category path."""
    ignore = {"начало", "home", "shop", "магазин", "всички продукти"}

    # WooCommerce product metadata usually contains the real product categories.
    cats = [clean_text(a.get_text(" ", strip=True)) for a in soup.select(".posted_in a")]
    cats = [c for c in cats if c and c.casefold() not in ignore]

    # Breadcrumb is useful when .posted_in is missing and usually preserves hierarchy.
    if not cats:
        cats = [clean_text(a.get_text(" ", strip=True)) for a in soup.select("nav.woocommerce-breadcrumb a, .breadcrumb a")]
        cats = [c for c in cats if c and c.casefold() not in ignore]

    if not cats and product_json.get("category"):
        raw = clean_text(str(product_json.get("category")))
        if raw:
            # Accept either a single category or a delimited path.
            split = re.split(r"\s*(?:>|/|»|›)\s*", raw)
            cats = [c for c in split if c and c.casefold() not in ignore]

    cats = unique_text_keep_order(cats)
    main_category = cats[0] if cats else ""
    subcategory = cats[1] if len(cats) > 1 else ""
    category_path = " > ".join(cats)
    return main_category, subcategory, category_path


def extract_variation_skus(soup: BeautifulSoup, html_text: str) -> str:
    """Extract SKU/code values from WooCommerce product variations."""
    skus: List[str] = []

    for form in soup.select("form.variations_form[data-product_variations]"):
        raw = form.get("data-product_variations") or ""
        raw = html.unescape(raw).strip()
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except Exception:
            data = None
        if isinstance(data, list):
            for item in data:
                if isinstance(item, dict):
                    sku = clean_text(str(item.get("sku", "") or ""))
                    if sku:
                        skus.append(sku)

    # Fallback for themes that print variation JSON inside scripts.
    for m in re.finditer(r'"sku"\s*:\s*"([^"\\]*(?:\\.[^"\\]*)*)"', html_text):
        try:
            sku = json.loads('"' + m.group(1) + '"')
        except Exception:
            sku = m.group(1)
        sku = clean_text(sku)
        if sku:
            skus.append(sku)

    return " | ".join(unique_text_keep_order(skus))


def parse_product(url: str) -> Product:
    html = fetch(url)
    if not html:
        return Product(product_url=url)

    soup = BeautifulSoup(html, "lxml")
    product_json = first_product_json_ld(soup)

    def select_text(selector: str) -> str:
        el = soup.select_one(selector)
        return clean_text(el.get_text(" ", strip=True)) if el else ""

    name = clean_text(product_json.get("name", ""))
    if not name:
        name = select_text("h1.product_title") or select_text("h1.entry-title")
    if not name:
        og = soup.select_one('meta[property="og:title"]')
        name = clean_text(og.get("content", "")) if og else ""

    sku = clean_text(str(product_json.get("sku", "") or ""))
    if not sku:
        sku = select_text(".sku")
    if not sku:
        sku_meta = soup.select_one('[itemprop="sku"], meta[property="product:retailer_item_id"]')
        if sku_meta:
            sku = clean_text(sku_meta.get("content") or sku_meta.get_text(" ", strip=True))

    category, subcategory, category_path = parse_category_parts(soup, product_json)
    variation_skus = extract_variation_skus(soup, html)

    price = ""
    offers = product_json.get("offers") if isinstance(product_json, dict) else None
    if isinstance(offers, list) and offers:
        offers = offers[0]
    if isinstance(offers, dict):
        price = clean_text(str(offers.get("price", "") or ""))
    if not price:
        price = parse_price_eur(select_text(".summary .price") or select_text(".price") or soup.get_text(" ", strip=True))

    desc_parts = []
    if product_json.get("description"):
        desc_parts.append(clean_text(str(product_json.get("description"))))
    for selector in [".woocommerce-product-details__short-description", "#tab-description", ".product .summary", ".entry-content"]:
        txt = select_text(selector)
        if txt and txt not in desc_parts:
            desc_parts.append(txt)
    description = "\n\n".join(desc_parts)

    images = []
    img_field = product_json.get("image")
    if isinstance(img_field, str):
        images.append(img_field)
    elif isinstance(img_field, list):
        images.extend([str(x) for x in img_field])
    for meta_sel in ['meta[property="og:image"]', 'meta[name="twitter:image"]']:
        for meta in soup.select(meta_sel):
            if meta.get("content"):
                images.append(meta["content"])
    for el in soup.select(".woocommerce-product-gallery a[href], .woocommerce-product-gallery img, .product img"):
        for attr in ["href", "data-large_image", "data-src", "src"]:
            val = el.get(attr)
            if val:
                images.append(val)
    images = [i for i in unique_keep_order(images) if not i.startswith("data:")]

    return Product(
        product_level_code=sku,
        variation_level_code=variation_skus,
        product_name=name,
        category=category,
        subcategory=subcategory,
        category_path=category_path,
        price_eur=price,
        all_image_urls=" | ".join(images),
        description=description,
        product_url=url,
    )


def write_xlsx(products: List[Product], output_file: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    max_imgs = 0
    image_lists = []
    for p in products:
        imgs = [x.strip() for x in p.all_image_urls.split("|") if x.strip()]
        image_lists.append(imgs)
        max_imgs = max(max_imgs, len(imgs))

    headers = [
        "Код на ниво продукт",
        "Код на ниво вариация",
        "Име на продукт",
        "Категория",
        "Подкатегория",
        "Път на категориите",
        "Цена в евро",
        "URL на всички изображения",
        "Описание на продукта",
        "URL на продукта",
    ] + [f"Изображение {i}" for i in range(1, max_imgs + 1)]
    ws.append(headers)

    for p, imgs in zip(products, image_lists):
        ws.append([
            p.product_level_code,
            p.variation_level_code,
            p.product_name,
            p.category,
            p.subcategory,
            p.category_path,
            p.price_eur,
            p.all_image_urls,
            p.description,
            p.product_url,
            *imgs,
        ])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        1: 22,
        2: 28,
        3: 55,
        4: 32,
        5: 32,
        6: 50,
        7: 14,
        8: 70,
        9: 90,
        10: 55,
    }
    for col_idx in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = widths.get(col_idx, 45)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 70

    wb.save(output_file)


def main() -> None:
    log.info("Starting CarTechBG scraper for %s", BASE_URL)
    urls = collect_sitemap_product_urls()
    if not urls:
        log.warning("No product URLs from sitemap. Falling back to /shop/ crawling.")
        urls = collect_shop_product_urls()

    # Safety: merge sitemap and fallback if sitemap seems too small.
    if len(urls) < 20:
        urls = sorted(set(urls) | set(collect_shop_product_urls()))

    if MAX_PRODUCTS > 0:
        urls = urls[:MAX_PRODUCTS]

    log.info("Total product URLs to scrape: %s", len(urls))
    products: List[Product] = []
    for idx, url in enumerate(urls, 1):
        log.info("[%s/%s] %s", idx, len(urls), url)
        try:
            product = parse_product(url)
            products.append(product)
        except Exception as exc:
            log.exception("Failed parsing %s: %s", url, exc)
            products.append(Product(product_url=url))
        time.sleep(REQUEST_SLEEP)

    write_xlsx(products, OUTPUT_FILE)
    log.info("Done. Wrote %s products to %s", len(products), OUTPUT_FILE)


if __name__ == "__main__":
    main()
