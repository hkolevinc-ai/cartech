# CarTechBG catalog scraper

Scraper за `https://cartechbg.com/`, направен за GitHub Actions автоматизация.

## Какво извлича

- Код на ниво продукт / product SKU
- Код на ниво вариация / variation SKU, ако продуктът има вариации
- Име на продукт
- Категория
- Подкатегория
- Път на категориите
- Цена в евро
- URL на всички изображения
- Описание на продукта
- URL на продукта
- Отделни колони `Изображение 1`, `Изображение 2`, ...

## Как работи

1. Проверява sitemap-и:
   - `/sitemap.xml`
   - `/sitemap_index.xml`
   - `/product-sitemap.xml`
   - `/product-sitemap2.xml`, `/product-sitemap3.xml`, ...
   - WordPress sitemap формат `/wp-sitemap-posts-product-1.xml`
2. Събира всички `/product/...` URL-и.
3. Ако sitemap не върне продукти, обхожда `/shop/` pagination като fallback.
4. За всеки продукт извлича product SKU и, при вариативни продукти, variation SKU от WooCommerce `data-product_variations`.
5. Записва резултата в `cartechbg_all_products.xlsx`.

## Колони в Excel файла

1. `Код на ниво продукт`
2. `Код на ниво вариация`
3. `Име на продукт`
4. `Категория`
5. `Подкатегория`
6. `Път на категориите`
7. `Цена в евро`
8. `URL на всички изображения`
9. `Описание на продукта`
10. `URL на продукта`
11. `Изображение 1`, `Изображение 2`, ...

Ако има повече от една вариация, кодовете на вариациите се записват в една клетка, разделени с ` | `.

## Как се пуска в GitHub

Качи/замени тези файлове в repository-то:

- `scraper.py`
- `requirements.txt`
- `.github/workflows/scrape.yml`

После отвори:

`Actions` → `Scrape CarTechBG catalog` → `Run workflow`

След приключване Excel файлът ще е в `Artifacts` като `cartechbg_all_products`.

## Настройки

Могат да се променят от workflow файла или като environment variables:

- `BASE_URL` — по подразбиране `https://cartechbg.com/`
- `OUTPUT_FILE` — по подразбиране `cartechbg_all_products.xlsx`
- `REQUEST_SLEEP` — пауза между продуктите, по подразбиране `0.4` секунди
- `MAX_PRODUCTS` — лимит за тест, `0` означава без лимит
- `MAX_SHOP_PAGES` — максимум страници за fallback crawler

## Локален тест

```bash
pip install -r requirements.txt
MAX_PRODUCTS=5 python scraper.py
```
