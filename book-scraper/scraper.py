"""Парсер каталога книг (учебный сайт books.toscrape.com) с выгрузкой в Excel или CSV.

Примеры запуска:
    python scraper.py
    python scraper.py --pages 10 --output books.xlsx
    python scraper.py --pages 5 --min-rating 4 --max-price 30 --output cheap_good.csv
"""
import argparse
import csv
import logging
import re
import sys
import time
from dataclasses import astuple, dataclass
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

PAGE_URL = "https://books.toscrape.com/catalogue/page-{}.html"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; learning-scraper/1.0)"}
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
COLUMNS = ["Название", "Цена, £", "Рейтинг (1-5)", "Наличие", "Ссылка"]


@dataclass
class Book:
    title: str
    price: float
    rating: int
    availability: str
    url: str


def fetch_page(session: requests.Session, url: str, retries: int = 3) -> str | None:
    """Скачивает страницу. Возвращает None, если страницы нет (404)."""
    for attempt in range(1, retries + 1):
        try:
            response = session.get(url, headers=HEADERS, timeout=15)
            if response.status_code == 404:
                return None
            response.raise_for_status()
            response.encoding = "utf-8"  # иначе символ £ превращается в «Â£»
            return response.text
        except requests.RequestException as error:
            logging.warning("Попытка %d из %d не удалась: %s", attempt, retries, error)
            time.sleep(2 * attempt)
    raise RuntimeError(f"Не удалось загрузить страницу: {url}")


def parse_books(html: str, page_url: str) -> list[Book]:
    """Достаёт книги из HTML одной страницы каталога."""
    soup = BeautifulSoup(html, "html.parser")
    books = []
    for card in soup.select("article.product_pod"):
        link = card.select_one("h3 a")
        price_tag = card.select_one("p.price_color")
        if link is None or price_tag is None:
            continue  # битая карточка: пропускаем, а не роняем весь парсер

        price = float(re.sub(r"[^\d.]", "", price_tag.get_text()))

        rating = 0
        rating_tag = card.select_one("p.star-rating")
        if rating_tag is not None:
            for css_class in rating_tag.get("class", []):
                if css_class in RATINGS:
                    rating = RATINGS[css_class]

        availability_tag = card.select_one("p.availability")
        availability = availability_tag.get_text(strip=True) if availability_tag else ""

        books.append(
            Book(
                title=link.get("title", link.get_text(strip=True)),
                price=price,
                rating=rating,
                availability=availability,
                url=urljoin(page_url, link["href"]),
            )
        )
    return books


def filter_books(books: list[Book], min_rating: int, max_price: float | None) -> list[Book]:
    result = [b for b in books if b.rating >= min_rating]
    if max_price is not None:
        result = [b for b in result if b.price <= max_price]
    return result


def save_excel(books: list[Book], path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Книги"
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for book in books:
        ws.append(list(astuple(book)))

    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "0.00"

    widths = [50, 12, 14, 14, 70]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def save_csv(books: list[Book], path: str) -> None:
    # utf-8-sig, чтобы Excel правильно показал кириллицу и символ £
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(COLUMNS)
        for book in books:
            writer.writerow(astuple(book))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Парсер каталога книг в Excel/CSV")
    parser.add_argument("--pages", type=int, default=5, help="сколько страниц каталога собрать (по умолчанию 5)")
    parser.add_argument("--output", default="books.xlsx", help="файл результата: .xlsx или .csv")
    parser.add_argument("--min-rating", type=int, default=0, choices=range(0, 6), help="минимальный рейтинг 0-5")
    parser.add_argument("--max-price", type=float, default=None, help="максимальная цена в £")
    parser.add_argument("--delay", type=float, default=0.5, help="пауза между запросами, секунд")
    return parser.parse_args()


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    args = parse_args()

    if not args.output.lower().endswith((".xlsx", ".csv")):
        sys.exit("Файл результата должен заканчиваться на .xlsx или .csv")
    if args.pages < 1:
        sys.exit("--pages должно быть не меньше 1")

    all_books: list[Book] = []
    with requests.Session() as session:
        for page in range(1, args.pages + 1):
            url = PAGE_URL.format(page)
            html = fetch_page(session, url)
            if html is None:
                logging.info("Страницы %d нет, каталог закончился.", page)
                break
            found = parse_books(html, url)
            all_books.extend(found)
            logging.info("Страница %d: найдено %d книг", page, len(found))
            time.sleep(args.delay)  # не нагружаем сайт

    books = filter_books(all_books, args.min_rating, args.max_price)
    if not books:
        sys.exit("Ничего не найдено. Проверьте фильтры или подключение к интернету.")

    if args.output.lower().endswith(".xlsx"):
        save_excel(books, args.output)
    else:
        save_csv(books, args.output)

    average = sum(b.price for b in books) / len(books)
    logging.info("\nГотово: %d книг (до фильтров собрано %d)", len(books), len(all_books))
    logging.info("Средняя цена: £%.2f", average)
    logging.info("Файл сохранён: %s", args.output)


if __name__ == "__main__":
    main()
