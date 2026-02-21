# main.py
from scraper import YallaKoraScraper
from storage import save_to_excel
from config import OUTPUT_FILE


def run():
    date = input('أدخل التاريخ بتنسيق dd/mm/yyyy: ')

    scraper = YallaKoraScraper()
    soup = scraper.get_matches_page(date)

    print("⏳ جاري تحليل البيانات...")
    matches_data = scraper.extract_matches_details(soup)

    save_to_excel(matches_data, OUTPUT_FILE)


if __name__ == "__main__":
    run()
