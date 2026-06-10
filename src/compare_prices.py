from src.scraper_booking import scrape_booking
from src.scraper_agoda import scrape_agoda

def compare_prices(condition: dict) -> dict:
    """
    全サイトから価格を取得し、最安値順に並べて返す
    """
    all_results = []

    print("[比較] Booking.com を検索中...")
    try:
        booking_results = scrape_booking(condition)
        all_results.extend(booking_results)
    except Exception as e:
        print(f"[Booking.com] エラー: {e}")

    print("[比較] Agoda を検索中...")
    try:
        agoda_results = scrape_agoda(condition)
        all_results.extend(agoda_results)
    except Exception as e:
        print(f"[Agoda] エラー: {e}")

    all_results.sort(key=lambda x: x["price"])

    return {
        "condition": condition,
        "results": all_results,
        "cheapest": all_results[0] if all_results else None,
    }
