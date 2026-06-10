import logging
from datetime import datetime
from playwright.sync_api import sync_playwright

logger = logging.getLogger(__name__)

BASE_URL = "https://www.booking.com/searchresults.html"


def scrape_booking(condition: dict) -> list:
    """
    Booking.com から condition に合うホテルを検索し価格リストを返す。
    """
    query = condition.get("hotel_name") or condition.get("location", "Tokyo")
    checkin  = condition.get("checkin", "")
    checkout = condition.get("checkout", "")
    guests   = int(condition.get("guests", 2))
    rooms    = int(condition.get("rooms", 1))
    free_cancel = condition.get("free_cancellation", False)

    params = (
        f"?ss={_q(query)}"
        f"&checkin={checkin}&checkout={checkout}"
        f"&group_adults={guests}&no_rooms={rooms}"
        f"&nflt=ht_id%3D204"  # hotel type
    )
    if free_cancel:
        params += "%3Bfc%3D3"  # free cancellation filter
    if condition.get("breakfast") == "included":
        params += "%3Bmealplan%3D1"

    url = BASE_URL + params
    results = []

    logger.info("Booking.com: %s", url)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ))
            page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(3000)

            cards = page.query_selector_all('[data-testid="property-card"]')
            for card in cards[:20]:
                try:
                    name_el  = card.query_selector('[data-testid="title"]')
                    price_el = card.query_selector('[data-testid="price-and-discounted-price"]')
                    url_el   = card.query_selector('a[data-testid="title-link"]')
                    if not (name_el and price_el):
                        continue
                    name       = name_el.inner_text().strip()
                    price_text = price_el.inner_text().strip()
                    href       = url_el.get_attribute("href") if url_el else ""
                    price      = _parse_price(price_text)
                    if price is None:
                        continue
                    results.append({
                        "site":       "Booking.com",
                        "hotel_name": name,
                        "price":      price,
                        "currency":   "JPY",
                        "url":        href,
                    })
                except Exception as e:
                    logger.debug("card parse error: %s", e)

            browser.close()
    except Exception as e:
        logger.error("Booking.com scrape failed: %s", e)

    logger.info("Booking.com: %d 件取得", len(results))
    return results


def _q(s: str) -> str:
    from urllib.parse import quote
    return quote(s)


def _parse_price(text: str):
    digits = "".join(c for c in text if c.isdigit())
    return int(digits) if digits else None
