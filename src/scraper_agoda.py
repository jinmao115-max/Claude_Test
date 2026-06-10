import logging
from urllib.parse import quote
from playwright.sync_api import sync_playwright

logger = logging.getLogger(__name__)

BASE_URL = "https://www.agoda.com/search"


def scrape_agoda(condition: dict) -> list:
    """
    Agoda から condition に合うホテルを検索し価格リストを返す。
    """
    query    = condition.get("hotel_name") or condition.get("location", "Tokyo")
    checkin  = condition.get("checkin", "")
    checkout = condition.get("checkout", "")
    guests   = int(condition.get("guests", 2))
    rooms    = int(condition.get("rooms", 1))
    free_cancel = condition.get("free_cancellation", False)

    params = (
        f"?city=17254"          # Tokyo city id (fallback)
        f"&textToSearch={quote(query)}"
        f"&checkIn={checkin}&checkOut={checkout}"
        f"&rooms={rooms}&adults={guests}"
        f"&los=1&locale=ja-jp&currency=JPY"
    )
    if free_cancel:
        params += "&filterByFreeCancellation=true"
    if condition.get("breakfast") == "included":
        params += "&filterByBreakfastIncluded=true"

    url = BASE_URL + params
    results = []

    logger.info("Agoda: %s", url)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
                locale="ja-JP",
            )
            page = context.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(4000)

            cards = page.query_selector_all('[data-element-name="property-card"]')
            for card in cards[:20]:
                try:
                    name_el  = card.query_selector('[data-element-name="property-card-name"]')
                    price_el = card.query_selector('[data-element-name="final-price"]')
                    url_el   = card.query_selector("a")
                    if not (name_el and price_el):
                        continue
                    name       = name_el.inner_text().strip()
                    price_text = price_el.inner_text().strip()
                    href       = url_el.get_attribute("href") if url_el else ""
                    full_url   = href if href.startswith("http") else f"https://www.agoda.com{href}"
                    price      = _parse_price(price_text)
                    if price is None:
                        continue
                    results.append({
                        "site":       "Agoda",
                        "hotel_name": name,
                        "price":      price,
                        "currency":   "JPY",
                        "url":        full_url,
                    })
                except Exception as e:
                    logger.debug("card parse error: %s", e)

            browser.close()
    except Exception as e:
        logger.error("Agoda scrape failed: %s", e)

    logger.info("Agoda: %d 件取得", len(results))
    return results


def _parse_price(text: str):
    digits = "".join(c for c in text if c.isdigit())
    return int(digits) if digits else None
