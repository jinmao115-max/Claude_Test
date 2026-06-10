import logging
import re
from urllib.parse import quote
from playwright.sync_api import sync_playwright

logger = logging.getLogger(__name__)

BASE_URL = "https://www.agoda.com/ja-jp/search"


def scrape_agoda(condition: dict) -> list:
    """
    Agoda（日本語版・JPY）から指定ホテル/地域の価格を取得する。
    """
    hotel_name  = condition.get("hotel_name", "").strip()
    location    = condition.get("location", "").strip()
    query       = hotel_name or location or "東京"
    is_hotel    = bool(hotel_name)

    checkin     = condition.get("checkin", "")
    checkout    = condition.get("checkout", "")
    guests      = int(condition.get("guests", 2))
    rooms       = int(condition.get("rooms", 1))
    free_cancel = condition.get("free_cancellation", False)
    breakfast   = condition.get("breakfast", "any")

    url = (
        f"{BASE_URL}?textToSearch={quote(query)}"
        f"&checkIn={checkin}&checkOut={checkout}"
        f"&rooms={rooms}&adults={guests}"
        f"&currency=JPY&locale=ja-JP"
    )
    if free_cancel:
        url += "&filterByFreeCancellation=true"
    if breakfast == "included":
        url += "&filterByBreakfastIncluded=true"

    results = []
    logger.info("Agoda 検索: %s (モード: %s)", query, "ホテル名" if is_hotel else "地域")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                locale="ja-JP",
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
            )
            page = context.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(6000)

            cards = (
                page.query_selector_all('[data-element-name="property-card"]')
                or page.query_selector_all('[class*="PropertyCard"]')
                or page.query_selector_all('[data-testid="property-card"]')
            )
            logger.info("Agoda: %d 件のカードを検出", len(cards))

            for card in cards[:30]:
                try:
                    name_el = (
                        card.query_selector('[data-element-name="property-card-name"]')
                        or card.query_selector('h3')
                        or card.query_selector('[class*="PropertyName"]')
                    )
                    price_el = (
                        card.query_selector('[data-element-name="final-price"]')
                        or card.query_selector('[data-selenium="display-price"]')
                        or card.query_selector('[class*="Price"][class*="final"]')
                        or card.query_selector('[class*="price"]')
                    )
                    url_el = card.query_selector("a")

                    if not (name_el and price_el):
                        continue

                    h_name     = name_el.inner_text().strip()
                    price_text = price_el.inner_text().strip()
                    href       = url_el.get_attribute("href") if url_el else ""
                    full_url   = href if href.startswith("http") else f"https://www.agoda.com{href}"
                    price      = _parse_jpy(price_text)

                    if price is None or price < 1000:
                        continue

                    results.append({
                        "site":       "Agoda",
                        "hotel_name": h_name,
                        "price":      price,
                        "currency":   "JPY",
                        "url":        full_url,
                    })
                except Exception as e:
                    logger.debug("card parse error: %s", e)

            if not results:
                try:
                    page.screenshot(path="agoda_debug.png")
                    logger.warning("Agoda: 結果なし。agoda_debug.png を確認してください")
                except Exception:
                    pass

            browser.close()
    except Exception as e:
        logger.error("Agoda scrape failed: %s", e)

    results = _filter_results(results, query, is_hotel)
    logger.info("Agoda: %d 件取得", len(results))
    return results


def _parse_jpy(text: str):
    nums = re.findall(r"[\d,]+", text)
    for n in nums:
        val = int(n.replace(",", ""))
        if val >= 1000:
            return val
    return None


def _filter_results(results: list, query: str, is_hotel: bool) -> list:
    """フィルタリングを緩くして結果が0件にならないようにする"""
    if not results:
        return results

    # エリア検索（is_hotel=False）は全件返す
    if not is_hotel:
        return results

    # ホテル名検索の場合も全件返す（メールで確認できるように）
    return results
