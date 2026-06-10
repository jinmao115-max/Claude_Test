import logging
import re
from urllib.parse import quote
from playwright.sync_api import sync_playwright

logger = logging.getLogger(__name__)

BASE_URL = "https://www.booking.com/searchresults.ja.html"


def scrape_booking(condition: dict) -> list:
    """
    Booking.com（日本語版・JPY）から指定ホテル/地域の価格を取得する。
    """
    hotel_name  = condition.get("hotel_name", "").strip()
    location    = condition.get("location", "").strip()
    # ホテル名指定が優先、なければ地域名で検索
    query       = hotel_name or location or "東京"
    is_hotel    = bool(hotel_name)   # True=ホテル名指定, False=地域検索

    checkin     = condition.get("checkin", "")
    checkout    = condition.get("checkout", "")
    guests      = int(condition.get("guests", 2))
    rooms       = int(condition.get("rooms", 1))
    free_cancel = condition.get("free_cancellation", False)
    breakfast   = condition.get("breakfast", "any")

    url = (
        f"{BASE_URL}?ss={quote(query)}"
        f"&lang=ja&selected_currency=JPY"
        f"&checkin={checkin}&checkout={checkout}"
        f"&group_adults={guests}&no_rooms={rooms}"
    )
    if free_cancel:
        url += "&nflt=fc%3D3"
    if breakfast == "included":
        url += "%3Bmealplan%3D1"

    results = []
    logger.info("Booking.com 検索: %s (モード: %s)", query, "ホテル名" if is_hotel else "地域")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(
                locale="ja-JP",
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
            )
            page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(4000)

            try:
                page.click('[id="onetrust-accept-btn-handler"]', timeout=3000)
                page.wait_for_timeout(1000)
            except Exception:
                pass

            cards = page.query_selector_all('[data-testid="property-card"]')
            for card in cards[:30]:
                try:
                    name_el  = card.query_selector('[data-testid="title"]')
                    price_el = card.query_selector('[data-testid="price-and-discounted-price"]')
                    url_el   = card.query_selector('a[data-testid="title-link"]')
                    if not (name_el and price_el):
                        continue
                    h_name     = name_el.inner_text().strip()
                    price_text = price_el.inner_text().strip()
                    href       = url_el.get_attribute("href") if url_el else ""
                    price      = _parse_jpy(price_text)
                    if price is None or price < 1000:
                        continue
                    results.append({
                        "site":       "Booking.com",
                        "hotel_name": h_name,
                        "price":      price,
                        "currency":   "JPY",
                        "url":        href,
                    })
                except Exception as e:
                    logger.debug("card parse error: %s", e)

            browser.close()
    except Exception as e:
        logger.error("Booking.com scrape failed: %s", e)

    results = _filter_results(results, query, is_hotel)
    logger.info("Booking.com: %d 件取得", len(results))
    return results


def _parse_jpy(text: str):
    nums = re.findall(r"[\d,]+", text)
    for n in nums:
        val = int(n.replace(",", ""))
        if val >= 1000:
            return val
    return None


def _filter_results(results: list, query: str, is_hotel: bool) -> list:
    """
    is_hotel=True  → ホテル名の全キーワードが一致するものだけ返す（完全一致寄り）
    is_hotel=False → 地域名キーワードを含むものに絞る（0件なら全件返す）
    """
    if not query:
        return results

    keywords = re.split(r"[\s　]+", query.strip())
    keywords = [k.lower() for k in keywords if k]

    if is_hotel:
        # 全キーワードが hotel_name に含まれるもののみ
        filtered = [
            r for r in results
            if all(k in r["hotel_name"].lower() for k in keywords)
        ]
        # 完全一致ゼロなら部分一致（少なくとも1キーワード）にフォールバック
        if not filtered:
            filtered = [
                r for r in results
                if any(k in r["hotel_name"].lower() for k in keywords)
            ]
    else:
        # 地域検索：地域キーワードを含むものに絞る
        filtered = [
            r for r in results
            if any(k in r["hotel_name"].lower() for k in keywords)
        ]
        if not filtered:
            filtered = results  # フォールバック

    return filtered
