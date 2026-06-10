import logging
import re
from urllib.parse import quote
from playwright.sync_api import sync_playwright

logger = logging.getLogger(__name__)

# 日本語版・JPY固定
BASE_URL = "https://www.booking.com/searchresults.ja.html"


def scrape_booking(condition: dict) -> list:
    """
    Booking.com（日本語版・JPY）から指定ホテル/地域の価格を取得する。
    """
    query    = condition.get("hotel_name") or condition.get("location", "東京")
    checkin  = condition.get("checkin", "")
    checkout = condition.get("checkout", "")
    guests   = int(condition.get("guests", 2))
    rooms    = int(condition.get("rooms", 1))
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
    logger.info("Booking.com 検索: %s", query)

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
            # クッキー同意ポップアップ対策
            page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(4000)

            # クッキー同意ボタンがあれば閉じる
            try:
                page.click('[id="onetrust-accept-btn-handler"]', timeout=3000)
                page.wait_for_timeout(1000)
            except Exception:
                pass

            cards = page.query_selector_all('[data-testid="property-card"]')
            for card in cards[:20]:
                try:
                    name_el  = card.query_selector('[data-testid="title"]')
                    price_el = card.query_selector('[data-testid="price-and-discounted-price"]')
                    url_el   = card.query_selector('a[data-testid="title-link"]')
                    if not (name_el and price_el):
                        continue
                    hotel_name = name_el.inner_text().strip()
                    price_text = price_el.inner_text().strip()
                    href       = url_el.get_attribute("href") if url_el else ""
                    price      = _parse_jpy(price_text)
                    if price is None or price < 1000:  # 1000円未満は誤パース
                        continue
                    results.append({
                        "site":       "Booking.com",
                        "hotel_name": hotel_name,
                        "price":      price,
                        "currency":   "JPY",
                        "url":        href,
                    })
                except Exception as e:
                    logger.debug("card parse error: %s", e)

            browser.close()
    except Exception as e:
        logger.error("Booking.com scrape failed: %s", e)

    # ホテル名指定がある場合は名前で絞り込み
    results = _filter_by_name(results, query)
    logger.info("Booking.com: %d 件取得", len(results))
    return results


def _parse_jpy(text: str):
    """
    "¥12,345" や "12,345円" などから整数を取り出す。
    カンマ区切り対応。
    """
    # カンマを除いた連続する数字ブロックを全て抽出
    nums = re.findall(r"[\d,]+", text)
    for n in nums:
        val = int(n.replace(",", ""))
        if val >= 1000:   # 宿泊料金として妥当な下限
            return val
    return None


def _filter_by_name(results: list, query: str) -> list:
    """
    query のキーワードを少なくとも1つ含むホテルに絞り込む。
    一致ゼロの場合は全件返す（地域検索のフォールバック）。
    """
    if not query:
        return results
    # スペース・全角スペースで分割してキーワードリスト化
    keywords = re.split(r"[\s　]+", query.strip())
    keywords = [k.lower() for k in keywords if k]
    filtered = [
        r for r in results
        if any(k in r["hotel_name"].lower() for k in keywords)
    ]
    return filtered if filtered else results
