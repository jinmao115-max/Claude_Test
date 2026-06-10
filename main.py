import json
import sys
import logging
from src.compare_prices import compare_prices
from src.send_gmail import send_email

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

def normalize_condition(c):
    """WebUIのキー名をPythonコードのキー名に変換"""
    return {
        "hotel_name": c.get("hotel_name") or c.get("location", ""),
        "location": c.get("location", ""),
        "checkin": c.get("checkin", ""),
        "checkout": c.get("checkout", ""),
        "guests": c.get("guests") or c.get("adults", 2),
        "rooms": c.get("rooms", 1),
        "bed_type": c.get("bed_type") or c.get("bedType", "any"),
        "breakfast": c.get("breakfast", "any"),
        "free_cancellation": c.get("free_cancellation") or (c.get("cancelFree") == "free"),
    }

def main():
    try:
        with open("config/search_conditions.json", "r", encoding="utf-8") as f:
            config = json.load(f)
    except Exception as e:
        logger.error("設定ファイルの読み込みに失敗: %s", e)
        sys.exit(1)

    # searches / conditions どちらのキーにも対応
    searches = config.get("searches") or config.get("conditions", [])
    notify_email = (
        config.get("settings", {}).get("notify_email")
        or config.get("settings", {}).get("email")
        or "jinmao115@gmail.com"
    )

    if not searches:
        logger.warning("検索条件が設定されていません")
        sys.exit(0)

    logger.info("%d 件の検索条件を処理します", len(searches))

    all_results = []
    for i, condition in enumerate(searches, 1):
        condition = normalize_condition(condition)
        name = condition.get("hotel_name") or condition.get("location", "不明")
        logger.info("[%d/%d] %s を検索中...", i, len(searches), name)
        result = compare_prices(condition)
        all_results.append(result)
        if result["cheapest"]:
            cheapest = result["cheapest"]
            logger.info("  最安値: %s ¥%s", cheapest["site"], f"{cheapest['price']:,}")
        else:
            logger.info("  結果なし")

    logger.info("%s にメールを送信中...", notify_email)
    send_email(notify_email, all_results)
    logger.info("完了！")

if __name__ == "__main__":
    main()
