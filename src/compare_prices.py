"""
Price comparison logic.
Merges Booking.com and Agoda results, ranks by price, and builds an HTML report.
"""

from __future__ import annotations
from datetime import date


def compare_prices(
    booking: list[dict],
    agoda: list[dict],
) -> dict:
    """
    Returns a dict with:
      - date: today's date string
      - booking_top5: cheapest 5 from Booking.com
      - agoda_top5: cheapest 5 from Agoda
      - overall_cheapest: single cheapest across both sources
      - html_body: HTML string for email
    """
    booking_sorted = sorted(booking, key=lambda x: x["price"])[:5]
    agoda_sorted = sorted(agoda, key=lambda x: x["price"])[:5]

    all_hotels = booking + agoda
    overall_cheapest = min(all_hotels, key=lambda x: x["price"]) if all_hotels else None

    today = date.today().isoformat()
    html_body = _build_html(today, booking_sorted, agoda_sorted, overall_cheapest)

    return {
        "date": today,
        "booking_top5": booking_sorted,
        "agoda_top5": agoda_sorted,
        "overall_cheapest": overall_cheapest,
        "html_body": html_body,
    }


def _build_html(
    today: str,
    booking: list[dict],
    agoda: list[dict],
    cheapest: dict | None,
) -> str:
    def rows(hotels: list[dict]) -> str:
        if not hotels:
            return "<tr><td colspan='3'>データなし</td></tr>"
        return "".join(
            f"<tr><td><a href='{h['url']}'>{h['hotel']}</a></td>"
            f"<td>{h['price']:,} {h['currency']}</td></tr>"
            for h in hotels
        )

    cheapest_section = ""
    if cheapest:
        cheapest_section = (
            f"<p><strong>🏆 本日の最安値:</strong> "
            f"<a href='{cheapest['url']}'>{cheapest['hotel']}</a> "
            f"— {cheapest['price']:,} {cheapest['currency']}</p>"
        )

    return f"""
<html><body>
<h2>ホテル最安値レポート ({today})</h2>
{cheapest_section}
<h3>Booking.com TOP 5</h3>
<table border="1" cellpadding="4" cellspacing="0">
<tr><th>ホテル名</th><th>価格</th></tr>
{rows(booking)}
</table>
<h3>Agoda TOP 5</h3>
<table border="1" cellpadding="4" cellspacing="0">
<tr><th>ホテル名</th><th>価格</th></tr>
{rows(agoda)}
</table>
</body></html>
""".strip()
