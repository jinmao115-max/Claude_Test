"""
Booking.com scraper using Playwright.
Returns a list of {"hotel": str, "price": int, "currency": str, "url": str}.
"""

import logging
import os
from playwright.async_api import async_playwright

logger = logging.getLogger(__name__)

BOOKING_URL = os.environ.get(
    "BOOKING_URL",
    "https://www.booking.com/searchresults.html?ss=Tokyo&checkin=2025-07-01&checkout=2025-07-02&group_adults=2",
)


async def scrape_booking() -> list[dict]:
    results = []
    logger.info("Scraping Booking.com: %s", BOOKING_URL)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.goto(BOOKING_URL, wait_until="networkidle", timeout=60_000)

        cards = await page.query_selector_all('[data-testid="property-card"]')
        for card in cards[:20]:
            try:
                name_el = await card.query_selector('[data-testid="title"]')
                price_el = await card.query_selector('[data-testid="price-and-discounted-price"]')
                url_el = await card.query_selector('a[data-testid="title-link"]')

                if not (name_el and price_el):
                    continue

                name = (await name_el.inner_text()).strip()
                price_text = (await price_el.inner_text()).strip()
                url = await url_el.get_attribute("href") if url_el else ""

                price = int("".join(filter(str.isdigit, price_text.split(".")[0])))
                currency = "JPY" if "¥" in price_text else "USD"

                results.append({"hotel": name, "price": price, "currency": currency, "url": url})
            except Exception as exc:
                logger.debug("Failed to parse Booking.com card: %s", exc)

        await browser.close()

    logger.info("Booking.com: %d hotels found", len(results))
    return results
