"""
Agoda scraper using Playwright.
Returns a list of {"hotel": str, "price": int, "currency": str, "url": str}.
"""

import logging
import os
from playwright.async_api import async_playwright

logger = logging.getLogger(__name__)

AGODA_URL = os.environ.get(
    "AGODA_URL",
    "https://www.agoda.com/search?city=17254&checkIn=2025-07-01&checkOut=2025-07-02&rooms=1&adults=2",
)


async def scrape_agoda() -> list[dict]:
    results = []
    logger.info("Scraping Agoda: %s", AGODA_URL)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        )
        page = await context.new_page()
        await page.goto(AGODA_URL, wait_until="networkidle", timeout=60_000)

        cards = await page.query_selector_all('[data-element-name="property-card"]')
        for card in cards[:20]:
            try:
                name_el = await card.query_selector('[data-element-name="property-card-name"]')
                price_el = await card.query_selector('[data-element-name="final-price"]')
                url_el = await card.query_selector("a")

                if not (name_el and price_el):
                    continue

                name = (await name_el.inner_text()).strip()
                price_text = (await price_el.inner_text()).strip()
                href = await url_el.get_attribute("href") if url_el else ""
                url = href if href.startswith("http") else f"https://www.agoda.com{href}"

                price = int("".join(filter(str.isdigit, price_text.split(".")[0])))
                currency = "JPY" if "¥" in price_text else "USD"

                results.append({"hotel": name, "price": price, "currency": currency, "url": url})
            except Exception as exc:
                logger.debug("Failed to parse Agoda card: %s", exc)

        await browser.close()

    logger.info("Agoda: %d hotels found", len(results))
    return results
