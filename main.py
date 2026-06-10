"""
Hotel Price Checker - Main Entry Point
Scrapes hotel prices from Booking.com and Agoda, compares them, and sends a Gmail notification.
"""

import asyncio
import logging
from datetime import datetime

from src.scraper_booking import scrape_booking
from src.scraper_agoda import scrape_agoda
from src.compare_prices import compare_prices
from src.send_gmail import send_gmail

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


async def main():
    logger.info("Hotel price check started at %s", datetime.now().strftime("%Y-%m-%d %H:%M"))

    booking_prices = await scrape_booking()
    agoda_prices = await scrape_agoda()

    report = compare_prices(booking_prices, agoda_prices)

    send_gmail(report)
    logger.info("Done.")


if __name__ == "__main__":
    asyncio.run(main())
