import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from playwright.sync_api import sync_playwright

AUCTION_URL = "https://www.rbauction.com/heavy-equipment-auctions"
WORKBOOK_PATH = Path("Copy of RBA Auction Data V2.xlsx")
WORKSHEET_NAME = "Quarter Summaries"
START_DATA_ROW = 6

COL_YEAR = 2
COL_MONTH = 3
COL_TAG = 4
COL_DATE = 5
COL_FORMULA_F = 6
COL_AUCTION = 8
COL_DAYS = 9
COL_LOTS = 10
COL_QTD_DAYS = 11
COL_QTD_LOTS = 12


@dataclass
class AuctionRecord:
    auction_date_text: str
    auction_name: str
    location_text: str
    lots: int
    days: int

    @property
    def parsed_date(self):
        year = datetime.today().year
        month, day = re.findall(r"([A-Z][a-z]{2}) (\d{1,2})", self.auction_date_text)[0]
        return datetime.strptime(f"{month} {day} {year}", "%b %d %Y")

    @property
    def month_num(self):
        return self.parsed_date.month

    @property
    def tag_value(self):
        return f"{self.month_num}-{str(self.parsed_date.year)[-2:]}"


def scrape_auctions():

    records = []

    with sync_playwright() as p:

        browser = p.chromium.launch()
        page = browser.new_page()

        page.goto(AUCTION_URL)

        page.get_by_text("Past").click()

        page.wait_for_timeout(3000)

        cards = page.locator("h5")
        count = cards.count()

        for i in range(count):

            name = cards.nth(i).inner_text()

            card = cards.nth(i).locator("xpath=ancestor::a")

            text = card.inner_text()

            date = re.search(r"[A-Z][a-z]{2} \d{1,2}", text)

            lots = re.search(r"(\d{1,3}(,\d{3})*) Items", text)

            if not date or not lots:
                continue

            records.append(
                AuctionRecord(
                    auction_date_text=date.group(),
                    auction_name=name,
                    location_text="",
                    lots=int(lots.group(1).replace(",", "")),
                    days=1,
                )
            )

        browser.close()

    return records


class AuctionWorkbookUpdater:

    def __init__(self):
        self.wb = load_workbook(WORKBOOK_PATH)
        self.ws = self.wb[WORKSHEET_NAME]

    def existing_keys(self):

        keys = set()

        for row in range(START_DATA_ROW, self.ws.max_row + 1):

            date = self.ws.cell(row=row, column=COL_DATE).value
            name = self.ws.cell(row=row, column=COL_AUCTION).value

            if date and name:
                keys.add((str(date), name.lower()))

        return keys

    def next_row(self):

        row = START_DATA_ROW

        while self.ws.cell(row=row, column=COL_AUCTION).value:
            row += 1

        return row

    def copy_formula(self, row, col):

        src = self.ws.cell(row=row - 1, column=col)

        if isinstance(src.value, str) and src.value.startswith("="):
            self.ws.cell(row=row, column=col).value = src.value

    def update(self, records):

        existing = self.existing_keys()

        added = 0

        for r in records:

            key = (str(r.parsed_date), r.auction_name.lower())

            if key in existing:
                continue

            row = self.next_row()

            self.ws.cell(row=row, column=COL_YEAR).value = r.parsed_date.year
            self.ws.cell(row=row, column=COL_MONTH).value = r.month_num
            self.ws.cell(row=row, column=COL_TAG).value = r.tag_value
            self.ws.cell(row=row, column=COL_DATE).value = r.parsed_date
            self.ws.cell(row=row, column=COL_AUCTION).value = r.auction_name
            self.ws.cell(row=row, column=COL_DAYS).value = r.days
            self.ws.cell(row=row, column=COL_LOTS).value = r.lots

            self.copy_formula(row, COL_FORMULA_F)
            self.copy_formula(row, COL_QTD_DAYS)
            self.copy_formula(row, COL_QTD_LOTS)

            added += 1

        self.wb.save(WORKBOOK_PATH)

        return added


def main():

    records = scrape_auctions()

    updater = AuctionWorkbookUpdater()

    added = updater.update(records)

    print("Rows added:", added)


if __name__ == "__main__":
    main()
