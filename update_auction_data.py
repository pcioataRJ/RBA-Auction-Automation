import re
            return nums[-1] - nums[0] + 1
        return 1


class AuctionWorkbookUpdater:
    def __init__(self, workbook_path: Path, worksheet_name: str):
        self.workbook_path = workbook_path
        self.worksheet_name = worksheet_name

    def update(self, records: List[AuctionRecord]) -> int:
        wb = load_workbook(self.workbook_path)
        ws = wb[self.worksheet_name]

        existing = self._existing_keys(ws)
        rows_added = 0

        for record in sorted(records, key=lambda r: (r.parsed_date, r.auction_name)):
            key = self._record_key(record)
            if key in existing:
                continue

            row = self._next_empty_row(ws)
            ws.cell(row=row, column=COL_YEAR).value = record.parsed_date.year
            ws.cell(row=row, column=COL_MONTH).value = record.month_num
            ws.cell(row=row, column=COL_TAG).value = record.tag_value
            ws.cell(row=row, column=COL_DATE).value = record.date_for_excel
            ws.cell(row=row, column=COL_AUCTION).value = record.auction_name
            ws.cell(row=row, column=COL_DAYS).value = record.days
            ws.cell(row=row, column=COL_LOTS).value = record.lots

            self._copy_formula_down(ws, row, COL_FORMULA_F)
            self._copy_formula_down(ws, row, COL_QTD_DAYS)
            self._copy_formula_down(ws, row, COL_QTD_LOTS)

            fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
            for col in range(COL_YEAR, COL_QTD_LOTS + 1):
                ws.cell(row=row, column=col).fill = fill

            existing.add(key)
            rows_added += 1

        wb.save(self.workbook_path)
        return rows_added

    def _existing_keys(self, ws):
        keys = set()
        for row in range(START_DATA_ROW, ws.max_row + 1):
            date_val = ws.cell(row=row, column=COL_DATE).value
            name_val = ws.cell(row=row, column=COL_AUCTION).value
            if date_val and name_val:
                if isinstance(date_val, datetime):
                    date_key = date_val.strftime("%Y-%m-%d")
                else:
                    date_key = str(date_val)
                keys.add((date_key.strip(), str(name_val).strip().lower()))
        return keys

    def _record_key(self, record: AuctionRecord):
        return (record.parsed_date.strftime("%Y-%m-%d"), record.auction_name.strip().lower())

    def _next_empty_row(self, ws) -> int:
        row = START_DATA_ROW
        while ws.cell(row=row, column=COL_AUCTION).value:
            row += 1
        return row

    def _copy_formula_down(self, ws, row: int, col: int):
        if row <= START_DATA_ROW:
            return
        source = ws.cell(row=row - 1, column=col)
        target = ws.cell(row=row, column=col)
        if isinstance(source.value, str) and source.value.startswith("="):
            target.value = source.value


def main():
    scraper = RBAuctionScraper(headless=True)
    updater = AuctionWorkbookUpdater(WORKBOOK_PATH, WORKSHEET_NAME)

    records = scraper.scrape_past_auctions()
    print(f"Scraped {len(records)} past auctions")

    rows_added = updater.update(records)
    print(f"Added {rows_added} new rows to workbook: {WORKBOOK_PATH}")


if __name__ == "__main__":
    main()
