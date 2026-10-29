# Hotel Booking Data Cleaning & Analysis

A beginner-friendly data cleaning and reporting project using **Python**, **Pandas**, and **Excel**. The project works with a practice hotel-bookings dataset, identifies data-quality issues, and exports an Excel workbook with separate sheets for review.

## Project goals

- Remove duplicate records and rows missing all key identifying/date fields.
- Parse check-in and check-out dates and calculate stay lengths.
- Separate valid-duration bookings from records requiring review.
- Flag missing check-in dates, unusually long stays, and differences between recorded and expected prices.
- Create a multi-sheet Excel report with booking and financial summaries.

## Tools

- Python
- Pandas
- OpenPyXL
- Microsoft Excel (for final formatting)

## Data-cleaning workflow

1. Load `hotel_bookings_practice_fixed.csv`.
2. Identify and remove duplicates.
3. Remove rows where `Check_In`, `Check_Out`, and `Full_Name` are all missing.
4. Convert dates using `pd.to_datetime(..., format='mixed', errors='coerce')`.
5. Separate records with missing check-in dates.
6. Calculate `Nights` and separate non-positive stay lengths.
7. Flag stays longer than 30 nights.
8. Calculate an **assumed** expected price at **$40 per night** and fill missing prices for analysis, retaining a `Price_was_missing` indicator.
9. Compare recorded prices with assumed expected prices and export review sheets.

**Important:** The $40/night rate is a practice assumption, not a verified hotel rate. Estimated check-in dates and filled prices are not confirmed facts. Price differences are review flags, not proof of errors.

## Results

| Metric | Value |
| --- | ---: |
| Original records | 124 |
| Duplicate records removed | 2 |
| Rows removed with all three key fields missing | 2 |
| Records remaining after removal | 120 |
| Bookings with positive stay lengths | 106 |
| Bookings with non-positive stay lengths | 9 |
| Records with missing check-in dates | 5 |
| Long stays flagged (>30 nights) | 5 |
| Missing prices filled with estimates | 6 |

### Financial summary

| Category | Amount (USD) |
| --- | ---: |
| Recorded prices in positive-duration bookings | $15,040 |
| Estimated prices in positive-duration bookings | $1,080 |
| Total price value of positive-duration bookings | $16,120 |
| Price value of bookings missing check-in dates | $840 |
| Total included booking price value | $16,960 |

These amounts are **not verified revenue or profit**. They are sums of recorded and, where indicated, estimated booking prices. Invalid-duration bookings and removed records are excluded from these totals.

## Files

- `main.py` — Python data-cleaning and Excel-export script (or your script's actual filename).
- `hotel_bookings_practice_fixed.csv` — original practice input dataset.
- `Hotel_Report_Final.xlsx` — formatted multi-sheet Excel report (or the actual uploaded report filename).
- `requirements.txt` — Python dependencies.

## Run locally

1. Install Python 3 and download or clone this repository.
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Keep the input CSV in the path expected by the Python script.
4. Run the script (replace `main.py` if your filename is different):

   ```bash
   python main.py
   ```

5. Open the generated Excel workbook to inspect the reports. Final manual Excel formatting may be present only in the uploaded final workbook.

## Notes

This is a **practice/portfolio project**. The dataset should contain only synthetic or otherwise publicly shareable information. Do not publish real guests' personal details.
