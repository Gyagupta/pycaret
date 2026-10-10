# IMD bin-hours v3: deck, zone and national workbook, and station workbook

| File | What it is |
|---|---|
| `IMD_BinHours_CLASP_FINAL_v3.pptx` | The v2 deck with three changes. Slide 6 is now the population view (share of people vs share of stations). New slide 7 puts the proposed table and today's table side by side. New slide 10 shows hours at 40 °C and above, degree by degree. Two annex references were also corrected (slide 4 → slide 5). |
| `IMD_BinHours_Zone_National_v3.xlsx` | Zone and national tables built only from numbers in the deck. Totals, bands, shares and checks are Excel formulas. |
| `build_station_sheets.py` | Builds the per-station workbook (see below). It needs the analysis output, which the deck does not contain. |
| `extract_deck_data.py`, `deck_data.json`, `build_excel.py`, `build_deck_v3.py` | Reproduce the two files above from the v2 deck. |

## Per-station workbook: how to produce it

The per-station tables need the hourly values from the analysis:
- each station's bins 24 to 43 and 44+;
- hours at every degree from below 0 °C up to the station's highest degree;
- each station's highest temperature and when it was reached;
- how many stations reach 40, 45 °C and so on;
- day and night hours per station.

1. In the analysis notebook, after step 6 ("count hours in 1 °C bins"), export the hourly table with one row per station per hour of every used day:
   `station_code, station_name, zone, date, utc_hour, temp_c`
   ```python
   hourly.to_parquet("hourly_values.parquet")   # or .to_csv("hourly_values.csv", index=False)
   ```
2. Run:
   ```
   python build_station_sheets.py hourly_values.parquet deck_data.json IMD_BinHours_Stations_v3.xlsx
   ```
3. The script first checks its zone totals against Annex 2a/2b: stations, days used, and every bin from 24 to 44+, in every zone. It writes the workbook only if every one of these 161 numbers matches exactly. It also compares its zone averages with slide 5, B with 6,372, and the count of stations at 44 °C+ with 50.

Bin, average-year and day/night rules are the deck's own (slides 14–16 and 8):
- **Bin:** 24.0–24.9 °C counts as 24, and 44 °C or more counts as 44+.
- **Average year:** each month's hours per usable day × the days in that month.
- **Day and night:** night runs from 6:30 pm to 5:30 am IST.
