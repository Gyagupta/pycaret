# IMD bin-hour deck: revisions for BEE review

The original PPTX, `IMD_BinHours_by_Zone_2_WITH_SPLINE.xlsx` and the raw IMD CSV were not available in this cloud workspace. Everything below is built from the published methodology note ("Zone Bin-Hour Methodology": IMD 2016–2025, zone averages) and BEE Schedule 19, Table 1.

`BinHour_addon_slides.pptx` contains 6 slides to paste into the deck (PowerPoint: Paste → *Use Destination Theme*). Build script: `build_addon_slides.js`.

| Add-on slide | Goes into deck as | Status |
|---|---|---|
| 1. Selection basis and IMD data limits | Slide 2 (replaces the challenges box) | Done |
| 2. Method, simplified 14:30 line and comparison panel | Slide 3 | Done |
| 3. Zone table with "43 °C and above" column and clear units | Slide 4 | Done |
| 4. Options for 44 °C and above, with results | Slide 5 | Done. Check that the option labels match the deck |
| 5. Assumptions and next steps | Last content slide | Done |
| 6. Annex: rough comparison with 5,774 | After Thank You | Done (zone totals) |
| Annex: 10-year counted hours by zone, per bin (24…43, 44+, total) | After Thank You | **Needs the spline workbook** |
| Annex: per-bin zone average compared with BEE's 5,774 | After Thank You | **Needs the spline workbook** (the zone-total version is above) |

---

## Slide 1
No change.

## Slide 2: basis of selection, and data challenges in place of the Undecided stations
**Basis of selection (449 → 128):**
- Complete day: all 8 three-hourly readings present (85.8% of station-days for the 128 stations).
- Usable cooling season (1 Mar–31 Oct): ≥90% complete days and no run of missing days longer than 15.
- Tier by usable seasons out of 10: A 8–10 (43), B 5–7 (56), C 1–4 (29, flagged), D 0 (321, dropped).
- Typical-year test: ≥28 usable days in every calendar month over the 10 years. All 128 pass with the spline. Without the spline, Bilaspur and Visakhapatnam (A) fail.
- Zone tag: ECBC 2017 list (43), NBC 2016 regional rule (78), project precedent (1, Coimbatore), undecided (6).
- The rule is the same as in the earlier notebook, and all 449 stations reproduce the same tier.

**IMD data challenges (these replace "6 undecided stations"):**
- Only 8 readings a day, with no hourly record. The true daily peak can fall between the 14:30 and 17:30 IST readings.
- 321 of the 449 station codes (71%) have no usable cooling season in 2016–2025.
- Of the 128 stations' station-days, 14.2% are incomplete: 4.5% are fillable (1–2 non-adjacent readings missing), 2.7% have a 9-hour gap and 7.0% are missing three or more readings.
- Coverage is uneven: Warm-Humid has 59 stations, Temperate 2 and Cold 3. 29 stations are Tier C.
- The data needed cleaning: 4,209 blank rows, repeated readings and impossible dates were removed (5,312,208 rows → 5,305,133 valid readings).
- IMD metadata has no climatic-zone field, so zones were assigned from ECBC or NBC.

## Slide 3: the "14:30 hottest reading" line
**What the line meant:** IMD records temperature 8 times a day. When those 8 readings are averaged over April–May at every station, the 14:30 IST reading is the highest (34.18 °C). So one of IMD's fixed observation times falls close to the hottest part of the day, which means the hot bins rest on a real measurement and not a guess. It answers the expert question: "Doesn't 3-hourly data miss the afternoon peak?"

**Caveat on the old wording:** "the peak is observed directly" overstates it. The true maximum often falls between 14:30 and 17:30. In the Ahmedabad example, the readings were 42.2 °C and 42.6 °C and the spline peaks at 43.2 °C around 16:00.

**Replacement (outcome line):**
> IMD takes a reading at 14:30 IST, close to the hottest time of day, so the hot-hour bins rest on real measurements. The spline only fills the hours between readings.

**Side panel: who else builds bin-hour tables**

| Method | Data | Bins | Same as ours | Different |
|---|---|---|---|---|
| India, BEE Schedule 19 / IS 1391 (Part 2), in force | ISHRAE Indian Weather Data Handbook 2014, 54 cities | 24–43 °C, 1 °C; 5,774 h scaled to 1,600 | Bins, 24 °C cut-off, whole-year basis | One typical year per city; one national table; stops at 43 °C |
| ISO 16358-1:2013, adopted by Malaysia, Thailand, Indonesia, Philippines, Vietnam | Fixed international reference | 21–35 °C, 1,817 h | Same bin-hour formula ISEER uses | Not local data. Malaysia applies ISO Table 3 unchanged (annual energy at 4,380 h = 12 h × 365) |
| Saudi Arabia, SASO 2663:2021 | National weather data in the ISO 16358 framework | Bins above 20 °C | Country data replaces the ISO default and covers very hot hours | One national table |
| China, GB 21455 | National weather data | 24–38 °C, 1 °C (2013 edition) | 1 °C bins starting at 24 °C | Top bin 38 °C; one national table |
| US Engineering Weather Data (AFM 88-29, 1978) | Airfield observations, 3-hourly where hourly is missing | 5 °F | Each 3-hourly reading counts as 3 hours, as in our Method A | No interpolation |
| Buildings & Cities (2021), India ISEER study | Typical-year weather files (2003–17), 55 cities | 24–43 °C | City-level bins tested against ISEER | Hourly typical-year files, not a raw station record |

**On the earlier Malaysia report:** I could not see that report from here. On the public record, Malaysia (Energy Commission MEPS guide) does not derive its own bins. It uses the ISO 16358-1 Table 3 distribution. The LBNL/U4E ASEAN harmonisation work found that ASEAN-specific bins barely change CSPF compared with the ISO bins, which is why one ISO table was kept. For India the opposite holds: the ISO table stops at 35 °C, while our Composite and Hot-Dry zones have 690–1,020 hours a year at 35 °C or above. That is why local data matters more here.

**Bottom line for the panel:** we found no other scheme that builds zone-wise tables from a 10-year raw station record. Most use one national table, built either from typical-year files or from the ISO default.

## Slide 4: add a 43 °C+ column and make the units readable
Add this header note: *"Each figure = hours in ONE typical year at ONE station, averaged over the zone's stations. Built from 10 years of IMD readings (2016–2025); not a 10-year total."*

| Zone | Stations | All hours ≥24 °C | vs 5,774 | Hours ≥35 °C | 43–44 °C | 44 °C+ | **43 °C+ (new)** |
|---|---:|---:|---:|---:|---:|---:|---:|
| Composite | 42 | 5,598 | −3% | 692 | 15.4 | 12.2 | **27.6** |
| Hot-Dry | 16 | 6,222 | +8% | 1,019 | 20.8 | 26.0 | **46.8** |
| Warm-Humid | 59 | 7,226 | +25% | 256 | 0.7 | 0.3 | **1.0** |
| Temperate | 2 | 5,465 | −5% | 144 | 0.0 | 0.0 | **0.0** |
| Cold | 3 | 1,615 | −72% | 33 | 0.0 | 0.0 | **0.0** |
| Undecided | 6 | 6,379 | +10% | 696 | 7.0 | 2.6 | **9.6** |
| Table in force | 54 cities | 5,774 | — | ≈642 | ≈11 (43 bin) | not counted | ≈11 |

Whole year, with spline. The 43 °C+ figures for the cooling months (Mar–Oct) are identical.
≈ BEE per-bin annual hours = bin hours × 5,774/1,600 (factor 3.609). Replace them with the exact Table 1 column from Schedule 19.

## Slide 5: 44 °C and above, with results for each option
I assumed the options are A = cap at 43 °C and drop 44 °C+, B = make the top bin "43 °C and above", and C = add a separate 44 °C+ bin. Relabel if the deck uses different options.

Hours per year per station, zone average. The bracketed figure is the same hours on BEE's 1,600-hour scale. The table in force puts **3** of its 1,600 hours in the 43 °C bin.

| Zone | Today: 43–44 / 44+ | Option A | **Option B** | Option C |
|---|---|---|---|---|
| Hot-Dry | 20.8 / 26.0 | 43 bin 20.8 (5.3); 26.0 h lost | **43+ bin 46.8 (12.0), +125%** | 43 bin 20.8 (5.3) + 44+ bin 26.0 (6.7) |
| Composite | 15.4 / 12.2 | 43 bin 15.4 (4.4); 12.2 h lost | **43+ bin 27.6 (7.9), +79%** | 43 bin 15.4 (4.4) + 44+ bin 12.2 (3.5) |
| Undecided | 7.0 / 2.6 | 43 bin 7.0 (1.8); 2.6 h lost | **43+ bin 9.6 (2.4), +37%** | 43 bin 7.0 (1.8) + 44+ bin 2.6 (0.7) |
| Warm-Humid | 0.7 / 0.3 | 0.7 (0.2) | 1.0 (0.2) | 44+ bin 0.3 (0.1) |
| Temperate, Cold | 0 / 0 | no change | no change | no change |

What this means for each option:
- **Option A** keeps the table shape but drops the hottest hours, when ACs work hardest.
- **Option B** keeps the 20 bins and the formula. The top bin in Hot-Dry goes from 5.3 to 12.0 per 1,600 hours, 4× the 3 in force.
- **Option C** is the most precise, but it adds a 21st bin and means changing the ISEER calculation and the IS table.

## Last slide: assumptions (the first two old items were facts) and next steps
**Assumptions**
1. **Smooth daily curve.** Temperature between the 3-hourly readings follows a smooth daily cycle. The spline was tested at 0.94 °C mean error, with no bias at 14:30.
2. **The recent decade represents today's climate.** No adjustment is made for the warming trend, which would add hot hours.
3. **Missing days look like recorded days.** A station's missing days in a month are assumed to be similar to its recorded days in that month (the month-by-month scaling depends on this).
4. **Stations stand for their zone.** Each station counts equally in its zone average, with no weighting by population or AC stock.
5. **Station air is the air at the AC.** Airport and observatory readings are taken to represent the air at outdoor units. Urban heat and rooftop effects are not captured, so hot hours in dense cities are probably understated.

**Next steps**
- Hold an expert consultation on the method (BEE, BIS, ISHRAE, IMD): spline or not, bin labels, the months covered, the 44 °C option and zone weighting.
- Fill the data gaps: request IMD hourly AWS data and daily Tmax for the weak stations, and cross-check against hourly airport (METAR) records.
- Zones and stations that need more readings:
  1. **Temperate:** only Bengaluru (Tier C) and Coimbatore (Tier B, not on the ECBC list).
  2. **Cold:** Srinagar, Sundernagar and Shillong only. There is no station from Ladakh, the Uttarakhand hills or Sikkim.
  3. **Hot-Dry Tier C:** Ajmer, Aurangabad, Deesa, Solapur and Udaipur. This is the zone that sets the 43 °C+ hours.
  4. **Central India heat belt, Tier C:** Delhi (Palam), Gwalior, Jabalpur, Raipur and Bhopal, all large markets.
  5. **Undecided Deccan:** Gulbarga and Anantapur (Tier C). More readings are needed to settle the zone.

  Mumbai, Bengaluru and Delhi (Palam) are all Tier C. This matters if the zones are later weighted by AC stock.

## After Thank You: annex tables
**Rough comparison with 5,774 (whole year, with spline, hours per year per station)**

| Zone | Hours ≥24 °C | vs 5,774 | Hours ≥35 °C | vs ≈642 | Share ≥35 °C |
|---|---:|---:|---:|---:|---:|
| Composite (42) | 5,598 | −3% | 692 | +8% | 12.4% |
| Hot-Dry (16) | 6,222 | +8% | 1,019 | +59% | 16.4% |
| Warm-Humid (59) | 7,226 | +25% | 256 | −60% | 3.5% |
| Temperate (2) | 5,465 | −5% | 144 | −78% | 2.6% |
| Cold (3) | 1,615 | −72% | 33 | −95% | 2.1% |
| Undecided (6) | 6,379 | +10% | 696 | +8% | 10.9% |
| All 128, simple mean* | 6,368 | +10% | 508 | −21% | 8.0% |
| Table in force | 5,774 | — | ≈642 | — | 11.1% |

\*This counts each station once and is not a national weighting. 59 of the 128 stations are Warm-Humid, which pulls total hours up and hot hours down. The zone-weighting decision therefore drives the national table.

**Pending: the per-bin tables.** For the 10-year counted hours by zone per bin (24, 25 … 43, 44+, total) and the per-bin comparison with BEE's 5,774, I need `IMD_BinHours_by_Zone_2_WITH_SPLINE.xlsx`. Its zone sheets carry the "record totals" (unscaled hours over every usable day), which sum directly to these tables.

## Sources
- BEE Schedule 19 (29 June 2015), Table 1; BEE Schedule 24 (8 Sept 2021), Table 1: bin hours 146, 163, 177, 183, 167, 150, 125, 104, 86, 67, 54, 46, 36, 28, 22, 16, 12, 9, 6, 3 (sum 1,600). https://beeindia.gov.in/WriteReadData/RTF1984/RTF-PDF-fdced7c20445fece_1777956042.pdf
- Energy Commission Malaysia, Guide on MEPS for Air Conditioners (2021): CSPF bins per ISO 16358-1 Table 3. https://www.st.gov.my/contents/2021/MEPS/20210108%20-Guide%20on%20MEPS%20for%20AC%20(UPDATED).pdf
- LBNL, Enhancing the Energy Efficiency of Room ACs in Malaysia (2025). https://eta-publications.lbl.gov/sites/default/files/2025-01/malaysia_ac.pdf
- LBNL/U4E, Harmonizing Energy-Efficiency Standards for Room ACs in Southeast Asia (2021). https://united4efficiency.org/wp-content/uploads/2022/07/ASEAN-AC-EE-Harmonization_Final_Feb-2021.pdf
- ISO 16358-1:2013. https://cdn.standards.iteh.ai/samples/56467/90c463d2acea405e9eb3ee48f2d34b38/ISO-16358-1-2013.pdf
- SASO 2663:2021 (CSA Group summary). https://www.csagroup.org/global-certification-regulatory-update/saudi-arabia-saso-released-updated-saso-2663-standard/
- GB 21455 bins (2013 edition, via LBNL China AC study). https://escholarship.org/content/qt77k463qf/qt77k463qf.pdf
- AFM 88-29 Engineering Weather Data (1978), described in Oregon State EM 8364. https://ir.library.oregonstate.edu/downloads/2f75r842b
- Performance metrics for room air-conditioners, Buildings & Cities (2021). https://journal-buildingscities.org/articles/10.5334/bc.127

The GB 21455 and SASO bin ranges come from secondary summaries; this environment could not open the standards themselves. Check them before external use.
