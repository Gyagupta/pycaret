// Builds the add-on slides for the IMD bin-hour deck (BEE review).
// Numbers come from the methodology note "Zone Bin-Hour Methodology"
// (IMD 2016-2025, whole year, with spline, zone averages) and BEE Schedule 19 Table 1.
// Run: NODE_PATH=<dir with pptxgenjs> node build_addon_slides.js
const pptxgen = require("pptxgenjs");
const path = require("path");

const NAVY = "1F3864", TEAL = "1F6F78", INK = "222222", MUTED = "5A6670",
  RULE = "C9D3DA", BAND = "EEF2F5", WHITE = "FFFFFF", HIGHLIGHT = "FCEBD6", WARN = "8A5212";
const FONT = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.title = "IMD bin-hours by zone: add-on slides for BEE review";
pres.company = "CLASP";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: WHITE },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 12.3, h: 0.75,
      fontFace: FONT, fontSize: 24, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "IMD bin-hours by climatic zone · draft for discussion with BEE, BIS and ISHRAE",
      options: { x: 0.5, y: 7.05, w: 9, h: 0.3, fontFace: FONT, fontSize: 9, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.5, h: 0.3, fontFace: FONT, fontSize: 9, color: MUTED, align: "right" },
});

const hdr = (t, extra = {}) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 12, valign: "middle", ...extra } });
const cell = (t, extra = {}) => ({ text: t, options: { fontSize: 12, color: INK, valign: "middle", ...extra } });
const num = (t, extra = {}) => cell(t, { align: "right", ...extra });
const tableOpts = (o) => ({ fontFace: FONT, border: { type: "solid", pt: 0.75, color: RULE }, margin: [3, 6, 3, 6], ...o });
const subtitle = (s, text, y = 1.05) => s.addText(text, { x: 0.5, y, w: 12.3, h: 0.4, fontFace: FONT, fontSize: 14,
  color: MUTED, italic: true, margin: 0, isTextBox: true });
const sectionHead = (s, text, x, y, w) => s.addText(text, { x, y, w, h: 0.4, fontFace: FONT, fontSize: 16, bold: true,
  color: TEAL, margin: 0, isTextBox: true });

// ---------- 1. Slide 2 update: basis of selection + IMD data challenges ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("How the 128 stations were chosen, and IMD data limits", { placeholder: "title" });
  subtitle(s, "Replaces the 'challenges' box on slide 2. The six Undecided stations stay in the zone table; they are a zoning question, not a data gap.");

  sectionHead(s, "Basis of selection (449 station codes → 128 used)", 0.5, 1.6, 6.3);
  const rows = [
    [hdr("Test"), hdr("Rule"), hdr("Result")],
    [cell("Complete day", { bold: true }), cell("All 8 three-hourly readings present"), cell("85.8% of station-days")],
    [cell("Usable season", { bold: true }), cell("1 Mar–31 Oct: ≥90% complete days, no gap >15 days"), cell("Counted per year, 2016–25")],
    [cell("Tier", { bold: true }), cell("Usable seasons of 10: A 8–10 · B 5–7 · C 1–4 · D 0"), cell("A 43 · B 56 · C 29 used; D 321 dropped")],
    [cell("Monthly coverage", { bold: true }), cell("≥28 usable days in every calendar month over 10 years"), cell("All 128 pass (with spline)")],
    [cell("Zone tag", { bold: true }), cell("ECBC 2017 list · NBC 2016 regional rule · project precedent · undecided"), cell("43 · 78 · 1 · 6")],
  ];
  s.addTable(rows, tableOpts({ x: 0.5, y: 2.05, w: 6.3, colW: [1.45, 2.95, 1.9], fontSize: 12 }));
  s.addText("Same rule as the earlier notebook: all 449 stations reproduce the same tier.", {
    x: 0.5, y: 5.75, w: 6.3, h: 0.35, fontFace: FONT, fontSize: 12, color: MUTED, italic: true, margin: 0, isTextBox: true });

  sectionHead(s, "Data challenges in the IMD record", 7.2, 1.6, 5.6);
  const ch = [
    ["Only 8 readings a day. ", "No hourly record; the true daily peak can fall between the 14:30 and 17:30 IST readings."],
    ["Most stations unusable. ", "321 of 449 station codes (71%) have no usable cooling season in 2016–2025."],
    ["Gaps inside the 128. ", "14.2% of station-days are incomplete: 4.5% fillable (1–2 readings), 2.7% 9-hour gaps, 7.0% three or more missing."],
    ["Uneven coverage. ", "Warm-Humid has 59 stations; Temperate 2 and Cold 3. 29 stations are Tier C."],
    ["Cleaning needed. ", "4,209 blank rows, repeated readings and impossible dates removed (5.31 M rows → 5.305 M valid readings)."],
    ["No zone in IMD metadata. ", "Each station was placed using ECBC 2017 or NBC 2016."],
  ];
  const runs = [];
  ch.forEach(([b, t], i) => {
    runs.push({ text: b, options: { bold: true, color: NAVY, bullet: true } });
    runs.push({ text: t, options: { color: INK, breakLine: i < ch.length - 1, paraSpaceAfter: 6 } });
  });
  s.addText(runs, { x: 7.2, y: 2.05, w: 5.6, h: 4.6, fontFace: FONT, fontSize: 13, valign: "top", margin: [0, 0, 0, 0], isTextBox: true });
  s.addNotes("Source: methodology note sections 'Source data' and 'Which stations are used'. Station-day split is for the 128 stations, 434,165 station-days.");
}

// ---------- 2. Slide 3 update: plain outcome line + comparable methods side panel ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Method: how 3-hourly IMD readings become hours in each bin", { placeholder: "title" });

  sectionHead(s, "What we do", 0.5, 1.25, 5.4);
  const steps = [
    ["Method A, no spline: ", "each IMD reading counts for 3 hours. Nothing is filled in."],
    ["Method B, spline: ", "a smooth daily curve through the 8 readings gives 24 hourly values. Days missing 1–2 non-adjacent readings are filled from that day's own curve."],
    ["Tested: ", "hiding one real reading and predicting it gave the lowest error of 5 methods (0.94 °C) and almost no bias at 14:30 IST."],
  ];
  const r = [];
  steps.forEach(([b, t], i) => {
    r.push({ text: b, options: { bold: true, color: NAVY, bullet: true } });
    r.push({ text: t, options: { color: INK, breakLine: i < steps.length - 1, paraSpaceAfter: 8 } });
  });
  s.addText(r, { x: 0.5, y: 1.7, w: 5.4, h: 2.9, fontFace: FONT, fontSize: 14, valign: "top", margin: 0, isTextBox: true });

  // replacement for the "14:30 hottest reading" line
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.75, w: 5.4, h: 1.75, fill: { color: HIGHLIGHT },
    line: { color: HIGHLIGHT }, rectRadius: 0.08, objectName: "outcome-box" });
  s.addText([
    { text: "Outcome", options: { bold: true, color: WARN, fontSize: 12, breakLine: true } },
    { text: "IMD takes a reading at 14:30 IST, close to the hottest time of day, so the hot-hour bins rest on real measurements. The spline only fills the hours between readings.",
      options: { color: INK, fontSize: 14 } },
  ], { x: 0.7, y: 4.85, w: 5.0, h: 1.55, fontFace: FONT, valign: "top", margin: 0, isTextBox: true });

  // side panel: comparable methods
  s.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 1.25, w: 6.55, h: 5.6, fill: { color: BAND }, line: { color: BAND }, objectName: "compare-panel" });
  s.addText("Who else builds bin-hour tables, and how", { x: 6.5, y: 1.35, w: 6.2, h: 0.4, fontFace: FONT, fontSize: 15,
    bold: true, color: TEAL, margin: 0, isTextBox: true });
  const cmp = [
    [hdr("Method", { fontSize: 11 }), hdr("Same as ours", { fontSize: 11 }), hdr("Different", { fontSize: 11 })],
    [cell("India, BEE Sch. 19 / IS 1391 (in force)", { bold: true, fontSize: 11 }), cell("1 °C bins from 24 °C; whole-year hours (5,774 → 1,600)", { fontSize: 11 }), cell("One typical year per city (ISHRAE 2014, 54 cities); one national table; stops at 43 °C", { fontSize: 11 })],
    [cell("ISO 16358-1 (Malaysia, Thailand, Indonesia, Philippines, Vietnam)", { bold: true, fontSize: 11 }), cell("Same bin-hour formula ISEER uses", { fontSize: 11 }), cell("Fixed reference table, 21–35 °C, 1,817 h; Malaysia uses it unchanged, not local data", { fontSize: 11 })],
    [cell("Saudi Arabia, SASO 2663:2021", { bold: true, fontSize: 11 }), cell("National weather data replaces ISO default bins; covers very hot hours", { fontSize: 11 }), cell("One national table", { fontSize: 11 })],
    [cell("China, GB 21455", { bold: true, fontSize: 11 }), cell("1 °C bins starting at 24 °C", { fontSize: 11 }), cell("Top bin 38 °C (2013 edition); one national table", { fontSize: 11 })],
    [cell("US Engineering Weather Data (AFM 88-29)", { bold: true, fontSize: 11 }), cell("3-hourly readings counted as 3 hours each, as in our Method A", { fontSize: 11 }), cell("No interpolation; 5 °F bins", { fontSize: 11 })],
  ];
  s.addTable(cmp, tableOpts({ x: 6.5, y: 1.85, w: 6.15, colW: [2.0, 2.05, 2.1], fill: { color: WHITE } }));
  s.addText("No other scheme we found builds zone-wise tables from a 10-year raw station record; most use one national table or typical-year files.", {
    x: 6.5, y: 6.25, w: 6.15, h: 0.5, fontFace: FONT, fontSize: 11, italic: true, color: MUTED, margin: 0, isTextBox: true });
  s.addNotes("Sources: BEE Schedule 19 (2015) Table 1; ISO 16358-1:2013; Energy Commission Malaysia (Suruhanjaya Tenaga) Guide on MEPS for AC (2021): bins per ISO 16358-1 Table 3; LBNL/U4E ASEAN AC harmonization report (2021); SASO 2663:2021; GB 21455-2013/2019; AFM 88-29 Engineering Weather Data (1978). Verify the GB and SASO bin ranges against the standard text before external use.");
}

// ---------- 3. Slide 4 update: zone table with 43 °C+ column and clear units ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Hours per year in each zone, compared with the table in force", { placeholder: "title" });
  subtitle(s, "Each figure = hours in ONE typical year at ONE station, averaged over the zone's stations. Built from 10 years of IMD readings (2016–2025); not a 10-year total. Whole year, with spline.");

  const H = (t) => hdr(t, { align: "center" });
  const rows = [
    [H("Zone"), H("Stations"), H("All hours ≥24 °C"), H("vs 5,774"), H("Hours ≥35 °C"), H("43–44 °C"), H("44 °C and above"), H("43 °C and above (new)")],
    ["Composite", 42, "5,598", "−3%", "692", "15.4", "12.2", "27.6"],
    ["Hot-Dry", 16, "6,222", "+8%", "1,019", "20.8", "26.0", "46.8"],
    ["Warm-Humid", 59, "7,226", "+25%", "256", "0.7", "0.3", "1.0"],
    ["Temperate", 2, "5,465", "−5%", "144", "0.0", "0.0", "0.0"],
    ["Cold", 3, "1,615", "−72%", "33", "0.0", "0.0", "0.0"],
    ["Undecided", 6, "6,379", "+10%", "696", "7.0", "2.6", "9.6"],
  ].map((r, i) => i === 0 ? r : r.map((v, j) => j === 0 ? cell(String(v), { bold: true }) :
      num(String(v), j === 7 ? { bold: true, fill: { color: HIGHLIGHT } } : {})));
  rows.push([cell("Table in force (BEE Sch. 19)", { bold: true, fill: { color: BAND } }), num("54 cities", { fill: { color: BAND } }),
    num("5,774", { fill: { color: BAND } }), num("—", { fill: { color: BAND } }), num("≈642", { fill: { color: BAND } }),
    num("≈11 (43 bin)", { fill: { color: BAND } }), num("not counted", { fill: { color: BAND } }), num("≈11", { fill: { color: BAND, bold: true } })]);
  s.addTable(rows, tableOpts({ x: 0.5, y: 1.7, w: 12.3, colW: [2.6, 1.1, 1.6, 1.1, 1.5, 1.3, 1.5, 1.6], fontSize: 13, rowH: 0.42 }));

  s.addText([
    { text: "Reading the table: ", options: { bold: true, color: NAVY } },
    { text: "a Hot-Dry station has about 47 hours a year at 43 °C or above, four times the ≈11 hours in the table in force. Cooling-months (Mar–Oct) figures for 43 °C and above are identical, since all such hours fall in those months.", options: { color: INK } },
  ], { x: 0.5, y: 5.45, w: 12.3, h: 0.75, fontFace: FONT, fontSize: 13, margin: 0, valign: "top", isTextBox: true });
  s.addText("≈ BEE per-bin annual hours derived as bin hours × 5,774/1,600 (e.g. 43 °C: 3 × 3.609 ≈ 11). Replace with the exact Schedule 19 Table 1 column before circulation.", {
    x: 0.5, y: 6.35, w: 12.3, h: 0.45, fontFace: FONT, fontSize: 10, color: MUTED, italic: true, margin: 0, isTextBox: true });
}

// ---------- 4. Slide 5 update: 44 °C and above options with results ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Hours at 44 °C and above: results under each option", { placeholder: "title" });
  subtitle(s, "Hours per year per station, zone average (whole year, with spline). Bracket: the same hours on BEE's 1,600-hour scale; the table in force puts 3 of 1,600 hours in its 43 °C bin.");

  const H = (t) => hdr(t, { align: "center" });
  const rows = [
    [H("Zone"), H("Data today: 43–44 °C | 44 °C+"), H("Option A: keep 24–43 °C, drop 44 °C+"), H("Option B: top bin = 43 °C and above"), H("Option C: add a separate 44 °C+ bin")],
    ["Hot-Dry", "20.8 | 26.0", "43 bin 20.8 (5.3)\n26.0 h lost", "43+ bin 46.8 (12.0)\n+125%", "43 bin 20.8 (5.3)\n44+ bin 26.0 (6.7)"],
    ["Composite", "15.4 | 12.2", "43 bin 15.4 (4.4)\n12.2 h lost", "43+ bin 27.6 (7.9)\n+79%", "43 bin 15.4 (4.4)\n44+ bin 12.2 (3.5)"],
    ["Undecided", "7.0 | 2.6", "43 bin 7.0 (1.8)\n2.6 h lost", "43+ bin 9.6 (2.4)\n+37%", "43 bin 7.0 (1.8)\n44+ bin 2.6 (0.7)"],
    ["Warm-Humid", "0.7 | 0.3", "43 bin 0.7 (0.2)", "43+ bin 1.0 (0.2)", "44+ bin 0.3 (0.1)"],
    ["Temperate, Cold", "0 | 0", "No change", "No change", "No change"],
  ].map((r, i) => i === 0 ? r : r.map((v, j) => j === 0 ? cell(v, { bold: true }) :
      cell(v, { align: "center", ...(j === 3 && i <= 3 ? { fill: { color: HIGHLIGHT }, bold: true } : {}) })));
  s.addTable(rows, tableOpts({ x: 0.5, y: 1.7, w: 12.3, colW: [2.0, 2.3, 2.6, 2.7, 2.7], fontSize: 13 }));

  // implication cards
  const cards = [
    ["Option A", "Same table shape as today, but undercounts the hottest hours where ACs work hardest (26 h/yr lost in Hot-Dry)."],
    ["Option B", "Keeps 20 bins and the ISEER formula; the top bin in Hot-Dry grows from 5.3 to 12.0 of 1,600 hours, 4× the 3 in force."],
    ["Option C", "Most precise, but adds a 21st bin and a change to the ISEER calculation and the IS 1391 / IS 18154 table."],
  ];
  cards.forEach(([t, d], i) => {
    const x = 0.5 + i * 4.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.55, w: 3.9, h: 1.65, fill: { color: BAND }, line: { color: BAND }, rectRadius: 0.08, objectName: `option-card-${i}` });
    s.addText([{ text: t, options: { bold: true, color: TEAL, fontSize: 14, breakLine: true } },
      { text: d, options: { color: INK, fontSize: 12 } }], { x: x + 0.15, y: 4.63, w: 3.6, h: 1.5, fontFace: FONT, valign: "top", margin: 0, isTextBox: true });
  });
  s.addNotes("Bracketed values = zone hours scaled to 1,600 using that zone's own total hours ≥24 °C (e.g. Hot-Dry 46.8/6,222 × 1,600 = 12.0). If the deck's options are labelled differently, map them to: A = cap at 43 and drop, B = merge into top bin, C = new 44+ bin.");
}

// ---------- 5. Last content slide: assumptions + next steps ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Assumptions and next steps", { placeholder: "title" });

  sectionHead(s, "Assumptions behind the numbers", 0.5, 1.2, 6.0);
  const as = [
    ["Smooth daily curve. ", "Temperature between IMD's 3-hourly readings follows a smooth daily cycle (spline tested: 0.94 °C mean error, no bias at 14:30)."],
    ["Recent decade = today's climate. ", "2016–2025 represents current conditions; no adjustment for the warming trend, which would add hot hours."],
    ["Missing days look like recorded days. ", "A station's missing days in a month are assumed similar to its recorded days in that month."],
    ["Stations stand for their zone. ", "Each station counts equally in its zone average; no weighting by population or AC stock."],
    ["Station air = air at the AC. ", "Airport and observatory readings represent outdoor-unit air; urban heat and rooftop effects are not captured."],
  ];
  const ra = [];
  as.forEach(([b, t], i) => {
    ra.push({ text: `${i + 1}. ` + b, options: { bold: true, color: NAVY } });
    ra.push({ text: t, options: { color: INK, breakLine: i < as.length - 1, paraSpaceAfter: 8 } });
  });
  s.addText(ra, { x: 0.5, y: 1.65, w: 6.0, h: 5.2, fontFace: FONT, fontSize: 13, valign: "top", margin: 0, isTextBox: true });

  sectionHead(s, "Next steps", 6.9, 1.2, 5.9);
  s.addText([
    { text: "Expert consultation on the method ", options: { bold: true, color: NAVY, bullet: true } },
    { text: "(BEE, BIS, ISHRAE, IMD): spline or not, bin labels, months, the 44 °C option and zone weighting.", options: { color: INK, breakLine: true, paraSpaceAfter: 6 } },
    { text: "Fill the data gaps: ", options: { bold: true, color: NAVY, bullet: true } },
    { text: "request IMD hourly AWS data and daily maximum temperatures for weak stations; cross-check with hourly airport (METAR) records.", options: { color: INK } },
  ], { x: 6.9, y: 1.65, w: 5.9, h: 1.85, fontFace: FONT, fontSize: 13, valign: "top", margin: 0, isTextBox: true });

  const need = [
    [hdr("Where we need more readings", { fontSize: 12 }), hdr("Why", { fontSize: 12 })],
    [cell("Temperate zone", { bold: true }), cell("Only 2 stations: Bengaluru (Tier C) and Coimbatore (zone not on ECBC list)")],
    [cell("Cold zone", { bold: true }), cell("3 stations (Srinagar, Sundernagar, Shillong); none from Ladakh, Uttarakhand hills or Sikkim")],
    [cell("Hot-Dry, Tier C", { bold: true }), cell("Ajmer, Aurangabad, Deesa, Solapur, Udaipur: this zone sets the 43 °C+ hours")],
    [cell("Central India heat belt", { bold: true }), cell("Delhi (Palam), Gwalior, Jabalpur, Raipur, Bhopal: large markets, Tier C")],
    [cell("Deccan, Undecided", { bold: true }), cell("Gulbarga, Anantapur (Tier C): more stations needed to settle the zone")],
  ];
  s.addTable(need, tableOpts({ x: 6.9, y: 3.6, w: 5.9, colW: [2.0, 3.9], fontSize: 11 }));
}

// ---------- 6. Annex (after Thank you): rough result vs 5,774 ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Annex: first result against the table in force (5,774 hours)", { placeholder: "title" });
  subtitle(s, "Rough comparison, whole year, with spline. Hours per year per station, zone average. A national table still needs a zone-weighting decision.");

  const H = (t) => hdr(t, { align: "center", fontSize: 11 });
  const rows = [
    [H("Zone"), H("Hours ≥24 °C"), H("vs 5,774"), H("Hours ≥35 °C"), H("vs ≈642"), H("Share ≥35 °C")],
    ["Composite (42)", "5,598", "−3%", "692", "+8%", "12.4%"],
    ["Hot-Dry (16)", "6,222", "+8%", "1,019", "+59%", "16.4%"],
    ["Warm-Humid (59)", "7,226", "+25%", "256", "−60%", "3.5%"],
    ["Temperate (2)", "5,465", "−5%", "144", "−78%", "2.6%"],
    ["Cold (3)", "1,615", "−72%", "33", "−95%", "2.1%"],
    ["Undecided (6)", "6,379", "+10%", "696", "+8%", "10.9%"],
    ["All 128, simple mean*", "6,368", "+10%", "508", "−21%", "8.0%"],
    ["Table in force", "5,774", "—", "≈642", "—", "11.1%"],
  ].map((r, i) => i === 0 ? r : r.map((v, j) => {
    const shade = i >= 7 ? { fill: { color: BAND } } : {};
    return j === 0 ? cell(v, { bold: true, fontSize: 12, ...shade }) : num(v, { fontSize: 12, ...shade });
  }));
  s.addTable(rows, tableOpts({ x: 0.5, y: 1.7, w: 7.4, colW: [2.1, 1.1, 0.95, 1.15, 1.0, 1.1] }));
  s.addText("* Each of the 128 stations counted once. Not a national weighting: 59 of 128 stations are Warm-Humid, which pulls total hours up and hot hours down.", {
    x: 0.5, y: 5.7, w: 7.4, h: 0.6, fontFace: FONT, fontSize: 10, italic: true, color: MUTED, margin: 0, isTextBox: true });

  s.addChart(pres.charts.BAR, [{ name: "Hours ≥35 °C per year",
    labels: ["Cold", "Temperate", "Warm-Humid", "BEE table (in force)", "Composite", "Undecided", "Hot-Dry"],
    values: [33, 144, 256, 642, 692, 696, 1019] }], {
    x: 8.2, y: 1.6, w: 4.65, h: 4.9, barDir: "bar", showTitle: true, title: "Hours at 35 °C and above, per year",
    titleFontSize: 13, titleColor: NAVY, titleFontFace: FONT, chartColors: [TEAL, TEAL, TEAL, "9AA5AD", TEAL, TEAL, TEAL],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 10, dataLabelColor: INK, dataLabelFontFace: FONT,
    catAxisLabelColor: INK, catAxisLabelFontSize: 11, catAxisLabelFontFace: FONT, valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 45,
  });
  s.addText("Takeaway: totals are within 8% of 5,774 in Composite and Hot-Dry, but the hot tail is heavier in Hot-Dry. Which zones carry how much weight decides the national table.", {
    x: 0.5, y: 6.35, w: 12.3, h: 0.55, fontFace: FONT, fontSize: 13, bold: true, color: NAVY, margin: 0, isTextBox: true });
  s.addNotes("Per-bin version (24 to 43 and 44+, 10-year counted hours by zone) to follow once IMD_BinHours_by_Zone_2_WITH_SPLINE.xlsx is available in this workspace. BEE ≥35 °C annual hours derived as 178 of 1,600 bin hours × 5,774/1,600 = 642.");
}

const out = path.join(__dirname, "BinHour_addon_slides.pptx");
pres.writeFile({ fileName: out }).then((f) => console.log("wrote", f));
