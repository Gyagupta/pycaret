"""Turn IMD_BinHours_CLASP_FINAL_v2.pptx into v3.

Changes:
  * slide 6 becomes 'population' slide: share of stations vs share of people (native bar chart), A/B table kept
  * new slide 7: proposed table vs today's table, side by side (two tables)
  * new slide after 'Should the table count 44 °C and above?': hours at 40 °C and above, degree by degree
  * annex cross-references 'slide 4' corrected to 'slide 5' (the average-year table)
Every number comes from deck_data.json (extracted from the v2 deck). Tables, panels, labels and the navy
bar are cloned from the deck's own shapes so fonts, colours and spacing match.

Usage: python -I build_deck_v3.py <v2.pptx> <deck_data.json> <out.pptx>
"""
import copy
import json
import sys
from xml.sax.saxutils import escape

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.util import Pt

SRC, DATA, OUT = sys.argv[1:4]
d = json.load(open(DATA))
prs = Presentation(SRC)

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
EMU = 914400
NAVY, INK, MUTED, BORDER, GREY, PEACH, WHITE = "084159", "474C55", "6B7078", "D6DEE3", "F1F4F5", "FDEFD9", "FFFFFF"
ORANGE_CHART, TEAL_CHART = "CC8049", "308589"


def q(tag):
    pre, t = tag.split(":")
    return "{%s}%s" % (NS[pre], t)


def emu(inches):
    return str(int(round(inches * EMU)))


def shape(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    raise KeyError(name)


# ------------------------------------------------------------------ source shapes (captured before any edit)
S5, S6, S8 = prs.slides[4], prs.slides[5], prs.slides[7]
T = {
    "subtitle": copy.deepcopy(shape(S6, "TextBox 2")._element),
    "label": copy.deepcopy(shape(S6, "TextBox 3")._element),
    "table": copy.deepcopy(shape(S6, "Table 6")._element),
    "panel_box": copy.deepcopy(shape(S8, "Rounded Rectangle 10")._element),
    "panel_text": copy.deepcopy(shape(S8, "TextBox 11")._element),
    "bar_box": copy.deepcopy(shape(S6, "Rounded Rectangle 9")._element),
    "bar_text": copy.deepcopy(shape(S6, "TextBox 10")._element),
    "foot": copy.deepcopy(shape(S6, "TextBox 11")._element),
    "caps": copy.deepcopy(shape(S8, "TextBox 8")._element),
    "sublabel": copy.deepcopy(shape(S8, "TextBox 4")._element),
    "small": copy.deepcopy(shape(S8, "TextBox 7")._element),
}


def run_props(el):
    """rPr of every run in a template text shape, in order."""
    return [copy.deepcopy(r.find(q("a:rPr"))) for r in el.iter(q("a:r"))]


def para_props(el):
    out = []
    for p in el.iter(q("a:p")):
        ppr = p.find(q("a:pPr"))
        out.append(copy.deepcopy(ppr) if ppr is not None else None)
    return out


R = {k: run_props(v) for k, v in T.items() if k not in ("table", "panel_box", "bar_box")}
PP = {k: para_props(v) for k, v in T.items() if k not in ("table", "panel_box", "bar_box")}
ROLE = {
    "subtitle": R["subtitle"][0],          # 16 pt bold navy
    "label": R["label"][0],                # 13 pt bold navy
    "caps": R["caps"][0],                  # 11 pt bold brown caps
    "panel_head": R["panel_text"][0],      # 12 pt bold navy
    "panel_body": R["panel_text"][1],      # 10.5 pt ink
    "panel_bold": R["panel_text"][4],      # 10.5 pt bold ink
    "bar_lead": R["bar_text"][0],          # 11 pt bold orange
    "bar_text": R["bar_text"][1],          # 11 pt white
    "foot": R["foot"][0],                  # 8 pt muted
    "sub_bold": R["sublabel"][0],          # 10.5 pt bold navy
    "sub_text": R["sublabel"][1],          # 10.5 pt ink
}


def rewrite(el, paras, ppr_from=None):
    """Replace all paragraphs of a text shape. paras = [[(text, role), ...], ...]."""
    body = el.find(".//" + q("p:txBody"))
    old = body.findall(q("a:p"))
    pprs = PP[ppr_from] if ppr_from else [p.find(q("a:pPr")) for p in old]
    for p in old:
        body.remove(p)
    for i, runs in enumerate(paras):
        p = etree.SubElement(body, q("a:p"))
        tpl = pprs[min(i, len(pprs) - 1)] if pprs else None
        if tpl is not None:
            p.append(copy.deepcopy(tpl))
        for text, role in runs:
            r = etree.SubElement(p, q("a:r"))
            r.append(copy.deepcopy(ROLE[role]))
            t = etree.SubElement(r, q("a:t"))
            t.text = text
    return el


def place(slide, el, x=None, y=None, w=None, h=None, name=None):
    """Insert a copy of a shape element on a slide with a fresh id, optionally moved/resized."""
    el = copy.deepcopy(el)
    nv = el.find(".//" + q("p:cNvPr"))
    nv.set("id", str(slide.shapes._next_shape_id))
    if name:
        nv.set("name", name)
    off = el.find(".//" + q("a:off"))
    ext = el.find(".//" + q("a:ext"))
    if x is not None:
        off.set("x", emu(x))
    if y is not None:
        off.set("y", emu(y))
    if w is not None:
        ext.set("cx", emu(w))
    if h is not None:
        ext.set("cy", emu(h))
    slide.shapes._spTree.append(el)
    return el


# ------------------------------------------------------------------ tables in the deck's own cell style
LINES = "".join(f'<a:{s} w="9525"><a:solidFill><a:srgbClr val="{BORDER}"/></a:solidFill></a:{s}>'
                for s in ("lnL", "lnR", "lnT", "lnB"))


def cell(runs, fill, algn, sz):
    """runs: list of paragraphs, each a list of (text, bold, colour)."""
    paras = ""
    for para in runs:
        rs = "".join(
            f'<a:r><a:rPr lang="en-IN" sz="{sz}" b="{1 if b else 0}" dirty="0"><a:solidFill><a:srgbClr val="{c}"/>'
            f'</a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>{escape(t)}</a:t></a:r>'
            for t, b, c in para)
        paras += f'<a:p><a:pPr algn="{algn}"/>{rs}</a:p>'
    return (f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>{paras}</a:txBody>'
            f'<a:tcPr marL="64008" marR="64008" marT="18288" marB="18288" anchor="ctr">{LINES}'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></a:tcPr></a:tc>')


def build_table(slide, x, y, colw, header, rows, row_h=0.4, hdr_h=0.36, sz=1050, hdr_sz=950,
                value_algn="ctr", highlight=(), name="Table"):
    """header: list of str; rows: list of lists (str, or list-of-paragraphs of (text,bold) tuples).
    Rows alternate white/grey like the deck; rows whose index is in `highlight` get the peach fill
    with bold navy values."""
    xml = [f'<a:tbl xmlns:a="{NS["a"]}"><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>']
    xml += [f'<a:gridCol w="{emu(w)}"/>' for w in colw]
    xml.append(f'</a:tblGrid><a:tr h="{emu(hdr_h)}">')
    for j, h in enumerate(header):
        xml.append(cell([[(h, True, WHITE)]] if h else [[]], NAVY, "l" if j == 0 else "ctr", hdr_sz))
    xml.append("</a:tr>")
    for i, row in enumerate(rows):
        hi = i in highlight
        fill = PEACH if hi else (WHITE if i % 2 == 0 else GREY)
        xml.append(f'<a:tr h="{emu(row_h)}">')
        for j, v in enumerate(row):
            if isinstance(v, str):
                runs = [[(line, j == 0 or hi, NAVY if (j == 0 or hi) else INK)] for line in v.split("\n")]
            else:
                runs = [[(t, b, NAVY if b else INK) for t, b in para] for para in v]
            xml.append(cell(runs, fill, "l" if j == 0 else value_algn, sz))
        xml.append("</a:tr>")
    xml.append("</a:tbl>")
    gf = copy.deepcopy(T["table"])
    old = gf.find(".//" + q("a:tbl"))
    old.getparent().replace(old, etree.fromstring("".join(xml)))
    xfrm = gf.find(q("p:xfrm"))
    xfrm.find(q("a:off")).set("x", emu(x))
    xfrm.find(q("a:off")).set("y", emu(y))
    xfrm.find(q("a:ext")).set("cx", emu(sum(colw)))
    xfrm.find(q("a:ext")).set("cy", emu(hdr_h + row_h * len(rows)))
    nv = gf.find(".//" + q("p:cNvPr"))
    nv.set("id", str(slide.shapes._next_shape_id))
    nv.set("name", name)
    slide.shapes._spTree.append(gf)
    return gf


def new_slide_after(ref_slide, title):
    layout = ref_slide.slide_layout
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx != 0:
            ph._element.getparent().remove(ph._element)
    s.shapes.title.text = title
    lst = prs.slides._sldIdLst
    ids = list(lst)
    ref_idx = next(i for i, sid in enumerate(ids) if prs.part.related_part(sid.rId) is ref_slide.part)
    new_el = ids[-1]
    lst.remove(new_el)
    lst.insert(ref_idx + 1, new_el)
    return s


def style_chart(chart):
    chart.font.name = "Arial"
    chart.font.size = Pt(10)
    chart.font.color.rgb = RGBColor.from_string(INK)


# ------------------------------------------------------------------ numbers from the deck
def n(s):
    return float(str(s).replace(",", "").replace("\xa0", "").strip())


shares = []
for line in d["notes"]["s6"].splitlines():
    parts = line.split()
    if len(parts) >= 3 and parts[-1].endswith("%") and parts[-2].endswith("%"):
        shares.append((" ".join(parts[:-2]), n(parts[-2][:-1]) / 100, n(parts[-1][:-1]) / 100))
share = {z: (s, p) for z, s, p in shares}

ann = {r[0]: r[1:] for r in d["tables"]["s15_t1"][1:] + d["tables"]["s15_t2"][1:] if r[0] != "Bin °C"}
of1600 = {"A": {}, "B": {}, "T": {}}
hours = {"A": {}, "B": {}, "T": {}}
for b, v in ann.items():
    for key, (ih, i1600) in {"A": (0, 1), "B": (2, 3), "T": (4, 5)}.items():
        hours[key][b] = v[ih]
        of1600[key][b] = v[i1600]
BINS = [str(x) for x in range(24, 44)] + ["44+"]


def band(key, lo, hi):
    vals = [of1600[key][b] for b in BINS[BINS.index(lo):BINS.index(hi) + 1]]
    vals = [n(v) for v in vals if v.strip() not in ("–", "-", "")]
    return sum(vals) if vals else None


zt = {r[0]: r[1:] for r in d["tables"]["s16_t1"][1:] + d["tables"]["s17_t1"][1:]}
zhdr = [h.replace("\xa0", " ") for h in d["tables"]["s16_t1"][0]]
s6tab = {r[0].replace("\xa0", " "): r[1:] for r in d["tables"]["s6_t1"][1:]}
hot8 = list(zip(d["charts"]["s8_c1"]["categories"], d["charts"]["s8_c1"]["series"]["v"]))

# ------------------------------------------------------------------ slide 6 -> population slide (edited in place)
s6 = S6
rewrite(shape(s6, "TextBox 2")._element,
        [[(f"Composite: {share['Composite'][1]:.1%} of people, {share['Composite'][0]:.1%} of stations; "
          f"Warm-Humid: {share['Warm-Humid'][1]:.1%} of people, {share['Warm-Humid'][0]:.1%} of stations", "subtitle")]])
rewrite(shape(s6, "TextBox 3")._element, [[("Share of stations and share of people, by climate zone", "label")]])
old_chart = shape(s6, "Chart 4")
rid = old_chart._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart").get(q("r:id"))
old_chart._element.getparent().remove(old_chart._element)
s6.part.drop_rel(rid)

zones = ["Composite", "Warm-Humid", "Hot-Dry", "Undecided", "Temperate", "Cold"]
cd = CategoryChartData()
cd.categories = zones
cd.add_series("Share of people (A: by population)", [share[z][1] for z in zones])
cd.add_series("Share of stations (B: each station equal)", [share[z][0] for z in zones])
gf = s6.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, int(0.75 * EMU), int(2.40 * EMU), int(7.30 * EMU), int(3.45 * EMU), cd)
gf.name = "Chart: stations vs people"
ch = gf.chart
style_chart(ch)
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.TOP
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(10)
plot = ch.plots[0]
plot.gap_width = 55
plot.overlap = -5
for ser, col in zip(plot.series, (ORANGE_CHART, TEAL_CHART)):
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = RGBColor.from_string(col)
    ser.invert_if_negative = False
plot.has_data_labels = True
dl = plot.data_labels
dl.number_format = "0.0%"
dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END
dl.font.size = Pt(9.5)
dl.font.bold = True
dl.font.color.rgb = RGBColor.from_string(NAVY)
va = ch.value_axis
va.visible = False
va.has_major_gridlines = False
va.minimum_scale = 0
va.maximum_scale = 0.6
ca = ch.category_axis
ca.reverse_order = True
ca.has_major_gridlines = False
ca.tick_labels.font.size = Pt(10.5)
ca.tick_labels.font.color.rgb = RGBColor.from_string(INK)
ca.format.line.color.rgb = RGBColor.from_string(BORDER)

rewrite(shape(s6, "TextBox 10")._element, [[
    ("What weighting by people does: ", "bar_lead"),
    (f"Composite counts for more, so India has fewer warm hours ({s6tab['Hours at 24 °C or more, a year'][0]} vs "
     f"{s6tab['Hours at 24 °C or more, a year'][1]} a year) but more hot ones "
     f"({s6tab['Hours at 35 °C or more, a year'][0]} vs {s6tab['Hours at 35 °C or more, a year'][1]} a year at 35 °C or more).",
     "bar_text")]])
s6.notes_slide.notes_text_frame.text = (
    "Share of stations = stations in the zone ÷ 129. Share of people = each state's projected 2026 population "
    "(National Commission on Population, MoHFW, 2019), shared among its stations by Census 2011 district population, "
    "added by zone.\n" + d["notes"]["s6"])

# ------------------------------------------------------------------ new slide 7: proposed vs today
s7 = new_slide_after(s6, "Proposed table and today's table, side by side")
place(s7, rewrite(copy.deepcopy(T["subtitle"]), [[("Same 1 °C bins and 1,600-hour basis; new data, a written method and one more bin", "subtitle")]]),
      name="Subtitle")
place(s7, rewrite(copy.deepcopy(T["label"]), [[("What changes", "label")]]), x=0.85, y=2.10, w=7.20, name="Label left")
build_table(
    s7, 0.85, 2.42, [1.70, 2.35, 3.15],
    ["", "Today (IS 18154)", "Proposed (IMD 2016–2025)"],
    [
        ["Data", "54 cities; years not published", "129 IMD stations with full-day readings, 10 years"],
        ["How hours are counted", "Method not published", "8 readings a day → smooth curve → 24 hourly values"],
        ["Bins", "1 °C, 24 to 43 °C", "1 °C, 24 to 43 °C, plus '44 °C and above'"],
        ["Day and night", "Counted (5,774 h a year is more than daytime alone holds)", "Counted, all 24 hours"],
        ["How places are combined", "Not published", "A: by population, or B: each station equal"],
        ["Zone tables", "None", "6 zones plus India"],
    ],
    row_h=0.43, sz=1000, value_algn="l", name="Table: what changes")
place(s7, rewrite(copy.deepcopy(T["caps"]), [[("ON THE 1,600-HOUR BASIS", "caps")]]), x=8.30, y=2.10, w=4.18, name="Label right")


def f1(v):
    return "–" if v is None else f"{v:,.1f}"


def f0(v):
    return "–" if v is None else f"{v:,.0f}"


bands = [("24–29 °C", "24", "29"), ("30–34 °C", "30", "34"), ("35–39 °C", "35", "39"), ("40–43 °C", "40", "43"), ("44 °C and above", "44+", "44+")]
rows = [["Hours a year at 24 °C+ (= 1,600)", ann["Total"][4], ann["Total"][0], ann["Total"][2]]]
for lab, lo, hi in bands:
    rows.append([lab, f0(band("T", lo, hi)), f1(band("A", lo, hi)), f1(band("B", lo, hi))])
rows.append(["35 °C and above", s6tab["35 °C or more, of 1,600"][2], s6tab["35 °C or more, of 1,600"][0], s6tab["35 °C or more, of 1,600"][1]])
build_table(s7, 8.30, 2.42, [1.60, 0.84, 0.86, 0.88], ["", "Today", "A", "B"], rows,
            row_h=0.36, highlight=(len(rows) - 1,), name="Table: of 1,600")
place(s7, T["bar_box"], name="Navy bar")
place(s7, T["bar_text"], name="Navy bar text")    # the deck's own 'Why below today at 35–42 °C?' line
place(s7, rewrite(copy.deepcopy(T["foot"]), [
    [("Today: IS 18154 (Part 1):2023 Table 4A. Proposed: CLASP analysis of IMD Daily Surface Table-III and hourly data, 2016–2025, 129 stations.", "foot")],
    [("A and B as on slide 6. Bands add the 1 °C bins of Annex 1 (rounded values), so they can differ by ±0.2 from unrounded totals.", "foot")]]),
    name="Footnote")
s7.notes_slide.notes_text_frame.text = (
    "Left: what is different between today's table and the proposal. Right: the same hours grouped into 5-degree bands "
    "on BEE's 1,600-hour basis. Sources: slides 4, 6, 8, 14 and Annex 1 (slide 17).")

# ------------------------------------------------------------------ new slide after slide 8: 40 °C and above
s9 = new_slide_after(S8, "Hours at 40 °C and above, degree by degree")
place(s9, rewrite(copy.deepcopy(T["subtitle"]), [[("Hours fall fast above 40 °C; 50 of 129 stations reach 44 °C or more in an average year", "subtitle")]]),
      name="Subtitle")
place(s9, rewrite(copy.deepcopy(T["label"]), [[("Hours counted at each degree, 2016–2025, by zone", "label")]]), x=0.85, y=2.08, w=6.60, name="Label left")
place(s9, rewrite(copy.deepcopy(T["sublabel"]), [[("10-year total over all usable days ", "sub_bold"), ("(not a yearly average)", "sub_text")]]),
      x=0.85, y=2.38, w=6.60, name="Sub-label left")
HOT = ["40", "41", "42", "43", "44+"]
zcols = ["Composite", "Hot-Dry", "Warm-Humid", "Undecided", "Cold"]
zi = {z: zhdr.index(z) - 1 for z in zcols + ["All 129 stations"]}
zrows = [["44 and above" if b == "44+" else b] + [zt[b][zi[z]] for z in zcols + ["All 129 stations"]] for b in HOT]
tot = [sum(int(zt[b][zi[z]].replace(",", "")) for b in HOT) for z in zcols + ["All 129 stations"]]
zrows.append(["40 and above"] + [f"{v:,}" for v in tot])
build_table(s9, 0.85, 2.66, [1.20, 0.95, 0.85, 0.98, 0.92, 0.55, 1.15],
            ["Bin °C", "Composite", "Hot-Dry", "Warm-Humid", "Undecided", "Cold", "All 129"], zrows,
            row_h=0.36, sz=1000, hdr_sz=900, highlight=(len(zrows) - 1,), name="Table: zone counts")
place(s9, rewrite(copy.deepcopy(T["small"]), [[
    ("Stations: ", "sub_bold"),
    ("Composite 42, Hot-Dry 16, Warm-Humid 59, Undecided 7, Cold 3. Temperate (2) has no hours at 40 °C or more.", "sub_text")]]),
    x=0.85, y=5.30, w=6.60, h=0.50, name="Note left")

place(s9, rewrite(copy.deepcopy(T["caps"]), [[("AVERAGE HOURS IN ONE YEAR, INDIA", "caps")]]), x=7.75, y=2.08, w=4.73, name="Label right")
nrows = []
for b in HOT:
    t = hours["T"][b]
    nrows.append(["44 and above" if b == "44+" else b, "not counted" if t.strip() in ("–", "-") else t, hours["A"][b], hours["B"][b]])
sumh = lambda k: sum(n(hours[k][b]) for b in HOT if hours[k][b].strip() not in ("–", "-"))
nrows.append(["40 and above", f0(sumh("T")), f0(sumh("A")), f0(sumh("B"))])
build_table(s9, 7.75, 2.40, [1.30, 1.15, 1.13, 1.15], ["Bin °C", "Today's table", "A: by population", "B: each station equal"],
            nrows, row_h=0.34, hdr_h=0.42, sz=1000, hdr_sz=900, highlight=(len(nrows) - 1,), name="Table: India hours")
place(s9, T["panel_box"], x=7.75, y=4.96, w=4.73, h=0.90, name="Panel")
top3 = ", ".join(f"{s} {v:.0f}" for s, v in hot8[:3])
place(s9, rewrite(copy.deepcopy(T["panel_text"]), [
    [("Stations at 44 °C and above", "panel_head")],
    [("50 of 129 stations (39%) reach 44 °C or more for at least an hour in an average year; 79 never do. Most hours a year: "
      f"{top3}.", "panel_body")]], ppr_from="panel_text"),
    x=7.92, y=5.02, w=4.40, h=0.80, name="Panel text")
place(s9, T["bar_box"], name="Navy bar")
c44 = {z: int(zt["44+"][zi[z]].replace(",", "")) for z in zcols + ["All 129 stations"]}
share_cd = (c44["Composite"] + c44["Hot-Dry"]) / c44["All 129 stations"]
place(s9, rewrite(copy.deepcopy(T["bar_text"]), [[
    ("Viewpoint: ", "bar_lead"),
    (f"of the {int(zt['Total'][zi['All 129 stations']].replace(',', '')) / 1e6:.1f} million hours counted, "
     f"{tot[-1]:,} are at 40 °C or more; {share_cd:.0%} of the {c44['All 129 stations']:,} hours at 44 °C or more "
     "are in Composite and Hot-Dry.", "bar_text")]]), name="Navy bar text")
place(s9, rewrite(copy.deepcopy(T["foot"]), [
    [("Left: Annex 2b (10-year totals; zones have different numbers of stations and days). Right: Annex 1, average year "
      "(rounded values; today's table has no 44 °C+ bin, so its '40 and above' is 40–43 °C). Stations: slide 9.", "foot")]]),
    name="Footnote")
s9.notes_slide.notes_text_frame.text = (
    "Left: Annex 2b, bins 40 to 44+ (10-year totals over all usable days; zones differ in stations and days, so compare "
    "shapes, not sizes). Right: Annex 1, average year for India. 50 of 129: slide 9. "
    "Still to add from station-level data: the highest temperature reached at each station, how many stations reach "
    "45, 46 °C and above, and hours at each degree above 44 °C.")

# ------------------------------------------------------------------ fix stale cross-references in the annexes
for slide in prs.slides:
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if "(slide 4)" in r.text:
                    r.text = r.text.replace("(slide 4)", "(slide 5)")
                if "slide 4 of the main deck" in r.text:
                    r.text = r.text.replace("slide 4 of the main deck", "slide 5 of the main deck")

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
