"""Extract every table, chart cache and the slide-6 notes from the v2 deck into deck_data.json.
Usage: python -I extract_deck_data.py <deck.pptx> <out.json>"""
import json, sys
from pptx import Presentation

deck, out = sys.argv[1], sys.argv[2]
prs = Presentation(deck)
data = {"tables": {}, "charts": {}, "notes": {}}
for i, slide in enumerate(prs.slides, 1):
    t = c = 0
    for sh in slide.shapes:
        if sh.has_table:
            t += 1
            data["tables"][f"s{i}_t{t}"] = [[cell.text for cell in row.cells] for row in sh.table.rows]
        if getattr(sh, "has_chart", False) and sh.has_chart:
            c += 1
            ch = sh.chart
            plot = ch.plots[0]
            data["charts"][f"s{i}_c{c}"] = {
                "categories": [str(x) for x in plot.categories],
                "series": {s.name: list(s.values) for p in ch.plots for s in p.series},
            }
    if slide.has_notes_slide:
        data["notes"][f"s{i}"] = slide.notes_slide.notes_text_frame.text
json.dump(data, open(out, "w"), indent=1, ensure_ascii=False)
print({k: list(v) for k, v in data.items()})
