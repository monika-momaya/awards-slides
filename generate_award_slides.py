#!/usr/bin/env python3
from __future__ import annotations

import copy
import re
import sys
from collections import defaultdict

import openpyxl
from pptx import Presentation
from pptx.util import Pt

HEADER_MAP = {
    "award category": "award_category",
    "nominee name": "nominee_name",
    "winner name": "winner_name",
    "zone": "zone",
    "placeholder x": "placeholder_x",
    "placeholder y": "placeholder_y",
    "placeholder z": "placeholder_z",
}

TOKENS = {
    "<<ZONE>>": "ZONE",
    "<<AWARD CATEGORY>>": "AWARD CATEGORY",
    "<<NOMINEES>>": "NOMINEES",
    "<<WINNER>>": "WINNER",
    "<<nominees-word>>": "NOMINEES_WORD",
    "<<winner-word>>": "WINNER_WORD",
    "<<PLACEHOLDER X>>": "PLACEHOLDER_X",
    "<<PLACEHOLDER Y>>": "PLACEHOLDER_Y",
    "<<PLACEHOLDER Z>>": "PLACEHOLDER_Z",
}


def norm(v):
    return re.sub(r"\s+", " ", "" if v is None else str(v)).strip().lower()


def clean(v):
    return "" if v is None else re.sub(r"\s+", " ", str(v)).strip()


def read_rows(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    rows = []
    for ws in wb.worksheets:
        data = list(ws.iter_rows(values_only=True))
        if not data:
            continue
        headers = [clean(x) for x in data[0]]
        idx = {}
        for i, h in enumerate(headers):
            key = HEADER_MAP.get(norm(h))
            if key:
                idx[key] = i
        for r in data[1:]:
            if not any(v is not None and str(v).strip() for v in r):
                continue
            def g(key):
                i = idx.get(key)
                return clean(r[i]) if i is not None and i < len(r) else ""
            rows.append({
                "award_category": g("award_category"),
                "nominee_name": g("nominee_name"),
                "winner_name": g("winner_name"),
                "zone": g("zone") or ws.title,
                "placeholder_x": g("placeholder_x"),
                "placeholder_y": g("placeholder_y"),
                "placeholder_z": g("placeholder_z"),
            })
    return rows


def group_rows(rows):
    nominee_groups = defaultdict(list)
    winner_groups = defaultdict(list)
    for r in rows:
        key = (r["zone"], r["award_category"])
        if r["nominee_name"]:
            nominee_groups[key].append(r)
        if r["winner_name"]:
            winner_groups[key].append(r)
    return nominee_groups, winner_groups


def detect_role(slide):
    text = " ".join([shape.text for shape in slide.shapes if hasattr(shape, "text") and shape.text]).lower()
    if "nominees" in text:
        return "nominee"
    if "winner" in text:
        return "winner"
    return "generic"


def set_text(shape, text):
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(22)


def fill_slide(slide, mapping):
    for shape in slide.shapes:
        if not hasattr(shape, "text"):
            continue
        low = (shape.text or "").lower()
        for token, key in TOKENS.items():
            if token.lower() in low:
                set_text(shape, mapping.get(key, ""))
            elif key == "NOMINEES_WORD" and token.lower() == "<<nominees-word>>":
                set_text(shape, "NOMINEES")
            elif key == "WINNER_WORD" and token.lower() == "<<winner-word>>":
                set_text(shape, "WINNER")


def clone_template_slide(template_slide, out_prs):
    slide = out_prs.slides.add_slide(out_prs.slide_layouts[6])
    for shape in list(slide.shapes):
        el = shape.element
        el.getparent().remove(el)
    for shape in template_slide.shapes:
        slide.shapes._spTree.insert_element_before(copy.deepcopy(shape.element), 'p:extLst')
    return slide


def pick_template_slides(prs):
    nominee = None
    winner = None
    generic = prs.slides[0] if len(prs.slides) > 0 else None
    for s in prs.slides:
        txt = " ".join([shape.text for shape in s.shapes if hasattr(shape, "text") and shape.text]).lower()
        if nominee is None and "nominees" in txt:
            nominee = s
        if winner is None and "winner" in txt:
            winner = s
    if nominee is None:
        nominee = generic
    if winner is None:
        winner = generic if generic is not None else nominee
    return nominee, winner


def build_deck(excel_path, template_path):
    rows = read_rows(excel_path)
    nominee_groups, winner_groups = group_rows(rows)
    prs = Presentation(template_path)
    nominee_tpl, winner_tpl = pick_template_slides(prs)
    out = Presentation()
    out.slide_width = prs.slide_width
    out.slide_height = prs.slide_height

    keys = sorted(set(nominee_groups) | set(winner_groups), key=lambda x: (x[1], x[0]))
    for key in keys:
        zone, category = key
        if key in nominee_groups:
            entries = nominee_groups[key]
            slide = clone_template_slide(nominee_tpl, out)
            fill_slide(slide, {
                "ZONE": zone,
                "AWARD CATEGORY": category,
                "NOMINEES": "\n".join(e["nominee_name"] for e in entries),
                "NOMINEES_WORD": "NOMINEES",
                "PLACEHOLDER_X": entries[0]["placeholder_x"],
                "PLACEHOLDER_Y": entries[0]["placeholder_y"],
                "PLACEHOLDER_Z": entries[0]["placeholder_z"],
            })
        if key in winner_groups:
            entries = winner_groups[key]
            slide = clone_template_slide(winner_tpl, out)
            fill_slide(slide, {
                "ZONE": zone,
                "AWARD CATEGORY": category,
                "WINNER": "\n".join(e["winner_name"] for e in entries),
                "WINNER_WORD": "WINNER",
                "PLACEHOLDER_X": entries[0]["placeholder_x"],
                "PLACEHOLDER_Y": entries[0]["placeholder_y"],
                "PLACEHOLDER_Z": entries[0]["placeholder_z"],
            })
    return out


def main():
    if len(sys.argv) != 4:
        print("Usage: python generate_award_slides.py input.xlsx template.pptx output.pptx")
        raise SystemExit(1)
    deck = build_deck(sys.argv[1], sys.argv[2])
    deck.save(sys.argv[3])
    print(f"Saved {sys.argv[3]} ({len(deck.slides)} slides)")


if __name__ == "__main__":
    main()
