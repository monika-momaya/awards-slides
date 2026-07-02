#!/usr/bin/env python3
from __future__ import annotations

import copy
import re
import sys
from collections import defaultdict

import openpyxl
from pptx import Presentation
from pptx.util import Pt
from pptx.oxml.ns import qn

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
    "<<zone>>": "ZONE",
    "<<award category>>": "AWARD CATEGORY",
    "<<nominees>>": "NOMINEES",
    "<<winner>>": "WINNER",
    "<<nominee name>>": "NOMINEES",
    "<<winner name>>": "WINNER",
    "<<nominees-word>>": "NOMINEES_WORD",
    "<<winner-word>>": "WINNER_WORD",
    "<<placeholder x>>": "PLACEHOLDER_X",
    "<<placeholder y>>": "PLACEHOLDER_Y",
    "<<placeholder z>>": "PLACEHOLDER_Z",
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


def find_slide_index_with_token(prs, token):
    token = token.lower()
    for i, s in enumerate(prs.slides):
        txt = " ".join([shape.text for shape in s.shapes if hasattr(shape, "text") and shape.text]).lower()
        if token in txt:
            return i
    return None


def pick_template_indices(prs):
    nominee_idx = find_slide_index_with_token(prs, "<<nominee name>>")
    if nominee_idx is None:
        nominee_idx = find_slide_index_with_token(prs, "<<nominees>>")
    winner_idx = find_slide_index_with_token(prs, "<<winner name>>")
    if winner_idx is None:
        winner_idx = find_slide_index_with_token(prs, "<<winner>>")
    if nominee_idx is None and len(prs.slides) > 0:
        nominee_idx = 0
    if winner_idx is None and len(prs.slides) > 1:
        winner_idx = 1
    if winner_idx is None:
        winner_idx = nominee_idx
    if nominee_idx is None or winner_idx is None:
        raise RuntimeError("Template must contain at least one slide with nominee or winner placeholders.")
    return nominee_idx, winner_idx


def duplicate_slide(prs, index):
    """Duplicate slide at index inside prs, preserving all shapes, images, and background exactly."""
    source = prs.slides[index]
    blank_slide_layout = source.slide_layout
    dest = prs.slides.add_slide(blank_slide_layout)

    # Remove any placeholder shapes auto-added by the layout
    for shape in list(dest.shapes):
        shape._element.getparent().remove(shape._element)

    # Build a mapping of old rId -> new rId for all non-external relationships
    # (images, media, etc.) referenced by the source slide.
    rid_map = {}
    for rel_id, rel in source.part.rels.items():
        if rel.is_external:
            continue
        new_rid = dest.part.relate_to(rel.target_part, rel.reltype)
        rid_map[rel_id] = new_rid

    def remap_rids(element):
        for attr_name in ("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed",
                          "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link",
                          "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"):
            old_rid = element.get(attr_name)
            if old_rid and old_rid in rid_map:
                element.set(attr_name, rid_map[old_rid])
        for child in element:
            remap_rids(child)

    # Copy every shape element from source slide (preserves text, images, groups, formatting)
    for shape in source.shapes:
        new_el = copy.deepcopy(shape._element)
        remap_rids(new_el)
        dest.shapes._spTree.append(new_el)

    # Copy background (bg element) if present at slide level, remapping any image rIds too
    src_bg = source._element.find(qn('p:bg'))
    if src_bg is not None:
        new_bg = copy.deepcopy(src_bg)
        remap_rids(new_bg)
        dest_bg = dest._element.find(qn('p:bg'))
        if dest_bg is not None:
            dest._element.remove(dest_bg)
        dest._element.insert(0, new_bg)

    return dest


def move_slide(prs, old_index, new_index):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[old_index])
    xml_slides.insert(new_index, slides[old_index])


def delete_slide(prs, index):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    rId = slides[index].get(qn('r:id'))
    prs.part.drop_rel(rId)
    xml_slides.remove(slides[index])


def set_text(shape, text):
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    if run.font.size is None:
        run.font.size = Pt(22)


def fill_shapes_recursive(shapes, mapping):
    for shape in shapes:
        if shape.shape_type == 6 and hasattr(shape, "shapes"):  # group shape
            fill_shapes_recursive(shape.shapes, mapping)
            continue
        if not hasattr(shape, "text_frame") or not shape.has_text_frame:
            continue
        text = shape.text_frame.text
        if not text:
            continue
        low = text.lower()
        matched = False
        for token, key in TOKENS.items():
            if token in low:
                set_text(shape, mapping.get(key, ""))
                matched = True
                break


def fill_slide(slide, mapping):
    fill_shapes_recursive(slide.shapes, mapping)


def build_deck(excel_path, template_path):
    rows = read_rows(excel_path)
    nominee_groups, winner_groups = group_rows(rows)

    prs = Presentation(template_path)
    nominee_idx, winner_idx = pick_template_indices(prs)
    total_original = len(prs.slides)

    keys = sorted(set(nominee_groups) | set(winner_groups), key=lambda x: (x[0], x[1]))

    generated_indices = []
    for key in keys:
        zone, category = key
        if key in nominee_groups:
            entries = nominee_groups[key]
            new_slide = duplicate_slide(prs, nominee_idx)
            fill_slide(new_slide, {
                "ZONE": zone,
                "AWARD CATEGORY": category,
                "NOMINEES": "\n".join(e["nominee_name"] for e in entries),
                "NOMINEES_WORD": "NOMINEES",
                "PLACEHOLDER_X": entries[0]["placeholder_x"],
                "PLACEHOLDER_Y": entries[0]["placeholder_y"],
                "PLACEHOLDER_Z": entries[0]["placeholder_z"],
            })
            generated_indices.append(len(prs.slides.__iter__.__self__._sldIdLst) - 1)
        if key in winner_groups:
            entries = winner_groups[key]
            new_slide = duplicate_slide(prs, winner_idx)
            fill_slide(new_slide, {
                "ZONE": zone,
                "AWARD CATEGORY": category,
                "WINNER": "\n".join(e["winner_name"] for e in entries),
                "WINNER_WORD": "WINNER",
                "PLACEHOLDER_X": entries[0]["placeholder_x"],
                "PLACEHOLDER_Y": entries[0]["placeholder_y"],
                "PLACEHOLDER_Z": entries[0]["placeholder_z"],
            })
            generated_indices.append(len(prs.slides.__iter__.__self__._sldIdLst) - 1)

    # Remove original template slides (highest index first to keep indices valid)
    for idx in sorted({nominee_idx, winner_idx}, reverse=True):
        delete_slide(prs, idx)

    return prs


def main():
    if len(sys.argv) != 4:
        print("Usage: python generate_award_slides.py input.xlsx template.pptx output.pptx")
        raise SystemExit(1)
    deck = build_deck(sys.argv[1], sys.argv[2])
    deck.save(sys.argv[3])
    print(f"Saved {sys.argv[3]} ({len(deck.slides)} slides)")


if __name__ == "__main__":
    main()
