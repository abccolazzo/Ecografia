#!/usr/bin/env python3
"""Applica la grafica ispirata a borlizzi.it (arancione/azzurro, titoli condensati, logo in intestazione)
a un digest Word, senza cambiarne il testo.
Uso: python3 stile_borlizzi.py ingresso.docx uscita.docx [--logo assets/logo-borlizzi.png]"""
import sys, os, re, argparse
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

ORANGE = "F16523"; ORANGE_D = "C94A0E"; BLUE = "7FA9B8"; BLUE_D = "3E6F80"; INK = "22303A"; MUTED = "5D6D77"
DISPLAY = "Arial Narrow"; BODY = "Arial"
SECTION_PREFIXES = ("teranostica", "traccianti", "nuove molecole", "altro", "regolator", "approvazioni", "trial clinici")

def set_font(run, name, size, bold=None, italic=None, color=None, caps=None):
    run.font.name = name
    rpr = run._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), name)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rf.get(qn(a)) is not None: del rf.attrib[qn(a)]
    run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if italic is not None: run.font.italic = italic
    if color: run.font.color.rgb = RGBColor.from_string(color)
    if caps is not None: run.font.all_caps = caps

def border(par, side, color, sz=12, space=4):
    ppr = par._p.get_or_add_pPr()
    pb = ppr.find(qn("w:pBdr"))
    if pb is None:
        pb = OxmlElement("w:pBdr"); ppr.append(pb)
    el = OxmlElement("w:" + side)
    for k, v in (("val", "single"), ("sz", str(sz)), ("space", str(space)), ("color", color)):
        el.set(qn("w:" + k), v)
    pb.append(el)

def all_runs(par):
    return [docx.text.run.Run(r, par) for r in par._p.iter(qn("w:r"))]

def fmt(par, kind):
    pf = par.paragraph_format
    runs = all_runs(par)
    if kind == "title":
        pf.space_before = Pt(6); pf.space_after = Pt(4); pf.keep_with_next = True
        for r in runs: set_font(r, DISPLAY, 26, True, False, BLUE_D, True)
    elif kind == "intro":
        pf.space_after = Pt(2)
        for r in runs: set_font(r, BODY, 9.5, False, True, MUTED, False)
    elif kind == "heading":
        pf.space_before = Pt(18); pf.space_after = Pt(6); pf.keep_with_next = True
        for r in runs: set_font(r, DISPLAY, 15, True, False, BLUE_D, True)
        border(par, "bottom", ORANGE, 12, 3)
    elif kind == "item":
        pf.space_before = Pt(10); pf.space_after = Pt(1); pf.keep_with_next = True
        for r in runs: set_font(r, DISPLAY, 12, True, False, INK, False)
    elif kind == "meta":
        pf.space_after = Pt(2); pf.keep_with_next = True
        for r in runs: set_font(r, BODY, 9, False, True, MUTED, False)
    elif kind == "link":
        pf.space_after = Pt(4)
        for r in runs: set_font(r, BODY, 9, False, False, ORANGE_D, False)
    else:
        pf.space_after = Pt(3); pf.line_spacing = 1.1
        for r in runs: set_font(r, BODY, 10.5, None, None, INK, False)

def classify(paras):
    kinds = {}; first = None; seen_heading = False; prev = None
    for i, p in enumerate(paras):
        t = p.text.strip(); runs = [r for r in all_runs(p) if r.text.strip()]
        sname = (p.style.name if p.style is not None else "") or ""
        if not t:
            kinds[i] = "empty"; continue
        if first is None:
            first = i; kinds[i] = "title"; prev = "title"; continue
        low = t.lower()
        if sname.lower().startswith("heading") or (len(t) < 80 and low.startswith(SECTION_PREFIXES)):
            kinds[i] = "heading"; seen_heading = True
        elif re.match(r"^https?://\S+$", t):
            kinds[i] = "link"
        elif not seen_heading:
            kinds[i] = "intro"
        elif runs and all(r.bold for r in runs) and prev in ("heading", "link", "body", "item", "empty", None) and len(t) > 12:
            kinds[i] = "item"
        elif runs and all(r.italic for r in runs) and prev == "item":
            kinds[i] = "meta"
        else:
            kinds[i] = "body"
        prev = kinds[i]
    return kinds

def add_field(par, instr):
    r = par.add_run()
    for typ, txt in (("begin", None), (None, instr), ("end", None)):
        if typ:
            fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), typ); r._r.append(fc)
        else:
            it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = txt; r._r.append(it)
    return r

def decorate(d, logo):
    for s in d.sections:
        s.top_margin = Cm(2.4); s.bottom_margin = Cm(2.0)
        s.header_distance = Cm(0.9); s.footer_distance = Cm(0.9)
        h = s.header; h.is_linked_to_previous = False
        for p in list(h.paragraphs)[1:]:
            p._p.getparent().remove(p._p)
        hp = h.paragraphs[0]
        for ch in list(hp._p):
            if ch.tag != qn("w:pPr"): hp._p.remove(ch)
        if logo and os.path.exists(logo):
            hp.add_run().add_picture(logo, height=Cm(1.3))
        border(hp, "bottom", ORANGE, 24, 6)
        f = s.footer; f.is_linked_to_previous = False
        fp = f.paragraphs[0]
        for ch in list(fp._p):
            if ch.tag != qn("w:pPr"): fp._p.remove(ch)
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        border(fp, "top", BLUE, 6, 4)
        r = fp.add_run("Borlizzi Medical Center · Aggiornamento Ecografia · uso personale · pag. ")
        set_font(r, BODY, 8, False, False, MUTED)
        r2 = add_field(fp, "PAGE"); set_font(r2, BODY, 8, False, False, MUTED)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src"); ap.add_argument("dst"); ap.add_argument("--logo", default="assets/logo-borlizzi.png")
    a = ap.parse_args()
    d = docx.Document(a.src)
    paras = d.paragraphs
    kinds = classify(paras)
    for i, p in enumerate(paras):
        if kinds[i] != "empty": fmt(p, kinds[i])
        else: p.paragraph_format.space_after = Pt(0)
    decorate(d, a.logo)
    d.save(a.dst)
    from collections import Counter
    print(os.path.basename(a.src), dict(Counter(kinds.values())))

if __name__ == "__main__":
    main()
