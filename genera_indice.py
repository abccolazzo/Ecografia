#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera "Indice_Ecografia.html" a partire dai digest Word (.docx) di notizie
da SIUMB, EFSUMB e WFUMB
presenti nella stessa cartella di questo script.

Uso:  python3 genera_indice.py            (dalla cartella Ecografia)
      python3 genera_indice.py /percorso/della/cartella

Non richiede librerie esterne: usa solo la libreria standard di Python 3.
La pagina generata si apre con doppio clic, senza server; logo e font
sono nella sottocartella "assets" accanto alla pagina.
"""
import datetime
import hashlib
import html
import json
import os
import re
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

OUTPUT_NAME = "Indice_Ecografia.html"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"

MESI = {
    "gennaio": 1, "febbraio": 2, "marzo": 3, "aprile": 4, "maggio": 5, "giugno": 6,
    "luglio": 7, "agosto": 8, "settembre": 9, "ottobre": 10, "novembre": 11, "dicembre": 12,
}
MESI_NOMI = ["", "gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio",
             "agosto", "settembre", "ottobre", "novembre", "dicembre"]

# ---------------------------------------------------------------------------
# Raggruppamenti tematici: (gruppo, etichetta, espressione regolare).
# Si applicano a titolo + riferimento + riassunto di ogni articolo.
# Per aggiungere un tema basta aggiungere una riga.
# ---------------------------------------------------------------------------
TEMI = [
    # Società / fonte
    ("Società", "SIUMB", r"SIUMB|siumb\.it"),
    ("Società", "EFSUMB", r"EFSUMB|efsumb\.org|\bEUROSON\b|Ultraschall in der Medizin|\bEJUM\b|Ultrasound International Open"),
    ("Società", "WFUMB", r"WFUMB|wfumb\.org|Ultrasound in Medicine (?:&|and) Biology|\bUMB\b|WFUMB Ultrasound Open"),
    # Distretto / ambito clinico
    ("Ambito clinico", "Fegato e vie biliari", r"fegat|liver|hepat|epat|\bMASLD\b|\bNAFLD\b|cirros|cirrhos|\bHCC\b|biliar|colecist|gallbladder"),
    ("Ambito clinico", "Addome e gastroenterologia", r"addom|abdom|intestin|bowel|\bIBD\b|crohn|pancrea|milza|spleen|gastro"),
    ("Ambito clinico", "Tiroide e collo", r"tiroid|thyroid|paratiroid|parathyroid|\bcollo\b|\bneck\b|TI-?RADS"),
    ("Ambito clinico", "Mammella", r"mammell|\bbreast|senolog|BI-?RADS"),
    ("Ambito clinico", "Vascolare", r"\bvascol|\bvascular|carotid|venos|venous|\barteri|\bTVP\b|\bDVT\b|aneurism|aneurysm"),
    ("Ambito clinico", "Cuore", r"cardi|ecocardio|echocardio|\bheart\b"),
    ("Ambito clinico", "Polmone e torace", r"polmon|\blung|toraci|thorac|pleur"),
    ("Ambito clinico", "Rene e urologia", r"\brene\b|\breni\b|renal|kidney|urolog|prostat|vescic|bladder|scrot|testic"),
    ("Ambito clinico", "Muscoloscheletrico", r"muscolo|musculo|\bMSK\b|tendin|tendon|\barticola[rz]|\bjoint|reumat|rheumat"),
    ("Ambito clinico", "Ostetricia e ginecologia", r"ostetric|obstetric|ginecolog|gynaecolog|gynecolog|fetal|fetale|gravidanz|pregnan|prenatal"),
    ("Ambito clinico", "Pediatria", r"pediatr|paediatr|neonat|bambin|child"),
    ("Ambito clinico", "Emergenza / POCUS", r"POCUS|point-of-care|point of care|emergenz|emergency|urgenz|critical care|terapia intensiva|\bFAST\b"),
    ("Ambito clinico", "Medicina estetica / dermatologia", r"estetic|esthetic|aesthetic|dermatolog|\bskin\b|\bcute\b"),
    ("Ambito clinico", "Veterinaria", r"veterinar"),
    ("Ambito clinico", "Oncologia", r"oncolog|tumor|tumour|cancer|neoplas|lesioni focali|focal lesion"),
    # Tecnica
    ("Tecnica", "CEUS / mezzo di contrasto", r"CEUS|contrast|contrasto|microbubb|microbolle"),
    ("Tecnica", "Elastografia", r"elastograf|elastograph|shear.?wave|\bSWE\b|stiffness|rigidit"),
    ("Tecnica", "Doppler", r"doppler"),
    ("Tecnica", "Ecografia interventistica", r"interventist|interventional|biops|ablaz|ablation|drenagg|drainage|\bINVUS\b|ecoguidat|guided"),
    ("Tecnica", "Intelligenza artificiale", r"intelligenza artificiale|artificial intelligence|(?-i:\bAI\b|\bIA\b)|deep learning|machine learning"),
    ("Tecnica", "Teleecografia / robotica / handheld", r"tele-?ecograf|tele-?ultrasound|robot|handheld|portatil|pocket"),
    ("Tecnica", "Ecografia quantitativa / QUS", r"\bQUS\b|quantitativ|attenuation|attenuaz|steatos|fat fraction"),
    # Tipo di contenuto
    ("Tipo di contenuto", "Linea guida / raccomandazioni", r"linea guida|linee guida|guideline|raccomandaz|recommendation|position (?:paper|statement)|consensus|consenso|statement"),
    ("Tipo di contenuto", "Congresso", r"congress|EUROSON|simposio|symposium|meeting"),
    ("Tipo di contenuto", "Corso / formazione", r"\bcors[oi]\b|\bcourse|scuola|school|formazion|training|educat|teaching|\bECB\b|\bCOE\b|centre of excellence|centro di eccellenza|esame|exam|certificaz|certificat|diploma|\bECM\b"),
    ("Tipo di contenuto", "Webinar / online", r"webinar|online|e-?learning|podcast|video"),
    ("Tipo di contenuto", "Rivista / pubblicazione", r"rivista|journal|pubblicaz|publication|articol[oi]|article|\bissue\b|fascicolo|\bdoi\b"),
    ("Tipo di contenuto", "Borse / premi / bandi", r"\bbors[ae]\b|fellowship|grant|premi[oa]?\b|award|bando|call for|scholarship"),
    ("Tipo di contenuto", "Vita associativa", r"assemblea|general assembly|\belezion|\belection|consiglio direttivo|\bboard\b|membership|quota|soci\b|members|statuto|president"),
]
TEMI_RE = [(g, l, re.compile(rx, re.I)) for g, l, rx in TEMI]

SEZIONI_CANONICHE = [
    ("siumb", "SIUMB – Società Italiana"),
    ("efsumb", "EFSUMB – Federazione Europea"),
    ("wfumb", "WFUMB – Federazione Mondiale"),
    ("altro", "Altro"),
]


def canon_section(name):
    low = name.lower()
    for key, label in SEZIONI_CANONICHE:
        if key in low:
            return label
    return name


# ---------------------------------------------------------------------------
# Lettura .docx con la sola libreria standard
# ---------------------------------------------------------------------------
def _is_on(el):
    if el is None:
        return False
    val = el.get(W + "val")
    return val is None or val.lower() not in ("0", "false", "off")


def read_docx(path):
    with zipfile.ZipFile(path) as z:
        doc = ET.fromstring(z.read("word/document.xml"))
        rels = {}
        if "word/_rels/document.xml.rels" in z.namelist():
            for rel in ET.fromstring(z.read("word/_rels/document.xml.rels")).iter(PR + "Relationship"):
                rels[rel.get("Id")] = rel.get("Target")
        styles = {}
        if "word/styles.xml" in z.namelist():
            for st in ET.fromstring(z.read("word/styles.xml")).iter(W + "style"):
                name = st.find(W + "name")
                styles[st.get(W + "styleId")] = (name.get(W + "val") if name is not None else st.get(W + "styleId"))

    paragraphs = []
    body = doc.find(W + "body")
    for p in body.iter(W + "p"):
        ppr = p.find(W + "pPr")
        style = ""
        if ppr is not None and ppr.find(W + "pStyle") is not None:
            sid = ppr.find(W + "pStyle").get(W + "val")
            style = styles.get(sid, sid) or ""
        texts, bolds, links = [], [], []
        for el in p.iter():
            if el.tag == W + "hyperlink":
                rid = el.get(R + "id")
                if rid and rid in rels:
                    links.append(rels[rid])
            elif el.tag == W + "r":
                rpr = el.find(W + "rPr")
                run_text = ""
                for c in el:
                    if c.tag == W + "t":
                        run_text += c.text or ""
                    elif c.tag in (W + "tab", W + "br"):
                        run_text += " "
                if run_text.strip():
                    bolds.append(rpr is not None and _is_on(rpr.find(W + "b")))
                texts.append(run_text)
            elif el.tag == W + "instrText" and el.text and "HYPERLINK" in el.text:
                m = re.search(r'HYPERLINK\s+"([^"]+)"', el.text)
                if m:
                    links.append(m.group(1))
        text = re.sub(r"\s+", " ", "".join(texts)).strip()
        if not text:
            continue
        paragraphs.append({
            "text": text,
            "style": style.lower(),
            "bold": bool(bolds) and all(bolds),
            "links": links,
        })
    return paragraphs


def is_heading(p):
    s = p["style"]
    return s.startswith("heading") or s.startswith("titolo ") or re.match(r"^titolo\d", s or "") is not None


def is_title_style(p):
    return p["style"] in ("title", "titolo")


URL_RE = re.compile(r"https?://\S+|\bdoi\.org/\S+", re.I)


def is_link_para(p):
    t = p["text"]
    return bool(re.match(r"^(https?://|www\.|doi\.org|link\s*:|fonte\s*:\s*https?)", t, re.I)) or (
        p["links"] and len(t) < 160 and URL_RE.search(t) is not None and len(URL_RE.sub("", t).strip()) < 20)


def find_urls(paras):
    urls = []
    for p in paras:
        urls.extend(p["links"])
        urls.extend(u.rstrip(".,;)") for u in URL_RE.findall(p["text"]))
    out = []
    for u in urls:
        if u.lower().startswith("doi.org"):
            u = "https://" + u
        if u not in out:
            out.append(u)
    return out


# ---------------------------------------------------------------------------
# Date
# ---------------------------------------------------------------------------
def date_from_name(name):
    m = re.search(r"(20\d\d)[-_.](\d{1,2})[-_.](\d{1,2})", name)
    if m:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r"(\d{1,2})[-_.](\d{1,2})[-_.](20\d\d)", name)
    if m:
        return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.search(r"(\d{1,2})[\s_-]+([a-zA-Z]+)[\s_-]+(20\d\d)", name)
    if m and m.group(2).lower() in MESI:
        return datetime.date(int(m.group(3)), MESI[m.group(2).lower()], int(m.group(1)))
    return None


def date_from_text(text):
    m = re.search(r"settimana del\s+(\d{1,2})\s+([a-zA-Z]+)\s+(20\d\d)", text, re.I)
    if m and m.group(2).lower() in MESI:
        return datetime.date(int(m.group(3)), MESI[m.group(2).lower()], int(m.group(1)))
    m = re.search(r"settimana del\s+(\d{1,2})/(\d{1,2})/(20\d\d)", text, re.I)
    if m:
        return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    return None


def label_date(d):
    return "%d %s %d" % (d.day, MESI_NOMI[d.month], d.year)


# ---------------------------------------------------------------------------
# Parsing di un digest
# ---------------------------------------------------------------------------
def parse_digest(paras):
    header, notes, sections = [], [], []
    current_section = None
    current_article = None

    def close_article():
        nonlocal current_article
        if current_article is not None:
            current_section["articles"].append(current_article)
            current_article = None

    for p in paras:
        if is_heading(p):
            if current_section is not None:
                close_article()
            current_section = {"name": p["text"], "articles": []}
            sections.append(current_section)
            continue
        if current_section is None:
            header.append(p)
            continue
        if p["bold"] and not is_link_para(p):
            close_article()
            current_article = {"title": p["text"], "paras": [], "closed": False}
            continue
        if current_article is None or current_article["closed"]:
            close_article()
            notes.append(p["text"])
            continue
        current_article["paras"].append(p)
        if is_link_para(p):
            current_article["closed"] = True
    if current_section is not None:
        close_article()

    # Le note in coda all'ultima sezione (es. "Fonte: ...") restano note del file.
    out_sections = []
    for s in sections:
        arts = []
        for a in s["articles"]:
            body = [p for p in a["paras"] if not is_link_para(p)]
            meta = body[0]["text"] if body else ""
            summary = " ".join(p["text"] for p in body[1:])
            if len(body) == 1 and len(meta) > 220:
                meta, summary = "", meta
            meta_clean = re.sub(r"\s*\|\s*PubMed PMID:?\s*\d+\s*$", "", meta).strip()
            arts.append({
                "title": a["title"],
                "meta": meta_clean or meta,
                "summary": summary,
                "urls": find_urls(a["paras"]),
                "pmid": None,
            })
        out_sections.append({"name": s["name"], "articles": arts})
    return header, notes, out_sections


def article_key(a):
    txt = " ".join([a["meta"], a["summary"]] + a["urls"])
    for u in a["urls"]:
        m = re.search(r"pubmed\.ncbi\.nlm\.nih\.gov/(\d+)", u)
        if m:
            return "pmid:" + m.group(1)
    m = re.search(r"PMID:?\s*(\d{6,9})", txt)
    if m:
        return "pmid:" + m.group(1)
    for u in a["urls"]:
        m = re.search(r"(10\.\d{4,9}/[^\s?#]+)", u)
        if m:
            return "doi:" + m.group(1).lower().rstrip("/")
    t = unicodedata.normalize("NFKD", a["title"]).encode("ascii", "ignore").decode().lower()
    return "t:" + re.sub(r"[^a-z0-9]+", "", t)[:80]


def tag_article(a):
    hay = " ".join([a["title"], a["meta"], a["summary"]] + a["urls"])
    return [[g, l] for g, l, rx in TEMI_RE if rx.search(hay)]


# ---------------------------------------------------------------------------
def main():
    folder = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__)))
    names = sorted(os.listdir(folder), key=lambda n: (len(n), n))  # a parità di contenuto vince il nome più corto
    files, digests, seen_hash = [], [], {}

    for name in names:
        path = os.path.join(folder, name)
        if name.startswith(("~$", ".")) or not os.path.isfile(path):
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext not in (".docx", ".pdf", ".doc"):
            continue
        info = {"name": name, "ext": ext[1:], "size": os.path.getsize(path),
                "modified": datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d/%m/%Y"),
                "date": None, "status": "", "digest": None}
        d = date_from_name(name)
        if ext == ".docx":
            with open(path, "rb") as fh:
                h = hashlib.sha1(fh.read()).hexdigest()
            if h in seen_hash:
                info["status"] = "Copia identica di " + seen_hash[h]
            else:
                try:
                    paras = read_docx(path)
                    header, notes, sections = parse_digest(paras)
                except Exception as exc:  # file danneggiato o formato inatteso
                    info["status"] = "Non leggibile: %s" % exc
                    files.append(info)
                    continue
                seen_hash[h] = name
                head_text = " ".join(p["text"] for p in header)
                d = d or date_from_text(head_text) or datetime.date.fromtimestamp(os.path.getmtime(path))
                subtitle = [p["text"] for p in header if not re.match(r"^aggiornamento ecografia.*$", p["text"], re.I)]
                digests.append({"file": name, "date": d, "subtitle": subtitle,
                                "notes": notes, "sections": sections})
                info["digest"] = len(digests) - 1
        elif ext == ".pdf":
            info["status"] = "PDF (versione stampabile, non indicizzato)"
        else:
            info["status"] = "Formato .doc non supportato: salvarlo come .docx"
        info["date"] = d.isoformat() if d else None
        files.append(info)

    digests.sort(key=lambda x: x["date"], reverse=True)
    weeks, articles, by_key = [], [], {}
    for wi, dg in enumerate(digests):
        week = {"id": "w%d" % wi, "date": dg["date"].isoformat(), "label": label_date(dg["date"]),
                "file": dg["file"], "subtitle": dg["subtitle"], "notes": dg["notes"], "sections": []}
        for s in dg["sections"]:
            sec = {"name": s["name"], "canon": canon_section(s["name"]), "articles": []}
            for a in s["articles"]:
                aid = "a%d" % len(articles)
                key = article_key(a)
                art = {"id": aid, "week": week["id"], "section": s["name"], "canon": sec["canon"],
                       "title": a["title"], "meta": a["meta"], "summary": a["summary"],
                       "urls": a["urls"], "tags": tag_article(a), "key": key}
                articles.append(art)
                by_key.setdefault(key, []).append(aid)
                sec["articles"].append(aid)
            week["sections"].append(sec)
        weeks.append(week)
    for art in articles:
        art["also"] = [x for x in by_key[art["key"]] if x != art["id"]]
        del art["key"]

    for f in files:
        f["week"] = None
        for w in weeks:
            if w["file"] == f["name"]:
                f["week"] = w["id"]
    files.sort(key=lambda f: (f["date"] or "", f["name"]), reverse=True)

    groups = []
    for g, l, _ in TEMI:
        if g not in groups:
            groups.append(g)
    data = {
        "generated": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "files": files, "weeks": weeks, "articles": articles,
        "groups": groups, "sections": [l for _, l in SEZIONI_CANONICHE],
    }
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    out = TEMPLATE.replace("/*__DATA__*/null", payload)
    out_path = os.path.join(folder, OUTPUT_NAME)
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(out)
    unique = len(set(a["title"].lower() for a in articles))
    print("Creato %s: %d settimane, %d articoli (%d distinti), %d file elencati." % (
        out_path, len(weeks), len(articles), unique, len(files)))


TEMPLATE = r"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Indice Ecografia</title>
<style>
@font-face{font-family:"Oswald";font-weight:200 700;font-style:normal;font-display:swap;src:url("assets/Oswald-latin.woff2") format("woff2");unicode-range:U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD}
@font-face{font-family:"Oswald";font-weight:200 700;font-style:normal;font-display:swap;src:url("assets/Oswald-latin-ext.woff2") format("woff2");unicode-range:U+0100-02BA,U+02BD-02C5,U+02C7-02CC,U+02CE-02D7,U+02DD-02FF,U+1E00-1EFF,U+20A0-20AB,U+20AD-20C4,U+2C60-2C7F,U+A720-A7FF}
:root{
  --bg:#f4f6f7; --panel:#ffffff; --ink:#22303a; --muted:#5d6d77; --line:#dce3e6;
  --accent:#c94a0e; --accent-soft:#fdeee5; --mark:#fff0a8; --chip:#e8f0f3; --shadow:0 1px 2px rgba(0,0,0,.05);
  --brand:#f16523; --brand-d:#d9531a; --blue:#7fa9b8; --blue-d:#3e6f80; --display:"Oswald","Arial Narrow",Arial,sans-serif;
}
@media (prefers-color-scheme: dark){
  :root{ --bg:#141b1f; --panel:#1d272d; --ink:#e8eef0; --muted:#9fb0b8; --line:#2d3a41;
    --accent:#ff8247; --accent-soft:#3a2418; --mark:#6b5d12; --chip:#26343b; --shadow:none;
    --brand:#ff8247; --brand-d:#f16523; --blue:#8db8c6; --blue-d:#9cc5d2; }
}
*{box-sizing:border-box}
html,body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}
a{color:var(--accent)}
header{position:sticky;top:0;z-index:5;background:var(--panel);border-bottom:1px solid var(--line);padding:12px 20px;display:flex;flex-wrap:wrap;gap:12px;align-items:center}
header h1{font-size:18px;margin:0 12px 0 0;white-space:nowrap}
header h1 small{display:block;font-weight:400;font-size:12px;color:var(--muted)}
#q{flex:1 1 280px;min-width:0;padding:9px 12px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font-size:15px}
nav.tabs{display:flex;gap:4px;flex-wrap:wrap}
nav.tabs button{border:1px solid var(--line);background:var(--panel);color:var(--ink);padding:7px 12px;border-radius:8px;cursor:pointer;font-size:14px}
nav.tabs button.on{background:var(--accent);border-color:var(--accent);color:#fff}
.wrap{display:grid;grid-template-columns:320px 1fr;gap:0;min-height:calc(100vh - 62px)}
aside{border-right:1px solid var(--line);background:var(--panel);padding:14px 16px;position:sticky;top:62px;height:calc(100vh - 62px);overflow:auto}
main{padding:18px 24px 60px;min-width:0}
aside h3{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin:16px 0 6px}
aside h3:first-child{margin-top:0}
.toc details{margin:2px 0}
.toc summary{cursor:pointer;padding:3px 0;font-weight:600}
.toc .sec{font-size:12px;color:var(--muted);margin:6px 0 2px 14px;text-transform:uppercase;letter-spacing:.04em}
.toc a{display:block;margin:0 0 0 14px;padding:3px 6px;border-radius:5px;text-decoration:none;color:var(--ink);font-size:13px;line-height:1.35}
.toc a:hover{background:var(--accent-soft)}
.toc .count{color:var(--muted);font-weight:400;font-size:12px}
.chips{display:flex;flex-wrap:wrap;gap:5px}
.chip{border:1px solid transparent;background:var(--chip);color:var(--ink);padding:3px 9px;border-radius:999px;font-size:12.5px;cursor:pointer;line-height:1.5}
.chip.on{background:var(--accent);color:#fff}
.chip .n{opacity:.65;margin-left:3px}
.status{color:var(--muted);font-size:13px;margin:0 0 12px;display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.status button{border:1px solid var(--line);background:var(--panel);color:var(--ink);border-radius:6px;padding:3px 8px;cursor:pointer;font-size:12.5px}
h2.week{font-size:20px;margin:28px 0 4px;padding-top:6px}
h2.week:first-of-type{margin-top:0}
.weeksub{color:var(--muted);font-size:13px;margin:0 0 6px}
h3.sec{font-size:13px;text-transform:uppercase;letter-spacing:.05em;color:var(--accent);margin:18px 0 8px;border-bottom:1px solid var(--line);padding-bottom:4px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:12px 16px;margin:0 0 10px;box-shadow:var(--shadow);scroll-margin-top:80px}
.card.flash{outline:3px solid var(--accent)}
.card h4{margin:0 0 3px;font-size:15.5px;line-height:1.35}
.card .meta{color:var(--muted);font-size:13px;margin-bottom:6px}
.card .sum{margin:0 0 8px}
.card .foot{display:flex;flex-wrap:wrap;gap:6px;align-items:center;font-size:12.5px}
.badge{background:var(--chip);border-radius:5px;padding:1px 7px;color:var(--muted)}
.badge.wk{background:var(--accent-soft);color:var(--accent);cursor:pointer}
.tag{background:transparent;border:1px solid var(--line);border-radius:999px;padding:0 7px;color:var(--muted);cursor:pointer}
.tag:hover{border-color:var(--accent);color:var(--accent)}
.also{font-size:12.5px;color:var(--muted)}
.also a{cursor:pointer}
mark{background:var(--mark);color:inherit;border-radius:2px}
table.files{width:100%;border-collapse:collapse;background:var(--panel);border:1px solid var(--line);border-radius:10px;overflow:hidden}
table.files th,table.files td{text-align:left;padding:8px 12px;border-bottom:1px solid var(--line);font-size:14px;vertical-align:top}
table.files th{font-size:12px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted)}
.theme-group h2{font-size:18px;margin:26px 0 8px}
.theme-group:first-child h2{margin-top:0}
.theme-block{margin:0 0 14px}
.theme-block > summary{cursor:pointer;font-weight:600;font-size:15px;padding:6px 0}
.theme-block > summary .n{color:var(--muted);font-weight:400}
.empty{color:var(--muted);padding:30px 0}
.notes{font-size:12.5px;color:var(--muted);margin:4px 0 0}
#menuBtn{display:none}
@media (max-width: 860px){
  .wrap{grid-template-columns:1fr}
  aside{position:static;height:auto;border-right:none;border-bottom:1px solid var(--line);display:none}
  aside.open{display:block}
  #menuBtn{display:inline-block}
  main{padding:14px 16px 60px}
  header{padding:10px 16px;position:static}
}
@media print{ header,aside{display:none} .wrap{display:block} .card{break-inside:avoid} }

/* --- Grafica ispirata a borlizzi.it --- */
.brandbar{background:#fff;border-bottom:4px solid var(--brand);padding:14px 20px;display:flex;align-items:center;gap:20px;flex-wrap:wrap}
.brandbar img{height:64px;width:auto;max-width:100%;display:block}
.brandbar .slogan{font:400 15px var(--display);letter-spacing:.06em;text-transform:uppercase;color:var(--blue-d);margin:0}
@media (prefers-color-scheme: dark){ .brandbar{background:#fff} }
header{box-shadow:0 3px 0 var(--brand)}
header h1{font:500 19px/1.1 var(--display);letter-spacing:.04em;text-transform:uppercase;color:var(--blue-d)}
header h1 small{font:400 12px var(--body,inherit);letter-spacing:0;text-transform:none;color:var(--muted)}
nav.tabs button{font-family:var(--display);letter-spacing:.05em;text-transform:uppercase;font-size:13px;border-color:var(--blue)}
nav.tabs button.on{background:var(--brand-d);border-color:var(--brand-d);color:#fff}
.chip.on{background:var(--brand-d);color:#fff}
h2.week{font-family:var(--display);font-weight:500;letter-spacing:.03em;text-transform:uppercase;color:var(--blue-d);border-bottom:2px solid var(--brand);padding-bottom:6px}
h3.sec{font-family:var(--display);font-weight:500;color:var(--accent);border-bottom-color:var(--blue)}
.card{border-radius:6px}
.card h4{font-family:var(--display);font-weight:500;font-size:17px;letter-spacing:.01em}
aside h3{font-family:var(--display);font-weight:500;color:var(--blue-d)}
.theme-group h2{font-family:var(--display);font-weight:500;text-transform:uppercase;letter-spacing:.03em;color:var(--blue-d)}
.sitefoot{border-top:4px solid var(--brand);background:var(--panel);color:var(--muted);font-size:12.5px;padding:16px 20px}
@media (max-width:860px){ .brandbar{padding:10px 16px} .brandbar img{height:48px} }
</style>
</head>
<body>
<div class="brandbar"><img src="assets/logo-borlizzi.webp" alt="Borlizzi Medical Center"><p class="slogan">Aggiornamento ecografia · SIUMB, EFSUMB e WFUMB al servizio della pratica clinica</p></div>
<header>
  <h1>Aggiornamento Ecografia<small id="gen"></small></h1>
  <input id="q" type="search" placeholder="Cerca in titoli, autori, riviste e riassunti (es. CEUS fegato, EUROSON, linee guida)" autocomplete="off">
  <nav class="tabs" id="tabs">
    <button id="menuBtn" type="button">Indice e filtri</button>
    <button data-v="weeks" class="on" type="button">Per settimana</button>
    <button data-v="themes" type="button">Per tema</button>
    <button data-v="files" type="button">File</button>
  </nav>
</header>
<div class="wrap">
  <aside id="side">
    <h3>Indice dei contenuti</h3>
    <div class="toc" id="toc"></div>
    <h3>Sezione</h3>
    <div class="chips" id="secChips"></div>
    <div id="facetBox"></div>
  </aside>
  <main id="main"></main>
</div>
<div class="sitefoot">Fonti: siumb.it, efsumb.org, wfumb.org. Uso personale di studio e aggiornamento professionale.</div>
<script>
const DATA = /*__DATA__*/null;
const A = Object.fromEntries(DATA.articles.map(a => [a.id, a]));
const W = Object.fromEntries(DATA.weeks.map(w => [w.id, w]));
const state = { view: "weeks", q: "", tags: new Set(), secs: new Set(), themeGroup: DATA.groups[0] };

const norm = s => (s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
const esc = s => (s || "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const tagKey = t => t[0] + "|" + t[1];
DATA.articles.forEach(a => {
  a._hay = norm([a.title, a.meta, a.summary, a.section, W[a.week].label, a.tags.map(t => t[1]).join(" ")].join(" \n "));
  a._tags = new Set(a.tags.map(tagKey));
});
document.getElementById("gen").textContent =
  DATA.weeks.length + " settimane · " + DATA.articles.length + " voci · aggiornato il " + DATA.generated;

function terms() { return norm(state.q).split(/\s+/).filter(Boolean); }
function matches(a) {
  const t = terms();
  if (t.length && !t.every(x => a._hay.includes(x))) return false;
  for (const k of state.tags) if (!a._tags.has(k)) return false;
  if (state.secs.size && !state.secs.has(a.canon)) return false;
  return true;
}
function hl(text) {
  const t = terms();
  let out = esc(text);
  if (!t.length) return out;
  // evidenzia ignorando accenti: mappa posizioni sul testo normalizzato
  const src = text || ""; const n = norm(src);
  if (n.length !== src.length) return out;
  const marks = new Array(src.length).fill(false);
  t.forEach(x => { let i = 0; while ((i = n.indexOf(x, i)) !== -1) { for (let j = i; j < i + x.length; j++) marks[j] = true; i += x.length; } });
  let res = "", open = false;
  for (let i = 0; i < src.length; i++) {
    if (marks[i] && !open) { res += "<mark>"; open = true; }
    if (!marks[i] && open) { res += "</mark>"; open = false; }
    res += esc(src[i]);
  }
  return res + (open ? "</mark>" : "");
}
function linkLabel(u) {
  if (/pubmed\.ncbi/.test(u)) return "PubMed";
  if (/ncbi\.nlm\.nih\.gov\/pmc/.test(u)) return "PMC";
  if (/doi\.org/.test(u)) return "DOI";
  if (/siumb\.it/.test(u)) return "SIUMB";
  if (/efsumb\.org/.test(u)) return "EFSUMB";
  if (/wfumb\.org/.test(u)) return "WFUMB";
  try { return new URL(u).hostname.replace(/^www\./, ""); } catch (e) { return "Link"; }
}
function card(a, showWeek) {
  const w = W[a.week];
  const links = a.urls.map(u => `<a href="${esc(u)}" target="_blank" rel="noopener">Apri su ${esc(linkLabel(u))} ↗</a>`).join(" ");
  const tags = a.tags.map(t => `<span class="tag" data-tag="${esc(tagKey(t))}" title="${esc(t[0])}">${esc(t[1])}</span>`).join("");
  const also = a.also.length ? `<div class="also">Compare anche in: ${a.also.map(id => `<a data-go="${id}">${esc(W[A[id].week].label)}</a>`).join(", ")}</div>` : "";
  return `<article class="card" id="${a.id}">
    <h4>${hl(a.title)}</h4>
    ${a.meta ? `<div class="meta">${hl(a.meta)}</div>` : ""}
    ${a.summary ? `<p class="sum">${hl(a.summary)}</p>` : ""}
    <div class="foot">${showWeek ? `<span class="badge wk" data-week="${w.id}">${esc(w.label)}</span>` : ""}
      <span class="badge">${esc(a.canon)}</span>${tags} ${links}</div>${also}
  </article>`;
}
function statusBar(n) {
  const parts = [];
  if (state.q) parts.push(`ricerca «${esc(state.q)}»`);
  state.secs.forEach(s => parts.push(esc(s)));
  state.tags.forEach(k => parts.push(esc(k.split("|")[1])));
  const any = parts.length;
  return `<div class="status"><span>${n} voci${any ? " per " + parts.join(" + ") : ""}</span>${any ? '<button type="button" id="clear">Azzera filtri</button>' : ""}</div>`;
}

function renderWeeks() {
  let html = "", n = 0;
  DATA.weeks.forEach(w => {
    let body = "";
    w.sections.forEach(s => {
      const list = s.articles.map(id => A[id]).filter(matches);
      if (!list.length) return;
      n += list.length;
      body += `<h3 class="sec">${esc(s.name)}</h3>` + list.map(a => card(a, false)).join("");
    });
    if (body) html += `<h2 class="week" id="${w.id}">Settimana del ${esc(w.label)}</h2>
      <div class="weeksub">${esc(w.subtitle.join(" · "))} — file: ${esc(w.file)}</div>` + body;
  });
  return statusBar(n) + (html || '<div class="empty">Nessun risultato.</div>');
}
function renderThemes() {
  const list = DATA.articles.filter(matches);
  let html = statusBar(list.length) + '<div class="chips" style="margin-bottom:16px">' +
    DATA.groups.concat(["Sezione"]).map(g => `<button class="chip ${g === state.themeGroup ? "on" : ""}" data-group="${esc(g)}">${esc(g)}</button>`).join("") + "</div>";
  const buckets = new Map();
  list.forEach(a => {
    const labels = state.themeGroup === "Sezione" ? [a.canon] : a.tags.filter(t => t[0] === state.themeGroup).map(t => t[1]);
    (labels.length ? labels : ["(altri argomenti)"]).forEach(l => { if (!buckets.has(l)) buckets.set(l, []); buckets.get(l).push(a); });
  });
  const sorted = [...buckets.entries()].sort((x, y) => (x[0].startsWith("(") - y[0].startsWith("(")) || y[1].length - x[1].length);
  html += '<div class="theme-group">' + sorted.map(([l, arr]) => `<details class="theme-block" ${state.q || state.tags.size ? "open" : ""}>
     <summary>${esc(l)} <span class="n">(${arr.length})</span></summary>${arr.map(a => card(a, true)).join("")}</details>`).join("") + "</div>";
  return html;
}
function renderFiles() {
  const rows = DATA.files.map(f => {
    const w = f.week ? W[f.week] : null;
    const cnt = w ? w.sections.reduce((s, x) => s + x.articles.length, 0) : "";
    return `<tr><td><a href="${encodeURI(f.name)}">${esc(f.name)}</a></td>
      <td>${w ? `<a data-week="${w.id}" href="#${w.id}">${esc(w.label)}</a>` : (f.date ? esc(f.date.split("-").reverse().join("/")) : "")}</td>
      <td>${cnt}</td><td>${(f.size / 1024).toFixed(1)} KB</td>
      <td>${esc(f.status || (w ? w.sections.map(s => s.name).join("; ") : ""))}${w && w.notes.length ? `<div class="notes">${esc(w.notes.join(" "))}</div>` : ""}</td></tr>`;
  }).join("");
  return `<p class="status">${DATA.files.length} file nella cartella. Il nome del file apre il documento originale.</p>
    <table class="files"><thead><tr><th>File</th><th>Settimana</th><th>Voci</th><th>Dim.</th><th>Contenuto / note</th></tr></thead><tbody>${rows}</tbody></table>`;
}
function renderSide() {
  document.getElementById("toc").innerHTML = DATA.weeks.map((w, i) => {
    const secs = w.sections.map(s => {
      const list = s.articles.map(id => A[id]).filter(matches);
      if (!list.length) return "";
      return `<div class="sec">${esc(s.canon)}</div>` + list.map(a => `<a href="#${a.id}" data-go="${a.id}">${esc(a.title)}</a>`).join("");
    }).join("");
    const c = w.sections.reduce((s, x) => s + x.articles.filter(id => matches(A[id])).length, 0);
    return `<details ${i === 0 || (state.q && c) ? "open" : ""}><summary>${esc(w.label)} <span class="count">(${c})</span></summary>${secs}</details>`;
  }).join("");
  const counts = {}; DATA.articles.forEach(a => { counts[a.canon] = (counts[a.canon] || 0) + 1; a.tags.forEach(t => { const k = tagKey(t); counts[k] = (counts[k] || 0) + 1; }); });
  document.getElementById("secChips").innerHTML = DATA.sections.map(s =>
    `<button class="chip ${state.secs.has(s) ? "on" : ""}" data-sec="${esc(s)}">${esc(s)}<span class="n">${counts[s] || 0}</span></button>`).join("");
  const byGroup = {};
  DATA.articles.forEach(a => a.tags.forEach(t => { (byGroup[t[0]] = byGroup[t[0]] || new Set()).add(t[1]); }));
  document.getElementById("facetBox").innerHTML = DATA.groups.filter(g => byGroup[g]).map(g => {
    const labels = [...byGroup[g]].sort((x, y) => counts[g + "|" + y] - counts[g + "|" + x]);
    return `<h3>${esc(g)}</h3><div class="chips">` + labels.map(l => { const k = g + "|" + l;
      return `<button class="chip ${state.tags.has(k) ? "on" : ""}" data-tag="${esc(k)}">${esc(l)}<span class="n">${counts[k]}</span></button>`; }).join("") + "</div>";
  }).join("");
}
function render() {
  document.querySelectorAll("#tabs button[data-v]").forEach(b => b.classList.toggle("on", b.dataset.v === state.view));
  const m = document.getElementById("main");
  m.innerHTML = state.view === "weeks" ? renderWeeks() : state.view === "themes" ? renderThemes() : renderFiles();
  renderSide();
}
function go(id) {
  const a = A[id];
  if (!a) return;
  if (!matches(a)) { state.q = ""; state.tags.clear(); state.secs.clear(); document.getElementById("q").value = ""; }
  if (state.view === "files") state.view = "weeks";
  render();
  let el = document.getElementById(id);
  const det = el && el.closest("details"); if (det) det.open = true;
  if (el) { el.scrollIntoView({behavior: "smooth", block: "start"}); el.classList.add("flash"); setTimeout(() => el.classList.remove("flash"), 1600); }
  document.getElementById("side").classList.remove("open");
}
function goWeek(id) {
  state.view = "weeks"; render();
  const el = document.getElementById(id); if (el) el.scrollIntoView({behavior: "smooth"});
  document.getElementById("side").classList.remove("open");
}
let timer;
document.getElementById("q").addEventListener("input", e => { clearTimeout(timer); timer = setTimeout(() => { state.q = e.target.value.trim(); render(); }, 150); });
document.addEventListener("click", e => {
  const t = e.target.closest("[data-go],[data-week],[data-tag],[data-sec],[data-group],[data-v],#clear,#menuBtn");
  if (!t) return;
  if (t.id === "menuBtn") { document.getElementById("side").classList.toggle("open"); return; }
  if (t.id === "clear") { state.q = ""; state.tags.clear(); state.secs.clear(); document.getElementById("q").value = ""; render(); return; }
  if (t.dataset.v) { state.view = t.dataset.v; render(); window.scrollTo(0, 0); return; }
  if (t.dataset.go) { e.preventDefault(); go(t.dataset.go); return; }
  if (t.dataset.week) { e.preventDefault(); goWeek(t.dataset.week); return; }
  if (t.dataset.group) { state.themeGroup = t.dataset.group; render(); return; }
  if (t.dataset.tag) { const k = t.dataset.tag; state.tags.has(k) ? state.tags.delete(k) : state.tags.add(k); render(); return; }
  if (t.dataset.sec) { const k = t.dataset.sec; state.secs.has(k) ? state.secs.delete(k) : state.secs.add(k); render(); return; }
});
render();
if (location.hash && A[location.hash.slice(1)]) go(location.hash.slice(1));
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
