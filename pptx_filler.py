# -*- coding: utf-8 -*-
"""
Logica di compilazione del PowerPoint (indipendente da Streamlit, così è
testabile da sola).

Funzioni principali:
  - load_mapping(xlsx, regenerate=False) -> dict {segnaposto: testo}
  - extract_placeholders(pptx)           -> set di segnaposto presenti nel file
  - fill_pptx(pptx, mapping)             -> (BytesIO compilato, stats)

Gestisce in modo robusto:
  - segnaposto {{chiave}} spezzati su più "run" dentro lo stesso paragrafo
  - testo dentro tabelle
  - forme raggruppate (group shapes), in modo ricorsivo
"""
import io
import math
import re
from openpyxl import load_workbook
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# backend headless per generare il radar anche senza display
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as _plt
    _HAS_MPL = True
except Exception:                       # matplotlib assente -> radar saltato
    _HAS_MPL = False

PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")

# le 4 aree dell'assessment, nell'ordine usato dal radar
AREA_ORDER = ("strategia", "cultura", "processi", "mercato")


# ----------------------------------------------------------------- Excel -> mapping
def _records(ws):
    """Legge un foglio come lista di dict, header in prima riga (lower/strip)."""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [(str(h).strip().lower() if h is not None else "") for h in rows[0]]
    out = []
    for row in rows[1:]:
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        out.append({h: row[i] for i, h in enumerate(headers) if h})
    return out


def _wide_record(ws):
    """Legge un foglio 'wide': riga 1 = id (intestazioni colonna), riga 2 = livello.
    Ritorna {id: livello}."""
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return {}
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = rows[1]
    return {h: v for h, v in zip(headers, data) if h}


def _read_scala(wb):
    """{livello: (nome, descrizione)} dal foglio 'Scala', se presente."""
    scala = {}
    if "Scala" in wb.sheetnames:
        for rec in _records(wb["Scala"]):
            lv = rec.get("livello")
            if lv is None:
                continue
            nome = str(rec.get("nome", "")).strip()
            desc = "" if rec.get("descrizione") is None else str(rec["descrizione"]).strip()
            scala[int(lv)] = (nome, desc)
    return scala


def _derive_panoramica(mapping, scala):
    """Inietta i segnaposto panoramica.* : il livello macro aggregato
    (media dei livelli delle 4 aree). scala: {livello: (nome, descrizione)}.

    Le aree si riconoscono perché l'id NON contiene '_' (i sottogruppi sì),
    quindi la funzione è indipendente dalla fonte del mapping.
    """
    livelli = []
    for key, val in mapping.items():
        if not key.endswith(".risultato"):
            continue
        aid = key[: -len(".risultato")]
        if "_" in aid:                       # salta i sottogruppi
            continue
        m = re.match(r"\s*(\d+)", str(val))
        if m:
            livelli.append(int(m.group(1)))
    if not livelli or not scala:
        return
    avg = sum(livelli) / len(livelli)
    macro = max(1, min(4, int(avg + 0.5)))   # arrotondamento half-up, clamp 1..4
    nome = scala.get(macro, ("", ""))[0] or str(macro)
    desc = scala.get(macro, ("", ""))[1] or ""
    mapping["panoramica.risultato"] = f"{macro} - {nome}"
    mapping["panoramica.livello_medio"] = f"{avg:.2f}".replace(".", ",") + " / 4"
    mapping["panoramica.descrizione"] = desc


def _area_scores(mapping):
    """[(nome, livello)] delle 4 aree, ricavati dai <area>.risultato nel mapping."""
    scores = []
    for aid in AREA_ORDER:
        m = re.match(r"\s*(\d+)", str(mapping.get(f"{aid}.risultato", "")))
        scores.append((aid.upper(), int(m.group(1)) if m else 0))
    return scores


def _mapping_from_sheet(ws):
    mapping = {}
    for rec in _records(ws):
        key = rec.get("segnaposto")
        if key is None:
            continue
        key = str(key).strip()
        if key:
            mapping[key] = "" if rec.get("testo") is None else str(rec["testo"])
    return mapping


def _mapping_from_levels(wb):
    """Ricostruisce il mapping da Assessment + Libreria + Scala (dinamico)."""
    scala = _read_scala(wb)  # {livello: (nome, descrizione)}

    libreria = {}  # (id, livello) -> testo ; e tipo per id
    tipo_by_id = {}
    for rec in _records(wb["Libreria"]):
        _id = str(rec.get("id", "")).strip()
        lv = rec.get("livello")
        if not _id or lv is None:
            continue
        libreria[(_id, int(lv))] = "" if rec.get("testo") is None else str(rec["testo"])
        tipo_by_id[_id] = str(rec.get("tipo", "")).strip().lower()

    mapping = {}
    for rec in _records(wb["Assessment"]):
        _id = str(rec.get("id", "")).strip()
        lv = rec.get("livello")
        if not _id or lv is None:
            continue
        lv = int(lv)
        nome_liv = scala.get(lv, (str(lv), ""))[0]
        mapping[f"{_id}.risultato"] = f"{lv} - {nome_liv}"
        tipo = tipo_by_id.get(_id, "")
        campo = "commento" if tipo == "area" else "attivita"
        mapping[f"{_id}.{campo}"] = libreria.get((_id, lv), "")
    _derive_panoramica(mapping, scala)
    return mapping


def load_mapping(xlsx, regenerate=False):
    """xlsx: path o file-like. regenerate=True ricostruisce da Assessment+Libreria."""
    wb = load_workbook(xlsx, data_only=True)
    if not regenerate and "Mapping" in wb.sheetnames:
        m = _mapping_from_sheet(wb["Mapping"])
        if m:
            _derive_panoramica(m, _read_scala(wb))
            return m
    if "Assessment" in wb.sheetnames and "Libreria" in wb.sheetnames:
        return _mapping_from_levels(wb)            # panoramica già inclusa
    if "Mapping" in wb.sheetnames:
        m = _mapping_from_sheet(wb["Mapping"])
        _derive_panoramica(m, _read_scala(wb))
        return m
    raise ValueError(
        "Excel non valido: serve un foglio 'Mapping' (segnaposto|testo) "
        "oppure i fogli 'Assessment' + 'Libreria'."
    )


def mapping_from_survey(input_xlsx, mapping_xlsx):
    """
    Costruisce il mapping dai risultati della survey.

    input_xlsx  : file Excel wide (riga 1 = id delle aree/sottogruppi come
                  intestazioni colonna, riga 2 = livello 1-4 di ciascuno)
    mapping_xlsx: deia_mapping.xlsx con i fogli Libreria e Scala
    """
    # --- leggi Scala e Libreria dal file di mapping --------------------------
    wb_map = load_workbook(mapping_xlsx, data_only=True)

    scala = _read_scala(wb_map)  # {livello: (nome, descrizione)}

    libreria = {}
    tipo_by_id = {}
    if "Libreria" in wb_map.sheetnames:
        for rec in _records(wb_map["Libreria"]):
            _id = str(rec.get("id", "")).strip()
            lv = rec.get("livello")
            if not _id or lv is None:
                continue
            libreria[(_id, int(lv))] = "" if rec.get("testo") is None else str(rec["testo"])
            tipo_by_id[_id] = str(rec.get("tipo", "")).strip().lower()

    # --- leggi i risultati della survey (input, formato wide) ----------------
    wb_in = load_workbook(input_xlsx, data_only=True)
    ws_in = wb_in.active

    mapping = {}
    for _id, lv in _wide_record(ws_in).items():
        if lv is None:
            continue
        lv = int(lv)
        tipo = tipo_by_id.get(_id, "")
        nome_liv = scala.get(lv, (str(lv), ""))[0]
        mapping[f"{_id}.risultato"] = f"{lv} - {nome_liv}"
        campo = "commento" if tipo == "area" else "attivita"
        mapping[f"{_id}.{campo}"] = libreria.get((_id, lv), "")

    _derive_panoramica(mapping, scala)
    return mapping


# ----------------------------------------------------------------- PPTX traversal
def _iter_shapes(shapes):
    """Itera ricorsivamente tutte le forme, entrando nei gruppi."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(shape.shapes)
        else:
            yield shape


def _iter_text_frames(prs):
    for slide in prs.slides:
        for shape in _iter_shapes(slide.shapes):
            if shape.has_text_frame:
                yield shape.text_frame
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        yield cell.text_frame


def _apply_to_text(text, mapping, used, unresolved):
    def repl(m):
        key = m.group(1).strip()
        if key in mapping:
            used.add(key)
            return str(mapping[key])
        unresolved.add(key)
        return m.group(0)
    return PLACEHOLDER_RE.sub(repl, text)


def _replace_in_text_frame(tf, mapping, used, unresolved):
    for para in tf.paragraphs:
        runs = para.runs
        full = "".join(r.text for r in runs)
        if "{{" not in full:
            continue
        new = _apply_to_text(full, mapping, used, unresolved)
        if new == full:
            continue
        if runs:
            runs[0].text = new          # tutto il testo nel primo run
            for r in runs[1:]:
                r.text = ""             # svuota i run successivi
        else:
            para.text = new


# --------------------------------------------------------- radar dinamico
ORANGE_HEX = "#E86A17"
DARK_HEX = "#1A1A1A"
RADAR_RE = re.compile(r"\{\{\s*radar\.aree\s*\}\}")


def _radar_image(scores):
    """scores: [(label, valore 0..4), ...]. Ritorna BytesIO di un PNG radar."""
    labels = [s[0] for s in scores]
    vals = [min(s[1], 4) for s in scores]
    n = len(labels)
    angles = [i / n * 2 * math.pi for i in range(n)]
    angles_c = angles + angles[:1]
    values_c = vals + vals[:1]

    fig = _plt.figure(figsize=(6, 6), dpi=150)
    ax = fig.add_subplot(111, polar=True)
    ax.set_theta_offset(math.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_ylim(0, 4)
    ax.set_yticks([1, 2, 3, 4])
    ax.set_yticklabels(["1", "2", "3", "4"], fontsize=9, color="#999999")
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=12, fontweight="bold", color=DARK_HEX)
    ax.tick_params(axis="x", pad=16)
    ax.plot(angles_c, values_c, color=ORANGE_HEX, linewidth=2.2)
    ax.fill(angles_c, values_c, color=ORANGE_HEX, alpha=0.22)
    ax.scatter(angles, vals, color=ORANGE_HEX, s=35, zorder=5)
    for ang, val in zip(angles, vals):
        ax.text(ang, val + 0.28, str(val), color=ORANGE_HEX,
                fontsize=10, fontweight="bold", ha="center", va="center")
    ax.spines["polar"].set_color("#CCCCCC")
    ax.grid(color="#DDDDDD")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", transparent=True, bbox_inches="tight", pad_inches=0.2)
    _plt.close(fig)
    buf.seek(0)
    return buf


def _install_dynamic_images(prs, mapping):
    """Sostituisce i marcatori {{radar.aree}} con un'immagine radar generata
    dai punteggi delle aree presenti nel mapping."""
    if not _HAS_MPL:
        return
    for slide in prs.slides:
        for shape in list(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            full = "".join(r.text for p in shape.text_frame.paragraphs for r in p.runs)
            if not RADAR_RE.search(full):
                continue
            img = _radar_image(_area_scores(mapping))
            l, t, w, h = shape.left, shape.top, shape.width, shape.height
            size = min(w, h)                          # radar quadrato, centrato
            left = l + (w - size) // 2
            top = t + (h - size) // 2
            shape._element.getparent().remove(shape._element)
            slide.shapes.add_picture(img, left, top, size, size)


def extract_placeholders(pptx):
    prs = Presentation(pptx)
    found = set()
    for tf in _iter_text_frames(prs):
        for para in tf.paragraphs:
            full = "".join(r.text for r in para.runs)
            for m in PLACEHOLDER_RE.finditer(full):
                key = m.group(1).strip()
                if key.startswith("radar"):          # marcatore immagine, non testo
                    continue
                found.add(key)
    return found


def fill_pptx(pptx, mapping):
    """Compila il pptx. Ritorna (BytesIO, stats)."""
    prs = Presentation(pptx)
    _install_dynamic_images(prs, mapping)            # prima le immagini dinamiche
    used, unresolved = set(), set()
    for tf in _iter_text_frames(prs):
        _replace_in_text_frame(tf, mapping, used, unresolved)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    stats = {
        "sostituiti": sorted(used),
        "non_risolti": sorted(unresolved),
        "n_sostituiti": len(used),
        "n_non_risolti": len(unresolved),
        "inutilizzati": sorted(set(mapping) - used),
    }
    return buf, stats
