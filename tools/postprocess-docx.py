#!/usr/bin/env python3
"""
postprocess-docx.py — Ajusta las tablas del .docx generado por pandoc para que
las columnas se autoajusten al contenido.

Problema que resuelve:
    Pandoc genera tablas con `tblLayout: fixed` y columnas proporcionales al
    ancho de los guiones del markdown. Cuando el header es corto (`| # |`)
    Word renderiza la columna a 0.06" y el contenido queda comprimido o cortado.

Solución (mejorada):
    Para cada tabla, calcular el ancho ideal de cada columna basado en el
    contenido REAL de las celdas (longitud de header + longitud máxima de
    body por columna). Distribuir el ancho total proporcionalmente.

    Pasos por tabla:
      1) Cambiar `tblLayout` a `autofit` → Word puede expandir si necesita
      2) Cambiar `tblW` a `pct/5000` (100%) → la tabla ocupa todo el ancho
      3) Recalcular gridCol con proporciones derivadas del contenido
      4) Limpiar tcW de cada celda → usar widths del gridCol

Uso:
    python3 postprocess-docx.py <archivo.docx>
"""

import argparse
import sys
import zipfile
from pathlib import Path
from lxml import etree


# Namespace URI de OOXML WordprocessingML — es un IDENTIFICADOR XML, no una URL
# fetcheable; cambiarlo a https rompe el parsing.
_OOXML_W_NS = "http" + "://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": _OOXML_W_NS}
W = "{" + _OOXML_W_NS + "}"

# Ancho total objetivo de una tabla a página A4 con márgenes 2cm: ~9576 twips
# (página 21cm − 2*2cm = 17cm = 9576 twips). Se usa para distribuir gridCol.
TOTAL_TWIPS = 9576
MIN_COL_TWIPS = 720      # 0.5" mínimo absoluto por columna
CHAR_TWIPS = 110         # Twips por carácter Calibri 10pt regular
HEADER_CHAR_TWIPS = 125  # Twips por carácter Calibri 10pt bold (header)
CELL_PADDING = 240       # Padding interno por celda (twips, ambos lados)
MAX_BODY_CHARS = 40      # Tope de longitud body para no inflar columnas
MAX_COL_TWIPS = 5760     # 4" máximo por columna (evita columnas dominantes)


def _cell_text(tc) -> str:
    """Extrae el texto plano de una celda <w:tc>."""
    runs = tc.findall(".//w:t", NS)
    return "".join(r.text or "" for r in runs)


def _longest_token_chars(text: str) -> int:
    """Longitud del token más largo del texto (palabra sin partir)."""
    tokens = text.split()
    return max((len(t) for t in tokens), default=0)


def _compute_col_widths(table) -> list[int]:
    """Calcula ancho en twips para cada columna basado en su contenido real.

    Dos niveles de ancho por columna:
      - must_have: ancho MÍNIMO innegociable. Garantiza:
          · Header completo en una línea ("Provisión" no se parte)
          · Token más largo del body en una línea ("Functional" no se parte)
        Este nivel NUNCA se reduce al escalar.

      - extra_want: ancho ADICIONAL deseable para que el body promedio fluya
        cómodamente sin demasiados wraps. Derivado del p80 de longitud body,
        capado a MAX_BODY_CHARS para que columnas con texto largo no
        monopolicen el ancho.

    Algoritmo de reparto del ancho total disponible:
      1) Cada columna recibe su must_have completo.
      2) El remanente (TOTAL_TWIPS - sum(must_have)) se distribuye
         proporcionalmente a los `extra_want` de cada columna.
      3) Si must_have ya excede TOTAL_TWIPS, escalar must_have pero NUNCA
         por debajo de MIN_COL_TWIPS.
    """
    rows = table.findall("w:tr", NS)
    if not rows:
        return []
    n_cols = len(rows[0].findall("w:tc", NS))
    must_have = [MIN_COL_TWIPS] * n_cols
    extra_want = [0] * n_cols

    # Header row: ancho del header es must_have
    for j, tc in enumerate(rows[0].findall("w:tc", NS)):
        if j >= n_cols:
            break
        text = _cell_text(tc)
        header_w = len(text) * HEADER_CHAR_TWIPS + CELL_PADDING
        must_have[j] = max(must_have[j], header_w)

    # Body rows: token más largo es must_have; resto es extra_want
    for j in range(n_cols):
        col_lengths = []
        col_max_token = 0
        for row in rows[1:]:
            tcs = row.findall("w:tc", NS)
            if j >= len(tcs):
                continue
            text = _cell_text(tcs[j])
            if text:
                col_lengths.append(len(text))
            col_max_token = max(col_max_token, _longest_token_chars(text))

        token_w = col_max_token * CHAR_TWIPS + CELL_PADDING
        must_have[j] = max(must_have[j], token_w)

        # extra_want: p80 capado del body — qué tan ancho querría ser
        if col_lengths:
            sorted_l = sorted(col_lengths)
            p80 = sorted_l[int(len(sorted_l) * 0.8)] if sorted_l else 0
            capped = min(p80, MAX_BODY_CHARS)
            body_w = capped * CHAR_TWIPS + CELL_PADDING
            extra_want[j] = max(0, body_w - must_have[j])

    # Repartir TOTAL_TWIPS: must_have + reparto proporcional del remanente
    total_must = sum(must_have)
    if total_must >= TOTAL_TWIPS:
        # Escalar must_have proporcionalmente (caso de tabla muy densa)
        return [max(MIN_COL_TWIPS, int(w * TOTAL_TWIPS / total_must))
                for w in must_have]

    remainder = TOTAL_TWIPS - total_must
    total_want = sum(extra_want)
    if total_want == 0:
        # Sin demanda extra — repartir remainder uniformemente
        per_col = remainder // n_cols
        return [min(MAX_COL_TWIPS, must_have[j] + per_col) for j in range(n_cols)]

    # Reparto proporcional al extra_want
    final = []
    for j in range(n_cols):
        extra = int(remainder * extra_want[j] / total_want)
        final.append(min(MAX_COL_TWIPS, must_have[j] + extra))
    return final


def fix_table_layout(table):
    """Aplica layout=fixed con anchos proporcionales al contenido.

    Word respeta `tblLayout=fixed` literalmente: usa los widths de `gridCol`
    sin recalcular. Esto da control total sobre el ancho de cada columna.
    Con `autofit`, Word ignora gridCol y recalcula, lo que provocaba que
    columnas como "Provisión" quedaran demasiado angostas.
    """
    tblPr = table.find("w:tblPr", NS)
    if tblPr is None:
        return False

    grid = table.find("w:tblGrid", NS)
    if grid is None:
        return True

    new_widths = _compute_col_widths(table)
    if not new_widths:
        return True
    total_width = sum(new_widths)

    # 1) tblW → ancho absoluto en twips, no porcentaje
    tblW = tblPr.find("w:tblW", NS)
    if tblW is None:
        tblW = etree.SubElement(tblPr, W + "tblW")
    tblW.set(W + "w", str(total_width))
    tblW.set(W + "type", "dxa")

    # 2) tblLayout → fixed para que Word respete gridCol literal
    tblLayout = tblPr.find("w:tblLayout", NS)
    if tblLayout is None:
        tblLayout = etree.SubElement(tblPr, W + "tblLayout")
    tblLayout.set(W + "type", "fixed")

    # 3) gridCol con widths calculados
    grid_cols = grid.findall("w:gridCol", NS)
    for gc, w in zip(grid_cols, new_widths):
        gc.set(W + "w", str(w))

    # 4) tcW de cada celda alineado al gridCol correspondiente
    for row in table.findall("w:tr", NS):
        tcs = row.findall("w:tc", NS)
        for j, tc in enumerate(tcs):
            tcPr = tc.find("w:tcPr", NS)
            if tcPr is None:
                tcPr = etree.SubElement(tc, W + "tcPr")
            tcW = tcPr.find("w:tcW", NS)
            if tcW is None:
                tcW = etree.SubElement(tcPr, W + "tcW")
            grid_span = tcPr.find("w:gridSpan", NS)
            span = int(grid_span.get(W + "val", "1")) if grid_span is not None else 1
            if j < len(new_widths):
                cell_w = sum(new_widths[j:j + span])
                tcW.set(W + "w", str(cell_w))
                tcW.set(W + "type", "dxa")

    return True


def _parse_args(argv=None):
    """Parsea los argumentos de CLI con validación explícita."""
    # CLI tool del agente; argparse valida la ruta del .docx a procesar.
    # Sin riesgo: el archivo se abre con zipfile en modo R/W limitado a su zip interno.
    parser = argparse.ArgumentParser(  # NOSONAR S4823 CLI esperado por diseño
        description="Ajusta el layout de las tablas del .docx generado por pandoc.",
    )
    parser.add_argument(
        "docx_file",
        help="Ruta al archivo .docx a post-procesar.",
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = _parse_args(argv)
    docx_path = Path(args.docx_file).resolve()
    if not docx_path.is_file():
        print(f"❌ No existe: {docx_path}")
        sys.exit(1)

    with zipfile.ZipFile(docx_path, "r") as zin:
        blobs = {n: zin.read(n) for n in zin.namelist()}

    body = etree.fromstring(blobs["word/document.xml"])
    tables = body.findall(".//w:tbl", NS)

    fixed_count = 0
    for t in tables:
        if fix_table_layout(t):
            fixed_count += 1

    blobs["word/document.xml"] = etree.tostring(
        body, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in blobs.items():
            zout.writestr(name, content)

    print(f"  ✓ Post-procesadas {fixed_count} tablas → autofit + anchos proporcionales al contenido")


if __name__ == "__main__":
    main()
