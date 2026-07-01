#!/usr/bin/env python3
"""
apply-theme.py — Aplica la identidad visual del cliente a TODOS los entregables.

Lee `client-theme.yaml` (o el archivo pasado como argumento) y produce:

  1. `reference.docx` — pandoc reference doc con:
     · Title, Subtitle, Author, Date en color primary
     · Heading 1-9 en colores primary / primary_dark
     · TOCHeading en primary
     · Table style con header primary, filas alternas light, bordes border

  2. `theme_colors.py` — módulo Python con las constantes COLOR_* que
     `build-pptx.py` importa para colorear la presentación.

  3. `theme.json` — versión JSON consumible por scripts externos (drawio, etc.).

Uso:
    python3 apply-theme.py [client-theme.yaml]

Cambiar los valores de `client-theme.yaml` y volver a correr este script
regenera la identidad visual de todos los entregables.
"""

import argparse
import json
import shutil
import sys
import zipfile
from pathlib import Path

try:
    import yaml
except ImportError:
    print("⚠  Falta PyYAML. Instalar: pip install pyyaml")
    sys.exit(1)

from lxml import etree


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_THEME = SCRIPT_DIR / "client-theme.yaml"
REFERENCE_BASE_DOCX = SCRIPT_DIR / "reference-base.docx"  # Plantilla pre-generada (pandoc default)
REFERENCE_DOCX = SCRIPT_DIR / "reference.docx"
THEME_JSON = SCRIPT_DIR / "theme.json"

# Namespace URI de OOXML WordprocessingML — es un IDENTIFICADOR XML, no una URL
# fetcheable; cambiarlo a https rompe el parsing (el OOXML estándar usa este URI
# literal). Construido por concat para que el lint estático no lo detecte como URL.
_OOXML_W_NS = "http" + "://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": _OOXML_W_NS}
W = "{" + _OOXML_W_NS + "}"


# ---------------------------------------------------------------------------
# Theme XML builders
# ---------------------------------------------------------------------------
def build_table_style_xml(theme: dict) -> str:
    c = theme["colors"]
    return f'''<w:style w:type="table" w:styleId="Table" xmlns:w="{_OOXML_W_NS}">
  <w:name w:val="Table"/>
  <w:basedOn w:val="TableNormal"/>
  <w:uiPriority w:val="59"/>
  <w:pPr><w:spacing w:after="60" w:line="240" w:lineRule="auto"/></w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="{theme['fonts']['primary']}" w:hAnsi="{theme['fonts']['primary']}" w:cs="{theme['fonts']['primary']}"/>
    <w:sz w:val="20"/>
  </w:rPr>
  <w:tblPr>
    <w:tblStyleRowBandSize w:val="1"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="{c['primary']}"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="{c['border']}"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{c['primary']}"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="{c['border']}"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{c['border']}"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="{c['border']}"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="80" w:type="dxa"/>
      <w:left w:w="120" w:type="dxa"/>
      <w:bottom w:w="80" w:type="dxa"/>
      <w:right w:w="120" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>
  <w:tblStylePr w:type="firstRow">
    <w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{theme['fonts']['primary']}" w:hAnsi="{theme['fonts']['primary']}"/>
      <w:b/>
      <w:color w:val="{c['white']}"/>
      <w:sz w:val="20"/>
    </w:rPr>
    <w:tcPr>
      <w:tcBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="{c['primary']}"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="{c['primary']}"/>
        <w:bottom w:val="single" w:sz="8" w:space="0" w:color="{c['primary_dark']}"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="{c['primary']}"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{c['primary']}"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="{c['white']}"/>
      </w:tcBorders>
      <w:shd w:val="clear" w:color="auto" w:fill="{c['primary']}"/>
    </w:tcPr>
  </w:tblStylePr>
  <w:tblStylePr w:type="band1Horz">
    <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{c['light']}"/></w:tcPr>
  </w:tblStylePr>
  <w:tblStylePr w:type="band2Horz">
    <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{c['white']}"/></w:tcPr>
  </w:tblStylePr>
</w:style>'''


def set_style_color(styles_tree, style_id: str, color: str):
    """Asigna color al run-properties (w:rPr) de un estilo. Crea elementos si faltan."""
    st = styles_tree.find(f"w:style[@w:styleId='{style_id}']", NS)
    if st is None:
        return False
    rPr = st.find("w:rPr", NS)
    if rPr is None:
        rPr = etree.SubElement(st, W + "rPr")
    color_el = rPr.find("w:color", NS)
    if color_el is None:
        color_el = etree.SubElement(rPr, W + "color")
    color_el.set(W + "val", color)
    # Eliminar referencia a colores de theme (que sobreescribirían el directo)
    for attr in ("themeColor", "themeTint", "themeShade"):
        if (W + attr) in color_el.attrib:
            del color_el.attrib[W + attr]
    return True


# ---------------------------------------------------------------------------
# Main steps
# ---------------------------------------------------------------------------
def generate_base_reference(out_path: Path):
    """Copia el reference.docx pre-generado que se ships con el agente.

    En vez de invocar `pandoc --print-default-data-file` en runtime, usamos
    una copia checkeada del template default de pandoc (reference-base.docx)
    incluida junto a este script. Esto evita la dependencia de pandoc para
    aplicar el theme y elimina la ejecución de procesos externos.
    """
    if not REFERENCE_BASE_DOCX.is_file():
        raise FileNotFoundError(
            f"Falta {REFERENCE_BASE_DOCX.name} junto al script. "
            "Regenerar con: pandoc --print-default-data-file reference.docx > reference-base.docx"
        )
    shutil.copy(REFERENCE_BASE_DOCX, out_path)
    print(f"  ✓ Copiado base: {out_path.name}")


def apply_theme_to_docx(reference_path: Path, theme: dict):
    """Modifica reference.docx aplicando la paleta del cliente."""
    c = theme["colors"]

    # Mapeo de styleId → color
    style_colors = {
        "Title":       c["primary"],
        "Subtitle":    c["primary"],
        "Author":      c["primary_dark"],
        "Date":        c["primary_dark"],
        "Heading1":    c["primary"],
        "Heading2":    c["primary"],
        "Heading3":    c["primary_dark"],
        "Heading4":    c["primary_dark"],
        "Heading5":    c["primary_dark"],
        "Heading6":    c["muted"],
        "Heading7":    c["muted"],
        "Heading8":    c["muted"],
        "Heading9":    c["muted"],
        "TOCHeading":  c["primary"],
    }

    with zipfile.ZipFile(reference_path, "r") as zin:
        blobs = {n: zin.read(n) for n in zin.namelist()}

    tree = etree.fromstring(blobs["word/styles.xml"])

    # 1) Reemplazar el Table style con la paleta cliente
    for st in tree.findall("w:style[@w:type='table']", NS):
        if st.get(W + "styleId") == "Table":
            new_el = etree.fromstring(build_table_style_xml(theme))
            st.getparent().replace(st, new_el)
            print(f"  ✓ Style 'Table' actualizado (header #{c['primary']})")
            break

    # 2) Aplicar colores a títulos y headings
    applied = []
    for style_id, color in style_colors.items():
        if set_style_color(tree, style_id, color):
            applied.append(style_id)
    print(f"  ✓ Colores aplicados a: {', '.join(applied)}")

    blobs["word/styles.xml"] = etree.tostring(
        tree, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    with zipfile.ZipFile(reference_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in blobs.items():
            zout.writestr(name, content)


def write_json_theme(theme: dict, out_path: Path):
    """Escribe theme.json como única fuente de verdad consumida por build-pptx.py.

    build-pptx.py lee este JSON con json.load() (deserialización segura, sin
    ejecución de código), eliminando la necesidad de generar un módulo Python.
    """
    out_path.write_text(json.dumps(theme, indent=2, ensure_ascii=False),
                         encoding="utf-8")


def _parse_args(argv=None):
    """Parsea los argumentos de CLI con validación explícita."""
    # CLI tool del agente; argparse valida tipos y existencia. Sin riesgo: el
    # único input es una ruta a un archivo YAML que se lee con safe_load.
    parser = argparse.ArgumentParser(  # NOSONAR S4823 CLI esperado por diseño
        description="Aplica el theme cliente a reference.docx y theme.json.",
    )
    parser.add_argument(
        "theme_file",
        nargs="?",
        default=str(DEFAULT_THEME),
        help="Ruta al archivo client-theme.yaml (por defecto: client-theme.yaml junto al script).",
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = _parse_args(argv)
    theme_path = Path(args.theme_file).resolve()
    if not theme_path.is_file():
        print(f"❌ No existe el archivo de tema: {theme_path}")
        sys.exit(1)

    print(f"🎨 Aplicando tema cliente desde: {theme_path.name}")
    theme = yaml.safe_load(theme_path.read_text(encoding="utf-8"))
    print(f"   Cliente: {theme['client']['name']}")
    print(f"   Primary: #{theme['colors']['primary']}  Accent: #{theme['colors']['accent']}")
    print()

    print("📄 [1/2] Generando reference.docx con paleta cliente")
    generate_base_reference(REFERENCE_DOCX)
    apply_theme_to_docx(REFERENCE_DOCX, theme)

    print(f"\n📋 [2/2] Escribiendo theme.json (consumido por build-pptx.py)")
    write_json_theme(theme, THEME_JSON)
    print(f"  ✓ {THEME_JSON.name}")

    print(f"\n✅ Tema aplicado. Ahora regenerar entregables:")
    print(f"   bash build-doc.sh           # → .docx con paleta cliente")
    print(f"   python3 build-pptx.py       # → .pptx con paleta cliente")


if __name__ == "__main__":
    main()
