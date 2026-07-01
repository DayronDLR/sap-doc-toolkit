#!/usr/bin/env bash
# build-doc.sh — Genera el documento .docx desde el Markdown maestro
# Uso: bash build-doc.sh [--no-diagrams]
#
# Requisitos:
#   - pandoc (brew install pandoc / apt install pandoc)
#   - Opcional: mermaid-cli (npm install -g @mermaid-js/mermaid-cli)
#     Fallback automático a kroki.io API (requiere internet)
#
# Output: HOCOL-Workflow-Reservas-PM-Arquitectura.docx

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

PROJECT="HOCOL-Workflow-Reservas-PM"
MD_SRC="${PROJECT}-Arquitectura.md"
MD_TMP="${PROJECT}-Arquitectura.tmp.md"
DOCX_OUT="${PROJECT}-Arquitectura.docx"
IMG_DIR="diagrams"

NO_DIAGRAMS=false
for arg in "$@"; do
  [[ "$arg" == "--no-diagrams" ]] && NO_DIAGRAMS=true
done

echo "🔨 Build doc — HOCOL Workflow Reservas PM"

if ! command -v pandoc &> /dev/null; then
  echo "❌ pandoc no instalado. Instalar con: brew install pandoc (macOS) o apt install pandoc (Linux)"
  exit 1
fi

if [[ ! -f "$MD_SRC" ]]; then
  echo "❌ No se encuentra $MD_SRC"
  exit 1
fi

mkdir -p "$IMG_DIR"

INPUT_FILE="$MD_SRC"

if [[ "$NO_DIAGRAMS" == "false" ]]; then
  echo "📊 Extrayendo y renderizando diagramas Mermaid..."

  python3 << 'PYEOF'
import re, os, base64, urllib.request, urllib.parse, sys, zlib, json

src = open("HOCOL-Workflow-Reservas-PM-Arquitectura.md").read()
pattern = re.compile(r"```mermaid\n(.*?)\n```", re.DOTALL)
matches = list(pattern.finditer(src))
print(f"  Encontrados {len(matches)} diagramas Mermaid")

img_dir = "diagrams"
os.makedirs(img_dir, exist_ok=True)

mmdc_available = os.system("command -v mmdc > /dev/null 2>&1") == 0

def render_kroki(mmd_content, png_file):
    """Kroki.io — método 1: POST plain text"""
    try:
        req = urllib.request.Request(
            "https://kroki.io/mermaid/png",
            data=mmd_content.encode("utf-8"),
            headers={"Content-Type": "text/plain", "User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            with open(png_file, "wb") as f:
                f.write(resp.read())
        return True
    except Exception as e:
        return False

def render_kroki_get(mmd_content, png_file):
    """Kroki.io — método 2: GET URL con zlib+base64"""
    try:
        compressed = zlib.compress(mmd_content.encode("utf-8"), 9)
        encoded = base64.urlsafe_b64encode(compressed).decode("ascii")
        url = f"https://kroki.io/mermaid/png/{encoded}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            with open(png_file, "wb") as f:
                f.write(resp.read())
        return True
    except Exception as e:
        return False

successful = []
for i, m in enumerate(matches, start=1):
    mmd_content = m.group(1)
    mmd_file = f"{img_dir}/diagram_{i:02d}.mmd"
    png_file = f"{img_dir}/diagram_{i:02d}.png"

    with open(mmd_file, "w") as f:
        f.write(mmd_content)

    rendered = False
    method = ""

    if mmdc_available:
        ret = os.system(f"mmdc -i {mmd_file} -o {png_file} -t neutral -w 1400 -b white -s 2 2>/dev/null")
        if ret == 0 and os.path.exists(png_file):
            rendered = True
            method = "mmdc"

    if not rendered:
        if render_kroki(mmd_content, png_file):
            rendered = True
            method = "kroki POST"

    if not rendered:
        if render_kroki_get(mmd_content, png_file):
            rendered = True
            method = "kroki GET"

    if rendered:
        print(f"  ✓ diagram_{i:02d}.png ({method})")
        successful.append(i)
    else:
        print(f"  ⚠ diagram_{i:02d}: no se pudo renderizar (continuando sin él)")

# Reemplazar bloques mermaid por imágenes solo si se renderizaron
def replace_block(match, counter=[0]):
    counter[0] += 1
    if counter[0] in successful:
        return f"![Diagrama {counter[0]}]({img_dir}/diagram_{counter[0]:02d}.png)"
    else:
        return f"```mermaid\n{match.group(1)}\n```"  # mantener bloque tal cual

new_src = pattern.sub(replace_block, src)

with open("HOCOL-Workflow-Reservas-PM-Arquitectura.tmp.md", "w") as f:
    f.write(new_src)

print(f"  ✓ {len(successful)}/{len(matches)} diagramas renderizados")
PYEOF

  INPUT_FILE="$MD_TMP"
fi

echo "📄 Generando $DOCX_OUT con pandoc..."

# Aplicar identidad visual del cliente desde client-theme.yaml
THEME_FILE="client-theme.yaml"
REFERENCE_DOC="reference.docx"
if [[ -f "$THEME_FILE" ]]; then
  echo "🎨 Aplicando theme cliente desde $THEME_FILE"
  python3 apply-theme.py "$THEME_FILE" || {
    echo "⚠ apply-theme.py falló — continuando sin theme cliente"
  }
fi

PANDOC_REF_ARG=""
if [[ -f "$REFERENCE_DOC" ]]; then
  PANDOC_REF_ARG="--reference-doc=$REFERENCE_DOC"
fi

pandoc "$INPUT_FILE" \
  -o "$DOCX_OUT" \
  $PANDOC_REF_ARG \
  --toc --toc-depth=3 \
  -V geometry:margin=2cm \
  -V mainfont="Calibri" \
  -V monofont="Consolas" \
  --metadata title="HOCOL — Arquitectura y Estimación Workflow Reservas PM" \
  --metadata author="Tivit — SAP Tech Lead" \
  --metadata date="$(date +%Y-%m-%d)"

# Post-procesamiento: ajustar layout de tablas a autofit
if [[ -f "postprocess-docx.py" ]]; then
  echo "🔧 Ajustando layout de tablas (autofit)"
  python3 postprocess-docx.py "$DOCX_OUT" || echo "⚠ Post-proceso falló"
fi

rm -f "$MD_TMP"

echo ""
echo "✅ Documento generado: $DOCX_OUT"
echo "   Diagramas en: $IMG_DIR/"
echo ""
echo "📦 Para entregar al cliente:"
echo "   - $DOCX_OUT"
echo "   - ${PROJECT}-architecture.drawio (abrir en https://app.diagrams.net)"
echo ""
echo "💡 Tip: para regenerar sin internet (sin diagramas), usar:"
echo "   bash build-doc.sh --no-diagrams"
