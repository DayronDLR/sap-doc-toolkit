# sap-doc-toolkit

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

**English** · [Español](#español)

MIT toolchain for the **branded build** of SAP technical documentation: generates
`.docx` / `.pptx` with the client's visual identity (palette, fonts, logos) and
**draw.io architecture diagrams**. It's the optional companion to the `sap-doc`
agent of the
[SAP Enterprise Stack](https://github.com/DayronDLR/sap-enterprise-stack-plugin):
the agent writes the **content**, this toolkit does the **branded build**.

## Prerequisites — install everything

All free and cross-platform:

| Tool | For | Required |
| --- | --- | --- |
| **pandoc** 3.x | `.docx` generation | ✅ |
| **Python 3** + `python-pptx` + `lxml` | `.pptx` and theming | ✅ |
| **mermaid-cli** (`mmdc`) | Mermaid diagrams → PNG (in `build-doc.sh`) | ✅ for Mermaid |
| **Graphviz** (`dot`) | some auto-layout diagrams | optional |
| SAP BTP icon library | icons in draw.io diagrams | optional (proprietary SAP — you provide it) |

### macOS (Homebrew)

```bash
brew install pandoc python graphviz
pip3 install python-pptx lxml
pnpm add -g @mermaid-js/mermaid-cli    # or: npm i -g @mermaid-js/mermaid-cli
```

### Linux — Debian / Ubuntu

```bash
sudo apt-get update && sudo apt-get install -y pandoc python3 python3-pip graphviz
pip3 install python-pptx lxml
pnpm add -g @mermaid-js/mermaid-cli
```

### Linux — RHEL / Fedora

```bash
sudo dnf install -y pandoc python3 python3-pip graphviz
pip3 install python-pptx lxml
pnpm add -g @mermaid-js/mermaid-cli
```

### Windows (winget / PowerShell)

```powershell
winget install JohnMacFarlane.Pandoc Python.Python.3.12 Graphviz.Graphviz
pip install python-pptx lxml
pnpm add -g @mermaid-js/mermaid-cli
```

> **pandoc without sudo** (SAP BAS / restricted container) — the binary is static:
>
> ```bash
> curl -L -o pandoc.tar.gz https://github.com/jgm/pandoc/releases/download/3.9/pandoc-3.9-linux-amd64.tar.gz
> mkdir -p ~/.local/bin && tar xvzf pandoc.tar.gz --strip-components 2 -C ~/.local/bin
> echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc && source ~/.bashrc
> ```

Verify:

```bash
pandoc --version && python3 --version && mmdc --version
```

## Setup (once)

```bash
# 1. Your palette/theme (copy the example and edit it)
cp tools/client-theme.example.yaml docs/architecture/client-theme.yaml

# 2. pandoc's base reference.docx (not distributed — generate it)
pandoc --print-default-data-file reference.docx > tools/reference-base.docx

# 3. (Optional) SAP BTP icons in sap-btp-icons/extracted-icons.json (yours)
```

## Usage

```bash
# Apply the client's visual identity → reference.docx + theme.json + theme_colors.py
python3 tools/apply-theme.py docs/architecture/client-theme.yaml

# Branded Word build (Mermaid→PNG→pandoc→.docx + post-process)
bash tools/build-doc.sh

# Branded .pptx presentation
python3 tools/build-pptx.py

# Example draw.io diagrams (SAP icons needed for the SAP shapes)
python3 tools/generate-examples.py
```

## What's included

| Script | What it does |
| --- | --- |
| `apply-theme.py` | Generates `reference.docx`, `theme_colors.py` and `theme.json` from your `client-theme.yaml` |
| `build-doc.sh` | Mermaid→PNG→pandoc→`.docx`, applies the theme and post-processes |
| `build-pptx.py` | Client-parametrized `.pptx` presentation |
| `postprocess-docx.py` | Fixes `.docx` tables (autofit, minimum widths) |
| `sap-drawio-generator.py` | SAP BTP draw.io diagram generator (`SAPDiagramBuilder`) |
| `generate-examples.py` | Generates example diagrams |
| `client-theme.example.yaml` | Client palette template (colors, fonts, logos) |

## License

**MIT** (see [`LICENSE`](LICENSE)). The SAP BTP icons are **not** distributed (see
[`NOTICE`](NOTICE)). Companion to the
[SAP Enterprise Stack plugin](https://github.com/DayronDLR/sap-enterprise-stack-plugin).

---

<a name="español"></a>

## Español · [English](#sap-doc-toolkit)

Toolchain **MIT** para el *branded build* de documentación técnica SAP: genera
`.docx` / `.pptx` con la **identidad visual del cliente** (paleta, fuentes, logos)
y **diagramas de arquitectura draw.io**. Es el companion opcional del agente
`sap-doc` del
[SAP Enterprise Stack](https://github.com/DayronDLR/sap-enterprise-stack-plugin):
el agente genera el **contenido**, este toolkit hace el **build con branding**.

## Requisitos — instalá todo

Todo gratis y multiplataforma:

| Herramienta | Para qué | Requerida |
| --- | --- | --- |
| **pandoc** 3.x | Genera el `.docx` | ✅ |
| **Python 3** + `python-pptx` + `lxml` | `.pptx` y theming | ✅ |
| **mermaid-cli** (`mmdc`) | Diagramas Mermaid → PNG (en `build-doc.sh`) | ✅ para Mermaid |
| **Graphviz** (`dot`) | Algunos diagramas con auto-layout | opcional |
| Librería de iconos SAP BTP | Iconos en los diagramas draw.io | opcional (propietarios de SAP — los pones tú) |

### macOS (Homebrew)

```bash
brew install pandoc python graphviz
pip3 install python-pptx lxml
pnpm add -g @mermaid-js/mermaid-cli    # o: npm i -g @mermaid-js/mermaid-cli
```

### Linux — Debian / Ubuntu

```bash
sudo apt-get update && sudo apt-get install -y pandoc python3 python3-pip graphviz
pip3 install python-pptx lxml
pnpm add -g @mermaid-js/mermaid-cli
```

### Linux — RHEL / Fedora

```bash
sudo dnf install -y pandoc python3 python3-pip graphviz
pip3 install python-pptx lxml
pnpm add -g @mermaid-js/mermaid-cli
```

### Windows (winget / PowerShell)

```powershell
winget install JohnMacFarlane.Pandoc Python.Python.3.12 Graphviz.Graphviz
pip install python-pptx lxml
pnpm add -g @mermaid-js/mermaid-cli
```

> **pandoc sin sudo** (SAP BAS / contenedor restringido) — el binario es estático:
>
> ```bash
> curl -L -o pandoc.tar.gz https://github.com/jgm/pandoc/releases/download/3.9/pandoc-3.9-linux-amd64.tar.gz
> mkdir -p ~/.local/bin && tar xvzf pandoc.tar.gz --strip-components 2 -C ~/.local/bin
> echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc && source ~/.bashrc
> ```

Verificá: `pandoc --version && python3 --version && mmdc --version`

## Setup (una vez)

```bash
# 1. Tu paleta/tema (copia el ejemplo y edítalo)
cp tools/client-theme.example.yaml docs/architecture/client-theme.yaml

# 2. reference.docx base de pandoc (no se distribuye — se genera)
pandoc --print-default-data-file reference.docx > tools/reference-base.docx

# 3. (Opcional) iconos SAP BTP en sap-btp-icons/extracted-icons.json (tuyos)
```

## Uso

```bash
python3 tools/apply-theme.py docs/architecture/client-theme.yaml   # → reference.docx + theme.json
bash tools/build-doc.sh          # Mermaid→PNG→pandoc→.docx + post-proceso
python3 tools/build-pptx.py      # presentación .pptx branded
python3 tools/generate-examples.py  # diagramas de ejemplo draw.io
```

## Licencia

**MIT** (ver [`LICENSE`](LICENSE)). Los iconos SAP BTP **no** se distribuyen (ver
[`NOTICE`](NOTICE)). Companion del
[plugin SAP Enterprise Stack](https://github.com/DayronDLR/sap-enterprise-stack-plugin).
