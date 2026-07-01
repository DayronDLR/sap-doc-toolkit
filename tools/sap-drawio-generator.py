"""
SAP BTP Solution Diagram Generator for draw.io
Follows official SAP BTP Solution Diagram Guidelines (Horizon 2023)

Rules enforced:
1. Icon labels = PLAIN TEXT, format via style attributes (never HTML in icon values)
2. html=1 in style ONLY when value has HTML tags
3. &#10; for multiline in icons
4. SAP official colors (Primary, Semantic, Accent)
5. Legend box always present
6. Orthogonal connectors
7. arcSize=24, absoluteArcSize=1, strokeWidth=1.5
"""

import json, os

# Load extracted SAP BTP icons — search multiple locations
_SEARCH = [
    os.path.join(os.path.dirname(__file__), "..", "sap-btp-icons", "extracted-icons.json"),
    os.path.join(os.path.dirname(__file__), "sap-btp-icons", "extracted-icons.json"),
    os.path.join(os.getcwd(), "sap-btp-icons", "extracted-icons.json"),
]
ICONS_PATH = next((p for p in _SEARCH if os.path.exists(p)), None)
if ICONS_PATH:
    with open(ICONS_PATH) as f:
        SAP_ICONS = json.load(f)
else:
    # Degradación con gracia: la librería de iconos SAP BTP es un asset propietario
    # de SAP y NO se distribuye con el toolkit MIT. Sin ella, el generador sigue
    # funcionando para diagramas sin iconos (los lookups devuelven "").
    import sys
    print(
        "[sap-drawio-generator] Aviso: no se encontró sap-btp-icons/extracted-icons.json. "
        "Los diagramas se generan SIN iconos SAP BTP. Colocá la librería de iconos en "
        "sap-btp-icons/ para habilitarlos.",
        file=sys.stderr,
    )
    SAP_ICONS = {}

# SAP BTP Color System
COLORS = {
    # Primary
    "btp_border": "#0070F2", "btp_fill": "#EBF8FF",
    "ext_border": "#475E75", "ext_fill": "#F5F6F7",
    "title_color": "#1D2D3E", "text_color": "#556B82",
    # Semantic
    "green_border": "#188918", "green_fill": "#F5FAE5",
    "orange_border": "#C35500", "orange_fill": "#FFF8D6",
    "red_border": "#D20A0A", "red_fill": "#FFEAF4",
    # Accent
    "teal_border": "#07838F", "teal_fill": "#DAFDF5",
    "indigo_border": "#5D36FF", "indigo_fill": "#F1ECFF",
    "pink_border": "#CC00DC", "pink_fill": "#FFF0FA",
    # UI
    "legend_border": "#eaecee", "white": "#FFFFFF",
}

class SAPDiagramBuilder:
    def __init__(self, title, diagram_id="diagram-1"):
        self.title = title
        self.diagram_id = diagram_id
        self.cells = []
        self._id_counter = 100

    def _next_id(self, prefix="c"):
        self._id_counter += 1
        return f"{prefix}-{self._id_counter}"

    def _svg_for(self, key):
        """Get SVG base64 for a SAP BTP service icon."""
        if key in SAP_ICONS:
            return SAP_ICONS[key]["svg_base64"]
        return None

    # === AREAS ===
    # NOTE: container=0 on all areas. draw.io renders connectors between
    # any cells regardless of container status. Using container=0 avoids
    # drag-into-container behavior that can break layouts.

    def add_btp_area(self, x, y, w, h, cid=None):
        """BTP global container (outermost blue)"""
        cid = cid or self._next_id("btp")
        self.cells.append(f'''<mxCell id="{cid}" value="" style="rounded=1;whiteSpace=wrap;html=1;strokeColor={COLORS['btp_border']};fillColor={COLORS['btp_fill']};arcSize=24;absoluteArcSize=1;strokeWidth=1.5;container=0;" parent="1" vertex="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" /></mxCell>''')
        return cid

    def add_subaccount(self, x, y, w, h, label="Subaccount", subtitle="Multi-Cloud", cid=None):
        """Subaccount container (gray border, white fill) with title"""
        cid = cid or self._next_id("sub")
        val = f'{label}&lt;div&gt;&lt;font style=&quot;font-size: 12px;&quot;&gt;&lt;span style=&quot;font-weight: normal;&quot;&gt;{subtitle}&lt;/span&gt;&lt;/font&gt;&lt;/div&gt;'
        self.cells.append(f'''<mxCell id="{cid}" value="{val}" style="rounded=1;whiteSpace=wrap;html=1;strokeColor={COLORS['ext_border']};fillColor={COLORS['white']};arcSize=24;absoluteArcSize=1;strokeWidth=1.5;verticalAlign=top;align=left;fontSize=16;fontStyle=1;spacingLeft=10;spacingTop=10;fontFamily=Arial;fontColor={COLORS['title_color']};container=0;" parent="1" vertex="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" /></mxCell>''')
        return cid

    def add_service_area(self, x, y, w, h, label="", cid=None):
        """Application service area (blue border, light blue fill)"""
        cid = cid or self._next_id("svc")
        lbl = self._esc_plain(label) if label else ""
        lbl_style = "verticalAlign=top;align=left;fontSize=14;fontStyle=1;spacingLeft=10;spacingTop=8;fontFamily=Arial;fontColor=#1D2D3E;" if label else ""
        self.cells.append(f'''<mxCell id="{cid}" value="{lbl}" style="rounded=1;whiteSpace=wrap;html=1;strokeColor={COLORS['btp_border']};fillColor={COLORS['btp_fill']};arcSize=24;absoluteArcSize=1;strokeWidth=1.5;container=0;{lbl_style}" parent="1" vertex="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" /></mxCell>''')
        return cid

    def add_inner_area(self, x, y, w, h, cid=None):
        """Inner service box (blue border, white fill)"""
        cid = cid or self._next_id("inn")
        self.cells.append(f'''<mxCell id="{cid}" value="" style="rounded=1;whiteSpace=wrap;html=1;strokeColor={COLORS['btp_border']};fillColor={COLORS['white']};arcSize=16;absoluteArcSize=1;strokeWidth=1.5;container=0;" parent="1" vertex="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" /></mxCell>''')
        return cid

    def add_external_area(self, x, y, w, h, label="", subtitle="", cid=None):
        """External/non-SAP area (gray border, gray fill)"""
        cid = cid or self._next_id("ext")
        if label and subtitle:
            val = f'{label}&lt;div&gt;&lt;font style=&quot;font-size: 12px;&quot;&gt;&lt;span style=&quot;font-weight: normal;&quot;&gt;{subtitle}&lt;/span&gt;&lt;/font&gt;&lt;/div&gt;'
            self.cells.append(f'''<mxCell id="{cid}" value="{val}" style="rounded=1;whiteSpace=wrap;html=1;strokeColor={COLORS['ext_border']};fillColor={COLORS['ext_fill']};arcSize=24;absoluteArcSize=1;strokeWidth=1.5;verticalAlign=top;align=left;fontSize=16;fontStyle=1;spacingLeft=10;spacingTop=10;fontFamily=Arial;fontColor={COLORS['title_color']};container=0;" parent="1" vertex="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" /></mxCell>''')
        elif label:
            self.cells.append(f'''<mxCell id="{cid}" value="{self._esc_plain(label)}" style="rounded=1;whiteSpace=wrap;html=1;strokeColor={COLORS['ext_border']};fillColor={COLORS['ext_fill']};arcSize=24;absoluteArcSize=1;strokeWidth=1.5;verticalAlign=top;align=left;fontSize=16;fontStyle=1;spacingLeft=10;spacingTop=10;fontFamily=Arial;fontColor={COLORS['title_color']};container=0;" parent="1" vertex="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" /></mxCell>''')
        else:
            self.cells.append(f'''<mxCell id="{cid}" value="" style="rounded=1;whiteSpace=wrap;html=1;strokeColor={COLORS['ext_border']};fillColor={COLORS['ext_fill']};arcSize=24;absoluteArcSize=1;strokeWidth=1.5;container=0;" parent="1" vertex="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" /></mxCell>''')
        return cid

    def add_network_column(self, x, y, w, h, cid=None):
        """NETWORK column (right side)"""
        cid = cid or self._next_id("net")
        self.cells.append(f'''<mxCell id="{cid}" value="" style="rounded=1;whiteSpace=wrap;html=1;strokeColor={COLORS['ext_border']};fillColor={COLORS['ext_fill']};arcSize=24;absoluteArcSize=1;strokeWidth=1.5;container=0;" parent="1" vertex="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" /></mxCell>''')
        lbl = self._next_id("nlbl")
        self.cells.append(f'''<mxCell id="{lbl}" value="NETWORK" style="text;html=1;align=center;verticalAlign=middle;resizable=0;points=[];autosize=1;strokeColor=none;fillColor=none;fontSize=14;fontColor={COLORS['text_color']};fontStyle=1;fontFamily=Arial;" parent="1" vertex="1"><mxGeometry x="{x+w//2-40}" y="{y-25}" width="80" height="25" as="geometry" /></mxCell>''')
        return cid

    # === ICONS ===
    def add_icon(self, x, y, icon_key, label, size=32, parent="1", cid=None, bold=False):
        """SAP BTP service icon with PLAIN TEXT label. Format via style only."""
        cid = cid or self._next_id("ico")
        svg = self._svg_for(icon_key)
        if not svg:
            # Fallback: simple circle with label
            return self.add_text(x, y, 80, 50, label, bold=bold, color=COLORS['title_color'])

        font_style = "1" if bold else "0"
        self.cells.append(f'''<mxCell id="{cid}" value="{self._esc_plain(label)}" style="shape=image;verticalLabelPosition=bottom;labelBackgroundColor=default;verticalAlign=top;aspect=fixed;imageAspect=0;image=data:image/svg+xml,{svg};strokeWidth=1.5;fontFamily=Arial;fontColor={COLORS['title_color']};fontSize=12;fontStyle={font_style};" parent="{parent}" vertex="1"><mxGeometry x="{x}" y="{y}" width="{size}" height="{size}" as="geometry" /></mxCell>''')
        return cid

    def add_generic_icon(self, x, y, label, svg_data, size=28, parent="1", cid=None):
        """Generic icon (End User, Application Clients) with SVG."""
        cid = cid or self._next_id("gen")
        self.cells.append(f'''<mxCell id="{cid}" value="{self._esc_plain(label)}" style="shape=image;verticalLabelPosition=bottom;labelBackgroundColor=default;verticalAlign=top;aspect=fixed;imageAspect=0;image=data:image/svg+xml,{svg_data};strokeWidth=1.5;fontFamily=Arial;fontColor={COLORS['title_color']};fontSize=12;" parent="{parent}" vertex="1"><mxGeometry x="{x}" y="{y}" width="{size}" height="{size}" as="geometry" /></mxCell>''')
        return cid

    # === TEXT ===
    def add_text(self, x, y, w, h, text, color=None, size=12, bold=False, parent="1", align="center", cid=None):
        """Plain text label."""
        cid = cid or self._next_id("txt")
        color = color or COLORS['title_color']
        font_style = "1" if bold else "0"
        self.cells.append(f'''<mxCell id="{cid}" value="{self._esc_plain(text)}" style="text;html=1;align={align};verticalAlign=middle;resizable=0;points=[];autosize=1;strokeColor=none;fillColor=none;fontSize={size};fontColor={color};fontStyle={font_style};fontFamily=Arial;" parent="{parent}" vertex="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" /></mxCell>''')
        return cid

    def add_title(self, x, y, text, cid=None):
        """Diagram title (blue, bold, 16px)"""
        return self.add_text(x, y, len(text)*9, 30, text, color=COLORS['btp_border'], size=16, bold=True, cid=cid)

    # === PROTOCOL PILLS ===
    def add_pill(self, x, y, text, color_type="green", cid=None):
        """Protocol pill (SAML2/OIDC, HTTPS, SCIM)"""
        cid = cid or self._next_id("pill")
        border = COLORS[f'{color_type}_border']
        fill = COLORS[f'{color_type}_fill']
        self.cells.append(f'''<mxCell id="{cid}" value="{text}" style="rounded=1;whiteSpace=wrap;html=1;strokeColor={border};strokeWidth=1.5;arcSize=16;fillColor={fill};fontSize=10;fontStyle=1;fontColor={border};fontFamily=Arial;absoluteArcSize=1;" parent="1" vertex="1"><mxGeometry x="{x}" y="{y}" width="{max(len(text)*8, 60)}" height="16" as="geometry" /></mxCell>''')
        return cid

    # === CONNECTORS ===
    def add_connector(self, source, target, color_type="gray", label="", bidirectional=False, dashed=False, cid=None):
        """Connector with orthogonal routing.
        source/target should be icon IDs or area IDs returned by add_* methods."""
        cid = cid or self._next_id("edge")
        color_map = {
            "gray": COLORS['ext_border'], "blue": COLORS['btp_border'],
            "green": COLORS['green_border'], "indigo": COLORS['indigo_border'],
            "pink": COLORS['pink_border'], "red": COLORS['red_border'],
            "orange": COLORS['orange_border'], "teal": COLORS['teal_border'],
        }
        color = color_map.get(color_type, COLORS['ext_border'])
        start = "startArrow=blockThin;startFill=1;startSize=4;" if bidirectional else ""
        dash = "dashed=1;" if dashed else ""
        lbl_val = self._esc_plain(label) if label else ""
        self.cells.append(f'''<mxCell id="{cid}" value="{lbl_val}" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;endArrow=blockThin;endFill=1;endSize=4;startSize=4;strokeWidth=1.5;strokeColor={color};{start}{dash}fontFamily=Arial;fontSize=10;fontColor={COLORS['text_color']};" parent="1" source="{source}" target="{target}" edge="1"><mxGeometry relative="1" as="geometry" /></mxCell>''')
        return cid

    # === LEGEND ===
    def add_legend(self, x, y, entries=None):
        """Standard legend box."""
        if entries is None:
            entries = [
                ("Access", COLORS['ext_border']),
                ("Authentication", COLORS['green_border']),
                ("Deployment", COLORS['btp_border']),
            ]
        h = 30 + len(entries) * 20
        w = 200
        box = self._next_id("leg")
        self.cells.append(f'''<mxCell id="{box}" value="" style="rounded=0;whiteSpace=wrap;html=1;strokeColor={COLORS['legend_border']};strokeWidth=1.5;fillColor={COLORS['white']};absoluteArcSize=1;" parent="1" vertex="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" /></mxCell>''')
        self.add_text(x+10, y+5, 60, 20, "Legend", bold=True, size=12, align="left")
        for i, (name, color) in enumerate(entries):
            dot = self._next_id("dot")
            self.cells.append(f'''<mxCell id="{dot}" value="" style="ellipse;whiteSpace=wrap;html=1;aspect=fixed;strokeColor=none;fillColor={color};" parent="1" vertex="1"><mxGeometry x="{x+15}" y="{y+30+i*20}" width="10" height="10" as="geometry" /></mxCell>''')
            self.add_text(x+30, y+25+i*20, 150, 20, name, size=10, align="left")

    # === BUILD ===
    def build(self):
        cells_xml = "\n        ".join(self.cells)
        return f'''<mxfile host="Claude Code" agent="SAP Documentation Architect">
  <diagram name="{self.title}" id="{self.diagram_id}">
    <mxGraphModel dx="1400" dy="900" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="1169" pageHeight="827" math="0" shadow="0">
      <root>
        <mxCell id="0" />
        <mxCell id="1" parent="0" />
        {cells_xml}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>'''

    def save(self, path):
        with open(path, 'w') as f:
            f.write(self.build())
        print(f"  [OK] {path} ({os.path.getsize(path)//1024}KB)")

    # === HELPERS ===
    def _esc(self, text):
        """Escape for HTML-enabled values (used in areas with subtitles).
        HTML tags must be escaped as &lt; &gt; inside XML attributes.
        draw.io with html=1 will unescape them back to render as HTML."""
        return text.replace('&', '&amp;').replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')

    def _esc_plain(self, text):
        """Escape for plain text values (icons, simple labels).
        Preserves literal newlines (\n) which draw.io renders as line breaks.
        The \n must appear as actual newline character in the XML value attribute."""
        text = text.replace('&', '&amp;').replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')
        # Convert \n escape sequences to actual newline characters
        # draw.io reads real newlines in value= attributes as line breaks
        text = text.replace('\\n', '\n')
        return text


def get_icon_svg(key):
    return SAP_ICONS.get(key, {}).get("svg_base64", "")
