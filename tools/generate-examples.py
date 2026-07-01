#!/usr/bin/env python3
"""Generate 5 SAP BTP Solution Diagram examples using the generator."""

import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from importlib import import_module
gen = import_module("sap-drawio-generator")

SAPDiagramBuilder = gen.SAPDiagramBuilder
COLORS = gen.COLORS

outdir = os.path.join(os.path.dirname(__file__), "..", "examples")
os.makedirs(outdir, exist_ok=True)

print("Generating 5 SAP BTP Solution Diagram examples...\n")

# ============================================================
# EXAMPLE 1: CAP + HANA + Fiori Elements (Simple L1)
# ============================================================
d1 = SAPDiagramBuilder("CAP_Fiori_HANA_L1", "ex1")
d1.add_title(20, 10, "CAP + Fiori Elements + HANA Cloud — L1")
d1.add_btp_area(50, 50, 750, 480)
d1.add_subaccount(70, 100, 710, 410)

# Services row
d1.add_icon(100, 155, "hana-cloud", "SAP HANA\nCloud", 32)
d1.add_icon(210, 155, "authorization-trust", "XSUAA", 32)
d1.add_icon(320, 155, "destination-service", "Destination\nService", 32)
d1.add_icon(430, 155, "html5-app-repo", "HTML5 App\nRepository", 32)

# CAP Application area
cap_area = d1.add_service_area(100, 250, 300, 140, "CAP Application")
cap_icon = d1.add_icon(130, 300, "cap-model", "CAP Service\nNode.js", 28)
d1.add_text(270, 310, 110, 30, "OData V4\nEndpoints", size=10, color=COLORS['text_color'])

# Fiori App area
fiori_area = d1.add_service_area(440, 250, 300, 140, "Fiori Elements App")
fiori_icon = d1.add_icon(470, 300, "build-work-zone", "Fiori\nList Report", 28)

# External
user = d1.add_text(870, 220, 80, 30, "End User", bold=True)
s4 = d1.add_external_area(850, 320, 180, 100, "SAP S/4HANA", "On-Premise")

# Connectors — point to ICONS, not areas
d1.add_connector(fiori_icon, cap_icon, "blue", "OData V4")
d1.add_connector(cap_icon, s4, "gray", "RFC / OData")

d1.add_legend(560, 550, [("Access", COLORS['ext_border']), ("BTP Internal", COLORS['btp_border'])])
d1.save(f"{outdir}/01-cap-fiori-hana-L1.drawio")

# ============================================================
# EXAMPLE 2: Integration Suite + CPI (L1)
# ============================================================
d2 = SAPDiagramBuilder("Integration_Suite_L1", "ex2")
d2.add_title(20, 10, "SAP Integration Suite — CPI Flows — L1")
d2.add_btp_area(50, 50, 800, 420)
d2.add_subaccount(70, 100, 760, 350)

# Integration Suite area
d2.add_service_area(100, 160, 400, 220, "SAP Integration Suite")
cpi = d2.add_icon(130, 215, "cloud-connector", "Cloud\nIntegration", 28)
mesh = d2.add_icon(260, 215, "event-mesh", "Event\nMesh", 28)
apim = d2.add_icon(380, 215, "destination-service", "API\nManagement", 28)
d2.add_text(130, 310, 340, 25, "iFlows: Invoice Sync, Master Data Replication", size=10, color=COLORS['text_color'])

# Monitoring
alm = d2.add_icon(550, 215, "analytics-cloud", "SAP Cloud\nALM", 28)

# External systems
s4 = d2.add_external_area(930, 120, 180, 80, "SAP S/4HANA")
api = d2.add_external_area(930, 270, 180, 80, "3rd Party API")

# Connectors to ICONS
d2.add_connector(cpi, s4, "gray", "OData")
d2.add_connector(apim, api, "green", "REST/HTTPS")

d2.add_legend(620, 490, [("Data Flow", COLORS['ext_border']), ("Secure Channel", COLORS['green_border'])])
d2.save(f"{outdir}/02-integration-suite-L1.drawio")

# ============================================================
# EXAMPLE 3: Multi-Tenant SaaS on BTP (L1)
# ============================================================
d3 = SAPDiagramBuilder("MultiTenant_SaaS_L1", "ex3")
d3.add_title(20, 10, "Multi-Tenant SaaS Application — BTP — L1")
d3.add_btp_area(50, 50, 900, 500)

# Provider subaccount
d3.add_subaccount(70, 100, 400, 420, "Provider Subaccount")
cap_saas = d3.add_icon(110, 170, "cap-model", "CAP SaaS\nApplication", 28)
hana = d3.add_icon(240, 170, "hana-cloud", "HANA Cloud\n(HDI Shared)", 28)
xsuaa = d3.add_icon(370, 170, "authorization-trust", "XSUAA\n(Multi-tenant)", 28)
html5 = d3.add_icon(110, 290, "html5-app-repo", "HTML5 App\nRepository", 28)
jobs = d3.add_icon(240, 290, "job-scheduling", "Job\nScheduling", 28)

# Subscriber A
d3.add_subaccount(510, 100, 400, 160, "Subscriber Tenant A")
wz_a = d3.add_icon(540, 170, "build-work-zone", "Work Zone", 28)
d3.add_pill(660, 185, "Subscribed", "green")

# Subscriber B
d3.add_subaccount(510, 290, 400, 160, "Subscriber Tenant B")
wz_b = d3.add_icon(540, 360, "build-work-zone", "Work Zone", 28)
d3.add_pill(660, 375, "Subscribed", "green")

# Connectors
d3.add_connector(cap_saas, wz_a, "blue", "Subscription")
d3.add_connector(cap_saas, wz_b, "blue", "Subscription")
d3.add_connector(cap_saas, hana, "gray")

d3.add_legend(700, 570, [("Subscription", COLORS['green_border']), ("BTP Service", COLORS['btp_border']), ("Data Flow", COLORS['ext_border'])])
d3.save(f"{outdir}/03-multitenant-saas-L1.drawio")

# ============================================================
# EXAMPLE 4: S/4HANA Extension Side-by-Side (L2)
# ============================================================
d4 = SAPDiagramBuilder("S4_Extension_L2", "ex4")
d4.add_title(20, 10, "S/4HANA Side-by-Side Extension — BTP — L2")
d4.add_btp_area(50, 50, 760, 570)
d4.add_subaccount(70, 100, 720, 500)

# Services row
xsuaa4 = d4.add_icon(100, 155, "authorization-trust", "XSUAA", 28)
dest4 = d4.add_icon(210, 155, "destination-service", "Destination", 28)
conn4 = d4.add_icon(320, 155, "cloud-connector", "Connectivity", 28)
html5_4 = d4.add_icon(430, 155, "html5-app-repo", "HTML5 Repo", 28)

# CAP Extension area
d4.add_service_area(100, 260, 380, 140, "Extension Application")
cap4 = d4.add_icon(130, 310, "cap-model", "CAP Extension\nService", 28)
hana4 = d4.add_icon(280, 310, "hana-cloud", "HANA Cloud", 28)

# Fiori Launchpad area
d4.add_service_area(100, 430, 380, 110, "Fiori Launchpad")
wz4 = d4.add_icon(130, 460, "build-work-zone", "SAP Build\nWork Zone", 28)
ias4 = d4.add_icon(280, 460, "identity-authentication", "Identity\nAuthentication", 28)

# Protocol pills
d4.add_pill(520, 470, "SAML2/OIDC", "green")

# NETWORK column
d4.add_network_column(850, 80, 110, 540)
cc4 = d4.add_icon(865, 250, "cloud-connector", "Cloud\nConnector", 28)
d4.add_pill(860, 400, "HTTPS", "ext")

# S/4HANA
s4_4 = d4.add_external_area(1000, 180, 190, 160, "SAP S/4HANA", "On-Premise")
d4.add_text(1020, 270, 150, 30, "Custom CDS Views\n+ OData APIs", size=10, color=COLORS['text_color'])

# Connectors
d4.add_connector(cap4, cc4, "gray", "RFC")
d4.add_connector(cc4, s4_4, "green")
d4.add_connector(wz4, cap4, "blue")
d4.add_connector(ias4, xsuaa4, "green", "OIDC", dashed=True)

d4.add_legend(520, 640, [
    ("Access", COLORS['ext_border']),
    ("Authentication", COLORS['green_border']),
    ("BTP Internal", COLORS['btp_border']),
])
d4.save(f"{outdir}/04-s4-extension-L2.drawio")

# ============================================================
# EXAMPLE 5: Document AI + OCR Processing (L2)
# ============================================================
d5 = SAPDiagramBuilder("Document_AI_OCR_L2", "ex5")
d5.add_title(20, 10, "Document AI Processing Pipeline — BTP — L2")
d5.add_btp_area(50, 50, 830, 560)
d5.add_subaccount(70, 100, 790, 490)

# AI subaccount area
d5.add_service_area(100, 155, 220, 130, "AI Services")
dox = d5.add_icon(160, 200, "analytics-cloud", "SAP Document\nAI (DOX)", 28)

# Main app area
d5.add_service_area(350, 155, 490, 270, "Invoice Processing")
core = d5.add_icon(380, 210, "cap-model", "CORE\nService", 28)
ctrl = d5.add_icon(490, 210, "cap-model", "Country\nController", 28)
ocr = d5.add_icon(600, 210, "build-work-zone", "OCR Upload\nWizard", 28)
mon = d5.add_icon(710, 210, "build-work-zone", "Invoice\nMonitor", 28)
hana5 = d5.add_icon(380, 330, "hana-cloud", "HANA Cloud", 28)
sched5 = d5.add_icon(490, 330, "job-scheduling", "Job\nScheduler", 28)
xsuaa5 = d5.add_icon(600, 330, "authorization-trust", "XSUAA", 28)

# Identity Services area
d5.add_service_area(100, 450, 400, 90, "SAP Cloud Identity Services")
idauth = d5.add_icon(140, 470, "identity-authentication", "Identity\nAuth", 24)
idprov = d5.add_icon(280, 470, "identity-authentication", "Identity\nProvisioning", 24)

# Protocol pills
d5.add_pill(530, 470, "SAML2/OIDC", "green")
d5.add_pill(650, 470, "SCIM", "indigo")

# Network + External
d5.add_network_column(920, 80, 110, 520)
cc5 = d5.add_icon(940, 250, "cloud-connector", "Cloud\nConnector", 28)
d5.add_pill(935, 400, "HTTPS", "ext")

s4_5 = d5.add_external_area(1070, 180, 180, 200, "SAP S/4HANA", "On-Premise")

# Connectors
d5.add_connector(dox, core, "blue", "DOX Extract")
d5.add_connector(core, ctrl, "gray", "Rules")
d5.add_connector(core, hana5, "gray")
d5.add_connector(core, cc5, "gray", "OData APIs")
d5.add_connector(cc5, s4_5, "green")
d5.add_connector(idauth, xsuaa5, "green", "OIDC", dashed=True)

d5.add_legend(530, 630, [
    ("Access", COLORS['ext_border']),
    ("Authentication", COLORS['green_border']),
    ("Authorization", COLORS['indigo_border']),
    ("BTP Internal", COLORS['btp_border']),
])
d5.save(f"{outdir}/05-document-ai-ocr-L2.drawio")

print("\nDone! 5 diagrams generated.")
