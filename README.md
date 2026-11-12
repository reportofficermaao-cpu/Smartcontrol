Excel VBA Smart Control Center (SCC)
An enterprise-grade, modular productivity framework embedded directly into Microsoft Excel. It centralizes data hygiene, workspace navigation, automated MIS dashboards, multi-format reporting, and AI-assisted workflows into a single interactive dashboard.

Core Capabilities
Workspace Navigation: Dynamic hyperlinked Workbook Index, Fast Navigator, and bulk Sheet Manager with deletion safeguards.

In-Memory Data Hygiene: High-speed array processing for trimming spaces, stripping non-printable characters, standardizing dates/numbers, and removing duplicates—complete with real-time progress tracking and automatic sheet backups.

MIS Dashboards & Visualization: One-click generation of executive KPI cards and auto-dimensioned charts from any active dataset.

Multi-Format Reporting & Automation: Native PDF publishing, late-bound Word (.docx) and PowerPoint (.pptx) exports, and automated Outlook email drafting.

Formula Lab & AI Bridge: Searchable formula repository alongside an optional REST API bridge for natural-language formula generation (falls back gracefully to offline heuristics when no API key is configured).

Quick Setup
Open your workbook in desktop Microsoft Excel and save it as a Macro-Enabled Workbook (.xlsm).

Press Alt + F11 to open the VBA Editor and import the modSCC_* standard modules, frmSCC_* UserForms, and ThisWorkbook lifecycle events.

Return to Excel, press Alt + F8, and run BuildSmartControlCenter to generate the home interface and hidden configuration/audit sheets (SCC_Config and SCC_Log).
