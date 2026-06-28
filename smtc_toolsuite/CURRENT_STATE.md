# SMTC Toolsuite — current repository state

Scope: `smtc_toolsuite/` only. Snapshot reviewed on 2026-06-28. This is a descriptive inventory of the current code, not a target design.

## Runtime structure

`Home.py` is the Streamlit entry point. It registers four operational pages, one repository-index page, and one documentation page:

- `Herramientas/manufacturing_plan.py`: builds and validates a manufacturing plan, assigns work to machines and shifts, and creates machine reports.
- `Herramientas/local_shipments_l.py`: consolidates Korrus/EDI/shipment/tracker information and updates the selected OOR workbook.
- `Herramientas/carga_demanda_l.py`: converts customer POs through the BOM, updates Apollo demand, and updates the write forecast.
- `Herramientas/clear_to_build.py`: builds a clear-to-build workbook, suggests launches, and creates a gating-parts report.
- `Indice/scripts_catalog.py`: Streamlit button that runs `git pull` for the toolsuite directory.
- `Documentacion/plan_de_manufactura_doc.py`: embeds an external manufacturing-plan HTML document.

## Very brief module descriptions

| Module | Current role |
|---|---|
| `Home.py` | Streamlit multipage navigation and page registration. |
| `utils.py` | Shared file dialogs, Excel formatting, state-pickle, and DataFrame utilities. The operational pages currently duplicate many of these functions. |
| `Herramientas/carga_demanda_l.py` | Demand-load application: explodes Korrus demand through BOM, adjusts Apollo, and produces a revised write forecast. |
| `Herramientas/clear_to_build.py` | CTB application: combines demand, BOM, consumption, inventory, receipts, and optional alternates; then suggests launches. |
| `Herramientas/manufacturing_plan.py` | Planner application: prioritizes orders, maps routing/master machines, schedules capacity, validates assignments, and creates machine reports. |
| `Herramientas/local_shipments_l.py` | Shipment/OOR application: imports Outlook Korrus attachments, updates EDI and shipment history, rebuilds OOR data, and applies gating/status data. |
| `Herramientas/excel_normalizer.py` | Reusable mapping-driven normalizer for tables found across Excel files and sheets. Also contains a CLI entry point. |
| `Herramientas/dev_manuf_plan.py` | Development/notebook-style manufacturing-plan script that imports the production planner and reuses planner state. |
| `Herramientas/dev_shipments.py` | Development/notebook-style shipment script that imports the production shipment module and contains exploratory/debug flows. |
| `Herramientas/tst.py` | Minimal unused Streamlit test page. |
| `Indice/scripts_catalog.py` | Repository update page using a fixed `mingit` executable path. |
| `Documentacion/plan_de_manufactura_doc.py` | Loads and displays `Documentacion\manufacturing_plan.html` from outside this folder. |
| `__init__.py`, `Herramientas/__init__.py` | Empty package markers. |
| `SMTC Toolsuite.bat` | Starts the multipage application with `streamlit run Home.py`. |
| `Herramientas/manufacturing_plan.bat` | Installs missing dependencies and launches the planner directly. |
| `Herramientas/local_shipments.bat` | Installs missing dependencies and launches shipment tracking directly. |

## PKL configuration coverage

The PKLs contain local absolute paths and user selections; they are runtime state, not portable application configuration.

| Stateful application/module | PKL referenced by code | Present now | Result |
|---|---|---:|---|
| `manufacturing_plan.py` | `folder_state_planner.pkl` | Yes | Own application state exists. |
| `clear_to_build.py` | `folder_state_clear_to_build.pkl` | Yes | Own application state exists. |
| `local_shipments_l.py` | `folder_state_local_shipments.pkl` | Yes | Own application state exists. |
| `carga_demanda_l.py` | `state_carga_demanda.pkl` | **No** | Missing. Code creates default in memory and writes it after the first selection. |
| `dev_manuf_plan.py` | `folder_state_planner.pkl` | Yes | Shares planner state; no separate PKL. The relative path depends on the process working directory. |
| `dev_shipments.py` | default `folder_state_local_shipments.pkl` | Ambiguous | Shares shipment state by name, but the relative lookup depends on the process working directory. |
| Other Python modules | None | N/A | They do not maintain their own persisted UI state. |

Therefore, each operational tool does **not** currently have a present PKL: demand load is missing `state_carga_demanda.pkl`. Each Python module also does not have its own PKL by design; development modules share operational state and stateless modules need none.

### PKL schemas and current selection coverage

- `folder_state_planner.pkl`
  - Top-level keys now: `folder_output`, `selections`, `plan_name`, `folder_master`.
  - Selection keys now: `16_wk`, `sales`, `end_of_period`, `equivalencias_file`, `order_file`, `routing_file`, `available_hours`, `part_master`.
  - All are configured except `sales`, which is an empty string.
  - Dates used later by the application (`initial_date`, `limit_date`) are not present in the current PKL.
- `folder_state_clear_to_build.pkl`
  - Top-level keys now: `folder_output`, `selections`, `plan_name`.
  - Selection keys now: `consumption`, `independent_demands`, `korrus`, `on_hand_detail`, `pendiente`, `wos`, `bom`, `bom_detail`, `calendario`.
  - UI selectors not currently recorded: `alternos`, `po_wo_info`, `work_order_action`, `tablillas`, `component_allocation`.
- `folder_state_local_shipments.pkl`
  - Top-level keys now: `folder_output`, `selections`, `fecha_mail`, `fecha_shipments_elp`, `fecha_freeze`, `outlook_folder`.
  - Selection keys now: `OOR`, `Tracker`, `EDI Master`, `ELP Master log`, `InventoryStageBakup`, `Top Priority`.
  - `fecha_shipments_elp` and `outlook_folder` are currently `None`.
  - UI selectors/folders not currently recorded: `OH Max`, `Gating Parts`, `folder_yield`, `folder_prices`, and the three ETA-gap values.
- `state_carga_demanda.pkl`
  - Expected default schema from code: `folder_output`, `selections`; later it may also receive `korrus_date`.
  - Expected selection keys: `korrus_file`, `bom_file`, `write_forecast`, `apollo_file`, `shipment_file`.

## Current-state observations

- Several applications treat the selected “output” folder as a working folder that must already contain `columns and formatting.xlsx`.
- Some “mandatory” inputs are validated explicitly; others are accessed directly and are only effectively mandatory. The detailed I/O map keeps these cases separate.
- `local_shipments_l.py` and the development scripts create troubleshooting workbooks in the process current working directory, outside the selected working folder.
- `local_shipments_l.py` modifies selected `EDI Master`, `OOR`, and `columns and formatting.xlsx` files in place.
- `excel_normalizer.py` currently returns `(dataframe, errors)` from `normalize_folder()`, but its CLI calls `.to_csv()` on that tuple. The intended CLI output is clear, but the current CLI path will fail before writing it.
- `Documentacion/plan_de_manufactura_doc.py` points to an HTML file outside `smtc_toolsuite/`, and uses a backslash inside the filename component. That dependency is not present inside the reviewed scope.

