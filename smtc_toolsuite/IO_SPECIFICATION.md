# SMTC Toolsuite — folder and file I/O map

Scope: current code inside `smtc_toolsuite/` only. “Effectively mandatory” means the code reads the path without a complete pre-check, so the process cannot finish without it.

## Shared folder convention

The four applications persist user-selected absolute paths in their PKL state. Their main folder roles are:

| Folder role | Used by | Contents |
|---|---|---|
| Working/output folder | All four applications | Generated reports and, for three tools, required `columns and formatting.xlsx`. |
| Master reports folder | Manufacturing plan | Source `.xlsx`/`.xlsm` master files normalized through mappings in `columns and formatting.xlsx`. |
| Outlook mail folder | Local shipments | Source messages containing Korrus `.xlsx`/`.csv` attachments; Windows/Outlook only. |
| Yield reports folder | Local shipments | All entries are read as yield-report Excel files by the AAR process. |
| Price-list folder | Local shipments | Files whose names match fixture or accessory master-price naming rules. |

## `carga_demanda_l.py`

Brief flow: Korrus PO demand → BOM explosion → Apollo adjustment → write-forecast adjustment.

### Input folders and files

| Input | Status | Expected content |
|---|---|---|
| Working/output folder | Mandatory | Destination for every generated workbook. |
| Korrus file | Mandatory and explicitly checked | `.xlsx`, UTF-16 tab CSV, or comma CSV. Required columns include `PurchaseOrder`, `PODate`, `Quantity`, `ProductServiceID`. |
| BOM | Mandatory/effective | Excel; columns `BOM`, `Component`, `Component Description`, `Qty Per`. |
| Apollo | Mandatory/effective | Excel workbook; sheets containing `Forecast`; requires `Part Number`, `Description`, and dated forecast columns. |
| Write Forecast | Mandatory/effective | Excel with the complete `Write Forecast` column set defined in `mandatory_cols`. |
| Shipment | Optional | Excel; if selected, requires `Part No.`, `Customer PO#`, `Qty. Shipped` and is also filtered using `Site`. |
| `columns and formatting.xlsx` | Defined but not consumed by the main flow | Path is created by `set_paths()`, but the demand-generation function does not read it. |

### Outputs in the working folder

- `Pos_bom_raw.xlsx`: accumulated exploded PO/BOM demand; same extraction date is replaced.
- `Forecast de Apollo actualizado.xlsx`: revised Apollo forecast with formula columns.
- `Reportes.xlsx`: conditional warning sheets for missing BOMs and components absent from forecast.
- `KorrusFile_Pendientes_<YYYY-MM-DD>.xlsx`: Korrus rows for models lacking BOMs.
- `Write_forecast_updated.xlsx`: revised write forecast.
- `Write_forecast_changes.xlsx`: rows added or changed by PO demand and shipment deductions.

## `clear_to_build.py`

Brief flow: demand + launched work orders + BOM + consumption/inventory/receipts → CTB → launch suggestion → gating report.

### Input folders and mandatory files

| Input | Status in current code | Expected content |
|---|---|---|
| Working/output folder | Mandatory | Destination for intermediate and final workbooks. |
| `columns and formatting.xlsx` in working folder | Effectively mandatory at page initialization | Sheets `column_format`, `column_equivalence`, `master_operation_relation`; also supplies workbook formatting. |
| Korrus | Explicitly mandatory for analysis/suggestion | Sheet `Data`; Korrus columns defined in `mandatory_cols`. Used when Independent Demands is selected. |
| BOM | Explicitly mandatory | Sheet `Flat Bill Browser - Cost Roll U`; BOM columns defined in `mandatory_cols`. |
| WOS | Explicitly mandatory | One or more sheets containing `Seguimiento`; WOS columns defined in `mandatory_cols`. |
| Independent Demands **or** Tablillas | Exactly one is intended | Independent Demands uses sheet `Independent Demands`. Tablillas flow reads `Work Order Action Report` and later also reads its first/default sheet. Current validation reports bad selection but does not immediately stop. |
| Consumption | Effectively mandatory | Excel read with header row 2; used as the base workbook for CTB. |
| On Hand Detail | Effectively mandatory | Excel with `Part`, `Quantity`. |
| Pendiente | Effectively mandatory | Pending-receipt Excel. Current column-validation call passes the DataFrame rather than `df.columns`. |
| Calendario | Effectively mandatory | Excel with at least `monday_date`, `year`, `closing_month`. |

### Optional selected files

- `Alternos`: maps alternate components and on-hand quantities.
- `BOM Detail`: adds `Primary Stock`.
- `Component Allocation`: used only with the Tablillas/additional-components flow.
- `PO WO Info`: optional arrival dates/quantities for the suggestion and gating calculations.
- `Work Order Action`: exposed in the UI but not read under this selector key in the current module.

### Outputs in the working folder

- `Analisis_lanzamiento.xlsx` or `Analisis_lanzamiento_tablillas.xlsx`: demand/launch analysis; later overwritten with suggestion, shortage detail, and shortage summary sheets.
- `bom_demand.xlsx`: exploded demand/BOM intermediate.
- `rl_raw.xlsx`: normalized consumption demand and scheduled-receipt intermediate.
- `CTB KRS.xlsx` or `CTB KRS_tablillas.xlsx`: main clear-to-build workbook.
- `Gating Parts Report.xlsx`: sheets `Shorts`, `Ready`, and `Arrivals`.

The paths `status de ordenes.xlsx` and `reporte de manufactura.xlsx` are defined in `set_paths()` but are not used by the current CTB flow.

## `manufacturing_plan.py`

Brief flow: prioritize orders → verify routing → normalize master files and select machines → schedule capacity → report and validate.

### Input folders and files

| Input | Status | Expected use |
|---|---|---|
| Working/output folder | Mandatory | Contains configuration and generated plans/reports. |
| `columns and formatting.xlsx` in working folder | Mandatory | Sheets `column_format`, `column_equivalence`, `master_operation_relation`, plus formatting definitions used by the planner. |
| Master reports folder | Mandatory for machine selection | All top-level `.xlsx`/`.xlsm` files are scanned by `ExcelNormalizer`; matching files/sheets come from `column_equivalence`. |
| Lista de Ordenes | Explicitly mandatory for Create Plan | Order list; header is located using `Priority`. |
| Routing | Explicitly mandatory for Create Plan | Sheet `Operations`; header is located using `Routing`. |
| Equivalencias | Effectively mandatory for Verify Orders | Read using the `Equivalencias` table definition. |
| Horas disponibles | Effectively mandatory for machine selection/plan | Provides `dia`, `maquina`, and shift-capacity fields. |
| Part Master | Effectively mandatory for plan creation | Header located using `Site`; supplies unit selling price. |
| 16Wk Gap | Mandatory only for “Sugerir prioridades” | Renamed through the `16 WK` mapping. |
| Top ventas | Mandatory only for “Sugerir prioridades” | Sales ranking input. |
| End Of Period | Mandatory only for “Sugerir prioridades” | Header located using `Report Date`; source work-order demand. |
| Existing `<plan_name>.xlsx` | Optional existing state | Read and merged when present; otherwise created as a new plan. |

### Outputs in the working folder

- `Lista de Ordenes sugerida.xlsx`: output of priority suggestion.
- `<plan_name>.xlsx`: main manufacturing plan; updated by create/validate.
- `<plan_name>_assignment.xlsx`: validation result showing assigned/pending order status.
- `reporte de manufactura <machine> <YYYY-MM-DD_HH-MM-SS>.xlsx`: one timestamped workbook per machine.

`status de ordenes.xlsx` is defined in `set_paths()` but is not used by the current production planner.

## `local_shipments_l.py`

Brief flow: Outlook Korrus attachments → EDI/shipment consolidation → OOR rebuild → yield, gating, and status updates.

### Input folders

| Folder | Status | Use |
|---|---|---|
| Working/output folder | Mandatory | Contains configuration, downloaded attachment staging, consolidated files, and some reports. |
| Outlook folder | Mandatory only for “Explorar Outlook” | Messages since `fecha_mail`; only attachments whose filename contains `KORRUS` are saved. |
| Yield reports folder | Mandatory only for AAR | Every directory entry is passed to the Excel reader and mapped into daily-status sheets. |
| Price-list folder | Mandatory only for “Cargar precios” | Reads matching fixture/accessory master-price files into session memory. |

### Mandatory/effective files

| File | Status | Use |
|---|---|---|
| `columns and formatting.xlsx` in working folder | Mandatory | Column mappings, formats, locations, options, conditions, and status rules. It is modified in place by “Actualizar Status”. |
| OOR | Effective core input | Selected workbook; rebuilt and modified in place by OOR, AAR, gating, and status processes. |
| Tracker | Effective for “Actualizar OOR” | Reads sheets containing `Plan de produccion` or `TERMINADAS`. |
| EDI Master | Effective core input | Workbook containing at least sheets `EDI Master`, `Shipment to ELP`, and `Cancelled Orders`; modified in place by “Actualizar EDI Master”. |
| `Shipped to Cust.xlsx` in working folder | Effective for “Actualizar OOR” | Must already exist or be generated from ELP Master log during the EDI update. |

The code declares `Tracker`, `EDI Master`, and `OOR` as `mandatory_selectors`, but the current Streamlit process buttons do not call `verify_selections()`; missing paths generally fail when used.

### Optional/process-specific files

- `Top Priority`: optional `Changes Request` sheet merged into EDI.
- `ELP Master log`: optional source for `Shipped to Cust.xlsx`.
- `InventoryStageBakup`: optional source appended to `Shipment to ELP`.
- `OH Max`: optional inventory integration during OOR update; expected to have exactly six columns.
- `Gating Parts`: mandatory only for the Gating Parts action; expected sheets `Ready`, `Shorts`, `Arrivals`.
- Working-folder `Korrus_data.xlsx`: optional before first run; created when absent and then merged into EDI.

### Outputs and in-place updates

In the selected working folder:

- `attachments/`: newly downloaded Korrus attachments.
- `attachments/Done/`: processed attachments moved from the staging directory.
- `Korrus_list.xlsx`: attachment processing ledger.
- `Korrus_data.xlsx`: consolidated Korrus data.
- `Shipped to Cust.xlsx`: normalized customer shipment history.
- `oor_options.xlsx`: grouped fields used during status calculation.

Modified in place:

- Selected `EDI Master` workbook.
- Selected `OOR` workbook.
- Working-folder `columns and formatting.xlsx` (`options`, defined names, and condition validation).

Troubleshooting outputs written to the **process current working directory**, not necessarily the selected working folder:

- `modified_rows.xlsx`
- `new_rows.xlsx`
- `df_oor_to_update.xlsx`
- `df_oor_old.xlsx`
- `df_oor.xlsx`
- `status.xlsx`

## Helper and development module I/O

### `excel_normalizer.py`

- Mandatory inputs: mapping CSV with `file_name_like`, `sheet_name_like`, `std_name`, `other_name`; input folder containing top-level `.xlsx`/`.xlsm`.
- Optional input: explicit sheet-name list.
- Intended output: caller-selected CSV path.
- Current behavior: `normalize_folder()` returns `(DataFrame, errors)`, while `main()` treats it as a DataFrame, so the CLI currently fails at `.to_csv()`.

### `dev_manuf_plan.py`

- Reuses `folder_state_planner.pkl`, planner mappings, selected planner inputs, and the planner working/master folders.
- Contains older keys/paths not selected by the current production UI, including `master_file`.
- Writes or updates `manufacturing plan.xlsx`, `reporte de manufactura...xlsx`, the selected order file, and `Lista de Ordenes.xlsx` in the configured folder.
- It is a development script, not registered in `Home.py`.

### `dev_shipments.py`

- Reuses shipment state and production helper functions, but also contains hard-coded exploratory paths.
- Reads and modifies the same EDI/OOR/report inputs as `local_shipments_l.py`.
- Generates troubleshooting workbooks in the process current working directory.
- It is a development script, not registered in `Home.py`.

### Launchers and non-data pages

- `Home.py` reads no business files and delegates to registered pages.
- `Indice/scripts_catalog.py` runs `git pull` in `smtc_toolsuite/`; its output is captured and displayed, not written as a report.
- `Documentacion/plan_de_manufactura_doc.py` reads `Documentacion\manufacturing_plan.html` outside this folder.
- `tst.py` has no file I/O.
- The three `.bat` files launch Streamlit and may install Python packages.

