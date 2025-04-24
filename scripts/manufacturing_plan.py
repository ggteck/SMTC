# Manufacturing plan
# - V4. 2025-04-21
#     - Correccion a extraccion de routers, se agregan variables routers, turnos
# - V3. 2025-04-21
#     - Reportes en excel
# - V2. 2025-04-19
#     - Verificacion de ordenes ya programadas, manejo de status, integracion con plan existente
# - V1. 2025-04-10
#     - Version inicial, calculo de plan de produccion

import streamlit as st
import pickle
import os
from tkinter import Tk, filedialog as fd
import datetime
import pandas as pd
from copy import copy
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter

# =============================================================================
# File/Directory & System Utilities
# =============================================================================

def open_file_selection(initialdir='', filter_name='*.*'):
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    filetypes = (
        (filter_name, f"*{filter_name.split(' ')[0]}*.*"),
        ('All files', '*.*'),
        ('Excel', '*.xlsx'),
        ('CSV', '*.doc')
    )
    files = fd.askopenfilenames(filetypes=filetypes, initialdir=initialdir)
    root.destroy()
    return files

def select_directory(initialdir):
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    directory = fd.askdirectory(initialdir=initialdir)
    root.destroy()
    return directory

def set_paths(path):
    output_paths = {}
    output_paths['path_xl_format'] = os.path.join(path, 'columns and formatting.xlsx')
    output_paths['path_plan'] = os.path.join(path, 'manufacturing plan.xlsx')
    output_paths['path_report'] = os.path.join(path, 'reporte de manufactura.xlsx')
    return output_paths

def set_col_rel(output_paths):
    df_columns = read_excel(output_paths['path_xl_format'], sheet_name='column_format')
    df_col_rel = df_columns[~df_columns['std_name'].isnull()].copy()
    return {'col_rel': df_col_rel, 'columns': df_columns}

# =============================================================================
# Excel Management Functions
# =============================================================================

def read_excel(path=None, sheet_name=0, header=0, keep_default_na=True, dtype=None):
    with pd.ExcelFile(path) as xls:
        df = pd.read_excel(path, sheet_name=sheet_name, header=header, keep_default_na=keep_default_na, dtype=dtype)
    return df

def load_excel_with_header_key(file_path, sheet_name=0, key_text='', dtype=None, **kwargs):
    df = read_excel(file_path, sheet_name=sheet_name, keep_default_na=False, dtype=dtype)
    header_row = None
    if key_text in df.columns:
        header_row = 0
    else:
        for col in df.columns:
            for i, cell in df[col].items():
                if pd.notna(cell) and key_text in str(cell):
                    header_row = i
                    break
            if header_row is not None:
                break
        if header_row is None:
            raise ValueError(f"Key text '{key_text}' not found in the sheet.")
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)
    return df

# =============================================================================
# Persistence / State Management
# =============================================================================

def save_state_pickle(state, filename='folder_state_planner.pkl'):
    with open(filename, 'wb') as f:
        pickle.dump(state, f)

def load_state_pickle(filename='folder_state_planner.pkl'):
    try:
        with open(filename, 'rb') as f:
            return pickle.load(f)
    except FileNotFoundError:
        return {"folder_output": None, "selections": {}, "initial_date": datetime.date.today()}

# =============================================================================
# DataFrame Management & Data Processing
# =============================================================================

def rename_columns(df, df_col_rel, table_from='std_name', sheet_from='only', table_to='std_name', sheet_to='only'):
    required_cols = {'table', 'sheet', 'column_name', 'std_name'}
    missing_cols = required_cols - set(df_col_rel.columns)
    if missing_cols:
        raise ValueError(f"df_col_rel is missing required columns: {missing_cols}")
    mapping = {}
    if table_to == 'std_name':
        filtered = df_col_rel[(df_col_rel['table'] == table_from) & (df_col_rel['sheet'] == sheet_from)]
        if filtered.empty:
            raise ValueError("No matching mapping found for given table_from and sheet_from.")
        mapping = filtered.set_index('column_name')['std_name'].to_dict()
    elif table_from == 'std_name':
        filtered = df_col_rel[(df_col_rel['table'] == table_to) & (df_col_rel['sheet'] == sheet_to)]
        if filtered.empty:
            raise ValueError("No matching mapping found for given table_to and sheet_to.")
        mapping = filtered.set_index('std_name')['column_name'].to_dict()
    else:
        filtered_from = df_col_rel[(df_col_rel['table'] == table_from) & (df_col_rel['sheet'] == sheet_from)]
        if filtered_from.empty:
            raise ValueError("No matching mapping found for given table_from and sheet_from.")
        mapping_from = filtered_from.set_index('column_name')['std_name'].to_dict()
        df = df.rename(columns=mapping_from)
        filtered_to = df_col_rel[(df_col_rel['table'] == table_to) & (df_col_rel['sheet'] == sheet_to)]
        if filtered_to.empty:
            raise ValueError("No matching mapping found for given table_to and sheet_to.")
        mapping = filtered_to.set_index('std_name')['column_name'].to_dict()
    return df.rename(columns=mapping)

def format_dates(df, date_cols=[],type='iso'):
    for col in date_cols:
        if not col in df.columns:
            continue
        if type=='iso':
            df[col]=pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')
        else:
            df[col]=pd.to_datetime(df[col], errors='coerce')
    return df

def read_predefined_excel(path,df_columns=pd.DataFrame(),table='',sheet='only',check_mandatory=False):
    """
    If the file exists, it reads it and returns the dataframe, 
    otherwise it returns an empty dataframe but with the columns of the file
    it checks mandatory columns according to columns and formatting file but rises it only if indicated
    """
    if not os.path.exists(path):
        df=get_predefined_df(df_columns=df_columns,
                             table=table,
                             sheet=sheet)
    else:
        df=read_excel(path)    
        if check_mandatory:
            check_mandatory_columns_df(df.columns,df_columns=df_columns,table=table,sheet=sheet)
    return df

def get_predefined_df(df_columns=pd.DataFrame(),table='',sheet='only'):
    cols=df_columns[(df_columns['table']==table)&
                                        (df_columns['sheet']==sheet)]['column_name'].to_list()
    if len(cols)==0:
        st.error(f"Tabla: {table} no definida en archivo columns and formatting")
        st.stop()
    df=pd.DataFrame(columns=cols)
    return df

def check_mandatory_columns_df(cols=[],df_columns=pd.DataFrame(),table='',sheet='only'):
    """
    Check mandatory columns against the columns marked in
    columns and formatting file
    """
    mandatory_cols=df_columns[(df_columns['table']==table)&
                                        (df_columns['sheet']==sheet)&
                                        (~df_columns['mandatory_column'].isna())]['column_name'].to_list()
    missing_columns = [col for col in mandatory_cols if col not in cols]
    if len(missing_columns)>0:
        st.error(f"No se encontraron las siguientes columnas en el archivo {table}: {missing_columns}")
        st.stop()
def append_df_to_df(df_new=pd.DataFrame(),df_old=pd.DataFrame(),table='',keys=[],date_cols=[],allow_duplicates=False):
    """
    Appends a dataframe to an existing Dataframe
    keys: Fields that should not be duplicated, the old records are kept
    table: Table which is source of the dataframe, just to show it in the error
    date_cols: Columns transformed to date
    """
    if df_new.empty:
        return
    if not keys:
        keys=df_new.columns
    df_new=df_new.copy()
    df_new=format_dates(df_new,date_cols)
    df_old=format_dates(df_old,date_cols)
    df_new=get_common_records(df_new,df_old,keys=keys,how='uncommon')
    df_old=pd.concat([df_old,df_new])  
    df_old.reset_index(inplace=True,drop=True)
    df_grp=df_old.groupby(keys).count()
    if allow_duplicates==True:
        return df_old
    df_grp=df_grp[df_grp[df_grp.columns[0]]>1]
    if len(df_grp)>0:
        df_grp.reset_index(inplace=True)
        st.error(f"Hay duplicados en el archivo {table} para la llave {keys}:")
        st.error(df_grp[keys].sort_values(keys))
        st.stop()
    return df_old

def get_common_records(df_new=pd.DataFrame(),df_old=pd.DataFrame(),keys=[],how='common'):
    """
    Gets common records in two dataframes based on a key of columns
    when getting uncommon only records the records of the new are returned
    """
    df_old=df_old.copy()
    df_new=df_new.copy()
    df_old['composite_key'] = list(zip(*(df_old[col] for col in keys)))
    df_new['composite_key'] = list(zip(*(df_new[col] for col in keys)))
    if how=='common':
        df_new=df_new[df_new['composite_key'].isin(df_old['composite_key'])]
    else:
        df_new=df_new[~df_new['composite_key'].isin(df_old['composite_key'])]
    df_new.drop(columns=['composite_key'],inplace=True)
    return df_new


    
# =============================================================================
# Main Functions
# =============================================================================

def verify_order_list():
    #% Open order list
    path_order_list = st.session_state.selected_paths['order_file']
    df_order_list = load_excel_with_header_key(path_order_list, key_text='Priority')
    df_order_list = rename_columns(df_order_list, st.session_state.df_col_rel, table_from='Lista de ordenes')
    st.session_state.df_order_list=df_order_list
    #% Open old plan
    path_plan=st.session_state.output_paths['path_plan']
    df_plan_old=read_predefined_excel(path_plan,df_columns=st.session_state.df_columns,table='Manufacturing plan',check_mandatory=True)
    st.session_state.df_plan_old=df_plan_old
    if len(df_plan_old)==0:
        st.info("Ok")
        return
    df_plan_old=df_plan_old[~df_plan_old['status'].isna()]
    st.session_state.df_plan_old=df_plan_old
    #% Show orders already planned
    df_already_planned=get_common_records(df_new=df_order_list,df_old=df_plan_old,keys=['wo','pn'])
    if len(df_already_planned)==0:
        st.info("Ok")
        return
    st.info("Las siguientes ordenes ya estan planeadas, continue si desea agregarlas al nuevo plan con cantidad diferente")
    st.dataframe(df_already_planned)


def create_plan():
    
    if not os.path.exists(st.session_state.output_paths['path_xl_format']):
        st.error("No se encuentra el archivo: columns and formatting.xlsx")
        st.stop()

    if 'df_plan_old' not in st.session_state:
        st.error('Favor de verificar las ordenes')
        return
    
    path_master_doblado = st.session_state.selected_paths['master_file']
    df_master_doblado = load_excel_with_header_key(path_master_doblado, sheet_name='00. Formato para Master de WC', key_text='PN')
    check_mandatory_columns_df(df_master_doblado.columns,df_columns=st.session_state.df_columns,table='Master Doblado',sheet='00. Formato para Master de WC')
    df_master_doblado = rename_columns(df_master_doblado, st.session_state.df_col_rel, table_from='Master Doblado', sheet_from='00. Formato para Master de WC')

    path_routing = st.session_state.selected_paths['routing_file']
    df_routing = load_excel_with_header_key(path_routing, sheet_name='Operations', key_text='Routing')
    df_routing = rename_columns(df_routing, st.session_state.df_col_rel, table_from='Routing', sheet_from='Operations')

    path_order_list = st.session_state.selected_paths['order_file']
    df_order_list = load_excel_with_header_key(path_order_list, key_text='Priority')
    df_order_list = rename_columns(df_order_list, st.session_state.df_col_rel, table_from='Lista de ordenes')

    path_plan=st.session_state.output_paths['path_plan']
    df_plan_old=read_predefined_excel(path_plan,df_columns=st.session_state.df_columns,table='Manufacturing plan',check_mandatory=True)

    machine_status = {}
    assignments = []
    df_order_list.sort_values('priority', inplace=True)
    try:
        shifts = {'PRIMER TURNO': int(state['time_first_shift']), 'SEGUNDO TURNO': int(state['time_first_shift'])}
    except (TypeError, ValueError):
        st.error(f"Los turnos deben ser un número entero válido.")
        st.stop()
    for idx, order in df_order_list.iterrows():
        pn = order['pn']
        qty = order['pzas_x_hacer']
        wo = order['wo']
        pty = order['priority']
        pn_info = df_routing[(df_routing['pn'] == pn)&(df_routing['operation_description']==state['routing_name'])]
        if pn_info.empty:
            continue
        run_time = pn_info.iloc[0]['run_time']
        setup_time = pn_info.iloc[0]['setup_time']
        pn_machines = df_master_doblado[df_master_doblado['pn'] == pn]
        if pn_machines.empty:
            continue
        row = pn_machines.iloc[0]
        machines = []
        for key, item in row.items():
            if ('maq_opc' in key) and (item is not None) and (item != ''):
                machines.append(item)
        for m in machines:
            if m not in machine_status:
                machine_status[m] = {'day': 1, 'avail': copy(shifts), 'last_pn': None}
        if len(machines) == 0:
            print(f"Favor de asignar maquina al PN: {pn}")
            raise SystemExit()
        while qty > 0:
            assigned = False
            for m in machines:
                for shift in shifts.keys():
                    available_time = machine_status[m]['avail'][shift]
                    if machine_status[m]['last_pn'] != pn:
                        if run_time == 0:
                            pieces_to_assign = qty
                            time_used = setup_time
                        else:
                            if available_time < (setup_time + run_time):
                                continue
                            pieces_possible = 1 + int((available_time - (setup_time + run_time)) / run_time)
                            pieces_possible = max(pieces_possible, 1)
                            pieces_to_assign = min(pieces_possible, qty)
                            time_used = setup_time + pieces_to_assign * run_time
                    else:
                        if run_time == 0:
                            pieces_to_assign = qty
                            time_used = 0
                        else:
                            if available_time < run_time:
                                continue
                            pieces_possible = int(available_time // run_time)
                            pieces_to_assign = min(pieces_possible, qty)
                            time_used = pieces_to_assign * run_time

                    assignments.append({
                        'day': machine_status[m]['day'],
                        'machine': m,
                        'shift': shift,
                        'pn': pn,
                        'wo': wo,
                        'priority': pty,
                        'pzas_x_hacer': pieces_to_assign,
                        'time_used': time_used
                    })
                    machine_status[m]['avail'][shift] -= time_used
                    machine_status[m]['last_pn'] = pn
                    qty -= pieces_to_assign
                    assigned = True
                    if qty == 0:
                        break
                if qty == 0:
                    break
            if not assigned:
                for m in machines:
                    machine_status[m]['day'] += 1
                    machine_status[m]['avail'] = copy(shifts)
                    machine_status[m]['last_pn'] = None

    df_plan_new=get_predefined_df(df_columns=st.session_state.df_columns,table='Manufacturing plan')
    df_plan_new = pd.concat([df_plan_new,pd.DataFrame(assignments)],ignore_index=True)
    max_day = df_plan_new['day'].max()
    workdays = pd.bdate_range(start=pd.to_datetime(state['initial_date']), periods=max_day)
    day_to_date = {day: workdays[day - 1] for day in range(1, max_day + 1)}
    df_plan_new['date'] = df_plan_new['day'].map(day_to_date)
    df_plan_new.sort_values(['date', 'machine', 'shift', 'priority', 'wo'], inplace=True)
    df_plan_old=df_plan_old[~df_plan_old['status'].isnull()]
    df_plan_old=append_df_to_df(df_new=df_plan_new,df_old=df_plan_old,table='Manufacturing plan',keys=['wo','pn','pzas_x_hacer'],allow_duplicates=True)
    path_plan=st.session_state.output_paths['path_plan']
    df_plan_old.to_excel(path_plan,sheet_name='Manufacturing plan',index=False)
    st.session_state.df_plan_old=df_plan_old
    st.info("Plan creado")


def generate_reports():
    df_plan_old=st.session_state.df_plan_old.copy()
    machines=df_plan_old['machine'].drop_duplicates().tolist()
    group_cols=['date']
    for machine in machines:
        df=df_plan_old[df_plan_old['machine']==machine].copy()
        wb = Workbook()
        ws = wb.active
        ws.title = 'Report'

        # Write headers
        titles = list(df.columns)
        for idx, title in enumerate(titles, start=1):
            ws.cell(row=1, column=idx, value=title)

        current_row = 2
        black_fill = PatternFill(start_color='000000', end_color='000000', fill_type='solid')

        # Identify numeric columns for subtotals
        numeric_cols = ['pzas_x_hacer','time_used']

        groups = df.groupby(group_cols)
        for date, group in groups:
            start_row = current_row
            # Data rows
            for _, row in group.iterrows():
                for col_idx, col in enumerate(titles, start=1):
                    if col not in group_cols:
                        ws.cell(row=current_row, column=col_idx, value=row[col])
                current_row += 1
            end_data_row = current_row - 1

            # Subtotal row
            for col_idx, col in enumerate(titles, start=1):
                if col in numeric_cols:
                    col_letter = get_column_letter(col_idx)
                    formula = f"=SUM({col_letter}{start_row}:{col_letter}{end_data_row})"
                    ws.cell(row=current_row, column=col_idx, value=formula)
            current_row += 1
        
            # Blank line with black fill across data range
            for col_idx in range(1, len(titles) + 1):
                ws.cell(row=current_row, column=col_idx).fill = black_fill
            current_row += 1

            # Merge date cells and rotate text
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_data_row+1, end_column=1)
            date_cell = ws.cell(row=start_row, column=1, value=date[0])
            date_cell.alignment = Alignment(textRotation=90, horizontal='center', vertical='center')
            date_cell.number_format = 'mmmm, d, yyyy'
        base, ext = os.path.splitext(st.session_state.output_paths['path_report'])
        wb.save(f"{base} {machine}{ext}")

def manage_file_selector(selector_key, display_label, state):
    if not state["selections"].get(selector_key):
        if st.button(f"{display_label} File", key=f"select_{selector_key}"):
            files = open_file_selection(initialdir=state["folder_output"] or os.getcwd())
            if files:
                state["selections"][selector_key] = files[0]
                save_state_pickle(state)
                st.rerun()
        st.info(f"{display_label} no seleccionado.")
    else:
        st.success(f"Selected {display_label} File: {state['selections'][selector_key]}")
        if st.button(f"Change {display_label} File", key=f"change_{selector_key}"):
            state["selections"][selector_key] = ""
            save_state_pickle(state)
            st.rerun()

# =============================================================================
# Main Script: Streamlit App UI
# =============================================================================

state = load_state_pickle()
st.session_state.folder_output = state['folder_output']
st.session_state.selected_paths = state['selections']
st.session_state.output_paths = set_paths(st.session_state.folder_output)
st.session_state.col_rel = set_col_rel(st.session_state.output_paths)
st.session_state.df_col_rel = st.session_state.col_rel['col_rel']
st.session_state.df_columns = st.session_state.col_rel['columns']

st.set_page_config(page_title="Plan de manufactura", page_icon=":factory:")
st.markdown(
    r"""
    <style>
    .stAppDeployButton {
            visibility: hidden;
        }
    </style>
    """, unsafe_allow_html=True
)
st.title("Plan de manufactura")

state = load_state_pickle()

st.header("Parámetros de turno y routing")
config_items = [
    ("time_first_shift", "Time First Shift"),
    ("time_second_shift", "Time Second Shift"),
    ("routing_name", "Routing Name"),
]
for key, label in config_items:
    value = st.text_input(label, value=state.get(key, ""), key=key)
    if value != state.get(key, ""):
        state[key] = value
        save_state_pickle(state)
        st.rerun()

st.header("Seleccionar carpeta de trabajo")
if st.button("Seleccionar carpeta", key="select_folder"):
    folder = select_directory(initialdir=state["folder_output"] or os.getcwd())
    if folder:
        state["folder_output"] = folder
        save_state_pickle(state)
        st.rerun()
state = load_state_pickle()
if state["folder_output"]:
    st.success(f"Carpeta de trabajo: {state['folder_output']}")
else:
    st.info("No seleccionado.")

st.header("Fecha inicial de programacion")
selected_date = st.date_input("Fecha inicial de programacion", value=state.get("initial_date", datetime.date.today()))
if selected_date != state.get("initial_date", datetime.date.today()):
    state["initial_date"] = selected_date
    save_state_pickle(state)
    st.rerun()

st.header("Seleccion de archivos")
file_selectors = [
    ("master_file", "Master Doblado"),
    ("order_file", "Lista de Ordenes"),
    ("routing_file", "Routing")
]
for selector_key, display_label in file_selectors:
    manage_file_selector(selector_key, display_label, state)


st.header("Crear Plan")
if st.button("Verificar ordenes"):
    verify_order_list()

if st.button("Crear Plan"):
    if not (state.get("folder_output") and 
            state["selections"].get("master_file") and 
            state["selections"].get("order_file") and 
            state["selections"].get("routing_file")):
        st.error("Por favor seleccione los archivos mandatorios.")
    else:
        st.success("Creando plan...")
        create_plan()

if st.button("Generar reportes"):
    generate_reports()
