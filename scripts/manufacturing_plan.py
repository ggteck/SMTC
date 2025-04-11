# Manufacturing plan
# - V1. 2025-04-07
#     - Version inicial, calculo de plan de produccion

import streamlit as st
import pickle
import os
from tkinter import Tk, filedialog as fd
import datetime
import pandas as pd
from copy import copy
from IPython.display import display, Markdown

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
    return output_paths

def set_col_rel(output_paths):
    df_columns = read_excel(output_paths['path_xl_format'], sheet_name='column_format')
    df_col_rel = df_columns[~df_columns['std_name'].isnull()].copy()
    return {'col_rel': df_col_rel, 'columns': df_columns}

# =============================================================================
# User Interface Management
# =============================================================================

def show_popup_message(message, df=pd.DataFrame()):
    display(Markdown(f"### **{message}**"))
    if not df.empty:
        display(Markdown(df.to_markdown(index=False)))

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

# =============================================================================
# Main Functions
# =============================================================================

def create_plan():
    state = load_state_pickle()
    folder_output = state['folder_output']
    file_selectors = state['selections']
    output_paths = set_paths(folder_output)
    col_rel = set_col_rel(output_paths)
    df_col_rel = col_rel['col_rel']

    if not os.path.exists(output_paths['path_xl_format']):
        show_popup_message("No se encuentra el archivo: columns and formatting.xlsx")
        raise SystemExit()

    path_master_doblado = file_selectors['master_file']
    df_master_doblado = load_excel_with_header_key(path_master_doblado, sheet_name='00. Formato para Master de WC', key_text='PN')
    df_master_doblado = rename_columns(df_master_doblado, df_col_rel, table_from='Master Doblado', sheet_from='00. Formato para Master de WC')

    path_routing = file_selectors['routing_file']
    df_routing = load_excel_with_header_key(path_routing, sheet_name='Operations', key_text='Routing')
    df_routing = rename_columns(df_routing, df_col_rel, table_from='Routing', sheet_from='Operations')

    path_order_list = file_selectors['order_file']
    df_order_list = load_excel_with_header_key(path_order_list, key_text='Priority')
    df_order_list = rename_columns(df_order_list, df_col_rel, table_from='Lista de ordenes')

    machine_status = {}
    assignments = []
    df_order_list.sort_values('priority', inplace=True)
    shifts = {'PRIMER TURNO': 8.0, 'SEGUNDO TURNO': 7.0}
    for idx, order in df_order_list.iterrows():
        pn = order['pn']
        qty = order['pzas_x_hacer']
        wo = order['wo']
        pty = order['priority']
        pn_info = df_routing[df_routing['pn'] == pn]
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
            show_popup_message(f"Favor de asignar maquina al PN: {pn}")
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

    report = pd.DataFrame(assignments)
    max_day = report['day'].max()
    workdays = pd.bdate_range(start=pd.to_datetime(state['initial_date']), periods=max_day)
    day_to_date = {day: workdays[day - 1] for day in range(1, max_day + 1)}
    report['date'] = report['day'].map(day_to_date)
    report.sort_values(['date', 'machine', 'shift', 'priority', 'wo'], inplace=True)
    report.to_excel(output_paths['path_plan'], index=False)

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
if st.button("Crear Plan"):
    if not (state.get("folder_output") and 
            state["selections"].get("master_file") and 
            state["selections"].get("order_file") and 
            state["selections"].get("routing_file")):
        st.error("Por favor seleccione los archivos mandatorios.")
    else:
        st.success("Creando plan...")
        st.write("Output Folder:", state["folder_output"])
        st.write("Fecha inicial de programacion:", state.get("initial_date"))
        st.write("Master Doblado File:", state["selections"]["master_file"])
        st.write("Lista de Ordenes File:", state["selections"]["order_file"])
        st.write("Routing File:", state["selections"]["routing_file"])
        create_plan()
