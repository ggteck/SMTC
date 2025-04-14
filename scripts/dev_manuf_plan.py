#%%
import pandas as pd
import os
from copy import copy
from manufacturing_plan import (open_file_selection,
select_directory,
set_paths,
set_col_rel,
show_popup_message,
read_excel,
load_excel_with_header_key,
save_state_pickle,
load_state_pickle,
rename_columns,
create_plan,
manage_file_selector)

#%%
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
    pn_info = df_routing[(df_routing['pn'] == pn)&(df_routing['operation_description']=='DOBLADO')]
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

#%%
