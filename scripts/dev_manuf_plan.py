#%%
import pandas as pd
import os
from copy import copy
from importlib import reload
from manufacturing_plan import *
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter
state = load_state_pickle()

folder_output = state['folder_output']
file_selectors = state['selections']
output_paths = set_paths(folder_output)
col_rel = set_col_rel(output_paths)
df_col_rel = col_rel['col_rel']
df_columns=col_rel['columns']
#%%
#%%


if not os.path.exists(output_paths['path_xl_format']):
    print("No se encuentra el archivo: columns and formatting.xlsx")
    raise SystemExit()

path_available_hours=file_selectors['available_hours']
df_avail_hours=read_excel(path=path_available_hours)
df_avail_hours=df_avail_hours[['maquina','tiempo_primer_turno','tiempo_segundo_turno']]
dict_avail_hours=df_avail_hours.set_index('maquina').to_dict(orient='index')

#%%
path_master_doblado = file_selectors['master_file']
df_master_doblado = load_excel_with_header_key(path_master_doblado, sheet_name='00. Formato para Master de WC', key_text='PN')
df_master_doblado = rename_columns(df_master_doblado, df_col_rel, table_from='Master Doblado', sheet_from='00. Formato para Master de WC')
df_master_doblado.replace('/','_',regex=True,inplace=True)
#%%
path_routing = file_selectors['routing_file']
df_routing = load_excel_with_header_key(path_routing, sheet_name='Operations', key_text='Routing')
df_routing = rename_columns(df_routing, df_col_rel, table_from='Routing', sheet_from='Operations')
#%% Open order list
path_order_list = file_selectors['order_file']
df_order_list = load_excel_with_header_key(path_order_list, key_text='Priority')
df_order_list = rename_columns(df_order_list, df_col_rel, table_from='Lista de ordenes')

#%% Open old plan
path_plan=output_paths['path_plan']
df_plan_old=read_predefined_excel(path_plan,df_columns=df_columns,table='Manufacturing plan',check_mandatory=True)
df_plan_old=rename_columns(df_plan_old,df_col_rel=df_col_rel,table_from='Manufacturing plan')
df_plan_old=df_plan_old[~df_plan_old['status'].isna()]

#%% Show orders already planned
df_already_planned=get_common_records(df_new=df_order_list,df_old=df_plan_old,keys=['wo','pn'])
st.info("Las siguientes ordenes ya estan planeadas, continue si desea agregarlas al nuevo plan")
st.info(df_already_planned)

#%%
machine_status = {}
assignments = []
df_order_list.sort_values('priority', inplace=True)

shifts = {'PRIMER TURNO': int(state['time_first_shift']), 'SEGUNDO TURNO': int(state['time_first_shift'])}
for idx, order in df_order_list.iterrows():
    pn = order['pn']
    qty = order['pzas_x_hacer']
    wo = order['wo']
    pty = order['priority']
    machine_default=order['machine']
    pn_info = df_routing[(df_routing['pn'] == pn)&(df_routing['operation_description']==state['routing_name'])]
    if pn_info.empty:
        continue
    run_time = pn_info.iloc[0]['run_time']
    setup_time = pn_info.iloc[0]['setup_time']
    if machine_default:
        machines=[machine_default]
    else:
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
            if not m in dict_avail_hours:
                print(f"Falta definir horas disponibles en la maquina {m}")
                raise SystemExit()

            machine_status[m] = {'day': 1, 'avail': copy(dict_avail_hours[m]), 'last_pn': None}
    if len(machines) == 0:
        print(f"Favor de asignar maquina al PN: {pn}")
        raise SystemExit()
    
    while qty > 0:
        assigned = False
        for m in machines:
            for shift in machine_status[m]['avail'].keys():
                available_time = machine_status[m]['avail'][shift]
                if machine_status[m]['last_pn'] != pn:
                    if run_time == 0:
                        pieces_to_assign = qty
                        time_used = 0
                    else:
                        if available_time < (run_time):
                            continue
                        pieces_possible = 1 + int((available_time - (run_time)) / run_time)
                        pieces_possible = max(pieces_possible, 1)
                        pieces_to_assign = min(pieces_possible, qty)
                        time_used = pieces_to_assign * run_time
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
                    'time_used': time_used,
                    'setup_time':setup_time
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
                machine_status[m]['avail'] = copy(dict_avail_hours[m])
                machine_status[m]['last_pn'] = None

df_plan_new=get_predefined_df(df_columns=df_columns,table='Manufacturing plan')
df_plan_new = pd.concat([df_plan_new,pd.DataFrame(assignments)],ignore_index=True)
max_day = df_plan_new['day'].max()
workdays = pd.bdate_range(start=pd.to_datetime(state['initial_date']), periods=max_day)
day_to_date = {day: workdays[day - 1] for day in range(1, int(max_day) + 1)}
df_plan_new['date'] = df_plan_new['day'].map(day_to_date)
df_plan_new.sort_values(['date', 'machine', 'shift', 'priority', 'wo'], inplace=True)
df_plan_new['pzas_x_hora']=(df_plan_new['pzas_x_hacer']/df_plan_new['time_used']).astype(float).round(2)
df_plan_old=df_plan_old[~df_plan_old['status'].isnull()]
df_plan_old=append_df_to_df(df_new=df_plan_new,df_old=df_plan_old,table='Manufacturing plan',keys=['wo','pn'],allow_duplicates=True)

plan_cols=df_columns[(df_columns['table']=='Manufacturing plan')&
           (~df_columns['mandatory_column'].isna())]['std_name'].to_list()
df_plan_old=df_plan_old[plan_cols]
df_plan_old=rename_columns(df_plan_old,df_col_rel=df_col_rel,table_to='Manufacturing plan')
df_plan_old.to_excel(path_plan,index=False)
#%%
df_plan_old

#%%
df=df_plan_old
#%%


#%% Reportes
machines=df_plan_old['machine'].drop_duplicates().tolist()
group_cols=['date']
plan_report_cols=df_columns[(df_columns['table']=='Machine Report')&
    (~df_columns['mandatory_column'].isna())]['std_name'].to_list()
df_plan_old=df_plan_old[plan_report_cols]

#%%aqui voy
machine_col=df_columns[(df_columns['table']=='Machine Report')&
    (df_columns['std_name']=='machine')]['column_name'].to_list()[0]
#%%
df_columns[(df_columns['table']=='Machine Report')]
#%%

for machine in machines:
    df=df_plan_old[df_plan_old['machine']==machine].copy()
    if len(df)==0:
        continue
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
    base, ext = os.path.splitext(output_paths['path_report'])
    wb.save(f"{base} {machine}{ext}")
# %%

