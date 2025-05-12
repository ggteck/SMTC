#%%
import pandas as pd
import os
from copy import copy
from importlib import reload
from manufacturing_plan import *
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pandas.tseries.offsets import BDay
state = load_state_pickle()

folder_output = state['folder_output']
file_selectors = state['selections']
output_paths = set_paths(folder_output)
col_rel = set_col_rel(output_paths)
df_col_rel = col_rel['col_rel']
df_columns=col_rel['columns']
#%
if not os.path.exists(output_paths['path_xl_format']):
    print("No se encuentra el archivo: columns and formatting.xlsx")
    raise SystemExit()

path_available_hours=file_selectors['available_hours']
df_avail_hours=read_excel(path=path_available_hours)
df_avail_hours['dia']=pd.to_datetime(df_avail_hours['dia'],errors='coerce').dt.strftime('%Y-%m-%d')
df_avail_hours['dia'].fillna('default',inplace=True)
dict_avail_hours=df_avail_hours.set_index(['dia','maquina']).to_dict(orient='index')

#%
path_master_doblado = file_selectors['master_file']
df_master_doblado = load_excel_with_header_key(path_master_doblado, sheet_name='00. Formato para Master de WC', key_text='PN')
df_master_doblado = rename_columns(df_master_doblado, df_col_rel, table_from='Master Doblado', sheet_from='00. Formato para Master de WC')
df_master_doblado.replace('/','_',regex=True,inplace=True)
#%
path_routing = file_selectors['routing_file']
df_routing = load_excel_with_header_key(path_routing, sheet_name='Operations', key_text='Routing')
df_routing = rename_columns(df_routing, df_col_rel, table_from='Routing', sheet_from='Operations')
#%%
# Equivalencias
path_equiv = st.session_state.selected_paths['equivalencias_file']
df_equiv=read_predefined_excel(path_equiv,df_columns,table='Equivalencias')
dict_equiv=df_equiv.set_index('pn')['equivalencia'].to_dict()
dict_equiv_inv={v: k for k, v in dict_equiv.items()}
#% Open order list
path_order_list = file_selectors['order_file']
df_order_list = load_excel_with_header_key(path_order_list, key_text='Priority')
df_order_list = rename_columns(df_order_list, df_col_rel, table_from='Lista de ordenes')
df_order_list['pn_orig']=df_order_list['pn']
df_order_list['pn']=df_order_list['pn'].replace(dict_equiv)
check_mandatory_columns_df(df_order_list.columns,df_columns=df_columns,table='Lista de ordenes')

df_missing_rout=get_common_records(df_order_list,df_routing,keys=['pn','operation_description'],how='uncommon')
df_order_list['composite_key'] = list(zip(*(df_order_list[col] for col in ['pn','operation_description'])))
df_routing['composite_key'] = list(zip(*(df_routing[col] for col in ['pn','operation_description'])))
df_missing_rout=df_order_list[~df_order_list['composite_key'].isin(df_routing['composite_key'])]
df_missing_rout['composite_key'] = list(zip(*(df_missing_rout[col] for col in ['pn_orig','operation_description'])))
df_missing_rout=df_missing_rout[~df_missing_rout['composite_key'].isin(df_routing['composite_key'])]
df_missing_machine=df_order_list[(~df_order_list['pn'].isin(df_master_doblado['pn']))&
                                 (~df_order_list['pn_orig'].isin(df_master_doblado['pn']))&
                                 (df_order_list['machine']=='')]
df_order_list.drop(columns=['composite_key'],inplace=True)

#% Open part master
path_part_master = file_selectors['part_master']
df_part_master = load_excel_with_header_key(path_part_master, key_text='Site')
df_part_master = rename_columns(df_part_master, df_col_rel, table_from='Part Master')
check_mandatory_columns_df(df_part_master.columns,df_columns=df_columns,table='Part Master')
df_part_master=df_part_master[['pn','unit_selling_price']].drop_duplicates(subset=['pn'],keep='last')
df_part_master_orig=df_part_master.copy()
df_part_master['pn']=df_part_master['pn'].replace(dict_equiv_inv)
df_part_master=pd.concat([df_part_master_orig,df_part_master])
df_part_master=df_part_master.drop_duplicates(subset=['pn'],keep='last')
#%%
df_part_master[df_part_master['pn'].str.contains('ENC-1500-120-20')]
#%%

#%%
df_part_master['pn'].replace(dict_equiv)
#%%
# df_part_master=get_common_records(df_part_master,df_order_list,keys=['pn'])
#% Open old plan
path_plan=output_paths['path_plan']
df_plan_old=read_predefined_excel(path_plan,df_columns=df_columns,table='Manufacturing plan',check_mandatory=True)
df_plan_old=rename_columns(df_plan_old,df_col_rel=df_col_rel,table_from='Manufacturing plan')
df_plan_old=df_plan_old[~df_plan_old['status'].isna()]

#% Show orders already planned
df_already_planned=get_common_records(df_new=df_order_list,df_old=df_plan_old,keys=['wo','pn'])
st.info("Las siguientes ordenes ya estan planeadas, continue si desea agregarlas al nuevo plan")
st.info(df_already_planned)
#%%
machine_status = {}
assignments = []
df_order_list.sort_values('priority', inplace=True)
initial_date   = pd.to_datetime(state['initial_date']).strftime('%Y-%m-%d')
limit_date_ts  = pd.to_datetime(state['limit_date'])

for idx, order in df_order_list.iterrows():
    pn_orig        = order['pn_orig']
    pn             = order['pn']
    qty            = order['pzas_x_hacer']
    wo             = order['wo']
    pty            = order['priority']
    machine_default= order['machine']
    routing_name   = order['operation_description']

    # routing lookup
    pn_info = df_routing[(df_routing['pn']==pn) & (df_routing['operation_description']==routing_name)]
    if pn_info.empty:
        pn_info = df_routing[(df_routing['pn']==pn_orig) & (df_routing['operation_description']==routing_name)]
        if pn_info.empty:
            continue
    run_time   = pn_info.iloc[0]['run_time']
    setup_time = pn_info.iloc[0]['setup_time']

    # machine selection
    if machine_default:
        machines = [machine_default]
    else:
        rows = df_master_doblado[df_master_doblado['pn']==pn]
        if rows.empty:
            rows = df_master_doblado[df_master_doblado['pn']==pn_orig]
            if rows.empty:
                continue
        machines = [v for k,v in rows.iloc[0].items() if 'maq_opc' in k and v]

    # init status
    for m in machines:
        if m not in machine_status:
            key = (initial_date, m)
            if key not in dict_avail_hours and ('default',m) not in dict_avail_hours:
                print(f"Falta definir horas disponibles en la maquina {m}")
                raise SystemExit()
            base = dict_avail_hours.get(key, dict_avail_hours[('default',m)])
            machine_status[m] = {
                'date':    initial_date,
                'avail':   copy(base),
                'last_pn': None
            }

    if not machines:
        print(f"Favor de asignar maquina al PN: {pn}")
        raise SystemExit()

    # assignment per machine
    for m in machines:
        while qty > 0:
            current = pd.to_datetime(machine_status[m]['date'])
            if current > limit_date_ts:
                break
            assigned = False
            # iterate shifts in order
            for shift, avail in machine_status[m]['avail'].items():
                # decide setup
                setup = setup_time if machine_status[m]['last_pn']!=pn else 0
                # max pieces this shift
                if run_time>0:
                    max_pieces = int((avail - setup)//run_time) if avail>=setup+run_time else 0
                else:
                    max_pieces = qty
                if max_pieces<=0:
                    machine_status[m]['avail'][shift] = 0
                    continue  # not enough room
                pieces = min(max_pieces, qty)
                run_used = pieces*run_time
                total_used = run_used + setup
                # perform assignment
                assignments.append({
                    'date':  machine_status[m]['date'],
                    'machine': m,
                    'shift': shift,
                    'operation_description': routing_name,
                    'pn':    pn_orig,
                    'wo':    wo,
                    'priority': pty,
                    'pzas_x_hacer': pieces,
                    'time_used':    run_used,
                    'setup_time':   setup
                })
                # update status
                machine_status[m]['avail'][shift] -= total_used
                machine_status[m]['last_pn'] = pn
                qty -= pieces
                assigned = True
                break
            if not assigned:
                # advance to next day
                nd = (current + BDay(1)).strftime('%Y-%m-%d')
                machine_status[m]['date']  = nd
                key = (nd, m)
                base = dict_avail_hours.get(key, dict_avail_hours[('default',m)])
                machine_status[m]['avail'] = copy(base)
            if qty==0:
                break
        if qty==0:
            break

df_plan_new=get_predefined_df(df_columns=df_columns,table='Manufacturing plan')
df_plan_new = pd.concat([df_plan_new,pd.DataFrame(assignments)],ignore_index=True)
df_plan_new.sort_values(['date', 'machine', 'shift', 'priority', 'wo'], inplace=True)
df_plan_new['pzas_x_hora']=(df_plan_new['pzas_x_hacer']/df_plan_new['time_used']).astype(float).round(2)
df_plan_old=df_plan_old[~df_plan_old['status'].isnull()]
df_plan_old=append_df_to_df(df_new=df_plan_new,df_old=df_plan_old,table='Manufacturing plan',keys=['wo','pn'],allow_duplicates=True)

#%%
#%%
df_part_master
#%%
if 'unit_selling_price' in df_plan_old.columns:
    df_plan_old.drop(columns=['unit_selling_price'],inplace=True)
df_plan_old=df_plan_old.merge(df_part_master,on='pn',how='left')    
df_plan_old['unit_selling_price']=df_plan_old['unit_selling_price']*df_plan_old['pzas_x_hacer']
plan_cols=df_columns[(df_columns['table']=='Manufacturing plan')&
           (~df_columns['mandatory_column'].isna())]['std_name'].to_list()
df_plan_old=df_plan_old[plan_cols]
df_plan_old['date']=pd.to_datetime(df_plan_old['date'],format='mixed').dt.strftime('%B, %#d, %Y')
df_plan_old=rename_columns(df_plan_old,df_col_rel=df_col_rel,table_to='Manufacturing plan')
df_plan_old.to_excel(path_plan,index=False)
os.startfile(path_plan)
#%%
#%%
df_plan_old=rename_columns(df_plan_old,df_col_rel=st.session_state.df_col_rel,table_from='Manufacturing plan')    
#% Reportes
machines=df_plan_old['machine'].drop_duplicates().tolist()
group_cols=['date']
plan_report_cols=df_columns[(df_columns['table']=='Machine Report')&
    (~df_columns['mandatory_column'].isna())]['std_name'].to_list()
df_plan_old=df_plan_old[plan_report_cols]
machine_col=df_columns[(df_columns['table']=='Machine Report')&
    (df_columns['std_name']=='machine')]['column_name'].to_list()[0]

df_columns[(df_columns['table']=='Machine Report')]

dict_formats=get_xl_formatting()
special_formats=dict_formats['special_format']
col_sizes=None
for file in os.listdir(folder_output):
    if 'reporte de manufactura' in file:
        wb=load_workbook(os.path.join(folder_output,file))
        col_sizes=get_col_sizes(wb)
        break

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
    for date, group in sorted(groups, key=lambda x: pd.to_datetime(x[0])):
        start_row = current_row
        # Data rows
        for _, row in group.iterrows():
            for col_idx, col in enumerate(titles, start=1):
                if col not in group_cols:
                    cell=ws.cell(row=current_row, column=col_idx, value=row[col])
                    cell.value=row[col]
                    if row[col] in special_formats.keys():
                        format=special_formats[row[col]]
                        format_cell(cell,format)
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
    if col_sizes:
        wb=apply_col_sizes(wb,col_sizes)
    base, ext = os.path.splitext(output_paths['path_report'])
    wb.save(f"{base} {machine}{ext}")
# %%

#%% Validar plan
close_xl_if_open(path_order_list)
close_xl_if_open(path_plan)
df_plan=read_excel(path_plan)
df_plan=rename_columns(df_plan,df_col_rel=df_col_rel,table_from='Manufacturing plan')
df_plan_grp=df_plan.groupby(['pn','wo']).sum(['pzas_x_hacer'])[['pzas_x_hacer']]
df_plan_grp.reset_index(inplace=True)
df_order_assignment=df_order_list.merge(df_plan_grp,how='left',on=['pn','wo'],suffixes=('','_assigned'))
df_order_assignment['pzas_x_asignar']=df_order_assignment['pzas_x_hacer']-df_order_assignment['pzas_x_hacer_assigned']
df_order_assignment.loc[df_order_assignment['pzas_x_asignar']==0,'status']='Planeada'
df_order_assignment.loc[df_order_assignment['pzas_x_asignar']>0,'status']='Incompleta'
df_order_assignment.loc[df_order_assignment['pzas_x_asignar'].isna(),'status']='Pendiente'
df_order_assignment=df_order_assignment.drop(columns=['pzas_x_hacer_assigned','pzas_x_asignar'])
orders_cols=df_columns[(df_columns['table']=='Lista de ordenes')]['std_name'].to_list()
df_order_assignment=df_order_assignment[orders_cols]
df_order_assignment=rename_columns(df_order_assignment,df_col_rel,table_to='Lista de ordenes')
df_order_assignment.to_excel(path_order_list,index=False)
df_plan.loc[df_plan['status'].isna(),'status']='Planeada'
df_plan.to_excel(path_plan,sheet_name='Manufacturing plan',index=False)

