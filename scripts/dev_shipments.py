    #%%

import os, shutil
import pickle
import pandas as pd
from datetime import date, timedelta, datetime
from tkinter import Tk, filedialog as fd

import win32com.client
from local_shipments_l import  *

state=load_state_pickle()
#%%
state
# %%
#%%

st.write("Ejecutando función: Explorar Outlook")
# Ensure a folder was selected from Outlook
if "mail_folder" not in st.session_state or st.session_state.mail_folder is None:
    st.error("No se ha seleccionado un folder de Outlook.")
    
# Ensure the attachments folder exists
if not os.path.exists(output_paths.get("path_attachments", "")):
    os.makedirs(output_paths['path_attachments'])
# Get the date limit from state (fecha_mail)
if state.get("fecha_mail") is None:
    st.error("La fecha mail no está definida en el estado.")
    
date_limit = datetime.combine(state["fecha_mail"], datetime.min.time())
folder = st.session_state.mail_folder
try:
    items = folder.Items
except Exception as e:
    st.error("Error accediendo a los elementos del folder de Outlook.")
 #%%   
count = 0
for message in items:
    try:
        received_time = message.ReceivedTime
        # Convert COM date to Python datetime
        received_time = datetime.fromtimestamp(received_time.timestamp())
    except Exception:
        continue
    if received_time < date_limit:
        continue
    if message.Attachments.Count > 0:
        for i in range(1, message.Attachments.Count + 1):
            try:
                attachment = message.Attachments.Item(i)
                if "KORRUS" not in attachment.FileName.upper():
                    continue
                base, ext = os.path.splitext(attachment.FileName)
                new_name = f"{base}_{received_time.strftime('%Y-%m-%d-%H%M%S')}{ext}"
                save_path = os.path.join(output_paths['path_attachments'], new_name)
                attachment.SaveAsFile(save_path)
                st.write(f"Attachment saved: {save_path}")
                count += 1
            except Exception as e:
                st.write(f"Error saving attachment: {str(e)}")
                continue
st.success(f"Exploración completada. {count} attachments guardados.")

# Crear la lista de archivos descargados para manejar su integración.
if not os.path.exists(output_paths['path_attachments_done']):
    os.makedirs(output_paths['path_attachments_done'])

# Status maneja el proceso a realizar en los archivos: r=reprocesar, d=done, e=error
if os.path.exists(output_paths['path_korrus_list']):
    df_korrus_list = pd.read_excel(output_paths['path_korrus_list'])
else:
    df_korrus_list = pd.DataFrame(columns=['file_name', 'received_time', 'status'])
close_xl_if_open(output_paths['path_korrus_list'])

files = os.listdir(output_paths['path_attachments'])
df_korrus_list_new = pd.DataFrame(columns=['file_name'], data=files)
df_korrus_list_new['received_time'] = df_korrus_list_new['file_name'].str.extract(r'(\d{4}-\d{2}-\d{2}-\d{6})')
df_korrus_list_new = df_korrus_list_new.merge(df_korrus_list, how='outer', on=['file_name', 'received_time'])
df_korrus_list_new.dropna(subset=['received_time'], inplace=True)
df_korrus_list_new['status'].fillna('', inplace=True)

# Analizar archivos adjuntos
df_korrus_data_new = pd.DataFrame(columns=['origin_file'] + mandatory_cols['Korrus'])
for index, row in df_korrus_list_new[df_korrus_list_new['status'].isin(['', 'r'])].iterrows():
    filepath = os.path.join(output_paths['path_attachments'], row['file_name'])
    if (row['status'] == 'r') or (not os.path.exists(filepath)):
        filepath = os.path.join(output_paths['path_attachments_done'], row['file_name'])
    if not os.path.exists(filepath):
        df_korrus_list_new.loc[index, 'status'] = 'e'
        print(f"El archivo {row['file_name']} no existe")
        continue
    ext = os.path.splitext(os.path.basename(filepath))[1].lower()
    if ext == ".xlsx":
        df = read_excel(filepath)
    elif ext == ".csv":
        try:
            df = pd.read_csv(filepath, sep='\t', encoding='utf-16')
        except:
            df = pd.read_csv(filepath, sep=',')
    else:
        df_korrus_list_new.loc[index, 'status'] = 'e'
        print(f"El archivo {row['file_name']} no pudo ser procesado")
        continue
    if check_mandatory_cols(df.columns, selector_name='Korrus', raise_error=False):
        df['origin_file'] = row['file_name']
        df_korrus_data_new = pd.concat([df_korrus_data_new, df])
        df_korrus_data_new = df_korrus_data_new[['origin_file'] + mandatory_cols['Korrus']]
        df_korrus_list_new.loc[index, 'status'] = 'n'
    else:
        df_korrus_list_new.loc[index, 'status'] = 'e'

# Consolidar datos en un solo archivo
if os.path.exists(output_paths['path_korrus_data']):
    df_korrus_data = pd.read_excel(output_paths['path_korrus_data'])
else:
    df_korrus_data = pd.DataFrame(columns=['origin_file'] + mandatory_cols['Korrus'])
close_xl_if_open(output_paths['path_korrus_data'])
df_korrus_data = df_korrus_data[~df_korrus_data['origin_file'].isin(df_korrus_data_new['origin_file'])]
df_korrus_data_new = pd.concat([df_korrus_data, df_korrus_data_new])
#%%
# Mover archivos a la carpeta de archivos procesados
for index,row in df_korrus_list_new[df_korrus_list_new['status']=='n'].iterrows():
    df_korrus_list_new.loc[index,'status']='y'
    if os.path.exists(f'{output_paths['path_attachments']}/{row["file_name"]}'):
        shutil.move(f'{output_paths['path_attachments']}/{row["file_name"]}', f'{output_paths['path_attachments_done']}/{row["file_name"]}')
# Eliminar archivos ya procesados del folder de llegada
df_remove=df_korrus_list_new[(df_korrus_list_new['status']=='y') & (df_korrus_list_new['file_name'].isin(files))]
for index,row in df_remove.iterrows():
    if os.path.exists(os.path.join(output_paths['path_attachments'],row['file_name'])):
        os.remove(os.path.join(output_paths['path_attachments'],row['file_name']))
files=os.listdir(output_paths['path_attachments'])
# Guardar datos y lista de archivos procesados
save_df(df_korrus_list_new,filepath=output_paths['path_korrus_list'],sheet_name='Korrus List',index=False)
save_df(df_korrus_data_new,filepath=output_paths['path_korrus_data'],sheet_name='Korrus Data',index=False)
st.success("Proceso de integración completado.")

#%% Leer lista de precios
folder_prices=state.get('folder_prices', "Not selected")
if folder_prices=="Not selected":
    st.info(f'Favor de seleccionar los folder con listas de precios')
    raise SystemExit()
prices_files_lst=os.listdir(folder_prices)
df_prices=pd.DataFrame()
for file in prices_files_lst:
    if '~' in file:
        continue
    filepath=os.path.join(folder_prices,file)
    close_xl_if_open(filepath)
    if 'ACCESSORIES' in file:
        df=load_excel_with_header_key(filepath,key_text='Site')
        df=rename_columns(df,df_col_rel=df_col_rel,table_from='Price accessories')
    else:
        df=load_excel_with_header_key(filepath,key_text='Final SKU')
        df=rename_columns(df,df_col_rel=df_col_rel,table_from='Prices fixtures')
    df=df[['modelo','price']]
    df_prices=pd.concat([df_prices,df])
df_prices.drop_duplicates(['modelo'],inplace=True)
df_prices.reset_index(inplace=True,drop=True)

#%%
# %% [markdown]
# ## Generar reportes
# - Hay tres reportes EDI Master, Shipped to Cust, Shipped to ELP
# - Si hay archivos seleccionados se integran a estos reportes


# ### Consolida
# - Korrus_data --> EDI Master
# - InventoryStage --> Shipment to ELP
# - Shipment transactions: Standalone file
#

# Consolidar reportes 
path_ship_elp=get_path(state,'ELP Master')
close_xl_if_open(path_ship_elp)
path_oor_old=get_path(state,'OOR')
close_xl_if_open(path_oor_old)
wb_elp=load_workbook(path_ship_elp)
wb_elp._external_links.clear()
ws_edi=wb_elp['EDI Master']
ws_dict_edi=get_worksheet_df(ws_edi,key_text='PO',data_only=True)
df_edi=ws_dict_edi['df']
df_edi['ProductService ID']=df_edi['ProductService ID'].str.upper()
df_edi['Quantity']=df_edi['Quantity'].astype(float)
ws_ship_elp=wb_elp['Shipment to ELP']
ws_dict_ship_elp=get_worksheet_df(ws_ship_elp,key_text='CUU ship Date',data_only=True)
df_ship_elp=ws_dict_ship_elp['df']

df=pd.read_excel(path_ship_elp,sheet_name='Shipment to ELP')

if not os.path.exists(output_paths['path_xl_format']):
    st.info("No se encuentra el archivo: columns and formatting.xlsx")
    raise SystemExit()

# Demanda del cliente
if not os.path.exists(output_paths['path_korrus_data']):
    df=pd.DataFrame(columns=[mandatory_cols['Korrus']+['origin_file']])
    save_df(df,filepath=output_paths['path_korrus_data'],sheet_name='Korrus Data',index=False)

df_korrus_data_new=pd.read_excel(output_paths['path_korrus_data'])
if (len(df_korrus_data_new)>0):
    df_korrus_data_new.loc[:, ~df_korrus_data_new.columns.str.startswith('Unnamed:')]
    df_korrus_data_new=df_korrus_data_new[~df_korrus_data_new['PurchaseOrder'].str.contains('---')]
    df_korrus_data_new['PODate']=pd.to_datetime(df_korrus_data_new['PODate'],format='mixed', errors='coerce')
    df_korrus_data_new['LineNumber']=df_korrus_data_new['LineNumber'].astype(int)
    df_korrus_data_new['AssignedDropZone'].fillna('NULL',inplace=True)
    df_korrus_data_new['ProductServiceID']=df_korrus_data_new['ProductServiceID'].str.upper()

    key_korrus=['PurchaseOrder','LineNumber','PODate','ProductServiceID','AssignedDropZone']
    df_korrus_data_new.drop_duplicates(subset=key_korrus,keep='last',inplace=True)
    
    # Eliminar registros no requeridos
    df_korrus_data_new['EDI Received']=df_korrus_data_new['origin_file'].str.extract(r'(\d{4}-\d{2}-\d{2})')
    df_korrus_data_new['EDI Received']=pd.to_datetime(df_korrus_data_new['EDI Received'],format='mixed', errors='coerce')
    df_korrus_data_new.drop('origin_file',axis=1,inplace=True)
    df_korrus_data_new.drop_duplicates(inplace=True)
    df_korrus_data_new['PODate']=pd.to_datetime(df_korrus_data_new['PODate'])
    df_korrus_data_new.rename({
        'PurchaseOrder':'PO',
        'ProductServiceID':'ProductService ID'
        },axis=1,inplace=True)
    df_korrus_data_new['CONCAT (PN+PO)']=df_korrus_data_new['ProductService ID']+df_korrus_data_new['PO']
    df_korrus_data_new['CONCAT (PN+PO+DZ)']=df_korrus_data_new['ProductService ID']+df_korrus_data_new['PO']+df_korrus_data_new['AssignedDropZone']
    edi_keys=['PO','LineNumber','ProductService ID','AssignedDropZone']
    df_edi=append_df_to_df(df_new=df_korrus_data_new,df_old=df_edi,table='EDI Master',keys=edi_keys)

#%%
# Shipment transactions, lo embarcado al cliente
path_ship_cust_new=get_path(state,'Shipment transactions')
if path_ship_cust_new!='Not selected':
    df_ship_cust_new=load_excel_with_header_key(path_ship_cust_new,sheet_name='Embarques from ELP',key_text='Fecha de Embarque')
    df_ship_cust_new=rename_columns(df_ship_cust_new,df_col_rel=df_col_rel,table_from='ELP Master',sheet_from='Embarques from ELP')
    df_ship_cust_new=df_ship_cust_new[df_ship_cust_new['is_shipped'].str.upper().str.contains('YES')]
    df_ship_cust_new=df_ship_cust_new[~df_ship_cust_new['po'].isna()]
    df_ship_cust_new['modelo']=df_ship_cust_new['modelo'].str.upper()
    save_df(df_ship_cust_new,output_paths['path_ship_cust'],sheet_name='Shipped to Cust',index=False)
#%%
#%%
# InventoryStage, lo que se embarco a ELP 
path_ship_elp_new=get_path(state,'InventoryStageBakup')
if path_ship_elp_new!='Not selected':
    df_ship_elp_new=read_excel(path_ship_elp_new)
    if 'ShipDate Details' in path_ship_elp_new:
        df_ship_elp_new=rename_columns(df_ship_elp_new,df_col_rel,table_from='ShipDate Details')
        df_ship_elp_new['quantity']=1 
    elif 'InventoryStageBakup' in path_ship_elp_new:
        df_ship_elp_new=rename_columns(df_ship_elp_new,df_col_rel,table_from='InventoryStageBakup')
        df_ship_elp_new=df_ship_elp_new[['family','po','modelo','box_id','quantity']].ffill()
        df_ship_elp_new['shipment_date_elp']=state["fecha_shipments_elp"]
    df_ship_elp_new['BOX qty']=1
    df_ship_elp_new=df_ship_elp_new[df_ship_elp_new['family']!='Total']
    df_ship_elp_new=df_ship_elp_new[~df_ship_elp_new['po'].isna()]
    df=df_ship_elp_new['po'].str.split('|', expand=True)
    df_ship_elp_new['dz']=''
    if len(df.columns)>1:
        df_ship_elp_new['dz']=df[1].str.strip()
        df_ship_elp_new['po']=df[0].str.strip()
    df_ship_elp_new['modelo']=df_ship_elp_new['modelo'].str.upper()
    df_ship_elp_new['shipment_date_elp']=pd.to_datetime(df_ship_elp_new['shipment_date_elp']).dt.strftime('%m/%d/%Y')
    # df_ship_elp_new=rename_columns(df_ship_elp_new,df_col_rel,table_from='InventoryStageBakup',table_to='ELP Master',sheet_to='Shipment to ELP')
    df_ship_elp_new['dz'].fillna('NULL',inplace=True)
    df_ship_elp_new.loc[df_ship_elp_new['dz']=='NA','dz']='NULL'
    df_ship_elp_new=df_ship_elp_new.groupby(['po','modelo','dz','box_id']).agg({
        'quantity': 'sum',
        'BOX qty':'count',
        'shipment_date_elp':'last'
    }).reset_index()
    if 'ShipDate Details' in path_ship_elp_new:
        df_ship_elp_new['BOX qty']=1

    df_ship_elp=rename_columns(df_ship_elp,df_col_rel,table_from='ELP Master',sheet_from='Shipment to ELP')
    #%%
    df_ship_elp['dz'].fillna('NULL',inplace=True)
    df_ship_elp=append_df_to_df(df_new=df_ship_elp_new,df_old=df_ship_elp,table='Shipment to ELP',keys=['po','modelo','box_id','dz'])
    df_ship_elp['family'].fillna('',inplace=True)
    df_ship_elp['shipment_date_elp']=pd.to_datetime(df_ship_elp['shipment_date_elp'], errors='coerce').dt.strftime('%m/%d/%Y')
    df_ship_elp=set_family(df_ship_elp,column='modelo',dest_col='family')
    df_ship_elp=rename_columns(df_ship_elp,df_col_rel=df_col_rel,table_to='ELP Master',sheet_to='Shipment to ELP')
#%%
# Ordenes Canceladas   

df_cancelled=read_excel(path_ship_elp,sheet_name='Cancelled Orders')
df_cancelled=rename_columns(df_cancelled,df_col_rel,table_from='ELP Master',sheet_from='Cancelled Orders',table_to='ELP Master',sheet_to='EDI Master')
df_cancelled=df_cancelled[['PO','ProductService ID','LineNumber']].drop_duplicates()
df_cancelled['ProductService ID']=df_cancelled['ProductService ID'].str.upper()
df_cancelled['status_cancelled']=True
#%%

df_edi=df_edi.merge(df_cancelled,how='left',on=['PO','ProductService ID','LineNumber'])
df_edi['Order/Line cancelled?']=''
df_edi.loc[df_edi['status_cancelled']==True,'Order/Line cancelled?']='Cancelled'
df_edi.drop(columns='status_cancelled',inplace=True)
ws_dict_edi['df']=df_edi
ws_edi=append_to_sheet(ws_dict_edi,ws_edi)
ws_edi=update_column(ws_dict_edi,ws_edi,column='Order/Line cancelled?')
ws_edi=update_column(ws_dict_edi,ws_edi,column='EDI Received')
ws_dict_ship_elp['df']=df_ship_elp
#%%

ws_ship_elp=append_to_sheet(ws_dict_ship_elp,ws_ship_elp)
date_cols=df_columns[(df_columns['sheet']=='EDI Master')&(df_columns['data_type']=='date')]['column_name'].to_list()
wb_elp=format_xl_dates(wb_elp,sheet_name='EDI Master',date_columns=date_cols)
date_cols=df_columns[(df_columns['sheet']=='Shipment to ELP')&(df_columns['data_type']=='date')]['column_name'].to_list()
wb_elp=format_xl_dates(wb_elp,sheet_name='Shipment to ELP',date_columns=date_cols)
#%%
save_wb(wb_elp,path_ship_elp)
#%% Reload from files
path_oor_old=get_path(state,'OOR')
close_xl_if_open(path_oor_old)
path_ship_elp=get_path(state,'ELP Master')
close_xl_if_open(path_ship_elp)
df_edi=read_excel(path_ship_elp,sheet_name='EDI Master')

df_ship_elp=read_excel(path_ship_elp,sheet_name='Shipment to ELP')
df_ship_elp['CUU ship Date']=pd.to_datetime(df_ship_elp['CUU ship Date'],errors='coerce').dt.strftime('%m/%d/%Y')
#%%
# Reporte de work orders, que se encuentra en proceso de produccion
path_tracker=get_path(state,'Tracker')
xls = pd.ExcelFile(path_tracker)
wo_sheets=[sheet for sheet in xls.sheet_names if (('Plan de produccion' in sheet) or ('TERMINADAS' in sheet))]
df_wo=pd.DataFrame()
for sheet in wo_sheets:
    df_wo=pd.concat([df_wo,pd.read_excel(path_tracker,sheet_name=sheet)])
    df_wo.dropna(subset=['PO cliente','Modelo'],inplace=True)
check_mandatory_cols(df_wo.columns,'Tracker')
df_wo=rename_columns(df_wo,df_col_rel,table_from='Tracker',sheet_from='Plan de produccion')
df_wo['modelo']=df_wo['modelo'].str.upper()
df_wo=format_dates(df_wo,['Date','START DATE', 'FINISH DATE', 'reprogrammed_cuu','SHIP DATE'])
df_wo.rename(columns={"wo_qty":"quantity"},inplace=True)
df_wo=move_columns_to_front(df_wo,['po','modelo','wo','quantity','START DATE', 'FINISH DATE', 'reprogrammed_cuu','estimated_move_date_cuu'])
#%%
df_wo_wb=pd.read_excel(path_tracker,sheet_name=None)
df_wo=pd.DataFrame()
for sheet in df_wo_wb.keys():
    if ('Plan de produccion' in sheet) or ('TERMINADAS' in sheet) :
        df_wo=pd.concat([df_wo,df_wo_wb[sheet]])
        df_wo.dropna(subset=['PO cliente','Modelo'],inplace=True)
check_mandatory_cols(df_wo.columns,'Tracker')
df_wo=rename_columns(df_wo,df_col_rel,table_from='Tracker',sheet_from='Plan de produccion')
df_wo['modelo']=df_wo['modelo'].str.upper()
df_wo=format_dates(df_wo,['Date','START DATE', 'FINISH DATE', 'reprogrammed_cuu','SHIP DATE'])
df_wo.rename(columns={"wo_qty":"quantity"},inplace=True)
df_wo=move_columns_to_front(df_wo,['po','modelo','wo','quantity','START DATE', 'FINISH DATE', 'reprogrammed_cuu','estimated_move_date_cuu'])
#%%
#% Edi
df_edi=rename_columns(df_edi,df_col_rel,table_from='ELP Master',sheet_from='EDI Master')
df_edi['modelo']=df_edi['modelo'].str.upper()
#%%
# Procesar envios al cliente eliminando cantidades negativas. Se conserva la fecha mas nueva de envio.
if not os.path.exists(output_paths['path_ship_cust']):
    st.info('No hay envios al paso, integre al menos un Shipment Transactions')
    raise SystemExit()
df_ship_cust=read_excel(path=output_paths['path_ship_cust'])
df_ship_cust=rename_columns(df_ship_cust,df_col_rel,table_from='Shipment transactions')
df_ship_cust['modelo']=df_ship_cust['modelo'].str.upper()
df_ship_cust_grp=df_ship_cust.groupby(['po','modelo']).sum('Qty. Shipped').reset_index()
df_ship_cust_grp=df_ship_cust_grp[['po','modelo','quantity']]
df_ship_cust_dates=df_ship_cust.sort_values(['po','modelo','shipment_date_cust']).drop_duplicates(subset=['po','modelo'],keep='last')
df_ship_cust_dates.drop('quantity',axis=1,inplace=True)
df_ship_cust_dates=df_ship_cust_dates.merge(df_ship_cust_grp,on=['po','modelo'])
df_ship_cust_dates=df_ship_cust_dates[['po','modelo','quantity','shipment_date_cust']]
#%%
# Envios a ELP
df_ship_elp=rename_columns(df_ship_elp,df_col_rel,table_from='ELP Master',sheet_from='Shipment to ELP')
df_ship_elp['modelo']=df_ship_elp['modelo'].str.upper()
df_ship_elp['po']=df_ship_elp['po'].str.strip()
df_ship_elp['modelo']=df_ship_elp['modelo'].str.strip()
df_ship_elp['quantity'].fillna(0,inplace=True)
# df_ship_elp_grp=df_ship_elp.groupby(['po','modelo','shipment_date_elp']).sum('quantity').reset_index()
# df_ship_elp_grp=df_ship_elp_grp[['po','modelo','quantity','shipment_date_elp']]

#%%
# Merge EDI con work orders y embarques
df_edi=df_edi[df_edi['po_date']>pd.to_datetime(state["fecha_freeze"])]
df_edi_po_dates=df_edi[['po','modelo','po_date']].copy()
df_edi_po_dates.sort_values(['po_date'],inplace=True)
df_edi_po_dates.drop_duplicates(['po','modelo'],keep='last',inplace=True)
df_wo=df_wo.merge(df_edi_po_dates,how='left',on=['po','modelo'])
df_wo=df_wo[df_wo['po_date']>pd.to_datetime(state["fecha_freeze"])]
df_ship_cust_dates=df_ship_cust_dates.merge(df_edi_po_dates,how='left',on=['po','modelo'])
df_ship_cust_dates=df_ship_cust_dates[df_ship_cust_dates['po_date']>pd.to_datetime(state["fecha_freeze"])]
df_ship_elp=df_ship_elp.merge(df_edi_po_dates,how='left',on=['po','modelo'])
df_ship_elp=df_ship_elp[df_ship_elp['po_date']>pd.to_datetime(state["fecha_freeze"])]

#%%
assignments_wo=assign_quantities(df_pos=df_edi,df_to_assign=df_wo,additional_fields=['WO','START DATE','FINISH DATE','reprogrammed_cuu','estimated_move_date_cuu'])
df_edi_combined=assignments_wo['df_pos']
df_edi_combined.rename({'Assigned':'WO Qty'},axis=1,inplace=True)
assignments_shp_cust=assign_quantities(df_pos=df_edi_combined,df_to_assign=df_ship_cust_dates,additional_fields=['shipment_date_cust'])
df_edi_combined=assignments_shp_cust['df_pos']
df_edi_combined.rename({'Assigned':'Shipped to Cust'},axis=1,inplace=True)
assignments_shp_elp=assign_quantities(df_pos=df_edi_combined,df_to_assign=df_ship_elp,additional_fields=['shipment_date_elp'])
df_edi_combined=assignments_shp_elp['df_pos']
df_edi_combined.rename({'Assigned':'Shipped to Elp'},axis=1,inplace=True)

#%%
# ### OOR Report


df_assigned_wo = assignments_wo['df_assignments']
df_edi_combined=merge_additional_fields(df_assigned_wo,
                           df_edi_combined,
                           fields=['po','modelo','LineNumber', 'START DATE', 'FINISH DATE','estimated_move_date_cuu','reprogrammed_cuu'],
                           sort_fields=['estimated_move_date_cuu'],
                           key=['po','modelo','LineNumber'])
# Add WO field
df_wo_qty=df_assigned_wo[['po','modelo','LineNumber','WO','AvailQuantity']]
df_wo_qty['AvailQuantity']=df_wo_qty['AvailQuantity'].fillna(0)
df_wo_qty['WO']=df_wo_qty['WO'].astype(int).astype(str) + ' TotQty: ' + df_wo_qty['AvailQuantity'].astype(int).astype(str) + '.'
df_wo_qty=df_wo_qty.groupby(['po','modelo','LineNumber']).agg({
    'WO': ' '.join
}).reset_index()
df_edi_combined=merge_additional_fields(df_wo_qty,
                           df_edi_combined,
                           fields=['po','modelo','LineNumber', 'WO'],
                           sort_fields=['po'],
                           key=['po','modelo','LineNumber'])


df_assigned_shp_elp=assignments_shp_elp['df_assignments']
df_edi_combined=merge_additional_fields(df_assigned_shp_elp,
                           df_edi_combined,
                           fields=['po','modelo','LineNumber','shipment_date_elp'],
                           sort_fields=['shipment_date_elp'],
                           key=['po','modelo','LineNumber'])
df_assigned_shp_cust=assignments_shp_cust['df_assignments']
df_edi_combined=merge_additional_fields(df_assigned_shp_cust,
                           df_edi_combined,
                           fields=['po','modelo','LineNumber','shipment_date_cust'],
                           sort_fields=['shipment_date_cust'],
                           key=['po','modelo','LineNumber'])

#%%
# Set production, and shipped STATUS
df_edi_combined[['quantity','WO Qty','Shipped to Elp','Shipped to Cust']].fillna(0,inplace=True)
df_edi_combined['quantity']=df_edi_combined['quantity'].astype(float)
df_edi_combined['Shipped to Elp']=df_edi_combined['Shipped to Elp'].astype(float)
df_edi_combined['Shipped to Cust']=df_edi_combined['Shipped to Cust'].astype(float)
df_edi_combined.loc[(df_edi_combined['quantity']-df_edi_combined['Shipped to Elp']<=0),'Movement Status']='Moved'
df_edi_combined.loc[(df_edi_combined['quantity']-df_edi_combined['Shipped to Cust']<=0),'Status']='Shipped'
cols=['quantity','WO Qty','Shipped to Elp','Shipped to Cust']
df_po_status_cls=df_edi_combined[['po']+cols]
for col in cols:
    df_po_status_cls[col]=df_po_status_cls[col].astype(float)
df_po_status_cls=df_po_status_cls.groupby(['po']).sum()
df_po_status_cls.reset_index(inplace=True)


df_po_status=df_po_status_cls[['po','quantity','WO Qty','Shipped to Elp','Shipped to Cust']]
df_edi_combined.sort_values(['po','modelo','LineNumber','po_date'],inplace=True)
df_edi_combined.reset_index(drop=True,inplace=True)
df_edi_combined['Balance to ship']=df_edi_combined['quantity']-df_edi_combined['Shipped to Elp']

df_po_status_cls=df_po_status_cls.loc[(df_po_status_cls['quantity']-df_po_status_cls['Shipped to Cust'])<=0,['po']]
df_po_dates=df_assigned_shp_cust.sort_values(['po','shipment_date_cust']).drop_duplicates('po',keep='last')[['po','shipment_date_cust']]
df_po_status_cls=df_po_status_cls.merge(df_po_dates,how='left',on=['po'])
df_po_status_cls.rename(columns={'shipment_date_cust':'\nPO closure date'},inplace=True)
df_edi_combined=df_edi_combined.merge(df_po_status_cls,how='left',on=['po'])
df_edi_combined=set_family(df_edi_combined,column='modelo',dest_col='family')
df_po_type=df_edi_combined[['po','family']].drop_duplicates()
df_po_type.loc[df_po_type['family']=='Accesories','PO TYPE']='Accesories only'
df_po_type.loc[(~(df_po_type['family']=='Accesories'),'PO TYPE')]='Fixtures only'
df_po_type=df_po_type.drop_duplicates(['po','PO TYPE']).sort_values(['po','PO TYPE'])
df_po_type=df_po_type.groupby('po', as_index=False).agg({'PO TYPE': '/'.join})
df_po_type.loc[df_po_type['PO TYPE']=='Accesories only/Fixtures only','PO TYPE']='Both types'
df_edi_combined=df_edi_combined.merge(df_po_type,how='left',on='po')

df_edi_combined.loc[df_edi_combined['Order/Line cancelled?']=='Cancelled','Status']='Cancelled'

if not 'edi_received' in df_edi_combined.columns:
    df_edi_combined['edi_received']=''

#%%
#%%
# Read customer shipments and shipments to elp
df_ship_cust=read_excel(output_paths['path_ship_cust'])
df_ship_cust=rename_columns(df_ship_cust,df_col_rel,table_from='Shipment transactions')
df_ship_cust['modelo']=df_ship_cust['modelo'].str.upper()
df_ship_cust.sort_values(['po','modelo','shipment_date_cust'],inplace=True)
df_ship_cust.reset_index(drop=True,inplace=True)
df_ship_cust=format_dates(df_ship_cust,['shipment_date_cust'])
df_ship_cust=move_columns_to_front(df_ship_cust,['po','modelo','shipment_date_cust','Quantity'])
df_ship_elp=rename_columns(df_ship_elp,df_col_rel,table_from='ELP Master',sheet_from='Shipment to ELP')
df_ship_elp['modelo']=df_ship_elp['modelo'].str.upper()
df_ship_elp.sort_values(['po','modelo','shipment_date_elp'],inplace=True)
df_ship_elp.reset_index(drop=True,inplace=True)
df_ship_elp.loc[df_ship_elp['dz']=='','dz']='NULL'
df_ship_elp['dz'].fillna('NULL',inplace=True)
df_ship_elp=format_dates(df_ship_elp,['df_ship_elp'],type='slash')
df_ship_elp=move_columns_to_front(df_ship_elp,['po','modelo','df_ship_elp','Quantity'])
#%%
df_ship_elp
# Get index for hyperlinks
#%%%
df_wo.reset_index(inplace=True,drop=True)
df_ship_elp.reset_index(inplace=True,drop=True)
df_ship_cust.reset_index(inplace=True,drop=True)
df_wo_idx=get_df_idx(df_wo,idx_cols=['po','modelo'],idx_name='idx_wo')
df_ship_elp_idx=get_df_idx(df_ship_elp,idx_cols=['po','modelo'],idx_name='idx_elp')
df_ship_cust_idx=get_df_idx(df_ship_cust,idx_cols=['po','modelo'],idx_name='idx_cust')

df_edi_combined=df_edi_combined.merge(df_wo_idx,how='left',on=['po','modelo'])
df_edi_combined=df_edi_combined.merge(df_ship_elp_idx,how='left',on=['po','modelo'])
df_edi_combined=df_edi_combined.merge(df_ship_cust_idx,how='left',on=['po','modelo'])

# Get Hyperlinks
df_edi_combined=set_hyperlink(df_edi_combined,sheet_name='WorkOrder Detail',col_name='WO',idx_name='idx_wo')
df_edi_combined=set_hyperlink(df_edi_combined,sheet_name='Shipped to Cust',col_name='Shipped to Cust',idx_name='idx_cust',typ='int')
df_edi_combined=set_hyperlink(df_edi_combined,sheet_name='Shipped to Elp',col_name='Shipped to Elp',idx_name='idx_elp',typ='int')

df_edi_combined['WO'].fillna('',inplace=True)
#%% 
df_edi_combined[df_edi_combined['shipment_date_elp']=='04/25/2025']
#%%
# Get prices if selected
# path_prices=get_path(state,'Prices')
# df_prices=pd.DataFrame()
# if path_prices!='Not selected':
#     df_prices=read_excel(path_prices)
#     df_prices=rename_columns(df_prices,df_col_rel,table_from='Prices',table_to='OOR Report',sheet_to='OOR')
#     df_prices=df_prices[['ProductServiceID','Price']]
#     df_prices.drop_duplicates(['ProductServiceID'],keep='last',inplace=True)

#%%
# ### Actualizar OOR con datos nuevos

# Actualizar OOR
df_oor=rename_columns(df_edi_combined,df_col_rel,table_from='EDI Combined',table_to='OOR Report',sheet_to='OOR')

df_oor=set_family(df_oor,'ProductServiceID')
oor_cols=df_columns[df_columns['table']=='OOR Report']['column_name'].to_list()
# Calculo de valores por columna
check_duplicated_columns(df_oor)
df_oor=format_dates(df_oor,['START DATE','FINISH DATE','SHIP DATE','EDI Received','POUS Date','Actual Ship date to Customer\n','Actual Move date from CUU'],type='')
df_oor.loc[df_oor['AssignedDropZone']=='','AssignedDropZone']='NULL'
df_oor['AssignedDropZone'].fillna('NULL',inplace=True)
df_oor['NP+PO+DZ']=df_oor['ProductServiceID']+df_oor['PurchaseOrder']+df_oor['AssignedDropZone'].astype(str)
df_oor['CONCAT\nPN+PO']=df_oor['ProductServiceID']+df_oor['PurchaseOrder']

oor_current_cols=[x for x in oor_cols if x in df_oor]

df_oor=df_oor[oor_current_cols]
df_oor.sort_values(['Family ','PurchaseOrder','ProductServiceID','LineNumber'],inplace=True)
df_oor['PO Aging']='' # Formula
df_oor['Aging Category']='' # Formula
df_oor['Balance to move']='' # Formula
# Filtrar segun la seleccion de fecha
df_oor=df_oor[df_oor['POUS Date'] >= pd.to_datetime(state["fecha_freeze"])]

#Integrar datos al OOR Anterior
path_oor_old=get_path(state,'OOR')

close_xl_if_open(path_oor_old)
sheet_rel=extract_selected_sheets(path_oor_old,sheets_to_keep,keep_original=False)
wb_oor_old=load_workbook(path_oor_old)

ws_oor_old=wb_oor_old['OOR']
dict_oor_old=get_worksheet_df(ws_oor_old,'Family')
df_oor_old=dict_oor_old['df']
check_mandatory_cols(df_oor_old.columns,'OOR')
# Part of the OOR that will be updated, we will remove duplicates from this part
key_cols=['PurchaseOrder','ProductServiceID','LineNumber','AssignedDropZone']
df_edi_rec_dates=df_oor_old[key_cols+['EDI Received']].drop_duplicates(key_cols).copy()

df_oor_to_update=df_oor_old[df_oor_old['POUS Date'] >= pd.to_datetime(state["fecha_freeze"])]
# This part will be kept as is, duplicates are allowed
df_oor_old=df_oor_old[df_oor_old['POUS Date'] < pd.to_datetime(state["fecha_freeze"])]
except_columns=df_columns[(~df_columns['user_column'].isna())&(df_columns['sheet']=='OOR')]['column_name'].to_list()
df_oor_to_update.drop_duplicates(subset=key_cols,inplace=True)
df_oor[['Comment','Status']]=df_oor[['Comment','Status']].fillna('')
df_oor_to_update=update_dataframe(df_oor_to_update,df_oor.fillna(0),key_cols,exceptions=except_columns)
df_oor_old=pd.concat([df_oor_old,df_oor_to_update])
df_oor_old.reset_index(drop=True,inplace=True)
#%%

path_oh_max=get_path(state,'OH Max')
df_oh_max=read_excel(path_oh_max,header=None)

df_oh_max['oh_max']
#%%
if path_oh_max!="Not selected":
    # Se actualiza el OH Max para todo el workbook ya que la formula toma en cuenta los embarcados y duplicados
    df_oh_max=read_excel(path_oh_max)

    if df_oh_max.shape[1]!=6:
        print("Numero incorrecto de columnas en OH Max")
    df_oh_max.columns=['modelo','description','oh_max','uom','stock_id','zone']
    df_oh_max = df_oh_max[pd.to_numeric(df_oh_max['oh_max'], errors='coerce').notna()]
    df_oh_max['oh_max']=pd.to_numeric(df_oh_max['oh_max'])
    df_oh_max=df_oh_max[~df_oh_max['stock_id'].str.upper().str.strip().isin(['KITS','PURGE','OUT','PP'])]
    df_oh_max=rename_columns(df_oh_max,df_col_rel,table_from='OH Max',sheet_from='only',table_to='OOR Report',sheet_to='OOR')
    df_oh_max=df_oh_max.groupby(['ProductServiceID']).sum('OH MAX')
    df_oh_max.reset_index(inplace=True)
    if 'OH MAX' in df_oor_old.columns:
        df_oor_old.drop(columns='OH MAX',inplace=True)
    df_oor_old=df_oor_old.merge(df_oh_max,how='left',on='ProductServiceID')
    if not 'OH MAX' in df_oor_old.columns:
        df_oor_old['OH MAX']=0    
    df_oor_old['OH MAX'].fillna(0,inplace=True)

df_oor_old=df_oor_old.merge(df_edi_rec_dates,how='left',on=key_cols,suffixes=('', '_new'))
df_oor_old.loc[df_oor_old['EDI Received'].isna(),'EDI Received']=df_oor_old.loc[df_oor_old['EDI Received'].isna(),'EDI Received_new']
df_oor_old.loc[df_oor_old['EDI Received']==0,'EDI Received']=df_oor_old.loc[df_oor_old['EDI Received']==0,'EDI Received_new']
df_oor_old.drop(columns='EDI Received_new',inplace=True)


#%%
if 'df_prices' not in st.session_state:
    print('Lista de precios no disponible')    
else:
    # if 'Price' in df_oor_old.columns:
    #     df_oor_old.drop(columns=['Price'],inplace=True)
    print("Integrando precios")
    df_prices=st.session_state.df_prices
    df_prices=rename_columns(df_prices,df_col_rel=df_col_rel,table_to='OOR Report',sheet_to='OOR')
    df_oor_old=df_oor_old.merge(df_prices,how='left',on=['ProductServiceID'])
#%%
df_oor_old
#%%

ws_oor_old=update_sheet(dict_oor_old,ws_oor_old)
#%%
# Value A1=1 is just to check later if formulas are evaluated
ws_oor_old['A1'].value="=1"

# Put date in oh max column

cell=find_cell_by_text(ws_oor_old,'OH MAX')
if cell:
    cell=ws_oor_old[cell]
    cell.offset(-1,0).value=datetime.now().date()

# Formato excel

wb_oor={}
wb_oor_old=create_new_sheet(wb_oor_old,sheet_name='WorkOrder Detail',df=df_wo)
wb_oor_old=create_new_sheet(wb_oor_old,sheet_name='Shipped to Elp',df=df_ship_elp)
wb_oor_old=create_new_sheet(wb_oor_old,sheet_name='Shipped to Cust',df=df_ship_cust)


#Apply autofilter and freeze panes
wb_oor_old['WorkOrder Detail'].auto_filter.ref = wb_oor_old['WorkOrder Detail'].dimensions
wb_oor_old['WorkOrder Detail'].freeze_panes = 'A2'
wb_oor_old['Shipped to Elp'].auto_filter.ref = wb_oor_old['Shipped to Elp'].dimensions
wb_oor_old['Shipped to Elp'].freeze_panes = 'A2'
wb_oor_old['Shipped to Cust'].auto_filter.ref = wb_oor_old['Shipped to Cust'].dimensions
wb_oor_old['Shipped to Cust'].freeze_panes = 'A2'


# Format row when key values change
dict_special_formats=get_xl_formatting('special_format')
col_info1=get_column_info(ws_oor_old,'PurchaseOrder')['data_range']
col_info2=get_column_info(ws_oor_old,'ProductServiceID')['data_range']
format_on_change(zip(col_info1,col_info2),ws_oor_old,start_row=1,format1=dict_special_formats['on_change_a'],format2=dict_special_formats['on_change_b'])

# Format dates
wb_oor_old=format_xl_dates(wb_oor_old,sheet_name='OOR',date_columns=['POUS Date',
                                                     'EDI Received',
                                                     '\nPO closure date',
                                                     'Actual Move date from CUU',
                                                     'Actual Ship date to Customer\n',
                                                     'Estimated \nMOVE DATE\nCUU',
                                                     'Reprogrammed finish date CUU',
                                                     'START DATE',
                                                     'FINISH DATE',
                                                     'Actual Ship date to Customer\n'])



ws=wb_oor_old['WorkOrder Detail']
col_info1=get_column_info(ws,'po')['data_range']
col_info2=get_column_info(ws,'modelo')['data_range']
format_on_change(zip(col_info1,col_info2),ws,start_row=1,format1=dict_special_formats['on_change_a'],format2=dict_special_formats['on_change_b'])


ws=wb_oor_old['Shipped to Elp']
col_info1=get_column_info(ws,'po')['data_range']
col_info2=get_column_info(ws,'modelo')['data_range']
format_on_change(zip(col_info1,col_info2),ws,start_row=1,format1=dict_special_formats['on_change_a'],format2=dict_special_formats['on_change_b'])

ws=wb_oor_old['Shipped to Cust']
col_info1=get_column_info(ws,'po')['data_range']
col_info2=get_column_info(ws,'modelo')['data_range']
format_on_change(zip(col_info1,col_info2),ws,start_row=1,format1=dict_special_formats['on_change_a'],format2=dict_special_formats['on_change_b'])

save_wb(wb_oor_old,path_oor_old)
os.startfile(path_oor_old)



#%% # AAR

# %% jupyter={"source_hidden": true}
#Yield reports
path_oor_old=get_path(state,'OOR')
folder_yield=state.get('folder_yield', "Not selected")

if folder_yield=="Not selected":
    st.info(f'Favor de seleccionar los folder con Yield reports')
    raise SystemExit()
yield_files_lst=os.listdir(folder_yield)
df_yield=pd.DataFrame()
for file in yield_files_lst:
    df=read_excel(os.path.join(folder_yield,file))
    df=map_yield_report(df)
    df['file_name']=file
    df_yield=pd.concat([df,df_yield])


dict_xl_fields={"ENSAMBLE DE BISAGRA Y TAPAS":"Complete mechanical Assy output",
 "ANGLE/DIMMING TEST":"Angle/Dimming test output",
#  "DIMMING TEST":"Angle/Dimming test output",
 "ENSAMBLE DEL BRACKET":"Mechanical Assy"}

for process in df_yield['Process'].drop_duplicates():
    if not process in dict_xl_fields.keys():
        continue
    df_yield.loc[df_yield['Process']==process,'xl_field']=dict_xl_fields[process]


df_yield=df_yield.groupby(['date_from','xl_field']).sum('Total Tested').reset_index()
df_yield=df_yield[['date_from','xl_field','Total Tested']]
df_yield['date_from']=pd.to_datetime(df_yield['date_from']).dt.normalize()
# Llenar el formato
search_parms={
    "column":'date_from',
    "row":'xl_field',
    "offset":
            {"Category":
            {
            "scheduled":(0,0),
            "real":(1,0)
            },
            "Shift":
            {"1s":(0,0),
            "2s":(0,2)},
            }
}

wb=load_workbook(path_oor_old)
wb=fill_yield_report(df_yield,wb,sheet_name='Trov Daily Status',search_parms=search_parms)
wb=fill_yield_report(df_yield,wb,sheet_name='Rise Daily Status',search_parms=search_parms)
wb.save(path_oor_old)
os.startfile(path_oor_old)

#%%
# # Integrar Gating parts report
#

# %%
# Gating parts
dict_special_formats=get_xl_formatting('special_format')
format_more_than_one_value=dict_special_formats['more_than_one_value']
neutral_format=dict_special_formats['neutral_format']
path_gating=get_path(state,'Gating Parts')
if path_gating=="Not selected":
    st.info("Seleccionar reporte de Gating Parts")
    raise SystemExit()
dict_gating=read_excel(path_gating,sheet_name=None)
path_oor=get_path(state,'OOR')
close_xl_if_open(path_oor)
wb_oor=load_workbook(path_oor)

ws_oor=wb_oor['OOR']
dict_oor=get_worksheet_df(ws_oor,'Family',data_only=True)
df_oor=rename_columns(dict_oor['df'],df_col_rel,table_from='OOR Report',sheet_from='OOR')

# Ficture Gating part and ETA
df_ready=dict_gating['Ready'].drop_duplicates()
df_ready['rdy']='CTB'
df_ready=rename_columns(df_ready,df_col_rel,table_from='Gating Parts',sheet_from='Ready')

if 'rdy' in df_oor.columns:
    df_oor.drop(columns=['rdy'],inplace=True)
df_oor=df_oor.merge(df_ready[['po','modelo','rdy']],how='left',on=['po','modelo'])
df_oor.loc[df_oor['rdy']=='CTB','fixture_gating_part']='CTB'
df_short_detail=rename_columns(dict_gating['Shorts'],df_col_rel,table_from='Gating Parts',sheet_from='Shorts')
df_short_detail.sort_values(['po','modelo','gating_due'],inplace=True)

df_short=df_short_detail.drop_duplicates(['po','modelo'],keep='last').copy()
df_short=df_short[['po','modelo','component','gating_due']]
# More than one short gating part for a line, we keep the last line to keep the last date
df_short=df_short_detail.copy()
df_short=df_short[['po','modelo','component','gating_due']]
df_short['plus_one']=False
df_short.loc[df_short.duplicated(['po','modelo'],keep=False),'plus_one']=True
df_short=df_short.drop_duplicates(['po','modelo'],keep='last')

df_oor=df_oor.merge(df_short,how='left',on=['po','modelo'])
df_oor['fixture_gating_part']=''
df_oor.loc[~df_oor['component'].isnull(),'fixture_gating_part']=df_oor.loc[~df_oor['component'].isnull(),'component']
df_oor.loc[~df_oor['component'].isnull(),'fixt_gp_eta']=df_oor.loc[~df_oor['component'].isnull(),'gating_due']

# Short of fixture parts, not accessories
df_short_detail.sort_values(['po','modelo','component','short_qty','gating_due','arrival_qty'],inplace=True)
df_short_detail.reset_index(inplace=True,drop=True)
create_new_sheet(wb_oor,sheet_name='Short Detail',df=df_short_detail)
df_idx_short=get_df_idx(df_short_detail,idx_cols=['po','modelo'],idx_name='idx_short')
df_oor=df_oor.merge(df_idx_short,how='left',on=['po','modelo'])
df_oor=set_hyperlink(df_oor,sheet_name='Short Detail',col_name='fixture_gating_part',idx_name='idx_short')
df_arrivals=dict_gating['Arrivals']
df_arrivals=rename_columns(df_arrivals,df_col_rel,table_from='Gating Parts',sheet_from='Arrivals')
df_arrivals.sort_values(['modelo','arrival_due'], inplace=True)
df_arrivals['Cumulative Sum'] = df_arrivals.groupby(['modelo'])['arrival_qty'].cumsum()

df_acc_shorts=df_oor[(df_oor['accessory_gating_part']=='Acc Short')&
                     (df_oor['oor_status'].str.lower()!='cancelled')&
                     (df_oor['family'].str[0:5].str.lower()=='acces')]
df_acc_shorts['Cumulative Sum'] = df_acc_shorts.groupby(['modelo'])['quantity'].cumsum()
df_arrivals['Cumulative Sum']=df_arrivals['Cumulative Sum'].astype(int)
df_acc_shorts['Cumulative Sum']=df_acc_shorts['Cumulative Sum'].astype(int)

df_acc_shorts = pd.merge_asof(
    df_acc_shorts.sort_values('Cumulative Sum'),
    df_arrivals.sort_values('Cumulative Sum'),
    left_on='Cumulative Sum',
    right_on='Cumulative Sum',
    by='modelo',
    direction='forward'
)
df_acc_shorts.dropna(subset=['po','modelo','LineNumber','arrival_due','quantity','arrival_qty'],inplace=True)
df_acc_shorts=df_acc_shorts[['po','modelo','LineNumber','arrival_due','quantity','arrival_qty']]
df_acc_shorts.drop_duplicates(subset=['po','modelo','LineNumber'],inplace=True,keep='last')
if 'arrival_due' in df_oor.columns:
    df_oor.drop(columns=['arrival_due'])
df_oor['acc_gp_eta']=''
df_oor=df_oor.merge(df_acc_shorts[['po','modelo','LineNumber','arrival_due']],how='left',on=['po','modelo','LineNumber'])

ws_oor=wb_oor['OOR']
for index,row in df_oor[['fixture_gating_part','fixture_gating_part_cell','fixt_gp_eta','fixt_gp_eta_cell','plus_one','acc_gp_eta_cell','arrival_due']].iterrows():
    cell=ws_oor[row['fixture_gating_part_cell']]
    cell.value=row['fixture_gating_part']
    if row['plus_one']==True:    
        format_cell(cell,format_more_than_one_value)
    else:
        format_cell(cell,neutral_format)
    cell=ws_oor[row['fixt_gp_eta_cell']]
    cell.value=row['fixt_gp_eta']
    cell.number_format = 'mm/dd/yyyy'
    cell=ws_oor[row['acc_gp_eta_cell']]
    cell.value=row['arrival_due']
# Value A1=1 is just to check later if formulas are evaluated
ws_oor['A1'].value="=1"
save_wb(wb_oor,path_oor)
#%%
# Actualizar Status

# Generar grupos de filtros por PO
# Columnas tipo texto se hace un .join ordenado
# Columnas tipo fecha se toma la ultima fecha
# Columnas numericas se suman

path_oor=get_path(state,'OOR')
sheet_rel=extract_selected_sheets(path_oor,sheets_to_keep,keep_original=False)
# df_oor=load_excel_with_header_key(path_oor,sheet_name='OOR',key_text='Family ', dtype=str)
wb_oor=load_workbook(path_oor,data_only=True)
ws_oor=wb_oor['OOR']
if ws_oor['A1'].value is None:
    st.info(f"Favor de Guardar el archivo {path_oor}")
    raise SystemExit()

dict_oor=get_worksheet_df(ws_oor,'Family ',data_only=True)
df_oor=dict_oor['df']
df_oor_stat=rename_columns(df_oor,df_col_rel,table_from='OOR Report',sheet_from='OOR')
df_oor_stat=df_oor_stat.copy()

# Tipo de filtros
df_oor_cols=df_columns[df_columns['sheet']=='OOR']

all_filters=df_oor_cols.loc[(~df_oor_cols['filter'].isna()),'std_name'].drop_duplicates().to_list()

# Tipo has_value
has_val_cols=df_oor_cols.loc[(df_oor_cols['filter']=='has_value'),'std_name'].drop_duplicates().to_list()
for col in has_val_cols:
    df_oor_stat.loc[(df_oor_stat[col]=='') | (df_oor_stat[col].isna()),col]='blank'
    df_oor_stat.loc[(df_oor_stat[col]!='') & (~df_oor_stat[col].isna()) & (df_oor_stat[col]!='blank'),col]='has_value'


# Tipo lista:
list_cols=df_oor_cols.loc[(df_oor_cols['filter']=='list'),'std_name'].drop_duplicates().to_list()
for col in list_cols:
    df_oor_stat.loc[(df_oor_stat[col]=='') | (df_oor_stat[col].isna()),col]='blank'
    df_oor_stat[col]=df_oor_stat[col].astype(str).str.lower()
list_cols=list_cols+has_val_cols

# Tipo fecha
date_cols=df_oor_cols.loc[(df_oor_cols['filter']=='date'),'std_name'].drop_duplicates().to_list()
for col in date_cols:
    df_oor_stat[col] = pd.to_datetime(df_oor_stat[col], errors='coerce')
    df_oor_stat[col]=df_oor_stat[col].fillna(pd.Timestamp('2099-12-31'))
date_cols_first=[f"{x}_first" for x in date_cols]
date_cols_last=[f"{x}_last" for x in date_cols]

# Tipo cantidad
qty_cols=df_oor_cols.loc[(df_oor_cols['filter']=='qty'),'std_name'].drop_duplicates().to_list()


# Key
key_cols=['po']

# Grupos de propiedades
# 1. Copiar original df_oor_stat
# 2. Aplicar filtros
# 3. Agregar fechas first, last
# 4. Aggregate
# 4. Merge

# Agregar valor in plan, existe como columna ficticia
df_oor_stat['in_plan']='blank'
df_oor_stat.loc[(df_oor_stat['balance_to_move']==0)|(df_oor_stat['wo']=='has_value'),'in_plan']='yes'
if not 'in_plan' in list_cols:
    list_cols.append('in_plan')

# Agregar valor short, existe como columna ficticia
df_oor_stat['acc_short']='blank'
df_oor_stat.loc[(df_oor_stat['acc_gp_eta']>pd.to_datetime(date.today()+timedelta(days=15)))&(df_oor_stat['accessory_gating_part']=='acc short'),'acc_short']='yes'
if not 'acc_short' in list_cols:
    list_cols.append('acc_short')

## Propiedades de PO (Status diferente de Shipped o Cancelled)
df_oor_by_po_active=df_oor_stat.copy()
df_oor_by_po_active=df_oor_by_po_active[~df_oor_by_po_active['oor_status'].str.lower().str.strip().isin(["shipped","cancelled"])]
df_oor_by_po_active=apply_grouping(df_oor_by_po_active,prefix='po_active_all_',key_cols=key_cols,list_cols=list_cols,qty_cols=qty_cols,date_cols=date_cols)
# Propiedades de accesorios
df_oor_po_acc_active=df_oor_stat.copy()
df_oor_po_acc_active=df_oor_po_acc_active[(~df_oor_po_acc_active['oor_status'].str.lower().str.strip().isin(["shipped","cancelled"]))
                            & (df_oor_po_acc_active['family'].str[0:5].str.lower()=='acces')]
df_oor_po_acc_active=apply_grouping(df_oor_po_acc_active,prefix='po_active_acc_',key_cols=key_cols,list_cols=list_cols,qty_cols=qty_cols,date_cols=date_cols)
# Propiedades de fixturas
df_oor_po_fix_active=df_oor_stat.copy()
df_oor_po_fix_active=df_oor_po_fix_active[(~df_oor_po_fix_active['oor_status'].str.lower().str.strip().isin(["shipped","cancelled"]))
                            & (df_oor_po_fix_active['family'].str[0:5].str.lower()!='acces')]
df_oor_po_fix_active=apply_grouping(df_oor_po_fix_active,prefix='po_active_fix_',key_cols=key_cols,list_cols=list_cols,qty_cols=qty_cols,date_cols=date_cols)

# Merge
df_oor_stat=df_oor_stat.merge(df_oor_by_po_active.df,how='left',on=key_cols)
df_oor_stat=df_oor_stat.merge(df_oor_po_acc_active.df,how='left',on=key_cols)
df_oor_stat=df_oor_stat.merge(df_oor_po_fix_active.df,how='left',on=key_cols)


# Actualizar listas en columns and formatting.xlsx
col_options={}
for col in list_cols + df_oor_by_po_active.list_cols + df_oor_po_acc_active.list_cols + df_oor_po_fix_active.list_cols:
    col_options[col]=df_oor_stat[col].dropna().drop_duplicates().to_list()
    if 'blank' not in col_options[col]:
        col_options[col]=col_options[col]+['blank']
# Quantity and date
for col in qty_cols + df_oor_by_po_active.qty_cols + df_oor_po_acc_active.qty_cols + df_oor_po_fix_active.qty_cols:
    col_options[col]=["(quantity)"]
for col in date_cols + df_oor_by_po_active.date_cols + df_oor_po_acc_active.date_cols + df_oor_po_fix_active.date_cols:
    col_options[col]=["(date)","today + X days","no date"]

col_options['columns_list']=sorted(col_options.keys())
# Agregar opciones
wb_cols=load_workbook(output_paths['path_xl_format'])
ws=wb_cols['options']
dict_opt=get_worksheet_df(ws,data_only=True)
# Add options
for col,options in col_options.items():
    if col in dict_opt['df'].columns:
        df=dict_opt['df'][[col]]
        old_opt=df[~df[col].isnull()][col].astype(str).tolist()
    else:
        old_opt=[]
    new_opt=[str(option) for option in options]
    new_opt=sorted(set(old_opt+new_opt))
    df_new_opt=pd.DataFrame(columns=[col],data=new_opt)
    if col in dict_opt['header']:
        dict_opt['df']=pd.concat([dict_opt['df'].drop(columns=[col]),df_new_opt],axis=1)
    else:
        dict_opt['df']=pd.concat([dict_opt['df'],df_new_opt],axis=1)

ws=update_sheet(dict_opt,ws,apply_formats=False)
# Create defined names based on options sheet
dict_opt=get_worksheet_df(ws,data_only=True)
df=dict_opt['df']
for col in df.columns:
    if col.endswith('_cell'):
        continue
    coorda=df.loc[~df[col].isnull(),f'{col}_cell'].tolist()[0]
    coordb=df.loc[~df[col].isnull(),f'{col}_cell'].tolist()[-1]
    if col in wb_cols.defined_names:
        del(wb_cols.defined_names[col])
    wb_cols.defined_names.add(DefinedName(col, attr_text=f"options!${coorda}:${coordb}"))


# Agregar data validation para condiciones 
ws=wb_cols['conditions']

dict_condit=get_worksheet_df(ws,key_text="column",data_only=True)
if 'column_cell' in dict_condit['df'].columns:
    last_cell1=dict_condit['df']['column_cell'].iloc[-1]
    last_cell2=dict_condit['df']['options_cell'].iloc[-1]
    first_cell1=dict_condit['df']['column_cell'].iloc[0]
    first_cell2=dict_condit['df']['options_cell'].iloc[0]
else:
    last_cell1=ws[find_cell_by_text(ws,"column")].offset(1,0).coordinate
    last_cell2=ws[find_cell_by_text(ws,"options")].offset(1,0).coordinate
    first_cell1=last_cell1
    first_cell2=last_cell2

range1=ws[f"{ws[first_cell1].coordinate}:{ws[last_cell1].offset(11,0).coordinate}"]
range2=ws[f"{ws[first_cell2].coordinate}:{ws[last_cell2].offset(11,0).coordinate}"]
ws.data_validations.dataValidation=[]
options_str = '"' + ','.join(sorted(df_oor_stat.columns)) + '"'

dv1=DataValidation(type="list", formula1=f'=INDIRECT("columns_list")', allow_blank=True)
ws.add_data_validation(dv1)
dv1.add(f"{range1[0][0].coordinate}:{range1[-1][0].coordinate}")
for cols in zip(range1,range2):
    dv2=DataValidation(type="list", formula1=f"=INDIRECT({cols[0][0].coordinate})", allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add(cols[1][0])
save_wb(wb_cols,output_paths['path_xl_format'])


# Actualizar status en el OOR
df_conditions=read_excel(output_paths['path_xl_format'],sheet_name='conditions')
df_conditions=df_conditions[df_conditions['sheet']=='OOR']
df_oor_stat['status_new']=''
for col in date_cols+date_cols_first+date_cols_last:
    if not col in df_oor_stat.columns:
        continue
    df_oor_stat[col] = pd.to_datetime(df_oor_stat[col], errors='coerce')
df_oor_status_final=apply_dynamic_rules(df=df_oor_stat,rules_df=df_conditions,target_col='status_new')
df_oor_status_final=df_oor_status_final[['po','modelo','dz','LineNumber','status_new']]
df_oor_status_final.drop_duplicates(['po','modelo','dz','LineNumber'],inplace=True,keep='last')
# df_oor_status_final=rename_columns(df_oor_status_final,table_to='OOR Report',sheet_to='OOR',df_col_rel=df_col_rel)
df_oor=rename_columns(df_oor,df_col_rel,table_from='OOR Report',sheet_from='OOR')
if 'status_new' in df_oor.columns:
    df_oor.drop(columns=['status_new'],inplace=True)
df_oor=df_oor.merge(df_oor_status_final,how='left',on=['po','modelo','LineNumber','dz'])
idx=(~df_oor['oor_status'].str.lower().isin(['cancelled','shipped']))
df_oor=df_oor[idx]
df_oor['oor_status']=df_oor['status_new']


# Actualizar fecha: Last commit date
df_short_dates=df_oor.copy()
df_short_dates=df_short_dates[['po','fixt_gp_eta','acc_gp_eta','estimated_move_date_cuu']]
df_short_dates['fixt_gp_eta_7_days']=pd.to_datetime(df_short_dates['fixt_gp_eta'], errors='coerce')+7*BDay()
df_short_dates['acc_gp_eta_2_days']=pd.to_datetime(df_short_dates['acc_gp_eta'], errors='coerce')+2*BDay()
df_short_dates['estimated_move_date_cuu_2_days']=pd.to_datetime(df_short_dates['estimated_move_date_cuu'], errors='coerce')+2*BDay()
# Max dates for shortage status
df_short_dates['max_date_short']=df_short_dates[['fixt_gp_eta_7_days','acc_gp_eta_2_days']].max(axis=1)
# Max datefor In shipping plan Case2
df_short_dates['max_date_ship_c2']=df_short_dates[['estimated_move_date_cuu_2_days','acc_gp_eta_2_days']].max(axis=1)
df_short_dates.drop(columns=['fixt_gp_eta','acc_gp_eta','estimated_move_date_cuu'],inplace=True)
df_short_dates=df_short_dates.groupby('po').max().reset_index()
df_oor['latest_commit_date']=None
df_oor=df_oor.merge(df_short_dates,how='left',on=['po'])


# Update latest commit date
df_oor.loc[df_oor['oor_status']=='Shortage','latest_commit_date']=df_oor.loc[df_oor['oor_status']=='Shortage','max_date_short']
df_oor.loc[df_oor['oor_status']=='In shipping plan_c1','latest_commit_date']=df_oor.loc[df_oor['oor_status']=='In shipping plan_c1','estimated_move_date_cuu_2_days']
df_oor.loc[df_oor['oor_status']=='In shipping plan_c2','latest_commit_date']=df_oor.loc[df_oor['oor_status']=='In shipping plan_c2','max_date_ship_c2']
df_oor['oor_status']=df_oor['oor_status'].str.replace('_c1','').str.replace('_c2','')

# Update TAT Category

df_tat=df_oor_stat[['po','modelo','LineNumber','po_active_all_po_shipped_complete_category']].copy()
df_tat['tat_cat']=''
df_tat.loc[df_tat['po_active_all_po_shipped_complete_category']=='pending ship complete','tat_cat']='Open'
df_tat.loc[df_tat['po_active_all_po_shipped_complete_category']=='po ship complete','tat_cat']='Closed'
df_tat=df_tat[['po','modelo','LineNumber','tat_cat']].drop_duplicates(subset=['po','modelo','LineNumber'])
if 'tat_cat' in df_oor.columns:
    df_oor.drop(columns=['tat_cat'],inplace=True)
df_oor=df_oor.merge(df_tat,how='left',on=['po','modelo','LineNumber'])

for idx,row in df_oor[['oor_status','Status_cell','latest_commit_date','latest_commit_date_cell','tat_cat','TAT Category_cell']].iterrows():
    ws_oor[row['Status_cell']]=row['oor_status']
    ws_oor[row['latest_commit_date_cell']]=row['latest_commit_date']
    ws_oor[row['TAT Category_cell']]=row['tat_cat']

save_wb(wb_oor,path_oor)
os.startfile(path_oor)