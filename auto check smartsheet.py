# %% Load Modules
import subprocess
import pandas as pd
import smartsheet

# %%
import json
driveloc=r'C:\Users\dokim2\OneDrive - Kiss Products Inc'
desktoploc=r'C:\Users\dokim2\OneDrive - Kiss Products Inc\Desktop'
with open(desktoploc+'\IVYENT_DH\data.json', 'r') as f:
    data = json.load(f)

# %%

# Get order information from smartsheet
def check_smartsheet(desktoploc):
    smartsheet_client = smartsheet.Smartsheet('rjjjwNgTfxwAjE5R5YcSKu5OocAMyLAUJa2av')
    smartsheet_client.Sheets.get_sheet_as_csv(
  3482775064995716,           # sheet_id
  desktoploc+'\stock check practice')

    order_df = pd.read_csv(desktoploc+'\stock check practice\Stock Check Request.csv')
    order_df['Requester'] = order_df['Requester'].astype(str)

    condition = ((order_df['Company (IVY/RED)'] == 'IVY') | (order_df['Company (IVY/RED)'] == 'IVY & RED')) & (order_df['Completed'] != True)
    order_df = order_df[condition]
    order_df = order_df.reset_index()

    if len(order_df)>0:
        print(order_df)
        A=input("Press Any Key to exit")
    else:
        print(order_df)
        
        import time
        time.sleep(5)

check_smartsheet(desktoploc)

