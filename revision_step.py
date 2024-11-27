import pandas as pd
from datetime import timedelta

def delete_outdated_rows(final_path, start_date, end_date):
    df = pd.read_excel(final_path)

    #get Date Time Zone from excel_path and convert it to MM/DD/YYYY format from MM/DD/YYYY HH:MM:SS
    df['Date Time Zone'] = pd.to_datetime(df['Date Time Zone']).dt.strftime('%m/%d/%Y')
    
    #convert start_date and end_date from string to datetime format 
    start_date = pd.to_datetime(start_date).strftime('%m/%d/%Y')
    end_date = pd.to_datetime(end_date).strftime('%m/%d/%Y')

    #compare df with start_date and end_date and delete full row that are out of date range
    df = df[(df['Date Time Zone'] >= start_date) & (df['Date Time Zone'] <= end_date)]

    #log message
    print('Deleted outdated rows')

    #save the new df to the final_path
    df.to_excel(final_path, index = False, header=True)
    
def back_to_back_rev(final_path):
    df = pd.read_excel(final_path)

    #get Date Time Zone column from final_path
    df['Date Time Zone'] = pd.to_datetime(df['Date Time Zone'])
    
    df['Back to back'] = ''
    
    for i in range(len(df)):
        if i == 0:
            df['Back to back'][i] = 'Ok'
        else:
            current_feed_index = df.loc[i, 'Feed Index']
            previous_feed_index = df.loc[i-1, 'Feed Index']
            
            current_date_time_zone = df.loc[i, 'Date Time Zone']
            previous_date_time_zone = df.loc[i-1, 'Date Time Zone']
            
            seconds = timedelta(seconds=df.iloc[i-1, 'Duration'])
            
            if current_feed_index != previous_feed_index:
                df.loc[i, 'Back to back'] = 'Ok'
            else:
                if current_date_time_zone <= previous_date_time_zone + seconds:
                    df.loc[i, 'Back to back'] = 'Back to back'
                else:
                    df.loc[i, 'Back to back'] = 'Ok'
        
    #save the new df to the final_path
    with pd.ExcelWriter(final_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='Archivo Final Play Logger', startrow=0, startcol=25, index=False)
        
def full_revision(final_path, start_date, end_date):
    delete_outdated_rows(final_path, start_date, end_date)
    back_to_back_rev(final_path)
    print('Full revision completed')
        
