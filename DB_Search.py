import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook



complaints = pd.read_excel('MKE Complaint Data.xlsx')
labels = pd.read_excel('Complaint Label.xlsx')
order_numbers = pd.read_excel("Found Order Numbers.xlsx")



def key_word_search():
    index_one = -1
    for index, row in complaints.iterrows():
        index_one += 1 
        try:           
            if 'precipitation' in row['Problem Description'] or 'particulate' in row['Problem Description'] or 'Silica' in row['Problem Description'] or 'Olefin' in row['Problem Description'] or 'oxidi' in row['Problem Description'] or 'particle' in row['Problem Description']:                
                labels.loc[-1] = complaints.loc[index_one]
                labels.index = labels.index + 1
                labels = labels.sort_index()  
                
            else:                
                
                continue
                
        except TypeError:
            
            continue    
        
    key_word_data = labels.to_excel('Complaint_Data_Key.xlsx', index = True)
    
    
def order_number_search(dataframe_one):    
    index_one = -1
    emp_list = {'Index' : [], 'Order': []}
    for index, row in dataframe_one.iterrows():
        index_one += 1 
        try:
            
            for z in str(row['Problem Description']).split():
                if z.isdigit():
                    emp_list['Index'].append(index)
                    emp_list['Order'].append(z)
                else:
                    continue         
            
            
                
        except TypeError:
            
            continue    
     
    final_df = pd.DataFrame(emp_list)
    order_data = final_df.to_excel('Found Order Numbers.xlsx', index = True)


def filter_on (df_one):
    index_two = -1
    filter_list = {'Index' : [], 'Order': []}
    for index, row in df_one.iterrows():
        index_two += 1
        if len(str(row['Order'])) == 9:
            filter_list['Index'].append(row['Index'])
            filter_list['Order'].append(row['Order'])
    filter_df = pd.DataFrame(filter_list)       
    filter_data = filter_df.to_excel('Filtered Order Numbers.xlsx')
    
    
    
filter_on(order_numbers)