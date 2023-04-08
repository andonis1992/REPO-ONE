import time, datetime
from openpyxl import Workbook, load_workbook
import os, sys #Standard python lib
import pandas as pd


#Change path to current working directory
os.chdir(sys.path[0])

full_scope = pd.read_excel('Stability Full Scope.xlsx')
info_df = pd.read_excel('MQ100 to 300 concat.xlsx')



class my_dictionary(dict):
     
  # __init__ function
  def __init__(self):
    self = dict()
 
  # Function to add key:value
  def add(self, key, value):
    self[key] = value


df_column_labels = my_dictionary()
df_column_labels_two = my_dictionary()

#matches materials from full scope and adds material info to full database
def filter_on (full_scope_df, info_df_df, label_dict_info, label_dict_scope):
    index_two = -1
    for index, row in full_scope_df.iterrows():
        index_two += 1
        percent = 100 * (index_two / 8848)
        print(str(percent) + "%")
        for index_three, row_three in info_df_df.iterrows():
            if str(row['Material']) == str(row_three['WEB_INDEX']):
                temp_info = info_df_df.iloc[index_three]
                temp_scope = full_scope_df.iloc[index]
                label_dict_info.loc[-1] = temp_info
                label_dict_info.index = label_dict_info.index + 1
                label_dict_info = label_dict_info.sort_index()
                label_dict_scope.loc[-1] = temp_scope
                label_dict_info.scope = label_dict_scope.index + 1
                label_dict_scope = label_dict_scope.sort_index()        
                
                #print(info_df_df.iloc[index_three])
            else:
                continue


#adding column names from info df
for col in info_df.columns:
    df_column_labels.add(col, [])
df_column_labels = pd.DataFrame(df_column_labels)
#adding relevant column names from full scope list
df_column_labels_two.add('SLife', [])
df_column_labels_two.add('RSL', [])
df_column_labels_two = pd.DataFrame(df_column_labels_two)


filter_on(full_scope, info_df, df_column_labels, df_column_labels_two)
#Creating Template Dataframe to append with matches
final_df = df_column_labels.join(df_column_labels_two)

final_df.to_excel("Test_Data.xlsx")

