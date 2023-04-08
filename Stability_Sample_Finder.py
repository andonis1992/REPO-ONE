# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""


import time, datetime
from openpyxl import Workbook, load_workbook
import os, sys #Standard python lib
from tkinter import * 


def runSampleFinder(file_name):
    # Loading Family and Sample Tracker
    WB1 = load_workbook("HC 409 Product Family Condensed.xlsx", data_only = True)
    WB1_WS1 = WB1["contour-export"]
    WB2 = load_workbook(file_name+".xlsx", data_only = True)
    WB2_WS1 = WB2["Sheet1"]

    #Establishing timestamped Excel Save Name
    t = time.localtime()
    timestamp = time.strftime('%b-%d-%Y_%H%M', t)
    BACKUP_NAME = ("Product_List_Update-" + timestamp)


    # Index values to correspond to row values to search through
    index_start = 2
    index_end = 1
    product_family_index_end = 921

    #Finds Length of excel sheet for index referencing
    while WB2_WS1.cell(column = 6, row = index_end).value != None:
        index_end +=1
    index_end -= 1

    #creation of lists for each column  
    column_a = []
    column_b = []
    column_c = []
    column_d = []
    column_e = []
    column_f = []
    column_g = []
    column_h = []
    column_i = []
    column_j = []
    column_k = []
    column_l = []
    column_m = []
    column_n = []
    column_o = []
    column_p = []
    column_q = []
    column_r = []
    column_s = []



    list_of_products_forcast = []
    list_of_products_batch = []
    list_of_products_family = []

    forcast_index = 2


    #Collect Product Number and Batch Number from Forcast List
    while forcast_index <= index_end:
        list_of_products_forcast.append(WB2_WS1.cell(column = 6, row = forcast_index).value)
        list_of_products_batch.append(WB2_WS1.cell(column = 7, row = forcast_index).value)
        forcast_index += 1
        
    forcast_index = 2

    #Collect Product Numbers from Stability List
    while forcast_index <= product_family_index_end:
        list_of_products_family.append(WB1_WS1.cell(column = 1, row = forcast_index).value)
        forcast_index += 1    


    sorted_products_bulk = []
    sorted_products_sample = []
    sorted_batches_bulk = []
    sorted_batches_sample = []

    #Sorts products from forcast list that match with stability list
    check_index = 0
    for items in list_of_products_forcast:
        product_family_index = 2
        for products in list_of_products_family:        
            if items[0:7] == products: 
                if "BULK" in items:
                    sorted_products_bulk.append(items)
                    sorted_batches_bulk.append(list_of_products_batch[check_index])
                elif "SAMPLE" in items or "5G" in items:
                    sorted_products_sample.append(items)
                    sorted_batches_sample.append(list_of_products_batch[check_index])
                else:
                    continue
                            
            else:
                continue
                
        check_index +=1


    #Checks bulk batch list and matches it with sample batch list. Updated list of samples and batches created.(Final list before excel import)
    final_product_list = []
    final_batch_list = []
    check_index = 0

    for items in sorted_batches_bulk:
        sample_index = 0
        for sample in sorted_batches_sample:
            if items == sample:
                final_product_list.append(sorted_products_sample[sample_index])
                final_batch_list.append(sample)
                sample_index += 1
            else:
                sample_index += 1
                continue
            
        check_index += 1    



    products_bulk = dict(zip(final_product_list, final_batch_list))
    print(products_bulk)

    # matches final product list with stability list and pulls respective info and distributes in respective cells.
    product_family_index = 0
    for products in list_of_products_family:
        check_index = 0       
        for items in final_product_list:             
            if products == items[0:7]:
                column_a.append(items)
                column_b.append(final_batch_list[check_index])
                column_c.append(WB1_WS1.cell(column = 2, row = product_family_index+2).value)
                column_d.append(WB1_WS1.cell(column = 3, row = product_family_index+2).value)
                column_e.append(WB1_WS1.cell(column = 17, row = product_family_index+2).value)
                column_f.append(WB1_WS1.cell(column = 18, row = product_family_index+2).value)
                column_g.append(WB1_WS1.cell(column = 19, row = product_family_index+2).value)
                column_h.append(WB1_WS1.cell(column = 20, row = product_family_index+2).value)
                column_i.append(WB1_WS1.cell(column = 23, row = product_family_index+2).value)
                column_j.append(WB1_WS1.cell(column = 24, row = product_family_index+2).value)
                column_k.append(WB1_WS1.cell(column = 25, row = product_family_index+2).value)
                column_l.append(WB1_WS1.cell(column = 26, row = product_family_index+2).value)
                column_m.append(WB1_WS1.cell(column = 27, row = product_family_index+2).value)
                column_n.append(WB1_WS1.cell(column = 28, row = product_family_index+2).value)
                column_o.append(WB1_WS1.cell(column = 31, row = product_family_index+2).value)
                column_p.append(WB2_WS1.cell(column = 10, row = check_index+2).value)
                column_q.append(WB1_WS1.cell(column = 33, row = product_family_index+2).value)
                column_r.append(WB1_WS1.cell(column = 29, row = product_family_index+2).value)
                column_s.append(WB1_WS1.cell(column = 30, row = product_family_index+2).value)
                check_index += 1
                
            else:
                check_index += 1
                continue
                
        product_family_index +=1






    print(column_a)
    print(column_b)
    print(column_c)
    print(column_d)
    print(column_e)
    print(column_f)
    print(column_g)
    print(column_h)
    print(column_i)
    print(column_j)
    print(column_k)
    print(column_l)
    print(column_m)
    print(column_n)
    print(column_o)
    print(column_p)
    print(column_q)
    print(column_r)
    print(column_s)



    #Setting up new Workbook
    workbook = Workbook()
    Export = workbook.active
    Export.title = "Export"

    #Creation of column headers
    Export["A1"] = "Catalog#"
    Export["B1"] = "Batch"
    Export["C1"] = "Product Name"
    Export["D1"] = "CAS #"
    Export["E1"] = "Family"
    Export["F1"] = "Form "
    Export["G1"] = "Retest"
    Export["H1"] = "Expiry/Retest period"
    Export["I1"] = "Storage Conditions**"
    Export["J1"] = "Air/Water React*"
    Export["K1"] = "Humidity-Sens*"
    Export["L1"] = "Air-Sens*"
    Export["M1"] = "Light-Sens*"
    Export["N1"] = "Temp-Sens*"
    Export["O1"] = "Flash Point \n(C) "
    Export["P1"] = "Clear Date"
    Export["Q1"] = "Hazards"
    Export["R1"] = "Acute Hazards"
    Export["S1"] = "Chronic Hazards"


    #defining function to export a column list to a column in excel

    def export_to_excel(column_list, Letter):
        i = 2
        for strings in column_list:
            Export[Letter + str(i)] = strings
            i +=1

    #calling function for columns A-S
    export_to_excel(column_a, "A")
    export_to_excel(column_b, "B")
    export_to_excel(column_c, "C")
    export_to_excel(column_d, "D")
    export_to_excel(column_e, "E")
    export_to_excel(column_f, "F")
    export_to_excel(column_g, "G")
    export_to_excel(column_h, "H")
    export_to_excel(column_i, "I")
    export_to_excel(column_j, "J")
    export_to_excel(column_k, "K")
    export_to_excel(column_l, "L")
    export_to_excel(column_m, "M")
    export_to_excel(column_n, "N")
    export_to_excel(column_o, "O")
    export_to_excel(column_p, "P")
    export_to_excel(column_q, "Q")
    export_to_excel(column_r, "R")
    export_to_excel(column_s, "S")


    #Save workbook as backupname

    workbook.save(filename=BACKUP_NAME +".xlsx")
    
    
    
    
    

#Must always go first creates window screen
root = Tk()

#Defining input box
e = Entry(root, width =35)



#Defining click function
def myClick():
    myLabel3 = Label(root, text="Please Wait, Analyzing " + e.get())
    myLabel3.grid(row=5, column=0)
    runSampleFinder(e.get())
    myLabel3.grid_remove()
    myLabel6 = Label(root, text="Analysis Complete")
    myLabel6.grid(row=5, column=0)

#Creating label widget
#Define the label
myLabel1 = Label(root, text = 'Stability Sample Updater')
myLabel2= Label(root, text = 'Please enter a SQVI Table')
myLabel4= Label(root, text="Only enter file name")
#define Button
myButton = Button(root, text='Run', command=myClick, fg="blue", bg="yellow")
#Putting the label on the screen
myLabel1.grid(row=0, column=0)
myLabel2.grid(row=1, column=0)
myLabel4.grid(row=2, column=0)
myButton.grid(row=4, column=0)
e.grid(row=3, column=0)


root.mainloop()





