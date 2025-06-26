#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: yuangli
"""
# from tkinter import *
# pip install openpyxl
import tkinter as tk
from tkinter import ttk
import pandas as pd
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
import math
import os
import sys

# calculate total price: ProductPrice * ProductUnit
def calculate(unit_num,query,df,sframe):
    """
    Calculate the specific products price : ProductionPrice * ProductUnit

    Parameters:
        :param unit_num (int): User input of number of product
        :param query (str):  User input of product type
        :param df (pd.DataFrame): Product info data
        :param sframe (Frame): Scrollable frame
        :return: None
    """
    global price
    global counter
    global result_df
    
    selected_type = get_entry_value(query)
    product_info = df.loc[df["ProductType"]== selected_type].iloc[0] 
    unit_price = float(product_info["ProductPrice"])   
    product_name = product_info["ProductName"]
    product_unit = int(get_entry_value(unit_num))
    price = round(product_unit * unit_price,1)
    
    # for documents' usage
    product_info['Quantity'] = product_unit
    product_info['Net Amount'] = price
    result_df = pd.concat([result_df, product_info.to_frame().T], ignore_index=True)

    # list calculated prices, updated index
    counter += 1
    label_result = tk.Label(sframe, text=f"{counter}: {product_name} with ${unit_price}/unit and {product_unit} pcs is total ${price}")
    label_result.grid(row = counter, column = 0,columnspan = 4,sticky ="ew", padx=(0,50))

def load_type(df):
    type_list = df['ProductType'].tolist()
    return type_list

def load_name(df):
    name_list = df['ProductName'].tolilst()
    return name_list

def load_price(df):
    price_list = df['ProductPrice'].tolist()
    return price_list

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # Temp folder PyInstaller uses
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def generate_docs(df,entry_quoteDate,entry_quoteNum):
    """
    Generate quotation and packing list PDF files with A4 size

    :param df (pd.DataFrame): Product info data:
    :param entry_quoteDate:  User input of quote date
    :param entry_quoteNum:  User input of quote reference number
    :return: None

    """
    quote_num = get_entry_value(entry_quoteNum)
    quote_date = get_entry_value(entry_quoteDate)
    h = 792
    # Generate folder to store and manage files for users
    folder_name = 'Sales_Documents'
    parent_path = os.path.expanduser('~/Desktop')
    isExist = os.path.exists(os.path.expanduser('~/Desktop/Sales_Documents'))
    if isExist:
        print("Sales_Documents folder already exists")
    else:
        path = os.path.join(parent_path, folder_name)
        os.makedirs(path, exist_ok= True)
        print("Folder exist or not", isExist)
    # save_name = os.path.join(os.path.expanduser("~"), "~/Desktop/Sales_Documents/", f"Quotation_{quote_num}.pdf")
    save_name =os.path.join(os.path.expanduser('~/Desktop/Sales_Documents'),f"Quotation_#{quote_num}.pdf")



    # Generate quotation
    # set up header and title
    c = pdf_canvas.Canvas(save_name, pagesize=A4)
    # image_path = resource_path("header.png")
    image_path = resource_path("test_header.png")
    c.drawImage(image_path, 30, h- 180, width = 550, height = 230)
    title = c.beginText(230, h -200)
    title.setFont("Times-Roman", 20)
    title.textLine("Quotation Sheet")
    c.drawText(title)
    # referencce number & date
    quote_info = c.beginText(400, h-250)
    quote_info.setFont("Times-Roman", 12)
    quote_info.textLine(f"Eagled Order Number: {quote_num}")
    quote_info.textLine(f"Order Date: {quote_date}")
    c.drawText(quote_info)

    # Calculate packing size
    def cal_cbm(size):
        try:
            l, w, h = [int(x.strip()) for x in size.split('X')]
            return (l * w * h ) / 1000000
        except Exception:
            print(f"Erro parsing : {size}, {Exception}")
            return None

    df["Boxes"] = df["Quantity"] / df["PackingUnit"]  
    df["Boxes"] = df["Boxes"].astype(int)
    df["BoxesSize"] = df["PackingSize"].apply(cal_cbm)
    df["Quantity"] = df["Quantity"].astype(int)
    df["Cbm"] = df["Boxes"] * df["BoxesSize"]

    #shiping fee : 109 / cbm
    shipment_fee = df["Cbm"].sum() * 109
    # only get ProductName, ProductType , Quantity, Unit Price, Net Amount
    quote_df = df.iloc[:,[0,1,5,2,6]]
    # calculate the shipping fee
    new_data = {col: "/" for col in quote_df.columns}
    new_data["ProductName"] = "Shipping Fee to FF"
    # round up to int
    new_data["Net Amount"] = math.ceil(shipment_fee)
    new_row = pd.DataFrame([new_data])
    quote_df = pd.concat([quote_df, new_row], ignore_index=True)
    # Calculate total price
    new_data = {col: "/" for col in quote_df.columns}
    new_data["ProductName"] = "NET (CNY)"
    new_data["Net Amount"] = quote_df["Net Amount"].sum()
    new_row = pd.DataFrame([new_data])
    quote_df = pd.concat([quote_df, new_row], ignore_index=True)
    quote_df= quote_df.rename(columns={"ProductName": "Description",
                                       "ProductType": "Model",
                                       "Quantity" : "Net Quantity", 
                                       "ProductPrice" : "Unit Price",
                                       "Net Amount" : "Net Amount" 
                                      })

    x_start = 45
    y_start = h - 350
    row_height = 20
    col_width = 80
    first_col_width = 185
    x_pos = x_start
    # Draw header
    for i, col in enumerate(quote_df.columns):
        width = first_col_width if i == 0 else col_width
        c.setFillColor(colors.lightgrey)
        c.rect(x_pos, y_start, width, row_height, fill=1)
        c.setFillColor(colors.black)
        c.drawString(x_pos + 5, y_start + 5, str(col))
        x_pos += width 

    # Draw rows
    for row_index, row in quote_df.iterrows():
        y = y_start - (row_index + 1) * row_height
        x_pos = x_start
        for col_index, value in enumerate(row):
            width = first_col_width if col_index == 0 else col_width
            c.rect(x_pos,y,width,row_height)
            c.drawString(x_pos + 5, y+5, str(value))
            x_pos += width
    # Notes part
    quote_remarks = c.beginText(50, 100)
    quote_remarks.setFont("Times-Roman", 12)
    quote_remarks.textLine("Remarks:")
    quote_remarks.textLine("1. Terms of Price: total price excludes VA, Custom")
    quote_remarks.textLine("2. Terms of Payment: 50% before production, 50% before shipping")
    quote_remarks.textLine("3. Terms of Production Time:  around 4 business days after all payment")
    c.drawText(quote_remarks)
    c.save()

#  Packing list generation
    h = 792
    # save_name = os.path.join(os.path.expanduser("~"), "Desktop/Sales_Documents/", f"Packing List_{quote_num}.pdf")
    save_name = os.path.join(os.path.expanduser('~/Desktop/Sales_Documents'), f"Packing List_#{quote_num}.pdf")
    c = pdf_canvas.Canvas(save_name, pagesize=A4)
    # image_path = resource_path("header.png")
    c.drawImage(image_path, 30, h- 180, width = 550, height = 230)

    title = c.beginText(230, h -200)
    title.setFont("Times-Roman", 20)
    title.textLine("Packing List")
    c.drawText(title)

    quote_info = c.beginText(400, h-250)
    quote_info.setFont("Times-Roman", 12)
    quote_info.textLine(f"Eagled Order Number: {quote_num}")
    quote_info.textLine(f"Order Date: {quote_date}")
    c.drawText(quote_info)
    packing_df = df.iloc[:,[0,1,7,3,5,9,4]]

    # calculate the shipping fee
    new_data = {col: "/" for col in packing_df.columns}
    new_data["ProductName"] = "NET "
    new_data["Boxes"] = int(packing_df["Boxes"].sum())
    new_data["Quantity"] = int(packing_df["Quantity"].sum())
    new_data["Cbm"] = packing_df["Cbm"].sum()
    total_boxes = packing_df["Boxes"].sum()
    total_quantity = packing_df["Quantity"].sum()
    total_cbm = packing_df["Cbm"].sum()

    new_row = pd.DataFrame([new_data])
    packing_df = pd.concat([packing_df, new_row], ignore_index=True)
    packing_df= packing_df.rename(
    columns={
        "ProductName": "Description",
        "ProductType": "Model",
        "Boxes":"Cartons",
        "PackingUnit":"Quantity/ctn",
        "Quantity" : "Net Quantity",
        "Cbm": "Volume(cbm)",
        "PackingSize": "Size(cm)"
    })

    x_start = 3
    y_start = h - 350
    row_height = 20
    col_width = 73
    first_col_width = 174
    x_pos = x_start

    # update volume to only keep 3 decimal points
    packing_df['Volume(cbm)'] = round(packing_df['Volume(cbm)'],3)


    # Draw header
    for i, col in enumerate(packing_df.columns):
        if i == 0:
            width = first_col_width
        elif i == 1:
            width = 60
        elif i == 2:
            width = 43
        else: 
            width = col_width
        c.setFillColor(colors.lightgrey)
        c.rect(x_pos, y_start, width, row_height, fill=1)
        c.setFillColor(colors.black)
        c.drawString(x_pos + 5, y_start + 5, str(col))
        x_pos += width

    # Draw rows
    for row_index, row in packing_df.iterrows():
        y = y_start - (row_index + 1) * row_height
        x_pos = x_start
        for col_index, value in enumerate(row):

            if col_index == 0:
                width = first_col_width
            elif col_index == 1:
                width = 60
            elif col_index == 2:
                width = 43
            else: 
                width = col_width
            c.rect(x_pos,y,width,row_height)
            c.drawString(x_pos + 5, y+5, str(value))
            x_pos += width

    quote_remarks = c.beginText(50, 100)
    quote_remarks.setFont("Times-Roman", 12)
    quote_remarks.textLine("Remarks:")
    quote_remarks.textLine(f"1. Total Number of Boxes: {total_boxes}")
    quote_remarks.textLine(f"2. Total Order Quantity: {total_quantity}")
    quote_remarks.textLine(f"3. Total Order Volume: {total_cbm} (CBM)" )
    c.drawText(quote_remarks)
    c.save()

def back_home(parent):
    global backup_df
    global root
    clear_frame(parent)
    main_page(root)

# Claculation page
def cal_page(parent,df):
    global price
    global result_df
    clear_frame(parent)

  # responsive grids
    for col in range(0,4):
        parent.columnconfigure(col, weight=1, minsize=10)

    for row in range(0,4):
        parent.rowconfigure(row, weight=0, minsize=5)
    
    # Title
    label_title = tk.Label(parent, text= "Price Calculator")
    label_title.grid(row = 0, column = 0,columnspan = 4,sticky ="ew", pady=(0,30), padx= (0,0))
    
    # User input widget to collect product type
    query_name= tk.Label(parent,text ="Product Type")
    query_name.grid(row = 1, column = 0, padx=(50,2))
    # search query
    product_type = load_type(df)
    query = ttk.Combobox(parent,values= product_type)  
    query.set("Select or type...")
    query.grid(row = 1, column = 1)
    # bind event handler
    query.bind('<KeyRelease>', lambda event: matchkey(event, query, product_type))

    # User input widget to collect number of products
    label_unit = tk.Label(parent,text = "Product unit")
    label_unit.grid(row = 2, column = 0, padx=(50,2))
    # unit input entry
    entry_unit = tk.Entry(parent)
    entry_unit.grid(row = 2 , column = 1)

    # calculate price
    btn_start = tk.Button(parent,text="Calculate",command= lambda:
                       calculate(entry_unit,query,df,scrollable_frame)).grid(row =3, column = 2, padx = (40,2), sticky = "ew")

    # clear results
    btn_clear = tk.Button(parent,text="Clear Result",command= lambda:
                        clear_result(scrollable_frame)).grid(row =3, column = 3, padx = (40,2), sticky = "ew")

    # scrollable section starts
    # show results
    tk.Label(parent, text="Results").grid(row=4, column=0, columnspan=4)
    # === Scrollable Results Frame ===
    results_container = tk.Frame(parent)
    results_container.grid(row = 5, column = 0, columnspan = 4, sticky = "nsew", padx = 20, pady = 10)
    canvas_frame = tk.Canvas(results_container, height = 200)
    scrollbar = tk.Scrollbar(results_container, orient= "vertical", command= canvas_frame.yview)
    scrollable_frame = tk.Frame(canvas_frame)
    scrollable_frame.bind("<Configure>", lambda e: canvas_frame.configure(scrollregion= canvas_frame.bbox("all")))
    canvas_frame.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas_frame.configure(yscrollcommand=scrollbar.set)
    canvas_frame.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # generate quotation & packing list
    # Collect user input of quotation reference number
    label_quoteNum = tk.Label(parent,text = "Quotation Number:")
    label_quoteNum.grid(row = 6, column = 0, padx=(40,2))
    entry_quoteNum = tk.Entry(parent)
    entry_quoteNum.grid(row = 6 , column = 1)
    # Collect user input of quotation date
    label_quoteDate = tk.Label(parent,text = "Quotation Date:")
    label_quoteDate.grid(row = 6, column = 2, padx=(40,2))
    entry_quoteDate = tk.Entry(parent)
    entry_quoteDate.grid(row = 6 , column = 3)

    btn_quote = tk.Button(parent, text="Generate Quotation & Packing List", command= lambda: generate_docs(result_df,entry_quoteDate,entry_quoteNum))
    btn_quote.grid(row = 7, column =3,  pady=20, sticky = "we")
    btn_packing = tk.Button (parent,text="Home", command= lambda: back_home(parent) )
    btn_packing.grid(row = 7, column = 0,  pady= 20, sticky = "we")

# Filterable product type input widget
def matchkey(event,query,product_type):
    # get user's input 
    value = get_entry_value(query)
    # show all products
    if value == '' or value == 'Select or type...':
        input_data = product_type
    # filter product types
    else:
        input_data = []
        for i in product_type:
            if value.lower() in i.lower():
                input_data.append(i)

        if len(input_data) == 0:
            input_data = [f"Didn't find products with {get_entry_value(query)} type"]
    # update dropdown list
    query.config(values = input_data)

# clear all price results & reset index 
def clear_result(parent):
    
    global result_df
    global counter
    result_df = pd.DataFrame()
    print("cleared result df", result_df)
    for widget in parent.winfo_children():
        widget.destroy()
    counter = 0

def clear_frame(parent):
    for widget in parent.winfo_children():
        widget.destroy()

def exit_file(var):
    var.destroy

def get_entry_value(entry):
    value = entry.get()
    return value

# Main page
result_df = pd.DataFrame()
counter = 0
df = pd.read_excel('~/PythonProject/salesapp/data.xlsx')
backup_df = df
root = tk.Tk()

root.title("Ealged Sales Team")

# define size of GUI frame
root.geometry('830x800')
# main_frame = tk.Frame(root)
menubar = tk.Menu(root)
# set up calculate page menu 
cal = tk.Menu(menubar,tearoff= 0)
menubar.add_cascade(label ='Calculation', menu = cal)
cal.add_command(label ='calculate', command = lambda: cal_page(root,df))
cal.add_command(label = 'Exit', command= root.destroy)
product_info = tk.Menu(menubar,tearoff = 0)
menubar.add_cascade(label = 'Product info', menu = product_info)
root.config(menu = menubar)

# Edit 
n_instance = 0
cells = {}
result_list = []
main_frame = tk.Frame(root)

def main_page(root):
    global cells
    global result_list
    global n_instance
    global counter
    clear_frame(root)

    counter = 0
    df = pd.read_excel('~/PythonProject/salesapp/data.xlsx')
    backup_df = df

    # Set up menu
    menubar = tk.Menu(root)
    cal = tk.Menu(menubar,tearoff= 0)
    menubar.add_cascade(label ='Calculation', menu = cal)
    cal.add_command(label ='calculate', command = lambda: cal_page(root,df))
    cal.add_command(label = 'Exit', command= root.destroy)
    product_info = tk.Menu(menubar,tearoff = 0)
    menubar.add_cascade(label = 'Product info', menu = product_info)
    root.config(menu = menubar)

    main_frame = tk.Frame(root)
    main_frame.pack(expand=True, fill="both", padx=10, pady=10)
    # Configure grid weights for canvas
    main_frame.grid_rowconfigure(0, weight=1)
    main_frame.grid_columnconfigure(0, weight=1)

    # Set up scrollable area
    canvas = tk.Canvas(main_frame)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=0, column=1, sticky="ns")
    canvas.grid(row=0, column=0, sticky="nsew")
    scrollable_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    scrollable_frame.bind("<Configure>", on_configure)

    # set up scrollable content
    df = backup_df
    col_name = df.columns
    n_rows = df.shape[0]
    n_cols = df.shape[1]
    cells = {}
    result_list = []
    n_instance = n_rows

    # will only reset the data based on data.xlsx
    def display_data(scrollable_frame, col_name,n_rows,n_cols,backup_df):

        df = backup_df
        # draw table headers
        for j, col in enumerate(col_name):
            text = tk.Text(scrollable_frame, width=20, height=1, bg="#9BC2E6")
            text.grid(row=0, column=j, padx=1, pady=1)
            text.insert(tk.INSERT, col)
            cells[(0,j)] = text

        # Data table cells
        for i in range(n_rows):
            for j in range(n_cols):
                text = tk.Text(scrollable_frame, width=20, height=1)
                text.grid(row=i+1, column=j, padx=1, pady=1)
                text.insert(tk.INSERT, df.iloc[i, j])
                cells[(i,j)] = text

    display_data(scrollable_frame, col_name,n_rows,n_cols,df)

    # Button to save changes into new excel file
    btn_save = tk.Button(root, text="Save changes", command= lambda: save_changes(n_cols,df,cells,col_name))
    btn_save.pack(pady=20, side= tk.RIGHT)
    # Button to reload remove changes
    btn_cancel = tk.Button (root,text="Cancel edits", command= lambda: display_data(scrollable_frame, col_name,n_rows,n_cols,df))
    btn_cancel.pack(pady= 20, side =tk.RIGHT)


    def save_changes(n_cols,df,cells,col_name):
        """
        This function saves the changes of product info into excel file

        :param n_cols (int): number of columns
        :param df(pd.DataFrame): product data):
        :param cells (dict): data instances of all products
        :param col_name(str): column name
        :return:
        """
        global n_instance
        print ("Number of instance in save changes", n_instance)
        for i in range (n_instance):
            print ("iteration",i)
            for j in range(n_cols):
                try:
                    value = cells[(i, j)].get("1.0", "end-1c")
                    df.at[i,col_name[j]] = value
                    print ("Vlaue is this", value)
                except Exception as e:
                    print(f"Error at row {i}, column {col_name[j]}: {e}")
        try:
            # Saved in a new file to avoid affecting correct product info data
            # TODO: Save to cloud file
            df.to_excel("New_test_data.xlsx", index = False)
            print("New Data Saved")
        except Exception as e: 
            print (f"Failed to save file : {e}")

    # Button to add new data
    btn_insert = tk.Button(root, text = "Add products", command= lambda: add_space(scrollable_frame))
    btn_insert.pack(pady= 20, side = tk.RIGHT)

    # Add new empty row
    def add_space(scrollable_frame):
        global n_instance
        global cells
        print("before adding n_instance is ,",n_instance)
        n_instance += 1
        for j in range(n_cols):
            text = tk.Text(scrollable_frame, width=20, height=1)
            text.grid(row = n_instance , column=j, padx=1, pady=1)
            text.insert(tk.INSERT, f'NaN{n_instance}')
            cells[(n_instance -1,j)] = text
        print ("NUM of current instance", n_instance)

    # Button to enter Calculation page
    btn_insert = tk.Button(root, text = "Report", command= lambda: cal_page(root, df))
    btn_insert.pack(pady= 20, side = tk.LEFT)

main_page(root)
root.mainloop()


