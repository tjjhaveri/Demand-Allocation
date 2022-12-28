from tkinter import *
import pandas as pd
from tkinter import ttk, filedialog
import MIPFormulation

root = Tk()
root.iconbitmap('tesla.icns')
root.title('Allocation Optimization')
root.geometry('1000x1200')

# File open function
def file_open():
    filename = filedialog.askopenfilename(
        initialdir = '/',
        title = 'Open a .csv file',
        filetypes=(('Excel files','*.xlsx'),('All files','*.*'))
        )
    if filename:
        try:
            filename = r"{}".format(filename)
            global df
            df = pd.read_excel(filename)
        except ValueError:
            warning_label.config(text='File could not be opened...\nTry again!')
        except FileNotFoundError:
            warning_label.config(text='File could not be opened...\nTry again!')
        print(filename)
    #Clear old tree view
    default_tree.destroy()
    clear_tree()
    
    #Set up new tree view
    my_tree['column'] = list(df.columns)
    my_tree['show'] = 'headings'
    
    #loop thru column list
    for column in my_tree['column']:
        my_tree.heading(column, text=column)
    
    # put data in treeview
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        my_tree.insert("","end",values=row)
    my_tree.pack()

def clear_tree():
    my_tree.delete(*my_tree.get_children())

def check_empty():
    global flag
    flag = 1
    try:
        df
    except NameError:
        # Add warning label
        global B6
        flag = 0
        B6 = Label(processor_frame, text='Please upload a dataframe!')
        B6.grid(column=1,row=5)

def show_result():
    # Clear old tree view
    default_result.destroy()
    result_tree.delete(*result_tree.get_children())

    # Set up new tree view
    result_tree['column'] = list(processed_df.columns)
    result_tree['show'] = 'headings'

    # loop thru column list
    for column in result_tree['column']:
        result_tree.heading(column, text=column)

    # put data in treeview
    df_rows = processed_df.to_numpy().tolist()
    for row in df_rows:
        result_tree.insert("", "end", values=row)
    result_tree.pack()



def get_value():
    check_empty()
    if flag == 1:
        try:
            B6.destroy()
        except:
            pass
        ps_max = float(B2_val.get())/100
        ps_min = float(B3_val.get())/100
        lt_level = str(B4.get())
        msu_score = float(B5_val.get())/100
        global processed_df
        processed_df = MIPFormulation.Function(df,ps_max,ps_min,msu_score,lt_level)
        show_result()
        print(ps_max, ps_min, lt_level, msu_score)
        print(processed_df.columns)

def save_file():
    save_to = filedialog.asksaveasfile(
        initialdir="/",
        title='Save to',
        filetypes=(("Excel file", "*.xlsx"), ("All Files", "*.*")),
        defaultextension='.xlsx'
    )
    if save_to:
        try:
            processed_df.to_excel(save_to.name)
        except FileNotFoundError:
            warning_label.config(text='File could not be opened...\nTry again!')

'''
Sec 1--------------------------------------------------------------------<
'''
# Create Tree Frame
data_frame = Frame(root)
data_frame.pack(pady=30)

# Add input header
input_header = Label(data_frame,text='Input Data:')
input_header.configure(font=("bold", 16))
input_header.pack(side=TOP, anchor=NW)

# Add Scrollbar
ytree_scroll = Scrollbar(data_frame)
ytree_scroll.pack(side=RIGHT,fill=Y)

xtree_scroll = Scrollbar(data_frame,orient='horizontal')
xtree_scroll.pack(side=BOTTOM,fill=X)

# Add default view
columns = ('Vendor_Name', 'TPN', 'Component_Category','Supplier_Preference','Pricing_per_Item','Lead_Time_wks','Capacity','Tesla_Demand')
default_tree = ttk.Treeview(data_frame, columns=columns, show='headings')
#loop thru column list
for column in default_tree['columns']:
    default_tree.heading(column, text=column)
default_tree.pack()

#Create treeview
my_tree = ttk.Treeview(data_frame,yscrollcommand=ytree_scroll.set,xscrollcommand=xtree_scroll.set)

# Add a menu
my_menu = Menu(root)
root.config(menu=my_menu)

file_menu = Menu(my_menu,tearoff = False)
my_menu.add_cascade(label='File',menu=file_menu)
file_menu.add_command(label='Open',command=file_open)

# Config Scrollbar
ytree_scroll.config(command=my_tree.yview)
xtree_scroll.config(command=my_tree.xview)

# Add warning label
warning_label = Label(data_frame,text='')
warning_label.pack(side=BOTTOM, anchor=SE)

'''
Sec 2--------------------------------------------------------------------<
'''

# Create Processor Frame
processor_frame = Frame(root)
processor_frame.pack(pady=10)

# Upper limits
# Title
A1 = Label(processor_frame, text = "In any category, total maximum allocation:")
A1.grid(column=0,row=0)
# Front Label
A2 = Label(processor_frame, text = "Preferred Supplier:")
A2.grid(column=0,row=1)
A3 = Label(processor_frame, text = "Unknown/Not-too-preferred Supplier:")
A3.grid(column=0,row=2)
# Prefered Max allocation Entry
B2 = StringVar()
B2.set("70")
B2_val = Entry(processor_frame,textvariable = B2)
B2_val.grid(column=1,row=1)
# Not Preferred Max allocation Entry
B3 = StringVar()
B3.set("20")
B3_val = Entry(processor_frame,textvariable = B3)
B3_val.grid(column=1,row=2)
# Unit Label
C2 = Label(processor_frame, text = "%")
C2.grid(column=2,row=1)
C3 = Label(processor_frame, text = "%")
C3.grid(column=2,row=2)

# Lead time importance level
# Label
A4 = Label(processor_frame, text = "Lead time importance level:")
A4.grid(column=0,row=3)
# Dropdown
B4 = StringVar()
B4.set("Medium")
B4_drop = OptionMenu(processor_frame,B4,"High","Medium","Low")
B4_drop.grid(column=1,row=3)

# Multi-sourcing utilization score
# Label
A5 = Label(processor_frame, text = "Multi-sourcing utilization score:")
A5.grid(column=0,row=4)
# Entry
B5 = StringVar()
B5.set("95")
B5_val = Entry(processor_frame,textvariable = B5)
B5_val.grid(column=1,row=4)

# Unit Label
C5 = Label(processor_frame, text = "%")
C5.grid(column=2,row=4)

# Run Button
C6 = Button(processor_frame,text='Run',command=get_value)
C6.grid(column=2,row=5)
'''
Sec 3--------------------------------------------------------------------<
'''
# Create Result Frame
result_frame = Frame(root)
result_frame.pack(pady=10)

# Add input header
result_header = Label(result_frame,text='Result Preview:')
result_header.configure(font=("bold", 16))
result_header.pack(side=TOP, anchor=NW)

# Add Scrollbar
yresult_scroll = Scrollbar(result_frame)
yresult_scroll.pack(side=RIGHT,fill=Y)

xresult_scroll = Scrollbar(result_frame,orient='horizontal')
xresult_scroll.pack(side=BOTTOM,fill=X)

# Add default view
result_columns = ('Telsa PN', 'Vendors', '2020 Price', 'Lead Time', 'Supplier Preference', 'Capacity',\
       'Allocation Fraction', 'Allocation Quantity')
default_result = ttk.Treeview(result_frame, columns=result_columns, show='headings')
#loop thru column list
for column in default_result['columns']:
    default_result.heading(column, text=column)
default_result.pack()

#Create treeview
result_tree = ttk.Treeview(result_frame,yscrollcommand=yresult_scroll.set,xscrollcommand=xresult_scroll.set)

# Config Scrollbar
yresult_scroll.config(command=result_tree.yview)
xresult_scroll.config(command=result_tree.xview)

# Create Save button
Save = Button(root,text='Save',command=save_file)
Save.pack(anchor=NE)

root.mainloop()