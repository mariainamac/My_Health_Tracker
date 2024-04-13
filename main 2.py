import tkinter as tk
from tkinter import ttk
import openpyxl

#
##
###
####
#***************CREATING FUNCTIONS*********************

#***************CREATING SWITCH MODE FUNCTIONS*********************

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-dark")
    else:
        style.theme_use("forest-light")

#***************CREATING SUBMIT BUTTON FUNCTIONS*********************

def insert_row():
    name = name_entry.get()
    age = int(age_entry.get())
    weight = float(weight_entry.get())
    food = food_entry.get()
    meal = meal_entry.get()
    calories = float(calorie_entry.get())
    exercise = "Yes" if a.get() else "No"

    row_value = [name, age, weight, food, meal, calories, exercise]

    treeview.insert('', tk.END, values=row_value)

    if row_value not in excel_data:
        excel_data.append(row_value)
        save_data()

    clear_entry_fields()


#***************CREATING LOADING OF DATA FUNCTIONS*********************

def load_data():
    global excel_data
    path = r"C:\Users\Lenovo\PycharmProjects\Health_Tracker_Project\My_Health_Tracker.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    excel_data = list(sheet.values)[1:]

    for col_name in list(sheet.values)[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in excel_data:
        treeview.insert('', tk.END, values=value_tuple)

def save_data():
    path = r"C:\Users\Lenovo\PycharmProjects\Health_Tracker_Project\My_Health_Tracker.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    for row in excel_data:
        sheet.append(row)
    workbook.save(path)

def clear_entry_fields():
    name_entry.delete(0, 'end')
    name_entry.insert(0, 'Name')
    age_entry.delete(0, 'end')
    age_entry.insert(0, 'Age')
    weight_entry.delete(0, 'end')
    weight_entry.insert(0, 'Weight')
    food_entry.delete(0, 'end')
    food_entry.insert(0, 'Food Description')
    meal_entry.delete(0, 'end')
    meal_entry.insert(0, 'Meal (B/L/S/D)')
    calorie_entry.delete(0, 'end')
    calorie_entry.insert(0, 'Calorie Intake')
    checkbutton.state(["!selected"])

#
##
###
####
#***************CREATING THE PARENT WINDOW*********************

root = tk.Tk()
style = ttk.Style(root)

#***************THEME OF THE PARENT WINDOW*********************

root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")

style.theme_use("forest-light")

#
##
###
####
#***************CREATING FRAME INSIDE THE PARENT WINDOW*********************

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text='My Calorie Intake')
widgets_frame.grid(row=0, column=0, padx=0, pady=10)

#
##
###
####
#***************CREATION OF 'NAME' FIELD*********************

name_entry = ttk.Entry(widgets_frame)

name_entry.insert(0, "Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky='ew')

#***************CREATION OF 'AGE' FIELD*********************

age_entry = ttk.Entry(widgets_frame)

age_entry.insert(0, "Age")
age_entry.bind("<FocusIn>", lambda e: age_entry.delete('0', 'end'))
age_entry.grid(row=1, column=0, padx=5, pady=(0, 5), sticky='ew')

#***************CREATION OF 'WEIGHT' FIELD*********************

weight_entry = ttk.Entry(widgets_frame)

weight_entry.insert(0, "Weight")
weight_entry.bind("<FocusIn>", lambda e: weight_entry.delete('0', 'end'))
weight_entry.grid(row=2, column=0, padx=5, pady=(0, 5), sticky='ew')

#***************CREATION OF 'FOOD DESCIPTION' FIELD*********************

food_entry = ttk.Entry(widgets_frame)

food_entry.insert(0, "Food Description")
food_entry.bind("<FocusIn>", lambda e: food_entry.delete('0', 'end'))
food_entry.grid(row=3, column=0, padx=5, pady=(0, 5), sticky='ew')

#***************CREATION OF 'MEAL' FIELD*********************

meal_entry = ttk.Entry(widgets_frame)

meal_entry.insert(0, "Meal (B/L/S/D)")
meal_entry.bind("<FocusIn>", lambda e: meal_entry.delete('0', 'end'))
meal_entry.grid(row=4, column=0, padx=5, pady=(0, 5), sticky='ew')

#***************CREATION OF 'CALORIE' FIELD*********************

calorie_entry = ttk.Entry(widgets_frame)

calorie_entry.insert(0, "Calorie Intake")
calorie_entry.bind("<FocusIn>", lambda e: calorie_entry.delete('0', 'end'))
calorie_entry.grid(row=5, column=0, padx=5, pady=(0, 5), sticky='ew')

#***************CREATION OF 'EXERCISE?' FIELD*********************

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Exercised?", variable = a)
checkbutton.grid(row=6, column=0, sticky='nsew')


#***************CREATION OF SUBMIT BUTTON*********************

submit_button = ttk.Button(widgets_frame, text="Submit", command=insert_row)
submit_button.grid(row=7, column=0, padx=5, pady=5, sticky='nsew')


#***************CREATION OF SEPARATOR　LINE*********************

seperator = ttk.Separator(widgets_frame)
seperator.grid(row=8, column=0, padx=5, pady=5, sticky='nsew')

#***************CREATION OF DARK & LIGHT SWITCH*********************

mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style='Switch',
                              command=toggle_mode)
mode_switch.grid(row=9, column=0, padx=5, pady=5, sticky='nsew')

#
##
###
####
#***************RIGHT SIDE TABLE*********************

#***************VIEWABLE TABLE ON THE PARENT WINDOW*********************

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)

#***************SCROLL BAR & MOVING WITH TBALE CONTENT*********************

treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

#***************ADDING THE HEADER NAMES*********************

cols = ("Name", "Age", "Weight", "Food Description", "Meal", "Calories", "Exercise")
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=13)
for col_name in cols:
    treeview.heading(col_name, text=col_name)
treeview.pack()

#***************ADDING THE CONTENT OF THE TABLE*********************
treeview.column("Name", width=100)
treeview.column("Age", width=100)
treeview.column("Weight", width=100)
treeview.column("Food Description", width=100)
treeview.column("Meal", width=100)
treeview.column("Calories", width=100)
treeview.column("Exercise", width=100)
treeview.pack()

treeScroll.config(command=treeview.yview)

excel_data = []

load_data()

root.mainloop()
