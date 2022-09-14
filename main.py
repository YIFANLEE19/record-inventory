import tkinter as tk
from openpyxl import Workbook,load_workbook
from tkinter import END, ttk,messagebox
from pyparsing import col
from tkcalendar import DateEntry

#workbook location 
wb = load_workbook('myInventory.xlsx')
ws = wb.active

def errorMessage():
    tk.messagebox.showinfo("Error", "Please make sure important information is filled in！")

def excel():
    ws.append([ent_monitor_model.get(),
        int(ent_monitor_size.get()),
        ent_serial_number.get(),
        ent_return_from.get(),
        ent_old_user.get(),
        str(date_return.get_date()),
        ent_new_user.get(),
        str(date_received.get_date()),
        ent_remark.get(),
        cb_status.get()]
    )
    wb.save('CCO Stock Test.xlsx')

def finishSubmit():
    ent_monitor_model.delete(0,END)
    ent_monitor_size.delete(0,END)
    ent_serial_number.delete(0,END)
    ent_return_from.delete(0,END)
    ent_old_user.delete(0,END)
    ent_new_user.delete(0,END)
    ent_remark.delete(0,END)

#clear btn function
def clear_btn(event):
    ent_monitor_model.delete(0,END)
    ent_monitor_size.delete(0,END)
    ent_serial_number.delete(0,END)
    ent_return_from.delete(0,END)
    ent_old_user.delete(0,END)
    ent_new_user.delete(0,END)
    ent_remark.delete(0,END)

#submit btn function
def submit_btn(event):
    if ent_monitor_model.get() == "" or ent_monitor_size.get() == "" or ent_serial_number.get() == "" or ent_new_user.get() == "" or ent_old_user.get() == "" or ent_return_from.get() == "":
        errorMessage()
    else:
        excel()
        finishSubmit()
    
window = tk.Tk()
window.title('CCO Inventory Stock')
window.resizable(0,0)
#create a frame for frm_form
frm_form = tk.Frame(
    relief=tk.FLAT,
    borderwidth=3
)
frm_form.pack(
    padx=10
)

#label and entry widgets
#monitor model
lbl_monitor_model = tk.Label(
    master=frm_form,
    text="Monitor Model: "
)
ent_monitor_model = tk.Entry(
    master=frm_form,
    width=50
)
lbl_monitor_model.grid(
    row=0,
    column=0,
    sticky="e",
    pady=5
)
ent_monitor_model.grid(
    row=0,
    column=1
)

#monitor size
lbl_monitor_size = tk.Label(
    master=frm_form,
    text="Size: "
)
ent_monitor_size = tk.Entry(
    master=frm_form,
    width=50
)
lbl_monitor_size.grid(
    row=1,
    column=0,
    sticky="e",
    pady=5
)
ent_monitor_size.grid(
    row=1,
    column=1
)

#serial number
lbl_serial_number = tk.Label(
    master=frm_form,
    text="SN: "
)
ent_serial_number = tk.Entry(
    master=frm_form,
    width=50
)
lbl_serial_number.grid(
    row=2,
    column=0,
    sticky="e",
    pady=5
)
ent_serial_number.grid(
    row=2,
    column=1
)

#return from (department)
lbl_return_from = tk.Label(
    master=frm_form,
    text="Return From: "
)
ent_return_from = tk.Entry(
    master=frm_form,
    width=50
)
lbl_return_from.grid(
    row=3,
    column=0,
    sticky="e",
    pady=5
)
ent_return_from.grid(
    row=3,
    column=1
)

#old user (user)
lbl_old_user = tk.Label(
    master=frm_form,
    text="Old User: "
)
ent_old_user = tk.Entry(
    master=frm_form,
    width=50
)
lbl_old_user.grid(
    row=4,
    column=0,
    sticky="e",
    pady=5
)
ent_old_user.grid(
    row=4,
    column=1
)

#date return
lbl_date_return = tk.Label(
    master=frm_form,
    text="Date Return: "
)
date_return = DateEntry(frm_form,selectmode='day')
lbl_date_return.grid(
    row=5,
    column=0,
    sticky="e",
    pady=5
)
date_return.grid(
    row=5,
    column=1,
    sticky='w'
)

#new user (user)
lbl_new_user = tk.Label(
    master=frm_form,
    text="New User: "
)
ent_new_user = tk.Entry(
    master=frm_form,
    width=50
)
lbl_new_user.grid(
    row=6,
    column=0,
    sticky="e",
    pady=5
)
ent_new_user.grid(
    row=6,
    column=1
)

#date received
lbl_date_received = tk.Label(
    master=frm_form,
    text="Date Received: "
)
date_received = DateEntry(frm_form,selectmode='day')
lbl_date_received.grid(
    row=7,
    column=0,
    sticky="e",
    pady=5
)
date_received.grid(
    row=7,
    column=1,
    sticky="w"
)

#remark
lbl_remark = tk.Label(
    master=frm_form,
    text="Remark: "
)
ent_remark = tk.Entry(
    master=frm_form,
    width=50
)
lbl_remark.grid(
    row=8,
    column=0,
    sticky="e",
    pady=5
)
ent_remark.grid(
    row=8,
    column=1
)

#status
lbl_status = tk.Label(
    master=frm_form,
    text="Status: "
)
n = tk.StringVar()
cb_status = ttk.Combobox(
    master=frm_form, 
    width = 27, 
    textvariable = n
)
cb_status['values'] = (
    'usable',
    'spoiled'
)
cb_status['state'] = 'readonly'
cb_status.current(0)
lbl_status.grid(
    row=9,
    column=0,
    sticky="e",
    pady=5
)
cb_status.grid(
    row=9,
    column=1,
    sticky="w"
)

#create new frame for buttons
frm_buttons = tk.Frame()
frm_buttons.pack(
    fill=tk.X,
    ipadx=5,
    ipady=5,
    padx=10
)

#button
btn_submit = tk.Button(
    master=frm_buttons,
    text="Submit",
    bg="blue",
    fg="white"
)
btn_submit.pack(
    side=tk.RIGHT,
    ipadx=5,
    ipady=5,
)
btn_clear = tk.Button(
    master=frm_buttons,
    text="Clear"
)
btn_clear.pack(
    side=tk.RIGHT,
    padx=5,
    ipady=5
)

btn_submit.bind("<Button-1>",submit_btn)
btn_clear.bind("<Button-1>",clear_btn)


window.mainloop()
