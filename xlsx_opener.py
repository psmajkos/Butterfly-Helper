import openpyxl
import tkinter 
from tkinter import Tk
import tkinter as tk
from tkinter import ttk, Checkbutton, LabelFrame, RIDGE, W, BooleanVar

root = Tk()
root.state('zoomed') 
#root.attributes("-fullscreen", True)
#root.geometry("800x600")
def do_nothing():
    print("do nothing")
menu_bar = tk.Menu(root)

# Add menus to the menu bar
file_menu = tk.Menu(menu_bar, tearoff=0)
edit_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)
menu_bar.add_cascade(label="Edit", menu=edit_menu)

# Add menu items to the "File" menu
file_menu.add_command(label="New", command=do_nothing)
file_menu.add_command(label="Open", command=do_nothing)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)

# Add menu items to the "Edit" menu
edit_menu.add_command(label="Cut", command=do_nothing)
edit_menu.add_command(label="Copy", command=do_nothing)
edit_menu.add_command(label="Paste", command=do_nothing)

# Attach the menu bar to the root window
root.config(menu=menu_bar)
def todo():
    # Open the workbook
    workbook = openpyxl.load_workbook('Plany P15 - v.2 .xlsm', keep_vba=True)

    # Get the active worksheet
    worksheet = workbook['290 Act. 32']

    to_do_lf = LabelFrame(root, text='To-Do', bd=2, relief=RIDGE)

    
    # Create a canvas with a scrollbar
    canvas = tkinter.Canvas(to_do_lf, width=800, height=350)
    scrollbar = ttk.Scrollbar(to_do_lf, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    def on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    canvas.bind_all("<MouseWheel>", on_mousewheel)

    def bind_scroll(event):
        canvas.bind_all("<MouseWheel>", on_mousewheel)

    def unbind_scroll(event):
        canvas.unbind_all("<MouseWheel>")

    canvas.bind("<Enter>", bind_scroll)
    canvas.bind("<Leave>", unbind_scroll)

    # var = BooleanVar()
    # # Read data from the worksheet
    # i=0
    # for row in worksheet.iter_rows(min_row=0, min_col=0, values_only=True):
    #     for j in range(0,len(row)):
    #         data=ttk.Label(scrollable_frame,width=20,text=row[j],
    #         anchor='center')
    #         data.grid(row=i,column=j)
    #         done=ttk.Label(scrollable_frame,width=20, text="Done",
    #         anchor='center')
    #         done.grid(row=0,column=3)

    #         def print_selected():
    #             for child in scrollable_frame.winfo_children():
    #                 if isinstance(child, ttk.Checkbutton):
    #                     if child.instate(['selected']):
    #                         print(row)

    #         done_checkbutton = ttk.Checkbutton(scrollable_frame, text="Option {}".format(i+1), variable=var, onvalue=1, offvalue=0, command=print_selected)
    #         done_checkbutton.grid(row=i, column=j+1)
    #     i += 1

    for i, row in enumerate(worksheet.iter_rows(min_row=5, min_col=5,max_col=9, values_only=True)):
        for j, value in enumerate(row):
            data = ttk.Label(scrollable_frame, width=20, text=value, anchor='center')
            data.grid(row=i, column=j)
            done = ttk.Label(scrollable_frame, width=20, text="Done", anchor='center')
            done.grid(row=0, column=3)

            var = BooleanVar()
            def print_selected(var, val, row_index):
                if var.get() == 1:
                    root.update_idletasks()
                    # Load the workbook
                    workbook = openpyxl.load_workbook('done.xlsx')

                    # # Create a new sheet
                    # sheet = workbook.create_sheet("Sheet1")

                    # Select the sheet you want to add a record to
                    sheet = workbook['Sheet']

                    # Add a new record to the sheet
                    new_record = [worksheet.cell(row=row_index+1, column=1).value, val, 30]
                    sheet.append(new_record)

                    # Save the workbook
                    workbook.save('done.xlsx')

                    print("Costumer Order:", val)
                    print("Zlecenie Produkcji:", worksheet.cell(row=row_index+1, column=1).value)

                    
            done_checkbutton = ttk.Checkbutton(scrollable_frame, variable=var, onvalue=1, offvalue=0, command=lambda var_value=var, val=value, row_index=i: print_selected(var_value, val, row_index))

            done_checkbutton.grid(row=i, column=j+1)
    # Close the workbook
    workbook.close()

    to_do_lf.pack()

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
todo()
separator = ttk.Separator(root, orient='horizontal')
separator.pack(fill='x', pady=20)

def done():
    # # Create a new workbook
    # workbook = openpyxl.Workbook()

    # Open the workbook
    workbook = openpyxl.load_workbook('done.xlsx')

    # Get the active worksheet
    worksheet = workbook.active
    done_lf = LabelFrame(root, text='Done', bd=2, relief=RIDGE)
    
    # Create a canvas with a scrollbar
    canvas_done = tkinter.Canvas(done_lf, width=800, height=350)
    scrollbar_done = ttk.Scrollbar(done_lf, orient="vertical", command=canvas_done.yview)
    scrollable_frame_done = ttk.Frame(canvas_done)

    scrollable_frame_done.bind(
        "<Configure>",
        lambda e: canvas_done.configure(
            scrollregion=canvas_done.bbox("all")
        )
    )

    canvas_done.create_window((0, 0), window=scrollable_frame_done, anchor="nw")
    canvas_done.configure(yscrollcommand=scrollbar_done.set)

    def on_mousewheel_done(event):
        canvas_done.yview_scroll(int(-1*(event.delta/120)), "units")

    canvas_done.bind_all("<MouseWheel>", on_mousewheel_done)

    def bind_scroll(event):
        canvas_done.bind_all("<MouseWheel>", on_mousewheel_done)

    def unbind_scroll(event):
        canvas_done.unbind_all("<MouseWheel>")

    canvas_done.bind("<Enter>", bind_scroll)
    canvas_done.bind("<Leave>", unbind_scroll)

    # Read data from the worksheet
    i=0
    for row in worksheet.iter_rows(min_row=0, min_col=0, values_only=True):
        for j in range(0,len(row)):
            e=ttk.Label(scrollable_frame_done,width=20,text=row[j],
            anchor='center')
            e.grid(row=i,column=j)
            done=ttk.Label(scrollable_frame_done,width=20, text="Done",
            anchor='center')
            done.grid(row=0,column=3)

            done_checkbutton = Checkbutton(scrollable_frame_done)
            done_checkbutton.grid(row=i, column=j+1)
        i += 1
        
    # Close the workbook
    workbook.close()

    done_lf.pack()

    scrollbar_done.pack(side="right", fill="y")
    canvas_done.pack(side="left", fill="both", expand=True)
done()

root.mainloop()

