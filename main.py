from tkinter import Tk
from tkinter import ttk, IntVar, CENTER

def main():
    root = Tk()
    root.geometry("800x600")

    def login_screen():
        user = IntVar()

        login_label = ttk.Label(root, text="Enter or scan your employee ID")
        login_label.pack()
        login_entry = ttk.Entry(root, textvariable=user)
        login_entry.pack(anchor=CENTER)
        login_entry.focus()
        def login():
            if user == 5632:
                print("ok")
            else:
                login_label.configure(text="wrong id")

        root.bind('<Return>', lambda event=None: login())
    login_screen()


    root.mainloop()
main()