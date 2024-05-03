from tkinter import *


class PasswordManager:
    def handler(self, event):
        self.entry = self.password.get()
        self.window.destroy()

    def __init__(self):
        self.entry = ""
        self.window = Tk()
        self.window.title("Password Manager")
        self.window.config(padx=50, pady=50)

        self.label = Label(text="Access Password: ")
        self.label.grid(row=0, column=0)

        self.password = Entry(width=25)
        self.password.grid(row=0, column=1)


        self.password.bind('<Return>', self.handler)

        self.window.mainloop()


