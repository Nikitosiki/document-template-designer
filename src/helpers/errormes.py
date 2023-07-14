from tkinter import messagebox
import sys

# Отображение ошибки и завершение работы
def show_error(error_mess):
    messagebox.showerror("Error! Document-template-designer", error_mess)
    sys.exit()