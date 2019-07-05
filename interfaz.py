from tkinter import *


window = Tk()
window.title("Universidad de Antioquia")
window.geometry('350x200')

lbl = Label(window, text="Materias",font=("Arial Bold", 15))
lbl.grid(column=0, row=0)

def clicked():
	lbl.configure(text="Ha seleccionado 1 cupo")

btn = Button(window, text="Seleccionar", bg="white", fg="black", command=clicked)
btn.grid(column=1, row=0)
window.mainloop()
