from tkinter import *
import openpyxl
import numpy as np

#------------------Lectura de Datos-----------------------------
excel_document = openpyxl.load_workbook('programa.xlsx')
sheet1 = excel_document['MateriasyCupos']
sheet2 = excel_document['EstudiantesMatriculados']


Sub=[]
Cupos=[]

multiple_cells1 = sheet1['A1':'A8']
for row in multiple_cells1:
    for cell in row:
        Sub.append(cell.value)


multiple_cells2 = sheet1['B1':'B8']
for row in multiple_cells2:
    for cell in row:
        Cupos.append(cell.value)


#---------------------Interfaz------------------------------------

window = Tk()
window.title("Universidad de Antioquia")
window.geometry('450x400')

def clicked():
	lbl.configure(text="Ha seleccionado 1 cupo")
	

def saveinfo():
	name_info = name.get()
	print(name_info)
	
	file=open("user.txt", "w")
	file.write(name_info)
	name_entry.delete(0,END)

	print("user", name_info, "has been regristred successfully")


nuevoscupos=[]
def selectcupos():
	nuevoscupos=Cupos - 1	
	print(nuevoscupos)


	

	
	

for i in range(len(Sub)):
	a=Sub[i]
	b=Cupos[i]

	lbl = Label(window, text=a,font=("Arial Bold", 15))
	lbl.grid(column=0, row=i)
	
	lbl1 = Label(window, text=b,font=("Arial Bold", 15))
	lbl1.grid(column=1, row=i)
	
	btn = Button(window, text="Seleccionar", bg="white", fg="black", command=selectcupos)
	btn.grid(column=2, row=i)


name = 	StringVar()
name_entry = Entry(textvariable= name, width=40)	
name_entry.place(x=1, y= 245)
name_text=Label(text= "Ingrese su nombre ",font=("Arial Bold", 15))
name_text.place(x=1, y=215)

Accept = Button(window, text = "Aceptar", bg="grey", command= saveinfo)
Accept.place(x= 150, y= 280)


window.mainloop()

