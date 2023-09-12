import datetime
from tkinter import Tk, Button, Label, Scrollbar, Listbox, StringVar, Entry, W,E,S,N, END
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime
from sqlserver import dbConfig
import pypyodbc as pyo

date = datetime.now()
year = date.year
month = date.month
day = date.day
con = pyo.connect(**dbConfig) #Conecta a la base de datos
now = (f"{day}-{month}-{year}")

root = Tk() #Crea la ventana de la app.
cursor = con.cursor() #Para ejecutar comandos de sql en app.

class VentasDB:
    def __init__(self):
        self.con = pyo.connect(**dbConfig)
        self.cursor = con.cursor()
        print("Te conectaste exitosamente a: ", con)

    def __del__(self): #Para cerrar la conexion con la base de datos.
        self.con.close()

    def view(self):
        self.cursor.execute("SELECT id, Nombre, Empresa, Correo, Curso, Telefono, Comentario, Fecha, Seguimiento FROM cursos;")
        rows = self.cursor.fetchall()
        return rows

    def insert(self, Nombre, Empresa, Correo, Curso, Telefono, Comentario, Fecha, Seguimiento): #Definimos metodos y parametros.
        sql = ("INSERT INTO ventas.dbo.cursos(Nombre, Empresa, Correo, Curso, Telefono, Comentario, Fecha, Seguimiento) VALUES(?,?,?,?,?,?,?,?)")
        values = [Nombre, Empresa, Correo, Curso, Telefono, Comentario, Fecha, Seguimiento]
        self.cursor.execute(sql,values) #Va a ejecutar las variables con los comandos en "cursor"
        self.con.commit() #Envia los datos a la base.
        messagebox.showinfo(title="Base de datos de Ventas", message="Se ha ingresado un nuevo cliente!")

    def update(self, id, Nombre, Empresa, Correo, Curso, Telefono, Comentario, Fecha, Seguimiento):
        update_value = messagebox.askokcancel(title="Base de datos de Ventas", message="Desear modificar el cliente: " + nombre_entry.get() + "?")
        if update_value is True:
            tsql = "UPDATE ventas.dbo.cursos SET Nombre= ?, Empresa= ?, Correo= ?, Curso= ?, Telefono= ?, Comentario= ?, Fecha= ?, Seguimiento= ? WHERE id=?"
            self.cursor.execute(tsql, [Nombre, Empresa, Correo, Curso, Telefono, Comentario, Fecha, Seguimiento, id])
            self.con.commit()
            messagebox.showinfo(title="Base de datos de Ventas", message="Base de datos ha sido actualizada!")
        else:
            messagebox.showinfo(title="Base de datos de Ventas", message="No se ha modificado ningun dato")

    def delete(self, id):  # Solo es necesario el ID para eliminar un valor.
        delete_value = messagebox.askokcancel(title="Base de datos de Ventas", message="Desear eliminar el cliente: " + nombre_entry.get() + "?")
        if delete_value is True:
            delquery = "DELETE FROM ventas.dbo.cursos WHERE id = ?"
            self.cursor.execute(delquery, [id])
            self.con.commit()
            messagebox.showinfo(title="Base de datos de Ventas", message="Se ha eliminado el cliente!")
        else:
            messagebox.showinfo(title="Base de datos de Ventas", message="No se ha eliminado ningun dato")

baseDatos = VentasDB() #Se hace referencia a la clase para poder interactuar con ella.

#Sacar datos de la list box y poder usarlos.
def get_row(event): #Los datos dentro de la caja de datos.
    global selected_tuple #Se utiliza global porque los datos van a ser editados fuera de la funcion.
    index = list_bx.curselection()[0] #Usada para darnos la lista de clientes, devuelve los indices en un tuple.
    selected_tuple = list_bx.get(index) #Se le asigna un valor a la variable para referencia del metodo y devuelve el indice de un valor.
    nombre_entry.delete(0, "end") #Si se selecciona una linea, la borra y anade la otra seleccionada.
    nombre_entry.insert("end", selected_tuple[1])
    empresa_entry.delete(0,"end")
    empresa_entry.insert("end", selected_tuple[2])
    correo_entry.delete(0,"end")
    correo_entry.insert("end", selected_tuple[3])
    curso_entry.delete(0,"end")
    curso_entry.insert("end", selected_tuple[4])
    tel_entry.delete(0,"end")
    tel_entry.insert("end", selected_tuple[5])
    coment_entry.delete(0,"end")
    coment_entry.insert("end", selected_tuple[6])
    fecha_text.insert("end", selected_tuple[7])
    fechaSiguiente_entry.delete(0,"end")
    fechaSiguiente_entry.insert("end",selected_tuple[8])




def view_clients():
    list_bx.delete(0,"end") #Limpia la caja de clientes
    for row in baseDatos.view():
        list_bx.insert("end",row) #Mete todos los datos encontrados en la base de datos hacia la caja.

def add_clients():
    baseDatos.insert(nombre_text.get(),empresa_text.get(),correo_text.get(), curso_text.get(), tel_text.get(),
                     coment_text.get(), now, fechaSiguiente_text.get())
    list_bx.delete(0,"end")
    list_bx.insert("end", (nombre_text.get(),empresa_text.get(),correo_text.get(), curso_text.get(), tel_text.get(),
                           coment_text.get(), now, fechaSiguiente_text.get()))
    nombre_entry.delete(0,"end")
    empresa_entry.delete(0, "end")
    correo_entry.delete(0, "end")
    curso_entry.delete(0, "end")
    tel_entry.delete(0, "end")
    coment_entry.delete(0, "end")
    fechaSiguiente_entry.delete(0,"end")
    con.commit()

def delete_clients():
    baseDatos.delete(selected_tuple[0]) #Toma el indice seleccionado.
    con.commit()

def clear_screen():
    list_bx.delete(0,"end")
    nombre_entry.delete(0,"end")
    empresa_entry.delete(0, "end")
    correo_entry.delete(0, "end")
    curso_entry.delete(0, "end")
    tel_entry.delete(0, "end")
    coment_entry.delete(0, "end")
    fechaSiguiente_entry.delete(0, "end")

def update_clients():
    baseDatos.update(selected_tuple[0],nombre_text.get(),empresa_text.get(),correo_text.get(),
                     curso_text.get(), tel_text.get(), coment_text.get(), now, fechaSiguiente_text.get())
    list_bx.delete(0, "end")
    nombre_entry.delete(0, "end")
    empresa_entry.delete(0, "end")
    correo_entry.delete(0, "end")
    curso_entry.delete(0, "end")
    tel_entry.delete(0, "end")
    coment_entry.delete(0, "end")
    fechaSiguiente_entry.delete(0, "end")
    con.commit()

def exit():
    done = baseDatos
    cerrar= messagebox.askokcancel("Salir","Desea cerrar el programa?")
    if cerrar is True:
        root.destroy()



#Titulo y fondo
root.title("Seguimiento de Ventas") #Titulo del app
root.configure(background= "white") #Fondo
root.geometry("950x550") #Tamano
root.resizable(width=False, height=False) #Que no se le pueda cambiar el tamano

#Botones (Tienen que estar referidos a "root" ya que es la app)
#Boton de nombre
nombre_label = ttk.Label(root, text = "Nombre", background="white", font=("TkDefaultFont", 12))
nombre_label.grid(row=0, column=0, sticky=W)
nombre_text= StringVar() #Va a ser usados para poder enviarlos a la base de datos.
nombre_entry= ttk.Entry(root, width=20, textvariable=nombre_text)
nombre_entry.grid(row=0,column=1, sticky=W)

#Boton de Empresa
empresa_label = ttk.Label(root, text= "Empresa", background="white", font=("TkDefaultFont", 12))
empresa_label.grid(row=0,column=2, sticky=E)
empresa_text = StringVar()
empresa_entry = ttk.Entry(root, width=20, textvariable=empresa_text)
empresa_entry.grid(row=0, column=3, sticky=W)

#Boton de Correo
correo_label = ttk.Label(root, text= "Correo", background="white", font=("TkDefaultFont", 12))
correo_label.grid(row=0,column=4, sticky=W)
correo_text = StringVar()
correo_entry = ttk.Entry(root, width=20, textvariable=correo_text)
correo_entry.grid(row= 0, column=5, sticky=W)

#Boton de Nombre de Curso
curso_label = ttk.Label(root, text="Curso", background="white", font=("TkDefaultFont", 12))
curso_label.grid(row=0, column=6, sticky=W)
curso_text = StringVar()
curso_entry = ttk.Entry(root, width=20, textvariable=curso_text)
curso_entry.grid(row=0, column=7, sticky=W)

#Boton de Telefono
tel_label = ttk.Label(root, text= "Telefono", background="white", font=("TkDefaultFont", 12))
tel_label.grid(row=1, column=0, sticky=W)
tel_text = StringVar()
tel_entry = ttk.Entry(root, width=20, textvariable=tel_text)
tel_entry.grid(row=1, column=1, sticky=W)

#Boton de comentario
coment_label = ttk.Label(root, text= "Comentarios", background="white", font=("TkDefaultFont", 12))
coment_label.grid(row=1, column=2, sticky=W)
coment_text = StringVar()
coment_entry = ttk.Entry(root, width=20, textvariable=coment_text)
coment_entry.grid(row=1, column=3, sticky=W)

#Fecha
fecha_label = ttk.Label(root, text= "Fecha", background="white", font=("TkDefaultFont", 12))
fecha_label.grid(row=1, column=6, sticky=W)
fecha_text = now
fecha_num = ttk.Label(root, text=str(now), background="white", font=("TkDefaultFont", 12))
fecha_num.grid(row=1, column=7, sticky=W)

#Fecha siguiente contacto
fechaSiguiente_label = ttk.Label(root, text= "Seguimiento", background="white", font=("TkDefaultFont", 12))
fechaSiguiente_label.grid(row=1, column=4, sticky=W)
fechaSiguiente_text = StringVar()
fechaSiguiente_entry = ttk.Entry(root, width=20, textvariable=fechaSiguiente_text)
fechaSiguiente_entry.grid(row=1, column=5, sticky=W)

#Boton para anadir a la base de datos
add_btn = Button(root, text="Agregar", bg="blue", fg="white", font="helvetica 10 bold", command=add_clients)
add_btn.grid(row=4, column=0, sticky=W)

#Espacio para desplegar datos
list_bx = Listbox(root, height=14, width=40, font="Helvetica 13", bg="light blue")
list_bx.grid(row=5, column=0, columnspan=14, sticky=W+E, pady=40, padx=15) #Columnspam para abarcar mas columnas.
list_bx.bind("<<ListboxSelect>>", get_row) #Unir el row que seleccionamos con la clase.

# #Espacio para desplegar datos
# list_bx2 = Listbox(root, height=16, width=45, font="Helvetica 13", bg="light blue")
# list_bx2.grid(row=5, column=5, columnspan=14, sticky=W, pady=30, padx=15) #Columnspam para abarcar mas columnas.
# list_bx2.bind("<<ListboxSelect>>", get_row) #Unir el row que seleccionamos con la clase.

#Scroll bar
#scroll_bar= Scrollbar(root)
#scroll_bar.grid(row=1, column=12, rowspan=14, sticky=W)
#list_bx.configure(yscrollcommand=scroll_bar.set)
#scroll_bar.configure(command=list_bx.yview())

#Boton de Modificar
modify_btn= Button(root, text="Modificar Cliente", bg="black", fg= "white", font="helvestica 10 bold", command=update_clients)
modify_btn.grid(row=13, column=1)

#Boton de Borrar
delete_btn= Button(root,text="Eliminar Cliente", bg="black", fg="white", font="helvestica 10 bold", command=delete_clients)
delete_btn.grid(row=13, column=3)

#Boton de Ver
view_btn= Button(root, text= "Ver Clientes", bg= "black", fg="white", font="helvestica 10 bold", command= view_clients)
view_btn.grid(row=13, column=5)

#Boton para limpiar pantalla
clear_btn= Button(root, text="Limpiar Datos", bg="black", fg="white", font="helvestica 10 bold", command= clear_screen)
clear_btn.grid(row=13, column=7)

#Boton para cerrar app
exit_btn= Button(root, text= "Cerrar Applicacion", bg="red", fg="white", font="helvestica 10 bold", command=root.destroy)
exit_btn.grid(row=13, column=9)



root.mainloop() #Para poder correr la app.