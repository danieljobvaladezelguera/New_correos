#Librerías para la creación visual del programa
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox, scrolledtext,font, Menu
from tkinter.ttk import Progressbar, Combobox
from openpyxl import load_workbook #Importamos libreria para el uso de excel
from smtplib import SMTP #Libreria dentro de python para el envio de correos
import webbrowser as web #Librería para poner una url en el programa
from socket import gaierror #Librería con el fin de buscar errores
from pathlib import Path #Usar para revisar archivos
import Pmw
from os import makedirs, path # Crear una carpeta para guardar información

#https://myaccount.google.com/lesssecureapps?pli=1&rapt=AEjHL4PDAcXeJzQCiyuIDVjmt4tfaMySFl40ei6dVTFTvz67ZJJl5s9nZV18NkNjNRd9jJ06jljezI9sUpZK8na8IUGjK6omA2_tjCXlK0doz2pBF2O7S3s
#https://myaccount.google.com/apppasswords?utm_source=google-account&utm_medium=myaccountsecurity&utm_campaign=tsv-settings&rapt=AEjHL4Mk5lLF8MX793jG6SdOJKKflJoFb5rDbg3TutbXD3LAXWwxUTF5Q5gKJZZqw4geHCcZTzVym_Gr33b-nhnP67f1jVtWir6glV4dKvbg3EqS7IhLniM

#Variables que no cambiarán en todo el correo
n_link = "https://myaccount.google.com/apppasswords?utm_source=google-account&utm_medium=myaccountsecurity&utm_campaign=tsv-settings&rapt=AEjHL4Mk5lLF8MX793jG6SdOJKKflJoFb5rDbg3TutbXD3LAXWwxUTF5Q5gKJZZqw4geHCcZTzVym_Gr33b-nhnP67f1jVtWir6glV4dKvbg3EqS7IhLniM"    
pasw = "C:/UsPwd"
arch = "C:/UsPwd/uspwd.txt"
asun = "Asunto"
salu = "Saludo"
body = "Cuerpo del Correo"
mail = "Correo Electrónico"
mail_port = 'smtp.gmail.com'
pasword = "Contraseña"

#Inicio de funciones
#link para que visiten la página web para la contraseña de google
def link():
    web.open(n_link)
#Función para que cambie de usuario
def cambio_usuario_inicio():
    arch = "C:/UsPwd/uspwd.txt"

    def on_entry_mail(event):
        while entry_mail.get() == mail:
            entry_mail.delete(0, "end")

    def on_entry_pass(event):
        if entry_password.get() == pasword:
            entry_password.delete(0, "end")

    def escribir():
            try:
                ag1 = entry_mail.get()
                ag2 = entry_password.get()
                pos = valores.index(combo.get())
                web_smtp = 'smtp.'+servidor[pos]
                if (ag1 != '' and ag2 != '' and pos != ""):
                    uspwd = open(arch,"w")
                    uspwd.write(ag1)#correo
                    uspwd.write(",")
                    uspwd.write(ag2)#contraseña
                    uspwd.write(",")
                    uspwd.write(web_smtp)#servidor
                    uspwd.write(",")
                    uspwd.write(500)#valor de cantidad de correos
                    uspwd.write(",")
                    messagebox.showinfo("Aviso", "Datos cargados correctamente")
                    uspwd.close()
                    correo.destroy()
                else:
                    messagebox.showinfo("Aviso", "Faltan datos por llenar")
            except Exception as e:
                messagebox.showerror("Error", f"Ha ocurrido un error: {e}")
                pass

    uspwd = open(arch, "r")
    uspwdread = uspwd.readlines()
    cantidad = len(uspwdread)
    if(cantidad == 0):
        valores=["1&1 IONOS",'AOL','GMAIL','ICLOUD MAIL','OFFICE 365',
            'ORANGE','OUTLOOK','YAHOO','ZOHO MAIL']
        servidor=['1and1.com','aol.com','gmail.com','gmx.com','mail.me.com',
            'office365.com','orange.net','live.com','mail.yahoo.com','zoho.com']
        correo = Tk()
        correo.geometry("450x250")
        correo.title("Cambio de Usuario")
        Pmw.initialise(correo)

        entry_mail =  Entry(correo, width = 45)
        entry_mail.insert('insert',mail)
        entry_mail.bind("<Button-1>", on_entry_mail)

        entry_password =  Entry(correo, width = 45)
        entry_password.insert('insert',pasword)
        entry_password.bind("<Button-1>", on_entry_pass)

        tooltip = Pmw.Balloon(correo)
        texto = Label(correo, text = "Selecciona sevidor de correo")
        combo = Combobox(correo,state='readonly',values= valores)

        btn_envio =  Button(correo,
                            text ="Guardar",
                            width = 45, 
                            height = 1,
                            bg = "gray", 
                            command = escribir)
        btn_insert = Button(correo, 
                            text = "Link para saber tu contraseña",
                            width = 45,
                            height= 1,
                            bg= "gray",
                            command = link)    
        
        tooltip.bind(texto, "Si tiene duda al elegir, es en dónde abre su correo"+
        "\nNo hablo del navegador, hablo de la página web donde la abre")
        entry_mail.pack(padx=10, pady=5)
        entry_password.pack(padx=10, pady=5)
        texto.pack(padx=10, pady=5)
        combo.pack(padx=10, pady=5)
        btn_envio.pack(padx=10, pady=5)
        btn_insert.pack(padx=10, pady=5)             
        correo.mainloop()
    else:
        uspwd.close()
    uspwd.close()
#Ambas funciones son las mismas, pero para mejorarla en un futuro se necesitan unos detalles
def cambio_usuario_button():
    arch = "C:/UsPwd/uspwd.txt"

    def on_entry_mail(event):
        while entry_mail.get() == mail:
            entry_mail.delete(0, "end")

    def on_entry_pass(event):
        if entry_password.get() == pasword:
            entry_password.delete(0, "end")

    def escribir():
            try:
                ag1 = entry_mail.get()
                ag2 = entry_password.get()
                pos = valores.index(combo.get())
                web_smtp = 'smtp.'+servidor[pos]
                if (ag1 != '' and ag2 != '' and pos != ""):
                    uspwd = open(arch,"w")
                    uspwd.write(ag1)#correo
                    uspwd.write(",")
                    uspwd.write(ag2)#contraseña
                    uspwd.write(",")
                    uspwd.write(web_smtp)#servidor
                    uspwd.write(",")
                    uspwd.write(500)#valor de cantidad de correos
                    uspwd.write(",")
                    messagebox.showinfo("Aviso", "Datos cargados correctamente")
                    uspwd.close()
                    correo.destroy()
                else:
                    messagebox.showinfo("Aviso", "Faltan datos por llenar")
            except Exception as e:
                messagebox.showerror("Error", f"Ha ocurrido un error: {e}")
                pass

    uspwd = open(arch, "r")
    uspwdread = uspwd.readlines()
    cantidad = len(uspwdread)
    valores=["1&1 IONOS",'AOL','GMAIL','ICLOUD MAIL','OFFICE 365',
            'ORANGE','OUTLOOK','YAHOO','ZOHO MAIL']
    servidor=['1and1.com','aol.com','gmail.com','gmx.com','mail.me.com',
            'office365.com','orange.net','live.com','mail.yahoo.com','zoho.com']
    correo = Tk()
    correo.geometry("450x250")
    correo.title("Cambio de Usuario")
    Pmw.initialise(correo)

    entry_mail =  Entry(correo, width = 45)
    entry_mail.insert('insert',mail)
    entry_mail.bind("<Button-1>", on_entry_mail)

    entry_password =  Entry(correo, width = 45)
    entry_password.insert('insert',pasword)
    entry_password.bind("<Button-1>", on_entry_pass)

    tooltip = Pmw.Balloon(correo)
    texto = Label(correo, text = "Selecciona sevidor de correo")
    combo = Combobox(correo,state='readonly',values= valores)

    btn_envio =  Button(correo,
                            text ="Guardar",
                            width = 45, 
                            height = 1,
                            bg = "gray", 
                            command = escribir)
    btn_insert = Button(correo, 
                            text = "Link para saber tu contraseña",
                            width = 45,
                            height= 1,
                            bg= "gray",
                            command = link)    
        
    tooltip.bind(texto, "Si tiene duda al elegir, es en dónde abre su correo"+
        "\nNo hablo del navegador, hablo de la página web donde la abre")
    entry_mail.pack(padx=10, pady=5)
    entry_password.pack(padx=10, pady=5)
    texto.pack(padx=10, pady=5)
    combo.pack(padx=10, pady=5)
    btn_envio.pack(padx=10, pady=5)
    btn_insert.pack(padx=10, pady=5)             
    correo.mainloop()
#Función para escribir los correos
def enviar():
    arch = "C:/UsPwd/uspwd.txt"
    l=h=cont=tot_cor=0
    a = "@"
    lis_er_str = ""
    lis_er = []

    def cant_corr(nombre, new_cantidad):
        # Leer el archivo y cargar los datos en una lista de líneas
        with open(arch, "r") as archivo:
            lineas = archivo.readlines()
    
    # Modificar la línea correspondiente
        for i in range(len(lineas)):
        # Separar cada línea por comas
            columnas = lineas[i].strip().split(",")
            if columnas[0] == nombre:  # Buscar el usuario
                columnas[3] = str(new_cantidad)  # Actualizar la cantidad de correos
                lineas[i] = ",".join(columnas) + "\n"  # Reensamblar la línea
    # Escribir el archivo nuevamente con la modificación
        with open(arch, "w") as archivo:
            archivo.writelines(lineas[i])
        pass

    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    name_archivo= path.basename(archivo)
    dato = open(arch, "r")
    info = dato.read()
    array = info.split(",")
    correo = array[0]#Correo
    contra = array[1]#Contraseña
    server_smtp = array[2]#Servidor Smtp
    cant_correo = array[3]#Cantidad de correos x día
    dato.close()
    cant_correo_2 = int(float(cant_correo))#conversión de correos
    
    if archivo:
        libro = load_workbook(filename=archivo)
        hoja = libro.active
        rango = (hoja['A1':'E500']) #Revisar cómo contar los datos dentro del excel antes de pasarlos a un 
        nombre= []
        lis_co= []
        pdf= []
        extra= []
        id = []
        for j in range(0,5):
            for k in range(0,cant_correo_2):
                while(j == 0):#Busca el id
                    if(rango[k][j].value != None and
                       rango[k][j].value != 'id'):
                        id.append(rango[k][j].value)
                    break
                while(j == 1):#busca el nombre
                    if(rango[k][j].value != 'nombre' and
                        rango[k][j].value != None):
                        nombre.append(rango[k][j].value)
                    break
                while(j == 2):#Busca el correo
                    if(rango[k][j].value != 'correo' and
                        rango[k][j].value != None):
                        lis_co.append(rango[k][j].value)
                    break
                while(j == 3):#Busca el archivo
                    if(rango[k][j].value != 'archivo' and
                        rango[k][j].value != None):
                        pdf.append(rango[k][j].value)
                    break
                while(j == 4):#Busca algo extra
                    if(rango[k][j].value != 'extra' and
                        rango[k][j].value != None):
                        extra.append(str(rango[k][j].value))
                    break
        
        messagebox.showinfo("Aviso",f"Datos del Archivo {name_archivo} cargados correctamente")
        #Revisión de los correos
        for i in lis_co:    
            try:
                duda = i.rfind(a)#Se busca un caracter en específico, el @
                if(duda == -1): #Se usa un -1 porque significa q no encontró lo que buscaba
                    cont = cont + 1   
                    lis_er.append(i)
                for h in lis_er:#Crea la lista para guardar los correos malos
                    lis_er_str = lis_er_str + h + ", "
                asf = lis_er_str
            except Exception as AtributeError:
                messagebox.showerror("ERROR", "Tomó un Excel erróneo  ó \n"+
                "no está lleno correctamente.")
                break
            except  Exception as e:
                messagebox.showerror("Error", f"Ha ocurrido un error: {e}")

        if(len(lis_er)> 0):#Revisión de correos con problemas
            messagebox.showinfo("Aviso", "Los siguientes correos tiene problemas, "
                                +"favor de revisarlos: "
                                + "\n"+ asf)
        else:     #Si no hay problemas, inicia la toma de datos para el envío           
            new_asunto = Entry_asun.get() #Asunto
            new_saludo = Entry_salu.get() #Saludo
            new_cuerpo = Entry_body_mail.get("1.0", "end-1c") #Cuerpo del correo

            num_cor = len(id)

            #Aquí se toman los datos del archivo txt para la cantidad de correos, usuario y contraseña del correo
            if(new_asunto != None or new_asunto != '' or new_saludo != None or new_saludo != ''):
                message = 'Subject: {}\n\n{}'.format(new_asunto, new_saludo + ": ")
                dato.close()
                try:#Mientras funcione, se enviarán los correos
                    server = SMTP(server_smtp,587) #toma los puertos para poder funcionar, por defecto se establecen cómo 587 y smtp.gmail.com
                    server.starttls()#manda un helo para conectar al servidor
                    server.login(correo, contra)#toma el usuario y contraseña antes guardados
                    pregunta = messagebox.askyesnocancel("Aviso","Inicio de envio de correos "
                                + "\n Total de correos a enviar: " + str(len(id))
                                + "\n\n De: " + correo
                                + "\n\n Para: " + lis_co[0]
                                + "\n\n"+ new_saludo + nombre[0]
                                + "\n\n"+ new_cuerpo
                                + "\n\n" + extra[0]
                                + "\n\n" + pdf[0]
                                + "\n\n ¿Se encuentra el correo en orden?")                                
                    if pregunta:
                        messagebox.showinfo("Aviso", "Inicio del envío de correos")
                        while(l != num_cor and tot_cor != cant_correo_2):
                            server.sendmail(correo, 
                                            lis_co[l], 
                                            message + nombre[l]
                                            +"\n"+ new_cuerpo
                                            + "\n\n Archivo adjunto: " + extra[l]
                                            + "\n\n " + pdf[l]) 
                            l = l + 1
                            tot_cor = tot_cor + 1
                        else:
                            messagebox.showinfo("Aviso", "Envio de correos finalizado")
                            new_cantidad = cant_correo_2 - tot_cor
                            cant_corr(correo, new_cantidad)
                            dato.close
                    elif "no" or "cancel":
                        messagebox.Message("Envío de correos cancelado")
                        new_cantidad = cant_correo_2 - tot_cor
                        cant_corr(correo, new_cantidad)
                        dato.close()
                except gaierror:
                    messagebox.showerror("Error de conexión", 
                                "No hay conexión a Internet." + 
                                "\nPor favor, verifica tu conexión e inténtalo de nuevo.")
                    new_cantidad = cant_correo_2 - tot_cor
                    cant_corr(correo, new_cantidad)
                    dato.close()
                except Exception as e:
                    messagebox.showerror("Error", f"Ha ocurrido un error: {e}")
                    new_cantidad = cant_correo_2 - tot_cor
                    cant_corr(correo, new_cantidad)
                    dato.close()
            else:
                messagebox.showerror("ALTO", "Datos de asunto y/o saludo deben de llenarse")
    else:
        messagebox.showinfo("Aviso", "Favor de seleccionar un archivo Excel")
#Función para borrar el texto explicativo
def on_entry_asun(event):
    if Entry_asun.get() == asun:
            Entry_asun.delete(0, "end")
#Función para borrar el texto explicativo
def on_entry_salu(event):
    if Entry_salu.get() == salu:
            Entry_salu.delete(0, "end")
#Función para borrar el texto explicativo
def on_entry_body(event):
    if Entry_body_mail.get("1.0", "end-1c") == body:
            Entry_body_mail.delete("1.0", "end-1c")
#Salir del programa
def salir():
    ven_correos.destroy()
#Conteo de la cantidad de correos que pueden salir
def cantidad_correos():
    dato = open(arch, "r")
    info = dato.read()
    array = info.split(",")
    cantidad = array[3]
    info = Tk()
    info.geometry("350x75")
    info.title("Cantidad de correoss")
    info.config(background='gray15', cursor="")
    
    if float(cantidad) > 0:
        texto = "Puede enviar una cantidad de " + cantidad + " correos"
        info_texto = Label(info, text= texto, width=60, background='gray15',
                font=font.Font(family="Arial",size=17),
                foreground="white")
    else:
        texto ="Ya no puede enviar correos. \n Favor de esperar hasta el día siguiente."
        info_texto = Label(info, text= texto, width=60, background='gray15',
                font=font.Font(family="Arial",size=17),
                foreground="white")
    
    

    info_texto.pack(padx=10, pady=10)
    info.mainloop()
    pass
#Inicio del programa

#Crea una carpeta y un archivo
Ruta = Path(pasw)
makedirs(Ruta, exist_ok=True)
uspwd_file = Path(arch)
if not uspwd_file.exists():
    with open(arch, "w") as file:
        file.write("")

#Leer usuario y contraseña del archivo de texto
cambio_usuario_inicio()

uspwd = open(arch, "r")
uspwdread = uspwd.readlines()
cantidad = len(uspwdread)
if (cantidad == 1):
    dato = open(arch, "r")
    info = dato.read()
    array = info.split(",")
    correo = array[0]
else:
    correo = 'Sin usuario'
usuario = 'Usuario: ' + correo

#Inicio del programa de envío de correos
ven_correos = Tk()
ven_correos.geometry("640x480")
ven_correos.title("Envío de Correos Masivos")
ven_correos.config(background='gray15', cursor="")
menu = Menu(ven_correos)
new_item = Menu(menu, border='50', tearoff=False)
ven_correos.config(menu=menu)

#Items dentro del menú desplegable
menu.add_cascade(label='Configuración', menu=new_item)
new_item.add_command(label='Cambio de usuario', command=cambio_usuario_button)
new_item.add_command(label='Cantidad de correos para enviar', command=cantidad_correos)
new_item.add_separator()
new_item.add_command(label='Envío de correos', command=enviar)
new_item.add_separator()
new_item.add_command(label='Salir', command=salir)

#Mostrar el usuario a turno
usuario = Label(ven_correos, text = usuario, width=60, background='gray15',
                font=font.Font(family="Arial",size=15),
                foreground="White")
#Lugar donde entrará la información
Entry_asun =  Entry(ven_correos, text= "", width=60, background='gray15',
                font=font.Font(family="Arial",size=17),
                foreground="white", insertbackground='white')
Entry_asun.insert(0, asun)
Entry_asun.bind("<Button-1>", on_entry_asun)

Entry_salu =  Entry(ven_correos, text= "", width=60, background='gray15',
                font=font.Font(family="Arial",size=17),
                foreground="white", insertbackground='white')
Entry_salu.insert(0, salu)
Entry_salu.bind("<Button-1>", on_entry_salu)

Entry_body_mail= scrolledtext.ScrolledText(ven_correos,width=60, height=30, background='gray15',
                font=font.Font(family="Arial",size=17),
                foreground="White", insertbackground='white')
Entry_body_mail.insert('insert', body)
Entry_body_mail.bind('<Button-1>', on_entry_body)
#Posicionamiento de los datos de arriba
usuario.pack(padx=10, pady=10)
Entry_asun.pack(padx=10, pady=10)
Entry_salu.pack(padx=10, pady=10)
Entry_body_mail.pack(padx=10, pady=10)

ven_correos.mainloop()