#Librería para la creación visual del programa
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox, scrolledtext,font, Menu, ttk
from openpyxl import load_workbook #Importamos libreria para el uso de excel
from smtplib import SMTP #Libreria dentro de python para el envio de correos
import webbrowser as web #Librería para poner una url en el programa
from socket import gaierror #Librería con el fin de buscar errores
from pathlib import Path #Usar para revisar archivos
from os import makedirs # Crear una carpeta para guardar información

#https://myaccount.google.com/lesssecureapps?pli=1&rapt=AEjHL4PDAcXeJzQCiyuIDVjmt4tfaMySFl40ei6dVTFTvz67ZJJl5s9nZV18NkNjNRd9jJ06jljezI9sUpZK8na8IUGjK6omA2_tjCXlK0doz2pBF2O7S3s
#https://myaccount.google.com/apppasswords?utm_source=google-account&utm_medium=myaccountsecurity&utm_campaign=tsv-settings&rapt=AEjHL4Mk5lLF8MX793jG6SdOJKKflJoFb5rDbg3TutbXD3LAXWwxUTF5Q5gKJZZqw4geHCcZTzVym_Gr33b-nhnP67f1jVtWir6glV4dKvbg3EqS7IhLniM

#Zona de 
n_link = "https://myaccount.google.com/apppasswords?utm_source=google-account&utm_medium=myaccountsecurity&utm_campaign=tsv-settings&rapt=AEjHL4Mk5lLF8MX793jG6SdOJKKflJoFb5rDbg3TutbXD3LAXWwxUTF5Q5gKJZZqw4geHCcZTzVym_Gr33b-nhnP67f1jVtWir6glV4dKvbg3EqS7IhLniM"    
pasw = "C:/UsPwd"
arch = "C:/UsPwd/uspwd.txt"
asun = "Zona del Asunto ..."
salu = "Zona del Saludo ..."
body = "Zona del Cuerpo del Correo ..."
mail = "Insertar Correo Electrónico"
pasword = "Insertar Contraseña"

#Inicio de funciones
#Función para sobre/escribir el correo y la contraseña
def escribir(entry_mail,  entry_password, user, pwd):
    res = messagebox.askyesnocancel(title="Aviso", message="Se sobre escribiran los datos\n"
                                    +"Desea seguir")
    uspwd = open(arch, "r")
    uspwdread = uspwd.readlines()
    cantidad = len(uspwdread)
    if cantidad == 0:
        if res:
            ag1 = entry_mail.get()
            ag2 = entry_password.get()
            user["text"] = ag1
            pwd["text"] = ag2
            nuser = str(ag1)
            npwd = str(ag2)
            uspwd = open(arch,"w")
            uspwd.write(nuser)
            uspwd.write(",")
            uspwd.write(npwd)
            messagebox.showinfo("Aviso", "Datos cargados correctamente")
            uspwd.close()
            pass
        elif "cancel":
            return
        else:
            pass
        uspwd.destroy()
#link para que visiten la página web para la contraseña de google
def link():    
    web.open(n_link)
#Función para que se pueda agregar otro tipo de correo que no sea el de google
def cambio_server():
    def seleccion():
        select = combo.get()
        if(select != ''):
            pos = valores.index(select)
            num_port = puerto[pos]
            web_smtp = 'smtp.'+servidor[pos]
            messagebox.showinfo('Aviso', str(num_port)+ 
                                ' ' + web_smtp +' ' + select)
            w_user.destroy()
        else:
            messagebox.showinfo('Aviso', 'No ha seleccionado su correo')
        pass

    def avanzado():
        url_server.pack()
        entry_url.pack()
        port.pack()
        entry_port.pack()
        combo.configure(state='disable')
        select_correo.config(state='disable')
        pass

    tam= '300x250'
    valores=["1&1 IONOS",'AOL','GMAIL','ICLOUD MAIL','OFFICE 365',
             'ORANGE','OUTLOOK','YAHOO','ZOHO MAIL']
    puerto=[587,465,465,587,587,587,465,587,465,465]
    servidor=['1and1.com','aol.com','gmail.com','gmx.com','mail.me.com','office365.com',
              'orange.net','live.com','mail.yahoo.com','zoho.com']
    #link_contraseña=[]
    w_user = Tk()
    w_user.title('Cambio de servidor')
    w_user.geometry(tam)
    select_correo = Button(w_user, text="Seleccionar", command=seleccion)
    combo = ttk.Combobox(w_user,state='readonly',values= valores)

    url_server =  Label(w_user, text = "URL del servidor SMTP: ")
    user_email =  Label(w_user, text = "Usuario de correo: ")
    pwd_email =  Label(w_user, text = "Contraseña del correo: ")
    port =  Label(w_user, text = "Puerto de entrada: ")
    entry_url =  Entry(w_user, text = "", width = 45)
    entry_us_ =  Entry(w_user, text = "", width = 45)
    entry_pwd =  Entry(w_user, text = "", width = 45)
    entry_port =  Entry(w_user, text = "", width = 15) 
    cambio_ser =  Button(w_user, text='Opciones avanzadas', command=avanzado)

    combo.pack()
    select_correo.pack()
    cambio_ser.pack()
    pass
#Función para que cambie de usuario 
def cambio_usuario():
        def on_entry_mail(event):
            if entry_mail.get() == mail:
                entry_mail.delete(0, "end")

        def on_entry_pass(event):
            if entry_password.get() == pasword:
                    entry_password.delete(0, "end")

        def escribir():
            res = messagebox.askyesnocancel(title="Aviso", message="Se sobre escribiran los datos\n"
                                    +"Desea seguir")
            if cantidad == 0:
                if res:
                    ag1 = entry_mail.get()
                    ag2 = entry_password.get()
                    uspwd = open(arch,"w")
                    uspwd.write(ag1)
                    uspwd.write(",")
                    uspwd.write(ag2)
                    messagebox.showinfo("Aviso", "Datos cargados correctamente")
                    uspwd.close()
                    pass
                elif "cancel":
                    correo.destroy()
                    return
                else:
                    pass
                correo.destroy()

        uspwd = open(arch, "r")
        uspwdread = uspwd.readlines()
        cantidad = len(uspwdread)
        if(cantidad != 0):
            cambio_user = messagebox.askyesno("Aviso", "Ya hay datos dentro del usuario. \n"+                                              
                                "¿Desea cambiar de usuario?")
            if(cambio_user):
                open(arch, "w+")
                correo = Tk()
                correo.geometry("450x250")
                correo.title("Carga de datos")

                entry_mail =  Entry(correo, width = 45)
                entry_mail.insert('insert',mail)
                entry_mail.bind("<Button-1>", on_entry_mail)

                entry_password =  Entry(correo, width = 45)
                entry_password.insert('insert',pasword)
                entry_password.bind("<Button-1>", on_entry_pass)

                botton1 =  Button(correo,
                            text ="Guardar",
                            width = 45, 
                            height = 1,
                            bg = "gray", 
                            command = escribir)
                botton2 = Button(correo, 
                            text = "Link para saber tu contraseña",
                            width = 45,
                            height= 1,
                            bg= "green",
                            command = link)    
                
                entry_mail.pack(padx=10, pady=5)
                entry_password.pack(padx=10, pady=5)
                botton1.pack(padx=10, pady=5)
                botton2.pack(padx=10, pady=5)            
                correo.mainloop()                
            else:
                uspwd.close()
        uspwd.close()
    

#Función para escribir los correos
def enviar():
    l=0
    h=0
    cont=0
    tot_cor = 0
    puerto = 587
    link_smtp ='smtp.gmail.com'
    a = "@"
    lis_er_str = ""
    lis_er = []
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])

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
            for k in range(0,500):
                while(j == 0):
                    if(rango[k][j].value != None and
                       rango[k][j].value != 'id'):
                        id.append(rango[k][j].value)
                    break
                while(j == 1):
                    if(rango[k][j].value != 'nombre' and
                        rango[k][j].value != None):
                        nombre.append(rango[k][j].value)
                    break
                while(j == 2):
                    if(rango[k][j].value != 'correo' and
                        rango[k][j].value != None):
                        lis_co.append(rango[k][j].value)
                    break
                while(j == 3):
                    if(rango[k][j].value != 'archivo' and
                        rango[k][j].value != None):
                        pdf.append(rango[k][j].value)
                    break
                while(j == 4):
                    if(rango[k][j].value != 'extra' and
                        rango[k][j].value != None):
                        extra.append(str(rango[k][j].value))
                    break
        messagebox.showinfo("Aviso", "Datos del Archivo Excel cargados correctamente")
        #Revisión de los correos
        for i in lis_co:    
            duda = i.rfind(a)
            if(duda == -1): #Se usa un -1 porque significa q no encontró lo que buscaba
                cont = cont + 1   
                lis_er.append(i)
        for h in lis_er:
            lis_er_str = lis_er_str + h + ", "
        asf = lis_er_str
        
        if(len(lis_er)> 0):
            messagebox.showinfo("Aviso", "Los siguientes correos tiene problemas, "
                                +"favor de revisarlos: "
                                + "\n"+ asf)
        else:                
            new_asunto = Entry_asun.get() #Asunto
            new_saludo = Entry_salu.get() #Saludo
            new_cuerpo = Entry_body_mail.get("1.0", "end-1c") #Cuerpo del correo

            num_cor = len(id)

            #Aquí se toman los datos del archivo txt para el usuario y contraseña del correo
            if(new_asunto != None or new_asunto != '' or new_saludo != None or new_saludo != ''):
                dato = open(arch, "r")
                info = dato.read()
                array = info.split(",")
                correo = array[0]
                contra = array[1]
                message = 'Subject: {}\n\n{}'.format(new_asunto, new_saludo + ": ")
                dato.close()
                try:
                    server = SMTP(link_smtp, puerto)
                    server.starttls()
                    server.login(correo,contra)
                    pregunta = messagebox.askyesnocancel("Aviso","Inicio de envio de correos "
                                + "\n Total de correos a enviar: " + str(len(id))
                                + "\n De: " + correo
                                + "\n Para: " + lis_co[0]
                                + "\n"+ new_saludo + nombre[0]
                                + "\n"+ new_cuerpo
                                + "\n" + extra[0]
                                + "\n" + pdf[0]
                                + "\n ¿Se encuentra el correo en orden?")                                
                    if pregunta:
                        while(l != num_cor and tot_cor != 500):
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

                    if "cancel":
                        return
                    else:
                        pass
                except gaierror:
                    messagebox.showerror("Error de conexión", 
                                "No hay conexión a Internet." + 
                                "\nPor favor, verifica tu conexión e inténtalo de nuevo.")            
                except Exception as e:
                    messagebox.showerror("Error", f"Ha ocurrido un error: {e}")
                    pass
            else:
                messagebox.showerror("ALTO", "Datos de asunto y/o saludo deben de llenarse")
                pass
    else:
        messagebox.showinfo("Aviso", "Favor de seleccionar un archivo Excel")
        pass

#Función para borrar el texto explicativo
def on_entry_asun(event):
    if Entry_asun.get() == asun:
            Entry_asun.delete(0, "end")

def on_entry_salu(event):
    if Entry_salu.get() == salu:
            Entry_salu.delete(0, "end")

def on_entry_body(event):
    if Entry_body_mail.get("1.0", "end-1c") == body:
            Entry_body_mail.delete("1.0", "end-1c")

def on_entry_mail(event, entry_mail):
    if entry_mail.get() == mail:
            entry_mail.delete(0, "end")
    pass

def on_entry_pass(event, entry_password):
    if entry_password.get() == pasword:
            entry_password.delete(0, "end")
    pass

#Inicio del programa

#Crea una carpeta y un archivo
Ruta = Path(pasw)
makedirs(Ruta, exist_ok=True)
uspwd_file = Path(arch)
if not uspwd_file.exists():
    with open(arch, "w") as file:
        file.write("")

"""
#Leer usuario y contraseña del archivo de texto
uspwd = open(arch, "r")
uspwdread = uspwd.readlines()
cantidad = len(uspwdread)
if(cantidad == 0):
    uspwdwrite = open(arch, "w+")
    correo = Tk()
    correo.geometry("450x250")
    correo.title("Carga de datos")

    entry_mail =  Entry(correo, width = 45)
    entry_mail.insert('insert', mail)
    entry_mail.bind("<Button-1>", on_entry_mail)

    entry_password =  Entry(correo, width = 45)
    entry_password.insert('insert', pasword)
    entry_password.bind("<Button-1>", on_entry_pass)

    user =  Label(correo, text = "")
    pwd  =  Label(correo, text = "")
    guardar =  Button(correo,
                      text ="Guardar",
                      width = 45, 
                      height = 1,
                      bg = "gray", 
                      command = escribir)
    button_link = Button(correo, 
                     text = "Link para saber tu contraseña",
                     width = 45,
                     height= 1,
                     bg= "green",
                     command = link)
    
    entry_mail.pack(padx=10, pady=5)
    entry_password.pack(padx=10, pady=5)
    guardar.pack(padx=10, pady=5)
    button_link.pack(padx=10, pady=5)

    correo.mainloop()
else:
    pass
"""

cambio_usuario()

#Inicio del programa de envío de correos
ven_correos = Tk()
ven_correos.geometry("500x650")
ven_correos.configure(background='gray')
ven_correos.title("Envío de Correos Masivos")
menu = Menu(ven_correos)
new_item = Menu(menu, border='50', tearoff=False)
ven_correos.config(menu=menu)
#ven_correos.iconphoto(True, PhotoImage(file="C:/Users/hp/Desktop/Carta.png"))
frame_button = ttk.Frame(ven_correos)
#División del programa en 2 partes, en la entrada de texto y los botones

#Items dentro del menú desplegable
menu.add_cascade(label='Menú', menu=new_item)
new_item.add_command(label='Cambio de servidor', command=cambio_server)
new_item.add_separator()
new_item.add_command(label='Registro de usuario', command=cambio_usuario)

#Lugar donde entrará la información
Entry_asun =  Entry(ven_correos, text= "", width=50, justify="center", border=5,
                font=font.Font(family="Arial",size=12,slant=font.ITALIC),
                foreground="Black")
Entry_asun.insert(0, asun)
Entry_asun.bind("<Button-1>", on_entry_asun)

Entry_salu =  Entry(ven_correos, text= "", width=50, justify="center",border=5,
                font=font.Font(family="Arial",size=12,slant=font.ITALIC),
                foreground="Black")
Entry_salu.insert(0, salu)
Entry_salu.bind("<Button-1>", on_entry_salu)

Entry_body_mail= scrolledtext.ScrolledText(ven_correos,width=50, border=5,
                font=font.Font(family="Arial",size=12,slant=font.ITALIC),
                foreground="Black")
Entry_body_mail.insert('insert', body)
Entry_body_mail.bind('<Button-1>', on_entry_body)

btn_envio =  Button(frame_button, text ="Envío", width = 25,
                height = 2, bg = "green", 
                command=enviar, font=font.Font(
                    font=("Arial", 13)))

btn_images = Button(frame_button, text ="Insertar imagen", width = 25,
                height = 2, bg = "green", font=font.Font(
                    font=("Arial", 13)))

barra = ttk.Progressbar(ven_correos, orient='horizontal', length='100')


#Lugar donde se meta
Entry_asun.pack(padx=10, pady=5)
Entry_salu.pack(padx=10, pady=5)
Entry_body_mail.pack(padx=10, pady=5, )
frame_button.pack(expand=True, padx=5, pady=5)
btn_envio.pack(side= 'left',padx=10, pady=10)
btn_images.pack(side= 'left',pady=10, padx=10)
ven_correos.mainloop()