from tkinter import Tk, Label, Button, Entry, filedialog, messagebox, scrolledtext, PhotoImage #Librería para la creación visual del programa
from openpyxl import load_workbook #Importamos libreria para el uso de excel
from smtplib import SMTP #Libreria dentro de python para el envio de correos
import webbrowser as web #Librería para poner una url en el programa
from socket import gaierror #Librería con el fin de poner 
from pathlib import Path
from os import makedirs
#from os import path


#https://myaccount.google.com/lesssecureapps?pli=1&rapt=AEjHL4PDAcXeJzQCiyuIDVjmt4tfaMySFl40ei6dVTFTvz67ZJJl5s9nZV18NkNjNRd9jJ06jljezI9sUpZK8na8IUGjK6omA2_tjCXlK0doz2pBF2O7S3s
#https://myaccount.google.com/apppasswords?utm_source=google-account&utm_medium=myaccountsecurity&utm_campaign=tsv-settings&rapt=AEjHL4Mk5lLF8MX793jG6SdOJKKflJoFb5rDbg3TutbXD3LAXWwxUTF5Q5gKJZZqw4geHCcZTzVym_Gr33b-nhnP67f1jVtWir6glV4dKvbg3EqS7IhLniM

#Links y rutas de archivos donde se crearán y guardarán datos del usuario
n_link = "https://myaccount.google.com/apppasswords?utm_source=google-account&utm_medium=myaccountsecurity&utm_campaign=tsv-settings&rapt=AEjHL4Mk5lLF8MX793jG6SdOJKKflJoFb5rDbg3TutbXD3LAXWwxUTF5Q5gKJZZqw4geHCcZTzVym_Gr33b-nhnP67f1jVtWir6glV4dKvbg3EqS7IhLniM"    
pasw = "C:/UsPwd"
arch = "C:/UsPwd/uspwd.txt"

#Inicio de funciones

def escribir():
    res = messagebox.askyesnocancel(title="Aviso", message="Se sobre escribiran los datos\n" +"Desea seguir")
    if res:
        ag1 = entry_1.get()
        ag2 = entry_2.get()
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
    correo.destroy()

def link():    
    web.open(n_link)

def preparar():
    l = 0    
    h = 0
    cont = 0
    tot_cor = 0
    a = "@"
    lis_er = []
    lis_er_str = ""
    
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        libro = load_workbook(filename=archivo)
        hoja = libro.active
        rango = (hoja['A1':'E500'])
        nombre = []
        lis_co = []
        pdf = []
        curso = []
        id = []
        for j in range(0,5):
            for k in range(0,500):
                while(j == 0):
                    if(rango[k][j].value != None and
                       rango[k][j].value != 'id'):
                        id.append(rango[k][j].value)
                    break
                while (j == 1):
                    if(rango[k][j].value != 'nombre' and
                        rango[k][j].value != None):
                        nombre.append(rango[k][j].value)
                    break
                while(j == 2):
                    if(rango[k][j].value != 'correo' and
                        rango[k][j].value != None):
                        lis_co.append(rango[k][j].value)
                    break
                while (j == 3):
                    if(rango[k][j].value != 'archivo' and
                        rango[k][j].value != None):
                        pdf.append(rango[k][j].value)
                    break
                while (j == 4):
                    if(rango[k][j].value != 'extra' and
                        rango[k][j].value != None):
                        curso.append(str(rango[k][j].value))
                    break
        messagebox.showinfo("Aviso", "Datos del Archivo Excel cargados correctamente")
    else:
        messagebox.showinfo("Aviso", "Favor de seleccionar un archivo Excel")
        
    for i in lis_co:    
        duda = i.rfind(a)
        if(duda == -1):
            cont = cont + 1   
            lis_er.append(i)
        else:
            pass
    
    for h in lis_er:
        lis_er_str = lis_er_str + h + ", "
    
    asf = lis_er_str
        
    if(len(lis_er)> 0):
        messagebox.showinfo("Aviso", "Los siguientes correos tiene problemas, favor de revisarlos: "
                        + "\n"+ asf)
        pass
    else:                
        agarrar1 = c_n_asunto.get() #Asunto
        agarrar2 = c_n_saludo.get() #Saludo
        agarrar3 = c_ccorreo.get("1.0", "end-1c") #Cuerpo del correo
        agarrar4 = c_despedida.get() #Despedida
        texto_f1["text"] = agarrar1
        texto_f2["text"] = agarrar2
        texto_f3["text"] = agarrar3
        texto_f4["text"] = agarrar4
        newagarrar1 = str(agarrar1)
        newagarrar2 = str(agarrar2+ ":  ")
        newagarrar3 = str(agarrar3)
        newagarrar4 = str(agarrar4)
        num_cor = len(id)
        if(newagarrar1 != None or newagarrar1 != '' or
                    newagarrar2 != None or newagarrar2 != ''):
                dato = open(arch, "r")
                info = dato.readline()
                dato.close()
                array = info.split(",")
                print(array)
                correo = array[0]
                contra = array[1]
                message = 'Subject: {}\n\n{}'.format(newagarrar1, newagarrar2)
                try:
                    server = SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(correo,contra)
                    pregunta = messagebox.askyesnocancel("Aviso", "Inicio de envio de correos "
                                                    + "\n Total de correos a enviar: " + str(len(id))
                                                    + "\n De: " + correo
                                                    + "\n Para: " + lis_co[0]
                                                    + "\n"+ newagarrar2 + nombre[0]
                                                    + "\n"+ newagarrar3
                                                    + "\nCurso tomado: " + curso[0]
                                                    + "\nMuchas gracias por estar con nosotros "
                                                    + "\n" + pdf[0]
                                                    + "\n" + newagarrar4)
                    if pregunta:
                        while(l != num_cor and tot_cor != 500):
                            server.sendmail(correo, lis_co[l], message + nombre[l]
                                                                +"\n"+ newagarrar3
                                                                + "\n\n Curso tomado: " + curso[l]
                                                                + "\n Muchas gracias por estar con nosotros "
                                                                + "\n\n " + pdf[l]
                                                                + "\n\n " + newagarrar4)
                            l = l + 1
                            tot_cor = tot_cor + 1
                        else:
                            messagebox.showinfo("Aviso", "Envio de correos finalizado")
                    elif "cancel":
                        return
                    else:
                        pass
                except gaierror:
                    messagebox.showerror("Error de conexión", "No hay conexión a Internet. Por favor, verifica tu conexión e inténtalo de nuevo.")            
                except Exception as e:
                    messagebox.showerror("Error", f"Ha ocurrido un error: {e}")
                    pass
        else:
            messagebox.showerror("ALTO", "Datos de asunto y/o saludo deben de llenarse")
            pass

#Inicio del programa raíz

#Crea una carpeta
Ruta = Path(pasw)
makedirs(Ruta, exist_ok=True)

#Crea un archivo

uspwd_file = Path(arch)
if not uspwd_file.exists():
    with open(arch, "w") as file:
        file.write("")

#Usuario y Contraseña       
uspwd = open(arch, "r")
uspwdread = uspwd.readlines()
cantidad = len(uspwdread)
if(cantidad == 0):
    uspwdwrite = open(arch, "w+")
    correo = Tk()
    correo.geometry("500x200")
    correo.title("Envío de correos masivos")
    #correo.iconbitmap("C:/Users/hp/Desktop/Carta.ico")
    text1 =  Label(correo, text = "Favor de introducir\n el correo")
    text2 =  Label(correo, text = "Favor de introducir\n la contraseña")
    entry_1 =  Entry(correo, text = "", width = 45)
    entry_2 =  Entry(correo, text = "", width = 45)
    user =  Label(correo, text = "")
    pwd  =  Label(correo, text = "")
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
    
    text1.grid(row = 0, column = 0)
    text2.grid(row = 2, column = 0)
    
    botton1.grid(row = 4, column = 0, columnspan = 3)
    botton2.grid(row = 5, column = 0, columnspan = 3)
    
    entry_1.grid(row = 1, column = 1)
    entry_2.grid(row = 3, column = 1)
    
    correo.mainloop()
    cantidad = len(uspwdread)
    print(cantidad)
else:
    uspwd.close()
uspwd.close()

#Inicio del programa de envío de correos
ventana = Tk()
ventana.geometry("400x300")
ventana.configure(background='gray')
ventana.title("Envío de Correos Masivos")
#ventana.iconphoto(True, PhotoImage(file="C:/Users/hp/Desktop/Carta.png"))

#Lugar donde entrará la información
c_n_asunto =  Entry(ventana, text = "")
c_n_saludo =  Entry(ventana, text = "")
c_ccorreo =  Entry(ventana, text = "")
c_ccorreo= scrolledtext.ScrolledText(ventana,width=30,height=10)
c_despedida =  Entry(ventana, text = "")

#Lugar donde se darán las instrucciones
text1 =  Label(ventana, text = "Zona del asunto ->")
text2 =  Label(ventana, text = "Zona del saludo ->")
text3 =  Label(ventana, text = "Zona del cuerpo ->\n del correo")
text4 =  Label(ventana, text = "Zona de la despedida ->")
text5 =  Label(ventana, text = "Programa hecho por Daniel Job Valadez Elguera", bg = "gray")

#Lugar donde se imprimen los resultados
texto_f1 =  Label(ventana, text = "")
texto_f2 =  Label(ventana, text = "")
texto_f3 =  Label(ventana, text = "")
texto_f4 =  Label(ventana, text = "")

#Botón de envío de correos
botton4 =  Button(ventana, text ="Envío de correos", width = 45, height = 1, bg = "green", command=preparar)
botton4.grid(row = 7, column = 0, columnspan = 4)

text1.grid(row = 2, column = 1)
text2.grid(row = 3, column = 1)
text3.grid(row = 4, column = 1)
text4.grid(row = 5, column = 1)
text5.grid(row = 8, column = 2)

c_n_asunto.grid(row = 2, column = 2)
c_n_saludo.grid(row = 3, column = 2)
c_ccorreo.grid(row = 4, column = 2)
c_despedida.grid(row = 5, column = 2)

ventana.mainloop()