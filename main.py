import sqlite3
import tkinter
from tkinter.filedialog import askopenfilename
from tkinter import ttk
import openpyxl
from exchangelib import DELEGATE, Account, Credentials, Configuration, Message, Mailbox
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os

db = sqlite3.connect("db.sqlite")
db.execute("CREATE TABLE IF NOT EXISTS"
           " clientes (name TEXT, _id PRIMARY KEY NOT NULL, email TEXT, balance_corto TEXT, balance_largo TEXT,"
           " marketing TEXT, cobranza TEXT)")
db.execute("CREATE TABLE IF NOT EXISTS exchange (usuario TEXT, clave TEXT, server TEXT, email TEXT)")
db.execute("CREATE TABLE IF NOT EXISTS smtp (email TEXT, clave TEXT, server TEXT)")


class ScrollBox(tkinter.Listbox):

    def __init__(self, window, **kwargs):
        super().__init__(window, **kwargs)
        self.scrollbar = ttk.Scrollbar(window, orient=tkinter.VERTICAL, command=self.yview)

    def grid(self, row, column, sticky="nsw", rowspan=1, columnspan=1, **kwargs):
        super().grid(row=row, column=column, sticky=sticky, rowspan=rowspan, columnspan=columnspan, **kwargs)
        self.scrollbar.grid(row=row, column=column, sticky="nse", rowspan=rowspan)
        self['yscrollcommand'] = self.scrollbar.set


class DataListBox(ScrollBox):

    def __init__(self, window, connection, table, sort, **kwargs):
        super(DataListBox, self).__init__(window, **kwargs)

        self.cursor = connection.cursor()
        self.table = table
        self.sort = sort

        self.sql_select = "SELECT * FROM " + self.table
        self.sql_sort = " ORDER BY " + self.sort

        sql = self.sql_select + self.sql_sort
        self.cursor.execute(sql)
        print(sql)
        for valueC in self.cursor:
            maped = map(str, valueC)
            self.insert(tkinter.END, " | ".join(maped))


# ----------- VENTANA EMAIL --------------
def guardaremailexchange():
    usuario_data = usuario.get()
    clave_data = clave.get()
    server_data = server.get()
    email_data = emailusuario.get()
    cursor = db.cursor()

    if len(usuario_data) != 0:
        if len(clave_data) != 0:
            if len(server_data) != 0:
                if len(email_data) != 0:
                    db.execute("DELETE FROM exchange")
                    db.execute("INSERT INTO exchange VALUES ('{}', '{}', '{}', '{}')".format(
                        usuario_data,
                        email_data,
                        clave_data,
                        server_data
                    ))
                    db.commit()
                    print("SUCCES")
                    usuario.delete(0, 'end')
                    clave.delete(0, 'end')
                    server.delete(0, 'end')
                    emailusuario.delete(0, 'end')
                    cursor.execute("SELECT * FROM exchange")
                    for usuariodata, emaildata, clavedata, serverdata in cursor:
                        longitudclaveexchange = len(clavedata)
                        codigoexchange = "*" * longitudclaveexchange
                        cuentaexchange.set("{} | {} | {} | {}".format(
                            usuariodata, emaildata, codigoexchange, serverdata))


def guardaremailsmtp():
    email_data = emailusuariosmtp.get()
    clave_data = clavesmtp.get()
    server_data = serversmtp.get()
    cursor = db.cursor()

    if len(server_data) != 0:
        if len(clave_data) != 0:
            if len(email_data) != 0:
                db.execute("DELETE FROM smtp")
                db.execute("INSERT INTO smtp VALUES ('{}', '{}', '{}')".format(
                    email_data,
                    clave_data,
                    server_data,
                ))
                db.commit()
                print("SUCCES")
                emailusuariosmtp.delete(0, 'end')
                clavesmtp.delete(0, 'end')
                serversmtp.delete(0, 'end')
                cursor.execute("SELECT * FROM smtp")
                for emaildatasmtp, clavedatasmtp, serverdatasmtp in cursor:
                    longitudclavesmtp = len(clavedatasmtp)
                    codigosmtp = "*" * longitudclavesmtp
                    cuentasmtp.set("{} | {} | {}".format(emaildatasmtp, codigosmtp, serverdatasmtp))


def emailwindow():
    global usuario, clave, server, emailusuario
    global emailusuariosmtp, clavesmtp, serversmtp
    global ventana, addexchange, addsmtp
    global cuentasmtp, cuentaexchange
    cursor = db.cursor()

    ventana = tkinter.Toplevel()
    ventana.geometry("600x300")
    ventana.title("Configurar Email")
    ventana.config(bg="light grey")

    principal = ttk.Notebook(ventana, width=600, height=300)
    principal.grid(sticky="nswe")

    # ----------------- EXCHANGE --------------------
    exchange = ttk.Frame(principal, style='frame.TFrame')
    exchange.columnconfigure(0, weight=4)
    exchange.columnconfigure(1, weight=4)
    exchange.columnconfigure(2, weight=4)
    exchange.rowconfigure(0, weight=4)
    exchange.rowconfigure(1, weight=4)
    exchange.rowconfigure(2, weight=4)

    addexchange = ttk.LabelFrame(exchange, style='fieldframe.TLabelframe', text="Add Exchange")
    addexchange.grid(row=0, column=1, sticky="nwe", pady=10)

    usuario = ttk.Entry(addexchange, width=30)
    usuario.grid(row=1, column=1, pady=5, padx=5)
    emailusuario = ttk.Entry(addexchange, width=30)
    emailusuario.grid(row=2, column=1, pady=5, padx=5)
    clave = ttk.Entry(addexchange, width=30, show="*")
    clave.grid(row=3, column=1, pady=5, padx=5)
    server = ttk.Entry(addexchange, width=30)
    server.grid(row=4, column=1, pady=5, padx=5)

    usuariolabel = ttk.Label(addexchange, style='label.TLabel', text="Usuario:")
    usuariolabel.grid(row=1, column=0, sticky="e")

    emailusuariolabel = ttk.Label(addexchange, style='label.TLabel', text="Email:")
    emailusuariolabel.grid(row=2, column=0, sticky="e")

    clavelabel = ttk.Label(addexchange, style='label.TLabel', text="Clave:")
    clavelabel.grid(row=3, column=0, sticky="e")

    serverlabel = ttk.Label(addexchange, style='label.TLabel', text="Servidor:")
    serverlabel.grid(row=4, column=0, sticky="e")

    savebutton = ttk.Button(addexchange, style='a.TButton', text='Guardar', command=guardaremailexchange)
    savebutton.grid(row=5, column=1, sticky="n")

    cursor.execute("SELECT * FROM exchange")
    cuentaexchange = tkinter.StringVar()
    for usuariodata, emaildata, clavedata, serverdata in cursor:
        longitudclaveexchange = len(clavedata)
        codigoexchange = "*" * longitudclaveexchange
        cuentaexchange.set("{} | {} | {} | {}".format(usuariodata, emaildata, codigoexchange, serverdata))
    cuenta = ttk.Label(exchange, style="greylabel.TLabel", textvariable=cuentaexchange)
    cuenta.grid(row=1, column=1)
    # ------------------------------------------------------

    smtpemail = ttk.Frame(principal, style='frame.TFrame')
    smtpemail.columnconfigure(0, weight=4)
    smtpemail.columnconfigure(1, weight=4)
    smtpemail.columnconfigure(2, weight=4)
    smtpemail.rowconfigure(0, weight=4)
    smtpemail.rowconfigure(1, weight=4)
    smtpemail.rowconfigure(2, weight=4)

    addsmtp = ttk.LabelFrame(smtpemail, style='fieldframe.TLabelframe', text="Add smtp", width=400)
    addsmtp.grid(row=0, column=1, sticky="nwe", pady=10)

    emailusuariosmtp = ttk.Entry(addsmtp, width=30)
    emailusuariosmtp.grid(row=1, column=1, pady=5, padx=5)
    clavesmtp = ttk.Entry(addsmtp, width=30, show="*")
    clavesmtp.grid(row=2, column=1, pady=5, padx=5)
    serversmtp = ttk.Entry(addsmtp, width=30)
    serversmtp.grid(row=3, column=1, pady=5, padx=5)

    emailusuariolabelsmtp = ttk.Label(addsmtp, style='label.TLabel', text="Email:")
    emailusuariolabelsmtp.grid(row=1, column=0, sticky="e")

    clavelabelsmtp = ttk.Label(addsmtp, style='label.TLabel', text="Clave:")
    clavelabelsmtp.grid(row=2, column=0, sticky="e")

    usuariolabelsmtp = ttk.Label(addsmtp, style='label.TLabel', text="Servidor:")
    usuariolabelsmtp.grid(row=3, column=0, sticky="e")

    savebuttonsmtp = ttk.Button(addsmtp, style='a.TButton', text='Guardar', command=guardaremailsmtp)
    savebuttonsmtp.grid(row=4, column=1, sticky="n")

    cursor.execute("SELECT * FROM smtp")
    cuentasmtp = tkinter.StringVar()
    for emaildatasmtp, clavedatasmtp, serverdatasmtp in cursor:
        longitudclavesmtp = len(clavedatasmtp)
        codigosmtp = "*" * longitudclavesmtp
        cuentasmtp.set("{} | {} | {}".format(emaildatasmtp, codigosmtp, serverdatasmtp))
    cuenta_smtp = ttk.Label(smtpemail, style="greylabel.TLabel", textvariable=cuentasmtp)
    cuenta_smtp.grid(row=1, column=1)

    principal.add(exchange, text="Exchange")
    principal.add(smtpemail, text="SMTP")
# -----------------------------------------


def html(nombre, balance_c, balance_l, es_cobranza):
    cobranza_si = ""
    if es_cobranza.lower() == "si":
        cobranza_si = "<p><span>Costas de Abogado serán calculadas al coordinar reunión</span></p>"
    css = """
                <style type="text/css">
                    * {box-sizing: border-box;}
                    .contenedor {
                        width: 700px; 
                        height: auto; 
                        padding: 5px; 
                        padding-bottom: 25px;
                        border: 2px solid #922E8D;
                        border-radius: 10px;
                    }
                    p span {display: block; font-size: 18px;}
                    body {
                        font-family: Arial;
                        font-size: 20px;
                    }
                    h3 {font-size: 24px;}
                    table {
                        border: 1px solid rgba(0,0,0, .4);
                        border-collapse: collapse;
                        width: 500px;
                        text-align: center;
                    }
                    th {
                        border: 1px solid rgba(0,0,0, .4); 
                        padding: 10px 20px; 
                        background: #D46C18; 
                        color: white;
                    }
                    td {
                        border: 1px solid rgba(0,0,0, .4); 
                        padding: 10px 20px;
                    }
                    tr:hover {background: #941D8E; color: white;}    
                    .final {
                        width: 400px;
                        padding: 6px;
                        border: 3px solid #922E8D;
                        border-radius: 15px;
                    }
                </style>
                """
    body_html = """
                <!DOCTYPE html>
                <html lang="es">
                <head>
                    <meta charset="utf-8">
                    <title>Mail</title>
                    {}
                </head>
                <body>
                    <div class="contenedor">
                        <h3>Estimado {}:</h3>
                        <p>
                            <span>Mi nombre es Claudia Castro, soy ejecutiva de normalización del Banco Security</span>
                            <span>Estoy encargada de regularizar su deuda, que el día de hoy es:</span>
                        </p>
    
                        <table>
                            <tr>
                                <th>Deuda Corto Plazo</th>
                                <th>Deuda Largo Plazo</th>
                            </tr>
                            <tr>
                                <td>{}</td>
                                <td>{}</td>
                            </tr>
                        </table>
                        {}
                        <p><span>Las Alternativas de pago son</span></p>
                        <p>
                            <span>1) Al contado, cheque por el total de la deuda vencida</span>
                            <span>2) 30% de abono y saldo en un crédito plazo por Definir</span>
                        </p>
                        <p><span>Para ello es necesario:</span></p>
                        <p>
                            <span>- Completar estado de situación adjunto</span>
                            <span>- Acreditar ingresos</span>
                            <span>* Sujeto a evaluación</span>
                        </p>
                        <p><span>Quedo atento a sus comentarios</span></p>
                        
                        <div class="final">
                            <p>
                                <span>Claudia Castro</span>
                                <span>Ejecutiva de Normalización | Banco Security</span>
                                <span>Mail: Claudia.Castro@security.cl</span>
                                <span>Fono: 225844072</span>
                            </p>
                        </div>   
                    </div>
                </body>
                </html>
                """.format(css, nombre, balance_c, balance_l, cobranza_si)
    return body_html


# ------------- ARCHIVO ----------------------
def openfile():
    archivo = askopenfilename(initialdir="C://", title="Elije Un archivo", filetypes=(
        ("All files", "*.*"),
        ("Excel", "*.xlxs*")
    )
                              )
    if len(archivo) != 0:
        _, extension = os.path.splitext(archivo)
        print(extension)
        if extension == ".xlxs":
            excel = openpyxl.load_workbook(archivo)
            sheet = excel.get_sheet_by_name('Hoja1')
            result = []
            loop_row = 1
            saltar = 1
            for row in sheet.iter_rows():
                for cell in row:

                    if loop_row == 8:
                        break
                    if saltar == 1:
                        print("Saltando Primera Linea")
                    else:
                        result.append(cell.value)
                    loop_row += 1
                    saltar += 1
                try:
                    nombre, key, mail, balance_c, balance_l, marketin, cobranz = result
                    nombre = str(nombre)
                    key = str(key)
                    mail = str(mail)
                    balance_c = str(balance_c)
                    balance_l = str(balance_l)
                    marketin = str(marketin)
                    cobranz = str(cobranz)

                except ValueError:
                    print("Sobran o faltan valores")
                else:
                    if key == "None":
                        print("vacio")
                    else:
                        try:
                            keyint = int(key)
                        except ValueError:
                            print("No es un entero")
                        else:
                            if nombre == "None":
                                nombre = ""
                            if mail == "None":
                                mail = ""
                            if balance_c == "None":
                                balance_c = ""
                            if balance_l == "None":
                                balance_l = ""
                            if marketin == "None":
                                marketin = ""
                            if cobranz == "None":
                                cobranz = ""

                            try:
                                db.execute("INSERT INTO clientes VALUES('{}', {}, '{}', '{}', '{}', '{}', '{}')".format(
                                    nombre,
                                    keyint,
                                    mail,
                                    balance_c,
                                    balance_l,
                                    marketin,
                                    cobranz
                                ))
                            except sqlite3.IntegrityError:
                                print("Ese id ya existe")
                                errorvalue(fileLabelText, "Algun id ya existe", 3000)
                                break
                            else:
                                db.commit()
                                errorvalue(fileLabelText, "Excel Transferido correctamente", 5000)
                result = []
                loop_row = 1
        else:
            print("No es un Archivo Excel")
            errorvalue(fileLabelText, "No es un archivo excel", 4000)
    refresh_data()
# --------------------------------------------


# EMAIL
def send_email():
    global sv
    global account
    mailstotales = 0
    mailsenviados = 0
    tipo_email = value.get()
    cursor = db.cursor()
    estado_conectado = False

    # ----------------- TIPO 1 ---------------------
    if tipo_email == 1:
        cursor.execute("SELECT * FROM exchange")
        datos = cursor.fetchone()
        if datos is not None:
            user, contra, servidor, mailusuario = datos
            credentials = Credentials(
                username=user,
                password=contra
            )
            config = Configuration(server=servidor, credentials=credentials)
            try:
                account = Account(
                    primary_smtp_address=mailusuario,
                    config=config,
                    autodiscover=False,
                    access_type=DELEGATE
                )
            except:
                print("Error de Conexion")
                errorvalue(emailSuccesString, "Error de conexion", 4000)
            else:
                estado_conectado = True

            if estado_conectado:
                for nombre, key, mail, balance_c, balance_l, marketin, cobranz in cursor:
                    if marketin != "+":
                        if len(mail) != 0:

                            body_html = html(nombre, balance_c, balance_l, cobranz)
                            mensaje = Message(
                                account=account,
                                subject=key,
                                body=body_html,
                                to_recipients=[Mailbox(email_address=mail)]
                            )
                            try:
                                mensaje.send()
                            except:
                                print("Error al enviar")
                                mailstotales += 1
                            else:
                                mailsenviados += 1
                                mailstotales += 1
                        else:
                            print("No Hay Email")
        else:
            print("No Hay Cuenta")
            errorvalue(emailSuccesString, "No hay cuenta Exchange Configurada", 3000)
    # ------------------------------------------------

    # ---------------- TIPO 2 ------------------------
    if tipo_email == 2:
        cursor.execute("SELECT * FROM smtp")
        datos = cursor.fetchone()
        if datos is not None:
            mailusuario, contra, servidor = datos
            try:
                sv = smtplib.SMTP(servidor, 587)
                sv.starttls()
                sv.login(mailusuario, contra)
            except:
                print("Error de conexion")
                errorvalue(emailSuccesString, "Error de conexion", 4000)
            else:
                print("Conexion Exitosa")
                estado_conectado = True

            if estado_conectado:
                sql = "SELECT * FROM clientes ORDER BY name"
                cursor.execute(sql)
                for nombre, key, mail, balance_c, balance_l, marketin, cobranz in cursor:
                    if marketin != "+":
                        if len(mail) != 0:
                            msg = MIMEMultipart()
                            msg['From'] = mailusuario
                            msg['To'] = mail
                            msg['Subject'] = "{}".format(key)

                            body_html = html(nombre, balance_c, balance_l, cobranz)

                            try:
                                archivo = MIMEApplication(open('archivo.pdf', 'rb').read())
                            except FileNotFoundError:
                                pass
                            else:
                                archivo.add_header('Content-Disposition', 'attachment', filename='pdfPrueba.pdf')
                                msg.attach(archivo)

                            body = MIMEText(body_html, 'html')
                            msg.attach(body)

                            try:
                                sv.sendmail(mailusuario, mail, msg.as_string())
                            except:
                                print("Error")
                                mailstotales += 1
                            else:
                                print("Mail enviado")
                                mailsenviados += 1
                                mailstotales += 1
                        else:
                            print("no hay email")
        else:
            errorvalue(emailSuccesString, "No hay cuenta smtp Configurada", 3000)
            if estado_conectado:
                sv.quit()
            print("Conexion Cerrada")
    errorvalue(emailSuccesString, "Enviado {} Mails de {}".format(mailsenviados, mailstotales), 3000)
    # ----------------------------------------------------------


# BASE DE DATOS
def save_data():

    namevalue = name.get()
    emailvalue = email.get()
    balance_corto_value = balance_corto.get()
    balance_largo_value = balance_largo.get()
    llavevalue = llave.get()
    marketingvalue = marketing.get()
    cobranzavalue = cobranza.get()
    es_entero = False
    if len(llavevalue) != 0:
        try:
            inter = int(llavevalue)
            print(inter)
        except ValueError:
            print("no es un numero")
            errorvalue(errorString, "Asegurate que sea un numero", 3000)
        else:
            pass
            es_entero = True

        if es_entero:
            db.execute("INSERT INTO clientes VALUES('{}', {}, '{}', '{}', '{}', '{}', '{}')".format(
                namevalue,
                llavevalue,
                emailvalue,
                balance_corto_value,
                balance_largo_value,
                marketingvalue,
                cobranzavalue
            ))
            db.commit()
            print("Exito al guardar")
            name.delete(0, 'end')
            llave.delete(0, 'end')
            email.delete(0, 'end')
            balance_corto.delete(0, 'end')
            balance_largo.delete(0, 'end')
            marketing.delete(0, 'end')
            cobranza.delete(0, 'end')
            refresh_data()
    else:
        print("No hay id")
        errorvalue(errorString, "No Hay id", 3000)


def refresh_data():
    listaClientes.delete(0, 'end')
    cursor = db.cursor()
    sql = "SELECT * FROM clientes ORDER BY name"
    cursor.execute(sql)
    for value in cursor:
        maped = map(str, value)
        listaClientes.insert(tkinter.END, " | ".join(maped))
    print("Refreshed")


def update_data():
    namevalue = name.get()
    emailvalue = email.get()
    balance_corto_value = balance_corto.get()
    balance_largo_value = balance_largo.get()
    llavevalue = llave.get()
    marketingvalue = marketing.get()
    cobranzavalue = cobranza.get()
    es_entero = False

    if len(llavevalue) != 0:
        try:
            inter = int(llavevalue)
            print(inter)
        except ValueError:
            print("no es un numero")
            errorvalue(errorString, "No es un numero", 3000)
        else:
            es_entero = True

        if es_entero:
            if len(emailvalue) != 0:
                db.execute("UPDATE clientes SET email = '{}' WHERE _id={}".format(emailvalue, llavevalue))

            if len(namevalue) != 0:
                db.execute("UPDATE clientes SET name = '{}' WHERE _id={}".format(namevalue, llavevalue))

            if len(balance_corto_value) != 0:
                db.execute("UPDATE clientes SET balance_corto = '{}' WHERE _id={}".format(
                    balance_corto_value,
                    llavevalue))

            if len(balance_largo_value) != 0:
                db.execute("UPDATE clientes SET balance_largo = '{}' WHERE _id={}".format(
                    balance_largo_value,
                    llavevalue
                ))

            if len(marketingvalue) != 0:
                db.execute("UPDATE clientes SET marketing = '{}' WHERE _id={}".format(marketingvalue, llavevalue))

            if len(cobranzavalue) != 0:
                db.execute("UPDATE clientes SET cobranza = '{}' WHERE _id={}".format(cobranzavalue, llavevalue))

            db.commit()
            print("Exito al guardar")
            name.delete(0, 'end')
            llave.delete(0, 'end')
            email.delete(0, 'end')
            balance_corto.delete(0, 'end')
            balance_largo.delete(0, 'end')
            marketing.delete(0, 'end')
            cobranza.delete(0, 'end')
            refresh_data()
    else:
        print("No hay id")
        errorvalue(errorString, "No hay id", 3000)


def encontrar():
    llavevalue = llave.get()
    cursor = db.cursor()
    es_entero = False
    if len(llavevalue) != 0:
        try:
            key = int(llavevalue)
            print(key)
        except ValueError:
            print("No es un numero")
            errorvalue(errorString, "No es un numero", 3000)
        else:
            es_entero = True

        if es_entero:
            listaClientes.delete(0, 'end')
            sql = "SELECT * FROM clientes WHERE _id={}".format(llavevalue)
            cursor.execute(sql)
            for value in cursor:
                maped = map(str, value)
                listaClientes.insert(tkinter.END, " | ".join(maped))
            llave.delete(0, 'end')
            print("Encontrado")
    else:
        print("No hay id")
        errorvalue(errorString, "No hay id", 3000)


def borrarbase():
    db.execute("DELETE FROM clientes")
    errorvalue(fileLabelText, "Base Borrada completamente", 3000)
    refresh_data()
    db.commit()


# ERRORES DE BASE DE DATOS
def errorvalue(label, mensaje="", segundos=0):
    label.set(mensaje)
    mainWindow.update()
    mainWindow.after(segundos, refresherror(label))


def refresherror(label):
    label.set("")


if __name__ == '__main__':

    # ----------- CONFIGURACION INICIAL -----------
    mainWindow = tkinter.Tk()
    mainWindow.title("Software Clientes")
    mainWindow.geometry("800x600")
    try:
        mainWindow.iconbitmap("img/app.ico")
    except:
        print("No se pudo cargar la imagen")
    menu = tkinter.Menu(mainWindow)
    mainWindow.config(menu=menu, bg="light grey")
    menuPrincipal = tkinter.Menu(menu, tearoff=False)
    menu.add_cascade(label="Options", menu=menuPrincipal)

    menuPrincipal.add_command(label="Email", command=emailwindow)

    mainWindow.columnconfigure(0, weight=20)
    mainWindow.columnconfigure(1, weight=2)
    mainWindow.columnconfigure(2, weight=4)
    mainWindow.columnconfigure(3, weight=2)

    mainWindow.rowconfigure(0, weight=2)
    mainWindow.rowconfigure(1, weight=5)
    mainWindow.rowconfigure(2, weight=4)
    mainWindow.rowconfigure(3, weight=2)
    # -------------------------------------------

    # ---------- ESTILOS ----------
    listaClientesLabel = ttk.Style()
    listaClientesLabel.configure('lista.TLabel', foreground="black", font=('Arial', 14), background="light grey")

    fieldFrameStyle = ttk.Style()
    fieldFrameStyle.configure('fieldframe.TLabelframe')

    fieldFrameTextStyle = ttk.Style()
    fieldFrameTextStyle.configure('texto.TLabel', font=('Arial', 12))

    botonStyle = ttk.Style()
    botonStyle.configure('a.TButton', foreground='black', background="black", font=('Arial', 10))

    errorStyle = ttk.Style()
    errorStyle.configure('errorstyle.TLabel', foreground="red", font=('Arial', 12), background="light grey")

    labelStyle = ttk.Style()
    labelStyle.configure('label.TLabel', foreground='black', font=('Arial', 10))

    frameStyle = ttk.Style()
    frameStyle.configure('frame.TFrame', background="light grey")

    greylabelStyle = ttk.Style()
    greylabelStyle.configure('greylabel.TLabel', background="light grey", foreground="Black", font=('Arial', 11))
    # -------------------------------------------

    # ----------- FRAMES ----------

    fieldFrameText = ttk.Label(text="Clientes", style='texto.TLabel')
    fieldFrame = ttk.LabelFrame(mainWindow, labelwidget=fieldFrameText, style='fieldframe.TLabelframe')
    fieldFrame.grid(row=2, column=2, sticky="sew", padx=10, pady=10, ipadx=10, ipady=10)

    fileFrameText = ttk.Label(text="Archivo", style='texto.TLabel')
    fileFrame = ttk.LabelFrame(mainWindow, labelwidget=fileFrameText, style='fieldframe.TLabelframe')
    fileFrame.grid(row=1, column=2, sticky="new", padx=(10, 10), pady=(10, 10))
    fileFrame.columnconfigure(0, weight=4)
    fileFrame.columnconfigure(1, weight=6)
    fileFrame.columnconfigure(2, weight=4)

    botonFrameText = ttk.Label(style='texto.TLabel')
    botonFrame = ttk.LabelFrame(mainWindow, labelwidget=botonFrameText, style='fieldframe.TLabelframe')
    botonFrame.grid(row=3, column=0, sticky="sew", padx=(30, 10), pady=(0, 15))
    botonFrame.columnconfigure(0, weight=4)
    botonFrame.columnconfigure(1, weight=4)
    botonFrame.columnconfigure(2, weight=4)
    # --------------------------------------------

    # ---------- LABELS ----------
    ttk.Label(mainWindow, text="Lista Clientes", style='lista.TLabel').grid(row=0, column=0)

    nombreLabel = tkinter.Label(fieldFrame, text="Nombre:").grid(row=0, column=1)
    idLabel = tkinter.Label(fieldFrame, text="Id:").grid(row=1, column=1)
    emailLabel = tkinter.Label(fieldFrame, text="Email:").grid(row=2, column=1)
    balanceCortoLabel = tkinter.Label(fieldFrame, text="Deuda Corto Plazo:").grid(row=3, column=1)
    balanceLargoLabel = tkinter.Label(fieldFrame, text="Deuda Largo Plazo:").grid(row=4, column=1)

    marketingLabel = tkinter.Label(fieldFrame, text="Marketing:").grid(row=5, column=1)
    instruccionesMarketingLabel = tkinter.Label(fieldFrame, text=" '+' Si es Correcto")
    instruccionesMarketingLabel.grid(row=5, column=2, sticky="e")

    cobranzaLabel = tkinter.Label(fieldFrame, text="Cobranza:").grid(row=6, column=1)
    cobranzaLabelIn = tkinter.Label(fieldFrame, text=" 'Si' o 'No' ").grid(row=6, column=2, sticky="e")
    # --------------------------------------------

    # ---------- LISTA -----------
    listaClientes = DataListBox(mainWindow, db, "clientes", "name")
    listaClientes.grid(row=1, column=0, sticky="nsew", rowspan=2, padx=(30, 0), pady=(0, 10))
    listaClientes.config(border=2, relief="sunken")
    # --------------------------------------------

    # ---------- ENTRYS ----------
    name = ttk.Entry(fieldFrame, width=50)
    name.grid(row=0, column=2, pady=3)

    llave = ttk.Entry(fieldFrame, width=50)
    llave.grid(row=1, column=2, pady=3)

    email = ttk.Entry(fieldFrame, width=50)
    email.grid(row=2, column=2, pady=3)

    balance_corto = ttk.Entry(fieldFrame, width=50)
    balance_corto.grid(row=3, column=2, pady=3)

    balance_largo = ttk.Entry(fieldFrame, width=50)
    balance_largo.grid(row=4, column=2, pady=3)

    marketing = ttk.Entry(fieldFrame, width=30)
    marketing.grid(row=5, column=2, sticky="w", pady=3)

    cobranza = ttk.Entry(fieldFrame, width=30)
    cobranza.grid(row=6, column=2, sticky="w", pady=5)

    # -------------------------------------------------

    # ---------- BOTONES ----------
    sumbit = ttk.Button(fieldFrame, text="Save", command=save_data, width=6, style='a.TButton')
    sumbit.grid(row=7, column=2, sticky="w")

    update = ttk.Button(fieldFrame, text="Actualizar", command=update_data, style='a.TButton')
    update.grid(row=7, column=2)

    find = ttk.Button(fieldFrame, text="Buscar", width=6, command=encontrar, style='a.TButton')
    find.grid(row=7, column=2, sticky="e")

    refresh = ttk.Button(botonFrame, text="Refresh", command=refresh_data, style='a.TButton')
    refresh.grid(row=0, column=0, pady=(8, 8))

    enviar = ttk.Button(botonFrame, text="Enviar", command=send_email, style='a.TButton')
    enviar.grid(row=0, column=2, pady=(8, 8))

    filebotton = ttk.Button(fileFrame, text="Abrir Archivo", command=openfile, style='a.TButton')
    filebotton.grid(row=1, column=0, padx=(8, 8), pady=(8, 8))

    borrar = ttk.Button(fileFrame, text="Borrar Base de datos", style="a.TButton", command=borrarbase)
    borrar.grid(row=1, column=2, padx=(8, 8), pady=(8, 8))
    # --------------------------------------------------

    # ---------- ERROR LABELS ----------

    errorString = tkinter.StringVar()
    errorString.set("")
    errorLabel = ttk.Label(mainWindow, textvariable=errorString, style='errorstyle.TLabel')
    errorLabel.grid(row=3, column=2, sticky="n")

    fileLabelText = tkinter.StringVar()
    fileLabelText.set("")
    fileLabel = ttk.Label(mainWindow, textvariable=fileLabelText, style='errorstyle.TLabel')
    fileLabel.grid(row=0, column=2, sticky="s")

    emailSuccesString = tkinter.StringVar()
    emailSuccesString.set("")
    emailSuccesLabel = ttk.Label(mainWindow, textvariable=emailSuccesString, style='errorstyle.TLabel')
    emailSuccesLabel.grid(row=3, column=2, sticky="w")
    # ---------------------------------------------------

    # ----------------- RADIO BUTTON ------------------
    value = tkinter.IntVar()
    value.set(1)
    exchangebutton = ttk.Radiobutton(botonFrame, text="Exchange", variable=value, value=1)
    exchangebutton.grid(row=2, column=0)
    smtpbutton = ttk.Radiobutton(botonFrame, text="SMTP", variable=value, value=2)
    smtpbutton.grid(row=2, column=2)
    # -------------------------------------------------

    mainWindow.mainloop()
    db.close()
