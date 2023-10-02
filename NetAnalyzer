import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import json
import os
from netmiko import ConnectHandler, NetmikoTimeoutException, NetmikoAuthenticationException
import re
import sqlite3
from openpyxl import Workbook
import xlsxwriter

# Obtener la ruta absoluta del directorio donde se encuentra el script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Construir la ruta completa al archivo switches.db y users.json
DATABASE_NAME = "switches.db"
DATABASE_NAME = os.path.join(BASE_DIR, "switches.db")
USERS_FILE = os.path.join(BASE_DIR, "users.json")

# Variables globales
connection = None
reporte = {}  # <-- Aquí declaramos reporte
tabla_reporte = None  # <-- Añadido
entry_busqueda = None

# Cargar usuarios del archivo
if os.path.exists(USERS_FILE):
    with open(USERS_FILE, 'r') as file:
        users = json.load(file)
else:
    users = {}



# Funciones de de gestiones de comandos def

def ver_informe():
    informe_win = tk.Toplevel(app)
    informe_win.title("Informe de Switches")
    informe_win.geometry("1000x600")

    # Función de búsqueda
    def buscar():
        for item in tree.get_children():
            tree.delete(item)

        ip_buscar = entry_buscar.get()

        conn = sqlite3.connect(DATABASE_NAME)
        cursor = conn.cursor()
        
        query = """
        SELECT 
            SwitchMaster.ip, 
            SwitchMaster.nemonico, 
            SwitchMaster.version, 
            SwitchDetails.model_number, 
            SwitchDetails.serial_number, 
            SwitchDetails.mac_address 
            
        FROM 
            SwitchMaster 
        JOIN 
            SwitchDetails 
        ON 
            SwitchMaster.id = SwitchDetails.master_id 
        WHERE 
            SwitchMaster.ip LIKE ?
        ORDER BY 
            SwitchMaster.ip, 
            SwitchDetails.id
        """
        cursor.execute(query, (f"%{ip_buscar}%",))
        
        switch_counter = 0
        last_ip = None
        for row in cursor.fetchall():
            ip = row[0]
            if ip != last_ip:  # Si la IP cambió, reiniciar el contador
                switch_counter = 1
                last_ip = ip
            else:
                switch_counter += 1
            
            tree.insert("", "end", values=(ip, *row[1:], switch_counter))
        
        conn.close()
    
    def exportar_a_excel():
        filename = "Informe_Switches.xlsx"
        
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()

        headers = ["IP", "Hostname", "Version", "Modelo", "Serial", "Mac", "Switcher"]
        for col_num, header in enumerate(headers):
            worksheet.write(0, col_num, header)

        for row_num, row in enumerate(tree.get_children(), 1):
            values = tree.item(row, "values")
            for col_num, cell_value in enumerate(values):
                worksheet.write(row_num, col_num, cell_value)

        workbook.close()

        messagebox.showinfo("Información", f"Informe exportado exitosamente como {filename}")

    frame_buscar = ttk.Frame(informe_win)
    frame_buscar.pack(pady=20)

    label_buscar = ttk.Label(frame_buscar, text="Buscar por IP:")
    label_buscar.grid(row=0, column=0)
    
    entry_buscar = ttk.Entry(frame_buscar)
    entry_buscar.grid(row=0, column=1, padx=10)
    btn_buscar = ttk.Button(frame_buscar, text="Buscar", command=buscar)
    btn_buscar.grid(row=0, column=2)
    btn_exportar = ttk.Button(frame_buscar, text="Exportar a Excel", command=exportar_a_excel)
    btn_exportar.grid(row=0, column=3, padx=10)

    columns = ("IP", "Hostname", "Version", "Modelo", "Serial", "Mac", "Switcher")
    tree = ttk.Treeview(informe_win, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    tree.pack(fill="both", expand=True, pady=20)

def buscar_principal():
    ip_buscar = entry_busqueda.get()
    
    # Se abre una nueva conexión a la base de datos
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()

    cursor.execute("SELECT ip, nemonico, version, num_switches FROM SwitchMaster WHERE ip LIKE ?", (f"%{ip_buscar}%",))
    registros = cursor.fetchall()

    tabla_reporte.delete(*tabla_reporte.get_children())  # Limpiar tabla antes de agregar registros
    for registro in registros:
        tabla_reporte.insert("", tk.END, values=registro)
    
    # Cerramos la conexión
    conn.close()

def buscar_detalles(ip):
    """Retorna los detalles de los switches asociados a una IP en particular."""
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    cursor.execute("SELECT id FROM SwitchMaster WHERE ip = ?", (ip,))
    master_id = cursor.fetchone()
    if not master_id:
        return []
    
    cursor.execute("SELECT mac_address, model_number, serial_number FROM SwitchDetails WHERE master_id = ?", (master_id[0],))
    detalles = cursor.fetchall()
    
    conn.close()
    return detalles

def mostrar_detalles(event):
    global detalle_win

    # Obtener la IP seleccionada
    seleccionado = tabla_reporte.selection()[0]
    ip_seleccionada = tabla_reporte.item(seleccionado)["values"][0]

    # Crear o levantar la ventana de detalles
    if detalle_win is None or not tk.Toplevel.winfo_exists(detalle_win):
        detalle_win = tk.Toplevel(app)
        detalle_win.title(f"Detalles de {ip_seleccionada}")
        detalle_win.geometry("800x600")

        tabla_detalle = ttk.Treeview(detalle_win, columns=("Switcher", "Modelo", "Serial", "Mac"), show="headings")
        tabla_detalle.heading("Switcher", text="Switcher")
        tabla_detalle.heading("Modelo", text="Modelo")
        tabla_detalle.heading("Serial", text="Serial")
        tabla_detalle.heading("Mac", text="Mac")
        tabla_detalle.column("Switcher", width=100)
        tabla_detalle.column("Modelo", width=200)
        tabla_detalle.column("Serial", width=200)
        tabla_detalle.column("Mac", width=200)
        tabla_detalle.pack(fill="both", expand=True, pady=20)

        conn = sqlite3.connect(DATABASE_NAME)
        cursor = conn.cursor()

        cursor.execute("SELECT model_number, serial_number, mac_address FROM SwitchDetails WHERE master_id = ?", (ip_seleccionada,))
        
        # Añadir enumeración basada en la IP
        for index, (model, serial, mac) in enumerate(cursor.fetchall(), start=1):
            tabla_detalle.insert("", "end", values=(index, model, serial, mac))

        conn.close()

def init_db():
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS SwitchMaster (
        id INTEGER PRIMARY KEY,
        ip TEXT NOT NULL,
        nemonico TEXT,
        version TEXT,
        num_switches INTEGER
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS SwitchDetails (
        id INTEGER PRIMARY KEY,
        master_id INTEGER,
        mac_address TEXT NOT NULL,
        model_number TEXT,
        serial_number TEXT,
        FOREIGN KEY (master_id) REFERENCES SwitchMaster(id)
    )
    """)

    conn.commit()
    conn.close()

init_db()

def guardar_en_db(data):
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    # Verificar si ya existe un registro con la misma IP
    cursor.execute("SELECT id FROM SwitchMaster WHERE ip = ?", (data["ip"],))
    existente = cursor.fetchone()
    if existente:
        # Si existe, mostramos un mensaje y no hacemos nada más
        messagebox.showwarning("Advertencia", "Ya existe un reporte para esta IP. No se ha guardado ningún dato nuevo.")
        conn.close()
        return
    
    # Si no existe, insertamos el nuevo registro
    cursor.execute("""
    INSERT INTO SwitchMaster (ip, nemonico, version, num_switches)
    VALUES (?, ?, ?, ?)
    """, (data["ip"], data["nemonico"], data["version"], data["num_switches"]))
    master_id = cursor.lastrowid  # ID del registro que acabamos de insertar

    for switch in data["switches"]:
        cursor.execute("""
        INSERT INTO SwitchDetails (master_id, mac_address, model_number, serial_number)
        VALUES (?, ?, ?, ?)
        """, (master_id, switch["mac_address"], switch.get("model_number", ""), switch.get("serial_number", "")))

    conn.commit()
    conn.close()
    
    # Mostrar un mensaje de éxito
    messagebox.showinfo("Información", "Reporte guardado exitosamente.")

def conectar():
    global connection
    ip = entry_ip.get()
    user = combo_users.get()
    password = entry_password.get()

    device = {
        'device_type': 'cisco_ios',  # Asumiendo que es un Cisco IOS, pero esto puede cambiar según el switch
        'ip': ip,
        'username': user,
        'password': password,
    }

    try:
        connection = ConnectHandler(**device)
        output.insert(tk.END, f"Conectado exitosamente a {ip}\n")
    except NetmikoTimeoutException:
        output.insert(tk.END, f"Tiempo de espera excedido al intentar conectar a {ip}\n")
    except NetmikoAuthenticationException:
        output.insert(tk.END, "Error de autenticación. Verifica usuario y contraseña.\n")

def actualizar_password(event):
    """Actualiza el campo de contraseña basado en el usuario seleccionado"""
    user = combo_users.get()
    if user in users:
        entry_password.delete(0, tk.END)
        entry_password.insert(0, users[user]) 
           
def enviar_comando():
    output.delete(1.0, tk.END)  # Limpiar el área de output
    if connection:
        comando = entry_comando.get()
        respuesta = connection.send_command(comando)
        output.insert(tk.END, f"Enviado: {comando}\nRecibido: {respuesta}\n")
    else:
        output.insert(tk.END, "Por favor, conecta primero al dispositivo.\n")

def agregar_usuario():
    global users  # Es importante referenciar la variable global

    user = combo_users.get()
    password = entry_password.get()

    if user and password:
        # Guardar en el archivo de usuarios
        users[user] = password  # Actualizamos el diccionario en memoria
        with open(USERS_FILE, 'w') as file:
            json.dump(users, file)  # Y luego guardamos ese diccionario en el archivo

        combo_users['values'] = list(users.keys())
        output.insert(tk.END, f"Usuario {user} agregado exitosamente\n")
    else:
        messagebox.showwarning("Advertencia", "Por favor, ingresa un usuario y contraseña")

def generar_reporte():
    global reporte  # <-- Declaramos reporte como global aquí
    if connection:
        comando = "show version"
        respuesta = connection.send_command(comando)
        reporte = procesar_show_version(respuesta)
        reporte["ip"] = entry_ip.get()
        output.delete(1.0, tk.END)  # Limpiar el output
        output.insert(tk.END, json.dumps(reporte, indent=4))  
    else:
        output.insert(tk.END, "Por favor, conecta primero al dispositivo.\n")

def procesar_show_version(output):
    reporte = {"nemonico": "", "version": "", "num_switches": 0, "switches": []}

    # 1. Obtener el nemónico
    nemonico_match = re.search(r"(\w+-\w+-\w+) uptime", output)
    if nemonico_match:
        reporte["nemonico"] = nemonico_match.group(1)

    # 2. Obtener la versión del software
    version_match = re.search(r"Cisco IOS Software,.*Version (\S+),", output)
    if version_match:
        reporte["version"] = version_match.group(1)

    # 3. Extraer secciones de cada switch
    switch_sections = re.split(r"Switch \d+", output)
    
    # 4. Obtener información de cada switch
    for section_str in switch_sections:
        mac_address_match = re.search(r"Base ethernet MAC Address\s+:\s+([\w:]+)", section_str)
        if mac_address_match:
            switch_data = {"mac_address": mac_address_match.group(1)}
            model_match = re.search(r"Model number\s+:\s+(\w+-\w+-\w+-\w+)", section_str)
            if model_match:
                switch_data["model_number"] = model_match.group(1)
            serial_match = re.search(r"System serial number\s+:\s+(\w+)", section_str)
            if serial_match:
                switch_data["serial_number"] = serial_match.group(1)
            reporte["switches"].append(switch_data)

    # Si no se encuentra nemónico en el formato usual, se busca el formato alternativo (e.g., "SW10-ALB")
    if not reporte["nemonico"]:
        nemonico_alt_match = re.search(r"(\w+-\w+) uptime", output)
        if nemonico_alt_match:
            reporte["nemonico"] = nemonico_alt_match.group(1)

    reporte["num_switches"] = len(reporte["switches"])        
    return reporte


# Interfaz del Programas

app = tk.Tk()
app.title("NetAnalyzer")
app.geometry("1000x800")
app.grid_rowconfigure(2, weight=1)
app.grid_columnconfigure(0, weight=1)


# label, botoenes, frame y combobox

frame = ttk.LabelFrame(app, text="Conexión")
frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew", columnspan=2)

label_ip = ttk.Label(frame, text="IP:")
label_ip.grid(row=0, column=0, sticky="w")
entry_ip = ttk.Entry(frame)
entry_ip.grid(row=0, column=1, sticky="ew")

label_user = ttk.Label(frame, text="Usuario:")
label_user.grid(row=1, column=0, sticky="w")
combo_users = ttk.Combobox(frame, values=list(users.keys()))
combo_users.grid(row=1, column=1, sticky="ew")
combo_users.set("Otro")
combo_users.bind("<<ComboboxSelected>>", actualizar_password)

label_password = ttk.Label(frame, text="Contraseña:")
label_password.grid(row=2, column=0, sticky="w")
entry_password = ttk.Entry(frame, show="*")
entry_password.grid(row=2, column=1, sticky="ew")

btn_conectar = ttk.Button(frame, text="Conectar", command=conectar)
btn_conectar.grid(row=3, column=1, pady=10, sticky="ew")
btn_agregar = ttk.Button(frame, text="Agregar Usuario", command=agregar_usuario)
btn_agregar.grid(row=3, column=2, pady=10, sticky="ew")

entry_comando = ttk.Entry(app, width=80)
entry_comando.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
btn_enviar = ttk.Button(app, text="Enviar Comando", command=enviar_comando)
btn_enviar.grid(row=1, column=1, padx=10, pady=10)

# Botón para generar reporte
btn_reporte = ttk.Button(app, text="Generar Reporte", command=generar_reporte)
btn_reporte.grid(row=1, column=3, padx=10, pady=5, sticky="ew")

# Botón para guardar el reporte
btn_guardar = ttk.Button(app, text="Guardar Reporte", command=lambda: guardar_en_db(reporte))
btn_guardar.grid(row=2, column=3, padx=10, pady=5, sticky="ew")

# Área de texto donde se mostrará el output
output = scrolledtext.ScrolledText(app, wrap=tk.WORD, width=80, height=30)
output.grid(row=3, column=0, padx=10, pady=10, sticky="nsew", columnspan=3, rowspan=2)  # rowspan=2 para que ocupe dos filas

# Botón para abrir ventana de informes
btn_informe = ttk.Button(app, text="Informe", command=ver_informe)
btn_informe.grid(row=3, column=3, padx=10, pady=5, sticky="ew")

app.mainloop()
