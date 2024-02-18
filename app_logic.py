# Variable global para almacenar la referencia a la ventana principal
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import simpledialog
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from tkinter import ttk, Scrollbar, Entry
import time
from selenium.common.exceptions import NoSuchElementException
import csv
from tkinter import BooleanVar


contactType_dict={'Buzón':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[3]',
    'No contesta llamada': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[4]',
    'No enlaza llamada': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[5]',
    'Fuera de servicio': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[6]',
    'Número no existe': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[7]',
    'Llamada sin audio': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[8]',
    'Número equivocado': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[9]',
    'Falta de interacción en chat': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[10]',
    'No contesta correo': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[11]',
    'Cuelga llamada': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[12]',
    'Volver a llamar': '//*[@id="mG61Hd"]/div[2]/div/div [2]/div[3]/div/div/div[2]/div/div[2]/div[13]',
    'Aliado no desea continuar su proceso': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[14]',
    'Ya no es restaurante': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[15]',
    'Presenta Bug/Incidencia en plataforma': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[16]',
    'Pendiente por revisión': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[17]',
    'Ayuda subir información': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[18]',
    'Compromete subir información': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[19]',
    'Problemas con credenciales': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[20]',
    'Ayuda por rechazo': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[21]',
    'Incidencias onboarding': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[22]',
    'Envía correo': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[23]',
    'Resuelven dudas': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[24]'
}

contactFalta = {
    'TYC':'//*[@id="i21"]',
    'Inf Ban.': '//*[@id="i27"]',
    'Hor.': '//*[@id="i39"]',
    'Menu': '//*[@id="i30"]',
    'H&L': '//*[@id="i33"]',
    'Doc':'//*[@id="i36"]',
    '':'//*[@id="i42"]'
}


class AppLogic:
    def llenar_forms(self):
        ruta_archivo = self.entry_ruta.get()
        df= pd.read_csv(''+ruta_archivo+'')

        # Configura el path al controlador de Chrome
        service = Service(executable_path='chromedriver.exe')

        # Configura las opciones del navegador (puedes ajustar según tus necesidades)
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')  # Inicia la ventana maximizada

        # Crea una instancia del navegador Chrome
        driver = webdriver.Chrome(service=service, options=options)

        # Navega a la página de lupe.rappi.com
        driver.get('https://docs.google.com/forms/d/e/1FAIpQLSdlsZY3VlD7CfqiB9Ftm4X8cEuvpVU76D-Ku8u9NNhu_Z5FYg/viewform')

        messagebox.showinfo("Proceso completado","En cuanto inicie sesión por favor dar click aquí")

        for row, data in df.iterrows():
            zcrm = data['ZCRM']
            storeID= data['Store ID']
            contactB = data['Contestó?']
            callsPE = data['que problemas calls?']
            chatPE = data['Problemas Whats?']
            falta = data['Estado']
            
            driver.find_element('xpath', '//*[@id="i5"]').click() #Correo
            if pd.isna(storeID):
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(zcrm) #set nombre
            else:
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(storeID) #set StoreID
            driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div[1]/div[1]').click() #Tipo de contacto
            time.sleep(1)
            if contactB == 'No' and pd.isna(callsPE):
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[4]').click() #Tipo de contacto
                time.sleep(1)
            elif contactB == 'No' and not pd.isna(callsPE):
                driver.find_element('xpath', contactType_dict[callsPE]).click() #Tipo de contacto
                time.sleep(1)
            elif contactB == 'Si' and pd.isna(callsPE):
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[19]').click() #Tipo de contacto
                time.sleep(1)
            elif contactB == 'Si' and not pd.isna(callsPE):
                driver.find_element('xpath', contactType_dict[callsPE]).click() #Tipo de contacto
                time.sleep(1)

            # Separa las columnas que faltan
            columnas_faltantes = [columna.strip() for columna in str(falta).replace("Falta:", "").split(",")]

            # Itera sobre las columnas que faltan y completa el formulario
            for columna in columnas_faltantes:
                # Supongamos que el formulario tiene un campo de entrada para cada columna
                
                if columna=='nan':
                    driver.find_element('xpath',contactFalta['']).click()
                else:
                    driver.find_element('xpath',contactFalta[columna]).click()

            time.sleep(1)
            driver.find_element('xpath', '//*[@id="i58"]').click() #Canal Comunicación (Llamada)
            driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[2]/div[1]/div/span').click() #Enviar (Llamada)
            time.sleep(1)
            #
            driver.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div/div[4]/a').click() #again
            #WhatsApp
            driver.find_element('xpath', '//*[@id="i5"]').click() #Correo
            time.sleep(1)
            if pd.isna(storeID):
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(zcrm) #put nombre
            else:
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(storeID) #put StoreID
            driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div[1]/div[1]').click() #Tipo de contacto
            time.sleep(1)
            if pd.isna(chatPE):
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[10]').click() #Tipo de contacto
                time.sleep(1)
            elif not pd.isna(chatPE):
                driver.find_element('xpath', contactType_dict[chatPE]).click() #Tipo de contacto
                time.sleep(1)
            

            # Separa las columnas que faltan
            columnas_faltantes = [columna.strip() for columna in str(falta).replace("Falta:", "").split(",")]

            # Itera sobre las columnas que faltan y completa el formulario
            for columna in columnas_faltantes:
                # Supongamos que el formulario tiene un campo de entrada para cada columna
                if columna=='nan':
                    driver.find_element('xpath',contactFalta['']).click()
                else:
                    driver.find_element('xpath',contactFalta[columna]).click()
            time.sleep(1)
            driver.find_element('xpath', '//*[@id="i52"]').click() #Canal Comunicación (WhatsApp)
            driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[2]/div[1]/div/span').click() #Enviar (WhatsApp)
            time.sleep(1)
            driver.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div/div[4]/a').click() #again
            time.sleep(1)
    def __init__(self, master):
        def cargar_archivo():
            ruta_archivo = filedialog.askopenfilename(title="Seleccionar archivo CSV", filetypes=[("Archivos CSV", "*.csv")])
            if ruta_archivo:
                self.entry_ruta.delete(0, tk.END)
                self.entry_ruta.insert(0, ruta_archivo)

        def cambio_ventana():
            self.ventana_carga.destroy()  # Cerrar la ventana de carga
            self.ventana_principal.deiconify()  # Mostrar la ventana principal
        # Crear la ventana de carga
        self.ventana_carga = tk.Tk()
        self.ventana_carga.title("Cargando...")

        # Puedes personalizar la pantalla de carga según tus necesidades
        label_carga = tk.Label(self.ventana_carga, text="Cargando la aplicación, por favor espera...")
        label_carga.pack(pady=50)

        # Mostrar la ventana de carga
        self.ventana_carga.update()
        self.ventana_carga.after(1000, cambio_ventana)  # Después de 5 segundos, ejecutar la función cambio_ventana

        # Crear la ventana principal
        self.ventana_principal = master
        self.ventana_principal.withdraw()
        self.ventana_principal.title("Automatización Rappi")
        # Ocultar la ventana principal al inicio


        # Crear y posicionar los elementos en la ventana principal
        label_instrucciones = tk.Label(self.ventana_principal, text="Instrucciones de la aplicación:\n1. Carga el archivo CSV.\n2. Presiona el botón 'Procesar Datos' para comenzar el trabajo.")
        label_instrucciones.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        label_ruta = tk.Label(self.ventana_principal, text="Ruta del archivo CSV:")
        label_ruta.grid(row=1, column=0, padx=10, pady=10)

        self.entry_ruta = tk.Entry(self.ventana_principal, width=50)
        self.entry_ruta.grid(row=1, column=1, padx=10, pady=10)

        boton_cargar = tk.Button(self.ventana_principal, text="Cargar Archivo", command=cargar_archivo)
        boton_cargar.grid(row=1, column=2, padx=10, pady=10)

        boton_procesar = tk.Button(self.ventana_principal, text="Procesar Datos", command=self.procesar_datos)
        boton_procesar.grid(row=3, column=2, columnspan=1, pady=10)
        
        boton_llenar_forms = tk.Button(self.ventana_principal, text="Llenar Forms", command=self.llenar_forms)
        boton_llenar_forms.grid(row=3, column=0, columnspan=1, pady=10)

        # Iniciar el bucle principal de la interfaz gráfica
        self.ventana_principal.mainloop()
    def procesar_datos(self):
        # Lógica para procesar los datos
        ruta_archivo = self.entry_ruta.get()

        try:
            # Cargar el archivo CSV en un DataFrame
            datos_clientes = pd.read_csv(ruta_archivo)

            # Extraer nombres de agentes no repetidos
            agentes = datos_clientes['Agente'].unique()

            # Mostrar una ventana para que el usuario elija un agente
            agente_seleccionado = self.elegir_agente(agentes)

            # Filtrar los datos según el agente seleccionado
            datos_filtrados = datos_clientes[datos_clientes['Agente'] == agente_seleccionado]

            # Verificar si la columna 'Onboarding Name' está presente
            if 'Onboarding Name' not in datos_filtrados.columns:
                raise ValueError("La columna 'Onboarding Name' no está presente en los datos filtrados.")

            # Filtrar los datos que tienen valores nulos en la columna 'Comentario'
            datos_filtrados = datos_filtrados[datos_filtrados['Comentario'].isnull()]

            # Ordenar los datos por la columna 'Prioridad'
            datos_filtrados = datos_filtrados.sort_values(by='Aging',kind='stable')
            num_filas, num_columnas = datos_filtrados.shape

            # Pregunta cuántas llamadas hará la persona
            llamadas = simpledialog.askinteger("Cantidad de llamadas", f"¿Cuántas llamadas hará {agente_seleccionado}? (Max {num_filas})")
            self.abrir_navegador(data=datos_filtrados, nCalls=llamadas, agente=agente_seleccionado)

        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar los datos: {str(e)}")

    def elegir_agente(self, agentes):
        # Variable para almacenar la selección del usuario
        agente_seleccionado_var = tk.StringVar()

        # Función para manejar la selección y cerrar la ventana
        def confirmar_seleccion():
            agente_seleccionado_var.set(lista_agentes.get(tk.ACTIVE))
            ventana_elegir_agente.destroy()

        # Mostrar una ventana con una lista de agentes y permitir al usuario elegir uno
        ventana_elegir_agente = tk.Toplevel(self.ventana_principal)
        ventana_elegir_agente.title("Elegir Agente")

        label_instrucciones = tk.Label(ventana_elegir_agente, text="Selecciona un agente:")
        label_instrucciones.pack(pady=10)

        # Cuadro de lista para mostrar los agentes
        lista_agentes = tk.Listbox(ventana_elegir_agente)
        for agente in agentes:
            lista_agentes.insert(tk.END, agente)
        lista_agentes.pack(pady=10)

        # Botón para confirmar la selección
        boton_confirmar = tk.Button(
            ventana_elegir_agente,
            text="Confirmar",
            command=confirmar_seleccion
        )
        boton_confirmar.pack(pady=10)

        # Esperar hasta que el usuario elija un agente
        ventana_elegir_agente.wait_window()

        # Devolver el agente seleccionado
        return agente_seleccionado_var.get()
    def abrir_navegador(self, data, nCalls, agente):
        # Configura el path al controlador de Chrome
        service = Service(executable_path='chromedriver.exe')

        # Configura las opciones del navegador (puedes ajustar según tus necesidades)
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')  # Inicia la ventana maximizada

        # Crea una instancia del navegador Chrome
        driver = webdriver.Chrome(service=service, options=options)

        # Navega a la página de lupe.rappi.com
        driver.get('https://lupe.rappi.com/')
        messagebox.showinfo("Alerta","Al iniciar sesión hacer click aquí")
       # Crea una nueva ventana Tkinter para mostrar datos relevantes
        self.mostrar_datos_relevantes(data, nCalls=nCalls, agente=agente, driver=driver)


    def mostrar_datos_relevantes(self, data, nCalls, agente, driver):
        # Crear una nueva ventana de Tkinter
        ventana_datos_relevantes = tk.Toplevel(self.ventana_principal)
        ventana_datos_relevantes.title("Datos Relevantes")
        ventana_datos_relevantes.geometry("1150x600")

        # Crear un canvas para la barra de desplazamiento
        self.canvas = tk.Canvas(ventana_datos_relevantes)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Crear un frame para contener los datos y agregarlo al canvas
        frame_datos = ttk.Frame(self.canvas)
        frame_datos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Crear una barra de desplazamiento y conectarla con el canvas
        scrollbarV = Scrollbar(ventana_datos_relevantes, orient="vertical", command=self.canvas.yview)
        scrollbarV.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=scrollbarV.set)
        
        # Configurar el canvas para que la barra de desplazamiento funcione correctamente
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=frame_datos, anchor="nw")
        
        # Permitir desplazamiento con la rueda del ratón
        ventana_datos_relevantes.bind_all("<MouseWheel>", self.on_mousewheel)

        # Botón "Enviar WhatsApps" en la parte superior izquierda
        boton_whatsapp = tk.Button(frame_datos, text="Enviar WhatsApps", command=lambda: self.sendWhats(database=data, agente=agente, nCalls=nCalls,driver=driver))
        boton_whatsapp.grid(row=0, column=0, columnspan=3, pady=10, padx=10, sticky="w")
        
        # Botón para exportar CSV
        boton_exportar_csv = tk.Button(frame_datos, text="Exportar CSV", command=lambda: self.exportar_a_csv(entry_widgets, nCalls, checkbutton_vars, data, stringPCall_vars, stringPWhats_vars))
        boton_exportar_csv.grid(row=0, column=1, pady=10, padx=10, sticky="e")

        # Problemas call boolean
        label_seleccionar = tk.Label(frame_datos, text="¿Problemas call?")
        label_seleccionar.grid(row=1, column=5, padx=10, pady=5)

        # Problemas Whats boolean
        label_seleccionar = tk.Label(frame_datos, text="¿Problemas Whats?")
        label_seleccionar.grid(row=1, column=6, padx=10, pady=5)

        entry_widgets = []
        checkbutton_vars = []
        stringPCall_vars = []
        stringPWhats_vars = []

        for call in range(nCalls):
            for i, columna in enumerate(["Onboarding Name", "Telefono 1", "Store ID"]):
                # Evitar la repetición de los nombres de las columnas
                if call == 0:
                    label_columna = tk.Label(frame_datos, text=columna)
                    label_columna.grid(row=call * 4 + 1, column=i, padx=10, pady=5)

            # Agregar Label encima de "Comentarios"
            label_comentarios = tk.Label(frame_datos, text="Comentarios")
            label_comentarios.grid(row=1, column=3, padx=10, pady=5)

            # Agregar Label encima de "¿Contestó?"
            label_contesto = tk.Label(frame_datos, text="¿Contestó?")
            label_contesto.grid(row=1, column=4, padx=0, pady=0)
            

            for i, columna in enumerate(["Onboarding Name", "Telefono 1", "Store ID"]):
            # Evitar la repetición de los nombres de las columnas
                if call == 0:
                    label_columna = tk.Label(frame_datos, text=columna)
                    label_columna.grid(row=call * 4 + 1, column=i, padx=10, pady=5)

                # Obtener los datos relevantes de la fila correspondiente al número de llamada
                valor_dato = str(data[columna].iloc[call])

                # Usar Entry en lugar de Label
                entry = Entry(frame_datos, width=15)
                entry.insert(tk.END, valor_dato)
                entry.grid(row=call * 4 + 2, column=i, padx=10, pady=5)

                # Añadir el widget Entry a la lista para futuras referencias
                entry_widgets.append(entry)

            # Agregar columna "Comentarios"
            entry_comentarios = Entry(frame_datos, width=15)
            entry_comentarios.grid(row=call * 4 + 2, column=3, padx=10, pady=5)
            entry_widgets.append(entry_comentarios)

            progreso = self.calcular_progreso(data.iloc[call])
            process_bar = ttk.Progressbar(frame_datos, orient='horizontal', mode='determinate')
            process_bar['value']=(progreso[1]/progreso[0])*100
            process_bar.grid(row=call * 4 + 2, column=7, padx=0, pady=0)
            entry_widgets.append(process_bar)

            label_datos_Faltantes = tk.Label(frame_datos, text="Falta: "+self.obtener_columnas_no(data.iloc[call]))
            label_datos_Faltantes.grid_remove()
            entry_widgets.append(label_datos_Faltantes)

            checkbutton_var = BooleanVar(value=False)
            checkbutton = tk.Checkbutton(frame_datos, variable=checkbutton_var)
            checkbutton.grid(row=call * 4 + 2, column=4, padx=0, pady=0)
            entry_widgets.append(checkbutton)
            checkbutton_vars.append(checkbutton_var)

            # Agregar columna "Entry problemas call"
            stringPCall_var = tk.StringVar()
            entry_problemasCall = tk.OptionMenu(frame_datos,stringPCall_var, *['','Buzón','No enlaza llamada','Fuera de servicio','Número no existe','Llamada sin audio','Número equivocado','Cuelga llamada','Volver a llamar', 'Aliado no desea continuar su proceso', 'Ya no es restaurante','Presenta Bug/Incidencia en plataforma', 'Pendiente por revisión','Ayuda subir información','Problemas con credenciales','Ayuda por rechazo', 'Incidencias onboarding','Resuelven dudas'])
            entry_problemasCall.config(width=5)
            entry_problemasCall.grid(row=call * 4 + 2, column=5, padx=0, pady=0)
            entry_problemasCall.nombre_opcion = "problemas_call"
            entry_widgets.append(entry_problemasCall)
            stringPCall_vars.append(stringPCall_var)


            # Agregar columna "Entry problemas Whats"
            stringPWhats_var = tk.StringVar()
            entry_problemasWhats = tk.OptionMenu(frame_datos, stringPWhats_var, *['', 'Aliado no desea continuar su proceso', 'Ya no es restaurante','Presenta Bug/Incidencia en plataforma', 'Pendiente por revisión','Ayuda subir información','Problemas con credenciales','Ayuda por rechazo', 'Incidencias onboarding','Resuelven dudas'])
            entry_problemasWhats.config(width=5)
            entry_problemasWhats.grid(row=call * 4 + 2, column=6, padx=0, pady=0)
            entry_problemasWhats.nombre_opcion = "problemas_whats"
            entry_widgets.append(entry_problemasWhats)
            stringPWhats_vars.append(stringPWhats_var)


        # ... Puedes agregar más elementos según tus necesidades

        # Hacer que la ventana siempre esté al frente
        ventana_datos_relevantes.attributes('-topmost', True)
        ventana_datos_relevantes.lift()

        # Configurar la barra de desplazamiento para que funcione con el canvas
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        # Permitir la selección de texto en los widgets Entry
        for entry_widget in entry_widgets:
            entry_widget.bind("<FocusIn>", lambda event: event.widget.select_range(0, tk.END))

        # Iniciar el bucle principal de la ventana de Tkinter
        ventana_datos_relevantes.mainloop()

    def on_mousewheel(self, event):
            if event.delta > 0:
                self.canvas.yview_scroll(-1, "units")
            elif event.delta < 0:
                self.canvas.yview_scroll(1, "units")

    
    def calcular_progreso(self, fila):
        columnas_progreso = ["TYC", "Inf Ban.", "Hor.", "Menu", "H&L", "Doc"]
        total_columnas = len(columnas_progreso)
        completadas = sum(fila[col] == 'Si' for col in columnas_progreso)
        return total_columnas, completadas
    
    def obtener_columnas_no(self, fila):
        columnas_evaluadas = ["TYC", "Inf Ban.", "Hor.", "Menu", "H&L", "Doc"]
        columnas_no = [col for col in columnas_evaluadas if fila[col] == 'No']
        return ', '.join(columnas_no)


    def sendWhats(self, database, agente, nCalls, driver):

        # Abre una nueva pestaña en el navegador con la página https://sales.treble.ai/
        driver.execute_script("window.open('https://sales.treble.ai/', '_blank')")

        # Cambia el enfoque a la nueva pestaña
        driver.switch_to.window(driver.window_handles[1])

        messagebox.showinfo("WhatsApp", "Al iniciar sesión en treble por favor dar click en Ok")
        time.sleep(1)

        driver.find_element('xpath', '/html/body/div/main/div/div/div[1]/div[1]/div[2]').click() #Click en contactos

        for iteration, (index, data) in enumerate(database.iterrows()):
            if iteration >= nCalls:
                break  # Salir del bucle después de nCalls iteraciones
            # Resto del código...
            cell = data['Telefono 1']
            nombre = data['Onboarding Name']
            pais =data['País']

            driver.find_element('xpath', '/html/body/div/main/div/div/div[1]/div[2]/div/div[1]/button').click()#Click en agregar contactos
            time.sleep(1)
            # Dividir la cadena en función del carácter especial
            subcadenas = str(nombre).split("|")

            # Verificar si se encontró el carácter especial
            if len(subcadenas) > 1:
                # Si se encontró, tomar la primera subcadena
                parte_deseada = subcadenas[0]
            else:
                # Si no se encontró, mantener la cadena original
                parte_deseada = nombre
            
            driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[1]/div/div/input').send_keys(parte_deseada)#Sendkeys nombre del contacto
            time.sleep(1)
                
            
            if pais == "Colombia":
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/input').send_keys(str(cell)[2:])#send keys celular (por defecto en colombia)
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[4]/div/button[2]').click()#click agregar contacto
                time.sleep(1)
            elif pais == "Mexico":
                #
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/input').send_keys(str(cell)[2:])#send keys celular (por defecto en colombia)
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/div/div/div/div/div/a').click()
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="Mexico"]').click()
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[4]/div/button[2]').click()#click agregar contacto
                time.sleep(1)
            elif pais == "Peru":
                #
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/input').send_keys(str(cell)[2:])#send keys celular (por defecto en colombia)
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/div/div/div/div/div/a').click()
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="Peru"]').click()
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[4]/div/button[2]').click()#click agregar contacto
                time.sleep(1)
            elif pais == "Argentina":
                #
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/input').send_keys(str(cell)[2:])#send keys celular (por defecto en colombia)
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/div/div/div/div/div/a').click()
                driver.find_element('xpath', '//*[@id="Argentina"]').click()
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[4]/div/button[2]').click()#click agregar contacto
                time.sleep(1)
            elif pais == "Costa Rica":
                #
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/input').send_keys(str(cell)[3:])#send keys celular (por defecto en colombia)
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/div/div/div/div/div/a').click()
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="Costa Rica"]').click()
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[4]/div/button[2]').click()#click agregar contacto
                time.sleep(1)
            elif pais == "Uruguay":
                #
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/input').send_keys(str(cell)[3:])#send keys celular (por defecto en colombia)
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/div/div/div/div/div/a').click()
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="Uruguay"]').click()
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[4]/div/button[2]').click()#click agregar contacto
                time.sleep(1)
            elif pais == "Chile":
                #
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/input').send_keys(str(cell)[2:])#send keys celular (por defecto en colombia)
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/div/div/div/div/div/a').click()
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="Chile"]').click()
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[4]/div/button[2]').click()#click agregar contacto
                time.sleep(1)
            elif pais == "Ecuador":
                #
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/input').send_keys(str(cell)[3:])#send keys celular (por defecto en colombia)
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/div/div/div/div/div/div/a').click()
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="Ecuador"]').click()
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[4]/div/button[2]').click()#click agregar contacto
                time.sleep(1)
            try:
                driver.find_element('xpath', '/html/body/div/main/div/div/div[1]/div[2]/div/div[2]').click()#click en buscar
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[1]/div[2]/div/input').send_keys(parte_deseada)#SendKeys nombre del contacto
                time.sleep(3)
                driver.find_element('xpath', '//*[@id="contacts-scrollable"]/div/div/div[1]').click()#click en el primero que aparezca
                time.sleep(4)
            
                try:

                    driver.find_element('xpath', '/html/body/div[1]/div[3]/div[3]/div[2]/div/div/div[2]/div[2]/div/div[1]/div/button').click()# si el gestor de platillas está en el centro
                    time.sleep(1)
                except NoSuchElementException:
                    # Si no se encuentra el elemento con el primer XPath, intenta con el segundo
                    try:
                        driver.find_element('xpath', '/html/body/div/main/div/div/div[2]/div/div/div[2]/div[2]/div/div[1]/div/button').click()# si el gestor de platillas está abajo
                        time.sleep(1)
                    except NoSuchElementException:
                        print(f"No se pudo encontrar el elemento con ninguno de los XPaths proporcionados.")
                        time.sleep(1)
                driver.find_element('xpath', '/html/body/div[1]/main/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div[2]/input').send_keys("R2S")#SendKeys R2S
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div[1]/main/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div[1]').click()#Click en la plantilla
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[3]/div[2]/div/input').send_keys(agente)#SendKeys el nombre del agente
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[3]/div[2]/div[4]/div/div[2]').click()#click en enviar
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div/div[1]/div[2]/div/div[2]').click()# darle a la x
                time.sleep(1)
            except:
                driver.find_element('xpath', '/html/body/div/main/div/div/div[1]/div[2]/div/div[2]').click()# darle a la x
                time.sleep(1)
            #Se repite :D
    def exportar_a_csv(self, entry_widgets, nCalls, checkbutton_vars, data, stringPCall_vars, stringPWhats_vars):

        # Nombre del archivo CSV
        nombre_archivo = "datos_relevantes.csv"

        # Lista para almacenar los datos
        datos_exportar = []

        # Encabezados del archivo CSV
        encabezados = ["ZCRM" ,"Onboarding Name", "Telefono 1", "Store ID", "Comentarios", "Estado", "Contestó?","que problemas calls?","Problemas Whats?"]

        # Agregar encabezados a la lista de datos
        datos_exportar.append(encabezados)

        # Iterar sobre las entradas para recopilar datos
        for call in range(nCalls):
            fila_datos = []

            # Agregar los datos de la columna "ZCRM"
            valor_zcrm = str(data["ZCRM"].iloc[call])
            fila_datos.append(valor_zcrm)

            for i, entry_widget in enumerate(entry_widgets[call * 9:(call + 1) * 9]):
                if isinstance(entry_widget, tk.Entry):
                    valor = entry_widget.get()
                    fila_datos.append(valor)
                elif isinstance(entry_widget, tk.Label):
                    valor = entry_widget.cget("text")
                    fila_datos.append(valor)
                elif isinstance(entry_widget, ttk.Progressbar):
                    valor=None
                elif isinstance(entry_widget, tk.Checkbutton):
                    valor = 'Si' if checkbutton_vars[call].get() else 'No'
                    fila_datos.append(valor)
                elif isinstance(entry_widget, tk.OptionMenu):
                    if entry_widget.nombre_opcion == "problemas_call":
                        valor = stringPCall_vars[call].get()
                        fila_datos.append(valor)
                    elif entry_widget.nombre_opcion == "problemas_whats":
                        valor = stringPWhats_vars[call].get()
                        fila_datos.append(valor)
                else:
                    valor = ""
                    fila_datos.append(valor)

            # Agregar la fila de datos a la lista
            datos_exportar.append(fila_datos)

        # Escribir datos en el archivo CSV
        with open(nombre_archivo, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(datos_exportar)

        messagebox.showinfo("Exportación exitosa", f"Datos exportados a {nombre_archivo}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppLogic(root)

    root.mainloop()









