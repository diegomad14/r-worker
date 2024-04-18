# Variable global para almacenar la referencia a la ventana principal
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import simpledialog
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from tkinter import ttk, Scrollbar, Entry
import time
from selenium.common.exceptions import NoSuchElementException
import csv
from tkinter import BooleanVar
from datetime import datetime
from PIL import Image, ImageTk
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


contactCountry_dict={'Argentina':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[3]',
    'Brasil': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[4]',
    'Chile': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[5]',
    'Colombia': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[6]',
    'Costa Rica': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[7]',
    'Ecuador': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[8]',
    'Mexico': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[9]',
    'Peru': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[10]',
    'Uruguay': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[11]',
}
contactType_dict={'Buzón':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[3]',
    'No contesta llamada': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[4]',
    'No enlaza llamada': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[5]',
    'Fuera de servicio': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[6]',
    'Número no existe': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[7]',
    'Llamada sin audio': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[8]',
    'Número equivocado': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[9]',
    'Falta de interacción en chat': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[10]',
    'No contesta correo': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[11]',
    'Cuelga llamada': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[12]',
    'Volver a llamar': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[13]',
    'Aliado no desea continuar su proceso': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[14]',
    'Ya no es restaurante': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[15]',
    'Presenta Bug/Incidencia en plataforma': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[16]',
    'Pendiente por revisión': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[17]',
    'Ayuda subir información': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[18]',
    'Compromete subir información': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[19]',
    'Problemas con credenciales': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[20]',
    'Ayuda por rechazo': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[21]',
    'Incidencias onboarding': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[22]',
    'Envía correo': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[23]',
    'Resuelven dudas': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[24]',
    'Rechazada por Flujo de BE':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[25]',
    'Fuera de Cobertura':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div[26]'
}

contactFalta = {
    'TYC':'//*[@id="i25"]',
    'Inf Ban.': '//*[@id="i31"]',
    'Hor.': '//*[@id="i43"]',
    'Menu': '//*[@id="i34"]',
    'H&L': '//*[@id="i37"]',
    'Doc':'//*[@id="i40"]',
    '':'//*[@id="i46"]'
}


class App(ttk.Frame):
    def __init__(self, parent):
        ttk.Frame.__init__(self)
        for index in [0, 1, 2]:
            self.columnconfigure(index=index, weight=1)
            self.rowconfigure(index=index, weight=1)
        self.var_0 = tk.BooleanVar()
        self.var_1 = tk.BooleanVar(value=True)
    
        self.setup_widgets()
    def setup_widgets(self):   
        # Create a Frame for the Opciones forms
        self.check_frame = ttk.LabelFrame(self, text="Opciones Forms", padding=(20, 10))
        self.check_frame.grid(
            row=2, column=0, padx=(20, 10), pady=(20, 10), sticky="nsew"
        )
        # Opciones Forms
        self.check_1 = ttk.Checkbutton(
            self.check_frame, text="Llamada", variable=self.var_0
        )
        self.check_1.grid(row=0, column=0, padx=5, pady=10, sticky="nsew")

        self.check_2 = ttk.Checkbutton(
            self.check_frame, text="Whats", variable=self.var_1
        )
        self.check_2.grid(row=0, column=1, padx=5, pady=10, sticky="nsew")

        # Create a Frame for the entryFrame
        self.entry_frame = ttk.LabelFrame(self, text="Archivo Maestro", padding=(20, 10))
        self.entry_frame.grid(row=0, column=0, padx=(20, 10), pady=(20, 10), sticky="nsew")
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        image_path = os.path.join(script_dir, "logo.png")
        # Load the logo image
        logo_image = Image.open(image_path)
        logo_photo = logo_image.resize((70, 70), Image.Resampling.LANCZOS)

        # Convert the image to Tkinter-compatible format
        logo_tk = ImageTk.PhotoImage(logo_photo)

        # Create a Label to display the logo
        logo_label = ttk.Label(self.entry_frame, image=logo_tk)
        logo_label.grid(row=0, column=0, padx=10, pady=10)

        # Set the logo image as a property of the label to prevent it from being garbage collected
        logo_label.image = logo_tk
        label_instrucciones = ttk.Label(self.entry_frame, text="Instrucciones de la aplicación:\n1. Carga el archivo CSV.\n2. Presiona el botón 'Procesar Datos' para comenzar el trabajo.")
        label_instrucciones.grid(row=0, column=0, columnspan=3, padx=10, pady=10) 

        label_ruta = ttk.Label(self.entry_frame, text="Ruta del archivo CSV:")
        label_ruta.grid(row=1, column=0, padx=10, pady=10)

        self.entry_ruta = ttk.Entry(self.entry_frame, width=50)
        self.entry_ruta.grid(row=1, column=1, padx=10, pady=10)

        boton_cargar = ttk.Button(self.entry_frame, text="Cargar Archivo", command=self.cargar_archivo)
        boton_cargar.grid(row=1, column=2, padx=10, pady=10)

        boton_procesar = ttk.Button(self.entry_frame, text="Procesar Datos", command=self.procesar_datos)
        boton_procesar.grid(row=4, column=2, columnspan=1, pady=10)
        
        boton_llenar_forms = ttk.Button(self.check_frame, text="Llenar Forms", command=self.llenar_forms)
        boton_llenar_forms.grid(row=1, column=0, columnspan=1, pady=10)


    def llenar_forms(self):
        ruta_archivo = self.entry_ruta.get()
        df= pd.read_csv(''+ruta_archivo+'')
        script_dir = os.path.dirname(os.path.abspath(__file__))
        driver_path = os.path.join(script_dir, "chromedriver")
        # Configura el path al controlador de Chrome
        service = Service(executable_path=driver_path)

        # Configura las opciones del navegador (puedes ajustar según tus necesidades)
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')  # Inicia la ventana maximizada

        # Crea una instancia del navegador Chrome
        driver = webdriver.Chrome(service=service, options=options)

        # Navega a la página del formulario
        driver.get('https://docs.google.com/forms/d/e/1FAIpQLSdlsZY3VlD7CfqiB9Ftm4X8cEuvpVU76D-Ku8u9NNhu_Z5FYg/viewform')

        messagebox.showinfo("Proceso completado","En cuanto inicie sesión por favor dar click aquí")       
        for row, data in df.iterrows():
            zcrm = data['ZCRM']
            country = data['País']
            storeID= data['Store ID']
            contactB = data['Contestó?']
            callsPE = data['que problemas calls?']
            chatPE = data['Problemas Whats?']
            falta = data['Estado']
            
            if self.var_0.get() and self.var_1.get():
                #Llamada
                driver.find_element('xpath', '//*[@id="i5"]').click() #Correo
                if pd.isna(storeID):
                    time.sleep(1)
                    driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(zcrm) #set nombre
                else:
                    time.sleep(1)
                    driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(storeID) #set StoreID
                    
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div[1]/div[1]').click() #País
                time.sleep(1)
                driver.find_element('xpath', contactCountry_dict[country]).click() # seleccionar País
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div[1]/div[1]').click() #Tipo de contacto
                time.sleep(1)
                if contactB == 'No' and pd.isna(callsPE):
                    driver.find_element('xpath', contactType_dict['No contesta llamada']).click() #Tipo de contacto
                    time.sleep(1)
                elif contactB == 'No' and not pd.isna(callsPE):
                    driver.find_element('xpath', contactType_dict[callsPE]).click() #Tipo de contacto
                    time.sleep(1)
                elif contactB == 'Si' and pd.isna(callsPE):
                    driver.find_element('xpath', contactType_dict['Compromete subir información']).click() #Tipo de contacto
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
                driver.find_element('xpath', '//*[@id="i62"]').click() #Canal Comunicación (Llamada)
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
                    
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div[1]/div[1]').click() #País
                time.sleep(1)
                driver.find_element('xpath', contactCountry_dict[country]).click() # seleccionar País
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div[1]/div[1]').click() #Tipo de contacto
                time.sleep(1)
                if pd.isna(chatPE):
                    driver.find_element('xpath', contactType_dict['Falta de interacción en chat']).click() #Tipo de contacto
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
                driver.find_element('xpath', '//*[@id="i56"]').click() #Canal Comunicación (WhatsApp)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[2]/div[1]/div/span').click() #Enviar (WhatsApp)
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div/div[4]/a').click() #again
                time.sleep(1)
            elif self.var_0.get() and not self.var_1.get():
                #Llamada
                driver.find_element('xpath', '//*[@id="i5"]').click() #Correo
                if pd.isna(storeID):
                    time.sleep(1)
                    driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(zcrm) #set nombre
                else:
                    time.sleep(1)
                    driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(storeID) #set StoreID
                    
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div[1]/div[1]').click() #País
                time.sleep(1)
                driver.find_element('xpath', contactCountry_dict[country]).click() # seleccionar País
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div[1]/div[1]').click() #Tipo de contacto
                time.sleep(1)
                if contactB == 'No' and pd.isna(callsPE):
                    driver.find_element('xpath', contactType_dict['No contesta llamada']).click() #Tipo de contacto
                    time.sleep(1)
                elif contactB == 'No' and not pd.isna(callsPE):
                    driver.find_element('xpath', contactType_dict[callsPE]).click() #Tipo de contacto
                    time.sleep(1)
                elif contactB == 'Si' and pd.isna(callsPE):
                    driver.find_element('xpath', contactType_dict['Compromete subir información']).click() #Tipo de contacto
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
                driver.find_element('xpath', '//*[@id="i62"]').click() #Canal Comunicación (Llamada)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[2]/div[1]/div/span').click() #Enviar (Llamada)
                time.sleep(1)
                #
                driver.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div/div[4]/a').click() #again
            elif not self.var_0.get() and self.var_1.get():
                #WhatsApp
                driver.find_element('xpath', '//*[@id="i5"]').click() #Correo
                time.sleep(1)
                if pd.isna(storeID):
                    time.sleep(1)
                    driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(zcrm) #put nombre
                else:
                    time.sleep(1)
                    driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(storeID) #put StoreID
                    
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div[1]/div[1]').click() #País
                time.sleep(1)
                driver.find_element('xpath', contactCountry_dict[country]).click() # seleccionar País
                time.sleep(1)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div[1]/div[1]').click() #Tipo de contacto
                time.sleep(1)
                if pd.isna(chatPE):
                    driver.find_element('xpath', contactType_dict['Falta de interacción en chat']).click() #Tipo de contacto
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
                driver.find_element('xpath', '//*[@id="i56"]').click() #Canal Comunicación (WhatsApp)
                driver.find_element('xpath', '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[2]/div[1]/div/span').click() #Enviar (WhatsApp)
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div/div[4]/a').click() #again
                time.sleep(1)
            else:
                driver.quit()
    def cargar_archivo(self):
        ruta_archivo = filedialog.askopenfilename(title="Seleccionar archivo CSV", filetypes=[("Archivos CSV", "*.csv")])
        if ruta_archivo:
            self.entry_ruta.delete(0, tk.END)
            self.entry_ruta.insert(0, ruta_archivo)
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
            llamadas = simpledialog.askinteger("Cantidad de Contactos", f"¿Cuántas contactos hará {agente_seleccionado}? (Max {num_filas})")
            self.mostrar_datos_relevantes(data=datos_filtrados, nCalls=llamadas, agente=agente_seleccionado)

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
        ventana_elegir_agente = tk.Toplevel(self)
        ventana_elegir_agente.title("Elegir Agente")

        label_instrucciones = ttk.Label(ventana_elegir_agente, text="Selecciona un agente:")
        label_instrucciones.pack(pady=10)

        # Cuadro de lista para mostrar los agentes
        lista_agentes = tk.Listbox(ventana_elegir_agente)
        for agente in agentes:
            lista_agentes.insert(tk.END, agente)
        lista_agentes.pack(pady=10)

        # Botón para confirmar la selección
        boton_confirmar = ttk.Button(
            ventana_elegir_agente,
            text="Confirmar",
            command=confirmar_seleccion
        )
        boton_confirmar.pack(pady=10)

        # Esperar hasta que el usuario elija un agente
        ventana_elegir_agente.wait_window()

        # Devolver el agente seleccionado
        return agente_seleccionado_var.get()


    def mostrar_datos_relevantes(self, data, nCalls, agente):
        # Crear una nueva ventana de Tkinter
        ventana_datos_relevantes = tk.Toplevel(self)
        ventana_datos_relevantes.title("Datos Relevantes")
        ventana_datos_relevantes.geometry("980x300")

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
        boton_whatsapp = ttk.Button(frame_datos, text="Enviar WhatsApps", command=lambda: self.sendWhats(database=data, agente=agente, nCalls=nCalls))
        boton_whatsapp.grid(row=0, column=0, columnspan=3, pady=10, padx=10, sticky="w")
        
        # Botón para exportar CSV
        boton_exportar_csv = ttk.Button(frame_datos, text="Exportar CSV", command=lambda: self.exportar_a_csv(entry_widgets, nCalls, checkbutton_vars, data, stringPCall_vars, stringPWhats_vars))
        boton_exportar_csv.grid(row=0, column=1, pady=10, padx=10, sticky="e")

        # Problemas call boolean
        label_seleccionar = ttk.Label(frame_datos, text="¿Problemas call?")
        label_seleccionar.grid(row=1, column=5, padx=10, pady=5)

        # Problemas Whats boolean
        label_seleccionar = ttk.Label(frame_datos, text="¿Contacto WhatsApp?")
        label_seleccionar.grid(row=1, column=6, padx=10, pady=5)

        entry_widgets = []
        checkbutton_vars = []
        stringPCall_vars = []
        stringPWhats_vars = []

        for call in range(nCalls):
            for i, columna in enumerate(["Onboarding Name", "Telefono 1", "Store ID"]):
                # Evitar la repetición de los nombres de las columnas
                if call == 0:
                    label_columna = ttk.Label(frame_datos, text=columna)
                    label_columna.grid(row=call * 4 + 1, column=i, padx=10, pady=5)

            # Agregar Label encima de "Comentarios"
            label_comentarios = ttk.Label(frame_datos, text="Comentarios")
            label_comentarios.grid(row=1, column=3, padx=10, pady=5)

            # Agregar Label encima de "¿Contestó?"
            label_contesto = ttk.Label(frame_datos, text="¿Contestó?")
            label_contesto.grid(row=1, column=4, padx=0, pady=0)
            

            for i, columna in enumerate(["Onboarding Name", "Telefono 1", "Store ID"]):
            # Evitar la repetición de los nombres de las columnas
                if call == 0:
                    label_columna = ttk.Label(frame_datos, text=columna)
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

            label_datos_Faltantes = ttk.Label(frame_datos, text="Falta: "+self.obtener_columnas_no(data.iloc[call]))
            label_datos_Faltantes.grid_remove()
            entry_widgets.append(label_datos_Faltantes)

            checkbutton_var = BooleanVar(value=False)
            checkbutton = ttk.Checkbutton(frame_datos, variable=checkbutton_var)
            checkbutton.grid(row=call * 4 + 2, column=4, padx=0, pady=0)
            entry_widgets.append(checkbutton)
            checkbutton_vars.append(checkbutton_var)

            # Agregar columna "Entry problemas call"
            stringPCall_var = tk.StringVar()
            entry_problemasCall = ttk.OptionMenu(frame_datos,stringPCall_var, *['','','Buzón','No enlaza llamada','Fuera de servicio','Número no existe','Llamada sin audio','Número equivocado','Cuelga llamada','Volver a llamar', 'Aliado no desea continuar su proceso', 'Ya no es restaurante','Presenta Bug/Incidencia en plataforma', 'Pendiente por revisión','Ayuda subir información','Problemas con credenciales','Ayuda por rechazo', 'Incidencias onboarding','Resuelven dudas','Rechazada por Flujo de BE','Fuera de Cobertura'])
            entry_problemasCall.config(width=5)
            entry_problemasCall.grid(row=call * 4 + 2, column=5, padx=10, pady=10)
            entry_problemasCall.nombre_opcion = "problemas_call"
            entry_widgets.append(entry_problemasCall)
            stringPCall_vars.append(stringPCall_var)


            # Agregar columna "Entry problemas Whats"
            stringPWhats_var = tk.StringVar()
            entry_problemasWhats = ttk.OptionMenu(frame_datos, stringPWhats_var, *['','','Aliado no desea continuar su proceso', 'Ya no es restaurante','Presenta Bug/Incidencia en plataforma', 'Pendiente por revisión','Ayuda subir información','Problemas con credenciales','Ayuda por rechazo', 'Incidencias onboarding','Resuelven dudas','Rechazada por Flujo de BE','Fuera de Cobertura'])
            entry_problemasWhats.config(width=5)
            entry_problemasWhats.grid(row=call * 4 + 2, column=6, padx=10, pady=10)
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

        entry_busqueda = tk.Entry(ventana_datos_relevantes)
        entry_busqueda.place(x=800, y=10)  # Colocar en la esquina superior derecha

        # Asociar el evento Enter a la función de búsqueda
        entry_busqueda.bind("<Return>", lambda event: buscar())
        # Asociar el evento KP_Enter (Enter numérico) a la función de búsqueda
        entry_busqueda.bind("<KP_Enter>", lambda event: buscar())

        # Función para buscar dentro de los datos
        def buscar(*args):
            # Obtener el término de búsqueda ingresado por el usuario
            termino_busqueda = entry_busqueda.get().lower()

            # Recorrer todos los widgets dentro de la ventana
            for widget in entry_widgets:
                # Verificar si el widget es un Entry o un Text
                if isinstance(widget, (tk.Entry, tk.Text)):
                    # Obtener el texto del widget y convertirlo a minúsculas para comparar
                    texto_widget = widget.get().lower()
                    # Verificar si el término de búsqueda está contenido en el texto del widget
                    if termino_busqueda in texto_widget:
                        # Resaltar o seleccionar el texto que coincide con el término de búsqueda
                        widget.focus_set()

        # Función para mover el cuadro de búsqueda cuando cambie el tamaño de la ventana
        def actualizar_posicion_busqueda(event):
            entry_busqueda.place(x=ventana_datos_relevantes.winfo_width() - 180, y=10)

        # Asociar la función de actualización de posición al evento de cambio de tamaño de la ventana
        ventana_datos_relevantes.bind("<Configure>", actualizar_posicion_busqueda)
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


    def sendWhats(self, database, agente, nCalls):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        driver_path = os.path.join(script_dir, "chromedriver")
        # Configura el path al controlador de Chrome
        service = Service(executable_path=driver_path)

        # Configura las opciones del navegador (puedes ajustar según tus necesidades)
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')  # Inicia la ventana maximizada

        # Crea una instancia del navegador Chrome
        driver = webdriver.Chrome(service=service, options=options)

        # Abre una nueva pestaña en el navegador con la página https://sales.treble.ai/
        driver.get('https://sales.treble.ai/')

        messagebox.showinfo("WhatsApp", "Al iniciar sesión en treble por favor dar click en Ok")
        time.sleep(1)

        driver.find_element('xpath', '/html/body/div/main/div/div[1]/header/nav/a[2]').click() #Click en contactos
        time.sleep(1)
        for iteration, (index, data) in enumerate(database.iterrows()):
            if iteration >= nCalls:
                break  # Salir del bucle después de nCalls iteraciones
            # Resto del código...
            cell = data['Telefono 1']
            nombre = data['Onboarding Name']
            pais =data['País']

            driver.find_element('xpath', '/html/body/div/main/div/div[1]/header/div/div/a').click()#Click en agregar contactos
            time.sleep(0.9)
            # Dividir la cadena en función del carácter especial
            subcadenas = str(nombre).split("|")

            # Verificar si se encontró el carácter especial
            if len(subcadenas) > 1:
                # Si se encontró, tomar la primera subcadena
                parte_deseada = subcadenas[0]
            else:
                # Si no se encontró, mantener la cadena original
                parte_deseada = nombre
            
            driver.find_element('xpath', '//*[@id="name"]').send_keys(parte_deseada)#Sendkeys nombre del contacto
            time.sleep(0.9)
                
            
            if pais == "Colombia":
                driver.find_element('xpath', '//*[@id="cellphone"]').send_keys(str(cell)[2:])#send keys celular (por defecto en colombia)
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/dialog/div/footer/button[2]').click()#click agregar contacto
                time.sleep(0.9)
            elif pais == "Mexico":
                #
                driver.find_element('xpath', '//*[@id="cellphone"]').send_keys(str(cell)[2:])#send keys celular (por defecto en colombia)
                time.sleep(0.9)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/button').click()
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/div/label/div/div/input').send_keys("Mexico")#buscar pais
                time.sleep(1)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/button[1]').click()
                time.sleep(0.9)

            elif pais == "Peru":
                #
                driver.find_element('xpath', '//*[@id="cellphone"]').send_keys(str(cell)[2:])#send keys celular (por defecto en colombia)
                time.sleep(0.9)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/button').click()
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/div/label/div/div/input').send_keys("Peru")#buscar pais
                time.sleep(1)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/button[1]').click()
                time.sleep(0.9)

            elif pais == "Argentina":
                #
                driver.find_element('xpath', '//*[@id="cellphone"]').send_keys(str(cell)[2:])#send keys celular (por defecto en colombia)
                time.sleep(0.9)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/button').click()
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/div/label/div/div/input').send_keys("Argentina")#buscar pais
                time.sleep(1)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/button[1]').click()
                time.sleep(0.9)

            elif pais == "Costa Rica":
                #
                driver.find_element('xpath', '//*[@id="cellphone"]').send_keys(str(cell)[3:])#send keys celular (por defecto en colombia)
                time.sleep(0.9)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/button').click()
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/div/label/div/div/input').send_keys("Costa Rica")#buscar pais
                time.sleep(1)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/button').click()
                time.sleep(0.9)

            elif pais == "Uruguay":
                #
                driver.find_element('xpath', '//*[@id="cellphone"]').send_keys(str(cell)[3:])#send keys celular (por defecto en colombia)
                time.sleep(0.9)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/button').click()
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/div/label/div/div/input').send_keys("Uruguay")#buscar pais
                time.sleep(1)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/button[1]').click()
                time.sleep(0.9)

            elif pais == "Chile":
                #
                driver.find_element('xpath', '//*[@id="cellphone"]').send_keys(str(cell)[2:])#send keys celular (por defecto en colombia)
                time.sleep(0.9)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/button').click()
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/div/label/div/div/input').send_keys("Chile")#buscar pais
                time.sleep(1)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/button[1]').click()
                time.sleep(0.9)
            elif pais == "Ecuador":
                #
                driver.find_element('xpath', '//*[@id="cellphone"]').send_keys(str(cell)[3:])#send keys celular (por defecto en colombia)
                time.sleep(0.9)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/button').click()
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/div/label/div/div/input').send_keys("Ecuador")#buscar pais
                time.sleep(1)
                driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/form/div/div/div[1]/div/button[1]').click()
                time.sleep(0.9)
            
            driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/footer/button[2]').click()#click agregar contacto
            time.sleep(1)
            # Verificar si el elemento está presente en la página
            try:
                if driver.find_element(By.XPATH, '//*[@id="cellphone-error"]'):
                    driver.find_element(By.XPATH, '/html/body/div/main/div/dialog/div/footer/button[1]').click()#click en cancelar
                    time.sleep(0.9)
                else:
                    print("El elemento no está presente en la página")
                    time.sleep(0.9)
            except NoSuchElementException:
                print("El elemento no está presente en la página")
                time.sleep(0.9)

            try:
                driver.find_element('xpath', '/html/body/div/main/div/div[1]/header/div/div/button').click()#click en buscar
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/div[1]/header/div/label/div/div/input').send_keys(cell)#SendKeys nombre del contacto
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div[1]/header/div/label/div/div/input').send_keys(Keys.ENTER)#SendKeys nombre del contacto
                time.sleep(1)
                driver.find_element('xpath', '/html/body/div/main/div/div[1]/div/div[1]/a[2]/div').click()#click en el primero que aparezca
                time.sleep(2)
            
                try:

                    driver.find_element('xpath', '/html/body/div/main/div/div[2]/div/div/a').click()# si el gestor de platillas está en el centro
                    time.sleep(1.3)
                except NoSuchElementException:
                    # Si no se encuentra el elemento con el primer XPath, intenta con el segundo
                    try:
                        driver.find_element('xpath', '/html/body/div/main/div/div[2]/div/div/div[2]/a').click()# si el gestor de platillas está abajo
                        time.sleep(1.3)
                    except NoSuchElementException:
                        print(f"No se pudo encontrar el elemento con ninguno de los XPaths proporcionados.")
                        time.sleep(1.3)
                driver.find_element('xpath', '/html/body/div/main/div/div[2]/div/div/label/div/div/input').send_keys("interacción_r2s")#SendKeys R2S
                driver.find_element('xpath', '/html/body/div/main/div/div[2]/div/div/label/div/div/input').send_keys(Keys.ENTER)#SendKeys nombre del contacto
                time.sleep(2)
                driver.find_element('xpath', '/html/body/div/main/div/div[2]/div/div/div/a').click()#Click en la plantilla
                time.sleep(1.3)
                driver.find_element('xpath', '/html/body/div/main/div/div[2]/div/div/button').click()#Click en enviar plantilla
                time.sleep(1.3)
                driver.find_element('xpath', '/html/body/div/main/div/div[2]/div/div/dialog/div/div/div/label[1]').click()#click en el número
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/div[2]/div/div/dialog/div/footer/button[2]').click()#click siguiente
                time.sleep(0.9)
                driver.find_element('xpath', '//*[@id="agente"]').send_keys(agente)#SendKeys el nombre del agente
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/div[2]/div/div/dialog/div/footer/button[2]').click()#click en enviar
                time.sleep(0.9)
                driver.find_element('xpath', '/html/body/div/main/div/div[1]/header/div/label/div/div/button').click()# darle a la x
                time.sleep(0.9)
            except:
                driver.find_element('xpath', '/html/body/div/main/div/div[1]/header/div/label/div/div/button').click()# darle a la x
                time.sleep(1.3)
            #Se repite :D
    def exportar_a_csv(self, entry_widgets, nCalls, checkbutton_vars, data, stringPCall_vars, stringPWhats_vars):
            # Generar un nombre sugerido basado en la fecha y hora actual
        nombre_sugerido = "datos_relevantes_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".csv"

            # Mostrar el diálogo de guardado de archivo con el nombre sugerido
        nombre_archivo = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=nombre_sugerido, filetypes=[("Archivos CSV", "*.csv")])
            
        if not nombre_archivo:  # Si el usuario cancela la operación de guardado
            return
        # Lista para almacenar los datos
        datos_exportar = []

        # Encabezados del archivo CSV
        encabezados = ["ZCRM" , "País", "Onboarding Name", "Telefono 1", "Store ID", "Comentarios", "Estado", "Contestó?","que problemas calls?","Problemas Whats?"]

        # Agregar encabezados a la lista de datos
        datos_exportar.append(encabezados)

        # Iterar sobre las entradas para recopilar datos
        for call in range(nCalls):
            fila_datos = []

            # Agregar los datos de la columna "ZCRM"
            valor_zcrm = str(data["ZCRM"].iloc[call])
            fila_datos.append(valor_zcrm)

            # Agregar los datos de la columna "Pais"
            valor_pais = str(data["País"].iloc[call])
            fila_datos.append(valor_pais)

            for i, entry_widget in enumerate(entry_widgets[call * 9:(call + 1) * 9]):
                if isinstance(entry_widget, tk.Entry):
                    valor = entry_widget.get()
                    fila_datos.append(valor)
                elif isinstance(entry_widget, ttk.Label):
                    valor = entry_widget.cget("text")
                    fila_datos.append(valor)
                elif isinstance(entry_widget, ttk.Progressbar):
                    valor=None
                elif isinstance(entry_widget, ttk.Checkbutton):
                    valor = 'Si' if checkbutton_vars[call].get() else 'No'
                    fila_datos.append(valor)
                elif isinstance(entry_widget, ttk.OptionMenu):
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

def main():
    root = tk.Tk()
    root.title("")

    # Obtener la ruta del directorio donde se encuentra este script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    theme_file = os.path.join(script_dir, "Azure", "azure.tcl")

    # Verificar si el archivo de tema existe antes de cargarlo
    if os.path.exists(theme_file):
        root.tk.call("source", theme_file)
        root.tk.call("set_theme", "light")
    else:
        print("¡Error! No se pudo encontrar el archivo de tema.")

    app = App(root)
    app.pack(fill="both", expand=True)

    # Set a minsize for the window, and place it in the middle
    root.update()
    root.minsize(root.winfo_width(), root.winfo_height())
    x_cordinate = int((root.winfo_screenwidth() / 2) - (root.winfo_width() / 2))
    y_cordinate = int((root.winfo_screenheight() / 2) - (root.winfo_height() / 2))
    root.geometry("+{}+{}".format(x_cordinate, y_cordinate-20))

    root.mainloop()

if __name__ == "__main__":
    main()






