#!/usr/bin/env python
# coding: utf-8

# In[6]:


import os
import pandas as pd
import tkinter as tk
import json
import numpy as np
import requests
import pyodbc
import warnings
from datetime import datetime, timedelta
from sqlalchemy import create_engine
from PIL import Image, ImageTk
from openpyxl.styles import PatternFill
from openpyxl import Workbook
import xlsxwriter



def cargar_archivos():
    # Borra todo el contenido del Text widget
    text_output.delete('1.0', tk.END)

    ruta_carpeta = ruta_input.get() # Obtener la ruta de la carpeta

    if not os.path.exists(ruta_carpeta):
        mensaje_resultado.config(text="La ruta no existe.", fg="red")
        text_output.insert(tk.END, "La ruta no existe.\n")  # Mostrar mensaje en el Text widget
        
        return
    
    fecha = fecha_input.get()  # Obtener la fecha del Entry
    
    try:
        # Intentar convertir la fecha ingresada al formato AAAA-MM-DD
        fecha_valida = datetime.strptime(fecha, "%Y-%m-%d")
        text_output.insert(tk.END, f"La fecha {fecha} es válida.\n")  # Mostrar mensaje en el Text widget

        archivo = str(fecha) + ".xlsx"
        ruta_archivo = os.path.join(ruta_carpeta, archivo)
        mensaje = f"Leyendo archivo: {ruta_archivo}\n"
        text_output.insert(tk.END, mensaje)  # Mostrar mensaje en el Text widget
        
        # Leer el archivo Excel (hoja específica y rango)
        try:
            df = pd.read_excel(ruta_archivo, sheet_name="Efectividad")
            usuario = 'saocomct_camaras'
            contraseña = '1t&F)DQG6BLq'
            host = '190.90.160.5'
            puerto = '3306'
            base_de_datos = 'saocomct_camaras'
    
            # Crea la cadena de conexión
            cadena_conexion = f'mysql+mysqlconnector://{usuario}:{contraseña}@{host}:{puerto}/{base_de_datos}'
    
            try:
                # Crea el motor de conexión
                motor = create_engine(cadena_conexion)
                df.to_sql('efectividades_pesv', con=motor, if_exists='append', index=False)
                mensaje_resultado.config(text="Los datos se han insertado correctamente en la base de datos", fg="green")
                text_output.insert(tk.END, "Los datos se han insertado correctamente en la base de datos\n")  # Mostrar mensaje de error en el Text widget
    
                try: 
                    os.remove(ruta_archivo)
                    text_output.insert(tk.END, f"El archivo {ruta_archivo} ha sido eliminado correctamente.\n")  # Mostrar mensaje de error en el Text widget
                except Exception as e: 
                    text_output.insert(tk.END, f"Ocurrió un error al intentar eliminar el archivo: {e}\n")  # Mostrar mensaje de error en el Text widget
    
            except Exception as e: 
                text_output.insert(tk.END, "Ocurrió un error al insertar los datos\n")  # Mostrar mensaje de error en el Text widget
                text_output.insert(tk.END, f"{e}\n")  # Mostrar mensaje de error en el Text widget
                
        except Exception as e:
            mensaje = f"Error al leer el archivo {archivo}: {e}\n"
            text_output.insert(tk.END, mensaje)  # Mostrar mensaje de error en el Text widget

    except ValueError:
        # Si no se puede convertir, la fecha es inválida
        text_output.insert(tk.END, f"La fecha {fecha} no es válida. Debe tener el formato AAAA-MM-DD.\n")  # Mostrar mensaje en el Text widget

def programacion_revision():
    import tkinter as tk
    text_output.delete('1.0', tk.END)  

    # Conexión con la base de datos
    try:
        engine = create_engine('mysql+pymysql://saocomct_camaras:1t&F)DQG6BLq@190.90.160.5/saocomct_camaras')
        # Realiza la consulta
        df_ult_consulta = pd.read_sql("SELECT BUS FROM revision_pesv WHERE fecha_consulta = (SELECT MAX(fecha_consulta) FROM revision_pesv)" , con=engine)
        text_output.insert(tk.END, "Conexión exitosa con la base de datos\n")  # Mostrar mensaje en el Text widget            

            # Diccionario del total de alarmas y alarmas de interes del PESV
        data_alarmas = {
            'type': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 20, 29, 36, 47, 58, 59, 60, 61, 62, 63, 64, 74, 100, 160, 161, 162, 163, 164, 165, 166, 169, 172, 173, 174, 175, 176, 392],
            'Center description': [
                'Video loss alarm', 'Motion detection alarm', 'Camera-covering alarm', 'Abnormal storage alarm',
                'IO 1', 'IO 2', 'IO 3', 'IO 4', 'IO 5', 'IO 6', 'IO 7', 'IO 8', 'Emergency alarm', 
                'High-speed alarm', 'Low voltage alarm', 'Accelerometer alarm', 'Geo-fencing alarm', 
                'Illegal shutdown', 'Temperature alarm', 'Distance alarm', 
                'Alarm for abnormal temperature changes', 'Driver Fatigue', 'No driver', 
                'Phone Detection', 'Smoking Detection', 'Driver Distraction', 'Lane departure', 
                'Forward Collision Warning', 'Abnormal boot alarm', 'GPS Alarm', 'Speeding Alarm', 
                'Impeding violation', 'Following Distance Monitoring', 'Pedestrian Collision Warning', 
                'Yawning Detection', 'Left blind spot detection', 'Right blind spot detection', 
                'Seat Belt Detection', 'Rolling Stop Alarm', 'Left BSD warning', 'Left BSD alarm', 
                'Right BSD warning', 'Right BSD alarm', 'Forward blind area'
            ],
            'English description': [
                'Video loss alarm', 'Motion detection alarm', 'Camera-covering alarm', 'Abnormal storage alarm',
                'IO 1', 'IO 2', 'IO 3', 'IO 4', 'IO 5', 'IO 6', 'IO 7', 'IO 8', 'Emergency alarm', 
                'High-speed alarm', 'Low voltage alarm', 'Accelerometer alarm', 'Geo-fencing alarm', 
                'Illegal shutdown', 'Temperature alarm', 'Distance alarm', 
                'Alarm for abnormal temperature changes', 'Driver Fatigue', 'No driver', 
                'Phone Detection', 'Smoking Detection', 'Driver Distraction', 'Lane departure', 
                'Forward Collision Warning', 'Abnormal boot alarm', 'GPS Alarm', 'Speeding Alarm', 
                'Impeding violation', 'Following Distance Monitoring', 'Pedestrian Collision Warning', 
                'Yawning Detection', 'Left blind spot detection', 'Right blind spot detection', 
                'Seat Belt Detection', 'Rolling Stop Alarm', 'Left BSD warning', 'Left BSD alarm', 
                'Right BSD warning', 'Right BSD alarm', 'Forward blind area'
            ]
        }


        # Crear el DataFrame
        df_alarmas = pd.DataFrame(data_alarmas)
        
        # Crear un DataFrame con los datos proporcionados
        alarmas_pesv = {
            "Center description": [
                "High-speed alarm", 
                "Accelerometer alarm", 
                "Seat Belt Detection", 
                "Phone Detection", 
                "Driver Fatigue", 
                "Camera-covering alarm"
            ],
            "Tipo alarma": [
                "Alta velocidad", 
                "Aceleración", 
                "Cinturón de seguridad", 
                "Detección de celular", 
                "Fatiga del conductor", 
                "Video cubierto"
            ]
        }
        
        df_alarmas_pesv = pd.DataFrame(alarmas_pesv)

        # Sacar el Key para acceder a la API
        username = 'admin'
        password = 'admin'
        
        if __name__ == '__main__':
            url1_1 = 'http://181.143.106.68:12056/api/v1/basic/key?username=' + username + '&password=' + password
            response1_1 = requests.get(url1_1)
            
            if response1_1.status_code == 200:
                content1_1 = response1_1.content
        
        dic1_1 = json.loads(content1_1.decode('utf-8'))
        
        # Extraer la key
        key = dic1_1['data']['key']

        # Sacar el listado de dispositivos y dejar solo los de SAO
        if __name__ == '__main__':
            url2_2 = 'http://181.143.106.68:12056/api/v1/basic/devices?key=' + key
            response2_2 = requests.get(url2_2)
            
            if response2_2.status_code == 200:
                content2_2 = response2_2.content
        
        dic2_2 = json.loads(content2_2.decode('utf-8'))
        
        # Extraer los datos del diccionario
        data2_2 = dic2_2['data']
        
        # Crear el DataFrame
        df2_2 = pd.DataFrame(data2_2)
        
        df_dispositivos = df2_2.copy()
        
        df_dispositivos['BUS'] = df_dispositivos['carlicence'].str[:6]
        
        # Filtrar las filas que NO comiencen con 'MDO'
        df_dispositivos = df_dispositivos[~df_dispositivos['BUS'].str.startswith('MDO')]

        # Sacar la fecha del día anterior
        fecha_actual = datetime.now()
        fecha_dia_anterior = fecha_actual - timedelta(days=1)
        fecha_dia1 = fecha_actual + timedelta(days=1)
        fecha_dia2 = fecha_actual + timedelta(days=2)
        fecha_dia3 = fecha_actual + timedelta(days=3)

        lista_ult_consulta = df_ult_consulta['BUS'].tolist()
        lista_ult_consulta = sorted(lista_ult_consulta)
        
        # Sacar el ultimo bus consultado
        ultimo_bus = lista_ult_consulta[-1]
        
        df_dispositivos = df_dispositivos[['deviceid', 'BUS' ]]
        # Ordenar el dataframe de dispositivos
        df_dispositivos = df_dispositivos.sort_values(by='BUS', ascending=True)
        df_dispositivos = df_dispositivos.reset_index(drop=True)
        
        # Encontrar la posición del último bus consultado
        pos_ultimo_bus = df_dispositivos[df_dispositivos['BUS'] == ultimo_bus].index[0]
        
        buses_consultar_anterior = df_dispositivos.iloc[pos_ultimo_bus + 1: pos_ultimo_bus + 11]
        
        # Verificar si hay menos de 10 vehículos
        if len(buses_consultar_anterior) < 10:
            # Calcular cuántos vehículos faltan
            faltantes = 10 - len(buses_consultar_anterior)
            
            # Reutilizar las filas desde el principio del DataFrame para completar
            filas_extra = df_dispositivos.iloc[:faltantes]
            
            # Concatenar las filas restantes con las filas adicionales
            buses_consultar_anterior = pd.concat([buses_consultar_anterior, filas_extra])
            buses_consultar_anterior = buses_consultar_anterior.reset_index(drop=True)

        buses_consultar_anterior = buses_consultar_anterior.copy()
        buses_consultar_anterior['Fecha'] = fecha_actual.strftime('%Y-%m-%d')

        # Ordenar el dataframe de dispositivos
        df_dispositivos = df_dispositivos.sort_values(by='BUS', ascending=True)
        df_dispositivos = df_dispositivos.reset_index(drop=True)
        
        # Encontrar la posición del último bus consultado
        ultimo_bus = buses_consultar_anterior['BUS'].iloc[-1]
        
        pos_ultimo_bus = df_dispositivos[df_dispositivos['BUS'] == ultimo_bus].index[0]
        
        buses_consultar_actual = df_dispositivos.iloc[pos_ultimo_bus + 1: pos_ultimo_bus + 11]
        
        # Verificar si hay menos de 10 vehículos
        if len(buses_consultar_actual) < 10:
            # Calcular cuántos vehículos faltan
            faltantes = 10 - len(buses_consultar_actual)
            
            # Reutilizar las filas desde el principio del DataFrame para completar
            filas_extra = df_dispositivos.iloc[:faltantes]
            
            # Concatenar las filas restantes con las filas adicionales
            buses_consultar_actual = pd.concat([buses_consultar_actual, filas_extra])
            buses_consultar_actual = buses_consultar_actual.reset_index(drop=True)

        buses_consultar_actual = buses_consultar_actual.copy()
        buses_consultar_actual['Fecha'] = fecha_dia1.strftime('%Y-%m-%d')

        # Ordenar el dataframe de dispositivos
        df_dispositivos = df_dispositivos.sort_values(by='BUS', ascending=True)
        df_dispositivos = df_dispositivos.reset_index(drop=True)
        
        # Encontrar la posición del último bus consultado
        ultimo_bus = buses_consultar_actual['BUS'].iloc[-1]
        
        pos_ultimo_bus = df_dispositivos[df_dispositivos['BUS'] == ultimo_bus].index[0]
        
        buses_consultar_dia1 = df_dispositivos.iloc[pos_ultimo_bus + 1: pos_ultimo_bus + 11]
        
        # Verificar si hay menos de 10 vehículos
        if len(buses_consultar_dia1) < 10:
            # Calcular cuántos vehículos faltan
            faltantes = 10 - len(buses_consultar_dia1)
            
            # Reutilizar las filas desde el principio del DataFrame para completar
            filas_extra = df_dispositivos.iloc[:faltantes]
            
            # Concatenar las filas restantes con las filas adicionales
            buses_consultar_dia1 = pd.concat([buses_consultar_dia1, filas_extra])
            buses_consultar_dia1 = buses_consultar_dia1.reset_index(drop=True)

        buses_consultar_dia1 = buses_consultar_dia1.copy()
        buses_consultar_dia1['Fecha'] = fecha_dia2.strftime('%Y-%m-%d')

        # Ordenar el dataframe de dispositivos
        df_dispositivos = df_dispositivos.sort_values(by='BUS', ascending=True)
        df_dispositivos = df_dispositivos.reset_index(drop=True)
        
        # Encontrar la posición del último bus consultado
        ultimo_bus = buses_consultar_dia1['BUS'].iloc[-1]
        
        pos_ultimo_bus = df_dispositivos[df_dispositivos['BUS'] == ultimo_bus].index[0]
        
        buses_consultar_dia2 = df_dispositivos.iloc[pos_ultimo_bus + 1: pos_ultimo_bus + 11]
        
        # Verificar si hay menos de 10 vehículos
        if len(buses_consultar_dia2) < 10:
            # Calcular cuántos vehículos faltan
            faltantes = 10 - len(buses_consultar_dia2)
            
            # Reutilizar las filas desde el principio del DataFrame para completar
            filas_extra = df_dispositivos.iloc[:faltantes]
            
            # Concatenar las filas restantes con las filas adicionales
            buses_consultar_dia2 = pd.concat([buses_consultar_dia2, filas_extra])
            buses_consultar_dia2 = buses_consultar_dia2.reset_index(drop=True)

        buses_consultar_dia2 = buses_consultar_dia2.copy()
        buses_consultar_dia2['Fecha'] = fecha_dia3.strftime('%Y-%m-%d')

        df_programacion_alarmas = pd.concat([buses_consultar_anterior, buses_consultar_actual, buses_consultar_dia1, buses_consultar_dia2], ignore_index=True)
        df_programacion_alarmas = df_programacion_alarmas[['BUS', 'Fecha']] 

        ruta_carpeta = r'Z:\OPERACIONES\PUBLICA\SEGURIDAD OPERACIONAL\2. FACTOR HUMANO\VILLA\E-M-P\E-M-P\E-M-P CAMARAS APLICATIVO CEIBA II\CAMARAS 2024\1. ALARMAS\Efectividad'
        
        try:
            # Guardar el archivo de salida
            archivo_guardado = os.path.join(ruta_carpeta, 'Programacion_revisiones.xlsx')
            with pd.ExcelWriter(archivo_guardado, engine="xlsxwriter") as writer:
                df_programacion_alarmas.to_excel(writer, sheet_name='Revision', index=False)
                
                # Obtener objetos de Workbook y Worksheet
                workbook = writer.book
                worksheet = writer.sheets['Revision']
                
                # Formato Calibri 11
                formato_calibri_11 = workbook.add_format({'font_name': 'Calibri', 'font_size': 11})
                
                # Formato encabezado
                header_format1 = workbook.add_format({
                    'font_name': 'Calibri',
                    'font_size': 11,
                    'bold': True,
                    'bg_color': '#D3D3D3',  # Gris claro
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1
                })
                
                # Aplicar formato a los encabezados
                for col_num, header in enumerate(df_programacion_alarmas.columns):
                    worksheet.write(0, col_num, header, header_format1)
                
                # Ajustar el ancho de las columnas
                worksheet.set_column('A:A', 8)
                worksheet.set_column('B:B', 10)
        
                # Formato de celda con fondo amarillo
                formato_amarillo = workbook.add_format({'bg_color': '#FFFF00'})
        
                # Obtener la fecha de mañana
                fecha_manana = (datetime.now() + timedelta(days=1)).date()
        
                # Escribir los datos con formato
                for row_num, row_data in enumerate(df_programacion_alarmas.values, start=1):  # Datos inician en fila 1 (fila 0 es encabezado)
                    for col_num, cell_value in enumerate(row_data):
                        if col_num == 1:  # Suponiendo que la columna 2 (col_num = 1) es la fecha
                            try:
                                fecha_celda = pd.to_datetime(cell_value).date()
                                if fecha_celda == fecha_manana:
                                    worksheet.write(row_num, col_num, cell_value, formato_amarillo)
                                else:
                                    worksheet.write(row_num, col_num, cell_value, formato_calibri_11)
                            except Exception:
                                worksheet.write(row_num, col_num, cell_value, formato_calibri_11)
                        else:
                            worksheet.write(row_num, col_num, cell_value, formato_calibri_11)
        
            text_output.insert(tk.END, f"Archivo revision guardado como: {archivo_guardado}\n")
            mensaje_resultado.config(text=f"Archivo revision guardado como: {archivo_guardado}", fg="green")
        
        except Exception as e:
            text_output.insert(tk.END, f"Error al guardar el archivo: {e}\n")
            mensaje_resultado.config(text=f"Error al guardar el archivo: {e}", fg="red")


    except Exception as ex:
        mensaje2 = ex
        text_output.insert(tk.END, mensaje2)  # Mostrar mensaje en el Text widget            
        return        

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Efectividades pesv")  # Título de la ventana
ventana.geometry("1000x700")  # Tamaño de la ventana
ventana.config(bg="white")  # Fondo blanco

# Cargar la imagen
ruta_imagen = r"C:\Users\natalia.sanchez\Documents\3. Proyectos terminados\PESV alarmas\logo sao.png"
imagen = Image.open(ruta_imagen)

# Obtener las dimensiones originales de la imagen
ancho_original, alto_original = imagen.size

# Redimensionar la imagen a 250 píxeles de ancho manteniendo la relación de aspecto
nuevo_ancho = 250
nuevo_alto = int((nuevo_ancho / ancho_original) * alto_original)
imagen_redimensionada = imagen.resize((nuevo_ancho, nuevo_alto))

# Convertir la imagen redimensionada en un formato compatible con Tkinter
imagen_tk = ImageTk.PhotoImage(imagen_redimensionada)

# Crear un Label para la imagen y colocarla en la esquina superior izquierda
etiqueta_imagen = tk.Label(ventana, image=imagen_tk, bg="white")
etiqueta_imagen.place(x=10, y=10)  # Ubicar la imagen en la esquina superior izquierda

# Crear un Label para el título "EFECTIVIDADES PESV" con color verde oscuro y una fuente más delgada
titulo = tk.Label(ventana, text="EFECTIVIDADES PESV", font=("Aptos Narrow", 20, "bold"), bg="white", fg="#005639")
# Centrar el título en la ventana, debajo de la imagen
titulo.place(relx=0.5, rely=0.2, anchor="center")

# Crear un Label para el mensaje debajo del título
mensaje = tk.Label(ventana, text="Ingrese la ruta donde están los archivos de efectividades", font=("Aptos Narrow", 12), bg="white", fg="black")
# Centrar el mensaje en la ventana, justo debajo del título
mensaje.place(relx=0.5, rely=0.3, anchor="center")

# Crear un cuadro de texto (input) para ingresar la ruta con tamaño específico (por ejemplo, 40 caracteres de ancho)
ruta_input = tk.Entry(ventana, font=("Aptos Narrow", 11), width=80, bd=2, relief="solid", highlightthickness=2, highlightbackground="#746f74", highlightcolor="#746f74")
# Colocar el cuadro de texto debajo del mensaje
ruta_input.place(relx=0.5, rely=0.35, anchor="center")

# Crear un Label para el mensaje debajo del título
mensaje2 = tk.Label(ventana, text="Ingrese la fecha que procesó: (AAAA-MM-DD)", font=("Aptos Narrow", 12), bg="white", fg="black")
# Centrar el mensaje en la ventana, justo debajo del título
mensaje2.place(relx=0.5, rely=0.45, anchor="center")

# Crear un cuadro de texto (input) para ingresar la ruta con tamaño específico (por ejemplo, 40 caracteres de ancho)
fecha_input = tk.Entry(ventana, font=("Aptos Narrow", 11), width=40, bd=2, relief="solid", highlightthickness=2, highlightbackground="#746f74", highlightcolor="#746f74")
# Colocar el cuadro de texto debajo del mensaje
fecha_input.place(relx=0.5, rely=0.5, anchor="center")

# Crear un botón debajo del Entry con texto "Cargar Novedades", fondo verde medio y texto blanco
boton_cargar = tk.Button(ventana, text="CARGAR ARCHIVO", font=("Aptos Narrow", 11, "bold"), bg="#46913C", fg="white", relief="flat", width=21, height=2, command=cargar_archivos)
# Colocar el botón un poco más hacia el centro
boton_cargar.place(relx=0.37, rely=0.6, anchor="center")

# Crear un segundo botón alineado con el primero
boton_archivo_plano = tk.Button(ventana, text="PROG. REVISIÓN", font=("Aptos Narrow", 11, "bold"), bg="#46913C", fg="white", relief="flat", width=21, height=2, command=programacion_revision)
# Colocar el botón a la derecha del primero, más centrado
boton_archivo_plano.place(relx=0.63, rely=0.6, anchor="center")

# Crear un Label para mostrar el mensaje debajo del botón
mensaje_resultado = tk.Label(ventana, text="", font=("Aptos Narrow", 8), bg="white", fg="black")
# Centrar el mensaje debajo del botón
mensaje_resultado.place(relx=0.5, rely=0.7, anchor="center")

# Crear un Text widget para mostrar los mensajes de print
text_output = tk.Text(ventana, font=("Aptos Narrow", 10), height=10, width=100, wrap=tk.WORD, bd=2, relief="solid")
# Colocar el Text widget debajo del mensaje_resultado
text_output.place(relx=0.5, rely=0.85, anchor="center")

# Iniciar la interfaz
ventana.mainloop()


# In[ ]:




