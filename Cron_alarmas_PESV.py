#!/usr/bin/env python
# coding: utf-8

# ## Automatización PESV

# In[79]:


import pandas as pd
from sqlalchemy import create_engine

# Crea el motor de SQLAlchemy
# Asegúrate de reemplazar 'dialecto', 'usuario', 'contraseña', 'host', 'puerto' y 'nombre_bd' con tus datos reales
engine = create_engine('mysql+pymysql://saocomct_camaras:1t&F)DQG6BLq@190.90.160.5/saocomct_camaras')

# Realiza la consulta
df_ult_consulta = pd.read_sql("SELECT BUS FROM revision_pesv WHERE fecha_consulta = (SELECT MAX(fecha_consulta) FROM revision_pesv)" , con=engine)


# In[81]:


import pandas as pd
import json
import numpy as np
import requests
from datetime import datetime, timedelta
import os

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


# In[85]:


# Sacar el Key para acceder a la API
username = 'admin'
password = 'admin'

if __name__ == '__main__':
    url1_1 = 'http://192.168.90.68:12056/api/v1/basic/key?username=' + username + '&password=' + password
    response1_1 = requests.get(url1_1)
    
    if response1_1.status_code == 200:
        content1_1 = response1_1.content

dic1_1 = json.loads(content1_1.decode('utf-8'))

# Extraer la key
key = dic1_1['data']['key']
print(key)


# In[87]:


# Sacar el listado de dispositivos y dejar solo los de SAO
if __name__ == '__main__':
    url2_2 = 'http://192.168.90.68:12056/api/v1/basic/devices?key=' + key
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


# In[89]:


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

buses_consultar = df_dispositivos.iloc[pos_ultimo_bus + 1: pos_ultimo_bus + 11]

# Verificar si hay menos de 10 vehículos
if len(buses_consultar) < 10:
    # Calcular cuántos vehículos faltan
    faltantes = 10 - len(buses_consultar)
    
    # Reutilizar las filas desde el principio del DataFrame para completar
    filas_extra = df_dispositivos.iloc[:faltantes]
    
    # Concatenar las filas restantes con las filas adicionales
    buses_consultar = pd.concat([buses_consultar, filas_extra])

# Sacar la lista de los seriales de los dispositivos
main_terid = buses_consultar['deviceid'].tolist()


# In[91]:


# Sacar la fecha del día anterior
fecha_actual = datetime.now()
fecha_dia_anterior = fecha_actual - timedelta(days=1)

# Construir las fechas para la consulta de las alarmas
inicio_dia = datetime(fecha_dia_anterior.year, fecha_dia_anterior.month, fecha_dia_anterior.day, 0, 0, 0)  # 00:00:00
fin_dia = datetime(fecha_dia_anterior.year, fecha_dia_anterior.month, fecha_dia_anterior.day, 23, 59, 59)  # 23:59:59

start_time = inicio_dia.strftime('%Y-%m-%d %H:%M:%S')
end_time = fin_dia.strftime('%Y-%m-%d %H:%M:%S')


# In[93]:


df_detail_alarmas = pd.DataFrame()

if __name__ == '__main__':
    url4_2 = 'http://192.168.90.68:12056/api/v1/basic/alarm/detail'
    rq4_2 =  {
"key": key,
"terid": main_terid,
"type": [], 
"starttime": start_time,
"endtime": end_time
}

    response4_2 = requests.post(url4_2, json=rq4_2)
    
    if response4_2.status_code == 200:
        content4_2 = response4_2.content

# Creación del dataframe
import json
dic4_2 = json.loads(content4_2.decode('utf-8'))

import pandas as pd

# Extraer los datos del diccionario
data4_2 = dic4_2['data']

# Crear el DataFrame
df4_2 = pd.DataFrame(data4_2)
df_detail_alarmas = pd.concat([df_detail_alarmas, df4_2], ignore_index=True)

# Agregar el nombre de las alarmas
df_detail_alarmas = pd.merge(df_detail_alarmas, df_alarmas, on='type', how='left')


# In[94]:


# Filtrar solo por las alarmas del PESV
lista_alarmas_pesv = df_alarmas_pesv['Center description'].tolist()
lista_alarmas_pesv_español = df_alarmas_pesv['Tipo alarma'].tolist()

df_detail_alarmas = df_detail_alarmas[df_detail_alarmas['Center description'].isin(lista_alarmas_pesv)]
df_detail_alarmas = df_detail_alarmas.reset_index(drop=True)


# In[95]:


# Dar formato
df_detail_alarmas['Fecha'] = fecha_dia_anterior.strftime('%Y-%m-%d')

# Merge para agregar el BUS
df_detail_alarmas = pd.merge(df_detail_alarmas, df_dispositivos, left_on='terid', right_on='deviceid', how='left')

# Merge para agregar el Tipo alarma
df_detail_alarmas = pd.merge(df_detail_alarmas, df_alarmas_pesv, on='Center description', how='left')


# In[96]:


# Armar dataframe con el detalle que voy a guardar
df_detalle = df_detail_alarmas[['Fecha', 'BUS', 'Tipo alarma', 'time', 'content', 'cmdtype', 'speed']]

# Cambiar los nombres
df_detalle = df_detalle.rename(columns={'time': 'Hora', 'content': 'Contenido', 'cmdtype': 'Canal', 'speed': 'Velocidad'})


# In[97]:


# Hacer el groupby para sacar la cantidad de alarmas
df_grouped = df_detail_alarmas.groupby(['Fecha', 'BUS', 'Tipo alarma']).size().reset_index(name='Total alarmas')


# In[103]:


# Crear un DataFrame que tenga todos los tipos de alarmas
fechas_unidades = buses_consultar[['BUS']]
fechas_unidades['Fecha'] = fecha_dia_anterior.strftime('%Y-%m-%d')

idx = pd.MultiIndex.from_product(
    [fechas_unidades['Fecha'].drop_duplicates(), fechas_unidades['BUS'], lista_alarmas_pesv_español],
    names=['Fecha', 'BUS', 'Tipo alarma']
)

# Crear un DataFrame completo con todas las combinaciones posibles
df_completo = pd.DataFrame(index=idx).reset_index()

# Combinar el DataFrame original con el DataFrame completo
df_pesv = pd.merge(df_completo, df_grouped, on=['Fecha', 'BUS', 'Tipo alarma'], how='left')

# Rellenar los valores faltantes en 'Total alarmas' con 0
df_pesv['Total alarmas'] = df_pesv['Total alarmas'].fillna(0).astype(int)


# In[105]:


# Calcular la muestra 
df_pesv['Muestra'] = np.ceil(df_pesv['Total alarmas'] * 0.5).astype(int)
df_pesv['Aciertos'] = np.nan


# In[107]:


nombre_archivo = fecha_dia_anterior.strftime('%Y-%m-%d')
ruta = r"Z:\OPERACIONES\PUBLICA\SEGURIDAD OPERACIONAL\2. FACTOR HUMANO\VILLA\E-M-P\E-M-P\E-M-P CAMARAS APLICATIVO CEIBA II\CAMARAS 2024\1. ALARMAS\Efectividad"

try:
    # Guardar el archivo de salida
    archivo_completo = os.path.join(ruta, f"{nombre_archivo}.xlsx")
    with pd.ExcelWriter(archivo_completo, engine="xlsxwriter") as writer:
        # Escribir el DataFrame en la hoja
        df_pesv.to_excel(writer, sheet_name='Efectividad', index=False)
        df_detalle.to_excel(writer, sheet_name='Alarmas', index=False)

        # Obtener el objeto de la hoja
        workbook = writer.book
        worksheet = writer.sheets['Efectividad']
        
        # Formato para los encabezados con fondo gris claro
        header_format1 = workbook.add_format({
            'font_name': 'Calibri',
            'font_size': 11,
            'bold': True,
            'bg_color': '#D3D3D3',  # Color gris claro
            'align': 'center',      # Alinear al centro
            'valign': 'vcenter',    # Alinear verticalmente al centro
            'border': 1             # Bordes finos
        })

        # Formato para los encabezados con fondo azul claro
        header_format2 = workbook.add_format({
            'font_name': 'Calibri',
            'font_size': 11,
            'bold': True,
            'bg_color': '#C4D79B',  # Color verde claro
            'align': 'center',      # Alinear al centro
            'valign': 'vcenter',    # Alinear verticalmente al centro
            'border': 1             # Bordes finos
        })
        
        # Aplicar formato a los encabezados
        for col_num, header in enumerate(df_pesv.columns):
            worksheet.write(0, col_num, header, header_format1)

        # Rango específico (ejemplo: La columna F)
        rango_columnas = range(5, 6)  # Índices de las columnas a formatear
        
        # Aplicar formato a los encabezados en el rango especificado
        for col_num in rango_columnas:
            header = df_pesv.columns[col_num]
            worksheet.write(0, col_num, header, header_format2)
            
        # Ajustar el ancho de las columnas
        worksheet.set_column('A:A', 10)
        worksheet.set_column('B:B', 7)
        worksheet.set_column('C:C', 19)
        worksheet.set_column('D:D', 12)
        worksheet.set_column('E:E', 8)
        worksheet.set_column('F:F', 8)

        # Obtener el objeto de la hoja
        workbook = writer.book
        worksheet = writer.sheets['Alarmas']

        # Definir el formato para Calibri 11
        formato_calibri_9 = workbook.add_format({'font_name': 'Calibri', 'font_size': 9})
        
        # Aplicar el formato a todas las celdas
        for row_num, row_data in enumerate(df_detalle.values, start=1):  # Enumerar filas con datos
            for col_num in range(len(df_detalle.columns)):  # Enumerar columnas
                worksheet.write(row_num, col_num, row_data[col_num], formato_calibri_9)
        
        # Formato para los encabezados con fondo gris claro
        header_format1 = workbook.add_format({
            'font_name': 'Calibri',
            'font_size': 9,
            'bold': True,
            'bg_color': '#D3D3D3',  # Color gris claro
            'align': 'center',      # Alinear al centro
            'valign': 'vcenter',    # Alinear verticalmente al centro
            'border': 1             # Bordes finos
        })

        # Aplicar formato a los encabezados
        for col_num, header in enumerate(df_detalle.columns):
            worksheet.write(0, col_num, header, header_format1)

        # Ajustar el ancho de las columnas
        worksheet.set_column('A:A', 8)
        worksheet.set_column('B:B', 6)
        worksheet.set_column('C:C', 15)
        worksheet.set_column('D:D', 14)
        worksheet.set_column('E:E', 33)
        worksheet.set_column('F:F', 4)
        worksheet.set_column('G:G', 7)
    
    print(f"Archivo revision guardado como: {archivo_completo}")
except Exception as e:
    print(f"Error al guardar el archivo: {e}")


# In[109]:


fechas_unidades = fechas_unidades.reset_index(drop=True)
# Usar .tail() para obtener el último registro
ultimo_registro = fechas_unidades.tail(1)


# In[ ]:


# Crear una base de datos con los buses y la ultima fecha que se reviso la efectividad de las alarmas
import requests
from sqlalchemy import create_engine
# Guardar la información en la base de datos
from sqlalchemy import create_engine

usuario = 'saocomct_camaras'
contraseña = '1t&F)DQG6BLq'
host = '190.90.160.5'
puerto = '3306'
base_de_datos = 'saocomct_camaras'


# Crea la cadena de conexión
cadena_conexion = f'mysql+mysqlconnector://{usuario}:{contraseña}@{host}:{puerto}/{base_de_datos}'

# Crea el motor de conexión
motor = create_engine(cadena_conexion)

fecha_actual = datetime.now()


# Agregar la columna con la fecha y hora actuales a todas las filas
ultimo_registro['fecha_consulta'] = fecha_actual


ultimo_registro.to_sql('revision_pesv', con=motor, if_exists='append', index=False)


# In[ ]:




