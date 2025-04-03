# 📌 PESV ALARMAS EFECTIVIDADES

[![Estado del Proyecto](https://img.shields.io/badge/status-terminado-ogreen.svg)]()
[![Licencia](https://img.shields.io/badge/licencia-MIT-blue.svg)]()

## 🚀 Descripción  

**PESV - Efectividad de alarmas** El script automatiza la consulta, procesamiento y exportación de alarmas de los vehículos para el día vencido, facilitando la verificación de efectividad de las alarmas.

### 🎯 Propósito  
Crear un archivo de excel con el nombre del día anterior donde almacena las alarmas reportadas en 10 vehiculos.  

### 👥 Público objetivo  
Está dirigido a la **aprendiz de operaciones** de la empresa **Sistema Alimentador Oriental**, facilitando la gestión de revisión de efectividad de alarmas en la empresa Sistema Alimentador Oriental.  

## 📌 Funcionalidades  Cron_alarmas_PESV.py
### Información general
Esta instalado en el servidor de Power BI y se ejecuta diariamente a las 04:00 con un programador de tareas

### **1. Conexión a la base de datos MySQL**
- Se conecta a una base de datos MySQL que almacena información sobre revisiones del PESV.
- Obtiene la lista de los últimos buses consultados.

### **2. Definición de tipos de alarmas**
- Se crea un DataFrame con todos los tipos de alarmas disponibles y otro con las alarmas de interés para el PESV.

### **3. Autenticación en la API de monitoreo**
- Se solicita una clave de autenticación (`key`) para acceder a los datos de los dispositivos (cámaras en los buses).

### **4. Obtención del listado de dispositivos**
- Se consultan todos los dispositivos registrados en la API.
- Se filtran los buses que pertenecen a la empresa SAO (Sistema Alimentador Oriental).

### **5. Selección de buses a consultar**
- Se identifica el último bus consultado y se seleccionan los siguientes 10 buses en la lista ordenada.
- Si hay menos de 10 buses disponibles, se toman los primeros de la lista para completar el número requerido.

### **6. Consulta de alarmas**
- Se calcula la fecha del día anterior y se obtiene el rango de tiempo (00:00:00 - 23:59:59).
- Se consultan las alarmas de los buses seleccionados mediante la API.
- Se almacena la información en un DataFrame.

### **7. Procesamiento de alarmas**
- Se filtran las alarmas que corresponden al PESV.
- Se agregan datos como la fecha, el bus y la descripción de la alarma en español.
- Se genera un DataFrame con el detalle de alarmas.

### **8. Análisis de alarmas**
- Se cuenta el número total de alarmas por tipo, bus y fecha.
- Se genera una muestra del 50% de las alarmas para verificación manual.

### **9. Exportación a Excel**
- Se guardan los datos en un archivo Excel en una ruta específica con el nombre de la fecha del día anterior.
- Se aplican formatos a las hojas:
  - **"Efectividad"** (resumen de alarmas con conteo y muestra).
  - **"Alarmas"** (detalle de cada alarma reportada).

### **10. Último registro**
- Se obtiene el último bus analizado para continuar la consulta en la siguiente ejecución.

## 🔄 Entradas y Salidas  

### 📤 **Salidas**  
El sistema genera un archivo:  
1. **AAAA-MM-DD (`.xlsx`)**:  
   - Hoja **Efectividad** Lista de buses y alarmas que se revisan con la cantidad total de día y el calculo de la muestra. 
   - Hoja **Alarmas** detalle de cada alarma reportada.

## 📌 Requisitos  

Para ejecutar el proyecto, asegúrate de tener **Python 3.8 o superior** instalado y las siguientes librerías:  

### 📦 **Librerías necesarias**  

| 📦 Librería   | 🔍 Descripción |
|--------------|--------------|
| `pandas`     | Para la manipulación y análisis de datos mediante estructuras como DataFrame y Series |
| `requests`     | Para realizar peticiones HTTP a la API de monitoreo y obtener datos de los dispositivos y alarmas.. |
| `sqlalchemy`     | Para conectarse a la base de datos MySQL de manera eficiente y realizar consultas SQL sin necesidad de escribir sentencias SQL manualmente. |
| `numpy`   | Para cálculos numéricos y manejo de arrays, útil en el análisis de datos y optimización de operaciones en los DataFrame de pandas. |
| `json`   | Para manejar datos en formato JSON, especialmente cuando se reciben respuestas de la API de monitoreo. |
| `xlsxwriter`   | Para generar archivos Excel con formato, permitiendo agregar estilos y hojas a los reportes generados. |
| `pymysql`   | Un conector de MySQL para Python, usado para ejecutar consultas directas a la base de datos MySQL. |

### 🔧 **Instalación de librerías externas**  
Las siguientes librerías deben instalarse manualmente:  

```bash
pip install pandas requests sqlalchemy numpy json xlsxwriter pymysql
```

## 📌 Funcionalidades Efectividades PESV a BD.py

### Información general
Esta instalado en la carpeta Z:\OPERACIONES\PUBLICA\SEGURIDAD OPERACIONAL\2. FACTOR HUMANO\VILLA\E-M-P\E-M-P\E-M-P CAMARAS APLICATIVO CEIBA II\CAMARAS 2024\1. ALARMAS\Efectividad

### **1. Cargar y procesar archivos Excel para insertar datos en una base de datos MySQL (`Cargar archivos`)**
- Recibe la ruta de una carpeta donde se guarda el archivo de revisiones y una fecha ingresada por el usuario.
- Verifica si la ruta existe y si la fecha tiene el formato correcto (AAAA-MM-DD).
- Busca un archivo Excel con el nombre de la fecha (YYYY-MM-DD.xlsx) en la carpeta especificada.
- Intenta leer la hoja "Efectividad" del archivo y cargar los datos en una base de datos MySQL (saocomct_camaras).
- Si la carga es exitosa, elimina el archivo Excel de la carpeta.

### **2. Programar revisiones de buses para evaluar alarmas (`programacion_revision`)**
- Se conecta a la base de datos MySQL para obtener el último grupo de buses revisados.
- Define un conjunto de alarmas relevantes para el Programa de Seguridad Vial (PESV).
- Obtiene una clave (key) de autenticación para acceder a una API externa.
- Consulta la lista de dispositivos desde la API y extrae información sobre buses.
- Filtra y ordena los buses, seleccionando los próximos 10 que serán revisados diariamente durante los próximos 3 días.
- Guarda esta programación en un archivo Excel (Programacion_revisiones.xlsx) dentro de una ruta compartida de red.

## 📌 Requisitos  

Para ejecutar el proyecto, asegúrate de tener **Python 3.8 o superior** instalado y las siguientes librerías:  

### 📦 **Librerías necesarias**  

| 📦 Librería   | 🔍 Descripción |
|--------------|--------------|
| `pandas`     | Para la manipulación y análisis de datos mediante estructuras como DataFrame y Series |
| `requests`     | Para realizar peticiones HTTP a la API de monitoreo y obtener datos de los dispositivos y alarmas.. |
| `sqlalchemy`     | Para conectarse a la base de datos MySQL de manera eficiente y realizar consultas SQL sin necesidad de escribir sentencias SQL manualmente. |
| `Tkinter`   | Interfaz gráfica para capturar la ruta y fecha de entrada. |
| `xlsxwriter`   | Para generar archivos Excel con formato, permitiendo agregar estilos y hojas a los reportes generados. |



# 📖 Manual de Uso - Interfaz gráfica
![alt text](<interfaz.png>)


## 1️⃣ Instrucciones de Uso

1. Presionar doble click en **Efectividades PESV a BD.py**
2. Ingresar la ruta donde están los archivos de efectividades
3. Ingresar la fecha de las efectividades procesadas (AAAA-MM-DD)
4. Presionar el botón de **CARGAR ARCHIVO**
5. Revisar que en el cuadro de mensajes se haya leido correctamente la ruta y el nombre del archivo y que se haya guardado exitosamente en la base de datos
6. Presionar el botón de **PROG. REVISIÓN**
7. Abrir el archivo Programacion_revisiones.xlsx 


## 2️⃣ Contacto y Soporte

Para dudas o soporte técnico, contactar a la profesional de Mejoramiento Continuo de Sistema Alimentador Oriental.