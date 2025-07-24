# -*- coding: utf-8 -*-
import requests
from datetime import datetime, timedelta
import os


# !Obtiene los tipos de cambio del Banco de México para "Para solventar obligaciones" y 
# !"Determinación" para el día actual y el día anterior.
# Token de Banxico
token = "5f05e502aba65738d90d9ee9c1ccb65ab52f100d37499a42404c7f9bfdb1dc64"

# Fecha actual y fecha anterior en formato YYYY-MM-DD
fecha_actual = datetime.now().strftime("%Y-%m-%d")


fecha_ayer = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

# Encabezados
headers = {
    "Bmx-Token": token
}

# --- Consulta 1: SF60653 y SF43718 para el día actual ---
series_dia_actual = ["SF60653", "SF43718"]
series_str_actual = ",".join(series_dia_actual)
url_actual = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_str_actual}/datos/{fecha_actual}/{fecha_actual}"

response_actual = requests.get(url_actual, headers=headers)

if response_actual.status_code == 200:
    data = response_actual.json()
    series = data.get("bmx", {}).get("series", [])
    tc = ""
    for serie in series:
        if serie["idSerie"] == "SF60653":
            serie["titulo"] = "Para solventar obligaciones"
        elif serie["idSerie"] == "SF43718":
                serie["titulo"] = "Determinacion"
        nombre = serie["titulo"]
        try:
            # Si hay datos, extraer el valor y la fecha
            valor = serie["datos"][0]["dato"]
            fecha = serie["datos"][0]["fecha"]
        except:
            # Si no hay datos, asignar valores por defecto
            valor = "N/E"
            fecha = datetime.now().strftime("%d/%m/%Y")
        #Concatenear serie_id, nombre, valor y fecha para que se una sola variable
        if tc:  # If tc already has content, add separator
            tc = f"{nombre}:{valor}:{fecha}" + ";" + tc
        else:  # First iteration
            tc = f"{nombre}:{valor}:{fecha}"
else:
    t = f"Error en consulta actual: {response_actual.status_code}"

    # --- Consulta 2: Sólo SF43718 para el día anterior ---
serie_ayer = "SF43718"
url_ayer = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{serie_ayer}/datos/{fecha_ayer}/{fecha_ayer}"

response_ayer = requests.get(url_ayer, headers=headers)

if response_ayer.status_code == 200:
    data = response_ayer.json()
    tmp = "0"
    serie = data.get("bmx", {}).get("series", [])[0]
    serie["titulo"] = "DOF"
    nombre = serie["titulo"]
    try:
        valor = serie["datos"][0]["dato"]
    except:
        valor = "N/E"
    tdc = f"{nombre}:{valor}:{tmp}"
    tmp =  f"{tc};{tdc}"

    print(tmp)

    # obtener la ruta de este script
    ruta_actual = os.path.dirname(os.path.abspath(__file__))
    print(f"Ruta actual: {ruta_actual}")
    
    # Escribir la variable tmp en un archivo de texto
    try:
        with open(os.path.join(ruta_actual, "tipos_de_cambio_usd.txt"), "w", encoding="utf-8") as archivo:
            archivo.write(tmp)
        print("Archivo 'tipos_de_cambio.txt' creado exitosamente.")
    except Exception as e:
        print(f"Error al escribir el archivo: {e}")
        
else:
    tmp = f"Error en consulta de ayer: {response_ayer.status_code}"
    print(tmp)
    
    # Escribir el mensaje de error en el archivo también
    try:
        with open("tipos_de_cambio.txt", "w", encoding="utf-8") as archivo:
            archivo.write(tmp)
        print("Archivo 'tipos_de_cambio.txt' creado con mensaje de error.")
    except Exception as e:
        print(f"Error al escribir el archivo: {e}")