# -*- coding: utf-8 -*-
import requests
from datetime import datetime, timedelta
import os

def get_tdc_banxico():
    # !Obtiene los tipos de cambio del Banco de México para EUR/USD, EUR/MXN, GBP/USD y GBP/MXN.
    # Token generado desde Banxico
    token = "36662b701117b841f269b53fda9a936a029bafee8aa375a149b41a67617ccba5"

    # SF57922 = 	Tipos de Cambio para Revalorización de Balance del Banco de México, EUR U.Mon.Europea (EUR/Euro 4/), Dólares por divisa
    # SF57923 = 	Tipos de Cambio para Revalorización de Balance del Banco de México, EUR U.Mon.Europea (EUR/Euro 4/), Tipo en Pesos
    # SF57814 = 	GBP Gran Bretaña (Libra esterlina), Dólares por Divisa
    # SF57815 = 	GBP Gran Bretaña (Libra esterlina), Tipo en Pesos

    # ID de series a consultar
    series_ids = ["SF57922", "SF57923", "SF57814", "SF57815"]
    series_str = ",".join(series_ids)

    # Endpoint REST
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series_str}/datos/oportuno"

    # Encabezados con token de autenticación
    headers = {
        "Bmx-Token": token
    }

    # Hacer la solicitud
    response = requests.get(url, headers=headers)

    # Validar respuesta
    if response.status_code == 200:
        data = response.json()
        series = data.get("bmx", {}).get("series", [])
        tc = ""  # Initialize tc variable
        for serie in series:
            # Si serie_id es SF57922 concatenar con "EUR/USD" pero si es SF57923 concatenar con "EUR/MXN" 
            # pero si es SF57814 concatenar con "GBP/USD" y si es SF57815 concatenar con "GBP/MXN"
            if serie["idSerie"] == "SF57922":
                serie["idSerie"] += " EUR/USD"
            elif serie["idSerie"] == "SF57923":
                serie["idSerie"] += " EUR/MXN"
            elif serie["idSerie"] == "SF57814":
                serie["idSerie"] += " GBP/USD"
            elif serie["idSerie"] == "SF57815":
                serie["idSerie"] += " GBP/MXN"
            # Extraer información de la serie
            serie_id = serie["idSerie"]
            nombre = serie["titulo"]
            valor = serie["datos"][0]["dato"]
            fecha = serie["datos"][0]["fecha"]
            #Concatenear serie_id, nombre, valor y fecha para que se una sola variable
            if tc:  # If tc already has content, add separator
                tc = f"{serie_id}:{valor}:{fecha}" + ";" + tc
            else:  # First iteration
                tc = f"{serie_id}:{valor}:{fecha}"
        return tc
    else:
        return f"Error al consultar Banxico: {response.status_code}"
    

def get_tdc_banxico_usd():
    # !Obtiene los tipos de cambio del Banco de México para "Para solventar obligaciones" y 
    # !"Determinación" para el día actual y el día anterior.
    # Token de Banxico
    token = "36662b701117b841f269b53fda9a936a029bafee8aa375a149b41a67617ccba5"

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
        return f"Error en consulta actual: {response_actual.status_code}"

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
        return f"{tc};{tdc}"
    else:
        return f"Error en consulta de ayer: {response_ayer.status_code}"
    
# Ejecutar ambas funciones y guardar los resultados en un archvivo txt con salto de línea
if __name__ == "__main__":
    tdc_banxico = get_tdc_banxico()
    
    # Obtener la ruta de este script
    ruta_actual = os.path.dirname(os.path.abspath(__file__))
    
    # Escribir los resultados en un archivo de texto
    try:
        with open(os.path.join(ruta_actual, "tipos_de_cambio_usd.txt"), "w", encoding="utf-8") as archivo:
            archivo.write(tdc_banxico)
        print("Archivo 'tipos_de_cambio_usd.txt' creado exitosamente.")
    except Exception as e:
        print(f"Error al crear el archivo: {e}")