Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

' Abrir el archivo
Set wb = objExcel.Workbooks.Open("C:\ruta\archivo.xlsx")
Set ws = wb.Sheets(1) ' Ajusta la hoja

' Determinar última fila (buscando por UUID en col AG)
lastRow = ws.Cells(ws.Rows.Count, 33).End(-4162).Row ' 33 = col AG

' Diccionario para almacenar si grupo tiene negativo
Set dictNeg = CreateObject("Scripting.Dictionary")
' Diccionario para almacenar la primera fila de cada UUID
Set dictPrimeraFila = CreateObject("Scripting.Dictionary")

' Paso 1: Detectar grupos con negativos y registrar primera fila
For i = 2 To lastRow ' Asumiendo encabezados en fila 1
    uuid = Trim(CStr(ws.Cells(i, 33).Value)) ' Columna AG
    valorAA = ws.Cells(i, 27).Value          ' Columna AA
    
    If Not dictNeg.Exists(uuid) Then
        dictNeg(uuid) = False
        dictPrimeraFila(uuid) = i ' Guardar primera fila donde aparece
    End If
    
    ' Marcar si el grupo tiene al menos un negativo
    If IsNumeric(valorAA) Then
        If valorAA < 0 Then
            dictNeg(uuid) = True
        End If
    End If
Next

' Paso 2: Asignar valores en col 49
For Each uuid In dictNeg.Keys
    If dictNeg(uuid) = True Then
        primeraFila = dictPrimeraFila(uuid)
        ' Poner 0 en la primera fila del grupo
        ws.Cells(primeraFila, 49).Value = 0
        
        ' Poner 1 en el resto de las filas del grupo
        For i = primeraFila + 1 To lastRow
            If Trim(CStr(ws.Cells(i, 33).Value)) = uuid Then
                ws.Cells(i, 49).Value = 1
            End If
        Next
    End If
Next

' Guardar y cerrar
wb.Save
wb.Close False
objExcel.Quit

MsgBox "Proceso completado"
