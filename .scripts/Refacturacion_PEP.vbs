On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathRexmex = objArgs(0)
WorkbookPathRef = objArgs(1)
ActualMonth = objArgs(2)
TipoDocto = objArgs(3)

'WorkbookPathRexmex = "C:\ReporteRegulatorioRpa\Input\REXMEX - Cuenta Operativa 072025.xlsx"
'WorkbookPathRef = "C:\ReporteRegulatorioRpa\Input\Layout refacturación.xlsx"
'ActualMonth = 7
'TipoDocto = "Factura"

WorkbookSheetRexmex = "Cuenta Operativa"
WorkbookSheetLayout = "Layout"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = True
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = True
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set objWorkbookPathRef = objExcel.Workbooks.Open(WorkbookPathRef, 0)
Set objWorkbookSheetRefL = objWorkbookPathRef.Worksheets(WorkbookSheetLayout)
Set objWorkbookSheetRef = objWorkbookPathRef.Worksheets("BL29")

' --- Optimización de rendimiento: lectura y escritura por lotes ---
Dim wbsDict, wbsDictH, pepCounterDict
Set wbsDict     = CreateObject("Scripting.Dictionary")
Set wbsDictH    = CreateObject("Scripting.Dictionary")
Set pepCounterDict = CreateObject("Scripting.Dictionary")
Set negative = CreateObject("Scripting.Dictionary")

Dim data, rowNum, totalRows
Dim wbsValue, agValue, sKey
Dim results()

lastRow = objWorkbookSheetRef.Cells(objWorkbookSheetRef.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

If lastRow > 1 Then
    objWorkbookSheetRef.Activate
    Set rng = objWorkbookSheetRef.Range("A1:AV" & lastRow)
    ' Limpiar campos de ordenamiento anteriores
    objWorkbookSheetRef.Sort.SortFields.Clear

    ' Agregar primer criterio: AG2:AG260 ascendente
    objWorkbookSheetRef.Sort.SortFields.Add rng.Columns(33), 0, 1, , 0   ' 0 = xlSortOnValues, 1 = xlAscending, último 0 = xlSortNormal

    ' Agregar segundo criterio: AA2:AA260 descendente
    objWorkbookSheetRef.Sort.SortFields.Add rng.Columns(27), 0, 2, , 0   ' 0 = xlSortOnValues, 2 = xlDescending, último 0 = xlSortNormal

    ' Aplicar ordenamiento
    With objWorkbookSheetRef.Sort
        .SetRange rng
        .Header = 1              ' 1 = xlYes
        .MatchCase = False
        .Orientation = 1         ' 1 = xlTopToBottom
        .SortMethod = 1          ' 1 = xlPinYin
        .Apply
    End With

    If TipoDocto = "Factura" Then
        ' Diccionario para almacenar si grupo tiene negativo
        Set dictNeg = CreateObject("Scripting.Dictionary")
        ' Diccionario para almacenar la primera fila de cada UUID
        Set dictPrimeraFila = CreateObject("Scripting.Dictionary")

        ' Paso 1: Detectar grupos con negativos y registrar primera fila
        For i = 2 To lastRow ' Asumiendo encabezados en fila 1
            uuid = Trim(CStr(objWorkbookSheetRef.Cells(i, 33).Value)) ' Columna AG
            valorAA = objWorkbookSheetRef.Cells(i, 27).Value          ' Columna AA
            
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
                objWorkbookSheetRef.Cells(primeraFila, 49).Value = 0
                
                ' Poner 1 en el resto de las filas del grupo
                For i = primeraFila + 1 To lastRow
                    If Trim(CStr(objWorkbookSheetRef.Cells(i, 33).Value)) = uuid Then
                        objWorkbookSheetRef.Cells(i, 49).Value = 1
                    End If
                Next
            End If
        Next
    End If
    ' Leer todas las filas a un array (más rápido que trabajar directo con Cells)
    data = objWorkbookSheetRef.Range("A2:AW" & lastRow).Value ' A = col 1, AG = col 33
    totalRows = UBound(data, 1)
    ReDim results(totalRows) ' Guardar filas que se deben ocultar

    For rowNum = 1 To totalRows ' Ya que empezamos en A2, este es índice 1
        wbsValue = Trim(data(rowNum, 3))     ' Columna C
        agValue  = Trim(data(rowNum, 33))    ' Columna AG
        aaValue = CDbl(Trim(data(rowNum, 27)))    ' Columna AA
        tmpValue = Trim(data(rowNum, 49))

        If wbsValue <> "" And agValue <> "" And tmpValue <> "1" And tmpValue <> "0" Then
            sKey = agValue & "|" & wbsValue
            ' Duplicado exacto UUID + WBS
            If wbsDictH.Exists(sKey) Then
                results(rowNum) = True ' Marcar para ocultar
            Else
                wbsDictH.Add sKey, 1
            End If

            ' Contador por UUID
            If Not wbsDict.Exists(agValue) Then
                wbsDict.Add agValue, 1
                pepCounterDict.Add agValue, 0
            ElseIf Not results(rowNum) And Not data(rowNum, 49) = "1" Then
                pepCounterDict(agValue) = pepCounterDict(agValue) + 1
                data(rowNum, 33) = agValue & "*pep" & pepCounterDict(agValue)
                'results(rowNum) = True ' Marcar para ocultar
            End If
        End If
        If tmpValue = "1" Then
                results(rowNum) = True ' Marcar para ocultar
        End If
    Next

    ' Escribir los datos modificados de vuelta
    objWorkbookSheetRef.Range("A2:AW" & lastRow).Value = data

    ' Ocultar filas en un solo paso
    For rowNum = 1 To totalRows
        If results(rowNum) = True Then
            objWorkbookSheetRef.Rows(rowNum + 1).Hidden = True ' +1 por offset a partir de fila 2
        End If
    Next

    If TipoDocto = "Factura" Then
        ' Itera sobre la columna AA y oculta las filas que tenga valores negativos
        Dim cell
        For Each cell In objWorkbookSheetRef.Range("AA2:AA" & lastRow)
            If Not IsEmpty(cell.Value) And Not IsNull(cell.Value) Then
                If cell.Value < 0 Then
                    If InStr(1, cell.offset(0, 6).Value, "*pep", vbTextCompare) > 0 Then
                        posicion = InStr(1, cell.offset(0, 6).Value, "*pep", vbTextCompare) - 1
                        cell.offset(0, 6).Value = Left(cell.offset(0, 6).Value, posicion)
                    End If
                    cell.EntireRow.Hidden = True
                End If
            End If
        Next

    End If
Else
    On Error GoTo 0
End If
' Guardar y cerrar el libro de refacturación
objWorkbookPathRef.Save
objWorkbookPathRef.Close
' Cerrar la aplicación de Excel
objExcel.Quit

'Devuelve el error en caso de
If Err.Number <> 0 Then
    ' Guardar y cerrar el libro de refacturación
    objWorkbookPathRef.Save
    objWorkbookPathRef.Close
    ' Cerrar la aplicación de Excel
    objExcel.Quit
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.StdOut.WriteLine Msg
Else
    WScript.StdOut.WriteLine "Script executed successfully."
End if
