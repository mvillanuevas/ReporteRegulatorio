On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathRef = objArgs(0)
WorkbookPathOH = objArgs(1)
periodo = objArgs(2)
Folio = objArgs(3)

'WorkbookPathRef = "C:\ReporteRegulatorioRpa\Input\Refacturacion_regular_v2.xlsm"
'WorkbookPathOH = "C:\ReporteRegulatorioRpa\Input\REXM - Overhead facturable julio 2025.xlsx"
'periodo = "Julio 2025"
'Folio = "618"

Folio = CInt(Folio)
'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

Set objWorkbookPathRef = objExcel.Workbooks.Open(WorkbookPathRef, 0)
Set objWorkbookSheetRef= objWorkbookPathRef.Worksheets("OH")
Set objWorkbookSheetCat= objWorkbookPathRef.Worksheets("Catalogo")

Set objWorkbookPathOH = objExcel.Workbooks.Open(WorkbookPathOH, 0)
Set objWorkbookSheetOH = objWorkbookPathOH.Worksheets("Sheet2")

Const xlPart = 2
Const xlValues = -4163

' Establece rango de busqueda en la columna C
Set nRange = objWorkbookSheetOH.Range("A:A")
Set cRange = objWorkbookSheetCat.Range("A:A")
' En la columna A de la hoja OH, buscar el valor "BL29" y obtener la fila
Dim foundCell : Set foundCell = nRange.Find("BL29",,xlValues,xlPart)
Dim CatCell 
fechaActual = DateSerial(Year(Date), Month(Date), Day(Date))

' Si la longitud de la variable "Folio" es menor a 6, agregar ceros a la izquierda con la cantidad de dígitos necesarios
If Len(Folio) < 6 Then
    Folio = Right(String(6, "0") & Folio, 6)
End If


If Not foundCell Is Nothing Then
    rowNumber = foundCell.Row
    filaOH = 2
    'Iterar cada 3 columnas a partir de la columna F
    For i = 6 To objWorkbookSheetOH.Cells(rowNumber, objWorkbookSheetOH.Columns.Count).End(-4159).Column Step 3
        ' Verificar si la celda contiene un valor numérico
        If objWorkbookSheetOH.Cells(rowNumber, i).Value <> 0 And  objWorkbookSheetOH.Cells(rowNumber, i).Value <> "-" Then
            objWorkbookSheetRef.Cells(filaOH, 3).Value = "'" & Folio
            ' Fecha actual
            'objWorkbookSheetRef.Cells(filaOH, 4).Value = Right("0" & Day(Date), 2) & "/" & Right("0" & Month(Date), 2) & "/" & Year(Date)
            objWorkbookSheetRef.Cells(filaOH, 4).Value = fechaActual
            ' Reemplazar el valor " <<Periodo>>" de la columna 28 por la variable "Periodo"
            objWorkbookSheetRef.Cells(filaOH, 28).Value = Replace(objWorkbookSheetRef.Cells(filaOH, 28).Value, "<<Periodo>>", periodo)
            proveedor = Trim(objWorkbookSheetOH.Cells(4, i).Value)
            Set CatCell = cRange.Find(proveedor,,xlValues,xlPart)
            If Not CatCell Is Nothing Then
                CatCellRow = CatCell.Row
                rfc = objWorkbookSheetCat.Cells(CatCellRow, 2).Value
                razonSocial = objWorkbookSheetCat.Cells(CatCellRow, 3).Value
                cp = objWorkbookSheetCat.Cells(CatCellRow, 4).Value
                ' Si cp tiene menos de 4 digitos, agregar ceros a la izquierda hasta tener 4 digitos
                If Len(cp) < 4 Then
                    cp = "'" & Right(String(4, "0") & cp, 4)
                End If
                regimen = objWorkbookSheetCat.Cells(CatCellRow, 5).Value
            End If
            objWorkbookSheetRef.Cells(filaOH, 16).Value = rfc
            objWorkbookSheetRef.Cells(filaOH, 17).Value = razonSocial
            objWorkbookSheetRef.Cells(filaOH, 18).Value = cp
            objWorkbookSheetRef.Cells(filaOH, 21).Value = regimen
            objWorkbookSheetRef.Cells(filaOH, 29).Value = objWorkbookSheetOH.Cells(rowNumber, i).Value
            Folio = Folio + 1
            filaOH = filaOH + 1
            ' Si la longitud de la variable "Folio" es menor a 6, agregar ceros a la izquierda con la cantidad de dígitos necesarios
            If Len(Folio) < 6 Then
                Folio = Right(String(6, "0") & Folio, 6)
            End If
        End If
    Next
    ' Ultima fila
    lastRowOH = objWorkbookSheetRef.Cells(objWorkbookSheetRef.Rows.Count, 3).End(-4162).Row
    objWorkbookSheetRef.Range("A2:B2").AutoFill objWorkbookSheetRef.Range("A2:B" & lastRowOH), 1
    objWorkbookSheetRef.Range("E2:O2").AutoFill objWorkbookSheetRef.Range("E2:O" & lastRowOH), 1
    objWorkbookSheetRef.Range("S2:T2").AutoFill objWorkbookSheetRef.Range("S2:T" & lastRowOH), 1
    objWorkbookSheetRef.Range("V2:AB2").AutoFill objWorkbookSheetRef.Range("V2:AB" & lastRowOH), 1
    objWorkbookSheetRef.Range("AD2:GC2").AutoFill objWorkbookSheetRef.Range("AD2:GC" & lastRowOH), 1

    objWorkbookSheetRef.Range("D2:D" & lastRowOH).NumberFormat = "dd/mm/yyyy"
    objWorkbookSheetRef.Range("AC2:AC" & lastRowOH).NumberFormat = "#,##0.00"

    ' Guardar y cerrar el libro de refacturación
    objWorkbookPathRef.Save
    objWorkbookPathRef.Close

    objWorkbookPathOH.Save
    objWorkbookPathOH.Close
    ' Cerrar la aplicación de Excel
    objExcel.Quit

Else
    ' Guardar y cerrar el libro de refacturación
    objWorkbookPathRef.Save
    objWorkbookPathRef.Close

    objWorkbookPathOH.Save
    objWorkbookPathOH.Close
    ' Cerrar la aplicación de Excel
    objExcel.Quit
    Msg = "Error was generated by " & "Valor 'BL29' no encontrado en la columna A."
    WScript.StdOut.WriteLine Msg
End If

'Devuelve el error en caso de
If Err.Number <> 0 Then
    ' Guardar y cerrar el libro de refacturación
    objWorkbookPathRef.Save
    objWorkbookPathRef.Close

    objWorkbookPathOH.Save
    objWorkbookPathOH.Close
    ' Cerrar la aplicación de Excel
    objExcel.Quit
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.StdOut.WriteLine Msg
Else
    WScript.StdOut.WriteLine Folio & ":Script executed successfully."
End if
