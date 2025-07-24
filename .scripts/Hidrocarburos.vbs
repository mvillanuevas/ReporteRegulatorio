On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathRef = objArgs(0)

'WorkbookPathRef = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\Timbrado\Refacturacion_regular_v2.xlsm"

WorkbookSheetRexmex = "Cuenta Operativa"
WorkbookSheetLayout = "Layout"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

WorkbookPathLayout = "Layout40"
GH = "GastoHidrocarburos"

Set objWorkbookPathRef = objExcel.Workbooks.Open(WorkbookPathRef, 0)
Set objWorkbookSheetLayout = objWorkbookPathRef.Worksheets(WorkbookPathLayout)
Set objWorkbookSheetGH = objWorkbookPathRef.Worksheets(GH)

' Verificar si los filtros están activos en la fila 1, si no, activarlos
If objWorkbookSheetLayout.AutoFilterMode Then
    objWorkbookSheetLayout.AutoFilterMode = False
End If

If Not objWorkbookSheetLayout.AutoFilterMode Then
    objWorkbookSheetLayout.Rows(1).AutoFilter
End If

lastRow = objWorkbookSheetLayout.Cells(objWorkbookSheetLayout.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

objWorkbookSheetLayout.Range(objWorkbookSheetLayout.Cells(1, 32), objWorkbookSheetLayout.Cells(lastRow, 32)).AutoFilter _
                            32, "=02"
     
Set dRange = objWorkbookSheetLayout.Range(objWorkbookSheetLayout.Cells(2, 1), objWorkbookSheetLayout.Cells(lastRow, 67)).SpecialCells(12)

dRange.Copy
objWorkbookSheetGH.Range("A2").PasteSpecial -4163 ' -4163 = xlPasteValues
' Desactivar el modo de copia
objExcel.CutCopyMode = False

' Iterar la columna 28 y reemplazar "Documento" por "IVA"
For Each cell In objWorkbookSheetGH.Range("AB2:AB" & objWorkbookSheetGH.Cells(objWorkbookSheetGH.Rows.Count, 28).End(-4162).Row)
    If InStr(cell.Value, "Documento") > 0 Then
        cell.Value = Replace(cell.Value, "Documento", "IVA")
    End If
Next
' Iterar la columna 29 y 30 y multiplicar por 0.16
For Each cell In objWorkbookSheetGH.Range("AC2:AD" & objWorkbookSheetGH.Cells(objWorkbookSheetGH.Rows.Count, 29).End(-4162).Row)
    If IsNumeric(cell.Value) Then
        cell.Value = CDbl(cell.Value) * 0.16
    End If
Next
' Establecer "01" en la columna 32
objWorkbookSheetGH.Range("AF2:AF" & objWorkbookSheetGH.Cells(objWorkbookSheetGH.Rows.Count, 1).End(-4162).Row).Value = "'01"
' Aplicar formato de color a las celdas copiadas
objWorkbookSheetGH.Range("A2:BO" & objWorkbookSheetGH.Cells(objWorkbookSheetGH.Rows.Count, 1).End(-4162).Row).Interior.Color = RGB(255, 230, 153)
' Obtener ultima fila con datos en la columna A
lastRowGH = objWorkbookSheetGH.Cells(objWorkbookSheetGH.Rows.Count, 1).End(-4162).Row + 1

dRange.Copy
objWorkbookSheetGH.Range("A" & lastRowGH).PasteSpecial -4163 ' -4163 = xlPasteValues
' Desactivar el modo de copia
objExcel.CutCopyMode = False

lastRowGH = objWorkbookSheetGH.Cells(objWorkbookSheetGH.Rows.Count, 1).End(-4162).Row
' Ordenar por la columna 3 (Columna C) la tabla de la coolumna A hasta la columna BO
With objWorkbookSheetGH.Sort
    .SortFields.Clear
    .SortFields.Add objWorkbookSheetGH.Range("C2"), 0, 1  ' 0 = xlSortOnValues, 1 = xlAscending
    .SetRange objWorkbookSheetGH.Range("A1", objWorkbookSheetGH.Cells(lastRowGH, 67))
    .Header = 1  ' xlYes (Con encabezados)
    .MatchCase = False
    .Orientation = 1 ' xlTopToBottom
    .Apply
End With

' Iterar la columna 32 y si contiene 02 aplicar formateo de relleno de color
For Each cell In objWorkbookSheetGH.Range("AF2:AF" & lastRowGH)
    If cell.Value = "02" Then
        outRow = cell.Row
        ' Verde
        objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 1), objWorkbookSheetGH.Cells(outRow, 2)).Interior.Color = RGB(198, 224, 180)
        objWorkbookSheetGH.Cells(outRow, 5).Interior.Color = RGB(198, 224, 180)
        objWorkbookSheetGH.Cells(outRow, 7).Interior.Color = RGB(198, 224, 180)
        objWorkbookSheetGH.Cells(outRow, 10).Interior.Color = RGB(198, 224, 180)
        objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 13), objWorkbookSheetGH.Cells(outRow, 15)).Interior.Color = RGB(198, 224, 180)
        objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 22), objWorkbookSheetGH.Cells(outRow, 23)).Interior.Color = RGB(198, 224, 180)
        objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 25), objWorkbookSheetGH.Cells(outRow, 26)).Interior.Color = RGB(198, 224, 180)

            ' Rojas
        objWorkbookSheetGH.Cells(outRow, 5).Interior.Color = RGB(244, 176, 132)
        objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 8), objWorkbookSheetGH.Cells(outRow, 9)).Interior.Color = RGB(244, 176, 132)
        objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 16), objWorkbookSheetGH.Cells(outRow, 18)).Interior.Color = RGB(244, 176, 132)
        objWorkbookSheetGH.Cells(outRow, 21).Interior.Color = RGB(244, 176, 132)
        objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 28), objWorkbookSheetGH.Cells(outRow, 36)).Interior.Color = RGB(244, 176, 132)
        objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 43), objWorkbookSheetGH.Cells(outRow, 66)).Interior.Color = RGB(244, 176, 132)

            ' Azul
        objWorkbookSheetGH.Cells(outRow, 3).Interior.Color = RGB(155, 194, 230)
        ' Establecer "01" en la columna 32
        objWorkbookSheetGH.Cells(outRow, 32) = "'01"
    End If
Next

objWorkbookSheetLayout.Range(objWorkbookSheetLayout.Cells(1, 32), objWorkbookSheetLayout.Cells(lastRow, 32)).AutoFilter _
                            32, "=01"

lastRowGH = objWorkbookSheetGH.Cells(objWorkbookSheetGH.Rows.Count, 1).End(-4162).Row + 1

Set dRange = objWorkbookSheetLayout.Range(objWorkbookSheetLayout.Cells(2, 1), objWorkbookSheetLayout.Cells(lastRow, 67)).SpecialCells(12)

dRange.Copy
objWorkbookSheetGH.Range("A" & lastRowGH).PasteSpecial -4163 ' -4163 = xlPasteValues
' Desactivar el modo de copia
objExcel.CutCopyMode = False

lastRowGHH = objWorkbookSheetGH.Cells(objWorkbookSheetGH.Rows.Count, 1).End(-4162).Row

' Iterar la columna 32 y si contiene 02 aplicar formateo de relleno de color
For Each cell In objWorkbookSheetGH.Range("AF" & lastRowGH & ":AF" & lastRowGHH)

    outRow = cell.Row
        ' Verde
    objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 1), objWorkbookSheetGH.Cells(outRow, 2)).Interior.Color = RGB(198, 224, 180)
    objWorkbookSheetGH.Cells(outRow, 5).Interior.Color = RGB(198, 224, 180)
    objWorkbookSheetGH.Cells(outRow, 7).Interior.Color = RGB(198, 224, 180)
    objWorkbookSheetGH.Cells(outRow, 10).Interior.Color = RGB(198, 224, 180)
    objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 13), objWorkbookSheetGH.Cells(outRow, 15)).Interior.Color = RGB(198, 224, 180)
    objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 22), objWorkbookSheetGH.Cells(outRow, 23)).Interior.Color = RGB(198, 224, 180)
    objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 25), objWorkbookSheetGH.Cells(outRow, 26)).Interior.Color = RGB(198, 224, 180)

            ' Rojas
    objWorkbookSheetGH.Cells(outRow, 5).Interior.Color = RGB(244, 176, 132)
    objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 8), objWorkbookSheetGH.Cells(outRow, 9)).Interior.Color = RGB(244, 176, 132)
    objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 16), objWorkbookSheetGH.Cells(outRow, 18)).Interior.Color = RGB(244, 176, 132)
    objWorkbookSheetGH.Cells(outRow, 21).Interior.Color = RGB(244, 176, 132)
    objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 28), objWorkbookSheetGH.Cells(outRow, 36)).Interior.Color = RGB(244, 176, 132)
    objWorkbookSheetGH.Range(objWorkbookSheetGH.Cells(outRow, 43), objWorkbookSheetGH.Cells(outRow, 66)).Interior.Color = RGB(244, 176, 132)

            ' Azul
    objWorkbookSheetGH.Cells(outRow, 3).Interior.Color = RGB(155, 194, 230)
Next

' Formato fecha columna 4
objWorkbookSheetGH.Range("D2:D" & lastRowGHH).NumberFormat = "dd/mm/yyyy"

' Autofill de la columna BR
objWorkbookSheetGH.Range("BR2").AutoFill objWorkbookSheetGH.Range("BR2:BR" & lastRowGHH)
' AutoFill de la columna BS - BT
objWorkbookSheetGH.Range("BS2:BT2").AutoFill objWorkbookSheetGH.Range("BS2:BT" & lastRowGHH),1
' AutoFill de la columna BU - CJ
objWorkbookSheetGH.Range("BU2:CJ2").AutoFill objWorkbookSheetGH.Range("BU2:CJ" & lastRowGHH)
' AutoFill de la columna CK
objWorkbookSheetGH.Range("CK2").AutoFill objWorkbookSheetGH.Range("CK2:CK" & lastRowGHH), 1
' AutoFill de la columna CL - CR
objWorkbookSheetGH.Range("CL2:CR2").AutoFill objWorkbookSheetGH.Range("CL2:CR" & lastRowGHH)
' AutoFill de la columna CS
objWorkbookSheetGH.Range("CS2").AutoFill objWorkbookSheetGH.Range("CS2:CS" & lastRowGHH), 1
' AutoFill de la columna CR - EP
objWorkbookSheetGH.Range("CT2:EP2").AutoFill objWorkbookSheetGH.Range("CT2:EP" & lastRowGHH)
' AutoFill de la columna EQ
objWorkbookSheetGH.Range("EQ2").AutoFill objWorkbookSheetGH.Range("EQ2:EQ" & lastRowGHH), 1
' AutoFill de la columna ER - IA
objWorkbookSheetGH.Range("ER2:IA2").AutoFill objWorkbookSheetGH.Range("ER2:IA" & lastRowGHH)

' Ordenar de manera ascendente la columna C de la hoja objWorkbookSheetLayout
With objWorkbookSheetLayout.Sort
    .SortFields.Clear
    .SortFields.Add objWorkbookSheetLayout.Range("C2"), 0, 1  ' 0 = xlSortOnValues, 1 = xlAscending
    .SetRange objWorkbookSheetLayout.Range("A1", objWorkbookSheetLayout.Cells(lastRow, 67))
    .Header = 1  ' xlYes (Con encabezados)
    .MatchCase = False
    .Orientation = 1 ' xlTopToBottom
    .Apply
End With

With objWorkbookSheetGH.Sort
    .SortFields.Clear
    .SortFields.Add objWorkbookSheetGH.Range("C2"), 0, 1  ' 0 = xlSortOnValues, 1 = xlAscending
    .SetRange objWorkbookSheetGH.Range("A1", objWorkbookSheetGH.Cells(lastRow, 67))
    .Header = 1  ' xlYes (Con encabezados)
    .MatchCase = False
    .Orientation = 1 ' xlTopToBottom
    .Apply
End With

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