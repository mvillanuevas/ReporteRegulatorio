On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathRexmex = objArgs(0)
WorkbookPathRef = objArgs(1)
ActualMonth = objArgs(2)
FechaRefacturacion = objArgs(3)

'WorkbookPathRexmex = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\REXMEX - Cuenta Operativa 2025_090625.xlsx"
'WorkbookPathRef = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Refacturacion_Test\Layout refacturación.xlsx"
'ActualMonth = 3
'FechaRefacturacion = "Marzo 2025"

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

'Abre libro Excel
Set objWorkbookPathRef = objExcel.Workbooks.Open(WorkbookPathRef, 0)
Set objWorkbookSheetRefL = objWorkbookPathRef.Worksheets(WorkbookSheetLayout)
Set objWorkbookSheetRef = objWorkbookPathRef.Worksheets("BL29")

' Arreglo de hojas de refacturación
Dim proveedores
proveedores = Array("PC CARIGALI", "PTTEP", "REPSOL", "SIERRA NEVADA")

saveLastRow = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row + 1

For i = LBound(proveedores) To UBound(proveedores)
    ' Copiar columnas específicas de objWorkbookSheetRef a objWorkbookSheetRefL

    Dim copyLastRow, pasteLastRow
    copyLastRow = objWorkbookSheetRef.Cells(objWorkbookSheetRef.Rows.Count, 1).End(-4162).Row
    If copyLastRow >= 2 Then
        ' Si la ultima fila con datos es 6, entonces pasteLastRow es 7
        If objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row = 6 Then
            pasteLastRow = 7
        Else
            pasteLastRow = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row + 2
        End If
        
        ' AP (col 42) -> D (col 4)
        objWorkbookSheetRef.Range("AP2:AP" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("D" & pasteLastRow).PasteSpecial -4163

        ' AG (col 33) -> E (col 5)
        objWorkbookSheetRef.Range("AG2:AG" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("E" & pasteLastRow).PasteSpecial -4163
            
        'Iterar la columna E y validar si la longitud del valor del la celda es menor a 16 y si es asi, cortar los valores hacia la columna F
        ' Esto para identificar un UUID extragero
        Dim longcell
        For Each cell In objWorkbookSheetRefL.Range("E" & pasteLastRow & ":E" & pasteLastRow + copyLastRow - 2)
            ' Si el valor de la celda contiene el valor "pep" restar 3 a la longitud del valor de la celda
            If InStr(1, cell.Value, "*pep", vbTextCompare) > 0 Then
                longcell = Len(cell.value) - 3
            Else
                longcell = Len(cell.Value)
            End If
            If longcell < 16 Then
                cell.Offset(0, 1).Value = cell.Value ' Mover el valor a la columna F
                cell.Value = "" ' Limpiar la celda original
            End If
        Next

        ' B (col 2) -> L (col 12)
        RowCount = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row
        
        objWorkbookSheetRefL.Range("E" & pasteLastRow & ":E" & RowCount).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("L" & pasteLastRow).PasteSpecial -4163

        ' AI (col 35) -> I (col 9)
        'objWorkbookSheetRef.Range("AI2:AI" & copyLastRow).SpecialCells(12).Copy
        'objWorkbookSheetRefL.Range("I" & pasteLastRow).PasteSpecial -4163

        ' AH (col 34) -> M (col 13)
        ' Formato fecha columna AH
        objWorkbookSheetRef.Range("AH2:AH" & copyLastRow).NumberFormat = "dd-mm-yyyy"
        objWorkbookSheetRef.Range("AH2:AH" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("M" & pasteLastRow).PasteSpecial -4163
        
        ' N (col 14)
        objWorkbookSheetRefL.Range("N" & pasteLastRow).Value = "'" & FechaRefacturacion
        ' AutoFill N (col 14) to the last row
        objWorkbookSheetRefL.Range("N" & pasteLastRow).AutoFill objWorkbookSheetRefL.Range("N" & pasteLastRow & ":N" & RowCount), 1

        ' N (col 14) -> O (col 15)
        objWorkbookSheetRef.Range("N2:N" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("O" & pasteLastRow).PasteSpecial -4163

        ' AE (col 31) -> R (col 18)
        objWorkbookSheetRef.Range("AE2:AE" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("R" & pasteLastRow).PasteSpecial -4163

        ' L (col 12) -> V (col 22)
        objWorkbookSheetRef.Range("L2:L" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("V" & pasteLastRow).PasteSpecial -4163

        ' F (col 6) -> X (col 24)
        objWorkbookSheetRef.Range("F2:F" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("X" & pasteLastRow).PasteSpecial -4163

        ' G (col 7) -> Y (col 25)
        objWorkbookSheetRef.Range("G2:G" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("Y" & pasteLastRow).PasteSpecial -4163

        ' H (col 8) -> Z (col 26)
        objWorkbookSheetRef.Range("H2:H" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("Z" & pasteLastRow).PasteSpecial -4163

        ' I (col 9) -> AA (col 27)
        objWorkbookSheetRef.Range("I2:I" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("AA" & pasteLastRow).PasteSpecial -4163

        ' C (col 3) -> AG (col 33)
        objWorkbookSheetRef.Range("C2:C" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("AG" & pasteLastRow).PasteSpecial -4163

        ' D (col 4) -> AH (col 34)
        objWorkbookSheetRef.Range("D2:D" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("AH" & pasteLastRow).PasteSpecial -4163

        ' E (col 5) -> AI (col 35)
        objWorkbookSheetRef.Range("E2:E" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("AI" & pasteLastRow).PasteSpecial -4163

        ' K (col 11) -> AJ (col 36)
        objWorkbookSheetRef.Range("K2:K" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("AJ" & pasteLastRow).PasteSpecial -4163

        ' AO (col 41) -> AK (col 37)
        objWorkbookSheetRef.Range("AO2:AO" & copyLastRow).SpecialCells(12).Copy
        objWorkbookSheetRefL.Range("AK" & pasteLastRow).PasteSpecial -4163

        objExcel.CutCopyMode = False

        ' Realizar autofill de fórmulas en las columnas S, T, U, AD, AE de objWorkbookSheetRefL
        
        Dim fillLastRow
        fillLastRow = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row

        ' C (col 3)
        objWorkbookSheetRefL.Range("C7").FormulaR1C1 = "=IFNA(VLOOKUP(RC[-2],Contratos!C[-2]:C[-1],2,FALSE)," & Chr(34) & Chr(34) & ")"
        objWorkbookSheetRefL.Range("C7").AutoFill objWorkbookSheetRefL.Range("C7:C" & fillLastRow)
        ' S (col 19)
        objWorkbookSheetRefL.Range("S7").AutoFill objWorkbookSheetRefL.Range("S7:S" & fillLastRow)
        ' T (col 20)
        objWorkbookSheetRefL.Range("T7").AutoFill objWorkbookSheetRefL.Range("T7:T" & fillLastRow)
        ' U (col 21)
        objWorkbookSheetRefL.Range("U7").FormulaR1C1 = "=IF(RC[-19]=" & Chr(34) & "PC CARIGALI" & Chr(34) & ",SUMIF(BL29!C[12],Layout!RC[-16],BL29!C[-1])+SUMIF(BL29!C[12],Layout!RC[-15],BL29!C[-1]),IF(RC[-19]=" & Chr(34) & "PTTEP" & Chr(34) & ",SUMIF(BL29!C[12],Layout!RC[-16],BL29!C[3])+SUMIF(BL29!C[12],Layout!RC[-15],BL29!C[3]),IF(RC[-19]=" & Chr(34) & "SIERRA NEVADA" & Chr(34) & ",SUMIF(BL29!C[12],Layout!RC[-16],BL29!C[5])+SUMIF(BL29!C[12],Layout!RC[-15],BL29!C[5]),SUMIF(BL29!C[12],Layout!RC[-16],BL29!C[-5])+SUMIF(BL29!C[12],Layout!RC[-15],BL29!C[-5]))))"
        objWorkbookSheetRefL.Range("U7").AutoFill objWorkbookSheetRefL.Range("U7:U" & fillLastRow)
        ' Q (col 17)
        objWorkbookSheetRefL.Range("Q7").FormulaR1C1 = "=IF(RC[-15]=" & Chr(34) & "PC CARIGALI" & Chr(34) & ",SUMIF(BL29!C[16],Layout!RC[-12],BL29!C[2])+SUMIF(BL29!C[16],Layout!RC[-11],BL29!C[2]),IF(RC[-15]=" & Chr(34) & "PTTEP" & Chr(34) & ",SUMIF(BL29!C[16],Layout!RC[-12],BL29!C[6])+SUMIF(BL29!C[16],Layout!RC[-11],BL29!C[6]),IF(RC[-15]=" & Chr(34) & "SIERRA NEVADA" & Chr(34) & ",SUMIF(BL29!C[16],Layout!RC[-12],BL29!C[8])+SUMIF(BL29!C[16],Layout!RC[-11],BL29!C[8]),SUMIF(BL29!C[16],Layout!RC[-12],BL29!C[-2])+SUMIF(BL29!C[16],Layout!RC[-11],BL29!C[-2]))))"
        objWorkbookSheetRefL.Range("Q7").AutoFill objWorkbookSheetRefL.Range("Q7:Q" & fillLastRow)
        ' W (col 23s)
        objWorkbookSheetRefL.Range("W7").AutoFill objWorkbookSheetRefL.Range("W7:W" & fillLastRow)
        ' AD (col 30)
        objWorkbookSheetRefL.Range("AD7").AutoFill objWorkbookSheetRefL.Range("AD7:AD" & fillLastRow)
        ' AE (col 31)
        objWorkbookSheetRefL.Range("AE7").AutoFill objWorkbookSheetRefL.Range("AE7:AE" & fillLastRow)

        ' Limpiar una fila vacía antes de pegar los datos
        If pasteLastRow > 7 Then
            objWorkbookSheetRefL.Rows(pasteLastRow - 1).ClearContents
        End If

        ' Rellenar con autofill el valor REP en la columna B de objWorkbookSheetRefL
        Dim bStart, bEnd
        bStart = pasteLastRow
        bEnd = fillLastRow

        ' Rellenar con autofill el valor "BLOQUE 29" en la columna A de objWorkbookSheetRefL
        objWorkbookSheetRefL.Range("A" & bStart).Value = "BLOQUE 29"
        objWorkbookSheetRefL.Range("A" & bStart & ":A" & bEnd).Value = "BLOQUE 29"
        ' Rellenar con autofill el valor del proveedor actual en la columna B de objWorkbookSheetRefL
        objWorkbookSheetRefL.Range("B" & bStart).Value = proveedores(i)
        objWorkbookSheetRefL.Range("B" & bStart & ":B" & bEnd).Value = proveedores(i)

        ' Rellenar con autofill el valor "Bloque 29, AP-CS-G10, Cuenca Salina / Administración General" en la columna AC de objWorkbookSheetRefL
        objWorkbookSheetRefL.Range("AC" & bStart).Value = "Bloque 29, AP-CS-G10, Cuenca Salina / Administración General"
        objWorkbookSheetRefL.Range("AC" & bStart & ":AC" & bEnd).Value = "Bloque 29, AP-CS-G10, Cuenca Salina / Administración General"
        
    End If
Next
On Error GoTo 0

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
    WScript.StdOut.WriteLine CStr(saveLastRow)
End if
