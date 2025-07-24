On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathRexmex = objArgs(0)
WorkbookPathRef = objArgs(1)
ActualMonth = CInt(objArgs(2))
TipoDocto = objArgs(3)
anio = CInt(objArgs(4))

'WorkbookPathRexmex = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Refacturacion_Test\REXMEX - Cuenta Operativa 2025_080725.xlsx"
'WorkbookPathRef = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Refacturacion_Test\Layout refacturación.xlsx"
'ActualMonth = 6
'TipoDocto = "<>NC"
'anio = 2025

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

Set objWorkbookPathRexmex = objExcel.Workbooks.Open(WorkbookPathRexmex, 0)
Set objWorkbookSheetRexmex = objWorkbookPathRexmex.Worksheets(WorkbookSheetRexmex)


' Verificar si la hoja WorkbookSheetLayout existe, si existe eliminarla y duplicar la hoja Template
If SheetExists(objWorkbookPathRef, WorkbookSheetLayout) Then
    objWorkbookPathRef.Worksheets(WorkbookSheetLayout).Delete
End If
' Duplicar la hoja Template y renombrarla a WorkbookSheetLayout
If SheetExists(objWorkbookPathRef, "Template") Then
    objWorkbookPathRef.Worksheets("Template").Copy objWorkbookPathRef.Worksheets(objWorkbookPathRef.Worksheets.Count)
    objWorkbookPathRef.Worksheets(objWorkbookPathRef.Worksheets.Count - 1).Name = WorkbookSheetLayout
    ' Unhide hoja WorkbookSheetLayout
    objWorkbookPathRef.Worksheets(WorkbookSheetLayout).Visible = -1 ' -1 = xlSheetVisible
End If

objWorkbookSheetRexmex.ShowAllData

flag = ReemplazarGuionesSimilares(objWorkbookSheetRexmex, 56)

' Referencia a la hoja de Layout refacturación
Set objWorkbookSheetRefL = objWorkbookPathRef.Worksheets(WorkbookSheetLayout)

' Arreglo de hojas de refacturación
Dim refacturacionSheets, bloque
refacturacionSheets = Array("BL29", "BL10", "BL11", "BL14")

' Iteraer sobre las hojas de refacturación
Dim i

For i = LBound(refacturacionSheets) To UBound(refacturacionSheets)
    Dim sheetName
    sheetName = refacturacionSheets(i)
    
    ' Verificar si la hoja existe en el libro de refacturación
    If SheetExists(objWorkbookPathRef, sheetName) Then
        Set objWorkbookSheetRef = objWorkbookPathRef.Worksheets(sheetName)

        ' Verificar si los filtros están activos en la fila 1, si no, activarlos
        If objWorkbookSheetRexmex.AutoFilterMode Then
            objWorkbookSheetRexmex.AutoFilterMode = False
        End If

        If Not objWorkbookSheetRexmex.AutoFilterMode Then
            objWorkbookSheetRexmex.Rows(1).AutoFilter
        End If

        ActualMonth = CInt(ActualMonth)
        Dim ultimoDiaMes
        ultimoDiaMes = DateSerial(anio, ActualMonth + 1, 0)
        ultimoDiaMes =  Right("0" & Day(ultimoDiaMes),2) & "-" & Right("0" & Month(ultimoDiaMes),2) & "-" & Year(ultimoDiaMes)

        primerDiaMes = DateSerial(anio, ActualMonth, 1)
        primerDiaMes = Right("0" & Day(primerDiaMes),2) & "-" & Right("0" & Month(primerDiaMes),2) & "-" & Year(primerDiaMes)


        ' Encontrar la última fila con datos en la columna a filtrar 
        lastRow = objWorkbookSheetRexmex.Cells(objWorkbookSheetRexmex.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

        objWorkbookSheetRexmex.Range("AA2:AA" & lastRow).NumberFormat = "dd/mm/yyyy"

        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 27), objWorkbookSheetRexmex.Cells(lastRow, 27)).AutoFilter _
                                    27, ">=" & CDbl(CDate(primerDiaMes)), 1, "<=" & CDbl(CDate(ultimoDiaMes))

        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 34), objWorkbookSheetRexmex.Cells(lastRow, 34)).AutoFilter _
                                    34, "<>OVERHEAD"

        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 55), objWorkbookSheetRexmex.Cells(lastRow, 55)).AutoFilter _
                                    55, TipoDocto

        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 24), objWorkbookSheetRexmex.Cells(lastRow, 24)).AutoFilter _
                                    24, "=" & sheetName
        
        Set dRange = objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 24), objWorkbookSheetRexmex.Cells(lastRow, 71)).SpecialCells(12)

        If Not dRange Is Nothing Then
            ' Encontrar la última fila con datos en la hoja de Layout refacturación
            lastRowR = objWorkbookSheetRef.Cells(objWorkbookSheetRef.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

            ' Copiar los valores de las celdas visibles a la hoja de REXMEX
            dRange.Copy
            ' Mostrar todas las filas de la hoja (quitar ocultamiento)
            objWorkbookSheetRef.Rows.Hidden = False
            ' Mostrar todas las columnas de la hoja (quitar ocultamiento)
            objWorkbookSheetRef.Columns.Hidden = False

            objWorkbookSheetRef.Range("A" & lastRowR + 1).PasteSpecial -4163 ' -4163 = xlPasteAll
            ' Quitar el modo de corte/copia
            objExcel.CutCopyMode = False
        Else
            Err.Clear
            ' No se encontraron celdas visibles, continuar con la siguiente hoja
        End If
    End If
    Set dRange = Nothing
Next
On Error GoTo 0
' Encontrar la última fila con datos en la hoja de Layout refacturación
lastRowL = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

' Ocultar las filas que no cumplen con el criterio de la columna 1 (BL29)
Dim j
For j = 7 To lastRowL
    If objWorkbookSheetRefL.Cells(j, 1).Value <> "BLOQUE 29" Then
        objWorkbookSheetRefL.Rows(j).Hidden = True
    End If
Next

Set objWorkbookSheetRef = objWorkbookPathRef.Worksheets("BL29")
' Encontrar la última fila con datos en la columna a filtrar 
lastRow = objWorkbookSheetRef.Cells(objWorkbookSheetRef.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

'Aplica Text to columns en formate General
objWorkbookSheetRef.Range("AG:AG").TextToColumns
' Ordenar de manera ascendente la columna AG (columna 33) "UUID"
With objWorkbookSheetRef.Sort
    .SortFields.Clear
    .SortFields.Add objWorkbookSheetRef.Range("AG2:AG" & lastRow), 0, 1 ' 0 = xlSortOnValues, 1 = xlAscending
    .SetRange objWorkbookSheetRef.Range("A1:AV" & lastRow) ' Ajusta el rango según tus datos
    .Header = 1 ' 1 = xlYes (hay encabezado)
    .Apply
End With

' Cerrar el libro de REXMEX sin guardar cambios
objWorkbookPathRexmex.Save
objWorkbookPathRexmex.Close

' Guardar y cerrar el libro de refacturación
objWorkbookPathRef.Save
objWorkbookPathRef.Close
' Cerrar la aplicación de Excel
objExcel.Quit

'Devuelve el error en caso de
If Err.Number <> 0 Then
    ' Cerrar el libro de REXMEX sin guardar cambios
    objWorkbookPathRexmex.Save
    objWorkbookPathRexmex.Close

    ' Guardar y cerrar el libro de refacturación
    objWorkbookPathRef.Save
    objWorkbookPathRef.Close
    ' Cerrar la aplicación de Excel
    objExcel.Quit
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.StdOut.WriteLine Msg
Else
    WScript.StdOut.WriteLine "Script executed successfully."
    WScript.Quit 0
End if

'____________________________________________________________________________________________________________________________________________
' Función para validar si una hoja existe en un libro de Excel
Function SheetExists(wb, sheetName)
    Dim s
    SheetExists = False
    For Each s In wb.Sheets
        If StrComp(s.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

Function ReemplazarGuionesSimilares(objHoja, columnaObjetivo)

    ' Obtener última fila usada en la columna deseada
    lastRow = objHoja.Cells(objHoja.Rows.Count, columnaObjetivo).End(-4162).Row ' -4162 = xlUp

    ' Caracteres similares a guión
    guionesSimilares = Array("–", "—", "?", "?", ChrW(8208), ChrW(8211), ChrW(8212), ChrW(8722))

    For i = 2 To lastRow
        celda = objHoja.Cells(i, columnaObjetivo).Value

        If Not IsEmpty(celda) Then
            For Each g In guionesSimilares
                If InStr(celda, g) > 0 Then
                    celda = Replace(celda, g, "-") ' Reemplaza el guión similar por un guión normal
                End If
            Next
            objHoja.Cells(i, columnaObjetivo).Value = celda
        End If
    Next
    ReemplazarGuionesSimilares = True
End Function