On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathRexmex = objArgs(0)
WorkbookSheetRexmex = objArgs(1)
WorkbookOverhead = objArgs(2)
WorkbookSheetOverhead = objArgs(3)
anio = CInt(objArgs(4))
mes = CInt(objArgs(5))

'WorkbookPathRexmex = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\REXMEX - Cuenta Operativa 2025_090625.xlsx"
'WorkbookSheetRexmex = "Cuenta Operativa"
'WorkbookOverhead = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Refacturacion_Test\REXM - Overhead facturable Mayo 2025.xlsx"
'WorkbookSheetOverhead = "Sheet2"
'anio = 2025
'mes = 3

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set objWorkbookPathRexmex = objExcel.Workbooks.Open(WorkbookPathRexmex, 0)
Set objWorkbookSheetRexmex = objWorkbookPathRexmex.Worksheets(WorkbookSheetRexmex)

Set objWorkbookPathOverhead = objExcel.Workbooks.Open(WorkbookOverhead, 0)
Set objWorkbookSheetOverhead = objWorkbookPathOverhead.Worksheets(WorkbookSheetOverhead)

' Verificar si los filtros están activos en la fila 1, si no, activarlos
If objWorkbookSheetRexmex.AutoFilterMode Then
    objWorkbookSheetRexmex.AutoFilterMode = False
End If

If Not objWorkbookSheetRexmex.AutoFilterMode Then
    objWorkbookSheetRexmex.Rows(1).AutoFilter
End If

' Ultimo día del mes actual
Dim ultimoDiaMes
ultimoDiaMes = DateSerial(anio, mes + 1, 0)
ultimoDiaMes =  Right("0" & Day(ultimoDiaMes),2) & "-" & Right("0" & Month(ultimoDiaMes),2) & "-" & Year(ultimoDiaMes)

' Filtrar la columna AA de la hoja objWorkbookSheetRexmex por fecha
Dim lastRowRexmex, lastRowOverhead

lastRowRexmex = objWorkbookSheetRexmex.Cells(objWorkbookSheetRexmex.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp
lastRowOverhead = objWorkbookSheetOverhead.Cells(objWorkbookSheetOverhead.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

' Filtrar la columna 27 (AA) de la hoja objWorkbookSheetOverhead por fecha y tipo de documento
objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 24), objWorkbookSheetRexmex.Cells(lastRowRexmex, 24)).AutoFilter _
                                24, "=BL29"
objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 27), objWorkbookSheetRexmex.Cells(lastRowRexmex, 27)).AutoFilter _
                                27, "=" & CDate(ultimoDiaMes)
objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 34), objWorkbookSheetRexmex.Cells(lastRowRexmex, 34)).AutoFilter _
                                34, "=*OVERHEAD*"

' Tomar el valor filtrado de la columna 36 ' (AJ) de la hoja objWorkbookSheetRexmex
Dim overheadValue
overheadValue = objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 36), objWorkbookSheetRexmex.Cells(lastRowRexmex, 36)).SpecialCells(12).Value ' 12 = xlCellTypeVisible

If Trim(overheadValue) = "" Then
' No se encontró el valor de Overhead para el mes y año especificados
    ' Limpiar los filtros de la hoja objWorkbookSheetRexmex
    If objWorkbookSheetRexmex.AutoFilterMode Then
        objWorkbookSheetRexmex.AutoFilterMode = False
    End If
    If Not objWorkbookSheetRexmex.AutoFilterMode Then
        objWorkbookSheetRexmex.Rows(1).AutoFilter
    End If

    ' Guardar el libro de Overhead
    objWorkbookPathOverhead.Save
    objWorkbookPathOverhead.Close

    objWorkbookPathRexmex.Save
    objWorkbookPathRexmex.Close

    objExcel.Quit

    Msg = "Error was generated. " & "No se encontró el valor de Overhead para el mes " & mes & " del año " & anio
    WScript.StdOut.WriteLine Msg
Else
' Se encontró el valor de Overhead para el mes y año especificados
    ' Limpiar los filtros de la hoja objWorkbookSheetRexmex
    If objWorkbookSheetRexmex.AutoFilterMode Then
        objWorkbookSheetRexmex.AutoFilterMode = False
    End If
    If Not objWorkbookSheetRexmex.AutoFilterMode Then
        objWorkbookSheetRexmex.Rows(1).AutoFilter
    End If

    ' Pegar valo de Overhead en la celda C8 de la hoja objWorkbookSheetOverhead
    objWorkbookSheetOverhead.Range("C8").Value = overheadValue

    ' Guardar el libro de Overhead
    objWorkbookPathOverhead.Save
    objWorkbookPathOverhead.Close

    objWorkbookPathRexmex.Save
    objWorkbookPathRexmex.Close

    objExcel.Quit

    WScript.StdOut.WriteLine "Script executed successfully."
End If