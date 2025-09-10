' Module: Run_Data_Cleaning
' Purpose: Consolidate all cleaning steps in a single macro
' Tabs: Data_Ingestion (input) -> Data_Cleaning (output)
' Actions: remove duplicates, handle missing values, normalize text/dates, lookups

Option Explicit

Public Sub Run_Data_Cleaning()
    On Error GoTo CleanFail

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim lastRow As Long, lastCol As Long

    Set wsSrc = ThisWorkbook.Worksheets("Data_Ingestion")
    Set wsDst = ThisWorkbook.Worksheets("Data_Cleaning")

    wsDst.Cells.Clear

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Copy wsDst.Range("A1")

    With wsDst
        ' Trim and standardize text in used range
        Dim rngUsed As Range, cell As Range
        Set rngUsed = .Range("A2", .Cells(.Rows.Count, 1).End(xlUp)).EntireRow
        For Each cell In rngUsed
            If VarType(cell.Value) = vbString Then cell.Value = Application.WorksheetFunction.Trim(cell.Value)
        Next cell

        ' Normalize dates (example: Date in column E)
        On Error Resume Next
        .Columns("E").NumberFormat = "yyyy-mm-dd"
        On Error GoTo 0

        ' Remove exact duplicates based on all columns
        Dim colCount As Long
        colCount = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Dim dupCols As Variant
        dupCols = Evaluate("ROW(1:" & colCount & ")")
        .Range("A1").CurrentRegion.RemoveDuplicates Columns:=dupCols, Header:=xlYes

        ' Handle missing values: fill zeros for numeric, "Unknown" for text
        Dim dataRegion As Range
        Set dataRegion = .Range("A2").CurrentRegion
        For Each cell In dataRegion
            If Len(CStr(cell.Value)) = 0 Then
                If IsNumeric(cell.Offset(0, 0).Value) Then
                    cell.Value = 0
                Else
                    cell.Value = "Unknown"
                End If
            End If
        Next cell

        ' Optional enrichment: City from Zip via XLOOKUP on Lookup_ZipToCity
        If SheetExists("Lookup_ZipToCity") Then
            Dim lastDataRow As Long
            lastDataRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            Dim cityCol As Long: cityCol = GetOrCreateColumn(.Rows(1), "City")
            Dim zipCol As Long: zipCol = FindColumn(.Rows(1), "Zip")
            If zipCol > 0 Then
                .Range(.Cells(2, cityCol), .Cells(lastDataRow, cityCol)).Formula2R1C1 = _
                    "=IFERROR(XLOOKUP(RC" & zipCol & ",Lookup_ZipToCity!C1,Lookup_ZipToCity!C2,\"Unknown\"),RC" & cityCol & ")"
            End If
        End If
    End With

    MsgBox "Data cleaning completed.", vbInformation
    Exit Sub

CleanFail:
    MsgBox "Data cleaning failed: " & Err.Description, vbCritical
End Sub

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    SheetExists = False
End Function

Private Function FindColumn(headerRow As Range, headerName As String) As Long
    Dim cell As Range
    For Each cell In headerRow.Cells
        If LCase(Trim(CStr(cell.Value))) = LCase(Trim(headerName)) Then
            FindColumn = cell.Column
            Exit Function
        End If
        If Len(CStr(cell.Value)) = 0 Then Exit For
    Next cell
    FindColumn = 0
End Function

Private Function GetOrCreateColumn(headerRow As Range, headerName As String) As Long
    Dim col As Long
    col = FindColumn(headerRow, headerName)
    If col > 0 Then
        GetOrCreateColumn = col
    Else
        GetOrCreateColumn = headerRow.Cells(1, headerRow.Cells(1, headerRow.Columns.Count).End(xlToLeft).Column + 1).Column
        headerRow.Cells(1, GetOrCreateColumn).Value = headerName
    End If
End Function
