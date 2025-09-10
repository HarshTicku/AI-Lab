' Module: Setup_Workbook
' Purpose: Create/rename tabs and insert Documentation sheet with guidance

Option Explicit

Public Sub Setup_Workbook()
    EnsureSheet "Data_Ingestion"
    EnsureSheet "Data_Cleaning"
    EnsureSheet "Analysis"
    EnsureSheet "Dashboard"
    EnsureDocumentation
    MsgBox "Workbook setup complete.", vbInformation
End Sub

Private Sub EnsureDocumentation()
    Dim ws As Worksheet
    Dim exists As Boolean
    exists = SheetExists("Documentation")
    If Not exists Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Documentation"
    Else
        Set ws = ThisWorkbook.Worksheets("Documentation")
    End If

    With ws
        .Cells.Clear
        .Range("A1").Value = "Documentation"
        .Range("A3").Value = "Tabs & Purpose"
        .Range("A4").Value = "Data_Ingestion: Paste/import raw listings data."
        .Range("A5").Value = "Data_Cleaning: Output of Run_Data_Cleaning macro."
        .Range("A6").Value = "Analysis: KPIs and staging tables for pivots."
        .Range("A7").Value = "Dashboard: Pivot charts, KPI cards, slicers."
        .Range("A9").Value = "Inputs"
        .Range("A10").Value = "Paste CSV/Excel extracts into Data_Ingestion (A1). Keep headers consistent."
        .Range("A12").Value = "Key Formulas"
        .Range("A13").Value = "ADR = NetRevenue / NightsOccupied"
        .Range("A14").Value = "RevPAR = ADR * OccupancyRate"
        .Range("A15").Value = "ROI = AnnualProfit / Investment"
        .Range("A17").Value = "Data Flow: Data_Ingestion → Data_Cleaning → Analysis → Dashboard"
        .Range("A19").Value = "Refresh Checklist"
        .Range("A20").Value = "1) Run Run_Data_Cleaning  2) Data → Refresh All  3) Validate slicers"
    End With
End Sub

Private Sub EnsureSheet(sheetName As String)
    If Not SheetExists(sheetName) Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
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
