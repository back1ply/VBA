Attribute VB_Name = "Single_Pivot_Code"
Sub InsertPivotTable()
'Macro By ExcelChamps.com

'Declare Variables
    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim PCache As PivotCache
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim LastRow As Long
    Dim LastCol As Long

'Insert a New Blank Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("PivotTable").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Sheets.Add Before:=ActiveSheet
    ActiveSheet.Name = "PivotTable"
    Set PSheet = Worksheets("PivotTable")
    Set DSheet = Worksheets("Sales_Data")
    
'Define Data Range
    LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
    LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)
    
'Define Pivot Cache
    Set PCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=PRange)

'Insert Pivot Table
    Set PTable = PCache.CreatePivotTable( _
        TableDestination:=PSheet.Cells(2, 2), TableName:="SalesPivotTable")
    
'Insert Row Fields
    With PTable.PivotFields("Region")
        .Orientation = xlRowField
        .Position = 1
    End With
    With PTable.PivotFields("Salesperson")
        .Orientation = xlRowField
        .Position = 2
    End With

'Insert Column Fields
    With PTable.PivotFields("Product")
        .Orientation = xlColumnField
        .Position = 1
    End With

'Insert Data Field
    With PTable.PivotFields("Total Sales")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "#,##0"
        .Name = "Revenue"
    End With

'Format Pivot Table
    PTable.ShowTableStyleRowStripes = True
    PTable.TableStyle2 = "PivotStyleMedium9"
    
End Sub




