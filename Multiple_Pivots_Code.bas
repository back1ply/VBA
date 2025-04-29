Attribute VB_Name = "Multiple_Pivots_Code"
Sub Insert_Multiple_Pivot_Tables()

    'Declare Variables
    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim PCache As PivotCache
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim LastRow As Long
    Dim LastCol As Long
    Dim pvt As PivotTable

    'Delete and Add Worksheet
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

    'Create Pivot Cache Once
    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange)

    'Pivot 1 – Region-Wise Total Sales
    Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Range("A1"), TableName:="Pivot1_RegionSales")
    With PTable.PivotFields("Region")
        .Orientation = xlRowField
        .Position = 1
    End With
    With PTable.PivotFields("Total Sales")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Pivot1_RegionSales"
    End With

    'Pivot 2 – Product-Wise Total Sales
    Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Range("A8"), TableName:="Pivot2_ProductSales")
    With PTable.PivotFields("Product")
        .Orientation = xlRowField
        .Position = 1
    End With
    With PTable.PivotFields("Total Sales")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Pivot2_ProductSales"
    End With

    'Pivot 3 – Payment Mode Wise Total Sales
    Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Range("D1"), TableName:="Pivot3_PaymentSales")
    With PTable.PivotFields("Payment Mode")
        .Orientation = xlRowField
        .Position = 1
    End With
    With PTable.PivotFields("Total Sales")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Pivot3_PaymentSales"
    End With

    'Pivot 4 – Delivery Status Wise Units Sold
    Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Range("D9"), TableName:="Pivot4_DeliveryUnits")
    With PTable.PivotFields("Delivery Status")
        .Orientation = xlRowField
        .Position = 1
    End With
    With PTable.PivotFields("Units Sold")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Pivot4_DeliveryUnits"
    End With

    'Pivot 5 – Customer Type Wise Total Sales
    Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Range("G1"), TableName:="Pivot5_CustomerSales")
    With PTable.PivotFields("Customer Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    With PTable.PivotFields("Total Sales")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Pivot5_CustomerSales"
    End With

    'Pivot 6 – Order Priority Wise Units Sold
    Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Range("A19"), TableName:="Pivot6_PriorityUnits")
    With PTable.PivotFields("Order Priority")
        .Orientation = xlRowField
        .Position = 1
    End With
    With PTable.PivotFields("Units Sold")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Pivot6_PriorityUnits"
    End With

    'Pivot 7 – Warranty Wise Units Sold
    Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Range("D16"), TableName:="Pivot7_WarrantyUnits")
    With PTable.PivotFields("Warranty")
        .Orientation = xlRowField
        .Position = 1
    End With
    With PTable.PivotFields("Units Sold")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Pivot7_WarrantyUnits"
    End With

    'Pivot 8 – Return Eligibility Wise Units Sold
    Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Range("G8"), TableName:="Pivot8_ReturnUnits")
    With PTable.PivotFields("Return Eligibility")
        .Orientation = xlRowField
        .Position = 1
    End With
    With PTable.PivotFields("Units Sold")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Pivot8_ReturnUnits"
    End With

    'Loop through each Pivot Table on the PivotTable worksheet to apply formatting
    For Each pvt In PSheet.PivotTables
        pvt.ShowTableStyleRowStripes = True
        pvt.TableStyle2 = "PivotStyleMedium9"
    Next pvt

End Sub

