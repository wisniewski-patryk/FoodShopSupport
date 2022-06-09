Attribute VB_Name = "SummaryStatement"
Sub GetProductsSet()
    Dim filePath As String

    filePath = GetFile
        'Debug.Print filePath
    
    'Get data workbook
    If filePath = "" Then
        Exit Sub
    End If
    
    Dim fileName As String
    fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
    
    Workbooks.Open filePath
    Dim wb As Workbook
    Set wb = Workbooks(fileName)
    wb.Activate
    
    'Get data worksheet
    Dim wsZamowienia As Worksheet
    Set wsZamowienia = wb.Worksheets("Zamówienia")
    
    'Set first and last row
    Dim FirstDataRow As Integer
    FirstDataRow = 2
    Dim LastDataRow As Integer
    LastDataRow = GetLastRow(wsZamowienia, 17) 'kolumna 'Nazwa produktu'
    
        'Debug.Print "FirstDataRow   = " & FirstDataRow
        'Debug.Print "LastDataRow    = " & LastDataRow
    
    'Scrap data
    Dim ProductName As String
    Dim ProductQuantity As Integer
    Dim ProductId As String
    
    Dim i As Integer
    For i = FirstDataRow To LastDataRow
        ProductName = wsZamowienia.Cells(i, 17).Value
        ProductId = wsZamowienia.Cells(i, 25).Value
        ProductQuantity = wsZamowienia.Cells(i, 19).Value
        Call AddProductToList(ProductName, ProductId, ProductQuantity)
        
    Next i
    
    wb.Close
    
    
    
End Sub

Private Function FindLastEmptyRow(WsCol As Worksheet, colNb As Integer) As Integer

    FindLastEmptyRow = WsCol.Cells(Rows.Count, colNb).End(xlUp).Row + 1
    
End Function

Private Function GetLastRow(WsCol As Worksheet, colNb As Integer) As Integer

    GetLastRow = WsCol.Cells(Rows.Count, colNb).End(xlUp).Row
    
End Function

Private Sub AddProductToList(Prod As String, ProdId As String, ProdQuantity As Integer)
    Dim wb As Workbook
    Set wb = Workbooks("ZestawienieIlosciowe.xlsm")
    Dim ProductListStart As Integer
    ProductListStart = 2
    Dim ProductListEnd As Integer
    ProductListEnd = GetLastRow(wb.Worksheets("Sheet1"), 1)
    Dim i As Integer
    For i = ProductListStart To ProductListEnd
        If wb.Worksheets("Sheet1").Cells(i, 2).Value = ProdId Then
            wb.Worksheets("Sheet1").Cells(i, 3).Value = wb.Worksheets("Sheet1").Cells(i, 3).Value + ProdQuantity
            Exit Sub
        End If
    Next i
    Dim nextRow As Integer
    nextRow = FindLastEmptyRow(wb.Worksheets("Sheet1"), 1)
    wb.Worksheets("Sheet1").Cells(nextRow, 1).Value = Prod
    wb.Worksheets("Sheet1").Cells(nextRow, 2).Value = ProdId
    wb.Worksheets("Sheet1").Cells(nextRow, 3).Value = ProdQuantity
    
End Sub

Private Function GetFile() As String
    Dim fileDialog As fileDialog
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    fileDialog.AllowMultiSelect = False
    fileDialog.InitialFileName = "C:\Users\wisniewski-patryk\AppData\Local\obsidian-vaults\wisniewski-patryk\_PRIVATE\FUCHY\Sklep\"
    fileDialog.Filters.Add "Excel file", "*.xlsx"
    fileDialog.Filters.Add "All files", "*.*"
    If fileDialog.Show = -1 Then
        GetFile = fileDialog.SelectedItems(1)
    Else
        GetFile = ""
    End If
    
End Function

Sub Clean()
    'Set first and last row
    Dim FirstDataRow As Integer
    FirstDataRow = 2
    Dim LastDataRow As Integer
    LastDataRow = GetLastRow(ActiveWorkbook.Worksheets("Sheet1"), 1)
    Dim i As Integer
    For i = FirstDataRow To LastDataRow
        ActiveWorkbook.Worksheets("Sheet1").Cells(i, 1).ClearContents
        ActiveWorkbook.Worksheets("Sheet1").Cells(i, 2).ClearContents
        ActiveWorkbook.Worksheets("Sheet1").Cells(i, 3).ClearContents
    Next i
    
End Sub
