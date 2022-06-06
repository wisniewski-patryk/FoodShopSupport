Attribute VB_Name = "SummaryStatement"
Sub GetProductsSet()
    Dim filePath As String

    filePath = GetFile
        'Debug.Print filePath
    
    'Get data workbook
    If filePath = "" Then
        MsgBox "Nie wybrano, lub wybrano niewspierany plik"
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
    Dim i As Integer
    For i = FirstDataRow To LastDataRow
        AddProductToList (wsZamowienia.Cells(i, 17).Value)
    Next i
    
    wb.Close
    
    
    
End Sub

Private Function FindLastEmptyRow(WsCol As Worksheet, colNb As Integer) As Integer

    FindLastEmptyRow = WsCol.Cells(Rows.Count, colNb).End(xlUp).Row + 1
    
End Function

Private Function GetLastRow(WsCol As Worksheet, colNb As Integer) As Integer

    GetLastRow = WsCol.Cells(Rows.Count, colNb).End(xlUp).Row
    
End Function

Private Sub AddProductToList(Prod As String)
    Dim wb As Workbook
    Set wb = Workbooks("ZestawienieIlosciowe.xlsm")
    Dim ProductListStart As Integer
    ProductListStart = 2
    Dim ProductListEnd As Integer
    ProductListEnd = GetLastRow(wb.Worksheets("Sheet1"), 1)
    Dim i As Integer
    For i = ProductListStart To ProductListEnd
        If wb.Worksheets("Sheet1").Cells(i, 1).Value = Prod Then
            wb.Worksheets("Sheet1").Cells(i, 2).Value = wb.Worksheets("Sheet1").Cells(i, 2).Value + 1
            Exit Sub
        End If
    Next i
    Dim nextRow As Integer
    nextRow = FindLastEmptyRow(wb.Worksheets("Sheet1"), 1)
    wb.Worksheets("Sheet1").Cells(nextRow, 1).Value = Prod
    wb.Worksheets("Sheet1").Cells(nextRow, 2).Value = 1
    
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
    Next i
    
End Sub
