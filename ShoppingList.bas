Attribute VB_Name = "ShoppingList"
Sub ListaZakupow()
    Dim fpZestawienieIlosciowe As String
    fpZestawienieIlosciowe = GetFile
        Debug.Print fpZestawienieIlosciowe
    
    'Get data workbook
    If fpZestawienieIlosciowe = "" Then
        Exit Sub
    End If
    
    Dim fileNameZestaweienieIlosciowe As String
    fileNameZestaweienieIlosciowe = Right(fpZestawienieIlosciowe, Len(fpZestawienieIlosciowe) - InStrRev(fpZestawienieIlosciowe, "\"))
    
    Workbooks.Open fpZestawienieIlosciowe
    Dim wbZestawienieIlosciowe As Workbook
    Set wbZestawienieIlosciowe = Workbooks(fileNameZestaweienieIlosciowe)
    wbZestawienieIlosciowe.Activate
    
    'Get data worksheet
    Dim wsZestawienieIlosciowe As Worksheet
    Set wsZestawienieIlosciowe = wbZestawienieIlosciowe.Worksheets("Sheet1")
    
    'Set first and last row of data
    Dim FirstDataRow As Integer
    FirstDataRow = 2
    Dim LastDataRow As Integer
    LastDataRow = GetLastRow(wsZestawienieIlosciowe, 1)
    
        Debug.Print "FirstDataRow   = " & FirstDataRow
        Debug.Print "LastDataRow    = " & LastDataRow
    
    'set target variable
    Dim wbLista As Workbook
    Set wbLista = Workbooks("ListaZakupow.xlsm")
    
    Dim wsLista As Worksheet
    Set wsLista = wbLista.Worksheets("Sheet1")
    
    Dim wsDb As Worksheet
    Set wsDb = wbLista.Worksheets("sety_db")
    
    'scrap data
    Dim Product As String
    Dim ProductId As String
    Dim ProductToAdd As String
    Dim ProductQuantity As Integer
    Dim PrducctQuantityToAdd As Integer
    Dim ProductWeightToAdd As Double
    
    
    Dim dbStart, dbEnd, dbElement As Integer
    dbStart = 2
    dbEnd = GetLastRow(wsDb, 1)
    
    For i = FirstDataRow To LastDataRow
        Product = wsZestawienieIlosciowe.Cells(i, 1)
        ProductId = wsZestawienieIlosciowe.Cells(i, 2)
        ProductQuantity = wsZestawienieIlosciowe.Cells(i, 3)
        
        For dbElement = dbStart To dbEnd
            If ProductId = wsDb.Cells(dbElement, 2) Then
                ProductToAdd = wsDb.Cells(dbElement, 3)
                PrducctQuantityToAdd = ProductQuantity * wsDb.Cells(dbElement, 5).Value
                ProductWeightToAdd = ProductQuantity * wsDb.Cells(dbElement, 6).Value
                
                Call AddToList(ProductToAdd, PrducctQuantityToAdd, ProductWeightToAdd)
                
            End If
            
        Next dbElement
        
    Next i
    
    
    wbZestawienieIlosciowe.Close
    
    
End Sub

Private Function GetFile() As String
    Dim fileDialog As fileDialog
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    fileDialog.AllowMultiSelect = False
    fileDialog.InitialFileName = "C:\Users\wisniewski-patryk\AppData\Local\obsidian-vaults\wisniewski-patryk\_PRIVATE\FUCHY\Sklep\"
    fileDialog.Filters.Add "Excel file with Macros", "*.xlsm"
    fileDialog.Filters.Add "Excel file", "*.xlsx"
    fileDialog.Filters.Add "All files", "*.*"
    If fileDialog.Show = -1 Then
        GetFile = fileDialog.SelectedItems(1)
    Else
        GetFile = ""
    End If
    
End Function

Private Function FindLastEmptyRow(WsCol As Worksheet, colNb As Integer) As Integer

    FindLastEmptyRow = WsCol.Cells(Rows.Count, colNb).End(xlUp).Row + 1
    
End Function

Private Function GetLastRow(WsCol As Worksheet, colNb As Integer) As Integer

    GetLastRow = WsCol.Cells(Rows.Count, colNb).End(xlUp).Row
    
End Function

Private Sub AddToList(Product As String, quantity As Integer, weight As Double)
    Dim ws As Worksheet
    Set ws = Workbooks("ListaZakupow.xlsm").Worksheets("Sheet1")
    
    
    Dim ListBegining, ListEnd, nextRow, i As Integer
    ListBeginig = 2
    ListEnd = GetLastRow(ws, 1)
    
    If ListEnd = 1 Then ListEnd = 2
    
    For i = ListBeginig To ListEnd Step 1
        If ws.Cells(i, 1).Value = Product Then
            ws.Cells(i, 2) = ws.Cells(i, 2).Value + quantity
            ws.Cells(i, 3) = ws.Cells(i, 3).Value + weight
            Exit Sub
        End If
        
    Next i
    
    nextRow = FindLastEmptyRow(ws, 1)
    ws.Cells(nextRow, 1).Value = Product
    ws.Cells(nextRow, 2).Value = quantity
    ws.Cells(nextRow, 3).Value = weight
        
    
End Sub

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

