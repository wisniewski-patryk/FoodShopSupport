Attribute VB_Name = "Etykiety"
Sub MakeLabels()

    Dim filePath As String
    filePath = GetFile
        'Debug.Print filePath
    
    'Get data workbook
    If filePath = "" Then
        Exit Sub
    End If
    
    Dim fileName As String
    fileName = GetFileName(filePath)
        Debug.Print "FileName= " & fileName
    
    Workbooks.Open filePath
    Dim wb As Workbook
    Set wb = Workbooks(fileName)
        Debug.Print "Workbook " & fileName & " is set as wb"
    wb.Activate
        Debug.Print "wb is Active"
        
    Dim wbLabels As Workbook
    Set wbLabels = Workbooks("Etykiety.xlsm")
    
    Dim wsLabels As Worksheet
        
    'Get data worksheet
    Dim wsOrders As Worksheet
    Set wsOrders = wb.Worksheets("Zamówienia")
        Debug.Print "wsOrders is set as worksheet zamówienia"
        
    Dim FirstDataRowInOrders As Integer
    FirstDataRowInOrders = 2
        Debug.Print "FirstDataRowInOrders   = " & FirstDataRowInOrders
        
    Dim LastDataRowInOrders As Integer
    LastDataRowInOrders = GetLastRow(wsOrders, 17) 'Row 'Q'
        Debug.Print "LastDataRowInOrders    = " & LastDataRowInOrders
        
    Dim LabelsPage, LabelRow, LabelColumn As Integer
    LabelsPage = 0
    LabelRow = -1
    LabelColumn = 0
    
    Dim newLabelsSetName As String
    
    Dim IdOrder As Integer          'row 1  - 'A'
    Dim ClientName As String        'row 26 - 'Z'
    Dim ClientForname As String     'row 27 - 'AA'
    Dim ClientAdress As String      'row 41 - 'AO'
    Dim ClientCity As String        'row 40 - 'AN'
    Dim ClientPhone As String       'row 39 - 'AM'
    Dim ClientZipCode As String     'row 43 - 'AQ'
    Dim ClientMsg As String         'row 44 - 'AR'
    Dim Recycling As String         'row 45 - 'AS' 1 = N ; 0 = T
    Dim PaymentMethod As String     'row 5  - 'E'
    
    Dim ProductName As String       'row 18 - 'R' -> zmiana na col 25 'Y'
    Dim ProductQuantity As Integer  'row 19 - 'S'
    
    Dim ProductOffsetRow, ProductOffsetColumn As Integer
    ProductOffsetRow = 0
    ProductOffsetColumn = 0


    Dim i As Integer

    For i = FirstDataRowInOrders To LastDataRowInOrders
    
        'making new sheet
        If (LabelRow = -1) Then
            LabelsPage = LabelsPage + wbLabels.Worksheets.Count
            newLabelsSetName = "Labels " & LabelsPage
            NewLabelsSet (newLabelsSetName)
            Set wsLabels = wbLabels.Worksheets(newLabelsSetName)
            LabelRow = 0
        End If
        
        'getting data
        If Not IsEmpty(wsOrders.Cells(i, 1)) Then
            IdOrder = wsOrders.Cells(i, 1)
            ClientName = wsOrders.Cells(i, 26)
            ClientForname = wsOrders.Cells(i, 27)
            ClientAdress = wsOrders.Cells(i, 41)
            ClientCity = wsOrders.Cells(i, 40)
            ClientPhone = wsOrders.Cells(i, 39)
            ClientZipCode = wsOrders.Cells(i, 43)
            ClientMsg = wsOrders.Cells(i, 44)
            If wsOrders.Cells(i, 45) = 1 Then
                Recycling = "N"
            Else
                Recycling = "T"
            End If
            PaymentMethod = wsOrders.Cells(i, 5)
            
            Dim RowTemp, ColTemp As Integer
            RowTemp = 1 + (8 * LabelRow)
            ColTemp = 1 + (6 * LabelColumn)
            
            wsLabels.Cells(RowTemp, ColTemp).Select
            If PaymentMethod = "Cash on delivery" Or PaymentMethod = "P?atno?? przy odbiorze" Then
                Selection.Offset(5, 1) = "Przy odbiorze"
            Else
                Selection.Offset(5, 1) = PaymentMethod
            End If
            Selection.Offset(6, 0) = ClientName
            Selection.Offset(6, 2) = ClientPhone
            Selection.Offset(7, 0) = ClientAdress
            Selection.Offset(5, 3) = ClientMsg
            Selection.Offset(4, 5) = Recycling
            
            If LabelColumn = 0 Then
                LabelColumn = 1
            Else
                LabelColumn = 0
                LabelRow = LabelRow + 1
            End If
            
            'If LabelRow = 7 Then LabelRow = -1
            
            ProductOffsetColumn = 0
            ProductOffsetRow = 0
            
        End If
                
        ProductName = wsOrders.Cells(i, 25)
        ProductQuantity = wsOrders.Cells(i, 19)
        
        Selection.Offset(ProductOffsetRow, ProductOffsetColumn) = ProductName
        Selection.Offset(ProductOffsetRow, ProductOffsetColumn + 1) = ProductQuantity
        
        ProductOffsetRow = ProductOffsetRow + 1
        
        If ProductOffsetRow > 4 Then
            ProductOffsetColumn = ProductOffsetColumn + 2
            ProductOffsetRow = 0
        End If
        
        
    Next i
   wb.Close
   
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
Private Function GetFileName(filePath As String) As String
    GetFileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
End Function

Private Function FindLastEmptyRow(WsCol As Worksheet, colNb As Integer) As Integer
    FindLastEmptyRow = WsCol.Cells(Rows.Count, colNb).End(xlUp).Row + 1
End Function

Private Function GetLastRow(WsCol As Worksheet, colNb As Integer) As Integer
    GetLastRow = WsCol.Cells(Rows.Count, colNb).End(xlUp).Row
End Function

Private Sub NewLabelsSet(LabelName As String)
    Dim workbookName, worksheetTempletName As String
    workbookName = "Etykiety.xlsm"
    worksheetTempetName = "SZABLON"
    Workbooks(workbookName).Worksheets(worksheetTempetName).Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = LabelName
End Sub
