Option Explicit

Sub SplitExcelFileByColumnPSum()
    ' Khai bao bien
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim sheetName As String
    Dim filePath As String
    Dim fileDialog As fileDialog
    Dim lastRow As Long
    Dim currentRow As Long
    Dim startRow As Long
    Dim fileCounter As Integer
    Dim sumValue As Double
    Dim targetFilePath As String
    Dim sourceFolder As String
    Dim i As Long
    Dim sourceFileName As String
    
    ' Khoi tao bien
    fileCounter = 1
    startRow = 4 ' Bat dau tu dong 4
    currentRow = 4
    sumValue = 0
    
    ' Buoc 1: Cho nguoi dung chon file Excel
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Chon file Excel can xu ly"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsm", 1
        .AllowMultiSelect = False
        
        If .Show = -1 Then ' Nguoi dung da chon file
            filePath = .SelectedItems(1)
        Else
            MsgBox "Ban chua chon file. Chuong trinh se dung lai.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Lay thu muc chua file goc
    sourceFolder = Left(filePath, InStrRev(filePath, "\"))
    
    ' Lay ten file goc (khong co extension)
    Dim tempFileName As String
    tempFileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    sourceFileName = Left(tempFileName, InStrRev(tempFileName, ".") - 1)
    
    ' Buoc 2: Nhap ten sheet can xu ly
    sheetName = InputBox("Nhap ten sheet can xu ly:", "Ten Sheet")
    If sheetName = "" Then
        MsgBox "Ban chua nhap ten sheet. Chuong trinh se dung lai.", vbExclamation
        Exit Sub
    End If
    
    ' Mo file Excel da chon
    On Error Resume Next
    Set wbSource = Workbooks.Open(filePath)
    If wbSource Is Nothing Then
        MsgBox "Khong the mo file Excel!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Kiem tra sheet co ton tai khong
    On Error Resume Next
    Set wsSource = wbSource.Worksheets(sheetName)
    If wsSource Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' khong ton tai trong file!", vbExclamation
        wbSource.Close False
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Tim dong cuoi cung co du lieu trong cot P
    lastRow = wsSource.Cells(wsSource.Rows.Count, "P").End(xlUp).Row
    
    ' Kiem tra neu khong co du lieu tu P4 tro di
    If lastRow < 4 Then
        MsgBox "Khong co du lieu trong cot P tu dong 4 tro di!", vbExclamation
        wbSource.Close False
        Exit Sub
    End If
    
    ' Tat cap nhat man hinh de tang toc do
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Buoc 3: Xu ly du lieu va tao file con
    Dim columnLValue As String
    Dim pairSum As Double
    Dim shouldSplit As Boolean
    Dim endRowForCurrentFile As Long
    Dim valueRow1 As Double
    Dim valueRow2 As Double
    
    pairSum = 0
    valueRow1 = 0
    valueRow2 = 0
    
    Do While currentRow <= lastRow
        columnLValue = Trim(CStr(wsSource.Range("L" & currentRow).Value))
        
        valueRow1 = wsSource.Range("P" & currentRow).Value
        valueRow2 = wsSource.Range("P" & currentRow + 1).Value
        
        ' Tinh tong cap ZPOS-ZNEG
        pairSum = valueRow1 + valueRow2
        
        ' Thu them cap vao sumValue de kiem tra
        Dim testSum As Double
        testSum = sumValue + pairSum
        
        shouldSplit = False
        endRowForCurrentFile = 0
        
        ' Kiem tra dieu kien
        If testSum = 1400000000 Then
            ' BANG 1.4 ty: them cap vao sum, tao file den het ZNEG
            shouldSplit = True
            endRowForCurrentFile = currentRow + 1
        End If
        
        If testSum > 1400000000 Then
            ' LON HON 1.4 ty: KHONG them cap vao sum, tao file den truoc ZPOS
            shouldSplit = True
            endRowForCurrentFile = currentRow - 1
        End If
        
        Dim currentLastRow As Long
        currentLastRow = currentRow + 1
        If currentLastRow >= lastRow Then
            shouldSplit = True
            endRowForCurrentFile = currentRow
        End If
        
        ' Tao file neu can
        If shouldSplit Then
            ' Tao workbook moi
            Set wbTarget = Workbooks.Add
            Set wsTarget = wbTarget.Worksheets(1)
            
            ' Dat ten sheet giong sheet goc
            wsTarget.Name = sheetName
            
            ' Copy header (3 dong dau luon luon)
            wsSource.Rows("1:3").Copy wsTarget.Rows("1:3")
            
            ' Copy du lieu tu startRow den endRowForCurrentFile vao dong 4 cua file moi
            wsSource.Rows(startRow & ":" & endRowForCurrentFile).Copy wsTarget.Range("A4")
            
            ' 1. Dem so dong tu P4 den cuoi cot P co du lieu, ghi vao P2
            Dim lastRowP As Long
            Dim rowCount As Long
            lastRowP = wsTarget.Cells(wsTarget.Rows.Count, "P").End(xlUp).Row
            rowCount = lastRowP - 3 ' Tru di 3 dong header
            wsTarget.Range("P2").Value = Format(rowCount, "000000")
            
            ' 2. Them "_" + fileCounter vao cuoi o B2
            If wsTarget.Range("B2").Value <> "" Then
                wsTarget.Range("B2").Value = wsTarget.Range("B2").Value & "_" & fileCounter
            End If
            
            ' 3. Them "_" + fileCounter vao tat ca o tu B4 den cuoi cot B co du lieu
            Dim lastRowB As Long
            Dim iRow As Long
            lastRowB = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
            For iRow = 4 To lastRowB
                If wsTarget.Range("B" & iRow).Value <> "" Then
                    wsTarget.Range("B" & iRow).Value = wsTarget.Range("B" & iRow).Value & "_" & fileCounter
                End If
            Next iRow
            
            ' Dieu chinh do rong cot
            wsTarget.Columns.AutoFit
            
            ' Dat ten va luu file con
            targetFilePath = sourceFolder & sourceFileName & "_" & fileCounter & ".xlsx"
            wbTarget.SaveAs fileName:=targetFilePath, FileFormat:=xlOpenXMLWorkbook
            wbTarget.Close False
            
            ' Thong bao tien trinh
            Debug.Print "Da tao file: " & targetFilePath & " (Tong: " & Format(sumValue, "#,##0") & ")"
            
            ' Tang bo dem file
            fileCounter = fileCounter + 1
            
            ' Reset cho file moi
            startRow = endRowForCurrentFile + 1
            sumValue = 0
            pairSum = 0
            valueRow1 = 0
            valueRow2 = 0
        Else
            ' Chua dat nguong: them cap vao sum, tiep tuc
            sumValue = testSum
        End If

        ' Tang dong hien tai
        If shouldSplit Then
            currentRow = endRowForCurrentFile + 1 ' Di den cap tiep theo
        Else
            currentRow = currentRow + 2 ' Neu con 1 dong cuoi, di
        End If
    Loop
    
    ' Dong file goc
    wbSource.Close False
    
    ' Bat lai cap nhat man hinh
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Thong bao hoan thanh
    MsgBox "Hoan thanh! Da tao " & (fileCounter - 1) & " file Excel con." & vbCrLf & _
           "Cac file duoc luu tai: " & sourceFolder, vbInformation, "Thanh cong"
    
End Sub

' Ham phu: Hien thi form nhap lieu (neu can giao dien dep hon)
Sub ShowInputForm()
    ' Co the tao UserForm de giao dien chuyen nghiep hon
    ' Hien tai su dung InputBox va FileDialog
    Call SplitExcelFileByColumnPSum
End Sub
