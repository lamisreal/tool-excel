Option Explicit

' ============================================================
' VBA_MainTool.bas
' Chuong trinh chinh - Main Entry Point cho tat ca tool cua Nhi
' Tao: 2026 | Cap nhat: 25/02/2026
' ============================================================

' ------------------------------------------------------------
' ENTRY POINT: Hien thi form chinh
' Goi sub nay de mo man hinh quan ly tool
' ------------------------------------------------------------
Sub ShowMainTool()

    Dim choice As String

    Do
        choice = InputBox( _
            "============================================" & vbCrLf & _
            "       TAT CA CAC TOOL CUA NHI             " & vbCrLf & _
            "============================================" & vbCrLf & _
            "" & vbCrLf & _
            "  1. Tach file (Split Excel by Sum cot P)  " & vbCrLf & _
            "  2. Tao sheet bang sort cot               " & vbCrLf & _
            "" & vbCrLf & _
            "Nhap so thu tu cong cu (1 hoac 2):" & vbCrLf & _
            "(De trong hoac bam Cancel de thoat)" & vbCrLf & _
            "" & vbCrLf & _
            Chr(169) & " 2026 Nhi's VBA Toolkit  |  v1.0", _
            "Tool cua Nhi")

        Select Case Trim(choice)
            Case "1"
                Call ShowInputForm
                Exit Do
            Case "2"
                Call SortAndCreateSheets
                Exit Do
            Case ""   ' Bam Cancel hoac de trong: thoat
                Exit Do
            Case Else
                MsgBox "Vui long nhap 1 hoac 2.", vbExclamation, "Lua chon khong hop le"
        End Select
    Loop

End Sub


' ============================================================
' TOOL 1: TACH FILE THEO TONG COT P
' Giu nguyen toan bo logic goc tu VBA_SplitExcelBySum.bas
' ============================================================

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
    Call SplitExcelFileByColumnPSum
End Sub


' ============================================================
' TOOL 2: TAO SHEET BANG SORT COT
' Buoc 1: Chon file Excel dau vao
' Buoc 2: Nhap ten sheet can doc
' Buoc 3: Nhap cot can doc
' Buoc 4: Sort ASC -> moi gia tri unique = 1 sheet (nhieu dong
'         trung nhau van gop vao chung 1 sheet do)
' Buoc 5: Luu file moi, xoa Sheet1 mac dinh, mo len cho xem
' ============================================================

Sub SortAndCreateSheets()
    ' ---- Khai bao bien ----
    Dim filePath     As String
    Dim sheetName    As String
    Dim colLetter    As String
    Dim wbSource     As Workbook
    Dim wsSource     As Worksheet
    Dim wbDest       As Workbook
    Dim lastRow      As Long
    Dim colIndex     As Long
    Dim newFilePath  As String
    Dim sourceFolder As String

    ' ---- Buoc 1: Chon file Excel dau vao ----
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Chon file Excel dau vao"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsm", 1
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "Ban chua chon file. Chuong trinh se dung lai.", vbExclamation
            Exit Sub
        End If
    End With

    sourceFolder = Left(filePath, InStrRev(filePath, "\" ))

    ' ---- Buoc 2: Nhap ten sheet ----
    sheetName = Trim(InputBox("Nhap ten sheet can doc:", "Ten Sheet"))
    If sheetName = "" Then
        MsgBox "Ban chua nhap ten sheet. Chuong trinh se dung lai.", vbExclamation
        Exit Sub
    End If

    ' ---- Buoc 3: Nhap cot can doc ----
    colLetter = UCase(Trim(InputBox("Nhap ky tu cot can doc:" & vbCrLf & "Vi du: A, B, C, ...", "Cot can doc")))
    If colLetter = "" Then
        MsgBox "Ban chua nhap cot. Chuong trinh se dung lai.", vbExclamation
        Exit Sub
    End If

    ' ---- Mo file nguon ----
    On Error Resume Next
    Set wbSource = Workbooks.Open(filePath)
    If wbSource Is Nothing Then
        MsgBox "Khong the mo file Excel dau vao!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    On Error Resume Next
    Set wsSource = wbSource.Worksheets(sheetName)
    If wsSource Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' khong ton tai trong file!", vbExclamation
        wbSource.Close False
        Exit Sub
    End If
    On Error GoTo 0

    ' Kiem tra cot hop le
    On Error Resume Next
    colIndex = wsSource.Columns(colLetter).Column
    If Err.Number <> 0 Then
        MsgBox "Cot '" & colLetter & "' khong hop le!", vbExclamation
        wbSource.Close False
        Exit Sub
    End If
    On Error GoTo 0

    ' Kiem tra du lieu tu dong 2 tro di
    lastRow = wsSource.Cells(wsSource.Rows.Count, colLetter).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Khong co du lieu tu dong 2 tro di trong cot " & colLetter & "!", vbExclamation
        wbSource.Close False
        Exit Sub
    End If

    ' ---- Tao workbook dich moi ----
    Set wbDest = Workbooks.Add

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---- Buoc 5a: Thu thap gia tri cot tu dong 2 den lastRow ----
    Dim totalRows As Long
    totalRows = lastRow - 1   ' so rows tu dong 2 den lastRow

    Dim colValues() As String
    ReDim colValues(1 To totalRows)
    Dim r As Long
    For r = 2 To lastRow
        colValues(r - 1) = Trim(CStr(wsSource.Cells(r, colLetter).Value))
    Next r

    ' ---- Buoc 5b: Sort ASC (Bubble Sort) ----
    Dim ii As Long, jj As Long, swapVal As String
    For ii = 1 To totalRows - 1
        For jj = 1 To totalRows - ii
            If colValues(jj) > colValues(jj + 1) Then
                swapVal = colValues(jj)
                colValues(jj) = colValues(jj + 1)
                colValues(jj + 1) = swapVal
            End If
        Next jj
    Next ii

    ' ---- Buoc 5c: Lay danh sach gia tri unique (da sort) ----
    Dim uniqueVals() As String
    Dim uniqueCount  As Long
    ReDim uniqueVals(1 To totalRows)
    uniqueCount = 0
    Dim prevVal As String
    prevVal = Chr(0)   ' Gia tri bat dau khac moi du lieu
    For r = 1 To totalRows
        If colValues(r) <> prevVal Then
            uniqueCount = uniqueCount + 1
            uniqueVals(uniqueCount) = colValues(r)
            prevVal = colValues(r)
        End If
    Next r

    ' ---- Buoc 5d: Tao sheet rieng cho tung gia tri unique ----
    Dim u            As Long
    Dim createdCount As Long
    Dim wsNew        As Worksheet
    createdCount = 0

    For u = 1 To uniqueCount
        Dim targetVal As String
        targetVal = uniqueVals(u)

        ' Xu ly ten sheet hop le (<= 31 ky tu, khong co ky tu dac biet)
        Dim safeSheetName As String
        safeSheetName = Left(targetVal, 31)
        safeSheetName = Application.WorksheetFunction.Substitute(safeSheetName, "/", "-")
        safeSheetName = Application.WorksheetFunction.Substitute(safeSheetName, "\\", "-")
        safeSheetName = Application.WorksheetFunction.Substitute(safeSheetName, "*", "")
        safeSheetName = Application.WorksheetFunction.Substitute(safeSheetName, "[", "")
        safeSheetName = Application.WorksheetFunction.Substitute(safeSheetName, "]", "")
        safeSheetName = Application.WorksheetFunction.Substitute(safeSheetName, ":", "-")
        safeSheetName = Application.WorksheetFunction.Substitute(safeSheetName, "?", "")
        If Trim(safeSheetName) = "" Then safeSheetName = "Sheet_" & u

        ' Xu ly trung ten sheet trong wbDest
        Dim finalName As String
        finalName = safeSheetName
        Dim nameIdx    As Integer
        Dim nameExists As Boolean
        nameIdx = 1
        Do
            nameExists = False
            Dim wsChk As Worksheet
            For Each wsChk In wbDest.Worksheets
                If wsChk.Name = finalName Then
                    nameExists = True
                    Exit For
                End If
            Next wsChk
            If nameExists Then
                nameIdx = nameIdx + 1
                finalName = Left(safeSheetName, 28) & "_" & nameIdx
            End If
        Loop While nameExists

        ' Tao sheet moi trong wbDest
        Set wsNew = wbDest.Worksheets.Add(After:=wbDest.Worksheets(wbDest.Worksheets.Count))
        wsNew.Name = finalName

        ' Copy dong tieu de (dong 1) tu wsSource
        wsSource.Rows(1).Copy wsNew.Rows(1)

        ' Copy cac dong co gia tri cot = targetVal
        Dim destRowIdx As Long
        destRowIdx = 2
        For r = 2 To lastRow
            If Trim(CStr(wsSource.Cells(r, colLetter).Value)) = targetVal Then
                wsSource.Rows(r).Copy wsNew.Rows(destRowIdx)
                destRowIdx = destRowIdx + 1
            End If
        Next r

        wsNew.Columns.AutoFit
        createdCount = createdCount + 1
        Debug.Print "Da tao sheet: " & finalName & " (" & (destRowIdx - 2) & " dong du lieu)"
    Next u

    ' ---- Luu file moi va mo len cho nguoi dung xem ----
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Xoa sheet mac dinh "Sheet1" neu ton tai va con sheet khac
    Dim wsDelete As Worksheet
    On Error Resume Next
    Set wsDelete = wbDest.Worksheets("Sheet1")
    On Error GoTo 0
    If Not wsDelete Is Nothing Then
        If wbDest.Worksheets.Count > 1 Then
            Application.DisplayAlerts = False
            wsDelete.Delete
            Application.DisplayAlerts = True
        End If
    End If

    newFilePath = sourceFolder & "SortedSheets_" & Format(Now(), "yyyymmdd_hhmmss") & ".xlsx"
    wbDest.SaveAs Filename:=newFilePath, FileFormat:=xlOpenXMLWorkbook

    ' Dong file nguon, giu file moi mo de nguoi dung xem
    Call SafeCloseWb(wbSource)
    wbDest.Activate
    wbDest.Worksheets(1).Activate

    MsgBox "Hoan thanh! Da tao " & createdCount & " sheet moi." & vbCrLf & _
           "File duoc luu tai: " & newFilePath, vbInformation, "Thanh cong"

End Sub

' ============================================================
' HELPER: Dong workbook an toan, bo qua loi neu da dong roi
' ============================================================
Private Sub SafeCloseWb(wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close False
    On Error GoTo 0
End Sub
