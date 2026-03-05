Option Explicit

Sub TaoSheet_TheoDivision_ChiLayData_XoaDongDau_And_Clean_WithFormatAndRename()
    Dim ws As Worksheet, wb As Workbook
    Dim pt As PivotTable, pf As PivotField
    Dim piLoop As PivotItem
    Dim itemName As String, wsMoi As Worksheet, rngCopy As Range
    Dim tenSheetGoc As String, tenSheetCuoi As String
    Dim laOLAP As Boolean
    
    '=== THAM SO ===
    Const TEN_SHEET_PIVOT As String = "PivotAR"
    Const TEN_PIVOT As String = ""                ' "" -> PivotTables(1)
    Const TEN_PAGE_FIELD As String = "Division"   ' Doi neu field ten khac
    Const DIA_CHI_O_B1 As String = "B1"
    Const GHI_DE_SHEET_CU As Boolean = True
    Const CHI_LAY_VUNG_TABLE_RANGE_1 As Boolean = False ' True = chi "data chinh"; False = toan bo (TableRange2)
    '===============
    
    On Error GoTo LOI
    
    ' 0) Lay sheet/pivot
    Set ws = Worksheets(TEN_SHEET_PIVOT)
    Set wb = ws.Parent
    
    If Len(TEN_PIVOT) > 0 Then
        Set pt = ws.PivotTables(TEN_PIVOT)
    Else
        If ws.PivotTables.Count = 0 Then Err.Raise vbObjectError + 1, , _
            "Khong tim thay PivotTable tren sheet " & TEN_SHEET_PIVOT
        Set pt = ws.PivotTables(1)
    End If
    
    Set pf = pt.PivotFields(TEN_PAGE_FIELD)
    If pf Is Nothing Then Err.Raise vbObjectError + 2, , "Khong tim thay PivotField '" & TEN_PAGE_FIELD & "'."
    If pf.Orientation <> xlPageField Then Err.Raise vbObjectError + 3, , _
        "Field '" & TEN_PAGE_FIELD & "' khong thuoc Report Filter (Page field)."
    
    laOLAP = pt.PivotCache.OLAP
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' >>> (I) XOA HET SHEET NGOAI TRU "PivotAR_NAME" VA "PivotAR", "Summary"
    XoaSheetKhongNamTrongDanhSach Array("PivotAR_NAME", "PivotAR", "Summary")
    
    ' 1) Bo filter cot E tren sheet PivotAR (neu co)
    Goi_BoFilterCotE ws
    
    ' 2) Non-OLAP -> dat Page field ve don lua chon
    If Not laOLAP Then
        If pf.EnableMultiplePageItems Then pf.EnableMultiplePageItems = False
    End If
    
    ' 3) Lap tung gia tri trong Division
    For Each piLoop In pf.PivotItems
        itemName = piLoop.Name
        If LCase$(itemName) = "(all)" Or LCase$(itemName) = "(blank)" Then GoTo SkipItem
        
        ' Hien thi de quan sat
        ws.Range(DIA_CHI_O_B1).Value = itemName
        
        ' Dat filter bang caption; neu OLAP khong ho tro -> bo qua
        On Error Resume Next
        pf.CurrentPage = itemName
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo LOI
            GoTo SkipItem
        End If
        On Error GoTo LOI
        
        ' KHONG refresh de tranh mo file nguon
        ' pt.RefreshTable
        
        ' Lay vung hien thi
        If CHI_LAY_VUNG_TABLE_RANGE_1 Then
            Set rngCopy = pt.TableRange1
        Else
            Set rngCopy = pt.TableRange2
        End If
        
        If Not rngCopy Is Nothing Then
            tenSheetGoc = DonTenSheet(itemName): If tenSheetGoc = "" Then tenSheetGoc = "KQ"
            
            ' Neu sheet ten goc da ton tai -> xoa hoac tao ten khong trung
            If SheetTonTaiTrongWB(wb, tenSheetGoc) Then
                If GHI_DE_SHEET_CU Then
                    wb.Worksheets(tenSheetGoc).Delete
                Else
                    tenSheetGoc = TaoTenSheet_KhongTrungTrongWB(wb, tenSheetGoc)
                End If
            End If
            
            ' Tao sheet moi trong chinh workbook
            Set wsMoi = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            wsMoi.Name = tenSheetGoc
            
            ' 4) Dan VALUES (chi data)
            wsMoi.Range("A1").Resize(rngCopy.Rows.Count, rngCopy.Columns.Count).Value = rngCopy.Value
            
            ' 5) Xoa dong 1-3
            wsMoi.Rows("1:3").Delete
            
            ' 6) Header + chen cot B
            wsMoi.Range("A1").Value = "Customer code"
            wsMoi.Columns("B:B").Insert Shift:=xlToRight
            wsMoi.Range("B1").Value = "Customer name"
            ' Header bo sung
            wsMoi.Range("C1").Value = "Doc date"
            wsMoi.Range("D1").Value = "Due date"
            
            ' 7) Dien cot B theo quy tac A la so -> B(r) = A(r+1)
            Goi_DienCotB_TheoCotA wsMoi
            
            ' 8) Xoa dong khong thoa dieu kien giu lai
            Goi_XoaDongKhongThoa wsMoi
            
            ' 8.5) Dam bao cac cot aging G:K va header dung thu tu
            Goi_DamBaoCotAging wsMoi
            
            ' 9) Dinh dang bo sung (dong 1 dam; dong A la so -> dam; to do cac cot aging; Grand Total -> vang)
            Goi_DinhDangBoSung wsMoi
            
            ' 9.5) Tính tong Grand Total (I + J)
            Goi_TinhTongGrandTotal_IJ wsMoi
            
            ' 10) Dinh dang so #,##0 cho cac cot F:K
            On Error Resume Next
            wsMoi.Range("F:K").NumberFormat = "#,##0"
            On Error GoTo 0
            
            ' 11) Doi ten sheet theo mapping
            tenSheetCuoi = TenSheetSauMapping(wsMoi.Name)
            If tenSheetCuoi <> wsMoi.Name Then
                tenSheetCuoi = TaoTenSheet_KhongTrungTrongWB(wb, tenSheetCuoi) ' dam bao khong trung
                wsMoi.Name = tenSheetCuoi
            End If
            
            ' (Tuy chon) Can cot
            wsMoi.UsedRange.Columns.AutoFit
        End If
        
SkipItem:
    Next piLoop
    
    ' >>> (M?I 2) Sau khi tao xong: loai bo "(1)" trong ten sheet neu co
    XoaChuoi_1_TrongTenSheet
    
    ' 12) Tra lai All
    On Error Resume Next
    pf.ClearAllFilters
    ws.Range(DIA_CHI_O_B1).Value = "(All)"
    On Error GoTo 0
    
    ' 13) tao sheet Pastdue
    Call Tao_Pastdue_Over60Days_1
    
    ' 14)
    Call Dien_Summary_GrandTotal
    
    ' 15)
    Dim arr As Variant
    arr = Array("FB", "IN", "QSR", "PP", "BASIC INDUSTRIES (HEAVY)", "MANUFACTURING (LIGHT)", "MINING", "Inter-Company",
    "Downstream", "High Tech")
    Call Tao_Sheet_CRITICAL(arr)

    ' 16)
    Call ToMau_Tab_Sheet("CRITICAL", 255, 0, 0)          ' tab do
    Call ToMau_Tab_Sheet("Pastdue over 60 days_1", 218, 165, 32)  ' tab xanh nuoc bien


    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Da tao xong", vbInformation
    Exit Sub

LOI:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Loi: " & Err.Description, vbExclamation
End Sub

'======================
'   CAC THU TUC PHU
'======================

' (M?I) Xoa tat ca sheet KHONG nam trong danh sach cho phep
Private Sub XoaSheetKhongNamTrongDanhSach(ByVal arrKeep As Variant)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim i As Long
    Dim dictKeep As Object: Set dictKeep = CreateObject("Scripting.Dictionary")
    dictKeep.CompareMode = vbTextCompare
    For i = LBound(arrKeep) To UBound(arrKeep)
        dictKeep(arrKeep(i)) = True
    Next i
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not dictKeep.Exists(ws.Name) Then
            ' Khong xoa neu workbook con 1 sheet (Excel khong cho xoa het)
            If wb.Worksheets.Count > 1 Then ws.Delete
        End If
    Next ws
End Sub

' (M?I) Doi ten sheet: neu chua "(1)" thi bo di; neu trung ten thi tao ten khong trung
Private Sub XoaChuoi_1_TrongTenSheet()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    Dim tenMoi As String
    
    For Each ws In wb.Worksheets
        tenMoi = Replace(ws.Name, "(1)", "")   ' bo tat ca "(1)" trong ten
        tenMoi = Trim$(tenMoi)
        If tenMoi = "" Then tenMoi = "Sheet"
        
        If tenMoi <> ws.Name Then
            tenMoi = TaoTenSheet_KhongTrungTrongWB(wb, tenMoi)
            ws.Name = tenMoi
        End If
    Next ws
End Sub

' Bo tat ca filter, dac biet cot E
Private Sub Goi_BoFilterCotE(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    If Not ws.AutoFilter Is Nothing Then ws.AutoFilter.ShowAllData
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If Not lo.AutoFilter Is Nothing Then lo.AutoFilter.ShowAllData
    Next lo
    On Error GoTo 0
End Sub

' Dien cot B: neu A toan ky tu so -> B(r) = A(r+1); nguoc lai de trong
Private Sub Goi_DienCotB_TheoCotA(ByVal ws As Worksheet)
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    For r = 2 To lastRow
        Dim s As String
        s = Trim$(CStr(ws.Cells(r, "A").Value))
        If IsAllDigits(s) Then
            ws.Cells(r, "B").Value = ws.Cells(r + 1, "A").Value
        Else
            ws.Cells(r, "B").ClearContents
        End If
    Next r
End Sub

' Xoa dong khong thoa: GIU dong co A la so hoac bat dau "C2" hoac = "Grand Total"
Private Sub Goi_XoaDongKhongThoa(ByVal ws As Worksheet)
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Application.ScreenUpdating = False
    For r = lastRow To 2 Step -1
        Dim s As String, sU As String
        s = Trim$(CStr(ws.Cells(r, "A").Value))
        sU = UCase$(s)
        
        Dim keepRow As Boolean
        keepRow = False
        
        If IsAllDigits(s) Then
            keepRow = True
        ElseIf Left$(sU, 2) = "C2" Then
            keepRow = True
        ElseIf sU = "GRAND TOTAL" Then
            keepRow = True
        End If
        
        If Not keepRow Then ws.Rows(r).Delete
    Next r
    Application.ScreenUpdating = True
End Sub

' DAM BAO CAC COT AGING o G:K theo thu tu bat buoc
' F1="_Current"; G1="1-30 days"; H1="31-60 days"; I1="61-90 days"; J1="90-180 days"; K1="Grand Total"
Private Sub Goi_DamBaoCotAging(ByVal ws As Worksheet)
    Dim headers As Variant, targets As Variant
    Dim i As Long, foundCol As Long
    
    headers = Array("_Current", "1-30 days", "31-60 days", "61-90 days", "90-180 days", "Grand Total")
    targets = Array(6, 7, 8, 9, 10, 11) ' F,G,H,I,J,K
    
    Application.CutCopyMode = False
    
    For i = LBound(headers) To UBound(headers)
        foundCol = TimCotTheoHeader(ws, headers(i))
        
        If foundCol > 0 Then
            If foundCol <> CLng(targets(i)) Then
                ws.Columns(foundCol).Cut
                ws.Columns(CLng(targets(i))).Insert Shift:=xlToRight
                Application.CutCopyMode = False
                ws.Cells(1, CLng(targets(i))).Value = headers(i)
            Else
                ws.Cells(1, CLng(targets(i))).Value = headers(i)
            End If
        Else
            ws.Columns(CLng(targets(i))).Insert Shift:=xlToRight
            ws.Cells(1, CLng(targets(i))).Value = headers(i)
        End If
    Next i
End Sub

' Tim cot theo header tren dong 1 (khong phan biet hoa thuong)
Private Function TimCotTheoHeader(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long, h As String
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = Trim$(CStr(ws.Cells(1, c).Value))
        If StrComp(h, headerText, vbTextCompare) = 0 Then
            TimCotTheoHeader = c
            Exit Function
        End If
    Next c
    TimCotTheoHeader = 0
End Function

' Dinh dang bo sung (dong 1 dam; dong A la so -> dam; to do cac cot aging; Grand Total -> vang)
Private Sub Goi_DinhDangBoSung(ByVal ws As Worksheet)
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    Dim head As String
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 1 Then Exit Sub
    
    ' Dong 1 to dam
    ws.Rows(1).Font.Bold = True
    
    ' Dong co A toan ky tu so -> to dam ca dong
    For r = 2 To lastRow
        Dim s As String
        s = Trim$(CStr(ws.Cells(r, "A").Value))
        If IsAllDigits(s) Then
            ws.Rows(r).Font.Bold = True
        End If
    Next r
    
    ' To do cac cot co header aging
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        head = UCase$(Trim$(CStr(ws.Cells(1, c).Value)))
        If head = "1-30 DAYS" Or head = "31-60 DAYS" Or head = "61-90 DAYS" Or head = "90-180 DAYS" Then
            ws.Range(ws.Cells(1, c), ws.Cells(lastRow, c)).Font.Color = vbRed
        End If
    Next c
    
    ' Hang co A chua "Grand Total" -> highlight tu cot A den cot K (khong qua K)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        If InStr(1, UCase$(Trim$(CStr(ws.Cells(r, "A").Value))), "GRAND TOTAL", vbTextCompare) > 0 Then
            ws.Range(ws.Cells(r, "A"), ws.Cells(r, "K")).Interior.Color = vbYellow
            ws.Rows(r).Font.Bold = True
        End If
    Next r

End Sub

' Mapping doi ten sheet
Private Function TenSheetSauMapping(ByVal tenHienTai As String) As String
    Dim s As String
    s = tenHienTai
    
    Select Case UCase$(s)
        Case "F&B", "F&amp;B"
            TenSheetSauMapping = "FB"
            Exit Function
        Case "INSTITUTIONAL"
            TenSheetSauMapping = "IN"
            Exit Function
        Case "INTERCO"
            TenSheetSauMapping = "Inter-Company"
            Exit Function
        Case "PAPER"
            TenSheetSauMapping = "PP"
            Exit Function
        Case Else
            ' Neu khong trung mapping cu the nao, VIET HOA toan bo
            TenSheetSauMapping = UCase$(s)
    End Select
End Function

' Kiem tra chuoi toan ky tu so
Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    s = Trim$(s)
    If Len(s) = 0 Then IsAllDigits = False: Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then IsAllDigits = False: Exit Function
    Next i
    IsAllDigits = True
End Function

' Helpers khac
Private Function SheetTonTaiTrongWB(ByVal wb As Workbook, ByVal ten As String) As Boolean
    Dim wsCheck As Worksheet
    On Error Resume Next
    Set wsCheck = wb.Worksheets(ten)
    SheetTonTaiTrongWB = Not wsCheck Is Nothing
    On Error GoTo 0
End Function

Private Function DonTenSheet(ByVal s As String) As String
    Dim x: x = Array("\", "/", "?", "*", "[", "]", ":", Chr(0))
    Dim ch
    For Each ch In x
        s = Replace$(s, CStr(ch), " ")
    Next
    If Len(s) > 31 Then s = Left$(s, 31)
    DonTenSheet = Trim$(s)
End Function

Private Function TaoTenSheet_KhongTrungTrongWB(ByVal wb As Workbook, ByVal goc As String) As String
    Dim i As Long, t As String: t = goc: i = 1
    Do While SheetTonTaiTrongWB(wb, t)
        t = DonTenSheet(goc & " (" & i & ")")
        i = i + 1
        If i > 1000 Then Exit Do
    Loop
    TaoTenSheet_KhongTrungTrongWB = t
End Function

'============================================
' TINH TONG GRAND TOTAL: COT I + COT J
' Ghi vao: COT I, dong (dongGrandTotal + 1)
'============================================
Private Sub Goi_TinhTongGrandTotal_IJ(ByVal ws As Worksheet)
    Dim lastRow As Long, r As Long
    Dim dongGrandTotal As Long
    Dim dataI As Double, dataJ As Double
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    dongGrandTotal = 0
    
    ' Tim dong chua "Grand Total" o cot A
    For r = 2 To lastRow
        If InStr(1, UCase$(Trim$(CStr(ws.Cells(r, "A").Value))), "GRAND TOTAL", vbTextCompare) > 0 Then
            dongGrandTotal = r
            Exit For
        End If
    Next r
    
    If dongGrandTotal = 0 Then Exit Sub   ' Khong tim thay
    
    ' Lay du lieu cot I va J tai dong Grand Total
    dataI = CDbl(Val(ws.Cells(dongGrandTotal, "I").Value))
    dataJ = CDbl(Val(ws.Cells(dongGrandTotal, "J").Value))
    
    ' Ghi tong xuong dong ke tiep (cot I)
    ws.Cells(dongGrandTotal + 1, "I").Value = dataI + dataJ
    ws.Cells(dongGrandTotal + 1, "I").NumberFormat = "#,##0"
    ws.Cells(dongGrandTotal + 1, "I").Font.Color = vbRed
    
End Sub

Public Sub Tao_Pastdue_Over60Days_1()
    Dim wb As Workbook
    Dim wsPD As Worksheet
    Dim wsSrc As Worksheet
    Dim shName As String
    Dim arrSBU As Variant
    Dim i As Long
    Dim sbu As String
    Dim lastRowI As Long
    Dim valI As Double
    Dim wsSummary As Worksheet
    
    Set wb = ThisWorkbook
    shName = "Pastdue over 60 days_1"
    
    ' Lay sheet Summary lam moc chen
    On Error Resume Next
    Set wsSummary = wb.Worksheets("Summary")
    On Error GoTo 0
    If wsSummary Is Nothing Then
        MsgBox "Khong tim thay sheet 'Summary' de chen sheet tong hop.", vbExclamation
        Exit Sub
    End If
    
    ' Xoa sheet cu neu ton tai
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(shName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Tao moi NGAY SAU sheet Summary
    Set wsPD = wb.Worksheets.Add(After:=wsSummary)
    wsPD.Name = shName
    
    ' Dua ve vi tri thu 4 (neu co it nhat 4 sheet)
    If wb.Worksheets.Count >= 4 Then
        If wsPD.Index <> 4 Then
            wsPD.Move Before:=wb.Worksheets(4)
        End If
    End If
    
    ' Header
    wsPD.Range("A1").Value = "SBU"
    wsPD.Range("B1").Value = "Current"
    wsPD.Rows(1).Font.Bold = True
    wsPD.Range("A1:B1").Interior.Color = RGB(0, 176, 240) ' xanh nuoc bien
    
    ' Danh sach SBU A2..A10
    arrSBU = Array( _
        "FB", _
        "IN", _
        "QSR", _
        "PP", _
        "BASIC INDUSTRIES (HEAVY)", _
        "MANUFACTURING (LIGHT)", _
        "MINING", _
        "Inter-Company", _
        "Downstream", _
        "High Tech", _
        "Total" _
    )
    For i = LBound(arrSBU) To UBound(arrSBU)
        wsPD.Cells(2 + i, "A").Value = CStr(arrSBU(i))
    Next i
    
    ' B2..B9: tu ten sheet o cot A -> lay gia tri cot I o dong cuoi co du lieu
    For i = 0 To 9   ' A2..A9 (bo qua "Total")
        sbu = CStr(arrSBU(i))
        
        Set wsSrc = Nothing
        On Error Resume Next
        Set wsSrc = wb.Worksheets(sbu)
        On Error GoTo 0
        
        If Not wsSrc Is Nothing Then
            lastRowI = wsSrc.Cells(wsSrc.Rows.Count, "I").End(xlUp).Row
            If lastRowI < 2 And Len(Trim$(CStr(wsSrc.Cells(lastRowI, "I").Value))) = 0 Then
                valI = 0
            Else
                If IsNumeric(wsSrc.Cells(lastRowI, "I").Value) Then
                    valI = CDbl(wsSrc.Cells(lastRowI, "I").Value)
                Else
                    valI = CDbl(Val(wsSrc.Cells(lastRowI, "I").Value))
                End If
            End If
        Else
            valI = 0
        End If
        
        ' Neu <= 0 -> de trong
        If valI <= 0 Then
            wsPD.Cells(2 + i, "B").ClearContents
        Else
            wsPD.Cells(2 + i, "B").Value = valI
        End If
    Next i
    
    ' B12 = SUM(B2:B11)
    wsPD.Range("B12").Formula = "=SUM(B2:B11)"
    
    ' Dinh dang so
    wsPD.Range("B2:B12").NumberFormat = "#,##0"
    
    ' Can cot
    wsPD.Columns("A:B").AutoFit
End Sub

' Xoa du lieu tren 1 sheet theo dia chi range truyen vao
' - sheetName: ten sheet dich (neu de rong -> dung ActiveSheet)
' - addr: dia chi vung/ o (vd: "C8:G15" hoac "C8" hoac nhieu vung "C8:C10,E8:E10")
' - clearFormats: True = xoa ca dinh dang (Clear); False = chi xoa noi dung (ClearContents)
Public Sub Xoa_Vung_Tu_ThamSo(ByVal sheetName As String, ByVal addr As String, Optional ByVal clearFormats As Boolean = False)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    
    Set wb = ThisWorkbook
    
    On Error GoTo LOI
    
    ' Chon sheet muc tieu
    If Len(Trim$(sheetName)) = 0 Then
        Set ws = ActiveSheet
    Else
        Set ws = wb.Worksheets(sheetName)
    End If
    
    ' Bo khoang trang thua
    addr = Trim$(addr)
    If Len(addr) = 0 Then
        Err.Raise vbObjectError + 1001, , "Dia chi range dang rong."
    End If
    
    ' Lay range theo dia chi truyen vao
    Set rng = ws.Range(addr)
    
    ' Xoa noi dung hoac ca dinh dang
    If clearFormats Then
        rng.Clear            ' xoa ca noi dung + dinh dang
    Else
        rng.ClearContents    ' chi xoa noi dung
    End If
    
    Exit Sub

LOI:
    MsgBox "Khong the xoa vung. Chi tiet: " & Err.Description, vbExclamation
End Sub

Public Sub Dien_Summary_GrandTotal()
    Dim wb As Workbook
    Dim wsSum As Worksheet
    Dim rFB As Long, rIN As Long, rInter As Long
    Dim rPP As Long, rQSR As Long, rMIN As Long
    Dim rLIGHT As Long, rHEAVY As Long
    Dim rDownstream As Long, rHighTech As Long
    
    Set wb = ThisWorkbook
    
    ' Tao sheet Summary neu chua co
    On Error Resume Next
    Set wsSum = wb.Worksheets("Summary")
    On Error GoTo 0
    If wsSum Is Nothing Then
        Set wsSum = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsSum.Name = "Summary"
    End If
    
    Call Xoa_Vung_Tu_ThamSo("Summary", "C8:G17")
    
    ' Tim dong "Grand Total" tren cac sheet
    rFB = TimDongGrandTotal("FB")
    rIN = TimDongGrandTotal("IN")
    rInter = TimDongGrandTotal("Inter-Company")
    rPP = TimDongGrandTotal("PP")
    rQSR = TimDongGrandTotal("QSR")
    rMIN = TimDongGrandTotal("MINING")
    rLIGHT = TimDongGrandTotal("MANUFACTURING (LIGHT)")
    rHEAVY = TimDongGrandTotal("BASIC INDUSTRIES (HEAVY)")
    rDownstream = TimDongGrandTotal("Downstream")
    rHighTech = TimDongGrandTotal("High Tech")
    
    ' === GHI CONG THUC LEN SHEET SUMMARY ===
    ' 2) Hang 8 - dung dong cua FB
    GhiHangSummary wsSum, 8, rFB, "FB"
    ' 3) Hang 9 - dung dong cua IN (tham chieu FB! cot G,H,I,J,F theo yeu cau)
    GhiHangSummary wsSum, 9, rIN, "IN"
    ' 4) Hang 10 - dung dong cua PP
    GhiHangSummary wsSum, 10, rQSR, "QSR"
    ' 5) Hang 11 - dung dong cua PP (lap lai theo yeu cau)
    GhiHangSummary wsSum, 11, rPP, "PP"
    ' 6) Hang 12 - dung dong cua HEAVY
    GhiHangSummary wsSum, 12, rHEAVY, "BASIC INDUSTRIES (HEAVY)"
    ' 7) Hang 13 - dung dong cua LIGHT
    GhiHangSummary wsSum, 13, rLIGHT, "MANUFACTURING (LIGHT)"
    ' 8) Hang 14 - dung dong cua MINING
    GhiHangSummary wsSum, 14, rMIN, "MINING"
    ' 9) Hang 15 - dung dong cua Inter-Company
    GhiHangSummary wsSum, 15, rInter, "Inter-Company"
    ' 10) Hang 16 - dung dong cua Downstream
    GhiHangSummary wsSum, 16, rDownstream, "Downstream"
    ' 11) Hang 17 - dung dong cua High Tech
    GhiHangSummary wsSum, 17, rHighTech, "High Tech"
    
    ' Dinh dang so (neu can)
    wsSum.Range("C8:G17").NumberFormat = "#,##0"
End Sub

'--- Helper: tim dong chua chuoi "Grand Total" o cot A cua sheetName ---
Private Function TimDongGrandTotal(ByVal sheetName As String) As Long
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    TimDongGrandTotal = 0
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 1 To lastRow
        If InStr(1, UCase$(Trim$(CStr(ws.Cells(r, "A").Value))), "GRAND TOTAL", vbTextCompare) > 0 Then
            TimDongGrandTotal = r
            Exit For
        End If
    Next r
End Function

'--- Helper: ghi cong thuc cho 1 hang tren Summary
' Theo yeu cau: C = FB!G, D = FB!H, E = FB!I, F = FB!J, G = FB!F voi so dong = rowNum
Private Sub GhiHangSummary(ByVal wsSum As Worksheet, ByVal targetRow As Long, ByVal rowNum As Long, ByVal sheetName As String)
    If rowNum > 0 Then
        wsSum.Cells(targetRow, "C").Formula = "='" & sheetName & "'!G" & rowNum
        wsSum.Cells(targetRow, "D").Formula = "='" & sheetName & "'!H" & rowNum
        wsSum.Cells(targetRow, "E").Formula = "='" & sheetName & "'!I" & rowNum
        wsSum.Cells(targetRow, "F").Formula = "='" & sheetName & "'!J" & rowNum
        wsSum.Cells(targetRow, "G").Formula = "='" & sheetName & "'!F" & rowNum
    Else
        ' Neu khong tim thay dong -> de trong
        wsSum.Range("C" & targetRow & ":G" & targetRow).ClearContents
    End If
End Sub

Public Sub Tao_Sheet_CRITICAL(ByVal wsMoi As Variant)
    Dim wb As Workbook
    Dim wsC As Worksheet, wsAfter As Worksheet
    Dim i As Long, r As Long, ci As Long
    Dim wsSrc As Worksheet
    Dim lastRow As Long, nextRow As Long
    Dim valB As String
    Dim v As Variant, num As Double, sumRow As Double
    Dim shName As String
    Dim arrDataCols As Variant
    Dim dataStartCol As Long, sumCol As Long, actionCol As Long
    Dim colNum As Long
    
    Set wb = ThisWorkbook
    
    ' === CAC COT NGUON CAN LAY (tuy chinh o day) ===
    arrDataCols = Array("I", "J")
    ' ===============================================
    
    ' Lay sheet moc chen
    On Error Resume Next
    Set wsAfter = wb.Worksheets("Pastdue over 60 days_1")
    On Error GoTo 0
    
    ' Xoa sheet CRITICAL neu ton tai
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("CRITICAL").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Tao CRITICAL sau sheet "Pastdue over 60 days_1" (neu co), nguoc lai tao o cuoi
    If Not wsAfter Is Nothing Then
        Set wsC = wb.Worksheets.Add(After:=wsAfter)
    Else
        Set wsC = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    End If
    wsC.Name = "CRITICAL"
    
    ' Header co dinh
    wsC.Cells(1, "A").Value = "Customer code"
    wsC.Cells(1, "B").Value = "Customer name"
    wsC.Cells(1, "C").Value = "Division"
    
    ' Header cac cot data tu arrDataCols (bat dau tu D)
    dataStartCol = 4 ' D
    For ci = LBound(arrDataCols) To UBound(arrDataCols)
        wsC.Cells(1, dataStartCol + ci).Value = HeaderChoCot(UCase$(CStr(arrDataCols(ci))))
    Next ci
    
    ' Cot tong "61-180 days" + cot "Action"
    sumCol = dataStartCol + (UBound(arrDataCols) - LBound(arrDataCols)) + 1
    wsC.Cells(1, sumCol).Value = "61-180 days"
    actionCol = sumCol + 1
    wsC.Cells(1, actionCol).Value = "Action"
    
    wsC.Rows(1).Font.Bold = True
    
    nextRow = 2
    
    ' --- Duyet danh sach sheet nguon ---
    If IsArray(wsMoi) Then
        For i = LBound(wsMoi) To UBound(wsMoi)
            Set wsSrc = Nothing
            If TypeName(wsMoi(i)) = "Worksheet" Then
                Set wsSrc = wsMoi(i)
            Else
                shName = CStr(wsMoi(i))
                On Error Resume Next
                Set wsSrc = wb.Worksheets(shName)
                On Error GoTo 0
            End If
            
            If Not wsSrc Is Nothing Then
                lastRow = MaxRow_MultiCols(wsSrc, arrDataCols)
                If lastRow >= 2 Then
                    For r = 2 To lastRow
                        valB = Trim$(CStr(wsSrc.Cells(r, "B").Value))
                        If Len(valB) > 0 Then
                            ' Ghi A,B,C
                            wsC.Cells(nextRow, "A").Value = wsSrc.Cells(r, "A").Value
                            wsC.Cells(nextRow, "B").Value = wsSrc.Cells(r, "B").Value
                            wsC.Cells(nextRow, "C").Value = wsSrc.Name
                            
                            ' Ghi tung cot data + tinh tong
                            sumRow = 0
                            For ci = LBound(arrDataCols) To UBound(arrDataCols)
                                colNum = ColNumFromLetter(CStr(arrDataCols(ci)))
                                v = wsSrc.Cells(r, colNum).Value
                                If IsNumeric(v) Then
                                    num = CDbl(v)
                                Else
                                    num = CDbl(Val(v))
                                End If
                                wsC.Cells(nextRow, dataStartCol + ci).Value = num
                                sumRow = sumRow + num
                            Next ci
                            
                            ' Ghi tong vao cot "61-180 days"
                            wsC.Cells(nextRow, sumCol).Value = sumRow
                            
                            ' Gan chu "Action" theo yeu cau
                            If Len(Trim$(CStr(wsC.Cells(nextRow, "A").Value))) > 0 Then
                                wsC.Cells(nextRow, actionCol).Value = "Sale team follow up with customer about the payment schedule"
                            End If
                            
                            nextRow = nextRow + 1
                        End If
                    Next r
                End If
            End If
        Next i
    Else
        ' Truong hop chi truyen 1 sheet/ten sheet
        Set wsSrc = Nothing
        If TypeName(wsMoi) = "Worksheet" Then
            Set wsSrc = wsMoi
        Else
            shName = CStr(wsMoi)
            On Error Resume Next
            Set wsSrc = wb.Worksheets(shName)
            On Error GoTo 0
        End If
        
        If Not wsSrc Is Nothing Then
            lastRow = MaxRow_MultiCols(wsSrc, arrDataCols)
            If lastRow >= 2 Then
                For r = 2 To lastRow
                    valB = Trim$(CStr(wsSrc.Cells(r, "B").Value))
                    If Len(valB) > 0 Then
                        wsC.Cells(nextRow, "A").Value = wsSrc.Cells(r, "A").Value
                        wsC.Cells(nextRow, "B").Value = wsSrc.Cells(r, "B").Value
                        wsC.Cells(nextRow, "C").Value = wsSrc.Name
                        
                        sumRow = 0
                        For ci = LBound(arrDataCols) To UBound(arrDataCols)
                            colNum = ColNumFromLetter(CStr(arrDataCols(ci)))
                            v = wsSrc.Cells(r, colNum).Value
                            If IsNumeric(v) Then
                                num = CDbl(v)
                            Else
                                num = CDbl(Val(v))
                            End If
                            wsC.Cells(nextRow, dataStartCol + ci).Value = num
                            sumRow = sumRow + num
                        Next ci
                        wsC.Cells(nextRow, sumCol).Value = sumRow
                        
                        If Len(Trim$(CStr(wsC.Cells(nextRow, "A").Value))) > 0 Then
                            wsC.Cells(nextRow, actionCol).Value = "Sale team follow up with customer about the payment schedule"
                        End If
                        
                        nextRow = nextRow + 1
                    End If
                Next r
            End If
        End If
    End If
    
    ' Dinh dang so cho cac cot data + cot tong
    If nextRow > 2 Then
        wsC.Range(wsC.Cells(2, dataStartCol), wsC.Cells(nextRow - 1, sumCol)).NumberFormat = "#,##0"
    End If
    
    ' Xoa cac dong co tong <= 0 o cot "61-180 days"
    Dim rDel As Long
    For rDel = nextRow - 1 To 2 Step -1
        If Val(wsC.Cells(rDel, sumCol).Value) <= 0 Then
            wsC.Rows(rDel).Delete
        End If
    Next rDel
    
    ' Sau khi xoa, dam bao cot Action dien dung tren moi dong co A co data
    Dim lastOutRow As Long
    lastOutRow = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).Row
    If lastOutRow >= 2 Then
        For r = 2 To lastOutRow
            If Len(Trim$(CStr(wsC.Cells(r, "A").Value))) > 0 Then
                wsC.Cells(r, actionCol).Value = "Sale team follow up with customer about the payment schedule"
            End If
        Next r
    End If
    
    ' Can cot
    wsC.Columns("A:" & ColLetterFromNumber(actionCol)).AutoFit
End Sub

'--- Helper: mapping header cho cot du lieu ---
Private Function HeaderChoCot(ByVal colLetter As String) As String
    Select Case UCase$(Trim$(colLetter))
        Case "G": HeaderChoCot = "1-30 days"
        Case "H": HeaderChoCot = "31-60 days"
        Case "I": HeaderChoCot = "61-90 days"
        Case "J": HeaderChoCot = "90-180 days"
        Case "K": HeaderChoCot = "Grand Total"
        Case Else: HeaderChoCot = colLetter
    End Select
End Function

'--- Helper: tra ve dong cuoi du tren cot B va cac cot data ---
Private Function MaxRow_MultiCols(ByVal ws As Worksheet, ByVal arrCols As Variant) As Long
    Dim rMax As Long, c As Variant, rTmp As Long
    rMax = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For Each c In arrCols
        rTmp = ws.Cells(ws.Rows.Count, CStr(c)).End(xlUp).Row
        If rTmp > rMax Then rMax = rTmp
    Next c
    MaxRow_MultiCols = rMax
End Function

'--- Helper: chu cai cot -> so cot ---
Private Function ColNumFromLetter(ByVal colLetter As String) As Long
    Dim i As Long, res As Long
    colLetter = UCase$(Trim$(colLetter))
    For i = 1 To Len(colLetter)
        res = res * 26 + (Asc(Mid$(colLetter, i, 1)) - Asc("A") + 1)
    Next i
    ColNumFromLetter = res
End Function

'--- Helper: so cot -> chu cai cot ---
Private Function ColLetterFromNumber(ByVal colNum As Long) As String
    Dim s As String, n As Long
    n = colNum
    Do While n > 0
        s = Chr$(((n - 1) Mod 26) + 65) & s
        n = (n - 1) \ 26
    Loop
    ColLetterFromNumber = s
End Function

' To mau tab cua 1 sheet theo mau RGB
Public Sub ToMau_Tab_Sheet(ByVal sheetName As String, ByVal r As Long, ByVal g As Long, ByVal b As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Khong tim thay sheet: " & sheetName, vbExclamation
        Exit Sub
    End If
    ws.Tab.Color = RGB(r, g, b)
End Sub

' Bo mau tab (tro ve mac dinh)
Public Sub BoMau_Tab_Sheet(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ws.Tab.ColorIndex = xlColorIndexNone
End Sub
