Attribute VB_Name = "DoiSoat"
'==========================================================================
' TOOL DOI SOAT NCC <-> MB51 (GR) - BAN VBA CHO MICROSOFT EXCEL
' Phat trien boi Nguyen Viet Duong
'
' Nut tren sheet Control:
'   CHECK CHENH LECH   -> CheckChenhLech
'   Chon file NCC/MB51 -> ChonFileNCC / ChonFileMB51
'   Luu ket qua        -> LuuKetQua
' Dung lai giao dien: Alt+F8 -> CaiDatGiaoDien -> Run
'
' File nguon nay viet tieng Viet UTF-8; script build doi chuoi co dau thanh U("\uXXXX")
' truoc khi nap vao Excel (VBA chi doc file .bas theo bang ma ANSI).
'==========================================================================
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Public gSilent As Boolean          ' chay thu tu dong: thong bao ghi vao o Z1, khong mo hop thoai
Private Const TRACE_ON As Boolean = False   ' ban test bat True: ghi vet tung buoc vao %TEMP%\doisoat_trace.txt

Private Const R_PATH As Long = 4, R_SHEET As Long = 5, R_START As Long = 6, R_MAT As Long = 7, R_PO As Long = 8
Private Const R_BB As Long = 9, R_SEQ As Long = 10, R_QTY As Long = 11, R_CAR As Long = 12, R_BATCH As Long = 13
Private Const R_CODES As Long = 14, R_TRIP As Long = 15, R_NOTE As Long = 17
Private Const C_NCC As Long = 2, C_MB As Long = 3, MAX_DETAIL As Long = 30, SUM_NC As Long = 11
Private Const PASS As String = ""
' vi tri truong trong 1 dong du lieu
Private Const F_R As Long = 0, F_MAT As Long = 1, F_PO As Long = 2, F_BB As Long = 3, F_QTY As Long = 4, F_CAR As Long = 5

Private CTRL As String, SUM_NAME As String, WARN_NAME As String, DATA_N As String, DATA_M As String
Private DETAIL_TAG As String, AUTO As String, NO_CAR As String, COL_LIST As String
Private MODELS As Variant, KW As Object, DEFV As Object
Private CLR_GREEN_BG As Long, CLR_GREEN_FT As Long, CLR_RED_BG As Long, CLR_RED_FT As Long
Private CLR_WARN_BG As Long, CLR_WARN_FT As Long, CLR_HDR_BG As Long, CLR_HDR2_BG As Long
Private CLR_WHITE As Long, CLR_YELLOW As Long, CLR_LBLUE As Long, CLR_GRID As Long
Private gCarOf As Object           ' dung cho ham so sanh khi sap xep PO theo dong xe

Private Sub InitK()
    CTRL = "Control": SUM_NAME = U("T\u1ed5ng h\u1ee3p"): WARN_NAME = U("C\u1ea3nh b\u00e1o")
    DATA_N = "NCC_Data": DATA_M = "MB51_Data"
    DETAIL_TAG = U("\u0110\u1ed0I SO\u00c1T CHI TI\u1ebeT"): AUTO = U("T\u1ef1 \u0111\u1ed9ng"): NO_CAR = U("(ch\u01b0a g\u00e1n)")
    MODELS = Array("VF5", "VF6", "VF7", "VF7 Pro", "VF8", "VF9", "Limo")
    COL_LIST = ColList()
    Set DEFV = NewDict()
    DEFV(R_SHEET) = Array("Sheet1", "Sheet1"): DEFV(R_START) = Array(2, 2)
    DEFV(R_MAT) = Array(AUTO, AUTO): DEFV(R_PO) = Array(AUTO, AUTO): DEFV(R_BB) = Array(AUTO, AUTO)
    DEFV(R_SEQ) = Array("-", "-"): DEFV(R_QTY) = Array(AUTO, AUTO): DEFV(R_CAR) = Array(AUTO, AUTO)
    DEFV(R_BATCH) = Array(5, 5): DEFV(R_CODES) = Array(4, 4): DEFV(R_TRIP) = Array(4, 4)
    ' Tu khoa nhan cot theo tieu de: Array(nhan, khoa khop, tu loai tru)
    Set KW = NewDict()
    KW("mat") = Array("Material", Split(U("material|m\u00e3 material|ma material|m\u00e3 v\u1eadt t\u01b0|ma vat tu|v\u1eadt t\u01b0|vat tu|m\u00e3 h\u00e0ng|ma hang|matnr|part no|part number|m\u00e3 linh ki\u1ec7n|ma linh kien"), "|"), _
                      Split(U("desc|document|doc.|m\u00f4 t\u1ea3|mo ta|t\u00ean|group|type|nh\u00f3m"), "|"))
    KW("po") = Array("PO/SA", Split(U("po/sa|sa/po|po|sa|purchase order|purchasing document|purchasing doc|pur. doc|po number|s\u1ed1 po|so po|\u0111\u01a1n h\u00e0ng|don hang|scheduling agreement"), "|"), _
                     Split(U("item|date|ng\u00e0y|ngay|type"), "|"))
    KW("bb") = Array(U("S\u1ed1 BB"), Split(U("s\u1ed1 bb|so bb|bb|bi\u00ean b\u1ea3n|bien ban|s\u1ed1 bi\u00ean b\u1ea3n|so bien ban|reference|ref|ref.|delivery note"), "|"), _
                     Split(U("date|ng\u00e0y|ngay"), "|"))
    KW("qty") = Array(U("S\u1ed1 l\u01b0\u1ee3ng"), Split(U("s\u1ed1 l\u01b0\u1ee3ng|so luong|sl|quantity|qty|qty in une|quantity in une|qty in un. of entry"), "|"), _
                      Split(U("unit|\u0111\u01a1n v\u1ecb|don vi|uom|price|amount|gi\u00e1"), "|"))
    KW("car") = Array(U("D\u00f2ng xe"), Split(U("d\u00f2ng xe|dong xe|lo\u1ea1i xe|loai xe|model|car model|vehicle|xe|d\u1ef1 \u00e1n|du an|project"), "|"), _
                      Split(U("date|ng\u00e0y|ngay"), "|"))
    CLR_GREEN_BG = RGB(198, 239, 206): CLR_GREEN_FT = RGB(0, 97, 0)
    CLR_RED_BG = RGB(255, 199, 206): CLR_RED_FT = RGB(156, 0, 6)
    CLR_WARN_BG = RGB(255, 235, 156): CLR_WARN_FT = RGB(156, 87, 0)
    CLR_HDR_BG = RGB(0, 32, 96): CLR_HDR2_BG = RGB(68, 114, 196)
    CLR_WHITE = RGB(255, 255, 255): CLR_YELLOW = RGB(255, 242, 204): CLR_LBLUE = RGB(221, 235, 247)
    CLR_GRID = RGB(200, 200, 200)
End Sub

' Giai ma chuoi "\uXXXX" (script build tao ra) thanh ky tu Unicode
Private Function U(ByVal s As String) As String
    Dim i As Long, n As Long, out As String
    n = Len(s): i = 1
    Do While i <= n
        If Mid$(s, i, 2) = "\u" And i + 5 <= n Then
            out = out & ChrW(CLng("&H" & Mid$(s, i + 2, 4)))
            i = i + 6
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = out
End Function

Private Function ColList() As String
    Dim a As String, i As Long
    For i = 0 To 25: a = a & Chr(65 + i) & ",": Next
    For i = 0 To 25: a = a & "A" & Chr(65 + i) & ",": Next
    ColList = Left$(a, Len(a) - 1)
End Function

'========================= CAC NUT BAM =========================
Public Sub CheckChenhLech()
    Dim calc As Long, em As String
    InitK
    Application.ScreenUpdating = False
    Trace "CheckChenhLech start"
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    calc = Application.Calculation
    Application.Calculation = xlCalculationManual
    On Error GoTo EH
    RunCheck
    GoTo Done
EH:
    em = Err.Description
    Trace "ERROR: " & Err.Number & " " & em
    On Error Resume Next
    Application.Calculation = calc
    Application.ScreenUpdating = True
    Msg U("L\u1ed7i: ") & em
Done:
    On Error Resume Next
    Application.Calculation = calc
    Dim c As Worksheet: Set c = GetSheet(ThisWorkbook, CTRL)
    If Not c Is Nothing Then ProtectSheet c
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Public Sub ChonFileNCC()
    InitK: ChonFile C_NCC, "Chon file NCC"
End Sub

Public Sub ChonFileMB51()
    InitK: ChonFile C_MB, "Chon file MB51"
End Sub

Private Sub ChonFile(ByVal c As Long, ByVal title As String)
    Dim ws As Worksheet, f As String
    Set ws = EnsureControl(ThisWorkbook)
    f = PickFile(title)
    If f <> "" Then ws.Cells(R_PATH, c).Value2 = f
    ProtectSheet ws
End Sub

Public Sub LuuKetQua()
    Dim def As String, path As String, f As Variant
    InitK
    If GetSheet(ThisWorkbook, SUM_NAME) Is Nothing Then Msg U("Ch\u01b0a c\u00f3 k\u1ebft qu\u1ea3. H\u00e3y b\u1ea5m CHECK CH\u00caNH L\u1ec6CH tr\u01b0\u1edbc."): Exit Sub
    def = ThisWorkbook.path & Application.PathSeparator & "DoiSoat_" & Format(Now, "yyyy-mm-dd_hhnn") & ".xlsx"
    If gSilent Then Exit Sub
    f = Application.GetSaveAsFilename(InitialFileName:=def, FileFilter:="Excel (*.xlsx),*.xlsx", Title:="Luu ket qua doi soat")
    If VarType(f) = vbBoolean Then Exit Sub
    path = CStr(f)
    On Error GoTo EH
    ExportResult path
    Msg U("\u0110\u00e3 l\u01b0u k\u1ebft qu\u1ea3:") & vbLf & path
    Exit Sub
EH:
    Msg U("L\u1ed7i khi l\u01b0u: ") & Err.Description
End Sub

Public Sub CaiDatGiaoDien()
    Dim ws As Worksheet
    InitK
    Set ws = EnsureControl(ThisWorkbook)
    ProtectSheet ws
    ws.Activate
End Sub

' Dung khi chay thu tu dong (khong mo hop thoai)
Public Sub SetTestMode(ByVal on_ As Boolean)
    gSilent = on_
End Sub

Public Sub XuatKetQua(ByVal path As String)
    InitK
    ExportResult path
End Sub

' Chay thu tron goi trong 1 lan goi (dung cho test tu dong): dat duong dan, doi soat, (xuat file), luu ban sao ket qua
Public Sub ChayThu(ByVal nccPath As String, ByVal mbPath As String, ByVal savePath As String, ByVal exportPath As String)
    Dim c As Worksheet
    gSilent = True
    InitK
    Set c = EnsureControl(ThisWorkbook)
    c.Range("Z1").Value2 = ""
    c.Cells(R_PATH, C_NCC).Value2 = nccPath: c.Cells(R_PATH, C_MB).Value2 = mbPath
    ProtectSheet c
    CheckChenhLech
    If exportPath <> "" Then
        On Error Resume Next
        ExportResult exportPath
        If Err.Number <> 0 Then Msg "EXPORT ERR: " & Err.Description
        On Error GoTo 0
    End If
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs savePath
End Sub

'========================= DOI SOAT =========================
Private Sub RunCheck()
    Dim t0 As Double, ws As Worksheet, nPath As String, mPath As String, f As String
    t0 = Timer
    Set ws = EnsureControl(ThisWorkbook)
    Trace "ensureControl ok"
    nPath = TrimS(ws.Cells(R_PATH, C_NCC).Value2)
    mPath = TrimS(ws.Cells(R_PATH, C_MB).Value2)
    If nPath = "" Then
        f = PickFile("Chon file NCC")
        If f <> "" Then ws.Cells(R_PATH, C_NCC).Value2 = f: nPath = f
    End If
    If mPath = "" Then
        f = PickFile("Chon file MB51")
        If f <> "" Then ws.Cells(R_PATH, C_MB).Value2 = f: mPath = f
    End If
    If nPath = "" Or mPath = "" Then
        ws.Activate
        Msg U("Ch\u01b0a c\u00f3 \u0111\u01b0\u1eddng d\u1eabn file.") & vbLf & vbLf & U("B\u1ea5m 'Ch\u1ecdn file NCC' v\u00e0 'Ch\u1ecdn file MB51' (ho\u1eb7c d\u00e1n \u0111\u01b0\u1eddng d\u1eabn v\u00e0o 2 \u00f4 v\u00e0ng), r\u1ed3i b\u1ea5m CHECK CH\u00caNH L\u1ec6CH.")
        Exit Sub
    End If
    Dim batch As Long, cpc As Long, cpt As Long
    batch = CLng(ToNum(ws.Cells(R_BATCH, C_NCC).Value2)): If batch < 1 Then batch = 5
    cpc = Int(ToNum(ws.Cells(R_CODES, C_NCC).Value2)): If cpc < 0 Then cpc = 0
    cpt = Int(ToNum(ws.Cells(R_TRIP, C_NCC).Value2)): If cpt < 0 Then cpt = 0
    Trace "params ok: " & nPath & " | " & mPath

    ' Doc 2 file (ca vung 1 lan) + nhan cot + loc dong loi
    Dim warns As New Collection, infoN As String, infoM As String, carN As Boolean, carM As Boolean
    Dim rawN As Variant, rawM As Variant, N As Variant, M As Variant, nN As Long, nM As Long
    rawN = ReadSource(ws, C_NCC, nPath, "NCC", infoN, carN)
    Trace "read NCC ok"
    rawM = ReadSource(ws, C_MB, mPath, "MB51", infoM, carM)
    Trace "read MB51 ok"
    N = CleanRows(rawN, "NCC", warns, nN)
    M = CleanRows(rawM, "MB51", warns, nM)
    Trace "clean ok nN=" & nN & " nM=" & nM
    If nN = 0 Then Msg U("File NCC kh\u00f4ng c\u00f3 d\u1eef li\u1ec7u h\u1ee3p l\u1ec7.") & vbLf & nPath: Exit Sub
    If nM = 0 Then Msg U("File MB51 kh\u00f4ng c\u00f3 d\u1eef li\u1ec7u h\u1ee3p l\u1ec7.") & vbLf & mPath: Exit Sub

    ' Gom theo SA/PO -> tung material ben trong PO
    Dim byPO As Object: Set byPO = NewDict()
    AddRows byPO, N, nN, True
    AddRows byPO, M, nM, False
    Trace "addRows ok, PO=" & byPO.Count
    Dim carOf As Object, hasCar As Boolean
    Set carOf = CarOfPO(N, nN, M, nM): hasCar = carN Or carM
    Dim pos As Variant: pos = byPO.Keys
    Set gCarOf = carOf
    SortArr pos, 1
    Trace "sort ok"

    ClearResultSheets ThisWorkbook
    Trace "clear ok"

    ' Kiem tra bo cau + sheet chi tiet cho PO co van de
    Dim sheetOf As Object, boOf As Object, o As Object, bo As Object
    Dim i As Long, j As Long, po As String, mat As Variant, mats As Variant, bad As Boolean
    Dim nBo As Long, nSheet As Long, skipped As Long
    Dim carS As String, mdName As String
    Set sheetOf = NewDict(): Set boOf = NewDict()
    For i = 0 To UBound(pos)
        po = pos(i): Set o = byPO(po)
        mats = o("mats").Keys: SortArr mats, 0: o("matsArr") = mats
        bad = False
        For j = 0 To UBound(mats)
            mat = mats(j)
            If Not o("inN").Exists(mat) Or Not o("inM").Exists(mat) Then
                bad = True: Exit For
            ElseIf Round6(DGet(o("n"), mat) - DGet(o("m"), mat)) <> 0 Then
                bad = True: Exit For
            End If
        Next
        Set bo = CheckBo(o, cpc, cpt)
        If i Mod 50 = 0 Then Trace "PO " & i & "/" & UBound(pos)
        boOf.Add po, bo
        If Not bo("ok") Then bad = True: nBo = nBo + 1
        If bad Then
            If nSheet >= MAX_DETAIL Then
                skipped = skipped + 1
            Else
                carS = DGetS(carOf, po)
                mdName = MakeDetailSheet(ThisWorkbook, po, o, batch, carS, bo)
                sheetOf(po) = mdName
                nSheet = nSheet + 1
            End If
        End If
    Next

    WriteDataSheet ThisWorkbook, DATA_N, N, nN, "NCC"
    WriteDataSheet ThisWorkbook, DATA_M, M, nM, "MB51"
    Trace "data sheets ok"
    Dim colNote As String, res As Object, sumWs As Worksheet, carArg As Object
    colNote = U("C\u1ed9t NCC: ") & infoN & U("   |   C\u1ed9t MB51: ") & infoM
    If hasCar Then Set carArg = carOf
    Set res = WriteSummary(ThisWorkbook, ws, N, nN, M, nM, sheetOf, warns.Count, carArg, pos, boOf, cpc, colNote)
    Set sumWs = res("ws")
    Trace "summary ok"
    If warns.Count > 0 Then WriteWarnings ThisWorkbook, sumWs, warns
    Application.Calculate

    ' Thong bao gon
    Dim s As String, secs As Double
    secs = Round(Timer - t0, 1)
    s = U("\u0110\u1ed1i so\u00e1t xong ") & (UBound(pos) + 1) & " SA/PO  (" & Format(secs, "0.0") & U(" gi\u00e2y)") & vbLf & vbLf & _
        U("\u2714 Kh\u1edbp: ") & res("khop") & U("     \u2718 L\u1ec7ch: ") & res("lech") & U("     \u26a0 M\u1ed9t b\u00ean: ") & res("motBen")
    If cpc > 0 Then
        s = s & vbLf & U("B\u1ed9 c\u1ea7u: ") & IIf(nBo > 0, nBo & U(" PO l\u1ec7ch b\u1ed9 / chuy\u1ebfn l\u1ed7i"), U("t\u1ea5t c\u1ea3 \u0111\u1ee7 b\u1ed9")) & "   (" & res("xeNCC") & U(" xe \u0111\u00e3 giao)")
        If nBo * 2 > UBound(pos) + 1 Then s = s & vbLf & U("\u2192 Qu\u00e1 n\u1eeda PO l\u1ed7i b\u1ed9: ki\u1ec3m tra 'S\u1ed1 m\u00e3 c\u1ea7u / 1 xe' (\u0111\u1eb7t 0 n\u1ebfu d\u1eef li\u1ec7u kh\u00f4ng theo b\u1ed9).")
    End If
    If warns.Count > 0 Then s = s & vbLf & U("\u26a0 ") & warns.Count & U(" d\u00f2ng d\u1eef li\u1ec7u l\u1ed7i (sheet ") & WARN_NAME & ")"
    If skipped > 0 Then s = s & vbLf & U("(Ch\u1ec9 t\u1ea1o ") & MAX_DETAIL & U(" sheet chi ti\u1ebft, ") & skipped & U(" PO c\u00f2n l\u1ea1i xem \u1edf T\u1ed5ng h\u1ee3p)")
    sumWs.Activate
    sumWs.Range("A1").Select
    Application.ScreenUpdating = True
    Msg s
End Sub

' Can tham chieu "Microsoft Scripting Runtime" (script build tu them) - New nhanh hon CreateObject rat nhieu
Private Function NewDict() As Object
    Set NewDict = New Scripting.Dictionary
End Function

Private Function DGet(d As Object, k As Variant) As Double
    If d.Exists(k) Then DGet = d(k) Else DGet = 0
End Function

Private Function DGetS(d As Object, k As Variant) As String
    If d Is Nothing Then Exit Function
    If d.Exists(k) Then DGetS = d(k)
End Function

Private Sub AddRows(byPO As Object, rows As Variant, ByVal n As Long, ByVal isN As Boolean)
    Dim i As Long, r As Variant, o As Object, po As String, lastPo As String, mat As String
    Dim dMats As Scripting.Dictionary, dQty As Scripting.Dictionary, dIn As Scripting.Dictionary, cRows As Collection
    lastPo = vbNullChar
    For i = 0 To n - 1
        r = rows(i): po = r(F_PO): mat = r(F_MAT)
        If po <> lastPo Then   ' du lieu thuong xep lien theo PO -> chi tra Dictionary khi doi PO
            If Not byPO.Exists(po) Then
                Set o = NewDict()
                o.Add "n", NewDict(): o.Add "m", NewDict(): o.Add "inN", NewDict(): o.Add "inM", NewDict()
                o.Add "mats", NewDict(): o.Add "rowsN", New Collection: o.Add "rowsM", New Collection
                byPO.Add po, o
            Else
                Set o = byPO(po)
            End If
            Set dMats = o("mats")
            If isN Then
                Set dQty = o("n"): Set dIn = o("inN"): Set cRows = o("rowsN")
            Else
                Set dQty = o("m"): Set dIn = o("inM"): Set cRows = o("rowsM")
            End If
            lastPo = po
        End If
        If Not dMats.Exists(mat) Then dMats.Add mat, 1
        If dQty.Exists(mat) Then dQty(mat) = dQty(mat) + r(F_QTY) Else dQty.Add mat, r(F_QTY)
        dIn(mat) = 1
        cRows.Add r
    Next
End Sub

' Kiem tra bo cau 1 PO: 1 xe = cpc ma (moi ma 1 cai); 1 chuyen (1 BB) = cpt xe
Private Function CheckBo(o As Object, ByVal cpc As Long, ByVal cpt As Long) As Object
    ' Bien khai bao dung kieu Scripting.Dictionary / Collection (goi truc tiep, nhanh hon nhieu so voi As Object)
    Dim res As Scripting.Dictionary, dN As Scripting.Dictionary, dM As Scripting.Dictionary
    Dim inN As Scripting.Dictionary, inM As Scripting.Dictionary, trips As Collection
    Dim mats As Variant, i As Long, mat As Variant, q As Double, nTripBad As Long
    Dim nN As Long, nM As Long, minN As Double, maxN As Double, minM As Double, issues As String
    Set res = New Scripting.Dictionary
    Set dN = o("n"): Set dM = o("m"): Set inN = o("inN"): Set inM = o("inM")
    Set trips = New Collection
    res("ok") = True: res("nCodes") = 0: res("carsN") = "": res("carsM") = "": res("issues") = ""
    res("cpc") = cpc: res("cpt") = cpt: res("nTripBad") = 0
    res.Add "trips", trips
    mats = o("matsArr"): minN = -1: minM = -1
    For i = 0 To UBound(mats)
        mat = mats(i)
        If inN.Exists(mat) Then
            nN = nN + 1: q = dN(mat)
            If minN < 0 Or q < minN Then minN = q
            If q > maxN Then maxN = q
        End If
        If inM.Exists(mat) Then
            nM = nM + 1: q = dM(mat)
            If minM < 0 Or q < minM Then minM = q
        End If
    Next
    res("nCodes") = nN
    If cpc <= 0 Then Set CheckBo = res: Exit Function   ' tat kiem tra bo
    ' So xe = so bo du = SL nho nhat trong cac ma (thieu ma -> 0 xe)
    If nN >= cpc Then res("carsN") = IIf(minN > 0, minN, 0) Else res("carsN") = 0
    If nM >= cpc Then res("carsM") = IIf(minM > 0, minM, 0) Else res("carsM") = 0
    If nN > 0 And nN <> cpc Then
        issues = U("Thi\u1ebfu m\u00e3 ") & nN & "/" & cpc
    ElseIf nN > 0 And Round6(maxN - minN) <> 0 Then
        issues = U("L\u1ec7ch b\u1ed9 (") & minN & U("\u2013") & maxN & U(" c\u00e1i/m\u00e3)")
    End If

    ' Tung chuyen (BB) cua NCC: du ma? cac ma bang nhau? tong = cpc x cpt?
    Dim byBB As Scripting.Dictionary, tTot As Scripting.Dictionary, tq As Scripting.Dictionary, t As Scripting.Dictionary
    Dim rowsN As Collection, r As Variant, bb As Variant, bbs As Variant, k As Variant, tot As Double
    Dim cnt As Long, mn As Double, mx As Double, notes As String, need As Double
    Set byBB = New Scripting.Dictionary: Set tTot = New Scripting.Dictionary: need = cpc * cpt
    Set rowsN = o("rowsN")
    For Each r In rowsN
        bb = r(F_BB)
        If byBB.Exists(bb) Then
            Set tq = byBB(bb)
        Else
            Set tq = New Scripting.Dictionary: byBB.Add bb, tq: tTot.Add bb, 0
        End If
        If tq.Exists(r(F_MAT)) Then tq(r(F_MAT)) = tq(r(F_MAT)) + r(F_QTY) Else tq.Add r(F_MAT), r(F_QTY)
        tTot(bb) = tTot(bb) + r(F_QTY)
    Next
    If byBB.Count > 0 Then
        bbs = byBB.Keys: SortArr bbs, 0
        For i = 0 To UBound(bbs)
            Set tq = byBB(bbs(i)): tot = tTot(bbs(i)): cnt = 0: mn = -1: mx = 0: notes = ""
            For Each k In tq.Keys
                cnt = cnt + 1: q = tq(k)
                If mn < 0 Or q < mn Then mn = q
                If q > mx Then mx = q
            Next
            If cnt <> cpc Then
                notes = U("c\u00f3 ") & cnt & "/" & cpc & U(" m\u00e3")
            ElseIf Round6(mx - mn) <> 0 Then
                notes = U("c\u00e1c m\u00e3 kh\u00f4ng b\u1eb1ng nhau (") & mn & U("\u2013") & mx & ")"
            End If
            If cpt > 0 And Round6(tot - need) <> 0 Then notes = JoinNote(notes, U("t\u1ed5ng ") & tot & "/" & need & U(" c\u00e1i"))
            If notes <> "" Then nTripBad = nTripBad + 1
            Set t = New Scripting.Dictionary
            t("bb") = bbs(i): t.Add "qty", tq: t("total") = tot: t("note") = notes
            trips.Add t
        Next
    End If
    res("nTripBad") = nTripBad
    If nTripBad > 0 Then issues = JoinNote(issues, U("Chuy\u1ebfn l\u1ed7i: ") & nTripBad)
    res("issues") = issues
    res("ok") = (issues = "")
    Set CheckBo = res
End Function

Private Function JoinNote(ByVal a As String, ByVal b As String) As String
    If a = "" Then JoinNote = b Else JoinNote = a & "; " & b
End Function

' Doc file nguon theo thong so tren Control -> mang dong tho Array(r, mat, po, bb, qty, car)
Private Function ReadSource(ws As Worksheet, ByVal c As Long, ByVal path As String, ByVal label As String, _
                            ByRef colInfo As String, ByRef hasCar As Boolean) As Variant
    Dim sheetName As String, startRow As Long, hdrRow As Long
    sheetName = TrimS(ws.Cells(R_SHEET, c).Value2)
    startRow = CLng(ToNum(ws.Cells(R_START, c).Value2)): If startRow < 1 Then startRow = 1
    hdrRow = startRow - 1
    Dim setting As Object: Set setting = NewDict()
    setting("mat") = ws.Cells(R_MAT, c).Value2: setting("po") = ws.Cells(R_PO, c).Value2
    setting("bb") = ws.Cells(R_BB, c).Value2: setting("qty") = ws.Cells(R_QTY, c).Value2
    setting("car") = ws.Cells(R_CAR, c).Value2

    Dim wb As Workbook, opened As Boolean, src As Worksheet
    Set wb = FindOpenWorkbook(path)
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True)
        On Error GoTo 0
        If wb Is Nothing Then Err.Raise vbObjectError + 1, , U("Kh\u00f4ng m\u1edf \u0111\u01b0\u1ee3c file ") & label & ":" & vbLf & path
        opened = True
    End If
    On Error GoTo Fail
    Set src = GetSheet(wb, sheetName)
    If src Is Nothing And wb.Worksheets.Count = 1 Then Set src = wb.Worksheets(1)   ' file chi co 1 sheet -> dung luon
    If src Is Nothing Then Err.Raise vbObjectError + 2, , U("Kh\u00f4ng t\u00ecm th\u1ea5y sheet '") & sheetName & "' trong file " & label & ":" & vbLf & path

    Dim ur As Range, lastRow As Long, lastCol As Long
    Set ur = src.UsedRange
    lastRow = ur.Row + ur.Rows.Count - 1: lastCol = ur.Column + ur.Columns.Count - 1

    ' Dong tieu de (de tu nhan cot)
    Dim hdr As Variant, hb As Variant, k As Long
    ReDim hdr(1 To lastCol)
    If hdrRow >= 1 Then
        hb = ReadBlock(src, hdrRow, hdrRow, lastCol)
        For k = 1 To lastCol: hdr(k) = hb(1, k): Next
    End If

    ' Xac dinh cot
    Dim cols As Object, used As Object, order As Variant, key As Variant, sv As String, info As String
    Set cols = NewDict(): Set used = NewDict()
    order = Array("qty", "bb", "po", "mat")
    For Each key In order
        sv = TrimS(setting(key))
        If sv = "" Or sv = AUTO Then
            cols(key) = DetectCol(hdr, KW(key), used)
            If cols(key) = 0 Then Err.Raise vbObjectError + 3, , "File " & label & U(": kh\u00f4ng t\u1ef1 nh\u1eadn \u0111\u01b0\u1ee3c c\u1ed9t '") & KW(key)(0) & U("' t\u1eeb d\u00f2ng ti\u00eau \u0111\u1ec1 (d\u00f2ng ") & hdrRow & ")." & vbLf & _
                U("H\u00e3y ch\u1ecdn ch\u1eef c\u00e1i c\u1ed9t cho '") & KW(key)(0) & U("' tr\u00ean sheet Control, ho\u1eb7c ki\u1ec3m tra 'D\u00f2ng b\u1eaft \u0111\u1ea7u \u0111\u1ecdc'.")
        Else
            cols(key) = ColIndex(sv)
            If cols(key) = 0 Then Err.Raise vbObjectError + 4, , "File " & label & U(": c\u1ed9t '") & KW(key)(0) & U("' kh\u00f4ng h\u1ee3p l\u1ec7: '") & sv & "'."
        End If
        used(cols(key)) = 1
    Next
    ' Cot Dong xe: KHONG bat buoc ("-" = khong co; tu dong khong thay thi bo qua)
    sv = TrimS(setting("car"))
    If sv = "-" Then
        cols("car") = 0
    ElseIf sv = "" Or sv = AUTO Then
        cols("car") = DetectCol(hdr, KW("car"), used)
    Else
        cols("car") = ColIndex(sv)
    End If
    For Each key In Array("mat", "po", "bb", "qty", "car")
        If cols(key) = 0 Then
            info = info & IIf(info = "", "", ", ") & KW(key)(0) & U("=(kh\u00f4ng c\u00f3)")
        Else
            info = info & IIf(info = "", "", ", ") & KW(key)(0) & "=" & ColLetter(cols(key))
            If cols(key) <= lastCol Then If TrimS(hdr(cols(key))) <> "" Then info = info & " (""" & TrimS(hdr(cols(key))) & """)"
        End If
    Next
    colInfo = info: hasCar = (cols("car") > 0)

    ' Doc du lieu 1 lan ca vung (nhanh hon doc tung o rat nhieu)
    Dim maxCol As Long, blk As Variant, nRows As Long, out() As Variant, i As Long
    maxCol = lastCol
    For Each key In Array("mat", "po", "bb", "qty", "car")
        If cols(key) > maxCol Then maxCol = cols(key)
    Next
    If lastRow >= startRow Then
        blk = ReadBlock(src, startRow, lastRow, maxCol)
        nRows = lastRow - startRow + 1
        ReDim out(0 To nRows - 1)
        For i = 1 To nRows
            out(i - 1) = Array(startRow + i - 1, blk(i, cols("mat")), blk(i, cols("po")), blk(i, cols("bb")), blk(i, cols("qty")), _
                               IIf(cols("car") > 0, blk(i, IIf(cols("car") > 0, cols("car"), 1)), Empty))
        Next
        ReadSource = out
    Else
        ReadSource = Empty
    End If
    If opened Then wb.Close SaveChanges:=False
    Exit Function
Fail:
    Dim en As Long, ed As String: en = Err.Number: ed = Err.Description
    If opened Then wb.Close SaveChanges:=False
    Err.Raise en, , ed
End Function

Private Function FindOpenWorkbook(ByVal path As String) As Workbook
    Dim w As Workbook
    For Each w In Application.Workbooks
        If LCase$(w.FullName) = LCase$(path) Then Set FindOpenWorkbook = w: Exit Function
    Next
End Function

' Doc 1 vung -> luon tra mang 2 chieu (chi so tu 1)
Private Function ReadBlock(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal nCol As Long) As Variant
    Dim v As Variant, a() As Variant
    v = ws.Range(ws.Cells(r1, 1), ws.Cells(r2, nCol)).Value2
    If IsArray(v) Then
        ReadBlock = v
    Else
        ReDim a(1 To 1, 1 To 1): a(1, 1) = v: ReadBlock = a
    End If
End Function

' Tim cot theo tieu de: khop nguyen van > khop nguyen tu > chua cum tu; bo cot da dung / co tu loai tru
Private Function DetectCol(hdr As Variant, kwd As Variant, used As Object) As Long
    Dim c As Long, h As String, toks As Variant, s As Long, best As Long, bestScore As Long
    Dim e As Variant, k As Variant, skip As Boolean
    For c = LBound(hdr) To UBound(hdr)
        If Not used.Exists(c) Then
            h = LCase$(TrimS(hdr(c)))
            If h <> "" Then
                skip = False
                For Each e In kwd(2)
                    If InStr(1, h, e, vbBinaryCompare) > 0 Then skip = True: Exit For
                Next
                If Not skip Then
                    toks = Tokens(h): s = 0
                    For Each k In kwd(1)
                        If h = k Then
                            If s < 100 Then s = 100
                        ElseIf InStr(k, " ") = 0 And InArr(toks, k) Then
                            If s < 50 Then s = 50
                        ElseIf Len(k) >= 4 And InStr(1, h, k, vbBinaryCompare) > 0 Then
                            If s < 30 Then s = 30
                        End If
                    Next
                    If s > bestScore Then bestScore = s: best = c
                End If
            End If
        End If
    Next
    DetectCol = best
End Function

Private Function Tokens(ByVal h As String) As Variant
    Dim seps As Variant, sp As Variant
    seps = Array("/", "(", ")", ".", ",", "-", "_", ":", "#", vbTab)
    For Each sp In seps: h = Replace(h, sp, " "): Next
    Do While InStr(h, "  ") > 0: h = Replace(h, "  ", " "): Loop
    Tokens = Split(Trim$(h), " ")
End Function

Private Function InArr(a As Variant, ByVal v As String) As Boolean
    Dim x As Variant
    For Each x In a
        If x = v Then InArr = True: Exit Function
    Next
End Function

' Lam sach: bo dong trong, ghi canh bao dong loi (kem so dong trong file goc) -> Array(r, mat, po, bb, qty, car)
Private Function CleanRows(raw As Variant, ByVal label As String, warns As Collection, ByRef n As Long) As Variant
    Dim out() As Variant, i As Long, x As Variant, mat As String, po As String, bb As String
    Dim rawQ As Variant, qty As Double, prob As String, sq As String
    n = 0
    If IsEmpty(raw) Then CleanRows = Empty: Exit Function
    ReDim out(0 To UBound(raw))
    For i = 0 To UBound(raw)
        x = raw(i)
        mat = CleanKey(x(F_MAT)): po = CleanKey(x(F_PO)): bb = CleanKey(x(F_BB))
        rawQ = x(F_QTY): sq = TrimS(rawQ)
        If mat = "" And po = "" And bb = "" And sq = "" Then GoTo NextRow   ' dong trong
        prob = "": qty = 0
        If VarType(rawQ) = vbDouble Or VarType(rawQ) = vbCurrency Or VarType(rawQ) = vbLong Or VarType(rawQ) = vbInteger Then
            qty = CDbl(rawQ)
        ElseIf sq = "" Then
            prob = JoinNote(prob, U("S\u1ed1 l\u01b0\u1ee3ng tr\u1ed1ng (t\u00ednh = 0)"))
        ElseIf IsNumericStr(sq) Then
            qty = Val(sq)
        Else
            prob = JoinNote(prob, U("S\u1ed1 l\u01b0\u1ee3ng kh\u00f4ng ph\u1ea3i s\u1ed1 (t\u00ednh = 0)"))
        End If
        If po = "" Then prob = JoinNote(prob, U("Thi\u1ebfu PO/SA -> B\u1ece QUA d\u00f2ng n\u00e0y"))
        If mat = "" Then prob = JoinNote(prob, U("Thi\u1ebfu Material"))
        If bb = "" Then prob = JoinNote(prob, U("Thi\u1ebfu s\u1ed1 BB"))
        If prob <> "" Then warns.Add Array(label, x(F_R), mat, po, bb, sq, prob)
        If po <> "" Then
            out(n) = Array(x(F_R), mat, po, bb, qty, NormCar(x(F_CAR)))
            n = n + 1
        End If
NextRow:
    Next
    If n = 0 Then CleanRows = Empty: Exit Function
    ReDim Preserve out(0 To n - 1)
    CleanRows = out
End Function

Private Function IsNumericStr(ByVal s As String) As Boolean
    Dim i As Long, ch As String, dot As Boolean, dig As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9]" Then
            dig = True
        ElseIf ch = "." And Not dot Then
            dot = True
        ElseIf ch = "-" And i = 1 Then
        Else
            Exit Function
        End If
    Next
    IsNumericStr = dig
End Function

' Chuan hoa ten dong xe: "vf7pro" / "VF7 PRO" -> "VF7 Pro"; ten la giu nguyen
Private Function NormCar(ByVal v As Variant) As String
    Dim s As String, k As String, i As Long
    s = CleanKey(v): If s = "" Then Exit Function
    k = LCase$(Replace(s, " ", ""))
    For i = 0 To UBound(MODELS)
        If LCase$(Replace(MODELS(i), " ", "")) = k Then NormCar = MODELS(i): Exit Function
    Next
    NormCar = s
End Function

' Thu tu dong xe: theo MODELS, xe la sau, "(chua gan)" cuoi
Private Function CarRank(ByVal c As String) As Long
    Dim i As Long
    For i = 0 To UBound(MODELS)
        If MODELS(i) = c Then CarRank = i: Exit Function
    Next
    If c = NO_CAR Or c = "" Then CarRank = 999 Else CarRank = 500
End Function

' Dong xe cua tung PO: lay tu NCC truoc, thieu thi MB51; 1 PO nhieu xe -> ghep "A / B"
Private Function CarOfPO(N As Variant, ByVal nN As Long, M As Variant, ByVal nM As Long) As Object
    Dim d As Object, i As Long, r As Variant, cur As String
    Set d = NewDict()
    For i = 0 To nN - 1
        r = N(i): If r(F_CAR) <> "" Then GoSub AddCar
    Next
    For i = 0 To nM - 1
        r = M(i): If r(F_CAR) <> "" Then GoSub AddCar
    Next
    Set CarOfPO = d
    Exit Function
AddCar:
    If Not d.Exists(r(F_PO)) Then
        d(r(F_PO)) = r(F_CAR)
    Else
        cur = d(r(F_PO))
        If InStr(1, " / " & cur & " / ", " / " & r(F_CAR) & " / ", vbBinaryCompare) = 0 Then d(r(F_PO)) = cur & " / " & r(F_CAR)
    End If
    Return
End Function

' Sheet du lieu: Material | PO/SA | So BB | So luong | Dong file goc | Dong xe  (dong 1 = tieu de)
Private Sub WriteDataSheet(wb As Workbook, ByVal nm As String, rows As Variant, ByVal n As Long, ByVal label As String)
    Dim ws As Worksheet, data() As Variant, i As Long, r As Variant
    Set ws = AddSheetLast(wb): ws.Name = nm
    ws.Columns(3).NumberFormat = "@"   ' giu so BB dang chu (0011)
    ReDim data(1 To n + 1, 1 To 6)
    data(1, 1) = "Material": data(1, 2) = "PO/SA": data(1, 3) = U("S\u1ed1 BB"): data(1, 4) = U("S\u1ed1 l\u01b0\u1ee3ng (") & label & ")"
    data(1, 5) = U("D\u00f2ng file g\u1ed1c"): data(1, 6) = U("D\u00f2ng xe")
    For i = 0 To n - 1
        r = rows(i)
        data(i + 2, 1) = ToCell(r(F_MAT)): data(i + 2, 2) = ToCell(r(F_PO)): data(i + 2, 3) = r(F_BB)
        data(i + 2, 4) = r(F_QTY): data(i + 2, 5) = r(F_R): data(i + 2, 6) = r(F_CAR)
    Next
    ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 6)).Value2 = data
    FmtHeader ws, 1, 6
    AutoFitCols ws, 6
End Sub

' Ma toan so (<=15 chu so, khong bat dau bang 0) ghi dang SO de SUMIFS khop; con lai giu chu
Private Function ToCell(ByVal s As String) As Variant
    If Len(s) >= 1 And Len(s) <= 15 And Left$(s, 1) Like "[1-9]" And IsAllDigits(s) Then ToCell = CDbl(s) Else ToCell = s
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Not Mid$(s, i, 1) Like "[0-9]" Then Exit Function
    Next
    IsAllDigits = (Len(s) > 0)
End Function

' So sanh: so theo so, chu theo chu
Private Function CmpVal(ByVal a As Variant, ByVal b As Variant) As Long
    Dim sa As String, sb As String, da As Double, db As Double
    sa = CStr(a): sb = CStr(b)
    If sa <> "" And sb <> "" And IsNumericStr(sa) And IsNumericStr(sb) Then
        da = Val(sa): db = Val(sb)
        If da < db Then CmpVal = -1 Else If da > db Then CmpVal = 1 Else CmpVal = 0
    Else
        CmpVal = StrComp(sa, sb, vbBinaryCompare)
    End If
End Function

' mode 0: CmpVal; mode 1: PO theo dong xe (gCarOf) roi theo so PO
Private Function CmpMode(ByVal a As Variant, ByVal b As Variant, ByVal mode As Long) As Long
    Dim ca As String, cb As String, d As Long
    If mode = 1 Then
        If gCarOf.Exists(a) Then ca = gCarOf(a) Else ca = ""
        If gCarOf.Exists(b) Then cb = gCarOf(b) Else cb = ""
        d = CarRank(IIf(ca = "", NO_CAR, ca)) - CarRank(IIf(cb = "", NO_CAR, cb))
        If d = 0 Then d = CmpVal(ca, cb)
        If d = 0 Then d = CmpVal(a, b)
        CmpMode = d
    Else
        CmpMode = CmpVal(a, b)
    End If
End Function

' Sap xep tron (on dinh) mang 1 chieu
Private Sub SortArr(arr As Variant, ByVal mode As Long)
    Dim tmp As Variant
    If Not IsArray(arr) Then Exit Sub
    If UBound(arr) <= LBound(arr) Then Exit Sub
    tmp = arr
    MergeSort arr, tmp, LBound(arr), UBound(arr), mode
End Sub

Private Sub MergeSort(arr As Variant, tmp As Variant, ByVal lo As Long, ByVal hi As Long, ByVal mode As Long)
    Dim mid_ As Long, i As Long, j As Long, k As Long
    If hi <= lo Then Exit Sub
    mid_ = (lo + hi) \ 2
    MergeSort arr, tmp, lo, mid_, mode
    MergeSort arr, tmp, mid_ + 1, hi, mode
    i = lo: j = mid_ + 1: k = lo
    Do While i <= mid_ And j <= hi
        If CmpMode(arr(j), arr(i), mode) < 0 Then
            tmp(k) = arr(j): j = j + 1
        Else
            tmp(k) = arr(i): i = i + 1
        End If
        k = k + 1
    Loop
    Do While i <= mid_: tmp(k) = arr(i): i = i + 1: k = k + 1: Loop
    Do While j <= hi: tmp(k) = arr(j): j = j + 1: k = k + 1: Loop
    For k = lo To hi: arr(k) = tmp(k): Next
End Sub

' Trang thai 1 PO/material -> Array(chu, loai mau)
Private Function PoStatus(ByVal inN As Boolean, ByVal inM As Boolean, ByVal gap As Double) As Variant
    If Not inM Then PoStatus = Array(U("\u26a0 Ch\u01b0a GR (ch\u1ec9 c\u00f3 NCC)"), "warn"): Exit Function
    If Not inN Then PoStatus = Array(U("\u26a0 Ch\u1ec9 c\u00f3 GR (NCC kh\u00f4ng c\u00f3)"), "warn"): Exit Function
    If Round6(gap) = 0 Then PoStatus = Array(U("\u2714 Kh\u1edbp"), "ok"): Exit Function
    If gap > 0 Then PoStatus = Array(U("\u2718 NCC th\u1eeba / GR thi\u1ebfu"), "bad") Else PoStatus = Array(U("\u2718 GR th\u1eeba / NCC thi\u1ebfu"), "bad")
End Function

Private Function AddSheetLast(wb As Workbook) As Worksheet
    Set AddSheetLast = AddSheetAfter(wb, wb.Sheets(wb.Sheets.Count))
End Function

' Them sheet moi ngay sau 1 sheet
Private Function AddSheetAfter(wb As Workbook, ByVal afterSh As Object) As Worksheet
    Set AddSheetAfter = wb.Worksheets.Add(After:=afterSh)
End Function

'========================= SHEET CONTROL =========================
Private Function EnsureControl(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheet(wb, CTRL)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Sheets(1))
        ws.Name = CTRL
        BuildControl ws
    Else
        UnprotectSheet ws
    End If
    ApplyUI ws
    Set EnsureControl = ws
End Function

Private Sub BuildControl(ws As Worksheet)
    Dim t As Range, h As Range, rr As Variant
    ws.Cells.Clear
    ws.Cells(1, 1).Value2 = U("TOOL \u0110\u1ed0I SO\u00c1T NCC  <->  MB51 (GR)")
    Set t = MergeRange(ws, 1, 1, 1, 3)
    t.Font.Bold = True: t.Font.Size = 14: t.Font.Color = CLR_WHITE
    t.Interior.Color = CLR_HDR_BG: t.HorizontalAlignment = xlCenter
    ws.Rows(1).RowHeight = 26
    ws.Cells(3, 1).Value2 = U("TH\u00d4NG S\u1ed0"): ws.Cells(3, 2).Value2 = "FILE NCC": ws.Cells(3, 3).Value2 = "FILE MB51"
    Set h = ws.Range(ws.Cells(3, 1), ws.Cells(3, 3))
    h.Font.Bold = True: h.Interior.Color = CLR_LBLUE: h.HorizontalAlignment = xlCenter
    ws.Cells(R_PATH, 1).Value2 = U("\u0110\u01b0\u1eddng d\u1eabn file")
    ws.Cells(R_SHEET, 1).Value2 = U("T\u00ean sheet")
    ws.Cells(R_START, 1).Value2 = U("D\u00f2ng b\u1eaft \u0111\u1ea7u \u0111\u1ecdc")
    ws.Cells(R_MAT, 1).Value2 = U("C\u1ed9t: M\u00e3 material")
    ws.Cells(R_PO, 1).Value2 = U("C\u1ed9t: M\u00e3 PO/SA")
    ws.Cells(R_BB, 1).Value2 = U("C\u1ed9t: S\u1ed1 bi\u00ean b\u1ea3n (BB)")
    ws.Cells(R_SEQ, 1).Value2 = U("C\u1ed9t: Seq")
    ws.Cells(R_QTY, 1).Value2 = U("C\u1ed9t: S\u1ed1 l\u01b0\u1ee3ng")
    For Each rr In Array(R_SHEET, R_START, R_MAT, R_PO, R_BB, R_SEQ, R_QTY, R_CAR)
        ws.Cells(rr, C_NCC).Value2 = DEFV(rr)(0)
        ws.Cells(rr, C_MB).Value2 = DEFV(rr)(1)
    Next
    ws.Cells(R_BATCH, C_NCC).Value2 = DEFV(R_BATCH)(0)
    ws.Range(ws.Cells(R_PATH, C_NCC), ws.Cells(R_CAR, C_MB)).Interior.Color = CLR_YELLOW
    ws.Cells(R_BATCH, C_NCC).Interior.Color = CLR_YELLOW
    With ws.Range(ws.Cells(3, 1), ws.Cells(R_CAR, 3)).Borders
        .LineStyle = xlContinuous: .Color = RGB(180, 180, 180)
    End With
    ws.Columns(1).ColumnWidth = 24: ws.Columns(2).ColumnWidth = 40: ws.Columns(3).ColumnWidth = 40
End Sub

' Khoa/mo o + danh sach chon + goi y mac dinh + cac nut (chay lai moi lan, khong doi gia tri nguoi dung nhap)
Private Sub ApplyUI(ws As Worksheet)
    Dim c As Long, r As Variant, f As String, cell As Range, pi As Variant, lb As Range
    ws.Cells.Locked = True
    ws.Cells(R_CAR, 1).Value2 = U("C\u1ed9t: D\u00f2ng xe")
    For c = C_NCC To C_MB
        If TrimS(ws.Cells(R_CAR, c).Value2) = "" Then ws.Cells(R_CAR, c).Value2 = DEFV(R_CAR)(c - C_NCC)
    Next
    With ws.Range(ws.Cells(3, 1), ws.Cells(R_CAR, 3)).Borders
        .LineStyle = xlContinuous: .Color = RGB(180, 180, 180)
    End With
    For c = C_NCC To C_MB
        If c = C_NCC Then f = "NCC" Else f = "MB51"
        ws.Cells(R_PATH, c).Locked = False
        AddVal ws.Cells(R_PATH, c), xlValidateInputOnly, "", "", U("\u0110\u01b0\u1eddng d\u1eabn file ") & f, _
            U("B\u1ea5m n\u00fat 'Ch\u1ecdn file ") & f & U("' ho\u1eb7c d\u00e1n \u0111\u01b0\u1eddng d\u1eabn \u0111\u1ea7y \u0111\u1ee7, VD: D:\u005CData\u005C") & f & ".xlsx"
        ws.Cells(R_SHEET, c).Locked = False
        AddVal ws.Cells(R_SHEET, c), xlValidateInputOnly, "", "", U("T\u00ean sheet ch\u1ee9a d\u1eef li\u1ec7u"), _
            U("M\u1eb7c \u0111\u1ecbnh: ") & DEFV(R_SHEET)(c - C_NCC) & U(". File ch\u1ec9 c\u00f3 1 sheet th\u00ec tool t\u1ef1 d\u00f9ng sheet \u0111\u00f3.")
        ws.Cells(R_START, c).Locked = False
        AddVal ws.Cells(R_START, c), xlValidateWholeNumber, "1", "1048576", U("D\u00f2ng b\u1eaft \u0111\u1ea7u \u0111\u1ecdc"), _
            U("D\u00f2ng d\u1eef li\u1ec7u \u0111\u1ea7u ti\u00ean. D\u00f2ng ngay tr\u00ean l\u00e0 d\u00f2ng TI\u00caU \u0110\u1ec0 (d\u00f9ng \u0111\u1ec3 t\u1ef1 nh\u1eadn c\u1ed9t). M\u1eb7c \u0111\u1ecbnh: ") & DEFV(R_START)(c - C_NCC)
        For Each r In Array(R_MAT, R_PO, R_BB, R_SEQ, R_QTY, R_CAR)
            Set cell = ws.Cells(r, c)
            cell.Locked = False: cell.Interior.Color = CLR_YELLOW
            If r = R_SEQ Then
                AddVal cell, xlValidateList, "-," & COL_LIST, "", U("C\u1ed9t Seq (kh\u00f4ng b\u1eaft bu\u1ed9c)"), U("Tool kh\u00f4ng d\u00f9ng c\u1ed9t n\u00e0y \u0111\u1ec3 t\u00ednh. \u0110\u1ec3 '-' n\u1ebfu kh\u00f4ng c\u00f3.")
            ElseIf r = R_CAR Then
                AddVal cell, xlValidateList, AUTO & ",-," & COL_LIST, "", U("C\u1ed9t D\u00f2ng xe (kh\u00f4ng b\u1eaft bu\u1ed9c)"), _
                    U("C\u1ed9t ghi d\u00f2ng xe: ") & Join(MODELS, ", ") & ". '" & AUTO & U("' = t\u1ef1 t\u00ecm theo ti\u00eau \u0111\u1ec1; '-' = file kh\u00f4ng c\u00f3 c\u1ed9t n\u00e0y.")
            Else
                AddVal cell, xlValidateList, AUTO & "," & COL_LIST, "", U("Ch\u1ecdn c\u1ed9t"), _
                    "'" & AUTO & U("' = tool t\u1ef1 t\u00ecm c\u1ed9t theo ti\u00eau \u0111\u1ec1. Ho\u1eb7c ch\u1ecdn ch\u1eef c\u00e1i c\u1ed9t. M\u1eb7c \u0111\u1ecbnh: ") & DEFV(r)(c - C_NCC)
            End If
        Next
    Next
    ws.Cells(R_BATCH, 1).Value2 = U("S\u1ed1 t\u1edd chu\u1ea9n / 1 l\u1ea7n GR")
    ws.Cells(R_BATCH, C_NCC).Locked = False
    AddVal ws.Cells(R_BATCH, C_NCC), xlValidateWholeNumber, "1", "100", U("S\u1ed1 t\u1edd chu\u1ea9n / 1 l\u1ea7n GR"), _
        U("Ch\u1ec9 d\u00f9ng \u0111\u1ec3 C\u1ea2NH B\u00c1O khi 1 l\u1ea7n GR g\u1ed9p kh\u00e1c s\u1ed1 t\u1edd. Kho\u1ea3ng BB \u0111\u01b0\u1ee3c chia theo t\u1edd \u0111\u1ea7u th\u1eadt trong MB51. M\u1eb7c \u0111\u1ecbnh: ") & DEFV(R_BATCH)(0)
    ' Thong so bo cau: 1 xe = N ma cau, 1 chuyen (1 BB) = M xe. 0 = tat kiem tra
    For Each pi In Array(Array(R_CODES, U("S\u1ed1 m\u00e3 c\u1ea7u / 1 xe"), U("M\u1ed7i xe c\u1ea7n \u0111\u1ee7 bao nhi\u00eau m\u00e3 c\u1ea7u kh\u00e1c nhau (m\u1ed7i m\u00e3 1 c\u00e1i). 0 = t\u1eaft ki\u1ec3m tra b\u1ed9. M\u1eb7c \u0111\u1ecbnh: 4")), _
                         Array(R_TRIP, U("S\u1ed1 xe / 1 chuy\u1ebfn (1 BB)"), U("1 bi\u00ean b\u1ea3n (BB) = 1 chuy\u1ebfn giao bao nhi\u00eau xe. VD 4 xe x 4 m\u00e3 = 16 c\u00e1i/chuy\u1ebfn. 0 = kh\u00f4ng ki\u1ec3m tra chuy\u1ebfn. M\u1eb7c \u0111\u1ecbnh: 4")))
        Set lb = ws.Cells(pi(0), 1)
        lb.Value2 = pi(1): lb.Font.Italic = False: lb.Font.Color = 0
        If TrimS(ws.Cells(pi(0), C_NCC).Value2) = "" Then ws.Cells(pi(0), C_NCC).Value2 = DEFV(pi(0))(0)
        ws.Cells(pi(0), C_NCC).Locked = False: ws.Cells(pi(0), C_NCC).Interior.Color = CLR_YELLOW
        AddVal ws.Cells(pi(0), C_NCC), xlValidateWholeNumber, "0", "50", pi(1), pi(2)
    Next
    With ws.Range(ws.Cells(R_BATCH, 1), ws.Cells(R_TRIP, 2)).Borders
        .LineStyle = xlContinuous: .Color = RGB(180, 180, 180)
    End With
    ws.Cells(2, 1).Value2 = U("\u25b6 Ch\u1ecdn 2 file (n\u00fat b\u00ean ph\u1ea3i) -> b\u1ea5m CHECK CH\u00caNH L\u1ec6CH. C\u1ed9t \u0111\u1ec3 '") & AUTO & U("' th\u00ec tool t\u1ef1 nh\u1eadn theo ti\u00eau \u0111\u1ec1.")
    ws.Cells(2, 1).Font.Bold = True: ws.Cells(2, 1).Font.Color = RGB(192, 0, 0)
    ws.Cells(R_NOTE, 1).Value2 = U("Ghi ch\u00fa: 1 BB = 1 chuy\u1ebfn giao. BB c\u1ee7a MB51 = s\u1ed1 BB T\u1edc \u0110\u1ea6U khi GR; tool chia kho\u1ea3ng BB c\u1ee7a NCC theo c\u00e1c t\u1edd \u0111\u1ea7u n\u00e0y. Ch\u1ec9 s\u1eeda \u0111\u01b0\u1ee3c \u00f4 v\u00e0ng.")
    ws.Cells(R_NOTE, 1).Font.Italic = True: ws.Cells(R_NOTE, 1).Font.Color = RGB(110, 110, 110)
    ws.Cells(R_NOTE + 1, 1).Value2 = U("Ph\u00e1t tri\u1ec3n b\u1edfi Nguy\u1ec5n Vi\u1ebft D\u01b0\u01a1ng")
    ws.Cells(R_NOTE + 1, 1).Font.Size = 9: ws.Cells(R_NOTE + 1, 1).Font.Color = RGB(128, 128, 128)
    ' O nhap canh trai het (so mac dinh canh phai -> nhin lech so voi chu)
    ws.Range(ws.Cells(R_PATH, C_NCC), ws.Cells(R_CAR, C_MB)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(R_BATCH, C_NCC), ws.Cells(R_TRIP, C_NCC)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(R_PATH, C_NCC), ws.Cells(R_PATH, C_MB)).NumberFormat = "@"
    ' Bang nut ben phai (cot E)
    ws.Columns(4).ColumnWidth = 3: ws.Columns(5).ColumnWidth = 30
    EnsureButton ws, "btnCheck", U("CHECK CH\u00caNH L\u1ec6CH"), "CheckChenhLech", 3, 6, RGB(0, 128, 64), 15
    EnsureButton ws, "btnFileNCC", U("Ch\u1ecdn file NCC"), "ChonFileNCC", 7, 8, RGB(47, 84, 150), 11
    EnsureButton ws, "btnFileMB51", U("Ch\u1ecdn file MB51"), "ChonFileMB51", 9, 10, RGB(47, 84, 150), 11
    EnsureButton ws, "btnSave", U("L\u01b0u k\u1ebft qu\u1ea3 (.xlsx)"), "LuuKetQua", 12, 13, RGB(0, 32, 96), 11
End Sub

' Nut (hinh chu nhat bo goc gan macro) trong cot E, tu dong rTop den rBot; co roi thi cap nhat lai
Private Sub EnsureButton(ws As Worksheet, ByVal nm As String, ByVal text As String, ByVal macro As String, _
                         ByVal rTop As Long, ByVal rBot As Long, ByVal color As Long, ByVal size As Single)
    Dim b As Shape, s As Shape, cTop As Range, cBot As Range, L As Single, T As Single, W As Single, H As Single
    For Each s In ws.Shapes
        If s.Name = nm Then Set b = s: Exit For
    Next
    Set cTop = ws.Cells(rTop, 5): Set cBot = ws.Cells(rBot, 5)
    L = cTop.Left + 2: T = cTop.Top + 3: W = cTop.Width - 4: H = cBot.Top + cBot.Height - cTop.Top - 6
    If b Is Nothing Then
        Set b = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, T, W, H)
        b.Name = nm
    Else
        b.Left = L: b.Top = T: b.Width = W: b.Height = H
    End If
    b.Fill.Visible = msoTrue: b.Fill.Solid: b.Fill.ForeColor.RGB = color
    b.Line.Visible = msoFalse
    On Error Resume Next
    b.Shadow.Visible = msoFalse
    On Error GoTo 0
    With b.TextFrame2
        .TextRange.Text = text
        .TextRange.Font.Bold = msoTrue: .TextRange.Font.Size = size
        .TextRange.Font.Fill.ForeColor.RGB = CLR_WHITE
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
    End With
    b.Placement = xlFreeFloating
    b.OnAction = macro
End Sub

Private Sub AddVal(cell As Range, ByVal typ As Long, ByVal f1 As String, ByVal f2 As String, ByVal title As String, ByVal info As String)
    On Error Resume Next
    With cell.Validation
        .Delete
        If typ = xlValidateInputOnly Then
            .Add Type:=typ
        ElseIf typ = xlValidateList Then
            .Add Type:=typ, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=f1
        Else
            .Add Type:=typ, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=f1, Formula2:=f2
        End If
        .IgnoreBlank = True
        If typ = xlValidateList Then .InCellDropdown = True
        .InputTitle = Left$(title, 32): .InputMessage = Left$(info, 255): .ShowInput = True
        .ErrorTitle = U("Gi\u00e1 tr\u1ecb kh\u00f4ng h\u1ee3p l\u1ec7"): .ErrorMessage = Left$(info, 255): .ShowError = True
    End With
    On Error GoTo 0
End Sub

Private Sub ProtectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=PASS, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Private Sub UnprotectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect PASS
End Sub

'========================= SHEET CHI TIET 1 SA/PO =========================
Private Function MakeDetailSheet(wb As Workbook, ByVal po As String, o As Object, ByVal batch As Long, ByVal car As String, bo As Object) As String
    Dim ws As Worksheet, NC As Long, r As Long, tt As Range, mats As Variant, i As Long, mat As Variant
    Dim gap As Double, st As Variant, nOk As Long, nBad As Long, nWarn As Long, badMats As New Collection
    Set ws = AddSheetLast(wb)
    ws.Name = SafeSheetName(wb, po)
    NC = 8: r = 1
    ws.Cells(r, 1).Value2 = DETAIL_TAG & U(" \u2014 SA/PO: ") & po & IIf(car <> "", U("   |   D\u00f2ng xe: ") & car, "")
    Set tt = MergeRange(ws, r, 1, r, NC)
    tt.Font.Bold = True: tt.Font.Size = 13: tt.Font.Color = CLR_WHITE
    tt.Interior.Color = CLR_HDR_BG: tt.HorizontalAlignment = xlCenter
    ws.Rows(r).RowHeight = 24
    AddLink ws, ws.Cells(2, 1), SUM_NAME, U("\u2190 V\u1ec1 ") & SUM_NAME
    r = r + 2
    WriteRow ws, r, Array("Material", U("T\u1ed5ng NCC"), U("T\u1ed5ng GR (MB51)"), U("L\u1ec7ch (NCC-GR)"), U("Tr\u1ea1ng th\u00e1i"))
    FmtHeader ws, r, 5: r = r + 1

    ' Moi material trong PO (cung ma -> cong don): khop xanh / lech do / 1 ben vang
    mats = o("matsArr")
    For i = 0 To UBound(mats)
        mat = mats(i)
        gap = Round6(DGet(o("n"), mat) - DGet(o("m"), mat))
        st = PoStatus(o("inN").Exists(mat), o("inM").Exists(mat), gap)
        ws.Cells(r, 1).NumberFormat = "@"
        WriteRow ws, r, Array(IIf(mat = "", U("(tr\u1ed1ng)"), mat), DGet(o("n"), mat), DGet(o("m"), mat), gap, st(0))
        ColorRowKind ws, r, 5, st(1)
        Select Case st(1)
            Case "ok": nOk = nOk + 1
            Case "bad": nBad = nBad + 1: badMats.Add mat
            Case Else: nWarn = nWarn + 1: badMats.Add mat
        End Select
        r = r + 1
    Next
    ws.Cells(r, 1).Value2 = U("PO n\u00e0y: ") & (UBound(mats) + 1) & U(" material \u2014 ") & nOk & U(" kh\u1edbp, ") & nBad & U(" l\u1ec7ch, ") & nWarn & U(" ch\u1ec9 c\u00f3 1 b\u00ean.")
    ws.Cells(r, 1).Font.Italic = True
    r = r + 2

    ' Kiem tra bo cau
    Dim hk As Range, t As Variant, k As Variant, parts As String, nP As Long, blk() As Variant, nb As Long, tb As Range
    If bo("cpc") > 0 Then
        ws.Cells(r, 1).Value2 = U("\u25ba KI\u1ec2M TRA B\u1ed8 C\u1ea6U   (1 xe = ") & bo("cpc") & U(" m\u00e3 c\u1ea7u") & _
            IIf(bo("cpt") > 0, U(", 1 chuy\u1ebfn = ") & bo("cpt") & " xe = " & (bo("cpc") * bo("cpt")) & U(" c\u00e1i"), "") & ")"
        Set hk = MergeRange(ws, r, 1, r, NC): hk.Font.Bold = True: hk.Interior.Color = CLR_LBLUE
        r = r + 1
        WriteRow ws, r, Array(U("S\u1ed1 m\u00e3 c\u1ea7u"), "Xe giao (NCC)", U("Xe \u0111\u00e3 GR"), U("K\u1ebft qu\u1ea3"))
        FmtHeader ws, r, 4: r = r + 1
        ws.Cells(r, 1).NumberFormat = "@"   ' "4/4" khong bi hieu la ngay
        WriteRow ws, r, Array(bo("nCodes") & "/" & bo("cpc"), bo("carsN"), bo("carsM"), IIf(bo("ok"), U("\u2714 \u0110\u1ee7 b\u1ed9"), U("\u26a0 ") & bo("issues")))
        ws.Cells(r, 1).HorizontalAlignment = xlLeft
        ColorRowKind ws, r, 4, IIf(bo("ok"), "ok", "warn")
        r = r + 2
        If bo("nTripBad") > 0 Then
            WriteRow ws, r, Array(U("Chuy\u1ebfn (BB)"), U("S\u1ed1 m\u00e3"), U("S\u1ed1 l\u01b0\u1ee3ng t\u1eebng m\u00e3"), U("T\u1ed5ng"), U("Ghi ch\u00fa"))
            FmtHeader ws, r, 5: r = r + 1
            ReDim blk(1 To bo("nTripBad"), 1 To 5): nb = 0
            For Each t In bo("trips")
                If t("note") <> "" Then
                    parts = "": nP = 0
                    For Each k In t("qty").Keys
                        parts = parts & IIf(parts = "", "", "   ") & k & ": " & t("qty")(k): nP = nP + 1
                    Next
                    nb = nb + 1
                    blk(nb, 1) = t("bb"): blk(nb, 2) = nP: blk(nb, 3) = parts: blk(nb, 4) = t("total"): blk(nb, 5) = t("note")
                End If
            Next
            ws.Range(ws.Cells(r, 1), ws.Cells(r + nb - 1, 1)).NumberFormat = "@"
            Set tb = ws.Range(ws.Cells(r, 1), ws.Cells(r + nb - 1, 5))
            tb.Value2 = blk   ' ghi ca khoi 1 lan cho nhanh
            tb.Interior.Color = CLR_WARN_BG: tb.Font.Color = CLR_WARN_FT
            tb.Borders.LineStyle = xlContinuous: tb.Borders.Color = CLR_GRID
            r = r + nb
            ws.Cells(r, 1).Value2 = (bo("trips").Count - bo("nTripBad")) & "/" & bo("trips").Count & U(" chuy\u1ebfn c\u00f2n l\u1ea1i \u0111\u1ee7 b\u1ed9.")
            ws.Cells(r, 1).Font.Italic = True
            r = r + 2
        ElseIf bo("trips").Count > 0 Then
            ws.Cells(r, 1).Value2 = U("T\u1ea5t c\u1ea3 ") & bo("trips").Count & U(" chuy\u1ebfn \u0111\u1ec1u \u0111\u1ee7 ") & bo("cpc") & U(" m\u00e3") & IIf(bo("cpt") > 0, " / " & (bo("cpc") * bo("cpt")) & U(" c\u00e1i."), ".")
            ws.Cells(r, 1).Font.Italic = True
            r = r + 2
        End If
    End If

    ' Soi tung material lech: chia khoang BB theo TO DAU that trong MB51
    Dim grp As Collection, g As Variant, gi As Long, hb As Range, d As Double
    For Each mat In badMats
        ws.Cells(r, 1).Value2 = U("\u25ba Material: ") & IIf(mat = "", U("(tr\u1ed1ng)"), mat) & U("   \u2014 chi ti\u1ebft theo bi\u00ean b\u1ea3n (BB)")
        Set hb = MergeRange(ws, r, 1, r, NC): hb.Font.Bold = True: hb.Interior.Color = CLR_LBLUE
        r = r + 1
        WriteRow ws, r, Array(U("Kho\u1ea3ng"), U("BB \u0111\u1ea7u (ti\u00eau \u0111\u1ec1 GR)"), U("C\u00e1c BB trong kho\u1ea3ng"), U("S\u1ed1 t\u1edd"), "NCC giao", "GR (MB51)", U("L\u1ec7ch"), U("Ghi ch\u00fa"))
        FmtHeader ws, r, NC: r = r + 1
        Set grp = SplitByGR(o("rowsN"), o("rowsM"), CStr(mat), batch)
        gi = 0
        For Each g In grp
            gi = gi + 1
            ws.Cells(r, 2).NumberFormat = "@"   ' giu so BB dang chu (0011)
            d = Round6(g("nSum") - g("grQ"))
            WriteRow ws, r, Array(gi, g("title"), ShowBBs(g("members")), g("members").Count, g("nSum"), g("grQ"), d, g("note"))
            If d = 0 Then
                ColorRowKind ws, r, NC, IIf(g("note") <> "", "warn", "ok")
            Else
                ColorRowKind ws, r, NC, "bad"
            End If
            r = r + 1
        Next
        r = r + 1
    Next
    AutoFitCols ws, NC
    ws.Columns(3).ColumnWidth = 34: ws.Columns(8).ColumnWidth = 40
    MakeDetailSheet = ws.Name
End Function

' Chia BB cua NCC theo cac to dau co trong MB51: moi BB thuoc to dau gan nhat <= no (rows: cac dong cua 1 PO)
Private Function SplitByGR(ByVal rowsN As Object, ByVal rowsM As Object, ByVal mat As String, ByVal batch As Long) As Collection
    Dim dBB As Object, dT As Object, r As Variant, out As New Collection, bbs As Variant, ts As Variant
    Dim i As Long, j As Long, mem As Collection, g As Object, notes As String, hasBB As Boolean, hasT As Boolean
    Set dBB = NewDict(): Set dT = NewDict()
    For Each r In rowsN
        If r(F_MAT) = mat Then dBB(r(F_BB)) = DGet(dBB, r(F_BB)) + r(F_QTY)
    Next
    For Each r In rowsM
        If r(F_MAT) = mat Then dT(r(F_BB)) = DGet(dT, r(F_BB)) + r(F_QTY)
    Next
    hasBB = dBB.Count > 0: hasT = dT.Count > 0
    If hasBB Then bbs = dBB.Keys: SortArr bbs, 0
    If hasT Then ts = dT.Keys: SortArr ts, 0

    If Not hasT Then
        Set mem = New Collection
        If hasBB Then
            For i = 0 To UBound(bbs): mem.Add bbs(i): Next
        End If
        out.Add NewGroup(U("\u2014"), mem, SumBB(dBB, mem), 0, U("Ch\u01b0a GR (MB51 kh\u00f4ng c\u00f3 material n\u00e0y trong PO)"))
        Set SplitByGR = out: Exit Function
    End If

    Set mem = New Collection
    If hasBB Then
        For i = 0 To UBound(bbs)
            If CmpVal(bbs(i), ts(0)) < 0 Then mem.Add bbs(i)
        Next
    End If
    If mem.Count > 0 Then out.Add NewGroup(U("(tr\u01b0\u1edbc t\u1edd \u0111\u1ea7u)"), mem, SumBB(dBB, mem), 0, U("BB n\u1eb1m tr\u01b0\u1edbc t\u1edd \u0111\u1ea7u GR \u0111\u1ea7u ti\u00ean -> ch\u01b0a \u0111\u01b0\u1ee3c GR"))

    For j = 0 To UBound(ts)
        Set mem = New Collection
        If hasBB Then
            For i = 0 To UBound(bbs)
                If CmpVal(bbs(i), ts(j)) >= 0 Then
                    If j = UBound(ts) Then
                        mem.Add bbs(i)
                    ElseIf CmpVal(bbs(i), ts(j + 1)) < 0 Then
                        mem.Add bbs(i)
                    End If
                End If
            Next
        End If
        notes = ""
        If ts(j) = "" Then
            notes = U("GR kh\u00f4ng c\u00f3 s\u1ed1 BB")
        ElseIf Not dBB.Exists(ts(j)) Then
            notes = U("T\u1edd \u0111\u1ea7u GR kh\u00f4ng c\u00f3 trong NCC")
        End If
        If mem.Count = 0 Then
            notes = JoinNote(notes, U("NCC kh\u00f4ng c\u00f3 BB n\u00e0o trong kho\u1ea3ng"))
        ElseIf mem.Count <> batch Then
            notes = JoinNote(notes, U("G\u1ed9p ") & mem.Count & U(" t\u1edd (chu\u1ea9n ") & batch & ")")
        End If
        out.Add NewGroup(ts(j), mem, SumBB(dBB, mem), dT(ts(j)), notes)
    Next
    Set SplitByGR = out
End Function

Private Function NewGroup(ByVal title As String, ByVal mem As Object, ByVal nSum As Double, ByVal grQ As Double, ByVal note As String) As Object
    Dim g As Object: Set g = NewDict()
    g("title") = title: g.Add "members", mem: g("nSum") = nSum: g("grQ") = grQ: g("note") = note
    Set NewGroup = g
End Function

Private Function SumBB(ByVal dBB As Object, ByVal mem As Object) As Double
    Dim x As Variant, s As Double
    For Each x In mem: s = s + dBB(x): Next
    SumBB = s
End Function

' Danh sach BB gon: nhieu qua thi rut "dau ... cuoi"
Private Function ShowBBs(ByVal mem As Object) As String
    Dim i As Long, s As String
    If mem.Count <= 8 Then
        For i = 1 To mem.Count: s = s & IIf(s = "", "", "  ") & mem(i): Next
    Else
        s = mem(1) & "  " & mem(2) & "  " & mem(3) & U("  \u2026  ") & mem(mem.Count - 1) & "  " & mem(mem.Count)
    End If
    ShowBBs = s
End Function

'========================= SHEET TONG HOP (ngay sau Control) =========================
' Cot: A SA/PO | B Dong xe | C NCC da giao | D GR (MB51) | E Gap | F Trang thai |
'      G So ma cau | H Xe giao (NCC) | I Xe da GR | J Bo cau | K Sheet chi tiet
' Ben phai (cot M..T): bang tong theo dong xe (neu file co cot Dong xe)
Private Function WriteSummary(wb As Workbook, ctrlWs As Worksheet, N As Variant, ByVal nN As Long, M As Variant, ByVal nM As Long, _
                              sheetOf As Object, ByVal nWarn As Long, carOf As Object, pos As Variant, boOf As Object, _
                              ByVal cpc As Long, ByVal colNote As String) As Object
    Dim sumN As Scripting.Dictionary, sumM As Scripting.Dictionary, inN As Scripting.Dictionary, inM As Scripting.Dictionary
    Dim i As Long, r As Variant, po As String
    Set sumN = New Scripting.Dictionary: Set sumM = New Scripting.Dictionary
    Set inN = New Scripting.Dictionary: Set inM = New Scripting.Dictionary
    For i = 0 To nN - 1
        r = N(i): po = r(F_PO)
        If sumN.Exists(po) Then sumN(po) = sumN(po) + r(F_QTY) Else sumN.Add po, r(F_QTY): inN.Add po, 1
    Next
    For i = 0 To nM - 1
        r = M(i): po = r(F_PO)
        If sumM.Exists(po) Then sumM(po) = sumM(po) + r(F_QTY) Else sumM.Add po, r(F_QTY): inM.Add po, 1
    Next
    Dim NC As Long: NC = SUM_NC

    Dim ws As Worksheet, tt As Range
    Set ws = AddSheetAfter(wb, ctrlWs)
    ws.Name = SafeSheetName(wb, SUM_NAME)
    ws.Cells(1, 1).Value2 = U("T\u1ed4NG H\u1ee2P THEO SA/PO \u2014 NCC \u0111\u00e3 giao vs GR (MB51)   (") & Format(Now, "dd/mm/yyyy hh:nn") & ")"
    Set tt = MergeRange(ws, 1, 1, 1, NC)
    tt.Font.Bold = True: tt.Font.Size = 13: tt.Font.Color = CLR_WHITE
    tt.Interior.Color = CLR_HDR_BG: tt.HorizontalAlignment = xlCenter
    ws.Rows(1).RowHeight = 24
    ws.Cells(2, 1).Value2 = U("S\u1ed1 SA/PO: ") & (UBound(pos) + 1) & U("   |   S\u1ed1 li\u1ec7u = SUMIFS t\u1eeb sheet ") & DATA_N & " / " & DATA_M & U(" (m\u1ed7i file = 1 ng\u00e0y)")
    ws.Cells(2, 1).Font.Italic = True
    AddLink ws, ws.Cells(2, NC), CTRL, U("\u2190 V\u1ec1 ") & CTRL
    If nWarn > 0 Then
        AddLink ws, ws.Cells(3, 5), WARN_NAME, U("\u26a0 ") & nWarn & U(" d\u00f2ng d\u1eef li\u1ec7u l\u1ed7i \u2192 xem")
        ColorCell ws.Cells(3, 5), False
    End If
    If colNote <> "" Then   ' cot tool da dung
        ws.Cells(4, 1).Value2 = colNote
        With ws.Cells(4, 1).Font: .Italic = True: .Size = 9: .Color = RGB(128, 128, 128): End With
    End If

    Dim rHdr As Long, r1 As Long, rN As Long, nP As Long
    nP = UBound(pos) + 1
    rHdr = 5: r1 = 6: rN = r1 + nP - 1
    WriteRow ws, rHdr, Array("SA/PO", U("D\u00f2ng xe"), U("NCC \u0111\u00e3 giao"), "GR (MB51)", "Gap (NCC - GR)", U("Tr\u1ea1ng th\u00e1i"), _
        U("S\u1ed1 m\u00e3 c\u1ea7u"), "Xe giao (NCC)", U("Xe \u0111\u00e3 GR"), U("B\u1ed9 c\u1ea7u"), U("Sheet chi ti\u1ebft"))
    FmtHeader ws, rHdr, NC

    Dim col() As Variant, boCol() As Variant, cars As Object, car As String, st As Variant, bo As Object
    Dim nOk As Long, nBad As Long, nWarnPO As Long, xeNCC As Double
    ReDim col(1 To nP, 1 To 2): ReDim boCol(1 To nP, 1 To 4)
    Set cars = NewDict()
    For i = 0 To nP - 1
        po = pos(i)
        If carOf Is Nothing Then
            car = ""
        Else
            car = DGetS(carOf, po): If car = "" Then car = NO_CAR
            If Not cars.Exists(car) Then cars.Add car, 1
        End If
        col(i + 1, 1) = ToCell(po): col(i + 1, 2) = car
        st = PoStatus(inN.Exists(po), inM.Exists(po), DGet(sumN, po) - DGet(sumM, po))
        Select Case st(1)
            Case "ok": nOk = nOk + 1
            Case "bad": nBad = nBad + 1
            Case Else: nWarnPO = nWarnPO + 1
        End Select
        Set bo = boOf(po)
        If cpc > 0 Then
            boCol(i + 1, 1) = bo("nCodes") & "/" & cpc: boCol(i + 1, 2) = bo("carsN"): boCol(i + 1, 3) = bo("carsM")
            boCol(i + 1, 4) = IIf(bo("ok"), U("\u2714 \u0110\u1ee7 b\u1ed9"), U("\u26a0 ") & bo("issues"))
            xeNCC = xeNCC + bo("carsN")
        Else
            boCol(i + 1, 1) = bo("nCodes"): boCol(i + 1, 2) = "": boCol(i + 1, 3) = "": boCol(i + 1, 4) = U("\u2014")
        End If
    Next
    ws.Range(ws.Cells(r1, 1), ws.Cells(rN, 2)).Value2 = col
    ws.Range(ws.Cells(r1, 7), ws.Cells(rN, 7)).NumberFormat = "@"   ' "4/4" khong bi hieu la ngay
    ws.Range(ws.Cells(r1, 7), ws.Cells(rN, 10)).Value2 = boCol
    SetSummaryFormulas ws, r1, rN

    For i = 0 To nP - 1
        If sheetOf.Exists(pos(i)) Then AddLink ws, ws.Cells(r1 + i, NC), sheetOf(pos(i)), U("\u2192 Xem chi ti\u1ebft ") & pos(i)
    Next
    If cpc > 0 Then
        ws.Cells(rN + 1, 8).Formula = "=SUM(H" & r1 & ":H" & rN & ")"
        ws.Cells(rN + 1, 9).Formula = "=SUM(I" & r1 & ":I" & rN & ")"
    End If
    With ws.Range(ws.Cells(rN + 1, 1), ws.Cells(rN + 1, NC))
        .Font.Bold = True: .Interior.Color = CLR_LBLUE
    End With
    ColorCell ws.Cells(3, 1), True: ColorCell ws.Cells(3, 2), False
    With ws.Cells(3, 3): .Interior.Color = CLR_WARN_BG: .Font.Color = CLR_WARN_FT: .Font.Bold = True: End With

    ' Mau A..F theo ky tu dau cua Trang thai (? xanh / ? do / ? vang) - tu doi khi so lieu doi
    Dim body As Range, main6 As Range, fc As Object, cf As Variant, x As Variant
    Set body = ws.Range(ws.Cells(r1, 1), ws.Cells(rN, NC - 1))
    body.Borders.LineStyle = xlContinuous: body.Borders.Color = CLR_GRID
    Set main6 = ws.Range(ws.Cells(r1, 1), ws.Cells(rN, 6))
    ws.Activate: ws.Cells(r1, 1).Select   ' cong thuc dieu kien tinh tuong doi theo o dang chon
    main6.FormatConditions.Delete
    For Each cf In Array(Array(U("\u2714"), CLR_GREEN_BG, CLR_GREEN_FT), Array(U("\u2718"), CLR_RED_BG, CLR_RED_FT), Array(U("\u26a0"), CLR_WARN_BG, CLR_WARN_FT))
        Set fc = main6.FormatConditions.Add(Type:=xlExpression, Formula1:="=LEFT($F" & r1 & ",1)=""" & cf(0) & """")
        fc.Interior.Color = cf(1): fc.Font.Color = cf(2)
    Next
    ' Mau G..J theo ket qua bo cau
    If cpc > 0 Then
        For i = 0 To nP - 1
            With ws.Range(ws.Cells(r1 + i, 7), ws.Cells(r1 + i, 10))
                If boOf(pos(i))("ok") Then
                    .Interior.Color = CLR_GREEN_BG: .Font.Color = CLR_GREEN_FT
                Else
                    .Interior.Color = CLR_WARN_BG: .Font.Color = CLR_WARN_FT
                End If
            End With
        Next
    End If
    ws.Range(ws.Cells(r1, 7), ws.Cells(rN, 9)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(rHdr, 1), ws.Cells(rN, NC)).AutoFilter   ' loc theo Dong xe / Trang thai / Bo cau
    If Not carOf Is Nothing Then WriteCarTable ws, cars.Keys, r1, rN
    AutoFitCols ws, NC
    ws.Columns(1).ColumnWidth = 16: ws.Columns(2).ColumnWidth = 12: ws.Columns(3).ColumnWidth = 15
    ws.Columns(4).ColumnWidth = 15: ws.Columns(5).ColumnWidth = 17: ws.Columns(7).ColumnWidth = 13
    ws.Columns(8).ColumnWidth = 15: ws.Columns(9).ColumnWidth = 13
    ws.Cells(1, 1).Select

    Dim res As Object: Set res = NewDict()
    res.Add "ws", ws: res("khop") = nOk: res("lech") = nBad: res("motBen") = nWarnPO: res("xeNCC") = xeNCC
    Set WriteSummary = res
End Function

' Bang tong theo dong xe (cot M..T) - toan cong thuc tren bang PO ben trai
Private Sub WriteCarTable(ws As Worksheet, ByVal carsArr As Variant, ByVal r1 As Long, ByVal rN As Long)
    Dim C0 As Long, NCt As Long, i As Long, r As Long, c As Long, t As Range, h As Range, hdr As Variant
    Dim B As String, Cc As String, D As String, F As String, H_ As String, key As String, rT As Long
    C0 = SUM_NC + 2: NCt = 8
    If IsArray(carsArr) Then If UBound(carsArr) >= 0 Then SortCars carsArr
    ws.Cells(4, C0).Value2 = U("T\u1ed4NG THEO D\u00d2NG XE")
    Set t = MergeRange(ws, 4, C0, 4, C0 + NCt - 1)
    t.Font.Bold = True: t.Font.Color = CLR_WHITE: t.Interior.Color = CLR_HDR_BG: t.HorizontalAlignment = xlCenter
    hdr = Array(U("D\u00f2ng xe"), U("S\u1ed1 PO"), "Xe giao", U("NCC \u0111\u00e3 giao"), "GR (MB51)", "Gap", U("\u2714 Kh\u1edbp"), U("C\u00f3 v\u1ea5n \u0111\u1ec1"))
    For i = 0 To UBound(hdr): ws.Cells(5, C0 + i).Value2 = hdr(i): Next
    Set h = ws.Range(ws.Cells(5, C0), ws.Cells(5, C0 + NCt - 1))
    h.Font.Bold = True: h.Font.Color = CLR_WHITE: h.Interior.Color = CLR_HDR2_BG: h.HorizontalAlignment = xlCenter
    h.Borders.LineStyle = xlContinuous: h.Borders.Color = CLR_WHITE
    B = "$B$" & r1 & ":$B$" & rN: Cc = "$C$" & r1 & ":$C$" & rN: D = "$D$" & r1 & ":$D$" & rN
    F = "$F$" & r1 & ":$F$" & rN: H_ = "$H$" & r1 & ":$H$" & rN
    Dim nCars As Long: nCars = 0
    If IsArray(carsArr) Then nCars = UBound(carsArr) + 1
    For i = 0 To nCars - 1
        r = 6 + i: key = "$" & CL(C0, 0) & r
        ws.Cells(r, C0).Value2 = carsArr(i)
        ws.Cells(r, C0 + 1).Formula = "=COUNTIF(" & B & "," & key & ")"
        ws.Cells(r, C0 + 2).Formula = "=SUMIFS(" & H_ & "," & B & "," & key & ")"
        ws.Cells(r, C0 + 3).Formula = "=SUMIFS(" & Cc & "," & B & "," & key & ")"
        ws.Cells(r, C0 + 4).Formula = "=SUMIFS(" & D & "," & B & "," & key & ")"
        ws.Cells(r, C0 + 5).Formula = "=" & CL(C0, 3) & r & "-" & CL(C0, 4) & r
        ws.Cells(r, C0 + 6).Formula = "=COUNTIFS(" & B & "," & key & "," & F & U(",""\u2714*"")")
        ws.Cells(r, C0 + 7).Formula = "=" & CL(C0, 1) & r & "-" & CL(C0, 6) & r
    Next
    rT = 6 + nCars
    ws.Cells(rT, C0).Value2 = U("T\u1ed4NG")
    For c = 1 To NCt - 1
        If c = 5 Then
            ws.Cells(rT, C0 + c).Formula = "=" & CL(C0, 3) & rT & "-" & CL(C0, 4) & rT
        Else
            ws.Cells(rT, C0 + c).Formula = "=SUM(" & CL(C0, c) & "6:" & CL(C0, c) & (rT - 1) & ")"
        End If
    Next
    With ws.Range(ws.Cells(rT, C0), ws.Cells(rT, C0 + NCt - 1)): .Font.Bold = True: .Interior.Color = CLR_LBLUE: End With
    For c = 0 To NCt - 1: ws.Columns(C0 + c).ColumnWidth = IIf(c = 0, 12, 13): Next
    ws.Columns(C0 - 1).ColumnWidth = 3
    If nCars = 0 Then Exit Sub
    Dim body As Range, g As Object, bdc As Object
    Set body = ws.Range(ws.Cells(6, C0), ws.Cells(rT - 1, C0 + NCt - 1))
    body.Borders.LineStyle = xlContinuous: body.Borders.Color = CLR_GRID
    ws.Cells(6, C0).Select
    body.FormatConditions.Delete
    Set g = body.FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & CL(C0, 7) & "6=0"): g.Interior.Color = CLR_GREEN_BG: g.Font.Color = CLR_GREEN_FT
    Set bdc = body.FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & CL(C0, 7) & "6>0"): bdc.Interior.Color = CLR_RED_BG: bdc.Font.Color = CLR_RED_FT
End Sub

Private Function CL(ByVal c0 As Long, ByVal k As Long) As String
    CL = ColLetter(c0 + k)
End Function

' Sap xep danh sach dong xe theo thu tu MODELS
Private Sub SortCars(arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = 1 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= 0
            If CarRank(arr(j)) - CarRank(tmp) > 0 Or (CarRank(arr(j)) = CarRank(tmp) And CmpVal(arr(j), tmp) > 0) Then
                arr(j + 1) = arr(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = tmp
    Next
End Sub

' Cong thuc Tong hop (dung lai khi xuat file) - dong r1..rN la cac PO, rN+1 la dong TONG
Private Sub SetSummaryFormulas(ws As Worksheet, ByVal r1 As Long, ByVal rN As Long)
    Dim a As String, e As String, rT As Long, f As String
    a = "$A" & r1: e = "E" & r1
    ws.Range(ws.Cells(r1, 3), ws.Cells(rN, 3)).Formula = "=SUMIFS(" & DATA_N & "!$D:$D," & DATA_N & "!$B:$B," & a & ")"
    ws.Range(ws.Cells(r1, 4), ws.Cells(rN, 4)).Formula = "=SUMIFS(" & DATA_M & "!$D:$D," & DATA_M & "!$B:$B," & a & ")"
    ws.Range(ws.Cells(r1, 5), ws.Cells(rN, 5)).Formula = "=C" & r1 & "-D" & r1
    ws.Range(ws.Cells(r1, 6), ws.Cells(rN, 6)).Formula = "=IF(COUNTIF(" & DATA_M & "!$B:$B," & a & U(")=0,""\u26a0 Ch\u01b0a GR (ch\u1ec9 c\u00f3 NCC)"",") & _
        "IF(COUNTIF(" & DATA_N & "!$B:$B," & a & U(")=0,""\u26a0 Ch\u1ec9 c\u00f3 GR (NCC kh\u00f4ng c\u00f3)"",") & _
        "IF(ROUND(" & e & U(",6)=0,""\u2714 Kh\u1edbp"",IF(") & e & U(">0,""\u2718 NCC th\u1eeba / GR thi\u1ebfu"",""\u2718 GR th\u1eeba / NCC thi\u1ebfu""))))")
    rT = rN + 1
    ws.Cells(rT, 1).Value2 = U("T\u1ed4NG")
    ws.Cells(rT, 3).Formula = "=SUM(C" & r1 & ":C" & rN & ")"
    ws.Cells(rT, 4).Formula = "=SUM(D" & r1 & ":D" & rN & ")"
    ws.Cells(rT, 5).Formula = "=C" & rT & "-D" & rT
    f = "F" & r1 & ":F" & rN
    ws.Cells(3, 1).Formula = U("=""\u2714 Kh\u1edbp: ""&COUNTIF(") & f & U(",""\u2714*"")")
    ws.Cells(3, 2).Formula = U("=""\u2718 L\u1ec7ch: ""&COUNTIF(") & f & U(",""\u2718*"")")
    ws.Cells(3, 3).Formula = U("=""\u26a0 M\u1ed9t b\u00ean: ""&COUNTIF(") & f & U(",""\u26a0*"")")
End Sub

'========================= SHEET CANH BAO DU LIEU =========================
Private Sub WriteWarnings(wb As Workbook, afterWs As Worksheet, warns As Collection)
    Dim ws As Worksheet, tt As Range, data() As Variant, i As Long, w As Variant, body As Range
    Set ws = AddSheetAfter(wb, afterWs)
    ws.Name = SafeSheetName(wb, WARN_NAME)
    ws.Cells(1, 1).Value2 = U("C\u1ea2NH B\u00c1O D\u1eee LI\u1ec6U \u2014 ") & warns.Count & U(" d\u00f2ng c\u00f3 v\u1ea5n \u0111\u1ec1 (ki\u1ec3m tra l\u1ea1i file g\u1ed1c)")
    Set tt = MergeRange(ws, 1, 1, 1, 7)
    tt.Font.Bold = True: tt.Font.Size = 13: tt.Font.Color = CLR_WHITE
    tt.Interior.Color = RGB(156, 87, 0): tt.HorizontalAlignment = xlCenter
    AddLink ws, ws.Cells(2, 1), SUM_NAME, U("\u2190 V\u1ec1 ") & SUM_NAME
    ws.Range("C:F").NumberFormat = "@"
    ReDim data(1 To warns.Count + 1, 1 To 7)
    data(1, 1) = "File": data(1, 2) = U("D\u00f2ng trong file g\u1ed1c"): data(1, 3) = "Material": data(1, 4) = "PO/SA"
    data(1, 5) = U("S\u1ed1 BB"): data(1, 6) = U("S\u1ed1 l\u01b0\u1ee3ng (g\u1ed1c)"): data(1, 7) = U("V\u1ea5n \u0111\u1ec1")
    i = 1
    For Each w In warns
        i = i + 1
        data(i, 1) = w(0): data(i, 2) = w(1): data(i, 3) = w(2): data(i, 4) = w(3): data(i, 5) = w(4): data(i, 6) = w(5): data(i, 7) = w(6)
    Next
    ws.Range(ws.Cells(4, 1), ws.Cells(4 + warns.Count, 7)).Value2 = data
    FmtHeader ws, 4, 7
    Set body = ws.Range(ws.Cells(5, 1), ws.Cells(4 + warns.Count, 7))
    body.Interior.Color = CLR_WARN_BG: body.Font.Color = CLR_WARN_FT
    body.Borders.LineStyle = xlContinuous: body.Borders.Color = CLR_GRID
    ws.Range(ws.Cells(4, 1), ws.Cells(4 + warns.Count, 7)).AutoFilter
    AutoFitCols ws, 7: ws.Columns(1).ColumnWidth = 10
End Sub

'========================= XUAT KET QUA RA FILE =========================
Private Sub ExportResult(ByVal path As String)
    Dim wb As Workbook, nb As Workbook, sh As Worksheet, names As New Collection, nm As Variant, s As Worksheet
    Dim r As Long, lk As Variant, i As Long
    Set wb = ThisWorkbook
    For Each sh In wb.Worksheets
        If IsToolSheet(sh) Then names.Add sh.Name
    Next
    If names.Count = 0 Then Err.Raise vbObjectError + 10, , U("Ch\u01b0a c\u00f3 k\u1ebft qu\u1ea3 \u0111\u1ec3 l\u01b0u.")
    Application.DisplayAlerts = False
    For Each nm In names
        Set sh = wb.Worksheets(nm)
        If nb Is Nothing Then
            sh.Copy: Set nb = ActiveWorkbook
        Else
            sh.Copy After:=nb.Sheets(nb.Sheets.Count)
        End If
    Next
    ' Cong thuc copy sang bi tro ve file tool -> ghi lai de tro vao sheet du lieu trong file moi
    Set s = GetSheet(nb, SUM_NAME)
    If Not s Is Nothing Then
        r = 6
        Do While TrimS(s.Cells(r, 1).Value2) <> U("T\u1ed4NG") And r < 1048570: r = r + 1: Loop
        If r > 6 Then SetSummaryFormulas s, 6, r - 1
        s.Cells(2, SUM_NC).Clear   ' bo link "Ve Control" (file moi khong co Control)
    End If
    lk = nb.LinkSources(xlExcelLinks)
    If Not IsEmpty(lk) Then
        For i = LBound(lk) To UBound(lk): nb.BreakLink Name:=lk(i), Type:=xlLinkTypeExcelLinks: Next
    End If
    nb.Sheets(1).Activate
    nb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    nb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub

Private Sub AddLink(ws As Worksheet, cell As Range, ByVal sheetName As String, ByVal text As String)
    On Error Resume Next
    ws.Hyperlinks.Add Anchor:=cell, Address:="", SubAddress:="'" & sheetName & "'!A1", TextToDisplay:=text
    If Err.Number <> 0 Then cell.Value2 = text
    On Error GoTo 0
End Sub

'========================= HAM PHU TRO =========================
Private Function GetSheet(wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = nm Then Set GetSheet = ws: Exit Function
    Next
End Function

' Sheet do tool tao: ten co dinh, hoac sheet chi tiet (A1 bat dau bang DETAIL_TAG)
Private Function IsToolSheet(sh As Worksheet) As Boolean
    Dim a1 As String
    If sh.Name = SUM_NAME Or sh.Name = WARN_NAME Or sh.Name = DATA_N Or sh.Name = DATA_M Then IsToolSheet = True: Exit Function
    On Error Resume Next
    a1 = CStr(sh.Cells(1, 1).Value2)
    On Error GoTo 0
    IsToolSheet = (Left$(a1, Len(DETAIL_TAG)) = DETAIL_TAG)
End Function

' Chi xoa sheet do tool tao - sheet rieng cua nguoi dung giu nguyen
Private Sub ClearResultSheets(wb As Workbook)
    Dim names As New Collection, sh As Worksheet, nm As Variant
    For Each sh In wb.Worksheets
        If sh.Name <> CTRL And IsToolSheet(sh) Then names.Add sh.Name
    Next
    Application.DisplayAlerts = False
    On Error Resume Next
    For Each nm In names: wb.Worksheets(nm).Delete: Next
    On Error GoTo 0
End Sub

Private Function ColIndex(ByVal v As Variant) As Long
    Dim s As String, n As Long, i As Long
    s = UCase$(TrimS(v))
    If s = "" Or s = "-" Then Exit Function
    If IsAllDigits(s) Then ColIndex = CLng(s): Exit Function
    If Not (s Like "[A-Z]" Or s Like "[A-Z][A-Z]" Or s Like "[A-Z][A-Z][A-Z]") Then Exit Function
    For i = 1 To Len(s): n = n * 26 + (Asc(Mid$(s, i, 1)) - 64): Next
    ColIndex = n
End Function

Private Function ColLetter(ByVal n As Long) As String
    Dim s As String, m As Long
    Do While n > 0
        m = (n - 1) Mod 26: s = Chr(65 + m) & s: n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Function CleanKey(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        If v = Int(v) And Abs(v) < 1E+15 Then CleanKey = CStr(CDec(v)) Else CleanKey = CStr(v)
    Else
        CleanKey = TrimS(v)
    End If
End Function

Private Function ToNum(ByVal v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then ToNum = CDbl(v)
End Function

Private Function Round6(ByVal x As Double) As Double
    Round6 = Round(x, 6)
End Function

Private Function TrimS(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    TrimS = Trim$(CStr(v))
End Function

Private Function SafeSheetName(wb As Workbook, ByVal s As String) As String
    Dim x As String, base As String, i As Long, ch As Variant
    x = s
    For Each ch In Array("\", "/", "?", "*", "[", "]", ":"): x = Replace(x, ch, "_"): Next
    If Len(x) > 31 Then x = Left$(x, 31)
    If x = "" Then x = "Sheet"
    base = x: i = 1
    Do While Not GetSheet(wb, x) Is Nothing
        x = Left$(base, 28) & "_" & i: i = i + 1
    Loop
    SafeSheetName = x
End Function

Private Function PickFile(ByVal title As String) As String
    Dim f As Variant
    If gSilent Then Exit Function
    f = Application.GetOpenFilename(FileFilter:="Excel (*.xls;*.xlsx;*.xlsm;*.csv),*.xls;*.xlsx;*.xlsm;*.csv", Title:=title)
    If VarType(f) = vbBoolean Then Exit Function
    PickFile = CStr(f)
End Function

Private Sub WriteRow(ws As Worksheet, ByVal r As Long, arr As Variant)
    Dim i As Long
    For i = 0 To UBound(arr): ws.Cells(r, i + 1).Value2 = arr(i): Next
End Sub

Private Function MergeRange(ws As Worksheet, ByVal r1 As Long, ByVal c1 As Long, ByVal r2 As Long, ByVal c2 As Long) As Range
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    On Error Resume Next
    rng.Merge
    On Error GoTo 0
    Set MergeRange = rng
End Function

Private Sub FmtHeader(ws As Worksheet, ByVal r As Long, ByVal nCol As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, nCol))
        .Font.Bold = True: .Font.Color = CLR_WHITE: .Interior.Color = CLR_HDR2_BG
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Color = CLR_WHITE
    End With
End Sub

Private Sub ColorCell(cell As Range, ByVal ok As Boolean)
    cell.Interior.Color = IIf(ok, CLR_GREEN_BG, CLR_RED_BG)
    cell.Font.Color = IIf(ok, CLR_GREEN_FT, CLR_RED_FT)
    cell.Font.Bold = True
End Sub

' kind: "ok" xanh / "bad" do / "warn" vang
Private Sub ColorRowKind(ws As Worksheet, ByVal r As Long, ByVal nCol As Long, ByVal kind As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, nCol))
        Select Case kind
            Case "ok": .Interior.Color = CLR_GREEN_BG: .Font.Color = CLR_GREEN_FT
            Case "warn": .Interior.Color = CLR_WARN_BG: .Font.Color = CLR_WARN_FT
            Case Else: .Interior.Color = CLR_RED_BG: .Font.Color = CLR_RED_FT
        End Select
        .Borders.LineStyle = xlContinuous: .Borders.Color = CLR_GRID
    End With
End Sub

Private Sub AutoFitCols(ws As Worksheet, ByVal nCol As Long)
    Dim c As Long
    On Error Resume Next
    For c = 1 To nCol: ws.Columns(c).AutoFit: Next
    On Error GoTo 0
End Sub

Private Sub Trace(ByVal s As String)
    Dim f As Integer
    If Not TRACE_ON Then Exit Sub
    f = FreeFile
    Open Environ$("TEMP") & "\doisoat_trace.txt" For Append As #f
    Print #f, Format(Now, "hh:nn:ss") & "." & Format((Timer * 1000) Mod 1000, "000") & " " & s
    Close #f
End Sub

Private Sub Msg(ByVal s As String)
    Dim w As Worksheet, pr As Boolean
    Trace "MSG: " & Left$(s, 200)
    If gSilent Then
        Set w = GetSheet(ThisWorkbook, CTRL)
        If Not w Is Nothing Then
            pr = w.ProtectContents
            If pr Then w.Unprotect PASS
            If TrimS(w.Range("Z1").Value2) = "" Then w.Range("Z1").Value2 = s Else w.Range("Z1").Value2 = w.Range("Z1").Value2 & " || " & s
            If pr Then ProtectSheet w
        End If
    Else
        Dim cap As String: cap = U("Tool \u0110\u1ed1i So\u00e1t NCC <-> MB51")
        MessageBoxW Application.hWnd, StrPtr(s), StrPtr(cap), &H40 Or &H40000
    End If
End Sub

