Attribute VB_Name = "Module1"
Option Explicit

Const ROW_TITLE    As Integer = 1
Const ROW_DAY_HDR  As Integer = 2
Const ROW_DATE     As Integer = 3
Const ROW_YOTEI    As Integer = 4
Const PERIOD_START As Integer = 5
Const ROW_AFTER    As Integer = 19
Const ROW_DIV      As Integer = 20
Const ROW_SUM_HDR  As Integer = 21
Const ROW_KOMA     As Integer = 22
Const ROW_IMPL     As Integer = 23
Const ROW_PLAN     As Integer = 24
Const ROW_CUM      As Integer = 25
Const COL_MON_SCI As Integer = 2
Const COL_TUE_SCI As Integer = 5
Const COL_WED_SCI As Integer = 8
Const COL_THU_SCI As Integer = 11
Const COL_FRI_SCI As Integer = 14
Const COL_MON_CHG As Integer = 4
Const COL_TUE_CHG As Integer = 7
Const COL_WED_CHG As Integer = 10
Const COL_THU_CHG As Integer = 13
Const COL_FRI_CHG As Integer = 16
' ----
' 書式設定ヘルパー
' ----
Private Sub SF(rng As Range, v As String, b As Boolean, sz As Integer, fc As Long, BG As Long, ha As Integer, va As Integer)
    rng.Value = v: rng.Font.Name = "メイリオ": rng.Font.Bold = b
    rng.Font.Size = sz: rng.Font.Color = fc: rng.Interior.Color = BG
    rng.HorizontalAlignment = ha: rng.VerticalAlignment = va
End Sub
' ----
' 【1】設定完了 ─ 全シート・ボタンを自動生成
' ----
Sub 設定を完了()
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Sheets("基本設定")
    On Error Resume Next: cfg.Unprotect "": On Error GoTo 0
    If cfg.Range("B2").Value = "" Then MsgBox "年度未入力", vbExclamation: Exit Sub
    If cfg.Range("B6").Value = "" Then MsgBox "クラス未入力", vbExclamation: Exit Sub
    If MsgBox("システムを構築します。よろしいですか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Application.ScreenUpdating = False
    Call 自動生成シートを削除
    Call 時間割マスターを作成
    Call 週案テンプレートを作成
    Call 基本設定にボタンを追加
    Dim ul As Variant: cfg.Cells.Locked = True
    For Each ul In Array("B2:B4", "B6:B13", "B16:B20", "B23:B35", "B38:B43"): cfg.Range(CStr(ul)).Locked = False: Next ul

    休業設定を準備 cfg
    cfg.Protect ""
    Application.ScreenUpdating = True
    MsgBox "セットアップ完了！", vbInformation
End Sub

Private Sub 自動生成シートを削除()
    Application.DisplayAlerts = False
    Dim ws As Worksheet, n() As String, c As Integer: c = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "基本設定" And ws.Name <> "VBAコード" Then
            ReDim Preserve n(c): n(c) = ws.Name: c = c + 1
        End If
    Next ws
    Dim i As Integer
    For i = 0 To c - 1: ThisWorkbook.Sheets(n(i)).Delete: Next i
    Application.DisplayAlerts = True
End Sub
' ----
' 【2】時間割マスターを作成
' ----
Private Sub 時間割マスターを作成()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("基本設定"))
    ws.Name = "時間割マスター": ws.Tab.Color = RGB(30, 132, 73)
    ws.Range("A1:F1").Merge
    With ws.Range("A1")
        .Value = "時間割マスター" & ChrW(12288) & "※プルダウンからクラスを選択"
        .Font.Name = "メイリオ": .Font.Bold = True: .Font.Size = 9
        .Font.Color = vbWhite: .Interior.Color = RGB(30, 132, 73)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 18
    ws.Cells(2, 1).Value = "校時"
    Dim days As Variant: days = Array("月", "火", "水", "木", "金")
    Dim d As Integer
    For d = 0 To 4: ws.Cells(2, d + 2).Value = days(d): Next d
    With ws.Range("A2:F2")
        .Font.Name = "メイリオ": .Font.Bold = True: .Font.Size = 10
        .Font.Color = vbWhite: .Interior.Color = RGB(30, 132, 73)
        .HorizontalAlignment = xlCenter: .RowHeight = 20
    End With
    Dim p As Integer
    For p = 1 To 7
        ws.Cells(p + 2, 1).Value = p
        With ws.Cells(p + 2, 1)
            .Font.Name = "メイリオ": .Font.Bold = True: .Font.Size = 10
            .Font.Color = vbWhite: .Interior.Color = RGB(46, 134, 171)
            .HorizontalAlignment = xlCenter: .RowHeight = 22
        End With
        For d = 2 To 6
            With ws.Cells(p + 2, d)
                .Font.Name = "メイリオ": .Font.Bold = True: .Font.Size = 10
                .Font.Color = RGB(26, 35, 126): .Interior.Color = RGB(235, 245, 251)
                .HorizontalAlignment = xlCenter: .NumberFormat = "@"
            End With
        Next d
    Next p
    With ws.Range("A2:F9").Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(187, 187, 187)
    End With
    ws.Columns("A").ColumnWidth = 6
    For d = 2 To 6: ws.Columns(d).ColumnWidth = 10: Next d
    With ws.Range("B3:F9").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="='基本設定'!$B$6:$B$20"
        .IgnoreBlank = True: .ShowError = False
    End With
    ws.Cells.Locked = True: ws.Range("B3:F9").Locked = False: ws.Protect ""
End Sub
' ----
' 【3】週案テンプレートを作成（v35レイアウト完全再現）
' ----
Private Sub 週案テンプレートを作成()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("時間割マスター"))
    ws.Name = "週案_テンプレート": ws.Tab.Color = RGB(192, 57, 43)
    Dim yr As String: yr = CStr(ThisWorkbook.Sheets("基本設定").Range("B2").Value)
    With ws.PageSetup
        .PaperSize = xlPaperA4: .Orientation = xlPortrait
        .FitToPagesWide = 1: .FitToPagesTall = 1: .Zoom = False
        Dim mg As Double: mg = Application.InchesToPoints(20 / 25.4)
        .LeftMargin = mg: .RightMargin = mg: .TopMargin = mg: .BottomMargin = mg
        .HeaderMargin = 0: .FooterMargin = 0
        .PrintArea = "A1:P25": .CenterHorizontally = True
    End With
    Call テンプレートレイアウト(ws, yr)
    ws.Columns("A").ColumnWidth = 4.91
    Dim c As Integer
    For c = 2 To 16: ws.Columns(c).ColumnWidth = 5.09: Next c
    ws.Rows(1).RowHeight = 50
    ws.Rows(2).RowHeight = 15.5
    ws.Rows(3).RowHeight = 15#
    ws.Rows(4).RowHeight = 55
    Dim p As Integer
    For p = 0 To 6
        ws.Rows(5 + p * 2).RowHeight = 18#
        ws.Rows(6 + p * 2).RowHeight = 62#
    Next p
    ws.Rows(19).RowHeight = 55
    ws.Rows(20).RowHeight = 0
    Dim r As Integer
    For r = 21 To 25: ws.Rows(r).RowHeight = 13.5: Next r
    ws.Cells.Locked = True: ws.Range("B3:P19").Locked = False: ws.Range("B25:N25").Locked = False: ws.Protect ""
End Sub

Private Sub テンプレートレイアウト(ws As Worksheet, yr As String)
    Dim dS(4) As Integer, dE(4) As Integer, dc(4) As Integer
    dS(0) = 2: dE(0) = 3: dc(0) = 4: dS(1) = 5: dE(1) = 6: dc(1) = 7
    dS(2) = 8: dE(2) = 9: dc(2) = 10: dS(3) = 11: dE(3) = 12: dc(3) = 13
    dS(4) = 14: dE(4) = 15: dc(4) = 16
    Dim dn(4) As String
    dn(0) = "月": dn(1) = "火": dn(2) = "水": dn(3) = "木": dn(4) = "金"
    Dim BG(4) As Long
    BG(0) = RGB(219, 234, 254): BG(1) = RGB(254, 249, 195): BG(2) = RGB(220, 252, 231)
    BG(3) = RGB(255, 228, 230): BG(4) = RGB(243, 232, 255)
    Dim cn As Long: cn = RGB(47, 85, 151)
    Dim CD As Long: CD = RGB(84, 110, 122)
    ws.Range("A1:G1").Merge
    Call SF(ws.Range("A1"), yr & "年度" & ChrW(12288) & "第" & ChrW(12288) & ChrW(12288) & "週（" & ChrW(12288) & ChrW(12288) & "月" & ChrW(12288) & ChrW(12288) & "日?" & ChrW(12288) & ChrW(12288) & "月" & ChrW(12288) & ChrW(12288) & "日）" & ChrW(12288) & "進度予定表", True, 11, vbWhite, cn, xlCenter, xlCenter)

    Call 決裁欄設定(ws)
    ws.Range("A2:A3").Merge
    Call SF(ws.Range("A2"), "時", True, 9, vbWhite, CD, xlCenter, xlCenter)
    Dim d As Integer, p As Integer, sR As Integer, cr As Integer
    For d = 0 To 4
        ws.Range(ws.Cells(2, dS(d)), ws.Cells(2, dc(d))).Merge
        Call SF(ws.Range(ws.Cells(2, dS(d)), ws.Cells(2, dc(d))), dn(d), True, 11, RGB(0, 0, 0), BG(d), xlCenter, xlCenter)
        ws.Range(ws.Cells(3, dS(d)), ws.Cells(3, dc(d))).Merge
        With ws.Range(ws.Cells(3, dS(d)), ws.Cells(3, dc(d)))
            .Value = ChrW(12288) & "/" & ChrW(12288) & "（" & dn(d) & "）"
            .NumberFormat = "@": .Font.Name = "メイリオ": .Font.Size = 8
            .Font.Color = RGB(85, 85, 85): .Interior.Color = BG(d)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
    Next d
    Call SF(ws.Range("A4"), "予定", True, 7, vbWhite, CD, xlCenter, xlCenter)
    For d = 0 To 4
        ws.Range(ws.Cells(4, dS(d)), ws.Cells(4, dc(d))).Merge
        With ws.Range(ws.Cells(4, dS(d)), ws.Cells(4, dc(d)))
            .Interior.Color = BG(d): .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop: .WrapText = True
            .Font.Name = "メイリオ": .Font.Size = 8
        End With
    Next d
    For p = 1 To 7
        sR = 5 + (p - 1) * 2: cr = sR + 1
        ws.Range(ws.Cells(sR, 1), ws.Cells(cr, 1)).Merge
        Call SF(ws.Range(ws.Cells(sR, 1), ws.Cells(cr, 1)), CStr(p), True, 12, vbWhite, CD, xlCenter, xlCenter)
        For d = 0 To 4
            ws.Range(ws.Cells(sR, dS(d)), ws.Cells(sR, dE(d))).Merge
            With ws.Range(ws.Cells(sR, dS(d)), ws.Cells(sR, dE(d)))
                .Font.Name = "メイリオ": .Font.Bold = True: .Font.Size = 9
                .Font.Color = RGB(26, 35, 126): .Interior.Color = RGB(227, 242, 253)
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .NumberFormat = "@"
            End With
            With ws.Cells(sR, dc(d))
                .Font.Name = "メイリオ": .Font.Size = 7: .Font.Color = RGB(85, 85, 85)
                .Interior.Color = RGB(191, 219, 254)
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            End With
            ws.Range(ws.Cells(cr, dS(d)), ws.Cells(cr, dc(d))).Merge
            With ws.Range(ws.Cells(cr, dS(d)), ws.Cells(cr, dc(d)))
                .Font.Name = "メイリオ": .Font.Size = 8: .Font.Color = RGB(33, 33, 33)
                .Interior.Color = vbWhite: .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop: .WrapText = True
            End With
        Next d
    Next p
    Call SF(ws.Range("A19"), "授業後", True, 7, vbWhite, RGB(69, 90, 100), xlCenter, xlCenter)
    For d = 0 To 4
        ws.Range(ws.Cells(19, dS(d)), ws.Cells(19, dc(d))).Merge
        With ws.Range(ws.Cells(19, dS(d)), ws.Cells(19, dc(d)))
            .Interior.Color = RGB(236, 239, 241): .Font.Name = "メイリオ": .Font.Size = 8
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlTop: .WrapText = True
        End With
    Next d
    Dim ci As Integer
    For ci = 1 To 16: ws.Cells(20, ci).Interior.Color = cn: Next ci
    Call 集計表を構築(ws)
    Call 罫線を追加(ws)
    For p = 1 To 7
        Dim dvR As Integer: dvR = 5 + (p - 1) * 2
        For d = 0 To 4
            With ws.Cells(dvR, dS(d)).Validation
                .Delete
                .Add Type:=xlValidateList, Formula1:="='基本設定'!$B$6:$B$20"
                .IgnoreBlank = True: .ShowError = False
            End With
        Next d
    Next p
    For d = 0 To 4
        With ws.Cells(19, dS(d)).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="='基本設定'!$B$6:$B$20"
            .IgnoreBlank = True: .ShowError = False
        End With
    Next d
    Dim chgList As String
    chgList = ChrW(215) & ",月1,月2,月3,月4,月5,月6,月7,火1,火2,火3,火4,火5,火6,火7,水1,水2,水3,水4,水5,水6,水7,木1,木2,木3,木4,木5,木6,木7,金1,金2,金3,金4,金5,金6,金7"
    For p = 1 To 7
        dvR = 5 + (p - 1) * 2
        For d = 0 To 4
            With ws.Cells(dvR, dc(d)).Validation
                .Delete
                .Add Type:=xlValidateList, Formula1:=chgList
                .IgnoreBlank = True: .ShowError = False
            End With
        Next d
    Next p
End Sub
' ----
' 【4】集計表を構築（R21-R25）
' ----
Private Sub 集計表を構築(ws As Worksheet)
    Dim cn As Long: cn = RGB(47, 85, 151)
    Dim CT As Long: CT = RGB(191, 54, 12)
    Call SF(ws.Cells(ROW_SUM_HDR, 1), "授業", True, 8, vbWhite, cn, xlCenter, xlCenter)
    Dim col As Integer
    For col = 2 To 14
        Dim sR As Integer: If col <= 9 Then sR = col + 4 Else sR = col + 6
        ws.Cells(ROW_SUM_HDR, col).Formula = "=IF(基本設定!B" & sR & "="""","""",基本設定!B" & sR & ")"
        With ws.Cells(ROW_SUM_HDR, col)
            .Font.Name = "メイリオ": .Font.Bold = True: .Font.Size = 8
            .Font.Color = vbWhite: .Interior.Color = cn
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
    Next col
    ws.Cells(ROW_SUM_HDR, 15).Interior.Color = RGB(221, 221, 221)
    Call SF(ws.Cells(ROW_SUM_HDR, 16), "計", True, 9, vbWhite, CT, xlCenter, xlCenter)
    With ws.Cells(ROW_KOMA, 1)
        .Value = "週コマ数": .Font.Name = "メイリオ": .Font.Size = 7
        .Font.Color = RGB(26, 35, 126): .Interior.Color = RGB(227, 242, 253)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .ShrinkToFit = True
    End With
    Dim hdr As String
    For col = 2 To 14
        hdr = ws.Cells(ROW_SUM_HDR, col).Address(False, False)
        ws.Cells(ROW_KOMA, col).Formula = "=IF(" & hdr & "="""","""",IFERROR(VALUE(VLOOKUP(" & hdr & ",基本設定!$A$23:$B$35,2,FALSE())),""""))"
        With ws.Cells(ROW_KOMA, col)
            .Font.Name = "メイリオ": .Font.Size = 9
            .Interior.Color = RGB(227, 242, 253): .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next col
    ws.Cells(ROW_KOMA, 15).Interior.Color = RGB(238, 238, 238)
    ws.Cells(ROW_KOMA, 16).Formula = "=IF(SUM(B" & ROW_KOMA & ":N" & ROW_KOMA & ")=0,"""",SUM(B" & ROW_KOMA & ":N" & ROW_KOMA & "))"
    With ws.Cells(ROW_KOMA, 16)
        .Interior.Color = RGB(255, 243, 224): .Font.Name = "メイリオ"
        .Font.Color = CT: .Font.Bold = True: .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(ROW_IMPL, 1), ws.Cells(ROW_PLAN, 1)).Merge
    With ws.Cells(ROW_IMPL, 1)
        .Value = "実施/予定": .Font.Name = "メイリオ": .Font.Size = 7
        .Font.Color = RGB(0, 0, 0): .Interior.Color = RGB(245, 245, 245)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .ShrinkToFit = True
    End With
    Dim sciR As Variant: sciR = Array(5, 7, 9, 11, 13, 15, 17)
    Dim dS(4) As Integer, dE(4) As Integer, dc(4) As Integer
    dS(0) = 2: dE(0) = 3: dS(1) = 5: dE(1) = 6: dS(2) = 8: dE(2) = 9: dS(3) = 11: dE(3) = 12: dS(4) = 14: dE(4) = 15
    dc(0) = 4: dc(1) = 7: dc(2) = 10: dc(3) = 13: dc(4) = 16
    Dim r As Integer, dd As Integer
    For col = 2 To 14
        hdr = ws.Cells(ROW_SUM_HDR, col).Address(False, False)
        Dim pp As String: pp = ""
        Dim mp As String: mp = ""
        For r = 0 To UBound(sciR)
            For dd = 0 To 4
                Dim sAddr As String: sAddr = ws.Cells(sciR(r), dS(dd)).Address(False, False)
                Dim eAddr As String: eAddr = ws.Cells(sciR(r), dE(dd)).Address(False, False)
                Dim cAddr As String: cAddr = ws.Cells(sciR(r), dc(dd)).Address(False, False)
                If pp <> "" Then pp = pp & "+"
                pp = pp & "COUNTIF(" & sAddr & ":" & eAddr & "," & hdr & ")"
                pp = pp & "+COUNTIF(" & sAddr & ":" & eAddr & ",""*,""&" & hdr & "&"",*"")"
                If mp <> "" Then mp = mp & "+"
                mp = mp & "SUMPRODUCT((" & sAddr & ":" & eAddr & "=" & hdr & ")*(" & cAddr & "=""×""))"
            Next dd
        Next r
        ws.Cells(ROW_PLAN, col).Formula = "=IF(" & hdr & "="""","""",IF(" & pp & "=0,""""," & pp & "))"
        With ws.Cells(ROW_PLAN, col)
            .Font.Name = "メイリオ": .Font.Size = 9: .Font.Color = RGB(121, 85, 0)
            .Interior.Color = RGB(255, 249, 196)
            .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
        End With
        Dim pRef As String: pRef = ws.Cells(ROW_PLAN, col).Address(False, False)
        ws.Cells(ROW_IMPL, col).Formula = "=IF(" & hdr & "="""","""",IF(" & pRef & "="""",""""," & pRef & "-(" & mp & ")))"
        With ws.Cells(ROW_IMPL, col)
            .Font.Name = "メイリオ": .Font.Size = 9: .Font.Bold = True
            .Font.Color = RGB(27, 94, 32): .Interior.Color = vbWhite
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
    Next col
    ws.Cells(ROW_IMPL, 15).Interior.Color = RGB(238, 238, 238)
    ws.Cells(ROW_PLAN, 15).Interior.Color = RGB(238, 238, 238)
    ws.Cells(ROW_IMPL, 16).Formula = "=IF(SUM(B" & ROW_IMPL & ":N" & ROW_IMPL & ")=0,"""",SUM(B" & ROW_IMPL & ":N" & ROW_IMPL & "))"
    ws.Cells(ROW_PLAN, 16).Formula = "=IF(SUM(B" & ROW_PLAN & ":N" & ROW_PLAN & ")=0,"""",SUM(B" & ROW_PLAN & ":N" & ROW_PLAN & "))"
    With ws.Cells(ROW_IMPL, 16)
        .Interior.Color = RGB(255, 243, 224): .Font.Name = "メイリオ"
        .Font.Color = CT: .Font.Bold = True: .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    With ws.Cells(ROW_PLAN, 16)
        .Interior.Color = RGB(255, 243, 224): .Font.Name = "メイリオ"
        .Font.Color = CT: .Font.Bold = True: .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    With ws.Cells(ROW_CUM, 1)
        .Value = "累" & ChrW(12288) & "計": .Font.Name = "メイリオ": .Font.Bold = True: .Font.Size = 7
        .Font.Color = RGB(136, 14, 79): .Interior.Color = RGB(252, 228, 236)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .ShrinkToFit = True
    End With
    For col = 2 To 14
        With ws.Cells(ROW_CUM, col)
            .Interior.Color = RGB(252, 228, 236): .Font.Name = "メイリオ"
            .Font.Size = 9: .Font.Bold = True: .Font.Color = RGB(136, 14, 79)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
    Next col
    ws.Cells(ROW_CUM, 15).Interior.Color = RGB(238, 238, 238)
    ws.Cells(ROW_CUM, 16).Formula = "=IF(SUM(B" & ROW_CUM & ":N" & ROW_CUM & ")=0,"""",SUM(B" & ROW_CUM & ":N" & ROW_CUM & "))"
    With ws.Cells(ROW_CUM, 16)
        .Interior.Color = RGB(255, 243, 224): .Font.Name = "メイリオ"
        .Font.Color = CT: .Font.Bold = True: .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Call 集計補正(ws)
End Sub

Private Sub 罫線を追加(ws As Worksheet)
    Dim r As Integer, c As Integer
    For r = 2 To 20
        For c = 1 To 16
            With ws.Cells(r, c).Borders
                .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(187, 187, 187)
            End With
        Next c
    Next r
    For r = ROW_SUM_HDR To ROW_CUM
        For c = 1 To 16
            With ws.Cells(r, c).Borders
                .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(187, 187, 187)
            End With
        Next c
    Next r
End Sub
' ----
' 【5】基本設定にボタンを追加
' ----
Sub 基本設定にボタンを追加()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("基本設定")
    On Error Resume Next: ws.Unprotect "": On Error GoTo 0
    Dim shp As Shape, n() As String, c As Integer: c = 0
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl And Left(shp.Name, 4) = "Auto" Then ReDim Preserve n(c): n(c) = shp.Name: c = c + 1
    Next shp
    Dim i As Integer
    For i = 0 To c - 1: ws.Shapes(n(i)).Delete: Next i
    Dim bL As Double: bL = ws.Columns("E").Left + 5
    Dim d(6, 1) As String
    d(0, 0) = "新しい週を追加": d(0, 1) = "ボタン_新しい週"
    d(1, 0) = "1学期PDF出力": d(1, 1) = "ボタン_1学期PDF"
    d(2, 0) = "2学期PDF出力": d(2, 1) = "ボタン_2学期PDF"
    d(3, 0) = "3学期PDF出力": d(3, 1) = "ボタン_3学期PDF"
    d(4, 0) = "次年度ファイルを作成": d(4, 1) = "ボタン_次年度"
    d(5, 0) = "別ファイルからインポート": d(5, 1) = "ボタン_旧ファイル移行"
    d(6, 0) = "全データ消去": d(6, 1) = "ボタン_全データ消去"
    For i = 0 To 6
        Dim btn As Shape
        Set btn = ws.Shapes.AddFormControl(xlButtonControl, bL, ws.Rows(2 + i * 2).Top + 3, 120, 22)
        With btn: .Name = "Auto_" & i: .TextFrame.Characters.Text = d(i, 0)
        .TextFrame.Characters.Font.Size = 9: .TextFrame.Characters.Font.Name = "メイリオ"
        .OLEFormat.Object.OnAction = d(i, 1): End With
    Next i
    ws.Protect ""
End Sub
' ----
' 【6】新しい週のシートを作成
' ----
Sub 新しい週を追加()
    Dim yr As String: yr = CStr(ThisWorkbook.Sheets("基本設定").Range("B2").Value)
    Dim wn As Integer: wn = 次の週番号()
    Dim md As String: md = 次の月曜日()
    If md = "" Then
        md = InputBox("月曜(例:4/7)", "入力")
        If md = "" Then Exit Sub
    End If
    Dim a As Integer
    a = MsgBox("第" & Format(wn, "00") & "週（" & md & ChrW(&HFF5E) & "）を追加します。" & vbLf & "よろしいですか？（いいえ＝修正）", vbYesNoCancel + vbQuestion)
    If a = vbCancel Then Exit Sub
    If a = vbNo Then
        Dim t As String
        t = InputBox("週番号", "修正", Format(wn, "00")): If t = "" Then Exit Sub
        wn = CInt(t)
        t = InputBox("日付(例:4/7)", "修正", md): If t = "" Then Exit Sub
        md = t
    End If
    md = 休業スキップ調整(CStr(ThisWorkbook.Sheets("基本設定").Range("B2").Value), md)
    If md = "" Then Exit Sub
    Dim tm As Worksheet: Set tm = ThisWorkbook.Sheets("週案_テンプレート")
    tm.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Dim nw As Worksheet: Set nw = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    nw.Unprotect ""
    nw.Name = "第" & Format(wn, "00") & "週"
    nw.Range("A1").Value = yr & "年度  第" & Format(wn, "00") & "週(" & md & "~)  進度予定表"
    Call 日付を入力(nw, md): Call 時間割展開(nw): Call 斜線を描画(nw)
    Dim cc(4) As Integer, cl As String, pi As Integer, di As Integer
    cc(0) = 4: cc(1) = 7: cc(2) = 10: cc(3) = 13: cc(4) = 16
    cl = ChrW(215) & ",月1,月2,月3,月4,月5,月6,月7,火1,火2,火3,火4,火5,火6,火7,水1,水2,水3,水4,水5,水6,水7,木1,木2,木3,木4,木5,木6,木7,金1,金2,金3,金4,金5,金6,金7"
    On Error Resume Next
    For pi = 1 To 7: For di = 0 To 4
        nw.Cells(5 + (pi - 1) * 2, cc(di)).Validation.Delete
        nw.Cells(5 + (pi - 1) * 2, cc(di)).Validation.Add Type:=xlValidateList, Formula1:=cl
        nw.Cells(5 + (pi - 1) * 2, cc(di)).Validation.IgnoreBlank = True
        nw.Cells(5 + (pi - 1) * 2, cc(di)).Validation.ShowError = False
    Next di: Next pi
    On Error GoTo 0
    Call ボタンを配置(nw): nw.Protect "": nw.Activate
    MsgBox "第" & Format(wn, "00") & "週 作成完了", vbInformation
End Sub
' ----
' 【7】ボタンを配置（全ボタン削除後にR列に3つ配置）
' ----
Sub ボタンを配置(tws As Worksheet)
    UP tws
    Dim shp As Shape, n() As String, c As Integer: c = 0
    For Each shp In tws.Shapes
        If shp.Type = msoFormControl Then
            ReDim Preserve n(c): n(c) = shp.Name: c = c + 1
        End If
    Next shp
    Dim i As Integer
    For i = 0 To c - 1: tws.Shapes(n(i)).Delete: Next i
    Dim bL As Double: bL = tws.Columns("R").Left + 5
    Dim defs(2, 1) As String
    defs(0, 0) = "新しい週を追加": defs(0, 1) = "ボタン_新しい週"
    defs(1, 0) = "時間割を再展開": defs(1, 1) = "ボタン_時間割展開"
    defs(2, 0) = "累計を更新":     defs(2, 1) = "ボタン_累計更新"
    Dim btnTop As Double: btnTop = tws.Rows(4).Top + 3
    For i = 0 To 2
        Dim btn As Shape
        Set btn = tws.Shapes.AddFormControl(xlButtonControl, bL, btnTop + i * 32, 90, 24)
        With btn
            .Name = "Btn_" & i
            .TextFrame.Characters.Text = defs(i, 0)
            .TextFrame.Characters.Font.Size = 9
            .TextFrame.Characters.Font.Name = "メイリオ"
            .OLEFormat.Object.OnAction = defs(i, 1)
        End With
    Next i
    LP tws
End Sub

' ----
' 【8】斜線を描画
' ----
Sub 斜線を描画(Optional targetWs As Worksheet = Nothing)
    Dim ws As Worksheet
    If targetWs Is Nothing Then
        Set ws = ActiveSheet
        If Left(ws.Name, 1) <> "第" And ws.Name <> "週案_テンプレート" Then
            MsgBox "週案シートを選択", vbExclamation: Exit Sub
        End If
    Else: Set ws = targetWs
    End If
    If targetWs Is Nothing Then UP ws
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 8) = "DiagLine" Then shp.Delete
    Next shp
    Dim cols As Variant: cols = Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
    Dim i As Integer
    For i = 0 To UBound(cols)
        Dim col As Integer: col = cols(i)
        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        x1 = ws.Cells(ROW_IMPL, col).Left + ws.Cells(ROW_IMPL, col).Width
        y1 = ws.Cells(ROW_IMPL, col).Top
        x2 = ws.Cells(ROW_PLAN, col).Left
        y2 = ws.Cells(ROW_PLAN, col).Top + ws.Cells(ROW_PLAN, col).Height
        Dim ln As Shape: Set ln = ws.Shapes.AddLine(x1, y1, x2, y2)
        With ln
            .Name = "DiagLine" & (i + 1)
            .Line.ForeColor.RGB = RGB(180, 180, 180): .Line.Weight = 0.5
        End With
    Next i
    If targetWs Is Nothing Then LP ws
End Sub
' ----
' 【9】時間割マスターから展開
' ----
Sub 時間割展開(Optional targetWs As Worksheet = Nothing)
    Dim ws As Worksheet, master As Worksheet
    If targetWs Is Nothing Then
        Set ws = ActiveSheet
        If Left(ws.Name, 1) <> "第" Then MsgBox "週案シートを選択", vbExclamation: Exit Sub
    Else: Set ws = targetWs
    End If
    Set master = ThisWorkbook.Sheets("時間割マスター")
    Application.ScreenUpdating = False
    If targetWs Is Nothing Then UP ws
    Dim sC(4) As Integer, cc(4) As Integer, mC(4) As Integer
    sC(0) = COL_MON_SCI: cc(0) = COL_MON_CHG: mC(0) = 2
    sC(1) = COL_TUE_SCI: cc(1) = COL_TUE_CHG: mC(1) = 3
    sC(2) = COL_WED_SCI: cc(2) = COL_WED_CHG: mC(2) = 4
    sC(3) = COL_THU_SCI: cc(3) = COL_THU_CHG: mC(3) = 5
    sC(4) = COL_FRI_SCI: cc(4) = COL_FRI_CHG: mC(4) = 6
    Dim d As Integer, period As Integer
    For d = 0 To 4
        For period = 1 To 7
            Dim sR As Integer: sR = PERIOD_START + (period - 1) * 2
            Dim cv As String: cv = Trim(ws.Cells(sR, cc(d)).Value)
            Dim clv As String
            If cv = "" Or cv = ChrW(215) Or cv = "x" Or cv = "X" Then
                clv = CStr(master.Cells(period + 2, mC(d)).Text)
            Else
                clv = GetClassFromChange(master, cv)
            End If
            With ws.Cells(sR, sC(d)): .NumberFormat = "@": .Value = clv: End With
        Next period
    Next d
    If targetWs Is Nothing Then LP ws
    Application.ScreenUpdating = True
    If targetWs Is Nothing Then MsgBox "展開完了", vbInformation
End Sub

Private Function GetClassFromChange(master As Worksheet, cv As String) As String
    If Len(cv) < 2 Then GetClassFromChange = "": Exit Function
    Dim dc As String: dc = Left(cv, 1)
    Dim pn As Integer
    On Error GoTo eH: pn = CInt(Mid(cv, 2))
    Dim mC As Integer
    Select Case dc
        Case "月": mC = 2: Case "火": mC = 3: Case "水": mC = 4
        Case "木": mC = 5: Case "金": mC = 6: Case Else: GetClassFromChange = "": Exit Function
    End Select
    If pn < 1 Or pn > 7 Then GetClassFromChange = "": Exit Function
    GetClassFromChange = CStr(master.Cells(pn + 2, mC).Text): Exit Function
eH: GetClassFromChange = ""
End Function
Private Function 次の月曜日() As String
    Dim ws As Worksheet, mx As Integer, lw As Worksheet: mx = 0
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 1) = "第" And Right(ws.Name, 1) = "週" Then
            Dim n As String: n = Mid(ws.Name, 2, Len(ws.Name) - 2)
            If IsNumeric(n) Then: If CInt(n) > mx Then mx = CInt(n): Set lw = ws
        End If
    Next ws
    If mx = 0 Then 次の月曜日 = "": Exit Function
    On Error GoTo eH
    Dim s As String: s = CStr(lw.Cells(ROW_DATE, COL_MON_SCI).Value)
    Dim p As Integer: p = InStr(s, "("): If p > 0 Then s = Left(s, p - 1)
    Dim pt() As String: pt = Split(s, "/")
    Dim bd As Date: bd = DateSerial(CInt(ThisWorkbook.Sheets("基本設定").Range("B2").Value), CInt(pt(0)), CInt(pt(1))) + 7
    次の月曜日 = Month(bd) & "/" & Day(bd): Exit Function
eH: 次の月曜日 = ""
End Function
' ----
' 【10】日付を自動入力
' ----
Private Sub 日付を入力(ws As Worksheet, md As String)
    On Error GoTo eH
    Dim pts() As String: pts = Split(md, "/")
    If UBound(pts) < 1 Then GoTo eH
    Dim m As Integer: m = CInt(pts(0))
    Dim d As Integer: d = CInt(pts(1))
    Dim bd As Date: bd = DateSerial(CInt(ThisWorkbook.Sheets("基本設定").Range("B2").Value), m, d)
    Dim dn(4) As String, sC(4) As Integer
    dn(0) = "月": sC(0) = COL_MON_SCI: dn(1) = "火": sC(1) = COL_TUE_SCI
    dn(2) = "水": sC(2) = COL_WED_SCI: dn(3) = "木": sC(3) = COL_THU_SCI
    dn(4) = "金": sC(4) = COL_FRI_SCI
    Dim i As Integer
    For i = 0 To 4
        With ws.Cells(ROW_DATE, sC(i))
            .NumberFormat = "@"
            .Value = Month(bd + i) & "/" & Day(bd + i) & "(" & dn(i) & ")"
        End With
    Next i
    Exit Sub
eH:
    With ws.Cells(ROW_DATE, COL_MON_SCI): .NumberFormat = "@": .Value = md & "(月)": End With
End Sub

Private Function 次の週番号() As Integer
    Dim ws As Worksheet, mx As Integer: mx = 0
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 1) = "第" And Right(ws.Name, 1) = "週" Then
            Dim ns As String: ns = Mid(ws.Name, 2, Len(ws.Name) - 2)
            If IsNumeric(ns) Then
                If CInt(ns) > mx Then mx = CInt(ns)
            End If
        End If
    Next ws
    次の週番号 = mx + 1
End Function
' ----
' 【★】テスト時 対象クラス入力ダイアログ
' ----
Sub HandleTestEntry(ws As Worksheet, target As Range)
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Sheets("基本設定")
    Dim defaultList As String: defaultList = ""
    Dim i As Integer
    For i = 6 To 13
        Dim c As String: c = Trim(CStr(cfg.Range("B" & i).Value))
        If c <> "" Then
            If defaultList = "" Then
                defaultList = c
            Else
                defaultList = defaultList & "," & c
            End If
        End If
    Next i
    Dim resp As Variant
    resp = Application.InputBox("テストでカウントするクラスをカンマ区切りで入力してください" & vbCrLf & "(不要なクラスを削除してください)", "テスト対象クラス", defaultList, Type:=2)
    Application.EnableEvents = False
    On Error Resume Next: ws.Unprotect "": On Error GoTo 0
    If VarType(resp) = vbBoolean Then
        target.Value = ""
        target.NumberFormat = "@"
    Else
        Dim s As String: s = Trim(CStr(resp))
        s = Replace(s, " ", "")
        s = Replace(s, "、", ",")
        If s = "" Then
            target.Value = ""
            target.NumberFormat = "@"
        Else
            target.NumberFormat = ";;;""テスト"""
            target.Value = "テスト[," & s & ",]"
        End If
    End If
    ws.Protect ""
    Application.EnableEvents = True
End Sub
' ----
' 【11】累計を全週シートから集計
' ----
Sub 累計を更新()
    Dim tw As Worksheet: Set tw = ActiveSheet
    If Left(tw.Name, 1) <> "第" Then MsgBox "週案シートを選択", vbExclamation: Exit Sub
    Dim cwn As Integer
    Dim cn As String: cn = Mid(tw.Name, 2, Len(tw.Name) - 2)
    If Not IsNumeric(cn) Then MsgBox "週シート名が不正", vbExclamation: Exit Sub
    cwn = CInt(cn)
    Application.ScreenUpdating = False
    UP tw
    Dim col As Integer
    For col = 2 To 15
        Dim subj As String: subj = tw.Cells(ROW_SUM_HDR, col).Value
        If subj = "" Then GoTo nxtC
        Dim total As Long: total = 0
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Sheets
            If Left(ws.Name, 1) = "第" And Right(ws.Name, 1) = "週" Then
                Dim ns As String: ns = Mid(ws.Name, 2, Len(ws.Name) - 2)
                If IsNumeric(ns) Then
                    If CInt(ns) <= cwn Then
                        Dim v As Variant: v = ws.Cells(ROW_IMPL, col).Value
                        If IsNumeric(v) And v <> "" Then total = total + CLng(v)
                    End If
                End If
            End If
        Next ws
        tw.Cells(ROW_CUM, col).Value = IIf(total = 0, "", total)
nxtC:
    Next col
    斜線を描画 tw
    LP tw
    Application.ScreenUpdating = True
    MsgBox "更新完了", vbInformation
End Sub
' ----
' 【12】学期末PDF出力
' ----
Sub 学期末PDF_1学期(): Call 学期PDF出力(4, 7, "1学期"): End Sub
Sub 学期末PDF_2学期(): Call 学期PDF出力(9, 12, "2学期"): End Sub
Sub 学期末PDF_3学期(): Call 学期PDF出力(1, 3, "3学期"): End Sub

Private Sub 学期PDF出力(sM As Integer, eM As Integer, tn As String)
    Dim ws As Worksheet, tg() As String, c As Integer: c = 0
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 1) = "第" Then
            Dim dS As String: dS = CStr(ws.Cells(ROW_DATE, COL_MON_SCI).Value)
            Dim m As Integer: m = ExtractMonth(dS)
            Dim ok As Boolean
            If sM <= eM Then ok = (m >= sM And m <= eM) Else ok = (m >= sM Or m <= eM)
            If ok Or m = 0 Then
                ReDim Preserve tg(c): tg(c) = ws.Name: c = c + 1
            End If
        End If
    Next ws
    If c = 0 Then MsgBox tn & "該当なし", vbExclamation: Exit Sub
    Dim pp As String
    pp = ThisWorkbook.Path & "\" & CStr(ThisWorkbook.Sheets("基本設定").Range("B2").Value) & "年度_" & tn & "_進度予定表.pdf"
    ThisWorkbook.Sheets(tg).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pp, Quality:=xlQualityStandard, IgnorePrintAreas:=False
    ThisWorkbook.Sheets("基本設定").Select
    MsgBox tn & "PDF出力完了", vbInformation
End Sub

Private Function ExtractMonth(s As String) As Integer
    On Error GoTo eH
    s = Trim(s)
    If s = "" Then ExtractMonth = 0: Exit Function
    If InStr(s, "/") > 0 Then ExtractMonth = CInt(Split(s, "/")(0)): Exit Function
    If InStr(s, "月") > 0 Then ExtractMonth = CInt(Split(s, "月")(0)): Exit Function
    ExtractMonth = 0: Exit Function
eH: ExtractMonth = 0
End Function
' ----
' 【13】次年度ファイルを作成
' ----
Sub 次年度ファイルを作成()
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Sheets("基本設定")
    Dim cy As Integer: cy = CInt(cfg.Range("B2").Value)
    If MsgBox(cy & "年度を保存し" & (cy + 1) & "年度を作成。よろしいですか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    ThisWorkbook.Save
    Dim np As String
    np = Application.GetSaveAsFilename(InitialFileName:=Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_" & (cy + 1) & "年度.xlsm", FileFilter:="Excelマクロ有効ブック (*.xlsm), *.xlsm")
    If np = "False" Then Exit Sub
    ThisWorkbook.SaveCopyAs np
    Dim nw As Workbook: Set nw = Workbooks.Open(np)
    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    nw.Sheets("基本設定").Range("B2").Value = cy + 1
    Dim ws As Worksheet, dn() As String, c As Integer: c = 0
    For Each ws In nw.Sheets
        If Left(ws.Name, 1) = "第" Then
            ReDim Preserve dn(c): dn(c) = ws.Name: c = c + 1
        End If
    Next ws
    Dim i As Integer
    For i = 0 To c - 1: nw.Sheets(dn(i)).Delete: Next i
    On Error Resume Next: nw.Sheets("VBAコード").Delete: On Error GoTo 0
    Dim nc As Worksheet: Set nc = nw.Sheets("基本設定")
    On Error Resume Next: nc.Unprotect ""
    Dim Sh As Shape, dl() As String, dc2 As Integer: dc2 = 0
    For Each Sh In nc.Shapes
        If Sh.Type = msoFormControl Then ReDim Preserve dl(dc2): dl(dc2) = Sh.Name: dc2 = dc2 + 1
    Next Sh
    For i = 0 To dc2 - 1: nc.Shapes(dl(i)).Delete: Next i
    nc.Activate: Dim sb As Shape
    Set sb = nc.Shapes.AddFormControl(xlButtonControl, nc.Range("E34").Left, nc.Range("E34").Top, 120, 22)
    With sb: .Name = "BtnSetup": .TextFrame.Characters.Text = "設定を完了"
    .TextFrame.Characters.Font.Size = 9: .TextFrame.Characters.Font.Name = "メイリオ"
    .OLEFormat.Object.OnAction = "ボタン_設定完了"
    End With: On Error GoTo 0
    On Error Resume Next
    nw.Sheets("時間割マスター").Range("B3:F9").ClearContents
    Dim yr As String: yr = CStr(cy + 1)
    nw.Sheets("週案_テンプレート").Range("A1").Value = yr & "年度" & ChrW(12288) & "第" & ChrW(12288) & ChrW(12288) & "週（" & ChrW(12288) & ChrW(12288) & "月" & ChrW(12288) & ChrW(12288) & "日?" & ChrW(12288) & ChrW(12288) & "月" & ChrW(12288) & ChrW(12288) & "日）" & ChrW(12288) & "進度予定表"
    On Error GoTo 0
    Application.DisplayAlerts = True
    nw.Save: Application.ScreenUpdating = True
    MsgBox (cy + 1) & "年度ファイルを作成しました。", vbInformation
End Sub
' ----
Private Sub UP(w As Worksheet): On Error Resume Next: w.Unprotect "": End Sub
Private Sub LP(w As Worksheet): w.Protect "": End Sub
' ----
' 【14】ボタン用マクロ
' ----
Sub ボタン_設定完了():     Call 設定を完了:        End Sub
Sub ボタン_新しい週():     Call 新しい週を追加:    End Sub
Sub ボタン_時間割展開():   Call 時間割展開:        End Sub
Sub ボタン_累計更新():     Call 累計を更新:        End Sub
Sub ボタン_1学期PDF():     Call 学期末PDF_1学期:   End Sub
Sub ボタン_2学期PDF():     Call 学期末PDF_2学期:   End Sub
Sub ボタン_3学期PDF():     Call 学期末PDF_3学期:   End Sub
Sub ボタン_次年度():       Call 次年度ファイルを作成:    End Sub
Sub ボタン_旧ファイル移行(): Call 旧ファイルから移行:        End Sub
Sub ボタン_全データ消去(): Call 全データ消去:        End Sub

' ================================================================
' 決裁欄を設定（L?P列）
' ================================================================
Sub 決裁欄設定(ws As Worksheet)
    ' H1: 決裁ラベル
    With ws.Cells(1, 8)
        .Value = "決裁": .Font.Name = "メイリオ": .Font.Size = 7: .Font.Bold = True
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(150, 150, 150)
    End With
    ' I-J:校長, K-L:教頭, M-N:主幹, O-P:教務（各2列結合）
    Dim kn As Variant: kn = Array("校長", "教頭", "主幹", "教務")
    Dim kc As Integer
    For kc = 0 To 3
        Dim c1 As Integer: c1 = 9 + kc * 2
        Dim c2 As Integer: c2 = c1 + 1
        ws.Range(ws.Cells(1, c1), ws.Cells(1, c2)).Merge
        With ws.Cells(1, c1)
            .Value = kn(kc): .Font.Name = "メイリオ": .Font.Size = 7
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlTop
        End With
        ws.Range(ws.Cells(1, c1), ws.Cells(1, c2)).Borders.LineStyle = xlContinuous
        ws.Range(ws.Cells(1, c1), ws.Cells(1, c2)).Borders.Weight = xlThin
        ws.Range(ws.Cells(1, c1), ws.Cells(1, c2)).Borders.Color = RGB(150, 150, 150)
    Next kc
End Sub

' ================================================================
' 集計補正：特活+学活合算、活動を右詰め配置、背景修正
' ================================================================
Sub 集計補正(ws As Worksheet)
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Sheets("基本設定")
    Dim col As Integer, ai As Integer

    ' --- 学活カウント方式（基本設定B4）：「含む」=特活に合算 / それ以外=学活で独立 ---
    Dim mergeGaku As Boolean
    mergeGaku = (InStr(CStr(cfg.Range("B4").Value), "含") > 0)
    ' --- 活動名を収集（mergeGaku のときのみ学活を特活に統合）---
    Dim actNames() As String, actCount As Integer: actCount = 0
    Dim srcRows As Variant: srcRows = Array(16, 17, 18, 19, 20)
    Dim hasToku As Boolean, hasGaku As Boolean
    hasToku = False: hasGaku = False
    Dim ri As Integer
    For ri = 0 To UBound(srcRows)
        Dim nm As String: nm = ""
        On Error Resume Next: nm = CStr(cfg.Range("B" & srcRows(ri)).Value): On Error GoTo 0
        If nm = "特活" Then hasToku = True
        If nm = "学活" Then hasGaku = True
        Dim skipIt As Boolean: skipIt = (mergeGaku And nm = "学活")
        If nm <> "" And Not skipIt Then
            ReDim Preserve actNames(actCount)
            If mergeGaku And nm = "特活" And hasGaku Then
                actNames(actCount) = "特活(学活含む)"
            Else
                actNames(actCount) = nm
            End If
            actCount = actCount + 1
        End If
    Next ri
    ' 2回目パス：特活が後に来る場合の対応（合算モードのみ）
    If mergeGaku And hasToku And hasGaku Then
        For ai = 0 To actCount - 1
            If actNames(ai) = "特活" Then actNames(ai) = "特活(学活含む)"
        Next ai
    End If

    ' --- 最後のクラス列を検出 ---
    Dim lastClass As Integer: lastClass = 1
    For col = 2 To 9
        If CStr(ws.Cells(ROW_SUM_HDR, col).Value) <> "" Then lastClass = col
    Next col

    ' --- 活動の配置先（右詰め：col14から左へ）---
    Dim firstAct As Integer: firstAct = 16 - actCount

    ' --- 行ごとの標準背景色 ---
    Dim rowBG(25) As Long
    rowBG(ROW_SUM_HDR) = RGB(47, 85, 151)
    rowBG(ROW_KOMA) = RGB(227, 242, 253)
    rowBG(ROW_IMPL) = RGB(255, 255, 255)
    rowBG(ROW_PLAN) = RGB(255, 249, 196)
    rowBG(ROW_CUM) = RGB(252, 228, 236)

    ' --- クラス列より右?col15をクリア・背景リセット ---
    Dim r As Integer
    For col = lastClass + 1 To 15
        For r = ROW_SUM_HDR To ROW_CUM
            ws.Cells(r, col).Value = ""
            ws.Cells(r, col).ClearContents
            ws.Cells(r, col).Interior.Color = rowBG(r)
            ws.Cells(r, col).Font.Name = "メイリオ"
            ws.Cells(r, col).Font.Size = 9
            ws.Cells(r, col).Font.Bold = False
            ws.Cells(r, col).HorizontalAlignment = xlCenter
            ws.Cells(r, col).VerticalAlignment = xlCenter
            ws.Cells(r, col).ShrinkToFit = False
        Next r
        ' ヘッダー行の文字色
        ws.Cells(ROW_SUM_HDR, col).Font.Color = RGB(255, 255, 255)
        ws.Cells(ROW_SUM_HDR, col).Font.Bold = True
        ws.Cells(ROW_SUM_HDR, col).Font.Size = 8
        ' 各行の文字色
        ws.Cells(ROW_KOMA, col).Font.Color = RGB(26, 35, 126)
        ws.Cells(ROW_IMPL, col).Font.Color = RGB(27, 94, 32)
        ws.Cells(ROW_IMPL, col).Font.Bold = True
        ws.Cells(ROW_IMPL, col).HorizontalAlignment = xlLeft
        ws.Cells(ROW_PLAN, col).Font.Color = RGB(121, 85, 0)
        ws.Cells(ROW_PLAN, col).HorizontalAlignment = xlRight
        ws.Cells(ROW_CUM, col).Font.Color = RGB(136, 14, 79)
        ws.Cells(ROW_CUM, col).Font.Bold = True
    Next col

    ' --- 活動を右詰めで配置 ---
    Dim sciR As Variant: sciR = Array(5, 7, 9, 11, 13, 15, 17)
    Dim dS(4) As Integer, dE(4) As Integer, dc(4) As Integer
    dS(0) = 2: dE(0) = 3: dS(1) = 5: dE(1) = 6: dS(2) = 8: dE(2) = 9
    dS(3) = 11: dE(3) = 12: dS(4) = 14: dE(4) = 15
    dc(0) = 4: dc(1) = 7: dc(2) = 10: dc(3) = 13: dc(4) = 16

    For ai = 0 To actCount - 1
        Dim tc As Integer: tc = firstAct + ai
        Dim actNm As String: actNm = actNames(ai)
        Dim isCombined As Boolean: isCombined = (actNm = "特活(学活含む)")

        ' ヘッダー
        ws.Cells(ROW_SUM_HDR, tc).Value = actNm
        If isCombined Then
            ws.Cells(ROW_SUM_HDR, tc).Font.Size = 7
            ws.Cells(ROW_SUM_HDR, tc).ShrinkToFit = True
        End If

        ' 予定数式
        Dim pp As String: pp = ""
        Dim mp As String: mp = ""
        Dim si As Integer, dd As Integer
        For si = 0 To UBound(sciR)
            For dd = 0 To 4
                Dim sA As String: sA = ws.Cells(sciR(si), dS(dd)).Address(False, False)
                Dim eA As String: eA = ws.Cells(sciR(si), dE(dd)).Address(False, False)
                Dim cA As String: cA = ws.Cells(sciR(si), dc(dd)).Address(False, False)
                If pp <> "" Then pp = pp & "+"
                If isCombined Then
                    pp = pp & "COUNTIF(" & sA & ":" & eA & ",""特活"")+COUNTIF(" & sA & ":" & eA & ",""学活"")"
                    If mp <> "" Then mp = mp & "+"
                    mp = mp & "SUMPRODUCT((" & sA & ":" & eA & "=""特活"")*(" & cA & "=""×""))+SUMPRODUCT((" & sA & ":" & eA & "=""学活"")*(" & cA & "=""×""))"
                Else
                    Dim hRef As String: hRef = ws.Cells(ROW_SUM_HDR, tc).Address(False, False)
                    pp = pp & "COUNTIF(" & sA & ":" & eA & "," & hRef & ")"
                    If mp <> "" Then mp = mp & "+"
                    mp = mp & "SUMPRODUCT((" & sA & ":" & eA & "=" & hRef & ")*(" & cA & "=""×""))"
                End If
            Next dd
        Next si

        ws.Cells(ROW_PLAN, tc).Formula = "=IF(" & pp & "=0,""""," & pp & ")"
        Dim pRef As String: pRef = ws.Cells(ROW_PLAN, tc).Address(False, False)
        ws.Cells(ROW_IMPL, tc).Formula = "=IF(" & pRef & "="""",""""," & pRef & "-(" & mp & "))"

        ' 週コマ数
        If isCombined Then
            Dim tK As Variant, gK As Variant
            tK = 0: gK = 0
            On Error Resume Next
            tK = Application.VLookup("特活", cfg.Range("A23:B35"), 2, False)
            gK = Application.VLookup("学活", cfg.Range("A23:B35"), 2, False)
            On Error GoTo 0
            If IsError(tK) Then tK = 0
            If IsError(gK) Then gK = 0
            Dim kTotal As Long: kTotal = CLng(Val(CStr(tK))) + CLng(Val(CStr(gK)))
            ws.Cells(ROW_KOMA, tc).Value = IIf(kTotal = 0, "", kTotal)
        Else
            Dim hR2 As String: hR2 = ws.Cells(ROW_SUM_HDR, tc).Address(False, False)
            ws.Cells(ROW_KOMA, tc).Formula = "=IF(" & hR2 & "="""","""",IFERROR(VALUE(VLOOKUP(" & hR2 & ",基本設定!$A$23:$B$35,2,FALSE())),""""))"
        End If
    Next ai

    ' --- P列(16)の合計数式を更新 ---
    Dim sumR As Variant
    For Each sumR In Array(ROW_KOMA, ROW_IMPL, ROW_PLAN, ROW_CUM)
        ws.Cells(CLng(sumR), 16).Formula = "=IF(SUM(B" & CLng(sumR) & ":O" & CLng(sumR) & ")=0,"""",SUM(B" & CLng(sumR) & ":O" & CLng(sumR) & "))"
    Next sumR
End Sub

' ----
' 【インポート】旧ファイルから週シートデータを移行
' ----
Sub 旧ファイルから移行()
    Dim fp As Variant
    fp = Application.GetOpenFilename("Excelマクロ有効ブック (*.xlsm),*.xlsm", , "旧ファイルを選択")
    If VarType(fp) = vbBoolean Then Exit Sub
    If MsgBox("選択した旧ファイルから週シートのデータをインポートします。" & vbLf & "同名の週シートは上書きされます。よろしいですか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Dim ob As Workbook
    On Error GoTo openErr
    Set ob = Workbooks.Open(CStr(fp), ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0
    Dim cnt As Integer: cnt = 0
    Dim oldWs As Worksheet, newWs As Worksheet
    Dim tm As Worksheet: Set tm = ThisWorkbook.Sheets("週案_テンプレート")
    On Error Resume Next
    ThisWorkbook.Sheets("時間割マスター").Unprotect ""
    ThisWorkbook.Sheets("時間割マスター").Range("B3:F9").Value = ob.Sheets("時間割マスター").Range("B3:F9").Value
    ThisWorkbook.Sheets("時間割マスター").Protect ""
    On Error GoTo 0
    For Each oldWs In ob.Sheets
        If Left(oldWs.Name, 1) = "第" And Right(oldWs.Name, 1) = "週" Then
            Set newWs = Nothing
            On Error Resume Next: Set newWs = ThisWorkbook.Sheets(oldWs.Name): On Error GoTo 0
            If newWs Is Nothing Then
                tm.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set newWs = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                newWs.Name = oldWs.Name
                newWs.Unprotect ""
                Call 斜線を描画(newWs)
                Call ボタンを配置(newWs)
            End If
            newWs.Unprotect ""
            newWs.Range("A1").Value = oldWs.Range("A1").Value
            Dim rr As Integer, cc As Integer, v As Variant, nf As String
            For rr = 3 To 19
                For cc = 2 To 16
                    v = oldWs.Cells(rr, cc).Value
                    nf = oldWs.Cells(rr, cc).NumberFormat
                    If Not IsEmpty(v) Then
                        newWs.Cells(rr, cc).NumberFormat = nf
                        newWs.Cells(rr, cc).Value = v
                    End If
                Next cc
            Next rr
            For cc = 2 To 14
                v = oldWs.Cells(25, cc).Value
                If Not IsEmpty(v) Then newWs.Cells(25, cc).Value = v
            Next cc
            newWs.Protect ""
            cnt = cnt + 1
        End If
    Next oldWs
    ob.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox cnt & " 週シートをインポートしました。", vbInformation
    Exit Sub
openErr:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "ファイルを開けませんでした。", vbCritical
End Sub

' ----
' 【ファイル開いた時】基本設定のボタンを最新版に差し替え
' ----
Sub Auto_Open()
    On Error Resume Next
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Sheets("基本設定")
    If cfg Is Nothing Then Exit Sub
    基本設定にボタンを追加
    cfg.Unprotect ""
    学活方式を準備 cfg
    休業設定を準備 cfg
    cfg.Protect ""
End Sub

Private Sub 休業設定を準備(cfg As Worksheet)
    On Error Resume Next
    ' セクション見出し（A37 を既存見出し A5 と同じ書式に）
    cfg.Range("A37").Value = "【長期休業期間】"
    cfg.Range("A5").Copy: cfg.Range("A37").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    Dim labels As Variant
    labels = Array("夏休み開始", "夏休み終了", "冬休み開始", "冬休み終了", "春休み開始", "春休み終了")
    Dim exs As Variant
    exs = Array("7/18", "8/31", "12/25", "1/6", "3/25", "4/6")
    Dim i As Integer
    For i = 0 To 5
        Dim r As Long: r = 38 + i
        ' ラベルセル（A38.. を A6 と同じ書式に、例を併記）
        cfg.Range("A" & r).Value = labels(i)
        cfg.Range("A6").Copy: cfg.Range("A" & r).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        ' 入力セル（B38.. を B2 と同じ書式、テキスト（M/D）入力、ツールチップ付き）
        cfg.Range("B2").Copy: cfg.Range("B" & r).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        cfg.Range("B" & r).NumberFormat = "@"
        cfg.Range("B" & r).Locked = False
        With cfg.Range("B" & r).Validation
            .Delete
            .Add Type:=xlValidateInputOnly
            .InputTitle = "入力例"
            .InputMessage = "「月/日」の形式で入力してください。" & vbLf & "例: " & exs(i)
            .ShowInput = True
        End With
    Next i
End Sub
' ----
' 【長期休業】mdが休業期間内なら明けの月曜に進める
' 戻り値：調整後の md 文字列（"M/D"形式）。ユーザーがキャンセルしたら ""
' ----
Function 休業スキップ調整(yr As String, md As String) As String
    On Error GoTo eH
    休業スキップ調整 = md
    If md = "" Or yr = "" Then Exit Function
    Dim pts() As String: pts = Split(md, "/")
    If UBound(pts) < 1 Then Exit Function
    Dim yrN As Integer: yrN = CInt(yr)
    Dim bd As Date: bd = DateSerial(yrN, CInt(pts(0)), CInt(pts(1)))
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Sheets("基本設定")
    Dim iter As Integer
    For iter = 1 To 10
        Dim inBreak As Boolean: inBreak = False
        Dim r As Integer
        For r = 38 To 42 Step 2
            Dim sTxt As String: sTxt = Trim(CStr(cfg.Range("B" & r).Value))
            Dim eTxt As String: eTxt = Trim(CStr(cfg.Range("B" & (r + 1)).Value))
            If sTxt <> "" And eTxt <> "" Then
                Dim sp() As String: sp = Split(sTxt, "/")
                Dim ep() As String: ep = Split(eTxt, "/")
                If UBound(sp) >= 1 And UBound(ep) >= 1 Then
                    Dim sM As Integer, sD As Integer, eM As Integer, eD As Integer
                    sM = CInt(sp(0)): sD = CInt(sp(1))
                    eM = CInt(ep(0)): eD = CInt(ep(1))
                    Dim bO As Long: bO = Month(bd) * 100 + Day(bd)
                    Dim sO As Long: sO = sM * 100 + sD
                    Dim eO As Long: eO = eM * 100 + eD
                    Dim hit As Boolean
                    If sO <= eO Then
                        hit = (bO >= sO And bO <= eO)
                    Else
                        hit = (bO >= sO Or bO <= eO)
                    End If
                    If hit Then
                        ' 休業明けの日（end+1）を含む週の月曜へ
                        Dim nxt As Date: nxt = DateSerial(Year(bd), eM, eD) + 1
                        If nxt <= bd Then nxt = DateSerial(Year(bd) + 1, eM, eD) + 1
                        Do While Weekday(nxt, vbMonday) <> 1
                            nxt = nxt - 1
                        Loop
                        bd = nxt
                        inBreak = True
                        Exit For
                    End If
                End If
            End If
        Next r
        If Not inBreak Then Exit For
    Next iter
    Dim newMd As String: newMd = Month(bd) & "/" & Day(bd)
    If newMd <> md Then
        Dim ans As Integer
        ans = MsgBox("入力された " & md & " は長期休業期間内のため、" & newMd & " に自動調整します。よろしいですか？", vbYesNoCancel + vbQuestion)
        If ans = vbCancel Then 休業スキップ調整 = "": Exit Function
        If ans = vbNo Then 休業スキップ調整 = md: Exit Function
        休業スキップ調整 = newMd
    End If
    Exit Function
eH:
    休業スキップ調整 = md
End Function

Public Sub 学活方式を準備(cfg As Worksheet)
    On Error Resume Next
    cfg.Range("A4").Value = "集計方式：学活"
    If Trim(CStr(cfg.Range("B4").Value)) = "" Then cfg.Range("B4").Value = "学活で独立"
    cfg.Range("B4").Locked = False
    With cfg.Range("B4").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="特活に含める,学活で独立"
        .IgnoreBlank = True: .ShowError = False
    End With
    ' 集計方式に応じて特別活動①/②と説明欄を同期
    Dim isMerge As Boolean
    isMerge = (InStr(CStr(cfg.Range("B4").Value), "含") > 0)
    If isMerge Then
        cfg.Range("B16").Value = "特活"
        cfg.Range("B17").Value = "学活"
        cfg.Range("C16").Value = "集計では学活と合算されます"
        cfg.Range("C17").ClearContents
    Else
        ' 独立モード：特別活動①/②を自動表示しない（自動設定値のみ消去）
        If CStr(cfg.Range("B16").Value) = "特活" Then cfg.Range("B16").ClearContents
        If CStr(cfg.Range("B17").Value) = "学活" Then cfg.Range("B17").ClearContents
        cfg.Range("C16").ClearContents
        cfg.Range("C17").ClearContents
    End If
End Sub

Public Sub 学活方式チェンジ処理(Sh As Object, target As Range)
    If Sh.Name <> "基本設定" Then Exit Sub
    If target.Cells.Count <> 1 Then Exit Sub
    If target.Address(False, False) <> "B4" Then Exit Sub
    Application.EnableEvents = False
    On Error Resume Next
    Sh.Unprotect ""
    学活方式を準備 Sh
    Sh.Protect ""
    Application.EnableEvents = True
End Sub

Sub 全データ消去()
    If MsgBox("全データを消去します。元に戻せません。", vbOKCancel + vbExclamation + vbDefaultButton2, "確認") <> vbOK Then Exit Sub
    Application.ScreenUpdating = False: Application.DisplayAlerts = False: Application.EnableEvents = False
    On Error Resume Next
    Dim ws As Worksheet, n() As String, c As Long: c = 0
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 1) = "第" And Right(ws.Name, 1) = "週" Then ReDim Preserve n(c): n(c) = ws.Name: c = c + 1
    Next ws
    Dim i As Long
    For i = 0 To c - 1: ThisWorkbook.Sheets(n(i)).Delete: Next i
    Dim tm As Worksheet: Set tm = ThisWorkbook.Sheets("時間割マスター")
    tm.Unprotect "": tm.Range("B3:F9").ClearContents: tm.Protect ""
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Sheets("基本設定")
    cfg.Unprotect ""
    Dim rs As Variant: rs = Array("B2:B4", "B6:B14", "B16:B20", "C16:C17", "B23:B35", "B38:B43")
    For i = 0 To UBound(rs): cfg.Range(CStr(rs(i))).ClearContents: Next i
    学活方式を準備 cfg: 休業設定を準備 cfg: cfg.Protect ""
    Application.EnableEvents = True: Application.DisplayAlerts = True: Application.ScreenUpdating = True
End Sub
