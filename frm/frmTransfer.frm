VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTransfer 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmTransfer.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_handlers As Collection
Public m_cn       As Object

Private Sub UserForm_Initialize()
    Set m_handlers = New Collection
    Set m_cn = GetConnection()

    Me.Caption = "転記"
    Me.Width = 400
    Me.BackColor = RGB(240, 244, 251)

    Const LX As Integer = 16
    Const TX As Integer = 150
    Const LW As Integer = 130
    Const TW As Integer = 210
    Const TH As Integer = 24
    Const SP As Integer = 32
    Dim topY As Integer: topY = 14
    Dim addL As Integer: addL = 0
    '── タイトル ──
    With Lbl("lblTitle", "転記する書類を選択してください", LX, topY, 360, 22)
        .Font.Size = 11: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 34

    '── チェックボックスと注記ラベル ──
    Dim chkMit As MSForms.CheckBox
    Set chkMit = Me.Controls.Add("Forms.CheckBox.1", "chkMit", True)
    With chkMit
        .Caption = SHEET_MITSUMORI
        .Left = LX: .top = topY
        .Width = 64: .Height = 22
        .Font.Size = 10: .Value = True
    End With
    addL = chkMit.Left + chkMit.Width - 4
    Lbl "lblNoteMit", "", addL, topY + 5, 50, 16
    Me.Controls("lblNoteMit").TextAlign = fmTextAlignLeft

    Dim chkNou As MSForms.CheckBox
    Set chkNou = Me.Controls.Add("Forms.CheckBox.1", "chkNou", True)
    With chkNou
        .Caption = SHEET_NOUHINN
        .Left = LX + 130: .top = topY
        .Width = 64: .Height = 22
        .Font.Size = 10: .Value = True
    End With
    addL = chkNou.Left + chkNou.Width - 4
    Lbl "lblNoteNou", "", addL, topY + 5, 50, 16
    Me.Controls("lblNoteNou").TextAlign = fmTextAlignLeft
    
    Dim chkSei As MSForms.CheckBox
    Set chkSei = Me.Controls.Add("Forms.CheckBox.1", "chkSei", True)
    With chkSei
        .Caption = SHEET_SEIKYUU
        .Left = LX + 260: .top = topY
        .Width = 64: .Height = 22
        .Font.Size = 10: .Value = True
    End With
    addL = chkSei.Left + chkSei.Width - 4
    Lbl "lblNoteSei", "", addL, topY + 5, 50, 16
    Me.Controls("lblNoteSei").TextAlign = fmTextAlignLeft

    Dim chkRyo As MSForms.CheckBox
    Set chkRyo = Me.Controls.Add("Forms.CheckBox.1", "chkRyo", True)
    With chkRyo
        .Caption = SHEET_RYOUSYUU
        .Left = LX + 130: .top = topY + 28
        .Width = 64: .Height = 22
        .Font.Size = 10: .Value = True
    End With
    addL = chkRyo.Left + chkRyo.Width - 4
    Lbl "lblNoteRyo", "", addL, topY + 28 + 5, 50, 16
    Me.Controls("lblNoteRyo").TextAlign = fmTextAlignLeft

    topY = topY + 60

    '── 区切り ──
    With Lbl("lblSep1", "", LX, topY, 355, 1)
        .BackColor = RGB(180, 180, 180)
        .BackStyle = fmBackStyleOpaque
    End With
    topY = topY + 12

    '── 注文情報セクションタイトル ──
    With Lbl("lblSec", "注文情報", LX, topY, 200, 20)
        .Font.Size = 10: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 28

    '── 件名 ──
    Lbl "lblSubject", "件名", LX, topY + 4, LW, 18
    txt "txtSubject", "", "hiragana", TX, topY, TW, TH
    topY = topY + SP

    '── 納期 ──
    Lbl "lblDelivery", "納期", LX, topY + 4, LW, 18
    txt "txtDelivery", "", "hankaku", TX, topY, 120, TH
    topY = topY + SP

    '── 見積有効期限 ──
    Lbl "lblExpiry", "見積有効期限", LX, topY + 4, LW, 18
    txt "txtExpiry", "", "hankaku", TX, topY, 120, TH
    topY = topY + SP

    '── 支払条件 ──
    Lbl "lblPayTerms", "支払条件", LX, topY + 4, LW, 18
    txt "txtPayTerms", "", "hiragana", TX, topY, TW, TH
    topY = topY + SP

    '── 支払期限 ──
    Lbl "lblPayDue", "支払期限", LX, topY + 4, LW, 18
    txt "txtPayDue", "", "hankaku", TX, topY, 120, TH
    topY = topY + SP + 8

    '── 区切り ──
    With Lbl("lblSep2", "", LX, topY, 355, 1)
        .BackColor = RGB(180, 180, 180)
        .BackStyle = fmBackStyleOpaque
    End With
    topY = topY + 12

    '── ボタン ──
    Const BW As Integer = 96
    Const BH As Integer = 30
    Const BTop As Integer = 10

    Dim bCancel As MSForms.CommandButton
    Set bCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancelTrans", True)
    With bCancel
        .Caption = "キャンセル"
        .top = topY + BTop
        .Left = Me.InsideWidth / 2 - BW - 16
        .Width = BW: .Height = BH
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
        .Font.Size = 11
    End With

    Dim bExec As MSForms.CommandButton
    Set bExec = Me.Controls.Add("Forms.CommandButton.1", "btnExecTrans", True)
    With bExec
        .Caption = "転記実行"
        .top = topY + BTop
        .Left = Me.InsideWidth / 2 + 16
        .Width = BW: .Height = BH
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 11
        .Font.Bold = True
    End With

    '── イベントハンドラ登録 ──
    Dim h As clsBtnHandler
    Set h = New clsBtnHandler
    Set h.btn = bCancel: h.act = "CANCEL_TRANS": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = bExec: h.act = "EXEC_TRANS": Set h.frm = Me
    m_handlers.Add h

    Me.Height = topY + BH + BTop + 52

    '── 既存データの読み込み ──
    Call LoadOrderData
End Sub

'── 注文番号で既存データを検索して初期値をセット ──
Public Sub LoadOrderData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("入力")
    Dim orderNum As String
    orderNum = ws.Range(SRC_RANGE_OUTORDRNUM).Value
    If orderNum = "" Then Exit Sub

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM orders WHERE order_number = '" & orderNum & "' LIMIT 1", m_cn

    If Not rs.EOF Then
        Me.Controls("txtSubject").Text = rs("subject") & ""
        Me.Controls("txtDelivery").Text = IIf(IsNull(rs("delivery_date")), "", Format(rs("delivery_date"), "yyyy/mm/dd"))
        Me.Controls("txtExpiry").Text = IIf(IsNull(rs("expiry_date")), "", Format(rs("expiry_date"), "yyyy/mm/dd"))
        Me.Controls("txtPayTerms").Text = rs("payment_terms") & ""
        Me.Controls("txtPayDue").Text = IIf(IsNull(rs("payment_due_date")), "", Format(rs("payment_due_date"), "yyyy/mm/dd"))

        '── 発行済みフラグに合わせて注記ラベルを表示 ──
        Dim flagMap(3) As Boolean
        flagMap(0) = CBool(rs("is_quoted"))
        flagMap(1) = CBool(rs("is_delivered"))
        flagMap(2) = CBool(rs("is_invoiced"))
        flagMap(3) = CBool(rs("is_receipted"))

        Dim noteNames(3) As String
        noteNames(0) = "lblNoteMit"
        noteNames(1) = "lblNoteNou"
        noteNames(2) = "lblNoteSei"
        noteNames(3) = "lblNoteRyo"

        Dim i As Integer
        For i = 0 To 3
            If flagMap(i) Then
                Me.Controls(noteNames(i)).Caption = "発行済み"
                Me.Controls(noteNames(i)).ForeColor = RGB(180, 0, 0)
            Else
                Me.Controls(noteNames(i)).Caption = ""
            End If
        Next i
    End If

    rs.Close
    Set rs = Nothing
End Sub
Private Sub UserForm_Terminate()
    If Not m_cn Is Nothing Then
        If m_cn.State = 1 Then m_cn.Close
    End If
    Set m_cn = Nothing
End Sub

Private Function Lbl(nm$, cap$, l%, t%, w%, h%) As MSForms.Label
    Set Lbl = Me.Controls.Add("Forms.Label.1", nm, True)
    With Lbl
        .Caption = cap: .Left = l: .top = t
        .Width = w: .Height = h: .Font.Size = 10
    End With
End Function

Private Function txt(nm$, val$, ime$, l%, t%, w%, h%) As MSForms.TextBox
    Set txt = Me.Controls.Add("Forms.TextBox.1", nm, True)
    With txt
        .Text = val: .Left = l: .top = t
        .Width = w: .Height = h: .Font.Size = 10
        If ime = "hiragana" Then
            .IMEMode = xlIMEModeHiragana
        Else
            .IMEMode = xlIMEModeDisable
        End If
    End With
End Function
