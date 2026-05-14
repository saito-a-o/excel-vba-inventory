VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClientEdit 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmClientEdit.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmClientEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_mode As String
Public m_cn As Object
Public m_clientId As Long
Public m_parentFrm As Object

Private m_handlers As Collection

Private Sub UserForm_Initialize()
    Set m_handlers = New Collection

    '── フォーム外観 ──
    Me.Width = 460
    Me.BackColor = RGB(240, 244, 251)

    '── レイアウト定数 ──
    Const LX As Integer = 14
    Const TX As Integer = 160
    Const LW As Integer = 140
    Const TW As Integer = 268
    Const TH As Integer = 24
    Const SP As Integer = 34

    Dim topY As Integer: topY = 14

    '── タイトルラベル ──
    With Lbl("lblTitle", Me.Caption, LX, topY, 400, 28)
        .Font.Size = 13: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 38

    '── 各フィールド ──
    Lbl "lblName", "取引先名", LX, topY + 4, LW, 18
    txt "txtName", "", "hiragana", TX, topY, TW, TH
    topY = topY + SP

    Lbl "lblZip", "郵便番号", LX, topY + 4, LW, 18
    With txt("txtZip3", "", "hankaku", TX, topY, 52, TH)
        .MaxLength = 3: .TextAlign = fmTextAlignCenter
    End With
    With Lbl("lblHyp", "－", TX + 58, topY + 4, 18, 18)
        .Font.Bold = True: .TextAlign = fmTextAlignCenter
    End With
    With txt("txtZip4", "", "hankaku", TX + 80, topY, 68, TH)
        .MaxLength = 4: .TextAlign = fmTextAlignCenter
    End With
    topY = topY + SP

    Lbl "lblAdr1", "住所①", LX, topY + 4, LW, 18
    txt "txtAdr1", "", "hiragana", TX, topY, TW, TH
    topY = topY + SP

    Lbl "lblAdr2", "住所②", LX, topY + 4, LW, 18
    txt "txtAdr2", "", "hiragana", TX, topY, TW, TH
    topY = topY + SP

    Lbl "lblTel", "電話番号", LX, topY + 4, LW, 18
    txt "txtTel", "", "hankaku", TX, topY, TW, TH
    topY = topY + SP

    Lbl "lblEmail", "メールアドレス", LX, topY + 4, LW, 18
    txt "txtEmail", "", "hankaku", TX, topY, TW, TH
    topY = topY + SP

    Lbl "lblContact", "担当者名", LX, topY + 4, LW, 18
    txt "txtContact", "", "hiragana", TX, topY, TW, TH
    topY = topY + SP + 8

    '── ボタン ──
    Const BW As Integer = 96
    Const BH As Integer = 30

    Dim bCancel As MSForms.CommandButton
    Set bCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    With bCancel
        .Caption = "キャンセル"
        .top = topY: .Left = Me.InsideWidth / 2 - BW - 16
        .Width = BW: .Height = BH
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
        .Font.Size = 11
    End With

    Dim bOK As MSForms.CommandButton
    Set bOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    With bOK
        .Caption = "保存"
        .top = topY: .Left = Me.InsideWidth / 2 + 16
        .Width = BW: .Height = BH
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 11
        .Font.Bold = True
    End With

    '── イベントハンドラ登録 ──
    Dim h As clsBtnHandler
    Set h = New clsBtnHandler
    Set h.btn = bCancel: h.act = "CANCEL_CLIENT_EDIT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = bOK: h.act = "OK_CLIENT_EDIT": Set h.frm = Me
    m_handlers.Add h

    Me.Height = topY + BH + 52
End Sub

'── ヘルパー関数 ──
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

Public Sub LoadData()
    If m_mode = "EDIT" Then
    
        Dim rs As Object
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM clients WHERE client_id = " & m_clientId, m_cn
        If Not rs.EOF Then
            Me.Controls("txtName").Text = rs("client_name") & ""
            Dim zip As String
            zip = rs("postal_code") & ""
            Dim zipArr As Variant
            zipArr = Split(zip, "-")
            If UBound(zipArr) >= 1 Then
                Me.Controls("txtZip3").Text = zipArr(0)
                Me.Controls("txtZip4").Text = zipArr(1)
            End If
            Me.Controls("txtAdr1").Text = rs("address1") & ""
            Me.Controls("txtAdr2").Text = rs("address2") & ""
            Me.Controls("txtTel").Text = rs("tel") & ""
            Me.Controls("txtEmail").Text = rs("email") & ""
            Me.Controls("txtContact").Text = rs("contact_name") & ""
        End If
        rs.Close
        Set rs = Nothing
    End If
End Sub
