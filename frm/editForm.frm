VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} editForm 
   Caption         =   "発行者情報の編集"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "editForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "editForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ws       As Worksheet
Private m_handlers As Collection   ' ボタンイベント保持用

'────────────────────────────────────────────────────────────────
'  フォーム初期化（コントロールをすべて動的に生成します）
'────────────────────────────────────────────────────────────────
Private Sub UserForm_Initialize()
    Set m_ws = ThisWorkbook.Sheets("入力")
    Set m_handlers = New Collection

    '── フォーム外観 ──
    Me.Caption = "発行者情報の編集"
    Me.Width = 460
    Me.BackColor = RGB(240, 244, 251)

    '── レイアウト定数 ──
    Const LX As Integer = 14   ' ラベル左端
    Const TX As Integer = 160  ' テキストボックス左端
    Const LW As Integer = 140  ' ラベル幅
    Const TW As Integer = 268  ' テキストボックス幅
    Const TH As Integer = 24   ' コントロール高さ
    Const SP As Integer = 34   ' 行間隔

    Dim topY As Integer: topY = 14
    Dim hira As String: hira = "hiragana"
    Dim hank As String: hank = "hankaku"
    '── タイトルラベル ──
    With Lbl("lblTitle", "発行者情報の編集", LX, topY, 400, 28)
        .Font.Size = 13: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 38

    '── 会社名 ──
    Lbl "lblComp", "会社名", LX, topY + 4, LW, 18
    txt "txtComp", m_ws.Cells(4, 2).Value, hira, TX, topY, TW, TH
    topY = topY + SP

    '── 郵便番号 ──
    Lbl "lblZip", "郵便番号", LX, topY + 4, LW, 18
    With txt("txtZip3", "", hank, TX, topY, 52, TH)
        .MaxLength = 3: .TextAlign = fmTextAlignCenter
    End With
    With Lbl("lblHyp", "－", TX + 58, topY + 4, 18, 18)
        .Font.Bold = True: .TextAlign = fmTextAlignCenter
    End With
    With txt("txtZip4", "", hank, TX + 80, topY, 68, TH)
        .MaxLength = 4: .TextAlign = fmTextAlignCenter
    End With

    ' 現在値を読み込む
    Dim z3 As String, c As Integer
    For c = 3 To 5:  z3 = z3 & m_ws.Cells(5, c).Value: Next c
    Me.Controls("txtZip3").Text = z3
    Dim z4 As String
    For c = 7 To 10: z4 = z4 & m_ws.Cells(5, c).Value: Next c
    Me.Controls("txtZip4").Text = z4
    topY = topY + SP

    '── 住所① ──
    Lbl "lblAdr1", "住所①", LX, topY + 4, LW, 18
    txt "txtAdr1", m_ws.Cells(6, 2).Value, hira, TX, topY, TW, TH
    topY = topY + SP

    '── 住所② ──
    Lbl "lblAdr2", "住所②", LX, topY + 4, LW, 18
    txt "txtAdr2", m_ws.Cells(7, 2).Value, hira, TX, topY, TW, TH
    topY = topY + SP

    '── 電話番号 ──
    Lbl "lblTel", "電話番号", LX, topY + 4, LW, 18
    txt "txtTel", m_ws.Cells(8, 2).Value, hank, TX, topY, TW, TH
    topY = topY + SP

    '── FAX番号 ──
    Lbl "lblFax", "FAX番号", LX, topY + 4, LW, 18
    txt "txtFax", m_ws.Cells(9, 2).Value, hank, TX, topY, TW, TH
    topY = topY + SP

    '── インボイス登録番号 ──
    Lbl "lblInv", "インボイス登録番号", LX, topY + 4, LW, 18
    txt "txtInv", m_ws.Cells(10, 2).Value, hank, TX, topY, TW, TH
    topY = topY + SP

    '── 担当者 ──
    Lbl "lblPrs", "担当者", LX, topY + 4, LW, 18
    txt "txtPrs", m_ws.Cells(11, 2).Value, hira, TX, topY, TW, TH
    topY = topY + SP + 8
    
    '── 振込先：銀行名 ──
    Lbl "lblBank", "銀行名・支店名", LX, topY + 4, LW, 18
    txt "txtBank", "", hira, TX, topY, TW, TH
    topY = topY + SP

    '── 振込先：口座種別 ──
    Lbl "lblAccType", "口座種別", LX, topY + 4, LW, 18
    txt "txtAccType", "", hira, TX, topY, TW, TH
    topY = topY + SP

    '── 振込先：口座番号 ──
    Lbl "lblAccNum", "口座番号", LX, topY + 4, LW, 18
    txt "txtAccNum", "", hank, TX, topY, TW, TH
    topY = topY + SP

    '── 振込先：口座名義 ──
    Lbl "lblAccName", "口座名義", LX, topY + 4, LW, 18
    txt "txtAccName", "", hira, TX, topY, TW, TH
    topY = topY + SP + 8

    '── DBから発行者情報を読み込む
    Dim cn As Object
    Dim rs As Object
    Set cn = GetConnection
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM my_company WHERE company_id = 1;", cn
    
    If Not rs.EOF Then
        Me.Controls("txtBank").Text = rs("bank_name") & ""
        Me.Controls("txtAccType").Text = rs("account_type") & ""
        Me.Controls("txtAccNum").Text = rs("account_number") & ""
        Me.Controls("txtAccName").Text = rs("account_name") & ""
    End If
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

    '── ボタン（終了 / 確定） ──
    Const BW As Integer = 96
    Const BH As Integer = 30
    Const BTop As Integer = 10  ' ボタン上余白

    Dim bCancel As MSForms.CommandButton
    Set bCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    With bCancel
        .Caption = "終了"
        .top = topY + BTop
        .Left = Me.InsideWidth / 2 - BW - 16
        .Width = BW: .Height = BH
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
        .Font.Size = 11
    End With

    Dim bOK As MSForms.CommandButton
    Set bOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    With bOK
        .Caption = "確定"
        .top = topY + BTop
        .Left = Me.InsideWidth / 2 + 16
        .Width = BW: .Height = BH
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 11
        .Font.Bold = True
    End With

    '── イベントハンドラをコレクションに登録 ──
    Dim h As clsBtnHandler
    Set h = New clsBtnHandler
    Set h.btn = bCancel: h.act = "CANCEL": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = bOK: h.act = "OK": Set h.frm = Me
    m_handlers.Add h

    Me.Height = topY + BH + BTop + 52
End Sub

'────────────────────────────────────────────────────────────────
'  ヘルパー関数（ラベル / テキストボックス生成）
'────────────────────────────────────────────────────────────────
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

