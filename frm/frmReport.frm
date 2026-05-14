VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_handlers     As Collection
Private m_dateHandlers As Collection

Private Sub UserForm_Initialize()
    Set m_handlers = New Collection
    Set m_dateHandlers = New Collection

    '── フォーム外観 ──
    Me.Caption = "レポート出力"
    Me.Width = 380
    Me.BackColor = RGB(240, 244, 251)

    Const LX As Integer = 16
    Const TX As Integer = 130
    Const TW As Integer = 200
    Const TH As Integer = 24
    Const SP As Integer = 34
    Dim topY As Integer: topY = 14

    '── タイトル ──
    With Lbl("lblTitle", "レポート出力", LX, topY, 340, 28)
        .Font.Size = 13: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 38

    '── 区切り ──
    With Lbl("lblSep1", "", LX, topY, 335, 1)
        .BackColor = RGB(180, 180, 180)
        .BackStyle = fmBackStyleOpaque
    End With
    topY = topY + 12

    '── 集計期間セクション ──
    With Lbl("lblSec1", "集計期間", LX, topY, 200, 20)
        .Font.Size = 10: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 28

    '── 開始日 ──
    Lbl "lblStart", "開始日", LX, topY + 4, 110, 18
    Dim txtStart As MSForms.TextBox
    Set txtStart = Me.Controls.Add("Forms.TextBox.1", "txtStart", True)
    With txtStart
        .Left = TX: .top = topY
        .Width = TW: .Height = TH
        .Font.Size = 10
        .IMEMode = xlIMEModeDisable
        .Text = Format(Date, "yyyy/mm/dd")
        .ForeColor = RGB(180, 180, 180)
    End With
    topY = topY + SP

    '── 終了日 ──
    Lbl "lblEnd", "終了日", LX, topY + 4, 110, 18
    Dim txtEnd As MSForms.TextBox
    Set txtEnd = Me.Controls.Add("Forms.TextBox.1", "txtEnd", True)
    With txtEnd
        .Left = TX: .top = topY
        .Width = TW: .Height = TH
        .Font.Size = 10
        .IMEMode = xlIMEModeDisable
        .Text = Format(Date, "yyyy/mm/dd")
        .ForeColor = RGB(180, 180, 180)
    End With
    topY = topY + SP + 8

    '── 区切り ──
    With Lbl("lblSep2", "", LX, topY, 335, 1)
        .BackColor = RGB(180, 180, 180)
        .BackStyle = fmBackStyleOpaque
    End With
    topY = topY + 12

    '── 出力先セクション ──
    With Lbl("lblSec2", "出力先フォルダ", LX, topY, 200, 20)
        .Font.Size = 10: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 28

    '── フォルダパス ──
    Dim txtFolder As MSForms.TextBox
    Set txtFolder = Me.Controls.Add("Forms.TextBox.1", "txtFolder", True)
    With txtFolder
        .Left = LX: .top = topY
        .Width = 260: .Height = TH
        .Font.Size = 9
        .Locked = True
        .BackColor = RGB(245, 245, 245)
        .SpecialEffect = fmSpecialEffectSunken
    End With

    Dim btnBrowse As MSForms.CommandButton
    Set btnBrowse = Me.Controls.Add("Forms.CommandButton.1", "btnBrowse", True)
    With btnBrowse
        .Caption = "参照..."
        .Left = LX + 268: .top = topY
        .Width = 56: .Height = TH
        .Font.Size = 9
    End With
    topY = topY + SP + 8

    '── ボタン ──
    Const BW As Integer = 96
    Const BH As Integer = 30
    Const BTop As Integer = 10

    Dim bCancel As MSForms.CommandButton
    Set bCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancelReport", True)
    With bCancel
        .Caption = "キャンセル"
        .top = topY + BTop
        .Left = Me.InsideWidth / 2 - BW - 16
        .Width = BW: .Height = BH
        .BackColor = RGB(200, 200, 200)
        .ForeColor = RGB(50, 50, 50)
        .Font.Size = 11
    End With

    Dim bOutput As MSForms.CommandButton
    Set bOutput = Me.Controls.Add("Forms.CommandButton.1", "btnOutputReport", True)
    With bOutput
        .Caption = "出力"
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
    Set h.btn = btnBrowse: h.act = "BROWSE_REPORT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = bCancel: h.act = "CANCEL_REPORT": Set h.frm = Me
    m_handlers.Add h

    Set h = New clsBtnHandler
    Set h.btn = bOutput: h.act = "OUTPUT_REPORT": Set h.frm = Me
    m_handlers.Add h

    '── 日付TextBoxハンドラ登録 ──
    Dim hd As clsDateTxtHandler
    Set hd = New clsDateTxtHandler
    Set hd.txt = txtStart
    m_dateHandlers.Add hd

    Set hd = New clsDateTxtHandler
    Set hd.txt = txtEnd
    m_dateHandlers.Add hd

    Me.Height = topY + BH + BTop + 52
End Sub

Private Function Lbl(nm$, cap$, l%, t%, w%, h%) As MSForms.Label
    Set Lbl = Me.Controls.Add("Forms.Label.1", nm, True)
    With Lbl
        .Caption = cap: .Left = l: .top = t
        .Width = w: .Height = h: .Font.Size = 10
    End With
End Function
