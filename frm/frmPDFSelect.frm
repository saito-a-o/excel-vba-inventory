VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPDFSelect 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmPDFSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPDFSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================
' 【frmPDFSelect】─ ユーザーフォーム コードウィンドウ
'================================================================
 
Private m_handlers As Collection
 
Private Sub UserForm_Initialize()
    Set m_handlers = New Collection
 
    Me.Caption = "PDF出力 - シート選択"
    Me.Width = 390
    Me.BackColor = RGB(240, 244, 251)
 
    Const LX As Integer = 16
    Const SP As Integer = 30
    Dim topY As Integer: topY = 14
 
    '── タイトル ──
    With Lbl("lblTitle", "出力するシートを選択してください", LX, topY, 350, 22)
        .Font.Size = 11: .Font.Bold = True
        .ForeColor = RGB(31, 56, 100)
    End With
    topY = topY + 34
 
    '── 「入力」以外のシートをチェックボックスで列挙 ──
    Dim ws As Worksheet
    Dim idx As Integer: idx = 1
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "入力" Then
            Dim chk As MSForms.CheckBox
            Set chk = Me.Controls.Add("Forms.CheckBox.1", "chk_" & idx, True)
            With chk
                .Caption = ws.Name
                .top = topY
                .Left = LX + 10
                .Width = 320
                .Height = 22
                .Font.Size = 10
                .Value = True
            End With
            topY = topY + SP
            idx = idx + 1
        End If
    Next ws
    topY = topY + 6
 
    '── 区切り ──
    With Lbl("lblSep", "", LX, topY, 350, 1)
        .BackColor = RGB(180, 180, 180)
        .BackStyle = fmBackStyleOpaque
    End With
    topY = topY + 12
 
    '── 出力先フォルダ ──
    With Lbl("lblFolder", "出力先フォルダ", LX, topY + 5, 110, 18)
        .Font.Size = 10: .Font.Bold = True
    End With
 
    Dim txtF As MSForms.TextBox
    Set txtF = Me.Controls.Add("Forms.TextBox.1", "txtFolder", True)
    With txtF
        .Text = ""
        .top = topY
        .Left = LX + 118
        .Width = 170
        .Height = 24
        .Font.Size = 9
        .Locked = True
        .BackColor = RGB(245, 245, 245)
        .SpecialEffect = fmSpecialEffectSunken
    End With
 
    Dim btnBrowse As MSForms.CommandButton
    Set btnBrowse = Me.Controls.Add("Forms.CommandButton.1", "btnBrowse", True)
    With btnBrowse
        .Caption = "参照..."
        .top = topY: .Left = LX + 296
        .Width = 56: .Height = 24
        .Font.Size = 9
    End With
    topY = topY + 38
 
    '── キャンセル / PDF出力 ──
    Const BW As Integer = 96
    Const BH As Integer = 30
 
    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    With btnCancel
        .Caption = "キャンセル"
        .top = topY
        .Left = Me.InsideWidth / 2 - BW - 14
        .Width = BW: .Height = BH
        .BackColor = RGB(200, 200, 200)
        .Font.Size = 10
    End With
 
    Dim btnOutput As MSForms.CommandButton
    Set btnOutput = Me.Controls.Add("Forms.CommandButton.1", "btnOutput", True)
    With btnOutput
        .Caption = "PDF出力"
        .top = topY
        .Left = Me.InsideWidth / 2 + 14
        .Width = BW: .Height = BH
        .BackColor = RGB(47, 84, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 10: .Font.Bold = True
    End With
 
    '── イベント登録 ──
    Dim h As clsBtnHandler
    Set h = New clsBtnHandler: Set h.btn = btnBrowse:  h.act = "BROWSE":     Set h.frm = Me: m_handlers.Add h
    Set h = New clsBtnHandler: Set h.btn = btnCancel:  h.act = "CANCEL_PDF":  Set h.frm = Me: m_handlers.Add h
    Set h = New clsBtnHandler: Set h.btn = btnOutput:  h.act = "OUTPUT":      Set h.frm = Me: m_handlers.Add h
 
    Me.Height = topY + BH + 52
End Sub
 
Private Function Lbl(nm$, cap$, l%, t%, w%, h%) As MSForms.Label
    Set Lbl = Me.Controls.Add("Forms.Label.1", nm, True)
    With Lbl
        .Caption = cap: .Left = l: .top = t
        .Width = w: .Height = h: .Font.Size = 10
    End With
End Function
 
 
