VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTxtHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Public rowIdx As Integer ' 何行目のTextBoxか
Public frm As Object '     親フォーム（frmItemSelect）


Private Sub txt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                          ByVal X As Single, ByVal Y As Single)
    frm.SelectRow rowIdx
End Sub
