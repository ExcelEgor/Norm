VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsmShiftEnter_Elektroeroziya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Public WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
Public WithEvents cmd As MSForms.CommandButton
Attribute cmd.VB_VarHelpID = -1
Public WithEvents tgl As MSForms.ToggleButton
Attribute tgl.VB_VarHelpID = -1
Public WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1
Public WithEvents opt As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1
Public WithEvents tabs As MSForms.TabStrip
Attribute tabs.VB_VarHelpID = -1
Public WithEvents lst As MSForms.ListBox
Attribute lst.VB_VarHelpID = -1
Public WithEvents img As MSForms.Image
Attribute img.VB_VarHelpID = -1
Public WithEvents fra As MSForms.Frame
Attribute fra.VB_VarHelpID = -1
Public WithEvents Spn As MSForms.SpinButton
Attribute Spn.VB_VarHelpID = -1

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub

Private Sub Cbo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub

Private Sub Cmd_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub

Private Sub Tgl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub

Private Sub Chk_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub

Private Sub Opt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub

Private Sub Tabs_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub

Private Sub Lst_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub

Private Sub Img_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub

Private Sub Fra_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub

Private Sub Spn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmElektroEroziya.AddInTable
End Sub


 
