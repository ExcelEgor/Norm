VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsmMass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Public WithEvents mlt As MSForms.MultiPage
Attribute mlt.VB_VarHelpID = -1
Public WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1

Public Property Set TextBoxControl(vControl As MSForms.TextBox): Set txt = vControl: End Property

Public Property Set MultiPageControl(vControl As MSForms.MultiPage): Set mlt = vControl: End Property

Public Property Set ComboBoxControl(vControl As MSForms.ComboBox): Set cbo = vControl: End Property

Private Sub cbo_Change(): Call frmMass.GlavnyyRaschot: End Sub

Private Sub mlt_Change(): Call frmMass.GlavnyyRaschot: End Sub

Private Sub txt_Change(): Call frmMass.GlavnyyRaschot: End Sub



