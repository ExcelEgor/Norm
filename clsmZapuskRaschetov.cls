VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsmZapuskRaschetov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents txtRaschotPloshchadi As MSForms.TextBox
Attribute txtRaschotPloshchadi.VB_VarHelpID = -1

Public WithEvents cboGidroabraziv As MSForms.ComboBox
Attribute cboGidroabraziv.VB_VarHelpID = -1
Public WithEvents txtGidroabraziv As MSForms.TextBox
Attribute txtGidroabraziv.VB_VarHelpID = -1

Public WithEvents cboSvarka As MSForms.ComboBox
Attribute cboSvarka.VB_VarHelpID = -1
Public WithEvents txtSvarka As MSForms.TextBox
Attribute txtSvarka.VB_VarHelpID = -1
Public WithEvents chkSvarka As MSForms.CheckBox
Attribute chkSvarka.VB_VarHelpID = -1

Public WithEvents txtSlesar As MSForms.TextBox
Attribute txtSlesar.VB_VarHelpID = -1
Public WithEvents cboSlesar As MSForms.ComboBox
Attribute cboSlesar.VB_VarHelpID = -1
Public WithEvents optSlesar As MSForms.OptionButton
Attribute optSlesar.VB_VarHelpID = -1
Public WithEvents chkSlesar As MSForms.CheckBox
Attribute chkSlesar.VB_VarHelpID = -1
Public WithEvents mltSlesar As MSForms.MultiPage
Attribute mltSlesar.VB_VarHelpID = -1

Public WithEvents txtKoordinat As MSForms.TextBox
Attribute txtKoordinat.VB_VarHelpID = -1
Public WithEvents cboKoordinat As MSForms.ComboBox
Attribute cboKoordinat.VB_VarHelpID = -1
Public WithEvents optKoordinat As MSForms.OptionButton
Attribute optKoordinat.VB_VarHelpID = -1
Public WithEvents chkKoordinat As MSForms.CheckBox
Attribute chkKoordinat.VB_VarHelpID = -1
Public WithEvents mltKoordinat As MSForms.MultiPage
Attribute mltKoordinat.VB_VarHelpID = -1

Public WithEvents txtTermograf As MSForms.TextBox
Attribute txtTermograf.VB_VarHelpID = -1
Public WithEvents chkTermograf As MSForms.CheckBox
Attribute chkTermograf.VB_VarHelpID = -1

Private Sub cboGidroabraziv_Change()
    Call frmGidroabraziv.GlavnyyRaschot
End Sub

Private Sub cboKoordinat_Change()
    Call frmKoordinat.RaschotVremeni
End Sub

Private Sub chkKoordinat_Click()
    Call frmKoordinat.RaschotVremeni
End Sub

Private Sub chkTermograf_Click()
    Call frmTermografiya.RaschotTermografiya
End Sub

Private Sub mltKoordinat_Change()
    Call frmKoordinat.RaschotVremeni
End Sub

Private Sub optKoordinat_Click()
    Call frmKoordinat.RaschotVremeni
End Sub

Private Sub txtGidroabraziv_Change()
    Call frmGidroabraziv.GlavnyyRaschot
End Sub

Private Sub txtKoordinat_Change()
    Call frmKoordinat.RaschotVremeni
End Sub

Private Sub txtMetallograf_Change()
    frmMetallografika.lstMetallografika.Clear
End Sub

Private Sub txtRaschotPloshchadi_Change()
    frmRaschotPloshchadi.GlavnyyRaschot
End Sub

Private Sub txtTermograf_Change()
    Call frmTermografiya.RaschotTermografiya
End Sub

