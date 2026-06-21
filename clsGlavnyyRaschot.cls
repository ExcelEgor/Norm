VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGlavnyyRaschot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents mlt As MSForms.MultiPage
Attribute mlt.VB_VarHelpID = -1
Private WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1
Private WithEvents opt As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1
Private WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1

Public Property Set TextBoxControl(vControl As MSForms.TextBox): Set txt = vControl: End Property

Public Property Set MultiPageControl(vControl As MSForms.MultiPage): Set mlt = vControl: End Property

Public Property Set CheckBoxControl(vControl As MSForms.CheckBox): Set chk = vControl: End Property

Public Property Set OptionButtonControl(vControl As MSForms.OptionButton): Set opt = vControl: End Property

Public Property Set ComboBoxControl(vControl As MSForms.ComboBox): Set cbo = vControl: End Property

Private Sub mlt_Change(): GetForm.GlavnyyRaschot: End Sub

Private Sub opt_Click(): GetForm.GlavnyyRaschot: End Sub

Private Sub txt_Change(): GetForm.GlavnyyRaschot: End Sub

Private Sub chk_Click(): GetForm.GlavnyyRaschot: End Sub

Private Sub cbo_Change(): GetForm.GlavnyyRaschot: End Sub

Private Function GetForm() As Object
    
    Dim ParentObject As Object
    
    If Not txt Is Nothing Then
        Set ParentObject = txt.Parent
    ElseIf Not mlt Is Nothing Then
        Set ParentObject = mlt.Parent
    ElseIf Not chk Is Nothing Then
        Set ParentObject = chk.Parent
    ElseIf Not opt Is Nothing Then
        Set ParentObject = opt.Parent
    ElseIf Not cbo Is Nothing Then
        Set ParentObject = cbo.Parent
    End If
    
    If Not ParentObject Is Nothing Then
        Do While TypeName(ParentObject) = "Frame" _
            Or TypeName(ParentObject) = "MultiPage" _
            Or TypeName(ParentObject) = "Page"
            Set ParentObject = ParentObject.Parent
        Loop
        
        Set GetForm = ParentObject
    End If
    
End Function
