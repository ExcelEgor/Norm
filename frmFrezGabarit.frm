VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFrezGabarit 
   Caption         =   "Фрезерование габаритов"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8460.001
   OleObjectBlob   =   "frmFrezGabarit.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFrezGabarit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Handler As clsFrezerovaniyeGabaritov
Private ControlHandlers As Collection

Private Sub UserForm_Initialize()
    
    Call ZapolneniyeElementovUpravleniya
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    Call ZagruzkaElementovVKlass

End Sub

Private Sub ZapolneniyeElementovUpravleniya()

    Dim MATERIALY
    MATERIALY = Array(EnumMaterialy.ALUMINIYEVYYE_SPLAVY, EnumMaterialy.MEDNYYE_SPLAVY, EnumMaterialy.STAL_UGLERODISTAYA, _
        EnumMaterialy.STAL_LEGIROVANNAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA, EnumMaterialy.TITANOVYYE_SPLAVY)
    
    Call DobavitMaterialyVListBox(MATERIALY, cboMaterial)
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.Name Like "cboRa_*" Then
            ctrl.List = Split("20 10 5 2,5 1,25")
            ctrl.ListIndex = 1
        ElseIf ctrl.Name Like "cboFrez_*" Then
            ctrl.List = Split("Концевая Торцевая")
            If ctrl.Name = "cboFrez_h" Then
                ctrl.ListIndex = 1
            Else
                ctrl.ListIndex = 0
            End If
        End If
    Next
    
End Sub

Private Sub ZagruzkaElementovVKlass()
    
    Set ControlHandlers = New Collection
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If LCase(ctrl.Tag) <> "notclsm" Then
            Set Handler = New clsFrezerovaniyeGabaritov
            If TypeOf ctrl Is MSForms.TextBox Then
                Set Handler.TextBoxControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.TabStrip Then
                Set Handler.TabStripControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.ComboBox Then
                Set Handler.ComboBoxControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.CheckBox Then
                Set Handler.CheckBoxControl = ctrl
            End If
            
            ControlHandlers.Add Handler
            
        End If
    Next
    
End Sub
