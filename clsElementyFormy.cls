VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsElementyFormy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParams As Object
Private pForma As MSForms.UserForm

Private Sub Class_Initialize()
    Set pParams = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Init(frm As MSForms.UserForm)
    Set pForma = frm
End Sub

Public Sub SaveFromForm()

    pParams.RemoveAll

    ' —охран€ем значени€ контролов активной страницы mlt
    Dim ctrl As Control
    For Each ctrl In pForma.Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            pParams(ctrl.Name) = ctrl.ListIndex
        ElseIf TypeOf ctrl Is MSForms.TextBox Or _
            TypeOf ctrl Is MSForms.CheckBox Or _
            TypeOf ctrl Is MSForms.OptionButton Or _
            TypeOf ctrl Is MSForms.MultiPage Then
            On Error Resume Next
            pParams(ctrl.Name) = ctrl.Value
            On Error GoTo 0
        End If
    Next ctrl
    
End Sub

Public Sub LoadToForm()

    Dim Key As Variant
    For Each Key In pParams.Keys
        On Error Resume Next
        pForma.Controls(Key).Value = pParams(Key)
        On Error GoTo 0
    Next Key
    
End Sub

Public Sub SaveToExcel(wsSave As Worksheet)
    
    Dim Key As Variant
    Dim i As Long
    i = 1
    For Each Key In pParams.Keys
        wsSave.Cells(i, 1) = pForma.Controls(Key).Name
        wsSave.Cells(i, 2) = pForma.Controls(Key).Value
        i = i + 1
    Next Key
        
End Sub
