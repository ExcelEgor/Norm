VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsmKontrolVvodaChisel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Public WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1

Public WithEvents txtClearOnFocus As MSForms.TextBox
Attribute txtClearOnFocus.VB_VarHelpID = -1
Public oldTxt As Control


Private Sub Cbo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call EvaluateKeyDown(KeyCode, cbo)
End Sub

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call EvaluateKeyDown(KeyCode, txt)
End Sub

Private Sub txt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If VvedenoChislo(KeyAscii, txt) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub EvaluateKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ctrl As Control)
    'Если в элемент управления введена формула, то произвести расчет по нажатию клавишы Enter
    
    If KeyCode = vbKeyReturn Then
        On Error Resume Next
        If ctrl.text Like "=*" Then
            ctrl.text = Replace(ctrl.text, ",", ".")
            ctrl.text = Application.Evaluate(Replace(ctrl.text, "=", ""))
            ctrl.text = Replace(ctrl.text, ".", ",")
        End If
    End If
    
End Sub

Function VvedenoChislo(ByVal KeyAscii As MSForms.ReturnInteger, ctrl As Control) As Boolean
    'KeyAscii - код нажатой клавиши
    '40 (
    '41 )
    '42 *
    '43 +
    '44 ,
    '45 -
    '46 .
    '47 /
    '94 ^
    
    VvedenoChislo = True
    Select Case KeyAscii
        'Если нажата клавина "=", то очистить элемент управления и написать "="
        Case 61
            ctrl.text = Empty
            'Если в поле  введена точка, то заменить ее на запятую
        Case 44, 46
            KeyAscii = 44
            'Вторую точку нельзя
            If InStr(1, ctrl.text, ",") And Not ctrl.text Like "=*" Then
                VvedenoChislo = False
            End If
            'Разрешать математические символы если введена формула
        Case 40, 41, 42, 43, 45, 47, 94
            If ctrl.text Like "=*" Then
                Exit Function
            Else
                VvedenoChislo = False
            End If
        Case 48 To 57
            Exit Function
        Case Else
            VvedenoChislo = False
    End Select
    
End Function
