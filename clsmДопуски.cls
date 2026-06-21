VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsmДопуски"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents txtNum As MSForms.TextBox
Attribute txtNum.VB_VarHelpID = -1
Public WithEvents LstIT As MSForms.ListBox
Attribute LstIT.VB_VarHelpID = -1

Private Sub lstIT_Click()
    
    Dim id As Long
    Dim i As Long
    
    Dim UpperDeviation As Double    'Верхнее отклонение
    Dim LowerDeviation As Double    'Нижнее отклонение
    Dim Tolerance As Double         'Допуск
       
    id = LstIT.List(LstIT.ListIndex, 1) * 1
    
    Dim ctrl As Control
    For Each ctrl In frmDopuski.Frame1.Controls
        If ctrl.Name <> LstIT.Name Then ctrl.ListIndex = -1
    Next
    
    With tbDeviation
        For i = 1 To .ListRows.Count
            If frmDopuski.txtSize_1 * 1 > .DataBodyRange(i, 1) And frmDopuski.txtSize_1 * 1 <= .DataBodyRange(i, 2) Then
                If .DataBodyRange(i, 3) = id Then
                    UpperDeviation = .DataBodyRange(i, 4) / 1000
                    LowerDeviation = .DataBodyRange(i + 1, 4) / 1000
                    Tolerance = UpperDeviation - LowerDeviation
                    Exit For
                End If
            End If
        Next
    End With
    
    With frmDopuski
        .txtUpperDeviation = IIf(UpperDeviation > 0, "+", "") & Format(UpperDeviation, "General Number")
        .txtLowerDeviation = IIf(LowerDeviation > 0, "+", "") & Format(LowerDeviation, "General Number")
        .txtTolerance_1 = Format(Tolerance, "General Number")
    End With
    
End Sub

Private Sub TxtNum_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        On Error Resume Next
        If txtNum.Value Like "=*" Then
            txtNum = Replace(txtNum, ",", ".")
            txtNum = Application.Evaluate(Replace(txtNum, "=", ""))
            txtNum = Replace(txtNum, ".", ",")
        End If
    End If
End Sub

Private Sub TxtNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii = 61 And Not txtNum Like "=*" Then Exit Sub
    
    Select Case KeyAscii
        Case 44, 46        ' обработка десятичного разделителя (44 - код запятой, 46 - код точки)
            KeyAscii = 44        ' если в поле  введена запятая, то заменим ее на точку
            If InStr(1, txtNum, ",") And Not txtNum Like "=*" Then KeyAscii = 0        ' вторую точку нельзя
    End Select


    If txtNum Like "=*" Then
        Select Case KeyAscii
            Case 40, 41, 42, 43, 45, 47, 94
                Exit Sub
            Case Else
                If KeyAscii <> 44 Then
                    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
                End If
        End Select
    Else
        If InStr(1, txtNum, ",") > 0 Then
            If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
        Else
            If KeyAscii <> 44 Then
                If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
            End If
        End If
    End If
    
End Sub

