VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTokFrez_Frm"
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
Public WithEvents opt As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1
Public WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1
Public WithEvents tabs As MSForms.TabStrip
Attribute tabs.VB_VarHelpID = -1
Public WithEvents lst As MSForms.ListBox
Attribute lst.VB_VarHelpID = -1
Public WithEvents fra As MSForms.Frame
Attribute fra.VB_VarHelpID = -1
Public WithEvents Spn As MSForms.SpinButton
Attribute Spn.VB_VarHelpID = -1

Public WithEvents LblDrugieProg As MSForms.Label
Attribute LblDrugieProg.VB_VarHelpID = -1


Private Sub cbo_Change()
    If frmTokFrez.StopEvents Then Exit Sub
    If cbo.text <> "" Or cbo.Tag = "num" Then Call GlavnyyRaschot_TokFrez
End Sub

Private Sub cbo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If cbo.Tag = "num" Then
        Select Case KeyAscii
            Case 48 To 57, 8    ' цифры 0 - 9 и <Backspace> (эти символы всегда разрешены)
            Case 44, 46         ' обработка десятичного разделителя (44 - код запятой, 46 - код точки)
                KeyAscii = 44   ' если в поле  введена точка, то заменим ее на запятую
                If InStr(1, cbo.text, ",") Then KeyAscii = 0 ' вторую точку нельзя
            Case Else
                KeyAscii = 0    ' остальные символы запрещены
        End Select
    End If
    
End Sub

Private Sub chk_Click()
    If frmTokFrez.StopEvents Then Exit Sub
    Call GlavnyyRaschot_TokFrez
End Sub

Private Sub Chk_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmTokFrez.DobavitVTablitsu
End Sub

Private Sub Cmd_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmTokFrez.DobavitVTablitsu
End Sub


Private Sub Fra_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 1 Then Call frmTokFrez.DobavitVTablitsu
End Sub

Private Sub opt_Click()
    If frmTokFrez.StopEvents Then Exit Sub
    Call GlavnyyRaschot_TokFrez
End Sub

Private Sub Tabs_Change()
    If frmTokFrez.StopEvents Or frmTokFrez.ZagruzkaFormy Then Exit Sub
    Call GlavnyyRaschot_TokFrez
End Sub

Private Sub Tgl_Click()
    If frmTokFrez.StopEvents Then Exit Sub
    Call GlavnyyRaschot_TokFrez
End Sub

Private Sub txt_Change()

    If frmTokFrez.StopEvents Then Exit Sub
    
    Select Case True
        Case txt.Name Like "txtSum*"
        Case txt.Name Like "txtTime*"
            txt.text = txt.text
        Case txt.Name Like "txtDesc*"
        Case txt.Name Like "txtFileName*"
        Case txt.Name Like "txtZachistZaus_Shlic"
        Case Else
            Call GlavnyyRaschot_TokFrez
    End Select

End Sub

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 13 And Shift = 1 Then Call frmTokFrez.DobavitVTablitsu
    
    Dim num As Double
    If KeyCode = vbKeyReturn Then
        On Error Resume Next
        If txt.Value Like "=*" Then
            txt = Replace(txt, ",", ".")
            num = Application.Evaluate(Replace(txt, "=", ""))
            txt = Format(num, "0.00")
        End If
    End If
    
End Sub

Private Sub txt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case True
    
        Case txt.Name Like "txtTime*"
        Case txt.Name Like "txtDesc*"
        Case txt.Name Like "txtFileName*"
        Case Else
    
            If KeyAscii = 61 And txt.text = Empty Then Exit Sub
        
            Select Case KeyAscii
                Case 44, 46        ' обработка десятичного разделителя (44 - код запятой, 46 - код точки)
                    KeyAscii = 44        ' если в поле  введена запятая, то заменим ее на точку
                    If InStr(1, txt, ",") And Not txt Like "=*" Then KeyAscii = 0        ' вторую точку нельзя
            End Select
    
    
            If txt Like "=*" Then
                Select Case KeyAscii
                    Case 40, 41, 42, 43, 45, 47, 94
                        Exit Sub
                    Case Else
                        If KeyAscii <> 44 Then
                            If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
                        End If
                End Select
            Else
                If InStr(1, txt, ",") > 0 Then
                    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
                Else
                    If KeyAscii <> 44 Then
                        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
                    End If
                End If
            End If
    
    End Select
    
End Sub


