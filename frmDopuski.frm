VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDopuski 
   Caption         =   "Допуски и посадки"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9075.001
   OleObjectBlob   =   "frmDopuski.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDopuski"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbQuality As ListObject
Dim ArrLstIT() As New clsmДопуски

Private Sub txtDeviation_Change()
    Call FindDeviation
End Sub

Private Sub txtSize_2_Change()
    Call FindDeviation
End Sub

Private Sub txtSize_1_Change()
    
    Dim lstQuality As Control
    Dim ctrl As Control
    Dim i As Long
    
    For Each ctrl In Frame1.Controls
        ctrl.Clear
    Next

    txtUpperDeviation = ""
    txtLowerDeviation = ""
    txtTolerance_1 = ""
    
    If Not IsNumeric(txtSize_1) Then Exit Sub
    
    With tbQuality
        For i = 1 To .ListRows.Count
            If txtSize_1 * 1 > .DataBodyRange(i, 3) And txtSize_1 * 1 <= .DataBodyRange(i, 4) Then
                Select Case .DataBodyRange(i, 5)
                    Case 5: Set lstQuality = ListBox1
                    Case 6: Set lstQuality = ListBox2
                    Case 7: Set lstQuality = ListBox3
                    Case 8: Set lstQuality = ListBox4
                    Case 9: Set lstQuality = ListBox5
                    Case 10: Set lstQuality = ListBox6
                    Case 11: Set lstQuality = ListBox7
                    Case 12: Set lstQuality = ListBox8
                    Case 14: Set lstQuality = ListBox9
                    Case 16: Set lstQuality = ListBox10
                    Case 17: Set lstQuality = ListBox11
                End Select
                With lstQuality
                    .AddItem
                    .List(.ListCount - 1, 0) = tbQuality.DataBodyRange(i, 2)
                    .List(.ListCount - 1, 1) = tbQuality.DataBodyRange(i, 1)
                End With
            End If
        Next
    End With
    
End Sub
Private Sub UserForm_Initialize()
       
    Dim ctrl As Control, fraCtrl As Control
    
    With wsDopuski
        Set tbQuality = .ListObjects("Квалитет")
        Set tbDeviation = .ListObjects("Отклонение")
        Set tbSize = .ListObjects("tbDopuskiPosadki_Otkloneniya")
    End With
       
    Dim iTxt As Integer, iLst As Integer
    iTxt = 1: iLst = 1
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ReDim Preserve ArrTxtKontrolVvodaChisel(iTxt_KontrolVvodaChisel)
            Set ArrTxtKontrolVvodaChisel(iTxt_KontrolVvodaChisel).txt = ctrl
        ElseIf ctrl.Name = "Frame1" Then
            For Each fraCtrl In Frame1.Controls
                iLst = iLst + 1
                ReDim Preserve ArrLstIT(iLst)
                Set ArrLstIT(iLst).LstIT = fraCtrl
            Next
        End If
    Next
    
    Dim W As Double
    W = 39
    For Each ctrl In Frame1.Controls
        ctrl.Width = W
        ctrl.Left = (CInt(Replace(ctrl.Name, "ListBox", "")) - 1) * ctrl.Width
    Next
    
    Frame1.Width = (W * 11) + 9
    
End Sub
