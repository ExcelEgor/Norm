Attribute VB_Name = "mДопускиПосадки"
Option Explicit

Public tbSize As ListObject

Public tbDeviation As ListObject

Sub FindDeviation()

    Dim QualityFind As Boolean
    
    Dim lstDeviation As Control
    Dim txtSize_2 As Control
    Dim txtTolerance_2 As Control
    Dim txtDeviation As Control
    
    Dim i As Long
    
    With frmDopuski
        Set lstDeviation = .lstDeviation
        Set txtSize_2 = .txtSize_2
        Set txtTolerance_2 = .txtTolerance_2
        Set txtDeviation = .txtDeviation
    End With
    
    lstDeviation.Clear
    txtTolerance_2 = ""
    QualityFind = False
    
    If Not IsNumeric(txtSize_2) Or Not IsNumeric(txtDeviation) Then Exit Sub
    
    With tbSize
        For i = 1 To .ListRows.Count
            If txtSize_2 * 1 > .DataBodyRange(i, 1) And txtSize_2 * 1 <= .DataBodyRange(i, 2) Then
                lstDeviation.AddItem
                lstDeviation.List(lstDeviation.ListCount - 1, 0) = .DataBodyRange(i, 4)
                lstDeviation.List(lstDeviation.ListCount - 1, 1) = CStr(.DataBodyRange(i, 3) / 1000)
                If QualityFind = False Then
                    If txtDeviation * 1000 = .DataBodyRange(i, 3) Then
                        txtTolerance_2 = "Квалитет равен " & .DataBodyRange(i, 4): QualityFind = True
                    ElseIf txtDeviation * 1000 > .DataBodyRange(i, 3) And txtDeviation * 1 < .DataBodyRange(i - 1, 3) Then
                        txtTolerance_2 = "Квалитет между " & .DataBodyRange(i, 4) & " и " & .DataBodyRange(i - 1, 4): QualityFind = True
                    ElseIf txtDeviation * 1000 < .DataBodyRange(i, 3) And .DataBodyRange(i + 1, 3) > .DataBodyRange(i, 3) Then
                        txtTolerance_2 = "Квалитет ниже " & .DataBodyRange(i, 4): QualityFind = True
                    ElseIf txtDeviation * 1000 > .DataBodyRange(i, 3) And (.DataBodyRange(i - 1, 3) < .DataBodyRange(i, 3) Or i = 1) Then
                        txtTolerance_2 = "Квалитет выше " & .DataBodyRange(i, 4): QualityFind = True
                    End If
                End If
                
            End If
        Next
    End With
    
End Sub

Function DopuskPoKvalitetu(ByVal razmer As Double, ByVal Kvalitet As Integer) As Double
    
    Dim tbDopuskiPosadki_Razmery As ListObject: Set tbDopuskiPosadki_Razmery = wsDopuski.ListObjects("tbDopuskiPosadki_Razmery")
    Dim tbDopuskiPosadki_Otkloneniya As ListObject: Set tbDopuskiPosadki_Otkloneniya = wsDopuski.ListObjects("tbDopuskiPosadki_Otkloneniya")
    
    Dim i As Long
    
    With tbDopuskiPosadki_Razmery
        For i = 1 To .ListRows.Count
            If .DataBodyRange(i, 1) >= razmer Then
                razmer = .DataBodyRange(i, 1)
                Exit For
            End If
        Next
    End With
    
    Dim id As Long
    id = CLng(razmer & Kvalitet)
    With tbDopuskiPosadki_Otkloneniya
        DopuskPoKvalitetu = .DataBodyRange(WorksheetFunction.Match(id, .ListColumns(5).DataBodyRange, 0), 3)
    End With
  

End Function


















