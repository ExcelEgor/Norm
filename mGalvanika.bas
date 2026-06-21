Attribute VB_Name = "mGalvanika"
Option Explicit

Function GalvanikaNovayaFunktsiya(Pokrytiye As String, ByVal PloshchadM2 As Double, Optional ByVal KolVo As Double, Optional Partiya As Boolean) As Variant
    
    If KolVo = 0 Then KolVo = 1
    If PolozhitelnyyeChisla(PloshchadM2, KolVo) = False Then Exit Function
    If PloshchadM2 > Galvanika_MaksPloshchadM2(Pokrytiye) Then Exit Function

    Dim Normativ
    Normativ = ZagruzitNormativ("galvanika")
    
    PloshchadM2 = PloshchadM2 / KolVo
    
    Dim Normy(1 To 2) As Double

    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = Pokrytiye Then
            If Normativ(i, 2) >= PloshchadM2 Then
                Normy(1) = IIf(Partiya, 1, 2) * KolVo * Normativ(i, 3)
                Exit For
            End If
            If Normativ(i, 1) <> Normativ(i + 1, 1) Then Exit For
        End If
    Next
    
    If Normy(1) > 0 Then
        Normy(2) = IIf(Partiya, 0.7, 1.3) * KolVo * VisualnyyKontrolPoPloshadi(PloshchadM2 * PloshchadM2)
    End If
    
    GalvanikaNovayaFunktsiya = Normy
    
End Function

Function Galvanika_MaksPloshchadM2(Pokrytiye As String) As Double

    Dim Normativ
    Normativ = ZagruzitNormativ("galvanika")

    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = Pokrytiye Then
            If Normativ(i, 2) > Galvanika_MaksPloshchadM2 Then Galvanika_MaksPloshchadM2 = Normativ(i, 2)
            If PoslednyayaStrokaVMassive(Normativ, i, 1) Then Exit Function
        End If
    Next
    
End Function

Function Galvanika(Pokrytiye As String, ByVal PloshchadM2 As Double) As Double

    If PolozhitelnyyeChisla(PloshchadM2) = False Then Exit Function

    Dim Normativ
    Normativ = ZagruzitNormativ("galvanika")

    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = Pokrytiye Then
            If Normativ(i, 2) >= PloshchadM2 Then
                Galvanika = Normativ(i, 3)
                Exit For
            End If
            If Normativ(i, 1) <> Normativ(i + 1, 1) Then Exit For
        End If
    Next
    
End Function

Function OpredelitRazryad_Galvanika(Pokrytiye As String) As Integer

    Select Case Pokrytiye
        Case "Обезжиривание", "Травление"
            OpredelitRazryad_Galvanika = 2
        Case "Ан. Окс. тв"
            OpredelitRazryad_Galvanika = 4
        Case Else
            OpredelitRazryad_Galvanika = 3
    End Select
    
End Function

