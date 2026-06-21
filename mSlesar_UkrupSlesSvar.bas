Attribute VB_Name = "mSlesar_UkrupSlesSvar"
Option Explicit
'ÓÊÐÓÏÍÅÍÍÛÅ ÎÄÍÎÑÒÐÎ×ÍÛÅ ÍÎÐÌÀÒÈÂÛ ÂÐÅÌÅÍÈ ÍÀ ÑËÅÑÀÐÍÛÅ, ÑËÅÑÀÐÍÎ-ÊÀÐÊÀÑÍÛÅ È ÑÂÀÐÎ×ÍÛÅ ÐÀÁÎÒÛ Â ÏÐÈÁÎÐÎÑÒÐÎÅÍÈÈ. 1972

'Ìàòåðèàëû:
'1 - Ñòàëü 35
'2 - Ìåäíûå ñïëàâû
'3 - Àëþìèíèåâûå ñïëàâû
'4 - 40Õ
'5 - 12ÕÍ3À
'6 - 20ÕÃ, 40ÕÃÀ, 38ÕÌÞÀ
'7 - 20Õ13
'8 - ÂÒ1-0
'9 - 36ÍÕÒÞ, 12Õ21Í5Ò
'10 - 12Õ18Í9Ò, ÂÍÌ-3, ÂÍÌ5-3

Private Const IMYA_KNIGI As String = "slesar 1972.xlsb"

Dim NumRow As Integer
Dim NumCol As Integer
Dim Smax As Double

Public Const DmaxPritiCilindr_Ukrup As Integer = 100
Public Const LmaxPritiCilindr_Ukrup As Integer = 50

Public Enum EnumVidShva
    VStyk = 1
    VTavr = 2
    VNakhlest = 3
End Enum

Dim varPritiCilindr

Const PerehodPriSverlNarez As Integer = 2

Private Normativ_SvarkaSoSkosom
Private Normativ_SvarkaBezSkosa

Private Const MaksTolshchinaSvarki As Integer = 15

Function SvarkaArgonoDugovaya(ByVal Material As EnumMaterialy, Tolshchina As Double, Dlina As Double, _
    Optional SkosKromok As Boolean = True, Optional VidShva As EnumVidShva = VTavr, _
    Optional NeudobnyyeUsloviya As Boolean = False, Optional KrivolineynyyShov As Boolean = False)
    
    If ProverkaParametrov_Svarka(Material, VidShva, SkosKromok, Tolshchina, Dlina) = False Then Exit Function
    
    If Material = STAL_LEGIROVANNAYA Or Material = STAL_UGLERODISTAYA Then Material = STAL_NERZHAVEYUSHCHAYA
    
    Dim Normativ
    Normativ = ZagruzitNormativ(IIf(SkosKromok, "svarka_so_skosom", "svarka_bez_skosa"))
    
    Dim i As Integer, a As Double, b As Double, tSvarka As Double
    For i = LBound(Normativ) To UBound(Normativ)
        If Normativ(i, 1) = Material And Normativ(i, 2) >= Tolshchina And Normativ(i, 3) = VidShva Then
            a = Normativ(i, 4)
            b = Normativ(i, 5)
            tSvarka = a * Dlina + b
            Exit For
        End If
    Next

    SvarkaArgonoDugovaya = PoprKoefSvarka(NeudobnyyeUsloviya, KrivolineynyyShov) * tSvarka

End Function

Private Function ProverkaParametrov_Svarka(Material As EnumMaterialy, VidShva As EnumVidShva, SkosKromok As Boolean, Tolshchina As Double, Dlina As Double) As Boolean
    
    ProverkaParametrov_Svarka = False
    
    If MaterialKorrektnyy_Svarka(Material) = False Then Exit Function
    
    If PolozhitelnyyeChisla(Tolshchina, Dlina) = False Then Exit Function
    
    If Tolshchina > MaksTolshchinaSvarki Then Exit Function
    
    Select Case VidShva
    
        Case VStyk, VTavr
        
        Case VNakhlest
            If SkosKromok = True Then Exit Function
            
        Case Else
            Exit Function
            
    End Select
    
    ProverkaParametrov_Svarka = True
    
End Function

Private Function PoprKoefSvarka(NeudobnyyeUsloviya As Boolean, KrivolineynyyShov As Boolean) As Double

    PoprKoefSvarka = IIf(KrivolineynyyShov, 1.2, 1) * IIf(NeudobnyyeUsloviya, 1.3, 1)
    
End Function

Private Function MaterialKorrektnyy_Svarka(Material As EnumMaterialy) As Boolean
    Select Case Material
        Case ALUMINIYEVYYE_SPLAVY, MEDNYYE_SPLAVY, STAL_LEGIROVANNAYA, STAL_NERZHAVEYUSHCHAYA, STAL_UGLERODISTAYA, TITANOVYYE_SPLAVY
            MaterialKorrektnyy_Svarka = True
        Case Else
            MaterialKorrektnyy_Svarka = False
    End Select
End Function

Function PreobrazovatMaterialVGruppuMateriala(Material As Integer)
    Dim Gruppa As Integer
    If Material > 2 Then
        Gruppa = 3
    Else
        Gruppa = Material
    End If
    PreobrazovatMaterialVGruppuMateriala = Gruppa
End Function


Function PritirDovodCilindrPoverh(Material As Integer, d As Double, L As Double, Ra As Double, _
    Otkloneniye As Double, Optional VUpor As Boolean = False, Optional Konus As Boolean = False, _
    Optional AzotNitr As Boolean = False, Optional Sfera As Boolean = False, Optional Tonkosten As Boolean = False) As Double
    
    Dim MinOtklonenie As Double:    MinOtklonenie = 0.001
    Dim MinRa As Double:            MinRa = 0.025
    
    If Material < 1 Or Material > 10 Or _
        d > DmaxPritiCilindr_Ukrup Or d <= 0 Or _
        L <= 0 Or L > LmaxPritiCilindr_Ukrup Or _
        Otkloneniye < MinOtklonenie Or Ra < MinRa Then Exit Function
       
    If IsEmpty(varPritiCilindr) Then
        varPritiCilindr = wsUkrup_PritirCilindr.ListObjects("tbPritirCilindr").DataBodyRange.Value
    End If
    
    Dim i As Integer
    Dim Thsh As Double
    For i = LBound(varPritiCilindr) To UBound(varPritiCilindr)
        If varPritiCilindr(i, 1) >= d Then
            If varPritiCilindr(i, 2) >= L Then
                If varPritiCilindr(i, 3) <= Ra Then
                    If varPritiCilindr(i, 4) <= Otkloneniye Then
                        Thsh = varPritiCilindr(i, 5)
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    
    Dim K(1 To 6) As Double
    K(1) = IIf(VUpor, 1.2, 1)
    K(2) = IIf(Konus, 1.5, 1)
    K(3) = IIf(AzotNitr, 1.2, 1)
    K(4) = IIf(Sfera, 1.4, 1)
    K(5) = IIf(Tonkosten, 1.3, 1)
    
    Select Case Material
        Case 1:    K(6) = 1
        Case 2:    K(6) = 0.8
        Case 3:    K(6) = 1
        Case Else: K(6) = 1.4
    End Select
        
    PritirDovodCilindrPoverh = Thsh * WorksheetFunction.Product(K)

End Function

Function PritirDovodCilindrPoverh_SvobodnyyRazmer(Material As Integer, d As Double, L As Double, Ra As Double, _
    Optional VUpor As Boolean = False, Optional Konus As Boolean = False, _
    Optional AzotNitr As Boolean = False, Optional Sfera As Boolean = False, Optional Tonkosten As Boolean = False) As Double
    
    Dim Otkloneniye As Double:      Otkloneniye = 0.01
    Dim KoefSvobodnyy As Double:    KoefSvobodnyy = 0.5
    
    PritirDovodCilindrPoverh_SvobodnyyRazmer = PritirDovodCilindrPoverh(Material, d, L, Ra, Otkloneniye, VUpor, Konus, AzotNitr, Sfera, Tonkosten) * KoefSvobodnyy
    
End Function

Function Svar_PrihvatArgon_m35(Mat As Integer, S As Double, VidSoed As Integer, KolTochek As Double) As Double
    
    Dim tFirst As Double, tNext As Double
    
    Select Case Mat
        Case 1
            'ÌÀ1; ÌÀ8
            tFirst = 0.9: tNext = 0.15
        Case 2
            'ÀÄ; ÀÄ-1; ÀÌÖ; ÀÌÖÀÌ; ÀÌÃ; ÀÌÃ6; Ä20; Ä16Ò; Ä16ÀÒ; Ä16ÀÂ; ÀÌÃ6Ò
            If S <= 3 Then
                tFirst = IIf(VidSoed = 1, 1, 1.1)
                tNext = IIf(VidSoed = 1, 0.2, 0.25)
            Else
                tFirst = IIf(VidSoed = 1, 1.1, 1.2)
                tNext = IIf(VidSoed = 1, 0.3, 0.35)
            End If
        Case 3
            '0Õ18Í9; 1Õ18Í9Ò; Õ17Í2 (ÝÈ268); Õ23Í18 (ÝÈ417); ÝÈ415; ÕÍ78Ò; ÝÈ703; ÝÈ654; ÝÈ652; ÝÈ602; ÂÆ98; 2Õ13 (Æ2); ÝÈ402
            If S <= 3 Then
                tFirst = IIf(VidSoed = 1, 1.1, 1.2)
                tNext = IIf(VidSoed = 1, 0.25, 0.3)
            Else
                tFirst = IIf(VidSoed = 1, 1.2, 1.3)
                tNext = IIf(VidSoed = 1, 0.35, 0.4)
            End If
        Case Else
            '25ÕÃÑÀ; 30ÕÃÑÀ; ÝÈ659; ÝÈ415; ÝÈ100
            If S <= 3 Then
                tFirst = IIf(VidSoed = 1, 1.1, 1.2)
                tNext = 0.3
            Else
                tFirst = IIf(VidSoed = 1, 1.3, 1.4)
                tNext = 0.4
            End If
    End Select
    
    Svar_PrihvatArgon_m35 = tFirst + (tNext * (KolTochek - 1))

End Function


Function Svar_ZachistNapil_m44(VidShva As Integer, L As Double, Mat As Integer) As Double

    If VidShva < 1 Or VidShva > 3 Then Exit Function
    
    Dim a As Double, X As Double, K As Double

    Select Case VidShva
        Case 1: a = 5.55713712205289E-02:   X = 0.828532602756075
        Case 2: a = 0.119702864666647:      X = 0.740975150607911
        Case 3: a = 0.194890961000119:      X = 0.690645379638459
    End Select
    
    If Mat = 1 Then K = 1 Else K = IIf(Mat = 2, 0.8, 0.6)
    Svar_ZachistNapil_m44 = (a * L ^ X) * K
    
End Function

Function Svar_ZachistZapodlico_m44(VidShva As Integer, L As Double, Mat As Integer) As Double

    If VidShva < 1 Or VidShva > 3 Then Exit Function
    
    Dim a As Double, X As Double, K As Double
    
    Select Case VidShva
        Case 1: a = 0.010358974: X = 0.441025641
        Case 2: a = 0.012307692: X = 0.692307692
        Case 3: a = 0.013538462: X = 0.961538462
    End Select

    If Mat = 1 Then K = 1 Else K = IIf(Mat = 2, 0.8, 0.6)
    Svar_ZachistZapodlico_m44 = (a * L + X) * K
    
End Function



Function Svar_ZavarOknoPaz_m38(b As Double, L As Double, S As Double) As Double

    If b > 50 Or L > 100 Or S > 10 Then Exit Function
    
    Static rB As Range, rL As Range, rs As Range, rTime As Range
    With ThisWorkbook.Worksheets("Óêðóï_38_3")
        Set rB = .Range("C6:C12")
        Set rL = .Range("D6:T12")
        Set rs = .Range("C14:C17")
        Set rTime = .Range("D14:T17")
    End With
    
    For NumRow = 1 To rB.Rows.Count
        If rB(NumRow, 1) >= b Then Exit For
    Next

    For NumCol = 1 To rL.Columns.Count
        If rL(NumRow, NumCol) > 0 And rL(NumRow, NumCol) >= L Then Exit For
    Next
    
    For NumRow = 1 To rs.Rows.Count
        If rs(NumRow, 1) >= S Then Exit For
    Next
        
    Svar_ZavarOknoPaz_m38 = rTime(NumRow, NumCol)
    
End Function

Function Svar_ZavarOtv_m38(d As Double, S As Double) As Double

    If S > 10 Or d > 50 Then Exit Function
    
    Select Case d
        Case Is <= 20
            Select Case S
                Case Is <= 3:   Svar_ZavarOtv_m38 = 1.1
                Case Is <= 5:   Svar_ZavarOtv_m38 = 1.3
                Case Is <= 8:   Svar_ZavarOtv_m38 = 1.6
                Case Else:      Svar_ZavarOtv_m38 = 1.9
            End Select
        Case Is <= 25
            Select Case S
                Case Is <= 3:   Svar_ZavarOtv_m38 = 1.4
                Case Is <= 5:   Svar_ZavarOtv_m38 = 1.7
                Case Is <= 8:   Svar_ZavarOtv_m38 = 2.2
                Case Else:      Svar_ZavarOtv_m38 = 2.7
            End Select
        Case Is <= 30
            Select Case S
                Case Is <= 3:   Svar_ZavarOtv_m38 = 1.8
                Case Is <= 5:   Svar_ZavarOtv_m38 = 2.2
                Case Is <= 8:   Svar_ZavarOtv_m38 = 2.9
                Case Else:      Svar_ZavarOtv_m38 = 3.9
            End Select
        Case Is <= 35
            Select Case S
                Case Is <= 3:   Svar_ZavarOtv_m38 = 2.2
                Case Is <= 5:   Svar_ZavarOtv_m38 = 2.8
                Case Is <= 8:   Svar_ZavarOtv_m38 = 3.7
                Case Else:      Svar_ZavarOtv_m38 = 4.8
            End Select
        Case Is <= 40
            Select Case S
                Case Is <= 3:   Svar_ZavarOtv_m38 = 2.6
                Case Is <= 5:   Svar_ZavarOtv_m38 = 3.4
                Case Is <= 8:   Svar_ZavarOtv_m38 = 4.8
                Case Else:      Svar_ZavarOtv_m38 = 6
            End Select
        Case Is <= 45
            Select Case S
                Case Is <= 3:   Svar_ZavarOtv_m38 = 3.2
                Case Is <= 5:   Svar_ZavarOtv_m38 = 4.4
                Case Is <= 8:   Svar_ZavarOtv_m38 = 6
                Case Else:      Svar_ZavarOtv_m38 = 7.2
            End Select
        Case Else
            Select Case S
                Case Is <= 3:   Svar_ZavarOtv_m38 = 3.8
                Case Is <= 5:   Svar_ZavarOtv_m38 = 5.2
                Case Is <= 8:   Svar_ZavarOtv_m38 = 7.2
                Case Else:      Svar_ZavarOtv_m38 = 8.8
            End Select
    End Select
    
End Function

