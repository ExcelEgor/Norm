Attribute VB_Name = "mLazernayaGravirovka"
Option Explicit

Public Enum TipyLazernoyGravirovki
    Nadpis = 1
    LogotipAvrora = 2
    Zemlya = 3
    Napryazheniye = 4
    KolVoSimvolov = 5
    Liniya = 6
End Enum

Function RaschotLazernoyGravirovki(Material As EnumMaterialy, PodPokrasku As Boolean, _
    Tip As TipyLazernoyGravirovki, text As String, ShriftShirinaVysota As Double, ByVal KolVo As Double) As Variant
    
    Select Case Material
        Case ALUMINIYEVYYE_SPLAVY, STAL_UGLERODISTAYA, STAL_NERZHAVEYUSHCHAYA, TITANOVYYE_SPLAVY
        Case Else
            Exit Function
    End Select
    
    If ShriftShirinaVysota <= 0 Then Exit Function
    
    If KolVo = 0 Then KolVo = 1
    
    Dim tLaser As Double, tZatirka As Double, tKontrol As Double

    If Tip = Nadpis Then
        tLaser = LazerGrav_Nadpis(text, Material, ShriftShirinaVysota, PodPokrasku)
    
    ElseIf Tip = LogotipAvrora Then
        tLaser = LazerGrav_LogotipAvrora(Material, ShriftShirinaVysota, PodPokrasku) * KolVo
    
    ElseIf Tip = Zemlya Then
        tLaser = LazerGrav_Zemlya(Material, ShriftShirinaVysota, PodPokrasku) * KolVo
    
    ElseIf Tip = Napryazheniye Then
        tLaser = LazerGrav_Napryazheniye(Material, ShriftShirinaVysota, PodPokrasku) * KolVo
    
    ElseIf Tip = KolVoSimvolov Then
        tLaser = LazerGrav_KolVoSimvolol(Material, ShriftShirinaVysota, KolVo, PodPokrasku)
    
    ElseIf Tip = Liniya Then
        tLaser = LazerGrav_Liniya(Material, ShriftShirinaVysota, PodPokrasku) * KolVo
    Else
        Exit Function
    End If
    
    If Tip = Liniya Then
        tZatirka = ZapolGravEmal_Riska_45(ShriftShirinaVysota, 2, 2, KolVo)
    Else
        tZatirka = ZapolGravEmal_Znak_45(ShriftShirinaVysota, KolVo, 2, 2)
    End If
    tKontrol = VisualnyyKontrol_Zatirka(ShriftShirinaVysota, KolVo, 0, 3)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tLaser
    Normy(2) = 1.3 * tZatirka
    Normy(3) = 1.3 * tKontrol
    
    RaschotLazernoyGravirovki = Normy
    
End Function

Function LazerGrav_Liniya(Material As EnumMaterialy, Dlina As Double, Optional PodPokrasku As Boolean = False) As Double

    If ProverkaOshovnykhParametrov(Material, Dlina) = False Then Exit Function
    
    Dim t100 As Double

    If PodPokrasku Then

    Else
        If Material = ALUMINIYEVYYE_SPLAVY Then
            t100 = 1.69
        Else
            t100 = 5.57
        End If
    End If
    
    LazerGrav_Liniya = t100 * Dlina / 100
    
End Function

Private Function ProverkaOshovnykhParametrov(Material As EnumMaterialy, razmer As Double) As Boolean

    If Material <> ALUMINIYEVYYE_SPLAVY And Material <> STAL_NERZHAVEYUSHCHAYA And Material <> STAL_UGLERODISTAYA And Material <> TITANOVYYE_SPLAVY Or razmer <= 0 Then
        ProverkaOshovnykhParametrov = False
    Else
        ProverkaOshovnykhParametrov = True
    End If
    
End Function


Function LazerGrav_Zemlya(Material As EnumMaterialy, VysotaSimvola As Double, Optional PodPokrasku As Boolean = False) As Double

    If ProverkaOshovnykhParametrov(Material, VysotaSimvola) = False Then Exit Function
    
    Dim a As Double
    Dim b As Double
    
    If PodPokrasku Then
        Select Case Material
            Case ALUMINIYEVYYE_SPLAVY
                a = 0.01471
                b = 1.70346
            Case STAL_UGLERODISTAYA
                a = 0.03814
                b = 1.75737
            Case STAL_NERZHAVEYUSHCHAYA
                a = 0.04694
                b = 1.76643
            Case TITANOVYYE_SPLAVY
                a = 0.04871
                b = 1.76785
        End Select
    Else
        If Material = ALUMINIYEVYYE_SPLAVY Then
            a = 0.01075
            b = 1.61555
        Else
            a = 0.02566
            b = 1.68778
        End If
    End If
    
    LazerGrav_Zemlya = a * VysotaSimvola ^ b

End Function

Function LazerGrav_Nadpis(ByVal text As String, Material As EnumMaterialy, RazmerShrifta As Double, Optional PodPokrasku As Boolean = False) As Double
    
    text = Replace(text, " ", "")
    If ProverkaOshovnykhParametrov(Material, RazmerShrifta) = False Or text = Empty Then Exit Function
    
    Dim i As Integer
    Dim tNaSimvol As Double

    For i = 1 To Len(text)
        tNaSimvol = 0
        If Not Mid(text, i, 1) = " " Then
            Select Case Asc(Mid(text, i, 1))
                
                'Символы
                Case 1 To 47, 58 To 64, 91 To 96, 123 To 191
                    tNaSimvol = LazerGrav_Simvoli(Material, RazmerShrifta, PodPokrasku)
                
                    'Цифры
                Case 48 To 57
                    tNaSimvol = LazerGrav_Tsifry(Material, RazmerShrifta, PodPokrasku)
                
                    'Большие буквы
                Case 65 To 90, 192 To 223
                    tNaSimvol = BolshieBukvy(Material, RazmerShrifta, PodPokrasku)
                
                    'Маленькие буквы
                Case 97 To 122, 224 To 255
                    tNaSimvol = MalenkieBukvy(Material, RazmerShrifta, PodPokrasku)

            End Select
        End If
        
        LazerGrav_Nadpis = LazerGrav_Nadpis + tNaSimvol
    Next

End Function
Function LazerGrav_KolVoSimvolol(Material As EnumMaterialy, RazmerShrifta As Double, KolVoSimvolol As Double, Optional PodPokrasku As Boolean = False) As Double

    If ProverkaOshovnykhParametrov(Material, RazmerShrifta) = False Or Not KolVoSimvolol > 0 Then Exit Function
    
    LazerGrav_KolVoSimvolol = BolshieBukvy(Material, RazmerShrifta, PodPokrasku) * KolVoSimvolol
    
End Function

Function LazerGrav_Napryazheniye(Material As EnumMaterialy, VysotaSimvola As Double, Optional PodPokrasku As Boolean = False) As Double

    If ProverkaOshovnykhParametrov(Material, VysotaSimvola) = False Then Exit Function

    Dim a As Double
    Dim b As Double
    
    If PodPokrasku Then
        Select Case Material
            Case ALUMINIYEVYYE_SPLAVY
                a = 0.01063
                b = 1.95238
            Case STAL_UGLERODISTAYA
                a = 0.03126
                b = 1.95488
            Case STAL_NERZHAVEYUSHCHAYA
                a = 0.03926
                b = 1.95604
            Case TITANOVYYE_SPLAVY
                a = 0.04088
                b = 1.95623
        End Select
    Else
        If Material = ALUMINIYEVYYE_SPLAVY Then
            a = 0.00642
            b = 1.93745
        Else
            a = 0.01782
            b = 1.95755
        End If
    End If
    
    LazerGrav_Napryazheniye = a * VysotaSimvola ^ b

End Function


Function LazerGrav_LogotipAvrora(Material As EnumMaterialy, Shirina As Double, Optional PodPokrasku As Boolean = False)
    
    If ProverkaOshovnykhParametrov(Material, Shirina) = False Then Exit Function
    
    Dim a As Double
    Dim b As Double
    
    If PodPokrasku Then
       
        Select Case Material
            Case ALUMINIYEVYYE_SPLAVY
                a = 0.01768
                b = 1.92989
            Case STAL_UGLERODISTAYA
                a = 0.05484
                b = 1.92088
            Case STAL_NERZHAVEYUSHCHAYA
                a = 0.06878
                b = 1.92256
            Case TITANOVYYE_SPLAVY
                a = 0.07159
                b = 1.92274
        End Select
       
    Else
    
        If Material = ALUMINIYEVYYE_SPLAVY Then
            a = 0.01078
            b = 1.9115
        Else
            a = 0.02717
            b = 1.93873
        End If
        
    End If
    
    LazerGrav_LogotipAvrora = a * Shirina ^ b
    
End Function

Private Function BolshieBukvy(Material As EnumMaterialy, RazmerShrifta As Double, Optional PodPokrasku As Boolean = False)
    
    Dim a As Double
    Dim b As Double
    
    If PodPokrasku Then
        Select Case Material
            Case ALUMINIYEVYYE_SPLAVY
                a = 0.04019
                b = 1.91658
            Case STAL_UGLERODISTAYA
                a = 0.12002
                b = 1.91634
            Case STAL_NERZHAVEYUSHCHAYA
                a = 0.15044
                b = 1.91752
            Case TITANOVYYE_SPLAVY
                a = 0.1565
                b = 1.91774
        End Select
    Else
        If Material = ALUMINIYEVYYE_SPLAVY Then
            a = 0.02498
            b = 1.90188
        Else
            a = 0.07051
            b = 1.91267
        End If
    End If
    
    BolshieBukvy = a * RazmerShrifta ^ b

End Function


Private Function MalenkieBukvy(Material As EnumMaterialy, RazmerShrifta As Double, Optional PodPokrasku As Boolean = False)
    
    Dim a As Double
    Dim b As Double
    
    If PodPokrasku Then
        Select Case Material
            Case ALUMINIYEVYYE_SPLAVY
                a = 0.02991
                b = 1.90886
            Case STAL_UGLERODISTAYA
                a = 0.08933
                b = 1.90862
            Case STAL_NERZHAVEYUSHCHAYA
                a = 0.11197
                b = 1.90979
            Case TITANOVYYE_SPLAVY
                a = 0.11648
                b = 1.91002
        End Select

    Else
        If Material = ALUMINIYEVYYE_SPLAVY Then
            a = 0.01859
            b = 1.89416
        Else
            a = 0.05247
            b = 1.90494
        End If

    End If
    
    MalenkieBukvy = a * RazmerShrifta ^ b
    
End Function

Private Function LazerGrav_Tsifry(Material As EnumMaterialy, RazmerShrifta As Double, Optional PodPokrasku As Boolean = False) As Double
    
    Dim a As Double
    Dim b As Double
    
    If PodPokrasku Then
        Select Case Material
            Case ALUMINIYEVYYE_SPLAVY
                a = 0.03476
                b = 1.69754
            Case STAL_UGLERODISTAYA
                a = 0.08449
                b = 1.89934
            Case STAL_NERZHAVEYUSHCHAYA
                a = 0.10611
                b = 1.90055
            Case TITANOVYYE_SPLAVY
                a = 0.11038
                b = 1.90082
        End Select
    Else
        If Material = ALUMINIYEVYYE_SPLAVY Then
            a = 0.01773
            b = 1.87061
        Else
            a = 0.04779
            b = 1.86798
        End If
    End If
    
    LazerGrav_Tsifry = a * RazmerShrifta ^ b

End Function

Private Function LazerGrav_Simvoli(Material As EnumMaterialy, RazmerShrifta As Double, Optional PodPokrasku As Boolean = False) As Double
    
    Dim a As Double
    Dim b As Double
    
    If PodPokrasku Then
        Select Case Material
            Case ALUMINIYEVYYE_SPLAVY
                a = 0.01437
                b = 1.90341
            Case STAL_UGLERODISTAYA
                a = 0.0429
                b = 1.90317
            Case STAL_NERZHAVEYUSHCHAYA
                a = 0.05377
                b = 1.90434
            Case TITANOVYYE_SPLAVY
                a = 0.05594
                b = 1.90457
        End Select
    Else
        If Material = ALUMINIYEVYYE_SPLAVY Then
            a = 0.00893
            b = 1.88871
        Else
            a = 0.0252
            b = 1.89949
        End If
    End If
    
    LazerGrav_Simvoli = a * RazmerShrifta ^ b

End Function
