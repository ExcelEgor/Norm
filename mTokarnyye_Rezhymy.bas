Attribute VB_Name = "mTokarnyye_Rezhymy"
Option Explicit

Const SmmOb_CHERN As Double = 0.25
Const T_OBS As Double = 1.1
Const Stoykost As Integer = 180
Const K_MARKA_TVERDOGO_SPLAVA As Integer = 0.87

Public Const RAZMER_MIN_POP As Integer = 50

Const GLUBINA_REZANIYA_MIN As Integer = 1


Const GLUBINA_REZANIYA_CHISTOVAYA As Double = 0.5
Const KOEFFITSIENT_SERIYNOST As Double = 0.7

Private Type PokazeteliStepeni
    Cv As Double
    Xv As Double
    Yv As Double
    m As Double
End Type

'Материалы
'1 - Медные сплавы
'2 - Алюминиевые сплавы
'3 - Сталь 35
'4 - 25Х13Н2
'5 - Титан ВТ1-0
'6 - 12Х18Н9Т
'7 - ВНМ 3-2


Private Sub Test_Prodolnoye()
    Dim Material As EnumMaterialy, diameter_zagotovki As Double, diameter_detali As Double, GLUBINA_REZANIYA As Double
    Dim dlina_tocheniya As Double, Ra As Double, Kvalitet As Integer, Korka As Boolean, Udar As Boolean, podrezka_tortsa As Boolean
    
    Material = 6
    diameter_zagotovki = 100
    diameter_detali = 10
    GLUBINA_REZANIYA = 2
    dlina_tocheniya = 300
    Ra = 3.2
    Kvalitet = 9
    Korka = True
    Udar = True
    podrezka_tortsa = True
    
    Debug.Print ProdolnoyeTocheniye_PoRezhimam(Material, diameter_zagotovki, diameter_detali, GLUBINA_REZANIYA, dlina_tocheniya, Ra, Kvalitet, Korka, Udar, podrezka_tortsa)
    
End Sub

Private Function v_rastachivaniye(Material As EnumMaterialy, d As Double, ByVal GLUBINA_REZANIYA As Double, SmmOb As Double, _
    Udar As Boolean, Korka As Boolean) As Double
        
    If GLUBINA_REZANIYA < GLUBINA_REZANIYA_MIN Then GLUBINA_REZANIYA = GLUBINA_REZANIYA_MIN
        
    Dim k_v As Double
    Select Case d
        Case Is <= 75:  k_v = 0.7
        Case Is <= 150: k_v = 0.8
        Case Is <= 250: k_v = 0.9
        Case Else:      k_v = 1
    End Select
    
    v_rastachivaniye = k_v * RaschotSkorostiRezaniya(CInt(Material), GLUBINA_REZANIYA, SmmOb, Udar, Korka)

End Function

Private Function v_poperechnoye_tocheniye(Material As EnumMaterialy, d As Double, dlina_tocheniya As Double, ByVal GLUBINA_REZANIYA As Double, SmmOb As Double, _
    Udar As Boolean, Korka As Boolean) As Double
    
    If GLUBINA_REZANIYA < GLUBINA_REZANIYA_MIN Then GLUBINA_REZANIYA = GLUBINA_REZANIYA_MIN
    
    Dim k_v As Double
    Dim d2 As Double
    d2 = 2 * dlina_tocheniya + d
    Select Case d2 / d
        Case Is <= 0.4: k_v = 1.25
        Case Is <= 0.7: k_v = 1.2
        Case Else:      k_v = 1.05
    End Select
    
    v_poperechnoye_tocheniye = k_v * RaschotSkorostiRezaniya(CInt(Material), GLUBINA_REZANIYA, SmmOb, Udar, Korka)

End Function

Private Function RaschotSkorostiRezaniya(Material As EnumMaterialy, ByVal GLUBINA_REZANIYA As Double, SmmOb As Double, _
    Udar As Boolean, Korka As Boolean) As Double
        
    If GLUBINA_REZANIYA < GLUBINA_REZANIYA_MIN Then GLUBINA_REZANIYA = GLUBINA_REZANIYA_MIN
    
    Dim Pokazateli As PokazeteliStepeni
    Pokazateli = PoluchitiPokazateliStepeni(Material, SmmOb)

    With Pokazateli
        RaschotSkorostiRezaniya = .Cv / (Stoykost ^ .m * GLUBINA_REZANIYA ^ .Xv * SmmOb ^ .Yv)
    End With
    
    Dim kMaterial As Double
    If Material = STAL_LEGIROVANNAYA Or Material = TITANOVYYE_SPLAVY Then
        kMaterial = 0.8
    Else
        kMaterial = 1
    End If
    
    Dim kV_1 As Double
    If Korka Then
        If Udar Then
            kV_1 = 0.65
        Else
            kV_1 = 0.85
        End If
    Else
        If Udar Then
            kV_1 = 0.7
        Else
            kV_1 = 1
        End If
    End If

    RaschotSkorostiRezaniya = K_MARKA_TVERDOGO_SPLAVA * kV_1 * kMaterial * RaschotSkorostiRezaniya

End Function

Private Function PoluchitiPokazateliStepeni(Material As EnumMaterialy, SmmOb As Double) As PokazeteliStepeni
    
    With PoluchitiPokazateliStepeni
        If Material = MEDNYYE_SPLAVY Then
            If SmmOb <= 0.4 Then
                .Cv = 550
                .Xv = 0.13
                .Yv = 0.2
            Else
                .Cv = 535
                .Xv = 0.2
                .Yv = 0.4
            End If
            .m = 0.2
            
        ElseIf Material = ALUMINIYEVYYE_SPLAVY Then
            .Cv = 582
            .Xv = 0.12
            .Yv = 0.5
            .m = 0.2
            
        ElseIf Material = EnumMaterialy.STAL_LEGIROVANNAYA Or Material = STAL_UGLERODISTAYA Then
            Select Case SmmOb
                Case Is <= 0.3
                    .Cv = 252
                    .Yv = 0.2
                Case Is <= 0.75
                    .Cv = 209
                    .Yv = 0.35
                Case Else
                    .Cv = 203
                    .Yv = 0.45
            End Select
            .Xv = 0.15
            .m = 0.2
            
        ElseIf Material = STAL_NERZHAVEYUSHCHAYA Or Material = EnumMaterialy.TITANOVYYE_SPLAVY Then
            .Cv = 104
            .Xv = 0.2
            .Yv = 0.45
            .m = 0.15
            
        Else
            Exit Function
            
        End If
    End With
    
End Function

Private Function t_mash_prodolnoye_tocheniye(Material As EnumMaterialy, GLUBINA_REZANIYA As Double, SmmOb As Double, d As Double, L As Double, _
    Udar As Boolean, Korka As Boolean) As Double
    
    If Not (GLUBINA_REZANIYA > 0 And SmmOb > 0 And d > 0 And L > 0) Then Exit Function
    
    Dim V As Double, N As Double
    
    V = RaschotSkorostiRezaniya(CInt(Material), GLUBINA_REZANIYA, SmmOb, Udar, Korka)
    N = RaschotOborotov(V, d)
    t_mash_prodolnoye_tocheniye = (L + L_NEGOBEGA_PEREBEGA_TOCHENIYE) / (N * SmmOb)
    
End Function

Private Function t_mash_rastachivaniye(Material As EnumMaterialy, GLUBINA_REZANIYA As Double, SmmOb As Double, d As Double, dlina_rastachivaniya As Double, _
    Udar As Boolean, Korka As Boolean) As Double
    
    Dim V As Double, N As Double
    
    V = v_rastachivaniye(Material, d, GLUBINA_REZANIYA, SmmOb, Udar, Korka)
    N = RaschotOborotov(V, d)
    t_mash_rastachivaniye = (dlina_rastachivaniya + L_NEGOBEGA_PEREBEGA_TOCHENIYE) / (N * SmmOb)
    
End Function

Private Function t_mash_poperechnoye_tocheniye(Material As EnumMaterialy, GLUBINA_REZANIYA As Double, SmmOb As Double, Diametr As Double, dlina_tocheniya As Double, _
    Udar As Boolean, Korka As Boolean) As Double
    
    Dim V As Double, N As Double
    
    V = v_poperechnoye_tocheniye(Material, Diametr, dlina_tocheniya, GLUBINA_REZANIYA, SmmOb, Udar, Korka)
    N = RaschotOborotov(V, Diametr)
    
    t_mash_poperechnoye_tocheniye = (dlina_tocheniya + L_NEGOBEGA_PEREBEGA_TOCHENIYE) / (N * SmmOb)
    
End Function

Private Function vybor_podachi_v_zavisimosti_ot_ra(Ra As Double) As Double

    Dim SmmOb As Double
    Select Case Ra
        Case Is <= 1.25
            SmmOb = 0.05
        Case Is <= 2.5
            SmmOb = 0.1
        Case Is <= 3.2
            SmmOb = 0.2
        Case Else:
            SmmOb = SmmOb_CHERN
    End Select
    
    vybor_podachi_v_zavisimosti_ot_ra = SmmOb
    
End Function

Private Function kolvo_prokhodov_v_zavisimosti_ot_kvaliteta(Kvalitet As Integer) As Integer
    kolvo_prokhodov_v_zavisimosti_ot_kvaliteta = IIf(Kvalitet > 11, 1, 2)
End Function

Private Function prodolnoye_tocheniye_chistovoye_po_rezhimam(Material As EnumMaterialy, d_detali As Double, dlina_tocheniya As Double, Ra As Double, Kvalitet As Integer, _
     Udar As Boolean) As Double
    
    Dim SmmOb As Double
    SmmOb = vybor_podachi_v_zavisimosti_ot_ra(Ra)
    
    Dim kolvo_prokhodov As Integer
    kolvo_prokhodov = kolvo_prokhodov_v_zavisimosti_ot_kvaliteta(Kvalitet)

    Dim t_mash As Double
    Dim t_vsp As Double

    t_mash = t_mash_prodolnoye_tocheniye(Material, GLUBINA_REZANIYA_CHISTOVAYA, SmmOb, d_detali, dlina_tocheniya, Udar, False)
    t_vsp = Tvsp_Prodolnoye_Tocheniye_Chistovoye(d_detali, dlina_tocheniya, Kvalitet)
    
    prodolnoye_tocheniye_chistovoye_po_rezhimam = T_OBS * (kolvo_prokhodov * t_mash + t_vsp)
    
End Function

Private Function rastachivaniye_chistovoye_po_rezhimam(Material As EnumMaterialy, d_detali As Double, dlina_rastachivaniya As Double, Ra As Double, Kvalitet As Integer, _
     Udar As Boolean) As Double
    
    Dim SmmOb As Double
    SmmOb = vybor_podachi_v_zavisimosti_ot_ra(Ra)
    
    Dim kolvo_prokhodov As Integer
    kolvo_prokhodov = kolvo_prokhodov_v_zavisimosti_ot_kvaliteta(Kvalitet)

    Dim t_mash As Double
    Dim t_vsp As Double

    t_mash = t_mash_rastachivaniye(Material, GLUBINA_REZANIYA_CHISTOVAYA, SmmOb, d_detali, dlina_rastachivaniya, Udar, False)
    t_vsp = Tvsp_Prodolnoye_Tocheniye_Chistovoye(d_detali, dlina_rastachivaniya, Kvalitet)
    
    rastachivaniye_chistovoye_po_rezhimam = T_OBS * (kolvo_prokhodov * t_mash + t_vsp)
    
End Function

Private Function poperechnoye_tocheniye_chernovoye_po_rezhimam(Material As EnumMaterialy, Diametr As Double, dlina_tocheniya As Double, GLUBINA_REZANIYA As Double, _
     Udar As Boolean, Korka As Boolean) As Double

    Dim t_mash As Double
    Dim t_vsp As Double

    t_mash = t_mash_poperechnoye_tocheniye(Material, GLUBINA_REZANIYA, SmmOb_CHERN, Diametr, dlina_tocheniya, Udar, Korka)
    t_vsp = Tvsp_PoperechnoyeTocheniye_Chernovoye(dlina_tocheniya)
    poperechnoye_tocheniye_chernovoye_po_rezhimam = T_OBS * (t_mash + t_vsp)
    
End Function

Private Function poperechnoye_tocheniye_chistovoye_po_rezhimam(Material As EnumMaterialy, Diametr As Double, dlina_tocheniya As Double, Ra As Double, Kvalitet As Integer, _
     izmeryaemyy_razmer As Double, Udar As Boolean) As Double
    
    Dim SmmOb As Double
    SmmOb = vybor_podachi_v_zavisimosti_ot_ra(Ra)
    
    Dim kolvo_prokhodov As Integer
    kolvo_prokhodov = kolvo_prokhodov_v_zavisimosti_ot_kvaliteta(Kvalitet)

    Dim t_mash As Double
    Dim t_vsp As Double
    
    t_mash = t_mash_poperechnoye_tocheniye(Material, GLUBINA_REZANIYA_CHISTOVAYA, SmmOb, Diametr, dlina_tocheniya, Udar, False)
    t_vsp = Tvsp_PoperechnoyeTocheniye_Chistovoye(izmeryaemyy_razmer, dlina_tocheniya, Kvalitet)
    
    poperechnoye_tocheniye_chistovoye_po_rezhimam = T_OBS * (kolvo_prokhodov * t_mash + t_vsp)
    
End Function

Function PoperechnoyeTocheniyeKompleks_PoRezhimam(Material As EnumMaterialy, D1 As Double, d2 As Double, Pripusk As Double, GLUBINA_REZANIYA As Double, _
    Ra As Double, Kvalitet As Integer, izmeryaemyy_razmer As Double, tocheniye_diametra As Boolean, Korka As Boolean, Udar As Boolean) As Double
    
    Dim t_chern As Double, t_chern_po_korke As Double, t_chist As Double
    Dim kolvo_prokhodov As Integer
    Dim dlina_tocheniya As Double
    Dim t_diametr As Double
    
    t_chern = 0
    dlina_tocheniya = 0.5 * (D1 - d2)
    
    If Pripusk > 2 * GLUBINA_REZANIYA_CHISTOVAYA Then

        kolvo_prokhodov = WorksheetFunction.RoundUp(Pripusk / GLUBINA_REZANIYA, 0)
        GLUBINA_REZANIYA = Pripusk / kolvo_prokhodov

        t_chern_po_korke = poperechnoye_tocheniye_chernovoye_po_rezhimam(Material, D1, dlina_tocheniya, GLUBINA_REZANIYA, Udar, Korka)

        If kolvo_prokhodov > 0 Then
            kolvo_prokhodov = kolvo_prokhodov - 1
            t_chern = kolvo_prokhodov * poperechnoye_tocheniye_chernovoye_po_rezhimam(Material, D1, dlina_tocheniya, GLUBINA_REZANIYA, Udar, False)
        End If
        
        t_chern = t_chern_po_korke + t_chern
        
    End If
    
    If tocheniye_diametra And d2 > 0 Then
        t_diametr = prodolnoye_tocheniye_chistovoye_po_rezhimam(Material, d2, dlina_tocheniya, Ra, Kvalitet, Udar)
    End If
    
    t_chist = poperechnoye_tocheniye_chistovoye_po_rezhimam(Material, D1, dlina_tocheniya, Ra, Kvalitet, izmeryaemyy_razmer, Udar)

    PoperechnoyeTocheniyeKompleks_PoRezhimam = t_chern + t_chist + t_diametr
    
End Function

Function ProdolnoyeTocheniye_PoRezhimam(Material As EnumMaterialy, diameter_zagotovki As Double, diameter_detali As Double, GLUBINA_REZANIYA As Double, _
    dlina_tocheniya As Double, Ra As Double, Kvalitet As Integer, _
    ByVal Korka As Boolean, Udar As Boolean, podrezka_tortsa As Boolean) As Double
    
    If diameter_zagotovki < diameter_detali Then Exit Function
    
    Dim t_chern As Double, t_chist As Double, t_podrezka As Double, t_na_prokhod As Double
    Dim Pripusk As Double
    Dim d_chern As Double, d_obrabotki As Double
    
    t_chern = 0
    Pripusk = (diameter_zagotovki - diameter_detali) / 2
    If Pripusk > 2 * GLUBINA_REZANIYA_CHISTOVAYA Then
        d_chern = diameter_detali + 2 * GLUBINA_REZANIYA_CHISTOVAYA
    
        d_obrabotki = diameter_zagotovki
        Do While d_obrabotki > d_chern
            t_na_prokhod = ProdolnoyeTocheniyeChernovoye_PoRezhimam(Material, d_obrabotki, GLUBINA_REZANIYA, dlina_tocheniya, Udar, Korka)
            Korka = False
            t_chern = t_chern + t_na_prokhod
            d_obrabotki = d_obrabotki - 2 * GLUBINA_REZANIYA
        Loop
        
    End If
    
    t_chist = prodolnoye_tocheniye_chistovoye_po_rezhimam(Material, diameter_detali, dlina_tocheniya, Ra, Kvalitet, Udar)
    
    If podrezka_tortsa Then
        t_podrezka = poperechnoye_tocheniye_chistovoye_po_rezhimam(Material, diameter_zagotovki, 0.5 * (diameter_zagotovki - diameter_detali), Ra, Kvalitet, RAZMER_MIN_POP, Udar)
    End If
    
    ProdolnoyeTocheniye_PoRezhimam = t_chern + t_chist + t_podrezka
    
End Function

Function RastachivaniyeKompleks_PoRezhimam(Material As EnumMaterialy, diameter_detali As Double, diameter_zagotovki As Double, GLUBINA_REZANIYA As Double, _
    dlina_tocheniya As Double, Ra As Double, Kvalitet As Integer, _
    ByVal Korka As Boolean, Udar As Boolean, podrezka_tortsa As Boolean) As Double
    
    If diameter_detali < diameter_zagotovki Then Exit Function
    
    Dim t_chern As Double, t_chist As Double, t_podrezka As Double, t_na_prokhod As Double
    Dim Pripusk As Double
    Dim d_chern As Double, d_obrabotki As Double
    
    t_chern = 0
    Pripusk = (diameter_detali - diameter_zagotovki) / 2
    If Pripusk > 2 * GLUBINA_REZANIYA_CHISTOVAYA Then
        d_chern = diameter_detali - 2 * GLUBINA_REZANIYA_CHISTOVAYA
    
        d_obrabotki = diameter_zagotovki
        Do While d_obrabotki < d_chern
            t_na_prokhod = rastachivaniye_chernovoye_po_rezhimam(Material, d_obrabotki, GLUBINA_REZANIYA, dlina_tocheniya, Udar, Korka)
            Korka = False
            t_chern = t_chern + t_na_prokhod
            d_obrabotki = d_obrabotki + 2 * GLUBINA_REZANIYA
        Loop
        
    End If
    
    t_chist = rastachivaniye_chistovoye_po_rezhimam(Material, diameter_detali, dlina_tocheniya, Ra, Kvalitet, Udar)
    
    If podrezka_tortsa Then
        t_podrezka = poperechnoye_tocheniye_chistovoye_po_rezhimam(CInt(Material), diameter_detali, 0.5 * (diameter_detali - diameter_zagotovki), Ra, Kvalitet, RAZMER_MIN_POP, Udar)
    End If
    
    RastachivaniyeKompleks_PoRezhimam = t_chern + t_chist + t_podrezka
    
End Function



Private Function ProdolnoyeTocheniyeChernovoye_PoRezhimam(Material As EnumMaterialy, Diametr As Double, GLUBINA_REZANIYA As Double, _
    dlina_tocheniya As Double, Udar As Boolean, Korka As Boolean) As Double

    Dim t_mash As Double
    Dim t_vsp As Double

    t_mash = t_mash_prodolnoye_tocheniye(Material, GLUBINA_REZANIYA, SmmOb_CHERN, Diametr, dlina_tocheniya, Udar, Korka)
    t_vsp = Tvsp_ProdolnoyeTocheniye_Chernovoye(dlina_tocheniya)
    ProdolnoyeTocheniyeChernovoye_PoRezhimam = T_OBS * (t_mash + t_vsp)
    
End Function

Private Function rastachivaniye_chernovoye_po_rezhimam(Material As EnumMaterialy, Diametr As Double, GLUBINA_REZANIYA As Double, _
    dlina_rastachivaniya As Double, Udar As Boolean, Korka As Boolean) As Double

    Dim t_mash As Double
    Dim t_vsp As Double

    t_mash = t_mash_rastachivaniye(Material, GLUBINA_REZANIYA, SmmOb_CHERN, Diametr, dlina_rastachivaniya, Udar, Korka)
    t_vsp = Tvsp_ProdolnoyeTocheniye_Chernovoye(dlina_rastachivaniya)
    rastachivaniye_chernovoye_po_rezhimam = T_OBS * (t_mash + t_vsp)
    
End Function

Private Function RaschotOborotov(V As Double, d As Double) As Double
    
    If Not (V > 0 And d > 0) Then Exit Function
    
    Dim N As Double
    Dim dopustimoye_uvelicheniye As Double
    Dim spisok_oborotov_16k20 As Variant
    Dim i As Integer
    Dim n_max_chpu As Integer
    
    N = (1000 * V) / (PI * d)

    dopustimoye_uvelicheniye = 1.05
    spisok_oborotov_16k20 = Array(12.5, 16, 20, 25, 31.5, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 500, 630, 800, 1000, 1250, 1600)
    If N * dopustimoye_uvelicheniye > spisok_oborotov_16k20(UBound(spisok_oborotov_16k20) - 1) Then
        N = spisok_oborotov_16k20(UBound(spisok_oborotov_16k20) - 1)
    Else
        For i = UBound(spisok_oborotov_16k20) To 0 Step -1
            If spisok_oborotov_16k20(i) <= N * dopustimoye_uvelicheniye Then
                N = spisok_oborotov_16k20(i)
                Exit For
            End If
        Next
    End If

    
    RaschotOborotov = N

End Function
