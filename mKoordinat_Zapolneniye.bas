Attribute VB_Name = "mKoordinat_Zapolneniye"
Option Explicit
Private frm As frmKoordinat

Public Sub ZapolnitElementyUpravleniya_Koordinat()

    Set frm = frmKoordinat

    Dim SovmOsi
    SovmOsi = Array("По нониусной или масштабной линейке", "По штихмасу или измерительными плитками", "Без совмещения осей")

    Dim MATERIALY
    MATERIALY = Array(EnumMaterialy.ALUMINIYEVYYE_SPLAVY, EnumMaterialy.MEDNYYE_SPLAVY, EnumMaterialy.STAL_UGLERODISTAYA, _
                        EnumMaterialy.STAL_LEGIROVANNAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA, EnumMaterialy.TITANOVYYE_SPLAVY)
                        
    With frm
    
        Call DobavitMaterialyVListBox(MATERIALY, .cboMaterial)
        
        Call ZapolnitParametryGlavnogoKoeffitsiyenta
        
        Call ZapolnitSposobyUstanovki
        
        With .cboTochnostViverki_Koordinat
            .AddItem "±0,005 мм"
            .AddItem "±0,01 мм"
            .AddItem "±0,05 мм"
            .AddItem "без выверки"
            .ListIndex = 1
        End With
    
        'Растачивание
        With .cboDopusk_Rastachivaniye
            .AddItem "-"
            .AddItem "0,002"
            .AddItem "0,005"
            .AddItem "0,01"
            .ListIndex = 0
        End With
        With .cboIT_Rastachivaniye
            .List = Split("11-12 9 7-8 6")
            .ListIndex = 0
        End With
        
        'Растачивание выточек
        With .cboIT_Vytochki
            .List = Split("14 11-12 9 7 6")
            .ListIndex = 0
        End With
        .cboDopusk_Vytochki.List = .cboDopusk_Rastachivaniye.List
        .cboDopusk_Vytochki.ListIndex = 0
    
    End With
    
End Sub

Private Sub ZapolnitParametryGlavnogoKoeffitsiyenta()
    
    With frm
    
        With .cboTipProizv_Koordinat
            .AddItem "Единичное"
            .AddItem "Мелкосерийное"
            .AddItem "Среднесерийное"
            .AddItem "Массовое и крупносерийное"
            .ListIndex = 0
        End With
        
        With .cboTipStanka_Koordinat
            .List = Split("Одностоечный Двухстоечный")
            .ListIndex = 0
        End With
        
        With .cboVozrast_Koordinat
            .List = Split("До 10 лет;Свыше 10 до 20 лет;Свыше 20 лет", ";")
            .ListIndex = .ListCount - 1
        End With
        
        With .cboMaterialInstrument_Koordinat
            .AddItem "Т15К6, ВК8, Р18"
            .AddItem "Р6М5"
            .AddItem "9ХС, У10, У12"
            .ListIndex = 0
        End With
    
    End With
        
End Sub

Private Sub ZapolnitSposobyUstanovki()

    Dim Normativ
    Normativ = ZagruzitNormativ("koordinat_ustanovka")

    Dim i As Integer
    For i = LBound(Normativ) To UBound(Normativ)
        Call DobavitUnikalnuyuZapisVListBoks(Normativ, frm.lstSposobUst_Koordinat, i, Array(1, 2))
    Next
    
End Sub

Public Sub ZapolnitKharakterVyverki_Koordinat(frm As frmKoordinat)
    
    With frm
    
        .cboHarakterViverki_Koordinat.Clear
        If .lstSposobUst_Koordinat.ListIndex = -1 Then Exit Sub
    
        Dim m As Double
        m = DblFromCtrl(.txtMass_Koordinat)
    
        Dim Normativ
        Normativ = ZagruzitNormativ("koordinat_ustanovka")
        
        Dim i As Integer
        For i = LBound(Normativ) To UBound(Normativ)
            If Normativ(i, 2) = CInt(.lstSposobUst_Koordinat.List(.lstSposobUst_Koordinat.ListIndex, 1)) Then
                Call DobavitUnikalnuyuZapisVListBoks(Normativ, .cboHarakterViverki_Koordinat, i, Array(3, 4))
            End If
        Next
        
        frm.cboHarakterViverki_Koordinat.ListIndex = 0
    
    End With

End Sub

Private Sub DobavitUnikalnuyuZapisVListBoks(Normativ, lst As Control, i As Integer, NomeraStolbtsov)

    If lst.ListCount = 0 Then
        Call DobavitZapis(Normativ, lst, i, NomeraStolbtsov)
    Else
        If Normativ(i, NomeraStolbtsov(LBound(NomeraStolbtsov))) <> Normativ(i - 1, NomeraStolbtsov(LBound(NomeraStolbtsov))) Then
            Call DobavitZapis(Normativ, lst, i, NomeraStolbtsov)
        End If
    End If
    
End Sub

Private Sub DobavitZapis(Normativ, lst As Control, i As Integer, NomeraStolbtsov)
    
    Dim j As Integer
    With lst
        .AddItem
        For j = LBound(NomeraStolbtsov) To UBound(NomeraStolbtsov)
            .List(.ListCount - 1, j) = Normativ(i, NomeraStolbtsov(j))
        Next
    End With
    
End Sub

