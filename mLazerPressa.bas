Attribute VB_Name = "mLazerPressa"
Option Explicit

Function ZachistkaZausentsev_IzTablitsy(Material As EnumMaterialy, Massa, Diametr As Double, Kontur As Double, a As Double, b As Double, _
    ByVal KolVoPoverkhnostey As Double, ByVal KolVoStoron As Double) As Double
    
    If KolVoPoverkhnostey = 0 Then KolVoPoverkhnostey = 1
    If KolVoStoron = 0 Then KolVoStoron = 2
    
    Dim tOtv As Double, tKontur As Double, tPerimetr As Double
    
    tOtv = ZachistkaZausencevSOtverstiiVruchnuyu(Material, Diametr, 2, 1, Massa > 20)
    
    tKontur = ZachistkaZausencev_PoKonturu_Napilnikom(Material, Kontur, False, Massa > 20)
    
    tPerimetr = ZachistkaZausencev_PoKonturu_Napilnikom(Material, 2 * (a + b), False, Massa > 20)
    
    ZachistkaZausentsev_IzTablitsy = KolVoPoverkhnostey * KolVoStoron * (tOtv + tKontur + tPerimetr)
    
End Function

Function Kontrol_IzTablitsy(TolshchinaLista As Double, Dotv As Double, Kontur As Double, a As Double, b As Double, ByVal KolVoPoverkhnostey As Double)
    
    If KolVoPoverkhnostey = 0 Then KolVoPoverkhnostey = 1

    Dim tOtv As Double, tKontur As Double, tPerimetr As Double
    
    tOtv = IzmerenieOtverstiiProbkoi(Dotv, TolshchinaLista, , KolVoPoverkhnostey)
    tKontur = IzmerShtangenCircul(Kontur / 4, Kontur / 4, False, KolVoPoverkhnostey)
    tPerimetr = IzmerShtangenCircul(a, b, False, KolVoPoverkhnostey)
    
    Kontrol_IzTablitsy = tOtv + tKontur + tPerimetr

End Function

