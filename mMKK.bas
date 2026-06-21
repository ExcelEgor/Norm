Attribute VB_Name = "mMKK"
Option Explicit

Function MKK_FrezerovatIzCilindra2Obraztsa(DiametrZagotovki As Double, DlinaZagotovki As Double, TolshchinaObraztsa As Double, ShirinaObraztsa As Double, DlinaObraztsa As Double, _
    Optional Material As EnumMaterialy = EnumMaterialy.STAL_NERZHAVEYUSHCHAYA) As Variant
    'Фрезерование из цилиндрической заготовки 2 образца на универсальном фрезерном станке

    Dim MassaZagotovki As Double
    MassaZagotovki = MassaKruga(DiametrZagotovki, DlinaZagotovki, Material)
    
    Dim tUstanovka As Double
    tUstanovka = Ustanov_Vtiskah(MassaZagotovki) * 6
    
    Dim Dfrezy As Integer:      Dfrezy = 80
    Dim sFrezy As Double:       sFrezy = 16
    Dim tRezPoDline As Double:  tRezPoDline = (DlinaZagotovki + Dfrezy) / sFrezy
    
    Dim PripuskNaStoronu As Double
    
    PripuskNaStoronu = ((DiametrZagotovki / 2) - TolshchinaObraztsa) / 2
    Dim tFrezTolshchina As Double
    tFrezTolshchina = FrezTortsovymyFrezami_Kompleks(2, CInt(Material), PripuskNaStoronu, DiametrZagotovki, DlinaZagotovki, 11, 5)(1)
    
    PripuskNaStoronu = (DiametrZagotovki - ShirinaObraztsa) / 2
    Dim tFrezShirina As Double
    tFrezShirina = FrezTortsovymyFrezami_Kompleks(2, CInt(Material), PripuskNaStoronu, DiametrZagotovki, DlinaZagotovki, 11, 5)(1)
    
    PripuskNaStoronu = (DlinaZagotovki - DlinaObraztsa) / 2
    Dim tFrezDlina As Double
    tFrezDlina = FrezTortsovymyFrezami_Kompleks(2, CInt(Material), PripuskNaStoronu, DiametrZagotovki, DiametrZagotovki, 11, 5)(1)
    
    Dim K As Double:                K = 1.3
    Dim KolVoObraztsov As Integer:  KolVoObraztsov = 2
    Dim KolVoOStoron As Integer:    KolVoOStoron = 2
    
    Dim Tsht_Frezernaya As Double
    Tsht_Frezernaya = K * KolVoOStoron * KolVoObraztsov * (tFrezTolshchina + tFrezShirina + tFrezDlina)
    
    Dim Tsht_Slesarnaya As Double
    Tsht_Slesarnaya = K * KolVoObraztsov * SnyatiyeZausentsevList(DlinaObraztsa, ShirinaObraztsa, TolshchinaObraztsa, CInt(Material))
    
    Dim Tsht_Kontrolnaya As Double
    Tsht_Kontrolnaya = K * KolVoObraztsov * _
    (IzmerShtangenCircul(TolshchinaObraztsa, DlinaZagotovki) + IzmerShtangenCircul(ShirinaObraztsa, DlinaZagotovki) + IzmerShtangenCircul(DlinaZagotovki, ShirinaObraztsa))
    
    Dim MassivTsht(1 To 3) As Double
    MassivTsht(1) = Tsht_Frezernaya
    MassivTsht(2) = Tsht_Slesarnaya
    MassivTsht(3) = Tsht_Kontrolnaya
    
    MKK_FrezerovatIzCilindra2Obraztsa = MassivTsht

End Function

Function MKK_VyrezatIzObraztsaNeobkhodimoyeKolVoPlastin(KolVoPlastin As Double, TolshchinaZagotovki As Double, DlinaObraztsa As Double, ShirinaObraztsa As Double, TolshchinaObraztsa As Double, _
    Optional Material As EnumMaterialy = EnumMaterialy.STAL_NERZHAVEYUSHCHAYA) As Variant
    
    Dim MassaObraztsa As Double
    MassaObraztsa = MassaLista(DlinaObraztsa, ShirinaObraztsa, TolshchinaObraztsa, Material)
    
    Dim tUstanovka As Double
    tUstanovka = 6 * Ustanov_Vtiskah(MassaObraztsa)
    
    Dim tVyrezka As Double
    tVyrezka = FrezPazov_Kompleks(VertikalnoGorizontalnoFrezernyy, CInt(Material), 2 * (DlinaObraztsa + ShirinaObraztsa), _
        16, TolshchinaZagotovki, 14, 12.5, False, False, False, False, False)(1)
        
    Dim tFrezTolshchina As Double
    tFrezTolshchina = 2 * FrezTortsovymyFrezami_Kompleks(VertikalnoGorizontalnoFrezernyy, CInt(Material), _
        0.5 * (TolshchinaZagotovki - TolshchinaObraztsa), ShirinaObraztsa, DlinaObraztsa, 11, 5)(1)
    
    Dim tFrezDlina As Double
    tFrezDlina = 2 * FrezTortsovymyFrezami_Kompleks(2, CInt(Material), 1, TolshchinaZagotovki, DlinaObraztsa, 11, 5)(1)
    
    Dim tFrezShirina As Double
    tFrezShirina = 2 * FrezTortsovymyFrezami_Kompleks(2, CInt(Material), 1, TolshchinaZagotovki, ShirinaObraztsa, 11, 5)(1)
    
    Dim kEdinichnoye As Double
    kEdinichnoye = 1.3
    
    Dim tFrezernaya As Double
    tFrezernaya = kEdinichnoye * KolVoPlastin * (tUstanovka + tVyrezka + tFrezTolshchina + tFrezDlina + tFrezShirina)
    
    Dim tSlesarnaya As Double
    tSlesarnaya = kEdinichnoye * KolVoPlastin * SnyatiyeZausentsevList(DlinaObraztsa, ShirinaObraztsa, TolshchinaObraztsa, CInt(Material))
    
    Dim tKontrolnaya As Double
    tKontrolnaya = kEdinichnoye * KolVoPlastin * _
    (IzmerShtangenCircul(TolshchinaObraztsa, DlinaObraztsa) + IzmerShtangenCircul(ShirinaObraztsa, DlinaObraztsa) + IzmerShtangenCircul(DlinaObraztsa, ShirinaObraztsa))
    
    MKK_VyrezatIzObraztsaNeobkhodimoyeKolVoPlastin = Array(tFrezernaya, tSlesarnaya, tKontrolnaya)
    
End Function

