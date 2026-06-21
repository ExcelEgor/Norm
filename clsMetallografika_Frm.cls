VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMetallografika_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Function ProverkaParametrov( _
    ByVal DlinaDetali As Double, _
    ByVal ShirinaDetali As Double, _
    ByVal DlinaZagotovki As Double, _
    ByVal ShirinaZagotovki As Double, _
    ByVal TolshchinaZagotovki As Double, _
    ByRef Oshibki As Collection) As Boolean
    
    Dim ZapolnenoVerno As Boolean
    ZapolnenoVerno = True
    Set Oshibki = New Collection
    
    If DlinaZagotovki > MAX_DLINA_GILOTIONA Then Oshibki.Add "txtL_Zagotovka": ZapolnenoVerno = False
    If ShirinaZagotovki > MAX_SHIRINA_GILOTIONA Then Oshibki.Add "txtB_Zagotovka": ZapolnenoVerno = False
    If TolshchinaZagotovki > 1 Then Oshibki.Add "txtS_Zagotovka": ZapolnenoVerno = False
    
    If DlinaDetali > DlinaZagotovki Then Oshibki.Add "txtL_Detal": ZapolnenoVerno = False
    If ShirinaDetali > ShirinaZagotovki Then Oshibki.Add "txtB_Detal": ZapolnenoVerno = False

    If Not PolozhitelnyyeChisla(DlinaDetali, ShirinaDetali, DlinaZagotovki, ShirinaZagotovki, TolshchinaZagotovki) Then
        ZapolnenoVerno = False
    End If
    
    ProverkaParametrov = ZapolnenoVerno
    
End Function

Public Function VychislitMarshrut( _
    ByVal DlinaDetali As Double, _
    ByVal ShirinaDetali As Double, _
    ByVal DlinaZagotovki As Double, _
    ByVal ShirinaZagotovki As Double, _
    ByVal TolshchinaZagotovki As Double, _
    ByVal IsDatronEnabled As Boolean, _
    ByVal SumDatron As Double, _
    ByVal SumSles As Double, _
    ByVal SumKontrol As Double, _
    ByVal SumVyrez As Double, _
    ByVal KolVoDetaley As Long, _
    ByVal Pritupleniye As Boolean) As Variant

    Dim tDatron As Double, tSles As Double, tKontrol As Double, tVyrez As Double
    
    If IsDatronEnabled Then
        Dim NormyDatron_NaruzhKontur As Variant
        NormyDatron_NaruzhKontur = Datron_tMashPerimetr(DlinaDetali, ShirinaDetali, False, Pritupleniye, 1)
            
        tDatron = Tsht_GravirovkaDatron_Metallografika(DlinaDetali, ShirinaDetali, SumDatron + NormyDatron_NaruzhKontur(1), 4, 1)
        tSles = SumSles + NormyDatron_NaruzhKontur(2)
        tKontrol = SumKontrol + NormyDatron_NaruzhKontur(3)
        tVyrez = SumVyrez + NormyDatron_NaruzhKontur(4)
    End If
        
    Dim Metallograf As New clsMetallografika
    With Metallograf
        .Init DlinaZagotovki, ShirinaZagotovki, TolshchinaZagotovki, _
              DlinaDetali, ShirinaDetali, KolVoDetaley, _
              tDatron, tSles, tKontrol, tVyrez
    End With
    
    VychislitMarshrut = Metallograf.Marshrut
    Set Metallograf = Nothing
    
End Function

