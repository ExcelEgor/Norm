VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTokFrez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As String
Public OpisaniyePerekhoda As String
Public D1 As Double
Public d2 As Double
Public DBObnizki As Double
Public DlinaObnizki As Double
Public FrezPosleSverl As Boolean
Public GlubinaObnizki As Double
Public GlubinaPaza As Double
Public GlubinaRezaniya As Double
Public IT As Integer
Public KalibrRezbSlesar As Boolean
Public KharkterVyverki As Integer
Public KolVoDetVPrutke As Double
Public KolVoPoverkhnostey As Double
Public KolVoProkhodov As Double
Public Korka As Boolean
Public L As Double
Public Lrezby As Double
Public Material As EnumMaterialy
Public Massa As Double
Public Otkloneniye As Double
Public PodrezkaTortsa As Boolean
Public Pripusk As Double
Public Ra As Double
Public ShagRezby As Double
Public ShirinaPoverhnosti As Double
Public Sposob As Integer
Public TipPoverkhnosti As Integer
Public TipStanka As EnumTipStankaRastFrez
Public TipTsentrovaniya As Integer
Public TocheniyeDiametra As Boolean
Public TochnostVyverki As Double
Public Udar As Boolean
Public Vconst As Boolean
Public VUpor As Boolean
Public UchestFasku As Boolean
Public ZachistkaBokovykhStoron As Boolean
Public UchestSlesar As Boolean
Public UchestKontrol As Boolean
Public Tonkostennaya As Boolean
Public PoRezhimam As Boolean
Public CHPU As Boolean

Public AvtoRezhimy As Boolean

Public Dfrezy As Double

Public Rezhimy As clsMillingRegimes
      
Public Function RaschotNormTokFrezIzParametrov() As RezultatyTokFrez

    Dim Normy As Variant
    Dim OpisaniyePerekhoda As String
    
    If KolVoPoverkhnostey = 0 Then KolVoPoverkhnostey = 1
    If KolVoProkhodov = 0 Then KolVoProkhodov = 1


    Select Case id
            
        Case "pUst"
            Normy = Ustanov_TokFrez(Sposob, Massa, KharkterVyverki, TochnostVyverki)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Установить", KolVoPoverkhnostey)
            
        Case "pOtrezka"
            Normy = OtrezkaNaTokarnomStanke(Material, D1, L)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Отрезать Ф" & D1, KolVoPoverkhnostey)
            
        Case "pPoperechnoye"
            Normy = PoperechnoeTocheniye(Material, D1, d2, Pripusk, IT, Ra, GlubinaRezaniya, Korka, Udar, Tonkostennaya, Vconst, TocheniyeDiametra, PoRezhimam)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Подрезать торец Ф" & D1, KolVoPoverkhnostey)
            
        Case "pProdolnoye"
            Normy = ProdolnoeTochenie(Material, D1, d2, L, IT, Ra, GlubinaRezaniya, Korka, Udar, Tonkostennaya, PodrezkaTortsa, PoRezhimam)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Точить Ф" & d2 & "H" & IT, KolVoPoverkhnostey)
            
        Case "pRastachivaniye"
            Normy = Rastachivanie(Material, D1, d2, L, IT, Ra, GlubinaRezaniya, Korka, VUpor, Udar, Tonkostennaya, PodrezkaTortsa, PoRezhimam)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Расточить Ф" & D1 & "H" & IT, KolVoPoverkhnostey)
            
        Case "pKonus"
            Normy = Konus(Material, D1, L, Ra, IT, TipPoverkhnosti, Tonkostennaya)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Точить конус Ф" & D1 & "H" & IT, KolVoPoverkhnostey)
            
        Case "pFaski"
            Normy = ObrabotkaFasok(Material, D1)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Точить фаску Ф" & D1, KolVoPoverkhnostey)
            
        Case "pGaltel"
            Normy = ObtachivaniyeGalteley(Material, D1, L, Ra)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Точить галтель R" & L & " Ф" & D1, KolVoPoverkhnostey)
            
        Case "pSverl_Tok"
            Normy = ObrabotkaOtverstiy_Tokarnyy(TipObrabotkiOtverstiya_Tokar.Sverleniye, Material, D1, L, VUpor)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Сверлить Ф" & D1, KolVoPoverkhnostey)
            
        Case "pRassverl_Tok"
            Normy = ObrabotkaOtverstiy_Tokarnyy(Rassverlivaniye, Material, D1, L, VUpor)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Рассверлить Ф" & D1, KolVoPoverkhnostey)
            
        Case "pZenker_Tok"
            Normy = ObrabotkaOtverstiy_Tokarnyy(Zenkerovaniye, Material, D1, L, VUpor)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Зенкеровать Ф" & D1, KolVoPoverkhnostey)
            
        Case "pRazvert_Tok"
            Normy = ObrabotkaOtverstiy_Tokarnyy(IIf(Ra > 125, Razvertyvaniye_Ra2_5, Razvertyvaniye_Ra1_25), Material, D1, L, VUpor)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Развернуть Ф" & D1 & " Ra" & Ra, KolVoPoverkhnostey)
                
        Case "pNakativaniyeRifleniy"
            Normy = NakativaniyeRifleniy(Material, D1, L)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Накатать рифление Ф" & D1, KolVoPoverkhnostey)
            
        Case "pKanavka"
            Normy = Kanavka(Material, TipPoverkhnosti, D1, L, ShirinaPoverhnosti, IT, Ra)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey(IIf(TipPoverkhnosti = 1, "Точить ", "Расточить ") & "канавку Ф" & D1, KolVoPoverkhnostey)
            
        Case "pMetchik"
            Normy = NarezaniyeRezbyMetchikom(Material, D1, L, IT, VUpor, ShagRezby)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Нарезать резьбу М" & D1 & "x" & ShagRezby, KolVoPoverkhnostey)
                
        Case "pPlashka"
            Normy = NarezaniyeRezbyPlashkoy(Material, D1, L, ShagRezby, IT)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Нарезать резьбу М" & D1 & "x" & ShagRezby, KolVoPoverkhnostey)
                
        Case "pMetrRezb"
            Normy = NarezanieMetricheskoiRezbiRezcom(TipPoverkhnosti, Material, D1, L, IT, VUpor, ShagRezby)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Нарезать резьбу М" & D1 & "x" & ShagRezby, KolVoPoverkhnostey)
                
        Case "RezbaAvtoVybor"
            Normy = NarezaniyeRezbyNaTokarnomStankeAvtoVyborSposoba(Material, D1, L, TipPoverkhnosti, VUpor, IT, ShagRezby)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Нарезать резьбу М" & D1 & "x" & ShagRezby, KolVoPoverkhnostey)
              
        Case "pDovodka"
            Normy = Dovodka(Material, TipPoverkhnosti, VUpor, D1, L, Ra, Otkloneniye, Tonkostennaya)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Довести поверхность Ф" & D1 & "до Ra" & Ra, KolVoPoverkhnostey)
            
        Case "pObrabOtvSNulya"
            Normy = ObrabotkaOtverstiySNulya(Material, D1, L, VUpor, IT, Ra, GlubinaRezaniya, Tonkostennaya, Vconst)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Обработать отверстие Ф" & D1, KolVoPoverkhnostey)
            
        Case "pTsentrovaniye"
            Normy = Tsentrovaniye(Material, D1, TipTsentrovaniya, Ra)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Центровать Ф" & D1, KolVoPoverkhnostey)
            
        Case "pPritupleniye"
            Normy = PritupleniyeKromok(Material, D1)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Притупить кромку Ф" & D1, KolVoPoverkhnostey)
            
        Case "pPolirovaniyeTok"
            Normy = PolirovaniyeNaTokranomStankeShkurkoy(D1, L, Ra, TipPoverkhnosti)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Полировать поверхность Ф" & D1 & " Ra" & Ra, KolVoPoverkhnostey)
            
        Case "pDovodkaTsentrov"
            Normy = DovodkaTsentrov()
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Довести центр", KolVoPoverkhnostey)
            
            'ФРЕЗЕРОВАНИЕ
        Case "pUstRastFrez"
            Normy = Ustanovka_Frezernyye_i_Rastochnyye(Massa, Sposob, KharkterVyverki, TochnostVyverki)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Установить", KolVoProkhodov)
                
        Case "pFrezPloskKonc"
            ' Создаем и инициализируем контекст
            Dim ctx As New clsMillingContext
            With ctx
                .AvtoRezhimy = AvtoRezhimy
                .CHPU = CHPU
                .DlinaPoverkhnosti = L
                .IT = IT
                .Korka = Korka
                .Material = Material
                .NaUdar = Udar
                .PoRezhimam = PoRezhimam
                .Pripusk = Pripusk
                .Ra = Ra
                .ShirinaPoverkhnosti = ShirinaPoverhnosti
                .TipStanka = TipStanka
                .TonkostennayaKonfiguraciya = Tonkostennaya
                If Not AvtoRezhimy Then
                    .Dfrezy = Dfrezy
                    Set .Rezhimy = Rezhimy
                End If
            End With
            
            ' Передаем один объект вместо 15 переменных
            Normy = FrezKontsevymiFrezami_Kompleks(ctx)
            
        Case "pFrezPloskTorc"
            Normy = FrezTortsovymyFrezami_Kompleks(TipStanka, Material, Pripusk, ShirinaPoverhnosti, L, IT, Ra, Udar, Korka, Tonkostennaya)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Фрез-ть пл-ть торцевой фр.", KolVoPoverkhnostey)
            
        Case "pFrezUstup":
            Normy = FrezUstupov_Kompleks(TipStanka, Material, L, ShirinaPoverhnosti, Pripusk, IT, Ra, Udar, Korka, Tonkostennaya)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Фрез уступ " & ShirinaPoverhnosti & "x" & Pripusk, KolVoPoverkhnostey)
                
        Case "pFrezPazov"
            Normy = FrezPazov_Kompleks(TipStanka, Material, L, ShirinaPoverhnosti, GlubinaPaza, IT, Ra, VUpor, Udar, Korka, Tonkostennaya, ZachistkaBokovykhStoron)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Фрез-ть паз " & ShirinaPoverhnosti & "x" & GlubinaPaza, KolVoPoverkhnostey)
            
        Case "pFrezShlic"
            Normy = FrezShlitsevymiFrezami_Kompleks(TipStanka, Material, ShirinaPoverhnosti, L, GlubinaRezaniya, Ra, Udar, Tonkostennaya)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Фрез-ть шлиц " & ShirinaPoverhnosti & "x" & GlubinaRezaniya, KolVoPoverkhnostey)
            
        Case "pFrezOtverstii":
            Normy = ObrabotkaOtverstiiNaFrezernomStanke(TipStanka, Material, D1, L, VUpor, DBObnizki, GlubinaObnizki, IT, Ra, Tonkostennaya, FrezPosleSverl, UchestFasku)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey(IIf(FrezPosleSverl, "Сверл., фрез.", "Сверлить") & " отв. Ф" & D1, KolVoPoverkhnostey)
            
        Case "pObrabotkaRezbovihOtv"
            Normy = ObrabotkaRezbovihOtverstiiNaFrezernimStanke(TipStanka, Material, D1, L, Lrezby, VUpor, DBObnizki, GlubinaObnizki, IT, ShagRezby, Tonkostennaya, KalibrRezbSlesar, FrezPosleSverl)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Обработать резьб отв. М" & D1, KolVoPoverkhnostey)

        Case "pFrezOkon"
            Normy = FrezOkon(Material, L, ShirinaPoverhnosti, Pripusk, VUpor, DlinaObnizki, DBObnizki, GlubinaObnizki, IT, Ra, Tonkostennaya, TipStanka)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Фрез-ть окно " & L & "x" & ShirinaPoverhnosti, KolVoPoverkhnostey)
                
        Case "pFrezKnopki"
            Normy = FrezOtverstiyPodKvadratnyieKnopki(L, Pripusk, Material)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Фрез-ть отв. под кнопку " & L, KolVoPoverkhnostey)
       
        Case "pShestigrannik"
            Normy = FrezShestigrannikaIzKrugaKoncevoyFrezoy(Material, ShagRezby, ShirinaPoverhnosti, Massa, IT, Ra, D1, VUpor)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Фрез-ть шестигранник", KolVoPoverkhnostey)

    End Select

    With RaschotNormTokFrezIzParametrov
        If IsArray(Normy) Then
            .Normy(1) = Normy(1) * KolVoPoverkhnostey * KolVoProkhodov
            .Normy(2) = IIf(UchestSlesar, Normy(2) * KolVoPoverkhnostey, 0)
            .Normy(3) = IIf(UchestKontrol, Normy(3) * KolVoPoverkhnostey, 0)
        Else
            .Normy(1) = Normy * KolVoPoverkhnostey * KolVoProkhodov
            .Normy(2) = 0
            .Normy(3) = 0
        End If
        .OpisaniyePerekhoda = OpisaniyePerekhoda
    End With

        
End Function


