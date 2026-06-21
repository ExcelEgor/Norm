VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFrmSpeedy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public KolVoStrok As Integer

Private pForma As frmSpeedy
Private WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
Private WithEvents mlt As MSForms.MultiPage
Attribute mlt.VB_VarHelpID = -1

Private Enum TipKontura
    ePerimetr = 1
    eDiametr = 2
    eProizvolnyy = 3
End Enum

Private Type ElemtyDopolnitelnoyVyrezki
    cboTip As Control
    txtA As Control
    txtB As Control
    txtKontur As Control
    txtOtv As Control
    txtKolVo As Control
End Type

Private Type GabaritySpeedy
    Dlina As Double
    Shirina As Double
    Diametr As Double
    Tolshchina As Double
End Type

Private Type ParametryVyrezki
    Kontur As TipKontura
    Dlina As Double
    Shirina As Double
    d As Double
    DlinaProizvolnogoKontura As Double
    KolVo As Double
End Type

Private Type RezultatyVyrezki
    DlinaRezaniya As Double
    tKontrol As Double
End Type

Private Type ItogovyyeResultaty
    tVyrezka As Double
    tProvymvka As Double
    tKontrol As Double
End Type

Public Property Set MultiPageControl(vControl As MSForms.MultiPage): Set mlt = vControl: End Property

Public Property Set TextBoxControl(vControl As MSForms.TextBox): Set txt = vControl: End Property

Public Property Set ComboBoxControl(vControl As MSForms.ComboBox): Set cbo = vControl: End Property

Public Property Set Forma(Value As MSForms.UserForm): Set pForma = Value: End Property

Private Sub mlt_Change(): Call GlavnyyRaschot: End Sub

Private Sub txt_Change(): Call GlavnyyRaschot: End Sub

Private Sub cbo_Change(): Call GlavnyyRaschot: End Sub

Public Sub GlavnyyRaschot()

    With pForma
        
        Dim Material As EnumMaterialy
        Material = MaterialFromCboMaterial(.cboMaterial)
        
        Dim MaksimalnayaTolshchina As Double
        MaksimalnayaTolshchina = RaschotMakcsimalnoyTolshchiny_Speedy(Material)
        
        Call OchistitFormu(MaksimalnayaTolshchina)
    
        Dim TipZagotovki As TipKontura
        TipZagotovki = OpredelitTipZagotovki
        
        Dim Gabarity As GabaritySpeedy
        Gabarity = RaschotGabaritovOsnovnoyVyrezki(TipZagotovki)
        
        Call DostupnostDopolnitelnoyVyrezki(Gabarity, MaksimalnayaTolshchina)
        
        If Gabarity.Tolshchina > MaksimalnayaTolshchina Then Exit Sub

        Call ZapisatMinutnuyuPodachu(Material, Gabarity.Tolshchina)
        
        Dim Rezultaty As RezultatyVyrezki
        Rezultaty = RaschotRezultatovVyrezki(Gabarity, TipZagotovki)
    
        If Rezultaty.DlinaRezaniya <= 0 Then Exit Sub
    
        Dim Itogi As ItogovyyeResultaty
        Itogi = RaschotItogov(Material, Gabarity, Rezultaty)

        Call ZapolnitTablitsuItogami(Itogi)
        
    End With

End Sub

Private Sub OchistitFormu(MaksimalnayaTolshchina As Double)

    With pForma
        .lstOper.Clear
        .txtPodacha.text = Empty
        .lblMaksTolshchina.Caption = "До " & MaksimalnayaTolshchina & "мм"
    
    End With
    
End Sub

Private Sub DostupnostDopolnitelnoyVyrezki(Gabarity As GabaritySpeedy, MaksimalnayaTolshchina As Double)
    
    Dim Dostupno As Boolean
    Dostupno = Gabarity.Tolshchina > 0 And Gabarity.Tolshchina <= MaksimalnayaTolshchina
    
    pForma.fraDopolnitelnaya.Enabled = Dostupno
        
    Dim ctrl As Control
    For Each ctrl In pForma.fraDopolnitelnaya.Controls
        ctrl.Enabled = Dostupno
    Next
    
    If Not Dostupno Then Exit Sub
    
    Dim i As Integer
    For i = 1 To KolVoStrok
    
        Dim Elementy As ElemtyDopolnitelnoyVyrezki
        Elementy = PrisvoitElementyDopolnitelnoyeVyrezki(i)
    
        Call DostupnostElementov(Elementy)
            
    Next

    
End Sub

Private Function OpredelitTipZagotovki() As TipKontura

    Select Case pForma.mltTipZagotovki.SelectedItem.Name
        Case "pDiametr"
            OpredelitTipZagotovki = eDiametr
        Case "pPerimetr"
            OpredelitTipZagotovki = ePerimetr
        Case "pKontur"
            OpredelitTipZagotovki = eProizvolnyy
    End Select

End Function

Private Sub ZapisatMinutnuyuPodachu(Material As EnumMaterialy, Tolshchina As Double)

    Dim Smm_min As Integer

    pForma.txtPodacha.text = Empty
    If Tolshchina > 0 Then
        Smm_min = RaschotMinutnoyPodachi_Speedy(Material, Tolshchina)
        pForma.txtPodacha = Smm_min
    End If

End Sub

Private Function RaschotRezultatovVyrezki(Gabarity As GabaritySpeedy, TipZagotovki As TipKontura) As RezultatyVyrezki

    Dim Osnovnaya As RezultatyVyrezki
    Osnovnaya = RaschotOsnovnoyVyrezki(Gabarity, TipZagotovki)
    
    Dim Dopolnitelnaya As RezultatyVyrezki
    Dopolnitelnaya = RaschotDopolnitelnoyVyrezki(Gabarity.Tolshchina)

    With RaschotRezultatovVyrezki
        .DlinaRezaniya = Osnovnaya.DlinaRezaniya + Dopolnitelnaya.DlinaRezaniya
        .tKontrol = Osnovnaya.tKontrol + Dopolnitelnaya.tKontrol
    End With
    
End Function

Private Function RaschotGabaritovOsnovnoyVyrezki(TipZagotovki As TipKontura) As GabaritySpeedy

    With RaschotGabaritovOsnovnoyVyrezki

        If TipZagotovki = eDiametr Then
            .Diametr = DblFromCtrl(pForma.txtD_Diametr)
            .Dlina = DblFromCtrl(pForma.txtTolshchina)
            .Shirina = .Diametr
        ElseIf TipZagotovki = ePerimetr Then
            .Dlina = DblFromCtrl(pForma.txtA_Perimetr)
            .Shirina = DblFromCtrl(pForma.txtB_Perimetr)
        ElseIf TipZagotovki = eProizvolnyy Then
            .Dlina = DblFromCtrl(pForma.txtA_Kontur)
            .Shirina = DblFromCtrl(pForma.txtB_Kontur)
        End If

        .Tolshchina = DblFromCtrl(pForma.txtTolshchina)

    End With

End Function

Private Function RaschotOsnovnoyVyrezki(Gabarity As GabaritySpeedy, TipZagotovki As TipKontura) As RezultatyVyrezki

    With RaschotOsnovnoyVyrezki
    
        If TipZagotovki = eDiametr Then
            .DlinaRezaniya = PI * Gabarity.Diametr
            .tKontrol = IzmerShtangenCircul(Gabarity.Diametr, Gabarity.Tolshchina)
        ElseIf TipZagotovki = ePerimetr Then
            .DlinaRezaniya = 2 * (Gabarity.Dlina + Gabarity.Shirina)
            .tKontrol = IzmerShtangenCircul_CPL(Gabarity.Dlina, Gabarity.Shirina)
        ElseIf TipZagotovki = eProizvolnyy Then
            .DlinaRezaniya = DblFromCtrl(pForma.txtL_Kontur)
            If DblFromCtrl(pForma.txtKolVoRazmerov) > 2 Then
                .tKontrol = IzmerShtangenCircul(Gabarity.Dlina) + IzmerShtangenCircul(Gabarity.Shirina) * (DblFromCtrl(pForma.txtKolVoRazmerov) - 1)
            Else
                .tKontrol = IzmerShtangenCircul_CPL(Gabarity.Dlina, Gabarity.Shirina)
            End If
        End If
        
    End With

End Function

Private Function RaschotDopolnitelnoyVyrezki(Tolshchina As Double) As RezultatyVyrezki

    Dim i As Integer
    For i = 1 To KolVoStrok

        Dim Elementy As ElemtyDopolnitelnoyVyrezki
        Elementy = PrisvoitElementyDopolnitelnoyeVyrezki(i)

        Dim Parametry As ParametryVyrezki
        Parametry = PoluchitParametryVyrezki(Elementy)
        
        If Parametry.Kontur <> 0 Then
            
            Dim Resultaty As RezultatyVyrezki
            Resultaty = RaschotStrokiuDopolnitelnoyVyrezki(Parametry, Tolshchina)
            
            With RaschotDopolnitelnoyVyrezki
                .DlinaRezaniya = .DlinaRezaniya + Resultaty.DlinaRezaniya
                .tKontrol = .tKontrol + Resultaty.tKontrol
            End With
            
        End If
        
    Next

End Function

Private Function PrisvoitElementyDopolnitelnoyeVyrezki(Storka As Integer) As ElemtyDopolnitelnoyVyrezki

    With PrisvoitElementyDopolnitelnoyeVyrezki
        Set .cboTip = pForma.Controls("cboTip_" & Storka)
        Set .txtA = pForma.Controls("txtA_" & Storka)
        Set .txtB = pForma.Controls("txtB_" & Storka)
        Set .txtKontur = pForma.Controls("txtKontur_" & Storka)
        Set .txtOtv = pForma.Controls("txtOtv_" & Storka)
        Set .txtKolVo = pForma.Controls("txtKolvo_" & Storka)
    End With
    
End Function

Private Function RaschotStrokiuDopolnitelnoyVyrezki(Parametry As ParametryVyrezki, Tolshchina As Double) As RezultatyVyrezki
    
    Dim Resultaty As RezultatyVyrezki
    With Parametry
    
        If .Kontur = ePerimetr Then
            If .Dlina > 0 And .Shirina > 0 Then
                Resultaty.DlinaRezaniya = .KolVo * 2 * (.Dlina + .Shirina)
                Resultaty.tKontrol = .KolVo * IzmerShtangenCircul_CPL(.Dlina, .Shirina)
            End If
    
        ElseIf .Kontur = eProizvolnyy Then
            If .DlinaProizvolnogoKontura > 0 Then
                Resultaty.DlinaRezaniya = .KolVo * .DlinaProizvolnogoKontura
                Resultaty.tKontrol = .KolVo * IzmerShtangenCircul_CPL(.Dlina / 4, .Shirina / 4)
            End If
    
        ElseIf .Kontur = eDiametr Then
            If .d > 0 Then
                Resultaty.DlinaRezaniya = .KolVo * .d * PI
                If .d > DMAX_IZMERENIE_PROBKOI Then
                    Resultaty.tKontrol = .KolVo * IzmerShtangenCircul(.d, Tolshchina)
                Else
                    Resultaty.tKontrol = .KolVo * IzmerenieOtverstiiProbkoi(.d, Tolshchina)
                End If
            End If
    
        Else
            Resultaty.DlinaRezaniya = 0
            Resultaty.tKontrol = 0
        End If
    
    End With
    
    RaschotStrokiuDopolnitelnoyVyrezki = Resultaty
            
End Function

Private Sub DostupnostElementov(Elementy As ElemtyDopolnitelnoyVyrezki)
    
    Dim arrAll As Variant
    With Elementy
        arrAll = Array(.txtA, .txtB, .txtKolVo, .txtKontur, .txtOtv)
    End With
    
    Dim i As Integer
    For i = 0 To UBound(arrAll)
        arrAll(i).Enabled = True
        arrAll(i).BorderColor = &H8000000D
    Next

    Dim Kontur As TipKontura
    Kontur = OpredelitTipKontura(Elementy.cboTip)
    
    Dim arrDisable As Variant
    With Elementy
        If Kontur = eDiametr Then
            arrDisable = Array(.txtA, .txtB, .txtKontur)

        ElseIf Kontur = ePerimetr Then
            arrDisable = Array(.txtKontur, .txtOtv)
            
        ElseIf Kontur = eProizvolnyy Then
            arrDisable = Array(.txtA, .txtB, .txtOtv)
        Else
            arrDisable = arrAll
        End If
        
    End With
    
    For i = LBound(arrDisable) To UBound(arrDisable)
        arrDisable(i).Enabled = False
        arrDisable(i).text = Empty
        arrDisable(i).BorderColor = &HE0E0E0
    Next

End Sub

Private Function RaschotItogov(Material As EnumMaterialy, Gabarity As GabaritySpeedy, Rezultaty As RezultatyVyrezki) As ItogovyyeResultaty

    With RaschotItogov
        .tVyrezka = VyrezkaSpeedy(Material, Gabarity.Tolshchina, Rezultaty.DlinaRezaniya)
        .tProvymvka = 1.3 * Promivka_m64(1, 2, Gabarity.Shirina, Gabarity.Dlina, 2)
        .tKontrol = 1.3 * Rezultaty.tKontrol
    End With
        
End Function

Private Sub ZapolnitTablitsuItogami(Itogi As ItogovyyeResultaty)

    pForma.lstOper.Clear

    Dim Operations As Variant
    With Itogi
        Operations = Array( _
            Array("Комплектования", 10, 0), _
            Array("Слесарная", 10, 0), _
            Array("Лазерная резка неметаллов", 20, .tVyrezka), _
            Array("Слесарная", 5, .tProvymvka), _
            Array("Контрольная", 0, .tKontrol) _
            )
    End With

    Call ZapolnitListBoksIzMassiva(Operations, pForma.lstOper)

End Sub



Private Function PoluchitParametryVyrezki(Elementy As ElemtyDopolnitelnoyVyrezki) As ParametryVyrezki

    Dim Parametry As ParametryVyrezki
    With Parametry
        .Kontur = OpredelitTipKontura(Elementy.cboTip)
        .Dlina = DblFromCtrl(Elementy.txtA)
        .Shirina = DblFromCtrl(Elementy.txtB)
        .DlinaProizvolnogoKontura = DblFromCtrl(Elementy.txtKontur)
        .d = DblFromCtrl(Elementy.txtOtv)
        .KolVo = KolVoPoverhnosteiProhodov(Elementy.txtKolVo)
    End With

    PoluchitParametryVyrezki = Parametry

End Function

Private Function OpredelitTipKontura(cboTip As MSForms.ComboBox) As TipKontura
    Select Case cboTip.text
        Case "Периметр"
            OpredelitTipKontura = ePerimetr
        Case "Контур"
            OpredelitTipKontura = eProizvolnyy
        Case "Отверстие"
            OpredelitTipKontura = eDiametr
    End Select
End Function
