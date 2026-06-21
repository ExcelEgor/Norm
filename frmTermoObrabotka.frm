VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTermoObrabotka 
   Caption         =   "“ермоќбработка"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6000
   OleObjectBlob   =   "frmTermoObrabotka.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTermoObrabotka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboMaterial_Change()
    Call RaschetZakalkiOtpuska
    Call RaschetTermoObrabotki
End Sub

Private Sub cboVidSecheniya_Change()
    Call RaschetZakalkiOtpuska
    Call RaschetTermoObrabotki
End Sub





Private Sub lbl_btn_EN_Click()
    tglEN.Value = Not tglEN.Value
End Sub


Private Sub tglEN_Click()
    If tglEN.Value = True Then
        lbl_btn_EN.SpecialEffect = fmSpecialEffectSunken
    Else
        lbl_btn_EN.SpecialEffect = fmSpecialEffectRaised
    End If
    Call RaschetZakalkiOtpuska
    Call RaschetTermoObrabotki
End Sub

Private Sub txtKolVo_Change()
    txtKolVo.BackColor = vbWhite
    If IsNumeric(txtKolVo) Then
        If CDbl(txtKolVo) < 3 Then
            txtKolVo.BackColor = &HFF&
        End If
    End If
    Call RaschetZakalkiOtpuska
    Call RaschetTermoObrabotki
End Sub

Private Sub txtMassa_Change()
    Call RaschetZakalkiOtpuska
    Call RaschetTermoObrabotki
End Sub



Private Sub txtOsnovnoeVremya_Change()
    Call RaschetTermoObrabotki
End Sub

Private Sub txtSechenie_Change()
    Call RaschetZakalkiOtpuska
    Call RaschetTermoObrabotki
End Sub


Private Sub UserForm_Initialize()
    
    With cboMaterial
        .List = Split("”глеродиста€ Ћегированна€")
        .ListIndex = 0
    End With
    
    With cboVidSecheniya
        .List = Split(" руглое  вадратное ѕр€моугольное")
        .ListIndex = 0
    End With
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
End Sub


Private Sub RaschetZakalkiOtpuska()

    Dim ћассаќднойƒетали As Double: ћассаќднойƒетали = DblFromCtrl(txtMassa)
    Dim —ечение As Double:          —ечение = DblFromCtrl(txtSechenie)
    Dim  ол¬оƒеталей As Double:      ол¬оƒеталей = DblFromCtrl(txtKolVo)
    
    Call EdinicaNormirovaniya( ол¬оƒеталей)
    
    If ћассаќднойƒетали > 0 And —ечение > 0 Then
        txtZakalka.text = VremyaEN(tglEN.Value, Zakalka(cboMaterial.ListIndex + 1, cboVidSecheniya.ListIndex + 1, CDbl(txtSechenie), CDbl(txtMassa), ,  ол¬оƒеталей),  ол¬оƒеталей)
        txtOtpusk.text = VremyaEN(tglEN.Value, Otpusk(cboMaterial.ListIndex + 1, CDbl(txtMassa), ,  ол¬оƒеталей),  ол¬оƒеталей)
    Else
        txtZakalka.text = Empty
        txtOtpusk.text = Empty
    End If
    
End Sub

Private Sub RaschetTermoObrabotki()
    
    Dim ќсновное¬рем€ As Double:        ќсновное¬рем€ = DblFromCtrl(txtOsnovnoeVremya)
    Dim ћассаќднойƒетали As Double:     ћассаќднойƒетали = DblFromCtrl(txtMassa)
    Dim  ол¬оƒеталей As Double:          ол¬оƒеталей = DblFromCtrl(txtKolVo)
    
    Call EdinicaNormirovaniya( ол¬оƒеталей)
    
    If ќсновное¬рем€ > 0 And ћассаќднойƒетали > 0 Then
        Dim ¬спогательное¬рем€ As Double
        ¬спогательное¬рем€ = ZagruzkaNaPodPechi(ћассаќднойƒетали,  ол¬оƒеталей) + VygruzkaIzPechiVTaru(ћассаќднойƒетали,  ол¬оƒеталей)
        txtTsht.text = VremyaEN(tglEN.Value, TermoObrabotka(ќсновное¬рем€, ¬спогательное¬рем€, ,  ол¬оƒеталей),  ол¬оƒеталей)
    Else
        txtTsht.text = Empty
    End If
    
End Sub

Private Function VremyaEN(≈Ќ As Boolean, “шт As Double,  ол¬оƒеталей As Double) As Variant
    
    Dim t As Double
    
    If ≈Ќ Then
        t = “шт *  ол¬оƒеталей
    Else
        t = “шт
    End If
    
    VremyaEN = CStr(OkruglenieTsht(t))
    
End Function

Private Sub EdinicaNormirovaniya( ол¬оƒеталей As Double)

    Dim KolVo As String: KolVo = CStr(IIf( ол¬оƒеталей = 0, 1,  ол¬оƒеталей))

    If tglEN.Value = True Then
        txtEN_DrugieOper.text = KolVo
        txtEN_Otpusk.text = KolVo
        txtEN_Zakalka.text = KolVo
    Else
        txtEN_DrugieOper.text = 1
        txtEN_Otpusk.text = 1
        txtEN_Zakalka.text = 1
    End If
    
End Sub












