VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLentochnoOtreznieStanki 
   Caption         =   "Отрезка заготовок на ленточно-отрезных станках"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   OleObjectBlob   =   "frmLentochnoOtreznieStanki.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLentochnoOtreznieStanki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Handler As clsGlavnyyRaschot
Private img As clsLentOtrez_Img

Private ControlHandlers As Collection

Private Figura As ITechnologicalShape

Private UstanovTokar As Boolean, UstanovFrez As Boolean

Private Material As EnumMaterialy
Private Massa As Double
Private D1 As Double, d2 As Double, TocheniyeNaUdar As Boolean
Private Shirina As Double, Vysota As Double

Private Sub cmdUstanovit_Click()

    If UstanovTokar Then
        Me.RaschotTokarnoyOperatsii
    ElseIf UstanovFrez Then
        Me.FrezerovatTorcy
    End If
    
End Sub

Private Sub mltLentOtrez_Change()

    Select Case mltLentOtrez.SelectedItem.Name
        Case "pUgolnik", "pShveller", "pShveller2", "pPryamougolTruba", "pPramougol", "pShestigran"
            chkPravka.Visible = True
        Case Else
            chkPravka.Visible = False
    End Select
    
End Sub

Private Sub cboShveller_Nomer_Change()

    Dim ArrParametryShvellera As Variant
    ArrParametryShvellera = ParametryShvellera(cboShveller_Nomer)

    Dim a As Integer, b As Integer
    Dim S As Double
    
    a = ArrParametryShvellera(1, 2)
    b = ArrParametryShvellera(1, 3)
    S = ArrParametryShvellera(1, 4)
    
    lblShveller_A.Caption = a
    lblShveller_B.Caption = b
    lblShveller_s.Caption = S
    
End Sub

Private Function ParametryShvellera(NomerShvellera As String)

    Dim tbShveller As ListObject
    Set tbShveller = wsShveller.ListObjects(1)
    
    Dim ArrParametryShvellera()
    ArrParametryShvellera = tbShveller.ListRows(WorksheetFunction.Match(NomerShvellera, tbShveller.ListColumns(1).DataBodyRange, 0)).Range.Value
    ParametryShvellera = ArrParametryShvellera
    
End Function

Private Sub UserForm_Initialize()

    mltLentOtrez.Style = fmTabStyleNone
        
    With cboShveller_Nomer
        .List = Split("5У 6.5У 8У 10У 12У 14У 16У 16аУ 18У 18аУ 20У 22У 24У 27У 30У 33У 36У 40У")
        .ListIndex = 0
    End With

    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    Set ControlHandlers = New Collection
    DobavitVKlass_GlavnyyRaschot Me, ControlHandlers, Handler
    
    Dim ctrl As Control
    For Each ctrl In fraKnopki.Controls
        If ctrl.Name Like "img_p*" Then
            Set img = New clsLentOtrez_Img
            Set img.ImageControl = ctrl
            
            ControlHandlers.Add img
        End If
    Next
    
    VyravnitElementy
    
    OtkrytVkladku img_pKrug
    
End Sub

Private Sub VyravnitElementy()

    VyravnitKartinki

    With fraMaterial
        .Top = 3
        .Left = 6
    End With

    With fraKnopki
        .Top = fraMaterial.Top + fraMaterial.Height + 3
        .Left = fraMaterial.Left
        .Width = 240
    End With
    
    With fraParametry
        .Top = fraKnopki.Top + fraKnopki.Height + 3
        .Left = fraMaterial.Left
        .Width = fraKnopki.Width
    End With
    
    With mltLentOtrez
        .Left = -1
        .Top = 6
        .Width = fraParametry.Width
        .Height = 135
    End With
    
    With txtShveller_L
        .Left = 147
        .Top = mltLentOtrez.Top + mltLentOtrez.Height + 6
    End With
    
    With lblLentOtrez_Shveller_L
        .Left = imgUgolnik.Width + imgUgolnik.Left - .Width
        .Top = txtShveller_L.Top
    End With
    
    fraParametry.Height = txtShveller_L.Top + txtShveller_L.Height + 15
    
    With fraRaschot
        .Top = fraKnopki.Top
        .Left = fraKnopki.Left + fraKnopki.Width + 6
        .Height = fraParametry.Top + fraParametry.Height - .Top
    End With

    With cmdUstanovit
        .Top = fraRaschot.Height - .Height - 15
        .Left = lstLentOtrezOperations.Left
    End With
    
    fraMaterial.Width = fraRaschot.Left + fraRaschot.Width - fraMaterial.Left
    
    With Me
        .Width = fraRaschot.Left + fraRaschot.Width + 18
        .Height = fraRaschot.Top + fraRaschot.Height + 36
    End With
    

End Sub

Private Sub VyravnitKartinki()

    Dim i As Integer, ctrl As Control, img As MSForms.Image
    For i = 0 To mltLentOtrez.Pages.Count - 1
    
        For Each ctrl In mltLentOtrez.Pages(i).Controls
        
            If TypeName(ctrl) = "Label" Then
                If mltLentOtrez.Pages(i).Name <> "pShveller" Then
                    ctrl.Left = 132
                Else
                    If ctrl.Name = "lblLentOtrez_Shveller" Then ctrl.Left = 147
                End If
            ElseIf TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
            
                ctrl.Left = 147
            
            ElseIf TypeName(ctrl) = "Image" Then
                Set img = ctrl
                With img
                    .Left = 6
                    .Width = 120
                    .Height = 120
                    .Top = 6
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = &H8000000A
                    .PictureSizeMode = fmPictureSizeModeClip
                    .BackColor = vbWhite
                End With
            End If
            
        Next
        
    Next
End Sub

Public Sub GlavnyyRaschot()

    Const K_TIP_PROIZV As Double = 1.3
    
    Material = OpredelitMaterial
    
    OtobrazitKnopkuUstanovki
    OchistitParametry
    IzmenitTsvetFormy Material
    
    Set Figura = SozdatFiguru
    
    If Not Figura Is Nothing Then
    
        Dim Marshrutator As New clsLentOtrez_Raschot
        Dim Operations As Variant
        
        Operations = Marshrutator.RaschotMarshrutaFigury(Figura, chkPravka.Value)
        
        Massa = Round(Figura.Massa, 3)
        txtMassa.text = CStr(Massa)

        Call ZapolnitListBoksIzMassiva(Operations, lstLentOtrezOperations)
        
    End If

    
End Sub

Private Sub OtobrazitKnopkuUstanovki()
    
    UstanovTokar = False
    UstanovFrez = False
    cmdUstanovit.Visible = True
    
    Select Case mltLentOtrez.SelectedItem.Name
        Case "pKrug", "pShestigran", "pTruba"
            UstanovTokar = True
            cmdUstanovit.Caption = "Расчет токарной операции"
        Case "pUgolnik", "pShveller", "pShveller2"
            UstanovFrez = True
            cmdUstanovit.Caption = "Фрезеровать торцы"
        Case Else
            cmdUstanovit.Caption = "-"
            cmdUstanovit.Visible = False
    End Select
    
End Sub

Private Sub OchistitParametry()

    txtMassa.text = "-"
    lstLentOtrezOperations.Clear
    
    D1 = 0
    d2 = 0
    TocheniyeNaUdar = False
    Shirina = 0
    Vysota = 0
    Set Figura = Nothing
    
End Sub

Private Function OpredelitMaterial() As EnumMaterialy

    If optAlyuminiy Then
        OpredelitMaterial = ALUMINIYEVYYE_SPLAVY
    ElseIf optLegirovannayaStal Then
        OpredelitMaterial = STAL_LEGIROVANNAYA
    ElseIf optMednyyeSplavy Then
        OpredelitMaterial = MEDNYYE_SPLAVY
    ElseIf optNerzaveyka Then
        OpredelitMaterial = STAL_NERZHAVEYUSHCHAYA
    ElseIf optStal Then
        OpredelitMaterial = STAL_UGLERODISTAYA
    End If

End Function

Private Function SozdatFiguru() As ITechnologicalShape

    Dim DlinaZagotovki As Double
    DlinaZagotovki = DblFromCtrl(txtShveller_L)
    If DlinaZagotovki = 0 Then Exit Function
    
    Dim Figura As ITechnologicalShape
    
    Select Case mltLentOtrez.SelectedItem.Name
        Case "pKrug"
            D1 = DblFromCtrl(txtKrug_D)
            d2 = 0
            TocheniyeNaUdar = False
            
            Dim Krug As New clsFigura_Krug
            Krug.Init Material, D1, DlinaZagotovki
            Set Figura = Krug
            
        Case "pPramougol"
            Dim List As New clsFigura_List
            List.Init Material, DblFromCtrl(txtPramougol_A), DblFromCtrl(txtPramougol_B), DlinaZagotovki
            Set Figura = List
            
        Case "pShestigran"
            Dim Shestigrnnik As New clsFigura_Shestigrnnik
            Shestigrnnik.Init Material, DblFromCtrl(txtShestigran_S), DlinaZagotovki
            Set Figura = Shestigrnnik
            
            D1 = Shestigrnnik.DiametrOpisannoyOkruzhnosti
            d2 = 0
            TocheniyeNaUdar = True
            
        Case "pPryamougolTruba"
            Dim PryamougolnayaTruba As New clsFigura_PryamougolnayaTruba
            PryamougolnayaTruba.Init Material, DblFromCtrl(txtPryamougolTruba_A), DblFromCtrl(txtPryamougolTruba_B), DblFromCtrl(txtPryamougolTruba_s), DlinaZagotovki
            Set Figura = PryamougolnayaTruba
            
        Case "pTruba"
            D1 = DblFromCtrl(txtTruba_D)
            d2 = DblFromCtrl(txtTruba_D) - 2 * DblFromCtrl(txtTruba_s)
            TocheniyeNaUdar = False
            
            Dim Truba As New clsFigura_Truba
            Truba.Init Material, D1, DblFromCtrl(txtTruba_s), DlinaZagotovki
            Set Figura = Truba

        Case "pUgolnik"
            Shirina = DblFromCtrl(txtUgolnik_Shirina)
            Vysota = DblFromCtrl(txtUgolnik_Vysota)
            
            Dim Ugolnik As New clsFigura_Ugolnik
            Ugolnik.Init Material, Shirina, Vysota, DblFromCtrl(txtUgolnik_a), DblFromCtrl(txtUgolnik_b), DlinaZagotovki
            Set Figura = Ugolnik
            
        Case "pShveller2"
            Shirina = DblFromCtrl(txtShvellerP_Shirina)
            Vysota = DblFromCtrl(txtShvellerP_Vysota)
            
            Dim ShvellerP As New clsFigura_Shveller_P
            ShvellerP.Init Material, Shirina, Vysota, DblFromCtrl(txtShvellerP_TolshchinaStenki), DlinaZagotovki
            Set Figura = ShvellerP

        Case "pShveller"
            Dim ShvellerU As New clsFigura_Shveller_U
            ShvellerU.Init Material, cboShveller_Nomer.text, DlinaZagotovki
            Set Figura = ShvellerU
            
            Shirina = ShvellerU.Shirina
            Vysota = ShvellerU.Vysota
    End Select
    
    Set SozdatFiguru = Figura
    
End Function

Private Sub IzmenitTsvetFormy(Material As EnumMaterialy)

    Dim MaterialColor As Double
    MaterialColor = TsvetMateriala(Material)
    
    lblOperatsiya.BackColor = MaterialColor
    lblTpz.BackColor = MaterialColor
    lblTsht.BackColor = MaterialColor
    Me.Repaint

End Sub

Public Sub RaschotTokarnoyOperatsii()

    Dim Tokarnaya As New clsLentOtrez_TokPOp
    
    With Tokarnaya
        .Init ListIndexMaterial, Massa, D1, d2, Figura.Dlina, TocheniyeNaUdar
        .RaschotTokarnoyOperatsii
    End With
    
    Call RaschotGalvaniki
    
End Sub

Private Function ListIndexMaterial()

    Dim i As Integer
    With frmTokFrez.cboMaterial
        For i = 0 To .ListCount - 1
            If CInt(.List(i, 1)) = Material Then
                ListIndexMaterial = i
                Exit For
            End If
        Next
    End With
    
End Function

Public Sub FrezerovatTorcy()

    If Shirina > 0 And Vysota > 0 Then
            
        With frmTokFrez
            
            If (.tabOperatsii.tabs.Count = 1 And .lstTokarnoFrezer.ListCount > 0) Or .tabOperatsii.tabs.Count > 1 Then Call frmTokFrez.Raschot.DobavitOperatsiyu
            
            .cboMaterial.ListIndex = ListIndexMaterial
            
            .tabForMlt.Value = 1
            .mltNormativRastFrez.Value = .mltNormativRastFrez.Pages("pUstRastFrez").Index
            
            .txtMass_RastFrez.text = CStr(Massa)
            .txtKolVoProhodov_RastFrez.text = 1
            .lstSposobUst_RastFrez.ListIndex = 2
            .cboKharakterVyverki_RastFrez.ListIndex = 1
        
            Call frmTokFrez.DobavitVTablitsu
        
            .mltNormativRastFrez.Value = .mltNormativRastFrez.Pages("pFrezPloskKonc").Index
            
            .chkPerimetr_FrezPloskKonc.Value = False
            .txtL_FrezPloskKonc = CStr(WorksheetFunction.Max(Shirina, Vysota))
            .txtB_FrezPloskKonc = CStr(WorksheetFunction.Min(Shirina, Vysota))
            .cboIT_FrezPloskKonc = 11
            .txtPripusk_FrezPloskKonc = 5
            .txtKolVo_FrezPloskKonc = 2
            .txtPerekhod = "Фрезеровать торцы"
        
            Call frmTokFrez.DobavitVTablitsu
            
            'Промывка
            .tabPromyvka.Value = 0
            .txtPromivka_D.text = CStr(WorksheetFunction.Max(Shirina, Vysota))
            .txtPromivka_L.text = CStr(Figura.Dlina)
            
            .Show
            
        End With
        
        Call RaschotGalvaniki
        
    End If
    
End Sub

Public Sub OtkrytVkladku(img As MSForms.Image)
    
    Dim ctrl As Control
    For Each ctrl In fraKnopki.Controls
        If ctrl.Name = img.Name Then
            ctrl.BorderColor = &H8000000D
        Else
            ctrl.BorderColor = &H8000000A
        End If
    Next
    
    Dim i As Integer
    For i = 0 To mltLentOtrez.Pages.Count - 1
        If mltLentOtrez.Pages(i).Name = Replace(img.Name, "img_", "") Then
            mltLentOtrez.Value = i
            Exit For
        End If
    Next

End Sub

