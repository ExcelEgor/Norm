VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSvarka 
   Caption         =   "Аргоно-дуговая сварка"
   ClientHeight    =   10110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11910
   OleObjectBlob   =   "frmSvarka.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmSvarka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const KOLVO_STROK_SVARKA As Integer = 11
Private Handler As clsFrmSvarka
Private ControlHandlers As Collection

Private Sub cboMaterial_Change(): Call IzmenitNadpisGalvanika: End Sub

Private Sub cmdNew_Click(): Call NovyyRaschot: End Sub


Private Sub UserForm_Initialize()

    Call ZapolnitSpiski

    Call SozdatTablitsuSvarki
    
    Call VyravnitElementy
    
    Call ZagruzkaElementovVKlass
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    Handler.GlavnyyRaschot
       
End Sub

Private Sub ZapolnitSpiski()

    cboTipProizvodstva.List = Split("Eдиничное,Мелкосерийное,Серийное", ",")
    cboTipProizvodstva.ListIndex = 0
    
    Dim MATERIALY As Variant
    MATERIALY = Array(EnumMaterialy.ALUMINIYEVYYE_SPLAVY, EnumMaterialy.MEDNYYE_SPLAVY, _
        EnumMaterialy.STAL_UGLERODISTAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA, EnumMaterialy.TITANOVYYE_SPLAVY)
    
    Call DobavitMaterialyVListBox(MATERIALY, cboMaterial)
    
End Sub

Private Sub SozdatTablitsuSvarki()
    
    Dim txtNum As MSForms.TextBox
    Dim txtTolshchina As MSForms.TextBox
    Dim txtDlina As MSForms.TextBox
    Dim txtKolVo As MSForms.TextBox
    Dim txtVremya As MSForms.TextBox
    
    Dim Stroka As Integer
    For Stroka = 2 To KOLVO_STROK_SVARKA
        
        With fraSvarka
        
            Set txtNum = .Controls.Add("Forms.TextBox.1", "txtNum_" & Stroka)
            Set txtTolshchina = .Controls.Add("Forms.TextBox.1", "txtTolshchina_" & Stroka)
            Set txtDlina = .Controls.Add("Forms.TextBox.1", "txtDlina_" & Stroka)
            Set txtKolVo = .Controls.Add("Forms.TextBox.1", "txtKolVo_" & Stroka)
            Set txtVremya = .Controls.Add("Forms.TextBox.1", "txtVremya_" & Stroka)
            
        End With
        
        Dim Elementy
        Elementy = Array(txtNum, txtTolshchina, txtDlina, txtKolVo, txtVremya)
        
        Call VyravnitElementyTablitsy(Elementy, Stroka, Me)
        
    Next
End Sub

Private Sub VyravnitElementy()

    fraZagolovok.Left = -1
    fraZagolovok.Width = Me.Width + 1
    
    lblZagolovov.Left = 0
    lblZagolovov.Width = fraZagolovok.Width
    
    imgZagolovok.Left = lblZagolovov.Width / 5
    imgZagolovok.Top = 3
    
End Sub

Private Sub ZagruzkaElementovVKlass()

    Dim ctrl As Control
    
    Set ControlHandlers = New Collection
    
    For Each ctrl In Me.Controls
    
        If LCase(ctrl.Tag) <> "notclsm" And Not ctrl.Name Like "txtVremya_*" Then
            
            Set Handler = New clsFrmSvarka

            If TypeOf ctrl Is MSForms.TextBox Then
                Set Handler.TextBoxControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.ComboBox Then
                Set Handler.ComboBoxControl = ctrl
            End If
            
            ControlHandlers.Add Handler
            Handler.KolVoStrokSvarka = KOLVO_STROK_SVARKA
            
        End If
    Next
    
    
End Sub

Private Sub NovyyRaschot()

    Dim ctrl As Control
    
    For Each ctrl In fraSvarka.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.text = Empty
        End If
    Next
    
    For Each ctrl In fraParametry.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.text = Empty
        ElseIf TypeName(ctrl) = "ComboBox" Then
            ctrl.ListIndex = 0
        End If
    Next

    Handler.GlavnyyRaschot
    
End Sub

Private Sub IzmenitNadpisGalvanika()

    Dim Material As EnumMaterialy
    Material = CInt(cboMaterial.List(cboMaterial.ListIndex, 1))
    
    If Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Then
        lblPloschad.Caption = "Площадь травления, м2"
        lblGalvaniika.Caption = "Травление"
    Else
        lblPloschad.Caption = "Площадь обезжиривания, м2"
        lblGalvaniika.Caption = "Обезжиривание"
    End If
End Sub

