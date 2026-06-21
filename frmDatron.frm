VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatron 
   Caption         =   "Datron"
   ClientHeight    =   11985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21990
   OleObjectBlob   =   "frmDatron.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDatron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public ParametryDatron As clsDatron
Public RezhimRedaktirovaniya As Boolean

Private Handler As clsDatron_Frm_Raschoty
Private ControlHandlers As Collection

Private Tablitsa As clsDatron_Frm_Tablitsa

Private Sub cboMaterial_Change(): Call Tablitsa.ZapistatMaterial: End Sub

Private Sub cmdAdd_Click(): Call Tablitsa.DobavitZapis(ParametryDatron, RezhimRedaktirovaniya): End Sub

Private Sub cmdDelete_Click(): Call Tablitsa.UdalitStroku: End Sub

Private Sub cmdEdit_Click(): Call Tablitsa.RedaktirovatPerekhod: End Sub

Private Sub cmdNew_Click(): Tablitsa.NovyyRaschot: End Sub

Private Sub cmdNewOp_Click()
    Tablitsa.NovayaOperatsiya
End Sub

Private Sub cmdPolirovaniye_Click(): frmPolirovaniye.Show: End Sub

Private Sub cmdGalvanika_Click()
    With frmGalvanika
        .cboGalvanika.ListIndex = 3
        .Show
    End With
End Sub

Private Sub cmdPromyvka_Click()
    With frmPromyvka
        .optPloskaya = True
        .Show
    End With
End Sub

Private Sub optGlukhoye_otv_Click(): Call IzmenenitTipaOtverstiya: End Sub

Private Sub optGlukhoyeSObnizh_otv_Click(): Call IzmenenitTipaOtverstiya: End Sub

Private Sub optSkvoznoye_otv_Click(): Call IzmenenitTipaOtverstiya: End Sub

Private Sub optSkvoznoyeSObnizh_otv_Click(): Call IzmenenitTipaOtverstiya: End Sub
    
Private Sub tabOperatsii_Change()
    Tablitsa.PerekluchitOperatsiyu
End Sub

Private Sub UserForm_Initialize()

    Set Tablitsa = New clsDatron_Frm_Tablitsa

    Call IzmenenitTipaOtverstiya
    
    Call ZapolnitSpiski
    
    Call VyravnivanitElementov
    
    Call VyravnitPoLevomuKrayuNazvaniyaVkladok(mltDatron)
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    Call ZagruzitElementyVKlass

End Sub


Private Sub IzmenenitTipaOtverstiya()
    
    imgSkvoznoye_otv.Picture = imgSkvoznoye_otv_chb.Picture
    imgGlukhoye_otv.Picture = imgGlukhoye_otv_chb.Picture
    imgSkvoznoyeSObnizh_otv.Picture = imgSkvoznoyeSObnizh_otv_chb.Picture
    imgGlukhoyeSObnizh_otv.Picture = imgGlukhoyeSObnizh_otv_chb.Picture
    
    lblDobnizh_otv.Enabled = False
    lblLobnizh_otv.Enabled = False
    
    With txtLobnizh_otv
        .Enabled = False
        .BorderColor = &H8000000A
        .text = ""
    End With
    With txtDobnizh_otv
        .Enabled = False
        .BorderColor = &H8000000A
        .text = ""
    End With

    If optSkvoznoye_otv Then
        imgSkvoznoye_otv.Picture = imgSkvoznoye_otv_sin.Picture
    ElseIf optGlukhoye_otv Then
        imgGlukhoye_otv.Picture = imgGlukhoye_otv_sin.Picture
    ElseIf optSkvoznoyeSObnizh_otv Then
        imgSkvoznoyeSObnizh_otv.Picture = imgSkvoznoyeSObnizh_otv_sin.Picture
    ElseIf optGlukhoyeSObnizh_otv Then
        imgGlukhoyeSObnizh_otv.Picture = imgGlukhoyeSObnizh_otv_sin.Picture
    End If
    
    If optSkvoznoyeSObnizh_otv Or optGlukhoyeSObnizh_otv Then
        lblDobnizh_otv.Enabled = True
        txtDobnizh_otv.Enabled = True
        txtDobnizh_otv.BorderColor = &H8000000D
        lblLobnizh_otv.Enabled = True
        txtLobnizh_otv.Enabled = True
        txtLobnizh_otv.BorderColor = &H8000000D
    End If

    Me.Repaint
    
End Sub

Private Sub ZapolnitSpiski()

    Call ZapolnitSpisokSposobovUstanovki
    Call ZapolnitSpisokMaterialov
    
End Sub

Private Sub ZapolnitSpisokSposobovUstanovki()

    With lstUstanovka
        .AddItem
        .List(.ListCount - 1, 0) = "На вакуумном столе"
        .List(.ListCount - 1, 1) = EnumUstanovkaDatron.VakuumnyyStol
        
        .AddItem
        .List(.ListCount - 1, 0) = "В тисках"
        .List(.ListCount - 1, 1) = EnumUstanovkaDatron.Tiski
        
        .AddItem
        .List(.ListCount - 1, 0) = "В специально приспособлении"
        .List(.ListCount - 1, 1) = EnumUstanovkaDatron.Prisposobleniye
        
        .AddItem
        .List(.ListCount - 1, 0) = "В патроне для фрезерования в 4-ой оси"
        .List(.ListCount - 1, 1) = EnumUstanovkaDatron.Patron4Os
        
        .ListIndex = 0
    End With
    
End Sub

Private Sub ZapolnitSpisokMaterialov()

    Dim MATERIALY
    MATERIALY = Array(ALUMINIYEVYYE_SPLAVY, POLIAMID, ORGSTEKLO)
    
    Call DobavitMaterialyVListBox(MATERIALY, cboMaterial)
    
    With cboMaterial
        .AddItem
        .List(.ListCount - 1, 0) = "Текстолит/Стеклотекстолит"
        .List(.ListCount - 1, 1) = TEKSTOLIT
        .ListIndex = 0
    End With
    
End Sub

Private Sub VyravnivanitElementov()

    Dim ctrlWidth As Double
    ctrlWidth = mltDatron.Width - mltDatron.TabFixedWidth - 6
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Select Case True
            Case ctrl.Name Like "fraMain_*"
                ctrl.Left = 3
                ctrl.Top = 21
                ctrl.Width = ctrlWidth
                ctrl.Height = mltDatron.Height - ctrl.Top - 3
            Case ctrl.Name Like "lblZagolovok_*"
                ctrl.Left = 3
                ctrl.Top = 3
                ctrl.Width = ctrlWidth
            Case ctrl.Name Like "img_*"
                ctrl.Top = 3
                ctrl.Left = ctrlWidth - ctrl.Width - 6
        End Select
    Next
    
    With lblZagolovok
        .Left = 0
        .Width = fraZagolovok.Width
        .TextAlign = fmTextAlignCenter
    End With
    
    With lblDatron
        .Left = 0
        .Width = fraZagolovok.Width
        .TextAlign = fmTextAlignCenter
    End With
    
End Sub

Private Sub ZagruzitElementyVKlass()
    
    Set ControlHandlers = New Collection
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If LCase(ctrl.Tag) <> "notclsm" Then
            Set Handler = New clsDatron_Frm_Raschoty
            
            If TypeOf ctrl Is MSForms.TextBox Then
                Set Handler.TextBoxControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.ComboBox Then
                Set Handler.ComboBoxControl = ctrl
            ElseIf TypeName(ctrl) = "CheckBox" Then
                Set Handler.CheckBoxControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.MultiPage Then
                Set Handler.MultiPageControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.OptionButton Then
                Set Handler.OptionButtonControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.ListBox Then
                Set Handler.ListBoxControl = ctrl
            End If
            
            ControlHandlers.Add Handler
            
        End If
    Next
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If MsgBox("Вы уверены, что хотите закрыть форму?", vbQuestion + vbYesNo, "Datron") = vbNo Then
        Cancel = True
    Else
        Tablitsa.ZakrytRaschot
    End If
End Sub

