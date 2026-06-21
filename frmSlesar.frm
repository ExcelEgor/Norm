VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSlesar 
   Caption         =   "Расчёт слесарных работ"
   ClientHeight    =   11985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   21990
   OleObjectBlob   =   "frmSlesar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSlesar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Handler As clsSlesarForm
Private ControlHandlers As Collection
Private RezhimRedaktirovaniya As Boolean

Const KOLVO_STROK_OBR_OTV As Integer = 30

Dim Operations As Collection

Private Sub cmdNovayaOperatsiya_Click(): Call NovayaOperatsiya: End Sub

Private Sub cmdDelete_Click(): Call UdalitPerkhod: End Sub

Private Sub cboMaterial_Change(): Call IzmeneniyeMateriala: End Sub

Private Sub UserForm_Initialize()

    Set Operations = New Collection
    
    Call ZapolnitElementyUpravleniya
    
    Call SozdatTablitsuObrabotkaOtverstiy
    
    Call VyravnivaniyeElementov
    
    Call NovayaOperatsiya
    
    Call ZagruzkaElementovVKlass

    Handler.RaschotTsht
    
End Sub

Private Sub ZapolnitElementyUpravleniya()

    Dim MATERIALY As Variant
    MATERIALY = Array(EnumMaterialy.ALUMINIYEVYYE_SPLAVY, EnumMaterialy.MEDNYYE_SPLAVY, EnumMaterialy.STAL_UGLERODISTAYA, _
        EnumMaterialy.STAL_LEGIROVANNAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA)
   
    Call DobavitMaterialyVListBox(MATERIALY, cboMaterial)
    
    With cboTochnost_PravkaList
        .AddItem "0,3"
        .AddItem "0,5"
        .AddItem "1,0"
        .AddItem "1,5"
        .ListIndex = 0
    End With
    
    With cboRa_Zachistka
        .AddItem "Rz40 - Rz20"
        .AddItem "Ra 2,5..1,25"
        .ListIndex = 0
    End With

    Call ZapolnitObrabotkuOtverstiy
    
    Call ZapolnitElementyUpravleniya_Pritirka

End Sub
Private Sub ZapolnitElementyUpravleniya_Pritirka()

    With cboOtkloneniye_PritirPlosk
        .List = Split("0,001 0,002 0,005 0,010")
        .ListIndex = 0
    End With
    
    With cboTolshchina_PritirPlosk
        .AddItem: .List(.ListCount - 1, 0) = "Свыше 5мм":   .List(.ListCount - 1, 1) = 6
        .AddItem: .List(.ListCount - 1, 0) = "До 5 мм":     .List(.ListCount - 1, 1) = 5
        .AddItem: .List(.ListCount - 1, 0) = "До 3 мм":     .List(.ListCount - 1, 1) = 3
        .AddItem: .List(.ListCount - 1, 0) = "До 1 мм":     .List(.ListCount - 1, 1) = 1
        .ListIndex = 0
    End With
    
    With cboRa_PritirPlosk
        .RowSource = wsVspomogatelnyy.Range("J3:M6").Address(external:=True)
        .ListIndex = 0
    End With
    

End Sub

Private Sub ZapolnitObrabotkuOtverstiy()
    
    With cboPerekhod_1
    
        .AddItem
        .List(.ListCount - 1, 0) = "Сверлить"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya.Sverleniye
        
        .AddItem
        .List(.ListCount - 1, 0) = "Развернуть"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya.Razvertyvaniye
        
        .AddItem
        .List(.ListCount - 1, 0) = "Нарезать резьбу"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya.NarezaniyeRezby
        
        .AddItem
        .List(.ListCount - 1, 0) = "Сверлить+нарезать резьбу"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya.SverleniyeRezba

        .AddItem
        .List(.ListCount - 1, 0) = "Сверлить+развернуть"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya.SverleniyeRazvertyvaniye

        .AddItem
        .List(.ListCount - 1, 0) = "Калибровать"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya.KalibrovaniyeRezby

        .AddItem
        .List(.ListCount - 1, 0) = "Пуклевать"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya.Puklevka

        .AddItem
        .List(.ListCount - 1, 0) = "Зенковать"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya.ZenkovaniyeFasok

    End With
    
End Sub


Private Sub ZagruzkaElementovVKlass()
    
    Set ControlHandlers = New Collection
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If LCase(ctrl.Tag) <> "notclsm" Then
            Set Handler = New clsSlesarForm
            If ctrl.Name Like "cmdUdalitStroku_*" Then
                Set Handler.CmdClearControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.TextBox Then
                If Not ctrl.Name Like "txtVremya_*" Then
                    Set Handler.TextBoxControl = ctrl
                End If
            ElseIf TypeOf ctrl Is MSForms.MultiPage Then
                Set Handler.MultiPageControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.ComboBox Then
                If ctrl.Name Like "cboPerekhod_*" Then
                    Set Handler.ComboBoxPerekhodControl = ctrl
                Else
                    Set Handler.ComboBoxControl = ctrl
                End If
            ElseIf TypeOf ctrl Is MSForms.OptionButton Then
                Set Handler.OptionButtonControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.CheckBox Then
                Set Handler.CheckBoxControl = ctrl
            End If
            
            Handler.Init KOLVO_STROK_OBR_OTV
            
            ControlHandlers.Add Handler
            
        End If
    Next
    
End Sub

Private Sub VyravnivaniyeElementov()

    mltKarta_ObrabotkaOtverstiy.Style = fmTabStyleNone

    With fraNazvaniye
        .Left = 6
        .Top = 3
    End With
    
    With lblNazvaniye_1
        .Left = 0
        .Width = fraNazvaniye.Width
        .Top = 0
    End With
    
    With lblNazvaniye_2
        .Top = lblNazvaniye_1.Height
        .Width = fraNazvaniye.Width
    End With
    
    With fraMain
        .Left = fraNazvaniye.Left
        .Width = fraNazvaniye.Width
        .Top = fraNazvaniye.Top + fraNazvaniye.Height + 3
    End With
    
    With fraPerekhod
        .Left = fraNazvaniye.Left
        .Width = fraNazvaniye.Width
        .Top = fraMain.Top + fraMain.Height + 3
    End With

    With fraCmdAdd
        .Left = fraNazvaniye.Left
        .Width = fraNazvaniye.Width
        .Top = fraPerekhod.Top + fraPerekhod.Height + 3
    End With
    
    With fraRaschot
        .Top = fraNazvaniye.Top
        .Height = fraCmdAdd.Top + fraCmdAdd.Height - fraRaschot.Top
        .Left = fraNazvaniye.Left + fraNazvaniye.Width + 6
    End With
    
    Call VyravnitElementy_GlavnyyeKnopki
    
    With mltSlesar
        .Top = 3
        .Left = 3
        .Width = fraMain.Width - mltSlesar.Left - 6
        .Height = fraMain.Height - mltSlesar.Top - 6
    End With
    
    Call VyravnitElementy_VoVkladkakh
    
End Sub


Private Sub VyravnitElementy_GlavnyyeKnopki()
    Dim ctrl As Control
    With fraGlavnyyeKnopki
        .Left = 0
        .Top = 0
        For Each ctrl In .Controls
            ctrl.Width = .Width / .Count - 1
        Next
        .SpecialEffect = fmSpecialEffectFlat
    End With
    cmdNew.Left = 0
    cmdOtkryt.Left = cmdNew.Left + cmdNew.Width + 1
    cmdSaveAsNew.Left = cmdOtkryt.Left + cmdOtkryt.Width + 1
    cmdSave.Left = cmdSaveAsNew.Left + cmdSaveAsNew.Width + 1
    
    With fraMaterial
        .Left = fraGlavnyyeKnopki.Left
        .Top = fraGlavnyyeKnopki.Top + fraGlavnyyeKnopki.Height + 1
        .Width = fraRaschot.Width - 3
    End With
    
    With fraKnopkiOperatsii
        .Left = fraGlavnyyeKnopki.Left + 3
        .Top = fraMaterial.Top + fraMaterial.Height + 3
    End With
    
    fraKnopkiRaschot.Left = fraRaschot.Left - fraRaschot.Width + fraKnopkiRaschot.Width - 6
    lstRaschot.Left = fraKnopkiOperatsii.Left
    lstRaschot.Width = fraKnopkiRaschot.Left - lstRaschot.Left - 1
    
    Dim ColumnWidths
    ColumnWidths = Array(0.07 * lstRaschot.Width, 0.58 * lstRaschot.Width, 0.17 * lstRaschot.Width, 0.16 * lstRaschot.Width)
    
    lblTabl_1.Width = ColumnWidths(0)
    lblTabl_2.Width = ColumnWidths(1)
    lblTabl_4.Width = ColumnWidths(2)

    lblTabl_1.Left = lstRaschot.Left
    lblTabl_2.Left = lblTabl_1.Left + lblTabl_1.Width
    lblTabl_4.Left = lblTabl_2.Left + lblTabl_2.Width
    
    lblTabl_5.Left = lblTabl_4.Left + lblTabl_4.Width
    lblTabl_5.Width = lstRaschot.Width - lblTabl_5.Left + lstRaschot.Left
    
    lstRaschot.ColumnWidths = CStr(ColumnWidths(0)) & ";" & _
        CStr(ColumnWidths(1)) & ";" & _
        CStr(ColumnWidths(2)) & ";" & _
        CStr(ColumnWidths(3))
    
End Sub

Private Sub VyravnitElementy_VoVkladkakh()

    Dim ctrl As Control

    Dim Left_lblZagolovok As Integer:   Left_lblZagolovok = 3
    Dim Top_lblZagolovok As Integer:    Top_lblZagolovok = 3
    Dim Width_lblZagolovok As Double:   Width_lblZagolovok = mltSlesar.Width - mltSlesar.TabFixedWidth - Left_lblZagolovok - 3
    Dim Height_lblZagolovok As Double:  Height_lblZagolovok = 18
    
    Dim Width_fraParametry: Width_fraParametry = Width_lblZagolovok - 9
    
    For Each ctrl In Me.Controls
        If ctrl.Name Like "lblZagolovok_*" Then
            ctrl.Left = Left_lblZagolovok
            ctrl.Top = Top_lblZagolovok
            ctrl.Width = Width_lblZagolovok
            ctrl.Height = Height_lblZagolovok
        ElseIf ctrl.Name Like "mltKarta_*" Then
            ctrl.Top = Top_lblZagolovok + Height_lblZagolovok + 3
            ctrl.Left = 3
            ctrl.Width = Width_lblZagolovok - 1.5
            ctrl.Height = mltSlesar.Height - ctrl.Top - 6
        ElseIf ctrl.Name Like "fraParametry_*" Then
            If ctrl.Name = "fraParametry_ObrabotkaOtverstiy" Then
                ctrl.Top = 0
                ctrl.Left = 0
                ctrl.Width = Width_lblZagolovok - 3
                ctrl.Height = mltSlesar.Height - mltSlesar.Top - ctrl.Top - 30
            Else
                ctrl.Top = 3
                ctrl.Left = 0
                ctrl.Width = Width_fraParametry
                ctrl.Height = mltSlesar.Height - 55
            End If
        ElseIf ctrl.Name Like "lblPodZagolovok_*" Then
            ctrl.Left = 0
            ctrl.Width = Width_fraParametry
        End If
    Next
    
End Sub

Private Sub cmdAdd_Click()
    If RezhimRedaktirovaniya Then
        Call ZapisatVListBox(lstRaschot.ListIndex)
    Else
        Call DobavitVListboxNovuyuZapis
    End If
End Sub

Private Sub cmdEdit_Click()
    If lstRaschot.ListIndex < 0 Then Exit Sub
    Call ZagruzitZnacheniyaElementov(lstRaschot.ListIndex)
End Sub

Private Sub cmdSaveAsNew_Click()
    
    Dim ElementyFormy As New clsElementyFormy

    Dim wsSave As Worksheet
    Workbooks.Add
    Set wsSave = ActiveWorkbook.Worksheets(1)
    
    Dim i As Long
    For i = 0 To lstRaschot.ListCount - 1
        Set ElementyFormy = Operations(i + 1)
        ElementyFormy.SaveToExcel wsSave
    Next

End Sub

Private Sub NovayaOperatsiya()

    With tabOperatsii
        .tabs.Add (Format(5 * (.tabs.Count + 1), "000"))
        .Value = .tabs.Count - 1
    End With

End Sub

Private Sub DobavitVListboxNovuyuZapis()
   
    If Not IsNumeric(txt_tSlesar) Then Exit Sub
    
    Dim Stroka As Integer
    lstRaschot.AddItem
    Stroka = lstRaschot.ListCount - 1
    
    Call ZapisatVListBox(Stroka)
    
    Dim ElementyFormy As New clsElementyFormy
    With ElementyFormy
        .Init Me
        .SaveFromForm
    End With

    Operations.Add ElementyFormy

End Sub

Private Sub ZapisatVListBox(Stroka As Integer)

    With lstRaschot
        .List(Stroka, 0) = .ListCount
        .List(Stroka, 1) = txtPerekhod
        .List(Stroka, 2) = txt_tSlesar.text
        .List(Stroka, 3) = "-"
    End With
    
    RezhimRedaktirovaniya = False
    
End Sub

Private Sub ZagruzitZnacheniyaElementov(i As Integer)

    RezhimRedaktirovaniya = True

    Dim ElementyFormy As New clsElementyFormy
    Set ElementyFormy = Operations(i + 1)
    ElementyFormy.LoadToForm
    
End Sub

Private Sub PereschotVsegoRaschota()
    
    If lstRaschot.ListIndex < 0 Or RezhimRedaktirovaniya Then Exit Sub

    Dim liMaterial As Integer
    liMaterial = cboMaterial.ListIndex

    Dim i As Integer
    For i = 0 To lstRaschot.ListCount - 1
        Call ZagruzitZnacheniyaElementov(i)
        cboMaterial.ListIndex = liMaterial
        Handler.RaschotTsht
        Call ZapisatVListBox(i)
    Next
    
End Sub

Private Sub UdalitPerkhod()
    
    If lstRaschot.ListIndex = -1 Then Exit Sub

    Dim selectedIndex As Integer
    selectedIndex = lstRaschot.ListIndex

    On Error Resume Next
    Operations.Remove (selectedIndex + 1)
    On Error GoTo 0

    lstRaschot.RemoveItem selectedIndex
    
End Sub

Private Sub IzmeneniyeMateriala()
    Call IzmenitTsvetFormy
    Call PereschotVsegoRaschota
End Sub

Private Sub IzmenitTsvetFormy()
    
    Dim Tsvet As Double
    Tsvet = TsvetMateriala(CInt(cboMaterial.List(cboMaterial.ListIndex, 1)))
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.Name Like "lblZagolovok_*" Or ctrl.Tag = "color" Or ctrl.Name Like "lblPodZagolovok_*" Then
            ctrl.BackColor = Tsvet
        End If
    Next

End Sub

Private Sub SozdatTablitsuObrabotkaOtverstiy()

    Dim cboPerekhod As MSForms.ComboBox
    Dim txtDotv As MSForms.TextBox
    Dim txtLotv As MSForms.TextBox
    Dim txtLrezb As MSForms.TextBox
    Dim txtKolVoOtv As MSForms.TextBox
    Dim cboITotv As MSForms.ComboBox
    Dim chkGlukhoyeOtv As MSForms.CheckBox
    Dim cmdUdalitStroku As MSForms.CommandButton
    Dim txtGlukhoye As MSForms.TextBox
    Dim txtVremya As MSForms.TextBox
    
    Dim Stroka As Integer
    For Stroka = 2 To KOLVO_STROK_OBR_OTV
        
        With fraParametry_ObrabotkaOtverstiy
        
            Set cboPerekhod = .Controls.Add("Forms.ComboBox.1", "cboPerekhod_" & Stroka)
            Set txtDotv = .Controls.Add("Forms.TextBox.1", "txtDotv_" & Stroka)
            Set txtLotv = .Controls.Add("Forms.TextBox.1", "txtLotv_" & Stroka)
            Set txtLrezb = .Controls.Add("Forms.TextBox.1", "txtLrezb_" & Stroka)
            Set txtKolVoOtv = .Controls.Add("Forms.TextBox.1", "txtKolVoOtv_" & Stroka)
            Set cboITotv = .Controls.Add("Forms.ComboBox.1", "cboITotv_" & Stroka)
            Set txtGlukhoye = .Controls.Add("Forms.TextBox.1", "txtGlukhoye_" & Stroka)
            Set chkGlukhoyeOtv = .Controls.Add("Forms.CheckBox.1", "chkGlukhoyeOtv_" & Stroka)
            Set cmdUdalitStroku = .Controls.Add("Forms.CommandButton.1", "cmdUdalitStroku_" & Stroka)
            Set txtVremya = .Controls.Add("Forms.TextBox.1", "txtVremya_" & Stroka)
            
        End With
        
        Dim Elementy
        Elementy = Array(cboPerekhod, txtDotv, txtLotv, txtLrezb, txtKolVoOtv, cboITotv, chkGlukhoyeOtv, cmdUdalitStroku, txtGlukhoye, txtVremya)
        
        Call VyravnitElementyTablitsy(Elementy, Stroka, Me)
        
    Next
    
    fraParametry_ObrabotkaOtverstiy.ScrollHeight = cboPerekhod.Top + cboPerekhod.Height + 3
    

End Sub









