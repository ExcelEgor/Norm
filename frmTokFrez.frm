VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTokFrez 
   Caption         =   "Нормирование"
   ClientHeight    =   14715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23760
   OleObjectBlob   =   "frmTokFrez.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmTokFrez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public KartinkaKakKnopka As clsKartinkiKakKnopki
Public KollektsiyaKartinokKakKnopok As Collection

#If VBA7 Then
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
#If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If
#Else
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Public StopEvents As Boolean

Public ZagruzkaFormy As Boolean

Public Parametry_TokarnoFrezer As clsTokFrez

Private VspomRaschoty As New clsTokFrez_VspomRaschoty
Private Spiski As New clsTokFrez_Spiski

Public tTokFrez As Double
Public tSles As Double
Public tKontrol As Double
Public TipStanka As Integer

Public PoRezhimam As Boolean
Public Material As EnumMaterialy
Public Tonkostennaya As Boolean
Public CHPU As Boolean
Public mltNormativ As MSForms.MultiPage

Public Raschot As clsRaschotVKnige
Public Kniga As clsKniga

Public RezhimRedaktirovaniya As Boolean

Public OpName As String                      'Обозначение операции в frmTokFrez.mltOperations

Const GWL_STYLE As Long = (-16)
Const WS_SYSMENU As Long = &H80000
Const WS_MINIMIZEBOX As Long = &H20000
Const WS_MAXIMIZEBOX As Long = &H10000
Const SW_SHOWMAXIMIZED As Long = 3
Const SW_SHOWMINIMIZED As Long = 2
Const SW_SHOWNORMAL As Long = 1

''''''''''''''''''''''''''''''ЗАПОЛНЕНИЕ RA ПРИ ТОЧЕНИИ''''''''''''''''''''''''''''''
Private Sub cboIT_Pop_Change(): Call ZapolnitRaRz_TokFrez(cboIT_Pop, cboRaRz_Pop): End Sub

Private Sub cboIT_Prod_Change(): Call ZapolnitRaRz_TokFrez(cboIT_Prod, cboRaRz_Prod): End Sub

Private Sub cboIT_Rast_Change(): Call ZapolnitRaRz_TokFrez(cboIT_Rast, cboRaRz_Rast): End Sub

Private Sub cboIT_ObrabOtvSNulya_Change(): Call ZapolnitRaRz_TokFrez(cboIT_ObrabOtvSNulya, cboRaRz_ObrabOtvSNulya): End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cboKharakterVyverki_Tok_Change(): Call Spiski.ZapolnittSpisok_TochnostVyverki_Tokarnyy: End Sub






Private Sub lstSposobUst_Tok_Change(): Call Spiski.ZapolnittSpisok_KharakterVyverki_Tokarnyy: End Sub

Private Sub cboMaterial_Change(): Call IzmeneniyeMaterialaTipaProizvodstvaTipaStanka: End Sub

Private Sub cboKharakterVyverki_RastFrez_Change(): Call Spiski.ZapolnittSpisok_TochnostVyverki_Frezernyy: End Sub

Private Sub lstSposobUst_RastFrez_Change(): Call Spiski.ZapolnittSpisok_KharakterVyverki_Frezernyy: End Sub

Private Sub cboIT_Kanav_Change(): Call ZapolnitSpisokRaKanavki: End Sub


Private Sub optNaruzhnaya_Kanavka_Click()
    imgKanavka_Naruzh.BorderColor = &H8000000D
    imgKanavka_Vnutr.BorderColor = &H808080
End Sub

Private Sub optVnutrennyaya_Kanavka_Click()
    imgKanavka_Vnutr.BorderColor = &H8000000D
    imgKanavka_Naruzh.BorderColor = &H808080
End Sub

Private Sub TextBox5_Change()

End Sub

''''''''''''''''''''''''''''''ЗАПОЛНЕНИЕ ШАГА РЕЗЬБЫ''''''''''''''''''''''''''''''
Private Sub txtD1_MetrRezb_Change()
    Call ZapolnitShagRezby_TokFrez(txtD1_MetrRezb, cboS_MetrRezb)
End Sub

Private Sub txtD1_FrezRezb_Change()
    cboS_FrezRezb.Clear
    If IsNumeric(txtD1_FrezRezb) Then
        If CDbl(txtD1_FrezRezb) <= DMAX_METCHIK Then
            Call ZapolnitShagRezby_TokFrez(txtD1_FrezRezb, cboS_FrezRezb)
        End If
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''РАБОТА С РАСЧЕТНОЙ ТАБЛИЦЕЙ''''''''''''''''''''''''''''''
Private Sub cmdCopy_Click():   Call Raschot.KopirovatStroki: End Sub

Private Sub cmdCopyOp_Click():  Call Raschot.KopirovatOperatsiyu: End Sub

Private Sub cmdDelete_Click(): Call Raschot.UdalitStroki: End Sub

Private Sub cmdDeletePage_Click():  Call Raschot.UdalitOperatsiyu: End Sub

Private Sub cmdDown_Click(): Call Raschot.PeremestitStrokuVniz: End Sub

Private Sub cmdEdit_Click(): Call Redaktirovaniye_TokFrez: End Sub

Private Sub cmdEditExit_Click(): Call VykhodIzRezhimaRedaktirovaniya: End Sub

Private Sub cmdNew_Click(): Call NovyyRaschot: End Sub

Private Sub cmdOtkryt_Click(): Call Raschot.OtkryRaschot(wsTablitsa_TokarnoFrezer.ListObjects(1)): End Sub

Private Sub cmdSave_Click(): Call Sokhranit: End Sub

Private Sub cmdSaveAsNew_Click(): Call SokhranitKak: End Sub

Private Sub cmdUp_Click(): Call Raschot.PeremestitStrokuVverkh: End Sub

Private Sub cmdAdd_Click(): Call DobavitVTablitsu: End Sub

Private Sub cmdNewOp_Click(): Call Raschot.DobavitOperatsiyu: End Sub

Private Sub tabOperatsii_Change():  Call PereklyucheniyeOperatsiy: End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub chkGlukhoye_Rast_Click(): chkPodrezka_Rast.Value = chkGlukhoye_Rast.Value: End Sub

Private Sub chkOkruglyonnyyeNormy_Click(): Call Raschot.ZapisatTsht: End Sub

Private Sub chkPoRezhimam_Click(): Call IzmeneniyeMaterialaTipaProizvodstvaTipaStanka: End Sub

Private Sub chkTonkosten_Click(): Call IzmeneniyeMaterialaTipaProizvodstvaTipaStanka: End Sub

Private Sub mltNormativRastFrez_Change(): Call GlavnyyRaschot_TokFrez: End Sub

Private Sub mltNormativTok_Change(): Call GlavnyyRaschot_TokFrez: End Sub

Private Sub mltMain_change(): Call OpredeleniyeVybrannogoNormativa: End Sub

Private Sub cmdRaschotTpz_Click(): Call Raschot.RaschotTpz: End Sub

Private Sub tabOperatsii_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then Call PereimenovatOperatsiyu
End Sub

''''''''''''''''''''''''''''''КАРТИНКИ/КНОПКИ ДЛЯ ОТКРЫТИЯХ ДРУГИХ ФОРМ''''''''''''''''''''''''''''''
Private Sub imgElektroeroziya_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): frmElektroEroziya.Show: End Sub

Private Sub imgGeometrRascheti_Click(): frmGeometrRascheti.Show: End Sub

Private Sub imgGidroabraziv_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): frmGidroabraziv.Show: End Sub

Private Sub imgKIM_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): frmKIM.Show: End Sub

Private Sub imgLentochnoOtreznieStanki_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): frmLentochnoOtreznieStanki.Show: End Sub

Private Sub imgMassa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): frmMass.Show: End Sub

Private Sub imgOkrashivaniye_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): frmOkrashivaniye.Show: End Sub

Private Sub imgSvarka_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): frmSvarka.Show: End Sub

Private Sub imgFrezGabarit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): frmFrezGabarit.Show: End Sub

Private Sub fraMainMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):   Call UbratRabmkuUKartinok: End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):  Call UbratRabmkuUKartinok: End Sub

Private Sub mltMain_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call UbratRabmkuUKartinok: End Sub

Private Sub mltOperations_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):    Call UbratRabmkuUKartinok: End Sub

Private Sub UbratRabmkuUKartinok()
    Dim Kartinka
    For Each Kartinka In KollektsiyaKartinokKakKnopok
        Kartinka.img.BorderStyle = fmBorderStyleNone
    Next
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''ВСПОМОГАТЕЛЬНЫЕ РАСЧЕТЫ''''''''''''''''''''''''''''''
Private Sub cboMaterial_ZausShestigran_Change():    Call VspomRaschoty.RashotZachistkiShestigrannika: End Sub

Private Sub txtL_ZausShestigran_Change():           Call VspomRaschoty.RashotZachistkiShestigrannika: End Sub

Private Sub txtS_ZausShestigran_Change():           Call VspomRaschoty.RashotZachistkiShestigrannika: End Sub

Private Sub txtS_Shestigrannik_Change(): txtDopis_Shestigrannik.text = CStr(WorksheetFunction.Round(DiametrShestigrannika(DblFromCtrl(txtS_Shestigrannik)), 3)): End Sub

Private Sub chkSDvukhStoron_Click():                Call VspomRaschoty.RashotZachistkiShestigrannika: End Sub

Private Sub cboKomplekt_Massa_Change():             Call VspomRaschoty.RashotKopmplektovaniya: End Sub

Private Sub txtKomplekt_KolVo_Change():             Call VspomRaschoty.RashotKopmplektovaniya: End Sub

Private Sub chkKomplekt_Transport_Click():          Call VspomRaschoty.RashotKopmplektovaniya: End Sub

Private Sub txtGalvanika_S_Change():                Call VspomRaschoty.RaschotGalvaniki: End Sub

Private Sub cboGalvanika_Change():                  Call VspomRaschoty.RaschotGalvaniki: End Sub

Private Sub chkGalvanikaPartiya_Click():           Call VspomRaschoty.RaschotGalvaniki: End Sub

Private Sub txtPromivka_D_Change():                 Call VspomRaschoty.RaschotPromyvkiObertyvaniya: End Sub

Private Sub txtPromivka_L_Change():                 Call VspomRaschoty.RaschotPromyvkiObertyvaniya: End Sub

Private Sub tabPromyvka_Change():                   Call VspomRaschoty.RaschotPromyvkiObertyvaniya: End Sub

Private Sub txtPerimetrA_FrezPloskKonc_Change():    Call VspomRaschoty.RaschotPerimetra_FrezPloskKonc: End Sub

Private Sub txtPerimetrB_FrezPloskKonc_Change():    Call VspomRaschoty.RaschotPerimetra_FrezPloskKonc: End Sub

Private Sub chkPerimetr_FrezPloskKonc_Click():      Call VspomRaschoty.FrezerovaniyePerimetra_FrezPloskKonc: End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''УСТАНОВКА ОБЩИХ ПАРАМЕТРОВ ТОЧЕНИЯ''''''''''''''''''''''''''''''
Private Sub cboObshchRaRzTok_Change(): VspomRaschoty.UstanovitObshchiyeParametry: End Sub

Private Sub cboObshchRaRz_Tok_Change(): Call VspomRaschoty.UstanovitObshchiyeParametry: End Sub

Private Sub cboObshchIT_Tok_Change()
    Call ZapolnitRaRz_TokFrez(cboObshchIT_Tok, cboObshchRaRz_Tok)
    Call VspomRaschoty.UstanovitObshchiyeParametry
End Sub

Private Sub chkObshchParam_Tok_Click()
    Call VspomRaschoty.VklucheniyeObshchikhParametrov
End Sub

''''''''''''''''''''''''''''''ЗАПОЛНЕНИЕ RA ПРИ ЦЕНТРОВАНИИ''''''''''''''''''''''''''''''
Private Sub optTsentrovaniye_1_Click(): ZapolnitRaTsentrovaniye (1): End Sub

Private Sub optTsentrovaniye_2_Click(): ZapolnitRaTsentrovaniye (2): End Sub

Private Sub optTsentrovaniye_3_Click(): ZapolnitRaTsentrovaniye (3): End Sub

Private Function ZapolnitRaTsentrovaniye(TipTsentrovaniya As Integer)
    
    Dim ArrRa
    Select Case TipTsentrovaniya
        Case 1: ArrRa = "-"
        Case 2: ArrRa = "2,5"
        Case 3: ArrRa = "1,25 0,63"
    End Select

    With cboRa_Tsentrovaniye
        .Enabled = (TipTsentrovaniya = 3)
        .Clear
        .List = Split(ArrRa, " ")
        .ListIndex = 0
    End With
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserForm_Initialize()

#If VBA7 Then
    Dim MyHwnd As LongPtr
    Dim MyStyle As LongPtr
#Else
    Dim MyHwnd As Long
    Dim MyStyle As Long
#End If
    MyHwnd = FindWindow(vbNullString, Me.Caption)
#If VBA7 Then
#If Win64 Then
    MyStyle = GetWindowLongPtr(MyHwnd, GWL_STYLE) Or WS_SYSMENU Or WS_MINIMIZEBOX
#Else
    MyStyle = GetWindowLong(MyHwnd, GWL_STYLE) Or WS_SYSMENU Or WS_MINIMIZEBOX
#End If
    SetWindowLongPtr MyHwnd, GWL_STYLE, (MyStyle)
#Else
    MyStyle = GetWindowLong(MyHwnd, GWL_STYLE) Or WS_SYSMENU Or WS_MINIMIZEBOX
    SetWindowLong MyHwnd, GWL_STYLE, (MyStyle)
#End If
    DrawMenuBar MyHwnd
    
    StopEvents = True
    ZagruzkaFormy = True
    
    Call ZagruzitFormu_TokFrez
    
    Call OpredeleniyeVybrannogoNormativa
        
    chkPoRezhimam.ControlTipText = "Точный расчёт времени по режимам резания и нормативам вспомогательного времени. " _
        & vbNewLine & _
        "Пока работает только для растачивания, продольного и попереченого точения."
    
    StopEvents = False
    ZagruzkaFormy = False
    
    tabForMlt.Value = 1
    tabForMlt.Value = 0
    
    optTsentrovaniye_1.Value = True
    fraObniz_FrezOkon.SpecialEffect = fmSpecialEffectFlat
    
    Call GlavnyyRaschot_TokFrez
        
End Sub

Private Sub chkObniz_FrezOkon_Click()
    Dim ctrl As Control
    For Each ctrl In fraObniz_FrezOkon.Controls
        If ctrl.Name <> "chkObniz_FrezOkon" Then
            ctrl.Enabled = chkObniz_FrezOkon.Value
            If TypeName(ctrl) = "TextBox" Then
                ctrl.BorderColor = IIf(chkObniz_FrezOkon, &H8000000D, &H8000000A)
                If chkObniz_FrezOkon.Value = False Then ctrl.text = Empty
            End If
        End If
    Next
End Sub

Private Sub chkZachistkaBokovykhStoron_Click()
    cboIT_FrezPazov.ListIndex = cboIT_FrezPazov.ListCount - 1
    cboRaRz_FrezPazov.ListIndex = 0
    
    cboIT_FrezPazov.Enabled = chkZachistkaBokovykhStoron
    cboRaRz_FrezPazov.Enabled = chkZachistkaBokovykhStoron
End Sub

Private Sub optSkvoznoye_Shestigrannik_Click()
    imgSkvoznoye_Shestigrannik.BorderStyle = fmBorderStyleSingle
    imgVUpor_Shestigrannik.BorderStyle = fmBorderStyleNone
End Sub

Private Sub optVUpor_Shestigrannik_Click()
    imgVUpor_Shestigrannik.BorderStyle = fmBorderStyleSingle
    imgSkvoznoye_Shestigrannik.BorderStyle = fmBorderStyleNone
End Sub

Private Sub tabForMlt_Change()

    If tabForMlt.Value = 2 Then
        mltMain.Value = 1
    Else
        mltMain.Value = tabForMlt.Value
    End If

    Call IzmeneniyeMaterialaTipaProizvodstvaTipaStanka
 
End Sub

Private Sub txtDopis_Shestigrannik_Change()
    txtD1_Shestigrannik.text = txtDopis_Shestigrannik.text
End Sub


Private Sub cboTipPov_Dovod_Change()
    optVUpor_Dovodka.Enabled = True
    optSkvoznoye_Dovodka.Enabled = True
    If cboTipPov_Dovod.ListIndex = 2 Then
        optSkvoznoye_Dovodka.Value = True
        optVUpor_Dovodka.Enabled = False
        optSkvoznoye_Dovodka.Enabled = False
    End If

End Sub

Private Sub cboTipProizv_Change()

    With lblKoefTipProizv
    
        Select Case cboTipProizv.ListIndex
            Case 0:     .Caption = "K=1,3"
            Case 1:     .Caption = "K=1,0"
            Case Else:  .Caption = "K=0,7"
        End Select
        
    End With
    
    Call IzmeneniyeMaterialaTipaProizvodstvaTipaStanka

End Sub

Private Sub chkTonkoSten_Konus_Click()

    StopEvents = True
    
    With cboRaRz_Konus
        .Clear
        .AddItem
        .List(.ListCount - 1, 0) = "5,0...2,5"
        .List(.ListCount - 1, 1) = 5
        
        .AddItem
        .List(.ListCount - 1, 0) = "1,25"
        .List(.ListCount - 1, 1) = "1,25"

        If chkTonkosten.Value = False Then
            .AddItem
            .List(.ListCount - 1, 0) = "0,63"
            .List(.ListCount - 1, 1) = "0,63"
        End If
        
        .ListIndex = 0
    
    End With
    
    StopEvents = False
    
End Sub

Private Sub txtTime_Change()
    If IsNumeric(txtTime) Then
        txtTime.ControlTipText = Format(txtTime / 60, "0.00") & " ч."
    Else
        txtTime.ControlTipText = ""
    End If
End Sub

Public Sub IzmenitRezhimFormy()

    Dim i As Integer, ctrl As Control
        
    With frmTokFrez
        
        If .RezhimRedaktirovaniya Then
            .cmdAdd.Caption = "Редактировать"
        Else
            .cmdAdd.Caption = "Добавить переход в расчёт"
        End If
        
        .cmdEditExit.Visible = .RezhimRedaktirovaniya

        For Each ctrl In .fraRaschot.Controls
            ctrl.Enabled = Not .RezhimRedaktirovaniya
        Next

        With .mltNormativ
            For i = 0 To .Pages.Count - 1
                If i <> .SelectedItem.Index Then .Pages(i).Enabled = Not frmTokFrez.RezhimRedaktirovaniya
            Next
        End With
        
    End With

End Sub

Public Sub IzmeneniyeTsvetaFormy()

    Dim ctrl As Control
    Dim matColor As Double
    
    For Each ctrl In frmTokFrez.Controls
        If ctrl.Tag = "color" Or ctrl.Name Like "lblHead*" Then
        
            Select Case frmTokFrez.cboMaterial.ListIndex + 1
                Case 1, 2                   'Алюминиевые, бронзовые и медные сплавы
                    matColor = &HDAEFE2     'Зеленый
                Case 3, 4                   'Углеродистые стали и легированные стали
                    matColor = &HF7EBDD     'Синий
                Case 5                      'Нержавеющие стали
                    matColor = &H80000018   'Жёлтый
                Case 6                      'Титановые сплавы
                    matColor = &HF2F2F2     'Серый
            
            End Select
            
            ctrl.BackColor = matColor
            
        End If
    Next
    
End Sub

Public Sub VykhodIzRezhimaRedaktirovaniya()

    frmTokFrez.RezhimRedaktirovaniya = False
    Call IzmenitRezhimFormy
    
End Sub

Private Sub PereimenovatOperatsiyu()
    
    With frmInputBox
        .Init frmTokFrez.Raschot.wsRaschot
        .Show
    End With
    
End Sub

Private Sub OpredeleniyeVybrannogoNormativa()

    Select Case frmTokFrez.mltMain.SelectedItem.Name
        Case "pTok"
            Set frmTokFrez.mltNormativ = frmTokFrez.mltNormativTok
            frmTokFrez.tabForMlt.Value = 0
        Case "pRastFrez"
            Set frmTokFrez.mltNormativ = frmTokFrez.mltNormativRastFrez
    End Select
    
    Call GlavnyyRaschot_TokFrez
    
End Sub

Public Sub OchistitKolVoPriRaboteIzPrutka()

    With frmTokFrez
        If .mltMain.SelectedItem.Name = "pTok" Then
            If .mltNormativTok.SelectedItem.Name = "pProdolnoye" Then
                .txtKolVo_Prod.text = Empty
            End If
        End If
    End With
    
End Sub

Public Sub DobavitVTablitsu()
        
    Call Raschot.DobavitVTablitsu(txtPerekhod.text, Parametry_TokarnoFrezer, tTokFrez, RezhimRedaktirovaniya)
    
    Call OchistitKolVoPriRaboteIzPrutka
    
    Call VykhodIzRezhimaRedaktirovaniya

End Sub

Public Sub SokhranitKak()
    Kniga.SokhranitKak
    Call ZapisatImyaSohkrannenogoFayla
End Sub

Public Sub Sokhranit()
    Kniga.Sokhranit
    Call ZapisatImyaSohkrannenogoFayla
End Sub

Private Sub ZapisatImyaSohkrannenogoFayla()
    frmTokFrez.txtFilePath.text = Kniga.wbRaschot.FullName
    frmTokFrez.txtFileDateTime.text = FileDateTime(Kniga.wbRaschot.FullName)
End Sub

Private Sub PereklyucheniyeOperatsiy()

    Raschot.PereklyucheniyeOperatsiy
    
    Call IzmeneniyeTsvetaFormy
    
End Sub

Private Sub IzmeneniyeMaterialaTipaProizvodstvaTipaStanka()
    
    If ZagruzkaFormy Then Exit Sub
    
    Raschot.ZapisatParametryOperatsiiVListRaschota

    Call IzmeneniyeTsvetaFormy
    
End Sub

Private Sub NovyyRaschot()
    
    Dim i As Integer
    
    If MsgBox("Ткущий расчёт будет удалён. Вы уверены что хотите начать новый расчёт?", vbYesNo + vbQuestion, "Новый расчёт") = vbYes Then
        For i = tabOperatsii.tabs.Count - 1 To 1 Step -1
            tabOperatsii.tabs.Remove (i)
        Next

        Call Raschot.SozdatNovyyRaschot(Kniga)
        
        txtFileDateTime.text = Empty
        txtFilePath.text = Empty

    End If

End Sub

Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If MsgBox("Вы уверены, что хотите закрыть форму?", vbQuestion + vbYesNo, "Нормирование") = vbNo Then
        Cancel = True
    Else
        Call ZakrytRaschot
    End If
    
End Sub

Private Sub ZakrytRaschot()

    If frmTokFrez.Raschot.wbRaschot.Path <> "" Then
        If MsgBox("Сохранить текущий расчёт перед закрытием?", vbQuestion + vbYesNo, "Нормирование") = vbYes Then
            Call Sokhranit
        End If
    End If
    
    Raschot.wbRaschot.Close (False)

End Sub



