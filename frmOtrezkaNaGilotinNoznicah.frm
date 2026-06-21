VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtrezkaNaGilotinNoznicah 
   Caption         =   "Разметка и отрезка заготовок на гильотинных ножницах"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10065
   OleObjectBlob   =   "frmOtrezkaNaGilotinNoznicah.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtrezkaNaGilotinNoznicah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Material As EnumMaterialy

Private Sub UserForm_Initialize()

    Dim MATERIALY
    MATERIALY = Array(EnumMaterialy.ALUMINIYEVYYE_SPLAVY, EnumMaterialy.MEDNYYE_SPLAVY, EnumMaterialy.STAL_LEGIROVANNAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA, _
        EnumMaterialy.STAL_UGLERODISTAYA, EnumMaterialy.STEKLOTEKSTOLIT, EnumMaterialy.GETINAKS)
    
    Call DobavitMaterialyVListBox(MATERIALY, cboMaterial)
    
    cboMaterial.ListIndex = 0

    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
End Sub

Private Sub cboMaterial_Change()
    Material = MaterialFromCboMaterial(cboMaterial)
    Select Case Material
        Case EnumMaterialy.STAL_LEGIROVANNAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA, EnumMaterialy.STAL_UGLERODISTAYA, _
            EnumMaterialy.MEDNYYE_SPLAVY, EnumMaterialy.ALUMINIYEVYYE_SPLAVY
            chkRihtovka.Enabled = True
        Case Else
            chkRihtovka.Enabled = False
            chkRihtovka.Value = False
    End Select
    Call RaschetOtrezkiNaGilotinnyhNoznicah
End Sub

Private Sub cboMaterial_enter()
    Call SdelatShriftZhirnimPriFocuse(Me, cboMaterial, True)
End Sub

Private Sub cboMaterial_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call SdelatShriftZhirnimPriFocuse(Me, cboMaterial, False)
End Sub

Private Sub chkRihtovka_Click()
    Call RaschetOtrezkiNaGilotinnyhNoznicah
    chkRihtovka.Font.Bold = chkRihtovka.Value
End Sub

Private Sub txtDlina_Change()
    Call BolsheMaksimalnogo(MAX_DLINA_GILOTIONA, txtDlina)
    Call RaschetOtrezkiNaGilotinnyhNoznicah
End Sub
Private Sub BolsheMaksimalnogo(MaksZnachenie, txt As MSForms.TextBox)
    
    Dim Znachenie As Double
    Znachenie = DblFromCtrl(txt)
    
    Dim lblMaksZnachenie As Control
    Set lblMaksZnachenie = Me.Controls(Replace(txt.Name, "txt", "lbl") & "_Lmax")
    
    If Znachenie > MaksZnachenie Then
        lblMaksZnachenie.ForeColor = vbRed
    Else
        lblMaksZnachenie.ForeColor = &HC0C0C0
    End If
    
End Sub


Private Sub txtDlina_Enter()
    Call SdelatShriftZhirnimPriFocuse(Me, txtDlina, True)
End Sub

Private Sub txtDlina_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call SdelatShriftZhirnimPriFocuse(Me, txtDlina, False)
End Sub

Private Sub txtShirina_Change()
    Call BolsheMaksimalnogo(MAX_SHIRINA_GILOTIONA, txtShirina)
    Call RaschetOtrezkiNaGilotinnyhNoznicah
End Sub

Private Sub txtShirina_Enter()
    Call SdelatShriftZhirnimPriFocuse(Me, txtShirina, True)
End Sub

Private Sub txtShirina_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call SdelatShriftZhirnimPriFocuse(Me, txtShirina, False)
End Sub

Private Sub txtTolschina_Change()
    Call BolsheMaksimalnogo(MAX_TOLSHCHINA_GILOTIONA, txtTolschina)
    Call RaschetOtrezkiNaGilotinnyhNoznicah
End Sub

Private Sub txtTolschina_Enter()
    Call SdelatShriftZhirnimPriFocuse(Me, txtTolschina, True)
End Sub

Private Sub txtTolschina_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call SdelatShriftZhirnimPriFocuse(Me, txtTolschina, False)
End Sub

Private Sub RaschetOtrezkiNaGilotinnyhNoznicah()
    
    On Error GoTo ErrorHandler
      
    lstOtrez.Clear
    txtMassa.text = Empty
    
    Dim Dlina As Double, Shirina As Double, Tolshchina As Double
    Dlina = WorksheetFunction.Max(DblFromCtrl(txtDlina), DblFromCtrl(txtShirina))
    Shirina = WorksheetFunction.Min(DblFromCtrl(txtDlina), DblFromCtrl(txtShirina))
    Tolshchina = DblFromCtrl(txtTolschina)
        
    If Dlina > MAX_DLINA_GILOTIONA Or Shirina > MAX_SHIRINA_GILOTIONA Or Tolshchina > MAX_TOLSHCHINA_GILOTIONA Then Exit Sub
            
    Dim List As New clsFigura_List
    With List
        .Init Material, WorksheetFunction.Min(DblFromCtrl(txtDlina), DblFromCtrl(txtShirina)), DblFromCtrl(txtTolschina), WorksheetFunction.Max(DblFromCtrl(txtDlina), DblFromCtrl(txtShirina))

        txtMassa.text = Format(.Massa, "0.0")

        'Расчет слесарной операции
        Dim tZaus As Double
        tZaus = .tZaus
       
        Dim tRihtovka As Double
        If chkRihtovka.Value = True Then
            tRihtovka = .tPravka
        End If
    
        Dim tSlesar As Double
        tSlesar = tZaus + tRihtovka
    
        Dim tPeskostruy As Double
        tPeskostruy = .tPeskostruy
        
        Call DobavleniyeStrokVListBox(lstOtrez, "Отрезная", 10, .tOtrezkiNaGilotinnykhNozhnitsakhPoRazmetke)
        Call DobavleniyeStrokVListBox(lstOtrez, "Контрольная", 0, IzmerShtangenCircul_CPL(Dlina, Shirina) * 1.3)
        Call DobavleniyeStrokVListBox(lstOtrez, "Маркирование (этикеткой)", 0, 3)
        Call DobavleniyeStrokVListBox(lstOtrez, "Упаковывание", 0, UpakovyvaniyePramougolnika(Dlina, Shirina, Tolshchina) * 1.3)
        Call DobavleniyeStrokVListBox(lstOtrez, "Контрольная", 0, VisualnyyKontrol(Shirina, Dlina) * 1.3)
        Call DobavleniyeStrokVListBox(lstOtrez, "Пескоструйная", 5, 1.3 * tPeskostruy)
        Call DobavleniyeStrokVListBox(lstOtrez, "Слесарная", 5, 1.3 * tSlesar)
        Call DobavleniyeStrokVListBox(lstOtrez, "Контрольная", 0, VisualnyyKontrol(Shirina, Dlina) * 1.3)
    
        Exit Sub
    
    End With
    
ErrorHandler:
    lstOtrez.Clear
    Debug.Print "Ошибка." & " Источник: " & Err.Source & " Описание: " & Err.Description

End Sub



