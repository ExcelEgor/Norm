VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmObrabotkaOtverstiy 
   Caption         =   "Слесарная обработка. Обработка отверстий"
   ClientHeight    =   11280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10710
   OleObjectBlob   =   "frmObrabotkaOtverstiy.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmObrabotkaOtverstiy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Handler As clsFrmObrabotkaOtverstiy
Private ControlHandlers As Collection

Private Elementy()
Private iElement As Long
Private Stroka As Integer

Public KolVoStrok As Integer

Private Sub cboMaterial_Change()
    Call IzmenitTsvetFormy
End Sub

Private Sub chktObshchayaTolshchina_Click()
    txttObshchayaTolshchina.Enabled = chktObshchayaTolshchina.Value
    txttObshchayaTolshchina.text = Empty
End Sub

Private Sub chkUstanovka_Click()
    txtMassa.Enabled = chkUstanovka.Value
    txtTust.text = Empty
    If chkUstanovka.Value = False Then txtMassa.text = Empty
End Sub

Private Sub cmdClear_Click()
    Call OtkrytFormyVTomzheMeste(Me, Me.Top, Me.Left, Me.Name)
    
End Sub

Private Sub UserForm_Initialize()
    
    KolVoStrok = 20

    Call ZapolnitSpiski

    Call SozdatTablitsuObrabotkaOtverstiy

    Call VyravnitItog
    
    Call ZagruzkaElementovVKlass
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
End Sub

Private Sub SozdatTablitsuObrabotkaOtverstiy()


    For Stroka = 2 To KolVoStrok
        
        iElement = 0
        
        SozdatElement cboPerekhod_1
        SozdatElement txtKolVoOtv_1
        SozdatElement txtDotv_1
        SozdatElement txtLotv_1
        SozdatElement txtLrezb_1
        SozdatElement cboITotv_1
        SozdatElement chkGlukhoyeOtv_1
        SozdatElement lblGlukhoyeBorder_1
        SozdatElement txtVremya_1
        SozdatElement lblOchistit_1
    
        VyravnitElementyTablitsy Elementy, Stroka, Me

    Next
        

End Sub

Private Sub SozdatElement(ctrl As Control)

    Dim newCtrl As Control
    Set newCtrl = Me.Controls.Add("Forms." & TypeName(ctrl) & ".1", Split(ctrl.Name, "_")(0) & "_" & Stroka)
    
    ReDim Preserve Elementy(iElement)
    Set Elementy(iElement) = newCtrl
    iElement = iElement + 1
    
End Sub

Private Sub VyravnitItog()

    Dim PosledniyCtrlVremya As Control
    Set PosledniyCtrlVremya = Me.Controls("txtVremya_" & KolVoStrok)
    
    txtItogo.Left = PosledniyCtrlVremya.Left
    txtItogo.Top = PosledniyCtrlVremya.Top + PosledniyCtrlVremya.Height + 3
    
    lblItogo.Left = txtItogo.Left - lblItogo.Width - 3
    lblItogo.Top = txtItogo.Top + txtItogo.Height / 9
    
    txtTkontrol.Left = txtItogo.Left
    txtTkontrol.Top = txtItogo.Top + txtItogo.Height + 3
      
    lblTkontrol.Left = lblItogo.Left - (lblTkontrol.Width - lblItogo.Width)
    lblTkontrol.Top = txtTkontrol.Top + txtTkontrol.Height / 9
    
    Me.Height = txtTkontrol.Top + txtItogo.Height + 33
    
End Sub

Private Sub ZapolnitSpiski()

    Dim MATERIALY As Variant
    MATERIALY = Array(EnumMaterialy.ALUMINIYEVYYE_SPLAVY, EnumMaterialy.MEDNYYE_SPLAVY, EnumMaterialy.STAL_UGLERODISTAYA, _
        EnumMaterialy.STAL_LEGIROVANNAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA)
   
    Call DobavitMaterialyVListBox(MATERIALY, cboMaterial)
    
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
        
            Set Handler = New clsFrmObrabotkaOtverstiy

            If ctrl.Name Like "lblOchistit_*" Then
                Set Handler.LabelControl = ctrl
                
            ElseIf TypeOf ctrl Is MSForms.TextBox Then
                If Not ctrl.Name Like "txtVremya_*" Then
                    Set Handler.TextBoxControl = ctrl
                End If
                
            ElseIf TypeOf ctrl Is MSForms.ComboBox Then
                If ctrl.Name Like "cboPerekhod_*" Then
                    Set Handler.ComboBoxPerekhodControl = ctrl
                Else
                    Set Handler.ComboBoxControl = ctrl
                End If
                
            ElseIf TypeOf ctrl Is MSForms.CheckBox Then
                Set Handler.CheckBoxControl = ctrl
            End If

            ControlHandlers.Add Handler
            
        End If
    Next

End Sub

Private Sub IzmenitTsvetFormy()
    
    Dim ctrl As Control
    Dim Tsvet As Double
    Tsvet = TsvetMateriala(MaterialFromCboMaterial(cboMaterial))
    For Each ctrl In Me.Controls
        If ctrl.Tag = "color" Or ctrl.Name = "txtItogo" Or ctrl.Name = "txtTkontrol" Then
            ctrl.BackColor = Tsvet
        End If
    Next

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set Handler = Nothing
End Sub
