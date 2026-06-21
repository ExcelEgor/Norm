VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpeedy 
   Caption         =   "ЛАЗЕРНАЯ РЕЗКА SPEEDY"
   ClientHeight    =   11040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19110
   OleObjectBlob   =   "frmSpeedy.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmSpeedy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const KOLVO_STROK As Integer = 5
Private Handler As clsFrmSpeedy
Private ControlHandlers As Collection

Private Sub cmdClear_Click(): Call OchistitRaschot: End Sub

Private Sub UserForm_Initialize()
    
    Call ZapolnitSpiski
    
    Call SozdatTablitsuRaschotaDopontilenoyVyrezki
    
    Call VyravnitElementyFormy

    Call ZagruzkaElementovVKlass

    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    Handler.GlavnyyRaschot

End Sub

Private Sub ZapolnitSpiski()

    Dim MATERIALY
    MATERIALY = Array(PARONIT, KARTON, FIBRA, REZINA, ORGSTEKLO)

    Call DobavitMaterialyVListBox(MATERIALY, cboMaterial)
    
    cboTip_1.List = Split("Периметр Отверстие Контур -")
    cboTip_1.ListIndex = cboTip_1.ListCount - 1
    
End Sub

Private Sub SozdatTablitsuRaschotaDopontilenoyVyrezki()

    Dim cboTip As MSForms.ComboBox
    Dim txtA As MSForms.TextBox, txtB As MSForms.TextBox, txtKontur As MSForms.TextBox, txtOtv As MSForms.TextBox, txtKolVo As MSForms.TextBox

    Dim i As Integer
    For i = 2 To KOLVO_STROK

        Set cboTip = fraDopolnitelnaya.Controls.Add("Forms.ComboBox.1", "cboTip_" & i)
        cboTip.Style = fmStyleDropDownList

        With fraDopolnitelnaya
            Set txtA = .Controls.Add("Forms.TextBox.1", "txtA_" & i)
            Set txtB = .Controls.Add("Forms.TextBox.1", "txtB_" & i)
            Set txtKontur = .Controls.Add("Forms.TextBox.1", "txtKontur_" & i)
            Set txtOtv = .Controls.Add("Forms.TextBox.1", "txtOtv_" & i)
            Set txtKolVo = .Controls.Add("Forms.TextBox.1", "txtKolvo_" & i)
        End With
        
        Dim Elementy
        Elementy = Array(cboTip, txtA, txtB, txtKontur, txtOtv, txtKolVo)
        
        Call VyravnitElementyTablitsy(Elementy, i, Me)

    Next

End Sub

Private Sub VyravnitElementyFormy()
    
    fraZagolovok.Left = 0
    fraZagolovok.Top = 0
    
    fraZagotovka.Left = 3
    fraZagotovka.Top = fraZagolovok.Top + fraZagolovok.Height + 3
    fraZagotovka.Width = 312
    
    With fraDopolnitelnaya
        .Left = fraZagotovka.Left + fraZagotovka.Width + 3
        .Top = fraZagotovka.Top
        .Width = fraZagotovka.Width
        .Height = 141
    End With
    
    With fraRaschot
        .Left = fraDopolnitelnaya.Left
        .Top = fraDopolnitelnaya.Top + fraDopolnitelnaya.Height + 3
        .Width = fraDopolnitelnaya.Width
    End With
    
    With fraOsnParametry
        .Left = fraZagotovka.Left
        .Top = fraZagotovka.Top + fraZagotovka.Height + 3
        .Width = fraZagotovka.Width
        .Height = fraRaschot.Top + fraRaschot.Height - fraOsnParametry.Top
    End With
    
    cmdClear.Top = fraOsnParametry.Top + fraOsnParametry.Height + 3
    
    lblTipVyrezKontura.Left = 3
    lblTipVyrezKontura.Top = 9
    
    With mltTipZagotovki
        .Top = lblTipVyrezKontura.Top + lblTipVyrezKontura.Height + 3
        .Left = 0
        .Width = fraOsnParametry.Width
        .Height = fraOsnParametry.Height - .Top
    End With

    
    fraZagolovok.Width = fraRaschot.Left + fraRaschot.Width
    lblZagolovov.Left = 0
    lblZagolovov.Width = fraZagolovok.Width
    
    cmdClear.Left = fraZagotovka.Left
    
    Me.Height = cmdClear.Top + cmdClear.Height + 30
    Me.Width = fraRaschot.Left + fraRaschot.Width + 6
    
    With mltTipZagotovki
        .Left = -1
        .Width = fraZagotovka.Width + 1
        .TabFixedWidth = fraZagotovka.Width / 3.05
    End With
    
End Sub

Private Sub ZagruzkaElementovVKlass()

    Dim ctrl As Control
    
    Set ControlHandlers = New Collection
    
    For Each ctrl In Me.Controls
    
        If LCase(ctrl.Tag) <> "notcls" Then
            
            Set Handler = New clsFrmSpeedy

            If TypeOf ctrl Is MSForms.TextBox Then
                Set Handler.TextBoxControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.ComboBox Then
                Set Handler.ComboBoxControl = ctrl
            ElseIf TypeOf ctrl Is MSForms.MultiPage Then
                Set Handler.MultiPageControl = ctrl
            End If
            
            ControlHandlers.Add Handler
            Handler.KolVoStrok = KOLVO_STROK
            Set Handler.Forma = Me
            
        End If
    Next
    
End Sub

Private Sub OchistitRaschot()

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.text = Empty
        ElseIf TypeName(ctrl) = "ComboBox" Then
            If ctrl.Name <> cboMaterial.Name Then
                ctrl.ListIndex = -1
            End If
        End If
    Next
    
End Sub














