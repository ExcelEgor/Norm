Attribute VB_Name = "mTokFrez_Zapolneniye_Tokarnyye"
Option Explicit
Option Private Module

Sub ZapolnitTokFrez_TokarnyyeStanki()

    ZapolnitUstanovku
    ZapolnitTipProizvodstva
    ZapolnitSpisokKvalitetov
    ZapolnitGlubinuRezaniya
    ZapolnitKonus
    ZapolnitObrabotkuOtverstiy
    ZapolnitProtachivaniyeKanavok
    ZapolnitDovodku
    ZapolnitGaltel
    ZapolnitPolirovaniye

End Sub
    
 Private Sub ZapolnitUstanovku()
    With frmTokFrez.lstSposobUst_Tok
        .AddItem "В самоцентрирующем патроне"
        .AddItem "В четырёхкулачковом патроне (детали цилиндрической формы)"
        .AddItem "В четырёхкулачковом патроне (детали фасонной и коробчатой формы)"
    End With
End Sub

Private Sub ZapolnitTipProizvodstva()
    With frmTokFrez.cboTipProizv
        .List = Split("Единичное Мелкосерийное Серийное")
        .ListIndex = 0
    End With
End Sub

Private Sub ZapolnitSpisokKvalitetov()

    With frmTokFrez
        Dim ArrCboIT(1 To 5)
        Set ArrCboIT(1) = .cboIT_Prod
        Set ArrCboIT(2) = .cboIT_Pop
        Set ArrCboIT(3) = .cboIT_Rast
        Set ArrCboIT(4) = .cboIT_ObrabOtvSNulya
        Set ArrCboIT(5) = .cboObshchIT_Tok
    End With
    
    Dim i As Integer
    Dim SpisokKvalitetov As String
    SpisokKvalitetov = "6 7 8 9 11 12 14"
    For i = LBound(ArrCboIT) To UBound(ArrCboIT)
        ArrCboIT(i).List = Split(SpisokKvalitetov)
        ArrCboIT(i).ListIndex = ArrCboIT(i).ListCount - 3
    Next
    
    frmTokFrez.cboObshchIT_Tok.ListIndex = -1
    frmTokFrez.cboObshchRaRz_Tok.ListIndex = -1
    
End Sub

Private Sub ZapolnitGlubinuRezaniya()

    With frmTokFrez
    
        Dim ArrCboGlubinaRez(1 To 2)
        Set ArrCboGlubinaRez(1) = .cboGlubinaRez_Rast
        Set ArrCboGlubinaRez(2) = .cboGlubinaRez_ObrabOtvSNulya
        
        Dim GlubinaRezPriRastachivanii As String
        GlubinaRezPriRastachivanii = "2 3 5"
        
        Dim i As Integer
        For i = LBound(ArrCboGlubinaRez) To UBound(ArrCboGlubinaRez)
            ArrCboGlubinaRez(i).List = Split(GlubinaRezPriRastachivanii)
            ArrCboGlubinaRez(i).ListIndex = 0
        Next
    
        .cboGlubinaRez_Prod.List = Split("2 4 6")
        .cboGlubinaRez_Prod.ListIndex = 0
        
        .cboGlubinaRez_Pop.List = Split("1 2 3 4")
        .cboGlubinaRez_Pop.ListIndex = 1
        
    End With
    
End Sub

Private Sub ZapolnitKonus()

    With frmTokFrez
        With .cboIT_Konus
            .List = Split("11 9 7")
            .ListIndex = 0
        End With
        With .cboRaRz_Konus
            .AddItem "5...2,5"
            .AddItem "1,25"
            .AddItem "0,63"
            .ListIndex = 0
        End With
    End With
        
End Sub
        
Private Sub ZapolnitObrabotkuOtverstiy()

    With frmTokFrez.cboPerehod_Sverl
        .AddItem
        .List(.ListCount - 1, 0) = "Сверление"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya_Tokar.Sverleniye
        
        .AddItem
        .List(.ListCount - 1, 0) = "Рассверливание"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya_Tokar.Rassverlivaniye
        
        .AddItem
        .List(.ListCount - 1, 0) = "Зенкерование"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya_Tokar.Zenkerovaniye
        
        .AddItem
        .List(.ListCount - 1, 0) = "Развертывание H9"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya_Tokar.Razvertyvaniye_Ra2_5
        
        .AddItem
        .List(.ListCount - 1, 0) = "Развертывание H7"
        .List(.ListCount - 1, 1) = TipObrabotkiOtverstiya_Tokar.Razvertyvaniye_Ra1_25

        .ListIndex = 0
    End With
        
End Sub
        
Private Sub ZapolnitProtachivaniyeKanavok()
    With frmTokFrez.cboIT_Kanav
        .List = Split("14 11 9")
        .ListIndex = 1
    End With
End Sub

Private Sub ZapolnitDovodku()

    With frmTokFrez
        With .cboTipPov_Dovod
            .AddItem "Наружная"
            .AddItem "Внутренняя"
            .AddItem "Коническая"
            .ListIndex = 0
        End With
        With .cboRaRz_Dovod
            .List = Split("0,63 0,32 0,16 0,08 0,04")
            .ListIndex = 0
        End With
        With .cboOtkl_Dovod
            .List = Split("0,001 0,002 0,005 0,010 0,050")
            .ListIndex = .ListCount - 1
        End With
    End With
    
End Sub
        
Private Sub ZapolnitGaltel()

    With frmTokFrez.cboRaRz_Galtel
        .AddItem
        .List(.ListCount - 1, 0) = "10...2,5"
        .List(.ListCount - 1, 1) = "2,5"
        
        .AddItem
        .List(.ListCount - 1, 0) = "1,25"
        .List(.ListCount - 1, 1) = "1,25"
        
        .AddItem
        .List(.ListCount - 1, 0) = "0,63"
        .List(.ListCount - 1, 1) = "0,63"
        
        
        .ListIndex = 1
    End With
End Sub
        
Private Sub ZapolnitPolirovaniye()
    With frmTokFrez.cboRaRz_Polirov
        .AddItem "1,25"
        .AddItem "0,63"
        .AddItem "0,32"
        .ListIndex = 0
    End With
End Sub

Sub ZapolnitSpisokRaKanavki()

    With frmTokFrez.cboRaRz_Kanav
        .Clear
        If CInt(frmTokFrez.cboIT_Kanav) > 11 Then
            .AddItem
            .List(.ListCount - 1, 0) = "20"
            .List(.ListCount - 1, 1) = "20"
        Else
            .AddItem
            .List(.ListCount - 1, 0) = "10 - 2,5"
            .List(.ListCount - 1, 1) = "10"
            .AddItem
            .List(.ListCount - 1, 0) = "1,25"
            .List(.ListCount - 1, 1) = "1,25"
        End If
        .ListIndex = 0
    End With
    
End Sub

Public Sub ZapolnitRaRz_TokFrez(cboIT As Control, cboRaRz As Control)

    'Список шероховатостей в зависимости от квалитета
    If cboIT.ListIndex <> -1 Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("Вспомогательный")
        With cboRaRz
            .ListIndex = -1
            If cboIT = "-" Then
                .RowSource = ws.Range("A2:D3").Address(external:=True)
            Else
                Select Case CInt(cboIT)
                    Case Is <= 5
                        .RowSource = ws.Range("A10:D10").Address(external:=True)
                    Case Is <= 10
                        .RowSource = ws.Range("A13:D16").Address(external:=True)
                    Case Else
                        .RowSource = ws.Range("A3:D7").Address(external:=True)
                End Select
            End If
            .ListIndex = 0
        End With
    End If
    
End Sub

Public Sub ZapolnitShagRezby_TokFrez(txtD As MSForms.TextBox, cboS As MSForms.ComboBox)

    frmTokFrez.ZagruzkaFormy = True

    cboS.Clear
    If IsNumeric(txtD) Then
        If CDbl(txtD) <= 120 Then
            Dim tbMetrRezb_Shag As ListObject
            Set tbMetrRezb_Shag = ThisWorkbook.Worksheets("Ток_23, 24, 25, 26").ListObjects("tbMetrRezb_Shag")
            Dim i As Integer
            With tbMetrRezb_Shag
                For i = 1 To .ListRows.Count
                    If .DataBodyRange(i, 1) >= CDbl(txtD) Then
                        cboS.AddItem CStr(.DataBodyRange(i, 2))
                        If .DataBodyRange(i, 1) <> .DataBodyRange(i + 1, 1) Then Exit For
                    End If
                Next
            End With
            cboS.ListIndex = 0
        End If
    End If
    
    frmTokFrez.ZagruzkaFormy = False
    Call GlavnyyRaschot_TokFrez
    
End Sub
