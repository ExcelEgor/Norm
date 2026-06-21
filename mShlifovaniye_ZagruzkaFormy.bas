Attribute VB_Name = "mShlifovaniye_ZagruzkaFormy"
Option Explicit

Option Private Module

Const dopusk As String = "- 0,01 0,005 0,002"
Const ShirinaKruga As String = "до 50;63;80"
Const Galtel As String = "Нет;С 1 стороны;С 2 сторон"

Sub ZagruzkaFormy_Shlifovaniye()

    Dim i As Integer
    Dim ctrl As Control
    
    frmShlifonaiye.tglIdotZagruzkaFormy.Value = True
    
    Call VyravnitZagolovki
    Call ZapolnitOsnovnyyeParametry

    Call ZagruzkaFormy_Naruzhnoye
    Call ZagruzkaFormy_Vnutrenneye
    Call ZagruzkaFormy_Ploskoye

    frmShlifonaiye.tglIdotZagruzkaFormy.Value = False
    
End Sub

Private Sub ZapolnitOsnovnyyeParametry()

    Dim MATERIALY
    MATERIALY = Array(STAL_LEGIROVANNAYA, STAL_NERZHAVEYUSHCHAYA, STAL_UGLERODISTAYA, TITANOVYYE_SPLAVY)
    
    Call DobavitMaterialyVListBox(MATERIALY, frmShlifonaiye.cboMaterial)
    
    With frmShlifonaiye
    
        .cboTipProizvodstva.List = Split("Единичный Мелкосерийный Среднесерийный Крупносерийный")
        .cboTipProizvodstva.ListIndex = 0

        With .cboVozrast
            .AddItem "до 10 лет"
            .AddItem "свыше 10 до 20 лет"
            .AddItem "свыше 20 лет"
            .ListIndex = 2
        End With
        
    End With

End Sub

Private Sub VyravnitZagolovki()

    Dim i As Integer, ctrl As Control
    With frmShlifonaiye
        For i = 0 To .mltShlifonaiye.Pages.Count - 1
            For Each ctrl In .mltShlifonaiye.Pages(i).Controls
                With ctrl
                    If .Name Like ("lblZagolovok_*") Then
                        .Left = 3
                        .Top = 3
                        .Width = frmShlifonaiye.mltShlifonaiye.Width - frmShlifonaiye.mltShlifonaiye.TabFixedWidth
                        .TextAlign = 2
                    End If
                End With
            Next
        Next
    End With
        
End Sub

Private Sub ZagruzkaFormy_Naruzhnoye()

    With frmShlifonaiye

        With .cboGaltel_Naruzhnoye
            .ListWidth = .Width & " pt"
            .List = Split(Galtel, ";")
            .ListIndex = 0
        End With
        
        .cboDupusk_Naruzhnoye.List = Split(dopusk)
        .cboDupusk_Naruzhnoye.ListIndex = 0
        
        Call SpisokKvalitetov(.cboIT_Naruzhnoye)
        
        With .cboShirinaKruga_Naruzhnoye
            .List = Split(ShirinaKruga, ";")
            .ListIndex = 0
        End With
        
    End With

End Sub

Private Sub ZagruzkaFormy_Vnutrenneye()

    With frmShlifonaiye

        .cboDopusk_Vnutrenneye.List = Split(dopusk)
        .cboDopusk_Vnutrenneye.ListIndex = 0
        
        With .cboIT_Vnutrenneye
            .ColumnCount = 2
            .ListWidth = .Width & " pt"
            .ColumnWidths = .Width & " pt;0 pt"
            
            .AddItem
            .List(.ListCount - 1, 0) = 9
            .List(.ListCount - 1, 1) = 9
            
            .AddItem
            .List(.ListCount - 1, 0) = "7, 8"
            .List(.ListCount - 1, 1) = 7
            
            .AddItem
            .List(.ListCount - 1, 0) = 6
            .List(.ListCount - 1, 1) = 6
            
            .ListIndex = 2
        End With
        
        With .cboShirinaKruga_Vnutrenneye
            .List = Split(ShirinaKruga, ";")
            .ListIndex = 0
        End With
        
    End With

End Sub

Private Sub ZagruzkaFormy_Ploskoye()

    With frmShlifonaiye
    
        With .cboPripusk_Ploskoye
            .List = Split("0,3 0,4 0,5")
            .ListIndex = 0
        End With
        
        .cboDopuskPloskParallel.List = Split(dopusk)
        .cboDopuskPloskParallel.ListIndex = 0
        
        With .cboTolschinaDetali_Ploskoye
            .AddItem "<=2"
            .AddItem "<=5"
            .AddItem ">5"
            .ListIndex = 2
        End With
        
        With .cboShirinaKruga_Ploskoye
            .List = Split(ShirinaKruga, ";")
            .ListIndex = 0
        End With

        With .cboIT_Ploskoye
            .AddItem 9
            .AddItem "6, 7"
            .AddItem "5"
            .ListIndex = 0
        End With
        
    End With

End Sub

Private Sub SpisokKvalitetov(cboIT As MSForms.ComboBox)

    With cboIT
        .ColumnCount = 2
        .ListWidth = .Width & " pt"
        .ColumnWidths = .Width & " pt;0 pt"
        
        .AddItem
        .List(.ListCount - 1, 0) = "11, 12"
        .List(.ListCount - 1, 1) = 11
        
        .AddItem
        .List(.ListCount - 1, 0) = 9
        .List(.ListCount - 1, 1) = 9
        
        .AddItem
        .List(.ListCount - 1, 0) = "6, 7"
        .List(.ListCount - 1, 1) = 6
        
        .AddItem
        .List(.ListCount - 1, 0) = 5
        .List(.ListCount - 1, 1) = 5
        
        .ListIndex = 2

    End With

End Sub


Private Sub VyravnivaniyeElementovFormy()

    Dim ctrl As Control
    Dim lstShlifovaniye As Control, txtTsht_Kontrol As Control, txtTsht_Sles As Control, txtTsht As Control

    With frmShlifonaiye

        .fraZagolovok.Width = .Width
    
        For Each ctrl In .Controls
            Select Case True
                Case ctrl.Name Like "fraTablitsa_*"
                
                    With ctrl
                        .Top = 0
                        .Width = 261
                        .Left = frmShlifonaiye.mltShlifonaiye.Width - .Width - frmShlifonaiye.mltShlifonaiye.TabFixedWidth - 3
                        .Height = frmShlifonaiye.mltShlifonaiye.Height - 3 - .Top
                        .SpecialEffect = fmSpecialEffectFlat
                    End With
                    
                    With .Controls(Replace(ctrl.Name, "fraTablitsa", "lblZagolovok"))
                        .Left = 6
                        .Width = ctrl.Left
                        .TextAlign = fmTextAlignLeft
                    End With
                    
                    Set lstShlifovaniye = .Controls(Replace(ctrl.Name, "fraTablitsa", "lstShlifovaniye"))
                    With lstShlifovaniye
                        .Left = 6
                        .Top = 18
                        .Width = 252
                        .Height = 225
                        .SpecialEffect = fmSpecialEffectEtched
                    End With
                    
                    Set txtTsht_Kontrol = .Controls(Replace(ctrl.Name, "fraTablitsa", "txtTsht_Kontrol"))
                    With txtTsht_Kontrol
                        .Top = lstShlifovaniye.Top + lstShlifovaniye.Height + 3
                        .Width = 48
                        .Left = lstShlifovaniye.Left + lstShlifovaniye.Width - .Width * 2
                        .SpecialEffect = fmSpecialEffectEtched
                    End With
                    
                    Set txtTsht_Sles = .Controls(Replace(ctrl.Name, "fraTablitsa", "txtTsht_Sles"))
                    With txtTsht_Sles
                        .Top = lstShlifovaniye.Top + lstShlifovaniye.Height + 3
                        .Width = 48
                        .Left = lstShlifovaniye.Left + lstShlifovaniye.Width - .Width
                        .SpecialEffect = fmSpecialEffectEtched
                    End With
                    
                    Set txtTsht = .Controls(Replace(ctrl.Name, "fraTablitsa", "txtTsht"))
                    With txtTsht
                        .Left = txtTsht_Kontrol.Left - .Width
                        .Top = lstShlifovaniye.Top + lstShlifovaniye.Height + 3
                        .Width = 48
                        .SpecialEffect = fmSpecialEffectEtched
                    End With
                
                Case ctrl.Name Like "fraZagolovok_*"
                    With ctrl
                        .Left = 0
                        .Top = 0
                        .Width = .Width
                        .SpecialEffect = fmSpecialEffectFlat
                    End With

            End Select
        Next

    End With

End Sub
