Attribute VB_Name = "mKoordinat_Vyravnivaniye"
Option Explicit
Option Private Module

Public Sub VyravnitElementyFormy_Koordinat(ByVal frm As Object)
    
    Const MARGIN As Single = 3
    Const MARGIN_LARGE As Single = 15
    Const MARGIN_BOTTOM As Single = 30
    
    Dim ctrl As Control
    
    With frm
        
        With .lblZagolovok
            .Top = -1
            .Left = -1
            .Width = frm.Width + 1
        End With
        
        With .fraMaterial
            .Left = MARGIN
            .Top = frm.lblZagolovok.Top + frm.lblZagolovok.Height + MARGIN
        End With
        
        With .fraMainKoef
            .Left = frm.fraMaterial.Left + frm.fraMaterial.Width + MARGIN
            .Top = frm.fraMaterial.Top
            .Width = frm.Width - .Left - MARGIN_LARGE
        End With

        With .fraVremya
            .Top = frm.Height - .Height - MARGIN_BOTTOM
            .Left = frm.fraMaterial.Left
            .Width = frm.Width - .Left - frm.fraRaschot.Width - 18
        End With
        
        With .fraRaschot
            .Top = frm.fraMaterial.Top + frm.fraMaterial.Height + MARGIN
            .Left = frm.Width - .Width - MARGIN_LARGE
            .Height = frm.fraVremya.Top + frm.fraVremya.Height - .Top
        End With
        
        With .fraMlt
            .Left = frm.fraMaterial.Left
            .Top = frm.fraRaschot.Top
            .Width = frm.fraVremya.Width
            .Height = frm.fraVremya.Top - .Top - MARGIN
        End With
        
        .lstKoordinat.Height = frm.fraMlt.Height - frm.lstKoordinat.Top
        
        With .mltKoordinat
            .Left = 0
            .Top = 0
            .Width = frm.fraMlt.Width
            .Height = frm.fraMlt.Height
        End With
        
        For Each ctrl In .Controls
            With ctrl
                Select Case True
                    Case .Name Like "lblZagolovok_*"
                        .Top = MARGIN
                        .Left = MARGIN
                        .Width = frm.mltKoordinat.Width - frm.mltKoordinat.TabFixedWidth - 12
                        
                    Case .Name Like "fraSovmOsi_*"
                        .Top = 21
                        .Left = MARGIN
                        .Width = frm.mltKoordinat.Width - frm.mltKoordinat.TabFixedWidth - 12
                        
                    Case .Name Like "txtTsht_*"
                        .Top = frm.lstKoordinat.Top + frm.lstKoordinat.Height
                        
                    Case .Name Like "fraKoef_*"
                        .Width = frm.mltKoordinat.Width - frm.mltKoordinat.TabFixedWidth - 9 - .Left
                End Select
            End With
        Next ctrl
        
    End With
    
End Sub
