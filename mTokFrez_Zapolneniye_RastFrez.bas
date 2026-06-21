Attribute VB_Name = "mTokFrez_Zapolneniye_RastFrez"
Option Explicit

Sub ZapolnitTokFrez_RastochnyyeFrezernyyeStanki()
   
    With frmTokFrez
            
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("Вспомогательный")
        
        Dim RangeRaKontsevyye As Range
        Set RangeRaKontsevyye = wsVspomogatelnyy.Range("J11:M15")
        
        ZapolnitKontsevyyeFrezy RangeRaKontsevyye
        ZapolnitTortsevyyeFrezy RangeRaKontsevyye
        ZapolnitShestigrannik RangeRaKontsevyye
        ZapolnitFrezerovaniyeUstupov
        ZapolnitFrezerovaniyeOkon
        ZapolnitShlitsevyyeFrezy
        ZapolnitFrezerovaniyePazov
        ZapolntiObrabotkuOtverstiy
        ZapolnitSbosobyUstanovki
        
    End With
    
End Sub

Private Sub ZapolnitKontsevyyeFrezy(RangeRa As Range)

    With frmTokFrez

        With .cboRaRz_FrezPloskKonc
            .RowSource = RangeRa.Address(external:=True)
            .ColumnHeads = True
            .ListIndex = 1
        End With
        With .cboIT_FrezPloskKonc
            .List = Split("7 9 11 14")
            .ListIndex = 2
        End With
        
    End With
    
End Sub

Private Sub ZapolnitTortsevyyeFrezy(RangeRa As Range)

    With frmTokFrez
        With .cboRaRz_FrezPloskTorc
            .RowSource = RangeRa.Address(external:=True)
            .ListIndex = 1
        End With
        With .cboIT_FrezPloskTorc
            .List = frmTokFrez.cboIT_FrezPloskKonc.List
            .ListIndex = frmTokFrez.cboIT_FrezPloskKonc.ListIndex
        End With
    End With
    
End Sub

Private Sub ZapolnitFrezerovaniyeUstupov()

    With frmTokFrez
        With .cboRaRz_FrezUstup
            .RowSource = wsVspomogatelnyy.Range("A3:D6").Address(external:=True)
            .ListIndex = 1
        End With
        
        With .cboIT_FrezUstup
            .List = frmTokFrez.cboIT_FrezPloskKonc.List
            .ListIndex = frmTokFrez.cboIT_FrezPloskKonc.ListIndex
        End With
    End With
    
End Sub

Private Sub ZapolnitShestigrannik(RangeRa As Range)

    With frmTokFrez
        
        With .cboRaRz_Shestigrannik
            .RowSource = RangeRa.Address(external:=True)
            .ListIndex = 1
        End With
        With .cboIT_Shestigrannik
            .List = frmTokFrez.cboIT_FrezPloskKonc.List
            .ListIndex = .ListCount - 1
        End With
        
    End With
    
End Sub

Private Sub ZapolnitFrezerovaniyeOkon()
    With frmTokFrez.cboIT_FrezOkon
        .List = Split("7 8 9 11 14")
        .ListIndex = .ListCount - 1
    End With
    With frmTokFrez.cboRaRz_FrezOkon
        .RowSource = wsVspomogatelnyy.Range("A3:D5").Address(external:=True)
        .ColumnHeads = True
        .ListIndex = 0
    End With
End Sub

Private Sub ZapolnitShlitsevyyeFrezy()
    
    With frmTokFrez
        With .cboB_FrezShlic
            .List = Split("4 2 0,8")
            .ListIndex = 1
        End With
        
        With .cboRaRz_FrezShlic
            .RowSource = wsVspomogatelnyy.Range("A24:D25").Address(external:=True)
            .ListIndex = 0
        End With
        
        With .cboGlubinaRez_FrezShlic
            .List = Split("2 4 6 10 15")
            .ListIndex = 0
        End With
    End With
    
End Sub
      
Private Sub ZapolnitFrezerovaniyePazov()
    With frmTokFrez
        With .cboIT_FrezPazov
            .List = frmTokFrez.cboIT_FrezPloskKonc.List
            .ListIndex = .ListCount - 1
        End With
        
        With .cboRaRz_FrezPazov
            .RowSource = wsVspomogatelnyy.Range("A3:D5").Address(external:=True)
            .ListIndex = 1
        End With
    End With
End Sub

Private Sub ZapolntiObrabotkuOtverstiy()
    With frmTokFrez
        With .cboRaRz_FrezOtv
            .RowSource = wsVspomogatelnyy.Range("A3:D5").Address(external:=True)
            .ColumnHeads = True
            .ListIndex = 0
        End With
        With .cboIT_FrezOtv
            .List = frmTokFrez.cboIT_FrezPloskKonc.List
            .ListIndex = .ListCount - 1
        End With
    End With
End Sub
        
Private Sub ZapolnitSbosobyUstanovki()

    frmTokFrez.lstSposobUst_RastFrez.List = wsUst_Frez.Range("UstFrez_Sposob").Value
    
End Sub
