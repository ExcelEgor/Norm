VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatron_Frm_Redaktirovaniye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Parametry As clsDatron

Private StrokaRedaktirovaniya As Range

Public Sub RedaktirovatPerekhod(wsRaschot As Worksheet, tbRaschot As ListObject)
    
    frmDatron.RezhimRedaktirovaniya = True
    frmDatron.lstRaschot.Enabled = False

    Set StrokaRedaktirovaniya = tbRaschot.ListRows(frmDatron.lstRaschot.ListIndex + 1).Range
    Set Parametry = New clsDatron
    Set Parametry = Parametry.PrisvoitParametryIzTablitsy(CInt(wsRaschot.Range("B1")), tbRaschot.HeaderRowRange, StrokaRedaktirovaniya)
    
    Call OtkrytNuzhnuyuVkladku
    Call UstanovitMaterial
    
    With frmDatron
        
        Select Case .mltDatron.SelectedItem.Name
            
            Case "Ustanovka"
                Call Ustanovka
                
            Case "FrezerovaniyeTortsevoyFrezoy"
                Call FrezerovaniyeTortsevoyFrezoy

            Case "ObrabotkaOtverstiy"
                Call ObrabotkaOtverstiy

            Case "ObrabotkaRezbovykhOtverstiy"
                Call ObrabotkaRezbovykhOtverstiy

            Case "VyrezaniyePoKonturu"
                Call VyrezaniyePoKonturu

            Case "FrezerovaniyeObnizheniy"
                Call FrezerovaniyeObnizheniy

            Case "StupenchatoyeFrezerovaniye"
                Call StupenchatoyeFrezerovaniye

            Case "FrezPazov"
                Call FrezPazov

            Case "FrezUstup"
                Call FrezUstup
                
            .txtPerekhod.text = Parametry.OpisaniyePerekhoda
            .chk_tSlec.Value = Parametry.UchestSlesar
            .chk_tKontr.Value = Parametry.UchestKontrol
        
        End Select

    End With
    
End Sub

Private Sub OtkrytNuzhnuyuVkladku()

    Dim i As Integer
    For i = 0 To frmDatron.mltDatron.Pages.Count - 1
        If frmDatron.mltDatron.Pages(i).Name = Parametry.id Then
            frmDatron.mltDatron.Value = i
            Exit For
        End If
    Next
    
End Sub

Private Sub UstanovitMaterial()
    Dim i As Integer
    For i = 0 To frmDatron.cboMaterial.ListCount - 1
        If CInt(frmDatron.cboMaterial.List(i, 1)) = CInt(Parametry.Material) Then
            frmDatron.cboMaterial.ListIndex = i
            Exit For
        End If
    Next
End Sub

Private Sub Ustanovka()

    With Parametry
        frmDatron.txtMassa = CStr(.Massa)
        frmDatron.lstUstanovka.ListIndex = CInt(.SposobUstanovki) - 1
        frmDatron.txtKolVo_Ustanovka = CStr(.KolVoPoverkhnostey)
    End With
    
End Sub

Private Sub FrezerovaniyeTortsevoyFrezoy()

    With Parametry
        frmDatron.txtDlina_FrezTorts = CStr(.Dlina)
        frmDatron.txtShirina_FrezTorts = CStr(.Shirina)
        frmDatron.txtPripusk_FrezTorts = CStr(.Pripusk)
        frmDatron.txtKolVo_FrezTorts = CStr(.KolVoPoverkhnostey)
    End With

End Sub

Private Sub ObrabotkaOtverstiy()

    With Parametry
        frmDatron.txtD_otv = CStr(.Diametr)
        frmDatron.txtL_otv = CStr(.Glubina)
        If .Glukhoye = True Then
            If .Dobnizheniya > 0 Then
                frmDatron.optGlukhoyeSObnizh_otv.Value = True
            Else
                frmDatron.optGlukhoye_otv.Value = True
            End If
        End If
        frmDatron.txtVysotaFaski = CStr(.VysotaFaski)
        frmDatron.txtDobnizh_otv = CStr(.Dobnizheniya)
        frmDatron.txtLobnizh_otv = CStr(.Lobnizheniya)
        frmDatron.tktKolVo_otv = CStr(.KolVoPoverkhnostey)
    End With

End Sub

Private Sub ObrabotkaRezbovykhOtverstiy()

    With Parametry
        frmDatron.txtD_RezbOtv = CStr(.Diametr)
        frmDatron.txtLotv_RezbOtv = CStr(.Glubina)
        frmDatron.txtLrezb_RezbOtv = CStr(.GlubinaRezby)
        frmDatron.optGlukhoye_RezbOtv.Value = .Glukhoye
        frmDatron.txtDobn_RezbOtv = CStr(.Dobnizheniya)
        frmDatron.txtLobn_RezbOtv = CStr(.Lobnizheniya)
        frmDatron.txtKolVo_RezbOtv = CStr(.KolVoPoverkhnostey)
    End With
    
End Sub

Private Sub VyrezaniyePoKonturu()

    With Parametry
        frmDatron.txtDlina_Vyrezaniye = CStr(.Dlina)
        frmDatron.txtGlubina_Vyrezaniye = CStr(.Glubina)
        frmDatron.txtKolVo_Vyrezaniye = CStr(.KolVoPoverkhnostey)
    End With

End Sub

Private Sub FrezerovaniyeObnizheniy()

    With Parametry
        frmDatron.txtDlina_Obnizheniye = CStr(.Dlina)
        frmDatron.txtShirina_Obnizheniye = CStr(.Shirina)
        frmDatron.txtGlubina_Obnizheniye = CStr(.Glubina)
        frmDatron.txtKolVo_Obnizheniye = CStr(.KolVoPoverkhnostey)
        frmDatron.txtRadius_Obnizheniye = CStr(.Radius)
        frmDatron.chkBezZachistki = .BezZachistki
    End With

End Sub

Private Sub StupenchatoyeFrezerovaniye()

    With Parametry
        frmDatron.txtLfrez = CStr(.Dlina)
        frmDatron.txtLkrivoy = CStr(.DlinaKrivoy)
        frmDatron.txtKolVo_StupFrez = CStr(.KolVoPoverkhnostey)
    End With

End Sub

Private Sub FrezPazov()

    With Parametry
        frmDatron.txtDlina_Paz = CStr(.Dlina)
        frmDatron.txtShirina_Paz = CStr(.Shirina)
        frmDatron.txtGlubina_Paz = CStr(.Glubina)
        frmDatron.txtKolVo_Paz = CStr(.KolVoPoverkhnostey)
    End With

End Sub

Private Sub FrezUstup()

    With Parametry
        frmDatron.txtDlina_Ustup = CStr(.Dlina)
        frmDatron.txtShirina_Ustup = CStr(.Shirina)
        frmDatron.txtVysota_Ustup = CStr(.Glubina)
        frmDatron.txtKolVo_Ustup = CStr(.KolVoPoverkhnostey)
    End With
    
End Sub
