Attribute VB_Name = "mTokFrez_Zagruzka"
Option Explicit
Option Private Module

Private ArrTxt() As New clsTokFrez_Frm
Private ArrCbo() As New clsTokFrez_Frm
Private ArrCmd() As New clsTokFrez_Frm
Private ArrTgl() As New clsTokFrez_Frm
Private ArrChk() As New clsTokFrez_Frm
Private ArrOpt() As New clsTokFrez_Frm
Private ArrTabs() As New clsTokFrez_Frm
Private ArrLst() As New clsTokFrez_Frm
Private ArrFra() As New clsTokFrez_Frm
Private ArrSpn() As New clsTokFrez_Frm

Sub ZagruzitFormu_TokFrez()

    frmTokFrez.Top = Application.Top
    frmTokFrez.Left = Application.Left

    Call ZapolnitElementyUpravleniya_TokFrez
    
    Call VyravnitElementyFormy_TokFrez
    
    Call ZagruzkaElementovVKlass
    
    Call PerveyaZagruzkaTablitsty

End Sub

Private Sub ZagruzkaElementovVKlass()

    Dim ctrl As Control
    
    With frmTokFrez
        Set .KollektsiyaKartinokKakKnopok = New Collection
        For Each ctrl In .fraMainMenu.Controls
            If TypeOf ctrl Is MSForms.Image Then
                Set .KartinkaKakKnopka = New clsKartinkiKakKnopki
                Set .KartinkaKakKnopka.ImageControl = ctrl
                .KollektsiyaKartinokKakKnopok.Add .KartinkaKakKnopka
            End If
        Next

        Dim iTxt As Integer, iOpt As Integer, iCbo As Integer, iLst As Integer, iChk As Integer, iCmd As Integer, _
            iFra As Integer, iTgl As Integer, iTab As Integer, iLbl As Integer, iSpn As Integer
        
        For Each ctrl In .Controls
            If Not ctrl.Name Like "*RaschetShestigran*" Then
                Select Case TypeName(ctrl)
                    Case "TextBox"
                        If ctrl.Tag <> "NotCLSM" Then
                            ReDim Preserve ArrTxt(iTxt)
                            Set ArrTxt(iTxt).txt = ctrl
                            iTxt = iTxt + 1
                        End If
                    Case "OptionButton"
                        ReDim Preserve ArrOpt(iOpt)
                        Set ArrOpt(iOpt).opt = ctrl
                        iOpt = iOpt + 1
                    Case "ComboBox"
                        If ctrl.Tag <> "NotCLSM" Then
                            ReDim Preserve ArrCbo(iCbo)
                            Set ArrCbo(iCbo).cbo = ctrl
                            iCbo = iCbo + 1
                        End If
                    Case "ListBox"
                        ReDim Preserve ArrLst(iLst)
                        Set ArrLst(iLst).lst = ctrl
                        iLst = iLst + 1
                    Case "CheckBox"
                        ReDim Preserve ArrChk(iChk)
                        Set ArrChk(iChk).chk = ctrl
                        iChk = iChk + 1
                    Case "CommandButton"
                        ReDim Preserve ArrCmd(iCmd)
                        Set ArrCmd(iCmd).cmd = ctrl
                        iCmd = iCmd + 1
                    Case "Frame"
                        ReDim Preserve ArrFra(iFra)
                        Set ArrFra(iFra).fra = ctrl
                        iFra = iFra + 1
                    Case "ToggleButton"
                        ReDim Preserve ArrTgl(iTgl)
                        Set ArrTgl(iTgl).tgl = ctrl
                        iTgl = iTgl + 1
                    Case "TabStrip"
                        ReDim Preserve ArrTabs(iTab)
                        Set ArrTabs(iTab).tabs = ctrl
                        iTab = iTab + 1
                    Case "SpinButton"
                        ReDim Preserve ArrSpn(iSpn)
                        Set ArrSpn(iSpn).Spn = ctrl
                        iSpn = iSpn + 1
                End Select
            End If
        Next
    
    End With
End Sub

Private Sub PerveyaZagruzkaTablitsty()
    
    Set frmTokFrez.Raschot = New clsRaschotVKnige
    
    With frmTokFrez
        
        With .Raschot
            Set .ListBox = frmTokFrez.lstTokarnoFrezer
            Set .ChkOkruglennyyeNomvy = frmTokFrez.chkOkruglyonnyyeNormy
            Set .txtTsht_Kontrol = frmTokFrez.txtTsht_Kontrol
            Set .txtTsht_Slesar = frmTokFrez.txtTsht_Slesar
            Set .txtTsht_TokFrez = frmTokFrez.txtTsht_TokFrez
            Set .cboMaterial = frmTokFrez.cboMaterial
            Set .cboTipProizv = frmTokFrez.cboTipProizv
            Set .chkPoRezhimam = frmTokFrez.chkPoRezhimam
            Set .chkTonkosten = frmTokFrez.chkTonkosten
            Set .chkCHPU = frmTokFrez.chkCHPU
            Set .tabOperatsii = frmTokFrez.tabOperatsii
        End With
        
        Set .Kniga = New clsKniga
        Call .Kniga.Init(wsTablitsa_TokarnoFrezer)
        
        Call .Raschot.SozdatNovyyRaschot(.Kniga)
    
    End With
    
End Sub
