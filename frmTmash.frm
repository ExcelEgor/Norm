VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTmash 
   Caption         =   "Тшт: Пробивочно-вырезная"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4020
   OleObjectBlob   =   "frmTmash.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTmash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboMaterial_Change()
    Call RaschetTshtNaProbivochnoVireznomStanke
End Sub

Private Sub optAlyuminiy_Click()
    cboMaterial.ListIndex = 0
End Sub

Private Sub optStla_Click()
    cboMaterial.ListIndex = 1
End Sub

Private Sub txtDet_A_Change()
    Call RaschetTshtNaProbivochnoVireznomStanke
End Sub

Private Sub txtDet_B_Change()
    Call RaschetTshtNaProbivochnoVireznomStanke
End Sub

Private Sub txtTmash_Min_Change()
    Call RaschetTshtNaProbivochnoVireznomStanke
End Sub

Private Sub txtTmash_Sek_Change()
    Call RaschetTshtNaProbivochnoVireznomStanke
End Sub

Private Sub txtZag_A_Change()
    Call RaschetTshtNaProbivochnoVireznomStanke
End Sub

Private Sub txtZag_B_Change()
    Call RaschetTshtNaProbivochnoVireznomStanke
End Sub

Private Sub txtZag_C_Change()
    Call RaschetTshtNaProbivochnoVireznomStanke
End Sub

Private Sub UserForm_Initialize()

    Dim tbMaterial As ListObject
    Set tbMaterial = wsMaterial.ListObjects("tbMaterial_LaserPress")
    cboMaterial.List = tbMaterial.DataBodyRange.Value
    cboMaterial.ListIndex = 0
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "TextBox"
                If ctrl.Name <> "txtTsht" Then
                    ReDim Preserve ArrTxtKontrolVvodaChisel(iTxt_KontrolVvodaChisel)
                    Set ArrTxtKontrolVvodaChisel(iTxt_KontrolVvodaChisel).txt = ctrl
                    iTxt_KontrolVvodaChisel = iTxt_KontrolVvodaChisel + 1
                End If
            Case "ComboBox"
                ReDim Preserve ArrCboKontrolVvodaChisel(iCbo_KontrolVvodaChisel)
                Set ArrCboKontrolVvodaChisel(iCbo_KontrolVvodaChisel).cbo = ctrl
                iCbo_KontrolVvodaChisel = iCbo_KontrolVvodaChisel + 1
        End Select
    Next
    
    optAlyuminiy.Value = True
    
    txtZag_B.SetFocus
    If Num_NOt_Stable() Then Application.SendKeys "+{NUMLOCK}"
    Application.SendKeys "{TAB}"
    
End Sub

Private Sub RaschetTshtNaProbivochnoVireznomStanke()
    
    With frmTmash
    
        .txtTsht.text = Empty
       
        Dim DlinaZagotovki As Double, ShirinaZagotovki As Double, TolschinaZagotovkiotovki As Double
        DlinaZagotovki = WorksheetFunction.Max(DblFromCtrl(.txtZag_A), DblFromCtrl(.txtZag_B))
        ShirinaZagotovki = WorksheetFunction.Min(DblFromCtrl(.txtZag_A), DblFromCtrl(.txtZag_B))
        TolschinaZagotovkiotovki = DblFromCtrl(.txtZag_C)
        
        Dim DlinaDetali As Double, ShirinaDetali As Double
        DlinaDetali = WorksheetFunction.Max(DblFromCtrl(.txtDet_A), DblFromCtrl(.txtDet_B))
        ShirinaDetali = WorksheetFunction.Min(DblFromCtrl(.txtDet_A), DblFromCtrl(.txtDet_B))
        
        Dim tMash As Double
        tMash = DblFromCtrl(.txtTmash_Min) + (DblFromCtrl(.txtTmash_Sek) / 60)
        
        If DlinaZagotovki > 0 And ShirinaZagotovki > 0 And TolschinaZagotovkiotovki > 0 And _
            DlinaDetali > 0 And ShirinaDetali > 0 And tMash > 0 And DlinaZagotovki > DlinaDetali And ShirinaZagotovki > ShirinaDetali Then

            Dim Tsht As Double
            Tsht = TshtLaserPress(DlinaZagotovki, ShirinaZagotovki, TolschinaZagotovkiotovki, DlinaDetali, ShirinaDetali, _
                DblFromCtrl(.txtTmash_Min), DblFromCtrl(.txtTmash_Sek), IIf(.optAlyuminiy, ALUMINIYEVYYE_SPLAVY, STAL_UGLERODISTAYA), DblFromCtrl(.txtKolVo))
            
            .txtTsht.text = CStr(Tsht)
            
        End If
    
    End With
    
End Sub


