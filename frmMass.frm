VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMass 
   Caption         =   "׀אסקוע לאססû"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900.001
   OleObjectBlob   =   "frmMass.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmMass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Handler As clsGlavnyyRaschot
Private ControlHandlers As Collection

Private Massa As Double, Ploschad As Double, PloschadPoperSech As Double

Private Sub UserForm_Initialize()

    Dim tbMaterial As ListObject
    Set tbMaterial = wsMaterial.ListObjects("tbMainMaterial")
    cbo_material.List = tbMaterial.DataBodyRange.Value
    cbo_material.ListIndex = 0
    
    With cboEdIzm_M
        .List = Split("ד ךד")
        .ListIndex = 1
    End With

    With cboEdIzm_S
        .List = Split("לל2 סל2 הל2 ל2")
        .ListIndex = 3
    End With
    
    With cboEdinitsaIzmereniya_PloschadPoperSech
        .List = Split("לל2 סל2 הל2 ל2")
        .ListIndex = 0
    End With
    
    DobavitVKlass_KontrolVvodaChisel Me
    
    Set ControlHandlers = New Collection
    DobavitVKlass_GlavnyyRaschot Me, ControlHandlers, Handler
    
    GlavnyyRaschot
    
End Sub


Public Sub GlavnyyRaschot()
    
    Dim Material As EnumMaterialy
    Material = CInt(cbo_material.List(cbo_material.ListIndex, 1))
    
    lblPlotnost.Caption = "P=" & PlotnostMateriala(CInt(Material)) & " ךד/ל" & ChrW(179)
    
    If Not IsNumeric(cbo_material.Column(1)) Then Exit Sub
    
    Select Case MultiPage1.SelectedItem.Name
        Case "pTruba"
            Dim Truba As New clsFigura_Truba
            Truba.Init Material, DblFromCtrl(txtDiametr_Truba), DblFromCtrl(txtTolshchinaStenki_Truba), DblFromCtrl(txtDlina_Truba)
            RaschitatFigury Truba
            
        Case "pPrutok"
            Dim Krug As New clsFigura_Krug
            Krug.Init Material, DblFromCtrl(txtD_Prutok), DblFromCtrl(txtL_Prutok)
            RaschitatFigury Krug
            
        Case "pList"
            Dim List As New clsFigura_List
            List.Init Material, DblFromCtrl(txt_b_list), DblFromCtrl(txt_h_list), DblFromCtrl(txt_a_list)
            RaschitatFigury List
        
        Case "pUgolnik"
            Dim Ugolnik As New clsFigura_Ugolnik
            Ugolnik.Init Material, DblFromCtrl(txtUgolnik_W), DblFromCtrl(txtUgolnik_H), DblFromCtrl(txtUgolnik_a), DblFromCtrl(txtUgolnik_b), DblFromCtrl(txtUgolnik_L)
            RaschitatFigury Ugolnik
    
        Case "pShestigran"
            Dim Shestigrannik As New clsFigura_Shestigrnnik
            Shestigrannik.Init Material, DblFromCtrl(txtShestigran_S), DblFromCtrl(txtShestigran_L)
            RaschitatFigury Shestigrannik
            
    End Select

    Call ZapisatRezultatyRaschotov

End Sub

Private Sub RaschitatFigury(Figura As ITechnologicalShape)
    Massa = Figura.Massa
    Ploschad = Figura.PloshchadPolnaya
    PloschadPoperSech = Figura.PloshchadPoperSech
End Sub

Private Sub ZapisatRezultatyRaschotov()

    Dim DelitelPloschad As Long, DelitelPloschadPoperSech As Long, DelitelMassa As Integer
    
    Select Case cboEdIzm_S.text
        Case "לל2": DelitelPloschad = 1
        Case "סל2": DelitelPloschad = 100
        Case "הל2": DelitelPloschad = 10000
        Case "ל2":  DelitelPloschad = 1000000
    End Select
    txtPloschad = Format(Ploschad / DelitelPloschad, "0.000")
    
    Select Case cboEdinitsaIzmereniya_PloschadPoperSech.text
        Case "לל2": DelitelPloschadPoperSech = 1
        Case "סל2": DelitelPloschadPoperSech = 100
        Case "הל2": DelitelPloschadPoperSech = 10000
        Case "ל2":  DelitelPloschadPoperSech = 1000000
    End Select
    txtPloschadPoperSech = Format(PloschadPoperSech / DelitelPloschadPoperSech, "0.000")
    
    Select Case cboEdIzm_M.text
        Case "ד": DelitelMassa = 1000
        Case "ךד": DelitelMassa = 1
    End Select
    
    txtMassa = Format(Massa / DelitelMassa, "0.000")
        

End Sub
