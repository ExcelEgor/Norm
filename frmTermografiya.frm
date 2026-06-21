VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTermografiya 
   Caption         =   "Термография 300х400"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15150
   OleObjectBlob   =   "frmTermografiya.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTermografiya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ArrTxt() As New clsmZapuskRaschetov
Dim ArrChk() As New clsmZapuskRaschetov

Const KolVoStrok As Integer = 7

Private Sub txtMinuty_Enter(): Call EnterTextBox(txtMinuty): End Sub

Private Sub txtMinuty_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ExitTextBox(txtMinuty): End Sub

Private Sub txtSekundy_Enter(): Call EnterTextBox(txtSekundy): End Sub

Private Sub txtSekundy_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ExitTextBox(txtSekundy): End Sub

Private Sub lblTTP_Click()
    Dim obj As Object
    Set obj = CreateObject("Shell.Application")
    obj.Open ("\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\ТТП\ДАИЕ.02221.00001 - ТТП изготовления планок с надписью с наненсением изображения термографическим способом.pdf")

End Sub

Private Sub UserForm_Initialize()

    Call SozdatTablitsu

    Dim iTxt As Integer, iChk As Integer
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If Not LCase(ctrl.Tag) = "notclsm" Then
            Select Case TypeName(ctrl)
                Case "TextBox"
                    ReDim Preserve ArrTxt(iTxt)
                    Set ArrTxt(iTxt).txtTermograf = ctrl
                    iTxt = iTxt + 1
                Case "CheckBox"
                    ReDim Preserve ArrChk(iChk)
                    Set ArrChk(iChk).chkTermograf = ctrl
                    iChk = iChk + 1
            End Select
        End If
    Next
    
    NastroitRazmerFormy Me, fraRaschot, fraRaschot
    lblZagolovok.Left = -1
    lblZagolovok.Width = Me.Width + 1
    
    lblTTP.Left = 0
    lblTTP.Width = Me.Width
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)

End Sub

Private Sub SozdatTablitsu()

    Dim Elementy()
    Dim iElement As Integer
    
    Dim Stroka As Integer
    For Stroka = 2 To KolVoStrok
        
        iElement = 0
        
        SozdatElement chkFaska_1, Elementy, iElement, Stroka
        SozdatElement txtD_1, Elementy, iElement, Stroka
        SozdatElement txtKolVo_1, Elementy, iElement, Stroka
        SozdatElement lblFaska_1, Elementy, iElement, Stroka

        VyravnitElementyTablitsy Elementy, Stroka, Me

    Next
        
End Sub

Private Sub SozdatElement(ctrl As Control, ByRef Elementy As Variant, iElement As Integer, Stroka As Integer)

    Dim newCtrl As Control
    Set newCtrl = fraZausentsy.Controls.Add("Forms." & TypeName(ctrl) & ".1", Split(ctrl.Name, "_")(0) & "_" & Stroka)
    
    ReDim Preserve Elementy(iElement)
    Set Elementy(iElement) = newCtrl
    iElement = iElement + 1
    
End Sub

Sub RaschotTermografiya()
        
    Dim KolVo_Planka As Double, A_Planka As Double, B_Planka As Double, Minuty As Double, Sekundy As Double
    
    KolVo_Planka = DblFromCtrl(txtKolVo_Planka)
    A_Planka = DblFromCtrl(txtA_Planka)
    B_Planka = DblFromCtrl(txtB_Planka)
    Minuty = DblFromCtrl(txtMinuty)
    Sekundy = DblFromCtrl(txtSekundy)
    
    Dim i As Integer, j As Integer
    Dim ArrDiametry(1 To KolVoStrok) As Double
    Dim ArrKolVo(1 To KolVoStrok) As Double
    Dim ArrZenkovaniye(1 To KolVoStrok) As Boolean

    For i = 1 To KolVoStrok
        ArrZenkovaniye(i) = Me.Controls("chkFaska_" & i).Value
        ArrDiametry(i) = DblFromCtrl(Me.Controls("txtD_" & i))
        ArrKolVo(i) = DblFromCtrl(Me.Controls("txtKolVo_" & i))
    Next
    
    Dim Termograf As New clsTermograf300_400
    Dim Normy As Variant
    
    With Termograf
        .ArrDiametry = ArrDiametry
        .ArrKolVo = ArrKolVo
        .ArrZenkovaniye = ArrZenkovaniye
        .DlinaPlanki = A_Planka
        .KolVoDet = KolVo_Planka
        .ShirnaPlanki = B_Planka
        .tMin = Minuty
        .tSek = Sekundy
        
        Normy = .NormyVremeni
        
        If IsArray(Normy) Then
            lstRaschot.List = Normy
        End If
        
        
    End With
    
    
End Sub

