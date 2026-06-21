VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTermografiya_2 
   Caption         =   "Термография 40х297"
   ClientHeight    =   9315.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15960
   OleObjectBlob   =   "frmTermografiya_2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTermografiya_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const KolVoStrok As Integer = 16
Dim tbTermografiya_Vremya As ListObject
Dim tbZaucenciTermograf As ListObject
Dim tbObrabotkaOtverstiyTermograf As ListObject

Private Sub cmdRaschot_Click()
    With mltTermografiya
        If .Value = 0 Then
            If OsnovnyyeParametryZapolneny() = True Then
                .Value = 1
                With cmdRaschot
                    .BackColor = &HFFFFC0
                    .Caption = "Назад к параметрам"
                End With
                Call RaschotTermografiya
            End If
        Else
            .Value = 0
            With cmdRaschot
                .BackColor = &HC0FFC0
                .Caption = "Расчёт"
            End With
        End If
    End With
    
End Sub
    

Private Sub lblTTP_Click()
    Dim obj As Object
    Set obj = CreateObject("Shell.Application")
    obj.Open ("\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\ТТП\ДАИЕ.02221.00001 - ТТП изготовления планок с надписью с наненсением изображения термографическим способом.pdf")

End Sub


Private Sub UserForm_Initialize()

    With wsTermografiya_2
        Set tbTermografiya_Vremya = .ListObjects("tbTermografiya_2")
        Set tbZaucenciTermograf = .ListObjects("tbZaucenciTermograf_2")
        Set tbObrabotkaOtverstiyTermograf = .ListObjects("tbObrabotkaOtvTermograf_2")
    End With
    
    cboPerekhod_1.List = wsKomplektovaniye.ListObjects("tbTipObrOtv").DataBodyRange.Value
    
    Call SozdaniyeElementov
    
    With lblZagolovok
        .Left = -1
        .Width = Me.Width - .Left
    End With
    
    With lblTTP
        .Left = (Me.Width - .Width) / 2
    End With
    
    
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
End Sub

Private Sub SozdaniyeElementov()

    Dim chkFaska As MSForms.CheckBox
    Dim txtD As MSForms.TextBox, txtA As MSForms.TextBox, txtB As MSForms.TextBox, txtKolVo As MSForms.TextBox, txtL As MSForms.TextBox
    Dim cboPerekhod As MSForms.ComboBox
    
    Dim i As Integer
    
    For i = 2 To KolVoStrok
    
        With fraZausentsy.Controls
            Set chkFaska = .Add("Forms.CheckBox.1", "chkFaska" & "_" & i)
            Set txtD = .Add("Forms.TextBox.1", "txtD" & "_" & i)
            Set txtA = .Add("Forms.TextBox.1", "txtA" & "_" & i)
            Set txtB = .Add("Forms.TextBox.1", "txtB" & "_" & i)
            Set txtKolVo = .Add("Forms.TextBox.1", "txtKolvo" & "_" & i)
            Set txtL = .Add("Forms.TextBox.1", "txtL" & "_" & i)
        End With
        
        With fraObrabotkaOtrverstiy
            Set cboPerekhod = .Add("Forms.ComboBox.1", "cboPerekhod" & "_" & i)
            Set txtD = .Add("Forms.TextBox.1", "txtDotv" & "_" & i)
            Set txtKolVo = .Add("Forms.TextBox.1", "txtKolVoOtv" & "_" & i)
        End With
    Next
    
    Call VyravnivanieElementov(fraZausentsy)
    Call VyravnivanieElementov(fraObrabotkaOtrverstiy)

End Sub

Private Sub VyravnivanieElementov(fra As MSForms.Frame)

    Dim ctrl As Control
    Dim num As Integer
    Dim prevCtrl As Control
    
    For Each ctrl In fra.Controls
        If TypeName(ctrl) <> "Label" And TypeName(ctrl) <> "Frame" Then
            num = NumControl(ctrl)
            If num > 1 Then
                With ctrl
                    Set prevCtrl = Me.Controls(Left(.Name, Len(.Name) - Len(CStr(num))) & num - 1)
                    On Error Resume Next
                    .Font.Size = 10
                    .Font.Name = prevCtrl.Font.Name
                    .ForeColor = prevCtrl.ForeColor
                    .SpecialEffect = prevCtrl.SpecialEffect
                    .BorderStyle = prevCtrl.BorderStyle
                    .BorderColor = prevCtrl.BorderColor
                    .Width = prevCtrl.Width
                    .Left = prevCtrl.Left
                    .Top = 21 * num + 6
                    .SelectionMargin = prevCtrl.SelectionMargin
                    .Locked = prevCtrl.Locked
                    .TextAlign = prevCtrl.TextAlign
                    .Style = 2
                    .List = prevCtrl.List
                    On Error GoTo 0
                End With
            End If
        End If
    Next

    fra.Font.Name = "Segoe Ui Semibold"
    
End Sub
Private Function OsnovnyyeParametryZapolneny() As Boolean
    
    OsnovnyyeParametryZapolneny = True
    Dim ctrl As Control
    For Each ctrl In fraOsnovnyyeParametry.Controls
        If TypeName(ctrl) = "TextBox" And ctrl.Name <> "txtMinuty" And ctrl.Name <> "txtSekundy" Then
            If DblFromCtrl(ctrl) = 0 Then
                OsnovnyyeParametryZapolneny = False
                Exit Function
            End If
        End If
        
    Next
    
End Function

Sub RaschotTermografiya()
        
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim KolVoDetaley As Double, DlinaZagotovki As Double, ShirinaZagotovki As Double
    Dim DlinaDetali As Double, ShirinaDetali As Double
    Dim Tolshchina As Double
    Dim Minuty As Double, Sekundy As Double
    
    KolVoDetaley = DblFromCtrl(txtKolVoDetaley)
    DlinaZagotovki = DblFromCtrl(txtDlinaZagotovki)
    ShirinaZagotovki = DblFromCtrl(txtShirinaZagotovki)
    DlinaDetali = DblFromCtrl(txtDlinaDetali)
    ShirinaDetali = DblFromCtrl(txtShirinaDetali)
    Tolshchina = DblFromCtrl(txtTolshchina)
    Minuty = DblFromCtrl(txtMinuty)
    Sekundy = DblFromCtrl(txtSekundy)
    
    With wsTermografiya_2
        .Range("KolVoDetaley") = KolVoDetaley
        .Range("DlinaZagotovki") = DlinaZagotovki
        .Range("ShirinaZagotovki") = ShirinaZagotovki
        .Range("Tolshchina") = Tolshchina
        .Range("DlinaDetali") = DlinaDetali
        .Range("ShirinaDetali") = ShirinaDetali
        .Range("Minuty") = Minuty
        .Range("Sekundy") = Sekundy
        .Range("KolVoPeremychek") = DblFromCtrl(txtKolVoPeremychek)
    End With
    
    Dim i As Integer, j As Integer
    
    Dim Faska As String
    Dim d As Double, a As Double, b As Double, L As Double, KolVo As Double
    With tbZaucenciTermograf
        For i = 1 To KolVoStrok
            Faska = IIf(Me.Controls("chkFaska_" & i) = True, "Да", "Нет")
            d = DblFromCtrl(Me.Controls("txtD_" & i))
            a = DblFromCtrl(Me.Controls("txtA_" & i))
            b = DblFromCtrl(Me.Controls("txtB_" & i))
            L = DblFromCtrl(Me.Controls("txtL_" & i))
            KolVo = DblFromCtrl(Me.Controls("txtKolVo_" & i))
            
            If d > 0 Or a > 0 Or b > 0 Or L > 0 Then
                .DataBodyRange(i, .ListColumns("Зенкование").Index) = Faska
                .DataBodyRange(i, .ListColumns("D, мм").Index) = d
                .DataBodyRange(i, .ListColumns("i поверх.").Index) = KolVo
                .DataBodyRange(i, .ListColumns("L, мм").Index) = L
                .DataBodyRange(i, .ListColumns("A, мм").Index) = a
                .DataBodyRange(i, .ListColumns("B, мм").Index) = b
            Else
                For j = 1 To .ListColumns.Count - 2
                    .DataBodyRange(i, j).ClearContents
                Next
            End If
        Next
    End With
    
    Dim Perekhod As String
    With tbObrabotkaOtverstiyTermograf
        For i = 1 To KolVoStrok
            Perekhod = Me.Controls("CboPerekhod_" & i)
            d = DblFromCtrl(Me.Controls("txtDotv_" & i))
            KolVo = DblFromCtrl(Me.Controls("txtKolVoOtv_" & i))
            
            If d > 0 Then
                .DataBodyRange(i, 1) = Perekhod
                .DataBodyRange(i, 3) = d
                .DataBodyRange(i, 4) = KolVo
            Else
                For j = 1 To .ListColumns.Count - 2
                    .DataBodyRange(i, j).ClearContents
                Next
            End If
        Next
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    With lstRaschot
        .List = tbTermografiya_Vremya.DataBodyRange.Value
        For i = 0 To .ListCount - 1
            .List(i, .ColumnCount - 1) = CStr(.List(i, .ColumnCount - 1))
        Next
    End With

    
End Sub



