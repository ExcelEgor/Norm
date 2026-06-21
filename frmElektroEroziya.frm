VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmElektroEroziya 
   Caption         =   "Электроэрозия"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16890
   OleObjectBlob   =   "frmElektroEroziya.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmElektroEroziya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tElektroeroziya As Double, tSlesar_elektro As Double, tKontrol_elektro As Double

Dim tbElektroEroziya_Raschot As ListObject
Dim tbKvalitet As ListObject

Private Material As Integer
Private Kvalitet As Integer

Private Handler As clsGlavnyyRaschot
Private ControlHandlers As Collection

Private IdyotZagruzkaFormy As Boolean

Private Sub optAlyuminiy_Click()
    Material = 1
    Call SostavitSpisokKvalitetov(Material)
End Sub

Private Sub optMednyyeSplavy_Click()
    Material = 2
    Call SostavitSpisokKvalitetov(Material)
End Sub

Private Sub optStal_Click()
    Material = 3
    Call SostavitSpisokKvalitetov(Material)
End Sub

Private Sub UserForm_Initialize()

    IdyotZagruzkaFormy = True
    
    Set tbElektroEroziya_Raschot = wsElektroEroziya.ListObjects("tbElektroEroziya_Raschot")
    With tbElektroEroziya_Raschot
        If .ListRows.Count > 0 Then
            .DataBodyRange.Rows.Delete
        End If
    End With
    
    Set tbKvalitet = wsElektroEroziya.ListObjects("tbElektroEroziya_Kvalitet")
    
    With cboSloznostUstanovki
        .AddItem "Простая"
        .AddItem "Средняя"
        .AddItem "Сложная"
        .ListIndex = 0
    End With
    
    Set ControlHandlers = New Collection
    DobavitVKlass_GlavnyyRaschot Me, ControlHandlers, Handler
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    IdyotZagruzkaFormy = False
    
    mltElektroeroziya.Value = 0
    
    optAlyuminiy.Value = True
    
    Call VyravnitElementy
    
    Call GlavnyyRaschot
    
End Sub

Private Sub VyravnitElementy()

    Const Shirina As Integer = 342

    With lblZagolovok
        .Left = 0
        .Top = 0
    End With
    
    With fraMaterial
        .Left = 6
        .Top = lblZagolovok.Top + lblZagolovok.Height + 9
        .Width = Shirina
    End With
    
    With fraParametry
        .Left = fraMaterial.Left
        .Top = fraMaterial.Top + fraMaterial.Height + 3
        .Width = Shirina
    End With
    
    With mltElektroeroziya
        .Top = 9
        .Left = 0
        .Width = fraParametry.Width
        .Height = fraParametry.Height - .Top
    End With
    
    VyravnitPoLevomuKrayuNazvaniyaVkladok mltElektroeroziya
    
    With fraVremya
        .Left = fraMaterial.Left
        .Top = fraParametry.Top + fraParametry.Height + 3
        .Width = Shirina
    End With
    
    With fraRaschot
        .Left = fraMaterial.Left + fraMaterial.Width + 3
        .Top = fraMaterial.Top
        .Height = fraVremya.Top + fraVremya.Height - .Top
    End With
    
    With fraNormy
        .Left = 3
        .Width = lstRaschet.Width
        .Top = fraRaschot.Height - .Height - 12
    End With
    
    With lstRaschet
        .Left = 3
        .Height = fraNormy.Top - .Top - 3
    End With
    
    With Me
        .Width = fraRaschot.Left + fraRaschot.Width + 18
        .Height = fraVremya.Top + fraVremya.Height + 36
    End With
    
    lblZagolovok.Width = Me.Width

End Sub

Sub GlavnyyRaschot()

    If IdyotZagruzkaFormy = True Then Exit Sub
    
    Kvalitet = 0
    tElektroeroziya = 0
    txtVremya.text = Empty

    Call DostupnostKonfiguratsii
    
    Call ZagruzkaIzTablitsy_v_ListBox
    
    Select Case mltElektroeroziya.SelectedItem.Name
        Case "pUstanovka"
            Call RaschetVremeni_Ustanovka
        Case "pKontur"
            Call RaschetVremeni_Kontrur(Material)
        Case "pDorabotka"
            Call RaschetVremeni_Dorabotka(Material)
    End Select
    
    If tElektroeroziya > 0 Then
        txtVremya.text = Format(tElektroeroziya, "0.0")
    End If
    
    
End Sub

Private Sub SostavitSpisokKvalitetov(Material)

    Dim i As Integer, ProshlyyListIndex As Integer

    With cboKvalitet
    
        ProshlyyListIndex = .ListIndex
        
        .Clear

        For i = 1 To tbKvalitet.ListRows.Count - IIf(Material = 1, 0, 1)
            .AddItem
            .List(.ListCount - 1, 0) = tbKvalitet.DataBodyRange(i, 1)
            .List(.ListCount - 1, 1) = tbKvalitet.DataBodyRange(i, 2)
            .List(.ListCount - 1, 2) = tbKvalitet.DataBodyRange(i, 3)
        Next
        
        If ProshlyyListIndex <> -1 And .ListCount - 1 > ProshlyyListIndex Then
            .ListIndex = ProshlyyListIndex
        Else
            .ListIndex = .ListCount - 1
        End If
        
    End With
    
End Sub

Private Sub RaschetVremeni_Kontrur(Material As Integer)
    
    txtS_Kontur.BackColor = vbWhite
    
    If cboKvalitet.ListIndex <> -1 Then
    
        With cboKvalitet
        
            If .ListIndex = -1 Then Exit Sub
        
            Dim KolVoProkhodov As Integer
            KolVoProkhodov = CInt(.List(.ListIndex, 2))
            lblKolVoProhodovKontur = KolVoProkhodov & IIf(KolVoProkhodov > 1, " прохода", " проход")
            
            Kvalitet = CInt(.List(.ListIndex, 1))
            
        End With

        Dim Smax As Double
        Smax = ElektroEroz_Kontur_Smax(Material, Kvalitet, IIf(optKonfig_1_Kontur, 1, 2))
        lblSmax_Kontur.Caption = "до " & Smax
        
        Dim S As Double:        S = DblFromCtrl(txtS_Kontur)
        Dim L As Double:        L = DblFromCtrl(txtL_Kontur)
        Dim KolVo As Double:    KolVo = KolVoPoverhnosteiProhodov(txtKolVo_Kontur)

        If S > Smax Then
            txtS_Kontur.BackColor = vbRed
        Else
            If S > 0 And L > 0 Then
                tElektroeroziya = ElektroEroz_Kontur(Material, Kvalitet, S, L, IIf(optKonfig_1_Kontur, 1, 2)) * KolVoPoverhnosteiProhodov(txtKolVo_Kontur)
                tSlesar_elektro = 1.3 * KolVo * ZachistkaZausencev(CInt(Material), L, 4)
                tKontrol_elektro = 1.3 * KolVo * IzmerShtangenCircul(, L)
            End If
        End If
        
    End If

End Sub

Private Sub RaschetVremeni_Dorabotka(Material As Integer)

    txtS_Dorabotka.BackColor = vbWhite
    
    Dim Smax As Double
    Smax = ElektroEroz_Dorabotka_Smax(Material, IIf(optKonfig_1_Dorabotka, 1, 2))
    lblSmax_Dorabotka.Caption = "до " & Smax
    
    Dim S As Double:        S = DblFromCtrl(txtS_Dorabotka)
    Dim L As Double:        L = DblFromCtrl(txtL_Dorabotka)
    Dim KolVo As Double:    KolVo = KolVoPoverhnosteiProhodov(txtKolVo_Dorabotka)
    
    If S > Smax Then
        txtS_Dorabotka.BackColor = vbRed
    Else
        If S > 0 And L > 0 Then
            tElektroeroziya = ElektroEroz_Dorabotka(Material, S, L, IIf(optKonfig_1_Dorabotka, 1, 2)) * KolVoPoverhnosteiProhodov(txtKolVo_Dorabotka)
            tSlesar_elektro = 1.3 * KolVo * ZachistkaZausencev(CInt(Material), L, 4)
            tKontrol_elektro = 1.3 * KolVo * IzmerShtangenCircul(, L)
        End If
    End If
        

End Sub

Private Sub RaschetVremeni_Ustanovka()

    tElektroeroziya = ElektroEroz_Ustanov(cboSloznostUstanovki.ListIndex + 1) * KolVoPoverhnosteiProhodov(txtKolVo_Ustanovka)
    
End Sub

Private Sub cmdAdd_Click()
    Call AddInTable
End Sub

Sub AddInTable()

    Dim Stroka As Integer
    Dim Perehod As String
    Dim Konfig As Integer
    Dim S As Double, L As Double, KolVo As Double
    Dim SlozhnostUstanovki As Integer

    If tElektroeroziya > 0 Then
        
        Select Case mltElektroeroziya.SelectedItem.Name
            Case "pUstanovka":  Perehod = "Установить"
                SlozhnostUstanovki = cboSloznostUstanovki.ListIndex + 1
                KolVo = DblFromCtrl(txtKolVo_Ustanovka)
            Case "pKontur":     Perehod = "Вырезать контур"
                Konfig = IIf(optKonfig_2_Kontur.Value = True, 2, 1)
                S = DblFromCtrl(txtS_Kontur)
                L = DblFromCtrl(txtL_Kontur)
                KolVo = DblFromCtrl(txtKolVo_Kontur)
            Case "pDorabotka":  Perehod = "Доработать паз"
                Konfig = IIf(optKonfig_2_Dorabotka.Value = True, 2, 1)
                S = DblFromCtrl(txtS_Dorabotka)
                L = DblFromCtrl(txtL_Dorabotka)
                KolVo = DblFromCtrl(txtKolVo_Dorabotka)
        End Select
        
        With tbElektroEroziya_Raschot
        
            .ListRows.Add
            Stroka = .ListRows.Count
            
            .DataBodyRange(Stroka, .ListColumns("id").Index) = mltElektroeroziya.SelectedItem.Name
            .DataBodyRange(Stroka, .ListColumns("Переход").Index) = Perehod
            
            If Kvalitet > 0 Then .DataBodyRange(Stroka, .ListColumns("idIT").Index) = Kvalitet
            If Konfig > 0 Then .DataBodyRange(Stroka, .ListColumns("Конфиг.").Index) = Konfig
            If S > 0 Then .DataBodyRange(Stroka, .ListColumns("S, мм").Index) = S
            If L > 0 Then .DataBodyRange(Stroka, .ListColumns("L, мм").Index) = L
            If KolVo > 0 Then .DataBodyRange(Stroka, .ListColumns("Кол-во").Index) = KolVo
            If SlozhnostUstanovki > 0 Then .DataBodyRange(Stroka, .ListColumns("Сложн. уст.").Index) = SlozhnostUstanovki

        End With
        
    End If
    
    Call ZagruzkaIzTablitsy_v_ListBox
    
End Sub

Private Sub ZagruzkaIzTablitsy_v_ListBox()

    wsElektroEroziya.Range("B1") = Material
    
    Dim i As Integer
    Dim Perekhod As String
    Dim tElektro As Double, tSles As Double, tKontrol As Double
    
    lstRaschet.Clear
    txtRazryad.text = Empty
    
    With tbElektroEroziya_Raschot
        For i = 1 To .ListRows.Count
            Perekhod = .DataBodyRange(i, .ListColumns("Переход").Index)
            tElektro = .DataBodyRange(i, .ListColumns("tЭлектро").Index)
            tSles = .DataBodyRange(i, .ListColumns("tСлес").Index)
            tKontrol = .DataBodyRange(i, .ListColumns("tКонтр").Index)
            With lstRaschet
                .AddItem
                .List(.ListCount - 1, 0) = Perekhod
                .List(.ListCount - 1, 1) = Format(tElektro, "0.0")
                .List(.ListCount - 1, 2) = Format(tSles, "0.0")
                .List(.ListCount - 1, 3) = Format(tKontrol, "0.0")
            End With
        Next
    End With
    
    txtRazryad.text = wsElektroEroziya.Range("D1")
    
    Call RaschetTsht

End Sub

Private Sub RaschetTsht()
    
    txtTsht.text = Empty
    
    Dim i As Integer
    Dim Tsht As Double, Tsht_Slesar As Double, Tsht_Kontrol As Double
    
    With lstRaschet
        
        If .ListCount > 0 Then
            For i = 0 To .ListCount - 1
                Tsht = Tsht + CDbl(.List(i, 1))
                Tsht_Slesar = Tsht_Slesar + CDbl(.List(i, 2))
                Tsht_Kontrol = Tsht_Kontrol + CDbl(.List(i, 3))
            Next
            txtTsht.text = OkruglenieTsht(Tsht)
            txtTsht_Slesar.text = OkruglenieTsht(Tsht_Slesar)
            txtTsht_Kontrol.text = OkruglenieTsht(Tsht_Kontrol)
        End If
        
    End With
    
End Sub

Private Sub cmdDelete_Click()

    If lstRaschet.ListIndex = -1 Then Exit Sub
        
    Dim i As Integer
    Dim StrokaUdaleniya As Integer
    
    With lstRaschet
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                StrokaUdaleniya = i + 1
                tbElektroEroziya_Raschot.ListRows(StrokaUdaleniya).Delete
            End If
        Next
        Call ZagruzkaIzTablitsy_v_ListBox
    End With
    
End Sub


Private Sub DostupnostKonfiguratsii()
    
    optKonfig_1_Dorabotka.Enabled = True
    optKonfig_2_Dorabotka.Enabled = True
    optKonfig_1_Kontur.Enabled = True
    optKonfig_2_Kontur.Enabled = True
    
    optKonfig_1_Dorabotka.Value = True
    optKonfig_1_Kontur.Value = True
    
    Select Case Material
        Case 1
            optKonfig_1_Kontur.Value = True
            optKonfig_2_Kontur.Enabled = False
            
            optKonfig_1_Dorabotka.Value = True
            optKonfig_2_Dorabotka.Enabled = False
        Case 2
            Select Case cboKvalitet.ListIndex
                Case 0: optKonfig_1_Kontur.Value = True: optKonfig_2_Kontur.Enabled = False
                Case 3: optKonfig_2_Kontur.Value = True: optKonfig_1_Kontur.Enabled = False
            End Select
    End Select

End Sub
