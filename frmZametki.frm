VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZametki 
   Caption         =   "Заметки"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10470
   OleObjectBlob   =   "frmZametki.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmZametki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const PutKZametkam As String = "\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\Для надстройки\Заметки.xlsx"
Dim wbZametki As Workbook
Dim tbZametki As ListObject
Dim ArrZametki

Private Sub UserForm_Initialize()
    
    Call OtkrytKnigu
    Call ZagruzkaVListBox
    
    wbZametki.Close (False)
    
End Sub
Private Sub OtkrytKnigu()

    If KnigaOtkryta = False Then
        Set wbZametki = Workbooks.Open(PutKZametkam)
        wbZametki.Windows(1).Visible = False
        Set tbZametki = wbZametki.Worksheets("Заметки").ListObjects(1)
    End If
    
End Sub

Private Sub ZagruzkaVListBox()

    lstZametki.Clear
    Dim i As Long
    For i = 1 To tbZametki.ListRows.Count
        lstZametki.AddItem tbZametki.DataBodyRange(i, 1)
    Next
    
    ArrZametki = tbZametki.DataBodyRange.Value
    
End Sub

Private Sub cmdNovayaZametka_Click()
    
    With lstZametki
        .AddItem
        .ListIndex = lstZametki.ListCount - 1
        .Enabled = False
    End With
    
    txtZametki.text = Empty
    
    cmdNovayaZametka.Enabled = False
    cmdSokhranit.Enabled = True
    
    Call VkluchitRezhimRedaktirovaniya
    
End Sub

Private Sub cmdUdalit_Click()

    If lstZametki.ListIndex <> -1 Then
    
        Call OtkrytKnigu
        
        tbZametki.ListRows(lstZametki.ListIndex + 1).Delete
        lstZametki.RemoveItem (lstZametki.ListIndex)
        cmdSokhranit.Enabled = True
        
        Call ZagruzkaVListBox
        
    End If
    
End Sub
Private Sub cmdSokhranit_Click()

    cmdNovayaZametka.Enabled = True
    cmdSokhranit.Enabled = False
    lstZametki.Enabled = True
    
    Call VykluchitRezhimRedaktirovaniya
    
    Call OtkrytKnigu
    
    If txtGruppa.text <> Empty And txtZametki.text <> Empty Then
        Call DobavitZapisVTablitsu
    End If
    
    Call ZagruzkaVListBox
    
    wbZametki.Close (True)
    
End Sub
Private Sub DobavitZapisVTablitsu()

    Dim Stroka As Long
    With tbZametki
        .ListRows.Add
        Stroka = .ListRows.Count
        .DataBodyRange(Stroka, 1) = txtGruppa.text
        .DataBodyRange(Stroka, 2) = txtZametki.text
    End With

End Sub

Private Sub lstZametki_Change()

    txtGruppa.text = Empty
    txtZametki.text = Empty
    
    If lstZametki.ListIndex <> -1 And lstZametki.ListIndex < UBound(ArrZametki) Then
        txtGruppa.text = ArrZametki(lstZametki.ListIndex + 1, 1)
        txtZametki.text = ArrZametki(lstZametki.ListIndex + 1, 2)
    End If
    
End Sub

Private Sub VkluchitRezhimRedaktirovaniya()
    Call SvoystvaTekstBoksa(txtGruppa, True)
    Call SvoystvaTekstBoksa(txtZametki, True)
End Sub

Private Sub VykluchitRezhimRedaktirovaniya()
    Call SvoystvaTekstBoksa(txtGruppa, False)
    Call SvoystvaTekstBoksa(txtZametki, False)
End Sub

Private Sub SvoystvaTekstBoksa(txt As MSForms.TextBox, RezhimRedaktirovaniya As Boolean)
    If RezhimRedaktirovaniya Then
        txt.Locked = False
        txt.SpecialEffect = fmSpecialEffectSunken
    Else
        txt.Locked = True
        txt.SpecialEffect = fmSpecialEffectEtched
    End If
End Sub

Private Function KnigaOtkryta() As Boolean
    
    KnigaOtkryta = False

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.FullName = PutKZametkam Then
            KnigaOtkryta = True
            Exit For
        End If
    Next
    
End Function
