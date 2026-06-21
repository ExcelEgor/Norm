VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSlovarSborki 
   Caption         =   "Сборка"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6165
   OleObjectBlob   =   "frmSlovarSborki.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSlovarSborki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim KolVoNovykhZapisey As Integer
Dim tbKomplektovaniye As ListObject
Dim tbSlovar As ListObject


Private Sub cdmAdd_Click()
    Call DobovitVShablon
End Sub
Private Sub ZagruzkaNulevyhZnacheniy()
    
    txtNaimenovaniye.text = Empty
    txtVremya.text = Empty
    
    Dim i As Long
    With tbKomplektovaniye
        For i = 1 To .ListRows.Count
            If .DataBodyRange(i, .ListColumns("Есть в словаре").Index) = False Then
                txtNaimenovaniye = .DataBodyRange(i, .ListColumns("Наименование").Index)
                Exit For
            End If
        Next
    End With
    
    txtVremya.SetFocus
    
End Sub

Private Sub DobovitVShablon()

    If IsNumeric(txtVremya) And txtNaimenovaniye <> "" Then
    
        If WorksheetFunction.CountIf(tbSlovar.ListColumns(1).DataBodyRange, txtNaimenovaniye) > 0 Then
        
            MsgBox txtNaimenovaniye & " уже есть словаре. Запись не добавлена!", vbExclamation, "Ошибка"
            
        Else
    
            Application.ScreenUpdating = False
            
            Dim Stroka As Long
            With tbSlovar
                .ListRows.Add
                Stroka = .ListRows.Count
                .DataBodyRange(Stroka, 1) = txtNaimenovaniye
                .DataBodyRange(Stroka, 2) = CDbl(txtVremya)
            End With
            
            Dim i As Long
            With tbKomplektovaniye
                For i = 1 To .ListRows.Count
                    .DataBodyRange(i, 6).FormulaLocal = .DataBodyRange(i, 6).FormulaLocal
                    .DataBodyRange(i, 7).FormulaLocal = .DataBodyRange(i, 7).FormulaLocal
                Next
            End With
            
            Application.Calculate
            
            Application.ScreenUpdating = True
            
            MsgBox "Добавлена новая запись!", vbInformation, "Новая запись"
            
            KolVoNovykhZapisey = KolVoNovykhZapisey + 1
            
            Call ZagruzkaNulevyhZnacheniy
            
        End If
        
    End If
    
End Sub

Private Sub UserForm_Initialize()

    KolVoNovykhZapisey = 0
    
    Set tbSlovar = wsSborka.ListObjects("tbSborka")
    Set tbKomplektovaniye = ActiveSheet.ListObjects(1)
    
    Call ZagruzkaNulevyhZnacheniy
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If KolVoNovykhZapisey > 0 Then
        ThisWorkbook.Save
    End If
End Sub


