VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPlankiSNadpisyu 
   Caption         =   "Планки с надписью"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13350
   OleObjectBlob   =   "frmPlankiSNadpisyu.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmPlankiSNadpisyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ImagePaths() As String
Private ImageCount As Integer
Private CurrentIndex As Integer

Private Sub cmdLeft_Click()
    If ImageCount > 0 Then
        If CurrentIndex > 1 Then
            CurrentIndex = CurrentIndex - 1
            ShowImage
        End If
    End If
End Sub

Private Sub cmdRight_Click()
    If ImageCount > 0 Then
        If CurrentIndex <= ImageCount Then
            CurrentIndex = CurrentIndex + 1
            ShowImage
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Dim folderPath As String
    Dim fileName As String
    Dim Ext As String
    
    folderPath = PoluchitPutKNormativam & "Планки с надписью\"
    
    ImageCount = 0
    
    fileName = Dir(folderPath & "*.*")
    
    Do While fileName <> ""
        Ext = LCase(Right(fileName, 4))
        If Ext = ".jpg" Or Ext = ".bmp" Or Ext = ".gif" Or Ext = ".emf" Then
            ImageCount = ImageCount + 1
            ReDim Preserve ImagePaths(1 To ImageCount)
            ImagePaths(ImageCount) = folderPath & fileName
        End If
        fileName = Dir()
    Loop
    
    If ImageCount > 0 Then
        CurrentIndex = 1
        ShowImage
    Else
        MsgBox "В указанной папке нет картинок (.jpg, .bmp, .gif, .emf)!", vbExclamation
        cmdLeft.Enabled = False
        cmdRight.Enabled = False
    End If
    
End Sub

Private Sub ShowImage()

    Image1.Picture = LoadPicture(ImagePaths(CurrentIndex))
    lblCount.Caption = CurrentIndex & " из " & ImageCount
    
    Dim fileName
    fileName = Split(ImagePaths(CurrentIndex), "\")
    fileName = fileName(UBound(fileName))
    fileName = Left(fileName, Len(fileName) - 4)
    lblOboznacheniye.Caption = fileName

End Sub
