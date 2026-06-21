VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKniga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWbRaschot As Workbook
Private pWsShablonTablisty As Worksheet

Public Property Get wbRaschot() As Workbook: Set wbRaschot = pWbRaschot: End Property

Public Sub Init(WsShablonTablisty As Worksheet)
    Set pWsShablonTablisty = WsShablonTablisty
End Sub

Public Sub NovyyRaschot()
    
    If Not pWbRaschot Is Nothing Then pWbRaschot.Close (False)
    
    pWsShablonTablisty.Copy
    Set pWbRaschot = ActiveWorkbook
    pWbRaschot.Windows(1).Visible = False

End Sub

Public Sub Sokhranit()

    If pWbRaschot.Path = "" Then
        Call SokhranitKak
    Else
        pWbRaschot.Save
    End If
    
End Sub

Public Sub SokhranitKak()

    Dim FilePath As String
    FilePath = Application.GetSaveAsFilename( _
        InitialFileName:="Расчёт.xlsx", _
        fileFilter:="Excel Files (*.xlsx), *.xlsx", _
        Title:="Сохранить книгу как")

    If FilePath = "False" Then Exit Sub
    
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase(wb.Name) = LCase(Dir(FilePath)) Then
            MsgBox "Файл с таким именем уже открыт. Выберите другое имя или закройте файл.", vbExclamation
            wb.Windows(1).Visible = True
            Exit Sub
        End If
    Next wb
    
    Application.DisplayAlerts = False
    pWbRaschot.Windows(1).Visible = True
    
    pWbRaschot.SaveAs fileName:=FilePath, FileFormat:=xlOpenXMLWorkbook
    
    Application.DisplayAlerts = True
    
    pWbRaschot.Windows(1).Visible = False
    
End Sub

Public Sub OtkrytRaschot()

    Dim ImyaKnigi As String
    ImyaKnigi = Application.GetOpenFilename
    
    If ImyaKnigi = "False" Then Exit Sub
    
    If Not pWbRaschot Is Nothing Then pWbRaschot.Close (False)
    
    Set pWbRaschot = Workbooks.Open(ImyaKnigi)
    pWbRaschot.Windows(1).Visible = False
    
End Sub

Public Sub DobavitList()

    Application.ScreenUpdating = False
    
    With pWbRaschot
        .Windows(1).Visible = True
        .pWsShablonTablisty.Copy after:=.Worksheets(.Worksheets.Count)
        .Windows(1).Visible = False
    End With
    
    Application.ScreenUpdating = True

End Sub

Public Sub KopirovatList(NomerLista As Integer)

    Application.ScreenUpdating = False
    
    With pWbRaschot
        .Windows(1).Visible = True
        .Worksheets(NomerLista).Copy after:=.Worksheets(.Worksheets.Count)
        .Windows(1).Visible = False
    End With
    
    Application.ScreenUpdating = True

End Sub

Public Sub UdalitList(NomerLista As Integer)

    If NomerLista < 1 Then Exit Sub

    Application.DisplayAlerts = False
    pWbRaschot.Worksheets(NomerLista).Delete
    Application.DisplayAlerts = True

End Sub
