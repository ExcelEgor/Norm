Attribute VB_Name = "mRabotaSBazoyDannykh"
Option Explicit
Option Private Module

Private BaseCache As Object

Private pTB_MATERIALY As ListObject
Private pTB_NORMATIVY As ListObject

Public Property Get TB_MATERIALY() As ListObject
    If pTB_MATERIALY Is Nothing Then
        Set pTB_MATERIALY = wsMaterial.ListObjects("tbMainMaterial")
    End If
    Set TB_MATERIALY = pTB_MATERIALY
End Property

Private Property Get TB_NORMATIVY() As ListObject
    If pTB_NORMATIVY Is Nothing Then
        Set pTB_NORMATIVY = ThisWorkbook.Worksheets("Нормативы").ListObjects("tbNormativy")
    End If
    Set TB_NORMATIVY = pTB_NORMATIVY
End Property

Public Function ZagruzitNormativ(ByVal Индентификатор As String) As Variant

    If BaseCache Is Nothing Then Set BaseCache = CreateObject("Scripting.Dictionary")
    
    If Not BaseCache.Exists(Индентификатор) Then
        Dim ImyaKnigi As String, ImyaLista As String
        
        On Error Resume Next
        ImyaKnigi = WorksheetFunction.VLookup(Индентификатор, TB_NORMATIVY.DataBodyRange, 2, False)
        ImyaLista = WorksheetFunction.VLookup(Индентификатор, TB_NORMATIVY.DataBodyRange, 3, False)
        On Error GoTo 0
        
        If ImyaKnigi = "" Then
            MsgBox "Ключи " & Индентификатор & " не найден в таблице нормативов tbNormativy", vbCritical
            Exit Function
        End If
        
        BaseCache(Индентификатор) = ZapisatTablitsuVMassiv(ImyaKnigi, ImyaLista)

    End If
    
    ZagruzitNormativ = BaseCache(Индентификатор)
    
End Function

Private Function ZapisatTablitsuVMassiv(ByVal ImyaKnigi As String, ByVal ImyaLista As String)

    Dim Connection As Object, RecordSet As Object
    Set Connection = CreateObject("ADODB.Connection")
    
    Connection.Open "Provider=Microsoft.ACE.OLEDB.12.0.;Data Source=" & PoluchitPutKNormativam & ImyaKnigi & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"

    Set RecordSet = CreateObject("ADODB.Recordset")
    RecordSet.Open "SELECT * FROM [" & ImyaLista & "$]", Connection
    
    If Not RecordSet.EOF Then
        ZapisatTablitsuVMassiv = TransposeArray(RecordSet.GetRows())
    Else
        ZapisatTablitsuVMassiv = Empty
    End If
    
    RecordSet.Close
    Connection.Close
    
End Function

Public Function PoluchitPutKNormativam() As String
    If Environ("UserName") = "hardf" Then
        PoluchitPutKNormativam = "E:\Работа Аврора\Расчёт норм\Для надстройки\Нормативы БД\"
    Else
        PoluchitPutKNormativam = "\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\Для надстройки\Нормативы БД\"
    End If
End Function

Public Function PoluchitPutKShablonam() As String
    If Environ("UserName") = "hardf" Then
        PoluchitPutKShablonam = "E:\Работа Аврора\Расчёт норм\Для надстройки\Шаблоны\"
    Else
        PoluchitPutKShablonam = "\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\Для надстройки\Шаблоны\"
    End If
End Function

Public Function PoslednyayaStrokaVMassive(ByRef Normativ As Variant, ByVal i As Long, ByVal StolbetsSravneniya As Integer) As Boolean
    PoslednyayaStrokaVMassive = False
    If i = UBound(Normativ) Then
        PoslednyayaStrokaVMassive = True
    ElseIf Normativ(i, StolbetsSravneniya) <> Normativ(i + 1, StolbetsSravneniya) Then
         PoslednyayaStrokaVMassive = True
    End If
End Function

Private Function TransposeArray(ByVal SourceArray As Variant) As Variant
    Dim X As Long, Y As Long
    Dim TargetArray As Variant
    
    ' Исходный массив из GetRows всегда 0-базовый: (0 To Сolumns, 0 To Rows)
    ' Новый массив делаем строго 1-базовым: (1 To Rows + 1, 1 To Columns + 1)
    ReDim TargetArray(1 To UBound(SourceArray, 2) + 1, 1 To UBound(SourceArray, 1) + 1)
    
    For X = 0 To UBound(SourceArray, 1)
        For Y = 0 To UBound(SourceArray, 2)
            TargetArray(Y + 1, X + 1) = SourceArray(X, Y)
        Next Y
    Next X
    
    TransposeArray = TargetArray
    
End Function
