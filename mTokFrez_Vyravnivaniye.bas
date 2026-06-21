Attribute VB_Name = "mTokFrez_Vyravnivaniye"
Option Explicit
Option Private Module

Sub VyravnitElementyFormy_TokFrez()
    
    Call VyravnitPoLevomuKrayuNazvaniyaVkladok
    
    Call VyravnitElementyFormy

    Call VyravnitElementyVnutriKarty
    
End Sub

Private Sub VyravnitPoLevomuKrayuNazvaniyaVkladok()

    Dim i As Integer, j As Integer
    Dim mlt As MSForms.MultiPage
    Dim KolVoSimvolov As Integer, SamoeDlinnoyeNazvaniye As Integer
    
    For i = 0 To frmTokFrez.mltMain.Pages.Count - 1
    
        Select Case frmTokFrez.mltMain.Pages(i).Name
            Case "pTok"
                Set mlt = frmTokFrez.mltNormativTok
            Case "pRastFrez"
                Set mlt = frmTokFrez.mltNormativRastFrez
        End Select
        
        For j = 0 To mlt.Pages.Count - 1
            KolVoSimvolov = Len(mlt.Pages(j).Caption)
            If KolVoSimvolov > SamoeDlinnoyeNazvaniye Then
                SamoeDlinnoyeNazvaniye = KolVoSimvolov
            End If
        Next
        
    Next
    
    Dim Nazvaniye As String
    For i = 0 To frmTokFrez.mltMain.Pages.Count - 1
        Select Case frmTokFrez.mltMain.Pages(i).Name
            Case "pTok"
                Set mlt = frmTokFrez.mltNormativTok
            Case "pRastFrez"
                Set mlt = frmTokFrez.mltNormativRastFrez
        End Select
        For j = 0 To mlt.Pages.Count - 1
            Nazvaniye = mlt.Pages(j).Caption
            KolVoSimvolov = Len(Nazvaniye)
            Nazvaniye = Nazvaniye & Space(SamoeDlinnoyeNazvaniye - KolVoSimvolov)
            mlt.Pages(j).Caption = Nazvaniye
        Next
    Next
    
End Sub

Private Sub VyravnitElementyFormy()

    With frmTokFrez

        'Основные группы по левому краю
        Dim LevyKray As Integer
        LevyKray = 3
        .fraMainMenu.Left = LevyKray
        .fraForMltMain.Left = LevyKray
        .fraPerekhod.Left = LevyKray
        .fraCmdAdd.Left = LevyKray
        
        'Вспомогательный расчёт по правому краю
        .framVspomogatRaschety.Left = frmTokFrez.Width - .framVspomogatRaschety.Width - 15
        
        'Расчётную таблицу к левому краю вспомогательных расчётов
        .fraRaschot.Left = .framVspomogatRaschety.Left - .fraRaschot.Width - 3
        
        'Ширина меню до левого края вспогательных расчётов
        .fraMainMenu.Width = .framVspomogatRaschety.Left - .fraMainMenu.Left - 3
        
        'Ширина основных групп до левого края расчётной таблицы
        Dim Shirina As Integer
        Shirina = .fraRaschot.Left - LevyKray - 3
        .fraForMltMain.Width = Shirina
        .fraPerekhod.Width = Shirina
        .fraCmdAdd.Width = Shirina
        
        'Выравнивание по верхнему краю
        Dim VerkhniyKray As Integer
        
        VerkhniyKray = 3
        .fraMainMenu.Top = VerkhniyKray
        .framVspomogatRaschety.Top = VerkhniyKray
        
        .fraForMltMain.Top = .fraMainMenu.Top + .fraMainMenu.Height + 6
        .fraRaschot.Top = .fraMainMenu.Top + .fraMainMenu.Height + 6
        
        'Выравнить .fraCmdAdd по нижнему краю форму
        .fraCmdAdd.Top = frmTokFrez.Height - .fraCmdAdd.Height - 30
        
        'Разместить .fraPerekhod выше .fraCmdAdd
        .fraPerekhod.Top = .fraCmdAdd.Top - .fraPerekhod.Height - 3
        
        'Выравнить высоту .fraForMltMain до .fraPerekhod
        .fraForMltMain.Height = .fraPerekhod.Top - .fraForMltMain.Top - 3
        
        'Выравнить высоту .fraRaschot и .framVspomogatRaschety до .fraCmdAdd
        .framVspomogatRaschety.Height = .fraCmdAdd.Top + .fraCmdAdd.Height - .framVspomogatRaschety.Top
        .fraRaschot.Height = .fraCmdAdd.Top + .fraCmdAdd.Height - .fraRaschot.Top
        
        'Вырывнить высоту lstTokarnoFrezer по нижнему краю .fraPerekhod
        .lstTokarnoFrezer.Height = .fraPerekhod.Top + .fraPerekhod.Height - .fraRaschot.Top - .lstTokarnoFrezer.Top
        
        'Выравнить .fraTsht по нижнему краю lstTokarnoFrezer
        .fraTsht.Top = .lstTokarnoFrezer.Top + .lstTokarnoFrezer.Height + 3
        
        'Выравнить cmdRaschotTpz ниже .fraTsht
        .cmdRaschotTpz.Top = .fraTsht.Top + .fraTsht.Height - 1
        
        With .mltMain
            .Left = 0
            .Top = frmTokFrez.tabForMlt.Top + frmTokFrez.tabForMlt.Height + 3
            .Width = frmTokFrez.fraForMltMain.Width
            .Height = frmTokFrez.fraForMltMain.Height
            .Style = fmTabStyleNone
        End With
        
        With frmTokFrez.fraObshchParam_Tok
            .Top = 3
            .Left = 3
            .Width = frmTokFrez.mltMain.Width - 11
        End With
    
    End With
    
End Sub

Private Sub VyravnitElementyVnutriKarty()

    Dim ArrMltNormativy
    With frmTokFrez
        ArrMltNormativy = Array(.mltNormativTok, .mltNormativRastFrez)
    End With
    
    Dim i As Integer, mltNormativ As Control
    For i = LBound(ArrMltNormativy) To UBound(ArrMltNormativy)
        Set mltNormativ = ArrMltNormativy(i)
        
        With mltNormativ
            .Top = 39
            .Left = 0
            .Height = frmTokFrez.mltMain.Height - .Top
            .Width = frmTokFrez.mltMain.Width
        End With
        
    Next

    VyravnitZagolovki
    
End Sub

Private Sub VyravnitZagolovki()

    Dim MltWidth As Double
    MltWidth = frmTokFrez.mltMain.Width - 14
    
    Dim ctrl As Control, img As Control
    For Each ctrl In frmTokFrez.Controls
        With ctrl
            If .Name Like "lblHead*" Then
                .Top = 3
                .Left = 3
                .Width = frmTokFrez.mltMain.Width - .Left - 9 - frmTokFrez.mltNormativTok.TabFixedWidth
                .Font.Name = "Segoe UI Semibold"
                .Font.Size = 10
                .Height = 18
                .ForeColor = &H404040
                
                On Error Resume Next
                    Set img = frmTokFrez.Controls(Replace(ctrl.Name, "lblHead", "img"))
                    img.Left = ctrl.Left + ctrl.Width - img.Width - 27
                On Error GoTo 0
            ElseIf .Name Like "lblMainHead_*" Then
                .Left = 6
                .Width = MltWidth
            End If
        End With
    Next
    
End Sub
