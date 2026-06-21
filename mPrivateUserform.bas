Attribute VB_Name = "mPrivateUserform"
'Вспомогательные функции для работы с пользовательской формой

Option Explicit
Option Private Module

Public Const SMESHCHENIYE As Integer = 3

Public ArrTxtKontrolVvodaChisel() As New clsmKontrolVvodaChisel
Public iTxt_KontrolVvodaChisel As Long

Public ArrCboKontrolVvodaChisel() As New clsmKontrolVvodaChisel
Public iCbo_KontrolVvodaChisel As Long


Sub RaschotGalvaniki()
    With frmTokFrez
        Select Case .cboMaterial
            Case "Алюминиевые сплавы"
                .cboGalvanika.text = "Ан. Окс. хр"
            Case "Медные сплавы"
                .cboGalvanika.text = "О-Ви"
            Case "Сталь углеродистая"
                .cboGalvanika.text = "Цинкование"
            Case "Сталь нержавеющая"
                .cboGalvanika.text = "Хим. Пас"
        End Select
    End With
End Sub

Sub DobavitVKlass_KontrolVvodaChisel(frm As UserForm)
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If LCase(ctrl.Tag) <> "notclsm" And LCase(ctrl.Tag) <> "notnum" Then
            Select Case TypeName(ctrl)
                Case "TextBox"
                    ReDim Preserve ArrTxtKontrolVvodaChisel(iTxt_KontrolVvodaChisel)
                    Set ArrTxtKontrolVvodaChisel(iTxt_KontrolVvodaChisel).txt = ctrl
                    iTxt_KontrolVvodaChisel = iTxt_KontrolVvodaChisel + 1
            End Select
        End If
    Next
End Sub

Sub DobavitVKlass_KontrolVvodaChisel_Ctrl(ctrl As Control)
    Select Case TypeName(ctrl)
        Case "TextBox"
            ReDim Preserve ArrTxtKontrolVvodaChisel(iTxt_KontrolVvodaChisel)
            Set ArrTxtKontrolVvodaChisel(iTxt_KontrolVvodaChisel).txt = ctrl
            iTxt_KontrolVvodaChisel = iTxt_KontrolVvodaChisel + 1
    End Select
End Sub

Sub UdaleniyeStrokIzLisboksa_i_Tablitsy(lst As MSForms.ListBox, tb As ListObject)

    If lst.ListIndex = -1 Then Exit Sub
        
    Dim i As Integer
    Dim StrokaUdaleniya As Integer
    
    With lst
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                StrokaUdaleniya = i + 1
                tb.ListRows(StrokaUdaleniya).Delete
            End If
        Next
    End With
    
End Sub

Public Sub DobavitVKlass_GlavnyyRaschot(Forma As MSForms.UserForm, ByRef ControlHandlers As Collection, ByRef Handler As clsGlavnyyRaschot)

    Dim ctrl As Control
    
    For Each ctrl In Forma.Controls
    
        If Not LCase(ctrl.Tag) Like "notcls*" Then
        
            Select Case TypeName(ctrl)
                Case "TextBox", "CheckBox", "MultiPage", "OptionButton", "ComboBox"
                    Set Handler = New clsGlavnyyRaschot
                    
                    Select Case TypeName(ctrl)
                        Case "TextBox"
                            Set Handler.TextBoxControl = ctrl
                        Case "CheckBox"
                            Set Handler.CheckBoxControl = ctrl
                        Case "MultiPage"
                            Set Handler.MultiPageControl = ctrl
                        Case "OptionButton"
                            Set Handler.OptionButtonControl = ctrl
                        Case "ComboBox"
                            Set Handler.ComboBoxControl = ctrl
                    End Select
            
                    ControlHandlers.Add Handler
            End Select

            
        End If
    Next
    

End Sub


Sub VyravnitPoLevomuKrayuNazvaniyaVkladok(mlt As MultiPage)

    Dim i As Integer
    
    Dim SamoeDlinnoyeNazvaniye As Integer
    Dim KolVoSimvolov As Integer
    For i = 0 To mlt.Pages.Count - 1
        KolVoSimvolov = Len(mlt.Pages(i).Caption)
        If KolVoSimvolov > SamoeDlinnoyeNazvaniye Then
            SamoeDlinnoyeNazvaniye = KolVoSimvolov
        End If
    Next
    
    Dim tabW As Double
    tabW = 5.8
    
    If mlt.TabFixedWidth > SamoeDlinnoyeNazvaniye * tabW Then SamoeDlinnoyeNazvaniye = CInt(mlt.TabFixedWidth / tabW)
    
    Dim Nazvaniye As String
    For i = 0 To mlt.Pages.Count - 1
        Nazvaniye = mlt.Pages(i).Caption
        KolVoSimvolov = Len(Nazvaniye)
        Nazvaniye = Nazvaniye & Space(SamoeDlinnoyeNazvaniye - KolVoSimvolov)
        mlt.Pages(i).Caption = Nazvaniye
    Next
    
End Sub

Public Sub RazmestitNizhe(Element As Control, ElementVyshe As Control, Optional DopSmeshcheniye As Double = 0)
    Element.Top = ElementVyshe.Top + ElementVyshe.Height + SMESHCHENIYE + DopSmeshcheniye
End Sub

Public Sub RazmestitPravee(Element As Control, ElementSleva As Control, Optional DopSmeshcheniye As Double = 0)
    Element.Top = ElementSleva.Top
    Element.Left = ElementSleva.Left + ElementSleva.Width + SMESHCHENIYE + DopSmeshcheniye
End Sub

Public Sub NastroitRazmerFormy(Forma As Object, NizhniyElemnt As Control, PravyyElement As Control)
    Forma.Width = PravyyElement.Left + PravyyElement.Width + PLUS_SHIRINA
    Forma.Height = NizhniyElemnt.Top + NizhniyElemnt.Height + PLUS_VYSOTA
End Sub

Function DblFromCtrl(ctrl As Control) As Double
    'Преобразование элемента управления в число
    
    If IsNumeric(ctrl) Then
        DblFromCtrl = CDbl(ctrl)
    Else
        DblFromCtrl = 0
    End If
    
End Function

Sub SdelatFonKrasnym(ctrl As Control, MaksimalnoyeZnacheniye As Double)

    If DblFromCtrl(ctrl) > MaksimalnoyeZnacheniye And ctrl <> "" Then
        ctrl.BackColor = vbRed
    Else
        ctrl.BackColor = vbWhite
    End If
End Sub

Sub SdelatShriftZhirnimPriFocuse(frm As UserForm, ctrl As Control, FontBold As Boolean)

    Dim lblPodpis As MSForms.Label
    Set lblPodpis = frm.Controls(Replace(ctrl.Name, Left(ctrl.Name, 3), "lbl"))
    lblPodpis.Font.Bold = FontBold

End Sub

Sub OtkrytFormyVTomzheMeste(frm As Object, oldTop As Double, oldLeft As Double, frmName As String)

    Unload frm
    
    Dim newFrm As Object
    Set newFrm = VBA.UserForms.Add(frmName)
    
    With newFrm
        .startupposition = 0
        .Top = oldTop
        .Left = oldLeft
        .Show
    End With
    
End Sub

Sub VyravnitElementyTablitsy(Elementy, Stroka As Integer, Form As MSForms.UserForm)

    Dim Element
    Dim ctrl As Control
    Dim num As Integer
    
    For Each Element In Elementy
        
        Set ctrl = Element
        
        num = NumControl(ctrl)
        Dim prevCtrl As Control
    
        With ctrl

            Set prevCtrl = Form.Controls(Split(.Name, "_")(0) & "_" & Stroka - 1)
            
            .Top = prevCtrl.Top + prevCtrl.Height + 2
            .Height = prevCtrl.Height
            .Width = prevCtrl.Width
            .Left = prevCtrl.Left
            
            On Error Resume Next

            .ForeColor = prevCtrl.ForeColor
            .SpecialEffect = prevCtrl.SpecialEffect
            .BorderStyle = prevCtrl.BorderStyle
            .BorderColor = prevCtrl.BorderColor
            .List = prevCtrl.List
            .ListIndex = prevCtrl.ListIndex
            .Style = prevCtrl.Style
            .ListWidth = prevCtrl.ListWidth
            .ColumnWidths = prevCtrl.ColumnWidths
            .Caption = prevCtrl.Caption
            .Font.Name = prevCtrl.Font.Name
            .Font.Size = prevCtrl.Font.Size
            .Enabled = prevCtrl.Enabled
            .Picture = prevCtrl.Picture
            .BackColor = prevCtrl.BackColor
            .BackStyle = prevCtrl.BackStyle
            .Locked = prevCtrl.Locked
            .TextAlign = prevCtrl.TextAlign
            .SelectionMargin = prevCtrl.SelectionMargin
            .TabIndex = 2 * prevCtrl.TabIndex + 1
           
           On Error GoTo 0

        End With
    Next

    Form.Repaint
    
End Sub

Function NumControl(ctrl As Control) As Integer
    If ctrl.Name Like "*_#" Or ctrl.Name Like "*_##" Then
        NumControl = CInt(Split(ctrl.Name, "_")(1))
    End If
End Function

Sub ProverkaNaMaksimalnoyeZnacheniye(txt As MSForms.TextBox, MaksimalnoyeZnacheniye As Double, frm As UserForm)
    
    If DblFromCtrl(txt) > MaksimalnoyeZnacheniye Then
        txt.BorderColor = vbRed
    Else
        txt.BorderColor = vbActiveBorder
    End If
    
    frm.Repaint

End Sub

Function EnterTextBox(txt As MSForms.TextBox)

    If Not IsNumeric(txt) Then
        txt.text = Empty
    End If
    txt.ForeColor = &H404040
    
End Function

Sub ZapolnitListBoksIzMassiva(MassivNaimOperTpzTsh As Variant, lstTablitsa As MSForms.ListBox)
    
    If Not IsArray(MassivNaimOperTpzTsh) Then Exit Sub

    If lstTablitsa.ColumnCount < 3 Then lstTablitsa.ColumnCount = 3
    
    Dim i As Integer, NaimenovaniyeOperatsii As String, Tpz As Integer, Tsht As Double
    
    For i = LBound(MassivNaimOperTpzTsh) To UBound(MassivNaimOperTpzTsh)
        
        NaimenovaniyeOperatsii = CStr(MassivNaimOperTpzTsh(i)(0))
        If Not NaimenovaniyeOperatsii = "Масса" Then
        
            Tpz = CInt(MassivNaimOperTpzTsh(i)(1))
            Tsht = CDbl(MassivNaimOperTpzTsh(i)(2))
                
            Call DobavleniyeStrokVListBox(lstTablitsa, NaimenovaniyeOperatsii, Tpz, Tsht)
            
        End If
        
    Next
    
End Sub

Sub DobavleniyeStrokVListBox(lstTablitsa As MSForms.ListBox, NaimenovaniyeOperatsii As String, Tpz As Integer, Tsht As Double)

    With lstTablitsa
        .AddItem
        .List(.ListCount - 1, 0) = NaimenovaniyeOperatsii
        .List(.ListCount - 1, 1) = Tpz
        .List(.ListCount - 1, 2) = CStr(OkruglenieTsht(Tsht))
    End With
    
End Sub

Function ExitTextBox(txt As MSForms.TextBox)
    If Not IsNumeric(txt) Then
        txt.text = txt.Tag
        txt.ForeColor = &H80000011
    End If
End Function

Sub DobavitMaterialyVListBox(SpisokMaterialov, cbo As MSForms.ComboBox)
    
    Dim i As Integer
    With cbo
        If .ColumnCount < 2 Then .ColumnCount = 2
        For i = LBound(SpisokMaterialov) To UBound(SpisokMaterialov)
            .AddItem
            .List(.ListCount - 1, 0) = NazvaniyeMateriala(CInt(SpisokMaterialov(i)))
            .List(.ListCount - 1, 1) = SpisokMaterialov(i)
        Next
        .ListIndex = 0
        .ColumnWidths = "; 0pt; 0pt"
    End With
    
End Sub

