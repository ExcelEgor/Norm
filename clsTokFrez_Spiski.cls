VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTokFrez_Spiski"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ZapolnittSpisok_KharakterVyverki_Tokarnyy()

    With frmTokFrez
        .cboKharakterVyverki_Tok.Clear
        If .lstSposobUst_Tok.ListIndex <> -1 Then
            Select Case .lstSposobUst_Tok.ListIndex
                Case 0
                    .cboKharakterVyverki_Tok.List = Split("Без выверки;С выверкой по диаметру и торцу", ";")
                Case Else
                    .cboKharakterVyverki_Tok.List = Split("Без выверки;С выверкой по диаметру;С выверкой по диаметру и торцу", ";")
            End Select
            .cboKharakterVyverki_Tok.ListIndex = .cboKharakterVyverki_Tok.ListCount - 1
        End If
    End With
    
End Sub

Public Sub ZapolnittSpisok_TochnostVyverki_Tokarnyy()

    With frmTokFrez
        .cboTochnostVyverki_Tok.Clear
        If .cboKharakterVyverki_Tok.ListIndex <> -1 Then
            If .cboKharakterVyverki_Tok.ListIndex = 0 Then
                .cboTochnostVyverki_Tok.List = Split("-", ";")
            Else
                Select Case .lstSposobUst_Tok.ListIndex
                    Case 0
                        .cboTochnostVyverki_Tok.List = Split("0,5;0,1", ";")
                    Case Else
                        .cboTochnostVyverki_Tok.List = Split("1,0;0,5;0,1;0,05;0,01", ";")
                End Select
            End If
            .cboTochnostVyverki_Tok.ListIndex = .cboTochnostVyverki_Tok.ListCount - 1
        End If
    End With
    
End Sub

Public Sub ZapolnittSpisok_TochnostVyverki_Frezernyy()

    Dim rSposobyUstanovkiRastFrez As Range
    Set rSposobyUstanovkiRastFrez = wsUst_Frez.Range("UstFrez")

    Dim i As Integer
    With frmTokFrez.cboTochnostVyverki_RastFrez
        .Clear
        For i = 1 To rSposobyUstanovkiRastFrez.Rows.Count
            If rSposobyUstanovkiRastFrez(i, 1) = frmTokFrez.lstSposobUst_RastFrez And rSposobyUstanovkiRastFrez(i, 2) = frmTokFrez.cboKharakterVyverki_RastFrez Then
                .AddItem IIf(rSposobyUstanovkiRastFrez.Cells(i, 3) = 0, "-", CStr(rSposobyUstanovkiRastFrez.Cells(i, 3)))
            End If
        Next
        .ListIndex = .ListCount - 1
    End With
    
End Sub

Public Sub ZapolnittSpisok_KharakterVyverki_Frezernyy()

    Dim rSposobyUstanovkiRastFrez As Range
    Set rSposobyUstanovkiRastFrez = wsUst_Frez.Range("UstFrez")

    Dim i As Integer
    With frmTokFrez.cboKharakterVyverki_RastFrez
        .Clear
        .AddItem
        For i = 1 To rSposobyUstanovkiRastFrez.Rows.Count
            If rSposobyUstanovkiRastFrez(i, 1) = frmTokFrez.lstSposobUst_RastFrez Then
                If .ListCount = 1 Or .List(.ListCount - 1, 0) <> rSposobyUstanovkiRastFrez.Cells(i, 2) Then
                    .AddItem
                    .List(.ListCount - 1, 0) = rSposobyUstanovkiRastFrez.Cells(i, 2)
                    .List(.ListCount - 1, 1) = WorksheetFunction.VLookup(.List(.ListCount - 1, 0), wsUst_Frez.Range("UstFrez_Vyverka"), 2, 0)
                End If
            End If
        Next
        .RemoveItem (0)
        .ListIndex = .ListCount - 1
    End With
    
End Sub



