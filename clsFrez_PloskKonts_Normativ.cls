VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFrez_PloskKonts_Normativ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Frezerovaniye

Option Explicit

Public Function RashotNorm(ctx As clsMillingContext) As Double
    
    If ctx.TipStanka <> GorizontalnoRastochnoy And ctx.TipStanka <> VertikalnoGorizontalnoFrezernyy Or _
        MaterialKorrektnyy_RastFrez(ctx.Material) = False Or ctx.Ra < MIN_RA_FREZ Or ctx.Pripusk <= 0 Or ctx.ShirinaPoverkhnosti <= 0 Or ctx.DlinaPoverkhnosti <= 0 Then
        Exit Function
    End If
    
    Dim hChist As Double
    If ctx.Ra <= 10 Then
        hChist = WorksheetFunction.Min(IIf(ctx.Ra > 2.5, 2, 1), ctx.Pripusk)
    Else
        hChist = 0
    End If

    Dim Bfrez As Double
    Bfrez = ShirinaFrezerovaniya_KontsevyyeFrezy(ctx.ShirinaPoverkhnosti)

    Dim iB As Long
    iB = WorksheetFunction.RoundUp(ctx.ShirinaPoverkhnosti / Bfrez, 0)
    
    Dim tChist As Double
    If hChist > 0 Then
        tChist = IIf(ctx.IT <= 11, 2, 1) * iB * VremyaNaPokhod(ctx.Material, ctx.TipStanka, ctx.DlinaPoverkhnosti, ctx.Ra, Bfrez, hChist)
    End If
    
    Dim hChern As Double
    hChern = ctx.Pripusk - hChist
    
    Dim tChernKork As Double, tChern As Double
    
    If hChern > 0 Then
        Dim GlubinaRez As Double
        GlubinaRez = GlubinaRez_KontsevyyeFrezy(hChern)
        
        Dim iH As Long
        iH = WorksheetFunction.RoundUp(hChern / GlubinaRez, 0)
        
        Dim tChernNaProkhod As Double
        tChernNaProkhod = VremyaNaPokhod(ctx.Material, ctx.TipStanka, ctx.DlinaPoverkhnosti, 20, Bfrez, GlubinaRez)
        
        If ctx.Korka Then
            tChernKork = K_KORKA_FREZ_PLOSK * iB * tChernNaProkhod
            iH = iH - 1
        End If
        
        If iH > 0 Then
            tChern = iH * iB * tChernNaProkhod
        End If
    End If
    
    RashotNorm = RaschotObshchegoKoeffitsienta_RastFrez(ctx.Material, ctx.TipStanka, ctx.NaUdar, ctx.TonkostennayaKonfiguraciya) * (tChernKork + tChern + tChist)
    
End Function

Private Function VremyaNaPokhod(Material As EnumMaterialy, TipStanka As EnumTipStankaRastFrez, DlinaPoverkhnosti As Double, Ra As Double, ShirinaFrezerovaniya As Double, GlubinaRez As Double) As Double
    'Фрезерование плоскостей концевыми фрезами (карты 24, 25)
        
    Dim TipStanka_DlyaPoiska As EnumTipStankaRastFrez
    TipStanka_DlyaPoiska = IzmenitTipStanka(TipStanka, Material)
    
    Dim Material_DlyaPoiska As EnumMaterialy
    Material_DlyaPoiska = IzmenitMaterialDlyaZaprosa_RastFrez(Material, False)
    
    Dim Normativ
    Normativ = ZagruzitNormativ("frez_kontsevyye")

    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = TipStanka_DlyaPoiska And Normativ(i, 2) = Material_DlyaPoiska And _
            Normativ(i, 3) >= ShirinaFrezerovaniya And Normativ(i, 4) <= Ra And Normativ(i, 5) >= GlubinaRez Then
            a = Normativ(i, 6)
            b = Normativ(i, 7)
            Exit For
        End If
    Next

    VremyaNaPokhod = a * DlinaPoverkhnosti + b
End Function

Private Function ShirinaFrezerovaniya_KontsevyyeFrezy(ShirinaPoverkhnosti As Double) As Integer
    Select Case ShirinaPoverkhnosti
        Case Is <= 10: ShirinaFrezerovaniya_KontsevyyeFrezy = 10
        Case Is <= 20: ShirinaFrezerovaniya_KontsevyyeFrezy = 20
        Case Is <= 30: ShirinaFrezerovaniya_KontsevyyeFrezy = 30
        Case Else
            With WorksheetFunction
                If .RoundUp(ShirinaPoverkhnosti / 30, 0) = .RoundUp(ShirinaPoverkhnosti / 40, 0) And ShirinaPoverkhnosti / 30 <> Int(ShirinaPoverkhnosti / 30) Then
                    ShirinaFrezerovaniya_KontsevyyeFrezy = 30
                Else
                    ShirinaFrezerovaniya_KontsevyyeFrezy = 40
                End If
            End With
    End Select
End Function

Private Function GlubinaRez_KontsevyyeFrezy(Pripusk As Double) As Double
    Select Case Pripusk
        Case Is <= 3:   GlubinaRez_KontsevyyeFrezy = 3
        Case Is <= 5:   GlubinaRez_KontsevyyeFrezy = 5
        Case Is <= 7:   GlubinaRez_KontsevyyeFrezy = 7
        Case Else:
            If WorksheetFunction.RoundUp(Pripusk / 7, 0) = _
                WorksheetFunction.RoundUp(Pripusk / 10, 0) And _
                Pripusk / 7 <> Int(Pripusk / 7) Then
                GlubinaRez_KontsevyyeFrezy = 7
            Else
                GlubinaRez_KontsevyyeFrezy = 10
            End If
    End Select
End Function

