Attribute VB_Name = "mFrez_Kontsevyye"
'@Folder Frezerovaniye

Option Explicit

Public Function FrezKontsevymiFrezami_Kompleks(ctx As clsMillingContext) As Variant

    If Not ctx.IsValid Then
        Exit Function
    End If
    
    Dim tFrez As Double
    
    If ctx.PoRezhimam Then
        If ctx.AvtoRezhimy Then
            tFrez = RaschotNormPoRezhimam_AvtoRezhimy(ctx)
        Else
            tFrez = RaschotNromPoRezhimam_RuchnVvodRezhimov(ctx)
        End If
    Else
        tFrez = RaschotNromPoNormativu(ctx)
    End If

    Dim Normy(1 To 3) As Double
    Normy(1) = tFrez
    Normy(2) = ZachistkaZausencev_PoKonturu_Napilnikom(ctx.Material, 2 * (ctx.ShirinaPoverkhnosti + ctx.DlinaPoverkhnosti), False, False)
    Normy(3) = IzmereniyePosleFrezerovaniya(ctx.ShirinaPoverkhnosti, ctx.DlinaPoverkhnosti, ctx.Pripusk, ctx.IT) / 2
    
    FrezKontsevymiFrezami_Kompleks = Normy
    
End Function

Private Function RaschotNormPoRezhimam_AvtoRezhimy(ctx As clsMillingContext) As Double

    Dim RaschotRezhimov As New clsRezhimy_FrezKonts
    
    ctx.Dfrezy = RaschotRezhimov.OpredelitDfrezy_FrezPloskKonts(ctx.ShirinaPoverkhnosti)
    Set ctx.Rezhimy = RaschotRezhimov.RaschotRezhimov(ctx.Material, ctx.Pripusk, ctx.Dfrezy, ctx.ShirinaPoverkhnosti, ctx.IT, ctx.Ra, ctx.Korka)
    RaschotNormPoRezhimam_AvtoRezhimy = RaschotRezhimov.RaschotVremeni_FrezPloskKonts(ctx.Rezhimy, ctx.TipStanka, ctx.Dfrezy, ctx.DlinaPoverkhnosti, ctx.ShirinaPoverkhnosti, ctx.IT, ctx.Ra, ctx.CHPU)
    
End Function

Private Function RaschotNromPoRezhimam_RuchnVvodRezhimov(ctx As clsMillingContext) As Double
    
    Dim RaschotRezhimov As New clsRezhimy_FrezKonts
    RaschotNromPoRezhimam_RuchnVvodRezhimov = RaschotRezhimov.RaschotVremeni_FrezPloskKonts(ctx.Rezhimy, ctx.TipStanka, ctx.Dfrezy, ctx.DlinaPoverkhnosti, ctx.ShirinaPoverkhnosti, ctx.IT, ctx.Ra, ctx.CHPU)

End Function

Private Function RaschotNromPoNormativu(ctx As clsMillingContext) As Double

    Dim Raschot As New clsFrez_PloskKonts_Normativ
    RaschotNromPoNormativu = Raschot.RashotNorm(ctx)

End Function

