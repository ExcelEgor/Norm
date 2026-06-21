Attribute VB_Name = "mTokFrez_Zapolneniye"
Option Explicit
Option Private Module

Sub ZapolnitElementyUpravleniya_TokFrez()
    
    Dim MATERIALY
    MATERIALY = Array(ALUMINIYEVYYE_SPLAVY, MEDNYYE_SPLAVY, STAL_UGLERODISTAYA, STAL_LEGIROVANNAYA, STAL_NERZHAVEYUSHCHAYA, TITANOVYYE_SPLAVY)
    
    DobavitMaterialyVListBox MATERIALY, frmTokFrez.cboMaterial

    ZapolnitTokFrez_TokarnyyeStanki
    
    ZapolnitTokFrez_RastochnyyeFrezernyyeStanki
    
    ZapolnitZachistkuShestigrannika
    
    ZapolnitGalvaniku
    
    ZapolnitKomplektovaniye

End Sub

Private Sub ZapolnitZachistkuShestigrannika()

    With frmTokFrez.cboMaterial_ZausShestigran
    
        .AddItem
        .List(.ListCount - 1, 0) = "Медные и алюминиевые сплавы"
        .List(.ListCount - 1, 1) = EnumMaterialy.ALUMINIYEVYYE_SPLAVY
        
        .AddItem
        .List(.ListCount - 1, 0) = "Сталь углеродистая"
        .List(.ListCount - 1, 1) = EnumMaterialy.STAL_UGLERODISTAYA
        
        .AddItem
        .List(.ListCount - 1, 0) = "Сталь нержавеющая"
        .List(.ListCount - 1, 1) = EnumMaterialy.STAL_NERZHAVEYUSHCHAYA

        .ListIndex = 0
    End With
    
End Sub

Private Sub ZapolnitGalvaniku()

    With frmTokFrez
        .cboGalvanika.List = Split("Обезжиривание;Травление;Цинкование;Ан. Окс. хр;Ан. Окс. тв;Ан. Окс. Черн;Хим. Пас;Н12Х;Н18;О-Ви", ";")
        .cboGalvanika.ListIndex = 0
    End With
    
End Sub

Private Sub ZapolnitKomplektovaniye()

    With frmTokFrez.cboKomplekt_Massa
        .List = Split("5 15 30 50 70 100 200")
        .ListIndex = 0
    End With
    
End Sub



