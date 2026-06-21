Attribute VB_Name = "mVspomVremya_Drugoye"
Option Explicit


Function Tvsp_CHPU(DlinaObrabatyvaemoyPoverkhnosti As Double, DlinaNedobegaPerebega As Double) As Double

    Dim Pausa As Double
    Dim UskorennyyKhod As Integer
    
    UskorennyyKhod = 3000
    Pausa = 3 / 60
    
    Tvsp_CHPU = Pausa + (DlinaObrabatyvaemoyPoverkhnosti + DlinaNedobegaPerebega) / UskorennyyKhod
    
End Function

