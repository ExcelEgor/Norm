VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRaschotTmash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDlinaPoverkhosti As Double
Private pDlinaNedobegaPerebega As Double
Private pSmm_min As Double
Private pKolVoProkhodov As Double

Public Property Let DlinaPoverkhosti(Value As Double)
    pDlinaPoverkhosti = Value
End Property

Public Property Let DlinaNedobegaPerebega(Value As Double)
    If Value > 0 Then
        pDlinaNedobegaPerebega = Value
    Else
        Err.Raise vbObjectError + 1
    End If
    
End Property

Public Property Let Smm_min(Value As Double)
    If Value > 0 Then
        pSmm_min = Value
    Else
        Err.Raise vbObjectError + 1
    End If
    
End Property

Public Property Let KolVoProkhodov(Value As Double)
    pKolVoProkhodov = Value
End Property

Public Function tMashinnoye() As Double

    tMashinnoye = pKolVoProkhodov * (pDlinaPoverkhosti + pDlinaNedobegaPerebega) / pSmm_min
    
End Function





