VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMillingStage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Frezerovaniye

Option Explicit

Public iB As Long
Public iH As Long
Public Bfrez As Double
Public GlubRez As Double
Public Pokazateli As clsRezhimIndicators

Private Sub Class_Initialize()
    Set Pokazateli = New clsRezhimIndicators
End Sub
