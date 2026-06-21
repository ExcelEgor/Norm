VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMillingRegimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Frezerovaniye

Option Explicit

Public ChernKork As clsMillingStage
Public Chern As clsMillingStage
Public Chist As clsMillingStage

Private Sub Class_Initialize()
    Set ChernKork = New clsMillingStage
    Set Chern = New clsMillingStage
    Set Chist = New clsMillingStage
End Sub
