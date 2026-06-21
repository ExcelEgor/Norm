VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLentOtrez_Img"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents img As MSForms.Image
Attribute img.VB_VarHelpID = -1

Public Property Set ImageControl(vControl As MSForms.Image): Set img = vControl: End Property

Private Sub img_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    frmLentochnoOtreznieStanki.OtkrytVkladku img
End Sub
