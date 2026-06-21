VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKniga2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub NovyyRaschot(ByRef wbRaschot As Workbook, wsShablon As Worksheet, lstRaschot As MSForms.ListBox)
    
    wbRaschot.Close (False)
    wsShablon.Copy
    Set wbRaschot = ActiveWorkbook
    
    Dim tbRaschot As ListObject
    Set tbRaschot = wbRaschot.Worksheets(1).ListObjects(1)
    
    lstRaschot.Clear
    
End Sub
