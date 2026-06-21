VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOtverstiye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMaterial As EnumMaterialy
Private mDiametr As Double
Private mGlubina As Double
Private mGlukhoye As Boolean
Private mRa As Double
Private mIT As Integer

Public Property Let Material(ByVal Value As EnumMaterialy)
    mMaterial = Value
End Property

Public Property Get Material() As EnumMaterialy
    Material = mMaterial
End Property

Public Property Let Diametr(ByVal Value As Double)
    mDiametr = Value
End Property

Public Property Get Diametr() As Double
    Diametr = mDiametr
End Property

Public Property Let Glubina(ByVal Value As Double)
    mGlubina = Value
End Property

Public Property Get Glubina() As Double
    Glubina = mGlubina
End Property

Public Property Let Glukhoye(ByVal Value As Boolean)
    mGlukhoye = Value
End Property

Public Property Get Glukhoye() As Boolean
    Glukhoye = mGlukhoye
End Property

Public Property Let Ra(ByVal Value As Double)
    mRa = Value
End Property

Public Property Get Ra() As Double
    Ra = mRa
End Property

Public Property Let IT(ByVal Value As Integer)
    mIT = Value
End Property

Public Property Get IT() As Integer
    IT = mIT
End Property

Public Function tKontrol() As Double
    tKontrol = KontrolOtverstiy(Diametr, Glubina, IT)
End Function

Public Function tSlesar() As Double
    tSlesar = ZachistkaZausencevSOtverstiiVruchnuyu(Material, Diametr, 2, IIf(Glukhoye, 1, 2))
End Function

