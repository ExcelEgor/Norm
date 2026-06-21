VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPoverkhost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMaterial As EnumMaterialy
Private mDlina As Double
Private mShirina As Double
Private mPripusk As Boolean
Private mRa As Double
Private mIT As Integer

Public Property Let Material(ByVal Value As EnumMaterialy)
    mMaterial = Value
End Property

Public Property Get Material() As EnumMaterialy
    Material = mMaterial
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

Public Property Let Dlina(ByVal Value As Integer)
    mDlina = Value
End Property

Public Property Get Dlina() As Integer
    Dlina = mDlina
End Property

Public Property Let Shirina(ByVal Value As Integer)
    mShirina = Value
End Property

Public Property Get Shirina() As Integer
    Shirina = mShirina
End Property

Public Property Let Pripusk(ByVal Value As Integer)
    mPripusk = Value
End Property

Public Property Get Pripusk() As Integer
    Pripusk = mPripusk
End Property


