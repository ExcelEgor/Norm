VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsObnizheniye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mGlubina As Double
Private mDlina As Double
Private mShirina As Double
Private mRadius As Double

Public Property Let Glubina(ByVal Value As Double)
    mGlubina = Value
End Property

Public Property Get Glubina() As Double
    Glubina = mGlubina
End Property

Public Property Let Dlina(ByVal Value As Double)
    mDlina = Value
End Property

Public Property Get Dlina() As Double
    Dlina = mDlina
End Property

Public Property Let Shirina(ByVal Value As Double)
    mShirina = Value
End Property

Public Property Get Shirina() As Double
    Shirina = mShirina
End Property

Public Property Let Radius(ByVal Value As Double)
    mRadius = Value
End Property

Public Property Get Radius() As Double
    Radius = mRadius
End Property
Function Vremya() As Double
    Vremya = OsonovnoyeFrezerovaniye + FrezerovaniyeRadiusov + PritupleniyeOstrykhKromok
End Function

Private Function OsonovnoyeFrezerovaniye()

    Dim GlubinaRezaniya As Double
    GlubinaRezaniya = 2.5
    
    Dim KolVoProkhodovPoGlubine As Integer
    KolVoProkhodovPoGlubine = WorksheetFunction.RoundUp(Glubina / GlubinaRezaniya, 0)
    
    Dim Dfrezy As Integer
    Dfrezy = 40
    
    Dim Kperekrytiya As Double
    Kperekrytiya = 0.4
    
    Dim KolVoProkhodovPoShirine As Integer
    KolVoProkhodovPoShirine = WorksheetFunction.RoundUp(WorksheetFunction.Min(Shirina, Dlina) / (Dfrezy * Kperekrytiya), 0)
    
    Dim KolVoProkhodov As Integer
    KolVoProkhodov = KolVoProkhodovPoGlubine * KolVoProkhodovPoShirine
    
    Dim DlinaRezaniya As Double
    DlinaRezaniya = Glubina + KolVoProkhodov * WorksheetFunction.Max(Shirina, Dlina)
    
    Dim Podacha As Integer
    Podacha = 1300
    
    OsonovnoyeFrezerovaniye = DlinaRezaniya / Podacha
    
End Function

Private Function FrezerovaniyeRadiusov()

    Dim GlubinaRezaniya As Double
    GlubinaRezaniya = 4
    
    Dim KolVoProkhodov As Integer
    KolVoProkhodov = WorksheetFunction.RoundUp(Glubina / GlubinaRezaniya, 0)
    
    Dim KoefDlinaRadiusa    'Коэффициента для расчёта длины резания по радиусу, полученный опытным путём
    KoefDlinaRadiusa = 16
    
    Dim DlinaRezaniya As Double
    DlinaRezaniya = KoefDlinaRadiusa * PI * Radius * KolVoProkhodov + 2 * (Dlina + Shirina)
    
    Dim Podacha As Integer
    Podacha = 800
    
    FrezerovaniyeRadiusov = DlinaRezaniya / Podacha
    
End Function

Private Function PritupleniyeOstrykhKromok()
    
    Dim DlinaRezaniya As Double
    DlinaRezaniya = 2 * (Dlina + Shirina)
    
    Dim Podacha As Integer
    Podacha = 700
    
    PritupleniyeOstrykhKromok = DlinaRezaniya / Podacha
    
End Function

