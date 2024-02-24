VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComplex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private x_  As Double   ' cartesisch - x
Private y_  As Double   ' cartesisch - y
Private t_  As Double   ' polar - Theta <-- !!! in Radian !!!
Private r_  As Double   ' polar - Rho
'----------------------------------------------------

Public Sub SetCartesian(x As Double, y As Double)
    ' Koordinaten zuweisen
    x_ = x
    y_ = y
    
    ' und Polarkoordinaten errechnen
    Call CalcPolar
End Sub

Public Sub SetPolar(Theta As Double, Rho As Double)
    ' Koordinaten zuweisen
    t_ = Theta
    r_ = Rho
    
    ' und cartesische Koordinaten errechnen
    Call CalcCartesian
End Sub
'----------------------------------------------------

Public Property Get x() As Double
    x = x_
End Property

Public Property Get y() As Double
    y = y_
End Property

Public Property Get Theta() As Double
    Theta = t_
End Property

Public Property Get Rho() As Double
    Rho = r_
End Property
'----------------------------------------------------

Private Sub CalcCartesian()
    ' z = x + iy = r(cost + isint)
    ' =>
    ' Umrechnung polar -> cartesisch:
    ' x = rcost
    ' y = rsint
    x_ = Cos(Theta) * r_
    y_ = Sin(Theta) * r_
End Sub

Private Sub CalcPolar()
    ' z = x + iy = r(cost + isint)
    ' =>
    ' Umrechnung cartesisch -> polar:
    ' r = sqr(x² + y²)
    ' cost = x/r }
    ' sint = y/r } <=>
    ' tant = y/x   =>
    ' t = atn(y/x)
    
    ' Modul errechnen (Rho)
    r_ = Sqr(x_ * x_ + y_ * y_)
    
    ' Winkel errechnen...
    ' Merke: O (null) hat kein Theta!
    
    If x_ <> 0 Then
        t_ = Atn(y_ / x_)
    Else
        ' mit dummer VB-Atn-Funktion richtigen atn emulieren
        If y_ > 0 Then
            t_ = Pi / 2
        ElseIf y_ < 0 Then
            t_ = -Pi / 2
        End If
    End If
    
    ' nochmal nachbessern, damit t_ positiv ist:
    If x_ < 0 Then _
        t_ = t_ + Pi
End Sub
