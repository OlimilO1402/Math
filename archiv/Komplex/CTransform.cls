VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Translation
Private TransVector_    As CComplex ' Translationsvektor

' Rotation
Private RotCentre_      As CComplex ' Rotationszentrum
Private Angle_          As Double   ' Rotationswinkel <-- !!! in Radian !!!

' Streckung (Homothetie)
Private HomCentre_      As CComplex ' Zentrum der Streckung
Private Factor_         As Double   ' Streckungsfaktor

' gibt an, ob die jeweilige Transformation eingeschaltet ist
Private bTranslate_     As Boolean
Private bRotate_        As Boolean
Private bStretch_       As Boolean
'----------------------------------------------------

Public Property Get Enabled() As Boolean
    ' gibt zurück, ob irgendeine Transformation aktiviert ist
    Enabled = bTranslate_ Or bRotate_ Or bStretch_
End Property
'----------------------------------------------------

' fügt eine Translation hinzu
Public Sub AddTranslation(Translation As CComplex)
    bTranslate_ = True
    
    Set TransVector_ = Translation
End Sub

' fügt eine Rotation hinzu
Public Sub AddRotation(Centre As CComplex, Angle As Double)
    bRotate_ = True
    
    Set RotCentre_ = Centre
    Angle_ = Angle
End Sub

' fügt eine zentrische Streckung hinzu
Public Sub AddHomothetia(Centre As CComplex, Factor As Double)
    bStretch_ = True
    
    Set HomCentre_ = Centre
    Factor_ = Factor
End Sub

Public Sub RemoveTranslation()
    bTranslate_ = False
End Sub

Public Sub RemoveRotation()
    bRotate_ = False
End Sub

Public Sub RemoveHomothetia()
    bStretch_ = False
End Sub
'----------------------------------------------------

Public Function Transform(Point As CComplex) As CComplex
    Dim tmpPoint As CComplex
    
    Set Transform = New CComplex
    
    ' Ursprüngliche Koordinaten setzen
    Transform.SetCartesian Point.x, Point.y
    
    ' WICHTIG:
    ' die Reihenfolge, in der Rotation und Homothetie auf den Punkt
    ' angewendet werden, ist egal. Bei der Translation spielt das leider
    ' eine Rolle, daher wird sie in diesem Beispiel zuletzt angewendet.
    
    ' Rotation:
    ' z' - w = e^(i*t) * (z - w), wobei "w" das Zentrum
    '                             und t der Winkel der Rotation ist
    ' Das ist leider etwas schwer. Um das Problem zu bewältigen, nimmt
    ' man zuerst eine Translation vor, um das Zentrum der Rotation zum Punkt
    ' (0|0) zu machen, dann dreht man und dann schiebt zuletzt die Punkte zurück...
    If bRotate_ Then
        ' Translation zum Ursprung...
        Set tmpPoint = New CComplex
        tmpPoint.SetCartesian Point.x - RotCentre_.x, _
          Point.y - RotCentre_.y
        
        ' ... Rotation um angegebenen Winkel...
        Transform.SetPolar Angle_ + tmpPoint.Theta, tmpPoint.Rho
        
        ' ... Translation zurück.
        Transform.SetCartesian Transform.x + RotCentre_.x, _
          Transform.y + RotCentre_.y
    End If
    
    ' Homothetie:
    ' z' - w = k(z - w) <=>
    ' z' = k(z - w) + w, wobei k der Streckungsfaktor ist
    ' hier isses einfach, da wir uns nicht um die Polarkoordinaten kümmern
    ' müssen, die Streckung geht in einem einzigen Schritt
    If bStretch_ Then _
        Transform.SetCartesian _
          Factor_ * (Transform.x - HomCentre_.x) + HomCentre_.x, _
          Factor_ * (Transform.y - HomCentre_.y) + HomCentre_.y
    
    
    ' Translation:
    ' z' = z + a + ib, wobei "a + ib" den Translationsvektor darstellt
    If bTranslate_ Then _
        Transform.SetCartesian Transform.x + TransVector_.x, Transform.y + TransVector_.y
End Function
