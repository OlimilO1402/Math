Attribute VB_Name = "modConstructor"
Option Explicit

Public Const Pi As Double = 3.14159265358979

'----------------------------------------------------

Public Points      As Collection    ' alle Punkte
Public lPointsCnt   As Long
Public Transform    As CTransform   ' Transformation
'----------------------------------------------------

' Konstruktoren für die CComplex-Klasse
Public Function CartesianPoint(x As Double, y As Double) As CComplex
    Set CartesianPoint = New CComplex: CartesianPoint.SetCartesian x, y
End Function

Public Function PolarPoint(Theta As Double, Rho As Double) As CComplex
    Set PolarPoint = New CComplex
    PolarPoint.SetPolar Theta, Rho
End Function
