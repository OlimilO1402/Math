Attribute VB_Name = "MMath"
Option Explicit ' OM: 2024-11-30 lines 1478
' OM: 2025-07-22 lines 1826
'#####################  v  for Bit Shifting v   ####################
' by Paul - wpsjr1@syix.com
' http://www.syix.com/wpsjr1/index.html

' Author's comments:  use ShiftLeft04 or ShiftRightZ05 without the wrappers if you need more speed,
' they're 25% faster than these.

' NOTE: YOU *MUST* CALL InitFunctionsShift() BEFORE USING THESE FUNCTIONS

Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const SHLCode As String = "8A4C240833C0F6C1E075068B442404D3E0C20800"  ' shl eax, cl = D3 E0
Private Const SHRCode As String = "8A4C240833C0F6C1E075068B442404D3E8C20800"  ' shr eax, cl = D3 E8
Private Const SARCode As String = "8A4C240833C0F6C1E075068B442404D3F8C20800"  ' sar eax, cl = D3 F8
Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Dim bHoldSHL() As Byte
Dim bHoldSHR() As Byte
Dim bHoldSAR() As Byte
Dim lCompiled As Long
'#####################  ^  for Bit Shifting ^   ####################


Public INDef  As Double 'not defined, undefined like 0 / 0
Public posINF As Double 'positive infinity like 1 / 0
Public negINF As Double 'negative infinity like -1 / 0
Public NaN    As Double 'Not a Number

Public Const Epsilon = 0.0000001
Public Const EpsilonDec As Variant = 1E-16               ' 1E-16
Public Const EpsilonDbl As Double = 0.000000000001 ' 1E-12
Public Const EpsilonSng As Single = 0.00000001     ' 1E-08

'Complex number in cartesian coordinates
Public Type Complex
    Re As Double 'real part of the complex number
    Im As Double 'imaginary part
End Type

'Complex number in polar coordinates or euler form
Public Type ComplexP
    r   As Double 'radius r
    phi As Double 'angle phi
End Type

'this types just for conversions
Private Type TLong
    Value As Long
End Type
Private Type TSingle
    Value As Single
End Type
Private Type TLong2
    Value0 As Long
    Value1 As Long
End Type
Private Type TDouble
    Value As Double
End Type

'value range Byte (unsigned int8)
'0 .. 255
Public Const MinByte    As Byte = 0
Public Const MaxByte    As Byte = 255

'value range Integer (short) (signed int16)
'-32768 .. 32767
Public Const MinInteger As Integer = &H8000 '-32768
Public Const MaxInteger As Integer = &H7FFF ' 32767

'value range Long (Int32)
'-2147483648 .. 2147483647
Public Const MinLong    As Long = &H80000000  '-2147483648
Public Const MaxLong    As Long = &H7FFFFFFF  ' 2147483647

Public Const MinSingle  As Single = -3.40282347E+38 '-3.40282347E+38;
Public Const MaxSingle  As Single = 3.40282347E+38  ' 3.40282347E+38

Public Const MinDouble As Double = -1.7976931348623E+308 ' -1.7976931348623157E+308;
Public Const MaxDouble As Double = 1.7976931348623E+308  ' = 1.7976931348623157E+308;

'value range Currency (Int64 / 10000)
Public Const MinCurrency As Currency = -922337203685477.5807@
Public Const MaxCurrency As Currency = 922337203685477.5807@

'value range Currency (Int64)
'    Dim Cur1: bigDec1 = CDec("9223372036854775807")   ' No overflow.
'    Dim Cur2: bigDec2 = CDec("9223372036854775808")   ' No overflow.
'    Dim Cur3: bigDec3 = CDec("9223372036854775809")   ' No overflow.

'value range Decimal
'    Dim bigDec1: bigDec1 = CDec("79228162514264337593543950335")
'    Dim bigDec2: bigDec2 = CDec("-79228162514264337593543950335")
'    Dim bigDec3: bigDec3 = CDec("7,9228162514264337593543950335")
'    Dim bigDec4: bigDec4 = CDec("-7,9228162514264337593543950335")
'    Dim bigDec5: bigDec5 = CDec("0,0000000000000000000000000001")
'    Dim bigDec6: bigDec6 = CDec("-0,0000000000000000000000000001")


'Mathematical constants
'https://de.wikipedia.org/wiki/Mathematische_Konstante
'https://de.wikipedia.org/wiki/Liste_besonderer_Zahlen
Public Pi          ' As Decimal
Public Pihalf      ' As Decimal
Public Pi2         ' As Decimal
Public Euler       ' As Decimal
Public SquareRoot2 ' As Decimal
Public SquareRoot3 ' As Decimal
Public GoldenRatio ' As Decimal

'Physical constants
Public SpeedOfLight    ' Lichtgeschwindigkeit im Vakuum      c   = 299792458 m/s
Public ElemCharge      ' Elementarladung (des Protons)       e   = 1,602176634 * 10^-19 C (Coulomb)
Public MassElektron    ' Ruhemasse des Elektrons             m_e = 9,109*10^-31 kg
Public MassProton      ' Ruhemasse des Protons               m_p = 1,6726215813 · 10^-27 kg
Public PlanckQuantum   ' Plancksches Wirkungsquantum         h   = 6,62607015 * 10^(-37) m² * kg / s
Public Avogadro        ' Avogadro-Konstante                  N_A = 6,022 * 10^23
Public Gravitation     ' Newtonsche Gravitationskonstante    G   = 6,6743 * 10^-11 m³ / (kg * s²)
Public Boltzmann       ' Boltzmann-Konstante                 k_B = 1,38064852 × 10^-23 m² kg / (s² * K)
Public StefanBoltzmann ' Stefan Boltzmann-Konstante          sigma = 5,670374419 × 10^-8 W/(m² * K^4)
Public MagnPermittvy   ' magnetische Feldkonstante           mue_0 = µ0 ˜ 1.2566370621219 * 10 ^(-6) N/A²
Public ElecPermittvy   ' elektrische Feldkonstante           eps_0 = 8.8541878128(13)e-12 (A s)/(V m)
Public QuantumAlpha    ' FineStructureConstant

Public Const TempCelsius_AbsolutZero As Double = 273.15

Private m_Factorials()   ' As Decimal
Public Primes()  As Long ' contains all primes up to 100000
Public PrimesX() As Long ' a distinct selection of primes

Public Fibonacci() As Long



'NTSYSAPI SIZE_T RtlCompareMemory(
'  [in] const VOID *Source1,
'  [in] const VOID *Source2,
'  [in] SIZE_T     Length
');
'Private Declare Function RtlCompareMemory Lib "ntdll" (pSrc1 As Long, pSrc2 As Long, ByVal Length As Long) As Long

'#####################  v  for Bit Shifting v   ####################
Private Sub InitFunctionsShift() ' call this in your Sub Main or Form_Load
    If Compiled Then
        SubstituteCode bHoldSHL, SHLCode, AddressOf ShiftLeft
        SubstituteCode bHoldSHR, SHRCode, AddressOf ShiftRightZ
        SubstituteCode bHoldSAR, SARCode, AddressOf ShiftRight
    End If
End Sub

Public Sub Init()
        'Pi = CDec("3,1415926535897932384626433832795") '0288419716939937510582097494459230781640628620899862803482534211706798214")
                  ' 3,1415926535897932384626433833
     'Euler = CDec("2,7182818284590452353602874713526") '6249775724709369995957496696762772407663035354759457138217852516642742746")
                  ' 2,7182818284590452353602874714


'https://oeis.org/A000796
         Pi = Constant_Parse(1, "3, 1, 4, 1, 5, 9, 2, 6, 5, 3, 5, 8, 9, 7, 9, 3, 2, 3, 8, 4, 6, 2, 6, 4, 3, 3, 8, 3, 2, 7, 9, 5, 0, 2, 8, 8, 4, 1, 9, 7, 1, 6, 9, 3, 9, 9, 3, 7, 5, 1, 0, 5, 8, 2, 0, 9, 7, 4, 9, 4, 4, 5, 9, 2, 3, 0, 7, 8, 1, 6, 4, 0, 6, 2, 8, 6, 2, 0, 8, 9, 9, 8, 6, 2, 8, 0, 3, 4, 8, 2, 5, 3, 4, 2, 1, 1, 7, 0, 6, 7, 9, 8, 2, 1, 4")
     
         Pi = CDec(CDec(4) * CDec(Atn(1)))
     
     Pihalf = Pi / CDec(2) 'Constant_Parse(1, "")
        Pi2 = Constant_Parse(1, "6, 2, 8, 3, 1, 8, 5, 3, 0, 7, 1, 7, 9, 5, 8, 6, 4, 7, 6, 9, 2, 5, 2, 8, 6, 7, 6, 6, 5, 5, 9, 0, 0, 5, 7, 6, 8, 3, 9, 4, 3, 3, 8, 7, 9, 8, 7, 5, 0, 2, 1, 1, 6, 4, 1, 9, 4, 9, 8, 8, 9, 1, 8, 4, 6, 1, 5, 6, 3, 2, 8, 1, 2, 5, 7, 2, 4, 1, 7, 9, 9, 7, 2, 5, 6, 0, 6, 9, 6, 5, 0, 6, 8, 4, 2, 3, 4, 1, 3")
        
'https://oeis.org/A001113
      Euler = Constant_Parse(1, "2, 7, 1, 8, 2, 8, 1, 8, 2, 8, 4, 5, 9, 0, 4, 5, 2, 3, 5, 3, 6, 0, 2, 8, 7, 4, 7, 1, 3, 5, 2, 6, 6, 2, 4, 9, 7, 7, 5, 7, 2, 4, 7, 0, 9, 3, 6, 9, 9, 9, 5, 9, 5, 7, 4, 9, 6, 6, 9, 6, 7, 6, 2, 7, 7, 2, 4, 0, 7, 6, 6, 3, 0, 3, 5, 3, 5, 4, 7, 5, 9, 4, 5, 7, 1, 3, 8, 2, 1, 7, 8, 5, 2, 5, 1, 6, 6, 4, 2, 7, 4, 2, 7, 4, 6")
    
'https://oeis.org/A002193
SquareRoot2 = Constant_Parse(1, "1, 4, 1, 4, 2, 1, 3, 5, 6, 2, 3, 7, 3, 0, 9, 5, 0, 4, 8, 8, 0, 1, 6, 8, 8, 7, 2, 4, 2, 0, 9, 6, 9, 8, 0, 7, 8, 5, 6, 9, 6, 7, 1, 8, 7, 5, 3, 7, 6, 9, 4, 8, 0, 7, 3, 1, 7, 6, 6, 7, 9, 7, 3, 7, 9, 9, 0, 7, 3, 2, 4, 7, 8, 4, 6, 2, 1, 0, 7, 0, 3, 8, 8, 5, 0, 3, 8, 7, 5, 3, 4, 3, 2, 7, 6, 4, 1, 5, 7")

'https://oeis.org/A002194
SquareRoot3 = Constant_Parse(1, "1, 7, 3, 2, 0, 5, 0, 8, 0, 7, 5, 6, 8, 8, 7, 7, 2, 9, 3, 5, 2, 7, 4, 4, 6, 3, 4, 1, 5, 0, 5, 8, 7, 2, 3, 6, 6, 9, 4, 2, 8, 0, 5, 2, 5, 3, 8, 1, 0, 3, 8, 0, 6, 2, 8, 0, 5, 5, 8, 0, 6, 9, 7, 9, 4, 5, 1, 9, 3, 3, 0, 1, 6, 9, 0, 8, 8, 0, 0, 0, 3, 7, 0, 8, 1, 1, 4, 6, 1, 8, 6, 7, 5, 7, 2, 4, 8, 5, 7, 5, 6, 7, 5, 6, 2, 6, 1, 4, 1, 4, 1, 5, 4")

'https://oeis.org/A001622
GoldenRatio = Constant_Parse(1, "1, 6, 1, 8, 0, 3, 3, 9, 8, 8, 7, 4, 9, 8, 9, 4, 8, 4, 8, 2, 0, 4, 5, 8, 6, 8, 3, 4, 3, 6, 5, 6, 3, 8, 1, 1, 7, 7, 2, 0, 3, 0, 9, 1, 7, 9, 8, 0, 5, 7, 6, 2, 8, 6, 2, 1, 3, 5, 4, 4, 8, 6, 2, 2, 7, 0, 5, 2, 6, 0, 4, 6, 2, 8, 1, 8, 9, 0, 2, 4, 4, 9, 7, 0, 7, 2, 0, 7, 2, 0, 4, 1, 8, 9, 3, 9, 1, 1, 3, 7, 4, 8, 4, 7, 5")



'https://oeis.org/A003678
SpeedOfLight = Constant_Parse(9, "2, 9, 9, 7, 9, 2, 4, 5, 8") ' m / sec

ElemCharge = Constant_Parse(1, "1, 6, 0, 2, 1, 7, 6, 6, 3, 4") * 10 ^ -19 'C (Coulomb)

'https://oeis.org/A081801
MassElektron = Constant_Parse(1, "9, 1, 0, 9, 3, 8") * 10 ^ -31 'kg

'1,67262192 * 10 ^ (-27) kg
'https://oeis.org/A070059
MassProton = Constant_Parse(1, "1, 6, 7, 2, 6, 2, 1, 9, 2") * 10 ^ -27 'kg

'https://oeis.org/A003676
PlanckQuantum = 6.62607015 * 10 ^ (-34)

'Teilchenzahl
'N_A
'https://oeis.org/A322578
Avogadro = Constant_Parse(1, "6, 0, 2, 2, 1, 4, 0, 7, 6") * 10 ^ 23 '1/mol

'https://oeis.org/A070058
'6.674 30(15) * 10^(-11) m^3 kg^(-1) s^(-2)
Gravitation = Constant_Parse(6, "6, 6, 7, 4, 3, 9, 0") * 10 ^ -11

'Massen von Atomen oder Molekülen werden in der Einheit u angegeben
'Ein u ist als ein Zwölftel der Masse eines Atoms des Kohlenstoffisotops C12 definiert
'1 u = 1/12 * m(12^_C)
'man kann u in Gramm umrechnen
'1g = 1 / (1,66*10^-24) u = 6,022 * 10^23 u
'Wenn man u in g umrechnen will muss man den Wert durch 6,022 mal 10^23 teilen
'ein Blick ins Periodensystem verrät uns wie schwer 1 Aluminiumatom ist, ca 27u
'Prinzipiell müssen wir nur 1g durch 27u teilen, nur können wir das nur machen,
'wenn wir dieselben Einheiten haben
'
'https://oeis.org/A070063
' Boltzmann-Konstante                 k_B = 1,38064852 × 10-23 m2 kg s-2 K-1
Boltzmann = Constant_Parse(1, "1, 3, 8, 0, 6, 4, 9") * 10 ^ -23

'https://oeis.org/A081820
' Stefan Boltzmann-Konstante    sigma = 5,670374419 × 10-8 W/(m^2 * K^4)
StefanBoltzmann = Constant_Parse(1, "5, 6, 7, 0, 3, 7, 4, 4, 1, 9, 1, 8, 4, 4, 2, 9, 4, 5, 3, 9, 7, 0, 9, 9, 6, 7, 3, 1, 8, 8, 9, 2, 3, 0, 8, 7, 5, 8, 4, 0, 1, 2, 2, 9, 7, 0, 2, 9, 1, 3, 0, 3, 6, 8, 2, 4, 0, 5, 4, 6, 1, 7, 3, 7, 0, 5, 3, 9, 4, 8, 1, 6, 0, 6, 2, 6, 5, 2, 3, 3, 2, 6, 0, 8, 2, 5, 7, 1, 8, 5, 7, 7, 7, 0, 4, 4, 6, 8, 8, 7, 0, 3") * 10 ^ -8

MagnPermittvy = Constant_Parse(1, "1, 2, 5, 6, 6, 3, 7, 0, 6, 1, 4, 3, 5, 9, 1, 7, 2, 9, 5, 3, 8, 5, 0, 5, 7, 3, 5, 3, 3, 1, 1, 8, 0, 1, 1, 5, 3, 6, 7, 8, 8, 6, 7, 7, 5, 9, 7, 5, 0, 0, 4, 2, 3, 2, 8, 3, 8, 9, 9, 7, 7, 8, 3, 6, 9, 2, 3, 1, 2, 6, 5, 6, 2, 5, 1, 4, 4, 8, 3, 5, 9, 9, 4, 5, 1, 2, 1, 3, 9, 3, 0, 1, 3, 6, 8, 4, 6, 8, 2") * 10 ^ -6   '

ElecPermittvy = Constant_Parse(1, "8, 8, 5, 4, 1, 8, 7, 8, 1, 7, 6, 2, 0, 3, 8, 9, 8, 5, 0, 5, 3, 6, 5, 6, 3, 0, 3, 1, 7, 1, 0, 7, 5, 0, 2, 6, 0, 6, 0, 8, 3, 7, 0, 1, 6, 6, 5, 9, 9, 4, 4, 9, 8, 0, 8, 1, 0, 2, 4, 1, 7, 1, 5, 2, 4, 0, 5, 3, 9, 5, 0, 9, 5, 4, 5, 9, 9, 8, 2, 1, 1, 4, 2, 8, 5, 2, 8, 9, 1, 6, 0, 7, 1, 8, 2, 0, 0, 8, 9, 3, 2, 8, 6, 7") * 10 ^ -12   '

QuantumAlpha = CDec(CDec(1) / CDec(137))

    InitFactorials
    
    'ReDim m_Primes(0 To 9591)
    'ReadPrimesTxt
    'SavePrimesBin
    'https://de.wikibooks.org/wiki/Primzahlen:_Tabelle_der_Primzahlen_(2_-_100.000)
#If ReadPrimesFromFile Then
    ReadPrimesBin
#End If
    
    InitPrimeX
    'InitDedekind
    InitFibonacci
    InitINF
    InitFunctionsShift
End Sub

Public Sub InitINF()
    GetINDef INDef
    posINF = GetINF
    negINF = GetINF(-1)
    GetNaN NaN
End Sub

Private Sub InitFibonacci()
    Fibonacci = FibonacciA
End Sub

Private Sub InitFactorials()
    ReDim m_Factorials(0 To 171)
    Dim i As Long, F: F = CDec(1)
    m_Factorials(0) = F
    m_Factorials(1) = F
    For i = 2 To 27
        F = F * CDec(i)
        m_Factorials(i) = F
    Next
    F = CDbl(F)
    For i = 28 To 170
        F = F * CDbl(i)
        m_Factorials(i) = F
    Next
    m_Factorials(171) = GetINFE
End Sub

Public Function Fact(ByVal n As Long) As Variant 'As Decimal
    If n > 170 Then n = 171
    Fact = m_Factorials(n)
End Function

Public Function ATan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then        'egal ob y > 0 oder y < 0    '1. Quadrant und 4. Quadrant
        ATan2 = VBA.Math.Atn(y / x)
    ElseIf x < 0 Then
        If y > 0 Then                '2. Quadrant
            ATan2 = VBA.Math.Atn(y / x) + Pi
        ElseIf y < 0 Then            '3. Quadrant
            ATan2 = VBA.Math.Atn(y / x) - Pi
        Else                         'neg x-Achse
            ATan2 = Pi
        End If
    Else
        If y > 0 Then                'pos y-Achse
            ATan2 = 0.5 * Pi
        ElseIf y < 0 Then            'neg y-Achse
            ATan2 = -0.5 * Pi
        Else                         'Nullpunkt
            ATan2 = 0#
        End If
    End If
End Function

Public Function Square(ByVal n As Double) As Double
   Square = n * n
End Function

Private Function Constant_Parse(ByVal nDigsVkst As Byte, ByVal sc As String) As String
    sc = Replace(sc, ", ", "")
    Dim s As String: s = Left(sc, nDigsVkst)
    If Len(sc) > nDigsVkst Then s = s & "," & Mid(sc, nDigsVkst + 1)
    Constant_Parse = s
End Function

' v ############################## v '    ggT and kgV-functions    ' v ############################## v '

'function ggT(a,b:integer):integer;
'   var c:integer;
'Begin
'    repeat
'      c:=a mod b;
'      a:=b;
'      b:=c;
'    until c=0;
'    result:=a
'end;
Function ggT(ByVal x As Long, ByVal y As Long) As Long
    'ggT = größter gemeinsamer Teiler
   Do While x <> y
      If x > y Then
         x = x - y
      Else
         y = y - x
      End If
   Loop 'Wend
   ggT = x
End Function

Public Function GreatestCommonDivisor(ByVal x As Long, ByVal y As Long) As Long
   GreatestCommonDivisor = ggT(x, y)
End Function

'function kgV(a,b:integer):integer;
'Begin
'  result:=a*b div ggT(a,b);
'end;
Public Function kgV(ByVal x As Long, ByVal y As Long) As Long
    kgV = (x * y) \ ggT(x, y)
End Function

Public Function LeastCommonMultiple(ByVal x As Long, ByVal y As Long) As Long
    'kgV = kleinstes gemeinsames Vielfaches
    LeastCommonMultiple = kgV(x, y)
End Function

Public Function PFZ(ByVal n As Long) As String
    Dim s As String
    Dim i As Long: i = 2 'CDec(2)
    Do
        While n Mod i = 0
            n = n / i
            If s <> vbNullString Then s = s & "*" 'first time wo *
            s = s & CStr(i)
        Wend
        'If i = 2 Then i = i + CDec(1) Else i = i + 2 'CDec(2)
        If i = 2 Then i = i + 1 Else i = i + 2 'CDec(2)
        If i > Int(Sqr(n)) Then i = n '//ohne diese Zeile:Kaffeepause!
    Loop Until n = 1
    If InStr(s, "*") = 0 Then s = "Primzahl"
    PFZ = s
End Function

'Bruch = Zähler / Nenner   'fraction = numerator / denominator
Public Function CancelFraction(numerator_inout As Long, denominator_inout As Long) As Boolean
    Dim t As Long: t = ggT(numerator_inout, ByVal denominator_inout)
    numerator_inout = numerator_inout \ t
    denominator_inout = denominator_inout \ t
    If denominator_inout < 0 Then
        numerator_inout = -numerator_inout
        denominator_inout = -denominator_inout
    End If
    CancelFraction = True
End Function
    
'procedure kuerze(var z,n:integer);
'  var t:integer;
'Begin
'  t:=ggT(z,n);
'  z:=z div t;
'  n:=n div t;
'  If n < 0 Then Begin
'    n:=-n; //positiv
'    z:=-z;
'  End;
'end;

Public Function IsPowerOfTwo(ByVal x As Long) As Boolean
    IsPowerOfTwo = (x <> 0) And ((x And (x - 1)) = 0)
End Function


'Public Shared Function Pow(ByVal x As Double, ByVal y As Double) As Double
Public Static Function Pow(ByVal x As Double, ByVal y As Double) As Double 'cDouble
    'Set Pow = New cDouble
    Pow = x ^ y
End Function

Public Static Function Powr2(ByVal Exponent As Long) As Long
    Powr2 = Pow2(Exponent)
End Function
 
Public Static Function Pow2(ByVal Exponent As Long) As Long
    ' by Donald, donald@xbeat.net, 20001217
    ' * Power205
    Dim alPow2(0 To 31) As Long
    Dim i As Long
    
    Select Case Exponent
    Case 0 To 31
        ' initialize lookup table
        If alPow2(0) = 0 Then
            alPow2(0) = 1
            For i = 1 To 30
                alPow2(i) = alPow2(i - 1) * 2
            Next
            alPow2(31) = &H80000000
        End If
        ' return
        Pow2 = alPow2(Exponent)
    End Select
End Function

Private Function Compiled() As Long
    On Error Resume Next
    Debug.Print 1 \ 0
    Compiled = (Err.number = 0)
End Function

' ^ ############################## ^ '    ggT and kgV-functions    ' ^ ############################## ^ '

' v ############################## v '    Linear interpolation    ' v ############################## v '
'x1 ' y1
'x2 ' LinIPol
'x3 ' y3
Public Function LinIPol(ByVal y1 As Double, _
                        ByVal y3 As Double, _
                        ByVal x1 As Double, _
                        ByVal x2 As Double, _
                        ByVal x3 As Double) As Double

    ' errechnet einen Wert y2 zu dem Wert x2 durch lineare Interpolation
    If (x3 - x1) = 0 Then
        LinIPol = y1
    Else
        LinIPol = y1 + (y3 - y1) / (x3 - x1) * (x2 - x1)
    End If

End Function

'     |  x1  |  x   |  x2
'-------------------------
' y1  | f11     ?     f21
' y   |  ?     ???     ?
' y2  | f12     ?     f22
'
Public Function BilIPol(ByVal x1 As Double, ByVal x As Double, ByVal x2 As Double, ByVal y1 As Double, ByVal y As Double, ByVal y2 As Double, ByVal f11 As Double, ByVal f12 As Double, ByVal f21 As Double, ByVal f22 As Double) As Double
    'https://de.wikipedia.org/wiki/Bilineare_Filterung
    Dim x2Minx1 As Double: x2Minx1 = x2 - x1
    Dim x2MinxDivx2Minx1 As Double: x2MinxDivx2Minx1 = (x2 - x) / x2Minx1
    Dim xMinx1Divx2Minx1 As Double: xMinx1Divx2Minx1 = (x - x1) / x2Minx1
    Dim R1 As Double: R1 = x2MinxDivx2Minx1 * f11 + xMinx1Divx2Minx1 * f21
    Dim R2 As Double: R2 = x2MinxDivx2Minx1 * f12 + xMinx1Divx2Minx1 * f22
    BilIPol = (y2 - y) / (y2 - y1) * R1 + (y - y1) / (y2 - y1) * R2
End Function


' ^ ############################## ^ '    Linear interpolation    ' ^ ############################## ^ '

' v ############################## v '    prime-functions    ' v ############################## v '

Public Function GetPrime(ByVal Min As Long) As Long
    If (Min < 0) Then
        'Throw New ArgumentException(Environment.GetResourceString("Arg_HTCapacityOverflow"))
        MsgBox "Arg_HTCapacityOverflow"
        Exit Function
    End If
    Dim i As Long
    Dim num2 As Long
    For i = 0 To UBound(Primes)
        num2 = Primes(i)
        If (num2 >= Min) Then
            GetPrime = num2
            'Debug.Print "min: " & CStr(min) & " GetPrime 1 " & CStr(GetPrime)
            Exit Function
        End If
    Next i
    Dim j As Long: j = (Min Or 1)
    Do While (j < &H7FFFFFFF)
        If MMath.IsPrime(j) Then
            GetPrime = j
            'Debug.Print "min: " & CStr(min) & " GetPrime 2 " & CStr(GetPrime)
            Exit Function
        End If
        j = (j + 2)
    Loop
    'Return min
    GetPrime = Min
End Function

'Public Function IsPrimeA(ByVal Value As Long) As Boolean
'    If Value < 99992 Then
'        Dim i As Long
'        For i = 0 To UBound(m_Primes)
'            If Value = m_Primes(i) Then IsPrimeA = True: Exit Function
'        Next
'    Else
'        '
'    End If
'End Function

'Function IsPrimeN(ByVal Value As Long) As Boolean
'    If Value = 2 Or Value = 3 Or Value = 5 Or Value = 7 Then
'        IsPrimeN = True
'        Exit Function
'    End If
'    If Value = 1 Or Value Mod 2 = 0 Or Value Mod 3 = 0 Or Value Mod 7 = 0 Or Value Mod 5 = 0 Then
'        'IsPrimeN = False
'        Exit Function
'    End If
'    Dim max As Long: max = CInt(Math.Sqr(Value))
'    Dim z   As Long: z = 5
'    'Do While z <= max
'    While z <= max
'        If Value Mod z = 0 Or Value Mod (z + 2) = 0 Then
'            'IsPrimeN = False
'            Exit Function
'        End If
'        z = z + 6
'    Wend 'While
'    'Loop
'    IsPrimeN = True
'End Function

Function IsPrime(ByVal Value As Long) As Boolean
'    If Value < 200 Then
'        Select Case Value
'        Case 2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47, 53, 59, 61, 67, 71, 73, 79, 83, 89, 97, 101, _
'              103, 107, 109, 113, 127, 131, 137, 139, 149, 151, 157, 163, 167, 173, 179, 181, 191, 193, 197, 199
'            IsPrime = True:        Exit Function
'        End Select
'    End If
    If (Value And 1) = 0 Then Exit Function
    Dim div As Long: div = 3
    Dim squ As Long: squ = 9
    Do While squ < Value
        If Value Mod div = 0 Then Exit Function
        div = div + 2
        squ = div * div
    Loop
    If squ <> Value Then
        IsPrime = True
    End If
    InitPrimeX
End Function

Public Function IsPrimeX(ByVal number As Long) As Boolean
    'handelt es sich bei einer Zahl um eine Primzahl?
    If ((number And 1) = 0) Then
        IsPrimeX = (number = 2)
        Exit Function
    End If
    Dim n As Long: n = CLng(VBA.Math.Sqr(CDbl(number)))
    Dim i As Long: i = 3
    Do While (i <= n)
        If ((number Mod i) = 0) Then
            IsPrimeX = False
            Exit Function
        End If
        i = i + 2
    Loop
    IsPrimeX = True
End Function

Public Sub InitPrimeX()
   Call FillArray(PrimesX, 3, 7, 11, &H11, &H17, &H1D, &H25, &H2F, &H3B, &H47, &H59, &H6B, &H83, &HA3, &HC5, &HEF, _
   &H125, &H161, &H1AF, &H209, &H277, &H2F9, &H397, &H44F, &H52F, &H63D, &H78B, &H91D, &HAF1, &HD2B, &HFD1, _
   &H12FD, &H16CF, &H1B65, &H20E3, &H2777, &H2F6F, &H38FF, &H446F, &H521F, &H628D, &H7655, &H8E01, &HAA6B, _
   &HCC89, &HF583, &H126A7, &H1619B, &H1A857, &H1FD3B, &H26315, &H2DD67, &H3701B, &H42023, &H4F361, &H5F0ED, _
   &H72125, &H88E31, &HA443B, &HC51EB, &HEC8C1, &H11BDBF, &H154A3F, &H198C4F, &H1EA867, &H24CA19, &H2C25C1, _
   &H34FA1B, &H3F928F, &H4C4987, &H5B8B6F, &H6DDA89)
End Sub

Private Sub FillArray(ByRef arr() As Long, ParamArray params())
    ReDim arr(0 To UBound(params))
    Dim i As Long
    For i = 0 To UBound(params)
        arr(i) = CLng(params(i))
    Next
End Sub

Public Function NextPrime(ByVal number As Long) As Long
    'liefert zu einer Zahl die nächsthöhere positive Primzahl
    'oder die Zahl sebst falls sie schon eine Primzahl ist.
    'Nur positive Primzahlen
    number = CLng(Abs(number))
    Dim i As Long, p As Long
    For i = 0 To UBound(Primes)
        p = Primes(i)
        If (p >= number) Then
            NextPrime = p
            Exit Function
        End If
    Next
    p = (number Or 1)
    Do While (p < &H7FFFFFFF)
        If IsPrime(p) Then
            NextPrime = p
            Exit Function
        End If
        p = p + 2
    Loop
    NextPrime = number
End Function

Public Function GetRandomPrime() As Long
    Dim i As Long: i = Rnd * 100000
    GetRandomPrime = Prime(i)
End Function

Public Property Get Prime(ByVal Index As Long) As Long
    Prime = Primes(Index)
End Property

Sub ReadPrimesBin()
    Dim FNm As String:  FNm = App.Path & "\" & "Primes100000.Int32"
    Dim FNr As Integer: FNr = FreeFile
    Open FNm For Binary Access Read As FNr
    Dim n As Long: n = LOF(FNr) / 4
    'Debug.Print n
    'ReDim Primes(0 To 9591)
    ReDim Primes(0 To n - 1)
    Get FNr, , Primes
    Close FNr
End Sub

'Sub SavePrimesBin()
'    Dim FNm As String:  FNm = App.Path & "\" & "Primes100000.Int32"
'    Dim FNr As Integer: FNr = FreeFile
'    Open FNm For Binary Access Write As FNr
'    'Dim FContent As String: FContent = Space(LOF(FNr))
'    'Get FNr, , FContent
'    Put FNr, , m_Primes
'    Close FNr
'End Sub
'
'Sub ReadPrimesTxt()
'    Dim FNm As String:  FNm = App.Path & "\" & "Primes100000.txt"
'    Dim FNr As Integer: FNr = FreeFile
'    Open FNm For Binary Access Read As FNr
'    Dim FContent As String: FContent = Space(LOF(FNr))
'    Get FNr, , FContent
'    Close FNr
'    Dim lines() As String: lines = Split(FContent, vbCrLf)
'    Dim i As Long, j As Long, c As Long, line As String, sa() As String, s As String
'    For i = 0 To UBound(lines)
'        line = lines(i)
'        sa = Split(line, ",")
'        For j = 0 To UBound(sa)
'            s = Trim(sa(j))
'            If Len(s) Then
'                If IsNumeric(s) Then
'                    m_Primes(c) = CLng(s)
'                    c = c + 1
'                End If
'            End If
'        Next
'    Next
'End Sub

' ^ ############################## ^ '    prime-functions    ' ^ ############################## ^ '


' v ############################## v '    Min-Max-functions    ' v ############################## v '

Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
Public Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function

Public Function MinArr(ParamArray params())
    Dim i As Long
    MinArr = params(0)
    For i = 1 To UBound(params) - 1
        MinArr = Min(MinArr, params(i))
    Next
End Function
Public Function MaxArr(ParamArray params())
    Dim i As Long
    MaxArr = params(0)
    For i = 1 To UBound(params) - 1
        MaxArr = Max(MaxArr, params(i))
    Next
End Function

Public Function MinByt(ByVal V1 As Byte, ByVal V2 As Byte) As Byte
    If V1 < V2 Then MinByt = V1 Else MinByt = V2
End Function
Public Function MaxByt(ByVal V1 As Byte, ByVal V2 As Byte) As Byte
    If V1 > V2 Then MaxByt = V1 Else MaxByt = V2
End Function
Public Function MinByt3(ByVal V1 As Byte, ByVal V2 As Byte, ByVal V3 As Byte) As Byte
    If V1 < V2 Then
        If V1 < V3 Then MinByt3 = V1 Else MinByt3 = V3
    Else
        If V2 < V3 Then MinByt3 = V2 Else MinByt3 = V3
    End If
End Function
Public Function MaxByt3(ByVal V1 As Byte, ByVal V2 As Byte, ByVal V3 As Byte) As Byte
    If V1 > V2 Then
        If V1 > V3 Then MaxByt3 = V1 Else MaxByt3 = V3
    Else
        If V2 > V3 Then MaxByt3 = V2 Else MaxByt3 = V3
    End If
End Function

Public Function MinInt(ByVal V1 As Integer, ByVal V2 As Integer) As Integer
    If V1 < V2 Then MinInt = V1 Else MinInt = V2
End Function
Public Function MaxInt(ByVal V1 As Integer, ByVal V2 As Integer) As Integer
    If V1 > V2 Then MaxInt = V1 Else MaxInt = V2
End Function

Public Function MinLng(ByVal V1 As Long, ByVal V2 As Long) As Long
    If V1 < V2 Then MinLng = V1 Else MinLng = V2
End Function
Public Function MaxLng(ByVal V1 As Long, ByVal V2 As Long) As Long
    If V1 > V2 Then MaxLng = V1 Else MaxLng = V2
End Function

Public Function MinSng(ByVal V1 As Single, ByVal V2 As Single) As Single
    If V1 < V2 Then MinSng = V1 Else MinSng = V2
End Function
Public Function MaxSng(ByVal V1 As Single, ByVal V2 As Single) As Single
    If V1 > V2 Then MaxSng = V1 Else MaxSng = V2
End Function
Public Function MinSng3(ByVal V1 As Single, ByVal V2 As Single, ByVal V3 As Single) As Single
    If V1 < V2 Then
        If V1 < V3 Then MinSng3 = V1 Else MinSng3 = V3
    Else
        If V2 < V3 Then MinSng3 = V2 Else MinSng3 = V3
    End If
End Function
Public Function MaxSng3(ByVal V1 As Single, ByVal V2 As Single, ByVal V3 As Single) As Single
    If V1 > V2 Then
        If V1 > V3 Then MaxSng3 = V1 Else MaxSng3 = V3
    Else
        If V2 > V3 Then MaxSng3 = V2 Else MaxSng3 = V3
    End If
End Function

Public Function MinDbl(ByVal V1 As Double, ByVal V2 As Double) As Double
    If V1 < V2 Then MinDbl = V1 Else MinDbl = V2
End Function
Public Function MaxDbl(ByVal V1 As Double, ByVal V2 As Double) As Double
    If V1 > V2 Then MaxDbl = V1 Else MaxDbl = V2
End Function

Public Function MinCur(ByVal V1 As Currency, ByVal V2 As Currency) As Currency
    If V1 < V2 Then MinCur = V1 Else MinCur = V2
End Function
Public Function MaxCur(ByVal V1 As Currency, ByVal V2 As Currency) As Currency
    If V1 > V2 Then MaxCur = V1 Else MaxCur = V2
End Function

'Einen Wert in Schranken zwingen, obere Werte auf Max reduzieren, untere Werte auf Min heben
'MinMax, Clamp,
'its like a sound compressor: lower sounds gets louder, louder sounds must not oversteer over a maxvalue
Public Function Clamp(Value, MinVal, MaxVal)
    Clamp = Max(MinVal, Min(MaxVal, Value))
End Function

Public Function ClampByt(ByVal Value As Byte, ByVal MinVal As Byte, ByVal MaxVal As Byte) As Byte
    ClampByt = MaxByt(MinVal, MinByt(MaxVal, Value))
End Function

Public Function ClampInt(ByVal Value As Integer, ByVal MinVal As Integer, ByVal MaxVal As Integer) As Integer
    ClampInt = MaxInt(MinVal, MinInt(MaxVal, Value))
End Function

Public Function ClampLng(ByVal Value As Long, ByVal MinVal As Long, ByVal MaxVal As Long) As Long
    ClampLng = MaxLng(MinVal, MinLng(MaxVal, Value))
End Function

Public Function ClampSng(ByVal Value As Single, ByVal MinVal As Single, ByVal MaxVal As Single) As Single
    ClampSng = MaxSng(MinVal, MinSng(MaxVal, Value))
End Function

Public Function ClampDbl(ByVal Value As Double, ByVal MinVal As Double, ByVal MaxVal As Double) As Double
    ClampDbl = MaxDbl(MinVal, MinDbl(MaxVal, Value))
End Function

' ^ ############################## ^ '    Min-Max-functions    ' ^ ############################## ^ '

'Private Sub InitDedekind()
'    Dim i As Long: i = 1
'    Dedekind(i) = 3:                     i = i + 1
'    Dedekind(i) = 6:                     i = i + 1
'    Dedekind(i) = 20:                    i = i + 1
'    Dedekind(i) = 168:                   i = i + 1
'    Dedekind(i) = 7581:                  i = i + 1
'    Dedekind(i) = 7828354:               i = i + 1
'    Dedekind(i) = CDec("2414682040998"): i = i + 1
'    Dedekind(i) = CDec("56130437228687557907788"): i = i + 1
'    'Dedekind(i) = CDec("286386577668298411128469151667598498812366"): i = i + 1
'End Sub

' v ############################## v '    Fibonacci functions    ' v ############################## v '
Function FibonacciA(Optional ByVal n As Long = -1) As Long()
    If n <= 0 Then n = 46
    'Calculates the Fibonacci-series for a given number n
    ReDim fib(0 To n) As Long: fib(1) = 1
    Dim i As Long
    For i = 2 To n
        fib(i) = fib(i - 1) + fib(i - 2)
    Next
    FibonacciA = fib
End Function

Private Function FibonacciR(ByVal n As Long) As Long
    'Calculates the Fibonacci number to any given number (may be slow for values higher than 20)
    If n > 46 Then Exit Function
    If n > 1 Then FibonacciR = FibonacciR(n - 1) + FibonacciR(n - 2) Else FibonacciR = n
End Function
' ^ ############################## ^ '    Fibonacci functions    ' ^ ############################## ^ '

' v ############################## v '    additional functions    ' v ############################## v '
Public Function SinusCardinalis(ByVal x As Double) As Double ' aka sinc
    If x = 0 Then
        SinusCardinalis = 1
    Else
        SinusCardinalis = VBA.Math.Sin(x) / x
    End If
End Function

Public Function BigMul(ByVal A As Long, ByVal b As Long) As Variant
    BigMul = CDec(A) * CDec(b)
End Function
' ^ ############################## ^ '    additional functions    ' ^ ############################## ^ '


' v ############################## v '    Logarithm functions    ' v ############################## v '

'    Dim num As Double, e As Double: e = Exp(1) ' e = 2,71828182845905
'    Debug.Print e
'    Dim s As String
'    num = e * e: s = s & "LN(" & num & ")     = " & MMath.LN(num) & vbCrLf      ' = 2
'    num = 1000:  s = s & "Log10(" & num & ")  = " & MMath.Log10(num) & vbCrLf   ' = 3
'    num = 10000: s = s & "Log10(" & num & ")  = " & MMath.LogN(num) & vbCrLf    ' = 4
'    num = 32:    s = s & "Log(" & num & ", 2) = " & MMath.LogN(num, 2) & vbCrLf ' = 5
'    MsgBox s
'number          |  base        | xl-function     | result | description
'    7.389056099 |  2,718281828 | LN(Zahl)        =   2    | LN aka ln  := Logarithm to base  e
' 1000           | 10           | Log10(Zahl)     =   3    | Log10      := Logarithm to base 10, with the excelfunction LOG10
'10000           | 10           | Log(Zahl)       =   4    | Log aka lg := Logarithm to base 10, with the excelfunction Log, base not explicitely given
'   32           |  2           | Log(Zahl;Basis) =   5    | Log        := Logarithm to base  2, if the base 2 was explicitely given


'Logarithmus naturalis, logarithm to base e
Public Function LN(ByVal d As Double) As Double
    LN = VBA.Math.Log(d)
End Function

'Logarithm to the base 10
Public Function Log10(ByVal d As Double) As Double
    If d = 0 Then Exit Function
    Log10 = VBA.Math.Log(d) / VBA.Math.Log(10)
End Function

'Logarithm to a given base
Public Function LogN(ByVal x As Double, Optional ByVal base As Double = 10#) As Double
                     'base must not be 1 or 0
    If base = 1 Or base = 0 Then Exit Function
    LogN = VBA.Math.Log(x) / VBA.Math.Log(base)
End Function
'Public Shared Function Log(ByVal d As Double) As Double
'Public Shared Function Log(ByVal a As Double, ByVal newBase As Double) As Double
Public Static Function Log(ByVal d As Double, ByVal newBase As Double) As Double 'cDouble
    'Set Log = New cDouble
    Log = VBA.Math.Log(d) / VBA.Math.Log(newBase)
End Function

' ^ ############################## ^ '    Logarithm functions    ' ^ ############################## ^ '

' v ############################## v '    Rounding functions     ' v ############################## v '
Public Function Floor(ByVal A As Double) As Double
    Floor = VBA.Conversion.Int(A)
End Function

Public Function Ceiling(ByVal A As Double) As Double
    Ceiling = VBA.Conversion.Fix(A)
    If A > 0 Then Ceiling = Ceiling + 1
    'If a > 0 Then Ceiling = CDbl(Int(a) + 1#) Else Ceiling = CDbl(Fix(a))
    'If a <> 0 Then If Abs(Ceiling / a) <> 1 Then Ceiling = Ceiling + 1
End Function

Public Function RoundUp(ByVal Value As Double, Optional ByVal NumDigitsAfterDecimal As Byte = 0) As Double
    If Value < 0 Then
        RoundUp = Math.Round(Value, NumDigitsAfterDecimal)
        If Value < RoundUp Then RoundUp = RoundUp - 10 ^ -NumDigitsAfterDecimal
    Else
        RoundUp = Math.Round(Value, NumDigitsAfterDecimal)
        If RoundUp < Value Then RoundUp = RoundUp + 10 ^ -NumDigitsAfterDecimal
    End If
End Function

Public Function RoundDown(ByVal Value As Double, Optional ByVal NumDigitsAfterDecimal As Byte = 0) As Double
    If Value < 0 Then
        RoundDown = Math.Round(Value, NumDigitsAfterDecimal)
        If RoundDown < Value Then RoundDown = RoundDown + 10 ^ -NumDigitsAfterDecimal
    Else
        RoundDown = Math.Round(Value, NumDigitsAfterDecimal)
        If Value < RoundDown Then RoundDown = RoundDown - 10 ^ -NumDigitsAfterDecimal
    End If
End Function
' ^ ############################## ^ '    Rounding functions     ' ^ ############################## ^ '

' v ############################## v ' IEEE754-INFINITY functions ' v ############################## v '
' v ############################## v '      Create functions      ' v ############################## v '
'either with error handling:
Public Function GetINFE(Optional ByVal Sign As Long = 1) As Double
Try: On Error Resume Next
    GetINFE = Sgn(Sign) / 0
Catch: On Error GoTo 0
End Function

' or without error handling:
Public Function GetINF(Optional ByVal Sign As Long = 1) As Double
    Dim l(1 To 2) As Long
    If Sgn(Sign) > 0 Then
        l(2) = &H7FF00000
    ElseIf Sgn(Sign) < 0 Then
        l(2) = &HFFF00000
    End If
    Call RtlMoveMemory(GetINF, l(1), 8)
End Function

Public Sub GetNaN(ByRef Value As Double)
    Dim l(1 To 2) As Long
    l(1) = 1
    l(2) = &H7FF00000
    Call RtlMoveMemory(Value, l(1), 8)
End Sub

Public Sub GetINDef(ByRef Value As Double)
Try: On Error Resume Next
    Value = 0# / 0#
Catch: On Error GoTo 0
End Sub
' ^ ############################## ^ '     Create functions     ' ^ ############################## ^ '

' v ############################## v '      Bool functions      ' v ############################## v '
Public Function IsINDef(ByRef Value As Double) As Boolean
Try: On Error Resume Next
    IsINDef = (CStr(Value) = CStr(INDef))
Catch: On Error GoTo 0
End Function

Public Function IsNaN(ByRef Value As Double) As Boolean
    Dim b(0 To 7) As Byte
    Dim i As Long
    
    RtlMoveMemory b(0), Value, 8
    
    If (b(7) = &H7F) Or (b(7) = &HFF) Then
        If (b(6) >= &HF0) Then
            For i = 0 To 5
                If b(i) <> 0 Then
                    IsNaN = True
                    Exit Function
                End If
            Next
        End If
    End If
End Function

Public Function IsPosINF(ByVal Value As Double) As Boolean
    IsPosINF = (Value = posINF)
End Function

Public Function IsNegINF(ByVal Value As Double) As Boolean
    IsNegINF = (Value = negINF)
End Function

Public Function IsZero(Value) As Boolean
    Select Case VarType(Value)
    Case VbVarType.vbSingle:  IsZero = Abs(Value) <= EpsilonSng
    Case VbVarType.vbDouble:  IsZero = Abs(Value) <= EpsilonDbl
    Case VbVarType.vbDecimal: IsZero = Abs(Value) <= EpsilonDec
    Case Else:                IsZero = Abs(Value) <= Epsilon
    End Select
End Function

Public Function IsZeroDbl(ByVal Value As Double) As Boolean
    IsZeroDbl = Abs(Value) <= EpsilonDbl
End Function

Public Function IsZeroSng(ByVal Value As Single) As Boolean
    IsZeroSng = Abs(Value) <= EpsilonSng
End Function

Public Function IsEqual(V1, V2) As Boolean
    IsEqual = Abs(V1 - V2) <= Epsilon 'Dbl
End Function

Public Function IsEqualDbl(ByVal V1 As Double, ByVal V2 As Double) As Boolean
    IsEqualDbl = Abs(V1 - V2) <= EpsilonDbl
End Function

Public Function IsEqualSng(ByVal V1 As Single, ByVal V2 As Single) As Boolean
    IsEqualSng = Abs(V1 - V2) <= EpsilonSng
End Function

Public Function IsOdd(ByVal Value As Long) As Boolean
    IsOdd = Value And 1& ' Mod 2 <> 0
End Function

Public Function IsEven(ByVal Value As Long) As Boolean
    IsEven = Not Value And 1& ' Mod 2 = 0
End Function

' ^ ############################## ^ '      Bool functions      ' ^ ############################## ^ '

' v ############################## v '    Bit-Shifting functions    ' v ############################## v '

' this is the Murphy McCauley method which I modified slightly, http://www.fullspectrum.com/deeth/
Private Sub SubstituteCode(StoreHere() As Byte, CodeString As String, ByVal AddressOfFunctionToReplace As Long)
    Dim OldProtection As Long
    Dim s As String
    Dim i As Long
      
    ReDim StoreHere(Len(CodeString) \ 2 - 1)
    
    For i = 0 To Len(CodeString) \ 2 - 1
        StoreHere(i) = Val("&H" & Mid$(CodeString, i * 2 + 1, 2))
    Next
    
    VirtualProtect ByVal AddressOfFunctionToReplace, 21, PAGE_EXECUTE_READWRITE, OldProtection
    RtlMoveMemory ByVal AddressOfFunctionToReplace, &H90, 1 ' nop to insure our first line is not concated with the previous instruction
    RtlMoveMemory ByVal AddressOfFunctionToReplace + 1, StoreHere(0), 20 ' shr/shl code substitution
    VirtualProtect ByVal AddressOfFunctionToReplace, 21, OldProtection, OldProtection
    
    ' alternately, if the code is much longer use this instead:
    
    ' VirtualProtect ByVal AddressOfFunctionToReplace, 7, PAGE_EXECUTE_READWRITE, OldProtection
    ' RtlMoveMemory ByVal AddressOfFunctionToReplace, &HB8, 1  ' mov eax, PointerToCode
    ' RtlMoveMemory ByVal AddressOfFunctionToReplace + 1, Varptr(StoreHere(0)),4
    ' RtlMoveMemory ByVal AddressOfFunctionToReplace + 5, &HE0FF&, 2 ' jmp eax
    ' VirtualProtect ByVal AddressOfFunctionToReplace, 7, OldProtection, OldProtection
End Sub

' Leave these placeholder functions, and their code
Public Function ShiftLeft(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20001215
    Dim mask As Long
    
    Select Case ShiftCount
    Case 1 To 31
        ' mask out bits that are pushed over the edge anyway
        mask = Pow2(31 - ShiftCount)
        ShiftLeft = Value And (mask - 1)
        ' shift
        ShiftLeft = ShiftLeft * Pow2(ShiftCount)
        ' set sign bit
        If Value And mask Then
          ShiftLeft = ShiftLeft Or &H80000000
        End If
    Case 0
        ' ret unchanged
        ShiftLeft = Value
    End Select
End Function

Public Function ShiftRightZ(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20001215
    Select Case ShiftCount
    Case 1 To 31
        If Value And &H80000000 Then
            ShiftRightZ = (Value And Not &H80000000) \ 2
            ShiftRightZ = ShiftRightZ Or &H40000000
            ShiftRightZ = ShiftRightZ \ Pow2(ShiftCount - 1)
        Else
            ShiftRightZ = Value \ Pow2(ShiftCount)
        End If
    Case 0
        ' ret unchanged
        ShiftRightZ = Value
    End Select
End Function

Public Static Function ShiftRight(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20011009
    Dim lPow2(0 To 30) As Long
    Dim i As Long
    
    Select Case ShiftCount
    Case 0:      ShiftRight = Value
    Case 1 To 30
        If i = 0 Then
            lPow2(0) = 1
            For i = 1 To 30
                lPow2(i) = 2 * lPow2(i - 1)
            Next
        End If
        If Value And &H80000000 Then
            ShiftRight = Value \ lPow2(ShiftCount)
            If ShiftRight * lPow2(ShiftCount) <> Value Then
                ShiftRight = ShiftRight - 1
            End If
        Else
            ShiftRight = Value \ lPow2(ShiftCount)
        End If
    Case 31
        If Value And &H80000000 Then
            ShiftRight = -1
        Else
            ShiftRight = 0
        End If
    End Select
End Function
' ^ ############################## ^ '    Bit-Shifting functions    ' ^ ############################## ^ '

' v ############################## v '     Output functions     ' v ############################## v '
Public Function INDefToString() As String
    On Error Resume Next
    INDefToString = CStr(INDef)
    On Error GoTo 0
End Function

Public Function NaNToString() As String
    On Error Resume Next
    If App.LogMode = 0 Then
        NaNToString = "1.#QNAN"
    Else
        NaNToString = CStr(NaN)
    End If
    On Error GoTo 0
End Function

Public Function PosINFToString() As String
    PosINFToString = CStr(posINF)
End Function

Public Function NegINFToString() As String
    NegINFToString = CStr(negINF)
End Function
' ^ ############################## ^ '     Output functions     ' ^ ############################## ^ '

' v ############################## v '      Input function      ' v ############################## v '
'nope from 2024-07-14 on ou can find this function inside MString
'Public Function Double_TryParse(s As String, Value_out As Double) As Boolean
'Try: On Error GoTo Catch
'    If Len(s) = 0 Then Exit Function
'    s = Replace(s, ",", ".")
'    If StrComp(s, "1.#QNAN") = 0 Then
'        GetNaN Value_out
'    ElseIf StrComp(s, "1.#INF") = 0 Then
'        Value_out = GetINF
'    ElseIf StrComp(s, "-1.#INF") = 0 Then
'        Value_out = GetINF(-1)
'    ElseIf StrComp(s, "-1.#IND") = 0 Then
'        GetINDef Value_out
'    Else
'        Value_out = Val(s)
'    End If
'    Double_TryParse = True
'Catch:
'End Function
' ^ ############################## ^ '       Input function       ' ^ ############################## ^ '
' ^ ############################## ^ ' IEEE754-INFINITY functions ' ^ ############################## ^ '

' v ############################## v '      Trigonometric functions      ' v ############################## v '

Public Function DegToRad(ByVal angleInDegrees As Double) As Double
    DegToRad = angleInDegrees * MMath.Pi / 180#
End Function
Public Function RadToDeg(ByVal angleInRadians As Double) As Double
    RadToDeg = angleInRadians * 180# / MMath.Pi
End Function


Public Static Function Sin(ByVal A As Double) As Double      'aka Sinus
    Sin = VBA.Math.Sin(A)
End Function
Public Static Function Cos(ByVal A As Double) As Double      'aka Cosinus
    Cos = VBA.Math.Cos(A)
End Function
Public Static Function Tan(ByVal A As Double) As Double      'aka Tangens
    Tan = VBA.Math.Tan(A)
End Function

Public Static Function Csc(ByVal A As Double) As Double      'aka Cosecans
    Csc = 1 / VBA.Math.Sin(A)
End Function
Public Static Function Sec(ByVal A As Double) As Double      'aka Secans
    Sec = 1 / VBA.Math.Cos(A)
End Function
Public Static Function Cot(ByVal A As Double) As Double      'aka Cotangens
    Cot = 1 / VBA.Math.Tan(A)
End Function


Public Static Function ASin(ByVal d As Double) As Double     'aka ArcusSinus
    ASin = Atn(d / (Sqr(1 - d ^ 2)))
End Function
Public Static Function ACos(ByVal d As Double) As Double     'aka ArcusCosinus
    ACos = (3.14159265358979 / 2) - Atn(d / (Sqr(1 - d ^ 2)))
End Function
Public Static Function ATan(ByVal d As Double) As Double     'aka ArcusTangens
    ATan = VBA.Math.Atn(d)
End Function

Public Static Function ACsc(ByVal y As Double) As Double     'aka ArcusCosecans
    ACsc = ASin(1 / y)
End Function
Public Static Function ASec(ByVal x As Double) As Double     'aka ArcusSecans
    ASec = ACos(1 / x)
End Function
Public Static Function ACot(ByVal t As Double) As Double   'aka ArcusCotangens
    ACot = Pi * 0.5 - VBA.Math.Atn(t)
End Function


Public Static Function Sinh(ByVal Value As Double) As Double 'aka SinusHyperbolicus
    Sinh = (VBA.Math.Exp(Value) - VBA.Math.Exp(-Value)) / 2
End Function
Public Static Function Cosh(ByVal Value As Double) As Double 'aka CosinusHyperbolicus
    Cosh = (VBA.Math.Exp(Value) + VBA.Math.Exp(-Value)) / 2
End Function
Public Static Function Tanh(ByVal Value As Double) As Double 'aka TangensHyperbolicus
    Tanh = (VBA.Math.Exp(Value) - VBA.Math.Exp(-Value)) / (VBA.Math.Exp(Value) + VBA.Math.Exp(-Value))
End Function

Public Static Function CscH(ByVal y As Double) As Double     'aka CosecansHyperbolicus
    CscH = 2 / (VBA.Math.Exp(y) - VBA.Math.Exp(-y))
End Function
Public Static Function SecH(ByVal x As Double) As Double     'aka SecansHyperbolicus
    SecH = 2 / (Exp(x) + Exp(-x))
End Function
Public Static Function CotH(ByVal t As Double) As Double     'aka CotangensHyperbolicus
    CotH = (VBA.Math.Exp(t) + VBA.Math.Exp(-t)) / (VBA.Math.Exp(t) - VBA.Math.Exp(-t))
End Function

Public Static Function ArSinH(ByVal y As Double) As Double   'aka AreaSinusHyperbolicus
    ArSinH = VBA.Math.Log(y + Sqr(y * y + 1))
End Function
Public Static Function ArCosH(ByVal x As Double) As Double   'aka AreaCosinusHyperbolicus
    ArCosH = VBA.Math.Log(x + Sqr(x * x - 1))
End Function
Public Static Function ArTanH(ByVal t As Double) As Double   'aka AreaTangensHyperbolicus
    ArTanH = VBA.Math.Log((1 + t) / (1 - t)) / 2
End Function

Public Static Function ArCscH(ByVal x As Double) As Double   'aka AreaCosecansHyperbolicus
    ArCscH = VBA.Math.Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
End Function
Public Static Function ArSecH(ByVal x As Double) As Double   'aka AreaSecansHyperbolicus
    ArSecH = VBA.Math.Log((Sqr(-x * x + 1) + 1) / x)
End Function
Public Static Function ArCotH(ByVal x As Double) As Double   'aka AreaCotangensHyperbolicus
    ArCotH = VBA.Math.Log((x + 1) / (x - 1)) / 2
End Function


'Public Shared Function Abs(ByVal value As Decimal) As Decimal
'Public Shared Function Abs(ByVal value As Double) As Double
'Public Shared Function Abs(ByVal value As Integer) As Integer
'Public Shared Function Abs(ByVal value As Long) As Long
'Public Shared Function Abs(ByVal value As Short) As Short
'Public Shared Function Abs(ByVal value As Single) As Single
'Public Shared Function Abs(ByVal value As System.SByte) As System.SByte
Public Function Abs_(ByVal varValue As Variant) As Variant
    Abs_ = VBA.Math.Abs(varValue)
End Function

'Und was ist mit ACot ????? =Pi/2 - Atan(x)

'Public Shared Function Atan2(ByVal y As Double, ByVal x As Double) As Double
'Public Static Function ATan2(ByVal y As Double, ByVal x As Double) As Double 'cDouble
'    'Set Atan2 = New cDouble
'    ATan2 = Atn(y / x)
'End Function

'Public Shared Function BigMul(ByVal a As Integer, ByVal b As Integer) As Long
'Public Static Function BigMul(ByVal a As Long, ByVal b As Long) As Variant 'As Long
'    'vergiss es
'    BigMul = CVar(a) * CVar(b)
'End Function

'Public Shared Function Ceiling(ByVal a As Double) As Double
'Public Static Function Ceiling(ByVal a As Double) As Double 'cDouble
'    'Set Ceiling = New cDouble
'    Ceiling = Int(a)
'End Function



'Public Shared Function DivRem(ByVal a As Integer, ByVal b As Integer, ByRef result As Integer) As Integer
'Public Shared Function DivRem(ByVal a As Long, ByVal b As Long, ByRef result As Long) As Long
Public Static Function DivRem(ByVal A As Long, ByVal b As Long, ByRef Result As Long) As Long 'cInteger 'Long
    'Set DivRem = New cInteger
End Function

'Public Shared Function Exp(ByVal d As Double) As Double
Public Static Function Exp(ByVal d As Double) As Double 'cDouble
    'Set Exp = New cDouble
    Exp = VBA.Math.Exp(d)
End Function

'Public Shared Function Floor(ByVal d As Double) As Double
'Public Function Floor(ByVal d As Double) As Double 'cDouble
'    'Set Floor = New cDouble
'End Function

'Public Shared Function IEEERemainder(ByVal x As Double, ByVal y As Double) As Double
'Public Static Function IEEERemainder(ByVal x As Double, ByVal y As Double) As Double 'cDouble
'    'Set IEEERemainder = New cDouble
'End Function


'Public Shared Function Sign(ByVal value As Decimal) As Integer
'Public Shared Function Sign(ByVal value As Double) As Integer
'Public Shared Function Sign(ByVal value As Integer) As Integer
'Public Shared Function Sign(ByVal value As Long) As Integer
'Public Shared Function Sign(ByVal value As Short) As Integer
'Public Shared Function Sign(ByVal value As Single) As Integer
'Public Shared Function Sign(ByVal value As System.SByte) As Integer
Public Static Function Sign(ByVal varValue As Variant) As Variant
    Sign = Sgn(varValue)
End Function


'Public Shared Function Sqrt(ByVal d As Double) As Double
Public Static Function Sqrt(ByVal d As Double) As Double
    Sqrt = VBA.Math.Sqr(d)
End Function

' ^ ############################## ^ '      Trigonometric functions      ' ^ ############################## ^ '

'#######  for Bit Shifting ##########
' v ############################## v '    Bit Shifting functions    ' v ############################## v '
Public Function ShL(Shifting As Long, Shifter As Long) As Long
    ShL = ShiftLeft(Shifting, Shifter)
End Function
Public Function ShRz(Shifting As Long, Shifter As Long) As Long
    ShRz = ShiftRightZ(Shifting, Shifter)
End Function
Public Function ShR(Shifting As Long, Shifter As Long) As Long
    ShR = ShiftRight(Shifting, Shifter)
End Function

Public Sub Increment(ByRef LngVal As Long) 'As Long
    LngVal = LngVal + 1
End Sub

Public Sub Decrement(ByRef LngVal As Long) 'As Long
    LngVal = LngVal - 1
End Sub
' ^ ############################## ^ '    Bit Shifting functions    ' ^ ############################## ^ '


' v ############################## v '     solving quadratic & cubic formula     ' v ############################## v '
Public Function Quadratic(ByVal A As Double, ByVal b As Double, ByVal c As Double, ByRef x1_out As Double, ByRef x2_out As Double) As Boolean
Try: On Error GoTo Catch
    
    'maybe the midnight-formula:
    'Dim w As Double: w = VBA.Sqr(b ^ 2 - 4 * a * c)
    'x1_out = (-b + w) / (2 * a)
    'x2_out = (-b - w) / (2 * a)
    
    'or maybe the pq-formula
    Dim p    As Double:   p = b / A
    Dim q    As Double:   q = c / A
    Dim p_2  As Double: p_2 = p / 2
    Dim p2_4 As Double: p2_4 = p_2 * p_2
    Dim W As Double: W = VBA.Sqr(p2_4 - q)
    
    x1_out = -p_2 + W
    x2_out = -p_2 - W
    
    Quadratic = True
    Exit Function
Catch:
End Function

Public Function Quadratic_ToStr(ByVal A As Double, ByVal b As Double, ByVal c As Double) As String
    Quadratic_ToStr = A & "x²" & GetOp(b) & "x" & GetOp(c) & " = 0"
End Function
Public Function Cubic_ToStr(ByVal A As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As String
    Cubic_ToStr = A & "x³" & GetOp(b) & "x²" & GetOp(c) & "x" & GetOp(d) & " = 0"
End Function
Private Function GetOp(ByVal v As Double) As String
    Select Case v
    Case -1:     GetOp = " - ": Exit Function
    Case 0:      Exit Function
    Case 1:      GetOp = " + ": Exit Function
    Case Else:   If v < 0 Then GetOp = " - " & Abs(v) Else GetOp = " + " & v
    End Select
End Function

'https://www.youtube.com/watch?v=xhjNRQxqJTM '&t=116s
'https://www.youtube.com/watch?v=q14F6fZf5kc '&t=1658s
'https://www.youtube.com/watch?v=N-KXStupwsc '&t=4s
Public Function Cubic(ByVal A As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, ByRef x1_out As Double, ByRef x2_out As Double, ByRef i2_out As Double, ByRef x3_out As Double, ByRef i3_out As Double) As Boolean
Try: On Error GoTo Catch
    'Scipione del Ferro (1465-1526), Nicolo Tartaglia (1500-1557), Gerolamo Cardano (1501-1576)
    'Rafael Bombelli (1526-1572)
    If A = 0 Then
        Cubic = Quadratic(b, c, d, x1_out, x2_out)
        Exit Function
    End If
    Dim a2 As Double: a2 = A * A
    Dim a3 As Double: a3 = A * a2
    Dim b2 As Double: b2 = b * b
    Dim b3 As Double: b3 = b * b2
    'Dim bc As Double: bc = b * c
    Dim b3_27a3 As Double: b3_27a3 = b3 / (27 * a3)
    'Dim bc_6a2  As Double:  bc_6a2 = bc / (6 * a2)
    Dim cb_3a2  As Double:  cb_3a2 = c * b / (3 * a2)
    Dim b2_3a2  As Double:  b2_3a2 = b2 / (3 * a2)
    'Dim b2_9a2  As Double:  b2_9a2 = b2 / (9 * a2)
    Dim d_a  As Double: d_a = d / A
    'Dim d_2a As Double: d_2a = d / (2 * a)
    Dim c_a  As Double: c_a = c / A
    'Dim c_3a As Double: c_3a = c / (3 * a)
    Dim b_3a As Double: b_3a = b / (3 * A)
    
    'Dim w As Double: w = VBA.Sqr((-b3_27a3 + bc_6a2 - d_2a) ^ 2 + (-b2_9a2 + c_3a) ^ 3)

    'x1_out = (-b3_27a3 + bc_6a2 - d_2a + w) ^ (1 / 3) + _
    '        (-b3_27a3 + bc_6a2 - d_2a - w) ^ (1 / 3) - b_3a
    
    'x2_out
    
'Discriminant D: q²/4 + p³/27 > = < 0
'1: q²/4 + p³/27 > 0
'   1 real
'   2 complex
'2: q²/4 + p³/27 = 0
'   2 real = 3 real overall, with 2 repeated: where a local minimum or local maximum touches the x-axis
'3:  q²/4 + p³/27 < 0
'   3 real all distinct
'
    
    Dim p   As Double:   p = -b2_3a2 + c_a
    Dim q   As Double:   q = 2 * b3_27a3 - cb_3a2 + d_a
    Dim DD  As Double:  DD = (q ^ 2) / 4 + (p ^ 3) / 27
    Dim q_2 As Double: q_2 = q / 2
    Dim p_3 As Double: p_3 = p / 3
    
    Select Case DD
    Case Is < 0
    Case Is = 0
    Case Is > 0
    End Select
    
    'Dim b_3a As Double: b_3a = b / (3 * a)
    Dim W As Double
    W = VBA.Math.Sqr(q_2 ^ 2 + p_3 ^ 3) '^ (1 / 2)
    
    'x1_out = -b_3a + (-q_2 + W) ^ (1 / 3) + (-q_2 - W) ^ (1 / 3)
    
    x1_out = -b_3a + CubeRoot(-q_2 + W) + CubeRoot(-q_2 - W) '^ (1 / 3)    '^ (1 / 3)
    
    Cubic = True
    Exit Function
Catch:
End Function

Public Function SqrH(ByVal n As Double) As Double
    'SquareRoot due to Heron algorithm
    Dim s As Long:   s = Sgn(n): n = Abs(n)
    Dim r As Double: r = n
    Dim i As Long
    SqrH = 1
    Do
        i = i + 1
        SqrH = (SqrH + r) / 2
        r = n / SqrH
        If (SqrH - r) < 0.0000000000001 Then Exit Do
        If i = 20 Then Exit Do
    Loop
    'Debug.Print i
End Function

'Public Function CubRt(ByVal v As Double, ByRef i_out As Double) As Double
Public Function CubeRoot(ByVal d As Double) As Double
    'CubeRoot due to Halley
    Dim a3 As Double
    Dim t As TDouble: t.Value = d
    Dim p As TLong2:   LSet p = t: p.Value1 = p.Value1 \ 3 + 715094163
    Dim A As Double:   LSet t = p: A = t.Value
    a3 = A * A * A
    A = A * (a3 + d + d) / (a3 + a3 + d)
    a3 = A * A * A
    A = A * (a3 + d + d) / (a3 + a3 + d)
    a3 = A * A * A
    CubeRoot = A * (a3 + d + d) / (a3 + a3 + d)
    
End Function

'/// This Program will find the cube root of any number n.
'#include<iostream>
'#include<math.h>
'using namespace std;
'double nrAlgorithm(double m,double n)
'{
'    if(n==0)
'        return 0;
'    double g=1,x=1;
'    int i=0;
'    While (True)
'    {
'        x=((m-1)*pow(x,m)+n)/(m*pow(x,m-1));
'        if(x==g)
'        {
'            break;
'        }
'            g=x;
'        i++;
'    }
'    return g;
'
'}
'int main()
'{
'    int t,n;
'    cin>>t;
'    while(t--)
'    {
'        cin>>n;
'        cout<<nrAlgorithm(3,n)<<endl;
'    }
'}

Public Function PascalTriangle(ByVal nrows As Integer) As Variant()
    If nrows > 1030 Then
        MsgBox nrows & ": overflow!" & vbCrLf & "The pascal-triangle will be computed with a maximum of 1030 rows."
        'Exit Function
    End If
    nrows = Min(nrows, 1030)
'We make an 1d-array with variants for every rows each row is a variant wchich is an array of arising size
    ReDim p(1 To nrows)
    Dim r() 'As Long
    'create the first row which contains one array with only one element with the number one
    ReDim r(1 To 1): r(1) = 1
    p(1) = r
    'create the second row which contains one array with two element each with the number one
    ReDim r(1 To 2): r(1) = 1: r(2) = 1
    p(2) = r
    Dim i As Long, j As Long
    For i = 3 To nrows
        ReDim r(1 To i)
        r(1) = 1: r(i) = 1
        For j = 2 To i - 1
            r(j) = p(i - 1)(j - 1) + p(i - 1)(j)
        Next
        p(i) = r
    Next
    PascalTriangle = p
End Function

Public Function PascalTriangle_ToStr(p() As Variant) As String
    Dim nrows As Long: nrows = UBound(p)
    ReDim sa(1 To nrows) As String
    sa(1) = "1"
    sa(2) = "1 1"
    '': s = "1" & vbCrLf & "1  1" & vbCrLf
    Dim i As Long, j As Long
    If nrows < 3 Then
        PascalTriangle_ToStr = Join(sa, vbCrLf)
        Exit Function
    End If
    Dim s As String
    For i = 3 To nrows
        s = ""
        For j = 1 To i
            s = s & CStr(p(i)(j)) & IIf(j < nrows, " ", "")
        Next
        sa(i) = s
    Next
    PascalTriangle_ToStr = Join(sa, vbCrLf)
End Function
'
'                        1
'                      1   1
'                    1   2   1
'                  1   3   3   1
'                1   4   6   4   1
'              1   5  10  10   5   1
'            1   6  15  20  15   6   1
'          1   7  21  25  25  21   7   1
'        1   8  28  46  50  46  28   8   1
'      1   9  36  74  96  96  74  36   9   1
'    1  10  45 120 170 192 170 120  45  10   1
'  1  11  55 165 330 362 462 330 165  55  11   1
'1  12  66 220 495 792 924 792 495 220  66  12   1


' v ############################## v '    Complex numbers    ' v ############################## v '
'Many thanks to Loay
'https://www.youtube.com/watch?v=kIjgFYLymJw
'x²+1=0 => x²=-1; x_1,2 = +-sqrt(-1); sqrt(-1) = i; i² = -1; (-i)² = -1
'Many thanks to MathePeter
'https://www.youtube.com/watch?v=zB2VwWzpYx4
'https://www.youtube.com/watch?v=G_FRNyHpzrk
'https://www.youtube.com/watch?v=Z-mKZECwOgg
Public Function Real_ToComplex(ByVal v As Double) As Complex
    Real_ToComplex.Re = v
    'Real_ToComplex.Im = 0
End Function

Public Function Real_ToComplexP(ByVal v As Double) As ComplexP
    With Real_ToComplexP: .r = v: .phi = CDbl(Pi2): End With
End Function

Public Function Complex(ByVal Re As Double, ByVal Im As Double) As Complex
    With Complex: .Re = Re: .Im = Im: End With
End Function

Public Function Complex_ToStr(z As Complex) As String
    'With z: Complex_ToStr = .Re & " + i*" & .Im: End With
    With z: Complex_ToStr = .Re & " + " & .Im & "*i": End With
End Function

'Verschiebung, Translations:
Public Function Complex_Add(z1 As Complex, z2 As Complex) As Complex
    With Complex_Add:       .Re = z1.Re + z2.Re: .Im = z1.Im + z2.Im:    End With
End Function

Public Function Complex_Subt(z1 As Complex, z2 As Complex) As Complex
    With Complex_Subt:      .Re = z1.Re - z2.Re: .Im = z1.Im - z2.Im:    End With
End Function

'Streckung/Stauchung oder Drehung, Rotation
Public Function Complex_Mul(z1 As Complex, z2 As Complex) As Complex
    With Complex_Mul
        .Re = z1.Re * z2.Re - (z1.Im * z2.Im)
        .Im = z1.Im * z2.Re + z1.Re * z2.Im
    End With
End Function

Public Function Complex_Div(z1 As Complex, z2 As Complex) As Complex
    Dim d As Double: d = Complex_Abs2(z2)
    Dim z2_ As Complex: z2_ = Complex_Conj(z2)
    Complex_Div = Complex_Mul(z1, z2_)
    With Complex_Div
        .Re = .Re / d
        .Im = .Im / d
    End With
End Function

'Spiegelung
'complex conjugation
Public Function Complex_Neg(z As Complex) As Complex
    'mirroring at center (0,0)
    With Complex_Neg:       .Re = -z.Re:         .Im = -z.Im:    End With
End Function

Public Function Complex_Conj(z As Complex) As Complex
    'mirroring at x-axis
    With Complex_Conj:      .Re = z.Re:          .Im = -z.Im:    End With
End Function

Public Function Complex_NegConj(z As Complex) As Complex
    'mirroring at y-axis
    With Complex_NegConj:   .Re = -z.Re:         .Im = z.Im:    End With
End Function

Public Function Complex_Abs(z As Complex) As Double
    Complex_Abs = VBA.Math.Sqr(Complex_Abs2(z))
End Function

Public Function Complex_Abs2(z As Complex) As Double
    With z: Complex_Abs2 = .Re * .Re + .Im * .Im:    End With
End Function

Public Function Complex_ToComplexP(c As Complex) As ComplexP
    With Complex_ToComplexP
        .r = VBA.Math.Sqr(Abs(c.Re * c.Re + c.Im * c.Im))
        .phi = ATan2(c.Im, c.Re)
        'oder:
        '.phi = Sgn(c.Im) * Arccos(c.Re / .r)
    End With
End Function

'polar form
Public Function ComplexP(ByVal r As Double, ByVal phi As Double) As ComplexP
    With ComplexP: .r = r: .phi = phi: End With
End Function

Public Function ComplexP_ToStr(p As ComplexP) As String
    With p
        'ComplexP_ToStr = .r & " + "e^(i*phi
        ComplexP_ToStr = .r & " +(cos(" & .phi & ")+sin(" & .phi & ")*i)"
    End With
End Function

'euler-form
Public Function ComplexP_ToStrE(p As ComplexP) As String
    With p
        ComplexP_ToStrE = .r & " * e^(" & .phi & "*i)"
    End With
End Function

Public Function ComplexP_Add(P1 As ComplexP, P2 As ComplexP) As ComplexP
    'Dim z1 As Complex: z1 = ComplexP_ToComplex(p1)
    'Dim z2 As Complex: z2 = ComplexP_ToComplex(p2)
    'Dim z3 As Complex: z2 = Complex_Add(z1, z2)
    'ComplexP_Add = Complex_ToComplexP(z3)
    ComplexP_Add = Complex_ToComplexP(Complex_Add(ComplexP_ToComplex(P1), ComplexP_ToComplex(P2)))
End Function

Public Function ComplexP_Mul(P1 As ComplexP, P2 As ComplexP) As ComplexP
    With ComplexP_Mul
        .r = P1.r * P2.r
        .phi = P1.phi + P2.phi
    End With
End Function

'r*e^(phi+2pi*k); k € Z (ganze Zahl)
'z^n = r^n * e^(i*(n*phi+2pi*k*n)); k gnd n sind ganze zahlen
Public Function ComplexP_Powi(p As ComplexP, ByVal expon As Long) As ComplexP
    'bei Ganzzahligen Exponenten spricht man vom Satz von Moivre, Potenzgesetze
    Dim ppi2 As Double: ppi2 = 8 * VBA.Math.Atn(1)
    With ComplexP_Powi
        .r = p.r ^ expon
        .phi = ModDbl(expon * p.phi, ppi2)
        'if .phi<0
    End With
End Function

'n = p / q € Q : rationale Zahlen
'Exponent n = p / q
'p € Z, q € N>=2
'k=0,1,...,q-1
'Z = ganze negative und positive Zahlen inkl 0
'N = ganze nur positive Zahlen
Public Function ComplexP_Pow(p As ComplexP, ByVal expon_p As Long, ByVal expon_q As Long) As ComplexP()
    Dim ccp() As ComplexP
    If expon_q < 2 Then
        If expon_q = 1 Then
            ReDim ccp(0 To 0): ccp(0) = ComplexP_Powi(p, expon_p)
            ComplexP_Pow = ccp()
        End If
        Exit Function 'q € N,
    End If
    If Not MMath.CancelFraction(expon_p, expon_q) Then
        '
        Exit Function
    End If
    If ggT(expon_p, expon_q) <> 1 Then
        '
        Exit Function
    End If
    Dim ppi2 As Double: ppi2 = Pi2 ' 8 * VBA.Math.Atn(1)
    Dim pq   As Double:   pq = expon_p / expon_q
    Dim r    As Double:    r = p.r ^ pq
    'Dim phi As Double:   phi = ModDbl(expon_q * p.phi, ppi2)
    Dim phi As Double:   phi = p.phi * pq
    Dim dp  As Double:    dp = ppi2 / expon_q
    ReDim ccp(0 To expon_q - 1)
    ccp(0) = ComplexP(r, phi)
    Dim i As Long
    For i = 1 To expon_q - 1
        phi = phi + dp
        ccp(i) = ComplexP(r, phi)
    Next
    ComplexP_Pow = ccp
End Function

Public Function ComplexP_NthRoot(p As ComplexP, ByVal n As Long) As ComplexP()
    'berechnet die n-te Wurzel einer komplexem zahl p
    'computes the nth-root of the complex number p
    Dim ccp() As ComplexP
    If n = 0 Then Exit Function 'q € N w/o 0
    If n = 1 Then
        ReDim ccp(0 To 0): ccp(0) = p
        ComplexP_NthRoot = ccp()
        Exit Function
    End If
    Dim N_1  As Double:  N_1 = 1 / n
    Dim r    As Double:    r = p.r ^ N_1
    Dim phi  As Double:  phi = p.phi * N_1
    Dim ppi2 As Double: ppi2 = Pi2
    Dim dp   As Double:   dp = ppi2 / n
    ReDim ccp(0 To n - 1)
    ccp(0) = ComplexP(r, phi)
    Dim i As Long
    For i = 1 To n - 1
        phi = phi + dp
        ccp(i) = ComplexP(r, phi)
    Next
    ComplexP_NthRoot = ccp
End Function

Public Function ComplexP_ToComplex(p As ComplexP) As Complex
    With ComplexP_ToComplex
        .Re = p.r * VBA.Math.Cos(p.phi)
        .Im = p.r * VBA.Math.Sin(p.phi)
    End With
End Function

' ^ ############################## ^ '    Complex numbers    ' ^ ############################## ^ '

' v ############################## v '  Modulo op on floats  ' v ############################## v '
Public Function ModF(ByVal Value As Double, ByVal div As Double) As Double
   ModF = Value - (Int(Value / div) * div)
End Function

Public Function ModDbl(v As Double, d As Double) As Double
    'berechnet das Überbleibsel bei der division
    Dim s As Long: s = Sgn(v)
    v = Abs(v)
    Dim i As Long: i = Int(v / d)
    ModDbl = s * (v - i * d)
End Function
' ^ ############################## ^ '  Modulo op on floats  ' ^ ############################## ^ '

' v ############################## v '   calculation of Pi   ' v ############################## v '
Public Function CalcPi() 'As Variant 'As Decimal
    Dim sqr3: sqr3 = CDec(SquareRoot3)  'CDec("1,7320508075688772935274463415") '058723669428052538103806280558069794519330169088000370811461867572485756")
    Dim sum: sum = CDec(0)
    'On Error Resume Next
    Dim n As Long
    For n = 1 To 11 '40
        sum = sum + -Fact(CDec(2) * CDec(n) - CDec(2)) / (CDec(2) ^ (CDec(4) * CDec(n) - CDec(2)) * Fact(CDec(n) - CDec(1)) ^ CDec(2) * (CDec(2) * CDec(n) - CDec(3)) * (CDec(2) * CDec(n) + CDec(1)))
    Next
    Dim Pi 'As Double
    Pi = CDec(3) * sqr3 / CDec(4) + CDec(24) * CDec(sum)
    CalcPi = Pi
End Function
' ^ ############################## ^ '   calculation of Pi   ' ^ ############################## ^ '

' v ############################## v '  unsigned arithmetic  ' v ############################## v '

Public Function UnsignedAdd(ByVal Value As Long, ByVal Incr As Long) As Long
' This function is useful when doing pointer arithmetic,
' but note it only works for positive values of Incr
' all credits for this function are going to the incredible Steve McMahon aka vbAccelerator
   If Value And &H80000000 Then  'Start < 0
       UnsignedAdd = Value + Incr
   ElseIf (Value Or &H80000000) < -Incr Then
       UnsignedAdd = Value + Incr
   Else
       UnsignedAdd = (Value + &H80000000) + (Incr + &H80000000)
   End If
   
End Function

' ^ ############################## ^ '  unsigned arithmetic  ' ^ ############################## ^ '

' v ############################## v ' Celsius, Fahrenheit, Kelvin ' v ############################## v '
'Public Const TempCelsius_AbsolutZero As Double = 273.15

'e.g.: -40 °C = -40 °F
Public Function TempCelsius_ToFahrenheit(ByVal TempCelsius As Double) As Double
    TempCelsius_ToFahrenheit = TempCelsius * 1.8 + 32
End Function
Public Function TempCelsius_ToKelvin(ByVal TempCelsius As Double) As Double
    TempCelsius_ToKelvin = 273.15 + TempCelsius
End Function


Public Function TempFahrenheit_ToCelsius(ByVal TempFahrenheit As Double) As Double
    TempFahrenheit_ToCelsius = (TempFahrenheit - 32) * 5 / 9
End Function
Public Function TempFahrenheit_ToKelvin(ByVal TempFahrenheit As Double) As Double
    TempFahrenheit_ToKelvin = (TempFahrenheit - 32) * 5 / 9 + 273.15
End Function


Public Function TempKelvin_ToCelsius(ByVal TempKelvin As Double) As Double
    TempKelvin_ToCelsius = TempKelvin - 273.15
End Function
Public Function TempKelvin_ToFahrenheit(ByVal TempKelvin As Double) As Double
    TempKelvin_ToFahrenheit = (TempKelvin - 273.15) * 1.8 + 32
End Function
' ^ ############################## ^ ' Celsius, Fahrenheit, Kelvin ' ^ ############################## ^ '

