Attribute VB_Name = "MMath"
Option Explicit ' OM: 2024-03-04 lines 1264

Public INDef  As Double
Public posINF As Double
Public negINF As Double
Public NaN    As Double

'Complex number in cartesian coordinates
Public Type Complex
    Re As Double
    Im As Double
End Type

'Complex number in polar coordinates or euler form
Public Type ComplexP
    r   As Double
    phi As Double
End Type

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bLength As Long)

'value range Byte (unsigned int8)
'0-255

'value range Integer (signed int16)
'-32768 - 32767


'value range Currency (Int64)
'    Dim Cur1: bigDec1 = CDec("9223372036854775807")   ' No overflow.
'    Dim Cur2: bigDec2 = CDec("9223372036854775808")   ' No Overflow.
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
Public SpeedOfLight   ' Lichtgeschwindigkeit im Vakuum      c   = 299792458 m/s
Public ElemCharge     ' Elementarladung (des Protons)       e   = 1,602176634 * 10^-19 C (Coulomb)
Public MassElektron   ' Ruhemasse des Elektrons             m_e = 9,109*10^-31 kg
Public MassProton     ' Ruhemasse des Protons               m_p = 1,6726215813 · 10^-27 kg
Public PlanckQuantum  ' Plancksches Wirkungsquantum         h   = 6,62607015 * 10^(-37) m² * kg / s
Public Avogadro       ' Avogadro-Konstante                  N_A = 6,022 * 10^23
Public Gravitation    ' Newtonsche Gravitationskonstante    G   = 6,6743 * 10^-11 m³ / (kg * s²)
Public BoltzmannConst ' Boltzmann-Konstante                 k_B = 1,38064852 × 10-23 m2 kg s-2 K-1
Public MagnPermittvy  ' magnetische Feldkonstante           mue_0 = µ0 ˜ 1.2566370621219 * 10 ^(-6) N/A²
Public ElecPermittvy  ' elektrische Feldkonstante           eps_0 = 8.8541878128(13)e-12 (A s)/(V m)
Public QuantumAlpha   ' FineStructureConstant

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


Public Sub Init()
        'Pi = CDec("3,1415926535897932384626433832795") '0288419716939937510582097494459230781640628620899862803482534211706798214")
                  ' 3,1415926535897932384626433833
     'Euler = CDec("2,7182818284590452353602874713526") '6249775724709369995957496696762772407663035354759457138217852516642742746")
                  ' 2,7182818284590452353602874714


'https://oeis.org/A000796
         Pi = Constant_Parse(1, "3, 1, 4, 1, 5, 9, 2, 6, 5, 3, 5, 8, 9, 7, 9, 3, 2, 3, 8, 4, 6, 2, 6, 4, 3, 3, 8, 3, 2, 7, 9, 5, 0, 2, 8, 8, 4, 1, 9, 7, 1, 6, 9, 3, 9, 9, 3, 7, 5, 1, 0, 5, 8, 2, 0, 9, 7, 4, 9, 4, 4, 5, 9, 2, 3, 0, 7, 8, 1, 6, 4, 0, 6, 2, 8, 6, 2, 0, 8, 9, 9, 8, 6, 2, 8, 0, 3, 4, 8, 2, 5, 3, 4, 2, 1, 1, 7, 0, 6, 7, 9, 8, 2, 1, 4")
     
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

'https://oeis.org/A081801
MassElektron = Constant_Parse(1, "9, 1, 0, 9, 3, 8") * 10 ^ -31 'kg

'1,67262192 * 10 ^ (-27) kg
'https://oeis.org/A070059
MassProton = Constant_Parse(1, "1, 6, 7, 2, 6, 2, 1, 9, 2") * 10 ^ -27 'kg


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

'Teilchenzahl
'N_A
'https://oeis.org/A322578
Avogadro = Constant_Parse(1, "6, 0, 2, 2, 1, 4, 0, 7, 6") * 10 ^ 23 '1/mol

ElemCharge = Constant_Parse(1, "1, 6, 0, 2, 1, 7, 6, 6, 3, 4") * 10 ^ -19 'C (Coulomb)

'https://oeis.org/A003676
PlanckQuantum = 6.62607015 * 10 ^ (-34)

QuantumAlpha = CDec(CDec(1) / CDec(137))

MagnPermittvy = Constant_Parse(1, "1, 2, 5, 6, 6, 3, 7, 0, 6, 1, 4, 3, 5, 9, 1, 7, 2, 9, 5, 3, 8, 5, 0, 5, 7, 3, 5, 3, 3, 1, 1, 8, 0, 1, 1, 5, 3, 6, 7, 8, 8, 6, 7, 7, 5, 9, 7, 5, 0, 0, 4, 2, 3, 2, 8, 3, 8, 9, 9, 7, 7, 8, 3, 6, 9, 2, 3, 1, 2, 6, 5, 6, 2, 5, 1, 4, 4, 8, 3, 5, 9, 9, 4, 5, 1, 2, 1, 3, 9, 3, 0, 1, 3, 6, 8, 4, 6, 8, 2") * 10 ^ -6   '

ElecPermittvy = Constant_Parse(1, "8, 8, 5, 4, 1, 8, 7, 8, 1, 7, 6, 2, 0, 3, 8, 9, 8, 5, 0, 5, 3, 6, 5, 6, 3, 0, 3, 1, 7, 1, 0, 7, 5, 0, 2, 6, 0, 6, 0, 8, 3, 7, 0, 1, 6, 6, 5, 9, 9, 4, 4, 9, 8, 0, 8, 1, 0, 2, 4, 1, 7, 1, 5, 2, 4, 0, 5, 3, 9, 5, 0, 9, 5, 4, 5, 9, 9, 8, 2, 1, 1, 4, 2, 8, 5, 2, 8, 9, 1, 6, 0, 7, 1, 8, 2, 0, 0, 8, 9, 3, 2, 8, 6, 7") * 10 ^ -12   '

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
    Dim i As Long, f: f = CDec(1)
    m_Factorials(0) = f
    m_Factorials(1) = f
    For i = 2 To 27
        f = f * CDec(i)
        m_Factorials(i) = f
    Next
    f = CDbl(f)
    For i = 28 To 170
        f = f * CDbl(i)
        m_Factorials(i) = f
    Next
    m_Factorials(171) = GetINFE
End Sub

Public Function Fact(ByVal N As Long) As Variant 'As Decimal
    If N > 170 Then N = 171
    Fact = m_Factorials(N)
End Function

Public Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then        'egal ob y > 0 oder y < 0    '1. Quadrant und 4. Quadrant
        Atan2 = VBA.Math.Atn(y / x)
    ElseIf x < 0 Then
        If y > 0 Then                '2. Quadrant
            Atan2 = VBA.Math.Atn(y / x) + Pi
        ElseIf y < 0 Then            '3. Quadrant
            Atan2 = VBA.Math.Atn(y / x) - Pi
        Else                         'neg x-Achse
            Atan2 = Pi
        End If
    Else
        If y > 0 Then                'pos y-Achse
            Atan2 = 0.5 * Pi
        ElseIf y < 0 Then            'neg y-Achse
            Atan2 = -0.5 * Pi
        Else                         'Nullpunkt
            Atan2 = 0#
        End If
    End If
End Function

Public Function Square(ByVal N As Double) As Double
   Square = N * N
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

Public Function PFZ(ByVal N As Long) As String
    Dim s As String
    Dim i As Long: i = 2 'CDec(2)
    Do
        While N Mod i = 0
            N = N / i
            If s <> vbNullString Then s = s & "*" 'first time wo *
            s = s & CStr(i)
        Wend
        'If i = 2 Then i = i + CDec(1) Else i = i + 2 'CDec(2)
        If i = 2 Then i = i + 1 Else i = i + 2 'CDec(2)
        If i > Int(Sqr(N)) Then i = N '//ohne diese Zeile:Kaffeepause!
    Loop Until N = 1
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

' ^ ############################## ^ '    ggT and kgV-functions    ' ^ ############################## ^ '

' v ############################## v '    Linear interpolation    ' v ############################## v '
'x1 ' y1
'x2 ' LinIPol
'x3 ' y3
Private Function LinIPol(ByVal y1 As Double, _
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

Function IsPrime(ByVal value As Long) As Boolean
'    If Value < 200 Then
'        Select Case Value
'        Case 2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47, 53, 59, 61, 67, 71, 73, 79, 83, 89, 97, 101, _
'              103, 107, 109, 113, 127, 131, 137, 139, 149, 151, 157, 163, 167, 173, 179, 181, 191, 193, 197, 199
'            IsPrime = True:        Exit Function
'        End Select
'    End If
    If (value And 1) = 0 Then Exit Function
    Dim div As Long: div = 3
    Dim squ As Long: squ = 9
    Do While squ < value
        If value Mod div = 0 Then Exit Function
        div = div + 2
        squ = div * div
    Loop
    If squ <> value Then
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
    Dim N As Long: N = CLng(VBA.Math.Sqr(CDbl(number)))
    Dim i As Long: i = 3
    Do While (i <= N)
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
    Dim N As Long: N = LOF(FNr) / 4
    'Debug.Print n
    'ReDim Primes(0 To 9591)
    ReDim Primes(0 To N - 1)
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
Function FibonacciA(Optional ByVal N As Long = -1) As Long()
    If N <= 0 Then N = 46
    'Calculates the Fibonacci-series for a given number n
    ReDim fib(0 To N) As Long: fib(1) = 1
    Dim i As Long
    For i = 2 To N
        fib(i) = fib(i - 1) + fib(i - 2)
    Next
    FibonacciA = fib
End Function

Private Function FibonacciR(ByVal N As Long) As Long
    'Calculates the Fibonacci number to any given number (may be slow for values higher than 20)
    If N > 46 Then Exit Function
    If N > 1 Then FibonacciR = FibonacciR(N - 1) + FibonacciR(N - 2) Else FibonacciR = N
End Function
' ^ ############################## ^ '    Fibonacci functions    ' ^ ############################## ^ '

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
Public Function LogN(ByVal x As Double, _
                     Optional ByVal base As Double = 10#) As Double
                     'base must not be 1 or 0
    If base = 1 Or base = 0 Then Exit Function
    LogN = VBA.Math.Log(x) / VBA.Math.Log(base)
End Function

' ^ ############################## ^ '    Logarithm functions    ' ^ ############################## ^ '

' v ############################## v '    Rounding functions     ' v ############################## v '
Public Function Floor(ByVal a As Double) As Double
    Floor = CDbl(Int(a))
End Function

Public Function Ceiling(ByVal a As Double) As Double
    Ceiling = CDbl(Int(a))
    If a <> 0 Then If Abs(Ceiling / a) <> 1 Then Ceiling = Ceiling + 1
End Function
' ^ ############################## ^ '    Rounding functions     ' ^ ############################## ^ '

' v ############################## v ' IEEE754-INFINITY functions ' v ############################## v '
' v ############################## v '      Create functions      ' v ############################## v '
'either with error handling:
Public Function GetINFE(Optional ByVal sign As Long = 1) As Double
Try: On Error Resume Next
    GetINFE = Sgn(sign) / 0
Catch: On Error GoTo 0
End Function

' or without error handling:
Public Function GetINF(Optional ByVal sign As Long = 1) As Double
    Dim L(1 To 2) As Long
    If Sgn(sign) > 0 Then
        L(2) = &H7FF00000
    ElseIf Sgn(sign) < 0 Then
        L(2) = &HFFF00000
    End If
    Call RtlMoveMemory(GetINF, L(1), 8)
End Function

Public Sub GetNaN(ByRef value As Double)
    Dim L(1 To 2) As Long
    L(1) = 1
    L(2) = &H7FF00000
    Call RtlMoveMemory(value, L(1), 8)
End Sub

Public Sub GetINDef(ByRef value As Double)
Try: On Error Resume Next
    value = 0# / 0#
Catch: On Error GoTo 0
End Sub
' ^ ############################## ^ '    Create functions    ' ^ ############################## ^ '

' v ############################## v '     Bool functions     ' v ############################## v '
Public Function IsINDef(ByRef value As Double) As Boolean
Try: On Error Resume Next
    IsINDef = (CStr(value) = CStr(INDef))
Catch: On Error GoTo 0
End Function

Public Function IsNaN(ByRef value As Double) As Boolean
    Dim b(0 To 7) As Byte
    Dim i As Long
    
    RtlMoveMemory b(0), value, 8
    
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

Public Function IsPosINF(ByVal value As Double) As Boolean
    IsPosINF = (value = posINF)
End Function

Public Function IsNegINF(ByVal value As Double) As Boolean
    IsNegINF = (value = negINF)
End Function
' ^ ############################## ^ '     Bool functions     ' ^ ############################## ^ '

' v ############################## v '    Output functions    ' v ############################## v '
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
' ^ ############################## ^ '    Output functions    ' ^ ############################## ^ '

' v ############################## v '     Input function     ' v ############################## v '
Public Function Double_TryParse(s As String, Value_out As Double) As Boolean
Try: On Error GoTo Catch
    If Len(s) = 0 Then Exit Function
    s = Replace(s, ",", ".")
    If StrComp(s, "1.#QNAN") = 0 Then
        GetNaN Value_out
    ElseIf StrComp(s, "1.#INF") = 0 Then
        Value_out = GetINF
    ElseIf StrComp(s, "-1.#INF") = 0 Then
        Value_out = GetINF(-1)
    ElseIf StrComp(s, "-1.#IND") = 0 Then
        GetINDef Value_out
    Else
        Value_out = Val(s)
    End If
    Double_TryParse = True
Catch:
End Function
' ^ ############################## ^ '       Input function       ' ^ ############################## ^ '
' ^ ############################## ^ ' IEEE754-INFINITY functions ' ^ ############################## ^ '

' v ############################## v '     solving quadratic & cubic formula     ' v ############################## v '
Public Function Quadratic(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByRef x1_out As Double, ByRef x2_out As Double) As Boolean
Try: On Error GoTo Catch
    
    'maybe the midnight-formula:
    'Dim w As Double: w = VBA.Sqr(b ^ 2 - 4 * a * c)
    'x1_out = (-b + w) / (2 * a)
    'x2_out = (-b - w) / (2 * a)
    
    'or maybe the pq-formula
    Dim p    As Double:   p = b / a
    Dim q    As Double:   q = c / a
    Dim p_2  As Double: p_2 = p / 2
    Dim p2_4 As Double: p2_4 = p_2 * p_2
    Dim W As Double: W = VBA.Sqr(p2_4 - q)
    
    x1_out = -p_2 + W
    x2_out = -p_2 - W
    
    Quadratic = True
    Exit Function
Catch:
End Function

Public Function Quadratic_ToStr(ByVal a As Double, ByVal b As Double, ByVal c As Double) As String
    Quadratic_ToStr = a & "x²" & GetOp(b) & "x" & GetOp(c) & " = 0"
End Function
Public Function Cubic_ToStr(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As String
    Cubic_ToStr = a & "x³" & GetOp(b) & "x²" & GetOp(c) & "x" & GetOp(d) & " = 0"
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
Public Function Cubic(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, ByRef x1_out As Double, ByRef x2_out As Double, ByRef i2_out As Double, ByRef x3_out As Double, ByRef i3_out As Double) As Boolean
Try: On Error GoTo Catch
    'Scipione del Ferro (1465-1526), Nicolo Tartaglia (1500-1557), Gerolamo Cardano (1501-1576)
    'Rafael Bombelli (1526-1572)
    If a = 0 Then
        Cubic = Quadratic(b, c, d, x1_out, x2_out)
        Exit Function
    End If
    Dim a2 As Double: a2 = a * a
    Dim a3 As Double: a3 = a * a2
    Dim b2 As Double: b2 = b * b
    Dim b3 As Double: b3 = b * b2
    'Dim bc As Double: bc = b * c
    Dim b3_27a3 As Double: b3_27a3 = b3 / (27 * a3)
    'Dim bc_6a2  As Double:  bc_6a2 = bc / (6 * a2)
    Dim cb_3a2  As Double:  cb_3a2 = c * b / (3 * a2)
    Dim b2_3a2  As Double:  b2_3a2 = b2 / (3 * a2)
    'Dim b2_9a2  As Double:  b2_9a2 = b2 / (9 * a2)
    Dim d_a  As Double: d_a = d / a
    'Dim d_2a As Double: d_2a = d / (2 * a)
    Dim c_a  As Double: c_a = c / a
    'Dim c_3a As Double: c_3a = c / (3 * a)
    Dim b_3a As Double: b_3a = b / (3 * a)
    
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
    W = (q_2 ^ 2 + p_3 ^ 3) ^ (1 / 2)

    x1_out = -b_3a + (-q_2 + W) ^ (1 / 3) _
                   + (-q_2 - W) ^ (1 / 3)
    
    Cubic = True
    Exit Function
Catch:
End Function

Public Function SqrH(ByVal N As Double) As Double
    'SquareRoot due to Heron algorithm
    Dim s As Long:   s = Sgn(N): N = Abs(N)
    Dim r As Double: r = N
    Dim i As Long
    SqrH = 1
    Do
        i = i + 1
        SqrH = (SqrH + r) / 2
        r = N / SqrH
        If (SqrH - r) < 0.0000000000001 Then Exit Do
        If i = 20 Then Exit Do
    Loop
    'Debug.Print i
End Function

Public Function CubRt(ByVal v As Double, ByRef i_out As Double) As Double
    'CubRt(1) = 1, -1/2+- SquRt(3) / 2 i
    If v = 0 Then Exit Function
    If v > 0 Then
        'Root = VBA.Sqr(v)
        Exit Function
    End If
    v = Abs(v)
    CubRt = VBA.Sqr(v)
    
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
        MsgBox nrows & ": overflow!" & vbCrLf & "The triangle will be computed with the maximum of only 1030 rows."
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
        .phi = Atan2(c.Im, c.Re)
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

Public Function ComplexP_Add(p1 As ComplexP, p2 As ComplexP) As ComplexP
    'Dim z1 As Complex: z1 = ComplexP_ToComplex(p1)
    'Dim z2 As Complex: z2 = ComplexP_ToComplex(p2)
    'Dim z3 As Complex: z2 = Complex_Add(z1, z2)
    'ComplexP_Add = Complex_ToComplexP(z3)
    ComplexP_Add = Complex_ToComplexP(Complex_Add(ComplexP_ToComplex(p1), ComplexP_ToComplex(p2)))
End Function

Public Function ComplexP_Mul(p1 As ComplexP, p2 As ComplexP) As ComplexP
    With ComplexP_Mul
        .r = p1.r * p2.r
        .phi = p1.phi + p2.phi
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

Public Function ComplexP_NthRoot(p As ComplexP, ByVal N As Long) As ComplexP()
    'berechnet die n-te Wurzel einer komplexem zahl p
    'computes the nth-root of the complex number p
    Dim ccp() As ComplexP
    If N = 0 Then Exit Function 'q € N w/o 0
    If N = 1 Then
        ReDim ccp(0 To 0): ccp(0) = p
        ComplexP_NthRoot = ccp()
        Exit Function
    End If
    Dim N_1  As Double:  N_1 = 1 / N
    Dim r    As Double:    r = p.r ^ N_1
    Dim phi  As Double:  phi = p.phi * N_1
    Dim ppi2 As Double: ppi2 = Pi2
    Dim dp   As Double:   dp = ppi2 / N
    ReDim ccp(0 To N - 1)
    ccp(0) = ComplexP(r, phi)
    Dim i As Long
    For i = 1 To N - 1
        phi = phi + dp
        ccp(i) = ComplexP(r, phi)
    Next
    ComplexP_NthRoot = ccp
End Function

Public Function ModF(ByVal value As Double, ByVal div As Double) As Double
   ModF = value - (Int(value / div) * div)
End Function

Public Function ModDbl(v As Double, d As Double) As Double
    'berechnet das Überbleibsel bei der division
    Dim s As Long: s = Sgn(v)
    v = Abs(v)
    Dim i As Long: i = Int(v / d)
    ModDbl = s * (v - i * d)
End Function

Public Function ComplexP_ToComplex(p As ComplexP) As Complex
    With ComplexP_ToComplex
        .Re = p.r * VBA.Math.Cos(p.phi)
        .Im = p.r * VBA.Math.Sin(p.phi)
    End With
End Function

' ^ ############################## ^ '    Complex numbers    ' ^ ############################## ^ '

Public Function CalcPi()
    Dim sqr3: sqr3 = CDec("1,7320508075688772935274463415") '058723669428052538103806280558069794519330169088000370811461867572485756")
    Dim sum: sum = CDec(0)
    Dim N As Long
    For N = 1 To 40
        sum = sum + -Fact(2 * N - 2) / (2 ^ (4 * N - 2) * Fact(N - 1) ^ 2 * (2 * N - 3) * (2 * N + 1))
    Next
    Dim Pi
    Pi = 3 * sqr3 / 4 + 24 * sum
    CalcPi = Pi
End Function

