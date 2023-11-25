Attribute VB_Name = "MMath"
Option Explicit
'Wertebereich Currency (Int64)
'    Dim bigDec1: bigDec1 = CDec("9223372036854775807")   ' No overflow.
'    Dim bigDec2: bigDec2 = CDec("9223372036854775808")   ' No Overflow.
'    Dim bigDec3: bigDec3 = CDec("9223372036854775809")   ' No overflow.

'Wertebereich Decimal
'    Dim bigDec1: bigDec1 = CDec("79228162514264337593543950335")
'    Dim bigDec2: bigDec2 = CDec("-79228162514264337593543950335")
'    Dim bigDec3: bigDec3 = CDec("7,9228162514264337593543950335")
'    Dim bigDec4: bigDec4 = CDec("-7,9228162514264337593543950335")
'    Dim bigDec5: bigDec5 = CDec("0,0000000000000000000000000001")
'    Dim bigDec6: bigDec6 = CDec("-0,0000000000000000000000000001")


'Mathematische Konstanten
'https://de.wikipedia.org/wiki/Mathematische_Konstante
'https://de.wikipedia.org/wiki/Liste_besonderer_Zahlen
Public Pi          ' As Variant As Decimal
Public Pihalf      ' As Variant As Decimal
Public Pi2         ' As Variant As Decimal
Public Euler       ' As Variant As Decimal
Public SquareRoot2 ' As Variant As Decimal
Public SquareRoot3 ' As Variant As Decimal
Public GoldenRatio ' As Variant As Decimal

'Physikalische Konstanten
Public SpeedOfLight   'Lichtgeschwindigkeit im Vakuum      c   = 299792458 m/s
Public ElemCharge     'Elementarladung (des Protons)       e   = 1,602176634 * 10^-19 C (Coulomb)
Public MassElektron   'Ruhemasse des Elektrons             m_e = 9,109*10^-31 kg
Public MassProton     'Ruhemasse des Protons               m_p = 1,6726215813 · 10^-27 kg
Public PlanckQuantum  'Plancksches Wirkungsquantum         h   = 6,62607015 * 10^(-37) m² * kg / s
Public Avogadro       'Avogadro-Konstante                  N_A = 6,022 * 10^23
Public Gravitation    'Newtonsche Gravitationskonstante    G   = 6,6743 * 10^-11 m³ / (kg * s²)
Public BoltzmannConst 'Boltzmann-Konstante                 k_B = 1,38064852 × 10-23 m2 kg s-2 K-1
Public MagnPermittvy  'magnetische Feldkonstante           mue_0 = µ0 ˜ 1.2566370621219 * 10 ^(-6) N/A²
Public ElecPermittvy  'elektrische Feldkonstante           eps_0 = 8.8541878128(13)e-12 (A s)/(V m)
Public QuantumAlpha   'FineStructureConstant

Private m_Factorials() 'As Variant 'As Decimal
Public Primes()  As Long 'contains all primes up to 100000    'As Variant 'As Decimal

Public PrimesX() As Long 'a distinct selection of primes

Public Fibonacci() As Long

'Public Dedekind(1 To 9)

'NTSYSAPI SIZE_T RtlCompareMemory(
'  [in] const VOID *Source1,
'  [in] const VOID *Source2,
'  [in] SIZE_T     Length
');
'Private Declare Function RtlCompareMemory Lib "ntdll" (pSrc1 As Long, pSrc2 As Long, ByVal Length As Long) As Long


'Boltzmann-Konstante kB
'magnetische und elektrische Feldkonstante   µ0, e0

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
End Sub

Private Sub InitFibonacci()
    Fibonacci = FibonacciA
End Sub

Private Sub InitFactorials()
    ReDim m_Factorials(0 To 171)
    Dim i As Long, f: f = CDec(1)
    m_Factorials(0) = CDec(0)
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

Private Function GetINFE(Optional ByVal sign As Long = 1) As Double
    On Error Resume Next
    GetINFE = Sgn(sign) / 0
    On Error GoTo 0
End Function

Public Function Fact(ByVal n As Long) As Variant 'As Decimal
    If n > 170 Then n = 171
    Fact = m_Factorials(n)
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
    'ggT = größter gemeinsamer Teiler
   Do While x <> y
      If x > y Then
         x = x - y
      Else
         y = y - x
      End If
   Loop 'Wend
   GreatestCommonDivisor = x
End Function

Public Function kgV(ByVal x As Long, ByVal y As Long) As Long
    kgV = (x * y) / ggT(x, y)
End Function

Public Function LeastCommonMultiple(ByVal x As Long, ByVal y As Long) As Long
    'kgV = kleinstes gemeinsames Vielfaches
    
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
'
'function kgV(a,b:integer):integer;
'Begin
'  result:=a*b div ggT(a,b);
'end;
'
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
    Debug.Print n
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

