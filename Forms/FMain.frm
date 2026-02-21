VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14655
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   14655
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   360
      Width           =   14535
   End
   Begin VB.CommandButton BtnTestPrimes2 
      Caption         =   "Test Primes 2"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton BtnTestPrimes 
      Caption         =   "Testing Primes"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Me.Caption = "Math: " & App.FileDescription & " v" & App.Major & "." & App.Minor & "." & App.Revision
    MMath.Init
    Tests
'value range Currency (Int64) / 10000
    Dim Cur1 As Currency: Cur1 = 922337203685477.5807@ 'CCur("922337203685477,5807")   ' No overflow.
    'Dim Cur2 As Currency: Cur2 = 922337203685477.5808@ 'CCur("922337203685477,5808")   ' Overflow.
    'Dim Cur3 As Currency: Cur3 = 'CCur("922337203685477,5809")   ' overflow.
    
''value range Currency (signed Int64)
''Currency (skalierte Ganzzahl)
''8 Bytes -922.337.203.685.477,5808 bis 922.337.203.685.477,5807
'
'    Dim c As Currency
'
'    c = -1234567890123.46    'mit Double nur max 2 Stellen nach dem Komma in der IDE
'    MsgBox c
'    c = -1234567890123.4567@ 'mit Currency@ gehen alle 4 Nachkommastellen
'    MsgBox c
'
'    Dim MinCur As Currency: MinCur = CCur("-922337203685477,5808")   ' No overflow.
'    MsgBox MinCur
'    MinCur = -922337203685477.5807@
'    MsgBox MinCur
'
'    Dim MaxCur As Currency: MaxCur = CCur("922337203685477,5807")   ' No Overflow.
'    MsgBox MaxCur
'    MaxCur = 922337203685477.5807@
'    MsgBox MaxCur
End Sub

Private Sub Form_Resize()
    Dim l As Single
    Dim t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then
        Text1.Move l, t, W, H
    End If
End Sub

Private Sub BtnTestPrimes_Click()
    Dim n As Long: n = 100000
    Dim t As Single: t = Timer
    Dim i As Long, b As Boolean
    Dim p As Long
    Dim c As Long
    For i = 0 To n
        b = IsPrime(i)
        If b Then
            p = i
            c = c + 1
        End If
    Next
    t = Timer - t
    MsgBox "Testing numbers from 0 to " & Format(n, "#,##0") & " whether it's a prime." & vbCrLf & _
           "Found " & c & " primes. This took about " & t & " seconds." & vbCrLf & "The last prime was " & Format(p, "#,##0")
End Sub

Private Sub BtnTestPrimes2_Click()
    Dim dt As Single
    
    Dim p As Long
    Dim i As Long
    Dim np As Long: np = 9591
    Dim ni As Long: ni = 1000000
    Dim b As Boolean
    
    dt = Timer
    For i = 0 To np
        p = MMath.Primes(i)
        b = IsPrime(p)
    Next
    For i = 0 To ni
        b = IsPrime(i)
    Next
    
    dt = Timer - dt
    
    Label2.Caption = dt
    
    dt = Timer
    For i = 0 To np
        p = MMath.Primes(i)
        b = IsPrimeX(p)
    Next
    For i = 0 To ni
        b = IsPrimeX(i)
    Next
    dt = Timer - dt
    
    Label3.Caption = dt
    
End Sub

Sub Tests()
    TestMinMaxValues
    TestConstants
    TestTempConvs
    TestGgT_KgV_PFZ_FraC
    TestFactorials
    TestPrimes
    TestQuadraticCubic
    TestPascalTriangle
    TestComplex
    TestMinMax
    TestFibonacci
    TestFloorCeiling
    TestRounding
    TestRandom
    TestTrigono
End Sub

Sub TestMinMaxValues()
    Dim n As Long: n = 21
    AddItem "TestNumericMinMaxValues:"
    AddItem "========================"
    AddItem "Min. Byte     = " & MString.PadCentered(MMath.MinByte, n)
    AddItem "Max. Byte     = " & MString.PadCentered(MMath.MaxByte, n)
    AddItem "Min. Integer  = " & MString.PadCentered(MMath.MinInteger, n)
    AddItem "Max. Integer  = " & MString.PadCentered(MMath.MaxInteger, n)
    AddItem "Min. Long     = " & MString.PadCentered(MMath.MinLong, n)
    AddItem "Max. Long     = " & MString.PadCentered(MMath.MaxLong, n)
    AddItem "Min. Currency = " & MString.PadCentered(MMath.MinCurrency, n)
    AddItem "Max. Currency = " & MString.PadCentered(MMath.MaxCurrency, n)
    AddItem "Min. Decimal  = " & MString.PadCentered(MMath.MinDecimal, n)
    AddItem "Max. Decimal  = " & MString.PadCentered(MMath.MaxDecimal, n)
    AddItem "Min.Dec.Fract+= " & MString.PadCentered(MMath.MinDecimalFractionPositive, n)
    AddItem "Max.Dec.Fract-= " & MString.PadCentered(MMath.MinDecimalFractionNegative, n)
    AddItem ""
    
End Sub

Sub TestConstants()
    'With aLB
    AddItem "TestConstants:"
    AddItem "=============="
    
    AddItem "Pi          =" & MMath.Pi
    AddItem "Pi calced   =" & MMath.CalcPi
    AddItem "2*Pi        =" & MMath.Pi2
    AddItem "Pi/2        =" & MMath.Pihalf
    AddItem "Euler       =" & MMath.Euler       ' As Variant As Decimal
    AddItem "SquareRoot2 =" & MMath.SquareRoot2 ' As Variant As Decimal
    AddItem "SquareRoot3 =" & MMath.SquareRoot3 ' As Variant As Decimal
    AddItem "GoldenRatio =" & MMath.GoldenRatio ' As Variant As Decimal

'Physikalische Konstanten
    AddItem "SpeedOfLight     =" & MMath.SpeedOfLight & " m/s"  'Lichtgeschwindigkeit im Vakuum      c
    AddItem "MassOfElektron   =" & MMath.MassElektron   'Ruhemasse des Elektrons             me
    AddItem "MassOfProton     =" & MMath.MassProton     'Ruhemasse des Protons               mp
    AddItem "Gravitation      =" & MMath.Gravitation    'Newtonsche Gravitationskonstante    G
    AddItem "Avogadro         =" & MMath.Avogadro       'Avogadro-Konstante                  NA
    AddItem "Boltzmann        =" & MMath.Boltzmann      'Boltzmann-Konstante                 k_B
    AddItem "StefanBoltzmann  =" & MMath.StefanBoltzmann 'Stefan Boltzmann-Konstante         sigma
    
    AddItem "ProtonCharge     =" & MMath.ElemCharge     'Elementarladung (des Protons)       e
    AddItem "PlanckQuantum    =" & MMath.PlanckQuantum  'Plancksches Wirkungsquantum         h
    AddItem "QuantumAlpha     =" & MMath.QuantumAlpha
    AddItem "ElectricPermittivity = " & MMath.ElecPermittvy ' Dielectrizitäts-Konstante  eps_0
    AddItem ""
    
End Sub

Sub TestTempConvs()
    
    AddItem "Test Temperature Conversions:"
    AddItem "============================="
    
    Dim c As Double, F As Double, k As Double
    
    c = 40:   F = TempCelsius_ToFahrenheit(c):     AddItem c & " °C = " & F & " °F"        ''  40 °C ==  104 °F
    c = -40:  F = TempCelsius_ToFahrenheit(c):     AddItem c & " °C = " & F & " °F"        '' -40 °C ==  -40 °F
    
    c = 40:   k = TempCelsius_ToKelvin(c):         AddItem c & " °C = " & k & " °K"        ''  40 °C ==  313,15 °K
    c = -40:  k = TempCelsius_ToKelvin(c):         AddItem c & " °C = " & k & " °K"        '' -40 °C ==  233,15 °K
    
    
    F = 104:  c = TempFahrenheit_ToCelsius(F):     AddItem F & " °F = " & c & " °C"        '' 104 °F ==   40 °C
    F = -104: c = TempFahrenheit_ToCelsius(F):     AddItem F & " °F = " & c & " °C"        ''-104 °F ==  -75,556 °C
    
    F = 104:  k = TempFahrenheit_ToKelvin(F):      AddItem F & " °F = " & k & " °K"        '' 104 °F ==  313,15 °K
    F = -104: k = TempFahrenheit_ToKelvin(F):      AddItem F & " °F = " & k & " °K"        ''-104 °F ==  197,59 °K
    
    
    k = 104:  c = TempKelvin_ToCelsius(k):         AddItem k & " °K = " & c & " °C"        '' 104 °K == -169,15 °C
    k = 40:   c = TempKelvin_ToCelsius(k):         AddItem k & " °K = " & c & " °C"        ''  40 °K == -233,15 °C
    
    k = 104:  F = TempKelvin_ToFahrenheit(k):      AddItem k & " °K = " & F & " °F"        '' 104 °K == -272,47 °F
    k = 40:   F = TempKelvin_ToFahrenheit(k):      AddItem k & " °K = " & F & " °F"        ''  40 °K == -387,67 °F
    
    AddItem ""
    
End Sub

Sub TestGgT_KgV_PFZ_FraC()
    
    'MsgBox PFZ(6442450938@)
    
    AddItem "TestGgT_KgV_PFZ_FraC"
    AddItem "===================="
    Dim n As Long: n = 2147483644
    AddItem "The prime factors of " & Format(n, "#,##0") & " are " & PFZ(n)
    
    Dim n1 As Long: n1 = 1234
    Dim n2 As Long: n2 = 56
    AddItem "Primefactors(" & n1 & ") = " & PFZ(n1)
    AddItem "GreatestCommonDivisor(" & n1 & ", " & n2 & ") = " & MMath.GreatestCommonDivisor(n1, n2)
    AddItem "LeastCommonMultiple(" & n1 & ", " & n2 & ") = " & MMath.LeastCommonMultiple(n1, n2)
    n2 = 3456
    Dim nn As Long: nn = n1
    Dim za As Long: za = n2
    MMath.CancelFraction nn, za
    AddItem "CancelFraction(" & n1 & ", " & n2 & ") = " & nn & " / " & za
    AddItem ""
    
End Sub

Sub TestFactorials()
    AddItem "TestFactorials"
    AddItem "=============="
    
    Dim i As Long, s As String
    For i = 0 To 5
        AddItem CStr(i) & "! = " & MMath.Fact(i)
    Next
    For i = 22 To 27
        AddItem CStr(i) & "! = " & MMath.Fact(i)
    Next
    For i = 28 To 30
        AddItem CStr(i) & "! = " & MMath.Fact(i)
    Next
    For i = 168 To 171
        AddItem CStr(i) & "! = " & MMath.Fact(i)
    Next
    AddItem ""
End Sub

Sub TestPrimes()
    AddItem "TestPrimes"
    AddItem "=========="
    'Dim n As Long: n = 5
    'Dim n As Long: n = 96211
    'Dim n As Long: n = 99991
    Dim n As Long: n = 99991
    AddItem "IsPrimeA(" & n & ") = " & MMath.IsPrime(n)
    AddItem ""
End Sub

Sub TestQuadraticCubic()
    AddItem "TestQuadraticCubic"
    AddItem "=================="
    Dim A As Double, b As Double, c As Double, d As Double
    Dim x1 As Double
    Dim x2 As Double, i2 As Double
    Dim x3 As Double, i3 As Double
    'a = 2 ' 1
    'b = 4 '0 '-1
    'c = -2 '6 '-4
    'd = -4 '21 '4
    
    A = 1: b = 8: c = -20
    x1 = 0: x2 = 0: x3 = 0
    
    AddItem Quadratic_ToStr(A, b, c)
    If Quadratic(A, b, c, x1, x2) Then
        'x1 = 2; x2 = -10
        AddItem "x1 = " & x1 & "; x2 = " & x2
    End If
    
    A = 2: b = -6: c = -4: d = -4
    AddItem Quadratic_ToStr(A, b, c)
    If Quadratic(A, b, c, x1, x2) Then
        'x1 = -1; x2 = -2
        AddItem "x1 = " & x1 & "; x2 = " & x2
    End If
    
    'a = 0.25: b = 0.75: c = -1.5: d = -2
    A = 2: b = 6: c = -4: d = -24
    x1 = 0: x2 = 0: i2 = 0: x3 = 0: i3 = 0
    
    AddItem Cubic_ToStr(A, b, c, d)
    
    If MMath.Cubic(A, b, c, d, x1, x2, i2, x3, i3) Then
        
        AddItem "x1 = " & x1
        
    End If
    AddItem ""
End Sub

Sub TestPascalTriangle()
    AddItem "TestPascalTriangle"
    AddItem "=================="
    Dim pt(): pt = MMath.PascalTriangle(14) 'max 1030 rows
    AddItem MMath.PascalTriangle_ToStr(pt)
    AddItem ""
End Sub

Sub TestComplex()
    
    AddItem "TestComplexNumbers"
    AddItem "=================="
    Dim z1 As Complex
    z1 = MMath.Real_ToComplex(8)
    AddItem "v=" & 8
    AddItem "z1(v)=" & MMath.Complex_ToStr(z1)
    
    z1 = MMath.Complex(1, 0.5)
    AddItem "In cartes. coords.:"
    AddItem "z1 = " & Complex_ToStr(z1)
    Dim z2 As Complex: z2 = MMath.Complex(2, 3)
    AddItem "z2 = " & Complex_ToStr(z2)
    Dim z As Complex
    z = MMath.Complex_Add(z1, z2)
    AddItem "z = z1 + z2; z = " & MMath.Complex_ToStr(z1) & " + (" & MMath.Complex_ToStr(z2) & ") = " & MMath.Complex_ToStr(z)
    
    z = MMath.Complex_Mul(z1, z2)
    AddItem "z = z1 * z2; z = " & MMath.Complex_ToStr(z1) & " * (" & MMath.Complex_ToStr(z2) & ") = " & MMath.Complex_ToStr(z)
    
    z1 = MMath.Complex(-3, 4)
    z2 = MMath.Complex(1, 3)
    z = MMath.Complex_Div(z1, z2)
    AddItem "z = z1 / z2; z = " & MMath.Complex_ToStr(z1) & " / (" & MMath.Complex_ToStr(z2) & ") = " & MMath.Complex_ToStr(z)
    
    
    z1 = MMath.Complex(Sqr(2) / 2, -Sqr(2) / 2)
    AddItem "z1 = " & MMath.Complex_ToStr(z1)
    AddItem "In polar coordinates: "
    Dim zp1 As ComplexP: zp1 = MMath.Complex_ToComplexP(z1)
    AddItem "zp1 = " & MMath.ComplexP_ToStr(zp1)
    AddItem "Or in euler form: "
    AddItem "zp1 = " & MMath.ComplexP_ToStrE(zp1)
    
    Dim ex As Long: ex = 2020
    Dim zp As ComplexP: zp = MMath.ComplexP_Powi(zp1, ex)
    AddItem "zp = zp1 ^ 2020; zp = (" & MMath.ComplexP_ToStrE(zp1) & ") ^ (" & ex & ") = " & MMath.ComplexP_ToStrE(zp)
    
    Dim r   As Double: r = 2
    Dim phi As Double ': phi=0
    Dim p   As Long: p = 4
    Dim q   As Long: q = 3
    Dim i As Long
    zp = MMath.ComplexP(r, phi)
    
    Dim zzp() As ComplexP
    zzp = MMath.ComplexP_Pow(zp, p, q)
    
    For i = 0 To q - 1
        AddItem ComplexP_ToStrE(zzp(i))
    Next
    
    phi = Pihalf
    zp = MMath.ComplexP(r, phi)
    
    Dim n As Long: n = 5
    zzp = ComplexP_NthRoot(zp, n)
    
    For i = 0 To n - 1
        AddItem ComplexP_ToStrE(zzp(i))
    Next
    
    AddItem ""
End Sub

Private Sub TestMinMax()
    Dim m 'As Long
    m = MinArr(15, 12, 22, 45, 100, 72, 11, 83, 46, 25, 35)
    AddItem "The Minimum out of (15, 12, 22, 45, 100, 72, 11, 83, 46, 25, 35) is " & m '11
    m = MaxArr(15, 12, 22, 45, 100, 72, 11, 83, 46, 25, 35)
    AddItem "The Maximum out of (15, 12, 22, 45, 100, 72, 11, 83, 46, 25, 35) is " & m '100
    
    AddItem ""
End Sub

Private Sub TestFibonacci()
    AddItem "TestFibonacci"
    AddItem "============="
    AddItem "Fibonacci(15) = " & MMath.Fibonacci(15)
    
    AddItem ""
End Sub

Private Sub TestFloorCeiling()
    AddItem "TestFloorCeiling"
    AddItem "================"
    'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/int-fix-functions
    Dim MyNumber As Double
    MyNumber = Int(99.8)     ' Returns 99.   'Floor
    MyNumber = Fix(99.2)     ' Returns 99.   'Floor
    
    MyNumber = Int(-99.8)    ' Returns -100. 'Floor
    MyNumber = Fix(-99.8)    ' Returns -99.  'Ceiling
    
    MyNumber = Int(-99.2)    ' Returns -100. 'Floor
    MyNumber = Fix(-99.2)    ' Returns -99.  'Ceiling
    
    MyNumber = 99.8: AddItem (MyNumber & "; Floor=" & MMath.Floor(MyNumber) & "; Ceiling=  " & MMath.Ceiling(MyNumber))  ' 99.8; Floor=100; Ceiling=  99
    MyNumber = 99.2: AddItem (MyNumber & "; Floor=" & MMath.Floor(MyNumber) & "; Ceiling=  " & MMath.Ceiling(MyNumber))  ' 99.2; Floor=100; Ceiling=  99
    
    MyNumber = -99.8: AddItem (MyNumber & "; Floor=" & MMath.Floor(MyNumber) & "; Ceiling=" & MMath.Ceiling(MyNumber))   '-99.8; Floor=-99; Ceiling=-100
    MyNumber = -99.2: AddItem (MyNumber & "; Floor=" & MMath.Floor(MyNumber) & "; Ceiling=" & MMath.Ceiling(MyNumber))   '-99.2; Floor=-99; Ceiling=-100
    
    AddItem ""
End Sub

Private Sub TestRounding()
    
    AddItem "TestRounding"
    AddItem "============"
    
    Dim d As Double
    Dim r As Double
    Dim n As Byte
    Dim s As String
    
    n = 2
    d = 123.456
    r = Round(d, n):     s = s & "Round(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundUp(d, n):   s = s & "RoundUp(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundDown(d, n): s = s & "RoundDown(" & d & ", " & n & ") = " & r & vbCrLf
    AddItem s
    s = ""
    d = -d
    r = Round(d, n):     s = s & "Round(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundUp(d, n):   s = s & "RoundUp(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundDown(d, n): s = s & "RoundDown(" & d & ", " & n & ") = " & r & vbCrLf
    AddItem s
    s = ""
    
    d = 12.3456
    r = Round(d, n):     s = s & "Round(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundUp(d, n):   s = s & "RoundUp(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundDown(d, n): s = s & "RoundDown(" & d & ", " & n & ") = " & r & vbCrLf
    AddItem s
    s = ""
    d = -d
    r = Round(d, n):     s = s & "Round(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundUp(d, n):   s = s & "RoundUp(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundDown(d, n): s = s & "RoundDown(" & d & ", " & n & ") = " & r & vbCrLf
    AddItem s
    s = ""
    
    
    n = 0
    d = 123.456
    r = Round(d, n):     s = s & "Round(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundUp(d, n):   s = s & "RoundUp(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundDown(d, n): s = s & "RoundDown(" & d & ", " & n & ") = " & r & vbCrLf
    AddItem s
    s = ""
    d = -d
    r = Round(d, n):     s = s & "Round(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundUp(d, n):   s = s & "RoundUp(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundDown(d, n): s = s & "RoundDown(" & d & ", " & n & ") = " & r & vbCrLf
    AddItem s
    s = ""
    
    d = 12.3456
    r = Round(d, n):     s = s & "Round(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundUp(d, n):   s = s & "RoundUp(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundDown(d, n): s = s & "RoundDown(" & d & ", " & n & ") = " & r & vbCrLf
    AddItem s
    s = ""
    d = -d
    r = Round(d, n):     s = s & "Round(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundUp(d, n):   s = s & "RoundUp(" & d & ", " & n & ") = " & r & vbCrLf
    r = RoundDown(d, n): s = s & "RoundDown(" & d & ", " & n & ") = " & r & vbCrLf
    AddItem s
    s = ""
    
    
End Sub

Private Sub TestRandom()
    
    AddItem "TestRandom"
    AddItem "=========="
    
    Dim s As String
    Dim b As Byte
    Dim i As Integer
    Dim l As Long
    Dim d
    
    i = MPtr.RndInt8:     s = "RndInt8  (       -128 ..        128): " & i:    AddItem s
    b = MPtr.RndUInt8:    s = "RndUInt8 (          0 ..        255): " & b:    AddItem s
    i = MPtr.RndInt16:    s = "RndInt16 (     -32768 ..      32767): " & i:    AddItem s
    l = MPtr.RndUInt16:   s = "RndUInt16(          0 ..      65536): " & l:    AddItem s
    l = MPtr.RndInt32:    s = "RndInt32 (-2147483648 .. 2147483647): " & l:    AddItem s
    d = MPtr.RndUInt32:   s = "RndUInt32(          0 .. 4294967296): " & d:    AddItem s
    d = MPtr.RndInt64:    s = "RndInt64 (          0 .. 4294967296): " & d:    AddItem s
    AddItem ""
    
End Sub

Private Sub TestTrigono()
    
    AddItem "Test trigonometric functions"
    AddItem "============================"
    
    Dim ad As Double: ad = 60
    Dim ar As Double: ar = MMath.DegToRad(ad)
    Dim fmt As String: fmt = "0.0####"
    Dim s As String
    Dim inn As Double, res As Double
    
    inn = ar:    res = MMath.Sin(inn): s = "Sin(" & ad & "°) = " & Format(res, fmt):     AddItem s
    inn = ar:    res = MMath.Cos(inn): s = "Cos(" & ad & "°) = " & Format(res, fmt):     AddItem s
    inn = ar:    res = MMath.Tan(inn): s = "Tan(" & ad & "°) = " & Format(res, fmt):     AddItem s
    AddItem ""
    
    inn = ar:    res = MMath.Csc(inn): s = "Csc(" & ad & "°) = " & Format(res, fmt):     AddItem s
    inn = ar:    res = MMath.Sec(inn): s = "Sec(" & ad & "°) = " & Format(res, fmt):     AddItem s
    inn = ar:    res = MMath.Cot(inn): s = "Cot(" & ad & "°) = " & Format(res, fmt):     AddItem s
    AddItem ""
    AddItem ""
    
    
    inn = Sin(ar): res = MMath.ASin(inn): s = "ASin(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt) & "°":   AddItem s
    inn = Cos(ar): res = MMath.ACos(inn): s = "ACos(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt) & "°":   AddItem s
    inn = Tan(ar): res = MMath.ATan(inn): s = "ATan(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt) & "°":   AddItem s
    AddItem ""
    'ar = ar - 0.5
    inn = Csc(ar): res = MMath.ACsc(inn): s = "ACsc(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt) & "°":   AddItem s
    inn = Sec(ar): res = MMath.ASec(inn): s = "ASec(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt) & "°":   AddItem s
    inn = Cot(ar): res = MMath.ACot(inn): s = "ACot(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt) & "°":   AddItem s
    AddItem ""
    AddItem ""
    
    
    inn = ar: res = MMath.Sinh(inn): s = "SinH(" & ad & ") = " & Format(res, fmt):       AddItem s
    inn = ar: res = MMath.Cosh(inn): s = "CosH(" & ad & ") = " & Format(res, fmt):       AddItem s
    inn = ar: res = MMath.Tanh(inn): s = "TanH(" & ad & ") = " & Format(res, fmt):       AddItem s
    AddItem ""
    
    inn = ar: res = MMath.CscH(inn): s = "CscH(" & ad & ") = " & Format(res, fmt):       AddItem s
    inn = ar: res = MMath.SecH(inn): s = "SecH(" & ad & ") = " & Format(res, fmt):       AddItem s
    inn = ar: res = MMath.CotH(inn): s = "CotH(" & ad & ") = " & Format(res, fmt):       AddItem s
    AddItem ""
    AddItem ""
    
    
    inn = Sinh(ar): res = MMath.ArSinH(inn): s = "ArSinH(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt):    AddItem s
    inn = Cosh(ar): res = MMath.ArCosH(inn): s = "ArCosH(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt):    AddItem s
    inn = Tanh(ar): res = MMath.ArTanH(inn): s = "ArTanH(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt):    AddItem s
    AddItem ""
    
    inn = CscH(ar): res = MMath.ArCscH(inn): s = "ArCscH(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt):    AddItem s
    inn = SecH(ar): res = MMath.ArSecH(inn): s = "ArSecH(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt):    AddItem s
    inn = CotH(ar): res = MMath.ArCotH(inn): s = "ArCotH(" & Format(inn, fmt) & ") = " & Format(MMath.RadToDeg(res), fmt):    AddItem s
    AddItem ""
    
End Sub

Sub AddItem(s As String)
    Text1.text = Text1.text & s & vbCrLf
End Sub

