VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6255
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
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
      Height          =   5655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   8
      Top             =   480
      Width           =   14535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   4680
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Constants"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim n As Long: n = 100000
    Dim t As Single: t = Timer
    Dim i As Long, b As Boolean
    Dim c As Long
    For i = 0 To n
        b = IsPrime(i)
        If b Then c = c + 1
    Next
    t = Timer - t
    MsgBox "time: " & t & "    " & c
End Sub

Private Sub Command2_Click()
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

Private Sub Command3_Click()
    'MsgBox MMath.Dedekind(8)
    
    'MsgBox PFZ(6442450938@)
    Dim n As Long: n = 2147483644
    MsgBox "n = " & PFZ(n) & " = " & n
    
End Sub

Private Sub Command4_Click()
    Dim m 'As Long
    
    m = MinArr(15, 12, 22, 45, 100, 72, 11, 83, 46, 25, 35)
    MsgBox m '11
    m = MaxArr(15, 12, 22, 45, 100, 72, 11, 83, 46, 25, 35)
    MsgBox m '100
End Sub

Private Sub Command5_Click()
    MsgBox "Fibonacci(15) = " & MMath.Fibonacci(15)
End Sub

Private Sub Form_Load()
    MMath.Init
    
    'Debug.Print CalcPi
    'Debug.Print Fact(78)
    'Debug.Print CDec(4) * Atn(CDec(1))
    'Debug.Print CDec("3,141592653589792")
    'Debug.Print CDec("3,1415926535897932384626433832795") '02884197169399375105820974944592")
    'Debug.Print MMath.Pi
    'Debug.Print MMath.Euler
    Tests
    'Constants_ToTextBox LBConstants
    
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then
        Text1.Move L, t, W, H
    End If
End Sub

Sub Tests()
    TestConstants
    TestGgT_KgV_PFZ_FraC
    TestFactorials
    TestPrimes
    TestComplexNumbers
    TestQuadraticCubic
    TestPascalTriangle
    TestComplex
    
End Sub

Sub TestConstants()
    'With aLB
    AddItem "TestConstants:"
    AddItem "=============="
    
    AddItem "Pi          =" & MMath.Pi
    AddItem "2*Pi        =" & MMath.Pi2
    AddItem "Pi/2        =" & MMath.Pihalf
    AddItem "Euler       =" & MMath.Euler       ' As Variant As Decimal
    AddItem "SquareRoot2 =" & MMath.SquareRoot2 ' As Variant As Decimal
    AddItem "SquareRoot3 =" & MMath.SquareRoot3 ' As Variant As Decimal
    AddItem "GoldenRatio =" & MMath.GoldenRatio ' As Variant As Decimal

'Physikalische Konstanten
    AddItem "SpeedOfLight   =" & MMath.SpeedOfLight & " m/s"  'Lichtgeschwindigkeit im Vakuum      c
    AddItem "MassOfElektron =" & MMath.MassElektron   'Ruhemasse des Elektrons             me
    AddItem "MassOfProton   =" & MMath.MassProton     'Ruhemasse des Protons               mp
    AddItem "Gravitation    =" & MMath.Gravitation    'Newtonsche Gravitationskonstante    G
    AddItem "Avogadro       =" & MMath.Avogadro       'Avogadro-Konstante                  NA
    AddItem "ProtonCharge   =" & MMath.ElemCharge     'Elementarladung (des Protons)       e
    AddItem "PlanckQuantum  =" & MMath.PlanckQuantum  'Plancksches Wirkungsquantum         h
    AddItem "QuantumAlpha   =" & MMath.QuantumAlpha
    AddItem "ElectricPermittivity = " & MMath.ElecPermittvy ' Dielectrizitäts-Konstante  eps_0
    AddItem ""
    
End Sub

Sub TestGgT_KgV_PFZ_FraC()
    AddItem "TestGgT_KgV_PFZ_FraC"
    AddItem "===================="
    
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

Sub TestComplexNumbers()
    AddItem "TestComplexNumbers"
    AddItem "=================="
    
    Dim z1 As Complex: z1 = MMath.Real_ToComplex(8)
    AddItem "v=" & 8
    AddItem "z1(v)=" & MMath.Complex_ToStr(z1)
    
    AddItem ""
End Sub

Sub TestQuadraticCubic()
    AddItem "TestQuadraticCubic"
    AddItem "=================="
    Dim a As Double, b As Double, c As Double, d As Double
    Dim x1 As Double
    Dim x2 As Double, i2 As Double
    Dim x3 As Double, i3 As Double
    'a = 2 ' 1
    'b = 4 '0 '-1
    'c = -2 '6 '-4
    'd = -4 '21 '4
    
    a = 1: b = 8: c = -20
    x1 = 0: x2 = 0: x3 = 0
    
    AddItem Quadratic_ToStr(a, b, c)
    If Quadratic(a, b, c, x1, x2) Then
        'x1 = 2; x2 = -10
        AddItem "x1 = " & x1 & "; x2 = " & x2
    End If
    
    a = 2: b = -6: c = -4: d = -4
    AddItem Quadratic_ToStr(a, b, c)
    If Quadratic(a, b, c, x1, x2) Then
        'x1 = -1; x2 = -2
        AddItem "x1 = " & x1 & "; x2 = " & x2
    End If
    
    'a = 0.25: b = 0.75: c = -1.5: d = -2
    a = 2: b = 6: c = -4: d = -24
    x1 = 0: x2 = 0: i2 = 0: x3 = 0: i3 = 0
    
    AddItem Cubic_ToStr(a, b, c, d)
    
    If MMath.Cubic(a, b, c, d, x1, x2, i2, x3, i3) Then
        
        AddItem "x1 = " & x1
        
    End If
    AddItem ""
End Sub

Sub TestPascalTriangle()
    AddItem "TestPascalTriangle"
    AddItem "=================="
    Dim pt(): pt = MMath.PascalTriangle(12) 'max 1030 rows
    AddItem MMath.PascalTriangle_ToStr(pt)
    AddItem ""
End Sub

Sub TestComplex()
    AddItem "TestComplexNumbers"
    AddItem "=================="
    Dim z1 As Complex: z1 = MMath.Complex(1, 0.5)
    Dim z2 As Complex: z2 = MMath.Complex(2, 3)
    Dim z As Complex
    z = MMath.Complex_Add(z1, z2)
    AddItem "z = z1 + z2; z = " & MMath.Complex_ToStr(z1) & " + (" & MMath.Complex_ToStr(z2) & ") = " & MMath.Complex_ToStr(z)
    
    z = MMath.Complex_Mul(z1, z2)
    AddItem "z = z1 * z2; z = " & MMath.Complex_ToStr(z1) & " * (" & MMath.Complex_ToStr(z2) & ") = " & MMath.Complex_ToStr(z)
    
    z1 = MMath.Complex(-3, 4)
    z2 = MMath.Complex(1, 3)
    z = MMath.Complex_Div(z1, z2)
    AddItem "z = z1 / z2; z = " & MMath.Complex_ToStr(z1) & " / (" & MMath.Complex_ToStr(z2) & ") = " & MMath.Complex_ToStr(z)
    
    AddItem ""
End Sub

Sub AddItem(s As String)
    Text1.Text = Text1.Text & s & vbCrLf
End Sub
