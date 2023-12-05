VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13575
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
   ScaleHeight     =   6135
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox LBConstants 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   13575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   4680
      TabIndex        =   4
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
      TabIndex        =   1
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
    Constants_ToListBox LBConstants
    
End Sub


Sub Constants_ToListBox(aLB As ListBox)
    With aLB
        .Clear
        .AddItem "Pi          =" & MMath.Pi
        .AddItem "2*Pi        =" & MMath.Pi2
        .AddItem "Pi/2        =" & MMath.Pihalf
        .AddItem "Euler       =" & MMath.Euler       ' As Variant As Decimal
        .AddItem "SquareRoot2 =" & MMath.SquareRoot2 ' As Variant As Decimal
        .AddItem "SquareRoot3 =" & MMath.SquareRoot3 ' As Variant As Decimal
        .AddItem "GoldenRatio =" & MMath.GoldenRatio ' As Variant As Decimal

'Physikalische Konstanten
        .AddItem "SpeedOfLight   =" & MMath.SpeedOfLight & " m/s"  'Lichtgeschwindigkeit im Vakuum      c
        .AddItem "MassOfElektron =" & MMath.MassElektron   'Ruhemasse des Elektrons             me
        .AddItem "MassOfProton   =" & MMath.MassProton     'Ruhemasse des Protons               mp
        .AddItem "Gravitation    =" & MMath.Gravitation    'Newtonsche Gravitationskonstante    G
        .AddItem "Avogadro       =" & MMath.Avogadro       'Avogadro-Konstante                  NA
        .AddItem "ProtonCharge   =" & MMath.ElemCharge     'Elementarladung (des Protons)       e
        .AddItem "PlanckQuantum  =" & MMath.PlanckQuantum  'Plancksches Wirkungsquantum         h
        .AddItem "QuantumAlpha   =" & MMath.QuantumAlpha
        .AddItem "ElectricPermittivity = " & MMath.ElecPermittvy ' Dielectrizitäts-Konstante  eps_0
        Dim n1 As Long: n1 = 1234
        Dim n2 As Long: n2 = 56
        .AddItem "Primefactors(" & n1 & ") = " & PFZ(n1)
        .AddItem "GreatestCommonDivisor(" & n1 & ", " & n2 & ") = " & MMath.GreatestCommonDivisor(n1, n2)
        .AddItem "LeastCommonMultiple(" & n1 & ", " & n2 & ") = " & MMath.LeastCommonMultiple(n1, n2)
        n2 = 3456
        Dim nn As Long: nn = n1
        Dim za As Long: za = n2
        MMath.CancelFraction nn, za
        .AddItem "CancelFraction(" & n1 & ", " & n2 & ") = " & nn & " / " & za
        Dim i As Long, s As String
        For i = 0 To 5
            .AddItem CStr(i) & "! = " & MMath.Fact(i)
        Next
        For i = 22 To 27
            .AddItem CStr(i) & "! = " & MMath.Fact(i)
        Next
        For i = 28 To 30
            .AddItem CStr(i) & "! = " & MMath.Fact(i)
        Next
        For i = 168 To 171
            .AddItem CStr(i) & "! = " & MMath.Fact(i)
        Next
        'Dim n As Long: n = 5
        'Dim n As Long: n = 96211
        'Dim n As Long: n = 99991
        Dim n As Long: n = 99991
        .AddItem "IsPrimeA(" & n & ") = " & MMath.IsPrime(n)
    End With
End Sub

Public Function CalcPi()
    Dim sqr3: sqr3 = CDec("1,7320508075688772935274463415") '058723669428052538103806280558069794519330169088000370811461867572485756")
    Dim sum: sum = CDec(0)
    Dim n As Long
    For n = 1 To 40
        sum = sum + -Fact(2 * n - 2) / (2 ^ (4 * n - 2) * Fact(n - 1) ^ 2 * (2 * n - 3) * (2 * n + 1))
    Next
    Dim Pi
    Pi = 3 * sqr3 / 4 + 24 * sum
    CalcPi = Pi
End Function

