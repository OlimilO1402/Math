VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9255
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
      Height          =   6135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Private Enum ECuberootfnd
'    ecr_cbrt_5d
'    ecr_pow_cbrtd
'    ecr_halley_cbrt1d
'    ecr_halley_cbrt2d
'    ecr_halley_cbrt3d
'    ecr_newton_cbrt1d
'    ecr_newton_cbrt2d
'    ecr_newton_cbrt3d
'    ecr_newton_cbrt4d
'End Enum

Private Sub Form_Load()
'    // a million uniform steps through the range from 0.0 to 1.0
'    // (doing uniform steps in the log scale would be better)
'    double a = 0.0;
    Dim a As Double
'    double b = 1.0;
    Dim b As Double: b = 1
'    int n = 1000000;
    Dim N As Long: N = 1000000

    'printf("32-bit float tests\n");
    'printf("----------------------------------------\n");
    'TestCubeRootf("cbrt_5f", cbrt_5f, a, b, n);
    'TestCubeRootf("pow", pow_cbrtf, a, b, n);
    'TestCubeRootf("halley x 1", halley_cbrt1f, a, b, n);
    'TestCubeRootf("halley x 2", halley_cbrt2f, a, b, n);
    'TestCubeRootf("newton x 1", newton_cbrt1f, a, b, n);
    'TestCubeRootf("newton x 2", newton_cbrt2f, a, b, n);
    'TestCubeRootf("newton x 3", newton_cbrt3f, a, b, n);
    'TestCubeRootf("newton x 4", newton_cbrt4f, a, b, n);
    'printf("\n\n");

    'printf("64-bit double tests\n");
    printf "64-bit double tests\n"
    'printf("----------------------------------------\n");
    printf "----------------------------------------"
    TestCubeRootd "cbrt_5d", ecr_cbrt_5d, a, b, N
    TestCubeRootd "pow", ecr_pow_cbrtd, a, b, N
    TestCubeRootd "halley x 1", ecr_halley_cbrt1d, a, b, N
    TestCubeRootd "halley x 2", ecr_halley_cbrt2d, a, b, N
    TestCubeRootd "halley x 3", ecr_halley_cbrt3d, a, b, N
    TestCubeRootd "newton x 1", ecr_newton_cbrt1d, a, b, N
    TestCubeRootd "newton x 2", ecr_newton_cbrt2d, a, b, N
    TestCubeRootd "newton x 3", ecr_newton_cbrt3d, a, b, N
    TestCubeRootd "newton x 4", ecr_newton_cbrt4d, a, b, N
    printf ""

    'getchar();

    'return 0;
End Sub

Sub printf(ByVal s As String)
    Text1.Text = Text1.Text & s & vbCrLf
End Sub

'// get accurate timer (Win32)
Function GetTimer() As Double
    'LARGE_INTEGER F, N;
    Dim f As Currency, N As Currency
    QueryPerformanceFrequency f
    QueryPerformanceCounter N
    'return double(N.QuadPart)/double(F.QuadPart);
    GetTimer = N / f
End Function

Function TestCubeRootd(ByVal szName As String, ByVal cbrt As ECuberootfnd, ByVal rA As Double, ByVal rB As Double, ByVal rN As Long) As Double

    Dim N As Long: N = rN
    
    Dim dd As Double: dd = (rB - rA) / N

    Dim i As Long 'int i=0

    Dim t As Double: t = GetTimer
    
    Dim s As Double '= 0.0
    Dim d As Double '= 0.0
    d = rA
    For i = 0 To N - 1
        d = d + dd
        s = s + cbrt(d)
    Next
    
    t = GetTimer() - t

    printf "%-10s %5.1f ms ", szName, t * 1000#

    Dim bits As Double '= 0.0;
    Dim maxre As Double '= 0.0;
    Dim worstx As Double '= 0.0;
    Dim worsty As Double '= 0.0;
    Dim minbits As Long: minbits = 64
    
    Dim i As Long
    d = rA
    For i = 0 To N - 1
        d = d + dd
        Dim a As Double: a = cbrt(d)
        Dim b As Double: b = pow(d, 1# / 3#)

        Dim bc As Long: bc = bits_of_precision(a, b) ' // min(53, count_matching_bitsd(a, b) - 12);
        bits = bits + bc

        If b > 0.000001 Then
            If (bc < minbits) Then
            
                bits_of_precision a, b
                minbits = bc
                worstx = d
                worsty = a
            End If
        End If
    Next

    bits = bits / N

    printf " %3d mbp  %6.3f abp\n", minbits, bits

    TestCubeRootd = s
End Function
