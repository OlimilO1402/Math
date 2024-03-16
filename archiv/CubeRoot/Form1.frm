VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   11910
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnRootTests 
      Caption         =   "Root Tests"
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton BtnTestValues 
      Caption         =   "Test Values"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton BtnTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
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
      Height          =   9135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.xbeat.net/vbspeed/
'https://learn.microsoft.com/de-de/cpp/c-runtime-library/reference/crt-alphabetical-function-reference?view=msvc-170

'https://www.youtube.com/watch?v=HekvDjzDZCo
'https://de.wikipedia.org/wiki/X87

'Private Enum ECbrt
'
'    AddressOf_cbrt_5f
'    AddressOf_pow_cbrtf
'    AddressOf_halley_cbrt1f
'    AddressOf_halley_cbrt2f
'    AddressOf_halley_cbrt3f
'    AddressOf_newton_cbrt1f
'    AddressOf_newton_cbrt2f
'    AddressOf_newton_cbrt3f
'    AddressOf_newton_cbrt4f
'
'    AddressOf_cbrt_5d
'    AddressOf_pow_cbrtd
'    AddressOf_halley_cbrt1d
'    AddressOf_halley_cbrt2d
'    AddressOf_halley_cbrt3d
'    AddressOf_newton_cbrt1d
'    AddressOf_newton_cbrt2d
'    AddressOf_newton_cbrt3d
'    AddressOf_newton_cbrt4d
'
'End Enum
Option Explicit
Private Declare Function Dbl_FSqrt Lib "FPUx87" (ByVal Value As Double) As Double
Private m_n As Long
Private m_Tests() As MRootTests.TRootTest

Private Sub Command3_Click()
    Dim v As Double: v = 2
    v = Dbl_FSqrt(v)
    
    MsgBox v
End Sub

Private Sub Form_Load()
    m_n = 1000000 '0 'ten no one million times
    ReDim m_Tests(0 To 7)
    Dim c As Long
    m_Tests(0) = MRootTests.New_RootTest(m_n)
    MRootTests.RootTest_InitRandomNumbers m_Tests(0):   c = c + 1
    m_Tests(c) = MRootTests.RootTest_Clone(m_Tests(0)): c = c + 1
    m_Tests(c) = MRootTests.RootTest_Clone(m_Tests(0)): c = c + 1
    m_Tests(c) = MRootTests.RootTest_Clone(m_Tests(0)): c = c + 1
    m_Tests(c) = MRootTests.RootTest_Clone(m_Tests(0)): c = c + 1
    m_Tests(c) = MRootTests.RootTest_Clone(m_Tests(0)): c = c + 1
    m_Tests(c) = MRootTests.RootTest_Clone(m_Tests(0)): c = c + 1
    m_Tests(c) = MRootTests.RootTest_Clone(m_Tests(0)): c = c + 1

End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then Text1.Move L, t, W, H
End Sub

Private Sub BtnRootTests_Click()
    Dim c As Long
    Dim i As Long
    
    'c = c + 1
    With m_Tests(c)
        .StartTime = GetTimer
        For i = 0 To .Count - 1
            With .Test(i)
                .Result = .Value ^ (1 / 2)
            End With
        Next
        .EndTime = GetTimer
    End With
    printf RootTest_ToStr(m_Tests(c), "n ^ (1 / 2)        ")
    DoEvents
    
    c = c + 1
    With m_Tests(c)
        .StartTime = GetTimer
        For i = 0 To .Count - 1
            With .Test(i)
                .Result = VBA.Math.Sqr(.Value)
            End With
        Next
        .EndTime = GetTimer
    End With
    printf RootTest_ToStr(m_Tests(c), "VBA.Math.Sqr(n)    ")
    DoEvents
    
    If Not MRootTests.RootTest_ResultsAreEqual(m_Tests(c), m_Tests(0)) Then
        printf "The results are not all equal"
    End If
    
    c = c + 1
    With m_Tests(c)
        .StartTime = GetTimer
        For i = 0 To .Count - 1
            With .Test(i)
                .Result = Dbl_FSqrt(.Value)
            End With
        Next
        .EndTime = GetTimer
    End With
    printf RootTest_ToStr(m_Tests(c), "Dbl_FSqrt(n)       ")
    DoEvents
    
    If Not MRootTests.RootTest_ResultsAreEqual(m_Tests(c), m_Tests(0)) Then
        printf "The results are not all equal"
    End If
    
    c = c + 1
    With m_Tests(c)
        .StartTime = GetTimer
        For i = 0 To .Count - 1
            With .Test(i)
                .Result = .Value ^ (1 / 3)
            End With
        Next
        .EndTime = GetTimer
    End With
    printf RootTest_ToStr(m_Tests(c), "n ^ (1 / 3)        ")
    DoEvents
    
    c = c + 1
    With m_Tests(c)
        .StartTime = GetTimer
        For i = 0 To .Count - 1
            With .Test(i)
                .Result = MCubeRoot.newton_cbrt4d(.Value)
            End With
        Next
        .EndTime = GetTimer
    End With
    printf RootTest_ToStr(m_Tests(c), "cbrt.newton_cbrt4d ")
    DoEvents
    
    If Not MRootTests.RootTest_ResultsAreEqual(m_Tests(c), m_Tests(3)) Then
        printf "The results are not all equal"
    End If
    
    c = c + 1
    With m_Tests(c)
        .StartTime = GetTimer
        For i = 0 To .Count - 1
            With .Test(i)
                .Result = MCubeRoot.halley_cbrt3d(.Value)
            End With
        Next
        .EndTime = GetTimer
    End With
    printf RootTest_ToStr(m_Tests(c), "cbrt.halley_cbrt3d ")
    DoEvents
    
    If Not MRootTests.RootTest_ResultsAreEqual(m_Tests(c), m_Tests(3)) Then
        printf "The results are not all equal"
    End If
    
    c = c + 1
    With m_Tests(c)
        .StartTime = GetTimer
        For i = 0 To .Count - 1
            With .Test(i)
                .Result = MCubeRoot2.newton_cbrt4d(.Value)
            End With
        Next
        .EndTime = GetTimer
    End With
    printf RootTest_ToStr(m_Tests(c), "cbrt2.newton_cbrt4d")
    DoEvents
    
    If Not MRootTests.RootTest_ResultsAreEqual(m_Tests(c), m_Tests(3)) Then
        printf "The results are not all equal"
    End If
    
    c = c + 1
    With m_Tests(c)
        .StartTime = GetTimer
        For i = 0 To .Count - 1
            With .Test(i)
                .Result = MCubeRoot2.halley_cbrt3d(.Value)
            End With
        Next
        .EndTime = GetTimer
    End With
    printf RootTest_ToStr(m_Tests(c), "cbrt2.halley_cbrt3d")
    DoEvents
    
    If Not MRootTests.RootTest_ResultsAreEqual(m_Tests(c), m_Tests(3)) Then
        printf "The results are not all equal"
    End If
    
End Sub

Sub printClear()
    Text1.Text = vbNullString
End Sub

Sub printf(ByVal s As String)
    Text1.Text = Text1.Text & s & vbCrLf
End Sub
'
'Private Sub BtnTest_Click()
'    printClear
'
''    // a million uniform steps through the range from 0.0 to 1.0
''    // (doing uniform steps in the log scale would be better)
''    double a = 0.0;
'    Dim a As Double
''    double b = 1.0;
'    Dim b As Double: b = 1
''    int n = 1000000;
'    Dim n As Long: n = m_n ' one million
'
'    printf "32-bit float tests n=" & n & " times"
'    printf "----------------------------------"
'    TestCubeRootf "cbrt_5f", AddressOf_cbrt_5f, a, b, n
'    TestCubeRootf "pow", AddressOf_pow_cbrtf, a, b, n
'    TestCubeRootf "halley x 1", AddressOf_halley_cbrt1f, a, b, n
'    TestCubeRootf "halley x 2", AddressOf_halley_cbrt2f, a, b, n
'    TestCubeRootf "halley x 3", AddressOf_halley_cbrt3f, a, b, n
'    TestCubeRootf "newton x 1", AddressOf_newton_cbrt1f, a, b, n
'    TestCubeRootf "newton x 2", AddressOf_newton_cbrt2f, a, b, n
'    TestCubeRootf "newton x 3", AddressOf_newton_cbrt3f, a, b, n
'    TestCubeRootf "newton x 4", AddressOf_newton_cbrt4f, a, b, n
'    printf ""
'
'    printf "64-bit double tests n=" & n & " times"
'    printf "-----------------------------------"
'    TestCubeRootd "cbrt_5d", AddressOf_cbrt_5d, a, b, n
'    TestCubeRootd "pow", AddressOf_pow_cbrtd, a, b, n
'    TestCubeRootd "halley x 1", AddressOf_halley_cbrt1d, a, b, n
'    TestCubeRootd "halley x 2", AddressOf_halley_cbrt2d, a, b, n
'    TestCubeRootd "halley x 3", AddressOf_halley_cbrt3d, a, b, n
'    TestCubeRootd "newton x 1", AddressOf_newton_cbrt1d, a, b, n
'    TestCubeRootd "newton x 2", AddressOf_newton_cbrt2d, a, b, n
'    TestCubeRootd "newton x 3", AddressOf_newton_cbrt3d, a, b, n
'    TestCubeRootd "newton x 4", AddressOf_newton_cbrt4d, a, b, n
'    printf ""
'
'End Sub
'
'Private Sub BtnTestValues_Click()
'    printClear
'
'    Dim d As Double
'    printf "Test halley_cbrt1d"
'    printf "------------------"
'    d = 8:                    printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt1d(d)
'    d = 12:                   printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt1d(d)
'    d = 12345:                printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt1d(d)
'    d = 1234567:              printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt1d(d)
'    d = 123456789:            printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt1d(d)
'    d = 12345678901#:         printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt1d(d)
'    d = 1234567890123#:       printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt1d(d)
'    d = 123456789012345#:     printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt1d(d)
'    d = 1.23456789012346E+16: printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt1d(d)
'
'    printf ""
'    printf "Test halley_cbrt2d"
'    printf "------------------"
'    d = 8:                    printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt2d(d)
'    d = 12:                   printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt2d(d)
'    d = 12345:                printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt2d(d)
'    d = 1234567:              printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt2d(d)
'    d = 123456789:            printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt2d(d)
'    d = 12345678901#:         printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt2d(d)
'    d = 1234567890123#:       printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt2d(d)
'    d = 123456789012345#:     printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt2d(d)
'    d = 1.23456789012346E+16: printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt2d(d)
'
'    printf ""
'
'    d = 8:                    printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt3d(d)
'    d = 12:                   printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt3d(d)
'    d = 12345:                printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt3d(d)
'    d = 1234567:              printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt3d(d)
'    d = 123456789:            printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt3d(d)
'    d = 12345678901#:         printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt3d(d)
'    d = 1234567890123#:       printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt3d(d)
'    d = 123456789012345#:     printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt3d(d)
'    d = 1.23456789012346E+16: printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.halley_cbrt3d(d)
'
'    printf ""
'
'    d = 8:                    printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.newton_cbrt4d(d)
'    d = 12:                   printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.newton_cbrt4d(d)
'    d = 12345:                printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.newton_cbrt4d(d)
'    d = 1234567:              printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.newton_cbrt4d(d)
'    d = 123456789:            printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.newton_cbrt4d(d)
'    d = 12345678901#:         printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.newton_cbrt4d(d)
'    d = 1234567890123#:       printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.newton_cbrt4d(d)
'    d = 123456789012345#:     printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.newton_cbrt4d(d)
'    d = 1.23456789012346E+16: printf "cbrt(" & d & ") = " & MCubeRoot2.pow_cbrtd(d) & " " & MCubeRoot2.newton_cbrt4d(d)
'
'End Sub
'
'Private Sub Command1_Click()
'    printClear
'
'    Dim i As Long, u As Long: u = m_n '10000000 'ten million
'    Dim d As Double: d = 123456789012345#
'    Dim v As Double
'    Dim dt0 As Double
'    Dim dt1 As Double
'    Dim dt2 As Double
'    Dim dt3 As Double
'    Dim dt4 As Double
'
'    printf "Testing " & u & " times"
'    printf "-----------------------"
'    DoEvents
'
'    printf "Test: d ^ (1 / 2) vs VBA.MAth.Sqr(d)"
'    printf "------------------------------------"
'    DoEvents
'
'    dt1 = GetTimer
'    For i = 0 To u
'        v = d ^ (1 / 2)
'    Next
'    dt1 = GetTimer - dt1
'    printf "v=" & v
'
'    dt2 = GetTimer
'    For i = 0 To u
'        v = VBA.Math.Sqr(d)
'    Next
'    dt2 = GetTimer - dt2
'    printf "v=" & v
'    printf "dt1 = " & dt1 & "  dt2 = " & dt2
'
'    printf ""
'
'    printf "d ^ (1 / 3) vs cbr.newton_cbrt4d(d) vs cbr2.newton_cbrt4d(d) vs cbr.halley_cbrt3d(d) vs cbr2.halley_cbrt3d(d)"
'    printf "-------------------------------------------------------------------------------------------------------------"
'    DoEvents
'
'    dt0 = GetTimer
'    For i = 0 To u
'        v = d ^ (1 / 3)
'    Next
'    dt0 = GetTimer - dt0
'    printf "v=" & v
'
'    dt1 = GetTimer
'    For i = 0 To u
'        v = MCubeRoot.newton_cbrt4d(d)
'    Next
'    dt1 = GetTimer - dt1
'    printf "v=" & v
'
'    dt2 = GetTimer
'    For i = 0 To u
'        v = MCubeRoot2.newton_cbrt4d(d)
'    Next
'    dt2 = GetTimer - dt2
'    printf "v=" & v
'
'    dt3 = GetTimer
'    For i = 0 To u
'        v = MCubeRoot.halley_cbrt3d(d)
'    Next
'    dt3 = GetTimer - dt3
'    printf "v=" & v
'
'    dt4 = GetTimer
'    For i = 0 To u
'        v = MCubeRoot2.halley_cbrt3d(d)
'    Next
'    dt4 = GetTimer - dt4
'    printf "v=" & v
'
'    printf "dt0 = " & dt0 * 1000 & " ms;  dt1 = " & dt1 * 1000 & " ms;  dt2 = " & dt2 * 1000 & " ms;  dt3 = " & dt3 * 1000 & " ms;  dt4 = " & dt4
'
'End Sub
'
'Private Sub Command2_Click()
'    Dim r As Double
'    Dim d0 As Double
'    Dim d1 As Double
'    Dim i As Long, n As Long: n = m_n '10000000 ten million
'    MsgBox "Test calculating " & n & " cuberoots or random numbers with halley_cbrt3d"
'    Randomize Timer
'    For i = 0 To 10000000
'        r = Rnd * 123456789012345#
'        d0 = r ^ (1 / 3)
'        d1 = MCubeRoot2.halley_cbrt3d(r)
'        If Abs(d1 - d0) > 0.0000000001 Then
'            MsgBox "d0: " & d0 & " d1: " & d1
'        End If
'    Next
'    MsgBox "OK"
'End Sub
'
'Private Function TestCubeRootf(ByVal szName As String, ByVal e_cbrt As ECbrt, ByVal rA As Single, ByVal rB As Single, ByVal rN As Long) As Single
'
'    Dim n As Long: n = rN
'    Dim dd As Single: dd = (rB - rA) / n
'    Dim i As Long 'int i=0
'    Dim s As Single '= 0.0
'    Dim d As Single '= 0.0
'    d = rA
'
'    Dim t As Single: t = GetTimer
'    Select Case e_cbrt
'    Case AddressOf_cbrt_5f
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.cbrt_5f(d): Next
'    Case AddressOf_pow_cbrtf
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.pow_cbrtf(d): Next
'    Case AddressOf_halley_cbrt1f
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.halley_cbrt1f(d): Next
'    Case AddressOf_halley_cbrt2f
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.halley_cbrt2f(d): Next
'    Case AddressOf_halley_cbrt3f
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.newton_cbrt3f(d): Next
'    Case AddressOf_newton_cbrt1f
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.newton_cbrt1f(d): Next
'    Case AddressOf_newton_cbrt2f
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.newton_cbrt2f(d): Next
'    Case AddressOf_newton_cbrt3f
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.newton_cbrt3f(d): Next
'    Case AddressOf_newton_cbrt4f
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.newton_cbrt4f(d): Next
'    End Select
'
'    t = GetTimer() - t
'
'    printf szName & " " & t * 1000# & " ms"
'
''    Dim bits    As Single '= 0.0;
''    Dim maxre   As Single '= 0.0;
''    Dim worstx  As Single '= 0.0;
''    Dim worsty  As Single '= 0.0;
''    Dim minbits As Long: minbits = 32
''    Dim a As Single
''    Dim b As Single
''    Dim bc As Long
''
''    d = rA
''
''    Select Case e_cbrt
''    Case AddressOf_cbrt_5f
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.cbrt_5f(d): b = d ^ (1# / 3#): bc = bits_of_precisionS(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionS a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_pow_cbrtf
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.pow_cbrtf(d): b = d ^ (1# / 3#): bc = bits_of_precisionS(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionS a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_halley_cbrt1f
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.halley_cbrt1f(d): b = d ^ (1# / 3#): bc = bits_of_precisionS(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionS a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_halley_cbrt2f
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.halley_cbrt2f(d): b = d ^ (1# / 3#): bc = bits_of_precisionS(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionS a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_halley_cbrt3f
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.newton_cbrt3f(d): b = d ^ (1# / 3#): bc = bits_of_precisionS(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionS a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_newton_cbrt1f
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.newton_cbrt1f(d): b = d ^ (1# / 3#): bc = bits_of_precisionS(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionS a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_newton_cbrt2f
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.newton_cbrt2f(d): b = d ^ (1# / 3#): bc = bits_of_precisionS(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionS a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_newton_cbrt3f
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.newton_cbrt3f(d): b = d ^ (1# / 3#): bc = bits_of_precisionS(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionS a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_newton_cbrt4f
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.newton_cbrt4f(d): b = d ^ (1# / 3#): bc = bits_of_precisionS(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionS a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    End Select
''
''    bits = bits / N
''
''    printf minbits & " minbits " & bits & " actualbits"
'
'    TestCubeRootf = s
'End Function
'
'Private Function TestCubeRootd(ByVal szName As String, ByVal e_cbrt As ECbrt, ByVal rA As Double, ByVal rB As Double, ByVal rN As Long) As Double
'
'    Dim n As Long: n = rN
'    Dim dd As Double: dd = (rB - rA) / n
'    Dim i As Long
'    Dim t As Double: t = GetTimer
'    Dim s As Double '= 0.0
'    Dim d As Double '= 0.0
'    d = rA
'
'    Select Case e_cbrt
'    Case AddressOf_cbrt_5d
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.cbrt_5d(d): Next
'    Case AddressOf_pow_cbrtd
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.pow_cbrtd(d): Next
'    Case AddressOf_halley_cbrt1d
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.halley_cbrt1d(d): Next
'    Case AddressOf_halley_cbrt2d
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.halley_cbrt2d(d): Next
'    Case AddressOf_halley_cbrt3d
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.halley_cbrt3d(d): Next
'    Case AddressOf_newton_cbrt1d
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.newton_cbrt1d(d): Next
'    Case AddressOf_newton_cbrt2d
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.newton_cbrt2d(d): Next
'    Case AddressOf_newton_cbrt3d
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.newton_cbrt3d(d): Next
'    Case AddressOf_newton_cbrt4d
'        For i = 0 To n - 1: d = d + dd: s = s + MCubeRoot2.newton_cbrt4d(d): Next
'    End Select
'
'    t = GetTimer() - t
'
'    printf szName & " " & t * 1000# & " ms"
'
''    Dim bits    As Double '= 0.0;
''    Dim maxre   As Double '= 0.0;
''    Dim worstx  As Double '= 0.0;
''    Dim worsty  As Double '= 0.0;
''    Dim minbits As Long: minbits = 64
''    Dim a As Double
''    Dim b As Double
''    Dim bc As Long
''
''    d = rA
''    Select Case e_cbrt
''    Case AddressOf_cbrt_5d
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.cbrt_5d(d): b = d ^ (1# / 3#): bc = bits_of_precisionD(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionD a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_pow_cbrtd
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.pow_cbrtd(d): b = d ^ (1# / 3#): bc = bits_of_precisionD(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionD a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_halley_cbrt1d
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.halley_cbrt1d(d): b = d ^ (1# / 3#): bc = bits_of_precisionD(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionD a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_halley_cbrt2d
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.halley_cbrt2d(d): b = d ^ (1# / 3#): bc = bits_of_precisionD(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionD a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_halley_cbrt3d
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.halley_cbrt3d(d): b = d ^ (1# / 3#): bc = bits_of_precisionD(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionD a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_newton_cbrt1d
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.newton_cbrt1d(d): b = d ^ (1# / 3#): bc = bits_of_precisionD(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionD a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_newton_cbrt2d
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.newton_cbrt2d(d): b = d ^ (1# / 3#): bc = bits_of_precisionD(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionD a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_newton_cbrt3d
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.newton_cbrt3d(d): b = d ^ (1# / 3#): bc = bits_of_precisionD(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionD a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    Case AddressOf_newton_cbrt4d
''        For i = 0 To N - 1
''            d = d + dd: a = MCubeRoot2.newton_cbrt4d(d): b = d ^ (1# / 3#): bc = bits_of_precisionD(a, b): bits = bits + bc
''            If b > 0.000001 Then If (bc < minbits) Then bits_of_precisionD a, b: minbits = bc: worstx = d: worsty = a
''        Next
''    End Select
''
''    bits = bits / N
''
''    printf minbits & " minbits " & bits & " actualbits"
'
'    TestCubeRootd = s
'End Function
'
