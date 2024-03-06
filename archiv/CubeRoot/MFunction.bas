Attribute VB_Name = "MFunction"
Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function CallWindowProcSng Lib "user32" Alias "CallWindowProcA" (ByRef lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Single
Private Declare Function CallWindowProcDbl Lib "user32" Alias "CallWindowProcA" (ByRef lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Double

Dim m_Args()  As Long
Dim m_UBnd    As Long
Dim m_nParams As Long
Dim m_Address As Long
' ASM Quellcode '
'Static m_asm(6) As Long
Dim m_asm(6) As Long

'// get accurate timer (Win32)
Public Function GetTimer() As Double
    'LARGE_INTEGER F, N;
    Dim f As Currency, N As Currency
    QueryPerformanceFrequency f
    QueryPerformanceCounter N
    'return double(N.QuadPart)/double(F.QuadPart);
    GetTimer = N / f
End Function

Public Function FncPtr(ByVal pFunc As Long) As Long
    FncPtr = pFunc
End Function

Public Sub Init(ByVal Address As Long, ByVal nParams As Long)
    
    m_Address = Address
    
    If m_asm(0) = 0 Then
        m_asm(0) = &H8BEC8B55:  m_asm(1) = &H5D8B0C4D:  m_asm(2) = &H744B4310
        m_asm(3) = &H50018B08:  m_asm(4) = &HEB04C183:  m_asm(5) = &H855FFF5
        m_asm(6) = &H10C2C9
    End If
    
    ' Ist das Array dimensioniert? '
    'On Error GoTo NoArgs
    m_nParams = nParams
    m_UBnd = nParams - 1 'VarArgs(0)
    'On Error GoTo 0
    
    ' Es gibt mindestens 1 Argument '
    ' Alle Argumente werden verkehrtherum in einen Long-Array geladen '
    
    'UBnd = UBound(VarArgs)
    If m_UBnd >= 0 Then
        ReDim m_Args(m_UBnd)
    End If
    
End Sub

Public Function InvokeSng(ParamArray VarArgs() As Variant) As Single
    
    Dim i As Long
    For i = 0 To m_UBnd
        Select Case VarType(VarArgs(i))
        Case vbLong, vbInteger, vbByte
            m_Args(m_UBnd - i) = VarArgs(i)
        Case vbString
            m_Args(m_UBnd - i) = StrPtr(StrConv(VarArgs(i), vbFromUnicode))
        Case Else
            m_Args(m_UBnd - i) = VarPtr(VarArgs(i))
        End Select
    Next
    
    ' Die Funktion wird per ASM aufgerufen '
    InvokeSng = CallWindowProcSng(m_asm(0), m_Address, m_Args(0), m_nParams, ByVal 0&)
    
'    Exit Function
'
'NoArgs:
'    ' Der Array ist nicht dimensioniert; es gibt also überhaupt keine Argumente! '
'    InvokeSng = CallWindowProcSng(asm(0), Address, 0, ByVal 0&, ByVal 0&)

End Function


Public Function InvokeDbl(ParamArray VarArgs() As Variant) As Double
    
'    Dim Args()  As Long
'    Dim UBnd    As Long
'    Dim I1      As Integer
'
'    ' ASM Quellcode '
'    Static asm(6) As Long
'
'    If asm(0) = 0 Then
'        asm(0) = &H8BEC8B55:  asm(1) = &H5D8B0C4D:  asm(2) = &H744B4310
'        asm(3) = &H50018B08:  asm(4) = &HEB04C183:  asm(5) = &H855FFF5
'        asm(6) = &H10C2C9
'    End If
'
'    ' Ist der Array dimensioniert? '
'    On Error GoTo NoArgs
'    UBnd = VarArgs(0)
'    On Error GoTo 0
    
    ' Es gibt mindestens 1 Argument '
    ' Alle Argumente werden verkehrtherum in einen Long-Array geladen '
    
'    UBnd = UBound(VarArgs)
    
    'ReDim Args(UBnd)
    Dim i As Long
    For i = 0 To m_UBnd
        Select Case VarType(VarArgs(i))
        Case vbLong, vbInteger, vbByte
            m_Args(m_UBnd - i) = VarArgs(i)
        Case vbString
            m_Args(m_UBnd - i) = StrPtr(StrConv(VarArgs(i), vbFromUnicode))
        Case Else
            m_Args(m_UBnd - i) = VarPtr(VarArgs(i))
        End Select
    Next
    
    ' Die Funktion wird per ASM aufgerufen '
    InvokeDbl = CallWindowProcDbl(m_asm(0), m_Address, m_Args(0), m_nParams, ByVal 0&)
    
'    Exit Function
'
'NoArgs:
'    ' Der Array ist nicht dimensioniert; es gibt also überhaupt keine Argumente! '
'    InvokeDbl = CallWindowProcDbl(asm(0), Address, 0, ByVal 0&, ByVal 0&)

End Function

