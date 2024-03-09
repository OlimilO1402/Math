Attribute VB_Name = "MCubeRoot"
Option Explicit
Type TLong
    Value As Long
End Type
Type TSingle
    Value As Single
End Type
Type TLong2
    Value0 As Long
    Value1 As Long
End Type
Type TDouble
    Value As Double
End Type

'// estimate bits of precision (32-bit float case)
'inline int bits_of_precision(float a, float b)
Public Function bits_of_precisionS(ByVal a As Single, ByVal b As Single) As Long
    
    Dim kd As Single: kd = 1# / VBA.Math.Log(2#)
    
    If a = b Then
        bits_of_precisionS = 23
        Exit Function
    End If
    
    Dim kdmin As Single: kdmin = 2 ^ -23
    
    Dim d As Single: d = Abs(a - b)
    If (d < kdmin) Then
        bits_of_precisionS = 23
        Exit Function
    End If
    
    bits_of_precisionS = Int(-Log(d) * kd)
    
End Function

''// estiamte bits of precision (64-bit double case)
'inline int bits_of_precision(double a, double b)
Function bits_of_precisionD(ByVal a As Double, ByVal b As Double) As Long
    Dim kd As Double: kd = 1# / VBA.Math.Log(2#)
    
    If a = b Then
        bits_of_precisionD = 52
        Exit Function
    End If
    
    Dim kdmin As Double: kdmin = 2# ^ -52#
    
    Dim d As Double: d = Abs(a - b)
    If d < kdmin Then
        bits_of_precisionD = 52
        Exit Function
    End If
    
    bits_of_precisionD = Int(-Log(d) * kd)
    
End Function

'// cube root via x^(1/3)
Function pow_cbrtf(ByVal x As Single) As Single
    pow_cbrtf = x ^ (1! / 3!)
End Function

'// cube root via x^(1/3)
Function pow_cbrtd(ByVal x As Double) As Double
    pow_cbrtd = x ^ (1# / 3#)
End Function

'// cube root approximation using bit hack for 32-bit float
'__forceinline float cbrt_5f(float f)
Function cbrt_5f(ByVal f As Single) As Single
    Dim t As TSingle: t.Value = f
    Dim p As TLong:    LSet p = t: p.Value = p.Value \ 3 + 709921077
    LSet t = p
    cbrt_5f = t.Value
End Function
'      8 ->   1,96252  correct is   2
'     12 ->   2,258374 correct is   2,2894284851066637356160844238794
'  12345 ->  23,43744  correct is  23,111618749807268680871973329588
'1234567 -> 108,052    correct is 107,27657218553581512232724447941
 
'// cube root approximation using bit hack for 64-bit float
'// adapted from Kahan's cbrt
'__forceinline double cbrt_5d(double d)
'{
'    const unsigned int B1 = 715094163;
'    double t = 0.0;
'    unsigned int* pt = (unsigned int*) &t;
'    unsigned int* px = (unsigned int*) &d;
'    pt[1]=px[1]/3+B1;
'    return t;
'}
Function cbrt_5d(ByVal d As Double) As Double
    Dim t As TDouble: t.Value = d
    Dim p As TLong2:   LSet p = t: p.Value1 = p.Value1 \ 3 + 715094163
    LSet t = p
    cbrt_5d = t.Value
End Function
'      8 ->   1,96693706512451 correct is   2
'     12 ->   2,26720809936523 correct is   2,2894284851066637356160844238794
'  12345 ->  23,5081024169922  correct is  23,111618749807268680871973329588
'1234567 -> 108,334655761719   correct is 107,27657218553581512232724447941
'
'// cube root approximation using bit hack for 64-bit float
'// adapted from Kahan's cbrt
'__forceinline double quint_5d(double d)
'{
'    return sqrt(sqrt(d));
'
'    const unsigned int B1 = 71509416*5/3;
'    double t = 0.0;
'    unsigned int* pt = (unsigned int*) &t;
'    unsigned int* px = (unsigned int*) &d;
'    pt[1]=px[1]/5+B1;
'    return t;
'}
Function quint_5d(ByVal d As Double) As Double
    quint_5d = VBA.Math.Sqr(VBA.Math.Sqr(d))
    Exit Function
End Function

'
'// iterative cube root approximation using Halley's method (float)
'__forceinline float cbrta_halleyf(const float a, const float R)
'{
'    const float a3 = a*a*a;
'    const float b= a * (a3 + R + R) / (a3 + a3 + R);
'    return b;
'}
Function cbrta_halleyf(ByVal a As Single, ByVal r As Single) As Single
    Dim a3 As Single: a3 = a * a * a
    cbrta_halleyf = a * (a3 + r + r) / (a3 + a3 + r)
End Function

'
'// iterative cube root approximation using Halley's method (double)
'__forceinline double cbrta_halleyd(const double a, const double R)
'{
'    const double a3 = a*a*a;
'    const double b= a * (a3 + R + R) / (a3 + a3 + R);
'    return b;
'}
Function cbrta_halleyd(ByVal a As Double, ByVal r As Double) As Double
    Dim a3 As Double: a3 = a * a * a
    cbrta_halleyd = a * (a3 + r + r) / (a3 + a3 + r)
End Function

'
'// iterative cube root approximation using Newton's method (float)
'__forceinline float cbrta_newtonf(const float a, const float x)
'{
'//    return (1.0 / 3.0) * ((a + a) + x / (a * a));
'    return a - (1.0f / 3.0f) * (a - x / (a*a));
'}
Function cbrta_newtonf(ByVal a As Single, ByVal x As Single) As Single
'//    return (1.0 / 3.0) * ((a + a) + x / (a * a));
    cbrta_newtonf = a - (1! / 3!) * (a - x / (a * a))
End Function

'
'// iterative cube root approximation using Newton's method (double)
'__forceinline double cbrta_newtond(const double a, const double x)
'{
'    return (1.0/3.0) * (x / (a*a) + 2*a);
'}
Function cbrta_newtond(ByVal a As Double, ByVal x As Double) As Double
    cbrta_newtond = (1# / 3#) * (x / (a * a) + 2 * a)
End Function

'
'// cube root approximation using 1 iteration of Halley's method (double)
'double halley_cbrt1d(double d)
'{
'    double a = cbrt_5d(d);
'    return cbrta_halleyd(a, d);
'}
Function halley_cbrt1d(ByVal d As Double) As Double
    Dim a As Double: a = cbrt_5d(d)
    halley_cbrt1d = cbrta_halleyd(a, d)
End Function

'
'// cube root approximation using 1 iteration of Halley's method (float)
'float halley_cbrt1f(float d)
'{
'    float a = cbrt_5f(d);
'    return cbrta_halleyf(a, d);
'}
Function halley_cbrt1f(ByVal f As Single) As Single
    Dim a As Single: a = cbrt_5f(f)
    halley_cbrt1f = cbrta_halleyf(a, f)
End Function

'
'// cube root approximation using 2 iterations of Halley's method (double)
'double halley_cbrt2d(double d)
'{
'    double a = cbrt_5d(d);
'    a = cbrta_halleyd(a, d);
'    return cbrta_halleyd(a, d);
'}
Function halley_cbrt2d(ByVal d As Double) As Double
    Dim a As Double: a = cbrt_5d(d)
    a = cbrta_halleyd(a, d)
    halley_cbrt2d = cbrta_halleyd(a, d)
End Function

'
'// cube root approximation using 3 iterations of Halley's method (double)
'double halley_cbrt3d(double d)
'{
'    double a = cbrt_5d(d);
'    a = cbrta_halleyd(a, d);
'    a = cbrta_halleyd(a, d);
'    return cbrta_halleyd(a, d);
'}

Function halley_cbrt3d(ByVal d As Double) As Double
    Dim a As Double: a = cbrt_5d(d)
    a = cbrta_halleyd(a, d)
    a = cbrta_halleyd(a, d)
    halley_cbrt3d = cbrta_halleyd(a, d)
End Function

'
'
'// cube root approximation using 2 iterations of Halley's method (float)
'float halley_cbrt2f(float d)
'{
'    float a = cbrt_5f(d);
'    a = cbrta_halleyf(a, d);
'    return cbrta_halleyf(a, d);
'}
Function halley_cbrt2f(ByVal d As Single) As Single
    Dim a As Single: a = cbrt_5f(d)
    a = cbrta_halleyf(a, d)
    halley_cbrt2f = cbrta_halleyf(a, d)
End Function

'
'// cube root approximation using 1 iteration of Newton's method (double)
'double newton_cbrt1d(double d)
'{
'    double a = cbrt_5d(d);
'    return cbrta_newtond(a, d);
'}
Function newton_cbrt1d(ByVal d As Double) As Double
    Dim a As Double: a = cbrt_5d(d)
    newton_cbrt1d = cbrta_newtond(a, d)
End Function

'
'// cube root approximation using 2 iterations of Newton's method (double)
'double newton_cbrt2d(double d)
'{
'    double a = cbrt_5d(d);
'    a = cbrta_newtond(a, d);
'    return cbrta_newtond(a, d);
'}
Function newton_cbrt2d(d As Double) As Double
    Dim a As Double: a = cbrt_5d(d)
    a = cbrta_newtond(a, d)
    newton_cbrt2d = cbrta_newtond(a, d)
End Function

'
'// cube root approximation using 3 iterations of Newton's method (double)
'double newton_cbrt3d(double d)
'{
'    double a = cbrt_5d(d);
'    a = cbrta_newtond(a, d);
'    a = cbrta_newtond(a, d);
'    return cbrta_newtond(a, d);
'}
Function newton_cbrt3d(ByVal d As Double) As Double
    Dim a As Double: a = cbrt_5d(d)
    a = cbrta_newtond(a, d)
    a = cbrta_newtond(a, d)
    newton_cbrt3d = cbrta_newtond(a, d)
End Function

'
'// cube root approximation using 4 iterations of Newton's method (double)
'double newton_cbrt4d(double d)
'{
'    double a = cbrt_5d(d);
'    a = cbrta_newtond(a, d);
'    a = cbrta_newtond(a, d);
'    a = cbrta_newtond(a, d);
'    return cbrta_newtond(a, d);
'}
Function newton_cbrt4d(ByVal d As Double) As Double
    Dim a As Double: a = cbrt_5d(d)
    a = cbrta_newtond(a, d)
    a = cbrta_newtond(a, d)
    a = cbrta_newtond(a, d)
    newton_cbrt4d = cbrta_newtond(a, d)
End Function

'
'// cube root approximation using 2 iterations of Newton's method (float)
'float newton_cbrt1f(float d)
'{
'    float a = cbrt_5f(d);
'    return cbrta_newtonf(a, d);
'}
Function newton_cbrt1f(ByVal d As Single) As Single
    Dim a As Single: a = cbrt_5f(d)
    newton_cbrt1f = cbrta_newtonf(a, d)
End Function

'
'// cube root approximation using 2 iterations of Newton's method (float)
'float newton_cbrt2f(float d)
'{
'    float a = cbrt_5f(d);
'    a = cbrta_newtonf(a, d);
'    return cbrta_newtonf(a, d);
'}
Function newton_cbrt2f(ByVal d As Single) As Single
    Dim a As Single: a = cbrt_5f(d)
    a = cbrta_newtonf(a, d)
    newton_cbrt2f = cbrta_newtonf(a, d)
End Function

'
'// cube root approximation using 3 iterations of Newton's method (float)
'float newton_cbrt3f(float d)
'{
'    float a = cbrt_5f(d);
'    a = cbrta_newtonf(a, d);
'    a = cbrta_newtonf(a, d);
'    return cbrta_newtonf(a, d);
'}
Function newton_cbrt3f(ByVal d As Single) As Single
    Dim a  As Single: a = cbrt_5f(d)
    a = cbrta_newtonf(a, d)
    a = cbrta_newtonf(a, d)
    newton_cbrt3f = cbrta_newtonf(a, d)
End Function

'
'// cube root approximation using 4 iterations of Newton's method (float)
'float newton_cbrt4f(float d)
'{
'    float a = cbrt_5f(d);
'    a = cbrta_newtonf(a, d);
'    a = cbrta_newtonf(a, d);
'    a = cbrta_newtonf(a, d);
'    return cbrta_newtonf(a, d);
'}

Function newton_cbrt4f(ByVal d As Single) As Single
    Dim a As Single: a = cbrt_5f(d)
    a = cbrta_newtonf(a, d)
    a = cbrta_newtonf(a, d)
    a = cbrta_newtonf(a, d)
    newton_cbrt4f = cbrta_newtonf(a, d)
End Function
