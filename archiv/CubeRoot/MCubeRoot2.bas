Attribute VB_Name = "MCubeRoot2"
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
    Dim t As TDouble: t.Value = d
    Dim p As TLong2:   LSet p = t: p.Value1 = p.Value1 \ 3 + 715094163
    Dim a As Double:   LSet t = p: a = t.Value
    Dim a3 As Double: a3 = a * a * a
    halley_cbrt1d = a * (a3 + d + d) / (a3 + a3 + d)
End Function

'
'// cube root approximation using 1 iteration of Halley's method (float)
'float halley_cbrt1f(float d)
'{
'    float a = cbrt_5f(d);
'    return cbrta_halleyf(a, d);
'}
Function halley_cbrt1f(ByVal f As Single) As Single
    Dim a3 As Single
    Dim t As TSingle: t.Value = f
    Dim p As TLong:    LSet p = t: p.Value = p.Value \ 3 + 709921077
    Dim a As Single:   LSet t = p: a = t.Value
    a3 = a * a * a
    halley_cbrt1f = a * (a3 + f + f) / (a3 + a3 + f)
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
    
    Dim a3 As Double
    Dim t As TDouble: t.Value = d
    Dim p As TLong2:   LSet p = t: p.Value1 = p.Value1 \ 3 + 715094163
    Dim a As Double:   LSet t = p: a = t.Value
    a3 = a * a * a
    a = a * (a3 + d + d) / (a3 + a3 + d)
    a3 = a * a * a
    halley_cbrt2d = a * (a3 + d + d) / (a3 + a3 + d)
    
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
    
    Dim a3 As Double
    Dim t As TDouble: t.Value = d
    Dim p As TLong2:   LSet p = t: p.Value1 = p.Value1 \ 3 + 715094163
    Dim a As Double:   LSet t = p: a = t.Value
    a3 = a * a * a
    a = a * (a3 + d + d) / (a3 + a3 + d)
    a3 = a * a * a
    a = a * (a3 + d + d) / (a3 + a3 + d)
    a3 = a * a * a
    halley_cbrt3d = a * (a3 + d + d) / (a3 + a3 + d)
    
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
Function halley_cbrt2f(ByVal f As Single) As Single
    
    Dim a3 As Single
    Dim t As TSingle: t.Value = f
    Dim p As TLong:    LSet p = t: p.Value = p.Value \ 3 + 709921077
    Dim a As Single:   LSet t = p: a = t.Value
    a3 = a * a * a
    a = a * (a3 + f + f) / (a3 + a3 + f)
    a3 = a * a * a
    halley_cbrt2f = a * (a3 + f + f) / (a3 + a3 + f)
    
End Function

'
'// cube root approximation using 1 iteration of Newton's method (double)
'double newton_cbrt1d(double d)
'{
'    double a = cbrt_5d(d);
'    return cbrta_newtond(a, d);
'}
Function newton_cbrt1d(ByVal d As Double) As Double
    
    Dim t As TDouble: t.Value = d
    Dim p As TLong2:   LSet p = t: p.Value1 = p.Value1 \ 3 + 715094163
    Dim a As Double:   LSet t = p: a = t.Value
    newton_cbrt1d = (1# / 3#) * (d / (a * a) + 2 * a)
    
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
    
    Dim t As TDouble: t.Value = d
    Dim p As TLong2:   LSet p = t: p.Value1 = p.Value1 \ 3 + 715094163
    Dim a As Double:   LSet t = p: a = t.Value
    a = (1# / 3#) * (d / (a * a) + 2 * a)
    newton_cbrt2d = (1# / 3#) * (d / (a * a) + 2 * a)
    
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
    
    Dim t As TDouble: t.Value = d
    Dim p As TLong2:   LSet p = t: p.Value1 = p.Value1 \ 3 + 715094163
    Dim a As Double:   LSet t = p: a = t.Value
    a = (1# / 3#) * (d / (a * a) + 2 * a)
    a = (1# / 3#) * (d / (a * a) + 2 * a)
    newton_cbrt3d = (1# / 3#) * (d / (a * a) + 2 * a)
    
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
    
    Dim t As TDouble: t.Value = d
    Dim p As TLong2:   LSet p = t: p.Value1 = p.Value1 \ 3 + 715094163
    Dim a As Double:   LSet t = p: a = t.Value
    a = (1# / 3#) * (d / (a * a) + 2 * a)
    a = (1# / 3#) * (d / (a * a) + 2 * a)
    a = (1# / 3#) * (d / (a * a) + 2 * a)
    newton_cbrt4d = (1# / 3#) * (d / (a * a) + 2 * a)
    
End Function

'
'// cube root approximation using 2 iterations of Newton's method (float)
'float newton_cbrt1f(float d)
'{
'    float a = cbrt_5f(d);
'    return cbrta_newtonf(a, d);
'}
Function newton_cbrt1f(ByVal d As Single) As Single
    
    Dim t As TSingle: t.Value = d
    Dim p As TLong:    LSet p = t: p.Value = p.Value \ 3 + 709921077
    Dim a As Single:   LSet t = p: a = t.Value
    a = t.Value
    newton_cbrt1f = a - (1! / 3!) * (a - d / (a * a))
    
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
    
    Dim t As TSingle: t.Value = d
    Dim p As TLong:    LSet p = t: p.Value = p.Value \ 3 + 709921077
    Dim a As Single:   LSet t = p: a = t.Value
    a = t.Value
    a = a - (1! / 3!) * (a - d / (a * a))
    newton_cbrt2f = a - (1! / 3!) * (a - d / (a * a))
    
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
    
    Dim t As TSingle: t.Value = d
    Dim p As TLong:    LSet p = t: p.Value = p.Value \ 3 + 709921077
    Dim a As Single:   LSet t = p: a = t.Value
    a = t.Value
    a = a - (1! / 3!) * (a - d / (a * a))
    a = a - (1! / 3!) * (a - d / (a * a))
    newton_cbrt3f = a - (1! / 3!) * (a - d / (a * a))
    
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
    
    Dim t As TSingle: t.Value = d
    Dim p As TLong:    LSet p = t: p.Value = p.Value \ 3 + 709921077
    Dim a As Single:   LSet t = p: a = t.Value
    a = a - (1! / 3!) * (a - d / (a * a))
    a = a - (1! / 3!) * (a - d / (a * a))
    a = a - (1! / 3!) * (a - d / (a * a))
    newton_cbrt4f = a - (1! / 3!) * (a - d / (a * a))
    
End Function
