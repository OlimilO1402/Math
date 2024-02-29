// CubeRoot.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <windows.h>
#include <math.h>

// get accurate timer (Win32)
double GetTimer()
{
	LARGE_INTEGER F, N;
	QueryPerformanceFrequency(&F);
	QueryPerformanceCounter(&N);
	return double(N.QuadPart)/double(F.QuadPart);
}

typedef float  (*cuberootfnf) (float);
typedef double (*cuberootfnd) (double);

// estimate bits of precision (32-bit float case)
inline int bits_of_precision(float a, float b)
{
	const double kd = 1.0 / log(2.0);

	if (a==b)
		return 23;

	const double kdmin = pow(2.0, -23.0);

	double d = fabs(a-b);
	if (d < kdmin)
		return 23;

	return int(-log(d)*kd);
}

// estiamte bits of precision (64-bit double case)
inline int bits_of_precision(double a, double b)
{
	const double kd = 1.0 / log(2.0);

	if (a==b)
		return 52;

	const double kdmin = pow(2.0, -52.0);

	double d = fabs(a-b);
	if (d < kdmin)
		return 52;

	return int(-log(d)*kd);
}

// cube root via x^(1/3)
float pow_cbrtf(float x)
{
	return pow(x, 1.0f/3.0f);
}

// cube root via x^(1/3)
double pow_cbrtd(double x)
{
	return pow(x, 1.0/3.0);
}

// cube root approximation using bit hack for 32-bit float
__forceinline float cbrt_5f(float f)
{
	unsigned int* p = (unsigned int *) &f;
	*p = *p/3 + 709921077;
	return f;
}

// cube root approximation using bit hack for 64-bit float 
// adapted from Kahan's cbrt
__forceinline double cbrt_5d(double d)
{
	const unsigned int B1 = 715094163;
	double t = 0.0;
	unsigned int* pt = (unsigned int*) &t;
	unsigned int* px = (unsigned int*) &d;
	pt[1]=px[1]/3+B1;
	return t;
}

// cube root approximation using bit hack for 64-bit float 
// adapted from Kahan's cbrt
__forceinline double quint_5d(double d)
{
	return sqrt(sqrt(d));

	const unsigned int B1 = 71509416*5/3;
	double t = 0.0;
	unsigned int* pt = (unsigned int*) &t;
	unsigned int* px = (unsigned int*) &d;
	pt[1]=px[1]/5+B1;
	return t;
}

// iterative cube root approximation using Halley's method (float)
__forceinline float cbrta_halleyf(const float a, const float R)
{
	const float a3 = a*a*a;
    const float b= a * (a3 + R + R) / (a3 + a3 + R);
	return b;
}

// iterative cube root approximation using Halley's method (double)
__forceinline double cbrta_halleyd(const double a, const double R)
{
	const double a3 = a*a*a;
    const double b= a * (a3 + R + R) / (a3 + a3 + R);
	return b;
}

// iterative cube root approximation using Newton's method (float)
__forceinline float cbrta_newtonf(const float a, const float x)
{
//    return (1.0 / 3.0) * ((a + a) + x / (a * a));
	return a - (1.0f / 3.0f) * (a - x / (a*a));
}

// iterative cube root approximation using Newton's method (double)
__forceinline double cbrta_newtond(const double a, const double x)
{
	return (1.0/3.0) * (x / (a*a) + 2*a);
}

// cube root approximation using 1 iteration of Halley's method (double)
double halley_cbrt1d(double d)
{
	double a = cbrt_5d(d);
	return cbrta_halleyd(a, d);
}

// cube root approximation using 1 iteration of Halley's method (float)
float halley_cbrt1f(float d)
{
	float a = cbrt_5f(d);
	return cbrta_halleyf(a, d);
}

// cube root approximation using 2 iterations of Halley's method (double)
double halley_cbrt2d(double d)
{
	double a = cbrt_5d(d);
	a = cbrta_halleyd(a, d);
	return cbrta_halleyd(a, d);
}

// cube root approximation using 3 iterations of Halley's method (double)
double halley_cbrt3d(double d)
{
	double a = cbrt_5d(d);
	a = cbrta_halleyd(a, d);
	a = cbrta_halleyd(a, d);
	return cbrta_halleyd(a, d);
}


// cube root approximation using 2 iterations of Halley's method (float)
float halley_cbrt2f(float d)
{
	float a = cbrt_5f(d);
	a = cbrta_halleyf(a, d);
	return cbrta_halleyf(a, d);
}

// cube root approximation using 1 iteration of Newton's method (double)
double newton_cbrt1d(double d)
{
	double a = cbrt_5d(d);
	return cbrta_newtond(a, d);
}

// cube root approximation using 2 iterations of Newton's method (double)
double newton_cbrt2d(double d)
{
	double a = cbrt_5d(d);
	a = cbrta_newtond(a, d);
	return cbrta_newtond(a, d);
}

// cube root approximation using 3 iterations of Newton's method (double)
double newton_cbrt3d(double d)
{
	double a = cbrt_5d(d);
	a = cbrta_newtond(a, d);
	a = cbrta_newtond(a, d);
	return cbrta_newtond(a, d);
}

// cube root approximation using 4 iterations of Newton's method (double)
double newton_cbrt4d(double d)
{
	double a = cbrt_5d(d);
	a = cbrta_newtond(a, d);
	a = cbrta_newtond(a, d);
	a = cbrta_newtond(a, d);
	return cbrta_newtond(a, d);
}

// cube root approximation using 2 iterations of Newton's method (float)
float newton_cbrt1f(float d)
{
	float a = cbrt_5f(d);
	return cbrta_newtonf(a, d);
}

// cube root approximation using 2 iterations of Newton's method (float)
float newton_cbrt2f(float d)
{
	float a = cbrt_5f(d);
	a = cbrta_newtonf(a, d);
	return cbrta_newtonf(a, d);
}

// cube root approximation using 3 iterations of Newton's method (float)
float newton_cbrt3f(float d)
{
	float a = cbrt_5f(d);
	a = cbrta_newtonf(a, d);
	a = cbrta_newtonf(a, d);
	return cbrta_newtonf(a, d);
}

// cube root approximation using 4 iterations of Newton's method (float)
float newton_cbrt4f(float d)
{
	float a = cbrt_5f(d);
	a = cbrta_newtonf(a, d);
	a = cbrta_newtonf(a, d);
	a = cbrta_newtonf(a, d);
	return cbrta_newtonf(a, d);
}

double TestCubeRootf(const char* szName, cuberootfnf cbrt, double rA, double rB, int rN)
{
	const int N = rN;
 	
	float dd = float((rB-rA) / N);

	// calculate 1M numbers
	int i=0;
	float d = (float) rA;

	double t = GetTimer();
	double s = 0.0;

	for(d=(float) rA, i=0; i<N; i++, d += dd)
	{
		s += cbrt(d);
	}

	t = GetTimer() - t;

	printf("%-10s %5.1f ms ", szName, t*1000.0);

	double maxre = 0.0;
	double bits = 0.0;
	double worstx=0.0;
	double worsty=0.0;
	int minbits=64;

	for(d=(float) rA, i=0; i<N; i++, d += dd)
	{
		float a = cbrt((float) d);	
		float b = (float) pow((double) d, 1.0/3.0);

		int bc = bits_of_precision(a, b);
		bits += bc;

		if (b > 1.0e-6)
		{
			if (bc < minbits)
			{
				minbits = bc;
				worstx = d;
				worsty = a;
			}
		}
	}

	bits /= N;

    printf(" %3d mbp  %6.3f abp\n", minbits, bits);

	return s;
}


double TestCubeRootd(const char* szName, cuberootfnd cbrt, double rA, double rB, int rN)
{
	const int N = rN;
	
	double dd = (rB-rA) / N;

	int i=0;

	double t = GetTimer();
	
	double s = 0.0;
	double d = 0.0;

	for(d=rA, i=0; i<N; i++, d += dd)
	{
		s += cbrt(d);
	}

	t = GetTimer() - t;

	printf("%-10s %5.1f ms ", szName, t*1000.0);

	double bits = 0.0;
	double maxre = 0.0;	
	double worstx = 0.0;
	double worsty = 0.0;
	int minbits = 64;
	for(d=rA, i=0; i<N; i++, d += dd)
	{
		double a = cbrt(d);	
		double b = pow(d, 1.0/3.0);

		int bc = bits_of_precision(a, b); // min(53, count_matching_bitsd(a, b) - 12);
		bits += bc;

		if (b > 1.0e-6)
		{
			if (bc < minbits)
			{
				bits_of_precision(a, b);
				minbits = bc;
				worstx = d;
				worsty = a;
			}
		}
	}

	bits /= N;

    printf(" %3d mbp  %6.3f abp\n", minbits, bits);

	return s;
}

int _tmain(int argc, _TCHAR* argv[])
{
	// a million uniform steps through the range from 0.0 to 1.0
	// (doing uniform steps in the log scale would be better)
	double a = 0.0;
	double b = 1.0;
	int n = 1000000;

	printf("32-bit float tests\n");
	printf("----------------------------------------\n");
	TestCubeRootf("cbrt_5f", cbrt_5f, a, b, n);
	TestCubeRootf("pow", pow_cbrtf, a, b, n);
	TestCubeRootf("halley x 1", halley_cbrt1f, a, b, n);
	TestCubeRootf("halley x 2", halley_cbrt2f, a, b, n);
	TestCubeRootf("newton x 1", newton_cbrt1f, a, b, n);
	TestCubeRootf("newton x 2", newton_cbrt2f, a, b, n);
	TestCubeRootf("newton x 3", newton_cbrt3f, a, b, n);
	TestCubeRootf("newton x 4", newton_cbrt4f, a, b, n);
	printf("\n\n");

	printf("64-bit double tests\n");
	printf("----------------------------------------\n");
	TestCubeRootd("cbrt_5d", cbrt_5d, a, b, n);
	TestCubeRootd("pow", pow_cbrtd, a, b, n);
	TestCubeRootd("halley x 1", halley_cbrt1d, a, b, n);
	TestCubeRootd("halley x 2", halley_cbrt2d, a, b, n);
	TestCubeRootd("halley x 3", halley_cbrt3d, a, b, n);
	TestCubeRootd("newton x 1", newton_cbrt1d, a, b, n);
	TestCubeRootd("newton x 2", newton_cbrt2d, a, b, n);
	TestCubeRootd("newton x 3", newton_cbrt3d, a, b, n);
	TestCubeRootd("newton x 4", newton_cbrt4d, a, b, n);
	printf("\n\n");

	getchar();

	return 0;
}