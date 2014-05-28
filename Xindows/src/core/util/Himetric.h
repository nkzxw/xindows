
#ifndef __XINDOWS_CORE_UTIL_HIMETRIC_H__
#define __XINDOWS_CORE_UTIL_HIMETRIC_H__

//+---------------------------------------------------------------------
//
//  Routines to convert Pixels to Himetric and vice versa
//
//----------------------------------------------------------------------
#define HIMETRIC_PER_INCH           2540L
#define POINT_PER_INCH              72L
#define TWIPS_PER_POINT             20L
#define TWIPS_PER_INCH              (POINT_PER_INCH*TWIPS_PER_POINT)
#define TWIPS_FROM_POINTS(points)   ((TWIPS_PER_INCH*(points))/POINT_PER_INCH)

long    HimetricFromHPix(int iPix);
long    HimetricFromVPix(int iPix);
int     HPixFromHimetric(long lHi);
int     VPixFromHimetric(long lHi);
float   UserFromHimetric(long lValue);
long    HimetricFromUser(float flt);

#endif //__XINDOWS_CORE_UTIL_HIMETRIC_H__