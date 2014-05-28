
#ifndef __XINDOWSUTIL_GLOBAL_MATH_H__
#define __XINDOWSUTIL_GLOBAL_MATH_H__

inline BOOL InRange(const TCHAR ch, const TCHAR chMin, const TCHAR chMax)
{
    return (unsigned)(ch-chMin)<=(unsigned)(chMax-chMin);
}

inline int MulDivQuick(int nMultiplicand, int nMultiplier, int nDivisor)
{
    Assert(nDivisor);
    return (!nDivisor-1)&MulDiv(nMultiplicand, nMultiplier, nDivisor);
}

#endif //__XINDOWSUTIL_GLOBAL_MATH_H__