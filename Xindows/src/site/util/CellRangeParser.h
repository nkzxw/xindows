
#ifndef __XINDOWS_SITE_UTIL_CELLRANGEPARSER_H__
#define __XINDOWS_SITE_UTIL_CELLRANGEPARSER_H__

class CCellRangeParser
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();

    CCellRangeParser(LPCTSTR szRange);

    void    GetRangeRect(RECT* pRect) { *pRect = _RangeRect; }
    BOOL    Failed(void) { return _fFailed;}
    LPCTSTR GetNormalizedRangeName() { return _strNormString; }
private:
    RECT     _RangeRect;
    BOOL     _fFailed;
    BOOL     _fOneValue;
    CString  _strNormString;

    void Normalize(LPCTSTR);
    BOOL GetColumn(int* pnCurIndex, long* pnCol);
    BOOL GetNumber(int* pnCurIndex, long* pnRow);
};

#endif //__XINDOWS_SITE_UTIL_CELLRANGEPARSER_H__