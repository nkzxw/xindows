
#include "stdafx.h"
#include "teststring.h"

XINDOWS_PUBLIC void TestString()
{
    CString str;
    str.Set(_T("ÎÒ²Ý"));

    CString* pStr = new CString();
    pStr->Set(_T("Wo cao"));

    CString* pStr2 = NULL;
    pStr->Clone(&pStr2);

    delete pStr;
    delete pStr2;

    CBuffer sb;
    sb.Append(_T("Hello, "));
    sb.Append(_T("World!"));

    CBuffer2 sb2;
    sb2.Append(_T("Hello, "));
    sb2.Append(_T("World!"));

    CDataListEnumerator de(sb);
    LPCTSTR lpNext;
    int nLen;
    while(de.GetNext(&lpNext, &nLen))
    {
        Trace((lpNext));
    }
    Trace0("\n");

    CBufferedString* pBS = new CBufferedString;
    delete pBS;
}