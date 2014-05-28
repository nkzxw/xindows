
#ifndef __XINDOWS_SITE_UTIL_FONTINFO_H__
#define __XINDOWS_SITE_UTIL_FONTINFO_H__

class CFontInfo
{
public:
    CString _cstrFaceName;
    SCRIPT_IDS _sids;
};

class CFontInfoCache : public CDataAry<CFontInfo>
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()
    CFontInfoCache() : CDataAry<CFontInfo>() {}
    void Free();

    HRESULT AddInfoToAtomTable(LPCTSTR pch, LONG* plIndex);
    HRESULT GetAtomFromName(LPCTSTR pch, LONG* plIndex);
    HRESULT GetInfoFromAtom(long lIndex, CFontInfo** ppfi);
    HRESULT SetScriptIDsOnAtom(LONG lIndex, SCRIPT_IDS sids)
    {
        CFontInfo* pfi;
        HRESULT hr = GetInfoFromAtom(lIndex, &pfi);
        if(!hr)
        {
            pfi->_sids = sids;
        }
        RRETURN(hr);
    }
    HRESULT AddScriptIDOnAtom(LONG lIndex, SCRIPT_ID sid)
    {
        CFontInfo* pfi;
        HRESULT hr = GetInfoFromAtom(lIndex, &pfi);
        if(!hr)
        {
            pfi->_sids |= ScriptBit(sid);
        }
        RRETURN(hr);
    }
};

#endif //__XINDOWS_SITE_UTIL_FONTINFO_H__