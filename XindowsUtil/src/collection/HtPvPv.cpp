
#include "stdafx.h"
#include "HtPvPv.h"

#define HtKeyEqual(pvKey1, pvKey2)  (((void*)((DWORD_PTR)pvKey1&~1L)) == (pvKey2))
#define HtKeyInUse(pvKey)           ((DWORD_PTR)pvKey > 1L)
#define HtKeyTstFree(pvKey)         (pvKey == NULL)
#define HtKeyTstBridged(pvKey)      ((DWORD_PTR)pvKey & 1L)
#define HtKeySetBridged(pvKey)      ((void*)((DWORD_PTR)pvKey | 1L))
#define HtKeyClrBridged(pvKey)      ((void*)((DWORD_PTR)pvKey & ~1L))
#define HtKeyTstRehash(pvKey)       ((DWORD_PTR)pvKey & 2L)
#define HtKeySetRehash(pvKey)       ((void*)((DWORD_PTR)pvKey | 2L))
#define HtKeyClrRehash(pvKey)       ((void*)((DWORD_PTR)pvKey & ~2L))
#define HtKeyTstFlags(pvKey)        ((DWORD_PTR)pvKey & 3L)
#define HtKeyClrFlags(pvKey)        ((void*)((DWORD_PTR)pvKey & ~3L))

CHtPvPv::CHtPvPv()
{
    _pEnt        = &_EntEmpty;
    _pEntLast    = &_EntEmpty;
    _cEntMax     = 1;
    _cStrideMask = 1;
}

CHtPvPv::~CHtPvPv()
{
    if(_pEnt != &_EntEmpty)
    {
        MemFree(_pEnt);
    }
}

UINT CHtPvPv::ComputeStrideMask(UINT cEntMax)
{
    UINT iMask;
    for(iMask=1; iMask<cEntMax; iMask<<=1);
    return ((iMask>>1)-1);
}

HRESULT CHtPvPv::Grow()
{
    HRESULT hr;
    DWORD*  pdw;
    UINT    cEntMax;
    UINT    cEntGrow;
    UINT    cEntShrink;
    HTENT*  pEnt;

    extern DWORD s_asizeAssoc[];

    for(pdw=s_asizeAssoc; *pdw<=_cEntMax; pdw++) ;

    cEntMax    = *pdw;
    cEntGrow   = cEntMax * 8L / 10L;
    cEntShrink = (pdw>s_asizeAssoc) ? *(pdw-1)*4L/10L : 0;
    pEnt       = (_pEnt==&_EntEmpty) ? NULL : _pEnt;

    hr = MemRealloc((void**)&pEnt, cEntMax*sizeof(HTENT));

    if(hr == S_OK)
    {
        _pEnt           = pEnt;
        _pEntLast       = &_EntEmpty;
        _cEntGrow       = cEntGrow;
        _cEntShrink     = cEntShrink;
        _cStrideMask    = ComputeStrideMask(cEntMax);

        memset(&_pEnt[_cEntMax], 0, (cEntMax-_cEntMax)*sizeof(HTENT));

        if(_cEntMax == 1)
        {
            memset(_pEnt, 0, sizeof(HTENT));
        }

        Rehash(cEntMax);
    }

    Trace((_T("Growing to cEntMax=%ld (cEntShrink=%ld,cEnt=%ld,cEntGrow=%ld)\n"),
        _cEntMax, _cEntShrink, _cEnt, _cEntGrow));

    RRETURN(hr);
}

extern DWORD s_asizeAssoc[];
void CHtPvPv::Shrink()
{
    DWORD*  pdw;
    UINT    cEntMax;
    UINT    cEntGrow;
    UINT    cEntShrink;

    for(pdw=s_asizeAssoc; *pdw<_cEntMax; pdw++) ;

    cEntMax    = *--pdw;
    cEntGrow   = cEntMax * 8L / 10L;
    cEntShrink = (pdw>s_asizeAssoc) ? *(pdw-1)*4L/10L : 0;

    Assert(_cEnt < cEntGrow);
    Assert(_cEnt > cEntShrink);

    _pEntLast       = &_EntEmpty;
    _cEntGrow       = cEntGrow;
    _cEntShrink     = cEntShrink;
    _cStrideMask    = ComputeStrideMask(cEntMax);

    Rehash(cEntMax);

    Verify(MemRealloc((void**)&_pEnt, cEntMax*sizeof(HTENT)) == S_OK);

    Trace((_T("Shrinking to cEntMax=%ld (cEntShrink=%ld,cEnt=%ld,cEntGrow=%ld)\n"),
        _cEntMax, _cEntShrink, _cEnt, _cEntGrow));
}

void CHtPvPv::Rehash(UINT cEntMax)
{
    UINT    iEntScan    = 0;
    UINT    cEntScan    = _cEntMax;
    HTENT*  pEntScan    = _pEnt;
    UINT    iEnt;
    UINT    cEnt;
    HTENT*  pEnt;

    _cEntDel = 0;
    _cEntMax = cEntMax;

    for(; iEntScan<cEntScan; ++iEntScan,++pEntScan)
    {
        if(HtKeyInUse(pEntScan->pvKey))
        {
            pEntScan->pvKey = HtKeyClrBridged(HtKeySetRehash(pEntScan->pvKey));
        }
        else
        {
            pEntScan->pvKey = NULL;
        }
        Assert(!HtKeyTstBridged(pEntScan->pvKey));
    }

    iEntScan = 0;
    pEntScan = _pEnt;

    for(; iEntScan<cEntScan; ++iEntScan,++pEntScan)
    {
repeat:
        if(HtKeyTstRehash(pEntScan->pvKey))
        {
            pEntScan->pvKey = HtKeyClrRehash(pEntScan->pvKey);

            iEnt = ComputeProbe(pEntScan->pvKey);
            cEnt = ComputeStride(pEntScan->pvKey);

            for(;;)
            {
                pEnt = &_pEnt[iEnt];

                if(pEnt == pEntScan)
                {
                    break;
                }

                if(pEnt->pvKey == NULL)
                {
                    *pEnt = *pEntScan;
                    pEntScan->pvKey = NULL;
                    break;
                }

                if(HtKeyTstRehash(pEnt->pvKey))
                {
                    void* pvKey1 = HtKeyClrBridged(pEnt->pvKey);
                    void* pvKey2 = HtKeyClrBridged(pEntScan->pvKey);
                    void* pvVal1 = pEnt->pvVal;
                    void* pvVal2 = pEntScan->pvVal;

                    if(HtKeyTstBridged(pEntScan->pvKey))
                    {
                        pvKey1 = HtKeySetBridged(pvKey1);
                    }

                    if(HtKeyTstBridged(pEnt->pvKey))
                    {
                        pvKey2 = HtKeySetBridged(pvKey2);
                    }

                    pEntScan->pvKey = pvKey1;
                    pEntScan->pvVal = pvVal1;
                    pEnt->pvKey = pvKey2;
                    pEnt->pvVal = pvVal2;

                    goto repeat;
                }

                pEnt->pvKey = HtKeySetBridged(pEnt->pvKey);

                iEnt += cEnt;

                if(iEnt >= _cEntMax)
                {
                    iEnt -= _cEntMax;
                }
            }
        }
    }
}

void* CHtPvPv::Lookup(void* pvKey)
{
    HTENT*  pEnt;
    UINT    iEnt;
    UINT    cEnt;

    Assert(!HtKeyTstFree(pvKey) && !HtKeyTstFlags(pvKey));

    if(HtKeyEqual(_pEntLast->pvKey, pvKey))
    {
        return _pEntLast->pvVal;
    }

    iEnt = ComputeProbe(pvKey);
    pEnt = &_pEnt[iEnt];

    if(HtKeyEqual(pEnt->pvKey, pvKey))
    {
        _pEntLast = pEnt;
        return(pEnt->pvVal);
    }

    if(!HtKeyTstBridged(pEnt->pvKey))
    {
        return(NULL);
    }

    cEnt = ComputeStride(pvKey);

    for(;;)
    {
        iEnt += cEnt;

        if(iEnt >= _cEntMax)
        {
            iEnt -= _cEntMax;
        }

        pEnt = &_pEnt[iEnt];

        if(HtKeyEqual(pEnt->pvKey, pvKey))
        {
            _pEntLast = pEnt;
            return(pEnt->pvVal);
        }

        if(!HtKeyTstBridged(pEnt->pvKey))
        {
            return(NULL);
        }
    }
}

#ifdef _DEBUG
BOOL CHtPvPv::IsPresent(void* pvKey)
{
    HTENT*  pEnt;
    UINT    iEnt;
    UINT    cEnt;

    Assert(!HtKeyTstFree(pvKey) && !HtKeyTstFlags(pvKey));

    iEnt = ComputeProbe(pvKey);
    pEnt = &_pEnt[iEnt];

    if(HtKeyEqual(pEnt->pvKey, pvKey))
    {
        return TRUE;
    }

    if(!HtKeyTstBridged(pEnt->pvKey))
    {
        return FALSE;
    }

    cEnt = ComputeStride(pvKey);

    for(;;)
    {
        iEnt += cEnt;

        if(iEnt >= _cEntMax)
        {
            iEnt -= _cEntMax;
        }

        pEnt = &_pEnt[iEnt];

        if(HtKeyEqual(pEnt->pvKey, pvKey))
        {
            return TRUE;
        }

        if(!HtKeyTstBridged(pEnt->pvKey))
        {
            return FALSE;
        }
    }
}
#endif

HRESULT CHtPvPv::Insert(void* pvKey, void* pvVal)
{
    HTENT*  pEnt;
    UINT    iEnt;
    UINT    cEnt;

    Assert(!HtKeyTstFree(pvKey) && !HtKeyTstFlags(pvKey));
    Assert(!IsPresent(pvKey));

    if(_cEnt+_cEntDel >= _cEntGrow)
    {
        if(_cEntDel > (_cEnt>>2))
        {
            Rehash(_cEntMax);
        }
        else
        {
            HRESULT hr = Grow();

            if(hr)
            {
                RRETURN(hr);
            }
        }
    }

    iEnt = ComputeProbe(pvKey);
    pEnt = &_pEnt[iEnt];

    if(!HtKeyInUse(pEnt->pvKey))
    {
        goto insert;
    }

    pEnt->pvKey = HtKeySetBridged(pEnt->pvKey);

    cEnt = ComputeStride(pvKey);

    for(;;)
    {
        iEnt += cEnt;

        if(iEnt >= _cEntMax)
        {
            iEnt -= _cEntMax;
        }

        pEnt = &_pEnt[iEnt];

        if(!HtKeyInUse(pEnt->pvKey))
        {
            goto insert;
        }

        pEnt->pvKey = HtKeySetBridged(pEnt->pvKey);
    }

insert:
    if(HtKeyTstBridged(pEnt->pvKey))
    {
        _cEntDel -= 1;
        pEnt->pvKey = HtKeySetBridged(pvKey);
    }
    else
    {
        pEnt->pvKey = pvKey;
    }

    pEnt->pvVal = pvVal;

    _pEntLast = pEnt;

    _cEnt += 1;

    return S_OK;
}

void* CHtPvPv::Remove(void* pvKey)
{
    HTENT*  pEnt;
    UINT    iEnt;
    UINT    cEnt;

    Assert(!HtKeyTstFree(pvKey) && !HtKeyTstFlags(pvKey));

    iEnt = ComputeProbe(pvKey);
    pEnt = &_pEnt[iEnt];

    if(HtKeyEqual(pEnt->pvKey, pvKey))
    {
        goto remove;
    }

    if(!HtKeyTstBridged(pEnt->pvKey))
    {
        return(NULL);
    }

    cEnt = ComputeStride(pvKey);

    for(;;)
    {
        iEnt += cEnt;

        if(iEnt >= _cEntMax)
        {
            iEnt -= _cEntMax;
        }

        pEnt = &_pEnt[iEnt];

        if(HtKeyEqual(pEnt->pvKey, pvKey))
        {
            goto remove;
        }

        if(!HtKeyTstBridged(pEnt->pvKey))
        {
            return(NULL);
        }
    }

remove:
    if(HtKeyTstBridged(pEnt->pvKey))
    {
        pEnt->pvKey = HtKeySetBridged(NULL);
        _cEntDel += 1;
    }
    else
    {
        pEnt->pvKey = NULL;
    }

    pvKey = pEnt->pvVal;

    _cEnt -= 1;

    if(_cEnt < _cEntShrink)
    {
        Shrink();
    }

    return pvKey;
}