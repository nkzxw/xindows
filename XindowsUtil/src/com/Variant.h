
#ifndef __XINDOWSUTIL_COM_VARIANT_H__
#define __XINDOWSUTIL_COM_VARIANT_H__

class XINDOWS_PUBLIC CVariant : public VARIANT
{
public:
    CVariant()           { ZeroVariant(); }
    CVariant(VARTYPE vt) { ZeroVariant(); V_VT(this) = vt; }
    ~CVariant()          { Clear(); }

    void ZeroVariant()
    {
        memset(this, 0, sizeof(CVariant));
    }

    HRESULT Copy(VARIANT* pVar)
    {
        Assert(pVar);
        return VariantCopy(this, pVar);
    }

    HRESULT Clear()
    {
        return VariantClear(this);
    }

    // Coerce from an arbitrary variant into this. (Copes with VATIANT BYREF's within VARIANTS).
    HRESULT CoerceVariantArg(VARIANT* pArgFrom, WORD wCoerceToType);
    // Coerce current variant into itself
    HRESULT CoerceVariantArg(WORD wCoerceToType);
    BOOL CoerceNumericToI4();
    BOOL IsEmpty()    { return (vt==VT_NULL || vt==VT_EMPTY); }
    BOOL IsOptional() { return (IsEmpty()   || vt==VT_ERROR); }
};

XINDOWS_PUBLIC HRESULT VariantChangeTypeSpecial(VARIANT* pvargDest, VARIANT* pVArg, VARTYPE vt, IServiceProvider* pSrvProvider=NULL, DWORD dwFlags=0);
XINDOWS_PUBLIC HRESULT VARIANTARGToCVar(VARIANT* pvarg, BOOL* pfAlloc, VARTYPE vt, void* pv, IServiceProvider* pSrvProvider=NULL, WORD wMaxstrlen=0);
XINDOWS_PUBLIC HRESULT VARIANTARGToIndex(VARIANT* pvarg, long* plIndex);
XINDOWS_PUBLIC void    CVarToVARIANTARG(void* pv, VARTYPE vt, VARIANTARG* pvarg);

#endif //__XINDOWSUTIL_COM_VARIANT_H__