
#include "stdafx.h"
#include "style.h"

#define _cxx_
#include "../gen/style.hdl"

// {783466AA-52BB-4d2a-A83B-409D3AFFF450}
static const GUID CLSID_CStyle = { 0x783466aa, 0x52bb, 0x4d2a, { 0xa8, 0x3b, 0x40, 0x9d, 0x3a, 0xff, 0xf4, 0x50 } };

//+------------------------------------------------------------------------
//
//  Member:     CStyle::CStyle
//
//-------------------------------------------------------------------------
CStyle::CStyle(CElement* pElem, DISPID dispID, DWORD dwFlags)
{
    _pElem = pElem;
    _dispIDAA = dispID;

    if(pElem)
    {
        if(dispID != DISPID_UNKNOWN)
        {
            _pAA = *pElem->CreateStyleAttrArray(_dispIDAA);
        }
    }

    ClearFlag(STYLE_MASKPROPERTYCHANGES);
    SetFlag(dwFlags);
    if(TestFlag(STYLE_SEPARATEFROMELEM))
    {
        Assert(_pElem);
        _pElem->SubAddRef();
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CStyle::~CStyle
//
//-------------------------------------------------------------------------
CStyle::~CStyle()
{
    if(!TestFlag(STYLE_SEPARATEFROMELEM))
    {
        Passivate();
    }
}

void CStyle::Passivate()
{
    if(_pStyleSource)
    {
        _pStyleSource->Release();
        _pStyleSource = NULL;
    }

    if(TestFlag(STYLE_SEPARATEFROMELEM))
    {
        _pElem->SubRelease();
        _pElem = NULL;
    }
    else
    {
        // Don't leave this alone, or CBase::Passivate will try to manage its destruction.
        // Since it doesn't actually belong to us (it belongs to its entry in _pElem->_pAA),
        // this would cause problems.
        _pAA = NULL;
    }
    super::Passivate();
}


ULONG CStyle::PrivateAddRef(void)
{
    Assert(_pElem);
    if(!TestFlag(STYLE_SEPARATEFROMELEM))
    {
        _pElem->AddRef();
    }
    else
    {
        super::PrivateAddRef();
    }

    return 0;
}

ULONG CStyle::PrivateRelease(void)
{
    Assert(_pElem);
    if(!TestFlag(STYLE_SEPARATEFROMELEM))
    {
        _pElem->Release();
    }
    else
    {
        super::PrivateRelease();
    }

    return 0;
}

const CStyle::CLASSDESC CStyle::s_classdesc =
{
    {
        &CLSID_HTMLStyle,                   // _pclsid
        NULL,                               // _pcpi
        0,                                  // _dwFlags
        &IID_IHTMLStyle,                    // _piidDispinterface
        &s_apHdlDescs,                      // _apHdlDesc
    },
    (void*)s_apfnpdIHTMLStyle,              // _apfnTearOff
};

BEGIN_TEAROFF_TABLE(CStyle, IServiceProvider)
    TEAROFF_METHOD(CStyle, &QueryService, queryservice, (REFGUID guidService, REFIID riid, void** ppvObject))
END_TEAROFF_TABLE()

CAtomTable* CStyle::GetAtomTable(BOOL* pfExpando)
{
    return _pElem->GetAtomTable(pfExpando);
}

//+------------------------------------------------------------------------
//
//  Member:     CStyle::PrivateQueryInterface, IUnknown
//
//  Synopsis:   Private unknown QI.
//
//-------------------------------------------------------------------------
HRESULT CStyle::PrivateQueryInterface(REFIID iid, void** ppv)
{
    HRESULT hr = S_OK;
    // All interfaces derived from IDispatch must be handled
    // using the ElementDesc()->_apfnTearOff tearoff interface.
    // This allows classes such as COleSite to override the
    // implementation of IDispatch methods.
    *ppv = NULL;
    switch(iid.Data1)
    {
        QI_INHERITS((IPrivateUnknown*)this, IUnknown)
        QI_TEAROFF(this, IObjectIdentity, NULL)
        QI_HTML_TEAROFF(this, IHTMLStyle2, NULL)
        QI_TEAROFF(this, IServiceProvider, NULL)
        QI_CASE(IConnectionPointContainer)
    default:
        {
            const CLASSDESC* pclassdesc = (CLASSDESC*)BaseDesc();

            if(iid == CLSID_CStyle)
            {
                *ppv = this; // Weak ref
                return S_OK;
            }

            if(pclassdesc && pclassdesc->_apfnTearOff &&
                pclassdesc->_classdescBase._piidDispinterface &&
                (iid==IID_IHTMLStyle || DispNonDualDIID(iid)))
            {
                hr = CreateTearOffThunk(this,
                    (void*)(pclassdesc->_apfnTearOff),
                    NULL,
                    ppv,
                    (void*)(CStyle::s_ppropdescsInVtblOrderIHTMLStyle));
            }
        }
    }
    if(!hr)
    {
        if(*ppv)
        {
            (*(IUnknown**)ppv)->AddRef();
        }
        else
        {
            hr = E_NOINTERFACE;
        }
    }
    RRETURN(hr);
}

void CStyle::MaskPropertyChanges(BOOL fMask)
{ 
    if(fMask)
    {
        SetFlag(STYLE_MASKPROPERTYCHANGES);
    }
    else
    {
        ClearFlag(STYLE_MASKPROPERTYCHANGES);
    }
}

//+------------------------------------------------------------------------
//
//  Member:     OnPropertyChange
//
//  Note:       Called after a property has changed to notify derived classes
//              of the change.  All properties (except those managed by the
//              derived class) are consistent before this call.
//
//              Also, fires a property change notification to the site.
//
//-------------------------------------------------------------------------
HRESULT CStyle::OnPropertyChange(DISPID dispid, DWORD dwFlags)
{
    HRESULT hr = S_OK;

    if(TestFlag(STYLE_MASKPROPERTYCHANGES))
    {
        goto Cleanup;
    }

    if(_pElem)
    {
        if(dispid == DISPID_A_POSITION)
        {
            CDocument* pDoc = _pElem->Doc();

            if(!pDoc->NeedRegionCollection())
            {
                DWORD dwVal;
                CAttrArray* pAA = *GetAttrArray();
                Assert(pAA);
                BOOL fFound = pAA->FindSimple(DISPID_A_POSITION, &dwVal);

                if(fFound && ((stylePosition)dwVal==stylePositionrelative || 
                    (stylePosition)dwVal==stylePositionabsolute))
                {
                    pDoc->_fRegionCollection = TRUE;
                }
            }
        }

        // allow Element to make a decision based on distinguishing the same DISPID from
        // element vs in-line style, by OR'ing in the ELEMCHNG_INLINESTYLE_PROPERTY
        if(_dispIDAA == DISPID_INTERNAL_INLINESTYLEAA)
        {
            dwFlags |= ELEMCHNG_INLINESTYLE_PROPERTY;
        }

        if(TestFlag(STYLE_NOCLEARCACHES))
        {
            // BUGBUG: (anandra) HACK ALERT.  This is completely a workaround
            // for the fact that CElement::OnPropertyChange turns around and calls 
            // ComputeFormat.  This causes recursion for the behavior onstyleapply stuff.
            // By not clearing the formats here, we prevent this from happening.  
            // CPeerHolder::ApplyStyleMulti will clear the caches correctly.
            dwFlags &= ~(ELEMCHNG_CLEARCACHES|ELEMCHNG_CLEARFF);
        }

        hr = _pElem->OnPropertyChange(dispid, dwFlags);
    }

Cleanup:
    RRETURN(hr);
}

STDMETHODIMP CStyle::QueryService(REFGUID guidService, REFIID riid, void** ppvObject)
{
    Assert(_pElem);
    RRETURN(SetErrorInfo(_pElem->QueryService(guidService, riid, ppvObject)));
}

CAttrArray** CStyle::GetAttrArray(void) const
{
    CAttrArray** ppAA;

    if(!TestFlag(STYLE_SEPARATEFROMELEM))
    {
        return _pElem->CreateStyleAttrArray(_dispIDAA);
    }

    ppAA = const_cast<CAttrArray**>(&_pAA);
    if(*ppAA)
    {
        return ppAA;
    }

    *ppAA = new CAttrArray;
    return ppAA;
}

//+---------------------------------------------------------------------------
//
//  Helper Function:    InvokeSourceGet
//
//----------------------------------------------------------------------------
HRESULT InvokeSourceGet(IDispatch* pdisp, DISPID dispid, VARTYPE varTypeResultRequested, void* pv)
{
    return E_FAIL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CStyle::NeedToDelegateGet
//
//----------------------------------------------------------------------------
BOOL CStyle::NeedToDelegateGet(DISPID dispid)
{
    CAttrArray* pAA = *GetAttrArray();
    return (_pStyleSource &&
        AA_IDX_UNKNOWN==pAA->FindAAIndex(dispid, CAttrValue::AA_Attribute));
}

//+---------------------------------------------------------------------------
//
//  Member:     CStyle::DelegateGet
//
//----------------------------------------------------------------------------
HRESULT CStyle::DelegateGet(DISPID dispid, VARTYPE varType, void* pv)
{
    return InvokeSourceGet(_pStyleSource, dispid, varType, pv);
}

//+---------------------------------------------------------------------------
//
//  Member:     CStyle::getValueHelper, public
//
//  Synopsis:   Helper function to implement get_ top/left/width/height
//
//  Arguments:  [plValue] -- Place to put value
//              [dwFlags] -- Flags indicating whether its an X or Y attribute
//              [puv]     -- CUnitValue containing the value we want to return
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
enum putValueFlags
{
    PUTVF_X     = 0x1,
    PUTVF_SIZE  = 0x2,
    PUTVF_Y     = 0x4,
    PUTVF_POS   = 0x8
};

HRESULT CStyle::getValueHelper(long* plValue, DWORD dwFlags, const PROPERTYDESC* pPropertyDesc)
{
    HRESULT     hr = S_OK;
    RECT        rcParent;
    CUnitValue  uv;
    CUnitValue  uvTemp;
    DWORD       dwVal;
    CLayout*    pLayoutParent;
    long        lParentSize;

    if(!plValue)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    if(_pElem)
    {    
        hr = _pElem->EnsureInMarkup();
        if(hr)
        {
            goto Cleanup;
        }

        {
            CAttrArray::FindSimple(
                *GetAttrArray(), pPropertyDesc, &dwVal);
            uvTemp.SetRawValue(dwVal);
        }
    }

    if(!uvTemp.IsNull() && _pElem)
    {
        CDocInfo DCI;

        pLayoutParent = _pElem->GetFirstBranch()->Parent()->GetUpdatedNearestLayout();
        if(!pLayoutParent && _pElem->GetMarkup())
        {
            CElement* pElement = _pElem->GetMarkup()->GetElementClient();

            if(pElement)
            {
                pLayoutParent = pElement->GetUpdatedLayout();
            }
        }

        if(pLayoutParent || _pElem->GetMarkup())
        {
            DCI.Init(pLayoutParent ? pLayoutParent->ElementOwner() : _pElem->GetMarkup()->Root());
        }
        // else, the element is not in a tree. e.g. el = new Image();
        if(pLayoutParent)
        {
            pLayoutParent->GetClientRect(&rcParent);

            if(dwFlags & PUTVF_X)
            {
                lParentSize = rcParent.right - rcParent.left;
            }
            else
            {
                lParentSize = rcParent.bottom - rcParent.top;
            }
        }
        else
        {
            // if there is no parent layout and no Element client layout
            // rcParent is empty, therefore lParentSize is NULL
            lParentSize = 0;
        }

        if(dwFlags & PUTVF_X)
        {
            *plValue = uvTemp.XGetPixelValue(&DCI, lParentSize, 
                _pElem->GetFirstBranch()->GetFontHeightInTwips(&uv));
        }
        else
        {
            *plValue = uvTemp.YGetPixelValue(&DCI, lParentSize, 
                _pElem->GetFirstBranch()->GetFontHeightInTwips(&uv));
        }
    }
    else
    {
        *plValue = 0;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
//  Member:     CStyle::putValueHelper, public
//
//  Synopsis:   Helper function for implementing put_ top/left/width/height
//
//  Arguments:  [lValue]  -- New value to store
//              [dwFlags] -- Flags indicating direction and what it affects
//              [dispid]  -- DISPID of property
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------

// BUGBUG (carled) there is a compiler bug that manifest with this function
//  do NOT reorder the local variables below or else dwFlags will incorrectly
//  get reset to 0 and the site postion
HRESULT CStyle::putValueHelper(long lValue, DWORD dwFlags, DISPID dispid)
{
    BOOL        fChanged    = FALSE;
    RECT        rcParent;
    long        delta;
    DWORD       dwPropFlags = ELEMCHNG_CLEARCACHES;
    HRESULT     hr          = S_OK;
    CLayout*    pParentLayout;

    if(!_pElem)
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

    pParentLayout = _pElem->GetUpdatedParentLayout();

    if(dwFlags & PUTVF_SIZE)
    {
        dwPropFlags |= ELEMCHNG_SIZECHANGED;
    }
    else
    {
        dwPropFlags |= ELEMCHNG_SITEPOSITION;
    }

    if(_dispIDAA == DISPID_INTERNAL_INLINESTYLEAA)
    {
        dwPropFlags |= ELEMCHNG_INLINESTYLE_PROPERTY;
    }

// wlw note
//#ifndef NO_EDIT
//    {
//        CUndoPropChangeNotificationPlaceHolder notfholder(TRUE, _pElem, dispid, dwPropFlags);
//#endif // NO_EDIT

        if(!pParentLayout && _pElem->GetMarkup())
        {
            CElement* pElement = _pElem->GetMarkup()->GetElementClient();
            // if there is no parent layout.
            // we are either a body ot framesetsite.
            // we should get client element rect
            if(pElement)
            {
                pParentLayout = pElement->GetUpdatedLayout();
            }
        }

        if(pParentLayout)
        {
            pParentLayout->GetClientRect(&rcParent);

            if(dwFlags & PUTVF_X)
            {
                delta = rcParent.right - rcParent.left;
            }
            else
            {
                delta = rcParent.bottom - rcParent.top;
            }
        }
        else
        {
            // if there is no parent layout and no Element client layout
            // rcParent is empty, therefore delta is NULL
            delta = 0;
        }

        hr = _pElem->SetDim(
            dispid,
            (float)lValue,
            CUnitValue::UNIT_PIXELS,
            delta,
            GetAttrArray(),
            _dispIDAA==DISPID_INTERNAL_INLINESTYLEAA,
            &fChanged);
        if(hr)
        {
            goto Cleanup;
        }

        if(fChanged)
        {
            _pElem->OnPropertyChange(dispid, dwPropFlags );
        }

// wlw note
//#ifndef NO_EDIT
//        notfholder.SetHR(fChanged?hr:S_FALSE);
//    }
//#endif // NO_EDIT

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CStyle::get_pixelWidth(long* plValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_pixelWidth))
    {
        return DelegateGet(DISPID_CStyle_pixelWidth, VT_I4, plValue);
    }

    RRETURN(getValueHelper(plValue, PUTVF_X, &s_propdescCStylewidth.a));
}

HRESULT CStyle::put_pixelWidth(long lValue)
{
    RRETURN(putValueHelper(lValue, PUTVF_X|PUTVF_SIZE, STDPROPID_XOBJ_WIDTH));
}

HRESULT CStyle::get_pixelHeight(long* plValue)
{

    if(NeedToDelegateGet(DISPID_CStyle_pixelHeight))
    {
        return DelegateGet(DISPID_CStyle_pixelHeight, VT_I4, plValue);
    }

    RRETURN(getValueHelper(plValue, PUTVF_Y, &s_propdescCStyleheight.a));
}

HRESULT CStyle::put_pixelHeight(long lValue)
{
    RRETURN(putValueHelper(lValue, PUTVF_Y|PUTVF_SIZE, STDPROPID_XOBJ_HEIGHT));
}

HRESULT CStyle::get_pixelLeft(long* plValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_pixelLeft))
    {
        return DelegateGet(DISPID_CStyle_pixelLeft, VT_I4, plValue);
    }

    RRETURN(getValueHelper(plValue, PUTVF_X, &s_propdescCStyleleft.a));
}

HRESULT CStyle::put_pixelLeft(long lValue)
{
    RRETURN(putValueHelper(lValue, PUTVF_X|PUTVF_POS, STDPROPID_XOBJ_LEFT));
}

HRESULT CStyle::get_pixelRight(long* plValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_pixelRight))
    {
        return DelegateGet(DISPID_CStyle_pixelRight, VT_I4, plValue);
    }

    RRETURN(getValueHelper(plValue, PUTVF_X, &s_propdescCStyleright.a));
}

HRESULT CStyle::put_pixelRight(long lValue)
{
    RRETURN(putValueHelper(lValue, PUTVF_X|PUTVF_POS, STDPROPID_XOBJ_RIGHT));
}

HRESULT CStyle::get_pixelTop(long* plValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_pixelTop))
    {
        return DelegateGet(DISPID_CStyle_pixelTop, VT_I4, plValue);
    }

    RRETURN(getValueHelper(plValue, PUTVF_Y, &s_propdescCStyletop.a));
}

HRESULT CStyle::put_pixelTop(long lValue)
{
    RRETURN(putValueHelper(lValue, PUTVF_Y|PUTVF_POS, STDPROPID_XOBJ_TOP));
}

HRESULT CStyle::get_pixelBottom(long* plValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_pixelBottom))
    {
        return DelegateGet(DISPID_CStyle_pixelBottom, VT_I4, plValue);
    }

    RRETURN(getValueHelper(plValue, PUTVF_Y, &s_propdescCStylebottom.a));
}

HRESULT CStyle::put_pixelBottom(long lValue)
{
    RRETURN(putValueHelper(lValue, PUTVF_Y|PUTVF_POS, STDPROPID_XOBJ_BOTTOM));
}

HRESULT CStyle::getfloatHelper(float* pfValue, DWORD dwFlags, const PROPERTYDESC* pPropertyDesc)
{
    HRESULT                   hr = S_OK;
    CUnitValue::UNITVALUETYPE uvt;
    CUnitValue                uv;
    CUnitValue                uvTemp;
    DWORD                     dwVal;

    if(!pfValue)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    if(_pElem)
    {
        hr = _pElem->EnsureInMarkup();
        if(hr)
        {
            goto Cleanup;
        }

        {
            CAttrArray::FindSimple(
                *GetAttrArray(), pPropertyDesc, &dwVal);
            uvTemp.SetRawValue(dwVal);
        }
    }

    if(!uvTemp.IsNull() && _pElem)
    {
        uvt = uvTemp.GetUnitType();

        if(dwFlags & PUTVF_X)
        {
            *pfValue = uvTemp.XGetFloatValueInUnits(
                uvt, _pElem->GetFirstBranch()->GetFontHeightInTwips(&uv));
        }
        else
        {
            *pfValue = uvTemp.YGetFloatValueInUnits(
                uvt, _pElem->GetFirstBranch()->GetFontHeightInTwips(&uv));
        }
    }
    else
    {
        *pfValue = 0;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CStyle::putfloatHelper(float fValue, DWORD dwFlags, DISPID dispid)
{
    HRESULT hr          = S_OK;
    DWORD   dwPropFlags = ELEMCHNG_CLEARCACHES;
    BOOL    fChanged    = FALSE;
    RECT    rcParent;
    long    delta;

    if(!_pElem)
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

    if(!_pElem->GetUpdatedParentLayout())
    {
        delta = 0;
    }
    else
    {
        _pElem->GetUpdatedParentLayout()->GetClientRect(&rcParent);

        if(dwFlags & PUTVF_X)
        {
            delta = rcParent.right - rcParent.left;
        }
        else
        {
            delta = rcParent.bottom - rcParent.top;
        }
    }

    hr = _pElem->SetDim(
        dispid,
        fValue,
        CUnitValue::UNIT_NULLVALUE,
        delta,
        GetAttrArray(),
        _dispIDAA==DISPID_INTERNAL_INLINESTYLEAA,
        &fChanged);
    if(hr)
    {
        goto Cleanup;
    }

    if(fChanged)
    {
        if(dwFlags & PUTVF_SIZE)
        {
            dwPropFlags |= ELEMCHNG_SIZECHANGED;
        }
        else
        {
            dwPropFlags |= ELEMCHNG_SITEPOSITION;
        }

        if(_dispIDAA == DISPID_INTERNAL_INLINESTYLEAA)
        {
            dwPropFlags |= ELEMCHNG_INLINESTYLE_PROPERTY;
        }

        _pElem->OnPropertyChange(dispid, dwPropFlags);
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CStyle::get_posWidth(float* pfValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_posWidth))
    {
        Assert(FALSE && "Not implemented");
        return E_NOTIMPL;
    }

    RRETURN(getfloatHelper(pfValue, PUTVF_X, &s_propdescCStylewidth.a));
}

HRESULT CStyle::put_posWidth(float fValue)
{
    RRETURN(putfloatHelper(fValue, PUTVF_X|PUTVF_SIZE, STDPROPID_XOBJ_WIDTH));
}

HRESULT CStyle::get_posHeight(float* pfValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_posHeight))
    {
        Assert(FALSE && "Not implemented");
        return E_NOTIMPL;
    }

    RRETURN(getfloatHelper(pfValue, PUTVF_Y, &s_propdescCStyleheight.a));
}

HRESULT CStyle::put_posHeight(float fValue)
{
    RRETURN(putfloatHelper(fValue, PUTVF_Y|PUTVF_SIZE, STDPROPID_XOBJ_HEIGHT));
}

HRESULT CStyle::get_posLeft(float* pfValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_posLeft))
    {
        Assert (FALSE && "Not implemented");
        return E_NOTIMPL;
    }

    RRETURN(getfloatHelper(pfValue, PUTVF_X, &s_propdescCStyleleft.a));
}

HRESULT CStyle::put_posLeft(float fValue)
{
    RRETURN(putfloatHelper(fValue, PUTVF_X|PUTVF_POS, STDPROPID_XOBJ_LEFT));
}

HRESULT CStyle::get_posRight(float* pfValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_posRight))
    {
        Assert (FALSE && "Not implemented");
        return E_NOTIMPL;
    }

    RRETURN(getfloatHelper(pfValue, PUTVF_X, &s_propdescCStyleright.a));
}

HRESULT CStyle::put_posRight(float fValue)
{
    RRETURN(putfloatHelper(fValue, PUTVF_X|PUTVF_POS, STDPROPID_XOBJ_RIGHT));
}

HRESULT CStyle::get_posTop(float* pfValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_posTop))
    {
        Assert (FALSE && "Not implemented");
        return E_NOTIMPL;
    }

    RRETURN(getfloatHelper(pfValue, PUTVF_Y, &s_propdescCStyletop.a));
}

HRESULT CStyle::put_posTop(float fValue)
{
    RRETURN(putfloatHelper(fValue, PUTVF_Y|PUTVF_POS, STDPROPID_XOBJ_TOP));
}

HRESULT CStyle::get_posBottom(float* pfValue)
{
    if(NeedToDelegateGet(DISPID_CStyle_posBottom))
    {
        Assert (FALSE && "Not implemented");
        return E_NOTIMPL;
    }

    RRETURN(getfloatHelper(pfValue, PUTVF_Y, &s_propdescCStylebottom.a));
}

HRESULT CStyle::put_posBottom(float fValue)
{
    RRETURN(putfloatHelper(fValue, PUTVF_Y|PUTVF_POS, STDPROPID_XOBJ_BOTTOM));
}

// All putters/getters must not have a pointer into the element attrArray is it could move.
// Use the below macros to guarantee we're pointing to a local variable which is pointing to the
// style sheet attrArray and not pointing to the attrValue on the element attrArray which can
// move if the elements attrArray has attrValues added to or deleted from.
#define GETATTR_ARRAY   CAttrArray *pTempStyleAA = *GetAttrArray();
#define USEATTR_ARRAY   &pTempStyleAA

STDMETHODIMP CStyle::put_StyleComponent(BSTR v)
{
    GET_THUNK_PROPDESC
    GETATTR_ARRAY
    return put_StyleComponentHelper(v, pPropDesc, USEATTR_ARRAY);
}

STDMETHODIMP CStyle::put_Url(BSTR v)
{
    GET_THUNK_PROPDESC
    GETATTR_ARRAY
    return put_UrlHelper(v, pPropDesc, USEATTR_ARRAY);
}

STDMETHODIMP CStyle::put_String(BSTR v)
{
    GET_THUNK_PROPDESC
    GETATTR_ARRAY
    return put_StringHelper(v, pPropDesc, USEATTR_ARRAY);
}

STDMETHODIMP CStyle::put_Long(long v)
{
    GET_THUNK_PROPDESC
    GETATTR_ARRAY
    return put_LongHelper(v, pPropDesc, USEATTR_ARRAY);
}

STDMETHODIMP CStyle::put_Bool(VARIANT_BOOL v)
{
    GET_THUNK_PROPDESC
    GETATTR_ARRAY
    return put_BoolHelper(v, pPropDesc, USEATTR_ARRAY);
}

STDMETHODIMP CStyle::put_Variant(VARIANT var)
{
    GET_THUNK_PROPDESC
    GETATTR_ARRAY

    return put_VariantHelper(var, pPropDesc, USEATTR_ARRAY);
}

STDMETHODIMP CStyle::get_Url(BSTR* pbstr)
{
    GET_THUNK_PROPDESC
    GETATTR_ARRAY

    if (NeedToDelegateGet(pPropDesc->GetDispid()))
    {
        return DelegateGet(pPropDesc->GetDispid(), VT_BSTR, pbstr);
    }

    return get_UrlHelper(pbstr, pPropDesc, USEATTR_ARRAY);
}

STDMETHODIMP CStyle::get_StyleComponent(BSTR* pbstr)
{
    GET_THUNK_PROPDESC
    GETATTR_ARRAY

    if (NeedToDelegateGet(pPropDesc->GetDispid()))
    {
        return DelegateGet(pPropDesc->GetDispid(), VT_BSTR, pbstr);
    }

    return get_StyleComponentHelper(pbstr, pPropDesc, USEATTR_ARRAY);
}

STDMETHODIMP CStyle::get_Property(void* pv)
{
    GET_THUNK_PROPDESC
    GETATTR_ARRAY

    if(NeedToDelegateGet(pPropDesc->GetDispid()))
    {
        return DelegateGet(pPropDesc->GetDispid(), VT_VARIANT, pv);
    }

    return get_PropertyHelper(pv, pPropDesc, USEATTR_ARRAY);
}

STDMETHODIMP CStyle::put_textDecorationNone(VARIANT_BOOL v)
{
    return s_propdescCStyletextDecorationNone.b.SetStyleComponentBooleanProperty(v, this, (CVoid *)(void *)(GetAttrArray()));
}
STDMETHODIMP CStyle::get_textDecorationNone(VARIANT_BOOL* p)
{
    return s_propdescCStyletextDecorationNone.b.GetStyleComponentBooleanProperty(p, this, (CVoid *)(void *)(GetAttrArray()));
}
STDMETHODIMP CStyle::put_textDecorationUnderline(VARIANT_BOOL v)
{
    return s_propdescCStyletextDecorationUnderline.b.SetStyleComponentBooleanProperty(v, this, (CVoid *)(void *)(GetAttrArray()));
}
STDMETHODIMP CStyle::get_textDecorationUnderline(VARIANT_BOOL* p)
{
    return s_propdescCStyletextDecorationUnderline.b.GetStyleComponentBooleanProperty(p, this, (CVoid *)(void *)(GetAttrArray()));
}
STDMETHODIMP CStyle::put_textDecorationOverline(VARIANT_BOOL v)
{
    return s_propdescCStyletextDecorationOverline.b.SetStyleComponentBooleanProperty(v, this, (CVoid *)(void *)(GetAttrArray()));
}
STDMETHODIMP CStyle::get_textDecorationOverline(VARIANT_BOOL* p)
{
    return s_propdescCStyletextDecorationOverline.b.GetStyleComponentBooleanProperty(p, this, (CVoid *)(void *)(GetAttrArray()));
}
STDMETHODIMP CStyle::put_textDecorationLineThrough(VARIANT_BOOL v)
{
    return s_propdescCStyletextDecorationLineThrough.b.SetStyleComponentBooleanProperty(v, this, (CVoid *)(void *)(GetAttrArray()));
}
STDMETHODIMP CStyle::get_textDecorationLineThrough(VARIANT_BOOL* p)
{
    return s_propdescCStyletextDecorationLineThrough.b.GetStyleComponentBooleanProperty(p, this, (CVoid *)(void *)(GetAttrArray()));
}
STDMETHODIMP CStyle::put_textDecorationBlink(VARIANT_BOOL v)
{
    return s_propdescCStyletextDecorationBlink.b.SetStyleComponentBooleanProperty(v, this, (CVoid *)(void *)(GetAttrArray()));
}
STDMETHODIMP CStyle::get_textDecorationBlink(VARIANT_BOOL* p)
{
    return s_propdescCStyletextDecorationBlink.b.GetStyleComponentBooleanProperty(p, this, (CVoid *)(void *)(GetAttrArray()));
}

HRESULT CStyle::toString(BSTR* String)
{
    RRETURN(super::toString(String));
}


typedef struct
{
    const TCHAR* szName;
    Esysfont eType;
} SystemFontName;

const SystemFontName asSystemFontNames[] =
{
    { _T("caption"), sysfont_caption },
    { _T("icon"), sysfont_icon },
    { _T("menu"), sysfont_menu },
    { _T("messagebox"), sysfont_messagebox },
    { _T("smallcaption"), sysfont_smallcaption },
    { _T("statusbar"), sysfont_statusbar }
};

int CompareSysFontPairsByName(const void* pv1, const void* pv2)
{
    return StrCmpIC(((SystemFontName*)pv1)->szName, ((SystemFontName*)pv2)->szName);
}

Esysfont FindSystemFontByName(const TCHAR* szString)
{
    SystemFontName sPotentialSysFont;

    sPotentialSysFont.szName = szString;

    SystemFontName* pFont = (SystemFontName*)bsearch(
        &sPotentialSysFont,
        asSystemFontNames,
        ARRAYSIZE(asSystemFontNames),
        sizeof(SystemFontName),
        CompareSysFontPairsByName);
    if(pFont)
    {
        return pFont->eType;
    }
    return sysfont_non_system;
}

//+------------------------------------------------------------------------
//
//  Helper Function:    NextSize
//
//  Synopsis:
//      This function tokenizes a font-size; we can't just tokenize on space,
//  in case we have a string like "12 pt".  This function returns a pointer
//  to the first character that is not part of the first size specifier in the
//  string.  It may or may not skip whitespace following the size specifier.
//
//  Return Values:
//      NULL on error (no size at the head of this string)
//      "" if the end of the string was reached.
//      pointer to following characters if success.
//-------------------------------------------------------------------------
TCHAR* NextSize(TCHAR* pszSize, const NUMPROPPARAMS* ppp)
{
    TCHAR* pszChar = pszSize;
    TCHAR* pszLastSpace = NULL;

    Assert(pszSize != NULL);

    // Skip any leading whitespace
    while(_istspace(*pszChar))
    {
        pszChar++;
    }

    // it's okay for the first character to be '-' or '+'
    if((*pszChar==_T('-')) || (*pszChar==_T('+')))
    {
        pszChar++;
    }

    // if the first character (after any '-') is not a digit, check if it's an enumerated size, then error.
    if(!_istdigit(*pszChar) && (*pszChar!=_T('.')))
    {
        if(ppp -> bpp.dwPPFlags & PROPPARAM_ENUM)
        {
            TCHAR chTerm;
            TCHAR* pszToken = pszChar;
            long lEnum;

            while(_istalpha(*pszChar) || (*pszChar==_T('-')))
            {
                pszChar++;
            }

            chTerm = *pszChar;
            *pszChar = 0;

            HRESULT hr = LookupEnumString(ppp, pszToken, &lEnum);
            *pszChar = chTerm;
            if(hr == S_OK)
            {
                return pszChar;
            }
        }
        return NULL;
    }

    // Now we know we have at least one digit, so we can pull a size out of it.
    
    // Skip over all the digits
    while(_istdigit(*pszChar) || (*pszChar==_T('.')))
    {
        pszChar++;
    }

    //  Skip any whitespace between the last digit and the (potential) unit string
    pszLastSpace = pszChar;
    while(_istspace(*pszChar))
    {
        pszChar++;
    }

    //  If the string is "in", parse it as inches
    if((pszChar[0] && pszChar[1]) // Make sure at least two chars remain
        && (!_tcsnicmp(_T("in"), 2, pszChar, 2)
        ||  !_tcsnicmp(_T("cm"), 2, pszChar, 2)
        ||  !_tcsnicmp(_T("mm"), 2, pszChar, 2)
        ||  !_tcsnicmp(_T("em"), 2, pszChar, 2)
        ||  !_tcsnicmp(_T("ex"), 2, pszChar, 2)
        ||  !_tcsnicmp(_T("pt"), 2, pszChar, 2)
        ||  !_tcsnicmp(_T("pc"), 2, pszChar, 2)
        ||  !_tcsnicmp(_T("px"), 2, pszChar, 2)))
    {
        return (pszChar+2);
    }

    if(*pszChar == _T('%')) // If the string ends with '%', it's a percentage
    {
        return pszChar+1;
    }

    // Default is to treat the string as "px" and parse it as pixels
    return pszLastSpace;
}

//+------------------------------------------------------------------------
//
//  Helper Function:    ValidStyleUrl
//
//  Synopsis:
//      This function determines if a string is a valid CSS-style URL
//  functional notation string (e.g. "url(http://www.foo.com/bar)").
//
//  Return Values:
//      zero if invalid, otherwise the length of the string inside
//-------------------------------------------------------------------------
static const TCHAR strURLBeg[] = _T("url(");
size_t ValidStyleUrl(TCHAR* pch)
{
    size_t nLen = pch ? _tcslen(pch) : 0;
    if((nLen>=5) && (0==_tcsnicmp(pch, ARRAYSIZE(strURLBeg)-1, strURLBeg, ARRAYSIZE(strURLBeg)-1))
        && (_T(')')==pch[nLen-1]))
    {
        return nLen;
    }
    return 0;
}

#define SINGLEQUOTE     _T('\'')
#define DOUBLEQUOTE     _T('\"')
//+------------------------------------------------------------------------
//
//  Helper Function:    RemoveQuotes
//
//  Synopsis:
//      This function determines if a string is quoted and removes the matching
//  quotes from the string. If the quots do not match or there are no quotes
//  it returns S_FALSE
//-------------------------------------------------------------------------
HRESULT RemoveQuotes(LPTSTR* ppszStr)
{
    TCHAR   chQuote;
    LPTSTR  pstrProp;
    LPTSTR  pstrPropEnd;
    HRESULT hr = S_FALSE;

    // Remove the spaces before the quote (if there are spaces and a quote)
    while(_istspace(**ppszStr)) (*ppszStr)++;

    pstrProp = *ppszStr;

    // Skip the quote if it is present and remember the quote type
    if((*pstrProp!=DOUBLEQUOTE) && (*pstrProp!=SINGLEQUOTE))
    {
        // The string does not start with a quote, ignore it
        goto Cleanup;
    }

    chQuote = *pstrProp;
    pstrProp++;

    // Scan for a matching quote
    while(*pstrProp!=_T('\0') && *pstrProp!=chQuote)
    {
        pstrProp++;
    }

    if(*pstrProp == _T('\0'))
    {
        // No matching quote
        goto Cleanup;
    }

    // Save the ending quote position
    pstrPropEnd = pstrProp;

    pstrProp++;

    // Check to see if the quote is the last thing in the string 
    while(_istspace(*pstrProp)) pstrProp++;
    if(*pstrProp != _T('\0'))
    {
        goto Cleanup;
    }

    /// Remove the starting quote
    (*ppszStr)++;
    // Remove the ending quote
    *pstrPropEnd = _T('\0');
    hr = S_OK;

Cleanup:
    if(hr == S_FALSE)
    {
        // Remove the spaces at the end of the string
        int nUrlLen = _tcslen(*ppszStr);
        while(nUrlLen>0 && _istspace((*ppszStr)[--nUrlLen]))
        {
            (*ppszStr)[nUrlLen] = _T('\0');
        }
    }

    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Helper Function:    RemoveStyleUrlFromStr
//
//  Synopsis:
//      This function determines if a string is a valid CSS-style URL
//  functional notation string (e.g. "url(http://www.foo.com/bar)") and
//  returns the url string without url( and ) and surrounding quotes
//  !!!! This function modifies the parameter string pointer
//
//  Return Values:
//      S_OK if there was a url and S_FALSE it there was no URL string in front
//      
//-------------------------------------------------------------------------
HRESULT RemoveStyleUrlFromStr(TCHAR** ppszURL)
{
    HRESULT hr = S_OK;
    int     nUrlLen;

    // Remove the leading spaces
    while(_istspace(**ppszURL)) (*ppszURL)++;
    // and the trailing spaces
    nUrlLen = _tcslen(*ppszURL);
    while(nUrlLen>0  && _istspace((*ppszURL)[--nUrlLen]))
    {
        (*ppszURL)[nUrlLen] = _T('\0');
    }

    // Check if our URL is a CSS url(xxx) string.  If so, strip "url(" off front,
    // and ")" off back.  Otherwise assume it'a a straight URL.
    if(!ValidStyleUrl(*ppszURL))
    {
        RemoveQuotes(ppszURL);
        hr = S_FALSE;
        goto Cleanup;
    }

    // Skip the "url(" (-1 for the terminating 0)
    *ppszURL += ARRAYSIZE(strURLBeg) - 1;

    // Now cut from the end the closing )
    nUrlLen = _tcslen(*ppszURL);
    (*ppszURL)[--nUrlLen] = _T('\0');

    // Remove the leading spaces
    while(_istspace(**ppszURL)) { (*ppszURL)++; nUrlLen--; }
    // Now remove the trailing spaces
    while(nUrlLen>0 && _istspace((*ppszURL)[--nUrlLen]))
    {
        (*ppszURL)[nUrlLen] = _T('\0');
    }

    RemoveQuotes(ppszURL);

Cleanup:
    // Remove the leading spaces
    while(_istspace(**ppszURL)) { (*ppszURL)++; nUrlLen--;  }
    // Now remove the trailing spaces
    while(nUrlLen>0 && _istspace((*ppszURL)[--nUrlLen]))
    {
        (*ppszURL)[nUrlLen] = _T('\0');
    }

    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Helper Function:    NextParenthesizedToken
//
//  Synopsis:
//      This function tokenizes a word or functional notation.
//  (BUGBUG: does not handle quoted strings!)  Expects string to have no
//  leading whitespace.
//
//  Return Values:
//      "" if the end of the string was reached.
//      pointer to following characters if success.
//-------------------------------------------------------------------------
TCHAR* NextParenthesizedToken(TCHAR* pszToken)
{
    int iParenDepth = 0;
    BOOL fInQuotes = FALSE;
    TCHAR quotechar = _T('"');

    while(*pszToken && (iParenDepth || !_istspace(*pszToken)))
    {
        if(!fInQuotes)
        {
            if((*pszToken==_T('\'')) || (*pszToken==_T('"')))
            {
                fInQuotes = TRUE;
                quotechar = *pszToken;
            }
            else if(iParenDepth && (*pszToken==_T(')')))
            {
                iParenDepth--;
            }
            else if(*pszToken == _T('('))
            {
                iParenDepth++;
            }
        }
        else if(quotechar == *pszToken)
        {
            fInQuotes = FALSE;
        }
        pszToken++;
    }
    return pszToken;
}

//+------------------------------------------------------------------------
//
//  Function:   EscapeQuotes
//
//  Synopsis:   Changes all the double quotes in the string to single quotes
//              Future implementations of this function must escape the quotes,
//                  so the information is saved. They might also need to allocate
//                  and return another string.
//-------------------------------------------------------------------------
void EscapeQuotes(LPTSTR lpPropValue)
{
    if(!lpPropValue)
    {
        return;
    }

    while(*lpPropValue != 0)
    {
        if(*lpPropValue == DOUBLEQUOTE)
        {
            *lpPropValue = SINGLEQUOTE;
        }
        lpPropValue++;
    }
}

HRESULT GetExpandingPropdescs(
          DWORD             dwDispId,
          PROPERTYDESC**    pppdTop,
          PROPERTYDESC**    pppdRight,
          PROPERTYDESC**    pppdBottom,
          PROPERTYDESC**    pppdLeft)
{
    switch(dwDispId)
    {
    case DISPID_A_PADDING:
        *pppdTop    = (PROPERTYDESC*)&s_propdescCStylepaddingTop.a;
        *pppdRight  = (PROPERTYDESC*)&s_propdescCStylepaddingRight.a;
        *pppdBottom = (PROPERTYDESC*)&s_propdescCStylepaddingBottom.a;
        *pppdLeft   = (PROPERTYDESC*)&s_propdescCStylepaddingLeft.a;
        break;
    case DISPID_A_MARGIN:
        *pppdTop    = (PROPERTYDESC*)&s_propdescCStylemarginTop.a;
        *pppdRight  = (PROPERTYDESC*)&s_propdescCStylemarginRight.a;
        *pppdBottom = (PROPERTYDESC*)&s_propdescCStylemarginBottom.a;
        *pppdLeft   = (PROPERTYDESC*)&s_propdescCStylemarginLeft.a;
        break;
    case DISPID_A_BORDERCOLOR:
        *pppdTop    = (PROPERTYDESC*)&s_propdescCStyleborderTopColor.a;
        *pppdRight  = (PROPERTYDESC*)&s_propdescCStyleborderRightColor.a;
        *pppdBottom = (PROPERTYDESC*)&s_propdescCStyleborderBottomColor.a;
        *pppdLeft   = (PROPERTYDESC*)&s_propdescCStyleborderLeftColor.a;
        break;
    case DISPID_A_BORDERWIDTH:
        *pppdTop    = (PROPERTYDESC*)&s_propdescCStyleborderTopWidth.a;
        *pppdRight  = (PROPERTYDESC*)&s_propdescCStyleborderRightWidth.a;
        *pppdBottom = (PROPERTYDESC*)&s_propdescCStyleborderBottomWidth.a;
        *pppdLeft   = (PROPERTYDESC*)&s_propdescCStyleborderLeftWidth.a;
        break;
    case DISPID_A_BORDERSTYLE:
        *pppdTop    = (PROPERTYDESC*)&s_propdescCStyleborderTopStyle.a;
        *pppdRight  = (PROPERTYDESC*)&s_propdescCStyleborderRightStyle.a;
        *pppdBottom = (PROPERTYDESC*)&s_propdescCStyleborderBottomStyle.a;
        *pppdLeft   = (PROPERTYDESC*)&s_propdescCStyleborderLeftStyle.a;
        break;
    case DISPID_A_CLIP:
        *pppdTop    = (PROPERTYDESC*)&s_propdescCStyleclipTop.a;
        *pppdRight  = (PROPERTYDESC*)&s_propdescCStyleclipRight.a;
        *pppdBottom = (PROPERTYDESC*)&s_propdescCStyleclipBottom.a;
        *pppdLeft   = (PROPERTYDESC*)&s_propdescCStyleclipLeft.a;
        break;

    default:
        Assert("Unrecognized expansion property!" && FALSE);
        return S_FALSE;
    }
    return S_OK;
}

HRESULT GetBorderSidePropdescs(DWORD dwDispId, PROPERTYDESC** pppdStyle, PROPERTYDESC** pppdColor, PROPERTYDESC** pppdWidth)
{
    Assert("Must have PD pointers!" && pppdStyle && pppdColor && pppdWidth);

    PROPERTYDESC* borderpropdescs[3][4] =
    {
        {
            (PROPERTYDESC*)&s_propdescCStyleborderTopStyle.a,
            (PROPERTYDESC*)&s_propdescCStyleborderRightStyle.a,
            (PROPERTYDESC*)&s_propdescCStyleborderBottomStyle.a,
            (PROPERTYDESC*)&s_propdescCStyleborderLeftStyle.a
        },
        {
            (PROPERTYDESC*)&s_propdescCStyleborderTopColor.a,
            (PROPERTYDESC*)&s_propdescCStyleborderRightColor.a,
            (PROPERTYDESC*)&s_propdescCStyleborderBottomColor.a,
            (PROPERTYDESC*)&s_propdescCStyleborderLeftColor.a
        },
        {
            (PROPERTYDESC*)&s_propdescCStyleborderTopWidth.a,
            (PROPERTYDESC*)&s_propdescCStyleborderRightWidth.a,
            (PROPERTYDESC*)&s_propdescCStyleborderBottomWidth.a,
            (PROPERTYDESC*)&s_propdescCStyleborderLeftWidth.a
        }
    };

    switch(dwDispId)
    {
    case DISPID_A_BORDERTOP:
        *pppdStyle = borderpropdescs[0][0];
        *pppdColor = borderpropdescs[1][0];
        *pppdWidth = borderpropdescs[2][0];
        break;
    case DISPID_A_BORDERRIGHT:
        *pppdStyle = borderpropdescs[0][1];
        *pppdColor = borderpropdescs[1][1];
        *pppdWidth = borderpropdescs[2][1];
        break;
    case DISPID_A_BORDERBOTTOM:
        *pppdStyle = borderpropdescs[0][2];
        *pppdColor = borderpropdescs[1][2];
        *pppdWidth = borderpropdescs[2][2];
        break;
    case DISPID_A_BORDERLEFT:
        *pppdStyle = borderpropdescs[0][3];
        *pppdColor = borderpropdescs[1][3];
        *pppdWidth = borderpropdescs[2][3];
        break;
    default:
        Assert("Not a DISPID for a valid border property!");
        return S_FALSE;
    }
    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteStyleToBSTR
//
//  Synopsis:   Converts a collection of style properties to a BSTR for display on grid
//
//  Note:       Look at size member of GlobalAlloc on large style sheets
//
//-------------------------------------------------------------------------
HRESULT WriteStyleToBSTR(CBase* pObject, CAttrArray* pAA, BSTR* pbstr, BOOL fQuote)
{
    const PROPERTYDESC* const* pppropdescs;
    HRESULT             hr                          = S_OK;
    BASICPROPPARAMS*    pbpp;
    LPCTSTR             pchLastOut                  = NULL;
    const CAttrValue*   pAV;
    CString             cstrStyle;
    BSTR                bstr                        = NULL;
    BOOL                fUseCompositeBackground     = FALSE;
    BOOL                fUseCompositeBGPosition     = FALSE;
    BOOL                fUseCompositeFont           = FALSE;
    BOOL                fUseCompositeBorderTop      = FALSE;
    BOOL                fUseCompositeBorderRight    = FALSE;
    BOOL                fUseCompositeBorderBottom   = FALSE;
    BOOL                fUseCompositeBorderLeft     = FALSE;
    BOOL                fUseCompositeMargin         = FALSE;
    BOOL                fUseCompositeLayoutGrid     = FALSE;

    int                 iLen;
    int                 idx;
    LPCTSTR             lpPropName = NULL;
    LPCTSTR             lpPropValue;
    BSTR                bstrTemp;

    if(!pbstr)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pbstr = NULL;

    if(!(pAA && pAA->HasAnyAttribute(TRUE)))
    {
        goto Cleanup;
    }

    //Quote style string?
    if(fQuote)
    {
        cstrStyle.Append(_T("\""));
    }

    if(pAA->Find(DISPID_A_BACKGROUNDPOSX, CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_BACKGROUNDPOSY, CAttrValue::AA_Attribute))
    {
        fUseCompositeBGPosition = TRUE;
    }
    if(fUseCompositeBGPosition &&
        pAA->Find(DISPID_A_BACKGROUNDIMAGE , CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_BACKCOLOR, CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_BACKGROUNDREPEAT, CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_BACKGROUNDATTACHMENT, CAttrValue::AA_Attribute))
    {
        fUseCompositeBackground = TRUE;
    }

    if(pAA->Find(DISPID_A_FONTWEIGHT,   CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_FONTSTYLE,   CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_FONTVARIANT, CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_FONTSIZE,    CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_LINEHEIGHT,  CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_FONTFACE,    CAttrValue::AA_Attribute))
    {
        fUseCompositeFont = TRUE;
    }

    if(pAA->Find(DISPID_A_LAYOUTGRIDMODE,    CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_LAYOUTGRIDTYPE,   CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_LAYOUTGRIDLINE,   CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_LAYOUTGRIDCHAR,   CAttrValue::AA_Attribute))
    {
        fUseCompositeLayoutGrid = TRUE;
    }

    if(pAA->Find(DISPID_A_BORDERTOPSTYLE,   CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_BORDERTOPWIDTH,  CAttrValue::AA_Attribute))
    {
        fUseCompositeBorderTop = TRUE;
    }

    if(pAA->Find(DISPID_A_BORDERRIGHTSTYLE,   CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_BORDERRIGHTWIDTH,  CAttrValue::AA_Attribute))
    {
        fUseCompositeBorderRight = TRUE;
    }

    if(pAA->Find(DISPID_A_BORDERBOTTOMSTYLE,   CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_BORDERBOTTOMWIDTH,  CAttrValue::AA_Attribute))
    {
        fUseCompositeBorderBottom = TRUE;
    }

    if(pAA->Find(DISPID_A_BORDERLEFTSTYLE,   CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_BORDERLEFTWIDTH,  CAttrValue::AA_Attribute))
    {
        fUseCompositeBorderLeft = TRUE;
    }

    if(pAA->Find(DISPID_A_MARGINTOP,     CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_MARGINRIGHT,  CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_MARGINBOTTOM, CAttrValue::AA_Attribute) &&
        pAA->Find(DISPID_A_MARGINLEFT,   CAttrValue::AA_Attribute))
    {
        fUseCompositeMargin = TRUE;
    }

    for(pppropdescs=CStyle::s_apHdlDescs.ppPropDescs; *pppropdescs; pppropdescs++)
    {
        const PROPERTYDESC* ppropdesc = (*pppropdescs);
        DWORD dispid = ppropdesc->GetBasicPropParams()->dispid;

        pAV = NULL;

        // BUGBUG for now check for the method pointer because of old property implementation...
        if(!ppropdesc->pfnHandleProperty)
        {
            continue;
        }

        pbpp = (BASICPROPPARAMS*)((*pppropdescs) + 1);

        if((pbpp->dwPPFlags&PROPPARAM_NOPERSIST) || (!(pbpp->dwPPFlags&PROPPARAM_ATTRARRAY)))
        {
            continue;
        }

        switch(dispid)
        {
        case DISPID_A_BACKGROUNDPOSX:
        case DISPID_A_BACKGROUNDPOSY:
            if(fUseCompositeBGPosition)
            {
                continue;
            }
            break;
        case DISPID_A_BACKGROUNDIMAGE:
        case DISPID_BACKCOLOR:
        case DISPID_A_BACKGROUNDREPEAT:
        case DISPID_A_BACKGROUNDATTACHMENT:
            if(fUseCompositeBackground)
            {
                continue;
            }
            break;
        case DISPID_A_FONTWEIGHT:
        case DISPID_A_FONTSTYLE:
        case DISPID_A_FONTVARIANT:
        case DISPID_A_FONTSIZE:
        case DISPID_A_LINEHEIGHT:
        case DISPID_A_FONTFACE:
            if(fUseCompositeFont)
            {
                continue;
            }
            break;
        case DISPID_A_LAYOUTGRIDMODE:
        case DISPID_A_LAYOUTGRIDTYPE:
        case DISPID_A_LAYOUTGRIDLINE:
        case DISPID_A_LAYOUTGRIDCHAR:
            if(fUseCompositeLayoutGrid)
            {
                continue;
            }
            break;
        case DISPID_A_BORDERTOPSTYLE:
        case DISPID_A_BORDERTOPWIDTH:
        case DISPID_A_BORDERTOPCOLOR:
            if(fUseCompositeBorderTop)
            {
                continue;
            }
            break;
        case DISPID_A_BORDERRIGHTSTYLE:
        case DISPID_A_BORDERRIGHTWIDTH:
        case DISPID_A_BORDERRIGHTCOLOR:
            if(fUseCompositeBorderRight)
            {
                continue;
            }
            break;
        case DISPID_A_BORDERBOTTOMSTYLE:
        case DISPID_A_BORDERBOTTOMWIDTH:
        case DISPID_A_BORDERBOTTOMCOLOR:
            if(fUseCompositeBorderBottom)
            {
                continue;
            }
            break;
        case DISPID_A_BORDERLEFTSTYLE:
        case DISPID_A_BORDERLEFTWIDTH:
        case DISPID_A_BORDERLEFTCOLOR:
            if(fUseCompositeBorderLeft)
            {
                continue;
            }
            break;
        case DISPID_A_MARGINTOP:
        case DISPID_A_MARGINRIGHT:
        case DISPID_A_MARGINBOTTOM:
        case DISPID_A_MARGINLEFT:
            if(fUseCompositeMargin)
            {
                continue;
            }
            break;

        case DISPID_A_CLIPRECTTOP:
        case DISPID_A_CLIPRECTRIGHT:
        case DISPID_A_CLIPRECTBOTTOM:
        case DISPID_A_CLIPRECTLEFT:
            // We do not write out the components
            continue;

        case DISPID_A_CLIP:
            if(pAA->Find(DISPID_A_CLIPRECTTOP,     CAttrValue::AA_Attribute) ||
                pAA->Find(DISPID_A_CLIPRECTRIGHT,  CAttrValue::AA_Attribute) ||
                pAA->Find(DISPID_A_CLIPRECTBOTTOM, CAttrValue::AA_Attribute) ||
                pAA->Find(DISPID_A_CLIPRECTLEFT,   CAttrValue::AA_Attribute))
            {
                goto WriteOutName;
            }
            continue;

        case DISPID_A_BACKGROUNDPOSITION:
            if(fUseCompositeBGPosition && !fUseCompositeBackground)
            {
                goto WriteOutName;
            }
            continue;
        case DISPID_A_BACKGROUND:
            if(fUseCompositeBackground)
            {
                goto WriteOutName;
            }
            continue;
        case DISPID_A_FONT:
            if(fUseCompositeFont)
            {
                goto WriteOutName;
            }
            break; // Font may be in the attr array, if it's a system font.
        case DISPID_A_LAYOUTGRID:
            if(fUseCompositeLayoutGrid)
            {
                goto WriteOutName;
            }
            break;
        case DISPID_A_BORDERTOP:
            if(fUseCompositeBorderTop)
            {
                goto WriteOutName;
            }
            continue;
        case DISPID_A_BORDERRIGHT:
            if(fUseCompositeBorderRight)
            {
                goto WriteOutName;
            }
            continue;
        case DISPID_A_BORDERBOTTOM:
            if(fUseCompositeBorderBottom)
            {
                goto WriteOutName;
            }
            continue;
        case DISPID_A_BORDERLEFT:
            if(fUseCompositeBorderLeft)
            {
                goto WriteOutName;
            }
            continue;
        case DISPID_A_LISTSTYLE:
            if(pAA->Find(DISPID_A_LISTSTYLETYPE,      CAttrValue::AA_Attribute) ||
                pAA->Find(DISPID_A_LISTSTYLEPOSITION, CAttrValue::AA_Attribute) ||
                pAA->Find(DISPID_A_LISTSTYLEIMAGE,    CAttrValue::AA_Attribute))
            {
                goto WriteOutName;
            }
            continue;
        case DISPID_A_MARGIN:
            if(fUseCompositeMargin)
            {
                goto WriteOutName;
            }
            continue;
        }
        // does this property exist in the aa?  We may skip this test if
        // this is a composite property (hence the gotos above), since they
        // may not actually be in the AA.
        if((pAV=pAA->Find(ppropdesc->GetDispid(), CAttrValue::AA_Attribute)) != NULL )
        {

WriteOutName:
            //Write out property name
            if(!pchLastOut || 0!=StrCmpC(pchLastOut, ppropdesc->pstrName))
            {
                if(pchLastOut)
                {
                    cstrStyle.Append(_T("; "));
                }
                cstrStyle.Append(ppropdesc->pstrName);
                cstrStyle.Append(_T(": "));
                pchLastOut = ppropdesc->pstrName;
            }
            else
            {
                cstrStyle.Append(_T(" "));
            }

            //Write out value
            hr = CALL_METHOD(*pppropdescs, (*pppropdescs)->pfnHandleProperty, (HANDLEPROP_AUTOMATION|(PROPTYPE_BSTR<<16),
                (void*)&bstr, NULL, (CVoid*)(void*)&pAA));
            if(hr == S_OK)
            {
                if(fQuote)
                {
                    // We are outputing to an inline style. For now we have to replace the 
                    //  double quotes to single, because of external pair of double quotes. Later
                    //   we might want to escape the quotes. We might also want to change how we pass
                    // in and out the string, because changing inplace will not work when escaping
                    // changes the string length, of course.
                    EscapeQuotes((LPTSTR)bstr);
                }
                cstrStyle.Append(bstr);
                FormsFreeString(bstr);
            }

            // BUGBUG: Don't handle "! important" on composite properties.
            if(pAV && pAV->IsImportant())
            {
                cstrStyle.Append(_T("! important"));
            }
        }

        // See if there's an expression
        if((pAV=pAA->Find(ppropdesc->GetDispid(), CAttrValue::AA_Expression)) != NULL)
        {
            // BUGBUG (michaelw) this will probably screw up if expressions are applied to composite properties
            //
            // The problem is actually caused by the brain dead way we loop through every possible attribute
            // apparently to make composite props come out right.  Sheesh!
            if(pchLastOut)
            {
                cstrStyle.Append(_T("; "));
            }

            cstrStyle.Append(_T("; "));
            cstrStyle.Append(ppropdesc->pstrName);
            cstrStyle.Append(_T(": expression("));
            cstrStyle.Append(pAV->GetLPWSTR());
            cstrStyle.Append(_T(")"));
            pchLastOut = ppropdesc->pstrName;
        }
    }

    // Look for all expandos & dump them out
    if(pObject)
    {
        pAV = (CAttrValue*)(*pAA);
        iLen = pAA->Size();
        for(idx=0; idx<iLen; idx++)
        {
            if((pAV->AAType() == CAttrValue::AA_Expando))
            {
                hr = pObject->GetExpandoName(pAV->GetDISPID(), &lpPropName);
                if(hr)
                {
                    continue;
                }

                if(pAV->GetIntoString(&bstrTemp, &lpPropValue))
                {
                    continue; // Can't convert to string
                }
                if(fQuote)
                {
                    // We are outputing to an inline style. For now we have to replace the 
                    //  double quotes to single, because of external pair of double quotes. Later
                    //   we might want to escape the quotes.
                    EscapeQuotes((LPTSTR)lpPropValue);
                }

                if(pchLastOut)
                {
                    cstrStyle.Append( _T("; "));
                }
                cstrStyle.Append(lpPropName);
                pchLastOut = lpPropName;
                cstrStyle.Append(_T(": "));
                cstrStyle.Append(lpPropValue);
                if(bstrTemp)
                {
                    SysFreeString(bstrTemp);
                }
            }
            pAV++;
        }
    }

    // Quote style string?
    if(fQuote)
    {
        cstrStyle.Append(_T("\""));
    }

    if(cstrStyle.Length())
    {
        cstrStyle.AllocBSTR(pbstr);
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteBackgroundStyleToBSTR
//
//  Synopsis:   Converts 'background' property to a BSTR
//
//-------------------------------------------------------------------------
HRESULT WriteBackgroundStyleToBSTR(CAttrArray* pAA, BSTR* pbstr)
{
    CString     cstrBackground;
    BSTR        bstr = NULL;
    CAttrValue* pAV;

    if(S_OK==s_propdescCStylebackgroundImage.b.GetStyleComponentProperty(&bstr, NULL, (CVoid*)&pAA) && bstr && *bstr)
    {
        if(_tcsicmp(bstr, _T("none")))
        {
            cstrBackground.Append((TCHAR*)bstr);
        }
    }
    else
    {
        goto Cleanup;
    }

    pAV = pAA->Find(DISPID_BACKCOLOR, CAttrValue::AA_Attribute);
    if(pAV)
    {
        TCHAR szBuffer[64];
        CColorValue cvColor = (CColorValue)pAV->GetLong();

        if(S_OK == cvColor.FormatBuffer(szBuffer, ARRAYSIZE(szBuffer), NULL))
        {
            if(_tcsicmp(szBuffer, _T("transparent")))
            {
                if(cstrBackground.Length())
                {
                    cstrBackground.Append(_T(" "));
                }
                cstrBackground.Append(szBuffer);
            }
        }
    }
    else
    {
        cstrBackground.Set(_T(""));
        goto Cleanup;
    }

    FormsFreeString(bstr);
    bstr = NULL;

    if(S_OK==s_propdescCStylebackgroundAttachment.b.GetEnumStringProperty(&bstr, NULL, (CVoid*)&pAA) && bstr && *bstr)
    {
        if(_tcsicmp( bstr, _T("scroll")))
        {
            if(cstrBackground.Length())
            {
                cstrBackground.Append(_T(" "));
            }
            cstrBackground.Append((TCHAR*)bstr);
        }
    }
    else
    {
        cstrBackground.Set(_T(""));
        goto Cleanup;
    }

    FormsFreeString(bstr);
    bstr = NULL;

    if(S_OK==s_propdescCStylebackgroundRepeat.b.GetEnumStringProperty(&bstr, NULL, (CVoid *)&pAA) && bstr && *bstr)
    {
        if(_tcsicmp(bstr, _T("repeat")))
        {
            if(cstrBackground.Length())
            {
                cstrBackground.Append(_T(" "));
            }
            cstrBackground.Append((TCHAR*)bstr);
        }
    }
    else
    {
        cstrBackground.Set(_T(""));
        goto Cleanup;
    }

    FormsFreeString(bstr);
    bstr = NULL;

    if(S_OK == WriteBackgroundPositionToBSTR(pAA, &bstr) && bstr && *bstr)
    {
        if(_tcsicmp(bstr, _T("0% 0%")))
        {
            if(cstrBackground.Length())
            {
                cstrBackground.Append(_T(" "));
            }
            cstrBackground.Append((TCHAR*)bstr);
        }
    }
    else
    {
        cstrBackground.Set(_T(""));
        goto Cleanup;
    }

    // If we got this far, we have all the right properties in the AA... but if all of them were default,
    if(!cstrBackground.Length())
    {   // We want to put all the defaults in the string.
        cstrBackground.Set(_T("none transparent scroll repeat 0% 0%"));
    }

Cleanup:
    FormsFreeString(bstr);
    bstr = NULL;

    return cstrBackground.AllocBSTR(pbstr);
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteFontToBSTR
//
//  Synopsis:   Cooks up a BSTR of all the font properties in CSS format.
//              We will only build a non-empty string and return it if
//              there are the minimum set of font properties (if the font
//              string is valid according to the CSS spec, which requires
//              at least a size and a font-family).
//
//-------------------------------------------------------------------------
HRESULT WriteFontToBSTR(CAttrArray* pAA, BSTR* pbstr)
{
    CString cstrFont;
    BSTR    bstr = NULL;
    BOOL    fIsValid = TRUE;
    VARIANT v;

    Assert(pbstr != NULL);
    VariantInit(&v);

    if(S_OK==s_propdescCStylefontWeight.b.GetEnumStringProperty(&bstr, NULL, (CVoid*)&pAA) && bstr && *bstr)
    {
        if(_tcsicmp(bstr, _T("normal")))
        {
            cstrFont.Append((TCHAR*)bstr);
            cstrFont.Append(_T(" "));
        }
    }
    else
    {
        fIsValid = FALSE;
        goto Cleanup;
    }

    FormsFreeString(bstr);
    bstr = NULL;

    if(S_OK==s_propdescCStylefontStyle.b.GetEnumStringProperty(&bstr, NULL, (CVoid*)&pAA) && bstr && *bstr)
    {
        if(_tcsicmp(bstr, _T("normal")))
        {
            cstrFont.Append((TCHAR*)bstr);
            cstrFont.Append(_T(" "));
        }
    }
    else
    {
        fIsValid = FALSE;
        goto Cleanup;
    }

    FormsFreeString(bstr);
    bstr = NULL;

    if(S_OK==s_propdescCStylefontVariant.b.GetEnumStringProperty(&bstr, NULL, (CVoid*)&pAA) && bstr && *bstr)
    {
        if(_tcsicmp(bstr, _T("normal")))
        {
            cstrFont.Append((TCHAR*)bstr);
            cstrFont.Append(_T(" "));
        }
    }
    else
    {
        fIsValid = FALSE;
        goto Cleanup;
    }

    FormsFreeString( bstr );
    bstr = NULL;

    if(S_OK==s_propdescCStylefontSize.a.HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v, NULL, (CVoid*)&pAA) && (v.vt==VT_BSTR && (LPTSTR)v.byref && *(LPTSTR)v.byref))
    {
        cstrFont.Append((LPTSTR)v.byref);
    }
    else
    {
        fIsValid = FALSE;
        goto Cleanup;
    }

    VariantClear(&v);

    if(S_OK==s_propdescCStylelineHeight.a.HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v, NULL, (CVoid*)&pAA) && (v.vt==VT_BSTR && (LPTSTR)v.byref && *(LPTSTR)v.byref))
    {
        if(_tcsicmp((LPTSTR)v.byref, _T("normal")))
        {
            cstrFont.Append(_T("/"));
            cstrFont.Append((LPTSTR)v.byref);
        }
    }
    else
    {
        fIsValid = FALSE;
        goto Cleanup;
    }
    cstrFont.Append(_T(" "));

    bstr = NULL;
    if(S_OK==s_propdescCStylefontFamily.b.GetStringProperty(&bstr, NULL, (CVoid*)&pAA) &&
        bstr && *bstr)
    {
        cstrFont.Append((TCHAR*)bstr);
    }
    else
    {
        fIsValid = FALSE;
    }

Cleanup:
    FormsFreeString(bstr);
    VariantClear(&v);
    if(!fIsValid)
    {
        cstrFont.Set(_T(""));
    }
    return cstrFont.AllocBSTR(pbstr);
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteLayoutGridToBSTR
//
//  Synopsis:
//
//-------------------------------------------------------------------------
HRESULT WriteLayoutGridToBSTR(CAttrArray* pAA, BSTR* pbstr)
{
    CString     cstrLayoutGrid;
    BSTR        bstrMode = NULL;
    BSTR        bstrType = NULL;
    BOOL        fIsValid = TRUE;
    CVariant    v1, v2;

    Assert(pbstr != NULL);

    if(S_OK==s_propdescCStylelayoutGridMode.b.GetEnumStringProperty(&bstrMode, NULL, (CVoid*)&pAA) && bstrMode && *bstrMode)
    {
        if(_tcsicmp(bstrMode, _T("none")))
        {
            cstrLayoutGrid.Append((TCHAR*)bstrMode);
        }
    }
    else
    {
        fIsValid = FALSE;
        goto Cleanup;
    }

    if(S_OK==s_propdescCStylelayoutGridType.b.GetEnumStringProperty(&bstrType, NULL, (CVoid*)&pAA) && bstrType && *bstrType)
    {
        if(_tcsicmp(bstrType, _T("loose")))
        {
            cstrLayoutGrid.Append(_T(" "));
            cstrLayoutGrid.Append((TCHAR*)bstrType);
        }
    }
    else
    {
        fIsValid = FALSE;
        goto Cleanup;
    }

    if(S_OK==s_propdescCStylelayoutGridLine.a.HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v1, NULL, (CVoid*)&pAA) && (v1.vt==VT_BSTR && v1.bstrVal && *v1.bstrVal) )
    {
        cstrLayoutGrid.Append(_T(" "));
        cstrLayoutGrid.Append((TCHAR*)v1.bstrVal);
        fIsValid = TRUE;
    }
    else
    {
        fIsValid = FALSE;
        goto Cleanup;
    }

    if(S_OK==s_propdescCStylelayoutGridChar.a.HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v2, NULL, (CVoid*)&pAA) && (v2.vt==VT_BSTR && v2.bstrVal && *v2.bstrVal))
    {
        cstrLayoutGrid.Append(_T(" "));
        cstrLayoutGrid.Append((TCHAR*)v2.bstrVal);
    }
    else
    {
        fIsValid = FALSE;
        goto Cleanup;
    }

Cleanup:
    FormsFreeString(bstrMode);
    FormsFreeString(bstrType);
    if(!fIsValid)
    {
        cstrLayoutGrid.Set(_T(""));  
    }
    return cstrLayoutGrid.AllocBSTR(pbstr);
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteTextDecorationToBSTR
//
//  Synopsis:   Cooks up a BSTR of all the text-decoration properties in CSS format.
//
//-------------------------------------------------------------------------
HRESULT WriteTextDecorationToBSTR(CAttrArray* pAA, BSTR* pbstr)
{
    CString cstrTextDecoration;
    CAttrValue* pAV;

    Assert(pbstr != NULL);
    if(!pAA)
    {
        return S_FALSE;
    }

    pAV = pAA->Find(DISPID_A_TEXTDECORATION, CAttrValue::AA_Attribute);
    if(pAV)
    {
        // We've got one!
        if(pAV->GetLong() & TD_NONE)
        {
            cstrTextDecoration.Append(_T("none "));
        }
        if(pAV->GetLong() & TD_UNDERLINE)
        {
            cstrTextDecoration.Append(_T("underline "));
        }
        if(pAV->GetLong() & TD_OVERLINE)
        {
            cstrTextDecoration.Append(_T("overline "));
        }
        if(pAV->GetLong() & TD_LINETHROUGH)
        {
            cstrTextDecoration.Append(_T("line-through "));
        }
        if(pAV->GetLong() & TD_BLINK)
        {
            cstrTextDecoration.Append(_T("blink "));
        }
        cstrTextDecoration.TrimTrailingWhitespace();
    }

    RRETURN(cstrTextDecoration.AllocBSTR(pbstr));
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteTextAutospaceToBSTR
//
//  Synopsis:   Cooks up a BSTR of all the text-autospace properties in CSS format.
//
//-------------------------------------------------------------------------
HRESULT WriteTextAutospaceToBSTR(CAttrArray* pAA, BSTR* pbstr)
{
    CString cstrTextAutospace;
    CAttrValue* pAV;

    Assert(pbstr != NULL);
    if(!pAA)
    {
        return S_FALSE;
    }

    pAV = pAA->Find(DISPID_A_TEXTAUTOSPACE, CAttrValue::AA_Attribute);
    if(pAV)
    {
        WriteTextAutospaceFromLongToBSTR(pAV->GetLong(), pbstr, FALSE);
    }
    else
    {
        cstrTextAutospace.Set(_T(""));
        cstrTextAutospace.AllocBSTR(pbstr);
    }

    RRETURN(S_OK);
}

//+----------------------------------------------------------------
//
//  member : WriteTextAutospaceFromLongToBSTR
//
//  Synopsis : given the current textAutospace property, this
//             will write it out to a string
//
//+----------------------------------------------------------------
void WriteTextAutospaceFromLongToBSTR(LONG lTextAutospace, BSTR* pbstr, BOOL fWriteNone)
{
    CString cstrTA;

    if(!lTextAutospace)
    {
        if(fWriteNone)
        {
            cstrTA.Set(_T("none"));
        }
        else
        {
            cstrTA.Set(_T(""));
        }
    }
    else
    {
        if(lTextAutospace & TEXTAUTOSPACE_ALPHA)
        {
            cstrTA.Append(_T("ideograph-alpha "));
        }
        if(lTextAutospace & TEXTAUTOSPACE_NUMERIC)
        {
            cstrTA.Append(_T("ideograph-numeric "));
        }
        if(lTextAutospace & TEXTAUTOSPACE_SPACE)
        {
            cstrTA.Append(_T("ideograph-space "));
        }
        if(lTextAutospace & TEXTAUTOSPACE_PARENTHESIS)
        {
            cstrTA.Append(_T("ideograph-parenthesis"));
        }
        cstrTA.TrimTrailingWhitespace();
    }

    cstrTA.AllocBSTR(pbstr);
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteBorderToBSTR
//
//  Synopsis:   Cooks up a BSTR of the border properties in CSS format.
//
//-------------------------------------------------------------------------
// BUGBUG: This function could be more efficient by doing all the Find()s itself
// and collapsing the values directly. - CWilso
HRESULT WriteBorderToBSTR(CAttrArray* pAA, BSTR* pbstr)
{
    CString cstrBorder;
    BSTR    bstrTemp;
    HRESULT hr = S_OK;

    Assert(pbstr != NULL);
    hr = WriteExpandedPropertyToBSTR(DISPID_A_BORDERCOLOR, pAA, &bstrTemp);
    if((hr==S_OK) && bstrTemp)
    {
        if(!_tcschr(bstrTemp, _T(' ')))
        {
            cstrBorder.Append((TCHAR*)bstrTemp);
        }
        else
        {
            hr = S_FALSE;
        }
        FormsFreeString(bstrTemp);
    }
    if(hr != S_OK)
    {
        goto Cleanup;
    }

    hr = WriteExpandedPropertyToBSTR(DISPID_A_BORDERWIDTH, pAA, &bstrTemp);
    if(hr == S_OK)
    {
        if(!_tcschr(bstrTemp, _T(' ')))
        {
            if(StrCmpC((TCHAR*)bstrTemp, _T("medium")))
            {
                if(cstrBorder.Length())
                {
                    cstrBorder.Append(_T(" "));
                }
                cstrBorder.Append((TCHAR*)bstrTemp);
            }
        }
        else
        {
            hr = S_FALSE;
        }
        FormsFreeString(bstrTemp);
    }
    if(hr != S_OK)
    {
        goto Cleanup;
    }

    hr = WriteExpandedPropertyToBSTR(DISPID_A_BORDERSTYLE, pAA, &bstrTemp);
    if(hr == S_OK)
    {
        if(!_tcschr(bstrTemp, _T(' ')))
        {
            if(StrCmpC((TCHAR*)bstrTemp, _T("none")))
            {
                if(cstrBorder.Length())
                {
                    cstrBorder.Append(_T(" "));
                }
                cstrBorder.Append((TCHAR*)bstrTemp);
            }
        }
        else
        {
            hr = S_FALSE;
        }
        FormsFreeString(bstrTemp);
    }

Cleanup:
    if(hr == S_FALSE)
    {
        cstrBorder.Set(_T(""));
    }
    return cstrBorder.AllocBSTR(pbstr);
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteExpandedPropertyToBSTR
//
//  Synopsis:   Cooks up a BSTR of an aggregate expanded property (e.g.
//              margin, padding, etc.) if the aggregate can be built.  Will
//              accomplish minization (e.g. "10px 10px 10px 10px" will be
//              written as "10px").
//
//-------------------------------------------------------------------------
HRESULT WriteExpandedPropertyToBSTR(DWORD dwDispId, CAttrArray* pAA, BSTR* pbstr)
{
    PROPERTYDESC*   ppdTop;
    PROPERTYDESC*   ppdRight;
    PROPERTYDESC*   ppdBottom;
    PROPERTYDESC*   ppdLeft;
    LPTSTR          pszTop          = NULL;
    LPTSTR          pszRight        = NULL;
    LPTSTR          pszBottom       = NULL;
    LPTSTR          pszLeft         = NULL;
    VARIANT         v;
    CString         cstrRetVal;
    BOOL            fWriteRightLeft = FALSE;
    BOOL            fWriteBottom    = FALSE;
    BOOL            fWriteLeft      = FALSE;
    BOOL            fTopInAA, fRightInAA, fBottomInAA, fLeftInAA;
    HRESULT         hr              = S_OK;

    Assert(pbstr != NULL);
    Assert(pAA && "Must have AttrArray!");
    if(!pAA)
    {
        goto Error;
    }

    if(GetExpandingPropdescs(dwDispId, &ppdTop, &ppdRight, &ppdBottom, &ppdLeft))
    {
        goto Error;
    }

    fTopInAA =    NULL != pAA->Find(ppdTop->GetDispid());
    fRightInAA =  NULL != pAA->Find(ppdRight->GetDispid());
    fBottomInAA = NULL != pAA->Find(ppdBottom->GetDispid());
    fLeftInAA =   NULL != pAA->Find(ppdLeft->GetDispid());

    if(!fTopInAA || !fRightInAA || !fBottomInAA || !fLeftInAA)
    {
        if((dwDispId==DISPID_A_BORDERCOLOR) &&
            !fTopInAA && !fRightInAA && !fBottomInAA && !fLeftInAA)
        {
            goto Cleanup;
        }
        else
        {
            goto Error; // Need all four sides to cook up expanded property.
        }
    }

    if(CALL_METHOD(ppdTop, ppdTop->pfnHandleProperty, (HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v, NULL, (CVoid*)&pAA)))
    {
        goto Error;
    }
    pszTop = (LPTSTR)v.byref;


    if(CALL_METHOD(ppdRight, ppdRight->pfnHandleProperty, (HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v, NULL, (CVoid*)&pAA)))
    {
        goto Error;
    }
    pszRight = (LPTSTR)v.byref;

    if(CALL_METHOD(ppdBottom, ppdBottom->pfnHandleProperty, (HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v, NULL, (CVoid*)&pAA)))
    {
        goto Error;
    }
    pszBottom = (LPTSTR)v.byref;

    if(CALL_METHOD(ppdLeft, ppdLeft->pfnHandleProperty, (HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v, NULL, (CVoid*)&pAA)))
    {
        goto Error;
    }
    pszLeft = (LPTSTR)v.byref;

    if(!pszTop || !pszRight || !pszBottom || !pszLeft)
    {
        goto Error;
    }

    cstrRetVal.Append(pszTop); // We always have the top string
    if(_tcsicmp( pszRight, pszLeft))
    {
        // Right and left don't match - write out everything.
        fWriteRightLeft = TRUE;
        fWriteBottom = TRUE;
        fWriteLeft = TRUE;
    }
    else
    {
        if(_tcsicmp(pszTop, pszBottom))
        {
            fWriteBottom = TRUE; // Top and bottom don't match
            fWriteRightLeft = TRUE;
        }
        else if(_tcsicmp(pszTop, pszRight))
        {
            fWriteRightLeft = TRUE;
        }
    }

    if(fWriteRightLeft)
    {
        cstrRetVal.Append(_T(" "));
        cstrRetVal.Append(pszRight);    // Write out the right string (may be left also)
    }
    if(fWriteBottom)
    {
        cstrRetVal.Append(_T(" "));
        cstrRetVal.Append(pszBottom);   // Write out the bottom string
    }
    if(fWriteLeft)
    {
        cstrRetVal.Append(_T(" "));
        cstrRetVal.Append(pszLeft);     // Write out the left string
    }

Cleanup:
    if(pszLeft)
    {
        FormsFreeString(pszLeft);
    }
    if(pszRight)
    {
        FormsFreeString(pszRight);
    }
    if(pszTop)
    {
        FormsFreeString(pszTop);
    }
    if(pszBottom)
    {
        FormsFreeString(pszBottom);
    }

    if(hr == S_OK)
    {
        hr = cstrRetVal.AllocBSTR(pbstr);
    }

    RRETURN1(hr, S_FALSE);

Error:
    hr = S_FALSE;
    goto Cleanup;
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteListStyleToBSTR
//
//  Synopsis:   Cooks up a BSTR of the list item properties in CSS format.
//
//-------------------------------------------------------------------------
HRESULT WriteListStyleToBSTR(CAttrArray* pAA, BSTR* pbstr)
{
    HRESULT hr = S_FALSE;
    CString cstrListStyle;
    BSTR    bstr = NULL;

    Assert(pbstr != NULL);
    s_propdescCStylelistStyleType.b.GetEnumStringProperty(&bstr, NULL, (CVoid*)&pAA);
    if(bstr && *bstr)
    {
        cstrListStyle.Append((TCHAR*)bstr);
        hr = S_OK;
    }

    FormsFreeString(bstr);
    bstr = NULL;

    if(S_OK == s_propdescCStylelistStyleImage.b.GetStyleComponentProperty(&bstr, NULL, (CVoid*)&pAA))
    {
        if(bstr && *bstr)
        {
            if(hr == S_OK)
            {
                cstrListStyle.Append(_T(" "));
            }
            cstrListStyle.Append((TCHAR*)bstr);
            hr = S_OK;
        }
    }

    FormsFreeString(bstr);
    bstr = NULL;

    s_propdescCStylelistStylePosition.b.GetEnumStringProperty(&bstr, NULL, (CVoid*)&pAA);
    if(bstr && *bstr)
    {
        if(hr == S_OK)
        {
            cstrListStyle.Append(_T(" "));
        }
        cstrListStyle.Append((TCHAR*)bstr);
        hr = S_OK;
    }
    if(hr == S_OK)
    {
        cstrListStyle.AllocBSTR(pbstr);
    }

    FormsFreeString(bstr);

    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteBorderSidePropertyToBSTR
//
//  Synopsis:   Cooks up a BSTR of all the border properties applied to a
//              particular side in CSS format.
//
//-------------------------------------------------------------------------
HRESULT WriteBorderSidePropertyToBSTR(DWORD dispid, CAttrArray* pAA, BSTR* pbstr)
{
    PROPERTYDESC*   ppdStyle;
    PROPERTYDESC*   ppdColor;
    PROPERTYDESC*   ppdWidth;
    CString         cstrBorder;
    BSTR            bstr = NULL;
    HRESULT         hr = S_OK;
    VARIANT         v;

    Assert(pbstr != NULL);
    VariantInit(&v);

    if(S_FALSE == GetBorderSidePropdescs(dispid, &ppdStyle, &ppdColor, &ppdWidth))
    {
        return S_FALSE;
    }

    hr = ((PROPERTYDESC_BASIC*)ppdColor)->b.GetColor((CVoid*)&pAA, &cstrBorder);
    if(hr != S_OK)
    {
        goto Cleanup;
    }

    hr = ppdWidth->HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v, NULL, (CVoid*)&pAA);
    if((hr==S_OK) && (v.vt==VT_BSTR) && v.byref)
    {
        if(_tcsicmp((LPTSTR)v.byref, _T("medium")))
        {
            if(cstrBorder.Length())
            {
                cstrBorder.Append(_T(" "));
            }
            cstrBorder.Append((LPTSTR)v.byref);
        }
    }
    else
    {
        cstrBorder.Set(_T(""));
        goto Cleanup;
    }

    hr = ((PROPERTYDESC_NUMPROP*)ppdStyle)->b.GetEnumStringProperty(&bstr, NULL, (CVoid*)&pAA);
    if(hr != S_OK)
    {
        goto Cleanup;
    }
    if(bstr && *bstr)
    {
        if(_tcsicmp(bstr, _T("none")))
        {
            if(cstrBorder.Length())
            {
                cstrBorder.Append(_T(" "));
            }
            cstrBorder.Append((TCHAR*)bstr);
        }
    }
    else
    {
        cstrBorder.Set(_T(""));
        goto Cleanup;
    }


    if(!cstrBorder.Length()) // All defaults
    {
        cstrBorder.Set(_T("medium none"));
    }

Cleanup:
    VariantClear(&v);
    FormsFreeString(bstr);
    if(hr == S_OK)
    {
        hr = cstrBorder.AllocBSTR(pbstr);
    }
    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteBackgroundPositionToBSTR
//
//  Synopsis:   Cooks up a BSTR of the background position.
//
//-------------------------------------------------------------------------
HRESULT WriteBackgroundPositionToBSTR(CAttrArray* pAA, BSTR* pbstr)
{
    CString cstrBGPos;
    VARIANT v;

    Assert(pbstr != NULL);
    VariantInit( &v );

    if(pAA->Find(DISPID_A_BACKGROUNDPOSX, CAttrValue::AA_Attribute) ||
        pAA->Find(DISPID_A_BACKGROUNDPOSY, CAttrValue::AA_Attribute))
    {
        if(S_OK == s_propdescCStylebackgroundPositionX.a.HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            &v, NULL, (CVoid*)&pAA))
        {
            cstrBGPos.Append((LPTSTR)v.byref);
            VariantClear(&v);

            if(S_OK == s_propdescCStylebackgroundPositionY.a.HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
                &v, NULL, (CVoid*)&pAA))
            {
                cstrBGPos.Append(_T(" "));
                cstrBGPos.Append((LPTSTR)v.byref);
                VariantClear(&v);
            }
        }
    }

    RRETURN(cstrBGPos.AllocBSTR(pbstr));
}

//+------------------------------------------------------------------------
//
//  Function:   ::WriteClipToBSTR
//
//  Synopsis:   Cooks up a BSTR of the "clip" region property.
//
//-------------------------------------------------------------------------
HRESULT WriteClipToBSTR(CAttrArray* pAA, BSTR* pbstr)
{
    CString cstrClip;
    VARIANT v;
    HRESULT hr;

    Assert(pbstr != NULL);
    VariantInit(&v);
    hr = s_propdescCStyleclipTop.a.HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v, NULL, (CVoid*)&pAA);
    if(hr==S_OK && v.byref)
    {
        cstrClip.Append(_T("rect("));
        cstrClip.Append((LPTSTR)v.byref);
        VariantClear(&v);
        hr = s_propdescCStyleclipRight.a.HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
            &v, NULL, (CVoid*)&pAA);
        if(hr==S_OK && v.byref)
        {
            cstrClip.Append(_T(" "));
            cstrClip.Append((LPTSTR)v.byref);
            VariantClear(&v);
            hr = s_propdescCStyleclipBottom.a.HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
                &v, NULL, (CVoid*)&pAA);
            if(hr==S_OK && v.byref)
            {
                cstrClip.Append(_T(" "));
                cstrClip.Append((LPTSTR)v.byref);
                VariantClear(&v);
                hr = s_propdescCStyleclipLeft.a.HandleUnitValueProperty(HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
                    &v, NULL, (CVoid*)&pAA);
                if(hr==S_OK && v.byref)
                {
                    cstrClip.Append(_T(" "));
                    cstrClip.Append((LPTSTR)v.byref);
                    cstrClip.Append(_T(")"));
                    VariantClear(&v);
                    RRETURN(cstrClip.AllocBSTR(pbstr));
                }
            }
        }
    }
    cstrClip.Free();
    RRETURN(cstrClip.AllocBSTR(pbstr));
}

//+------------------------------------------------------------------------
//
//  Function:     ::ParseBackgroundProperty
//
//  Synopsis:
//      This function reads a background property string from the given
//  data string, setting the internal style data of the CAttrArray to reflect any
//  given background styling.
//-------------------------------------------------------------------------
HRESULT ParseBackgroundProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszBGString, BOOL bValidate)
{
    LPTSTR pszString, pszCopy;
    LPTSTR pszNextToken;
    HRESULT hr = S_OK;
    BOOL fSeenXPos = FALSE;
    BOOL fSeenYPos = FALSE;
    PROPERTYDESC* pPropertyDesc;
    PROPERTYDESC* ppdPosX   = (PROPERTYDESC*)&s_propdescCStylebackgroundPositionX.a;
    PROPERTYDESC* ppdPosY   = (PROPERTYDESC*)&s_propdescCStylebackgroundPositionY.a;
    BOOL fSeenBGColor       = FALSE;
    BOOL fSeenBGImage       = FALSE;
    BOOL fSeenBGRepeat      = FALSE;
    BOOL fSeenBGAttachment  = FALSE;
    TCHAR* pszLastXToken    = NULL;

    if(!pcszBGString)
    {
        pcszBGString = _T("");
    }

    pszCopy = pszNextToken = pszString = new TCHAR[_tcslen(pcszBGString)+1];

    if(!pszCopy)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }
    _tcscpy(pszCopy, pcszBGString);

    for(; pszString&&*pszString; pszString=pszNextToken)
    {
        while(_istspace(*pszString))
        {
            pszString++;
        }

        pszNextToken = NextParenthesizedToken(pszString);

        if(*pszNextToken)
        {
            *pszNextToken++ = _T('\0');
        }

        hr = s_propdescCStylebackgroundRepeat.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);

        if(hr != S_OK)
        {
            hr = s_propdescCStylebackgroundAttachment.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);

            if(hr != S_OK)
            {
                if(ValidStyleUrl(pszString)>0 || !_tcsicmp(pszString, _T("none")))
                {
                    hr = s_propdescCStylebackgroundImage.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
                    fSeenBGImage = TRUE;
                }
                else
                {
                    // Try it as a color
                    hr = s_propdescCStylebackgroundColor.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
                    if(hr == S_OK)
                    {
                        // We decided this was a color.
                        fSeenBGColor = TRUE;
                        continue;
                    }
                    else // Not a color.
                    {
                        // First see if it's a measurement string
                        BOOL fTriedOneWay = FALSE;

                        // If two numeric values given, horizontal comes first (CSS spec).  If keywords
                        // are used ("top", "left" etc), order is irrelevant.
                        if(!fSeenXPos)
                        {
TryAsX:
                            hr = s_propdescCStylebackgroundPositionX.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
                            if(hr == S_OK)
                            {
                                fSeenXPos = TRUE;
                                if(pszLastXToken && !fSeenYPos)
                                {
                                    if(s_propdescCStylebackgroundPositionY.a.TryLoadFromString(
                                        dwOpCode, pszLastXToken, pObject, ppAA) == S_OK)
                                    {
                                        fSeenYPos = TRUE;
                                    }
                                }

                                pszLastXToken = pszString;
                                continue;
                            }
                            else if(!fTriedOneWay)
                            {
                                fTriedOneWay = TRUE;
                                goto TryAsY;
                            }
                        }

                        if(!fSeenYPos)
                        {
TryAsY:
                            hr = s_propdescCStylebackgroundPositionY.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
                            if(hr == S_OK)
                            {
                                fSeenYPos = TRUE;
                                continue;
                            }
                            else if(!fTriedOneWay)
                            {
                                fTriedOneWay = TRUE;
                                goto TryAsX;
                            }
                        }
                    }

                    // Not a background string... return if we're in validate mode.
                    if(bValidate)
                    {
                        hr = E_INVALIDARG;
                        break;
                    }
                } // end if else test for color
            } // end if test for background-attachment failure
            else
            {
                fSeenBGAttachment = TRUE;
            }
        } // end if test for background-repeat failure
        else
        {
            fSeenBGRepeat = TRUE;
        }
    } // end for next token
    if(fSeenXPos^fSeenYPos)
    {
        // If only one set, must set the other to "50%" as per the CSS spec.
        pPropertyDesc = fSeenXPos ? ppdPosY : ppdPosX;

        hr = CALL_METHOD(pPropertyDesc, pPropertyDesc->pfnHandleProperty,
            ((dwOpCode|(PROPTYPE_LPWSTR<<16)),
            (CVoid*)_T("50%"), pObject, (CVoid*)ppAA));
    }

    if(!fSeenXPos && !fSeenYPos)
    {
        s_propdescCStylebackgroundPositionX.a.TryLoadFromString(dwOpCode, _T("0%"), pObject, ppAA);
        s_propdescCStylebackgroundPositionY.a.TryLoadFromString(dwOpCode, _T("0%"), pObject, ppAA);
    }

    if(!fSeenBGImage)
    {
        s_propdescCStylebackgroundImage.a.TryLoadFromString(dwOpCode, _T("none"), pObject, ppAA);
    }
    if(!fSeenBGColor)
    {
        s_propdescCStylebackgroundColor.a.TryLoadFromString(dwOpCode, _T("transparent"), pObject, ppAA);
    }
    if(!fSeenBGRepeat)
    {
        s_propdescCStylebackgroundRepeat.a.TryLoadFromString(dwOpCode, _T("repeat"), pObject, ppAA);
    }
    if(!fSeenBGAttachment)
    {
        s_propdescCStylebackgroundAttachment.a.TryLoadFromString(dwOpCode, _T("scroll"), pObject, ppAA);
    }

Cleanup:
    delete[] pszCopy ;
    RRETURN1(hr, E_INVALIDARG);
}

//+------------------------------------------------------------------------
//
//  Function:     ::ParseFontProperty
//
//  Synopsis:
//      This function reads an aggregate font property string from the given
//  data string, setting the internal data of the CAttrArray to reflect any
//  font weight, style, variant, size, line-height, or families set by the string.
//-------------------------------------------------------------------------
HRESULT ParseFontProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszFontString)
{
    LPTSTR pszString, pszCopy;
    LPTSTR pszNextToken;
    HRESULT hResult = S_OK;
    TCHAR chTerminator;
    BOOL fFontWeight = FALSE;
    BOOL fFontStyle = FALSE;
    BOOL fFontVariant = FALSE;

    if(!pcszFontString || !*pcszFontString)
    {
        return E_INVALIDARG;
    }

    pszCopy = pszNextToken = pszString = new TCHAR[_tcslen(pcszFontString)+1];
    if(!pszString)
    {
        hResult = E_OUTOFMEMORY;
        goto Cleanup;
    }
    _tcscpy(pszCopy, pcszFontString);

    // Peel off (and handle) any font weight, style, or variant strings at the beginning
    for(; pszString&&*pszString; pszString=pszNextToken)
    {
        while(_istspace(*pszString))
        {
            pszString++;
        }

        pszNextToken = pszString;

        while(*pszNextToken && !_istspace(*pszNextToken))
        {
            pszNextToken++;
        }

        if(*pszNextToken)
        {
            *pszNextToken++ = _T('\0');
        }

        // Try font weight handler first
        hResult = s_propdescCStylefontWeight.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
        if(hResult == S_OK)
        {
            fFontWeight = TRUE;
            continue; // token was a font weight - go on to next token
        }

        // Try font style handler next
        hResult = s_propdescCStylefontStyle.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
        if(hResult == S_OK)
        {
            fFontStyle = TRUE;
            continue; // token was a font style - go on to next token
        }

        // Try font variant handler last
        hResult = s_propdescCStylefontVariant.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
        if(hResult == S_OK)
        {
            fFontVariant = TRUE;
            continue; // token was a font variant - go on to next token
        }

        break; // String was unrecognized; must be a font-size.
    }

    if(!*pszString)
    {
        hResult = E_INVALIDARG;
        goto Cleanup;
    }

    if(*pszNextToken) // If we replaced a space with a NULL, unreplace it.
    {
        *(pszNextToken-1) = _T(' ');
    }

    pszNextToken = NextSize(pszString, &s_propdescCStylefontSize.b);
    if(!pszNextToken)
    {
        hResult = E_INVALIDARG;
        goto Cleanup;
    }

    chTerminator = *pszNextToken;
    *pszNextToken = _T('\0');

    hResult = s_propdescCStylefontSize.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
    if(hResult != S_OK)
    {
        goto Cleanup;
    }

    *pszNextToken = chTerminator;
    while(_istspace(*pszNextToken))
    {
        pszNextToken++;
    }

    if(*pszNextToken == _T('/'))
    {
        // There is a line-height value present
        pszString = ++pszNextToken;
        pszNextToken = NextSize(pszString, &s_propdescCStylelineHeight.b);

        if(!pszNextToken)
        {
            hResult = E_INVALIDARG;
            goto SetDefaults;
        }

        chTerminator = *pszNextToken;
        *pszNextToken = _T('\0');

        hResult = s_propdescCStylelineHeight.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
        if(hResult != S_OK)
        {
            goto SetDefaults;
        }

        *pszNextToken = chTerminator;
        while(_istspace(*pszNextToken))
        {
            pszNextToken++;
        }
    }
    else
    {
        hResult = s_propdescCStylelineHeight.a.TryLoadFromString(dwOpCode, _T("normal"), pObject, ppAA);
        if(hResult != S_OK)
        {
            goto SetDefaults;
        }
    }

    if(!*pszNextToken)
    {
        hResult = E_INVALIDARG;
        goto SetDefaults;
    }

    if(OPCODE(dwOpCode) == HANDLEPROP_AUTOMATION)
    {
        BSTR bstr;

        hResult = FormsAllocString(pszNextToken, &bstr);
        if(hResult)
        {
            goto SetDefaults;
        }

        hResult = s_propdescCStylefontFamily.b.SetStringProperty(bstr, pObject, (CVoid*)ppAA);

        FormsFreeString(bstr);

        if(hResult)
        {
            goto SetDefaults;
        }
    }
    else
    {
        hResult = s_propdescCStylefontFamily.a.TryLoadFromString(dwOpCode, pszNextToken, pObject, ppAA);
    }

SetDefaults:
    // Set the default font weight
    if(!fFontWeight)
    {
        s_propdescCStylefontWeight.a.TryLoadFromString(dwOpCode, _T("normal"), pObject, ppAA);
    }

    // Set the default font style
    if(!fFontStyle)
    {
        s_propdescCStylefontStyle.a.TryLoadFromString(dwOpCode, _T("normal"), pObject, ppAA);
    }

    // Set the default font variant
    if(!fFontVariant)
    {
        s_propdescCStylefontVariant.a.TryLoadFromString(dwOpCode, _T("normal"), pObject, ppAA);
    }

Cleanup:
    delete[] pszCopy;

    RRETURN1(hResult, E_INVALIDARG);
}

//+------------------------------------------------------------------------
//
//  Function:     ::ParseLayoutGridProperty
//
//  Synopsis:
//      This function reads an aggregate layout-grid property string from the given
//  data string, setting the internal data of the CAttrArray to reflect any
//  layout grid values set by the string.
//-------------------------------------------------------------------------
HRESULT ParseLayoutGridProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszGridString)
{
    LPTSTR pszString, pszCopy = NULL;
    LPTSTR pszNextToken;
    HRESULT hResult = S_OK;
    TCHAR chTerminator = _T('\0');
    BOOL fGridMode = FALSE;
    BOOL fGridType = FALSE;

    if(!pcszGridString || !*pcszGridString)
    {
        goto Cleanup;
    }

    pszCopy = pszNextToken = pszString = (LPTSTR)pcszGridString;

    // Peel off (and handle) any grid mode or grid type strings at the beginning
    for(; pszString&&*pszString; pszString=pszNextToken)
    {
        while(_istspace(*pszString))
        {
            pszString++;
        }

        pszNextToken = pszString;

        while(*pszNextToken && !_istspace(*pszNextToken))
        {
            pszNextToken++;
        }

        if(*pszNextToken)
        {
            chTerminator = *pszNextToken;
            *pszNextToken++ = _T('\0');
        }

        // Try layout grid mode first
        hResult = s_propdescCStylelayoutGridMode.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
        if(hResult == S_OK)
        {
            fGridMode = TRUE;
            if(*pszNextToken) // If we replaced a space with a NULL, unreplace it.
            {
                *(pszNextToken-1) = chTerminator;
            }
            continue; // token was a layout grid mode - go on to next token
        }

        // Try layout grid type next
        hResult = s_propdescCStylelayoutGridType.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
        if(hResult == S_OK)
        {
            fGridType = TRUE;
            if(*pszNextToken) // If we replaced a space with a NULL, unreplace it.
            {
                *(pszNextToken-1) = chTerminator;
            }
            continue; // token was a layout grid type - go on to next token
        }

        break; // String was unrecognized; must be a layout grid line amount, possibly followed
        // by a layout grid char amount.
    }

    if(!*pszString)
    {
        goto Cleanup;
    }

    if(*pszNextToken) // If we replaced a space with a NULL, unreplace it.
    {
        *(pszNextToken-1) = chTerminator;
    }

    // Now we can read in the grid line size
    pszNextToken = NextSize(pszString, &s_propdescCStylelayoutGridLine.b);
    if(!pszNextToken)
    {
        hResult = E_INVALIDARG;
        goto Cleanup;
    }

    chTerminator = *pszNextToken;
    *pszNextToken = _T('\0');
    hResult = s_propdescCStylelayoutGridLine.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
    *pszNextToken = chTerminator;

    if(hResult != S_OK)
    {
        goto Cleanup;
    }

    while(_istspace(*pszNextToken))
    {
        pszNextToken++;
    }

    if(!*pszNextToken)
    {
        goto Cleanup;
    }

    // Now we can read in the grid char size
    pszString = pszNextToken;

    pszNextToken = NextSize(pszString, &s_propdescCStylelayoutGridChar.b);
    if(!pszNextToken)
    {
        hResult = E_INVALIDARG;
        goto Cleanup;
    }

    chTerminator = *pszNextToken;
    *pszNextToken = _T('\0');
    hResult = s_propdescCStylelayoutGridChar.a.TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
    *pszNextToken = chTerminator;

Cleanup:
    RRETURN1(hResult, E_INVALIDARG);
}

//+------------------------------------------------------------------------
//
//  Function:     ::ParseExpandProperty
//
//  Synopsis:
//      This function reads a "margin" or "padding" property value from the
//  given data string, setting the internal data of the CAttrArray to reflect
//  any given margins or padding.  This function handles expansion of the
//  margin or padding values if less than four values are present.
//-------------------------------------------------------------------------
HRESULT ParseExpandProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszBGString, DWORD dwDispId, BOOL fIsMeasurements)
{
    LPTSTR pszCopy;
    LPTSTR pszTop, pszRight, pszBottom, pszLeft, pszEnd;
    HRESULT hr = S_OK;
    CUnitValue cuv;
    PROPERTYDESC* ppdTop;
    PROPERTYDESC* ppdRight;
    PROPERTYDESC* ppdBottom;
    PROPERTYDESC* ppdLeft;

    if(!pcszBGString)
    {
        pcszBGString = _T("");
    }

    if(GetExpandingPropdescs(dwDispId, &ppdTop, &ppdRight, &ppdBottom, &ppdLeft))
    {
        return S_FALSE;
    }

    pszTop = pszRight = pszBottom = pszLeft = pszCopy = new TCHAR[_tcslen(pcszBGString)+1];
    if(!pszTop)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }
    _tcscpy(pszCopy, pcszBGString);

    if(*pszCopy)
    {
        while(_istspace(*pszTop))
        {
            pszTop++; // Skip any leading whitespace
        }
        if(!fIsMeasurements)
        {
            pszRight = NextParenthesizedToken(pszTop);
        }
        else
        {
            pszRight = NextSize(pszTop, (NUMPROPPARAMS*)(ppdTop+1));
        }

        if(pszRight && !*pszRight)
        {
            pszRight = NULL;
        }

        if(pszRight)
        {
            *pszRight++ = _T('\0');
        }

        if(fIsMeasurements)
        {
            hr = cuv.FromString(pszTop, ppdTop);
        }
        if(hr)
        {
            goto Cleanup;
        }

        if(pszRight)
        {
            // Got the top value - let's check for the right value.
            while(_istspace(*pszRight))
            {
                pszRight++; // Skip any leading whitespace
            }
            if(!fIsMeasurements)
            {
                pszBottom = NextParenthesizedToken(pszRight);
            }
            else
            {
                pszBottom = NextSize(pszRight, (NUMPROPPARAMS*)(ppdRight+1));
            }

            if(pszBottom && !*pszBottom)
            {
                pszBottom = NULL;
            }

            if(pszBottom)
            {
                *pszBottom++ = _T('\0');
            }

            if(fIsMeasurements)
            {
                hr = cuv.FromString(pszRight, ppdRight);
            }
            if(hr == S_OK)
            {
                pszLeft = pszRight;

                if(pszBottom)
                {
                    // Got the right value - let's check for the bottom value.
                    while(_istspace(*pszBottom))
                    {
                        pszBottom++; // Skip any leading whitespace
                    }
                    if(!fIsMeasurements)
                    {
                        pszLeft = NextParenthesizedToken(pszBottom);
                    }
                    else
                    {
                        pszLeft = NextSize(pszBottom, (NUMPROPPARAMS*)(ppdBottom+1));
                    }

                    if(pszLeft && !*pszLeft)
                    {
                        pszLeft = NULL;
                    }

                    if(pszLeft)
                    {
                        *pszLeft++ = _T('\0');

                        while(_istspace(*pszLeft))
                        {
                            pszLeft++; // Skip any leading whitespace
                        }
                        if(pszLeft && !*pszLeft)
                        {
                            pszLeft = NULL;
                        }
                    }

                    if(fIsMeasurements)
                    {
                        hr = cuv.FromString(pszBottom, ppdBottom);
                    }
                    if(hr == S_OK)
                    {
                        if(pszLeft)
                        {
                            if(!fIsMeasurements)
                            {
                                pszEnd = NextParenthesizedToken(pszLeft);
                            }
                            else
                            {
                                pszEnd = NextSize(pszLeft, (NUMPROPPARAMS*)(ppdLeft+1));
                            }

                            if(pszEnd && !*pszEnd)
                            {
                                pszEnd = NULL;
                            }

                            if(pszEnd)
                            {
                                *pszEnd = _T('\0');
                            }

                            // Got the bottom value - let's check for the left value.
                            if(fIsMeasurements)
                            {
                                hr = cuv.FromString(pszLeft, ppdLeft);
                            }
                            if(hr != S_OK)
                            {
                                pszLeft = pszRight; // Failed to get left value
                            }
                        }
                        else // No left value present
                        {
                            pszLeft = pszRight;
                        }
                    }
                    else // Failed to get bottom value
                    {
                        pszBottom = pszTop;
                    }
                }
                else // No bottom or left value present
                {
                    pszBottom = pszTop;
                }
            }
            else // Failed to get right value
            {
                pszRight = pszTop;
            }
        }
        else // Only one margin present
        {
            pszRight = pszTop;
        }
    }
    else
    {   // Empty string should still clear the properties.
        pszTop = pszRight = pszBottom = pszLeft = pszCopy;
    }

    hr = ppdTop->TryLoadFromString(dwOpCode, pszTop, pObject, ppAA);
    if(hr)
    {
        goto Cleanup;
    }
    hr = ppdRight->TryLoadFromString(dwOpCode, pszRight, pObject, ppAA);
    if(hr)
    {
        goto Cleanup;
    }
    hr = ppdBottom->TryLoadFromString(dwOpCode, pszBottom, pObject, ppAA);
    if(hr)
    {
        goto Cleanup;
    }
    hr = ppdLeft->TryLoadFromString(dwOpCode, pszLeft, pObject, ppAA);

Cleanup:
    delete[] pszCopy ;
    RRETURN1(hr, E_INVALIDARG);
}

//+------------------------------------------------------------------------
//
//  Function:     ::ParseBorderProperty
//
//  Synopsis:
//      This function reads an aggregate border property string from the given
//  data string, setting the internal data of the CAttrArray to reflect any
//  border styles, widths, or colors set by the string.
//-------------------------------------------------------------------------
HRESULT ParseBorderProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszBorderString)
{
    HRESULT hr = S_OK;


    hr = s_propdescCStyleborderTop.a.TryLoadFromString (dwOpCode, pcszBorderString,pObject, ppAA);
    if(hr!=S_OK && hr!=E_INVALIDARG)
    {
        return hr;
    }

    hr = s_propdescCStyleborderRight.a.TryLoadFromString(dwOpCode, pcszBorderString,pObject, ppAA);
    if(hr!=S_OK && hr!=E_INVALIDARG)
    {
        return hr;
    }

    hr = s_propdescCStyleborderBottom.a.TryLoadFromString(dwOpCode, pcszBorderString,pObject, ppAA);
    if(hr!=S_OK && hr!=E_INVALIDARG)
    {
        return hr;
    }

    hr = s_propdescCStyleborderLeft.a.TryLoadFromString(dwOpCode, pcszBorderString,pObject, ppAA);

    return hr;
}

//+------------------------------------------------------------------------
//
//  Function:     ::ParseAndExpandBorderSideProperty
//
//  Synopsis:
//      This function takes an aggregate border side property string (e.g. the
//  string for "border-top") and sets the internal data of the CAttrArray to
//  reflect any border style, width or color set by the string.
//-------------------------------------------------------------------------
HRESULT ParseAndExpandBorderSideProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszBorderString, DWORD dwDispId)
{
    LPTSTR          pszString, pszCopy, pszNextToken;
    HRESULT         hr = S_OK;
    PROPERTYDESC*   ppdStyle;
    PROPERTYDESC*   ppdColor;
    PROPERTYDESC*   ppdWidth;
    int             nSeenStyle = 0;
    int             nSeenWidth = 0;
    int             nSeenColor = 0;

    if(S_FALSE == GetBorderSidePropdescs(dwDispId, &ppdStyle, &ppdColor, &ppdWidth))
    {
        return S_FALSE;
    }

    if(!pcszBorderString)
    {
        pcszBorderString = _T("");
    }

    pszCopy = pszNextToken = pszString = new TCHAR[_tcslen(pcszBorderString)+1];
    if(!pszCopy)
    {
        return E_OUTOFMEMORY;
    }
    _tcscpy(pszCopy, pcszBorderString);

    for(; pszString&&*pszString; pszString=pszNextToken)
    {
        while ( _istspace( *pszString ) )
            pszString++;

        pszNextToken = NextParenthesizedToken(pszString);

        if(*pszNextToken)
        {
            *pszNextToken++ = _T('\0');
        }

        hr = ppdStyle->TryLoadFromString(dwOpCode, pszString, pObject, ppAA);

        if(hr != S_OK)
        {
            // Let's see if it's a measurement string
            hr = ppdWidth->TryLoadFromString(dwOpCode, pszString, pObject, ppAA);
            if(hr != S_OK)
            {
                // Try it as a color
                hr = ppdColor->TryLoadFromString(dwOpCode, pszString, pObject,ppAA);
                if(hr != S_OK)
                {
                    // Not a valid border string token
                    hr = E_INVALIDARG;
                    goto Cleanup;
                }
                else
                {
                    nSeenColor++;
                }
            }
            else
            {
                nSeenWidth++;
            }
        }
        else
        {
            nSeenStyle++;
        }
    }

Cleanup:
    if(nSeenStyle == 0)
    {
        ppdStyle->TryLoadFromString(dwOpCode, _T("none"), pObject, ppAA);
    }
    if(nSeenWidth == 0)
    {
        ppdWidth->TryLoadFromString(dwOpCode, _T("medium"), pObject, ppAA);
    }
    if(nSeenColor == 0)
    {
        DWORD dwVal;
        if(*ppAA)
        {
            (*ppAA)->FindSimpleInt4AndDelete(ppdColor->GetBasicPropParams()->dispid, &dwVal, CAttrValue::AA_StyleAttribute);
        }
    }

    if(!hr && (nSeenStyle>1 || nSeenWidth>1 || nSeenColor>1))
    {
        hr = E_INVALIDARG;
    }

    delete[] pszCopy;

    RRETURN1(hr, E_INVALIDARG);
}