
#ifndef __XINDOWS_SITE_TEXT_BODYELEMENT_H__
#define __XINDOWS_SITE_TEXT_BODYELEMENT_H__

#define _hxx_
#include "../gen/body.hdl"

class CBodyLayout;

//+---------------------------------------------------------------------------
//
// CBodyElement
//
//----------------------------------------------------------------------------
class CBodyElement : public CTxtSite
{
    DECLARE_CLASS_TYPES(CBodyElement, CTxtSite)

public:
    enum { NUM_COMMON_PROPS = 5 };

    DECLARE_MEMCLEAR_NEW_DELETE()

    STDMETHOD(PrivateQueryInterface)(REFIID iid, void** ppv);

    CBodyElement(CDocument* pDoc);

    static  HRESULT CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElementResult);

    virtual HRESULT Init2(CInit2Context* pContext);

    virtual void    Notify(CNotification* pNF);

    virtual HRESULT ApplyDefaultFormat(CFormatInfo* pCFI);

    virtual DWORD   GetBorderInfo(CDocInfo* pdci, CBorderInfo* pborderinfo, BOOL fAll=FALSE);

    DECLARE_LAYOUT_FNS(CBodyLayout)

    virtual HRESULT GetFocusShape(long lSubDivision, CDocInfo* pdci, CShape** ppShape);

    // The following are used by the parser to determine whether or not to
    // apply the attrs of a BODY tag to an existing body element. If the
    // body was synthesized because something was seen that needed to be
    // parented by a body, and then a BODY tag was seen, the attrs of the
    // late-arriving tag will be applied to the synthesized element (which
    // will then no longer be considered synthesized).
    void            SetSynthetic(BOOL fSynthetic) { _fSynthetic = fSynthetic;}
    BOOL            GetSynthetic() { return _fSynthetic; }

    virtual HRESULT OnPropertyChange(DISPID dispid, DWORD dwFlags);

    void            WaitForRecalc();

#define _CBodyElement_
#include "../gen/body.hdl"

    static CElement::ACCELS s_AccelsBodyDesign;
    static CElement::ACCELS s_AccelsBodyRun;

protected:
    DECLARE_CLASSDESC_MEMBERS;

    static const CLSID* s_apclsidPages[];

private:
    unsigned    _fSynthetic:1;

public:
    NO_COPY(CBodyElement);
};

#endif //__XINDOWS_SITE_TEXT_BODYELEMENT_H__