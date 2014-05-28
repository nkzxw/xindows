
#ifndef __XINDOWS_SITE_BUILTIN_IMGELEMENT_H__
#define __XINDOWS_SITE_BUILTIN_IMGELEMENT_H__

#define _hxx_
#include "../gen/img.hdl"

class CShape;
class CImgElementLayout;

//+---------------------------------------------------------------------------
//
// CImgElement
//
//----------------------------------------------------------------------------
class CImgElement : public CSite
{
    friend class CImgElementLayout;

    DECLARE_CLASS_TYPES(CImgElement, CSite)

    friend class CDBindMethodsImg;

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    CImgElement(ELEMENT_TAG etag, CDocument* pDoc);

    virtual void    Notify(CNotification* pNF);
    virtual HRESULT Init();
    HRESULT EnterTree();

    void EnsureTabStop();

    DECLARE_LAYOUT_FNS(CImgElementLayout)

    STDMETHOD(PrivateQueryInterface)(REFIID iid, void** ppv);

    STDMETHOD(QueryStatus)(
        GUID*       pguidCmdGroup,
        ULONG       cCmds,
        MSOCMD      rgCmds[],
        MSOCMDTEXT* pcmdtext);

    STDMETHOD(Exec)(
        GUID*        pguidCmdGroup,
        DWORD       nCmdID,
        DWORD       nCmdexecopt,
        VARIANTARG* pvarargIn,
        VARIANTARG* pvarargOut);

    static HRESULT CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElementResult);

    virtual void Passivate();
    virtual HRESULT STDMETHODCALLTYPE HandleMessage(CMessage* pMessage);
    NV_DECLARE_MOUSECAPTURE_METHOD(HandleCaptureMessageForImage, handlecapturemessageforimage,
        (CMessage* pMessage));
    NV_DECLARE_MOUSECAPTURE_METHOD(HandleCaptureMessageForArea, handlecapturemessageforarea,
        (CMessage* pMessage));

    virtual HRESULT ClickAction(CMessage* pMessage);

    virtual HRESULT ApplyDefaultFormat(CFormatInfo *pCFI);
    HRESULT Save(CStreamWriteBuff* pStmWrBuff, BOOL fEnd);

    virtual HRESULT DoSubDivisionEvents(long lSubDivision, DISPID dispidMethod, DISPID dispidProp, VARIANT* pvb, BYTE* pbTypes, ...);
    virtual HRESULT GetSubDivisionCount(long* pc);
    virtual HRESULT GetSubDivisionTabs(long* pTabs, long c);
    virtual HRESULT SubDivisionFromPt(POINT pt, long *piSub);
    virtual HRESULT GetFocusShape(long lSubDivision, CDocInfo* pdci, CShape** ppShape);
    virtual HRESULT ShowTooltip(CMessage* pmsg, POINT pt);

    virtual HRESULT OnPropertyChange(DISPID dispid, DWORD dwFlags);

    // InvokeEx to validate ready state.
    NV_DECLARE_TEAROFF_METHOD(ContextThunk_InvokeExReady, invokeexready, (
        DISPID id,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS* pdp,
        VARIANT* pvarRes,
        EXCEPINFO* pei,
        IServiceProvider* pSrvProvider));

    //base implementation prototypes
    NV_DECLARE_TEAROFF_METHOD(put_src, PUT_src, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(get_src, GET_src, (BSTR* p));
    HRESULT STDMETHODCALLTYPE put_height(long v);
    HRESULT STDMETHODCALLTYPE get_height(long* p);
    HRESULT STDMETHODCALLTYPE put_width(long v);
    HRESULT STDMETHODCALLTYPE get_width(long* p);

    STDMETHOD(put_hspace)(long v);
    STDMETHOD(get_hspace)(long* p);

    BOOL    IsHSpaceDefined() const ;

    HRESULT putHeight(long l);
    HRESULT putWidth(long l);
    HRESULT GetHeight(long* pl);
    HRESULT GetWidth(long* pl);

    NV_DECLARE_TEAROFF_METHOD(get_onload, GET_onload, (VARIANT*));
    NV_DECLARE_TEAROFF_METHOD(put_onload, PUT_onload, (VARIANT));

    NV_DECLARE_TEAROFF_METHOD(get_readyState, get_readyState, (BSTR* pBSTR));
    NV_DECLARE_TEAROFF_METHOD(get_readyState, get_readyState, (VARIANT* pReadyState));
    NV_DECLARE_TEAROFF_METHOD(get_readyStateValue, get_readyStateValue, (long* plRetValue));

#define _CImgElement_
#include "../gen/img.hdl"

    CImgHelper* _pImage;

protected:
    DECLARE_CLASSDESC_MEMBERS;
    static const CLSID* s_apclsidPages[];

private:
    RECT            _rcWobbleZone;
    BOOL            _fCanClickImage;
    BOOL            _fBelowAnchor;

    NO_COPY(CImgElement);
};


class CImageElementFactory : public CElementFactory
{
    DECLARE_CLASS_TYPES(CImageElementFactory, CElementFactory)

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    CImageElementFactory() {}
    ~CImageElementFactory() {}

#define _CImageElementFactory_
#include "../gen/img.hdl"

protected:
    DECLARE_CLASSDESC_MEMBERS;
};

#endif //__XINDOWS_SITE_BUILTIN_IMGELEMENT_H__