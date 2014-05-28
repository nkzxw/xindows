
#ifndef __XINDOWS_SITE_BUILTIN_TEXTAREA_H__
#define __XINDOWS_SITE_BUILTIN_TEXTAREA_H__

#define _hxx_
#include "../gen/textarea.hdl"

#define TEXT_INSET_DEFAULT_TOP      1
#define TEXT_INSET_DEFAULT_BOTTOM   1
#define TEXT_INSET_DEFAULT_RIGHT    1
#define TEXT_INSET_DEFAULT_LEFT     1

//+---------------------------------------------------------------------------
//
// CRichtext: HTML textarea
//
//----------------------------------------------------------------------------
class CRichtext: public CTxtSite
{
    DECLARE_CLASS_TYPES(CRichtext, CTxtSite)

public:
    CRichtext(ELEMENT_TAG etag, CDocument* pDoc) : CTxtSite(etag, pDoc)
    {
        _vStatus.vt = VT_NULL;
        _dwLastCP = 0;
    };

    static HRESULT CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElementResult);
    virtual HRESULT Init2(CInit2Context* pContext);
    virtual HRESULT GetFocusShape(long lSubDivision, CDocInfo* pdci, CShape** ppShape)
    {
        // Never want a focus rect
        Assert(ppShape);
        *ppShape = NULL;
        return S_FALSE;
    }

    DECLARE_LAYOUT_FNS(CRichtextLayout)

    virtual HRESULT SetDefaultValueHelper(CString& cstrDefault) { return _cstrDefaultValue.Set(cstrDefault); }
    virtual HRESULT SetDefaultValueHelper(TCHAR* psz) { return _cstrDefaultValue.Set(psz); }

    virtual HRESULT DoReset(void);

    // property helpers
    NV_DECLARE_PROPERTY_METHOD(GetValueHelper, GETValueHelper, (CString* pbstr));
    NV_DECLARE_PROPERTY_METHOD(SetValueHelper, SETValueHelper, (CString* bstr));

    HRESULT SetValueHelperInternal(CString* pstr, BOOL fOM=TRUE);

    //+-----------------------------------------------------------------------
    //  CTxtSite methods
    //------------------------------------------------------------------------
    virtual BOOL GetWordWrap() const;
    virtual HRESULT OnPropertyChange(DISPID dispid, DWORD dwFlags);
    virtual HRESULT STDMETHODCALLTYPE HandleMessage(CMessage* pMessage);
    
    //+-----------------------------------------------------------------------
    // Helper methods
    //------------------------------------------------------------------------
    virtual void    Notify(CNotification* pNF);

    //+------------------------------------------------------------------------
    // CElement methods
    //-----------------------------------------------------------------------
    virtual HRESULT ApplyDefaultFormat(CFormatInfo* pCFI);

    virtual HRESULT YieldCurrency(CElement* pElemNew);
    virtual HRESULT RequestYieldCurrency(BOOL fForce);
    virtual HRESULT BecomeUIActive();
    
    //+-----------------------------------------------------------------------
    //  Interface methods
    //------------------------------------------------------------------------

    // IOleCommandTarget methods
    HRESULT STDMETHODCALLTYPE QueryStatus(
        GUID*       pguidCmdGroup,
        ULONG       cCmds,
        MSOCMD      rgCmds[],
        MSOCMDTEXT* pcmdtext);
    HRESULT STDMETHODCALLTYPE Exec(
        GUID*       pguidCmdGroup,
        DWORD       nCmdID,
        DWORD       nCmdexecopt,
        VARIANTARG* pvarargIn,
        VARIANTARG* pvarargOut);

    // default value html property
    CString     _cstrDefaultValue;
    CString     _cstrLastValue;
    DWORD       _dwLastCP;
    unsigned    _fTextChanged:1;
    unsigned    _fFiredValuePropChange:1;
    unsigned    _fDoReset:1;
    unsigned    _fInSave:1;
    unsigned    _fLastValueSet:1;
    unsigned    _fChangingEncoding:1;
    unsigned    _iHistoryIndex:16;

#define _CRichtext_
#include "../gen/textarea.hdl"

protected:
    DECLARE_CLASSDESC_MEMBERS

    VARIANT     _vStatus;

    static ACCELS s_AccelsTextareaDesign;
    static ACCELS s_AccelsTextareaRun;

private:
    NO_COPY(CRichtext);
};

//+---------------------------------------------------------------------------
//
// CTextarea: plain text textarea
//
//----------------------------------------------------------------------------
class CTextArea : public CRichtext
{
    DECLARE_CLASS_TYPES(CTextArea, CRichtext)
    friend class CHTMLTextAreaParser;

public:
    CTextArea(ELEMENT_TAG etag, CDocument* pDoc) : CRichtext(etag, pDoc) {}
    static HRESULT CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElementResult);

    virtual HRESULT Save(CStreamWriteBuff* pStreamWrBuff, BOOL fEnd);

    DECLARE_LAYOUT_FNS(CTextAreaLayout)

    long GetPlainTextLengthWithBreaks();
    void GetPlainTextWithBreaks(TCHAR* pchBuff);

    // property helpers
    virtual HRESULT GetSubmitValue(CString* pstr);

    //+------------------------------------------------------------------------
    // CElement methods
    //-----------------------------------------------------------------------
    virtual HRESULT ApplyDefaultFormat(CFormatInfo* pCFI);


    // default value html property
    CString _cstrDefaultValue;

#define _CTextArea_
#include "../gen/textarea.hdl"

protected:
    DECLARE_CLASSDESC_MEMBERS

private:
    NO_COPY(CTextArea);
};

#endif //__XINDOWS_SITE_BUILTIN_TEXTAREA_H__