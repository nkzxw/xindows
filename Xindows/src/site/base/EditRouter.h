
#ifndef __XINDOWS_SITE_BASE_EDITROUTER_H__
#define __XINDOWS_SITE_BASE_EDITROUTER_H__

class CEditRouter
{
public:
    CEditRouter();
    ~CEditRouter();

    void Passivate();

    HRESULT ExecEditCommand(
        GUID*       pguidCmdGroup,
        DWORD       nCmdID,
        DWORD       nCmdexecopt,
        VARIANTARG* pvarargIn,
        VARIANTARG* pvarargOut,
        IUnknown*   punkContext,
        CDocument*  pDoc);

    HRESULT QueryStatusEditCommand(
        GUID*       pguidCmdGroup,
        ULONG       cCmds,
        MSOCMD      rgCmds[],
        MSOCMDTEXT* pcmdtext,
        IUnknown*   punkContext,
        CDocument*  pDoc);

private:
    HRESULT SetHostEditHandler(IUnknown* punkContext, CDocument* pDoc);
    HRESULT SetInternalEditHandler(IUnknown* punkContext, CDocument* pDoc, BOOL fForceCreate);

    IOleCommandTarget*  _pHostCmdTarget; // Alternate Editing Interface provided by the host
    BOOL                _fHostChecked;
    IOleCommandTarget*  _pInternalCmdTarget;
};

#endif //__XINDOWS_SITE_BASE_EDITROUTER_H__