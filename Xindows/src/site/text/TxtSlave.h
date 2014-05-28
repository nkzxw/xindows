
#ifndef __XINDOWS_SITE_TEXT_TXTSLAVE_H__
#define __XINDOWS_SITE_TEXT_TXTSLAVE_H__

class CTxtSlave : public CElement
{
private:
    DECLARE_MEMCLEAR_NEW_DELETE()
    DECLARE_CLASS_TYPES(CTxtSlave, CElement)
public:
    CTxtSlave(ELEMENT_TAG etag, CDocument* pDoc) : super(etag, pDoc)
    {
    }

    static HRESULT CreateElement(CHtmTag* pht, CDocument*pDoc, CElement** ppElementResult);

    virtual HRESULT Init2(CInit2Context* pContext);
    virtual BOOL    IsEnabled() { return MarkupMaster()->IsEnabled(); }

    virtual HRESULT ComputeFormats(CFormatInfo* pCFI, CTreeNode* pNodeTarget);
    virtual DWORD   GetBorderInfo(CDocInfo* pdci, CBorderInfo* pborderinfo, BOOL fAll);

    virtual HRESULT STDMETHODCALLTYPE HandleMessage(CMessage* pMessage)
    {
        // Kill the message if the master is dead. That can happen
        // if an event handler nuked the master (#30760)
        if(!IsInMarkup() || !MarkupMaster())
        {
            return S_OK;
        }

        // super will pass to the master
        return super::HandleMessage(pMessage);
    }       

    HRESULT STDMETHODCALLTYPE QueryStatus(
        GUID* pguidCmdGroup,
        ULONG cCmds,
        MSOCMD rgCmds[],
        MSOCMDTEXT* pcmdtext)
    {
        if(IsInMarkup())
        {
            CElement* pElemMaster = MarkupMaster();

            if(pElemMaster)
            {
                RRETURN_NOTRACE(pElemMaster->QueryStatus(
                    pguidCmdGroup,
                    cCmds,
                    rgCmds,
                    pcmdtext));
            }
        }

        RRETURN_NOTRACE(super::QueryStatus(
            pguidCmdGroup,
            cCmds,
            rgCmds,
            pcmdtext));
    }

    HRESULT STDMETHODCALLTYPE Exec(
        GUID* pguidCmdGroup,
        DWORD nCmdID,
        DWORD nCmdexecopt,
        VARIANTARG* pvarargIn,
        VARIANTARG* pvarargOut)
    {
        if(IsInMarkup())
        {
            CElement* pElemMaster = MarkupMaster();

            if(pElemMaster)
            {
                RRETURN_NOTRACE(pElemMaster->Exec(
                    pguidCmdGroup,
                    nCmdID,
                    nCmdexecopt,
                    pvarargIn,
                    pvarargOut));
            }
        }
        RRETURN_NOTRACE(super::Exec(
            pguidCmdGroup,
            nCmdID,
            nCmdexecopt,
            pvarargIn,
            pvarargOut));
    }

    virtual HRESULT CreateLayout() { return S_OK; /* no layout! */ }

protected:
    DECLARE_CLASSDESC_MEMBERS
};

#endif //__XINDOWS_SITE_TEXT_TXTSLAVE_H__