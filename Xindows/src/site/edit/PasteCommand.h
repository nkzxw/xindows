
#ifndef __XINDOWS_SITE_EDIT_PASTECOMMAND_H__
#define __XINDOWS_SITE_EDIT_PASTECOMMAND_H__

class CXindowsEditor;
class CSpringLoader;

#include "delcmd.h"

//+---------------------------------------------------------------------------
//
//  CPasteCommand Class
//
//----------------------------------------------------------------------------

class CPasteCommand : public CDeleteCommand
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()

    CPasteCommand(DWORD cmdId, CXindowsEditor* pEd) : CDeleteCommand(cmdId, pEd)
    {
    }

    virtual ~CPasteCommand()
    {
    }

    HRESULT Paste(IMarkupPointer* pStart, IMarkupPointer* pEnd, CSpringLoader* psl, BSTR bstrText=NULL);   

    HRESULT PasteFromClipboard(IMarkupPointer* pStart, IMarkupPointer* pEnd, IDataObject* pDO, CSpringLoader* psl);

    HRESULT InsertText(OLECHAR* pchText, long cch, IMarkupPointer* pPointerTarget, CSpringLoader* psl);

protected:
    HRESULT PrivateExec(DWORD nCmdexecopt, VARIANTARG* pvarargIn, VARIANTARG* pvarargOut);

    HRESULT PrivateQueryStatus(OLECMD* pcmd, OLECMDTEXT* pcmdtext);

private:
    HRESULT IsPastePossible(IMarkupPointer* pStart, IMarkupPointer* pEnd, BOOL* pfResult);
};

#endif //__XINDOWS_SITE_EDIT_PASTECOMMAND_H__