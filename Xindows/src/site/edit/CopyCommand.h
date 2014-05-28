
#ifndef __XINDOWS_SITE_EDIT_COPYCOMMAND_H__
#define __XINDOWS_SITE_EDIT_COPYCOMMAND_H__

class CXindowsEditor;
class CSpringLoader;

//+---------------------------------------------------------------------------
//
//  CCopyCommand Class
//
//----------------------------------------------------------------------------
class CCopyCommand : public CCommand
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()

    CCopyCommand(DWORD cmdId, CXindowsEditor* pEd) : CCommand(cmdId, pEd)
    {
    }

    virtual ~CCopyCommand()
    {
    }

protected:
    HRESULT PrivateExec(DWORD nCmdexecopt, VARIANTARG* pvarargIn, VARIANTARG* pvarargOut);

    HRESULT PrivateQueryStatus(OLECMD* pcmd, OLECMDTEXT* pcmdtext);
};

#endif //__XINDOWS_SITE_EDIT_COPYCOMMAND_H__