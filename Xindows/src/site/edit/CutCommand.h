
#ifndef __XINDOWS_SITE_EDIT_CUTCOMMAND_H__
#define __XINDOWS_SITE_EDIT_CUTCOMMAND_H__

class CXindowsEditor;
class CSpringLoader;

#include "delcmd.h"

//+---------------------------------------------------------------------------
//
//  CCutCommand Class
//
//----------------------------------------------------------------------------
class CCutCommand : public CDeleteCommand
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()

    CCutCommand(DWORD cmdId, CXindowsEditor* pEd) : CDeleteCommand(cmdId, pEd)
    {
    }

    virtual ~CCutCommand()
    {
    }

protected:
    HRESULT PrivateExec(DWORD nCmdexecopt, VARIANTARG* pvarargIn, VARIANTARG* pvarargOut);

    HRESULT PrivateQueryStatus(OLECMD* pcmd, OLECMDTEXT* pcmdtext);
};

#endif //__XINDOWS_SITE_EDIT_CUTCOMMAND_H__