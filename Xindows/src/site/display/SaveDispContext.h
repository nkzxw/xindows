
#ifndef __XINDOWS_SITE_DISPLAY_SAVEDISPCONTEXT_H__
#define __XINDOWS_SITE_DISPLAY_SAVEDISPCONTEXT_H__

//+---------------------------------------------------------------------------
//
//  Contents:   Utility class to save interesting parts of the display context
//              during tree traversal.
//
//----------------------------------------------------------------------------
class CSaveDispContext
{
public:
	CSaveDispContext(CDispContext* pContext)
	{
		_pOriginalContext = pContext;
		_saveContext = *pContext;
	}
	~CSaveDispContext()
	{
		*_pOriginalContext = _saveContext;
	}

	CDispContext*   _pOriginalContext;
	CDispContext    _saveContext;
};

#endif //__XINDOWS_SITE_DISPLAY_SAVEDISPCONTEXT_H__
