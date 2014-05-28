
#ifndef __XINDOWS_CORE_BASE_INPLACE_H__
#define __XINDOWS_CORE_BASE_INPLACE_H__

class COffScreenContext;

//+---------------------------------------------------------------------------
//
//  Class:      CInPlace (cinpl)
//
//  Purpose:    Encapsulates the in-memory data storage of an in-place or
//              UI active object.  All interface methods are forwarded to
//              equivalents on CServer.  Any member data used only when inplace
//              should be put on this object.
//
//----------------------------------------------------------------------------
class CInPlace
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE();

	CInPlace();
	virtual ~CInPlace();

	POINT					_ptWnd;					// Location of window in container
	RECT					_rcPos;					// Position
	RECT					_rcClip;				// Clip

	HWND					_hwnd;					// Our window

	unsigned				_fIPNoDraw:1;
	unsigned				_fIPPaintBkgnd:1;
	unsigned				_fIPOffScreen:1;
	unsigned				_fDeactivating:1;		// are we deactivating?
	unsigned				_fFocus:1;				// True if we have the focus.
	unsigned				_fUsingWindowRgn:1;		// True if clipping with window region.
	unsigned				_fBubbleInsideOut:1;	// True if bubbling was caused from an inside control

	COffScreenContext*		_pOSC;
	IDataObject*			_pDataObj;				// saved data obj from begin of drop.
};

#endif //__XINDOWS_CORE_BASE_INPLACE_H__