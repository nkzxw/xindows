
#ifndef __XINDOWS_SITE_DISPLAY_DISPCONTEXT_H__
#define __XINDOWS_SITE_DISPLAY_DISPCONTEXT_H__

class CDispContextStack;
class CDispSurface;
class CDispExtras;
class CDispInfo;

//+---------------------------------------------------------------------------
//
//  Class:      CDispContext
//
//  Synopsis:   Base class for context objects passed throughout display tree.
//
//----------------------------------------------------------------------------
class CDispContext
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()

	CDispContext() {}
	CDispContext(const CRect& rcClip, const CSize& offset)
	{
		_rcClip = rcClip;
		_offset = offset;
	}
	CDispContext(const CDispContext& c)
	{
		_rcClip = c._rcClip;
		_offset = c._offset;
	}
	~CDispContext() {}

	CDispContext& operator=(const CDispContext& c)
	{
		_rcClip = c._rcClip;
		_offset = c._offset;
		return *this;
	}

	// SetNoClip() sets the clip rect to a really big rect, but not
	// the maximum rect, because otherwise translations may cause it to
	// over/underflow.  (For speed reasons, we aren't using CRect::SafeOffset)
	void SetNoClip()
	{
		static const LONG bigVal = 100000000;
		_rcClip.SetRect(-bigVal, -bigVal, bigVal, bigVal);
	}

	void SetToIdentity()
	{
		_offset = _afxGlobalData._Zero.size; SetNoClip();
	}

	void SetClipRect(const CRect& rcClip) { _rcClip = rcClip; }
	const CRect& GetClipRect() const { return _rcClip; }

	// transform a point or region or rect to new coordinate system, with
	// clipping for rects and regions
	void Transform(CPoint* ppt) const { *ppt += _offset; }

	void Transform(CRegion* prgn) const
	{
		prgn->Intersect(_rcClip);
		prgn->Offset(_offset);
	}

	void Transform(CRect* prc) const
	{
		prc->IntersectRect(_rcClip);
		prc->OffsetRect(_offset);
	}

	void GetTransformedClipRect(CRect* prcClip) const
	{
		prcClip->left = _rcClip.left + _offset.cx;
		prcClip->top = _rcClip.top + _offset.cy;
		prcClip->right = _rcClip.right + _offset.cx;
		prcClip->bottom = _rcClip.bottom + _offset.cy;
	}

	// reverse transforms
	void Untransform(CPoint* ppt) const { *ppt -= _offset; }

	void Untransform(CRect* prc) const
	{
		prc->OffsetRect(-_offset);
		prc->IntersectRect(_rcClip);
	}

	void GetUntransformedClipRect(CRect* prcClip) const
	{
		*prcClip = _rcClip;
	}
	const CRect& GetUntransformedClipRect() const
	{
		return _rcClip;
	}

	// to transform from one coordinate system to another, first clip,
	// using _rcClip (which is in source coordinates), then add _offset
	CRect _rcClip;	// in source coordinates
	CSize _offset;	// add to produce destination coordinates
};



//+---------------------------------------------------------------------------
//
//  Class:      CDispHitContext
//
//  Synopsis:   Context used for hit testing.
//
//----------------------------------------------------------------------------
class CDispHitContext : public CDispContext
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()

	CDispHitContext() {}
	~CDispHitContext() {}

	BOOL RectIsHit(const CRect& rc) const
	{
		return rc.Contains(_ptHitTest);
	}

	// BUGBUG. FATHIT. marka - Fix for Bug 65015 - enabling "Fat" hit testing on tables.
	// Edit team is to provide a better UI-level way of dealing with this problem for post IE5.
	//
	// BUGBUG - when this change is done - signature of FuzzyRectIsHit *must* be changed
	// to not take an extra param saying whether to do FatHitTesting
	BOOL FuzzyRectIsHit(const CRect& rc, BOOL fFatHitTest); 

	// BUGBUG. FATHIT. marka - Fix for Bug 65015 - enabling "Fat" hit testing on tables.
	// Edit team is to provide a better UI-level way of dealing with this problem for post IE5.
	BOOL FatRectIsHit(const CRect& rc);


	// hit testing
	CPoint  _ptHitTest;     // hit test point
	void*   _pClientData;   // client supplied data
	long    _cFuzzyHitTest; // pixels to extend hit test rects
};


//+---------------------------------------------------------------------------
//
//  Class:      CDispDrawContext
//
//  Synopsis:   Context used for PreDraw and Draw.
//
//----------------------------------------------------------------------------
class CDispDrawContext : public CDispContext
{
	friend class CDispNode;
	friend class CDispLeafNode;
	friend class CDispInteriorNode;
	friend class CDispContentNode;
	friend class CDispItemPlus;
	friend class CDispRoot;
	friend class CDispContainer;
	friend class CDispContainerPlus;
	friend class CDispScroller;
	friend class CDispScrollerPlus;
	friend class CDispSurface;
	friend class CDispLayerContext;

public:
    DECLARE_MEMALLOC_NEW_DELETE()

	CDispDrawContext() : _drawSelector(CDispFlags::s_drawSelector), _fBypassFilter(FALSE) {}
	~CDispDrawContext() {}


	// client data
	void SetClientData(void* pClientData) { _pClientData = pClientData; }
	void* GetClientData() { return _pClientData; }

	CDispSurface* GetDispSurface();
	void SetDispSurface(CDispSurface* pSurface)
	{
		_pDispSurface = pSurface; SetSurfaceRegion();
	}

protected:
	CDispContext& operator=(const CDispContext& c)
	{
		_rcClip = c._rcClip;
		_offset = c._offset;
		return *this;
	}

	// modify the current offset to global coordinates
	void AddGlobalOffset(const CSize& offset)
	{
		_offset += offset;
	}
	void SubtractGlobalOffset(const CSize& offset)
	{
		_offset -= offset;
	}
	const CSize& GetGlobalOffset() const
	{
		return _offset;
	}
	void SetGlobalOffset(const CSize& offset)
	{
		_offset = offset;
	}

	void SetRedrawRegion(CRegion* prgn)
	{
		_prgnRedraw = prgn;
	}
	CRegion* GetRedrawRegion()
	{
		return _prgnRedraw;
	}
	void AddToRedrawRegion(const CRect& rc)
	{
		CRect rcGlobal(rc);
		Transform(&rcGlobal);
		AddToRedrawRegionGlobal(rcGlobal);
	}
	void AddToRedrawRegionGlobal(const CRect& rcGlobal);
	void IntersectRedrawRegion(CRegion* prgn)
	{
		Transform(prgn);
		prgn->Intersect(*_prgnRedraw);
		prgn->Offset(-_offset);
	}
	void IntersectRedrawRegion(CRect *prc)
	{
		Transform(prc);
		prc->IntersectRect(_prgnRedraw->GetBounds());
		prc->OffsetRect(-_offset);
	}
	BOOL IntersectsRedrawRegion(const CRect& rc)
	{
		CRect rcGlobal(rc);
		Transform(&rcGlobal);
		return _prgnRedraw->Intersects(rcGlobal);
	}
	BOOL IsRedrawRegionEmpty() const
	{
		return _prgnRedraw->IsEmpty();
	}
	BOOL RectContainsRedrawRegion(const CRect& rc) const
	{
		CRect rcRedraw;
		_prgnRedraw->GetBounds(&rcRedraw);
		rcRedraw.OffsetRect(-_offset); // local coords.
		return rc.Contains(rcRedraw);
	}

	// the following methods deal with the region redraw stack, which is used
	// during PreDraw traversal to save the redraw region before opaque and
	// buffered items are subtracted from it
	void SetRedrawRegionStack(CRegionStack* pRegionStack)
	{
		_pRedrawRegionStack = pRegionStack;
	}
	CRegionStack* GetRedrawRegionStack() const
	{
		return _pRedrawRegionStack;
	}
	BOOL PushRedrawRegion(const CRegion& rgn, void* key);
	BOOL PopRedrawRegionForKey(void* key)
	{
		if(!_pRedrawRegionStack->PopRegionForKey(key, &_prgnRedraw))
		{
			return FALSE;
		}
		SetSurfaceRegion();
		return TRUE;
	}
	void* PopFirstRedrawRegion()
	{
		void* key = _pRedrawRegionStack->PopFirstRegion(&_prgnRedraw);
		SetSurfaceRegion();
		return key;
	}
	void RestorePreviousRedrawRegion()
	{
		delete _prgnRedraw;
		_prgnRedraw = _pRedrawRegionStack->RestorePreviousRegion();
	}

	// manage context stack for drawing
	void SetContextStack(CDispContextStack* pContextStack)
	{
		_pContextStack = pContextStack;
	}
	CDispContextStack* GetContextStack() const { return _pContextStack; }
	BOOL PopContext(CDispNode* pDispNode);
	void FillContextStack(CDispNode* pDispNode);
	void SaveContext(const CDispNode* pDispNode, const CDispContext& context);

	// internal methods
	HDC GetRawDC();

	BOOL            _fBypassFilter;         // Bypass filter for rendering
protected:

	// redraw region
	CRegion*        _prgnRedraw;            // redraw region (global coords.)
	CRegion         _rgnRedrawClipped;      // redraw region clipped to current band
	CRegionStack*   _pRedrawRegionStack;    // stack of redraw regions
	CDispNode*      _pFirstDrawNode;        // first node to draw
	CDispRoot*      _pRootNode;             // root node for this tree

	// context stack
	CDispContextStack* _pContextStack;      // saved context information for Draw

	// display surface
	CDispSurface*   _pDispSurface;          // display surface

	// client data
	void*           _pClientData;           // client data

	// draw selector
	CDispFlags      _drawSelector;          // choose which nodes to draw

private:
	void SetSurfaceRegion();
};



//+---------------------------------------------------------------------------
//
//  Class:      CDispContextStack
//
//  Synopsis:   Store interesting parts of the display context (such as
//              clipping rects and offsets) for a branch of the display
//              tree, so that we don't have n^2 calculation of this
//              information while drawing.
//
//----------------------------------------------------------------------------
class CDispContextStack
{
public:
	CDispContextStack() { _stackIndex = _stackMax = 0; }
	~CDispContextStack() {}

	void Init() { _stackIndex = _stackMax = 0; }
	void Restore() { _stackIndex = 0; }

	BOOL PopContext(CDispContext* pContext, const CDispNode* pDispNode)
	{
		if(_stackIndex == _stackMax)  // run out of stack
		{
			return FALSE;
		}
#ifdef _DEBUG
		Assert(_stack[_stackIndex]._pDispNode == pDispNode);
#endif
		*pContext = _stack[_stackIndex++]._context;
		return TRUE;
	}

	BOOL MoreToPop() const { return _stackIndex<_stackMax; }
	BOOL IsEmpty() const { return _stackMax==0; }

	void ReserveSlot(const CDispNode* pDispNode)
	{
#ifdef _DEBUG
		if(_stackIndex < CONTEXTSTACKSIZE)
		{
			_stack[_stackIndex]._pDispNode = pDispNode;
		}
#endif
		_stackIndex++;
		if(_stackMax < CONTEXTSTACKSIZE)
		{
			_stackMax++;
		}
	}

	void PushContext(const CDispContext& context, const CDispNode* pDispNode)
	{
		if(--_stackIndex >= CONTEXTSTACKSIZE)
		{
			return;
		}
#ifdef _DEBUG
		Assert(_stackIndex>=0 && _stackIndex<_stackMax);
		Assert(_stack[_stackIndex]._pDispNode == pDispNode);
#endif
		_stack[_stackIndex]._context = context;
	}

	void SaveContext(const CDispContext& context, const CDispNode* pDispNode)
	{
		if(_stackIndex >= CONTEXTSTACKSIZE) return;
#ifdef _DEBUG
		_stack[_stackIndex]._pDispNode = pDispNode;
#endif
		_stack[_stackIndex++]._context = context;
		Assert(_stackIndex > _stackMax);
		_stackMax = _stackIndex;
	}

#ifdef _DEBUG
	const CDispNode* GetTopNode()
	{
		return ((_stackIndex>=0&&_stackIndex<_stackMax)? _stack[_stackIndex]._pDispNode : NULL);
	}
#endif

private:
	enum { CONTEXTSTACKSIZE = 32 };
	struct stackElement
	{
		CDispContext        _context;
#ifdef _DEBUG
		const CDispNode*    _pDispNode;
#endif
	};

	int             _stackIndex;
	int             _stackMax;
	stackElement    _stack[CONTEXTSTACKSIZE];
};



//+---------------------------------------------------------------------------
//
//  Class:      CApplyUserClip
//
//  Synopsis:   A stack-based object which incorporates optional user clip for
//              the given display node into the total clip rect in the context,
//              and removes it when destructed.
//
//----------------------------------------------------------------------------
class CApplyUserClip
{
public:
	CApplyUserClip(CDispNode* pNode, CDispContext* pContext);
	~CApplyUserClip()
	{
		if(_fHasUserClip)
		{
			_pContext->_rcClip = _rcClipSave;
		}
	}

private:
	BOOL            _fHasUserClip;
	CDispContext*   _pContext;
	CRect           _rcClipSave;
};

#endif //__XINDOWS_SITE_DISPLAY_DISPCONTEXT_H__
