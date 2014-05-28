
#ifndef __XINDOWS_SITE_BASE_SAVER_H__
#define __XINDOWS_SITE_BASE_SAVER_H__

class CElement;
class CStreamWriteBuff;

// The CTreeSaver class generates HTML for a piece of a tree.  It can also generate
// text, CF_HTML, or a number of other formats.  It's used for clipboard operations
// and OM operations like innerHTML. 
// The CTreeSaver specifically can only save pieces of a tree between two treepos's.
// It can't save partial text runs.  This more advanced case is handled by the
// CRangeSaver.
class CTreeSaver
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    CTreeSaver(CElement* pelToSave, CStreamWriteBuff* pswb, CElement* pContainer=NULL);
    virtual HRESULT Save();

    void SetTextFragSave(BOOL fSave) { _fSaveTextFrag = !!fSave; }
    BOOL GetTextFragSave()           { return _fSaveTextFrag; }

protected:
    CTreeSaver(CMarkup* pMarkup) { _fSaveTextFrag = FALSE; }

    virtual HRESULT SaveSelection(BOOL fEnd) { return S_OK; }
    HRESULT SaveElement(CElement* pel, BOOL fEnd);
    HRESULT SaveTextPos(CTreePos* ptp);

    // IE4 Compat helpers
    DWORD   LineBreakChar(CTreePosGap* ptpg);

    BOOL    IsElementBlockInContainer(CElement* pElement);
    BOOL    ScopesLeftOfStart(CElement* pel);
    BOOL    ScopesRightOfEnd(CElement* pel);

    CLineBreakCompat    _breaker;

    CMarkup*            _pMarkup;
    CElement*           _pelFragment;
    CElement*           _pelContainer;

    CTreePosGap         _tpgStart;
    CTreePosGap         _tpgEnd;

    CStreamWriteBuff*   _pswb;

    CElement*           _pelLastBlockScope;

    // Flags
    DWORD              _fPendingNBSP : 1;
    DWORD              _fSymmetrical : 1;
    DWORD              _fLBStartLeft : 1;
    DWORD              _fLBEndLeft   : 1;
    DWORD              _fSaveTextFrag: 1;

protected:
    // Protected helpers
    HRESULT ForceClose(CElement* pel);
};

#endif //__XINDOWS_SITE_BASE_SAVER_H__