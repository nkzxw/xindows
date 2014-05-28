
#ifndef __XINDOWS_SITE_BASE_MESSAGE_H__
#define __XINDOWS_SITE_BASE_MESSAGE_H__

class CElement;
class CTreeNode;
class CDispNode;

DWORD FormsGetKeyState(); // KeyState helper

class HITTESTRESULTS
{
public:
    HITTESTRESULTS() { memset(this, 0, sizeof(HITTESTRESULTS)); };

    BOOL    _fPseudoHit;    // set to true if pseudohit (i.e. NOT text hit)
    BOOL    _fWantArrow;
    BOOL    _fRightOfCp;
    LONG    _cpHit;
    LONG    _iliHit;
    LONG    _ichHit;
    LONG    _cchPreChars;
};

struct FOCUS_ITEM
{
    CElement*   pElement;
    long        lTabIndex;
    long        lSubDivision;
};

class CMessage : public MSG
{
    void CommonCtor();
    NO_COPY(CMessage);

public:
    DECLARE_MEMALLOC_NEW_DELETE();

    CMessage(const CMessage* pMessage);
    CMessage(const MSG* pmsg);
    CMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    CMessage() { CommonCtor(); }
    ~CMessage();

    void SetNodeHit(CTreeNode* pNodeHit);
    void SetNodeClk(CTreeNode* pNodeClk);

    BOOL IsContentPointValid() const
    {
        return (pDispNode != NULL);
    }
    void SetContentPoint(const CPoint& pt, CDispNode* pDispNodeContent)
    {
        ptContent = pt;
        pDispNode = pDispNodeContent;
    }
    void SetContentPoint(const POINT& pt, CDispNode* pDispNodeContent)
    {
        ptContent = pt;
        pDispNode = pDispNodeContent;
    }

    CPoint      pt;                     // Global  coordinates
    CPoint      ptContent;              // Content coordinates
    CDispNode*  pDispNode;              // CDispNode associated with ptContent
    DWORD       dwKeyState;
    CTreeNode*  pNodeHit;
    CTreeNode*  pNodeClk;
    DWORD       dwClkData;              // Used to pass in any data for DoClick().
    HTC         htc;
    LRESULT     lresult;
    HITTESTRESULTS resultsHitTest;
    long        lSubDivision;

    unsigned    fNeedTranslate:1;       // TRUE if Trident should manually
                                        // call TranslateMessage, raid 44891
    unsigned    fRepeated:1;
    unsigned    fEventsFired:1;         // because we can have situation when FireStdEventsOnMessage
                                        // called twice with the same pMessage, this bit is used to
                                        // prevent firing same events 2nd time
    unsigned    fSelectionHMCalled:1;   // set once the selection has had a shot at the mess. perf.
    unsigned    fStopForward:1;         // prevent CElement::HandleMessage()
};

BOOL IsTabKey(CMessage* pMessage);
BOOL IsFrameTabKey(CMessage* pMessage);

#endif //__XINDOWS_SITE_BASE_MESSAGE_H__