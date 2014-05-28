
#include "stdafx.h"
#include "Message.h"

DWORD FormsGetKeyState()
{
    static int vk[] =
    {
        VK_LBUTTON,     // MK_LBUTTON = 0x0001
        VK_RBUTTON,     // MK_RBUTTON = 0x0002
        VK_SHIFT,       // MK_SHIFT   = 0x0004
        VK_CONTROL,     // MK_CONTROL = 0x0008
        VK_MBUTTON,     // MK_MBUTTON = 0x0010
        VK_MENU,        // MK_ALT     = 0x0020
    };

    DWORD dwKeyState = 0;

    for(int i=0; i<ARRAYSIZE(vk); i++)
    {
        if(GetKeyState(vk[i]) & 0x8000)
        {
            dwKeyState |= (1 << i);
        }
    }

    return dwKeyState;
}

//+---------------------------------------------------------------------------
//
// Function: IsTabKey
//
// Test if CMessage is a TAB KEYDOWN message.
//
//-----------------------------------------------------------------------------
BOOL IsTabKey(CMessage* pMessage)
{
    BOOL fTabOrder = FALSE;

    if(pMessage->message == WM_KEYDOWN)
    {
        if(pMessage->wParam == VK_TAB)
        {
            fTabOrder = (pMessage->dwKeyState&FCONTROL) ? (FALSE) : (TRUE);
        }
    }

    return fTabOrder;
}

//+---------------------------------------------------------------------------
//
// Function: IsFrameTabKey
//
// Test if CMessage is a CTRL+TAB or F6 KEYDOWN message.
//
//-----------------------------------------------------------------------------
BOOL IsFrameTabKey(CMessage* pMessage)
{
    BOOL fResult = FALSE;

    if(pMessage->message == WM_KEYDOWN)
    {
        if((pMessage->wParam==VK_TAB && (pMessage->dwKeyState&FCONTROL)) || (pMessage->wParam==VK_F6))
        {
            fResult = TRUE;
        }
    }
    return fResult;
}

void CMessage::CommonCtor()
{
    memset(this, 0, sizeof(*this));
    dwKeyState = FormsGetKeyState();
    resultsHitTest._cpHit = -1;
}

CMessage::CMessage(const CMessage* pMessage)
{
    if(pMessage)
    {
        memcpy(this, pMessage, sizeof(CMessage));
        // NEWTREE: we don't have subrefs on nodes...? regular add ref?
        if(pNodeHit)
        {
            pNodeHit->NodeAddRef();
        }
    }
}

CMessage::CMessage(const MSG* pmsg)
{
    CommonCtor();
    if(pmsg)
    {
        memcpy(this, pmsg, sizeof(MSG));
        htc = HTC_YES;
    }
}

CMessage::CMessage(HWND hwndIn, UINT msg, WPARAM wParamIn, LPARAM lParamIn)
{
    CommonCtor();
    hwnd      = hwndIn;
    message   = msg;
    wParam    = wParamIn;
    lParam    = lParamIn;
    htc       = HTC_YES;
    time      = GetMessageTime();
    DWORD  dw = GetMessagePos();
    MSG::pt.x = (short)LOWORD(dw);
    MSG::pt.y = (short)HIWORD(dw);

    Assert(!fEventsFired);
    Assert(!fSelectionHMCalled);
}

CMessage::~CMessage()
{
    // NEWTREE: same subref note here
    CTreeNode::ReplacePtr(&pNodeHit, NULL);
}

//+---------------------------------------------------------------------------
//
// Member : CMessage : SetNodeHit
//
//  Synopsis : things like the tracker cache the message and then access it
//  on a timer callback. inorder to ensure that the elements that are in the
//  message are still there, we need to sub(?) addref the element.
//
//----------------------------------------------------------------------------
void CMessage::SetNodeHit(CTreeNode* pNodeHitIn)
{
    // NEWTREE: same subref note here
    CTreeNode::ReplacePtr(&pNodeHit, pNodeHitIn);
}

//+-----------------------------------------------------------------------------
//
//  Member : CMessage::SetElementClk
//
//  Synopsis : consolidating the click firing code requires a helper function
//      to set the click element member of the message structure.
//             this function should only be called from handling a mouse buttonup
//      event message which could fire off a click (i.e. LButton only)
//
//------------------------------------------------------------------------------
void CMessage::SetNodeClk(CTreeNode* pNodeClkIn)
{
    CTreeNode* pNodeDown = NULL;

    Assert(pNodeClkIn && !pNodeClk || pNodeClkIn==pNodeClk);

    // get the element that this message is related to or if there
    // isn't one use the pointer passed in
    pNodeClk = (pNodeHit) ? pNodeHit : pNodeClkIn;

    Assert(pNodeClk);

    // if this is not a LButtonUP just return, using the value
    // set (i.e. we do not need to look for a common ancester)
    if(message != WM_LBUTTONUP)
    {
        return;
    }

    // now go get the _pEltGotButtonDown from the doc and look for
    // the first common ancester between the two. this is the lowest
    // element that the mouse went down and up over.
    pNodeDown = pNodeClk->Element()->Doc()->_pNodeGotButtonDown;

    if(!pNodeDown)
    {
        // Button down was not on this doc, or cleared due to capture.
        // so this is not a click
        pNodeClk = NULL;
    }
    else
    {
        // Convert both nodes from from slave to master before comparison
        if(pNodeDown->Tag() == ETAG_TXTSLAVE)
        {
            pNodeDown = pNodeDown->GetMarkup()->Master()->GetFirstBranch();
        }
        if(pNodeClk->Tag() == ETAG_TXTSLAVE)
        {
            pNodeClk = pNodeClk->GetMarkup()->Master()->GetFirstBranch();
        }

        if(!pNodeDown)
        {
            pNodeClk = NULL;
        }
        if(!pNodeClk)
        {
            return;
        }

        if(pNodeDown != pNodeClk)
        {
            if(!pNodeDown->Element()->TestClassFlag(CElement::ELEMENTDESC_NOANCESTORCLICK))
            {
                // The mouse up is on a different element than us.  This
                // can only happen if someone got capture and
                // forwarded the message to us.  Now we find the first
                // common anscestor and fire the click event from
                // there.
                pNodeClk = pNodeDown->GetFirstCommonAncestor(pNodeClk, NULL);
            }
            else
            {
                pNodeClk = NULL;
            }
        }
    }
}
