
#include "stdafx.h"
#include "EvtParam.h"

//---------------------------------------------------------------------------
//
//  Member:     EVENTPARAM::EVENTPARAM
//
//  Synopsis:   constructor
//
//  Parameters: pDoc            Doc to bind to
//              fInitState      If true, then initialize state type
//                              members.  (E.g. x, y, keystate, etc).
//
//---------------------------------------------------------------------------
EVENTPARAM::EVENTPARAM(CDocument* pNewDoc, BOOL fInitState, BOOL fPush/*=TRUE*/)
{
    memset(this, 0, sizeof(*this));

    pDoc = pNewDoc;
    pDoc->SubAddRef(); // balanced in destructor

    _fOnStack = FALSE;

    if(fPush)
    {
        Push();
    }

    _lKeyCode = 0;

    if(fInitState)
    {
        POINT pt;

        ::GetCursorPos(&pt);
        _screenX= pt.x;
        _screenY= pt.y;

        {
            if(pDoc->_pInPlace)
            {
                ::ScreenToClient(pDoc->_pInPlace->_hwnd, &pt);
            }
        }
        _clientX = pt.x;
        _clientY = pt.y;
        _sKeyState = VBShiftState();
    }
}

//---------------------------------------------------------------------------
//
//  Member:     EVENTPARAM::~EVENTPARAM
//
//  Synopsis:   destructor
//
//---------------------------------------------------------------------------
EVENTPARAM::~EVENTPARAM()
{
    Pop();

    pDoc->SubRelease(); // to balance SubAddRef in constructor
}

//---------------------------------------------------------------------------
//
//  Member:     EVENTPARAM::Push
//
//---------------------------------------------------------------------------
void EVENTPARAM::Push()
{
    if(!_fOnStack)
    {
        _fOnStack = TRUE;
        pparamPrev = pDoc->_pparam;
        pDoc->_pparam = this;
    }
}

//---------------------------------------------------------------------------
//
//  Member:     EVENTPARAM::Pop
//
//---------------------------------------------------------------------------
void EVENTPARAM::Pop()
{
    if(_fOnStack)
    {
        _fOnStack = FALSE;
        pDoc->_pparam = pparamPrev;
    }
}

//---------------------------------------------------------------------------
//
//  Member:     EVENTPARAM::IsCancelled
//
//---------------------------------------------------------------------------
BOOL EVENTPARAM::IsCancelled()
{
    HRESULT hr;

    if(VT_EMPTY != V_VT(&varReturnValue))
    {
        hr = VariantChangeTypeSpecial(&varReturnValue, &varReturnValue, VT_BOOL);
        if(hr)
        {
            return FALSE; // assume event was not cancelled if changing type to bool failed
        }

        return (VT_BOOL==V_VT(&varReturnValue) && VB_FALSE==V_BOOL(&varReturnValue));
    }
    else
    {
        // if the variant is empty, we consider it is not cancelled by default
        return FALSE;
    }
}

void EVENTPARAM::SetNodeAndCalcCoordinates(CTreeNode *pNewNode)
{
    if(_pNode != pNewNode)
    {
        _pNode = pNewNode;
        if(_pNode) CalcRestOfCoordinates();
    }
}

//+-------------------------------------------------------------------------
//
//  Method:     EVENTPARAM::GetParentCoordinates
//
//  Synopsis:   Helper function for getting the parent coordinates
//              notset/statis -- doc coordinates
//              relative      --  [styleleft, styletop] +
//                                [x-site.rc.x, y-site.rc.y]
//              absolute      --  [A_parent.rc.x, A_parent.rc.y,] +
//                                [x-site.rc.x, y-site.rc.y]
//
//  Parameters : px, py - the return point 
//               fContainer :  TRUE - COORDSYS_CONTAINER
//                             FALSE - COORDSYS_CONTENT
//
//  these parameters are here for NS compat. and as such the parent defined
//  for the positioning are only ones that are absolutely or relatively positioned
//  or the body.
//
//--------------------------------------------------------------------------
HRESULT EVENTPARAM::GetParentCoordinates(long* px, long* py)
{
    CPoint      pt(0, 0);
    HRESULT     hr = S_OK;
    CLayout*    pLayout;
    CElement*   pElement;
    CDispNode*  pDispNode;
    CRect       rc;
    CTreeNode*  pZNode = _pNode;

    if(!_pNode || !_pNode->_pNodeParent)
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

    // First, determine if we need to climg out of a slavetree
    //---------------------------------------------------------
    if(_pNode->Tag() == ETAG_TXTSLAVE)
    {
        Assert(_pNode->GetMarkup());

        CElement* pElemMaster;

        pElemMaster = _pNode->GetMarkup()->Master();
        if(!pElemMaster)
        {
            hr = DISP_E_MEMBERNOTFOUND;
            goto Cleanup;
        }

        pZNode = pElemMaster->GetFirstBranch();
        if(!pZNode)
        {
            hr = DISP_E_MEMBERNOTFOUND;
            goto Cleanup;
        }
    }

    // now the tricky part. We have a Node for the element that 
    // the event is on, but the "parent" that we are reporting
    // position information for is not constant. The parent is the 
    // first ZParent of this node that is relatively positioned
    // and the body is always reported as relative.  Absolutely 
    // positioned elements DON'T COUNT.
    //---------------------------------------------------------

    // walk up looking for a positioned thing.
    while(pZNode && !pZNode->IsRelative() && 
        !(pZNode->Tag()==ETAG_ROOT) && !(pZNode->Tag()==ETAG_BODY))
    {
        pZNode = pZNode->ZParentBranch();
    }

    pElement = pZNode ? pZNode->Element() : NULL;

    // now we know the element that we are reporting a position wrt
    // and so we just need to get the dispnode and the position info
    //---------------------------------------------------------
    if(pElement)
    {
        pLayout = pElement->GetUpdatedNearestLayout();

        if(pLayout)
        {
            pDispNode = pLayout->GetElementDispNode(pElement);

            if(pDispNode)
            {
                pDispNode->TransformPoint(&pt, COORDSYS_CONTAINER, COORDSYS_GLOBAL);
            }
        }
    }

    // adjust for the offset of the mouse wrt the postition of the parent
    pt.x = _clientX - pt.x;
    pt.y = _clientY - pt.y;


    // and return the values.
    if(px)
    {
        *px = pt.x;
    }
    if(py)
    {
        *py = pt.y;
    }

Cleanup:
    RRETURN1(hr, DISP_E_MEMBERNOTFOUND);
}

HRESULT EVENTPARAM::CalcRestOfCoordinates()
{
    HRESULT     hr = S_OK;
    CLayout*    pLayout = NULL;

    if(_pNode)
    {
        hr = GetParentCoordinates(&_x, &_y);
        if(hr)
        {
            goto Cleanup;
        }

        pLayout = _pNode->GetUpdatedNearestLayout();
        if(!pLayout)
        {
            _offsetX = -1;
            _offsetY = -1;
            goto Cleanup;
        }

        // BUGBUG: Do this a better way? (brendand)
        // the _offsetX and _offsetY are reported always INSIDE the borders of the offsetParent
        // but for perf reasons we don't subtract out the borders until they are asked for
        //------------------------------------------------------------------------------------
        _offsetX = _clientX - pLayout->GetPositionLeft(COORDSYS_GLOBAL) + pLayout->GetXScroll();
        _offsetY = _clientY - pLayout->GetPositionTop(COORDSYS_GLOBAL) + pLayout->GetYScroll();
    }

Cleanup:
    RRETURN(hr);
}