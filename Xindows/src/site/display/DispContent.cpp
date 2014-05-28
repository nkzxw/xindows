
#include "stdafx.h"
#include "DispContent.h"

//+---------------------------------------------------------------------------
//
//  Member:     CDispContentNode::GetZOrder
//              
//  Synopsis:   Return the z order of this node.
//              
//  Arguments:  none
//              
//  Returns:    Z order of this node.
//              
//  Notes:      This method shouldn't be called unless the node is
//              in the negative or positive Z layers.
//              
//----------------------------------------------------------------------------
LONG CDispContentNode::GetZOrder() const
{
    Assert(GetLayerType()==DISPNODELAYER_NEGATIVEZ || GetLayerType()==DISPNODELAYER_POSITIVEZ);
    
    return _pDispClient->GetZOrderForSelf();
}