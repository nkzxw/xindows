
#include "stdafx.h"
#include "ElementFactory.h"

HRESULT CElementFactory::PrivateQueryInterface(REFIID iid, void** ppv)
{
    HRESULT hr = S_OK;

    *ppv = NULL;

    switch(iid.Data1)
    {
        QI_INHERITS((IPrivateUnknown*)this, IUnknown)
        QI_TEAROFF(this, IObjectIdentity, NULL)

    default:
        {
            const CLASSDESC* pclassdesc = (const CLASSDESC*)BaseDesc();
            if(pclassdesc &&
                pclassdesc->_apfnTearOff &&
                pclassdesc->_classdescBase._piidDispinterface &&
                (iid==*pclassdesc->_classdescBase._piidDispinterface))
            {
                hr = CreateTearOffThunk(this, (void*)(pclassdesc->_apfnTearOff), NULL, ppv);
            }
        }
    }

    if(!hr)
    {
        if(*ppv)
        {
            (*(IUnknown**)ppv)->AddRef();
        }
        else
        {
            hr = E_NOINTERFACE;
        }
    }
    RRETURN(hr);
}