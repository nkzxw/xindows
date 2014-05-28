
#ifndef __XINDOWS_SITE_MISCELEM_PARAELEMENT_H__
#define __XINDOWS_SITE_MISCELEM_PARAELEMENT_H__

#include "BlockElement.h"

#define _hxx_
#include "../gen/para.hdl"

class CParaElement : public CBlockElement
{
    DECLARE_CLASS_TYPES(CParaElement, CBlockElement)

public:
    CParaElement(CDocument* pDoc) : CBlockElement(ETAG_P, pDoc) {}
    ~CParaElement() {}

    static  HRESULT CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElementResult);

    virtual HRESULT Save(CStreamWriteBuff* pStreamWrBuff, BOOL fEnd);

#define _CParaElement_
#include "../gen/para.hdl"

protected:
    DECLARE_CLASSDESC_MEMBERS;

private:
    NO_COPY(CParaElement);
};

#endif //__XINDOWS_SITE_MISCELEM_PARAELEMENT_H__