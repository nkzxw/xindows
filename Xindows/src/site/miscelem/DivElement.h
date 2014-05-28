
#ifndef __XINDOWS_SITE_MISCELEM_DIVELEMENT_H__
#define __XINDOWS_SITE_MISCELEM_DIVELEMENT_H__

#include "BlockElement.h"

#define _hxx_
#include "../gen/div.hdl"

class CDivElement : public CBlockElement
{
    DECLARE_CLASS_TYPES(CDivElement, CBlockElement)

public:
    CDivElement(CDocument* pDoc) : CBlockElement(ETAG_DIV, pDoc) {}

    ~CDivElement() {}

    static HRESULT CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElementResult);

    STDMETHOD(PrivateQueryInterface)(REFIID iid, void** ppv);

#define _CDivElement_
#include "../gen/div.hdl"

protected:
    DECLARE_CLASSDESC_MEMBERS;

private:
    NO_COPY(CDivElement);
};

#endif //__XINDOWS_SITE_MISCELEM_DIVELEMENT_H__