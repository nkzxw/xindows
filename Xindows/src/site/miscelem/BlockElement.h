
#ifndef __XINDOWS_SITE_MISCELEM_BLOCKELEMENT_H__
#define __XINDOWS_SITE_MISCELEM_BLOCKELEMENT_H__

#define _hxx_
#include "../gen/block.hdl"

class CBlockElement : public CElement
{
    DECLARE_CLASS_TYPES(CBlockElement, CElement)

public:
    DECLARE_MEMCLEAR_NEW_DELETE()
    CBlockElement (ELEMENT_TAG etag, CDocument* pDoc) : CElement(etag, pDoc) {}
    ~CBlockElement() {}

    STDMETHOD(PrivateQueryInterface) (REFIID iid, void** ppv);

    static  HRESULT CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElementResult);

    virtual HRESULT ApplyDefaultFormat(CFormatInfo* pCFI);

    HRESULT Save(CStreamWriteBuff* pStmWrBuff, BOOL fEnd);


#define _CBlockElement_
#include "../gen/block.hdl"

protected:
    DECLARE_CLASSDESC_MEMBERS;

private:
    NO_COPY(CBlockElement);
};

#endif //__XINDOWS_SITE_MISCELEM_BLOCKELEMENT_H__