
#include "stdafx.h"
#include "XindowsAPI.h"

XINDOWS_PUBLIC IHTMLDocument2* CreateDocument(HWND hWndParent)
{
    CEnsureThreadState ets;

    InitDocClass();

    CDocument* pDoc = new CDocument();
    pDoc->Init();
    pDoc->InitNew();
    pDoc->RunningToInPlace(hWndParent, NULL);

    CElement* pBody = NULL;
    pDoc->_pPrimaryMarkup->CreateElement(ETAG_BODY, &pBody);
    pDoc->_pPrimaryMarkup->Root()->InsertAdjacent(CElement::AfterBegin, pBody);

    pBody->Release();

    return pDoc;
}

XINDOWS_PUBLIC void SetDocumentRect(IHTMLDocument2* pIDocument, RECT& rc)
{
    CDocument* pDoc = NULL;
    pIDocument->QueryInterface(CLSID_HTMLDocument, (void**)&pDoc);

    Assert(pDoc);

    if(pDoc)
    {
        pDoc->SetObjectRects(&rc, &rc);
    }
}

XINDOWS_PUBLIC void DeleteDocument(IHTMLDocument2* pIDocument)
{
    CDocument* pDoc = NULL;
    pIDocument->QueryInterface(CLSID_HTMLDocument, (void**)&pDoc);

    Assert(pDoc);

    if(pDoc)
    {
        pDoc->InPlaceToRunning();
        pDoc->Release();
    }
}

XINDOWS_PUBLIC BOOL AppendElement(IHTMLElement* pIParent, IHTMLElement* pIChild)
{
    if(!pIParent || !pIChild)
    {
        return FALSE;
    }

    CElement* pParent = NULL;
    CElement* pChild = NULL;

    pIParent->QueryInterface(CLSID_CElement, (void**)&pParent);
    pIChild->QueryInterface(CLSID_CElement, (void**)&pChild);

    Assert(pParent && pChild);

    return SUCCEEDED(pParent->InsertAdjacent(CElement::BeforeEnd, pChild));
}