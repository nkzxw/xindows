
#ifndef __XINDOWS_API_XINDOWSAPI_H__
#define __XINDOWS_API_XINDOWSAPI_H__

interface IHTMLDocument2;
interface IHTMLElement;

XINDOWS_PUBLIC IHTMLDocument2* CreateDocument(HWND hWndParent);
XINDOWS_PUBLIC void SetDocumentRect(IHTMLDocument2* pIDocument, RECT& rc);
XINDOWS_PUBLIC void DeleteDocument(IHTMLDocument2* pIDocument);

XINDOWS_PUBLIC BOOL AppendElement(IHTMLElement* pIParent, IHTMLElement* pIChild);

#endif //__XINDOWS_API_XINDOWSAPI_H__