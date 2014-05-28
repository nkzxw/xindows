
#ifndef __XINDOWS_CORE_COM_ERRORINFO_H__
#define __XINDOWS_CORE_COM_ERRORINFO_H__

//+----------------------------------------------------------------
//
//  Enum:       EPART
//
//  Synopsis:   Error message parts. The parts are composed into
//              a message using the following template.  Square
//              brackets are used to indication optional components
//              of the message.
//
//                 [<EPART_ACTION>] <EPART_ERROR>
//
//                 [Solution:
//                 <EPART_SOLUTION>]
//
//              If EPART_ERROR is not specified, it is derived from
//              the HRESULT.
//
//-----------------------------------------------------------------
enum EPART
{
    EPART_ACTION,
    EPART_ERROR,
    EPART_SOLUTION,
    EPART_LAST
};

//+----------------------------------------------------------------
//
//  Class:      CErrorInfo
//
//  Synopsis:   Rich error information object. Use GetErrorInfo() to get
//              the current error information object.  Use ClearErrorInfo()
//              to delete the current error information object.  Use
//              CloseErrorInfo to set the per thread error information object.
//
//-----------------------------------------------------------------
class CErrorInfo : public IErrorInfo
{
public:
    CErrorInfo();
    ~CErrorInfo();

    DECLARE_MEMCLEAR_NEW_DELETE();

    // IUnknown methods
    DECLARE_STANDARD_IUNKNOWN(CErrorInfo);

    // IErrorInfo methods
    STDMETHOD(GetGUID)(GUID* pguid);
    STDMETHOD(GetSource)(BSTR* pbstrSource);
    STDMETHOD(GetDescription)(BSTR* pbstrDescription);
    STDMETHOD(GetHelpFile)(BSTR* pbstrHelpFile);
    STDMETHOD(GetHelpContext)(DWORD* pdwHelpContext);

    // IErrorInfo2 methods
    STDMETHOD(GetDescriptionEx)(BSTR* pbstrDescription, BSTR* pbstrSolution);

    // Methods and members for setting error info
    void         SetTextV(EPART epart, UINT ids, void* pvArgs);
    void __cdecl SetText(EPART epart, UINT ids, ...);

    HRESULT     _hr;
    TCHAR*      _apch[EPART_LAST];

    DWORD       _dwHelpContext;     // Help context in Forms3 help file
    GUID        _clsidSource;       // Used to generate progid for GetSource

    IID         _iidInvoke;         // Used to generate action part of message
    DISPID      _dispidInvoke;      //      if set and _apch[EPART_ACTION]
    INVOKEKIND  _invkind;           //      not set.

    // Private helper functions
private:
    HRESULT GetMemberName(BSTR* pbstrName);
};

CErrorInfo*  GetErrorInfo();
void         ClearErrorInfo(THREADSTATE* pts=NULL);
void         CloseErrorInfo(HRESULT hr, REFCLSID clsidSource);
void __cdecl PutErrorInfoText(EPART epart, UINT ids, ...);

HRESULT      GetErrorText(HRESULT hrError, LPTSTR pstr, int cch);

#endif //__XINDOWS_CORE_COM_ERRORINFO_H__