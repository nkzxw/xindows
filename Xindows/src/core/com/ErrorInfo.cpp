
#include "stdafx.h"
#include "ErrorInfo.h"

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::CErrorInfo
//
//  Synopsis:   Constructor
//
//----------------------------------------------------------------------------
CErrorInfo::CErrorInfo()
{
    _ulRefs = 1;
}

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::~CErrorInfo
//
//  Synopsis:   Destructor
//
//----------------------------------------------------------------------------
CErrorInfo::~CErrorInfo()
{
    int c;
    TCHAR** ppch;

    for(ppch=_apch,c=EPART_LAST; c>0; ppch++,c--)
    {
        MemFree(*ppch);
    }
}

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::QueryInterface, IUnknown
//
//----------------------------------------------------------------------------
HRESULT CErrorInfo::QueryInterface(REFIID iid, void** ppv)
{
    if(iid==IID_IErrorInfo || iid==IID_IUnknown)
    {
        *ppv = (IErrorInfo*)this;
        ((IUnknown*)*ppv)->AddRef();
        return S_OK;
    }
    else
    {
        return E_NOINTERFACE;
    }
}

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::SetTextV/SetText
//
//  Synopsis:   Set text of error description.
//
//  Arguments:  epart   Part of error message, taken from EPART enum.
//              ids     String id of format string.
//              pvArgs  Arguments
//
//----------------------------------------------------------------------------
void CErrorInfo::SetTextV(EPART epart, UINT ids, void* pvArgs)
{
    if(_apch[epart])
    {
        MemFree(_apch[epart]);
        _apch[epart] = NULL;
    }

    VFormat(_afxGlobalData._hInstResource, FMT_OUT_ALLOC, &_apch[epart], 0, MAKEINTRESOURCE(ids), pvArgs);
}

void __cdecl CErrorInfo::SetText(EPART epart, UINT ids, ...)
{
    va_list arg;
    va_start(arg, ids);
    SetTextV(epart, ids, &arg);
    va_end(arg);    
}

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::GetGUID, IErrorInfo
//
//  Synopsis:   Return iid of interface defining error code.
//
//----------------------------------------------------------------------------
HRESULT CErrorInfo::GetGUID(GUID* pguid)
{
    // Assume OS defined errors only.
    *pguid = _afxGlobalData._Zero.clsid;
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::GetSource, IErrorInfo
//
//  Synopsis:   Get progid of error source.
//
//----------------------------------------------------------------------------
HRESULT CErrorInfo::GetSource(BSTR* pbstrSource)
{
    HRESULT hr; 
    OLECHAR* pstrProgID = NULL;

    *pbstrSource = 0;
    if(_clsidSource == _afxGlobalData._Zero.clsid)
    {
        RRETURN(E_FAIL);
    }

    hr = ProgIDFromCLSID(_clsidSource, &pstrProgID);
    if(hr)
    {
        goto Cleanup;
    }

    hr = FormsAllocString(pstrProgID, pbstrSource);

Cleanup:
    CoTaskMemFree(pstrProgID);
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::GetHelpFile, IErrorInfo
//
//  Synopsis:   Get help file describing error.
//
//----------------------------------------------------------------------------
HRESULT CErrorInfo::GetHelpFile(BSTR* pbstrHelpFile)
{
    RRETURN(E_NOTIMPL);
}

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::GetHelpContext, IErrorInfo
//
//  Synopsis:   Get help context describing error.
//
//----------------------------------------------------------------------------
HRESULT CErrorInfo::GetHelpContext(DWORD* pdwHelpContext)
{
    *pdwHelpContext = _dwHelpContext;
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::GetDescription, IErrorInfo
//
//  Synopsis:   Get description of the error.
//
//----------------------------------------------------------------------------
HRESULT CErrorInfo::GetDescription(BSTR* pbstrDescription)
{
    HRESULT hr;
    BSTR    bstrSolution = NULL;
    BSTR    bstrDescription = NULL;

    hr = GetDescriptionEx(&bstrDescription, &bstrSolution);
    if(hr)
    {
        goto Cleanup;
    }

    if(bstrDescription && bstrSolution)
    {
        hr = FormsAllocStringLen(
            (LPCTSTR)NULL, 
            FormsStringLen(bstrDescription)+2+FormsStringLen(bstrSolution),
            pbstrDescription);
        if(hr)
        {
            goto Cleanup;
        }

        _tcscpy(*pbstrDescription, bstrDescription);
        _tcscat(*pbstrDescription, TEXT("  ")); // Review: FE grammar?
        _tcscat(*pbstrDescription, bstrSolution);
    }
    else
    {
        *pbstrDescription = bstrDescription;
        bstrDescription = NULL;
    }

Cleanup:
    FormsFreeString(bstrDescription);
    FormsFreeString(bstrSolution);

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::GetMemberName
//
//  Synopsis:   Get name of member _dispidInvoke in interface _iidInvoke.
//
//----------------------------------------------------------------------------
HRESULT CErrorInfo::GetMemberName(BSTR* pbstrName)
{
    RRETURN(E_FAIL);
}

//+---------------------------------------------------------------------------
//
//  Method:     CErrorInfo::GetDescriptionEx, IErrorInfo2
//
//  Synopsis:   
//
//  Arguments:  
//
//----------------------------------------------------------------------------
HRESULT CErrorInfo::GetDescriptionEx(BSTR* pbstrDescription, BSTR* pbstrSolution)
{
    HRESULT hr = S_OK;
    TCHAR   achBufAction[FORMS_BUFLEN];
    TCHAR   achBufError[FORMS_BUFLEN];
    TCHAR*  pchAction;
    TCHAR*  pchError;
    TCHAR*  pchSolution;
    BSTR    bstrMemberName = NULL;

    *pbstrDescription = NULL;
    *pbstrSolution = NULL;

    if(_apch[EPART_ACTION])
    {
        pchAction = _apch[EPART_ACTION];
    }
    else if(_invkind==INVOKE_PROPERTYPUT && _hr==E_INVALIDARG)
    {
        pchAction = NULL;
    }
    else if(_invkind)
    {
        hr = GetMemberName(&bstrMemberName);
        if(hr)
        {
            goto Error;
        }

        hr = Format(
            _afxGlobalData._hInstResource,
            0, 
            achBufAction, 
            ARRAYSIZE(achBufAction),
            _invkind==INVOKE_PROPERTYPUT ? 
            _T("Could not set property <0s>")
            :(_invkind==INVOKE_PROPERTYGET ? 
            _T("Could not get property <0s>")
            :_T("Could not call method <0s>")),
            bstrMemberName);
        if(hr)
        {
            goto Error;
        }

        pchAction = achBufAction;
    }
    else
    {
        pchAction = NULL;
    }

    if(_apch[EPART_ERROR])
    {
        pchError = _apch[EPART_ERROR];
    }
    else if(_invkind==INVOKE_PROPERTYPUT && _hr==E_INVALIDARG)
    {
        hr = GetMemberName(&bstrMemberName);
        if(hr)
        {
            goto Error;
        }

        hr = Format(
            _afxGlobalData._hInstResource,
            0, 
            achBufError, 
            ARRAYSIZE(achBufError),
            _T("The value entered is not valid <0s>"),
            bstrMemberName);
        if(hr)
        {
            goto Error;
        }

        pchError = achBufError;
    }
    else
    {
        hr = GetErrorText(_hr, achBufError, ARRAYSIZE(achBufError));
        if(hr)
        {
            goto Error;
        }

        pchError = achBufError;
    }

    if(pchAction)
    {
        hr = FormsAllocStringLen(
            pchAction,
            _tcslen(pchAction)+_tcslen(pchError)+2,
            pbstrDescription);
        if(!*pbstrDescription)
        {
            goto MemoryError;
        }
        _tcscat(*pbstrDescription, TEXT(" "));
        _tcscat(*pbstrDescription, pchError);
    }
    else
    {
        hr = FormsAllocString(pchError,pbstrDescription);
        if(!*pbstrDescription)
        {
            goto MemoryError;
        }
    }

    if(_apch[EPART_SOLUTION])
    {
        pchSolution = _apch[EPART_SOLUTION];
    }
    else
    {
        pchSolution = NULL;
    }

    hr = FormsAllocString(pchSolution, pbstrSolution);
    if(hr)
    {
        goto Error;
    }

Cleanup:
    SysFreeString(bstrMemberName);
    RRETURN(hr);

MemoryError:
    hr = E_OUTOFMEMORY;

Error:
    SysFreeString(*pbstrDescription);
    SysFreeString(*pbstrSolution);
    *pbstrDescription = NULL;
    *pbstrSolution = NULL;
    goto Cleanup;
}

//+---------------------------------------------------------------------------
//
//  Function:   GetErrorInfo
//
//----------------------------------------------------------------------------
CErrorInfo* GetErrorInfo()
{
    THREADSTATE* pts = GetThreadState();

    if(!pts->pEI)
    {
        pts->pEI = new CErrorInfo;
    }

    return pts->pEI;
}

//+---------------------------------------------------------------------------
//
//  Function:   ClearErrorInfo
//
//----------------------------------------------------------------------------
void ClearErrorInfo(THREADSTATE* pts)
{
    if(!pts)
    {
        pts = GetThreadState();
    }
    ClearInterface(&(pts->pEI));
}

//+---------------------------------------------------------------------------
//
//  Function:   CloseErrorInfo
//
//----------------------------------------------------------------------------
void CloseErrorInfo(HRESULT hr, REFCLSID clsidSource)
{
    IErrorInfo* pErrorInfo;

    Assert(FAILED(hr));

    if(GetErrorInfo() == NULL)
    {
        // There's an error, but we couldn't allocate
        // an error info object.  Release the global error
        // object so that our caller's caller will
        // not be confused.
        hr = GetErrorInfo(0, &pErrorInfo);
        if(!hr)
        {
            ReleaseInterface(pErrorInfo);
        }
    }
    else
    {
        THREADSTATE* pts = GetThreadState();
        pts->pEI->_hr = hr;
        pts->pEI->_clsidSource = clsidSource;
        SetErrorInfo(0, pts->pEI);

        ClearInterface(&(pts->pEI));
    }
}

//+---------------------------------------------------------------------------
//
//  Function:   PutErrorInfoText
//
//----------------------------------------------------------------------------
void PutErrorInfoText(EPART epart, UINT ids, ...)
{
    va_list arg;

    if(GetErrorInfo() != NULL)
    {
        va_start(arg, ids);
        TLS(pEI)->SetTextV(epart, ids, &arg);
        va_end(arg);
    }
}

// Each entry in the table has the following format:
//     StartError, NumberOfErrorsToMap, ErrorMessage
//
// The last entry must be all 0s.
struct ERRTOMSG
{
    SCODE   error;      // the starting error code
    USHORT  cErrors;    // number of errors in range to map
    TCHAR*  message;    // the error msg
};

// Disable warning: Truncation of constant value.  There is
// no way to specify a short integer constant.
ERRTOMSG g_aetmError[] =
{
    E_UNEXPECTED,                       1,  _T("����ص����˷��������Է��ʡ�"),
    E_FAIL,                             1,  _T("δָ���Ĵ���"),
    E_INVALIDARG,                       1,  _T("������Ч��"),
    CTL_E_CANTMOVEFOCUSTOCTRL,          1,  _T("���ڸÿؼ�Ŀǰ���ɼ���δ���û������Ͳ���������޷���������������"),
    CTL_E_CONTROLNEEDSFOCUS,            1,  _T("�ÿؼ���Ҫ��ȷ���Ľ��㡣"),
    CTL_E_INVALIDPICTURE,               1,  _T("���ֵ����ͼƬ����"),
    CTL_E_INVALIDPICTURETYPE,           1,  _T("��ͼƬ������Ч��"),
    CTL_E_INVALIDPROPERTYARRAYINDEX,    1,  _T("��Ч�����������±ꡣ"),
    CTL_E_INVALIDPROPERTYVALUE,         1,  _T("��Ч������ֵ��"),
    CTL_E_METHODNOTAPPLICABLE,          1,  _T("�÷����޷�������һ�������С�"),
    CTL_E_OVERFLOW,                     1,  _T("���"),
    CTL_E_PERMISSIONDENIED,             1,  _T("Ȩ�ޱ��ܾ���"),
    CTL_E_SETNOTSUPPORTEDATRUNTIME,     1,  _T("����������ʱ�������ԡ�"),
    CTL_E_INVALIDPASTETARGET,           1,  _T("�ò�����Ŀ��Ԫ����Ч��"),
    CTL_E_INVALIDPASTESOURCE,           1,  _T("���ڴ��������ԭʼ HTML ��Ч��"),
    CTL_E_MISMATCHEDTAG,                1,  _T("���ĵ�������ƥ��ı�ǣ����δ����ȫ���ء�"),
    CTL_E_INCOMPATIBLEPOINTERS,         1,  _T("���ָ����˲��������ݡ�"),
    CTL_E_UNPOSITIONEDPOINTER,          1,  _T("�˲������õı��ָ���λ��δ֪��"),
    CTL_E_UNPOSITIONEDELEMENT,          1,  _T("�ò�������Ԫ�ص�λ��δ֪��"),
    CLASS_E_NOTLICENSED,                1,  _T("����û�еõ���ȷ��Ȩ������޷������ÿؼ���"),
    MSOCMDERR_E_NOTSUPPORTED,           1,  _T("��֧�ָ����"),
    HRESULT_FROM_WIN32(ERROR_INTERNET_INVALID_URL),       1, _T("�õ�ַ��Ч�������ַ��Ȼ������һ�Ρ�"),
    HRESULT_FROM_WIN32(ERROR_INTERNET_NAME_NOT_RESOLVED), 1, _T("δ�ҵ�վ�㡣��ȷ�ϵ�ַ�Ƿ���ȷ��Ȼ������һ�Ρ�"),
    // Terminate table with nulls.
    0,                                  0,  0,
};

//+---------------------------------------------------------------------------
//
//  Function:   GetErrorText
//
//  Synopsis:   Gets the text of an error from a message resource.
//
//  Arguments:  [hr]   --   The error.  Must not be 0.
//              [pstr] --   Buffer for the message.
//              [cch]  --   Size of the buffer in characters.
//
//  Returns:    HRESULT.
//
//  Modifies:   [*pstr]
//
//----------------------------------------------------------------------------
HRESULT GetErrorText(HRESULT hrError, LPTSTR pstr, int cch)
{
    HRESULT     hr  = S_OK;
    ERRTOMSG*   petm;
    DWORD       dwFlags;
    LPCVOID     pvSource;
    BOOL        fHrCode = FALSE;

    Assert(pstr);
    Assert(cch >= 0);

    for(petm=g_aetmError; petm->error; petm++)
    {
        if(petm->error<=hrError && hrError<petm->error+petm->cErrors)
        {
            _tcscpy(pstr, petm->message);
            return S_OK;
        }
    }

    if(hrError>=HRESULT_FROM_WIN32(INTERNET_ERROR_BASE) &&
        hrError<=HRESULT_FROM_WIN32(INTERNET_ERROR_LAST))
    {
        dwFlags = FORMAT_MESSAGE_FROM_HMODULE | FORMAT_MESSAGE_IGNORE_INSERTS;
        pvSource = GetModuleHandleA("wininet.dll");
        fHrCode = TRUE;
    }
    else if(hrError>=INET_E_ERROR_FIRST && hrError<=INET_E_ERROR_LAST)
    {
        dwFlags = FORMAT_MESSAGE_FROM_HMODULE | FORMAT_MESSAGE_IGNORE_INSERTS;
        pvSource = GetModuleHandleA("urlmon.dll");
    }
    else
    {
        dwFlags = FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
        pvSource = NULL;
    }

    if(FormatMessage(
        dwFlags,
        pvSource,
        fHrCode?HRESULT_CODE(hrError):hrError,
        LANG_SYSTEM_DEFAULT,
        pstr,
        cch,
        NULL))
    {
        return S_OK;
    }

    // Show the error string and the hex code.
    hr = Format(_afxGlobalData._hInstResource, 0, pstr, cch, _T("���ڳ��ִ��� <0x> �����´�������޷���ɡ�"), hrError);
    RRETURN(hr);
}