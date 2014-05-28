
#ifndef __XINDOWS_SITE_BASE_XBAG_H__
#define __XINDOWS_SITE_BASE_XBAG_H__

enum
{
    ICF_EMBEDDEDOBJECT,
    ICF_EMBEDSOURCE,
    ICF_LINKSOURCE,
    ICF_LINKSRCDESCRIPTOR,
    ICF_OBJECTDESCRIPTOR,
    ICF_FORMSCLSID,
    ICF_FORMSTEXT
};
// Array of common clip formats, indexed by ICF_xxx enum
extern CLIPFORMAT g_acfCommon[];

// Initialize g_cfStandard.
void RegisterClipFormats();

#define CF_COMMON(icf)  (CLIPFORMAT)(CF_PRIVATEFIRST+icf)

// FORMATETC helpers
HRESULT CloneStgMedium(const STGMEDIUM* pcstgmedSrc, STGMEDIUM* pstgmedDest);
BOOL    DVTARGETDEVICEMatchesRequest(const DVTARGETDEVICE* pcdvtdRequest, const DVTARGETDEVICE* pcdvtdActual);
BOOL    TYMEDMatchesRequest(TYMED tymedRequest, TYMED tymedActual);
BOOL    FORMATETCMatchesRequest(const FORMATETC* pcfmtetcRequest, const FORMATETC* pcfmtetcActual);
int     FindCompatibleFormat(FORMATETC FmtTable[], int iSize, FORMATETC& formatetc);

// OBJECTDESCRIPTOR clipboard format helpers
HRESULT GetObjectDescriptor(LPDATAOBJECT pDataObj, LPOBJECTDESCRIPTOR pDescOut);
HRESULT UpdateObjectDescriptor(LPDATAOBJECT pDataObj, POINTL& ptl, DWORD dwAspect);

//+-------------------------------------------------------------------------
//
//  Transfer: bagged list
//
//--------------------------------------------------------------------------
typedef struct
{
    RECTL rclBounds;
    ULONG ulID;
} CTRLABSTRACT, FAR *LPCTRLABSTRACT;

#define CLSID_STRLEN    38

HRESULT FindLegalCF(IDataObject* pDO);
HRESULT GetcfCLSIDFmt(LPDATAOBJECT pDataObj, TCHAR* tszClsid);


//////////////////////////////////////////////////////////////////////////////////////////////
//
// CDropSource
// Implements IDropSource interface.
class CDropSource : public IDropSource
{
private:
    DECLARE_MEMALLOC_NEW_DELETE()
public:
    STDMETHOD(QueryContinueDrag)(BOOL fEscapePressed, DWORD grfKeyState);
    STDMETHOD(GiveFeedback)(DWORD dwEffect);
protected:
    DWORD       _dwButton;  // Mouse btn state info for dragging
    CDocument*  _pDoc;
};

//////////////////////////////////////////////////////////////////////////////////////////////
//
// CDummyDropSource
// Wraps a trivial instantiable class around CDropSource
class CDummyDropSource : public CDropSource
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()
    DECLARE_STANDARD_IUNKNOWN(CDummyDropSource);
    static HRESULT Create(DWORD dwKeyState, CDocument* pDoc, IDropSource** ppDropSrc);
private:
    CDummyDropSource() { _ulRefs = 1; }
};

// command IDs used by IOleCommandTarget interface, for programatic paste
// security check
#define IDM_SETSECURITYDOMAIN   1000
#define IDM_CHECKSECURITYDOMAIN 1001

// An abstract class for all the classes that implement
// IDataObject (the text xbag,  the 2d xbag and the
// generic data object)
class CBaseBag : public CDropSource, public IDataObject, public IOleCommandTarget
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()

    DECLARE_STANDARD_IUNKNOWN(CBaseBag);

    //IDataObject
    STDMETHODIMP DAdvise(
        FORMATETC FAR* pFormatetc,
        DWORD advf,
        LPADVISESINK pAdvSink,
        DWORD FAR* pdwConnection) { return OLE_E_ADVISENOTSUPPORTED; }
    STDMETHODIMP DUnadvise( DWORD dwConnection) { return OLE_E_ADVISENOTSUPPORTED; }
    STDMETHODIMP EnumDAdvise(LPENUMSTATDATA FAR* ppenumAdvise) { return OLE_E_ADVISENOTSUPPORTED; }

    STDMETHODIMP EnumFormatEtc(
        DWORD dwDirection,
        LPENUMFORMATETC FAR* ppenumFormatEtc) { return E_NOTIMPL; }

    STDMETHODIMP GetCanonicalFormatEtc(
        LPFORMATETC pformatetc,
        LPFORMATETC pformatetcOut)
    {
        pformatetcOut->ptd = NULL;
        return E_NOTIMPL;
    }

    STDMETHODIMP GetData(LPFORMATETC pformatetcIn, LPSTGMEDIUM pmedium)
    {
        return E_NOTIMPL;
    }
    STDMETHODIMP GetDataHere(LPFORMATETC pformatetc, LPSTGMEDIUM pmedium)
    {
        return E_NOTIMPL;
    }
    STDMETHODIMP QueryGetData(LPFORMATETC pformatetc )
    {
        return E_NOTIMPL;
    }
    STDMETHODIMP SetData(LPFORMATETC pformatetc, STGMEDIUM FAR* pmedium, BOOL fRelease)
    {
        return E_NOTIMPL;
    }

    // IOleCommandTarget
    virtual HRESULT STDMETHODCALLTYPE QueryStatus(
        const GUID* pguidCmdGroup,
        ULONG cCmds,
        OLECMD rgCmds[],
        OLECMDTEXT* pcmdtext);
    virtual HRESULT STDMETHODCALLTYPE Exec(
        const GUID* pguidCmdGroup,
        DWORD nCmdID,
        DWORD nCmdexecopt,
        VARIANTARG* pvarargIn,
        VARIANTARG* pvarargOut);

    // Other
    virtual HRESULT DeleteObject() { return E_FAIL; }

    IDataObject* _pLinkDataObj;
    BOOL _fLinkFormatAppended;

    virtual ~CBaseBag()
    {
        if(_pLinkDataObj)
        {
            _pLinkDataObj->Release();
            _pLinkDataObj = NULL;
        }
    }
    BYTE _abSID[MAX_SIZE_SECURITY_ID];
};

class CGenDataObject : public CBaseBag
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()
    CGenDataObject(CDocument* pDoc);
    ~CGenDataObject();

    void SetPreferredEffect(DWORD dwPreferredEffect) { _dwPreferredEffect = dwPreferredEffect; }

    DWORD GetPreferredEffect() const { return _dwPreferredEffect; }

    HRESULT AppendFormatData(CLIPFORMAT cfFormat, HGLOBAL hGlobal);
    HRESULT DeleteFormatData(CLIPFORMAT cfFormat);

    //IDataObject
    STDMETHODIMP EnumFormatEtc(DWORD dwDirection, LPENUMFORMATETC FAR* ppenumFormatEtc);
    STDMETHODIMP GetData(LPFORMATETC pformatetcIn, LPSTGMEDIUM pmedium);
    STDMETHODIMP QueryGetData(LPFORMATETC pformatetc);
    STDMETHODIMP SetData(LPFORMATETC pformatetc, STGMEDIUM FAR* pmedium, BOOL fRelease);

    void SetBtnState(DWORD dwKeyState);

    CDataAry<FORMATETC> _rgfmtc;
    CDataAry<STGMEDIUM> _rgstgmed;

private:
    FORMATETC _fmtcPreferredEffect;
    DWORD     _dwPreferredEffect;
};

/*
*  class CEnumFormatEtc
*
*  Purpose:
*      implements a generic format enumerator for IDataObject
*/
class CEnumFormatEtc : public IEnumFORMATETC
{
public:
    STDMETHOD(QueryInterface)(REFIID riid, void** ppvObj);
    STDMETHOD_(ULONG,AddRef)(void);
    STDMETHOD_(ULONG,Release)(void);

    STDMETHOD(Next)(ULONG celt, FORMATETC* rgelt, ULONG* pceltFetched);
    STDMETHOD(Skip)(ULONG celt);
    STDMETHOD(Reset)(void);
    STDMETHOD(Clone)(IEnumFORMATETC** ppenum);

    static HRESULT Create(FORMATETC* prgFormats, DWORD cFormats, IEnumFORMATETC** ppenum);

private:
    DECLARE_MEMALLOC_NEW_DELETE()

    CEnumFormatEtc();
    ~CEnumFormatEtc();

    ULONG       _crefs;
    ULONG       _iCurrent;      // current clipboard format
    ULONG       _cTotal;        // total number of formats
    FORMATETC*  _prgFormats;    // array of available formats
};

#endif //__XINDOWS_SITE_BASE_XBAG_H__