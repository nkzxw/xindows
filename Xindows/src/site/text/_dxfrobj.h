
#ifndef __XINDOWS_SITE_TEXT_DXFROBJ_H__
#define __XINDOWS_SITE_TEXT_DXFROBJ_H__

class CSelDragDropSrcInfo;

typedef enum tagDataObjectInfo
{
    DOI_NONE            = 0,
    DOI_CANUSETOM       = 1,    // TOM<-->TOM optimized data transfers
    DOI_CANPASTEPLAIN   = 2,    // plain text pasting available
    DOI_CANPASTERICH    = 4,    // rich text pasting available  (TODO: alexgo)
    DOI_CANPASTEOLE     = 8     // object may be pasted as an OLE embedding
    // (note that this flag may be combined with
    // others). (TODO: alexgo)
    //TODO (alexgo): more possibilites:  CANPASTELINK, CANPASTESTATICOLE
} DataObjectInfo;

/*
*  CTextXBag
*
*  Purpose:
*      holds a "snapshot" of some text data that can be used
*      for drag drop or clipboard operations
*
*  Notes:
*      TODO (alexgo): add in support for TOM<-->TOM optimized data
*      transfers
*/
class CTextXBag : public CBaseBag
{
    typedef CBaseBag super;

public:
    STDMETHODIMP QueryInterface(REFIID iid, LPVOID* ppv);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    STDMETHOD(EnumFormatEtc)(DWORD dwDirection, IEnumFORMATETC** ppenumFormatEtc);
    STDMETHOD(GetData)(FORMATETC* pformatetcIn, STGMEDIUM* pmedium);
    STDMETHOD(QueryGetData)(FORMATETC* pformatetc);
    STDMETHOD(SetData)(LPFORMATETC pformatetc, STGMEDIUM FAR* pmedium, BOOL fRelease);

    static HRESULT Create(
        CMarkup*                pMarkup,
        BOOL                    fSupportsHTML,
        ISegmentList*           pSegmentList, 
        BOOL                    fDragDrop,
        CTextXBag**             ppTextXBag,
        CSelDragDropSrcInfo*    pSelDragDropSrcInfo=NULL);

    static HRESULT  GetDataObjectInfo(IDataObject* pdo, DWORD* pDOIFlags);

private:
    // NOTE: private cons/destructor, may not be allocated on the stack as
    // this would break OLE's current object liveness rules
    CTextXBag();
    virtual ~CTextXBag();

    HRESULT SetKeyState();
    HRESULT FillWithFormats(
        CMarkup*        pMarkup,
        BOOL            fSupportsHTML,
        ISegmentList*   pSegmentList);

    HRESULT SetTextHelper(
        CMarkup*        pMarkup,
        ISegmentList*   pSegmentList,
        DWORD           dwSaveHtmlFlags,
        CODEPAGE        cp,
        DWORD           dwStmWrBuffFlags,
        HGLOBAL*        phGlobalText,
        int             iFETCIndex);

    HRESULT SetText(
        CMarkup*        pMarkup,
        BOOL            fSupportsHTML,
        ISegmentList*   pSegmentList);

    HRESULT SetUnicodeText(
        CMarkup*        pMarkup,
        BOOL            fSupportsHTML,
        ISegmentList*   pSegmentList);

    HRESULT SetCFHTMLText(
        CMarkup*        pMarkup,
        BOOL            fSupportsHTML,
        ISegmentList*   pSegmentList);

    HRESULT GetHTMLText(
        HGLOBAL*        phGlobal, 
        ISegmentList*   pSegmentList,
        CMarkup*        pMarkup, 
        DWORD           dwSaveHtmlMode,
        CODEPAGE        codepage, 
        DWORD           dwStrWrBuffFlags);

    DECLARE_MEMCLEAR_NEW_DELETE()

    long        _cFormatMax;    // maximum formats the array can store
    long        _cTotal;        // total number of formats in the array
    FORMATETC*  _prgFormats;    // the array of supported formats
    CSelDragDropSrcInfo* _pSelDragDropSrcInfo;

public:
    HGLOBAL     _hText;         // handle to the ansi plain text
    HGLOBAL     _hUnicodeText;  // handle to the plain UNICODE text
    HGLOBAL     _hRTFText;      // handle to the RTF text
    HGLOBAL     _hCFHTMLText;   // handle to the CFHTML (in utf-8)
};


//  Some globally useful FORMATETCs
extern FORMATETC    g_rgFETC[];
extern DWORD        g_rgDOI[];
extern int          CFETC;

enum FETCINDEX                          // Keep in sync with g_rgFETC[]
{
    iHTML,                              // HTML (in ANSI)
    iRtfFETC,                           // RTF
    iUnicodeFETC,                       // Unicode plain text
    iAnsiFETC,                          // ANSI plain text
    iRtfAsTextFETC,                     // Pastes RTF as text
    iFileDescA,                         // FileGroupDescriptor
    iFileDescW,                         // FileGroupDescriptorW
    iFileContents,                      // FileContents
    iUniformResourceLocator             // UniformResourceLocator
};

#define cf_HTML                     g_rgFETC[iHTML].cfFormat
#define cf_RTF                      g_rgFETC[iRtfFETC].cfFormat
#define cf_RTFASTEXT                g_rgFETC[iRtfAsTextFETC].cfFormat
#define cf_FILEDESCA                g_rgFETC[iFileDescA].cfFormat
#define cf_FILEDESCW                g_rgFETC[iFileDescW].cfFormat
#define cf_FILECONTENTS             g_rgFETC[iFileContents].cfFormat
#define cf_UNIFORMRESOURCELOCATOR   g_rgFETC[iUniformResourceLocator].cfFormat

#endif //__XINDOWS_SITE_TEXT_DXFROBJ_H__