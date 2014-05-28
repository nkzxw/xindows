
#ifndef __XINDOWS_SITE_BASE_PROGSINK_H__
#define __XINDOWS_SITE_BASE_PROGSINK_H__

struct PROGDATA
{
    BYTE    bClass;         // See PROGSINK_CLASS_* definition
    BYTE    bState;         // See PROGSINK_STATE_* definition
    BYTE    bBelow;         // Progress is coming from sub-frame
    BYTE    bFlags;         // See PDF_* below
    DWORD   dwPos;          // Thermometer position
    DWORD   dwMax;          // Thermometer max
    TCHAR*  pchText;        // Status text
    DWORD   dwIds;          // Resource string for formatting
    TCHAR*  pchFormat;      // Formatted progress string
};

#define PDF_FREE        0x01
#define PDF_NOSPIN      0x02

class CProgSink : public CBaseFT, public IProgSink
{
public:
    typedef CBaseFT super;

    DECLARE_MEMCLEAR_NEW_DELETE()

    // Methods called only on the creation thread
    CProgSink(CDocument* pDoc, CMarkup* pMarkup);
    ~CProgSink();

    virtual void Passivate();
    void Detach();
    void SetCounters(DWORD dwClass, BOOL fFwd, BOOL fAdd, BOOL fSpin);

    BOOL UpdateProgress(PROGDATA* ppd, DWORD dwFlags, DWORD dwPos, DWORD dwMax);
    void AdjustProgress(DWORD dwClass, BOOL fFwd, BOOL fAdd);

    void Signal(UINT wChange);
    NV_DECLARE_ONCALL_METHOD(OnMethodCall, onmethodcall, (DWORD_PTR dwContext));

    UINT GetClassCounter(DWORD dwClass, BOOL fBelow=FALSE);

    // IUnknown methods (called on any thread)

    STDMETHOD(QueryInterface)(REFIID iid, void** ppv);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IProgSink methods (called on any thread)

    STDMETHOD(AddProgress)(DWORD dwClass, DWORD* pdwCookie);
    STDMETHOD(SetProgress)(DWORD dwCookie, DWORD dwFlags,
        DWORD dwState, LPCTSTR pchText, DWORD dwIds, DWORD dwPos, DWORD dwMax);
    STDMETHOD(DelProgress)(DWORD dwCookie);

private:

    WORD                _fPassive:1;
    WORD                _fSentSignal:1;
    WORD                _fSendParseDone:1;
    WORD                _fSentParseDone:1;
    WORD                _fSentQuickDone:1;
    WORD                _fSendProgDone:1;
    WORD                _fSentDone:1;
    WORD                _fSpin:1;
    WORD                _fGotDefault:1;
    WORD                _fShowItemsRemaining:1;

    BYTE                _bBelowCur;
    BYTE                _bStateCur;
    DWORD               _dwSlotCur;
    PROGDATA            _pdDefault;
    CDocument*          _pDoc;
    CMarkup*            _pMarkup;
    IProgSink*          _pProgSinkFwd;
    THREADSTATE*        _pts;

    DECLARE_CDataAry(CAryProg, PROGDATA)
    CAryProg            _aryProg;

    UINT                _cFree;
    UINT                _cActive;
    UINT                _cSpin;
    UINT                _acClass[PROGSINK_CLASS_FRAME+1];
    UINT                _acClassBelow[PROGSINK_CLASS_FRAME+1];
    UINT                _uChange;

    DWORD               _dwProgMax;    // next max pos on the progress bar
    DWORD               _dwProgCur;    // current pos on the progress bar
    DWORD               _dwProgDelta;  // progress delta
    DWORD               _dwMaxTotal;   // sum of all valid max 
    UINT                _cOngoing;     // number of ongoing downloads

    UINT                _cPotential;   // new download potential
    UINT                _cPotTotal;    // total download potential (new + consumed)
    UINT                _cPotDelta;    // download potential currently consumed

    CRITICAL_SECTION    _cs;
};

#endif //__XINDOWS_SITE_BASE_PROGSINK_H__