
#ifndef __XINDOWS_SITE_TEXT_TREEINFO_H__
#define __XINDOWS_SITE_TEXT_TREEINFO_H__

class COneRun;

class CTreeInfo
{
public:
    CTreeInfo(CMarkup* pMarkup) : _tpFrontier(pMarkup) { _fInited = FALSE; };
    HRESULT InitializeTreeInfo(CFlowLayout* pFlowLayout, BOOL fIsEditable, LONG cp, CTreePos* ptp);
    void SetupCFPF(BOOL fIniting, CTreeNode* pNode);
    BOOL AdvanceTreePos();
    BOOL AdvanceTxtPtr();
    void AdvanceFrontier(COneRun* por);

    CTreePos*           _ptpFrontier;
    long                _cchRemainingInTreePos;

    BOOL                _fHasNestedElement;
    BOOL                _fHasNestedLayout;
    BOOL                _fHasNestedRunOwner;
    const CParaFormat*  _pPF;
    long                _iPF;
    BOOL                _fInnerPF;
    const CCharFormat*  _pCF;
    BOOL                _fInnerCF;
    const CFancyFormat* _pFF;
    long                _iFF;

    long                _lscpFrontier;
    long                _cpFrontier;
    long                _chSynthsBefore;

    // Following are the text related state variables
    CTxtPtr            _tpFrontier;
    long               _cchValid;
    const TCHAR       *_pchFrontier;

    CMarkup            *_pMarkup;
    CFlowLayout        *_pFlowLayout;
    LONG                _cpLayoutFirst;
    LONG                _cpLayoutLast;
    CTreePos           *_ptpLayoutLast;
    BOOL                _fIsEditable;

    BOOL _fInited;
};

#endif //__XINDOWS_SITE_TEXT_TREEINFO_H__