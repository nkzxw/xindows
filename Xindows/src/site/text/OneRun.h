
#ifndef __XINDOWS_SITE_TEXT_ONERUN_H__
#define __XINDOWS_SITE_TEXT_ONERUN_H__

#include "LineServices.h"

class COneRun;

class CComplexRun
{
public:
    CComplexRun()
    {
        _ThaiWord.cp = 0;
        _ThaiWord.cch = 0;
        _pGlyphBuffer = NULL;
    }
    ~CComplexRun()
    {
        MemFree(_pGlyphBuffer);
    }
    DECLARE_MEMCLEAR_NEW_DELETE();
    
    void ComputeAnalysis(
        const CFlowLayout* pFlowLayout,
        BOOL fRTL,
        BOOL fForceGlyphing,
        WCHAR chNext,
        WCHAR chPassword,
        COneRun* por,
        COneRun* porTail,
        DWORD uLangDigits);
    BOOL IsNumericSeparatorRun(COneRun* por, COneRun* porTail);
    void ThaiTypeBrkcls(CMarkup* pMarkup, LONG cp, BRKCLS* pbrkclsAfter, BRKCLS* pbrkclsBefore);
    SCRIPT_ANALYSIS* GetAnalysis() { return &_Analysis; }
    WORD GetScript() { return _Analysis.eScript; }

    inline void NukeAnalysis() { _Analysis.eScript = 0; }
    inline void CopyPtr(WORD* pGlyphBuffer)
    {
        if(_pGlyphBuffer)
        {
            delete _pGlyphBuffer;
        }
        _pGlyphBuffer = pGlyphBuffer;
    }

private:
    SCRIPT_ANALYSIS _Analysis;
    struct
    {
        DWORD cp    : 26;
        DWORD cch   : 6;
    } _ThaiWord;
    WORD* _pGlyphBuffer;
};

class COneRun
{
public:
    enum ORTYPES
    {
        OR_NORMAL    = 0x00,
        OR_SYNTHETIC = 0x01,
        OR_ANTISYNTHETIC = 0x02
    };

    enum charWidthClass
    {
        charWidthClassFullWidth,
        charWidthClassHalfWidth,
        charWidthClassCursive,
        charWidthClassUnknown
    };

    ~COneRun()
    {
        Deinit();
    }

    void Deinit();

    LONG Cp()
    {
        return _lscpBase-_chSynthsBefore;
    }
    const CCharFormat* GetCF()
    {
        Assert(_pCF);
        return _pCF;
    }
    CCharFormat* GetOtherCF();
    const CParaFormat* GetPF()
    {
        Assert(_pPF);
        return _pPF;
    }

    const CFancyFormat* GetFF()
    {
        Assert(_pFF);
        return _pFF;
    }

    CComplexRun* GetComplexRun()
    {
        return _pComplexRun;
    }
    CComplexRun* GetNewComplexRun()
    {
        if(_pComplexRun != NULL)
        {
            delete _pComplexRun;
        }
        _pComplexRun = new CComplexRun;
        return _pComplexRun;
    }
    void SetConvertMode(CONVERTMODE cm)
    {
        AssertSz(_bConvertMode==CVM_UNINITED, "Setting cm more than once!");
        _bConvertMode = cm;
    }
    LPTSTR SetCharacter(TCHAR ch)
    {
        _cstrRunChars.Set(&ch, 1);
        return _cstrRunChars;
    }
    LPTSTR SetString(LPTSTR pstr)
    {
        _cstrRunChars.Set(pstr);
        return _cstrRunChars;
    }
    BOOL IsSelected() { return _fSelected; }
    void Selected(CFlowLayout* pFlowLayout, BOOL fSelected=TRUE);
    void ImeHighlight( CFlowLayout* pFlowLayout, DWORD dwImeHighlight);
    BOOL HasBgColor() { return _fHasBackColor; }    // allow override of formats
    BOOL HasTextColor() { return _fHasTextColor; }  // allow override of formats
    void SetCurrentBgColor(CFlowLayout* pFlowLayout);
    CColorValue GetCurrentBgColor() { return _ccvBackColor; }
    void SetCurrentTextColor(CFlowLayout* pFlowLayout);
    charWidthClass GetCharWidthClass();
    BOOL IsSameGridClass(COneRun* por);
    BOOL IsOneCharPerGridCell();

    CLayout* GetLayout(CFlowLayout* pFlowLayout)
    {
        CTreeNode* pNode = Branch();
        CLayout* pLayout = pNode->NeedsLayout() ? pNode->GetUpdatedLayout() : pFlowLayout;
        Assert(pLayout && pLayout == pNode->GetUpdatedNearestLayout());
        return pLayout;
    }
    // BUGBUG SLOWBRANCH: GetBranch is **way** too slow to be used here.
    CTreeNode* Branch() { return _ptp->GetBranch(); }
    COneRun* Clone(COneRun* porClone);
    void SetSidFromTreePos(const CTreePos* ptp)
    {
        _sid = (_ptp&&_ptp->IsText()) ? _ptp->Sid() : sidDefault;
    }

    BOOL IsNormalRun()        { return _fType==OR_NORMAL;}
    BOOL IsSyntheticRun()     { return _fType==OR_SYNTHETIC; }
    BOOL IsAntiSyntheticRun() { return _fType==OR_ANTISYNTHETIC; }
    BOOL WantsLSCPStop()      { return (_fType!=OR_SYNTHETIC || CLineServices::s_aSynthData[_synthType].fLSCPStop); }

    void MakeRunNormal()        {_fType = OR_NORMAL; }
    void MakeRunSynthetic()     {_fType = OR_SYNTHETIC; }
    void MakeRunAntiSynthetic() {_fType = OR_ANTISYNTHETIC; }
    void FillSynthData(CLineServices::SYNTHTYPE synthtype);

public:
    COneRun* _pNext;
    COneRun* _pPrev;

    LONG _lscpBase;                     // The lscp of this run
    LONG _lscch;                        // The number of characters in this run
    LONG _lscchOriginal;                // The number of original characters in this run
    const TCHAR* _pchBase;              // The ptr to the characters
    const TCHAR* _pchBaseOriginal;      // The original ptr to the characters
    LSCHP _lsCharProps;                 // Char props to pass back to LS
    LONG _chSynthsBefore;               // How many synthetics before this run?

    union
    {
        DWORD _dwProps;
        struct
        {
            // These flags are specific to the list and do not mean
            // anything semantically
            unsigned int _fType: 2;             // Normal/Synthetic/Antisynthetic run?
            BOOL _fGlean: 1;                    // Do we need to glean this run at all?
            BOOL _fNotProcessedYet: 1;          // The run has not been processed yet
            BOOL _fCannotMergeRuns: 1;          // Cannot merge runs
            BOOL _fCharsForNestedElement: 1;    // Chars for nested elements
            BOOL _fCharsForNestedLayout: 1;     // Chars for nested layouts
            BOOL _fCharsForNestedRunOwner: 1;   // Run has nested run owners
            BOOL _fMustDeletePcf: 1;            // Did we allocate a pcf for this run?
            BOOL _fMakeItASpace:1;              // Convert into space?

            // These runs have some semantic meaning for LS.
            BOOL _fSelected: 1;                 // Is the run selected
            BOOL _fHasBackColor: 1;             // For selection and IME
            BOOL _fHasTextColor: 1;             // For IME
            BOOL _fHidden: 1;                   // Is the run hidden
            BOOL _fNoTextMetrics: 1;            // This run has no text metrics
            DWORD _dwImeHighlight: 3;           // IME highlight
            BOOL _fUnderlineForIME: 1;          // IME highlight
            DWORD _brkopt:6;                    // Use break options for GetBreakingClasses
            DWORD _csco: 6;                     // CSCOPTION
        };
    };

    // The Formats specific stuff here
    CCharFormat*        _pCF;
    BOOL                _fInnerCF;
    const CParaFormat*  _pPF;
    BOOL                _fInnerPF;
    const CFancyFormat* _pFF;

    CONVERTMODE         _bConvertMode;
    CString             _cstrRunChars;
    CColorValue         _ccvBackColor;
    CColorValue         _ccvTextColor;
    CComplexRun*        _pComplexRun;
    CLineServices::SYNTHTYPE _synthType;

    union
    {
        DWORD _dwProps2;
        struct
        {
            DWORD _sid : 8;
            DWORD _fIsStartOrEndOfObj : 1;
            DWORD _fIsArtificiallyStartedNOBR : 1;
            DWORD _fIsArtificiallyTerminatedNOBR : 1;
            DWORD _dwUnused : 21;
        };
    };

    // Run width specific stuff here.
    // These members are valid only for runs that have grid layout.
    // For all others runs these members are set to 0.
    LONG _xWidth;
    LONG _xDrawOffset;
    LONG _xOverhang; /* Can be a short */

    // Finally the tree pointer. This is where the current run is
    // WRT our tree.
    CTreePos* _ptp;
};

inline BOOL COneRun::IsSameGridClass(COneRun* por)
{
    return (por
        && por->_ptp->IsText() 
        && por->GetCF()->HasCharGrid(TRUE)
        && por->GetCF()->GetLayoutGridType(TRUE)==GetCF()->GetLayoutGridType(TRUE) 
        && por->GetCharWidthClass()==GetCharWidthClass());
}

inline BOOL COneRun::IsOneCharPerGridCell()
{
    return (GetCharWidthClass()==charWidthClassFullWidth 
        || (GetCF()->GetLayoutGridType(TRUE)==styleLayoutGridTypeFixed 
        && GetCharWidthClass()==COneRun::charWidthClassHalfWidth ));
}

#endif //__XINDOWS_SITE_TEXT_ONERUN_H__