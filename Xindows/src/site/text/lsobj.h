
#ifndef __XINDOWS_SITE_TEXT_LSOBJ_H__
#define __XINDOWS_SITE_TEXT_LSOBJ_H__

#include "LineServices.h"

#include "msls/lsdnset.h"

// This class replaces struct dobj in a C++ context.  It is the
// abstract base class for all the specific types of ILS objects
// to subclass.
// Each CILSObjBase subclass will have a corresponding subclass
// of this class.
class CDobjBase
{
private:
    DECLARE_MEMALLOC_NEW_DELETE()

    // typedefs
private:
    typedef CILSObjBase* PILSOBJ;
    typedef COneRun* PLSRUN;
    typedef CDobjBase DOBJ;
    typedef DOBJ* PDOBJ;

public:
    // Data members
    PILSOBJ  _pilsobj; // pointer back to the CILSObjBase that created us
    PLSDNODE _plsdnTop;

    // We cache the difference between the minimum and maximum width
    // in this variable.  After a min-max pass, we enumerate our dobj
    // and compute the true minimum for those sites where the min and
    // max widths differ (e.g. table)
    LONG _dvMinMaxDelta;

    // Methods
protected:
    CDobjBase(PILSOBJ pilsobjNew, PLSDNODE plsdn);
    virtual void UselessVirtualToMakeTheCompilerHappy() {}

public:
    CLineServices* GetPLS() { return _pilsobj->_pLS; };
    PLSC GetPLSC() { return GetPLS()->_plsc;  };

    LSERR QueryObjDim(POBJDIM pobjdim)
    {
        return LsdnQueryObjDimRange(GetPLSC(), _plsdnTop, _plsdnTop, pobjdim);
    };
};

class CEmbeddedDobj: public CDobjBase
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()

    // typedefs
private:
    typedef CILSObjBase* PILSOBJ;
    typedef COneRun* PLSRUN;
    typedef CDobjBase DOBJ;
    typedef DOBJ* PDOBJ;

public:
    // Data Members

    // IE3-only sites (MARQUEE/OBJECT) will break between themselves and
    // text.  We note whether we have such a site in the dobj to facilitate
    // breaking.
    BOOL _fIsBreakingSite : 1;

    // Methods
    CEmbeddedDobj(PILSOBJ pilsobjNew, PLSDNODE plsdn);
    static LSERR WINAPI FindPrevBreakChunk(PCLOCCHNK, PCPOSICHNK, BRKCOND, PBRKOUT);
    static LSERR WINAPI FindNextBreakChunk(PCLOCCHNK, PCPOSICHNK, BRKCOND, PBRKOUT);

    // Line services Callbacks
    static LSERR WINAPI DestroyDObj(PDOBJ);
    static LSERR WINAPI Enum(PDOBJ, PLSRUN, PCLSCHP, LSCP, LSDCP, LSTFLOW, BOOL, BOOL, const POINT*, PCHEIGHTS, long);
    static LSERR WINAPI Display(PDOBJ, PCDISPIN);
    static LSERR WINAPI QueryPointPcp(PDOBJ, PCPOINTUV, PCLSQIN, PLSQOUT);
    static LSERR WINAPI QueryCpPpoint(PDOBJ, LSDCP, PCLSQIN, PLSQOUT);
    static LSERR WINAPI GetModWidthPrecedingChar(PDOBJ, PLSRUN, PLSRUN, PCHEIGHTS, WCHAR, MWCLS, long*);
    static LSERR WINAPI GetModWidthFollowingChar(PDOBJ, PLSRUN, PLSRUN, PCHEIGHTS, WCHAR, MWCLS, long*);
    static LSERR WINAPI SetBreak(PDOBJ, BRKKIND, DWORD, BREAKREC*, DWORD*);
    static LSERR WINAPI GetSpecialEffectsInside(PDOBJ, UINT*);
    static LSERR WINAPI FExpandWithPrecedingChar(PDOBJ, PLSRUN, PLSRUN, WCHAR, MWCLS, BOOL*);
    static LSERR WINAPI FExpandWithFollowingChar(PDOBJ, PLSRUN, PLSRUN, WCHAR, MWCLS, BOOL*);
    static LSERR WINAPI CalcPresentation(PDOBJ, long, LSKJUST, BOOL);
};

class CNobrDobj : public CDobjBase
{
public:
    DECLARE_MEMALLOC_NEW_DELETE()

    // typedefs
private:
    typedef CILSObjBase* PILSOBJ;
    typedef COneRun* PLSRUN;
    typedef CDobjBase DOBJ;
    typedef DOBJ* PDOBJ;

public:
    // Data Members
    PLSSUBL _plssubline;    // the subline representing this nobr block
    BRKCOND _brkcondBefore; // brkcond at the beginning of the NOBR
    BRKCOND _brkcondAfter;  // brkcond at the end of the NOBR

    // Methods
    CNobrDobj(PILSOBJ pilsobjNew, PLSDNODE plsdn);
    BRKCOND GetBrkcondBefore(COneRun* por);
    BRKCOND GetBrkcondAfter(COneRun* por, LONG dcp);
    static LSERR WINAPI FindPrevBreakChunk(PCLOCCHNK, PCPOSICHNK, BRKCOND, PBRKOUT);
    static LSERR WINAPI FindNextBreakChunk(PCLOCCHNK, PCPOSICHNK, BRKCOND, PBRKOUT);

    // Line services Callbacks
    static LSERR WINAPI DestroyDObj(PDOBJ);
    static LSERR WINAPI Enum(PDOBJ, PLSRUN, PCLSCHP, LSCP, LSDCP, LSTFLOW, BOOL, BOOL, const POINT*, PCHEIGHTS, long);
    static LSERR WINAPI Display(PDOBJ, PCDISPIN);
    static LSERR WINAPI QueryPointPcp(PDOBJ, PCPOINTUV, PCLSQIN, PLSQOUT);
    static LSERR WINAPI QueryCpPpoint(PDOBJ, LSDCP, PCLSQIN, PLSQOUT);
    static LSERR WINAPI GetModWidthPrecedingChar(PDOBJ, PLSRUN, PLSRUN, PCHEIGHTS, WCHAR, MWCLS, long*);
    static LSERR WINAPI GetModWidthFollowingChar(PDOBJ, PLSRUN, PLSRUN, PCHEIGHTS, WCHAR, MWCLS, long*);
    static LSERR WINAPI SetBreak(PDOBJ, BRKKIND, DWORD, BREAKREC*, DWORD*);
    static LSERR WINAPI GetSpecialEffectsInside(PDOBJ, UINT*);
    static LSERR WINAPI FExpandWithPrecedingChar(PDOBJ, PLSRUN, PLSRUN, WCHAR, MWCLS, BOOL*);
    static LSERR WINAPI FExpandWithFollowingChar(PDOBJ, PLSRUN, PLSRUN, WCHAR, MWCLS, BOOL*);
    static LSERR WINAPI CalcPresentation(PDOBJ, long, LSKJUST, BOOL);
};

#endif //__XINDOWS_SITE_TEXT_LSOBJ_H__