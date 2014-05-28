
#ifndef __XINDOWS_SITE_TEXT_SPLICE_H__
#define __XINDOWS_SITE_TEXT_SPLICE_H__

//+---------------------------------------------------------------------------
//
//  Struct:     CSpliceRecord
//
//  Synposis:   we use this structure to store poslist information from the
//              area that is being spliced out.  This is used internally
//              in SpliceTree and also by undo
//
//----------------------------------------------------------------------------
struct CSpliceRecord
{
    CTreePos::EType _type;

    union
    {
        // Begin/end edge
        struct
        {
            CElement*   _pel;
            long        _cIncl;
            BOOL        _fSkip;
        };

        // Text
        struct
        {
            unsigned long   _cch:26;    // [Text] number of characters I own directly
            unsigned long   _sid:6;     // [Text] the script id for this run
            long            _lTextID;   // [Text] Text ID for DOM text nodes
        };

        // Pointer
        struct
        {
            CMarkupPointer* _pPointer;
        };
    };
};

//+---------------------------------------------------------------------------
//
//  Class:      CSpliceRecordList
//
//  Synposis:   A list of above records
//
//----------------------------------------------------------------------------
class CSpliceRecordList : public CDataAry<CSpliceRecord>
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();

    CSpliceRecordList() {}

    ~CSpliceRecordList();
};

#endif //__XINDOWS_SITE_TEXT_SPLICE_H__