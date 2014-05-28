
#ifndef __XINDOWSUTIL_COLLECTION_ASSOC_H__
#define __XINDOWSUTIL_COLLECTION_ASSOC_H__

// Hashing function decls
XINDOWS_PUBLIC DWORD HashString(const TCHAR* pch, DWORD len, DWORD hash);
XINDOWS_PUBLIC DWORD HashStringCi(const TCHAR* pch, DWORD len, DWORD hash);

//+---------------------------------------------------------------------------
//
//  Class:      CAssoc
//
//  Purpose:    A single association in an associative array mapping
//              strings -> DWORD_PTRs.
//
//              The class is designed to be an aggregate so that it is
//              statically initializable. Therefore, it has no base
//              class, no constructor/destructors, and no private members
//              or methods.
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CAssoc
{
public:
	void* operator new(size_t cb)
	{
		return MemAllocClear(cb);
	}
	void* operator new(size_t cb, size_t cbExtra)
	{
		return MemAllocClear(cb+cbExtra);
	}
	void operator delete(void* pv) { MemFree(pv); }

	void Init(DWORD_PTR number, const TCHAR* pch, int length, DWORD hash)
	{
		_number = number;
		_hash = hash;
		memcpy(_ach, pch, (length+1)*sizeof(TCHAR));
	}

	TCHAR* String()     { return _ach; }
	DWORD_PTR Number()  { return _number; }
	DWORD Hash()        { return _hash; }

	DWORD_PTR   _number;
	DWORD       _hash;
	TCHAR       _ach[];
};

//+---------------------------------------------------------------------------
//
//  Class:      CAssocArray
//
//  Purpose:    A hash table implementation mapping strings -> DWORDs
//
//              The class is designed to be an aggregate so that it is
//              statically initializable. Therefore, it has no base
//              class, no constructor/destructors, and no private members
//              or methods.
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CAssocArray
{
public:
	void Init();
	void Deinit();

	CAssoc* AssocFromString(const TCHAR* pch, DWORD len, DWORD hash);
	CAssoc* AssocFromStringCi(const TCHAR* pch, DWORD len, DWORD hash);
	CAssoc* AddAssoc(DWORD_PTR number, const TCHAR* pch, DWORD len, DWORD hash);

	CAssoc**    _pHashTable;    // Assocs hashed by string
	DWORD       _cHash;         // entries in the hash table
	DWORD       _mHash;         // hash table modulus
	DWORD       _sHash;         // hash table stride mask
	DWORD       _maxHash;       // maximum entries allowed
	DWORD       _iSize;         // prime array index

	union
	{
		BOOL	_fStatic;       // TRUE for a static Assoc table
		CAssoc*	_pAssocOne;     // NULL for a dynamic Assoc table
	};

	DWORD EmptyHashIndex(DWORD hash);
	HRESULT ExpandHash();
};

//+---------------------------------------------------------------------------
//
//  Class:      CImplPtrBag
//
//  Purpose:    Implements an associative array of strings->pointers.
//
//              Implemented as a CAssocArray. Can be intialized to point to
//              a static associative array, or can be dynamic.
//
//----------------------------------------------------------------------------
class XINDOWS_PUBLIC CImplPtrBag : protected CAssocArray
{
public:
	CImplPtrBag() { Init(); }
	CImplPtrBag(const CAssocArray* pTable)
	{
		Assert(pTable->_fStatic);
		*(CAssocArray*)this = *pTable;
	}

	~CImplPtrBag() { if(!_fStatic) { Deinit(); } }

	HRESULT SetImpl(const TCHAR* pch, DWORD cch, DWORD hash, void* e);
	void* GetImpl(const TCHAR* pch, DWORD cch, DWORD hash);

	HRESULT SetCiImpl(const TCHAR* pch, DWORD cch, DWORD hash, void* e);
	void* GetCiImpl(const TCHAR* pch, DWORD cch, DWORD hash);
};

//+---------------------------------------------------------------------------
//
//  Class:      CPtrBag
//
//  Purpose:    A case-sensitive ptrbag.
//
//              This template class declares a concrete derived class
//              of CImplPtrBag.
//
//----------------------------------------------------------------------------
template<class ELEM> class CPtrBag : protected CImplPtrBag
{
public:
	CPtrBag() : CImplPtrBag() { Assert(sizeof(ELEM)<=sizeof(void*)); }
	CPtrBag(const CAssocArray* pTable) : CImplPtrBag(pTable)
	{
		Assert(sizeof(ELEM)<=sizeof(void*));
	}

	~CPtrBag() {}

	CPtrBag& operator=(const CPtrBag&);

	HRESULT Set(const TCHAR* pch, ELEM e)
	{
		DWORD dw = _tcslen(pch);
		return Set(pch, dw, e);
	}
	HRESULT Set(const TCHAR* pch, DWORD cch, ELEM e) { return Set(pch, cch, HashString(pch, cch, 0), e); }
	HRESULT Set(const TCHAR* pch, DWORD cch, DWORD hash, ELEM e)
	{
		return CImplPtrBag::SetImpl(pch, cch, hash, (void*)(DWORD_PTR)e);
	}
	ELEM Get(const TCHAR* pch)
	{
		DWORD dw = _tcslen(pch);
		return Get(pch, dw);
	}
	ELEM Get(const TCHAR* pch, DWORD cch)
	{
		return Get(pch, cch, HashString(pch, cch, 0));
	}
	ELEM Get(const TCHAR* pch, DWORD cch, DWORD hash)
	{
		return (ELEM)(DWORD_PTR)CImplPtrBag::GetImpl(pch, cch, hash);
	}
};

//+---------------------------------------------------------------------------
//
//  Class:      CPtrBagCi
//
//  Purpose:    A case-insensitive ptrbag.
//
//              Supplies case-insensitive GetCi/SetCi methods in addition to
//              the  case-sensitive GetCs/SetCs methods.
//
//----------------------------------------------------------------------------
template<class ELEM> class CPtrBagCi : protected CImplPtrBag
{
public:
	CPtrBagCi() : CImplPtrBag() { Assert(sizeof(ELEM)<=sizeof(void*)); }
	CPtrBagCi(const CAssocArray* pTable) : CImplPtrBag(pTable)
	{
		Assert(sizeof(ELEM)<=sizeof(void*));
	}

	~CPtrBagCi() {}

	CPtrBagCi& operator=(const CPtrBagCi&);

	HRESULT SetCs(const TCHAR* pch, ELEM e)
	{
		DWORD dw = _tcslen(pch);
		return SetCs(pch, dw, e);
	}
	HRESULT SetCs(const TCHAR* pch, DWORD cch, ELEM e) { return SetCs(pch, cch, HashStringCi(pch, cch, 0), e); }
	HRESULT SetCs(const TCHAR* pch, DWORD cch, DWORD hash, ELEM e)
	{
		return CImplPtrBag::SetImpl(pch, cch, hash, (void*)(DWORD_PTR)e);
	}
	ELEM GetCs(const TCHAR* pch)
	{
		DWORD dw = _tcslen(pch);
		return GetCs(pch, dw);
	}
	ELEM GetCs(const TCHAR* pch, DWORD cch) { return GetCs(pch, cch, HashStringCi(pch, cch, 0)); }
	ELEM GetCs(const TCHAR* pch, DWORD cch, DWORD hash) { return (ELEM)(DWORD_PTR)CImplPtrBag::GetImpl(pch, cch, hash); }

	HRESULT SetCi(const TCHAR* pch, ELEM e)
	{
		DWORD dw = _tcslen(pch);
		return SetCi(pch, dw, e);
	}
	HRESULT SetCi(const TCHAR* pch, DWORD cch, ELEM e) { return SetCi(pch, cch, HashStringCi(pch, cch, 0), e); }
	HRESULT SetCi(const TCHAR* pch, DWORD cch, DWORD hash, ELEM e)
	{
		return CImplPtrBag::SetCiImpl(pch, cch, hash, (void*)(DWORD_PTR)e);
	}
	ELEM GetCi(const TCHAR* pch)
	{
		DWORD dw = _tcslen(pch);
		return GetCi(pch, dw);
	}
	ELEM GetCi(const TCHAR* pch, DWORD cch) { return GetCi(pch, cch, HashStringCi(pch, cch, 0)); }
	ELEM GetCi(const TCHAR* pch, DWORD cch, DWORD hash)
	{
		return (ELEM)(DWORD_PTR)CImplPtrBag::GetCiImpl(pch, cch, hash);
	}
};

#endif	//__XINDOWSUTIL_COLLECTION_ASSOC_H__