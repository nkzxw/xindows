
#include "stdafx.h"
#include "testcollection.h"

XINDOWS_PUBLIC void TestAssocArray()
{
    CPtrBag<int*>* pPtrBag = new CPtrBag<int*>;
    pPtrBag->Set(_T("0"), new int(0));
    pPtrBag->Set(_T("1"), new int(1));
    Assert(*(pPtrBag->Get(_T("0"))) == 0);
    Assert(*(pPtrBag->Get(_T("1"))) == 1);
    delete pPtrBag->Get(_T("0"));
    delete pPtrBag->Get(_T("1"));
    delete pPtrBag;

    CPtrBagCi<int*>* pPtrBagCi = new CPtrBagCi<int*>;
    pPtrBagCi->SetCi(_T("0"), new int(0));
    delete pPtrBag->Get(_T("0"));
    pPtrBagCi->SetCi(_T("0"), new int(1));
    Assert(*(pPtrBagCi->GetCi(_T("0"))) == 1);
    delete pPtrBag->Get(_T("0"));
    delete pPtrBagCi;
}

XINDOWS_PUBLIC void TestArray()
{
    CPtrAry<void*> ptrArray;
    ptrArray.EnsureSize(10);
    ptrArray.Append(MemAlloc(10));
    ptrArray.Append(MemAlloc(20));
    ptrArray.Append(MemAlloc(30));
    ptrArray.Append(MemAlloc(40));

    CStackCStrAry strArray;
    strArray.Append()->Set(_T("Hi"));
    strArray.Append()->Set(_T("Hello"));
    strArray.Append()->Set(_T("Mexi"));
}