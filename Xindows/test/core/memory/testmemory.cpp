
#include "stdafx.h"
#include "testmemory.h"

XINDOWS_PUBLIC void TestMemAlloc()
{
    void* pv = MemAllocClear(1000);
    MemSetName((pv, "��һ�η��䣨��Ů���䣩"));
    memcpy(pv, "wo ai ni", 8);
    void* pv2 = MemAlloc(1000);

    HRESULT hr = MemRealloc(&pv, 10);
    hr = MemRealloc(&pv, 10000);
    hr = MemRealloc(&pv, 10000);
    hr = MemRealloc(&pv, 0);

    MemFree(pv2);

    for(int i=0; i<100; i++)
    {
        MemAlloc(1000);
    }
}