
#include "stdafx.h"
#include "testbase.h"

XINDOWS_PUBLIC void TestAttrAry()
{
    CAttrArray* pAttVal = NULL;
    CVariant v;
    v.vt = VT_I4;
    v.intVal = 100;
    CAttrArray::Set(&pAttVal, 0, &v);
    CAttrArray::Set(&pAttVal, 1, &v);
    CAttrArray::Set(&pAttVal, 2, &v);
    CAttrArray::Set(&pAttVal, 3, &v);
    delete pAttVal;
}