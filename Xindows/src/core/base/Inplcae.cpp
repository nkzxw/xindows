
#include "stdafx.h"
#include "Inplcae.h"

CInPlace::CInPlace()
{
}

CInPlace::~CInPlace()
{
	ClearInterface(&_pDataObj);
	Assert(!_hwnd);
}