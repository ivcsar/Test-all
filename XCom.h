#pragma once
#include "stdafx.h"

BOOL IsVarVal(VARIANT * pVar);
BSTR get_BSTR(VARIANT pvariant);
void set_EXCEPINFO(EXCEPINFO *pExcepInfo, LPWSTR text, DWORD rc);
HINSTANCE get_HINSTANCE();
LONG IncrementRef();
LONG DecrementRef();