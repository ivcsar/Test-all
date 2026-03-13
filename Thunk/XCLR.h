#ifndef __XCLR_H__
#define __XCLR_H__

#include "..\stdafx.h"

HRESULT LoadAssembly(BSTR, VARIANT*);
HRESULT LoadAssemblyObject(BSTR, BSTR, VARIANT*);
BSTR ErrorDescription(HRESULT hResult);
HRESULT UnloadAssemblyObject(BSTR dllname);

#endif
