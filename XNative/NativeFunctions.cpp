#pragma unmanaged
#include "stdafx.h"
#include "NativeBase.h"
#include "NativeAddIn.h"

long __cdecl GetClassObject(const WCHAR_T* wsName, IComponentBase** pInterface)
{
    if(!*pInterface)
    {
		if(_wcsicmp(wsName, L"CAddinNative") == 0)
			*pInterface= new CAddInNative();
		//else if(_wcsicmp(wsName, L"Pipe") == 0)
		//	*pInterface= new CPipe();
		else
			*pInterface= new CAddInNative();
        return (long)*pInterface;
    }
    return 0;
}

void showPtr(char *eventname, void *ptr)
{
	//char cbuf[256] = {0};
	//sprintf_s(cbuf, 256, "%p", ptr);
	//::MessageBoxA(NULL, cbuf, eventname, 0);
}
//---------------------------------------------------------------------------//
long __cdecl DestroyObject(IComponentBase** pIntf)
{
	showPtr("object : destroy", *pIntf);

	if(!*pIntf)
		return -1;
	delete (*pIntf);
	*pIntf = 0;

	return 0;
}
//---------------------------------------------------------------------------//
const WCHAR_T* __cdecl GetClassNames()
{
    //static WCHAR_T* names = 0;
    //if (!names)
    //    ::convToShortWchar(&names, L"CAddInNative", 0);
    //return names;
	return L"CAddinNative";
}