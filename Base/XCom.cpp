
#include "XCom.h"
#include "Thunk\XInit.h"

//1cv8.exe | ole32.dll | CoCreateInstance | 764C9C5B | 09901E70
//core83.dll | ole32.dll | CLSIDFromProgID | 764A4F5C | 09901D80
//core83.dll | ole32.dll | CoCreateInstance | 764C9C5B | 09901E70


HINSTANCE gl_hInstance;
LONG gl_nRef;
LPDISPATCH gl_pDispatch;

BOOL WINAPI DllMain(HINSTANCE hinstance, DWORD dwReason, LPVOID lpReserved)
{
	gl_hInstance = hinstance;
	
	if (dwReason == DLL_PROCESS_ATTACH)			// 1
	{
		//DisableThreadLibraryCalls(hinstance);
		gl_nRef++;
	}
	else if (dwReason == DLL_THREAD_ATTACH)		// 2
	{
		//gl_nRef++;
	}
	else if (dwReason == DLL_THREAD_DETACH)		// 3
	{
		//gl_nRef--;
	}
	else if (dwReason == DLL_PROCESS_DETACH)	// 0
	{
		gl_nRef--;
	}
	return TRUE;    // ok
}
STDAPI DllCanUnloadNow()
{
	return S_OK;
}
STDAPI DllRegisterServer()
{
	XInit::Register(NULL, L"GetInitDone");

	return S_OK;
}

BOOL IsVarVal(VARIANT * pVar)
{
	if ((void*)pVar)	// для размещения результата передан не NULL
		return true;
	else
		return false;
}

BSTR get_BSTR(VARIANT pvariant)
{
	int vt = (pvariant.vt & 0xFF);
	int isRef = (pvariant.vt & 0xFF00);
	if(vt == VT_BSTR)
	{
		if (isRef == VT_BYREF)
			return (*pvariant.pbstrVal);
		else
			return pvariant.bstrVal;
	}
	else if(vt==VT_VARIANT)
	{
		if(isRef == VT_BYREF)
		{
			if(pvariant.pvarVal->vt==VT_BSTR)
				return pvariant.pvarVal->bstrVal;
		}
	}
	return L"";
}

void set_EXCEPINFO(EXCEPINFO *pExcepInfo, LPWSTR source, DWORD rc)
{
	wchar_t * errbuf = new wchar_t [2048];
	DWORD ln = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, 
								NULL, 
								rc, 
								MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 
								errbuf, 
								2048, 
								NULL);
	if(ln > 0)
		pExcepInfo->bstrDescription = ::SysAllocString(errbuf);
	else
		pExcepInfo->bstrDescription = ::SysAllocString(L"Unknown error");
	delete errbuf;

	pExcepInfo->wCode = 1006;
	pExcepInfo->scode = 0;
	if(source)
		pExcepInfo->bstrSource = ::SysAllocString(source);
	else
		pExcepInfo->bstrSource = ::SysAllocString(L"XCom");
}

HINSTANCE get_HINSTANCE()
{
	return gl_hInstance;
}

LONG IncrementRef()
{
	return 0;
}
LONG DecrementRef()
{
	return 0;
}
