#include "stdafx.h"
#include "NativeAddInXCom.h"
#include "Thunk\XInit.h"
#include "Thunk\XCLR.h"

extern LONG gl_nRef;


enum ENUM_METHOD_NUMBERS
{
    emRegister = 0,
	emUnload,
	emUnloadAssembly,
	emTestError,
	emVersion
};
CMethod gl_Methods [] =
{
	CMethod(L"Register", emRegister, 2, true), 
	CMethod(L"Unload",	 emUnload,	 1, false),
	CMethod(L"UnloadAssembly", emUnloadAssembly, 1, false),
	CMethod(L"TestError", emTestError, 0, false),
	CMethod(L"Version",   emVersion, 0, true)

}; 

// CAddInNative
CAddInNative::CAddInNative()
{
	this->m_cMethods = gl_Methods;
	this->m_nMethods = sizeof(gl_Methods) / sizeof(CMethod);

	//FILE *fp = fopen("\\\\svu\\File_DBS\\ExtProc\\XCom\\XCom.txt", "a+");
	//fprintf(fp, "Register %u %u %u\n", ::GetCurrentProcessId(), GetCurrentThreadId(), gl_nRef);
	//fclose(fp);

}
//---------------------------------------------------------------------------//
CAddInNative::~CAddInNative()
{
}
// ILanguageExtenderBase
bool CAddInNative::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{ 
	return AllocWSTR(wsExtensionName, L"XInitDone");
}
//---------------------------------------------------------------------------//
bool CAddInNative::GetParamDefValue(const long lMethodNum, 
									const long lParamNum,
									tVariant *pvarParamDefValue)
{ 
	bool result = false;
	if(lMethodNum == emRegister)
	{
		if(lParamNum==1 || lParamNum==0)
		{
			TV_VT(pvarParamDefValue)= VTYPE_PWSTR;
			TV_WSTR(pvarParamDefValue) = L"";
		}
	}

    return result;
} 
//---------------------------------------------------------------------------//
bool CAddInNative::CallAsProc(	const long lMethodNum,
								tVariant* paParams, 
								const long lSizeArray)
{ 
	if(lMethodNum == emUnload)
	{
		Unload(paParams[0].pwstrVal);
		return true;
	}
	else if (lMethodNum == emUnloadAssembly)
	{
		UnloadAssemblyObject(paParams[0].pwstrVal);
		return true;
	}
	else if(lMethodNum == emTestError)
	{
		m_iConnect->AddError(1004, L"XNative", L"тест интерфейса AddinDefBase", 0);
		//::MessageBoxA(NULL, "Test Native", 0, 0);
		return true;
	}
    return false;
}
//---------------------------------------------------------------------------
bool CAddInNative::CallAsFunc(const long lMethodNum, tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{ 
    bool ret = false;
	CLSID clsid;
	HRESULT rc;

    if(lMethodNum == emRegister)
    { 
		rc = XInit::Register(paParams[0].pwstrVal, paParams[1].pwstrVal);
		if(rc)
		{
			this->CallException(rc, L"Регистрация");
		}
		ret = true;
    }
	else if (lMethodNum == emVersion)
	{
		wchar_t *pws = NULL;
		
		AllocWSTR(&pws, L"Native");

		SetWSTR(pvarRetValue, pws);

		ret = true;
	}
	return ret;
}
void CAddInNative::Unload(wchar_t *dllname)
{
	HMODULE hModule;
	HRESULT hr;
	HRESULT (__stdcall* pDllCanUnloadNow)();

	if(dllname && *dllname != L'\0')
		hModule = ::GetModuleHandleW(dllname);
	else
		hModule = ::GetModuleHandleA("XFunctions.dll");
	if(hModule)
	{
		pDllCanUnloadNow = (HRESULT (__stdcall*)())GetProcAddress(hModule, "DllCanUnloadNow");
		if(pDllCanUnloadNow)
		{
			hr = pDllCanUnloadNow();
			if(hr == S_OK)
				::FreeLibrary(hModule);
		}
	}
}

long __cdecl GetClassObject(const WCHAR_T* wsName, IComponentBase** pInterface)
{
	if (!*pInterface)
	{
		if (_wcsicmp(wsName, L"CAddinNative") == 0)
			*pInterface = new CAddInNative();
		//else if(_wcsicmp(wsName, L"Pipe") == 0)
		//	*pInterface= new CPipe();
		else
			*pInterface = new CAddInNative();
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

	if (!*pIntf)
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
