#include "..\Xcom.h"
#include "XInit.h"
#include "XCLR.h"

using namespace XInit;

/************ функции на замену ***************************/
HRESULT __stdcall XInit::CLSIDFromProgID_Wrap(LPCOLESTR lpszProgID, LPCLSID pclsid)
{
	HRESULT rc;
	CXInit * pInit;
	pInit = CXInit::GetCXInit();
	rc = pInit->ClsidFromProgId(lpszProgID, pclsid);
	if(rc != S_OK)
		XInit::UnlockInit();
	return rc;
}
HRESULT __stdcall XInit::CreateInstance_Wrap(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext, REFIID riid, void FAR* FAR* ppv)
{
	HRESULT rc = S_OK;
	HINSTANCE hInstance;

	CXInit * pInit = CXInit::GetCXInit();

	//char *ptrfun = (char*)CLSIDFromProgID_Wrap;  
	//void (__stdcall *pUnkfun)();
	//pUnkfun = (void (__stdcall*)())(ptrfun + 81168);

	if(pInit->m_registration == XInit::vrg_XCom)
	{
		int cs = strlen(pInit->m_classID) + 1;
		char *funcName = new char[cs];
		memmove(funcName, pInit->m_classID, cs-1);
		funcName[cs-1] = 0;
		XInit::UnlockInit();

		HRESULT (__stdcall *get_Object)(void**);

		get_Object = (HRESULT (__stdcall*)(void**)) GetProcAddress(get_HINSTANCE(), funcName);
		delete funcName;
		if(get_Object)
			rc = get_Object(ppv);
		else
			//REGDB_E_CLASSNOTREG A specified class is not registered in the registration database. Also can indicate that the type of server you requested in the CLSCTX enumeration is not registered or the values for the server types in the registry are corrupt.
			//CLASS_E_NOAGGREGATION This class cannot be created as part of an aggregate.
			//E_NOINTERFACE The specified class does not implement the requested interface, or the controlling IUnknown does not expose the requested interface.
			rc = REGDB_E_CLASSNOTREG;
	}
	else if(pInit->m_registration==XInit::vrg_AttachAddin)		// подключение внешней компоненты
	{
		wchar_t *wszDllname = ::SysAllocString(pInit->m_wModulename); 
		if(pInit->m_pCoClassList->Count() <= 0)
			XInit::UnlockInit();

		hInstance = ::GetModuleHandleW(wszDllname);
		if(!hInstance)
			hInstance = ::CoLoadLibrary(wszDllname, TRUE);
		::SysFreeString(wszDllname);
		rc = XInit::CreateInstance(hInstance, rclsid, riid, ppv);
	}
	else if(pInit->m_registration==XInit::vrg_NET)
	{
		wchar_t *wszDllname = ::SysAllocString(pInit->m_wModulename); 
		wchar_t *wClassname = ::SysAllocString(pInit->m_wClassID);
		XInit::UnlockInit();

		VARIANT rvar = {0};
		rc = LoadAssemblyObject(wszDllname, wClassname, &rvar);
		if(rc == S_OK)
		{
			*ppv = rvar.pdispVal;
		}
	}
	else
	{
		XInit::UnlockInit();
		rc = REGDB_E_CLASSNOTREG;
	}
	return rc;
}
FARPROC __stdcall XInit::GetProcAddress_Wrap(HMODULE hModule, LPCSTR pFunName)
{
	FARPROC procAddr;

	procAddr = ::GetProcAddress(hModule, pFunName);
	if(procAddr)
	{
		CXInit * pInit = CXInit::GetCXInit();
		HMODULE _hModule = GetModuleHandleW(pInit->m_wModulename);
		if(_hModule == hModule)
		{
			int lFunName = strlen(pFunName);
			if(lFunName==10 && memcmp(pFunName, "DllInstall", lFunName)==0)
				procAddr = (FARPROC)XInit::DllInstall_Wrap;
			else if(lFunName==17 && memcmp(pFunName, "DllRegisterServer", lFunName)==0)
				procAddr = (FARPROC)XInit::DllRegisterServer_Wrap;
		}
		else
			procAddr = ::GetProcAddress(hModule, pFunName);
	}
	return procAddr;
}

void __stdcall XInit::CoFreeLibrary_Wrap(HINSTANCE hInstance)
{
	CXInit * pInit = CXInit::GetCXInit();
	if(hInstance == get_HINSTANCE())
		return;
	//HMODULE _hModule = GetModuleHandleW(pInit->m_wModulename);
	//if(_hModule == hInstance)
	//{
	//	return;
	//}
	else
		::CoFreeLibrary(hInstance);
}

HRESULT __stdcall XInit::DllInstall_Wrap(BOOL bInstall, LPCWSTR pszCmdLine)
{
	return S_OK;
}

HRESULT __stdcall XInit::DllRegisterServer_Wrap()
{
	return S_OK;
}

//
//LONG __stdcall thisRegQueryInfoKey(HKEY hKey, LPWSTR  lpClass, LPDWORD lpcbClass, LPDWORD lpReserved, LPDWORD lpcSubKeys, LPDWORD lpcbMaxSubKeyLen, LPDWORD lpcbMaxClassLen, LPDWORD lpcValues, LPDWORD lpcbMaxValueNameLen, LPDWORD lpcbMaxValueLen, LPDWORD lpcbSecurityDescriptor, PFILETIME lpftLastWriteTime)
//{
//	return ::RegQueryInfoKeyW(hKey, lpClass, lpcbClass, lpReserved, lpcSubKeys, lpcbMaxSubKeyLen, lpcbMaxClassLen, lpcValues, lpcbMaxValueNameLen, lpcbMaxValueLen, lpcbSecurityDescriptor, lpftLastWriteTime);
//}
//
//LONG __stdcall thisRegOpenKeyEx(HKEY hKey, LPCWSTR lpSubKey, DWORD ulOptions, REGSAM samDesired, PHKEY phkResult)
//{
//	return ::RegOpenKeyExW(hKey, lpSubKey, ulOptions, samDesired, phkResult);
//}
//LSTATUS __stdcall thisRegEnumKeyEx(HKEY hKey, DWORD dwIndex, LPWSTR lpName, LPDWORD lpcchName, LPDWORD lpReserved, LPWSTR lpClass, LPDWORD lpcchClass, PFILETIME pFileTime)
//{
//	return ::RegEnumKeyExW(hKey, dwIndex, lpName, lpcchName, lpReserved, lpClass, lpcchClass, pFileTime);
//}
//HRESULT __stdcall thisCoGetClassObject(REFCLSID rclsid, 
//							WORD dwClsContext, 
//							COSERVERINFO* pServerInfo, 
//							REFIID riid, 
//							LPVOID* ppv)
//{
//	return ::CoGetClassObject(rclsid, dwClsContext, pServerInfo, riid, ppv);
//}
//
//HRESULT __stdcall thisCoRegisterClassObject(REFCLSID rclsid, IUnknown * pUnk, DWORD dwClsContext, DWORD flags, LPDWORD  lpdwRegister)
//{
//	return ::CoRegisterClassObject(rclsid, pUnk, dwClsContext, flags, lpdwRegister);
//}
//HRSRC __stdcall thisFindResourceW(HMODULE hm, LPCWSTR lpName, LPCWSTR lpType, WORD wLanguage)
//{
//	return ::FindResourceExW(hm, lpName, lpType, wLanguage);
//	return ::FindResourceW(hm, lpName, lpType);
//}
//HGLOBAL __stdcall thisLoadResource(HMODULE hm, HRSRC hResinfo)
//{
//	return ::LoadResource(hm, hResinfo);
//}
//HINSTANCE __stdcall thisCoLoadLibrary(LPOLESTR lpszLibName, BOOL bAutoFree)
//{
//	return ::CoLoadLibrary(lpszLibName, bAutoFree);
//}
int __stdcall XInit::LoadStringW_Wrap(HINSTANCE hInstance, UINT uID, LPWSTR lpBuffer, int cchMax)
{
	int ln = ::LoadStringW(hInstance, uID, lpBuffer, cchMax);
	if(uID == 100)
	{
		CXInit *pInit = CXInit::GetCXInit();
		pInit->ClassCorrection(lpBuffer);
	}
	return ln;
}
//int __stdcall thislstrcmpiW(LPCWSTR lpString1, LPCWSTR lpString2)
//{
//	int ln = ::lstrcmpiW(lpString1, lpString2);
//	return ln;
//}
