#pragma once
#include "..\stdafx.h"

HRESULT __stdcall CLSIDFromProgID_Stub(LPCOLESTR lpszProgID, LPCLSID pclsid);

HRESULT __stdcall CoCreateInstance_Stub(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext, REFIID riid, void FAR* FAR* ppv);

STDAPI DllInstall_Stub(BOOL bInstall, LPCWSTR pszCmdLine);

STDAPI DllRegisterServer_Stub();

FARPROC __stdcall GetProcAddress_Stub(HMODULE hModule, LPCSTR pFunName);

void __stdcall CoFreeLibrary_Stub(HINSTANCE hInstance);

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
//int __stdcall thisLoadStringW(HINSTANCE hInstance, UINT uID, LPWSTR lpBuffer, int cchMax)
//{
//	int ln = ::LoadStringW(hInstance, uID, lpBuffer, cchMax);
//	return ln;
//}
//int __stdcall thislstrcmpiW(LPCWSTR lpString1, LPCWSTR lpString2)
//{
//	int ln = ::lstrcmpiW(lpString1, lpString2);
//	return ln;
//}
