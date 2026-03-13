#pragma once
#include "..\stdafx.h"
#include "Replacement.h"

//struct __declspec(uuid("77FE4A3A-A562-46FC-8E89-43BC5FF258E1")) IID_XCom_Clss;
//const GUID IID_XCom = __uuidof(IID_XCom_Clss);
const GUID IID_XCom =
//	{ 0x00000000, 0x0000, 0x0000, { 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB }};
	{ 0x77FE4A3A, 0xA562, 0x46FC, { 0x8E, 0x89, 0x43, 0xBC, 0x5F, 0xF2, 0x58, 0xE1 }};


namespace XInit
{
	class CoClassInfo
	{
	public:
		CLSID m_clsid;
		BSTR  m_bstrClassName;
		bool m_isReg;
		CoClassInfo()
		{
			m_bstrClassName = NULL;
			m_clsid = GUID_NULL;
			m_isReg = false;
		}
		~CoClassInfo()
		{
			if(m_bstrClassName)
				::SysFreeString(m_bstrClassName);
		}
	};
	class CoClassList : public TList <CoClassInfo>
	{
		BOOL Compare(CoClassInfo *pCoClassInfo, void *ptr)
		{
			if(wcsicmp(pCoClassInfo->m_bstrClassName, (LPCWSTR)ptr) == 0)
				return TRUE;
			else
				return FALSE;
		}
	};
	enum VARIANT_OF_REGISTRATIION
	{
		vrg_XCom = 1,
		vrg_AttachAddin,
		vrg_NET
	};
	class CXInit
	{
		HANDLE m_hMutex;
	public:
		int m_registration;
		wchar_t *m_wModulename;
		char    *m_classID; 
		wchar_t *m_wClassID;
		Replacement::CFuncReplace *m_pFuncReplace;
		CoClassList *m_pCoClassList; 
		static CXInit *& GetCXInit()
		{
			static CXInit * ptr;
			return ptr;
		}
	public:
		CXInit(wchar_t *wszFuncname);
		CXInit(wchar_t *dllname, wchar_t *classID);
		~CXInit();
	protected:
		void Init();
		void ClassID(OLECHAR *pClassID);
	public:
		HRESULT ClsidFromProgId(LPCOLESTR lpszProgID, LPCLSID pclsid);
		void ClassCorrection(LPWSTR lpwBuffer);
	};
	HRESULT __stdcall CLSIDFromProgID_Wrap(LPCOLESTR lpszProgID, LPCLSID pclsid);
	HRESULT __stdcall CreateInstance_Wrap(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext, REFIID riid, void FAR* FAR* ppv);
	FARPROC __stdcall GetProcAddress_Wrap(HMODULE hModule, LPCSTR pFunName);
	void __stdcall CoFreeLibrary_Wrap(HINSTANCE hInstance);
	HRESULT __stdcall DllInstall_Wrap(BOOL bInstall, LPCWSTR pszCmdLine);
	HRESULT __stdcall DllRegisterServer_Wrap();
	int __stdcall LoadStringW_Wrap(HINSTANCE hInstance, UINT uID, LPWSTR lpBuffer, int cchMax);

	HRESULT __cdecl CreateInstance(HMODULE hModule, const CLSID & clsid, REFIID riid, void ** ppv);
	HRESULT __cdecl Register(wchar_t *dllname, wchar_t *classID);
	HRESULT TypeInfo(wchar_t *dllname, CoClassList*);
	BOOL LockInit();
	void UnlockInit();
};