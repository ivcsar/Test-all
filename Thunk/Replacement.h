#pragma once
#include "..\stdafx.h"
#include "..\TList.h"
#ifndef __REPLACEMENT_H__
#define __REPLACEMENT_H__

namespace Replacement
{
	class REPLACEMENT
	{
		public:
			void FAR * m_oldptr;
			void FAR * m_funplace;
	};
	struct ENUMFUNCTIONSPARS
	{
		const char *funName;
		void *pFunc;
		std::ofstream *file;
	};
	typedef int EnumFunctionsProc(LPCSTR dll_1c, LPCSTR dll_win, LPCSTR funName, void * pFunc, void * parameters);
	typedef int EnumDllsProc(LPCSTR dll_1c, LPCSTR dll_win, void * parameters);

	int FunctionsList(LPCSTR dll_1c, LPCSTR dll_win, void*);
	static void EnumFunctions(LPCSTR dll_1c, LPCSTR dll_win, EnumFunctionsProc *, void*);

	static void *GetThunkAddress(LPCSTR strDll, LPCSTR strExporter, LPCSTR strFunc);
	void  EnumDlls(EnumDllsProc * proc, void * pars, HMODULE hExcludeModule);
	int   get_FunThunkAddr(LPCSTR dll_1c, LPCSTR dll_win, LPCSTR funName, void * pFunc, void * parameters);
	void * ReplaceFunctionAddr(LPCSTR strDll, LPCSTR strExporter, LPCSTR strFunc, void * funAddress);

	int ReplaceDllFunctions(LPCSTR c_dllname, LPCSTR win_dllname, void * pars);
	
	class CFuncDescription
	{
	public:
		CFuncDescription()
		{
			m_pFuncname = NULL;
			m_pNewFuncAddr = NULL;
		}
		~CFuncDescription()
		{
			if(m_pFuncname)
			{
				delete m_pFuncname;
				m_pFuncname = NULL;
			}
		}
	public:
		char *m_pFuncname;
		void *m_pNewFuncAddr; 
	};
	class CFuncReplace
	{
		TList <REPLACEMENT> *m_pReplacementList;
		TList <CFuncDescription> *m_pFuncList; 
	public:
		CFuncReplace();
		~CFuncReplace();
		void Add(LPCWSTR pFuncName, void *pNewFuncAddr);
		void Replace();
		void ReplaceFuncs(LPCSTR szAppDllname, LPCSTR szWinDllname);
	};
};
#endif