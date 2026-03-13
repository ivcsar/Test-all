#include "Replacement.h"

using namespace Replacement;
using namespace std;

CFuncReplace::CFuncReplace()
{
	m_pFuncList = new TList<CFuncDescription>();
	m_pReplacementList = new TList<REPLACEMENT>();
}
CFuncReplace::~CFuncReplace()
{
	REPLACEMENT * repl;
	DWORD protectAttr = 0;
	
	m_pReplacementList->MoveFirst();
	while((repl = m_pReplacementList->Next()))
	{
		VirtualProtect(repl->m_funplace, sizeof(void*), PAGE_EXECUTE_READWRITE, &protectAttr);
		memmove(repl->m_funplace, &repl->m_oldptr, sizeof(void*));
		//*(repl->m_funplace) = repl->m_oldptr;
		VirtualProtect(repl->m_funplace, sizeof(void*), protectAttr, &protectAttr);
	}
	delete m_pReplacementList;
	delete m_pFuncList;
}

void CFuncReplace::Add(LPCWSTR wszFunctionName, void *pFunctionAddr)
{
	BOOL bUsedDefaultChar = FALSE;
	char *szName = NULL;

	int iBufSize = ::WideCharToMultiByte(CP_ACP, 0, wszFunctionName, -1, NULL, 0, NULL, &bUsedDefaultChar);
	if(iBufSize > 0)
	{
		CFuncDescription *pFuncDescription = new CFuncDescription();
		pFuncDescription->m_pFuncname = new char[iBufSize + 1];
		iBufSize = ::WideCharToMultiByte(0, 0, wszFunctionName, -1, pFuncDescription->m_pFuncname, iBufSize, NULL, &bUsedDefaultChar);
		pFuncDescription->m_pNewFuncAddr = pFunctionAddr;
		m_pFuncList->Add(pFuncDescription);
	}
}

int Replacement::ReplaceDllFunctions(LPCSTR c_dllname, LPCSTR win_dllname, void * pars)
{
	CFuncReplace *pFuncReplace = (CFuncReplace*)pars;
		
	pFuncReplace->ReplaceFuncs(c_dllname, win_dllname);

	return 0;
}

void CFuncReplace::Replace()
{
	//ofstream *of = new ofstream("\\\\dmsk0137\\Extproc\\Xcom\\Xcom.txt", ios::trunc);
	//of->close();
	EnumDlls(Replacement::ReplaceDllFunctions, (void*)this, NULL);	// меняет все вызовы CoCreateInstance и CLSIDFromProgID
	//Replacement::ReplaceDllFunctions("1cv8.exe",    "ole32.dll", this);
	//Replacement::ReplaceDllFunctions("core83.dll",  "ole32.dll", this);
	//Replacement::ReplaceDllFunctions("frame.dll",   "ole32.dll", this);
	//Replacement::ReplaceDllFunctions("mngui.dll",   "ole32.dll", this);
	//Replacement::ReplaceDllFunctions("addncom.dll", "ole32.dll", this);

	return;
}

void CFuncReplace::ReplaceFuncs(LPCSTR szAppDllname, LPCSTR szWinDllname)
{
	CFuncDescription * pFuncDescription = NULL;
	m_pFuncList->MoveFirst();
	while((pFuncDescription = m_pFuncList->Next()))
	{
		REPLACEMENT *pRepl = (REPLACEMENT*)ReplaceFunctionAddr(szAppDllname, szWinDllname, pFuncDescription->m_pFuncname, pFuncDescription->m_pNewFuncAddr);
		if (pRepl)
		{
			this->m_pReplacementList->Add(pRepl);
			//ofstream *of = new ofstream("\\\\dmsk0137\\Extproc\\Xcom\\Xcom.txt", ios::app);
			//*of << szAppDllname << "|" << szWinDllname << "|" << pFuncDescription->m_pFuncname << "|" << pRepl->m_oldptr << "|" << pFuncDescription->m_pNewFuncAddr << std::endl;
			//of->close();
		}
	}
}
