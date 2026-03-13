#include "..\stdafx.h"
#include "..\XCom.h"
#include <Psapi.h>
#include "Replacement.h"

#pragma comment (lib, "psapi.lib")

/****************** список функций в модуле *****************************************/
void Replacement::EnumFunctions(LPCSTR dll_1c, LPCSTR dll_win, EnumFunctionsProc enumfunproc, void FAR * parameters)
{
	HMODULE hDll;

	if (_strncoll("api-ms", dll_1c, 6) == 0)
		return;
	
	hDll = GetModuleHandle(dll_1c);
	if(!hDll) 
		return;

	BYTE *pimage = reinterpret_cast<BYTE*>(hDll);

	PIMAGE_OPTIONAL_HEADER ioh = reinterpret_cast<PIMAGE_OPTIONAL_HEADER>(pimage + 
								reinterpret_cast<PIMAGE_DOS_HEADER>(hDll)->e_lfanew + 4 + 
								sizeof(IMAGE_FILE_HEADER));
	
	PIMAGE_IMPORT_DESCRIPTOR iid = reinterpret_cast<PIMAGE_IMPORT_DESCRIPTOR>(pimage +
		(ioh->DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT].VirtualAddress));
	
	// В таблице импорта ищем соответствующий элемент для библиотеки
	while(iid->Name)  //до тех пор пока поле структуры не содержит 0
	{
		const char * curname = reinterpret_cast<const char*>(pimage + iid->Name);
		if(_strcmpi(curname, (const char*)dll_win)==0)
			break;
		iid++;
	}
	if(iid->Name == 0)
		return;
	PIMAGE_THUNK_DATA isd = reinterpret_cast<PIMAGE_THUNK_DATA>(pimage + iid->OriginalFirstThunk);
	BYTE FAR * pReal = reinterpret_cast<BYTE FAR*>(pimage + iid->FirstThunk);
	LPCSTR fun_name = "";
	int retf;
	char wbuf[24];
	unsigned int i;
	for(i=0; isd->u1.Function; isd++, i++)
	{
		if(IMAGE_SNAP_BY_ORDINAL(isd->u1.Ordinal))
		{
			WORD szEntry = IMAGE_ORDINAL(isd->u1.Ordinal);
			sprintf_s(wbuf, sizeof(wbuf), "%d", (unsigned int)szEntry);
			fun_name = wbuf;
		}
		else
		{
			PIMAGE_IMPORT_BY_NAME iin = (PIMAGE_IMPORT_BY_NAME)isd->u1.AddressOfData;
			fun_name=(LPCSTR)(pimage + (DWORD)(iin->Name));
		}
		retf = enumfunproc(dll_1c, dll_win, fun_name, pReal + i*sizeof(void*), parameters);
		if(retf == 1)
			break;
	}
/*
	if (!IMAGE_SNAP_BY_ORDINAL(pNamesTable->u1.Ordinal))
    {
		PIMAGE_IMPORT_BY_NAME pName = MakePtr(PIMAGE_IMPORT_BY_NAME, hModule, pNamesTable->u1.AddressOfData);
		if(0 == ::lstrcmpiA(LPSTR(pName->Name), szEntry))
        {
			hr = S_OK;
			break;
        }
    }
*/
/*Получим адрес таблицы импорта
	PIMAGE_THUNK_DATA pThunk = (PIMAGE_THUNK_DATA)((PBYTE) hmodCaller + pImportDesc->FirstThunk);
	//Переберём все импортируемые функции
	for (; pThunk->u1.Function; pThunk++) 
	{
		PROC* ppfn = (PROC*) &pThunk->u1.Function; //Получим адрес функции
		BOOL fFound = (*ppfn == pfnCurrent);     //Его ищем?
		if (!fFound && (*ppfn > sm_pvMaxAppAddr)) 
		{
			// Если не нашли, то поищем поглубже. 
			// Если мы в Win98 под отладчиком, то
			// здесь может быть push с адресом нашей функции
			PBYTE pbInFunc = (PBYTE) *ppfn;
			if (pbInFunc[0] == cPushOpCode) 
			{
				//Да, здесь PUSH
				ppfn = (PROC*) &pbInFunc[1];
				//Наш адрес?
				fFound = (*ppfn == pfnCurrent);
			}
		}
		if (fFound) 
		{ 
			//Нашли!!!
			DWORD dwDummy;
			//Разрешим запись в эту страницу
			VirtualProtect(ppfn, sizeof(ppfn), PAGE_EXECUTE_READWRITE, &dwDummy);
			//Сменим адрес на свой
			WriteProcessMemory(GetCurrentProcess(), ppfn, &pfnNew, 
			sizeof(pfnNew), NULL);
			//Восстановим атрибуты
			VirtualProtect(ppfn, sizeof(ppfn), dwDummy , &dwDummy);
			//Готово!!!
			return;
		}
	}
*/
}
/********************** список функций в библиотеках *******************************************************/
void Replacement::EnumDlls(EnumDllsProc * proc, void * parameters, HMODULE hExcludeModule)
{

	struct ENUMDLLS_WORKSTRUCT
	{
		char baseModuleName[512];
		char moduleName[512];
		HMODULE hModules[500];
	} *wstr;

	DWORD i;
	DWORD lBaseModule, lModule, lPath;
	
	wstr = new ENUMDLLS_WORKSTRUCT;
	memset(wstr, 0, sizeof(*wstr));

	HANDLE hProcess = GetCurrentProcess();

	lBaseModule = ::GetModuleFileNameA(NULL, wstr->baseModuleName, sizeof(wstr->baseModuleName));
	lBaseModule--;
	while(wstr->baseModuleName[lBaseModule] != '\\')
	{
		lBaseModule--;
	}

	DWORD cbNeeded = 0;

	::EnumProcessModules(hProcess, wstr->hModules, sizeof(wstr->hModules), &cbNeeded);
			
	TList <char> *wdlls = new TList <char> ();
	TList <char> *cdlls = new TList <char> ();
	for(i=0; i < cbNeeded / sizeof(HMODULE); i++)
	{
		if(wstr->hModules[i] == 0) 
			break;
		lModule = GetModuleFileName(wstr->hModules[i], wstr->moduleName, sizeof(wstr->moduleName));
		if(lModule == 0)
			break;
		if(wstr->hModules[i] == get_HINSTANCE())
		{
			continue;
		}
		lPath = lModule - 1;
		while(wstr->moduleName[lPath] != '\\')
		{
			lPath--;
		}
		if(lPath == lBaseModule && memcmp(wstr->baseModuleName, wstr->moduleName, lPath)==0)
		{
			cdlls->Add(wstr->moduleName + lPath + 1, lModule - lPath);
		}
		else
		{
			wdlls->Add(wstr->moduleName + lPath + 1, lModule - lPath);
		}
	}
	
	delete wstr;

	char *cdllname, *windllname;
	cdlls->MoveFirst();
	while((cdllname=cdlls->Next()))
	{
		wdlls->MoveFirst();
		while((windllname = wdlls->Next()))
		{
			proc(cdllname, windllname, parameters);
		}
	}
			
	delete cdlls;
	delete wdlls;
}
/***************** получить адрес функции ***********************************/
int Replacement::get_FunThunkAddr(LPCSTR dll_1c, LPCSTR dll_win, LPCSTR funName, void * pFunc, void * parameters)
{
	if(_stricmp(funName, ((ENUMFUNCTIONSPARS*)parameters)->funName) == 0)
	{
		((ENUMFUNCTIONSPARS*)parameters)->pFunc = pFunc;
		return 1;
	}
	else
		return 0;
}
void* Replacement::GetThunkAddress(LPCSTR strDll, LPCSTR strExporter, LPCSTR strFunc)
{
	ENUMFUNCTIONSPARS pars = {0};
	pars.funName = strFunc;

	Replacement::EnumFunctions(strDll, strExporter, Replacement::get_FunThunkAddr, &pars);

	return pars.pFunc;
}
//******** ReplaceFun ******************************/
void* Replacement::ReplaceFunctionAddr(LPCSTR strDll, LPCSTR strExporter, LPCSTR strFunc, void * funAddress)
{
	DWORD op;
	void * pFunc;
	REPLACEMENT *repl = NULL;

	pFunc = Replacement::GetThunkAddress(strDll, strExporter, strFunc);
	if(pFunc) 
	{
		repl = new REPLACEMENT();
		//repl->m_oldptr  = *pFunc;
		//repl->m_funplace = pFunc;
		memmove(&repl->m_oldptr, pFunc, sizeof(void*));
		repl->m_funplace = pFunc;

		BOOL retv = VirtualProtect(pFunc, sizeof(void FAR*), PAGE_EXECUTE_READWRITE, &op);
		if(retv)
		{
			// *(repl->m_funplace) = reinterpret_cast<DWORD>(funAddress);
			memmove(pFunc, &funAddress, sizeof(void*));
			//WriteProcessMemory(GetCurrentProcess(), repl->m_funplace, &funAddress, sizeof(DWORD), NULL);
			VirtualProtect(pFunc, sizeof(void FAR*), op, &op);
		}
		else
		{
			delete repl;
			repl = NULL;
		}
	}
	return repl;
}