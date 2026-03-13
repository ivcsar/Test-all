#include "WinFunctions.h"

enum ENUM_METHOD_NUMBERS
{
	emSleep = 1,
	emNotifyIcon 
};


CMethod WinFunctions[] =
{
	CMethod(L"Sleep", emSleep, 1, false),
	CMethod(L"NotifyIcon", emNotifyIcon, 0, false),
};

// функция создания объекта
STDAPI GetWinFunctions(LPVOID *ppv)
{
	CWinFunctions *pWinFunctions = new CWinFunctions();
	*ppv = (void*)pWinFunctions;
	return S_OK;
}

CWinFunctions::CWinFunctions()
{
	AddRef();
	m_pComMethods = WinFunctions;
	m_nComMethodCount = sizeof(WinFunctions) / sizeof(CMethod);
}

CWinFunctions::~CWinFunctions()
{
}

HRESULT CWinFunctions::Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *)
{
	HRESULT rc = S_FALSE;
	VARIANT var = { 0 };

	if (nf == emSleep)
	{
		::Sleep(parameters->rgvarg[0].pvarVal->lVal);
		rc = S_OK;
	}
	else if (nf == emNotifyIcon)
	{
		NOTIFYICONDATA nidata = { 0 };
		::Shell_NotifyIcon(NIM_ADD, &nidata);
		rc = S_OK;
	}
	return rc;
}
