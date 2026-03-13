#pragma once
#include "ComBase.h"

class CWinFunctions :
	public ComBase
{
public:
	CWinFunctions();
	~CWinFunctions();
	HRESULT __stdcall Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *);
};

