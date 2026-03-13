#pragma once
#include "..\ComBase.h"

class CLock : public ComBase
{
public:
	CLock(void);
	~CLock(void);

	HANDLE m_hMutex;

	HRESULT __stdcall Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *);
	int FileInformation(BSTR filename);
};

