#pragma once
#include "combase.h"
#include "..\XNative\NativeBase.h"

class CDest :
	public CNativeBase,
	public ComBase 
{
	wchar_t *m_str;
public:
	CDest(void);
	~CDest(void);
    bool ADDIN_API RegisterExtensionAs(WCHAR_T**);
    virtual bool ADDIN_API CallAsProc(const long lMethodNum,
                    tVariant* paParams, const long lSizeArray);
    virtual bool ADDIN_API CallAsFunc(const long lMethodNum,
                tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray);

	HRESULT __stdcall Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *);
};

