
#pragma once
#ifndef __NATIVEBASE_H__
#define __NATIVEBASE_H__
#define BASE_ERRNO  7

#pragma once

#include "stdafx.h"

#include "..\..\include\AddInDefBase.h"
#include "..\..\include\IMemoryManager.h"
#include "..\..\include\ComponentBase.h"

struct CMethod
{
	wchar_t *method;
	int lw;
	DISPID id;
	int nParams;
	bool hr;
	CMethod(wchar_t *pMethod, DISPID memberId, int pNParams, bool pHasRetValue)
	{
		method = pMethod;
		id = memberId;
		nParams = pNParams;
		hr = pHasRetValue;
		lw = wcslen(method);
	}
};

class CNativeBase : public IComponentBase
{
public:
    CNativeBase(void);
    virtual ~CNativeBase();
    // IInitDoneBase
    virtual bool ADDIN_API Init(void*);
    virtual bool ADDIN_API setMemManager(void* mem);
    virtual long ADDIN_API GetInfo();
    virtual void ADDIN_API Done();
    // ILanguageExtenderBase
	virtual bool ADDIN_API RegisterExtensionAs(WCHAR_T**) = 0;
    virtual long ADDIN_API GetNProps();
    virtual long ADDIN_API FindProp(const WCHAR_T* wsPropName);
    virtual const WCHAR_T* ADDIN_API GetPropName(long lPropNum, long lPropAlias);
    virtual bool ADDIN_API GetPropVal(const long lPropNum, tVariant* pvarPropVal);
    virtual bool ADDIN_API SetPropVal(const long lPropNum, tVariant* varPropVal);
    virtual bool ADDIN_API IsPropReadable(const long lPropNum);
    virtual bool ADDIN_API IsPropWritable(const long lPropNum);
    virtual long ADDIN_API GetNMethods();
    virtual long ADDIN_API FindMethod(const WCHAR_T* wsMethodName);
    virtual const WCHAR_T* ADDIN_API GetMethodName(const long lMethodNum, 
                            const long lMethodAlias);
    virtual long ADDIN_API GetNParams(const long lMethodNum);
    virtual bool ADDIN_API GetParamDefValue(const long lMethodNum, const long lParamNum,
                            tVariant *pvarParamDefValue);   
    virtual bool ADDIN_API HasRetVal(const long lMethodNum);
    virtual bool ADDIN_API CallAsProc(const long lMethodNum,
                    tVariant* paParams, const long lSizeArray);
    virtual bool ADDIN_API CallAsFunc(const long lMethodNum,
                tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray);
    // LocaleBase
    virtual void ADDIN_API SetLocale(const WCHAR_T* loc);
protected:
    void addError(uint32_t wcode, const wchar_t* source, 
                  const wchar_t* descriptor, long code);
    // Attributes
    IAddInDefBase      *m_iConnect;
    IMemoryManager     *m_iMemory;

	struct CMethod *m_cMethods;
	DWORD   m_nMethods;
	void CallException(DWORD errc, wchar_t*);
	void SetWSTR(tVariant *ptVar, wchar_t*);
	bool AllocWSTR(WCHAR_T ** pwstr, WCHAR_T *wstr);
};
#endif //
