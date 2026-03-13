#pragma once
#include "..\XNative\NativeBase.h"

#ifndef __XNATIVE_H__
#define __XNATIVE_H__

// class CAddInNative
class CAddInNative : public CNativeBase
{
public:
    CAddInNative(void);
    virtual ~CAddInNative();
    // ILanguageExtenderBase
    virtual bool ADDIN_API RegisterExtensionAs(WCHAR_T**);
    virtual bool ADDIN_API GetParamDefValue(const long lMethodNum, const long lParamNum,
                            tVariant *pvarParamDefValue);   
    virtual bool ADDIN_API CallAsProc(const long lMethodNum,
                    tVariant* paParams, const long lSizeArray);
    virtual bool ADDIN_API CallAsFunc(const long lMethodNum,
                tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray);
private:
	void Unload(wchar_t *dllname);
	//HRESULT TypeInfo(wchar_t *dllname, CLSID & clsid);
};
#endif //__ADDINNATIVE_H__
