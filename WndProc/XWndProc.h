///////////// CXWndProc 
#pragma once
#include "..\stdafx.h"

#pragma pack(push, 1)
struct KEYBOARD_LPARAM
{
	unsigned int count		:16;	// 0-15 The repeat count. The value is the number of times the keystroke is repeated as a result of the user's holding down the key.
	unsigned int code		: 8;		//16-23 The scan code. The value depends on the OEM.
	unsigned int isExtended : 1;	//   24 Indicates whether the key is an extended key, such as a function key or a key on the numeric keypad. The value is 1 if the key is an extended key; otherwise, it is 0.
	unsigned int reserved	: 4;	//25-28 Reserved.
	unsigned int isAlt		: 1;	//   29 The context code. The value is 1 if the ALT key is down; otherwise, it is 0.
	unsigned int preKeyState: 1;	//   30 The previous key state. The value is 1 if the key is down before the message is sent; it is 0 if the key is up.
	unsigned int transState	: 1;	//   31 The transition state. The value is 0 if the key is being pressed and 1 if it is being released.
};
//struct KEYBOARD_LPARAM
//{
//	int transState         : 1;	//   31 The transition state. The value is 0 if the key is being pressed and 1 if it is being released.
//	int preKeyState:1;	//   30 The previous key state. The value is 1 if the key is down before the message is sent; it is 0 if the key is up.
//	bool isAlt:1;		//   29 The context code. The value is 1 if the ALT key is down; otherwise, it is 0.
//	unsigned int reserved  :  4;	//25-28 Reserved.
//	bool isExtended		   :  1;	//   24 Indicates whether the key is an extended key, such as a function key or a key on the numeric keypad. The value is 1 if the key is an extended key; otherwise, it is 0.
//	unsigned __int8  code  :  8;	//16-23 The scan code. The value depends on the OEM.
//	unsigned __int16 count : 16;	// 0-15 The repeat count. The value is the number of times the keystroke is repeated as a result of the user's holding down the key.
#pragma pack(pop)

class CXWndProc : public IDispatch
{
	LONG m_nRef;
public:
	CXWndProc();
	~CXWndProc();
protected:
	HRESULT __stdcall QueryInterface(const IID &riid, void **ppv);
	ULONG __stdcall AddRef(void);
	ULONG __stdcall Release(void);
	HRESULT __stdcall GetTypeInfoCount(UINT *pInfoCount);
	HRESULT __stdcall GetTypeInfo(UINT nType, LCID lcid, ITypeInfo **pInfo);
public:
	HRESULT __stdcall GetIDsOfNames(const IID &, LPOLESTR *methodName, UINT nMethodName, LCID lcid, DISPID * dispid);
	HRESULT __stdcall Invoke(DISPID nf, const IID &, LCID, WORD, DISPPARAMS * parameters, VARIANT * pVarResult, EXCEPINFO *pExcepInfo, UINT *);
private:
	HRESULT SetWndProcedure(DISPPARAMS * parameters, VARIANT * pVar);
	HRESULT SetKeyboardProc(DISPPARAMS * parameters);
	void SetCaption(BSTR bstrCaption);
};