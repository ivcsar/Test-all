#pragma once
#include "..\stdafx.h"

class CHandler
{
	IDispatch *m_pDispatch;
	HRESULT m_rc;
	DISPID m_Dispid;
public:
	CHandler(LPDISPATCH pObject, BSTR method);
	~CHandler();
	void CallHandler(int code, WPARAM wParam, LPARAM lParam);
};

class CHandlerParent
{
	void *m_parent;
	TList <CHandler> *m_handlers;
public:
	CHandlerParent(void *parent);
	~CHandlerParent();
	int Compare(void * ptr);
	void CallHandlers(int code, WPARAM wParam, LPARAM lParam);
	void Add(LPDISPATCH pObject, BSTR method);
	BOOL IsEmpty();
};
