#pragma once
#include "..\stdafx.h"
#include "TListKey.h"

class CHandlerParent;
class CXWndProc;
class CHookHandler
{
public:
	HHOOK m_hook;
	HOOKPROC m_HookProc;
	TListKey <CHandlerParent, CXWndProc*> *m_parents;
public:
	CHookHandler(HHOOK hook, HOOKPROC proc);
	~CHookHandler();
	int Compare(HOOKPROC hookProc)
	{
		if(m_HookProc == hookProc)
			return 0;
		else
			return 1;
	}
	void AddHandler(void *parent, IDispatch *pDispatch, BSTR method);
	void RemoveHandlerParent(void *parent);
	void CallWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	void CallHandlers(int code, WPARAM wParam, LPARAM lParam);
	BOOL IsEmpty();
};