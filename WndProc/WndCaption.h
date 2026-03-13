#ifndef _WNDCAPTION_H_
#define _WNDCAPTION_H_
#include <windows.h>

class CXWndCaption
{
	HWND m_hwnd;
	wchar_t * m_wCaption;
	WNDPROC m_WndProc;
	static CXWndCaption *m_Ptr;
public:
	CXWndCaption(BSTR wCaption, VARIANT *wincap);
	~CXWndCaption();
	static void Release();
	static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	static BOOL setCaptionHook(BSTR caption, VARIANT *wincap);
	// by Hook - not use
	static LRESULT CALLBACK CallGetMsgProc(int code, WPARAM wParam, LPARAM lParam);
	static void SetWndCaptionByHook();
};
#endif
