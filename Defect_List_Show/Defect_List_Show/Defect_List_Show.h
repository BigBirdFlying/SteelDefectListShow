// Defect_List_Show.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CDefect_List_ShowApp:
// �йش����ʵ�֣������ Defect_List_Show.cpp
//

class CDefect_List_ShowApp : public CWinApp
{
public:
	CDefect_List_ShowApp();

// ��д
	public:
	virtual BOOL InitInstance();

protected:
	HANDLE m_hMutex;

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CDefect_List_ShowApp theApp;