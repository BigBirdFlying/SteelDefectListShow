// Defect_List_Show.cpp : ����Ӧ�ó��������Ϊ��
//

#include "stdafx.h"
#include "Defect_List_Show.h"
#include "Defect_List_ShowDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CDefect_List_ShowApp

BEGIN_MESSAGE_MAP(CDefect_List_ShowApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CDefect_List_ShowApp ����

CDefect_List_ShowApp::CDefect_List_ShowApp()
{
	// TODO: �ڴ˴���ӹ�����룬
	// ��������Ҫ�ĳ�ʼ�������� InitInstance ��
}


// Ψһ��һ�� CDefect_List_ShowApp ����

CDefect_List_ShowApp theApp;


// CDefect_List_ShowApp ��ʼ��

BOOL CDefect_List_ShowApp::InitInstance()
{
	// ���һ�������� Windows XP �ϵ�Ӧ�ó����嵥ָ��Ҫ
	// ʹ�� ComCtl32.dll �汾 6 ����߰汾�����ÿ��ӻ���ʽ��
	//����Ҫ InitCommonControlsEx()�����򣬽��޷��������ڡ�
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// ��������Ϊ��������Ҫ��Ӧ�ó�����ʹ�õ�
	// �����ؼ��ࡣ
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();

	///
	//����������
	m_hMutex = ::CreateMutex(NULL, FALSE, _T("Defect_List_ShowApp"));
	//�жϻ������Ƿ����
	if (GetLastError() == ERROR_ALREADY_EXISTS)	
	{
		AfxMessageBox(L"Ӧ�ó����Ѿ����С�");
		return FALSE;	
	}
	else
	{
		//AfxMessageBox("Ӧ�ó����һ�����С�");
		;
	}
	///

	AfxEnableControlContainer();

	// ��׼��ʼ��
	// ���δʹ����Щ���ܲ�ϣ����С
	// ���տ�ִ���ļ��Ĵ�С����Ӧ�Ƴ�����
	// ����Ҫ���ض���ʼ������
	// �������ڴ洢���õ�ע�����
	// TODO: Ӧ�ʵ��޸ĸ��ַ�����
	// �����޸�Ϊ��˾����֯��
	SetRegistryKey(_T("Ӧ�ó��������ɵı���Ӧ�ó���"));

	CDefect_List_ShowDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: �ڴ˷��ô����ʱ��
		//  ��ȷ�������رնԻ���Ĵ���
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: �ڴ˷��ô����ʱ��
		//  ��ȡ�������رնԻ���Ĵ���
	}

	if (!AfxOleInit())
	{
		AfxMessageBox(L"excel");
		return FALSE;
	}

	// ���ڶԻ����ѹرգ����Խ����� FALSE �Ա��˳�Ӧ�ó���
	//  ����������Ӧ�ó������Ϣ�á�
	return FALSE;
}
