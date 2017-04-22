// Defect_List_ShowDlg.h : ͷ�ļ�
//

#pragma once
#include "afxcmn.h"
#include "ListCtrlCl.h"
#include "afxwin.h"
#include "afxdtctl.h"

//#include "ArNTISDataStruct.h"
//#include "CommonFunc.h"
//#include "ArNTCommonClass.h"
//#include "ArNTOnlineDetectClass.h"

// CDefect_List_ShowDlg �Ի���
class CDefect_List_ShowDlg : public CDialog
{
// ����
public:
	CDefect_List_ShowDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_DEFECT_LIST_SHOW_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	//ArNTXMLOperator		m_XMLOperator;
	//LONGSTR strSteelWarn;
	CString m_strDBServer;

	int ReadConfigFile();

	void ReadDBAccess();
	void ReadDBSqlSever();
public:
	CListCtrlCl m_ListCtrl;
	afx_msg void OnBnClickedQuery();
	CComboBox c_combo_begin_hour;
	CComboBox c_combo_end_hour;
	CComboBox c_combo_end_min;
	CComboBox c_combo_begin_min;
	CDateTimeCtrl m_DateTimeStart;
	CDateTimeCtrl m_DateTimeEnd;
	afx_msg void OnBnClickedButtonReport();
	afx_msg void OnNMClickListDefects(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnClose();
	afx_msg void OnBnClickedRadioYes();
	afx_msg void OnBnClickedRadioNo();
	afx_msg void OnBnClickedButtonSet();
	CListCtrl m_DefectList;
};
