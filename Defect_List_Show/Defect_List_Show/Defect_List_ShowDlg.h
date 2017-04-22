// Defect_List_ShowDlg.h : 头文件
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

// CDefect_List_ShowDlg 对话框
class CDefect_List_ShowDlg : public CDialog
{
// 构造
public:
	CDefect_List_ShowDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_DEFECT_LIST_SHOW_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
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
