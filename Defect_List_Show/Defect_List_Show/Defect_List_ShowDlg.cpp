// Defect_List_ShowDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "math.h"

#include "Defect_List_Show.h"
#include "Defect_List_ShowDlg.h"

#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRange.h"
#include "CFont0.h"
#include "Cnterior.h"
#include "CBorders.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CDefect_List_ShowDlg �Ի���




CDefect_List_ShowDlg::CDefect_List_ShowDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDefect_List_ShowDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_strDBServer=L"";
}

void CDefect_List_ShowDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_DEFECTS, m_ListCtrl);
	DDX_Control(pDX, IDC_COMBO_HOUR_START, c_combo_begin_hour);
	DDX_Control(pDX, IDC_COMBO_HOUR_END, c_combo_end_hour);
	DDX_Control(pDX, IDC_COMBO_MIN_END, c_combo_end_min);
	DDX_Control(pDX, IDC_COMBO_MIN_START, c_combo_begin_min);
	DDX_Control(pDX, IDC_DATETIMEPICKER_START, m_DateTimeStart);
	DDX_Control(pDX, IDC_DATETIMEPICKER_END, m_DateTimeEnd);
	DDX_Control(pDX, IDC_LIST_DEFECTLIST, m_DefectList);
}

BEGIN_MESSAGE_MAP(CDefect_List_ShowDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON_QUERY, &CDefect_List_ShowDlg::OnBnClickedQuery)
	ON_BN_CLICKED(IDC_BUTTON_REPORT, &CDefect_List_ShowDlg::OnBnClickedButtonReport)
	ON_NOTIFY(NM_CLICK, IDC_LIST_DEFECTS, &CDefect_List_ShowDlg::OnNMClickListDefects)
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDC_RADIO_YES, &CDefect_List_ShowDlg::OnBnClickedRadioYes)
	ON_BN_CLICKED(IDC_RADIO_NO, &CDefect_List_ShowDlg::OnBnClickedRadioNo)
	ON_BN_CLICKED(IDC_BUTTON_SET, &CDefect_List_ShowDlg::OnBnClickedButtonSet)
END_MESSAGE_MAP()


// CDefect_List_ShowDlg ��Ϣ�������

BOOL CDefect_List_ShowDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	ReadConfigFile();
	//�õ���ʾ����С
	int   cx,cy;
	cx   =   GetSystemMetrics(SM_CXSCREEN);
	cy   =   GetSystemMetrics(SM_CYSCREEN)-50;
	//����MoveWindow
	CRect   rcTemp;
	rcTemp.BottomRight()   =   CPoint(cx,   cy);
	rcTemp.TopLeft()   =   CPoint(0, 0);
	MoveWindow(&rcTemp);     

	CRect   rcTemp0;
	rcTemp0.TopLeft()=CPoint(0, cy/4);
	rcTemp0.BottomRight()=CPoint(cx,  cy);
	m_ListCtrl.MoveWindow(&rcTemp0);  


	//�鿴ģ��
	CRect rect_Info;
	this->GetWindowRect(&rect_Info);
	ScreenToClient(rect_Info);
	rect_Info.bottom/=4;
	GetDlgItem(IDC_STATIC_GROUP_QUERY)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/64,(float)(rect_Info.bottom-rect_Info.top)/32, (float)(rect_Info.right-rect_Info.left)*8/16, (float)(rect_Info.bottom-rect_Info.top)*26/32);
	CRect rect_query;
	GetDlgItem(IDC_STATIC_GROUP_QUERY)->GetWindowRect(&rect_query);
	GetDlgItem(IDC_STATIC_BEGIN_TIME)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*2/12, 
																				100, 
																				25);
	GetDlgItem(IDC_STATIC_END_TIME)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*5/12, 
																				100, 
																				25);
	GetDlgItem(IDC_DATETIMEPICKER_START)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*7/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*2/12, 
																				100, 
																				25);
	GetDlgItem(IDC_DATETIMEPICKER_END)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*7/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*5/12, 
																				100, 
																				25);
	GetDlgItem(IDC_COMBO_HOUR_START)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*12/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*2/12, 
																				50, 
																				25);
	GetDlgItem(IDC_COMBO_HOUR_END)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*12/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*5/12, 
																				50, 
																				25);
	GetDlgItem(IDC_STATIC_BEGIN_HOUR)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*15/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*2/12, 
																				50, 
																				25);
	GetDlgItem(IDC_STATIC_END_HOUR)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*15/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*5/12, 
																				50, 
																				25);
	GetDlgItem(IDC_COMBO_MIN_START)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*18/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*2/12, 
																				50, 
																				25);
	GetDlgItem(IDC_COMBO_MIN_END)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*18/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*5/12, 
																				50, 
																				25);
	GetDlgItem(IDC_STATIC_BEGIN_MIN)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*21/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*2/12, 
																				50, 
																				25);
	GetDlgItem(IDC_STATIC_END_MIN)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*21/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*5/12, 
																				50, 
																				25);
	GetDlgItem(IDC_BUTTON_QUERY)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*27/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*2/12, 
																				100, 
																				25);
	GetDlgItem(IDC_BUTTON_REPORT)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*27/32,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*5/12, 
																				100, 
																				25);
	//������
	GetDlgItem(IDC_STATIC_STEEL)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)/32,
																			(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*8/12, 
																			100, 
																			25);
	GetDlgItem(IDC_EDIT_STEEL)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*7/32,
																			(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*8/12, 
																			100, 
																			25);
	GetDlgItem(IDC_STATIC_ALERM)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*12/32,
																			(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*8/12, 
																			50, 
																			25);
	GetDlgItem(IDC_EDIT_ALERM)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*15/32,
																			(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*8/12, 
																			50, 
																			25);
	GetDlgItem(IDC_RADIO_YES)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*18/32,
																			(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*8/12, 
																			50, 
																			25);
	GetDlgItem(IDC_RADIO_NO)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*21/32,
																			(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*8/12, 
																			50, 
																			25);
	GetDlgItem(IDC_BUTTON_SET)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)*27/32,
																			(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_query.bottom-rect_query.top)*8/12, 
																			100, 
																			25);

	GetDlgItem(IDC_STATIC_GROUP_DEFECTLIST)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*34/64,(float)(rect_Info.bottom-rect_Info.top)/32, (float)(rect_Info.right-rect_Info.left)*29/64, (float)(rect_Info.bottom-rect_Info.top)*26/32);
	CRect rect_defectlist;
	GetDlgItem(IDC_STATIC_GROUP_DEFECTLIST)->GetWindowRect(&rect_defectlist);
	GetDlgItem(IDC_LIST_DEFECTLIST)->MoveWindow( (float)(rect_Info.right-rect_Info.left)*1/128+(float)(rect_query.right-rect_query.left)+(float)(rect_defectlist.left-rect_query.right)+(float)(rect_defectlist.right-rect_defectlist.left)*3/64,
																				(float)(rect_Info.bottom-rect_Info.top)/32+(float)(rect_defectlist.bottom-rect_defectlist.top)*3/24, 
																				(float)(rect_defectlist.right-rect_defectlist.left)*60/64, 
																				(float)(rect_defectlist.bottom-rect_defectlist.top)*20/24);


	//m_ListCtrl.SetWindowPos(NULL,0,0,rcTemp0.Width(),rcTemp0.Height(),SWP_NOZORDER|SWP_NOMOVE);
	m_ListCtrl.SetHeaderHeight(1.5);          //����ͷ���߶�
	m_ListCtrl.SetHeaderFontHW(16,0);         //����ͷ������߶�,�Ϳ��,0��ʾȱʡ������Ӧ 
	m_ListCtrl.SetHeaderTextColor(RGB(255,200,100)); //����ͷ��������ɫ
	m_ListCtrl.SetHeaderBKColor(128,255,255,8); //����ͷ������ɫ

	m_ListCtrl.SetBkColor(RGB(128,128,128));        //���ñ���ɫRGB(128,64,0)

	m_ListCtrl.SetRowHeigt(25);               //�����и߶�
	m_ListCtrl.SetHeaderHeight(1.5);          //����ͷ���߶�
	m_ListCtrl.SetFontHW(15,0);               //��������߶ȣ��Ϳ��,0��ʾȱʡ���

	int n=9;
	float m=0.05;
	m_ListCtrl.InsertColumn(0,_T("���"),LVCFMT_CENTER,rcTemp0.Width()*m);
	m_ListCtrl.InsertColumn(1,_T("��ˮ��"),LVCFMT_CENTER,rcTemp0.Width()*m);
	m_ListCtrl.InsertColumn(2,_T("�ְ��"),LVCFMT_CENTER,rcTemp0.Width()*2*m);
	m_ListCtrl.InsertColumn(3,_T("���"),LVCFMT_CENTER,rcTemp0.Width()*m);
	m_ListCtrl.InsertColumn(4,_T("���"),LVCFMT_CENTER,rcTemp0.Width()*m);
	m_ListCtrl.InsertColumn(5,_T("����"),LVCFMT_CENTER,rcTemp0.Width()*m);
	//m_ListCtrl.InsertColumn(6,_T("����ʱ��"),LVCFMT_CENTER,rcTemp0.Width()*2.5*m);
	m_ListCtrl.InsertColumn(6,_T("���ʱ��"),LVCFMT_CENTER,rcTemp0.Width()*2.5*m);
	m_ListCtrl.InsertColumn(7,_T("ȱ���Զ�����"),LVCFMT_CENTER,rcTemp0.Width()*4.5*m);
	m_ListCtrl.InsertColumn(8,_T("�˹�ȱ������"),LVCFMT_CENTER,rcTemp0.Width()*6*m);

	CRect rect_temp2;
	GetDlgItem(IDC_LIST_DEFECTLIST)->GetWindowRect(&rect_temp2);
	m_DefectList.InsertColumn(0,_T("���"),LVCFMT_CENTER,rect_temp2.Width()*0.1);
	m_DefectList.InsertColumn(1,_T("���"),LVCFMT_CENTER,rect_temp2.Width()*0.2);
	m_DefectList.InsertColumn(2,_T("����"),LVCFMT_CENTER,rect_temp2.Width()*0.1);
	m_DefectList.InsertColumn(3,_T("����λ��(mm)"),LVCFMT_CENTER,rect_temp2.Width()*0.2);
	m_DefectList.InsertColumn(4,_T("����λ��(mm)"),LVCFMT_CENTER,rect_temp2.Width()*0.2);
	m_DefectList.InsertColumn(5,_T("���(mm2)"),LVCFMT_CENTER,rect_temp2.Width()*0.2);

	//��ʼ��
	for(int i= 23; i >=0; i--)
	{
		CString strInfo;
		strInfo.Format(L"%d", i);
		c_combo_begin_hour.InsertString(0,strInfo);
		c_combo_end_hour.InsertString(0,strInfo);
	}
	for(int i= 59; i >=0; i--)
	{
		CString strInfo;
		strInfo.Format(L"%d", i);
		c_combo_begin_min.InsertString(0,strInfo);
		c_combo_end_min.InsertString(0,strInfo);
	}
	c_combo_begin_hour.SetCurSel(0);
	c_combo_end_hour.SetCurSel(23);
	c_combo_begin_min.SetCurSel(0);
	c_combo_end_min.SetCurSel(59);

	SetWindowLong(m_ListCtrl.m_hWnd ,GWL_EXSTYLE,WS_EX_CLIENTEDGE);
	m_ListCtrl.SetExtendedStyle(LVS_EX_GRIDLINES);                     //������չ���Ϊ����
	::SendMessage(m_ListCtrl.m_hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE,LVS_EX_FULLROWSELECT, LVS_EX_FULLROWSELECT);
	//PostMessage(WM_COMMAND,MAKEWPARAM(IDC_BUTTON_SET,BN_CLICKED));
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CDefect_List_ShowDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CDefect_List_ShowDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
				//----------���������ñ���ͼƬ----------------------------
		CPaintDC dc(this); 
		CRect rect;
		GetClientRect(&rect);
		CDC   dcMem;   
		dcMem.CreateCompatibleDC(&dc);   
		CBitmap   bmpBackground;   
		bmpBackground.LoadBitmap(IDC_BACKGROUND);   //IDB_BITMAP�����Լ���ͼ��Ӧ��ID 

		BITMAP   bitmap;   
		bmpBackground.GetBitmap(&bitmap);   
		CBitmap   *pbmpOld=dcMem.SelectObject(&bmpBackground);   
		dc.StretchBlt(0,0,rect.Width(),rect.Height(),&dcMem,0,0,bitmap.bmWidth,bitmap.bmHeight,SRCCOPY);  

		//��仰�����±߲���ʹͼ����ʾ����
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CDefect_List_ShowDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CDefect_List_ShowDlg::OnBnClickedQuery()
{
	CString strIP=::GetCommandLine();
	int iPos=strIP.Find(L"-IP");
	if(iPos>0)
	{
		CString strIPdata = _T("");
		CString S=L"-IP ";
		int len=S.GetLength();
		strIPdata = strIP.Right(strIP.GetLength()-iPos-len);
		m_strDBServer.Format(L"%s",strIPdata);
		if(m_strDBServer != L"")
		{
			ReadDBSqlSever();
		}
		else
		{
			AfxMessageBox(L"IP��ַΪ��");
		}
	}
	else
	{
		AfxMessageBox(L"δָ��IP");
	}
	//ReadDBAccess();
	
	
}

void CDefect_List_ShowDlg::ReadDBSqlSever()
{
	m_ListCtrl.DeleteAllItems();
	//m_ListCtrl.SetBkColor(RGB(0,0,255));	
	//m_ListCtrl.SetTextColor(RGB(255,255,0));

	//CString strDBServer=L"192.168.0.100";
	//CString strDBServer=L"127.0.0.1";//172.16.17.252
	CString strDBServer=L"";
	strDBServer.Format(L"%s",m_strDBServer);
	CString strDBName=L"SteelRecord";
	CString strUser=L"ARNTUSER";
	CString strPassWd=L"ARNTUSER";

	::CoInitialize(NULL);
	CString strTableContent;
 
	_ConnectionPtr m_pConnection;
	_variant_t RecordsAffected;
	_RecordsetPtr m_pRecordset;
 
	try
	{
		m_pConnection.CreateInstance(__uuidof(Connection));
		CString sql;
		//sql.Format(L"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s",strSteelWarn);
		sql.Format(L"Provider=SQLOLEDB.1;Password=%s;Persist Security Info=True; User ID=%s;Initial Catalog=%s;Data Source=%s",strPassWd,strUser,strDBName,strDBServer);	
		_bstr_t strCmd=sql;
		m_pConnection->Open(strCmd,"","",adModeUnknown);
	}
	catch(_com_error e)
	{
		CString errormessage;
		errormessage.Format(CString("�������ݿ�ʧ��!\r������Ϣ:%s"),e.ErrorMessage());
		AfxMessageBox(errormessage);
		return;
	}

	SYSTEMTIME tmStart;
	m_DateTimeStart.GetTime(&tmStart);
	SYSTEMTIME tmEnd;
	m_DateTimeEnd.GetTime(&tmEnd);

	int iStartHour = GetDlgItemInt(IDC_COMBO_HOUR_START);
	int iEndHour = GetDlgItemInt(IDC_COMBO_HOUR_END);
	int iStartMin = GetDlgItemInt(IDC_COMBO_MIN_START);
	int iEndMin = GetDlgItemInt(IDC_COMBO_MIN_END);
 
	try
	{
		m_pRecordset.CreateInstance("ADODB.Recordset"); //ΪRecordset���󴴽�ʵ��
		CString sql;
		//m_DefectsMDBSet.strDefectsMDB[i].iDefectsNumMax=200;
		//sql.Format(L"select * from steel where TopDetectTime between \'%04d-%02d-%02d %02d:%02d:%02d.000\' and \'%04d-%02d-%02d %02d:%02d:%02d.000\' order by TopDetectTime desc",tmStart.wYear, tmStart.wMonth, tmStart.wDay, iStartHour, iStartMin,0,tmEnd.wYear, tmEnd.wMonth, tmEnd.wDay, iEndHour,iEndMin,0);
		sql.Format(L"select steel.SequeceNo,steel.SteelID,steel.TopDetectTime,SteelID.Thick,SteelID.Width,SteelID.Length,SteelID.AddTime,SteelGrade.Dsc\
					from SteelID inner join steel on SteelID.ID=steel.SteelID inner join SteelGrade on SteelGrade.SequeceNo=steel.SequeceNo \
					where TopDetectTime between \'%04d-%02d-%02d %02d:%02d:%02d.000\' and \'%04d-%02d-%02d %02d:%02d:%02d.000\' order by TopDetectTime desc",
					tmStart.wYear, tmStart.wMonth, tmStart.wDay, iStartHour, iStartMin,0,tmEnd.wYear, tmEnd.wMonth, tmEnd.wDay, iEndHour,iEndMin,0);
		_bstr_t strCmd=sql;				
		
		m_pRecordset=m_pConnection->Execute(strCmd,&RecordsAffected,adCmdText);
		int a=0;
 
	}
	catch(_com_error &e)
	{
		AfxMessageBox(e.Description());
	}
 
	_variant_t vInfo;
	int Index=0;
	try
	{
		while(!m_pRecordset->adoEOF)
		{
            int iSequeceNo=m_pRecordset->GetCollect("SequeceNo");		
			CString  strSteelID=m_pRecordset->GetCollect("SteelID");
			CString  strDateTime=m_pRecordset->GetCollect("TopDetectTime");
			int iThick=m_pRecordset->GetCollect("Thick");
			int iWidth=m_pRecordset->GetCollect("Width");
			int iLength=m_pRecordset->GetCollect("Length");
			CString  strAddTime=m_pRecordset->GetCollect("AddTime");
			vInfo=m_pRecordset->GetCollect("Dsc");
			CString  strDsc=L"";
			if(vInfo.vt!=VT_NULL)
			{
				strDsc.Format(L"%s",vInfo.bstrVal);
			}

			if(strDsc.IsEmpty() != true)
			{
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(255,0,0));  //���õ�Ԫ��������ɫ
			}
			else
			{
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(0,255,0));  //���õ�Ԫ��������ɫ
			}
			//�����б�
			CString strIndex;
			strIndex.Format(L"%d",Index);
			m_ListCtrl.InsertItem(Index,strIndex);

			CString strSequeceNo;
			strSequeceNo.Format(L"%d",iSequeceNo);
			m_ListCtrl.SetItemText(Index,1,strSequeceNo);

			m_ListCtrl.SetItemText(Index,2,strSteelID);

			CString strThick;
			strThick.Format(L"%d",iThick);
			m_ListCtrl.SetItemText(Index,3,strThick);

			CString strWidth;
			strWidth.Format(L"%d",iWidth);
			m_ListCtrl.SetItemText(Index,4,strWidth);

			CString strLength;
			strLength.Format(L"%d",iLength);
			m_ListCtrl.SetItemText(Index,5,strLength);

			//m_ListCtrl.SetItemText(Index,6,strAddTime);

			m_ListCtrl.SetItemText(Index,6,strDateTime);

			m_ListCtrl.SetItemText(Index,8,strDsc);


			Index++;
			if(Index>1024)
			{
				break;
			}
			m_pRecordset->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		AfxMessageBox(e.Description());
	}
}

void CDefect_List_ShowDlg::ReadDBAccess()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	m_ListCtrl.DeleteAllItems();
	CString strDBName=L"SteelDefectList.mdb";

	::CoInitialize(NULL);
	CString strTableContent;
 
	_ConnectionPtr m_pConnection;
	_variant_t RecordsAffected;
	_RecordsetPtr m_pRecordset;
 
	try
	{
		m_pConnection.CreateInstance(__uuidof(Connection));
		CString sql;
		sql.Format(L"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s",strDBName);	
		_bstr_t strCmd=sql;
		m_pConnection->Open(strCmd,"","",adModeUnknown);
	}
	catch(_com_error e)
	{
		CString errormessage;
		errormessage.Format(CString("�������ݿ�ʧ��!\r������Ϣ:%s"),e.ErrorMessage());
		AfxMessageBox(errormessage);
		return;
	}

	SYSTEMTIME tmStart;
	m_DateTimeStart.GetTime(&tmStart);
	SYSTEMTIME tmEnd;
	m_DateTimeEnd.GetTime(&tmEnd);

	int iStartHour = GetDlgItemInt(IDC_COMBO_HOUR_START);
	int iEndHour = GetDlgItemInt(IDC_COMBO_HOUR_END);
	int iStartMin = GetDlgItemInt(IDC_COMBO_MIN_START);
	int iEndMin = GetDlgItemInt(IDC_COMBO_MIN_END);
 
	try
	{
		m_pRecordset.CreateInstance("ADODB.Recordset"); //ΪRecordset���󴴽�ʵ��
		CString sql;
		//m_DefectsMDBSet.strDefectsMDB[i].iDefectsNumMax=200;
		//sql.Format(L"select * from steel");
		sql.Format(L"select * from steel where DetectTime between \'%04d-%02d-%02d %02d:%02d:00\' and  \'%04d-%02d-%02d %02d:%02d:00\' order by DetectTime desc",tmStart.wYear, tmStart.wMonth, tmStart.wDay, iStartHour, iStartMin,tmEnd.wYear, tmEnd.wMonth, tmEnd.wDay, iEndHour,iEndMin);
		//sql.Format(L"select steel.SequeceNo,steel.SteelID,steel.TopDetectTime,SteelID.Thick,SteelID.Width,SteelID.Length,SteelID.AddTime,SteelGrade.Dsc\
		//			from SteelID inner join steel on SteelID.ID=steel.SteelID inner join SteelGrade on SteelGrade.SequeceNo=steel.SequeceNo \
		//			where TopDetectTime between \'%04d-%02d-%02d %02d:%02d:%02d.000\' and \'%04d-%02d-%02d %02d:%02d:%02d.000\' order by TopDetectTime desc",
		//			tmStart.wYear, tmStart.wMonth, tmStart.wDay, iStartHour, iStartMin,0,tmEnd.wYear, tmEnd.wMonth, tmEnd.wDay, iEndHour,iEndMin,0);
		_bstr_t strCmd=sql;				
		
		m_pRecordset=m_pConnection->Execute(strCmd,&RecordsAffected,adCmdText);
		int a=0;
 
	}
	catch(_com_error &e)
	{
		AfxMessageBox(e.Description());
	}
 
	_variant_t vInfo;
	int Index=0;
	try
	{
		while(!m_pRecordset->adoEOF)
		{
            int iSequeceNo=m_pRecordset->GetCollect("SequeceNo");		
			CString  strSteelID=m_pRecordset->GetCollect("SteelID");
			CString  strDateTime=m_pRecordset->GetCollect("DetectTime");
			int iThick=m_pRecordset->GetCollect("Thick");
			int iWidth=m_pRecordset->GetCollect("Width");
			int iLength=m_pRecordset->GetCollect("Length");
			CString  strAddTime=m_pRecordset->GetCollect("AddTime");
			vInfo=m_pRecordset->GetCollect("Dsc");
			CString  strDsc=L"";
			if(vInfo.vt!=VT_NULL)
			{
				strDsc.Format(L"%s",vInfo.bstrVal);
			}

			if(strDsc.IsEmpty() != true)
			{
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(255,0,0));  //���õ�Ԫ��������ɫ
			}
			else
			{
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(0,255,0));  //���õ�Ԫ��������ɫ
			}
			//�����б�
			CString strIndex;
			strIndex.Format(L"%d",Index);
			m_ListCtrl.InsertItem(Index,strIndex);

			CString strSequeceNo;
			strSequeceNo.Format(L"%d",iSequeceNo);
			m_ListCtrl.SetItemText(Index,1,strSequeceNo);

			m_ListCtrl.SetItemText(Index,2,strSteelID);

			CString strThick;
			strThick.Format(L"%d",iThick);
			m_ListCtrl.SetItemText(Index,3,strThick);

			CString strWidth;
			strWidth.Format(L"%d",iWidth);
			m_ListCtrl.SetItemText(Index,4,strWidth);

			CString strLength;
			strLength.Format(L"%d",iLength);
			m_ListCtrl.SetItemText(Index,5,strLength);

			m_ListCtrl.SetItemText(Index,6,strAddTime);

			m_ListCtrl.SetItemText(Index,7,strDateTime);

			m_ListCtrl.SetItemText(Index,9,strDsc);


			Index++;
			if(Index>1024)
			{
				break;
			}
			m_pRecordset->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		AfxMessageBox(e.Description());
	}
}


int CDefect_List_ShowDlg::ReadConfigFile()
{
	//LONGSTR strConfigFile = {0};
	//LONGSTR strAppPath = {0};
	//LONGSTR strAppName = {0};
	//CCommonFunc::GetAppPath(strAppPath, STR_LEN(strAppPath));
	//CCommonFunc::GetAppName(strAppName, STR_LEN(strAppName));
	////CCommonFunc::SafeStringPrintf(m_strLogPath, STR_LEN(m_strLogPath), L"%s\\Log\\%s", strAppPath, strAppName);
	////CCommonFunc::SafeStringPrintf(m_strDefectInfoPath, STR_LEN(m_strDefectInfoPath), L"%s\\Defects_Info", strAppPath, strAppName);
	//CCommonFunc::SafeStringPrintf(strConfigFile, STR_LEN(strConfigFile), L"%s\\Defect_List_Show.xml", strAppPath);
	//if(m_XMLOperator.LoadXML(strConfigFile) == ArNT_FALSE)
	//{
	//	CCommonFunc::ErrorBox(L"�޷����ز����ļ�:%s", strConfigFile);
	//	return ArNT_FALSE;
	//}

	//LONGSTR strValue = {0};	
	//LONGSTR strName = {0};

	////���ݿⲿ��
	////strSteelWarn={0};
	//CCommonFunc::SafeStringPrintf(strName, STR_LEN(strName), L"�ְ�ȱ����Ϣչʾ����#ȱ�ݱ�����Ϣ���ݿ�����#���ݿ�·��");
	//if(m_XMLOperator.GetValueByString(strName, strValue))
	//{
	//	CCommonFunc::SafeStringCpy(strSteelWarn, STR_LEN(strSteelWarn), strValue);
	//}

	return true;
}
void CDefect_List_ShowDlg::OnBnClickedButtonReport()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	if(m_ListCtrl.GetItemCount()<=0)
	{
		//return;
	}


	CString strFile;// = _T("D:\\WriteListToExcelTest.xlsx");
	SYSTEMTIME tmStart;
	m_DateTimeStart.GetTime(&tmStart);
	SYSTEMTIME tmEnd;
	m_DateTimeEnd.GetTime(&tmEnd);

	int iStartHour = GetDlgItemInt(IDC_COMBO_HOUR_START);
	int iEndHour = GetDlgItemInt(IDC_COMBO_HOUR_END);
	int iStartMin = GetDlgItemInt(IDC_COMBO_MIN_START);
	int iEndMin = GetDlgItemInt(IDC_COMBO_MIN_END);
	strFile.Format(L"D:\\%04d��%02d��%02d��%02dʱ%02d��__%04d��%02d��%02d��%02dʱ%02d��.xlsx",tmStart.wYear,tmStart.wMonth,tmStart.wDay,iStartHour,iStartMin,tmEnd.wYear,tmEnd.wMonth,tmEnd.wDay,iEndHour,iEndMin);

	COleVariant 
	covTrue((short)TRUE), 
	covFalse((short)FALSE), 
	covOptional((long)DISP_E_PARAMNOTFOUND,   VT_ERROR); 
 
	CApplication app;
	CWorkbook book;
	CWorkbooks books;
	CWorksheet sheet;
	CWorksheets sheets;
	CRange range;
	CFont0 font;
	Cnterior interior;//����ɫ
 
	if (!app.CreateDispatch(_T("Excel.Application")))
	{
		MessageBox(_T("Error!Creat Excel Application Server Faile!"));
		//exit(1);
	}
	books = app.get_Workbooks();
	book = books.Add(covOptional);
	sheets = book.get_Worksheets();
	sheet = sheets.get_Item(COleVariant((short)1));
	//�õ�ȫ��Cells 
	range.AttachDispatch(sheet.get_Cells()); 
	CString sText[]={_T("���"),_T("��ˮ��"),_T("�ְ��"),_T("���"),_T("���"),_T("����"),_T("���ʱ��"),_T("ȱ���Զ�����"),_T("�˹�ȱ������")};
	for (int setnum=0;setnum<m_ListCtrl.GetItemCount()+1;setnum++)
	{
		for (int num=0;num<9;num++)
		{
			if (!setnum)
			{
				range.put_Item(_variant_t((long)(setnum+1)),_variant_t((long)(num+1)),_variant_t(sText[num]));
				range.AttachDispatch(sheet.get_Range(COleVariant(L"A1"),COleVariant(L"I1")),TRUE);
				interior =range.get_Interior();//��ɫ
				interior.put_Color(_variant_t(100));
				font.AttachDispatch(range.get_Font());
				font.put_Color(COleVariant((long)0xFF0000)); 
				range.AttachDispatch(sheet.get_Cells()); 
			}
			else
			{
				range.put_Item(_variant_t((long)(setnum+1)),_variant_t((long)(num+1)),_variant_t(m_ListCtrl.GetItemText(setnum-1,num)));
				
				//range.AttachDispatch(sheet.get_Cells()); 
				////range.AttachDispatch((range.get_Item(COleVariant(long(1)), COleVariant(long(1)))).pdispVal);
				//font.AttachDispatch(range.get_Font());
				//font.put_Color(COleVariant((long)0x0000FF)); 
				
				CString strTemp;
				strTemp=m_ListCtrl.GetItemText(setnum-2,num);
				if(strTemp.Compare(L"����")==0)
				{
					CString strLeft,strRight;
					strLeft.Format(L"A%d",setnum+1);
					strRight.Format(L"I%d",setnum+1);
					range.AttachDispatch(sheet.get_Range(COleVariant(strLeft),COleVariant(strRight)),TRUE);
					interior =range.get_Interior();//��ɫ
					interior.put_Color(COleVariant((long)0x884400));
					font.AttachDispatch(range.get_Font());
					font.put_Color(COleVariant((long)0x0000FF)); 
					range.AttachDispatch(sheet.get_Cells()); 
				}
				else
				{
					CString strLeft,strRight;
					strLeft.Format(L"A%d",setnum+1);
					strRight.Format(L"I%d",setnum+1);
					range.AttachDispatch(sheet.get_Range(COleVariant(strLeft),COleVariant(strRight)),TRUE);
					interior =range.get_Interior();//��ɫ
					interior.put_Color(COleVariant((long)0x884400));
					font.AttachDispatch(range.get_Font());
					font.put_Color(COleVariant((long)0x00FF00)); 
					range.AttachDispatch(sheet.get_Cells()); 
				}
			}
		}
	}
	//���ñ��
	range.put_RowHeight(_variant_t(20));//��

	CRange cols;
	cols.AttachDispatch(range.get_Item(COleVariant((long)1),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //�����п�
	cols.AttachDispatch(range.get_Item(COleVariant((long)2),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //�����п�
	cols.AttachDispatch(range.get_Item(COleVariant((long)3),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //�����п�
	cols.AttachDispatch(range.get_Item(COleVariant((long)4),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //�����п�
	cols.AttachDispatch(range.get_Item(COleVariant((long)5),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //�����п�
	cols.AttachDispatch(range.get_Item(COleVariant((long)6),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //�����п�
	cols.AttachDispatch(range.get_Item(COleVariant((long)7),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //�����п�
	cols.AttachDispatch(range.get_Item(COleVariant((long)8),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //�����п�
	cols.AttachDispatch(range.get_Item(COleVariant((long)9),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //�����п�
	//cols.AttachDispatch(range.get_Item(COleVariant((long)10),vtMissing).pdispVal,TRUE);
	//cols.put_ColumnWidth(COleVariant((long)20)); //�����п�
	range.put_HorizontalAlignment(COleVariant((long)-4131)); 
	

	//����
	book.SaveCopyAs(COleVariant(strFile)); 
	book.put_Saved(true);
	app.put_Visible(true); 
 
	//�ͷŶ��� 
	range.ReleaseDispatch(); 
	sheet.ReleaseDispatch(); 
	sheets.ReleaseDispatch(); 
	book.ReleaseDispatch(); 
	books.ReleaseDispatch();
	app.ReleaseDispatch(); 
	app.Quit(); 
}

void CDefect_List_ShowDlg::OnNMClickListDefects(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<NMITEMACTIVATE*>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString strSteel;    // ѡ�����Ե������ַ��� 
	CString strSequeceNo;    // ѡ�����Ե������ַ���
	CString strAlerm;    // ѡ�����Ե������ַ���   
    NMLISTVIEW *pNMListView = (NMLISTVIEW*)pNMHDR;   
  
    if (-1 != pNMListView->iItem)        // ���iItem����-1����˵�����б��ѡ��   
    {   
        // ��ȡ��ѡ���б����һ��������ı�  
		strSequeceNo = m_ListCtrl.GetItemText(pNMListView->iItem, 1);
		int iSequeceNo=_wtoi(strSequeceNo);
        strSteel = m_ListCtrl.GetItemText(pNMListView->iItem, 2);   
		strAlerm = m_ListCtrl.GetItemText(pNMListView->iItem, 7);   
        // ��ѡ���������ʾ��༭����   
        SetDlgItemText(IDC_EDIT_STEEL, strSteel);  
		SetDlgItemText(IDC_EDIT_ALERM, strAlerm);  
		if(strAlerm.Compare(L"����")==0)
		{
			CButton* radio1=(CButton*)GetDlgItem(IDC_RADIO_YES);
			radio1->SetCheck(1);
			CButton* radio2=(CButton*)GetDlgItem(IDC_RADIO_NO);
			radio2->SetCheck(0);
		}
		else
		{
			CButton* radio1=(CButton*)GetDlgItem(IDC_RADIO_YES);
			radio1->SetCheck(0);
			CButton* radio2=(CButton*)GetDlgItem(IDC_RADIO_NO);
			radio2->SetCheck(1);
		}

		//��ȱ�����ݿ�
		m_DefectList.DeleteAllItems();
		int Index=0;
		for(int n=1;n<=6;n++)
		{
			CString strDBServer=L"";
			strDBServer.Format(L"%s",m_strDBServer);
			CString strDBName=L"";
			strDBName.Format(L"ClientDefectDB%d",n);
			CString strUser=L"ARNTUSER";
			CString strPassWd=L"ARNTUSER";

			::CoInitialize(NULL);
			CString strTableContent;
	 
			_ConnectionPtr m_pConnection;
			_variant_t RecordsAffected;
			_RecordsetPtr m_pRecordset;
	 
			try
			{
				m_pConnection.CreateInstance(__uuidof(Connection));
				CString sql;
				//sql.Format(L"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s",strSteelWarn);
				sql.Format(L"Provider=SQLOLEDB.1;Password=%s;Persist Security Info=True; User ID=%s;Initial Catalog=%s;Data Source=%s",strPassWd,strUser,strDBName,strDBServer);	
				_bstr_t strCmd=sql;
				m_pConnection->Open(strCmd,"","",adModeUnknown);
			}
			catch(_com_error e)
			{
				CString errormessage;
				errormessage.Format(CString("�������ݿ�ʧ��!\r������Ϣ:%s"),e.ErrorMessage());
				AfxMessageBox(errormessage);
				return;
			}
	 
			try
			{
				m_pRecordset.CreateInstance("ADODB.Recordset"); //ΪRecordset���󴴽�ʵ��
				CString sql;
				//m_DefectsMDBSet.strDefectsMDB[i].iDefectsNumMax=200;
				//sql.Format(L"select * from steel where TopDetectTime between \'%04d-%02d-%02d %02d:%02d:%02d.000\' and \'%04d-%02d-%02d %02d:%02d:%02d.000\' order by TopDetectTime desc",tmStart.wYear, tmStart.wMonth, tmStart.wDay, iStartHour, iStartMin,0,tmEnd.wYear, tmEnd.wMonth, tmEnd.wDay, iEndHour,iEndMin,0);
				sql.Format(L"select * from defect where SteelNo=%d",iSequeceNo);//

				_bstr_t strCmd=sql;				
				m_pRecordset=m_pConnection->Execute(strCmd,&RecordsAffected,adCmdText);
	 
			}
			catch(_com_error &e)
			{
				AfxMessageBox(e.Description());
			}
	 
			_variant_t vInfo;
			try
			{
				while(!m_pRecordset->adoEOF)
				{
					int iClass=m_pRecordset->GetCollect("Class");	
					int iFace=m_pRecordset->GetCollect("CameraNo");
					int iSetX=m_pRecordset->GetCollect("LeftInSteel");
					int iSetY=m_pRecordset->GetCollect("TopInSteel");
					int iArea=m_pRecordset->GetCollect("Area");
					//�����б�
					CString strIndex;
					strIndex.Format(L"%d",Index);
					m_DefectList.InsertItem(Index,strIndex);

					CString strClass;
					switch(iClass)
					{
					case 0:
						strClass=L"������";
						break;
					case 1:
						strClass=L"��ӡ";
						break;
					case 2:
						strClass=L"��Ƥ";
						break;
					case 3:
						strClass=L"����";
						break;
					case 4:
						strClass=L"����";
						break;
					case 5:
						strClass=L"����";
						break;
					case 6:
						strClass=L"����ֲ�";
						break;
					case 7:
						strClass=L"����ѹ��";
						break;
					case 8:
						strClass=L"������";
						break;
					case 9:
						strClass=L"���";
						break;
					default:
						break;
						
					}
					//strClass.Format(L"%d",iClass);
					m_DefectList.SetItemText(Index,1,strClass);

					CString strFace;
					if(iFace<=3)
					{
						strFace=L"�ϱ���";
					}
					else
					{
						strFace=L"�±���";
					}
					//strFace.Format(L"%d",iFace);
					m_DefectList.SetItemText(Index,2,strFace);

					CString strSetX;
					strSetX.Format(L"%d",iSetX);
					m_DefectList.SetItemText(Index,3,strSetX);

					CString strSetY;
					strSetY.Format(L"%d",iSetY);
					m_DefectList.SetItemText(Index,4,strSetY);

					CString strArea;
					strArea.Format(L"%d",iArea);
					m_DefectList.SetItemText(Index,5,strArea);

					Index++;
					if(Index>1024)
					{
					break;
					}
					m_pRecordset->MoveNext();
				}
			}
			catch(_com_error &e)
			{
				AfxMessageBox(e.Description());
			}
		}  
	}
	*pResult = 0;
	
}

void CDefect_List_ShowDlg::OnClose()
{
	// TODO: �ڴ������Ϣ�����������/�����Ĭ��ֵ
	//PostMessage(WM_COMMAND,MAKEWPARAM(IDC_BUTTON_SET,BN_CLICKED));
	CDialog::OnClose();
}

void CDefect_List_ShowDlg::OnBnClickedRadioYes()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	SetDlgItemText(IDC_EDIT_ALERM, L"����");   
}

void CDefect_List_ShowDlg::OnBnClickedRadioNo()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	SetDlgItemText(IDC_EDIT_ALERM, L"*");   
}

void CDefect_List_ShowDlg::OnBnClickedButtonSet()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString strSteel;
	GetDlgItemText(IDC_EDIT_STEEL, strSteel); 
	if(strSteel.Compare(L"")==0)
	{
		return;
	}

	m_ListCtrl.DeleteAllItems();

	CString strDBServer=L"";
	strDBServer.Format(L"%s",m_strDBServer);
	CString strDBName=L"SteelRecord";
	CString strUser=L"ARNTUSER";
	CString strPassWd=L"ARNTUSER";

	::CoInitialize(NULL);
	CString strTableContent;
 
	_ConnectionPtr m_pConnection;
	_variant_t RecordsAffected;
	_RecordsetPtr m_pRecordset;
 
	try
	{
		m_pConnection.CreateInstance(__uuidof(Connection));
		CString sql;
		//sql.Format(L"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s",strSteelWarn);
		sql.Format(L"Provider=SQLOLEDB.1;Password=%s;Persist Security Info=True; User ID=%s;Initial Catalog=%s;Data Source=%s",strPassWd,strUser,strDBName,strDBServer);	
		_bstr_t strCmd=sql;
		m_pConnection->Open(strCmd,"","",adModeUnknown);
	}
	catch(_com_error e)
	{
		CString errormessage;
		errormessage.Format(CString("�������ݿ�ʧ��!\r������Ϣ:%s"),e.ErrorMessage());
		AfxMessageBox(errormessage);
		return;
	}
 
	try
	{
		m_pRecordset.CreateInstance("ADODB.Recordset"); //ΪRecordset���󴴽�ʵ��
		CString sql;
		//m_DefectsMDBSet.strDefectsMDB[i].iDefectsNumMax=200;
		//sql.Format(L"select * from steel where TopDetectTime between \'%04d-%02d-%02d %02d:%02d:%02d.000\' and \'%04d-%02d-%02d %02d:%02d:%02d.000\' order by TopDetectTime desc",tmStart.wYear, tmStart.wMonth, tmStart.wDay, iStartHour, iStartMin,0,tmEnd.wYear, tmEnd.wMonth, tmEnd.wDay, iEndHour,iEndMin,0);
		sql.Format(L"select steel.SequeceNo,steel.SteelID,steel.TopDetectTime,SteelID.Thick,SteelID.Width,SteelID.Length,SteelID.AddTime,SteelGrade.Dsc\
					from SteelID inner join steel on SteelID.ID=steel.SteelID inner join SteelGrade on SteelGrade.SequeceNo=steel.SequeceNo \
					where SteelID LIKE \'%%%s%%\'",strSteel);//

		_bstr_t strCmd=sql;				
		
		m_pRecordset=m_pConnection->Execute(strCmd,&RecordsAffected,adCmdText);
		int a=0;
 
	}
	catch(_com_error &e)
	{
		AfxMessageBox(e.Description());
	}
 
	_variant_t vInfo;
	int Index=0;
	try
	{
		while(!m_pRecordset->adoEOF)
		{
            int iSequeceNo=m_pRecordset->GetCollect("SequeceNo");		
			CString  strSteelID=m_pRecordset->GetCollect("SteelID");
			CString  strDateTime=m_pRecordset->GetCollect("TopDetectTime");
			int iThick=m_pRecordset->GetCollect("Thick");
			int iWidth=m_pRecordset->GetCollect("Width");
			int iLength=m_pRecordset->GetCollect("Length");
			CString  strAddTime=m_pRecordset->GetCollect("AddTime");
			vInfo=m_pRecordset->GetCollect("Dsc");
			CString  strDsc=L"";
			if(vInfo.vt!=VT_NULL)
			{
				strDsc.Format(L"%s",vInfo.bstrVal);
			}

			if(strDsc.IsEmpty() != true)
			{
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(255,0,0));  //���õ�Ԫ��������ɫ
			}
			else
			{
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(0,255,0));  //���õ�Ԫ��������ɫ
			}
			//�����б�
			CString strIndex;
			strIndex.Format(L"%d",Index);
			m_ListCtrl.InsertItem(Index,strIndex);

			CString strSequeceNo;
			strSequeceNo.Format(L"%d",iSequeceNo);
			m_ListCtrl.SetItemText(Index,1,strSequeceNo);

			m_ListCtrl.SetItemText(Index,2,strSteelID);

			CString strThick;
			strThick.Format(L"%d",iThick);
			m_ListCtrl.SetItemText(Index,3,strThick);

			CString strWidth;
			strWidth.Format(L"%d",iWidth);
			m_ListCtrl.SetItemText(Index,4,strWidth);

			CString strLength;
			strLength.Format(L"%d",iLength);
			m_ListCtrl.SetItemText(Index,5,strLength);

			m_ListCtrl.SetItemText(Index,6,strDateTime);

			m_ListCtrl.SetItemText(Index,8,strDsc);


			Index++;
			if(Index>1024)
			{
				break;
			}
			m_pRecordset->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		AfxMessageBox(e.Description());
	}
}
