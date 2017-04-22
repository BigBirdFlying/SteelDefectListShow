// Defect_List_ShowDlg.cpp : 实现文件
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


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CDefect_List_ShowDlg 对话框




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


// CDefect_List_ShowDlg 消息处理程序

BOOL CDefect_List_ShowDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	ReadConfigFile();
	//得到显示器大小
	int   cx,cy;
	cx   =   GetSystemMetrics(SM_CXSCREEN);
	cy   =   GetSystemMetrics(SM_CYSCREEN)-50;
	//再用MoveWindow
	CRect   rcTemp;
	rcTemp.BottomRight()   =   CPoint(cx,   cy);
	rcTemp.TopLeft()   =   CPoint(0, 0);
	MoveWindow(&rcTemp);     

	CRect   rcTemp0;
	rcTemp0.TopLeft()=CPoint(0, cy/4);
	rcTemp0.BottomRight()=CPoint(cx,  cy);
	m_ListCtrl.MoveWindow(&rcTemp0);  


	//查看模块
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
	//第三行
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
	m_ListCtrl.SetHeaderHeight(1.5);          //设置头部高度
	m_ListCtrl.SetHeaderFontHW(16,0);         //设置头部字体高度,和宽度,0表示缺省，自适应 
	m_ListCtrl.SetHeaderTextColor(RGB(255,200,100)); //设置头部字体颜色
	m_ListCtrl.SetHeaderBKColor(128,255,255,8); //设置头部背景色

	m_ListCtrl.SetBkColor(RGB(128,128,128));        //设置背景色RGB(128,64,0)

	m_ListCtrl.SetRowHeigt(25);               //设置行高度
	m_ListCtrl.SetHeaderHeight(1.5);          //设置头部高度
	m_ListCtrl.SetFontHW(15,0);               //设置字体高度，和宽度,0表示缺省宽度

	int n=9;
	float m=0.05;
	m_ListCtrl.InsertColumn(0,_T("序号"),LVCFMT_CENTER,rcTemp0.Width()*m);
	m_ListCtrl.InsertColumn(1,_T("流水号"),LVCFMT_CENTER,rcTemp0.Width()*m);
	m_ListCtrl.InsertColumn(2,_T("钢板号"),LVCFMT_CENTER,rcTemp0.Width()*2*m);
	m_ListCtrl.InsertColumn(3,_T("厚度"),LVCFMT_CENTER,rcTemp0.Width()*m);
	m_ListCtrl.InsertColumn(4,_T("宽度"),LVCFMT_CENTER,rcTemp0.Width()*m);
	m_ListCtrl.InsertColumn(5,_T("长度"),LVCFMT_CENTER,rcTemp0.Width()*m);
	//m_ListCtrl.InsertColumn(6,_T("生产时间"),LVCFMT_CENTER,rcTemp0.Width()*2.5*m);
	m_ListCtrl.InsertColumn(6,_T("检测时间"),LVCFMT_CENTER,rcTemp0.Width()*2.5*m);
	m_ListCtrl.InsertColumn(7,_T("缺陷自动分析"),LVCFMT_CENTER,rcTemp0.Width()*4.5*m);
	m_ListCtrl.InsertColumn(8,_T("人工缺陷描述"),LVCFMT_CENTER,rcTemp0.Width()*6*m);

	CRect rect_temp2;
	GetDlgItem(IDC_LIST_DEFECTLIST)->GetWindowRect(&rect_temp2);
	m_DefectList.InsertColumn(0,_T("序号"),LVCFMT_CENTER,rect_temp2.Width()*0.1);
	m_DefectList.InsertColumn(1,_T("类别"),LVCFMT_CENTER,rect_temp2.Width()*0.2);
	m_DefectList.InsertColumn(2,_T("表面"),LVCFMT_CENTER,rect_temp2.Width()*0.1);
	m_DefectList.InsertColumn(3,_T("横向位置(mm)"),LVCFMT_CENTER,rect_temp2.Width()*0.2);
	m_DefectList.InsertColumn(4,_T("纵向位置(mm)"),LVCFMT_CENTER,rect_temp2.Width()*0.2);
	m_DefectList.InsertColumn(5,_T("面积(mm2)"),LVCFMT_CENTER,rect_temp2.Width()*0.2);

	//初始化
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
	m_ListCtrl.SetExtendedStyle(LVS_EX_GRIDLINES);                     //设置扩展风格为网格
	::SendMessage(m_ListCtrl.m_hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE,LVS_EX_FULLROWSELECT, LVS_EX_FULLROWSELECT);
	//PostMessage(WM_COMMAND,MAKEWPARAM(IDC_BUTTON_SET,BN_CLICKED));
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CDefect_List_ShowDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
				//----------给窗口设置背景图片----------------------------
		CPaintDC dc(this); 
		CRect rect;
		GetClientRect(&rect);
		CDC   dcMem;   
		dcMem.CreateCompatibleDC(&dc);   
		CBitmap   bmpBackground;   
		bmpBackground.LoadBitmap(IDC_BACKGROUND);   //IDB_BITMAP是你自己的图对应的ID 

		BITMAP   bitmap;   
		bmpBackground.GetBitmap(&bitmap);   
		CBitmap   *pbmpOld=dcMem.SelectObject(&bmpBackground);   
		dc.StretchBlt(0,0,rect.Width(),rect.Height(),&dcMem,0,0,bitmap.bmWidth,bitmap.bmHeight,SRCCOPY);  

		//这句话放在下边才能使图像显示出来
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
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
			AfxMessageBox(L"IP地址为空");
		}
	}
	else
	{
		AfxMessageBox(L"未指定IP");
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
		errormessage.Format(CString("连接数据库失败!\r错误信息:%s"),e.ErrorMessage());
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
		m_pRecordset.CreateInstance("ADODB.Recordset"); //为Recordset对象创建实例
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
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(255,0,0));  //设置单元格字体颜色
			}
			else
			{
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(0,255,0));  //设置单元格字体颜色
			}
			//填入列表
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
	// TODO: 在此添加控件通知处理程序代码
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
		errormessage.Format(CString("连接数据库失败!\r错误信息:%s"),e.ErrorMessage());
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
		m_pRecordset.CreateInstance("ADODB.Recordset"); //为Recordset对象创建实例
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
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(255,0,0));  //设置单元格字体颜色
			}
			else
			{
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(0,255,0));  //设置单元格字体颜色
			}
			//填入列表
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
	//	CCommonFunc::ErrorBox(L"无法加载参数文件:%s", strConfigFile);
	//	return ArNT_FALSE;
	//}

	//LONGSTR strValue = {0};	
	//LONGSTR strName = {0};

	////数据库部分
	////strSteelWarn={0};
	//CCommonFunc::SafeStringPrintf(strName, STR_LEN(strName), L"钢板缺陷信息展示配置#缺陷报警信息数据库设置#数据库路径");
	//if(m_XMLOperator.GetValueByString(strName, strValue))
	//{
	//	CCommonFunc::SafeStringCpy(strSteelWarn, STR_LEN(strSteelWarn), strValue);
	//}

	return true;
}
void CDefect_List_ShowDlg::OnBnClickedButtonReport()
{
	// TODO: 在此添加控件通知处理程序代码
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
	strFile.Format(L"D:\\%04d年%02d月%02d日%02d时%02d分__%04d年%02d月%02d日%02d时%02d分.xlsx",tmStart.wYear,tmStart.wMonth,tmStart.wDay,iStartHour,iStartMin,tmEnd.wYear,tmEnd.wMonth,tmEnd.wDay,iEndHour,iEndMin);

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
	Cnterior interior;//背景色
 
	if (!app.CreateDispatch(_T("Excel.Application")))
	{
		MessageBox(_T("Error!Creat Excel Application Server Faile!"));
		//exit(1);
	}
	books = app.get_Workbooks();
	book = books.Add(covOptional);
	sheets = book.get_Worksheets();
	sheet = sheets.get_Item(COleVariant((short)1));
	//得到全部Cells 
	range.AttachDispatch(sheet.get_Cells()); 
	CString sText[]={_T("序号"),_T("流水号"),_T("钢板号"),_T("厚度"),_T("宽度"),_T("长度"),_T("检测时间"),_T("缺陷自动分析"),_T("人工缺陷描述")};
	for (int setnum=0;setnum<m_ListCtrl.GetItemCount()+1;setnum++)
	{
		for (int num=0;num<9;num++)
		{
			if (!setnum)
			{
				range.put_Item(_variant_t((long)(setnum+1)),_variant_t((long)(num+1)),_variant_t(sText[num]));
				range.AttachDispatch(sheet.get_Range(COleVariant(L"A1"),COleVariant(L"I1")),TRUE);
				interior =range.get_Interior();//底色
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
				if(strTemp.Compare(L"存在")==0)
				{
					CString strLeft,strRight;
					strLeft.Format(L"A%d",setnum+1);
					strRight.Format(L"I%d",setnum+1);
					range.AttachDispatch(sheet.get_Range(COleVariant(strLeft),COleVariant(strRight)),TRUE);
					interior =range.get_Interior();//底色
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
					interior =range.get_Interior();//底色
					interior.put_Color(COleVariant((long)0x884400));
					font.AttachDispatch(range.get_Font());
					font.put_Color(COleVariant((long)0x00FF00)); 
					range.AttachDispatch(sheet.get_Cells()); 
				}
			}
		}
	}
	//设置表格
	range.put_RowHeight(_variant_t(20));//磅

	CRange cols;
	cols.AttachDispatch(range.get_Item(COleVariant((long)1),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //设置列宽
	cols.AttachDispatch(range.get_Item(COleVariant((long)2),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //设置列宽
	cols.AttachDispatch(range.get_Item(COleVariant((long)3),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //设置列宽
	cols.AttachDispatch(range.get_Item(COleVariant((long)4),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //设置列宽
	cols.AttachDispatch(range.get_Item(COleVariant((long)5),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //设置列宽
	cols.AttachDispatch(range.get_Item(COleVariant((long)6),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //设置列宽
	cols.AttachDispatch(range.get_Item(COleVariant((long)7),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //设置列宽
	cols.AttachDispatch(range.get_Item(COleVariant((long)8),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //设置列宽
	cols.AttachDispatch(range.get_Item(COleVariant((long)9),vtMissing).pdispVal,TRUE);
	cols.put_ColumnWidth(COleVariant((long)20)); //设置列宽
	//cols.AttachDispatch(range.get_Item(COleVariant((long)10),vtMissing).pdispVal,TRUE);
	//cols.put_ColumnWidth(COleVariant((long)20)); //设置列宽
	range.put_HorizontalAlignment(COleVariant((long)-4131)); 
	

	//保存
	book.SaveCopyAs(COleVariant(strFile)); 
	book.put_Saved(true);
	app.put_Visible(true); 
 
	//释放对象 
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
	// TODO: 在此添加控件通知处理程序代码
	CString strSteel;    // 选择语言的名称字符串 
	CString strSequeceNo;    // 选择语言的名称字符串
	CString strAlerm;    // 选择语言的名称字符串   
    NMLISTVIEW *pNMListView = (NMLISTVIEW*)pNMHDR;   
  
    if (-1 != pNMListView->iItem)        // 如果iItem不是-1，就说明有列表项被选择   
    {   
        // 获取被选择列表项第一个子项的文本  
		strSequeceNo = m_ListCtrl.GetItemText(pNMListView->iItem, 1);
		int iSequeceNo=_wtoi(strSequeceNo);
        strSteel = m_ListCtrl.GetItemText(pNMListView->iItem, 2);   
		strAlerm = m_ListCtrl.GetItemText(pNMListView->iItem, 7);   
        // 将选择的语言显示与编辑框中   
        SetDlgItemText(IDC_EDIT_STEEL, strSteel);  
		SetDlgItemText(IDC_EDIT_ALERM, strAlerm);  
		if(strAlerm.Compare(L"存在")==0)
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

		//读缺陷数据库
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
				errormessage.Format(CString("连接数据库失败!\r错误信息:%s"),e.ErrorMessage());
				AfxMessageBox(errormessage);
				return;
			}
	 
			try
			{
				m_pRecordset.CreateInstance("ADODB.Recordset"); //为Recordset对象创建实例
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
					//填入列表
					CString strIndex;
					strIndex.Format(L"%d",Index);
					m_DefectList.InsertItem(Index,strIndex);

					CString strClass;
					switch(iClass)
					{
					case 0:
						strClass=L"待分类";
						break;
					case 1:
						strClass=L"辊印";
						break;
					case 2:
						strClass=L"翘皮";
						break;
					case 3:
						strClass=L"划伤";
						break;
					case 4:
						strClass=L"裂纹";
						break;
					case 5:
						strClass=L"凹坑";
						break;
					case 6:
						strClass=L"表面粗糙";
						break;
					case 7:
						strClass=L"异物压入";
						break;
					case 8:
						strClass=L"暗划伤";
						break;
					case 9:
						strClass=L"麻点";
						break;
					default:
						break;
						
					}
					//strClass.Format(L"%d",iClass);
					m_DefectList.SetItemText(Index,1,strClass);

					CString strFace;
					if(iFace<=3)
					{
						strFace=L"上表面";
					}
					else
					{
						strFace=L"下表面";
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
	// TODO: 在此添加消息处理程序代码和/或调用默认值
	//PostMessage(WM_COMMAND,MAKEWPARAM(IDC_BUTTON_SET,BN_CLICKED));
	CDialog::OnClose();
}

void CDefect_List_ShowDlg::OnBnClickedRadioYes()
{
	// TODO: 在此添加控件通知处理程序代码
	SetDlgItemText(IDC_EDIT_ALERM, L"存在");   
}

void CDefect_List_ShowDlg::OnBnClickedRadioNo()
{
	// TODO: 在此添加控件通知处理程序代码
	SetDlgItemText(IDC_EDIT_ALERM, L"*");   
}

void CDefect_List_ShowDlg::OnBnClickedButtonSet()
{
	// TODO: 在此添加控件通知处理程序代码
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
		errormessage.Format(CString("连接数据库失败!\r错误信息:%s"),e.ErrorMessage());
		AfxMessageBox(errormessage);
		return;
	}
 
	try
	{
		m_pRecordset.CreateInstance("ADODB.Recordset"); //为Recordset对象创建实例
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
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(255,0,0));  //设置单元格字体颜色
			}
			else
			{
				m_ListCtrl.SetRowItemColor(Index,(DWORD)RGB(0,255,0));  //设置单元格字体颜色
			}
			//填入列表
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
