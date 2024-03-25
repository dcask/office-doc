// FOutFilter.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FOutFilter.h"
#include "Kan.h"
#include "FDeclarant.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FOutFilter dialog


FOutFilter::FOutFilter(CWnd* pParent /*=NULL*/)
	: CDialog(FOutFilter::IDD, pParent)
{
	//{{AFX_DATA_INIT(FOutFilter)
	m_csAuthor = _T("");
	m_csContent = _T("");
	m_csReciever = _T("");
	//}}AFX_DATA_INIT
	m_ToolTip=NULL;
}


void FOutFilter::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FOutFilter)
	DDX_CBString(pDX, IDC_AUTHOR, m_csAuthor);
	DDX_Text(pDX, IDC_CONTENT, m_csContent);
	DDX_Text(pDX, IDC_RECIEVER, m_csReciever);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FOutFilter, CDialog)
	//{{AFX_MSG_MAP(FOutFilter)
	ON_WM_SHOWWINDOW()
	ON_BN_CLICKED(IDC_PROMT, OnPromt)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FOutFilter message handlers

void FOutFilter::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	
	if(!bShow) return;
	
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	bt->LoadComboBox(GetDlgItem(IDC_AUTHOR),"[ComboPersons]");
	GetDlgItem(IDC_AUTHOR)->SetFocus();
	m_ToolTip =new CToolTipCtrl();
	if (!m_ToolTip->Create(this))
	{
		TRACE("Unable To create ToolTip\n");           
		return;
	}

	m_ToolTip->AddTool(GetDlgItem(IDC_PROMT),"Выбрать заявителя");
	m_ToolTip->Activate(TRUE);
}	

BOOL FOutFilter::DestroyWindow() 
{
	if(m_ToolTip) 	m_ToolTip->DestroyWindow();
	delete m_ToolTip;
	
	return CDialog::DestroyWindow();
}

BOOL FOutFilter::PreTranslateMessage(MSG* pMsg) 
{
	if (NULL != m_ToolTip)            
      m_ToolTip->RelayEvent(pMsg);
	
	return CDialog::PreTranslateMessage(pMsg);
}

void FOutFilter::OnPromt() 
{
	FDeclarant dlg;
	GetDlgItemText(IDC_RECIEVER, dlg.m_csString);
	CComboBox* cb=(CComboBox*) GetDlgItem(IDC_RECIEVER);
	dlg.m_csQuery="[GetPlaces]";
	dlg.m_csWindowName="Адресат";
	if(dlg.DoModal()==IDOK)
	{
		SetDlgItemText(IDC_RECIEVER, dlg.m_csString);
	}
	
}
