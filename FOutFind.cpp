// FOutFind.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FOutFind.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FOutFind dialog


FOutFind::FOutFind(CWnd* pParent /*=NULL*/)
	: CDialog(FOutFind::IDD, pParent)
{
	//{{AFX_DATA_INIT(FOutFind)
	m_csRegNum = _T("");
	m_uiYear = atoi((COleDateTime::GetCurrentTime()).Format("%Y"));
	//}}AFX_DATA_INIT
}


void FOutFind::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FOutFind)
	DDX_Control(pDX, IDC_SPIN, m_Spin);
	DDX_Text(pDX, IDC_REG_NUM, m_csRegNum);
	DDX_Text(pDX, IDC_YEAR, m_uiYear);
	DDV_MinMaxUInt(pDX, m_uiYear, 1900, 2200);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FOutFind, CDialog)
	//{{AFX_MSG_MAP(FOutFind)
	ON_WM_SHOWWINDOW()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FOutFind message handlers

void FOutFind::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	
	m_Spin.SetBase(0);
	m_Spin.SetRange(1950,2200);
	m_Spin.SetPos(m_uiYear);
	GetDlgItem(IDC_REG_NUM)->SetFocus();
	
}
