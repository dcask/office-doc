// FWait.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FWait.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FWait dialog


FWait::FWait(CWnd* pParent /*=NULL*/)
	: CDialog(FWait::IDD, pParent)
{
	//{{AFX_DATA_INIT(FWait)
	m_csInfo = _T("");
	//}}AFX_DATA_INIT
}


void FWait::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FWait)
	DDX_Text(pDX, IDC_INFO, m_csInfo);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FWait, CDialog)
	//{{AFX_MSG_MAP(FWait)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FWait message handlers

void FWait::SetInfo(CString csVal)
{
	this->
	SetDlgItemText(IDC_INFO,csVal);
}
