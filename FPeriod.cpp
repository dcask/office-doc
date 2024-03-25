// FPeriod.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FPeriod.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FPeriod dialog


FPeriod::FPeriod(CWnd* pParent /*=NULL*/)
	: CDialog(FPeriod::IDD, pParent)
{
	//{{AFX_DATA_INIT(FPeriod)
	m_Begin = COleDateTime::GetCurrentTime();
	m_End = COleDateTime::GetCurrentTime();
	//}}AFX_DATA_INIT
}


void FPeriod::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FPeriod)
	DDX_DateTimeCtrl(pDX, IDC_BEGIN, m_Begin);
	DDX_DateTimeCtrl(pDX, IDC_END, m_End);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FPeriod, CDialog)
	//{{AFX_MSG_MAP(FPeriod)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FPeriod message handlers
