// FChoosDate.cpp : implementation file
//

#include "stdafx.h"
#include "kan.h"
#include "FChoosDate.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FChoosDate dialog


FChoosDate::FChoosDate(CWnd* pParent /*=NULL*/)
	: CDialog(FChoosDate::IDD, pParent)
{
	//{{AFX_DATA_INIT(FChoosDate)
	m_Date = COleDateTime::GetCurrentTime();
	//}}AFX_DATA_INIT
}


void FChoosDate::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FChoosDate)
	DDX_DateTimeCtrl(pDX, IDC_DATE, m_Date);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FChoosDate, CDialog)
	//{{AFX_MSG_MAP(FChoosDate)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FChoosDate message handlers

void FChoosDate::OnOK() 
{

	CDialog::OnOK();
}
