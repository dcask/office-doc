// FTheme.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FTheme.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FTheme dialog


FTheme::FTheme(CWnd* pParent /*=NULL*/)
	: CDialog(FTheme::IDD, pParent)
{
	//{{AFX_DATA_INIT(FTheme)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void FTheme::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FTheme)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FTheme, CDialog)
	//{{AFX_MSG_MAP(FTheme)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FTheme message handlers
