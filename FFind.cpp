// FFind.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FFind.h"
#include "FDeclarant.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FFind dialog


FFind::FFind(CWnd* pParent /*=NULL*/)
	: CDialog(FFind::IDD, pParent)
{
	//{{AFX_DATA_INIT(FFind)
	m_uiYear = 0;
	m_csRegNum = _T("");
	m_csExtInNum = _T("");
	m_csDeclarant = _T("");
	//}}AFX_DATA_INIT
	m_uiYear = atoi((COleDateTime::GetCurrentTime()).Format("%Y"));
}


void FFind::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FFind)
	DDX_Control(pDX, IDC_SPIN, m_Spin);
	DDX_Text(pDX, IDC_YEAR, m_uiYear);
	DDV_MinMaxUInt(pDX, m_uiYear, 1950, 2200);
	DDX_Text(pDX, IDC_REG_NUM, m_csRegNum);
	DDX_Text(pDX, IDC_X_INNUM, m_csExtInNum);
	DDX_Text(pDX, IDC_DECLARANT, m_csDeclarant);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FFind, CDialog)
	//{{AFX_MSG_MAP(FFind)
	ON_WM_SHOWWINDOW()
	ON_BN_CLICKED(IDC_DECL_FIND, OnDeclFind)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FFind message handlers

void FFind::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	m_Spin.SetBase(0);
	m_Spin.SetRange(1950,2200);
	m_Spin.SetPos(m_uiYear);
	GetDlgItem(IDC_REG_NUM)->SetFocus();
	
	
}


void FFind::OnDeclFind() 
{
	FDeclarant dlg;
	GetDlgItemText(IDC_DECLARANT, dlg.m_csString);
	CComboBox* cb=(CComboBox*) GetDlgItem(IDC_DECLARANT);
	if(dlg.DoModal()==IDOK)
		SetDlgItemText(IDC_DECLARANT, dlg.m_csString);
}
