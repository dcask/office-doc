// FDeclarant.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FDeclarant.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FDeclarant dialog


FDeclarant::FDeclarant(CWnd* pParent /*=NULL*/)
	: CDialog(FDeclarant::IDD, pParent)
{
	//{{AFX_DATA_INIT(FDeclarant)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	m_csQuery="[GetDeclarants]";
	m_iDeclType=-1;
	m_csWindowName="Заявители";
}


void FDeclarant::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FDeclarant)
	DDX_Control(pDX, IDC_LIST, m_List);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FDeclarant, CDialog)
	//{{AFX_MSG_MAP(FDeclarant)
	ON_WM_SHOWWINDOW()
	ON_LBN_DBLCLK(IDC_LIST, OnDblclkList)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FDeclarant message handlers

void FDeclarant::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	
	if(!bShow) return;

	this->SetWindowText(m_csWindowName);
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CListBox* lb=(CListBox*) GetDlgItem(IDC_LIST);
	CADOCommand pCmd(bt->GetDB(), m_csQuery);
	CString csVal;
	//csVal="%";
	csVal=m_csString;
	//csVal+="%";
	pCmd.AddParameter("Decl",CADORecordset::typeChar,
				CADOParameter::paramInput,200, csVal.Left(199));
	if(m_iDeclType>-1) csVal.Format("%d", m_iDeclType);
	else csVal="%";
	
	pCmd.AddParameter("DeclType",CADORecordset::typeChar,
					CADOParameter::paramInput,50, csVal);
	try
	{
		bt->GetRS()->Execute(&pCmd);
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		return;
	}
	if(bt->GetRecordsCount()>0) bt->GetRS()->MoveFirst();
	for(UINT i=0; i< bt->GetRecordsCount(); ++i)
	{
		bt->GetRS()->GetFieldValue(0,csVal);
		lb->AddString(csVal);
		bt->GetRS()->MoveNext();
	}
	bt->GetRS()->Close();
	GetDlgItem(IDC_LIST)->SetFocus();
}

void FDeclarant::OnOK() 
{
	CListBox* lb=(CListBox*) GetDlgItem(IDC_LIST);
	if(lb->GetCurSel()>=0) lb->GetText(lb->GetCurSel(),m_csString);
	
	CDialog::OnOK();
}

void FDeclarant::OnDblclkList() 
{
	OnOK();
}
