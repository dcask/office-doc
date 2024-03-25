// FFilter.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FFilter.h"
#include "FDeclarant.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FFilter dialog


FFilter::FFilter(CWnd* pParent /*=NULL*/)
	: CDialog(FFilter::IDD, pParent)
{
	//{{AFX_DATA_INIT(FFilter)
	m_csDeclarant = _T("");
	m_csContent = _T("");
	//}}AFX_DATA_INIT
	m_ToolTip=NULL;
}


void FFilter::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FFilter)
	DDX_Text(pDX, IDC_DECLARANT, m_csDeclarant);
	DDX_Text(pDX, IDC_CONTENT, m_csContent);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FFilter, CDialog)
	//{{AFX_MSG_MAP(FFilter)
	ON_BN_CLICKED(IDC_PROMT, OnPromt)
	ON_WM_SHOWWINDOW()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FFilter message handlers

void FFilter::OnPromt() 
{
	FDeclarant dlg;
	GetDlgItemText(IDC_DECLARANT, dlg.m_csString);
	dlg.m_iDeclType=-1;
	if(dlg.DoModal()==IDOK)
	{
		SetDlgItemText(IDC_DECLARANT, dlg.m_csString);
	}
	
}

void FFilter::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	
	GetDlgItem(IDC_DECLARANT)->SetFocus();
	m_ToolTip =new CToolTipCtrl();
	if (!m_ToolTip->Create(this))
	{
		TRACE("Unable To create ToolTip\n");           
		return;
	}

	m_ToolTip->AddTool(GetDlgItem(IDC_PROMT),"Выбрать заявителя");
	m_ToolTip->Activate(TRUE);
}

BOOL FFilter::PreTranslateMessage(MSG* pMsg) 
{
	CWnd* focus=GetFocus();
	if (NULL != m_ToolTip)            
      m_ToolTip->RelayEvent(pMsg);
	if(WM_KEYDOWN == pMsg->message)
	{
		int id=focus->GetDlgCtrlID();
		switch(pMsg->wParam)
		{
			case VK_RETURN:
					if(id!=IDC_DECLARANT) 
						EndDialog(IDOK);
					else 
						OnPromt();
				break;
			case VK_ESCAPE:
				EndDialog(IDCANCEL);
				break;
			default:
				return CDialog::PreTranslateMessage(pMsg);
		}
		return TRUE;       // запрет дальнейшей обработки
	}
	
	return CDialog::PreTranslateMessage(pMsg);
}

BOOL FFilter::DestroyWindow() 
{
	if(m_ToolTip) m_ToolTip->DestroyWindow();
	delete m_ToolTip;
	
	return CDialog::DestroyWindow();
}
