// FResolution.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FResolution.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FResolution dialog


FResolution::FResolution(CWnd* pParent /*=NULL*/)
	: CDialog(FResolution::IDD, pParent)
{
	//{{AFX_DATA_INIT(FResolution)
	m_csResolution = _T("");
	m_OutDate = COleDateTime::GetCurrentTime();
	//}}AFX_DATA_INIT
}


void FResolution::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FResolution)
	DDX_Control(pDX, IDC_AUTHOR, m_Author);
	DDX_Text(pDX, IDC_RESOLUTION, m_csResolution);
	DDX_Control(pDX, IDC_MOUT_DATE, m_mOutDate);
	DDX_DateTimeCtrl(pDX, IDC_OUT_DATE, m_OutDate);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FResolution, CDialog)
	//{{AFX_MSG_MAP(FResolution)
	ON_WM_SHOWWINDOW()
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_OUT_DATE, OnDatetimechangeOutDate)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FResolution message handlers

void FResolution::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	if(!bShow) return;
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	bt->LoadComboBox(GetDlgItem(IDC_AUTHOR),"[ComboPersons]");
	m_OutDate=doc->m_ResolutionDate;
	if( m_OutDate.m_status !=COleDateTime::null) m_mOutDate.SetText(m_OutDate.Format("%x"));
	m_csResolution=doc->m_csResolution;
	UpdateData(FALSE);
	m_Author.SetCurSel(m_Author.FindString(0,doc->m_csAuthor));
}

BOOL FResolution::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	CString csFormatDate=((CKanApp*) AfxGetApp())->m_csFormatDate;
	CString csVal;
		
	for(int i=0; i<csFormatDate.GetLength(); ++i) 
		if(csFormatDate.GetAt(i)<='A') csVal+=csFormatDate.GetAt(i); 
			else csVal+="9";


	m_mOutDate.SetMask(csVal);
	m_mOutDate.SetFormat(csFormatDate);
	
	

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void FResolution::OnDatetimechangeOutDate(NMHDR* pNMHDR, LRESULT* pResult) 
{
	if( m_OutDate.Format("%x") !="") m_mOutDate.SetText(m_OutDate.Format("%x"));
	UpdateData();
	*pResult = 0;
}


BOOL FResolution::DestroyWindow() 
{
	UpdateData(TRUE);

	m_mOutDate.SetPromptInclude(FALSE);

	if(m_mOutDate.GetText()!="")
		doc->m_ResolutionDate.ParseDateTime(m_mOutDate.GetFormattedText());
	else doc->m_ResolutionDate.m_status=COleDateTime::null;

	
	m_mOutDate.SetPromptInclude(TRUE);

	doc->m_csResolution=m_csResolution;
	if(m_Author.GetCurSel()>=0) 
	{
		m_Author.GetLBText(m_Author.GetCurSel(),doc->m_csAuthor);
		doc->m_iAuthor=m_Author.GetItemData(m_Author.GetCurSel());
		
	}
	else 
	{
		doc->m_iAuthor=-1;
		doc->m_csAuthor=_T("");
	}
	
	return CDialog::DestroyWindow();
}

BOOL FResolution::PreTranslateMessage(MSG* pMsg) 
{
	if(WM_KEYDOWN == pMsg->message)
	{
		switch(pMsg->wParam)
		{
			case VK_RETURN:
			case VK_ESCAPE:
				return TRUE;
				break;
			default: break;
		}
	}
	
	return CDialog::PreTranslateMessage(pMsg);
}
