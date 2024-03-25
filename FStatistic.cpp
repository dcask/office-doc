// FStatistic.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FStatistic.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FStatistic dialog


FStatistic::FStatistic(CWnd* pParent /*=NULL*/)
	: CDialog(FStatistic::IDD, pParent)
{
	//{{AFX_DATA_INIT(FStatistic)
	m_csInfo = _T("");
	m_Begin=COleDateTime::GetCurrentTime();
	m_End=COleDateTime::GetCurrentTime();
	//}}AFX_DATA_INIT
}


void FStatistic::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FStatistic)
	DDX_Control(pDX, IDC_DATA, m_Data);
	DDX_Text(pDX, IDC_INFO, m_csInfo);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FStatistic, CDialog)
	//{{AFX_MSG_MAP(FStatistic)
	ON_WM_SHOWWINDOW()
	ON_BN_CLICKED(IDC_UNLOAD, OnMUnload)
	ON_WM_CLOSE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FStatistic message handlers

void FStatistic::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	if(!bShow) return;
	CDialog::OnShowWindow(bShow, nStatus);
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOFieldInfo info;
	CADOCommand pCmd(bt->GetDB(), m_csQuery);
	CString csVal;
	long row=0;
	pCmd.AddParameter("BDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_Begin.Format(VAR_DATEVALUEONLY));
	pCmd.AddParameter("EDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_End.Format(VAR_DATEVALUEONLY));
	try
	{
		bt->GetRS()->Execute(&pCmd);
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		return;
	}
	m_Data.SetRedraw(FALSE);
	// оформление грида
	csVal.Format("Статистика за период с %s по %s\n",m_Begin.Format("%x"),m_End.Format("%x"));
	csVal+=m_csTitle;
	SetDlgItemText(IDC_INFO,csVal);
	m_Data.SetRows(bt->GetRecordsCount()+1);
	m_Data.SetCols(bt->GetRS()->GetFieldCount()+1);
	m_Data.SetColWidth(0,100);
	for(int i=0; i<bt->GetRS()->GetFieldCount(); i++)
	{
		bt->GetRS()->GetFieldInfo(i,&info);
		m_Data.SetTextMatrix(0,i+1,info.m_strName);
		m_Data.SetColWidth(i+1,2500);
	}
	//данные
	while(!bt->GetRS()->IsEof())
	{
		for(long l=0; l<bt->GetRS()->GetFieldCount(); ++l)
		{
			csVal=bt->GetStringValue(l);
			csVal.TrimLeft(); csVal.TrimRight();
			if(csVal=="") csVal="Не определено";
			m_Data.SetTextMatrix((row)+1,l+1,csVal);
		}
		row++;
		bt->GetRS()->MoveNext();
	}
	m_Data.SetRedraw(TRUE);
	m_Data.Refresh();
	
}

void FStatistic::OnMUnload() 
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	if(bt->GetRS()->IsOpen())
		::SendMessage(AfxGetApp()->GetMainWnd()->GetSafeHwnd(),WM_COMMAND,ID_REPORTS_EXCEL,NULL);
}

void FStatistic::OnClose() 
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	bt->GetRS()->Close();
	
	CDialog::OnClose();

}
