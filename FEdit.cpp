// FEdit.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FEdit.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FEdit dialog


FEdit::FEdit(CWnd* pParent /*=NULL*/)
	: CDialog(FEdit::IDD, pParent)
{
	//{{AFX_DATA_INIT(FEdit)
	m_csQuery = _T("");
	//}}AFX_DATA_INIT
}


void FEdit::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FEdit)
	DDX_Control(pDX, IDC_DATA, m_Data);
	DDX_Text(pDX, IDC_QUERY, m_csQuery);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FEdit, CDialog)
	//{{AFX_MSG_MAP(FEdit)
	ON_BN_CLICKED(IDC_RUN, OnRun)
	ON_WM_SHOWWINDOW()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FEdit message handlers

void FEdit::OnRun() 
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOFieldInfo info;
	CString csVal;
	long row=0;
	UpdateData();
	try
	{
		bt->Execute(m_csQuery);
	}
	catch(...)
	{
		return;
	}
	// оформление грида
	m_Data.SetRedraw(FALSE);
	m_Data.SetRows(bt->GetRecordsCount()+1);
	m_Data.SetCols(bt->GetRS()->GetFieldCount()+1);
	int* size=new int[bt->GetRS()->GetFieldCount()]; 
	m_Data.SetColWidth(0,300);
	
	for(int i=0; i<bt->GetRS()->GetFieldCount(); i++)
	{
		bt->GetRS()->GetFieldInfo(i,&info);
		m_Data.SetTextMatrix(0,i+1,info.m_strName);
		size[i]=strlen(info.m_strName);
	}
	//данные
	if(bt->GetRecordsCount()>0) bt->GetRS()->MoveFirst();
	while(!bt->GetRS()->IsEof())
	{
		csVal.Format("%d",row);
		m_Data.SetTextMatrix(row+1,0,csVal);
		for(long l=0; l<bt->GetRS()->GetFieldCount(); ++l)
		{
			csVal=bt->GetStringValue(l);
			csVal.TrimLeft(); csVal.TrimRight();
			if(csVal=="") csVal="Не определено";
			m_Data.SetTextMatrix(row+1,l+1,csVal);
			
			if(size[l]<csVal.GetLength()) size[l]=csVal.GetLength();
		}
		row++;
		bt->GetRS()->MoveNext();
	}
	// ширина
	for(i=0; i<bt->GetRS()->GetFieldCount(); ++i)
	{
		m_Data.SetColWidth(i+1,size[i]*100);
	}

	bt->GetRS()->Close();
	m_Data.SetRedraw(TRUE);
	delete[] size;
}

void FEdit::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	
	if(bShow) GetDlgItem(IDC_QUERY)->SetFocus();
	
}

BOOL FEdit::PreTranslateMessage(MSG* pMsg) 
{
	CWnd* focus=GetFocus();
	if(WM_KEYDOWN == pMsg->message)
	{
		int id=focus->GetDlgCtrlID();
		switch(pMsg->wParam)
		{
			case VK_RETURN:
					if(id!=IDC_QUERY) 
						OnRun();
					else
						return CDialog::PreTranslateMessage(pMsg);
				break;
			default:
				return CDialog::PreTranslateMessage(pMsg);
		}
		return TRUE;       // запрет дальнейшей обработки
	}
	return CDialog::PreTranslateMessage(pMsg);
}
