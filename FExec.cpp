// FExec.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FExec.h"
#include "FExecList.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FExec dialog


#define DATE_COLLUMN 3
#define DAYS_COLLUMN 2
#define ID_COLLUMN 5

FExec::FExec(CWnd* pParent /*=NULL*/)
	: CDialog(FExec::IDD, pParent)
{
	//{{AFX_DATA_INIT(FExec)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	m_iSize=1500;
	m_bAnswer=FALSE;
	m_ToolTip=NULL;
	m_lInLetterID=0;
}


void FExec::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FExec)
	DDX_Control(pDX, IDC_DAYS, m_Days);
	DDX_Control(pDX, IDC_SCHEDULE, m_Schedule);
	DDX_Control(pDX, IDC_FACT, m_Fact);
	DDX_Control(pDX, IDC_DATA, m_Data);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FExec, CDialog)
	//{{AFX_MSG_MAP(FExec)
	ON_BN_CLICKED(IDC_CHANGE, OnChange)
	ON_WM_SHOWWINDOW()
	ON_NOTIFY(NM_KILLFOCUS, IDC_SCHEDULE, OnKillfocusSchedule)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FExec message handlers

void FExec::OnChange() 
{
	FExecList dlg;
	dlg.m_paExecuters=m_paExecuters;
	dlg.DoModal();
	View();
}

void FExec::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CString csVal;

	if(!bShow) return;
	/*Если есть ответ*/
	if(m_bAnswer) 
	{
		CADOCommand pCmdIn(bt->GetDB(), "[GetReply]");
	
		pCmdIn.AddParameter("[InID]",CADORecordset::typeBigInt,
					CADOParameter::paramInput,sizeof(long),m_lInLetterID);	
		try
		{
			bt->GetRS()->Execute(&pCmdIn);
		}catch(CADOException* e)
		{
			bt->ThrowError(e);
			return;
		}
		if(bt->GetRecordsCount()==1) csVal=bt->GetStringValue("FullNum");
		bt->GetRS()->Close();
		GetDlgItem(IDC_ANSWER)->SetWindowText("Ответ-"+csVal);
		GetDlgItem(IDC_ANSWER)->ShowWindow(SW_SHOW);
	}
	View();
	m_ToolTip =new CToolTipCtrl();
	if (!m_ToolTip->Create(this))
	{
		TRACE("Unable To create ToolTip\n");           
		return;
	}

	m_ToolTip->AddTool(GetDlgItem(IDC_CHANGE),"Изменить список исполнителей");
	m_ToolTip->Activate(TRUE);
	
}

void FExec::View()
{
	Executer* exec;
	int count=1;
	CString csVal;
	m_Data.Clear();
	m_Data.SetCols(6);
	m_Data.SetRows(1);
	m_Data.SetColWidth(0,100);
	m_Data.SetColWidth(1,4800);m_Data.SetTextMatrix(0,1,"Исполнитель");
	m_Data.SetColWidth(2,m_iSize/2); m_Data.SetTextMatrix(0,2,"Дней");
	m_Data.SetColWidth(3,m_iSize); m_Data.SetTextMatrix(0,3,"План");
	m_Data.SetColWidth(4,m_iSize); m_Data.SetTextMatrix(0,4,"Факт");
	m_Data.SetColWidth(5,0);m_Data.SetTextMatrix(0,5,"");
	for(int i=0; i<(*m_paExecuters).GetSize(); i++)
	{
		exec=(Executer*)((*m_paExecuters).GetAt(i));
		if(exec->state!=2)
		{
			count++;
			m_Data.SetRows(count);
			m_Data.SetTextMatrix(count-1,1,exec->csName);
			m_Data.SetTextMatrix(count-1,3,exec->scheduleDate.Format("%x"));
			m_Data.SetTextMatrix(count-1,4,exec->factDate.Format("%x"));
			csVal.Format("%d",i);
			m_Data.SetTextMatrix(count-1,5,csVal);
			m_Data.SetRowHeight(count-1,320);
		}
	}
	m_Data.SetCol(1);m_Data.SetColSel(1);
	m_Data.SetSort(1);
}

BEGIN_EVENTSINK_MAP(FExec, CDialog)
    //{{AFX_EVENTSINK_MAP(FExec)
	ON_EVENT(FExec, IDC_DATA, 71 /* EnterCell */, OnEnterCellData, VTS_NONE)
	ON_EVENT(FExec, IDC_DATA, 72 /* LeaveCell */, OnLeaveCellData, VTS_NONE)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

void FExec::OnEnterCellData() 
{
	COleDateTime data;
	BOOL result=TRUE;
	int row = m_Data.GetRow();
	int col = m_Data.GetCol();
	CString csVal=m_Data.GetTextMatrix(row,col);
	if(row<1 || (col !=DATE_COLLUMN&&col !=DAYS_COLLUMN)) return;
	

	if (col==DATE_COLLUMN)
	{
		if(csVal!="") result=data.ParseDateTime(csVal);
		else data.m_status=COleDateTime::null;
		if(result) m_Schedule.SetTime(data);// else m_Schedule.SetTime(data);
		
	}
	CDC* pDC = GetDC();
	int lx = pDC->GetDeviceCaps(LOGPIXELSX);
	int ly = pDC->GetDeviceCaps(LOGPIXELSY);
	ReleaseDC(pDC);

	int h = (m_Data.GetRowHeight(row) * ly) / 1440-2;
	int w = (m_Data.GetColWidth(col) * lx) / 1440-2;

	CRect r;
	m_Data.GetWindowRect(&r);
	ScreenToClient(&r);
	int t1 = m_Data.GetCellLeft();
	int t2 = m_Data.GetCellTop();
	int sx = (t1 * ly) / 1440 + r.left;
	int sy = (t2 * lx) / 1440 + r.top;

	if(col==DATE_COLLUMN)
	{
		m_Schedule.SetWindowPos(NULL, sx+1, sy+1, w, h, SWP_SHOWWINDOW);
		m_Schedule.ShowWindow(SW_SHOW);
		m_Schedule.SetFocus();
	}
	if (col==DAYS_COLLUMN)
	{
		m_Days.SetWindowPos(NULL, sx+1, sy+1, w, h, SWP_SHOWWINDOW);
		m_Days.ShowWindow(SW_SHOW);
		m_Days.SetFocus();
	}
	
	
}

void FExec::OnLeaveCellData() 
{
	

	int row = m_Data.GetRow();
	int col = m_Data.GetCol();
	COleDateTime data;
	COleDateTimeSpan ts;
	CString days;
	int r;
	Executer* ex;
	if(row<1 || (col !=DATE_COLLUMN&&col !=DAYS_COLLUMN)) return;
	if (col==DATE_COLLUMN)
	{
		m_Schedule.GetTime(data);
		m_Data.SetTextMatrix(row,DATE_COLLUMN,data.Format("%x"));
		m_Data.SetFocus();
		m_Schedule.ShowWindow(SW_HIDE);
		r=atoi(m_Data.GetTextMatrix(row,ID_COLLUMN));
		ex=(Executer*)((*m_paExecuters).GetAt(r));
		ex->scheduleDate=data;
	}
	if (col==DAYS_COLLUMN)
	{
		m_Days.GetWindowText(days);
		if(days!="")
		{
			data=m_ProcessDate;
			ts=atoi(days);
			data+=ts;
			m_Data.SetFocus();
			m_Data.SetTextMatrix(row,DATE_COLLUMN,data.Format("%x"));
			m_Data.SetFocus();
			r=atoi(m_Data.GetTextMatrix(row,ID_COLLUMN));
			ex=(Executer*)((*m_paExecuters).GetAt(r));
			ex->scheduleDate=data;
		}

		m_Days.SetWindowText(_T(""));
		m_Days.ShowWindow(SW_HIDE);
		
	}
	
}

void FExec::OnKillfocusSchedule(NMHDR* pNMHDR, LRESULT* pResult) 
{
	if(! (m_Schedule.GetMonthCalCtrl()) )
	{
		OnLeaveCellData() ;
		m_Data.SetCol(1);
	}
	
	*pResult = 0;
}

BOOL FExec::PreTranslateMessage(MSG* pMsg) 
{
	if (NULL != m_ToolTip)            
      m_ToolTip->RelayEvent(pMsg);
	if(WM_KEYDOWN == pMsg->message)
	{
		switch(pMsg->wParam)
		{
			case VK_RETURN:

				if(this->GetFocus()->GetDlgCtrlID()==IDC_DAYS)
				{
					OnLeaveCellData();
					m_Data.SetCol(1);
					return TRUE;       // запрет дальнейшей обработки
				}
				break;
			default:break;
		}
	}
	
	return CDialog::PreTranslateMessage(pMsg);
}

/*void FExec::AddDate(int days)
{
	
}*/

BOOL FExec::DestroyWindow() 
{
	if(m_ToolTip) m_ToolTip->DestroyWindow();
	delete m_ToolTip;
	
	return CDialog::DestroyWindow();
}
