// FOutbox.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FOutbox.h"
#include "FOutMessage.h"
#include "FWait.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FOutbox dialog


FOutbox::FOutbox(CWnd* pParent /*=NULL*/)
	: CDialog(FOutbox::IDD, pParent)
{
	//{{AFX_DATA_INIT(FOutbox)
	m_Begin = COleDateTime::GetCurrentTime();
	m_End = COleDateTime::GetCurrentTime();
	//}}AFX_DATA_INIT
	CString csVal;
	csVal="01/01/"+(COleDateTime::GetCurrentTime()).Format("%y");
	m_Begin.ParseDateTime(csVal);
	m_bKeypressed=FALSE;
	m_bExportSendMessage=TRUE;
	m_ToolTip=NULL;
	m_lIdleTime=0;
}


void FOutbox::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FOutbox)
	DDX_Control(pDX, IDC_LETTERLIST, m_LetterList);
	DDX_DateTimeCtrl(pDX, IDC_BEGIN, m_Begin);
	DDX_DateTimeCtrl(pDX, IDC_END, m_End);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FOutbox, CDialog)
	//{{AFX_MSG_MAP(FOutbox)
	ON_WM_SIZE()
	ON_WM_SHOWWINDOW()
	ON_WM_MOUSEWHEEL()
	ON_BN_CLICKED(IDC_EQUAL, OnEqual)
	ON_NOTIFY(DTN_CLOSEUP, IDC_BEGIN, OnCloseupBegin)
	ON_NOTIFY(DTN_CLOSEUP, IDC_END, OnCloseupEnd)
	ON_BN_CLICKED(IDC_EXPORT, OnExport)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FOutbox message handlers

BOOL FOutbox::OnInitDialog() 
{
	CRect rect;
	CDialog::OnInitDialog();
	CWnd *list=this->GetDlgItem(IDC_LETTERLIST);
	
	GetClientRect(&rect);
	m_iCx=rect.right-rect.left;
	m_iCy=rect.bottom-rect.top;

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void FOutbox::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	
	WINDOWPLACEMENT pm;
	
	CWnd *list=this->GetDlgItem(IDC_LETTERLIST);
	CWnd *button=this->GetDlgItem(IDC_COLOR);
	CWnd *export=this->GetDlgItem(IDC_EXPORT);
	if(list)
	{
		list->GetWindowPlacement(&pm);
		pm.rcNormalPosition.right+=	cx-m_iCx;
		pm.rcNormalPosition.bottom+=cy-m_iCy;
		m_iWidth=pm.rcNormalPosition.right-pm.rcNormalPosition.left;
		list->SetWindowPlacement(&pm);
	}
	if(button)
	{
		button->GetWindowPlacement(&pm);
		pm.rcNormalPosition.right+=	cx-m_iCx;
		pm.rcNormalPosition.left+=	cx-m_iCx;
		button->SetWindowPlacement(&pm);
	}
	if(export)
	{
		export->GetWindowPlacement(&pm);
		pm.rcNormalPosition.right+=	cx-m_iCx;
		pm.rcNormalPosition.left+=	cx-m_iCx;
		export->SetWindowPlacement(&pm);
	}
	m_iCx=cx;m_iCy=cy;
	
}

void FOutbox::Show()
{
	long row=0;
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), "[OutBox]");
	CADOFieldInfo info;
	CString csVal,attrib;
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	BOOL bOnly=((CKanApp*) AfxGetApp())->m_bOnlyMine;
	int user=((CKanApp*) AfxGetApp())->m_iUserID;
	int replyCol=GetPrivateProfileInt("Outbox","REPLY",3,csAppPath+"\\kan.ini");
	int size=0,alig=4;
	bt->SetDateFormat("%x");
		
	pCmd.AddParameter("BDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_Begin.Format(VAR_DATEVALUEONLY));
	pCmd.AddParameter("EDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_End.Format(VAR_DATEVALUEONLY));
	
	csVal="%"+m_outFilterAuthor;
	csVal+="%";
	pCmd.AddParameter("Author",CADORecordset::typeChar,
				CADOParameter::paramInput,csVal.GetLength(),csVal);
	csVal="%"+m_outFilterContent;
	csVal+="%";
	pCmd.AddParameter("Cont",CADORecordset::typeChar,
				CADOParameter::paramInput,csVal.GetLength(),csVal);
	csVal="%"+m_outFilterReciever;
	csVal+="%";
	pCmd.AddParameter("Rcvr",CADORecordset::typeChar,
				CADOParameter::paramInput,csVal.GetLength(),csVal);
	FWait* waitdlg=new FWait();
	waitdlg->Create(FWait::IDD);
	waitdlg->ShowWindow(SW_SHOW);
	waitdlg->SetInfo("Формирование...");
	GetDlgItem(IDC_GREEN)->SetWindowPos(NULL, 0, 0, 12, 12, SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE);
	//запрос
	bt->Execute(&pCmd);
	// оформление грида
	m_LetterList.Clear();
	m_LetterList.SetRows(bt->GetRecordsCount()+1);
	m_LetterList.SetCols(bt->GetRS()->GetFieldCount()+1);
	m_LetterList.SetColWidth(0,100);
	csVal.Format("Исходящие. Всего %d писем за период", bt->GetRecordsCount());
	SetDlgItemText(IDC_INFO,csVal);
	m_LetterList.SetRedraw(FALSE);
	for(int i=0; i<bt->GetRS()->GetFieldCount(); i++)
	{
		csVal.Format("Size%d",i+1);
		size=GetPrivateProfileInt("Outbox",csVal,0,csAppPath+"\\kan.ini");
		bt->GetRS()->GetFieldInfo(i,&info);
		m_LetterList.SetTextMatrix(0,i+1,info.m_strName);
		m_LetterList.SetColWidth(i+1,size);
		csVal.Format("Alig%d",i+1);
		alig=GetPrivateProfileInt("Outbox",csVal,4,csAppPath+"\\kan.ini");
		if(size>50) m_LetterList.SetColAlignment(i+1,alig);
	}

	//данные
	while(!bt->GetRS()->IsEof())
	{
		/*отражать только документы текущего пользователя*/
		if(bOnly) 
			if(atoi(bt->GetStringValue("UserID"))!=user) 
			{ 
				bt->GetRS()->MoveNext();
				continue;
			}

		for(long l=0; l<bt->GetRS()->GetFieldCount(); ++l)
		{
			csVal=bt->GetStringValue(l);
			csVal.TrimLeft(); csVal.TrimRight();
			m_LetterList.SetTextMatrix((row)+1,l+1,csVal);
			attrib=bt->GetStringValue("_Count");
			m_LetterList.SetRow(row+1);m_LetterList.SetRowSel(row+1);
			m_LetterList.SetCol(replyCol);m_LetterList.SetColSel(replyCol);
			if(atol(attrib)>0)
				m_LetterList.SetCellBackColor(RGB(199,228,197));
		}
		row++;
		bt->GetRS()->MoveNext();
	}
	if(bt->GetRecordsCount()>0)
	{
		m_LetterList.SetRow(1);m_LetterList.SetRowSel(1);
		m_LetterList.SetCol(1);
		m_LetterList.SetColSel(bt->GetRS()->GetFieldCount());
	}
	bt->GetRS()->Close();
	m_LetterList.SetRows(row+1);
	waitdlg->DestroyWindow(); delete waitdlg;
	m_LetterList.SetRedraw(TRUE);
}

void FOutbox::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	
	if(!bShow) return;
	Show();
	m_ToolTip =new CToolTipCtrl();
	if (!m_ToolTip->Create(this))
	{
		TRACE("Unable To create ToolTip\n");           
		return;
	}

	m_ToolTip->AddTool(GetDlgItem(IDC_EQUAL),"Установить одинаковые даты");
	m_ToolTip->AddTool(GetDlgItem(IDC_LETTERLIST),"Поле");
	m_ToolTip->SetMaxTipWidth(10000000);
	m_ToolTip->SetDelayTime(TTDT_AUTOPOP ,10000);
	m_ToolTip->SetDelayTime(TTDT_INITIAL ,500);
	m_ToolTip->SetDelayTime(TTDT_RESHOW ,10000);
	m_ToolTip->SetTipBkColor(0xEFE0DAL);
	m_ToolTip->SetTipTextColor(0xA00000L);
	m_ToolTip->Activate(TRUE);
	((CKanApp*) AfxGetApp())->GetMainWnd()->SetWindowText("Канцелярия - Исходящие");
}

BOOL FOutbox::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	long row=m_LetterList.GetTopRow();
	if(zDelta<0&&row<m_LetterList.GetRows()-1)
		m_LetterList.SetTopRow(row+1);
	if(zDelta>0&&row>1)
		m_LetterList.SetTopRow(row-1);
	
	return CDialog::OnMouseWheel(nFlags, zDelta, pt);
}

BEGIN_EVENTSINK_MAP(FOutbox, CDialog)
    //{{AFX_EVENTSINK_MAP(FOutbox)
	ON_EVENT(FOutbox, IDC_LETTERLIST, -600 /* Click */, OnClickLetterlist, VTS_NONE)
	ON_EVENT(FOutbox, IDC_LETTERLIST, -601 /* DblClick */, OnDblClickLetterlist, VTS_NONE)
	ON_EVENT(FOutbox, IDC_LETTERLIST, 69 /* SelChange */, OnSelChangeLetterlist, VTS_NONE)
	ON_EVENT(FOutbox, IDC_LETTERLIST, -606 /* MouseMove */, OnMouseMoveLetterlist, VTS_I2 VTS_I2 VTS_I4 VTS_I4)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

void FOutbox::OnClickLetterlist() 
{
	long rows=m_LetterList.GetRows();
	if(rows<2) return;
	long col=m_LetterList.GetMouseCol();
	long row=m_LetterList.GetMouseRow();
	long mode=m_LetterList.GetSelectionMode();
	m_LetterList.SetSelectionMode(0);
	if(row==0)
	{
		m_LetterList.SetCol(col);
		m_LetterList.SetColSel(col);
		m_bSortInc=!m_bSortInc;
		if(m_bSortInc) m_LetterList.SetSort(2); else m_LetterList.SetSort(1);
		UpdateData();
		m_LetterList.SetCol(0);
		m_LetterList.SetColSel(m_LetterList.GetCols() - 1);
	}

	m_LetterList.SetSelectionMode(mode);
	
}

void FOutbox::OnDblClickLetterlist() 
{
	long row=m_LetterList.GetRow();
	long mouserow=m_LetterList.GetMouseRow();
	if(mouserow==0&&!m_bKeypressed||row==0) return;
	FOutMessage dlg;
	doc->m_lLetterId=atol(m_LetterList.GetTextMatrix(row,1));
	doc->LoadOutData();
	dlg.doc=doc;
	if( dlg.DoModal()==IDOK)
	{
		doc->SaveOutData();
		Show();
	}
	m_bKeypressed=FALSE;
	
}

void FOutbox::OnEqual() 
{
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	m_Begin=m_End;
	UpdateData(FALSE);
	WritePrivateProfileString("Outbox","Begin",m_Begin.Format("%x"),csAppPath+"\\kan.ini");
	WritePrivateProfileString("Outbox","End",m_End.Format("%x"),csAppPath+"\\kan.ini");
	Show();
	
}

void FOutbox::OnCloseupBegin(NMHDR* pNMHDR, LRESULT* pResult) 
{
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	UpdateData();
	WritePrivateProfileString("Outbox","Begin",m_Begin.Format("%x"),csAppPath+"\\kan.ini");
	Show();
	
	*pResult = 0;
}

void FOutbox::OnCloseupEnd(NMHDR* pNMHDR, LRESULT* pResult) 
{
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	UpdateData();
	WritePrivateProfileString("Outbox","End",m_End.Format("%x"),csAppPath+"\\kan.ini");
	Show();
	
	*pResult = 0;
}

BOOL FOutbox::PreTranslateMessage(MSG* pMsg) 
{
	if (NULL != m_ToolTip)            
      m_ToolTip->RelayEvent(pMsg);
	if(WM_KEYDOWN == pMsg->message)
	{
		switch(pMsg->wParam)
		{
			case VK_RETURN:
				if(this->GetFocus()->GetDlgCtrlID()==IDC_LETTERLIST)
				{
					m_bKeypressed=TRUE;
					OnDblClickLetterlist();
					return TRUE;       // запрет дальнейшей обработки
				}
				break;
			case VK_ESCAPE: return TRUE;
			case VK_F2:
			case VK_F4:
			case VK_F5:
			case VK_F8:
			case VK_F6:
			case VK_F7:
				AfxGetApp()->GetMainWnd()->PostMessage(pMsg->message,
					pMsg->wParam,pMsg->lParam);
				break;
			default:break;
		}
	}
	
	return CDialog::PreTranslateMessage(pMsg);
}

BOOL FOutbox::DestroyWindow() 
{
	CString name,size;
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	for(int i=1; i< m_LetterList.GetCols(); i++ )
	{
		name.Format("Size%d",i);
		size.Format("%d",m_LetterList.GetColWidth(i));
		WritePrivateProfileString("Outbox",name,size,csAppPath+"\\kan.ini");

	}
	WritePrivateProfileString("Outbox","Begin",m_Begin.Format("%x"),csAppPath+"\\kan.ini");
	WritePrivateProfileString("Outbox","End",m_End.Format("%x"),csAppPath+"\\kan.ini");
	if(m_ToolTip) m_ToolTip->DestroyWindow();
	delete m_ToolTip;
	return CDialog::DestroyWindow();
}

void FOutbox::OnExport() 
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), "[OutboxExcel]");
	CString csVal;
	
	pCmd.AddParameter("BDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_Begin.Format(VAR_DATEVALUEONLY));
	pCmd.AddParameter("EDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_End.Format(VAR_DATEVALUEONLY));
	
	csVal="%"+m_outFilterAuthor;
	csVal+="%";
	pCmd.AddParameter("Author",CADORecordset::typeBSTR,
				CADOParameter::paramInput,512,csVal);
	csVal="%"+m_outFilterContent;
	csVal+="%";
	pCmd.AddParameter("Cont",CADORecordset::typeBSTR,
				CADOParameter::paramInput,512,csVal);

	//запрос
	try
	{
		bt->GetRS()->Execute(&pCmd);
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		return;
	}

	if(m_bExportSendMessage)
		::SendMessage(AfxGetApp()->GetMainWnd()->GetSafeHwnd(),WM_COMMAND,ID_REPORTS_EXCEL,NULL);	
}

void FOutbox::GetSetForExport()
{
	m_bExportSendMessage=FALSE;
	OnExport();
	m_bExportSendMessage=TRUE;
}

void FOutbox::OnSelChangeLetterlist() 
{
	m_LetterList.SetRowSel(m_LetterList.GetRow());
	
}

void FOutbox::OnMouseMoveLetterlist(short Button, short Shift, long x, long y) 
{
	CString str=_T("");
	long row=m_LetterList.GetMouseRow();
	int count=m_LetterList.GetCols();
	if(m_lMouseX==x && m_lMouseY==y)
	{
		if( m_lIdleTime==0)
		{
			for(int i=1; i<count; ++i)
			{
				if(m_LetterList.GetColWidth(i)==0) continue;
				if(str.GetLength()>0) str+="\n";
				str+=m_LetterList.GetTextMatrix(0,i);
				str+=":  ";
				str+=m_LetterList.GetTextMatrix(row,i);
			}
			m_ToolTip->UpdateTipText(str,GetDlgItem(IDC_LETTERLIST));
			m_lIdleTime=1;
		}
	}
	else
	{
		m_lMouseX=x;
		m_lMouseY=y;
		m_ToolTip->Pop();
		m_lIdleTime=0;
	}	
}
