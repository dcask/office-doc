// FInbox.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FInbox.h"
#include "FInMessage.h"
#include "FWait.h"
//#include "mmsystem.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FInbox dialog


FInbox::FInbox(CWnd* pParent /*=NULL*/)
	: CDialog(FInbox::IDD, pParent)
{
	//{{AFX_DATA_INIT(FInbox)
	m_Begin = COleDateTime::GetCurrentTime();
	m_End = COleDateTime::GetCurrentTime();
	//}}AFX_DATA_INIT
	// определяем дату начала года и конца
	CString csVal;
	csVal="01/01/"+(COleDateTime::GetCurrentTime()).Format("%y");
	m_Begin.ParseDateTime(csVal);
	//csVal="31/12/"+(COleDateTime::GetCurrentTime()).Format("%y");
	m_bKeypressed=FALSE;
	m_bExportSendMessage=TRUE;
	m_csQueryName="[InBox]";
	m_ToolTip=NULL;
	m_lIdleTime=0;
}


void FInbox::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FInbox)
	DDX_Control(pDX, IDC_LETTERLIST, m_LetterList);
	DDX_DateTimeCtrl(pDX, IDC_BEGIN, m_Begin);
	DDX_DateTimeCtrl(pDX, IDC_END, m_End);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FInbox, CDialog)
	//{{AFX_MSG_MAP(FInbox)
	ON_WM_SHOWWINDOW()
	ON_WM_SIZE()
	ON_WM_MOUSEWHEEL()
	ON_NOTIFY(DTN_CLOSEUP, IDC_BEGIN, OnCloseupBegin)
	ON_NOTIFY(DTN_CLOSEUP, IDC_END, OnCloseupEnd)
	ON_BN_CLICKED(IDC_EQUAL, OnEqual)
	ON_BN_CLICKED(IDC_EXPORT, OnExport)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FInbox message handlers
void FInbox::OnShowWindow(BOOL bShow, UINT nStatus) 


{
	CDialog::OnShowWindow(bShow, nStatus);
	if(!bShow) return;
	this->Show();
	m_LetterList.SetFocus();
	m_ToolTip =new CToolTipCtrl();
	if (!m_ToolTip->Create(this))
	{
		TRACE("Unable To create ToolTip\n");           
		return;
	}

	m_ToolTip->AddTool(GetDlgItem(IDC_EQUAL),"Установить одинаковые даты");
	m_ToolTip->AddTool(GetDlgItem(IDC_EXPORT),"Выгрузить список в Excel");
	m_ToolTip->AddTool(GetDlgItem(IDC_LETTERLIST),"Поле");
	m_ToolTip->SetMaxTipWidth(10000000);
	m_ToolTip->SetDelayTime(TTDT_AUTOPOP ,10000);
	m_ToolTip->SetDelayTime(TTDT_INITIAL ,500);
	m_ToolTip->SetDelayTime(TTDT_RESHOW ,10000);
	m_ToolTip->SetTipBkColor(0xEFE0DAL);
	m_ToolTip->SetTipTextColor(0xA00000L);
	m_ToolTip->Activate(TRUE);
	((CKanApp*) AfxGetApp())->GetMainWnd()->SetWindowText("Канцелярия - Входящие");
}

BOOL FInbox::OnInitDialog() 
{
	CRect rect;
	CRect rc;
	CDialog::OnInitDialog();
	
	
	GetClientRect(&rect);
	m_iCx=rect.right-rect.left;
	m_iCy=rect.bottom-rect.top;
	/*удаление из временного хранилища, если было некорректно */
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), "[DelAllTemp]");
	pCmd.AddParameter("User",CADORecordset::typeInteger,
					CADOParameter::paramInput,sizeof(int),((CKanApp*)AfxGetApp())->m_iUserID);
	try
	{
		bt->GetRS()->Execute(&pCmd);
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		return FALSE;
	}
	bt->GetRS()->Close();

	/**/

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void FInbox::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	
	WINDOWPLACEMENT pm;
	
	CWnd *list=this->GetDlgItem(IDC_LETTERLIST);
	CWnd *button=this->GetDlgItem(IDC_COLOR);
	CWnd *begin=this->GetDlgItem(IDC_BEGIN);
	CWnd *export=this->GetDlgItem(IDC_EXPORT);
	CWnd *end=this->GetDlgItem(IDC_END);
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

BOOL FInbox::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	long row=m_LetterList.GetTopRow();
	if(zDelta<0&&row<m_LetterList.GetRows()-1)
		m_LetterList.SetTopRow(row+1);
	if(zDelta>0&&row>1)
		m_LetterList.SetTopRow(row-1);
	
	return CDialog::OnMouseWheel(nFlags, zDelta, pt);
}

BEGIN_EVENTSINK_MAP(FInbox, CDialog)
    //{{AFX_EVENTSINK_MAP(FInbox)
	ON_EVENT(FInbox, IDC_LETTERLIST, -600 /* Click */, OnClickLetterlist, VTS_NONE)
	ON_EVENT(FInbox, IDC_LETTERLIST, -601 /* DblClick */, OnDblClickLetterlist, VTS_NONE)
	ON_EVENT(FInbox, IDC_LETTERLIST, 69 /* SelChange */, OnSelChangeLetterlist, VTS_NONE)
	ON_EVENT(FInbox, IDC_LETTERLIST, -606 /* MouseMove */, OnMouseMoveLetterlist, VTS_I2 VTS_I2 VTS_I4 VTS_I4)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

void FInbox::OnClickLetterlist() 
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

void FInbox::OnDblClickLetterlist() 
{
	long row=m_LetterList.GetRow();
	long mouserow=m_LetterList.GetMouseRow();
	if(mouserow==0&&!m_bKeypressed||row==0) return;
	FInMessage dlg;
	doc->m_lLetterId=atol(m_LetterList.GetTextMatrix(row,1));
	doc->LoadInData();
	dlg.doc=doc;

	if(dlg.DoModal()==IDOK)
	{
		doc->SaveInData();
		Show();
	}
	m_bKeypressed=FALSE;
}

void FInbox::Show()
{
	long row=0;
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), m_csQueryName);
	CADOFieldInfo info;
	CString csVal,control,attrib;
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	BOOL bOnly=((CKanApp*) AfxGetApp())->m_bOnlyMine;
	int user=((CKanApp*) AfxGetApp())->m_iUserID;
	int controlCol=GetPrivateProfileInt("Inbox","CONTROL",4,csAppPath+"\\kan.ini");
	int completeCol=GetPrivateProfileInt("Inbox","COMPLETE",6,csAppPath+"\\kan.ini");
	int size=0, alig=4;

	bt->SetDateFormat("%x");
		
	pCmd.AddParameter("BDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_Begin.Format(VAR_DATEVALUEONLY));
	pCmd.AddParameter("EDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_End.Format(VAR_DATEVALUEONLY));
	
	csVal="%"+m_inFilterDeclarant;
	csVal+="%";
	pCmd.AddParameter("Decl",CADORecordset::typeChar,
				CADOParameter::paramInput,csVal.GetLength(),csVal);
	csVal="%"+m_inFilterContent;
	csVal+="%";
	pCmd.AddParameter("Cont",CADORecordset::typeChar,
				CADOParameter::paramInput,csVal.GetLength(),csVal);
	FWait* waitdlg=new FWait();
	waitdlg->Create(FWait::IDD);
	waitdlg->ShowWindow(SW_SHOW);
	waitdlg->SetInfo("Формирование списка...");
	/**/
	GetDlgItem(IDC_RED)->SetWindowPos(NULL, 0, 0, 12, 12, SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE);
	GetDlgItem(IDC_GREEN)->SetWindowPos(NULL, 0, 0, 12, 12, SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE);
	GetDlgItem(IDC_GRAY)->SetWindowPos(NULL, 0, 0, 12, 12, SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE);
	/**/
	bt->Execute(&pCmd);
	// оформление грида
	m_LetterList.Clear();
	m_LetterList.SetRows(bt->GetRecordsCount()+1);
	m_LetterList.SetCols(bt->GetRS()->GetFieldCount()+1);
	m_LetterList.SetColWidth(0,100);
	csVal.Format("Входящие. Всего %d писем за период", bt->GetRecordsCount());
	SetDlgItemText(IDC_INFO,csVal);
	m_LetterList.SetRedraw(FALSE);
	for(int i=0; i<bt->GetRS()->GetFieldCount(); i++)
	{
		csVal.Format("Size%d",i+1);
		size=GetPrivateProfileInt("Inbox",csVal,0,csAppPath+"\\kan.ini");
		bt->GetRS()->GetFieldInfo(i,&info);
		m_LetterList.SetTextMatrix(0,i+1,info.m_strName);
		m_LetterList.SetColWidth(i+1,size);
		csVal.Format("Alig%d",i+1);
		alig=GetPrivateProfileInt("Inbox",csVal,4,csAppPath+"\\kan.ini");
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
			control=bt->GetStringValue("Control");
			attrib=bt->GetStringValue("Attribute");
			m_LetterList.SetRow(row+1);m_LetterList.SetRowSel(row+1);
			m_LetterList.SetCol(controlCol);m_LetterList.SetColSel(controlCol);
			if(control=="1")
				m_LetterList.SetCellBackColor(RGB(228,199,197));
			else 
				m_LetterList.SetCellBackColor(RGB(199,228,197));
			if(attrib=="1")
			{
				m_LetterList.SetRow(row+1);m_LetterList.SetRowSel(row+1);
				m_LetterList.SetCol(completeCol);m_LetterList.SetColSel(completeCol);
				m_LetterList.SetCellBackColor(RGB(215,210,210));
			}
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
	m_LetterList.SetRows(row+1);
	waitdlg->DestroyWindow(); delete waitdlg;
	bt->GetRS()->Close();
	m_LetterList.SetRedraw(TRUE);
}

void FInbox::OnCloseupBegin(NMHDR* pNMHDR, LRESULT* pResult) 
{
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	UpdateData();
	WritePrivateProfileString("Inbox","Begin",m_Begin.Format("%x"),csAppPath+"\\kan.ini");
	Show();
	
	*pResult = 0;
}

void FInbox::OnCloseupEnd(NMHDR* pNMHDR, LRESULT* pResult) 
{
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	UpdateData();

	WritePrivateProfileString("Inbox","End",m_End.Format("%x"),csAppPath+"\\kan.ini");
	Show();
	
	*pResult = 0;
}

void FInbox::OnEqual() 
{
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	m_Begin=m_End;
	UpdateData(FALSE);
	WritePrivateProfileString("Inbox","Begin",m_Begin.Format("%x"),csAppPath+"\\kan.ini");
	WritePrivateProfileString("Inbox","End",m_End.Format("%x"),csAppPath+"\\kan.ini");
	Show();
}

BOOL FInbox::PreTranslateMessage(MSG* pMsg) 
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
			case VK_ESCAPE:return TRUE;
			case VK_F2:
			case VK_F4:
			case VK_F5:
			case VK_F6:
			case VK_F7:
			case VK_F8:
				AfxGetApp()->GetMainWnd()->PostMessage(pMsg->message,
					pMsg->wParam,pMsg->lParam);
				break;
			default:break;
		}
	}
    

	
	return CDialog::PreTranslateMessage(pMsg);
}

BOOL FInbox::DestroyWindow() 
{
	CString name,size;
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	for(int i=1; i< m_LetterList.GetCols(); i++ )
	{
		name.Format("Size%d",i);
		size.Format("%d",m_LetterList.GetColWidth(i));
		WritePrivateProfileString("Inbox",name,size,csAppPath+"\\kan.ini");

	}
	WritePrivateProfileString("Inbox","Begin",m_Begin.Format("%x"),csAppPath+"\\kan.ini");
	WritePrivateProfileString("Inbox","End",m_End.Format("%x"),csAppPath+"\\kan.ini");
	m_ToolTip->DestroyWindow();
	if(m_ToolTip) delete m_ToolTip;
	return CDialog::DestroyWindow();
}

void FInbox::OnExport() 
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), "[InboxExcel]");
	CString csVal;
	
	pCmd.AddParameter("BDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_Begin.Format(VAR_DATEVALUEONLY));
	pCmd.AddParameter("EDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_End.Format(VAR_DATEVALUEONLY));
	
	csVal="%"+m_inFilterDeclarant;
	csVal+="%";
	pCmd.AddParameter("Decl",CADORecordset::typeBSTR,
				CADOParameter::paramInput,200,csVal.Left(199));
	csVal="%"+m_inFilterContent;
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

void FInbox::GetSetForExport()
{
	m_bExportSendMessage=FALSE;
	OnExport();
	m_bExportSendMessage=TRUE;
}

void FInbox::OnSelChangeLetterlist() 
{
	m_LetterList.SetRowSel(m_LetterList.GetRow());
	
}

void FInbox::OnMouseMoveLetterlist(short Button, short Shift, long x, long y) 
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
