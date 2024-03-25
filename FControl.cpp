// FControl.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FControl.h"
#include "FInMessage.h"
#include "FWait.h"
#include <lm.h>
#include <atlconv.h>

#define UNICODE

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#ifdef UNICODE
#define WTEXT(s) L ## s
#else
#define WTEXT(s) s
#endif

/////////////////////////////////////////////////////////////////////////////
// FControl dialog


FControl::FControl(CWnd* pParent /*=NULL*/)
	: CDialog(FControl::IDD, pParent)
{
	//{{AFX_DATA_INIT(FControl)
	m_Begin = COleDateTime::GetCurrentTime();
	m_End = COleDateTime::GetCurrentTime();
	m_csInfo = _T("");
	//}}AFX_DATA_INIT
	// определяем дату начала года и конца
	CString csVal;
	csVal="01/01/"+(COleDateTime::GetCurrentTime()).Format("%y");
	m_Begin.ParseDateTime(csVal);
	//csVal="31/12/"+(COleDateTime::GetCurrentTime()).Format("%y");
	m_bKeypressed=FALSE;
	m_csQueryName="[ControlList]";
	m_bExportSendMessage=TRUE;
	m_ToolTip=NULL;
	m_lIdleTime=0;
}


void FControl::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FControl)
	DDX_DateTimeCtrl(pDX, IDC_BEGIN, m_Begin);
	DDX_DateTimeCtrl(pDX, IDC_END, m_End);
	DDX_Text(pDX, IDC_INFO, m_csInfo);
	DDX_Control(pDX, IDC_LETTERLIST, m_LetterList);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FControl, CDialog)
	//{{AFX_MSG_MAP(FControl)
	ON_WM_SHOWWINDOW()
	ON_WM_SIZE()
	ON_WM_MOUSEWHEEL()
	ON_BN_CLICKED(IDC_EQUAL, OnEqual)
	ON_NOTIFY(DTN_CLOSEUP, IDC_BEGIN, OnCloseupBegin)
	ON_NOTIFY(DTN_CLOSEUP, IDC_END, OnCloseupEnd)
	ON_BN_CLICKED(IDC_SEND, OnSend)
	ON_BN_CLICKED(IDC_EXPORT, OnExport)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FControl message handlers

void FControl::OnShowWindow(BOOL bShow, UINT nStatus) 
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

	m_ToolTip->AddTool(GetDlgItem(IDC_SEND),"Оповестить всех из списка");
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
	((CKanApp*) AfxGetApp())->GetMainWnd()->SetWindowText("Канцелярия - Контроль");
	
}

void FControl::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	
	WINDOWPLACEMENT pm;
	
	CWnd *list=this->GetDlgItem(IDC_LETTERLIST);
	CWnd *button=this->GetDlgItem(IDC_SEND);
	CWnd *export=this->GetDlgItem(IDC_EXPORT);
	CWnd *begin=this->GetDlgItem(IDC_BEGIN);
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

BOOL FControl::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	CRect rect;
	//WINDOWPLACEMENT pm;

	GetClientRect(&rect);
	m_iCx=rect.right-rect.left;
	m_iCy=rect.bottom-rect.top;

	/**/
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

BOOL FControl::PreTranslateMessage(MSG* pMsg) 
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
			case VK_F6:
				AfxGetApp()->GetMainWnd()->PostMessage(pMsg->message,
					pMsg->wParam,pMsg->lParam);
				break;
			default:break;
		}
	}
		
	return CDialog::PreTranslateMessage(pMsg);
}

BOOL FControl::DestroyWindow() 
{
	CString name,size;
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	for(int i=1; i< m_LetterList.GetCols(); i++ )
	{
		name.Format("Size%d",i);
		size.Format("%d",m_LetterList.GetColWidth(i));
		WritePrivateProfileString("Control",name,size,csAppPath+"\\kan.ini");

	}
	if(m_ToolTip) m_ToolTip->DestroyWindow();
	delete m_ToolTip;
	return CDialog::DestroyWindow();
	
}

BOOL FControl::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	long row=m_LetterList.GetTopRow();
	if(zDelta<0&&row<m_LetterList.GetRows()-1)
		m_LetterList.SetTopRow(row+1);
	if(zDelta>0&&row>1)
		m_LetterList.SetTopRow(row-1);
	
	return CDialog::OnMouseWheel(nFlags, zDelta, pt);
}

BEGIN_EVENTSINK_MAP(FControl, CDialog)
    //{{AFX_EVENTSINK_MAP(FControl)
	ON_EVENT(FControl, IDC_LETTERLIST, -600 /* Click */, OnClickLetterlist, VTS_NONE)
	ON_EVENT(FControl, IDC_LETTERLIST, -601 /* DblClick */, OnDblClickLetterlist, VTS_NONE)
	ON_EVENT(FControl, IDC_LETTERLIST, 69 /* SelChange */, OnSelChangeLetterlist, VTS_NONE)
	ON_EVENT(FControl, IDC_LETTERLIST, -606 /* MouseMove */, OnMouseMoveLetterlist, VTS_I2 VTS_I2 VTS_I4 VTS_I4)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

void FControl::OnClickLetterlist() 
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

void FControl::OnDblClickLetterlist() 
{
	long row=m_LetterList.GetRow();
	long mouserow=m_LetterList.GetMouseRow();
	if(mouserow==0&&!m_bKeypressed) return;
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

void FControl::Show()
{
	long row=0;
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), "[ControlList]");
	CADOFieldInfo info;
	CString csVal,control,attrib, netname;
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	BOOL bOnly=((CKanApp*) AfxGetApp())->m_bOnlyMine;
	int user=((CKanApp*) AfxGetApp())->m_iUserID;
	int size=0, alig=4;
	bt->SetDateFormat("%x");
		
	pCmd.AddParameter("BDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_Begin.Format(VAR_DATEVALUEONLY));
	pCmd.AddParameter("EDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_End.Format(VAR_DATEVALUEONLY));
	
	csVal="%"+m_filterDeclarant;
	csVal+="%";
	pCmd.AddParameter("Decl",CADORecordset::typeChar,
				CADOParameter::paramInput,200,csVal.Left(199));
	FWait* waitdlg=new FWait();
	waitdlg->Create(FWait::IDD);
	waitdlg->ShowWindow(SW_SHOW);
	waitdlg->SetInfo("Формирование...");

	//запрос
	try
	{
		bt->GetRS()->Execute(&pCmd);
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		return;
	}
	// оформление грида
		m_LetterList.Clear();
		m_LetterList.SetRows(bt->GetRecordsCount()+1);
		m_LetterList.SetCols(bt->GetRS()->GetFieldCount()+1);
		m_LetterList.SetColWidth(0,100);
	csVal.Format("Контроль. Всего %d человек за период", bt->GetRecordsCount());
	SetDlgItemText(IDC_INFO,csVal);
	m_LetterList.SetRedraw(FALSE);
	for(int i=0; i<bt->GetRS()->GetFieldCount(); i++)
	{
		csVal.Format("Size%d",i+1);
		size=GetPrivateProfileInt("Control",csVal,0,csAppPath+"\\kan.ini");
		bt->GetRS()->GetFieldInfo(i,&info);
		m_LetterList.SetTextMatrix(0,i+1,info.m_strName);
		m_LetterList.SetColWidth(i+1,size);
		csVal.Format("Alig%d",i+1);
		alig=GetPrivateProfileInt("Control",csVal,4,csAppPath+"\\kan.ini");
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

void FControl::OnEqual() 
{
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	m_End=m_Begin;
	UpdateData(FALSE);
	Show();
	
}

void FControl::OnCloseupBegin(NMHDR* pNMHDR, LRESULT* pResult) 
{
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	UpdateData();
	Show();
	
	*pResult = 0;
}

void FControl::OnCloseupEnd(NMHDR* pNMHDR, LRESULT* pResult) 
{
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	UpdateData();

	Show();
	
	*pResult = 0;
}

void FControl::OnSend() 
{
	if(!((CKanApp*)AfxGetApp())->m_bMessage) return;
	if(AfxMessageBox("Отослать сообщение всем из списка ?", MB_YESNO | MB_ICONQUESTION)!=IDYES) return;

	// ERROR_ACCESS_DENIED NERR_Success. ERROR_NOT_SUPPORTED NERR_NameNotFound NERR_NetworkError
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	CString netname, date, num, decl, regdate, content;
	CString tmp,message,templ;
	int netnamecol=GetPrivateProfileInt("Control","NETNAMECOL",9,csAppPath+"\\kan.ini");
	int iNum=GetPrivateProfileInt("Control","NUMCOL",1,csAppPath+"\\kan.ini");
	int iDate=GetPrivateProfileInt("Control","DATECOL",2,csAppPath+"\\kan.ini");
	int iDecl=GetPrivateProfileInt("Control","DECLCOL",3,csAppPath+"\\kan.ini");
	int iRegDate=GetPrivateProfileInt("Control","RDATECOL",4,csAppPath+"\\kan.ini");
	int iContent=GetPrivateProfileInt("Control","CONTCOL",5,csAppPath+"\\kan.ini");
	GetPrivateProfileString("Control","Message","Напоминание",templ.GetBuffer(512),512,csAppPath+"\\kan.ini");
	templ.ReleaseBuffer();
	int len=0;

	NET_API_STATUS nasStatus;
	USES_CONVERSION;
	for(int i=1;i<m_LetterList.GetRows(); i++)
	{
		message=templ;
		num=m_LetterList.GetTextMatrix(i,iNum);
		date=m_LetterList.GetTextMatrix(i,iDate);
		decl=m_LetterList.GetTextMatrix(i,iDecl);
		regdate=m_LetterList.GetTextMatrix(i,iRegDate);
		content=m_LetterList.GetTextMatrix(i,iContent);
		message.Replace("[%DocNum%]",num);
		message.Replace("[%ExecDate%]",date);
		message.Replace("[%Declarant%]",decl);
		message.Replace("[%RegDate%]",regdate);
		message.Replace("[%Content%]",content);
		tmp=message.Left(511);
		message=tmp;
		len=message.GetLength();
		netname=m_LetterList.GetTextMatrix(i,netnamecol);
		nasStatus = NetMessageBufferSend(NULL , T2CW(netname.GetBuffer(0)) , NULL, LPBYTE(T2CW(message.GetBuffer(0))), len*2+1);
		netname.ReleaseBuffer();message.ReleaseBuffer();
		tmp.Format("%d ",nasStatus);
		tmp+="Сообщение для "+m_LetterList.GetTextMatrix(i,netnamecol);
		if(nasStatus==0) AfxMessageBox(tmp+" отправлено"); else AfxMessageBox(tmp+" не отправлено");
	}	

	//Stringto
}

void FControl::OnExport() 
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), "[ControlExcel]");
	CString csVal;
	
	pCmd.AddParameter("BDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_Begin.Format(VAR_DATEVALUEONLY));
	pCmd.AddParameter("EDate",CADORecordset::typeDate,
				CADOParameter::paramInput,12,m_End.Format(VAR_DATEVALUEONLY));
	
	csVal="%"+m_filterDeclarant;
	csVal+="%";
	pCmd.AddParameter("Decl",CADORecordset::typeChar,
				CADOParameter::paramInput,200,csVal.Left(199));
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

void FControl::GetSetForExport()
{
	m_bExportSendMessage=FALSE;
	OnExport();
	m_bExportSendMessage=TRUE;
}

void FControl::OnSelChangeLetterlist() 
{
	m_LetterList.SetRowSel(m_LetterList.GetRow());
	
}

void FControl::OnMouseMoveLetterlist(short Button, short Shift, long x, long y) 
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
