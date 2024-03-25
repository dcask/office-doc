// KanView.cpp : implementation of the CKanView class
//

#include "stdafx.h"
#include "Kan.h"

#include "KanDoc.h"
#include "KanView.h"
#include "FFind.h"
#include "FOutFind.h"
#include "FFilter.h"
#include "FOutFilter.h"
#include "FInMessage.h"
#include "FOutMessage.h"
#include "FStatistic.h"
#include "FPeriod.h"
#include "FChoosDate.h"
#include "FWait.h"
#include "FEdit.h"


#include "ChildFrm.h"
#include "msword.h"
#include "excel.h"
#include "FInbox.h"
#include "FOutbox.h"
#include "FControl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


#define ID_TIMER_1 100
/////////////////////////////////////////////////////////////////////////////
// CKanView

IMPLEMENT_DYNCREATE(CKanView, CFormView)

BEGIN_MESSAGE_MAP(CKanView, CFormView)
	//{{AFX_MSG_MAP(CKanView)
	ON_WM_SIZE()
	ON_WM_SHOWWINDOW()
	ON_NOTIFY(TCN_SELCHANGE, IDC_MAINTAB, OnSelchangeMaintab)
	ON_COMMAND(ID_LETTER_CREATE, OnLetterCreate)
	ON_COMMAND(ID_LETTER_DELETE, OnLetterDelete)
	ON_COMMAND(ID_LETTER_FIND, OnLetterFind)
	ON_COMMAND(ID_LETTER_FILTER, OnLetterFilter)
	ON_COMMAND(ID_REPORTS_INVALID, OnReportsInvalid)
	ON_COMMAND(ID_STAT_BYUSER, OnStatByuser)
	ON_COMMAND(ID_STAT_BYCONTROL, OnStatBycontrol)
	ON_COMMAND(ID_STAT_BYPERSON, OnStatByperson)
	ON_COMMAND(ID_STAT_BYAUTHOR, OnStatByauthor)
	ON_COMMAND(ID_HELP_COMPACT, OnHelpCompact)
	ON_COMMAND(ID_EDIT_QUERY, OnEditQuery)
	ON_COMMAND(ID_EDIT_MINE, OnEditMine)
	ON_COMMAND(ID_REPORTS_EXCEL, OnReportsExcel)
	ON_WM_TIMER()
	ON_WM_CREATE()
	ON_COMMAND(ID_STAT_ONLINE, OnStatOnline)
	ON_COMMAND(ID_LETTER_REFRESH, OnLetterRefresh)
	ON_COMMAND(ID_LETTER_UNFILTER, OnLetterUnfilter)
	ON_COMMAND(ID_REPORTS_SIGN, OnReportsSign)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CKanView construction/destruction

CKanView::CKanView()
	: CFormView(CKanView::IDD)
{

	CString csVal;
	m_inFilterDeclarant=_T("");
	m_outFilterAuthor=_T("");
	m_End = COleDateTime::GetCurrentTime();
	csVal="01/01/"+(COleDateTime::GetCurrentTime()).Format("%y");
	m_Begin.ParseDateTime(csVal);
	m_ToolTipIn=NULL;
	m_ToolTipOut=NULL;
	m_bTimer=FALSE;
}

CKanView::~CKanView()
{
	if(m_pImgListIn) delete m_pImgListIn;
	if(m_pImgListIn_d) delete m_pImgListIn_d;
	if(m_pImgListOut) delete m_pImgListOut;
	if(m_pImgListOut_d) delete m_pImgListOut_d;
	if(m_ToolTipIn) 
	{
		m_ToolTipIn->DestroyWindow();
		delete m_ToolTipIn;
	}
	if(m_ToolTipOut) 
	{
		m_ToolTipOut->DestroyWindow();
		delete m_ToolTipOut;
	}

}

void CKanView::DoDataExchange(CDataExchange* pDX)
{
	CFormView::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CKanView)
	DDX_Control(pDX, IDC_MAINTAB, m_Tab);
	//}}AFX_DATA_MAP
}

BOOL CKanView::PreCreateWindow(CREATESTRUCT& cs)
{
	return CFormView::PreCreateWindow(cs);
}

void CKanView::OnInitialUpdate()
{
	NMHDR hdr;
	
	TCITEM tcItem;
	CRect rect;
	SIZE sizeTotal;
	WINDOWPLACEMENT pm;

	GetDocument()->m_uiID=1;

	CFormView::OnInitialUpdate();
	CTabCtrl *tab=(CTabCtrl*)GetDlgItem(IDC_MAINTAB);
	int xpstyle=((CKanApp*) AfxGetApp())->m_iXPStyle;

	/*Оформление табов*/
	m_pTabImgList = new CImageList();
	ASSERT(m_pTabImgList!= NULL);    // serious allocation failure checking
	m_pTabImgList->Create(48, 144, ILC_COLOR32, 1, 1);
	if(xpstyle!=1) m_pTabImgList->SetBkColor(GetSysColor(COLOR_3DFACE));
	m_pTabImgList->Add(AfxGetApp()->LoadIcon(IDI_INBOX));
	m_pTabImgList->Add(AfxGetApp()->LoadIcon(IDI_OUTBOX));
	m_pTabImgList->Add(AfxGetApp()->LoadIcon(IDI_COMPLETE));
	
	tab->SetImageList(m_pTabImgList);

	/*Создание списков изображения для тулбаров*/
	m_pImgListIn=new CImageList();
	m_pImgListIn_d=new CImageList();
	m_pImgListOut=new CImageList();
	m_pImgListOut_d=new CImageList();

	m_pImgListIn->Create(96, 32, ILC_COLOR32, 4, 1);
	if(xpstyle!=1)	m_pImgListIn->SetBkColor(GetSysColor(COLOR_3DFACE)); 
	m_pImgListIn->Add(AfxGetApp()->LoadIcon(IDI_IN_CREATE));
	m_pImgListIn->Add(AfxGetApp()->LoadIcon(IDI_DELETE));
	m_pImgListIn->Add(AfxGetApp()->LoadIcon(IDI_SEARCH));
	m_pImgListIn->Add(AfxGetApp()->LoadIcon(IDI_FILTER));
	m_pImgListIn->Add(AfxGetApp()->LoadIcon(IDI_UNFIL));
	m_pImgListIn->Add(AfxGetApp()->LoadIcon(IDI_REFRESH));

	m_pImgListIn_d->Create(96, 32, ILC_COLOR32, 4, 1);
	if(xpstyle!=1)	m_pImgListIn_d->SetBkColor(GetSysColor(COLOR_3DFACE)); 
	m_pImgListIn_d->Add(AfxGetApp()->LoadIcon(IDI_IN_CREATE));
	m_pImgListIn_d->Add(AfxGetApp()->LoadIcon(IDI_DELETE));
	m_pImgListIn_d->Add(AfxGetApp()->LoadIcon(IDI_SEARCH));
	m_pImgListIn_d->Add(AfxGetApp()->LoadIcon(IDI_FILTER));
	m_pImgListIn_d->Add(AfxGetApp()->LoadIcon(IDI_UNFIL));
	m_pImgListIn_d->Add(AfxGetApp()->LoadIcon(IDI_REFRESH));

	m_pImgListOut->Create(96, 32, ILC_COLOR32, 4, 1);
	if(xpstyle!=1)	m_pImgListOut->SetBkColor(GetSysColor(COLOR_3DFACE)); 
	m_pImgListOut->Add(AfxGetApp()->LoadIcon(IDI_OUT_CREATE));
	m_pImgListOut->Add(AfxGetApp()->LoadIcon(IDI_DELETE));
	m_pImgListOut->Add(AfxGetApp()->LoadIcon(IDI_SEARCH));
	m_pImgListOut->Add(AfxGetApp()->LoadIcon(IDI_FILTER));
	m_pImgListOut->Add(AfxGetApp()->LoadIcon(IDI_UNFIL));
	m_pImgListOut->Add(AfxGetApp()->LoadIcon(IDI_REFRESH));

	m_pImgListOut_d->Create(96, 32, ILC_COLOR32, 4, 1);
	if(xpstyle!=1)	m_pImgListOut_d->SetBkColor(GetSysColor(COLOR_3DFACE)); 

	m_pImgListOut_d->Add(AfxGetApp()->LoadIcon(IDI_OUT_CREATE));
	m_pImgListOut_d->Add(AfxGetApp()->LoadIcon(IDI_DELETE));
	m_pImgListOut_d->Add(AfxGetApp()->LoadIcon(IDI_SEARCH));
	m_pImgListOut_d->Add(AfxGetApp()->LoadIcon(IDI_FILTER));
	m_pImgListOut_d->Add(AfxGetApp()->LoadIcon(IDI_UNFIL));
	m_pImgListOut_d->Add(AfxGetApp()->LoadIcon(IDI_REFRESH));
	
	hdr.code = TCN_SELCHANGE;
	hdr.hwndFrom = tab->m_hWnd;
	
	tcItem.mask = TCIF_TEXT | TCIF_IMAGE   ;
	tcItem.pszText = _T("");tcItem.iImage=0;tab->InsertItem(0, &tcItem);
	tcItem.pszText = _T("");tcItem.iImage=1;tab->InsertItem(1, &tcItem);
	tcItem.pszText = _T("");tcItem.iImage=2;tab->InsertItem(2, &tcItem);
	GetClientRect(&rect);
	
	sizeTotal.cx=rect.right;// размеры области для прокрутки внутри дочернего окна
	sizeTotal.cy=rect.bottom;//
	m_iCx=rect.right;// размеры области для прокрутки внутри дочернего окна
	m_iCy=rect.bottom;//
	
	SetScrollSizes(MM_TEXT,sizeTotal); 
	CSize size;
	size.cx=2;
	size.cy=2;

	tab->GetWindowPlacement(&pm);
	pm.rcNormalPosition.bottom=rect.bottom-10;
	pm.rcNormalPosition.right=rect.right-10;
	tab->SetWindowPlacement(&pm);
	tab->SetCurSel(0);
	tab->SetPadding(size);
	size.cx=10;
	size.cy=55;
	tab->SetItemSize(size);
	
	SendMessage ( WM_NOTIFY, tab->GetDlgCtrlID(), (LPARAM)&hdr );
}

/////////////////////////////////////////////////////////////////////////////
// CKanView diagnostics

#ifdef _DEBUG
void CKanView::AssertValid() const
{
	CFormView::AssertValid();
}

void CKanView::Dump(CDumpContext& dc) const
{								
	CFormView::Dump(dc);
}

CKanDoc* CKanView::GetDocument() // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CKanDoc)));
	return (CKanDoc*)m_pDocument;
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CKanView message handlers

void CKanView::OnSize(UINT nType, int cx, int cy) 
{
	CFormView::OnSize(nType, cx, cy);
	CKanDoc* doc=GetDocument();

	WINDOWPLACEMENT pm;
	CRect rc; 

	CTabCtrl *tab=(CTabCtrl*)GetDlgItem(IDC_MAINTAB);
	if(tab)
	{
		tab->GetWindowPlacement(&pm);
		pm.rcNormalPosition.bottom+=cy-m_iCy;
		pm.rcNormalPosition.right+=cx-m_iCx;
		tab->SetWindowPlacement(&pm);
		tab->GetWindowRect (&rc); // получаем "рабочую область"
		tab->ScreenToClient (&rc); // преобразуем в относительные координаты
		tab->AdjustRect (FALSE, &rc); 
		if(doc->m_pTabDialog)
			doc->m_pTabDialog->MoveWindow (&rc);
		m_iCx=cx;	m_iCy=cy;
	}
}


void CKanView::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CFormView::OnShowWindow(bShow, nStatus);
	CString csVal;
	CChildFrame* pChild= (CChildFrame*) (this->GetOwner());
	pChild->m_iToolBar=IDR_IN;
}

void CKanView::OnSelchangeMaintab(NMHDR* pNMHDR, LRESULT* pResult) 
{
	int id; // ID диалога
	CKanDoc* doc=GetDocument();
	CChildFrame* pChild= (CChildFrame*) (this->GetOwner());
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	COleDateTimeSpan ts =GetPrivateProfileInt("Control","Period",5, csAppPath+"\\kan.ini");;
	// надо сначала удалить предыдущий диалог в Tab Control'е:
	if (doc->m_pTabDialog)
	{
		doc->m_pTabDialog->DestroyWindow();
		delete doc->m_pTabDialog;
		doc->m_pTabDialog=NULL;
	}

	// теперь в зависимости от того, какая закладка выбрана, 
	// выбираем соотв. диалог
	switch( m_Tab.GetCurSel()+1 ) // +1 для того, чтобы порядковые номера закладок
	// и диалогов совпадали с номерами в case
	{
	// первая закладка
		case 1 :
			id = IDD_INBOX;
			doc->m_pTabDialog = new FInbox;
			((FInbox*)(doc->m_pTabDialog))->doc=doc;
			GetPeriod("Inbox", ((FInbox*)(doc->m_pTabDialog))->m_Begin,
				((FInbox*)(doc->m_pTabDialog))->m_End);
			LoadToolBar(1,&(pChild->m_wndToolBar));
			break;
			// вторая закладка
		case 2 :
			id = IDD_OUTBOX;
			doc->m_pTabDialog = new FOutbox;
			((FOutbox*)(doc->m_pTabDialog))->doc=doc;
			GetPeriod("OutBox", ((FOutbox*)(doc->m_pTabDialog))->m_Begin,
				((FOutbox*)(doc->m_pTabDialog))->m_End);
			LoadToolBar(2,&(pChild->m_wndToolBar));
			break;
		case 3 :
			id = IDD_CONTROL;
			doc->m_pTabDialog = new FControl;
			((FControl*)(doc->m_pTabDialog))->doc=doc;
			((FControl*)(doc->m_pTabDialog))->m_Begin=COleDateTime::GetCurrentTime();
			((FControl*)(doc->m_pTabDialog))->m_End=((FControl*)(doc->m_pTabDialog))->m_Begin;
			((FControl*)(doc->m_pTabDialog))->m_End+=ts;
			LoadToolBar(3,&(pChild->m_wndToolBar));
			break;
			break;
		default:
			doc->m_pTabDialog = new CDialog; // новый пустой диалог
			return;

	} // end switch

	// создаем диалог

	doc->m_pTabDialog->Create (id, (CWnd*)&m_Tab); //параметры: ресурс диалога и родитель

	CRect rc; 
	doc->m_uiCurSelectedTab=m_Tab.GetCurSel();

	m_Tab.GetWindowRect (&rc); // получаем "рабочую область"
	m_Tab.ScreenToClient (&rc); // преобразуем в относительные координаты

	// исключаем область, где отображаются названия закладок:
	m_Tab.AdjustRect (FALSE, &rc); 

	// помещаем диалог на место..
	doc->m_pTabDialog->MoveWindow (&rc);

	// и показываем:
	doc->m_pTabDialog->ShowWindow ( SW_SHOWNORMAL );
	doc->m_pTabDialog->UpdateWindow();
	doc->m_pTabDialog->SetFocus();
	*pResult = 0;
}

BOOL CKanView::LoadToolBar(UINT id, CToolBar* tb)
{
	unsigned int buttons[6];
	
	switch(id)
	{
		case 1: // inbox
			buttons[0]=ID_LETTER_CREATE;
			buttons[1]=ID_LETTER_DELETE;
			buttons[2]=ID_LETTER_FIND;
			buttons[3]=ID_LETTER_FILTER;
			buttons[4]=ID_LETTER_UNFILTER;
			buttons[5]=ID_LETTER_REFRESH;
			tb->GetToolBarCtrl().SetImageList(m_pImgListIn);
			tb->GetToolBarCtrl().SetDisabledImageList(m_pImgListIn_d);
			tb->SetButtons(buttons,6);
			tb->ShowWindow(SW_SHOW);
			break;
		case 2: // outbox
			buttons[0]=ID_LETTER_CREATE;
			buttons[1]=ID_LETTER_DELETE;
			buttons[2]=ID_LETTER_FIND;
			buttons[3]=ID_LETTER_FILTER;
			buttons[4]=ID_LETTER_UNFILTER;
			buttons[5]=ID_LETTER_REFRESH;
			tb->GetToolBarCtrl().SetImageList(m_pImgListOut);
			tb->GetToolBarCtrl().SetDisabledImageList(m_pImgListOut_d);
			tb->SetButtons(buttons,6);
			tb->ShowWindow(SW_SHOW);
			break;
		case 3: // Control
			buttons[0]=ID_LETTER_CREATE;
			buttons[1]=ID_LETTER_DELETE;
			buttons[2]=ID_LETTER_FIND;
			buttons[3]=ID_LETTER_FILTER;
			buttons[4]=ID_LETTER_UNFILTER;
			buttons[5]=ID_LETTER_REFRESH;
			tb->GetToolBarCtrl().SetImageList(m_pImgListIn_d);
			tb->GetToolBarCtrl().SetDisabledImageList(m_pImgListIn_d);
			tb->SetButtons(buttons,6);
			break;
		default:
			break;
	}
	return TRUE;	
}

void CKanView::OnActivateView(BOOL bActivate, CView* pActivateView, CView* pDeactiveView) 
{
	if(!bActivate) this->GetParent()->SetWindowPos(&CWnd::wndBottom, 0, 0, 0, 0,
      SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE|SWP_NOSENDCHANGING);
	
	CFormView::OnActivateView(bActivate, pActivateView, pDeactiveView);
}

void CKanView::OnLetterCreate() 
{
	FInMessage inDlg;
	FOutMessage outDlg;
	double Num,NumTmp;
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmdIn(bt->GetDB(), "[InMax]");
	CADOCommand pCmdOut(bt->GetDB(), "[OutMax]");
	CADOCommand pCmdTmpMax(bt->GetDB(), "[TempMax]");
	CADOCommand pCmdTmpNum(bt->GetDB(), "[AddTempNum]");
	CADOCommand pCmdTmpDel(bt->GetDB(), "[DelTempNum]");
	COleDateTime date;
	date.ParseDateTime("01/01/"+COleDateTime::GetCurrentTime().Format("%Y"));
	pCmdIn.AddParameter("BDate",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),date);
	pCmdOut.AddParameter("BDate",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),date);
	date.ParseDateTime("31/12/"+COleDateTime::GetCurrentTime().Format("%Y"));
	pCmdIn.AddParameter("EDate",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),date);
	pCmdOut.AddParameter("EDate",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),date);
	switch(m_Tab.GetCurSel()+1)
	{
		case 1://inbox
				if(!((CKanApp*)AfxGetApp())->m_bEditIn) return;
				inDlg.doc=GetDocument();
				inDlg.doc->m_lLetterId=0;
				inDlg.doc->LoadInData();
				/*максимальный во временных входящих*/
				pCmdTmpMax.AddParameter("Type",CADORecordset::typeInteger,
					CADOParameter::paramInput,sizeof(int),0);
				try
				{
					bt->GetRS()->Execute(&pCmdTmpMax);
				}catch(CADOException* e)
				{
					bt->ThrowError(e);
					return;
				}
				bt->GetRS()->MoveFirst();
				bt->GetRS()->GetFieldValue("Num",NumTmp);
				bt->GetRS()->Close();
				
				/*максимальный во входящих*/
				try
				{
					bt->GetRS()->Execute(&pCmdIn);
				}catch(CADOException* e)
				{
					bt->ThrowError(e);
					return;
				}
				bt->GetRS()->MoveFirst();
				bt->GetRS()->GetFieldValue("Num",Num);
				bt->GetRS()->Close();
				/*выбор большего номера, дабы предотвратить пересечение*/
				if(NumTmp>Num) Num=NumTmp;
				inDlg.m_csRegNum.Format("%d",int(Num));
				/*внесение номера во временное хранилище*/
				pCmdTmpNum.AddParameter("Num",CADORecordset::typeInteger,
					CADOParameter::paramInput,sizeof(int),int(Num));
				pCmdTmpNum.AddParameter("Type",CADORecordset::typeInteger,
					CADOParameter::paramInput,sizeof(int),0);
				pCmdTmpNum.AddParameter("User",CADORecordset::typeInteger,
					CADOParameter::paramInput,sizeof(int),((CKanApp*)AfxGetApp())->m_iUserID);
				try
				{
					bt->GetRS()->Execute(&pCmdTmpNum);
				}catch(CADOException* e)
				{
					bt->ThrowError(e);
					return;
				}
				bt->GetRS()->Close();
				/*диалог*/
				inDlg.m_bNew=TRUE;
				if(inDlg.DoModal()==IDOK)
				{
					/*Добавляем новый*/
					bt->GetRS()->Open(((CKanApp*)AfxGetApp())->m_csInLetter,CADORecordset::openTable);
					bt->GetRS()->AddNew();
					bt->GetRS()->SetFieldValue(1,0);
					bt->GetRS()->Update();
					bt->GetRS()->GetFieldValue("ID",inDlg.doc->m_lLetterId);
					inDlg.doc->SaveInData();
					((FInbox*)(GetDocument()->m_pTabDialog))->Show();
				}
				/*удаляем временный*/
				try
				{
					pCmdTmpDel.AddParameter("Num",CADORecordset::typeInteger,
						CADOParameter::paramInput,sizeof(int),int(Num));
					bt->GetRS()->Execute(&pCmdTmpDel);
				}catch(CADOException* e)
				{
					bt->ThrowError(e);
					return;
				}
				bt->GetRS()->Close();
				break;

		case 2://outbox
				if(!((CKanApp*)AfxGetApp())->m_bEditOut) return;
				outDlg.doc=GetDocument();
				outDlg.doc->m_lLetterId=0; 
				outDlg.doc->LoadOutData(); // подготовка данных
				/*максимальный во временных исходящих*/
				pCmdTmpMax.AddParameter("Type",CADORecordset::typeInteger,
					CADOParameter::paramInput,sizeof(int),1);
				try
				{
					bt->GetRS()->Execute(&pCmdTmpMax);
				}catch(CADOException* e)
				{
					bt->ThrowError(e);
					return;
				}
				bt->GetRS()->MoveFirst();
				bt->GetRS()->GetFieldValue("Num",NumTmp);
				bt->GetRS()->Close();
				/*максимальный в исходящих*/
				try
				{
					bt->GetRS()->Execute(&pCmdOut);
				}catch(CADOException* e)
				{
					bt->ThrowError(e);
					return;
				}
				bt->GetRS()->MoveFirst();
				bt->GetRS()->GetFieldValue("Num",Num);
				outDlg.doc->m_lInLetterID=0;	// нет связывания для нового
				bt->GetRS()->Close();
				/*выбор большего номера, дабы предотвратить пересечение*/
				if(NumTmp>Num) Num=NumTmp;
				outDlg.m_csRegNum.Format("%d",int(Num));
				/*внесение номера во временное хранилище*/
				pCmdTmpNum.AddParameter("Num",CADORecordset::typeInteger,
					CADOParameter::paramInput,sizeof(int),int(Num));
				pCmdTmpNum.AddParameter("Type",CADORecordset::typeInteger,
					CADOParameter::paramInput,sizeof(int),1);
				pCmdTmpNum.AddParameter("User",CADORecordset::typeInteger,
					CADOParameter::paramInput,sizeof(int),((CKanApp*)AfxGetApp())->m_iUserID);
				try
				{
					bt->GetRS()->Execute(&pCmdTmpNum);
				}catch(CADOException* e)
				{
					bt->ThrowError(e);
					return;
				}
				bt->GetRS()->Close();
				/*диалог*/
				outDlg.m_bNew=TRUE;
				if(outDlg.DoModal()==IDOK)
				{
					/*Добавляем новый*/
					bt->GetRS()->Open(((CKanApp*)AfxGetApp())->m_csOutLetter,CADORecordset::openTable);
					bt->GetRS()->AddNew();
					bt->GetRS()->SetFieldValue(1,0);
					bt->GetRS()->Update();
					bt->GetRS()->GetFieldValue("ID",outDlg.doc->m_lLetterId);
					outDlg.doc->SaveOutData();
					((FOutbox*)(GetDocument()->m_pTabDialog))->Show();
				}
				/*удаляем временный*/
				try
				{
					pCmdTmpDel.AddParameter("Num",CADORecordset::typeInteger,
						CADOParameter::paramInput,sizeof(int),int(Num));
					bt->GetRS()->Execute(&pCmdTmpDel);
				}catch(CADOException* e)
				{
					bt->ThrowError(e);
					return;
				}
				bt->GetRS()->Close();
				break;
		default:break;
	}
	
}

void CKanView::OnLetterDelete() 
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmdIn(bt->GetDB(), "[InDelete]");
	CADOCommand pCmdInNum(bt->GetDB(), "[DelInNum]");
	CADOCommand pCmdOutNum(bt->GetDB(), "[DelOutNum]");
	CADOCommand pCmdEx(bt->GetDB(), "[ExDelete]");
	CADOCommand pCmdOut(bt->GetDB(), "[OutDelete]");
	CKanDoc* doc=GetDocument();
	long row;
	CString csVal;
	switch(m_Tab.GetCurSel()+1)
	{
		case 1://inbox
				if(!((CKanApp*)AfxGetApp())->m_bEditIn) return;
				if(AfxMessageBox("Удалить выбраное письмо ?", MB_YESNO)==IDNO) return;
				row=((FInbox*)(doc->m_pTabDialog))->m_LetterList.GetRow();
				if(row<1) return; // если список пуст
				csVal=((FInbox*)(doc->m_pTabDialog))->m_LetterList.GetTextMatrix(row,1);
				pCmdIn.AddParameter("InID",CADORecordset::typeBigInt,
						CADOParameter::paramInput,sizeof(long),atol(csVal));
				pCmdEx.AddParameter("InID",CADORecordset::typeBigInt,
						CADOParameter::paramInput,sizeof(long),atol(csVal));
				pCmdEx.AddParameter("ExID",CADORecordset::typeBigInt,
						CADOParameter::paramInput,sizeof(long),-1);
				pCmdInNum.AddParameter("_ID",CADORecordset::typeBigInt,
						CADOParameter::paramInput,sizeof(long),long(-1));
				pCmdInNum.AddParameter("InID",CADORecordset::typeBigInt,
						CADOParameter::paramInput,sizeof(long),atol(csVal));
				bt->Execute(&pCmdIn);
				bt->Execute(&pCmdEx);
				bt->Execute(&pCmdInNum);
				bt->GetRS()->Close();
				
				((FInbox*)(GetDocument()->m_pTabDialog))->Show();
				break;
				
		case 2://outbox
				if(!((CKanApp*)AfxGetApp())->m_bEditOut) return;
				if(AfxMessageBox("Удалить выбраное письмо ?", MB_YESNO)==IDNO) return;
				row=((FOutbox*)(doc->m_pTabDialog))->m_LetterList.GetRow();
				if(row<1) return; // если список пуст
				csVal=((FOutbox*)(doc->m_pTabDialog))->m_LetterList.GetTextMatrix(row,1);
				pCmdOut.AddParameter("OutID",CADORecordset::typeBigInt,
						CADOParameter::paramInput,sizeof(long),atol(csVal));
				pCmdInNum.AddParameter("_ID",CADORecordset::typeBigInt,
						CADOParameter::paramInput,sizeof(long),long(-1));
				pCmdInNum.AddParameter("OutID",CADORecordset::typeBigInt,
						CADOParameter::paramInput,sizeof(long),atol(csVal));
				bt->Execute(&pCmdOut);
				bt->Execute(&pCmdOutNum);
				bt->GetRS()->Close();
				((FOutbox*)(GetDocument()->m_pTabDialog))->Show();
				break;
		default:break;
	}
	
}

void CKanView::OnLetterFind() 
{
	FFind inFind;
	FOutFind outFind;
	FInMessage inDlg;
	FOutMessage outDlg;
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmdIn(bt->GetDB(), "[FindIn]");
	CADOCommand pCmdInNum(bt->GetDB(), "[FindInNum]");
	CADOCommand pCmdOut(bt->GetDB(), "[FindOut]");


	CString csVal;
	switch(m_Tab.GetCurSel()+1)
	{
		case 1://inbox
				if(inFind.DoModal()==IDOK)
				{
					if(inFind.m_csExtInNum=="")
					{
						inFind.m_csRegNum.TrimLeft();inFind.m_csRegNum.TrimRight();
						pCmdIn.AddParameter("InRegNum",CADORecordset::typeChar,
								CADOParameter::paramInput,50,inFind.m_csRegNum);
						pCmdIn.AddParameter("RegDateYear",CADORecordset::typeUnsignedInt,
								CADOParameter::paramInput,sizeof(unsigned int),int(inFind.m_uiYear));
						bt->Execute(&pCmdIn);
					}
					if(inFind.m_csRegNum=="")
					{
						inFind.m_csDeclarant.TrimLeft();inFind.m_csDeclarant.TrimRight();
						csVal="%"+inFind.m_csDeclarant.Left(199)+"%";
						inFind.m_csExtInNum.TrimLeft();inFind.m_csExtInNum.TrimRight();
						pCmdInNum.AddParameter("RegNum",CADORecordset::typeChar,
								CADOParameter::paramInput,50,inFind.m_csExtInNum);
						pCmdInNum.AddParameter("Decl",CADORecordset::typeChar,
								CADOParameter::paramInput,202,csVal);
						pCmdInNum.AddParameter("_Year",CADORecordset::typeUnsignedInt,
								CADOParameter::paramInput,sizeof(unsigned int),int(inFind.m_uiYear));
						bt->Execute(&pCmdInNum);
					}
					
					
					if(bt->GetRecordsCount()<1)
					{
						AfxMessageBox("Не найден");
						break;
					}
					if(bt->GetRecordsCount()>1)
					{
						AfxMessageBox("Несколько вариантов. Просмотр первого найденного");
						//break;
					}
					bt->GetRS()->MoveFirst();
					bt->GetRS()->GetFieldValue("ID",GetDocument()->m_lLetterId);
					bt->GetRS()->Close();
					inDlg.doc=GetDocument();
					inDlg.doc->LoadInData();
					if(inDlg.DoModal()==IDOK)
					{
						inDlg.doc->SaveInData();
						((FInbox*)(GetDocument()->m_pTabDialog))->Show();
					}
				}
				break;
			case 2://outbox
				if(outFind.DoModal()==IDOK)
				{
					if(outFind.m_csRegNum=="") outFind.m_csRegNum=((CKanApp*)AfxGetApp())->m_TemplateSymbol;
					pCmdOut.AddParameter("OutRegNum",CADORecordset::typeChar,
							CADOParameter::paramInput,50,outFind.m_csRegNum);
					pCmdOut.AddParameter("RegDateYear",CADORecordset::typeUnsignedInt,
							CADOParameter::paramInput,sizeof(unsigned int),int(outFind.m_uiYear));
					try
					{
						bt->GetRS()->Execute(&pCmdOut);
					}catch(CADOException* e)
					{
						bt->ThrowError(e);
						return;
					}
					if(bt->GetRecordsCount()<1)
					{
						AfxMessageBox("Не найден");
						break;
					}
					if(bt->GetRecordsCount()>1)
					{
						AfxMessageBox("Несколько вариантов. Просмотр первого найденного");
						//break;
					}
					bt->GetRS()->MoveFirst();
					bt->GetRS()->GetFieldValue("ID",GetDocument()->m_lLetterId);
					bt->GetRS()->Close();
					outDlg.doc=GetDocument();
					outDlg.doc->LoadOutData();
					if(outDlg.DoModal()==IDOK)
					{
						outDlg.doc->SaveOutData();
						((FOutbox*)(GetDocument()->m_pTabDialog))->Show();
					}
				}
				break;
		default:break;
	}
}
	
void CKanView::OnLetterFilter() 
{
	FFilter inFilter;
	FOutFilter outFilter;
	CKanDoc* doc=GetDocument();
	inFilter.m_csDeclarant=m_inFilterDeclarant;
	outFilter.m_csAuthor=m_outFilterAuthor;
	inFilter.m_csContent=m_inFilterContent;
	outFilter.m_csContent=m_outFilterContent;
	outFilter.m_csReciever=m_outFilterReciever;
	switch(m_Tab.GetCurSel()+1)
	{
		case 1://inbox
			if(inFilter.DoModal()==IDOK)
			{
				m_inFilterDeclarant=inFilter.m_csDeclarant;
				m_inFilterContent=inFilter.m_csContent;
				((FInbox*)(doc->m_pTabDialog))->m_inFilterDeclarant=inFilter.m_csDeclarant;
				((FInbox*)(doc->m_pTabDialog))->m_inFilterContent=inFilter.m_csContent;
				((FInbox*)(doc->m_pTabDialog))->Show();
			}
			break;


		case 2://outbox
			if(outFilter.DoModal()==IDOK)
			{
				m_outFilterAuthor=outFilter.m_csAuthor;
				m_outFilterContent=outFilter.m_csContent;
				m_outFilterReciever=outFilter.m_csReciever;
				((FOutbox*)(doc->m_pTabDialog))->m_outFilterAuthor=outFilter.m_csAuthor;
				((FOutbox*)(doc->m_pTabDialog))->m_outFilterContent=outFilter.m_csContent;
				((FOutbox*)(doc->m_pTabDialog))->m_outFilterReciever=outFilter.m_csReciever;
				((FOutbox*)(doc->m_pTabDialog))->Show();
			}
			break;
		default:break;
	}

	
}

void CKanView::OnReportsInvalid() 
{
	if(!((CKanApp*)AfxGetApp())->m_bReport) return;
	_EApplication oApp;
	Workbooks oBooks; //объявляем коллекцию "рабочих книг"
	_Workbook oBook; //объявляем конкретный экземпляр (класс _Workbook)
	Worksheets oSheets; //объявляем коллекцию "рабочих листов"
	_Worksheet oSheet; //объявляем конкретный экземпляр (класс _Worksheet)
	Range oRange; //объявляем экземпляр класса Range, отвечающего за заполнение ячеек
	Interior oInterior;
	CString col,num;
	COleVariant x,y,res;
	CString query;
	CString formula,resformula;
	CString name;
	unsigned short firstLetter,secondLetter;
	unsigned long nFields;

	COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant  covTrue(short(1), VT_BOOL);
	COleVariant  covFalse(short(0), VT_BOOL);
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), "[ReportBreach]");
	FPeriod dlg;
	try{
		if(dlg.DoModal()!=IDOK) return;

		pCmd.AddParameter("BDate",CADORecordset::typeDate,
					CADOParameter::paramInput,12,dlg.m_Begin.Format(VAR_DATEVALUEONLY));
		pCmd.AddParameter("EDate",CADORecordset::typeDate,
					CADOParameter::paramInput,12,dlg.m_End.Format(VAR_DATEVALUEONLY));
		
		bt->Execute(&pCmd);
		if(!oApp.CreateDispatch("Excel.Application")) //запустить сервер
		{
			AfxMessageBox("Невозможно запустить Excel.");
			return;
		}

		oBooks = oApp.GetWorkbooks(); // и получаем список книг
		oBook = oBooks.Add(covOptional);
		
		oSheets = oBook.GetWorksheets(); //и получаем список листов
		
		oSheet =oSheets.GetItem(COleVariant(long(1)));

		FWait* waitdlg;
		waitdlg=new FWait();
		waitdlg->Create(FWait::IDD);
		waitdlg->ShowWindow(SW_SHOW);
		waitdlg->SetInfo("Формирование документа...");

		long counter=0,row=0,rows;
		
			rows=bt->GetRS()->GetRecordCount();

			if(bt->GetRS()->IsOpen())
			{
				nFields=bt->GetRS()->GetFieldCount();

				/*проверяем количество записей результате*/
				if(rows<1) rows=0;		
				else bt->GetRS()->MoveFirst();
		
				CADOFieldInfo info;
				/*список полей в шапку*/
				
				for(unsigned long i=0; i<nFields; ++i)
				{
					firstLetter=unsigned short (i/26.0);
					secondLetter=unsigned short(i-firstLetter*26);
					if(firstLetter<1) col.Format("%c1",secondLetter+'A');
					else col.Format("%c%c1",firstLetter+'A'-1,secondLetter+'A');
					y=COleVariant(col);
					oRange = oSheet.GetRange(y,y); //например

					bt->GetRS()->GetFieldInfo(i,&info);
					oRange.SetValue(covOptional,COleVariant(info.m_strName));
					oRange.BorderAround(COleVariant(short(2)),1,-4105,COleVariant(long(-4105)));
					//oRange.SetShrinkToFit(covTrue);
				}
				/*данные*/
				row=2;
				while(!bt->GetRS()->IsEof()&&rows>0)
				{
					for(unsigned long i=0; i<nFields; ++i)
					{
						firstLetter=unsigned short (i/26.0);
						secondLetter=unsigned short(i-firstLetter*26);
						if(firstLetter<1) col.Format("%c%d",secondLetter+'A',row);
						else col.Format("%c%c%d",firstLetter+'A'-1,secondLetter+'A',row);
						y=COleVariant(col);
						oRange = oSheet.GetRange(y,y); //например
						oRange.SetValue(covOptional,COleVariant(bt->GetStringValue(i)));
						oRange.BorderAround(COleVariant(short(2)),1,-4105,COleVariant(long(-4105)));
						//oRange.SetShrinkToFit(covTrue);
					}
					col.Format("Формирование... %d строк(%d",row,int(row*100/rows));
					col+="%)";
					waitdlg->SetInfo(col);
					bt->GetRS()->MoveNext(); row++;
				}
			}
		waitdlg->DestroyWindow(); delete waitdlg;
		x=COleVariant("A1");
		firstLetter=unsigned short ((nFields-1)/26.0);
		secondLetter=unsigned short((nFields-1)-firstLetter*26);
		if(firstLetter<1) col.Format("%c%d",secondLetter+'A',rows+1);
		else col.Format("%c%c%d",firstLetter+'A'-1,secondLetter+'A',rows+1);
		y=COleVariant(col);
		oRange = oSheet.GetRange(x,y);
		oRange.BorderAround(COleVariant(short(1)),2,-4105,COleVariant(long(-4105)));
		//oRange.SetShrinkToFit(covTrue);
		oRange=oRange.GetEntireColumn();
		oRange.AutoFit();
		if(firstLetter<1) col.Format("%c1",secondLetter+'A');
		else col.Format("%c%c1",firstLetter+'A'-1,secondLetter+'A');
		y=COleVariant(col);
		oRange = oSheet.GetRange(x,y);
		oInterior=oRange.GetInterior();
		oInterior.SetColorIndex(COleVariant(long(15)));
		oSheet.Activate();
		oApp.SetVisible(TRUE);
		oApp.SetUserControl(TRUE);
		bt->GetRS()->Close();	
	}//end try
	catch(CADOException* e)
	{
		AfxMessageBox(e->GetErrorMessage(),MB_OK|MB_ICONERROR);
	}
	catch(...)
	{
		AfxMessageBox(bt->GetDB()->GetErrorDescription(),MB_OK|MB_ICONERROR);
	}
}

void CKanView::OnStatByuser() 
{
	if(!((CKanApp*)AfxGetApp())->m_bState) return;
	FStatistic dlg;
	FPeriod period;
	if(period.DoModal()!=IDOK) return;
	dlg.m_Begin=period.m_Begin;
	dlg.m_End=period.m_End;
	dlg.m_csQuery="[GetInByUser]";
	dlg.m_csTitle="Количество писем внесённых пользователями программы";
	dlg.DoModal();
	
}

void CKanView::OnStatBycontrol() 
{
	if(!((CKanApp*)AfxGetApp())->m_bState) return;
	FStatistic dlg;
	FPeriod period;
	if(period.DoModal()!=IDOK) return;
	dlg.m_Begin=period.m_Begin;
	dlg.m_End=period.m_End;
	dlg.m_csQuery="[GetInByControl]";
	dlg.m_csTitle="Количество писем по видам контроля";
	dlg.DoModal();
	
}

void CKanView::GetPeriod(CString name, COleDateTime &begin, COleDateTime &end)
{
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	CString csVal;
	COleDateTimeSpan ts=GetPrivateProfileInt(name,"Period",30,csAppPath+"\\kan.ini");
	end=COleDateTime::GetCurrentTime();
	begin=end;
	begin-=ts;

}

void CKanView::OnStatByperson() 
{
	if(!((CKanApp*)AfxGetApp())->m_bState) return;
	FStatistic dlg;
	FPeriod period;
	if(period.DoModal()!=IDOK) return;
	dlg.m_Begin=period.m_Begin;
	dlg.m_End=period.m_End;
	dlg.m_csQuery="[GetInByPerson]";
	dlg.m_csTitle="Количество писем переданных исполнителям";
	dlg.DoModal();
	
}

void CKanView::OnStatByauthor() 
{
	if(!((CKanApp*)AfxGetApp())->m_bState) return;
	FStatistic dlg;
	FPeriod period;
	if(period.DoModal()!=IDOK) return;
	dlg.m_Begin=period.m_Begin;
	dlg.m_End=period.m_End;
	dlg.m_csQuery="[GetOutByAuthor]";
	dlg.m_csTitle="Количество писем от исполнителей";
	dlg.DoModal();
	
}

void CKanView::OnHelpCompact() 
{
	if(!((CKanApp*)AfxGetApp())->m_bBaseTool) return;
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	if(AfxMessageBox("Упаковать базу?", MB_YESNO | MB_ICONQUESTION)==IDYES) 
	{
		FWait* waitdlg=new FWait();
		waitdlg->Create(FWait::IDD);
		waitdlg->ShowWindow(SW_SHOW);
		waitdlg->SetInfo("Сжатие...");

		bt->Compact();

		waitdlg->DestroyWindow(); delete waitdlg;
	}
	
}

void CKanView::OnEditQuery() 
{
	if (((CKanApp*)AfxGetApp())->m_bBaseTool)
	{
		FEdit dlg;
		dlg.DoModal();
	}
	
}

void CKanView::OnEditMine() 
{
	NMHDR hdr;

	((CKanApp*) AfxGetApp())->m_bOnlyMine=!(((CKanApp*) AfxGetApp())->m_bOnlyMine);
	CMenu* menu=AfxGetApp()->GetMainWnd()->GetMenu();
	CMenu* submenu=menu->GetSubMenu(4);
	UINT state = submenu->GetMenuState(ID_EDIT_MINE, MF_BYCOMMAND);
	if (state & MF_CHECKED)	submenu->CheckMenuItem(ID_EDIT_MINE, MF_UNCHECKED | MF_BYCOMMAND);
	else submenu->CheckMenuItem(ID_EDIT_MINE, MF_CHECKED | MF_BYCOMMAND);

	CTabCtrl *tab=(CTabCtrl*)GetDlgItem(IDC_MAINTAB);
	hdr.code = TCN_SELCHANGE;
	hdr.hwndFrom = tab->m_hWnd;
	SendMessage ( WM_NOTIFY, tab->GetDlgCtrlID(), (LPARAM)&hdr );
}

BOOL CKanView::PreTranslateMessage(MSG* pMsg) 
{
	if (NULL != m_ToolTipIn)            
      m_ToolTipIn->RelayEvent(pMsg);
	switch(pMsg->message)
	{
		case WM_NOTIFY:
			if(pMsg->wParam==9999)
			{
				StartSession();
			}
			break;
		default:break;
	}
	return CFormView::PreTranslateMessage(pMsg);
}

void CKanView::OnReportsExcel() 
{
	if(!((CKanApp*)AfxGetApp())->m_bReport) return;
	_EApplication oApp;
	Workbooks oBooks; //объявляем коллекцию "рабочих книг"
	_Workbook oBook; //объявляем конкретный экземпляр (класс _Workbook)
	Worksheets oSheets; //объявляем коллекцию "рабочих листов"
	_Worksheet oSheet; //объявляем конкретный экземпляр (класс _Worksheet)
	Range oRange; //объявляем экземпляр класса Range, отвечающего за заполнение ячеек
	Interior oInterior;
	CString col,num;
	COleVariant x,y,res;
	CString query;
	CString formula,resformula;
	CString name;
	unsigned short firstLetter,secondLetter;
	unsigned long nFields;

	COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant  covTrue(short(1), VT_BOOL);
	COleVariant  covFalse(short(0), VT_BOOL);
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CKanDoc* doc=GetDocument();
	try{
		if(!bt->GetRS()->IsOpen())
		{
			switch(m_Tab.GetCurSel()+1)
			{
				case 1://inbox
					((FInbox*)(doc->m_pTabDialog))->GetSetForExport();
					break;
				case 2://outbox
					((FOutbox*)(doc->m_pTabDialog))->GetSetForExport();
					break;
				case 3://control
					((FControl*)(doc->m_pTabDialog))->GetSetForExport();
					break;
				default:
					break;
			}
		}
		if(!oApp.CreateDispatch("Excel.Application")) //запустить сервер
		{
			AfxMessageBox("Невозможно запустить Excel.");
			return;
		}

		oBooks = oApp.GetWorkbooks(); // и получаем список книг
		oBook = oBooks.Add(covOptional);
		
		oSheets = oBook.GetWorksheets(); //и получаем список листов
		
		oSheet =oSheets.GetItem(COleVariant(long(1)));

		FWait* waitdlg;
		waitdlg=new FWait();
		waitdlg->Create(FWait::IDD);
		waitdlg->ShowWindow(SW_SHOW);
		waitdlg->SetInfo("Формирование документа...");

		long counter=0,row=0,rows;
		rows=bt->GetRS()->GetRecordCount();

		if(bt->GetRS()->IsOpen())
		{
			nFields=bt->GetRS()->GetFieldCount();
			/*проверяем количество записей результате*/
			if(rows<1) rows=0;		
			else bt->GetRS()->MoveFirst();

			CADOFieldInfo info;
			/*список полей в шапку*/
			
			for(unsigned long i=0; i<nFields; ++i)
			{
				firstLetter=unsigned short (i/26.0);
				secondLetter=unsigned short(i-firstLetter*26);
				if(firstLetter<1) col.Format("%c1",secondLetter+'A');
				else col.Format("%c%c1",firstLetter+'A'-1,secondLetter+'A');
				y=COleVariant(col);
				oRange = oSheet.GetRange(y,y); //например
				bt->GetRS()->GetFieldInfo(i,&info);
				oRange.SetValue(covOptional,COleVariant(info.m_strName));
				oRange.BorderAround(COleVariant(short(2)),1,-4105,COleVariant(long(-4105)));
				//oRange.SetShrinkToFit(covTrue);
			}
			/*данные*/
			row=2;
			while(!bt->GetRS()->IsEof()&&rows>0)
			{
				for(unsigned long i=0; i<nFields; ++i)
				{
					firstLetter=unsigned short (i/26.0);
					secondLetter=unsigned short(i-firstLetter*26);
					if(firstLetter<1) col.Format("%c%d",secondLetter+'A',row);
					else col.Format("%c%c%d",firstLetter+'A'-1,secondLetter+'A',row);
					y=COleVariant(col);
					oRange = oSheet.GetRange(y,y); //например
					oRange.SetValue(covOptional,COleVariant(bt->GetStringValue(i)));
					oRange.BorderAround(COleVariant(short(2)),1,-4105,COleVariant(long(-4105)));
					//oRange.SetShrinkToFit(covTrue);
				}
				col.Format("Формирование... %d строк(%d",row,int(row*100/rows));
				col+="%)";
				waitdlg->SetInfo(col);
				bt->GetRS()->MoveNext(); row++;
			}
		}
	
		waitdlg->DestroyWindow(); delete waitdlg;
		x=COleVariant("A1");
		firstLetter=unsigned short ((nFields-1)/26.0);
		secondLetter=unsigned short((nFields-1)-firstLetter*26);
		if(firstLetter<1) col.Format("%c%d",secondLetter+'A',rows+1);
		else col.Format("%c%c%d",firstLetter+'A'-1,secondLetter+'A',rows+1);
		y=COleVariant(col);
		oRange = oSheet.GetRange(x,y);
		oRange.BorderAround(COleVariant(short(1)),2,-4105,COleVariant(long(-4105)));
		//oRange.SetShrinkToFit(covTrue);
		oRange=oRange.GetEntireColumn();
		oRange.AutoFit();
		if(firstLetter<1) col.Format("%c1",secondLetter+'A');
		else col.Format("%c%c1",firstLetter+'A'-1,secondLetter+'A');
		y=COleVariant(col);
		oRange = oSheet.GetRange(x,y);
		oInterior=oRange.GetInterior();
		oInterior.SetColorIndex(COleVariant(long(15)));
		oSheet.Activate();
		oApp.SetVisible(TRUE);
		oApp.SetUserControl(TRUE);
		bt->GetRS()->Close();
	}//end try
	catch(CADOException* e)
	{
		AfxMessageBox(e->GetErrorMessage(),MB_OK|MB_ICONERROR);
	}
	catch(...)
	{
		AfxMessageBox(bt->GetDB()->GetErrorDescription(),MB_OK|MB_ICONERROR);
	}
}

void CKanView::OnTimer(UINT nIDEvent) 
{
	CString UserName;
	DWORD len=100;
	GetUserName(UserName.GetBuffer(100),&len);
	UserName.ReleaseBuffer();
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), "[SetActiveUser]");
	pCmd.AddParameter("[UserID]",CADORecordset::typeInteger,
					CADOParameter::paramInput,sizeof(int),((CKanApp*)AfxGetApp())->m_iUserID);
	pCmd.AddParameter("[UserName]",CADORecordset::typeChar,
					CADOParameter::paramInput,50,UserName);
	switch(nIDEvent)
	{
	case ID_TIMER_1:
		
		try
		{
			bt->GetRS()->Execute(&pCmd);
		}catch(CADOException* e)
		{	
			bt->ThrowError(e);
			return;
		}
		bt->GetRS()->Close();
		break;
	default:break;
	}
	
	CFormView::OnTimer(nIDEvent);
}

void CKanView::StartSession()
{
	OnTimer(ID_TIMER_1);
	SetTimer(ID_TIMER_1,60000,NULL);
	m_bTimer=TRUE;
}

int CKanView::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CFormView::OnCreate(lpCreateStruct) == -1)
		return -1;
	return 0;
}

void CKanView::OnStatOnline() 
{
	if(!((CKanApp*)AfxGetApp())->m_bState) return;
	FStatistic dlg;
	dlg.m_csQuery="[GetOnlineUsers]";
	dlg.m_csTitle="Пользователи, работающие с программой в данный момент ";
	dlg.DoModal();
	
}

void CKanView::OnLetterRefresh() 
{
	CKanDoc* doc=GetDocument();
	switch(m_Tab.GetCurSel()+1)
	{
		case 1://inbox
			((FInbox*)(doc->m_pTabDialog))->Show();
			break;
		case 2://outbox
			((FOutbox*)(doc->m_pTabDialog))->Show();
			break;
		case 3://outbox
			((FControl*)(doc->m_pTabDialog))->Show();
			break;
		default:break;
	}
	
}

void CKanView::OnLetterUnfilter() 
{
	CKanDoc* doc=GetDocument();
	switch(m_Tab.GetCurSel()+1)
	{
		case 1://inbox
			if(m_inFilterContent==""&&
				m_inFilterDeclarant=="") return;
			((FInbox*)(doc->m_pTabDialog))->m_inFilterDeclarant=_T("");
			((FInbox*)(doc->m_pTabDialog))->m_inFilterContent=_T("");
			((FInbox*)(doc->m_pTabDialog))->Show();
			m_inFilterDeclarant=_T("");
			m_inFilterContent=_T("");
			break;
		case 2://outbox
			if(m_outFilterAuthor==""&&
				m_outFilterContent==""&&
				m_outFilterReciever=="") return;
			((FOutbox*)(doc->m_pTabDialog))->m_outFilterAuthor=_T("");
			((FOutbox*)(doc->m_pTabDialog))->m_outFilterContent=_T("");
			((FOutbox*)(doc->m_pTabDialog))->m_outFilterReciever=_T("");
			((FOutbox*)(doc->m_pTabDialog))->Show();
			m_outFilterAuthor=_T("");
			m_outFilterContent=_T("");
			m_outFilterReciever=_T("");
			break;
		default:break;
	}

	
}

void CKanView::OnReportsSign() 
{
	if(!((CKanApp*)AfxGetApp())->m_bReport) return;
	_EApplication oApp;
	Workbooks oBooks; //объявляем коллекцию "рабочих книг"
	_Workbook oBook; //объявляем конкретный экземпляр (класс _Workbook)
	Worksheets oSheets; //объявляем коллекцию "рабочих листов"
	_Worksheet oSheet; //объявляем конкретный экземпляр (класс _Worksheet)
	Range oRange; //объявляем экземпляр класса Range, отвечающего за заполнение ячеек
	Interior oInterior;
	CString col,num;
	COleVariant x,y,res;
	//CString query;
	CString oldNum;
	unsigned short firstLetter,secondLetter;
	unsigned long nFields;

	COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant  covTrue(short(1), VT_BOOL);
	COleVariant  covFalse(short(0), VT_BOOL);
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), "[ReportSign]");
	FChoosDate dlg;
	try{
		if(dlg.DoModal()!=IDOK) return;

		pCmd.AddParameter("BDate",CADORecordset::typeDate,
					CADOParameter::paramInput,12,dlg.m_Date.Format(VAR_DATEVALUEONLY));
		bt->Execute(&pCmd);
		if(!oApp.CreateDispatch("Excel.Application")) //запустить сервер
		{
			AfxMessageBox("Невозможно запустить Excel.");
			return;
		}
		
		oBooks = oApp.GetWorkbooks(); // и получаем список книг
		//oBook = oBooks.Add(covOptional);
		oBook=oBooks.Open(((CKanApp*)AfxGetApp())->m_csAppPath+"\\на_подпись.xls", covOptional, covOptional, covOptional, covOptional, covOptional, 
		covOptional, covOptional,covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional);
		oSheets = oBook.GetWorksheets(); //и получаем список листов
		
		oSheet =oSheets.GetItem(COleVariant(long(1)));

		FWait* waitdlg;
		waitdlg=new FWait();
		waitdlg->Create(FWait::IDD);
		waitdlg->ShowWindow(SW_SHOW);
		waitdlg->SetInfo("Формирование документа...");

		long counter=0,row=0,rows;
		
			rows=bt->GetRS()->GetRecordCount();

			if(bt->GetRS()->IsOpen())
			{
				nFields=bt->GetRS()->GetFieldCount();

				/*проверяем количество записей результате*/
				if(rows<1) rows=0;		
				else bt->GetRS()->MoveFirst();
		
				CADOFieldInfo info;
				/*список полей в шапку*/
				
				for(unsigned long i=0; i<nFields; ++i)
				{
					firstLetter=unsigned short (i/26.0);
					secondLetter=unsigned short(i-firstLetter*26);
					if(firstLetter<1) col.Format("%c1",secondLetter+'A');
					else col.Format("%c%c1",firstLetter+'A'-1,secondLetter+'A');
					y=COleVariant(col);
					oRange = oSheet.GetRange(y,y); //например

					bt->GetRS()->GetFieldInfo(i,&info);
					oRange.SetValue(covOptional,COleVariant(info.m_strName));
					oRange.BorderAround(COleVariant(short(2)),1,-4105,COleVariant(long(-4105)));
					//oRange.SetShrinkToFit(covTrue);
				}
				/*данные*/
				row=2;
				oldNum=_T("");
				while(!bt->GetRS()->IsEof()&&rows>0)
				{
					for(unsigned long i=0; i<nFields; ++i)
					{
						if(i==0)
						{
							if (oldNum==bt->GetStringValue(0))	i=3;
							oldNum=bt->GetStringValue(i);
						}
						firstLetter=unsigned short (i/26.0);
						secondLetter=unsigned short(i-firstLetter*26);
						if(firstLetter<1) col.Format("%c%d",secondLetter+'A',row);
						else col.Format("%c%c%d",firstLetter+'A'-1,secondLetter+'A',row);
						y=COleVariant(col);
						oRange = oSheet.GetRange(y,y); //например
						oRange.SetValue(covOptional,COleVariant(bt->GetStringValue(i)));
						oRange.BorderAround(COleVariant(short(2)),1,-4105,COleVariant(long(-4105)));
						//oRange.SetShrinkToFit(covTrue);
					}
					col.Format("Формирование... %d строк(%d",row,int(row*100/rows));
					col+="%)";
					waitdlg->SetInfo(col);
					bt->GetRS()->MoveNext(); row++;
				}
			}
		waitdlg->DestroyWindow(); delete waitdlg;
		x=COleVariant("A1");
		firstLetter=unsigned short ((nFields-1)/26.0);
		secondLetter=unsigned short((nFields-1)-firstLetter*26);
		if(firstLetter<1) col.Format("%c%d",secondLetter+'A',rows+1);
		else col.Format("%c%c%d",firstLetter+'A'-1,secondLetter+'A',rows+1);
		y=COleVariant(col);
		oRange = oSheet.GetRange(x,y);
		oRange.BorderAround(COleVariant(short(1)),2,-4105,COleVariant(long(-4105)));
		//oRange.SetShrinkToFit(covTrue);
		oRange=oRange.GetEntireColumn();
		//oRange.AutoFit();
		if(firstLetter<1) col.Format("%c1",secondLetter+'A');
		else col.Format("%c%c1",firstLetter+'A'-1,secondLetter+'A');
		/*y=COleVariant(col);
		oRange = oSheet.GetRange(x,y);
		oInterior=oRange.GetInterior();
		oInterior.SetColorIndex(COleVariant(long(15)));*/
		oSheet.Activate();
		oApp.SetVisible(TRUE);
		oApp.SetUserControl(TRUE);
		bt->GetRS()->Close();	
	}//end try
	catch(CADOException* e)
	{
		AfxMessageBox(e->GetErrorMessage(),MB_OK|MB_ICONERROR);
	}
	catch(...)
	{
		AfxMessageBox(bt->GetDB()->GetErrorDescription(),MB_OK|MB_ICONERROR);
	}
	
}
