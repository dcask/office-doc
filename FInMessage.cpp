// FInMessage.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FInMessage.h"
#include "FExec.h"
#include "FTheme.h"
#include "FResolution.h"
#include "FDeclarant.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FInMessage dialog


FInMessage::FInMessage(CWnd* pParent /*=NULL*/)
	: CDialog(FInMessage::IDD, pParent)
{
	//{{AFX_DATA_INIT(FInMessage)
	m_csContent = _T("");
	m_csControlType = _T("");
	m_csDeclarant = _T("");
	m_OutDate = COleDateTime::GetCurrentTime();
	m_RegDate = COleDateTime::GetCurrentTime();
	m_ProcessDate = COleDateTime::GetCurrentTime();
	m_csRegNum = _T("");
	m_csExtInNum = _T("");
	m_csDeclarantType = _T("");
	m_ProcessDate = COleDateTime::GetCurrentTime();
	m_csFullNum = _T("");
	//}}AFX_DATA_INIT
	m_pTabDialog=NULL;
	//m_OutDate.m_status = COleDateTime::null;
	m_hIcon = AfxGetApp()->LoadIcon(IDI_INLETTER);
	m_ToolTip=NULL;
	m_bNew=FALSE;
	m_ProcessDate.m_status=COleDateTime::null;
	m_OutDate.m_status=COleDateTime::null;
}


void FInMessage::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FInMessage)
	DDX_Control(pDX, IDC_LIST, m_List);
	DDX_Control(pDX, IDC_TAB, m_Tab);
	DDX_Text(pDX, IDC_CONTENT, m_csContent);
	DDX_CBString(pDX, IDC_CONTROLTYPE, m_csControlType);
	DDX_Text(pDX, IDC_FROM_PERSON, m_csDeclarant);
	DDX_Control(pDX, IDC_MOUT_DATE, m_mOutDate);
	DDX_Control(pDX, IDC_MREG_DATE, m_mRegDate);
	DDX_DateTimeCtrl(pDX, IDC_OUT_DATE, m_OutDate);
	DDX_DateTimeCtrl(pDX, IDC_REG_DATE, m_RegDate);
	DDX_Text(pDX, IDC_REG_NUM, m_csRegNum);
	DDX_Text(pDX, IDC_X_INNUM, m_csExtInNum);
	DDX_CBString(pDX, IDC_DECL_TYPE, m_csDeclarantType);
	DDX_DateTimeCtrl(pDX, IDC_PROC_DATE, m_ProcessDate);
	DDX_Text(pDX, IDC_REGNUM_FULL, m_csFullNum);
	DDX_Control(pDX, IDC_MPROC_DATE, m_mProcessDate);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FInMessage, CDialog)
	//{{AFX_MSG_MAP(FInMessage)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB, OnSelchangeTab)
	ON_WM_SHOWWINDOW()
	ON_WM_CREATE()
	ON_BN_CLICKED(IDC_SAVE, OnSave)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_REG_DATE, OnDatetimechangeRegDate)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_OUT_DATE, OnDatetimechangeOutDate)
	ON_BN_CLICKED(IDC_DECL_FIND, OnDeclFind)
	ON_CBN_SELCHANGE(IDC_CONTROLTYPE, OnSelchangeControltype)
	ON_CBN_SELCHANGE(IDC_DECL_TYPE, OnSelchangeDeclType)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_PROC_DATE, OnDatetimechangeProcDate)
	ON_EN_CHANGE(IDC_REG_NUM, OnChangeRegNum)
	ON_EN_CHANGE(IDC_FROM_PERSON, OnChangeFromPerson)
	ON_BN_CLICKED(IDC_ADD, OnAdd)
	ON_BN_CLICKED(IDC_CANCEL, OnCancel)
	ON_BN_CLICKED(IDC_DEL, OnDel)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FInMessage message handlers

BOOL FInMessage::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	NMHDR hdr;
	TCITEM tcItem;
	CRect rect;
	/*ToolBar*/
	GetClientRect(&rect); 
	unsigned int buttons[2];
	
	buttons[0]=IDC_SAVE;
	buttons[1]=IDC_CANCEL;

	m_pImgList=new CImageList();
	m_pImgList_d=new CImageList();

	m_pImgList->Create(96, 32, ILC_COLOR32, 4, 1);
	m_pImgList->SetBkColor(GetSysColor(COLOR_3DFACE)); 
	m_pImgList->Add(AfxGetApp()->LoadIcon(IDI_SAVE));
	m_pImgList->Add(AfxGetApp()->LoadIcon(IDI_CANCEL));
	
	m_pImgList_d->Create(96, 32, ILC_COLOR32, 4, 1);
	m_pImgList_d->SetBkColor(GetSysColor(COLOR_3DFACE)); 
	m_pImgList_d->Add(AfxGetApp()->LoadIcon(IDI_OK));
	m_pImgList_d->Add(AfxGetApp()->LoadIcon(IDI_CANCEL));

	m_wndToolBar.GetToolBarCtrl().SetImageList(m_pImgList);
	m_wndToolBar.GetToolBarCtrl().SetDisabledImageList(m_pImgList_d);
	m_wndToolBar.SetButtons(buttons,2);
	
	m_wndToolBar.MoveWindow(0, 0, rect.Width(), 40);
	/**/
	CTabCtrl *tab=(CTabCtrl*)GetDlgItem(IDC_TAB);
	
	hdr.code = TCN_SELCHANGE;
	hdr.hwndFrom = tab->m_hWnd;
	
	tcItem.mask = TCIF_TEXT | TCIF_IMAGE   ;
	tcItem.pszText = _T("»сполнители");tcItem.iImage=0;tab->InsertItem(0, &tcItem);
	tcItem.pszText = _T("–езолюци€");tcItem.iImage=1;tab->InsertItem(1, &tcItem);
	//tcItem.pszText = _T("“емы");tcItem.iImage=2;tab->InsertItem(2, &tcItem);
	
	tab->SetCurSel(0);
	SendMessage ( WM_NOTIFY, tab->GetDlgCtrlID(), (LPARAM)&hdr );

	SetIcon(m_hIcon, FALSE);        // ”станавливаем маленькую иконку
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void FInMessage::OnSelchangeTab(NMHDR* pNMHDR, LRESULT* pResult) 
{
	int id; // ID диалога
	// надо сначала удалить предыдущий диалог в Tab Control'е:
	if (m_pTabDialog)
	{
		m_pTabDialog->DestroyWindow();
		delete m_pTabDialog;
		m_pTabDialog=NULL;
	}

	// теперь в зависимости от того, кака€ закладка выбрана, 
	// выбираем соотв. диалог
	switch( m_Tab.GetCurSel()+1 ) // +1 дл€ того, чтобы пор€дковые номера закладок
	// и диалогов совпадали с номерами в case
	{
	// перва€ закладка
		case 2 :
			id = IDD_RESOLUTION;
			m_pTabDialog = new FResolution();
			((FResolution*) m_pTabDialog)->doc=doc;
			// тип указател€ соответствует нужному диалогу,
			// иначе добавленный в диалог код не будет функционировать
			break;
			// втора€ закладка
		case 1 :
			id = IDD_EXECUTOR;
			m_pTabDialog = new FExec();
			((FExec*) m_pTabDialog)->m_paExecuters=&(doc->m_aExecuters);
			((FExec*) m_pTabDialog)->m_bAnswer=(doc->m_iAttrib!=0 ? TRUE: FALSE );
			((FExec*) m_pTabDialog)->m_ProcessDate=m_ProcessDate;
			((FExec*) m_pTabDialog)->m_lInLetterID=doc->m_lLetterId;
			break;
		case 3 :
			id = IDD_THEME;
			m_pTabDialog = new FTheme;
			break;
		default:
			m_pTabDialog = new CDialog; // новый пустой диалог
			return;

	} // end switch

	// создаем диалог

	m_pTabDialog->Create (id, (CWnd*)&m_Tab); //параметры: ресурс диалога и родитель

	CRect rc; 
	m_uiCurSelectedTab=m_Tab.GetCurSel();

	m_Tab.GetWindowRect (&rc); // получаем "рабочую область"
	m_Tab.ScreenToClient (&rc); // преобразуем в относительные координаты

	// исключаем область, где отображаютс€ названи€ закладок:
	m_Tab.AdjustRect (FALSE, &rc); 

	// помещаем диалог на место..
	m_pTabDialog->MoveWindow (&rc);

	// и показываем:
	m_pTabDialog->ShowWindow ( SW_SHOWNORMAL );
	m_pTabDialog->UpdateWindow();
	OnSelchangeControltype();
	*pResult = 0;
}

void FInMessage::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	int iDtype=0;
	if(bShow)
	{

		/*«аполнение комбобоксов*/
		CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
		bt->LoadComboBox(GetDlgItem(IDC_CONTROLTYPE),"[ComboControlType]");
		bt->LoadComboBox(GetDlgItem(IDC_DECL_TYPE),"[ComboDeclType]");
		CComboBox* ct=(CComboBox*) GetDlgItem(IDC_CONTROLTYPE);
		CComboBox* dt=(CComboBox*) GetDlgItem(IDC_DECL_TYPE);
		ct->SetCurSel(1);
		dt->SetCurSel(0);
		/**/
		if(doc->m_lLetterId!=0)
		{
			m_csContent=doc->m_csContent;
			m_csControlType=doc->m_csControlType;
			m_csDeclarantType=doc->m_csDeclarantType;
			ct->SetCurSel(ct->FindString(0,m_csControlType));
			iDtype=dt->FindString(0,m_csDeclarantType);
			if(iDtype>=0) dt->SetCurSel(iDtype); else dt->SetCurSel(0);
			m_csDeclarant=doc->m_csDeclarant;
			m_csRegNum=doc->m_csRegNum;
			m_csFullNum=doc->m_csFullNum;
			m_RegDate=doc->m_RegDate;
			m_ProcessDate=doc->m_ProcessDate;

			UpdateData(FALSE);
		}
		else OnSelchangeDeclType();
			
		GetDlgItem(IDC_MREG_DATE)->SetFocus();

		CString csFormatDate=((CKanApp*) AfxGetApp())->m_csFormatDate;
		CString csVal;
		
		for(int i=0; i<csFormatDate.GetLength(); ++i) 
				if(csFormatDate.GetAt(i)<='A') csVal+=csFormatDate.GetAt(i); 
				else csVal+="9";


		m_mOutDate.SetMask(csVal);
		m_mRegDate.SetMask(csVal);
		m_mProcessDate.SetMask(csVal);
		
		m_mOutDate.SetFormat(csFormatDate);
		m_mRegDate.SetFormat(csFormatDate);
		m_mProcessDate.SetFormat(csFormatDate);

		if( m_OutDate.Format("%x") !="") m_mOutDate.SetText(m_OutDate.Format("%x"));
		if( m_RegDate.Format("%x") !="") m_mRegDate.SetText(m_RegDate.Format("%x"));
		if( m_ProcessDate.Format("%x") !="") m_mProcessDate.SetText(m_ProcessDate.Format("%x"));

		OnSelchangeControltype();
		m_ToolTip =new CToolTipCtrl();
		if (!m_ToolTip->Create(this))
		{
			TRACE("Unable To create ToolTip\n");           
			return;
		}
		
		/*Ќомера*/
		InNum* innum;
		for(i=0; i<doc->m_aExtNum.GetSize(); ++i)
		{
			innum=(InNum*) doc->m_aExtNum.GetAt(i);
			csVal.Format("є %s от %s",innum->csRegNum,innum->RegDate.Format("%x"));
			m_List.SetItemDataPtr(m_List.AddString(csVal),innum);
		}

		m_ToolTip->AddTool(GetDlgItem(IDC_DECL_FIND),"¬ыбрать за€вител€");
		m_ToolTip->AddTool(GetDlgItem(IDC_DEL),"”далить указанную св€зь");
		m_ToolTip->Activate(TRUE);
		if(!m_bNew)
			((CEdit*) GetDlgItem(IDC_REG_NUM))->SetReadOnly();
	}
}

int FInMessage::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	if (!m_wndToolBar.CreateEx(this, TBSTYLE_FLAT| TBSTYLE_TRANSPARENT, WS_CHILD | WS_VISIBLE | CBRS_TOP
		| CBRS_GRIPPER | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC))
	{
		TRACE0("Failed to create toolbar\n");
		return -1;      // fail to create
	}
	m_wndToolBar.GetToolBarCtrl().SetButtonSize(CSize(96,40));
	m_wndToolBar.GetToolBarCtrl().SetBitmapSize(CSize(96,32));
	return 0;
}

void FInMessage::OnCancel() 
{
	this->EndDialog(IDCANCEL);
	
}

void FInMessage::OnSave() 
{
	if(!((CKanApp*)AfxGetApp())->m_bEditIn) return;
	UpdateData();
	if(AfxMessageBox("—охранить ?", MB_YESNO | MB_ICONQUESTION)!=IDYES) return;
	
	if (m_pTabDialog)
			m_pTabDialog->DestroyWindow();
	
	if(m_csExtInNum!="") OnAdd();
	
		CString tmp=_T("%");
		int iYear=atoi(COleDateTime::GetCurrentTime().Format("%Y"));
		CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
		CADOCommand pCmd(bt->GetDB(), "[FindIn]");
		CComboBox* cb=(CComboBox*) GetDlgItem(IDC_CONTROLTYPE);
		/*проверка на наличие номера если новый*/
		pCmd.AddParameter("InRegNum",CADORecordset::typeChar,
				CADOParameter::paramInput,50,m_csRegNum);
		pCmd.AddParameter("OutRegNum",CADORecordset::typeChar,
				CADOParameter::paramInput,50,tmp);
		pCmd.AddParameter("Decl",CADORecordset::typeChar,
				CADOParameter::paramInput,50,tmp);
		pCmd.AddParameter("RegDateYear",CADORecordset::typeUnsignedInt,
				CADOParameter::paramInput,sizeof(unsigned int),iYear);
		if(m_bNew)
		{
			try
			{
				bt->GetRS()->Execute(&pCmd);
			}catch(CADOException* e)
			{
				bt->ThrowError(e);
				return;
			}
			if(bt->GetRecordsCount()>0)
			{
				AfxMessageBox("Ќомер уже зан€т");
				return;
			}
		}
		bt->GetRS()->Close();

		doc->m_csContent=m_csContent;
		doc->m_csControlType=m_csControlType;
		doc->m_iControlType=cb->GetItemData(cb->GetCurSel());
		m_csDeclarant.TrimLeft();m_csDeclarant.TrimRight();
		while (m_csDeclarant.Replace("  "," ")>0);
		m_csDeclarant.Remove('%');m_csDeclarant.Remove('_');
		doc->m_csDeclarant=m_csDeclarant;
		cb=(CComboBox*) GetDlgItem(IDC_DECL_TYPE);
		doc->m_iDeclarantType=cb->GetItemData(cb->GetCurSel());
		m_csRegNum.Remove('%');m_csRegNum.Remove('_');
		doc->m_csRegNum=m_csRegNum;
		m_csFullNum.Remove('%');m_csFullNum.Remove('_');
		doc->m_csFullNum=m_csFullNum;
		doc->m_RegDate.ParseDateTime(m_mRegDate.GetFormattedText());
		doc->m_ProcessDate.ParseDateTime(m_mProcessDate.GetFormattedText());
		if(m_csExtInNum!="") doc->m_csOutRegNum_F=m_csExtInNum;
		else doc->m_csOutRegNum_F="б/н";
		doc->m_OutRegDate_F.ParseDateTime(m_mOutDate.GetFormattedText());
		this->EndDialog(IDOK);

	
}

void FInMessage::OnDatetimechangeRegDate(NMHDR* pNMHDR, LRESULT* pResult) 
{
	if( m_RegDate.Format("%x") !="") m_mRegDate.SetText(m_RegDate.Format("%x"));
	UpdateData();
	*pResult = 0;
}

void FInMessage::OnDatetimechangeOutDate(NMHDR* pNMHDR, LRESULT* pResult) 
{
	if( m_OutDate.Format("%x") !="") m_mOutDate.SetText(m_OutDate.Format("%x"));
	UpdateData();
	*pResult = 0;
}

BOOL FInMessage::PreTranslateMessage(MSG* pMsg) 
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
					if(id==IDC_FROM_PERSON) 
						OnDeclFind();
					else
					{
						if(id==IDC_X_INNUM)
							OnAdd();
						else OnSave();
					}
					
						
				break;
			case VK_ESCAPE:
				OnCancel();
				break;
			default:
				return CDialog::PreTranslateMessage(pMsg);
		}
		return TRUE;       // запрет дальнейшей обработки
	}
	return CDialog::PreTranslateMessage(pMsg);
}

void FInMessage::OnDeclFind() 
{
	FDeclarant dlg;
	GetDlgItemText(IDC_FROM_PERSON, dlg.m_csString);
	CComboBox* cb=(CComboBox*) GetDlgItem(IDC_DECL_TYPE);
	dlg.m_iDeclType=cb->GetItemData(cb->GetCurSel());
	if(dlg.DoModal()==IDOK)
	{
		SetDlgItemText(IDC_FROM_PERSON, dlg.m_csString);
	}
	
}

void FInMessage::OnSelchangeControltype() 
{
	CComboBox* cb=(CComboBox*) GetDlgItem(IDC_CONTROLTYPE);
	int size=1500;
	if(cb->GetCurSel()==1) size=0;
	switch( m_Tab.GetCurSel()+1 ) 
	{
		case 2 :
			break;
		case 1 :
			((FExec*) m_pTabDialog)->m_iSize=size;
			((FExec*) m_pTabDialog)->View();
			break;
		case 3 :
			break;
		default:break;
	} // end switch
}

void FInMessage::OnSelchangeDeclType() 
{
	UpdateData(TRUE);
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmd(bt->GetDB(), "[GetTemplate]");
	CComboBox* cb=(CComboBox*) GetDlgItem(IDC_DECL_TYPE);
	int iCur=cb->GetCurSel();
	int iTmp=cb->GetItemData(iCur);
	
	CString num_template;
	pCmd.AddParameter("TempID",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),iTmp);
	//запрос
	try
	{
		bt->GetRS()->Execute(&pCmd);
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		return;
	}
	bt->GetRS()->MoveFirst();
	num_template=bt->GetStringValue("Template");
	bt->GetRS()->Close();
	m_csFullNum=ApplyTemplate(num_template,m_csRegNum, m_csDeclarant);
	UpdateData(FALSE);
	cb->SetCurSel(iCur);
	
}

void FInMessage::OnDatetimechangeProcDate(NMHDR* pNMHDR, LRESULT* pResult) 
{
	if( m_ProcessDate.Format("%x") !="") m_mProcessDate.SetText(m_ProcessDate.Format("%x"));
	UpdateData();
	*pResult = 0;
}

void FInMessage::OnChangeRegNum() 
{
	OnSelchangeDeclType();	
}

CString FInMessage::ApplyTemplate(CString Templ, CString Num, CString Decl)
{
	CString res,cstmp;
	res=Templ;
	Decl.TrimLeft();
	if(Decl.GetLength()>0) cstmp=Decl.GetAt(0); else cstmp=_T("");
	cstmp.MakeUpper();
	res.Replace("%n",Num);
	res.Replace("%F",cstmp);
	return res;
}

void FInMessage::OnChangeFromPerson() 
{
	OnSelchangeDeclType();
}

BEGIN_EVENTSINK_MAP(FInMessage, CDialog)
    //{{AFX_EVENTSINK_MAP(FInMessage)
	ON_EVENT(FInMessage, IDC_MPROC_DATE, 1 /* Change */, OnChangeMprocDate, VTS_NONE)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

void FInMessage::OnChangeMprocDate() 
{
	switch( m_Tab.GetCurSel()+1 ) 
	{
		case 1 :
			((FExec*) m_pTabDialog)->m_ProcessDate.ParseDateTime(m_mProcessDate.GetFormattedText());
			break;
		default:
			return;

	} // end switch
	
}

BOOL FInMessage::DestroyWindow() 
{
	delete m_pImgList;
	delete m_pImgList_d;
	delete m_pTabDialog;
	if(m_ToolTip) m_ToolTip->DestroyWindow();
	delete m_ToolTip;
	
	return CDialog::DestroyWindow();
}

void FInMessage::OnAdd() 
{

	InNum*	innum;
	CString	csVal,tmp;

	if(!UpdateData()) return;
	
	innum=new InNum();
	innum->csRegNum=m_csExtInNum;
	innum->RegDate.ParseDateTime(m_mOutDate.GetFormattedText());
	innum->state=1;
	csVal.Format("є %s от %s",m_csExtInNum,m_mOutDate.GetFormattedText());
	m_List.SetItemDataPtr(m_List.AddString(csVal),innum);
	doc->m_aExtNum.Add(innum);
	m_csExtInNum = _T("");
	UpdateData(FALSE);
}

void FInMessage::OnDel() 
{
	InNum*	innum;
	int pos=m_List.GetCurSel();
	if(pos<0) return;
	innum=(InNum*) m_List.GetItemDataPtr(pos);
	if(innum->state==0) innum->state=2;
	else innum->state=-1; 
	m_List.DeleteString(pos);

	
}
