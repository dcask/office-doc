// FOutMessage.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FOutMessage.h"
#include "FDeclarant.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FOutMessage dialog


FOutMessage::FOutMessage(CWnd* pParent /*=NULL*/)
	: CDialog(FOutMessage::IDD, pParent)
{
	//{{AFX_DATA_INIT(FOutMessage)
	m_csAuthor = _T("");
	m_csContent = _T("");
	m_RegDate = COleDateTime::GetCurrentTime();
	m_csRegNum = _T("");
	m_csReplyOn = _T("");
	m_iYear = 0;
	m_csPrefix = _T("");
	//}}AFX_DATA_INIT
	m_hIcon = AfxGetApp()->LoadIcon(IDI_OUTLETTER);
	m_ToolTip=NULL;
	m_bNew=FALSE;
	m_iYear = atoi(COleDateTime::GetCurrentTime().Format("%Y"));
}


void FOutMessage::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FOutMessage)
	DDX_Control(pDX, IDC_INLIST, m_InList);
	DDX_Control(pDX, IDC_SPIN1, m_Spin);
	DDX_Control(pDX, IDC_TREE, m_Tree);
	DDX_CBString(pDX, IDC_AUTHOR, m_csAuthor);
	DDX_Text(pDX, IDC_CONTENT, m_csContent);
	DDX_Control(pDX, IDC_MREG_DATE, m_mRegDate);
	DDX_DateTimeCtrl(pDX, IDC_REG_DATE, m_RegDate);
	DDX_Text(pDX, IDC_REG_NUM, m_csRegNum);
	DDX_Text(pDX, IDC_REPLY_NUM, m_csReplyOn);
	DDX_Text(pDX, IDC_YEAR, m_iYear);
	DDV_MinMaxInt(pDX, m_iYear, 1900, 2200);
	DDX_Text(pDX, IDC_PREFIX, m_csPrefix);
	DDX_Text(pDX, IDC_TOPLACE, m_csToPlace);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FOutMessage, CDialog)
	//{{AFX_MSG_MAP(FOutMessage)
	ON_BN_CLICKED(IDC_REPLY_FIND, OnReplyFind)
	ON_WM_SHOWWINDOW()
	ON_CBN_SELCHANGE(IDC_AUTHOR, OnSelchangeAuthor)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_REG_DATE, OnDatetimechangeRegDate)
	ON_WM_CREATE()
	ON_BN_CLICKED(IDC_SAVE, OnSave)
	ON_BN_CLICKED(IDC_BUTTON_ALL, OnButtonAll)
	ON_EN_CHANGE(IDC_REG_NUM, OnChangeRegNum)
	ON_BN_CLICKED(IDC_ADR_FIND, OnAdrFind)
	ON_BN_CLICKED(IDC_CANCEL, OnCancel)
	ON_BN_CLICKED(IDC_DEL, OnDel)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FOutMessage message handlers

void FOutMessage::OnReplyFind() 
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmdIn(bt->GetDB(), "[FindIn]");

	Reply*	reply;
	CString	csVal,tmp;

	if(!UpdateData()) return;
	
	long id;
	csVal="%";
	if(m_csReplyOn!="")	// если связь ещё не установлена
	{
		pCmdIn.AddParameter("InRegNum",CADORecordset::typeChar,
								CADOParameter::paramInput,50,m_csReplyOn);
		pCmdIn.AddParameter("RegDateYear",CADORecordset::typeUnsignedInt,
								CADOParameter::paramInput,sizeof(unsigned int),m_iYear);
		try
		{
			bt->GetRS()->Execute(&pCmdIn);
			}catch(CADOException* e)
		{
			bt->ThrowError(e);
			return;
		}
		if(bt->GetRecordsCount()<1)
		{
			AfxMessageBox("Не найден");
			return;
		}
		bt->GetRS()->MoveFirst();
		bt->GetRS()->GetFieldValue("ID",doc->m_lInLetterID);
		bt->GetRS()->Close();
	
		reply=new Reply();
		reply->inLetterID=doc->m_lInLetterID;
		reply->state=1;
		tmp.Format("№ %s за %d год",m_csReplyOn,m_iYear);
		m_InList.SetItemDataPtr(m_InList.AddString(tmp),reply);
		doc->m_aReplyOn.Add(reply);

		id=doc->m_lLetterId; // сохраняем ИД
		doc->m_lLetterId=doc->m_lInLetterID; // заменяем ИД для функции загрузки
		doc->LoadInData();
		LoadExecutors();
		doc->m_lLetterId=id; // восстанавливаем ИД
		m_csReplyOn=_T("");
		UpdateData(FALSE);
	}
}

BOOL FOutMessage::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	m_Tree.ModifyStyle( TVS_CHECKBOXES, 0 );
	m_Tree.ModifyStyle( 0, TVS_CHECKBOXES );

	/*ToolBar*/
	CRect rect;
	GetClientRect(&rect); 
	unsigned int buttons[2];
	
	buttons[0]=IDC_SAVE;
	buttons[1]=IDC_CANCEL;

	m_pImgList=new CImageList();
	m_pImgList_d=new CImageList();
	m_pImgList_Exec=new CImageList();

	m_pImgList->Create(96, 32, ILC_COLOR32, 4, 1);
	m_pImgList->SetBkColor(GetSysColor(COLOR_3DFACE)); 
	m_pImgList->Add(AfxGetApp()->LoadIcon(IDI_SAVE));
	m_pImgList->Add(AfxGetApp()->LoadIcon(IDI_CANCEL));

	m_pImgList_Exec->Create(16, 16, ILC_COLOR32, 4, 1);
	//m_pImgList_Exec->SetBkColor(GetSysColor(COLOR_3DFACE)); 
	m_pImgList_Exec->Add(AfxGetApp()->LoadIcon(IDI_PERSON));
	m_pImgList_Exec->Add(AfxGetApp()->LoadIcon(IDI_CONTROL));
	
	m_pImgList_d->Create(96, 32, ILC_COLOR32, 4, 1);
	m_pImgList_d->SetBkColor(GetSysColor(COLOR_3DFACE)); 
	m_pImgList_d->Add(AfxGetApp()->LoadIcon(IDI_OK));
	m_pImgList_d->Add(AfxGetApp()->LoadIcon(IDI_CANCEL));

	m_wndToolBar.GetToolBarCtrl().SetImageList(m_pImgList);
	m_wndToolBar.GetToolBarCtrl().SetDisabledImageList(m_pImgList_d);
	m_wndToolBar.SetButtons(buttons,2);
	
	m_wndToolBar.MoveWindow(0, 0, rect.Width(), 40);
	SetIcon(m_hIcon, FALSE);        // Устанавливаем маленькую иконку
	m_Tree.SetImageList(m_pImgList_Exec,TVSIL_NORMAL);
	/**/
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void FOutMessage::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	
	if(!bShow) return;
	
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	bt->LoadComboBox(GetDlgItem(IDC_AUTHOR),"[ComboPersons]");
	bt->LoadComboBox(GetDlgItem(IDC_PREFIX),"[ComboPrefix]");
	if(doc->m_lLetterId!=0)
	{
		m_csPrefix=doc->m_csRegPrefix;
		m_csRegNum=doc->m_csRegNum;
		m_csToPlace=doc->m_csToPlace;
		m_RegDate=doc->m_RegDate;
		m_csAuthor=doc->m_csAuthor;
		m_csContent=doc->m_csContent;
		UpdateData(FALSE);
	}
	m_Spin.SetRange(1900,2200);
	m_Spin.SetPos(m_iYear);

	/*Дата*/
	CString csFormatDate=((CKanApp*) AfxGetApp())->m_csFormatDate;
	CString csVal;
	
	for(int i=0; i<csFormatDate.GetLength(); ++i) 
			if(csFormatDate.GetAt(i)<='A') csVal+=csFormatDate.GetAt(i); 
			else csVal+="9";


	m_mRegDate.SetMask(csVal);
	
	m_mRegDate.SetFormat(csFormatDate);

	if( m_RegDate.Format("%x") !="") m_mRegDate.SetText(m_RegDate.Format("%x"));
	OnChangeRegNum();

	m_ToolTip =new CToolTipCtrl();
	if (!m_ToolTip->Create(this))
	{
		TRACE("Unable To create ToolTip\n");           
		return;
	}

	m_ToolTip->AddTool(GetDlgItem(IDC_REPLY_FIND),"Связать с входящим письмом");
	m_ToolTip->AddTool(GetDlgItem(IDC_BUTTON_ALL),"Снять со всех");
	m_ToolTip->AddTool(GetDlgItem(IDC_DEL),"Удалить указанную связь");
	m_ToolTip->Activate(TRUE);
	GetDlgItem(IDC_MREG_DATE)->SetFocus();
	if(!m_bNew)
		((CEdit*) GetDlgItem(IDC_REG_NUM))->SetReadOnly();
	/*Ответы*/
	Reply* reply;
	for(i=0; i<doc->m_aReplyOn.GetSize(); ++i)
	{
		reply=(Reply*) doc->m_aReplyOn.GetAt(i);
		csVal.Format("№ %s за %d год",reply->RegNum, reply->year);
		m_InList.SetItemDataPtr(m_InList.AddString(csVal),reply);
		
		long id=doc->m_lLetterId; // сохраняем ИД
		doc->m_lLetterId=reply->inLetterID; // заменяем ИД для йункции загрузки
		doc->LoadInData();
		LoadExecutors();
		doc->m_lLetterId=id; // восстанавливаем ИД
	}
}

void FOutMessage::LoadExecutors()
{
	CString csVal;
	HTREEITEM hItem,hItemBranch;
	BOOL found;
	Executer* exec;
	DWORD dwPersonId;
	/*заполнение дерева*/
	for(UINT i=0; i< UINT(doc->m_aExecuters.GetSize()); i++)
	{
		found=FALSE;
		exec=((Executer*) doc->m_aExecuters.GetAt(i));
		hItem = m_Tree.GetNextItem(m_Tree.GetFirstVisibleItem(),TVGN_ROOT);
		dwPersonId=exec->dwPersonId;
		for (UINT j=0;j < m_Tree.GetCount() && hItem;j++)
		{
			if(dwPersonId==m_Tree.GetItemData(hItem))
			{
				found=TRUE; // нашёл
				break;
			}
			hItem = m_Tree.GetNextItem(hItem,TVGN_NEXT);
		}

		if(!found)
		{
			csVal=exec->csName;
			hItem=m_Tree.InsertItem(csVal);
			m_Tree.SetItemImage(hItem,0,0);
			m_Tree.SetItemData(hItem,dwPersonId);
		}
	}
	/*заполнение веток*/
	ExecLetter* execLetter;
	for(i=0; i< UINT(doc->m_aExecuters.GetSize()); i++)
	{
		hItem = m_Tree.GetNextItem(m_Tree.GetFirstVisibleItem(),TVGN_ROOT);
		exec=((Executer*) doc->m_aExecuters.GetAt(i));
		dwPersonId=exec->dwPersonId;
		for (UINT j=0;j < m_Tree.GetCount() && hItem;j++)
		{
			found=FALSE;
			if(dwPersonId==m_Tree.GetItemData(hItem))
			{
				// добавляем в ветку
				csVal="Номер: "+doc->m_csRegNum;
				if(exec->scheduleDate.m_status!=COleDateTime::null)
				{
					csVal+="    Срок исполнения: ";
					csVal+=exec->scheduleDate.Format("%x");
				}
				hItemBranch=m_Tree.InsertItem(csVal,hItem);
				m_Tree.SetItemImage(hItemBranch,1,1);
				execLetter=new ExecLetter();
				execLetter->dwID=exec->dwID;
				execLetter->inLetterID=exec->dwInLetterId;
				execLetter->scheduleDate=exec->scheduleDate;
				m_Tree.SetItemData(hItemBranch,(DWORD)execLetter);
				if(exec->factDate.m_status==COleDateTime::null)
				{
					m_Tree.SetCheck(hItemBranch,TRUE);
					found=TRUE;
				}
				else
					m_Tree.SetCheck(hItemBranch,FALSE);
				if(found) m_Tree.SetCheck(hItem,TRUE);
				m_Tree.Expand(hItem,TVE_EXPAND);
				break;
			}
			hItem = m_Tree.GetNextItem(hItem,TVGN_NEXT);
		}
	}
}

void FOutMessage::OnSelchangeAuthor() 
{
	CComboBox* cb=(CComboBox*) GetDlgItem(IDC_AUTHOR);
	CString csVal;
	long id;
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	cb->GetLBText(cb->GetCurSel(),csVal);
	id=cb->GetItemData(cb->GetCurSel());
	CADOCommand pCmd(bt->GetDB(),"[GetPrefix]");
	pCmd.AddParameter("PersonID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),id);
	try
	{
		bt->GetRS()->Execute(&pCmd);
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		return;
	}

	if(bt->GetRecordsCount()==1)
	{
		bt->GetRS()->MoveFirst();
		csVal=bt->GetStringValue("Prefix");
		SetDlgItemText(IDC_PREFIX, csVal);
	}
	bt->GetRS()->Close();
	OnChangeRegNum();
}



void FOutMessage::OnDatetimechangeRegDate(NMHDR* pNMHDR, LRESULT* pResult) 
{
	if( m_RegDate.Format("%x") !="") m_mRegDate.SetText(m_RegDate.Format("%x"));
	UpdateData();
	
	*pResult = 0;
}

int FOutMessage::OnCreate(LPCREATESTRUCT lpCreateStruct) 
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

void FOutMessage::OnCancel() 
{
	this->EndDialog(IDCANCEL);
	
}

void FOutMessage::OnSave() 
{
	if(!((CKanApp*)AfxGetApp())->m_bEditOut) return;
	UpdateData();
	if(m_csAuthor=="")
	{
		AfxMessageBox("Укажите исполнителя");
		return;
	}
	if(AfxMessageBox("Сохранить ?", MB_YESNO | MB_ICONQUESTION)!=IDYES) return;


	if(m_csReplyOn!="") OnReplyFind();
	
		CString tmp=_T("%");
		int iYear=atoi(COleDateTime::GetCurrentTime().Format("%Y"));
		CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
		CADOCommand pCmd(bt->GetDB(), "[FindOut]");
		CComboBox* cb=((CComboBox*) GetDlgItem(IDC_AUTHOR));
		/*проверка на наличие номера*/
		pCmd.AddParameter("OutRegNum",CADORecordset::typeChar,
							CADOParameter::paramInput,50,m_csRegNum);
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
				AfxMessageBox("Номер уже занят");
				return;
			}
		}

		bt->GetRS()->Close();
		
		doc->m_csContent=m_csContent;
		doc->m_csRegPrefix=m_csPrefix;
		m_csRegNum.Remove('%');m_csRegNum.Remove('_');
		m_csToPlace.Remove('%');m_csToPlace.Remove('_');
		while(m_csToPlace.Replace("  "," ")>0);
		m_csToPlace.TrimLeft();m_csToPlace.TrimRight();
		doc->m_csRegNum=m_csRegNum;
		doc->m_csToPlace=m_csToPlace;
		doc->m_RegDate.ParseDateTime(m_mRegDate.GetFormattedText());
		doc->m_csAuthor=m_csAuthor;
		doc->m_iAuthor=cb->GetItemData(cb->GetCurSel());
		m_csReplyOn.Remove('%');m_csReplyOn.Remove('_');
		doc->m_csOutRegNum=m_csReplyOn;
		/////////////////////////////////////////////////////////////////
		// перезаполняем список исполнителей на основе дереве
		for(UINT i=0; i< UINT(doc->m_aExecuters.GetSize()); i++)
		delete ((Executer*)doc->m_aExecuters.GetAt(i));
		
		doc->m_aExecuters.RemoveAll();

		UINT uCount = m_Tree.GetCount();
		HTREEITEM hItem = m_Tree.GetNextItem(m_Tree.GetFirstVisibleItem(),TVGN_ROOT);
		HTREEITEM hItemChild;
		ExecLetter* execLetter;
		Executer* exec;
		for (i=0;i < uCount && hItem;i++)
		{
			hItemChild=m_Tree.GetNextItem(hItem,TVGN_CHILD);
			for(UINT j=0; j<uCount && hItemChild; ++j)
			{

				execLetter=(ExecLetter*) m_Tree.GetItemData(hItemChild);
				exec=new Executer();
				exec->dwID=execLetter->dwID;
				exec->dwInLetterId=execLetter->inLetterID;
				exec->dwPersonId=m_Tree.GetItemData(hItem);
				exec->scheduleDate=execLetter->scheduleDate;
				if(!m_Tree.GetCheck(hItemChild)) exec->factDate=doc->m_RegDate;
				else exec->factDate.m_status=COleDateTime::null;
				(Executer*)doc->m_aExecuters.Add(exec);
				hItemChild=m_Tree.GetNextItem(hItemChild,TVGN_NEXT);
			}
				
			hItem = m_Tree.GetNextItem(hItem,TVGN_NEXT);
		}
		this->EndDialog(IDOK);

	
}

void FOutMessage::OnButtonAll() 
{
	UINT uCount = m_Tree.GetCount();
	HTREEITEM hItem = m_Tree.GetNextItem(m_Tree.GetFirstVisibleItem(),TVGN_ROOT);
	HTREEITEM hItemChild;

	for (UINT i=0;i < uCount && hItem;i++)
	{
		m_Tree.SetCheck(hItem,FALSE);
		hItemChild=m_Tree.GetNextItem(hItem,TVGN_CHILD);
		for(UINT j=0; j<uCount && hItemChild; ++j)
		{
			m_Tree.SetCheck(hItemChild,FALSE);
			hItemChild=m_Tree.GetNextItem(hItemChild,TVGN_NEXT);
		}
			
		hItem = m_Tree.GetNextItem(hItem,TVGN_NEXT);
	}
}

BOOL FOutMessage::PreTranslateMessage(MSG* pMsg) 
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
					if(id==IDC_REPLY_NUM) OnReplyFind();
					else
						if(id==IDC_TOPLACE) OnAdrFind();
							else
								OnSave();
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

void FOutMessage::OnChangeRegNum() 
{
	CString num,prefix;
	GetDlgItem(IDC_REG_NUM)->GetWindowText(num);
	GetDlgItem(IDC_PREFIX)->GetWindowText(prefix);
	GetDlgItem(IDC_REGNUM_FULL)->SetWindowText(prefix+"/"+num);
}

BOOL FOutMessage::DestroyWindow() 
{
	delete m_pImgList;
	delete m_pImgList_d;
	delete m_pImgList_Exec;
	if(m_ToolTip) m_ToolTip->DestroyWindow();
	delete m_ToolTip;

	ExecLetter* execList;

	UINT uCount = m_Tree.GetCount();
	HTREEITEM hItem = m_Tree.GetNextItem(m_Tree.GetFirstVisibleItem(),TVGN_ROOT);
	HTREEITEM hItemChild;

	for (UINT i=0;i < uCount && hItem;i++)
	{
		hItemChild=m_Tree.GetNextItem(hItem,TVGN_CHILD);
		for(UINT j=0; j<uCount && hItemChild; ++j)
		{

			execList=(ExecLetter*) m_Tree.GetItemData(hItemChild);
			delete execList;
			hItemChild=m_Tree.GetNextItem(hItemChild,TVGN_NEXT);
		}
			
		hItem = m_Tree.GetNextItem(hItem,TVGN_NEXT);
	}
	return CDialog::DestroyWindow();
}

void FOutMessage::OnAdrFind() 
{
	FDeclarant dlg;
	GetDlgItemText(IDC_TOPLACE, dlg.m_csString);
	CComboBox* cb=(CComboBox*) GetDlgItem(IDC_TOPLACE);
	dlg.m_csQuery="[GetPlaces]";
	dlg.m_csWindowName="Адресат";
	if(dlg.DoModal()==IDOK)
	{
		SetDlgItemText(IDC_TOPLACE, dlg.m_csString);
	}
	
}

void FOutMessage::OnDel() 
{
	Reply*	reply;
	int pos=m_InList.GetCurSel();
	if(pos<0) return;
	reply=(Reply*) m_InList.GetItemDataPtr(pos);
	if(reply->state==0) reply->state=2;
	else reply->state=-1; 
	m_InList.DeleteString(pos);
	ExecLetter* execList;
	////////////////////////////////////////////////////////////	
	UINT uCount = m_Tree.GetCount();
	HTREEITEM hItem = m_Tree.GetNextItem(m_Tree.GetFirstVisibleItem(),TVGN_ROOT);
	HTREEITEM hItemChild;
	HTREEITEM hItemTemp;
	BOOL flag;

	for (UINT i=0;i < uCount && hItem;i++)
	{
		hItemChild=m_Tree.GetNextItem(hItem,TVGN_CHILD);
		flag=FALSE;
		for(UINT j=0; j<uCount && hItemChild; ++j)
		{
			DWORD temp=m_Tree.GetItemData(hItemChild);
			execList=(ExecLetter*) m_Tree.GetItemData(hItemChild);
			if(execList->inLetterID==reply->inLetterID) 
			{
				// удадяем элемент и проверяем родителя, если нет детей, удаляем
				delete execList;
				hItemTemp=hItemChild;
				hItemChild=m_Tree.GetNextItem(hItemChild,TVGN_NEXT);

				m_Tree.DeleteItem(hItemTemp);
				if(!m_Tree.ItemHasChildren(hItem)) flag=TRUE;
			}
			else hItemChild=m_Tree.GetNextItem(hItemChild,TVGN_NEXT);
		}
		hItemTemp=hItem;
		hItem = m_Tree.GetNextItem(hItem,TVGN_NEXT);
		if(flag) m_Tree.DeleteItem(hItemTemp);
	}
}
