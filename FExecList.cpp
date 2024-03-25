// FExecList.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
#include "FExecList.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FExecList dialog


FExecList::FExecList(CWnd* pParent /*=NULL*/)
	: CDialog(FExecList::IDD, pParent)
{
	//{{AFX_DATA_INIT(FExecList)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void FExecList::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FExecList)
	DDX_Control(pDX, IDC_TREE, m_Tree);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FExecList, CDialog)
	//{{AFX_MSG_MAP(FExecList)
	ON_WM_SHOWWINDOW()
	ON_WM_CLOSE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FExecList message handlers

void FExecList::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	
	if(!bShow) return;
	CString execs;
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	GetPrivateProfileString("Settings","Execs","%",execs.GetBuffer(5),5,csAppPath+"\\kan.ini");
	execs.ReleaseBuffer();
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	int post;
	CADOCommand pCmd(bt->GetDB(), "[PersonList]");
	pCmd.AddParameter("Attr",CADORecordset::typeChar,
				CADOParameter::paramInput,5,execs);
	CString csVal;
	DWORD	dwExecId;
	HTREEITEM hItem;
	
	bt->Execute(&pCmd);
	/*заполнение дерева*/
	if(bt->GetRecordsCount()>0) bt->GetRS()->MoveFirst();
	for(UINT i=0; i< bt->GetRecordsCount(); i++)
	{
		bt->GetRS()->GetFieldValue("ID",dwExecId);
		csVal=bt->GetStringValue("Name");
		csVal+=" : ";
		csVal+=bt->GetStringValue("Post");
		csVal.TrimLeft();csVal.TrimRight();
		bt->GetRS()->GetFieldValue("PostCode",post);
		if(post<0||post>3) post=2;
		if(dwExecId>0)
		{
			hItem=m_Tree.InsertItem(csVal);
			m_Tree.SetItemData(hItem,dwExecId);
			m_Tree.SetItemImage(hItem,post,post);
		}
		bt->GetRS()->MoveNext();
	}
	/*заполнение чекбоксов*/
	UINT uCount = m_Tree.GetCount();
	hItem = m_Tree.GetNextItem(m_Tree.GetFirstVisibleItem(),TVGN_ROOT);

	for (i=0;i < uCount && hItem;i++)
	{
		ASSERT(hItem != NULL);
		for(int j=0; j<(*m_paExecuters).GetSize(); ++j)
		{
			dwExecId=m_Tree.GetItemData(hItem);
			if( ((Executer*)((*m_paExecuters).GetAt(j)))->dwPersonId ==dwExecId)
				if(((Executer*)((*m_paExecuters).GetAt(j)))->state!=2) m_Tree.SetCheck(hItem);
		}
			
		hItem = m_Tree.GetNextItem(hItem,TVGN_NEXT);
	}
	bt->GetRS()->Close();
}

BOOL FExecList::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	m_Tree.ModifyStyle( TVS_CHECKBOXES, 0 );
	m_Tree.ModifyStyle( 0, TVS_CHECKBOXES );
	m_pImgList=new CImageList();
	m_pImgList->Create(16, 16, ILC_COLOR32, 4, 1);
	//m_pImgList_Exec->SetBkColor(GetSysColor(COLOR_3DFACE)); 
	m_pImgList->Add(AfxGetApp()->LoadIcon(IDI_POST1));
	m_pImgList->Add(AfxGetApp()->LoadIcon(IDI_POST2));
	m_pImgList->Add(AfxGetApp()->LoadIcon(IDI_POST3));
	m_pImgList->Add(AfxGetApp()->LoadIcon(IDI_POST4));
	m_Tree.SetImageList(m_pImgList,TVSIL_NORMAL);
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void FExecList::OnOK() 
{
	UINT uCount = m_Tree.GetCount();
	HTREEITEM hItem = m_Tree.GetNextItem(m_Tree.GetFirstVisibleItem(),TVGN_ROOT);
	DWORD dwExecId;
	BOOL bExist;
	Executer* exec;
	CString csVal;
	/*перенос исполнителей в основной массив*/
	for (UINT i=0;i < uCount && hItem;i++)
	{
		ASSERT(hItem != NULL);
		dwExecId=m_Tree.GetItemData(hItem);
		if(m_Tree.GetCheck(hItem))
		{
			/*поиск уже имеющегося*/
			bExist=FALSE;
			for(int j=0; j<(*m_paExecuters).GetSize(); ++j)
			{
				exec=(Executer*)((*m_paExecuters).GetAt(j));
				if(exec->dwPersonId ==dwExecId)
				{
					if(exec->state==2) exec->state=0; // если удаляли и снова отметили, восстанавливаем
					bExist=TRUE;
				}
			}
			if(!bExist)
			{
				exec=new Executer();
				csVal=m_Tree.GetItemText(hItem);
				exec->csName=csVal.Left(csVal.Find(":")-1);
				exec->dwPersonId=dwExecId;
				exec->scheduleDate.m_status=COleDateTime::null;
				exec->factDate.m_status=COleDateTime::null;
				exec->state=1;

				(*m_paExecuters).Add(exec);
			}

		}
		else // если не отмечен
		{
			/*поиск уже имеющегося*/
			for(int j=0; j<(*m_paExecuters).GetSize(); ++j)
			{
				exec=(Executer*)((*m_paExecuters).GetAt(j));
				
				if(exec->dwPersonId ==dwExecId)
					exec->state=2;
			}
		}
			
		hItem = m_Tree.GetNextItem(hItem,TVGN_NEXT);
	}
	
	CDialog::OnOK();
}

void FExecList::OnClose() 
{
	delete m_pImgList;
	
	CDialog::OnClose();
}
