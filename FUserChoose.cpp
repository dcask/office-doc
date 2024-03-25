// FUserChoose.cpp : implementation file
//

#include "stdafx.h"
#include "Kan.h"
//#include "KanView.h"
#include "FUserChoose.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FUserChoose dialog


FUserChoose::FUserChoose(CWnd* pParent /*=NULL*/)
	: CDialog(FUserChoose::IDD, pParent)
{
	//{{AFX_DATA_INIT(FUserChoose)
	m_csUser = _T("");
	//}}AFX_DATA_INIT
	m_iID=0;
}


void FUserChoose::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FUserChoose)
	DDX_CBString(pDX, IDC_USER, m_csUser);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FUserChoose, CDialog)
	//{{AFX_MSG_MAP(FUserChoose)
	ON_WM_SHOWWINDOW()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FUserChoose message handlers

void FUserChoose::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);
	
	if(!bShow) return;

	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CComboBox* cb=(CComboBox* ) GetDlgItem(IDC_USER);
	bt->LoadComboBox(cb,"[ComboUser]");
	for(int i=0; i< cb->GetCount(); i++ )
		if(cb->GetItemData(i)==UINT(m_iID))
		{
			cb->SetCurSel(i);
			break;
		}
}

void FUserChoose::OnOK() 
{
	
	CComboBox* cb=(CComboBox*) GetDlgItem(IDC_USER);
	CString userName, currentUserName;
	DWORD rights=0;
	DWORD len=100;
	int pos=cb->GetCurSel();
	if(pos<0)
	{
		AfxMessageBox("������������ �� ����������",MB_ICONSTOP);
		return;
	}
	m_iID=cb->GetItemData(pos);

	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	/* ������� ���������� ������*/
	CADOCommand pCmdClear(bt->GetDB(), "[ClearActiveUsers]");
	try
	{
		bt->GetRS()->Execute(&pCmdClear);
	}catch(CADOException* e)
	{	
		bt->ThrowError(e);
		return;
	}
	bt->GetRS()->Close();
	/*�������� "������"*/
	BOOL res=GetUserName(currentUserName.GetBuffer(100),&len);
	currentUserName.ReleaseBuffer();
	CADOCommand pCmdPassword(bt->GetDB(), "[GetUserName]");
	pCmdPassword.AddParameter("UserID",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),m_iID);
	try
	{
		bt->GetRS()->Execute(&pCmdPassword);
	}catch(CADOException* e)
	{	
		bt->ThrowError(e);
		return;
	}
	
	userName=bt->GetStringValue("UserName");
	bt->GetRS()->GetFieldValue("UserRights",rights);
	((CKanApp*)AfxGetApp())->m_bUniq=rights & 0x1L; // ���� ��������� � ����
	((CKanApp*)AfxGetApp())->m_bEditIn=rights & 0x2L; // ������ � �������������� ��������
	((CKanApp*)AfxGetApp())->m_bEditOut=rights & 0x4L; // ������ � �������������� ���������
	((CKanApp*)AfxGetApp())->m_bReport=rights & 0x8L; // ������ � �������
	((CKanApp*)AfxGetApp())->m_bBaseTool=rights & 0x10L; //������ � �������������� ����
	((CKanApp*)AfxGetApp())->m_bState=rights & 0x20L; //������ � ������� ����������
	((CKanApp*)AfxGetApp())->m_bMessage=rights & 0x40L; //������ � �������� ���������

	bt->GetRS()->Close();
	if(userName!=currentUserName && userName!="")
	{
		AfxMessageBox("���� ������������ �� �����\n �������� �� ������ ���������� ",MB_ICONSTOP);
		return;
	}
	/*��������*/
	CADOCommand pCmd(bt->GetDB(), "[GetActiveUser]");
	pCmd.AddParameter("UserID",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),m_iID);
	try
	{
		bt->GetRS()->Execute(&pCmd);
	}catch(CADOException* e)
	{	
		bt->ThrowError(e);
		return;
	}

	if(bt->GetRecordsCount()>0 &&	((CKanApp*)AfxGetApp())->m_bUniq)
	{
			bt->GetRS()->Close();
			AfxMessageBox("������������ ��� ���������.", MB_ICONSTOP);
			return;
	}
	bt->GetRS()->Close();
	CDialog::OnOK();
}
