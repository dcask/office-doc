// KanDoc.cpp : implementation of the CKanDoc class
//

#include "stdafx.h"
#include "Kan.h"

#include "KanDoc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CKanDoc

IMPLEMENT_DYNCREATE(CKanDoc, CDocument)

BEGIN_MESSAGE_MAP(CKanDoc, CDocument)
	//{{AFX_MSG_MAP(CKanDoc)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CKanDoc, CDocument)
	//{{AFX_DISPATCH_MAP(CKanDoc)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//      DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IKan to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {88FF54F9-C87F-4E73-800B-6B2069BEE433}
static const IID IID_IKan =
{ 0x88ff54f9, 0xc87f, 0x4e73, { 0x80, 0xb, 0x6b, 0x20, 0x69, 0xbe, 0xe4, 0x33 } };

BEGIN_INTERFACE_MAP(CKanDoc, CDocument)
	INTERFACE_PART(CKanDoc, IID_IKan, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CKanDoc construction/destruction

CKanDoc::CKanDoc()
{
	m_pTabDialog=NULL;
	EnableAutomation();
	AfxOleLockApp();
}

CKanDoc::~CKanDoc()
{
	AfxOleUnlockApp();
}

BOOL CKanDoc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;
	CString newTitle="Unnamed";
	CString title=GetTitle();
	if(title.Find("Main")>=0) newTitle="Главное";
	if(title.Find("In")>=0) newTitle="Входящее письмо";
	if(title.Find("Out")>=0) newTitle="Исходящее письмо";
	SetPathName(newTitle,FALSE);
	m_lLetterId=((CKanApp*)AfxGetApp())->m_lLetterId;
	return TRUE;
}



/////////////////////////////////////////////////////////////////////////////
// CKanDoc serialization

void CKanDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}

/////////////////////////////////////////////////////////////////////////////
// CKanDoc diagnostics

#ifdef _DEBUG
void CKanDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CKanDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CKanDoc commands

void CKanDoc::OnCloseDocument() 
{
	for(int i=0; i< m_aExecuters.GetSize(); i++)
	{
		delete ((Executer*)m_aExecuters.GetAt(i));
	}
	for(i=0; i< m_aThemes.GetSize(); i++)
	{
		delete ((Theme*)m_aThemes.GetAt(i));
	}
	for(i=0; i< m_aExtNum.GetSize(); i++)
	{
		delete ((InNum*)m_aExtNum.GetAt(i));
	}
	for(i=0; i< m_aReplyOn.GetSize(); i++)
	{
		delete ((Reply*)m_aReplyOn.GetAt(i));
	}
	/* Очищаем невалидные сессии*/
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmdClear(bt->GetDB(), "[DeleteActiveUser]");
	pCmdClear.AddParameter("User_ID",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),((CKanApp*)AfxGetApp())->m_iUserID);
	try
	{
		bt->GetRS()->Execute(&pCmdClear);
	}catch(CADOException* e)
	{	
		bt->ThrowError(e);
		return;
	}
	bt->GetRS()->Close();
	CDocument::OnCloseDocument();
}

BOOL CKanDoc::LoadInData()
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;

	/*Очищение*/
	for(UINT i=0; i< UINT(m_aExecuters.GetSize()); i++)
		delete ((Executer*)m_aExecuters.GetAt(i));
	for(i=0; i< UINT(m_aThemes.GetSize()); i++)
		delete ((Theme*)m_aThemes.GetAt(i));
	for(i=0; i< UINT(m_aExtNum.GetSize()); i++)
		delete ((InNum*)m_aExtNum.GetAt(i));
	m_aExecuters.RemoveAll();
	m_aThemes.RemoveAll();
	m_aExtNum.RemoveAll();

	m_csRegNum=_T("");
	m_RegDate.m_status=COleDateTime::null;
	m_csDeclarant=_T("");
	m_csDeclarantType=_T("");
	m_iDeclarantType=0;
	m_csUse=_T("");
	m_iUse=0;
	m_csDistrict=_T("");
	m_iDistrict=0;
	m_csFrom=_T("");
	m_iFrom=0;
	m_csOutRegNum_F=_T("");
	m_OutRegDate_F.m_status=COleDateTime::null;
	m_csContent=_T("");
	m_csAuthor=_T("");
	m_iAuthor=0;
	m_ResolutionDate.m_status=COleDateTime::null;
	m_csResolution=_T("");
	m_csExecCourse=_T("");
	m_csControlType=_T("");
	m_iControlType=0;
	m_csResult=_T("");
	m_iResult=0;
	m_csTheme=_T("");
	m_iTheme=0;
	m_iAttrib=0;
	/**/
	if(m_lLetterId==0) return TRUE;
	CADOCommand pCmd(bt->GetDB(), "[GetIn]");

	CString csVal;
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	bt->SetDateFormat("%x");
			
	pCmd.AddParameter("InID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lLetterId);
	
	//запрос
	try
	{
		if(!bt->GetRS()->Execute(&pCmd)) return FALSE;
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		return FALSE;
	}
	
	if(bt->GetRecordsCount()<1) // не найден
	{
		AfxMessageBox("Не найден");
		::SendMessage(AfxGetApp()->GetMainWnd()->GetSafeHwnd(),WM_COMMAND,ID_FILE_CLOSE,NULL);
	}
	m_csRegNum=bt->GetStringValue("RegNum");
	m_csFullNum=bt->GetStringValue("RegFull");
	bt->GetRS()->GetFieldValue("RegDate",m_RegDate);
	bt->GetRS()->GetFieldValue("ProcDate",m_ProcessDate);
	m_csDeclarant=bt->GetStringValue("tDeclarant");
	m_csDeclarantType=bt->GetStringValue("tDeclarantType");
	bt->GetRS()->GetFieldValue("iDeclarantType",m_iDeclarantType);
	m_csUse=bt->GetStringValue("tUse");
	bt->GetRS()->GetFieldValue("iUse",m_iUse);
	m_csDistrict=bt->GetStringValue("tDistrict");
	bt->GetRS()->GetFieldValue("iDistrict",m_iDistrict);
	m_csFrom=bt->GetStringValue("tFrom");
	bt->GetRS()->GetFieldValue("iFrom",m_iFrom);
	m_csContent=bt->GetStringValue("tContent");
	m_csAuthor=bt->GetStringValue("Author");
	bt->GetRS()->GetFieldValue("iAuthor",m_iAuthor);
	bt->GetRS()->GetFieldValue("ResolutionDate",m_ResolutionDate);
	m_csResolution=bt->GetStringValue("tResolution");
	m_csExecCourse=bt->GetStringValue("ExecCourse");
	m_csControlType=bt->GetStringValue("tControlType");
	bt->GetRS()->GetFieldValue("iControlType",m_iControlType);
	m_csResult=bt->GetStringValue("tResult");
	bt->GetRS()->GetFieldValue("iResult",m_iResult);
	m_csTheme=bt->GetStringValue("tTheme");
	bt->GetRS()->GetFieldValue("iTheme",m_iTheme);
	bt->GetRS()->GetFieldValue("Attrib",m_iAttrib);
	bt->GetRS()->Close();
	/*Исполнители*/
	CADOCommand pCmd1(bt->GetDB(), "[GetExecuters]");
	pCmd1.AddParameter("InID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lLetterId);
	//запрос
	try
	{
		if(!bt->GetRS()->Execute(&pCmd1)) return FALSE;
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		return FALSE;
	}
	Executer* exec;
	if(bt->GetRecordsCount()>0) bt->GetRS()->MoveFirst();
	for(i=0; i<bt->GetRecordsCount(); ++i)
	{
		exec=new Executer();
		bt->GetRS()->GetFieldValue("PersonID",exec->dwPersonId);
		bt->GetRS()->GetFieldValue("ID",exec->dwID);
		exec->csName=bt->GetStringValue("Name");
		bt->GetRS()->GetFieldValue("ScheduleDate",exec->scheduleDate);
		bt->GetRS()->GetFieldValue("FactDate",exec->factDate);
		exec->dwInLetterId=m_lLetterId;
		exec->state=0;
		m_aExecuters.Add(exec);
		bt->GetRS()->MoveNext();
	}
	bt->GetRS()->Close();
	CADOCommand pCmd2(bt->GetDB(), "[GetInNum]");
		
	pCmd2.AddParameter("InID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lLetterId);
	bt->Execute(&pCmd2);
	InNum* innum;
	for(i=0; i<bt->GetRecordsCount(); ++i)
	{
		innum=new InNum();
		bt->GetRS()->GetFieldValue("ID",innum->dwId);
		innum->csRegNum=bt->GetStringValue("RegNum");
		bt->GetRS()->GetFieldValue("RegDate",innum->RegDate);
		innum->state=0;
		m_aExtNum.Add(innum);
		bt->GetRS()->MoveNext();
	}
	bt->GetRS()->Close();
	return TRUE;
}

BOOL CKanDoc::LoadOutData()
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;

	/*Очищение*/
	for(UINT i=0; i< UINT(m_aExecuters.GetSize()); i++)
		delete ((Executer*)m_aExecuters.GetAt(i));
	for(i=0; i< UINT(m_aReplyOn.GetSize()); i++)
		delete ((Reply*)m_aReplyOn.GetAt(i));
	for(i=0; i< UINT(m_aThemes.GetSize()); i++)
		delete ((Theme*)m_aThemes.GetAt(i));
	m_aExecuters.RemoveAll();
	m_aThemes.RemoveAll();
	m_aReplyOn.RemoveAll();

	m_csRegNum=_T("");
	m_RegDate.m_status=COleDateTime::null;
	m_csRegPrefix=_T("");
	m_csUse=_T("");
	m_iUse=0;
	m_lInLetterID=0;
	m_csContent=_T("");
	m_csAuthor=_T("");
	m_iAuthor=0;
	m_csTheme=_T("");
	m_csToPlace=_T("");
	m_iTheme=0;
	m_csOutRegNum=_T("");
	m_OutRegDate.m_status=COleDateTime::null;
	m_csOutRegNum_F=_T("");
	m_OutRegDate_F.m_status=COleDateTime::null;
	/**/
	if(m_lLetterId==0) return TRUE;
	CADOCommand pCmd(bt->GetDB(), "[GetOut]");
	//CADOFieldInfo info;
	CString csVal;
	CString csAppPath=((CKanApp*) AfxGetApp())->m_csAppPath;
	bt->SetDateFormat("%x");
			
	pCmd.AddParameter("OutID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lLetterId);
	
	//запрос
	try
	{
		if(!bt->GetRS()->Execute(&pCmd)) return FALSE;
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		return FALSE;
	}
	
	if(bt->GetRecordsCount()<1) // не найден
	{
		AfxMessageBox("Не найден");
		::SendMessage(AfxGetApp()->GetMainWnd()->GetSafeHwnd(),WM_COMMAND,ID_FILE_CLOSE,NULL);

	}
	m_csRegNum=bt->GetStringValue("RegNum");
	bt->GetRS()->GetFieldValue("RegDate",m_RegDate);
	m_csRegPrefix=bt->GetStringValue("RegPrefix");
	m_csUse=bt->GetStringValue("tUse");
	bt->GetRS()->GetFieldValue("iUse",m_iUse);
	bt->GetRS()->GetFieldValue("InLetterID",m_lInLetterID);
	m_csContent=bt->GetStringValue("tContent");
	m_csAuthor=bt->GetStringValue("tAuthor");
	bt->GetRS()->GetFieldValue("iAuthor",m_iAuthor);
	bt->GetRS()->GetFieldValue("ToPlace",m_csToPlace);
	m_csTheme=bt->GetStringValue("tTheme");
	bt->GetRS()->GetFieldValue("iTheme",m_iTheme);
	m_csOutRegNum=bt->GetStringValue("InRegNum");
	bt->GetRS()->GetFieldValue("InRegDate",m_OutRegDate);
	m_csOutRegNum_F=bt->GetStringValue("ExtInRegNum");
	bt->GetRS()->GetFieldValue("ExtInRegDate",m_OutRegDate_F);
	bt->GetRS()->Close();
	
	CADOCommand pCmd1(bt->GetDB(), "[GetOutNum]");
		
	pCmd1.AddParameter("OutID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lLetterId);
	bt->Execute(&pCmd1);
	Reply* reply;
	for(i=0; i<bt->GetRecordsCount(); ++i)
	{
		reply=new Reply();
		bt->GetRS()->GetFieldValue("ID",reply->dwId);
		bt->GetRS()->GetFieldValue("InID",reply->inLetterID);
		reply->RegNum=bt->GetStringValue("RegNum");
		bt->GetRS()->GetFieldValue("RegYear",reply->year);
		reply->state=0;
		m_aReplyOn.Add(reply);
		bt->GetRS()->MoveNext();
	}
	bt->GetRS()->Close();
	return TRUE;
}

BOOL CKanDoc::SaveOutData()
{

	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmdOut(bt->GetDB(), "[UpdateOutLetter]");
	
	pCmdOut.AddParameter("[Prefix]",CADORecordset::typeChar,
				CADOParameter::paramInput,100,m_csRegPrefix);
	
	pCmdOut.AddParameter("[RegNum]",CADORecordset::typeChar,
				CADOParameter::paramInput,50,m_csRegNum);

	pCmdOut.AddParameter("[:RegDate]",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),m_RegDate);

	pCmdOut.AddParameter("[:iAuthor]",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_iAuthor);

	pCmdOut.AddParameter("[:InID]",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lInLetterID);
	
	pCmdOut.AddParameter("[:Place]",CADORecordset::typeBSTR,
				CADOParameter::paramInput,1000,m_csToPlace);

	pCmdOut.AddParameter("[:tContent]",CADORecordset::typeBSTR,
				CADOParameter::paramInput,1000,m_csContent);

	pCmdOut.AddParameter("[:UserID]",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),((CKanApp*) AfxGetApp())->m_iUserID);

	pCmdOut.AddParameter("[:OutID]",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lLetterId);

	try
	{
		if(!(bt->GetRS()->Execute(&pCmdOut))) 
		{
			AfxMessageBox("Ошибка записи");
			return FALSE;
		}
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		AfxMessageBox("Ошибка записи");
		return FALSE;
	}
	bt->GetRS()->Close();
	Executer*  ex;
	//int count=0; // количество закрытых исполнителей

	for(UINT i=0; i< UINT(m_aExecuters.GetSize()); ++i)
	{
			ex=(Executer*) (m_aExecuters.GetAt(i));
			CADOCommand pCmdEx(bt->GetDB(), "[UpdateExecuters]");
			pCmdEx.AddParameter("InID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),ex->dwInLetterId);
			pCmdEx.AddParameter("PersonID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),long(ex->dwPersonId));
			if(ex->scheduleDate.m_status==COleDateTime::null)
			{
				ex->scheduleDate.m_dt=0.0;
				ex->scheduleDate.m_status=COleDateTime::valid;
			}
			pCmdEx.AddParameter("Schedule",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),ex->scheduleDate);
			if(ex->factDate.m_status==COleDateTime::null)
			{
				ex->factDate.m_dt=0.0;
				ex->factDate.m_status=COleDateTime::valid;
			}
			pCmdEx.AddParameter("Fact",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),ex->factDate);
			pCmdEx.AddParameter("UserID",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),((CKanApp*) AfxGetApp())->m_iUserID);
			pCmdEx.AddParameter("ExID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),long(ex->dwID));
			bt->Execute(&pCmdEx);
			bt->GetRS()->Close();
	}

	// номера
	Reply* reply;
	for(i=0; i< UINT(m_aReplyOn.GetSize()); ++i)
	{
		//////////////// insert
			reply=(Reply*) (m_aReplyOn.GetAt(i));
			CADOCommand pCmdReplyAdd(bt->GetDB(), "[AddOutNum]");
			CADOCommand pCmdReplyDel(bt->GetDB(), "[DelOutNum]");
			pCmdReplyAdd.AddParameter("OutID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lLetterId);
			pCmdReplyAdd.AddParameter("InID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),long(reply->inLetterID));
			pCmdReplyAdd.AddParameter("UserID",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),((CKanApp*) AfxGetApp())->m_iUserID);
			pCmdReplyDel.AddParameter("_ID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),long(reply->dwId));
			pCmdReplyDel.AddParameter("OutID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),long(-100));
			if(reply->state==1)
			{
				bt->Execute(&pCmdReplyAdd);
				CADOCommand pCmdIn(bt->GetDB(), "[SetReply]");
				pCmdIn.AddParameter("[Attr]",CADORecordset::typeBigInt,
					CADOParameter::paramInput,sizeof(long),long(1));
				pCmdIn.AddParameter("[InID]",CADORecordset::typeBigInt,
					CADOParameter::paramInput,sizeof(long),reply->inLetterID);
				bt->Execute(&pCmdIn);
				bt->GetRS()->Close();
			}
			if(reply->state==2)
			{
				bt->Execute(&pCmdReplyDel);
				CADOCommand pCmdIn(bt->GetDB(), "[SetReply]");
				pCmdIn.AddParameter("[Attr]",CADORecordset::typeBigInt,
					CADOParameter::paramInput,sizeof(long),long(0));
				pCmdIn.AddParameter("[InID]",CADORecordset::typeBigInt,
					CADOParameter::paramInput,sizeof(long),reply->inLetterID);
				bt->Execute(&pCmdIn);
				bt->GetRS()->Close();
			}
			// признак
			
	}
	return TRUE;

}


BOOL CKanDoc::SaveInData()
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	CADOCommand pCmdIn(bt->GetDB(), "[UpdateInLetter]");

	pCmdIn.AddParameter("[RegNum]",CADORecordset::typeChar,
				CADOParameter::paramInput,50,m_csRegNum);

	pCmdIn.AddParameter("[RegFull]",CADORecordset::typeChar,
				CADOParameter::paramInput,50,m_csFullNum);

	pCmdIn.AddParameter("[ProcDate]",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),m_ProcessDate);

	pCmdIn.AddParameter("[RegDate]",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),m_RegDate);

	pCmdIn.AddParameter("[Decl]",CADORecordset::typeChar,
				CADOParameter::paramInput,200,m_csDeclarant.Left(199));

	pCmdIn.AddParameter("[DeclTypeID]",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),m_iDeclarantType);

	pCmdIn.AddParameter("[OutRegNum]",CADORecordset::typeChar,
				CADOParameter::paramInput,50,m_csOutRegNum_F);

	pCmdIn.AddParameter("[OutRegDate]",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),m_OutRegDate_F);

	pCmdIn.AddParameter("[tContent]",CADORecordset::typeBSTR,
				CADOParameter::paramInput,1000,m_csContent.Left(999));

	pCmdIn.AddParameter("[Control]",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),m_iControlType);

	pCmdIn.AddParameter("[Res]",CADORecordset::typeBSTR,
				CADOParameter::paramInput,1000,m_csResolution.Left(999));

	pCmdIn.AddParameter("[ResAuthor]",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),m_iAuthor);

	pCmdIn.AddParameter("[ResDate]",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),m_ResolutionDate);

	pCmdIn.AddParameter("[UserID]",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),((CKanApp*) AfxGetApp())->m_iUserID);

	pCmdIn.AddParameter("[InID]",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lLetterId);

	try
	{
		if(!(bt->GetRS()->Execute(&pCmdIn))) 
		{
			AfxMessageBox("Ошибка записи");
			return FALSE;
		}
	}catch(CADOException* e)
	{
		bt->ThrowError(e);
		AfxMessageBox("Ошибка записи");
		return FALSE;
	}
	bt->GetRS()->Close();
	Executer*  ex;

	for(UINT i=0; i< UINT(m_aExecuters.GetSize()); ++i)
	{
		ex=(Executer*) (m_aExecuters.GetAt(i));
		if(ex->state<2) // обновить
		{
			if(ex->state==1) // новый
			{
				bt->GetRS()->Open(((CKanApp*)AfxGetApp())->m_csExecutor,CADORecordset::openTable);
				bt->GetRS()->AddNew();
				bt->GetRS()->SetFieldValue(1,0);
				bt->GetRS()->Update();
				bt->GetRS()->GetFieldValue("ID",ex->dwID);
			}
			CADOCommand pCmdEx(bt->GetDB(), "[UpdateExecuters]");			
			pCmdEx.AddParameter("InID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lLetterId);
			pCmdEx.AddParameter("PersonID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),long(ex->dwPersonId));
			if(ex->scheduleDate.m_status==COleDateTime::null)
			{
				ex->scheduleDate.m_dt=0.0;
				ex->scheduleDate.m_status=COleDateTime::valid;
			}
			pCmdEx.AddParameter("Schedule",CADORecordset::typeDate,
					CADOParameter::paramInput,sizeof(COleDateTime),ex->scheduleDate);
			if(ex->factDate.m_status==COleDateTime::null)
			{
				ex->factDate.m_dt=0.0;
				ex->factDate.m_status=COleDateTime::valid;
			}
			pCmdEx.AddParameter("Fact",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime),ex->factDate);
			pCmdEx.AddParameter("UserID",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),((CKanApp*) AfxGetApp())->m_iUserID);
			pCmdEx.AddParameter("ExID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),long(ex->dwID));
			try
			{
				bt->GetRS()->Execute(&pCmdEx);
			}catch(CADOException* e)
			{
				bt->ThrowError(e);
				return FALSE;
			}
			bt->GetRS()->Close();
		}
		if(ex->state==2) // удаляем
		{
			CADOCommand pCmdEx(bt->GetDB(), "[ExDelete]");
			pCmdEx.AddParameter("InID",CADORecordset::typeBigInt,
						CADOParameter::paramInput,sizeof(long),-1);
			pCmdEx.AddParameter("ExID",CADORecordset::typeBigInt,
						CADOParameter::paramInput,sizeof(long),long(ex->dwID));
			try
			{
				bt->GetRS()->Execute(&pCmdEx);
			}catch(CADOException* e)
			{	
				bt->ThrowError(e);
				return FALSE;
			}
			bt->GetRS()->Close();
		}
		
	}
	// номера
	InNum* innum;
	for(i=0; i< UINT(m_aExtNum.GetSize()); ++i)
	{
			innum=(InNum*) (m_aExtNum.GetAt(i));
			CADOCommand pCmdNumAdd(bt->GetDB(), "[AddInNum]");
			CADOCommand pCmdNumDel(bt->GetDB(), "[DelInNum]");
			pCmdNumAdd.AddParameter("InID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),m_lLetterId);
			pCmdNumAdd.AddParameter("_Num",CADORecordset::typeChar,
				CADOParameter::paramInput,49, innum->csRegNum.Left(49));
			pCmdNumAdd.AddParameter("_Date",CADORecordset::typeDate,
				CADOParameter::paramInput,sizeof(COleDateTime), innum->RegDate);
			pCmdNumAdd.AddParameter("UserID",CADORecordset::typeInteger,
				CADOParameter::paramInput,sizeof(int),((CKanApp*) AfxGetApp())->m_iUserID);
			pCmdNumDel.AddParameter("_ID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),long(innum->dwId));
			pCmdNumDel.AddParameter("InID",CADORecordset::typeBigInt,
				CADOParameter::paramInput,sizeof(long),long(-100));
			if(innum->state==1) bt->Execute(&pCmdNumAdd);
			if(innum->state==2) bt->Execute(&pCmdNumDel);
			bt->GetRS()->Close();
	}
	return TRUE;
}