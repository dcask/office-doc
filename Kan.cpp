// Kan.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "Kan.h"

#include "MainFrm.h"
#include "ChildFrm.h"
#include "KanDoc.h"
#include "KanView.h"
#include "FUserChoose.h"
#include "eh.h"
#include <locale.h>
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CKanApp

BEGIN_MESSAGE_MAP(CKanApp, CWinApp)
	//{{AFX_MSG_MAP(CKanApp)
	ON_COMMAND(ID_APP_ABOUT, OnAppAbout)
	ON_COMMAND(ID_FILE_NEW, OnFileNew)
	//}}AFX_MSG_MAP
	// Standard file based document commands
	ON_COMMAND(ID_FILE_NEW, CWinApp::OnFileNew)
	ON_COMMAND(ID_FILE_OPEN, CWinApp::OnFileOpen)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CKanApp construction

CKanApp::CKanApp()
{
	m_bEditIn=FALSE;
	m_bEditOut=FALSE;
	m_bReport=FALSE;
	m_bBaseTool=FALSE;
	m_bUniq=FALSE;
	m_bState=FALSE;
	m_bMessage=FALSE;

}

/////////////////////////////////////////////////////////////////////////////
// The one and only CKanApp object

CKanApp theApp;

// This identifier was generated to be statistically unique for your app.
// You may change it if you prefer to choose a specific identifier.

// {5CD806FA-472E-4A91-B651-822365B20DA8}
static const CLSID clsid =
{ 0x5cd806fa, 0x472e, 0x4a91, { 0xb6, 0x51, 0x82, 0x23, 0x65, 0xb2, 0xd, 0xa8 } };

/////////////////////////////////////////////////////////////////////////////
// CKanApp initialization

BOOL CKanApp::InitInstance()
{
	//Init base
	CString path1,path2,basepath;
	set_terminate(MegaFail);
	GetModuleFileName(GetModuleHandle( NULL), path1.GetBuffer(256), 256);
	path1.ReleaseBuffer();
			
	LPTSTR pszFileName = NULL;
	
	m_bOnlyMine=FALSE;

	GetFullPathName( path1, 250, path2.GetBuffer(256), &pszFileName);
	path2.ReleaseBuffer();
			
	path2=path1.Left(path1.Find(pszFileName)-1);
	m_csAppPath=path2;

	CFile f;
	if( !f.Open(m_csAppPath+"\\kan.ini",CFile::modeReadWrite))
	{
		AfxMessageBox("Не могу найти INI файл");
		return FALSE;
	}
	else f.Close();
	/*  Connection String  */
	GetPrivateProfileString("Settings","Path","",
			basepath.GetBuffer(512),512,m_csAppPath+"\\kan.ini");
	basepath.ReleaseBuffer();
	GetPrivateProfileString("Settings","Base","",
			path2.GetBuffer(512),512,m_csAppPath+"\\kan.ini");
	path2.ReleaseBuffer();
	if (basepath=="") path2.Replace("[%pathname%]",m_csAppPath);
	else path2.Replace("[%pathname%]",basepath);
	/*    Template    */
	GetPrivateProfileString("Settings","TemplSymb","*",
			m_TemplateSymbol.GetBuffer(2),2,m_csAppPath+"\\kan.ini");
	m_TemplateSymbol.ReleaseBuffer();
	/* IN_TABLE */
	GetPrivateProfileString("Settings","IN_TABLE","InLetter",
			m_csInLetter.GetBuffer(50),50,m_csAppPath+"\\kan.ini");
	m_csInLetter.ReleaseBuffer();
	/*   OUT_TABLE  */
	GetPrivateProfileString("Settings","OUT_TABLE","OutLetter",
			m_csOutLetter.GetBuffer(50),50,m_csAppPath+"\\kan.ini");
	m_csOutLetter.ReleaseBuffer();
	/*   EXEC_TABLE  */
	GetPrivateProfileString("Settings","EXEC_TABLE","Executor",
			m_csExecutor.GetBuffer(50),50,m_csAppPath+"\\kan.ini");
	m_csExecutor.ReleaseBuffer();
	/*    USER  */
	m_iUserID=GetPrivateProfileInt("Settings","User",0,m_csAppPath+"\\kan.ini");
	m_iXPStyle=GetPrivateProfileInt("Settings","XPStyle",1,m_csAppPath+"\\kan.ini");
	m_bt=NULL;
	m_bt=new CADOBaseTool(path2);
	if(!m_bt)
	{
		AfxMessageBox("Ошибка открытия базы. ");
		return FALSE;
	}
	m_bt->ShowError(TRUE);
	/*региональные настройки*/
	HKEY Key;
	char buf[256];
	DWORD dws = 0x100;
	DWORD dwt = 0;
	LONG result = 0;

	result = RegOpenKeyEx(HKEY_CURRENT_USER, "Control Panel\\International", 0, KEY_ALL_ACCESS, &Key);
	
	RegQueryValueEx(Key, "sShortDate",0, 0, (LPBYTE)buf, &dws);
	m_csFormatDate=buf;
	// Initialize OLE libraries
	if (!AfxOleInit())
	{
		AfxMessageBox(IDP_OLE_INIT_FAILED);
		return FALSE;
	}

	AfxEnableControlContainer();

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	//  of your final executable, you should remove from the following
	//  the specific initialization routines you do not need.

#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

	// Change the registry key under which our settings are stored.
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization.
	SetRegistryKey(_T("Local AppWizard-Generated Applications"));

	LoadStdProfileSettings();  // Load standard INI file options (including MRU)
	setlocale( LC_ALL, "Russian" ); 
	// Register the application's document templates.  Document templates
	//  serve as the connection between documents, frame windows and views.
	//RegisterLib();
	CMultiDocTemplate* pDocTemplateMain;
	/*CMultiDocTemplate* pDocTemplateIn;
	CMultiDocTemplate* pDocTemplateOut;*/

	pDocTemplateMain = new CMultiDocTemplate(
		IDR_KANTYPE,
		RUNTIME_CLASS(CKanDoc),
		RUNTIME_CLASS(CChildFrame), // custom MDI child frame
		RUNTIME_CLASS(CKanView));
	AddDocTemplate(pDocTemplateMain);

	/*pDocTemplateOut = new CMultiDocTemplate(
		IDR_OUTTYPE,
		RUNTIME_CLASS(CKanDoc),
		RUNTIME_CLASS(CChildFrame), // custom MDI child frame
		RUNTIME_CLASS(FOutLetter));
	AddDocTemplate(pDocTemplateOut);

	pDocTemplateIn = new CMultiDocTemplate(
		IDR_INTYPE,
		RUNTIME_CLASS(CKanDoc),
		RUNTIME_CLASS(CChildFrame), // custom MDI child frame
		RUNTIME_CLASS(FInLetter));
	AddDocTemplate(pDocTemplateIn);*/

	// Connect the COleTemplateServer to the document template.
	//  The COleTemplateServer creates new documents on behalf
	//  of requesting OLE containers by using information
	//  specified in the document template.
	m_server.ConnectTemplate(clsid, pDocTemplateMain, FALSE);

	// Register all OLE server factories as running.  This enables the
	//  OLE libraries to create objects from other applications.
	COleTemplateServer::RegisterAll();
		// Note: MDI applications register all server objects without regard
		//  to the /Embedding or /Automation on the command line.
	m_csDocType="Main";
	// create main MDI Frame window
	CMainFrame* pMainFrame = new CMainFrame;
	if (!pMainFrame->LoadFrame(IDR_MAINFRAME))
		return FALSE;
	m_pMainWnd = pMainFrame;

	// Parse command line for standard shell commands, DDE, file open
	CCommandLineInfo cmdInfo;
	ParseCommandLine(cmdInfo);

	// Check to see if launched as OLE server
	if (cmdInfo.m_bRunEmbedded || cmdInfo.m_bRunAutomated)
	{
		// Application was run with /Embedding or /Automation.  Don't show the
		//  main window in this case.
		return TRUE;
	}

	// When a server application is launched stand-alone, it is a good idea
	//  to update the system registry in case it has been damaged.
	m_server.UpdateRegistry(OAT_DISPATCH_OBJECT);
	COleObjectFactory::UpdateRegistryAll();

	// Dispatch commands specified on the command line
	if (!ProcessShellCommand(cmdInfo))
		return FALSE;

	// The main window has been initialized, so show and update it.
	pMainFrame->ShowWindow(SW_SHOWMAXIMIZED);
	pMainFrame->UpdateWindow();
	FUserChoose dlg;
	/* Пользователь*/
	CString csVal;
	dlg.m_iID=m_iUserID;
	while(dlg.DoModal()!=IDOK);
	m_iUserID=dlg.m_iID;
	csVal.Format("%d",dlg.m_iID);
	WritePrivateProfileString("Settings","User",csVal,m_csAppPath+"\\kan.ini");

	/*Запустить сессию*/
	CMDIFrameWnd *pFrame = (CMDIFrameWnd*)AfxGetApp()->m_pMainWnd;
	CMDIChildWnd *pChild = (CMDIChildWnd *) pFrame->GetActiveFrame();
	((CKanView*) pChild->GetActiveView())->StartSession();
	/*Активация пункта меню*/
	CMenu* menu=AfxGetApp()->GetMainWnd()->GetMenu();
	CMenu* submenu=menu->GetSubMenu(4);
	if(m_iUserID!=0) submenu->EnableMenuItem(3,MF_BYPOSITION | MF_GRAYED);
	return TRUE;
}


/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
		// No message handlers
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

// App command to run the dialog
void CKanApp::OnAppAbout()
{
	CAboutDlg aboutDlg;
	aboutDlg.DoModal();
}

/////////////////////////////////////////////////////////////////////////////
// CKanApp message handlers


void CKanApp::OnFileNew() 
{
	OpenNewDoc(m_csDocType);	
	
}

BOOL CKanApp::OpenNewDoc(const CString &strTarget)
{
	CString strDocName;
	CString winName="Unnamed";
	if(strTarget=="Main") winName="Главное"; 
	if(strTarget=="In") winName="Входящее письмо"; 
	if(strTarget=="Out") winName="Исходящее письмо"; 
	
	//((CMDIFrameWnd*)m_pMainWnd)->ge;
	
	CDocTemplate* pSelectedTemplate;
	POSITION pos = GetFirstDocTemplatePosition();
	while (pos != NULL)
	{
		pSelectedTemplate = (CDocTemplate*) GetNextDocTemplate(pos);
		pSelectedTemplate->GetDocString(strDocName, CDocTemplate::docName);
		if (strDocName == strTarget) // выбирается из строкового ресурса шаблона
		{
			CDocument* doc=pSelectedTemplate->OpenDocumentFile(NULL);
			doc->SetTitle(winName);
			return TRUE;
		}
	}
	return FALSE;
}

int CKanApp::ExitInstance() 
{
	if(m_bt)
	{
		if(m_bt->GetRS()->IsOpen()) m_bt->GetRS()->Close();
		if(m_bt->GetDB()->IsOpen()) m_bt->GetDB()->Close();
		delete m_bt;
	}
	return CWinApp::ExitInstance();
}

void CKanApp::RegisterLib()
{
	/*регистрация*/
	HKEY Key;
	DWORD dws = 0x100;
	DWORD dwt = 0;
	LONG result = 0;

	result = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Shared Tools\\msflxgrd.ocx", 0, KEY_ALL_ACCESS, &Key);
	if(result!=ERROR_SUCCESS)
	{
		CString pszDllName=m_csAppPath+"\\MSFLXGRD.OCX";
		FARPROC lpDllEntryPoint;
		HINSTANCE hLib = LoadLibrary(pszDllName);
		if (hLib < (HINSTANCE)HINSTANCE_ERROR)
			AfxMessageBox("Не могу открыть MSFLXGRD.OCX "+pszDllName);

		// Find the entry point.
		lpDllEntryPoint = GetProcAddress(hLib,_T("DllRegisterServer"));
		if (lpDllEntryPoint != NULL) (*lpDllEntryPoint)();
		else AfxMessageBox("Не найдена точка входа "+pszDllName);
	}
	RegCloseKey(Key);
	result = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Shared Tools\\msmask32.ocx", 0, KEY_ALL_ACCESS, &Key);
	if(result!=ERROR_SUCCESS)
	{
		CString pszDllName=m_csAppPath+"\\MSMASK32.OCX";
		FARPROC lpDllEntryPoint;
		HINSTANCE hLib = LoadLibrary(pszDllName);
		if (hLib < (HINSTANCE)HINSTANCE_ERROR)
			AfxMessageBox("Не могу открыть MSMASK32.OCX "+pszDllName);

		// Find the entry point.
		(FARPROC&)lpDllEntryPoint = GetProcAddress(hLib,_T("DllRegisterServer"));
		if (lpDllEntryPoint != NULL) (*lpDllEntryPoint)();
		else AfxMessageBox("Не найдена точка входа "+pszDllName);
	}
	RegCloseKey(Key);

}

