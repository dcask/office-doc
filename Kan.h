// Kan.h : main header file for the KAN application
//

#if !defined(AFX_KAN_H__67C01A05_E57A_4637_83AA_EEDA44B0202A__INCLUDED_)
#define AFX_KAN_H__67C01A05_E57A_4637_83AA_EEDA44B0202A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CKanApp:
// See Kan.cpp for the implementation of this class
//

class CKanApp : public CWinApp
{
public:
	CString m_csDocType;
	long	m_lLetterId;
	int		m_iLetterType;
	int		m_iXPStyle;
	CADOBaseTool*	m_bt;
	CString m_csAppPath;
	CString m_csFormatDate;
	CString m_TemplateSymbol;
	CString m_csInLetter;
	CString m_csExecutor;
	CString m_csOutLetter;
	BOOL	m_bEditIn;
	BOOL	m_bEditOut;
	BOOL	m_bMessage;
	BOOL	m_bReport;
	BOOL	m_bBaseTool;
	BOOL	m_bUniq;
	BOOL	m_bState;
	int		m_iUserID;
	BOOL	m_bOnlyMine;
public:
	void RegisterLib();
	BOOL OpenNewDoc(const CString &strTarget);
	CKanApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKanApp)
	public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();
	//}}AFX_VIRTUAL

// Implementation
	COleTemplateServer m_server;
		// Server object for document creation
	//{{AFX_MSG(CKanApp)
	afx_msg void OnAppAbout();
	afx_msg void OnFileNew();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KAN_H__67C01A05_E57A_4637_83AA_EEDA44B0202A__INCLUDED_)
