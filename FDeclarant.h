#if !defined(AFX_FDECLARANT_H__A4045BD9_5D04_4B1C_8D35_032ECF8D8C54__INCLUDED_)
#define AFX_FDECLARANT_H__A4045BD9_5D04_4B1C_8D35_032ECF8D8C54__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FDeclarant.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FDeclarant dialog

class FDeclarant : public CDialog
{
public:
	CString m_csString;
	int		m_iDeclType;
	CString m_csQuery;
	CString m_csWindowName;
// Construction
public:
	FDeclarant(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FDeclarant)
	enum { IDD = IDD_DECLARANT };
	CListBox	m_List;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FDeclarant)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FDeclarant)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	virtual void OnOK();
	afx_msg void OnDblclkList();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FDECLARANT_H__A4045BD9_5D04_4B1C_8D35_032ECF8D8C54__INCLUDED_)
