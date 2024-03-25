#if !defined(AFX_FFILTER_H__E53B5BDB_2CAA_486D_9654_2008C6C55D3E__INCLUDED_)
#define AFX_FFILTER_H__E53B5BDB_2CAA_486D_9654_2008C6C55D3E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FFilter.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FFilter dialog

class FFilter : public CDialog
{
// Construction
public:
	FFilter(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FFilter)
	enum { IDD = IDD_FILTER };
	CString	m_csDeclarant;
	CString	m_csContent;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FFilter)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL DestroyWindow();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CToolTipCtrl* m_ToolTip;
	// Generated message map functions
	//{{AFX_MSG(FFilter)
	afx_msg void OnPromt();
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FFILTER_H__E53B5BDB_2CAA_486D_9654_2008C6C55D3E__INCLUDED_)
