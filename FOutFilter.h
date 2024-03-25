#if !defined(AFX_FOUTFILTER_H__B9ED6C1E_7A77_4784_ACD8_629ADE3ABFA0__INCLUDED_)
#define AFX_FOUTFILTER_H__B9ED6C1E_7A77_4784_ACD8_629ADE3ABFA0__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FOutFilter.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FOutFilter dialog

class FOutFilter : public CDialog
{
// Construction
public:
	FOutFilter(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FOutFilter)
	enum { IDD = IDD_FILTER_OUT };
	CString	m_csAuthor;
	CString	m_csContent;
	CString	m_csReciever;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FOutFilter)
	public:
	virtual BOOL DestroyWindow();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CToolTipCtrl* m_ToolTip;
	// Generated message map functions
	//{{AFX_MSG(FOutFilter)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnPromt();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FOUTFILTER_H__B9ED6C1E_7A77_4784_ACD8_629ADE3ABFA0__INCLUDED_)
