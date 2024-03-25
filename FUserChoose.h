#if !defined(AFX_FUSERCHOOSE_H__4A1AAFE3_D46A_4897_BFF6_F14BF767FD7E__INCLUDED_)
#define AFX_FUSERCHOOSE_H__4A1AAFE3_D46A_4897_BFF6_F14BF767FD7E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FUserChoose.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FUserChoose dialog

class FUserChoose : public CDialog
{
public:
	int	m_iID;
// Construction
public:
	FUserChoose(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FUserChoose)
	enum { IDD = IDD_USER };
	CString	m_csUser;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FUserChoose)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FUserChoose)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FUSERCHOOSE_H__4A1AAFE3_D46A_4897_BFF6_F14BF767FD7E__INCLUDED_)
