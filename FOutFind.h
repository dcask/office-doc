#if !defined(AFX_FOUTFIND_H__CA7AB9D8_FC9E_4ABF_8678_49CFE574272E__INCLUDED_)
#define AFX_FOUTFIND_H__CA7AB9D8_FC9E_4ABF_8678_49CFE574272E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FOutFind.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FOutFind dialog

class FOutFind : public CDialog
{
// Construction
public:
	FOutFind(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FOutFind)
	enum { IDD = IDD_FIND_OUT };
	CSpinButtonCtrl	m_Spin;
	CString	m_csRegNum;
	UINT	m_uiYear;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FOutFind)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FOutFind)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FOUTFIND_H__CA7AB9D8_FC9E_4ABF_8678_49CFE574272E__INCLUDED_)
