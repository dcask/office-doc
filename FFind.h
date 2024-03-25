#if !defined(AFX_FFIND_H__5705B65C_80B3_4A5D_9EF1_BE02759E9CAE__INCLUDED_)
#define AFX_FFIND_H__5705B65C_80B3_4A5D_9EF1_BE02759E9CAE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FFind.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FFind dialog

class FFind : public CDialog
{
// Construction
public:
	FFind(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FFind)
	enum { IDD = IDD_FIND };
	CSpinButtonCtrl	m_Spin;
	UINT	m_uiYear;
	CString	m_csRegNum;
	CString	m_csExtInNum;
	CString	m_csDeclarant;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FFind)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FFind)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnDeclFind();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FFIND_H__5705B65C_80B3_4A5D_9EF1_BE02759E9CAE__INCLUDED_)
