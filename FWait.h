#if !defined(AFX_FWAIT_H__4E071C3C_29B9_4E3F_B14A_4C181EFEE583__INCLUDED_)
#define AFX_FWAIT_H__4E071C3C_29B9_4E3F_B14A_4C181EFEE583__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FWait.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FWait dialog

class FWait : public CDialog
{
// Construction
public:
	void SetInfo(CString csVal);
	FWait(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FWait)
	enum { IDD = IDD_WAIT };
	CString	m_csInfo;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FWait)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FWait)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FWAIT_H__4E071C3C_29B9_4E3F_B14A_4C181EFEE583__INCLUDED_)
