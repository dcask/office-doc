#if !defined(AFX_FPERIOD_H__1E232315_48B7_407B_9E7D_518AD66554F4__INCLUDED_)
#define AFX_FPERIOD_H__1E232315_48B7_407B_9E7D_518AD66554F4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FPeriod.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FPeriod dialog

class FPeriod : public CDialog
{
// Construction
public:
	FPeriod(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FPeriod)
	enum { IDD = IDD_PERIOD };
	COleDateTime	m_Begin;
	COleDateTime	m_End;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FPeriod)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FPeriod)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FPERIOD_H__1E232315_48B7_407B_9E7D_518AD66554F4__INCLUDED_)
