#if !defined(AFX_FCHOOSDATE_H__18FFDCC8_2C3E_4868_83A8_C880C27169F6__INCLUDED_)
#define AFX_FCHOOSDATE_H__18FFDCC8_2C3E_4868_83A8_C880C27169F6__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FChoosDate.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FChoosDate dialog

class FChoosDate : public CDialog
{
// Construction
public:
	FChoosDate(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FChoosDate)
	enum { IDD = IDD_DATE };
	COleDateTime	m_Date;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FChoosDate)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FChoosDate)
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FCHOOSDATE_H__18FFDCC8_2C3E_4868_83A8_C880C27169F6__INCLUDED_)
