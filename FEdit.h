//{{AFX_INCLUDES()
#include "msflexgrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_FEDIT_H__F438BCE0_8274_4289_8607_18CD4D5B88F4__INCLUDED_)
#define AFX_FEDIT_H__F438BCE0_8274_4289_8607_18CD4D5B88F4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FEdit.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FEdit dialog

class FEdit : public CDialog
{
// Construction
public:
	FEdit(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FEdit)
	enum { IDD = IDD_EDIT };
	CMSFlexGrid	m_Data;
	CString	m_csQuery;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FEdit)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FEdit)
	afx_msg void OnRun();
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FEDIT_H__F438BCE0_8274_4289_8607_18CD4D5B88F4__INCLUDED_)
