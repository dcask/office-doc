//{{AFX_INCLUDES()
#include "msmask.h"
//}}AFX_INCLUDES
#if !defined(AFX_FRESOLUTION_H__EF8C0CC4_D38B_4E9A_9377_FD64F503375C__INCLUDED_)
#define AFX_FRESOLUTION_H__EF8C0CC4_D38B_4E9A_9377_FD64F503375C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FResolution.h : header file
//
#include "KanDoc.h"
/////////////////////////////////////////////////////////////////////////////
// FResolution dialog

class FResolution : public CDialog
{
// Construction
public:
	CKanDoc *doc;
	FResolution(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FResolution)
	enum { IDD = IDD_RESOLUTION };
	CComboBox	m_Author;
	CString	m_csResolution;
	CMSMask	m_mOutDate;
	COleDateTime	m_OutDate;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FResolution)
	public:
	virtual BOOL DestroyWindow();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FResolution)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	virtual BOOL OnInitDialog();
	afx_msg void OnDatetimechangeOutDate(NMHDR* pNMHDR, LRESULT* pResult);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FRESOLUTION_H__EF8C0CC4_D38B_4E9A_9377_FD64F503375C__INCLUDED_)
