//{{AFX_INCLUDES()
#include "msflexgrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_FSTATISTIC_H__0BE0612E_C8F2_4236_B8EC_7EC4C20A92A1__INCLUDED_)
#define AFX_FSTATISTIC_H__0BE0612E_C8F2_4236_B8EC_7EC4C20A92A1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FStatistic.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FStatistic dialog

class FStatistic : public CDialog
{
public:
	CString	m_csQuery;
	CString m_csUser;
	CString m_csTitle;
	COleDateTime m_Begin;
	COleDateTime m_End;

// Construction
public:
	FStatistic(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FStatistic)
	enum { IDD = IDD_STAT };
	CMSFlexGrid	m_Data;
	CString	m_csInfo;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FStatistic)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FStatistic)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnMUnload();
	afx_msg void OnClose();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FSTATISTIC_H__0BE0612E_C8F2_4236_B8EC_7EC4C20A92A1__INCLUDED_)
