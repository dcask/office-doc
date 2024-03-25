//{{AFX_INCLUDES()
#include "msflexgrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_FEXEC_H__60B718AF_6F80_46B5_B6CA_10C6DFF70445__INCLUDED_)
#define AFX_FEXEC_H__60B718AF_6F80_46B5_B6CA_10C6DFF70445__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FExec.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FExec dialog

class FExec : public CDialog
{
public:
	CPtrArray*		m_paExecuters;
	int				m_iSize;
	BOOL			m_bAnswer;
	COleDateTime	m_ProcessDate;
	CToolTipCtrl*	m_ToolTip;
	long			m_lInLetterID;
// Construction
public:
	//void AddDate(int days);
	void View();
	FExec(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FExec)
	enum { IDD = IDD_EXECUTOR };
	CEdit	m_Days;
	CDateTimeCtrl	m_Schedule;
	CDateTimeCtrl	m_Fact;
	CMSFlexGrid	m_Data;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FExec)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL DestroyWindow();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(FExec)
	afx_msg void OnChange();
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnEnterCellData();
	afx_msg void OnLeaveCellData();
	afx_msg void OnKillfocusSchedule(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FEXEC_H__60B718AF_6F80_46B5_B6CA_10C6DFF70445__INCLUDED_)
