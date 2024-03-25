//{{AFX_INCLUDES()
#include "msflexgrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_FCONTROL_H__087911AB_FFE4_424B_9C75_D590EFFB08A6__INCLUDED_)
#define AFX_FCONTROL_H__087911AB_FFE4_424B_9C75_D590EFFB08A6__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FControl.h : header file
//
#include "KanDoc.h"
/////////////////////////////////////////////////////////////////////////////
// FControl dialog

class FControl : public CDialog
{
public:
	CKanDoc*	doc;
	CString m_filterDeclarant;
	BOOL	m_bUseDates;
	BOOL	m_bKeypressed;
	BOOL	m_bExportSendMessage;
	CString m_csQueryName;
	int			m_iWidth;
	CToolTipCtrl* m_ToolTip;
// Construction
public:
	void GetSetForExport();
	void Show();
	FControl(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FControl)
	enum { IDD = IDD_CONTROL };
	COleDateTime	m_Begin;
	COleDateTime	m_End;
	CString	m_csInfo;
	CMSFlexGrid	m_LetterList;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FControl)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL DestroyWindow();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	int m_iCx;	// Для контроля размеров данного представления
	int m_iCy;	// Для контроля размеров данного представления
	long m_lMouseX;	
	long m_lMouseY;	
	long m_lIdleTime;
	BOOL	m_bSortInc;	// Сортировка столбца возр/убыв
	// Generated message map functions
	//{{AFX_MSG(FControl)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	virtual BOOL OnInitDialog();
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnClickLetterlist();
	afx_msg void OnDblClickLetterlist();
	afx_msg void OnEqual();
	afx_msg void OnCloseupBegin(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnCloseupEnd(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSend();
	afx_msg void OnExport();
	afx_msg void OnSelChangeLetterlist();
	afx_msg void OnMouseMoveLetterlist(short Button, short Shift, long x, long y);
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FCONTROL_H__087911AB_FFE4_424B_9C75_D590EFFB08A6__INCLUDED_)
