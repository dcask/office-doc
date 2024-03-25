//{{AFX_INCLUDES()
#include "msflexgrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_FINBOX_H__05B32666_EA46_4DCF_B496_EC60482D5635__INCLUDED_)
#define AFX_FINBOX_H__05B32666_EA46_4DCF_B496_EC60482D5635__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FInbox.h : header file
//
#include "KanDoc.h"
/////////////////////////////////////////////////////////////////////////////
// FInbox dialog

class FInbox : public CDialog
{
public:
	CKanDoc*	doc;
	CString m_inFilterDeclarant;
	CString m_inFilterContent;
	BOOL	m_bKeypressed;
	BOOL	m_bExportSendMessage;
	CString m_csQueryName;
	int			m_iWidth;
	CToolTipCtrl* m_ToolTip;
// Construction
public:
	void GetSetForExport();
	void Show();
	FInbox(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FInbox)
	enum { IDD = IDD_INBOX };
	CMSFlexGrid	m_LetterList;
	COleDateTime	m_Begin;
	COleDateTime	m_End;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FInbox)
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
	//{{AFX_MSG(FInbox)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnClickLetterlist();
	afx_msg void OnDblClickLetterlist();
	afx_msg void OnCloseupBegin(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnCloseupEnd(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnEqual();
	afx_msg void OnExport();
	afx_msg void OnSelChangeLetterlist();
	afx_msg void OnMouseMoveLetterlist(short Button, short Shift, long x, long y);
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FINBOX_H__05B32666_EA46_4DCF_B496_EC60482D5635__INCLUDED_)
