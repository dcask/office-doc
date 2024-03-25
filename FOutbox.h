//{{AFX_INCLUDES()
#include "msflexgrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_FOUTBOX_H__7AF75EF9_7101_4BB9_A71C_DCC297A81F5C__INCLUDED_)
#define AFX_FOUTBOX_H__7AF75EF9_7101_4BB9_A71C_DCC297A81F5C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FOutbox.h : header file
//
#include "KanDoc.h"
/////////////////////////////////////////////////////////////////////////////
// FOutbox dialog

class FOutbox : public CDialog
{
public:
	CKanDoc*	doc;
	CString m_outFilterAuthor;
	CString m_outFilterContent;
	CString m_outFilterReciever;
	BOOL	m_bExportSendMessage;
	BOOL	m_bKeypressed;
	int			m_iWidth;
	CToolTipCtrl* m_ToolTip;
// Construction
public:
	void GetSetForExport();
	void Show();
	FOutbox(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FOutbox)
	enum { IDD = IDD_OUTBOX };
	CMSFlexGrid	m_LetterList;
	COleDateTime	m_Begin;
	COleDateTime	m_End;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FOutbox)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL DestroyWindow();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	int m_iCx;	// ƒл€ контрол€ размеров данного представлени€
	int m_iCy;	// ƒл€ контрол€ размеров данного представлени€
	long m_lMouseX;	
	long m_lMouseY;	
	long m_lIdleTime;
	BOOL	m_bSortInc;
	// Generated message map functions
	//{{AFX_MSG(FOutbox)
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnClickLetterlist();
	afx_msg void OnDblClickLetterlist();
	afx_msg void OnEqual();
	afx_msg void OnCloseupBegin(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnCloseupEnd(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnExport();
	afx_msg void OnSelChangeLetterlist();
	afx_msg void OnMouseMoveLetterlist(short Button, short Shift, long x, long y);
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FOUTBOX_H__7AF75EF9_7101_4BB9_A71C_DCC297A81F5C__INCLUDED_)
