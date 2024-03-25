//{{AFX_INCLUDES()
#include "msmask.h"
//}}AFX_INCLUDES
#if !defined(AFX_FOUTMESSAGE_H__6DBD6870_DC85_455B_8226_7CB2C637B033__INCLUDED_)
#define AFX_FOUTMESSAGE_H__6DBD6870_DC85_455B_8226_7CB2C637B033__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FOutMessage.h : header file
//
#include "KanDoc.h"
/////////////////////////////////////////////////////////////////////////////
// FOutMessage dialog

class FOutMessage : public CDialog
{
public:
	CKanDoc*	doc;
	CToolBar	m_wndToolBar;
	CImageList*	m_pImgList;
	CImageList*	m_pImgList_d;
	CImageList*	m_pImgList_Exec;
	CToolTipCtrl* m_ToolTip;
	BOOL		m_bNew;
// Construction
public:
	void LoadExecutors();
	FOutMessage(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FOutMessage)
	enum { IDD = IDD_OUT_LETTER_NEW };
	CListBox	m_InList;
	CSpinButtonCtrl	m_Spin;
	CTreeCtrl	m_Tree;
	CString	m_csAuthor;
	CString	m_csContent;
	CMSMask	m_mRegDate;
	COleDateTime	m_RegDate;
	CString	m_csRegNum;
	CString	m_csReplyOn;
	int		m_iYear;
	CString	m_csPrefix;
	CString	m_csToPlace;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FOutMessage)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL DestroyWindow();
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON	m_hIcon;
	// Generated message map functions
	//{{AFX_MSG(FOutMessage)
	afx_msg void OnReplyFind();
	virtual BOOL OnInitDialog();
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnSelchangeAuthor();
	afx_msg void OnDatetimechangeRegDate(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnCancel();
	afx_msg void OnSave();
	afx_msg void OnButtonAll();
	afx_msg void OnChangeRegNum();
	afx_msg void OnAdrFind();
	afx_msg void OnDel();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FOUTMESSAGE_H__6DBD6870_DC85_455B_8226_7CB2C637B033__INCLUDED_)
