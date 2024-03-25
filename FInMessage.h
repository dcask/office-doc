//{{AFX_INCLUDES()
#include "msmask.h"
//}}AFX_INCLUDES
#if !defined(AFX_FINMESSAGE_H__44CC688E_00E6_482B_8859_EE1591817107__INCLUDED_)
#define AFX_FINMESSAGE_H__44CC688E_00E6_482B_8859_EE1591817107__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FInMessage.h : header file
//
#include "KanDoc.h"
/////////////////////////////////////////////////////////////////////////////
// FInMessage dialog

class FInMessage : public CDialog
{
public:
	CKanDoc*	doc;
	CDialog*	m_pTabDialog;
	UINT m_uiCurSelectedTab;
	CToolBar	m_wndToolBar;
	CImageList*	m_pImgList;
	CImageList*	m_pImgList_d;
	BOOL		m_bNew;
	CToolTipCtrl* m_ToolTip;
	
// Construction
public:
	CString ApplyTemplate(CString Templ, CString Num, CString Decl);
	FInMessage(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FInMessage)
	enum { IDD = IDD_IN_LETTER_NEW };
	CListBox	m_List;
	CTabCtrl	m_Tab;
	CString	m_csContent;
	CString	m_csControlType;
	CString	m_csDeclarant;
	CMSMask	m_mOutDate;
	CMSMask	m_mRegDate;
	COleDateTime	m_OutDate;
	COleDateTime	m_RegDate;
	CString	m_csRegNum;
	CString	m_csExtInNum;
	CString	m_csDeclarantType;
	COleDateTime	m_ProcessDate;
	CString	m_csFullNum;
	CMSMask	m_mProcessDate;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FInMessage)
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
	//{{AFX_MSG(FInMessage)
	virtual BOOL OnInitDialog();
	afx_msg void OnSelchangeTab(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSave();
	afx_msg void OnCancel();
	afx_msg void OnDatetimechangeRegDate(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnDatetimechangeOutDate(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnDeclFind();
	afx_msg void OnSelchangeControltype();
	afx_msg void OnSelchangeDeclType();
	afx_msg void OnDatetimechangeProcDate(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnChangeRegNum();
	afx_msg void OnChangeFromPerson();
	afx_msg void OnChangeMprocDate();
	afx_msg void OnAdd();
	afx_msg void OnDel();
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FINMESSAGE_H__44CC688E_00E6_482B_8859_EE1591817107__INCLUDED_)
