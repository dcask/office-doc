// KanView.h : interface of the CKanView class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_KANVIEW_H__CF97589B_82AB_47EF_AB05_FE5DE385798A__INCLUDED_)
#define AFX_KANVIEW_H__CF97589B_82AB_47EF_AB05_FE5DE385798A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class CKanView : public CFormView
{
protected: // create from serialization only
	CKanView();
	DECLARE_DYNCREATE(CKanView)

public:
	//{{AFX_DATA(CKanView)
	enum { IDD = IDD_KAN_FORM };
	CTabCtrl	m_Tab;
	//}}AFX_DATA

// Attributes
public:
	CKanDoc* GetDocument();
	CImageList* m_pTabImgList;
	COleDateTime m_Begin;
	COleDateTime m_End;
	CToolTipCtrl* m_ToolTipIn;
	CToolTipCtrl* m_ToolTipOut;
// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKanView)
	public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual void OnInitialUpdate(); // called first time after construct
	virtual void OnActivateView(BOOL bActivate, CView* pActivateView, CView* pDeactiveView);
	//}}AFX_VIRTUAL

// Implementation
public:
	void StartSession();
	void GetPeriod(CString name, COleDateTime &begin, COleDateTime &end);
	BOOL LoadToolBar(UINT id,CToolBar* tb);
	virtual ~CKanView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:
	int m_iCx;	// Для контроля размеров данного представления
	int m_iCy;	// Для контроля размеров данного представления
	BOOL	m_bTimer;
	CToolBar	m_wndToolBar;
	CImageList*	m_pImgListIn;
	CImageList*	m_pImgListIn_d;
	CImageList*	m_pImgListOut;
	CImageList*	m_pImgListOut_d;

		/*для фильтра*/
	COleDateTime m_inFilterBegin;
	COleDateTime m_inFilterEnd;
	COleDateTime m_outFilterBegin;
	COleDateTime m_outFilterEnd;
	CString m_inFilterDeclarant;
	CString m_outFilterAuthor;
	CString m_outFilterReciever;
	CString m_inFilterContent;
	CString m_outFilterContent;
		/**/
// Generated message map functions
protected:
	//{{AFX_MSG(CKanView)
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnSelchangeMaintab(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnLetterCreate();
	afx_msg void OnLetterDelete();
	afx_msg void OnLetterFind();
	afx_msg void OnLetterFilter();
	afx_msg void OnReportsInvalid();
	afx_msg void OnStatByuser();
	afx_msg void OnStatBycontrol();
	afx_msg void OnStatByperson();
	afx_msg void OnStatByauthor();
	afx_msg void OnHelpCompact();
	afx_msg void OnEditQuery();
	afx_msg void OnEditMine();
	afx_msg void OnReportsExcel();
	afx_msg void OnTimer(UINT nIDEvent);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnStatOnline();
	afx_msg void OnLetterRefresh();
	afx_msg void OnLetterUnfilter();
	afx_msg void OnReportsSign();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

#ifndef _DEBUG  // debug version in KanView.cpp
inline CKanDoc* CKanView::GetDocument()
   { return (CKanDoc*)m_pDocument; }
#endif

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KANVIEW_H__CF97589B_82AB_47EF_AB05_FE5DE385798A__INCLUDED_)
