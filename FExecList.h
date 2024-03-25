#if !defined(AFX_FEXECLIST_H__C3D8A629_74D1_497B_B127_E225D46AFC90__INCLUDED_)
#define AFX_FEXECLIST_H__C3D8A629_74D1_497B_B127_E225D46AFC90__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FExecList.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FExecList dialog

class FExecList : public CDialog
{
public:
	CPtrArray*		m_paExecuters;
// Construction
public:
	FExecList(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FExecList)
	enum { IDD = IDD_EXEC_LIST };
	CTreeCtrl	m_Tree;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FExecList)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CImageList* m_pImgList;
	// Generated message map functions
	//{{AFX_MSG(FExecList)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	afx_msg void OnClose();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FEXECLIST_H__C3D8A629_74D1_497B_B127_E225D46AFC90__INCLUDED_)
