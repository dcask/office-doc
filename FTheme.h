#if !defined(AFX_FTHEME_H__C0F27599_BDFB_4016_8D5D_2BE369FA8AEE__INCLUDED_)
#define AFX_FTHEME_H__C0F27599_BDFB_4016_8D5D_2BE369FA8AEE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FTheme.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FTheme dialog

class FTheme : public CDialog
{
public:
	CPtrArray*		m_paThemes;
// Construction
public:
	FTheme(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FTheme)
	enum { IDD = IDD_THEME };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FTheme)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FTheme)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FTHEME_H__C0F27599_BDFB_4016_8D5D_2BE369FA8AEE__INCLUDED_)
