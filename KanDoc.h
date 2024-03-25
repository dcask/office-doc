// KanDoc.h : interface of the CKanDoc class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_KANDOC_H__1472E1C1_039C_4707_B9EA_32C7BE7BC38C__INCLUDED_)
#define AFX_KANDOC_H__1472E1C1_039C_4707_B9EA_32C7BE7BC38C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class CKanDoc : public CDocument
{
protected: // create from serialization only
	CKanDoc();
	DECLARE_DYNCREATE(CKanDoc)
// Attributes
public:
	UINT m_uiCurSelectedTab;	// Открытая закладка
	CDialog* m_pTabDialog;		// текущий диалог для таба
	UINT	m_uiID; // 1-главное, 2  - входящее, 3 - исходящее
	long	m_lLetterId;
	/*Данные документа*/
			//Входящий
	CString			m_csRegNum;
	CString			m_csFullNum;
	CString			m_csRegPrefix;
	COleDateTime	m_RegDate;
	CString			m_csDeclarant;
	CString			m_csDeclarantType;
	int				m_iDeclarantType;
	CString			m_csUse;
	int				m_iUse;
	CString			m_csDistrict;
	int				m_iDistrict;
	CString			m_csFrom;
	int				m_iFrom;
	CString			m_csOutRegNum_F;	// исходящий внешний
	CString			m_csOutRegNum;		// исходящий 
	COleDateTime	m_OutRegDate_F;		// дата исходящего внешнего
	COleDateTime	m_OutRegDate;		// дата исходящего
	COleDateTime	m_ProcessDate;		// дата передачи исполнителю 
	CString			m_csContent;
	CString			m_csAuthor;			
	int				m_iAuthor;			
	COleDateTime	m_ResolutionDate;
	CString			m_csResolution;
	CString			m_csExecCourse;		// ход исполнения
	CString			m_csControlType;
	int				m_iControlType;
	CString			m_csResult;
	int				m_iResult;
	CPtrArray		m_aExecuters;
	CPtrArray		m_aThemes;
	CPtrArray		m_aExtNum;
	CPtrArray		m_aReplyOn;
	CString			m_csTheme;
	int				m_iTheme;
	long			m_lInLetterID;
	int				m_iAttrib;
	CString			m_csToPlace;	// адресат для исходящего
	/**/
// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKanDoc)
	public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);
	virtual void OnCloseDocument();
	//}}AFX_VIRTUAL

// Implementation
public:
	BOOL SaveOutData();
	BOOL LoadOutData();
	BOOL SaveInData();
	BOOL LoadInData();
	virtual ~CKanDoc();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CKanDoc)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CKanDoc)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KANDOC_H__1472E1C1_039C_4707_B9EA_32C7BE7BC38C__INCLUDED_)
