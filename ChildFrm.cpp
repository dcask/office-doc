// ChildFrm.cpp : implementation of the CChildFrame class
//

#include "stdafx.h"
#include "Kan.h"
#include "Wingdi.h"
#include "ChildFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CChildFrame

IMPLEMENT_DYNCREATE(CChildFrame, CMDIChildWnd)

BEGIN_MESSAGE_MAP(CChildFrame, CMDIChildWnd)
	//{{AFX_MSG_MAP(CChildFrame)
	ON_WM_CREATE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CChildFrame construction/destruction

CChildFrame::CChildFrame()
{
	//m_iToolBar=IDR_OUT;
	
}

CChildFrame::~CChildFrame()
{
}

BOOL CChildFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	if( !CMDIChildWnd::PreCreateWindow(cs) )
		return FALSE;
	//cs.style &= ~WS_MAXIMIZEBOX;
	//cs.style &= ~WS_MINIMIZEBOX;
	cs.style &= WS_MAXIMIZE;
	return TRUE;
}



/////////////////////////////////////////////////////////////////////////////
// CChildFrame diagnostics

#ifdef _DEBUG
void CChildFrame::AssertValid() const
{
	CMDIChildWnd::AssertValid();
}

void CChildFrame::Dump(CDumpContext& dc) const
{
	CMDIChildWnd::Dump(dc);
}

#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CChildFrame message handlers

int CChildFrame::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CMDIChildWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	if (!m_wndToolBar.CreateEx(this, TBSTYLE_FLAT/*|TBSTYLE_TRANSPARENT*/, WS_CHILD | WS_VISIBLE | CBRS_TOP
		 | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC))
		//	!m_wndToolBar.LoadToolBar(m_iToolBar))
	{
		TRACE0("Failed to create toolbar\n");
		return -1;      // fail to create
	}
	m_wndToolBar.GetToolBarCtrl().SetButtonSize(CSize(100,36));
	m_wndToolBar.GetToolBarCtrl().SetBitmapSize(CSize(96,32));
	return 0;
}



BOOL CChildFrame::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult) 
{
	LPNMHDR pnmh = (LPNMHDR) lParam;
     if(pnmh->hwndFrom == m_wndToolBar.m_hWnd)
     {
		 if(((CKanApp*)AfxGetApp())->m_iXPStyle!=0)
		 {
			LPNMTBCUSTOMDRAW lpNMCustomDraw = (LPNMTBCUSTOMDRAW) lParam;
			CRect rect;
			m_wndToolBar.GetClientRect(rect);
			
			SelectObject(lpNMCustomDraw->nmcd.hdc,GetStockObject(DC_BRUSH));
			SetDCBrushColor(lpNMCustomDraw->nmcd.hdc,GetSysColor(COLOR_3DFACE)); 

			FillRect(lpNMCustomDraw->nmcd.hdc, rect, (HBRUSH)GetStockObject(DC_BRUSH));
		 }
		
     }

	
	return CMDIChildWnd::OnNotify(wParam, lParam, pResult);
}
