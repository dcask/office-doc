// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__DFB44E8D_2D16_4FE5_98F2_6BC22B6D2CF0__INCLUDED_)
#define AFX_STDAFX_H__DFB44E8D_2D16_4FE5_98F2_6BC22B6D2CF0__INCLUDED_

//#define WINVER 0x0500
#define _WIN32_WINNT 0x0600
//#define XPMANIFEST
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions
#include <afxdisp.h>        // MFC Automation classes
#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Common Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT

#include "BaseTool.h"
struct Executer
{
	DWORD dwPersonId;
	long dwInLetterId;
	DWORD dwID;
	CString csName;
	COleDateTime scheduleDate;
	COleDateTime factDate;
	int		state; // 0 - старый, 1 - новый, 2- удаление
};
struct ExecLetter
{
	long	dwID;
	long	inLetterID;
	COleDateTime scheduleDate;
};
struct Theme
{
	CString csName;
	DWORD	dwId;
};
struct Reply
{
	DWORD	dwId;
	long	inLetterID;
	CString RegNum;
	int		year;
	int		state; // 0 - старый, 1 - новый, 2- удаление
};
struct InNum
{
	DWORD			dwId;
	CString			csRegNum;
	COleDateTime	RegDate;
	int				state; // 0 - старый, 1 - новый, 2- удаление
	
};
void MegaFail();
//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__DFB44E8D_2D16_4FE5_98F2_6BC22B6D2CF0__INCLUDED_)
