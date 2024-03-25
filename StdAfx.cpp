// stdafx.cpp : source file that includes just the standard includes
//	Kan.pch will be the pre-compiled header
//	stdafx.obj will contain the pre-compiled type information

#include "stdafx.h"
#include "Kan.h"
void MegaFail()
{
	CADOBaseTool* bt=((CKanApp*)AfxGetApp())->m_bt;
	bt->Dump(((CKanApp*)AfxGetApp())->m_csAppPath+"\\errors.log");
	AfxMessageBox("Непредвиденная ошибка", MB_ICONSTOP);
}

