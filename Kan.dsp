# Microsoft Developer Studio Project File - Name="Kan" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Application" 0x0101

CFG=Kan - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "Kan.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "Kan.mak" CFG="Kan - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Kan - Win32 Release" (based on "Win32 (x86) Application")
!MESSAGE "Kan - Win32 Debug" (based on "Win32 (x86) Application")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "Kan - Win32 Release"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /FR /FD /c
# SUBTRACT CPP /YX /Yc /Yu
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x419 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x419 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /machine:I386
# ADD LINK32 netapi32.lib /nologo /subsystem:windows /machine:I386

!ELSEIF  "$(CFG)" == "Kan - Win32 Debug"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /FR /FD /GZ /c
# SUBTRACT CPP /YX /Yc /Yu
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x419 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x419 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept
# ADD LINK32 netapi32.lib /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept

!ENDIF 

# Begin Target

# Name "Kan - Win32 Release"
# Name "Kan - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\ado2.cpp
# End Source File
# Begin Source File

SOURCE=.\BaseTool.cpp
# End Source File
# Begin Source File

SOURCE=.\ChildFrm.cpp
# End Source File
# Begin Source File

SOURCE=.\excel.cpp
# End Source File
# Begin Source File

SOURCE=.\FChoosDate.cpp
# End Source File
# Begin Source File

SOURCE=.\FControl.cpp
# End Source File
# Begin Source File

SOURCE=.\FDeclarant.cpp
# End Source File
# Begin Source File

SOURCE=.\FEdit.cpp
# End Source File
# Begin Source File

SOURCE=.\FExec.cpp
# End Source File
# Begin Source File

SOURCE=.\FExecList.cpp
# End Source File
# Begin Source File

SOURCE=.\FFilter.cpp
# End Source File
# Begin Source File

SOURCE=.\FFind.cpp
# End Source File
# Begin Source File

SOURCE=.\FInbox.cpp
# End Source File
# Begin Source File

SOURCE=.\FInMessage.cpp
# End Source File
# Begin Source File

SOURCE=.\FOutbox.cpp
# End Source File
# Begin Source File

SOURCE=.\FOutFilter.cpp
# End Source File
# Begin Source File

SOURCE=.\FOutFind.cpp
# End Source File
# Begin Source File

SOURCE=.\FOutMessage.cpp
# End Source File
# Begin Source File

SOURCE=.\FPeriod.cpp
# End Source File
# Begin Source File

SOURCE=.\FResolution.cpp
# End Source File
# Begin Source File

SOURCE=.\FStatistic.cpp
# End Source File
# Begin Source File

SOURCE=.\FTheme.cpp
# End Source File
# Begin Source File

SOURCE=.\FUserChoose.cpp
# End Source File
# Begin Source File

SOURCE=.\FWait.cpp
# End Source File
# Begin Source File

SOURCE=.\Kan.cpp
# End Source File
# Begin Source File

SOURCE=.\Kan.odl
# End Source File
# Begin Source File

SOURCE=.\Kan.rc
# End Source File
# Begin Source File

SOURCE=.\KanDoc.cpp
# End Source File
# Begin Source File

SOURCE=.\KanView.cpp
# End Source File
# Begin Source File

SOURCE=.\MainFrm.cpp
# End Source File
# Begin Source File

SOURCE=.\msflexgrid.cpp
# End Source File
# Begin Source File

SOURCE=.\msmask.cpp
# End Source File
# Begin Source File

SOURCE=.\msword.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\ado2.h
# End Source File
# Begin Source File

SOURCE=.\BaseTool.h
# End Source File
# Begin Source File

SOURCE=.\ChildFrm.h
# End Source File
# Begin Source File

SOURCE=.\excel.h
# End Source File
# Begin Source File

SOURCE=.\FChoosDate.h
# End Source File
# Begin Source File

SOURCE=.\FControl.h
# End Source File
# Begin Source File

SOURCE=.\FDeclarant.h
# End Source File
# Begin Source File

SOURCE=.\FEdit.h
# End Source File
# Begin Source File

SOURCE=.\FExec.h
# End Source File
# Begin Source File

SOURCE=.\FExecList.h
# End Source File
# Begin Source File

SOURCE=.\FFilter.h
# End Source File
# Begin Source File

SOURCE=.\FFind.h
# End Source File
# Begin Source File

SOURCE=.\FInbox.h
# End Source File
# Begin Source File

SOURCE=.\FInMessage.h
# End Source File
# Begin Source File

SOURCE=.\FOutbox.h
# End Source File
# Begin Source File

SOURCE=.\FOutFilter.h
# End Source File
# Begin Source File

SOURCE=.\FOutFind.h
# End Source File
# Begin Source File

SOURCE=.\FOutMessage.h
# End Source File
# Begin Source File

SOURCE=.\FPeriod.h
# End Source File
# Begin Source File

SOURCE=.\FResolution.h
# End Source File
# Begin Source File

SOURCE=.\FStatistic.h
# End Source File
# Begin Source File

SOURCE=.\FTheme.h
# End Source File
# Begin Source File

SOURCE=.\FUserChoose.h
# End Source File
# Begin Source File

SOURCE=.\FWait.h
# End Source File
# Begin Source File

SOURCE=.\Kan.h
# End Source File
# Begin Source File

SOURCE=.\KanDoc.h
# End Source File
# Begin Source File

SOURCE=.\KanView.h
# End Source File
# Begin Source File

SOURCE=.\MainFrm.h
# End Source File
# Begin Source File

SOURCE=.\msflexgrid.h
# End Source File
# Begin Source File

SOURCE=.\msmask.h
# End Source File
# Begin Source File

SOURCE=.\msword.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\res\cancel.ico
# End Source File
# Begin Source File

SOURCE=.\res\complete.ico
# End Source File
# Begin Source File

SOURCE=.\res\complete1.ico
# End Source File
# Begin Source File

SOURCE=.\res\complite.ico
# End Source File
# Begin Source File

SOURCE=.\res\control.ico
# End Source File
# Begin Source File

SOURCE=.\res\delete.ico
# End Source File
# Begin Source File

SOURCE=.\res\filter.ico
# End Source File
# Begin Source File

SOURCE=.\res\find.ico
# End Source File
# Begin Source File

SOURCE=.\res\gray.ico
# End Source File
# Begin Source File

SOURCE=.\res\green.ico
# End Source File
# Begin Source File

SOURCE=.\res\ico00001.ico
# End Source File
# Begin Source File

SOURCE=.\res\icon1.ico
# End Source File
# Begin Source File

SOURCE=.\res\idr_inty.ico
# End Source File
# Begin Source File

SOURCE=.\res\idr_outt.ico
# End Source File
# Begin Source File

SOURCE=.\res\in.bmp
# End Source File
# Begin Source File

SOURCE=.\res\in_creat.ico
# End Source File
# Begin Source File

SOURCE=.\res\inbox.ico
# End Source File
# Begin Source File

SOURCE=.\res\inbox1.ico
# End Source File
# Begin Source File

SOURCE=.\res\Kan.ico
# End Source File
# Begin Source File

SOURCE=.\res\Kan.rc2
# End Source File
# Begin Source File

SOURCE=.\res\KanDoc.ico
# End Source File
# Begin Source File

SOURCE=.\res\ok.ico
# End Source File
# Begin Source File

SOURCE=.\res\out.bmp
# End Source File
# Begin Source File

SOURCE=.\res\out_crea.ico
# End Source File
# Begin Source File

SOURCE=.\res\outbox.ico
# End Source File
# Begin Source File

SOURCE=.\res\outbox1.ico
# End Source File
# Begin Source File

SOURCE=.\res\outlette.ico
# End Source File
# Begin Source File

SOURCE=.\res\person.ico
# End Source File
# Begin Source File

SOURCE=.\res\post1.ico
# End Source File
# Begin Source File

SOURCE=.\res\post2.ico
# End Source File
# Begin Source File

SOURCE=.\res\post3.ico
# End Source File
# Begin Source File

SOURCE=.\res\post4.ico
# End Source File
# Begin Source File

SOURCE=.\res\red.ico
# End Source File
# Begin Source File

SOURCE=.\res\refresh.ico
# End Source File
# Begin Source File

SOURCE=.\res\save.ico
# End Source File
# Begin Source File

SOURCE=.\res\search.ico
# End Source File
# Begin Source File

SOURCE=.\res\test.ico
# End Source File
# Begin Source File

SOURCE=.\res\Toolbar.bmp
# End Source File
# Begin Source File

SOURCE=.\res\unfil.ico
# End Source File
# End Group
# Begin Source File

SOURCE=.\Kan.reg
# End Source File
# Begin Source File

SOURCE=.\ReadMe.txt
# End Source File
# End Target
# End Project
# Section Kan : {C932BA85-4374-101B-A56C-00AA003668DC}
# 	2:21:DefaultSinkHeaderFile:msmask.h
# 	2:16:DefaultSinkClass:CMSMask
# End Section
# Section Kan : {6262D3A0-531B-11CF-91F6-C2863C385E30}
# 	2:21:DefaultSinkHeaderFile:msflexgrid.h
# 	2:16:DefaultSinkClass:CMSFlexGrid
# End Section
# Section Kan : {5F4DF280-531B-11CF-91F6-C2863C385E30}
# 	2:5:Class:CMSFlexGrid
# 	2:10:HeaderFile:msflexgrid.h
# 	2:8:ImplFile:msflexgrid.cpp
# End Section
# Section Kan : {4D6CC9A0-DF77-11CF-8E74-00A0C90F26F8}
# 	2:5:Class:CMSMask
# 	2:10:HeaderFile:msmask.h
# 	2:8:ImplFile:msmask.cpp
# End Section
