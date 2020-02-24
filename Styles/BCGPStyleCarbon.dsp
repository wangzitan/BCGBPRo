# Microsoft Developer Studio Project File - Name="Carbon" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=Carbon - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "BCGPStyleCarbon.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "BCGPStyleCarbon.mak" CFG="Carbon - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Carbon - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Carbon - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "Carbon - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "..\..\..\bin"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Ignore_Export_Lib 1
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "..\..\bin"
# PROP Intermediate_Dir "intermediate"
# PROP Ignore_Export_Lib 1
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O1 /D "NDEBUG" /D "WIN32" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /YX /FD /c
# ADD CPP /nologo /MT /W3 /D "NDEBUG" /D "WIN32" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /FD /c
# SUBTRACT CPP /Gf /Gy /YX
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# SUBTRACT BASE MTL /mktyplib203
# ADD MTL /nologo /D "NDEBUG" /win32
# SUBTRACT MTL /mktyplib203
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_BCGP_STYLE_DLL_"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /dll /machine:I386 /nodefaultlib /NOENTRY
# SUBTRACT BASE LINK32 /pdb:none
# ADD LINK32 /nologo /dll /pdb:none /machine:I386 /nodefaultlib /out:"..\..\bin/BCGPStyleCarbon280.dll" /NOENTRY
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Cmds=@Echo Copying Style DLL to 'bin64' folder.	copy ..\..\bin\BCGPStyleCarbon280.dll ..\..\bin64\BCGPStyleCarbon280.dll
# End Special Build Tool

!ELSEIF  "$(CFG)" == "Carbon - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "..\..\..\bin"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Ignore_Export_Lib 1
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "..\..\bin"
# PROP Intermediate_Dir "intermediate"
# PROP Ignore_Export_Lib 1
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O1 /D "NDEBUG" /D "WIN32" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /YX /FD /c
# ADD CPP /nologo /MT /W3 /D "NDEBUG" /D "WIN32" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /FD /c
# SUBTRACT CPP /Gf /Gy /YX
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# SUBTRACT BASE MTL /mktyplib203
# ADD MTL /nologo /D "NDEBUG" /win32
# SUBTRACT MTL /mktyplib203
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_BCGP_STYLE_DLL_"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /dll /machine:I386 /nodefaultlib /NOENTRY
# SUBTRACT BASE LINK32 /pdb:none
# ADD LINK32 /nologo /dll /pdb:none /machine:I386 /nodefaultlib /out:"..\..\bin/BCGPStyleCarbon280.dll" /NOENTRY
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Cmds=@Echo Copying Style DLL to 'bin64' folder.	copy ..\..\bin\BCGPStyleCarbon280.dll ..\..\bin64\BCGPStyleCarbon280.dll
# End Special Build Tool

!ENDIF 

# Begin Target

# Name "Carbon - Win32 Debug"
# Name "Carbon - Win32 Release"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\BCGPStyleCarbon.rc
# End Source File
# Begin Source File

SOURCE=.\BCGPStyleCarbon.rc2
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Group "AppCaption"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\Carbon\AppCaption\SysBtn_Back.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\AppCaption\SysBtn_Back_C.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\AppCaption\SysBtn_Back_C_H.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\AppCaption\SysBtn_Back_H.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\AppCaption\SysBtn_Close.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\AppCaption\SysBtn_Help.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\AppCaption\SysBtn_Maximize.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\AppCaption\SysBtn_Minimize.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\AppCaption\SysBtn_Restore.png
# End Source File
# End Group
# Begin Group "MainWnd"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\Carbon\MainWnd\Border.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\MainWnd\BorderCaption.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\MainWnd\borderL.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\MainWnd\BorderMDIChild.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\MainWnd\borderMini.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\MainWnd\borderMiniCaption.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\MainWnd\borderMiniTB.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\MainWnd\borderR.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\MainWnd\borderTB.png
# End Source File
# End Group
# Begin Group "MenuBar"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Carbon\MenuBar\Btn.png"
# End Source File
# End Group
# Begin Group "PopupBar"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Carbon\PopupBar\Border.png"
# End Source File
# Begin Source File

SOURCE=.\Carbon\PopupBar\btnCaption.png
# End Source File
# Begin Source File

SOURCE=".\Carbon\PopupBar\Gripper.png"
# End Source File
# Begin Source File

SOURCE=.\Carbon\PopupBar\Item_Back.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\PopupBar\Item_Marker_C.png
# End Source File
# Begin Source File

SOURCE=.\Carbon\PopupBar\Item_Marker_R.png
# End Source File
# Begin Source File

SOURCE=".\Carbon\PopupBar\Tear.png"
# End Source File
# End Group
# Begin Group "StatusBar"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Carbon\StatusBar\Back.png"
# End Source File
# Begin Source File

SOURCE=".\Carbon\StatusBar\Back_Ext.png"
# End Source File
# Begin Source File

SOURCE=".\Carbon\StatusBar\PaneBorder.png"
# End Source File
# Begin Source File

SOURCE=".\Carbon\StatusBar\SizeGrip.png"
# End Source File
# End Group
# Begin Group "ToolBar"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\Carbon\ToolBar\Btn.png
# End Source File
# End Group
# Begin Group "Ribbon"

# PROP Default_Filter ""
# Begin Group "Buttons"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\Carbon\Ribbon\Buttons\Group.PNG
# End Source File
# Begin Source File

SOURCE=.\Carbon\Ribbon\Buttons\Push.png
# End Source File
# End Group
# End Group
# Begin Source File

SOURCE=".\Carbon\Style.xml"
# End Source File
# End Group
# End Target
# End Project
