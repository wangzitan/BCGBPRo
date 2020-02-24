# Microsoft Developer Studio Project File - Name="Office 2007 (aqua)" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=Office 2007 (aqua) - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "BCGPStyle2007Aqua.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "BCGPStyle2007Aqua.mak" CFG="Office 2007 (aqua) - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Office 2007 (aqua) - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Office 2007 (aqua) - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "Office 2007 (aqua) - Win32 Debug"

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
# ADD LINK32 /nologo /dll /pdb:none /machine:I386 /nodefaultlib /out:"..\..\bin/BCGPStyle2007Aqua280.dll" /NOENTRY
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Cmds=@Echo Copying Style DLL to 'bin64' folder.	copy ..\..\bin\BCGPStyle2007Aqua280.dll ..\..\bin64\BCGPStyle2007Aqua280.dll
# End Special Build Tool

!ELSEIF  "$(CFG)" == "Office 2007 (aqua) - Win32 Release"

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
# ADD LINK32 /nologo /dll /pdb:none /machine:I386 /nodefaultlib /out:"..\..\bin/BCGPStyle2007Aqua280.dll" /NOENTRY
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Cmds=@Echo Copying Style DLL to 'bin64' folder.	copy ..\..\bin\BCGPStyle2007Aqua280.dll ..\..\bin64\BCGPStyle2007Aqua280.dll
# End Special Build Tool

!ENDIF 

# Begin Target

# Name "Office 2007 (aqua) - Win32 Debug"
# Name "Office 2007 (aqua) - Win32 Release"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\BCGPStyle2007Aqua.rc
# End Source File
# Begin Source File

SOURCE=.\BCGPStyle2007Aqua.rc2
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Group "AppCaption"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Close.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Close_S.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Help.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Help_S.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Maximize.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Maximize_S.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Minimize.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Minimize_S.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Restore.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\AppCaption\SysBtn_Restore_S.png"
# End Source File
# End Group
# Begin Group "ComboBox"

# PROP Default_Filter ""
# End Group
# Begin Group "MainWnd"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\MainWnd\Border.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\MainWnd\BorderCaption.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\MainWnd\BorderMDIChild.png"
# End Source File
# End Group
# Begin Group "MenuBar"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\MenuBar\Btn.png"
# End Source File
# End Group
# Begin Group "PopupBar"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\Border.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\Gripper.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\Item_Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\Item_Marker_C.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\Item_Marker_R.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\Item_Pin.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\MenuHighlight.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\MenuHighlightDisabled.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\MenuScrollBtn.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\ResizeBar.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\ResizeBar_HV.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\ResizeBar_HVT.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\ResizeBar_V.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\PopupBar\Tear.png"
# End Source File
# End Group
# Begin Group "StatusBar"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\StatusBar\Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\StatusBar\Back_Ext.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\StatusBar\PaneBorder.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\StatusBar\SizeGrip.png"
# End Source File
# End Group
# Begin Group "Tab"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Tab\3D.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Tab\Dot_E.PNG"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Tab\Dot_R.PNG"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Tab\Flat.png"
# End Source File
# End Group
# Begin Group "TaskPane"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\TaskPane\ScrollBtn.png"
# End Source File
# End Group
# Begin Group "ToolBar"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\ToolBar\Btn.png"
# End Source File
# End Group
# Begin Group "Ribbon"

# PROP Default_Filter ""
# Begin Group "Buttons"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Check.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Default.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Default_Image.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Default_QAT.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Default_QAT_Icon.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Group.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Group_F.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Group_L.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Group_M.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Group_S.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\GroupMenu_F_C.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\GroupMenu_F_M.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\GroupMenu_L_C.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\GroupMenu_L_M.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\GroupMenu_M_C.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\GroupMenu_M_M.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Launch_Icon.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Main.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Main_Panel.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Menu_H_C.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Menu_H_M.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Menu_V_C.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Menu_V_M.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Normal_B.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Normal_S.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Page_L.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Page_R.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Palette_B.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Palette_M.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Palette_T.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Push.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\Radio.PNG"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Buttons\StatusPane.png"
# End Source File
# End Group
# Begin Group "Category_Context"

# PROP Default_Filter ""
# Begin Group "Red"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Red\Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Red\Caption.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Red\Default.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Red\Tab.png"
# End Source File
# End Group
# Begin Group "Orange"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Orange\Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Orange\Caption.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Orange\Default.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Orange\Tab.png"
# End Source File
# End Group
# Begin Group "Yellow"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Yellow\Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Yellow\Caption.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Yellow\Default.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Yellow\Tab.png"
# End Source File
# End Group
# Begin Group "Green"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Green\Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Green\Caption.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Green\Default.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Green\Tab.png"
# End Source File
# End Group
# Begin Group "Blue"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Blue\Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Blue\Caption.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Blue\Default.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Blue\Tab.png"
# End Source File
# End Group
# Begin Group "Indigo"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Indigo\Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Indigo\Caption.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Indigo\Default.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Indigo\Tab.png"
# End Source File
# End Group
# Begin Group "Violet"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Violet\Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Violet\Caption.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Violet\Default.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Violet\Tab.png"
# End Source File
# End Group
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Panel_Back_B.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Panel_Back_T.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Context\Separator.png"
# End Source File
# End Group
# Begin Group "Slider"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Slider\Minus.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Slider\Plus.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Slider\Thumb.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Slider\Thumb_H.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Slider\Thumb_L.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Slider\Thumb_R.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Slider\Thumb_T.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Slider\Thumb_V.png"
# End Source File
# End Group
# Begin Group "Progress"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Progress\Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Progress\Infinity.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Progress\Normal.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Progress\Normal_Ext.png"
# End Source File
# End Group
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Border_Floaty.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Caption_QA.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Caption_QA_Glass.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Back.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Category_Tab.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\KeyTip.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Panel_Back_B.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Panel_Back_T.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Ribbon\Separator.png"
# End Source File
# End Group
# Begin Group "Scroll"

# PROP Default_Filter ""
# Begin Group "Horz"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Horz\Back_1.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Horz\Back_2.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Horz\Item_1.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Horz\Item_2.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Horz\Thumb.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Horz\ThumbIcon.png"
# End Source File
# End Group
# Begin Group "Vert"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Vert\Back_1.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Vert\Back_2.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Vert\Item_1.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Vert\Item_2.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Vert\Thumb.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Scroll\Vert\ThumbIcon.png"
# End Source File
# End Group
# End Group
# Begin Group "Slider No. 1"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Slider\Thumb_B.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Slider\Thumb_H.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Slider\Thumb_L.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Slider\Thumb_R.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Slider\Thumb_T.png"
# End Source File
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Slider\Thumb_V.png"
# End Source File
# End Group
# Begin Group "Calculator"

# PROP Default_Filter ""
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Calculator\Btn.png"
# End Source File
# End Group
# Begin Source File

SOURCE=".\Office 2007 (aqua)\Style.xml"
# End Source File
# End Group
# End Target
# End Project
