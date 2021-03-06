#ifndef __BCGCBPROCONFIG_H
#define __BCGCBPROCONFIG_H

// This is a part of the BCGControlBar Library
// Copyright (C) 1998-2018 BCGSoft Ltd.
// All rights reserved.
//
// This source code can be used, distributed or modified
// only under terms and conditions 
// of the accompanying license agreement.

//---------------------------------------------------
// Uncomment some of these definitions to exclude
// non-required features and reduce the library size:
//---------------------------------------------------

//#define BCGP_EXCLUDE_GRID_CTRL
//#define BCGP_EXCLUDE_GANTT
//#define BCGP_EXCLUDE_PLANNER
//#define BCGP_EXCLUDE_EDIT_CTRL
//#define BCGP_EXCLUDE_PROP_LIST
//#define BCGP_EXCLUDE_POPUP_WINDOW
//#define BCGP_EXCLUDE_SHELL
//#define BCGP_EXCLUDE_TOOLBOX
//#define BCGP_EXCLUDE_HOT_SPOT_IMAGE
//#define BCGP_EXCLUDE_ANIM_CTRL
//#define BCGP_EXCLUDE_TASK_PANE
//#define BCGP_EXCLUDE_RIBBON
//#define BCGP_EXCLUDE_PNG_SUPPORT
//#define BCGP_EXCLUDE_GDI_PLUS
//#define BCGP_USE_EXTERNAL_PNG_LIB
//#define BCGP_USE_EXTERNAL_ZLIB
//#define BCGP_USE_CIMAGE_FOR_PNG	// For VS.NET or higher only
//#define BCGP_CLEAN_DLL_PATH
//#define BCGP_EXCLUDE_OPENGL
//#define BCGP_EXCLUDE_SVG

#if (defined(_M_ARM) || defined(_M_ARM64)) && (!defined(BCGP_EXCLUDE_OPENGL))
	#define BCGP_EXCLUDE_OPENGL	// ARM/ARM64 doesn't have OpenGL support
#endif

#endif // __BCGCBPROCONFIG_H
