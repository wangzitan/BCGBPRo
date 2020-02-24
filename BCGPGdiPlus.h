//*******************************************************************************
// COPYRIGHT NOTES
// ---------------
// This is a part of the BCGControlBar Library (Professional Version)
// Copyright (C) 1998-2018 BCGSoft Ltd.
// All rights reserved.
//
// This source code can be used, distributed or modified
// only under terms and conditions 
// of the accompanying license agreement.
//*******************************************************************************

#ifndef BCGPGdiPlus_H
#define BCGPGdiPlus_H

#include "BCGCBPro.h"

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef BCGP_EXCLUDE_GDI_PLUS

	#if _MSC_VER < 1400
		#define __field_ecount_opt(count)
		#define __out_bcount(size)
		#define __out_ecount(size)
	#endif

	#pragma warning( push, 3 )

	#if _MSC_VER >= 1400
		#pragma push_macro("new")
		#undef new
	#endif

#if _MSC_VER < 1300
	#include "GDIPlus\gdiplus.h"
#else
	#include <gdiplus.h>
#endif

	#if _MSC_VER >= 1400
		#pragma pop_macro("new")
	#endif

	#pragma warning( pop )

#endif

#endif
