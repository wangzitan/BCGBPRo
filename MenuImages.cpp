//*******************************************************************************
// COPYRIGHT NOTES
// ---------------
// This is a part of the BCGControlBar Library
// Copyright (C) 1998-2018 BCGSoft Ltd.
// All rights reserved.
//
// This source code can be used, distributed or modified
// only under terms and conditions 
// of the accompanying license agreement.
//*******************************************************************************

// MenuImages.cpp: implementation of the CBCGPMenuImages class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MenuImages.h"
#include "BCGPLocalResource.h"
#include "BCGPDrawManager.h"
#include "bcgprores.h"
#include "bcgglobals.h"

static const COLORREF clrTransparent = RGB (255, 0, 255);

static CSize GetOriginalImageSize(UINT* pnResID = NULL, double* pdblScale = NULL)
{
	CSize size(9, 9);
	double dblScale = globalData.GetRibbonImageScale();
	UINT nResID = IDB_BCGBARRES_MENU_IMAGES;

	if (globalData.Is32BitIcons())
	{
#ifndef _BCGSUITE_
		if (dblScale > 1.51)
		{
			nResID = IDB_BCGBARRES_MENU_IMAGES24_200;
			dblScale = dblScale / 2.0;
			size = CSize(18, 18);
		}
		else if (dblScale > 1.26)
		{
			nResID = IDB_BCGBARRES_MENU_IMAGES24_150;
			dblScale = dblScale / 1.5;
			size = CSize(13, 13);
		}
		else if (dblScale > 1.1)
		{
			nResID = IDB_BCGBARRES_MENU_IMAGES24_125;
			dblScale = dblScale / 1.25;
			size = CSize(11, 11);
		}
		else
#endif
		{
			nResID = IDB_BCGBARRES_MENU_IMAGES24;
		}
	}

	if (pnResID != NULL)
	{
		*pnResID = nResID;
	}

	if (pdblScale != NULL)
	{
		*pdblScale = dblScale;
	}

	return size;
}

CBCGPToolBarImages CBCGPMenuImages::m_ImagesBlack;
CBCGPToolBarImages CBCGPMenuImages::m_ImagesDkGray;
CBCGPToolBarImages CBCGPMenuImages::m_ImagesGray;
CBCGPToolBarImages CBCGPMenuImages::m_ImagesLtGray;
CBCGPToolBarImages CBCGPMenuImages::m_ImagesWhite;
CBCGPToolBarImages CBCGPMenuImages::m_ImagesBlack2;

BOOL CBCGPMenuImages::m_bInitializing = FALSE;
BOOL CBCGPMenuImages::m_bCleannedUpGlobally = FALSE;

BOOL CBCGPMenuImages::Initialize ()
{
	if (m_bInitializing || m_bCleannedUpGlobally)
	{
		return FALSE;
	}

	if (m_ImagesBlack.IsValid ())
	{
		return TRUE;
	}

	m_bInitializing = TRUE;

	CBCGPLocalResource locaRes;
	
	UINT nImageResID = 0;
	double dblScale = 1.0;

	m_ImagesBlack.SetImageSize (GetOriginalImageSize(&nImageResID, &dblScale));

#ifndef _BCGSUITE_
	//----------------------------
	// Try to load SVG image list:
	//----------------------------
	if (!m_ImagesBlack.Load(IDB_BCGBARRES_MENU_IMAGES_SVG))
#endif
	{
		if (!m_ImagesBlack.Load(nImageResID))
		{
			TRACE(_T("CBCGPMenuImages. Can't load menu images %x\n"), nImageResID);
			m_bInitializing = FALSE;
			return FALSE;
		}
	}

#ifndef _BCGSUITE_
	if (globalData.m_bIsBlackHighContrast)
	{
		m_ImagesBlack.AdaptColors(RGB(0, 0, 0), RGB(255, 255, 255));
	}
#endif

	if (m_ImagesBlack.IsRTL ())
	{
		m_ImagesBlack.Mirror ();
	}

#ifndef _BCGSUITE_
	if (m_ImagesBlack.GetBitsPerPixel() < 32)
#endif
	{
		m_ImagesBlack.SetTransparentColor (clrTransparent);
	}

	CreateCopy (m_ImagesGray, RGB (128, 128, 128));
	CreateCopy (m_ImagesDkGray, RGB (72, 72, 72));
	CreateCopy (m_ImagesLtGray, RGB (192, 192, 192));
	CreateCopy (m_ImagesWhite, RGB (255, 255, 255));
	CreateCopy (m_ImagesBlack2, RGB (0, 0, 0));

#ifndef _BCGSUITE_
	if (m_ImagesBlack.IsValid() && dblScale != 1.0)
	{
		m_ImagesBlack.SmoothResize (dblScale);
		m_ImagesGray.SmoothResize (dblScale);
		m_ImagesDkGray.SmoothResize (dblScale);
		m_ImagesLtGray.SmoothResize (dblScale);
		m_ImagesWhite.SmoothResize (dblScale);
		m_ImagesBlack2.SmoothResize (dblScale);
	}
#endif

	m_bInitializing = FALSE;
	
	return TRUE;
}
//****************************************************************************************
CSize CBCGPMenuImages::Size()
{
	if (m_bCleannedUpGlobally)
	{
		return CSize(0, 0);
	}

	if (m_bInitializing)
	{
		double dblScale = 1.0;
		CSize size = GetOriginalImageSize(NULL, &dblScale);

		if (dblScale != 1.0)
		{
			size.cx = (int)(size.cx * dblScale);
			size.cy = (int)(size.cy * dblScale);
		}

		return size;
	}

	Initialize ();
	return m_ImagesBlack.GetImageSize ();
}
//****************************************************************************************
void CBCGPMenuImages::Draw (CDC* pDC, IMAGES_IDS id, const CPoint& ptImage,
						CBCGPMenuImages::IMAGE_STATE state,
						const CSize& sizeImage/* = CSize (0, 0)*/)
{
	if (!Initialize ())
	{
		return;
	}

	CBCGPDrawState ds;

	CBCGPToolBarImages& images = (state == ImageBlack) ? m_ImagesBlack :
					(state == ImageGray) ? m_ImagesGray : 
					(state == ImageDkGray) ? m_ImagesDkGray : 
					(state == ImageLtGray) ? m_ImagesLtGray : 
					(state == ImageWhite) ? m_ImagesWhite : m_ImagesBlack2;

	if (images.PrepareDrawImage (ds, sizeImage))
	{
		images.Draw (pDC, ptImage.x, ptImage.y, id);
		images.EndDrawImage (ds);
	}
}
//****************************************************************************************
void CBCGPMenuImages::Draw (CDC* pDC, IMAGES_IDS id, const CRect& rectImage,
						CBCGPMenuImages::IMAGE_STATE state,
						const CSize& sizeImageDest/* = CSize (0, 0)*/)
{
	const CSize sizeImage = 
		(sizeImageDest == CSize (0, 0)) ? Size () : sizeImageDest;

	CPoint ptImage (
		rectImage.left + (rectImage.Width () - sizeImage.cx) / 2 + ((rectImage.Width () - sizeImage.cx) % 2), 
		rectImage.top + (rectImage.Height () - sizeImage.cy) / 2 + ((rectImage.Height () - sizeImage.cy) % 2));

	Draw (pDC, id, ptImage, state, sizeImageDest);
}
//*************************************************************************************
void CBCGPMenuImages::CleanUp (BOOL bIsGlobalCleanUp/* = FALSE*/)
{
	if (m_bInitializing)
	{
		return;
	}

	if (m_ImagesBlack.GetCount () > 0)
	{
		m_ImagesBlack.Clear ();
		m_ImagesGray.Clear ();
		m_ImagesDkGray.Clear ();
		m_ImagesLtGray.Clear ();
		m_ImagesWhite.Clear ();
		m_ImagesBlack2.Clear ();
	}

	m_bCleannedUpGlobally = bIsGlobalCleanUp;
}
//***********************************************************************************
void CBCGPMenuImages::CreateCopy (CBCGPToolBarImages& images, COLORREF clr)
{
	if (m_bCleannedUpGlobally)
	{
		return;
	}

	m_ImagesBlack.CopyTo (images);

#ifndef _BCGSUITE_
	if (m_ImagesBlack.GetBitsPerPixel() == 32 || m_ImagesBlack.IsSVGList())
	{
		images.AdaptColors(RGB (0, 0, 0), clr, FALSE);
	}
	else
#endif
	{
		images.MapTo3dColors (TRUE, RGB (0, 0, 0), clr);
	}
}
//***********************************************************************************
void CBCGPMenuImages::SetColor (CBCGPMenuImages::IMAGE_STATE state,
							COLORREF color)
{
	if (m_bCleannedUpGlobally)
	{
		return;
	}

	Initialize ();

	CBCGPLocalResource locaRes;

	CBCGPToolBarImages imagesTmp;

	UINT nImageResID = 0;
	double dblScale = 1.0;

	imagesTmp.SetImageSize(GetOriginalImageSize(&nImageResID, &dblScale));
	
#ifndef _BCGSUITE_
	//----------------------------
	// Try to load SVG image list:
	//----------------------------
	if (!imagesTmp.Load(IDB_BCGBARRES_MENU_IMAGES_SVG))
#endif
	{
		imagesTmp.Load (nImageResID);
	}

#ifndef _BCGSUITE_
	if (globalData.m_bIsBlackHighContrast)
	{
		imagesTmp.AdaptColors(RGB(0, 0, 0), RGB(255, 255, 255));
	}
#endif

#ifndef _BCGSUITE_
	if (imagesTmp.GetBitsPerPixel() < 32)
#endif
	{
		imagesTmp.SetTransparentColor (clrTransparent);
	}

#ifndef _BCGSUITE_
	if (imagesTmp.IsRTL ())
	{
		CBCGPToolBarImages::MirrorBitmap (imagesTmp.m_hbmImageWell, imagesTmp.GetImageSize ().cx);
	}
#endif

	CBCGPToolBarImages& images = (state == ImageBlack) ? m_ImagesBlack :
					(state == ImageGray) ? m_ImagesGray : 
					(state == ImageDkGray) ? m_ImagesDkGray : 
					(state == ImageLtGray) ? m_ImagesLtGray : 
					(state == ImageWhite) ? m_ImagesWhite : m_ImagesBlack2;

	if (color != (COLORREF)-1)
	{
#ifndef _BCGSUITE_
		if (imagesTmp.GetBitsPerPixel() == 32 || imagesTmp.IsSVGList())
		{
			imagesTmp.AdaptColors(RGB (0, 0, 0), color, FALSE);
		}
		else
#endif
		{
			imagesTmp.MapTo3dColors (TRUE, RGB (0, 0, 0), color);
		}
	}

#ifndef _BCGSUITE_
	if (!m_bInitializing)
	{
		imagesTmp.SmoothResize(dblScale);
	}
#endif

	images.Clear ();
	imagesTmp.CopyTo (images);
}
//************************************************************************************
CBCGPMenuImages::IMAGE_STATE CBCGPMenuImages::GetStateByColor(COLORREF color, BOOL bIsBackgroundColor)
{
	if (color == (COLORREF)-1 || CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		return CBCGPMenuImages::ImageBlack;
	}
	
	BOOL bIsDarkColor = CBCGPDrawManager::IsDarkColor(color);

	if (bIsBackgroundColor)
	{
		return bIsDarkColor ? CBCGPMenuImages::ImageWhite : CBCGPMenuImages::ImageBlack;
	}
	else
	{
		return !bIsDarkColor ? CBCGPMenuImages::ImageWhite : CBCGPMenuImages::ImageBlack;
	}
}
