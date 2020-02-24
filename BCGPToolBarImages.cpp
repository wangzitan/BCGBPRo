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

// BCGPToolBarImages.cpp: implementation of the CBCGPToolBarImages class.
//
//////////////////////////////////////////////////////////////////////

#include "Stdafx.h"
#include "BCGGlobals.h"
#include "BCGPMath.h"
#include "BCGPToolBarImages.h"

#ifndef _BCGPCHART_STANDALONE
#include "BCGPToolBar.h"
#include "BCGPLocalResource.h"
#include "bcgprores.h"
#include "BCGPVisualManager.h"
#include "BCGPPngImage.h"
#endif

#include "BCGPDrawManager.h"
#include "BCGPImageProcessing.h"
#include "BCGPSVGImage.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#ifdef _DEBUG
	BOOL g_bToolBarImagesTrace = TRUE;
#endif

static BOOL WriteDIB( LPCTSTR szFile, HANDLE hDIB);
static HANDLE DDBToDIB (HBITMAP bitmap, DWORD dwCompression);

BOOL CBCGPToolBarImages::m_bDisableTrueColorAlpha = TRUE;
CCriticalSection CBCGPToolBarImages::g_cs;
BOOL CBCGPToolBarImages::m_bMultiThreaded = FALSE;
BOOL CBCGPToolBarImages::m_bIsDrawOnGlass = FALSE;
HINSTANCE CBCGPToolBarImages::m_hinstMSIMGDLL = NULL;
ALPHABLEND CBCGPToolBarImages::m_pfAlphaBlend = NULL;
TRANSPARENTBLT CBCGPToolBarImages::m_pfTransparentBlt = NULL;
BYTE CBCGPToolBarImages::m_nDisabledImageAlpha = 127;
BYTE CBCGPToolBarImages::m_nFadedImageAlpha = 150;
BOOL CBCGPToolBarImages::m_bIsRTL = FALSE;
double CBCGPToolBarImages::m_dblColorTolerance = 0.2;

// globals for fast drawing (shared globals)
static HDC hDCGlyphs = NULL;
static HDC hDCMono = NULL;

/*
	DIBs use RGBQUAD format:
		0xbb 0xgg 0xrr 0x00

	Reasonably efficient code to convert a COLORREF into an
	RGBQUAD is byte-order-dependent, so we need different
	code depending on the byte order we're targeting.
*/

#define RGB_TO_RGBQUAD(r,g,b)   (RGB(b,g,r))
#define CLR_TO_RGBQUAD(clr)     (RGB(GetBValue(clr), GetGValue(clr), GetRValue(clr)))
#define RGBQUAD_TO_CLR(clr)     (RGB(GetBValue(clr), GetGValue(clr), GetRValue(clr)))

// Raster Ops
#define ROP_DSPDxax  0x00E20746L
#define ROP_PSDPxax  0x00B8074AL

// Internal images:
#define BCGImage_Light	0
#define BCGImage_Shadow	1
#define BCGImage_Gray	2

/////////////////////////////////////////////////////////////////////////////
// Init / Term

void CBCGPToolBarImages::CleanUp ()
{
	if (hDCMono != NULL)
	{
		::DeleteDC (hDCMono);
		hDCMono = NULL;
	}

	if (hDCGlyphs != NULL)
	{
		::DeleteDC (hDCGlyphs);
		hDCGlyphs = NULL;
	}

	if (m_hinstMSIMGDLL != NULL)
	{
		::FreeLibrary (m_hinstMSIMGDLL);
		m_hinstMSIMGDLL = NULL;
	}

	m_pfAlphaBlend = NULL;
	m_pfTransparentBlt = NULL;
}

// a special struct that will cleanup automatically
struct _AFX_TOOLBAR_TERM
{
	~_AFX_TOOLBAR_TERM()
	{
		CBCGPToolBarImages::CleanUp ();
	}
};

static const _AFX_TOOLBAR_TERM toolbarTerm;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPToolBarImages::CBCGPToolBarImages()
{
	m_bModified = FALSE;
	m_bReadOnly = FALSE;
	m_bIsTemporary = FALSE;
	m_iCount = 0;
	m_bIsGray = FALSE;
	m_nGrayPercentage = 0;

	m_hbmImageWell = NULL;
	m_hbmImageLight = NULL;
	m_hbmImageShadow = NULL;
	m_hbmImageGray = NULL;

	m_bUserImagesList = FALSE;

	// initialize the toolbar drawing engine
	static BOOL bInitialized;
	if (!bInitialized)
	{
		hDCGlyphs = CreateCompatibleDC(NULL);

		// Mono DC and Bitmap for disabled image
		hDCMono = ::CreateCompatibleDC(NULL);

		if (hDCGlyphs == NULL || hDCMono == NULL)
			AfxThrowResourceException();

#if _MSC_VER < 1800
		OSVERSIONINFO osvi;
		osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
		
		::GetVersionEx (&osvi);
		
		BOOL bIsWindowsNT4 = ((osvi.dwPlatformId == VER_PLATFORM_WIN32_NT) &&
			(osvi.dwMajorVersion < 5));
		
		BOOL bIsOSAlphaBlendingSupport = (osvi.dwMajorVersion > 4) ||
			((osvi.dwMajorVersion == 4) && (osvi.dwMinorVersion > 0));
		
		if (bIsWindowsNT4)
		{
			bIsOSAlphaBlendingSupport = FALSE;
		}
		
		if (bIsOSAlphaBlendingSupport)
#endif
		{
			if ((m_hinstMSIMGDLL = LoadLibrary (_T("msimg32.dll"))) != NULL)
			{
				m_pfAlphaBlend = 
					(ALPHABLEND)::GetProcAddress (m_hinstMSIMGDLL, "AlphaBlend");

				if (!globalData.bIsWindows9x)
				{
					m_pfTransparentBlt = 
						(TRANSPARENTBLT)::GetProcAddress (m_hinstMSIMGDLL, "TransparentBlt");
				}
			}
		}

		bInitialized = TRUE;
	}

	m_clrTransparentOriginal = m_clrTransparent = (COLORREF) -1;

	// UISG standard sizes
	m_sizeImage = CSize (16, 15);
	m_sizeImageOriginal = CSize (0, 0);
	m_sizeImageDest = CSize (0, 0);
	m_rectLastDraw = CRect (0, 0, 0, 0);
	m_rectSubImage = CRect (0, 0, 0, 0);
	m_bStretch = FALSE;
	m_pBmpOriginal = NULL;

	m_bFadeInactive = FALSE;
	m_nBitsPerPixel = 0;

	m_nLightPercentage = 130;
	m_bAlwaysLight = FALSE;

	m_bMapTo3DColors = TRUE;
	m_bAutoCheckPremlt = FALSE;
	m_bCreateMonoDC = TRUE;

	m_dblScale = 1.0;
	m_bAutoGrayAfterReLoad = FALSE;
	m_dblGrayLumRatio = 1.0;
	
	m_nSVGCount = 0;

	OnSysColorChange ();
}
//*******************************************************************************
CBCGPToolBarImages::~CBCGPToolBarImages()
{
	ASSERT (m_dcMem.GetSafeHdc () == NULL);
	ASSERT (m_bmpMem.GetSafeHandle () == NULL);
	ASSERT (m_pBmpOriginal == NULL);

#ifndef BCGP_EXCLUDE_SVG
	CleanUpSVG();
#endif

	if (!m_bIsTemporary)
	{
		AfxDeleteObject((HGDIOBJ*)&m_hbmImageWell);
		CleanUpInternalImages();
	}
}
//*******************************************************************************
void CBCGPToolBarImages::CleanUpInternalImages()
{
	AfxDeleteObject((HGDIOBJ*)&m_hbmImageLight);
	m_hbmImageLight = NULL;

	AfxDeleteObject((HGDIOBJ*)&m_hbmImageShadow);
	m_hbmImageShadow = NULL;

	AfxDeleteObject((HGDIOBJ*)&m_hbmImageGray);
	m_hbmImageGray = NULL;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::Load (UINT uiResID, HINSTANCE hinstRes, BOOL bAdd)
{
	return LoadStr (MAKEINTRESOURCE (uiResID), hinstRes, bAdd);
}
//*******************************************************************************
BOOL CBCGPToolBarImages::LoadStr (LPCTSTR lpszResourceName, HINSTANCE hinstRes, BOOL bAdd)
{
	if (m_bIsTemporary)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	if (lpszResourceName == NULL)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	UINT uiResID = IS_INTRESOURCE (lpszResourceName) ? (UINT)((UINT_PTR)(lpszResourceName)) : 0;

	if (!bAdd)
	{
		AfxDeleteObject((HGDIOBJ*)&m_hbmImageWell);     // get rid of old one
		m_lstOrigResIds.RemoveAll ();
		m_lstOrigResInstances.RemoveAll ();
		m_mapOrigResOffsets.RemoveAll ();
		m_nBitsPerPixel = 0;

#ifndef BCGP_EXCLUDE_SVG
		CleanUpSVG();
#endif
	}
	else if (uiResID != 0 &&
		m_lstOrigResIds.Find (uiResID) != NULL)	// Already loaded, do nothing
	{
		return TRUE;
	}

	HBITMAP hbmp = NULL;

	//---------------------------------
	// 1-st try to load SVG image list:
	//---------------------------------
#ifndef BCGP_EXCLUDE_SVG

	CBCGPSVGImageList arSVG;
	if (arSVG.LoadStr(lpszResourceName, hinstRes))
	{
		const int nCount = (int)arSVG.GetSize();
		if (nCount > 0)
		{
			// Add empty BMP:
			hbmp = CBCGPDrawManager::CreateBitmap_32(CSize(m_sizeImage.cx * nCount, m_sizeImage.cy), NULL);

			if (hbmp == NULL)
			{
				ASSERT (FALSE);
				return FALSE;
			}
			
			for (int i = 0; i < nCount; i++)
			{
				CBCGPSVGImage* pSVG = arSVG[i];
				if (pSVG != NULL)
				{
					arSVG[i] = NULL;
					m_nSVGCount++;
				}

				m_arSVG.Add(pSVG);
			}
			
			m_nBitsPerPixel = 32;
		}
	}
	
#endif // BCGP_EXCLUDE_SVG

	if (hbmp == NULL)
	{
		//-----------------------
		// Try to load PNG image:
		//-----------------------
		CBCGPPngImage pngImage;
		if (pngImage.Load (lpszResourceName, hinstRes))
		{
			hbmp = (HBITMAP) pngImage.Detach ();
		}
		else
		{
			if (hinstRes == NULL)
			{
				hinstRes = AfxFindResourceHandle (lpszResourceName, RT_BITMAP);
			}

			UINT uiLoadImageFlags = LR_CREATEDIBSECTION;
			if (m_bMapTo3DColors && !globalData.m_bIsBlackHighContrast)
			{
				uiLoadImageFlags |= LR_LOADMAP3DCOLORS;
			}

			hbmp = (HBITMAP) ::LoadImage (
				hinstRes,
				lpszResourceName,
				IMAGE_BITMAP,
				0, 0,
				uiLoadImageFlags);
		}
	}

	if (hbmp == NULL)
	{
#ifdef _DEBUG
		if (g_bToolBarImagesTrace)
		{
			if (uiResID != 0)
			{
				TRACE(_T("Can't load image: %x. GetLastError() = %x\n"), 
					uiResID,
					GetLastError ());
			}
			else
			{
				TRACE(_T("Can't load bitmap: %s. GetLastError() = %x\n"), 
					lpszResourceName,
					GetLastError ());
			}
		}
#endif

		return FALSE;
	}

	BITMAP bmp;
	if (::GetObject (hbmp, sizeof (BITMAP), &bmp) == 0)
	{
		ASSERT (FALSE);
		::DeleteObject (hbmp);
		return FALSE;
	}

	if (bmp.bmBitsPixel >= 32)
	{
		PreMultiplyAlpha (hbmp);
	}
	else if ((bmp.bmBitsPixel > 8 && m_bMapTo3DColors) || globalData.m_bIsBlackHighContrast)
	{
		//------------------------------------------------
		// LR_LOADMAP3DCOLORS don't support > 8bpp images,
		// we should convert it now:
		//------------------------------------------------
		MapBmpTo3dColors (hbmp, FALSE);
	}

	if (bAdd && m_nBitsPerPixel < bmp.bmBitsPixel && m_hbmImageWell != NULL)
	{
		// Convert m_hbmImageWell to bmp.bmBitsPixel!
		ChangeBpp(hbmp);
	}

	m_nBitsPerPixel = max (m_nBitsPerPixel, bmp.bmBitsPixel);

	if (bAdd)
	{
		if (uiResID != 0)
		{
			m_mapOrigResOffsets.SetAt (uiResID, m_iCount);
		}

		AddImage (hbmp);

		if (uiResID != 0)
		{
			m_lstOrigResIds.AddTail (uiResID);
			m_lstOrigResInstances.AddTail (hinstRes);
		}

		::DeleteObject (hbmp);
	}
	else
	{
		m_hbmImageWell = hbmp;
	}

	UpdateCount ();
	CleanUpInternalImages();

#ifndef BCGP_EXCLUDE_SVG
	while (m_arSVG.GetSize() < m_iCount)
	{
		m_arSVG.Add(NULL);
	}
#endif

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::Load (LPCTSTR lpszBmpFileName)
{
	if (m_bIsTemporary)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	ASSERT (lpszBmpFileName != NULL);

	AfxDeleteObject((HGDIOBJ*)&m_hbmImageWell);     // get rid of old one

#ifndef BCGP_EXCLUDE_SVG
	CleanUpSVG();
#endif

	CString strPath = lpszBmpFileName;

	//-----------------------------------------------------------------------
	// If the image path is not defined, try to open it in the EXE directory:
	//-----------------------------------------------------------------------
	if (strPath.Find (_T("\\")) == -1 &&
		strPath.Find (_T("/")) == -1 &&
		strPath.Find (_T(":")) == -1)
	{
		TCHAR lpszFilePath [_MAX_PATH];
		if (::GetModuleFileName (NULL, lpszFilePath, _MAX_PATH) > 0)
		{
			TCHAR path_buffer[_MAX_PATH];   
			TCHAR drive[_MAX_DRIVE];
			TCHAR dir[_MAX_DIR];
			TCHAR fname[_MAX_FNAME];   
			TCHAR ext[_MAX_EXT];

#if _MSC_VER < 1400
			_tsplitpath (lpszFilePath, drive, dir, NULL, NULL);
			_tsplitpath (lpszBmpFileName, NULL, NULL, fname, ext);
			
			_tmakepath (path_buffer, drive, dir, fname, ext);
#else
			_tsplitpath_s (lpszFilePath, drive, _MAX_DRIVE, dir, _MAX_DIR, NULL, 0, NULL, 0);
			_tsplitpath_s (lpszBmpFileName, NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);

			_tmakepath_s (path_buffer, drive, dir, fname, ext);
#endif

			strPath = path_buffer;
		}
	}

	BOOL bPNG = FALSE;
	BOOL bSVG = FALSE;

	{
		CString ext;

#if _MSC_VER < 1400
		_tsplitpath (strPath, NULL, NULL, NULL, ext.GetBuffer (_MAX_EXT));
#else
		_tsplitpath_s (strPath, NULL, 0, NULL, 0, NULL, 0, ext.GetBuffer (_MAX_EXT), _MAX_EXT);
#endif

		ext.ReleaseBuffer ();
		
		bPNG = ext.CompareNoCase (_T(".png")) == 0;
		bSVG = ext.CompareNoCase (_T(".svg")) == 0;
	}

	//--------------------------------
	// Load images from the disk file:
	//--------------------------------
#ifndef _BCGPCHART_STANDALONE
	if (bSVG)
	{
#ifndef BCGP_EXCLUDE_SVG
		CBCGPSVGImageList arSVG;
		if (arSVG.LoadFromFile(strPath))
		{
			const int nCount = (int)arSVG.GetSize();
			if (nCount > 0)
			{
				// Add empty BMP:
				m_hbmImageWell = CBCGPDrawManager::CreateBitmap_32(CSize(m_sizeImage.cx * nCount, m_sizeImage.cy), NULL);
				
				if (m_hbmImageWell == NULL)
				{
					ASSERT (FALSE);
					return FALSE;
				}
				
				for (int i = 0; i < nCount; i++)
				{
					CBCGPSVGImage* pSVG = arSVG[i];
					if (pSVG != NULL)
					{
						arSVG[i] = NULL;
						m_nSVGCount++;
					}
					
					m_arSVG.Add(pSVG);
				}
				
				m_nBitsPerPixel = 32;
			}
		}
#endif // BCGP_EXCLUDE_SVG
	}
	else if (bPNG)
	{
		CBCGPPngImage pngImage;
		if (pngImage.LoadFromFile (strPath))
		{
			m_hbmImageWell = (HBITMAP) pngImage.Detach ();
		}
	}
	else
#endif
	{
		UINT uiLoadImageFlags = LR_LOADFROMFILE | LR_CREATEDIBSECTION;
		if (m_bMapTo3DColors)
		{
			uiLoadImageFlags |= LR_LOADMAP3DCOLORS;
		}

		m_hbmImageWell = (HBITMAP) ::LoadImage (
			AfxGetInstanceHandle (),
			strPath,
			IMAGE_BITMAP,
			0, 0,
			uiLoadImageFlags);
	}

	if (m_hbmImageWell == NULL)
	{
#ifdef _DEBUG
		if (g_bToolBarImagesTrace)
		{
			TRACE(_T("Can't load bitmap: %s. GetLastError() = %x\r\n"), 
				lpszBmpFileName,
				GetLastError ());
		}
#endif

		return FALSE;
	}

	BITMAP bmp;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
	{
		ASSERT (FALSE);
		::DeleteObject (m_hbmImageWell);
		m_hbmImageWell = NULL;
		return FALSE;
	}

	m_bUserImagesList = TRUE;
	m_strUDLPath = strPath;

	if (::GetFileAttributes (strPath) & FILE_ATTRIBUTE_READONLY)
	{
		m_bReadOnly = TRUE;
	}

	m_nBitsPerPixel = bmp.bmBitsPixel;

	if (bmp.bmBitsPixel >= 32)
	{
		PreMultiplyAlpha (m_hbmImageWell);
	}
	else if (bmp.bmBitsPixel > 8 && m_bMapTo3DColors)
	{
		//------------------------------------------------
		// LR_LOADMAP3DCOLORS don't support > 8bpp images,
		// we should convert it now:
		//------------------------------------------------
		MapTo3dColors (FALSE);
	}

	UpdateCount ();
	CleanUpInternalImages();

#ifndef BCGP_EXCLUDE_SVG
	while (m_arSVG.GetSize() < m_iCount)
	{
		m_arSVG.Add(NULL);
	}
#endif
	return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::PrepareDrawImage (CBCGPDrawState& ds,
										  CSize sizeImageDest,
										  BOOL bFadeInactive)
{
	if (m_hbmImageWell == NULL)
	{
		return FALSE;
	}

	if (m_bMultiThreaded)
	{
		g_cs.Lock ();
	}

	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		if (sizeImageDest.cx <= 0 || sizeImageDest.cy <= 0)
		{
			m_sizeImageDest = m_sizeImage;
		}
		else
		{
			m_sizeImageDest = sizeImageDest;
		}
		
		return TRUE;
	}

	if (bFadeInactive && m_nBitsPerPixel < 32 && m_hbmImageLight == NULL)
	{
		UpdateInternalImage (BCGImage_Light);
	}

	if (ds.m_bAutoGrayScale && m_hbmImageGray == NULL)
	{
		UpdateInternalImage (BCGImage_Gray);
	}

#ifndef _BCGPCHART_STANDALONE
	if (m_hbmImageShadow == NULL &&
		CBCGPVisualManager::GetInstance ()->IsShadowHighlightedImage () &&
		!globalData.IsHighContastMode ())
	{
		UpdateInternalImage (BCGImage_Shadow);
	}
#endif
	m_bStretch = FALSE;

	if (m_hbmImageLight == NULL ||
		(m_nBitsPerPixel > 4 && !m_bAlwaysLight) || m_nBitsPerPixel == 0)
	{
		// Down't fade 256+ or unknown bitmaps
		bFadeInactive = FALSE;
	}

	m_bFadeInactive = bFadeInactive;

	ASSERT(m_hbmImageWell != NULL);
	ASSERT (m_dcMem.GetSafeHdc () == NULL);
	ASSERT (m_bmpMem.GetSafeHandle () == NULL);
	ASSERT (m_pBmpOriginal == NULL);

	// We need to kick-start the bitmap selection process.
	ds.hbmOldGlyphs = (HBITMAP)SelectObject (hDCGlyphs, 
		(ds.m_bAutoGrayScale && m_hbmImageGray != NULL) ? m_hbmImageGray : bFadeInactive && m_nBitsPerPixel < 32 ? m_hbmImageLight : m_hbmImageWell);

	if (ds.hbmOldGlyphs == NULL)
	{
		TRACE0("Error: can't draw toolbar.\r\n");

		if (m_bMultiThreaded)
		{
			g_cs.Unlock ();
		}

		return FALSE;
	}

	if (m_bCreateMonoDC)
	{
		ds.hbmMono = CreateBitmap (m_sizeImage.cx + 2, m_sizeImage.cy + 2,
						1, 1, NULL);
		ds.hbmMonoOld = (HBITMAP)SelectObject(hDCMono, ds.hbmMono);

		if (ds.hbmMono == NULL || ds.hbmMonoOld == NULL)
		{
			TRACE0("Error: can't draw toolbar.\r\n");
			AfxDeleteObject((HGDIOBJ*)&ds.hbmMono);

			if (m_bMultiThreaded)
			{
				g_cs.Unlock ();
			}

			return FALSE;
		}
	}

	if (sizeImageDest.cx <= 0 || sizeImageDest.cy <= 0)
	{
		m_sizeImageDest = m_sizeImage;
	}
	else
	{
		m_sizeImageDest = sizeImageDest;
	}

	COLORREF clrTransparent = m_nBitsPerPixel == 32 && m_pfAlphaBlend != NULL ?
		(COLORREF) -1 : m_clrTransparent;

	if (m_sizeImageDest != m_sizeImage || clrTransparent != (COLORREF) -1)
	{
		CWindowDC dc (NULL);

		m_bStretch = (m_sizeImageDest != m_sizeImage);

		m_dcMem.CreateCompatibleDC (NULL);	// Assume display!
		m_bmpMem.CreateCompatibleBitmap (&dc, m_sizeImage.cx + 2, m_sizeImage.cy + 2);

		m_pBmpOriginal = m_dcMem.SelectObject (&m_bmpMem);
		ASSERT (m_pBmpOriginal != NULL);
	}

	return TRUE;
}
//*******************************************************************************
void CBCGPToolBarImages::EndDrawImage (CBCGPDrawState& ds)
{
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		m_sizeImageDest = CSize (0, 0);
		m_rectLastDraw = CRect (0, 0, 0, 0);

		if (m_bMultiThreaded)
		{
			g_cs.Unlock ();
		}
		
		return;
	}

	if (m_bCreateMonoDC)
	{
		SelectObject(hDCMono, ds.hbmMonoOld);
		AfxDeleteObject((HGDIOBJ*)&ds.hbmMono);
	}

	SelectObject(hDCGlyphs, ds.hbmOldGlyphs);

	m_sizeImageDest = CSize (0, 0);
	m_rectLastDraw = CRect (0, 0, 0, 0);

	COLORREF clrTransparent = m_nBitsPerPixel == 32 && m_pfAlphaBlend != NULL ?
		(COLORREF) -1 : m_clrTransparent;

	if (m_bStretch || clrTransparent != (COLORREF) -1)
	{
		ASSERT (m_pBmpOriginal != NULL);

		m_dcMem.SelectObject (m_pBmpOriginal);
		m_pBmpOriginal = NULL;

		::DeleteObject (m_bmpMem.Detach ());
		::DeleteDC (m_dcMem.Detach ());
	}

	m_bFadeInactive = FALSE;

	if (m_bMultiThreaded)
	{
		g_cs.Unlock ();
	}
}
//*******************************************************************************
void CBCGPToolBarImages::CreateMask(int iImage, BOOL bHilite, BOOL bHiliteShadow)
{
	// initalize whole area with 0's
	PatBlt(hDCMono, 0, 0, m_sizeImage.cx + 2, m_sizeImage.cy + 2, WHITENESS);

	COLORREF clrTransparent = m_nBitsPerPixel == 32 && m_pfAlphaBlend != NULL ?
		(COLORREF) -1 : m_clrTransparent;

	// create mask based on color bitmap
	// convert this to 1's
	SetBkColor (hDCGlyphs, 
		clrTransparent != -1 ? clrTransparent : globalData.clrBtnFace);

	::BitBlt(hDCMono, 0, 0, m_sizeImage.cx, m_sizeImage.cy,
		hDCGlyphs, iImage * m_sizeImage.cx, 0, SRCCOPY);

	if (bHilite)
	{
		// convert this to 1's
		SetBkColor(hDCGlyphs, globalData.clrBtnHilite);

		// OR in the new 1's
		::BitBlt(hDCMono, 0, 0, m_sizeImage.cx, m_sizeImage.cy,
			hDCGlyphs, iImage * m_sizeImage.cx, 0, SRCPAINT);

		if (bHiliteShadow)
		{
			::BitBlt(hDCMono, 1, 1, m_sizeImage.cx + 1, m_sizeImage.cy + 1,
				hDCMono, 0, 0, SRCAND);
		}
	}
}
//********************************************************************************
HBITMAP CBCGPToolBarImages::GetMask (int iImage)
{
	CBCGPDrawState	ds;

	PrepareDrawImage(ds, FALSE);
	CreateMask(iImage, FALSE, FALSE);

	CDC	memDCDest;
	CDC* pDCMono = CDC::FromHandle(hDCMono);
	ASSERT_VALID (pDCMono);

	memDCDest.CreateCompatibleDC (pDCMono);
	
	CBitmap bitmapMask;

	if (bitmapMask.CreateBitmap(m_sizeImage.cx, m_sizeImage.cy, 1, 1, NULL))
	{
		CBitmap* pOldBitmapDest	= memDCDest.SelectObject (&bitmapMask);

		memDCDest.BitBlt (0, 0, m_sizeImage.cx, m_sizeImage.cy, pDCMono, 0, 0, SRCCOPY);
		memDCDest.SelectObject (pOldBitmapDest);
	}
					
	EndDrawImage(ds);

	return (HBITMAP) bitmapMask.Detach ();
}
//*******************************************************************************
BOOL CBCGPToolBarImages::Draw (CDC* pDCDest, 
				int xDest, int yDest,
				int iImage,
				BOOL bHilite,
				BOOL bDisabled,
				BOOL bIndeterminate,
				BOOL bShadow,
				BOOL bInactive,
				BYTE alphaSrc/* = 255*/,
				BOOL bIsIgnoreAlpha/* = FALSE*/)
{
	if (iImage < 0 || iImage >= m_iCount)
	{
		return FALSE;
	}

	if (bShadow && globalData.m_nBitsPerPixel <= 8)
	{
		return TRUE;
	}

	m_rectLastDraw = CRect (CPoint (xDest, yDest), m_sizeImageDest);

	if (m_bStretch)
	{
		bHilite = FALSE;
		bIndeterminate = FALSE;
	}

#ifndef BCGP_EXCLUDE_SVG
	if (iImage < m_arSVG.GetSize())
	{
		CBCGPSVGImage* pSVG = m_arSVG[iImage];
		if (pSVG != NULL)
		{
			if (bShadow)
			{
				return TRUE;
			}

			pSVG->DoDraw(pDCDest, m_rectLastDraw, bDisabled, bDisabled ? 100 : 255, TRUE /* Cash bitmap*/);
			return TRUE;
		}
	}
#endif

	HBITMAP hBmpOriginal = NULL;
	if ((!bInactive || bDisabled) && m_bFadeInactive && m_nBitsPerPixel < 32)
	{
		hBmpOriginal = (HBITMAP) SelectObject (hDCGlyphs, m_hbmImageWell);
	}

	BOOL bStretchOld = m_bStretch;
	BOOL bAlphaStretch = 
		(m_nBitsPerPixel == 32 && m_pfAlphaBlend != NULL && m_bStretch);

	if (bAlphaStretch)
	{
		m_bStretch = FALSE;
	}

	COLORREF clrTransparent = (m_nBitsPerPixel == 32 || m_bIsDrawOnGlass) && 
								m_pfAlphaBlend != NULL
									? (COLORREF) -1
									: m_clrTransparent;

	BOOL bIsTransparent = (clrTransparent != (COLORREF) -1);

	COLORREF clrTransparentDisabled = clrTransparent;

	CDC* pDC = m_bStretch || bIsTransparent ? &m_dcMem : pDCDest;
	ASSERT_VALID(pDC);

	int x = m_bStretch || bIsTransparent ? 0 : xDest;
	int y = m_bStretch || bIsTransparent ? 0 : yDest;

	const int xOffset = m_rectSubImage.left;
	const int yOffset = m_rectSubImage.top;

	const int nWidth = m_rectSubImage.IsRectEmpty () ? m_sizeImage.cx : m_rectSubImage.Width ();
	const int nHeight = m_rectSubImage.IsRectEmpty () ? m_sizeImage.cy : m_rectSubImage.Height ();

	if (m_bStretch || bIsTransparent)
	{
		CRect rectImage (CPoint (0, 0), m_sizeImage);

		if (bIsTransparent && clrTransparent != globalData.clrBtnFace)
		{
			CBrush brBackgr (clrTransparent);
			pDC->FillRect (rectImage, &brBackgr);
		}
		else
		{
			pDC->FillRect (rectImage, &globalData.brBtnFace);
		}

		if (bDisabled && globalData.m_nBitsPerPixel == 16)
		{
			clrTransparentDisabled = pDC->GetPixel (rectImage.TopLeft ());
		}
	}

	BOOL bDisabledTrueColor = FALSE;

	if (bDisabled && m_nBitsPerPixel >= 24)
	{
		bDisabled = FALSE;
		bDisabledTrueColor = TRUE;
	}

	if (!bHilite && !bDisabled && !bShadow)
	{
		BOOL bIsReady = FALSE;

		if ((m_nBitsPerPixel == 32 || m_bIsDrawOnGlass) && m_pfAlphaBlend != NULL)
		{
			BLENDFUNCTION pixelblend = { AC_SRC_OVER, 0, 
				alphaSrc, (BYTE)(bIsIgnoreAlpha ? 0 : 1 /*AC_SRC_ALPHA*/ ) };

			if (bDisabledTrueColor)
			{
				pixelblend.SourceConstantAlpha = m_nDisabledImageAlpha;
			}

			if (bInactive && m_bFadeInactive)
			{
				pixelblend.SourceConstantAlpha = m_nFadedImageAlpha;
			}

			ASSERT (m_pfAlphaBlend != NULL);

			if (globalData.bIsWindows9x)
			{
				BITMAPINFO bi;

				// Fill in the BITMAPINFOHEADER
				bi.bmiHeader.biSize          = sizeof(BITMAPINFOHEADER);
				bi.bmiHeader.biWidth         = nWidth;
				bi.bmiHeader.biHeight        = nHeight;
				bi.bmiHeader.biPlanes        = 1;
				bi.bmiHeader.biBitCount      = 32;
				bi.bmiHeader.biCompression   = BI_RGB;
				bi.bmiHeader.biSizeImage     = nWidth * nHeight;
				bi.bmiHeader.biXPelsPerMeter = 0;
				bi.bmiHeader.biYPelsPerMeter = 0;
				bi.bmiHeader.biClrUsed       = 0;
				bi.bmiHeader.biClrImportant  = 0;

				CDC dcMem;
				dcMem.CreateCompatibleDC (NULL);

				COLORREF* pBits = NULL;
				HBITMAP hbmp = CreateDIBSection (
					dcMem.m_hDC, &bi, DIB_RGB_COLORS, (void **)&pBits,
					NULL, NULL);
				if (hbmp == NULL)
				{
					ASSERT (FALSE);
					return FALSE;
				}

				CBitmap bmpMem;
				bmpMem.Attach (hbmp);

				CBitmap* pBmpOriginal = dcMem.SelectObject (&bmpMem);

				dcMem.FillRect (CRect (0, 0, nWidth, nHeight), &globalData.brBtnFace);

				::BitBlt (dcMem.m_hDC, 0, 0,
					nWidth, nHeight,
					hDCGlyphs, iImage * m_sizeImage.cx + xOffset, yOffset, SRCCOPY);

				const CSize sizeDest = bAlphaStretch ? m_sizeImageDest : m_sizeImage;

				bIsReady = (*m_pfAlphaBlend) (pDC->m_hDC, x, y,
					sizeDest.cx, sizeDest.cy, 
					dcMem.m_hDC, 0, 0, 
					nWidth, nHeight, pixelblend);

				dcMem.SelectObject (pBmpOriginal);
			}
			else
			{
				const CSize sizeDest = bAlphaStretch ? m_sizeImageDest : m_sizeImage;

				if (m_nBitsPerPixel != 32)
				{
					BITMAPINFO bi;

					// Fill in the BITMAPINFOHEADER
					bi.bmiHeader.biSize          = sizeof(BITMAPINFOHEADER);
					bi.bmiHeader.biWidth         = nWidth;
					bi.bmiHeader.biHeight        = nHeight;
					bi.bmiHeader.biPlanes        = 1;
					bi.bmiHeader.biBitCount      = 32;
					bi.bmiHeader.biCompression   = BI_RGB;
					bi.bmiHeader.biSizeImage     = nWidth * nHeight;
					bi.bmiHeader.biXPelsPerMeter = 0;
					bi.bmiHeader.biYPelsPerMeter = 0;
					bi.bmiHeader.biClrUsed       = 0;
					bi.bmiHeader.biClrImportant  = 0;

					COLORREF* pBits = NULL;
					HBITMAP hbmp = CreateDIBSection (NULL, &bi, DIB_RGB_COLORS, (LPVOID*)&pBits, NULL, 0);

					if (hbmp == NULL)
					{
						ASSERT (FALSE);
						return FALSE;
					}

					CBitmap bmpMem;
					bmpMem.Attach (hbmp);

					CDC dcMem;
					dcMem.CreateCompatibleDC (NULL);
					CBitmap* pBmpOld = dcMem.SelectObject (&bmpMem);

					::BitBlt (dcMem.GetSafeHdc (), 
						0, 0, nWidth, nHeight, 
						hDCGlyphs, iImage * m_sizeImage.cx + xOffset, yOffset, SRCCOPY);

					int nSizeImage = nHeight * nWidth;

					if (m_clrTransparent == -1)
					{
						for (int i = 0; i < nSizeImage; i++)
						{
							*pBits |= 0xFF000000;
							pBits++;
						}
					}
					else
					{
						COLORREF clrTrans = RGB (GetBValue (m_clrTransparent),
												 GetGValue (m_clrTransparent),
												 GetRValue (m_clrTransparent));

						for (int i = 0; i < nSizeImage; i++)
						{
							if (*pBits != clrTrans)
							{
								*pBits |= 0xFF000000;
							}
							else
							{
								*pBits = (COLORREF)0;
							}

							pBits++;
						}
					}

					bIsReady = (*m_pfAlphaBlend) (pDC->m_hDC, x, y,
						sizeDest.cx, sizeDest.cy, 
						dcMem.GetSafeHdc (), 0, 0, 
						nWidth, nHeight, pixelblend);

					dcMem.SelectObject (pBmpOld);
				}
				else
				{
					bIsReady = (*m_pfAlphaBlend) (pDC->m_hDC, x, y,
						sizeDest.cx, sizeDest.cy, 
						hDCGlyphs, iImage * m_sizeImage.cx + xOffset, yOffset, 
						nWidth, nHeight, pixelblend);

					if (!bIsReady)
					{
						// The possible reason can be the following: the source rectangle must lie 
						// completely within the source surface, otherwise an error occurs and 
						// the function returns FALSE.
						BITMAP bitmap;
						if (::GetObject(m_hbmImageWell, sizeof(BITMAP), &bitmap) != 0)
						{
							int nSrcHeight = bitmap.bmHeight;

							bIsReady = (*m_pfAlphaBlend) (pDC->m_hDC, x, y,
								sizeDest.cx, sizeDest.cy, 
								hDCGlyphs, iImage * m_sizeImage.cx + xOffset, yOffset, 
								nWidth, nSrcHeight, pixelblend);
						}
					}
				}
			}
		}
		
		if (!bIsReady)
		{
			//----------------------
			// normal image version:
			//----------------------
			::BitBlt(pDC->m_hDC, x, y,
				nWidth, nHeight,
				hDCGlyphs, iImage * m_sizeImage.cx + xOffset, yOffset, SRCCOPY);

			if (bDisabledTrueColor)
			{
				CBCGPDrawManager dm (*pDC);
				dm.GrayRect (CRect (x, y, x + nWidth + 2, y + nHeight + 2), 
					-1, clrTransparentDisabled == -1 ? globalData.clrBtnFace : clrTransparentDisabled,
#ifdef _BCGPCHART_STANDALONE
					RGB(192, 192, 192));
#else
					CBCGPVisualManager::GetInstance ()->GetToolbarDisabledColor ());
#endif
			}
		}
	}
	else if (bShadow && m_hbmImageShadow != NULL)
	{
		HBITMAP hbmpCurr = 
			(HBITMAP) SelectObject (hDCGlyphs, m_hbmImageShadow);

		if (m_nBitsPerPixel == 32)
		{
			if (m_pfAlphaBlend != NULL)
			{
				BLENDFUNCTION pixelblend = { AC_SRC_OVER, 0, 
					alphaSrc, (BYTE)(bIsIgnoreAlpha ? 0 : 1 /*AC_SRC_ALPHA*/ ) };

				(*m_pfAlphaBlend) (pDC->m_hDC, x, y,
					nWidth, nHeight, 
					hDCGlyphs, iImage * m_sizeImage.cx + xOffset, yOffset,
					nWidth, nHeight, pixelblend);
			}
		}
		else
		{
			::BitBlt(pDC->m_hDC, x, y,
				nWidth, nHeight,
				hDCGlyphs, iImage * m_sizeImage.cx + xOffset, yOffset, SRCCOPY);
		}

		SelectObject (hDCGlyphs, hbmpCurr);
	}
	else
	{
		if (bDisabled || bIndeterminate || bShadow)
		{
			// disabled or indeterminate version
			CreateMask(iImage, TRUE, FALSE);

			pDC->SetTextColor(bShadow ? m_clrImageShadow : 0L); // 0's in mono -> 0 (for ROP)
			pDC->SetBkColor((COLORREF)0x00FFFFFFL); // 1's in mono -> 1

#ifndef _BCGPCHART_STANDALONE
			if (bDisabled && CBCGPVisualManager::GetInstance ()->IsEmbossDisabledImage ())
			{
				// disabled - draw the hilighted shadow
				HGDIOBJ hbrOld = pDC->SelectObject (globalData.hbrBtnHilite);
				if (hbrOld != NULL)
				{
					// draw hilight color where we have 0's in the mask
					::BitBlt(pDC->m_hDC, x + 1, y + 1,
						nWidth + 2, nHeight + 2,
						hDCMono, 0, 0, ROP_PSDPxax);

					pDC->SelectObject(hbrOld);
				}
			}
#endif
			//BLOCK: always draw the shadow
			{
				HGDIOBJ hbrOld = pDC->SelectObject(globalData.hbrBtnShadow);
				if (hbrOld != NULL)
				{
					// draw the shadow color where we have 0's in the mask
					::BitBlt(pDC->m_hDC, 
						x, y,
						nWidth + 2, nHeight + 2,
						hDCMono, 0, 0, ROP_PSDPxax);

					pDC->SelectObject(hbrOld);
				}
			}
		}

		// if it is checked do the dither brush avoiding the glyph
		if (bHilite || bIndeterminate)
		{
			CBrush* pBrOld = pDC->SelectObject(&globalData.brLight);
			if (pBrOld != NULL)
			{
				CreateMask(iImage, !bIndeterminate, bDisabled);

				pDC->SetTextColor(0L);              // 0 -> 0
				pDC->SetBkColor((COLORREF)0x00FFFFFFL); // 1 -> 1

				// only draw the dither brush where the mask is 1's
				::BitBlt(pDC->m_hDC, x, y, 
					nWidth, nHeight,
					hDCMono, 0, 0, ROP_DSPDxax);

				pDC->SelectObject(pBrOld);
			}
		}
	}

	if (m_bStretch)
	{
		TransparentBlt (pDCDest->GetSafeHdc (), xDest, yDest, 
						nWidth, nHeight, 
						pDC, 0, 0, 
						bIsTransparent ? clrTransparent : globalData.clrBtnFace,
						m_sizeImageDest.cx, m_sizeImageDest.cy);
	}
	else if (bIsTransparent)
	{
		TransparentBlt (pDCDest->GetSafeHdc (), xDest, yDest, 
						nWidth, nHeight, 
						pDC, 0, 0, clrTransparent);
	}

	if (hBmpOriginal != NULL)
	{
		SelectObject (hDCGlyphs, hBmpOriginal);
	}

	m_bStretch = bStretchOld;
	return TRUE;
}
//********************************************************************************
BOOL CBCGPToolBarImages::DrawEx (CDC* pDC,
				CRect rect,
				int iImageIndex,
				ImageAlignHorz horzAlign/* = ImageAlignHorzLeft*/,
				ImageAlignVert vertAlign/* = ImageAlignVertTop*/,
				CRect rectSrc/* = CRect (0, 0, 0, 0)*/,
				BYTE alphaSrc/* = 255*/,
				BOOL bIsIgnoreAlpha/* = FALSE*/)
{
	ASSERT_VALID (pDC);
	
	if (iImageIndex < 0)
	{
		return FALSE;
	}

	if (rectSrc.IsRectEmpty ())
	{
		rectSrc = CRect (CPoint (0, 0), m_sizeImage);
	}

    if (rectSrc.IsRectEmpty ())
    {
        return FALSE;
    }

    CRect rectDst (rect);

    if (horzAlign != ImageAlignHorzStretch)
    {
        BOOL bUpdate = TRUE;

        if (horzAlign == ImageAlignHorzLeft)
        {
            rectDst.right = rectDst.left + rectSrc.Width ();
        }
        else if (horzAlign == ImageAlignHorzRight)
        {
            rectDst.left = rectDst.right - rectSrc.Width ();
        }
        else if (horzAlign == ImageAlignHorzCenter)
        {
            rectDst.left += (rectDst.Width() - rectSrc.Width ()) / 2;
            rectDst.right = rectDst.left + rectSrc.Width ();
        }
        else
        {
            bUpdate = FALSE;
        }

        if (bUpdate)
        {
            CRect rt (rectDst);
            rectDst.IntersectRect (rectDst, rect);

            if (0 < rectDst.Width () && rectDst.Width () !=  rectSrc.Width ())
            {
                rectSrc.left += rectDst.left - rt.left;
                rectSrc.right = rectSrc.left + min(rectDst.Width (), rectSrc.Width ());
            }
        }
    }

    if (vertAlign != ImageAlignVertStretch)
    {
        BOOL bUpdate = TRUE;

        if (vertAlign == ImageAlignVertTop)
        {
            rectDst.bottom = rectDst.top + rectSrc.Height ();
        }
        else if (vertAlign == ImageAlignVertBottom)
        {
            rectDst.top = rectDst.bottom - rectSrc.Height ();
        }
        else if (vertAlign == ImageAlignVertCenter)
        {
            rectDst.top += (rectDst.Height() - rectSrc.Height ()) / 2;
            rectDst.bottom = rectDst.top + rectSrc.Height ();
        }
        else
        {
            bUpdate = FALSE;
        }

        if (bUpdate)
        {
            CRect rt (rectDst);
            rectDst.IntersectRect (rectDst, rect);

            if (0 < rectDst.Height () && rectDst.Height () !=  rectSrc.Height ())
            {
                rectSrc.top += rectDst.top - rt.top;
                rectSrc.bottom = rectSrc.top + min(rectDst.Height (), rectSrc.Height ());
            }
        }
    }

    if (rectSrc.IsRectEmpty () || rectDst.IsRectEmpty ())
    {
        return FALSE;
    }

#ifndef BCGP_EXCLUDE_SVG
	if (iImageIndex < m_arSVG.GetSize())
	{
		CBCGPSVGImage* pSVG = m_arSVG[iImageIndex];
		if (pSVG != NULL)
		{
			pSVG->DoDraw(pDC, rectDst, alphaSrc <= 128, alphaSrc, TRUE /* Cash bitmap*/);
			return TRUE;
		}
	}
#endif

	if (!globalData.bIsWindows9x)
	{
		if (m_bMultiThreaded)
		{
			g_cs.Lock ();
		}

		HBITMAP hbmOldGlyphs = (HBITMAP)SelectObject (hDCGlyphs, m_hbmImageWell);

		const int xOffset = rectSrc.left;
		const int yOffset = rectSrc.top;

		const int nWidth = rectSrc.IsRectEmpty () ? m_sizeImage.cx : rectSrc.Width ();
		const int nHeight = rectSrc.IsRectEmpty () ? m_sizeImage.cy : rectSrc.Height ();

		BOOL bRes = FALSE;
		BOOL bUseAlphaBlend = FALSE;

		if (m_pfAlphaBlend != NULL)
		{
			if (alphaSrc < 255 && m_nBitsPerPixel < 32)
			{
				bUseAlphaBlend = TRUE;
				bIsIgnoreAlpha = TRUE;
			}
			else if (m_nBitsPerPixel == 32)
			{
				bUseAlphaBlend = TRUE;
			}
		}

		if (bUseAlphaBlend)
		{
			BLENDFUNCTION pixelblend = { AC_SRC_OVER, 0, 
				alphaSrc, (BYTE)(bIsIgnoreAlpha ? 0 : 1 /*AC_SRC_ALPHA*/ ) };

			bRes = (*m_pfAlphaBlend) (pDC->m_hDC, rectDst.left, rectDst.top,
				rectDst.Width (), rectDst.Height (), 
				hDCGlyphs, iImageIndex * m_sizeImage.cx + xOffset, yOffset, 
				nWidth, nHeight, pixelblend);
		}
		else if (m_clrTransparent == -1)
		{
			bRes = ::StretchBlt (pDC->m_hDC, rectDst.left, rectDst.top,
				rectDst.Width (), rectDst.Height (), 
				hDCGlyphs, iImageIndex * m_sizeImage.cx + xOffset, yOffset, 
				nWidth, nHeight, SRCCOPY);
		}

		SelectObject (hDCGlyphs, hbmOldGlyphs);

		if (m_bMultiThreaded)
		{
			g_cs.Unlock ();
		}

		if (bRes)
		{
			return TRUE;
		}
	}

	BOOL bCreateMonoDC = m_bCreateMonoDC;
	m_bCreateMonoDC = FALSE;

	CBCGPDrawState ds;
	if (!PrepareDrawImage (ds, rectDst.Size ()))
	{
		m_bCreateMonoDC = bCreateMonoDC;
		return FALSE;
	}

	m_rectSubImage = rectSrc;

	BOOL bRes = Draw (pDC, rectDst.left, rectDst.top, iImageIndex, 
					FALSE, FALSE, FALSE, FALSE, FALSE, alphaSrc, bIsIgnoreAlpha);

	m_rectSubImage.SetRectEmpty ();

	EndDrawImage (ds);
	m_bCreateMonoDC = bCreateMonoDC;
	return bRes;
}
//********************************************************************************
void CBCGPToolBarImages::FillDitheredRect (CDC* pDC, const CRect& rect)
{
	ASSERT_VALID(pDC);

	if (m_bIsDrawOnGlass)
	{
		LOGBRUSH br; 
        globalData.brLight.GetLogBrush(&br);

		CBCGPDrawManager dm (*pDC);
		dm.DrawRect (rect, br.lbColor, (COLORREF)-1);
	}
	else
	{
		pDC->FillRect (&rect, &globalData.brLight);
	}
}
//********************************************************************************
void CBCGPToolBarImages::OnSysColorChange()
{
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return;
	}

	if (m_bIsTemporary)
	{
		return;
	}

	if (m_dblScale != 1.0)
	{
		m_dblScale = 1.0;
		m_nBitsPerPixel = 0;

		if (m_clrTransparentOriginal != (COLORREF)-1)
		{
			m_clrTransparent = m_clrTransparentOriginal;
			m_clrTransparentOriginal = (COLORREF)-1;
		}

		m_sizeImage = m_sizeImageOriginal;
		m_sizeImageOriginal = CSize (0, 0);
		m_sizeImageDest = CSize (0, 0);
		m_rectLastDraw = CRect (0, 0, 0, 0);
		m_rectSubImage = CRect (0, 0, 0, 0);
	}

	// re-color bitmap for toolbar
	if (m_hbmImageWell != NULL)
	{
		if (m_bUserImagesList)
		{
			Load (m_strUDLPath);
		}
		else
		{
			// Image was buit from the resources...
			if (m_lstOrigResIds.IsEmpty ())
			{
				return;
			}

			ASSERT (m_lstOrigResInstances.GetCount () == m_lstOrigResIds.GetCount ());

			AfxDeleteObject((HGDIOBJ*)&m_hbmImageWell);     // get rid of old one

			m_nBitsPerPixel = 0;

			POSITION posInst = m_lstOrigResInstances.GetHeadPosition ();

			for (POSITION pos = m_lstOrigResIds.GetHeadPosition (); 
				pos != NULL;)
			{
				ASSERT (posInst != NULL);

				UINT uiResId = m_lstOrigResIds.GetNext (pos);
				ASSERT (uiResId > 0);

				HINSTANCE hInst = m_lstOrigResInstances.GetNext (posInst);
				ASSERT (hInst != NULL);

				HBITMAP hbmp = NULL;

#ifndef BCGP_EXCLUDE_SVG
				CBCGPSVGImageList arSVG;
				if (arSVG.Load(uiResId, hInst))
				{
					const int nCount = (int)arSVG.GetSize();
					if (nCount > 0)
					{
						// Add empty BMP:
						hbmp = CBCGPDrawManager::CreateBitmap_32(CSize(m_sizeImage.cx * nCount, m_sizeImage.cy), NULL);
						
						if (hbmp == NULL)
						{
							ASSERT (FALSE);
						}
					}
				}
#endif
				if (hbmp == NULL)
				{
					CBCGPPngImage pngImage;
					if (pngImage.Load (uiResId, hInst))
					{
						hbmp = (HBITMAP) pngImage.Detach ();
					}
				}

				if (hbmp == NULL)
				{
					UINT uiLoadImageFlags = LR_CREATEDIBSECTION;

					if (m_bMapTo3DColors && !globalData.IsHighContastMode ())
					{
						uiLoadImageFlags |= LR_LOADMAP3DCOLORS;
					}

					hbmp = (HBITMAP) ::LoadImage (
						hInst,
						MAKEINTRESOURCE (uiResId),
						IMAGE_BITMAP,
						0, 0,
						uiLoadImageFlags);
				}

				BITMAP bmp;
				if (::GetObject (hbmp, sizeof (BITMAP), &bmp) == 0)
				{
					ASSERT (FALSE);
				}

				if (m_nBitsPerPixel < bmp.bmBitsPixel && m_hbmImageWell != NULL)
				{
					// Convert m_hbmImageWell to bmp.bmBitsPixel!
					ChangeBpp(hbmp);
				}

				m_nBitsPerPixel = max(bmp.bmBitsPixel, m_nBitsPerPixel);

				if (bmp.bmBitsPixel >= 32)
				{
					PreMultiplyAlpha (hbmp);
				}
				else if ((bmp.bmBitsPixel > 8 && m_bMapTo3DColors) || globalData.m_bIsBlackHighContrast)
				{
					//------------------------------------------------
					// LR_LOADMAP3DCOLORS don't support > 8bpp images,
					// we should convert it now:
					//------------------------------------------------
					MapBmpTo3dColors (hbmp, FALSE);
				}

				AddImage (hbmp);

				::DeleteObject (hbmp);
			}
		}

		if (m_bAutoGrayAfterReLoad)
		{
			ConvertToGrayScale(m_dblGrayLumRatio, TRUE);
		}
	}

	UpdateCount ();

	if (m_bIsRTL)
	{
		MirrorBitmap (m_hbmImageWell, m_sizeImage.cx);
	}

	CleanUpInternalImages();

	if (m_bIsGray)
	{
		GrayImages (m_nGrayPercentage);
	}

	m_clrImageShadow = globalData.clrBtnShadow;
}
//********************************************************************************
void CBCGPToolBarImages::UpdateCount ()
{
	if (m_hbmImageWell == NULL || m_sizeImage.cx == 0)
	{
		m_iCount = 0;
		return;
	}

	BITMAP bmp;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
	{
		m_iCount = 0;
		return;
	}

	int iWidth = bmp.bmWidth;
	m_iCount = iWidth / m_sizeImage.cx;
}

//////////////////////////////////////////
// Image editing methods:
//////////////////////////////////////////

int CBCGPToolBarImages::AddImage (HBITMAP hbmp, BOOL bSetBitPerPixel/* = FALSE*/)
{
	if (m_bIsTemporary)
	{
		ASSERT (FALSE);
		return -1;
	}

	BOOL bIsMirror = FALSE;

	if (m_bIsRTL)
	{
		bIsMirror = TRUE;

		hbmp = (HBITMAP) ::CopyImage (hbmp, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
		MirrorBitmap (hbmp, m_sizeImage.cx);
	}

	if (IsScaled ())
	{
		BITMAP bmpScale;
		if (::GetObject (hbmp, sizeof (BITMAP), &bmpScale) == 0)
		{
			return -1;
		}

		if (bmpScale.bmHeight != m_sizeImage.cy)
		{
			CBCGPToolBarImages imageForScale;
			imageForScale.m_hbmImageWell = hbmp;

			imageForScale.m_nBitsPerPixel = bmpScale.bmBitsPixel;

			imageForScale.SetImageSize (m_sizeImageOriginal);
			imageForScale.m_iCount = bmpScale.bmWidth / m_sizeImageOriginal.cx;
			imageForScale.SmoothResize (m_dblScale);
			imageForScale.m_bIsTemporary = TRUE;

			::DeleteObject (hbmp);
			hbmp = imageForScale.GetImageWell ();
		}
	}

	//-------------------------------------------------------
	// Create memory source DC and select an original bitmap:
	//-------------------------------------------------------
	CDC memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	HBITMAP hOldBitmapSrc = NULL;

	int iBitmapWidth;
	int iBitmapHeight;
	int iNewBitmapWidth;

	BITMAP bmp;
	if (::GetObject (hbmp, sizeof (BITMAP), &bmp) == 0)
	{
		return -1;
	}

	if (bSetBitPerPixel)	
	{
		m_nBitsPerPixel = bmp.bmBitsPixel;
	}

	iNewBitmapWidth = bmp.bmWidth;
	iBitmapHeight = bmp.bmHeight;

	if (m_hbmImageWell != NULL)
	{
		//-------------------------------
		// Get original bitmap attributes:
		//-------------------------------
		if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
		{
			return -1;
		}

		hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (m_hbmImageWell);
		if (hOldBitmapSrc == NULL)
		{
			return -1;
		}

		iBitmapWidth = bmp.bmWidth;
		iBitmapHeight = bmp.bmHeight;
	}
	else
	{
		iBitmapWidth = 0;

		hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (hbmp);
		if (hOldBitmapSrc == NULL)
		{
			return -1;
		}
	}

	//----------------------------------------------------------
	// Create a new bitmap compatible with the source memory DC
	// (original bitmap SHOULD BE ALREADY SELECTED!):
	//----------------------------------------------------------
	HBITMAP hNewBitmap = (HBITMAP) ::CreateCompatibleBitmap (memDCSrc,
									iBitmapWidth + iNewBitmapWidth,
									iBitmapHeight);
	if (hNewBitmap == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		return -1;
	}

	//------------------------------------------------------
	// Create memory destination DC and select a new bitmap:
	//------------------------------------------------------
	CDC memDCDst;
	memDCDst.CreateCompatibleDC (&memDCSrc);
	
	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hNewBitmap);
	if (hOldBitmapDst == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		::DeleteObject (hNewBitmap);
		return -1;
	}

	if (m_hbmImageWell != NULL)
	{
		//-----------------------------
		// Copy original bitmap to new:
		//-----------------------------
		memDCDst.BitBlt (0, 0, iBitmapWidth, iBitmapHeight,
					&memDCSrc, 0, 0, SRCCOPY);
	}
		
	//--------------------------------
	// Select a new image and copy it:
	//--------------------------------
	if (memDCSrc.SelectObject (hbmp) == NULL)
	{
		memDCDst.SelectObject (hOldBitmapDst);
		memDCSrc.SelectObject (hOldBitmapSrc);

		::DeleteObject (hNewBitmap);
		return -1;
	}

	memDCDst.BitBlt (iBitmapWidth, 0, iNewBitmapWidth, iBitmapHeight,
				&memDCSrc, 0, 0, SRCCOPY);

	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);

	if (m_hbmImageWell != NULL)
	{
		::DeleteObject (m_hbmImageWell);
	}

	m_hbmImageWell = hNewBitmap;
	m_bModified = TRUE;

	UpdateCount ();
	CleanUpInternalImages();

#ifndef BCGP_EXCLUDE_SVG
	while (m_arSVG.GetSize() < m_iCount)
	{
		m_arSVG.Add(NULL);
	}
#endif

	if (bIsMirror)
	{
		::DeleteObject (hbmp);
	}

	return m_iCount - 1;
}
//********************************************************************************
int CBCGPToolBarImages::AddImage (const CBCGPToolBarImages& imageList, int nIndex)
{
	if (nIndex < 0 || nIndex >= imageList.GetCount ())
	{
		ASSERT (FALSE);
		return -1;
	}

	CWindowDC dc (NULL);

	if (!IsScaled ())
	{
		m_sizeImage = imageList.m_sizeImage;
		m_sizeImageDest = imageList.m_sizeImageDest;	
		m_clrTransparent = imageList.m_clrTransparent;
		m_clrImageShadow = imageList.m_clrImageShadow;
		m_bFadeInactive = imageList.m_bFadeInactive;
		m_nBitsPerPixel = imageList.m_nBitsPerPixel;
	}

	CDC	memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (imageList.m_hbmImageWell);

	CDC	memDCDest;
	memDCDest.CreateCompatibleDC (NULL);

	CBitmap bitmap;
	DIBSECTION ds = {0};

	if (imageList.m_nBitsPerPixel >= 24 &&
		::GetObject (m_hbmImageWell, sizeof (DIBSECTION), &ds) == sizeof(DIBSECTION))
	{
		BITMAPINFO bi = {0};
		bi.bmiHeader.biSize        = sizeof (BITMAPINFOHEADER);
		bi.bmiHeader.biWidth       = imageList.m_sizeImage.cx;
		bi.bmiHeader.biHeight      = imageList.m_sizeImage.cy;
		bi.bmiHeader.biPlanes      = ds.dsBmih.biPlanes;
		bi.bmiHeader.biBitCount    = ds.dsBmih.biBitCount;
		bi.bmiHeader.biCompression = BI_RGB;

		COLORREF* pBits = NULL;
		HBITMAP hNewBitmap = ::CreateDIBSection (
			dc, &bi, DIB_RGB_COLORS, (void **)&pBits,
			NULL, NULL);
		bitmap.Attach (hNewBitmap);
	}
	else
	{
		bitmap.CreateCompatibleBitmap (&dc, imageList.m_sizeImage.cx, imageList.m_sizeImage.cy);
	}

	CBitmap* pOldBitmapDest	= memDCDest.SelectObject (&bitmap);

	memDCDest.BitBlt (0, 0, imageList.m_sizeImage.cx, imageList.m_sizeImage.cy, 
		&memDCSrc, nIndex * imageList.m_sizeImage.cx, 0, SRCCOPY);

	memDCDest.SelectObject (pOldBitmapDest);
	memDCSrc.SelectObject (hOldBitmapSrc);

	int nNewIndex = AddImage(bitmap, m_hbmImageWell == NULL || m_iCount == 0);

#ifndef BCGP_EXCLUDE_SVG
	if (nNewIndex >= 0)
	{
		if (nIndex < imageList.m_arSVG.GetSize() && imageList.m_arSVG[nIndex] != NULL)
		{
			m_arSVG[nNewIndex] = new CBCGPSVGImage(*imageList.m_arSVG[nIndex]);
			m_nSVGCount++;
		}
	}
#endif

	return nNewIndex;
}
//********************************************************************************
int CBCGPToolBarImages::AddImageList(const CBCGPToolBarImages& src, BOOL bSetBitPerPixel /*= FALSE*/)
{
#ifndef BCGP_EXCLUDE_SVG
	int nCount = m_iCount;
#endif

	int nRes = AddImage(src.GetImageWell(), bSetBitPerPixel);

#ifndef BCGP_EXCLUDE_SVG
	if (src.m_nSVGCount > 0)
	{
		for (int i = 0; i < src.m_iCount; i++)
		{
			if (i < src.m_arSVG.GetSize() && src.m_arSVG[i] != NULL)
			{
				m_arSVG[nCount + i] = new CBCGPSVGImage(*src.m_arSVG[i]);
				m_nSVGCount++;
			}
		}
	}
#endif

	return nRes;
}
//********************************************************************************
int CBCGPToolBarImages::AddIcon (HICON hIcon, BOOL bAlphaBlend/* = FALSE*/)
{
	CWindowDC dc (NULL);

	if (hIcon == NULL)
	{
		bAlphaBlend = FALSE;
	}

	CDC dcMem;
	dcMem.CreateCompatibleDC (NULL);

	CBitmap bmpMem;

	CSize sizeIcon = m_sizeImage;
	if (IsScaled ())
	{
		sizeIcon = m_sizeImageOriginal;
	}

	if (bAlphaBlend)
	{
		BITMAPINFO bi;

		// Fill in the BITMAPINFOHEADER
		bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bi.bmiHeader.biWidth = sizeIcon.cx;
		bi.bmiHeader.biHeight = sizeIcon.cy;
		bi.bmiHeader.biPlanes = 1;
		bi.bmiHeader.biBitCount = 32;
		bi.bmiHeader.biCompression = BI_RGB;
		bi.bmiHeader.biSizeImage = sizeIcon.cx * sizeIcon.cy;
		bi.bmiHeader.biXPelsPerMeter = 0;
		bi.bmiHeader.biYPelsPerMeter = 0;
		bi.bmiHeader.biClrUsed = 0;
		bi.bmiHeader.biClrImportant = 0;

		COLORREF* pBits = NULL;
		HBITMAP hbmp = CreateDIBSection (
			dcMem.m_hDC, &bi, DIB_RGB_COLORS, (void **)&pBits,
			NULL, NULL);
		if (hbmp == NULL)
		{
			ASSERT (FALSE);
			return -1;
		}

		bmpMem.Attach (hbmp);
	}
	else
	{
		bmpMem.CreateCompatibleBitmap (&dc, sizeIcon.cx, sizeIcon.cy);
	}

	CBitmap* pBmpOriginal = dcMem.SelectObject (&bmpMem);

	if (!bAlphaBlend)
	{
		dcMem.FillRect (CRect (0, 0, sizeIcon.cx, sizeIcon.cy), &globalData.brBtnFace);
	}

	if (hIcon != NULL)
	{
		dcMem.DrawState (CPoint (0, 0), sizeIcon, hIcon, DSS_NORMAL, (CBrush*) NULL);
	}

	dcMem.SelectObject (pBmpOriginal);

	if (bAlphaBlend)
	{
		m_nBitsPerPixel = 32;
		PreMultiplyAlpha (bmpMem, TRUE);
	}

	return AddImage (bmpMem);
}

#ifndef BCGP_EXCLUDE_SVG

int CBCGPToolBarImages::AddSVG(CBCGPSVGImage* pSVG)
{
	if (pSVG == NULL)
	{
		ASSERT(FALSE);
		return -1;
	}

	HBITMAP hbmp = CBCGPDrawManager::CreateBitmap_32(m_sizeImage, NULL);
	if (hbmp == NULL)
	{
		ASSERT(FALSE);
		return -1;
	}

	m_arSVG.Add(pSVG);
	m_nSVGCount++;

	int nRes = AddImage(hbmp, TRUE);
	DeleteObject(hbmp);

	return nRes;
}
//*******************************************************************************
int CBCGPToolBarImages::AddSVGList(const CBCGPSVGImageList& sprite)
{
	int nCount = (int)sprite.GetSize();
	if (nCount <= 0)
	{
		ASSERT(FALSE);
		return -1;
	}

	HBITMAP hbmp = CBCGPDrawManager::CreateBitmap_32(CSize(m_sizeImage.cx * nCount, m_sizeImage.cy), NULL);
	if (hbmp == NULL)
	{
		ASSERT(FALSE);
		return -1;
	}

	for (int i = 0; i < nCount; i++)
	{
		m_arSVG.Add(new CBCGPSVGImage(*sprite[i]));
		m_nSVGCount++;
	}

	int nRes = AddImage(hbmp, TRUE);
	DeleteObject(hbmp);

	return nRes;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::UpdateSVG(int iImage, CBCGPSVGImage* pSVG)
{
	if (iImage < 0 || iImage >= m_arSVG.GetSize())	// Wrong index
	{
		return FALSE;
	}
	
	if (m_arSVG[iImage] != NULL)
	{
		delete m_arSVG[iImage];
		m_arSVG[iImage] = NULL;
		m_nSVGCount--;
	}

	if (pSVG != NULL)
	{
		m_arSVG[iImage] = pSVG;
		m_nSVGCount++;
	}

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::ExportToSVGImageList(CBCGPSVGImageList& imageList)
{
	imageList.RemoveAll();
	
	for (int i = 0; i < m_iCount; i++)
	{
		CBCGPSVGImage* pSVG = (i < m_arSVG.GetSize() && m_arSVG[i] != NULL) ? new CBCGPSVGImage(*m_arSVG[i]) : new CBCGPSVGImage;
		imageList.Add(pSVG);
	}

	return TRUE;
}
//*******************************************************************************
void CBCGPToolBarImages::CleanUpSVG()
{
	if (!m_bIsTemporary)
	{
		for (int i = 0; i < (int)m_arSVG.GetSize(); i++)
		{
			CBCGPSVGImage* pSVG = m_arSVG[i];
			if (pSVG != NULL)
			{
				delete pSVG;
			}
		}
	}

	m_arSVG.RemoveAll();
	m_nSVGCount = 0;
}

#endif // BCGP_EXCLUDE_SVG

BOOL CBCGPToolBarImages::UpdateImage (int iImage, HBITMAP hbmp)
{
	if (m_bIsTemporary)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	if (!m_bUserImagesList)	// Only user images can be edited!
	{
		return FALSE;
	}

	CWindowDC	dc (NULL);
	CBitmap 	bitmap;
	CDC 		memDCSrc;
	CDC 		memDCDst;

	memDCSrc.CreateCompatibleDC (&dc);
	memDCDst.CreateCompatibleDC (&dc);
	
	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (m_hbmImageWell);
	HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (hbmp);

	memDCDst.BitBlt (m_sizeImage.cx * iImage, 0, m_sizeImage.cx, m_sizeImage.cy,
				&memDCSrc, 0, 0, SRCCOPY);

	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);

	m_bModified = TRUE;

	CleanUpInternalImages();

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::DeleteImage (int iImage)
{
	if (m_bIsTemporary)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	if (!m_bUserImagesList)	// Only user images can be edited!
	{
		return FALSE;
	}

	if (iImage < 0 || iImage >= GetCount ())	// Wrong index
	{
		return FALSE;
	}

	//-------------------------------
	// Get original bitmap attrbutes:
	//-------------------------------
	BITMAP bmp;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
	{
		return FALSE;
	}

	//-------------------------------------------------------
	// Create memory source DC and select an original bitmap:
	//-------------------------------------------------------
	CDC memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (m_hbmImageWell);
	if (hOldBitmapSrc == NULL)
	{
		return FALSE;
	}

	//----------------------------------------------------------
	// Create a new bitmap compatibel with the source memory DC
	// (original bitmap SHOULD BE ALREADY SELECTED!):
	//----------------------------------------------------------
	HBITMAP hNewBitmap = (HBITMAP) ::CreateCompatibleBitmap (memDCSrc,
									bmp.bmWidth - m_sizeImage.cx,
									bmp.bmHeight);
	if (hNewBitmap == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		return FALSE;
	}

	//------------------------------------------------------
	// Create memory destination DC and select a new bitmap:
	//------------------------------------------------------
	CDC memDCDst;
	memDCDst.CreateCompatibleDC (&memDCSrc);
	
	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hNewBitmap);
	if (hOldBitmapDst == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		::DeleteObject (hNewBitmap);
		return FALSE;
	}

	//-----------------------------
	// Copy original bitmap to new:
	//-----------------------------

	if (iImage != 0)
	{
		memDCDst.BitBlt (0, 0, 
					m_sizeImage.cx * iImage, bmp.bmHeight,
					&memDCSrc, 0, 0, SRCCOPY);
	}
	
	if (iImage != m_iCount - 1)
	{
		memDCDst.BitBlt (m_sizeImage.cx * iImage, 0, 
					(m_iCount - iImage - 1) * m_sizeImage.cx, bmp.bmHeight,
					&memDCSrc, 
					m_sizeImage.cx * (iImage + 1), 0, SRCCOPY);
	}

	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);

	::DeleteObject (m_hbmImageWell);

	m_hbmImageWell = hNewBitmap;
	m_bModified = TRUE;

	UpdateCount ();
	CleanUpInternalImages();

#ifndef BCGP_EXCLUDE_SVG
	CBCGPSVGImage* pSVG = m_arSVG[iImage];
	if (pSVG != NULL)
	{
		delete pSVG;
	}

	m_arSVG.RemoveAt(iImage);
#endif

	return TRUE;
}
//*******************************************************************************
HICON CBCGPToolBarImages::ExtractIcon(int nIndex)
{
	if (nIndex < 0 || nIndex >= GetCount ())	// Wrong index
	{
		return NULL;
	}

#ifndef BCGP_EXCLUDE_SVG
	if (nIndex < m_arSVG.GetSize())
	{
		CBCGPSVGImage* pSVG = m_arSVG[nIndex];
		if (pSVG != NULL)
		{
			return pSVG->ExportToIcon(m_sizeImage);
		}
	}
#endif

	CImageList imageList;
	if (ExportToImageList(imageList))
	{
		return imageList.ExtractIcon (nIndex);
	}

	return NULL;
}
//*******************************************************************************
HBITMAP CBCGPToolBarImages::ExtractBitmap(int nIndex)
{
#ifndef BCGP_EXCLUDE_SVG
	if (nIndex >= 0 && nIndex < m_arSVG.GetSize())
	{
		CBCGPSVGImage* pSVG = m_arSVG[nIndex];
		if (pSVG != NULL)
		{
			return pSVG->ExportToBitmap(m_sizeImage);
		}
	}
#endif

	HICON hIcon = ExtractIcon(nIndex);
	if (hIcon == NULL)
	{
		return NULL;
	}

	BOOL bAlphaBlend = m_nBitsPerPixel >= 32;

	CWindowDC dc(NULL);
	
	CDC dcMem;
	dcMem.CreateCompatibleDC(NULL);
	
	CBitmap bmpMem;
	
	if (bAlphaBlend)
	{
		HBITMAP hbmp = CBCGPDrawManager::CreateBitmap_32(m_sizeImage, NULL);
		if (hbmp == NULL)
		{
			ASSERT (FALSE);
			return NULL;
		}
		
		bmpMem.Attach(hbmp);
	}
	else
	{
		bmpMem.CreateCompatibleBitmap(&dc, m_sizeImage.cx, m_sizeImage.cy);
	}
	
	CBitmap* pBmpOriginal = dcMem.SelectObject(&bmpMem);
	
	if (!bAlphaBlend)
	{
		dcMem.FillRect(CRect (0, 0, m_sizeImage.cx, m_sizeImage.cy), &globalData.brBtnFace);
	}
	
	dcMem.DrawState(CPoint (0, 0), m_sizeImage, hIcon, DSS_NORMAL, (CBrush*)NULL);
	dcMem.SelectObject(pBmpOriginal);

	::DestroyIcon(hIcon);
	
	if (bAlphaBlend)
	{
		PreMultiplyAlpha(bmpMem, TRUE);
	}

	return (HBITMAP)bmpMem.Detach();
}
//*******************************************************************************
HICON CBCGPToolBarImages::ExtractSimplifiedIcon(int nIndex)
{
	if (nIndex < 0 || nIndex >= GetCount ())	// Wrong index
	{
		return NULL;
	}
	
#ifndef BCGP_EXCLUDE_SVG
	if (nIndex < m_arSVG.GetSize())
	{
		CBCGPSVGImage* pSVG = m_arSVG[nIndex];
		if (pSVG != NULL)
		{
			CBCGPSVGImage svg(*pSVG);
			svg.Simplify();

			return svg.ExportToIcon(m_sizeImage);
		}
	}
#endif	

	if (GetBitsPerPixel() < 32)
	{
		return NULL;
	}

	CBCGPToolBarImages images;
	images.SetImageSize(GetImageSize());

	HICON hIcon = ExtractIcon(nIndex);
	if (hIcon != NULL)
	{
		images.AddIcon(hIcon, TRUE);
		::DestroyIcon(hIcon);

		images.Simplify();
		return images.ExtractIcon(0);
	}

	return NULL;
}

#ifndef BCGP_EXCLUDE_SVG

CBCGPSVGImage* CBCGPToolBarImages::GetSVG(int nIndex) const
{
	if (nIndex < m_arSVG.GetSize())
	{
		return m_arSVG[nIndex];
	}
	
	return NULL;
}

#endif

COLORREF CBCGPToolBarImages::MapToSysColor (COLORREF color, BOOL bUseRGBQUAD)
{
	struct COLORMAP
	{
		// use DWORD instead of RGBQUAD so we can compare two RGBQUADs easily
		DWORD rgbqFrom;
		int iSysColorTo;
	};
	static const COLORMAP sysColorMap[] =
	{
		// mapping from color in DIB to system color
		{ RGB_TO_RGBQUAD(0x00, 0x00, 0x00),  COLOR_BTNTEXT },       // black
		{ RGB_TO_RGBQUAD(0x80, 0x80, 0x80),  COLOR_BTNSHADOW },     // dark grey
		{ RGB_TO_RGBQUAD(0xC0, 0xC0, 0xC0),  COLOR_BTNFACE },       // bright grey
		{ RGB_TO_RGBQUAD(0xFF, 0xFF, 0xFF),  COLOR_BTNHIGHLIGHT }   // white
	};
	const int nMaps = 4;

	// look for matching RGBQUAD color in original
	for (int i = 0; i < nMaps; i++)
	{
		if (color == sysColorMap[i].rgbqFrom)
		{
			return bUseRGBQUAD ? 
				CLR_TO_RGBQUAD (globalData.GetColor (sysColorMap[i].iSysColorTo)) :
				globalData.GetColor (sysColorMap[i].iSysColorTo);
		}
	}

	return color;
}
//********************************************************************************
COLORREF CBCGPToolBarImages::MapToSysColorAlpha (COLORREF color)
{
	BYTE r = GetRValue (color);
	BYTE g = GetGValue (color);
	BYTE b = GetBValue (color);

	const int nDelta = 10;

	if (abs (r - b) > nDelta || abs (r - g) > nDelta || abs (b - g) > nDelta)
	{
		return color;
	}

	return CBCGPDrawManager::PixelAlpha (globalData.clrBarFace,
		1. + ((double) r - 192) / 255,
		1. + ((double) g - 192) / 255,
		1. + ((double) b - 192) / 255);
}
//********************************************************************************
COLORREF CBCGPToolBarImages::MapFromSysColor (COLORREF color, BOOL bUseRGBQUAD)
{
	struct COLORMAP
	{
		// use DWORD instead of RGBQUAD so we can compare two RGBQUADs easily
		DWORD rgbTo;
		int iSysColorFrom;
	};
	static const COLORMAP sysColorMap[] =
	{
		// mapping from color in DIB to system color
		{ RGB(0x00, 0x00, 0x00),  COLOR_BTNTEXT },       // black
		{ RGB(0x80, 0x80, 0x80),  COLOR_BTNSHADOW },     // dark grey
		{ RGB(0xC0, 0xC0, 0xC0),  COLOR_BTNFACE },       // bright grey
		{ RGB(0xFF, 0xFF, 0xFF),  COLOR_BTNHIGHLIGHT }   // white
	};
	const int nMaps = 4;

	// look for matching RGBQUAD color in original
	for (int i = 0; i < nMaps; i++)
	{
		COLORREF clrSystem = globalData.GetColor (sysColorMap[i].iSysColorFrom);

		if (bUseRGBQUAD)
		{
			if (color == CLR_TO_RGBQUAD (clrSystem))
			{
				return CLR_TO_RGBQUAD (sysColorMap[i].rgbTo);
			}
		}
		else
		{
			if (color == clrSystem)
			{
				return sysColorMap[i].rgbTo;
			}
		}
	}

	return color;
}
//************************************************************************************
BOOL CBCGPToolBarImages::Save (LPCTSTR lpszBmpFileName)
{
	if (m_hbmImageWell == NULL)	// Not loaded yet!
	{
		return FALSE;
	}

	if (m_bReadOnly)
	{
		return FALSE;
	}

	CString strFile;
	if (lpszBmpFileName == NULL)
	{
		strFile = m_strUDLPath;
	}
	else
	{
		strFile = lpszBmpFileName;
	}

	if (!m_bModified && strFile == m_strUDLPath && m_nSVGCount == 0)
	{
		return TRUE;
	}
	
	HBITMAP hBmp = m_hbmImageWell;

#ifndef BCGP_EXCLUDE_SVG
	if (m_nSVGCount > 0)
	{
		hBmp = Copy(m_hbmImageWell);

		CWindowDC	dc(NULL);
		CDC 		memDCSrc;
		CDC 		memDCDst;
		
		memDCSrc.CreateCompatibleDC (&dc);
		memDCDst.CreateCompatibleDC (&dc);

		HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hBmp);

		for (int i = 0; i < m_iCount; i++)
		{
			if (i < m_arSVG.GetSize())
			{
				CBCGPSVGImage* pSVG = m_arSVG[i];
				if (pSVG != NULL)
				{
					HBITMAP hBmpSVG = pSVG->ExportToBitmap(m_sizeImage);
					if (hBmpSVG != NULL)
					{
						HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject(hBmpSVG);
						memDCDst.BitBlt (m_sizeImage.cx * i, 0, m_sizeImage.cx, m_sizeImage.cy, &memDCSrc, 0, 0, SRCCOPY);
						memDCSrc.SelectObject (hOldBitmapSrc);

						::DeleteObject(hBmpSVG);
					}
				}
			}
		}

		memDCDst.SelectObject (hOldBitmapDst);
	}
#endif

	CString ext;
	
#if _MSC_VER < 1400
	_tsplitpath(strFile, NULL, NULL, NULL, ext.GetBuffer (_MAX_EXT));
#else
	_tsplitpath_s(strFile, NULL, 0, NULL, 0, NULL, 0, ext.GetBuffer (_MAX_EXT), _MAX_EXT);
#endif
	
	ext.ReleaseBuffer ();
	
	if (ext.CompareNoCase (_T(".png")) == 0 && m_nBitsPerPixel >= 24)
	{
		CBCGPPngImage pngImage;
		pngImage.Attach(hBmp);

		BOOL bRes = pngImage.SaveToFile(strFile);

		pngImage.Detach();
		
		m_bModified = !bRes;

		if (m_nSVGCount > 0)
		{
			::DeleteObject(hBmp);
		}
		
		return bRes;
	}

	HANDLE hDib = DDBToDIB (hBmp, 0);
	if (hDib == NULL)
	{
		TRACE(_T("CBCGPToolBarImages::Save Can't convert DDB to DIB\n"));

		if (m_nSVGCount > 0)
		{
			::DeleteObject(hBmp);
		}
		
		return FALSE;
	}

	BOOL bSuccess = WriteDIB (strFile, hDib);
	::GlobalFree (hDib);

	if (m_nSVGCount > 0)
	{
		::DeleteObject(hBmp);
	}
	
	if (!bSuccess)
	{
		return FALSE;
	}

	m_bModified = FALSE;
	return TRUE;
}
//************************************************************************************
static BOOL WriteDIB( LPCTSTR szFile, HANDLE hDIB)
{
	BITMAPFILEHEADER	hdr;
	LPBITMAPINFOHEADER	lpbi;

	if (!hDIB)
		return FALSE;

	CFile file;
	if( !file.Open (szFile, CFile::modeWrite | CFile::modeCreate))
	{
		return FALSE;
	}

	lpbi = (LPBITMAPINFOHEADER) hDIB;

	int nColors = 1 << lpbi->biBitCount;
	if (nColors > 256 || lpbi->biBitCount == 32)
		nColors = 0;

	// Fill in the fields of the file header 
	hdr.bfType		= ((WORD) ('M' << 8) | 'B');	// is always "BM"
	hdr.bfSize		= (DWORD) (GlobalSize (hDIB) + sizeof (BITMAPFILEHEADER));
	hdr.bfReserved1 	= 0;
	hdr.bfReserved2 	= 0;
	hdr.bfOffBits		= (DWORD) (sizeof( hdr ) + lpbi->biSize +
						nColors * sizeof(RGBQUAD));

	// Write the file header 
	file.Write( &hdr, sizeof(hdr) );

	// Write the DIB header and the bits 
	file.Write( lpbi, (UINT) GlobalSize(hDIB) );

	return TRUE;
}
//********************************************************************************
static HANDLE DDBToDIB (HBITMAP bitmap, DWORD dwCompression) 
{
	BITMAP				bm;
	BITMAPINFOHEADER	bi;
	LPBITMAPINFOHEADER 	lpbi;
	DWORD				dwLen;
	HANDLE				hDIB;
	HANDLE				handle;
	HDC 				hDC;
	HPALETTE			hPal;

	// The function has no arg for bitfields
	if( dwCompression == BI_BITFIELDS)
		return NULL;

	hPal = (HPALETTE) GetStockObject(DEFAULT_PALETTE);

	// Get bitmap information
	::GetObject(bitmap, sizeof(bm),(LPSTR)&bm);

	// Initialize the bitmapinfoheader
	bi.biSize			= sizeof(BITMAPINFOHEADER);
	bi.biWidth			= bm.bmWidth;
	bi.biHeight 		= bm.bmHeight;
	bi.biPlanes 		= 1;
	bi.biBitCount		= (WORD) (bm.bmPlanes * bm.bmBitsPixel);
	bi.biCompression	= dwCompression;
	bi.biSizeImage		= 0;
	bi.biXPelsPerMeter	= 0;
	bi.biYPelsPerMeter	= 0;
	bi.biClrUsed		= 0;
	bi.biClrImportant	= 0;

	// Compute the size of the  infoheader and the color table
	int nColors = (1 << bi.biBitCount);
	if (nColors > 256 || bi.biBitCount == 32)
		nColors = 0;
	dwLen  = bi.biSize + nColors * sizeof(RGBQUAD);

	// We need a device context to get the DIB from
	hDC = ::CreateCompatibleDC (NULL);
	if (hDC == NULL)
	{
		return FALSE;
	}

	HBITMAP bmp = ::CreateBitmap (1, 1, 1, bi.biBitCount, NULL);
	if (bmp == NULL)
	{
		::DeleteDC(hDC);
		return NULL;
	}

	HBITMAP hOldBitmap = (HBITMAP)::SelectObject (hDC, bmp);

	hPal = SelectPalette(hDC,hPal,FALSE);
	RealizePalette(hDC);

	// Allocate enough memory to hold bitmapinfoheader and color table
	hDIB = GlobalAlloc(GMEM_FIXED,dwLen);

	if (!hDIB)
	{
		::SelectPalette (hDC,hPal,FALSE);

		if (hOldBitmap != NULL)
		{
			::SelectObject (hDC, hOldBitmap);
		}

		::DeleteObject (bmp);
		::DeleteDC(hDC);
		return NULL;
	}

	lpbi = (LPBITMAPINFOHEADER)hDIB;

	*lpbi = bi;

	// Call GetDIBits with a NULL lpBits param, so the device driver 
	// will calculate the biSizeImage field 
	GetDIBits(hDC, bitmap, 0L, (DWORD)bi.biHeight,
			(LPBYTE)NULL, (LPBITMAPINFO)lpbi, (DWORD)DIB_RGB_COLORS);

	bi = *lpbi;

	// If the driver did not fill in the biSizeImage field, then compute it
	// Each scan line of the image is aligned on a DWORD (32bit) boundary
	if (bi.biSizeImage == 0){
		bi.biSizeImage = ((((bi.biWidth * bi.biBitCount) + 31) & ~31) / 8) 
						* bi.biHeight;

		// If a compression scheme is used the result may infact be larger
		// Increase the size to account for this.
		if (dwCompression != BI_RGB)
			bi.biSizeImage = (bi.biSizeImage * 3) / 2;
	}

	// Realloc the buffer so that it can hold all the bits
	dwLen += bi.biSizeImage;
	handle = GlobalReAlloc(hDIB, dwLen, GMEM_MOVEABLE);
	if (handle != NULL)
		hDIB = handle;
	else{
		GlobalFree(hDIB);

		// Reselect the original palette
		SelectPalette(hDC,hPal,FALSE);
		if (hOldBitmap != NULL)
		{
			::SelectObject (hDC, hOldBitmap);
		}
		::DeleteObject (bmp);
		::DeleteDC(hDC);
		return NULL;
	}

	// Get the bitmap bits
	lpbi = (LPBITMAPINFOHEADER)hDIB;

	// FINALLY get the DIB
	BOOL bGotBits = GetDIBits( hDC, bitmap,
				0L,				// Start scan line
				(DWORD)bi.biHeight,		// # of scan lines
				(LPBYTE)lpbi 			// address for bitmap bits
				+ (bi.biSize + nColors * sizeof(RGBQUAD)),
				(LPBITMAPINFO)lpbi,		// address of bitmapinfo
				(DWORD)DIB_RGB_COLORS);		// Use RGB for color table

	if( !bGotBits )
	{
		GlobalFree(hDIB);
		
		SelectPalette(hDC,hPal,FALSE);
		if (hOldBitmap != NULL)
		{
			::SelectObject (hDC, hOldBitmap);
		}
		::DeleteObject (bmp);
		::DeleteDC(hDC);
		return NULL;
	}

	// Convert color table to the standard 3-d colors:
	DWORD* pColorTable = (DWORD*)(((LPBYTE)lpbi) + (UINT) lpbi->biSize);
	for (int iColor = 0; iColor < nColors; iColor ++)
	{
		pColorTable[iColor] = CBCGPToolBarImages::MapFromSysColor (pColorTable[iColor]);
	}

	SelectPalette(hDC,hPal,FALSE);

	if (hOldBitmap != NULL)
	{
		::SelectObject (hDC, hOldBitmap);
	}

	::DeleteObject (bmp);
	::DeleteDC(hDC);

	return hDIB;
}


/////////////////////////////////////////////////////////////////////////////
// CBCGPToolBarImages diagnostics

#ifdef _DEBUG
void CBCGPToolBarImages::AssertValid() const
{
	CObject::AssertValid();

	ASSERT(m_hbmImageWell != NULL);
}

void CBCGPToolBarImages::Dump(CDumpContext& dc) const
{
	CObject::Dump(dc);

	dc << "\nm_bUserImagesList = " << m_bUserImagesList;
	dc << "\nm_sizeImage = " << m_sizeImage;

	if (m_bUserImagesList)
	{
		dc << "\nm_strUDLPath = " << m_strUDLPath;
	}

	if (dc.GetDepth() > 0)
	{
	}

	dc << "\n";
}

#endif

BOOL CBCGPToolBarImages::CopyImageToClipboard (int iImage)
{
#ifndef _BCGPCHART_STANDALONE
	CBCGPLocalResource locaRes;
#endif

	try
	{
		CWindowDC dc (NULL);

		//----------------------
		// Create a bitmap copy:
		//----------------------
		CDC memDCDest;
		memDCDest.CreateCompatibleDC (NULL);
		
		CBitmap bitmapCopy;
		if (!bitmapCopy.CreateCompatibleBitmap (&dc, m_sizeImage.cx, m_sizeImage.cy))
		{
#ifndef _BCGPCHART_STANDALONE
			AfxMessageBox (IDP_BCGBARRES_CANT_COPY_BITMAP);
#endif
			return FALSE;
		}

		CBitmap* pOldBitmapDest = memDCDest.SelectObject (&bitmapCopy);

		memDCDest.FillRect (CRect (0, 0, m_sizeImage.cx, m_sizeImage.cy),
			&globalData.brBtnFace);

		CBCGPDrawState ds;
		PrepareDrawImage (ds, FALSE);
	
		Draw (&memDCDest, 0, 0, iImage);
		EndDrawImage (ds);

		memDCDest.SelectObject (pOldBitmapDest);

		if (!AfxGetMainWnd ()->OpenClipboard ())
		{
#ifndef _BCGPCHART_STANDALONE
			AfxMessageBox (IDP_BCGBARRES_CANT_COPY_BITMAP);
#endif
			return FALSE;
		}

		if (!::EmptyClipboard ())
		{
#ifndef _BCGPCHART_STANDALONE
			AfxMessageBox (IDP_BCGBARRES_CANT_COPY_BITMAP);
#endif
			::CloseClipboard ();
			return FALSE;
		}


		HANDLE hclipData = ::SetClipboardData (CF_BITMAP, bitmapCopy.Detach ());
		if (hclipData == NULL)
		{
#ifndef _BCGPCHART_STANDALONE
			AfxMessageBox (IDP_BCGBARRES_CANT_COPY_BITMAP);
#endif
			TRACE (_T("CBCGToolBar::OnBcgbarresCopyImage error. Error code = %x\n"), GetLastError ());
		}

		::CloseClipboard ();
		return TRUE;
	}
	catch (...)
	{
#ifndef _BCGPCHART_STANDALONE
		AfxMessageBox (IDP_BCGBARRES_INTERLAL_ERROR);
#endif
	}

	return FALSE;
}
//****************************************************************************************
BOOL CBCGPToolBarImages::CopyTo (CBCGPToolBarImages& dest)
{
	if (dest.m_bIsTemporary)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	if (dest.IsValid ())
	{
		dest.Clear ();
	}

	if (globalData.bIsWindowsVista)
	{
		BITMAP bmp;
		if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == sizeof (BITMAP))
		{
			CSize sizeImage (bmp.bmWidth, abs(bmp.bmHeight));

			//-------------------------------------------------------
			// Create memory source DC and select an original bitmap:
			//-------------------------------------------------------
			CDC memDCSrc;
			memDCSrc.CreateCompatibleDC (NULL);

			HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (m_hbmImageWell);
			if (hOldBitmapSrc != NULL)
			{

				//----------------------------------------------------------
				// Create a new bitmap compatible with the source memory DC
				// (original bitmap SHOULD BE ALREADY SELECTED!):
				//----------------------------------------------------------

				HBITMAP hNewBitmap = NULL;

				DIBSECTION ds = {0};
				if (bmp.bmBitsPixel >= 24 &&
					::GetObject (m_hbmImageWell, sizeof (DIBSECTION), &ds) != 0)
				{
					BITMAPINFO bi = {0};
					bi.bmiHeader.biSize        = sizeof (BITMAPINFOHEADER);
					bi.bmiHeader.biWidth       = bmp.bmWidth;
					bi.bmiHeader.biHeight      = bmp.bmHeight;
					bi.bmiHeader.biPlanes      = bmp.bmPlanes;
					bi.bmiHeader.biBitCount    = bmp.bmBitsPixel;
					bi.bmiHeader.biCompression = BI_RGB;

					COLORREF* pBits = NULL;
					hNewBitmap = ::CreateDIBSection (
						memDCSrc, &bi, DIB_RGB_COLORS, (void **)&pBits,
						NULL, NULL);
				}
				else
				{
					hNewBitmap = (HBITMAP) ::CreateCompatibleBitmap (memDCSrc,
												sizeImage.cx, sizeImage.cy);
				}

				if (hNewBitmap != NULL)
				{
					//------------------------------------------------------
					// Create memory destination DC and select a new bitmap:
					//------------------------------------------------------
					CDC memDCDst;
					memDCDst.CreateCompatibleDC (&memDCSrc);
					
					HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hNewBitmap);
					if (hOldBitmapDst != NULL)
					{
						//-----------------------------
						// Copy original bitmap to new:
						//-----------------------------
						memDCDst.BitBlt (0, 0, sizeImage.cx, sizeImage.cy,
									&memDCSrc, 0, 0, SRCCOPY);
							
						memDCDst.SelectObject (hOldBitmapDst);

						dest.m_hbmImageWell = hNewBitmap;
					}
					else
					{
						::DeleteObject (hNewBitmap);
					}
				}

				memDCSrc.SelectObject (hOldBitmapSrc);
			}
		}
	}
	else
	{
		dest.m_hbmImageWell = (HBITMAP) ::CopyImage (m_hbmImageWell, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
		if (m_hbmImageLight != NULL)
		{
			dest.m_hbmImageLight = (HBITMAP) ::CopyImage (m_hbmImageLight, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
		}
		if (m_hbmImageShadow != NULL)
		{
			dest.m_hbmImageShadow = (HBITMAP) ::CopyImage (m_hbmImageShadow, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
		}
	}

	dest.m_sizeImage = m_sizeImage;
	dest.m_sizeImageDest = m_sizeImageDest;
	dest.m_bUserImagesList = m_bUserImagesList;
	dest.m_strUDLPath = m_strUDLPath;
	dest.m_bModified = m_bModified;
	dest.m_iCount = m_iCount;
	dest.m_clrTransparent = m_clrTransparent;
	dest.m_bReadOnly = m_bReadOnly;
	dest.m_clrImageShadow = m_clrImageShadow;
	dest.m_bFadeInactive = m_bFadeInactive;
	dest.m_nBitsPerPixel = m_nBitsPerPixel;
	dest.m_dblScale = m_dblScale;
	dest.m_sizeImageOriginal = m_sizeImageOriginal;
	dest.m_bAutoGrayAfterReLoad = m_bAutoGrayAfterReLoad;
	dest.m_dblGrayLumRatio = m_dblGrayLumRatio;

	for (POSITION pos = m_lstOrigResIds.GetHeadPosition (); pos != NULL;)
	{
		UINT uiResId = m_lstOrigResIds.GetNext (pos);

		dest.m_lstOrigResIds.AddTail (uiResId);

		int iOffset = -1;
		if (m_mapOrigResOffsets.Lookup (uiResId, iOffset))
		{
			dest.m_mapOrigResOffsets.SetAt (uiResId, iOffset);
		}
	}

	for (POSITION posInst = m_lstOrigResInstances.GetHeadPosition (); posInst != NULL;)
	{
		HINSTANCE hInst = m_lstOrigResInstances.GetNext (posInst);
		dest.m_lstOrigResInstances.AddTail (hInst);
	}

#ifndef BCGP_EXCLUDE_SVG
	dest.CleanUpSVG();
	dest.m_nSVGCount = m_nSVGCount;

	for (int i = 0; i < (int)m_arSVG.GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = NULL;

		if (m_arSVG[i] != NULL)
		{
			pSVG = new CBCGPSVGImage(*m_arSVG[i]);
		}

		dest.m_arSVG.Add(pSVG);
	}
#endif
	
	return TRUE;
}
//****************************************************************************************
void CBCGPToolBarImages::Clear ()
{
	if (m_bIsTemporary)
	{
		ASSERT (FALSE);
		return;
	}

	AfxDeleteObject((HGDIOBJ*)&m_hbmImageWell);
	m_hbmImageWell = NULL;
	CleanUpInternalImages();

	m_lstOrigResIds.RemoveAll();
	m_mapOrigResOffsets.RemoveAll ();
	m_lstOrigResInstances.RemoveAll ();
	m_strUDLPath.Empty();

	m_bUserImagesList = FALSE;
	m_iCount = 0;
	m_bModified = FALSE;
	m_bIsGray = FALSE;
	m_nGrayPercentage = 0;
	m_nBitsPerPixel = 0;

#ifndef BCGP_EXCLUDE_SVG
	CleanUpSVG();
#endif

	if (m_dblScale != 1.0)
	{
		m_sizeImage = m_sizeImageOriginal;
		m_sizeImageOriginal = CSize (0, 0);
		m_dblScale = 1.0;
	}
}
//*************************************************************************************
void CBCGPToolBarImages::TransparentBlt( HDC hdcDest, int nXDest, int nYDest, int nWidth, 
			int nHeight, CDC* pDcSrc, int nXSrc, int nYSrc,
			COLORREF colorTransparent,
			int nWidthDest/* = -1*/, int nHeightDest/* = -1*/)
{
	int cx = nWidthDest == -1 ? nWidth : nWidthDest;
	int cy = nHeightDest == -1 ? nHeight : nHeightDest;

	if (m_pfTransparentBlt != NULL && !m_bIsRTL && !globalData.bIsWindows9x)
	{
		if ((*m_pfTransparentBlt) (hdcDest, nXDest, nYDest, cx, cy, 
			pDcSrc->GetSafeHdc (), nXSrc, nYSrc,
			nWidth, nHeight, colorTransparent))
		{
			return;
		}
	}

	CDC dc, memDC, maskDC;
	dc.Attach( hdcDest );
	maskDC.CreateCompatibleDC(&dc);
	CBitmap maskBitmap;

	//add these to store return of SelectObject() calls
	CBitmap* pOldMemBmp = NULL;
	CBitmap* pOldMaskBmp = NULL;
	
	memDC.CreateCompatibleDC(&dc);
	CBitmap bmpImage;
	bmpImage.CreateCompatibleBitmap( &dc, cx, cy);
	pOldMemBmp = memDC.SelectObject( &bmpImage );

	if (nWidthDest == -1 || (nWidthDest == nWidth && nHeightDest == nHeight))
	{
		memDC.BitBlt( 0,0,nWidth, nHeight, pDcSrc, nXSrc, nYSrc, SRCCOPY);
	}
	else
	{
		memDC.StretchBlt (0,0, nWidthDest, nHeightDest, pDcSrc, 
						nXSrc, nYSrc, nWidth, nHeight, SRCCOPY);
	}
	
	// Create monochrome bitmap for the mask
	maskBitmap.CreateBitmap (cx, cy, 1, 1, NULL );
	pOldMaskBmp = maskDC.SelectObject( &maskBitmap );
	memDC.SetBkColor( colorTransparent );
	
	// Create the mask from the memory DC
	maskDC.BitBlt (0, 0, cx, cy, &memDC, 0, 0, SRCCOPY);
	
	// Set the background in memDC to black. Using SRCPAINT with black 
	// and any other color results in the other color, thus making 
	// black the transparent color
	memDC.SetBkColor(RGB(0,0,0));
	memDC.SetTextColor(RGB(255,255,255));
	memDC.BitBlt(0, 0, cx, cy, &maskDC, 0, 0, SRCAND);
	
	// Set the foreground to black. See comment above.
	dc.SetBkColor(RGB(255,255,255));
	dc.SetTextColor(RGB(0,0,0));

	dc.BitBlt (nXDest, nYDest, cx, cy, &maskDC, 0, 0, SRCAND);
		
	// Combine the foreground with the background
	dc.BitBlt(nXDest, nYDest, cx, cy, &memDC, 
		0, 0, SRCPAINT);
	
	if (pOldMaskBmp)
		maskDC.SelectObject( pOldMaskBmp );
	if (pOldMemBmp)
		memDC.SelectObject( pOldMemBmp );
	
	dc.Detach();
}
//***********************************************************************************
BOOL CBCGPToolBarImages::MapBmpTo3dColors (	HBITMAP& hBmp, 
											BOOL bUseRGBQUAD/* = TRUE*/,
											COLORREF clrSrc/* = (COLORREF)-1*/,
											COLORREF clrDest/* = (COLORREF)-1*/)
{
	if (hBmp == NULL)
	{
		return FALSE;
	}

	if (clrSrc != (COLORREF)-1 && clrDest == (COLORREF)-1)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	//-------------------------------
	// Get original bitmap attrbutes:
	//-------------------------------
	BITMAP bmp;
	if (::GetObject (hBmp, sizeof (BITMAP), &bmp) == 0)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	if (bmp.bmBitsPixel == 24)
	{
		DIBSECTION ds;
		if (::GetObject (hBmp, sizeof (DIBSECTION), &ds) == 0)
		{
			ASSERT (FALSE);
			return FALSE;
		}
		
		if (ds.dsBm.bmBits != NULL)
		{
			if (ds.dsBm.bmBitsPixel < 24)
			{
				ASSERT (FALSE);
				return FALSE;
			}
			
			DWORD dwPitch = ((ds.dsBm.bmWidth * 24 + 31) & ~31) / 8;
			LPBYTE pBits = (LPBYTE)ds.dsBm.bmBits;
			for (int y = 0; y < ds.dsBm.bmHeight; y++)
			{
				LPBYTE pBitsRow = pBits;
				for (int x = 0; x < ds.dsBm.bmWidth; x++)
				{
					COLORREF clrOrig = RGB (pBitsRow[2], pBitsRow[1], pBitsRow[0]);
					COLORREF clrNew;
					
					if (clrSrc != (COLORREF)-1 && clrDest != (COLORREF)-1 && clrOrig == clrSrc)
					{
						clrNew = clrDest;
					}
					else
					{
						clrNew = !m_bDisableTrueColorAlpha ? MapToSysColorAlpha (clrOrig) : MapToSysColor (clrOrig, bUseRGBQUAD);
					}

					if (clrOrig != clrNew)
					{
						pBitsRow[2] = (BYTE) (GetRValue (clrNew));
						pBitsRow[1] = (BYTE) (GetGValue (clrNew));
						pBitsRow[0] = (BYTE) (GetBValue (clrNew));
					}
					
					pBitsRow += 3;
				}
				
				pBits += dwPitch;
			}

			return TRUE;
		}
	}

	//-------------------------------------------------------
	// Create memory source DC and select an original bitmap:
	//-------------------------------------------------------
	CDC memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	HBITMAP hOldBitmapSrc = NULL;

	int iBitmapWidth;
	int iBitmapHeight;

	hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (hBmp);
	if (hOldBitmapSrc == NULL)
	{
		return FALSE;
	}

	iBitmapWidth = bmp.bmWidth;
	iBitmapHeight = bmp.bmHeight;

	//----------------------------------------------------------
	// Create a new bitmap compatibel with the source memory DC:
	// (original bitmap SHOULD BE ALREADY SELECTED!):
	//----------------------------------------------------------
	HBITMAP hNewBitmap = (HBITMAP) ::CreateCompatibleBitmap (memDCSrc,
									iBitmapWidth,
									iBitmapHeight);
	if (hNewBitmap == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		return FALSE;
	}

	//------------------------------
	// Create memory destination DC:
	//------------------------------
	CDC memDCDst;
	memDCDst.CreateCompatibleDC (&memDCSrc);

	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hNewBitmap);
	if (hOldBitmapDst == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		::DeleteObject (hNewBitmap);
		return FALSE;
	}

	//-----------------------------
	// Copy original bitmap to new:
	//-----------------------------
	memDCDst.BitBlt (0, 0, iBitmapWidth, iBitmapHeight,
				&memDCSrc, 0, 0, SRCCOPY);
		
	//------------------------------------
	// Change a specific colors to system:
	//------------------------------------
	for (int x = 0; x < iBitmapWidth; x ++)
	{
		for (int y = 0; y < iBitmapHeight; y ++)
		{
			COLORREF clrOrig = ::GetPixel (memDCDst, x, y);

			if (clrSrc != (COLORREF)-1)
			{
				if (clrOrig == clrSrc)
				{
					::SetPixel (memDCDst, x, y, clrDest);
				}
			}
			else
			{
				COLORREF clrNew = bmp.bmBitsPixel == 24 && !m_bDisableTrueColorAlpha ?
					MapToSysColorAlpha (clrOrig) :
					MapToSysColor (clrOrig, bUseRGBQUAD);

				if (clrOrig != clrNew)
				{
					::SetPixel (memDCDst, x, y, clrNew);
				}
			}
		}
	}
	
	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);

	::DeleteObject (hBmp);
	hBmp = hNewBitmap;

	return TRUE;
}
//**********************************************************************************
BOOL CBCGPToolBarImages::MapTo3dColors (BOOL bUseRGBQUAD/* = TRUE*/,
										COLORREF clrSrc/* = (COLORREF)-1*/,
										COLORREF clrDest/* = (COLORREF)-1*/)
{
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return TRUE;
	}

	return MapBmpTo3dColors (m_hbmImageWell, bUseRGBQUAD, clrSrc, clrDest);
}
//*******************************************************************************
void CBCGPToolBarImages::CopyTemp (CBCGPToolBarImages& imagesDest)
{
	imagesDest.Clear ();
	imagesDest.m_bIsTemporary = TRUE;

	imagesDest.m_sizeImage = m_sizeImage;
	imagesDest.m_sizeImageDest = m_sizeImageDest;
	imagesDest.m_hbmImageWell = m_hbmImageWell;
	imagesDest.m_bUserImagesList = m_bUserImagesList;
	imagesDest.m_iCount = m_iCount;
	imagesDest.m_bReadOnly = TRUE;
	imagesDest.m_nBitsPerPixel = m_nBitsPerPixel;

#ifndef BCGP_EXCLUDE_SVG
	if (m_nSVGCount > 0)
	{
		imagesDest.m_nSVGCount = m_nSVGCount;
		imagesDest.m_arSVG.Append(m_arSVG);
	}
#endif
}
//*******************************************************************************
BOOL CBCGPToolBarImages::UpdateInternalImage (int nIndex)
{
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return TRUE;
	}

#ifdef _BCGPCHART_STANDALONE
	UNREFERENCED_PARAMETER(nIndex);
#else
	switch (nIndex)
	{
	case BCGImage_Shadow:
		if (!CBCGPVisualManager::GetInstance ()->IsShadowHighlightedImage ())
		{
			return FALSE;
		}
		break;

	case BCGImage_Light:
		if (!CBCGPVisualManager::GetInstance ()->IsFadeInactiveImage())
		{
			return FALSE;
		}
		break;
	}

	if (nIndex == BCGImage_Gray)
	{
		AfxDeleteObject((HGDIOBJ*)&m_hbmImageGray);
		m_hbmImageGray = NULL;

		if (m_hbmImageWell != NULL)
		{
			m_hbmImageGray = Copy(m_hbmImageWell);
			_ConvertToGrayScale(m_hbmImageGray, m_clrTransparent, CBCGPVisualManager::GetInstance()->GetGrayScaleLumRatio());
		}

		return TRUE;
	}

	HBITMAP& hbmpInternal = (nIndex == BCGImage_Light) ? 
		m_hbmImageLight : m_hbmImageShadow;

	if (nIndex == BCGImage_Light)
	{
		if ((m_nBitsPerPixel > 4 && !m_bAlwaysLight) || m_nBitsPerPixel == 0)
		{
			// Down't fade 256+ or unknown bitmaps
			return FALSE;
		}
	}
	
	AfxDeleteObject((HGDIOBJ*)&hbmpInternal);
	hbmpInternal = NULL;

	if (m_hbmImageWell == NULL)
	{
		return TRUE;
	}

	if (globalData.m_nBitsPerPixel <= 8 || !globalData.bIsWindows2000)
	{
		return TRUE;
	}

	//-------------------------------------------------------
	// Create memory source DC and select an original bitmap:
	//-------------------------------------------------------
	CDC memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	BITMAP bmp;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
	{
		return FALSE;
	}

	int iBitmapWidth = bmp.bmWidth;
	int iBitmapHeight = bmp.bmHeight;

	HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (m_hbmImageWell);
	if (hOldBitmapSrc == NULL)
	{
		return FALSE;
	}

	//------------------------------------------------------
	// Create memory destination DC and select a new bitmap:
	//------------------------------------------------------
	CDC memDCDst;
	memDCDst.CreateCompatibleDC (&memDCSrc);
	
	BITMAPINFO bi;

	// Fill in the BITMAPINFOHEADER
	bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bi.bmiHeader.biWidth = iBitmapWidth;
	bi.bmiHeader.biHeight = iBitmapHeight;
	bi.bmiHeader.biPlanes = 1;
	bi.bmiHeader.biBitCount = 32;
	bi.bmiHeader.biCompression = BI_RGB;
	bi.bmiHeader.biSizeImage = iBitmapWidth * iBitmapHeight;
	bi.bmiHeader.biXPelsPerMeter = 0;
	bi.bmiHeader.biYPelsPerMeter = 0;
	bi.bmiHeader.biClrUsed = 0;
	bi.bmiHeader.biClrImportant = 0;

	COLORREF* pBits = NULL;
	hbmpInternal = CreateDIBSection (
		memDCDst.m_hDC, &bi, DIB_RGB_COLORS, (void **)&pBits,
		NULL, NULL);

	if (hbmpInternal == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		return FALSE;
	}

	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hbmpInternal);
	if (hOldBitmapDst == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		::DeleteObject (hbmpInternal);
		hbmpInternal = NULL;
		return FALSE;
	}

	//-----------------------------
	// Copy original bitmap to new:
	//-----------------------------
	COLORREF clrTransparent = m_nBitsPerPixel == 32 && m_pfAlphaBlend != NULL ?
		(COLORREF) -1 : m_clrTransparent;

	memDCDst.BitBlt (0, 0, iBitmapWidth, iBitmapHeight,
				&memDCSrc, 0, 0, SRCCOPY);
		
	if (nIndex == BCGImage_Light)
	{
		CBCGPDrawManager dm (memDCDst);

		dm.HighlightRect (CRect (0, 0, iBitmapWidth, iBitmapHeight),
			m_nLightPercentage,
			clrTransparent == -1 ? globalData.clrBtnFace : clrTransparent);
	}
	else
	{
		COLORREF clrShadow = globalData.m_nBitsPerPixel <= 8
								? globalData.clrBtnShadow
								: CBCGPDrawManager::PixelAlpha (CBCGPVisualManager::GetInstance ()->GetToolbarHighlightColor (), 67);

		if (m_nBitsPerPixel == 32)
		{
			BYTE clrShadowR = GetRValue(clrShadow);
			BYTE clrShadowG = GetGValue(clrShadow);
			BYTE clrShadowB = GetBValue(clrShadow);

			CBCGPScanlinerBitmap scan;
			if (scan.Attach(hbmpInternal))
			{
				for (LONG y = 0; y < iBitmapHeight; y++)
				{
					LPBYTE pRow = scan.Get();

					for (LONG x = 0; x < iBitmapWidth; x++)
					{
						double a = (double)pRow[3];
						if (a != 0)
						{
							pRow[0] = clrShadowR;
							pRow[1] = clrShadowG;
							pRow[2] = clrShadowB;

							if (a < 255.0)
							{
								a /= 255.0;

								pRow[0] = (BYTE)bcg_clamp_to_byte(pRow[0] * a);
								pRow[1] = (BYTE)bcg_clamp_to_byte(pRow[1] * a);
								pRow[2] = (BYTE)bcg_clamp_to_byte(pRow[2] * a);
							}
						}

						pRow += 4;
					}

					scan++;
				}
			}
		}
		else
		{
			COLORREF clrTr = 
				clrTransparent == -1 ? globalData.clrBtnFace : clrTransparent;

			for (int x = 0; x < iBitmapWidth; x++)
			{
				for (int y = 0; y < iBitmapHeight; y++)
				{
					COLORREF clr = memDCDst.GetPixel (x, y);
					if (clr != clrTr)
					{
						memDCDst.SetPixel (x, y, clrShadow);
					}
				}
			}
		}
	}

	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);
#endif
	return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::PreMultiplyAlpha (HBITMAP hbmp, BOOL bAutoCheckPremlt)
{
	DIBSECTION ds = {0};
	if (::GetObject (hbmp, sizeof (DIBSECTION), &ds) != sizeof (DIBSECTION))
	{
		ASSERT (FALSE);
		return FALSE;
	}

	if (ds.dsBm.bmBits == NULL || ds.dsBm.bmBitsPixel != 32)
	{
		return FALSE;
	}

	int i = 0;

	RGBQUAD* pBits = (RGBQUAD*) ds.dsBm.bmBits;
	const int length = ds.dsBm.bmWidth * ds.dsBm.bmHeight;

	if (bAutoCheckPremlt)
	{
		BOOL bPremultiply = FALSE;
    
		RGBQUAD* pBit = pBits;
		for (i = 0; i < length; i++)
		{
			if (pBit->rgbRed   > pBit->rgbReserved ||
			    pBit->rgbGreen > pBit->rgbReserved ||
			    pBit->rgbBlue  > pBit->rgbReserved)
			{
				bPremultiply = TRUE;
				break;
			}

			pBit++;
		}

		if (!bPremultiply)
		{
			return TRUE;
		}
	}

	//----------------------------------------------------------------
	// Premultiply the R,G and B values with the Alpha channel values:
	//----------------------------------------------------------------
	RGBQUAD* pBit = pBits;
	for (i = 0; i < length; i++)
	{
		double alpha = (double)(pBit->rgbReserved) / 255.0;

		pBit->rgbRed   = (BYTE)min(bcg_round(pBit->rgbRed   * alpha), pBit->rgbReserved);
		pBit->rgbGreen = (BYTE)min(bcg_round(pBit->rgbGreen * alpha), pBit->rgbReserved);
		pBit->rgbBlue  = (BYTE)min(bcg_round(pBit->rgbBlue  * alpha), pBit->rgbReserved);
		pBit++;
	}

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::MultiplyAlpha (HBITMAP hbmp)
{
	DIBSECTION ds = {0};
	if (::GetObject (hbmp, sizeof (DIBSECTION), &ds) != sizeof (DIBSECTION))
	{
		ASSERT (FALSE);
		return FALSE;
	}

	if (ds.dsBm.bmBits == NULL || ds.dsBm.bmBitsPixel != 32)
	{
		return FALSE;
	}

	RGBQUAD* pBits = (RGBQUAD*)ds.dsBm.bmBits;
	const int length = ds.dsBm.bmWidth * ds.dsBm.bmHeight;

	//----------------------------------------------------------------
	// Multiply the R,G and B values with the Alpha channel values:
	//----------------------------------------------------------------
	RGBQUAD* pBit = pBits;
	for (int i = 0; i < length; i++)
	{
		if (pBit->rgbReserved != 0 && pBit->rgbReserved != 255)
		{
			double alpha = 255.0 / (double)pBit->rgbReserved;

			pBit->rgbRed   = (BYTE)min(bcg_round(pBit->rgbRed   * alpha), 255.0);
			pBit->rgbGreen = (BYTE)min(bcg_round(pBit->rgbGreen * alpha), 255.0);
			pBit->rgbBlue  = (BYTE)min(bcg_round(pBit->rgbBlue  * alpha), 255.0);
		}

		pBit++;
	}

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::PreMultiplyAlpha (HBITMAP hbmp)
{
	return PreMultiplyAlpha (hbmp, m_bAutoCheckPremlt);
}
//*******************************************************************************
BOOL CBCGPToolBarImages::CreateFromImageList (const CImageList& imageList)
{
	ASSERT (imageList.GetSafeHandle () != NULL);

	Clear ();

	if (imageList.GetImageCount () == 0)
	{
		return FALSE;
	}

	IMAGEINFO info;
	imageList.GetImageInfo (0, &info);

	CRect rectImage = info.rcImage;
	m_sizeImage = rectImage.Size ();

	BITMAP bmpObj;
	::GetObject(info.hbmImage, sizeof (BITMAP), &bmpObj);

	m_nBitsPerPixel = bmpObj.bmBitsPixel;

	for (int i = 0; i < imageList.GetImageCount (); i++)
	{
		HICON hIcon = ((CImageList&) imageList).ExtractIcon (i);
		ASSERT (hIcon != NULL);

		AddIcon (hIcon, m_nBitsPerPixel == 32);

		::DestroyIcon (hIcon);
	}

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::ExportToImageList (CImageList& imageList)
{
	if (imageList.GetSafeHandle() != NULL)
	{
		imageList.DeleteImageList();
	}

	BITMAP bmp;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
	{
		return FALSE;
	}

	UINT nFlags = (m_nBitsPerPixel == 32) ? 0 : ILC_MASK;

	switch (bmp.bmBitsPixel)
	{
	case 4:
	default:
		nFlags |= ILC_COLOR4;
		break;

	case 8:
		nFlags |= ILC_COLOR8;
		break;

	case 16:
		nFlags |= ILC_COLOR16;
		break;

	case 24:
		nFlags |= ILC_COLOR24;
		break;

	case 32:
		nFlags |= ILC_COLOR32;
		break;
	}

	if (!imageList.Create(m_sizeImage.cx, m_sizeImage.cy, nFlags, 0, 0))
	{
		return FALSE;
	}

	HBITMAP hbmImageWellCopy = Copy(m_hbmImageWell);

	if (hbmImageWellCopy != NULL)
	{
		if (m_nSVGCount > 0)
		{
			CWindowDC	dc (NULL);
			CDC 		memDCSrc;
			CDC 		memDCDst;
			
			memDCSrc.CreateCompatibleDC (&dc);
			memDCDst.CreateCompatibleDC (&dc);
			
			HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hbmImageWellCopy);

#ifndef BCGP_EXCLUDE_SVG
			for (int i = 0; i < m_arSVG.GetSize(); i++)
			{
				CBCGPSVGImage* pSVG = m_arSVG[i];
				if (pSVG != NULL)
				{
					HBITMAP hbmpSVG = pSVG->ExportToBitmap(m_sizeImage);
					if (hbmpSVG != NULL)
					{
						HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject(hbmpSVG);
						
						memDCDst.BitBlt(m_sizeImage.cx * i, 0, m_sizeImage.cx, m_sizeImage.cy, &memDCSrc, 0, 0, SRCCOPY);
						
						memDCSrc.SelectObject (hOldBitmapSrc);

						::DeleteObject(hbmpSVG);
					}
				}
			}
#endif

			memDCDst.SelectObject (hOldBitmapDst);
		}

		if ((nFlags & ILC_COLOR32) == ILC_COLOR32)
		{
			MultiplyAlpha(hbmImageWellCopy);
			imageList.Add(CBitmap::FromHandle(hbmImageWellCopy), (CBitmap*)NULL);
		}
		else
		{
			imageList.Add(CBitmap::FromHandle (hbmImageWellCopy), 
				m_clrTransparent == -1 ? globalData.clrBtnFace : m_clrTransparent);
		}

		AfxDeleteObject((HGDIOBJ*)&hbmImageWellCopy);
	}
	else
	{
		imageList.Add(CBitmap::FromHandle (m_hbmImageWell), 
			m_clrTransparent == -1 ? globalData.clrBtnFace : m_clrTransparent);
	}

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::Is32BitTransparencySupported ()
{
	return globalData.bIsOSAlphaBlendingSupport;
}
//********************************************************************************
BOOL CBCGPToolBarImages::GrayImages (int nGrayPercentage)
{
	m_bIsGray = TRUE;
	m_nGrayPercentage = nGrayPercentage;

	if (m_hbmImageWell == NULL)
	{
		return TRUE;
	}

	if (globalData.m_nBitsPerPixel <= 8 || !globalData.bIsWindows2000)
	{
		return TRUE;
	}

	//-------------------------------------------------------
	// Create memory source DC and select an original bitmap:
	//-------------------------------------------------------
	CDC memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	BITMAP bmp;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
	{
		return FALSE;
	}

	int iBitmapWidth = bmp.bmWidth;
	int iBitmapHeight = bmp.bmHeight;

	HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (m_hbmImageWell);
	if (hOldBitmapSrc == NULL)
	{
		return FALSE;
	}

	//------------------------------------------------------
	// Create memory destination DC and select a new bitmap:
	//------------------------------------------------------
	CDC memDCDst;
	memDCDst.CreateCompatibleDC (&memDCSrc);
	
	BITMAPINFO bi;

	// Fill in the BITMAPINFOHEADER
	bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bi.bmiHeader.biWidth = iBitmapWidth;
	bi.bmiHeader.biHeight = iBitmapHeight;
	bi.bmiHeader.biPlanes = 1;
	bi.bmiHeader.biBitCount = 32;
	bi.bmiHeader.biCompression = BI_RGB;
	bi.bmiHeader.biSizeImage = iBitmapWidth * iBitmapHeight;
	bi.bmiHeader.biXPelsPerMeter = 0;
	bi.bmiHeader.biYPelsPerMeter = 0;
	bi.bmiHeader.biClrUsed = 0;
	bi.bmiHeader.biClrImportant = 0;

	COLORREF* pBits = NULL;
	HBITMAP hNewBitmap = CreateDIBSection (
		memDCDst.m_hDC, &bi, DIB_RGB_COLORS, (void **)&pBits,
		NULL, NULL);

	if (hNewBitmap == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		return FALSE;
	}

	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hNewBitmap);
	if (hOldBitmapDst == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		::DeleteObject (hNewBitmap);
		hNewBitmap = NULL;
		return FALSE;
	}

	//-----------------------------
	// Copy original bitmap to new:
	//-----------------------------
	memDCDst.BitBlt (0, 0, iBitmapWidth, iBitmapHeight,
				&memDCSrc, 0, 0, SRCCOPY);

	int nPercentage = m_nGrayPercentage <= 0 ? 130 : m_nGrayPercentage;

	if (m_nBitsPerPixel == 32 && m_pfAlphaBlend != NULL)
	{
		DIBSECTION ds;
		if (::GetObject (hNewBitmap, sizeof (DIBSECTION), &ds) == 0)
		{
			ASSERT (FALSE);
			return FALSE;
		}

		if (ds.dsBm.bmBitsPixel != 32)
		{
			ASSERT (FALSE);
			return FALSE;
		}

		RGBQUAD* pBits32 = (RGBQUAD*) ds.dsBm.bmBits;

		//----------------------------------------------------------------
		// Premultiply the R,G and B values with the Alpha channel values:
		//----------------------------------------------------------------
		for (int i = 0; i < ds.dsBm.bmWidth * ds.dsBm.bmHeight; i++)
		{
			RGBQUAD* pBit = pBits32 + i;

			double H,S,L;
			CBCGPDrawManager::RGBtoHSL (RGB (pBit->rgbRed, pBit->rgbGreen, pBit->rgbBlue), &H, &S, &L);
			COLORREF color = CBCGPDrawManager::PixelAlpha (
				CBCGPDrawManager::HLStoRGB_ONE (H,L,0),
				.01 * nPercentage, .01 * nPercentage, .01 * nPercentage);

			pBit->rgbRed = (BYTE) (GetRValue (color) * pBit->rgbReserved / 255);
			pBit->rgbGreen = (BYTE) (GetGValue (color) * pBit->rgbReserved / 255);
			pBit->rgbBlue = (BYTE) (GetBValue (color) * pBit->rgbReserved / 255);
		}
	}
	else
	{
		CBCGPDrawManager dm (memDCDst);

		dm.GrayRect (CRect (0, 0, iBitmapWidth, iBitmapHeight),
			nPercentage,
			m_clrTransparent == -1 ? globalData.clrBtnFace : m_clrTransparent);
	}

	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);

	::DeleteObject (m_hbmImageWell);
	m_hbmImageWell = hNewBitmap;

	return TRUE;
}
//********************************************************************************
BOOL CBCGPToolBarImages::Simplify()
{
	if (globalData.m_nBitsPerPixel <= 8 || !globalData.bIsWindows2000 || m_nBitsPerPixel < 32)
	{
		return TRUE;
	}

#ifndef BCGP_EXCLUDE_SVG
	for (int i = 0; i < (int)m_arSVG.GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = m_arSVG[i];
		if (pSVG != NULL)
		{
			pSVG->Simplify();
		}
	}
	
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return TRUE;
	}
#endif

	//-------------------------------------------------------
	// Create memory source DC and select an original bitmap:
	//-------------------------------------------------------
	CDC memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	BITMAP bmp;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
	{
		return FALSE;
	}

	int iBitmapWidth = bmp.bmWidth;
	int iBitmapHeight = bmp.bmHeight;

	HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (m_hbmImageWell);
	if (hOldBitmapSrc == NULL)
	{
		return FALSE;
	}

	//------------------------------------------------------
	// Create memory destination DC and select a new bitmap:
	//------------------------------------------------------
	CDC memDCDst;
	memDCDst.CreateCompatibleDC (&memDCSrc);
	
	BITMAPINFO bi;

	// Fill in the BITMAPINFOHEADER
	bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bi.bmiHeader.biWidth = iBitmapWidth;
	bi.bmiHeader.biHeight = iBitmapHeight;
	bi.bmiHeader.biPlanes = 1;
	bi.bmiHeader.biBitCount = 32;
	bi.bmiHeader.biCompression = BI_RGB;
	bi.bmiHeader.biSizeImage = iBitmapWidth * iBitmapHeight;
	bi.bmiHeader.biXPelsPerMeter = 0;
	bi.bmiHeader.biYPelsPerMeter = 0;
	bi.bmiHeader.biClrUsed = 0;
	bi.bmiHeader.biClrImportant = 0;

	COLORREF* pBits = NULL;
	HBITMAP hNewBitmap = CreateDIBSection (
		memDCDst.m_hDC, &bi, DIB_RGB_COLORS, (void **)&pBits,
		NULL, NULL);

	if (hNewBitmap == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		return FALSE;
	}

	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hNewBitmap);
	if (hOldBitmapDst == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		::DeleteObject (hNewBitmap);
		hNewBitmap = NULL;
		return FALSE;
	}

	//-----------------------------
	// Copy original bitmap to new:
	//-----------------------------
	memDCDst.BitBlt (0, 0, iBitmapWidth, iBitmapHeight, &memDCSrc, 0, 0, SRCCOPY);

	if (m_nBitsPerPixel == 32 && m_pfAlphaBlend != NULL)
	{
		DIBSECTION ds;
		if (::GetObject (hNewBitmap, sizeof (DIBSECTION), &ds) == 0)
		{
			ASSERT (FALSE);
			return FALSE;
		}

		if (ds.dsBm.bmBitsPixel != 32)
		{
			ASSERT (FALSE);
			return FALSE;
		}

		RGBQUAD* pBits32 = (RGBQUAD*) ds.dsBm.bmBits;

		for (int i = 0; i < ds.dsBm.bmWidth * ds.dsBm.bmHeight; i++)
		{
			RGBQUAD* pBit = pBits32 + i;

			COLORREF clrCurr = RGB (pBit->rgbRed, pBit->rgbGreen, pBit->rgbBlue);

			double H,S,L;
			CBCGPDrawManager::RGBtoHSL (clrCurr, &H, &S, &L);

			if (L >= 0.87)
			{
				pBit->rgbReserved = 0;
				pBit->rgbRed   = 0;
				pBit->rgbGreen = 0;
				pBit->rgbBlue  = 0;
			}
			else
			{
				pBit->rgbRed = pBit->rgbReserved;
				pBit->rgbGreen = pBit->rgbReserved;
				pBit->rgbBlue = pBit->rgbReserved;
			}
		}
	}

	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);

	::DeleteObject (m_hbmImageWell);
	m_hbmImageWell = hNewBitmap;

	return TRUE;
}
//*******************************************************************************
HICON CBCGPToolBarImages::SimplifyIcon(HICON hIcon, BOOL bDefaultSize)
{
	const CSize sizeIcon = bDefaultSize ? globalUtils.ScaleByDPI(CSize(16, 16)) : globalUtils.GetIconSize(hIcon);

	CBCGPToolBarImages images;
	
	images.SetImageSize(sizeIcon);
	images.AddIcon(hIcon, TRUE);
	
	images.Simplify();

	return images.ExtractIcon(0);
}
//*******************************************************************************
BOOL CBCGPToolBarImages::MirrorBitmap (HBITMAP& hbmp, int cxImage)
{
	if (hbmp == NULL || cxImage <= 0)
	{
		return TRUE;
	}

	BITMAP bmp;
	if (::GetObject (hbmp, sizeof (BITMAP), &bmp) == 0)
	{
		return FALSE;
	}

	const int cx = bmp.bmWidth;
	const int cy = bmp.bmHeight;
	const int iCount = cx / cxImage;

	if (bmp.bmBitsPixel == 32)
	{
		DIBSECTION ds;
		if (::GetObject (hbmp, sizeof (DIBSECTION), &ds) == 0)
		{
			ASSERT (FALSE);
			return FALSE;
		}

		if (ds.dsBm.bmBitsPixel != 32)
		{
			ASSERT (FALSE);
			return FALSE;
		}

		DWORD* pBits = (DWORD*) ds.dsBm.bmBits;

		for (int iImage = 0; iImage < iCount; iImage++)
		{
			for (int y = 0; y < cy; y++)
			{
				DWORD* pRow1 = pBits + cx * y + iImage * cxImage;
                DWORD* pRow2 = pRow1 + cxImage - 1;

				for (int x = 0; x < cxImage / 2; x++)
				{
					DWORD color = *pRow1;

					*pRow1 = *pRow2;
					*pRow2 = color;

                    pRow1++;
                    pRow2--;
				}
			}
		}

		return TRUE;
	}

	CDC memDC;
	memDC.CreateCompatibleDC (NULL);

	HBITMAP hOldBitmap = (HBITMAP) memDC.SelectObject (hbmp);
	if (hOldBitmap == NULL)
	{
		return FALSE;
	}

	for (int iImage = 0; iImage < iCount; iImage++)
	{
		for (int y = 0; y < cy; y++)
		{
			int x1 = iImage * cxImage;
			int x2 = x1 + cxImage - 1;

			for (int x = 0; x < cxImage / 2; x++)
			{
				COLORREF color = memDC.GetPixel (x1, y);

				memDC.SetPixel (x1, y, memDC.GetPixel (x2, y));
				memDC.SetPixel (x2, y, color);

				x1++;
				x2--;
			}
		}
	}

	memDC.SelectObject (hOldBitmap);

    return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::Mirror ()
{
#ifndef BCGP_EXCLUDE_SVG
	for (int i = 0; i < (int)m_arSVG.GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = m_arSVG[i];
		if (pSVG != NULL)
		{
			pSVG->Mirror();
		}
	}
	
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return TRUE;
	}
#endif

	if (!MirrorBitmap (m_hbmImageWell, m_sizeImage.cx))
	{
		return FALSE;
	}

	if (m_hbmImageLight != NULL)
	{
		MirrorBitmap (m_hbmImageLight, m_sizeImage.cx);
	}

	if (m_hbmImageShadow != NULL)
	{
		MirrorBitmap (m_hbmImageShadow, m_sizeImage.cx);
	}

	return TRUE;
}
//****************************************************************************
BOOL CBCGPToolBarImages::MirrorBitmapVert (HBITMAP& hbmp, int cyImage)
{
	if (hbmp == NULL)
	{
		return TRUE;
	}

	BITMAP bmp;
	if (::GetObject (hbmp, sizeof (BITMAP), &bmp) == 0)
	{
		return FALSE;
	}

	const int cx = bmp.bmWidth;
	const int cy = bmp.bmHeight;
	const int iCount = cy / cyImage;

	if (bmp.bmBitsPixel >= 16)
	{
		DIBSECTION ds;
		if (::GetObject (hbmp, sizeof (DIBSECTION), &ds) == 0)
		{
			ASSERT (FALSE);
			return FALSE;
		}

		if (ds.dsBm.bmBitsPixel != bmp.bmBitsPixel)
		{
			ASSERT (FALSE);
			return FALSE;
		}

		LPBYTE pBits    = (LPBYTE)ds.dsBm.bmBits;
		DWORD pitch     = cx * ds.dsBm.bmBitsPixel / 8;
		if(pitch % 4)
		{
			pitch = (DWORD)(pitch / 4) * 4 + 4;
		}

		LPBYTE pRowTemp = new BYTE[pitch];

		for (int iImage = 0; iImage < iCount; iImage++)
		{
			LPBYTE pRowBits1 = pBits + iImage * cyImage * pitch;
			LPBYTE pRowBits2 = pRowBits1 + (cyImage - 1) * pitch;

			for (int y = 0; y < cyImage / 2; y++)
			{
				memcpy(pRowTemp, pRowBits1, pitch);
				memcpy(pRowBits1, pRowBits2, pitch);
				memcpy(pRowBits2, pRowTemp, pitch);

				pRowBits1 += pitch;
				pRowBits2 -= pitch;
			}
		}

		delete [] pRowTemp;

		return TRUE;
	}

	CDC memDC;
	memDC.CreateCompatibleDC (NULL);

	HBITMAP hOldBitmap = (HBITMAP) memDC.SelectObject (hbmp);
	if (hOldBitmap == NULL)
	{
		return FALSE;
	}

	for (int iImage = 0; iImage < iCount; iImage++)
	{
		int y1 = iImage * cyImage;
		int y2 = y1 + cyImage - 1;

		for (int y = 0; y < cyImage / 2; y++)
		{
			for (int x = 0; x < cx; x++)
			{
				COLORREF color = memDC.GetPixel (x, y1);

				memDC.SetPixel (x, y1, memDC.GetPixel (x, y2));
				memDC.SetPixel (x, y2, color);
			}

			y1++;
			y2--;
		}
	}

	memDC.SelectObject (hOldBitmap);

    return TRUE;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::MirrorVert ()
{
	if (!MirrorBitmapVert (m_hbmImageWell, m_sizeImage.cy))
	{
		return FALSE;
	}

	if (m_hbmImageLight != NULL)
	{
		MirrorBitmapVert (m_hbmImageLight, m_sizeImage.cy);
	}

	if (m_hbmImageShadow != NULL)
	{
		MirrorBitmapVert (m_hbmImageShadow, m_sizeImage.cy);
	}

	return TRUE;
}
//****************************************************************************
void CBCGPToolBarImages::EnableRTL (BOOL bIsRTL/* = TRUE*/)
{
	m_bIsRTL = bIsRTL;
}
//****************************************************************************
void CBCGPToolBarImages::AdaptColors (COLORREF clrBase, COLORREF clrTone, BOOL bClampHue/* = TRUE*/)
{
#ifndef BCGP_EXCLUDE_SVG
	for (int i = 0; i < (int)m_arSVG.GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = m_arSVG[i];
		if (pSVG != NULL)
		{
			pSVG->AdaptColors(clrBase, clrTone, bClampHue);
		}
	}

	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return;
	}
#endif

	double dSrcH, dSrcS, dSrcL;
	CBCGPDrawManager::RGBtoHSL (clrBase, &dSrcH, &dSrcS, &dSrcL);

	double dDestH, dDestS, dDestL;
	CBCGPDrawManager::RGBtoHSL (clrTone, &dDestH, &dDestS, &dDestL);

	double DH = dDestH - dSrcH;
	double DL = dDestL - dSrcL;
	double DS = dDestS - dSrcS;

	bClampHue = (dDestH > 0.5) && bClampHue;

	if (m_nBitsPerPixel >= 24)
	{
		DIBSECTION ds;
		if (::GetObject (m_hbmImageWell, sizeof (DIBSECTION), &ds) == 0)
		{
			ASSERT (FALSE);
			return;
		}

		if (ds.dsBm.bmBits != NULL)
		{
			if (ds.dsBm.bmBitsPixel < 24)
			{
				ASSERT (FALSE);
				return;
			}

			if (ds.dsBm.bmBitsPixel == 24)
			{
				COLORREF clrTransparent =
					m_clrTransparent == (COLORREF)-1 ? globalData.clrBtnFace : m_clrTransparent;

				DWORD dwPitch = ((ds.dsBm.bmWidth * 24 + 31) & ~31) / 8;
				LPBYTE pBits = (LPBYTE)ds.dsBm.bmBits;
				for (int y = 0; y < ds.dsBm.bmHeight; y++)
				{
					LPBYTE pBitsRow = pBits;
					for (int x = 0; x < ds.dsBm.bmWidth; x++)
					{
						COLORREF clrOrig = RGB (pBitsRow[2], pBitsRow[1], pBitsRow[0]);
						if (clrOrig == clrTransparent)
						{
							pBitsRow += 3;
							continue;
						}

						double H,S,L;
						CBCGPDrawManager::RGBtoHSL (clrOrig, &H, &S, &L);

						if (fabs (dSrcH - H) < m_dblColorTolerance)
						{
							double HNew = CBCGPDrawManager::MakeHue_ONE(H + DH);
							double SNew = max (0., min (1.00, S + DS));
							double LNew = max (0., min (1.00, L + DL));

							COLORREF color = CBCGPDrawManager::HLStoRGB_ONE (
								HNew, bClampHue ? L : LNew, SNew);

							pBitsRow[2] = (BYTE) (GetRValue (color));
							pBitsRow[1] = (BYTE) (GetGValue (color));
							pBitsRow[0] = (BYTE) (GetBValue (color));
						}

						pBitsRow += 3;
					}

					pBits += dwPitch;
				}
			}
			else
			{
				RGBQUAD* pBits = (RGBQUAD*) ds.dsBm.bmBits;
				for (int i = 0; i < ds.dsBm.bmWidth * ds.dsBm.bmHeight; i++)
				{
					RGBQUAD* pBit = pBits + i;

					if (pBit->rgbReserved == 0)
					{
						continue;
					}

					double a  = (double)pBit->rgbReserved / 255.0;
					double aR = 255.0 / (double)pBit->rgbReserved;
					COLORREF clrOrig = RGB (pBit->rgbRed * aR, pBit->rgbGreen * aR, pBit->rgbBlue * aR);

					double H,S,L;
					CBCGPDrawManager::RGBtoHSL (clrOrig, &H, &S, &L);

					if (fabs (dSrcH - H) < m_dblColorTolerance)
					{
						double HNew = CBCGPDrawManager::MakeHue_ONE(H + DH);
						double SNew = max (0., min (1.00, S + DS));
						double LNew = max (0., min (1.00, L + DL));

						COLORREF color = CBCGPDrawManager::HLStoRGB_ONE (
							HNew, bClampHue ? L : LNew, SNew);

						pBit->rgbRed   = (BYTE) (GetRValue (color) * a);
						pBit->rgbGreen = (BYTE) (GetGValue (color) * a);
						pBit->rgbBlue  = (BYTE) (GetBValue (color) * a);
					}
				}
			}

			return;
		}
	}

	//-------------------------------------------------------
	// Create memory source DC and select an original bitmap:
	//-------------------------------------------------------
	CDC memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	HBITMAP hOldBitmapSrc = NULL;

	int iBitmapWidth;
	int iBitmapHeight;

	//-------------------------------
	// Get original bitmap attrbutes:
	//-------------------------------
	BITMAP bmp;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
	{
		return;
	}

	hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (m_hbmImageWell);
	if (hOldBitmapSrc == NULL)
	{
		return;
	}

	iBitmapWidth = bmp.bmWidth;
	iBitmapHeight = bmp.bmHeight;

	//----------------------------------------------------------
	// Create a new bitmap compatibel with the source memory DC:
	// (original bitmap SHOULD BE ALREADY SELECTED!):
	//----------------------------------------------------------
	HBITMAP hNewBitmap = (HBITMAP) ::CreateCompatibleBitmap (memDCSrc,
									iBitmapWidth,
									iBitmapHeight);
	if (hNewBitmap == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		return;
	}

	//------------------------------
	// Create memory destination DC:
	//------------------------------
	CDC memDCDst;
	memDCDst.CreateCompatibleDC (&memDCSrc);

	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hNewBitmap);
	if (hOldBitmapDst == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		::DeleteObject (hNewBitmap);
		return;
	}

	//-----------------------------
	// Copy original bitmap to new:
	//-----------------------------
	memDCDst.BitBlt (0, 0, iBitmapWidth, iBitmapHeight,
				&memDCSrc, 0, 0, SRCCOPY);

	COLORREF clrTransparent =
		m_clrTransparent == (COLORREF)-1 ? globalData.clrBtnFace : m_clrTransparent;
		
	for (int x = 0; x < iBitmapWidth; x ++)
	{
		for (int y = 0; y < iBitmapHeight; y ++)
		{
			COLORREF clrOrig = ::GetPixel (memDCDst, x, y);

			if (clrOrig == clrTransparent)
			{
				continue;
			}

			double H, L, S;
			CBCGPDrawManager::RGBtoHSL (clrOrig, &H, &S, &L);

			if (fabs (dSrcH - H) < m_dblColorTolerance)
			{
				double HNew = CBCGPDrawManager::MakeHue_ONE(H + DH);
				double SNew = max (0., min (1.00, S + DS));
				double LNew = max (0., min (1.00, L + DL));

				COLORREF clrNew = CBCGPDrawManager::HLStoRGB_ONE (
					HNew, bClampHue ? L : LNew, SNew);

				if (clrOrig != clrNew)
				{
					::SetPixel (memDCDst, x, y, clrNew);
				}
			}
		}
	}
	
	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);

	::DeleteObject (m_hbmImageWell);
	m_hbmImageWell = hNewBitmap;

	CleanUpInternalImages();
}
//*****************************************************************************
void CBCGPToolBarImages::AdaptColors (COLORREF clrBase, COLORREF clrTone, double dOpacity/* = 1.0*/)
{
#ifndef BCGP_EXCLUDE_SVG
	for (int i = 0; i < (int)m_arSVG.GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = m_arSVG[i];
		if (pSVG != NULL)
		{
			pSVG->AdaptColors(clrBase, clrTone);
		}
	}
	
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return;
	}
#endif

	double dH, dS, dL;
	CBCGPDrawManager::RGBtoHSL (clrTone, &dH, &dS, &dL);

	BOOL bUseBase = clrBase != COLORREF(-1) && clrBase != CLR_DEFAULT;
	double dBaseH = 0.0, dBaseS, dBaseL;

	if (bUseBase)
	{
		CBCGPDrawManager::RGBtoHSL (clrBase, &dBaseH, &dBaseS, &dBaseL);
	}

	DWORD dwLum = RGB_LUM_COLOR (clrTone);

	if (m_nBitsPerPixel >= 24)
	{
		DIBSECTION ds;
		if (::GetObject (m_hbmImageWell, sizeof (DIBSECTION), &ds) == 0)
		{
			ASSERT (FALSE);
			return;
		}

		if (ds.dsBm.bmBits != NULL)
		{
			if (ds.dsBm.bmBitsPixel < 24)
			{
				ASSERT (FALSE);
				return;
			}

			if (ds.dsBm.bmBitsPixel == 24)
			{
				COLORREF clrTransparent =
					m_clrTransparent == (COLORREF)-1 ? globalData.clrBtnFace : m_clrTransparent;

				DWORD dwPitch = ((ds.dsBm.bmWidth * 24 + 31) & ~31) / 8;
				LPBYTE pBits = (LPBYTE)ds.dsBm.bmBits;
				for (int y = 0; y < ds.dsBm.bmHeight; y++)
				{
					LPBYTE pBitsRow = pBits;
					for (int x = 0; x < ds.dsBm.bmWidth; x++)
					{
						COLORREF clrOrig = RGB (pBitsRow[2], pBitsRow[1], pBitsRow[0]);
						if (clrOrig == clrTransparent)
						{
							pBitsRow += 3;
							continue;
						}

						if (bUseBase)
						{
							double H, L, S;
							CBCGPDrawManager::RGBtoHSL (clrOrig, &H, &S, &L);

							if (fabs (dBaseH - H) > m_dblColorTolerance)
							{
								pBitsRow += 3;
								continue;
							}
						}

						DWORD dwLumPix = RGB_LUM (pBitsRow[2], pBitsRow[1], pBitsRow[0]);
						if (dwLumPix <= 128)
						{
							dwLumPix = dwLum * dwLumPix / 128;
						}
						else
						{
							dwLumPix = (255 - dwLum) * (dwLumPix - 128) / 128 + dwLum;
						}

						COLORREF color = CBCGPDrawManager::HLStoRGB_ONE (
							dH, dwLumPix / 255.0, dS * dOpacity);

						pBitsRow[2] = (BYTE) (GetRValue (color));
						pBitsRow[1] = (BYTE) (GetGValue (color));
						pBitsRow[0] = (BYTE) (GetBValue (color));

						pBitsRow += 3;
					}

					pBits += dwPitch;
				}
			}
			else
			{
				RGBQUAD* pBits = (RGBQUAD*) ds.dsBm.bmBits;
				for (int i = 0; i < ds.dsBm.bmWidth * ds.dsBm.bmHeight; i++)
				{
					RGBQUAD* pBit = pBits + i;

					if (pBit->rgbReserved == 0)
					{
						continue;
					}

					if (bUseBase)
					{
						COLORREF clrOrig = RGB (pBit->rgbRed, pBit->rgbGreen, pBit->rgbBlue);

						double H, L, S;
						CBCGPDrawManager::RGBtoHSL (clrOrig, &H, &S, &L);

						if (fabs (dBaseH - H) > m_dblColorTolerance)
						{
							continue;
						}
					}

					double a  = (double)pBit->rgbReserved / 255.0;
					double aR = 255.0 / (double)pBit->rgbReserved;

					DWORD dwLumPix = RGB_LUM (pBit->rgbRed * aR, pBit->rgbGreen * aR, pBit->rgbBlue * aR);
					if (dwLumPix <= 128)
					{
						dwLumPix = dwLum * dwLumPix / 128;
					}
					else
					{
						dwLumPix = (255 - dwLum) * (dwLumPix - 128) / 128 + dwLum;
					}

					COLORREF color = CBCGPDrawManager::HLStoRGB_ONE (
						dH, dwLumPix / 255.0, dS * dOpacity);

					pBit->rgbRed   = (BYTE) (GetRValue (color) * a);
					pBit->rgbGreen = (BYTE) (GetGValue (color) * a);
					pBit->rgbBlue  = (BYTE) (GetBValue (color) * a);
				}
			}

			return;
		}
	}

	//-------------------------------------------------------
	// Create memory source DC and select an original bitmap:
	//-------------------------------------------------------
	CDC memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	HBITMAP hOldBitmapSrc = NULL;

	int iBitmapWidth;
	int iBitmapHeight;

	//-------------------------------
	// Get original bitmap attrbutes:
	//-------------------------------
	BITMAP bmp;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
	{
		return;
	}

	hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (m_hbmImageWell);
	if (hOldBitmapSrc == NULL)
	{
		return;
	}

	iBitmapWidth = bmp.bmWidth;
	iBitmapHeight = bmp.bmHeight;

	//----------------------------------------------------------
	// Create a new bitmap compatibel with the source memory DC:
	// (original bitmap SHOULD BE ALREADY SELECTED!):
	//----------------------------------------------------------
	HBITMAP hNewBitmap = (HBITMAP) ::CreateCompatibleBitmap (memDCSrc,
									iBitmapWidth,
									iBitmapHeight);
	if (hNewBitmap == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		return;
	}

	//------------------------------
	// Create memory destination DC:
	//------------------------------
	CDC memDCDst;
	memDCDst.CreateCompatibleDC (&memDCSrc);

	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hNewBitmap);
	if (hOldBitmapDst == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		::DeleteObject (hNewBitmap);
		return;
	}

	//-----------------------------
	// Copy original bitmap to new:
	//-----------------------------
	memDCDst.BitBlt (0, 0, iBitmapWidth, iBitmapHeight,
				&memDCSrc, 0, 0, SRCCOPY);

	COLORREF clrTransparent =
		m_clrTransparent == (COLORREF)-1 ? globalData.clrBtnFace : m_clrTransparent;
		
	for (int x = 0; x < iBitmapWidth; x ++)
	{
		for (int y = 0; y < iBitmapHeight; y ++)
		{
			COLORREF clrOrig = ::GetPixel (memDCDst, x, y);

			if (clrOrig == clrTransparent)
			{
				continue;
			}

			if (bUseBase)
			{
				double H, L, S;
				CBCGPDrawManager::RGBtoHSL (clrOrig, &H, &S, &L);

				if (fabs (dBaseH - H) > m_dblColorTolerance)
				{
					continue;
				}
			}

			DWORD dwLumPix = RGB_LUM_COLOR (clrOrig);
			if (dwLumPix <= 128)
			{
				dwLumPix = dwLum * dwLumPix / 128;
			}
			else
			{
				dwLumPix = (255 - dwLum) * (dwLumPix - 128) / 128 + dwLum;
			}

			COLORREF color = CBCGPDrawManager::HLStoRGB_ONE (
				dH, dwLumPix / 255.0, dS * dOpacity);

			if (clrOrig != color)
			{
				::SetPixel (memDCDst, x, y, color);
			}
		}
	}
	
	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);

	::DeleteObject (m_hbmImageWell);
	m_hbmImageWell = hNewBitmap;

	CleanUpInternalImages();
}
//****************************************************************************
void CBCGPToolBarImages::InvertColors()
{
#ifndef BCGP_EXCLUDE_SVG
	for (int i = 0; i < (int)m_arSVG.GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = m_arSVG[i];
		if (pSVG != NULL)
		{
			pSVG->InvertColors();
		}
	}
	
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return;
	}
#endif

	BCGPInvertBitmapColors(m_hbmImageWell, m_clrTransparent == (COLORREF)-1 ? globalData.clrBtnFace : m_clrTransparent);
	CleanUpInternalImages();
}
//*****************************************************************************
void CBCGPToolBarImages::ConvertToGrayScale(double dblLumRatio, BOOL bAutoConvertAfterReLoad)
{
#ifndef BCGP_EXCLUDE_SVG
	for (int i = 0; i < (int)m_arSVG.GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = m_arSVG[i];
		if (pSVG != NULL)
		{
			pSVG->ConvertToGrayScale(dblLumRatio);
		}
	}
	
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return;
	}
#endif

	if (m_hbmImageWell != NULL)
	{
		_ConvertToGrayScale(m_hbmImageWell, m_clrTransparent, dblLumRatio);
	}

	CleanUpInternalImages();

	m_bAutoGrayAfterReLoad = bAutoConvertAfterReLoad;
	m_dblGrayLumRatio = dblLumRatio;
}
//*****************************************************************************
void CBCGPToolBarImages::MakeLighter(double dblRatio)
{
#ifndef BCGP_EXCLUDE_SVG
	for (int i = 0; i < (int)m_arSVG.GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = m_arSVG[i];
		if (pSVG != NULL)
		{
			pSVG->MakeLighter(dblRatio);
		}
	}
	
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return;
	}
#endif

	BCGPModifyBitmapLuminosity(m_hbmImageWell, m_clrTransparent == (COLORREF)-1 ? globalData.clrBtnFace : m_clrTransparent, dblRatio + 1.0);
	CleanUpInternalImages();
}
//*****************************************************************************
void CBCGPToolBarImages::MakeDarker(double dblRatio)
{
#ifndef BCGP_EXCLUDE_SVG
	for (int i = 0; i < (int)m_arSVG.GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = m_arSVG[i];
		if (pSVG != NULL)
		{
			pSVG->MakeDarker(dblRatio);
		}
	}
	
	if (m_nSVGCount == m_iCount && m_iCount > 0)
	{
		return;
	}
#endif
	BCGPModifyBitmapLuminosity(m_hbmImageWell, m_clrTransparent == (COLORREF)-1 ? globalData.clrBtnFace : m_clrTransparent, dblRatio);
	CleanUpInternalImages();
}
//*****************************************************************************
void CBCGPToolBarImages::_ConvertToGrayScale(HBITMAP& hbmp, COLORREF clrTransparent, double dblLumRatio)
{
	BCGPDesaturateBitmap(hbmp, clrTransparent == (COLORREF)-1 ? globalData.clrBtnFace : clrTransparent, FALSE, dblLumRatio);
}
//*****************************************************************************
HRGN CBCGPToolBarImages::CreateRegionFromImage (HBITMAP hbmp, COLORREF clrTransparent)
{
	if (hbmp == NULL)
	{
		ASSERT (FALSE);
		return NULL;
	}

	BITMAP bmp;
	if (::GetObject (hbmp, sizeof (BITMAP), &bmp) == 0)
	{
		ASSERT (FALSE);
		return NULL;
	}

	CDC dc;
	dc.CreateCompatibleDC (NULL);

	HBITMAP hbmpOld = (HBITMAP) dc.SelectObject (hbmp);

	int iBitmapWidth = bmp.bmWidth;
	int iBitmapHeight = bmp.bmHeight;

	CRgn rgnAll;
	rgnAll.CreateRectRgn (0, 0, iBitmapWidth, iBitmapHeight);

	for (int y = 0; y < iBitmapHeight; y++)
	{
		for (int x = 0; x < iBitmapWidth; x++)
		{
			COLORREF color = dc.GetPixel (x, y);

			if (color == clrTransparent)
			{
				CRgn rgnPoint;
				rgnPoint.CreateRectRgn (x, y, x + 1, y + 1);

				rgnAll.CombineRgn (&rgnAll, &rgnPoint, RGN_DIFF);
			}
		}
	}

	dc.SelectObject (hbmpOld);

	return (HRGN) rgnAll.Detach ();
}
//*******************************************************************************
void CBCGPToolBarImages::SetSingleImage (BOOL bUpdateInternalImages)
{
#ifndef BCGP_EXCLUDE_SVG
	if (m_nSVGCount == 1 && m_arSVG.GetSize() == 1)
	{
		CBCGPSVGImage* pSVG = m_arSVG[0];
		if (pSVG != NULL)
		{
			m_iCount = 1;

			const CSize sizeSVG = pSVG->GetSize();

			if (m_sizeImage != sizeSVG)
			{
				m_sizeImage = sizeSVG;

				AfxDeleteObject((HGDIOBJ*)&m_hbmImageWell);
				m_hbmImageWell = CBCGPDrawManager::CreateBitmap_32(m_sizeImage, NULL);
			}
		}
		
		return;
	}
#endif

	if (m_hbmImageWell == NULL)
	{
		return;
	}

	BITMAP bmp;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmp) == 0)
	{
		ASSERT (FALSE);
		return;
	}

	m_sizeImage.cx = bmp.bmWidth;
	m_sizeImage.cy = bmp.bmHeight;

	m_iCount = 1;
	
	if (m_nSVGCount == 1)
	{
		return;
	}

	if (bUpdateInternalImages)
	{
		UpdateInternalImage (BCGImage_Light);
		UpdateInternalImage (BCGImage_Shadow);

		AfxDeleteObject((HGDIOBJ*)&m_hbmImageGray);
		m_hbmImageGray = NULL;
	}
	
#ifndef BCGP_EXCLUDE_SVG
	CleanUpSVG();
	m_arSVG.Add(NULL);
#endif
}
//*******************************************************************************
HBITMAP CBCGPToolBarImages::Copy (HBITMAP hbmpSrc)
{
	if (hbmpSrc == NULL)
	{
		ASSERT (FALSE);
		return NULL;
	}

	if (!globalData.bIsWindowsVista)
	{
		return (HBITMAP) ::CopyImage (hbmpSrc, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
	}

	//-------------------------------------------------------
	// Windows Vista has some problems in ::CopyImage method,
	// copy bitmap not using this method:
	//-------------------------------------------------------

	CDC memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (hbmpSrc);
	if (hOldBitmapSrc == NULL)
	{
		return NULL;
	}

	BITMAP bmp;
	::GetObject (hbmpSrc, sizeof (BITMAP), &bmp);

	//----------------------------------------------------------
	// Create a new bitmap compatible with the source memory DC:
	//----------------------------------------------------------
	HBITMAP hNewBitmap = NULL;
	DIBSECTION ds = {0};

	if (bmp.bmBitsPixel >= 24 &&
		::GetObject (hbmpSrc, sizeof (DIBSECTION), &ds) != 0)
	{
		BITMAPINFO bi = {0};
		bi.bmiHeader.biSize        = sizeof (BITMAPINFOHEADER);
		bi.bmiHeader.biWidth       = bmp.bmWidth;
		bi.bmiHeader.biHeight      = bmp.bmHeight;
		bi.bmiHeader.biPlanes      = bmp.bmPlanes;
		bi.bmiHeader.biBitCount    = bmp.bmBitsPixel;
		bi.bmiHeader.biCompression = BI_RGB;

		COLORREF* pBits = NULL;
		hNewBitmap = ::CreateDIBSection (
			memDCSrc, &bi, DIB_RGB_COLORS, (void **)&pBits,
			NULL, NULL);
	}

	if (hNewBitmap == NULL)
	{
		hNewBitmap = (HBITMAP) ::CreateCompatibleBitmap (memDCSrc,
											bmp.bmWidth,
											bmp.bmHeight);
	}

	if (hNewBitmap == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		return NULL;
	}

	//------------------------------------------------------
	// Create memory destination DC and select a new bitmap:
	//------------------------------------------------------
	CDC memDCDst;
	memDCDst.CreateCompatibleDC (&memDCSrc);
	
	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hNewBitmap);
	if (hOldBitmapDst == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		::DeleteObject (hNewBitmap);
		return NULL;
	}

	//-----------------------------
	// Copy original bitmap to new:
	//-----------------------------
	memDCDst.BitBlt (0, 0, bmp.bmWidth, bmp.bmHeight, &memDCSrc, 0, 0, SRCCOPY);
		
	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);

	return hNewBitmap;
}
//*******************************************************************************
BOOL CBCGPToolBarImages::SmoothResize (double dblImageScale)
{
	if (m_hbmImageWell == NULL)
	{
		return FALSE;
	}

	if (m_nBitsPerPixel < 24)
	{
		return FALSE;
	}

	if (dblImageScale == 0.0)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	if (dblImageScale == 1.0)
	{
		return TRUE;
	}

	CSize sizeNew ((int)(.5 + m_sizeImage.cx * dblImageScale), (int)(.5 + m_sizeImage.cy * dblImageScale));

	if (sizeNew == m_sizeImage ||
		m_sizeImage.cx <= 0 || m_sizeImage.cy <= 0 ||
		sizeNew.cx <= 0 || sizeNew.cy <= 0)
	{
		return TRUE;
	}

	int nImageCount = GetCount ();
	if (nImageCount == 0)
	{
		return TRUE;
	}

	int nCY = m_sizeImage.cy;

	BOOL bInvert = FALSE;
	CSize sizeIW(0, 0);
	{
		BITMAP bmp;
		if (::GetObject (GetImageWell (), sizeof (BITMAP), &bmp) == 0)
		{
			ASSERT (FALSE);
			return FALSE;
		}

		sizeIW.cx = bmp.bmWidth;
		sizeIW.cy = abs(bmp.bmHeight);
		bInvert = bmp.bmHeight < 0;

#ifdef _DEBUG
		if (sizeIW.cx != nImageCount * m_sizeImage.cx || (sizeIW.cy % m_sizeImage.cy) != 0)
		{
			if (nCY != sizeIW.cy)
			{
				nCY = sizeIW.cy;
			}

			TRACE0("CBCGPToolBarImages::SmoothResize: possible incorrect internal size of the image. Please check the image size(s).\n");
		}
#endif
	}

	double dblScaleOld = m_dblScale;

	m_dblScale *= dblImageScale;

	CPoint offSrc (m_sizeImage.cx, 0);
	CPoint offDst (0, 0);
	
	if (nImageCount == 1)
	{
		if (sizeIW.cy > m_sizeImage.cy)
		{
			nImageCount = sizeIW.cy / m_sizeImage.cy;
			offSrc = CPoint (0, m_sizeImage.cy);
		}
	}

	HBITMAP hBmpSrc = CBCGPDrawManager::CreateBitmap_32 (m_hbmImageWell, m_clrTransparent);
	if (hBmpSrc == NULL)
	{
		m_dblScale = dblScaleOld;
		TRACE0("CBCGPToolBarImages::SmoothResize: cannot convert 24bpp image list to 32bpp.\n");
		return FALSE;
	}

	CSize sizeNewIW (sizeNew);
	if (offSrc.x > 0)
	{
		sizeNewIW.cx *= nImageCount;
		offDst.x = sizeNew.cx;
	}
	else
	{
		sizeNewIW.cy *= nImageCount;
		offDst.y = sizeNew.cy;
	}

	if (bInvert)
	{
		sizeNewIW.cy = -sizeNewIW.cy;
	}

	HBITMAP hBmpDst = CBCGPDrawManager::CreateBitmap_32 (sizeNewIW, NULL);

	sizeNewIW.cy = abs(sizeNewIW.cy);

	if (hBmpDst == NULL)
	{
		ASSERT(FALSE);
		::DeleteObject (hBmpSrc);
		return FALSE;
	}

	if (m_nSVGCount < m_iCount)
	{
		CBCGPZoomKernel::XFilterType ft = dblImageScale < 1.0
											? CBCGPZoomKernel::e_FilterTypeLanczos3
											: CBCGPZoomKernel::e_FilterTypeMitchell;

		CBCGPScanlinerBitmap ms;
		ms.Attach (hBmpSrc);
		CBCGPScanlinerBitmap md;
		md.Attach (hBmpDst);

		DWORD channel = ms.GetChannels ();

		CBCGPZoomKernel KernelX;
		KernelX.Create(m_sizeImage.cx, sizeNew.cx, 0, m_sizeImage.cx, ft);

		CBCGPZoomKernel KernelY;
		KernelY.Create(nCY, sizeNew.cy, 0, nCY, ft);

		double values[4] = {0.0, 0.0, 0.0, 0.0};
		double values2[4] = {0.0, 0.0, 0.0, 0.0};

		CPoint offSrcSum (0, 0);
		CPoint offDstSum (0, 0);

		for (int index = 0; index < nImageCount; index++)
		{
			const DWORD val_size   = sizeof(double) * channel;
			const DWORD offsetDstX = offDstSum.x * channel;

			for (DWORD dy = 0; dy < (DWORD)sizeNew.cy; dy++)
			{
				const CBCGPZoomKernel::XKernelList& listY = KernelY[dy];

				LPBYTE pRowDst = md[dy + offDstSum.y] + offsetDstX;

				for (DWORD dx = 0; dx < (DWORD)sizeNew.cx; dx++)
				{
					const CBCGPZoomKernel::XKernelList& listX = KernelX[dx];

					memset(values, 0, val_size);

					for (DWORD sy = 0; sy < listY.count; sy++)
					{
						const CBCGPZoomKernel::XKernel& statY = listY.stat[sy];

						const LPBYTE pRowSrc = ms[statY.pixel + offSrcSum.y];
						double weight    = statY.weight;

						memset(values2, 0, val_size);

						for (DWORD sx = 0; sx < listX.count; sx++)
						{
							const CBCGPZoomKernel::XKernel& statX = listX.stat[sx];

							LPBYTE pRowSrc2 = pRowSrc + (statX.pixel + offSrcSum.x) * channel;
							double weight2    = statX.weight;

							for(DWORD c = 0; c < channel; c++)
							{
								values2[c] += (double)(*pRowSrc2) * weight2;
								pRowSrc2++;
							}
						}

						for(DWORD c = 0; c < channel; c++)
						{
							values[c] += values2[c] * weight;
						}
					}

					for(DWORD c = 0; c < channel; c++)
					{
						*pRowDst = (BYTE)bcg_clamp_to_byte(values[c]);
						pRowDst++;
					}

					if (channel == 4)
					{
						BCGPCorrectAlpha(pRowDst - 4, *(pRowDst - 1));
					}
				}
			}

			offSrcSum.x += offSrc.x;
			offSrcSum.y += offSrc.y;
			offDstSum.x += offDst.x;
			offDstSum.y += offDst.y;
		}
	}

	::DeleteObject (hBmpSrc);

	int nOldCount = m_iCount;

	if (m_sizeImageOriginal == CSize (0, 0))
	{
		m_sizeImageOriginal = m_sizeImage;
	}

	SetImageSize (sizeNew);
	m_clrTransparentOriginal = m_clrTransparent;
	m_clrTransparent = (COLORREF)-1;

	if (!m_bIsTemporary)
	{
		AfxDeleteObject((HGDIOBJ*)&m_hbmImageWell);
	}

	m_hbmImageWell = hBmpDst;

	m_iCount = nOldCount;
	m_nBitsPerPixel = 32;
	
	CleanUpInternalImages();

	return IsValid ();
}

BOOL CBCGPToolBarImages::ConvertTo32Bits (COLORREF clrTransparent)
{
	if (!IsValid ())
	{
		ASSERT(FALSE);
		return FALSE;
	}

	if (m_nBitsPerPixel == 32)
	{
		return TRUE;
	}

	HBITMAP hbmpNew = CBCGPDrawManager::CreateBitmap_32 (m_hbmImageWell, clrTransparent == (COLORREF)-1 ? m_clrTransparent : clrTransparent);
	if (hbmpNew == NULL)
	{
		return FALSE;
	}

	AfxDeleteObject((HGDIOBJ*)&m_hbmImageWell);

	m_hbmImageWell = hbmpNew;
	m_clrTransparent = (COLORREF)-1;
	m_nBitsPerPixel = 32;

	CleanUpInternalImages();

	return TRUE;
}

BOOL CBCGPToolBarImages::ConvertTo24Bits()
{
	if (!IsValid ())
	{
		ASSERT(FALSE);
		return FALSE;
	}

	if (m_nBitsPerPixel >= 24)
	{
		return TRUE;
	}

	HBITMAP hbmpNew = CBCGPDrawManager::CreateBitmap_24(m_hbmImageWell);
	if (hbmpNew == NULL)
	{
		return FALSE;
	}

	AfxDeleteObject((HGDIOBJ*)&m_hbmImageWell);

	m_hbmImageWell = hbmpNew;
	m_nBitsPerPixel = 24;

	CleanUpInternalImages();

	return TRUE;
}

BOOL CBCGPToolBarImages::ChangeBpp(HBITMAP hbmp)
{
	if (hbmp == NULL || m_hbmImageWell == NULL)
	{
		return FALSE;
	}

	BITMAP bmp;
	if (::GetObject (hbmp, sizeof (BITMAP), &bmp) == 0)
	{
		return FALSE;
	}

	BITMAP bmpOrig;
	if (::GetObject (m_hbmImageWell, sizeof (BITMAP), &bmpOrig) == 0)
	{
		return FALSE;
	}

	if (bmpOrig.bmBitsPixel >= bmp.bmBitsPixel)
	{
		return FALSE;
	}

	if (bmp.bmBitsPixel == 32)
	{
		return ConvertTo32Bits();
	}

	CDC memDCSrc;
	memDCSrc.CreateCompatibleDC (NULL);

	HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject (hbmp);
	if (hOldBitmapSrc == NULL)
	{
		return FALSE;
	}

	//----------------------------------------------------------
	// Create a new bitmap compatibel with the source memory DC:
	//----------------------------------------------------------
	HBITMAP hNewBitmap = (HBITMAP) ::CreateCompatibleBitmap (memDCSrc,
									bmpOrig.bmWidth,
									bmpOrig.bmHeight);
	if (hNewBitmap == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		return FALSE;
	}

	//------------------------------------------------------
	// Create memory destination DC and select a new bitmap:
	//------------------------------------------------------
	CDC memDCDst;
	memDCDst.CreateCompatibleDC (&memDCSrc);
	
	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hNewBitmap);
	if (hOldBitmapDst == NULL)
	{
		memDCSrc.SelectObject (hOldBitmapSrc);
		::DeleteObject (hNewBitmap);
		return FALSE;
	}

	memDCSrc.SelectObject (m_hbmImageWell);

	//-----------------------------
	// Copy original bitmap to new:
	//-----------------------------
	memDCDst.BitBlt (0, 0, bmpOrig.bmWidth, bmpOrig.bmHeight, &memDCSrc, 0, 0, SRCCOPY);
		
	memDCDst.SelectObject (hOldBitmapDst);
	memDCSrc.SelectObject (hOldBitmapSrc);

	AfxDeleteObject((HGDIOBJ*)&m_hbmImageWell);

	m_hbmImageWell = hNewBitmap;
	m_nBitsPerPixel = bmp.bmBitsPixel;

	CleanUpInternalImages();

	return TRUE;
}
