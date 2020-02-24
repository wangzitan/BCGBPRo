//*******************************************************************************
// COPYRIGHT NOTES
// ---------------
// This is a part of BCGControlBar Library Professional Edition
// Copyright (C) 1998-2018 BCGSoft Ltd.
// All rights reserved.
//
// This source code can be used, distributed or modified
// only under terms and conditions 
// of the accompanying license agreement.
//*******************************************************************************
//
// BCGPGdiPlusHelper.h: interface for the CBCGPGdiPlusHelper class.
//
//////////////////////////////////////////////////////////////////////

#ifndef BCGPGdiPlusHelper_H
#define BCGPGdiPlusHelper_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXMT_H__
	#include <afxmt.h>
#endif

#include "BCGCBPro.h"
#include "BCGPGdiPlus.h"

class CBCGPGdiPlusHelper
{
public:
							CBCGPGdiPlusHelper();
							~CBCGPGdiPlusHelper();

	static	BOOL 			IsReady()
							{
								return m_nCount > 0;
							}

			BOOL			IsInitialized() const
							{
								return m_dwToken != 0;
							}

			BOOL			Initialize(BOOL SuppressBackgroundThread = FALSE);
			void			Shutdown();

#ifndef BCGP_EXCLUDE_GDI_PLUS
	static	Gdiplus::Graphics*
							CreateGraphics(CDC* pDC);
	static	void			DeleteGraphics(Gdiplus::Graphics* object);

	static	Gdiplus::Pen*	CreatePen(const Gdiplus::Color& color, Gdiplus::REAL width = 1.0f);
	static	Gdiplus::Pen*	CreatePen(const Gdiplus::Brush* brush, Gdiplus::REAL width = 1.0f);
	static	void			DeletePen(Gdiplus::Pen* object);

	static	Gdiplus::Brush* CreateSolidBrush(const Gdiplus::Color& color);
	static	Gdiplus::LinearGradientBrush*
							CreateLinearGradientBrush(const Gdiplus::PointF& point1, const Gdiplus::PointF& point2, const Gdiplus::Color& color1, const Gdiplus::Color& color2);
	static	Gdiplus::PathGradientBrush*
							CreatePathGradientBrush(const Gdiplus::GraphicsPath* path);
	static	Gdiplus::TextureBrush*
							CreateTextureBrush(Gdiplus::Image* image, Gdiplus::WrapMode wrapMode, Gdiplus::REAL opacity);
	static	void			DeleteBrush(Gdiplus::Brush* object);

	static	Gdiplus::GraphicsPath*
							CreateGraphicsPath();
	static	Gdiplus::GraphicsPath*
							CreateRoundedRectanglePath(const Gdiplus::RectF& rectangle, Gdiplus::REAL radiusX, Gdiplus::REAL radiusY);
	static	void			DeleteGraphicsPath(Gdiplus::GraphicsPath* object);

	static	Gdiplus::Region*
							CreateRegion();
	static	Gdiplus::Region*
							CreateRegion(Gdiplus::GraphicsPath* path);
	static	void			DeleteRegion(Gdiplus::Region* object);

	static	Gdiplus::Image* CreateImage(HBITMAP hBitmap, BOOL ignoreAlpha);
	static	Gdiplus::Image* CreateImage(HICON hIcon);
	static	HBITMAP			CreateHBITMAP(Gdiplus::Image* image);
	static	void			DeleteImage(Gdiplus::Image* object);

	static	Gdiplus::ImageAttributes*
							CreateImageAttributes();
	static	void			DeleteImageAttributes(Gdiplus::ImageAttributes* object);

	static	Gdiplus::Font*	CreateFont(const CString& family, Gdiplus::REAL emSize, INT style, Gdiplus::Unit unit = Gdiplus::UnitWorld);
	static	void			DeleteFont(Gdiplus::Font* object);

	static	Gdiplus::StringFormat*
							CreateStringFormat(INT formatFlags, LANGID language);
	static	void			DeleteStringFormat(Gdiplus::StringFormat* object);

	static	Gdiplus::REAL*	CreateArrayF(int size);
	static	void			DeleteArrayF(Gdiplus::REAL* object);

	static	Gdiplus::PointF*
							CreateArrayPointF(int size);
	static	void			DeleteArrayPointF(Gdiplus::PointF* object);

	static	Gdiplus::Color* CreateArrayColor(int size);
	static	void			DeleteArrayColor(Gdiplus::Color* object);

	static	Gdiplus::Matrix*
							CreateMatrix();
	static	void			DeleteMatrix(Gdiplus::Matrix* object);

	static	Gdiplus::Color	Color(COLORREF clr, BYTE a = 255);
	static	Gdiplus::Color	Color(BYTE r, BYTE g, BYTE b, BYTE a = 255);

	static	void			NormalizeRectF(Gdiplus::RectF& rect);

private:
			Gdiplus::GdiplusStartupInput
							m_SI;
			Gdiplus::GdiplusStartupOutput
							m_SO;

#endif

private:
			ULONG_PTR		m_dwToken;
			ULONG_PTR		m_dwTokenHook;
			static int		m_nCount;
			static CCriticalSection g_cs;			// For multi-thread applications
};

#endif
