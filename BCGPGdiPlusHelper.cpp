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
// BCGPGdiPlusHelper.cpp: implementation of the CBCGPGdiPlusHelper class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BCGPGdiPlusHelper.h"
#include "BCGPMath.h"
#include "BCGPDrawManager.h"

int	CBCGPGdiPlusHelper::m_nCount = 0;
CCriticalSection CBCGPGdiPlusHelper::g_cs;

CBCGPGdiPlusHelper::CBCGPGdiPlusHelper()
	: m_dwToken    (0)
	, m_dwTokenHook(0)
{
}


CBCGPGdiPlusHelper::~CBCGPGdiPlusHelper()
{
}


BOOL
CBCGPGdiPlusHelper::Initialize(BOOL SuppressBackgroundThread)
{
#ifndef BCGP_EXCLUDE_GDI_PLUS

	if (!IsInitialized())
	{
		if (SuppressBackgroundThread)
		{
			m_SI.SuppressBackgroundThread = TRUE;
		}

		Gdiplus::Status status = Gdiplus::GdiplusStartup(&m_dwToken, &m_SI, &m_SO);

		if (status == Gdiplus::Ok)
		{
			if (SuppressBackgroundThread)
			{
				m_SO.NotificationHook(&m_dwTokenHook);
			}

			g_cs.Lock ();
			m_nCount++;
			g_cs.Unlock ();
		}
		else
		{
			m_dwToken = 0;
		}
	}

	return m_dwToken != 0;
#else

	UNREFERENCED_PARAMETER(SuppressBackgroundThread);

	return FALSE;

#endif
}


void
CBCGPGdiPlusHelper::Shutdown()
{
#ifndef BCGP_EXCLUDE_GDI_PLUS
	if (IsInitialized())
	{
		if (m_dwTokenHook != 0)
		{
			m_SO.NotificationUnhook(m_dwTokenHook);
		}

		Gdiplus::GdiplusShutdown(m_dwToken);

		g_cs.Lock ();
		if (m_nCount > 0)
		{
			m_nCount--;
		}
		g_cs.Unlock ();
	}

	m_dwToken = 0;
	m_dwTokenHook = 0;
#endif
}

#ifndef BCGP_EXCLUDE_GDI_PLUS

#define BCG_DELETE_GDIPLUS_OBJECT(object) if (object != NULL) { delete object; }

#define BCG_DELETE_GDIPLUS_ARRAY(array) if (object != NULL) { delete [] array; }

Gdiplus::Graphics*
CBCGPGdiPlusHelper::CreateGraphics(CDC* pDC)
{
	if (pDC->GetSafeHdc() == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}

	return Gdiplus::Graphics::FromHDC(pDC->GetSafeHdc());
}


void
CBCGPGdiPlusHelper::DeleteGraphics(Gdiplus::Graphics* object)
{
	BCG_DELETE_GDIPLUS_OBJECT(object);
}


Gdiplus::Pen*
CBCGPGdiPlusHelper::CreatePen(const Gdiplus::Color& color, Gdiplus::REAL width/* = 1.0f*/)
{
	return new Gdiplus::Pen(color, width);
}


Gdiplus::Pen*
CBCGPGdiPlusHelper::CreatePen(const Gdiplus::Brush* brush, Gdiplus::REAL width/* = 1.0f*/)
{
	return new Gdiplus::Pen(brush, width);
}


void
CBCGPGdiPlusHelper::DeletePen(Gdiplus::Pen* object)
{
	BCG_DELETE_GDIPLUS_OBJECT(object);
}


Gdiplus::Brush*
CBCGPGdiPlusHelper::CreateSolidBrush(const Gdiplus::Color& color)
{
	return new Gdiplus::SolidBrush(color);
}


Gdiplus::LinearGradientBrush*
CBCGPGdiPlusHelper::CreateLinearGradientBrush(const Gdiplus::PointF& point1, const Gdiplus::PointF& point2, const Gdiplus::Color& color1, const Gdiplus::Color& color2)
{
	return new Gdiplus::LinearGradientBrush(point1, point2, color1, color2);
}


Gdiplus::PathGradientBrush*
CBCGPGdiPlusHelper::CreatePathGradientBrush(const Gdiplus::GraphicsPath* path)
{
	return new Gdiplus::PathGradientBrush(path);
}


Gdiplus::TextureBrush*
CBCGPGdiPlusHelper::CreateTextureBrush(Gdiplus::Image* image, Gdiplus::WrapMode wrapMode, Gdiplus::REAL opacity)
{
	if (image == NULL)
	{
		return NULL;
	}

	Gdiplus::ImageAttributes* attr = NULL;

	opacity = (Gdiplus::REAL)bcg_clamp(opacity, 0.0, 1.0);
	if (opacity != 1.0)
	{
		Gdiplus::ColorMatrix clrMatrix;
		memset(clrMatrix.m, 0, sizeof(clrMatrix.m));
		clrMatrix.m[0][0] = (Gdiplus::REAL)1.0;
		clrMatrix.m[1][1] = (Gdiplus::REAL)1.0;
		clrMatrix.m[2][2] = (Gdiplus::REAL)1.0;
		clrMatrix.m[3][3] = (Gdiplus::REAL)opacity;
		clrMatrix.m[4][4] = (Gdiplus::REAL)1.0;

		attr = CBCGPGdiPlusHelper::CreateImageAttributes();
		if (attr != NULL)
		{
			attr->SetColorMatrix(&clrMatrix, Gdiplus::ColorMatrixFlagsDefault, Gdiplus::ColorAdjustTypeBitmap);
		}
	}

	Gdiplus::TextureBrush* brush = new Gdiplus::TextureBrush(image, Gdiplus::Rect(0, 0, image->GetWidth(), image->GetHeight()), attr);
	if (brush != NULL)
	{
		brush->SetWrapMode(wrapMode);
	}

	CBCGPGdiPlusHelper::DeleteImageAttributes(attr);

	return brush;
}


void
CBCGPGdiPlusHelper::DeleteBrush(Gdiplus::Brush* object)
{
	BCG_DELETE_GDIPLUS_OBJECT(object);
}


Gdiplus::GraphicsPath*
CBCGPGdiPlusHelper::CreateGraphicsPath()
{
	return new Gdiplus::GraphicsPath;
}


Gdiplus::GraphicsPath*
CBCGPGdiPlusHelper::CreateRoundedRectanglePath(const Gdiplus::RectF& rectangle, Gdiplus::REAL radiusX, Gdiplus::REAL radiusY)
{
	Gdiplus::RectF rect(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
	CBCGPGdiPlusHelper::NormalizeRectF(rect);

	if (rect.IsEmptyArea() || radiusX <= 0 || radiusY <= 0)
	{
		return NULL;
	}

	radiusX *= 2.0f;
	radiusY *= 2.0f;

	Gdiplus::GraphicsPath* pPath = CBCGPGdiPlusHelper::CreateGraphicsPath();
	if (pPath == NULL)
	{
		return NULL;
	}

	if (radiusX > rect.Width)
	{
		radiusX = rect.Width;
	}
	if (radiusY > rect.Height)
	{
		radiusY = rect.Height;
	}

	pPath->AddArc
		(
			rect.X,
			rect.Y,
			radiusX,
			radiusY,
			180.0,
			90.0
		);
	pPath->AddArc
		(
			rect.GetRight() - radiusX,
			rect.Y,
			radiusX,
			radiusY,
			270.0,
			90.0
		);
	pPath->AddArc
		(
			rect.GetRight() - radiusX,
			rect.GetBottom() - radiusY,
			radiusX,
			radiusY,
			0.0,
			90.0
		);
	pPath->AddArc
		(
			rect.X,
			rect.GetBottom() - radiusY,
			radiusX,
			radiusY,
			90.0,
			90.0
		);

	pPath->CloseFigure();

	return pPath;
}


void
CBCGPGdiPlusHelper::DeleteGraphicsPath(Gdiplus::GraphicsPath* object)
{
	BCG_DELETE_GDIPLUS_OBJECT(object);
}


Gdiplus::Region*
CBCGPGdiPlusHelper::CreateRegion()
{
	return new Gdiplus::Region;
}


Gdiplus::Region*
CBCGPGdiPlusHelper::CreateRegion(Gdiplus::GraphicsPath* path)
{
	return new Gdiplus::Region(path);
}


void
CBCGPGdiPlusHelper::DeleteRegion(Gdiplus::Region* object)
{
	BCG_DELETE_GDIPLUS_OBJECT(object);
}


Gdiplus::Image*
CBCGPGdiPlusHelper::CreateImage(HBITMAP hBitmap, BOOL ignoreAlpha)
{
	if (hBitmap == NULL)
	{
		return NULL;
	}

	Gdiplus::Bitmap* bitmap = NULL;

	BITMAP bm = {0};
	::GetObject(hBitmap, sizeof(BITMAP), &bm);

	if (bm.bmBitsPixel != 32 || ignoreAlpha)
	{
		bitmap = Gdiplus::Bitmap::FromHBITMAP(hBitmap, NULL);
	}
	else
	{
		HBITMAP hBitmapDst = NULL;
		LPBYTE lpData = NULL;
		BOOL bInverted = TRUE;

		DIBSECTION ds = {0};
		if (::GetObject(hBitmap, sizeof(DIBSECTION), &ds) == sizeof(DIBSECTION))
		{
			bInverted = ds.dsBm.bmHeight < 0;
			lpData = (LPBYTE)ds.dsBm.bmBits;
		}

		if (lpData == NULL)
		{
			hBitmapDst = CBCGPDrawManager::CreateBitmap_32(CSize(bm.bmWidth, -bm.bmHeight), (LPVOID*)&lpData);

			if (hBitmapDst != NULL)
			{
				HDC hdcSrc = ::CreateCompatibleDC(NULL);
				HGDIOBJ hbmpSrcOld = ::SelectObject(hdcSrc, hBitmap);

				HDC hdcDst = ::CreateCompatibleDC(NULL);
				HGDIOBJ hbmpDstOld = ::SelectObject(hdcDst, hBitmapDst);

				::BitBlt(hdcDst, 0, 0, bm.bmWidth, bm.bmHeight, hdcSrc, 0, 0, SRCCOPY);

				::SelectObject(hdcSrc, hbmpSrcOld);
				::DeleteDC(hdcSrc);

				::SelectObject(hdcDst, hbmpDstOld);
				::DeleteDC(hdcDst);
			}
		}

		if (lpData != NULL)
		{
			bitmap = new Gdiplus::Bitmap(bm.bmWidth, bm.bmHeight, PixelFormat32bppPARGB);

			if (bitmap != NULL)
			{
				Gdiplus::BitmapData data = {0};
				Gdiplus::Rect rect(0, 0, bm.bmWidth, bm.bmHeight);

				bitmap->LockBits(&rect, Gdiplus::ImageLockModeWrite, PixelFormat32bppPARGB, &data);
				if (data.Scan0 != NULL)
				{
					if (bInverted)
					{
						memcpy(data.Scan0, lpData, bm.bmWidth * bm.bmHeight * 4);
					}
					else
					{
						const DWORD pitch = bm.bmWidth * 4;
						lpData = lpData + (bm.bmHeight - 1) * pitch;
						LPBYTE pScan = (LPBYTE)data.Scan0;

						for (LONG y = 0; y < bm.bmHeight; y++)
						{
							memcpy(pScan, lpData, pitch);
							pScan += data.Stride;
							lpData -= pitch;
						}
					}

					bitmap->UnlockBits(&data);
				}
				else
				{
					delete bitmap;
					bitmap = NULL;
				}
			}
		}

		if (hBitmapDst != NULL)
		{
			::DeleteObject(hBitmapDst);
		}
	}

	return bitmap;
}


Gdiplus::Image*
CBCGPGdiPlusHelper::CreateImage(HICON hIcon)
{
	if (hIcon == NULL)
	{
		return NULL;
	}

	return Gdiplus::Bitmap::FromHICON(hIcon);
}


HBITMAP
CBCGPGdiPlusHelper::CreateHBITMAP(Gdiplus::Image* image)
{
	if (image == NULL)
	{
		return NULL;
	}

	HBITMAP hBitmap = NULL;

	Gdiplus::PixelFormat eSrcPixelFormat = image->GetPixelFormat();
	UINT nBPP = 32;
	Gdiplus::PixelFormat eDestPixelFormat = PixelFormat32bppRGB;

	if (eSrcPixelFormat & PixelFormatGDI)
	{
		nBPP = Gdiplus::GetPixelFormatSize(eSrcPixelFormat);
		eDestPixelFormat = eSrcPixelFormat;
	}

	if (Gdiplus::IsAlphaPixelFormat(eSrcPixelFormat))
	{
		nBPP = 32;
		eDestPixelFormat = PixelFormat32bppARGB;
	}

	const UINT width = image->GetWidth();
	const UINT height = image->GetHeight();
	int pitch = (((width * nBPP) + 31) / 32) * 4;

	LPBYTE lpBits = NULL;

	LPBITMAPINFO pbmi = NULL;
	UINT pbmiLength = sizeof(BITMAPINFOHEADER);

	if( nBPP <= 8 )
	{
		pbmiLength += sizeof(RGBQUAD) * 256;
	}

	pbmi = (LPBITMAPINFO)(new BYTE[pbmiLength]);
	if(pbmi != NULL)
	{
		memset(pbmi, 0, pbmiLength);

		pbmi->bmiHeader.biSize = sizeof(pbmi->bmiHeader);
		pbmi->bmiHeader.biWidth = width;
		pbmi->bmiHeader.biHeight = height;
		pbmi->bmiHeader.biPlanes = 1;
		pbmi->bmiHeader.biBitCount = USHORT(nBPP);
		pbmi->bmiHeader.biCompression = BI_RGB;

		hBitmap = ::CreateDIBSection(NULL, pbmi, DIB_RGB_COLORS, (LPVOID*)&lpBits, NULL, 0);

		delete [] pbmi;

		if (hBitmap != NULL)
		{
/*
			USES_ATL_SAFE_ALLOCA;
			Gdiplus::ColorPalette* pPalette = NULL;
			if (Gdiplus::IsIndexedPixelFormat(eSrcPixelFormat))
			{
				UINT nPaletteSize = image->GetPaletteSize();
				pPalette = static_cast<Gdiplus::ColorPalette*>(_ATL_SAFE_ALLOCA(nPaletteSize, _ATL_SAFE_ALLOCA_DEF_THRESHOLD));

				if( pPalette == NULL )
					return E_OUTOFMEMORY;

				image->GetPalette(pPalette, nPaletteSize);

				RGBQUAD argbPalette[256];
				ATLENSURE_RETURN((pPalette->Count > 0) && (pPalette->Count <= 256));
				for(UINT iColor = 0; iColor < pPalette->Count; iColor++)
				{
					Gdiplus::ARGB color = pPalette->Entries[iColor];
					argbPalette[iColor].rgbRed = (BYTE)((color>>RED_SHIFT) & 0xff);
					argbPalette[iColor].rgbGreen = (BYTE)((color>>GREEN_SHIFT) & 0xff);
					argbPalette[iColor].rgbBlue = (BYTE)((color>>BLUE_SHIFT) & 0xff);
					argbPalette[iColor].rgbReserved = 0;
				}

				SetColorTable(0, pPalette->Count, argbPalette);
			}
*/
			lpBits += (height - 1) * pitch;
			pitch = -pitch;

			if (eDestPixelFormat == eSrcPixelFormat && image->GetType() == Gdiplus::ImageTypeBitmap)
			{
				Gdiplus::Bitmap* bitmap = (Gdiplus::Bitmap*)image;

				// The pixel formats are identical, so just memcpy the rows.
				Gdiplus::BitmapData data;
				Gdiplus::Rect rect(0, 0, width, height);
				if (bitmap->LockBits(&rect, Gdiplus::ImageLockModeRead, eSrcPixelFormat, &data) == Gdiplus::Ok)
				{
					size_t nBytesPerRow = abs(pitch);
					BYTE* pbDestRow = lpBits;
					const BYTE* pbSrcRow = static_cast<const BYTE*>(data.Scan0);

					for(UINT y = 0; y < height; y++)
					{
						memcpy(pbDestRow, pbSrcRow, nBytesPerRow);

						pbDestRow += pitch;
						pbSrcRow += data.Stride;
					}

					bitmap->UnlockBits(&data);
				}
				else
				{
					::DeleteObject(hBitmap);
					hBitmap = NULL;
				}
			}
			else
			{
				// Let GDI+ work its magic
				Gdiplus::Bitmap bmDest(width, height, pitch, eDestPixelFormat, lpBits);
				Gdiplus::Graphics gDest(&bmDest);

				gDest.DrawImage(image, Gdiplus::RectF((Gdiplus::REAL)0, (Gdiplus::REAL)0, (Gdiplus::REAL)width, (Gdiplus::REAL)height));
			}
		}
	}

	return hBitmap;
}


void
CBCGPGdiPlusHelper::DeleteImage(Gdiplus::Image* object)
{
	BCG_DELETE_GDIPLUS_OBJECT(object);
}


Gdiplus::ImageAttributes*
CBCGPGdiPlusHelper::CreateImageAttributes()
{
	return new Gdiplus::ImageAttributes;
}


void
CBCGPGdiPlusHelper::DeleteImageAttributes(Gdiplus::ImageAttributes* object)
{
	BCG_DELETE_GDIPLUS_OBJECT(object);
}


Gdiplus::Font*
CBCGPGdiPlusHelper::CreateFont(const CString& family, Gdiplus::REAL emSize, INT style, Gdiplus::Unit unit/* = Gdiplus::UnitWorld*/)
{
	LPWSTR lpFamily = NULL;

#ifdef _UNICODE
	lpFamily = (LPWSTR)(LPCWSTR)family;
#else
	{
		int length = (int)family.GetLength() + 1;
		lpFamily = (LPWSTR) new WCHAR[length];
		lpFamily = AfxA2WHelper(lpFamily, (LPCSTR)family, length);
	}
#endif

	Gdiplus::Font* font = new Gdiplus::Font
		(
			lpFamily,
			emSize,
			style,
			unit
		);

	if ((LPVOID)lpFamily != (LPVOID)(LPCTSTR)family)
	{
		delete [] lpFamily;
	}

	return font;
}


void
CBCGPGdiPlusHelper::DeleteFont(Gdiplus::Font* object)
{
	BCG_DELETE_GDIPLUS_OBJECT(object);
}


Gdiplus::StringFormat*
CBCGPGdiPlusHelper::CreateStringFormat(INT formatFlags, LANGID language)
{
	return new Gdiplus::StringFormat(formatFlags, language);
}


void
CBCGPGdiPlusHelper::DeleteStringFormat(Gdiplus::StringFormat* object)
{
	BCG_DELETE_GDIPLUS_OBJECT(object);
}


Gdiplus::REAL*
CBCGPGdiPlusHelper::CreateArrayF(int size)
{
	if (size == 0)
	{
		return NULL;
	}

	return new Gdiplus::REAL[size];
}


void
CBCGPGdiPlusHelper::DeleteArrayF(Gdiplus::REAL* object)
{
	BCG_DELETE_GDIPLUS_ARRAY(object);
}


Gdiplus::PointF*
CBCGPGdiPlusHelper::CreateArrayPointF(int size)
{
	if (size == 0)
	{
		return NULL;
	}

	return new Gdiplus::PointF[size];
}


void
CBCGPGdiPlusHelper::DeleteArrayPointF(Gdiplus::PointF* object)
{
	BCG_DELETE_GDIPLUS_ARRAY(object);
}


Gdiplus::Color*
CBCGPGdiPlusHelper::CreateArrayColor(int size)
{
	if (size == 0)
	{
		return NULL;
	}

	return new Gdiplus::Color[size];
}


void
CBCGPGdiPlusHelper::DeleteArrayColor(Gdiplus::Color* object)
{
	BCG_DELETE_GDIPLUS_ARRAY(object);
}


Gdiplus::Matrix*
CBCGPGdiPlusHelper::CreateMatrix()
{
	return new Gdiplus::Matrix;
}


void
CBCGPGdiPlusHelper::DeleteMatrix(Gdiplus::Matrix* object)
{
	BCG_DELETE_GDIPLUS_OBJECT(object);
}


Gdiplus::Color
CBCGPGdiPlusHelper::Color(COLORREF clr, BYTE a)
{
	return Gdiplus::Color(a, GetRValue(clr), GetGValue(clr), GetBValue(clr));
}


Gdiplus::Color
CBCGPGdiPlusHelper::Color(BYTE r, BYTE g, BYTE b, BYTE a)
{
	return Gdiplus::Color(a, r, g, b);
}


void
CBCGPGdiPlusHelper::NormalizeRectF(Gdiplus::RectF& rect)
{
	if (rect.Width < 0.0f)
	{
		rect.Width = -rect.Width;
		rect.X -= rect.Width;
	}

	if (rect.Height < 0.0f)
	{
		rect.Height = -rect.Height;
		rect.Y -= rect.Height;
	}
}

#endif
