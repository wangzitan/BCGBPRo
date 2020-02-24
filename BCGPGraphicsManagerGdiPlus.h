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
// BCGPGraphicsManagerGdiPlus.h: interface for the CBCGPGraphicsManagerGdiPlus class.
//
//////////////////////////////////////////////////////////////////////

#ifndef BCGPGraphicsManagerGDIPlus_H
#define BCGPGraphicsManagerGDIPlus_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "BCGCBPro.h"
#include "BCGPGraphicsManager.h"

class BCGCBPRODLLEXPORT CBCGPGraphicsManagerGdiPlus : public CBCGPGraphicsManager
{
	friend struct BCGPGLOBAL_DATA;

	DECLARE_DYNCREATE(CBCGPGraphicsManagerGdiPlus)

public:
							CBCGPGraphicsManagerGdiPlus(CDC* pDC = NULL, BOOL bDoubleBuffering = TRUE, CBCGPGraphicsManagerParams* pParams = NULL);
	virtual 				~CBCGPGraphicsManagerGdiPlus();

	virtual CBCGPGraphicsManager*
							CreateOffScreenManager(const CBCGPRect& rect, CBCGPImage* pImageDest);

	static BOOL InitGDIPlus(BOOL bSuppressBackgroundThread = FALSE);
	static void Shutdown();

	static BOOL IsGDIPlusInitialized();

protected:
							CBCGPGraphicsManagerGdiPlus(const CBCGPRect& rectDest, CBCGPImage* pImageDest);

public:
	virtual BOOL			IsValid() const;

	virtual void			BindDC(CDC* pDC, BOOL bDoubleBuffering = TRUE);
	virtual void			BindDC(CDC* pDC, const CRect& rect);

	static BOOL m_bCheckLineOffsets;
	static BOOL m_bRectOffsets;

// overrides:
public:
	virtual BOOL			BeginDraw();
	virtual void			EndDraw();

	virtual void			Clear(const CBCGPColor& color = CBCGPColor());

	virtual void			DrawLine
							(
								const CBCGPPoint& ptFrom, const CBCGPPoint& ptTo, const CBCGPBrush& brush, 
								double lineWidth = 1., const CBCGPStrokeStyle* pStrokeStyle = NULL
							);

	virtual void			DrawLines
							(
								const CBCGPPointsArray& arPoints, const CBCGPBrush& brush, 
								double lineWidth = 1., const CBCGPStrokeStyle* pStrokeStyle = NULL
							);

	virtual void			DrawScatter
							(
								const CBCGPPointsArray& arPoints, const CBCGPBrush& brush, double dblPointSize, UINT nStyle = 0
							);

	virtual void			DrawRectangle
							(
								const CBCGPRect& rect, const CBCGPBrush& brush, 
								double lineWidth = 1., const CBCGPStrokeStyle* pStrokeStyle = NULL
							);

	virtual void			FillRectangle(const CBCGPRect& rect, const CBCGPBrush& brush);

	virtual void			DrawRoundedRectangle
							(
								const CBCGPRoundedRect& rect, const CBCGPBrush& brush, 
								double lineWidth = 1., const CBCGPStrokeStyle* pStrokeStyle = NULL
							);

	virtual void			FillRoundedRectangle(const CBCGPRoundedRect& rect, const CBCGPBrush& brush);

	virtual void			DrawEllipse
							(
								const CBCGPEllipse& ellipse, const CBCGPBrush& brush, 
								double lineWidth = 1., const CBCGPStrokeStyle* pStrokeStyle = NULL
							);

	virtual void			FillEllipse(const CBCGPEllipse& ellipse, const CBCGPBrush& brush);

	virtual void			DrawArc
							(
								const CBCGPEllipse& ellipse, double dblStartAngle, double dblFinishAngle, BOOL bIsClockwise,
								const CBCGPBrush& brush, double lineWidth = 1., const CBCGPStrokeStyle* pStrokeStyle = NULL
							);

	virtual void			DrawArc
							(
								const CBCGPPoint& ptFrom, const CBCGPPoint& ptTo, const CBCGPSize sizeRadius, 
								BOOL bIsClockwise, BOOL bIsLargeArc,
								const CBCGPBrush& brush, double lineWidth = 1., const CBCGPStrokeStyle* pStrokeStyle = NULL
							);

	virtual void			DrawGeometry
							(
								const CBCGPGeometry& geometry, const CBCGPBrush& brush, 
								double lineWidth = 1., const CBCGPStrokeStyle* pStrokeStyle = NULL
							);

	virtual void			FillGeometry(const CBCGPGeometry& geometry, const CBCGPBrush& brush);

	virtual void			DrawImage
							(
								const CBCGPImage& image, const CBCGPPoint& ptDest, const CBCGPSize& sizeDest = CBCGPSize(),
								double opacity = 1., CBCGPImage::BCGP_IMAGE_INTERPOLATION_MODE interpolationMode = CBCGPImage::BCGP_IMAGE_INTERPOLATION_MODE_LINEAR, 
								const CBCGPRect& rectSrc = CBCGPRect()
							);

	virtual void			DrawText
							(
								const CString& strText, const CBCGPRect& rectText, const CBCGPTextFormat& textFormat,
								const CBCGPBrush& foregroundBrush
							);

	virtual CBCGPSize		GetTextSize(const CString& strText, const CBCGPTextFormat& textFormat, double dblWidth = 0., BOOL bIgnoreTextRotation = FALSE);

	virtual void			SetClipArea(const CBCGPGeometry& geometry, int nFlags = RGN_COPY);
	virtual void			ReleaseClipArea();

	virtual void			CombineGeometry(CBCGPGeometry& geometryDest, const CBCGPGeometry& geometrySrc1, const CBCGPGeometry& geometrySrc2, int nFlags);
	virtual void			GetGeometryBoundingRect(const CBCGPGeometry& geometry, CBCGPRect& rect);

	virtual void			EnableAntialiasing(BOOL bEnable = TRUE);
	virtual BOOL			IsAntialiasingEnabled() const;

	virtual BOOL			IsTransformSupported() const;

	virtual void			ResetTransform();

	virtual void			PushTransform();
	virtual BOOL			PopTransform();

	virtual void			SetRotateTransform(double dblAngle, const CBCGPPoint& ptCenter);
	virtual void			SetScaleTransform(const CBCGPSize& size, const CBCGPPoint& ptCenter = CBCGPPoint(-1., -1.));
	virtual void			SetTranslateTransform(const CBCGPSize& size);
	virtual void			SetMatrixTransform(const CArray<double, double>& arValues);

	virtual void			SetMask(const CBCGPGeometry& /*geometry*/, const CBCGPBrush& /*opacityBrush*/, const CBCGPTransformMatrix& /*transform*/ = CBCGPTransformMatrix()) {}
	virtual void			SetMask(const CBCGPGeometry& /*geometry*/, double /*opacity*/, const CBCGPTransformMatrix& /*transform*/ = CBCGPTransformMatrix()) {}
	virtual void			ReleaseMask() {}

protected:
	virtual LPVOID			CreateGeometry(CBCGPGeometry& geometry);
	virtual BOOL			DestroyGeometry(CBCGPGeometry& geometry);

	virtual LPVOID			CreateTextFormat(CBCGPTextFormat& textFormat);
	virtual BOOL			DestroyTextFormat(CBCGPTextFormat& textFormat);

	virtual LPVOID			CreateBrush(CBCGPBrush& brush);
	virtual BOOL			DestroyBrush(CBCGPBrush& brush);
	virtual BOOL			SetBrushOpacity(CBCGPBrush& brush);

	virtual LPVOID			CreateImage(CBCGPImage& image);
	virtual BOOL			DestroyImage(CBCGPImage& image);
	virtual CBCGPSize		GetImageSize(CBCGPImage& image);
	virtual BOOL			CopyImage(CBCGPImage& imageSrc, CBCGPImage& imageDest, const CBCGPRect& rectSrc);
	virtual HBITMAP 		ExportImageToBitmap(CBCGPImage& image);

	virtual LPVOID			CreateStrokeStyle(CBCGPStrokeStyle& style);
	virtual BOOL			DestroyStrokeStyle(CBCGPStrokeStyle& style);

			void			PrepareBrush(CBCGPBrush& brush, LPVOID pBrushIn, const CBCGPRect& rectIn) const;
			void			PrepareTextFormat(CBCGPTextFormat& textFormat) const;

			BOOL			IsGraphicsManagerValid(CBCGPGraphicsManager* pGM);

// Attributes:
protected:
			CDC*			m_pDC;
			CDC*			m_pDCPaint;
			CBCGPMemDC* 	m_pMemDC;
			BOOL			m_bDoubleBuffering;
			CBitmap 		m_bmpScreen;
			CBitmap*		m_pBmpScreenOld;
			CDC 			m_dcScreen;
			LPVOID			m_pGraphics;

			CList<LPVOID, LPVOID>
							m_lstTransforms;
			CBCGPSize		m_sizeScaleTransform;
};

#endif
