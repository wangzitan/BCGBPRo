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
// BCGPGraphicsManagerGdiPlus.cpp : implementation of the CBCGPGraphicsManagerGdiPlus class.
//
//////////////////////////////////////////////////////////////////////

#include "StdAfx.h"
#include "BCGPGdiplusHelper.h"
#include "BCGPGraphicsManagerGDIPlus.h"
#include "BCGPMath.h"
#include "BCGPDrawManager.h"
#include "BCGPSVGImage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

CBCGPGdiPlusHelper g_GdiPlusHelper;

#ifndef BCGP_EXCLUDE_GDI_PLUS

#define GPGraphics(graphics) ((Gdiplus::Graphics*)graphics)

inline BOOL NeedPrepareBrush(const CBCGPBrush& brush)
{
	return brush.HasTextureImage() || 
			(brush.GetGradientType() != CBCGPBrush::BCGP_NO_GRADIENT && 
			brush.GetGradientType() != CBCGPBrush::BCGP_GRADIENT_BEVEL);
}

inline Gdiplus::Color GPColor(const CBCGPColor& value, double opacity = 1.0)
{
	return Gdiplus::Color
				(
					(BYTE)bcg_clamp_to_byte(value.a * opacity * 255.0), 
					(BYTE)bcg_clamp_to_byte(value.r * 255.0), 
					(BYTE)bcg_clamp_to_byte(value.g * 255.0), 
					(BYTE)bcg_clamp_to_byte(value.b * 255.0)
				);
}

inline Gdiplus::PointF GPPoint(const CBCGPGraphicsManager* pGM, double x, double y, BOOL bPixelOffset = FALSE)
{
	Gdiplus::PointF pointGP((Gdiplus::REAL)x, (Gdiplus::REAL)y);

	if (bPixelOffset)
	{
		pointGP.X = (Gdiplus::REAL)((int)pointGP.X + (Gdiplus::REAL)0.5);
		pointGP.Y = (Gdiplus::REAL)((int)pointGP.Y + (Gdiplus::REAL)0.5);
	}

	if (pGM->IsDrawShadowMode())
	{
		pointGP.X += (Gdiplus::REAL)pGM->GetShadowOffset().x;
		pointGP.Y += (Gdiplus::REAL)pGM->GetShadowOffset().y;
	}

	return pointGP;
}

inline Gdiplus::PointF GPPoint(const CBCGPGraphicsManager* pGM, const CBCGPPoint& value, BOOL bPixelOffset = FALSE)
{
	return GPPoint(pGM, (Gdiplus::REAL)value.x, (Gdiplus::REAL)value.y, bPixelOffset);
}

inline Gdiplus::SizeF GPSize(const CBCGPSize& value)
{
	return Gdiplus::SizeF((Gdiplus::REAL)value.cx, (Gdiplus::REAL)value.cy);
}

inline Gdiplus::RectF GPRect(const CBCGPGraphicsManager* pGM, const CBCGPRect& value, BOOL bPixelOffset = FALSE)
{
	CBCGPRect rectGP(value);

	if (bPixelOffset)
	{
		rectGP.left = (int)rectGP.left + .5;
		rectGP.right = (int)rectGP.right + .5;
		rectGP.top = (int)rectGP.top + .5;
		rectGP.bottom = (int)rectGP.bottom + .5;
	}

	if (pGM->IsDrawShadowMode())
	{
		rectGP.left += pGM->GetShadowOffset().x;
		rectGP.top += pGM->GetShadowOffset().y;
		rectGP.right += pGM->GetShadowOffset().x;
		rectGP.bottom += pGM->GetShadowOffset().y;
	}

	return Gdiplus::RectF((Gdiplus::REAL)rectGP.left, (Gdiplus::REAL)rectGP.top, (Gdiplus::REAL)rectGP.Width(), (Gdiplus::REAL)rectGP.Height());
}

inline double GetMatrixValue(const CArray<double, double>& arValues, int nIndex, double size = 0.0)
{
	double val = nIndex >= arValues.GetSize() ? 0.0 : arValues[nIndex];

	if (size > 0.0 && fabs(val) > 0 && fabs(val) <= 1.0)
	{
		val *= size;
	}

	return val;
}

inline Gdiplus::REAL GetMatrixValueR(const CArray<double, double>& arValues, int nIndex, double size = 0.0)
{
	return (Gdiplus::REAL)GetMatrixValue(arValues, nIndex, size);
}

inline void GPMatrix(Gdiplus::Matrix& matrix, const CArray<double, double>& arMatrix, const CBCGPSize& size = CBCGPSize())
{
	matrix.SetElements
				(
					GetMatrixValueR(arMatrix, 0), 
					GetMatrixValueR(arMatrix, 1),
					GetMatrixValueR(arMatrix, 2),
					GetMatrixValueR(arMatrix, 3),
					GetMatrixValueR(arMatrix, 4, size.cx),
					GetMatrixValueR(arMatrix, 5, size.cy)
				);
}

class CGPPen
{
public:
							CGPPen(Gdiplus::Brush* pBrush, double width, const CBCGPStrokeStyle* style)
								: m_pPen(NULL)
							{
								m_pPen = CBCGPGdiPlusHelper::CreatePen(pBrush, (Gdiplus::REAL)width);

								if (m_pPen != NULL && style != NULL && !style->IsEmpty())
								{
									m_pPen->SetDashStyle(ConvertDashStyle(style->GetDashStyle()));
									if (style->GetDashStyle() == CBCGPStrokeStyle::BCGP_DASH_STYLE_CUSTOM && 
										style->GetDashes().GetSize() > 0)
									{
										m_pPen->SetDashPattern(style->GetDashes().GetData(), (INT)style->GetDashes().GetSize());
									}
									m_pPen->SetDashCap(ConvertDashCap(style->GetDashCap()));
									m_pPen->SetDashOffset((Gdiplus::REAL)style->GetDashOffset());

									m_pPen->SetStartCap(ConvertLineCap(style->GetStartCap()));
									m_pPen->SetEndCap(ConvertLineCap(style->GetEndCap()));
									m_pPen->SetLineJoin(ConvertLineJoin(style->GetLineJoin()));
									m_pPen->SetMiterLimit((Gdiplus::REAL)style->GetMitterLimit());
								}
							}

							~CGPPen()
							{
								CBCGPGdiPlusHelper::DeletePen(m_pPen);
							}

							operator Gdiplus::Pen*()
							{
								return m_pPen;
							}
							operator const Gdiplus::Pen*() const
							{
								return m_pPen;
							}

	static	Gdiplus::DashStyle
							ConvertDashStyle(CBCGPStrokeStyle::BCGP_DASH_STYLE value)
							{
								return (Gdiplus::DashStyle)value;
							}
	static	Gdiplus::DashCap
							ConvertDashCap(CBCGPStrokeStyle::BCGP_CAP_STYLE value)
							{
								return (Gdiplus::DashCap)value;
							}
	static	Gdiplus::LineCap
							ConvertLineCap(CBCGPStrokeStyle::BCGP_CAP_STYLE value)
							{
								return (Gdiplus::LineCap)value;
							}
	static	Gdiplus::LineJoin
							ConvertLineJoin(CBCGPStrokeStyle::BCGP_LINE_JOIN value)
							{
								return (Gdiplus::LineJoin)value;
							}

private:
			Gdiplus::Pen* m_pPen;
};

class CGPStoreTransform
{
public:
							CGPStoreTransform(Gdiplus::Brush* pBrush)
								: m_pBrush	 (pBrush)
								, m_pGraphics(NULL)
							{
								if (m_pBrush != NULL)
								{
									Gdiplus::BrushType type = m_pBrush->GetType();
									if (type == Gdiplus::BrushTypeLinearGradient)
									{
										((Gdiplus::LinearGradientBrush*)m_pBrush)->GetTransform(&m_Matrix);
									}
									else if (type == Gdiplus::BrushTypePathGradient)
									{
										((Gdiplus::PathGradientBrush*)m_pBrush)->GetTransform(&m_Matrix);
									}
									else if (type == Gdiplus::BrushTypeTextureFill)
									{
										((Gdiplus::TextureBrush*)m_pBrush)->GetTransform(&m_Matrix);
									}
									else
									{
										m_pBrush = NULL;
									}
								}
							}
							CGPStoreTransform(Gdiplus::Graphics* pGraphics)
								: m_pBrush	 (NULL)
								, m_pGraphics(pGraphics)
							{
								if (m_pGraphics != NULL)
								{
									GPGraphics(m_pGraphics)->GetTransform(&m_Matrix);
								}
							}

							~CGPStoreTransform()
							{
								if (m_pBrush != NULL)
								{
									Gdiplus::BrushType type = m_pBrush->GetType();
									if (type == Gdiplus::BrushTypeLinearGradient)
									{
										((Gdiplus::LinearGradientBrush*)m_pBrush)->SetTransform(&m_Matrix);
									}
									else if (type == Gdiplus::BrushTypePathGradient)
									{
										((Gdiplus::PathGradientBrush*)m_pBrush)->SetTransform(&m_Matrix);
									}
									else if (type == Gdiplus::BrushTypeTextureFill)
									{
										((Gdiplus::TextureBrush*)m_pBrush)->SetTransform(&m_Matrix);
									}
								}
								else if (m_pGraphics != NULL)
								{
									GPGraphics(m_pGraphics)->SetTransform(&m_Matrix);
								}
							}

private:
			Gdiplus::Brush* m_pBrush;
			Gdiplus::Graphics*
							m_pGraphics;
			Gdiplus::Matrix m_Matrix;
};

class CGPStorePen
{
public:
							CGPStorePen(Gdiplus::Pen* pPen, double width, const CBCGPStrokeStyle* style)
								: m_pPen			(pPen)
								, m_pStyle			(style)
								, m_DashPatternCount(0)
								, m_DashPattern 	(NULL)
							{
								if (m_pPen != NULL)
								{
									m_Width = m_pPen->GetWidth();
									m_pPen->SetWidth((Gdiplus::REAL)width);

									if (m_pStyle != NULL)
									{
										m_DashStyle = m_pPen->GetDashStyle();
										m_pPen->SetDashStyle(CGPPen::ConvertDashStyle(m_pStyle->GetDashStyle()));

										if (m_DashStyle == Gdiplus::DashStyleCustom)
										{
											m_DashPatternCount = m_pPen->GetDashPatternCount();
											if (m_DashPatternCount > 0)
											{
												m_DashPattern = CBCGPGdiPlusHelper::CreateArrayF(m_DashPatternCount);
												m_pPen->GetDashPattern(m_DashPattern, m_DashPatternCount);
											}
										}
										if (m_pStyle->GetDashStyle() == CBCGPStrokeStyle::BCGP_DASH_STYLE_CUSTOM && 
											m_pStyle->GetDashes().GetSize() > 0)
										{
											m_pPen->SetDashPattern(m_pStyle->GetDashes().GetData(), (INT)m_pStyle->GetDashes().GetSize());
										}

										m_DashCap = m_pPen->GetDashCap();
										m_pPen->SetDashCap(CGPPen::ConvertDashCap(m_pStyle->GetDashCap()));

										m_DashOffset = m_pPen->GetDashOffset();
										m_pPen->SetDashOffset((Gdiplus::REAL)m_pStyle->GetDashOffset());

										m_LineCapStart = m_pPen->GetStartCap();
										m_pPen->SetStartCap(CGPPen::ConvertLineCap(m_pStyle->GetStartCap()));

										m_LineCapEnd = m_pPen->GetEndCap();
										m_pPen->SetEndCap(CGPPen::ConvertLineCap(m_pStyle->GetEndCap()));

										m_LineJoin = m_pPen->GetLineJoin();
										m_pPen->SetLineJoin(CGPPen::ConvertLineJoin(m_pStyle->GetLineJoin()));

										m_MitterLimit = m_pPen->GetMiterLimit();
										m_pPen->SetMiterLimit((Gdiplus::REAL)m_pStyle->GetMitterLimit());
									}
								}
							}

							~CGPStorePen()
							{
								if (m_pPen != NULL)
								{
									m_pPen->SetWidth(m_Width);

									if (m_pStyle != NULL)
									{
										m_pPen->SetDashStyle(m_DashStyle);
										if (m_DashPatternCount > 0 && m_DashPattern != NULL)
										{
											m_pPen->SetDashPattern(m_DashPattern, m_DashPatternCount);
											CBCGPGdiPlusHelper::DeleteArrayF(m_DashPattern);
										}
										m_pPen->SetDashCap(m_DashCap);
										m_pPen->SetDashOffset(m_DashOffset);

										m_pPen->SetStartCap(m_LineCapStart);
										m_pPen->SetEndCap(m_LineCapEnd);
										m_pPen->SetLineJoin(m_LineJoin);
										m_pPen->SetMiterLimit(m_MitterLimit);
									}
								}
							}

private:
			Gdiplus::Pen*	m_pPen;
			const CBCGPStrokeStyle*
							m_pStyle;

			Gdiplus::REAL	m_Width;

			Gdiplus::DashStyle
							m_DashStyle;
			Gdiplus::DashCap
							m_DashCap;
			Gdiplus::REAL	m_DashOffset;
			int 			m_DashPatternCount;
			Gdiplus::REAL*	m_DashPattern;

			Gdiplus::LineCap
							m_LineCapStart;
			Gdiplus::LineCap
							m_LineCapEnd;
			Gdiplus::LineJoin
							m_LineJoin;
			Gdiplus::REAL	m_MitterLimit;
};

class CGPGeometry
{
public:
							CGPGeometry(LPVOID object, bool region)
								: m_Object(object)
								, m_Region(region)
							{
							}
							~CGPGeometry()
							{
								DeleteObject();
							}

			void			DeleteObject()
							{
								if (m_Region)
								{
									CBCGPGdiPlusHelper::DeleteRegion((Gdiplus::Region*)m_Object);
								}
								else
								{
									CBCGPGdiPlusHelper::DeleteGraphicsPath((Gdiplus::GraphicsPath*)m_Object);
								}
							}

			void			CreateRegion(const CGPGeometry& geometry)
							{
								if (geometry.IsRegion() == m_Region && geometry.GetObject() == m_Object)
								{
									return;
								}

								LPVOID object = NULL;

								if (geometry.IsRegion())
								{
									object = (LPVOID)((Gdiplus::Region*)geometry.GetObject())->Clone();
								}
								else
								{
									object = (LPVOID)CBCGPGdiPlusHelper::CreateRegion((Gdiplus::GraphicsPath*)geometry.GetObject());
								}

								DeleteObject();

								m_Object = object;
								m_Region = true;
							}

			void			Combine(const CGPGeometry& geometry, int nFlags)
							{
								if (m_Object == NULL || !m_Region)
								{
									ASSERT(FALSE);
									return;
								}

								Gdiplus::Region* region = (Gdiplus::Region*)m_Object;

								if (geometry.IsRegion())
								{
									Gdiplus::Region* pObject = (Gdiplus::Region*)geometry.GetObject();

									switch (nFlags)
									{
									case RGN_OR:
										region->Union(pObject);
										break;

									case RGN_AND:
										region->Intersect(pObject);
										break;

									case RGN_XOR:
										region->Xor(pObject);
										break;

									case RGN_DIFF:
										region->Exclude(pObject);
										break;
									}
								}
								else
								{
									Gdiplus::GraphicsPath* pObject = (Gdiplus::GraphicsPath*)geometry.GetObject();

									switch (nFlags)
									{
									case RGN_OR:
										region->Union(pObject);
										break;

									case RGN_AND:
										region->Intersect(pObject);
										break;

									case RGN_XOR:
										region->Xor(pObject);
										break;

									case RGN_DIFF:
										region->Exclude(pObject);
										break;
									}
								}
							}

			LPVOID			GetObject()
							{
								return m_Object;
							}
			LPCVOID 		GetObject() const
							{
								return m_Object;
							}
			bool			IsRegion() const
							{
								return m_Region;
							}

			void			GetBounds(Gdiplus::Graphics* graphics, CBCGPRect& rect) const
							{
								if (graphics == NULL || m_Object == NULL)
								{
									return;
								}

								if (m_Region)
								{
									HRGN rgn = ((Gdiplus::Region*)m_Object)->GetHRGN(graphics);
									if (rgn != NULL)
									{
										CRect rt;
										::GetRgnBox(rgn, rt);
										::DeleteObject(rgn);

										rect = rt;
									}
								}
								else
								{
									Gdiplus::RectF rt;
									((Gdiplus::GraphicsPath*)m_Object)->GetBounds(&rt);

									rect.left	= rt.GetLeft();
									rect.top	= rt.GetTop();
									rect.right	= rt.GetRight();
									rect.bottom = rt.GetBottom();
								}
							}

			void			Draw(Gdiplus::Graphics* graphics, const Gdiplus::Pen* pen)
							{
								if (graphics == NULL || pen == NULL || m_Object == NULL)
								{
									return;
								}

								if (!m_Region)
								{
									graphics->DrawPath(pen, (Gdiplus::GraphicsPath*)m_Object);
								}
							}

			void			Fill(Gdiplus::Graphics* graphics, const Gdiplus::Brush* brush)
							{
								if (graphics == NULL || brush == NULL || m_Object == NULL)
								{
									return;
								}

								if (m_Region)
								{
									graphics->FillRegion(brush, (Gdiplus::Region*)m_Object);
								}
								else
								{
									graphics->FillPath(brush, (Gdiplus::GraphicsPath*)m_Object);
								}
							}

private:
			LPVOID			m_Object;
			bool			m_Region;
};


class CGPModifyBrushOpacity
{
public:
							CGPModifyBrushOpacity(const CBCGPBrush& brush, CBCGPGraphicsManager* pGM)
								: m_Brush		  ((CBCGPBrush&)brush)
								, m_CurrentOpacity(pGM->GetCurrentOpacity())
								, m_SavedOpacity  (brush.GetOpacity())
							{
								if (m_CurrentOpacity >= 0)
								{
									m_Brush.SetOpacity(m_CurrentOpacity * m_SavedOpacity);
								}
							}

							~CGPModifyBrushOpacity()
							{
								if (m_CurrentOpacity >= 0)
								{
									m_Brush.SetOpacity(m_SavedOpacity);
								}
							}

protected:
							CBCGPBrush& m_Brush;
							double		m_CurrentOpacity;
							double		m_SavedOpacity;
};


static Gdiplus::GraphicsPath* GPCreatePath(const CBCGPGraphicsManagerGdiPlus* pGM, const CBCGPGeometry& geometry)
{
	Gdiplus::GraphicsPath* pPath = CBCGPGdiPlusHelper::CreateGraphicsPath();
	if (pPath == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}

	if (geometry.IsKindOf(RUNTIME_CLASS(CBCGPRectangleGeometry)))
	{
		pPath->AddRectangle(GPRect(pGM, ((CBCGPRectangleGeometry&)geometry).GetRectangle()));
	}
	else if (geometry.IsKindOf(RUNTIME_CLASS(CBCGPEllipseGeometry)))
	{
		pPath->AddEllipse(GPRect(pGM, ((CBCGPEllipseGeometry&)geometry).GetEllipse()));
	}
	else if (geometry.IsKindOf(RUNTIME_CLASS(CBCGPRoundedRectangleGeometry)))
	{
		const CBCGPRoundedRect& rect = ((CBCGPRoundedRectangleGeometry&)geometry).GetRoundedRect();

		Gdiplus::GraphicsPath* pRoundedPath = CBCGPGdiPlusHelper::CreateRoundedRectanglePath
				(
					GPRect(pGM, rect.rect), 
					(Gdiplus::REAL)rect.radiusX, 
					(Gdiplus::REAL)rect.radiusY
				);
		if (pRoundedPath != NULL)
		{
			pPath->AddPath(pRoundedPath, FALSE);
		}
	}
	else if (geometry.IsKindOf(RUNTIME_CLASS(CBCGPPolygonGeometry)))
	{
		CBCGPPolygonGeometry& polygonGeometry = (CBCGPPolygonGeometry&)geometry;
		const CBCGPPointsArray& arPoints = polygonGeometry.GetPoints();
		const int nCount = (int)arPoints.GetSize();

		pPath->SetFillMode((geometry.GetFillMode() == CBCGPGeometry::BCGP_FILL_MODE_ALTERNATE)
									? Gdiplus::FillModeAlternate
									: Gdiplus::FillModeWinding);

		if (polygonGeometry.GetCurveType() == CBCGPPolygonGeometry::BCGP_CURVE_TYPE_BEZIER)
		{
			if (nCount > 0)
			{
				Gdiplus::PointF ptCurr = GPPoint(pGM, arPoints[0]);

				for (int i = 1; i < nCount; i++)
				{
					Gdiplus::PointF points[3];

					points[0] = GPPoint(pGM, arPoints[i]);
					if (i < nCount - 1)
					{
						i++;
					}

					points[1] = GPPoint(pGM, arPoints[i]);
					if (i < nCount - 1)
					{
						i++;
					}

					points[2] = GPPoint(pGM, arPoints[i]);

					pPath->AddBezier(ptCurr, points[0], points[1], points[2]);
					ptCurr = points[2];
				}
			}
		}
		else
		{
			const CBCGPPoint* points = arPoints.GetData();

			Gdiplus::PointF* pts = CBCGPGdiPlusHelper::CreateArrayPointF(nCount);
			Gdiplus::PointF* ppts = pts;

			for (int i = 0; i < nCount; i++)
			{
				ppts->X = (Gdiplus::REAL)points->x;
				ppts->Y = (Gdiplus::REAL)points->y;
				ppts++;
				points++;
			}

			pPath->AddLines(pts, nCount);

			CBCGPGdiPlusHelper::DeleteArrayPointF(pts);
		}

		if (polygonGeometry.IsClosed())
		{
			pPath->CloseFigure();
		}
	}
	else if (geometry.IsKindOf(RUNTIME_CLASS(CBCGPSplineGeometry)))
	{
		CBCGPSplineGeometry& splineGeometry = (CBCGPSplineGeometry&)geometry;
		const CBCGPPointsArray& arPoints = splineGeometry.GetPoints();
		int nCount = (int)arPoints.GetSize();

		nCount = (nCount - 1) / 3;
		if (nCount > 0)
		{
			pPath->SetFillMode((geometry.GetFillMode() == CBCGPGeometry::BCGP_FILL_MODE_ALTERNATE)
										? Gdiplus::FillModeAlternate
										: Gdiplus::FillModeWinding);

			pPath->StartFigure();

			CBCGPPoint ptCurr = arPoints[0];

			int index = 1;
			for (int i = 0; i < nCount; i++)
			{
				pPath->AddBezier(GPPoint(pGM, ptCurr), GPPoint(pGM, arPoints[index]),
						GPPoint(pGM, arPoints[index + 1]), GPPoint(pGM, arPoints[index + 2]));

				ptCurr = arPoints[index + 2];
				index += 3;
			}

			if (splineGeometry.IsClosed())
			{
				pPath->CloseFigure();
			}
		}
	}
	else if (geometry.IsKindOf(RUNTIME_CLASS(CBCGPComplexGeometry)))
	{
		CBCGPComplexGeometry& complexGeometry = (CBCGPComplexGeometry&)geometry;
		const CObArray& arSegments = complexGeometry.GetSegments();
		int nCount = (int)arSegments.GetSize();

		CBCGPPoint ptCurr = complexGeometry.GetStartPoint();

		pPath->SetFillMode((geometry.GetFillMode() == CBCGPGeometry::BCGP_FILL_MODE_ALTERNATE)
									? Gdiplus::FillModeAlternate
									: Gdiplus::FillModeWinding);

		pPath->StartFigure();

		for (int i = 0; i < nCount; i++)
		{
			CObject* pSegment = arSegments[i];
			ASSERT_VALID(pSegment);

			if (pSegment->IsKindOf(RUNTIME_CLASS(CBCGPLineSegment)))
			{
				CBCGPLineSegment* pLineSegment = (CBCGPLineSegment*)pSegment;

				pPath->AddLine(GPPoint(pGM, ptCurr), GPPoint(pGM, pLineSegment->m_Point));

				ptCurr = pLineSegment->m_Point;
			}
			else if (pSegment->IsKindOf(RUNTIME_CLASS(CBCGPBezierSegment)))
			{
				CBCGPBezierSegment* pBezierSegment = (CBCGPBezierSegment*)pSegment;

				pPath->AddBezier(GPPoint(pGM, ptCurr), GPPoint(pGM, pBezierSegment->m_Point1),
					GPPoint(pGM, pBezierSegment->m_Point2), GPPoint(pGM, pBezierSegment->m_Point3));

				ptCurr = pBezierSegment->m_Point3;
			}
			else if (pSegment->IsKindOf(RUNTIME_CLASS(CBCGPArcSegment)))
			{
				//ASSERT(FALSE);
				CBCGPArcSegment* pArcSegment = (CBCGPArcSegment*)pSegment;

				Gdiplus::GraphicsPath* pPathArc = NULL;
				if (pArcSegment->m_dblRotationAngle != 0.0)
				{
					pPathArc = CBCGPGdiPlusHelper::CreateGraphicsPath();
				}

				BOOL bIsLargeArc = pArcSegment->m_bIsLargeArc;

				double rX;
				double rY;

				CBCGPPoint ptCenter = pArcSegment->GetArcCenter(ptCurr, bIsLargeArc, rX, rY);
				if (rX == 0.0 || rY == 0.0)
				{
					break;
				}

				double startAngle  = bcg_normalize_deg(bcg_rad2deg(bcg_angle((ptCurr.x - ptCenter.x), (ptCenter.y - ptCurr.y))));
				double finishAngle = bcg_normalize_deg(bcg_rad2deg(bcg_angle((pArcSegment->m_Point.x - ptCenter.x), (ptCenter.y - pArcSegment->m_Point.y))));
				double sweepAngle  = finishAngle - startAngle;

				if (pArcSegment->m_bIsClockwise)
				{
					if (sweepAngle > 0.0)
					{
						sweepAngle -= 360.0;
					}
				}
				else
				{
					if (sweepAngle < 0.0)
					{
						sweepAngle += 360.0;
					}
				}

				if (pPathArc != NULL)
				{
					startAngle += pArcSegment->m_dblRotationAngle;

					pPathArc->AddArc
							(
								Gdiplus::RectF((Gdiplus::REAL)(-rX), (Gdiplus::REAL)(-rY), (Gdiplus::REAL)(rX * 2.0), (Gdiplus::REAL)(rY * 2.0)),
								(Gdiplus::REAL)(360.0 - startAngle),
								-(Gdiplus::REAL)sweepAngle
							);
					Gdiplus::Matrix m;
					m.Rotate((Gdiplus::REAL)pArcSegment->m_dblRotationAngle);
					m.Translate((Gdiplus::REAL)(ptCenter.x), (Gdiplus::REAL)(ptCenter.y), Gdiplus::MatrixOrderAppend);
					pPathArc->Transform(&m);

					pPath->AddPath(pPathArc, TRUE);
				}
				else
				{
					pPath->AddArc
							(
								Gdiplus::RectF((Gdiplus::REAL)(ptCenter.x - rX), (Gdiplus::REAL)(ptCenter.y - rY), (Gdiplus::REAL)(rX * 2.0), (Gdiplus::REAL)(rY * 2.0)),
								(Gdiplus::REAL)(360.0 - startAngle),
								-(Gdiplus::REAL)sweepAngle
							);
				}

				ptCurr = pArcSegment->m_Point;
			}
		}

		if (complexGeometry.IsClosed())
		{
			pPath->CloseFigure();
		}
	}
	else if (geometry.IsKindOf(RUNTIME_CLASS(CBCGPGeometryGroup)))
	{
		pPath->SetFillMode((geometry.GetFillMode() == CBCGPGeometry::BCGP_FILL_MODE_ALTERNATE)
									? Gdiplus::FillModeAlternate
									: Gdiplus::FillModeWinding);

		const CArray<CBCGPGeometry*, CBCGPGeometry*>& arGeometries = ((CBCGPGeometryGroup&)geometry).GetGeometries();

		for (int i = 0; i < (int)arGeometries.GetSize(); i++)
		{
			const CBCGPGeometry* pGeometry = arGeometries[i];
			if (pGeometry == NULL)
			{
				continue;
			}

			Gdiplus::GraphicsPath* p = GPCreatePath(pGM, *pGeometry);

			if (p != NULL)
			{
				pPath->AddPath(p, FALSE);
			}
		}
	}

	return pPath;
}

#endif


BOOL CBCGPGraphicsManagerGdiPlus::m_bCheckLineOffsets = TRUE;
BOOL CBCGPGraphicsManagerGdiPlus::m_bRectOffsets = TRUE;

IMPLEMENT_DYNCREATE(CBCGPGraphicsManagerGdiPlus, CBCGPGraphicsManager)


CBCGPGraphicsManagerGdiPlus::CBCGPGraphicsManagerGdiPlus(CDC* pDC, BOOL bDoubleBuffering, CBCGPGraphicsManagerParams* /*pParams*/)
	: m_pDCPaint		  (NULL)
	, m_pMemDC			  (NULL)
	, m_pBmpScreenOld	  (NULL)
	, m_pGraphics		  (NULL)
	, m_sizeScaleTransform(1.0, 1.0)
{
	m_Type = BCGP_GRAPHICS_MANAGER_GDI_PLUS;

	m_nSupportedFeatures = BCGP_GRAPHICS_MANAGER_COLOR_OPACITY |
						   BCGP_GRAPHICS_MANAGER_CAP_STYLE |
						   BCGP_GRAPHICS_MANAGER_LINE_JOIN |
						   BCGP_GRAPHICS_MANAGER_ANTIALIAS |
						   BCGP_GRAPHICS_MANAGER_LOCALE |
						   BCGP_GRAPHICS_MANAGER_SCALING;

	InitGDIPlus();
	BindDC(pDC, bDoubleBuffering);

	if (m_pDC != NULL)
	{
		BeginDraw();
	}
}


CBCGPGraphicsManagerGdiPlus::CBCGPGraphicsManagerGdiPlus(const CBCGPRect& rectDest, CBCGPImage* pImageDest)
	: CBCGPGraphicsManager(rectDest, pImageDest)
	, m_pDCPaint		  (NULL)
	, m_pMemDC			  (NULL)
	, m_pBmpScreenOld	  (NULL)
	, m_pGraphics		  (NULL)
	, m_sizeScaleTransform(1.0, 1.0)
{
	m_Type = BCGP_GRAPHICS_MANAGER_GDI_PLUS;

	m_nSupportedFeatures = BCGP_GRAPHICS_MANAGER_COLOR_OPACITY |
						   BCGP_GRAPHICS_MANAGER_CAP_STYLE |
						   BCGP_GRAPHICS_MANAGER_LINE_JOIN |
						   BCGP_GRAPHICS_MANAGER_ANTIALIAS |
						   BCGP_GRAPHICS_MANAGER_LOCALE |
						   BCGP_GRAPHICS_MANAGER_SCALING;
}


CBCGPGraphicsManagerGdiPlus::~CBCGPGraphicsManagerGdiPlus()
{
	EndDraw();

#ifndef BCGP_EXCLUDE_GDI_PLUS

	while (!m_lstTransforms.IsEmpty())
	{
		CBCGPGdiPlusHelper::DeleteMatrix((Gdiplus::Matrix*)m_lstTransforms.RemoveTail());
	}

#endif

	CleanResources();
}


CBCGPGraphicsManager*
CBCGPGraphicsManagerGdiPlus::CreateOffScreenManager(const CBCGPRect& rect, CBCGPImage* pImageDest)
{
	if (m_pDC == NULL)
	{
		return NULL;
	}

	if (rect.IsRectEmpty())
	{
		return NULL;
	}

	CBCGPGraphicsManagerGdiPlus* pGM = NULL;

#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(pImageDest);

#else

	pGM = new CBCGPGraphicsManagerGdiPlus(rect, pImageDest);

	pGM->m_nSupportedFeatures = m_nSupportedFeatures;
	pGM->EnableAntialiasing(IsAntialiasingEnabled());
	pGM->EnableTransparentGradient(IsTransparentGradient());
	pGM->m_pOriginal = this;

	if (!pGM->BeginDraw())
	{
		delete pGM;
		pGM = NULL;

		TRACE0("CBCGPGraphicsManagerGDIPlus::CreateOffScreenManager failed.\r\n");
	}

#endif

	return pGM;
}


void
CBCGPGraphicsManagerGdiPlus::BindDC(CDC* pDC, BOOL bDoubleBuffering)
{
	if (m_pImageDest != NULL)
	{
		ASSERT(FALSE);
		return;
	}

	m_pDC = pDC;
	m_bDoubleBuffering = m_pDC != NULL && bDoubleBuffering && m_pDC->GetWindow()->GetSafeHwnd() != NULL && !pDC->IsPrinting();
}


void
CBCGPGraphicsManagerGdiPlus::BindDC(CDC* pDC, const CRect& /*rect*/)
{
	BindDC(pDC);
}


BOOL
CBCGPGraphicsManagerGdiPlus::BeginDraw()
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	return FALSE;

#else

	if (m_pGraphics != NULL)
	{
		return FALSE;
	}

	if (m_pImageDest != NULL)
	{
		CSize sizeDest = m_rectDest.Size();
		if (sizeDest.cx <= 0 || sizeDest.cy <= 0)
		{
			return FALSE;
		}

		m_dcScreen.CreateCompatibleDC(NULL);

		HBITMAP hBitmap = CBCGPDrawManager::CreateBitmap_32(sizeDest, NULL);
		if (hBitmap == NULL)
		{
			return FALSE;
		}

		m_bmpScreen.Attach(hBitmap);
		m_pBmpScreenOld = m_dcScreen.SelectObject(&m_bmpScreen);

		m_dcScreen.FillSolidRect(CRect(0, 0, sizeDest.cx, sizeDest.cy), RGB(0, 0, 0));

		CBCGPGraphicsManagerGdiPlus* pGMOrig = (CBCGPGraphicsManagerGdiPlus*)m_pOriginal;
		if (pGMOrig != NULL && pGMOrig->m_pDCPaint != NULL)
		{
			m_dcScreen.BitBlt(0, 0, sizeDest.cx, sizeDest.cy, pGMOrig->m_pDCPaint, 
							(int)m_rectDest.left, (int)m_rectDest.top, SRCCOPY);
		}

		m_pDCPaint = &m_dcScreen;
	}
	else if (m_bDoubleBuffering)
	{
		ASSERT(m_pMemDC == NULL);

#ifndef _BCGSUITE_
		m_pMemDC = new CBCGPMemDC(*m_pDC, m_pDC->GetWindow(), 0, m_dblScale);
#else
		m_pMemDC = new CBCGPMemDC(*m_pDC, m_pDC->GetWindow());
#endif
		m_pDCPaint = &m_pMemDC->GetDC ();
	}
	else
	{
		m_pDCPaint = m_pDC;
	}

	m_pGraphics = Gdiplus::Graphics::FromHDC(m_pDCPaint->GetSafeHdc());

	GPGraphics(m_pGraphics)->SetPixelOffsetMode(Gdiplus::PixelOffsetModeHighQuality);
	//GPGraphics(m_pGraphics)->SetTextRenderingHint(Gdiplus::TextRenderingHintAntiAliasGridFit);
	GPGraphics(m_pGraphics)->SetPageUnit(Gdiplus::UnitPixel);

	EnableAntialiasing(TRUE);

	return TRUE;

#endif
}


void
CBCGPGraphicsManagerGdiPlus::EndDraw()
{
#ifndef BCGP_EXCLUDE_GDI_PLUS

	ReleaseClipArea();

	if (m_pGraphics != NULL)
	{
		CBCGPGdiPlusHelper::DeleteGraphics(GPGraphics(m_pGraphics));
		m_pGraphics = NULL;
	}

	if (m_pMemDC != NULL)
	{
		delete m_pMemDC;
		m_pMemDC = NULL;
	}

	if (m_pBmpScreenOld != NULL && m_dcScreen.GetSafeHdc() != NULL)
	{
		m_dcScreen.SelectObject(m_pBmpScreenOld);
		m_pBmpScreenOld = NULL;
	}

	if (m_bmpScreen.GetSafeHandle() != NULL && m_pImageDest != NULL)
	{
		m_dcScreen.DeleteDC();

		ASSERT_VALID(m_pImageDest);

		DestroyImage(*m_pImageDest);

		Gdiplus::Bitmap* pBitmap = Gdiplus::Bitmap::FromHBITMAP((HBITMAP)m_bmpScreen.GetSafeHandle(), NULL);

		if (pBitmap != NULL)
		{
			m_pImageDest->Set(this, (LPVOID)pBitmap);
		}

		m_bmpScreen.DeleteObject();
	}

	m_pDCPaint = NULL;

#endif
}


void CBCGPGraphicsManagerGdiPlus::Clear(const CBCGPColor& color)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(color);

#else

	if (m_pGraphics == NULL)
	{
		return;
	}

	GPGraphics(m_pGraphics)->Clear(color.IsNull() ? GPColor(CBCGPColor(globalData.clrWindow)) : GPColor(color));

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawLine(const CBCGPPoint& ptFrom, const CBCGPPoint& ptTo, const CBCGPBrush& brush, double lineWidth, const CBCGPStrokeStyle* pStrokeStyle)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(ptFrom);
	UNREFERENCED_PARAMETER(ptTo);
	UNREFERENCED_PARAMETER(brush);
	UNREFERENCED_PARAMETER(lineWidth);
	UNREFERENCED_PARAMETER(pStrokeStyle);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	CGPStoreTransform storeTransform(pBrush);
	PrepareBrush((CBCGPBrush&)brush, pBrush, CBCGPRect(ptFrom, ptTo));

	CGPPen pPen(pBrush, lineWidth, pStrokeStyle);
	if (pPen == NULL)
	{
		return;
	}

	BOOL bPixelOffset = m_bCheckLineOffsets ? (fabs(ptFrom.x - ptTo.x) < .1 || fabs(ptFrom.y - ptTo.y) < .1) : FALSE;

	GPGraphics(m_pGraphics)->DrawLine(pPen, GPPoint(this, ptFrom, bPixelOffset), GPPoint(this, ptTo, bPixelOffset));

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawLines(const CBCGPPointsArray& arPoints, const CBCGPBrush& brush, double lineWidth, const CBCGPStrokeStyle* pStrokeStyle)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(arPoints);
	UNREFERENCED_PARAMETER(brush);
	UNREFERENCED_PARAMETER(lineWidth);
	UNREFERENCED_PARAMETER(pStrokeStyle);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	const int nSize = (int)arPoints.GetSize();
	if (nSize < 2)
	{
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	//CGPStoreTransform storeTransform(pBrush);
	//PrepareBrush((CBCGPBrush&)brush, pBrush, CBCGPRect(ptFrom, ptTo));

	CGPPen pPen(pBrush, lineWidth, pStrokeStyle);
	if (pPen == NULL)
	{
		return;
	}

	for (int i = 0; i < nSize - 1; i++)
	{
		GPGraphics(m_pGraphics)->DrawLine(pPen, GPPoint(this, arPoints[i]), GPPoint(this, arPoints[i + 1]));
	}

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawScatter(const CBCGPPointsArray& arPoints, const CBCGPBrush& brush, double dblPointSize, UINT nStyle)
{
	UNREFERENCED_PARAMETER(nStyle);

#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(arPoints);
	UNREFERENCED_PARAMETER(brush);
	UNREFERENCED_PARAMETER(dblPointSize);

#else

	const int nSize = (int)arPoints.GetSize();
	if (nSize < 1)
	{
		return;
	}
	
	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	//CGPStoreTransform storeTransform(pBrush);
	//PrepareBrush((CBCGPBrush&)brush, pBrush, CBCGPRect(ptFrom, ptTo));

	CGPPen pPen(pBrush, max(1.0, 0.25 * dblPointSize), NULL);
	if (pPen == NULL)
	{
		return;
	}

	for (int i = 0; i < nSize; i++)
	{
		GPGraphics(m_pGraphics)->DrawLine
						(
							pPen,
							GPPoint(this, arPoints[i].x - dblPointSize, arPoints[i].y),
							GPPoint(this, arPoints[i].x + dblPointSize, arPoints[i].y)
						);

		GPGraphics(m_pGraphics)->DrawLine
						(
							pPen,
							GPPoint(this, arPoints[i].x, arPoints[i].y - dblPointSize),
							GPPoint(this, arPoints[i].x, arPoints[i].y + dblPointSize)
						);
	}

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawRectangle(const CBCGPRect& rect, const CBCGPBrush& brush, double lineWidth, const CBCGPStrokeStyle* pStrokeStyle)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(rect);
	UNREFERENCED_PARAMETER(brush);
	UNREFERENCED_PARAMETER(lineWidth);
	UNREFERENCED_PARAMETER(pStrokeStyle);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	CGPStoreTransform storeTransform(pBrush);
	PrepareBrush((CBCGPBrush&)brush, pBrush, rect);

	CGPPen pPen(pBrush, lineWidth, pStrokeStyle);
	if (pPen == NULL)
	{
		return;
	}

	GPGraphics(m_pGraphics)->DrawRectangle(pPen, GPRect(this, rect, m_bRectOffsets));

#endif
}


void
CBCGPGraphicsManagerGdiPlus::FillRectangle(const CBCGPRect& rect, const CBCGPBrush& brush)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(rect);
	UNREFERENCED_PARAMETER(brush);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	if (brush.GetGradientType() == CBCGPBrush::BCGP_GRADIENT_BEVEL)
	{
		DrawBeveledRectangle(rect, brush);
	}
	else
	{
		CGPStoreTransform storeTransform(pBrush);
		PrepareBrush((CBCGPBrush&)brush, pBrush, rect);

		GPGraphics(m_pGraphics)->FillRectangle(pBrush, GPRect(this, rect));
	}

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawRoundedRectangle(const CBCGPRoundedRect& rect, const CBCGPBrush& brush, double lineWidth, const CBCGPStrokeStyle* pStrokeStyle)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(rect);
	UNREFERENCED_PARAMETER(brush);
	UNREFERENCED_PARAMETER(lineWidth);
	UNREFERENCED_PARAMETER(pStrokeStyle);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	CGPStoreTransform storeTransform(pBrush);
	PrepareBrush((CBCGPBrush&)brush, pBrush, rect.rect);

	CGPPen pPen(pBrush, lineWidth, pStrokeStyle);
	if (pPen == NULL)
	{
		return;
	}

	Gdiplus::GraphicsPath* pGeometry = CBCGPGdiPlusHelper::CreateRoundedRectanglePath
				(
					GPRect(this, rect.rect), 
					(Gdiplus::REAL)rect.radiusX, 
					(Gdiplus::REAL)rect.radiusY
				);
	if (pGeometry == NULL)
	{
		return;
	}

	GPGraphics(m_pGraphics)->DrawPath(pPen, pGeometry);

	CBCGPGdiPlusHelper::DeleteGraphicsPath(pGeometry);

#endif
}


void
CBCGPGraphicsManagerGdiPlus::FillRoundedRectangle(const CBCGPRoundedRect& rect, const CBCGPBrush& brush)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(rect);
	UNREFERENCED_PARAMETER(brush);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	Gdiplus::GraphicsPath* pGeometry = CBCGPGdiPlusHelper::CreateRoundedRectanglePath
				(
					GPRect(this, rect.rect), 
					(Gdiplus::REAL)rect.radiusX, 
					(Gdiplus::REAL)rect.radiusY
				);
	if (pGeometry == NULL)
	{
		return;
	}

	CGPStoreTransform storeTransform(pBrush);
	PrepareBrush((CBCGPBrush&)brush, pBrush, rect.rect);

	GPGraphics(m_pGraphics)->FillPath(pBrush, pGeometry);

	CBCGPGdiPlusHelper::DeleteGraphicsPath(pGeometry);

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawEllipse(const CBCGPEllipse& ellipse, const CBCGPBrush& brush, double lineWidth, const CBCGPStrokeStyle* pStrokeStyle)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(ellipse);
	UNREFERENCED_PARAMETER(brush);
	UNREFERENCED_PARAMETER(lineWidth);
	UNREFERENCED_PARAMETER(pStrokeStyle);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	if (ellipse.radiusX <= 0 || ellipse.radiusY <= 0)
	{
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	CGPStoreTransform storeTransform(pBrush);
	PrepareBrush((CBCGPBrush&)brush, pBrush, CBCGPRect(ellipse));

	CGPPen pPen(pBrush, lineWidth, pStrokeStyle);
	if (pPen == NULL)
	{
		return;
	}

	GPGraphics(m_pGraphics)->DrawEllipse(pPen, GPRect(this, ellipse));

#endif
}


void
CBCGPGraphicsManagerGdiPlus::FillEllipse(const CBCGPEllipse& ellipse, const CBCGPBrush& brush)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(ellipse);
	UNREFERENCED_PARAMETER(brush);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	if (ellipse.radiusX <= 0 || ellipse.radiusY <= 0)
	{
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	if (brush.GetGradientType() == CBCGPBrush::BCGP_GRADIENT_BEVEL)
	{
		{
			Gdiplus::Brush* pBrush2 = (Gdiplus::Brush*)brush.GetHandle2();

			CGPStoreTransform storeTransform(pBrush2);
			PrepareBrush((CBCGPBrush&)brush, pBrush2, ellipse);
			GPGraphics(m_pGraphics)->FillEllipse(pBrush2, GPRect(this, ellipse));
		}

		double nDepth = min(ellipse.radiusX, ellipse.radiusY) * .05;

		CBCGPEllipse ellipseInternal = ellipse;
		ellipseInternal.radiusX -= nDepth;
		ellipseInternal.radiusY -= nDepth;

		{
			CGPStoreTransform storeTransform(pBrush);
			PrepareBrush((CBCGPBrush&)brush, pBrush, ellipseInternal);
			GPGraphics(m_pGraphics)->FillEllipse(pBrush, GPRect(this, ellipseInternal));
		}
	}
	else
	{
		CGPStoreTransform storeTransform(pBrush);
		PrepareBrush((CBCGPBrush&)brush, pBrush, ellipse);

		GPGraphics(m_pGraphics)->FillEllipse(pBrush, GPRect(this, ellipse));
	}

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawArc(const CBCGPEllipse& ellipse, double dblStartAngle, double dblFinishAngle, BOOL bIsClockwise, const CBCGPBrush& brush, double lineWidth, const CBCGPStrokeStyle* pStrokeStyle)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(ellipse);
	UNREFERENCED_PARAMETER(dblStartAngle);
	UNREFERENCED_PARAMETER(dblFinishAngle);
	UNREFERENCED_PARAMETER(bIsClockwise);
	UNREFERENCED_PARAMETER(brush);
	UNREFERENCED_PARAMETER(lineWidth);
	UNREFERENCED_PARAMETER(pStrokeStyle);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	if (ellipse.IsNull())
	{
		return;
	}

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	CGPStoreTransform storeTransform(pBrush);
	PrepareBrush((CBCGPBrush&)brush, pBrush, ellipse);

	CGPPen pPen(pBrush, lineWidth, pStrokeStyle);
	if (pPen == NULL)
	{
		return;
	}

	dblStartAngle  = bcg_normalize_deg(dblStartAngle);
	dblFinishAngle = bcg_normalize_deg(dblFinishAngle);
	double sweepAngle = dblFinishAngle - dblStartAngle;

	if (bIsClockwise)
	{
		if (sweepAngle > 0.0)
		{
			sweepAngle -= 360.0;
		}
	}
	else
	{
		if (sweepAngle < 0.0)
		{
			sweepAngle += 360.0;
		}
	}

	GPGraphics(m_pGraphics)->DrawArc(pPen, GPRect(this, ellipse), (Gdiplus::REAL)(360.0 - dblStartAngle), (Gdiplus::REAL)-sweepAngle);

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawArc(const CBCGPPoint& ptFrom, const CBCGPPoint& ptTo, const CBCGPSize sizeRadius, BOOL bIsClockwise, BOOL bIsLargeArc, const CBCGPBrush& brush, double lineWidth, const CBCGPStrokeStyle* pStrokeStyle)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(ptFrom);
	UNREFERENCED_PARAMETER(ptTo);
	UNREFERENCED_PARAMETER(sizeRadius);
	UNREFERENCED_PARAMETER(bIsClockwise);
	UNREFERENCED_PARAMETER(bIsLargeArc);
	UNREFERENCED_PARAMETER(brush);
	UNREFERENCED_PARAMETER(lineWidth);
	UNREFERENCED_PARAMETER(pStrokeStyle);

#else

	if (m_pGraphics == NULL)
	{
		return;
	}

	CBCGPArcSegment arcSegment(ptTo, sizeRadius, bIsClockwise, bIsLargeArc, 0);

	double rX = 0;
	double rY = 0;

	CBCGPPoint ptCenter = arcSegment.GetArcCenter(ptFrom, bIsLargeArc, rX, rY);
	if (rX == 0.0 || rY == 0.0)
	{
		return;
	}

	DrawArc
		(
			CBCGPEllipse(ptCenter, rX, rY), 
			bcg_rad2deg(bcg_angle((ptFrom.x - ptCenter.x), (ptCenter.y - ptFrom.y))),
			bcg_rad2deg(bcg_angle((ptTo.x - ptCenter.x), (ptCenter.y - ptTo.y))), 
			arcSegment.m_bIsClockwise, 
			brush, 
			lineWidth, 
			pStrokeStyle
		);

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawGeometry(const CBCGPGeometry& geometry, const CBCGPBrush& brush, double lineWidth, const CBCGPStrokeStyle* pStrokeStyle)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(geometry);
	UNREFERENCED_PARAMETER(brush);
	UNREFERENCED_PARAMETER(lineWidth);
	UNREFERENCED_PARAMETER(pStrokeStyle);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	CGPGeometry* pGeometry = (CGPGeometry*)CreateGeometry((CBCGPGeometry&)geometry);
	if (pGeometry == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	CGPStoreTransform storeTransform(pBrush);
	if (NeedPrepareBrush(brush))
	{
		CBCGPRect rectGradient = m_rectCurrGradient;
		if (rectGradient.IsRectEmpty())
		{
			GetGeometryBoundingRect(geometry, rectGradient);
		}

		PrepareBrush((CBCGPBrush&)brush, pBrush, rectGradient);
	}

	CGPPen pPen(pBrush, lineWidth, pStrokeStyle);
	if (pPen == NULL)
	{
		return;
	}

	pGeometry->Draw(GPGraphics(m_pGraphics), pPen);

	if (IsDrawShadowMode())
	{
		DestroyGeometry((CBCGPGeometry&)geometry);
	}

#endif
}


void
CBCGPGraphicsManagerGdiPlus::FillGeometry(const CBCGPGeometry& geometry, const CBCGPBrush& brush)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(geometry);
	UNREFERENCED_PARAMETER(brush);

#else

	if (m_pGraphics == NULL || brush.IsEmpty())
	{
		return;
	}

	CGPModifyBrushOpacity mo(brush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)brush);
	if (pBrush == NULL)
	{
		return;
	}

	CGPGeometry* pGeometry = (CGPGeometry*)CreateGeometry((CBCGPGeometry&)geometry);
	if (pGeometry == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	CGPStoreTransform storeTransform(pBrush);
	if (NeedPrepareBrush(brush))
	{
		CBCGPRect rectGradient = m_rectCurrGradient;
		if (rectGradient.IsRectEmpty())
		{
			GetGeometryBoundingRect(geometry, rectGradient);
		}

		PrepareBrush((CBCGPBrush&)brush, pBrush, rectGradient);
	}

	pGeometry->Fill(GPGraphics(m_pGraphics), pBrush);

	if (IsDrawShadowMode())
	{
		DestroyGeometry((CBCGPGeometry&)geometry);
	}

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawImage(const CBCGPImage& image, const CBCGPPoint& ptDest, const CBCGPSize& sizeDest, double opacity, CBCGPImage::BCGP_IMAGE_INTERPOLATION_MODE interpolationMode, const CBCGPRect& rectSrc)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(image);
	UNREFERENCED_PARAMETER(ptDest);
	UNREFERENCED_PARAMETER(sizeDest);
	UNREFERENCED_PARAMETER(opacity);
	UNREFERENCED_PARAMETER(interpolationMode);
	UNREFERENCED_PARAMETER(rectSrc);

#else

	if (m_pGraphics == NULL)
	{
		return;
	}

	if (sizeDest.cx < 0. || sizeDest.cy < 0)
	{
		return;
	}

	Gdiplus::Bitmap* pBitmap = (Gdiplus::Bitmap*)CreateImage((CBCGPImage&)image);
	if (pBitmap == NULL)
	{
		return;
	}

	CBCGPSize size = image.GetSize();
	CBCGPRect rectImage(CBCGPPoint(0.0, 0.0), size);
	if (!rectSrc.IsRectEmpty())
	{
		if (!rectImage.IntersectRect(rectImage, rectSrc))
		{
			return;
		}

		size = rectImage.Size();
	}

	if (!sizeDest.IsEmpty())
	{
		size = sizeDest;
	}
	CBCGPRect rectDest(ptDest, size);

	Gdiplus::InterpolationMode gpModeInterpolation = GPGraphics(m_pGraphics)->GetInterpolationMode();
	GPGraphics(m_pGraphics)->SetInterpolationMode(interpolationMode == CBCGPImage::BCGP_IMAGE_INTERPOLATION_MODE_LINEAR
										? Gdiplus::InterpolationModeHighQualityBilinear
										: Gdiplus::InterpolationModeNearestNeighbor);

	Gdiplus::CompositingMode gpModeCompositing = GPGraphics(m_pGraphics)->GetCompositingMode();
	if (image.IsIgnoreAlphaBitmap())
	{
		GPGraphics(m_pGraphics)->SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
	}

	Gdiplus::ImageAttributes* attr = NULL;

	if (GetCurrentOpacity() != -1.0)
	{
		opacity *= GetCurrentOpacity();
	}
	opacity = bcg_clamp(opacity, 0.0, 1.0);

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

	GPGraphics(m_pGraphics)->DrawImage(pBitmap, GPRect(this, rectDest), 
		(Gdiplus::REAL)rectImage.left,
		(Gdiplus::REAL)rectImage.top,
		(Gdiplus::REAL)rectImage.Width(),
		(Gdiplus::REAL)rectImage.Height(),
		Gdiplus::UnitPixel, attr, NULL, NULL);

	GPGraphics(m_pGraphics)->SetInterpolationMode(gpModeInterpolation);
	GPGraphics(m_pGraphics)->SetCompositingMode(gpModeCompositing);

	CBCGPGdiPlusHelper::DeleteImageAttributes(attr);

#endif
}


void
CBCGPGraphicsManagerGdiPlus::DrawText(const CString& strText, const CBCGPRect& rectText, const CBCGPTextFormat& textFormat, const CBCGPBrush& foregroundBrush)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(strText);
	UNREFERENCED_PARAMETER(rectText);
	UNREFERENCED_PARAMETER(textFormat);
	UNREFERENCED_PARAMETER(foregroundBrush);

#else

	if (m_pGraphics == NULL || strText.IsEmpty())
	{
		return;
	}

	CGPModifyBrushOpacity mo(foregroundBrush, this);

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)CreateBrush((CBCGPBrush&)foregroundBrush);
	if (pBrush == NULL)
	{
		return;
	}

	Gdiplus::StringFormat* pTextFormat = (Gdiplus::StringFormat*)CreateTextFormat((CBCGPTextFormat&)textFormat);
	if (pTextFormat == NULL)
	{
		return;
	}

	Gdiplus::Font* pFont = (Gdiplus::Font*)textFormat.GetHandle1();
	if (pFont == NULL)
	{
		return;
	}

	PrepareTextFormat((CBCGPTextFormat&)textFormat);

	double centerX = rectText.CenterPoint().x;
	double centerY = rectText.CenterPoint().y;

	switch (textFormat.GetTextAlignment())
	{
	case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_TRAILING:
		centerX = rectText.right;
		break;
	case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_LEADING:
		centerX = rectText.left;
		break;
	}

	switch (textFormat.GetTextVerticalAlignment())
	{
	case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_TRAILING:
		centerY = rectText.bottom;
		break;
		
	case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_LEADING:
		centerY = rectText.top;
		break;
	}

	CGPStoreTransform storeTransformBrush(pBrush);
	PrepareBrush((CBCGPBrush&)foregroundBrush, pBrush, rectText);

	CBCGPRect rt(rectText);

	if (textFormat.IsAdjustLineSpacing())
	{
		Gdiplus::FontFamily family;
		pFont->GetFamily(&family);

		const UINT nStyle = pFont->GetStyle();
		const UINT16 nEmHeight = family.GetEmHeight(nStyle);

		if (nEmHeight != 0)
		{
			const double offset = (GPGraphics(m_pGraphics)->GetDpiY() * pFont->GetSize() * family.GetCellDescent(nStyle)) / (72.0f * nEmHeight);
			rt.OffsetRect(0.0, (double)(int)(offset + 0.5));
		}
	}

	CGPStoreTransform storeTransformGraphics(GPGraphics(m_pGraphics));
	if (textFormat.GetDrawingAngle() != 0.0)
	{
		Gdiplus::Matrix matrix;
		matrix.RotateAt
				(
					-(Gdiplus::REAL)textFormat.GetDrawingAngle(), 
					GPPoint(this, centerX, centerY), 
					Gdiplus::MatrixOrderAppend
				);

		GPGraphics(m_pGraphics)->ResetTransform();
		GPGraphics(m_pGraphics)->SetTransform(&matrix);

		Gdiplus::PointF pts[4];
		pts[0] = GPPoint(this, rt.TopLeft());
		pts[1] = GPPoint(this, rt.TopRight());
		pts[2] = GPPoint(this, rt.BottomRight());
		pts[3] = GPPoint(this, rt.BottomLeft());

		matrix.Reset();
		matrix.RotateAt
				(
					(Gdiplus::REAL)textFormat.GetDrawingAngle(), 
					GPPoint(this, centerX, centerY), 
					Gdiplus::MatrixOrderAppend
				);
		matrix.TransformPoints(pts, 4);

		rt.left = min(min(min(pts[0].X, pts[1].X), pts[2].X), pts[3].X);
		rt.right = max(max(max(pts[0].X, pts[1].X), pts[2].X), pts[3].X);
		rt.top = min(min(min(pts[0].Y, pts[1].Y), pts[2].Y), pts[3].Y);
		rt.bottom = max(max(max(pts[0].Y, pts[1].Y), pts[2].Y), pts[3].Y);
	}

	const int strLength = strText.GetLength();
	LPWSTR lpText = NULL;

#ifdef _UNICODE
	lpText = (LPWSTR)(LPCWSTR)strText;
#else
	lpText = (LPWSTR)new WCHAR[strLength + 1];
	lpText = AfxA2WHelper(lpText, (LPCSTR)strText, strLength + 1);
#endif

	GPGraphics(m_pGraphics)->DrawString
		(
			lpText,
			strLength,
			pFont,
			GPRect(this, rt),
			pTextFormat,
			pBrush
		);

	if ((LPVOID)lpText != (LPVOID)(LPCTSTR)strText)
	{
		delete [] lpText;
	}

#endif
}


CBCGPSize
CBCGPGraphicsManagerGdiPlus::GetTextSize(const CString& strText, const CBCGPTextFormat& textFormat, double dblWidth, BOOL bIgnoreTextRotation)
{
	CBCGPSize size;

#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(strText);
	UNREFERENCED_PARAMETER(textFormat);
	UNREFERENCED_PARAMETER(dblWidth);
	UNREFERENCED_PARAMETER(bIgnoreTextRotation);

#else

	if (m_pGraphics == NULL || strText.IsEmpty())
	{
		return size;
	}

	if (textFormat.IsWordWrap() && dblWidth == 0.)
	{
		ASSERT(FALSE);
		return size;
	}

	Gdiplus::StringFormat* pTextFormat = (Gdiplus::StringFormat*)CreateTextFormat((CBCGPTextFormat&)textFormat);
	if (pTextFormat == NULL)
	{
		return size;
	}

	Gdiplus::Font* pFont = (Gdiplus::Font*)textFormat.GetHandle1();
	if (pFont == NULL)
	{
		return size;
	}

	PrepareTextFormat((CBCGPTextFormat&)textFormat);

	Gdiplus::SizeF sizeF((Gdiplus::REAL)0.0, (Gdiplus::REAL)0.0);

	const int strLength = strText.GetLength();
	LPWSTR lpText = NULL;

#ifdef _UNICODE
	lpText = (LPWSTR)(LPCWSTR)strText;
#else
	lpText = (LPWSTR)new WCHAR[strLength + 1];
	lpText = AfxA2WHelper(lpText, (LPCSTR)strText, strLength + 1);
#endif

	GPGraphics(m_pGraphics)->MeasureString
		(
			lpText,
			strLength,
			pFont,
			Gdiplus::SizeF(textFormat.IsWordWrap() ? (Gdiplus::REAL)dblWidth : REAL_MAX, REAL_MAX),
			pTextFormat,
			&sizeF
		);

	if ((LPVOID)lpText != (LPVOID)(LPCTSTR)strText)
	{
		delete [] lpText;
	}

	size.cx = (double)sizeF.Width;
	size.cy = (double)sizeF.Height;

	if (!bIgnoreTextRotation)
	{
		CBCGPTextFormat::AdjustTextSize(textFormat.GetDrawingAngle(), size);
	}

#endif

	return size;
}


void
CBCGPGraphicsManagerGdiPlus::SetClipArea(const CBCGPGeometry& geometry, int nFlags)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(geometry);
	UNREFERENCED_PARAMETER(nFlags);

#else

	if (m_pGraphics == NULL)
	{
		return;
	}

	CGPGeometry* pGeometry = (CGPGeometry*)CreateGeometry((CBCGPGeometry&)geometry);
	if (pGeometry == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	Gdiplus::CombineMode mode = Gdiplus::CombineModeReplace;
	switch (nFlags)
	{
	case RGN_OR:
		mode = Gdiplus::CombineModeUnion;
		break;

	case RGN_AND:
		mode = Gdiplus::CombineModeIntersect;
		break;

	case RGN_XOR:
		mode = Gdiplus::CombineModeXor;
		break;

	case RGN_DIFF:
		mode = Gdiplus::CombineModeExclude;
		break;
	}

	if (pGeometry->IsRegion())
	{
		GPGraphics(m_pGraphics)->SetClip((Gdiplus::Region*)pGeometry->GetObject(), mode);
	}
	else
	{
		GPGraphics(m_pGraphics)->SetClip((Gdiplus::GraphicsPath*)pGeometry->GetObject(), mode);
	}

#endif
}


void CBCGPGraphicsManagerGdiPlus::ReleaseClipArea()
{
#ifndef BCGP_EXCLUDE_GDI_PLUS

	if (m_pGraphics == NULL)
	{
		return;
	}

	GPGraphics(m_pGraphics)->ResetClip();

#endif
}


void
CBCGPGraphicsManagerGdiPlus::CombineGeometry(CBCGPGeometry& geometryDest, const CBCGPGeometry& geometrySrc1, const CBCGPGeometry& geometrySrc2, int nFlags)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(geometryDest);
	UNREFERENCED_PARAMETER(geometrySrc1);
	UNREFERENCED_PARAMETER(geometrySrc2);
	UNREFERENCED_PARAMETER(nFlags);

#else

	CGPGeometry* pGeometrySrc1 = (CGPGeometry*)CreateGeometry((CBCGPGeometry&)geometrySrc1);
	ASSERT(pGeometrySrc1 != NULL);

	CGPGeometry* pGeometrySrc2 = (CGPGeometry*)CreateGeometry((CBCGPGeometry&)geometrySrc2);
	ASSERT(pGeometrySrc2 != NULL);

	CGPGeometry* region = new CGPGeometry(NULL, false);
	region->CreateRegion(*pGeometrySrc1);

	if (region->GetObject() != NULL)
	{
		region->Combine(*pGeometrySrc2, nFlags);
	}
	else
	{
		delete region;
		region = NULL;
	}

	DestroyGeometry(geometryDest);
	geometryDest.Set(this, (LPVOID)region);

#endif
}


void
CBCGPGraphicsManagerGdiPlus::GetGeometryBoundingRect(const CBCGPGeometry& geometry, CBCGPRect& rectOut)
{
	rectOut.SetRectEmpty();

#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(geometry);

#else

	CGPGeometry* pGeometry = (CGPGeometry*)CreateGeometry((CBCGPGeometry&)geometry);
	if (pGeometry == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	pGeometry->GetBounds(GPGraphics(m_pGraphics), rectOut);

#endif
}


LPVOID
CBCGPGraphicsManagerGdiPlus::CreateGeometry(CBCGPGeometry& geometry)
{
	if (m_pOriginal != NULL)
	{
		return m_pOriginal->CreateGeometry(geometry);
	}

	if (geometry.GetHandle() != NULL)
	{
		CBCGPGraphicsManager* pGM = geometry.GetGraphicsManager();

		if (!IsGraphicsManagerValid(pGM))
		{
			return NULL;
		}

		return geometry.GetHandle();
	}

#ifndef BCGP_EXCLUDE_GDI_PLUS

	Gdiplus::GraphicsPath* pGeometry = GPCreatePath(this, geometry);

	if (pGeometry != NULL)
	{
		geometry.Set(this, (LPVOID)(new CGPGeometry((LPVOID)pGeometry, false)));
	}

#endif

	return geometry.GetHandle();
}


BOOL
CBCGPGraphicsManagerGdiPlus::DestroyGeometry(CBCGPGeometry& geometry)
{
#ifndef BCGP_EXCLUDE_GDI_PLUS

	CGPGeometry* pGeometry = (CGPGeometry*)geometry.GetHandle();
	if (pGeometry != NULL)
	{
		delete pGeometry;
	}

#endif

	geometry.Set(NULL, NULL);

	return TRUE;
}


LPVOID
CBCGPGraphicsManagerGdiPlus::CreateTextFormat(CBCGPTextFormat& textFormat)
{
	if (m_pOriginal != NULL)
	{
		return m_pOriginal->CreateTextFormat(textFormat);
	}

	if (textFormat.GetHandle() != NULL && textFormat.GetHandle1() != NULL)
	{
		CBCGPGraphicsManager* pGM = textFormat.GetGraphicsManager();

		if (!IsGraphicsManagerValid(pGM))
		{
			return NULL;
		}

		return textFormat.GetHandle();
	}

#ifndef BCGP_EXCLUDE_GDI_PLUS

	if (textFormat.GetFontSize() == 0. && textFormat.GetFontFamily().IsEmpty())
	{
		LOGFONT lf;
		memset(&lf, 0, sizeof(lf));

		globalData.fontRegular.GetLogFont(&lf);

		textFormat.CreateFromLogFont(lf);
	}

	INT style = 0; //Gdiplus::FontStyleRegular;
	if (textFormat.GetFontWeight() >= FW_BOLD)
	{
		style |= Gdiplus::FontStyleBold;
	}
	if (textFormat.GetFontStyle() != CBCGPTextFormat::BCGP_FONT_STYLE_NORMAL)
	{
		style |= Gdiplus::FontStyleItalic;
	}

	Gdiplus::Font* pFont = CBCGPGdiPlusHelper::CreateFont
		(
			textFormat.GetFontFamily(),
			(Gdiplus::REAL)fabs(textFormat.GetFontSize()),
			style,
			Gdiplus::UnitWorld
		);
	if (pFont == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}

	INT flags = 0;
	if (!textFormat.IsWordWrap())
	{
		flags |= Gdiplus::StringFormatFlagsNoWrap;
	}
	if (!textFormat.IsClipText())
	{
		flags |= Gdiplus::StringFormatFlagsNoClip;
	}

	Gdiplus::StringFormat* pTextFormat = CBCGPGdiPlusHelper::CreateStringFormat(flags, LANG_NEUTRAL);
	if (pTextFormat == NULL)
	{
		CBCGPGdiPlusHelper::DeleteFont(pFont);

		ASSERT(FALSE);
		return NULL;
	}

	textFormat.Set(this, (LPVOID)pTextFormat, (LPVOID)pFont);

	PrepareTextFormat(textFormat);

#endif

	return textFormat.GetHandle();
}


void
CBCGPGraphicsManagerGdiPlus::PrepareTextFormat(CBCGPTextFormat& textFormat) const
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(textFormat);

#else

	Gdiplus::StringFormat* pTextFormat = (Gdiplus::StringFormat*)textFormat.GetHandle();
	if (pTextFormat == NULL)
	{
		return;
	}

	INT flags = pTextFormat->GetFormatFlags();
	if (textFormat.IsWordWrap())
	{
		flags &= ~Gdiplus::StringFormatFlagsNoWrap;
	}
	if (textFormat.IsClipText())
	{
		flags &= ~Gdiplus::StringFormatFlagsNoClip;
	}
	pTextFormat->SetFormatFlags(flags);

	switch (textFormat.GetTextAlignment())
	{
	case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_TRAILING:
		pTextFormat->SetAlignment(Gdiplus::StringAlignmentFar);
		break;
	case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_CENTER:
		pTextFormat->SetAlignment(Gdiplus::StringAlignmentCenter);
		break;
	case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_LEADING:
		pTextFormat->SetAlignment(Gdiplus::StringAlignmentNear);
		break;
	}

	switch (textFormat.GetTextVerticalAlignment())
	{
	case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_TRAILING:
		pTextFormat->SetLineAlignment(Gdiplus::StringAlignmentFar);
		break;
	case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_CENTER:
		pTextFormat->SetLineAlignment(Gdiplus::StringAlignmentCenter);
		break;
	case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_LEADING:
		pTextFormat->SetLineAlignment(Gdiplus::StringAlignmentNear);
		break;
	}

#endif
}


BOOL
CBCGPGraphicsManagerGdiPlus::DestroyTextFormat(CBCGPTextFormat& textFormat)
{
#ifndef BCGP_EXCLUDE_GDI_PLUS

	CBCGPGdiPlusHelper::DeleteStringFormat((Gdiplus::StringFormat*)textFormat.GetHandle());
	CBCGPGdiPlusHelper::DeleteFont((Gdiplus::Font*)textFormat.GetHandle1());

#endif

	textFormat.Set(NULL, NULL);
	return TRUE;
}


void
CBCGPGraphicsManagerGdiPlus::PrepareBrush(CBCGPBrush& brush, LPVOID pBrushIn, const CBCGPRect& rectIn) const
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(brush);
	UNREFERENCED_PARAMETER(pBrushIn);
	UNREFERENCED_PARAMETER(rectIn);

#else

	if (!NeedPrepareBrush(brush))
	{
		return;
	}

	if (IsDrawShadowMode())
	{
		return;
	}

	Gdiplus::Brush* pBrush = (Gdiplus::Brush*)pBrushIn;

	CBCGPRect rect = !m_rectCurrGradient.IsRectEmpty() ? m_rectCurrGradient : rectIn;

	if (brush.HasTextureImage())
	{
		Gdiplus::TextureBrush* pTextureBrush = (Gdiplus::TextureBrush*)pBrush;

		//pTextureBrush->ScaleTransform((Gdiplus::REAL)rect.Width(), (Gdiplus::REAL)rect.Height(), Gdiplus::MatrixOrderAppend);
		pTextureBrush->TranslateTransform((Gdiplus::REAL)rect.left, (Gdiplus::REAL)rect.top, Gdiplus::MatrixOrderAppend);

		return;
	}

	CBCGPPoint ptFirst;
	CBCGPPoint ptSecond;

	const double dr = M_SQRT2 * 0.5;
	CBCGPPoint ptRadius(rect.Width() * dr, rect.Height() * dr);
	CBCGPPoint ptOffset(rect.CenterPoint());
	CBCGPPoint ptCenter(DBL_MAX, DBL_MAX);

	if (brush.GetGradientStops().GetSize() > 0)
	{
		ptFirst.x = brush.GetGradientPoint1().x;
		ptFirst.y = brush.GetGradientPoint1().y;
		if (!brush.GetGradientPoint1().bIsAbsolute)
		{
			ptFirst.x = rect.left + rect.Width() * ptFirst.x / 100.0;
			ptFirst.y = rect.top + rect.Height() * ptFirst.y / 100.0;
		}

		ptSecond.x = brush.GetGradientPoint2().x;
		ptSecond.y = brush.GetGradientPoint2().y;
		if (!brush.GetGradientPoint2().bIsAbsolute)
		{
			ptSecond.x = rect.left + rect.Width() * ptSecond.x / 100.0;
			ptSecond.y = rect.top + rect.Height() * ptSecond.y / 100.0;
		}

		if (pBrush->GetType() == Gdiplus::BrushTypePathGradient)
		{
			ptRadius.x = brush.GetGradientRadius().x;
			ptRadius.y = brush.GetGradientRadius().y;

			if (!brush.GetGradientRadius().bIsAbsolute)
			{
				ptRadius.x *= rect.Width() / 100;
				ptRadius.y *= rect.Height() / 100;
			}

			if (ptRadius.x == 0)
			{
				ptRadius.x = rect.Width() / 2;
			}
			
			if (ptRadius.y == 0)
			{
				ptRadius.y = rect.Height() / 2;
			}

			ptOffset = ptFirst;

			ptCenter = ptSecond - ptFirst;
			ptCenter.x = ptCenter.x / ptRadius.x;
			ptCenter.y = ptCenter.y / ptRadius.y;
		}
	}
	else
	{
		CBCGPBrush::BCGP_GRADIENT_TYPE type = brush.GetGradientType();

		if (type == CBCGPBrush::BCGP_GRADIENT_HORIZONTAL ||
			type == CBCGPBrush::BCGP_GRADIENT_CENTER_HORIZONTAL ||
			type == CBCGPBrush::BCGP_GRADIENT_PIPE_HORIZONTAL)
		{
			ptFirst.x = ptSecond.x = rect.CenterPoint().x;
			ptFirst.y = rect.top;
			ptSecond.y = rect.bottom;
		}

		if (type == CBCGPBrush::BCGP_GRADIENT_VERTICAL ||
			type == CBCGPBrush::BCGP_GRADIENT_CENTER_VERTICAL ||
			type == CBCGPBrush::BCGP_GRADIENT_PIPE_VERTICAL)
		{
			ptFirst.y = ptSecond.y = rect.CenterPoint().y;
			ptFirst.x = rect.right;
			ptSecond.x = rect.left;
		}

		if (type == CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT)
		{
			ptFirst.x = rect.left;
			ptFirst.y = rect.top;
			ptSecond.x = rect.right;
			ptSecond.y = rect.bottom;
		}

		if (type == CBCGPBrush::BCGP_GRADIENT_DIAGONAL_RIGHT)
		{
			ptFirst.x = rect.right;
			ptFirst.y = rect.top;
			ptSecond.x = rect.left;
			ptSecond.y = rect.bottom;
		}
	}

	Gdiplus::Matrix transform;
	transform.Reset();

	CBCGPTransformsArray& arTransforms = (CBCGPTransformsArray&)brush.GetGradientTransforms();
	if (arTransforms.GetSize() > 0)
	{
		for (int i = 0; i < (int)arTransforms.GetSize(); i++)
		{
			const CBCGPTransformMatrix& t = arTransforms[i];

			CBCGPPoint ptRotateCenter = t.m_ptRotateCenter;
			if (ptRotateCenter.x == DBL_MAX && ptRotateCenter.y == DBL_MAX)
			{
				ptRotateCenter = rect.CenterPoint();
			}
			else if (fabs(ptRotateCenter.x) > 0.0 && fabs(ptRotateCenter.x) < 1.0 && 
					fabs(ptRotateCenter.y) > 0.0 && fabs(ptRotateCenter.y) < 1.0)
			{
				ptRotateCenter.x = rect.Width() * ptRotateCenter.x;
				ptRotateCenter.y = rect.Height() * ptRotateCenter.y;
			}

			if (ptRotateCenter != CBCGPPoint(-DBL_MAX, -DBL_MAX))
			{
				transform.Translate(-(Gdiplus::REAL)ptRotateCenter.x, -(Gdiplus::REAL)ptRotateCenter.y, Gdiplus::MatrixOrderAppend);
			}

			Gdiplus::Matrix m;
			GPMatrix(m, t.m_arMatrix, t.m_bCheckRelativeTranslate ? rect.Size() : CBCGPSize());

			transform.Multiply(&m, Gdiplus::MatrixOrderAppend);

			if (ptRotateCenter != CBCGPPoint(-DBL_MAX, -DBL_MAX))
			{
				transform.Translate((Gdiplus::REAL)ptRotateCenter.x, (Gdiplus::REAL)ptRotateCenter.y, Gdiplus::MatrixOrderAppend);
			}
		}
	}

	if (pBrush->GetType() == Gdiplus::BrushTypeLinearGradient)
	{
		Gdiplus::LinearGradientBrush* pLinearBrush = (Gdiplus::LinearGradientBrush*)pBrush;

		const Gdiplus::REAL scale = (Gdiplus::REAL)bcg_distance(ptFirst, ptSecond);

		pLinearBrush->RotateTransform((Gdiplus::REAL)bcg_rad2deg(bcg_angle(ptFirst, ptSecond)), Gdiplus::MatrixOrderAppend);
		pLinearBrush->ScaleTransform(scale, scale, Gdiplus::MatrixOrderAppend);
		pLinearBrush->TranslateTransform((Gdiplus::REAL)ptFirst.x, (Gdiplus::REAL)ptFirst.y, Gdiplus::MatrixOrderAppend);

		if (!transform.IsIdentity())
		{
			pLinearBrush->MultiplyTransform(&transform, Gdiplus::MatrixOrderAppend);
		}
	}
	else
	{
		Gdiplus::PathGradientBrush* pPathBrush = (Gdiplus::PathGradientBrush*)pBrush;

		if (ptCenter.x != DBL_MAX && ptCenter.y != DBL_MAX)
		{
			pPathBrush->SetCenterPoint(Gdiplus::PointF((Gdiplus::REAL)ptCenter.x, (Gdiplus::REAL)ptCenter.y));
		}

		Gdiplus::Matrix world;
		GPGraphics(m_pGraphics)->GetTransform(&world);

		Gdiplus::PointF points[3];
		points[0].X = (Gdiplus::REAL)0.0;
		points[0].Y = (Gdiplus::REAL)0.0;
		points[1].X = (Gdiplus::REAL)1.0;
		points[1].Y = (Gdiplus::REAL)0.0;
		points[2].X = (Gdiplus::REAL)0.0;
		points[2].Y = (Gdiplus::REAL)1.0;

		world.TransformPoints(points, 3);

		Gdiplus::REAL scaleX = (Gdiplus::REAL)sqrt(bcg_sqr(points[1].X - points[0].X) + bcg_sqr(points[1].Y - points[0].Y));
		Gdiplus::REAL scaleY = (Gdiplus::REAL)sqrt(bcg_sqr(points[2].X - points[0].X) + bcg_sqr(points[2].Y - points[0].Y));

		if (scaleX >= 1.0)
		{
			scaleX = 1.0;
		}
		if (scaleY >= 1.0)
		{
			scaleY = 1.0;
		}

		pPathBrush->ScaleTransform((Gdiplus::REAL)ptRadius.x, (Gdiplus::REAL)ptRadius.y, Gdiplus::MatrixOrderAppend);
		pPathBrush->TranslateTransform((Gdiplus::REAL)ptOffset.x * scaleX, (Gdiplus::REAL)ptOffset.y * scaleY, Gdiplus::MatrixOrderAppend);

		Gdiplus::REAL m[6] = {Gdiplus::REAL(0)};
		transform.GetElements(m);
		transform.SetElements(m[0], m[1], m[2], m[3], m[4] * scaleX, m[5] * scaleY);

		pPathBrush->MultiplyTransform(&transform, Gdiplus::MatrixOrderAppend);
	}

#endif
}


LPVOID
CBCGPGraphicsManagerGdiPlus::CreateBrush(CBCGPBrush& brushIn)
{
	if (m_pOriginal != NULL)
	{
		return m_pOriginal->CreateBrush(brushIn);
	}

	CBCGPBrush& brush = IsDrawShadowMode() ? m_brShadow : brushIn;
	if (brush.GetHandle() != NULL)
	{
		CBCGPGraphicsManager* pGM = brush.GetGraphicsManager();

		if (!IsGraphicsManagerValid(pGM))
		{
			return NULL;
		}
	
		return brush.GetHandle();
	}

#ifndef BCGP_EXCLUDE_GDI_PLUS

	if (brush.HasTextureImage())
	{
		Gdiplus::Image* pImage = (Gdiplus::Image*)CreateImage(brush.GetTextureImage());
		if (pImage != NULL)
		{
			Gdiplus::TextureBrush* textureBrush = 
				CBCGPGdiPlusHelper::CreateTextureBrush(pImage, Gdiplus::WrapModeTile, (Gdiplus::REAL)bcg_clamp(brush.GetOpacity(), 0.0, 1.0));

			if (textureBrush == NULL)
			{
				TRACE(_T("CBCGPGraphicsManagerGDIPlus::CreateBrush: CreateTextureBrush Failed.\r\n"));
				return NULL;
			}

			brush.Set(this, (LPVOID)textureBrush);
			return brush.GetHandle();
		}
	}

	CBCGPBrush::BCGP_GRADIENT_TYPE type = brush.GetGradientType();
	if (type == CBCGPBrush::BCGP_GRADIENT_BEVEL)
	{
		CBCGPColor colorLight, colorDark;
		PrepareBevelColors(brush.GetColor(), colorLight, colorDark);

		CBCGPBrush br1(brush.GetColor(), CBCGPColor::White, CBCGPBrush::BCGP_GRADIENT_HORIZONTAL, brush.GetOpacity());
		CreateBrush(br1);

		CBCGPBrush br2(colorDark, brush.GetColor(), CBCGPBrush::BCGP_GRADIENT_HORIZONTAL, brush.GetOpacity());
		CreateBrush(br2);

		CBCGPBrush br3(colorDark, colorLight, CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT, brush.GetOpacity());
		CreateBrush(br3);

		brush.Set(this, br1.Detach(), br2.Detach(), br3.Detach());

		return brush.GetHandle();
	}

	Gdiplus::Brush* pBrush = NULL;

	const double opacity = brush.GetOpacity();
	const CBCGPBrushGradientStops& stops = brush.GetGradientStops();
	int stopsC = (int)stops.GetSize();

	Gdiplus::Color color1 = GPColor(brush.GetGradientColor(), opacity);
	Gdiplus::Color color2 = GPColor(brush.GetColor(), opacity);

	switch (type)
	{
	case CBCGPBrush::BCGP_NO_GRADIENT:
		pBrush = CBCGPGdiPlusHelper::CreateSolidBrush(GPColor(brush.GetColor(), opacity));
		break;

	case CBCGPBrush::BCGP_GRADIENT_HORIZONTAL:
	case CBCGPBrush::BCGP_GRADIENT_VERTICAL:
	case CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT:
	case CBCGPBrush::BCGP_GRADIENT_DIAGONAL_RIGHT:
	case CBCGPBrush::BCGP_GRADIENT_CENTER_HORIZONTAL:
	case CBCGPBrush::BCGP_GRADIENT_CENTER_VERTICAL:
	case CBCGPBrush::BCGP_GRADIENT_PIPE_HORIZONTAL:
	case CBCGPBrush::BCGP_GRADIENT_PIPE_VERTICAL:
		{
			Gdiplus::PointF pt1((Gdiplus::REAL)0.0, (Gdiplus::REAL)0.0);
			Gdiplus::PointF pt2((Gdiplus::REAL)1.0, (Gdiplus::REAL)0.0);

			if (stopsC > 0)
			{
				color1 = GPColor(stops[0].m_Color, opacity);
				color2 = GPColor(stops[stopsC - 1].m_Color, opacity);
			}

			Gdiplus::LinearGradientBrush* pLinearBrush = 
				CBCGPGdiPlusHelper::CreateLinearGradientBrush(pt1, pt2, color1, color2);

			if (pLinearBrush != NULL)
			{
				Gdiplus::Color* presetColors = NULL;
				Gdiplus::REAL* blendPosition = NULL;

				if (stopsC > 0)
				{
					presetColors  = CBCGPGdiPlusHelper::CreateArrayColor(stopsC);
					blendPosition = CBCGPGdiPlusHelper::CreateArrayF(stopsC);

					for (int i = 0; i < stopsC; i++)
					{
						presetColors[i]  = GPColor(stops[i].m_Color, opacity);
						blendPosition[i] = (Gdiplus::REAL)stops[i].m_dblOffset;
					}
				}
				else if (type == CBCGPBrush::BCGP_GRADIENT_CENTER_HORIZONTAL ||
						 type == CBCGPBrush::BCGP_GRADIENT_CENTER_VERTICAL)
				{
					stopsC = 3;
					presetColors  = CBCGPGdiPlusHelper::CreateArrayColor(stopsC);
					blendPosition = CBCGPGdiPlusHelper::CreateArrayF(stopsC);

					presetColors[0]  = color1;
					blendPosition[0] = (Gdiplus::REAL)0.0;

					presetColors[1]  = color2;
					blendPosition[1] = (Gdiplus::REAL)0.5;

					presetColors[2]  = color1;
					blendPosition[2] = (Gdiplus::REAL)1.0;
				}
				else if (type == CBCGPBrush::BCGP_GRADIENT_PIPE_HORIZONTAL ||
						 type == CBCGPBrush::BCGP_GRADIENT_PIPE_VERTICAL)
				{
					stopsC = 3;
					presetColors  = CBCGPGdiPlusHelper::CreateArrayColor(stopsC);
					blendPosition = CBCGPGdiPlusHelper::CreateArrayF(stopsC);

					presetColors[0]  = color2;
					blendPosition[0] = (Gdiplus::REAL)0.0;

					CBCGPColor clrLight = brush.GetGradientColor();
					if ((COLORREF)clrLight == CBCGPColor::White)
					{
						clrLight = brush.GetColor();
						clrLight.MakeLighter(.5);
					}

					presetColors[1]  = GPColor(clrLight, opacity);
					blendPosition[1] = (Gdiplus::REAL)(type == CBCGPBrush::BCGP_GRADIENT_PIPE_VERTICAL
																? 0.7 : 0.3);

					presetColors[2]  = color2;
					blendPosition[2] = (Gdiplus::REAL)1.0;
				}

				if (presetColors != NULL && blendPosition != NULL)
				{
					pLinearBrush->SetInterpolationColors(presetColors, blendPosition, (INT)stopsC);

					CBCGPGdiPlusHelper::DeleteArrayColor(presetColors);
					CBCGPGdiPlusHelper::DeleteArrayF(blendPosition);
				}

				pLinearBrush->SetWrapMode
								(
									brush.GetExtendMode() == CBCGPBrush::BCGP_GRADIENT_EXTEND_MODE_MIRROR
										? Gdiplus::WrapModeTileFlipXY
										: Gdiplus::WrapModeTile
								);
			}

			pBrush = pLinearBrush;
		}
		break;

	case CBCGPBrush::BCGP_GRADIENT_RADIAL_TOP:
	case CBCGPBrush::BCGP_GRADIENT_RADIAL_CENTER:
	case CBCGPBrush::BCGP_GRADIENT_RADIAL_BOTTOM:
	case CBCGPBrush::BCGP_GRADIENT_RADIAL_LEFT:
	case CBCGPBrush::BCGP_GRADIENT_RADIAL_RIGHT:
	case CBCGPBrush::BCGP_GRADIENT_RADIAL_TOP_LEFT:
	case CBCGPBrush::BCGP_GRADIENT_RADIAL_TOP_RIGHT:
	case CBCGPBrush::BCGP_GRADIENT_RADIAL_BOTTOM_LEFT:
	case CBCGPBrush::BCGP_GRADIENT_RADIAL_BOTTOM_RIGHT:
		{
			Gdiplus::PathGradientBrush* pPathBrush = NULL;

			Gdiplus::PointF ptCenter((Gdiplus::REAL)0.0, (Gdiplus::REAL)0.0);
			const Gdiplus::REAL dx = (Gdiplus::REAL)0.5;
			const Gdiplus::REAL dy = (Gdiplus::REAL)0.5;

			if (type == CBCGPBrush::BCGP_GRADIENT_RADIAL_TOP)
			{
				ptCenter.Y -= dy;
			}
			else if (type == CBCGPBrush::BCGP_GRADIENT_RADIAL_BOTTOM)
			{
				ptCenter.Y += dy;
			}
			else if (type == CBCGPBrush::BCGP_GRADIENT_RADIAL_TOP_LEFT)
			{
				ptCenter.X -= dx;
				ptCenter.Y -= dy;
			}
			else if (type == CBCGPBrush::BCGP_GRADIENT_RADIAL_TOP_RIGHT)
			{
				ptCenter.X += dx;
				ptCenter.Y -= dy;
			}
			else if (type == CBCGPBrush::BCGP_GRADIENT_RADIAL_BOTTOM_LEFT)
			{
				ptCenter.X -= dx;
				ptCenter.Y += dy;
			}
			else if (type == CBCGPBrush::BCGP_GRADIENT_RADIAL_BOTTOM_RIGHT)
			{
				ptCenter.X += dx;
				ptCenter.Y += dy;
			}
			else if (type == CBCGPBrush::BCGP_GRADIENT_RADIAL_LEFT)
			{
				ptCenter.X -= dx;
			}
			else if (type == CBCGPBrush::BCGP_GRADIENT_RADIAL_RIGHT)
			{
				ptCenter.X += dx;
			}

			Gdiplus::GraphicsPath path;
			path.AddEllipse(Gdiplus::RectF((Gdiplus::REAL)-1.0, (Gdiplus::REAL)-1.0, (Gdiplus::REAL)2.0, (Gdiplus::REAL)2.0));

			pPathBrush = CBCGPGdiPlusHelper::CreatePathGradientBrush(&path);

			if (pPathBrush != NULL)
			{
				if (stopsC > 0)
				{
					ptCenter.X = (Gdiplus::REAL)0.0;
					ptCenter.Y = (Gdiplus::REAL)0.0;

					Gdiplus::Color* presetColors  = CBCGPGdiPlusHelper::CreateArrayColor(stopsC);
					Gdiplus::REAL* blendPosition = CBCGPGdiPlusHelper::CreateArrayF(stopsC);

					for (int i = 0; i < stopsC; i++)
					{
						int index = stopsC - i - 1;
						presetColors[index]  = GPColor(stops[i].m_Color, opacity);
						blendPosition[index] = (Gdiplus::REAL)(1.0 - stops[i].m_dblOffset);
					}

					pPathBrush->SetInterpolationColors(presetColors, blendPosition, (INT)stopsC);

					CBCGPGdiPlusHelper::DeleteArrayColor(presetColors);
					CBCGPGdiPlusHelper::DeleteArrayF(blendPosition);
				}
				else
				{
					pPathBrush->SetCenterColor(color1);

					INT count = 1;
					pPathBrush->SetSurroundColors(&color2, &count);
				}

				pPathBrush->SetCenterPoint(ptCenter);

				pPathBrush->SetWrapMode(Gdiplus::WrapModeClamp);
			}

			pBrush = pPathBrush;
		}
		break;
	};

	if (pBrush == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}

	brush.Set(this, (LPVOID)pBrush);

#endif

	return brush.GetHandle();
}


BOOL
CBCGPGraphicsManagerGdiPlus::DestroyBrush(CBCGPBrush& brush)
{
#ifndef BCGP_EXCLUDE_GDI_PLUS

	CBCGPGdiPlusHelper::DeleteBrush((Gdiplus::Brush*)brush.GetHandle());
	CBCGPGdiPlusHelper::DeleteBrush((Gdiplus::Brush*)brush.GetHandle1());
	CBCGPGdiPlusHelper::DeleteBrush((Gdiplus::Brush*)brush.GetHandle2());

#endif

	brush.Set(NULL, NULL);
	return TRUE;
}


BOOL
CBCGPGraphicsManagerGdiPlus::SetBrushOpacity(CBCGPBrush& brush)
{
	if (brush.GetHandle() != NULL)
	{
		DestroyBrush(brush);
		CreateBrush(brush);
	}

	return TRUE;
}


LPVOID
CBCGPGraphicsManagerGdiPlus::CreateImage(CBCGPImage& image)
{
	if (m_pOriginal != NULL)
	{
		return (CBCGPGraphicsManagerGdiPlus*)m_pOriginal->CreateImage(image);
	}

	if (image.GetHandle() != NULL)
	{
		CBCGPGraphicsManager* pGM = image.GetGraphicsManager();

		if (!IsGraphicsManagerValid(pGM))
		{
			return NULL;
		}

		return image.GetHandle();
	}

#ifndef BCGP_EXCLUDE_GDI_PLUS

	CString strPath = image.GetPath();
	UINT uiResID = image.GetResourceID();
	HICON hIcon = image.GetIcon();
	HBITMAP hBitmap = image.GetBitmap();

	if (strPath.IsEmpty() && hIcon == NULL && hBitmap == NULL && uiResID == 0)
	{
		return NULL;
	}

	Gdiplus::Image* pImage = NULL;

	if (!strPath.IsEmpty())
	{
		CString strExt;
		
#if _MSC_VER < 1400
		_tsplitpath (strPath, NULL, NULL, NULL, strExt.GetBuffer (_MAX_EXT));
#else
		_tsplitpath_s (strPath, NULL, 0, NULL, 0, NULL, 0, strExt.GetBuffer (_MAX_EXT), _MAX_EXT);
#endif
		
		strExt.ReleaseBuffer ();
		
		if (strExt.CompareNoCase (_T(".svg")) == 0)
		{
#ifndef BCGP_EXCLUDE_SVG
			CBCGPSVGImage svg;
			if (svg.LoadFromFile(strPath))
			{
				HBITMAP hBmpSrc = svg.ExportToBitmap();
				if (hBmpSrc != NULL)
				{
					pImage = CBCGPGdiPlusHelper::CreateImage(hBmpSrc, image.IsIgnoreAlphaBitmap());
					::DeleteObject(hBmpSrc);
				}
			}
#endif			
		}
		else
		{
			LPWSTR lpPath = NULL;
	#ifdef _UNICODE
			lpPath = (LPWSTR)(LPCWSTR)strPath;
	#else
			{
				int length = strPath.GetLength() + 1;
				lpPath = (LPWSTR)new WCHAR[length];
				lpPath = AfxA2WHelper(lpPath, (LPCSTR)strPath, length);
			}
	#endif

			pImage = Gdiplus::Bitmap::FromFile(lpPath);

			if ((LPVOID)lpPath != (LPVOID)(LPCTSTR)strPath)
			{
				delete [] lpPath;
			}
		}
	}
	else if (hIcon != NULL)
	{
		pImage = CBCGPGdiPlusHelper::CreateImage(hIcon);
	}
	else if (hBitmap != NULL)
	{
		pImage = CBCGPGdiPlusHelper::CreateImage(hBitmap, image.IsIgnoreAlphaBitmap());
	}
	else if (uiResID != 0)
	{

		LPCTSTR lpszResourceName = MAKEINTRESOURCE(uiResID);
		ASSERT(lpszResourceName != NULL);

		// Locate the resource.
		LPCTSTR szResType = image.GetResourceType();

		HINSTANCE hInst = NULL;

		LPCTSTR szResources[] = {
#ifndef BCGP_EXCLUDE_SVG
			_T("SVG"),
#endif
			_T("PNG"), 
			RT_BITMAP, 
			NULL };

		if (szResType == NULL)
		{
			int i = 0;
			while (szResources[i] != NULL)
			{
				hInst = AfxFindResourceHandle(lpszResourceName, szResources[i]);
				if (hInst != NULL)
				{
					if (::FindResource(hInst, lpszResourceName, szResources[i]) != NULL)
					{
						szResType = szResources[i];
						break;
					}
				}

				hInst = NULL;
				i++;
			}
		}

		if (hInst == NULL && szResType != NULL)
		{
			hInst = AfxFindResourceHandle(lpszResourceName, szResType);
		}

		if (szResType != NULL && szResType > RT_HTML && lstrcmpi(szResType, _T("SVG")) == 0)
		{
#ifndef BCGP_EXCLUDE_SVG
			CBCGPSVGImageList svg;
			if (svg.LoadStr(lpszResourceName, hInst))
			{
				HBITMAP hBmp = svg.ExportToBitmap();
				if (hBmp != NULL)
				{
					pImage = CBCGPGdiPlusHelper::CreateImage(hBmp, image.IsIgnoreAlphaBitmap());
					::DeleteObject(hBmp);
				}
			}
#endif
		}
		else
		{
			CBCGPToolBarImages images;
			images.SetTransparentColor((COLORREF)-1);
			images.SetPreMultiplyAutoCheck();
			images.LoadStr(lpszResourceName, hInst);
			images.SetSingleImage();

			if (images.IsValid())
			{
				pImage = CBCGPGdiPlusHelper::CreateImage(images.GetImageWell(), false);
			}
		}
	}

	if (pImage != NULL)
	{
		CBCGPSize sizeSrc(pImage->GetWidth(), pImage->GetHeight());
		CBCGPSize sizeDest = image.GetDestSize();

		if (image.NeedToPrepare() || (!sizeDest.IsEmpty() && sizeSrc != sizeDest))
		{
			CBCGPGraphicsManagerGdiPlus gm;
			CSize size(sizeSrc);

			if (!sizeDest.IsEmpty())
			{
				size = sizeDest;
			}

			if (size.cx != 0 && size.cy != 0)
			{
				HBITMAP hBitmapNew = CBCGPDrawManager::CreateBitmap_32(size, NULL);
				if (hBitmapNew != NULL)
				{
					CDC dcMem;
					dcMem.CreateCompatibleDC (NULL);

					HBITMAP hbmpOld = (HBITMAP) dcMem.SelectObject (hBitmapNew);

					CBCGPImage imageOld;
					imageOld.Set(&gm, pImage);

					CBCGPRect rect(CBCGPPoint(), size);
					gm.BindDC(&dcMem, rect);
					gm.BeginDraw();
					GPGraphics(gm.m_pGraphics)->Clear(CBCGPGdiPlusHelper::Color(0, 0, 0, 0));
					gm.DrawImage(imageOld, CBCGPPoint(0.0, 0.0), CBCGPSize(size), 1.0, 
						sizeDest.IsEmpty() ? CBCGPImage::BCGP_IMAGE_INTERPOLATION_MODE_NEAREST_NEIGHBOR : CBCGPImage::BCGP_IMAGE_INTERPOLATION_MODE_LINEAR);
					gm.EndDraw();

					imageOld.Detach();

					dcMem.SelectObject(hbmpOld);

					Gdiplus::Image* pImageNew = NULL;

					BOOL bCorrect = !image.NeedToPrepare();
					if (!bCorrect)
					{
						bCorrect = PrepareImage(image, hBitmapNew);
					}

					if (bCorrect)
					{
						pImageNew = CBCGPGdiPlusHelper::CreateImage(hBitmapNew, image.IsIgnoreAlphaBitmap());
					}

					::DeleteObject(hBitmapNew);

					if (pImageNew != NULL)
					{
						CBCGPGdiPlusHelper::DeleteImage(pImage);
						pImage = pImageNew;
					}
				}
			}
		}
	}

	if (pImage == NULL)
	{
		TRACE(_T("CBCGPGraphicsManagerGdiPlus::CreateImage failed!\n"));
		return NULL;
	}

	image.Set(this, (LPVOID)pImage);

#endif

	return image.GetHandle();
}


BOOL
CBCGPGraphicsManagerGdiPlus::DestroyImage(CBCGPImage& image)
{
#ifndef BCGP_EXCLUDE_GDI_PLUS

	CBCGPGdiPlusHelper::DeleteImage((Gdiplus::Bitmap*)image.GetHandle());

#endif

	image.Set(NULL, NULL);
	return TRUE;
}


CBCGPSize
CBCGPGraphicsManagerGdiPlus::GetImageSize(CBCGPImage& image)
{
	CBCGPSize size;

#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(image);

#else

	Gdiplus::Bitmap* pBitmap = (Gdiplus::Bitmap*)CreateImage(image);
	if (pBitmap != NULL)
	{
		size.cx = pBitmap->GetWidth();
		size.cy = pBitmap->GetHeight();
	}

#endif

	return size;
}


BOOL
CBCGPGraphicsManagerGdiPlus::CopyImage(CBCGPImage& imageSrc, CBCGPImage& imageDest, const CBCGPRect& rectSrc)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(imageSrc);
	UNREFERENCED_PARAMETER(imageDest);
	UNREFERENCED_PARAMETER(rectSrc);

	return FALSE;

#else

	Gdiplus::Bitmap* pBitmapSrc = (Gdiplus::Bitmap*)CreateImage(imageSrc);
	if (pBitmapSrc == NULL)
	{
		return FALSE;
	}

	CBCGPRect rectImage(0.0, 0.0, pBitmapSrc->GetWidth(), pBitmapSrc->GetHeight());
	if (!rectSrc.IsRectEmpty())
	{
		if (!rectImage.IntersectRect(rectImage, rectSrc))
		{
			return FALSE;
		}
	}

	Gdiplus::Bitmap* pBitmapDest = pBitmapSrc->Clone(GPRect(this, rectImage), pBitmapSrc->GetPixelFormat());
	if (pBitmapDest == NULL)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	imageDest.Set(this, (LPVOID)pBitmapDest);

	return TRUE;

#endif
}


HBITMAP
CBCGPGraphicsManagerGdiPlus::ExportImageToBitmap(CBCGPImage& /*image*/)
{
	ASSERT(FALSE);
	return NULL;
}


LPVOID
CBCGPGraphicsManagerGdiPlus::CreateStrokeStyle(CBCGPStrokeStyle& /*style*/)
{
	return NULL;
}


BOOL
CBCGPGraphicsManagerGdiPlus::DestroyStrokeStyle(CBCGPStrokeStyle& /*style*/)
{
	return FALSE;
}


BOOL CBCGPGraphicsManagerGdiPlus::InitGDIPlus(BOOL bSuppressBackgroundThread /*= FALSE*/)
{
	return g_GdiPlusHelper.Initialize(bSuppressBackgroundThread);
}

void CBCGPGraphicsManagerGdiPlus::Shutdown()
{
	g_GdiPlusHelper.Shutdown();
}

BOOL CBCGPGraphicsManagerGdiPlus::IsGDIPlusInitialized()
{
	return g_GdiPlusHelper.IsInitialized();
}

BOOL CBCGPGraphicsManagerGdiPlus::IsValid() const
{
	return CBCGPGdiPlusHelper::IsReady();
}

BOOL CBCGPGraphicsManagerGdiPlus::IsGraphicsManagerValid(CBCGPGraphicsManager* pGM)
{
	if (pGM == NULL)
	{
		return TRUE;
	}

	while (pGM->GetOriginal() != NULL)
	{
		pGM = pGM->GetOriginal();
	}

	CBCGPGraphicsManager* pThis = this;

	while (pThis->GetOriginal() != NULL)
	{
		pThis = pThis->GetOriginal();
	}

	if (pGM == pThis)
	{
		return TRUE;
	}

	TRACE0("You cannot use the same GDI+ resource with the different graphics managers!\n");

	return FALSE;
}


void
CBCGPGraphicsManagerGdiPlus::EnableAntialiasing(BOOL bEnable)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(bEnable);

#else

	if (m_pGraphics == NULL)
	{
		return;
	}

	GPGraphics(m_pGraphics)->SetSmoothingMode(bEnable ? Gdiplus::SmoothingModeAntiAlias : Gdiplus::SmoothingModeNone);

#endif
}


BOOL
CBCGPGraphicsManagerGdiPlus::IsAntialiasingEnabled() const
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	return FALSE;

#else

	if (m_pGraphics == NULL)
	{
		return FALSE;
	}

	return GPGraphics(m_pGraphics)->GetSmoothingMode() != Gdiplus::SmoothingModeNone;

#endif
}


BOOL
CBCGPGraphicsManagerGdiPlus::IsTransformSupported() const
{
	return TRUE;
}


void
CBCGPGraphicsManagerGdiPlus::ResetTransform()
{
#ifndef BCGP_EXCLUDE_GDI_PLUS

	if (m_pGraphics == NULL)
	{
		return;
	}

	GPGraphics(m_pGraphics)->ResetTransform();
	m_sizeScaleTransform.SetSize(1.0, 1.0);

#endif
}


void
CBCGPGraphicsManagerGdiPlus::PushTransform()
{
#ifndef BCGP_EXCLUDE_GDI_PLUS

	if (m_pGraphics == NULL)
	{
		return;
	}

	Gdiplus::Matrix* pMatrix = CBCGPGdiPlusHelper::CreateMatrix();

	GPGraphics(m_pGraphics)->GetTransform(pMatrix);

	m_lstTransforms.AddTail((LPVOID)pMatrix);

#endif
}


BOOL
CBCGPGraphicsManagerGdiPlus::PopTransform()
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	return FALSE;

#else

	if (m_pGraphics == NULL)
	{
		return FALSE;
	}

	if (m_lstTransforms.IsEmpty())
	{
		TRACE0("CBCGPGraphicsManagerGdiPlus::PopTransform: transform is not pushed - call CBCGPGraphicsManagerGdiPlus::PushTransform first\n");
		ASSERT(FALSE);

		return FALSE;
	}

	Gdiplus::Matrix* pMatrix = (Gdiplus::Matrix*)m_lstTransforms.RemoveTail();
	ASSERT(pMatrix != NULL);

	GPGraphics(m_pGraphics)->SetTransform(pMatrix);

	CBCGPGdiPlusHelper::DeleteMatrix(pMatrix);

	return TRUE;

#endif
}


void
CBCGPGraphicsManagerGdiPlus::SetRotateTransform(double dblAngle, const CBCGPPoint& ptCenterIn)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(dblAngle);
	UNREFERENCED_PARAMETER(ptCenterIn);

#else

	if (m_pGraphics == NULL)
	{
		return;
	}

	if (dblAngle != 0.0)
	{
		CBCGPPoint ptCenter = ptCenterIn;
		if (ptCenter.x == DBL_MAX)
		{
			ptCenter.x = 0.0;
		}
		if (ptCenter.y == DBL_MAX)
		{
			ptCenter.y = 0.0;
		}

		Gdiplus::Matrix matrix;
		GPGraphics(m_pGraphics)->GetTransform(&matrix);

		matrix.RotateAt
				(
					-(Gdiplus::REAL)dblAngle, 
					GPPoint(this, ptCenter), 
					Gdiplus::MatrixOrderPrepend
				);

		GPGraphics(m_pGraphics)->SetTransform(&matrix);
	}

#endif
}


void
CBCGPGraphicsManagerGdiPlus::SetScaleTransform(const CBCGPSize& size, const CBCGPPoint& ptCenter /*= CBCGPPoint(-1., -1.)*/)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(size);
	UNREFERENCED_PARAMETER(ptCenter);

#else

	if (m_pGraphics == NULL)
	{
		return;
	}

	Gdiplus::Matrix matrix;
	GPGraphics(m_pGraphics)->GetTransform(&matrix);

	if (ptCenter == CBCGPPoint(-1., -1.))
	{
		matrix.Scale((Gdiplus::REAL)size.cx, (Gdiplus::REAL)size.cy, Gdiplus::MatrixOrderPrepend);
	}
	else
	{
		matrix.Translate((Gdiplus::REAL)ptCenter.x, (Gdiplus::REAL)ptCenter.y, Gdiplus::MatrixOrderPrepend);
		matrix.Scale((Gdiplus::REAL)size.cx, (Gdiplus::REAL)size.cy, Gdiplus::MatrixOrderPrepend);
		matrix.Translate(-(Gdiplus::REAL)ptCenter.x, -(Gdiplus::REAL)ptCenter.y, Gdiplus::MatrixOrderPrepend);
	}

	GPGraphics(m_pGraphics)->SetTransform(&matrix);

	m_sizeScaleTransform = size;

#endif
}


void
CBCGPGraphicsManagerGdiPlus::SetTranslateTransform(const CBCGPSize& size)
{
#ifdef BCGP_EXCLUDE_GDI_PLUS

	UNREFERENCED_PARAMETER(size);

#else

	if (m_pGraphics == NULL)
	{
		return;
	}
	
	Gdiplus::Matrix matrix;
	GPGraphics(m_pGraphics)->GetTransform(&matrix);

	matrix.Translate((Gdiplus::REAL)size.cx, (Gdiplus::REAL)size.cy, Gdiplus::MatrixOrderPrepend);
	
	GPGraphics(m_pGraphics)->SetTransform(&matrix);

#endif
}


void
CBCGPGraphicsManagerGdiPlus::SetMatrixTransform(const CArray<double, double>& arValues)
{
	if (m_pGraphics == NULL || arValues.GetSize() == 0)
	{
		return;
	}

#ifndef BCGP_EXCLUDE_GDI_PLUS

	Gdiplus::Matrix matrix;
	GPMatrix(matrix, arValues);

	GPGraphics(m_pGraphics)->MultiplyTransform(&matrix, Gdiplus::MatrixOrderPrepend);

#endif
}
