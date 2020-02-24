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
//
// BCGPGridSparkLine.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPGridSparkLine.h"
#include "BCGPChartVisualObject.h"

#ifndef BCGP_EXCLUDE_GRID_CTRL

/////////////////////////////////////////////////////////////////////////////
// CBCGPGridSparklineItem

IMPLEMENT_DYNCREATE (CBCGPGridSparklineItem, CBCGPGridItem)

CBCGPGridSparklineItem::CBCGPGridSparklineItem (SparklineType type, DWORD_PTR dwData) :
	CBCGPGridItem (_variant_t ((LPCTSTR) _T("")), dwData)
{
	m_bAllowEdit = FALSE;
	m_pChart = NULL;
	m_nMaxMarkerSize = 0;
	m_bDefaultSelColor = TRUE;

	if (type != SparklineTypeLine)
	{
		SetType(type);
	}
}
//*****************************************************************************************
CBCGPGridSparklineItem::CBCGPGridSparklineItem (const CBCGPDoubleArray& arData, SparklineType type, DWORD_PTR dwData) :
	CBCGPGridItem (_variant_t ((LPCTSTR) _T("")), dwData)
{
	m_bAllowEdit = FALSE;
	m_pChart = NULL;
	m_nMaxMarkerSize = 0;
	m_bDefaultSelColor = TRUE;

	if (type != SparklineTypeLine)
	{
		SetType(type);
	}

	if (arData.GetSize() > 0)
	{
		AddData(arData);
	}
}
//*****************************************************************************************
CBCGPGridSparklineItem::CBCGPGridSparklineItem(const CBCGPGridSparklineDataArray& arData, SparklineType type, DWORD_PTR dwData) :
	CBCGPGridItem (_variant_t ((LPCTSTR) _T("")), dwData)
{
	m_bAllowEdit = FALSE;
	m_pChart = NULL;
	m_nMaxMarkerSize = 0;
	m_bDefaultSelColor = TRUE;
	
	if (type != SparklineTypeLine)
	{
		SetType(type);
	}
	
	if (arData.GetSize() > 0)
	{
		AddData(arData);
	}
}
//*****************************************************************************************
CBCGPGridSparklineItem::~CBCGPGridSparklineItem()
{
	if (m_pChart != NULL)
	{
		delete m_pChart;
	}
}
//*****************************************************************************
CBCGPChartVisualObject* CBCGPGridSparklineItem::GetChart()
{
	if (m_pChart == NULL)
	{
		m_pChart = new CBCGPChartVisualObject;
		InitChart();
	}

	return m_pChart;
}
//*****************************************************************************
int CBCGPGridSparklineItem::SparklineTypeToChartType (SparklineType type)
{
	switch (type)
	{
	case SparklineTypeLine:
		return (int)BCGPChartLine;
		
	case SparklineTypePie:
		return (int)BCGPChartPie;
		
	case SparklineTypeColumn:
		return (int)BCGPChartColumn;
		
	case SparklineTypeBar:
		return (int)BCGPChartBar;

	case SparklineTypeBarSmart:
		return (int)BCGPChartBarSmart;
		
	case SparklineTypeArea:
		return (int)BCGPChartArea;
		
	case SparklineTypeBubble:
		return (int)BCGPChartBubble;
		
	case SparklineTypeDoughnut:
		return (int)BCGPChartDoughnut;
	}

	return (int)BCGPChartDefault;
}
//*****************************************************************************
void CBCGPGridSparklineItem::SetType(SparklineType type)
{
	CBCGPChartVisualObject*	pChart = GetChart();
	ASSERT_VALID(pChart);

	pChart->SetChartType((BCGPChartCategory)SparklineTypeToChartType(type));
}
//*****************************************************************************
CBCGPGridSparklineItem::SparklineType CBCGPGridSparklineItem::GetType()
{
	CBCGPChartVisualObject*	pChart = GetChart();
	ASSERT_VALID(pChart);

	switch (pChart->GetChartCategory())
	{
	case BCGPChartLine:
		return SparklineTypeLine;

	case BCGPChartPie:
		return SparklineTypePie;
		
	case BCGPChartColumn:
		return SparklineTypeColumn;
		
	case BCGPChartBar:
		return SparklineTypeBar;

	case BCGPChartBarSmart:
		return SparklineTypeBarSmart;
		
	case BCGPChartArea:
		return SparklineTypeArea;
		
	case BCGPChartBubble:
		return SparklineTypeBubble;
		
	case BCGPChartDoughnut:
		return SparklineTypeDoughnut;
	}

	return (SparklineType)-1;
}
//*****************************************************************************
COLORREF CBCGPGridSparklineItem::GetFillColor()
{
	return (COLORREF)-1;
}
//*****************************************************************************
CBCGPChartSeries* CBCGPGridSparklineItem::GetSeries(int nSeries, SparklineType type)
{
	CBCGPChartVisualObject*	pChart = GetChart();
	ASSERT_VALID(pChart);

	if (type != SparklineTypeDefault && type != GetType())
	{
		return pChart->CreateSeries(_T(""), CBCGPColor(), BCGP_CT_DEFAULT,
			(BCGPChartCategory)SparklineTypeToChartType(type), nSeries);
	}
	else
	{
		return pChart->GetSeries(nSeries, TRUE);
	}
}
//*****************************************************************************
void CBCGPGridSparklineItem::AddData(const CBCGPDoubleArray& arData, int nSeries, SparklineType type)
{
	CBCGPChartSeries* pSeries = GetSeries(nSeries, type);
	if (pSeries == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	ASSERT_VALID(pSeries);

	pSeries->AddDataPoints(arData);
}
//*****************************************************************************
void CBCGPGridSparklineItem::AddData(const CBCGPGridSparklineDataArray& arData, int nSeries, SparklineType type)
{
	if (arData.GetSize() == 0)
	{
		return;
	}

	CBCGPChartVisualObject*	pChart = GetChart();
	ASSERT_VALID(pChart);

	CBCGPChartSeries* pSeries = GetSeries(nSeries, type);
	if (pSeries == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	ASSERT_VALID(pSeries);

	if (type != SparklineTypeDefault && type != GetType())
	{
		pSeries->SetChartType((BCGPChartCategory)SparklineTypeToChartType(type));
	}

	int nMaxMarkerSizePrev = m_nMaxMarkerSize;

	for (int i = 0; i < (int)arData.GetSize(); i++)
	{
		const CBCGPBrush& brDataPointFill = arData[i].m_brDataPointFill;
		const CBCGPBrush& brDataPointBorder = arData[i].m_brDataPointBorder;
		
		BOOL bShowMarker = arData[i].m_MarkerOptions.m_bShowMarker;
		
		int nDPIndex = -1;

		if (!brDataPointFill.IsEmpty() || !brDataPointBorder.IsEmpty() || bShowMarker)
		{
			BCGPChartFormatSeries fs;
			
			if (!brDataPointFill.IsEmpty() || !brDataPointBorder.IsEmpty())
			{
				fs.SetSeriesFill(brDataPointFill);
				fs.SetSeriesLineColor(brDataPointBorder.IsEmpty() ? brDataPointFill : brDataPointBorder);
			}
			
			if (bShowMarker)
			{
				const CBCGPBrush& brMarkerFill = arData[i].m_brMarkerFill;
				const CBCGPBrush& brMarkerBorder = arData[i].m_brMarkerBorder;

				fs.m_markerFormat.m_options = arData[i].m_MarkerOptions;
				
				if (!brMarkerFill.IsEmpty() || !brMarkerBorder.IsEmpty())
				{
					fs.SetMarkerFill(brMarkerFill);
					fs.SetMarkerLineColor(brMarkerBorder.IsEmpty() ? brMarkerFill : brMarkerBorder);
				}

				m_nMaxMarkerSize = max(m_nMaxMarkerSize, fs.m_markerFormat.m_options.GetMarkerSize());
			}
			
			nDPIndex = pSeries->AddDataPoint(arData[i].m_dblValue, &fs);
		}
		else
		{
			nDPIndex = pSeries->AddDataPoint(arData[i].m_dblValue);
		}

		if (arData[i].m_dblValue1 != 0. && nDPIndex >= 0)
		{
			pSeries->SetDataPointValue(nDPIndex, arData[i].m_dblValue1, CBCGPChartData::CI_Y1);
		}
	}

	if (nMaxMarkerSizePrev != m_nMaxMarkerSize)
	{
		const double dblPadding = 0.5 * (double)(1.0 + m_nMaxMarkerSize);

		pChart->GetChartAreaFormat().SetContentPadding(CBCGPSize(dblPadding, dblPadding));
	}
}
//*****************************************************************************
BOOL CBCGPGridSparklineItem::UpdateDataPoint(int nIndex, double dblValue, int nSeries/* = 0*/)
{
	CBCGPChartVisualObject*	pChart = GetChart();
	ASSERT_VALID(pChart);

	CBCGPChartSeries* pSeries = pChart->GetSeries(nSeries);
	if (pSeries == NULL)
	{
		return FALSE;
	}

	if (!pSeries->SetDataPointValue(nIndex, dblValue))
	{
		return FALSE;
	}

	pChart->RecalcMinMaxValues();
	pChart->SetDirty();

	return TRUE;
}
//*****************************************************************************
BOOL CBCGPGridSparklineItem::ShowDataPointMarker(int nIndex, const BCGPChartMarkerOptions& markerOptions, 
						 const CBCGPBrush& brMarkerFill, const CBCGPBrush& brMarkerBorder, int nSeries)
{
	CBCGPChartVisualObject*	pChart = GetChart();
	ASSERT_VALID(pChart);

	CBCGPChartSeries* pSeries = pChart->GetSeries(nSeries);
	if (pSeries == NULL)
	{
		return FALSE;
	}

	if (nIndex < 0 || nIndex >= pSeries->GetDataPointCount())
	{
		return FALSE;
	}

	int nMaxMarkerSizePrev = m_nMaxMarkerSize;

	BCGPChartFormatSeries fs;
	fs.m_markerFormat.m_options = markerOptions;
	
	if (!brMarkerFill.IsEmpty() || !brMarkerBorder.IsEmpty())
	{
		fs.SetMarkerFill(brMarkerFill);
		fs.SetMarkerLineColor(brMarkerBorder.IsEmpty() ? brMarkerFill : brMarkerBorder);
	}

	m_nMaxMarkerSize = max(m_nMaxMarkerSize, fs.m_markerFormat.m_options.GetMarkerSize());

	pSeries->SetMarkerOptions(markerOptions, nIndex);
	pSeries->SetMarkerFill(brMarkerFill, nIndex);
	pSeries->SetMarkerLineColor(brMarkerBorder, nIndex);

	if (nMaxMarkerSizePrev != m_nMaxMarkerSize)
	{
		const double dblPadding = 0.5 * (double)(1.0 + m_nMaxMarkerSize);
		pChart->GetChartAreaFormat().SetContentPadding(CBCGPSize(dblPadding, dblPadding));
	}

	pChart->RecalcMinMaxValues();
	pChart->SetDirty();

	return TRUE;
}
//*****************************************************************************
void CBCGPGridSparklineItem::RemoveData(int nSeries)
{
	CBCGPChartVisualObject*	pChart = GetChart();
	ASSERT_VALID(pChart);

	pChart->CleanUpChartData(nSeries);

	if (nSeries == -1 || pChart->GetSeriesCount() == 1)
	{
		if (m_nMaxMarkerSize != 0)
		{
			m_nMaxMarkerSize = 0;
			pChart->GetChartAreaFormat().SetContentPadding(CBCGPSize());
		}
	}
}
//*****************************************************************************
void CBCGPGridSparklineItem::OnDrawValue (CDC* pDC, CRect rect)
{
	ASSERT_VALID (this);

	OnFillBackground(pDC, rect);

	CBCGPGridCtrl* pGrid = GetOwnerList();

	CBCGPChartVisualObject*	pChart = GetChart();
	ASSERT_VALID(pChart);

	if (pChart->GetRect().IsRectEmpty() || pGrid == NULL)
	{
		return;
	}

	CBCGPGraphicsManager* pGM = pGrid->GetGraphicsManager();
	if (pGM == NULL)
	{
		return;
	}

	CRect rectClient;
	pGrid->GetClientRect(rectClient);

	pGM->BindDC(pDC, rectClient);
	
	if (!pGM->BeginDraw())
	{
		return;
	}

	pChart->OnDraw(pGM, rect);

	pGM->EndDraw();
}
//*****************************************************************************
void CBCGPGridSparklineItem::OnPrintValue (CDC* pDC, CRect rect)
{
	ASSERT_VALID (this);

	CBCGPGridCtrl* pGrid = GetOwnerList();

	CBCGPChartVisualObject*	pChart = GetChart();
	ASSERT_VALID(pChart);

	if (pChart->GetRect().IsRectEmpty() || pGrid == NULL)
	{
		return;
	}

	CBCGPGraphicsManager* pGM = pGrid->GetGraphicsManager();
	if (pGM == NULL)
	{
		return;
	}


	ASSERT_VALID(pGM);

	const CBCGPSize sizeScaleOld = pChart->GetScaleRatio ();

	HDC hDCFrom = ::GetDC(NULL);
	double dXMul = (double) pDC->GetDeviceCaps(LOGPIXELSX);			// pixels in print dc
	double dXDiv = (double) ::GetDeviceCaps(hDCFrom, LOGPIXELSX);	// pixels in screen dc
	double dYMul = (double) pDC->GetDeviceCaps(LOGPIXELSY);			// pixels in print dc
	double dYDiv = (double) ::GetDeviceCaps(hDCFrom, LOGPIXELSY);	// pixels in screen dc
	::ReleaseDC(NULL, hDCFrom);

	pChart->SetScaleRatio (CBCGPSize (dXMul / dXDiv, dYMul / dYDiv));
	
	HBITMAP hmbpDib = pChart->ExportToBitmap(pGM);

	pChart->SetScaleRatio (sizeScaleOld);

	if (hmbpDib == NULL)
	{
		ASSERT (FALSE);
		return;
	}

	const CSize size = rect.Size();
	const CSize sizeImage = size;

	CDC dcMem;
	dcMem.CreateCompatibleDC (pDC);

	HBITMAP hbmpOld = (HBITMAP)dcMem.SelectObject (hmbpDib);
	if (hbmpOld != NULL)
	{
		pDC->StretchBlt (rect.left, rect.top, rect.Width (), rect.Height (), &dcMem, 0, 0, sizeImage.cx, sizeImage.cy, SRCCOPY);
		dcMem.SelectObject(hbmpOld);
	}

	::DeleteObject(hmbpDib);
}
//*****************************************************************************
BOOL CBCGPGridSparklineItem::OnEdit (LPPOINT)
{
	return FALSE;
}
//*****************************************************************************
void CBCGPGridSparklineItem::OnPosSizeChanged (CRect rectOld)
{
	ASSERT_VALID (this);

	CBCGPGridItem::OnPosSizeChanged(rectOld);

	CBCGPGridCtrl* pGrid = GetOwnerList ();
	if (pGrid != NULL && !pGrid->IsColumnBeingResized())
	{
		CBCGPChartVisualObject*	pChart = GetChart();
		ASSERT_VALID(pChart);

		pChart->SetScaleRatio(CBCGPSize(pGrid->GetScale(), pGrid->GetScale()));

		CBCGPRect rect = m_Rect;

		if (pChart->GetRect() != m_Rect)
		{
			pChart->SetRect(rect, TRUE);
			pChart->SetDirty (TRUE);
		}
	}
}
//*****************************************************************************
void CBCGPGridSparklineItem::InitChart()
{
	ASSERT_VALID(m_pChart);

	OnAfterChangeGridColors();

	m_pChart->SetThumbnailMode(TRUE, BCGP_CHART_THUMBNAIL_DRAW_MARKERS);
	m_pChart->ShowChartTitle(FALSE);
	m_pChart->GetChartAreaFormat().SetContentPadding(CBCGPSize());
	m_pChart->SetTitleAreaPadding(CBCGPSize());
	m_pChart->GetTitleLabelFormat().SetContentPadding(CBCGPSize());
	m_pChart->SetPlotAreaPadding(CBCGPRect());
	m_pChart->SetLegendAreaPadding(CBCGPSize());

	m_pChart->ShowAxisGridLines(BCGP_CHART_X_PRIMARY_AXIS, FALSE, FALSE);
	m_pChart->ShowAxisGridLines(BCGP_CHART_Y_PRIMARY_AXIS, FALSE, FALSE);

	CBCGPChartAxis* pAxisX = m_pChart->GetChartAxis(BCGP_CHART_X_PRIMARY_AXIS);
	ASSERT_VALID(pAxisX);
	pAxisX->m_axisNameFormat.SetContentPadding(CBCGPSize());
	pAxisX->m_axisLabelType = CBCGPChartAxis::ALT_NO_LABELS;

	CBCGPChartAxis* pAxisY = m_pChart->GetChartAxis(BCGP_CHART_Y_PRIMARY_AXIS);
	ASSERT_VALID(pAxisY);
	pAxisY->m_axisNameFormat.SetContentPadding(CBCGPSize());
	pAxisY->m_axisLabelType = CBCGPChartAxis::ALT_NO_LABELS;

	m_pChart->ShowAxis(BCGP_CHART_X_PRIMARY_AXIS, FALSE);
	m_pChart->ShowAxis(BCGP_CHART_Y_PRIMARY_AXIS, FALSE);
}
//*****************************************************************************
void CBCGPGridSparklineItem::OnAfterChangeGridColors()
{
	CBCGPGridItem::OnAfterChangeGridColors();

	CBCGPGridCtrl* pGrid = GetOwnerList();
	if (pGrid == NULL)
	{
		return;
	}

	CBCGPChartVisualObject*	pChart = GetChart();
	ASSERT_VALID(pChart);

	const CBCGPGridColors& colorsGrid = pGrid->GetColorTheme();

	CBCGPChartTheme colorsChart;
	CBCGPColor clrFill = colorsGrid.m_EvenColors.m_clrBackground == (COLORREF)-1 ? colorsGrid.m_clrBackground : colorsGrid.m_EvenColors.m_clrBackground;

	COLORREF clrGridText = globalData.clrWindowText;

	if (colorsGrid.m_EvenColors.m_clrText != (COLORREF)-1)
	{
		clrGridText = colorsGrid.m_EvenColors.m_clrText;
	}
	else if (colorsGrid.m_clrText != (COLORREF)-1)
	{
		clrGridText = colorsGrid.m_clrText;
	}

	CBCGPColor clrOutline = clrGridText;
	CBCGPColor clrText = clrGridText;

	if (clrOutline.IsDark())
	{
		clrOutline.MakeLighter(.5);
	}
	else
	{
		clrOutline.MakeDarker(.5);
	}

	CBCGPChartTheme::InitChartColors(colorsChart, 
		CBCGPColor(), clrOutline, clrText,
		CBCGPColor(), CBCGPColor(), .04, clrFill.IsDark());

	colorsChart.m_brPlotLineColor.Empty();
	colorsChart.m_brPlotFillColor.Empty();
	colorsChart.m_brChartFillColor.Empty();

	for (int i = 0; i < BCGP_GRID_SPARKLINES_CHART_SERIES_NUM; i++)
	{
		COLORREF clrSeriesFill = colorsGrid.m_SparklineSeriesColors[i].m_clrBackground;
		COLORREF clrSeriesBorder = colorsGrid.m_SparklineSeriesColors[i].m_clrBorder;

		if (clrSeriesFill == (COLORREF)-1 || clrSeriesBorder == (COLORREF)-1)
		{
			switch (i)
			{
			case 0:
			default:
				clrSeriesBorder = clrSeriesFill = RGB (1, 168, 220);
				break;

			case 1:
				clrSeriesBorder = clrSeriesFill = RGB(237, 125, 49);
				break;

			case 2:
				clrSeriesBorder = clrSeriesFill = RGB(112, 173, 71);
				break;

			case 3:
				clrSeriesBorder = clrSeriesFill = RGB(165, 165, 165);
				break;

			case 4:
				clrSeriesBorder = clrSeriesFill = RGB(255, 192, 0);
				break;
			}
		}

		colorsChart.m_seriesColors[i].m_brElementFillColor.SetColors(clrSeriesFill, clrSeriesFill, CBCGPBrush::BCGP_NO_GRADIENT);
		colorsChart.m_seriesColors[i].m_brElementLineColor.SetColor(clrSeriesBorder);
	}
	
	m_bDefaultSelColor = colorsGrid.m_bSparklineDefaultSelColor;

	pChart->SetColors(colorsChart);
	pChart->SetDirty (TRUE);
}

#endif // BCGP_EXCLUDE_GRID_CTRL
