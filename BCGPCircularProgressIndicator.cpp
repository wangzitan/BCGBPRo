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
// BCGPCircularProgressIndicator.cpp: implementation of the CBCGPCircularProgressIndicatorImpl class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <math.h>
#include "BCGPMath.h"
#include "BCGPCircularProgressIndicator.h"
#include "BCGPTextGaugeImpl.h"
#include "BCGPGlobalUtils.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#ifndef PBM_GETSTEP
	#define PBM_GETSTEP             (WM_USER+13)
#endif

IMPLEMENT_DYNCREATE(CBCGPCircularProgressIndicatorImpl, CBCGPCircularGaugeImpl)
IMPLEMENT_DYNAMIC(CBCGPCircularProgressIndicatorCtrl, CBCGPVisualCtrl)

LRESULT CBCGPCircularProgressIndicatorCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	CBCGPCircularProgressIndicatorImpl* pProgress = NULL;
	if (message >= PBM_SETRANGE && message <= PBM_GETSTEP)
	{
		pProgress = GetCircularProgressIndicator();
	}

	LRESULT lRes = CBCGPVisualCtrl::WindowProc(message, wParam, lParam);
	
	if (pProgress == NULL)
	{
		return lRes;
	}
	
	switch (message)
	{
	case PBM_SETRANGE:
		pProgress->SetRange((double)LOWORD(lParam), (double)HIWORD(lParam));
		lRes = 0L;
		break;
		
	case PBM_GETRANGE:
		{
			double l = 0.0;
			double h = 0.0;
			pProgress->GetRange(l, h);
			
			PPBRANGE range = (PPBRANGE)lParam;
			if (range != NULL)
			{
				range->iLow = (int)l;
				range->iHigh = (int)h;
			}
			
			lRes = (LRESULT)(wParam ? l : h);
		}
		break;
		
	case PBM_SETRANGE32:
		pProgress->SetRange((double)wParam, (double)lParam);
		lRes = 0L;
		break;
		
	case PBM_SETPOS:
		lRes = (LRESULT)pProgress->SetPos((double)wParam);
		break;
		
	case PBM_GETPOS:
		lRes = (LRESULT)pProgress->GetPos();
		break;
		
	case PBM_SETSTEP:
		lRes = (LRESULT)pProgress->SetStep((double)wParam);
		break;
		
	case PBM_GETSTEP:
		lRes = (LRESULT)pProgress->GetStep(FALSE);
		break;
		
	case PBM_DELTAPOS:
		lRes = (LRESULT)pProgress->OffsetPos((double)wParam);
		break;
		
	case PBM_STEPIT:
		lRes = (LRESULT)pProgress->StepIt();
		break;
		
	case PBM_SETMARQUEE:
		if (pProgress->GetOptions().m_bMarqueeStyle)
		{
			double delta = (double)lParam;
			if (delta == 0.0)
			{
				delta = 30.0;
			}
			
			pProgress->StartMarquee((BOOL)wParam, bcg_clamp(delta / 1000.0, 0.0, 1.0));
		}
		
		lRes = (LRESULT)TRUE;
		break;
	};
	
	return lRes;
}

void CBCGPCircularProgressIndicatorCtrl::OnAfterCreateWnd()
{
	CBCGPVisualCtrl::OnAfterCreateWnd();
	
	if (m_bIsPopup)
	{
		CBCGPCircularProgressIndicatorImpl* pProgress = GetCircularProgressIndicator();
		if (pProgress != NULL)
		{
			pProgress->EnableImageCache(FALSE);
		}
	}
}

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPCircularProgressIndicatorImpl::CBCGPCircularProgressIndicatorImpl(CBCGPVisualContainer* pContainer) :
	CBCGPCircularGaugeImpl(pContainer)
{
	m_pRenderer = NULL;

	m_bMarqueeStarted = FALSE;
	m_dblMarquee = 0.0;

	m_dblStep = 0.0;

	m_bIsPressed = FALSE;
	m_bIsSmall = FALSE;

	SetFrameSize(0);
	SetCapSize(0);
	SetTextLabelFormat(_T(""));
	SetTickMarkSize(0, TRUE);
	SetTickMarkSize(0, FALSE);
	
	SetClosedRange(0, 100);

	SetColors(CBCGPCircularGaugeColors::BCGP_CIRCULAR_GAUGE_VISUAL_MANAGER);
}
//*******************************************************************************
CBCGPCircularProgressIndicatorImpl::~CBCGPCircularProgressIndicatorImpl()
{
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::StartMarquee(BOOL bStart /*= TRUE*/, double dblSpeed /* = 0.0*/)
{
	if (!IsMarqueeStyle())
	{
		return;
	}

	if (m_bMarqueeStarted == bStart)
	{
		return;
	}

	m_bMarqueeStarted = bStart;
	m_dblMarquee = 0.0;

	if (m_bMarqueeStarted)
	{
		double dblDelay = 0.0;

		if (dblSpeed <= 0.0)
		{
			dblDelay = m_bIsSmall ? 0.01 : 0.03;
		}
		else
		{
			dblDelay = (1.0 - bcg_clamp(dblSpeed, 0.001, 0.9999)) * 0.1;
		}

		m_arData[0]->StartFlashAnimation(dblDelay);
	}
	else
	{
		m_arData[0]->StopAnimation();
	}
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::SetMarqueeValue(double dblValue)
{
	if (!IsMarqueeStyle())
	{
		return;
	}

	if (dblValue < 0.0)
	{
		m_bMarqueeStarted = FALSE;
		m_dblMarquee = 0.0;
	}
	else
	{
		m_bMarqueeStarted = TRUE;
		m_dblMarquee = bcg_clamp(dblValue, 0.0, 1.0);
	}
}
//*******************************************************************************
double CBCGPCircularProgressIndicatorImpl::SetPos(double nPos)
{
	if (IsMarqueeStyle())
	{
		return 0.0;
	}

	double dblPrev = GetPos();

	double start = GetStart();
	double finish = GetFinish();

	if (start < finish)
	{
		nPos = bcg_clamp(nPos, start, finish);
	}
	else
	{
		nPos = bcg_clamp(nPos, finish, start);
	}

	if (m_arSubGauges.GetSize() > 0 && !m_Options.m_strPercentageFormat.IsEmpty())
	{
		CBCGPTextGaugeImpl* pLabel1 = DYNAMIC_DOWNCAST(CBCGPTextGaugeImpl, m_arSubGauges[0]);
		ASSERT_VALID(pLabel1);

		pLabel1->SetText(OnPreparePercentageLabel(nPos));
	}
	
	SetLevelBarValue(nPos, 0, 0, TRUE);
	return dblPrev;
}
//*******************************************************************************
double CBCGPCircularProgressIndicatorImpl::GetPos() const
{
	if (m_arLevelBars.GetSize() > 0)
	{
		return (const_cast<CBCGPCircularProgressIndicatorImpl*>(this))->GetLevelBar(0)->GetValue();
	}

	return 0.0;
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::SetRange(double nLower, double nUpper)
{
	SetClosedRange(nLower, nUpper);
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::GetRange(double& nLower, double& nUpper) const
{
	nLower = GetStart();
	nUpper = GetFinish();
}
//*******************************************************************************
double CBCGPCircularProgressIndicatorImpl::OffsetPos(double nPos)
{
	if (IsMarqueeStyle())
	{
		return 0.0;
	}
	
	const double dblPrev = GetPos();

	if (nPos != 0.0)
	{
		SetPos(dblPrev + nPos);
	}

	return dblPrev;
}
//*******************************************************************************
double CBCGPCircularProgressIndicatorImpl::SetStep(double nStep)
{
	double dblPrevStep = m_dblStep;
	m_dblStep = nStep;

	return dblPrevStep;
}
//*******************************************************************************
double CBCGPCircularProgressIndicatorImpl::GetStep(BOOL bCalcIfZero /*= FALSE*/) const
{
	double dblStep = m_dblStep;
	if (dblStep == 0.0 && bCalcIfZero)
	{
		dblStep = (GetStart() < GetFinish()) ? 1.0 : -1.0;
	}

	return dblStep;
}
//*******************************************************************************
double CBCGPCircularProgressIndicatorImpl::StepIt()
{
	if (IsMarqueeStyle())
	{
		return 0.0;
	}
	
	const double dblPrev = GetPos();

	SetPos(dblPrev + GetStep(TRUE));
	return dblPrev;
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::CopyFrom(const CBCGPBaseVisualObject& srcObj)
{
	CBCGPCircularGaugeImpl::CopyFrom(srcObj);

	const CBCGPCircularProgressIndicatorImpl& src = (const CBCGPCircularProgressIndicatorImpl&)srcObj;

	m_strLabel = src.m_strLabel;
	m_Options = src.m_Options;
	m_ProgressColors = src.m_ProgressColors;
	m_brFocus = src.m_brFocus;
	m_brPressed = src.m_brPressed;
	m_bIsSmall = src.m_bIsSmall;
	m_dblStep = src.m_dblStep;
	
	SetColors(m_ProgressColors);
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::SetColors(CBCGPCircularGaugeColors::BCGP_CIRCULAR_GAUGE_COLOR_THEME theme)
{
	m_Colors.SetTheme(theme);
	
	if (theme == CBCGPCircularGaugeColors::BCGP_CIRCULAR_GAUGE_SILVER)
	{
		m_Colors.m_brPointerOutline = m_Colors.m_brText = CBCGPColor::DimGray;
	}

	m_ProgressColors.ConvertFrom(m_Colors);
	SetColors(m_ProgressColors);
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::SetColors(const CBCGPCircularProgressIndicatorColors& colors)
{
	m_ProgressColors = colors;	

	m_brFocus = m_ProgressColors.m_brProgressFill;
	
	m_brPressed = m_ProgressColors.m_brFill;
	if (m_brPressed.IsEmpty())
	{
		m_brPressed = CBCGPBrush(CBCGPColor::LightGray, 0.5);
	}
	else
	{
		if (m_brPressed.GetColor().IsDark())
		{
			m_brPressed.MakeLighter();
		}
		else
		{
			m_brPressed.MakeDarker(.05);
		}
	}
	
	if (m_arLevelBars.GetSize() == 0)
	{
		AddLevelBar(CBCGPBrush(), CBCGPBrush(), m_ProgressColors.m_brProgressFill, 0, max(0.0, m_Options.m_dblProgressWidth));
	}
	else
	{
		if (!m_ProgressColors.m_brProgressFill.IsEmpty())
		{
			GetLevelBar(0)->SetFillValueBrush(m_ProgressColors.m_brProgressFill);
		}
	}
	
	if (m_arSubGauges.GetSize() == 0)
	{
		CBCGPTextGaugeImpl* pLabel1 = new CBCGPTextGaugeImpl(OnPreparePercentageLabel(0), m_ProgressColors.m_brProgressLabel);
		CBCGPTextFormat tf1 = pLabel1->GetTextFormat();
		
		tf1.SetTextAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_CENTER);
		tf1.SetTextVerticalAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_CENTER);
		tf1.SetFontWeight(FW_ULTRALIGHT);
		
		pLabel1->SetTextFormat(tf1);
		
		AddSubGauge(pLabel1, CBCGPGaugeImpl::BCGP_SUB_GAUGE_NONE);
		
		CBCGPTextGaugeImpl* pLabel2 = new CBCGPTextGaugeImpl(m_strLabel, m_ProgressColors.m_brTextLabel);
		
		CBCGPTextFormat tf2 = pLabel2->GetTextFormat();
		
		tf2.SetTextAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_CENTER);
		tf2.SetTextVerticalAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_LEADING);
		tf2.SetWordWrap();
		tf2.SetFontWeight(FW_ULTRALIGHT);
		
		pLabel2->SetTextFormat(tf2);
		
		AddSubGauge(pLabel2, CBCGPGaugeImpl::BCGP_SUB_GAUGE_NONE);
	}
	else
	{
		CBCGPTextGaugeImpl* pLabel1 = DYNAMIC_DOWNCAST(CBCGPTextGaugeImpl, m_arSubGauges[0]);
		ASSERT_VALID(pLabel1);
		
		pLabel1->SetTextBrush(m_ProgressColors.m_brProgressLabel);
		
		CBCGPTextGaugeImpl* pLabel2 = DYNAMIC_DOWNCAST(CBCGPTextGaugeImpl, m_arSubGauges[1]);
		ASSERT_VALID(pLabel2);
		
		pLabel2->SetTextBrush(m_ProgressColors.m_brTextLabel);
	}
	
	m_Colors.m_brFill = m_ProgressColors.m_brFill;
	m_Colors.m_brFrameOutline = m_ProgressColors.m_brFrameOutline;

	m_Colors.m_brPointerFill = m_Colors.m_brPointerOutline = CBCGPBrush();

	SetDirty(TRUE, TRUE);
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::SetOptions(const CBCGPCircularProgressIndicatorOptions& options)
{
	if (m_Options.CompareWith(options))
	{
		return;
	}

	if (!options.m_bMarqueeStyle && m_bMarqueeStarted)
	{
		StartMarquee(FALSE);
	}

	const BOOL bPercFormatWasChanged = (m_Options.m_strPercentageFormat != options.m_strPercentageFormat);

	double start = 0.0;
	double finish = 100.0;

	GetRange(start, finish);

	m_Options = options;

	SetClosedRange(start, finish);
	SetLevelBarValue(IsMarqueeStyle() ? finish : start, 0, 0, TRUE);

	if (bPercFormatWasChanged && m_arSubGauges.GetSize() > 0)
	{
		CBCGPTextGaugeImpl* pLabel1 = DYNAMIC_DOWNCAST(CBCGPTextGaugeImpl, m_arSubGauges[0]);
		ASSERT_VALID(pLabel1);
		
		pLabel1->SetText(OnPreparePercentageLabel(GetPos()));
	}

	AdjustLayout(m_rect);
	SetDirty(TRUE, TRUE);
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::SetLabel(const CString& strLabel)
{
	if (m_strLabel == strLabel)
	{
		return;
	}
	
	m_strLabel = strLabel;
	
	CBCGPTextGaugeImpl* pLabel = DYNAMIC_DOWNCAST(CBCGPTextGaugeImpl, m_arSubGauges[1]);
	ASSERT_VALID(pLabel);
	
	pLabel->SetText(m_strLabel);
	SetDirty(TRUE, TRUE);
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::OnDraw(CBCGPGraphicsManager* pGM, const CBCGPRect& rectClip, DWORD dwFlags /*= BCGP_DRAW_STATIC | BCGP_DRAW_DYNAMIC*/)
{
	CBCGPBrush brFillSaved;

	if (m_bIsPressed)
	{
		brFillSaved = m_ProgressColors.m_brFill;
		m_Colors.m_brFill = m_brPressed;
	}

	CBCGPCircularGaugeImpl::OnDraw(pGM, rectClip, dwFlags);	

	if (m_bIsPressed)
	{
		m_Colors.m_brFill = brFillSaved;
	}

	if (!m_bIsInteractiveMode || (dwFlags & BCGP_DRAW_DYNAMIC) == 0)
	{
		return;
	}
	
	CBCGPCircularProgressIndicatorCtrl* pCtrl = DYNAMIC_DOWNCAST(CBCGPCircularProgressIndicatorCtrl, m_pWndOwner);
	if (pCtrl->GetSafeHwnd() == NULL || !pCtrl->IsFocused())
	{
		return;
	}
	
	CBCGPRect rect = m_rect;
	rect.DeflateRect(1.0, 1.0);
	
	const CBCGPPoint center = rect.CenterPoint();
	const double dblR = 0.5 * min(rect.Height(), rect.Width());
	
	CBCGPEllipse ellipse(center, dblR, dblR);
	
	CBCGPStrokeStyle strokeStyle;
	strokeStyle.SetDashStyle(CBCGPStrokeStyle::BCGP_DASH_STYLE_DASH);

	pGM->DrawEllipse(ellipse, m_brFocus, globalUtils.ScaleByDPI(1.0), &strokeStyle);
}
//*******************************************************************************
static BOOL CreateArcShape(CBCGPComplexGeometry& shape, BOOL bValue, double dblPerc, const CBCGPPoint& center, double r1, double r2,
						   double dblCos1, double dblSin1, double dblCos2, double dblSin2,
						   BOOL bClockwiseRotation, BOOL bIsLargeArc)
{
	BOOL bIsFullCircle = FALSE;

	if (bValue)
	{
		if (fabs(dblPerc) <= FLT_EPSILON)
		{
			return FALSE;
		}
		else if (fabs(dblPerc - 1.0) <= FLT_EPSILON)
		{
			bIsFullCircle = TRUE;
		}
	}
	else
	{
		if (fabs(dblPerc - 1.0) <= FLT_EPSILON)
		{
			return FALSE;
		}
		else if (fabs(dblPerc) <= FLT_EPSILON)
		{
			bIsFullCircle = TRUE;
		}

		bClockwiseRotation = !bClockwiseRotation;
		bIsLargeArc = !bIsLargeArc;
	}

	if (bIsFullCircle)
	{
		shape.SetFillMode(CBCGPGeometry::BCGP_FILL_MODE_ALTERNATE);
		
		shape.SetStart(CBCGPPoint(center.x - r1, center.y));
		shape.AddArc(CBCGPPoint(center.x + r1, center.y), CBCGPSize(r1, r1), TRUE, TRUE);
		shape.AddArc(CBCGPPoint(center.x - r1, center.y), CBCGPSize(r1, r1), TRUE, TRUE);
		
		shape.AddLine(CBCGPPoint(center.x - r2, center.y));
		shape.AddArc(CBCGPPoint(center.x + r2, center.y), CBCGPSize(r2, r2), TRUE, TRUE);
		shape.AddArc(CBCGPPoint(center.x - r2, center.y), CBCGPSize(r2, r2), TRUE, TRUE);
	}
	else
	{
		shape.SetStart(CBCGPPoint(center.x + dblCos1 * r1, center.y + dblSin1 * r1));
		
		shape.AddArc(
			CBCGPPoint(center.x + dblCos2 * r1, center.y + dblSin2 * r1),
			CBCGPSize(r1, r1), bClockwiseRotation, bIsLargeArc);
		
		shape.AddLine(CBCGPPoint(center.x + dblCos2 * r2, center.y + dblSin2 * r2));
		
		shape.AddArc(
			CBCGPPoint(center.x + dblCos1 * r2, center.y + dblSin1 * r2),
			CBCGPSize(r2, r2), !bClockwiseRotation, bIsLargeArc);
	}

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPCircularProgressIndicatorImpl::GetLevelBarShape(const CBCGPPoint& center, double radius, BOOL bForValue, CBCGPComplexGeometry& shape, double dblWidth, double dblOffsetFromFrame, int nScale, double dblValue)
{
	if (IsMarqueeStyle() && 
		(m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Gradient || m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Arc))
	{
		return CBCGPCircularGaugeImpl::GetLevelBarShape(center, radius, bForValue, shape, dblWidth, dblOffsetFromFrame, nScale, dblValue);
	}

	return TRUE;
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::OnDrawLevelBar(CBCGPGraphicsManager* pGM, const CBCGPComplexGeometry& shapeIn, const CBCGPBrush& /*brFill*/, const CBCGPBrush& /*brOutline*/, double /*lineWidth*/, BOOL bIsDynamic)
{
	if (!bIsDynamic || (IsMarqueeStyle() && !m_bMarqueeStarted))
	{
		return;
	}

	const double dblProgressWidth = m_arLevelBars.GetSize() > 0 ? m_arLevelBars[0]->GetWidth() : 0.0;
	if (dblProgressWidth == 0.0)
	{
		return;
	}

	const BOOL bClockwiseRotation = m_Options.m_bClockwiseRotation;

	double dblPerc = 0.0;
	double dblAngle = m_dblMarquee * 2.0 * M_PI;

	const BOOL bDrawNonCompleted = !IsMarqueeStyle() && !m_ProgressColors.m_brProgressFillNonComplete.IsEmpty() || !m_ProgressColors.m_brProgressOutlineNonComplete.IsEmpty();

	if (!IsMarqueeStyle())
	{
		double dblStart = GetStart();
		double dblFinish = GetFinish();
	
		double dblRange = fabs(dblFinish - dblStart);
		if (dblRange != 0.0)
		{
			dblPerc = (fabs(GetPos() - dblStart) / dblRange);
		}

		if (dblPerc == 0.0 && !bDrawNonCompleted)
		{
			return;
		}

		dblAngle = bClockwiseRotation ? -M_PI / 2.0 : M_PI / 2.0;

		if (m_Options.m_Shape != CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Gradient)
		{
			pGM->ReleaseClipArea();
		}
	}

	const double dblProgressWidth2 = dblProgressWidth / 2.0;

	CBCGPRect rect = m_rect;
	rect.DeflateRect(dblProgressWidth2, dblProgressWidth2);

	const CBCGPPoint center = rect.CenterPoint();

	double dblR = 0.5 * min(rect.Height(), rect.Width());

	const double dblMargin = (m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Gradient ||
		m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Arc) ? 0.0 : min(dblR / 20, globalUtils.ScaleByDPI(5.0));

	dblR -= dblMargin;
	if (dblR <= 0.0)
	{
		return;
	}

	if (m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Gradient)
	{
		CBCGPColor color = m_ProgressColors.m_brProgressFill.GetColor();
		
		CBCGPColor colorGradient = m_ProgressColors.m_brProgressFillGradient.GetColor();
		if (colorGradient.IsNull())
		{
			colorGradient = color;
			colorGradient.MakePale();
			
			color.MakeDarker(.2);
		}

		CBCGPBrushGradientStops stops;

		stops.Add(CBCGPBrushGradientStop(colorGradient));
		stops.Add(CBCGPBrushGradientStop(color, 1.0));

		CBCGPBrush br;

		if (!IsMarqueeStyle())
		{
			dblAngle += min(dblPerc, 0.99999) * M_PI * 2.0;;
		}

		const double dblCos = cos(dblAngle);
		const double dblSin = bClockwiseRotation ? sin(dblAngle) : -sin(dblAngle);

		br.SetLinearGradientStops(stops, 
			CBCGPGradientPoint(rect.CenterPoint().x - dblR * dblCos, rect.CenterPoint().y - dblR * dblSin, TRUE), 
			CBCGPGradientPoint(rect.CenterPoint().x + dblR * dblCos, rect.CenterPoint().y + dblR * dblSin, TRUE), CBCGPTransformsArray());
		
		br.SetOpacity(m_ProgressColors.m_brProgressFill.GetOpacity());

		pGM->FillGeometry(shapeIn, br);
		pGM->DrawGeometry(shapeIn, m_ProgressColors.m_brProgressOutline);
		return;
	}

	if (m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Arc)
	{
		const double dblCos = cos(dblAngle);
		const double dblSin = bClockwiseRotation ? sin(dblAngle) : -sin(dblAngle);

		CBCGPColor color = m_ProgressColors.m_brProgressFill.GetColor();
		CBCGPColor colorGradient = m_ProgressColors.m_brProgressFillGradient.GetColor();
		
		CBCGPBrush br;

		if (colorGradient.IsNull())
		{
			br = color;
		}
		else
		{
			CBCGPBrushGradientStops stops;
			
			stops.Add(CBCGPBrushGradientStop(color));
			stops.Add(CBCGPBrushGradientStop(colorGradient, 1.0));

			br.SetLinearGradientStops(stops, 
				CBCGPGradientPoint(rect.CenterPoint().x - dblR * dblCos, rect.CenterPoint().y - dblR * dblSin, TRUE), 
				CBCGPGradientPoint(rect.CenterPoint().x + dblR * dblCos, rect.CenterPoint().y + dblR * dblSin, TRUE), CBCGPTransformsArray());
		}
		
		br.SetOpacity(m_ProgressColors.m_brProgressFill.GetOpacity());

		const double r1 = dblR - dblProgressWidth2;
		const double r2 = dblR + dblProgressWidth2;

		if (IsMarqueeStyle())
		{
			dblPerc = 0.95;
		}

		const double dblAngle2 = dblAngle + min(dblPerc, 0.99999) * M_PI * 2.0;
		
		const double dblCos2 = cos(dblAngle2);
		const double dblSin2 = bClockwiseRotation ? sin(dblAngle2) : sin(-dblAngle2);
		
		const BOOL bIsLargeArc = dblPerc >= 0.5;

		CBCGPComplexGeometry shape;
		CBCGPComplexGeometry shapeNonComplete;

		if (bDrawNonCompleted && CreateArcShape(shapeNonComplete, FALSE, dblPerc, center, r1, r2, dblCos, dblSin, dblCos2, dblSin2, bClockwiseRotation, bIsLargeArc))
		{
			pGM->FillGeometry(shapeNonComplete, m_ProgressColors.m_brProgressFillNonComplete);
			pGM->DrawGeometry(shapeNonComplete, m_ProgressColors.m_brProgressOutlineNonComplete);
		}

		if (CreateArcShape(shape, TRUE, dblPerc, center, r1, r2, dblCos, dblSin, dblCos2, dblSin2, bClockwiseRotation, bIsLargeArc))
		{
			pGM->FillGeometry(shape, br);
			pGM->DrawGeometry(shape, m_ProgressColors.m_brProgressOutline);
		}
		
		return;
	}

	const double dblLen = 2.0 * M_PI * dblR;
	const double dblLineWidth = globalUtils.ScaleByDPI(m_bIsSmall ? 1.0 : 3.0);

	int nSteps = (int)(0.5 + dblLen / (1.5 * dblProgressWidth));

	if (m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Line)
	{
		nSteps = IsMarqueeStyle() ? (int)(0.5 + dblLen / (2.0 * dblLineWidth) - 1.0) : 50;
	}
	else if (m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Pie)
	{
		nSteps = max(4, (int)(0.5 + dblLen / 20));
	}

	if (nSteps < 4)
	{
		return;
	}
	
	double dblDeltaAngle = 2.0 * M_PI / nSteps;
	int iStart = 0;

	if (IsMarqueeStyle() && m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Pie && m_Options.m_bFadeSize)
	{
		iStart = 1;
	}

	const double dblPieDelta = M_PI / 100;
	const int nValuesSteps = (int)(dblPerc * nSteps);

	int nMaxSteps = nSteps;
	if (!IsMarqueeStyle() && !bDrawNonCompleted)
	{
		nMaxSteps = nValuesSteps;
	}

	for (int i = iStart; i < nMaxSteps; i++)
	{
		const double ratio = (double)(i + 1) / nSteps;
		const double dblCos = cos(dblAngle);
		const double dblSin = bClockwiseRotation ? sin(dblAngle) : -sin(dblAngle);

		CBCGPBrush brFillCurr;
		CBCGPBrush brOutline;

		if (!IsMarqueeStyle() && bDrawNonCompleted && i > nValuesSteps)
		{
			brFillCurr = m_ProgressColors.m_brProgressFillNonComplete;
			brOutline = m_ProgressColors.m_brProgressOutlineNonComplete;
		}
		else
		{
			brFillCurr = m_ProgressColors.m_brProgressFill;

			if (!m_ProgressColors.m_brProgressFillGradient.IsEmpty())
			{
				brFillCurr.Mix(m_ProgressColors.m_brProgressFillGradient, ratio);
			}

			brOutline = m_ProgressColors.m_brProgressOutline;

			if (!m_ProgressColors.m_brProgressOutlineGradient.IsEmpty())
			{
				brOutline.Mix(m_ProgressColors.m_brProgressOutlineGradient, ratio);
			}

			if (m_Options.m_bFadeColor)
			{
				brFillCurr.SetOpacity(min(1., 0.15 + 0.85 * ratio * brFillCurr.GetOpacity()));
				brOutline.SetOpacity(min(1., 0.15 + 0.85 * ratio * brOutline.GetOpacity()));
			}
		}

		double r = dblProgressWidth2;
		if (m_Options.m_bFadeSize)
		{
			r *= (0.3 + 0.7 * ratio);
		}

		if (m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Circle)
		{
			double x = center.x + dblCos * dblR;
			double y = center.y + dblSin * dblR;

			CBCGPEllipse circle(CBCGPPoint(x, y), r, r);

			pGM->FillEllipse(circle, brFillCurr);
			pGM->DrawEllipse(circle, brOutline);
		}
		else if (m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Line)
		{
			if ((i % 2) == 0)
			{
				const double dblOffsetCenter = m_bIsSmall ? globalUtils.ScaleByDPI(3.0) : dblR - r;

				double x1 = center.x + dblCos * dblOffsetCenter;
				double y1 = center.y + dblSin * dblOffsetCenter;

				double x2 = center.x + dblCos * (dblR + r);
				double y2 = center.y + dblSin * (dblR + r);

				pGM->DrawLine(x1, y1, x2, y2, brFillCurr, dblLineWidth);
			}
		}
		else if (m_Options.m_Shape == CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Pie)
		{
			CBCGPComplexGeometry shape;

			const double r1 = dblR - r;
			const double r2 = dblR + r;

			const double dblCos2 = cos(dblAngle + dblDeltaAngle - dblPieDelta);
			const double dblSin2 = bClockwiseRotation ? sin(dblAngle + dblDeltaAngle - dblPieDelta) : -sin(dblAngle + dblDeltaAngle - dblPieDelta);

			const BOOL bIsLargeArc = FALSE;

			shape.SetStart(CBCGPPoint(center.x + dblCos * r1, center.y + dblSin * r1));
			
			shape.AddArc(
				CBCGPPoint(center.x + dblCos2 * r1, center.y + dblSin2 * r1),
				CBCGPSize(r1, r1), bClockwiseRotation, bIsLargeArc);

			shape.AddLine(CBCGPPoint(center.x + dblCos2 * r2, center.y + dblSin2 * r2));

			shape.AddArc(
				CBCGPPoint(center.x + dblCos * r2, center.y + dblSin * r2),
				CBCGPSize(r2, r2), !bClockwiseRotation, bIsLargeArc);

			pGM->FillGeometry(shape, brFillCurr);
			pGM->DrawGeometry(shape, brOutline);
		}

		dblAngle += dblDeltaAngle;
	}
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::OnAnimation(CBCGPVisualDataObject* /*pDataObject*/)
{
	m_dblMarquee += 0.01;
	if (m_dblMarquee > 1.0)
	{
		m_dblMarquee = 0.0;
	}

	Redraw();

	if (m_pRenderer != NULL)
	{
		m_pRenderer->m_pCircularProgressWndOwner->RedrawWindow(m_pRenderer->m_rectCircularProgress.IsRectEmpty() ? NULL : &m_pRenderer->m_rectCircularProgress);
	}
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::OnAnimationFinished(CBCGPVisualDataObject* /*pDataObject*/)
{
	m_dblMarquee = 0.0;
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::SetRect(const CBCGPRect& rect, BOOL bRedraw /*= FALSE*/)
{
	CBCGPCircularGaugeImpl::SetRect(rect, bRedraw);	
	AdjustLayout(rect);
}
//*******************************************************************************
BOOL CBCGPCircularProgressIndicatorImpl::ReposAllSubGauges(CBCGPGraphicsManager* /*pGM*/)
{
	CBCGPRect rect = m_rect;
	const double dblProgressWidth = m_arLevelBars.GetSize() > 0 ? m_arLevelBars[0]->GetWidth() + 2.0 : 0.0;
	rect.DeflateRect(dblProgressWidth, dblProgressWidth);

	const double dblSize = min(rect.Height(), rect.Width());

	double left = rect.CenterPoint().x - dblSize / 2;
	double y = rect.CenterPoint().y - dblSize / 2;
	double height = dblSize;
	double width = dblSize;

	BOOL bNoProgress = IsMarqueeStyle() || m_Options.m_strPercentageFormat.IsEmpty();

	for (int i = 0; i < (int)m_arSubGauges.GetSize(); i++)
	{
		CBCGPTextGaugeImpl* pLabel = DYNAMIC_DOWNCAST(CBCGPTextGaugeImpl, m_arSubGauges[i]);
		ASSERT_VALID(pLabel);

		BOOL bHide = FALSE;

		if (i == 0)
		{
			if (bNoProgress || m_bIsSmall)
			{
				bHide = TRUE;
			}
		}
		else
		{
			if (m_strLabel.IsEmpty() || m_bIsSmall)
			{
				bHide = TRUE;
			}
		}

		pLabel->SetVisible(!bHide);

		if (!bHide)
		{
			CBCGPRect rectLabel(CBCGPPoint(left, y), CBCGPSize(width, height));

			if (i == 0 && !m_strLabel.IsEmpty())
			{
				height = dblSize / 3;
				width = dblSize / 2;
				
				left = rect.CenterPoint().x - width / 2;

				rectLabel.bottom -= height;
				y = rectLabel.bottom;
			}

			pLabel->SetRect(rectLabel);

			CBCGPTextFormat tf = pLabel->GetTextFormat();

			double ratio = (i == 0) ? 0.4 : 0.3;

			if (i == 1)
			{
				if (bNoProgress)
				{
					ratio = 0.1;
					tf.SetTextVerticalAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_CENTER);
				}
				else
				{
					tf.SetTextVerticalAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_LEADING);
				}
			}

			tf.SetFontSize((float)(ratio * rectLabel.Height()));
			
			pLabel->SetTextFormat(tf);
		}
	}

	return FALSE;
}
//*******************************************************************************
BOOL CBCGPCircularProgressIndicatorImpl::OnMouseDown(int nButton, const CBCGPPoint& pt)
{
	if (m_bIsPressed)
	{
		return TRUE;
	}

	m_bIsPressed = FALSE;

	if (!m_bIsInteractiveMode || nButton != 0)
	{
		return CBCGPCircularGaugeImpl::OnMouseDown(nButton, pt);
	}

	CBCGPRect rect = m_rect;
	rect.OffsetRect(-m_ptScrollOffset);
	
	const double radius = min(rect.Height(), rect.Width()) / 2.0;
	const CBCGPPoint center = rect.CenterPoint();
	
	const double dblDistanceToCenter = bcg_distance(pt, center);
	
	if (dblDistanceToCenter <= radius)
	{
		m_bIsPressed = TRUE;
		SetDirty(TRUE, TRUE);
		return TRUE;
	}

	return CBCGPCircularGaugeImpl::OnMouseDown(nButton, pt);
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::OnMouseUp(int nButton, const CBCGPPoint& pt)
{
	if (!m_bIsInteractiveMode || !m_bIsPressed)
	{
		CBCGPCircularGaugeImpl::OnMouseUp(nButton, pt);
		return;
	}

	if (nButton != 0)
	{
		return;
	}
	
	m_bIsPressed = FALSE;
	SetDirty(TRUE, TRUE);

	CBCGPRect rect = m_rect;
	rect.OffsetRect(-m_ptScrollOffset);
	
	const double radius = min(rect.Height(), rect.Width()) / 2.0;
	const CBCGPPoint center = rect.CenterPoint();
	
	const double dblDistanceToCenter = bcg_distance(pt, center);
	
	if (dblDistanceToCenter <= radius)
	{
		FireClickEvent(pt);
	}
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::OnMouseMove(const CBCGPPoint& /*pt*/)
{
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::OnMouseLeave()
{
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::OnCancelMode()
{
	if (!m_bIsPressed)
	{
		CBCGPCircularGaugeImpl::OnCancelMode();
		return;
	}

	m_bIsPressed = FALSE;
	SetDirty(TRUE, TRUE);
}
//*******************************************************************************
BOOL CBCGPCircularProgressIndicatorImpl::OnKeyboardDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (nChar == VK_SPACE || nChar == VK_RETURN)
	{
		FireClickEvent(CBCGPPoint(-1.0, -1.0));
		return TRUE;
	}

	return CBCGPCircularGaugeImpl::OnKeyboardDown(nChar, nRepCnt, nFlags);
}
//*******************************************************************************
void CBCGPCircularProgressIndicatorImpl::AdjustLayout(const CBCGPRect& rect)
{
	const double dblSize = min(rect.Height(), rect.Width());

	m_bIsSmall = dblSize > 0.0 && dblSize < globalUtils.ScaleByDPI(24.0);

	if (m_arLevelBars.GetSize() != 0)
	{
		double dblWidth = max(0.0, m_Options.m_dblProgressWidth);

		if (dblWidth == 0.0)
		{
			double dblPart = 15.0;

			switch (m_Options.m_Shape)
			{
			case CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Gradient:
			case CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Arc:
				dblPart = 30.0;
				break;

			case CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Line:
			case CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Pie:
				dblPart = m_bIsSmall ? 5.0 : 10.0;
				break;

			case CBCGPCircularProgressIndicatorOptions::BCGPCircularProgressIndicator_Circle:
				dblPart = m_bIsSmall ? 5.0 : 15.0;
				break;
			}

			dblWidth = max(globalUtils.ScaleByDPI(2.0), dblSize / dblPart);
		}
		else if (dblWidth > 0.0 && dblWidth < 1.0)
		{
			dblWidth *= dblSize / 4.0;
		}
		
		m_arLevelBars[0]->SetWidth(dblWidth);
	}
}
//*******************************************************************************
CString CBCGPCircularProgressIndicatorImpl::OnPreparePercentageLabel(double nPos) const
{
	CString strLabel;
	double dblPerc = 0;
	
	double dblStart = GetStart();
	double dblFinish = GetFinish();
	
	double dblRange = fabs(dblFinish - dblStart);
	if (dblRange != 0.0)
	{
		dblPerc = 100.0 * fabs(nPos - dblStart) / dblRange;
	}
	
	strLabel.Format(m_Options.m_strPercentageFormat, dblPerc);
	return strLabel;
}
//*******************************************************************************
HRESULT CBCGPCircularProgressIndicatorImpl::get_accRole(VARIANT varChild, VARIANT *pvarRole)
{
	HRESULT hr = CBCGPCircularGaugeImpl::get_accRole(varChild, pvarRole);
	if (SUCCEEDED(hr))
	{
		pvarRole->lVal = ROLE_SYSTEM_PROGRESSBAR;
	}
	
	return hr;
}
//*******************************************************************************
HRESULT CBCGPCircularProgressIndicatorImpl::get_accValue(VARIANT varChild, BSTR *pszValue)
{
	if (pszValue == NULL)
	{
		return E_INVALIDARG;
	}
	
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		CString strValue;
		
		if (IsMarqueeStyle())
		{
			strValue = m_strLabel;
		}
		else
		{
			if (m_Options.m_strPercentageFormat.IsEmpty())
			{
				strValue.Format(_T("%f"), GetPos());
			}
			else
			{
				strValue = OnPreparePercentageLabel(GetPos());
			}
		}
		
		if (!strValue.IsEmpty())
		{
			*pszValue = strValue.AllocSysString();
			return S_OK;
		}
	}
	
	return S_FALSE;
}
//*******************************************************************************
HRESULT CBCGPCircularProgressIndicatorImpl::get_accState(VARIANT varChild, VARIANT *pvarState)
{
	if (m_bIsInteractiveMode && (varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		pvarState->vt = VT_I4;
		
		if (m_bIsInteractiveMode)
		{
			pvarState->lVal = STATE_SYSTEM_FOCUSABLE;
			pvarState->lVal |= STATE_SYSTEM_SELECTABLE;

			if (m_bIsPressed)
			{
				pvarState->lVal |= STATE_SYSTEM_PRESSED;
			}
			
			CBCGPCircularProgressIndicatorCtrl* pCtrl = DYNAMIC_DOWNCAST(CBCGPCircularProgressIndicatorCtrl, m_pWndOwner);
			if (pCtrl->GetSafeHwnd() != NULL && pCtrl->IsFocused())
			{
				pvarState->lVal |= STATE_SYSTEM_FOCUSED;
				pvarState->lVal |= STATE_SYSTEM_SELECTED;
			}
		}
		else
		{
			pvarState->lVal = STATE_SYSTEM_READONLY;
		}

		return S_OK;
	}
	
	return S_FALSE;
}
