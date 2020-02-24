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
// BCGPGaugeImpl.cpp: implementation of the CBCGPGaugeImpl class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BCGPGaugeImpl.h"
#include "BCGPMath.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

UINT BCGM_ON_GAUGE_START_TRACK = ::RegisterWindowMessage (_T("BCGM_ON_GAUGE_START_TRACK"));
UINT BCGM_ON_GAUGE_TRACK = ::RegisterWindowMessage (_T("BCGM_ON_GAUGE_TRACK"));
UINT BCGM_ON_GAUGE_FINISH_TRACK = ::RegisterWindowMessage (_T("BCGM_ON_GAUGE_FINISH_TRACK"));
UINT BCGM_ON_GAUGE_CANCEL_TRACK = ::RegisterWindowMessage (_T("BCGM_ON_GAUGE_CANCEL_TRACK"));
UINT BCGM_ON_GAUGE_CLICK = ::RegisterWindowMessage (_T("BCGM_ON_GAUGE_CLICK"));

IMPLEMENT_DYNAMIC(CBCGPGaugeImpl, CBCGPBaseVisualObject)
IMPLEMENT_DYNCREATE(CBCGPGaugeDataObject, CBCGPVisualDataObject)
IMPLEMENT_DYNCREATE(CBCGPGaugeScaleObject, CObject)
IMPLEMENT_DYNCREATE(CBCGPGaugeColoredRangeObject, CObject)
IMPLEMENT_DYNCREATE(CBCGPGaugeLevelBarObject, CObject)

void CBCGPGaugeLevelBarObject::OnAnimationValueChanged(double dblOldValue, double dblNewValue)
{
	UNREFERENCED_PARAMETER(dblOldValue);
	UNREFERENCED_PARAMETER(dblNewValue);

	if (m_pParentGauge != NULL)
	{
		m_pParentGauge->Redraw();
	}
}

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPGaugeImpl::CBCGPGaugeImpl(CBCGPVisualContainer* pContainer) :
	CBCGPBaseVisualObject(pContainer)	
{
	m_nFrameSize = 6;
	m_Pos = BCGP_SUB_GAUGE_NONE;
	m_bIsSubGauge = FALSE;
	m_nTrackingScale = -1;

	m_brShadow.SetColor(CBCGPColor(.1, .1, .1, .2));
}
//*******************************************************************************
CBCGPGaugeImpl::~CBCGPGaugeImpl()
{
	RemoveAllColoredRanges();
	RemoveAllScales();
	RemoveAllLevelBars();
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::SetValue(double dblVal, int nIndex, UINT uiAnimationTime, BOOL bRedraw)
{
	if (nIndex < 0 || nIndex >= m_arData.GetSize())
	{
		ASSERT(FALSE);
		return FALSE;
	}

	CBCGPGaugeDataObject* pData = DYNAMIC_DOWNCAST(CBCGPGaugeDataObject, m_arData[nIndex]);
	ASSERT_VALID(pData);

	const int nScale = pData->GetScale();
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return FALSE;
	}

	CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
	ASSERT_VALID(pScale);

	if (dblVal < min(pScale->m_dblStart, pScale->m_dblFinish) || 
		dblVal > max(pScale->m_dblStart, pScale->m_dblFinish))
	{
		return FALSE;
	}

	return CBCGPBaseVisualObject::SetValue(dblVal, nIndex, uiAnimationTime, bRedraw);
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::SetLevelBarValue(double dblVal, int nIndex /*= 0*/, UINT uiAnimationTime /*= 100 /* Ms, 0 - no animation */, BOOL bRedraw /*= FALSE*/)
{
	CBCGPGaugeLevelBarObject* pLevelBar = GetLevelBar(nIndex);
	if (pLevelBar == NULL)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	ASSERT_VALID(pLevelBar);

	const int nScale = pLevelBar->GetScale();
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return FALSE;
	}
	
	CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
	ASSERT_VALID(pScale);
	
	if (dblVal < min(pScale->m_dblStart, pScale->m_dblFinish) || 
		dblVal > max(pScale->m_dblStart, pScale->m_dblFinish))
	{
		return FALSE;
	}

	const double dblRange = pScale->m_dblFinish - pScale->m_dblStart;

	pLevelBar->SetValue(dblVal, 0.0001 * uiAnimationTime * dblRange);

	if (bRedraw)
	{
		Redraw();
	}

	return TRUE;
}
//*******************************************************************************
int CBCGPGaugeImpl::AddColoredRange(double dblStartValue, double dblFinishValue,
		const CBCGPBrush& brFill, const CBCGPBrush& brFrame,
		int nScale,
		double dblStartWidth,
		double dblFinishWidth,
		double dblOffsetFromFrame,
		const CBCGPBrush& brTextLabel, 
		const CBCGPBrush& brTickMarkOutline, const CBCGPBrush& brTickMarkFill,
		BOOL bRedraw)
{
	return AddColoredRange(new CBCGPGaugeColoredRangeObject(
		dblStartValue, dblFinishValue, brFill, brFrame, nScale, 
		dblStartWidth, dblFinishWidth, dblOffsetFromFrame,
		brTextLabel, brTickMarkOutline, brTickMarkFill), bRedraw);
}
//*******************************************************************************
int CBCGPGaugeImpl::AddColoredRange(CBCGPGaugeColoredRangeObject* pRange, BOOL bRedraw)
{
	ASSERT_VALID(pRange);

	pRange->SetParentGauge (this);
	m_arRanges.Add(pRange);

	if (bRedraw)
	{
		Redraw();
	}

	return (int)m_arRanges.GetSize() - 1;
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::ModifyColoredRange(int nIndex, const CBCGPGaugeColoredRangeObject& range, BOOL bRedraw)
{
	if (nIndex < 0 || nIndex >= m_arRanges.GetSize())
	{
		return FALSE;
	}

	CBCGPGaugeColoredRangeObject* pRange = m_arRanges[nIndex];
	ASSERT_VALID(pRange);

	*pRange = range;

	pRange->m_Shape.Clear();
	pRange->m_Shape.Destroy();

	SetDirty();

	if (bRedraw)
	{
		Redraw();
	}

	return TRUE;
}
//*******************************************************************************
void CBCGPGaugeImpl::RemoveAllScales(BOOL bRedraw)
{
	for (int i = 0; i < m_arScales.GetSize(); i++)
	{
		CBCGPGaugeScaleObject* pScale = m_arScales[i];
		ASSERT_VALID(pScale);
		
		delete pScale;
	}

	m_arScales.RemoveAll();

	SetDirty();
	
	if (bRedraw)
	{
		Redraw();
	}
}
//*******************************************************************************
void CBCGPGaugeImpl::RemoveAllColoredRanges(BOOL bRedraw)
{
	for (int i = 0; i < m_arRanges.GetSize(); i++)
	{
		CBCGPGaugeColoredRangeObject* pRange = m_arRanges[i];
		ASSERT_VALID(pRange);

		delete pRange;
	}

	m_arRanges.RemoveAll();

	SetDirty();

	if (bRedraw)
	{
		Redraw();
	}
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::RemoveColoredRange(int nIndex, BOOL bRedraw)
{
	if (nIndex < 0 || nIndex >= m_arRanges.GetSize())
	{
		return FALSE;
	}

	CBCGPGaugeColoredRangeObject* pRange = m_arRanges[nIndex];
	ASSERT_VALID(pRange);

	delete pRange;
	m_arRanges.RemoveAt(nIndex);

	SetDirty();

	if (bRedraw)
	{
		Redraw();
	}

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::RemoveScale(int nIndex, BOOL bRedraw)
{
	if (nIndex < 0 || nIndex >= m_arScales.GetSize())
	{
		return FALSE;
	}

	CBCGPGaugeScaleObject* pScale = m_arScales[nIndex];
	ASSERT_VALID(pScale);

	delete pScale;
	m_arScales.RemoveAt(nIndex);

	SetDirty();

	if (bRedraw)
	{
		Redraw();
	}

	return TRUE;
}
//*******************************************************************************
const CBCGPGaugeColoredRangeObject* CBCGPGaugeImpl::GetColoredRangeByValue(double dblVal, int nScale, DWORD dwFlags) const
{
	for (int i = 0; i < m_arRanges.GetSize(); i++)
	{
		CBCGPGaugeColoredRangeObject* pRange = m_arRanges[i];
		ASSERT_VALID(pRange);

		if (pRange->m_nScale != nScale)
		{
			continue;
		}

		if ((dwFlags & BCGP_GAUGE_RANGE_TEXT_COLOR) == BCGP_GAUGE_RANGE_TEXT_COLOR &&
			pRange->m_brTextLabel.IsEmpty())
		{
			continue;
		}

		if ((dwFlags & BCGP_GAUGE_RANGE_TICKMARK_COLOR) == BCGP_GAUGE_RANGE_TICKMARK_COLOR &&
			pRange->m_brTickMarkFill.IsEmpty() && pRange->m_brTickMarkOutline.IsEmpty())
		{
			continue;
		}

		double dblMin = min(pRange->m_dblStartValue, pRange->m_dblFinishValue);
		double dblMax = max(pRange->m_dblStartValue, pRange->m_dblFinishValue);

		if (dblVal >= dblMin && dblVal <= dblMax)
		{
			return pRange;
		}
	}

	return NULL;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetRange(double dblStart, double dblFinish, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
	ASSERT_VALID(pScale);

	pScale->m_dblStart = dblStart;
	pScale->m_dblFinish = dblFinish;

	for (int i = 0; i < (int)m_arData.GetSize(); i++)
	{
		CBCGPGaugeDataObject* pData = DYNAMIC_DOWNCAST(CBCGPGaugeDataObject, m_arData[i]);
		if (pData != NULL && pData->GetScale() == nScale)
		{
			double dblVal = pData->GetValue();

			if (dblVal < dblStart)
			{
				pData->SetValue(dblStart);
			}
			else if (dblVal > dblFinish)
			{
				pData->SetValue(dblFinish);
			}
		}
	}

	m_arScales[nScale]->CleanUp();

	SetDirty();
}
//*******************************************************************************
double CBCGPGaugeImpl::GetStart(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return 0.;
	}

	return m_arScales[nScale]->m_dblStart;
}
//*******************************************************************************
double CBCGPGaugeImpl::GetFinish(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return 0.;
	}

	return m_arScales[nScale]->m_dblFinish;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetStep(double dblStep, int nScale)
{
	ASSERT(dblStep != 0.);

	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	m_arScales[nScale]->m_dblStep = fabs(dblStep);
	m_arScales[nScale]->CleanUp();

	SetDirty();
}
//*******************************************************************************
double CBCGPGaugeImpl::GetStep(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return 0.;
	}

	return m_arScales[nScale]->m_dblStep;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetTextLabelFormat(const CString& strFormat, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	m_arScales[nScale]->m_strLabelFormat = strFormat;

	SetDirty();
}
//*******************************************************************************
CString CBCGPGaugeImpl::GetTextLabelFormat(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return _T("");
	}

	return m_arScales[nScale]->m_strLabelFormat;
}
//*******************************************************************************
double CBCGPGaugeImpl::GetTickMarkSize(BOOL bMajor, int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return -1.;
	}

	return bMajor ? m_arScales[nScale]->m_dblMajorTickMarkSize : m_arScales[nScale]->m_dblMinorTickMarkSize;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetTickMarkSize(double dblSize, BOOL bMajor, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	if (bMajor)
	{
		m_arScales[nScale]->m_dblMajorTickMarkSize = dblSize;
	}
	else
	{
		m_arScales[nScale]->m_dblMinorTickMarkSize = dblSize;
	}

	m_arScales[nScale]->CleanUp();

	SetDirty();
}
//*******************************************************************************
void CBCGPGaugeImpl::SetMajorTickMarkStep(double dblMajorTickMarkStep, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	m_arScales[nScale]->m_dblMajorTickMarkStep = dblMajorTickMarkStep;
	m_arScales[nScale]->CleanUp();

	SetDirty();
}
//*******************************************************************************
double CBCGPGaugeImpl::GetMajorTickMarkStep(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return -1.;
	}

	return m_arScales[nScale]->m_dblMajorTickMarkStep;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetScaleFillBrush(const CBCGPBrush& brFill, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	m_arScales[nScale]->m_brFill = brFill;

	SetDirty();
}
//*******************************************************************************
CBCGPBrush CBCGPGaugeImpl::GetScaleFillBrush(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return CBCGPBrush();
	}

	return m_arScales[nScale]->m_brFill;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetScaleOutlineBrush(const CBCGPBrush& brOutline, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	m_arScales[nScale]->m_brOutline = brOutline;

	SetDirty();
}
//*******************************************************************************
CBCGPBrush CBCGPGaugeImpl::GetScaleOutlineBrush(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return CBCGPBrush();
	}

	return m_arScales[nScale]->m_brOutline;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetScaleTextBrush(const CBCGPBrush& brText, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	m_arScales[nScale]->m_brText = brText;

	SetDirty();
}
//*******************************************************************************
CBCGPBrush CBCGPGaugeImpl::GetScaleTextBrush(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return CBCGPBrush();
	}

	return m_arScales[nScale]->m_brText;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetScaleTickMarkBrush(const CBCGPBrush& brTickMark, BOOL bMajor, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	if (bMajor)
	{
		m_arScales[nScale]->m_brTickMarkMajor = brTickMark;
	}
	else
	{
		m_arScales[nScale]->m_brTickMarkMinor = brTickMark;
	}

	SetDirty();
}
//*******************************************************************************
CBCGPBrush CBCGPGaugeImpl::GetScaleTickMarkBrush(BOOL bMajor, int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return CBCGPBrush();
	}

	return bMajor ? m_arScales[nScale]->m_brTickMarkMajor : m_arScales[nScale]->m_brTickMarkMinor;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetScaleTickMarkOutlineBrush(const CBCGPBrush& brTickMark, BOOL bMajor, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	if (bMajor)
	{
		m_arScales[nScale]->m_brTickMarkMajorOutline = brTickMark;
	}
	else
	{
		m_arScales[nScale]->m_brTickMarkMinorOutline = brTickMark;
	}

	SetDirty();
}
//*******************************************************************************
CBCGPBrush CBCGPGaugeImpl::GetScaleTickMarkOutlineBrush(BOOL bMajor, int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return CBCGPBrush();
	}

	return bMajor ? m_arScales[nScale]->m_brTickMarkMajorOutline : m_arScales[nScale]->m_brTickMarkMinorOutline;
}
//*******************************************************************************
double CBCGPGaugeImpl::GetAnimationRange(int nIndex)
{
	if (nIndex < 0 || nIndex >= m_arData.GetSize())
	{
		ASSERT(FALSE);
		return 0.;
	}

	CBCGPGaugeDataObject* pData = DYNAMIC_DOWNCAST(CBCGPGaugeDataObject, m_arData[nIndex]);
	ASSERT_VALID(pData);

	const int nScale = pData->GetScale();
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return 0.;
	}

	CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
	ASSERT_VALID(pScale);

	return pScale->m_dblFinish - pScale->m_dblStart;
}
//*******************************************************************************
CBCGPSize CBCGPGaugeImpl::GetTickMarkTextLabelSize(CBCGPGraphicsManager* pGM, const CString& strLabel,
												   const CBCGPTextFormat& tf)
{
	ASSERT_VALID(pGM);
	return pGM->GetTextSize(strLabel, tf);
}
//*******************************************************************************
void CBCGPGaugeImpl::OnDrawTickMarkTextLabel(
	CBCGPGraphicsManager* pGM, 
	const CBCGPTextFormat& tf, const CBCGPRect& rectText, const CString& strLabel, double dblVal, int nScale, const CBCGPBrush& br)
{
	ASSERT_VALID(pGM);

	const CBCGPGaugeColoredRangeObject* pRange = GetColoredRangeByValue(dblVal, nScale, BCGP_GAUGE_RANGE_TEXT_COLOR);

	if (pRange != NULL && !pRange->m_brTextLabel.IsEmpty())
	{
		pGM->DrawText(strLabel, rectText, tf, pRange->m_brTextLabel);
	}
	else
	{
		pGM->DrawText(strLabel, rectText, tf, br);
	}
}
//*******************************************************************************
void CBCGPGaugeImpl::SetScaleOffsetFromFrame(double dblOffsetFromFrame, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
	ASSERT_VALID(pScale);

	pScale->m_dblOffsetFromFrame = dblOffsetFromFrame;
	pScale->CleanUp();
}
//*******************************************************************************
double CBCGPGaugeImpl::GetScaleOffsetFromFrame(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return 0.;
	}

	CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
	ASSERT_VALID(pScale);

	return pScale->m_dblOffsetFromFrame;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetTickMarkStyle(CBCGPGaugeScaleObject::BCGP_TICKMARK_STYLE style, 
											 BOOL bIsMajor, double dblSize, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
	ASSERT_VALID(pScale);

	if (bIsMajor)
	{
		pScale->m_MajorTickMarkStyle = style;
	}
	else
	{
		pScale->m_MinorTickMarkStyle = style;
	}

	if (dblSize >= 0.)
	{
		if (bIsMajor)
		{
			pScale->m_dblMajorTickMarkSize = dblSize;
		}
		else
		{
			pScale->m_dblMinorTickMarkSize = dblSize;
		}
	}

	SetDirty();
}
//*******************************************************************************
CBCGPGaugeScaleObject::BCGP_TICKMARK_STYLE CBCGPGaugeImpl::GetTickMarkStyle(BOOL bIsMajor, int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return (CBCGPGaugeScaleObject::BCGP_TICKMARK_STYLE)-1;
	}

	CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
	ASSERT_VALID(pScale);

	return bIsMajor ? pScale->m_MajorTickMarkStyle : pScale->m_MinorTickMarkStyle;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetMinorTickMarkPosition(CBCGPGaugeScaleObject::BCGP_TICKMARK_POSITION pos, int nScale)
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return;
	}

	CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
	ASSERT_VALID(pScale);

	if (pScale->m_MinorTickMarkPosition != pos)
	{
		pScale->m_MinorTickMarkPosition = pos;

		SetDirty();
	}
}
//*******************************************************************************
CBCGPGaugeScaleObject::BCGP_TICKMARK_POSITION CBCGPGaugeImpl::GetMinorTickMarkPosition(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		ASSERT(FALSE);
		return (CBCGPGaugeScaleObject::BCGP_TICKMARK_POSITION)-1;
	}

	CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
	ASSERT_VALID(pScale);

	return pScale->m_MinorTickMarkPosition;
}
//*******************************************************************************
CBCGPGaugeScaleObject* CBCGPGaugeImpl::GetScale(int nScale) const
{
	if (nScale < 0 || nScale >= m_arScales.GetSize())
	{
		return NULL;
	}

	return m_arScales[nScale];
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::OnMouseDown(int nButton, const CBCGPPoint& pt)
{
	if (m_bIsInteractiveMode && nButton == 0)
	{
		double dblVal = 0.;
		m_nTrackingScale = 0;	// Temp!

		if (m_nTrackingScale >= 0 && m_nTrackingScale < m_arData.GetSize())
		{
			if (HitTestValue(pt, dblVal, m_nTrackingScale))
			{
				if (!FireTrackEvent(BCGM_ON_GAUGE_START_TRACK, dblVal, pt))
				{
					SetValue(dblVal, 0, 0, TRUE);

					if (globalData.IsAccessibilitySupport())
					{
#ifdef _BCGSUITE_
						::NotifyWinEvent(EVENT_OBJECT_VALUECHANGE, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
#else
						globalData.NotifyWinEvent(EVENT_OBJECT_VALUECHANGE, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
#endif
					}
				}

				return TRUE;
			}
			else
			{
				m_nTrackingScale = -1;
			}
		}
		else
		{
			m_nTrackingScale = -1;
		}
	}

	return CBCGPBaseVisualObject::OnMouseDown(nButton, pt);
}
//*******************************************************************************
void CBCGPGaugeImpl::ClampValueToScaleRange(double& dblVal, int nScale) const
{
	ASSERT_VALID(this);

	CBCGPGaugeScaleObject* pScale = GetScale(nScale);
	if (pScale != NULL)
	{
		ASSERT_VALID(pScale);
		dblVal = bcg_clamp(dblVal, min(pScale->m_dblStart, pScale->m_dblFinish), max(pScale->m_dblStart, pScale->m_dblFinish));
	}
}
//*******************************************************************************
void CBCGPGaugeImpl::OnMouseUp(int nButton, const CBCGPPoint& pt)
{
	if (m_bIsInteractiveMode && m_nTrackingScale >= 0 && nButton == 0)
	{
		double dblVal = 0.;

		if (HitTestValue(pt, dblVal, m_nTrackingScale, FALSE))
		{
			ClampValueToScaleRange(dblVal, m_nTrackingScale);
			FireTrackEvent(BCGM_ON_GAUGE_FINISH_TRACK, dblVal, pt);

			SetValue(dblVal, 0, 0, TRUE);
			m_nTrackingScale = -1;

			if (globalData.IsAccessibilitySupport())
			{
#ifdef _BCGSUITE_
				::NotifyWinEvent(EVENT_OBJECT_VALUECHANGE, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
#else
				globalData.NotifyWinEvent (EVENT_OBJECT_VALUECHANGE, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
#endif
			}
			return;
		}

		m_nTrackingScale = -1;
	}

	CBCGPBaseVisualObject::OnMouseUp(nButton, pt);
}
//*******************************************************************************
void CBCGPGaugeImpl::OnMouseMove(const CBCGPPoint& pt)
{
	if (m_bIsInteractiveMode && m_nTrackingScale >= 0)
	{
		double dblVal = 0.;

		if (HitTestValue(pt, dblVal, m_nTrackingScale, FALSE))
		{
			ClampValueToScaleRange(dblVal, m_nTrackingScale);
			FireTrackEvent(BCGM_ON_GAUGE_TRACK, dblVal, pt);

			SetValue(dblVal, 0, 0, TRUE);

			if (globalData.IsAccessibilitySupport())
			{
#ifdef _BCGSUITE_
				::NotifyWinEvent(EVENT_OBJECT_VALUECHANGE, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
#else
				globalData.NotifyWinEvent (EVENT_OBJECT_VALUECHANGE, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
#endif
			}
			return;
		}
	}

	CBCGPBaseVisualObject::OnMouseMove(pt);
}
//*******************************************************************************
void CBCGPGaugeImpl::OnCancelMode()
{
	CBCGPBaseVisualObject::OnCancelMode();

	double dblVal = 0.;

	FireTrackEvent(BCGM_ON_GAUGE_CANCEL_TRACK, dblVal, CBCGPPoint(-1., -1.));

	m_nTrackingScale = -1;
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::FireTrackEvent(UINT nMessage, double& dblValue, const CBCGPPoint& pt)
{
	if (m_pWndOwner->GetSafeHwnd() == NULL)
	{
		return FALSE;
	}

	CBCGPGaugeTrackingData data;

	data.m_nMessage = nMessage;
	data.m_nScale = m_nTrackingScale;
	data.m_Point = pt;
	data.m_Value = dblValue;
	data.m_pGauge = this;

	WPARAM wp = (WPARAM)GetID();
	LPARAM lp = nMessage != BCGM_ON_GAUGE_CANCEL_TRACK ? (LPARAM)&data : NULL;

	if (m_pWndOwner->SendMessage(nMessage, wp, lp) != 0)
	{
		if (nMessage != BCGM_ON_GAUGE_CANCEL_TRACK)
		{
			dblValue = data.m_Value;
		}

		return TRUE;
	}

	CWnd* pOwner = m_pWndOwner->GetOwner();
	if (pOwner->GetSafeHwnd() == NULL)
	{
		return FALSE;
	}

	if (pOwner->SendMessage(nMessage, wp, lp) != 0)
	{
		if (nMessage != BCGM_ON_GAUGE_CANCEL_TRACK)
		{
			dblValue = data.m_Value;
		}

		return TRUE;
	}

	return FALSE;
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::OnGetToolTip(const CBCGPPoint& /*pt*/, CString& strToolTip, CString& strDescr)
{
	if (m_strToolTip.IsEmpty())
	{
		return FALSE;
	}

	strToolTip = m_strToolTip;
	strDescr = m_strDescription;

	return TRUE;
}
//*******************************************************************************
void CBCGPGaugeImpl::SetToolTip(LPCTSTR lpszToolTip, LPCTSTR lpszDescr)
{
	m_strToolTip = lpszToolTip == NULL ? _T("") : lpszToolTip;
	m_strDescription = lpszDescr == NULL ? _T("") : lpszDescr;
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::OnKeyboardDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (m_bIsInteractiveMode && m_arScales.GetSize() > 0)
	{
		int nScale = 0;
		
		CBCGPGaugeScaleObject* pScale = m_arScales[nScale];
		ASSERT_VALID(pScale);

		double dblVal = GetValue();
		double dblNewVal = dblVal;

		switch (nChar)
		{
		case VK_HOME:
			dblNewVal = pScale->m_dblStart;
			break;

		case VK_END:
			dblNewVal = pScale->m_dblFinish;
			break;

		case VK_SUBTRACT:
		case VK_LEFT:
			dblNewVal = max(min(pScale->m_dblStart, pScale->m_dblFinish), dblNewVal - pScale->m_dblStep);
			break;

		case VK_ADD:
		case VK_RIGHT:
			dblNewVal = min(max(pScale->m_dblStart, pScale->m_dblFinish), dblNewVal + pScale->m_dblStep);
			break;
		}

		if (dblNewVal != dblVal)
		{
			FireTrackEvent(BCGM_ON_GAUGE_FINISH_TRACK, dblNewVal, CBCGPPoint(-1., -1.));

			SetValue(dblNewVal, 0, 0, TRUE);

			if (globalData.IsAccessibilitySupport())
			{
#ifdef _BCGSUITE_
				::NotifyWinEvent(EVENT_OBJECT_VALUECHANGE, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
#else
				globalData.NotifyWinEvent (EVENT_OBJECT_VALUECHANGE, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
#endif
			}
			
			return TRUE;
		}
	}

	return CBCGPBaseVisualObject::OnKeyboardDown(nChar, nRepCnt, nFlags);
}
//*******************************************************************************
CBCGPBaseVisualObject* CBCGPGaugeImpl::CreateCopy () const
{
	ASSERT_VALID(this);

	CRuntimeClass* pRTC = GetRuntimeClass();
	if (pRTC == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}

	CObject* pNewObj = pRTC->CreateObject();
	if (pNewObj == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}

	CBCGPGaugeImpl* pNewGauge = DYNAMIC_DOWNCAST(CBCGPGaugeImpl, pNewObj);
	if (pNewGauge  == NULL)
	{
		ASSERT(FALSE);
		delete pNewObj;
		return NULL;
	}

	pNewGauge->CopyFrom(*this);
	return pNewGauge;
}
//*******************************************************************************
void CBCGPGaugeImpl::CopyFrom(const CBCGPBaseVisualObject& srcObj)
{
	ASSERT_VALID(this);

	CBCGPBaseVisualObject::CopyFrom(srcObj);

	const CBCGPGaugeImpl& src = (const CBCGPGaugeImpl&)srcObj;

	RemoveAllColoredRanges();
	RemoveAllScales();

	m_nFrameSize =		   src.m_nFrameSize;        
	m_Pos =				   src.m_Pos;               
	m_ptOffset =		   src.m_ptOffset;          
	m_nTrackingScale =	   src.m_nTrackingScale;    
	m_brShadow =		   src.m_brShadow;          
	m_strToolTip =		   src.m_strToolTip;        
	m_strDescription =	   src.m_strDescription;

	for (int iScale = 0; iScale < src.m_arScales.GetSize(); iScale++)
	{
		CBCGPGaugeScaleObject* pScaleSrc = src.m_arScales[iScale];
		ASSERT_VALID(pScaleSrc);

		CBCGPGaugeScaleObject* pScale = pScaleSrc->CreateCopy();
		if (pScale == NULL)
		{
			ASSERT(FALSE);
			continue;
		}

		pScale->SetParentGauge(this);
		m_arScales.Add(pScale);
	}

	for (int iColorRange = 0; iColorRange < src.m_arRanges.GetSize(); iColorRange++)
	{
		CBCGPGaugeColoredRangeObject* pColorRangeSrc = src.m_arRanges[iColorRange];
		ASSERT_VALID(pColorRangeSrc);
		
		CBCGPGaugeColoredRangeObject* pColorRange = pColorRangeSrc->CreateCopy();
		if (pColorRange == NULL)
		{
			ASSERT(FALSE);
			continue;
		}
		
		AddColoredRange(pColorRange, FALSE);
	}

	for (int iLevelBar = 0; iLevelBar < src.m_arLevelBars.GetSize(); iLevelBar++)
	{
		CBCGPGaugeLevelBarObject* pLevelBarSrc = src.m_arLevelBars[iLevelBar];
		ASSERT_VALID(pLevelBarSrc);
		
		CBCGPGaugeLevelBarObject* pLevelBar = pLevelBarSrc->CreateCopy();
		if (pLevelBar == NULL)
		{
			ASSERT(FALSE);
			continue;
		}
		
		AddLevelBar(pLevelBar, FALSE);
	}
}
//*******************************************************************************
int CBCGPGaugeImpl::AddLevelBar(const CBCGPBrush& brFill, const CBCGPBrush& brFrame, const CBCGPBrush& brFillValue, int nScale /*= 0*/, double dblWidth /*= 0.*/, double dblOffsetFromFrame /*= 0.*/, 
								CBCGPGaugeLevelBarObject::BCGP_GAUGE_LEVEL_OBJECT_STYLE style/* = CBCGPGaugeLevelBarObject::BCGP_GAUGE_LEVEL_OBJECT_RECT*/, BOOL bRedraw /*= FALSE*/)
{
	return AddLevelBar(new CBCGPGaugeLevelBarObject(
		brFill, brFillValue, brFrame, nScale, 
		dblWidth, dblOffsetFromFrame, style), bRedraw);
}
//*******************************************************************************
int CBCGPGaugeImpl::AddLevelBar(CBCGPGaugeLevelBarObject* pLevelBar, BOOL bRedraw /*= FALSE*/)
{
	ASSERT_VALID(pLevelBar);
	
	pLevelBar->SetParentGauge (this);
	m_arLevelBars.Add(pLevelBar);
	
	SetDirty(TRUE, bRedraw);
	return (int)m_arLevelBars.GetSize() - 1;
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::ModifyLevelBar(int nIndex, const CBCGPGaugeLevelBarObject& levelBar, BOOL bRedraw /*= FALSE*/)
{
	if (nIndex < 0 || nIndex >= m_arLevelBars.GetSize())
	{
		return FALSE;
	}
	
	CBCGPGaugeLevelBarObject* pLevelBar = m_arLevelBars[nIndex];
	ASSERT_VALID(pLevelBar);
	
	*pLevelBar = levelBar;
	
	pLevelBar->m_Shape.Clear();
	pLevelBar->m_Shape.Destroy();
	
	SetDirty(TRUE, bRedraw);
	return TRUE;
}
//*******************************************************************************
BOOL CBCGPGaugeImpl::RemoveLevelBar(int nIndex, BOOL bRedraw /*= FALSE*/)
{
	if (nIndex < 0 || nIndex >= m_arLevelBars.GetSize())
	{
		return FALSE;
	}
	
	CBCGPGaugeLevelBarObject* pLevelBar = m_arLevelBars[nIndex];
	ASSERT_VALID(pLevelBar);
	
	delete pLevelBar;
	m_arLevelBars.RemoveAt(nIndex);
	
	SetDirty(TRUE, bRedraw);
	return TRUE;
}
//*******************************************************************************
void CBCGPGaugeImpl::RemoveAllLevelBars(BOOL bRedraw /*= FALSE*/)
{
	for (int i = 0; i < m_arLevelBars.GetSize(); i++)
	{
		CBCGPGaugeLevelBarObject* pLevelBar = m_arLevelBars[i];
		ASSERT_VALID(pLevelBar);
		
		delete pLevelBar;
	}
	
	m_arLevelBars.RemoveAll();
	
	SetDirty(TRUE, bRedraw);
}
//*******************************************************************************
void CBCGPGaugeImpl::FireClickEvent(const CBCGPPoint& pt)
{
	if (m_pWndOwner->GetSafeHwnd() == NULL)
	{
		return;
	}

	if (globalData.IsAccessibilitySupport() && m_bIsInteractiveMode)
	{
#ifdef _BCGSUITE_
		::NotifyWinEvent (EVENT_OBJECT_STATECHANGE, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
		::NotifyWinEvent (EVENT_OBJECT_FOCUS, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
#else
		globalData.NotifyWinEvent (EVENT_OBJECT_STATECHANGE, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
		globalData.NotifyWinEvent (EVENT_OBJECT_FOCUS, m_pWndOwner->GetSafeHwnd (), OBJID_CLIENT , CHILDID_SELF);
#endif
	}
	
	m_pWndOwner->PostMessage(BCGM_ON_GAUGE_CLICK, (WPARAM)GetID(), MAKELPARAM(pt.x, pt.y));

	CWnd* pOwner = m_pWndOwner->GetOwner();
	if (pOwner->GetSafeHwnd() == NULL)
	{
		return;
	}

	m_pWndOwner->GetOwner()->PostMessage(BCGM_ON_GAUGE_CLICK, (WPARAM)GetID(), MAKELPARAM(pt.x, pt.y));
}
