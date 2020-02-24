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
// BCGPStaticGaugeImpl.cpp: implementation of the CBCGPStaticGaugeImpl class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <math.h>
#include "BCGPStaticGaugeImpl.h"
#include "BCGPVisualCtrl.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

UINT BCGM_ON_GAUGE_SCROLL_FINISHED = ::RegisterWindowMessage (_T("BCGM_ON_GAUGE_SCROLL_FINISHED"));

IMPLEMENT_DYNAMIC(CBCGPStaticGaugeImpl, CBCGPGaugeImpl)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPStaticGaugeImpl::CBCGPStaticGaugeImpl(CBCGPVisualContainer* pContainer) :
	CBCGPGaugeImpl(pContainer)
{
	AddData(new CBCGPGaugeDataObject);
	m_nFlashTime = 0;
	m_bOff = FALSE;
	m_bIsPressed = FALSE;
	m_dblOpacity = 1.0;
	m_dblScrollOffset = 0.0;
	m_nAnimationType = BCGPSTATICGAUGE_ANIMATION_NONE;
	m_bIsHorizontalScroll = TRUE;
	m_nScrollFlags = 0;
	m_dblScrollTime = 0.0;
	m_dblScrollDelay = 0.0;
	m_ScrollAnimationType = CBCGPAnimationManager::BCGPANIMATION_AccelerateDecelerate;
	m_bScrollInternal = FALSE;
	m_bScrollStopped = FALSE;
	m_bScrollPaused = FALSE;
	m_DefaultDrawFlags = BCGP_DRAW_STATIC;
}
//*******************************************************************************
CBCGPStaticGaugeImpl::~CBCGPStaticGaugeImpl()
{
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::StartFlashing(UINT nTime)
{
	StopContentScrolling();

	if (IsFlashing() || nTime == 0)
	{
		return;
	}

	m_arData[0]->StartFlashAnimation(0.001 * nTime);
	m_nFlashTime = nTime;
	m_nAnimationType = BCGPSTATICGAUGE_ANIMATION_FLASH;
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::StopFlashing()
{
	if (IsFlashing())
	{
		m_arData[0]->StopAnimation();
		m_bOff = FALSE;
		m_nAnimationType = BCGPSTATICGAUGE_ANIMATION_NONE;
		Redraw();
	}
}
//*******************************************************************************
BOOL CBCGPStaticGaugeImpl::StartContentScrolling(double dblScrollTime, double dblScrollDelay,
		BOOL bIsHorizontal, UINT nFlags,
		CBCGPAnimationManager::BCGPAnimationType type, 
		CBCGPAnimationManagerOptions* pOptions)
{
	StopFlashing();

	if (IsContentScrolling())
	{
		return FALSE;
	}

	BOOL bWasPaused = m_bScrollPaused;

	m_bScrollPaused = FALSE;
	m_bIsHorizontalScroll = bIsHorizontal;
	m_nScrollFlags = nFlags;
	m_ScrollAnimationType = type;
	m_dblScrollTime = dblScrollTime;
	m_dblScrollDelay = dblScrollDelay;
	m_bScrollStopped = FALSE;

	if (pOptions != NULL)
	{
		m_ScrollAnimationOptions = *pOptions;
	}

	CBCGPBaseVisualCtrl* pWnd = DYNAMIC_DOWNCAST(CBCGPBaseVisualCtrl, GetParentWnd());
	if (pWnd->GetSafeHwnd() == NULL)
	{
		return FALSE;
	}

	const CBCGPSize sizeTotal = GetContentTotalSize(pWnd->GetGraphicsManager());
	const double dblTotal = bIsHorizontal ? sizeTotal.cx : sizeTotal.cy;

	if (dblTotal <= 0.0)
	{
		return FALSE;
	}

	const double dblGaugeSize = bIsHorizontal ? m_rect.Width() : m_rect.Height();
	double dblFinishVal = dblTotal - dblGaugeSize;

	if (dblFinishVal <= 0.0)
	{
		return FALSE;
	}

	if (m_dblScrollOffset >= dblFinishVal)
	{
		if ((m_nScrollFlags & BCGPSTATICGAUGE_CONTENT_SCROLL_CYCLIC) == BCGPSTATICGAUGE_CONTENT_SCROLL_CYCLIC)
		{
			// Scroll back
			dblFinishVal = 0.0;
		}
	}

	if (bWasPaused)
	{
		// Adjust scroll time:
		dblScrollTime = dblScrollTime * fabs(dblFinishVal - m_dblScrollOffset) / (dblTotal - dblGaugeSize);
	}

	if (!m_arData[0]->StartAnimation(m_dblScrollOffset, dblFinishVal, dblScrollTime, 
		m_ScrollAnimationType, &m_ScrollAnimationOptions, m_bScrollInternal ? m_dblScrollDelay : 0.0))
	{
		return FALSE;
	}

	m_nAnimationType = BCGPSTATICGAUGE_ANIMATION_SCROLL;
	return TRUE;
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::StopContentScrolling(BOOL bResetOffset)
{
	if (IsContentScrolling())
	{
		m_bScrollStopped = TRUE;

		m_arData[0]->StopAnimation();
		m_bScrollInternal = FALSE;
		m_nAnimationType = BCGPSTATICGAUGE_ANIMATION_NONE;

		if (bResetOffset)
		{
			m_dblScrollOffset = 0.0;
		}

		Redraw();
	}
	else if (bResetOffset && m_dblScrollOffset != 0.0)
	{
		m_dblScrollOffset = 0.0;
		Redraw();
	}
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::OnAnimation(CBCGPVisualDataObject* pDataObject)
{
	if (m_nAnimationType == BCGPSTATICGAUGE_ANIMATION_SCROLL)
	{
		ASSERT_VALID(pDataObject);

		m_dblScrollOffset = pDataObject->GetAnimatedValue();
	}
	else if (m_nAnimationType == BCGPSTATICGAUGE_ANIMATION_FLASH)
	{
		m_bOff = !m_bOff;
	}
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::OnAnimationFinished(CBCGPVisualDataObject* /*pDataObject*/)
{
	if (m_nAnimationType == BCGPSTATICGAUGE_ANIMATION_SCROLL)
	{
		if ((m_nScrollFlags & BCGPSTATICGAUGE_CONTENT_SCROLL_CYCLIC) == 0)
		{
			m_bScrollInternal = FALSE;
			m_nAnimationType = BCGPSTATICGAUGE_ANIMATION_NONE;

			if (!m_bScrollStopped)
			{
				if (m_pWndOwner->GetSafeHwnd() != NULL)
				{
					ASSERT_VALID(m_pWndOwner);

					m_pWndOwner->PostMessage(BCGM_ON_GAUGE_SCROLL_FINISHED, (WPARAM)GetID());
				
					CWnd* pOwner = m_pWndOwner->GetOwner();
					if (pOwner->GetSafeHwnd() != NULL)
					{
						m_pWndOwner->GetOwner()->PostMessage(BCGM_ON_GAUGE_SCROLL_FINISHED, (WPARAM)GetID());
					}
				}
			}
		}
		else
		{
			m_bScrollInternal = TRUE;

			StartContentScrolling(m_dblScrollTime, m_dblScrollDelay, m_bIsHorizontalScroll, m_nScrollFlags,
				m_ScrollAnimationType, &m_ScrollAnimationOptions);
		}
	}
}
//*******************************************************************************
CWnd* CBCGPStaticGaugeImpl::SetOwner(CWnd* pWndOwner, BOOL bRedraw)
{
	BOOL bIsFlashing = IsFlashing();

	CWnd* pWndRes = CBCGPGaugeImpl::SetOwner(pWndOwner, bRedraw);

	if (bIsFlashing)
	{
		StartFlashing(m_nFlashTime);
	}

	return pWndRes;
}
//*******************************************************************************
BOOL CBCGPStaticGaugeImpl::OnMouseDown(int nButton, const CBCGPPoint& pt)
{
	if (!m_bIsInteractiveMode || nButton != 0)
	{
		return CBCGPGaugeImpl::OnMouseDown(nButton, pt);
	}

	m_bIsPressed = TRUE;
	return TRUE;
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::OnMouseUp(int nButton, const CBCGPPoint& pt)
{
	if (!m_bIsInteractiveMode || nButton != 0)
	{
		CBCGPGaugeImpl::OnMouseUp(nButton, pt);
		return;
	}

	if (!m_bIsPressed)
	{
		return;
	}

	m_bIsPressed = FALSE;
	FireClickEvent(pt);
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::OnMouseMove(const CBCGPPoint& pt)
{
	if ((m_nScrollFlags & BCGPSTATICGAUGE_CONTENT_SCROLL_PAUSE_ON_MOUSE) == BCGPSTATICGAUGE_CONTENT_SCROLL_PAUSE_ON_MOUSE &&
		IsContentScrolling())
	{
		StopContentScrolling(FALSE);
		m_bScrollPaused = TRUE;
	}

	if (!m_bIsPressed)
	{
		CBCGPGaugeImpl::OnMouseMove(pt);
	}
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::OnMouseLeave()
{
	if (m_bScrollPaused)
	{
		StartContentScrolling(m_dblScrollTime, m_dblScrollDelay, m_bIsHorizontalScroll, m_nScrollFlags,
			m_ScrollAnimationType, &m_ScrollAnimationOptions);
	}

	CBCGPGaugeImpl::OnMouseLeave();
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::OnCancelMode()
{
	if (!m_bIsPressed)
	{
		CBCGPGaugeImpl::OnCancelMode();
		return;
	}

	m_bIsPressed = FALSE;
}
//*******************************************************************************
BOOL CBCGPStaticGaugeImpl::OnSetMouseCursor(const CBCGPPoint& pt)
{
	if (m_bIsInteractiveMode)
	{
		::SetCursor (globalData.GetHandCursor());
		return TRUE;
	}

	return CBCGPGaugeImpl::OnSetMouseCursor(pt);
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::SetFillBrush(const CBCGPBrush& brush)
{
	m_brFill = brush;

	SetDirty();
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::SetOutlineBrush(const CBCGPBrush& brush)
{
	m_brOutline = brush;

	SetDirty();
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::SetOpacity(double opacity, BOOL bRedraw)
{
	m_dblOpacity = opacity;

	if (bRedraw)
	{
		Redraw();
	}
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::SetDefaultDrawFlags(DWORD dwDrawFlags, BOOL bRedraw)
{
	m_DefaultDrawFlags = dwDrawFlags;

	if (bRedraw)
	{
		Redraw();
	}
}
//*******************************************************************************
void CBCGPStaticGaugeImpl::CopyFrom(const CBCGPBaseVisualObject& srcObj)
{
	CBCGPGaugeImpl::CopyFrom(srcObj);

	const CBCGPStaticGaugeImpl& src = (const CBCGPStaticGaugeImpl&)srcObj;

	m_nFlashTime = src.m_nFlashTime;
	m_dblOpacity = src.m_dblOpacity;
	m_brFill = src.m_brFill;
	m_brOutline = src.m_brOutline;
	m_DefaultDrawFlags = src.m_DefaultDrawFlags;
}

//----------------------------------------------------------------------------------
// CBCGPStaticFrameImpl

IMPLEMENT_DYNCREATE(CBCGPStaticFrameImpl, CBCGPStaticGaugeImpl)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPStaticFrameImpl::CBCGPStaticFrameImpl(CBCGPVisualContainer* pContainer) :
	CBCGPStaticGaugeImpl(pContainer)
{
	m_dblFrameSize = 1.0;
	m_dblCornerRadius = 0.0;

	m_brOutline = CBCGPBrush(CBCGPColor::Black);
}
//*******************************************************************************
CBCGPStaticFrameImpl::~CBCGPStaticFrameImpl()
{
}
//*******************************************************************************
void CBCGPStaticFrameImpl::OnDraw(CBCGPGraphicsManager* pGM, const CBCGPRect& /*rectClip*/, DWORD dwFlags)
{
	ASSERT_VALID(pGM);

	if (m_rect.IsRectEmpty() || !m_bIsVisible)
	{
		return;
	}

	if ((dwFlags & m_DefaultDrawFlags) == 0)
	{
		return;
	}

	CBCGPRect rectDraw = m_rect;
	if (m_bIsSubGauge)
	{
		rectDraw.OffsetRect(-m_ptScrollOffset);
	}
	
	CBCGPRect rect = rectDraw;
	const double scaleRatio = GetScaleRatioMid();

	CBCGPStrokeStyle* pStrokeStyle = m_strokeStyle.IsEmpty() ? NULL : &m_strokeStyle;

	if (m_dblCornerRadius > 0.0)
	{
		CBCGPRoundedRect rectRounded(rect, m_dblCornerRadius * scaleRatio, m_dblCornerRadius * scaleRatio);

		pGM->FillRoundedRectangle(rectRounded, m_brFill);
		pGM->DrawRoundedRectangle(rectRounded, m_brOutline, m_dblFrameSize * scaleRatio, pStrokeStyle);
		
		if (pGM->GetType() == CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER_GDI)
		{
			rectRounded.rect.DeflateRect(m_sizeScaleRatio.cx, m_sizeScaleRatio.cy);
			pGM->DrawRoundedRectangle(rectRounded, m_brOutline, m_dblFrameSize * scaleRatio, pStrokeStyle);
		}
	}
	else
	{
		pGM->FillRectangle(rect, m_brFill);
		pGM->DrawRectangle(rect, m_brOutline, m_dblFrameSize * scaleRatio, pStrokeStyle);

		if (pGM->GetType() == CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER_GDI)
		{
			rect.DeflateRect(m_sizeScaleRatio.cx, m_sizeScaleRatio.cy);
			pGM->DrawRectangle(rect, m_brOutline, m_dblFrameSize * scaleRatio, pStrokeStyle);
		}
	}

	SetDirty (FALSE);
}
//*******************************************************************************
void CBCGPStaticFrameImpl::CopyFrom(const CBCGPBaseVisualObject& srcObj)
{
	CBCGPStaticGaugeImpl::CopyFrom(srcObj);

	const CBCGPStaticFrameImpl& src = (const CBCGPStaticFrameImpl&)srcObj;

	m_strokeStyle = src.m_strokeStyle;
	m_dblFrameSize = src.m_dblFrameSize;
	m_dblCornerRadius = src.m_dblCornerRadius;
}
//*******************************************************************************
void CBCGPStaticFrameImpl::SetFrameSize(double dblFrameSize, BOOL bRedraw)
{
	m_dblFrameSize = dblFrameSize;

	if (bRedraw)
	{
		Redraw();
	}
}
//*******************************************************************************
void CBCGPStaticFrameImpl::SetStrokeStyle(const CBCGPStrokeStyle& strokeStyle, BOOL bRedraw)
{
	m_strokeStyle = strokeStyle;

	if (bRedraw)
	{
		Redraw();
	}
}
//*******************************************************************************
void CBCGPStaticFrameImpl::SetCornerRadius(double dblCornerRadius, BOOL bRedraw)
{
	m_dblCornerRadius = dblCornerRadius;
	
	if (bRedraw)
	{
		Redraw();
	}
}

