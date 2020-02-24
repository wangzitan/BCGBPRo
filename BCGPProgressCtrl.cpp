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
// BCGPProgressCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPVisualManager.h"
#ifndef _BCGSUITE_
#include "BCGPRibbonProgressBar.h"
#include "BCGPToolBarImages.h"
#endif
#include "bcgglobals.h"
#include "BCGPProgressCtrl.h"
#include "BCGPDlgImpl.h"
#include "BCGPGlobalUtils.h"
#include "BCGPDrawManager.h"

#ifndef _BCGSUITE_
	#define visualManagerMFC	CBCGPVisualManager::GetInstance ()
#else
	#define visualManagerMFC	CMFCVisualManager::GetInstance ()
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static const int idTimerMarquee = 1001;
static const int nMarqueeDotSize = 5;
static const double dblMarqueeDotEllapse = 1.2;

/////////////////////////////////////////////////////////////////////////////
// CBCGPProgressMarqueeObject

CBCGPProgressMarqueeObject::CBCGPProgressMarqueeObject()
{
	m_pParent = NULL;
	m_nIndex = -1;
	m_nFinish = m_nOffset = -1;
	m_bIsFirstStep = FALSE;
}
//**************************************************************************
void CBCGPProgressMarqueeObject::Start(CBCGPProgressCtrl* pParent, int nIndex, int nOffset, int nMiddle, int nFinish)
{
	m_pParent = pParent;
	m_nIndex = nIndex;
	
	m_nFinish = nFinish;
	
	CBCGPAnimationManagerOptions options;
	options.m_dblAccelerationRatio = 0.0;
	options.m_dblDecelerationRatio = 1.0;
	
	m_bIsFirstStep = TRUE;
	
	StartAnimation(nOffset, nMiddle, dblMarqueeDotEllapse, CBCGPAnimationManager::BCGPANIMATION_AccelerateDecelerate, &options, 
		0.4 * m_nIndex);
}
//**************************************************************************
void CBCGPProgressMarqueeObject::OnAnimationValueChanged(double /*dblOldValue*/, double dblNewValue)
{
	ASSERT_VALID(m_pParent);

	m_nOffset = (int)dblNewValue;
	m_pParent->RedrawWindow();
}
//**************************************************************************
void CBCGPProgressMarqueeObject::OnAnimationFinished()
{
	if (!m_bIsFirstStep)
	{
		if (m_nIndex == BCGP_PROGRESS_MARQUEE_OBJECTS - 1)
		{
			m_pParent->PostMessage(PBM_SETMARQUEE, TRUE, 0);
		}
		
		return;
	}
	
	m_bIsFirstStep = FALSE;
	
	CBCGPAnimationManagerOptions options;
	options.m_dblAccelerationRatio = 1.0;
	options.m_dblDecelerationRatio = 0.0;
	
	StartAnimation(m_nOffset, m_nFinish, dblMarqueeDotEllapse, CBCGPAnimationManager::BCGPANIMATION_AccelerateDecelerate, &options,
		0.25 * BCGP_PROGRESS_MARQUEE_OBJECTS);
}
//**************************************************************************
void CBCGPProgressMarqueeObject::OnAnimationIdle()
{
	ASSERT_VALID(m_pParent);

	m_pParent->OnAnimationIdle();
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPProgressCtrl

IMPLEMENT_DYNAMIC(CBCGPProgressCtrl, CProgressCtrl)

CBCGPProgressCtrl::CBCGPProgressCtrl()
{
	m_bVisualManagerStyle = FALSE;
	m_bOnGlass = FALSE;
	m_bIsMarqueeStarted = FALSE;
	m_nMarqueeStep = 0;
	m_MarqueeStyle = BCGP_MARQUEE_DEFAULT;
	m_clrMarquee = (COLORREF)-1;
	m_clrMarqueeGradient = (COLORREF)-1;
}

CBCGPProgressCtrl::~CBCGPProgressCtrl()
{
}

BEGIN_MESSAGE_MAP(CBCGPProgressCtrl, CProgressCtrl)
	//{{AFX_MSG_MAP(CBCGPProgressCtrl)
	ON_WM_ERASEBKGND()
	ON_WM_PAINT()
	ON_WM_NCCALCSIZE()
	ON_WM_NCPAINT()
	ON_WM_TIMER()
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLVMMODE, OnBCGSetControlVMMode)
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLAERO, OnBCGSetControlAero)
	ON_MESSAGE(PBM_SETMARQUEE, OnSetMarquee)
	ON_MESSAGE(WM_PRINTCLIENT, OnPrintClient)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPProgressCtrl message handlers

BOOL CBCGPProgressCtrl::OnEraseBkgnd(CDC* /*pDC*/) 
{
	BOOL bCustomInfiniteMode = (GetStyle() & PBS_MARQUEE) != 0 && m_MarqueeStyle != BCGP_MARQUEE_DEFAULT;;
	
	if (!m_bVisualManagerStyle && !bCustomInfiniteMode)
	{
		return (BOOL)Default();
	}

	return TRUE;
}
//**************************************************************************
void CBCGPProgressCtrl::OnPaint() 
{
	BOOL bCustomInfiniteMode = (GetStyle() & PBS_MARQUEE) != 0 && m_MarqueeStyle != BCGP_MARQUEE_DEFAULT;

	if (!m_bVisualManagerStyle && !bCustomInfiniteMode)
	{
		Default();
		return;
	}

	CPaintDC dc(this); // device context for painting
	
	CBCGPMemDC memDC (dc, this);
	CDC* pDC = &memDC.GetDC ();

	DoPaint(pDC);
}
//**************************************************************************
void CBCGPProgressCtrl::DoPaint(CDC* pDC) 
{
	ASSERT_VALID(pDC);

	if (!CBCGPVisualManager::GetInstance ()->OnFillParentBarBackground (this, pDC))
	{
		globalData.DrawParentBackground (this, pDC);
	}

	CRect rectProgress;
	GetClientRect(rectProgress);

	BOOL bInfiniteMode = (GetStyle() & PBS_MARQUEE) != 0;

	int nMin = 0;
	int nMax = 100;
	int nPos = m_nMarqueeStep;
	
	if (!bInfiniteMode)
	{
		GetRange (nMin, nMax);
		nPos = GetPos();
	}

	BOOL bIsVertical = (GetStyle() & PBS_VERTICAL) != 0;

	CBCGPDrawOnGlass dog (m_bOnGlass);

	if (bInfiniteMode && m_MarqueeStyle == BCGP_MARQUEE_DOTS)
	{
		const int nMarqueeDotSizeCurr = globalUtils.ScaleByDPI(nMarqueeDotSize);

		int x = rectProgress.CenterPoint().x - nMarqueeDotSizeCurr / 2;
		int y = rectProgress.CenterPoint().y - nMarqueeDotSizeCurr / 2;
		
		for (int i = 0; i < BCGP_PROGRESS_MARQUEE_OBJECTS; i++)
		{
			int nOffset = m_MarqueeObjects[i].m_nOffset;

			if (nOffset > 0)
			{
				CRect rect(
					bIsVertical ? CPoint(x, rectProgress.bottom - nOffset - nMarqueeDotSizeCurr) : CPoint(nOffset, y), 
					CSize(nMarqueeDotSizeCurr, nMarqueeDotSizeCurr));
				rect.DeflateRect(1, 1);

				CBCGPVisualManager::GetInstance()->OnDrawProgressMarqueeDot(pDC, this, rect, i);
			}
		}
		return;
	}

	if (bInfiniteMode && m_MarqueeStyle == BCGP_MARQUEE_GRADIENT && !globalData.IsHighContastMode())
	{
		if (m_bIsMarqueeStarted)
		{
			CBCGPDrawManager dm(*pDC);

			COLORREF clr1 = m_clrMarquee;
			COLORREF clr2 = m_clrMarqueeGradient;

			if (clr1 == (COLORREF)-1)
			{
				clr1 = m_bVisualManagerStyle ? CBCGPVisualManager::GetInstance()->GetProgressMarqueeColor(this) : RGB(0, 128, 0);
			}

			if (clr2 == (COLORREF)-1)
			{
				if (CBCGPDrawManager::IsPaleColor(clr1))
				{
					clr2 = CBCGPDrawManager::ColorMakeDarker(clr1, .5);
				}
				else
				{
					clr2 = CBCGPDrawManager::ColorMakePale(clr1);
				}
			}

			dm.DrawRect(rectProgress, clr1, (COLORREF)-1);

			if (bIsVertical)
			{
				int nHeight = rectProgress.Height();
				
				rectProgress.bottom = (rectProgress.bottom + nHeight) - (rectProgress.Height() + nHeight) * m_nMarqueeStep / 100;
				rectProgress.top = rectProgress.bottom - nHeight;

				dm.Fill4ColorsGradient(rectProgress, clr2, clr1, clr1, clr2, bIsVertical, 50);
			}
			else
			{
				int nWidth = rectProgress.Width();

				rectProgress.left = (rectProgress.left - nWidth) + (rectProgress.Width() + nWidth) * m_nMarqueeStep / 100;
				rectProgress.right = rectProgress.left + nWidth;

				dm.Fill4ColorsGradient(rectProgress, clr1, clr2, clr2, clr1, bIsVertical, 50);
			}
		}
		return;
	}

#if (!defined BCGP_EXCLUDE_RIBBON) && (!defined _BCGSUITE_)

	CBCGPRibbonProgressBar dummy(0, 0, 0, bIsVertical);
	dummy.SetRange(nMin, nMax);
	dummy.SetPos(nPos, FALSE);
	dummy.SetInfiniteMode(bInfiniteMode);
	
	CRect rectChunk(0, 0, 0, 0);
	dummy.CalculateChunkRect(rectProgress, rectChunk);

	visualManagerMFC->OnDrawRibbonProgressBar(pDC, &dummy, rectProgress, rectChunk, bInfiniteMode);
#else
	visualManagerMFC->OnDrawStatusBarProgress (pDC, NULL,
										rectProgress, nMax - nMin, 
										GetPos (),
										globalData.clrHilite, 
										(COLORREF)-1, 
										(COLORREF)-1,
										FALSE);
#endif
}
//**************************************************************************
void CBCGPProgressCtrl::OnNcCalcSize(BOOL /*bCalcValidRects*/, NCCALCSIZE_PARAMS FAR* /*lpncsp*/) 
{
}
//**************************************************************************
void CBCGPProgressCtrl::OnNcPaint() 
{
}
//**************************************************************************
LRESULT CBCGPProgressCtrl::OnBCGSetControlVMMode (WPARAM wp, LPARAM)
{
	m_bVisualManagerStyle = (BOOL) wp;
	return 0;
}
//**************************************************************************
LRESULT CBCGPProgressCtrl::OnBCGSetControlAero (WPARAM wp, LPARAM)
{
	m_bOnGlass = (BOOL) wp;
	return 0;
}
//**************************************************************************
LRESULT CBCGPProgressCtrl::OnSetMarquee(WPARAM wp, LPARAM lp)
{
	m_bIsMarqueeStarted = (BOOL)wp;

	switch (m_MarqueeStyle)
	{
	case BCGP_MARQUEE_DEFAULT:
	case BCGP_MARQUEE_GRADIENT:
		if (m_bIsMarqueeStarted)
		{
			int nElapse = (lp == 0) ? 30 : (int)lp;
			SetTimer(idTimerMarquee, nElapse, NULL);
		}
		else
		{
			KillTimer(idTimerMarquee);
		}
		break;

	case BCGP_MARQUEE_DOTS:
		{
			CRect rectClient;
			GetClientRect(rectClient);
			
			const int nMarqueeDotSizeCurr = globalUtils.ScaleByDPI(nMarqueeDotSize);

			BOOL bIsVertical = (GetStyle() & PBS_VERTICAL) != 0;

			for (int i = 0; i < BCGP_PROGRESS_MARQUEE_OBJECTS; i++)
			{
				if (m_bIsMarqueeStarted)
				{
					int nMiddle = (bIsVertical ? rectClient.CenterPoint().y : rectClient.CenterPoint().x) -
						(i - BCGP_PROGRESS_MARQUEE_OBJECTS / 2) * (nMarqueeDotSizeCurr * 2);
					int nStart = bIsVertical ? rectClient.top : rectClient.left;
					int nFinish = bIsVertical ? rectClient.bottom : rectClient.right;

					m_MarqueeObjects[i].Start(this, i, nStart - nMarqueeDotSizeCurr / 2, nMiddle, nFinish);
				}
				else
				{
					m_MarqueeObjects[i].StopAnimation(FALSE);
					m_MarqueeObjects[i].m_nOffset = -1;
				}
			}
		}
		break;
	}

	if (!m_bIsMarqueeStarted)
	{
		RedrawWindow();
	}

	return Default();
}
//**************************************************************************
void CBCGPProgressCtrl::OnTimer(UINT_PTR nIDEvent) 
{
	CProgressCtrl::OnTimer(nIDEvent);

	if (nIDEvent == idTimerMarquee)
	{
		m_nMarqueeStep++;
		if (m_nMarqueeStep > 100)
		{
			m_nMarqueeStep = 0;
		}

		RedrawWindow();
	}
}
//**************************************************************************
LRESULT CBCGPProgressCtrl::OnPrintClient(WPARAM wp, LPARAM lp)
{
	BOOL bCustomInfiniteMode = (GetStyle() & PBS_MARQUEE) != 0 && m_MarqueeStyle != BCGP_MARQUEE_DEFAULT;;
	
	if (!m_bVisualManagerStyle && !bCustomInfiniteMode)
	{
		return Default();
	}

	if (lp & PRF_CLIENT)
	{
		CDC* pDC = CDC::FromHandle((HDC) wp);
		ASSERT_VALID(pDC);
		
		CRect rectClient;
		GetClientRect(rectClient);
		
		CBCGPMemDC memDC(*pDC, rectClient);
		
		DoPaint(&memDC.GetDC());
	}
	
	return 0;
}
//**************************************************************************
void CBCGPProgressCtrl::OnDestroy() 
{
	for (int i = 0; i < BCGP_PROGRESS_MARQUEE_OBJECTS; i++)
	{
		m_MarqueeObjects[i].StopAnimation(FALSE);
		m_MarqueeObjects[i].CleanUp();
	}

	CProgressCtrl::OnDestroy();
}
