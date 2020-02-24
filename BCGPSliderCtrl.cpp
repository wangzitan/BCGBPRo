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
// BCGPSliderCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPVisualManager.h"
#include "bcgglobals.h"
#include "BCGPSliderCtrl.h"
#include "BCGPDlgImpl.h"
#include "BCGPDrawManager.h"
#include "BCGPGlobalUtils.h"
#ifndef _BCGSUITE_
#include "BCGPBaseControlBar.h"
#include "BCGPToolBarImages.h"
#include "BCGPRibbonSlider.h"
#endif
#include "trackmouse.h"
#include "BCGPMath.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPSliderCtrl

static const UINT IdAutoCommand = 1;

#define UM_REDRAWFRAME	(WM_USER + 1001)

IMPLEMENT_DYNAMIC(CBCGPSliderCtrl, CSliderCtrl)

CBCGPSliderCtrl::CBCGPSliderCtrl()
{
	m_bVisualManagerStyle = FALSE;
	m_bOnGlass = FALSE;
	m_bTracked = FALSE;
	m_nTicFreq = 1;
	m_bIsThumbHighligted = FALSE;
	m_bIsThumPressed = FALSE;
	m_clrSelection = (COLORREF)-1;
	m_bZoomButtons = FALSE;
	m_rectZoomInBtn.SetRectEmpty();
	m_rectZoomOutBtn.SetRectEmpty();
	m_nZoomStepValue = 0;
	m_nZoomButtonSize = globalUtils.ScaleByDPI(20);
	m_nPressedZoomButton = -1;
	m_nHighlightedZoomButton = -1;
	m_nCapturedZoomButton = -1;
	m_nAutoRepeatDelay = 200;
	m_bAutoCommandSent = FALSE;
	m_bIsDelayedLayout = FALSE;
}

CBCGPSliderCtrl::~CBCGPSliderCtrl()
{
}

BEGIN_MESSAGE_MAP(CBCGPSliderCtrl, CSliderCtrl)
	//{{AFX_MSG_MAP(CBCGPSliderCtrl)
	ON_WM_ERASEBKGND()
	ON_WM_CANCELMODE()
	ON_WM_MOUSEMOVE()
	ON_WM_PAINT()
	ON_WM_LBUTTONUP()
	ON_WM_LBUTTONDOWN()
	ON_WM_NCCALCSIZE()
	ON_WM_NCMOUSEMOVE()
	ON_WM_NCPAINT()
	ON_WM_NCHITTEST()
	ON_WM_TIMER()
	ON_WM_CREATE()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_MOUSELEAVE, OnMouseLeave)
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLVMMODE, OnBCGSetControlVMMode)
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLAERO, OnBCGSetControlAero)
	ON_MESSAGE(WM_PRINTCLIENT, OnPrintClient)
	ON_MESSAGE(WM_PRINT, OnPrint)
	ON_MESSAGE(TBM_SETTICFREQ, OnSetTicFreq)
	ON_MESSAGE(UM_REDRAWFRAME, OnRedrawFrame)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPSliderCtrl message handlers

BOOL CBCGPSliderCtrl::OnEraseBkgnd(CDC* /*pDC*/)
{
	if (!m_bVisualManagerStyle)
	{
		return (BOOL) Default ();
	}

	return TRUE;
}
//**************************************************************************
LRESULT CBCGPSliderCtrl::OnBCGSetControlVMMode (WPARAM wp, LPARAM)
{
	m_bVisualManagerStyle = (BOOL) wp;
	RedrawWindow ();
	return 0;
}
//**************************************************************************
LRESULT CBCGPSliderCtrl::OnBCGSetControlAero (WPARAM wp, LPARAM)
{
	m_bOnGlass = (BOOL) wp;
	return 0;
}
//**************************************************************************
void CBCGPSliderCtrl::OnCancelMode() 
{
	CSliderCtrl::OnCancelMode();

	m_bIsThumbHighligted = FALSE;
	m_bIsThumPressed = FALSE;

	if (HasZoomButtons() && GetCapture()->GetSafeHwnd() == GetSafeHwnd())
	{
		ReleaseCapture ();

		if (m_nAutoRepeatDelay > 0)
		{
			KillTimer (IdAutoCommand);
			m_bAutoCommandSent = FALSE;
		}
	}
	
	m_nPressedZoomButton = -1;
	m_nCapturedZoomButton = -1;
	m_nHighlightedZoomButton = -1;
	
	RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
}
//*****************************************************************************************
void CBCGPSliderCtrl::OnMouseMove(UINT nFlags, CPoint point) 
{
	if (m_nCapturedZoomButton >= 0)
	{
		int nButtonPressed = HitTestZoomButton(point);
		if (nButtonPressed != m_nPressedZoomButton)
		{
			m_nPressedZoomButton = nButtonPressed;
			RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
		}
		
		return;
	}
	
	if (m_nHighlightedZoomButton >= 0)
	{
		int nButtonHighlighted = HitTestZoomButton(point);
		if (nButtonHighlighted != m_nHighlightedZoomButton)
		{
			m_nHighlightedZoomButton = -1;
			
			ReleaseCapture ();
			RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
		}
	}

	BOOL bIsThumbHighligted = m_bIsThumbHighligted;

	if (GetStyle() & TBS_NOTHUMB)
	{
		m_bIsThumbHighligted = FALSE;
		CSliderCtrl::OnMouseMove(nFlags, point);
		return;
	}

	CRect rectThumb;
	GetThumbRect (rectThumb);

	m_bIsThumbHighligted = rectThumb.PtInRect (point);

	CSliderCtrl::OnMouseMove(nFlags, point);

	if (bIsThumbHighligted != m_bIsThumbHighligted)
	{
		RedrawWindow ();
	}

	if (!m_bTracked)
	{
		m_bTracked = TRUE;
		
		TRACKMOUSEEVENT trackmouseevent;
		trackmouseevent.cbSize = sizeof(trackmouseevent);
		trackmouseevent.dwFlags = TME_LEAVE;
		trackmouseevent.hwndTrack = GetSafeHwnd();
		trackmouseevent.dwHoverTime = HOVER_DEFAULT;
		::BCGPTrackMouse (&trackmouseevent);	
	}
}
//*****************************************************************************************
LRESULT CBCGPSliderCtrl::OnMouseLeave(WPARAM,LPARAM)
{
	m_bTracked = FALSE;

	if (m_bIsThumbHighligted)
	{
		m_bIsThumbHighligted = FALSE;
		RedrawWindow ();
	}

	return 0;
}
//*****************************************************************************************
void CBCGPSliderCtrl::OnPaint() 
{
	if (!m_bVisualManagerStyle)
	{
		Default ();
		return;
	}

	CPaintDC dc(this); // device context for painting
	OnDraw(&dc);
}
//*****************************************************************************************
void CBCGPSliderCtrl::OnDraw(CDC* pDCIn) 
{
	CBCGPMemDC memDC (*pDCIn, this);
	CDC* pDC = &memDC.GetDC ();

	if (!CBCGPVisualManager::GetInstance ()->OnFillParentBarBackground (this, pDC))
	{
		globalData.DrawParentBackground (this, pDC, NULL);
	}

	DWORD dwStyle = GetStyle ();
	BOOL bVert = (dwStyle & TBS_VERT);
	BOOL bLeftTop = (dwStyle & TBS_BOTH) || (dwStyle & TBS_LEFT);
	BOOL bRightBottom = (dwStyle & TBS_BOTH) || ((dwStyle & TBS_LEFT) == 0);
	BOOL bIsFocused = GetSafeHwnd () == ::GetFocus ();
	BOOL bIsAutoTicks = (dwStyle & TBS_AUTOTICKS);

	CRect rectChannel;
	GetChannelRect (rectChannel);

	if (bVert)
	{
		CRect rect = rectChannel;

		rectChannel.left = rect.top;
		rectChannel.right = rect.bottom;
		rectChannel.top = rect.left;
		rectChannel.bottom = rect.right;
	}

	CRect rectSel = GetSelectionRect();

	CBCGPDrawOnGlass dog(m_bOnGlass);
	OnDrawChannel(pDC, bVert, rectChannel, m_bOnGlass);

	CRect rectThumb;
	GetThumbRect (rectThumb);

	int nTicSize = max(3, (bVert ? rectThumb.Height() : rectThumb.Width()) / 3);
	int nTicOffset = 1;

	int nNumTics = GetNumTics();
	int nRange = GetRangeMax() - GetRangeMin();
	int nTicFreq = (nNumTics == 2) ? nRange : m_nTicFreq;

	for (int i = 0; i < nNumTics; i++)
	{
		int nTicPos = GetTicPos(bIsAutoTicks ? i * nTicFreq : i);

		if (nTicPos < 0)
		{
			if (i == nNumTics - 2)
			{
				if (bVert)
				{
					nTicPos = rectChannel.top + rectThumb.Height() / 2;
				}
				else
				{
					nTicPos = rectChannel.left + rectThumb.Width() / 2;
				}
			}
			else if (i == nNumTics - 1)
			{
				if (bVert)
				{
					nTicPos = rectChannel.bottom - rectThumb.Height() / 2 - 1;
				}
				else
				{
					nTicPos = rectChannel.right - rectThumb.Width() / 2 - 1;
				}
			}
		}

		if (nTicPos >= 0)
		{
			CRect rectTic1(0, 0, 0, 0);
			CRect rectTic2(0, 0, 0, 0);

			if (bVert)
			{
				if (bLeftTop)
				{
					rectTic1 = CRect(rectThumb.left - nTicOffset - nTicSize, nTicPos, rectThumb.left - nTicOffset, nTicPos + 1);
				}
				
				if (bRightBottom)
				{
					rectTic2 = CRect(rectThumb.right + nTicOffset, nTicPos, rectThumb.right + nTicOffset + nTicSize, nTicPos + 1);
				}
			}
			else
			{
				if (bLeftTop)
				{
					rectTic1 = CRect(nTicPos, rectThumb.top - nTicOffset - nTicSize, nTicPos + 1, rectThumb.top - nTicOffset);
				}

				if (bRightBottom)
				{
					rectTic2 = CRect(nTicPos, rectThumb.bottom + nTicOffset, nTicPos + 1, rectThumb.bottom + nTicOffset + nTicSize);
				}
			}

			if (!rectTic1.IsRectEmpty())
			{
				CBCGPVisualManager::GetInstance ()->OnDrawSliderTic(pDC, this, rectTic1, bVert, TRUE, m_bOnGlass);
			}

			if (!rectTic2.IsRectEmpty())
			{
				CBCGPVisualManager::GetInstance ()->OnDrawSliderTic(pDC, this, rectTic2, bVert, FALSE, m_bOnGlass);
			}
		}
	}

	if (!rectSel.IsRectEmpty() && nNumTics >= 2)
	{
		for (int i = 0; i < 2; i++)
		{
			CRect rectSelMarker1(0, 0, 0, 0);
			CRect rectSelMarker2(0, 0, 0, 0);

			int nMarkerPos = 0;
			
			if (bVert)
			{
				nMarkerPos = (i == 0) ? rectSel.top : rectSel.bottom;

				if (bLeftTop)
				{
					rectSelMarker1 = CRect(rectThumb.left - nTicOffset - nTicSize, nMarkerPos, rectThumb.left - nTicOffset, nMarkerPos + 1);
				}
				
				if (bRightBottom)
				{
					rectSelMarker2 = CRect(rectThumb.right + nTicOffset, nMarkerPos, rectThumb.right + nTicOffset + nTicSize, nMarkerPos + 1);
				}
			}
			else
			{
				nMarkerPos = (i == 0) ? rectSel.left : rectSel.right;

				if (bLeftTop)
				{
					rectSelMarker1 = CRect(nMarkerPos - nTicSize / 2 - 2, rectThumb.top - nTicOffset - nTicSize, nMarkerPos + nTicSize / 2 - 1, rectThumb.top - nTicOffset);
				}
				
				if (bRightBottom)
				{
					rectSelMarker2 = CRect(nMarkerPos - nTicSize / 2 - 2, rectThumb.bottom + nTicOffset, nMarkerPos + nTicSize / 2 - 1, rectThumb.bottom + nTicOffset + nTicSize);
				}
			}

			if (i == 1)
			{
				if (bVert)
				{
					rectSelMarker1.OffsetRect(0, nTicSize);
					rectSelMarker2.OffsetRect(0, nTicSize);
				}
				else
				{
					rectSelMarker1.OffsetRect(nTicSize, 0);
					rectSelMarker2.OffsetRect(nTicSize, 0);
				}
			}
			
			if (!rectSelMarker1.IsRectEmpty())
			{
				CBCGPVisualManager::GetInstance ()->OnDrawSliderSelectionMarker(pDC, this, rectSelMarker1, i == 0, bVert, TRUE, m_bOnGlass);
			}
			
			if (!rectSelMarker2.IsRectEmpty())
			{
				CBCGPVisualManager::GetInstance ()->OnDrawSliderSelectionMarker(pDC, this, rectSelMarker2, i == 0, bVert, FALSE, m_bOnGlass);
			}
		}
	}

	if ((GetStyle() & TBS_NOTHUMB) == 0)
	{
		OnDrawThumb(pDC, rectThumb, m_bIsThumbHighligted || bIsFocused,
			m_bIsThumPressed, !IsWindowEnabled (),
			bVert, bLeftTop, bRightBottom, m_bOnGlass);
	}

	if (bIsFocused && m_bDrawFocus)
	{
		CRect rectFocus;
		GetClientRect (rectFocus);

		if (m_bOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawFocusRect(rectFocus);
		}
		else
		{
			pDC->DrawFocusRect (rectFocus);
		}
	}
}
//*****************************************************************************************
void CBCGPSliderCtrl::OnLButtonUp(UINT nFlags, CPoint point) 
{
	if (m_nCapturedZoomButton >= 0)
	{
		BOOL bIsZoomOut = (m_nPressedZoomButton == 0);

		ReleaseCapture ();
		globalUtils.NotifyReleaseCapture(this);
		
		if (m_nAutoRepeatDelay > 0)
		{
			KillTimer (IdAutoCommand);
			m_bAutoCommandSent = FALSE;
		}

		m_nCapturedZoomButton = m_nHighlightedZoomButton = m_nPressedZoomButton = -1;

		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);

		CRect rectBtn = bIsZoomOut ? m_rectZoomOutBtn : m_rectZoomInBtn;
		if (rectBtn.PtInRect(point) && !m_bAutoCommandSent)
		{
			OnClickZoomButton(bIsZoomOut);
		}

		return;
	}

	if (m_bIsThumPressed)
	{
		m_bIsThumPressed = FALSE;
		RedrawWindow ();
	}
	
	CSliderCtrl::OnLButtonUp(nFlags, point);
}
//*****************************************************************************************
void CBCGPSliderCtrl::OnLButtonDown(UINT nFlags, CPoint point) 
{
	if (HasZoomButtons())
	{
		int nZoomButton = HitTestZoomButton(point);
		if (nZoomButton >= 0)
		{
			SetFocus ();
			
			m_nPressedZoomButton = m_nCapturedZoomButton = nZoomButton;
			SetCapture ();
			
			if (m_nAutoRepeatDelay > 0)
			{
				m_bAutoCommandSent = FALSE;
				SetTimer(IdAutoCommand, m_nAutoRepeatDelay, NULL);
			}

			RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
			return;
		}
	}

	if ((GetStyle() & TBS_NOTHUMB) == 0)
	{
		CRect rectThumb;
		GetThumbRect (rectThumb);

		m_bIsThumPressed = rectThumb.PtInRect (point);
		
		if (m_bIsThumPressed)
		{
			RedrawWindow ();
		}
	}
	
	CSliderCtrl::OnLButtonDown(nFlags, point);
}
//*******************************************************************************
LRESULT CBCGPSliderCtrl::OnPrintClient(WPARAM wp, LPARAM lp)
{
	DWORD dwFlags = (DWORD)lp;

	if ((dwFlags & PRF_CLIENT) == PRF_CLIENT)
	{
		if (!m_bVisualManagerStyle)
		{
			return Default();
		}

		CDC* pDC = CDC::FromHandle((HDC) wp);
		ASSERT_VALID(pDC);

		OnDraw(pDC);
	}

	return 0;
}
//****************************************************************************
LRESULT CBCGPSliderCtrl::OnPrint(WPARAM wp, LPARAM lp)
{
	DWORD dwFlags = (DWORD)lp;

	CDC* pDC = CDC::FromHandle((HDC) wp);
	ASSERT_VALID(pDC);

	if ((dwFlags & PRF_NONCLIENT) == PRF_NONCLIENT)
	{
		if (m_bZoomButtons)
		{
			DoNcPaint(pDC);
		}
	}

	Default();

	if ((dwFlags & PRF_NONCLIENT) == PRF_NONCLIENT)
	{
		BOOL bHasBorders = (GetStyle () & WS_BORDER) != 0 || (GetExStyle () & WS_EX_CLIENTEDGE) != 0;
		
		if (bHasBorders)
		{
			CRect rect;
			GetWindowRect(rect);
			
			rect.bottom -= rect.top;
			rect.right -= rect.left;
			rect.left = rect.top = 0;
			
			CBCGPVisualManager::GetInstance()->OnDrawControlBorder(pDC, rect, this, FALSE);
		}
	}

	return 0;
}
//*******************************************************************************
LRESULT CBCGPSliderCtrl::OnSetTicFreq(WPARAM wp, LPARAM)
{
	m_nTicFreq = (int)wp;
	return Default();	
}
//*******************************************************************************
CRect CBCGPSliderCtrl::GetSelectionRect()
{
	DWORD dwStyle = GetStyle ();
	BOOL bVert = (dwStyle & TBS_VERT);
	BOOL bIsAutoTicks = (dwStyle & TBS_AUTOTICKS);
	BOOL bThumb = ((GetStyle() & TBS_NOTHUMB) == 0);

	CRect rectSel;
	rectSel.SetRectEmpty();

	if ((dwStyle & TBS_ENABLESELRANGE) == 0)
	{
		return rectSel;
	}

	int nRangeMin = GetRangeMin();
	int nRangeMax = GetRangeMax();
	
	if (nRangeMin >= nRangeMax)
	{
		return rectSel;
	}
	
	int nMin = 0;
	int nMax = 0;

	GetSelection(nMin, nMax);

	nMin = bcg_clamp(nMin, nRangeMin, nRangeMax);
	nMax = bcg_clamp(nMax, nRangeMin, nRangeMax);

	if (nMin >= nMax)
	{
		return rectSel;
	}

	CRect rectChannel;
	GetChannelRect (rectChannel);

	CRect rectThumb(0,0,0,0);
	
	if (bThumb)
	{
		GetThumbRect(rectThumb);
	}

	if (bVert)
	{
		CRect rect = rectChannel;

		rectChannel.left = rect.top;
		rectChannel.right = rect.bottom;
		rectChannel.top = rect.left;
		rectChannel.bottom = rect.right;

		double dblRatio = (double)(rectChannel.Height() - rectThumb.Height() / 2) / (nRangeMax - nRangeMin);
		
		rectSel = rectChannel;

		rectSel.DeflateRect(1, 0);
		rectSel.top += (int)(0.5 + dblRatio * (nMin - nRangeMin) + rectThumb.Height() / 2);
		rectSel.bottom = rectChannel.top + (int)(0.5 + dblRatio * (nMax - nRangeMin));
	}
	else
	{
		double dblRatio = (double)(rectChannel.Width() - rectThumb.Width() / 2) / (nRangeMax - nRangeMin);

		rectSel = rectChannel;

		rectSel.DeflateRect(0, 1);
		
		rectSel.left += (int)(0.5 + dblRatio * (nMin - nRangeMin)+ rectThumb.Width() / 2);
		rectSel.right = rectChannel.left + (int)(0.5 + dblRatio * (nMax - nRangeMin));
	}

	// Align to tick marks:

	int nNumTics = GetNumTics();
	int nTicFreq = (nNumTics == 2) ? (nRangeMax - nRangeMin) : m_nTicFreq;
	BOOL bStartIsReady = FALSE;
	int nPrevTicPos = -1;

	int nStartPos = bVert ? rectSel.top : rectSel.left;
	int nFinishPos = bVert ? rectSel.bottom : rectSel.right;
	
	for (int i = 0; i < nNumTics; i++)
	{
		int nTicPos = GetTicPos(bIsAutoTicks ? i * nTicFreq : i);
		if (nTicPos < 0)
		{
			continue;
		}

		if (!bStartIsReady)
		{
			if (nTicPos == nStartPos)
			{
				bStartIsReady = TRUE;
			}
			else
			{
				if (nTicPos > nStartPos)
				{
					if (nPrevTicPos >= 0)
					{
						if (bVert)
						{
							rectSel.top = nPrevTicPos + 1;
						}
						else
						{
							rectSel.left = nPrevTicPos + 1;
						}
					}

					bStartIsReady = TRUE;
				}
			}
		}

		if (nTicPos > nFinishPos)
		{
			if (bVert)
			{
				rectSel.bottom = nTicPos;
			}
			else
			{
				rectSel.right = nTicPos;
			}
			break;
		}

		nPrevTicPos = nTicPos;
	}

	return rectSel;
}
//*******************************************************************************
void CBCGPSliderCtrl::OnDrawThumb(CDC* pDC, CRect rectThumb, BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled, BOOL bVert, BOOL bLeftTop, BOOL bRightBottom, BOOL bDrawOnGlass)
{
	CBCGPVisualManager::GetInstance ()->OnDrawSliderThumb(
		pDC, this, rectThumb, bIsHighlighted, bIsPressed, bIsDisabled,
		bVert, bLeftTop, bRightBottom, bDrawOnGlass);
}
//*******************************************************************************
void CBCGPSliderCtrl::OnDrawChannel(CDC* pDC, BOOL bVert, CRect rect, BOOL bDrawOnGlass)
{
	CBCGPVisualManager::GetInstance ()->OnDrawSliderChannel (pDC, this, bVert, rect, bDrawOnGlass);
}
//*******************************************************************************
void CBCGPSliderCtrl::EnableZoomButtons(BOOL bEnable, UINT nAutoRepeatDelay, int nZoomStepValue)
{
	m_bZoomButtons = bEnable;
	m_rectZoomInBtn.SetRectEmpty();
	m_rectZoomOutBtn.SetRectEmpty();
	m_nAutoRepeatDelay = nAutoRepeatDelay;
	m_nZoomStepValue = nZoomStepValue;

	if (GetSafeHwnd() != NULL)
	{
		SetWindowPos(NULL, -1, -1, -1, -1, SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	}
	else
	{
		m_bIsDelayedLayout = TRUE;
	}
}
//*******************************************************************************
void CBCGPSliderCtrl::OnNcCalcSize(BOOL bCalcValidRects, NCCALCSIZE_PARAMS FAR* lpncsp) 
{
	CSliderCtrl::OnNcCalcSize(bCalcValidRects, lpncsp);

	if (m_bZoomButtons)
	{
		if (GetStyle() & TBS_VERT)
		{
			lpncsp->rgrc[0].top += m_nZoomButtonSize;
			lpncsp->rgrc[0].bottom -= m_nZoomButtonSize;
		}
		else
		{
			lpncsp->rgrc[0].left += m_nZoomButtonSize;
			lpncsp->rgrc[0].right -= m_nZoomButtonSize;
		}
	}
}
//*******************************************************************************
void CBCGPSliderCtrl::OnNcMouseMove(UINT nHitTest, CPoint point) 
{
	if (m_nCapturedZoomButton < 0 && HasZoomButtons())
	{
		CPoint ptClient = point;
		ScreenToClient (&ptClient);

		int nHighlightedButton = HitTestZoomButton(ptClient);
		
		if (nHighlightedButton >= 0)
		{
			SetCapture ();
			
			m_nHighlightedZoomButton = nHighlightedButton;
			RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
		}
	}
	
	CSliderCtrl::OnNcMouseMove(nHitTest, point);
}
//*******************************************************************************
void CBCGPSliderCtrl::OnNcPaint() 
{
	BOOL bHasBorders = (GetStyle () & WS_BORDER) != 0 || (GetExStyle () & WS_EX_CLIENTEDGE) != 0;

	if (!m_bZoomButtons)
	{
		if (bHasBorders)
		{
			CBCGPVisualManager::GetInstance()->OnDrawControlBorder(this);
		}

		return;
	}

	CWindowDC dc(this);
	DoNcPaint(&dc);
}
//*******************************************************************************
void CBCGPSliderCtrl::DoNcPaint(CDC* pDC)
{
	ASSERT_VALID(pDC);

	BOOL bIsVert = (GetStyle() & TBS_VERT) != 0;
	BOOL bHasBorders = (GetStyle () & WS_BORDER) != 0 || (GetExStyle () & WS_EX_CLIENTEDGE) != 0;

	CRect rectWindow;
	GetWindowRect(rectWindow);

	CRect rectClient;
	GetClientRect (rectClient);
	ClientToScreen (&rectClient);

	CRect rectChannel;
	GetChannelRect(rectChannel);
	
	m_rectZoomOutBtn = rectWindow;
	m_rectZoomInBtn = rectWindow;

	int nButtonOffsetX = 0;
	int nButtonOffsetY = 0;

	if (bIsVert)
	{
		m_rectZoomOutBtn.bottom = m_rectZoomOutBtn.top + m_nZoomButtonSize;
		m_rectZoomOutBtn.right = m_rectZoomOutBtn.left + m_nZoomButtonSize;
		
		m_rectZoomInBtn.top = m_rectZoomInBtn.bottom - m_nZoomButtonSize;
		m_rectZoomInBtn.right = m_rectZoomInBtn.left + m_nZoomButtonSize;
		
		nButtonOffsetX = max(0, rectChannel.CenterPoint().y - m_nZoomButtonSize / 2);
		nButtonOffsetY = rectClient.bottom + m_nZoomButtonSize - rectWindow.bottom;
	}
	else
	{
		m_rectZoomOutBtn.right = m_rectZoomOutBtn.left + m_nZoomButtonSize;
		m_rectZoomOutBtn.bottom = m_rectZoomOutBtn.top + m_nZoomButtonSize;
		
		m_rectZoomInBtn.left = m_rectZoomInBtn.right - m_nZoomButtonSize;
		m_rectZoomInBtn.bottom = m_rectZoomInBtn.top + m_nZoomButtonSize;

		nButtonOffsetX = rectClient.right + m_nZoomButtonSize - rectWindow.right;
		nButtonOffsetY = max(0, rectChannel.CenterPoint().y - m_nZoomButtonSize / 2);
	}

	if (bHasBorders)
	{
		nButtonOffsetX += 2;
		nButtonOffsetY += 2;
	}

	m_rectZoomInBtn.OffsetRect(nButtonOffsetX, nButtonOffsetY);
	m_rectZoomOutBtn.OffsetRect(nButtonOffsetX, nButtonOffsetY);
	
	if (bIsVert)
	{
		m_rectZoomInBtn.left += rectClient.left - rectWindow.left;
		m_rectZoomInBtn.right -= rectWindow.right - rectClient.right;
		
		m_rectZoomOutBtn.left += rectClient.left - rectWindow.left;
		m_rectZoomOutBtn.right -= rectWindow.right - rectClient.right;
	}
	else
	{
		m_rectZoomInBtn.top += rectClient.top - rectWindow.top;
		m_rectZoomInBtn.bottom -= rectWindow.bottom - rectClient.bottom;
		
		m_rectZoomOutBtn.top += rectClient.top - rectWindow.top;
		m_rectZoomOutBtn.bottom -= rectWindow.bottom - rectClient.bottom;
	}

	CRect rect = m_rectZoomInBtn;
	rect.OffsetRect (-rectWindow.left, -rectWindow.top);
	
    rectClient.OffsetRect(-rectWindow.TopLeft());

	pDC->ExcludeClipRect(rectClient);
	
	if (!CBCGPVisualManager::GetInstance ()->OnFillParentBarBackground (this, pDC, NULL, TRUE))
	{
		CRect rectWnd;
		GetWindowRect(rectWnd);

		ScreenToClient(rectWnd);

		CPoint pt = rectWnd.TopLeft();

		CWnd* pParent = GetParent();
		MapWindowPoints (pParent, &pt, 1);
		pt = pDC->OffsetWindowOrg (pt.x, pt.y);
		
		pParent->SendMessage (WM_ERASEBKGND, (WPARAM)pDC->m_hDC);
		
		pDC->SetWindowOrg(pt.x, pt.y);
	}

	CBCGPDrawOnGlass dog(m_bOnGlass);

	int nZoomOutIndex = 0;
	int nZoomInIndex = 1;

	if (GetExStyle() & WS_EX_LAYOUTRTL)
	{
		nZoomOutIndex = 1;
		nZoomInIndex = 0;
	}

	OnDrawZoomButton(pDC, rect, FALSE, m_nPressedZoomButton == nZoomInIndex, m_nHighlightedZoomButton == nZoomInIndex);

	rect = m_rectZoomOutBtn;
	rect.OffsetRect (-rectWindow.left, -rectWindow.top);

	OnDrawZoomButton(pDC, rect, TRUE, m_nPressedZoomButton == nZoomOutIndex, m_nHighlightedZoomButton == nZoomOutIndex);
	
	if (bHasBorders)
	{
		CRect rectWnd;
		GetWindowRect(rectWnd);
		
		rectWnd.bottom -= rectWnd.top;
		rectWnd.right -= rectWnd.left;
		rectWnd.left = rectWnd.top = 0;
		
		CBCGPVisualManager::GetInstance()->OnDrawControlBorder(pDC, rectWnd, this, m_bOnGlass);
	}

	ScreenToClient (&m_rectZoomInBtn);
	ScreenToClient (&m_rectZoomOutBtn);

	pDC->SelectClipRgn(NULL);
}
//*******************************************************************************
void CBCGPSliderCtrl::OnDrawZoomButton(CDC* pDC, CRect rect, BOOL bIsZoomOut, BOOL bIsPressed, BOOL bIsHighlighted)
{
	ASSERT_VALID(pDC);

#ifndef BCGP_EXCLUDE_RIBBON
#ifndef _BCGSUITE_
	BOOL bIsVert = (GetStyle() & TBS_VERT) != 0;

	CBCGPRibbonSlider slider(0, 100, bIsVert ? TBS_VERT : 0);
	CBCGPVisualManager::GetInstance()->OnDrawRibbonSliderZoomButton(pDC, &slider, rect, bIsZoomOut, bIsHighlighted, bIsPressed, !IsWindowEnabled());
#else
	CMFCRibbonSlider slider(0);
	CMFCVisualManager::GetInstance()->OnDrawRibbonSliderZoomButton(pDC, &slider, rect, bIsZoomOut, bIsHighlighted, bIsPressed, !IsWindowEnabled());
#endif
#endif
}
//*******************************************************************************
int CBCGPSliderCtrl::HitTestZoomButton(CPoint point)
{
	return m_rectZoomOutBtn.PtInRect(point) ? 0 : m_rectZoomInBtn.PtInRect(point) ? 1 : -1;
}
//*******************************************************************************
BCGNcHitTestType CBCGPSliderCtrl::OnNcHitTest(CPoint point) 
{
	if (HasZoomButtons())
	{
		CPoint ptClient = point;
		ScreenToClient (&ptClient);

		if (HitTestZoomButton(ptClient) >= 0)
		{
			return HTBORDER;
		}
	}
	
	return CSliderCtrl::OnNcHitTest(point);
}
//*******************************************************************************
void CBCGPSliderCtrl::OnClickZoomButton(BOOL bIsZoomOut)
{
	if (GetExStyle() & WS_EX_LAYOUTRTL)
	{
		bIsZoomOut = !bIsZoomOut;
	}

	int nPos = GetPos();
	int nZoomStepValue = max(1, (m_nZoomStepValue > 0 ? m_nZoomStepValue : GetPageSize()));

	if (bIsZoomOut)
	{
		if (nPos - nZoomStepValue <= GetRangeMin())
		{
			SendMessage(WM_KEYDOWN, VK_HOME);
		}
		else
		{
			SetPos(nPos - nZoomStepValue + 1);
			SendMessage(WM_KEYDOWN, VK_LEFT);
		}
	}
	else
	{
		if (nPos + nZoomStepValue >= GetRangeMax())
		{
			SendMessage(WM_KEYDOWN, VK_END);
		}
		else
		{
			SetPos(nPos + nZoomStepValue - 1);
			SendMessage(WM_KEYDOWN, VK_RIGHT);
		}
	}
}
//***********************************************************************************
void CBCGPSliderCtrl::OnTimer(UINT_PTR nIDEvent)
{
	if (nIDEvent == IdAutoCommand)
	{
		if (m_nPressedZoomButton == m_nHighlightedZoomButton && m_nPressedZoomButton >= 0)
		{
			OnClickZoomButton(m_nPressedZoomButton == 0);
			m_bAutoCommandSent = TRUE;
		}
	}

	CSliderCtrl::OnTimer(nIDEvent);
}
//*******************************************************************************
void CBCGPSliderCtrl::PreSubclassWindow() 
{
	CSliderCtrl::PreSubclassWindow();

	if (m_bIsDelayedLayout)
	{
		PostMessage(UM_REDRAWFRAME);
	}
}
//*******************************************************************************
int CBCGPSliderCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CSliderCtrl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	if (m_bIsDelayedLayout)
	{
		PostMessage(UM_REDRAWFRAME);
	}

	return 0;
}
//**************************************************************************
LRESULT CBCGPSliderCtrl::OnRedrawFrame(WPARAM, LPARAM)
{
	SetWindowPos(NULL, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
	RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	return 0;
}
//**************************************************************************
BOOL CBCGPSliderCtrl::PreTranslateMessage(MSG* pMsg) 
{
	if (GetFocus()->GetSafeHwnd() == GetSafeHwnd() && pMsg->message == WM_KEYDOWN && m_bZoomButtons && (pMsg->wParam == VK_ADD || pMsg->wParam == VK_SUBTRACT))
	{
		OnClickZoomButton(pMsg->wParam == VK_SUBTRACT);
		return TRUE;
	}
	
	return CSliderCtrl::PreTranslateMessage(pMsg);
}
