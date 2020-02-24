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
// BCGPScrollBar.cpp : implementation file
//

#include "stdafx.h"
#include "TrackMouse.h"
#include "BCGGlobals.h"
#include "BCGPScrollBar.h"
#include "BCGPDlgImpl.h"
#include "BCGPDrawManager.h"
#include "MenuImages.h"
#include "BCGPComboBox.h"
#include "BCGPGlobalUtils.h"

#ifndef _BCGSUITE_
#include "BCGPToolBarImages.h"
#else
#define CX_BORDER   1
#define CY_BORDER   1
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

UINT BCGM_ON_DRAW_SCROLLBAR = ::RegisterWindowMessage (_T("BCGM_ON_DRAW_SCROLLBAR"));
UINT BCGM_REDRAW_SCROLLBAR  = ::RegisterWindowMessage (_T("BCGM_REDRAW_SCROLLBAR"));

static const UINT SB_TIMER_DOWN		= 1111;
static const UINT SB_TIMER			= 2222;

static BOOL _CompareScrollInfo(SCROLLINFO& si1, SCROLLINFO& si2)
{
	return si1.nMin == si2.nMin &&
			si1.nMax == si2.nMax &&
			si1.nPage == si2.nPage &&
			si1.nPos == si2.nPos &&
			si1.nTrackPos == si2.nTrackPos;
}

static void ScreenToBar(CWnd* pWnd, CPoint& point)
{
	CRect rectWindow;
	pWnd->GetWindowRect(rectWindow);

	if (pWnd->GetExStyle() & WS_EX_LAYOUTRTL)
	{
		point.x = rectWindow.right - point.x;
	}
	else
	{
		point.x -= rectWindow.left;
	}

	point.y -= rectWindow.top;
}

#define MAX_CLASS_NAME 255

static DWORD _GetExcludeStyle(CWnd* pWnd)
{
	DWORD dwResult = 0;

	TCHAR lpszClassName[MAX_CLASS_NAME + 1];
	::GetClassName(pWnd->GetSafeHwnd(), lpszClassName, MAX_CLASS_NAME);
	if (_tcscmp(lpszClassName, _T("SysTreeView32")) == 0)
	{
		dwResult = TVS_NOSCROLL;
	}
	else if (_tcscmp(lpszClassName, _T("SysListView32")) == 0)
	{
		dwResult = LVS_NOSCROLL;
	}
	else if (_tcscmp(lpszClassName, _T("Edit")) == 0)
	{
		dwResult = ES_AUTOHSCROLL | ES_AUTOVSCROLL;
	}

	return dwResult;
};

/////////////////////////////////////////////////////////////////////////////
// CBCGPScrollBarImpl class

CBCGPScrollBarImpl::CBCGPScrollBarImpl()
{
	m_bInternal = TRUE;

	m_rect.SetRectEmpty();
	m_bHorz = FALSE;

	m_HitTest = CBCGPScrollBar::BCGP_HT_NOWHERE;
	m_HitPressed = CBCGPScrollBar::BCGP_HT_NOWHERE;

	m_rectTrackThumb.SetRectEmpty();
	m_rectTrackThumbStart.SetRectEmpty();

	m_TrackPos = -1;
	m_bTracked = FALSE;
	m_ptTrackStart = CPoint(0, 0);

	m_TimerDown = 0;
	m_Timer = 0;
	m_dwDisabledArrows = 0;

	m_pLastCapture = NULL;

	m_bSendPosBeforeEndThumb = TRUE;
}
//*******************************************************************************
BOOL CBCGPScrollBarImpl::_GetScrollInfo(CWnd* pWnd, SCROLLINFO& si) const
{
	ZeroMemory(&si, sizeof(SCROLLINFO));
	si.cbSize = sizeof(SCROLLINFO);
	si.fMask = SIF_ALL;

	int nBar = SB_CTL;
	if (m_bInternal)
	{
		nBar = m_bHorz ? SB_HORZ : SB_VERT;
	}

	return ::GetScrollInfo(pWnd->GetSafeHwnd(), nBar, &si);
}
//*******************************************************************************
int CBCGPScrollBarImpl::_GetScrollPos(CWnd* pWnd) const
{
	int nBar = SB_CTL;
	if (m_bInternal)
	{
		nBar = m_bHorz ? SB_HORZ : SB_VERT;
	}

	return ::GetScrollPos(pWnd->GetSafeHwnd(), nBar);
}
//*******************************************************************************
int CBCGPScrollBarImpl::_SetScrollPos(CWnd* pWnd, int nPos, BOOL bRedraw) const
{
	int nBar = SB_CTL;
	if (m_bInternal)
	{
		nBar = m_bHorz ? SB_HORZ : SB_VERT;
	}

	return ::SetScrollPos(pWnd->GetSafeHwnd(), nBar, nPos, bRedraw);
}
//*******************************************************************************
void CBCGPScrollBarImpl::_RedrawScrollBar(CWnd* pWnd) const
{
	if (m_bInternal)
	{
		pWnd->SendMessage(BCGM_REDRAW_SCROLLBAR, m_bHorz ? SB_HORZ : SB_VERT);
	}
	else
	{
		pWnd->RedrawWindow();
	}
}
//*******************************************************************************
void CBCGPScrollBarImpl::Render(CDC* pDC, BOOL bIsEnabled, const SCROLLINFO& si, const CPoint& ptOffset)
{
	if (m_rect.IsRectEmpty())
	{
		return;
	}

	CRect rect = m_rect;
	rect.OffsetRect(ptOffset);

	BOOL bDisabled = !bIsEnabled || (m_dwDisabledArrows & ESB_DISABLE_BOTH) == ESB_DISABLE_BOTH;
	if (!bDisabled)
	{
		bDisabled = si.nMax <= si.nMin || (si.nMax - si.nMin) < (int)si.nPage;
	}

	CRect rectThumb;
	if (!bDisabled)
	{
		rectThumb = GetThumbRect(bIsEnabled, si);
		rectThumb.OffsetRect(ptOffset);
	}

	CRect rectBtn[2];
	rectBtn[0] = rect;

	if (m_bHorz)
	{
		int nSize = ::GetSystemMetrics (SM_CXHSCROLL);
		if (rect.Width () < nSize * 2)
		{
			nSize = rect.Width () / 2;
		}

		rectBtn[0].right = rect.left + nSize;
		rectBtn[1] = rectBtn[0];
		rectBtn[1].OffsetRect (rect.Width () - rectBtn[1].Width (), 0);

		rect.left += rectBtn[0].Width ();
		rect.right -= rectBtn[1].Width ();

		if (rectThumb.left < rectBtn[0].right || rectBtn[1].left < rectThumb.right)
		{
			rectThumb.SetRectEmpty ();
		}
	}
	else
	{
		int nSize = ::GetSystemMetrics (SM_CYVSCROLL);
		if (rect.Height () < nSize * 2)
		{
			nSize = rect.Height () / 2;
		}

		rectBtn[0].bottom = rect.top + nSize;
		rectBtn[1] = rectBtn[0];
		rectBtn[1].OffsetRect (0, rect.Height () - rectBtn[1].Height ());

		rect.top += rectBtn[0].Height ();
		rect.bottom -= rectBtn[1].Height ();

		if (rectThumb.top < rectBtn[0].bottom || rectBtn[1].top < rectThumb.bottom)
		{
			rectThumb.SetRectEmpty ();
		}
	}

	DoPaintButton (pDC, rectBtn[0], m_bHorz, 
		m_HitTest == CBCGPScrollBar::BCGP_HT_BUTTON_LTUP || m_HitPressed == CBCGPScrollBar::BCGP_HT_BUTTON_LTUP, 
		m_HitPressed == CBCGPScrollBar::BCGP_HT_BUTTON_LTUP && m_HitTest == m_HitPressed, TRUE, 
		bDisabled || (m_dwDisabledArrows & ESB_DISABLE_LTUP) == ESB_DISABLE_LTUP);
	
	DoPaintButton (pDC, rectBtn[1], m_bHorz, 
		m_HitTest == CBCGPScrollBar::BCGP_HT_BUTTON_RTDN || m_HitPressed == CBCGPScrollBar::BCGP_HT_BUTTON_RTDN, 
		m_HitPressed == CBCGPScrollBar::BCGP_HT_BUTTON_RTDN && m_HitTest == m_HitPressed, FALSE, 
		bDisabled || (m_dwDisabledArrows & ESB_DISABLE_RTDN) == ESB_DISABLE_RTDN);

	if (!bDisabled && !rectThumb.IsRectEmpty ())
	{
		DoPaintThumb (pDC, rectThumb, m_bHorz, 
			m_HitTest == CBCGPScrollBar::BCGP_HT_THUMB || m_HitPressed == CBCGPScrollBar::BCGP_HT_THUMB,
			m_HitPressed == CBCGPScrollBar::BCGP_HT_THUMB, bDisabled);

		BOOL bDrawBack = FALSE;
		CRect rectBack (rect);
		if (m_bHorz)
		{
			rectBack.right = rectThumb.left;
			bDrawBack = rectBack.left < rectBack.right;
		}
		else
		{
			rectBack.bottom = rectThumb.top;
			bDrawBack = rectBack.top < rectBack.bottom;
		}

		if (bDrawBack)
		{
			DoPaintBackground (pDC, rectBack, m_bHorz, 
				m_HitTest == CBCGPScrollBar::BCGP_HT_CLIENT_LTUP || m_HitPressed == CBCGPScrollBar::BCGP_HT_CLIENT_LTUP, 
				m_HitPressed == CBCGPScrollBar::BCGP_HT_CLIENT_LTUP && m_HitTest == m_HitPressed, TRUE, bDisabled);
		}

		bDrawBack = FALSE;
		rectBack = rect;
		if (m_bHorz)
		{
			rectBack.left = rectThumb.right;
			bDrawBack = rectBack.left < rectBack.right;
		}
		else
		{
			rectBack.top = rectThumb.bottom;
			bDrawBack = rectBack.top < rectBack.bottom;
		}

		if (bDrawBack)
		{
			DoPaintBackground (pDC, rectBack, m_bHorz, 
				m_HitTest == CBCGPScrollBar::BCGP_HT_CLIENT_RTDN || m_HitPressed == CBCGPScrollBar::BCGP_HT_CLIENT_RTDN, 
				m_HitPressed == CBCGPScrollBar::BCGP_HT_CLIENT_RTDN && m_HitTest == m_HitPressed, FALSE, bDisabled);
		}
	}
	else if (!rect.IsRectEmpty ())
	{
		DoPaintBackground (pDC, rect, m_bHorz, FALSE, FALSE, TRUE, bDisabled);
	}
}
//*******************************************************************************
void CBCGPScrollBarImpl::DoPaintThumb (CDC* pDC, CRect rect, BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bDisabled)
{
	CBCGPScrollBar dummy;
	dummy.m_HitTest = m_HitTest != BCGP_HT_NOWHERE ? m_HitTest : m_HitPressed;

	CBCGPScrollBar* pBar = GetScrollBarCtrl();
	if (pBar == NULL)
	{
		pBar = &dummy;
	}

	CBCGPVisualManager::GetInstance ()->OnScrollBarDrawThumb (pDC, pBar, rect, 
		bHorz, bHighlighted, bPressed, bDisabled);
}
//*******************************************************************************
void CBCGPScrollBarImpl::DoPaintButton (CDC* pDC, CRect rect, BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bFirst, BOOL bDisabled)
{
	CBCGPScrollBar dummy;
	dummy.m_HitTest = m_HitTest != BCGP_HT_NOWHERE ? m_HitTest : m_HitPressed;

	CBCGPScrollBar* pBar = GetScrollBarCtrl();
	if (pBar == NULL)
	{
		pBar = &dummy;
	}

	CBCGPVisualManager::GetInstance ()->OnScrollBarDrawButton (pDC, pBar, rect, bHorz, 
							bHighlighted, bPressed, bFirst, bDisabled);
}
//*******************************************************************************
void CBCGPScrollBarImpl::DoPaintBackground (CDC* pDC, CRect rect, BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bFirst, BOOL bDisabled)
{
	CBCGPScrollBar dummy;
	dummy.m_HitTest = m_HitTest != BCGP_HT_NOWHERE ? m_HitTest : m_HitPressed;

	CBCGPScrollBar* pBar = GetScrollBarCtrl();
	if (pBar == NULL)
	{
		pBar = &dummy;
	}

	CBCGPVisualManager::GetInstance ()->OnScrollBarFillBackground (pDC, pBar, rect, bHorz, 
							bHighlighted, bPressed, bFirst, bDisabled);
}
//*****************************************************************************
CRect CBCGPScrollBarImpl::GetThumbRect(BOOL bIsEnabled, const SCROLLINFO& si) const
{
	CRect rectThumb (0, 0, 0, 0);
	
	if (m_HitPressed == CBCGPScrollBar::BCGP_HT_THUMB)
	{
		rectThumb = m_rectTrackThumb;
	}
	else
	{
		if (!bIsEnabled || 
			(m_dwDisabledArrows & ESB_DISABLE_BOTH) == ESB_DISABLE_BOTH)
		{
			return rectThumb;
		}
		
		rectThumb = m_rect;
		
		int nClientSize = 0;
		int nThumbSize	= 0;
		if (m_bHorz)
		{
			nThumbSize = ::GetSystemMetrics (SM_CXHSCROLL);
			rectThumb.DeflateRect (nThumbSize, 0);
			nClientSize = rectThumb.Width ();
		}
		else
		{
			nThumbSize = ::GetSystemMetrics (SM_CYVSCROLL);
			rectThumb.DeflateRect (0, nThumbSize);
			nClientSize = rectThumb.Height ();
		}
		
		const int c_ScrollMin = globalUtils.ScaleByDPI(4);

		if (nClientSize <= c_ScrollMin)
		{
			rectThumb.SetRectEmpty ();
			return rectThumb;
		}

        if ((si.nMax - si.nMin) < (int)si.nPage)
        {
            nThumbSize = 0;
        }

		int nThumbPos = nThumbSize;
		
		if (si.nPage != 0)
		{
			const int c_ScrollThumbMin = globalUtils.ScaleByDPI(8);
			nThumbSize = max (::MulDiv(nClientSize, si.nPage, si.nMax - si.nMin + 1), c_ScrollThumbMin);
		}
		
		if (nClientSize < nThumbSize || nThumbSize == 0)
		{
			rectThumb.SetRectEmpty ();
			return rectThumb;
		}
		else
		{
			nClientSize -= nThumbSize;
			int nMax = si.nMax - max (si.nPage - 1, 0);
			if (si.nMin < nMax)
			{
				nThumbPos += ::MulDiv(nClientSize, si.nPos - si.nMin, nMax - si.nMin);
			}
		}
		
		if (m_bHorz)
		{
			rectThumb.left	 = m_rect.left + nThumbPos;
			rectThumb.right = rectThumb.left + nThumbSize;
		}
		else
		{
			rectThumb.top	 = m_rect.top + nThumbPos;
			rectThumb.bottom = rectThumb.top + nThumbSize;
		}
	}
	
	return rectThumb;
}
//*****************************************************************************
BOOL CBCGPScrollBarImpl::ProcessEvent(CWnd* pWnd)
{
	CWnd* pWndNotify = m_bInternal ? pWnd : pWnd->GetParent ();
	if (pWndNotify->GetSafeHwnd () == NULL)
	{
		return FALSE;
	}

	BOOL bProcess = FALSE;

	UINT nMessage = m_bHorz ? WM_HSCROLL : WM_VSCROLL;
	WPARAM nWParam = 0;
	LPARAM nLParam = m_bInternal ? 0 : (LPARAM)pWnd->GetSafeHwnd ();

	if ((m_HitPressed == CBCGPScrollBar::BCGP_HT_BUTTON_LTUP ||
		 m_HitPressed == CBCGPScrollBar::BCGP_HT_BUTTON_RTDN) &&
		 m_HitPressed == m_HitTest)
	{
		BOOL bUp = m_HitPressed == CBCGPScrollBar::BCGP_HT_BUTTON_LTUP;

		if (m_bHorz)
		{
			nWParam = bUp ? SB_LINELEFT : SB_LINERIGHT;
		}
		else
		{
			nWParam = bUp ? SB_LINEUP : SB_LINEDOWN;
		}

		bProcess = TRUE;
	}
	else if ((m_HitPressed == CBCGPScrollBar::BCGP_HT_CLIENT_LTUP ||
			  m_HitPressed == CBCGPScrollBar::BCGP_HT_CLIENT_RTDN) &&
			  m_HitPressed == m_HitTest)
	{
		BOOL bUp = m_HitPressed == CBCGPScrollBar::BCGP_HT_CLIENT_LTUP;

		if (m_bHorz)
		{
			nWParam = bUp ? SB_PAGELEFT : SB_PAGERIGHT;
		}
		else
		{
			nWParam = bUp ? SB_PAGEUP : SB_PAGEDOWN;
		}

		bProcess = TRUE;
	}

	if (bProcess)
	{
		pWndNotify->SendMessage(nMessage, nWParam, nLParam);
	}

	return bProcess;
}
//*****************************************************************************
BOOL CBCGPScrollBarImpl::ProcessMouseDown(CWnd* pWnd, CPoint point)
{
	ASSERT_VALID(pWnd);

	m_pLastCapture = NULL;

	CBCGPDropDownListBox* pDropDownList = DYNAMIC_DOWNCAST(CBCGPDropDownListBox, pWnd);
	if (pDropDownList != NULL)
	{
		if (!pDropDownList->m_bIsReady)
		{
			pDropDownList->m_bIsReady = TRUE;
			return FALSE;
		}
	}

	CPoint ptScreen(point);

	if (!m_bInternal)
	{
		pWnd->ClientToScreen(&ptScreen);
	}

	CBCGPScrollBar::BCGPSB_HITTEST hit = DoHitTest(pWnd, ptScreen);

	if (hit != CBCGPScrollBar::BCGP_HT_NOWHERE)
	{
		if (pDropDownList == NULL)
		{
			m_pLastCapture = pWnd->SetCapture ();
		}

		_RedrawScrollBar(pWnd);

		if (hit != CBCGPScrollBar::BCGP_HT_THUMB)
		{
			m_HitPressed = hit;

			ProcessEvent(pWnd);

			if (m_TimerDown == 0)
			{
				const UINT nDelay = GetDoubleClickTime();
				m_TimerDown = pWnd->SetTimer(SB_TIMER_DOWN, max(nDelay, (UINT)200), NULL);
			}
		}
		else
		{
			SCROLLINFO si;
			_GetScrollInfo(pWnd, si);

			m_ptTrackStart = point;

			m_rectTrackThumbStart = GetThumbRect(pWnd->IsWindowEnabled(), si);
			m_rectTrackThumb = m_rectTrackThumbStart;

			m_TrackPos = _GetScrollPos(pWnd);

			m_HitPressed = hit;

			CWnd* pWndNotify = m_bInternal ? pWnd : pWnd->GetParent ();
			if (pWndNotify->GetSafeHwnd () != NULL)
			{
				pWndNotify->SendMessage (m_bHorz ? WM_HSCROLL : WM_VSCROLL, 
						MAKEWPARAM (SB_THUMBTRACK, m_TrackPos), (LPARAM)(m_bInternal ? 0 : pWnd->GetSafeHwnd ()));
			}
		}
	}

	return TRUE;
}
//*****************************************************************************
BOOL CBCGPScrollBarImpl::ProcessMouseUp(CWnd* pWnd, CPoint point)
{
	CBCGPDropDownListBox* pDropDownList = DYNAMIC_DOWNCAST(CBCGPDropDownListBox, pWnd);
	if (pDropDownList != NULL)
	{
		if (!pDropDownList->m_bIsReady)
		{
			return FALSE;
		}
	}

	BOOL bSendEvent = m_HitPressed != CBCGPScrollBar::BCGP_HT_NOWHERE;
	BOOL bThumb 	= m_HitPressed == CBCGPScrollBar::BCGP_HT_THUMB;
	int nTrackPos	= m_TrackPos;

	ProcessCancelMode(pWnd, point);

	CWnd* pWndNotify = m_bInternal ? pWnd : pWnd->GetParent ();
	if (bSendEvent && pWndNotify->GetSafeHwnd () != NULL)
	{
		UINT dwScroll = m_bHorz ? WM_HSCROLL : WM_VSCROLL;

		if (bThumb)
		{
			if (m_bSendPosBeforeEndThumb)
			{
				if (_GetScrollPos(pWnd) != nTrackPos)
				{
					_SetScrollPos(pWnd, nTrackPos, FALSE);
					nTrackPos = _GetScrollPos(pWnd);
				}
			}

			pWndNotify->SendMessage(dwScroll, MAKEWPARAM(SB_THUMBPOSITION, nTrackPos), (LPARAM)(m_bInternal ? 0 : pWnd->GetSafeHwnd()));
		}

		pWndNotify->SendMessage(dwScroll, SB_ENDSCROLL, (LPARAM)(m_bInternal ? 0 : pWnd->GetSafeHwnd()));
	}

	return TRUE;
}
//*****************************************************************************
BOOL CBCGPScrollBarImpl::ProcessMouseDblClk(CWnd* pWnd, CPoint point)
{
	return ProcessMouseDown(pWnd, point);
}
//*****************************************************************************
BOOL CBCGPScrollBarImpl::ProcessMouseMove(CWnd* pWnd, BOOL bIsMouseDown, CPoint point)
{
	ASSERT_VALID(pWnd);

	CPoint ptScreen(point);

	if (!m_bInternal)
	{
		pWnd->ClientToScreen(&ptScreen);
	}

	if (!bIsMouseDown)
	{
		if (!m_bTracked)
		{
			m_bTracked = TRUE;
			
			TRACKMOUSEEVENT trackmouseevent;
			trackmouseevent.cbSize = sizeof(trackmouseevent);
			trackmouseevent.dwFlags = TME_LEAVE;
			trackmouseevent.hwndTrack = pWnd->GetSafeHwnd();
			trackmouseevent.dwHoverTime = HOVER_DEFAULT;
			::BCGPTrackMouse (&trackmouseevent);
		}
	}

	CBCGPScrollBar::BCGPSB_HITTEST hit = DoHitTest(pWnd, ptScreen);

	if (m_HitTest != hit)
	{
		m_HitTest = hit;

		_RedrawScrollBar(pWnd);
	}

	if (bIsMouseDown && m_HitPressed == CBCGPScrollBar::BCGP_HT_THUMB)
	{
		CSize szClamp (0, 0);
		int nThumbLength = 0;
		int nScroll = 0;

		if (m_bHorz)
		{
			nScroll = ::GetSystemMetrics (SM_CXHSCROLL);
			szClamp.cx = m_rect.left;
			szClamp.cy = m_rect.right;
			nThumbLength = m_rectTrackThumbStart.Width();
		}
		else
		{
			nScroll = ::GetSystemMetrics (SM_CYVSCROLL);
			szClamp.cx = m_rect.top;
			szClamp.cy = m_rect.bottom;
			nThumbLength = m_rectTrackThumbStart.Height();
		}

		szClamp.cx += nScroll;
		szClamp.cy -= (nScroll + nThumbLength);

		CPoint ptOffset (point - m_ptTrackStart);
		CRect rectNew (m_rectTrackThumbStart);

		if (m_bHorz)
		{
			if (abs(ptOffset.y) < 150)
			{
				rectNew.OffsetRect (ptOffset.x, 0);

				if (rectNew.left < szClamp.cx)
				{
					rectNew.left = szClamp.cx;
					rectNew.right = rectNew.left + nThumbLength;
				}
				else if (szClamp.cy < rectNew.left)
				{
					rectNew.left = szClamp.cy;
					rectNew.right = rectNew.left + nThumbLength;
				}
			}
		}
		else
		{
			if (abs(ptOffset.x) < 150)
			{
				rectNew.OffsetRect (0, ptOffset.y);

				if (rectNew.top < szClamp.cx)
				{
					rectNew.top = szClamp.cx;
					rectNew.bottom = rectNew.top + nThumbLength;
				}
				else if (szClamp.cy < rectNew.top)
				{
					rectNew.top = szClamp.cy;
					rectNew.bottom = rectNew.top + nThumbLength;
				}
			}
		}

		if (rectNew != m_rectTrackThumb)
		{
			SCROLLINFO si;
			_GetScrollInfo(pWnd, si);

			m_rectTrackThumb = rectNew;

			int nPosNew = PosFromThumb 
				((m_bHorz ? m_rectTrackThumb.left : m_rectTrackThumb.top) - szClamp.cx, szClamp, si);

			if (m_TrackPos != nPosNew)
			{
				m_TrackPos = nPosNew;

				CWnd* pWndNotify = m_bInternal ? pWnd : pWnd->GetParent ();
				if (pWndNotify->GetSafeHwnd () != NULL)
				{
					pWndNotify->SendMessage (m_bHorz ? WM_HSCROLL : WM_VSCROLL, 
						MAKEWPARAM (SB_THUMBTRACK, m_TrackPos), (LPARAM)(m_bInternal ? 0 : pWnd->GetSafeHwnd ()));
				}
			}

			_RedrawScrollBar(pWnd);
		}
	}

	return TRUE;
}
//*****************************************************************************
void CBCGPScrollBarImpl::ProcessMouseLeave(CWnd* pWnd)
{
	if (m_bTracked)
	{
		m_bTracked = FALSE;
		m_HitTest = CBCGPScrollBar::BCGP_HT_NOWHERE;
		
		_RedrawScrollBar(pWnd);
	}
}
//*****************************************************************************
BOOL CBCGPScrollBarImpl::ProcessCancelMode(CWnd* pWnd, CPoint point)
{
	ASSERT_VALID(pWnd);

	if (CWnd::GetCapture () == pWnd && DYNAMIC_DOWNCAST(CBCGPDropDownListBox, pWnd) == NULL)
	{
		ReleaseCapture ();
	}
	
	if (m_pLastCapture != NULL)
	{
		m_pLastCapture->SetCapture ();
		m_pLastCapture = NULL;
	}
	
	if (m_TimerDown != 0)
	{
		pWnd->KillTimer (SB_TIMER_DOWN);
		m_TimerDown = 0;
	}
	
	if (m_Timer != 0)
	{
		pWnd->KillTimer (SB_TIMER);
		m_Timer = 0;
	}

	m_ptTrackStart = CPoint (0, 0);

	CPoint ptScreen(point);

	if (!m_bInternal)
	{
		pWnd->ClientToScreen(&ptScreen);
	}

	CBCGPScrollBar::BCGPSB_HITTEST hit = DoHitTest(pWnd, ptScreen);

	if (m_HitPressed != CBCGPScrollBar::BCGP_HT_NOWHERE ||
		m_HitTest != hit)
	{
		m_HitTest	 = hit;
		m_HitPressed = CBCGPScrollBar::BCGP_HT_NOWHERE;

		_RedrawScrollBar(pWnd);
	}

	return TRUE;
}
//*****************************************************************************
BOOL CBCGPScrollBarImpl::ProcessTimer(CWnd* pWnd, UINT_PTR nIDEvent)
{
	ASSERT_VALID(pWnd);

	if (nIDEvent == SB_TIMER_DOWN && m_TimerDown != 0)
	{
		m_TimerDown = 0;
		
		pWnd->KillTimer(SB_TIMER_DOWN);

		ProcessEvent(pWnd);

		const UINT nSpeed = GetDoubleClickTime() / 8;
		m_Timer = pWnd->SetTimer(SB_TIMER, max(nSpeed, (UINT)50), NULL);

		return TRUE;
	}
	else if (nIDEvent == SB_TIMER && m_Timer != 0)
	{
		ProcessEvent(pWnd);
		return TRUE;
	}

	return FALSE;
}
//*****************************************************************************
int CBCGPScrollBarImpl::PosFromThumb (int thumb, const CSize& szClamp, const SCROLLINFO& si)
{
	const int nRange1 = szClamp.cy - szClamp.cx;
	const int nMin	 = si.nMin;
	const int nMax	 = si.nMax - (si.nPage - 1);
	const int nRange2 = nMax - nMin;

	if (nRange2 < 0)
	{
		return 0;
	}

	int nPos = nMin + 
		(int)(((double)thumb) * ((double)nRange2) / (double)(nRange1) + 0.5);

	if (nPos < nMin)
	{
		nPos = nMin;
	}

	if (nMax < nPos)
	{
		nPos = nMax;
	}

	return nPos;
}
//*****************************************************************************
CBCGPScrollBarImpl::BCGPSB_HITTEST CBCGPScrollBarImpl::DoHitTest(CWnd* pWnd, const CPoint& pt) const
{
	CRect rect = m_rect;

	CPoint point(pt);

	if (!m_bInternal)
	{
		pWnd->ScreenToClient(&point);
	}

	CBCGPScrollBar::BCGPSB_HITTEST hit = CBCGPScrollBar::BCGP_HT_NOWHERE;

	if (rect.PtInRect (point))
	{
		hit = CBCGPScrollBar::BCGP_HT_CLIENT;

		SCROLLINFO si;
		_GetScrollInfo(pWnd, si);

		CRect rectThumb (GetThumbRect(pWnd->IsWindowEnabled(), si));

		CRect rectBtn[2];
		rectBtn[0] = rect;

		if (m_bHorz)
		{
			int nSize = ::GetSystemMetrics (SM_CXHSCROLL);
			if (rect.Width () < nSize * 2)
			{
				nSize = rect.Width () / 2;
			}

			rectBtn[0].right = rect.left + nSize;
			rectBtn[1] = rectBtn[0];
			rectBtn[1].OffsetRect (rect.Width () - rectBtn[1].Width (), 0);

			rect.left += rectBtn[0].Width ();
			rect.right -= rectBtn[1].Width ();

			if (rectThumb.left < rectBtn[0].right || rectBtn[1].left < rectThumb.right)
			{
				rectThumb.SetRectEmpty ();
			}
		}
		else
		{
			int nSize = ::GetSystemMetrics (SM_CYVSCROLL);
			if (rect.Height () < nSize * 2)
			{
				nSize = rect.Height () / 2;
			}

			rectBtn[0].bottom = rect.top + nSize;
			rectBtn[1] = rectBtn[0];
			rectBtn[1].OffsetRect (0, rect.Height () - rectBtn[1].Height ());
			
			rect.top += rectBtn[0].Height ();
			rect.bottom -= rectBtn[1].Height ();

			if (rectThumb.top < rectBtn[0].bottom || rectBtn[1].top < rectThumb.bottom)
			{
				rectThumb.SetRectEmpty ();
			}
		}

		if (rectBtn[0].PtInRect (point))
		{
			hit = m_bHorz ? CBCGPScrollBar::BCGP_HT_BUTTON_LEFT : CBCGPScrollBar::BCGP_HT_BUTTON_UP;
		}
		else
		{
			if (rectBtn[1].PtInRect (point))
			{
				hit = m_bHorz ? CBCGPScrollBar::BCGP_HT_BUTTON_RIGHT : CBCGPScrollBar::BCGP_HT_BUTTON_DOWN;
			}
			else if (!rectThumb.IsRectEmpty ())
			{
				if (rectThumb.PtInRect (point))
				{
					hit = CBCGPScrollBar::BCGP_HT_THUMB;
				}
				else
				{
					if (m_bHorz)
					{
						if (point.x < rectThumb.left)
						{
							hit = CBCGPScrollBar::BCGP_HT_CLIENT_LEFT;
						}
						else
						{
							hit = CBCGPScrollBar::BCGP_HT_CLIENT_RIGHT;
						}
					}
					else
					{
						if (point.y < rectThumb.top)
						{
							hit = CBCGPScrollBar::BCGP_HT_CLIENT_UP;
						}
						else
						{
							hit = CBCGPScrollBar::BCGP_HT_CLIENT_DOWN;
						}
					}
				}
			}
		}
	}

	return hit;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPScrollBar

IMPLEMENT_DYNAMIC(CBCGPScrollBar, CScrollBar)

CBCGPScrollBar::CBCGPScrollBar()
	: m_Style		 (CBCGPScrollBar::BCGP_SBSTYLE_VISUAL_MANAGER) 
	, m_bAutoDestoyWindow (FALSE)
	, m_bOnGlass	 (FALSE)
{
	m_bInternal = FALSE;
}

CBCGPScrollBar::~CBCGPScrollBar()
{
}


BEGIN_MESSAGE_MAP(CBCGPScrollBar, CScrollBar)
	//{{AFX_MSG_MAP(CBCGPScrollBar)
	ON_WM_CREATE()
	ON_WM_PAINT()
	ON_WM_CANCELMODE()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_LBUTTONDBLCLK()
	ON_WM_RBUTTONUP()
	ON_WM_MOUSEMOVE()
	ON_MESSAGE(WM_MOUSELEAVE, OnMouseLeave)
	ON_WM_TIMER()
	ON_WM_ERASEBKGND()
	ON_WM_NCDESTROY()
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLAERO, OnBCGSetControlAero)
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_PRINTCLIENT, OnPrintClient)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPScrollBar message handlers

void CBCGPScrollBar::DoPaint (CDC* pDC)
{
	CBCGPDrawOnGlass dog(m_bOnGlass);

	GetClientRect(m_rect);
	m_bHorz = (GetStyle() & SBS_VERT) == 0;

	SCROLLINFO si;
	si.cbSize = sizeof(SCROLLINFO);
	GetScrollInfo(&si, SIF_POS | SIF_RANGE | SIF_PAGE);
	
	Render(pDC, IsWindowEnabled(), si);
}
//*******************************************************************************
void CBCGPScrollBar::DoPaintThumb (CDC* pDC, CRect rect, BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bDisabled)
{
	if (m_Style == BCGP_SBSTYLE_OWNER_DRAW)
	{
		BCGPSCROLLBAR_DRAW_INFO info (pDC, this, BCGP_SCROLLBAR_THUMB, rect, 
			bHorz, bHighlighted, bPressed, FALSE, bDisabled);

		ASSERT_VALID (GetOwner ());

		GetOwner ()->SendMessage (BCGM_ON_DRAW_SCROLLBAR, (WPARAM) GetDlgCtrlID (), (LPARAM) &info);
		return;
	}

	if (m_Style == BCGP_SBSTYLE_CUSTOM)
	{
		DrawCustomBox(pDC, rect, bHorz, bHighlighted, bPressed, bDisabled);
		return;
	}

	CBCGPVisualManager::GetInstance ()->OnScrollBarDrawThumb (pDC, this, rect, 
		bHorz, bHighlighted, bPressed, bDisabled);
}
//*******************************************************************************
void CBCGPScrollBar::DoPaintButton (CDC* pDC, CRect rect, BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bFirst, BOOL bDisabled)
{
	if (m_Style == BCGP_SBSTYLE_OWNER_DRAW)
	{
		BCGPSCROLLBAR_DRAW_INFO info (pDC, this, BCGP_SCROLLBAR_BUTTON, rect, 
			bHorz, bHighlighted, bPressed, bFirst, bDisabled);

		ASSERT_VALID (GetOwner ());

		GetOwner ()->SendMessage (BCGM_ON_DRAW_SCROLLBAR, (WPARAM) GetDlgCtrlID (), (LPARAM) &info);
		return;
	}

	if (bHorz && (GetExStyle() & WS_EX_LAYOUTRTL))
	{
		bFirst = !bFirst;
	}

	if (m_Style == BCGP_SBSTYLE_CUSTOM)
	{
		COLORREF clrFill = DrawCustomBox(pDC, rect, bHorz, bHighlighted, bPressed, bDisabled);
		
		CBCGPMenuImages::IMAGES_IDS ids;
		if (bHorz)
		{
			ids = bFirst ? CBCGPMenuImages::IdArowLeftTab3d : CBCGPMenuImages::IdArowRightTab3d;
		}
		else
		{
			ids = bFirst ? CBCGPMenuImages::IdArowUpLarge : CBCGPMenuImages::IdArowDownLarge;
		}
		
		if (bDisabled)
		{
			CBCGPMenuImages::IMAGE_STATE state = CBCGPDrawManager::IsDarkColor(clrFill) ? CBCGPMenuImages::ImageDkGray : CBCGPMenuImages::ImageLtGray;
			CBCGPMenuImages::Draw (pDC, ids, rect, state);
		}
		else
		{
			CBCGPMenuImages::DrawByColor(pDC, ids, rect, clrFill);
		}

		return;
	}

	CBCGPVisualManager::GetInstance ()->OnScrollBarDrawButton (pDC, this, rect, bHorz, 
							bHighlighted, bPressed, bFirst, bDisabled);
}
//*******************************************************************************
void CBCGPScrollBar::DoPaintBackground (CDC* pDC, CRect rect, BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bFirst, BOOL bDisabled)
{
	if (m_Style == BCGP_SBSTYLE_OWNER_DRAW)
	{
		BCGPSCROLLBAR_DRAW_INFO info (pDC, this, BCGP_SCROLLBAR_BACKGROUND, rect, 
			bHorz, bHighlighted, bPressed, bFirst, bDisabled);

		ASSERT_VALID (GetOwner ());

		GetOwner ()->SendMessage (BCGM_ON_DRAW_SCROLLBAR, (WPARAM) GetDlgCtrlID (), (LPARAM) &info);
		return;
	}

	if (m_Style == BCGP_SBSTYLE_CUSTOM)
	{
		CBCGPDrawManager dm (*pDC);

		COLORREF clrFace = (bPressed && m_ColorTheme.m_clrFacePressed != (COLORREF)-1) ? m_ColorTheme.m_clrFacePressed : m_ColorTheme.m_clrFace;
		dm.DrawRect (rect, clrFace, bDisabled ? (COLORREF)-1 : m_ColorTheme.m_clrBorder);
		return;
	}

	CBCGPVisualManager::GetInstance ()->OnScrollBarFillBackground (pDC, this, rect, bHorz, 
							bHighlighted, bPressed, bFirst, bDisabled);
}
//*****************************************************************************
int CBCGPScrollBar::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CScrollBar::OnCreate(lpCreateStruct) == -1)
	{
		return -1;
	}
	
	return 0;
}
//*****************************************************************************
void CBCGPScrollBar::PreSubclassWindow() 
{
	CScrollBar::PreSubclassWindow();
}
//*****************************************************************************
void CBCGPScrollBar::OnPaint() 
{
	if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
		m_Style == BCGP_SBSTYLE_DEFAULT)
	{
		Default ();
		return;
	}

	CPaintDC dc(this); // device context for painting
	
	CBCGPMemDC memDC (dc, this);
	CDC* pDC = &memDC.GetDC ();

	DoPaint (pDC);
}
//*****************************************************************************
void CBCGPScrollBar::OnCancelMode() 
{
	CScrollBar::OnCancelMode();

	CPoint point;
	::GetCursorPos (&point);
	ScreenToClient(&point);

	ProcessCancelMode(this, point);
}
//*****************************************************************************
void CBCGPScrollBar::Process ()
{
	ProcessEvent(this);
}
//*****************************************************************************
void CBCGPScrollBar::OnLButtonDown(UINT /*nFlags*/, CPoint point) 
{
	if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
		m_Style == BCGP_SBSTYLE_DEFAULT)
	{
		Default ();
		return;
	}

	ProcessMouseDown(this, point);
}
//*****************************************************************************
void CBCGPScrollBar::OnLButtonUp(UINT /*nFlags*/, CPoint point) 
{
	if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
		m_Style == BCGP_SBSTYLE_DEFAULT)
	{
		Default ();
		return;
	}

	ProcessMouseUp(this, point);
}
//*****************************************************************************
void CBCGPScrollBar::OnLButtonDblClk(UINT /*nFlags*/, CPoint point) 
{
	if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
		m_Style == BCGP_SBSTYLE_DEFAULT)
	{
		Default ();
		return;
	}

	ProcessMouseDblClk(this, point);
}
//*****************************************************************************
void CBCGPScrollBar::OnRButtonUp(UINT /*nFlags*/, CPoint /*point*/)
{
	if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
		m_Style == BCGP_SBSTYLE_DEFAULT)
	{
		Default ();
	}
}
//*****************************************************************************
void CBCGPScrollBar::OnMouseMove(UINT nFlags, CPoint point) 
{
	if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
		m_Style == BCGP_SBSTYLE_DEFAULT)
	{
		Default ();
		return;
	}
	
	ProcessMouseMove(this, (nFlags & MK_LBUTTON) == MK_LBUTTON, point);
}
//*****************************************************************************
LRESULT CBCGPScrollBar::OnMouseLeave(WPARAM, LPARAM)
{
	if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
		m_Style == BCGP_SBSTYLE_DEFAULT)
	{
		return Default ();
	}

	ProcessMouseLeave(this);
	return 0;
}
//*****************************************************************************
LRESULT CBCGPScrollBar::WindowProc(UINT message, WPARAM wParam, LPARAM lParam) 
{
	if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
		m_Style == BCGP_SBSTYLE_DEFAULT)
	{
		return CScrollBar::WindowProc (message, wParam, lParam);
	}

	if (GetSafeHwnd () == NULL)
	{
		return 0L;
	}

	LRESULT lRes = 0L;

	if (message == SBM_SETPOS || message == SBM_SETRANGE ||
		message == SBM_SETRANGEREDRAW || message == SBM_SETSCROLLINFO)
	{
		BOOL bRedraw = FALSE;

		if (m_HitPressed != CBCGPScrollBar::BCGP_HT_THUMB)
		{
			SCROLLINFO si;
			si.cbSize = sizeof(SCROLLINFO);
			GetScrollInfo(&si);

			if (message == SBM_SETPOS)
			{
				if (lParam == TRUE)
				{
					bRedraw = (int)wParam != si.nPos;
				}
			}
			else if (message == SBM_SETRANGEREDRAW)
			{
				bRedraw = (int)wParam != si.nMin || (int)lParam != si.nMax;
			}
			else if (message == SBM_SETSCROLLINFO)
			{
				LPSCROLLINFO lpSI = (LPSCROLLINFO)lParam;

				if (wParam == TRUE && lpSI != NULL)
				{
					bRedraw = ((si.fMask & SIF_PAGE) == SIF_PAGE && si.nPage != lpSI->nPage) ||
						((si.fMask & SIF_POS) == SIF_POS && si.nPos != lpSI->nPos) ||
						((si.fMask & SIF_RANGE) == SIF_RANGE && (si.nMin != lpSI->nMin || si.nMax != lpSI->nMax)) ||
						((si.fMask & SIF_TRACKPOS) == SIF_TRACKPOS && si.nTrackPos != lpSI->nTrackPos) ||
						((si.fMask & SIF_DISABLENOSCROLL) != (lpSI->fMask & SIF_DISABLENOSCROLL));
				}
			}
		}

		lRes = CScrollBar::WindowProc
							(
								message == SBM_SETRANGEREDRAW ? SBM_SETRANGE : message, 
								message == SBM_SETSCROLLINFO ? 0 : wParam, 
								message == SBM_SETPOS ? 0 : lParam
							);

		if (bRedraw)
		{
			RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
		}
	}
	else
	{
		lRes = CScrollBar::WindowProc(message, wParam, lParam);

		if (message == SBM_ENABLE_ARROWS && m_dwDisabledArrows != (DWORD)wParam && lRes)
		{
			m_dwDisabledArrows = (DWORD)wParam;
			RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
		}

		if (m_HitPressed == CBCGPScrollBar::BCGP_HT_THUMB)
		{
			if (message == SBM_GETSCROLLINFO && lParam != NULL &&
				(((SCROLLINFO*)lParam)->fMask & SIF_TRACKPOS) == SIF_TRACKPOS && lRes)
			{
				((SCROLLINFO*)lParam)->nTrackPos = m_TrackPos;
			}
		}
	}

	return lRes;
}
//*****************************************************************************
BOOL CBCGPScrollBar::PreTranslateMessage (MSG* pMsg)
{
	if ((CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () || m_Style == BCGP_SBSTYLE_CUSTOM) &&
		m_Style != BCGP_SBSTYLE_DEFAULT)
	{
		if (pMsg->message == WM_MOUSEMOVE && m_HitPressed != BCGP_HT_NOWHERE)
		{
			OnMouseMove ((UINT) pMsg->wParam, CPoint ((DWORD)pMsg->lParam));
			return TRUE;
		}
	}

	return CScrollBar::PreTranslateMessage (pMsg);
}
//*****************************************************************************
void CBCGPScrollBar::OnTimer(UINT_PTR nIDEvent) 
{
	if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
		m_Style == BCGP_SBSTYLE_DEFAULT)
	{
		Default ();
		return;
	}

	if (!ProcessTimer(this, nIDEvent))
	{
		CScrollBar::OnTimer(nIDEvent);
	}
}
//*****************************************************************************
BOOL CBCGPScrollBar::OnEraseBkgnd(CDC* /*pDC*/) 
{
	if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
		m_Style == BCGP_SBSTYLE_DEFAULT)
	{
		return (BOOL) Default ();
	}

	return TRUE;
}
//*****************************************************************************
void CBCGPScrollBar::OnNcDestroy() 
{
	CScrollBar::OnNcDestroy();

	if (m_bAutoDestoyWindow)
	{
		delete this;
	}
}
//*****************************************************************************
LRESULT CBCGPScrollBar::OnBCGSetControlAero (WPARAM wp, LPARAM)
{
	m_bOnGlass = (BOOL) wp;
	return 0;
}
//*******************************************************************************
LRESULT CBCGPScrollBar::OnPrintClient(WPARAM wp, LPARAM lp)
{
	if (lp & PRF_CLIENT)
	{
		CDC* pDC = CDC::FromHandle((HDC) wp);
		ASSERT_VALID(pDC);

		if ((!CBCGPVisualManager::GetInstance ()->IsOwnerDrawScrollBar () && m_Style != BCGP_SBSTYLE_CUSTOM) ||
			m_Style == BCGP_SBSTYLE_DEFAULT)
		{
			return Default ();
		}
		
		DoPaint(pDC);
	}

	return 0;
}
//*******************************************************************************
CBCGPScrollBar::BCGPSB_HITTEST CBCGPScrollBar::HitTest (const CPoint& point) const
{
	return DoHitTest((CWnd*)this, point);
}
//*******************************************************************************
void CBCGPScrollBar::RebuildCustomColors()
{
	if (m_ColorTheme.m_clrButton == (COLORREF)-1)
	{
		if (CBCGPDrawManager::IsDarkColor(m_ColorTheme.m_clrFace))
		{
			m_ColorTheme.m_clrButton = CBCGPDrawManager::ColorMakeLighter(m_ColorTheme.m_clrFace, .3);
		}
		else
		{
			m_ColorTheme.m_clrButton = CBCGPDrawManager::ColorMakeDarker(m_ColorTheme.m_clrFace, .2);
		}
	}

	if (m_ColorTheme.m_bHighlightEffect)
	{
		if (CBCGPDrawManager::IsDarkColor(m_ColorTheme.m_clrButton))
		{
			m_ColorTheme.m_clrButtonHighlighted = CBCGPDrawManager::ColorMakeLighter(m_ColorTheme.m_clrButton);
		}
		else
		{
			m_ColorTheme.m_clrButtonHighlighted = CBCGPDrawManager::ColorMakeDarker(m_ColorTheme.m_clrButton);
		}

		if (m_ColorTheme.m_b3DEffect)
		{
			m_ColorTheme.m_clrButtonHighlightedGradient = CBCGPDrawManager::ColorMakeLighter(m_ColorTheme.m_clrButtonHighlighted, .5);
		}
	}

	if (m_ColorTheme.m_b3DEffect)
	{
		m_ColorTheme.m_clrButtonGradient = CBCGPDrawManager::ColorMakeLighter(m_ColorTheme.m_clrButton, .5);
		m_ColorTheme.m_clrButtonPressedGradient = CBCGPDrawManager::ColorMakeLighter(
			m_ColorTheme.m_clrButtonPressed != (COLORREF)-1 ? m_ColorTheme.m_clrButtonPressed : m_ColorTheme.m_clrButton, .5);
	}
}
//*******************************************************************************
COLORREF CBCGPScrollBar::DrawCustomBox(CDC* pDC, CRect rect, BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bDisabled)
{
	COLORREF clrFill = bDisabled ? m_ColorTheme.m_clrFace : bPressed ? (m_ColorTheme.m_clrButtonPressed != (COLORREF)-1 ? m_ColorTheme.m_clrButtonPressed : m_ColorTheme.m_clrButton) : (bHighlighted && m_ColorTheme.m_clrButtonHighlighted != (COLORREF)-1) ? m_ColorTheme.m_clrButtonHighlighted : m_ColorTheme.m_clrButton;
	COLORREF clrFillGradient = bDisabled ? (COLORREF)-1 : bPressed ? m_ColorTheme.m_clrButtonPressedGradient : (bHighlighted && m_ColorTheme.m_clrButtonHighlightedGradient != (COLORREF)-1) ? m_ColorTheme.m_clrButtonHighlightedGradient : m_ColorTheme.m_clrButtonGradient;
	COLORREF clrLine = bDisabled ? (COLORREF)-1 : (bPressed && m_ColorTheme.m_clrButtonBorderPressed != (COLORREF)-1) ? m_ColorTheme.m_clrButtonBorderPressed : m_ColorTheme.m_clrButtonBorder;

	CBCGPDrawManager dm (*pDC);
	
	if (clrFillGradient == (COLORREF)-1)
	{
		dm.DrawRect (rect, clrFill, clrLine);
	}
	else
	{
		if (bHorz)
		{
			dm.FillGradient(rect, clrFill, clrFillGradient);
		}
		else
		{
			dm.FillGradient(rect, clrFillGradient, clrFill, FALSE);
		}
		
		dm.DrawRect(rect, (COLORREF)-1, clrLine);
	}

	return clrFill;
}

BOOL CBCGPScrollBar::DoGesturePan(CPoint ptFrom, CPoint ptTo)
{
	ScreenToClient(&ptFrom);
	ScreenToClient(&ptTo);

	CRect rectClient;
	GetClientRect(rectClient);

	if (!rectClient.PtInRect(ptFrom))
	{
		return FALSE;
	}

	CBCGPRAII<BOOL> bHorz(m_bHorz, (GetStyle() & SBS_VERT) == 0);

	if (m_bHorz)
	{
		if (ptFrom.x == ptTo.x)
		{
			return TRUE;
		}
	}
	else
	{
		if (ptFrom.y == ptTo.y)
		{
			return TRUE;
		}
	}

	CRect rectSaved = m_rect;

	if (m_rect.IsRectEmpty())
	{
		m_rect = rectClient;
	}

	ProcessMouseMove(this, FALSE, ptFrom);
	ProcessMouseDown(this, ptFrom);
	ProcessMouseMove(this, TRUE, ptTo);
	ProcessMouseUp(this, ptTo);

	m_rect = rectSaved;
	return TRUE;
}


/////////////////////////////////////////////////////////////////////////////
// CBCGPInternalScrollBarWrapper

void CBCGPInternalScrollBarWrapper::_RedrawScrollBars(HRGN hRgn, int nBar)
{
	if (nBar == -1)
	{
		m_ScrollHorz.m_rect.SetRectEmpty();
		m_ScrollVert.m_rect.SetRectEmpty();
		m_rectSizeBox.SetRectEmpty();
	}
	else if (nBar == SB_HORZ)
	{
		m_ScrollHorz.m_rect.SetRectEmpty();
	}
	else if (nBar == SB_VERT)
	{
		m_ScrollVert.m_rect.SetRectEmpty();
	}

	CRect rectWindow;
	CRect rectClient;
	m_pWndOwner->GetWindowRect(rectWindow);
	m_pWndOwner->GetClientRect(rectClient);
	m_pWndOwner->ClientToScreen(rectClient);

	CPoint ptOffset(rectWindow.TopLeft());

	rectWindow = rectClient;
	rectWindow.OffsetRect(-ptOffset);

	if (m_pWndOwner->GetExStyle() & WS_EX_LAYOUTRTL)
	{
		rectWindow.OffsetRect(-GetSystemMetrics(SM_CXVSCROLL), 0);
	}

	CRect rectScrollBarH(0, 0, 0, 0);
	if ((m_dwScrool & WS_HSCROLL) == WS_HSCROLL && nBar != SB_VERT)
	{
		rectScrollBarH = rectWindow;
		rectScrollBarH.top = rectScrollBarH.bottom;
		rectScrollBarH.bottom += GetSystemMetrics(SM_CYHSCROLL);
	}

	CRect rectScrollBarV(0, 0, 0, 0);
	if ((m_dwScrool & WS_VSCROLL) == WS_VSCROLL && nBar != SB_HORZ)
	{
		rectScrollBarV = rectWindow;
		if (m_dwScroolEx & WS_EX_LEFTSCROLLBAR)
		{
			rectScrollBarV.right = rectScrollBarV.left;
			rectScrollBarV.left -= GetSystemMetrics(SM_CXVSCROLL);
		}
		else
		{
			rectScrollBarV.left = rectScrollBarV.right;
			rectScrollBarV.right += GetSystemMetrics(SM_CXVSCROLL);
		}
	}

	if ((m_dwScrool & WS_HSCROLL) && (m_dwScrool & WS_VSCROLL) && nBar == -1)
	{
		m_rectSizeBox = rectWindow;
		m_rectSizeBox.top = m_rectSizeBox.bottom;
		m_rectSizeBox.bottom += GetSystemMetrics(SM_CYHSCROLL);

		if (m_dwScroolEx & WS_EX_LEFTSCROLLBAR)
		{
			m_rectSizeBox.right = m_rectSizeBox.left;
			m_rectSizeBox.left -= GetSystemMetrics(SM_CXVSCROLL);
		}
		else
		{
			m_rectSizeBox.left = m_rectSizeBox.right;
			m_rectSizeBox.right += GetSystemMetrics(SM_CXVSCROLL);
		}
	}

	CRgn rgnScroll;

	if (!rectScrollBarH.IsRectEmpty())
	{
		rgnScroll.CreateRectRgnIndirect(rectScrollBarH);
	}

	if (!rectScrollBarV.IsRectEmpty())
	{
		if (rgnScroll.GetSafeHandle() == NULL)
		{
			rgnScroll.CreateRectRgnIndirect(rectScrollBarV);
		}
		else
		{
			CRgn rgnTemp;
			rgnTemp.CreateRectRgnIndirect(rectScrollBarV);
			rgnScroll.CombineRgn(&rgnScroll, &rgnTemp, RGN_OR);
		}
	}

	if (!m_rectSizeBox.IsRectEmpty())
	{
		if (rgnScroll.GetSafeHandle() == NULL)
		{
			rgnScroll.CreateRectRgnIndirect(m_rectSizeBox);
		}
		else
		{
			CRgn rgnTemp;
			rgnTemp.CreateRectRgnIndirect(m_rectSizeBox);
			rgnScroll.CombineRgn(&rgnScroll, &rgnTemp, RGN_OR);
		}
	}

	if (rgnScroll.GetSafeHandle() == NULL)
	{
		return;
	}

	rgnScroll.OffsetRgn(ptOffset);

	if (hRgn != NULL)
	{
		int nRes = rgnScroll.CombineRgn(&rgnScroll, CRgn::FromHandle(hRgn), RGN_AND);
		if (nRes == ERROR || nRes == NULLREGION)
		{
			return;
		}
	}

	CWindowDC dc(m_pWndOwner);

	rgnScroll.OffsetRgn(-ptOffset);

	dc.SelectClipRgn(&rgnScroll, RGN_COPY);

	if (!rectScrollBarH.IsRectEmpty())
	{
		SCROLLINFO si;
		si.cbSize = sizeof(SCROLLINFO);
		m_pWndOwner->GetScrollInfo(SB_HORZ, &si);

		m_ScrollHorz.SetRect(rectScrollBarH);

		CDC memDC;
		memDC.CreateCompatibleDC(&dc);
		CBitmap memBmp;
		memBmp.CreateCompatibleBitmap(&dc, rectScrollBarH.Width(), rectScrollBarH.Height());
		CBitmap* pOldBmp = (CBitmap*)memDC.SelectObject(&memBmp);

		m_ScrollHorz.Render(&memDC, m_pWndOwner->IsWindowEnabled(), si, -rectScrollBarH.TopLeft());

		dc.BitBlt(rectScrollBarH.left, rectScrollBarH.top, rectScrollBarH.Width(), rectScrollBarH.Height(), &memDC, 0, 0, SRCCOPY);

		memDC.SelectObject(pOldBmp);
	}

	if (!rectScrollBarV.IsRectEmpty())
	{
		SCROLLINFO si;
		si.cbSize = sizeof(SCROLLINFO);
		m_pWndOwner->GetScrollInfo(SB_VERT, &si);

		m_ScrollVert.SetRect(rectScrollBarV);

		CDC memDC;
		memDC.CreateCompatibleDC(&dc);
		CBitmap memBmp;
		memBmp.CreateCompatibleBitmap(&dc, rectScrollBarV.Width(), rectScrollBarV.Height());
		CBitmap* pOldBmp = (CBitmap*)memDC.SelectObject(&memBmp);

		m_ScrollVert.Render(&memDC, m_pWndOwner->IsWindowEnabled(), si, -rectScrollBarV.TopLeft());

		dc.BitBlt(rectScrollBarV.left, rectScrollBarV.top, rectScrollBarV.Width(), rectScrollBarV.Height(), &memDC, 0, 0, SRCCOPY);

		memDC.SelectObject(pOldBmp);
	}

	if (!m_rectSizeBox.IsRectEmpty())
	{
		dc.FillRect(m_rectSizeBox, &CBCGPVisualManager::GetInstance()->GetDlgBackBrush(NULL));
	}

	dc.SelectClipRgn(NULL, RGN_COPY);
}

LRESULT CBCGPInternalScrollBarWrapper::InternalSB_WrapperWindowProc(UINT message, WPARAM wParam, LPARAM lParam, BOOL& bProcessed)
{
	bProcessed = FALSE;

#ifdef _BCGSUITE_
	UNUSED_ALWAYS(message);
	UNUSED_ALWAYS(wParam);
	UNUSED_ALWAYS(lParam);

	return 0L;
#else
	if (globalData.m_nThemedScrollBars == BCGP_THEMED_SCROLLBAR_DISABLE || !IsInternalScrollBarThemed())
	{
		return 0L;
	}

	ASSERT_VALID(m_pWndOwner);

	if (message == BCGM_CHANGEVISUALMANAGER && !m_bSuppressMessage)
	{
		DWORD dwScroll = m_dwScrool;

		if (!CBCGPVisualManager::GetInstance()->IsOwnerDrawScrollBar())
		{
			m_dwScrool = 0;
			m_dwScroolEx = 0;

			m_ScrollHorz.m_rect.SetRectEmpty();
			m_ScrollVert.m_rect.SetRectEmpty();
			m_rectSizeBox.SetRectEmpty();

			for (int i = 0; i < 2; i++)
			{
				ZeroMemory(&m_ScrollInfo[i], sizeof(SCROLLINFO));
			}
		}
		else
		{
			m_dwScrool = m_pWndOwner->GetStyle() & (WS_HSCROLL | WS_VSCROLL);
			m_dwScroolEx = m_pWndOwner->GetExStyle() & (WS_EX_LEFTSCROLLBAR | WS_EX_LAYOUTRTL);
		}

		if (dwScroll != m_dwScrool)
		{
			m_pWndOwner->SetWindowPos(NULL, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOOWNERZORDER | SWP_FRAMECHANGED);
		}

		return 0L;
	}

	if (!CBCGPVisualManager::GetInstance()->IsOwnerDrawScrollBar())
	{
		return 0L;
	}

	if (m_pWndOwner->GetSafeHwnd() != NULL && ::IsWindow(m_pWndOwner->GetSafeHwnd()) && 
		(message == WM_LBUTTONDOWN || message == WM_LBUTTONDBLCLK || 
		 message == WM_RBUTTONDOWN || message == WM_RBUTTONDBLCLK || 
		 message == WM_MBUTTONDOWN || message == WM_MBUTTONDBLCLK))
	{
		CRect rectWindow;
		m_pWndOwner->GetWindowRect(rectWindow);

		CRect rectClient;
		m_pWndOwner->GetClientRect(rectClient);
		m_pWndOwner->ClientToScreen(rectClient);

		CPoint pt(BCG_GET_X_LPARAM(lParam), BCG_GET_Y_LPARAM(lParam));
		m_pWndOwner->ClientToScreen(&pt);

		if (rectWindow.PtInRect(pt) && !rectClient.PtInRect(pt))
		{
			LPARAM points;
			MAKEPOINTS(points).x = (SHORT)pt.x;
			MAKEPOINTS(points).y = (SHORT)pt.y;

			pt.x -= rectWindow.left;
			pt.y -= rectWindow.top;
			LRESULT hitTest = m_ScrollVert.DoHitTest(m_pWndOwner, pt) != CBCGPScrollBarImpl::BCGP_HT_NOWHERE
								? HTVSCROLL
								: m_ScrollHorz.DoHitTest(m_pWndOwner, pt) != CBCGPScrollBarImpl::BCGP_HT_NOWHERE
									? HTHSCROLL
									: SendMessage(m_pWndOwner->GetSafeHwnd(), WM_NCHITTEST, 0, points);

			if (hitTest == HTHSCROLL || hitTest == HTVSCROLL)
			{
				message = message - WM_LBUTTONDOWN + WM_NCLBUTTONDOWN;
				wParam = hitTest;
				lParam = points;
			}
		}
	}

	if (message == WM_CREATE)
	{
		LRESULT lResult = InternalSB_ControlWindowProc(message, wParam, lParam);

		if (lResult != -1)
		{
			m_dwScrool = m_pWndOwner->GetStyle() & (WS_HSCROLL | WS_VSCROLL);
			m_dwScroolEx = m_pWndOwner->GetExStyle() & (WS_EX_LEFTSCROLLBAR | WS_EX_LAYOUTRTL);
		}

		bProcessed = TRUE;
		return lResult;
	}
	else if (InternalSB_OnMessageRedraw(message, wParam, lParam) && m_dwScrool != 0)
	{
		HWND hwndOwner = m_pWndOwner->GetSafeHwnd();

		LRESULT lResult = InternalSB_ControlWindowProc(message, wParam, lParam);

		if (!::IsWindow(hwndOwner))
		{
			bProcessed = TRUE;
			return lResult;
		}

		SCROLLINFO si;

		BOOL bRepaint = FALSE;
		for (int i = 0; i < 2; i++)
		{
			ZeroMemory(&si, sizeof(SCROLLINFO));
			si.cbSize = sizeof(SCROLLINFO);

			m_pWndOwner->GetScrollInfo(i, &si);

			if (!_CompareScrollInfo(m_ScrollInfo[i], si))
			{
				m_ScrollInfo[i] = si;
				bRepaint = TRUE;
				break;
			}
		}

		if (bRepaint)
		{
			_RedrawScrollBars(NULL);
		}

		bProcessed = TRUE;
		return lResult;
	}
	else if (message == BCGM_REDRAW_SCROLLBAR)
	{
		_RedrawScrollBars(NULL, (int)wParam);
	}
	else if (message == WM_PRINT && !m_bSuppressMessage && m_dwScrool != 0)
	{
		DWORD dwFlags = (DWORD)lParam;
		LRESULT lResult = InternalSB_ControlWindowProc(message, wParam, lParam);
		bProcessed = TRUE;

		if ((dwFlags & PRF_NONCLIENT) == PRF_NONCLIENT)
		{
			CDC* pDC = CDC::FromHandle((HDC)wParam);
			ASSERT_VALID(pDC);
			
			if (!m_rectSizeBox.IsRectEmpty())
			{
				pDC->FillRect(m_rectSizeBox, &CBCGPVisualManager::GetInstance()->GetDlgBackBrush(NULL));
			}

			SCROLLINFO si;
			si.cbSize = sizeof(SCROLLINFO);

			m_pWndOwner->GetScrollInfo(SB_HORZ, &si);
			m_ScrollHorz.Render(pDC, m_pWndOwner->IsWindowEnabled(), si);

			m_pWndOwner->GetScrollInfo(SB_VERT, &si);
			m_ScrollVert.Render(pDC, m_pWndOwner->IsWindowEnabled(), si);
		}

		return lResult;
	}
	else if (message == WM_NCPAINT && !m_bSuppressMessage && m_dwScrool != 0)
	{
		m_bSuppressMessage = TRUE;

		DWORD dwStyle = m_pWndOwner->GetStyle();

		SetWindowLongPtr(m_pWndOwner->GetSafeHwnd(), GWL_STYLE, (dwStyle & ~m_dwScrool) | _GetExcludeStyle(m_pWndOwner));

		HWND hwndOwner = m_pWndOwner->GetSafeHwnd();

		LRESULT lResult = InternalSB_ControlWindowProc(message, wParam, lParam);

		if (!::IsWindow(hwndOwner))
		{
			bProcessed = TRUE;
			return lResult;
		}

		SetWindowLongPtr(m_pWndOwner->GetSafeHwnd(), GWL_STYLE, dwStyle);

		m_bSuppressMessage = FALSE;

		_RedrawScrollBars(NULL);

		bProcessed = TRUE;
		return lResult;
	}
	else if (message == WM_NCLBUTTONDOWN && (wParam == HTHSCROLL || wParam == HTVSCROLL))
	{
		CPoint point(BCG_GET_X_LPARAM(lParam), BCG_GET_Y_LPARAM(lParam));

		ScreenToBar(m_pWndOwner, point);

		if (wParam == HTHSCROLL)
		{
			bProcessed = m_ScrollHorz.ProcessMouseDown(m_pWndOwner, point);
		}
		else if (wParam == HTVSCROLL)
		{
			bProcessed = m_ScrollVert.ProcessMouseDown(m_pWndOwner, point);
		}

		return 0L;
	}
	else if (message == WM_NCLBUTTONUP && (wParam == HTHSCROLL || wParam == HTVSCROLL))
	{
		CPoint point(BCG_GET_X_LPARAM(lParam), BCG_GET_Y_LPARAM(lParam));

		ScreenToBar(m_pWndOwner, point);

		if (wParam == HTHSCROLL)
		{
			bProcessed = m_ScrollHorz.ProcessMouseUp(m_pWndOwner, point);
		}
		else if (wParam == HTVSCROLL)
		{
			bProcessed = m_ScrollVert.ProcessMouseUp(m_pWndOwner, point);
		}

		return 0L;
	}
	else if (message == WM_NCLBUTTONDBLCLK && (wParam == HTHSCROLL || wParam == HTVSCROLL))
	{
		CPoint point(BCG_GET_X_LPARAM(lParam), BCG_GET_Y_LPARAM(lParam));

		ScreenToBar(m_pWndOwner, point);

		if (wParam == HTHSCROLL)
		{
			m_ScrollHorz.ProcessMouseDblClk(m_pWndOwner, point);
		}
		else if (wParam == HTVSCROLL)
		{
			m_ScrollVert.ProcessMouseDblClk(m_pWndOwner, point);
		}

		bProcessed = TRUE;
		return 0L;
	}
	else if (message == WM_LBUTTONUP)
	{
		CPoint point(BCG_GET_X_LPARAM(lParam), BCG_GET_Y_LPARAM(lParam));

		SCROLLINFO si;
		ZeroMemory(&si, sizeof(SCROLLINFO));
		si.cbSize = sizeof(SCROLLINFO);

		if (m_ScrollVert.m_HitPressed != CBCGPScrollBarImpl::BCGP_HT_NOWHERE)
		{
			m_ScrollVert.ProcessMouseUp(m_pWndOwner, point);

			bProcessed = TRUE;
		}
		else if (m_ScrollHorz.m_HitPressed != CBCGPScrollBarImpl::BCGP_HT_NOWHERE)
		{
			m_ScrollHorz.ProcessMouseUp(m_pWndOwner, point);

			bProcessed = TRUE;
		}
		
		return 0L;
	}
	else if (message == WM_NCMOUSEMOVE && (wParam == HTHSCROLL || wParam == HTVSCROLL))
	{
		CPoint point(BCG_GET_X_LPARAM(lParam), BCG_GET_Y_LPARAM(lParam));

		ScreenToBar(m_pWndOwner, point);

		SCROLLINFO si;
		ZeroMemory(&si, sizeof(SCROLLINFO));
		si.cbSize = sizeof(SCROLLINFO);

		if (wParam == HTHSCROLL)
		{
			m_ScrollHorz.ProcessMouseMove(m_pWndOwner, m_ScrollHorz.m_HitPressed != CBCGPScrollBar::BCGP_HT_NOWHERE, point);
		}
		else if (wParam == HTVSCROLL)
		{
			m_ScrollVert.ProcessMouseMove(m_pWndOwner, m_ScrollVert.m_HitPressed != CBCGPScrollBar::BCGP_HT_NOWHERE, point);
		}

		bProcessed = TRUE;
		return 0L;
	}
	else if (message == WM_NCHITTEST && !m_bSuppressMessage)
	{
		HWND hwndOwner = m_pWndOwner->GetSafeHwnd();

		LRESULT lResult = InternalSB_ControlWindowProc(message, wParam, lParam);

		if (!::IsWindow(hwndOwner))
		{
			bProcessed = TRUE;
			return lResult;
		}

		if (lResult != HTNOWHERE)
		{
			bProcessed = TRUE;
			return lResult;
		}

		CPoint pt(BCG_GET_X_LPARAM(lParam), BCG_GET_Y_LPARAM(lParam));

		CRect rectClient;
		m_pWndOwner->GetClientRect(rectClient);
		m_pWndOwner->ClientToScreen(rectClient);

		CRect rectWindow;
		m_pWndOwner->GetWindowRect(rectWindow);

		if ((m_dwScrool & WS_VSCROLL) &&
			(((m_dwScroolEx & (WS_EX_LEFTSCROLLBAR | WS_EX_LAYOUTRTL)) && rectWindow.left <= pt.x && pt.x < rectClient.left) ||
				(rectClient.right < pt.x && pt.x <= rectWindow.right)))
		{
			lResult = HTVSCROLL;
		}

		if ((m_dwScrool & WS_HSCROLL) && rectClient.bottom < pt.y && pt.y <= rectWindow.bottom)
		{
			lResult = HTHSCROLL;
		}

		bProcessed = TRUE;
		return lResult;
	}
	else if (message == WM_NCCALCSIZE)
	{
		DWORD dwStyle = m_pWndOwner->GetStyle();

		if (!m_bSuppressMessage)
		{
			m_dwScrool = dwStyle & (WS_HSCROLL | WS_VSCROLL);
			m_dwScroolEx = m_pWndOwner->GetExStyle() & (WS_EX_LEFTSCROLLBAR | WS_EX_LAYOUTRTL);
		}

		if (m_dwScrool != 0)
		{
			m_bSuppressMessage = TRUE;

			SetWindowLongPtr(m_pWndOwner->GetSafeHwnd(), GWL_STYLE, (dwStyle & ~m_dwScrool) | _GetExcludeStyle(m_pWndOwner));

			HWND hwndOwner = m_pWndOwner->GetSafeHwnd();

			LRESULT lResult = InternalSB_ControlWindowProc(message, wParam, lParam);

			if (!::IsWindow(hwndOwner))
			{
				bProcessed = TRUE;
				return lResult;
			}

			SetWindowLongPtr(m_pWndOwner->GetSafeHwnd(), GWL_STYLE, dwStyle);

			m_bSuppressMessage = FALSE;

			LPNCCALCSIZE_PARAMS nccsp = (LPNCCALCSIZE_PARAMS)lParam;
			RECT* rect = &nccsp->rgrc[0];

			if (m_dwScrool & WS_VSCROLL)
			{
				if (((m_dwScroolEx & WS_EX_LEFTSCROLLBAR) != 0 && (m_dwScroolEx & WS_EX_LAYOUTRTL) == 0) ||
					((m_dwScroolEx & WS_EX_LEFTSCROLLBAR) == 0 && (m_dwScroolEx & WS_EX_LAYOUTRTL) != 0))
				{
					rect->left += GetSystemMetrics(SM_CXVSCROLL);
				}
				else
				{
					rect->right -= GetSystemMetrics(SM_CXVSCROLL);
				}
			}

			if (m_dwScrool & WS_HSCROLL)
			{
				rect->bottom -= GetSystemMetrics(SM_CYHSCROLL);
			}

			bProcessed = TRUE;
			return lResult;
		}
	}
	else if (message == WM_SIZE)
	{
		DWORD dwScroll = m_dwScrool;

		LRESULT lResult = InternalSB_ControlWindowProc(message, wParam, lParam);

		m_dwScrool = m_pWndOwner->GetStyle() & (WS_HSCROLL | WS_VSCROLL);
		if (dwScroll != m_dwScrool)
		{
			m_pWndOwner->PostMessage(BCGM_REDRAW_SCROLLBAR, (WPARAM)-1);
		}

		bProcessed = TRUE;
		return lResult;
	}
	else if ((message == WM_STYLECHANGING || message == WM_STYLECHANGED) && !m_bSuppressMessage)
	{
		LPSTYLESTRUCT lpStruct = (LPSTYLESTRUCT)lParam;

		if (wParam == GWL_STYLE)
		{
			m_dwScrool = lpStruct->styleNew & (WS_HSCROLL | WS_VSCROLL);
		}
		else if (wParam == GWL_EXSTYLE)
		{
			m_dwScroolEx = lpStruct->styleNew & (WS_EX_LEFTSCROLLBAR | WS_EX_LAYOUTRTL);
		}
	}
	else if (message == WM_TIMER && 
			(m_ScrollHorz.m_TimerDown == wParam || m_ScrollHorz.m_Timer == wParam ||
			m_ScrollVert.m_TimerDown == wParam || m_ScrollVert.m_Timer == wParam))
	{
		CPoint point;
		::GetCursorPos(&point);

		ScreenToBar(m_pWndOwner, point);

		BOOL bRedraw = FALSE;

		CBCGPScrollBarImpl* pImpl = NULL;
		if (m_ScrollHorz.TimerStarted())
		{
			pImpl = &m_ScrollHorz;
		}
		else if (m_ScrollVert.TimerStarted())
		{
			pImpl = &m_ScrollVert;
		}

		if (pImpl->m_Timer != 0)
		{
			CBCGPScrollBarImpl::BCGPSB_HITTEST hit = pImpl->DoHitTest(m_pWndOwner, point);

			if (hit != pImpl->m_HitTest)
			{
				pImpl->m_HitTest = hit;
				bRedraw = TRUE;
			}
		}

		pImpl->ProcessTimer(m_pWndOwner, wParam);

		if (bRedraw)
		{
			_RedrawScrollBars(NULL, pImpl->m_bHorz ? SB_HORZ : SB_VERT);
		}

		bProcessed = TRUE;
	}
	else if (message == WM_CANCELMODE)
	{
		CPoint point;
		::GetCursorPos (&point);

		m_ScrollVert.ProcessCancelMode(m_pWndOwner, point);
		m_ScrollHorz.ProcessCancelMode(m_pWndOwner, point);
	}
	else if (message == WM_MOUSELEAVE)
	{
		m_ScrollVert.ProcessMouseLeave(m_pWndOwner);
		m_ScrollHorz.ProcessMouseLeave(m_pWndOwner);
	}
	else if (message == WM_MOUSEMOVE)
	{
		CPoint point(BCG_GET_X_LPARAM(lParam), BCG_GET_Y_LPARAM(lParam));
		m_pWndOwner->ClientToScreen(&point);

		ScreenToBar(m_pWndOwner, point);

		BOOL bVertIsPressed = m_ScrollVert.m_HitPressed != CBCGPScrollBar::BCGP_HT_NOWHERE;
		BOOL bHorzIsPressed = m_ScrollHorz.m_HitPressed != CBCGPScrollBar::BCGP_HT_NOWHERE;

		if (m_ScrollVert.m_rect.PtInRect(point) || bVertIsPressed)
		{
			m_ScrollVert.ProcessMouseMove(m_pWndOwner, bVertIsPressed, point);
		}
		else
		{
			m_ScrollVert.ProcessMouseLeave(m_pWndOwner);
		}

		if (m_ScrollHorz.m_rect.PtInRect(point) || bHorzIsPressed)
		{
			m_ScrollHorz.ProcessMouseMove(m_pWndOwner, bHorzIsPressed, point);
		}
		else
		{
			m_ScrollHorz.ProcessMouseLeave(m_pWndOwner);
		}
	}
	else if ((message == WM_HSCROLL || message == WM_VSCROLL) && lParam == NULL)
	{
		if (LOWORD(wParam) == SB_THUMBTRACK)
		{
			bProcessed = InternalSB_OnThumbTrack(message == WM_HSCROLL ? SB_HORZ : SB_VERT, message == WM_HSCROLL ? m_ScrollHorz.m_TrackPos : m_ScrollVert.m_TrackPos);
		}
	}
	else if (message == WM_RBUTTONDOWN || message == WM_RBUTTONUP)
	{
		if (m_ScrollVert.m_bTracked || m_ScrollVert.m_HitPressed != CBCGPScrollBarImpl::BCGP_HT_NOWHERE || 
			m_ScrollHorz.m_bTracked || m_ScrollHorz.m_HitPressed != CBCGPScrollBarImpl::BCGP_HT_NOWHERE)
		{
			bProcessed = TRUE;
		}
	}

	return 0L;
#endif
}
