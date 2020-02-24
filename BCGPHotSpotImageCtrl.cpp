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
// BCGPHotSpotImageCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "bcgprores.h"
#include "bcgglobals.h"
#include "BCGPLocalResource.h"
#include "BCGPHotSpotImageCtrl.h"
#include "BCGPDrawManager.h"
#include "BCGPGlobalUtils.h"

#ifndef _BCGSUITE_
	#include "BCGPPopupMenu.h"
	#include "BCGPPngImage.h"
	#include "BCGPTooltipManager.h"
#endif

#ifndef BCGP_EXCLUDE_HOT_SPOT_IMAGE

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define ID_SCROLL_VERT	2
#define ID_SCROLL_HORZ	3

/////////////////////////////////////////////////////////////////////////////
// CBCGPHotSpot

IMPLEMENT_SERIAL(CBCGPHotSpot, CObject, 1)

CBCGPHotSpot::CBCGPHotSpot(UINT nID) :
	m_nID (nID)
{
	m_rect.SetRectEmpty ();
}
//*******************************************************************************
CBCGPHotSpot::~CBCGPHotSpot()
{
}
//*******************************************************************************
void CBCGPHotSpot::CreateRect (CRect rect)
{
	m_rect = rect;
}
//*******************************************************************************
BOOL CBCGPHotSpot::IsPointWithin (CPoint pt) const
{
	return m_rect.PtInRect (pt);
}
//********************************************************************************
void CBCGPHotSpot::Serialize(CArchive& ar)
{
	CObject::Serialize (ar);

	if (ar.IsLoading ())
	{
		ar >> (ULONG&) m_nID;
		ar >> m_rect;
	}
	else
	{
		ar << m_nID;
		ar << m_rect;
	}
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPHotSpotImageCtrl

CBCGPHotSpotImageCtrl::CBCGPHotSpotImageCtrl(BOOL bDPIAutoScale)
{
#if (!defined _BCGSUITE_) || (_MSC_VER >= 1600)
	m_bDPIAutoScale		= bDPIAutoScale;
#else
	UNREFERENCED_PARAMETER(bDPIAutoScale);
	m_bDPIAutoScale		= FALSE;
#endif
	m_sizeImage			= CSize (0, 0);
	m_clrBack			= RGB (255, 255, 255);

	m_pHot				= NULL;
	m_pClicked			= NULL;

	m_nSelectedNum		= 0;

	m_lineDev			= CSize (0, 0);
	m_pageDev			= CSize (0, 0);

	m_nHighlightPercentage = 0;
	m_nHighlightTolerance = 0;

	m_clrTransparentCurr = (COLORREF)-1;

	m_bScrollBars		= TRUE;
	m_ScrollBarStyle	= CBCGPScrollBar::BCGP_SBSTYLE_DEFAULT;

	m_rectView.SetRectEmpty ();

	m_pToolTip = NULL;
	m_bRebuildTooltips = FALSE;

	m_bKeepHotSpotsOnCleanup = FALSE;
}
//*******************************************************************************
CBCGPHotSpotImageCtrl::~CBCGPHotSpotImageCtrl()
{
	CleanUp ();
}

BEGIN_MESSAGE_MAP(CBCGPHotSpotImageCtrl, CButton)
	//{{AFX_MSG_MAP(CBCGPHotSpotImageCtrl)
	ON_WM_ERASEBKGND()
	ON_WM_SETCURSOR()
	ON_WM_MOUSEMOVE()
	ON_WM_MOUSEWHEEL()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_CANCELMODE()
	ON_WM_SIZE()
	ON_WM_VSCROLL()
	ON_WM_HSCROLL()
	ON_WM_CREATE()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_STYLECHANGING, OnStyleChanging)
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)
	ON_REGISTERED_MESSAGE(BCGM_UPDATETOOLTIPS, OnBCGUpdateToolTips)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnTTNeedTipText)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPHotSpotImageCtrl message handlers

BOOL CBCGPHotSpotImageCtrl::Create(const RECT& rect, CWnd* pParentWnd, UINT nID)
{
	return CButton::Create(_T(""), WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN, rect, pParentWnd, nID);
}
//*******************************************************************************
BOOL CBCGPHotSpotImageCtrl::PreCreateWindow(CREATESTRUCT& cs) 
{
	cs.style |= (BS_OWNERDRAW);
	cs.style &= ~BS_DEFPUSHBUTTON;
	cs.dwExStyle &= ~WS_EX_LAYOUTRTL;

	return CButton::PreCreateWindow(cs);
}
//********************************************************************************
void CBCGPHotSpotImageCtrl::PreSubclassWindow() 
{
	ModifyStyle (BS_DEFPUSHBUTTON, BS_OWNERDRAW);
	ModifyStyleEx (WS_EX_LAYOUTRTL, 0);

	CButton::PreSubclassWindow();

	UpdateScrollBars ();
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::DrawItem(LPDRAWITEMSTRUCT lpDIS) 
{
	ASSERT (lpDIS != NULL);
	ASSERT (lpDIS->CtlType == ODT_BUTTON);

	if (m_bRebuildTooltips)
	{
		RebuildToolTips();
	}

	CDC* pDCPaint = CDC::FromHandle (lpDIS->hDC);
	ASSERT_VALID (pDCPaint);

	OnPrepareDC (pDCPaint);

	CRect rect = lpDIS->rcItem;

	CPoint ptScrollOffset (GetScrollPos (SB_HORZ), GetScrollPos (SB_VERT));

	rect.right += ptScrollOffset.x;
	rect.bottom += ptScrollOffset.y;

	CRect rectClip;
	int nClipType = pDCPaint->GetClipBox (rectClip);
	if (nClipType != SIMPLEREGION)
	{
		rectClip = rect;
	}

	CBCGPMemDC memDC (*pDCPaint, this);
	CDC* pDC = &memDC.GetDC ();

	if (!m_rectView.IsRectEmpty ())
	{
		// Draw gripper
		CRect rectGripper;
		GetClientRect (rectGripper);

		if (globalData.m_bIsRTL)
		{
			rectGripper.right = m_rectView.left;
		}
		else
		{
			rectGripper.left = m_rectView.right;
		}
		rectGripper.top = m_rectView.bottom;
		rectGripper.OffsetRect (ptScrollOffset);

		pDC->FillRect (rectGripper, &globalData.brBarFace);
	}

	CRgn rgnClip;
	if (!m_rectView.IsRectEmpty ())
	{
		// Clip scrollbars
		CRect rectViewClip = m_rectView;
		rectViewClip.OffsetRect (ptScrollOffset);

 		rect.right = min(rectViewClip.right, rect.right);
 		rect.bottom = min(rectViewClip.bottom, rect.bottom);

		rgnClip.CreateRectRgnIndirect (&rectViewClip);
		pDC->SelectClipRgn (&rgnClip);
	}

	pDC->FillSolidRect (rectClip, m_clrBack);

	OnDrawImage (pDC, rect);

	CBCGPHotSpot* pHot = m_pHot;
	if (m_pClicked != NULL)
	{
		pHot = m_pClicked;
	}

	CRgn rgnHot;
	if (pHot != NULL)
	{
		CRect rectHotSpot = pHot->m_rect;
		m_clrTransparentCurr = pDC->GetPixel (rectHotSpot.TopLeft ());

		pDC->LPtoDP (&rectHotSpot);

		rgnHot.CreateRectRgnIndirect (rectHotSpot);
		pDC->SelectClipRgn (&rgnHot);

		OnDrawHotSpot (pDC, rect, pHot);
	}

	if (rgnClip.GetSafeHandle() != NULL || rgnHot.GetSafeHandle() != NULL)
	{
		pDC->SelectClipRgn (NULL);
	}
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::OnPrepareDC(CDC* pDC)
{
	ASSERT_VALID(pDC);

	pDC->SetMapMode(MM_TEXT);

	CPoint ptVpOrg(0, 0);       // assume no shift for printing

	// by default shift viewport origin in negative direction of scroll
	ptVpOrg = -GetDeviceScrollPosition();

	pDC->SetViewportOrg(ptVpOrg);
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::OnDrawImage (CDC* pDC, CRect rectImage)
{
	ASSERT_VALID (pDC);

	if (m_Image.GetCount() == 0)
	{
		return;
	}

	rectImage.right = min (rectImage.right, m_sizeImage.cx);
	rectImage.bottom = min (rectImage.bottom, m_sizeImage.cy);

	m_Image.DrawEx(pDC, rectImage, 0);
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::OnDrawHotSpot (CDC* pDC, CRect rectImage, CBCGPHotSpot* /*pHotSpot*/)
{
	ASSERT_VALID (pDC);

	rectImage.right = min (rectImage.right, m_sizeImage.cx);
	rectImage.bottom = min (rectImage.bottom, m_sizeImage.cy);

	if (m_ImageHot.GetCount() > 0)
	{
		m_ImageHot.DrawEx(pDC, rectImage, 0);
	}

	if (m_nHighlightPercentage > 0)
	{
		CBCGPDrawManager dm (*pDC);
		dm.HighlightRect (rectImage, m_nHighlightPercentage,
			m_clrTransparentCurr, m_nHighlightTolerance);
	}
}
//********************************************************************************
BOOL CBCGPHotSpotImageCtrl::SetImage (UINT uiBmpResId, UINT uiBmpHotResId/* = 0*/, BOOL bKeepHotSpots/* = FALSE*/)
{
	if (uiBmpResId == 0)
	{
		CBCGPRAII<BOOL> raii(m_bKeepHotSpotsOnCleanup, bKeepHotSpots);
		CleanUp ();
		return TRUE;
	}

	BOOL bIsBlackHighContrast = globalData.m_bIsBlackHighContrast;
	if (bIsBlackHighContrast)
	{
		globalData.m_bIsBlackHighContrast = FALSE;
	}

	CBCGPToolBarImages image;
	CBCGPToolBarImages imageHot;

	image.SetMapTo3DColors(FALSE);
	image.Load(uiBmpResId);

	if (uiBmpHotResId != 0)
	{
		imageHot.SetMapTo3DColors(FALSE);
		imageHot.Load(uiBmpHotResId);
	}

	if (bIsBlackHighContrast)
	{
		globalData.m_bIsBlackHighContrast = TRUE;
	}

	return SetImage (image.GetImageWell(), imageHot.GetImageWell(), FALSE, bKeepHotSpots);
}
//*******************************************************************************
BOOL CBCGPHotSpotImageCtrl::SetImage (HBITMAP hbmpImage, HBITMAP hbmpImageHot/* = NULL*/, BOOL bAutoDestroy/* = FALSE*/, BOOL bKeepHotSpots/* = FALSE*/)
{
	{
		CBCGPRAII<BOOL> raii(m_bKeepHotSpotsOnCleanup, bKeepHotSpots);
		CleanUp ();
	}

	ASSERT (hbmpImage != NULL);

	const BOOL bIsRTL = CBCGPToolBarImages::IsRTL();
	CBCGPToolBarImages::EnableRTL(FALSE);

	m_Image.Clear();
	m_ImageHot.Clear();

	m_Image.AddImage(hbmpImage, TRUE);
	m_Image.SetSingleImage();
	
	if (hbmpImageHot != NULL)
	{
		m_ImageHot.AddImage(hbmpImageHot, TRUE);
		m_ImageHot.SetSingleImage();
	}

	if (m_bDPIAutoScale && globalData.GetRibbonImageScale() > 1.0)
	{
		globalUtils.ScaleByDPI(m_Image);

		if (m_ImageHot.GetCount() > 0)
		{
			globalUtils.ScaleByDPI(m_ImageHot);
		}
	}

	m_sizeImage = m_Image.GetImageSize();

	m_pageDev.cx = m_sizeImage.cx / 10;
	m_pageDev.cy = m_sizeImage.cy / 10;
	m_lineDev.cx = m_pageDev.cx / 10;
	m_lineDev.cy = m_pageDev.cy / 10;

	if (m_ImageHot.GetCount() > 0)
	{
		// Hot and cold bitmaps should have the same size!
		ASSERT (m_sizeImage.cx == m_ImageHot.GetImageSize().cx);
		ASSERT (m_sizeImage.cy == m_ImageHot.GetImageSize().cy);
	}

	if (bAutoDestroy)
	{
		::DeleteObject(hbmpImage);
		if (hbmpImageHot != NULL)
		{
			::DeleteObject(hbmpImageHot);
		}
	}

	if (GetSafeHwnd () != NULL)
	{
		UpdateScrollBars ();
		RedrawWindow ();
	}

	CBCGPToolBarImages::EnableRTL(bIsRTL);
	return TRUE;
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::CleanUp ()
{
	CBCGPTooltipManager::DeleteToolTip(m_pToolTip);

	if (!m_bKeepHotSpotsOnCleanup)
	{
		while (!m_lstHotSpots.IsEmpty ())
		{
			delete m_lstHotSpots.RemoveHead ();
		}
	}

	m_sizeImage = CSize (0, 0);

	m_lineDev = CSize (0, 0);
	m_pageDev = CSize (0, 0);

	UpdateScrollBars ();
}
//********************************************************************************
BOOL CBCGPHotSpotImageCtrl::OnEraseBkgnd(CDC* /*pDC*/) 
{
	return TRUE;
}
//********************************************************************************
void CBCGPHotSpotImageCtrl::SetBackgroundColor (COLORREF color)
{
	m_clrBack = color;

	if (GetSafeHwnd () != NULL)
	{
		RedrawWindow ();
	}
}
//********************************************************************************
CBCGPHotSpot* CBCGPHotSpotImageCtrl::FindHotSpot (UINT nID) const
{
	for (POSITION pos = m_lstHotSpots.GetHeadPosition (); pos != NULL;)
	{
		CBCGPHotSpot* pHotSpot = m_lstHotSpots.GetNext (pos);
		ASSERT_VALID (pHotSpot);

		if (pHotSpot->m_nID == nID)
		{
			return pHotSpot;
		}
	}

	return NULL;
}
//*******************************************************************************
BOOL CBCGPHotSpotImageCtrl::AddHotSpot (UINT nID, CRect rect, LPCTSTR lpszToolTip, LPCTSTR lpszToolTipDescription)
{
	if (FindHotSpot (nID) != NULL)
	{
		// nID should be unique!
		ASSERT (FALSE);
		return FALSE;
	}

	ASSERT (nID != 0);

	CBCGPHotSpot* pHotSpot = new CBCGPHotSpot (nID);

	if (m_bDPIAutoScale && globalData.GetRibbonImageScale() > 1.0)
	{
		CPoint pt(globalUtils.ScaleByDPI(rect.left), globalUtils.ScaleByDPI(rect.top));
		CSize size = globalUtils.ScaleByDPI(rect.Size());

		pHotSpot->CreateRect(CRect(pt, size));
	}
	else
	{
		pHotSpot->CreateRect (rect);
	}

	if (lpszToolTip != NULL)
	{
		pHotSpot->m_strToolTip = lpszToolTip;
		pHotSpot->m_strToolTipDescription = (lpszToolTipDescription == NULL) ? _T("") : lpszToolTipDescription;

		m_bRebuildTooltips = TRUE;
	}

	m_lstHotSpots.AddTail (pHotSpot);

	return TRUE;
}
//********************************************************************************
void CBCGPHotSpotImageCtrl::SetHotSpotToolTip(UINT nID, LPCTSTR lpszToolTip, LPCTSTR lpszToolTipDescription)
{
	CBCGPHotSpot* pHotSpot = FindHotSpot(nID);
	if (pHotSpot != NULL)
	{
		ASSERT_VALID(pHotSpot);

		pHotSpot->m_strToolTip = (lpszToolTip == NULL) ? _T("") : lpszToolTip;
		pHotSpot->m_strToolTipDescription = (lpszToolTipDescription == NULL) ? _T("") : lpszToolTipDescription;

		m_bRebuildTooltips = TRUE;
	}
}
//********************************************************************************
BOOL CBCGPHotSpotImageCtrl::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	CPoint point;
	GetCursorPos (&point);
	ScreenToClient (&point);

	if (HitTest (point) != NULL)
	{
		::SetCursor (globalData.GetHandCursor ());
		return TRUE;
	}
	
	return CButton::OnSetCursor(pWnd, nHitTest, message);
}
//*********************************************************************************
CBCGPHotSpot* CBCGPHotSpotImageCtrl::HitTest (CPoint pt)
{
	CRect rectClient;
	GetClientRect (rectClient);

	if (!rectClient.PtInRect (pt))
	{
		return NULL;
	}

	CClientDC dc (this);
	OnPrepareDC (&dc);

	CPoint ptScreen = pt;
	ClientToScreen (&ptScreen);

	dc.DPtoLP (&pt);

	for (POSITION pos = m_lstHotSpots.GetHeadPosition (); pos != NULL;)
	{
		CBCGPHotSpot* pHotSpot = m_lstHotSpots.GetNext (pos);
		ASSERT_VALID (pHotSpot);

		if (pHotSpot->IsPointWithin (pt))
		{
			if (::WindowFromPoint (ptScreen) != GetSafeHwnd ())
			{
				return NULL;
			}

			return pHotSpot;
		}
	}

	return NULL;
}
//********************************************************************************
void CBCGPHotSpotImageCtrl::RedrawRect (CRect rect)
{
	CClientDC dc (this);
	OnPrepareDC (&dc);

	dc.LPtoDP (&rect);

	RedrawWindow (rect, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::OnMouseMove(UINT /*nFlags*/, CPoint point) 
{
	CBCGPHotSpot* pHot = HitTest (point);

	if (m_pHot == pHot)
	{
		return;
	}

	// ----------------
	// No new hot areas
	// ----------------
	if (pHot == NULL)
	{
		if (m_pHot != NULL)
		{
			// remove old hotlight
			CRect rectUpdate = m_pHot->m_rect;
			m_pHot = NULL;

			if (HasHotImage () || m_nHighlightPercentage > 0)
			{
				RedrawRect (rectUpdate);
			}

			if (m_pClicked == NULL)
			{
				ReleaseCapture ();
			}
		}
		
		return;
	}
	else if (pHot != NULL)
	{
		// ------------
		// New hot spot
		// ------------
		if (m_pHot == NULL)
		{
			if (GetCapture () != NULL)
			{
				return;
			}

			SetCapture ();
		}
		else
		{
			// remove old hotlight
			CRect rectHotSpot = m_pHot->m_rect;
			m_pHot = NULL;

			if (HasHotImage () || m_nHighlightPercentage > 0)
			{
				RedrawRect (rectHotSpot);
			}
		}
		
		// add new hotlight
		m_pHot = pHot;

		if (HasHotImage () || m_nHighlightPercentage > 0)
		{
			RedrawRect (pHot->m_rect);
		}
	}
}
//*******************************************************************************
BOOL CBCGPHotSpotImageCtrl::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	if (CBCGPPopupMenu::GetActiveMenu () != NULL)
	{
		return TRUE;
	}

	if (nFlags & (MK_SHIFT | MK_CONTROL))
	{
		return FALSE;
	}
	
	UNUSED_ALWAYS(pt);

	DWORD dwStyle = GetStyle();
	CScrollBar* pBar = GetScrollBarCtrl(SB_VERT);
	BOOL bHasVertBar = ((pBar != NULL) && pBar->IsWindowEnabled()) || (dwStyle & WS_VSCROLL);

	SCROLLINFO scrollInfo;
	ZeroMemory(&scrollInfo, sizeof(SCROLLINFO));

    scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_ALL;

	GetScrollInfo (bHasVertBar ? SB_VERT : SB_HORZ, &scrollInfo);

	if (scrollInfo.nPos >= scrollInfo.nMax)
	{
		return FALSE;
	}

	int nSteps = abs(zDelta) / WHEEL_DELTA * 4;

	for (int i = 0; i < nSteps; i++)
	{
		UINT nSBCode = zDelta < 0 ? SB_LINEDOWN : SB_LINEUP;

		if (bHasVertBar)
		{
			OnVScroll (nSBCode, 0, GetScrollBarCtrl(SB_VERT));
		}
		else
		{
			OnHScroll (nSBCode, 0, GetScrollBarCtrl(SB_HORZ));
		}
	}

	return TRUE;
}
//********************************************************************************
void CBCGPHotSpotImageCtrl::OnLButtonDown(UINT /*nFlags*/, CPoint point) 
{
	m_nSelectedNum = 0;
	m_pClicked = HitTest (point);

	SetFocus ();
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::OnLButtonUp(UINT /*nFlags*/, CPoint point) 
{
	if (m_pHot == NULL && m_pClicked == NULL)
	{
		return;
	}

	ReleaseCapture ();

	BOOL bIsHotSpotClick = (m_pHot != NULL && m_pHot == HitTest (point) &&
		m_pClicked == m_pHot);

	CRect rectHotSpot = (m_pHot != NULL) ? m_pHot->m_rect : CRect (0, 0, 0, 0);
	CRect rectClicked = (m_pClicked != NULL) ? m_pClicked->m_rect : CRect (0, 0, 0, 0);

	UINT nSelectedNum = m_pHot != NULL ? m_pHot->m_nID : 0;

	m_pHot = NULL;
	m_pClicked = NULL;

	if (HasHotImage () || m_nHighlightPercentage > 0)
	{
		RedrawRect (rectHotSpot);

		if (!rectClicked.IsRectEmpty () && rectClicked != rectHotSpot)
		{
			RedrawRect (rectClicked);
		}
	}

	if (bIsHotSpotClick)
	{
		CWnd* pParent = GetParent ();
		if (pParent != NULL)
		{
			m_nSelectedNum = nSelectedNum;

			HWND hwnd = GetSafeHwnd ();

			pParent->SendMessage (	WM_COMMAND,
									MAKEWPARAM (GetDlgCtrlID (), BN_CLICKED),
									(LPARAM) m_hWnd);

			if (::IsWindow (hwnd))
			{
				m_nSelectedNum = 0;
			}
		}
	}
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::OnCancelMode() 
{
	CButton::OnCancelMode();
	
	if (m_pHot != NULL || m_pClicked != NULL)
	{
		CRect rectHotSpot = m_pHot != NULL ? m_pHot->m_rect : CRect (0, 0, 0, 0);
		CRect rectClicked = (m_pClicked != NULL) ? m_pClicked->m_rect : CRect (0, 0, 0, 0);

		m_pHot = NULL;
		m_pClicked = NULL;
		ReleaseCapture ();

		if (HasHotImage () || m_nHighlightPercentage > 0)
		{
			RedrawRect (rectHotSpot);

			if (!rectClicked.IsRectEmpty () && rectClicked != rectHotSpot)
			{
				RedrawRect (rectClicked);
			}
		}
	}

	m_pClicked = NULL;
	m_nSelectedNum = 0;
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::OnSize(UINT nType, int cx, int cy) 
{
	CButton::OnSize(nType, cx, cy);
	
	UpdateScrollBars ();
	RebuildToolTips();
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	if (pScrollBar != NULL && pScrollBar->SendChildNotifyLastMsg())
		return;     // eat it

	// ignore scroll bar msgs from other controls
	if (pScrollBar != GetScrollBarCtrl(SB_VERT))
		return;

	OnScroll(MAKEWORD(-1, nSBCode), nPos);
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	if (pScrollBar != NULL && pScrollBar->SendChildNotifyLastMsg())
		return;     // eat it

	// ignore scroll bar msgs from other controls
	if (pScrollBar != GetScrollBarCtrl(SB_HORZ))
		return;

	OnScroll(MAKEWORD(nSBCode, -1), nPos);
}
//********************************************************************************
void CBCGPHotSpotImageCtrl::UpdateScrollBars ()
{
	if (GetSafeHwnd () == NULL)
	{
		return;
	}

	if (!m_bScrollBars)
	{
		SetScrollRange (SB_HORZ, 0, 0, TRUE);
		SetScrollRange (SB_VERT, 0, 0, TRUE);
		m_rectView.SetRectEmpty ();
		return;
	}

	static BOOL bInUpdate = FALSE;
	if (bInUpdate)
	{
		return;
	}

	bInUpdate = TRUE;

	CRect rectClient;
	GetClientRect (rectClient);

	const int cxScroll = ::GetSystemMetrics (SM_CXVSCROLL);
	const int cyScroll = ::GetSystemMetrics (SM_CYHSCROLL);

	m_rectView = rectClient;
	m_rectView.right -= cxScroll;
	m_rectView.bottom -= cyScroll;

	SCROLLINFO scrollInfo;
	scrollInfo.cbSize = sizeof (SCROLLINFO);
	scrollInfo.fMask = SIF_RANGE | SIF_PAGE;
	scrollInfo.nMin = 0;
	scrollInfo.nMax = m_sizeImage.cx;
	scrollInfo.nPage = m_rectView.Width ();
	BOOL bScrollHorz = (scrollInfo.nMax > 0 && (int)scrollInfo.nPage < scrollInfo.nMax);

	if (!SetScrollInfo (SB_HORZ, &scrollInfo))
	{
		SetScrollRange (SB_HORZ, 0, m_sizeImage.cx, TRUE);
	}

	scrollInfo.nMax = m_sizeImage.cy;
	scrollInfo.nPage = m_rectView.Height ();
	BOOL bScrollVert = (scrollInfo.nMax > 0 && (int)scrollInfo.nPage < scrollInfo.nMax);
		
	if (!SetScrollInfo (SB_VERT, &scrollInfo))
	{
		SetScrollRange (SB_VERT, 0, m_sizeImage.cy, TRUE);
	}

	m_rectView.right += cxScroll;
	m_rectView.bottom += cyScroll;

	if (bScrollHorz && bScrollVert && m_rectView.Width () >= m_sizeImage.cx && m_rectView.Height () >= m_sizeImage.cy)
	{
		bScrollHorz = bScrollVert = FALSE;
	}

	CSize szScroll (0, 0);
	if (bScrollVert)
	{
		if (globalData.m_bIsRTL)
		{
			m_rectView.left += cxScroll;
		}
		else
		{
			m_rectView.right -= cxScroll;
		}
		szScroll.cy = rectClient.Height ();
	}
	if (bScrollHorz)
	{
		m_rectView.bottom -= cyScroll;
		szScroll.cx = m_rectView.Width ();
		szScroll.cy = max (0, szScroll.cy - cyScroll);
	}

	if (m_wndScrollVert.GetSafeHwnd () != NULL)
	{
		m_wndScrollVert.EnableScrollBar (m_bScrollBars && bScrollVert ? ESB_ENABLE_BOTH : ESB_DISABLE_BOTH);
		m_wndScrollVert.EnableWindow(bScrollVert);

		if (bScrollVert)
		{
			m_wndScrollVert.SetWindowPos (NULL, globalData.m_bIsRTL ? rectClient.left : m_rectView.right, m_rectView.top,
				cxScroll, szScroll.cy, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
		}
		else
		{
			m_wndScrollVert.SetWindowPos (NULL, 0, 0,
				0, 0, SWP_NOZORDER | SWP_NOACTIVATE);
		}
	}

	if (m_wndScrollHorz.GetSafeHwnd () != NULL)
	{
		m_wndScrollHorz.EnableScrollBar (m_bScrollBars && bScrollHorz ? ESB_ENABLE_BOTH : ESB_DISABLE_BOTH);
		m_wndScrollHorz.EnableWindow(bScrollHorz);

		if (bScrollHorz)
		{
			m_wndScrollHorz.SetWindowPos (NULL, m_rectView.left, rectClient.bottom - cyScroll,
				szScroll.cx, cyScroll, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
		}
		else
		{
			m_wndScrollHorz.SetWindowPos (NULL, 0, 0,
				0, 0, SWP_NOZORDER | SWP_NOACTIVATE);
		}
	}

	bInUpdate = FALSE;
}
//*******************************************************************************
int CBCGPHotSpotImageCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CButton::OnCreate(lpCreateStruct) == -1)
		return -1;

	CRect rectDummy;
	rectDummy.SetRectEmpty ();

	m_wndScrollVert.Create (WS_CHILD | SBS_VERT, rectDummy, this, ID_SCROLL_VERT);
	m_wndScrollHorz.Create (WS_CHILD | SBS_HORZ, rectDummy, this, ID_SCROLL_HORZ);

	SetScrollBarsStyle (CBCGPScrollBar::BCGP_SBSTYLE_VISUAL_MANAGER);
	
	UpdateScrollBars ();
	return 0;
}
//*******************************************************************************
BOOL CBCGPHotSpotImageCtrl::OnScroll(UINT nScrollCode, UINT nPos, BOOL bDoScroll)
{
	if (m_pToolTip->GetSafeHwnd () != NULL)
	{
		m_pToolTip->Pop();
	}

	// calc new x position
	int x = GetScrollPos(SB_HORZ);
	int xOrig = x;

	switch (LOBYTE(nScrollCode))
	{
	case SB_TOP:
		x = 0;
		break;
	case SB_BOTTOM:
		x = INT_MAX;
		break;
	case SB_LINEUP:
		x -= m_lineDev.cx;
		break;
	case SB_LINEDOWN:
		x += m_lineDev.cx;
		break;
	case SB_PAGEUP:
		x -= m_pageDev.cx;
		break;
	case SB_PAGEDOWN:
		x += m_pageDev.cx;
		break;
	case SB_THUMBTRACK:
		x = nPos;
		break;
	}

	// calc new y position
	int y = GetScrollPos(SB_VERT);
	int yOrig = y;

	switch (HIBYTE(nScrollCode))
	{
	case SB_TOP:
		y = 0;
		break;
	case SB_BOTTOM:
		y = INT_MAX;
		break;
	case SB_LINEUP:
		y -= m_lineDev.cy;
		break;
	case SB_LINEDOWN:
		y += m_lineDev.cy;
		break;
	case SB_PAGEUP:
		y -= m_pageDev.cy;
		break;
	case SB_PAGEDOWN:
		y += m_pageDev.cy;
		break;
	case SB_THUMBTRACK:
		y = nPos;
		break;
	}

	BOOL bResult = OnScrollBy(CSize(x - xOrig, y - yOrig), bDoScroll);
	if (bResult && bDoScroll)
		UpdateWindow();

	return bResult;
}
//*******************************************************************************
BOOL CBCGPHotSpotImageCtrl::OnScrollBy(CSize sizeScroll, BOOL bDoScroll)
{
	if (m_pToolTip->GetSafeHwnd () != NULL)
	{
		m_pToolTip->Pop();
	}

	int xOrig, x;
	int yOrig, y;

	// don't scroll if there is no valid scroll range (ie. no scroll bar)
	CScrollBar* pBar;
	DWORD dwStyle = GetStyle();
	pBar = GetScrollBarCtrl(SB_VERT);
	if ((pBar != NULL && !pBar->IsWindowEnabled()) ||
		(pBar == NULL && !(dwStyle & WS_VSCROLL)))
	{
		// vertical scroll bar not enabled
		sizeScroll.cy = 0;
	}
	pBar = GetScrollBarCtrl(SB_HORZ);
	if ((pBar != NULL && !pBar->IsWindowEnabled()) ||
		(pBar == NULL && !(dwStyle & WS_HSCROLL)))
	{
		// horizontal scroll bar not enabled
		sizeScroll.cx = 0;
	}

	// adjust current x position
	xOrig = x = GetScrollPos(SB_HORZ);
	int xMax = GetScrollLimit(SB_HORZ);
	x += sizeScroll.cx;
	if (x < 0)
		x = 0;
	else if (x > xMax)
		x = xMax;

	// adjust current y position
	yOrig = y = GetScrollPos(SB_VERT);
	int yMax = GetScrollLimit(SB_VERT);
	y += sizeScroll.cy;
	if (y < 0)
		y = 0;
	else if (y > yMax)
		y = yMax;

	// did anything change?
	if (x == xOrig && y == yOrig)
		return FALSE;

	if (bDoScroll)
	{
		if (x != xOrig)
			SetScrollPos(SB_HORZ, x);
		if (y != yOrig)
			SetScrollPos(SB_VERT, y);

		if (!m_rectView.IsRectEmpty ())
		{
			RedrawWindow (&m_rectView, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_NOCHILDREN);
		}
		else
		{
			RedrawWindow ();
		}
	}

	RebuildToolTips();
	return TRUE;
}
//*******************************************************************************
CPoint CBCGPHotSpotImageCtrl::GetDeviceScrollPosition() const
{
	CPoint pt(GetScrollPos(SB_HORZ), GetScrollPos(SB_VERT));
	ASSERT(pt.x >= 0 && pt.y >= 0);

	return pt;
}
//******************************************************************************************
LRESULT CBCGPHotSpotImageCtrl::OnStyleChanging (WPARAM wp, LPARAM lp)
{
	int nStyleType = (int) wp; 
	LPSTYLESTRUCT lpStyleStruct = (LPSTYLESTRUCT) lp;

	CButton::OnStyleChanging (nStyleType, lpStyleStruct);

	if (nStyleType == GWL_EXSTYLE)
	{
		lpStyleStruct->styleNew &= ~WS_EX_LAYOUTRTL;
	}

	return 0;
}
//*******************************************************************************
void CBCGPHotSpotImageCtrl::SetHighlightHotArea (int nPercentage, int nTolerance/* = 0*/)
{
	m_nHighlightPercentage = nPercentage;
	m_nHighlightTolerance = nTolerance;
}
//*********************************************************************************
void CBCGPHotSpotImageCtrl::EnableScrollBars (BOOL bEnable/* = TRUE*/)
{
	if (m_bScrollBars == bEnable)
	{
		return;
	}

	m_bScrollBars = bEnable;
	UpdateScrollBars ();
}
//*********************************************************************************
CScrollBar* CBCGPHotSpotImageCtrl::GetScrollBarCtrl(int nBar) const
{
	if (nBar == SB_HORZ && m_wndScrollHorz.GetSafeHwnd () != NULL)
	{
		return (CScrollBar* ) &m_wndScrollHorz;
	}
	
	if (nBar == SB_VERT && m_wndScrollVert.GetSafeHwnd () != NULL)
	{
		return (CScrollBar* ) &m_wndScrollVert;
	}

	return NULL;
}
//*********************************************************************************
LRESULT CBCGPHotSpotImageCtrl::OnChangeVisualManager(WPARAM, LPARAM)
{
	return 0;
}
//*********************************************************************************
BOOL CBCGPHotSpotImageCtrl::PreTranslateMessage(MSG* pMsg) 
{
   	switch (pMsg->message)
	{
	case WM_KEYDOWN:
	case WM_SYSKEYDOWN:
	case WM_LBUTTONDOWN:
	case WM_RBUTTONDOWN:
	case WM_MBUTTONDOWN:
	case WM_LBUTTONUP:
	case WM_RBUTTONUP:
	case WM_MBUTTONUP:
	case WM_NCLBUTTONDOWN:
	case WM_NCRBUTTONDOWN:
	case WM_NCMBUTTONDOWN:
	case WM_NCLBUTTONUP:
	case WM_NCRBUTTONUP:
	case WM_NCMBUTTONUP:
	case WM_MOUSEMOVE:
		if (m_pToolTip->GetSafeHwnd () != NULL)
		{
			m_pToolTip->RelayEvent(pMsg);
		}
		break;
	}
	
	return CButton::PreTranslateMessage(pMsg);
}
//***********************************************************************************************************
BOOL CBCGPHotSpotImageCtrl::OnTTNeedTipText (UINT /*id*/, NMHDR* pNMH, LRESULT* /*pResult*/)
{
	static CString strTipText;

	int nItem = ((int)pNMH->idFrom) - 1;
    if (nItem < 0)
	{
        return FALSE;
	}

	int i = 0;
	
	for (POSITION pos = m_lstHotSpots.GetHeadPosition (); pos != NULL; i++)
	{
		CBCGPHotSpot* pHotSpot = m_lstHotSpots.GetNext (pos);
		ASSERT_VALID (pHotSpot);

		if (i == nItem)
		{
			if (pHotSpot->GetToolTip().IsEmpty())
			{
				return FALSE;
			}

			strTipText = pHotSpot->GetToolTip();

			LPNMTTDISPINFO	pTTDispInfo	= (LPNMTTDISPINFO) pNMH;
			ASSERT((pTTDispInfo->uFlags & TTF_IDISHWND) == 0);

			if (!pHotSpot->GetToolTipDescription().IsEmpty())
			{
				CBCGPToolTipCtrl* pToolTip = DYNAMIC_DOWNCAST(CBCGPToolTipCtrl, m_pToolTip);
				if (pToolTip != NULL)
				{
					ASSERT_VALID (pToolTip);
					pToolTip->SetDescription(pHotSpot->GetToolTipDescription());
				}
			}
			
			pTTDispInfo->lpszText = const_cast<LPTSTR> ((LPCTSTR)strTipText);
			return TRUE;
		}
	}

	return FALSE;
}
//**************************************************************************
LRESULT CBCGPHotSpotImageCtrl::OnBCGUpdateToolTips(WPARAM wp, LPARAM)
{
	UINT nTypes = (UINT) wp;

	if (m_pToolTip->GetSafeHwnd () == NULL)
	{
		return 0;
	}

	if (nTypes & BCGP_TOOLTIP_TYPE_DEFAULT)
	{
		RebuildToolTips();
	}

	return 0;
}
//**************************************************************************
void CBCGPHotSpotImageCtrl::RebuildToolTips()
{
	CBCGPTooltipManager::DeleteToolTip(m_pToolTip);
	CBCGPTooltipManager::CreateToolTip(m_pToolTip, this, BCGP_TOOLTIP_TYPE_DEFAULT);

	int i = 0;
	CPoint ptScrollOffset (GetScrollPos (SB_HORZ), GetScrollPos (SB_VERT));

	for (POSITION pos = m_lstHotSpots.GetHeadPosition (); pos != NULL; i++)
	{
		CBCGPHotSpot* pHotSpot = m_lstHotSpots.GetNext (pos);
		ASSERT_VALID (pHotSpot);

		if (!pHotSpot->GetToolTip().IsEmpty())
		{
		   CRect rectTT = pHotSpot->m_rect;
		   rectTT.OffsetRect(-ptScrollOffset);

			m_pToolTip->AddTool(this, LPSTR_TEXTCALLBACK, rectTT, i + 1);
		}
	}

	m_bRebuildTooltips = FALSE;
}

#endif // BCGP_EXCLUDE_HOT_SPOT_IMAGE
