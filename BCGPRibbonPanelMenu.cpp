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
// BCGPRibbonPanelMenu.cpp : implementation file
//

#include "stdafx.h"
#include "trackmouse.h"
#include "BCGPRibbonPanelMenu.h"
#include "BCGPRibbonPanel.h"
#include "BCGPRibbonCategory.h"
#include "BCGPRibbonBar.h"
#include "BCGPVisualManager.h"
#include "BCGPTooltipManager.h"
#include "BCGPToolTipCtrl.h"
#include "BCGPRibbonPaletteButton.h"
#include "BCGPRibbonFloaty.h"
#include "BCGPRibbonEdit.h"

#ifndef BCGP_EXCLUDE_RIBBON

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static const int uiPopupTimerEvent = 1;
static const int uiRemovePopupTimerEvent = 2;
static const UINT IdAutoCommand = 3;
static const int nScrollBarID = 1;
static const int nScrollBarHorzID = 2;

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonPanelMenuBar window

IMPLEMENT_DYNAMIC(CBCGPRibbonPanelMenuBar, CBCGPPopupMenuBar)

#pragma warning (disable : 4355)

CBCGPRibbonPanelMenuBar::CBCGPRibbonPanelMenuBar(CBCGPRibbonPanel* pPanel)
{
	ASSERT_VALID (pPanel);

	m_pPanel = DYNAMIC_DOWNCAST (CBCGPRibbonPanel,
		pPanel->GetRuntimeClass ()->CreateObject ());
	ASSERT_VALID (m_pPanel);

	m_pPanel->CopyFrom (*pPanel);

	CommonInit ();

	m_pPanelOrigin = pPanel;

	m_pPanel->SetParentMenuBar(this);

	m_pRibbonBar = m_pPanel->m_pParent->GetParentRibbonBar ();
}
//*******************************************************************************
CBCGPRibbonPanelMenuBar::CBCGPRibbonPanelMenuBar(
	CBCGPRibbonBar* pRibbonBar,
	const CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arButtons,
	BOOL bIsFloatyMode)
{
	m_pPanel = new CBCGPRibbonPanel;

	CommonInit ();
	AddButtons (pRibbonBar, arButtons, bIsFloatyMode);
}
//*******************************************************************************
CBCGPRibbonPanelMenuBar::CBCGPRibbonPanelMenuBar(CBCGPRibbonPaletteButton* pPaletteButton)
{
	ASSERT_VALID (pPaletteButton);

	m_pPanel = new CBCGPRibbonPanel (pPaletteButton);

	CommonInit ();

	// Create array without scroll buttons:
	CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arButtons;
	pPaletteButton->GetMenuItems (arButtons);

	AddButtons (pPaletteButton->GetTopLevelRibbonBar (), arButtons, FALSE);
}
//*******************************************************************************
CBCGPRibbonPanelMenuBar::CBCGPRibbonPanelMenuBar (CBCGPRibbonCategory* pCategory, CSize size)
{
	ASSERT_VALID (pCategory);

	m_pPanel = NULL;

	CommonInit ();

	m_pCategory = (CBCGPRibbonCategory*) pCategory->GetRuntimeClass ()->CreateObject ();
	ASSERT_VALID (m_pCategory);

	m_pCategory->CopyFrom (*pCategory);
	m_pCategory->m_pParentMenuBar = this;

	for (int iPanel = 0; iPanel < m_pCategory->GetPanelCount (); iPanel++)
	{
		CBCGPRibbonPanel* pPanel = m_pCategory->GetPanel (iPanel);
		ASSERT_VALID (pPanel);

		pPanel->SetParentMenuBar(this);
	}

	m_sizeCategory = size;
	m_pRibbonBar = m_pCategory->GetParentRibbonBar ();
}
//*******************************************************************************
CBCGPRibbonPanelMenuBar::CBCGPRibbonPanelMenuBar(CBCGPRibbonBar* pRibbonBar)
{
	ASSERT_VALID(pRibbonBar);

	m_pPanel = NULL;

	CommonInit();

	m_pRibbonBar = pRibbonBar;
	m_pRibbonBar->SetAutoHidePopup(this);
}
//*******************************************************************************
void CBCGPRibbonPanelMenuBar::AddButtons (CBCGPRibbonBar* pRibbonBar,
	const CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arButtons,
	BOOL bFloatyMode)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pRibbonBar);
	ASSERT_VALID(m_pPanel);

	m_bSimpleMode = TRUE;
	m_pRibbonBar = pRibbonBar;

	m_pPanel->m_pParentMenuBar = this;
	m_pPanel->m_bFloatyMode = bFloatyMode;
	m_pPanel->m_nXMargin = 2;
	m_pPanel->m_nYMargin = bFloatyMode ? globalUtils.ScaleByDPI(4) : 2;
	m_pPanel->RemoveAll ();

	for (int i = 0; i < arButtons.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pSrcButton = arButtons [i];
		ASSERT_VALID (pSrcButton);

		CBCGPBaseRibbonElement* pButton =
			(CBCGPBaseRibbonElement*) pSrcButton->GetRuntimeClass ()->CreateObject ();
		ASSERT_VALID (pButton);

		pButton->CopyFrom (*pSrcButton);
		pButton->SetOriginal (pSrcButton);
		pButton->m_bCompactMode = TRUE;

		pButton->SetCommandSearchMenu(pSrcButton->IsCommandSearchMenu());

		pButton->SetParentMenu (this);

		m_pPanel->Add (pButton);
	}
}
//*******************************************************************************
CBCGPRibbonPanelMenuBar::CBCGPRibbonPanelMenuBar()
{
	m_pPanel = new CBCGPRibbonPanel;
	CommonInit ();
}
//*******************************************************************************
void CBCGPRibbonPanelMenuBar::CommonInit ()
{
	if (m_pPanel != NULL)
	{
		ASSERT_VALID (m_pPanel);
		m_pPanel->m_pParentMenuBar = this;
	}

	m_pCategory = NULL;
	m_sizeCategory = CSize (0, 0);

	m_pDelayedCloseButton = NULL;
	m_pDelayedButton = NULL;
	m_pPressed = NULL;
	m_rectAutoCommand.SetRectEmpty ();

	m_bSimpleMode = FALSE;
	m_bIsMenuMode = FALSE;
	m_bIsDefaultMenuLook = FALSE;
	m_pPanelOrigin = NULL;
	m_pRibbonBar = NULL;

	m_bTracked = FALSE;
	m_pToolTip = NULL;
	m_bDisableSideBarInXPMode = TRUE;

	m_sizePrefered = CSize (0, 0);
	m_bIsQATPopup = FALSE;
	m_bCustomizeMenu = TRUE;
	m_bIsFloaty = FALSE;
	m_bSetKeyTips = FALSE;
	m_bHasKeyTips = FALSE;
	m_bAutoCommandTimer = FALSE;
	m_bIsOneRowFloaty = FALSE;

	m_ptStartMenu = CPoint (-1, -1);

	m_bIsCtrlMode = FALSE;
}

#pragma warning (default : 4355)

CBCGPRibbonPanelMenuBar::~CBCGPRibbonPanelMenuBar()
{
	if (m_pPanel != NULL)
	{
		ASSERT_VALID (m_pPanel);

		if (m_pRibbonBar != NULL && m_pRibbonBar->GetKeyboardNavLevelCurrent () == m_pPanel)
		{
			m_pRibbonBar->DeactivateKeyboardFocus (FALSE);
		}

		delete m_pPanel;
	}

	if (m_pCategory != NULL)
	{
		ASSERT_VALID (m_pCategory);

		if (m_pRibbonBar != NULL && m_pRibbonBar->GetKeyboardNavLevelCurrent () == m_pCategory)
		{
			m_pRibbonBar->DeactivateKeyboardFocus (FALSE);
		}

		delete m_pCategory;

		if (m_pRibbonBar != NULL && m_pRibbonBar->GetActiveCategory () != NULL)
		{
			// Redraw ribbon tab:
			ASSERT_VALID (m_pRibbonBar);
			ASSERT_VALID (m_pRibbonBar->GetActiveCategory ());

			if (!m_pRibbonBar->IsQuickAccessToolbarOnTop ())
			{
				CBCGPRibbonTab& tab = m_pRibbonBar->GetActiveCategory ()->m_Tab;
				
				tab.m_bIsDroppedDown = FALSE;
				tab.m_bIsHighlighted = FALSE;

				CRect rectRedraw = tab.GetRect ();
				rectRedraw.bottom = m_pRibbonBar->GetQuickAccessToolbarLocation ().bottom;
				rectRedraw.InflateRect (1, 1);

				m_pRibbonBar->RedrawWindow (rectRedraw);
			}
		}
	}

	if (m_bHasKeyTips)
	{
		CWnd* pMenu = CBCGPPopupMenu::GetActiveMenu ();

		if (pMenu != NULL && CWnd::FromHandlePermanent (pMenu->GetSafeHwnd ()) != NULL &&
			pMenu->IsWindowVisible ())
		{
			CBCGPPopupMenu::UpdateAllShadows ();
		}
	}
}

BEGIN_MESSAGE_MAP(CBCGPRibbonPanelMenuBar, CBCGPPopupMenuBar)
	//{{AFX_MSG_MAP(CBCGPRibbonPanelMenuBar)
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONUP()
	ON_WM_LBUTTONDOWN()
	ON_WM_CREATE()
	ON_WM_DESTROY()
	ON_WM_SIZE()
	ON_WM_TIMER()
	ON_WM_CONTEXTMENU()
	ON_WM_VSCROLL()
	ON_WM_LBUTTONDBLCLK()
	ON_WM_HSCROLL()
	ON_WM_SETFOCUS()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_MOUSELEAVE, OnMouseLeave)
	ON_REGISTERED_MESSAGE(BCGM_UPDATETOOLTIPS, OnBCGUpdateToolTips)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnNeedTipText)
	ON_MESSAGE(WM_GESTURE, OnGestureEvent)
	ON_MESSAGE(WM_TABLET_QUERYSYSTEMGESTURESTATUS, OnTabletQuerySystemGestureStatus)
END_MESSAGE_MAP()

void CBCGPRibbonPanelMenuBar::AdjustLocations ()
{
	ASSERT_VALID (this);

	if (m_bInUpdateShadow)
	{
		return;
	}

	if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
	{
		m_pRibbonBar->RecalcLayout();
		return;
	}

	CRect rectClient;
	GetClientRect (rectClient);

	if (m_bIsCtrlMode)
	{
		rectClient.DeflateRect(2, 1);
	}

	CClientDC dc (this);

	BOOL bIsBackStageView = m_pPanel != NULL && m_pPanel->IsBackstageView();

	CFont* pOldFont = dc.SelectObject (m_pRibbonBar == NULL ? 
		&globalData.fontRegular : 
		bIsBackStageView ? m_pRibbonBar->GetBackstageViewItemFont() : m_pRibbonBar->GetFont());
	ASSERT (pOldFont != NULL);

	if (m_pCategory != NULL)
	{
		ASSERT_VALID (m_pCategory);

		m_pCategory->m_rect = rectClient;
		m_pCategory->RecalcLayout (&dc);
	}
	else if (m_pPanel != NULL)
	{
		m_pPanel->m_bSizeIsLocked = m_bResizeTracking || m_bIsCtrlMode;

		if (!IsBackStageView())
		{
			m_pPanel->m_nScrollOffset = m_iOffset;
		}

		m_pPanel->Repos (&dc, rectClient);
		m_pPanel->OnAfterChangeRect (&dc);
		
		CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar ();
		if (pRibbonBar != NULL && pRibbonBar->GetKeyboardNavLevelCurrent () == m_pPanel)
		{
			pRibbonBar->ShowKeyTips (TRUE);
		}

		m_pPanel->m_bSizeIsLocked = FALSE;
	}

	dc.SelectObject (pOldFont);
}
//*****************************************************************************************
void CBCGPRibbonPanelMenuBar::SetPreferedSize (CSize size)
{
	ASSERT_VALID (this);

	CSize sizePalette (0, 0);

	if (m_pPanel != NULL)
	{
		ASSERT_VALID (m_pPanel);

		if (m_pPanel->m_pPaletteButton != NULL)
		{
			sizePalette = m_pPanel->GetPaletteMinSize ();
			sizePalette.cx -= ::GetSystemMetrics (SM_CXVSCROLL);
		}
	}

	m_sizePrefered = CSize (max (size.cx, sizePalette.cx), size.cy);
}
//*****************************************************************************************
CSize CBCGPRibbonPanelMenuBar::CalcSize (BOOL /*bVertDock*/)
{
	ASSERT_VALID (this);

	if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
	{
		// Auto hide mode
		CRect rectRibbon;
		m_pRibbonBar->GetWindowRect(rectRibbon);
		
		m_pRibbonBar->m_bRecalcCategoryHeight = TRUE;
		CSize size = m_pRibbonBar->CalcFixedLayout(FALSE, TRUE);
		
		if (size != CSize(0, 0))
		{
			size.cx = rectRibbon.Width();
		}
		
		return size;
	}

	if (m_pCategory != NULL)
	{
		ASSERT_VALID (m_pCategory);
		ASSERT (m_sizeCategory != CSize (0, 0));

		return m_sizeCategory;
	}
	
	if (m_pPanel == NULL)
	{
		return CSize(0, 0);
	}

	ASSERT_VALID (m_pRibbonBar);
	ASSERT_VALID (m_pPanel);

	m_pPanel->m_bIsQATPopup = m_bIsQATPopup;

	CClientDC dc (m_pRibbonBar);

	CFont* pOldFont = dc.SelectObject (m_pPanel->IsBackstageView() ? m_pRibbonBar->GetBackstageViewItemFont() : m_pRibbonBar->GetFont());
	ASSERT (pOldFont != NULL);

	if (m_bIsMenuMode)
	{
		m_pPanel->m_bMenuMode = TRUE;
		m_pPanel->m_bIsDefaultMenuLook = m_bIsDefaultMenuLook;
		
		m_pPanel->Repos (&dc, CRect (0, 0, m_sizePrefered.cx, m_sizePrefered.cy));

		dc.SelectObject (pOldFont);

		CSize size = m_pPanel->m_rect.Size ();

		if (m_sizePrefered != CSize (0, 0))
		{
			size.cx = max (m_sizePrefered.cx, size.cx);

			if (m_sizePrefered.cy <= 0)
			{
				size.cy = m_pPanel->m_rect.Size ().cy;
			}
			else
			{
				if (m_pPanel->m_pPaletteButton != NULL)
				{
					size.cy = max (size.cy, m_sizePrefered.cy);
				}
				else
				{
					if (size.cy > m_sizePrefered.cy)
					{
						CBCGPPopupMenu* pParentMenu = DYNAMIC_DOWNCAST (CBCGPPopupMenu, GetParent ());
						if (pParentMenu != NULL)
						{
							pParentMenu->m_bScrollable = TRUE;
						}
					}

					size.cy = m_sizePrefered.cy;
				}
			}
		}

		return size;
	}

	if (m_bSimpleMode && m_pPanel->m_arWidths.GetSize () == 0)
	{
		CWaitCursor wait;
		m_pPanel->RecalcWidths (&dc, 32767);
	}

	const int nWidthSize = (int) m_pPanel->m_arWidths.GetSize ();
	if (nWidthSize == 0)
	{
		dc.SelectObject (pOldFont);
		return CSize (10, 10);
	}

	if (m_pPanel->m_bAlignByColumn && !m_pPanel->m_bFloatyMode && !m_pPanel->IsFixedSize ())
	{
		int nCategoryHeight = m_pRibbonBar->GetCategoryHeight ();
		if (nCategoryHeight <= 0)
		{
			nCategoryHeight = m_pPanel->GetHeight (&dc) + m_pPanel->GetCaptionSize (&dc).cy;
		}

		const int nHeight = nCategoryHeight - 2 * m_pPanel->m_nYMargin;
		m_pPanel->Repos (&dc, CRect (0, 0, 32767, nHeight));
	}
	else if (m_bIsQATPopup)
	{
		int nWidth = m_pPanel->m_arWidths [0] + 2 * m_pPanel->m_nXMargin;
		m_pPanel->Repos (&dc, CRect (0, 0, nWidth, 32767));
	}
	else
	{
		int nWidth = 0;
		int nHeight = 0;

		if (!m_pPanel->m_bFloatyMode)
		{
			nWidth = m_pPanel->m_arWidths [0] + 4 * m_pPanel->m_nXMargin;

			int nCategoryHeight = m_pRibbonBar->GetCategoryHeight ();
			if (nCategoryHeight <= 0)
			{
				nCategoryHeight = m_pPanel->GetHeight (&dc) + m_pPanel->GetCaptionSize (&dc).cy;
			}

			nHeight = nCategoryHeight - 2 * m_pPanel->m_nYMargin;
		}
		else
		{
			nWidth = m_pPanel->m_arWidths [nWidthSize > 2 && !m_bIsOneRowFloaty ? 1 : 0] + 4 * m_pPanel->m_nXMargin;
			nHeight = 32767;
		}

		m_pPanel->Repos (&dc, CRect (0, 0, nWidth, nHeight));
	}

	CSize size = m_pPanel->m_rect.Size ();
	dc.SelectObject (pOldFont);

	if (m_bSimpleMode && m_pPanel->GetCount () > 0 && !m_bIsQATPopup)
	{
		int xMin = 32767;
		int xMax = 0;

		int yMin = 32767;
		int yMax = 0;

		for (int i = 0; i < m_pPanel->GetCount (); i++)
		{
			CBCGPBaseRibbonElement* pButton = m_pPanel->GetElement (i);
			ASSERT_VALID (pButton);

			CRect rectButton = pButton->GetRect ();

			xMin = min (xMin, rectButton.left);
			yMin = min (yMin, rectButton.top);

			xMax = max (xMax, rectButton.right);
			yMax = max (yMax, rectButton.bottom);
		}

		return CSize (
			xMax - xMin + 2 * m_pPanel->m_nXMargin, 
			yMax - yMin + 2 * m_pPanel->m_nYMargin);
	}

	return size;
}
//*****************************************************************************************
void CBCGPRibbonPanelMenuBar::DoPaint (CDC* pDCPaint)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDCPaint);

	CBCGPMemDC memDC (*pDCPaint, this);
	CDC* pDC = &memDC.GetDC ();

	CRect rectClip(0, 0, 0, 0);
	pDCPaint->GetClipBox (rectClip);

	CRgn rgnClip;

	if (!rectClip.IsRectEmpty ())
	{
		rgnClip.CreateRectRgnIndirect (rectClip);
		pDC->SelectClipRgn (&rgnClip);
	}
	
	BOOL bIsBackStageView = m_pPanel != NULL && m_pPanel->IsBackstageView();

	CFont* pOldFont = pDC->SelectObject (m_pRibbonBar == NULL ? 
		&globalData.fontRegular : 
		bIsBackStageView ? m_pRibbonBar->GetBackstageViewItemFont() : m_pRibbonBar->GetFont());

	ASSERT (pOldFont != NULL);

	pDC->SetBkMode (TRANSPARENT);

	CRect rectClient;
	GetClientRect (rectClient);

	CRect rectFill = rectClient;
	rectFill.InflateRect (3, 3);

	if (m_pCategory != NULL)
	{
		ASSERT_VALID (m_pCategory);

		CBCGPVisualManager::GetInstance ()->OnDrawRibbonCategory (
			pDC, m_pCategory, rectFill);

		m_pCategory->OnDraw (pDC);
	}
	else if (m_pPanel != NULL)
	{
		ASSERT_VALID (m_pPanel);

		if (m_bIsCtrlMode)
		{
			CBCGPVisualManager::GetInstance()->OnFillRibbonPanelMenuBarInCtrlMode(pDC, this, rectClient);
		}
		else
		{
			if (m_pPanel->m_pParent != NULL)
			{
				CBCGPVisualManager::GetInstance ()->OnDrawRibbonDropDownPanel(pDC, m_pPanel, rectFill);
			}
			else if (m_bIsQATPopup)
			{
				CBCGPVisualManager::GetInstance ()->OnFillRibbonQATPopup (
					pDC, this, rectClient);
			}
			else
			{
				CBCGPVisualManager::GetInstance ()->OnFillBarBackground (
					pDC, this, rectClient, rectClient);
			}
		}

		m_pPanel->DoPaint (pDC);
	}
	else if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
	{
		m_pRibbonBar->OnDraw(pDC);
	}

	if (m_bIsCtrlMode && (GetStyle() & WS_BORDER) != 0)
	{
		CBCGPVisualManager::GetInstance()->OnDrawControlBorder(pDC, rectClient, this, FALSE);
	}

	pDC->SelectObject (pOldFont);

	if (rgnClip.GetSafeHandle() != NULL)
	{
		pDC->SelectClipRgn (NULL);
	}
}
//*****************************************************************************************
void CBCGPRibbonPanelMenuBar::OnMouseMove(UINT nFlags, CPoint point) 
{
	CBCGPPopupMenuBar::OnMouseMove(nFlags, point);

	if (m_pPanel != NULL && globalData.IsAccessibilitySupport())
	{
		const int nIndex = m_pPanel->HitTestEx (point);
		if (nIndex != -1 && nIndex != m_iAccHotItem)
		{
			m_iAccHotItem = nIndex;

			KillTimer(uiAccNotifyEvent);
			SetTimer(uiAccNotifyEvent, uiAccTimerDelay, NULL);
		}
	}

	if (m_pPanel != NULL &&
		!m_pPanel->m_bMenuMode && m_pPanel->GetDroppedDown () != NULL)
	{
		return;
	}

	if (m_pCategory != NULL && m_pCategory->GetDroppedDown () != NULL)
	{
		return;
	}

	if (m_ptStartMenu != CPoint (-1, -1))
	{
		// Check if mouse was moved from the menu start point:

		CPoint ptCursor;
		::GetCursorPos (&ptCursor);

		if (abs (ptCursor.x - m_ptStartMenu.x) < 10 &&
			abs (ptCursor.y - m_ptStartMenu.y) < 10)
		{
			return;
		}

		m_ptStartMenu = CPoint (-1, -1);
	}

	if (point == CPoint (-1, -1))
	{
		m_bTracked = FALSE;
	}
	else if (!m_bTracked)
	{
		m_bTracked = TRUE;
		
		TRACKMOUSEEVENT trackmouseevent;
		trackmouseevent.cbSize = sizeof(trackmouseevent);
		trackmouseevent.dwFlags = TME_LEAVE;
		trackmouseevent.hwndTrack = GetSafeHwnd();
		trackmouseevent.dwHoverTime = HOVER_DEFAULT;
		::BCGPTrackMouse (&trackmouseevent);

		CBCGPBaseRibbonElement* pPressed = NULL;

		if (m_pCategory != NULL)
		{
		}
		else if (m_pPanel != NULL)
		{
			pPressed = m_pPanel->GetPressed ();
		}
		else if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
		{
			pPressed = m_pRibbonBar->m_pPressed;
		}

		if (pPressed != NULL && ((nFlags & MK_LBUTTON) == 0))
		{
			ASSERT_VALID (pPressed);
			pPressed->m_bIsPressed = FALSE;
		}
	}

	if (m_pCategory != NULL)
	{
		m_pCategory->OnMouseMove (point);
	}
	else if (m_pPanel != NULL)
	{
		BOOL bWasHighlighted = m_pPanel->IsHighlighted ();
		m_pPanel->Highlight (TRUE, point);

		if (!bWasHighlighted)
		{
			UINT uiRedrawFlags = IsBackStageView() ? RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE : RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN;
			RedrawWindow (NULL, NULL, uiRedrawFlags);
		}
	}
	else if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
	{
		m_pRibbonBar->OnMouseMove(nFlags, point);
	}
}
//*****************************************************************************************
LRESULT CBCGPRibbonPanelMenuBar::OnMouseLeave(WPARAM,LPARAM)
{
	CPoint point;
	::GetCursorPos (&point);
	ScreenToClient (&point);

	CRect rectClient;
	GetClientRect (rectClient);

	if (!rectClient.PtInRect (point))
	{
		OnMouseMove (0, CPoint (-1, -1));
		m_bTracked = FALSE;

		if (m_pPanel != NULL)
		{
			m_pPanel->Highlight (FALSE, CPoint (-1, -1));
		}

		UINT uiRedrawFlags = IsBackStageView() ? RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE : RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN;
		RedrawWindow (NULL, NULL, uiRedrawFlags);
	}

	m_bTracked = FALSE;
	return 0;
}
//*****************************************************************************************
void CBCGPRibbonPanelMenuBar::OnLButtonUp(UINT nFlags, CPoint point) 
{
	if (m_bAutoCommandTimer)
	{
		KillTimer (IdAutoCommand);
		m_bAutoCommandTimer = FALSE;
		m_pPressed = NULL;
		m_rectAutoCommand.SetRectEmpty ();
	}

	HWND hwndThis = GetSafeHwnd ();

	CBCGPPopupMenuBar::OnLButtonUp(nFlags, point);

	if (!::IsWindow (hwndThis))
	{
		return;
	}

	if (m_pCategory != NULL)
	{
		m_pCategory->OnLButtonUp (point);
	}
	else if (m_pPanel != NULL)
	{
		m_pPanel->MouseButtonUp (point);
	}
	else if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
	{
		m_pRibbonBar->OnLButtonUp(nFlags, point);
	}

	if (::IsWindow (hwndThis))
	{
		::GetCursorPos (&point);
		ScreenToClient (&point);

		OnMouseMove (nFlags, point);

		if (m_bIsCtrlMode)
		{
			RedrawWindow();
		}
	}
}
//*****************************************************************************************
void CBCGPRibbonPanelMenuBar::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CBCGPPopupMenuBar::OnLButtonDown(nFlags, point);

	CBCGPBaseRibbonElement* pDroppedDown = GetDroppedDown ();
	if (pDroppedDown != NULL)
	{
		ASSERT_VALID (pDroppedDown);
		pDroppedDown->ClosePopupMenu ();
	}

	OnMouseMove (nFlags, point);

	m_pPressed = NULL;
	m_rectAutoCommand.SetRectEmpty ();

	HWND hwndThis = GetSafeHwnd ();

	CBCGPBaseRibbonElement* pPressed = NULL;

	if (m_pCategory != NULL)
	{
		pPressed = m_pCategory->OnLButtonDown (point);
	}
	else if (m_pPanel != NULL)
	{
		pPressed = m_pPanel->MouseButtonDown (point);
	}
	else if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
	{
		m_pRibbonBar->OnLButtonDown(nFlags, point);

		if (::IsWindow (hwndThis))
		{
			pPressed = m_pRibbonBar->m_pPressed;
		}
	}

	if (!::IsWindow (hwndThis))
	{
		return;
	}

	m_pPressed = pPressed;

	if (m_pPressed != NULL)
	{
		ASSERT_VALID (m_pPressed);

		int nDelay = 100;

		CBCGPRibbonBar* pRibbonBar = m_pPressed->GetTopLevelRibbonBar();
		if (pRibbonBar != NULL)
		{
			nDelay = pRibbonBar->GetAutoRepeatCmdDelay();
		}

		if (m_pPressed->IsAutoRepeatMode (nDelay))
		{
			SetTimer (IdAutoCommand, nDelay, NULL);
			m_bAutoCommandTimer = TRUE;
			m_rectAutoCommand = m_pPressed->GetRect ();
		}
	}

	if (m_bIsCtrlMode)
	{
		SetFocus();
	}
}
//*******************************************************************************
void CBCGPRibbonPanelMenuBar::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	CBCGPPopupMenuBar::OnLButtonDblClk(nFlags, point);

	if (IsRibbonPanelInRegularMode ())
	{
		CBCGPRibbonButton* pDroppedDown = GetDroppedDown ();
		if (pDroppedDown != NULL)
		{
			pDroppedDown->ClosePopupMenu ();
		}
	}

	CBCGPBaseRibbonElement* pHit = HitTest (point);
	if (pHit != NULL)
	{
		ASSERT_VALID (pHit);

		pHit->OnLButtonDblClk (point);
	}
}
//*******************************************************************************
void CBCGPRibbonPanelMenuBar::OnClickButton (CBCGPRibbonButton* pButton, CPoint /*point*/)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pButton);

	pButton->m_bIsHighlighted = pButton->m_bIsPressed = FALSE;
	RedrawWindow (pButton->GetRect ());

	if (m_pPanel != NULL)
	{
		ASSERT_VALID (m_pPanel);

		if (m_pPanel->m_pPaletteButton != NULL)
		{
			ASSERT_VALID (m_pPanel->m_pPaletteButton);
			
			if (m_pPanel->m_pPaletteButton->OnClickPaletteSubItem (pButton, this))
			{
				return;
			}
		}
	}

	BOOL bInvoked = pButton->NotifyCommand (TRUE);

	if (m_bIsCtrlMode || (pButton->DontCloseParentPopupOnClick() && !pButton->IsCommandSearchMenu()))
	{
		return;
	}

	if (pButton->IsCommandSearchMenu())
	{
		CBCGPRibbonBar* pRibbonBar = pButton->GetTopLevelRibbonBar();
		if (pRibbonBar != NULL)
		{
			if (pRibbonBar->IsInAutoHideMode())
			{
				pRibbonBar->ShowAutoHidePopup(FALSE);
				return;
			}

			if (pRibbonBar->GetTopLevelFrame () != NULL)
			{
				pRibbonBar->GetTopLevelFrame ()->SetFocus ();
				return;
			}
		}
	}

	if (IsFloaty ())
	{
		CBCGPRibbonFloaty* pFloaty = DYNAMIC_DOWNCAST (
			CBCGPRibbonFloaty, GetParent ());

		if (pFloaty != NULL)
		{
			if (pFloaty->IsContextMenuMode ())
			{
				pFloaty->CancelContextMenuMode ();
			}

			return;
		}
	}

	if (bInvoked)
	{
		CBCGPRibbonPanelMenu* pPopupMenu = 
			DYNAMIC_DOWNCAST (CBCGPRibbonPanelMenu, GetParent ());
		if (pPopupMenu != NULL)
		{
			ASSERT_VALID(pPopupMenu);

			CBCGPRibbonPanelMenu* pParentMenu = 
				DYNAMIC_DOWNCAST (CBCGPRibbonPanelMenu, pPopupMenu->GetParentPopupMenu ());
			if (pParentMenu != NULL)
			{
				ASSERT_VALID(pParentMenu);
				pParentMenu->m_bForceClose = TRUE;
			}
		}
	}

	if (IsBackStageView())
	{
		if (bInvoked)
		{
			PostMessage(WM_CLOSE);
		}
		return;
	}

	CFrameWnd* pParentFrame = BCGPGetParentFrame (this);
	ASSERT_VALID (pParentFrame);

	pParentFrame->DestroyWindow ();
}
//*************************************************************************************
void CBCGPRibbonPanelMenuBar::OnChangeHighlighted (CBCGPBaseRibbonElement* pHot)
{
	ASSERT_VALID (this);

	if (m_pPanel == NULL || !m_pPanel->m_bMenuMode)
	{
		return;
	}

	CBCGPRibbonButton* pDroppedDown = DYNAMIC_DOWNCAST (
		CBCGPRibbonButton, m_pPanel->GetDroppedDown ());

	CBCGPRibbonButton* pHotButton = DYNAMIC_DOWNCAST (
		CBCGPRibbonButton, pHot);

	if (pDroppedDown != NULL && pHot == NULL)
	{
		return;
	}

	BOOL bHotWasChanged = pDroppedDown != pHot;

	if (pHotButton != NULL && pDroppedDown == pHotButton &&
		!pHotButton->GetMenuRect ().IsRectEmpty () &&
		!pHotButton->IsMenuAreaHighlighted ())
	{
		//------------------------------------------------
		// Mouse moved away from the menu area, hide menu:
		//------------------------------------------------
		bHotWasChanged = TRUE;
	}

	if (bHotWasChanged)
	{
		CBCGPRibbonPanelMenu* pParentMenu = 
			DYNAMIC_DOWNCAST (CBCGPRibbonPanelMenu, GetParent ());

		if (pDroppedDown != NULL)
		{
			const MSG* pMsg = GetCurrentMessage ();

			if (CBCGPToolBar::IsCustomizeMode () ||
				(pMsg != NULL && pMsg->message == WM_KEYDOWN))
			{
				KillTimer (uiRemovePopupTimerEvent);
				m_pDelayedCloseButton = NULL;

				pDroppedDown->ClosePopupMenu ();

				if (pParentMenu != NULL)
				{
					CBCGPPopupMenu::ActivatePopupMenu (
						BCGCBProGetTopLevelFrame (this), pParentMenu);
				}
			}
			else
			{
				m_pDelayedCloseButton = pDroppedDown;
				m_pDelayedCloseButton->m_bToBeClosed = TRUE;

#pragma warning (disable : 4296)
				SetTimer(uiRemovePopupTimerEvent, max (0, m_uiPopupTimerDelay - 1), NULL);
#pragma warning (default : 4296)

				pDroppedDown->Redraw();
			}
		}

		if (pHotButton != NULL && pHotButton->HasMenu ())
		{
			StopWaitForPopup();

			if ((m_pDelayedButton = pHotButton) != NULL)
			{
				if (m_pDelayedButton == m_pDelayedCloseButton)
				{
					BOOL bRestoreSubMenu = TRUE;
					
					CRect rectMenu = m_pDelayedButton->GetMenuRect ();

					if (!rectMenu.IsRectEmpty ())
					{
						CPoint point;
						
						::GetCursorPos (&point);
						ScreenToClient (&point);

						if (!rectMenu.PtInRect (point))
						{
							bRestoreSubMenu = FALSE;
						}
					}

					if (bRestoreSubMenu)
					{
						RestoreDelayedSubMenu ();
						m_pDelayedButton = NULL;
					}
				}
				else
				{
					SetTimer (uiPopupTimerEvent, m_uiPopupTimerDelay, NULL);
				}
			}
		}

		// Maybe, this menu will be closed by the parent menu bar timer proc.?
		CBCGPRibbonPanelMenuBar* pParentBar = NULL;

		if (pParentMenu != NULL &&
			(pParentBar = pParentMenu->GetParentRibbonMenuBar ()) != NULL &&
			pParentBar->m_pDelayedCloseButton == pParentMenu->GetParentRibbonElement ())
		{
			pParentBar->RestoreDelayedSubMenu ();
		}

		if (pParentMenu != NULL &&
			pParentMenu->GetParentRibbonElement () != NULL)
		{
			ASSERT_VALID (pParentMenu->GetParentRibbonElement ());
			pParentMenu->GetParentRibbonElement ()->OnChangeMenuHighlight (this, pHotButton);
		}
	}
	else if (pHotButton != NULL && pHotButton == m_pDelayedCloseButton)
	{
		m_pDelayedCloseButton->m_bToBeClosed = FALSE;
		m_pDelayedCloseButton = NULL;

		KillTimer (uiRemovePopupTimerEvent);
	}

	if (pHot == NULL)
	{
		CBCGPRibbonPanelMenu* pParentMenu = 
			DYNAMIC_DOWNCAST (CBCGPRibbonPanelMenu, GetParent ());

		if (pParentMenu != NULL &&
			pParentMenu->GetParentRibbonElement () != NULL)
		{
			ASSERT_VALID (pParentMenu->GetParentRibbonElement ());
			pParentMenu->GetParentRibbonElement ()->OnChangeMenuHighlight (this, NULL);
		}
	}
}
//*************************************************************************************
void CBCGPRibbonPanelMenuBar::OnUpdateCmdUI(CFrameWnd* pTarget, BOOL bDisableIfNoHndler)
{
	ASSERT_VALID (this);

	CBCGPRibbonCmdUI state;
	state.m_pOther = this;

	if (m_pCategory != NULL)
	{
		m_pCategory->OnUpdateCmdUI (&state, pTarget, bDisableIfNoHndler);
	}
	else if (m_pPanel != NULL)
	{
		m_pPanel->OnUpdateCmdUI (&state, pTarget, bDisableIfNoHndler);
	}

	// update the dialog controls added to the ribbon
	UpdateDialogControls(pTarget, bDisableIfNoHndler);

	if (bDisableIfNoHndler && m_bSetKeyTips && m_pRibbonBar != NULL)
	{
		if (m_pPanel != NULL)
		{
			if (m_pPanel->GetDroppedDown () == NULL)
			{
				m_pRibbonBar->SetKeyboardNavigationLevel (m_pPanel, FALSE);
			}
		}
		else if (m_pCategory != NULL)
		{
			m_pRibbonBar->SetKeyboardNavigationLevel (m_pCategory, FALSE);
		}

		m_bSetKeyTips = FALSE;
		CBCGPPopupMenu::UpdateAllShadows ();
	}
}
//*************************************************************************************
void CBCGPRibbonPanelMenuBar::OnDrawMenuBorder (CDC* pDC)
{
	ASSERT_VALID (this);

	if (m_pCategory != NULL)
	{
		m_pCategory->OnDrawMenuBorder (pDC, this);
	}
	else if (m_pPanel != NULL)
	{
		m_pPanel->OnDrawMenuBorder (pDC, this);
	}
}
//*************************************************************************************
int CBCGPRibbonPanelMenuBar::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CBCGPPopupMenuBar::OnCreate(lpCreateStruct) == -1)
		return -1;

	if (globalData.m_bShowTooltipsOnRibbonFloaty || !IsFloaty () || IsQATPopup ())
	{
		CBCGPTooltipManager::CreateToolTip (m_pToolTip, this,
			BCGP_TOOLTIP_TYPE_RIBBON);

		if (m_pToolTip->GetSafeHwnd () != NULL)
		{
			CRect rectClient;
			GetClientRect (&rectClient);

			m_pToolTip->SetMaxTipWidth (640);
			m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectClient, GetDlgCtrlID ());
		}
	}

	if (m_pPanel != NULL)
	{
		ASSERT_VALID(m_pPanel);

		if (m_pPanel->m_pPaletteButton != NULL || IsBackStageView())
		{
			m_wndScrollBarVert.Create (WS_CHILD | WS_VISIBLE | SBS_VERT,
				CRect (0, 0, 0, 0), this, nScrollBarID);
			m_pPanel->m_pScrollBar = &m_wndScrollBarVert;
		}

		if (IsBackStageView())
		{
			m_wndScrollBarHorz.Create (WS_CHILD | WS_VISIBLE | SBS_HORZ,
				CRect (0, 0, 0, 0), this, nScrollBarHorzID);
			m_pPanel->m_pScrollBarHorz = &m_wndScrollBarHorz;
		}
	}

	if (m_pRibbonBar != NULL && m_pRibbonBar->IsDialogMode())
	{
		CBCGPPopupMenu* pParentMenu = DYNAMIC_DOWNCAST (CBCGPPopupMenu, GetParent());
		if (pParentMenu != NULL)
		{
			pParentMenu->m_pMessageWnd = m_pRibbonBar->GetOwner();
		}
	}

	if (m_pRibbonBar != NULL && m_pRibbonBar->GetKeyboardNavigationLevel () >= 0)
	{
		m_bSetKeyTips = TRUE;
		m_bHasKeyTips = TRUE;
	}

	::GetCursorPos (&m_ptStartMenu);

	CBCGPGestureConfig gestureConfig;
	gestureConfig.EnablePan(TRUE, BCGP_GC_PAN_WITH_SINGLE_FINGER_VERTICALLY | BCGP_GC_PAN_WITH_SINGLE_FINGER_HORIZONTALLY | BCGP_GC_PAN_WITH_GUTTER | BCGP_GC_PAN_WITH_INERTIA);
	
	bcgpGestureManager.SetGestureConfig(GetSafeHwnd(), gestureConfig);
	return 0;
}
//*************************************************************************************
void CBCGPRibbonPanelMenuBar::OnDestroy() 
{
	KillTimer(uiAccNotifyEvent);

	if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
	{
		m_pRibbonBar->SetAutoHidePopup(NULL);
	}

	// Inform all elements that parent menu bar is going to be destroyed:
	CArray<CBCGPBaseRibbonElement*,CBCGPBaseRibbonElement*> arElems;

	if (m_pCategory != NULL)
	{
		ASSERT_VALID(m_pCategory);

		m_pCategory->GetElements(arElems);
	}
	else if (m_pPanel != NULL)
	{
		ASSERT_VALID(m_pPanel);
		m_pPanel->GetElements(arElems);
	}
	
	for (int i = 0; i < arElems.GetSize(); i++)
	{
		CBCGPBaseRibbonElement* pElem = arElems[i];
		ASSERT_VALID(pElem);

		pElem->OnBeforeDestroyParentMenuBar(this);
	}

	if (m_pToolTip != NULL)
	{
		CBCGPTooltipManager::DeleteToolTip (m_pToolTip);
	}

	CBCGPRibbonBar*	pTopRibbon = GetTopLevelRibbonBar ();
	if (pTopRibbon != NULL)
	{
		ASSERT_VALID(pTopRibbon);
		pTopRibbon->SetContextHelpActiveID(NULL);
	}

	CBCGPPopupMenuBar::OnDestroy();
}
//************************************************************************************
LRESULT CBCGPRibbonPanelMenuBar::OnGestureEvent(WPARAM wp, LPARAM lp)
{
	if (ProcessGestureEvent(this, wp, lp))
	{
		return Default();
	}
	
	return 0;
}
//************************************************************************************
LRESULT CBCGPRibbonPanelMenuBar::OnTabletQuerySystemGestureStatus(WPARAM /*wp*/, LPARAM /*lp*/)
{
	return (LRESULT)TABLET_DISABLE_PRESSANDHOLD;
}
//*************************************************************************************
void CBCGPRibbonPanelMenuBar::OnSize(UINT nType, int cx, int cy) 
{
	CBCGPPopupMenuBar::OnSize(nType, cx, cy);
	
	if (m_pToolTip->GetSafeHwnd () != NULL && cx != 0 && cy != 0)
	{
		m_pToolTip->SetToolRect (this, GetDlgCtrlID (), CRect (0, 0, cx, cy));
	}
}
//**********************************************************************************
BOOL CBCGPRibbonPanelMenuBar::OnNeedTipText(UINT /*id*/, NMHDR* pNMH, LRESULT* /*pResult*/)
{
	static CString strTipText;

	if (m_pToolTip->GetSafeHwnd () == NULL || pNMH->hwndFrom != m_pToolTip->GetSafeHwnd ())
	{
		return FALSE;
	}

	BOOL bIsHwnd = FALSE;

	LPNMTTDISPINFO pTTDispInfo = (LPNMTTDISPINFO) pNMH;
	if ((pTTDispInfo->uFlags & TTF_IDISHWND) == TTF_IDISHWND)
	{
		CWnd* pWndTT = CWnd::FromHandle((HWND)pNMH->idFrom);
		if (pWndTT->GetSafeHwnd() != NULL)
		{
			bIsHwnd = TRUE;

			if (!pWndTT->IsWindowEnabled())
			{
				return FALSE;
			}
		}
	}
	
	CBCGPPopupMenu* pPopupMenu = CBCGPPopupMenu::GetSafeActivePopupMenu();
	if (pPopupMenu->GetSafeHwnd() != NULL)
	{
		if (pPopupMenu != GetParent() && pPopupMenu->m_hwndConnectedFloaty != GetParent()->GetSafeHwnd())
		{
			return bIsHwnd;
		}
	}

	CBCGPRibbonBar*	pTopRibbon = GetTopLevelRibbonBar ();
	if (pTopRibbon != NULL && !pTopRibbon->IsToolTipEnabled ())
	{
		return TRUE;
	}

	CPoint point;
	
	::GetCursorPos (&point);
	ScreenToClient (&point);

	CBCGPBaseRibbonElement* pHit = HitTest (point);

	if (pTopRibbon != NULL)
	{
		ASSERT_VALID(pTopRibbon);
		pTopRibbon->SetContextHelpActiveID(pHit);
	}

	if (pHit == NULL)
	{
		return TRUE;
	}

	ASSERT_VALID (pHit);

	if (pHit->HasMenu () && IsMainPanel ())
	{
		return TRUE;
	}

	if (pTopRibbon == NULL || !pTopRibbon->OnGetCustomToolTip(pHit, strTipText))
	{
		strTipText = pHit->GetToolTipText();
	}

	if (strTipText.IsEmpty ())
	{
		return TRUE;
	}

	if (pTopRibbon != NULL)
	{
		ASSERT_VALID(pTopRibbon);

		CPoint ptToolTipLocation(-1, -1);

		if (!m_bIsMenuMode && !IsMainPanel () && !m_bIsCtrlMode && !pHit->IsQATMode() && !pHit->IsTabElement() && !pHit->IsCaptionButton())
		{
			CRect rectWindow;
			GetWindowRect (rectWindow);
			
			CRect rectElem = pHit->GetRect ();
			ClientToScreen (&rectElem);
			
			ptToolTipLocation = CPoint(rectElem.left, rectWindow.bottom);
		}

		pTopRibbon->SetupButtonToolTip(m_pToolTip, pHit, ptToolTipLocation);
	}

	if (m_bHasKeyTips)
	{
		m_pToolTip->SetWindowPos (&wndTopMost, -1, -1, -1, -1,
			SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
	}

	pTTDispInfo->lpszText = const_cast<LPTSTR> ((LPCTSTR) strTipText);
	return TRUE;
}
//**************************************************************************
LRESULT CBCGPRibbonPanelMenuBar::OnBCGUpdateToolTips (WPARAM wp, LPARAM)
{
	UINT nTypes = (UINT) wp;

	if ((nTypes & BCGP_TOOLTIP_TYPE_RIBBON) && (globalData.m_bShowTooltipsOnRibbonFloaty || !IsFloaty () || IsQATPopup ()))
	{
		CBCGPTooltipManager::CreateToolTip (m_pToolTip, this,
			BCGP_TOOLTIP_TYPE_RIBBON);

		CRect rectClient;
		GetClientRect (&rectClient);

		m_pToolTip->SetMaxTipWidth (640);
		
		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectClient, GetDlgCtrlID ());
	}

	return 0;
}
//**************************************************************************
void CBCGPRibbonPanelMenuBar::PopTooltip ()
{
	ASSERT_VALID (this);

	if (m_pToolTip->GetSafeHwnd () != NULL)
	{
		m_pToolTip->Pop ();
	}
}
//**************************************************************************
void CBCGPRibbonPanelMenuBar::SetActive (BOOL bIsActive)
{
	ASSERT_VALID (this);

	CBCGPRibbonPanelMenu* pParentMenu = DYNAMIC_DOWNCAST (
		CBCGPRibbonPanelMenu, GetParent ());
	if (pParentMenu != NULL)
	{
		ASSERT_VALID (pParentMenu);
		pParentMenu->SetActive (bIsActive);
	}
}
//**************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonPanelMenuBar::FindByOrigin (CBCGPBaseRibbonElement* pOrigin) const
{
	ASSERT_VALID (this);
	ASSERT_VALID (pOrigin);

	if (m_pPanel == NULL)
	{
		return NULL;
	}

	ASSERT_VALID (m_pPanel);

	CArray<CBCGPBaseRibbonElement*,CBCGPBaseRibbonElement*> arElems;
	m_pPanel->GetElements (arElems);

	for (int i = 0; i < arElems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pListElem = arElems [i];
		ASSERT_VALID (pListElem);

		CBCGPBaseRibbonElement* pElem = pListElem->FindByOriginal (pOrigin);

		if (pElem != NULL)
		{
			ASSERT_VALID (pElem);
			return pElem;
		}
	}

	return NULL;
}
//**************************************************************************
CBCGPRibbonBar*	CBCGPRibbonPanelMenuBar::GetTopLevelRibbonBar () const
{
	ASSERT_VALID (this);

	if (m_pRibbonBar != NULL)
	{
		ASSERT_VALID (m_pRibbonBar);
		return m_pRibbonBar;
	}
	else if (m_pPanelOrigin != NULL)
	{
		ASSERT_VALID (m_pPanelOrigin);
		ASSERT_VALID (m_pPanelOrigin->m_pParent);

		return m_pPanelOrigin->m_pParent->GetParentRibbonBar ();
	}

	return NULL;
}
//**************************************************************************
void CBCGPRibbonPanelMenuBar::OnTimer(UINT_PTR nIDEvent)
{
	CPoint ptCursor;
	::GetCursorPos (&ptCursor);
	ScreenToClient (&ptCursor);

	if (nIDEvent == uiPopupTimerEvent)
	{
		KillTimer (uiPopupTimerEvent);

		//---------------------------------
		// Remove current tooltip (if any):
		//---------------------------------
		if (m_pToolTip->GetSafeHwnd () != NULL)
		{
			m_pToolTip->ShowWindow (SW_HIDE);
		}

		if (m_pDelayedCloseButton != NULL &&
			m_pDelayedCloseButton->GetRect ().PtInRect (ptCursor))
		{
			return;
		}

		CloseDelayedSubMenu ();

		CBCGPRibbonButton* pDelayedPopupMenuButton = m_pDelayedButton;
		m_pDelayedButton = NULL;

		if (pDelayedPopupMenuButton != NULL &&
			pDelayedPopupMenuButton->IsHighlighted ())
		{
			if (pDelayedPopupMenuButton->IsPopupDialogEnabled())
			{
				ASSERT(FALSE);
			}
			else
			{
				pDelayedPopupMenuButton->OnShowPopupMenu ();
			}
		}
	}
	else if (nIDEvent == uiRemovePopupTimerEvent)
	{
		KillTimer (uiRemovePopupTimerEvent);

		if (m_pDelayedCloseButton != NULL)
		{
			ASSERT_VALID (m_pDelayedCloseButton);
			CBCGPPopupMenu* pParentMenu = DYNAMIC_DOWNCAST (CBCGPPopupMenu, GetParent ());

			CRect rectMenu = m_pDelayedCloseButton->GetRect ();

			if (rectMenu.PtInRect (ptCursor))
			{
				return;
			}

			m_pDelayedCloseButton->ClosePopupMenu ();
			m_pDelayedCloseButton = NULL;

			if (pParentMenu != NULL)
			{
				CBCGPPopupMenu::ActivatePopupMenu (BCGCBProGetTopLevelFrame (this), pParentMenu);
			}
		}
	}
	else if (nIDEvent == uiAccNotifyEvent)
	{
		KillTimer (uiAccNotifyEvent);

		CRect rc;
		GetClientRect (&rc);
		if (!rc.PtInRect (ptCursor))
		{
			return;
		}

		CPoint point;
		::GetCursorPos (&point);
		ScreenToClient (&point);

		CBCGPBaseRibbonElement* pHit = HitTest (point);

		if (pHit != NULL)
		{
			CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arButtons;

			if (m_pPanel != NULL)
			{
				m_pPanel->GetVisibleElements (arButtons);
			}
			else if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
			{
				m_pRibbonBar->GetVisibleElements(arButtons);
			}

			int nIndex = -1;

			for (int i = 0; i < arButtons.GetSize(); i++)
			{
				if (arButtons[i] == pHit)
				{
					nIndex = i;
					break;
				}
			}

			if (nIndex != -1)
			{	
				ASSERT_VALID (pHit);
				pHit->SetACCData (this, m_AccData);
				globalData.NotifyWinEvent (EVENT_OBJECT_FOCUS, GetSafeHwnd (), 
					OBJID_CLIENT , nIndex+1);
			}
		}
	}
	else if (nIDEvent == IdAutoCommand)
	{
		if (!m_rectAutoCommand.PtInRect (ptCursor))
		{
			m_pPressed = NULL;
			KillTimer (IdAutoCommand);
			m_rectAutoCommand.SetRectEmpty ();
		}
		else
		{
			if (m_pPressed != NULL)
			{
				ASSERT_VALID (m_pPressed);

				if (m_pPressed->GetRect ().PtInRect (ptCursor))
				{
					if (!m_pPressed->OnAutoRepeat ())
					{
						KillTimer (IdAutoCommand);
					}
				}
			}
		}
	}

}
//*******************************************************************************
void CBCGPRibbonPanelMenuBar::StopWaitForPopup()
{
	ASSERT_VALID (this);

	if (m_pDelayedButton != NULL)
	{
		KillTimer (uiPopupTimerEvent);
	}
}
//*******************************************************************************
void CBCGPRibbonPanelMenuBar::CloseDelayedSubMenu ()
{
	ASSERT_VALID (this);

	if (m_pDelayedCloseButton != NULL)
	{
		ASSERT_VALID (m_pDelayedCloseButton);

		KillTimer (uiRemovePopupTimerEvent);

		m_pDelayedCloseButton->ClosePopupMenu ();
		m_pDelayedCloseButton = NULL;
	}
}
//*******************************************************************************
void CBCGPRibbonPanelMenuBar::RestoreDelayedSubMenu ()
{
	ASSERT_VALID (this);

	if (m_pDelayedCloseButton == NULL || m_pPanel == NULL)
	{
		return;
	}

	ASSERT_VALID (m_pDelayedCloseButton);
	m_pDelayedCloseButton->m_bToBeClosed = FALSE;

	CBCGPBaseRibbonElement* pPrev = m_pPanel->GetHighlighted ();

	m_pPanel->Highlight (TRUE, m_pDelayedCloseButton->GetRect ().TopLeft ());

	BOOL bUpdate = FALSE;

	if (m_pDelayedCloseButton != pPrev)
	{
		if (m_pDelayedCloseButton != NULL)
		{
			ASSERT_VALID (m_pDelayedCloseButton);
			InvalidateRect (m_pDelayedCloseButton->GetRect ());
		}

		if (pPrev != NULL)
		{
			ASSERT_VALID (pPrev);
			InvalidateRect (pPrev->GetRect ());
		}

		bUpdate = TRUE;
	}

	m_pDelayedCloseButton = NULL;

	KillTimer (uiRemovePopupTimerEvent);

	if (bUpdate)
	{
		UpdateWindow ();
	}
}
//****************************************************************************************
BOOL CBCGPRibbonPanelMenuBar::OnKey (UINT nChar)
{
	ASSERT_VALID (this);

	if (nChar == VK_F10 && (0x8000 & GetKeyState(VK_SHIFT)) != 0 ||
		nChar == VK_APPS)
	{
		OnContextMenu(this, CPoint (-1, -1));
		return TRUE;
	}

	if (m_pRibbonBar != NULL && m_pRibbonBar->ProcessKey (nChar))
	{
		return TRUE;
	}

	if (m_pPanel != NULL)
	{
		ASSERT_VALID (m_pPanel);

		CBCGPDisableMenuAnimation disableMenuAnimation;
		return m_pPanel->OnKey (nChar);
	}

	if (m_pCategory != NULL)
	{
		ASSERT_VALID (m_pCategory);

		CBCGPBaseRibbonElement* pDroppedDown = m_pCategory->GetDroppedDown();
		if (pDroppedDown != NULL)
		{
			if (nChar == VK_RIGHT)
			{
				return FALSE;
			}

			if (nChar == VK_TAB)
			{
				if (pDroppedDown->IsCloseDropDownByTabKey())
				{
					pDroppedDown->ClosePopupMenu();
				}
				else
				{
					return FALSE;
				}
			}
		}

		CBCGPDisableMenuAnimation disableMenuAnimation;
		return m_pCategory->OnKey (nChar);
	}

	return FALSE;
}
//*****************************************************************************
void CBCGPRibbonPanelMenuBar::OnContextMenu(CWnd* /*pWnd*/, CPoint point) 
{
	ASSERT_VALID (this);

	if (m_pRibbonBar == NULL || IsBackStageView())
	{
		return;
	}

	if (m_pPanel != NULL && m_pPanel->m_pPaletteButton != NULL)
	{
		ASSERT_VALID(m_pPanel->m_pPaletteButton);

		if (m_pPanel->m_pPaletteButton->m_pParentControl != NULL)
		{
			return;
		}
	}

	ASSERT_VALID (m_pRibbonBar);

	CBCGPPopupMenu* pParentMenu = DYNAMIC_DOWNCAST (CBCGPPopupMenu, GetParent ());

	if (pParentMenu != NULL &&
		pParentMenu->GetParentRibbonElement () != NULL)
	{
		ASSERT_VALID (pParentMenu->GetParentRibbonElement ());

		if (pParentMenu->GetParentRibbonElement ()->m_bFloatyMode || pParentMenu->GetParentRibbonElement ()->IsCaptionButton())
		{
			return;
		}
	}

	if (m_bAutoCommandTimer)
	{
		KillTimer (IdAutoCommand);
		m_bAutoCommandTimer = FALSE;
		m_pPressed = NULL;
		m_rectAutoCommand.SetRectEmpty ();
	}

	if (IsRibbonPanel () && m_bCustomizeMenu)
	{
		if (IsFloaty () && !IsQATPopup ())
		{
			return;
		}

		if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)	// Left mouse button is pressed
		{
			return;
		}

		CPoint ptClient = point;
		ScreenToClient (&ptClient);

		CBCGPRibbonButton* pDroppedDown = GetDroppedDown ();
		if (pDroppedDown != NULL)
		{
			pDroppedDown->ClosePopupMenu ();
		}

		StopWaitForPopup();

		if (point == CPoint (-1, -1))
		{
			CBCGPBaseRibbonElement* pFocused = GetFocused ();
			if (pFocused != NULL)
			{
				ASSERT_VALID (pFocused);

				CRect rectFocus = pFocused->GetRect ();
				ClientToScreen (&rectFocus);

				m_pRibbonBar->OnShowRibbonContextMenu (this, rectFocus.left, rectFocus.top, pFocused);
				return;
			}
		}

		m_pRibbonBar->OnShowRibbonContextMenu (this, point.x, point.y, HitTest (ptClient));
	}
}
//**************************************************************************
void CBCGPRibbonPanelMenuBar::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	ASSERT_VALID (this);

	if (m_pPanel == NULL ||
		pScrollBar->GetSafeHwnd () != m_wndScrollBarVert.GetSafeHwnd () ||
		(m_pPanel->m_pPaletteButton == NULL && !IsBackStageView()))
	{
		CBCGPPopupMenuBar::OnVScroll(nSBCode, nPos, pScrollBar);
		return;
	}

	SCROLLINFO scrollInfo;
	ZeroMemory(&scrollInfo, sizeof(SCROLLINFO));

    scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_ALL;

	m_wndScrollBarVert.GetScrollInfo (&scrollInfo);

	int iOffset = m_pPanel->m_nScrollOffset;
	int nMaxOffset = scrollInfo.nMax;
	int nPage = scrollInfo.nPage;

	if (nMaxOffset - nPage <= 1)
	{
		return;
	}

	int nRowHeight = m_pPanel->IsBackstageView() ? globalData.GetTextHeight() : m_pPanel->m_pPaletteButton->GetMenuRowHeight ();

	switch (nSBCode)
	{
	case SB_LINEUP:
		iOffset -= nRowHeight;
		break;

	case SB_LINEDOWN:
		iOffset += nRowHeight;
		break;

	case SB_TOP:
		iOffset = 0;
		break;

	case SB_BOTTOM:
		iOffset = nMaxOffset;
		break;

	case SB_PAGEUP:
		iOffset -= nPage;
		break;

	case SB_PAGEDOWN:
		iOffset += nPage;
		break;

	case SB_THUMBPOSITION:
		iOffset = scrollInfo.nPos;
		break;

	case SB_THUMBTRACK:
		iOffset = scrollInfo.nTrackPos;
		break;

	default:
		return;
	}

	iOffset = min (max (0, iOffset), nMaxOffset - nPage);

	if (iOffset == m_pPanel->m_nScrollOffset)
	{
		return;
	}

	if (m_pPanel->IsBackstageView())
	{
		m_pPanel->m_nScrollOffset = iOffset;
		AdjustLocations();
	}
	else
	{
		m_pPanel->ScrollPalette (iOffset);
	}

	m_wndScrollBarVert.SetScrollPos (iOffset);
	PopTooltip();

	RedrawWindow ();
}
//**********************************************************************************
void CBCGPRibbonPanelMenuBar::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	ASSERT_VALID (this);

	if (m_pPanel == NULL ||
		pScrollBar->GetSafeHwnd () != m_wndScrollBarHorz.GetSafeHwnd () || !IsBackStageView())
	{
		CBCGPPopupMenuBar::OnHScroll(nSBCode, nPos, pScrollBar);
		return;
	}

	SCROLLINFO scrollInfo;
	ZeroMemory(&scrollInfo, sizeof(SCROLLINFO));

    scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_ALL;

	m_wndScrollBarHorz.GetScrollInfo (&scrollInfo);

	int iOffset = m_pPanel->m_nScrollOffsetHorz;
	int nMaxOffset = scrollInfo.nMax;
	int nPage = scrollInfo.nPage;

	if (nMaxOffset - nPage <= 1)
	{
		return;
	}

	int nColumnWidth = globalData.GetTextHeight();

	switch (nSBCode)
	{
	case SB_LINEUP:
		iOffset -= nColumnWidth;
		break;

	case SB_LINEDOWN:
		iOffset += nColumnWidth;
		break;

	case SB_TOP:
		iOffset = 0;
		break;

	case SB_BOTTOM:
		iOffset = nMaxOffset;
		break;

	case SB_PAGEUP:
		iOffset -= nPage;
		break;

	case SB_PAGEDOWN:
		iOffset += nPage;
		break;

	case SB_THUMBPOSITION:
	case SB_THUMBTRACK:
		iOffset = nPos;
		break;

	default:
		return;
	}

	iOffset = min (max (0, iOffset), nMaxOffset - nPage);

	if (iOffset == m_pPanel->m_nScrollOffsetHorz)
	{
		return;
	}

	m_pPanel->m_nScrollOffsetHorz = iOffset;
	AdjustLocations();

	m_wndScrollBarHorz.SetScrollPos (iOffset);
	RedrawWindow ();

	if (IsBackStageView())
	{
		CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar();
		if (pRibbonBar != NULL)
		{
			pRibbonBar->SetBackstageHorzOffset(iOffset);
		}
	}
}
//**********************************************************************************
CBCGPRibbonButton* CBCGPRibbonPanelMenuBar::GetDroppedDown () const
{
	if (m_pCategory != NULL)
	{
		ASSERT_VALID (m_pCategory);

		return DYNAMIC_DOWNCAST (
			CBCGPRibbonButton, m_pCategory->GetDroppedDown ());
	}
	else if (m_pPanel != NULL)
	{
		ASSERT_VALID (m_pPanel);

		return DYNAMIC_DOWNCAST (
			CBCGPRibbonButton, m_pPanel->GetDroppedDown ());
	}
	else if (IsRibbonPopupMode())
	{
		ASSERT_VALID(m_pRibbonBar);
		return DYNAMIC_DOWNCAST(CBCGPRibbonButton, m_pRibbonBar->GetDroppedDown());
	}

	return NULL;
}
//**********************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonPanelMenuBar::HitTest (CPoint point) const
{
	if (m_pCategory != NULL)
	{
		ASSERT_VALID (m_pCategory);
		return m_pCategory->HitTest (point, TRUE);
	}
	else if (m_pPanel != NULL)
	{
		ASSERT_VALID (m_pPanel);
		return m_pPanel->HitTest (point);
	}
	else if (IsRibbonPopupMode())
	{
		ASSERT_VALID (m_pRibbonBar);
		return m_pRibbonBar->HitTest(point, TRUE, TRUE);
	}

	return NULL;
}
//**********************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonPanelMenuBar::GetFocused () const
{
	if (m_pCategory != NULL)
	{
		ASSERT_VALID (m_pCategory);
		return m_pCategory->GetFocused ();
	}
	else if (m_pPanel != NULL)
	{
		ASSERT_VALID (m_pPanel);
		return m_pPanel->GetFocused ();
	}
	else if (IsRibbonPopupMode())
	{
		ASSERT_VALID (m_pRibbonBar);
		return m_pRibbonBar->GetFocused ();
	}

	return NULL;
}
//**********************************************************************************
int CBCGPRibbonPanelMenuBar::HitTestEx (CPoint point) const
{
	if (m_pCategory != NULL)
	{
		ASSERT_VALID (m_pCategory);
		return m_pCategory->HitTestEx (point);
	}
	else if (m_pPanel != NULL)
	{
		ASSERT_VALID (m_pPanel);
		return m_pPanel->HitTestEx (point);
	}
	else if (IsRibbonPopupMode())
	{
		ASSERT_VALID (m_pRibbonBar);
		ASSERT(FALSE);	// TBD
//		return m_pRibbonBar->HitTest(point);
	}

	return NULL;
}
//*************************************************************************************
BOOL CBCGPRibbonPanelMenuBar::IsRibbonPopupMode() const
{
	return m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup->GetSafeHwnd() == GetSafeHwnd();
}
//*************************************************************************************
BOOL CBCGPRibbonPanelMenuBar::OnSetAccData (long lVal)
{
	ASSERT_VALID (this);

	m_AccData.Clear ();

	CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arButtons;

	if (m_pCategory != NULL)
	{
		m_pCategory->GetVisibleElements (arButtons);
	}
	else if (m_pPanel != NULL)
	{
		m_pPanel->GetVisibleElements (arButtons);
	}
	else if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
	{
		m_pRibbonBar->GetVisibleElements (arButtons);
	}
	
	int nIndex = (int)lVal - 1;
	if (nIndex < arButtons.GetSize () && nIndex >= 0)
	{
		CBCGPBaseRibbonElement* pElem = arButtons [nIndex];
		ASSERT_VALID (pElem);
		
		BOOL retVal = pElem->SetACCData(this, m_AccData);
		return retVal;
	}

	return FALSE;
}
//****************************************************************************
HRESULT CBCGPRibbonPanelMenuBar::get_accChildCount(long *pcountChildren)
{
	if (!pcountChildren)
    {
        return E_INVALIDARG;
    }

	CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arButtons;

	if (m_pCategory != NULL)
	{
		m_pCategory->GetVisibleElements(arButtons);
	}
	else if (m_pPanel != NULL)
	{
		m_pPanel->GetVisibleElements(arButtons);
	}
	else if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
	{
		m_pRibbonBar->GetVisibleElements(arButtons);
	}
	
	*pcountChildren = (int)arButtons.GetSize();
	return S_OK;  
}
//****************************************************************************
HRESULT CBCGPRibbonPanelMenuBar::accHitTest(long xLeft, long yTop, VARIANT *pvarChild)
{
	if( !pvarChild )
	{
		return E_INVALIDARG;
	}

	pvarChild->vt = VT_I4;
	pvarChild->lVal = CHILDID_SELF;

	CPoint pt (xLeft, yTop);
	ScreenToClient (&pt);

	CBCGPBaseRibbonElement* pElement = HitTest (pt);

	if (pElement != NULL)
	{
		CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arButtons;

		if (m_pCategory != NULL)
		{
			m_pCategory->GetVisibleElements(arButtons);
		}
		else if (m_pPanel != NULL)
		{
			m_pPanel->GetVisibleElements(arButtons);
		}
		else if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
		{
			m_pRibbonBar->GetVisibleElements(arButtons);
		}

		int index = 0;

		for (int i = 0; i < arButtons.GetSize(); i++)
		{
			if (arButtons[i] == pElement)
			{
				index = i;
				break;
			}
		}
		pvarChild->lVal = index + 1;
		pElement->SetACCData(this, m_AccData);
	}

	return S_OK;
}
//**************************************************************************
BOOL CBCGPRibbonPanelMenuBar::PreTranslateMessage(MSG* pMsg) 
{
	if (pMsg->message == WM_KEYDOWN)
	{
		if (m_bIsCtrlMode)
		{
			switch (pMsg->wParam)
			{
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
				if (OnKey ((UINT)pMsg->wParam))
				{
					return TRUE;
				}
			}
		}
		else
		{
			if (pMsg->wParam == VK_TAB && OnKey (VK_TAB))
			{
				return TRUE;
			}
		}
	}

	if (pMsg->message == WM_LBUTTONDOWN)
	{
		CBCGPRibbonEditCtrl* pEdit = DYNAMIC_DOWNCAST (CBCGPRibbonEditCtrl, GetFocus ());
		if (pEdit != NULL)
		{
			ASSERT_VALID (pEdit);

			if (pEdit->GetOwnerRibbonEdit().IsCommandsCombo())
			{
				if (m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup == this)
				{
				}
				else
				{
					return TRUE;
				}
			}

			CPoint point;
			
			::GetCursorPos (&point);
			ScreenToClient (&point);

			pEdit->GetOwnerRibbonEdit ().PreLMouseDown (point);
		}
	}

	return CBCGPPopupMenuBar::PreTranslateMessage(pMsg);
}
//**************************************************************************
void CBCGPRibbonPanelMenuBar::OnSetFocus(CWnd* pOldWnd) 
{
	if (m_bIsCtrlMode)
	{
		return;
	}

	CBCGPPopupMenuBar::OnSetFocus(pOldWnd);
}
//***************************************************************
BOOL CBCGPRibbonPanelMenuBar::OnGestureEventPan(const CPoint& ptFrom, const CPoint& ptTo, CSize& sizeOverPan)
{
	if (CBCGPPopupMenu::GetActiveMenu() != NULL && CBCGPPopupMenu::GetActiveMenu() != GetParent())
	{
		return FALSE;
	}

	SendMessageToDescendants(WM_CANCELMODE);

	CPoint ptFromScreen = ptFrom;
	ClientToScreen(&ptFromScreen);
	
	CPoint ptToScreen = ptTo;
	ClientToScreen(&ptToScreen);

	if (ptFrom.x != ptTo.x && m_wndScrollBarHorz.GetSafeHwnd() != NULL && m_wndScrollBarHorz.IsWindowVisible())
	{
		if (m_wndScrollBarHorz.DoGesturePan(ptFromScreen, ptToScreen))
		{
			return FALSE;
		}

		int nDelta = ptFrom.x - ptTo.x;
		int nPos = m_wndScrollBarHorz.GetScrollPos();

		OnHScroll(SB_THUMBTRACK, nPos + nDelta, &m_wndScrollBarHorz);

		if (m_wndScrollBarHorz.GetScrollPos() == nPos)
		{
			const int nColumnWidth = globalData.GetTextHeight();
			sizeOverPan.cx = ptFrom.x > ptTo.x ? -nColumnWidth : nColumnWidth;

			return FALSE;
		}

		return TRUE;
	}
	
	if (ptFrom.y != ptTo.y && m_wndScrollBarVert.GetSafeHwnd() != NULL && m_wndScrollBarVert.IsWindowVisible())
	{
		if (m_wndScrollBarVert.DoGesturePan(ptFromScreen, ptToScreen))
		{
			return FALSE;
		}

		int nPos = m_wndScrollBarVert.GetScrollPos();

		if (ptFrom.y < ptTo.y)
		{
			OnVScroll(SB_LINEUP, 0, &m_wndScrollBarVert);
		}
		else
		{
			OnVScroll(SB_LINEDOWN, 0, &m_wndScrollBarVert);
		}

		if (m_wndScrollBarVert.GetScrollPos() == nPos)
		{
			const int nRowHeight = globalData.GetTextHeight();
			sizeOverPan.cy = ptFrom.y > ptTo.y ? -nRowHeight : nRowHeight;

			return FALSE;
		}
		
		return TRUE;
	}

	return FALSE;
}
//***************************************************************
BOOL CBCGPRibbonPanelMenuBar::IsStopMouseClickAfterCloseActiveMenu() const
{
	return m_pRibbonBar != NULL && m_pRibbonBar->m_pWndAutoHidePopup->GetSafeHwnd() == GetSafeHwnd();
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonPanelMenu

IMPLEMENT_DYNAMIC(CBCGPRibbonPanelMenu, CBCGPPopupMenu)

CBCGPRibbonPanelMenu::CBCGPRibbonPanelMenu (CBCGPRibbonPanel* pPanel) :
	m_wndRibbonBar (pPanel)
{
	m_bForceClose = FALSE;
	m_bInScroll = FALSE;
}

CBCGPRibbonPanelMenu::CBCGPRibbonPanelMenu (
		CBCGPRibbonBar* pRibbonBar,
		const CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>&	arButtons, BOOL bIsFloatyMode) :
	m_wndRibbonBar (pRibbonBar, arButtons, bIsFloatyMode)
{
	m_bForceClose = FALSE;
	m_bInScroll = FALSE;
}

CBCGPRibbonPanelMenu::CBCGPRibbonPanelMenu (
		CBCGPRibbonPaletteButton* pPaletteButton) :
	m_wndRibbonBar (pPaletteButton)
{
	ASSERT_VALID (pPaletteButton);

	m_bScrollable = TRUE;
	m_bForceClose = FALSE;
	m_bInScroll = FALSE;

	if (pPaletteButton->IsMenuResizeEnabled ())
	{
		ASSERT_VALID (m_wndRibbonBar.m_pPanel);

		CSize sizeMin = m_wndRibbonBar.m_pPanel->GetPaletteMinSize ();

		if (sizeMin.cx > 0 && sizeMin.cy > 0)
		{
			CSize sizeBorder = GetBorderSize ();

			sizeMin.cx += sizeBorder.cx * 2;
			sizeMin.cy += sizeBorder.cy * 2;
			
			if (pPaletteButton->IsMenuResizeVertical ())
			{
				EnableVertResize (sizeMin.cy);
			}
			else
			{
				EnableResize (sizeMin);
			}
		}
	}
}

CBCGPRibbonPanelMenu::CBCGPRibbonPanelMenu (
		CBCGPRibbonCategory* pCategory, CSize size) :
	m_wndRibbonBar (pCategory, size)
{
	m_bForceClose = FALSE;
	m_bInScroll = FALSE;
	m_bNormalizeLocationInScreen = FALSE;
}

CBCGPRibbonPanelMenu::CBCGPRibbonPanelMenu ()
{
	m_bForceClose = FALSE;
	m_bInScroll = FALSE;
}

CBCGPRibbonPanelMenu::CBCGPRibbonPanelMenu(CBCGPRibbonBar* pRibbonBar) :
	m_wndRibbonBar(pRibbonBar)
{
	m_bForceClose = FALSE;
	m_bInScroll = FALSE;
	m_bNormalizeLocationInScreen = FALSE;
}

CBCGPRibbonPanelMenu::~CBCGPRibbonPanelMenu()
{
}

BEGIN_MESSAGE_MAP(CBCGPRibbonPanelMenu, CBCGPPopupMenu)
	//{{AFX_MSG_MAP(CBCGPRibbonPanelMenu)
	ON_WM_KEYDOWN()
	ON_WM_MOUSEWHEEL()
	ON_WM_LBUTTONDOWN()
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

CBCGPRibbonPanelMenuBar* CBCGPRibbonPanelMenu::GetParentRibbonMenuBar () const
{
	ASSERT_VALID (this);

	CBCGPPopupMenu* pParentMenu = GetParentPopupMenu ();
	if (pParentMenu == NULL)
	{
		return NULL;
	}

	ASSERT_VALID (pParentMenu);

	return DYNAMIC_DOWNCAST (CBCGPRibbonPanelMenuBar,
		pParentMenu->GetMenuBar ());
}
//****************************************************************************************
BOOL CBCGPRibbonPanelMenu::OnDrawMenuImage(CDC* pDC, const CBCGPToolbarMenuButton* pMenuButton, const CRect& rectImage)
{
	ASSERT_VALID (this);

	CBCGPRibbonBar* pRibbonBar = m_wndRibbonBar.GetTopLevelRibbonBar();
	if (pRibbonBar != NULL)
	{
		ASSERT_VALID(pRibbonBar);
		return pRibbonBar->DrawMenuImage (pDC, pMenuButton, rectImage);
	}

	return FALSE;
}
//****************************************************************************************
void CBCGPRibbonPanelMenu::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	ASSERT_VALID (this);

	if (!m_wndRibbonBar.OnKey (nChar))
	{
		CBCGPPopupMenu::OnKeyDown (nChar, nRepCnt, nFlags);
	}
}
//**************************************************************************
BOOL CBCGPRibbonPanelMenu::OnMouseWheel(UINT /*nFlags*/, short zDelta, CPoint /*pt*/) 
{
	ASSERT_VALID (this);

	if (m_bInScroll)
	{
		return TRUE;
	}

	m_bInScroll = TRUE;

	const int nSteps = abs(zDelta) / WHEEL_DELTA;
	const int nOldOffset = m_wndRibbonBar.GetOffset();

	for (int i = 0; i < nSteps; i++)
	{
		if (IsScrollUpAvailable () || IsScrollDnAvailable ())
		{
			int iOffset = m_wndRibbonBar.GetOffset ();

			if (zDelta > 0)
			{
				if (IsScrollUpAvailable ())
				{
					m_wndRibbonBar.SetOffset (iOffset - 1);
					AdjustScroll ();
				}
			}
			else
			{
				if (IsScrollDnAvailable ())
				{
					m_wndRibbonBar.SetOffset (iOffset + 1);
					AdjustScroll ();
				}
			}
		}
		else
		{
			m_wndRibbonBar.OnVScroll (zDelta < 0 ? SB_LINEDOWN : SB_LINEUP, 0, 
				&m_wndRibbonBar.m_wndScrollBarVert);
		}
	}

	if (nOldOffset != m_wndRibbonBar.GetOffset())
	{
		m_wndRibbonBar.RedrawWindow();
	}

	m_bInScroll = FALSE;
	return TRUE;
}
//**************************************************************************
BOOL CBCGPRibbonPanelMenu::IsAlwaysClose () const
{
	return m_bForceClose || ((CBCGPRibbonPanelMenu*) this)->m_wndRibbonBar.IsMainPanel ();
}
//**************************************************************************
void CBCGPRibbonPanelMenu::DoPaint (CDC* pDC)
{
	CBCGPPopupMenu::DoPaint (pDC);
	m_wndRibbonBar.OnDrawMenuBorder (pDC);
}
//*****************************************************************************************
void CBCGPRibbonPanelMenu::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CBCGPPopupMenu::OnLButtonDown(nFlags, point);

	if (m_wndRibbonBar.IsMainPanel ())
	{
		ClientToScreen (&point);
		ScreenToClient (&point);

		m_wndRibbonBar.GetPanel ()->MouseButtonDown (point);
	}
}
//**************************************************************************
int CBCGPRibbonPanelMenu::GetBorderSize () const
{
	if (m_wndRibbonBar.IsRibbonPopupMode())
	{
		return 0;
	}

	return IsMenuMode () ? 
		CBCGPPopupMenu::GetBorderSize () : 
		CBCGPVisualManager::GetInstance ()->GetRibbonPopupBorderSize (this);
}
//**************************************************************************
BOOL CBCGPRibbonPanelMenu::IsScrollUpAvailable ()
{
	return m_wndRibbonBar.m_iOffset > 0;
}
//**************************************************************************
BOOL CBCGPRibbonPanelMenu::IsScrollDnAvailable ()
{
	return m_wndRibbonBar.m_pPanel == NULL ||
		m_wndRibbonBar.m_pPanel->m_bScrollDnAvailable;
}
//**************************************************************************
BOOL CBCGPRibbonPanelMenu::IsParentEditFocused()
{
	CBCGPPopupMenu* pParentMenu = GetParentPopupMenu ();
	if (pParentMenu == NULL)
	{
		return CBCGPPopupMenu::IsParentEditFocused();
	}

	return pParentMenu->IsParentEditFocused();
}
//**************************************************************************
BOOL CBCGPRibbonPanelMenu::IsDropListMode()
{
	CBCGPPopupMenu* pParentMenu = GetParentPopupMenu ();
	if (pParentMenu == NULL)
	{
		return CBCGPPopupMenu::IsDropListMode();
	}
	
	return pParentMenu->IsDropListMode();
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonPanelMenu message handlers

void CBCGPRibbonPanelMenu::OnDestroy() 
{
	if (m_bEscClose && m_wndRibbonBar.GetCategory () != NULL &&
		BCGCBProGetTopLevelFrame (&m_wndRibbonBar) != NULL)
	{
		BCGCBProGetTopLevelFrame (&m_wndRibbonBar)->SetFocus ();
	}

	CBCGPPopupMenu::OnDestroy();
	
}

#endif // BCGP_EXCLUDE_RIBBON

