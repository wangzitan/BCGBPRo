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

// ButtonsList.cpp : implementation file
//

#include "stdafx.h"

#include "afxadv.h"
#include <afxpriv.h>
#include "bcgprores.h"
#include "ButtonsList.h"
#include "BCGPToolbarButton.h"
#include "BCGPToolBarImages.h"
#include "BCGGlobals.h"
#include "BCGPVisualManager.h"
#include "trackmouse.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CButtonsList

CButtonsList::CButtonsList()
{
	m_pSelButton = NULL;
	m_pHighlightedButton = NULL;
	m_pImages = NULL;
	m_nButtonsInRow = 0;

	m_iScrollOffset = 0;
	m_iScrollTotal = 0;
	m_iScrollPage = 0;

	m_bEnableDragFromList = FALSE;

	m_bInited = FALSE;
	m_bTrack = FALSE;
}
//**************************************************************************************
CButtonsList::~CButtonsList()
{
}

BEGIN_MESSAGE_MAP(CButtonsList, CButton)
	//{{AFX_MSG_MAP(CButtonsList)
	ON_WM_ERASEBKGND()
	ON_WM_LBUTTONDOWN()
	ON_WM_VSCROLL()
	ON_WM_ENABLE()
	ON_WM_SYSCOLORCHANGE()
	ON_WM_SIZE()
	ON_WM_CTLCOLOR()
	ON_WM_KEYDOWN()
	ON_WM_GETDLGCODE()
	ON_WM_MOUSEWHEEL()
	ON_WM_MOUSEMOVE()
	ON_WM_CANCELMODE()
	ON_WM_KILLFOCUS()
	ON_WM_SETFOCUS()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_MOUSELEAVE, OnMouseLeave)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CButtonsList message handlers

BOOL CButtonsList::OnEraseBkgnd(CDC* pDC) 
{
	CRect rectClient;	// Client area rectangle
	GetClientRect (&rectClient);

	if (globalData.m_bUseVisualManagerInBuiltInDialogs)
	{
		pDC->FillRect(&rectClient, &CBCGPVisualManager::GetInstance()->GetDlgBackBrush(this));
	}
	else
	{
		pDC->FillSolidRect (&rectClient,  IsWindowEnabled () ? globalData.clrWindow : globalData.clrBtnFace);
	}

	return TRUE;
}
//*********************************************************************************************
void CButtonsList::DrawItem (LPDRAWITEMSTRUCT lpDIS)
{
	if (!m_bInited)
	{
		RebuildLocations ();
	}

	CDC* pDC = CDC::FromHandle (lpDIS->hDC);
	CRect rectClient = lpDIS->rcItem;

	if (m_pImages != NULL)
	{
		m_pImages->SetTransparentColor (globalData.clrBtnFace);

		CBCGPDrawState ds;
		if (!m_pImages->PrepareDrawImage (ds))
		{
			return;
		}

		for (POSITION pos = m_Buttons.GetHeadPosition (); pos != NULL;)
		{
			CBCGPToolbarButton* pButton = m_Buttons.GetNext (pos);
			ASSERT_VALID(pButton);

			CRect rect = pButton->Rect ();
			rect.OffsetRect (0, -m_iScrollOffset);

			if (rect.top >= rectClient.bottom)
			{
				break;
			}

			if (rect.bottom > rectClient.top)
			{
				int nSaveStyle = pButton->m_nStyle;
				BOOL bLocked = pButton->m_bLocked;
				BOOL bIsHighlight = FALSE;

				if (!IsWindowEnabled ())
				{
					pButton->m_nStyle |= TBBS_DISABLED;
				}
				else if (pButton == m_pHighlightedButton)
				{
					bIsHighlight = TRUE;
				}
				else if (pButton == m_pSelButton)
				{
					pButton->m_nStyle |= TBBS_CHECKED;
				}

				pButton->m_bLocked = TRUE;
				pButton->OnDraw (pDC, rect, m_pImages, TRUE, FALSE, bIsHighlight);
				pButton->m_nStyle = nSaveStyle;
				pButton->m_bLocked = bLocked;
			}
		}

		m_pImages->EndDrawImage (ds);
	}

	CBCGPVisualManager::GetInstance()->OnDrawControlBorder(pDC, rectClient, this, FALSE);
}
//*********************************************************************************************
void CButtonsList::OnLButtonDown(UINT /*nFlags*/, CPoint point) 
{
	SetFocus ();

	CBCGPToolbarButton* pButton = HitTest (point);
	if (pButton == NULL)
	{
		return;
	}

	SelectButton (pButton);

	if (m_bEnableDragFromList)
	{
		COleDataSource* pSrcItem = new COleDataSource();

		pButton->m_bDragFromCollection = TRUE;
		pButton->PrepareDrag (*pSrcItem);
		pButton->m_bDragFromCollection = TRUE;

		pSrcItem->DoDragDrop ();
		pSrcItem->InternalRelease();
	}
}
//*********************************************************************************************
void CButtonsList::SetImages (CBCGPToolBarImages* pImages)
{
	ASSERT_VALID (pImages);
	m_pImages = pImages;

	m_sizeButton.cx = m_pImages->GetImageSize ().cx + globalUtils.ScaleByDPI(8);
	m_sizeButton.cy = m_pImages->GetImageSize ().cy + globalUtils.ScaleByDPI(8);
	
	RemoveButtons ();
}
//*********************************************************************************************
void CButtonsList::AddButton (CBCGPToolbarButton* pButton)
{
	ASSERT_VALID (pButton);
	ASSERT (m_pImages != NULL);

	m_Buttons.AddTail (pButton);
	pButton->OnChangeParentWnd (this);

	RebuildLocations ();

	HWND hwnd = pButton->GetHwnd ();
	if (hwnd != NULL)
	{
		::EnableWindow (hwnd, FALSE);
	}
}
//*********************************************************************************************
void CButtonsList::RemoveButtons ()
{
	SelectButton ((CBCGPToolbarButton*) NULL);

	while (!m_Buttons.IsEmpty ())
	{
		CBCGPToolbarButton* pButton = m_Buttons.RemoveHead();
		ASSERT_VALID(pButton);

		pButton->OnChangeParentWnd (NULL);
	}

	m_iScrollOffset = 0;
	m_iScrollTotal = 0;
	m_iScrollPage = 0;

	EnableScrollBarCtrl (SB_VERT, FALSE);
	SetScrollRange (SB_VERT, 0, 0);
}
//*********************************************************************************************
CBCGPToolbarButton* CButtonsList::HitTest (POINT point) const
{
	CRect rectClient;
	GetClientRect (&rectClient);

	for (POSITION pos = m_Buttons.GetHeadPosition (); pos != NULL;)
	{
		CBCGPToolbarButton* pButton = m_Buttons.GetNext (pos);
		ASSERT_VALID(pButton);

		CRect rect = pButton->Rect ();
		rect.OffsetRect (0, -m_iScrollOffset);

		if (rect.PtInRect (point))
		{
			return pButton;
		}
	}

	return NULL;
}
//*********************************************************************************************
void CButtonsList::SelectButton (CBCGPToolbarButton* pButton)
{
	if (m_pSelButton == pButton)
	{
		RedrawSelection ();
		return;
	}

	CRect rectClient;
	GetClientRect (&rectClient);

	CBCGPToolbarButton* pOldSel = m_pSelButton;
	m_pSelButton = pButton;

	if (m_pSelButton != NULL && m_pSelButton->Rect().top - m_iScrollOffset < rectClient.top)
	{
		while (m_iScrollOffset > 0 && m_pSelButton->Rect().top - m_iScrollOffset < rectClient.top)
		{
			OnVScroll(SB_LINEUP, 0, &m_wndScrollBar);
		}

		RedrawWindow();
	}
	else if (m_pSelButton != NULL && m_pSelButton->Rect().bottom - m_iScrollOffset > rectClient.bottom)
	{
		while (m_iScrollOffset < m_iScrollTotal && m_pSelButton->Rect().bottom - m_iScrollOffset > rectClient.bottom)
		{
			OnVScroll(SB_LINEDOWN, 0, &m_wndScrollBar);
		}

		RedrawWindow();
	}
	else
	{
		if (pOldSel != NULL)
		{
			CRect rect = pOldSel->Rect ();
			rect.OffsetRect (0, -m_iScrollOffset);
			
			rect.InflateRect (2, 2);
			InvalidateRect (rect);
		}

		if (m_pSelButton != NULL)
		{
			CRect rect = m_pSelButton->Rect ();
			rect.OffsetRect (0, -m_iScrollOffset);
			
			rect.InflateRect (2, 2);
			
			InvalidateRect (rect);
		}
		
		UpdateWindow ();
	}

	//-------------------------------------------------------
	// Trigger mouse up event (to button click notification):
	//-------------------------------------------------------
	CWnd* pParent = GetParent ();
	if (pParent->GetSafeHwnd() != NULL)
	{
		pParent->SendMessage (	WM_COMMAND,
								MAKEWPARAM (GetDlgCtrlID (), BN_CLICKED), 
								(LPARAM) m_hWnd);
	}
}
//*********************************************************************************************
BOOL CButtonsList::SelectButton (int iImage)
{
	if (iImage < 0)
	{
		SelectButton ((CBCGPToolbarButton*) NULL);
		return TRUE;
	}

	for (POSITION pos = m_Buttons.GetHeadPosition (); pos != NULL;)
	{
		CBCGPToolbarButton* pListButton = m_Buttons.GetNext (pos);
		ASSERT_VALID(pListButton);

		if (pListButton->GetImage () == iImage)
		{
			SelectButton (pListButton);
			return TRUE;
		}
	}

	return FALSE;
}
//********************************************************************************
void CButtonsList::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* /*pScrollBar*/)
{
	int iScrollOffset = m_iScrollOffset;

	switch (nSBCode)
	{
	case SB_TOP:
		iScrollOffset = 0;
		break;

	case SB_BOTTOM:
		iScrollOffset = m_iScrollTotal;
		break;

	case SB_LINEUP:
		iScrollOffset -= m_sizeButton.cy;
		break;

	case SB_LINEDOWN:
		iScrollOffset += m_sizeButton.cy;
		break;

	case SB_PAGEUP:
		iScrollOffset -= m_iScrollPage * (m_sizeButton.cy);
		break;

	case SB_PAGEDOWN:
		iScrollOffset += m_iScrollPage * (m_sizeButton.cy);
		break;

	case SB_THUMBPOSITION:
		iScrollOffset = ((m_sizeButton.cy) / 2 + nPos) / 
			(m_sizeButton.cy) * (m_sizeButton.cy);
		break;

	default:
		return;
	}

	iScrollOffset = min (m_iScrollTotal, max (iScrollOffset, 0));

	if (iScrollOffset != m_iScrollOffset)
	{
		m_iScrollOffset = iScrollOffset;
		SetScrollPos (SB_VERT, m_iScrollOffset);

		CRect rectClient;
		GetClientRect (&rectClient);

		rectClient.right -= ::GetSystemMetrics (SM_CXVSCROLL) + 2;
		rectClient.DeflateRect(1, 1);

		RedrawWindow(rectClient);
	}
}
//********************************************************************************
CScrollBar* CButtonsList::GetScrollBarCtrl(int nBar) const
{
	if (nBar == SB_HORZ || m_wndScrollBar.GetSafeHwnd () == NULL)
	{
		return NULL;
	}

	return (CScrollBar* ) &m_wndScrollBar;
}
//********************************************************************************
void CButtonsList::OnEnable(BOOL bEnable) 
{
	CButton::OnEnable(bEnable);
	RedrawWindow ();
}
//********************************************************************************
void CButtonsList::OnSysColorChange() 
{
	if (m_pImages == NULL)
	{
		return;
	}

	m_pImages->OnSysColorChange ();
	RedrawWindow ();
}
//*********************************************************************************
void CButtonsList::RebuildLocations ()
{
	if (GetSafeHwnd () == NULL || m_Buttons.IsEmpty ())
	{
		m_nButtonsInRow = 0;
		return;
	}

	CRect rectClient;
	GetClientRect (&rectClient);

	CRect rectButtons = rectClient;

	rectButtons.right -= ::GetSystemMetrics (SM_CXVSCROLL) + 1;
	rectButtons.DeflateRect(1, 1);

	int x = rectButtons.left;
	int y = rectButtons.top - m_iScrollOffset;

	m_nButtonsInRow = 0;
	int nButtonsInCurRow = 0;

	CClientDC dc (this);

	for (POSITION pos = m_Buttons.GetHeadPosition (); pos != NULL;)
	{
		CBCGPToolbarButton* pButton = m_Buttons.GetNext(pos);
		ASSERT_VALID(pButton);

		CSize sizeButton = pButton->OnCalculateSize (&dc, m_sizeButton, TRUE);

		if (x + sizeButton.cx > rectButtons.right)
		{
			if (x == rectButtons.left)
			{
				sizeButton.cx = rectButtons.right - rectButtons.left;
			}
			else
			{
				x = rectButtons.left;
				y += sizeButton.cy;

				nButtonsInCurRow = 0;
			}
		}

		pButton->SetRect(CRect (CPoint (x, y), CSize (sizeButton.cx, m_sizeButton.cy)));

		x += sizeButton.cx;
		nButtonsInCurRow++;

		m_nButtonsInRow = max(m_nButtonsInRow, nButtonsInCurRow);
	}

	CBCGPToolbarButton* pLastButton = m_Buttons.GetTail();
	ASSERT_VALID(pLastButton);

	int iVisibleRows = rectButtons.Height () / (m_sizeButton.cy);
	int iTotalRows = pLastButton->Rect ().bottom / (m_sizeButton.cy);

	CRect rectSB;
	GetClientRect (&rectSB);

	rectSB.DeflateRect(1, 1);
	rectSB.left = rectSB.right - ::GetSystemMetrics (SM_CXVSCROLL) - 1;

	int iNonVisibleRows = iTotalRows - iVisibleRows;
	if (iNonVisibleRows > 0)	// Not enough space.
	{
		if (m_wndScrollBar.GetSafeHwnd () == NULL)
		{
			m_wndScrollBar.Create (WS_CHILD | WS_VISIBLE | SBS_VERT, rectSB, this, 1);
			m_wndScrollBar.SetVisualStyle(CBCGPScrollBar::BCGP_SBSTYLE_VISUAL_MANAGER, FALSE);
		}
		
		m_iScrollTotal = iNonVisibleRows * (m_sizeButton.cy);
		m_iScrollPage = iVisibleRows;
	}
	else
	{
		m_iScrollTotal = 0;
	}

	if (m_wndScrollBar.GetSafeHwnd () != NULL)
	{
		m_wndScrollBar.SetWindowPos(NULL, rectSB.left, rectSB.top, rectSB.Width(), rectSB.Height(),
			SWP_NOZORDER | SWP_NOACTIVATE);

		SetScrollRange (SB_VERT, 0, m_iScrollTotal);
	}

	m_bInited = TRUE;
}
//*********************************************************************************
void CButtonsList::OnSize(UINT nType, int cx, int cy) 
{
	CButton::OnSize(nType, cx, cy);

	if (m_wndScrollBar.GetSafeHwnd() != NULL)
	{
		m_wndScrollBar.SetScrollPos(0, FALSE);
	}

	m_iScrollOffset = 0;
	RebuildLocations ();
}
//*********************************************************************************
HBRUSH CButtonsList::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	HBRUSH hbr = CWnd::OnCtlColor (pDC, pWnd, nCtlColor);

	for (POSITION pos = m_Buttons.GetHeadPosition (); pos != NULL;)
	{
		CBCGPToolbarButton* pButton = m_Buttons.GetNext(pos);
		ASSERT_VALID(pButton);

		HWND hwdList = pButton->GetHwnd ();
		if (hwdList == NULL)	// No control
		{
			continue;
		}

		if (hwdList == pWnd->GetSafeHwnd () ||
			::IsChild (hwdList, pWnd->GetSafeHwnd ()))
		{
			HBRUSH hbrButton = pButton->OnCtlColor (pDC, nCtlColor);
			return (hbrButton == NULL) ? hbr : hbrButton;
		}
	}

	return hbr;
}
//*********************************************************************************
UINT CButtonsList::OnGetDlgCode() 
{
	return DLGC_WANTARROWS;
}
//*********************************************************************************
void CButtonsList::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	int nSteps = 1;

	switch (nChar)
	{
	case VK_UP:
		nSteps = m_nButtonsInRow;

	case VK_LEFT:
		if (m_pSelButton != NULL)
		{
			POSITION pos = m_Buttons.Find (m_pSelButton);
			if (pos != NULL)
			{
				for (int i = 0; pos != NULL && i < nSteps; i++)
				{
					m_Buttons.GetPrev(pos);
				}

				if (pos != NULL)
				{
					SelectButton(m_Buttons.GetAt(pos));
					return;
				}
			}
		}
		break;

	case VK_DOWN:
		nSteps = m_nButtonsInRow;

	case VK_RIGHT:
		if (m_pSelButton != NULL)
		{
			POSITION pos = m_Buttons.Find (m_pSelButton);
			if (pos != NULL)
			{
				for (int i = 0; pos != NULL && i < nSteps; i++)
				{
					m_Buttons.GetNext (pos);
				}

				if (pos != NULL)
				{
					SelectButton(m_Buttons.GetAt(pos));
					return;
				}
			}
		}
		break;

	case VK_HOME:
		if (!m_Buttons.IsEmpty ())
		{
			SelectButton(m_Buttons.GetHead ());
		}
		return;

	case VK_END:
		if (!m_Buttons.IsEmpty ())
		{
			SelectButton(m_Buttons.GetTail ());
		}
		return;

	default:
		CButton::OnKeyDown(nChar, nRepCnt, nFlags);
		return;
	}			
	
	if (m_pSelButton == NULL && !m_Buttons.IsEmpty())
	{
		SelectButton(m_Buttons.GetHead ());
	}
}
//*********************************************************************************
void CButtonsList::RedrawButton(CBCGPToolbarButton* pButton)
{
	if (pButton == NULL)
	{
		return;
	}
	
	CRect rect = pButton->Rect ();
	rect.OffsetRect (0, -m_iScrollOffset);
	
	rect.InflateRect (2, 2);
	
	RedrawWindow(rect);
}
//*********************************************************************************
void CButtonsList::RedrawSelection()
{
	RedrawButton(m_pSelButton);
}
//*********************************************************************************
BOOL CButtonsList::OnMouseWheel(UINT nFlags, short zDelta, CPoint /*pt*/) 
{
	if (nFlags & (MK_SHIFT | MK_CONTROL))
	{
		return FALSE;
	}

	if (m_wndScrollBar.GetSafeHwnd () == NULL)
	{
		return NULL;
	}
	
	const int nSteps = abs(zDelta) / WHEEL_DELTA;
	const UINT nSBCode = zDelta < 0 ? SB_LINEDOWN : SB_LINEUP;
	
	for (int i = 0; i < nSteps; i++)
	{
		OnVScroll(nSBCode, 0, &m_wndScrollBar);
	}
	
	return TRUE;
}
//*********************************************************************************
void CButtonsList::OnMouseMove(UINT nFlags, CPoint point) 
{
	if (!m_bTrack)
	{
		m_bTrack = TRUE;
		
		TRACKMOUSEEVENT trackmouseevent;
		trackmouseevent.cbSize = sizeof(trackmouseevent);
		trackmouseevent.dwFlags = TME_LEAVE;
		trackmouseevent.hwndTrack = GetSafeHwnd();
		trackmouseevent.dwHoverTime = HOVER_DEFAULT;
		::BCGPTrackMouse (&trackmouseevent);	
	}

	CBCGPToolbarButton* pButton = HitTest(point);
	if (pButton != m_pHighlightedButton)
	{
		CBCGPToolbarButton* pButtonPrev = m_pHighlightedButton;
		m_pHighlightedButton = pButton;

		RedrawButton(m_pHighlightedButton);
		RedrawButton(pButtonPrev);
	}
	
	CButton::OnMouseMove(nFlags, point);
}
//*****************************************************************************************
LRESULT CButtonsList::OnMouseLeave(WPARAM,LPARAM)
{
	m_bTrack = FALSE;

	if (m_pHighlightedButton != NULL)
	{
		CBCGPToolbarButton* pButtonPrev = m_pHighlightedButton;
		m_pHighlightedButton = NULL;
		
		RedrawButton(pButtonPrev);
	}

	return 0;
}
//*****************************************************************************************
void CButtonsList::OnCancelMode() 
{
	CButton::OnCancelMode();
	
	if (m_bTrack)
	{
		OnMouseLeave(0, 0);
	}
}
//*****************************************************************************************
void CButtonsList::OnKillFocus(CWnd* /*pNewWnd*/) 
{
}
//*****************************************************************************************
void CButtonsList::OnSetFocus(CWnd* /*pOldWnd*/) 
{
}
