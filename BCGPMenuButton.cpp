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

// BCGPMenuButton.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPMenuButton.h"
#include "BCGCBPro.h"
#include "BCGGlobals.h"
#include "BCGPGlobalUtils.h"
#include "BCGPDrawManager.h"
#include "BCGPPopupDlg.h"

#ifndef _BCGSUITE_
#include "MenuImages.h"
#include "BCGPContextMenuManager.h"
#include "BCGPPopupMenu.h"
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define IMAGE_HORZ_MARGIN globalUtils.ScaleByDPI(8)

/////////////////////////////////////////////////////////////////////////////
// CBCGPMenuButton

IMPLEMENT_DYNAMIC(CBCGPMenuButton, CBCGPButton)

CBCGPMenuButton::CBCGPMenuButton()
{
	m_bRightArrow = FALSE;
	m_hMenu = NULL;
	m_nMenuResult = 0;
	m_bMenuIsActive = FALSE;
	m_bStayPressed = FALSE;
	m_bOSMenu = TRUE;
	m_bDefaultClick = FALSE;
	m_bClickOnMenu = FALSE;
	m_pRTIPopupDlg = NULL;
	m_lpszPopupDlgTemplateName = NULL;
	m_pPopupDlg = NULL;
	m_bIsResizablePopup = FALSE;
	m_bIsRightAlignedPopup = FALSE;
}
//*****************************************************************************************
CBCGPMenuButton::~CBCGPMenuButton()
{
}


BEGIN_MESSAGE_MAP(CBCGPMenuButton, CBCGPButton)
	//{{AFX_MSG_MAP(CBCGPMenuButton)
	ON_WM_KEYDOWN()
	ON_WM_LBUTTONDOWN()
	ON_WM_GETDLGCODE()
	ON_WM_LBUTTONUP()
	ON_WM_KILLFOCUS()
	ON_WM_LBUTTONDBLCLK()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CBCGPMenuButton message handlers

CSize CBCGPMenuButton::SizeToContent(BOOL bCalcOnly)
{
	CSize size = CBCGPButton::SizeToContent(TRUE);
	size.cx += CBCGPMenuImages::Size ().cx;

	if (!bCalcOnly)
	{
		SetWindowPos (NULL, -1, -1, size.cx, size.cy, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
	}

	return size;
}
//*****************************************************************************************
void CBCGPMenuButton::OnDrawMenuArrow(CDC* pDC, const CRect& rect, UINT uiState)
{
	ASSERT_VALID(pDC);

	CBCGPMenuImages::Draw (pDC, m_bRightArrow ? CBCGPMenuImages::IdArowRightLarge : CBCGPMenuImages::IdArowDownLarge, 
		rect, (uiState & ODS_DISABLED) ? CBCGPMenuImages::ImageGray : 
		CBCGPMenuImages::GetStateByColor((m_bVisualManagerStyle && !m_bDontSkin && m_clrText != (COLORREF)-1) ? m_clrText : globalData.clrBtnText, FALSE));
}
//*****************************************************************************************
void CBCGPMenuButton::OnDraw (CDC* pDC, const CRect& rect, UINT uiState)
{
	ASSERT_VALID (pDC);

	CBCGPDrawOnGlass dog (m_bOnGlass);

	CSize sizeArrow = CBCGPMenuImages::Size ();

	CRect rectParent = rect;
	rectParent.right -= sizeArrow.cx + IMAGE_HORZ_MARGIN;

	CBCGPButton::OnDraw (pDC, rectParent, uiState);

	CRect rectArrow = rect;
	rectArrow.left = rectParent.right;
	
	OnDrawMenuArrow(pDC, rectArrow, uiState);

	if (m_bDefaultClick)
	{
		//----------------
		// Draw separator:
		//----------------
		CRect rectSeparator = rectArrow;
		rectSeparator.right = rectSeparator.left + 2;
		rectSeparator.DeflateRect (0, 2);

		if (!m_bWinXPTheme || m_bDontUseWinXPTheme)
		{
			rectSeparator.left += m_sizePushOffset.cx;
			rectSeparator.top += m_sizePushOffset.cy;
		}

		if (m_bOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawLine (rectSeparator.left, rectSeparator.top, rectSeparator.left, rectSeparator.bottom,
				globalData.clrBarDkShadow);
		}
		else if (m_bVisualManagerStyle && !m_bDontSkin)
		{
			pDC->Draw3dRect (rectSeparator, globalData.clrBarDkShadow, globalData.clrBarHilite);
		}
		else
		{
			pDC->Draw3dRect (rectSeparator, globalData.clrBtnDkShadow, globalData.clrBtnHilite);
		}
	}
}
//*****************************************************************************************
void CBCGPMenuButton::OnShowMenu () 
{
	CRect rectWindow;
	GetWindowRect (rectWindow);
	
	int x = (m_bRightArrow || m_bIsRightAlignedPopup) ? rectWindow.right : rectWindow.left;
	int y = m_bRightArrow ? rectWindow.top : rectWindow.bottom;
	
	if (m_pRTIPopupDlg != NULL)
	{
		if (m_pPopupDlg->GetSafeHwnd() != NULL)
		{
			ClosePopupDlg(FALSE);
		}
		else
		{
			ASSERT(m_lpszPopupDlgTemplateName != NULL);
			ASSERT(m_pRTIPopupDlg != NULL);
			
			CBCGPDlgPopupMenu* pPopupMenu = new CBCGPDlgPopupMenu(this, m_pRTIPopupDlg, m_lpszPopupDlgTemplateName);
			if (pPopupMenu != NULL)
			{
				pPopupMenu->m_bIsResizable = m_bIsResizablePopup;
				pPopupMenu->SetRightAlign(m_bIsRightAlignedPopup);
				
				m_pPopupDlg = pPopupMenu->m_pDlg;
				if (m_pPopupDlg != NULL)
				{
					m_pPopupDlg->m_pParentEdit = new CBCGPParentMenuButtonPtr(this);
				}
				
				CBCGPPopupMenu* pMenuActive = CBCGPPopupMenu::GetActiveMenu ();
				if (pMenuActive->GetSafeHwnd() != NULL)
				{
					pMenuActive->SendMessage(WM_CLOSE);
				}

				OnBeforeShowPopupDlg(m_pPopupDlg);

				if (m_bStayPressed)
				{
					m_bPushed = TRUE;
					m_bHighlighted = TRUE;
				}

				if (GetExStyle() & WS_EX_LAYOUTRTL)
				{
					x = (m_bRightArrow || m_bIsRightAlignedPopup) ? rectWindow.left : rectWindow.right;
				}

				m_bMenuIsActive = TRUE;
				pPopupMenu->Create(this, x, y, NULL, FALSE, FALSE);
			}
		}

		return;
	}

	if (m_hMenu == NULL || m_bMenuIsActive)
	{
		return;
	}

	if (m_bStayPressed)
	{
		m_bPushed = TRUE;
		m_bHighlighted = TRUE;
	}

	m_bMenuIsActive = TRUE;
	Invalidate ();

#ifndef _BCGSUITE_
	if (!m_bOSMenu && g_pContextMenuManager != NULL)
	{
		m_nMenuResult = g_pContextMenuManager->TrackPopupMenu (
			m_hMenu, x, y, this);
		SetFocus ();
	}
	else
#endif
	{
		m_nMenuResult = ::TrackPopupMenu (m_hMenu, 
			TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_NONOTIFY | TPM_RETURNCMD, 
			x, y, 0, GetSafeHwnd (), NULL);
	}

	if (m_nMenuResult != 0)
	{
		//-------------------------------------------------------
		// Trigger mouse up event (to button click notification):
		//-------------------------------------------------------
		CWnd* pParent = GetParent ();
		if (pParent != NULL)
		{
			pParent->SendMessage (	WM_COMMAND,
									MAKEWPARAM (GetDlgCtrlID (), BN_CLICKED),
									(LPARAM) m_hWnd);
		}
	}

	m_bPushed = FALSE;
	m_bHighlighted = FALSE;
	m_bMenuIsActive = FALSE;
	
	Invalidate ();
	UpdateWindow ();

	if (m_bCaptured)
	{
		ReleaseCapture ();
		m_bCaptured = FALSE;
	}
}
//*****************************************************************************************
void CBCGPMenuButton::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (nChar == VK_SPACE || nChar == VK_DOWN)
	{
		m_bClickOnMenu = TRUE;
		OnShowMenu ();
		return;
	}
	
	CButton::OnKeyDown(nChar, nRepCnt, nFlags);
}
//*****************************************************************************************
void CBCGPMenuButton::OnLButtonDown(UINT nFlags, CPoint point) 
{
	if (m_bMenuIsActive)
	{
		Default ();
		return;
	}

	m_bClickOnMenu = TRUE;

	if (m_bDefaultClick)
	{
		CRect rectClient;
		GetClientRect (rectClient);

		CRect rectArrow = rectClient;
		rectArrow.left = rectArrow.right - CBCGPMenuImages::Size ().cx - IMAGE_HORZ_MARGIN;

		if (!rectArrow.PtInRect (point))
		{
			m_bClickOnMenu = FALSE;
			m_nMenuResult = 0;
			CBCGPButton::OnLButtonDown (nFlags, point);
			return;
		}
	}

	SetFocus ();
	OnShowMenu ();
}
//*****************************************************************************************
UINT CBCGPMenuButton::OnGetDlgCode() 
{
	return DLGC_WANTARROWS;
}
//****************************************************************************************
void CBCGPMenuButton::OnLButtonUp(UINT nFlags, CPoint point) 
{
	if (m_bStayPressed && m_bMenuIsActive && m_bPushed)
	{
		m_bClickiedInside = FALSE;

		CButton::OnLButtonUp(nFlags, point);

		if (m_bCaptured)
		{
			ReleaseCapture ();
			m_bCaptured = FALSE;
		}
	}
	else if (!m_bClickOnMenu)
	{
		CBCGPButton::OnLButtonUp(nFlags, point);
	}
}
//***************************************************************************************
void CBCGPMenuButton::OnKillFocus(CWnd* pNewWnd) 
{
	if (m_bStayPressed && m_bMenuIsActive && m_bPushed)
	{
		CButton::OnKillFocus(pNewWnd);
		
		if (m_bCaptured)
		{
			ReleaseCapture ();
			m_bCaptured = FALSE;
		}
		
		m_bClickiedInside = FALSE;
		m_bHover = FALSE;
	}
	else
	{
		CBCGPButton::OnKillFocus(pNewWnd);
	}
}
//***************************************************************************************
BOOL CBCGPMenuButton::PreTranslateMessage(MSG* pMsg) 
{
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		CBCGPPopupMenu::GetActiveMenu () == NULL)
	{
		m_bClickOnMenu = TRUE;
		OnShowMenu ();
		return TRUE;
	}
	
	return CBCGPButton::PreTranslateMessage(pMsg);
}
//***************************************************************************************
void CBCGPMenuButton::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	if (!m_bMenuIsActive)
	{
		CBCGPButton::OnLButtonDblClk(nFlags, point);
	}

	if (m_bCaptured)
	{
		ReleaseCapture ();
		m_bCaptured = FALSE;
	}
}
//********************************************************************************
void CBCGPMenuButton::EnablePopupDialog(CRuntimeClass* pRTI, UINT nIDTemplate, BOOL bIsResizable, BOOL bIsRightAligned)
{
	EnablePopupDialog(pRTI, MAKEINTRESOURCE(nIDTemplate), bIsResizable, bIsRightAligned);
}
//********************************************************************************
void CBCGPMenuButton::EnablePopupDialog(CRuntimeClass* pRTI, LPCTSTR lpszTemplateName, BOOL bIsResizable, BOOL bIsRightAligned)
{
	ASSERT(pRTI != NULL);

	if (pRTI != NULL && pRTI->m_pfnCreateObject == NULL)
	{
		TRACE(_T("CBCGPMenuButton::EnablePopupDialog: you've to add DECLARE_DYNCREATE to your popup dialog class\n"));
		ASSERT(FALSE);
		
		pRTI = NULL;
	}

	m_pRTIPopupDlg = pRTI;
	m_lpszPopupDlgTemplateName = lpszTemplateName;
	m_bIsResizablePopup = bIsResizable;
	m_bIsRightAlignedPopup = bIsRightAligned;
}
//*****************************************************************************
void CBCGPMenuButton::ClosePopupDlg(BOOL /*bOK*/, DWORD_PTR /*dwUserData*/)
{
	if (m_pPopupDlg->GetSafeHwnd() != NULL)
	{
		m_pPopupDlg->GetParent()->PostMessage(WM_CLOSE);
		m_pPopupDlg = NULL;
	}
	
	m_bPushed = FALSE;
	m_bHighlighted = FALSE;
	m_bMenuIsActive = FALSE;

	RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
}
