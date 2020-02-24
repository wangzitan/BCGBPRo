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
// BCGPPopupWindow.cpp : implementation file
//

#include "stdafx.h"

#include "BCGPVisualManager.h"
#include "BCGPMath.h"
#include "BCGPPopupWindow.h"
#include "BCGPGlobalUtils.h"

#ifndef BCGP_EXCLUDE_POPUP_WINDOW

#include "BCGPDrawManager.h"

#ifndef _BCGSUITE_
	#include "BCGPContextMenuManager.h"
#else
	#define IMAGE_MARGIN			4
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

IMPLEMENT_DYNAMIC(CBCGPPopupWindow, CWnd)

// Timer IDs:
static const int iClosePopupTimerId = 1;
static const int iAnimTimerId = 2;
static const int iCheckActivityTimerId = 3;

static clock_t nLastAnimTime = 0;
static const int nSmallCaptionHeight = 7;

UINT BCGM_ON_CLOSEPOPUPWINDOW = ::RegisterWindowMessage (_T("BCGM_ON_CLOSEPOPUPWINDOW"));

/////////////////////////////////////////////////////////////////////////////
// CBCGPPopupWndButton window

void CBCGPPopupWndButton::SetupOwner()
{
	if (m_pOwner == NULL && GetParent() != NULL)
	{
		m_pOwner = DYNAMIC_DOWNCAST(CBCGPPopupWindow, GetParent()->GetParent());
		
		if (m_pOwner == NULL)
		{
			m_pOwner = DYNAMIC_DOWNCAST(CBCGPPopupWindow, GetParent());
		}
	}
}
//*************************************************************************************
void CBCGPPopupWndButton::OnFillBackground (CDC* pDC, const CRect& rectClient)
{
	SetupOwner();

	if (m_pOwner == NULL || !m_pOwner->OnFillButton(pDC, this, rectClient))
	{
		CBCGPVisualManager::GetInstance ()->OnErasePopupWindowButton (pDC, rectClient, this);
	}
}
//*************************************************************************************
void CBCGPPopupWndButton::OnDrawBorder (CDC* pDC, CRect& rectClient, UINT uiState)
{
	SetupOwner();

	if (m_pOwner == NULL || !m_pOwner->OnDrawButtonBorder(pDC, this, rectClient, uiState))
	{
		CBCGPVisualManager::GetInstance ()->OnDrawPopupWindowButtonBorder (pDC, rectClient, this);
	}
}
//*************************************************************************************
void CBCGPPopupWndButton::OnDraw (CDC* pDC, const CRect& rect, UINT uiState)
{
	SetupOwner();

	if ((m_pOwner == NULL || m_pOwner->GetTheme() == CBCGPPopupWindow::BCGPPopupWindowTheme_VisualManager) &&
		CBCGPVisualManager::GetInstance ()->IsDefaultWinXPPopupButton (this))
	{
		return;
	}

	CBCGPButton::OnDraw (pDC, rect, uiState);
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPPopupWindow

CBCGPPopupWindow::CBCGPPopupWindow()
{
	m_Theme = BCGPPopupWindowTheme_VisualManager;
	m_StemLocation = BCGPPopupWindowStemLocation_None;
	m_nStemSize = globalUtils.ScaleByDPI(15);
	m_nStemGap = 0;

	m_pWndOwner = NULL;
	m_bIsActive = FALSE;
	m_bDontChangeActiveStatus = FALSE;

	m_nAutoCloseTime = 3000;
	m_bSmallCaption = TRUE;
	m_bSmallCaptionGripper = TRUE;
	m_bLargeCaptionFont = FALSE;
	m_bHasCloseButton = TRUE;
	m_bHasPinButton = FALSE;
	m_hMenu = NULL;
	m_pWndDlg = NULL;
	m_uiDlgResID = 0;

	m_nBtnMarginVert = globalUtils.ScaleByDPI(2);
	m_nBtnMarginHorz = globalUtils.ScaleByDPI(2);

	m_nCornerRadius = 0;

	m_nTransparency = 255;	// Opaque

	m_AnimationType = CBCGPPopupMenu::NO_ANIMATION;
	m_AnimationSpeed = 30;
	m_nAnimationAlpha = 0;
	m_bAnimationIsDone = FALSE;
	m_AnimSize = CSize (0, 0);
	m_FinalSize = CSize (0, 0);
	m_bIsAnimRight = FALSE;
	m_bIsAnimDown = FALSE;
	m_bFadeOutAnimation = FALSE;
	m_ptLastPos = CPoint (-1, -1);
	m_bMoving = FALSE;
	m_ptStartMove = CPoint (-1, -1);

	m_pWndShadow = NULL;
	m_bShadow = FALSE;
	m_bShadowIsReady = FALSE;

	m_bIsPinned = FALSE;

	m_btnImageState = CBCGPMenuImages::ImageBlack;
}
//*******************************************************************************
CBCGPPopupWindow::~CBCGPPopupWindow()
{
}
//*******************************************************************************
void CBCGPPopupWindow::SetTheme(BCGPPopupWindowTheme theme)
{
	m_Theme = theme;

	switch (m_Theme)
	{
	case BCGPPopupWindowTheme_VisualManager:
		m_Colors = CBCGPPopupWindowColors();
		break;
		
	case BCGPPopupWindowTheme_Black:
		m_Colors.clrFill = RGB(80, 80, 80);
		m_Colors.clrBorder = RGB(28, 28, 28);
		m_Colors.clrText = RGB(226, 226, 226);
		m_Colors.clrLink = RGB(164, 212, 218);
		m_Colors.clrHoverLink = RGB(202, 236, 230);
		m_Colors.clrPressedButton = RGB(120, 120, 120);
		break;
		
	case BCGPPopupWindowTheme_White:
		m_Colors.clrFill = RGB(255, 255, 255);
		m_Colors.clrBorder = RGB(180, 180, 180);
		m_Colors.clrText = RGB(59, 59, 59);
		m_Colors.clrLink = RGB(43, 87, 154);
		m_Colors.clrHoverLink = RGB(0, 191, 255);
		m_Colors.clrPressedButton = RGB(230, 230, 230);
		break;
		
	case BCGPPopupWindowTheme_Gray:
		m_Colors.clrFill = RGB(241, 241, 241);
		m_Colors.clrBorder = RGB(200, 200, 200);
		m_Colors.clrText = RGB(30, 30, 30);
		m_Colors.clrLink = RGB(43, 87, 154);
		m_Colors.clrHoverLink = RGB(0, 191, 255);
		m_Colors.clrPressedButton = RGB(220, 220, 220);
		break;
	}
}
//*********************************************************************************************************
void CBCGPPopupWindow::SetStemLocation(BCGPPopupWindowStemLocation location, int nSize, int nStemGap)
{
	m_StemLocation = location;

	if (nSize > 0)
	{
		m_nStemSize = nSize;
	}

	m_nStemGap = nStemGap;
}
//*********************************************************************************************************
void CBCGPPopupWindow::SetCustomTheme(const CBCGPPopupWindowColors& colors)
{
	if (globalData.IsHighContastMode() || colors.clrFill == (COLORREF)-1 || colors.clrText == (COLORREF)-1)
	{
		m_Theme = BCGPPopupWindowTheme_VisualManager;
		m_Colors = CBCGPPopupWindowColors();
		return;
	}

	m_Theme = BCGPPopupWindowTheme_Custom;
	m_Colors = colors;

	if (m_Colors.clrPressedButton == (COLORREF)-1)
	{
		if (CBCGPDrawManager::IsDarkColor(m_Colors.clrFill))
		{
			m_Colors.clrPressedButton = CBCGPDrawManager::ColorMakeLighter(m_Colors.clrFill, 0.3);
		}
		else
		{
			m_Colors.clrPressedButton = CBCGPDrawManager::ColorMakeDarker(m_Colors.clrFill);
		}
	}
}

BEGIN_MESSAGE_MAP(CBCGPPopupWindow, CWnd)
	//{{AFX_MSG_MAP(CBCGPPopupWindow)
	ON_WM_TIMER()
	ON_WM_NCDESTROY()
	ON_WM_ERASEBKGND()
	ON_WM_PAINT()
	ON_WM_DESTROY()
	ON_WM_CLOSE()
	ON_WM_WINDOWPOSCHANGED()
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONUP()
	ON_WM_CANCELMODE()
	ON_WM_LBUTTONDOWN()
	ON_WM_SETFOCUS()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_PRINTCLIENT, OnPrintClient)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CBCGPPopupWindow message handlers

BOOL CBCGPPopupWindow::Create(CWnd* pWndOwner, UINT uiDlgResID, HMENU hMenu, CPoint ptPos,
							  CRuntimeClass* pRTIDlgBar, LPARAM lParam)
{
	ASSERT (pRTIDlgBar->IsDerivedFrom (RUNTIME_CLASS (CBCGPPopupDlg)));

	m_hMenu = hMenu;
	m_pWndOwner = pWndOwner;
	m_uiDlgResID = uiDlgResID;

	m_pWndDlg = (CBCGPPopupDlg*) pRTIDlgBar->CreateObject ();
	ASSERT_VALID (m_pWndDlg);

	m_pWndDlg->m_lCustomParam = lParam;

	return CommonCreate (ptPos);
}
//*******************************************************************************
BOOL CBCGPPopupWindow::Create(CWnd* pWndOwner, CBCGPPopupWndParams& params, HMENU hMenu, CPoint ptPos, LPARAM lParam)
{
	m_hMenu = hMenu;
	m_pWndOwner = pWndOwner;

	m_pWndDlg = new CBCGPPopupDlg;
	ASSERT_VALID (m_pWndDlg);

	m_pWndDlg->m_bDefault = TRUE;
	m_pWndDlg->m_lCustomParam = lParam;

	return CommonCreate (ptPos, &params);
}
//*******************************************************************************
CRect CBCGPPopupWindow::GetStemPoints(POINT* points, BOOL bForDraw)
{
	CRect rectPadding;
	rectPadding.SetRectEmpty();

	int delta = AreRoundedCorners() ? 0 : 1;
	int draw_delta = bForDraw ? 1 : 0;

	switch (m_StemLocation)
	{
	case BCGPPopupWindowStemLocation_TopCenter:
		points[0].x = m_FinalSize.cx / 2 - m_nStemSize;
		points[0].y = m_nStemSize;
		points[1].x = points[0].x + m_nStemSize;
		points[1].y = -1 + draw_delta;
		points[2].x = points[1].x + m_nStemSize;
		points[2].y = points[0].y;
		
		rectPadding.top = m_nStemSize;
		break;
		
	case BCGPPopupWindowStemLocation_TopLeft:
		points[0].x = m_nStemSize;
		points[0].y = m_nStemSize;
		points[1].x = points[0].x;
		points[1].y = -1 + draw_delta;
		points[2].x = points[1].x + m_nStemSize;
		points[2].y = points[0].y;
		
		rectPadding.top = m_nStemSize;
		break;
		
	case BCGPPopupWindowStemLocation_TopRight:
		points[0].x = m_FinalSize.cx - m_nStemSize - draw_delta;
		points[0].y = m_nStemSize;
		points[1].x = points[0].x;
		points[1].y = -1 + draw_delta;
		points[2].x = points[1].x - m_nStemSize - 1 + draw_delta;
		points[2].y = points[0].y;
		
		rectPadding.top = m_nStemSize;
		break;
		
	case BCGPPopupWindowStemLocation_BottomCenter:
		points[0].x = m_FinalSize.cx / 2 - m_nStemSize;
		points[0].y = m_FinalSize.cy - m_nStemSize - 1;
		points[1].x = points[0].x + m_nStemSize;
		points[1].y = m_FinalSize.cy - draw_delta;
		points[2].x = points[1].x + m_nStemSize;
		points[2].y = points[0].y;
		
		rectPadding.bottom = m_nStemSize + delta;
		break;
		
	case BCGPPopupWindowStemLocation_BottomLeft:
		points[0].x = m_nStemSize;
		points[0].y = m_FinalSize.cy - m_nStemSize - 1;
		points[1].x = points[0].x;
		points[1].y = m_FinalSize.cy - draw_delta;
		points[2].x = points[1].x + m_nStemSize;
		points[2].y = points[0].y;
		
		rectPadding.bottom = m_nStemSize + delta;
		break;
		
	case BCGPPopupWindowStemLocation_BottomRight:
		points[0].x = m_FinalSize.cx - m_nStemSize - draw_delta;
		points[0].y = m_FinalSize.cy - m_nStemSize - 1;
		points[1].x = points[0].x;
		points[1].y = m_FinalSize.cy - draw_delta;
		points[2].x = points[1].x - m_nStemSize - 1 + draw_delta;
		points[2].y = points[0].y;
		
		rectPadding.bottom = m_nStemSize + delta;
		break;
		
	case BCGPPopupWindowStemLocation_Left:
		points[0].x = m_nStemSize;
		points[0].y = m_FinalSize.cy / 2 - m_nStemSize;
		points[1].x = 0;
		points[1].y = points[0].y + m_nStemSize;
		points[2].x = points[0].x;
		points[2].y = points[1].y + m_nStemSize;
		
		rectPadding.left = m_nStemSize;
		break;
		
	case BCGPPopupWindowStemLocation_Right:
		points[0].x = m_FinalSize.cx - m_nStemSize - 1;
		points[0].y = m_FinalSize.cy / 2 - m_nStemSize;
		points[1].x = m_FinalSize.cx - draw_delta;
		points[1].y = points[0].y + m_nStemSize;
		points[2].x = points[0].x;
		points[2].y = points[1].y + m_nStemSize;
		
		rectPadding.right = m_nStemSize + delta;
		break;
	}

	return rectPadding;
}
//*******************************************************************************
int CBCGPPopupWindow::UpdateWindowRgn()
{
	CRect rectPadding;
	rectPadding.SetRectEmpty();

	if (m_StemLocation != BCGPPopupWindowStemLocation_None)
	{
		POINT points[3];
		rectPadding = GetStemPoints(points);

		CRgn rgn1;
		
		if (m_nCornerRadius > 0)
		{
			rgn1.CreateRoundRectRgn (rectPadding.left, rectPadding.top, m_FinalSize.cx + 1 - rectPadding.right, m_FinalSize.cy + 1 - rectPadding.bottom, m_nCornerRadius, m_nCornerRadius);
		}
		else
		{
			rgn1.CreateRectRgn (rectPadding.left, rectPadding.top, m_FinalSize.cx + 1 - rectPadding.right, m_FinalSize.cy + 1 - rectPadding.bottom);
		}
		
		CRgn rgn2;
		rgn2.CreatePolygonRgn (points, 3, WINDING);
		
		rgn2.CombineRgn(&rgn1, &rgn2, RGN_OR);
		SetWindowRgn(rgn2, FALSE);
	}
	else
	{
		if (m_nCornerRadius > 0)
		{
			CRgn rgn;
			rgn.CreateRoundRectRgn (0, 0, m_FinalSize.cx + 1, m_FinalSize.cy + 1, m_nCornerRadius, m_nCornerRadius);
			
			SetWindowRgn(rgn, FALSE);
		}
	}

	return rectPadding.left;
}
//*******************************************************************************
BOOL CBCGPPopupWindow::CommonCreate (CPoint ptPos, CBCGPPopupWndParams* pParams)
{
	m_ptLastPos = ptPos;
	m_bDontChangeActiveStatus = TRUE;

	m_nStemSize = bcg_clamp(m_nStemSize, 7, GetSystemMetrics(SM_CYCAPTION));

	int nCaptionHeight = GetCaptionHeight ();
	int nTopHeight = nCaptionHeight;

	switch (m_StemLocation)
	{
	case BCGPPopupWindowStemLocation_TopCenter:
	case BCGPPopupWindowStemLocation_TopLeft:
	case BCGPPopupWindowStemLocation_TopRight:
		nTopHeight += m_nStemSize;
		break;
	}

	HWND hwndFocus = ::GetFocus ();
	HWND hwndForeground = ::GetForegroundWindow ();

	CString strClassName = ::AfxRegisterWndClass (
			CS_SAVEBITS,
			::LoadCursor(NULL, IDC_ARROW),
			(HBRUSH)(COLOR_BTNFACE + 1), NULL);

	CRect rectDummy (0, 0, 0, 0);
	DWORD dwStyleEx = WS_EX_TOOLWINDOW | WS_EX_TOPMOST;

	if (!CWnd::CreateEx (dwStyleEx, strClassName, _T(""), WS_POPUP, rectDummy, NULL, 0))
	{
		m_bDontChangeActiveStatus = FALSE;
		return FALSE;
	}

	CSize sizeDialog;

	m_pWndDlg->m_bDontSetFocus = TRUE;

	if (m_uiDlgResID != 0)
	{
		if (!m_pWndDlg->Create (m_uiDlgResID, this) || m_pWndDlg->IsEmptyDlg())
		{
			m_bDontChangeActiveStatus = FALSE;
			return FALSE;
		}

		sizeDialog = GetDialogSize ();
	}
	else
	{
		ASSERT (pParams != NULL);

		if (!m_pWndDlg->CreateFromParams (*pParams, this))
		{
			m_bDontChangeActiveStatus = FALSE;
			return FALSE;
		}

		sizeDialog = m_pWndDlg->GetDlgSize ();
	}

	CSize sizeDlgExtra = m_pWndDlg->GetExtraSize();
	sizeDialog += sizeDlgExtra;

	m_pWndDlg->m_bDontSetFocus = FALSE;

	int nMenuImageScale = (!m_bSmallCaption && m_bLargeCaptionFont) ? 2 : 1;

	CSize sizeBtn = CSize(
		CBCGPMenuImages::Size().cx * nMenuImageScale, 
		CBCGPMenuImages::Size().cy * nMenuImageScale) + globalUtils.ScaleByDPI(CSize(6, 6));

	int nButtonsWidth = 0;
	BOOL bHasPinButton = m_bHasPinButton && m_nAutoCloseTime > 0;

	if (m_bHasCloseButton)
	{
		nButtonsWidth += sizeBtn.cx;
	}
	
	if (bHasPinButton)
	{
		nButtonsWidth += sizeBtn.cx;
	}
	
	if (m_hMenu != NULL)
	{
		nButtonsWidth += sizeBtn.cx;
	}

	BOOL bButtonsOnCaption = (sizeBtn.cy + 2 <= nCaptionHeight);

	if (nButtonsWidth > 0)
	{
		if (bButtonsOnCaption)
		{
			sizeDialog.cx = max(sizeDialog.cx, 3 * nButtonsWidth + m_nBtnMarginHorz);
		}
		else
		{
			sizeDialog.cx += m_nBtnMarginHorz + nButtonsWidth;
		}
	}

	CRect rectScreen;

	MONITORINFO mi;
	mi.cbSize = sizeof (MONITORINFO);

	if (ptPos == CPoint(-2, -2) && m_pWndOwner->GetSafeHwnd() != NULL && GetMonitorInfo(MonitorFromWindow(m_pWndOwner->GetSafeHwnd(), MONITOR_DEFAULTTONEAREST), &mi))
	{
		rectScreen = mi.rcWork;
	}
	else if (ptPos != CPoint(-1, -1) && GetMonitorInfo (MonitorFromPoint (ptPos, MONITOR_DEFAULTTONEAREST), &mi))
	{
		rectScreen = mi.rcWork;
	}
	else
	{
		::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectScreen, 0);
	}

	sizeDialog.cx = min (rectScreen.Width () - 2, sizeDialog.cx);
	sizeDialog.cy = min (rectScreen.Height () - nTopHeight - 2, sizeDialog.cy);

	m_FinalSize = sizeDialog;
	m_FinalSize.cy += nTopHeight + 2;
	m_FinalSize.cx += 2;

	int nTopButtons = nTopHeight;
	int nLeftButtons = 0;

	switch (m_StemLocation)
	{
	case BCGPPopupWindowStemLocation_TopCenter:
	case BCGPPopupWindowStemLocation_TopLeft:
	case BCGPPopupWindowStemLocation_TopRight:
		nTopButtons += m_nStemSize;
		break;

	case BCGPPopupWindowStemLocation_BottomCenter:
	case BCGPPopupWindowStemLocation_BottomLeft:
	case BCGPPopupWindowStemLocation_BottomRight:
		m_FinalSize.cy += m_nStemSize;
		break;

	case BCGPPopupWindowStemLocation_Left:
		if (bButtonsOnCaption)
		{
			nLeftButtons += m_nStemSize;
		}

	case BCGPPopupWindowStemLocation_Right:
		m_FinalSize.cx += m_nStemSize;
		break;
	}

	CWnd* pBtnParent = bButtonsOnCaption ? (CWnd*) this : m_pWndDlg;
	
	int nBtnVertOffset = bButtonsOnCaption ? (nTopButtons - sizeBtn.cy) / 2 + 1 : m_nBtnMarginVert;

	CRect rectBtn = CRect(CPoint (sizeDialog.cx - sizeBtn.cx - m_nBtnMarginHorz + nLeftButtons, nBtnVertOffset), sizeBtn);
	
	m_btnImageState = CBCGPMenuImages::ImageBlack;
	
	COLORREF clrText = GetTextColor(m_pWndDlg, bButtonsOnCaption, TRUE);
	if (clrText != (COLORREF)-1)
	{
		m_btnImageState = CBCGPMenuImages::GetStateByColor(clrText, FALSE);
	}

	if (m_bHasCloseButton)
	{
		m_btnClose.m_pOwner = this;
		m_btnClose.Create (_T("Close"), WS_CHILD | WS_VISIBLE, rectBtn, pBtnParent, (UINT) -1);

		m_btnClose.SetStdImage (CBCGPMenuImages::IdClose, m_btnImageState);
		m_btnClose.m_bDrawFocus = FALSE;
		m_btnClose.m_nFlatStyle = CBCGPButton::BUTTONSTYLE_FLAT;
		m_btnClose.SetDrawText(FALSE, FALSE);

		m_btnClose.m_bIsCaptionButton = bButtonsOnCaption;
		m_btnClose.m_bIsCloseButton = TRUE;

		rectBtn.OffsetRect (-sizeBtn.cx - 1, 0);
	}

	if (bHasPinButton)
	{
		m_btnPin.m_pOwner = this;
		m_btnPin.Create (_T("Pin"), WS_CHILD | WS_VISIBLE, rectBtn, pBtnParent, (UINT) -1);
		
		m_btnPin.SetStdImage (CBCGPMenuImages::IdPinHorz, m_btnImageState);
		m_btnPin.m_bDrawFocus = FALSE;
		m_btnPin.m_nFlatStyle = CBCGPButton::BUTTONSTYLE_FLAT;
		m_btnPin.SetDrawText(FALSE, FALSE);
		
		m_btnPin.m_bIsCaptionButton = bButtonsOnCaption;
		m_btnPin.m_bIsPinButton = TRUE;
		
		rectBtn.OffsetRect (-sizeBtn.cx - 1, 0);
	}

	if (m_hMenu != NULL)
	{
		m_btnMenu.m_pOwner = this;
		m_btnMenu.Create (_T("Options"), WS_CHILD | WS_VISIBLE, rectBtn, pBtnParent, (UINT) -1);

		m_btnMenu.SetStdImage (CBCGPMenuImages::IdArowDownLarge, m_btnImageState);
		m_btnMenu.m_bDrawFocus = FALSE;
		m_btnMenu.m_nFlatStyle = CBCGPButton::BUTTONSTYLE_FLAT;
		m_btnMenu.SetDrawText(FALSE, FALSE);

		m_btnMenu.m_bIsCaptionButton = bButtonsOnCaption;
	}

	if (ptPos == CPoint (-1, -1))
	{
		ptPos.x = rectScreen.right - m_FinalSize.cx;
		ptPos.y = rectScreen.bottom - m_FinalSize.cy;
	}
	else if (ptPos == CPoint (-2, -2))
	{
		CRect rectOwner = rectScreen;
		if (m_pWndOwner->GetSafeHwnd() != NULL)
		{
			m_pWndOwner->GetWindowRect(rectOwner);
		}

		ptPos.x = rectOwner.left + max(0, (rectOwner.Width() - m_FinalSize.cx) / 2);
		ptPos.y = rectOwner.top + max(0, (rectOwner.Height() - m_FinalSize.cy) / 2);

		if (ptPos.x + m_FinalSize.cx > rectScreen.right)
		{
			ptPos.x = rectScreen.right - m_FinalSize.cx;
		}

		if (ptPos.y + m_FinalSize.cy > rectScreen.bottom)
		{
			ptPos.y = rectScreen.bottom - m_FinalSize.cy;
		}
	}
	else
	{
		AdjustLocationByStem(ptPos, m_nStemGap);

		if (ptPos.x < rectScreen.left)
		{
			ptPos.x = rectScreen.left;
		}
		else if (ptPos.x + m_FinalSize.cx > rectScreen.right)
		{
			ptPos.x = rectScreen.right - m_FinalSize.cx;
		}

		if (ptPos.y < rectScreen.top)
		{
			ptPos.y = rectScreen.top;
		}
		else if (ptPos.y + m_FinalSize.cy > rectScreen.bottom)
		{
			ptPos.y = rectScreen.bottom - m_FinalSize.cy;
		}
	}

	int nLeftOffset = UpdateWindowRgn();

	BOOL bIsRTL = FALSE;

	if (m_pWndOwner->GetSafeHwnd() != NULL && (m_pWndOwner->GetExStyle () & WS_EX_LAYOUTRTL))
	{
		ModifyStyleEx(0, WS_EX_LAYOUTRTL);
		bIsRTL = TRUE;
	}

	OnBeforeShow (ptPos);

	if (m_nCornerRadius > 0 && (!bIsRTL || sizeDlgExtra == CSize(0, 0)))
	{
		CRgn rgn1;
		rgn1.CreateRoundRectRgn (0, 0, sizeDialog.cx + 1, sizeDialog.cy + 1, m_nCornerRadius - 1, m_nCornerRadius - 1);
		
		CRgn rgn;
		rgn.CreateRectRgn(0, 0, sizeDialog.cx + 1, sizeDialog.cy - m_nCornerRadius);
		
		rgn.CombineRgn(&rgn1, &rgn, RGN_OR);
		
		m_pWndDlg->SetWindowRgn(rgn, FALSE);
	}
	
	SetWindowPos (&wndTop, ptPos.x, ptPos.y, m_FinalSize.cx, m_FinalSize.cy, SWP_NOACTIVATE | SWP_SHOWWINDOW);

#ifndef _BCGSUITE_
	if (m_bShadow && globalData.IsWindowsLayerSupportAvailable())
	{
		int nDepth = m_nCornerRadius != 0 ? m_nCornerRadius : 5;

		POINT points[3];
		CRect rectPadding = GetStemPoints(points);

		m_pWndShadow = new CBCGPShadowWnd(this, 5, nDepth, CSize(nDepth, nDepth), rectPadding);

		if (m_StemLocation != BCGPPopupWindowStemLocation_None)
		{
			m_pWndShadow->SetStemRegion(points, 3);
		}

		m_pWndShadow->Create();
	}
#endif

	StartAnimation ();

	m_pWndDlg->SetWindowPos (NULL, nLeftOffset + 1, nTopHeight + 1, sizeDialog.cx, sizeDialog.cy, SWP_NOZORDER | SWP_NOACTIVATE);

	SetTimer (iCheckActivityTimerId, 100, NULL);

	if (::IsWindow (hwndForeground))
	{
		::SetForegroundWindow (hwndForeground);
	}

	if (::IsWindow (hwndFocus))
	{
		::SetFocus (hwndFocus);
	}

#ifndef _BCGSUITE_
	CBCGPPopupMenu* pActiveMenu = CBCGPPopupMenu::GetSafeActivePopupMenu();
	if (pActiveMenu != NULL)
	{
		pActiveMenu->SetWindowPos(&wndTop, -1, -1, -1, -1, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOREDRAW | SWP_NOSIZE);
	}

	if (m_pWndShadow->GetSafeHwnd() != NULL && GetActualAnimationType() == CBCGPPopupMenu::NO_ANIMATION)
	{
		m_pWndShadow->Repos();
		m_bShadowIsReady = TRUE;
	}
#endif

	m_bDontChangeActiveStatus = FALSE;
	return TRUE;
}
//*******************************************************************************
CSize CBCGPPopupWindow::GetDialogSize ()
{
	CDialogTemplate dlgt;
	if (!dlgt.Load (MAKEINTRESOURCE (m_uiDlgResID)))
	{
		ASSERT (FALSE);
		return CSize (0, 0);
	}

	CSize sizeDialog;
	dlgt.GetSizeInPixels (&sizeDialog);

	return sizeDialog;
}
//*******************************************************************************
void CBCGPPopupWindow::OnTimer(UINT_PTR nIDEvent) 
{
	const CBCGPPopupMenu::ANIMATION_TYPE animationType = GetActualAnimationType();

	switch (nIDEvent)
	{
	case iAnimTimerId:
		if (!m_bAnimationIsDone)
		{
			clock_t nCurrAnimTime = clock ();

			int nDuration = nCurrAnimTime - nLastAnimTime;
			int nSteps = (int) (.5 + (float) nDuration / m_AnimationSpeed);

			if (m_bFadeOutAnimation)
			{
				nSteps = -nSteps;
			}

			switch (animationType)
			{
			case CBCGPPopupMenu::UNFOLD:
				m_AnimSize.cx += nSteps * m_nAnimStepX;
				// no break intentionally

			case CBCGPPopupMenu::SLIDE:
				m_AnimSize.cy += nSteps * m_nAnimStepY;
				break;

			case CBCGPPopupMenu::FADE:
				m_iFadePercent += m_iFadeStep;

				if (m_iFadePercent > 100 + nSteps * m_iFadeStep)
				{
					m_iFadePercent = 101;
				}
				break;
			}

			m_AnimSize.cx = max (0, min (m_AnimSize.cx, m_FinalSize.cx));
			m_AnimSize.cy = max (0, min (m_AnimSize.cy, m_FinalSize.cy));

			if (m_bFadeOutAnimation &&
				!m_bIsActive &&
				(m_AnimSize.cx == 0 || m_AnimSize.cy == 0 ||
				(animationType == CBCGPPopupMenu::FADE && m_iFadePercent <= 50)))
			{
				SendMessage (WM_CLOSE);
				return;
			}

			if ((animationType != CBCGPPopupMenu::FADE && m_AnimSize.cy >= m_FinalSize.cy && m_AnimSize.cx >= m_FinalSize.cx) ||
				(animationType == CBCGPPopupMenu::UNFOLD && m_AnimSize.cx >= m_FinalSize.cx) ||
				(animationType == CBCGPPopupMenu::FADE && m_iFadePercent > 100) ||
				m_bIsActive)
			{
				m_AnimSize.cx = m_FinalSize.cx;
				m_AnimSize.cy = m_FinalSize.cy;

				KillTimer (iAnimTimerId);

				if (m_btnPin.GetSafeHwnd () != NULL)
				{
					m_btnPin.ShowWindow (SW_SHOWNOACTIVATE);
				}

				if (m_btnClose.GetSafeHwnd () != NULL)
				{
					m_btnClose.ShowWindow (SW_SHOWNOACTIVATE);
				}

				if (m_btnMenu.GetSafeHwnd () != NULL)
				{
					m_btnMenu.ShowWindow (SW_SHOWNOACTIVATE);
				}

#ifndef _BCGSUITE_
	BOOL bIsWindowsLayerSupportAvailable = globalData.IsWindowsLayerSupportAvailable();
#else
	BOOL bIsWindowsLayerSupportAvailable = TRUE;
#endif

				if (bIsWindowsLayerSupportAvailable && globalData.m_nBitsPerPixel > 8 &&
					m_nTransparency < 255)
				{
					ModifyStyleEx(0, WS_EX_LAYERED);
					BYTE nTransparency = m_bIsActive ? (BYTE) 255 : m_nTransparency;

					globalData.SetLayeredAttrib (GetSafeHwnd (), 0, nTransparency, LWA_ALPHA);

#ifndef _BCGSUITE_
					if (m_pWndShadow->GetSafeHwnd () != NULL)
					{
						m_pWndShadow->UpdateTransparency(nTransparency);
					}
#endif
				}

#ifndef _BCGSUITE_
				if (m_pWndShadow->GetSafeHwnd() != NULL)
				{
					m_pWndShadow->Repos();
					m_bShadowIsReady = TRUE;
				}
#endif
				m_pWndDlg->SetWindowPos (NULL,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_NOREDRAW|SWP_NOZORDER|SWP_SHOWWINDOW | SWP_NOACTIVATE);
				m_pWndDlg->ValidateRect (NULL);
    
				m_bAnimationIsDone = TRUE;

				if (m_btnPin.GetSafeHwnd () != NULL)
				{
					m_btnPin.RedrawWindow();
				}
				
				if (m_btnClose.GetSafeHwnd () != NULL)
				{
					m_btnClose.RedrawWindow();
				}
				
				if (m_btnMenu.GetSafeHwnd () != NULL)
				{
					m_btnMenu.RedrawWindow();
				}

				if (m_nAutoCloseTime > 0)
				{
					SetTimer (iClosePopupTimerId, m_nAutoCloseTime, NULL);
				}
			}

			if (m_bFadeOutAnimation && animationType != CBCGPPopupMenu::FADE)
			{
				CRect rectWnd;
				GetWindowRect (rectWnd);

				int x = m_bIsAnimRight ? 
					rectWnd.left : rectWnd.right - m_AnimSize.cx;
				int y = m_bIsAnimDown ? 
					rectWnd.top : rectWnd.bottom - m_AnimSize.cy;

				SetWindowPos (NULL, x, y, m_AnimSize.cx, m_AnimSize.cy,
					SWP_NOZORDER | SWP_NOACTIVATE);
			}
			else
			{
				RedrawWindow(NULL, NULL, RDW_ALLCHILDREN | RDW_ERASE | RDW_INVALIDATE | RDW_UPDATENOW);
			}

			nLastAnimTime = nCurrAnimTime;
		}
		break;

	case iClosePopupTimerId:
		if (!m_bIsActive && !m_bMoving && !m_bIsPinned)
		{
			KillTimer (iClosePopupTimerId);
			StartAnimation (FALSE);
		} 
		return;

	case iCheckActivityTimerId:
		if (!m_bMoving)
		{
			BOOL bWasActive = m_bIsActive;

			CRect rectWnd;
			GetWindowRect (rectWnd);

			CPoint ptCursor;
			GetCursorPos (&ptCursor);

			m_bIsActive = 
				rectWnd.PtInRect (ptCursor) || m_pWndDlg->HasFocus ();

#ifndef _BCGSUITE_
			BOOL bIsWindowsLayerSupportAvailable = globalData.IsWindowsLayerSupportAvailable();
#else
			BOOL bIsWindowsLayerSupportAvailable = TRUE;
#endif
			if (m_bIsActive != bWasActive &&
				bIsWindowsLayerSupportAvailable && 
				globalData.m_nBitsPerPixel > 8 &&
				m_nTransparency < 255)
			{
				BYTE nTransparency = m_bIsActive ? (BYTE) 255 : m_nTransparency;
				globalData.SetLayeredAttrib (GetSafeHwnd (), 0, 
					nTransparency, LWA_ALPHA);

#ifndef _BCGSUITE_
				if (m_pWndShadow->GetSafeHwnd () != NULL)
				{
					if (m_bIsActive && !m_bShadowIsReady)
					{
						m_pWndShadow->Repos();
						m_bShadowIsReady = TRUE;
					}

					m_pWndShadow->UpdateTransparency (nTransparency);
				}
#endif
			}
		}
	}

	CWnd::OnTimer(nIDEvent);
}
//*******************************************************************************
void CBCGPPopupWindow::OnNcDestroy() 
{
	CWnd::OnNcDestroy();
	delete this;
}
//*******************************************************************************
BOOL CBCGPPopupWindow::OnEraseBkgnd(CDC* /*pDC*/)
{
	return TRUE;
}
//*******************************************************************************
void CBCGPPopupWindow::OnPaint() 
{
	CPaintDC dc(this); // device context for painting

	if (!m_bAnimationIsDone)
	{
		DrawAnimation (&dc);
	}
	else
	{
		OnDraw (&dc);
	}
}
//*******************************************************************************
CRect CBCGPPopupWindow::GetCaptionRect ()
{
	int nCaptionHeight = GetCaptionHeight ();

	CRect rectClient;
	GetClientRect (&rectClient);
	
	CRect rectCaption = rectClient;

	rectCaption.DeflateRect (1, 1);

	switch (m_StemLocation)
	{
	case BCGPPopupWindowStemLocation_TopCenter:
	case BCGPPopupWindowStemLocation_TopLeft:
	case BCGPPopupWindowStemLocation_TopRight:
		rectCaption.top += m_nStemSize;
		break;

	case BCGPPopupWindowStemLocation_Left:
		rectCaption.left += m_nStemSize;
		break;
	}

	rectCaption.bottom = rectCaption.top + nCaptionHeight;

	return rectCaption;
}
//*******************************************************************************
BOOL CBCGPPopupWindow::ProcessCommand (HWND hwnd)
{
	ASSERT_VALID (m_pWndDlg);

	if (hwnd == m_btnClose.GetSafeHwnd ())
	{
		if (OnClickCloseButton())
		{
			SendMessage (WM_CLOSE);
		}

		return TRUE;
	}

	if (hwnd == m_btnPin.GetSafeHwnd ())
	{
		m_bIsPinned = !m_bIsPinned;
		m_btnPin.SetStdImage (m_bIsPinned ? CBCGPMenuImages::IdPinVert : CBCGPMenuImages::IdPinHorz, m_btnImageState);
		return TRUE;
	}

	if (hwnd == m_btnMenu.GetSafeHwnd () && m_hMenu != NULL)
	{
		CRect rectMenuBtn;
		m_btnMenu.GetWindowRect (rectMenuBtn);

		const int x = rectMenuBtn.left;
		const int y = rectMenuBtn.bottom;

		UINT nMenuResult = 0;

		m_pWndDlg->m_bMenuIsActive = TRUE;
		
#ifndef _BCGSUITE_
		if (g_pContextMenuManager != NULL)
		{
			const BOOL bMenuShadows = CBCGPMenuBar::IsMenuShadows ();
			CBCGPMenuBar::EnableMenuShadows (FALSE);

			nMenuResult = g_pContextMenuManager->TrackPopupMenu (
				m_hMenu, x, y, this);

			CBCGPMenuBar::EnableMenuShadows (bMenuShadows);
		}
		else
#endif
		{
			nMenuResult = ::TrackPopupMenu (m_hMenu, 
				TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_NONOTIFY | TPM_RETURNCMD, 
				x, y, 0, GetSafeHwnd (), NULL);
		}

		m_pWndDlg->m_bMenuIsActive = FALSE;
		
		if (nMenuResult != 0)
		{
			if (m_pWndOwner != NULL)
			{
				m_pWndOwner->PostMessage (WM_COMMAND, nMenuResult);
			}
			else
			{
				m_pWndDlg->PostMessage (WM_COMMAND, nMenuResult);
			}
		}

		OnCancelMode ();
		return TRUE;
	}

	return FALSE;
}
//*******************************************************************************
void CBCGPPopupWindow::AdjustLocationByStem(CPoint& point, int nDelta)
{
	if (m_StemLocation == BCGPPopupWindowStemLocation_None)
	{
		return;
	}

	switch (m_StemLocation)
	{
	case BCGPPopupWindowStemLocation_TopCenter:
		point.x -= m_FinalSize.cx / 2;
		point.y += nDelta;
		break;

	case BCGPPopupWindowStemLocation_TopLeft:
		point.x -= m_nStemSize;
		point.y += nDelta;
		break;
		
	case BCGPPopupWindowStemLocation_TopRight:
		point.x -= m_FinalSize.cx - m_nStemSize;
		point.y += nDelta;
		break;
		
	case BCGPPopupWindowStemLocation_BottomCenter:
		point.x -= m_FinalSize.cx / 2;
		point.y -= m_FinalSize.cy + nDelta;
		break;
		
	case BCGPPopupWindowStemLocation_BottomLeft:
		point.x -= m_nStemSize;
		point.y -= m_FinalSize.cy + nDelta;
		break;
		
	case BCGPPopupWindowStemLocation_BottomRight:
		point.x -= m_FinalSize.cx - m_nStemSize;
		point.y -= m_FinalSize.cy + nDelta;
		break;
		
	case BCGPPopupWindowStemLocation_Left:
		point.y -= m_FinalSize.cy / 2;
		point.x += nDelta;
		break;

	case BCGPPopupWindowStemLocation_Right:
		point.y -= m_FinalSize.cy / 2;
		point.x -= m_FinalSize.cx + nDelta;
		break;
	}
}
//*******************************************************************************
BOOL CBCGPPopupWindow::UpdateContent(const CString& strText, const CPoint& ptScreen)
{
	if (m_pWndDlg->GetSafeHwnd() == NULL || !m_pWndDlg->m_bDefault)
	{
		return FALSE;
	}

	CSize sizeDialog = m_pWndDlg->UpdateContent(strText);
	if (sizeDialog == CSize(0, 0))
	{
		return FALSE;
	}

	int nCaptionHeight = GetCaptionHeight();

	m_FinalSize = sizeDialog;
	m_FinalSize.cy += nCaptionHeight + 2;
	m_FinalSize.cx += 2;

	switch (m_StemLocation)
	{
	case BCGPPopupWindowStemLocation_TopCenter:
	case BCGPPopupWindowStemLocation_TopLeft:
	case BCGPPopupWindowStemLocation_TopRight:
	case BCGPPopupWindowStemLocation_BottomCenter:
	case BCGPPopupWindowStemLocation_BottomLeft:
	case BCGPPopupWindowStemLocation_BottomRight:
		m_FinalSize.cy += m_nStemSize;
		break;
		
	case BCGPPopupWindowStemLocation_Left:
	case BCGPPopupWindowStemLocation_Right:
		m_FinalSize.cx += m_nStemSize;
		break;
	}

	CPoint pt = ptScreen;
	AdjustLocationByStem(pt, m_nStemGap);

	int nLeftOffset = UpdateWindowRgn();

	if (m_nCornerRadius > 0)
	{
		CRgn rgn1;
		rgn1.CreateRoundRectRgn (0, 0, sizeDialog.cx + 1, sizeDialog.cy + 1, m_nCornerRadius - 1, m_nCornerRadius - 1);
		
		CRgn rgn2;
		rgn2.CreateRectRgn(0, 0, sizeDialog.cx + 1, sizeDialog.cy - m_nCornerRadius);
		
		rgn2.CombineRgn(&rgn1, &rgn2, RGN_OR);
		
		m_pWndDlg->SetWindowRgn(rgn2, FALSE);
	}
	
	SetWindowPos (&wndTop, pt.x, pt.y, m_FinalSize.cx, m_FinalSize.cy, SWP_NOACTIVATE | SWP_SHOWWINDOW);
	
	int nTopHeight = GetCaptionHeight();
	if (m_StemLocation == BCGPPopupWindowStemLocation_TopCenter || m_StemLocation == BCGPPopupWindowStemLocation_TopLeft || m_StemLocation == BCGPPopupWindowStemLocation_TopRight)
	{
		nTopHeight += m_nStemSize;
	}

	m_pWndDlg->SetWindowPos(NULL, nLeftOffset + 1, nTopHeight + 1, sizeDialog.cx, sizeDialog.cy, SWP_NOZORDER | SWP_NOACTIVATE);
	
#ifndef _BCGSUITE_
	if (m_pWndShadow->GetSafeHwnd() != NULL)
	{
		if (m_StemLocation != BCGPPopupWindowStemLocation_None)
		{
			POINT points[3];
			GetStemPoints(points);

			m_pWndShadow->SetStemRegion(points, 3);
		}

		m_pWndShadow->Repos();
	}
#endif

	RedrawWindow(NULL, NULL, RDW_ALLCHILDREN | RDW_ERASE | RDW_INVALIDATE | RDW_UPDATENOW);

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPPopupWindow::OnCommand(WPARAM wParam, LPARAM lParam) 
{
	if (ProcessCommand ((HWND)lParam))
	{
		return TRUE;
	}
	
	return CWnd::OnCommand(wParam, lParam);
}
//*******************************************************************************
void CBCGPPopupWindow::OnDestroy() 
{
#ifndef _BCGSUITE_
	if (m_pWndShadow->GetSafeHwnd () != NULL)
	{
		m_pWndShadow->ShowWindow (SW_HIDE);
		m_pWndShadow->DestroyWindow ();
		m_pWndShadow = NULL;
	}
#endif
	if (m_pWndDlg != NULL)
	{
		m_pWndDlg->DestroyWindow ();
		delete m_pWndDlg;
		m_pWndDlg = NULL;
	}

	CWnd::OnDestroy();
}
//*******************************************************************************
LRESULT CBCGPPopupWindow::OnPrintClient(WPARAM wp, LPARAM lp)
{
	if (lp & PRF_CLIENT)
	{
		CDC* pDC = CDC::FromHandle ((HDC) wp);
		ASSERT_VALID (pDC);

		OnDraw (pDC);
	}

	return 0;
}
//*******************************************************************************
void CBCGPPopupWindow::OnDraw (CDC* pDC)
{
	ASSERT_VALID (pDC);

	BCGPPopupWindowStemLocation stemLocation = m_bAnimationIsDone ? m_StemLocation : BCGPPopupWindowStemLocation_None;

	CRect rectClient;
	GetClientRect(&rectClient);

	const int nImageMargin = globalUtils.ScaleByDPI(IMAGE_MARGIN);

	if (m_nCornerRadius > 0)
	{
		CRect rectBottom = rectClient;
		rectBottom.top = rectBottom.bottom - m_nCornerRadius - 1;

		if (m_Theme == BCGPPopupWindowTheme_VisualManager || globalData.IsHighContastMode())
		{
			CBCGPVisualManager::GetInstance ()->OnFillPopupWindowBackground(pDC, rectBottom);
		}
		else
		{
			pDC->FillSolidRect(rectBottom, m_Colors.clrFill);
		}
	}

	CRect rectFill = rectClient;

	switch (stemLocation)
	{
	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_BottomCenter:
	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_BottomLeft:
	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_BottomRight:
		rectFill.top -= GetCaptionHeight() + m_nStemSize;
		break;

	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_Left:
		rectFill.left -= m_nStemSize;
		rectFill.OffsetRect(0, GetCaptionHeight());
		break;

	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_Right:
		rectFill.right += m_nStemSize;
		rectFill.OffsetRect(0, GetCaptionHeight());
		break;
	}

	if (rectFill != rectClient)
	{
		if (m_Theme == BCGPPopupWindowTheme_VisualManager || globalData.IsHighContastMode())
		{
			CBCGPVisualManager::GetInstance ()->OnFillPopupWindowBackground(pDC, rectFill);
		}
		else
		{
			pDC->FillSolidRect(rectFill, m_Colors.clrFill);
		}
	}

	CRect rectCaption = GetCaptionRect();

	COLORREF clrText = m_Colors.clrText == (COLORREF)-1 ? RGB(0, 0, 0) : m_Colors.clrText;

	if (m_Colors.clrFill == (COLORREF)-1 || globalData.IsHighContastMode())
	{
		clrText = CBCGPVisualManager::GetInstance ()->OnDrawPopupWindowCaption(pDC, rectCaption, this);
	}
	else
	{
		CRect rectFillCaption = rectCaption;

		switch (stemLocation)
		{
		case CBCGPPopupWindow::BCGPPopupWindowStemLocation_TopCenter:
		case CBCGPPopupWindow::BCGPPopupWindowStemLocation_TopLeft:
		case CBCGPPopupWindow::BCGPPopupWindowStemLocation_TopRight:
			rectFillCaption.top -= m_nStemSize;
			break;
		}

		pDC->FillSolidRect(rectFillCaption, m_Colors.clrFill);

		if (HasSmallCaption() && IsSmallCaptionGripper())
		{
			CBCGPDrawManager dm(*pDC);
			
			int dx = rectCaption.Height();
			
			for (int x = rectCaption.CenterPoint().x - 3 * dx; x <= rectCaption.CenterPoint().x + 3 * dx; x += dx)
			{
				CRect rect = rectCaption;
				rect.left = x;
				rect.right = x + dx;
				rect.DeflateRect(2, 2);
				
				dm.DrawEllipse(rect, clrText, (COLORREF)-1);
			}
		}
	}
	
	CRect rectBorder = rectClient;
	
	if (stemLocation != BCGPPopupWindowStemLocation_None)
	{
		CRect rectExclude(0, 0, 0, 0);

		switch (stemLocation)
		{
		case BCGPPopupWindowStemLocation_TopCenter:
			rectBorder.top += m_nStemSize;
			rectExclude = CRect(rectBorder.CenterPoint().x - m_nStemSize, rectBorder.top, rectBorder.CenterPoint().x + m_nStemSize, rectBorder.top + 2);
			break;

		case BCGPPopupWindowStemLocation_TopLeft:
			rectBorder.top += m_nStemSize;
			rectExclude = CRect(rectBorder.left + m_nStemSize, rectBorder.top, rectBorder.left + m_nStemSize + m_nStemSize, rectBorder.top + 2);
			break;
			
		case BCGPPopupWindowStemLocation_TopRight:
			rectBorder.top += m_nStemSize;
			rectExclude = CRect(rectBorder.right - m_nStemSize - m_nStemSize, rectBorder.top, rectBorder.right - m_nStemSize, rectBorder.top + 2);
			break;
			
		case BCGPPopupWindowStemLocation_BottomCenter:
			rectBorder.bottom -= m_nStemSize;
			rectExclude = CRect(rectBorder.CenterPoint().x - m_nStemSize, rectBorder.bottom - 2, rectBorder.CenterPoint().x + m_nStemSize, rectBorder.bottom);
			break;

		case BCGPPopupWindowStemLocation_BottomLeft:
			rectBorder.bottom -= m_nStemSize;
			rectExclude = CRect(m_nStemSize, rectBorder.bottom - 2, m_nStemSize + m_nStemSize, rectBorder.bottom);
			break;
			
		case BCGPPopupWindowStemLocation_BottomRight:
			rectBorder.bottom -= m_nStemSize;
			rectExclude = CRect(rectBorder.right - m_nStemSize - m_nStemSize, rectBorder.bottom - 2, rectBorder.right - m_nStemSize, rectBorder.bottom);
			break;
			
		case BCGPPopupWindowStemLocation_Left:
			rectBorder.left += m_nStemSize;
			rectExclude = CRect(rectBorder.left, rectBorder.CenterPoint().y - m_nStemSize, rectBorder.left + 2, rectBorder.CenterPoint().y + m_nStemSize);
			break;
			
		case BCGPPopupWindowStemLocation_Right:
			rectBorder.right -= m_nStemSize;
			rectExclude = CRect(rectBorder.right - 2, rectBorder.CenterPoint().y - m_nStemSize, rectBorder.right, rectBorder.CenterPoint().y + m_nStemSize);
			break;
		}

		pDC->ExcludeClipRect(rectExclude);
	}

	if (m_Colors.clrBorder == (COLORREF)-1 || globalData.IsHighContastMode())
	{
		if (m_nCornerRadius > 0)
		{
			CBCGPVisualManager::GetInstance ()->OnDrawPopupWindowRoundedBorder (pDC, rectBorder, this, m_nCornerRadius);
		}
		else
		{
			CBCGPVisualManager::GetInstance ()->OnDrawPopupWindowBorder (pDC, rectBorder);
		}
	}
	else
	{
		if (m_nCornerRadius > 0)
		{
			CPen pen (PS_SOLID, 1, m_Colors.clrBorder);
			CPen* pOldPen = pDC->SelectObject (&pen);
			CBrush* pOldBrush = (CBrush*) pDC->SelectStockObject (NULL_BRUSH);
			
			pDC->RoundRect (rectBorder.left, rectBorder.top, rectBorder.right, rectBorder.bottom, m_nCornerRadius + 1, m_nCornerRadius + 1);
			
			pDC->SelectObject (pOldPen);
			pDC->SelectObject (pOldBrush);
		}
		else
		{
			pDC->Draw3dRect(rectBorder, m_Colors.clrBorder, m_Colors.clrBorder);
		}
	}

	if (stemLocation != BCGPPopupWindowStemLocation_None)
	{
		pDC->SelectClipRgn(NULL);

		COLORREF clrBorder = m_Colors.clrBorder;
		if (clrBorder == (COLORREF)-1 || globalData.IsHighContastMode())
		{
			clrBorder = CBCGPVisualManager::GetInstance ()->GetPopupWindowBorderBorderColor();
		}

		POINT points[3];
		GetStemPoints(points, TRUE);

		CBCGPPenSelector ps(*pDC, clrBorder);
		pDC->Polyline(points, 3);
	}

	if (m_bSmallCaption)
	{
		return;
	}

	//--------------------
	// Draw icon and name:
	//--------------------
	CRect rectText = rectCaption;
	int nHorzMargin = rectCaption.Height() / 4;
	rectText.left += nHorzMargin;

	CWnd* pWndBtn = NULL;

	if (m_btnMenu.GetSafeHwnd () != NULL)
	{
		pWndBtn = &m_btnMenu;
	}
	else if (m_btnPin.GetSafeHwnd () != NULL)
	{
		pWndBtn = &m_btnPin;
	}
	else if (m_btnClose.GetSafeHwnd () != NULL)
	{
		pWndBtn = &m_btnClose;
	}

	if (pWndBtn != NULL)
	{
		CRect rectBtn;
		pWndBtn->GetWindowRect (&rectBtn);
		ScreenToClient (&rectBtn);

		rectText.right = rectBtn.left - nImageMargin;
	}

	HICON hIcon = GetIcon (FALSE);
	if (hIcon != NULL)
	{
		int nImagePadding = globalUtils.ScaleByDPI(m_bLargeCaptionFont ? 2 * nImageMargin : nImageMargin);

		CRect rectIcon(
			CPoint(
				rectCaption.left + nImagePadding, 
				rectCaption.top + max(0, (int)(0.5 + 0.5 * (rectCaption.Height() - globalData.m_sizeSmallIcon.cy)))),
			globalData.m_sizeSmallIcon);

		::DrawIconEx(pDC->GetSafeHdc (), 
			rectIcon.left,
			rectIcon.top,
			hIcon, globalData.m_sizeSmallIcon.cx, globalData.m_sizeSmallIcon.cy, 0, NULL,
			DI_NORMAL);

		rectText.left = rectIcon.right + nImagePadding;
	}

	CString strText;
	GetWindowText(strText);

	if (!strText.IsEmpty())
	{
		COLORREF clrTextOld = pDC->SetTextColor (clrText);
		pDC->SetBkMode (TRANSPARENT);

#ifndef _BCGSUITE_
		CFont* pOldFont = pDC->SelectObject(m_bLargeCaptionFont ? &globalData.fontCaption : &globalData.fontRegular);
#else
		CFont* pOldFont = pDC->SelectObject(&globalData.fontRegular);
#endif

		pDC->DrawText (strText, rectText, DT_LEFT | DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);

		pDC->SelectObject (pOldFont);
		pDC->SetTextColor (clrTextOld);
	}
}
//*******************************************************************************
BOOL CBCGPPopupWindow::OnDrawButtonBorder(CDC* pDC, CBCGPPopupWndButton* pButton, const CRect& rect, UINT /*nState*/)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	if (m_Theme == BCGPPopupWindowTheme_VisualManager || globalData.IsHighContastMode())
	{
		return FALSE;
	}

	COLORREF clrBorder = pButton->IsHighlighted() || pButton->IsPressed() ? m_Colors.clrBorder : m_Colors.clrFill;
	pDC->Draw3dRect(rect, clrBorder, clrBorder);

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPPopupWindow::OnFillButton(CDC* pDC, CBCGPPopupWndButton* pButton, const CRect& rect)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	if (m_Theme == BCGPPopupWindowTheme_VisualManager || globalData.IsHighContastMode())
	{
		return FALSE;
	}

	COLORREF clrFill = pButton->IsPressed() ? m_Colors.clrPressedButton : m_Colors.clrFill;
	pDC->FillSolidRect(rect, clrFill);

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPPopupWindow::OnFillBackground(CDC* pDC, CBCGPPopupDlg* /*pDlg*/, const CRect& rect)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);
	
	if (m_Theme == BCGPPopupWindowTheme_VisualManager || globalData.IsHighContastMode())
	{
		return FALSE;
	}
	
	pDC->FillSolidRect(rect, m_Colors.clrFill);
	return TRUE;
}
//*******************************************************************************
COLORREF CBCGPPopupWindow::GetTextColor(CBCGPPopupDlg* /*pDlg*/, BOOL bCaption, BOOL bButton)
{
	ASSERT_VALID(this);

	if (!bCaption)
	{
		return m_Colors.clrText != (COLORREF)-1 ? m_Colors.clrText : CBCGPVisualManager::GetInstance ()->GetPopupWindowTextColor(this);
	}

	if (m_Colors.clrFill == (COLORREF)-1 || globalData.IsHighContastMode())
	{
		return CBCGPVisualManager::GetInstance ()->GetPopupWindowCaptionTextColor(this, bButton);
	}

	return m_Colors.clrText;
}
//*******************************************************************************
COLORREF CBCGPPopupWindow::GetLinkTextColor(CBCGPPopupDlg* /*pDlg*/, BOOL bIsHover)
{
	ASSERT_VALID(this);

	COLORREF clr = bIsHover ? m_Colors.clrHoverLink : m_Colors.clrLink;
	if (clr != (COLORREF)-1 && !globalData.IsHighContastMode())
	{
		return clr;
	}
	
	return CBCGPVisualManager::GetInstance ()->GetPopupWindowLinkTextColor(this, bIsHover);
}
//*******************************************************************************
void CBCGPPopupWindow::StartAnimation (BOOL bShow/* = TRUE*/)
{
	const CBCGPPopupMenu::ANIMATION_TYPE animationType = GetActualAnimationType();

	if (animationType == CBCGPPopupMenu::NO_ANIMATION ||
		globalData.bIsRemoteSession ||
		(animationType == CBCGPPopupMenu::FADE && globalData.m_nBitsPerPixel <= 8))
	{
		if (!bShow)
		{
			SendMessage (WM_CLOSE);
			return;
		}

		if (m_btnPin.GetSafeHwnd () != NULL)
		{
			m_btnPin.ShowWindow (SW_SHOWNOACTIVATE);
		}

		if (m_btnClose.GetSafeHwnd () != NULL)
		{
			m_btnClose.ShowWindow (SW_SHOWNOACTIVATE);
		}

		if (m_btnMenu.GetSafeHwnd () != NULL)
		{
			m_btnMenu.ShowWindow (SW_SHOWNOACTIVATE);
		}

#ifndef _BCGSUITE_
		if (m_pWndShadow->GetSafeHwnd() != NULL)
		{
			m_pWndShadow->Repos();
			m_bShadowIsReady = TRUE;
		}

		BOOL bIsWindowsLayerSupportAvailable = globalData.IsWindowsLayerSupportAvailable();
#else
		BOOL bIsWindowsLayerSupportAvailable = TRUE;
#endif

		m_bAnimationIsDone = TRUE;

		if (bIsWindowsLayerSupportAvailable && globalData.m_nBitsPerPixel > 8 &&
			m_nTransparency < 255)
		{
			ModifyStyleEx(0, WS_EX_LAYERED);
		}

		if (m_nAutoCloseTime > 0)
		{
			SetTimer (iClosePopupTimerId, m_nAutoCloseTime, NULL);
		}

		m_pWndDlg->SetWindowPos (NULL,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_NOREDRAW|SWP_NOZORDER|SWP_SHOWWINDOW | SWP_NOACTIVATE);
		m_pWndDlg->ValidateRect (NULL);

		if (bIsWindowsLayerSupportAvailable && globalData.m_nBitsPerPixel > 8 &&
			m_nTransparency < 255)
		{
			globalData.SetLayeredAttrib (GetSafeHwnd (), 0, m_nTransparency, LWA_ALPHA);

#ifndef _BCGSUITE_
			if (m_pWndShadow->GetSafeHwnd () != NULL)
			{
				m_pWndShadow->UpdateTransparency (m_nTransparency);
			}
#endif
		}

		return;
	}

	m_bAnimationIsDone = FALSE;
	m_bFadeOutAnimation = !bShow;

	//-------------------------
	// Set animation direction:
	//-------------------------
	CRect rectScreen;
	::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectScreen, 0);

	CRect rectWindow;
	GetWindowRect (&rectWindow);

	m_bIsAnimRight = rectWindow.left < rectScreen.CenterPoint ().x;
	m_bIsAnimDown = rectWindow.top < rectScreen.CenterPoint ().y;

	m_iFadePercent = m_bFadeOutAnimation ? 100 : 0;
	m_iFadeStep = m_bFadeOutAnimation ? -5 : 5;
	
	if (m_FinalSize.cx > m_FinalSize.cy)
	{
		m_nAnimStepY = 10;
		m_nAnimStepX = max (1, m_nAnimStepY * m_FinalSize.cx / m_FinalSize.cy);
	}
	else
	{
		m_nAnimStepX = 10;
		m_nAnimStepY = max (1, m_nAnimStepX * m_FinalSize.cy / m_FinalSize.cx);
	}

	//--------------------------
	// Adjust initial menu size:
	//--------------------------
	m_AnimSize = m_FinalSize;

	if (bShow)
	{
		switch (animationType)
		{
		case CBCGPPopupMenu::UNFOLD:
			m_AnimSize.cx = m_nAnimStepX;

		case CBCGPPopupMenu::SLIDE:
			m_AnimSize.cy = m_nAnimStepY;
			break;
		}
	}

	if (m_pWndDlg != NULL && m_pWndDlg->IsWindowVisible ())
	{
		m_pWndDlg->ShowWindow (SW_HIDE);
	}

	SetTimer (iAnimTimerId, m_AnimationSpeed, NULL);
	nLastAnimTime = clock ();
}
//*******************************************************************************
void CBCGPPopupWindow::DrawAnimation (CDC* pPaintDC)
{
	CRect rectClient;
	GetClientRect (&rectClient);

	int cx = m_FinalSize.cx;
	int cy = m_FinalSize.cy;

	CDC dcMem;
	if (!dcMem.CreateCompatibleDC (pPaintDC))
	{
		return;
	}

	const CBCGPPopupMenu::ANIMATION_TYPE animationType = GetActualAnimationType();

	// create the three bitmaps if not done yet
	if (m_bmpScreenDst.GetSafeHandle() == NULL)
	{
		CBitmap* pBmpOld = NULL;

		if (animationType == CBCGPPopupMenu::FADE || globalData.m_nBitsPerPixel > 8)
		{
			// Fill in the BITMAPINFOHEADER
			BITMAPINFOHEADER bih;
			bih.biSize = sizeof(BITMAPINFOHEADER);
			bih.biWidth = cx;
			bih.biHeight = cy;
			bih.biPlanes = 1;
			bih.biBitCount = 32;
			bih.biCompression = BI_RGB;
			bih.biSizeImage = cx * cy;
			bih.biXPelsPerMeter = 0;
			bih.biYPelsPerMeter = 0;
			bih.biClrUsed = 0;
			bih.biClrImportant = 0;

			HBITMAP hmbpDib;
			// Create a DIB section and attach it to the source bitmap
			hmbpDib = CreateDIBSection (
				dcMem.m_hDC, (LPBITMAPINFO)&bih, DIB_RGB_COLORS, (void **)&m_cFadeSrcBits,
				NULL, NULL);
			if (hmbpDib == NULL || m_cFadeSrcBits == NULL)
			{
				return;
			}

			m_bmpScreenSrc.Attach( hmbpDib );

			// Create a DIB section and attach it to the destination bitmap
			hmbpDib = CreateDIBSection (
				dcMem.m_hDC, (LPBITMAPINFO)&bih, DIB_RGB_COLORS, (void **)&m_cFadeDstBits,
				NULL, NULL);
			if (hmbpDib == NULL || m_cFadeDstBits == NULL)
			{
				return;
			}
			m_bmpScreenDst.Attach( hmbpDib );

			// Create a DIB section and attach it to the temporary bitmap
			hmbpDib = CreateDIBSection (
				dcMem.m_hDC, (LPBITMAPINFO)&bih, DIB_RGB_COLORS, (void **)&m_cFadeTmpBits,
				NULL, NULL);
			if (hmbpDib == NULL || m_cFadeTmpBits == NULL)
			{
				return;
			}

			m_bmpScreenTmp.Attach( hmbpDib );

			// get source image, representing the window below the popup menu
			pBmpOld = dcMem.SelectObject (&m_bmpScreenSrc);
			dcMem.BitBlt (0, 0, cx, cy, pPaintDC, rectClient.left, rectClient.top, SRCCOPY);

			// copy it to the destination so that shadow will be ok
			memcpy (m_cFadeDstBits, m_cFadeSrcBits, sizeof(COLORREF)* cx*cy);
			dcMem.SelectObject (&m_bmpScreenDst);
		}
		else
		{
			m_bmpScreenDst.CreateCompatibleBitmap (pPaintDC, cx, cy);
			pBmpOld = dcMem.SelectObject (&m_bmpScreenDst);
		}

#ifndef _BCGSUITE_
		BOOL bIsWindowsLayerSupportAvailable = globalData.IsWindowsLayerSupportAvailable();
#else
		BOOL bIsWindowsLayerSupportAvailable = TRUE;
#endif

		if (bIsWindowsLayerSupportAvailable && globalData.m_nBitsPerPixel > 8 &&
			m_nTransparency < 255)
		{
			ModifyStyleEx(0, WS_EX_LAYERED);
			globalData.SetLayeredAttrib (GetSafeHwnd (), 0, m_nTransparency, LWA_ALPHA);
		}

		SendMessage (WM_PRINT, (WPARAM) dcMem.GetSafeHdc (), 
			PRF_CHILDREN | PRF_CLIENT | PRF_ERASEBKGND);
		
		if (m_btnPin.GetSafeHwnd () != NULL)
		{
			m_btnPin.ShowWindow (SW_HIDE);
		}
		
		if (m_btnClose.GetSafeHwnd () != NULL)
		{
			m_btnClose.ShowWindow (SW_HIDE);
		}

		if (m_btnMenu.GetSafeHwnd () != NULL)
		{
			m_btnMenu.ShowWindow (SW_HIDE);
		}
		
		CRect rect;
		m_pWndDlg->GetWindowRect (&rect);
		ScreenToClient (&rect);
		
		dcMem.SetViewportOrg (rect.TopLeft());

		m_pWndDlg->SendMessage (WM_PRINT, (WPARAM) dcMem.GetSafeHdc (), 
			PRF_CHILDREN | PRF_CLIENT | PRF_ERASEBKGND);

		dcMem.SetViewportOrg (CPoint (0,0));

		dcMem.SelectObject (pBmpOld);
	}
	
	COLORREF *src = m_cFadeSrcBits;
	COLORREF *dst = m_cFadeDstBits;
	COLORREF *tmp = m_cFadeTmpBits;

	CBitmap* pBmpOld = NULL;

	switch (animationType)
	{
	case CBCGPPopupMenu::UNFOLD:
	case CBCGPPopupMenu::SLIDE:
		pBmpOld = dcMem.SelectObject (&m_bmpScreenDst);

		pPaintDC->BitBlt (
			m_bIsAnimRight ? rectClient.left : rectClient.right - m_AnimSize.cx,
			m_bIsAnimDown ? rectClient.top : rectClient.bottom - m_AnimSize.cy,
			m_AnimSize.cx, m_AnimSize.cy, &dcMem, 0, 0, SRCCOPY);
		break;

	case CBCGPPopupMenu::FADE:
		pBmpOld = dcMem.SelectObject (&m_bmpScreenTmp);
		for (int pixel = 0; pixel < cx * cy; pixel++)
		{
			*tmp++ = CBCGPDrawManager::PixelAlpha (*src++, *dst++, 100 - m_iFadePercent);
		}

		pPaintDC->BitBlt (rectClient.left, rectClient.top, cx, cy, 
						&dcMem, 0, 0, SRCCOPY);
	}

	dcMem.SelectObject (pBmpOld);
}
//*******************************************************************************
int CBCGPPopupWindow::GetCaptionHeight ()
{
	if (m_bSmallCaption)
	{
		return globalUtils.ScaleByDPI(nSmallCaptionHeight);
	}
	else
	{
		int nButtonHeight = CBCGPMenuImages::Size().cy;
#ifndef _BCGSUITE_
		int nTextHeight = m_bLargeCaptionFont ? globalData.GetCaptionTextHeight() : globalData.GetTextHeight();
#else
		int nTextHeight = globalData.GetTextHeight();
#endif

		int nVertPadding = globalUtils.ScaleByDPI(8);
		return max (nButtonHeight, nTextHeight) + nVertPadding;
	}
}
//*******************************************************************************
void CBCGPPopupWindow::OnClose() 
{
	if (m_pWndOwner->GetSafeHwnd () != NULL)
	{
		m_pWndOwner->SendMessage (BCGM_ON_CLOSEPOPUPWINDOW, 0, (LPARAM) this);
	}

	CWnd::OnClose();
}
//*******************************************************************************
void CBCGPPopupWindow::OnWindowPosChanged(WINDOWPOS FAR* lpwndpos)
{
	CWnd::OnWindowPosChanged(lpwndpos);

#ifndef _BCGSUITE_
	if (m_pWndShadow->GetSafeHwnd () != NULL)
	{
		if (lpwndpos->flags & SWP_HIDEWINDOW)
		{
			m_pWndShadow->ShowWindow (SW_HIDE);
		}
		else
		{
			if (m_StemLocation != BCGPPopupWindowStemLocation_None)
			{
				POINT points[3];
				GetStemPoints(points);

				m_pWndShadow->SetStemRegion(points, 3);
			}

			m_pWndShadow->Repos ();
			m_bShadowIsReady = TRUE;
		}
	}
#endif
}
//*******************************************************************************
void CBCGPPopupWindow::StartWindowMove ()
{
	m_bMoving = TRUE;
	GetCursorPos (&m_ptStartMove);

	SetCapture();
	KillTimer (iClosePopupTimerId);
}
//*******************************************************************************
void CBCGPPopupWindow::OnMouseMove(UINT nFlags, CPoint point) 
{
	if (m_bMoving)
	{
		// ---------------------
		// Calc screen rectangle
		// ---------------------
		CRect rectScreen;

		CPoint ptCursor = point;
		ClientToScreen (&ptCursor);

		MONITORINFO mi;
		mi.cbSize = sizeof (MONITORINFO);
		if (GetMonitorInfo (MonitorFromPoint (ptCursor, MONITOR_DEFAULTTONEAREST), &mi))
		{
			rectScreen = mi.rcWork;
		}
		else
		{
			::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectScreen, 0);
		}

		CPoint ptMouse;
		GetCursorPos (&ptMouse);
		
		CPoint ptOffset = ptMouse - m_ptStartMove;
		m_ptStartMove = ptMouse;
		
		CRect rect;
		GetWindowRect (&rect);
		rect.OffsetRect (ptOffset);
		
		if (rect.left < rectScreen.left)
		{
			rect.OffsetRect (rectScreen.left - rect.left, 0);
		}
		else if (rect.right > rectScreen.right)
		{
			rect.OffsetRect (rectScreen.right - rect.right, 0);
		}

		if (rect.top < rectScreen.top)
		{
			rect.OffsetRect (0, rectScreen.top - rect.top);
		}
		else if (rect.bottom > rectScreen.bottom)
		{
			rect.OffsetRect (0, rectScreen.bottom - rect.bottom);
		}

		SetWindowPos (NULL, rect.left, rect.top, -1, -1,
			SWP_NOZORDER | SWP_NOSIZE | SWP_SHOWWINDOW);
	}
	
	CWnd::OnMouseMove(nFlags, point);
}
//*******************************************************************************
void CBCGPPopupWindow::OnLButtonUp(UINT nFlags, CPoint point) 
{
	if (m_bMoving)
	{
		ReleaseCapture ();
		m_bMoving = FALSE;

		if (m_nAutoCloseTime > 0)
		{
			SetTimer (iClosePopupTimerId, m_nAutoCloseTime, NULL);
		}

		CRect rectWnd;
		GetWindowRect (rectWnd);

		m_ptLastPos = rectWnd.TopLeft ();
	}
	
	CWnd::OnLButtonUp(nFlags, point);
}
//*******************************************************************************
void CBCGPPopupWindow::OnCancelMode() 
{
	CWnd::OnCancelMode();
	
	if (m_bMoving)
	{
		ReleaseCapture ();
		m_bMoving = FALSE;

		if (m_nAutoCloseTime > 0)
		{
			SetTimer (iClosePopupTimerId, m_nAutoCloseTime, NULL);
		}

		CRect rectWnd;
		GetWindowRect (rectWnd);

		m_ptLastPos = rectWnd.TopLeft ();
	}
}
//*******************************************************************************
void CBCGPPopupWindow::OnLButtonDown(UINT /*nFlags*/, CPoint /*point*/) 
{
	StartWindowMove ();
}
//*******************************************************************************
void CBCGPPopupWindow::OnSetFocus(CWnd* pOldWnd) 
{
	CWnd::OnSetFocus(pOldWnd);
	
	if (m_pWndDlg->GetSafeHwnd() != NULL)
	{
		m_pWndDlg->SetFocus();
	}
}

#endif // BCGP_EXCLUDE_POPUP_WINDOW
