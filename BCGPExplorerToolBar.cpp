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
// BCGPExplorerToolBar.cpp : implementation file
//

#include "stdafx.h"

#include "BCGCBPro.h"
#include "BCGPExplorerToolBar.h"

#ifndef _BCGSUITE_
#include "BCGPFrameWnd.h"
#include "BCGPMDIFrameWnd.h"
#include "BCGPToolBarButton.h"
#include "BCGPToolbarMenuButton.h"
#include "BCGPContextMenuManager.h"
#include "CustomizeButton.h"
#include "MenuImages.h"
#else
#include "afxcustomizebutton.h"
#endif

#include "BCGPLocalResource.h"
#include "BCGPVisualManager.h"
#include "BCGPDlgImpl.h"
#include "BCGPGlobalUtils.h"
#include "BCGPDrawManager.h"
#include "BCGProRes.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static const int nHorzMargin = 2;

#ifndef _BCGSUITE_
#define visualManager			CBCGPVisualManager
#define visualManagerInstance	CBCGPVisualManager::GetInstance ()
#else
#define visualManager			CMFCVisualManager
#define visualManagerInstance	CMFCVisualManager::GetInstance ()
#endif


/////////////////////////////////////////////////////////////////////////////
// CBCGPExplorerNavigationButton

enum BCGPNavButtonType
{
	BCGPNav_Home,
	BCGPNav_Back,
	BCGPNav_Forward,
	BCGPNav_Up,
	BCGPNav_Margin,
};

class CBCGPExplorerNavigationButton : public CBCGPToolbarButton
{
	DECLARE_DYNCREATE(CBCGPExplorerNavigationButton)

public:
	CBCGPExplorerNavigationButton (UINT nCmdID = 0, BCGPNavButtonType type = BCGPNav_Margin) :
		CBCGPToolbarButton (nCmdID, -1)
	{
		m_Type = type;
		m_xPreferredWidth = 0;
	}

	CBCGPExplorerNavigationButton (int xPreferredWidth) :
		CBCGPToolbarButton (0, -1)
	{
		m_Type = BCGPNav_Margin;
		m_xPreferredWidth = xPreferredWidth;
	}

	virtual void CopyFrom (const CBCGPToolbarButton& src);
	virtual void OnDraw (CDC* pDC, const CRect& rect, CBCGPToolBarImages* pImages,
						BOOL bHorz = TRUE, BOOL bCustomizeMode = FALSE,
						BOOL bHighlight = FALSE,
						BOOL bDrawBorder = TRUE,
						BOOL bGrayDisabledButtons = TRUE);
	virtual SIZE OnCalculateSize (CDC* pDC, const CSize& sizeDefault, BOOL bHorz);
	
	virtual BOOL IsExplorerNavigationButton() const
	{
		return TRUE;
	}

protected:
	BCGPNavButtonType	m_Type;
	int					m_xPreferredWidth;
};

IMPLEMENT_DYNCREATE(CBCGPExplorerNavigationButton, CBCGPToolbarButton)

void CBCGPExplorerNavigationButton::CopyFrom (const CBCGPToolbarButton& s)
{
	CBCGPToolbarButton::CopyFrom (s);

	const CBCGPExplorerNavigationButton& src = (const CBCGPExplorerNavigationButton&) s;

	m_Type = src.m_Type;
	m_xPreferredWidth = src.m_xPreferredWidth;
}

void CBCGPExplorerNavigationButton::OnDraw (CDC* pDC, const CRect& rect, CBCGPToolBarImages* /*pImages*/,
						BOOL /*bHorz*/, BOOL /*bCustomizeMode*/,
						BOOL bHighlight,
						BOOL /*bDrawBorder*/,
						BOOL /*bGrayDisabledButtons*/)
{
	ASSERT_VALID (this);

	if (m_Type == BCGPNav_Margin)
	{
		return;
	}

	CBCGPExplorerToolBar* pToolBar = DYNAMIC_DOWNCAST(CBCGPExplorerToolBar, m_pWndParent);
	if (pToolBar == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	BOOL bIsForward = m_Type == BCGPNav_Forward;
	BOOL bIsUp = m_Type == BCGPNav_Up;
	BOOL bIsHome = m_Type == BCGPNav_Home;

	if (!pToolBar->m_bSimplifiedNavigationButtons && !bIsUp)
	{
		pToolBar->OnDrawNavButton (pDC, this, bIsForward, rect, bHighlight);
		return;
	}

	FillInterior (pDC, rect, bHighlight);

#ifndef _BCGSUITE_
	visualManager::BCGBUTTON_STATE state = visualManager::ButtonsIsRegular;
#else
	visualManager::AFX_BUTTON_STATE state = visualManager::ButtonsIsRegular;
#endif
	
	if (bHighlight)
	{
		if (m_nStyle & (TBBS_PRESSED | TBBS_CHECKED))
		{
			//-----------------------
			// Pressed in or checked:
			//-----------------------
			state = visualManager::ButtonsIsPressed;
		}
		else
		{
			state = visualManager::ButtonsIsHighlighted;
		}
	}

	CRect rectText = rect;

#ifndef _BCGSUITE_
	CBCGPMenuImages::IMAGE_STATE stateImage = (CBCGPMenuImages::IMAGE_STATE)CBCGPVisualManager::GetInstance()->GetNavImageState(
		bHighlight, bHighlight && (m_nStyle & (TBBS_PRESSED | TBBS_CHECKED)), m_nStyle & TBBS_DISABLED, pToolBar->IsOnGlass());

	CSize sizeImage(0, 0);

	if (!pToolBar->m_bSimplifiedNavigationButtons && bIsHome)
	{
		sizeImage = globalUtils.ScaleByDPI(CSize(16, 16));
	}
		
	CBCGPMenuImages::Draw(pDC, bIsHome ? CBCGPMenuImages::IdExplorerHome : bIsUp ? CBCGPMenuImages::IdExplorerUp : bIsForward ? CBCGPMenuImages::IdExplorerForward : CBCGPMenuImages::IdExplorerBack, 
		rectText, stateImage, sizeImage);
#else
	COLORREF clrText = visualManagerInstance->GetToolbarButtonTextColor(this, state);
	if (bIsHome)
	{
		const int count = 7;
		POINT arPoints[count];

		int xCenter = rectText.CenterPoint().x;
		CSize sizeHome = globalUtils.ScaleByDPI(CSize(12, 9));

		rectText.left = xCenter - sizeHome.cx / 2;
		rectText.top = rectText.CenterPoint().y - sizeHome.cy / 2;
		rectText.right = rectText.left + sizeHome.cx;
		rectText.bottom = rectText.top + sizeHome.cy;

		int dx = rectText.Width() / 5;
		int yMid = rectText.top + rectText.Width() / 2;
		int yBottom = min(rectText.bottom, yMid + 2 * dx);

		arPoints[0] = CPoint(xCenter, rectText.top);
		arPoints[1] = CPoint(rectText.right, yMid);
		arPoints[2] = CPoint(rectText.right - dx, yMid);
		arPoints[3] = CPoint(rectText.right - dx, yBottom);
		arPoints[4] = CPoint(rectText.left + dx, yBottom);
		arPoints[5] = CPoint(rectText.left + dx, yMid);
		arPoints[6] = CPoint(rectText.left, yMid);

		CRgn rgn;
		rgn.CreatePolygonRgn(arPoints, count, ALTERNATE);
		
		CBrush br(clrText);
		pDC->FillRgn(&rgn, &br);
	}
	else
	{
		const CString strOut(bIsUp ? _T("^") : bIsForward ? _T(">") : _T("<"));

		pDC->SetTextColor(clrText);
		pDC->DrawText(strOut, rectText, DT_VCENTER | DT_SINGLELINE | DT_CENTER);
	}
#endif

	//--------------------
	// Draw button border:
	//--------------------
	visualManagerInstance->OnDrawButtonBorder (pDC, this, rect, state);
}

SIZE CBCGPExplorerNavigationButton::OnCalculateSize (CDC* /*pDC*/, const CSize& sizeDefault, BOOL /*bHorz*/)
{
	ASSERT_VALID (this);

	if (m_Type == BCGPNav_Margin)
	{
		return CSize (m_xPreferredWidth, 4);
	}

	if (m_nID == 0)
	{
		return CSize (0, 0);
	}

	CBCGPExplorerToolBar* pToolBar = DYNAMIC_DOWNCAST(CBCGPExplorerToolBar, m_pWndParent);
	if (pToolBar == NULL)
	{
		return CSize (0, 0);
	}

	ASSERT_VALID (pToolBar);

	BOOL bIsUp = m_Type == BCGPNav_Up;

	if (pToolBar->m_bSimplifiedNavigationButtons || bIsUp)
	{
		int nCheckSize = ::GetSystemMetrics (SM_CYMENUCHECK);
		return CSize(max(sizeDefault.cx, nCheckSize + nHorzMargin), max(sizeDefault.cy, nCheckSize));
	}

	const CSize sizeImage = pToolBar->GetNavImageSize ();

	int cx = sizeImage.cx;
	int cy = sizeImage.cy;

	if (m_Type == BCGPNav_Back)
	{
		cx += nHorzMargin;
	}
	else
	{
		cx += nHorzMargin + 1;
	}

	return CSize (cx, cy);
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPExplorerNavigationHistoryButton

class CBCGPExplorerNavigationHistoryButton : public CBCGPToolbarMenuButton
{
	DECLARE_DYNCREATE(CBCGPExplorerNavigationHistoryButton)

public:
	CBCGPExplorerNavigationHistoryButton (UINT nCmdID = 0, UINT uiMenuItemID = 0) :
		CBCGPToolbarMenuButton (nCmdID, NULL, -1)
	{
		m_listCommands.AddTail (new CBCGPToolbarMenuButton (uiMenuItemID, NULL, -1, _T("<HISTORY ITEM>")));
	}

	virtual void OnDraw (CDC* pDC, const CRect& rect, CBCGPToolBarImages* pImages,
						BOOL bHorz = TRUE, BOOL bCustomizeMode = FALSE,
						BOOL bHighlight = FALSE,
						BOOL bDrawBorder = TRUE,
						BOOL bGrayDisabledButtons = TRUE);
	virtual SIZE OnCalculateSize (CDC* pDC, const CSize& sizeDefault, BOOL bHorz);
	virtual BOOL OnClick (CWnd* pWnd, BOOL bDelay = TRUE);

	virtual const CRect GetInvalidateRect () const
	{
		CRect rect = CBCGPToolbarMenuButton::GetInvalidateRect ();
		rect.left -= 2 * nHorzMargin;

		return rect;
	}

	virtual BOOL IsExplorerNavigationButton() const
	{
		return TRUE;
	}
};

IMPLEMENT_DYNCREATE(CBCGPExplorerNavigationHistoryButton, CBCGPToolbarMenuButton)

void CBCGPExplorerNavigationHistoryButton::OnDraw (CDC* pDC, const CRect& rect, CBCGPToolBarImages* /*pImages*/,
						BOOL /*bHorz*/, BOOL /*bCustomizeMode*/,
						BOOL bHighlight,
						BOOL /*bDrawBorder*/,
						BOOL /*bGrayDisabledButtons*/)
{
	ASSERT_VALID (this);

	CBCGPExplorerToolBar* pToolBar = DYNAMIC_DOWNCAST(CBCGPExplorerToolBar, m_pWndParent);
	if (pToolBar == NULL)
	{
		return;
	}

	if (!pToolBar->m_bSimplifiedNavigationButtons && !globalData.IsHighContastMode())
	{
		return;
	}

	FillInterior (pDC, rect, bHighlight);
	
#ifndef _BCGSUITE_
	visualManager::BCGBUTTON_STATE state = visualManager::ButtonsIsRegular;
#else
	visualManager::AFX_BUTTON_STATE state = visualManager::ButtonsIsRegular;
#endif
	
	if (bHighlight || pToolBar->m_bIsHistoryDroppedDown)
	{
		if (m_nStyle & (TBBS_PRESSED | TBBS_CHECKED))
		{
			//-----------------------
			// Pressed in or checked:
			//-----------------------
			state = visualManager::ButtonsIsPressed;
		}
		else
		{
			state = visualManager::ButtonsIsHighlighted;
		}
	}
	
	CRect rectText = rect;
	
#ifndef _BCGSUITE_
	CBCGPMenuImages::IMAGE_STATE stateImage = (CBCGPMenuImages::IMAGE_STATE)CBCGPVisualManager::GetInstance()->GetNavImageState(
		bHighlight, bHighlight && (m_nStyle & (TBBS_PRESSED | TBBS_CHECKED)), m_nStyle & TBBS_DISABLED, pToolBar->IsOnGlass());

	CBCGPMenuImages::Draw(pDC, CBCGPMenuImages::IdArowDownLarge, rectText, stateImage);
#else
	CBCGPFontSelector fs(*pDC, &globalData.fontMarlett);

	pDC->SetTextColor(visualManagerInstance->GetToolbarButtonTextColor(this, state));

	CString strOut = _T("u");
	pDC->DrawText(strOut, rectText, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
#endif
	
	//--------------------
	// Draw button border:
	//--------------------
	visualManagerInstance->OnDrawButtonBorder (pDC, this, rect, state);
}

SIZE CBCGPExplorerNavigationHistoryButton::OnCalculateSize (CDC* /*pDC*/, const CSize& sizeDefault, BOOL /*bHorz*/)
{
	ASSERT_VALID (this);

	if (m_nID == 0)
	{
		return CSize (0, 0);
	}

	CBCGPExplorerToolBar* pToolBar = DYNAMIC_DOWNCAST(CBCGPExplorerToolBar, m_pWndParent);
	if (pToolBar == NULL)
	{
		return CSize (0, 0);
	}

	ASSERT_VALID (pToolBar);

	if (pToolBar->m_bSimplifiedNavigationButtons)
	{
		int nCheckSize = ::GetSystemMetrics (SM_CYMENUCHECK);
		return CSize(nCheckSize + 2 * nHorzMargin, max(sizeDefault.cy, nCheckSize));
	}

	const CSize sizeImage = pToolBar->GetNavImageSize ();
	const CSize sizeFrame = pToolBar->GetNavFrameSize ();

	return CSize (sizeFrame.cx - 2 * sizeImage.cx - nHorzMargin - 1, sizeImage.cy + 2);
}

BOOL CBCGPExplorerNavigationHistoryButton::OnClick (CWnd* pWnd, BOOL /*bDelay*/)
{
	ASSERT_VALID (this);

	CBCGPExplorerToolBar* pToolBar = DYNAMIC_DOWNCAST(CBCGPExplorerToolBar, m_pWndParent);
	if (pToolBar == NULL)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	pToolBar->m_bIsHistoryDroppedDown = TRUE;

	HMENU hMenu = CreateMenu ();

	CPoint point = CPoint (m_rect.left, m_rect.bottom - 1);
	pWnd->ClientToScreen (&point);

	UINT nMenuResult = 0;

#ifndef _BCGSUITE_
	if (g_pContextMenuManager != NULL)
	{
		nMenuResult = g_pContextMenuManager->TrackPopupMenu (hMenu, point.x, point.y, pWnd);
	}
#else
	CWinAppEx* pApp = DYNAMIC_DOWNCAST(CWinAppEx, AfxGetApp());
	CContextMenuManager* pMenuManager = (pApp != NULL) ? pApp->GetContextMenuManager () : NULL;

	if (pMenuManager != NULL)
	{
		nMenuResult = pMenuManager->TrackPopupMenu (hMenu, point.x, point.y, pWnd);
	}
#endif
	else
	{
		nMenuResult = ::TrackPopupMenu (hMenu, 
			TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_NONOTIFY | TPM_RETURNCMD, 
			point.x, point.y, 0, pWnd->GetSafeHwnd (), NULL);
	}

	::DestroyMenu (hMenu);

	pToolBar->m_bIsHistoryDroppedDown = FALSE;
	pToolBar->RedrawWindow (m_rect);

	if (nMenuResult != 0)
	{
		//-------------------------------------------------------
		// Trigger mouse up event (to button click notification):
		//-------------------------------------------------------
		CWnd* pParent = pToolBar->GetOwner ();
		if (pParent != NULL)
		{
			pParent->PostMessage (WM_COMMAND, nMenuResult);
		}
	}

	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPExplorerAddressButton

class CBCGPExplorerAddressButton : public CBCGPToolbarButton  
{
	DECLARE_SERIAL(CBCGPExplorerAddressButton)

public:
	CBCGPExplorerAddressButton();
	CBCGPExplorerAddressButton (UINT uiId, HWND hWndAddressBar, int nMinWidth);

	virtual ~CBCGPExplorerAddressButton();

// Overrides:
	virtual SIZE OnCalculateSize (CDC* pDC, const CSize& sizeDefault, BOOL bHorz);
	virtual void CopyFrom (const CBCGPToolbarButton& src);
	virtual void OnShow (BOOL bShow);
	virtual void OnMove ();
	virtual void OnSize (int iSize);

	virtual HWND GetHwnd ()
	{	
		return m_hWndAddressBar;
	}

	virtual BOOL CanBeStretched () const
	{	
		return TRUE;	
	}

	virtual BOOL HaveHotBorder () const
	{
		return FALSE;
	}

	virtual void OnDraw (CDC* /*pDC*/, const CRect& /*rect*/, CBCGPToolBarImages* /*pImages*/,
						BOOL /*bHorz*/ = TRUE, BOOL /*bCustomizeMode*/ = FALSE,
						BOOL /*bHighlight*/ = FALSE,
						BOOL /*bDrawBorder*/ = TRUE,
						BOOL /*bGrayDisabledButtons*/ = TRUE) {}

	virtual void SetStyle (UINT nStyle);

// Attributes:
public:
	HWND m_hWndAddressBar;
	int	m_nMinWidth;
};

IMPLEMENT_SERIAL(CBCGPExplorerAddressButton, CBCGPToolbarButton, 1)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPExplorerAddressButton::CBCGPExplorerAddressButton()
{
	m_hWndAddressBar = NULL;
	m_nMinWidth = 0;
}
//**************************************************************************************
CBCGPExplorerAddressButton::CBCGPExplorerAddressButton (UINT uiId, HWND hWndAddressBar, int nMinWidth) :
	CBCGPToolbarButton (uiId, -1)
{
	m_hWndAddressBar = hWndAddressBar;
	m_nMinWidth = nMinWidth;
}
//**************************************************************************************
CBCGPExplorerAddressButton::~CBCGPExplorerAddressButton()
{
}
//**************************************************************************************
SIZE CBCGPExplorerAddressButton::OnCalculateSize (CDC* /*pDC*/, const CSize& sizeDefault, BOOL /*bHorz*/)
{
	if (!IsVisible ())
	{
		if (m_hWndAddressBar != NULL)
		{
			::ShowWindow(m_hWndAddressBar, SW_HIDE);
		}

		return CSize (0,0);
	}

	if (m_hWndAddressBar != NULL && !IsHidden ())
	{
		::ShowWindow(m_hWndAddressBar, SW_SHOWNOACTIVATE);
	}

	return CSize (m_nMinWidth, sizeDefault.cy);
}
//**************************************************************************************
void CBCGPExplorerAddressButton::CopyFrom (const CBCGPToolbarButton& s)
{
	CBCGPToolbarButton::CopyFrom(s);

	const CBCGPExplorerAddressButton& src = (const CBCGPExplorerAddressButton&) s;

	m_hWndAddressBar = src.m_hWndAddressBar;
	m_nMinWidth = src.m_nMinWidth;
}
//**************************************************************************************
void CBCGPExplorerAddressButton::OnMove ()
{
	if (m_hWndAddressBar != NULL)
	{
		::SetWindowPos (m_hWndAddressBar, NULL,
			m_rect.left + 1, m_rect.top + 2, m_rect.Width () - 2, m_rect.Height () - 2,
			SWP_NOZORDER | SWP_NOACTIVATE);
	}
}
//**************************************************************************************
void CBCGPExplorerAddressButton::OnSize (int iSize)
{
	m_rect.right = m_rect.left + iSize;

	if (m_hWndAddressBar != NULL)
	{
		::SetWindowPos (m_hWndAddressBar, NULL,
			m_rect.left + 1, m_rect.top + 2, m_rect.Width () - 2, m_rect.Height () - 2,
			SWP_NOZORDER | SWP_NOACTIVATE);
	}
}
//****************************************************************************************
void CBCGPExplorerAddressButton::OnShow (BOOL bShow)
{
	if (m_hWndAddressBar != NULL)
	{
		if (bShow)
		{
			::ShowWindow(m_hWndAddressBar, SW_SHOWNOACTIVATE);
			OnMove ();
		}
		else
		{
			::ShowWindow(m_hWndAddressBar, SW_HIDE);
		}
	}
}
//****************************************************************************************
void CBCGPExplorerAddressButton::SetStyle (UINT nStyle)
{
	BOOL bWasDisabled = (m_nStyle & TBBS_DISABLED) != 0;
	BOOL bDisabled = (nStyle & TBBS_DISABLED) != 0;

	CBCGPToolbarButton::SetStyle (nStyle);

	if (bWasDisabled != bDisabled && m_hWndAddressBar != NULL)
	{
		::EnableWindow(m_hWndAddressBar, !bDisabled);
	}
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPExplorerToolBar

IMPLEMENT_DYNCREATE(CBCGPExplorerToolBar, CBCGPToolBar)

CBCGPExplorerToolBar::CBCGPExplorerToolBar()
{
	m_nBackID = 0;
	m_nForwardID = 0;
	m_nUpID = 0;
	m_nHomeID = 0;
	m_nHistoryID = 0;
	m_nHistoryItemID = 0;
	m_xRightMargin = 15;
	m_bSimplifiedNavigationButtons = FALSE;
	m_rectNavButtons.SetRectEmpty ();
	m_rectHistoryButton.SetRectEmpty ();
	m_bIsHistoryDroppedDown = FALSE;
	m_bIsAeroEnabled = TRUE;
	m_nStretchID = 0;
	m_nAddressBarID = 0;
	m_nAddressBarMinWidth = 0;
	m_hWndAddressBar = NULL;
	m_bIsHidden = FALSE;

#if (defined _BCGSUITE_) || (defined _BCGSUITE_INC_)
	m_bOnGlass = FALSE;
#endif
}

CBCGPExplorerToolBar::~CBCGPExplorerToolBar()
{
}


BEGIN_MESSAGE_MAP(CBCGPExplorerToolBar, CBCGPToolBar)
	//{{AFX_MSG_MAP(CBCGPExplorerToolBar)
	ON_WM_SHOWWINDOW()
	ON_WM_CREATE()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONDBLCLK()
	ON_WM_SETTINGCHANGE()
	//}}AFX_MSG_MAP
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CBCGPExplorerToolBar message handlers

BOOL CBCGPExplorerToolBar::LoadBitmap (UINT uiResID, UINT uiColdResID, UINT uiDisabledResID,
									   UINT uiLargeColdResID, UINT uiLargeDisabledResID, UINT uiLargeHotResID)
{
	CBCGPToolBarParams params;

	params.m_uiColdResID			= uiColdResID;
	params.m_uiHotResID				= uiResID;
	params.m_uiDisabledResID		= uiDisabledResID;
	params.m_uiLargeColdResID		= uiLargeColdResID;
	params.m_uiLargeDisabledResID	= uiLargeDisabledResID;
	params.m_uiLargeHotResID		= uiLargeHotResID == 0 ? uiLargeColdResID : uiLargeHotResID;

	return LoadBitmapEx (params, TRUE);
}
//********************************************************************************
BOOL CBCGPExplorerToolBar::LoadToolBar (UINT uiResID, UINT uiColdResID, UINT uiDisabledResID, UINT uiHotResID,
										UINT uiLargeColdResID, UINT uiLargeDisabledResID, UINT uiLargeHotResID)
{
	CBCGPToolBarParams params;

	params.m_uiColdResID			= uiColdResID;
	params.m_uiHotResID				= uiHotResID == 0 ? uiColdResID : uiHotResID;
	params.m_uiDisabledResID		= uiDisabledResID;
	params.m_uiLargeColdResID		= uiLargeColdResID;
	params.m_uiLargeDisabledResID	= uiLargeDisabledResID;
	params.m_uiLargeHotResID		= uiLargeHotResID == 0 ? uiLargeColdResID : uiLargeHotResID;

	return LoadToolBarEx (uiResID, params, TRUE);
}
//********************************************************************************
BOOL CBCGPExplorerToolBar::Create(CWnd* pParentWnd, DWORD dwStyle, UINT nID)
{
	return CreateEx(pParentWnd, 0, dwStyle, nID);
}
//********************************************************************************
BOOL CBCGPExplorerToolBar::CreateEx (CWnd* pParentWnd, DWORD /*dwCtrlStyle*/, DWORD dwStyle, 
							 UINT nID)
{
	ASSERT_VALID(pParentWnd);   // must have a parent

	m_dwStyle |= CBRS_HIDE_INPLACE;

	dwStyle &= ~CBRS_GRIPPER;

	// save the style
	SetBarAlignment (dwStyle & CBRS_ALL);

	// create the HWND
	CRect rect;
	rect.SetRectEmpty();

#ifndef _BCGSUITE_
	m_dwBCGStyle = 0; // can't float, resize, close, slide
#else
	m_dwControlBarStyle = 0; // can't float, resize, close, slide
#endif

	if (!CWnd::Create (globalData.RegisterWindowClass (_T("BCGPExplorerToolBar")),
		NULL, dwStyle | WS_CLIPSIBLINGS, rect, pParentWnd, nID))
	{
		return FALSE;
	}

#ifdef _BCGSUITE_
	if (pParentWnd->IsKindOf (RUNTIME_CLASS (CFrameWndEx)))
	{
		((CFrameWndEx*) pParentWnd)->AddControlBar (this);
	}
	else if (pParentWnd->IsKindOf (RUNTIME_CLASS (CMDIFrameWndEx)))
	{
		((CMDIFrameWndEx*) pParentWnd)->AddControlBar (this);
	}
#else
	if (pParentWnd->IsKindOf (RUNTIME_CLASS (CBCGPFrameWnd)))
	{
		((CBCGPFrameWnd*) pParentWnd)->AddControlBar (this);
	}
	else if (pParentWnd->IsKindOf (RUNTIME_CLASS (CBCGPMDIFrameWnd)))
	{
		((CBCGPMDIFrameWnd*) pParentWnd)->AddControlBar (this);
	}
#endif
	else
	{
		ASSERT (FALSE);
		return FALSE;
	}

	return TRUE;
}
//********************************************************************************
void CBCGPExplorerToolBar::EnableNavigationButtons (UINT nBackID, UINT nForwardID, 
												 UINT nHistoryID, UINT nHistoryItemID,
												 int xRightMargin, UINT nUpID, BOOL bSimplifiedNavigationButtons, UINT nHomeID)
{
	ASSERT_VALID (this);

	m_nBackID = nBackID;
	m_nForwardID = nForwardID;
	m_nHistoryID = nHistoryID;
	m_nHistoryItemID = nHistoryItemID;
	m_xRightMargin = xRightMargin;
	m_nUpID = nUpID;
	m_bSimplifiedNavigationButtons = bSimplifiedNavigationButtons;
	m_nHomeID = nHomeID;

	while (!m_Buttons.IsEmpty())
	{
		CBCGPToolbarButton* pButton = DYNAMIC_DOWNCAST(CBCGPToolbarButton, m_Buttons.GetHead());
		if (pButton == NULL)
		{
			break;
		}

#ifndef _BCGSUITE_
		if (!pButton->IsExplorerNavigationButton())
		{
			break;
		}
#else
		if (DYNAMIC_DOWNCAST(CBCGPExplorerNavigationButton, pButton) == NULL && DYNAMIC_DOWNCAST(CBCGPExplorerNavigationHistoryButton, pButton) == NULL)
		{
			break;
		}
#endif
		delete m_Buttons.RemoveHead();
	}
}
//********************************************************************************
void CBCGPExplorerToolBar::EnableAddressBar(UINT nID, HWND hwndAddressBar, BOOL bStretch, int nAddressBarMinWidth)
{
	ASSERT_VALID (this);
	
	if (m_hWndAddressBar != NULL && hwndAddressBar == NULL)
	{
		// Hide previously enabled address bar:
		int nIndex = CommandToIndex(m_nAddressBarID);
		if (nIndex >= 0)
		{
			::ShowWindow(m_hWndAddressBar, SW_HIDE);
			RemoveButton(nIndex);
		}
	}

	m_nStretchID = (hwndAddressBar != NULL && bStretch) ? nID : 0;
	m_nAddressBarID = hwndAddressBar != NULL ? nID : 0;
	m_hWndAddressBar = hwndAddressBar;
	m_nAddressBarMinWidth = nAddressBarMinWidth;
}
//********************************************************************************
void CBCGPExplorerToolBar::OnLoadNavImages ()
{
	CBCGPLocalResource locaRes;

	const CSize sizeImage (25, 25);
	const CSize sizeFrame (76, 29);

	const CSize sizeImageLarge (32, 32);
	const CSize sizeFrameLarge (87, 36);

	m_NavImages16.SetImageSize (sizeImage);
	m_NavImages16.SetTransparentColor (globalData.clrBtnFace);
	
	m_NavImages16.Load (IDB_BCGBARRES_NAV_BUTTONS_16);
	
	m_NavFrames16.SetImageSize (sizeFrame);
	m_NavFrames16.SetTransparentColor (globalData.clrBtnFace);
	m_NavFrames16.Load (IDB_BCGBARRES_NAV_FRAMES_16);

	m_NavImages.Clear();
	m_NavImages.SetImageSize (sizeImage);
	m_NavImages.Load (CBCGPVisualManager::GetInstance()->GetNavButtonsID(FALSE));
	globalUtils.ScaleByDPI(m_NavImages);
	
	m_NavFrames.Clear();
	m_NavFrames.SetImageSize (sizeFrame);
	m_NavFrames.Load (CBCGPVisualManager::GetInstance()->GetNavFrameID(FALSE));
	globalUtils.ScaleByDPI(m_NavFrames);

	m_NavImagesLarge.Clear();
	m_NavImagesLarge.SetImageSize (sizeImageLarge);
	m_NavImagesLarge.Load (CBCGPVisualManager::GetInstance()->GetNavButtonsID(TRUE));
	globalUtils.ScaleByDPI(m_NavImagesLarge);
	
	m_NavFramesLarge.Clear();
	m_NavFramesLarge.SetImageSize (sizeFrameLarge);
	m_NavFramesLarge.Load (CBCGPVisualManager::GetInstance()->GetNavFrameID(TRUE));
	globalUtils.ScaleByDPI(m_NavFramesLarge);
}
//**********************************************************************************
void CBCGPExplorerToolBar::AdjustLocations ()
{
	ASSERT_VALID(this);
	
	if (m_Buttons.IsEmpty () || GetSafeHwnd () == NULL)
	{
		return;
	}

	m_rectNavButtons.SetRectEmpty ();
	m_rectHistoryButton.SetRectEmpty ();

	CRect rectClient;
	GetClientRect (rectClient);

	if (m_nStretchID == 0)
	{
		CBCGPToolBar::AdjustLocations ();

		// Center buttons vertically:
		for (POSITION pos = m_Buttons.GetHeadPosition (); pos != NULL;)
		{
			CBCGPToolbarButton* pButton = (CBCGPToolbarButton*) m_Buttons.GetNext (pos);
			if (pButton == NULL || pButton == (CBCGPToolbarButton*)m_pCustomizeBtn)
			{
				break;
			}

			ASSERT_VALID (pButton);

			if (pButton->IsVisible ())
			{
				CRect rectButton = pButton->Rect ();
				int nButtonHeight = rectButton.Height ();

				if (rectClient.Height () > nButtonHeight && nButtonHeight > 0)
				{
					rectButton.top = rectClient.top + (rectClient.Height () - nButtonHeight) / 2;
					rectButton.bottom = rectButton.top + nButtonHeight;

					pButton->SetRect (rectButton);
				}
			}
		}
	}
	else
	{
		CClientDC dc (this);
		CFont* pOldFont = SelectDefaultFont (&dc);
		ASSERT (pOldFont != NULL);

		int nStretchButtonIndex = -1;
		POSITION pos = NULL;
		int nIndex = 0;

		CSize sizeGrid (GetColumnWidth (), GetRowHeight ());

		// First, calculate the free space:
		int cxTotalWidth = 0;

		for (pos = m_Buttons.GetHeadPosition (); pos != NULL; nIndex++)
		{
			CBCGPToolbarButton* pButton = (CBCGPToolbarButton*) m_Buttons.GetNext (pos);
			if (pButton == NULL || pButton == (CBCGPToolbarButton*)m_pCustomizeBtn)
			{
				break;
			}

			ASSERT_VALID (pButton);

			if (pButton->IsVisible ())
			{
				if (m_nStretchID != 0 && m_nStretchID == pButton->m_nID)
				{
					nStretchButtonIndex = nIndex;
				}

				cxTotalWidth += pButton->OnCalculateSize (&dc, sizeGrid, TRUE).cx;
			}
		}

		int cxFreeSpace = max (0, rectClient.Width () - cxTotalWidth);

		int x = rectClient.left;
		nIndex = 0;

		CRect rectRedraw = rectClient;

		// Repos buttons:
		for (pos = m_Buttons.GetHeadPosition (); pos != NULL; nIndex++)
		{
			CBCGPToolbarButton* pButton = (CBCGPToolbarButton*) m_Buttons.GetNext (pos);
			if (pButton == NULL || pButton == (CBCGPToolbarButton*)m_pCustomizeBtn)
			{
				break;
			}

			ASSERT_VALID (pButton);

			CRect rectButton (0, 0, 0, 0);

			if (pButton->IsVisible ())
			{
				CSize sizeButton = pButton->OnCalculateSize (&dc, sizeGrid, TRUE);
				if (pButton->m_bTextBelow)
				{
					sizeButton.cy = sizeGrid.cy;
				}

				if (nIndex == nStretchButtonIndex)
				{
					sizeButton.cx += cxFreeSpace;
					rectRedraw.left = x;
				}

				rectButton.left = x;
				rectButton.right = rectButton.left + sizeButton.cx;
				rectButton.top = rectClient.top + max (0, (rectClient.Height () - sizeButton.cy) / 2);
				rectButton.bottom = rectButton.top + sizeButton.cy;
				
				x += sizeButton.cx;
			}

			pButton->SetRect (rectButton);
		}
		
		dc.SelectObject (pOldFont);

		if (!rectRedraw.IsRectEmpty ())
		{
			RedrawWindow (rectRedraw);
		}
	}

	if (m_nBackID != 0)
	{
		CRect rectBack (0, 0, 0, 0);
		GetItemRect(CommandToIndex (m_nBackID), rectBack);

		m_rectNavButtons = rectBack;

		if (m_nForwardID != 0)
		{
			CRect rectForward (0, 0, 0, 0);
			GetItemRect(CommandToIndex (m_nForwardID), rectForward);

			m_rectNavButtons.right = rectForward.right + nHorzMargin;

			if (m_nHistoryID != 0)
			{
				GetItemRect(CommandToIndex (m_nHistoryID), m_rectHistoryButton);
			}
		}

		UpdateTooltips ();
	}

	if (m_pCustomizeBtn != NULL)
	{
		ASSERT_VALID (m_pCustomizeBtn);
		m_pCustomizeBtn->SetPipeStyle (FALSE);
	}

	RedrawCustomizeButton ();
}
//**********************************************************************************
void CBCGPExplorerToolBar::OnFillBackground (CDC* pDC)
{
	ASSERT_VALID(this);
	ASSERT_VALID (pDC);
	
	CBCGPToolBar::OnFillBackground (pDC);

	if (!m_rectNavButtons.IsRectEmpty () && !m_bSimplifiedNavigationButtons)
	{
		OnDrawNavButtonsFrame (pDC);
	}
}
//**********************************************************************************
void CBCGPExplorerToolBar::OnDrawNavButton (CDC* pDC, CBCGPToolbarButton* pButton, BOOL bIsForward, CRect rect, BOOL bHighlight)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	BOOL bIsHome = pButton->m_nID == m_nHomeID && m_nHomeID != 0;

	int nIndex = 1;

	if (pButton->m_nStyle & TBBS_DISABLED)
	{
		nIndex = 0;
	}
	else if (pButton->m_nStyle & TBBS_PRESSED)
	{
		nIndex = 3;
	}
	else if (bHighlight)
	{
		nIndex = 2;
	}

	if (bIsForward)
	{
		nIndex += 4;
	}
	else if (bIsHome)
	{
		nIndex += 8;
	}

	if (!bIsForward && !bIsHome)
	{
		rect.left += nHorzMargin;
	}

	CBCGPToolBarImages& images = 
		(globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ()) && m_NavImages16.GetCount () > 0 ?
		m_NavImages16 : 
		(m_bLargeIcons ? m_NavImagesLarge : m_NavImages);

	images.DrawEx (pDC, rect, nIndex, 
			bIsForward ? CBCGPToolBarImages::ImageAlignHorzRight : CBCGPToolBarImages::ImageAlignHorzCenter,
			CBCGPToolBarImages::ImageAlignVertCenter);
}
//**********************************************************************************
void CBCGPExplorerToolBar::OnDrawNavButtonsFrame (CDC* pDC)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);

	if (globalData.IsHighContastMode())
	{
		return;
	}

	const int cyFrame = GetNavFrameSize ().cy;

	CRect rectFrame = m_rectNavButtons;
	rectFrame.top = rectFrame.CenterPoint ().y - cyFrame / 2;
	rectFrame.bottom = rectFrame.top + cyFrame;

	int nIndex = 0;

	if (!m_rectHistoryButton.IsRectEmpty ())
	{
		rectFrame.right = m_rectHistoryButton.right;
		nIndex = 1;
		
		int nHistoryIndex = m_nHomeID == 0 ? 2 : 3;

		if (m_iHighlighted == nHistoryIndex)
		{
			nIndex = 2;
		}

		if (m_bIsHistoryDroppedDown)
		{
			nIndex = 3;
		}
	}

	CBCGPToolBarImages& images = 
		(globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ()) && m_NavFrames16.GetCount () > 0 ?
		m_NavFrames16 : 
		(m_bLargeIcons ? m_NavFramesLarge : m_NavFrames);

	images.DrawEx (pDC, rectFrame, nIndex);
}
//**********************************************************************************
int CBCGPExplorerToolBar::CalcMaxButtonHeight ()
{
	int nNextIndex = 0;

	if (m_nBackID != 0 && CommandToIndex (m_nBackID) < 0)
	{
		if (m_nHomeID != 0 && CommandToIndex (m_nHomeID) < 0)
		{
			InsertButton (CBCGPExplorerNavigationButton (m_nHomeID, BCGPNav_Home), nNextIndex++);
		}

		InsertButton (CBCGPExplorerNavigationButton (m_nBackID, BCGPNav_Back), nNextIndex++);

		if (m_nForwardID != 0)
		{
			InsertButton (CBCGPExplorerNavigationButton (m_nForwardID, BCGPNav_Forward), nNextIndex++);

			if (m_nHistoryID != 0)
			{
				InsertButton (CBCGPExplorerNavigationHistoryButton (m_nHistoryID, m_nHistoryItemID), nNextIndex++);
			}
		}

		if (m_nUpID != 0)
		{
			InsertButton (CBCGPExplorerNavigationButton (m_nUpID, BCGPNav_Up), nNextIndex++);
		}

		InsertButton (CBCGPExplorerNavigationButton(globalUtils.ScaleByDPI(m_xRightMargin)), nNextIndex++);
	}

	if (m_nAddressBarID != 0 && CommandToIndex(m_nAddressBarID) < 0)
	{
		::ShowWindow(m_hWndAddressBar, SW_SHOWNOACTIVATE);
		InsertButton (CBCGPExplorerAddressButton(m_nAddressBarID, m_hWndAddressBar, m_nAddressBarMinWidth), nNextIndex);
	}

	return CBCGPToolBar::CalcMaxButtonHeight ();
}
//**********************************************************************************
void CBCGPExplorerToolBar::AdjustLayout ()
{
	CBCGPToolBar::AdjustLayout ();

#if (!defined _BCGSUITE_)
	if (GetSafeHwnd () == NULL)
	{
		return;
	}

	BOOL bOnGlass = m_bOnGlass;

	m_bOnGlass = FALSE;

	if (globalData.DwmIsCompositionEnabled () && m_bIsAeroEnabled && !m_bIsHidden)
	{
		CRect rectWindow;
		GetWindowRect (rectWindow);

		BCGPMARGINS margins;
		margins.cxLeftWidth = 0;
		margins.cxRightWidth = 0;
		margins.cyTopHeight = CBCGPVisualManager::GetInstance()->IsDWMCaptionSupported () && !globalData.IsTitleBarAccentColor() ? rectWindow.Height () : 0;
		margins.cyBottomHeight = 0;

		if (globalData.DwmExtendFrameIntoClientArea(GetParent ()->GetSafeHwnd (), &margins))
		{
			m_bOnGlass = margins.cyTopHeight > 0;
		}
	}

	if (bOnGlass != m_bOnGlass)
	{
		SendMessage (BCGM_ONSETCONTROLAERO, (WPARAM) m_bOnGlass);
	}
#endif
}
//**********************************************************************************
void CBCGPExplorerToolBar::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CBCGPToolBar::OnShowWindow(bShow, nStatus);

	m_bIsHidden = !bShow;
	
	if (!bShow && m_bOnGlass)
	{
		BCGPMARGINS margins;
		margins.cxLeftWidth = 0;
		margins.cxRightWidth = 0;
		margins.cyTopHeight = 0;
		margins.cyBottomHeight = 0;

		BCGCBProDwmExtendFrameIntoClientArea(GetParent()->GetSafeHwnd(), &margins);
	}
}
//**********************************************************************************
void CBCGPExplorerToolBar::DWMCompositionChanged ()
{
	if (m_bIsAeroEnabled)
	{
		AdjustLayout ();
		RedrawWindow ();
	}
}
//**********************************************************************************
int CBCGPExplorerToolBar::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CBCGPToolBar::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	OnLoadNavImages ();
	return 0;
}
//**********************************************************************************
void CBCGPExplorerToolBar::SetStretchID (UINT nID)
{
	m_nStretchID = nID;
}
//**********************************************************************************
CSize CBCGPExplorerToolBar::CalcLayout (DWORD dwMode, int nLength)
{
	CSize size = CBCGPToolBar::CalcLayout (dwMode, nLength);

	if (m_nBackID != 0 && !m_bSimplifiedNavigationButtons)
	{
		size.cy = max (size.cy, GetNavFrameSize ().cy + 2);
	}

	return size;
}
//**********************************************************************************
void CBCGPExplorerToolBar::OnLButtonDown(UINT nFlags, CPoint point) 
{
	if (m_bIsHistoryDroppedDown)
	{
		return;
	}

#if (!defined _BCGSUITE_) && (!defined BCGP_EXCLUDE_RIBBON)
	if ((m_bOnGlass || CBCGPVisualManager::GetInstance()->DoesRibbonExtendCaptionToTabsArea()) && HitTest(point) < 0)
	{
		CPoint ptScreen = point;
		ClientToScreen(&ptScreen);

		GetParent ()->SendMessage (WM_NCLBUTTONDOWN, 
			(WPARAM) HTCAPTION, MAKELPARAM (ptScreen.x, ptScreen.y));

		return;
	}
#endif
	CBCGPToolBar::OnLButtonDown(nFlags, point);
}
//**********************************************************************************
void CBCGPExplorerToolBar::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
#if (!defined _BCGSUITE_) && (!defined BCGP_EXCLUDE_RIBBON)
	if ((m_bOnGlass || CBCGPVisualManager::GetInstance()->DoesRibbonExtendCaptionToTabsArea()) && HitTest(point) < 0)
	{
		CPoint ptScreen = point;
		ClientToScreen(&ptScreen);

		GetParent ()->SendMessage (WM_NCLBUTTONDBLCLK, 
			(WPARAM) HTCAPTION, MAKELPARAM (ptScreen.x, ptScreen.y));

		return;
	}
#endif	
	CBCGPToolBar::OnLButtonDblClk(nFlags, point);
}
//**********************************************************************************
LRESULT CBCGPExplorerToolBar::OnChangeVisualManager (WPARAM, LPARAM)
{
	OnLoadNavImages();
	return 0;
}
//**********************************************************************************
void CBCGPExplorerToolBar::OnSettingChange(UINT uFlags, LPCTSTR lpszSection) 
{
	CBCGPToolBar::OnSettingChange(uFlags, lpszSection);
	AdjustLayout();
}
