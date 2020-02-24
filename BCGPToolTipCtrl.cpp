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
// BCGPToolTipCtrl.cpp : implementation file
//

#include "stdafx.h"

#include "BCGPToolTipCtrl.h"
#include "BCGPToolBarImages.h"
#include "BCGPToolBar.h"
#include "BCGPToolBarButton.h"
#include "BCGPDrawManager.h"
#include "BCGPVisualManager.h"
#include "BCGPOutlookBarPane.h"
#include "BCGPRibbonButton.h"
#include "BCGPRibbonBar.h"
#include "BCGPGlobalUtils.h"
#include "BCGPMDIFrameWnd.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#ifndef TTS_BALLOON
#define TTS_BALLOON 0x40
#endif

#ifndef TTM_SETTITLE
#define TTM_SETTITLE (WM_USER + 32)
#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPToolTipCtrl

UINT BCGM_GET_TOOLTIP_IMAGE_SIZE	= ::RegisterWindowMessage (_T("BCGM_GET_TOOLTIP_IMAGE_SIZE"));
UINT BCGM_ON_DRAW_TOOLTIP_IMAGE		= ::RegisterWindowMessage (_T("BCGM_ON_DRAW_TOOLTIP_IMAGE"));

IMPLEMENT_DYNCREATE(CBCGPToolTipCtrl, CToolTipCtrl)

CBCGPToolTipCtrl::CBCGPToolTipCtrl(CBCGPToolTipParams* pParams/* = NULL*/)
{
	m_bResetDelayTime = FALSE;
	m_bIsOwnerDrawImage = FALSE;

	SetParams (pParams);

	m_pToolBar = NULL;
	m_pToolBarImages = NULL;
	m_pHotButton = NULL;
	m_sizeImage = CSize (0, 0);
	m_sizeCornerRadius = CSize (0, 0);
	m_ptMargin = CPoint (0, 0);
	m_ptLocation = CPoint (-1, -1);
	m_pRibbonButton = NULL;
	m_nRibbonImageType = 0;
	m_nFixedWidthRegular = 0;
	m_nFixedWidthWithImage = 0;
}
//**********************************************************************
CBCGPToolTipCtrl::~CBCGPToolTipCtrl()
{
}

BEGIN_MESSAGE_MAP(CBCGPToolTipCtrl, CToolTipCtrl)
	//{{AFX_MSG_MAP(CBCGPToolTipCtrl)
	ON_WM_PAINT()
	ON_WM_CREATE()
	ON_WM_ERASEBKGND()
	//}}AFX_MSG_MAP
	ON_NOTIFY_REFLECT(TTN_SHOW, OnShow)
	ON_NOTIFY_REFLECT(TTN_POP, OnPop)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPToolTipCtrl message handlers

void CBCGPToolTipCtrl::OnPaint() 
{
	if (m_Params.m_bBallonTooltip)
	{
		CToolTipCtrl::OnPaint ();
		return;
	}

	CPaintDC dcPaint(this); // device context for painting
	
	CBCGPMemDC memDC (dcPaint, this);
	CDC* pDC = &memDC.GetDC ();

	CRect rect;
	GetClientRect (rect);

	CRect rectMargin;
	GetMargin (rectMargin);

	CRect rectText = rect;
	CRect rectHelp = rectText;

	rectText.DeflateRect (rectMargin);
	rectText.DeflateRect (m_ptMargin.x, m_ptMargin.y);

	COLORREF clrLine = m_Params.m_clrBorder == (COLORREF)-1 ?
		::GetSysColor (COLOR_INFOTEXT) : m_Params.m_clrBorder;

	COLORREF clrText = m_Params.m_clrText == (COLORREF)-1 ?
		::GetSysColor (COLOR_INFOTEXT) : m_Params.m_clrText;

	//-----------------
	// Fill background:
	//-----------------
	OnFillBackground (pDC, rect, clrText, clrLine);

	CPen penLine (PS_SOLID, 1, clrLine);
	CPen* pOldPen = pDC->SelectObject (&penLine);

	//-------------
	// Draw border:
	//-------------
	OnDrawBorder (pDC, rect, clrLine);

	//-----------
	// Draw icon:
	//-----------
	int nImageBottom = 0;

	if (m_sizeImage != CSize (0, 0) && m_Params.m_bDrawIcon)
	{
		CRect rectImage = rectText;
		rectImage.right = rectImage.left + m_sizeImage.cx;
		rectImage.bottom = rectImage.top + m_sizeImage.cy;

		OnDrawIcon (pDC, rectImage);

		rectText.left += m_sizeImage.cx + m_ptMargin.x;
		nImageBottom = rectImage.bottom + m_ptMargin.y;
	}

	pDC->SetBkMode (TRANSPARENT);
	pDC->SetTextColor (clrText);

	//------------
	// Draw label:
	//------------
	int nTextHeight = OnDrawLabel (pDC, rectText, FALSE).cy;

	rectHelp.top = rectText.top + nTextHeight + 3 * m_ptMargin.y / 2;

	//------------------------------
	// Draw separator + description:
	//------------------------------
	if (!m_strDescription.IsEmpty () && m_Params.m_bDrawDescription)
	{
		CRect rectDescr = rectText;
		rectDescr.top = rectHelp.top;

		if (m_Params.m_bDrawSeparator)
		{
			OnDrawSeparator (pDC, rectDescr.left,
				rectDescr.right, rectDescr.top - m_ptMargin.y / 2);
		}

		CSize sizeDescr = OnDrawDescription (pDC, rectDescr, FALSE);

		rectHelp.top += sizeDescr.cy + m_ptMargin.y;
	}

	//--------------------------
	// Draw context help prompt:
	//--------------------------
	if (!m_strHelpPrompt.IsEmpty())
	{
		rectHelp.top = max(rectHelp.top, nImageBottom);

		int nPadding = globalUtils.ScaleByDPI(5);

		CBCGPDrawManager dm(*pDC);
		dm.DrawLine(rectHelp.left + nPadding, rectHelp.top, rectHelp.right - nPadding, rectHelp.top, 
			m_Params.m_clrBorder == (COLORREF)-1 ?
			CBCGPVisualManager::GetInstance()->GetSeparatorColor() : m_Params.m_clrBorder);

		rectHelp.DeflateRect(m_ptMargin.x, 0);
		OnDrawHelpPrompt(pDC, rectHelp, FALSE);
	}
	
	pDC->SelectObject (pOldPen);
}
//**********************************************************************
void CBCGPToolTipCtrl::OnShow(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
	*pResult = 0;

	m_bIsOwnerDrawImage = FALSE;

	if (m_Params.m_bVislManagerTheme)
	{
		CBCGPVisualManager::GetInstance ()->GetToolTipParams (m_Params);
		m_Params.m_bVislManagerTheme = TRUE;
	}

	if (m_Params.m_bBallonTooltip)
	{
		return;
	}

	CPoint ptCursor;
	::GetCursorPos (&ptCursor);

	GetHotButton ();

	m_sizeImage = m_Params.m_bDrawIcon ? GetIconSize () : CSize (0, 0);
	m_ptMargin = globalUtils.ScaleByDPI(CPoint (6, 4));

	m_ptMargin.x += m_Params.m_nPaddingX;
	m_ptMargin.y += m_Params.m_nPaddingY;

	CRect rectMargin;
	GetMargin (rectMargin);

	CRect rectText;
	GetClientRect (rectText);

	CClientDC dc (this);
	CSize sizeText = OnDrawLabel (&dc, rectText, TRUE);
	
	int cx = sizeText.cx;
	int cy = sizeText.cy;

	CSize sizeDescr (0, 0);
	
	if (!m_Params.m_bDrawDescription || m_strDescription.IsEmpty ())
	{
		cy = max (cy, m_sizeImage.cy);
	}
	else
	{
		sizeDescr = OnDrawDescription (&dc, rectText, TRUE);

		cy += sizeDescr.cy + 2 * m_ptMargin.y;
		cx = max (cx, sizeDescr.cx);

		cy = max (cy, m_sizeImage.cy);
	}

	CSize sizeHelp = OnDrawHelpPrompt(&dc, rectText, TRUE);
	if (sizeHelp != CSize(0, 0))
	{
		cy += sizeHelp.cy + m_ptMargin.y;
		cx = max (cx, sizeHelp.cx);
	}

	if (m_sizeImage.cx > 0 && m_Params.m_bDrawIcon)
	{
		cx += m_sizeImage.cx + m_ptMargin.x;
	}

	cx += 2 * m_ptMargin.x;
	cy += 2 * m_ptMargin.y;

	int nFixedWidth = GetFixedWidth ();
	if (nFixedWidth > 0 && sizeDescr != CSize (0, 0))
	{
		cx = max (cx, nFixedWidth);
	}

	CRect rectWindow;
	GetWindowRect (rectWindow);

	CRgn rgn;
	rgn.CreateRectRgn(rectWindow.left, rectWindow.top, rectWindow.right + 1, rectWindow.bottom + 1);
	
	SetWindowRgn(rgn, FALSE);

	int x = rectWindow.left;
	int y = rectWindow.top;

	if (m_ptLocation != CPoint (-1, -1))
	{
		x = m_ptLocation.x;
		y = m_ptLocation.y;

		*pResult = 1;
	}

	CRect rectScreen;

	MONITORINFO mi;
	mi.cbSize = sizeof (MONITORINFO);
	if (GetMonitorInfo (MonitorFromPoint (rectWindow.TopLeft (), MONITOR_DEFAULTTONEAREST),
		&mi))
	{
		rectScreen = mi.rcWork;
	}
	else
	{
		::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectScreen, 0);
	}

	int nBottom = max (
		ptCursor.y + cy + ::GetSystemMetrics (SM_CYCURSOR),
		y + cy + 2);

	if (nBottom > rectScreen.bottom)
	{
		y = ptCursor.y - cy - 1;

#ifndef BCGP_EXCLUDE_RIBBON
		if (m_pRibbonButton != NULL && m_ptLocation != CPoint (-1, -1))
		{
			ASSERT_VALID (m_pRibbonButton);

			CBCGPRibbonBar* pRibbon = m_pRibbonButton->GetTopLevelRibbonBar ();
			if (pRibbon->GetSafeHwnd () != NULL)
			{
				CRect rectRibbon;
				pRibbon->GetWindowRect (rectRibbon);

				y = rectRibbon.top - cy;
			}
		}
#endif

		*pResult = 1;
	}

	if (x + cx + 2 > rectScreen.right)
	{
		if ((*pResult) == 1)	// Y has been changed
		{
			x = ptCursor.x - cx - 1;
		}
		else
		{
			x = rectScreen.right - cx - 1;
			*pResult = 1;
		}
	}

	if ((*pResult) == 1)
	{
		SetWindowPos (NULL, x, y, cx, cy, SWP_NOZORDER | SWP_NOACTIVATE);
	}
	else
	{
		SetWindowPos (NULL, -1, -1, cx, cy, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
	}

	m_sizeCornerRadius = CBCGPVisualManager::GetInstance()->GetSystemToolTipCornerRadius(this);

	CWnd* pWndTop = CBCGPMDIFrameWnd::GetActiveTearOffFrame();
	SetWindowPos(pWndTop == NULL ? &wndTop : pWndTop, -1, -1, -1, -1, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOSIZE | SWP_DRAWFRAME);

    SetWindowRgn(NULL, FALSE);
}
//**********************************************************************
void CBCGPToolTipCtrl::OnPop(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
#ifndef BCGP_EXCLUDE_RIBBON
	if (m_pRibbonButton != NULL)
	{
		ASSERT_VALID(m_pRibbonButton);
		
		CBCGPRibbonBar* pRibbonBar = m_pRibbonButton->GetTopLevelRibbonBar();
		if (pRibbonBar != NULL)
		{
			pRibbonBar->SetContextHelpActiveID(NULL);
		}
	}
#endif

	m_strHelpPrompt.Empty();

	m_pHotButton = NULL;
	m_pToolBarImages = NULL;
	m_strDescription.Empty ();
	m_ptLocation = CPoint (-1, -1);
	m_pRibbonButton = NULL;
	m_nRibbonImageType = 0;
	*pResult = 0;
}
//**********************************************************************
void CBCGPToolTipCtrl::OnFillBackground (CDC* pDC, CRect rect, COLORREF& clrText, COLORREF& clrLine)
{
	ASSERT_VALID (pDC);

	if (m_Params.m_clrFill == (COLORREF)-1)
	{
		CBCGPVisualManager::GetInstance ()->OnFillToolTip (pDC, this, rect,
			clrText, clrLine);
	}
	else
	{
		if (m_Params.m_clrFillGradient == (COLORREF)-1)
		{
			CBrush br (m_Params.m_clrFill);
			pDC->FillRect (rect, &br);
		}
		else
		{
			CBCGPDrawManager dm (*pDC);

			dm.FillGradient2 (rect, 
				m_Params.m_clrFillGradient, m_Params.m_clrFill,
				m_Params.m_nGradientAngle == -1 ? 90 : m_Params.m_nGradientAngle);
		}
	}
}
//**********************************************************************
CSize CBCGPToolTipCtrl::GetIconSize()
{
	CWnd* pWndParent = GetParent();
	if (pWndParent->GetSafeHwnd() != NULL)
	{
		CSize size(0, 0);
		if (pWndParent->SendMessage(BCGM_GET_TOOLTIP_IMAGE_SIZE, 0, (LPARAM)&size))
		{
			m_bIsOwnerDrawImage = TRUE;
			return size;
		}
	}

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_pRibbonButton != NULL)
	{
		ASSERT_VALID (m_pRibbonButton);
		return m_pRibbonButton->GetToolTipImageSize(m_nRibbonImageType);
	}
#endif

	if (m_pHotButton == NULL || m_pToolBarImages == NULL ||
		m_pToolBarImages->GetCount () == 0)
	{
		return CSize (0, 0);
	}

	ASSERT_VALID (m_pHotButton);

	CSize size = globalUtils.ScaleByDPI(m_pHotButton->GetImage () >= 0 ? m_pToolBarImages->GetImageSize () : CSize (0, 0));
	return size;
}
//**********************************************************************
void CBCGPToolTipCtrl::OnDrawBorder (CDC* pDC, CRect rect,
									 COLORREF clrLine)
{
	ASSERT_VALID (pDC);

	if (globalData.bIsWindows8)
	{
		pDC->Draw3dRect (rect, clrLine, clrLine);
		return;
	}

	const int nOffsetX = m_sizeCornerRadius.cx / 2;
	const int nOffsetY = m_sizeCornerRadius.cy / 2;

	if (nOffsetX == 0 || nOffsetY == 0)
	{
		pDC->Draw3dRect (rect, clrLine, clrLine);
		return;
	}

	pDC->MoveTo (rect.left + nOffsetX, rect.top);
	pDC->LineTo (rect.right - nOffsetX - 1, rect.top);

	pDC->LineTo (rect.right - 1, rect.top + nOffsetY);
	pDC->LineTo (rect.right - 1, rect.bottom - 1 - nOffsetY);

	pDC->LineTo (rect.right - nOffsetX - 1, rect.bottom - 1);
	pDC->LineTo (rect.left + nOffsetX, rect.bottom - 1);

	pDC->LineTo (rect.left, rect.bottom - 1 - nOffsetY);
	pDC->LineTo (rect.left, rect.top + nOffsetY);

	pDC->LineTo (rect.left + nOffsetX, rect.top);
}
//**********************************************************************
BOOL CBCGPToolTipCtrl::OnDrawIcon (CDC* pDC, CRect rectImage)
{
	ASSERT_VALID (pDC);

	if (m_bIsOwnerDrawImage)
	{
		CWnd* pWndParent = GetParent();
		if (pWndParent->GetSafeHwnd() != NULL)
		{
			return (BOOL)pWndParent->SendMessage(BCGM_ON_DRAW_TOOLTIP_IMAGE, (WPARAM)pDC->GetSafeHdc(), (LPARAM)&rectImage);
		}
	}

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_pRibbonButton != NULL)
	{
		ASSERT_VALID (m_pRibbonButton);
		m_pRibbonButton->OnDrawTooltipImage(pDC, (CBCGPBaseRibbonElement::RibbonImageType) m_nRibbonImageType, rectImage);

		return TRUE;
	}
#endif

	if (m_pHotButton == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID (m_pHotButton);

	CBCGPToolBarImages* pImages = m_pToolBarImages;
	
	CBCGPToolBar* pToolBar = DYNAMIC_DOWNCAST(CBCGPToolBar, m_pHotButton->GetParentWnd());
	if (pToolBar != NULL)
	{
		ASSERT_VALID(pToolBar);
		pToolBar->CheckForButtonImages(m_pHotButton, &pImages);
	}

	if (pImages == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID(pImages);

	CBCGPDrawState ds(CBCGPVisualManager::GetInstance()->IsAutoGrayscaleImages());
	pImages->PrepareDrawImage (ds);

	UINT nSaveStyle = m_pHotButton->m_nStyle;
	BOOL bSaveText = m_pHotButton->m_bText;
	BOOL bSaveImage = m_pHotButton->m_bImage;
	BOOL bSaveTextBelow = m_pHotButton->m_bTextBelow;

	BOOL bSaveLargeIcons = CBCGPToolBar::m_bLargeIcons;
	CBCGPToolBar::m_bLargeIcons = FALSE;

	m_pHotButton->m_bText = FALSE;
	m_pHotButton->m_bImage = TRUE;
	m_pHotButton->m_bTextBelow = FALSE;

	m_pHotButton->m_nStyle = 0;

	m_pHotButton->CBCGPToolbarButton::OnDraw (pDC, rectImage, pImages);

	m_pHotButton->m_nStyle = nSaveStyle;
	m_pHotButton->m_bText = bSaveText;
	m_pHotButton->m_bImage = bSaveImage;
	m_pHotButton->m_bTextBelow = bSaveTextBelow;

	CBCGPToolBar::m_bLargeIcons = bSaveLargeIcons;

	pImages->EndDrawImage (ds);
	return TRUE;
}
//**********************************************************************
CString CBCGPToolTipCtrl::GetLabel ()
{
	CString strText;
	GetWindowText (strText);

	return strText;
}
//**********************************************************************
CSize CBCGPToolTipCtrl::OnDrawLabel (CDC* pDC, CRect rect, BOOL bCalcOnly)
{
	ASSERT_VALID (pDC);

	CSize sizeText (0, 0);

	CString strText (GetLabel ());

	strText.Replace (_T("\t"), _T("    "));

	if (globalUtils.m_bRemoveDoubleAmpersands)
	{
		strText.Replace(_T("&&"), _T("&"));
	}

	BOOL bDrawDescr = m_Params.m_bDrawDescription && !m_strDescription.IsEmpty ();

	CFont* pOldFont = (CFont*) pDC->SelectObject (
		m_Params.m_bBoldLabel && bDrawDescr ?
			&globalData.fontBold : &globalData.fontTooltip);

	if (strText.Find (_T('\n')) >= 0)	// Multi-line text
	{
		UINT nFormat = DT_NOPREFIX;
		if (bCalcOnly)
		{
			nFormat |= DT_CALCRECT;
		}

		int nHeight = pDC->DrawText (strText, rect, nFormat);
		sizeText = CSize (rect.Width (), nHeight);
	}
	else
	{
		if (bCalcOnly)
		{
			sizeText = pDC->GetTextExtent (strText);
		}
		else
		{
			UINT nFormat = DT_LEFT | DT_NOCLIP | DT_SINGLELINE;
		
			if (!bDrawDescr && m_strHelpPrompt.IsEmpty())
			{
				nFormat |= DT_VCENTER;
			}

			if (m_pRibbonButton != NULL)
			{
				nFormat |= DT_NOPREFIX;
			}

			sizeText.cy = pDC->DrawText (strText, rect, nFormat);
			sizeText.cx = rect.Width ();
		}
	}

	pDC->SelectObject (pOldFont);

	return sizeText;
}
//**********************************************************************
CSize CBCGPToolTipCtrl::OnDrawDescription (CDC* pDC, CRect rect, BOOL bCalcOnly)
{
	ASSERT_VALID (pDC);

	CSize sizeText (0, 0);

	if (!m_Params.m_bDrawDescription)
	{
		return sizeText;
	}

	CRect rectSaved = rect;

	CFont* pOldFont = pDC->SelectObject (&globalData.fontTooltip);
	int nFixedWidth = GetFixedWidth ();

	if (nFixedWidth > 0 && m_sizeImage.cx <= 32)
	{
		rect.right = rect.left + nFixedWidth;

		if (m_sizeImage.cx > 0 && m_Params.m_bDrawIcon)
		{
			rect.right -= m_sizeImage.cx + m_ptMargin.x;
		}
	}
	else
	{
		rect.right = rect.left + m_Params.m_nMaxDescrWidth;
	}

	if (rect.Width() < rectSaved.Width())
	{
		rect = rectSaved;
	}

	UINT nFormat = DT_WORDBREAK | DT_NOCLIP | DT_NOPREFIX;
	if (bCalcOnly)
	{
		nFormat |= DT_CALCRECT;
	}

	CString strDescr(m_strDescription);

	if (globalUtils.m_bRemoveDoubleAmpersands)
	{
		strDescr.Replace(_T("&&"), _T("&"));
	}

	int nDescrHeight = pDC->DrawText (strDescr, rect, nFormat);
	pDC->SelectObject (pOldFont);

	return CSize (rect.Width (), nDescrHeight);
}
//**********************************************************************
CSize CBCGPToolTipCtrl::OnDrawHelpPrompt(CDC* pDC, CRect rect, BOOL bCalcOnly)
{
	ASSERT_VALID (pDC);

	if (m_strHelpPrompt.IsEmpty())
	{
		return CSize(0, 0);
	}

	int nIconWidth = 0;

	if (m_Params.m_hiconHelp != NULL)
	{
		nIconWidth = globalData.m_sizeSmallIcon.cx + 2 * m_ptMargin.x;

		if (!bCalcOnly)
		{
			CPoint ptIcon(
					rect.left + m_ptMargin.x,
					rect.top + max(0, (int)(0.5 + 0.5 * (rect.Height() - globalData.m_sizeSmallIcon.cy))));
			
			::DrawIconEx(pDC->GetSafeHdc (), 
				ptIcon.x,
				ptIcon.y,
				m_Params.m_hiconHelp, globalData.m_sizeSmallIcon.cx, globalData.m_sizeSmallIcon.cy, 0, NULL,
				DI_NORMAL);
			
			rect.left += nIconWidth;
		}
	}

	CFont* pOldFont = pDC->SelectObject (&globalData.fontTooltip);

	UINT nFormat = DT_SINGLELINE | DT_LEFT | DT_VCENTER;
	if (bCalcOnly)
	{
		nFormat |= DT_CALCRECT;
	}

	int nHelpHeight = pDC->DrawText(m_strHelpPrompt, rect, nFormat);
	pDC->SelectObject (pOldFont);

	return CSize(nIconWidth + rect.Width(), nHelpHeight);
}
//**********************************************************************
void CBCGPToolTipCtrl::OnDrawSeparator (CDC* pDC, int x1, int x2, int y)
{
	ASSERT_VALID (pDC);

	pDC->MoveTo (x1, y);
	pDC->LineTo (x2, y);
}
//**********************************************************************
int CBCGPToolTipCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CToolTipCtrl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	m_pToolBar = DYNAMIC_DOWNCAST (CBCGPToolBar,
		CWnd::FromHandlePermanent (lpCreateStruct->hwndParent));

	if (m_pToolBar != NULL && m_pToolBar->IsKindOf (RUNTIME_CLASS (CBCGPOutlookBarPane)))
	{
		m_pToolBar = NULL;
	}

	ModifyStyle (WS_BORDER, 0);

	if (m_Params.m_bBallonTooltip)
	{
		ModifyStyle (0, TTS_BALLOON);
	}

	AdjustDelayTime();
	return 0;
}
//**********************************************************************
BOOL CBCGPToolTipCtrl::OnEraseBkgnd(CDC* pDC) 
{
	CRect rect;
	GetClientRect (rect);

	COLORREF clrDummy;

	OnFillBackground (pDC, rect, clrDummy, clrDummy);
	return TRUE;
}
//**********************************************************************
void CBCGPToolTipCtrl::SetParams (CBCGPToolTipParams* pParams)
{
	ASSERT_VALID (this);

	BOOL bBallonTooltipSaved = m_Params.m_bBallonTooltip;
	BOOL bDelayWasSet = m_Params.m_nDelayAutomatic != 0 || m_Params.m_nDelayAutoPop != 0 || m_Params.m_nDelayInitial != 0 || m_Params.m_nDelayReshow != 0;

	if (pParams == NULL)
	{
		if (bDelayWasSet)
		{
			m_bResetDelayTime = TRUE;
		}

		CBCGPToolTipParams paramsDefault;
		m_Params = paramsDefault;
	}
	else
	{
		m_Params = *pParams;

		if (bDelayWasSet && m_Params.m_nDelayAutomatic == 0 && m_Params.m_nDelayAutoPop == 0 && m_Params.m_nDelayInitial == 0 && m_Params.m_nDelayReshow == 0)
		{
			m_bResetDelayTime = TRUE;
		}
	}

	if (bBallonTooltipSaved != m_Params.m_bBallonTooltip && GetSafeHwnd() != NULL)
	{
		if (m_Params.m_bBallonTooltip)
		{
			ModifyStyle (0, TTS_BALLOON);
		}
		else
		{
			ModifyStyle (TTS_BALLOON, 0);
		}
	}

	AdjustDelayTime();
}
//**********************************************************************
void CBCGPToolTipCtrl::SetDescription (const CString strDescription)
{
	ASSERT_VALID (this);

	GetHotButton ();
	m_strDescription = strDescription;

	m_strDescription.Replace (_T("\t"), _T("    "));
}
//**********************************************************************
void CBCGPToolTipCtrl::GetHotButton ()
{
	m_pHotButton = NULL;
	m_pToolBarImages = NULL;

	if (m_pRibbonButton != NULL)
	{
		return;
	}

	if (m_pToolBar != NULL)
	{
		CPoint ptToolBar;

		::GetCursorPos (&ptToolBar);
		m_pToolBar->ScreenToClient (&ptToolBar);

		m_pHotButton = m_pToolBar->GetButton (m_pToolBar->HitTest (ptToolBar));

		if (m_pHotButton != NULL)
		{
			if (m_pToolBar->IsLocked ())
			{
				m_pToolBarImages = m_pToolBar->GetLockedMenuImages ();
			}
			else
			{
				if (m_pHotButton->m_bUserButton)
				{
					m_pToolBarImages = CBCGPToolBar::GetUserImages ();
				}
				else
				{
					m_pToolBarImages = CBCGPToolBar::GetMenuImages ();
					if (m_pToolBarImages == NULL || m_pToolBarImages->GetCount () <= 0)
					{
						m_pToolBarImages = CBCGPToolBar::GetImages ();
					}
				}
			}
		}
	}
}
//**********************************************************************
void CBCGPToolTipCtrl::SetLocation (CPoint pt)
{
	ASSERT_VALID (this);

	m_ptLocation = pt;
}
//**********************************************************************
void CBCGPToolTipCtrl::SetHotRibbonButton (CBCGPRibbonButton* pRibbonButton, LPCTSTR lpszHelpPrompt)
{
	ASSERT_VALID (this);

	m_pRibbonButton = pRibbonButton;

	m_strHelpPrompt.Empty();

#ifndef BCGP_EXCLUDE_RIBBON
	if (lpszHelpPrompt != NULL && pRibbonButton != NULL && 
		pRibbonButton->GetID() != 0 && pRibbonButton->GetID() != (UINT)-1)
	{
		m_strHelpPrompt = lpszHelpPrompt;
	}
#endif
}
//**********************************************************************
int CBCGPToolTipCtrl::GetFixedWidth ()
{
	ASSERT_VALID (this);

	if (m_sizeImage.cx <= (int) (globalData.GetRibbonImageScale () * 32))
	{
		return m_nFixedWidthRegular;
	}
	else
	{
		return m_nFixedWidthWithImage;
	}
}
//**********************************************************************
void CBCGPToolTipCtrl::AdjustDelayTime()
{
	if (GetSafeHwnd() == NULL)
	{
		return;
	}
	
	if (m_bResetDelayTime)
	{
		SetDelayTime(TTDT_AUTOMATIC, -1);
		m_bResetDelayTime = FALSE;
		return;
	}
	
	if (m_Params.m_nDelayAutomatic != 0)
	{
		SetDelayTime(TTDT_AUTOMATIC, m_Params.m_nDelayAutomatic);
		return;
	}
	
	if (m_Params.m_nDelayAutoPop != 0)
	{
		SetDelayTime(TTDT_AUTOPOP, m_Params.m_nDelayAutoPop);
	}
	
	if (m_Params.m_nDelayInitial != 0)
	{
		SetDelayTime(TTDT_INITIAL, m_Params.m_nDelayInitial);
	}
	
	if (m_Params.m_nDelayReshow != 0)
	{
		SetDelayTime(TTDT_RESHOW, m_Params.m_nDelayReshow);
	}
}
