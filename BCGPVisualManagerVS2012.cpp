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
// BCGPVisualManagerVS2012.cpp: implementation of the CBCGPVisualManagerVS2012 class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BCGPGlobalUtils.h"
#include "BCGPVisualManagerVS2012.h"
#include "BCGPToolbarMenuButton.h"
#include "BCGPDrawManager.h"
#include "BCGPTabWnd.h"
#include "BCGPAutoHideButton.h"
#include "BCGPColorBar.h"
#include "BCGPCalendarBar.h"
#include "bcgpstyle.h"
#include "BCGPAutoHideDockBar.h"
#include "BCGPSlider.h"
#include "BCGPStatusBar.h"
#include "BCGPToolbarEditBoxButton.h"
#include "BCGPToolBox.h"
#include "BCGPPropList.h"
#include "BCGPOutlookButton.h"
#include "CustomizeButton.h"
#include "BCGPCustomizeMenuButton.h"
#include "BCGPToolbarComboBoxButton.h"
#include "BCGPGroup.h"
#include "BCGPReBar.h"
#include "BCGPGridCtrl.h"
#include "BCGPMiniFrameWnd.h"
#include "BCGPCaptionBar.h"
#include "BCGPURLLinkButton.h"
#include "BCGPPopupWindow.h"
#include "BCGPBreadcrumb.h"
#include "BCGPFrameWnd.h"
#include "BCGPMDIFrameWnd.h"
#include "BCGPMDIChildWnd.h"
#include "BCGPPropertySheet.h"
#include "BCGPGanttItem.h"
#include "BCGPGanttChart.h"
#include "BCGPCalculator.h"
#include "BCGPDockBarRow.h"
#include "BCGPProgressCtrl.h"
#include "BCGPComboBox.h"
#include "BCGPSpinButtonCtrl.h"
#include "BCGPHeaderCtrl.h"
#include "BCGPSliderCtrl.h"
#include "BCGPDialog.h"
#include "BCGPTasksPane.h"
#include "BCGPChartFormat.h"
#include "BCGPCircularGaugeImpl.h"
#include "BCGPLinearGaugeImpl.h"
#include "BCGPSplitterWnd.h"
#include "BCGPSwitchImpl.h"

#ifndef BCGP_EXCLUDE_PLANNER
#include "BCGPPlannerViewDay.h"
#include "BCGPPlannerViewWeek.h"
#include "BCGPPlannerViewMonth.h"
#include "BCGPPlannerViewMulti.h"
#include "BCGPPlannerViewSchedule.h"
#endif

#ifndef BCGP_EXCLUDE_RIBBON
#include "BCGPRibbonCategory.h"
#include "BCGPRibbonMainPanel.h"
#include "BCGPRibbonPanel.h"
#include "BCGPRibbonStatusBar.h"
#include "BCGPRibbonStatusBarPane.h"
#include "BCGPRibbonHyperlink.h"
#include "BCGPRibbonEdit.h"
#include "BCGPRibbonProgressBar.h"
#include "BCGPRibbonSlider.h"
#include "BCGPRibbonBackstageView.h"
#endif

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CBCGPVisualManagerVS2012, CBCGPVisualManagerVS2010)
IMPLEMENT_DYNCREATE(CBCGPVisualManagerVS2013, CBCGPVisualManagerVS2012)
IMPLEMENT_DYNCREATE(CBCGPVisualManagerVS11, CBCGPVisualManagerVS2012)	// Backward compatibility

BOOL CBCGPVisualManagerVS2012::m_bAutoGrayscaleImages = FALSE;
BOOL CBCGPVisualManagerVS2012::m_bThinSlider = FALSE;
COLORREF CBCGPVisualManagerVS2012::m_clrAccent = RGB(0, 122, 204);
COLORREF CBCGPVisualManagerVS2012::m_clrFrame = (COLORREF)-1;
COLORREF CBCGPVisualManagerVS2012::m_clrFrameTextActive = (COLORREF)-1;
COLORREF CBCGPVisualManagerVS2012::m_clrFrameTextInactive = (COLORREF)-1;
CBCGPVisualManagerVS2012::AccentColor CBCGPVisualManagerVS2012::m_curAccentColor = CBCGPVisualManagerVS2012::VS2012_Blue;
CBCGPVisualManagerVS2012::AccentColor CBCGPVisualManagerVS2012::m_curStateColor = (CBCGPVisualManagerVS2012::AccentColor)INT_MAX;
COLORREF CBCGPVisualManagerVS2012::m_clrState = (COLORREF)-1;
CBCGPVisualManagerVS2012::Style CBCGPVisualManagerVS2012::m_Style = CBCGPVisualManagerVS2012::VS2012_Light;

static COLORREF CalculateHourLineColor (COLORREF clrBase, BOOL /*bWorkingHours*/, BOOL bHour)
{
	int nAlpha = bHour ? 85 : 95;
	return CBCGPDrawManager::PixelAlpha (clrBase, nAlpha);
}

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPVisualManagerVS2012::CBCGPVisualManagerVS2012()
{
	OnUpdateSystemColors ();

	m_bConnectMenuToParent = TRUE;
	m_bCheckedRibbonButtonFrame = TRUE;
	m_bTabFillGradient = FALSE;

	if (m_curAccentColor != (AccentColor)INT_MAX)
	{
		m_clrAccent = AccentColorToRGB(m_curAccentColor);
	}

	if (m_curStateColor != (AccentColor)INT_MAX)
	{
		m_clrState = AccentColorToRGB(m_curStateColor);
	}

	m_nVertMargin = globalUtils.ScaleByDPI(1);
	m_nHorzMargin = 0;
	m_nGroupVertOffset = 0;
	m_nGroupCaptionVertOffset = 0;
	m_nTasksMinHeight = globalUtils.ScaleByDPI(20);
	m_nTasksVertMargin = globalUtils.ScaleByDPI(2);
	m_nTasksHorzOffset = globalUtils.ScaleByDPI(3);
	m_nTasksIconHorzOffset = globalUtils.ScaleByDPI(9);
	m_nTasksIconVertOffset = 0;

	m_pCurrentNcWnd = NULL;
}
//**************************************************************************************************************
CBCGPVisualManagerVS2012::~CBCGPVisualManagerVS2012()
{
	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageWhite, (COLORREF)-1);
	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageBlack, (COLORREF)-1);
	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageGray, (COLORREF)-1);
	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageBlack2, (COLORREF)-1);
}
//**************************************************************************************************************
void CBCGPVisualManagerVS2012::SetFrameColor(COLORREF colorFill, COLORREF clrTextActive, COLORREF clrTextInactive, BOOL bRedraw)
{
	if (m_clrFrame == colorFill && m_clrFrameTextActive == clrTextActive && m_clrFrameTextInactive == clrTextInactive)
	{
		return;
	}

	m_clrFrame = colorFill;
	m_clrFrameTextActive = clrTextActive;
	m_clrFrameTextInactive = clrTextInactive;

	CBCGPVisualManagerVS2012* pThis = DYNAMIC_DOWNCAST(CBCGPVisualManagerVS2012, m_pVisManager);
	if (pThis != NULL)
	{
		pThis->OnUpdateSystemColors();

		if (bRedraw)
		{
			pThis->RedrawAll();
		}
	}
}
//**************************************************************************************************************
void CBCGPVisualManagerVS2012::SetAccentColor(AccentColor color)
{
	m_curAccentColor = color;
	m_clrAccent = AccentColorToRGB(color);
	
	CBCGPVisualManagerVS2012* pThis = DYNAMIC_DOWNCAST(CBCGPVisualManagerVS2012, m_pVisManager);
	if (pThis != NULL)
	{
		pThis->OnUpdateSystemColors();
	}
}
//**************************************************************************************************************
void CBCGPVisualManagerVS2012::SetAccentColorRGB(COLORREF color)
{
	m_curAccentColor = (AccentColor)INT_MAX;
	m_clrAccent = color;

	CBCGPVisualManagerVS2012* pThis = DYNAMIC_DOWNCAST(CBCGPVisualManagerVS2012, m_pVisManager);
	if (pThis != NULL)
	{
		pThis->OnUpdateSystemColors();
	}
}
//**************************************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetAccentColorRGB()
{
    return m_clrAccent;
}
//**************************************************************************************************************
void CBCGPVisualManagerVS2012::RedrawAppStatusBar()
{
	if (AfxGetMainWnd()->GetSafeHwnd() != NULL)
	{		
		CWnd* pStatusBar = AfxGetMainWnd()->GetDescendantWindow(AFX_IDW_STATUS_BAR, TRUE);
	
		if (pStatusBar->GetSafeHwnd () != NULL && pStatusBar->IsWindowVisible())
		{
			if (DYNAMIC_DOWNCAST (CBCGPStatusBar, pStatusBar) != NULL
#ifndef BCGP_EXCLUDE_RIBBON
				|| DYNAMIC_DOWNCAST (CBCGPRibbonStatusBar, pStatusBar) != NULL
#endif // BCGP_EXCLUDE_RIBBON
				)
			{
				pStatusBar->RedrawWindow();
				AfxGetMainWnd()->RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE);
			}
		}
	}
}
//**************************************************************************************************************
void CBCGPVisualManagerVS2012::SetStateColor(AccentColor color)
{
	m_curStateColor = color;
	m_clrState = AccentColorToRGB(color);
	
	CBCGPVisualManagerVS2012* pThis = DYNAMIC_DOWNCAST(CBCGPVisualManagerVS2012, m_pVisManager);
	if (pThis != NULL)
	{
		pThis->OnUpdateSystemColors();
		pThis->RedrawAppStatusBar();
	}
}
//**************************************************************************************************************
void CBCGPVisualManagerVS2012::SetStateColorRGB(COLORREF color)
{
	m_curStateColor = (AccentColor)INT_MAX;
	m_clrState = color;

	CBCGPVisualManagerVS2012* pThis = DYNAMIC_DOWNCAST(CBCGPVisualManagerVS2012, m_pVisManager);
	if (pThis != NULL)
	{
		pThis->OnUpdateSystemColors();
		pThis->RedrawAppStatusBar();
	}
}
//**************************************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetStateColorRGB()
{
    return m_clrState;
}
//**************************************************************************************************************
void CBCGPVisualManagerVS2012::SetStyle(Style style)
{
	m_Style = style;

	CBCGPVisualManagerVS2012* pThis = DYNAMIC_DOWNCAST(CBCGPVisualManagerVS2012, m_pVisManager);
	if (pThis != NULL)
	{
		pThis->OnUpdateSystemColors();
	}
}
//**************************************************************************************************************
COLORREF CBCGPVisualManagerVS2012::AccentColorToRGB(AccentColor color)
{
	switch (color)
	{
	case VS2012_Default:
		{
			COLORREF clr = globalData.GetWindowsAccentColor();
			if (clr != (COLORREF)-1)
			{
				return clr;
			}
		}
		// Use blue as default
		
	case VS2012_Blue:
	default:
		return RGB(0, 122, 204);

	case VS2012_Brown:
		return RGB(160,80,0);

	case VS2012_Green:
		return RGB(51,153,51);

	case VS2012_Lime:
		return RGB(137, 164, 48);

	case VS2012_Magenta:
		return RGB(216,0,115);

	case VS2012_Orange:
		return RGB(247, 99, 12);

	case VS2012_Pink:
		return RGB(230,113,184);

	case VS2012_Purple:
		return RGB(162,0,255);

	case VS2012_Red:
		return RGB(229,20,0);

	case VS2012_Teal:
		return RGB(3, 131, 135);
	}
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetNCBorderColor(BOOL bActive)
{
	return bActive ? m_clrAccentText : (m_clrFrame != (COLORREF)-1 ? m_clrFrame : m_clrNcBorder);
}
//*****************************************************************************
CSize CBCGPVisualManagerVS2012::GetSysBorderSize(CWnd* pWnd)
{
	CSize szSysBorder (globalUtils.GetSystemBorders (pWnd));
	
	BOOL bDialog = pWnd->IsKindOf (RUNTIME_CLASS (CBCGPDialog)) || pWnd->IsKindOf (RUNTIME_CLASS (CBCGPPropertySheet));
	BOOL bMDIChild = pWnd->IsKindOf (RUNTIME_CLASS (CBCGPMDIChildWnd));
	
	if (!bDialog && !bMDIChild && !IsDWMCaptionSupported() && IsSmallSystemBorders())
	{
		szSysBorder = CSize(3, 3);
	}

	return szSysBorder;
}
//*****************************************************************************
BOOL CBCGPVisualManagerVS2012::OnNcPaint (CWnd* pWnd, const CObList& lstSysButtons, CRect rectRedraw)
{
	ASSERT_VALID (pWnd);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnNcPaint (pWnd, lstSysButtons, rectRedraw);
	}

	if (pWnd->GetSafeHwnd () == NULL)
	{
		return FALSE;
	}

	CWindowDC dc (pWnd);

	if (dc.GetSafeHdc () != NULL)
	{
		CRgn rgn;
		if (!rectRedraw.IsRectEmpty () && !globalData.bIsRemoteSession)
		{
			rgn.CreateRectRgnIndirect (rectRedraw);
			dc.SelectClipRgn (&rgn);
		}

		BOOL bIsBackstageViewMode = FALSE;

#ifndef BCGP_EXCLUDE_RIBBON
		CBCGPRibbonBar* pBar = GetRibbonBar (pWnd);
		BOOL bRibbonCaption  = pBar != NULL && 
							   pBar->IsWindowVisible () &&
							   pBar->IsReplaceFrameCaption ();

		bIsBackstageViewMode = pBar != NULL && pBar->IsBackstageViewActive();
#else
		BOOL bRibbonCaption = FALSE;
#endif

		COLORREF clrRibbonFrame = (COLORREF)-1;
		int nRibbonTopAreaHeight = -1;

		CRect rtWindow;
		pWnd->GetWindowRect (rtWindow);
		pWnd->ScreenToClient (rtWindow);

		CRect rtClient;
		pWnd->GetClientRect (rtClient);

		CPoint ptWinTopLeft = rtWindow.TopLeft ();

		rtClient.OffsetRect (-ptWinTopLeft);
		dc.ExcludeClipRect (rtClient);

		rtWindow.OffsetRect (-rtWindow.TopLeft ());

        BOOL bActive = IsWindowActive (pWnd);

		const DWORD dwStyle = pWnd->GetStyle ();
		CRect rectCaption (rtWindow);
		CSize szSysBorder = GetSysBorderSize(pWnd);

		BOOL bDialog = pWnd->IsKindOf (RUNTIME_CLASS (CBCGPDialog)) || pWnd->IsKindOf (RUNTIME_CLASS (CBCGPPropertySheet));

		if (!pWnd->IsIconic ())
		{
			rectCaption.bottom = rectCaption.top + szSysBorder.cy;
		}

		BOOL bMaximized = (dwStyle & WS_MAXIMIZE) == WS_MAXIMIZE;

		if (!bRibbonCaption)
		{
			const DWORD dwStyleEx = pWnd->GetExStyle ();
			const BOOL bIsSmallCaption = (dwStyleEx & WS_EX_TOOLWINDOW) != 0;
			const BOOL bIsIconic = pWnd->IsIconic();

			if (!bIsIconic)
			{
				rectCaption.bottom += bIsSmallCaption ? ::GetSystemMetrics (SM_CYSMCAPTION) : ::GetSystemMetrics (SM_CYCAPTION);
			}
			else
			{
				rectCaption.bottom -= 1;
			}

			BOOL bDestroyIcon = FALSE;
			HICON hIcon = globalUtils.GetWndIcon (pWnd, &bDestroyIcon, FALSE);

			CString strText;
			pWnd->GetWindowText (strText);

			if (!bIsIconic && rtClient.top <= szSysBorder.cy)
			{
				rectCaption.bottom = rectCaption.top + szSysBorder.cy;
			}

			if (bMaximized)
			{
				rectCaption.InflateRect (szSysBorder.cx, szSysBorder.cy, szSysBorder.cx, 0);
			}

			m_pCurrentNcWnd = pWnd;
			DrawNcCaption (&dc, rectCaption, dwStyle, dwStyleEx, strText, hIcon, bActive, lstSysButtons, szSysBorder);
			m_pCurrentNcWnd = NULL;

			if (bDestroyIcon)
			{
				::DestroyIcon (hIcon);
			}

			if (bMaximized)
			{
				return TRUE;
			}
		}
#ifndef BCGP_EXCLUDE_RIBBON
		else
		{
			if (bMaximized)
			{
				return TRUE;
			}

			BOOL bHide  = (pBar->GetHideFlags () & BCGPRIBBONBAR_HIDE_ALL) != 0;

			if (bHide && !pBar->IsBackstageViewActive())
			{
				rectCaption.bottom -= szSysBorder.cy + 1;
			}
			else
			{
				rectCaption.bottom += pBar->GetCaptionHeight ();
			}

			CRect rtFrame (rectCaption);

			rtFrame.DeflateRect(1, 1, 1, 0);

			if (!pBar->IsBackstageViewActive())
			{
				clrRibbonFrame = GetRibbonFrameColor(pBar);
			}

			if (clrRibbonFrame != (COLORREF)-1)
			{
				dc.FillSolidRect(rtFrame, clrRibbonFrame);
				nRibbonTopAreaHeight = pBar->GetTabsHeight();
			}
			else
			{
				if (m_brFrame.GetSafeHandle() != NULL)
				{
					dc.FillRect(rtFrame, &m_brFrame);
				}
				else
				{
					dc.FillRect(rtFrame, &globalData.brBarFace);
				}
			}

			if (pBar->IsBackstageViewActive())
			{
				CRect rectBackstageLeft = rtFrame;

				rectBackstageLeft.right = rectBackstageLeft.left + pBar->GetBackstageCommandsAreaWidth() + szSysBorder.cx - 1 - pBar->GetBackstageHorzOffset();

				OnFillRibbonBackstageLeftPane(&dc, rectBackstageLeft);
			}

			rtFrame = rectCaption;
			rtFrame.bottom++;
			
			COLORREF clrBorder = GetNCBorderColor(bActive);
			dc.Draw3dRect(rtFrame, clrBorder, clrBorder);
		}
#endif // BCGP_EXCLUDE_RIBBON

		rtWindow.top = rectCaption.bottom;

		dc.ExcludeClipRect (rectCaption);

		CRect rtFrame (rtWindow);

		rtFrame.DeflateRect(1, 0, 1, 1);

		if (m_brFrame.GetSafeHandle() != NULL)
		{
			dc.FillRect(rtFrame, &m_brFrame);
		}
		else
		{
			dc.FillRect(rtFrame, bDialog ? &GetDlgBackBrush (pWnd) : &globalData.brBarFace);
		}

		if (clrRibbonFrame != (COLORREF)-1 && nRibbonTopAreaHeight > 0)
		{
			CRect rectRibbon = rtFrame;
			rectRibbon.bottom = rectRibbon.top + nRibbonTopAreaHeight;

			dc.FillSolidRect(rectRibbon, clrRibbonFrame);
		}

		rtFrame = rtWindow;
		rtFrame.top--;

		COLORREF clrBorder = GetNCBorderColor(bActive);
		dc.Draw3dRect(rtFrame, clrBorder, clrBorder);

		if (bDialog)
		{
			CBCGPPropertySheet* pPropSheet = DYNAMIC_DOWNCAST(CBCGPPropertySheet, pWnd);
			if (pPropSheet != NULL && pPropSheet->GetAeroHeight() > 0 && pPropSheet->IsVisualManagerStyle())
			{
				CRect rectAero;
				pPropSheet->GetClientRect (rectAero);
				
				int nHeight = pPropSheet->GetAeroHeight();
				rectAero.top = rtWindow.top;
				rectAero.bottom  = rectAero.top + nHeight;
				rectAero.left   = rtWindow.left + 1;
				rectAero.right  = rtWindow.right - 1;
			
				BOOL bDrawLine = TRUE;
				OnFillPropSheetHeaderArea(&dc, pPropSheet, rectAero, bDrawLine);
			}

			if (rgn.GetSafeHandle() != NULL)
			{
				dc.SelectClipRgn (NULL);
			}

			return TRUE;
		}

		//-------------------------------
		// Find status bar extended area:
		//-------------------------------
		if (m_brFrame.GetSafeHandle() == NULL)
		{
			BOOL bBottomFrame = FALSE;
			BOOL bIsStatusBar = FALSE;

			CWnd* pStatusBar = pWnd->GetDescendantWindow (AFX_IDW_STATUS_BAR, TRUE);

			if (pStatusBar->GetSafeHwnd () != NULL && pStatusBar->IsWindowVisible () && !bIsBackstageViewMode)
			{
				CBCGPStatusBar* pClassicStatusBar = DYNAMIC_DOWNCAST (
					CBCGPStatusBar, pStatusBar);
				if (pClassicStatusBar != NULL)
				{
					bIsStatusBar = TRUE;
				}
#ifndef BCGP_EXCLUDE_RIBBON
				else
				{
					CBCGPRibbonStatusBar* pRibbonStatusBar = DYNAMIC_DOWNCAST (
						CBCGPRibbonStatusBar, pStatusBar);
					if (pRibbonStatusBar != NULL)
					{
						bBottomFrame = pRibbonStatusBar->IsBottomFrame ();
						bIsStatusBar = TRUE;
					}
				}
#endif // BCGP_EXCLUDE_RIBBON
			}

			if (bIsStatusBar)
			{
				CRect rectStatus;
				pStatusBar->GetClientRect (rectStatus);
				
				int nHeight = rectStatus.Height ();
				rectStatus.bottom = rtWindow.bottom;
				rectStatus.top    = rectStatus.bottom - nHeight - (bBottomFrame ? 0 : szSysBorder.cy);
				rectStatus.left   = rtWindow.left;
				rectStatus.right  = rtWindow.right;

				OnFillStatusBarFrame(&dc, rectStatus, bActive);
			}
		}

#ifndef BCGP_EXCLUDE_RIBBON
		if (pBar != NULL && pBar->IsBackstageViewActive())
		{
			CRect rectBackstageLeft;
			pWnd->GetWindowRect(rectBackstageLeft);
			
			rectBackstageLeft.OffsetRect(-rectBackstageLeft.TopLeft());

			if (pBar->IsBackstageHorzScroll())
			{
				if (pBar->GetBackstageCommandsAreaWidth() + szSysBorder.cx > pBar->GetBackstageHorzOffset())
				{
					rectBackstageLeft.right = rectBackstageLeft.left + szSysBorder.cx;
				}
				else
				{
					rectBackstageLeft.SetRectEmpty();
				}
			}
			else
			{
				rectBackstageLeft.right = rectBackstageLeft.left + pBar->GetBackstageCommandsAreaWidth() + szSysBorder.cx;
			}

			if (!rectBackstageLeft.IsRectEmpty())
			{
				rectBackstageLeft.left++;
				OnFillRibbonBackstageLeftPane(&dc, rectBackstageLeft);
			}
		}
#endif
		if (rgn.GetSafeHandle() != NULL)
		{
			dc.SelectClipRgn (NULL);
		}

		return TRUE;
	}

	return CBCGPVisualManagerVS2010::OnNcPaint (pWnd, lstSysButtons, rectRedraw);
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnFillStatusBarFrame(CDC* pDC, CRect rectStatus, BOOL /*bActive*/)
{
	ASSERT_VALID(pDC);
	pDC->FillRect(rectStatus, &m_brStatusBar);
}
//*****************************************************************************
CSize CBCGPVisualManagerVS2012::GetNcBtnSize (BOOL /*bSmall*/) const
{
	if (globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode () && globalData.bIsWindows8)
	{
		NONCLIENTMETRICS ncm;
		globalData.GetNonClientMetrics(ncm);
		
		CSize size(ncm.iCaptionWidth, ncm.iCaptionHeight);
		if (size.cx > 0 && size.cy > 0)
		{
			return size;
		}
	}

	return CSize (::GetSystemMetrics (SM_CXMENUSIZE) + globalUtils.ScaleByDPI(8), ::GetSystemMetrics (SM_CYMENUSIZE));
}
//*****************************************************************************
CBrush* CBCGPVisualManagerVS2012::GetNcCaptionBrush(BOOL /*bActive*/, BOOL bSmallDlgCaption)
{
	return bSmallDlgCaption ? &GetDlgBackBrush(m_pCurrentNcWnd) : &globalData.brBarFace;
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::DrawNcCaption (CDC* pDC, CRect rectCaption, 
											   DWORD dwStyle, DWORD dwStyleEx,
											   const CString& strTitle,
											   HICON hIcon, BOOL bActive, 
											   const CObList& lstSysButtons,
											   const CSize& szSysBorder)
{
	const BOOL bIsRTL           = (dwStyleEx & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL;
	const BOOL bIsSmallCaption	= (dwStyleEx & WS_EX_TOOLWINDOW) != 0;
	const int nSysCaptionHeight = bIsSmallCaption ? ::GetSystemMetrics (SM_CYSMCAPTION) : ::GetSystemMetrics (SM_CYCAPTION);

    CDC memDC;
    memDC.CreateCompatibleDC (pDC);
    CBitmap memBmp;
    memBmp.CreateCompatibleBitmap (pDC, rectCaption.Width (), rectCaption.Height ());
    CBitmap* pBmpOld = memDC.SelectObject (&memBmp);

	BOOL bMaximized = (dwStyle & WS_MAXIMIZE) == WS_MAXIMIZE;

	CRect rectBorderCaption (rectCaption);
	if (bMaximized)
	{
		rectBorderCaption.OffsetRect (-rectBorderCaption.TopLeft ());
		rectBorderCaption.bottom -= szSysBorder.cy;
	}

	CRect rect(rectCaption);
	rect.DeflateRect (szSysBorder.cx, szSysBorder.cy, szSysBorder.cx, 0);

	CBrush* pBrush = &m_brFrame;

	if (pBrush->GetSafeHandle() == NULL)
	{
		BOOL bSmallDlgCaption = m_pCurrentNcWnd->GetSafeHwnd() != NULL && rect.Height() < nSysCaptionHeight  &&
			(m_pCurrentNcWnd->IsKindOf(RUNTIME_CLASS(CBCGPDialog)) || m_pCurrentNcWnd->IsKindOf(RUNTIME_CLASS(CBCGPPropertySheet)));

		pBrush = GetNcCaptionBrush(bActive, bSmallDlgCaption);
	}

	memDC.FillRect(rectBorderCaption, pBrush);

	rectBorderCaption.bottom++;

	COLORREF clrBorder = GetNCBorderColor(bActive);
	memDC.Draw3dRect(rectBorderCaption, clrBorder, clrBorder);

	if (nSysCaptionHeight <= rect.Height())
	{
		rect.top = rect.bottom - nSysCaptionHeight - 1;

		// Draw icon:
		if (hIcon != NULL && !bIsSmallCaption)
		{
			CSize szIcon (::GetSystemMetrics (SM_CXSMICON), ::GetSystemMetrics (SM_CYSMICON));

			long x = rect.left + (bMaximized ? szSysBorder.cx : 0) + 2;
			long y = rect.top + max (0, (nSysCaptionHeight - szIcon.cy) / 2);

			::DrawIconEx (memDC.GetSafeHdc (), x, y, hIcon, szIcon.cx, szIcon.cy,
				0, NULL, DI_NORMAL);

			rect.left = x + szIcon.cx + (bMaximized ? szSysBorder.cx : 4);
		}
		else
		{
			rect.left += (bMaximized ? szSysBorder.cx : 0);
		}

		// Draw system buttons:
		int xButtonsRight = rect.right;

		for (POSITION pos = lstSysButtons.GetHeadPosition (); pos != NULL;)
		{
			CBCGPFrameCaptionButton* pButton = (CBCGPFrameCaptionButton*)
				lstSysButtons.GetNext (pos);
			ASSERT_VALID (pButton);

			BCGBUTTON_STATE state = ButtonsIsRegular;

			if (pButton->m_bPushed && pButton->m_bFocused)
			{
				state = ButtonsIsPressed;
			}
			else if (pButton->m_bFocused)
			{
				state = ButtonsIsHighlighted;
			}

			UINT uiHit = pButton->GetHit ();
			UINT nButton = 0;

			switch (uiHit)
			{
			case HTCLOSE_BCG:
				nButton = SC_CLOSE;
				break;

			case HTMAXBUTTON_BCG:
				nButton = 
					(dwStyle & WS_MAXIMIZE) == WS_MAXIMIZE ? SC_RESTORE : SC_MAXIMIZE;
				break;

			case HTMINBUTTON_BCG:
				nButton = 
					(dwStyle & WS_MINIMIZE) == WS_MINIMIZE ? SC_RESTORE : SC_MINIMIZE;
				break;

			case HTHELPBUTTON_BCG:
				nButton = SC_CONTEXTHELP;
				break;
			}

			CRect rectBtn (pButton->GetRect ());
			if (bMaximized)
			{
				rectBtn.OffsetRect (szSysBorder.cx, szSysBorder.cy);
			}

			DrawNcBtn (&memDC, rectBtn, nButton, state, FALSE, bActive, FALSE, pButton->m_bEnabled);

			xButtonsRight = min (xButtonsRight, rectBtn.left);
		}

		// Draw text:
		if (!strTitle.IsEmpty () && rect.left < rect.right)
		{
			CFont* pOldFont = (CFont*)memDC.SelectObject(&GetNcCaptionTextFont());

			CRect rectText = rect;
			rectText.right = xButtonsRight - 1;

			DrawNcText (&memDC, rectText, strTitle, bActive, bIsRTL);

			memDC.SelectObject(pOldFont);
		}
	}

    pDC->BitBlt (rectCaption.left, rectCaption.top, rectCaption.Width (), rectCaption.Height (),
        &memDC, 0, 0, SRCCOPY);

    memDC.SelectObject (pBmpOld);
}
//*****************************************************************************
BOOL CBCGPVisualManagerVS2012::CanSetActivateFlag(CWnd* pWnd) const
{
	ASSERT_VALID(pWnd);

	return pWnd->GetSafeHwnd() != NULL;
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::SetActivateFlag(CWnd* pWnd, BOOL bActive)
{
	
	BOOL bIsMDIFrame = FALSE;
	BOOL bWasActive = FALSE;
	
	// If the active state of an owner-draw MDI frame window changes, we need to
	// invalidate the MDI client area so the MDI child window captions are redrawn.
	if (IsOwnerDrawCaption())
	{
		bIsMDIFrame = pWnd->IsKindOf(RUNTIME_CLASS(CMDIFrameWnd));
		bWasActive = IsWindowActive(pWnd);
	}
	
	m_ActivateFlag[(UINT_PTR)pWnd->GetSafeHwnd()] = bActive;
	pWnd->SendMessage(WM_NCPAINT, 0, 0);
	
	if (IsOwnerDrawCaption() && bIsMDIFrame && (bWasActive != bActive))
	{
		RedrawMDIArea((CMDIFrameWnd*)pWnd);
	}
}
//*****************************************************************************
CFont& CBCGPVisualManagerVS2012::GetNcCaptionTextFont()
{
	return globalData.fontRegular;
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetNcCaptionTextColor(BOOL bActive, BOOL bTitle) const
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::GetNcCaptionTextColor(bActive, bTitle);
	}

	if (m_clrFrameTextActive != (COLORREF)-1)
	{
		return (bActive || m_clrFrameTextInactive == (COLORREF)-1) ? m_clrFrameTextActive : m_clrFrameTextInactive;
	}

	return bActive ? m_clrNcTextActive : m_clrNcTextInactive;
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::DrawNcBtn(CDC* pDC, const CRect& rect, UINT nButton, 
										BCGBUTTON_STATE state, BOOL /*bSmall*/, 
										BOOL bActive, BOOL /*bMDI*/, BOOL bEnabled)
{
	ASSERT_VALID(pDC);

	if (state != ButtonsIsRegular)
	{
		if (state == ButtonsIsPressed)
		{
			OnFillHighlightedArea (pDC, rect, &m_brHighlightDnNC, NULL);
		}
		else
		{
			OnFillHighlightedArea (pDC, rect, &m_brHighlightNC, NULL);
		}
	}
	
	CBCGPMenuImages::IMAGES_IDS imageID;
	
	switch (nButton)
	{
	case SC_CLOSE:
		imageID = CBCGPMenuImages::IdClose;
		break;
		
	case SC_MINIMIZE:
		imageID = CBCGPMenuImages::IdMinimize;
		break;
		
	case SC_MAXIMIZE:
		imageID = CBCGPMenuImages::IdMaximize;
		break;
		
	case SC_RESTORE:
		imageID = CBCGPMenuImages::IdRestore;
		break;
		
	case SC_CONTEXTHELP:
		imageID = CBCGPMenuImages::IdHelp;
		break;

	default:
		return;
	}

	if (m_clrFrame != (COLORREF)-1)
	{
		if (state == ButtonsIsRegular)
		{
			CBCGPMenuImages::DrawByColor(pDC, imageID, rect, m_clrFrame, TRUE);
		}
		else
		{
			CBCGPMenuImages::DrawByColor(pDC, imageID, rect, state == ButtonsIsPressed ? m_clrHighlightDn : m_clrHighlight, TRUE);
		}
	}
	else
	{
		CBCGPMenuImages::Draw(pDC, imageID, rect, GetNcBtnState(state, bActive, bEnabled));

		if (m_Style == VS2012_LightBlue && state != ButtonsIsRegular)
		{
			pDC->Draw3dRect(rect, m_clrMenuItemBorder, m_clrMenuItemBorder);
		}
	}
}
//*****************************************************************************
CBCGPMenuImages::IMAGE_STATE CBCGPVisualManagerVS2012::GetNcBtnState(BCGBUTTON_STATE state, BOOL bActive, BOOL bEnabled)
{
	if (state == ButtonsIsHighlighted && m_Style != VS2012_Dark)
	{
		return CBCGPMenuImages::ImageBlack2;
	}
	else
	{
		return (bActive && bEnabled) ? ((m_Style == VS2012_Dark || (state == ButtonsIsPressed && m_Style == VS2012_Light)) ? CBCGPMenuImages::ImageWhite : CBCGPMenuImages::ImageBlack) : (m_Style == VS2012_Dark ? CBCGPMenuImages::ImageDkGray : CBCGPMenuImages::ImageLtGray);
	}
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::DrawNcText (CDC* pDC, CRect& rect, 
										 const CString& strTitle, 
										 BOOL bActive, BOOL bIsRTL)
{
	if (strTitle.IsEmpty () || rect.right <= rect.left)
	{
		return;
	}

	ASSERT_VALID (pDC);

	int nOldMode = pDC->SetBkMode (TRANSPARENT);
	COLORREF clrOldText = pDC->GetTextColor ();

	DWORD dwTextStyle = DT_END_ELLIPSIS | DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX |
		(bIsRTL ? DT_RTLREADING : 0);

	COLORREF clrText = GetNcCaptionTextColor(bActive, TRUE);

	int width = pDC->GetTextExtent (strTitle).cx;

	const int nTextMargin = 4;

	rect.left += nTextMargin;
	rect.right = min (rect.left + width, rect.right);

	if (rect.right > rect.left)
	{
		pDC->SetTextColor (clrText);
		pDC->DrawText (strTitle, rect, dwTextStyle);
	}

	pDC->SetBkMode    (nOldMode);
	pDC->SetTextColor (clrOldText);
}
//**************************************************************************************************************
void CBCGPVisualManagerVS2012::OnFillBackground(CDC* pDC, CRect rect, CWnd* pWnd)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnFillBackground(pDC, rect, pWnd);
		return;
	}

	pDC->FillRect(rect, &m_brTabs);
}
//**************************************************************************************************************
void CBCGPVisualManagerVS2012::OnFillBarBackground (CDC* pDC, CBCGPBaseControlBar* pBar,
								CRect rectClient, CRect rectClip,
								BOOL bNCArea)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnFillBarBackground (pDC, pBar, rectClient, rectClip, bNCArea);
		return;
	}

	ASSERT_VALID(pBar);
	ASSERT_VALID(pDC);

	if (pBar->IsOnGlass ())
	{
		pDC->FillSolidRect (rectClient, RGB (0, 0, 0));
		return;
	}

	if (pBar->IsKindOf (RUNTIME_CLASS (CBCGPStatusBar)))
	{
		pDC->FillRect(rectClient, &m_brStatusBar);
		return;
	}
	
	if (rectClip.IsRectEmpty ())
	{
		rectClip = rectClient;
	}

#ifndef BCGP_EXCLUDE_RIBBON
	if (pBar->IsKindOf (RUNTIME_CLASS (CBCGPRibbonStatusBar)))
	{
		pDC->FillRect(rectClient, &m_brStatusBar);
		return;
	}

	if (IsRibbonBackstageWhiteBackground() && pBar->IsKindOf (RUNTIME_CLASS (CBCGPRibbonBar)))
	{
		CBCGPRibbonBar* pRibbonBar = DYNAMIC_DOWNCAST(CBCGPRibbonBar, pBar);
		ASSERT_VALID(pRibbonBar);

		if (pRibbonBar->IsBackstageViewActive())
		{
			pDC->FillRect(rectClient, &m_brWhite);
			return;
		}

		if (pRibbonBar->IsDialogMode())
		{
			pDC->FillRect(rectClient, &GetDlgBackBrush(pRibbonBar));
			return;
		}
	}

#endif
	
	if (DYNAMIC_DOWNCAST (CBCGPReBar, pBar) != NULL || DYNAMIC_DOWNCAST (CBCGPReBar, pBar->GetParent ()))
	{
		FillRebarPane (pDC, pBar, rectClient);
		return;
	}

    if (DYNAMIC_DOWNCAST(CBCGPColorBar, pBar) != NULL || DYNAMIC_DOWNCAST(CBCGPCalculator, pBar) != NULL)
	{
		if (pBar->IsDialogControl ())
		{
			pDC->FillRect(rectClient, &globalData.brBtnFace);
			return;
		}
		else if (DYNAMIC_DOWNCAST(CDialog, pBar->GetParent()) != NULL)
		{
			pDC->FillRect(rectClient, &GetDlgBackBrush(pBar->GetParent()));
			return;
		}
		else if (m_Style == VS2012_LightBlue)
		{
			pDC->FillRect(rectClient, &m_brDlgBackground);
			return;
		}
	}

    if (DYNAMIC_DOWNCAST(CBCGPCalendarBar, pBar) != NULL && m_Style == VS2012_LightBlue)
	{
		pDC->FillRect(rectClient, &m_brDlgBackground);
		return;
	}

    if (DYNAMIC_DOWNCAST(CBCGPPopupMenuBar, pBar) != NULL)
	{
		if (m_Style == VS2012_LightBlue)
		{
			CBCGPVisualManagerVS2010::OnFillBarBackground (pDC, pBar, rectClient, rectClip, bNCArea);
		}
		else
		{
			pDC->FillRect (rectClient, &m_brMenuLight);
		}
		return;
	}

    if (pBar->IsKindOf (RUNTIME_CLASS (CBCGPCaptionBar)))
	{
		if (m_Style == VS2012_LightBlue)
		{
			CBCGPVisualManagerVS2010::OnFillBarBackground (pDC, pBar, rectClient, rectClip, bNCArea);
			return;
		}

		CBCGPCaptionBar* pCaptionBar = (CBCGPCaptionBar*) pBar;

		if (!pCaptionBar->IsMessageBarMode() && pCaptionBar->m_clrBarBackground != -1)
		{
			pDC->FillSolidRect(rectClip, pCaptionBar->m_clrBarBackground);
		}
		else
		{
			pDC->FillRect(rectClip, &globalData.brBarFace);
		}
		return;
	}

#if !defined(BCGP_EXCLUDE_TOOLBOX) && !defined(BCGP_EXCLUDE_TASK_PANE)
 	if (pBar->IsKindOf (RUNTIME_CLASS (CBCGPToolBoxPage)) || pBar->IsKindOf (RUNTIME_CLASS (CBCGPToolBox)) || pBar->IsKindOf (RUNTIME_CLASS (CBCGPToolBoxEx)))
	{
		pDC->FillRect(rectClip, &m_brControl);
		return;
	}
#endif

	if (pBar->IsKindOf(RUNTIME_CLASS (CBCGPOutlookBarPane)))
	{
		CBCGPOutlookBarPane* pOlBar = DYNAMIC_DOWNCAST (CBCGPOutlookBarPane, pBar);
		ASSERT_VALID (pOlBar);
		
		if (pOlBar->IsBackgroundTexture ())
		{
			CBCGPVisualManagerVS2010::OnFillBarBackground (pDC, pBar, rectClient, rectClip);
		}
		else
		{
			pDC->FillSolidRect(rectClip, globalData.clrBarLight);
		}
		return;
	}

	if (pBar->IsDialogControl ())
	{
		pDC->FillRect(rectClient, &globalData.brBtnFace);
		return;
	}

	if (m_Style == VS2012_LightBlue && (pBar->IsKindOf (RUNTIME_CLASS (CBCGPAutoHideDockBar)) || pBar->IsKindOf (RUNTIME_CLASS (CBCGPAutoHideToolBar))))
	{
		CBCGPVisualManagerVS2010::OnFillBarBackground (pDC, pBar, rectClient, rectClip, bNCArea);
		return;
	}

	BOOL bIsMenuBar = pBar->IsKindOf (RUNTIME_CLASS (CBCGPMenuBar));

	if (!bIsMenuBar && DYNAMIC_DOWNCAST(CBCGPToolBar, pBar) != NULL && m_clrToolBarGradientLight != globalData.clrBarFace)
	{
		CBrush br(m_clrToolBarGradientLight);
		pDC->FillRect(rectClient, &br);
	}
	else
	{
		pDC->FillRect(rectClient, &globalData.brBarFace);
	}

	// If menu bar is docked on top, under the caption bar, draw horizontal line on top
	if (IsOwnerDrawCaption() && bIsMenuBar && bNCArea && (pBar->GetCurrentAlignment() & CBRS_ALIGN_TOP) && IsDrawLineOnTopOfMenuBar())
	{
		CBCGPDockBarRow* pDockRow = pBar->GetDockRow();
		if (pDockRow != NULL && pDockRow->GetRowOffset() == 0)
		{
			CBCGPPenSelector ps(*pDC, globalData.clrBarShadow);
			
			pDC->MoveTo(rectClient.left, rectClient.top);
			pDC->LineTo(rectClient.right, rectClient.top);
		}
	}

	if (m_Style == VS2012_LightBlue && bNCArea)
	{
		CBCGPToolBar* pToolBar = DYNAMIC_DOWNCAST (CBCGPToolBar, pBar);

		if (pToolBar->GetSafeHwnd() != NULL && pToolBar->IsDocked () &&
			pToolBar->GetParentDockBar () != NULL && !bIsMenuBar)
		{
			COLORREF clrBorder = RGB(220, 224, 236);

			rectClient.DeflateRect(1, 0);
			pDC->Draw3dRect(rectClient, clrBorder, clrBorder);
		}
	}
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawControlBorder (CDC* pDC, CRect rect, CWnd* pWndCtrl, BOOL bDrawOnGlass)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawControlBorder (pDC, rect, pWndCtrl, bDrawOnGlass);
		return;
	}

	ASSERT_VALID (pDC);

	COLORREF clrInner = IsDarkTheme() ? globalData.clrBarLight : globalData.clrBarHilite;
	COLORREF clrBorder = m_clrSeparator;

	if (pWndCtrl->GetSafeHwnd() != NULL && (DYNAMIC_DOWNCAST(CBCGPEdit, pWndCtrl) != NULL || DYNAMIC_DOWNCAST(CBCGPComboBox, pWndCtrl) != NULL))
	{
		CBCGPComboBox* pCombo = DYNAMIC_DOWNCAST(CBCGPComboBox, pWndCtrl);
		BOOL bIsDroppedDown = (pCombo != NULL && pCombo->GetDroppedState());

		CWnd* pFocus = CWnd::GetFocus();

		if (bIsDroppedDown || (pFocus->GetSafeHwnd() != NULL && (pWndCtrl->IsChild(pFocus) || pFocus->GetSafeHwnd() == pWndCtrl->GetSafeHwnd())))
		{
			clrBorder = m_clrBorderActive;
		}
		else if (!pWndCtrl->IsWindowEnabled())
		{
			clrBorder = m_clrTextDisabled;
		}
	}

	if (bDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawRect (rect, (COLORREF)-1, clrBorder);

		rect.DeflateRect (1, 1);

		dm.DrawRect (rect, (COLORREF)-1, clrInner);
	}
	else
	{
		pDC->Draw3dRect (rect, clrBorder, clrBorder);

		rect.DeflateRect (1, 1);
		pDC->Draw3dRect (rect, clrInner, clrInner);
	}
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnHighlightMenuItem (CDC* pDC, CBCGPToolbarMenuButton* pButton,
											CRect rect, COLORREF& clrText)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnHighlightMenuItem(pDC, pButton, rect, clrText);
		return;
	}

	if ((pButton->m_nStyle & TBBS_DISABLED) && globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode ())
	{
		rect.DeflateRect(1, 0);
		pDC->Draw3dRect(rect, m_clrMenuSeparator, m_clrMenuSeparator);
		return;
	}

	CBrush* pBrush = (pButton->m_nStyle & TBBS_DISABLED) ? &m_brMenuLight : &m_brHighlight;
	
	rect.DeflateRect (2, 0);
	OnFillHighlightedArea (pDC, rect, pBrush, pButton);
	
	clrText = GetHighlightedMenuItemTextColor (pButton);
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnEraseTabsArea (CDC* pDC, CRect rect, const CBCGPBaseTabWnd* pTabWnd)
{
	if (m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnEraseTabsArea (pDC, rect, pTabWnd);
		return;
	}

	CBCGPVisualManagerXP::OnEraseTabsArea (pDC, rect, pTabWnd);
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnEraseTabsButton (CDC* pDC, CRect rect,
											  CBCGPButton* pButton,
											  CBCGPBaseTabWnd* pBaseTab)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);
	ASSERT_VALID (pBaseTab);
	
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnEraseTabsButton (pDC, rect, pButton, pBaseTab);
		return;
	}

	if (pBaseTab->IsPureFlatTab() && !pBaseTab->IsMDITab())
	{
		if (pButton->IsHighlighted())
		{
			pDC->FillRect(rect, pButton->IsPressed() ? &m_brHighlight : &m_brMenuItemCheckedHighlight);
		}
		else
		{
			if (pBaseTab->IsDialogControl())
			{
				pDC->FillRect(rect, &GetDlgBackBrush(pBaseTab->GetParent()));
			}
			else
			{
				pDC->FillRect(rect, &globalData.brBarFace);
			}
		}
	}
	else
	{
		if (pButton->IsPressed ())
		{
			pDC->FillRect(rect, &m_brHighlightDn);
		}
		else if (pButton->IsHighlighted())
		{
			pDC->FillRect(rect, &m_brHighlight);
		}
		else
		{
			CBCGPVisualManagerXP::OnEraseTabsButton (pDC, rect, pButton, pBaseTab);
		}
	}
}
//*********************************************************************************************************
BOOL CBCGPVisualManagerVS2012::OnEraseTabsFrame (CDC* pDC, CRect rect, const CBCGPBaseTabWnd* pTabWnd)
{
	ASSERT_VALID (pDC);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () ||
		pTabWnd->IsDialogControl (TRUE) ||
		pTabWnd->IsFlatTab () || pTabWnd->IsOneNoteStyle () || pTabWnd->IsVS2005Style () || pTabWnd->IsPureFlatTab())
	{
		return CBCGPVisualManagerVS2008::OnEraseTabsFrame (pDC, rect, pTabWnd);
	}

	pDC->FillRect(rect, &globalData.brBarFace);
	return TRUE;
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTabBorder(CDC* pDC, CBCGPTabWnd* pTabWnd, CRect rectBorder, COLORREF clrBorder,
										 CPen& penLine)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pTabWnd);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () ||
		(pTabWnd->IsDialogControl (TRUE) && !pTabWnd->IsPropertySheetTab()) || !pTabWnd->IsMDITab() ||
		pTabWnd->IsFlatTab () || pTabWnd->IsOneNoteStyle () || pTabWnd->IsVS2005Style ())
	{
		CBCGPVisualManagerVS2010::OnDrawTabBorder(pDC, pTabWnd, rectBorder, clrBorder, penLine);
		return;
	}

	COLORREF clrActiveTab = pTabWnd->GetTabBkColor (pTabWnd->GetActiveTab ());
	if (clrActiveTab == (COLORREF)-1)
	{
		clrActiveTab =  pTabWnd->IsMDIFocused() && pTabWnd->IsActiveInMDITabGroup() ? 
			(m_Style == VS2012_LightBlue ? m_clrHighlight : m_clrAccent) : 
			(m_Style == VS2012_LightBlue ? m_clrMenuLight : globalData.clrBarShadow);
	}

	CRect rectTop = rectBorder;
	rectTop.DeflateRect(pTabWnd->GetTabBorderSize() + 1, 0);
	rectTop.bottom = rectTop.top + pTabWnd->GetTabBorderSize();
	rectTop.top++;

	pDC->FillSolidRect(rectTop, pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM ? globalData.clrBarFace : clrActiveTab);

	CRect rectBottom = rectBorder;
	rectBottom.DeflateRect(pTabWnd->GetTabBorderSize() + 1, 0);
	rectBottom.top = rectBottom.bottom - pTabWnd->GetTabBorderSize();
	rectBottom.bottom--;

	pDC->FillSolidRect(rectBottom, pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM ? clrActiveTab : globalData.clrBarFace);
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTab (CDC* pDC, CRect rectTab,
						int iTab, BOOL bIsActive, const CBCGPBaseTabWnd* pTabWnd)
{
	ASSERT_VALID (pTabWnd);
	ASSERT_VALID (pDC);

	const BOOL bIsHighlight = (iTab == pTabWnd->GetHighlightedTab ());

	if (pTabWnd->IsVS2005Style ())
	{
		CBCGPVisualManagerXP::OnDrawTab (pDC, rectTab, iTab, bIsActive, pTabWnd);
		return;
	}

	const BOOL bIsColored3DTab = pTabWnd->GetTabBkColor(iTab) != (COLORREF)-1 && pTabWnd->Is3DStyle();

	if (globalData.m_nBitsPerPixel <= 8 || (globalData.IsHighContastMode () && !bIsColored3DTab) || (m_Style == VS2012_LightBlue && !bIsColored3DTab) ||
		(pTabWnd->IsDialogControl (TRUE) && !pTabWnd->IsPropertySheetTab()) ||
		pTabWnd->IsFlatTab () || pTabWnd->IsOneNoteStyle () || pTabWnd->IsLeftRightRounded() || pTabWnd->IsPureFlatTab() ||
		(pTabWnd->GetTabBkColor(iTab) != (COLORREF)-1 && !pTabWnd->Is3DStyle()))
	{
		COLORREF clrHighlightMenuItemSaved = m_clrHighlightMenuItem;
		if (bIsHighlight && pTabWnd->IsOneNoteStyle ())
		{
			if (m_Style != VS2012_Dark || bIsActive)
			{
				m_clrHighlightMenuItem = m_Style == VS2012_Dark ? globalData.clrBarFace : RGB(255, 255, 255);
			}
		}

		CBCGPVisualManagerVS2010::OnDrawTab (pDC, rectTab, iTab, bIsActive, pTabWnd);

		m_clrHighlightMenuItem = clrHighlightMenuItemSaved;
		return;
	}

	if (!bIsActive && iTab != 0 && !pTabWnd->IsMDITab() && pTabWnd->GetTabBkColor(iTab) == (COLORREF)-1)
	{
		if (pTabWnd->GetActiveTab() != iTab - 1)
		{
			CBCGPPenSelector ps(*pDC, m_Style == VS2012_Dark ? &m_penSeparatorLight : &m_penSeparator);

			pDC->MoveTo(rectTab.left, rectTab.top + 1);
			pDC->LineTo(rectTab.left, rectTab.bottom - 1);
		}
		
		rectTab.left++;
	}

	BOOL bIsActiveInMDITabGroup = pTabWnd->IsActiveInMDITabGroup();
	if (pTabWnd->IsMultipleSelection() && pTabWnd->IsMultipleSelectionSupported() && bIsActive)
	{
		if (pTabWnd->GetActiveTab() != iTab)
		{
			// Tab is selected
			bIsActiveInMDITabGroup = FALSE;
		}
	}

	if (bIsActive || bIsHighlight)
	{
		CRect rectFill = rectTab;

		COLORREF clrTab = pTabWnd->GetTabBkColor(iTab);
		if (clrTab == (COLORREF)-1 || pTabWnd->Is3DStyle())
		{
			clrTab = globalData.clrBarLight;

			if (bIsActive)
			{
				if (pTabWnd->IsMDITab())
				{
					if (pTabWnd->IsMDIFocused() && bIsActiveInMDITabGroup)
					{
						clrTab = (bIsColored3DTab && m_Style == VS2012_LightBlue) ? m_clrHighlightMenuItem : m_clrAccent;
					}
					else
					{
						clrTab = IsDarkTheme() ? globalData.clrBarLight : globalData.clrBarShadow;
					}
				}
				else
				{
					clrTab = m_clrHighlight;
				}
			}
			else if (pTabWnd->IsMDITab())
			{
				clrTab = (bIsColored3DTab && m_Style == VS2012_LightBlue) ? globalData.clrBarDkShadow : m_clrAccentLight;
			}
		}

		if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM)
		{
			rectFill.top--;
		}
		else
		{
			rectFill.bottom++;
		}

		if (bIsColored3DTab)
		{
			rectFill.DeflateRect(1, 0);
		}

		pDC->FillSolidRect(rectFill, clrTab);
	}

	COLORREF clrTabText = pTabWnd->GetTabTextColor(iTab);

	if (bIsColored3DTab)
	{
		CRect rectColor = rectTab;
		const int nColorBarHeight = max(3, rectTab.Height() / 6);

		if (pTabWnd->GetLocation() == CBCGPBaseTabWnd::LOCATION_BOTTOM)
		{
			rectColor.top = rectColor.bottom - nColorBarHeight;
			rectTab.bottom -= nColorBarHeight;
		}
		else
		{
			rectColor.bottom = rectColor.top + nColorBarHeight;
			rectTab.top += nColorBarHeight;
		}

		rectColor.DeflateRect(1, 0);

		CBrush br(pTabWnd->GetTabBkColor(iTab));
		pDC->FillRect(rectColor, &br);

		if (m_Style == VS2012_LightBlue && pTabWnd->IsMDITab())
		{
			if (!bIsActive)
			{
				clrTabText = RGB(255, 255, 255);
			}
			else
			{
				clrTabText = RGB(0, 0, 0);
			}
		}
	}

	if (clrTabText == (COLORREF)-1)
	{
		if (pTabWnd->IsMDITab())
		{
			if (bIsActive)
			{
				if (pTabWnd->IsMDIFocused() && bIsActiveInMDITabGroup)
				{
					clrTabText = RGB(255, 255, 255);
				}
				else
				{
					clrTabText = globalData.clrBarText;
				}
			}
			else if (bIsHighlight)
			{
				clrTabText = RGB(255, 255, 255);
			}
			else
			{
				clrTabText = globalData.clrBarText;
			}
		}
		else
		{
			clrTabText = (bIsActive || bIsHighlight) ? m_clrAccentText : globalData.clrBarText;
		}
	}

	OnDrawTabContent(pDC, rectTab, iTab, bIsActive, pTabWnd, clrTabText);
}
//**********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetActiveTabBackColor(const CBCGPBaseTabWnd* pTabWnd) const
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || 
		(pTabWnd->IsDialogControl(TRUE) && !pTabWnd->IsPropertySheetTab()) || !IsDarkTheme())
	{
		return CBCGPVisualManagerVS2010::GetActiveTabBackColor(pTabWnd);
	}

	return pTabWnd->IsFlatTab() ? globalData.clrBarLight : m_clrDlgBackground;
}
//**********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTabUnderline(CDC* pDC, CBCGPTabWnd* pTabWnd, CRect rect, int iTab)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawTabUnderline(pDC, pTabWnd, rect, iTab);
		return;
	}

	COLORREF clrFill = m_clrAccentHilight;
	
	if (pTabWnd->IsMDITabGroup() && !(pTabWnd->IsMDIFocused() && pTabWnd->IsActiveInMDITabGroup()))
	{
		clrFill = globalData.clrBarShadow;
	}
	
	CBrush brFill(clrFill);
	pDC->FillRect(rect, &brFill);
}
//**********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTabContent (CDC* pDC, CRect rectTab,
							int iTab, BOOL bIsActive, const CBCGPBaseTabWnd* pTabWnd,
							COLORREF clrText)
{
	COLORREF clrTextOld = (COLORREF)-1;

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue ||
		(pTabWnd->IsDialogControl (TRUE) && !pTabWnd->IsPropertySheetTab()) ||
		(pTabWnd->GetTabBkColor(iTab) != (COLORREF)-1 && !pTabWnd->Is3DStyle()) ||
		clrText != (COLORREF)-1)
	{
		if (globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode ())
		{
			if (pTabWnd->IsPureFlatTab())
			{
				clrTextOld = pDC->SetTextColor((bIsActive || pTabWnd->GetHighlightedTab() == iTab) ? m_clrAccentText : globalData.clrBarText);
			}
			else if (IsDarkTheme() && pTabWnd->GetHighlightedTab() == iTab && pTabWnd->IsOneNoteStyle() && pTabWnd->GetTabBkColor(iTab) != (COLORREF)-1)
			{
				clrTextOld = pDC->SetTextColor(globalData.clrBarText);
			}
		}
	}
	else
	{
		if (!bIsActive)
		{
			if (pTabWnd->IsPureFlatTab() && pTabWnd->GetHighlightedTab() == iTab)
			{
				clrTextOld = pDC->SetTextColor(m_clrAccentText);
			}
			else
			{
				clrTextOld = pDC->SetTextColor(globalData.clrBarText);
			}
		}
		else if (pTabWnd->IsPureFlatTab() || IsDarkTheme())
		{
			clrTextOld = pDC->SetTextColor(m_clrAccentText);
		}
	}

	CBCGPVisualManagerVS2010::OnDrawTabContent (pDC, rectTab, iTab, bIsActive, pTabWnd, clrText);

	if (clrTextOld != (COLORREF)-1)
	{
		pDC->SetTextColor(clrTextOld);
	}
}
//**********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTabButton (CDC* pDC, CRect rect, 
											   const CBCGPBaseTabWnd* pTabWnd,
											   int nID,
											   BOOL bIsHighlighted,
											   BOOL bIsPressed,
											   BOOL bIsDisabled)
{
	if (pTabWnd->IsVS2005Style ())
	{
		CBCGPVisualManagerXP::OnDrawTabButton (pDC, rect, pTabWnd, nID, bIsHighlighted, bIsPressed, bIsDisabled);
		return;
	}

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue ||
		((pTabWnd->GetTabBkColor(m_nCurrDrawedTab) != (COLORREF)-1 && !pTabWnd->Is3DStyle()) && !bIsHighlighted && !bIsPressed) ||
		((pTabWnd->IsOneNoteStyle () || pTabWnd->IsLeftRightRounded() || pTabWnd->IsPureFlatTab()) && m_nCurrDrawedTab == pTabWnd->GetActiveTab() && !bIsHighlighted))
	{
		CBCGPVisualManagerVS2010::OnDrawTabButton (pDC, rect, pTabWnd, nID, bIsHighlighted, bIsPressed, bIsDisabled);
		return;
	}

	rect.DeflateRect(1, 0);

	CBCGPMenuImages::IMAGE_STATE state = (bIsPressed && bIsHighlighted) || m_Style == VS2012_Dark ? CBCGPMenuImages::ImageWhite : 
		CBCGPMenuImages::ImageBlack;

	if (pTabWnd->IsMDITab())
	{
		BOOL bActive = m_nCurrDrawedTab == pTabWnd->GetActiveTab() && pTabWnd->IsMDIFocused() && pTabWnd->IsActiveInMDITabGroup();

		if (bActive || (pTabWnd->GetHighlightedTab() == m_nCurrDrawedTab && pTabWnd->Is3DStyle()))
		{
			state = CBCGPMenuImages::ImageWhite;
		}

		if (bIsPressed && bIsHighlighted)
		{
			if (bActive)
			{
				pDC->FillSolidRect (rect, m_clrAccentText);
			}
			else
			{
				pDC->FillRect(rect, &m_brAccent);
			}

			state = CBCGPMenuImages::ImageWhite;
		}
		else if (bIsHighlighted || bIsPressed)
		{
			pDC->FillRect (rect, bActive ? &m_brAccentLight : &m_brAccentHighlight);
			state = CBCGPMenuImages::ImageWhite;
		}
	}
	else
	{
		if (bIsHighlighted)
		{
			pDC->FillSolidRect(rect, bIsPressed ? m_clrAccent : globalData.clrBarShadow);
		}

		if (pTabWnd->IsFlatTab() && !bIsHighlighted && m_nCurrDrawedTab == pTabWnd->GetActiveTab())
		{
			state = CBCGPMenuImages::ImageBlack;
		}
	}

	CBCGPMenuImages::Draw (pDC, (CBCGPMenuImages::IMAGES_IDS)nID, rect, state);

	if (bIsHighlighted && !pTabWnd->IsMDITab())
	{
		pDC->Draw3dRect (rect, m_clrMenuItemBorder, m_clrMenuItemBorder);
	}
}
//*********************************************************************************************************
BOOL CBCGPVisualManagerVS2012::IsTabColorBar(CBCGPTabWnd* pTabWnd, int iTab)
{
	ASSERT_VALID (pTabWnd);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::IsTabColorBar(pTabWnd, iTab);
	}

	return (pTabWnd->IsPureFlatTab() || pTabWnd->Is3DStyle()) && pTabWnd->GetTabBkColor (iTab) != (COLORREF)-1;
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawAutoHideButtonBorder (CDC* pDC, CRect rect, CRect rectBorderSize, CBCGPAutoHideButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawAutoHideButtonBorder (pDC, rect, rectBorderSize, pButton);
		return;
	}

	DWORD dwAlign = (pButton->GetAlignment () & CBRS_ALIGN_ANY);
	CRect rectBar = rect;

	int nBarSize = 6;
	int nMargin = 3;

	switch (dwAlign)
	{
	case CBRS_ALIGN_LEFT:
		rectBar.right = rectBar.left + nBarSize;
		rectBar.DeflateRect(0, nMargin);
		break;

	case CBRS_ALIGN_RIGHT:
		rectBar.left = rectBar.right - nBarSize;
		rectBar.DeflateRect(0, nMargin);
		break;

	case CBRS_ALIGN_TOP:
		rectBar.bottom = rectBar.top + nBarSize;
		rectBar.DeflateRect(nMargin, 0);
		break;

	case CBRS_ALIGN_BOTTOM:
		rectBar.top = rectBar.bottom - nBarSize;
		rectBar.DeflateRect(nMargin, 0);
		break;
	}

	OnDrawAutohideBar(pDC, rectBar, pButton);
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawAutohideBar(CDC* pDC, CRect rectBar, CBCGPAutoHideButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	pDC->FillSolidRect(rectBar, pButton->IsHighlighted() ? m_clrAutohideBarActive : m_clrAutohideBarInactive);
}
//*********************************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetAutoHideButtonTextColor (CBCGPAutoHideButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetAutoHideButtonTextColor (pButton);
	}

	return pButton->IsHighlighted() ? m_clrAccentText : globalData.clrBarText;
}
//*********************************************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnDrawControlBarCaption (CDC* pDC, CBCGPDockingControlBar* pBar, 
			BOOL bActive, CRect rectCaption, CRect rectButtons)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::OnDrawControlBarCaption (pDC, pBar, bActive, rectCaption, rectButtons);
	}

	ASSERT_VALID (pDC);

	rectCaption.bottom++;

	if (bActive)
	{
		pDC->FillRect(rectCaption, &m_brAccent);
		return RGB(255, 255, 255);
	}
	else
	{
		pDC->FillRect(rectCaption, &globalData.brBarFace);
		return globalData.clrBarText;
	}
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawControlBarCaptionText (CDC* pDC, CBCGPDockingControlBar* pBar, BOOL bActive, const CString& strTitle, CRect& rectCaption)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawControlBarCaptionText(pDC, pBar, bActive, strTitle, rectCaption);
		return;
	}

	ASSERT_VALID (pDC);

	int nTextWidth = 0;

	if (!strTitle.IsEmpty())
	{
		CRect rectText = rectCaption;

		pDC->DrawText (strTitle, rectText, DT_LEFT | DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);

		CSize sizeText = pDC->GetTextExtent(strTitle);
		nTextWidth = sizeText.cx + sizeText.cy / 2;
	}

	CRect rectGripper = rectCaption;
	rectGripper.left += nTextWidth;

	if (rectGripper.Width() > 10 && m_Style != VS2012_LightBlue)
	{
		COLORREF clrGripper = m_clrGripper;
		COLORREF clrGripperBk = m_clrGripperBk;

		if (bActive)
		{
			m_clrGripper = IsDarkTheme() ? m_clrAccentText : m_clrAccentLight;
			m_clrGripperBk = m_clrAccent;
		}

		OnDrawBarGripper (pDC, rectGripper, TRUE, pBar);

		m_clrGripper = clrGripper;
		m_clrGripperBk = clrGripperBk;
	}
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawBarGripper (CDC* pDC, CRect rectGripper, BOOL bHorz, CBCGPBaseControlBar* pBar)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawBarGripper(pDC, rectGripper, bHorz, pBar);
		return;
	}

	ASSERT_VALID (pDC);

	COLORREF clrTextOld = pDC->SetTextColor (m_clrGripper);
	COLORREF clrBkOld = pDC->SetBkColor (m_clrGripperBk);

	const int nSize = 5;
	const int nSizeBrush = 4;

	if (rectGripper.Width() < 2 || rectGripper.Height() < 2)
	{
		return;
	}

	CRect rectFill = rectGripper;
	
	CBCGPToolBar* pToolBar = DYNAMIC_DOWNCAST (CBCGPToolBar, pBar);
	if (pToolBar != NULL)
	{
		if (bHorz)
		{
			int xCenter = rectFill.CenterPoint ().x;
			rectFill.left = xCenter - 1;
			rectFill.right = rectFill.left + nSize;
			rectFill.DeflateRect (0, 2);
		}
		else
		{
			int yCenter = rectFill.CenterPoint ().y;
			rectFill.top = yCenter - 1;
			rectFill.bottom = rectFill.top + nSize;
			rectFill.DeflateRect (2, 0);
		}
	}
	else
	{
		rectFill.top = rectFill.CenterPoint().y - nSize / 2;
		rectFill.bottom = rectFill.top + nSize;
		bHorz = FALSE;
	}

	int nDelta = (bHorz ? rectFill.Height() : rectFill.Width()) % nSizeBrush;
	if (nDelta != 1)
	{
		nDelta = (nDelta + nSizeBrush - 1) % nSizeBrush;
		if (bHorz)
		{
			rectFill.DeflateRect (0, nDelta / 2, 0, nDelta - nDelta / 2);
		}
		else
		{
			rectFill.DeflateRect (nDelta / 2, 0, nDelta - nDelta / 2, 0);
		}
	}

	CPoint ptBrushOrg(pDC->SetBrushOrg((rectFill.left % nSizeBrush), (rectFill.top % nSizeBrush)));

	CBrush* pBrush = bHorz ? &m_brGripperHorz : &m_brGripperVert;
	pBrush->UnrealizeObject();

	pDC->FillRect (rectFill, pBrush);

	pDC->SetTextColor (clrTextOld);
	pDC->SetBkColor (clrBkOld);
	pDC->SetBrushOrg (ptBrushOrg);
}
//*********************************************************************************************************
BOOL CBCGPVisualManagerVS2012::IsAutoGrayscaleImages()
{
	return m_bAutoGrayscaleImages && globalData.m_nBitsPerPixel > 8 && !globalData.m_bIsBlackHighContrast;
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::ModifyGlobalColors ()
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::ModifyGlobalColors();
		return;
	}

	if (m_Style == VS2012_Light)
	{
		globalData.clrBarFace = RGB (239, 239, 242);
		globalData.clrBarText = RGB (30, 30, 30);
		globalData.clrBarShadow = RGB (204, 206, 219);
		globalData.clrBarHilite = RGB (255, 255, 255);
		globalData.clrBarDkShadow = RGB(162, 164, 165);
		globalData.clrBarLight = RGB (254, 254, 254);
	}
	else if (m_Style == VS2012_Dark)
	{
		globalData.clrBarFace = RGB (45, 45, 48);
		globalData.clrBarText = RGB (241, 241, 241);
		globalData.clrBarShadow = RGB (120, 120, 120);
		globalData.clrBarHilite = RGB (104, 104, 104);
		globalData.clrBarDkShadow = RGB(190, 190, 190);
		globalData.clrBarLight = RGB (90, 90, 90);
	}
	else
	{
		globalData.clrBarFace = RGB (214, 219, 233);
		globalData.clrBarText = RGB (27, 41, 62);
		globalData.clrBarShadow = RGB (190, 195, 203);
		globalData.clrBarHilite = RGB (255, 255, 255);
		globalData.clrBarDkShadow = RGB(133, 145, 162);
		globalData.clrBarLight = RGB (241, 243, 248);
	}

	globalData.brBarFace.DeleteObject ();
	globalData.brBarFace.CreateSolidBrush (globalData.clrBarFace);
}
//************************************************************************************
void CBCGPVisualManagerVS2012::SetupColors()
{
	if (m_curAccentColor == VS2012_Default)
	{
		COLORREF clr = globalData.GetWindowsAccentColor();
		if (clr != (COLORREF)-1)
		{
			m_clrAccent = clr;
		}
	}

	m_clrAccentLight = CBCGPDrawManager::PixelAlpha(m_clrAccent, m_Style != VS2012_Dark ? 140 : 75);
	m_clrAccentHilight = CBCGPDrawManager::PixelAlpha(m_clrAccent, m_Style != VS2012_Dark ? 115 : 85);

	m_clrAccentText = CBCGPDrawManager::PixelAlpha (m_clrAccent, m_Style != VS2012_Dark ? 75 : 140);
	m_clrGripper = m_Style != VS2012_Dark ? RGB(70, 70, 74) : RGB (90, 91, 93);
	
	m_clrTextDisabled = m_Style == VS2012_Dark ? globalData.clrBarShadow : globalData.clrBarDkShadow;
	m_clrEditPrompt = m_Style == VS2012_Dark ? 
		CBCGPDrawManager::ColorMakeDarker(globalData.clrBarText, .35) : CBCGPDrawManager::ColorMakeLighter(globalData.clrWindowText, .45);
	
	m_clrMenuSeparator = m_Style == VS2012_Dark ? globalData.clrBarFace : globalData.clrBarShadow;
	
	m_clrBarBkgnd = globalData.clrBarLight;

	m_clrGripperBk = globalData.clrBarFace;
	m_clrAutohideBarActive = m_clrAccentText;

	if (m_Style == VS2012_LightBlue)
	{
		m_clrHighlight = RGB(255, 240, 208);
		
		m_clrHighlightChecked = m_clrHighlightDn = RGB(255, 232, 166);
		m_clrHighlightMenuItem = m_clrHighlight;
		
		m_clrHighlightGradientLight = m_clrHighlight;
		m_clrHighlightGradientDark = m_clrHighlight;
		
		m_clrHighlightDnGradientDark = m_clrHighlightDn;
		m_clrHighlightDnGradientLight =  m_clrHighlightDn;
		
		m_clrMenuLight = RGB (208, 215, 236);
		m_clrMenuLightGradient = m_clrMenuGutter = RGB (233, 236, 238);
		
		m_clrMenuItemBorder = m_clrPressedButtonBorder = RGB(229, 195, 101);
		m_clrMenuItemGradientDark = m_clrMenuItemGradient1 = m_clrMenuItemGradient2 = m_clrHighlight;

		m_clrEditBoxBorder = globalData.clrBarDkShadow;
		m_clrMenuBorder = RGB(155, 167, 183);

		m_clrControl = RGB(255, 255, 255);
		m_clrDlgBackground = CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace, .06);

		m_clrAutohideBarInactive = m_clrInactiveCaptionGradient2;

		m_clrCaptionBarGradientDark = m_clrInactiveCaptionGradient1;
		m_clrCaptionBarGradientLight = m_clrInactiveCaptionGradient2;
	}
	else
	{
		m_clrHighlight = (m_Style == VS2012_Dark) ? CBCGPDrawManager::ColorMakeDarker(globalData.clrBarLight) : globalData.clrBarLight;

		m_clrHighlightDn = m_clrAccent;
		m_clrHighlightChecked = globalData.clrBarFace;
		m_clrHighlightMenuItem = (m_Style == VS2012_Dark) ? RGB(51, 51, 52) : RGB(248, 249, 250);

		m_clrHighlightGradientLight = globalData.clrBarHilite;
		m_clrHighlightGradientDark = globalData.clrBarHilite;
		
		m_clrHighlightDnGradientDark = m_clrAccent;
		m_clrHighlightDnGradientLight =  m_clrAccent;
		
		m_clrMenuLight = m_Style == VS2012_Light ? RGB(231, 232, 236) : m_Style == VS2012_Dark ? RGB(27, 27, 27) : RGB(233, 236, 238);
		m_clrMenuLightGradient = m_clrMenuGutter = globalData.clrBarLight;
		
		m_clrMenuItemGradientDark = m_clrPressedButtonBorder = m_clrAccent;
		m_clrMenuItemBorder = m_clrHighlight;

		m_clrEditBoxBorder = (m_Style == VS2012_Dark) ? 
			CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace, .5) : 
			CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace, .09);

		m_clrMenuBorder = m_Style == VS2012_Dark ? CBCGPDrawManager::ColorMakeLighter(m_clrMenuSeparator, .2) : m_clrMenuSeparator;

		m_clrControl = m_Style != VS2012_Dark ? 
			CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace, .028) :
			CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace, .1);

		m_clrDlgBackground = m_Style == VS2012_Dark ? CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace, .2) : globalData.clrBarFace;
		m_clrAutohideBarInactive = (m_Style != VS2012_Dark) ? globalData.clrBarShadow : globalData.clrBarLight;

		m_clrCaptionBarGradientDark = m_clrHighlight;
		m_clrCaptionBarGradientLight = m_clrHighlight;
	}

	m_clrToolBarGradientDark = m_clrToolBarGradientLight = globalData.clrBarFace;

	m_clrToolbarDisabled = globalData.clrBarHilite;
	m_clrSeparator = globalData.clrBarShadow;

	m_clrCombo = m_Style == VS2012_Dark ? CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace, .15) : globalData.clrBarLight;
	m_clrComboDisabled = globalData.clrBarFace;

	if (IsDarkTheme())
	{
		m_clrEditCtrlSelectionBkActive = m_clrAccentLight;
		m_clrEditCtrlSelectionTextActive = RGB(218, 218, 218);
		
		m_clrEditCtrlSelectionBkInactive = globalData.clrBarLight;
		m_clrEditCtrlSelectionTextInactive = globalData.clrBarText;
	}
	else
	{
		m_clrEditCtrlSelectionBkActive = CBCGPDrawManager::ColorMakePale(m_clrAccent, .87);
		m_clrEditCtrlSelectionBkInactive = CBCGPDrawManager::ColorMakePale(m_clrAccent, .95);

		m_clrEditCtrlSelectionTextActive = CBCGPDrawManager::IsDarkColor(m_clrEditCtrlSelectionBkActive) ? globalData.clrBarHilite : globalData.clrBarText;
		m_clrEditCtrlSelectionTextInactive = CBCGPDrawManager::IsDarkColor(m_clrEditCtrlSelectionBkInactive) ? globalData.clrBarHilite : globalData.clrBarText;
	}
	
	m_clrReportGroupText = m_clrInactiveCaptionGradient1;

	m_clrBarGradientDark = m_clrBarGradientLight = globalData.clrBarFace;
	
	m_clrPlannerWork      = globalData.clrBarFace;
	m_clrPalennerLine     = globalData.clrBarShadow;
	m_clrPlannerTodayFill = m_clrAccent;
	
	m_clrPlannerTodayLine = m_clrAccentText;

	m_clrRibbonTabs = globalData.clrBarText;
	m_clrTabsBackground = globalData.clrBarFace;

	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageBlack, RGB(113, 113, 113));

	m_clrMenuImage = m_Style == VS2012_Dark ? RGB(153, 153, 153) : RGB(113, 113, 113);
	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageGray, m_clrMenuImage);

	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageBlack2, m_clrAccent);
	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageWhite, RGB(255, 255, 255));

	m_clrHighlighDownText = globalData.clrBarText;
	
	if (m_Style != VS2012_Dark)
	{
		if (GetRValue (m_clrHighlightDn) <= 128 ||
			GetGValue (m_clrHighlightDn) <= 128 ||
			GetBValue (m_clrHighlightDn) <= 128)
		{
			m_clrHighlighDownText = globalData.clrBarHilite;
		}
	}

	m_clrRibbonCategoryFill = globalData.clrBarFace;

	m_clrNcTextActive   = globalData.clrBarText;
	m_clrNcTextInactive = globalData.clrBarDkShadow;
	m_clrNcBorder       = m_clrMenuBorder;

	m_clrButtonsArea = globalData.clrBarLight;
	m_clrBorderActive = m_clrAccent;
	
	if (!IsDarkTheme())
	{
		m_clrButtonsArea = CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace, 
			m_Style == VS2012_LightBlue ? .1 : .05);
	}

	m_clrFace = m_Style == VS2012_LightBlue ? m_clrButtonsArea : m_clrToolBarGradientLight;
	m_clrComboHighlighted = globalData.clrBarFace;

	m_clrHighlightDnNC = m_clrHighlightDn;

	if (m_Style == VS2012_LightBlue)
	{
		m_clrHighlightNC = CBCGPDrawManager::ColorMakePale(m_clrHighlight);
	}
	else
	{
		m_clrHighlightNC = m_clrHighlight;
	}
}
//************************************************************************************
void CBCGPVisualManagerVS2012::OnUpdateSystemColors ()
{
	CBCGPVisualManagerVS2010::OnUpdateSystemColors ();

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		m_brDlgBackground.DeleteObject();
		return;
	}

	SetupColors();

	m_dblGrayScaleLumRatio = (m_Style != VS2012_Dark) ? 0.8 : 1.1;

	m_brFrame.DeleteObject();
	if (m_clrFrame != (COLORREF)-1)
	{
		m_brFrame.CreateSolidBrush(m_clrFrame);
	}

	m_brAccent.DeleteObject();
	m_brAccent.CreateSolidBrush(m_clrAccent);

	m_brAccentLight.DeleteObject();
	m_brAccentLight.CreateSolidBrush(m_clrAccentLight);

	m_brAccentHighlight.DeleteObject();
	m_brAccentHighlight.CreateSolidBrush(m_clrAccentHilight);

	m_brBarBkgnd.DeleteObject();
	m_brBarBkgnd.CreateSolidBrush (m_clrCombo);

	m_brTabs.DeleteObject();

	if (m_Style == VS2012_LightBlue)
	{
		m_brTabs.CreateSolidBrush(RGB(41, 58, 86));
	}
	else
	{
		m_brTabs.CreateSolidBrush(m_clrAutohideBarInactive);
	}

	m_brHighlight.DeleteObject();
	m_brHighlight.CreateSolidBrush(m_clrHighlight);

	m_brHighlightDn.DeleteObject();
	m_brHighlightDn.CreateSolidBrush(m_clrHighlightDn);

	m_brHighlightChecked.DeleteObject();
	m_brHighlightChecked.CreateSolidBrush(m_clrHighlightChecked);

	m_brMenuConnect.DeleteObject();

	if (m_Style == VS2012_LightBlue)
	{
		m_brMenuConnect.CreateSolidBrush(m_clrMenuGutter);
	}
	else
	{
		m_brMenuConnect.CreateSolidBrush(m_clrMenuLight);
	}

	m_penMenuItemBorder.DeleteObject ();
	m_penMenuItemBorder.CreatePen (PS_SOLID, 1, m_clrMenuItemBorder);

	m_brMenuLight.DeleteObject ();
	m_brMenuLight.CreateSolidBrush(m_Style == VS2012_LightBlue ? m_clrMenuGutter : m_clrMenuLight);

	m_brTabBack.DeleteObject ();
	m_brTabBack.CreateSolidBrush (m_clrTabsBackground);

	m_brFace.DeleteObject ();
	m_brFace.CreateSolidBrush (m_clrFace);

	m_penBottomLine.DeleteObject ();
	m_penBottomLine.CreatePen (PS_SOLID, 1, m_clrToolBarBottomLine);

	m_brMenuButtonDroppedDown.DeleteObject ();
	m_brMenuButtonDroppedDown.CreateSolidBrush (m_clrBarBkgnd);

	m_brMenuItemCheckedHighlight.DeleteObject ();
	m_brMenuItemCheckedHighlight.CreateSolidBrush (globalData.clrBarLight);

	m_penSeparator.DeleteObject ();
	m_penSeparator.CreatePen (PS_SOLID, 1, m_clrSeparator);

	m_brMenuGutter.DeleteObject();
	m_brMenuGutter.CreateSolidBrush(m_clrMenuGutter);

	m_brDlgBackground.DeleteObject();
	m_brDlgBackground.CreateSolidBrush(m_clrDlgBackground);

	m_brGridSelectionBkActive.DeleteObject();
	m_brGridSelectionBkActive.CreateSolidBrush(m_clrEditCtrlSelectionBkActive);
	m_brGridSelectionBkInactive.DeleteObject();
	m_brGridSelectionBkInactive.CreateSolidBrush(m_clrEditCtrlSelectionBkInactive);

	m_brAutohideButton.DeleteObject();
	m_brAutohideButton.CreateSolidBrush(globalData.clrBarFace);

	m_brDlgButtonsArea.DeleteObject();
	m_brDlgButtonsArea.CreateSolidBrush(m_clrButtonsArea);

	m_penPlannerTodayLine.DeleteObject ();
	m_penPlannerTodayLine.CreatePen(PS_SOLID, 1, m_clrPlannerTodayLine);

	m_brGripperHorz.DeleteObject();
	m_brGripperVert.DeleteObject();

	WORD bits[4] = {0x70, 0xF0, 0xD0, 0xF0};
		
	CBitmap bm;
	bm.CreateBitmap(4, 4, 1, 1, bits);
	
	m_brGripperHorz.CreatePatternBrush(&bm);
	m_brGripperVert.CreatePatternBrush(&bm);

	m_brStatusBar.DeleteObject();
	m_brStatusBar.CreateSolidBrush(m_clrState == (COLORREF)-1 ? m_clrAccent : m_clrState);

	m_brRibbonCategoryFill.DeleteObject();
	m_brRibbonCategoryFill.CreateSolidBrush(m_clrRibbonCategoryFill);

	m_brControl.DeleteObject();
	m_brControl.CreateSolidBrush(m_clrControl);

	m_clrFrameShadowBaseActive = m_clrAccent;

	m_brHighlightNC.DeleteObject();
	m_brHighlightNC.CreateSolidBrush(m_clrHighlightNC);

	m_brHighlightDnNC.DeleteObject();
	m_brHighlightDnNC.CreateSolidBrush(m_clrHighlightDnNC);

	m_clrDockingBarScrollButton = globalData.clrBarLight;

	m_clrTaskPaneHotText = m_clrAccentText;
}
//*********************************************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillCommandsListBackground (CDC* pDC, CRect rect, BOOL bIsSelected)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnFillCommandsListBackground (pDC, rect, bIsSelected);
	}

	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	rect.left = 0;

	if (bIsSelected)
	{
		if (m_Style == VS2012_LightBlue)
		{
			pDC->FillSolidRect(rect, m_clrHighlightDn);
			pDC->Draw3dRect(rect, m_clrMenuItemBorder, m_clrMenuItemBorder);
			return globalData.clrBarText;
		}
		else
		{
			pDC->FillRect(rect, &m_brAccent);
			pDC->Draw3dRect(rect, RGB(255, 255, 255), RGB(255, 255, 255));
			return RGB(255, 255, 255);
		}
	}
	else
	{
		pDC->FillRect (rect, &m_brControl);
		return globalData.clrBarText;
	}
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::GetTabFrameColors (const CBCGPBaseTabWnd* pTabWnd,
				   COLORREF& clrDark,
				   COLORREF& clrBlack,
				   COLORREF& clrHighlight,
				   COLORREF& clrFace,
				   COLORREF& clrDarkShadow,
				   COLORREF& clrLight,
				   CBrush*& pbrFace,
				   CBrush*& pbrBlack)
{
	ASSERT_VALID (pTabWnd);

	if (m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::GetTabFrameColors (pTabWnd,
			clrDark, clrBlack,
			clrHighlight, clrFace,
			clrDarkShadow, clrLight,
			pbrFace, pbrBlack);

		return;
	}
	
	CBCGPVisualManagerXP::GetTabFrameColors (pTabWnd,
			   clrDark, clrBlack,
			   clrHighlight, clrFace,
			   clrDarkShadow, clrLight,
			   pbrFace, pbrBlack);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pTabWnd->IsDialogControl (TRUE) ||
		pTabWnd->IsFlatTab ())
	{
		return;
	}

	clrFace = clrBlack;
	pbrFace  = &m_brTabs;
	clrBlack = IsDarkTheme() ? globalData.clrBarShadow : globalData.clrBarDkShadow;
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTabResizeBar (CDC* pDC, CBCGPBaseTabWnd* pWndTab, 
									BOOL bIsVert, CRect rect,
									CBrush* pbrFace, CPen* pPen)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pWndTab->IsFlatTab ())
	{
		CBCGPVisualManagerVS2010::OnDrawTabResizeBar (pDC, pWndTab, bIsVert, rect, pbrFace, pPen);
		return;
	}

	ASSERT_VALID (pDC);
	pDC->FillRect(rect, &globalData.brBarFace);
}
//**********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTabDot(CDC* pDC, CBCGPBaseTabWnd* pWndTab, int nShape, const CRect& rectIn, int iTab, BOOL bIsActive, BOOL bIsHighlighted)
{
	ASSERT_VALID(pWndTab);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue ||
		pWndTab->IsDialogControl (TRUE))
	{
		CBCGPVisualManagerVS2010::OnDrawTabDot(pDC, pWndTab, nShape, rectIn, iTab, bIsActive, bIsHighlighted);
		return;
	}

	ASSERT_VALID(this);
	ASSERT_VALID(pDC);
	
	CBCGPDrawManager dm (*pDC);
	
	CRect rect = rectIn;

	COLORREF clrFill = (bIsHighlighted || bIsActive) ? m_clrAccent : (IsDarkTheme() && m_Style != VS2012_Dark ? globalData.clrBarLight : globalData.clrBarDkShadow);
	COLORREF clrBorder = (COLORREF)-1;
	
	if (rect.Width() >= 8)
	{
		COLORREF clrTab = pWndTab->GetTabBkColor(iTab);
		if (clrTab != (COLORREF)-1)
		{
			if (!bIsActive && !bIsHighlighted)
			{
				rect.DeflateRect(2, 2);
			}
			
			clrFill = clrTab;
		}
	}

	switch ((CBCGPTabWnd::TabDotShape)nShape)
	{
	case CBCGPTabWnd::TAB_DOT_ELLIPSE:
		dm.DrawEllipse(rect, clrFill, clrBorder);
		break;
		
	case CBCGPTabWnd::TAB_DOT_SQUARE:
		dm.DrawRect(rect, clrFill, clrBorder);
		break;
	}
}
//*********************************************************************************************************
BCGP_SMARTDOCK_THEME CBCGPVisualManagerVS2012::GetSmartDockingTheme ()
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || !globalData.IsWindowsLayerSupportAvailable ())
	{
		return CBCGPVisualManagerVS2010::GetSmartDockingTheme ();
	}

	return m_Style == VS2012_LightBlue ? BCGP_SDT_VS2010 : BCGP_SDT_VS2012;
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawSlider (CDC* pDC, CBCGPSlider* pSlider, CRect rect, BOOL bAutoHideMode)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnDrawSlider (pDC, pSlider, rect, bAutoHideMode);
		return;
	}

	ASSERT_VALID (pDC);

	DWORD dwAlgn = pSlider->GetCurrentAlignment ();
	COLORREF clrSlider = globalData.clrBarShadow;
	
	if (m_bThinSlider)
	{
		pDC->FillRect(rect, &globalData.brBarFace);

		CBCGPPenSelector ps(*pDC, m_clrEditBoxBorder);

		if ((dwAlgn & CBRS_ALIGN_LEFT) || (dwAlgn & CBRS_ALIGN_RIGHT))
		{
			const int x = (dwAlgn & CBRS_ALIGN_LEFT) ? (rect.right - 1) : rect.left;

			pDC->MoveTo(x, rect.top);
			pDC->LineTo(x, rect.bottom);
		}
		else
		{
			const int y = (dwAlgn & CBRS_ALIGN_TOP) ? (rect.bottom - 1) : rect.top;
			
			pDC->MoveTo(rect.left, y);
			pDC->LineTo(rect.right, y);
		}

		return;
	}
	
	if ((dwAlgn & CBRS_ALIGN_LEFT) || (dwAlgn & CBRS_ALIGN_RIGHT) || rect.Height() > rect.Width() * 4)
	{
		pDC->FillRect(rect, &globalData.brBarFace);
		
		rect.left = rect.CenterPoint().x - 1;
		rect.right = rect.CenterPoint().x + 1;
		
		CRgn rgn;
		rgn.CreateRectRgnIndirect (rect);
		
		pDC->SelectClipRgn (&rgn);
		
		CFrameWnd* pMainFrame = pSlider == NULL ? NULL : BCGCBProGetTopLevelFrame (pSlider);
		if (pMainFrame->GetSafeHwnd () != NULL)
		{
			CRect rectMain;
			
			pMainFrame->GetClientRect (rectMain);
			pMainFrame->MapWindowPoints (pSlider, &rectMain);
			
			rect.top = rectMain.top;
			rect.bottom = rect.top + globalData.m_rectVirtual.Height () + 10;
		}
		
		CBCGPDrawManager dm(*pDC);
		dm.Fill4ColorsGradient (rect, clrSlider, globalData.clrBarFace, globalData.clrBarFace, clrSlider);
		
		pDC->SelectClipRgn (NULL);
	}
	else
	{
		pDC->FillSolidRect(rect, clrSlider);
	}
}
//*********************************************************************************************************
int CBCGPVisualManagerVS2012::GetTabButtonState (CBCGPTabWnd* pTab, CBCGTabButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetTabButtonState (pTab, pButton);
	}

	if (m_Style == VS2012_LightBlue && (!pTab->IsFlatTab() || pButton->IsPressed ()))
	{
		return CBCGPVisualManagerVS2010::GetTabButtonState (pTab, pButton);
	}

	return (int) ((m_Style == VS2012_Dark || pButton->IsPressed ()) ? CBCGPMenuImages::ImageWhite : CBCGPMenuImages::ImageBlack);
}
//*****************************************************************************
int CBCGPVisualManagerVS2012::GetMenuDownArrowState (CBCGPToolbarMenuButton* pButton, BOOL bHightlight, BOOL bPressed, BOOL bDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetMenuDownArrowState (pButton, bHightlight, bPressed, bDisabled);
	}

	CBCGPWnd* pWnd = DYNAMIC_DOWNCAST(CBCGPWnd, pButton->GetParentWnd());

	if (pWnd != NULL && pWnd->IsOnGlass())
	{
		return CBCGPVisualManagerVS2010::GetMenuDownArrowState (pButton, bHightlight, bPressed, bDisabled);
	}

	if (bDisabled)
	{
		return (int)CBCGPMenuImages::ImageLtGray;
	}

	if (pButton->IsDroppedDown())
	{
		return (int)CBCGPMenuImages::ImageWhite;
	}

	return (int) ((m_Style != VS2012_Dark) ? CBCGPMenuImages::ImageBlack : CBCGPMenuImages::ImageWhite);
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetStatusBarPaneTextColor (CBCGPStatusBar* pStatusBar, CBCGStatusBarPaneInfo* pPane)
{
	ASSERT (pPane != NULL);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pPane->clrText != (COLORREF)-1)
	{
		return CBCGPVisualManagerVS2010::GetStatusBarPaneTextColor (pStatusBar, pPane);
	}

	COLORREF clrFill = m_clrState == (COLORREF)-1 ? m_clrAccent : m_clrState;
	BOOL bIsDisabled = (pPane->nStyle & SBPS_DISABLED);

	const double perceptiveLuminance = 0.299 * GetRValue(clrFill) + 0.587 * GetGValue(clrFill) + 0.114 * GetBValue(clrFill);
	if (perceptiveLuminance < 165.0)
	{
		if (bIsDisabled)
		{
			COLORREF clrText = CBCGPDrawManager::ColorMakeLighter(clrFill, .5);
			if (CBCGPDrawManager::IsPaleColor(clrText))
			{
				clrText = CBCGPDrawManager::ColorMakeLighter(clrFill, .2);
			}

			return clrText;
		}

		return RGB(255, 255, 255);
	}

	return bIsDisabled ? m_clrTextDisabled : RGB(0, 0, 0);
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawStatusBarPaneBorder (CDC* pDC, CBCGPStatusBar* pBar,
					CRect rectPane, UINT uiID, UINT nStyle)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawStatusBarPaneBorder (pDC, pBar, rectPane, 
								uiID, nStyle);
	}
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawComboDropButton (CDC* pDC, CRect rect,
												BOOL bDisabled,
												BOOL bIsDropped,
												BOOL bIsHighlighted,
												CBCGPToolbarComboBoxButton* pButton)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID (this);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || CBCGPToolBarImages::m_bIsDrawOnGlass ||
		m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnDrawComboDropButton (pDC, rect, bDisabled, bIsDropped, bIsHighlighted, pButton);
		return;
	}

	if (!bDisabled)
	{
		if (pButton != NULL && pButton->IsCtrlButton())
		{
			COLORREF clrFill = m_clrCombo;
			COLORREF clrBorder = (COLORREF)-1;
			
			if (bIsDropped)
			{
				pDC->FillRect(rect, &m_brHighlightDn);
				clrBorder = m_clrBorderActive;
				clrFill = m_clrHighlightDn;
			}
			else if (bIsHighlighted)
			{
				pDC->FillRect (rect, &m_brAccentLight);
				clrBorder = m_clrBorderActive;
				clrFill = m_clrAccentLight;
			}
			else
			{
				pDC->FillRect (rect, &globalData.brBarFace);
				clrFill = globalData.clrBarFace;
			}
			
			CBCGPMenuImages::DrawByColor(pDC, CBCGPMenuImages::IdArowDown, rect, clrFill);

			if (clrBorder != (COLORREF)-1)
			{
				pDC->Draw3dRect(rect, clrBorder, clrBorder);
			}
		}
		else
		{
			if (bIsDropped)
			{
				if (!IsDarkTheme())
				{
					rect.InflateRect(0, 1, 1, 1);
				}

				pDC->FillRect(rect, &m_brHighlightDn);
			}
			else if (bIsHighlighted)
			{
				if (!IsDarkTheme())
				{
					rect.InflateRect(0, 1, 1, 1);
				}

				pDC->FillSolidRect (rect, m_clrComboHighlighted);

				CPen pen (PS_SOLID, 1, m_clrAccent);
				CPen* pOldPen = pDC->SelectObject (&pen);

				pDC->MoveTo(rect.left, rect.top);
				pDC->LineTo(rect.left, rect.bottom);

				pDC->SelectObject (pOldPen);
			}
			else
			{
				pDC->FillSolidRect (rect, m_clrCombo);
			}

			CBCGPMenuImages::Draw(pDC, CBCGPMenuImages::IdArowDown, rect, 
				bIsDropped ? CBCGPMenuImages::ImageWhite : bIsHighlighted ? CBCGPMenuImages::ImageBlack2 : CBCGPMenuImages::ImageGray);
		}
	}
	else
	{
		pDC->FillRect (rect, &globalData.brBarFace);

		CBCGPMenuImages::Draw(pDC, CBCGPMenuImages::IdArowDown, rect, CBCGPMenuImages::ImageLtGray);
	}
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawComboBorder (CDC* pDC, CRect rect,
												BOOL bDisabled,
												BOOL bIsDropped,
												BOOL bIsHighlighted,
												CBCGPToolbarComboBoxButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || CBCGPToolBarImages::m_bIsDrawOnGlass ||
		m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnDrawComboBorder (pDC, rect, bDisabled, bIsDropped, bIsHighlighted, pButton);
		return;
	}

	COLORREF clrFrame = bDisabled ? (IsDarkTheme() ? globalData.clrBarLight : globalData.clrBarShadow) : m_clrEditBoxBorder;
	pDC->Draw3dRect(rect, clrFrame, clrFrame);
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawEditBorder (CDC* pDC, CRect rect,
												BOOL bDisabled,
												BOOL bIsHighlighted,
												CBCGPToolbarEditBoxButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || !CBCGPToolbarEditBoxButton::IsFlatMode () || CBCGPToolBarImages::m_bIsDrawOnGlass ||
		m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnDrawEditBorder (pDC, rect, bDisabled, bIsHighlighted, pButton);
		return;
	}
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnFillSpinButton (CDC* pDC, CBCGPSpinButtonCtrl* pSpinCtrl, CRect rect, BOOL bDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPVisualManagerVS2010::OnFillSpinButton(pDC, pSpinCtrl, rect, bDisabled);
		return;
	}

	COLORREF clrBorder = bDisabled ? (COLORREF)-1 : globalData.clrBarHilite;
	BOOL bIsDateTimeCtrl = pSpinCtrl != NULL && pSpinCtrl->m_bIsDateTimeControl;

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawRect (rect, m_clrControl, clrBorder);
	}
	else
	{
		pDC->FillRect(rect, (bDisabled && !bIsDateTimeCtrl) ? &globalData.brBarFace : &m_brControl);

		if (!bDisabled)
		{
			pDC->Draw3dRect(rect, clrBorder, clrBorder);
		}
	}
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawSpinButtons (CDC* pDC, CRect rectSpin, 
											  int nState, BOOL bOrientation, CBCGPSpinButtonCtrl* pSpinCtrl)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawSpinButtons (pDC, rectSpin, nState, bOrientation, pSpinCtrl);
		return;
	}

	ASSERT_VALID(pDC);
	ASSERT_VALID (this);
	
	CRect rect [2];
	rect[0] = rect[1] = rectSpin;
	
	if (!bOrientation) 
	{
		rect[0].DeflateRect(0, 0, 0, rect[0].Height() / 2);
		rect[1].top = rect[0].bottom ;
	}
	else
	{
		rect[1].DeflateRect(0, 0, rect[0].Width() / 2, 0);
		rect[0].left = rect[1].right;
	}
	
	CBCGPMenuImages::IMAGES_IDS id[2][2] = {{CBCGPMenuImages::IdArowUp, CBCGPMenuImages::IdArowDown}, {CBCGPMenuImages::IdArowRight, CBCGPMenuImages::IdArowLeft}};
	
	BOOL bDisabled = (nState & SPIN_DISABLED);
	int idxPressed = bDisabled ? -1 : (nState & (SPIN_PRESSEDUP | SPIN_PRESSEDDOWN)) - 1;
	
	int idxHighlighted = -1;
	if (!bDisabled)
	{
		if (nState & SPIN_HIGHLIGHTEDUP)
		{
			idxHighlighted = 0;
		}
		else if (nState & SPIN_HIGHLIGHTEDDOWN)
		{
			idxHighlighted = 1;
		}
	}
	
	for (int i = 0; i < 2; i ++)
	{
		CBCGPMenuImages::IMAGE_STATE state = CBCGPMenuImages::GetStateByColor(m_clrControl);

		if (pSpinCtrl != NULL && pSpinCtrl->m_bIsDateTimeControl)
		{
			COLORREF clrFill = m_clrCombo;
			COLORREF clrBorder = (COLORREF)-1;
			
			if (idxPressed == i)
			{
				pDC->FillRect(rect [i], &m_brHighlightDn);
				clrBorder = m_clrBorderActive;
				clrFill = m_clrHighlightDn;
			}
			else if (idxHighlighted == i)
			{
				pDC->FillRect (rect [i], m_Style == VS2012_LightBlue ? &m_brHighlight : &m_brAccentLight);
				clrBorder = m_clrBorderActive;
				clrFill = m_Style == VS2012_LightBlue ? m_clrHighlight : m_clrAccentLight;
			}
			else
			{
				pDC->FillRect (rect [i], &m_brControl);
				clrFill = m_clrControl;
			}

			state = CBCGPMenuImages::GetStateByColor(clrFill);
			
			if (clrBorder != (COLORREF)-1)
			{
				pDC->Draw3dRect(rect [i], clrBorder, clrBorder);
			}
		}
		else
		{
			if (idxPressed == i || idxHighlighted == i)
			{
				OnFillHighlightedArea (pDC, rect [i], 
					(idxPressed == i) ? &m_brHighlightDn : &m_brHighlight, NULL);
				
				state = CBCGPMenuImages::GetStateByColor((idxPressed == i) ? m_clrHighlightDn : m_clrHighlight);
			}
			else
			{
				OnFillSpinButton (pDC, pSpinCtrl, rect[i], bDisabled);
			}
		}
		
		CBCGPMenuImages::Draw (pDC, id[bOrientation ? 1 : 0][i], rect[i],
			bDisabled ? CBCGPMenuImages::ImageLtGray : state);
	}
	
	if (idxHighlighted >= 0)
	{
		CRect rectHot = rect [idxHighlighted];
		
		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect (rectHot, (COLORREF)-1, m_clrMenuItemBorder);
		}
		else
		{
			pDC->Draw3dRect (rectHot, m_clrMenuItemBorder, m_clrMenuItemBorder);
		}
	}
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnFillButtonInterior (CDC* pDC,
	CBCGPToolbarButton* pButton, CRect rect, CBCGPVisualManager::BCGBUTTON_STATE state)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnFillButtonInterior(pDC, pButton, rect, state);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->IsExplorerNavigationButton())
	{
		return;
	}

	if (m_Style == VS2012_Dark)
	{
		CBCGPBaseControlBar* pParentBar = DYNAMIC_DOWNCAST(CBCGPBaseControlBar, pButton->GetParentWnd());
		if (pParentBar != NULL && pParentBar->IsDialogControl ())
		{
			if (state != CBCGPVisualManager::ButtonsIsRegular)
			{
				if (m_hThemeToolBar == NULL)
				{
					COLORREF clrFill = state == CBCGPVisualManager::ButtonsIsHighlighted ? globalData.clrBtnLight : globalData.clrBtnFace;
					COLORREF clrLine = state == CBCGPVisualManager::ButtonsIsHighlighted ? globalData.clrBtnShadow : globalData.clrBtnDkShadow;
					
					CBCGPDrawManager dm (*pDC);
					dm.DrawRect (rect, clrFill, clrLine);
				}
				else
				{
					int nState = TS_NORMAL;
					
					if (pButton->m_nStyle & TBBS_DISABLED)
					{
						nState = TS_DISABLED;
					}
					else if ((pButton->m_nStyle & TBBS_PRESSED) && state == ButtonsIsHighlighted)
					{
						nState = TS_PRESSED;
					}
					else if (pButton->m_nStyle & TBBS_CHECKED)
					{
						nState = (state == ButtonsIsHighlighted) ? TS_HOTCHECKED : TS_CHECKED;
					}
					else if (state == ButtonsIsHighlighted)
					{
						nState = TS_HOT;
					}
					
					(*m_pfDrawThemeBackground) (m_hThemeToolBar, pDC->GetSafeHdc(), TP_BUTTON, nState, &rect, 0);
				}
			}
			return;
		}
	}
	
	if (state == ButtonsIsHighlighted && m_Style == VS2012_LightBlue && (pButton->m_nStyle & TBBS_CHECKED))
	{
		pDC->FillSolidRect(rect, CBCGPDrawManager::ColorMakeLighter(m_clrHighlight, .12));
		return;
	}

	if (DYNAMIC_DOWNCAST (CCustomizeButton, pButton) != NULL && state == ButtonsIsRegular && m_Style == VS2012_LightBlue)
	{
		pDC->FillSolidRect(rect, RGB(220, 224, 236));
	}

	CBCGPToolbarMenuButton* pMenuButton = DYNAMIC_DOWNCAST (CBCGPToolbarMenuButton, pButton);

	CWnd* pWndParentMenu = (pMenuButton != NULL) ? pMenuButton->GetParentWnd() : NULL;
	BOOL bIsPopupMenuButton = pWndParentMenu != NULL && pWndParentMenu->IsKindOf (RUNTIME_CLASS (CBCGPPopupMenuBar));

	if (pMenuButton != NULL && !bIsPopupMenuButton && pMenuButton->IsDroppedDown ())
	{
		CBCGPMenuBar* pParentMenuBar = DYNAMIC_DOWNCAST(CBCGPMenuBar, pWndParentMenu);
		if (pParentMenuBar != NULL)
		{
			CRect rectFill = rect;

			switch (pMenuButton->GetPopupMenu()->GetDropDirection ())
			{
			case CBCGPPopupMenu::DROP_DIRECTION_BOTTOM:
				rectFill.DeflateRect(1, 1, 1, 0);
				break;

			case CBCGPPopupMenu::DROP_DIRECTION_TOP:
				rectFill.DeflateRect(1, 0, 1, 1);
				break;

			case CBCGPPopupMenu::DROP_DIRECTION_RIGHT:
				rectFill.DeflateRect(1, 1, 0, 1);
				break;

			case CBCGPPopupMenu::DROP_DIRECTION_LEFT:
				rectFill.DeflateRect(0, 1, 1, 1);
				break;

			default:
				if (pParentMenuBar->IsHorizontal())
				{
					rectFill.DeflateRect(1, 0, 1, 1);
				}
				else
				{
					rectFill.DeflateRect(1, 1, 1, 1);
				}
				break;
			}

			if (CBCGPToolBarImages::m_bIsDrawOnGlass)
			{
				CBCGPDrawManager dm (*pDC);

				dm.DrawRect (rect, (COLORREF)-1, m_clrMenuGutter);
				dm.DrawRect (rectFill, m_clrMenuLight, (COLORREF)-1);
			}
			else
			{
				pDC->Draw3dRect(rect, m_clrMenuBorder, m_clrMenuBorder);
				pDC->FillRect (rectFill, &m_brMenuLight);
			}
		}
		else
		{
			rect.DeflateRect(1, 0);

			if (CBCGPToolBarImages::m_bIsDrawOnGlass)
			{
				CBCGPDrawManager dm (*pDC);
				dm.DrawRect (rect, m_clrHighlightDn, (COLORREF)-1);
			}
			else
			{
				pDC->FillRect (rect, &m_brHighlightDn);
			}
		}

		return;
	}

	CBCGPVisualManagerXP::OnFillButtonInterior (pDC, pButton, rect, state);
}
//***********************************************************************************
void CBCGPVisualManagerVS2012::OnFillHighlightedArea (CDC* pDC, CRect rect, CBrush* pBrush, CBCGPToolbarButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pButton == NULL || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnFillHighlightedArea (pDC, rect, pBrush, pButton);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pBrush);
	ASSERT_VALID (pButton);

	CBCGPDrawManager dm (*pDC);

	if (pBrush == &m_brHighlight)
	{
		BOOL bIsPopupMenu = pButton->GetParentWnd () != NULL && pButton->GetParentWnd ()->IsKindOf (RUNTIME_CLASS (CBCGPPopupMenuBar));

		dm.DrawRect(rect, bIsPopupMenu ? m_clrHighlightMenuItem : m_clrHighlight, (COLORREF)-1);
	}
	else if (pBrush == &m_brHighlightDn)
	{
		dm.DrawRect(rect, (pButton->m_nStyle & TBBS_CHECKED) == 0 ? m_clrHighlightDn : m_clrHighlight, (COLORREF)-1);
	}
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2012::OnHighlightRarelyUsedMenuItems (CDC* pDC, CRect rectRarelyUsed)
{
	ASSERT_VALID (pDC);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnHighlightRarelyUsedMenuItems (pDC, rectRarelyUsed);
		return;
	}

	rectRarelyUsed.left --;
	rectRarelyUsed.right = rectRarelyUsed.left + CBCGPToolBar::GetMenuImageSize ().cx + 
		2 * GetMenuImageMargin () + 2;

	pDC->FillRect (rectRarelyUsed, &m_brHighlight);
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawButtonBorder (CDC* pDC, CBCGPToolbarButton* pButton, CRect rect, BCGBUTTON_STATE state)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawButtonBorder (pDC, pButton, rect, state);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->IsExplorerNavigationButton())
	{
		return;
	}

	if (m_Style == VS2012_Dark)
	{
		CBCGPBaseControlBar* pParentBar = DYNAMIC_DOWNCAST(CBCGPBaseControlBar, pButton->GetParentWnd());
		if (pParentBar != NULL && pParentBar->IsDialogControl ())
		{
			return;
		}
	}

	COLORREF clrBorder = (COLORREF)-1;

	if (m_Style == VS2012_LightBlue && state != CBCGPVisualManager::ButtonsIsRegular && !pButton->IsDroppedDown())
	{
		if (DYNAMIC_DOWNCAST (CCustomizeButton, pButton) == NULL)
		{
			clrBorder = m_clrMenuItemBorder;
		}
	}
	else if (pButton->m_nStyle & TBBS_CHECKED)
	{
		clrBorder = (!CBCGPToolBar::IsCustomizeMode () && (pButton->m_nStyle & TBBS_DISABLED)) ? globalData.clrBarShadow : 
					(m_Style == VS2012_LightBlue ? m_clrMenuItemBorder : m_clrAccentLight);
	}

	if (clrBorder != (COLORREF)-1)
	{
		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			
			dm.DrawRect (rect, (COLORREF)-1, clrBorder);
		}
		else
		{
			pDC->Draw3dRect(rect, clrBorder, clrBorder);
		}
	}
}
//*************************************************************************************
int CBCGPVisualManagerVS2012::GetTabExtraHeight(const CBCGPTabWnd* pTabWnd)
{
	ASSERT_VALID(pTabWnd);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () ||
		(pTabWnd->IsDialogControl (TRUE) && !pTabWnd->IsPropertySheetTab()) || !pTabWnd->IsMDITab() ||
		pTabWnd->IsFlatTab () || pTabWnd->IsOneNoteStyle () || pTabWnd->IsVS2005Style ())
	{
		return CBCGPVisualManagerVS2010::GetTabExtraHeight(pTabWnd);
	}

	return 0;
}
//*************************************************************************************
BOOL CBCGPVisualManagerVS2012::OnEraseMDIClientArea (CDC* pDC, CRect rectClient)
{
	if (globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode () && m_Style != VS2012_LightBlue)
	{
		CMDIFrameWnd* pMDIFrameWnd = DYNAMIC_DOWNCAST (CMDIFrameWnd, AfxGetMainWnd ());
		if (pMDIFrameWnd != NULL)
		{
			pDC->FillRect(rectClient, &globalData.brBarFace);
			return TRUE;
		}
	}

	return CBCGPVisualManagerVS2010::OnEraseMDIClientArea (pDC, rectClient);
}
//*************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPrintPreviewBackgroundColor(CBCGPPrintPreviewView* pPrintPreview)
{
	return CBCGPVisualManagerXP::GetPrintPreviewBackgroundColor(pPrintPreview);
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPrintPreviewFrameColor(BOOL bIsActive)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::GetPrintPreviewFrameColor(bIsActive);
	}

	return bIsActive ? m_clrAccent : globalData.clrBarDkShadow;
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnFillCombo (CDC* pDC, CRect rect,
										BOOL bDisabled,
										BOOL bIsDropped,
										BOOL bIsHighlighted,
										CBCGPToolbarComboBoxButton* pButton)
{
	ASSERT_VALID (pDC);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPVisualManagerXP::OnFillCombo(pDC, rect, bDisabled, bIsDropped, bIsHighlighted, pButton);
		return;
	}

	COLORREF clr = bDisabled ? m_clrComboDisabled : (bIsHighlighted || bIsDropped) && m_Style != VS2012_LightBlue ? globalData.clrBarHilite : m_clrCombo;

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);

		dm.DrawRect (rect, clr, (COLORREF)-1);
	}
	else
	{
		pDC->FillSolidRect (rect, clr);
	}
}
//************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetComboboxTextColor(CBCGPToolbarComboBoxButton* pButton, 
										BOOL bDisabled,
										BOOL bIsDropped,
										BOOL bIsHighlighted)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetComboboxTextColor(pButton, bDisabled, bIsDropped, bIsHighlighted);
	}

	return bDisabled ? globalData.clrGrayedText : globalData.clrBarText;
}
//************************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillMiniFrameCaption (CDC* pDC, 
								CRect rectCaption, 
								CBCGPMiniFrameWnd* pFrameWnd,
								BOOL bActive)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pFrameWnd);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::OnFillMiniFrameCaption (pDC, rectCaption, pFrameWnd, bActive);
	}

	if (DYNAMIC_DOWNCAST (CBCGPBaseToolBar, pFrameWnd->GetControlBar ()) != NULL)
	{
		bActive = FALSE;
	}

	if (bActive)
	{
		pDC->FillRect(rectCaption, &m_brAccent);
		return RGB(255, 255, 255);
	}
	else
	{
		pDC->FillRect(rectCaption, &globalData.brBarFace);
		return globalData.clrBarText;
	}
}
//************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawMiniFrameBorder (
										CDC* pDC, CBCGPMiniFrameWnd* pFrameWnd,
										CRect rectBorder, CRect rectBorderSize)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnDrawMiniFrameBorder (pDC, pFrameWnd, rectBorder, rectBorderSize);
		return;
	}

	ASSERT_VALID (pDC);

	CRect rectClip = rectBorder;
	rectClip.DeflateRect(rectClip);

	pDC->ExcludeClipRect(rectClip);

	if (pFrameWnd->GetSafeHwnd() != NULL && pFrameWnd->IsActive())
	{
		pDC->FillRect(rectBorder, &m_brAccent);
	}
	else
	{
		pDC->FillRect(rectBorder, &globalData.brBarFace);
		pDC->Draw3dRect(rectBorder, globalData.clrBarDkShadow, globalData.clrBarDkShadow);
	}

	pDC->SelectClipRgn (NULL);
}
//**************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetToolbarButtonTextColor (CBCGPToolbarButton* pButton, CBCGPVisualManager::BCGBUTTON_STATE state)
{
	ASSERT_VALID (pButton);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pButton->IsKindOf (RUNTIME_CLASS (CBCGPOutlookButton)) ||
		(m_Style == VS2012_LightBlue && !pButton->IsExplorerNavigationButton()))
	{
		return CBCGPVisualManagerVS2010::GetToolbarButtonTextColor (pButton, state);
	}

	if (pButton->GetParentWnd() != NULL && state == ButtonsIsRegular)
	{
		if (DYNAMIC_DOWNCAST (CBCGPReBar, pButton->GetParentWnd()) != NULL || DYNAMIC_DOWNCAST (CBCGPReBar, pButton->GetParentWnd()->GetParent ()))
		{
			return CBCGPVisualManagerVS2010::GetToolbarButtonTextColor (pButton, state);
		}
	}

	BOOL bDisabled = (CBCGPToolBar::IsCustomizeMode () && !pButton->IsEditable ()) ||
			(!CBCGPToolBar::IsCustomizeMode () && (pButton->m_nStyle & TBBS_DISABLED));

	CBCGPControlBar* pWnd = DYNAMIC_DOWNCAST(CBCGPControlBar, pButton->GetParentWnd());
	if (pWnd != NULL && pWnd->IsDialogControl())
	{
		return bDisabled ? globalData.clrGrayedText : globalData.clrBtnText;
	}

	COLORREF clrLight = m_Style != VS2012_Dark ? globalData.clrBarHilite : globalData.clrBarText;
	BOOL bIsDroppedDownOnToolbar = pButton->IsDroppedDown();

	if (bIsDroppedDownOnToolbar)
	{
		if (pButton->GetParentWnd() != NULL && pButton->GetParentWnd()->IsKindOf (RUNTIME_CLASS (CBCGPMenuBar)))
		{
			bIsDroppedDownOnToolbar = FALSE;
		}
	}

	if (pButton->IsExplorerNavigationButton())
	{
		COLORREF clrHot = m_Style == VS2012_Dark ? m_clrAccentText : m_clrAccentHilight;
		COLORREF clrPressed = m_Style == VS2012_Dark ? m_clrAccentLight : m_clrAccentText;

		return bDisabled ? m_clrTextDisabled : state == CBCGPVisualManager::ButtonsIsPressed ? clrPressed : state == CBCGPVisualManager::ButtonsIsHighlighted ? clrHot : globalData.clrBarText;
	}

	return bDisabled ? m_clrTextDisabled : bIsDroppedDownOnToolbar ? clrLight : 
		(pButton->m_nStyle & TBBS_PRESSED) ? ((pButton->m_nStyle & TBBS_CHECKED) ? m_clrAccentText : clrLight) : globalData.clrBarText;
}
//**************************************************************************************
BOOL CBCGPVisualManagerVS2012::OnDrawPushButton (CDC* pDC, CRect rect, CBCGPButton* pButton, COLORREF& clrText)
{
	ASSERT_VALID(pDC);

	BOOL bOnGlass = pButton->GetSafeHwnd() != NULL && pButton->m_bOnGlass;

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || bOnGlass ||
		m_brDlgBackground.GetSafeHandle () == NULL)
	{
		return CBCGPVisualManagerVS2010::OnDrawPushButton (pDC, rect, pButton, clrText);
	}

	ASSERT_VALID (pDC);

	clrText = globalData.clrBarText;

	BOOL bDefault = pButton->GetSafeHwnd() != NULL && pButton->IsDefaultButton () && pButton->IsWindowEnabled ();

	CPen pen (PS_SOLID, 1, bDefault ? m_clrAccentText : (m_Style == VS2012_Dark ? globalData.clrBarShadow : globalData.clrBarDkShadow));
	CPen penAccent (PS_SOLID, 1, m_clrAccent);
	CPen* pOldPen = pDC->SelectObject (&pen);

	CBrush* pOldBrush = (CBrush*)pDC->SelectObject(m_Style == VS2012_LightBlue ? &globalData.brBarFace : &m_brControl);

	if (pButton->GetSafeHwnd() != NULL && pButton->IsWindowEnabled ())
	{
		BOOL bIsFocused = CWnd::GetFocus () == pButton;

		if (pButton->GetCheck() && !bIsFocused && !pButton->IsPressed () && !pButton->IsHighlighted())
		{
			pDC->SelectObject(&m_brHighlightChecked);
			pDC->SelectObject(&penAccent);
		}
		else if (pButton->IsPressed ())
		{
			if (!pButton->GetCheck ())
			{
				pDC->SelectObject(&m_brHighlightDn);
				clrText = m_clrHighlighDownText;
			}
		}
		else if (pButton->IsHighlighted () || bIsFocused)
		{
			pDC->SelectObject(&m_brHighlight);
		}
	}

	pDC->Rectangle(rect);

	pDC->SelectObject(pOldPen);
	pDC->SelectObject(pOldBrush);

	return TRUE;
}
//**************************************************************************************
int CBCGPVisualManagerVS2012::GetPushButtonStdImageState(CBCGPButton* pButton, BOOL bIsDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || !bIsDisabled)
	{
		return CBCGPVisualManagerVS2010::GetPushButtonStdImageState(pButton, bIsDisabled);
	}

	return (int)(IsDarkTheme() ? CBCGPMenuImages::ImageDkGray : CBCGPMenuImages::ImageLtGray);
}
//**************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetToolbarDisabledTextColor ()
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::GetToolbarDisabledTextColor();
	}

	return m_clrTextDisabled;
}

#ifndef BCGP_EXCLUDE_PROP_LIST

COLORREF CBCGPVisualManagerVS2012::OnFillPropList(CDC* pDC, CBCGPPropList* pPropList, const CRect& rectClient, COLORREF& clrFill)
{
	ASSERT_VALID (pPropList);

	if (!pPropList->DrawControlBarColors() || globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::OnFillPropList(pDC, pPropList, rectClient, clrFill);
	}

	if (m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::OnFillPropList(pDC, pPropList, rectClient, clrFill);
	}

	pDC->FillRect(rectClient, &m_brControl);
	clrFill = m_clrControl;

	return globalData.clrBarText;
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnFillPropListToolbarArea(CDC* pDC, CBCGPPropList* pPropList, const CRect& rectToolBar)
{
	if ((pPropList != NULL && !pPropList->DrawControlBarColors()) || globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnFillPropListToolbarArea(pDC, pPropList, rectToolBar);
		return;
	}

	pDC->FillRect(rectToolBar, &globalData.brBarFace);
}
//****************************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillPropertyListSelectedItem(CDC* pDC, CBCGPProp* pProp, CBCGPPropList* pWndList, const CRect& rectFill, BOOL bFocused)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID(pProp);

	if (!pWndList->DrawControlBarColors() || globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::OnFillPropertyListSelectedItem(pDC, pProp, pWndList, rectFill, bFocused);
	}

	COLORREF clrText = (COLORREF)-1;

	if (bFocused)
	{
		if (m_Style == VS2012_LightBlue)
		{
			pDC->FillRect (rectFill, &m_brHighlight);
			
			if (pProp->IsEnabled())
			{
				clrText = globalData.clrBarText;
			}
			else
			{
				clrText = m_clrTextDisabled;
			}
		}
		else
		{
			pDC->FillRect (rectFill, &m_brAccentLight);

			BOOL bIsDarkBackground = (GetRValue (m_clrAccentLight) <= 128 ||
				GetGValue (m_clrAccentLight) <= 128 ||
				GetBValue (m_clrAccentLight) <= 128);

			if (pProp->IsEnabled())
			{
				if (bIsDarkBackground)
				{
					clrText = RGB(241, 241, 241);
				}
				else
				{
					clrText = RGB(30, 30, 30);
				}
			}
			else
			{
				if (bIsDarkBackground)
				{
					clrText = RGB(192, 192, 192);
				}
				else
				{
					clrText = RGB(82, 82, 82);
				}
			}
		}
	}
	else
	{
		pDC->FillSolidRect(rectFill, IsDarkTheme() ? globalData.clrBarLight : globalData.clrBarShadow);
		clrText = pProp->IsEnabled() ? globalData.clrBarText : globalData.clrBarDkShadow;
	}

	return clrText;
}
//****************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPropListGroupTextColor(CBCGPPropList* pPropList)
{
	ASSERT_VALID (pPropList);

	if (!pPropList->DrawControlBarColors () || globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::GetPropListGroupTextColor(pPropList);
	}

	return m_Style != VS2012_Dark ? RGB(30, 30, 30) : globalData.clrBarText;
}
//***********************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillPropListDescriptionArea(CDC* pDC, CBCGPPropList* pList, const CRect& rect)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pList);

	if (!pList->DrawControlBarColors () || globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::OnFillPropListDescriptionArea(pDC, pList, rect);
	}

	pDC->FillRect(rect, &m_brDlgBackground);
	return IsDarkTheme() ? m_clrEditBoxBorder : m_clrMenuBorder;
}
//***********************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillPropListCommandsArea(CDC* pDC, CBCGPPropList* pList, const CRect& rect)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pList);

	if (!pList->DrawControlBarColors () || globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::OnFillPropListCommandsArea(pDC, pList, rect);
	}

	pDC->FillRect(rect, &m_brDlgBackground);
	return IsDarkTheme() ? m_clrEditBoxBorder : m_clrMenuBorder;
}
//***********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPropListCommandTextColor (CBCGPPropList* pPropList)
{	
	ASSERT_VALID (pPropList);

	if (!pPropList->DrawControlBarColors () || globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetPropListCommandTextColor (pPropList);
	}

	return m_clrAccentText;
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPropListDisabledTextColor(CBCGPPropList* pPropList)
{
	if (!pPropList->DrawControlBarColors () || globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetPropListDisabledTextColor(pPropList);
	}

	return globalData.clrBarDkShadow;
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPropListGroupColor(CBCGPPropList* pPropList)
{
	ASSERT_VALID (pPropList);
	
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () ||
		!pPropList->DrawControlBarColors ())
	{
		return CBCGPVisualManagerVS2010::GetPropListGroupColor (pPropList);
	}
	
	return m_clrDlgBackground;
}

#endif

void CBCGPVisualManagerVS2012::OnDrawMenuBorder (CDC* pDC, CBCGPPopupMenu* pMenu, CRect rect)
{
	BOOL bConnectMenuToParent = m_bConnectMenuToParent;

	if (bConnectMenuToParent && pMenu != NULL)
	{
		CBCGPToolBar* pParentToolBar = pMenu->GetParentToolBar();

		if (DYNAMIC_DOWNCAST(CBCGPMenuBar, pParentToolBar) == NULL &&
			DYNAMIC_DOWNCAST(CBCGPPopupMenuBar, pParentToolBar) == NULL)
		{
			m_bConnectMenuToParent = FALSE;
		}
	}

	CBCGPVisualManagerXP::OnDrawMenuBorder (pDC, pMenu, rect);

	m_bConnectMenuToParent = bConnectMenuToParent;
}
//***********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawMenuScrollButton (CDC* pDC, CRect rect, BOOL bIsScrollDown, 
												 BOOL bIsHighlited, BOOL bIsPressed,
												 BOOL bIsDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawMenuScrollButton (pDC, rect, bIsScrollDown, bIsHighlited, bIsPressed, bIsDisabled);
		return;
	}

	ASSERT_VALID (pDC);

	rect.top --;

	pDC->FillRect(rect, bIsHighlited ? &m_brHighlight : &globalData.brBarFace);

	CBCGPMenuImages::DrawByColor(pDC, bIsScrollDown ? CBCGPMenuImages::IdArowDown : CBCGPMenuImages::IdArowUp, rect,
		bIsHighlited ? m_clrHighlight : globalData.clrBarFace);
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawMenuCheck (CDC* pDC, CBCGPToolbarMenuButton* pButton, CRect rect, BOOL bHighlight, BOOL bIsRadio)
{
	ASSERT_VALID (pButton);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawMenuCheck (pDC, pButton, rect, bHighlight, bIsRadio);
		return;
	}

	int iImage = bIsRadio ? CBCGPMenuImages::IdRadio : CBCGPMenuImages::IdCheck;
	rect.DeflateRect(0, 1);

	CBCGPMenuImages::IMAGE_STATE imageState = CBCGPMenuImages::ImageGray;

	if (pButton->m_nStyle & TBBS_DISABLED)
	{
		imageState = CBCGPMenuImages::ImageLtGray;
	}
	else if (bHighlight)
	{
		imageState = m_Style == VS2012_Dark ? CBCGPMenuImages::ImageLtGray : CBCGPMenuImages::ImageDkGray;
	}

	CBCGPMenuImages::Draw (pDC, (CBCGPMenuImages::IMAGES_IDS) iImage, rect, imageState);
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnFillMenuImageRect (CDC* pDC, CBCGPToolbarButton* pButton, CRect rect, CBCGPVisualManager::BCGBUTTON_STATE state)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnFillMenuImageRect(pDC, pButton, rect, state);
		return;
	}

	if ((pButton->m_nStyle & TBBS_CHECKED) && m_Style == VS2012_LightBlue)
	{
		ASSERT_VALID (pDC);
		ASSERT_VALID (pButton);

		if (pButton->GetImage() >= 0 && DYNAMIC_DOWNCAST(CBCGPCustomizeMenuButton, pButton) == NULL)
		{
			OnFillButtonInterior (pDC, pButton, rect, state);
		}
	}
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawMenuImageRectBorder (CDC* pDC, CBCGPToolbarButton* pButton, CRect rect, CBCGPVisualManager::BCGBUTTON_STATE state)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawMenuImageRectBorder(pDC, pButton, rect, state);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->m_nStyle & TBBS_CHECKED)
	{
		if (!CBCGPToolBar::IsCustomizeMode () && (pButton->m_nStyle & TBBS_DISABLED))
		{
			pDC->Draw3dRect(rect, globalData.clrBarShadow, globalData.clrBarShadow);
		}
		else if (pButton->GetImage() < 0 || DYNAMIC_DOWNCAST(CBCGPCustomizeMenuButton, pButton) != NULL)
		{
			COLORREF clrFrame = m_Style != VS2012_Dark ? RGB(113, 113, 113) : RGB(153, 153, 153);

			if (pButton->m_nStyle & TBBS_MARKED)
			{
				clrFrame = m_Style != VS2012_Dark ? RGB(30, 30, 30) : RGB(241, 241, 241);
			}

			rect.bottom++;

			pDC->Draw3dRect(rect, clrFrame, clrFrame);
			rect.DeflateRect(1, 1);
			pDC->Draw3dRect(rect, clrFrame, clrFrame);
		}
		else
		{
			COLORREF clrFrame = m_Style == VS2012_LightBlue ? m_clrMenuItemBorder : m_clrHighlightDn;
			pDC->Draw3dRect(rect, clrFrame, clrFrame);
		}
	}
}
//**************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetMenuItemTextColor(CBCGPToolbarMenuButton* pButton, BOOL bHighlighted, BOOL bDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetMenuItemTextColor (pButton, bHighlighted, bDisabled);
	}

	return bDisabled ? m_clrTextDisabled : globalData.clrBarText;
}
//**************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawCaptionButton (CDC* pDC, CBCGPCaptionButton* pButton, 
								BOOL bActive, BOOL bHorz, BOOL bMaximized, BOOL bDisabled,
								int nImageID /*= -1*/)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnDrawCaptionButton (pDC, pButton, bActive, bHorz, bMaximized, bDisabled, nImageID);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

    CRect rc = pButton->GetRect ();

	const BOOL bHighlight = (pButton->m_bFocused || pButton->m_bDroppedDown) && !bDisabled;

	if (pButton->m_bPushed && bHighlight)
	{
		if (bActive)
		{
			pDC->FillSolidRect (rc, m_Style == VS2012_Dark ? m_clrAccentLight : m_clrAccentText);
		}
		else
		{
			pDC->FillRect(rc, &m_brAccent);
		}
	}
	else if (bHighlight || pButton->m_bPushed)
	{
		if (bActive)
		{
			pDC->FillSolidRect (rc, m_Style == VS2012_Dark ? m_clrAccentText : m_clrAccentLight);
		}
		else
		{
			pDC->FillRect (rc, &m_brHighlight);
		}
	}

	CBCGPMenuImages::IMAGES_IDS id = (CBCGPMenuImages::IMAGES_IDS)-1;
	
	if (nImageID != -1)
	{
		id = (CBCGPMenuImages::IMAGES_IDS)nImageID;
	}
	else
	{
		id = pButton->GetIconID (bHorz, bMaximized);
	}

	if (id != (CBCGPMenuImages::IMAGES_IDS)-1)
	{
		CSize sizeImage = CBCGPMenuImages::Size ();
		CPoint ptImage (rc.left + (rc.Width () - sizeImage.cx) / 2,
						rc.top + (rc.Height () - sizeImage.cy) / 2);

		OnDrawCaptionButtonIcon (pDC, pButton, id, bActive, bDisabled, ptImage);
	}
}
//**********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawCaptionButtonIcon (CDC* pDC, 
													CBCGPCaptionButton* pButton,
													CBCGPMenuImages::IMAGES_IDS id,
													BOOL bActive, BOOL bDisabled,
													CPoint ptImage)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnDrawCaptionButtonIcon (pDC, pButton, id, bActive, bDisabled, ptImage);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	CBCGPMenuImages::IMAGE_STATE imageState = (m_Style == VS2012_Dark ? CBCGPMenuImages::ImageLtGray : CBCGPMenuImages::ImageBlack);
	
	if (bDisabled)
	{
		imageState = CBCGPMenuImages::ImageGray;
	}
	else if (bActive || ((pButton->m_bFocused && pButton->m_bPushed) && m_Style != VS2012_Dark))
	{
		imageState = CBCGPMenuImages::ImageWhite;
	}

	CBCGPMenuImages::Draw (pDC, id, ptImage, imageState);
}

#ifndef BCGP_EXCLUDE_TOOLBOX

BOOL CBCGPVisualManagerVS2012::OnEraseToolBoxButton (CDC* pDC, CRect rect,
											CBCGPToolBoxButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::OnEraseToolBoxButton(pDC, rect, pButton);
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->GetCheck ())
	{
		pDC->FillRect(rect, &m_brAccentLight);
	}
	else if (pButton->IsHighlighted ())
	{
		pDC->FillRect(rect, &m_brHighlight);
	}

	return TRUE;
}
//**********************************************************************************
BOOL CBCGPVisualManagerVS2012::OnDrawToolBoxButtonBorder (CDC* pDC, CRect& rect, 
												 CBCGPToolBoxButton* pButton, UINT uiState)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnDrawToolBoxButtonBorder(pDC, rect, pButton, uiState);
	}

	if (pButton->GetCheck () && m_Style == VS2012_LightBlue)
	{
		pDC->Draw3dRect(rect, m_clrMenuItemBorder, m_clrMenuItemBorder);
	}

	return TRUE;
}
//**********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetToolBoxButtonTextColor (CBCGPToolBoxButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetToolBoxButtonTextColor(pButton);
	}

	if (m_Style == VS2012_Dark || m_Style == VS2012_LightBlue)
	{
		return globalData.clrBarText;
	}

	ASSERT_VALID (pButton);
	return pButton->m_bIsChecked ? globalData.clrBarHilite : globalData.clrBarText;
}

#endif	// BCGP_EXCLUDE_TOOLBOX

#ifndef BCGP_EXCLUDE_TASK_PANE

void CBCGPVisualManagerVS2012::OnFillTasksPaneBackground(CDC* pDC, CRect rectWorkArea)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnFillTasksPaneBackground(pDC, rectWorkArea);
		return;
	}

	ASSERT_VALID(pDC);

	pDC->FillRect(rectWorkArea, &m_brMenuLight);
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTasksGroupCaption(CDC* pDC, CBCGPTasksGroup* pGroup, 
								BOOL bIsHighlighted, BOOL bIsSelected, 
								BOOL bCanCollapse)
{
#ifndef BCGP_EXCLUDE_TOOLBOX
	BOOL bIsToolBox = pGroup->m_pPage->m_pTaskPane != NULL &&
		(pGroup->m_pPage->m_pTaskPane->IsKindOf (RUNTIME_CLASS (CBCGPToolBoxEx)));
#else
	BOOL bIsToolBox = FALSE;
#endif

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnDrawTasksGroupCaption(
										pDC, pGroup, 
										bIsHighlighted, bIsSelected, bCanCollapse);
		return;
	}

	CRect rectGroup = pGroup->m_rect;

	// ------------------
	// Draw group caption
	// ------------------
	CRect rectFill = rectGroup;
	rectFill.DeflateRect (1, 1);

	pDC->FillRect(rectFill, m_Style == VS2012_LightBlue ? &m_brDlgBackground : &m_brControl);

	int nIconHOffset = 0;

	if (bCanCollapse)
	{
		//--------------------
		// Draw expanding box:
		//--------------------
		int nBoxSize = globalUtils.ScaleByDPI(9);
		int nBoxOffset = 6;

		CRect rectButton = rectFill;
		
		rectButton.left += nBoxOffset;
		rectButton.right = rectButton.left + nBoxSize;
		rectButton.top = rectButton.CenterPoint ().y - nBoxSize / 2;
		rectButton.bottom = rectButton.top + nBoxSize;

		nIconHOffset += nBoxOffset + nBoxSize;

		OnDrawExpandingBox (pDC, rectButton, !pGroup->m_bIsCollapsed, globalData.clrBarText);

		rectGroup.left = rectButton.right + nBoxOffset;
	}

	// ---------------------------
	// Draw an icon if it presents
	// ---------------------------
	BOOL bShowIcon = (!bIsToolBox && pGroup->m_hIcon != NULL 
		&& pGroup->m_sizeIcon.cx < rectGroup.Width () - rectGroup.Height());
	if (bShowIcon)
	{
		OnDrawTasksGroupIcon(pDC, pGroup, 5 + nIconHOffset, bIsHighlighted, bIsSelected, bCanCollapse);
	}

	// -----------------------
	// Draw group caption text
	// -----------------------
	CFont* pFontOld = pDC->SelectObject (&globalData.fontBold);
	COLORREF clrTextOld = pDC->GetTextColor ();
	int nBkModeOld = pDC->SetBkMode(TRANSPARENT);

	if (bIsToolBox)
	{
		pDC->SetTextColor (globalData.clrBarText);
	}
	else
	{
		if (bCanCollapse && bIsHighlighted)
		{
			pDC->SetTextColor (pGroup->m_clrTextHot == (COLORREF)-1 ?
				globalData.clrBarText : pGroup->m_clrTextHot);
		}
		else
		{
			pDC->SetTextColor (pGroup->m_clrText == (COLORREF)-1 ?
				globalData.clrBarText : pGroup->m_clrText);
		}
	}
	
	int nTaskPaneHOffset = pGroup->m_pPage->m_pTaskPane->GetGroupCaptionHorzOffset();
	int nTaskPaneVOffset = pGroup->m_pPage->m_pTaskPane->GetGroupCaptionVertOffset();
	int nCaptionHOffset = (nTaskPaneHOffset != -1 ? nTaskPaneHOffset : m_nGroupCaptionHorzOffset);
	
	CRect rectText = rectGroup;
	rectText.left += (bShowIcon ? pGroup->m_sizeIcon.cx	+ 5: nCaptionHOffset);
	rectText.top += (nTaskPaneVOffset != -1 ? nTaskPaneVOffset : m_nGroupCaptionVertOffset);
	rectText.right = max(rectText.left, rectText.right - nCaptionHOffset);

	pDC->DrawText (pGroup->m_strName, rectText, DT_SINGLELINE | DT_VCENTER);

	pDC->SetBkMode(nBkModeOld);
	pDC->SelectObject (pFontOld);
	pDC->SetTextColor (clrTextOld);
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnFillTasksGroupInterior(CDC* /*pDC*/, CRect /*rect*/, BOOL /*bSpecial*/)
{
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTasksGroupAreaBorder(CDC* /*pDC*/, CRect /*rect*/, BOOL /*bSpecial*/, BOOL /*bNoTitle*/)
{
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTask(CDC* pDC, CBCGPTask* pTask, CImageList* pIcons, BOOL bIsHighlighted, BOOL bIsSelected)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnDrawTask(pDC, pTask, pIcons, bIsHighlighted, bIsSelected);
		return;
	}

	ASSERT_VALID(pDC);
	ASSERT(pTask != NULL);

	CRect rect = pTask->m_rect;

	if (pTask->m_bIsSeparator)
	{
		CPen* pOldPen = (CPen*) pDC->SelectObject (&globalData.penBarShadow);
		ASSERT (pOldPen != NULL);

		pDC->MoveTo (rect.left, rect.CenterPoint ().y);
		pDC->LineTo (rect.right, rect.CenterPoint ().y);

		pDC->SelectObject (pOldPen);
		return;
	}

	ASSERT_VALID(pTask->m_pGroup);
	ASSERT_VALID(pTask->m_pGroup->m_pPage);

	CBCGPTasksPane* pTaskPane = pTask->m_pGroup->m_pPage->m_pTaskPane;
	ASSERT_VALID (pTaskPane);

	// ---------------
	// Fill background
	// ---------------

	if (bIsSelected)
	{
		pDC->Draw3dRect(rect, globalData.clrBarHilite, globalData.clrBarHilite);
	}

	// ---------
	// Draw icon
	// ---------
	CSize sizeIcon(0, 0);
	::ImageList_GetIconSize (pIcons->m_hImageList, (int*) &sizeIcon.cx, (int*) &sizeIcon.cy);
	if (pTask->m_nIcon >= 0 && sizeIcon.cx > 0)
	{
		CPoint ptIcon = rect.TopLeft();
		if (!pTaskPane->IsWrapTasksEnabled())
		{
			ptIcon.y += max(0, (rect.Height() - sizeIcon.cy) / 2);
		}

		pIcons->Draw (pDC, pTask->m_nIcon, ptIcon, ILD_TRANSPARENT);
	}
	int nTaskPaneOffset = pTaskPane->GetTasksIconHorzOffset();
	rect.left += sizeIcon.cx + (nTaskPaneOffset != -1 ? nTaskPaneOffset : m_nTasksIconHorzOffset);

	// ---------
	// Draw text
	// ---------
	BOOL bIsLabel = (pTask->m_uiCommandID == 0);

	int nBkModeOld = pDC->SetBkMode (TRANSPARENT);
	COLORREF clrTextOld = pDC->GetTextColor ();
	CFont* pFontOld = NULL;

	if (bIsLabel && pTask->m_bIsBold)
	{
		pFontOld = pDC->SelectObject (&globalData.fontBold);
	}
	else if (bIsHighlighted)
	{
		pFontOld = pDC->SelectObject (&globalData.fontUnderline);
	}
	else
	{
		pFontOld = pDC->SelectObject (&globalData.fontRegular);
	}

	COLORREF color;
	if (!bIsLabel && !pTask->m_bEnabled)
	{
		color = globalData.clrGrayedText;
	}
	else if (bIsHighlighted && !bIsLabel)
	{
		color = m_clrTaskPaneHotText;
	}
	else
	{
		color = globalData.clrBarText;
	}

	if (bIsHighlighted)
	{
		pDC->SetTextColor (pTask->m_clrTextHot == (COLORREF)-1 ? color : pTask->m_clrTextHot);
	}
	else
	{
		pDC->SetTextColor (pTask->m_clrText == (COLORREF)-1 ? color : pTask->m_clrText);
	}

	BOOL bMultiline = bIsLabel ? 
		pTaskPane->IsWrapLabelsEnabled () : pTaskPane->IsWrapTasksEnabled ();

	if (bMultiline)
	{
		pDC->DrawText (pTask->m_strName, rect, DT_WORDBREAK);
	}
	else
	{
		CString strText = pTask->m_strName;
		strText.Remove (_T('\n'));
		strText.Remove (_T('\r'));
		pDC->DrawText (strText, rect, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
	}

	pDC->SetBkMode(nBkModeOld);
	pDC->SetTextColor (clrTextOld);
	pDC->SelectObject (pFontOld);
}

#endif

void CBCGPVisualManagerVS2012::OnFillOutlookPageButton (CDC* pDC, const CRect& rect,
										BOOL bIsHighlighted, BOOL bIsPressed,
										COLORREF& clrText)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnFillOutlookPageButton (pDC, rect,
										bIsHighlighted, bIsPressed,
										clrText);
		return;
	}

	ASSERT_VALID (pDC);

	clrText = globalData.clrBarText;

	if (bIsPressed)
	{
		pDC->FillRect(rect, &m_brAccent);
		clrText = RGB(255, 255, 255);
	}
	else if (bIsHighlighted)
	{
		pDC->FillRect(rect, &m_brMenuLight);
	}
	else
	{
		pDC->FillRect(rect, &globalData.brBarFace);
	}
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawOutlookPageButtonBorder (CDC* pDC, 
							CRect& rectBtn, BOOL bIsHighlighted, BOOL bIsPressed)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawOutlookPageButtonBorder (pDC, 
							rectBtn, bIsHighlighted, bIsPressed);
		return;
	}

	ASSERT_VALID (pDC);

	COLORREF clrFrame = bIsHighlighted && bIsPressed ? globalData.clrBarHilite : globalData.clrBarShadow;
	pDC->Draw3dRect (rectBtn, clrFrame, clrFrame);
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawOutlookBarFrame(CDC* pDC, CRect rectFrame)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawOutlookBarFrame(pDC, rectFrame);
	}

	// Don't draw frame
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnFillOutlookBarCaption (CDC* pDC, CRect rectCaption, COLORREF& clrText)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnFillOutlookBarCaption (pDC, rectCaption, clrText);
		return;
	}

	rectCaption.right++;
	pDC->FillRect(rectCaption, &m_brMenuLight);

	clrText = globalData.clrBarText;
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnDrawOutlookPopupButton(CDC* pDC, CRect& rectBtn, BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsOnCaption)
{
	COLORREF clr = CBCGPVisualManagerVS2010::OnDrawOutlookPopupButton(pDC, rectBtn, bIsHighlighted, bIsPressed, bIsOnCaption);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ()  || m_Style == VS2012_LightBlue || m_Style == VS2012_Dark)
	{
		return clr;
	}

	return bIsPressed && bIsHighlighted ? globalData.clrBarHilite : globalData.clrBarText;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetCaptionBarTextColor (CBCGPCaptionBar* pBar)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetCaptionBarTextColor(pBar);
	}

	return globalData.clrBarText;
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawHeaderCtrlBorder (CBCGPHeaderCtrl* pCtrl, CDC* pDC, CRect& rect, BOOL bIsPressed, BOOL bIsHighlighted)
{
	BOOL bIsVMStyle = (pCtrl == NULL || pCtrl->m_bVisualManagerStyle);
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || !bIsVMStyle)
	{
		CBCGPVisualManagerVS2010::OnDrawHeaderCtrlBorder (pCtrl, pDC, rect, bIsPressed, bIsHighlighted);
		return;
	}

	if (bIsPressed)
	{
		pDC->FillRect(rect, &m_brHighlightDn);
	}
	else if (bIsHighlighted)
	{
		pDC->FillRect(rect, &m_brHighlight);
	}
	else
	{
		pDC->FillRect(rect, &m_brDlgBackground);
	}

	pDC->Draw3dRect(rect, globalData.clrBarShadow, globalData.clrBarShadow);
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawHeaderCtrlSortArrow(CBCGPHeaderCtrl* pCtrl, CDC* pDC, CRect& rect, BOOL bIsUp)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode())
	{
		CBCGPVisualManagerVS2010::OnDrawHeaderCtrlSortArrow(pCtrl, pDC, rect, bIsUp);
		return;
	}

	if (pCtrl != NULL && !pCtrl->m_bVisualManagerStyle)
	{
		CBCGPVisualManagerVS2010::OnDrawHeaderCtrlSortArrow(pCtrl, pDC, rect, bIsUp);
		return;
	}
	
	CBCGPMenuImages::Draw(pDC, bIsUp ? CBCGPMenuImages::IdArowUpLarge : CBCGPMenuImages::IdArowDownLarge, rect);
}
//****************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetHeaderCtrlTextColor (CBCGPHeaderCtrl* pCtrl, BOOL bIsPressed, BOOL bIsHighlighted)
{
	BOOL bIsVMStyle = (pCtrl == NULL || pCtrl->m_bVisualManagerStyle);
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || !bIsVMStyle)
	{
		return CBCGPVisualManagerVS2010::GetHeaderCtrlTextColor (pCtrl, bIsPressed, bIsHighlighted);
	}
	
	if (bIsPressed)
	{
		return m_clrHighlighDownText;
	}

	return globalData.clrBarText;
}
//****************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetURLLinkColor (CBCGPURLLinkButton* pButton, BOOL bHover)
{
	ASSERT_VALID (pButton);

	if (!pButton->m_bVisualManagerStyle || pButton->m_bOnGlass || globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetURLLinkColor (pButton, bHover);
	}

	if (pButton->GetSafeHwnd() != NULL && !pButton->IsWindowEnabled())
	{
		return m_clrTextDisabled;
	}
	
	return bHover ? m_clrAccentText : m_clrAccent;
}

#ifndef BCGP_EXCLUDE_POPUP_WINDOW

void CBCGPVisualManagerVS2012::OnFillPopupWindowBackground (CDC* pDC, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnFillPopupWindowBackground (pDC, rect);
		return;
	}

	pDC->FillRect(rect, &m_brDlgBackground);
}
//**********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPopupWindowLinkTextColor(CBCGPPopupWindow* pPopupWnd, BOOL bIsHover)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetPopupWindowLinkTextColor(pPopupWnd, bIsHover);
	}

	return bIsHover ? m_clrAccentText : m_clrAccent;
}
//**********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPopupWindowBorderBorderColor()
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::GetPopupWindowBorderBorderColor();
	}

	return m_clrAccent;
}
//**********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawPopupWindowBorder (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnDrawPopupWindowBorder (pDC, rect);
		return;
	}

	pDC->Draw3dRect(rect, m_clrAccent, m_clrAccent);
}
//**********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawPopupWindowRoundedBorder (CDC* pDC, CRect rect, CBCGPPopupWindow* pPopupWnd, int nCornerRadius)
{
	ASSERT_VALID (pDC);
	
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnDrawPopupWindowRoundedBorder(pDC, rect, pPopupWnd, nCornerRadius);
		return;
	}
	
	CPen pen (PS_SOLID, 1, m_clrAccent);
	CPen* pOldPen = pDC->SelectObject (&pen);
	CBrush* pOldBrush = (CBrush*) pDC->SelectStockObject (NULL_BRUSH);
	
	pDC->RoundRect (rect.left, rect.top, rect.right, rect.bottom, nCornerRadius + 1, nCornerRadius + 1);
	
	pDC->SelectObject (pOldPen);
	pDC->SelectObject (pOldBrush);
}
//**********************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnDrawPopupWindowCaption (CDC* pDC, CRect rectCaption, CBCGPPopupWindow* pPopupWnd)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::OnDrawPopupWindowCaption (pDC, rectCaption, pPopupWnd);
	}

	ASSERT_VALID(pDC);
	ASSERT_VALID(pPopupWnd);

	CRect rectFill = rectCaption;

	switch (pPopupWnd->GetStemLocation())
	{
	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_TopCenter:
	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_TopLeft:
	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_TopRight:
		rectFill.top -= pPopupWnd->GetStemSize();
		break;

	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_Left:
		rectFill.left -= pPopupWnd->GetStemSize();
		break;
	}

	pDC->FillRect(rectFill, &m_brDlgBackground);

	if (pPopupWnd->HasSmallCaption() && pPopupWnd->IsSmallCaptionGripper())
	{
		CBCGPDrawManager dm(*pDC);

		int dx = rectCaption.Height();
		int ellipse_size = min(dx, globalUtils.ScaleByDPI(3));

		for (int x = rectCaption.CenterPoint().x - 3 * dx; x <= rectCaption.CenterPoint().x + 3 * dx; x += dx)
		{
			CRect rect = rectCaption;
			
			rect.left = x + ellipse_size / 2;
			rect.right = rect.left + ellipse_size;
			rect.top = rect.CenterPoint().y - ellipse_size / 2;
			rect.bottom = rect.top + ellipse_size;

			dm.DrawEllipse(rect, IsDarkTheme() ? m_clrAccentLight : m_clrAccent, (COLORREF)-1);
		}
	}

	return globalData.clrBarText;
}
//**********************************************************************************
void CBCGPVisualManagerVS2012::OnErasePopupWindowButton (CDC* pDC, CRect rc, CBCGPPopupWndButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnErasePopupWindowButton(pDC, rc, pButton);
		return;
	}

	if (pButton->IsPressed() || pButton->IsHighlighted())
	{
		OnFillHighlightedArea (pDC, rc, pButton->IsPressed() ? (m_Style == VS2012_Light ? &m_brHighlightChecked : &m_brHighlightDn) : &m_brHighlight, NULL);
	}
	else
	{
		CRect rectParent;
		pButton->GetParent ()->GetClientRect (rectParent);
		
		pButton->GetParent ()->MapWindowPoints (pButton, rectParent);
		OnFillPopupWindowBackground (pDC, rectParent);
	}
}
//**********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawPopupWindowButtonBorder (CDC* pDC, CRect rc, CBCGPPopupWndButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnDrawPopupWindowButtonBorder (pDC, rc, pButton);
		return;
	}

	if (pButton->IsPressed () || pButton->IsHighlighted () || pButton->IsPushed ())
	{
		pDC->Draw3dRect(rc, m_clrAccent, m_clrAccent);
	}
}

#endif	// BCGP_EXCLUDE_POPUP_WINDOW

COLORREF CBCGPVisualManagerVS2012::OnFillListBox(CDC* pDC, CBCGPListBox* pListBox, CRect rectClient)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnFillListBox(pDC, pListBox, rectClient);
	}

	ASSERT_VALID (pDC);

	pDC->FillRect(rectClient, &m_brControl);
	return globalData.clrBarText;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillListBoxItem (CDC* pDC, CBCGPListBox* pListBox, int nItem, CRect rect, BOOL bIsHighlihted, BOOL bIsSelected)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::OnFillListBoxItem (pDC, pListBox, nItem, rect, bIsHighlihted, bIsSelected);
	}

	ASSERT_VALID (pDC);

	COLORREF clrText = (COLORREF)-1;
	BOOL bIsDisabled = pListBox != NULL && !pListBox->IsEnabled(nItem);

	if (bIsDisabled && (bIsSelected || bIsHighlihted))
	{
		pDC->FillSolidRect(rect, globalData.clrBarFace);
		pDC->Draw3dRect(rect, m_clrHighlight, m_clrHighlight);
	}
	else if (bIsSelected)
	{
		pDC->FillRect (rect, &m_brAccent);
		clrText = RGB(255, 255, 255);
	}
	else if (bIsHighlihted)
	{
		pDC->FillRect(rect, &m_brHighlight);
	}
	
	return clrText;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillComboBoxItem (CDC* pDC, CBCGPComboBox* pComboBox, int nItem, CRect rect, BOOL bIsHighlihted, BOOL bIsSelected)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::OnFillComboBoxItem (pDC, pComboBox, nItem, rect, bIsHighlihted, bIsSelected);
	}

	ASSERT_VALID (pDC);

	if (!bIsHighlihted && !bIsSelected)
	{
		pDC->FillRect(rect, &m_brControl);
		return globalData.clrBarText;
	}

	COLORREF clrText = (COLORREF)-1;

	if (bIsSelected)
	{
		pDC->FillRect (rect, &m_brAccent);
		clrText = RGB(255, 255, 255);
	}
	
	if (bIsHighlihted)
	{
		pDC->Draw3dRect(rect, RGB(255, 255, 255), RGB(255, 255, 255));
	}

	return clrText;
}
//*******************************************************************************
BOOL CBCGPVisualManagerVS2012::OnDrawComboBoxText(CDC* pDC, CBCGPComboBox* pComboBox)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnDrawComboBoxText(pDC, pComboBox);
	}

	ASSERT_VALID(pDC);
	ASSERT_VALID(pComboBox);

	CString strText;
	pComboBox->GetWindowText(strText);

	CRect rect;
	pComboBox->GetClientRect(rect);

	BOOL bIsFocused = pComboBox->GetSafeHwnd() == CWnd::GetFocus()->GetSafeHwnd();
	COLORREF clrText = OnFillComboBoxItem(pDC, pComboBox, pComboBox->GetCurSel(), rect, bIsFocused, bIsFocused);

	if (!pComboBox->IsWindowEnabled())
	{
		clrText = m_clrTextDisabled;
	}

	rect.left += 4;

	CFont* pOldFont = pDC->SelectObject(pComboBox->GetFont());
	ASSERT_VALID(pOldFont);

	int nOldBkMode = pDC->SetBkMode(TRANSPARENT);
	COLORREF clrTextOld = pDC->SetTextColor(clrText);

	const int cxDropDown = ::GetSystemMetrics (SM_CXVSCROLL);
	rect.right -= cxDropDown;

	pDC->DrawText(strText, rect, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);

	pDC->SelectObject(pOldFont);
	pDC->SetBkMode(nOldBkMode);
	pDC->SetTextColor(clrTextOld);

	return TRUE;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnDrawEditListCaption(CDC* pDC, CBCGPEditListBase* pList, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::OnDrawEditListCaption(pDC, pList, rect);
	}

	pDC->FillRect(rect, &m_brControl);
	pDC->Draw3dRect(rect, m_clrSeparator, m_clrSeparator);

	return globalData.clrBarText;
}
//*******************************************************************************
HBRUSH CBCGPVisualManagerVS2012::GetListControlFillBrush(CBCGPListCtrl* pListCtrl)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetListControlFillBrush(pListCtrl);
	}

	return (HBRUSH)m_brControl;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetListControlMarkedColor(CBCGPListCtrl* pListCtrl)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style != VS2012_Dark)
	{
		return CBCGPVisualManagerVS2010::GetListControlMarkedColor(pListCtrl);
	}
	
	return globalData.clrBarLight;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetListControlTextColor(CBCGPListCtrl* pListCtrl)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetListControlTextColor(pListCtrl);
	}

	return globalData.clrBarText;
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawShowAllMenuItems (CDC* pDC, CRect rect, CBCGPVisualManager::BCGBUTTON_STATE state)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawShowAllMenuItems (pDC, rect, state);
		return;
	}

	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdArowShowAll, rect,
		m_Style == VS2012_Dark ? CBCGPMenuImages::ImageWhite : CBCGPMenuImages::ImageBlack);
}
//*****************************************************************************
int CBCGPVisualManagerVS2012::GetShowAllMenuItemsHeight (CDC* pDC, const CSize& sizeDefault)
{
	return CBCGPVisualManagerXP::GetShowAllMenuItemsHeight(pDC, sizeDefault);
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawSeparator (CDC* pDC, CBCGPBaseControlBar* pBar, CRect rect, BOOL bIsHoriz)
{
#ifndef BCGP_EXCLUDE_RIBBON
	if (DYNAMIC_DOWNCAST(CBCGPRibbonStatusBar, pBar) != NULL)
	{
		return;
	}

	if (DYNAMIC_DOWNCAST(CBCGPRibbonBackstageView, pBar) != NULL && globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode())
	{
		CBCGPDrawManager dm(*pDC);

		int y = rect.CenterPoint().y;
		dm.DrawLine(rect.left, y, rect.right, y, CBCGPDrawManager::ColorMakeLighter(m_clrAccent, .2));
		return;
	}

#endif

	CBCGPVisualManagerXP::OnDrawSeparator(pDC, pBar, rect, bIsHoriz);
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::GetCalendarColors (const CBCGPCalendar* pCalendar, CBCGPCalendarColors& colors)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::GetCalendarColors(pCalendar, colors);
		return;
	}

	colors.clrCaption = globalData.clrBarFace;
	colors.clrCaptionText = globalData.clrBarText;
	colors.clrSelected = m_clrAccentLight;

	if (GetRValue (m_clrAccentLight) <= 128 ||
		GetGValue (m_clrAccentLight) <= 128 ||
		GetBValue (m_clrAccentLight) <= 128)
	{
		colors.clrSelectedText = RGB(255, 255, 255);
	}
	else
	{
		colors.clrSelectedText = RGB(0, 0, 0);
	}

	colors.clrTodayBorder = m_clrAccentText;

	if (pCalendar != NULL && pCalendar->IsVisualManagerStyle())
	{
		colors.clrBackground = m_clrControl;
		colors.clrText = globalData.clrBarText;
		colors.clrLine = m_clrSeparator;
	}

	if (pCalendar->GetSafeHwnd() != NULL && !pCalendar->IsWindowEnabled())
	{
		colors.clrCaptionText = globalData.clrGrayedText;
		colors.clrSelectedText = globalData.clrGrayedText;
		colors.clrText = globalData.clrGrayedText;
		colors.clrLine = globalData.clrBarShadow;
	}
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawDateTimeDropButton (CDC* pDC, CRect rect, 
	BOOL bDisabled, BOOL bPressed, BOOL bHighlighted, CBCGPDateTimeCtrl* pCtrl)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID (this);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || (m_Style == VS2012_LightBlue && (bHighlighted || bPressed)))
	{
		CBCGPVisualManagerVS2010::OnDrawDateTimeDropButton(pDC, rect, bDisabled, bPressed, bHighlighted, pCtrl);
		return;
	}

	if (!bDisabled)
	{
		COLORREF clrFill = m_clrCombo;
		
		if (bPressed)
		{
			pDC->FillRect(rect, &m_brHighlightDn);
			clrFill = m_clrHighlightDn;
		}
		else if (bHighlighted)
		{
			pDC->FillRect (rect, &m_brAccentLight);
			clrFill = m_clrAccentLight;
		}
		else
		{
			pDC->FillRect (rect, &m_brControl);
			clrFill = m_clrControl;
		}
		
		CBCGPMenuImages::DrawByColor(pDC, CBCGPMenuImages::IdArowDown, rect, clrFill);
	}
	else
	{
		pDC->FillRect (rect, &m_brControl);
		CBCGPMenuImages::Draw(pDC, CBCGPMenuImages::IdArowDown, rect, CBCGPMenuImages::ImageLtGray);
	}
}
//**************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawCaptionBarInfoArea (CDC* pDC, CBCGPCaptionBar* pBar, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pBar->m_clrBarBackground != (COLORREF)-1 || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnDrawCaptionBarInfoArea (pDC, pBar, rect);
		return;
	}

	pDC->FillRect(rect, &m_brMenuLight);
	pDC->Draw3dRect(rect, m_clrMenuBorder, m_clrMenuBorder);
}
//**************************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillCaptionBarButton (CDC* pDC, CBCGPCaptionBar* pBar,
											CRect rect, BOOL bIsPressed, BOOL bIsHighlighted, 
											BOOL bIsDisabled, BOOL bHasDropDownArrow,
											BOOL bIsSysButton)
{
	ASSERT_VALID (pBar);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::OnFillCaptionBarButton (pDC, pBar,
											rect, bIsPressed, bIsHighlighted, 
											bIsDisabled, bHasDropDownArrow, bIsSysButton);
	}

	if (bIsDisabled)
	{
		return (COLORREF)-1;
	}

	COLORREF clrText = globalData.clrBarText;

	if (bIsPressed)
	{
		OnFillHighlightedArea (pDC, rect, &m_brHighlightDn, NULL);
		
		if (GetRValue (m_clrHighlightDn) > 100 &&
			GetGValue (m_clrHighlightDn) > 100 &&
			GetBValue (m_clrHighlightDn) > 100)
		{
			clrText = RGB (0, 0, 0);
		}
		else
		{
			clrText = RGB (255, 255, 255);
		}
	}
	else if (bIsHighlighted)
	{
		OnFillHighlightedArea (pDC, rect, &m_brHighlight, NULL);

		if (GetRValue (m_clrHighlight) > 100 &&
			GetGValue (m_clrHighlight) > 100 &&
			GetBValue (m_clrHighlight) > 100)
		{
			clrText = RGB (0, 0, 0);
		}
		else
		{
			clrText = RGB (255, 255, 255);
		}
	}
	else if (!bIsSysButton)
	{
		pDC->FillRect (rect, &m_brMenuLight);
	}

	return clrText;
}
//**************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawCaptionBarButtonBorder (CDC* pDC, CBCGPCaptionBar* pBar,
											CRect rect, BOOL bIsPressed, BOOL bIsHighlighted, 
											BOOL bIsDisabled, BOOL bHasDropDownArrow,
											BOOL bIsSysButton)
{
	ASSERT_VALID (pBar);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::OnDrawCaptionBarButtonBorder (pDC, pBar,
											rect, bIsPressed, bIsHighlighted, 
											bIsDisabled, bHasDropDownArrow, bIsSysButton);
		return;
	}

	ASSERT_VALID (pDC);

	if (bIsHighlighted)
	{
		if (bIsSysButton && bIsPressed && m_clrPressedButtonBorder != (COLORREF)-1)
		{
			pDC->Draw3dRect (rect, m_clrPressedButtonBorder, m_clrPressedButtonBorder);
		}
		else
		{
			pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarDkShadow);
		}
	}
	else if (!bIsSysButton)
	{
		pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarDkShadow);
	}
}

#ifndef BCGP_EXCLUDE_GRID_CTRL

COLORREF CBCGPVisualManagerVS2012::OnFillGridGroupByBoxBackground (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::OnFillGridGroupByBoxBackground (pDC, rect);
	}

	pDC->FillRect (rect, &globalData.brBarFace);

	return globalData.clrBarText;
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillGridGroupByBoxTitleBackground (CDC* pDC, CRect rect)
{
	return CBCGPVisualManagerXP::OnFillGridGroupByBoxTitleBackground (pDC, rect);
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetGridGroupByBoxLineColor () const
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::GetGridGroupByBoxLineColor ();
	}

	return m_clrSeparator;
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawGridGroupByBoxItemBorder (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnDrawGridGroupByBoxItemBorder (pCtrl, pDC, rect);
		return;
	}

	ASSERT_VALID(pDC);

	pDC->FillRect (rect, &m_brMenuLight);
	pDC->Draw3dRect(rect, m_clrSeparator, m_clrSeparator);
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetGridLeftOffsetColor (CBCGPGridCtrl* pCtrl)
{
	ASSERT_VALID(pCtrl);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || !pCtrl->IsControlBarColors ())
	{
		return CBCGPVisualManagerXP::GetGridLeftOffsetColor (pCtrl);
	}

	return m_Style == VS2012_Dark ? globalData.clrBarDkShadow : m_clrHighlight;
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetGridItemSortedColor (CBCGPGridCtrl* pCtrl)
{
	return CBCGPVisualManagerXP::GetGridItemSortedColor(pCtrl);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnFillGridGroupBackground (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rectFill)
{
	CBCGPVisualManagerXP::OnFillGridGroupBackground(pCtrl, pDC, rectFill);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawGridGroupUnderline (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rectFill)
{
	CBCGPVisualManagerXP::OnDrawGridGroupUnderline(pCtrl, pDC, rectFill);
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillGridRowBackground (CBCGPGridCtrl* pCtrl, 
												  CDC* pDC, CRect rectFill, BOOL bSelected)
{
	return CBCGPVisualManagerXP::OnFillGridRowBackground(pCtrl, pDC, rectFill, bSelected);
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillGridItem (CBCGPGridCtrl* pCtrl, 
											CDC* pDC, CRect rectFill,
											BOOL bSelected, BOOL bActiveItem, BOOL bSortedColumn)
{
	return CBCGPVisualManagerXP::OnFillGridItem(pCtrl, pDC, rectFill, bSelected, bActiveItem, bSortedColumn);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawGridHeaderMenuButton (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rect, 
		BOOL bHighlighted, BOOL bPressed, BOOL bDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnDrawGridHeaderMenuButton(pCtrl, pDC, rect, bHighlighted, bPressed, bDisabled);
		return;
	}

	ASSERT_VALID (pDC);

	rect.DeflateRect (1, 1);

	if (bHighlighted)
	{
		OnFillHighlightedArea (pDC, rect, &m_brHighlight, NULL);
		pDC->Draw3dRect (rect, m_clrMenuItemBorder, m_clrMenuItemBorder);
	}
	else if (bPressed)
	{
		OnFillHighlightedArea (pDC, rect, &m_brHighlightDn, NULL);
		pDC->Draw3dRect (rect, m_clrMenuItemBorder, m_clrMenuItemBorder);
	}
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawGridSelectionBorder (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnDrawGridSelectionBorder (pCtrl, pDC, rect);
		return;
	}

	pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarDkShadow);
	rect.DeflateRect (1, 1);
	pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarDkShadow);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawGridDataBar(CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawGridDataBar(pCtrl, pDC, rect);
		return;
	}
	
	ASSERT_VALID(pDC);
	
	COLORREF clrFill = (COLORREF)-1;

	if (IsDarkTheme())
	{
		clrFill = RGB (62, 109, 171);
	}
	else
	{
		clrFill = RGB (183, 203, 228);
	}
	
	CBrush brFill(clrFill);
	pDC->FillRect (rect, &brFill);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::GetGridColorScaleBaseColors(CBCGPGridCtrl* pCtrl, COLORREF& clrLow, COLORREF& clrMid, COLORREF& clrHigh)
{
	CBCGPVisualManagerVS2010::GetGridColorScaleBaseColors(pCtrl, clrLow, clrMid, clrHigh);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return;
	}


	if (IsDarkTheme())
	{
		clrLow = RGB (50, 0, 0);
		clrMid = RGB (174, 112, 0);
		clrHigh = RGB (0, 100, 0);
	}
	else
	{
		double dblRatio = .2;

		clrLow = CBCGPDrawManager::ColorMakeLighter(clrLow, dblRatio);
		clrMid = CBCGPDrawManager::ColorMakeLighter(clrMid, dblRatio);
		clrHigh = CBCGPDrawManager::ColorMakeLighter(clrHigh, dblRatio);
	}
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetGridTreeLineColor (CBCGPGridCtrl* pCtrl)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode())
	{
		return CBCGPVisualManager::GetGridTreeLineColor(pCtrl);
	}
	
	if (IsDarkTheme() && pCtrl->IsVisualManagerStyle())
	{
		return CBCGPVisualManager::GetGridTreeLineColor(pCtrl);
	}

	return globalData.clrBarShadow;
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawGridSortArrow(CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rectArrow, BOOL bIsUp)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode())
	{
		CBCGPVisualManagerVS2010::OnDrawGridSortArrow(pCtrl, pDC, rectArrow, bIsUp);
		return;
	}

	CBCGPMenuImages::Draw (pDC, bIsUp ? CBCGPMenuImages::IdArowUpLarge : CBCGPMenuImages::IdArowDownLarge, rectArrow);
}
//********************************************************************************
BOOL CBCGPVisualManagerVS2012::OnSetGridColorTheme (CBCGPGridCtrl* pCtrl, BCGP_GRID_COLOR_DATA& theme)
{
	BOOL bRes = CBCGPVisualManagerVS2010::OnSetGridColorTheme (pCtrl, theme);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return bRes;
	}

	COLORREF clrTextSel;

	if (GetRValue (m_clrAccentLight) <= 128 ||
		GetGValue (m_clrAccentLight) <= 128 ||
		GetBValue (m_clrAccentLight) <= 128)
	{
		clrTextSel = RGB(255, 255, 255);
	}
	else
	{
		clrTextSel = RGB(0, 0, 0);
	}

	theme.m_clrBackground = pCtrl->IsTreeControl() ? GetTreeControlFillColor(NULL) : globalData.clrBarLight;
	theme.m_clrText = globalData.clrBarText;
	theme.m_clrHorzLine = theme.m_clrVertLine = globalData.clrBarShadow;
	
	theme.m_EvenColors.m_clrBackground = pCtrl->IsTreeControl() ? theme.m_clrBackground : CBCGPDrawManager::SmartMixColors(globalData.clrBarFace, globalData.clrBarHilite);
	theme.m_EvenColors.m_clrGradient = (COLORREF)-1;
	theme.m_EvenColors.m_clrText = globalData.clrBarText;

	if (pCtrl->IsTreeControl())
	{
		theme.m_OddColors.m_clrBackground = m_Style == VS2012_Dark ? 
			CBCGPDrawManager::ColorMakeLighter(theme.m_EvenColors.m_clrBackground, .2) : 
			CBCGPDrawManager::ColorMakeDarker(theme.m_EvenColors.m_clrBackground, .05);
	}
	else
	{
		theme.m_OddColors.m_clrBackground = m_Style == VS2012_Light? 
			CBCGPDrawManager::ColorMakeDarker(theme.m_EvenColors.m_clrBackground, .01) : 
			globalData.clrBarLight;
	}

	theme.m_OddColors.m_clrGradient = (COLORREF)-1;
	theme.m_OddColors.m_clrText = globalData.clrBarText;

	if (m_Style != VS2012_LightBlue)
	{
		theme.m_SelColors.m_clrBackground = m_clrAccentLight;
		theme.m_SelColors.m_clrGradient = (COLORREF)-1;
		theme.m_SelColors.m_clrText = clrTextSel;
	}

	theme.m_LeftOffsetColors.m_clrBackground = globalData.clrBarLight;
	theme.m_LeftOffsetColors.m_clrGradient = (COLORREF)-1;
	theme.m_LeftOffsetColors.m_clrBorder = globalData.clrBarShadow;
	theme.m_LeftOffsetColors.m_clrText = globalData.clrBarText;

	theme.m_HeaderColors.m_clrBackground = (COLORREF)-1;

	if (m_Style != VS2012_LightBlue)
	{
		theme.m_HeaderColors.m_clrBorder = globalData.clrBarShadow;
		theme.m_HeaderColors.m_clrGradient = (COLORREF)-1;
		theme.m_HeaderColors.m_clrText = globalData.clrBarText;

		theme.m_HeaderSelColors.m_clrBackground = m_clrAccentLight;
		theme.m_HeaderSelColors.m_clrBorder = globalData.clrBarShadow;
		theme.m_HeaderSelColors.m_clrGradient = (COLORREF)-1;
		theme.m_HeaderSelColors.m_clrText = clrTextSel;
	}

	theme.m_GroupSelColors.m_clrBorder = theme.m_GroupColors.m_clrBorder = globalData.clrBarShadow;
	theme.m_GroupSelColors.m_clrBackground = theme.m_GroupColors.m_clrBackground = globalData.clrBarFace;
	theme.m_GroupSelColors.m_clrGradient = theme.m_GroupColors.m_clrGradient = (COLORREF)-1;
	theme.m_GroupSelColors.m_clrText = theme.m_GroupColors.m_clrText = globalData.clrBarText;

	return TRUE;
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillReportCtrlRowBackground (CBCGPGridCtrl* pCtrl, 
												  CDC* pDC, CRect rectFill,
												  BOOL bSelected, BOOL bGroup)
{
	return CBCGPVisualManagerXP::OnFillReportCtrlRowBackground (pCtrl, pDC, rectFill, bSelected, bGroup);
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetReportCtrlGroupBackgoundColor ()
{
	return CBCGPVisualManagerXP::GetReportCtrlGroupBackgoundColor ();
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnFillGridBackground (CDC* pDC, CRect rect)
{
	CBCGPVisualManagerXP::OnFillGridBackground(pDC, rect);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnFillGridHeaderBackground (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rect)
{
	if (pCtrl == NULL || !pCtrl->IsVisualManagerStyle() ||
		globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnFillGridHeaderBackground (pCtrl, pDC, rect);
		return;
	}

	pDC->FillRect(rect, &globalData.brBarFace);
}
//********************************************************************************
BOOL CBCGPVisualManagerVS2012::OnDrawGridHeaderItemBorder (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rect, BOOL bPressed)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::OnDrawGridHeaderItemBorder (pCtrl, pDC, rect, bPressed);
	}

	if (bPressed)
	{
		CRect rectFill = rect;
		rectFill.right--;

		pDC->FillRect(rectFill, &m_brHighlightDn);
	}

	CPen* pOldPen = pDC->SelectObject (&m_penSeparator);
	
	ASSERT_VALID(pDC);
	
	int nHeight = rect.Height () / 5;
	
	if (pCtrl->GetHeaderFlags () & BCGP_GRID_HEADER_FORCESIMPLEBORDERS)
	{
		nHeight = 0;
	}
	
	pDC->MoveTo (rect.right - 1, rect.top + nHeight);
	pDC->LineTo (rect.right - 1, rect.bottom - nHeight);
	
	CPen pen (PS_SOLID, 1, globalData.clrBarDkShadow);
	pDC->SelectObject (&pen);

	pDC->MoveTo (rect.left, rect.top);
	pDC->LineTo (rect.right, rect.top);
	
	pDC->MoveTo (rect.left, rect.bottom - 1);
	pDC->LineTo (rect.right, rect.bottom - 1);
	
	pDC->SelectObject (pOldPen);

	return FALSE;
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnFillGridRowHeaderBackground (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rect)
{
	if (pCtrl == NULL || !pCtrl->IsVisualManagerStyle() ||
		globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnFillGridRowHeaderBackground (pCtrl, pDC, rect);
		return;
	}

	pDC->FillRect(rect, &globalData.brBarFace);
}
//********************************************************************************
BOOL CBCGPVisualManagerVS2012::OnDrawGridRowHeaderItemBorder (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rect, BOOL bPressed)
{
	ASSERT_VALID (pDC);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::OnDrawGridRowHeaderItemBorder (pCtrl, pDC, rect, bPressed);
	}

	CPen* pOldPen = pDC->SelectObject (&m_penSeparator);

	pDC->MoveTo (rect.left, rect.top);
	pDC->LineTo (rect.right - 1, rect.top);

	pDC->MoveTo (rect.left, rect.bottom);
	pDC->LineTo (rect.right - 1, rect.bottom);

	CPen pen (PS_SOLID, 1, globalData.clrBarDkShadow);
	pDC->SelectObject (&pen);

	pDC->MoveTo (rect.left, rect.top);
	pDC->LineTo (rect.left, rect.bottom);

	pDC->MoveTo (rect.right - 1, rect.top);
	pDC->LineTo (rect.right - 1, rect.bottom);

	pDC->SelectObject (pOldPen);

	return FALSE;
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnFillGridSelectAllAreaBackground (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rect, BOOL bPressed)
{
	if (pCtrl == NULL || !pCtrl->IsVisualManagerStyle() ||
		globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerXP::OnFillGridSelectAllAreaBackground (pCtrl, pDC, rect, bPressed);
		return;
	}

	pDC->FillRect(rect, &globalData.brBarFace);
}
//********************************************************************************
BOOL CBCGPVisualManagerVS2012::OnDrawGridSelectAllAreaBorder (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rect, BOOL bPressed)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerXP::OnDrawGridSelectAllAreaBorder (pCtrl, pDC, rect, bPressed);
	}

	pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarDkShadow);
	return FALSE;
}

#endif // BCGP_EXCLUDE_GRID_CTRL

//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetEditBackSidebarColor(CBCGPEditCtrl* pEdit)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetEditBackSidebarColor(pEdit);
	}

	return m_Style != VS2012_Dark ? globalData.clrBarFace : m_clrControl;
}
//********************************************************************************
HBRUSH CBCGPVisualManagerVS2012::GetEditBackgroundBrush(CBCGPEditCtrl* pEdit)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetEditBackgroundBrush(pEdit);
	}

	return m_Style != VS2012_Dark ? m_brWhite : m_brControl;
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetEditTextColor(CBCGPEditCtrl* pEdit)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetEditTextColor(pEdit);
	}

	return IsDarkTheme() ? RGB(200, 200, 200) : globalData.clrBarText;
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetEditLineNumbersBarBackColor(CBCGPEditCtrl* pEdit)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetEditLineNumbersBarBackColor(pEdit);
	}

	return m_Style == VS2012_LightBlue ? m_clrControl : m_clrControl;
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetEditLineNumbersBarTextColor(CBCGPEditCtrl* pEdit)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetEditLineNumbersBarTextColor(pEdit);
	}

	return m_clrAccentText;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetEditOutlineColor(CBCGPEditCtrl* pEdit)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetEditOutlineColor(pEdit);
	}

	return globalData.clrBarDkShadow;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetTreeControlFillColor(CBCGPTreeCtrl* pTreeCtrl, BOOL bIsSelected, BOOL bIsFocused, BOOL bIsDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetTreeControlFillColor(pTreeCtrl, bIsSelected, bIsFocused, bIsDisabled);
	}

	if (bIsSelected)
	{
		return bIsFocused ? m_clrAccent : (IsDarkTheme() ? globalData.clrBarLight : globalData.clrBarShadow);
	}

	return m_clrControl;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetTreeControlTextColor(CBCGPTreeCtrl* pTreeCtrl, BOOL bIsSelected, BOOL bIsFocused, BOOL bIsDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetTreeControlTextColor(pTreeCtrl, bIsSelected, bIsFocused, bIsDisabled);
	}

	if (bIsSelected)
	{
		return bIsFocused ? RGB(255, 255, 255) : globalData.clrBarText;
	}

	return globalData.clrBarText;
}
//*********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetIntelliSenseFillColor(CBCGPBaseIntelliSenseLB* pCtrl, BOOL bIsSelected)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetIntelliSenseFillColor(pCtrl, bIsSelected);
	}
	
	return bIsSelected ? m_clrAccentLight : m_clrControl;
}
//*********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetIntelliSenseTextColor(CBCGPBaseIntelliSenseLB* pCtrl, BOOL bIsSelected)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetIntelliSenseTextColor(pCtrl, bIsSelected);
	}

	if (bIsSelected)
	{
		BOOL bIsDarkBackground = (GetRValue (m_clrAccentLight) <= 128 ||
			GetGValue (m_clrAccentLight) <= 128 ||
			GetBValue (m_clrAccentLight) <= 128);
		
		return (bIsDarkBackground) ? RGB(241, 241, 241) : RGB(30, 30, 30);
	}
	
	return globalData.clrBarText;
}
//*********************************************************************************
CBrush& CBCGPVisualManagerVS2012::GetEditCtrlBackgroundBrush(CBCGPEdit* pEdit)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetEditCtrlBackgroundBrush(pEdit);
	}

	if (pEdit->GetSafeHwnd() != NULL && ((pEdit->GetStyle() & ES_READONLY) == ES_READONLY || !pEdit->IsWindowEnabled()))
	{
		return GetDlgBackBrush (pEdit->GetParent());
	}
	
	return m_brControl;
}
//*********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetEditCtrlTextColor(CBCGPEdit* pEdit)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetEditCtrlTextColor(pEdit);
	}

	return globalData.clrBarText;
}
//*************************************************************************************
HBRUSH CBCGPVisualManagerVS2012::GetToolbarEditColors(CBCGPToolbarButton* pButton, COLORREF& clrText, COLORREF& clrBk)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetToolbarEditColors(pButton, clrText, clrBk);
	}

	if (pButton != NULL && (pButton->m_nStyle & TBBS_DISABLED))
	{
		clrText = m_clrTextDisabled;
		clrBk = globalData.clrBarText;

		return (HBRUSH) globalData.brBarFace.GetSafeHandle ();
	}

	clrText = globalData.clrBarText;
	clrBk = m_clrCombo;

	return (HBRUSH) m_brBarBkgnd.GetSafeHandle ();
}
//*************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetToolbarEditPromptColor()
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetToolbarEditPromptColor();
	}

	return m_clrEditPrompt;
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::GetEditSimplifiedIconColors(COLORREF& clrSimple, COLORREF& clrHot, BOOL bIsCloseButton)
{
	CBCGPVisualManagerVS2010::GetEditSimplifiedIconColors(clrSimple, clrHot, bIsCloseButton);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return;
	}

	clrHot = m_clrAccentHilight;
}
//*************************************************************************************
COLORREF CBCGPVisualManagerVS2012::BreadcrumbFillBackground (CDC& dc, CBCGPBreadcrumb* pControl, CRect rectFill)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || !pControl->m_bVisualManagerStyle)
	{
		return CBCGPVisualManagerVS2010::BreadcrumbFillBackground(dc, pControl, rectFill);
	}

	COLORREF clrBackground = pControl->GetBackColor();
	if (clrBackground != CLR_INVALID)
	{
		return CBCGPVisualManagerVS2010::BreadcrumbFillBackground(dc, pControl, rectFill);
	}

	dc.FillRect(&rectFill, &m_brControl);
	return pControl->IsWindowEnabled() ? globalData.clrBarText : m_clrTextDisabled;
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::BreadcrumbFillProgress (CDC& dc, CBCGPBreadcrumb* pControl, CRect rectProgress, CRect rectChunk, int nValue)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || !pControl->m_bVisualManagerStyle || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::BreadcrumbFillProgress(dc, pControl, rectProgress, rectChunk, nValue);
		return;
	}
	
	if (rectChunk.IsRectEmpty ())
	{
		return;
	}

	if (pControl->IsOnGlass())
	{
		CBCGPDrawManager dm (dc);
		dm.DrawRect (rectChunk, m_Style == VS2012_Dark ? m_clrAccent : m_clrAccentLight, (COLORREF)-1);
	}
	else
	{
		dc.FillRect(rectChunk, m_Style == VS2012_Dark ? &m_brAccent : &m_brAccentLight);
	}
}
//*************************************************************************************
BOOL CBCGPVisualManagerVS2012::IsOwnerDrawScrollBar () const
{
	return globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode ();
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnScrollBarDrawThumb (CDC* pDC, CBCGPScrollBar* /*pScrollBar*/, CRect rect, 
		BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return;
	}

	ASSERT_VALID(pDC);

	CBCGPDrawManager dm (*pDC);
	dm.DrawRect (rect, m_Style == VS2012_Dark ? m_clrHighlight : globalData.clrBarFace, (COLORREF)-1);

	if (bDisabled)
	{
		return;
	}

	COLORREF clrFill = bPressed ? m_clrEditPrompt : bHighlighted ? globalData.clrBarDkShadow : globalData.clrBarShadow;

	if (bHorz)
	{
		rect.DeflateRect(0, 4);
	}
	else
	{
		rect.DeflateRect(4, 0);
	}

	dm.DrawRect (rect, clrFill, (COLORREF)-1);
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnScrollBarDrawButton (CDC* pDC, CBCGPScrollBar* pScrollBar, CRect rect, 
		BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bFirst, BOOL bDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return;
	}

	CBCGPDrawManager dm (*pDC);
	dm.DrawRect (rect, m_Style == VS2012_Dark ? m_clrHighlight : globalData.clrBarFace, (COLORREF)-1);

	CBCGPMenuImages::IMAGES_IDS ids;
	if (bHorz)
	{
		if (pScrollBar->GetSafeHwnd() != NULL && (pScrollBar->GetExStyle() & WS_EX_LAYOUTRTL))
		{
			bFirst = !bFirst;
		}

		ids = bFirst ? CBCGPMenuImages::IdArowLeftTab3d : CBCGPMenuImages::IdArowRightTab3d;
	}
	else
	{
		ids = bFirst ? CBCGPMenuImages::IdArowUpLarge : CBCGPMenuImages::IdArowDownLarge;
	}

	CBCGPMenuImages::IMAGE_STATE state = bDisabled ? (m_Style == VS2012_Dark ? CBCGPMenuImages::ImageDkGray : CBCGPMenuImages::ImageLtGray) : 
		bPressed ? (m_Style == VS2012_Dark ? CBCGPMenuImages::ImageWhite : CBCGPMenuImages::ImageBlack) :
		bHighlighted ? CBCGPMenuImages::ImageBlack2 : CBCGPMenuImages::ImageGray;

	CBCGPMenuImages::Draw (pDC, ids, rect, state);
}
//*************************************************************************************
void CBCGPVisualManagerVS2012::OnScrollBarFillBackground (CDC* pDC, CBCGPScrollBar* /*pScrollBar*/, CRect rect, 
		BOOL /*bHorz*/, BOOL /*bHighlighted*/, BOOL /*bPressed*/, BOOL /*bFirst*/, BOOL /*bDisabled*/)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return;
	}

	CBCGPDrawManager dm (*pDC);
	dm.DrawRect (rect, m_Style == VS2012_Dark ? m_clrHighlight : globalData.clrBarFace, (COLORREF)-1);
}

#ifndef BCGP_EXCLUDE_PLANNER
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnFillPlanner (CDC* pDC, CBCGPPlannerView* pView, 
		CRect rect, BOOL bWorkingArea)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pView);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnFillPlanner (pDC, pView, rect, bWorkingArea);
		return;
	}

	if (m_bPlannerBackItemSelected)
	{
		pDC->FillRect (rect, &m_brAccent);
	}
	else
	{
		CBCGPVisualManagerVS2010::OnFillPlanner (pDC, pView, rect, bWorkingArea);
	}

	if (m_bPlannerBackItemToday && DYNAMIC_DOWNCAST (CBCGPPlannerViewDay, pView) == NULL)
	{
		pDC->Draw3dRect (rect, m_clrPlannerTodayLine, m_clrPlannerTodayLine);
	}
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillPlannerCaption (CDC* pDC,
		CBCGPPlannerView* pView, CRect rect, BOOL bIsToday, BOOL bIsSelected,
		BOOL bNoBorder/* = FALSE*/, BOOL bHorz /*= TRUE*/)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnFillPlannerCaption (pDC,
			pView, rect, bIsToday, bIsSelected, bNoBorder, bHorz);
	}

	const BOOL bMonth = DYNAMIC_DOWNCAST(CBCGPPlannerViewMonth, pView) != NULL ||
		DYNAMIC_DOWNCAST(CBCGPPlannerViewWeek, pView) != NULL;;

	if (bMonth && m_bPlannerCaptionBackItemHeader)
	{
		return globalData.clrBarText;
	}

	ASSERT_VALID (pDC);

	BOOL bDay = FALSE;

	if (!bMonth)
	{
		bDay = pView->IsKindOf (RUNTIME_CLASS (CBCGPPlannerViewDay));

		if (bDay)
		{
			if (!bIsToday)
			{
				if (pView->IsKindOf (RUNTIME_CLASS (CBCGPPlannerViewMulti)))
				{
					rect.top++;
					rect.left++;
				}
				else
				{
					rect.left++;
				}
			}
		}
	}
	else
	{
		if (!bIsToday)
		{
			rect.bottom--;
		}
	}

	COLORREF clrText   = globalData.clrBarText;
	COLORREF clrBorder = CLR_DEFAULT;

	if (bIsToday)
	{
		CBrush br (m_clrPlannerTodayFill);
		pDC->FillRect (rect, &br);

		clrText = RGB(255, 255, 255);
		clrBorder = m_clrPlannerTodayLine;
	}
	else
	{
		pDC->FillRect (rect, &globalData.brBarFace);
	}

	if (clrBorder != CLR_DEFAULT && !bNoBorder)
	{
		pDC->Draw3dRect (rect, clrBorder, clrBorder);
	}

	return clrText;
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnDrawPlannerCaptionText (CDC* pDC, 
		CBCGPPlannerView* pView, CRect rect, const CString& strText, 
		COLORREF clrText, int nAlign, BOOL bHighlight)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawPlannerCaptionText (pDC,
			pView, rect, strText, clrText, nAlign, bHighlight);
		return;
	}

	const int nTextMargin = 2;

	rect.DeflateRect (nTextMargin, 0);

	COLORREF clrOld = pDC->SetTextColor (clrText);

	pDC->DrawText (strText, rect, DT_SINGLELINE | DT_VCENTER | nAlign);

	pDC->SetTextColor (clrOld);
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::GetPlannerAppointmentColors (CBCGPPlannerView* pView,
		BOOL bSelected, BOOL bSimple, DWORD dwDrawFlags, 
		COLORREF& clrBack1, COLORREF& clrBack2,
		COLORREF& clrFrame1, COLORREF& clrFrame2, COLORREF& clrText)
{
	CBCGPVisualManagerVS2010::GetPlannerAppointmentColors (pView,
		bSelected, bSimple, dwDrawFlags, clrBack1, clrBack2, clrFrame1, clrFrame2, clrText);

	if (globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode ())
	{
		clrFrame1 = m_clrPlannerTodayFill;
		clrFrame2 = !bSelected ? clrFrame1 : RGB(0, 0, 0);
	}
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPlannerAppointmentDurationColor(CBCGPPlannerView* pView, const CBCGPAppointment* pApp) const
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetPlannerAppointmentDurationColor(pView, pApp);
	}

	return pApp->GetDurationColor() == CLR_DEFAULT
			? m_clrPlannerTodayFill
			: pApp->GetDurationColor();
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPlannerAppointmentTimeColor (CBCGPPlannerView* pView,
		BOOL bSelected, BOOL bSimple, DWORD dwDrawFlags)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () ||
		(bSelected && (dwDrawFlags & BCGP_PLANNER_DRAW_APP_OVERRIDE_SELECTION) == 0))
	{
		return CBCGPVisualManagerVS2010::GetPlannerAppointmentTimeColor (pView,
			bSelected, bSimple, dwDrawFlags);
	}

	return CLR_DEFAULT;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPlannerHourLineColor (CBCGPPlannerView* pView,
		BOOL bWorkingHours, BOOL bHour)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () ||
		GetPlannerViewBackgroundColor (pView) != m_clrPlannerWork)
	{
		return CBCGPVisualManagerVS2010::GetPlannerHourLineColor (pView, 
			bWorkingHours, bHour);
	}

	return bHour
			? m_clrPalennerLine
			: CBCGPDrawManager::PixelAlpha
				(
					bWorkingHours
						? GetPlannerViewWorkingColor (pView)
						: GetPlannerViewNonWorkingColor (pView),
					m_clrPalennerLine, 50
				);
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPlannerViewWorkingColor (CBCGPPlannerView* pView)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () ||
		GetPlannerViewBackgroundColor (pView) != m_clrPlannerWork)
	{
		return CBCGPVisualManagerVS2010::GetPlannerViewWorkingColor (pView);
	}

	return globalData.clrBarHilite;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPlannerViewNonWorkingColor (CBCGPPlannerView* pView)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () ||
		GetPlannerViewBackgroundColor (pView) != m_clrPlannerWork)
	{
		return CBCGPVisualManagerVS2010::GetPlannerViewNonWorkingColor (pView);
	}

	return CBCGPDrawManager::PixelAlpha(globalData.clrBarHilite, globalData.clrBarFace, !IsDarkTheme() ? 15 : 50);
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPlannerSelectionColor (CBCGPPlannerView* pView)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetPlannerSelectionColor (pView);
	}

	return m_clrAccent;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetPlannerSeparatorColor (CBCGPPlannerView* pView)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetPlannerSeparatorColor (pView);
	}

	return m_clrPalennerLine;
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillPlannerTimeBar (CDC* pDC, 
		CBCGPPlannerView* pView, CRect rect, COLORREF& clrLine)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnFillPlannerTimeBar (pDC, pView, rect, clrLine);
	}

	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBarFace);
	
	clrLine = m_clrPalennerLine;

	return globalData.clrBarText;
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnDrawPlannerTimeLine (CDC* pDC, CBCGPPlannerView* pView, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawPlannerTimeLine (pDC, pView, rect);
		return;
	}

	ASSERT_VALID (pDC);

	if (DYNAMIC_DOWNCAST(CBCGPPlannerViewSchedule, pView) == NULL)
	{
		rect.top = rect.bottom - 3;

		pDC->FillRect(rect, &m_brAccentLight);

		CBCGPPenSelector sel(*pDC, m_clrAccent);

		pDC->MoveTo (rect.left, rect.bottom - 2);
		pDC->LineTo (rect.right, rect.bottom - 2);
	}
	else
	{
		rect.left = rect.right - 3;

		pDC->FillRect(rect, &m_brAccentLight);

		CBCGPPenSelector sel(*pDC, m_clrAccent);

		pDC->MoveTo (rect.right - 2, rect.top);
		pDC->LineTo (rect.right - 2, rect.bottom);
	}
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnFillPlannerWeekBar (CDC* pDC, 
		CBCGPPlannerView* pView, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnFillPlannerWeekBar (pDC, pView, rect);
		return;
	}

	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBarFace);
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnDrawPlannerHeader (CDC* pDC, 
	CBCGPPlannerView* pView, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawPlannerHeader (pDC, pView, rect);
		return;
	}

	ASSERT_VALID (pDC);

	COLORREF clr = GetPlannerSeparatorColor (pView);

	if (DYNAMIC_DOWNCAST(CBCGPPlannerViewMonth, pView) != NULL)
	{
		clr = globalData.clrBarFace;
	}

	CBrush br (clr);
	pDC->FillRect (rect, &br);

	if (pView->IsKindOf (RUNTIME_CLASS (CBCGPPlannerViewDay)))
	{
		if (rect.left == pView->GetAppointmentsRect().left)
		{
			CRect rect1 (rect);
			rect1.right = rect1.left + 1;

			if (pView->IsKindOf (RUNTIME_CLASS (CBCGPPlannerViewMulti)))
			{
				rect1.top++;
			}

			pDC->FillRect (rect1, &globalData.brBarFace);
		}
	}
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnDrawPlannerHeaderPane (CDC* pDC, 
	CBCGPPlannerView* pView, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawPlannerHeaderPane (pDC, pView, rect);
		return;
	}

	if (DYNAMIC_DOWNCAST(CBCGPPlannerViewMonth, pView) != NULL)
	{
		pDC->Draw3dRect (rect.right - 1, rect.top - 2, 1, rect.Height () + 4, 
			m_clrPalennerLine, m_clrPalennerLine);
	}
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnFillPlannerHeaderAllDay (CDC* pDC, 
		CBCGPPlannerView* pView, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnFillPlannerHeaderAllDay (pDC, pView, rect);
		return;
	}

	ASSERT_VALID (pDC);

	rect.bottom--;

	pDC->FillRect (rect, &globalData.brBarFace);

	CPen pen(PS_SOLID, 0, m_clrPalennerLine);
	CPen* pOldPen = (CPen*)pDC->SelectObject(&pen);

	pDC->MoveTo(rect.left, rect.bottom);
	pDC->LineTo(rect.right, rect.bottom);

	pDC->SelectObject(pOldPen);
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnDrawPlannerHeaderAllDayItem (CDC* pDC, 
		CBCGPPlannerView* pView, CRect rect, BOOL bIsToday, BOOL bIsSelected)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawPlannerHeaderAllDayItem (pDC, pView, rect, 
			bIsToday, bIsSelected);
		return;
	}

	ASSERT_VALID (pDC);

	rect.left++;
	if (bIsSelected)
	{
		CBrush br (GetPlannerSelectionColor (pView));
		pDC->FillRect (rect, &br);
	}

	if (bIsToday)
	{
 		rect.top--;
		rect.left--;
		rect.bottom++;
		pDC->Draw3dRect (rect, m_clrPlannerTodayLine, m_clrPlannerTodayLine);
	}
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::PreparePlannerBackItem (BOOL bIsToday, BOOL bIsSelected)
{
	m_bPlannerBackItemToday    = bIsToday;
	m_bPlannerBackItemSelected = bIsSelected;
}
//*******************************************************************************
DWORD CBCGPVisualManagerVS2012::GetPlannerDrawFlags () const
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetPlannerDrawFlags ();
	}

	return BCGP_PLANNER_DRAW_APP_OVERRIDE_SELECTION |
		   BCGP_PLANNER_DRAW_APP_DURATION_SHAPE | 
		   BCGP_PLANNER_DRAW_APP_NO_MULTIDAY_CLOCKS |
		   BCGP_PLANNER_DRAW_APP_NO_SHADOW |
		   BCGP_PLANNER_DRAW_APP_NO_IMAGES |
		   BCGP_PLANNER_DRAW_APP_DEFAULT_BKGND |
		   BCGP_PLANNER_DRAW_APP_DURATION_MULTIDAY |
		   BCGP_PLANNER_DRAW_VIEW_NO_DURATION |
		   BCGP_PLANNER_DRAW_VIEW_WEEK_BAR |
		   BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_BOLD_FC |
		   BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_COMPACT |
		   BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_UPPER |
		   BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_EXTENDED |
		   BCGP_PLANNER_DRAW_VIEW_DAYS_UPDOWN | 
		   BCGP_PLANNER_DRAW_VIEW_DAYS_UPDOWN_VCENTER |
		   BCGP_PLANNER_DRAW_VIEW_HEADER_BOLD_CUR |
		   BCGP_PLANNER_DRAW_VIEW_HEADER_UPPER |
		   BCGP_PLANNER_DRAW_VIEW_HEADER_EXTENDED |
		   BCGP_PLANNER_DRAW_VIEW_TIMELINE_FULL_WIDTH |
		   BCGP_PLANNER_DRAW_VIEW_TIME_BAR_NO_MINUTS;
}
//*******************************************************************************
int CBCGPVisualManagerVS2012::GetPlannerRowExtraHeight () const
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetPlannerRowExtraHeight ();
	}

	return 3;
}

#endif // BCGP_EXCLUDE_PLANNER

#ifndef BCGP_EXCLUDE_RIBBON

void CBCGPVisualManagerVS2012::OnDrawRibbonProgressBar (CDC* pDC, 
														CBCGPRibbonProgressBar* pProgress, 
														CRect rectProgress, CRect rectChunk,
														BOOL bInfiniteMode)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonProgressBar (pDC, pProgress, 
			rectProgress, rectChunk, bInfiniteMode);
		return;
	}
	
	ASSERT_VALID (pDC);

	BOOL bIsOnStatusBar = (DYNAMIC_DOWNCAST(CBCGPRibbonStatusBar, pProgress->GetParentRibbonBar()) != NULL);
	COLORREF clrBorder = bIsOnStatusBar ? globalData.clrBarHilite : globalData.clrBtnDkShadow;

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);

		if (bIsOnStatusBar || IsDarkTheme ())
		{
			dm.DrawRect (rectProgress, m_clrAccentLight, (COLORREF)-1);
		}

		if (!rectChunk.IsRectEmpty ())
		{
			rectChunk.DeflateRect(1, 1);
			dm.DrawRect (rectChunk, m_clrAccent, (COLORREF)-1);
		}

		dm.DrawRect (rectProgress, (COLORREF)-1, clrBorder);
	}
	else
	{
		if (bIsOnStatusBar || IsDarkTheme ())
		{
			pDC->FillRect(rectProgress, &m_brAccentLight);
		}

		if (!rectChunk.IsRectEmpty ())
		{
			rectChunk.DeflateRect(1, 1);
			pDC->FillRect (rectChunk, &m_brAccent);
		}

		pDC->Draw3dRect(rectProgress, clrBorder, clrBorder);
	}
}
//****************************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonCategory (
		CDC* pDC, 
		CBCGPRibbonCategory* pCategory,
		CRect rectCategory)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonCategory (pDC, pCategory, rectCategory);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pCategory);

	pDC->FillRect (rectCategory, &m_brRibbonCategoryFill);

	CRect rectActiveTab = pCategory->GetTabRect ();

	if (pCategory->GetParentMenuBar() != NULL && rectActiveTab.IsRectEmpty())
	{
		rectActiveTab = pCategory->GetOriginalTabRect();
	}

	CPen* pOldPen = pDC->SelectObject (&m_penSeparator);
	ASSERT (pOldPen != NULL);

	pDC->MoveTo (rectCategory.left, rectCategory.top);
	pDC->LineTo (rectActiveTab.left + 1, rectCategory.top);

	pDC->MoveTo (rectActiveTab.right - 2, rectCategory.top);
	pDC->LineTo (rectCategory.right - 1, rectCategory.top);
	pDC->LineTo (rectCategory.right - 1, rectCategory.bottom - 1);
	pDC->LineTo (rectCategory.left, rectCategory.bottom - 1);
	pDC->LineTo (rectCategory.left, rectCategory.top);

	pDC->SelectObject (pOldPen);
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnDrawRibbonPanel(
													 CDC* pDC,
													 CBCGPRibbonPanel* pPanel, 
													 CRect rectPanel,
													 CRect rectCaption)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnDrawRibbonPanel (pDC, pPanel, rectPanel, rectCaption);
	}

	if (pPanel->IsBackstageView())
	{
		rectPanel = ((CBCGPRibbonMainPanel*)pPanel)->GetCommandsFrame ();
		return OnFillRibbonBackstageLeftPane(pDC, rectPanel);
	}

	CPen* pOldPen = pDC->SelectObject (&m_penSeparator);
	ASSERT (pOldPen != NULL);
	
	if (pPanel->IsOnDialogBar())
	{
		pDC->MoveTo(rectPanel.left, rectPanel.bottom - 1);
		pDC->LineTo(rectPanel.right, rectPanel.bottom - 1);
	}
	else
	{
		pDC->MoveTo (rectPanel.right - 1, rectPanel.top);
		pDC->LineTo (rectPanel.right - 1, rectPanel.bottom);
	}
	
	pDC->SelectObject (pOldPen);

	return globalData.clrBarText;
}
//*****************************************************************************
BOOL CBCGPVisualManagerVS2012::IsSimplifiedBackstageViewIcon(CBCGPRibbonButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::IsSimplifiedBackstageViewIcon(pButton);
	}

	return CBCGPDrawManager::IsLightColor(OnGetRibbonBackstageLeftPaneTextColor());
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillRibbonBackstageLeftPane(CDC* pDC, CRect rectPanel)
{
	ASSERT_VALID(pDC);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnFillRibbonBackstageLeftPane(pDC, rectPanel);
	}

	pDC->FillRect(rectPanel, &m_brAccent);
	return OnGetRibbonBackstageLeftPaneTextColor();
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnGetRibbonBackstageLeftPaneTextColor() 
{ 
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnGetRibbonBackstageLeftPaneTextColor();
	}

	return RGB(255, 255, 255); 
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetRibbonBackstageCommandTextColor(CBCGPRibbonButton* pButton)
{
	ASSERT_VALID(pButton);
	return pButton->IsDisabled() ? RGB(192, 192, 192) : RGB(255, 255, 255);
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillRibbonPanelCaption (
										   CDC* pDC,
										   CBCGPRibbonPanel* pPanel, 
										   CRect rectCaption)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnFillRibbonPanelCaption(pDC, pPanel, rectCaption);
	}

	return globalData.clrBarText;
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnDrawRibbonCategoryTab (
					CDC* pDC, 
					CBCGPRibbonTab* pTab,
					BOOL bIsActive)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnDrawRibbonCategoryTab (pDC, 
											pTab, bIsActive);
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pTab);
	
	CBCGPRibbonCategory* pCategory = pTab->GetParentCategory ();
	ASSERT_VALID (pCategory);
	CBCGPRibbonBar* pBar = pCategory->GetParentRibbonBar ();
	ASSERT_VALID (pBar);
	
	bIsActive = bIsActive && 
		((pBar->GetHideFlags () & BCGPRIBBONBAR_HIDE_ELEMENTS) == 0 || pTab->GetDroppedDown () != NULL);
	
	const BOOL bIsFocused	= pTab->IsFocused () && (pBar->GetHideFlags () & BCGPRIBBONBAR_HIDE_ELEMENTS);
	const BOOL bIsHighlighted = (pTab->IsHighlighted () || bIsFocused) && !pTab->IsDroppedDown ();
	
	CRect rectTab = pTab->GetRect ();
	
	rectTab.top += 3;
	
	const int nTrancateRatio =
		pTab->GetParentCategory ()->GetParentRibbonBar ()->GetTabTrancateRatio ();
	
	if (nTrancateRatio > 0)
	{
		const int nPercent = m_Style == VS2012_Dark ?
			min (150, 50 + nTrancateRatio / 2) :
			max (10, 100 - nTrancateRatio / 2);
		
		COLORREF color = CBCGPDrawManager::PixelAlpha (
			m_Style == VS2012_Dark ? globalData.clrBarText : globalData.clrBarFace, nPercent);
		
		CPen pen (PS_SOLID, 1, color);
		CPen* pOldPen = pDC->SelectObject (&pen);
		
		pDC->MoveTo (rectTab.right - 1, rectTab.top);
		pDC->LineTo (rectTab.right - 1, rectTab.bottom);
		
		pDC->SelectObject (pOldPen);
	}
	
	if (!bIsActive && !pTab->IsFocused ())
	{
		return bIsHighlighted ? m_clrAccentText : globalData.clrBarText;
	}
	
	CPen* pOldPen = pDC->SelectObject(&m_penSeparator);
	ASSERT(pOldPen != NULL);

	rectTab.right -= 2;

	COLORREF clrTab = m_clrAccentText;
	COLORREF clrCurr = RibbonCategoryColorToRGB (pTab->GetParentCategory ()->GetTabColor ());

	if (clrCurr != (COLORREF)-1)
	{
		clrTab = CBCGPDrawManager::ColorMakeDarker(clrCurr, .4);
	}

	CPen penFocused;

	if (pTab->IsFocused ())
	{
		CRect rectFill = rectTab;
		rectFill.bottom++;

		COLORREF colorBorder = m_clrHighlightDn;

		if (clrCurr != (COLORREF)-1)
		{
			colorBorder = clrTab;

			CBrush brFill(CBCGPDrawManager::ColorMakePale(clrCurr, .9));
			pDC->FillRect(rectFill, &brFill);
		}
		else
		{
			OnFillHighlightedArea (pDC, rectFill, &m_brHighlight, NULL);
		}

		penFocused.CreatePen(PS_SOLID, 1, colorBorder);
		pDC->SelectObject(&penFocused);
	}
	else
	{
		pDC->FillRect (rectTab, &m_brRibbonCategoryFill);
	}

	pDC->MoveTo(rectTab.left, rectTab.bottom);
	pDC->LineTo(rectTab.left, rectTab.top);
	pDC->LineTo(rectTab.right, rectTab.top);
	pDC->LineTo(rectTab.right, rectTab.bottom);

	pDC->SelectObject (pOldPen);
	
	return clrTab;
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetRibbonEditBackgroundColor (
															   CBCGPRibbonEditCtrl* pEdit,
															   BOOL bIsHighlighted,
															   BOOL bIsPaneHighlighted,
															   BOOL bIsDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetRibbonEditBackgroundColor(pEdit, bIsHighlighted, bIsPaneHighlighted, bIsDisabled);
	}

	BOOL bIsReadOnly = pEdit->GetSafeHwnd() != NULL && (pEdit->GetStyle() & ES_READONLY) == ES_READONLY;
	return bIsDisabled ? m_clrComboDisabled : bIsReadOnly ? m_clrRibbonCategoryFill : m_clrCombo;
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetRibbonEditTextColor(
														   CBCGPRibbonEditCtrl* pEdit,
														   BOOL bIsHighlighted,
														   BOOL bIsDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || (!IsDarkTheme() && !bIsDisabled))
	{
		return CBCGPVisualManagerVS2010::GetRibbonEditTextColor(pEdit, bIsHighlighted, bIsDisabled);
	}

	return bIsDisabled ? m_clrTextDisabled : globalData.clrBarText;
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonMainButton (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonMainButton(pDC, pButton);
		return;
	}

	BOOL bIsHighlighted = pButton->IsHighlighted () || pButton->IsFocused ();
	BOOL bIsPressed = pButton->IsPressed () || pButton->IsDroppedDown ();

	if (pButton->IsDroppedDown ())
	{
		bIsPressed = TRUE;
		bIsHighlighted = TRUE;
	}

	CRect rect = pButton->GetRect ();

	rect.right -= 2;
	rect.top += 2;

	if (pButton->IsDisabled())
	{
		pDC->FillSolidRect(rect, globalData.clrBarShadow);
		return;
	}

	if (bIsHighlighted || IsMainRibbonButtonColored())
	{
		pDC->FillRect(rect, bIsHighlighted ? &m_brAccentHighlight : &m_brAccent);
	}

	if (pButton->IsFocused())
	{
		rect.DeflateRect(1, 1);

		if (pButton->GetTopLevelRibbonBar() != NULL && pButton->GetTopLevelRibbonBar()->GetHideFlags() == 0)
		{
			rect.bottom++;
		}

		pDC->Draw3dRect(rect, m_clrAccentLight, m_clrAccentLight);
	}
}
//***********************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnDrawRibbonCategoryCaption (
					CDC* pDC, 
					CBCGPRibbonContextCaption* pContextCaption)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pContextCaption);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnDrawRibbonCategoryCaption (
					pDC, pContextCaption);
	}

	COLORREF clr = RibbonCategoryColorToRGB (pContextCaption->GetColor ());
	CRect rect = pContextCaption->GetRect ();
	
	if (clr == (COLORREF)-1)
	{
		return globalData.clrBarText;
	}

	COLORREF clrTopLine = IsDarkTheme() ? CBCGPDrawManager::ColorMakeDarker(clr) : CBCGPDrawManager::ColorMakeLighter(clr);
	COLORREF clrFill = (m_Style == VS2012_LightBlue || m_Style == VS2012_Dark) ? CBCGPDrawManager::SmartMixColors(clr, globalData.clrBarFace, 1.0, 3, 1) : CBCGPDrawManager::ColorMakePale(clr);

	CBCGPRibbonBar* pRibbon = pContextCaption->GetParentRibbonBar();
	ASSERT_VALID(pRibbon);

	rect.bottom += pRibbon->GetTabsHeight();

	pDC->FillSolidRect(rect, clrFill);

	CRect rectTop = rect;
	rectTop.bottom = rectTop.top + globalUtils.ScaleByDPI(4);

	pDC->FillSolidRect(rectTop, clrTopLine);

	return CBCGPDrawManager::ColorMakeDarker(clr, .4);
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnDrawRibbonTabsFrame (
					CDC* pDC, 
					CBCGPRibbonBar* pWndRibbonBar, 
					CRect rectTab)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnDrawRibbonTabsFrame (pDC, 
											pWndRibbonBar, rectTab);
	}

	return m_clrRibbonTabs;
}
//**************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetRibbonBackstageTextColor()
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetRibbonBackstageTextColor();
	}

	return globalData.clrBarText;
}
//**************************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetRibbonBackstageInfoTextColor()
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetRibbonBackstageInfoTextColor();
	}

	return m_clrAccent;
}
//**************************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnDrawRibbonStatusBarPane(CDC* pDC, CBCGPRibbonStatusBar* pBar, CBCGPRibbonStatusBarPane* pPane)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pPane->IsDisabled())
	{
		return CBCGPVisualManagerVS2010::OnDrawRibbonStatusBarPane(pDC, pBar, pPane);
	}

	CRect rect = pPane->GetRect ();

	COLORREF clrFill = m_clrState != (COLORREF)-1 ? m_clrState : m_clrAccent;

	if (pPane->IsHighlighted ())
	{
		if (pPane->IsPressed ())
		{
			clrFill = m_clrState != (COLORREF)-1 ? CBCGPDrawManager::ColorMakeDarker(m_clrState) : m_clrAccentText;
			pDC->FillSolidRect(rect, clrFill);
		}
		else
		{
			clrFill = m_clrState != (COLORREF)-1 ? CBCGPDrawManager::ColorMakeLighter(m_clrState, .2) : m_clrAccentHilight;
			pDC->FillSolidRect(rect, clrFill);
		}

		return CBCGPDrawManager::GetContrastColor(clrFill);
	}
	else if (pPane->IsChecked())
	{
		clrFill = m_clrState != (COLORREF)-1 ? CBCGPDrawManager::ColorMakeDarker(m_clrState) : m_clrAccentText;
		pDC->FillSolidRect(rect, clrFill);

		return CBCGPDrawManager::GetContrastColor(clrFill);
	}

	return IsRibbonStatusBarDark() ? RGB(255, 255, 255) : RGB(0, 0, 0);
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonButtonBorder (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonButtonBorder (pDC, pButton);
		return;
	}

	if (pButton->IsKindOf (RUNTIME_CLASS(CBCGPRibbonEdit)))
	{
		CRect rect (pButton->GetRect ());

		COLORREF colorBorder = m_clrEditBoxBorder;

		if (pButton->IsDisabled ())
		{
			colorBorder = m_clrTextDisabled;
		}
		else if (pButton->IsHighlighted () || pButton->IsDroppedDown () || pButton->IsFocused ())
		{
			colorBorder = m_clrBorderActive;
		}

		rect.left = pButton->GetCommandRect ().left;

		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect (rect, (COLORREF)-1, colorBorder);
		}
		else
		{
			pDC->Draw3dRect (rect, colorBorder, colorBorder);
		}
	}
	else if ((m_bCheckedRibbonButtonFrame || pButton->IsHighlighted()) && pButton->IsChecked() && !pButton->IsDisabled())
	{
		CRect rect(pButton->GetRect());
		COLORREF colorBorder = pButton->IsHighlighted() ? m_clrHighlightDn : m_clrAccent;
		
		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect (rect, (COLORREF)-1, colorBorder);
		}
		else
		{
			pDC->Draw3dRect (rect, colorBorder, colorBorder);
		}
	}
}
//****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonDefaultPaneButton (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonDefaultPaneButton (pDC, pButton);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->IsQATMode () || pButton->IsSearchResultMode ())
	{
		OnFillRibbonButton (pDC, pButton);
	}
	else
	{
		CRect rect = pButton->GetRect();
		CRect rectImage = rect;

		rect.right--;
		
		COLORREF clrFill = (m_Style == VS2012_Dark) ? globalData.clrBarLight : globalData.clrBarShadow;

		if (pButton->IsDroppedDown())
		{
			pDC->FillSolidRect(rect, clrFill);
		}
		else if (pButton->IsHighlighted() || pButton->IsFocused())
		{
			pDC->FillSolidRect(rect, CBCGPDrawManager::PixelAlpha(clrFill, 103));
		}

		CSize sizeImage = pButton->GetImageSize (CBCGPRibbonButton::RibbonImageSmall);

		rectImage.top += 10;
		rectImage.bottom = rectImage.top + sizeImage.cy;

		rectImage.left = rectImage.CenterPoint().x - sizeImage.cx / 2;
		rectImage.right = rectImage.left + sizeImage.cx;

		rectImage.InflateRect(5, 5);

		pDC->FillRect(rectImage, &globalData.brBarFace);

		COLORREF clrBorder = (m_Style == VS2012_Dark && (pButton->IsHighlighted() || pButton->IsDroppedDown())) ? m_clrAccentText : m_clrMenuBorder;
		pDC->Draw3dRect(rectImage, clrBorder, clrBorder);
	}

	OnDrawRibbonDefaultPaneButtonContext (pDC, pButton);
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnDrawDefaultRibbonImage (
					CDC* pDC, CRect rectImage,
					BOOL bIsDisabled,
					BOOL bIsPressed,
					BOOL bIsHighlighted)
{
	ASSERT_VALID (pDC);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawDefaultRibbonImage (pDC, rectImage, bIsDisabled, bIsPressed, bIsHighlighted);
		return;
	}

	CRect rectBullet (rectImage.CenterPoint (), CSize (1, 1));

	int nRadius = globalUtils.ScaleByDPI(6);
	rectBullet.InflateRect(nRadius, nRadius);

	CBCGPDrawManager dm (*pDC);

	dm.DrawEllipse (rectBullet,
		bIsDisabled ? globalData.clrGrayedText : RGB (207, 224, 219),
		bIsDisabled ? globalData.clrBtnShadow : RGB (126, 172, 157));

	rectBullet.DeflateRect(1, 1);

	dm.DrawEllipse (rectBullet,
		(COLORREF)-1,
		bIsDisabled ? RGB(192, 192, 192) : RGB (244, 248, 246));
}
//****************************************************************************
BOOL CBCGPVisualManagerVS2012::IsRibbonStatusBarDark()
{	
	if (m_clrState != (COLORREF)-1)
	{
		if (m_clrState == 0)
		{
			return TRUE;
		}
		
		double H;
		double S;
		double L;
		
		CBCGPDrawManager::RGBtoHSL(m_clrState, &H, &S, &L);
		return L < 0.7;
	}

	return TRUE;	
}
//****************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillRibbonStatusBarButton(CDC* pDC, CBCGPRibbonButton* pButton)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	CRect rect = pButton->GetRect ();

	COLORREF clrFill = m_clrState != (COLORREF)-1 ? m_clrState : m_clrAccent;
	
	if (pButton->IsHighlighted ())
	{
		if (pButton->IsPressed ())
		{
			clrFill = m_clrState != (COLORREF)-1 ? CBCGPDrawManager::ColorMakeDarker(m_clrState) : m_clrAccentText;
			pDC->FillSolidRect(rect, clrFill);
		}
		else
		{
			clrFill = m_clrState != (COLORREF)-1 ? CBCGPDrawManager::ColorMakeLighter(m_clrState, .2) : m_clrAccentHilight;
			pDC->FillSolidRect(rect, clrFill);
		}

		return CBCGPDrawManager::GetContrastColor(clrFill);
	}
	else if (pButton->IsChecked())
	{
		clrFill = m_clrState != (COLORREF)-1 ? CBCGPDrawManager::ColorMakeDarker(m_clrState) : m_clrAccentText;
		pDC->FillSolidRect(rect, clrFill);

		return CBCGPDrawManager::GetContrastColor(clrFill);
	}
	
	return IsRibbonStatusBarDark() ? RGB(255, 255, 255) : RGB(0, 0, 0);
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonMainButtonLabel(CDC* pDC, CBCGPRibbonButton* pButton, CRect rectText, const CString strText)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonMainButtonLabel(pDC, pButton, rectText, strText);
		return;
	}

	CString strLabel = strText;
	if (IsTopLevelMenuItemUpperCase())
	{
		strLabel.MakeUpper();
	}
	
	pButton->DoDrawText(pDC, strLabel, rectText, DT_SINGLELINE | DT_CENTER | DT_VCENTER, GetRibbonMainButtonTextColor(pButton));
}
//*****************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetRibbonMainButtonTextColor(CBCGPRibbonButton* pButton)
{
	ASSERT_VALID(pButton);
	return pButton->IsDisabled() ? RGB(128, 128, 128) : RGB(253, 253, 253);
}
//****************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnFillRibbonButton (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnFillRibbonButton(pDC, pButton);
	}

	if (pButton->IsDefaultPanelButton () && !pButton->IsQATMode () && !pButton->IsSearchResultMode ())
	{
		return (COLORREF)-1;
	}

	CRect rect = pButton->GetRect ();

	const BOOL bIsMenuMode = pButton->IsMenuMode () && !pButton->IsPaletteIcon ();

	const BOOL bIsHighlighted = 
		((pButton->IsHighlighted () || pButton->IsDroppedDown ()) && 
		!pButton->IsDisabled ()) || pButton->IsFocused ();

	if (pButton->IsKindOf (RUNTIME_CLASS (CBCGPRibbonEdit)))
	{
		BOOL bIsReadOnly = ((CBCGPRibbonEdit*)pButton)->IsReadOnly();

		COLORREF clrBorder = globalData.clrBarShadow;
		COLORREF clrFill = pButton->IsDisabled() ? m_clrComboDisabled : bIsReadOnly ? m_clrRibbonCategoryFill : m_clrCombo;

		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect (rect, clrFill, clrBorder);
		}
		else
		{
			CBrush brFill(clrFill);
			pDC->FillRect (rect, &brFill);

			pDC->Draw3dRect (rect, clrBorder, clrBorder);
		}
		
		return globalData.clrBarText;
	}

	if (pButton->IsBackstageViewMode () && !pButton->IsQATMode())
	{
		if (bIsHighlighted)
		{
			pDC->FillSolidRect(rect, IsDarkTheme() ? CBCGPDrawManager::ColorMakeDarker(m_clrAccentHilight) : m_clrAccentText);
		}
		else if (!pButton->IsDisabled () && pButton->GetBackstageAttachedView() != NULL && pButton->IsChecked())
		{
			pDC->FillSolidRect(rect, m_clrAccentHilight);
		}

		return GetRibbonBackstageCommandTextColor(pButton);
	}

	if (!pButton->IsChecked () && !bIsHighlighted)
	{
		if (pButton->IsStatusBarMode())
		{
			return IsRibbonStatusBarDark() ? RGB(255, 255, 255) : (COLORREF)-1;
		}

		CBCGPRibbonCaptionButton* pCaptionButton = DYNAMIC_DOWNCAST(CBCGPRibbonCaptionButton, pButton);
		if (pCaptionButton != NULL && pCaptionButton->IsOnBackStageTopMenu())
		{
			return RGB(255, 255, 255);
		}

		return (COLORREF)-1;
	}

	if (pButton->IsChecked () && bIsMenuMode && !bIsHighlighted)
	{
		return (COLORREF)-1;
	}

	CRect rectMenu = pButton->GetMenuRect ();
	
	CRect rectCommand (0, 0, 0, 0);
	if (!rectMenu.IsRectEmpty ())
	{
		rectCommand = pButton->GetCommandRect ();

		if (pButton->GetLocationInGroup () != CBCGPBaseRibbonElement::RibbonElementNotInGroup)
		{
			rectMenu.DeflateRect (0, 1, 1, 1);
			rectCommand.DeflateRect (1, 1, 0, 1);
		}
	}

	COLORREF clrText = (COLORREF)-1;

	if (!rectMenu.IsRectEmpty () && bIsHighlighted)
	{
		if (pButton->IsCommandAreaHighlighted ())
		{
			OnFillHighlightedArea (pDC, rectCommand, pButton->IsPressed () ? &m_brHighlightDn : &m_brHighlight, NULL);

			if (pButton->IsChecked())
			{
				OnFillHighlightedArea(pDC, rectMenu, &m_brHighlightDn, NULL);
			}

			pDC->Draw3dRect(rectMenu, m_clrHighlightDn, m_clrHighlightDn);
		}
		else
		{
			if ((pButton->IsPressed () || pButton->IsDroppedDown ()))
			{
				OnFillHighlightedArea (pDC, rect, &m_brHighlightDn, NULL);
				return m_clrHighlighDownText;
			}
			else
			{
				if (pButton->IsChecked())
				{
					OnFillHighlightedArea (pDC, rectCommand, &m_brHighlightDn, NULL);
				}

				OnFillHighlightedArea (pDC, rectMenu, &m_brHighlight, NULL);
				pDC->Draw3dRect(rectCommand, m_clrHighlightDn, m_clrHighlightDn);
			}
		}
	}
	else if (pButton->IsStatusBarMode())
	{
		return OnFillRibbonStatusBarButton(pDC, pButton);
	}
	else
	{
		if (pButton->IsDisabled() && pButton->IsChecked())
		{
			pDC->Draw3dRect (rect, m_clrHighlightDn, m_clrHighlightDn);
			return (COLORREF)-1;
		}

		BOOL bIsHighlightDown = (pButton->IsPressed () || pButton->IsDroppedDown ()) && !bIsMenuMode;
		CBrush* pBrush = bIsHighlightDown ? &m_brHighlightDn : &m_brHighlight;

		CRect rectFill = rect;

		if (pButton->IsPaletteIcon())
		{
			rectFill.DeflateRect(1, 1);
		}
		
		if (pButton->IsChecked () && !bIsMenuMode)
		{
			pBrush = bIsHighlighted ? &m_brHighlightChecked : (pButton->IsPaletteIcon() ? &m_brHighlight : &m_brHighlightDn);

			if (pBrush == &m_brHighlightDn && rectCommand.IsRectEmpty())
			{
				clrText = m_clrHighlighDownText;
			}
		}

		OnFillHighlightedArea (pDC, rectFill, pBrush, NULL);

		if (bIsHighlightDown && !pButton->IsChecked())
		{
			clrText = m_clrHighlighDownText;
		}
	}

	return clrText;
}
//********************************************************************************
void CBCGPVisualManagerVS2012::GetRibbonSliderColors (CBCGPRibbonSlider* pSlider, 
					BOOL bIsHighlighted, 
					BOOL bIsPressed,
					BOOL bIsDisabled,
					COLORREF& clrLine,
					COLORREF& clrFill)
{
	BOOL bIsRibbonStatusBar = DYNAMIC_DOWNCAST(CBCGPRibbonStatusBar, pSlider->GetParentRibbonBar()) != NULL && IsRibbonStatusBarDark();

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || bIsDisabled || bIsHighlighted || bIsPressed)
	{
		CBCGPVisualManagerVS2010::GetRibbonSliderColors(pSlider, bIsHighlighted, bIsPressed, bIsDisabled, clrLine, clrFill);
		if ((bIsHighlighted || bIsPressed) && !bIsRibbonStatusBar)
		{
			clrLine = m_clrAccentText;
		}
		return;
	}

	clrFill = globalData.clrBarFace;

	if (bIsRibbonStatusBar)
	{
		clrLine = RGB(255, 255, 255);
	}
	else
	{
		clrLine = globalData.clrBarText;
	}
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonSliderZoomButton (
			CDC* pDC, CBCGPRibbonSlider* pSlider, 
			CRect rect, BOOL bIsZoomOut, 
			BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonSliderZoomButton (
			pDC, pSlider, rect, bIsZoomOut, bIsHighlighted, bIsPressed, bIsDisabled);
		return;
	}

	ASSERT_VALID (pDC);

	COLORREF clrLine = (COLORREF)-1;
	COLORREF clrFill = (COLORREF)-1;

	GetRibbonSliderColors(pSlider, bIsHighlighted, bIsPressed, bIsDisabled, clrLine, clrFill);

	CPoint ptCenter = rect.CenterPoint ();
	int nRadius = globalUtils.ScaleByDPI(7);

	// Draw +/- sign:
	CRect rectSign (CPoint (ptCenter.x - nRadius / 2, ptCenter.y - nRadius / 2), CSize (nRadius, nRadius));

	CBCGPDrawManager dm(*pDC);

	dm.DrawLine(rectSign.left, ptCenter.y, rectSign.right, ptCenter.y, clrLine);
	dm.DrawLine(rectSign.left, ptCenter.y + 1, rectSign.right, ptCenter.y + 1, clrLine);

	if (!bIsZoomOut)
	{
		dm.DrawLine(ptCenter.x, rectSign.top, ptCenter.x, rectSign.bottom, clrLine);
		dm.DrawLine(ptCenter.x + 1, rectSign.top, ptCenter.x + 1, rectSign.bottom, clrLine);
	}
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonSliderChannel (
			CDC* pDC, CBCGPRibbonSlider* pSlider, 
			CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonSliderChannel (
			pDC, pSlider, rect);
		return;
	}

	ASSERT_VALID (pDC);

	BOOL bIsVert = FALSE;

	if (pSlider != NULL)
	{
		ASSERT_VALID (pSlider);
		bIsVert = pSlider->IsVert ();
	}

	OnDrawSliderChannel (pDC, NULL, bIsVert, rect, CBCGPToolBarImages::m_bIsDrawOnGlass);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonSliderThumb (
			CDC* pDC, CBCGPRibbonSlider* pSlider, 
			CRect rect, BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonSliderThumb (
			pDC, pSlider, rect, bIsHighlighted, bIsPressed, bIsDisabled);
		return;
	}

	ASSERT_VALID (pDC);

	BOOL bIsVert = FALSE;
	BOOL bLeftTop = FALSE;
	BOOL bRightBottom = FALSE;

	if (pSlider != NULL)
	{
		ASSERT_VALID (pSlider);

		bIsVert = pSlider->IsVert ();
		bLeftTop = pSlider->IsThumbLeftTop ();
		bRightBottom = pSlider->IsThumbRightBottom ();
	}

	OnDrawSliderThumb (pDC, NULL, rect, bIsHighlighted, bIsPressed, bIsDisabled,
				bIsVert, bLeftTop, bRightBottom, CBCGPToolBarImages::m_bIsDrawOnGlass);
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::OnDrawRibbonButtonsGroup (
					CDC* pDC, CBCGPRibbonButtonsGroup* pGroup,
					CRect rect)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pGroup);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnDrawRibbonButtonsGroup (pDC, pGroup, rect);
	}

	return (COLORREF)-1;
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonGroupSeparator (CDC* pDC, CRect rectSeparator)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonGroupSeparator(pDC, rectSeparator);
		return;
	}

	ASSERT_VALID (pDC);

	CPen* pOldPen = pDC->SelectObject (&m_penSeparator);
	ASSERT (pOldPen != NULL);
	
	pDC->MoveTo (rectSeparator.CenterPoint().x, rectSeparator.top);
	pDC->LineTo (rectSeparator.CenterPoint().x, rectSeparator.bottom);
	
	pDC->SelectObject (pOldPen);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnFillRibbonPanelMenuBarInCtrlMode(CDC* pDC, CBCGPRibbonPanelMenuBar* pBar, CRect rectClient)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnFillRibbonPanelMenuBarInCtrlMode(pDC, pBar, rectClient);
		return;
	}

	ASSERT_VALID(pDC);
	pDC->FillRect(rectClient, &m_brControl);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonMinimizeButtonImage(CDC* pDC, CBCGPRibbonMinimizeButton* pButton, BOOL bRibbonIsMinimized)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonMinimizeButtonImage(pDC, pButton, bRibbonIsMinimized);
		return;
	}
	
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);
	
	CBCGPMenuImages::IMAGE_STATE state = pButton->IsHighlighted() ? CBCGPMenuImages::ImageBlack2 : CBCGPMenuImages::ImageBlack;
	
	if (pButton->IsDisabled())
	{
		state = CBCGPMenuImages::ImageLtGray;
	}
	else if (pButton->IsPressed() && pButton->IsHighlighted())
	{
		state = CBCGPMenuImages::ImageWhite;
	}
	
	CBCGPMenuImages::Draw(pDC, 
		bRibbonIsMinimized ? CBCGPMenuImages::IdMinimizeRibbon : CBCGPMenuImages::IdMaximizeRibbon, pButton->GetRect(), state);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonQATSeparator (CDC* pDC, CBCGPRibbonSeparator* pSeparator, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonQATSeparator(pDC, pSeparator, rect);
		return;
	}

	ASSERT_VALID (pDC);
	
	int x = rect.CenterPoint ().x;
	int y1 = rect.top;
	int y2 = rect.bottom - 1;
	
	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawLine (x, y1, x, y2, m_clrSeparator);
	}
	else
	{
		CPen* pOldPen = pDC->SelectObject (&m_penSeparator);
		ASSERT (pOldPen != NULL);
		
		pDC->MoveTo (x, y1);
		pDC->LineTo (x, y2);
		
		pDC->SelectObject (pOldPen);
	}
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetRibbonQATTextColor (BOOL bDisabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetRibbonQATTextColor (bDisabled);
	}
	
	return bDisabled ? m_clrTextDisabled : globalData.clrBarText;
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonMenuArrow(CDC* pDC, CBCGPRibbonButton* pButton, UINT idIn, const CRect& rectMenuArrow)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonMenuArrow(pDC, pButton, idIn, rectMenuArrow);
		return;
	}

	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	CBCGPMenuImages::IMAGES_IDS id = (CBCGPMenuImages::IMAGES_IDS)idIn;
	CBCGPMenuImages::IMAGE_STATE state = CBCGPMenuImages::ImageBlack;

	if (pButton->IsDisabled())
	{
		state = IsDarkTheme() ? CBCGPMenuImages::ImageDkGray : CBCGPMenuImages::ImageLtGray;
	}
	else if (pButton->IsDrawSimplifiedIcon())
	{
		state = CBCGPMenuImages::ImageWhite;
	}
	else if ((pButton->IsMenuAreaHighlighted() && pButton->IsPressed()) || pButton->IsDroppedDown())
	{
		if (!pButton->IsDefaultPanelButton())
		{
			CBCGPMenuImages::DrawByColor(pDC, id, rectMenuArrow, m_clrHighlightDn);
			return;
		}
	}

	CBCGPMenuImages::Draw (pDC, id, rectMenuArrow, state);
}
//****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonLaunchButton (
					CDC* pDC,
					CBCGPRibbonLaunchButton* pButton,
					CBCGPRibbonPanel* pPanel)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonLaunchButton(pDC, pButton, pPanel);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);
	ASSERT_VALID (pPanel);

	OnFillRibbonButton (pDC, pButton);

	CBCGPMenuImages::IMAGE_STATE state = pButton->IsHighlighted() ? CBCGPMenuImages::ImageBlack2 : CBCGPMenuImages::ImageBlack;
	CRect rect = pButton->GetRect ();
	CBCGPMenuImages::IMAGES_IDS id = CBCGPMenuImages::IdLaunchArrow;

	BOOL bIsReady = FALSE;
	
	if (pButton->IsDisabled())
	{
		state = IsDarkTheme() ? CBCGPMenuImages::ImageDkGray : CBCGPMenuImages::ImageLtGray;
	}
	else if (pButton->IsHighlighted() && pButton->IsPressed() || pButton->IsDroppedDown())
	{
		CBCGPMenuImages::DrawByColor(pDC, id, rect, m_clrHighlightDn);
		bIsReady = TRUE;
	}

	if (!bIsReady)
	{
		CBCGPMenuImages::Draw (pDC, id, rect, state);
	}

	OnDrawRibbonButtonBorder (pDC, pButton);
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetRibbonHyperlinkTextColor (CBCGPRibbonHyperlink* pHyperLink)
{
	ASSERT_VALID (pHyperLink);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pHyperLink->IsDisabled () || !IsRibbonStatusBarDark())
	{
		return CBCGPVisualManagerVS2010::GetRibbonHyperlinkTextColor (pHyperLink);
	}

	if (DYNAMIC_DOWNCAST(CBCGPRibbonStatusBar, pHyperLink->GetParentRibbonBar()) != NULL)
	{
		if (m_clrState != (COLORREF)-1)
		{
			COLORREF clrText = CBCGPDrawManager::ColorMakePale(m_clrState);
			return pHyperLink->IsHighlighted () ? CBCGPDrawManager::ColorMakeDarker(clrText) : clrText;
		}

		if (IsDarkTheme())
		{
			return pHyperLink->IsHighlighted() ? m_clrHighlight : m_clrAccentText;
		}

		return pHyperLink->IsHighlighted() ? m_clrAccentLight : m_clrHighlight;
	}

	if (IsDarkTheme())
	{
		return pHyperLink->IsHighlighted() ? m_clrAccent : m_clrAccentText;
	}

	return pHyperLink->IsHighlighted() ? m_clrAccentText : m_clrAccentHilight;
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonBackStageTopMenu(CDC* pDC, CRect rectBackstageTopArea)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonBackStageTopMenu(pDC, rectBackstageTopArea);
		return;
	}

	OnFillRibbonBackstageLeftPane(pDC, rectBackstageTopArea);
}
//********************************************************************************
int CBCGPVisualManagerVS2012::GetRibbonPopupBorderSize (const CBCGPRibbonPanelMenu* pPopup) const
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pPopup == NULL)
	{
		return CBCGPVisualManagerVS2010::GetRibbonPopupBorderSize(pPopup);
	}

	ASSERT_VALID (pPopup);

	const CBCGPRibbonPanelMenu* pRibbonMenu = DYNAMIC_DOWNCAST (CBCGPRibbonPanelMenu, pPopup);
	if (pRibbonMenu != NULL)
	{
		ASSERT_VALID(pRibbonMenu);

		if (pRibbonMenu->GetCategory () != NULL)
		{
			return 0;
		}
	}
	
	return CBCGPVisualManagerVS2010::GetRibbonPopupBorderSize(pPopup);
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonCategoryScroll (
					CDC* pDC, 
					CBCGPRibbonCategoryScroll* pScroll)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonCategoryScroll (pDC, pScroll);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pScroll);
	
	CRect rect = pScroll->GetRect ();
	rect.bottom--;
	
	OnFillHighlightedArea (pDC, rect, 
		pScroll->IsHighlighted () ? &m_brHighlight : &globalData.brBarFace, NULL);
	
	BOOL bIsLeft = pScroll->IsLeftScroll ();
	if (globalData.m_bIsRTL)
	{
		bIsLeft = !bIsLeft;
	}
	
	CBCGPMenuImages::Draw (pDC,
		pScroll->IsVertical() ? 
			(bIsLeft ? CBCGPMenuImages::IdArowUpLarge : CBCGPMenuImages::IdArowDownLarge) : 
			(bIsLeft ? CBCGPMenuImages::IdArowLeftLarge : CBCGPMenuImages::IdArowRightLarge), 
		rect);

	COLORREF clrBorder = pScroll->IsHighlighted() ? m_clrAccentHilight : m_clrMenuBorder;
	pDC->Draw3dRect (rect, clrBorder, clrBorder);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawRibbonColorPaletteBoxHotBorder (CDC* pDC, CBCGPRibbonColorButton* pColorButton,
																 CBCGPRibbonPaletteIcon* pIcon,
																 CRect rect, BOOL bIsHighlighted, BOOL bIsChecked)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawRibbonColorPaletteBoxHotBorder(pDC, pColorButton, pIcon, rect, bIsHighlighted, bIsChecked);
		return;
	}

	ASSERT_VALID (pDC);
	
	if (bIsHighlighted || bIsChecked)
	{
		COLORREF clrBorder = bIsHighlighted ? RGB(242, 148, 54) : RGB(239, 72, 16);
		pDC->Draw3dRect (rect, clrBorder, clrBorder);

		COLORREF clrIner = RGB(255, 226, 148);

		rect.DeflateRect(1, 1);
		pDC->Draw3dRect (rect, clrIner, clrIner);
	}
}

#endif

//********************************************************************************
BOOL CBCGPVisualManagerVS2012::DrawCheckBox (CDC *pDC, CRect rect, 
										 BOOL bHighlighted, 
										 int nState,
										 BOOL bEnabled,
										 BOOL bPressed)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::DrawCheckBox (pDC, rect, bHighlighted, nState, bEnabled, bPressed);
	}

	rect.DeflateRect(1, 1);

	if (rect.Width() <= 0 || rect.Height() <= 0)
	{
		return TRUE;
	}

	if (rect.Width() > rect.Height())
	{
		rect.left = rect.CenterPoint().x - rect.Height() / 2;
		rect.right = rect.left + rect.Height();
	}
	else if (rect.Width() < rect.Height())
	{
		rect.top = rect.CenterPoint().y - rect.Width() / 2;
		rect.bottom = rect.top + rect.Width();
	}

	if (bPressed)
	{
		pDC->FillRect(rect, &m_brHighlightDn);
	}
	else if (bHighlighted)
	{
		pDC->FillRect(rect, &m_brHighlight);
	}
	else if (!IsDarkTheme ())
	{
		pDC->FillRect(rect, &m_brWhite);
	}
	else if (m_Style != VS2012_Dark)
	{
		pDC->FillSolidRect(rect, globalData.clrBarHilite);
	}

	COLORREF clrBorder = (bHighlighted || bPressed) ? 
		(m_Style == VS2012_LightBlue ? m_clrMenuItemBorder : m_clrHighlightDn) : globalData.clrBarDkShadow;

	if (!bEnabled)
	{
		clrBorder = globalData.clrBarShadow;
	}
	
	pDC->Draw3dRect (&rect, clrBorder, clrBorder);
	
	if (nState == 1)
	{
		if (m_bScaleCheckRadio)
		{
			rect.DeflateRect(rect.Width() / 4, rect.Height() / 4);
			COLORREF clrLine = m_clrMenuImage;

			if (!bEnabled)
			{
				clrLine = globalData.clrBarShadow;
			}
			else if (bPressed)
			{
				if (CBCGPDrawManager::IsDarkColor(m_clrHighlightDn))
				{
					clrLine = RGB(255, 255, 255);
				}
				else
				{
					clrLine = m_clrAccent;
				}
			}
			else if (bHighlighted)
			{
				clrLine = m_clrAccent;
			}

			CBCGPDrawManager dm(*pDC);
			dm.DrawCheckMark(rect, clrLine, (int)(0.5 + m_dblScaleCheckRadio));
		}
		else
		{
			CSize sizeDest(0, 0);

			if (CBCGPMenuImages::Size().cx > rect.Width() || CBCGPMenuImages::Size().cy > rect.Height())
			{
				sizeDest = rect.Size();
			}
			else
			{
				rect.bottom--;
			}

			CBCGPMenuImages::IMAGE_STATE state = CBCGPMenuImages::ImageGray;

			if (!bEnabled)
			{
				state = IsDarkTheme() ? CBCGPMenuImages::ImageDkGray : CBCGPMenuImages::ImageLtGray;
			}
			else if (bPressed)
			{
				if (CBCGPDrawManager::IsDarkColor(m_clrHighlightDn))
				{
					state = CBCGPMenuImages::ImageWhite;
				}
				else
				{
					state = CBCGPMenuImages::ImageBlack2;
				}
			}
			else if (bHighlighted)
			{
				state = CBCGPMenuImages::ImageBlack2;
			}

			CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdCheck, rect, state, sizeDest);
		}
	}
	else if (nState == 2)
	{
		rect.DeflateRect (1, 1);
		
		WORD HatchBits [8] = { 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55 };
		
		CBitmap bmp;
		bmp.CreateBitmap (8, 8, 1, 1, HatchBits);
		
		CBrush br;
		br.CreatePatternBrush (&bmp);

		COLORREF clrTextOld = pDC->SetTextColor(globalData.clrBarFace);
		
		pDC->FillRect (rect, &br);

		pDC->SetTextColor(clrTextOld);
	}

	return TRUE;
}
//*************************************************************************************************
BOOL CBCGPVisualManagerVS2012::DrawRadioButton (CDC *pDC, CRect rect, 
										 BOOL bHighlighted, 
										 BOOL bChecked,								 
										 BOOL bEnabled,
										 BOOL bPressed)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::DrawRadioButton (pDC, rect, bHighlighted, bChecked, bEnabled, bPressed);
	}

	rect.DeflateRect(1, 1);

	if (rect.Width() <= 0 || rect.Height() <= 0)
	{
		return TRUE;
	}

	if (rect.Width() > rect.Height())
	{
		rect.left = rect.CenterPoint().x - rect.Height() / 2;
		rect.right = rect.left + rect.Height();
	}
	else if (rect.Width() < rect.Height())
	{
		rect.top = rect.CenterPoint().y - rect.Width() / 2;
		rect.bottom = rect.top + rect.Width();
	}

	CBCGPDrawManager dm (*pDC);
	
	if (bPressed)
	{
		dm.DrawEllipse (rect, m_clrHighlightDn, (COLORREF)-1);
	}
	else if (bHighlighted || bPressed)
	{
		dm.DrawEllipse (rect, m_clrHighlight, (COLORREF)-1);
	}
	else if (!IsDarkTheme ())
	{
		dm.DrawEllipse (rect, RGB(255, 255, 255), (COLORREF)-1);
	}
	else if (m_Style != VS2012_Dark)
	{
		dm.DrawEllipse (rect, globalData.clrBarHilite, (COLORREF)-1);
	}

	COLORREF clrBorder = (bHighlighted || bPressed) ? 
		(m_Style == VS2012_LightBlue ? m_clrMenuItemBorder : m_clrHighlightDn) : globalData.clrBarDkShadow;

	if (!bEnabled)
	{
		clrBorder = IsDarkTheme() ? CBCGPDrawManager::ColorMakeLighter(globalData.clrBarShadow, .7) : CBCGPDrawManager::ColorMakeDarker(globalData.clrBarShadow);
	}

	dm.DrawEllipse (rect,(COLORREF)-1, clrBorder);
	
	if (bChecked)
	{
		rect.DeflateRect (rect.Width () / 3, rect.Width () / 3);
		
		dm.DrawEllipse (rect,
			bEnabled ? (bHighlighted ? m_clrAccentText : globalData.clrBarText) : (m_Style == VS2012_Dark ? globalData.clrBarShadow : clrBorder),
			(COLORREF)-1);
	}

	return TRUE;
}
//*************************************************************************************************
CSize CBCGPVisualManagerVS2012::GetExpandButtonDefaultSize()
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetExpandButtonDefaultSize();
	}

	CSize size = CBCGPMenuImages::Size();
	size.cx *= 2;
	size.cy *= 2;

	return size;
}
//**********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawGroupExpandCollapse(CDC* pDC, CRect rect, 
									   BOOL bIsExpanded,
									   BOOL bHighlighted, 
									   BOOL bPressed,
									   BOOL bEnabled)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawGroupExpandCollapse(pDC, rect, bIsExpanded, bHighlighted, bPressed, bEnabled);
		return;
	}

	CBCGPDrawManager dm (*pDC);

	COLORREF clrFill = globalData.clrBarFace;
	COLORREF clrBorder = globalData.clrBarDkShadow;

	if (!bEnabled)
	{
		clrBorder = globalData.clrBarShadow;
	}
	else if (bPressed)
	{
		clrFill = m_clrHighlightDn;
	}
	else if (bHighlighted || bPressed)
	{
		clrFill = m_clrHighlight;
		clrBorder = m_Style == VS2012_LightBlue ? m_clrMenuItemBorder : m_clrHighlightDn;
	}
	else if (!IsDarkTheme ())
	{
		clrFill = RGB(255, 255, 255);
	}
	
	dm.DrawEllipse (rect, clrFill, clrBorder);

	CBCGPMenuImages::DrawByColor(pDC, bIsExpanded ? CBCGPMenuImages::IdArowUpLarge : CBCGPMenuImages::IdArowDownLarge, rect, clrFill);
}
//*************************************************************************************************
BOOL CBCGPVisualManagerVS2012::OnDrawCalculatorDisplay (CDC* pDC, CRect rect, const CString& strText, BOOL bMem, CBCGPCalculator* pCalculator)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnDrawCalculatorDisplay (pDC, rect, strText, bMem, pCalculator);
	}

	pDC->FillRect(rect, &m_brHighlight);

	COLORREF clrBorder = globalData.clrBarShadow;
	pDC->Draw3dRect (&rect, clrBorder, clrBorder);

	return TRUE;
}
//*************************************************************************************************
BOOL CBCGPVisualManagerVS2012::OnDrawCalculatorButton (CDC* pDC, 
	CRect rect, CBCGPToolbarButton* pButton, 
	CBCGPVisualManager::BCGBUTTON_STATE state, 
	int cmd,
	CBCGPCalculator* pCalculator)
{
	ASSERT_VALID (pButton);
	ASSERT_VALID (pDC);
	ASSERT_VALID (pCalculator);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pCalculator->IsDialogControl () || 
		m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::OnDrawCalculatorButton (pDC, rect, pButton, state, cmd, pCalculator);
	}

	switch (state)
	{
	case ButtonsIsPressed:
		pDC->FillRect(rect, &m_brHighlightDn);
		break;
		
	case ButtonsIsHighlighted:
		pDC->FillRect(rect, &m_brHighlight);
		break;
		
	default:
		pDC->FillRect(rect, 
			DYNAMIC_DOWNCAST(CBCGPPopupMenu, pCalculator->GetParent()) != NULL ? &globalData.brBarFace : &m_brMenuLight);
		break;
	}

	return TRUE;
}
//********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetCalculatorDisplayTextColor() 
{ 
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetCalculatorDisplayTextColor();
	}

	return m_clrAccentText; 
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetCalculatorButtonTextColor (BOOL bIsUserCommand, int nCmd)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		return CBCGPVisualManagerVS2010::GetCalculatorButtonTextColor(bIsUserCommand, nCmd);
	}
	
	CBCGPCalculator::CalculatorCommands cmd = (CBCGPCalculator::CalculatorCommands) nCmd;
	
	COLORREF clrText = globalData.clrBarText;
	
	if (bIsUserCommand)
	{
		clrText = RGB (255, 0, 255);
	}
	else if (cmd != CBCGPCalculator::idCommandNone && cmd != CBCGPCalculator::idCommandDot)
	{
		clrText = m_clrAccentText;
	}
	
	return clrText;
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawSliderChannel (CDC* pDC, CBCGPSliderCtrl* pSlider, BOOL bVert, CRect rect, BOOL bDrawOnGlass)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawSliderChannel (pDC, pSlider, bVert, rect, bDrawOnGlass);
		return;
	}

	ASSERT_VALID (pDC);

	COLORREF clrLine = globalData.clrBarDkShadow;
	CBCGPDrawManager dm(*pDC);

	if (pSlider->GetSafeHwnd() == NULL || (pSlider->GetStyle() & TBS_ENABLESELRANGE) == 0)
	{
		if (bVert)
		{
			dm.DrawLine(rect.CenterPoint().x, rect.top, rect.CenterPoint().x, rect.bottom, clrLine);
		}
		else
		{
			dm.DrawLine(rect.left, rect.CenterPoint().y, rect.right, rect.CenterPoint().y, clrLine);
		}

		return;
	}

	CRect rectSel = pSlider->GetSelectionRect();
	
	dm.DrawRect(rect, m_clrMenuLight, (COLORREF)-1);
		
	if (!rectSel.IsRectEmpty())
	{
		if (bVert)
		{
			rectSel.DeflateRect(1, 0);
		}
		else
		{
			rectSel.DeflateRect(0, 1);
		}

		COLORREF clrSelection = pSlider->IsWindowEnabled() ? pSlider->GetSelectionColor() : globalData.clrBarShadow;
		if (clrSelection == (COLORREF)-1)
		{
			clrSelection = m_clrHighlightDn;
		}
		
		dm.DrawRect(rectSel, clrSelection, (COLORREF)-1);
	}
	
	dm.DrawRect(rect, (COLORREF)-1, globalData.clrBarShadow);
}
//********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawSliderThumb (CDC* pDC, CBCGPSliderCtrl* pSlider, 
			CRect rect, BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled,
			BOOL bVert, BOOL bLeftTop, BOOL bRightBottom,
			BOOL bDrawOnGlass)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawSliderThumb (
			pDC, pSlider, rect, bIsHighlighted, bIsPressed, bIsDisabled,
			bVert, bLeftTop, bRightBottom, bDrawOnGlass);
		return;
	}

	ASSERT_VALID (pDC);

	COLORREF clrLine = bIsDisabled ? globalData.clrBarShadow : globalData.clrBarDkShadow;
	COLORREF clrFill = bIsDisabled ? globalData.clrBarDkShadow : m_Style == VS2012_Dark ? RGB(200, 200, 200) : globalData.clrBarFace;

	if (!bIsDisabled)
	{
		if (bIsPressed || bIsHighlighted)
		{
			clrFill = m_Style == VS2012_Dark ? m_clrAccentText : m_clrAccentLight;
			clrLine = m_Style == VS2012_LightBlue ? m_clrAccent : m_clrHighlightDn;
		}
	}

	if (bVert)
	{
		rect.DeflateRect (2, 1);

		rect.left = rect.CenterPoint ().x - rect.Height ();
		rect.right = rect.left + 2 * rect.Height ();

		rect.top = rect.CenterPoint ().y - 2;
		rect.bottom = rect.CenterPoint ().y + 2;
	}
	else
	{
		rect.DeflateRect (1, 2);

		rect.top = rect.CenterPoint ().y - rect.Width () + 1;
		rect.bottom = rect.top + 2 * rect.Width () - 2;

		rect.left = rect.CenterPoint ().x - 2;
		rect.right = rect.CenterPoint ().x + 2;
	}

	CBCGPDrawManager dm (*pDC);
	dm.DrawRect (rect, clrFill, clrLine);
}
//********************************************************************************
BOOL CBCGPVisualManagerVS2012::OnSetWindowRegion (CWnd* pWnd, CSize sizeWindow)
{ 
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::OnSetWindowRegion(pWnd, sizeWindow);
	}

	if (pWnd->GetSafeHwnd() == NULL)
	{
		return FALSE;
	}

#ifndef BCGP_EXCLUDE_RIBBON
	if (DYNAMIC_DOWNCAST (CBCGPRibbonBar, pWnd) != NULL)
	{
		return FALSE;
	}
#endif
	else
	{
		if (pWnd->IsZoomed ())
		{
			pWnd->SetWindowRgn (NULL, TRUE);
			return TRUE;
		}
	}

    CRgn rgn;

	if (rgn.CreateRectRgn (0, 0, sizeWindow.cx + 1, sizeWindow.cy + 1))
	{
		pWnd->SetWindowRgn ((HRGN)rgn.Detach (), TRUE);
		return TRUE;
	}

	return FALSE;
}
//********************************************************************************
BOOL CBCGPVisualManagerVS2012::IsOwnerDrawCaption ()
{	
	return globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode ();	
}

#if !defined (BCGP_EXCLUDE_GRID_CTRL) && !defined (BCGP_EXCLUDE_GANTT)

void CBCGPVisualManagerVS2012::DrawGanttHeaderCell (const CBCGPGanttChart* /*pChart*/, CDC& dc, const BCGP_GANTT_CHART_HEADER_CELL_INFO& cellInfo, BOOL /*bHilite*/)
{
	dc.FillRect (cellInfo.rectCell, &globalData.brBarFace);

	CBCGPPenSelector pen(dc, globalData.clrBarShadow);
	
	dc.MoveTo (cellInfo.rectCell.right - 1, cellInfo.rectCell.top);
	dc.LineTo (cellInfo.rectCell.right - 1, cellInfo.rectCell.bottom);
	
	dc.MoveTo (cellInfo.rectCell.left, cellInfo.rectCell.bottom - 1);
	dc.LineTo (cellInfo.rectCell.right, cellInfo.rectCell.bottom - 1);
	
}
//********************************************************************************
void CBCGPVisualManagerVS2012::GetGanttColors (const CBCGPGanttChart* pChart, BCGP_GANTT_CHART_COLORS& colors, COLORREF clrBack) const
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || m_Style == VS2012_LightBlue)
	{
		CBCGPVisualManagerVS2010::GetGanttColors(pChart, colors, clrBack);
		return;
	}

	if (pChart->GetSafeHwnd() != NULL && !pChart->IsVisualManagerStyle())
	{
		CBCGPVisualManagerVS2010::GetGanttColors(pChart, colors, clrBack);
		return;
	}

	BOOL bIsDefault = FALSE;

	if (clrBack == CLR_DEFAULT)
	{
		clrBack = globalData.clrBarLight;
		bIsDefault = TRUE;
	}	

    colors.clrBackground      = globalData.clrBarLight;
	colors.clrShadows         = m_clrMenuShadowBase;

	colors.clrRowBackground   = colors.clrBackground;

	colors.clrBarFill         = m_clrAccent;
	colors.clrBarComplete     = m_clrAccentLight;

	colors.clrRowBackground  = CBCGPDrawManager::SmartMixColors(globalData.clrBarFace, globalData.clrBarHilite);

	if (bIsDefault || IsDarkTheme())
	{
		colors.clrRowDayOff = IsDarkTheme() ? globalData.clrBarLight : CBCGPDrawManager::ColorMakeDarker(colors.clrRowBackground, .05);
	}
	else
	{
		colors.clrRowDayOff = CBCGPDrawManager::ColorMakePale(clrBack, .9);
	}

	colors.clrConnectorLines = globalData.clrBarDkShadow;

	colors.clrGridLine0 = IsDarkTheme() ? m_clrSeparator : CalculateHourLineColor (colors.clrRowDayOff, TRUE, TRUE);
	colors.clrGridLine1 = IsDarkTheme() ? m_clrSeparator : CalculateHourLineColor (colors.clrRowDayOff, TRUE, !bIsDefault);

	colors.clrSelection = m_clrAccentHilight;
	colors.clrSelectionBorder = m_clrAccent;
}

#endif // !defined (BCGP_EXCLUDE_GRID_CTRL) && !defined (BCGP_EXCLUDE_GANTT)

UINT CBCGPVisualManagerVS2012::GetNavButtonsID(BOOL bIsLarge)
{
	return bIsLarge ? IDB_BCGBARRES_NAV_BUTTONS_120_W8 : IDB_BCGBARRES_NAV_BUTTONS_W8;
}
//*****************************************************************************
UINT CBCGPVisualManagerVS2012::GetNavFrameID(BOOL bIsLarge)
{
	return bIsLarge ? IDB_BCGBARRES_NAV_FRAMES_120_W8 : IDB_BCGBARRES_NAV_FRAMES_W8;
}
//*****************************************************************************
int CBCGPVisualManagerVS2012::GetNavImageState(BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled, BOOL bIsDrawOnGlass)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetNavImageState(bIsHighlighted, bIsPressed, bIsDisabled, bIsDrawOnGlass);
	}

	CBCGPMenuImages::IMAGE_STATE stateImage = CBCGPMenuImages::ImageWhite;

	if (m_Style == VS2012_Dark)
	{
		if (bIsDisabled)
		{
			stateImage = CBCGPMenuImages::ImageDkGray;
		}
		else if (bIsHighlighted)
		{
			stateImage = CBCGPMenuImages::ImageBlack2;
		}
	}
	else
	{
		stateImage = CBCGPMenuImages::ImageDkGray;

		if (bIsDisabled)
		{
			stateImage = CBCGPMenuImages::ImageLtGray;
		}
		else if (bIsHighlighted)
		{
			stateImage = CBCGPMenuImages::ImageBlack2;
		}
	}

	return (int)stateImage;
}
//*********************************************************************************
int CBCGPVisualManagerVS2012::CreateAutoHideButtonRegion (CRect rect, 
								DWORD dwAlignment, LPPOINT& points)
{
	if (m_bRoundedAutohideButtons)
	{
		return CBCGPVisualManagerVS2010::CreateAutoHideButtonRegion(rect, dwAlignment, points);
	}

	switch (dwAlignment & CBRS_ALIGN_ANY)
	{
	case CBRS_ALIGN_LEFT:
		rect.right--;
		break;

	case CBRS_ALIGN_TOP:
		rect.bottom--;
		break;
	}

	CRect rectOrign = rect;
	DWORD dwAlignmentOrign = dwAlignment;

	if ((dwAlignment & CBRS_ALIGN_ANY) == CBRS_ALIGN_LEFT || 
		(dwAlignment & CBRS_ALIGN_ANY) == CBRS_ALIGN_RIGHT)
	{
		rect = CRect (0, 0, rectOrign.Height (), rectOrign.Width ());
		dwAlignment = (dwAlignment == CBRS_ALIGN_LEFT) ? CBRS_ALIGN_TOP : CBRS_ALIGN_BOTTOM;
	}

	CList<POINT, POINT> pts;

	rect.right--;

	pts.AddHead (CPoint (rect.left, rect.top));
	pts.AddHead (CPoint (rect.left, rect.bottom));
	pts.AddHead (CPoint (rect.right, rect.bottom));
	pts.AddHead (CPoint (rect.right, rect.top));

	points = new POINT [pts.GetCount ()];

	int i = 0;

	for (POSITION pos = pts.GetHeadPosition (); pos != NULL; i++)
	{
		points [i] = pts.GetNext (pos);

		switch (dwAlignmentOrign & CBRS_ALIGN_ANY)
		{
		case CBRS_ALIGN_BOTTOM:
			points [i].y = rect.bottom - (points [i].y - rect.top);
			break;

		case CBRS_ALIGN_RIGHT:
			{
				int x = rectOrign.right - points [i].y;
				int y = rectOrign.top + points [i].x;

				points [i] = CPoint (x, y);
			}
			break;

		case CBRS_ALIGN_LEFT:
			{
				int x = rectOrign.left + points [i].y;
				int y = rectOrign.top + points [i].x;

				points [i] = CPoint (x, y);
			}
			break;
		}
	}

	return (int) pts.GetCount ();
}
//***********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawCaptionBarBorder (CDC* pDC, 
	CBCGPCaptionBar* pBar, CRect rect, COLORREF clrBarBorder, BOOL bFlatBorder)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode() || m_Style == VS2012_LightBlue || clrBarBorder != (COLORREF) -1 || pBar == NULL ||
		(pBar != NULL && pBar->IsDialogControl ()))
	{
		CBCGPVisualManagerVS2010::OnDrawCaptionBarBorder(pDC, pBar, rect, clrBarBorder, bFlatBorder);
		return;
	}

	ASSERT_VALID (pDC);

	if (!pBar->IsMessageBarMode() && pBar->m_clrBarBackground != -1)
	{
		pDC->FillSolidRect(rect, pBar->m_clrBarBackground);
	}
	else
	{
		pDC->FillRect(rect, &globalData.brBarFace);
	}
}
//***********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawTearOffCaption (CDC* pDC, CRect rectCaption, BOOL bIsActive)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawTearOffCaption (pDC, rectCaption, bIsActive);
		return;
	}

	pDC->FillRect (rectCaption, &m_brTearOffCaption);

	CBCGPDrawManager dm(*pDC);

	int xPadding = globalUtils.ScaleByDPI(2);
	int yPadding = globalUtils.ScaleByDPI(3);
	int dx = rectCaption.Height() - xPadding;
	
	for (int x = rectCaption.CenterPoint().x - 3 * dx; x <= rectCaption.CenterPoint().x + 3 * dx; x += dx)
	{
		CRect rect = rectCaption;
		rect.left = x;
		rect.right = x + dx;
		rect.DeflateRect(xPadding, yPadding);
		
		dm.DrawEllipse(rect, bIsActive ? m_clrAccent : globalData.clrBarText, (COLORREF)-1);
	}
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::OnDrawProgressMarqueeDot(CDC* pDC, CBCGPProgressCtrl* pProgressCtrl, CRect rect, int nIndex)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pProgressCtrl);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || !pProgressCtrl->m_bVisualManagerStyle ||
		pProgressCtrl->m_clrMarquee != (COLORREF)-1)
	{
		CBCGPVisualManagerVS2010::OnDrawProgressMarqueeDot(pDC, pProgressCtrl, rect, nIndex);
		return;
	}

	if (pProgressCtrl->m_bOnGlass)
	{
		CBCGPDrawManager dm(*pDC);
		dm.DrawRect(rect, m_clrAccent, (COLORREF)-1);
	}
	else
	{
		pDC->FillRect(rect, &m_brAccent);
	}
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::GetInfoBoxColors(CBCGPInfoBoxRenderer* pCtrl, COLORREF& clrBackground, COLORREF& clrForeground, COLORREF& clrFrame, COLORREF& clrLink)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode())
	{
		CBCGPVisualManagerVS2010::GetInfoBoxColors(pCtrl, clrBackground, clrForeground, clrFrame, clrLink);
		return;
	}

	if (clrBackground == CLR_DEFAULT)
	{
		clrBackground = IsDarkTheme() ? m_clrControl : RGB(252, 245, 195);
	}
	
	if (clrFrame == CLR_DEFAULT)
	{
		clrFrame = IsDarkTheme() ? m_clrSeparator : RGB(220, 220, 220);
	}
	
	if (clrForeground == CLR_DEFAULT)
	{
		clrForeground = globalData.clrBarText;
	}

	if (clrLink == CLR_DEFAULT)
	{
		clrLink = IsDarkTheme() ? m_clrAccentText : globalData.clrHotText;
	}
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawSplitterBar(CDC* pDC, CBCGPSplitterWnd* pSplitterWnd, CRect& rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawSplitterBar(pDC, pSplitterWnd, rect);
		return;
	}

	CBCGPControlBar* pParent = pSplitterWnd->GetSafeHwnd() == NULL ? NULL : DYNAMIC_DOWNCAST(CBCGPControlBar, pSplitterWnd->GetParent());
	if (pParent != NULL)
	{
		OnFillBackground(pDC, rect, pParent);
	}
	else
	{
		CBrush br(globalData.clrBarShadow);
		pDC->FillRect(rect, &br);
	}
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::OnDrawEditClearButton(CDC* pDC, CBCGPButton* pButton, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawEditClearButton(pDC, pButton, rect);
		return;
	}

	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	CBCGPDrawManager dm (*pDC);

	if (pButton->IsPressed())
	{
		OnFillHighlightedArea (pDC, rect, &m_brHighlightDn, NULL);
		dm.DrawRect (rect, (COLORREF)-1, m_clrMenuItemBorder);
	}
	else if (pButton->IsHighlighted())
	{
		OnFillHighlightedArea (pDC, rect, &m_brHighlight, NULL);
		dm.DrawRect (rect, (COLORREF)-1, m_clrMenuItemBorder);
	}
	else
	{
		dm.DrawRect (rect, globalData.clrBarFace, (COLORREF)-1);
	}
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::BreadcrumbDrawItemBackground (CDC& dc, CBCGPBreadcrumb* pControl, BREADCRUMBITEMINFO* pItemInfo, CRect rectItem, UINT uState, COLORREF color)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || color != CLR_INVALID || !pControl->m_bVisualManagerStyle)
	{
		CBCGPVisualManagerXP::BreadcrumbDrawItemBackground (dc, pControl, pItemInfo, rectItem, uState, color);
		return;
	}
	
	if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
	{	
		CBCGPDrawManager dm(dc);

		COLORREF clrFill = uState == CDIS_SELECTED ? m_clrHighlightDn : uState == CDIS_HOT ? m_clrHighlight : m_clrHighlightChecked;
		dm.DrawRect(rectItem, clrFill, (COLORREF)-1);
	}
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::BreadcrumbDrawArrowBackground (CDC& dc, CBCGPBreadcrumb* pControl, BREADCRUMBITEMINFO* pItemInfo, CRect rectArrow, UINT uState, COLORREF color)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || color != CLR_INVALID || !pControl->m_bVisualManagerStyle)
	{
		CBCGPVisualManagerXP::BreadcrumbDrawArrowBackground (dc, pControl, pItemInfo, rectArrow, uState, color);
		return;
	}
	
	if (uState == CDIS_OTHERSIDEHOT) 
	{
		uState = CDIS_HOT;
	}

	if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
	{	
		CBCGPDrawManager dm(dc);

		COLORREF clrFill = uState == CDIS_SELECTED ? m_clrHighlightDn : uState == CDIS_HOT ? m_clrHighlight : m_clrHighlightChecked;
		dm.DrawRect(rectArrow, clrFill, (COLORREF)-1);
	}
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::BreadcrumbDrawLeftArrowBackground (CDC& dc, CBCGPBreadcrumb* pControl, CRect rect, UINT uState, COLORREF color)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || color != CLR_INVALID || !pControl->m_bVisualManagerStyle)
	{
		CBCGPVisualManager::BreadcrumbDrawLeftArrowBackground (dc, pControl, rect, uState, color);
		return;
	}
	
	if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
	{	
		CBCGPDrawManager dm(dc);

		COLORREF clrFill = uState == CDIS_SELECTED ? m_clrHighlightDn : uState == CDIS_HOT ? m_clrHighlight : m_clrHighlightChecked;
		dm.DrawRect(rect, clrFill, (COLORREF)-1);
	}
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::BreadcrumbDrawItem (CDC& dc, CBCGPBreadcrumb* pControl, BREADCRUMBITEMINFO* pItemInfo, CRect rectItem, UINT uState, COLORREF color)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || color != CLR_INVALID || !pControl->m_bVisualManagerStyle)
	{
		CBCGPVisualManager::BreadcrumbDrawItem (dc, pControl, pItemInfo, rectItem, uState, color);
		return;
	}
	
	if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
	{
		color = (uState == CDIS_HOT) ? m_clrAccentText : globalData.clrBarText;

		if (m_Style == VS2012_Light && uState == CDIS_SELECTED)
		{
			color = RGB(255, 255, 255);
		}
	}
	
	CBCGPVisualManager::BreadcrumbDrawItem (dc, pControl, pItemInfo, rectItem, uState, color);
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::BreadcrumbDrawArrow (CDC& dc, CBCGPBreadcrumb* pControl, BREADCRUMBITEMINFO* pItemInfo, CRect rect, UINT uState, COLORREF color)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || color != CLR_INVALID || !pControl->m_bVisualManagerStyle)
	{
		CBCGPVisualManager::BreadcrumbDrawArrow (dc, pControl, pItemInfo, rect, uState, color);
		return;
	}
	
	if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
	{
		color = (uState == CDIS_HOT) ? m_clrAccentText : globalData.clrBarText;

		if (m_Style == VS2012_Light && uState == CDIS_SELECTED)
		{
			color = RGB(255, 255, 255);
		}
	}
	
	CBCGPVisualManager::BreadcrumbDrawArrow (dc, pControl, pItemInfo, rect, uState, color);
}
//*********************************************************************************
void CBCGPVisualManagerVS2012::BreadcrumbDrawLeftArrow (CDC& dc, CBCGPBreadcrumb* pControl, CRect rect, UINT uState, COLORREF color)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || color != CLR_INVALID || !pControl->m_bVisualManagerStyle)
	{
		CBCGPVisualManager::BreadcrumbDrawLeftArrow (dc, pControl, rect, uState, color);
		return;
	}

	if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
	{
		color = (uState == CDIS_HOT) ? m_clrAccentText : globalData.clrBarText;
	}
	
	CBCGPVisualManager::BreadcrumbDrawLeftArrow (dc, pControl, rect, uState, color);
}
//*******************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetDlgHeaderTextColor(CWnd* pDlg)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return CBCGPVisualManagerVS2010::GetDlgHeaderTextColor(pDlg);
	}
	
	return m_clrAccentText;
}
//*******************************************************************************
void CBCGPVisualManagerVS2012::OnDrawMenuResizeBar (CDC* pDC, CRect rect, int nResizeFlags)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::OnDrawMenuResizeBar(pDC, rect, nResizeFlags);
		return;
	}

	pDC->FillRect(rect, &globalData.brBarFace);
	OnDrawMenuResizeGipper(pDC, rect, nResizeFlags, m_Style == VS2012_Dark ? globalData.clrBarText : m_clrMenuBorder);
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::GetChartColors(CBCGPChartColors& colors)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2010::GetChartColors(colors);
		return;
	}

	colors.m_clrFill = m_Style == VS2012_Dark ? globalData.clrBarFace : globalData.clrBarHilite;
	colors.m_clrText = m_Style == VS2012_Dark ? CBCGPDrawManager::ColorMakeDarker(globalData.clrBarText, .4) : globalData.clrBarText;
	colors.m_clrOutline = m_Style == VS2012_Dark ? (COLORREF)colors.m_clrText : IsDarkTheme() ? globalData.clrBarLight : globalData.clrBarDkShadow;
	colors.m_clrPlotter = m_Style == VS2012_Dark ? CBCGPDrawManager::ColorMakeDarker(globalData.clrBarLight, .2) : (COLORREF)colors.m_clrFill;
	colors.m_dblInterlaceOpacity = m_Style == VS2012_Dark ? 0.2 : IsDarkTheme() ? 0.04 : 0.08;
	
	if (m_Style == VS2012_Dark)
	{
		colors.m_colorScrollBar = globalData.clrBarLight;
		colors.m_colorScrollBarGradient = globalData.clrBarHilite;
	}
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::GetCircularGaugeColors(CBCGPCircularGaugeColors& colors)
{
	CBCGPVisualManagerVS2010::GetCircularGaugeColors(colors);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return;
	}

	LOGBRUSH lbr;
	m_brControl.GetLogBrush(&lbr);
	
	colors.m_brFill.SetColor(lbr.lbColor);
	colors.m_brText.SetColor(m_Style == VS2012_Dark ? CBCGPDrawManager::ColorMakeDarker(globalData.clrBarText, .4) : globalData.clrBarText);

	COLORREF clrOutline = globalData.clrBarDkShadow;

	colors.m_brPointerOutline.SetColor(m_clrAccentText);
	colors.m_brPointerFill.SetColor(m_clrAccent);

	colors.m_brFrameOutline.SetColor(clrOutline);
	colors.m_brFrameFill.SetColors(globalData.clrBarFace, globalData.clrBarHilite, CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT);

	colors.m_brCapOutline.SetColor(clrOutline);
	colors.m_brCapFill.SetColor(globalData.clrBarLight);
	
	colors.m_brTickMarkFill.SetColors(globalData.clrBarFace, globalData.clrBarHilite, CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT);
	colors.m_brTickMarkOutline.SetColor(clrOutline);
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::GetSwitchColors(CBCGPSwitchColors& colors, BOOL bIsVertical)
{
	CBCGPVisualManagerVS2010::GetSwitchColors(colors, bIsVertical);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return;
	}

	if (colors.IsGDICtrl() && m_Style == VS2012_Blue)
	{
		return;
	}

	colors.m_brFillOn = m_clrAccent;

	if (colors.IsGDICtrl())
	{
		colors.m_brFillOff = m_clrControl;
		colors.m_brOutline = globalData.clrBarDkShadow;
		return;
	}

	LOGBRUSH lbr;
	m_brControl.GetLogBrush(&lbr);
	
	colors.m_brFill.SetColor(lbr.lbColor);
	colors.m_brLabelOn = CBCGPDrawManager::GetContrastColor((COLORREF)colors.m_brFillOn.GetColor());
}
//*****************************************************************************
void CBCGPVisualManagerVS2012::GetLinearGaugeColors(CBCGPLinearGaugeColors& colors)
{
	CBCGPVisualManagerVS2010::GetLinearGaugeColors(colors);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return;
	}

	LOGBRUSH lbr;
	m_brControl.GetLogBrush(&lbr);
	
	colors.m_brFill.SetColor(lbr.lbColor);
	colors.m_brText.SetColor(m_Style == VS2012_Dark ? CBCGPDrawManager::ColorMakeDarker(globalData.clrBarText, .4) : globalData.clrBarText);

	COLORREF clrOutline = globalData.clrBarDkShadow;

	colors.m_brPointerOutline.SetColor(m_clrAccentText);
	colors.m_brPointerFill.SetColor(m_clrAccent);

	colors.m_brFrameOutline.SetColor(clrOutline);
	colors.m_brFrameFill.SetColors(globalData.clrBarFace, globalData.clrBarHilite, CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT);

	colors.m_brTickMarkFill.SetColors(globalData.clrBarFace, globalData.clrBarHilite, CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT);
	colors.m_brTickMarkOutline.SetColor(clrOutline);
}
//**********************************************************************************
COLORREF CBCGPVisualManagerVS2012::GetHighlightedColor(UINT nType)
{
	return CBCGPVisualManagerXP::GetHighlightedColor(nType);
}
//**********************************************************************************
CSize CBCGPVisualManagerVS2012::GetFrameCaptionButtonOffset(CWnd* pWnd)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pWnd->GetSafeHwnd() == NULL)
	{
		return CSize(0, 0);
	}

	if (pWnd->IsZoomed() || pWnd->IsIconic())
	{
		return CSize(0, 0);
	}

	CSize szSysBorder = GetSysBorderSize(pWnd);

	return CSize(szSysBorder.cx - 1, -szSysBorder.cy);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPVisualManagerVS2013

void CBCGPVisualManagerVS2013::ModifyGlobalColors ()
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2012::ModifyGlobalColors();
		return;
	}

	if (m_Style == VS2012_Light)
	{
		globalData.clrBarFace = RGB (239, 239, 242);
		globalData.clrBarText = RGB (30, 30, 30);
		globalData.clrBarShadow = RGB (204, 206, 219);
		globalData.clrBarHilite = RGB (255, 255, 255);
		globalData.clrBarDkShadow = RGB(162, 164, 165);
		globalData.clrBarLight = RGB (254, 254, 254);
	}
	else if (m_Style == VS2012_Dark)
	{
		globalData.clrBarFace = RGB (45, 45, 48);
		globalData.clrBarText = RGB (241, 241, 241);
		globalData.clrBarShadow = RGB (120, 120, 120);
		globalData.clrBarHilite = RGB (104, 104, 104);
		globalData.clrBarDkShadow = RGB(190, 190, 190);
		globalData.clrBarLight = RGB (70, 70, 75);
	}
	else
	{
		globalData.clrBarFace = RGB (214, 219, 233);
		globalData.clrBarText = RGB (27, 41, 62);
		globalData.clrBarShadow = RGB (190, 195, 203);
		globalData.clrBarHilite = RGB (252, 252, 252);
		globalData.clrBarDkShadow = RGB(133, 145, 162);
		globalData.clrBarLight = RGB (220, 224, 236);
	}

	globalData.brBarFace.DeleteObject ();
	globalData.brBarFace.CreateSolidBrush (globalData.clrBarFace);
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2013::SetupColors()
{
	CBCGPVisualManagerVS2012::SetupColors();

	m_clrAccentLight = CBCGPDrawManager::PixelAlpha(m_clrAccent, IsDarkTheme() ? 75 : 140);
	m_clrAccentHilight = CBCGPDrawManager::PixelAlpha(m_clrAccent, IsDarkTheme() ? 85 : 115);

	m_clrAccentText = CBCGPDrawManager::PixelAlpha (m_clrAccent, IsDarkTheme() ? 140 : 75);
	
	m_clrEditPrompt = IsDarkTheme() ? 
		CBCGPDrawManager::ColorMakeDarker(globalData.clrBarText, .35) : CBCGPDrawManager::ColorMakeLighter(globalData.clrWindowText, .45);
	
	m_clrMenuSeparator = m_Style == VS2012_Dark ? globalData.clrBarFace : globalData.clrBarShadow;
	
	m_clrBarBkgnd = globalData.clrBarLight;

	m_clrGripperBk = globalData.clrBarFace;

	if (m_Style == VS2012_LightBlue)
	{
		m_clrTextDisabled = RGB(128, 128, 128);
		m_clrHighlightChecked = m_clrHighlight = RGB(253, 244, 191);
		
		m_clrHighlightDn = RGB(255, 242, 157);
		m_clrHighlightMenuItem = m_clrHighlight;
		
		m_clrHighlightGradientLight = m_clrHighlight;
		m_clrHighlightGradientDark = m_clrHighlight;
		
		m_clrHighlightDnGradientDark = m_clrHighlightDn;
		m_clrHighlightDnGradientLight =  m_clrHighlightDn;
		
		m_clrMenuLightGradient = m_clrMenuLight = RGB (234, 240, 255);
		m_clrMenuGutter = RGB (242, 244, 254);
		
		m_clrMenuItemBorder = m_clrPressedButtonBorder = RGB(229, 195, 101);
		m_clrMenuItemGradientDark = m_clrMenuItemGradient1 = m_clrMenuItemGradient2 = m_clrHighlightDn;

		m_clrEditBoxBorder = globalData.clrBarDkShadow;
		m_clrMenuBorder = RGB(155, 167, 183);

		m_clrControl = globalData.clrBarHilite;
		m_clrDlgBackground = CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace, .06);

		m_clrInactiveCaptionGradient1 = m_clrInactiveCaptionGradient2 = RGB(77, 96, 130);
		m_clrToolBarGradientDark = m_clrToolBarGradientLight = RGB(207, 214, 229);

		m_clrAutohideBarActive = m_clrAccent;
		m_clrAutohideBarInactive = m_clrInactiveCaptionGradient1;

		m_clrCaptionBarGradientDark = m_clrInactiveCaptionGradient1;
		m_clrCaptionBarGradientLight = m_clrInactiveCaptionGradient2;
	}
	else
	{
		m_clrTextDisabled = m_Style == VS2012_Dark ? globalData.clrBarShadow : globalData.clrBarDkShadow;

		if (m_Style == VS2012_Dark)
		{
			m_clrHighlight = CBCGPDrawManager::ColorMakeDarker(globalData.clrBarLight);
		}
		else
		{
			if (m_curAccentColor == VS2012_Blue)
			{
				m_clrHighlight = RGB(201, 222, 245);
			}
			else
			{
				m_clrHighlight = CBCGPDrawManager::ColorMakePale(m_clrAccent, .90);
			}
		}

		m_clrHighlightDn = m_clrAccent;
		m_clrHighlightChecked = globalData.clrBarFace;
		m_clrHighlightMenuItem = (m_Style == VS2012_Dark) ? RGB(51, 51, 52) : m_clrHighlight;

		m_clrHighlightGradientLight = globalData.clrBarHilite;
		m_clrHighlightGradientDark = globalData.clrBarHilite;
		
		m_clrHighlightDnGradientDark = m_clrAccent;
		m_clrHighlightDnGradientLight =  m_clrAccent;
		
		m_clrMenuLight = m_Style == VS2012_Light ? RGB(246, 246, 246) : m_Style == VS2012_Dark ? RGB(27, 27, 27) : RGB(233, 236, 238);
		m_clrMenuLightGradient = m_clrMenuGutter = globalData.clrBarLight;
		
		m_clrMenuItemGradientDark = m_clrPressedButtonBorder = m_clrAccent;
		m_clrMenuItemBorder = m_clrHighlight;

		m_clrEditBoxBorder = (m_Style == VS2012_Dark) ? 
			CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace, .5) : 
			CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace, .09);

		m_clrMenuBorder = m_Style == VS2012_Dark ? CBCGPDrawManager::ColorMakeLighter(m_clrMenuSeparator, .2) : m_clrMenuSeparator;

		m_clrControl = m_Style != VS2012_Dark ? 
			CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace, .028) :
			CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace, .1);

		m_clrDlgBackground = m_Style == VS2012_Dark ? CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace, .2) : globalData.clrBarFace;
		m_clrToolBarGradientDark = m_clrToolBarGradientLight = globalData.clrBarFace;

		m_clrCaptionBarGradientDark = m_clrHighlight;
		m_clrCaptionBarGradientLight = m_clrHighlight;
	}

	m_clrToolbarDisabled = globalData.clrBarHilite;

	m_clrCombo = m_Style == VS2012_Dark ? CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace, .15) : globalData.clrBarHilite;
	m_clrComboDisabled = globalData.clrBarFace;
	
	m_clrReportGroupText = m_clrInactiveCaptionGradient1;

	m_clrBarGradientDark = m_clrBarGradientLight = globalData.clrBarFace;
	
	m_clrPlannerWork      = globalData.clrBarFace;
	m_clrPalennerLine     = globalData.clrBarShadow;
	m_clrPlannerTodayFill = m_clrAccent;
	
	m_clrPlannerTodayLine = m_clrAccentText;

	m_clrRibbonTabs = globalData.clrBarText;
	m_clrTabsBackground = globalData.clrBarFace;

	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageBlack, RGB(113, 113, 113));
	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageGray, m_Style == VS2012_Dark ? RGB(153, 153, 153) : RGB(113, 113, 113));

	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageBlack2, m_clrAccent);
	CBCGPMenuImages::SetColor (CBCGPMenuImages::ImageWhite, RGB(255, 255, 255));

	m_clrHighlighDownText = globalData.clrBarText;
	
	if (m_Style != VS2012_Dark)
	{
		if (GetRValue (m_clrHighlightDn) <= 128 ||
			GetGValue (m_clrHighlightDn) <= 128 ||
			GetBValue (m_clrHighlightDn) <= 128)
		{
			m_clrHighlighDownText = globalData.clrBarHilite;
		}

		m_clrComboHighlighted = m_clrHighlight;
	}

	m_clrRibbonCategoryFill = globalData.clrBarFace;

	m_clrNcTextActive   = globalData.clrBarText;
	m_clrNcTextInactive = globalData.clrBarDkShadow;
	m_clrNcBorder       = m_clrMenuBorder;

	m_clrButtonsArea = globalData.clrBarLight;
	
	if (!IsDarkTheme())
	{
		m_clrButtonsArea = CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace, 
			m_Style == VS2012_LightBlue ? .1 : .05);
	}

	m_clrFace = m_Style == VS2012_LightBlue ? m_clrButtonsArea : m_clrToolBarGradientLight;
}
//**********************************************************************************************************
void CBCGPVisualManagerVS2013::OnEraseTabsArea (CDC* pDC, CRect rect, const CBCGPBaseTabWnd* pTabWnd)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || pTabWnd->IsMDITab() ||
		(m_Style != VS2012_LightBlue || pTabWnd->IsPropertySheetTab() || !pTabWnd->Is3DStyle()) || pTabWnd->IsDialogControl())
	{
		CBCGPVisualManagerVS2012::OnEraseTabsArea (pDC, rect, pTabWnd);
		return;
	}

	CBCGPDrawManager dm (*pDC);
	dm.FillGradient(rect, m_clrInactiveCaptionGradient1, m_clrInactiveCaptionGradient2);
}
//*********************************************************************************************************
void CBCGPVisualManagerVS2013::OnDrawTab (CDC* pDC, CRect rectTab, int iTab, BOOL bIsActive, const CBCGPBaseTabWnd* pTabWnd)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || !pTabWnd->Is3DStyle() || pTabWnd->IsPropertySheetTab())
	{
		CBCGPVisualManagerVS2012::OnDrawTab (pDC, rectTab, iTab, bIsActive, pTabWnd);
		return;
	}

	if (pTabWnd->IsMDITab())
	{
		COLORREF clrHighlight = m_clrHighlight;

		if (m_Style == VS2012_LightBlue)
		{
			m_clrHighlight = m_clrHighlightDn;
		}

		CBCGPVisualManagerVS2012::OnDrawTab (pDC, rectTab, iTab, bIsActive, pTabWnd);

		m_clrHighlight = clrHighlight;
		return;
	}
	
	const BOOL bIsHighlight = iTab == pTabWnd->GetHighlightedTab ();
	
	if (bIsActive || bIsHighlight)
	{
		CRect rectFill = rectTab;
		
		if (bIsActive)
		{
			if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM)
			{
				rectFill.top--;
				rectTab.top--;
			}
			else
			{
				rectFill.bottom++;
				rectTab.bottom++;
			}

			pDC->FillSolidRect(rectFill, m_clrControl);
		}
		else
		{
			rectFill.DeflateRect(1, 1);

			if (m_Style == VS2012_LightBlue && !pTabWnd->IsDialogControl())
			{
				pDC->FillSolidRect(rectFill, CBCGPDrawManager::ColorMakeDarker(m_clrInactiveCaptionGradient1));
			}
			else
			{
				pDC->FillRect (rectFill, &m_brHighlight);
			}
		}
	}
	
	CRect rectTabFrame = rectTab;
	
	COLORREF clrBkTab = pTabWnd->GetTabBkColor(iTab);
	if (clrBkTab != (COLORREF)-1)
	{
		CRect rectColor = rectTab;
		const int nColorBarHeight = max(3, rectTab.Height() / 6);
		
		if (pTabWnd->GetLocation() == CBCGPBaseTabWnd::LOCATION_BOTTOM)
		{
			rectColor.top = rectColor.bottom - nColorBarHeight;
			rectTab.bottom -= nColorBarHeight;
		}
		else
		{
			rectColor.bottom = rectColor.top + nColorBarHeight;
			rectTab.top += nColorBarHeight;
		}
		
		rectColor.DeflateRect(bIsActive ? 1 : 2, 0);
		
		CBrush br(clrBkTab);
		pDC->FillRect(rectColor, &br);
	}
	
	if (bIsActive)
	{
		CBCGPPenSelector ps(*pDC, globalData.clrBarShadow);
		
		if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM)
		{
			pDC->MoveTo (rectTabFrame.left, rectTabFrame.top);
			pDC->LineTo (rectTabFrame.left, rectTabFrame.bottom);
			pDC->LineTo (rectTabFrame.right - 1, rectTabFrame.bottom);
			pDC->LineTo (rectTabFrame.right - 1, rectTabFrame.top);
		}
		else
		{
			pDC->MoveTo (rectTabFrame.left, rectTabFrame.bottom - 1);
			pDC->LineTo (rectTabFrame.left, rectTabFrame.top);
			pDC->LineTo (rectTabFrame.right - 1, rectTabFrame.top);
			pDC->LineTo (rectTabFrame.right - 1, rectTabFrame.bottom);
		}
	}
	
	COLORREF clrTabText = pTabWnd->GetTabTextColor(iTab);
	if (clrTabText == (COLORREF)-1)
	{
		if (m_Style == VS2012_LightBlue)
		{
			if (pTabWnd->IsDialogControl())
			{
				clrTabText = bIsActive ? m_clrAccentText : globalData.clrBarText;
			}
			else
			{
				clrTabText = bIsActive ? m_clrAccentText : globalData.clrBarLight;
			}
		}
		else
		{
			clrTabText = (bIsActive || bIsHighlight) ? m_clrAccentText : globalData.clrBarText;
		}
	}
	
	OnDrawTabContent(pDC, rectTab, iTab, bIsActive, pTabWnd, clrTabText);
}
//*************************************************************************************
void CBCGPVisualManagerVS2013::OnDrawComboBorder (CDC* pDC, CRect rect,
												BOOL bDisabled,
												BOOL bIsDropped,
												BOOL bIsHighlighted,
												CBCGPToolbarComboBoxButton* pButton)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || CBCGPToolBarImages::m_bIsDrawOnGlass ||
		m_Style != VS2012_Light)
	{
		CBCGPVisualManagerVS2012::OnDrawComboBorder (pDC, rect, bDisabled, bIsDropped, bIsHighlighted, pButton);
		return;
	}

	COLORREF clrFrame = bDisabled ? globalData.clrBarShadow : bIsHighlighted ? m_clrAccentText : m_clrEditBoxBorder;
	pDC->Draw3dRect(rect, clrFrame, clrFrame);
}
//****************************************************************************************
void CBCGPVisualManagerVS2013::OnFillMenuImageRect (CDC* pDC, CBCGPToolbarButton* pButton, CRect rect, CBCGPVisualManager::BCGBUTTON_STATE state)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		CBCGPVisualManagerVS2012::OnFillMenuImageRect(pDC, pButton, rect, state);
		return;
	}

	if ((pButton->m_nStyle & TBBS_CHECKED) && !(pButton->m_nStyle & TBBS_DISABLED))
	{
		ASSERT_VALID (pDC);
		ASSERT_VALID (pButton);

		if (DYNAMIC_DOWNCAST(CBCGPCustomizeMenuButton, pButton) == NULL)
		{
			if (state == CBCGPVisualManager::ButtonsIsHighlighted && m_Style == CBCGPVisualManagerVS2012::VS2012_Light)
			{
				return;
			}

			if (state == CBCGPVisualManager::ButtonsIsHighlighted)
			{
				rect.DeflateRect(1, 1);
				pDC->FillSolidRect(rect, CBCGPDrawManager::ColorMakeLighter(m_clrHighlight, .12));
			}
			else if (pButton->GetImage() >= 0)
			{
				OnFillButtonInterior (pDC, pButton, rect, state);
			}
			else
			{
				rect.DeflateRect(1, 1);
				pDC->FillSolidRect(rect, globalData.clrBarFace);
			}
		}
	}
}
//****************************************************************************************
void CBCGPVisualManagerVS2013::OnDrawMenuImageRectBorder (CDC* pDC, CBCGPToolbarButton* pButton, CRect rect, CBCGPVisualManager::BCGBUTTON_STATE state)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () ||
		(pButton != NULL && pButton->GetImage() >= 0) || DYNAMIC_DOWNCAST(CBCGPCustomizeMenuButton, pButton) != NULL)
	{
		CBCGPVisualManagerVS2012::OnDrawMenuImageRectBorder(pDC, pButton, rect, state);
	}
}
