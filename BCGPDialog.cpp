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

#include "stdafx.h"

#include "BCGPDialog.h"
#include "BCGPWnd.h"
#include "BCGPMDIFrameWnd.h"
#include "BCGPFrameWnd.h"
#include "bcgprores.h"

#ifndef _BCGSUITE_
#include "BCGPPopupMenu.h"
#include "BCGPToolbarMenuButton.h"
#include "BCGPPngImage.h"
#include "BCGPRibbonBar.h"
#include "BCGPTooltipManager.h"
#include "BCGPMiniFrameWnd.h"
#include "BCGPRibbonStatusBar.h"
#else
#define CBCGPRibbonBar				CMFCRibbonBar
#define CBCGPRibbonStatusBar		CMFCRibbonStatusBar
#define CBCGPBaseRibbonElement		CMFCRibbonBaseElement
#define CBCGPRibbonCaptionButton	CMFCRibbonCaptionButton
#endif

#include "BCGPVisualManager.h"
#include "BCGPLocalResource.h"
#include "BCGPGestureManager.h"
#include "BCGPDrawManager.h"
#include "BCGPGlobalUtils.h"
#include "BCGPEdit.h"

#include "BCGPMath.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define EXPAND_MARGIN_VERT globalUtils.ScaleByDPI(8)

/////////////////////////////////////////////////////////////////////////////
// CBCGPDialog dialog

IMPLEMENT_DYNAMIC(CBCGPDialog, CDialog)

#pragma warning (disable : 4355)

CBCGPDialog::CBCGPDialog() :
	m_Impl (*this)
{
	CommonConstruct ();
}
//*************************************************************************************
CBCGPDialog::CBCGPDialog (UINT nIDTemplate, CWnd *pParent/*= NULL*/) : 
				CDialog (nIDTemplate, pParent),
				m_Impl (*this)
{
	CommonConstruct ();
}
//*************************************************************************************
CBCGPDialog::CBCGPDialog (LPCTSTR lpszTemplateName, CWnd *pParentWnd/*= NULL*/) : 
				CDialog(lpszTemplateName, pParentWnd),
				m_Impl (*this)
{
	CommonConstruct ();
}

#pragma warning (default : 4355)

//*************************************************************************************
void CBCGPDialog::CommonConstruct ()
{
	m_bDisableShadows = FALSE;
	m_bIsLocal = FALSE;
	m_pLocaRes = NULL;
	m_bWasMaximized = FALSE;
	m_rectBackstageWatermark.SetRectEmpty();
	m_pBackstageWatermark = NULL;
	m_pParentEdit = NULL;
	m_bParentClosePopupDlgNotified = FALSE;
	m_nExpandedAreaHeight = -1;
	m_nExpandedCheckBoxHeight = -1;
	m_hwndExpandCheckBoxCtrl = NULL;
	m_bExpandAreaSpecialBackground = TRUE;
	m_bIsRibbonStartPage = FALSE;
	m_bIsExpanded = TRUE;
	m_bCancelModeCapturedWindow = FALSE;
}
//*************************************************************************************
void CBCGPDialog::SetControlInfoTip(UINT nCtrlID, LPCTSTR lpszInfoTip, DWORD dwVertAlign, BOOL bRedrawInfoTip, CBCGPControlInfoTip::BCGPControlInfoTipStyle style, BOOL bIsClickable, const CPoint& ptOffset)
{
	m_Impl.SetControlInfoTip(nCtrlID, lpszInfoTip, dwVertAlign, bRedrawInfoTip, style, bIsClickable, ptOffset);
}
//*************************************************************************************
void CBCGPDialog::SetControlInfoTip(CWnd* pWndCtrl, LPCTSTR lpszInfoTip, DWORD dwVertAlign, BOOL bRedrawInfoTip, CBCGPControlInfoTip::BCGPControlInfoTipStyle style, BOOL bIsClickable, const CPoint& ptOffset)
{
	m_Impl.SetControlInfoTip(pWndCtrl, lpszInfoTip, dwVertAlign, bRedrawInfoTip, style, bIsClickable, ptOffset);
}
//*************************************************************************************
void CBCGPDialog::SetBackgroundColor (COLORREF color, BOOL bRepaint)
{
	if (m_brBkgr.GetSafeHandle () != NULL)
	{
		m_brBkgr.DeleteObject ();
	}

	if (color != (COLORREF)-1)
	{
		m_brBkgr.CreateSolidBrush (color);
	}

	if (bRepaint && GetSafeHwnd () != NULL)
	{
		Invalidate ();
		UpdateWindow ();
	}
}
//*************************************************************************************
void CBCGPDialog::SetBackgroundImage (HBITMAP hBitmap, 
								BackgroundLocation location,
								BOOL bAutoDestroy,
								BOOL bRepaint)
{
	m_Impl.m_Background.Clear();

	if (hBitmap != NULL)
	{
		m_Impl.m_Background.SetMapTo3DColors(FALSE);
		m_Impl.m_Background.AddImage(hBitmap, TRUE);

#ifndef _BCGSUITE_
		m_Impl.m_Background.SetSingleImage(FALSE);
#else
		m_Impl.m_Background.SetSingleImage();
#endif
		if (bAutoDestroy)
		{
			::DeleteObject(hBitmap);
		}
	}

	m_Impl.m_BkgrLocation = (int)location;

	if (bRepaint && GetSafeHwnd () != NULL)
	{
		RedrawWindow();
	}
}
//*************************************************************************************
BOOL CBCGPDialog::SetBackgroundImage (UINT uiBmpResId,
									BackgroundLocation location,
									BOOL bRepaint)
{
	BOOL bRes = TRUE;

	m_Impl.m_Background.Clear();
	
	if (uiBmpResId != 0)
	{
		m_Impl.m_Background.SetMapTo3DColors(FALSE);
		bRes = m_Impl.m_Background.Load(uiBmpResId);

#ifndef _BCGSUITE_
		m_Impl.m_Background.SetSingleImage(FALSE);
#else
		m_Impl.m_Background.SetSingleImage();
#endif
	}
	
	m_Impl.m_BkgrLocation = (int)location;
	
	if (bRepaint && GetSafeHwnd () != NULL)
	{
		RedrawWindow();
	}

	return bRes;
}

BEGIN_MESSAGE_MAP(CBCGPDialog, CDialog)
	//{{AFX_MSG_MAP(CBCGPDialog)
	ON_WM_ACTIVATE()
	ON_WM_NCACTIVATE()
	ON_WM_ENABLE()	
	ON_WM_ERASEBKGND()
	ON_WM_DESTROY()
	ON_WM_CTLCOLOR()
	ON_WM_SYSCOLORCHANGE()
	ON_WM_SETTINGCHANGE()
	ON_WM_SIZE()
	ON_WM_NCPAINT()
	ON_WM_NCMOUSEMOVE()
	ON_WM_LBUTTONUP()
	ON_WM_LBUTTONDOWN()
	ON_WM_NCHITTEST()
	ON_WM_NCCALCSIZE()
	ON_WM_MOUSEMOVE()
	ON_WM_WINDOWPOSCHANGED()
	ON_WM_GETMINMAXINFO()
	ON_WM_CREATE()
	ON_WM_SHOWWINDOW()
	ON_WM_SYSCOMMAND()
	ON_WM_SETCURSOR()
	//}}AFX_MSG_MAP
	ON_WM_STYLECHANGED()
	ON_MESSAGE(WM_DWMCOMPOSITIONCHANGED, OnDWMCompositionChanged)
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)
	ON_MESSAGE(WM_SETTEXT, OnSetText)
	ON_MESSAGE(WM_POWERBROADCAST, OnPowerBroadcast)
	ON_REGISTERED_MESSAGE(BCGM_UPDATETOOLTIPS, OnBCGUpdateToolTips)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnNeedTipText)
	ON_MESSAGE(WM_DPICHANGED, OnDPIChanged)
	ON_MESSAGE(WM_THEMECHANGED, OnThemeChanged)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPDialog message handlers

void CBCGPDialog::OnStyleChanged(int nStyleType, LPSTYLESTRUCT lpStyleStruct)
{
	CDialog::OnStyleChanged(nStyleType, lpStyleStruct);
	
	if (nStyleType == GWL_EXSTYLE)
	{
		if (((lpStyleStruct->styleOld & WS_EX_LAYOUTRTL) != 0 && 
			(lpStyleStruct->styleNew & WS_EX_LAYOUTRTL) == 0 ||
			(lpStyleStruct->styleOld & WS_EX_LAYOUTRTL) == 0 && 
			(lpStyleStruct->styleNew & WS_EX_LAYOUTRTL) != 0))
		{
			OnRTLChanged ((lpStyleStruct->styleNew & WS_EX_LAYOUTRTL) != 0);
		}
	}
}
//*****************************************************************************
void CBCGPDialog::OnRTLChanged (BOOL bIsRTL)
{
	if (AfxGetMainWnd() == this)
	{
		globalData.m_bIsRTL = bIsRTL;
		CBCGPToolBarImages::EnableRTL(globalData.m_bIsRTL);
	}
}
//**************************************************************************
void CBCGPDialog::OnActivate(UINT nState, CWnd *pWndOther, BOOL /*bMinimized*/) 
{
	Default();
	m_Impl.OnActivate (nState, pWndOther);
}
//*************************************************************************************
BOOL CBCGPDialog::OnNcActivate(BOOL bActive) 
{
	//-----------------------------------------------------------
	// Do not call the base class because it will call Default()
	// and we may have changed bActive.
	//-----------------------------------------------------------
	BOOL bRes = (BOOL)DefWindowProc (WM_NCACTIVATE, bActive, 0L);

	if (bRes)
	{
		m_Impl.OnNcActivate (bActive);
	}

	return bRes;
}
//**************************************************************************************
void CBCGPDialog::OnEnable(BOOL bEnable) 
{
	CDialog::OnEnable (bEnable);
	m_Impl.OnNcActivate (bEnable);
}
//**************************************************************************************
BOOL CBCGPDialog::OnEraseBkgnd(CDC* pDC) 
{
	BOOL bRes = TRUE;

	CRect rectBottom;
	rectBottom.SetRectEmpty();

	CRect rectClient;
	GetClientRect (rectClient);

	if (m_nExpandedAreaHeight > 0 && m_nExpandedCheckBoxHeight > 0 && m_bExpandAreaSpecialBackground)
	{
		int nBottomHeight = m_nExpandedCheckBoxHeight + 2 * EXPAND_MARGIN_VERT;
		if (IsExpanded())
		{
			nBottomHeight += m_nExpandedAreaHeight;
		}
		
		rectBottom = rectClient;
		rectBottom.top = rectBottom.bottom - nBottomHeight;
	}

	if (m_brBkgr.GetSafeHandle () == NULL && !m_Impl.m_Background.IsValid() && !IsVisualManagerStyle ())
	{
		bRes = CDialog::OnEraseBkgnd (pDC);

		if (!rectBottom.IsRectEmpty() && !globalData.IsHighContastMode())
		{
			COLORREF clrLine = CBCGPDrawManager::ColorMakeDarker(::GetSysColor(COLOR_BTNFACE), .1);
			COLORREF clrLine1 = CBCGPDrawManager::ColorMakeLighter(::GetSysColor(COLOR_BTNFACE), .1);

			CBCGPDrawManager dm(*pDC);
			dm.DrawLine(rectBottom.left, rectBottom.top, rectBottom.right, rectBottom.top, clrLine);
			dm.DrawLine(rectBottom.left, rectBottom.top + 1, rectBottom.right, rectBottom.top + 1, clrLine1);
		}
	}
	else
	{
		ASSERT_VALID (pDC);

		if (m_Impl.m_BkgrLocation != 0 /* BACKGR_TILE */ || !m_Impl.m_Background.IsValid())
		{
			if (m_brBkgr.GetSafeHandle () != NULL)
			{
				pDC->FillRect (rectClient, &m_brBkgr);
			}
			else if (IsVisualManagerStyle ())
			{
#ifndef _BCGSUITE_
				if (IsBackstageMode())
				{
					CBCGPMemDC memDC(*pDC, this);
					CDC* pDCMem = &memDC.GetDC();

					CBCGPVisualManager::GetInstance ()->OnFillRibbonBackstageForm(pDCMem, this, rectClient);

#ifndef BCGP_EXCLUDE_RIBBON
					if (m_bIsRibbonStartPage)
					{
						// Syncronize the ribbon background image with the dialog background:
						OnDrawRibbonBackgroundImage(pDCMem, rectClient);

						int nLeftPaneWidth = GetRibbonStartPageLeftPaneWidth();
						if (nLeftPaneWidth > 0)
						{
							CRect rectLeftPane = rectClient;
							rectLeftPane.right = rectLeftPane.left + nLeftPaneWidth;

							CBCGPVisualManager::GetInstance()->OnFillRibbonBackstageLeftPane(pDCMem, rectLeftPane);
						}
					}
#endif
					if (!m_rectBackstageWatermark.IsRectEmpty() && !globalData.IsHighContastMode())
					{
						if (m_pBackstageWatermark != NULL)
						{
							ASSERT_VALID(m_pBackstageWatermark);
							m_pBackstageWatermark->DrawEx(pDCMem, m_rectBackstageWatermark, 0);
						}
						else
						{
							OnDrawBackstageWatermark(pDCMem, m_rectBackstageWatermark);
						}
					}

					bRes = TRUE;
				}
				else
#endif
				{
					if (IsWhiteBackground() && !globalData.IsHighContastMode() && !CBCGPVisualManager::GetInstance ()->IsDarkTheme())
					{
						pDC->FillSolidRect(rectClient, RGB(255, 255, 255));
					}
					else if (!CBCGPVisualManager::GetInstance ()->OnFillDialog (pDC, this, rectClient))
					{
						CDialog::OnEraseBkgnd (pDC);
					}
				}
			}
			else
			{
				CDialog::OnEraseBkgnd (pDC);
			}
		}

		if (!rectBottom.IsRectEmpty())
		{
			CBCGPVisualManager::GetInstance()->OnDrawDlgExpandedArea(pDC, this, rectBottom);
		}

		m_Impl.DrawBackgroundImage(pDC, rectClient);
	}

	m_Impl.DrawResizeBox(pDC);
	m_Impl.ClearAeroAreas(pDC);
	m_Impl.DrawGroupBoxes(pDC);
	m_Impl.DrawControlInfoTips(pDC);

	return bRes;
}
//**********************************************************************************
void CBCGPDialog::OnDestroy() 
{
	if (m_pParentEdit != NULL)
	{
		if (!m_bParentClosePopupDlgNotified)
		{
			m_pParentEdit->ClosePopupDlg(NULL, FALSE);
			m_bParentClosePopupDlgNotified = TRUE;
		}

		delete m_pParentEdit;
		m_pParentEdit = NULL;
	}

	m_Impl.OnDestroy ();
	CDialog::OnDestroy();
}
//***************************************************************************************
HBRUSH CBCGPDialog::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	if (m_brBkgr.GetSafeHandle () != NULL || m_Impl.m_Background.IsValid() ||
		IsVisualManagerStyle ())
	{
		HBRUSH hbr = m_Impl.OnCtlColor (pDC, pWnd, nCtlColor);
		if (hbr != NULL)
		{
			return hbr;
		}
	}	

	return CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
}
//**************************************************************************************
BOOL CBCGPDialog::PreTranslateMessage(MSG* pMsg) 
{
	if (m_Impl.PreTranslateMessage (pMsg))
	{
		return TRUE;
	}

	return CDialog::PreTranslateMessage(pMsg);
}
//**************************************************************************************
void CBCGPDialog::SetActiveMenu (CBCGPPopupMenu* pMenu)
{
	m_Impl.SetActiveMenu (pMenu);
}
//*************************************************************************************
BOOL CBCGPDialog::OnCommand(WPARAM wParam, LPARAM lParam) 
{
	if (m_hwndExpandCheckBoxCtrl != NULL && lParam == (LPARAM)m_hwndExpandCheckBoxCtrl && HIWORD(wParam) == BN_CLICKED)
	{
		Expand(!IsExpanded());
		return TRUE;
	}

	if (m_Impl.OnCommand (wParam, lParam))
	{
		return TRUE;
	}

	return CDialog::OnCommand(wParam, lParam);
}
//*************************************************************************************
INT_PTR CBCGPDialog::DoModal() 
{
	if (m_bIsLocal && m_pLocaRes == NULL)
	{
		m_pLocaRes = new CBCGPLocalResource ();
	}

	return CDialog::DoModal();
}
//*************************************************************************************
void CBCGPDialog::PreInitDialog()
{
	if (m_pLocaRes != NULL)
	{
		delete m_pLocaRes;
		m_pLocaRes = NULL;
	}
}
//*************************************************************************************
void CBCGPDialog::OnSysColorChange() 
{
	CDialog::OnSysColorChange();
	
	if (AfxGetMainWnd() == this)
	{
		globalData.UpdateSysColors ();
		CBCGPVisualManager::GetInstance ()->OnUpdateSystemColors ();
	}
}
//*************************************************************************************
void CBCGPDialog::OnSettingChange(UINT uFlags, LPCTSTR lpszSection) 
{
	CDialog::OnSettingChange(uFlags, lpszSection);
	
	if (AfxGetMainWnd() == this)
	{
		globalData.OnSettingChange ();
#ifndef _BCGSUITE_
		globalData.CheckForImmersiveColorSetChange(lpszSection);
#endif
	}
}
//*************************************************************************************
void CBCGPDialog::EnableVisualManagerStyle (BOOL bEnable, BOOL bNCArea, const CList<UINT,UINT>* plstNonSubclassedItems)
{
	ASSERT_VALID (this);

	if (IsLightBox())
	{
		bNCArea = FALSE;
	}

	m_Impl.EnableVisualManagerStyle (bEnable, bEnable && bNCArea, plstNonSubclassedItems);

	if (bEnable && bNCArea)
	{
		m_Impl.OnChangeVisualManager ();
	}
}
//*************************************************************************************
BOOL CBCGPDialog::OnInitDialog() 
{
#ifndef _BCGSUITE_
	CBCGPPopupMenu* pActivePopupMenu = CBCGPPopupMenu::GetSafeActivePopupMenu();
	if (pActivePopupMenu != NULL)
	{
		ASSERT_VALID(pActivePopupMenu);
		
		if (pActivePopupMenu->IsFloaty())
		{
			pActivePopupMenu->SendMessage(WM_CLOSE);
		}
	}
#endif

	if (AfxGetMainWnd() == this)
	{
		globalData.m_bIsRTL = (GetExStyle() & WS_EX_LAYOUTRTL);
		CBCGPToolBarImages::EnableRTL(globalData.m_bIsRTL);
	}

	if (m_bCancelModeCapturedWindow)
	{
		HWND hWndCapture = ::GetCapture();
		if (hWndCapture != NULL && ::IsWindow(hWndCapture))
		{
			::SendMessage(hWndCapture, WM_CANCELMODE, 0, 0);
		}
	}

	CDialog::OnInitDialog();

	m_Impl.m_bHasBorder = (GetStyle () & WS_BORDER) == WS_BORDER;
	m_Impl.m_bHasCaption = (GetStyle() & WS_CAPTION) == WS_CAPTION;

	if (IsVisualManagerStyle ())
	{
		m_Impl.EnableVisualManagerStyle (TRUE, IsVisualManagerNCArea ());
	}

	if (m_Impl.HasAeroMargins ())
	{
		m_Impl.EnableAero (m_Impl.m_AeroMargins);
	}

	if (IsBackstageMode())
	{
		m_Impl.EnableBackstageMode();
	}

	if (m_Impl.IsOwnerDrawCaption() && (!m_Impl.m_bHasBorder || !m_Impl.m_bHasCaption))
	{
		SetWindowPos(NULL, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
	}

	m_Impl.OnInit();

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
//*************************************************************************************
BOOL CBCGPDialog::EnableAero (BCGPMARGINS& margins)
{
	return m_Impl.EnableAero (margins);
}
//*************************************************************************************
void CBCGPDialog::GetAeroMargins (BCGPMARGINS& margins) const
{
	m_Impl.GetAeroMargins (margins);
}
//***************************************************************************
LRESULT CBCGPDialog::OnDWMCompositionChanged(WPARAM,LPARAM)
{
	m_Impl.OnDWMCompositionChanged ();
	return 0;
}
//***************************************************************************
void CBCGPDialog::OnSize(UINT nType, int cx, int cy) 
{
	BOOL bIsMinimized = (nType == SIZE_MINIMIZED);

	if (m_Impl.IsOwnerDrawCaption ())
	{
		CRect rectWindow;
		GetWindowRect (rectWindow);

		WINDOWPOS wndpos;
		wndpos.flags = SWP_FRAMECHANGED;
		wndpos.x     = rectWindow.left;
		wndpos.y     = rectWindow.top;
		wndpos.cx    = rectWindow.Width ();
		wndpos.cy    = rectWindow.Height ();

		m_Impl.OnWindowPosChanged (&wndpos);
	}

	m_Impl.UpdateCaption ();
	m_Impl.UpdateToolTipsRect();

	if (!bIsMinimized && nType != SIZE_MAXIMIZED && !m_bWasMaximized)
	{
		AdjustControlsLayout();
		return;
	}

	CDialog::OnSize(nType, cx, cy);

	m_bWasMaximized = (nType == SIZE_MAXIMIZED);

	AdjustControlsLayout();
}
//**************************************************************************
void CBCGPDialog::OnNcPaint() 
{
	if (!m_Impl.OnNcPaint ())
	{
		Default ();
	}
}
//**************************************************************************
void CBCGPDialog::OnNcMouseMove(UINT nHitTest, CPoint point) 
{
	m_Impl.OnNcMouseMove (nHitTest, point);
	CDialog::OnNcMouseMove(nHitTest, point);
}
//**************************************************************************
void CBCGPDialog::OnLButtonUp(UINT nFlags, CPoint point) 
{
	m_Impl.OnLButtonUp (point);
	CDialog::OnLButtonUp(nFlags, point);
}
//**************************************************************************
void CBCGPDialog::OnLButtonDown(UINT nFlags, CPoint point) 
{
	m_Impl.OnLButtonDown (point);
	CDialog::OnLButtonDown(nFlags, point);
}
//**************************************************************************
BCGNcHitTestType CBCGPDialog::OnNcHitTest(CPoint point) 
{
	BCGNcHitTestType nHit = m_Impl.OnNcHitTest (point);
	if (nHit != HTNOWHERE)
	{
		return nHit;
	}

	return CDialog::OnNcHitTest(point);
}
//**************************************************************************
void CBCGPDialog::OnNcCalcSize(BOOL bCalcValidRects, NCCALCSIZE_PARAMS FAR* lpncsp) 
{
	if (!m_Impl.OnNcCalcSize (bCalcValidRects, lpncsp))
	{
		CDialog::OnNcCalcSize(bCalcValidRects, lpncsp);
	}
}
//**************************************************************************
void CBCGPDialog::OnMouseMove(UINT nFlags, CPoint point) 
{
	m_Impl.OnMouseMove (point);

	CDialog::OnMouseMove(nFlags, point);
}
//**************************************************************************
void CBCGPDialog::OnWindowPosChanged(WINDOWPOS FAR* lpwndpos) 
{
	if ((lpwndpos->flags & SWP_FRAMECHANGED) == SWP_FRAMECHANGED)
	{
		m_Impl.OnWindowPosChanged (lpwndpos);
	}

	CDialog::OnWindowPosChanged(lpwndpos);

#ifndef _BCGSUITE_
	if (m_Impl.m_pShadow != NULL)
	{
		m_Impl.m_pShadow->Repos();
	}
#endif
}
//**************************************************************************
LRESULT CBCGPDialog::OnChangeVisualManager (WPARAM, LPARAM)
{
	m_Impl.OnChangeVisualManager ();
	return 0;
}
//**************************************************************************
void CBCGPDialog::OnGetMinMaxInfo(MINMAXINFO FAR* lpMMI) 
{
	m_Impl.OnGetMinMaxInfo (lpMMI);
	CDialog::OnGetMinMaxInfo(lpMMI);
}
//****************************************************************************
LRESULT CBCGPDialog::OnSetText(WPARAM, LPARAM) 
{
	LRESULT	lRes = Default();

	if (lRes && IsVisualManagerStyle () && IsVisualManagerNCArea ())
	{
		RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW);
	}

	return lRes;
}
//****************************************************************************
int CBCGPDialog::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	m_Impl.m_bDisableShadows = m_bDisableShadows;

	if (CDialog::OnCreate(lpCreateStruct) == -1)
	{
		return -1;
	}
	
    return m_Impl.OnCreate();
}
//****************************************************************************
void CBCGPDialog::EnableDragClientArea(BOOL bEnable)
{
	m_Impl.EnableDragClientArea(bEnable);
}
//****************************************************************************
void CBCGPDialog::EnableLayout(BOOL bEnable, CRuntimeClass* pRTC, BOOL bResizeBox)
{
	if (IsLightBox())
	{
		bResizeBox = FALSE;
	}

	m_Impl.EnableLayout(bEnable, pRTC, bResizeBox);
}
//************************************************************************
void CBCGPDialog::SetWhiteBackground(BOOL bSet)
{
	ASSERT_VALID (this);

	m_Impl.m_bIsWhiteBackground = bSet && !globalData.IsHighContastMode();
}
//****************************************************************************
void CBCGPDialog::SetGroupBoxesDrawByParent(BOOL bSet /*= TRUE*/)
{
	if (GetSafeHwnd() != NULL)
	{
		ASSERT(FALSE);
		return;
	}

	m_Impl.m_bGroupBoxesDrawByParent = bSet;
}
//****************************************************************************
void CBCGPDialog::OnShowWindow(BOOL bShow, UINT nStatus)
{
	m_Impl.LoadPlacement();
	CDialog::OnShowWindow(bShow, nStatus);
	
	if (!bShow && IsBackstageMode())
	{
		m_Impl.SetUnderlineKeyboardShortcuts(FALSE, FALSE);
	}
}
//****************************************************************************
BOOL CBCGPDialog::OnSetPlacement(WINDOWPLACEMENT& wp)
{
	return m_Impl.SetPlacement(wp);
}
//***************************************************************************
LRESULT CBCGPDialog::OnPowerBroadcast(WPARAM wp, LPARAM)
{
	LRESULT lres = Default ();

	if (wp == PBT_APMRESUMESUSPEND && AfxGetMainWnd()->GetSafeHwnd() == GetSafeHwnd())
	{
		globalData.Resume ();
#ifdef _BCGSUITE_
		bcgpGestureManager.Resume();
#endif
	}

	return lres;
}
//***************************************************************************
BOOL CBCGPDialog::Create(UINT nIDTemplate, CWnd* pParentWnd) 
{
#if _MSC_VER < 1400
	return Create(MAKEINTRESOURCE(nIDTemplate), pParentWnd);
#else
	if (m_bIsLocal && m_pLocaRes == NULL)
	{
		m_pLocaRes = new CBCGPLocalResource ();
	}

	return CDialog::Create(MAKEINTRESOURCE(nIDTemplate), pParentWnd);
#endif
}
//***************************************************************************
BOOL CBCGPDialog::Create(LPCTSTR lpszTemplateName, CWnd* pParentWnd) 
{
	if (m_bIsLocal && m_pLocaRes == NULL)
	{
		m_pLocaRes = new CBCGPLocalResource ();
	}
	
	return CDialog::Create(lpszTemplateName, pParentWnd);
}
//***************************************************************************
void CBCGPDialog::OnOK()
{
	if (m_pParentEdit != NULL && !m_bParentClosePopupDlgNotified)
	{
		m_pParentEdit->ClosePopupDlg(NULL, TRUE);
		m_bParentClosePopupDlgNotified = TRUE;
		return;
	}

	if (!IsBackstageMode())
	{
		CDialog::OnOK();
	}
}
//***************************************************************************
void CBCGPDialog::OnCancel()
{
	if (m_Impl.m_pToolTipInfo->GetSafeHwnd () != NULL)
	{
		m_Impl.m_pToolTipInfo->Pop();
	}

	if (m_pParentEdit != NULL && !m_bParentClosePopupDlgNotified)
	{
		m_pParentEdit->ClosePopupDlg(NULL, FALSE);
		m_bParentClosePopupDlgNotified = TRUE;
		return;
	}

	if (!IsBackstageMode())
	{
		CDialog::OnCancel();
	}
	else
	{
		CWnd* pParent = GetParent();
		if (pParent->GetSafeHwnd() != NULL)
		{
			pParent->SendMessage(WM_CLOSE);
		}
	}
}
//***************************************************************************
void CBCGPDialog::ClosePopupDlg(LPCTSTR lpszEditValue, BOOL bOK, DWORD_PTR dwUserData)
{
	if (m_pParentEdit == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	if (!m_bParentClosePopupDlgNotified)
	{
		m_pParentEdit->ClosePopupDlg(lpszEditValue, bOK, dwUserData);
		m_bParentClosePopupDlgNotified = TRUE;
	}
}
//***************************************************************************
void CBCGPDialog::OnSysCommand(UINT nID, LPARAM lParam)
{
	m_Impl.OnSysCommand(nID, lParam);
	CDialog::OnSysCommand(nID, lParam);
}
//***************************************************************************
void CBCGPDialog::EnableExpand(UINT nExpandCheckBoxCtrlID, LPCTSTR lpszExpandLabel, LPCTSTR lszCollapseLabel)
{
	EnableExpand(GetDlgItem(nExpandCheckBoxCtrlID)->GetSafeHwnd(), lpszExpandLabel, lszCollapseLabel);
}
//***************************************************************************
void CBCGPDialog::EnableExpand(HWND hwndExpandCheckBoxCtrl, LPCTSTR lpszExpandLabel, LPCTSTR lszCollapseLabel)
{
	ASSERT(GetSafeHwnd() != NULL);

	if (GetLayout() != NULL)
	{
		ASSERT(FALSE);
		return;
	}

	if (!m_bIsExpanded)
	{
		Expand();
	}

	m_strExpandLabel = lpszExpandLabel == NULL ? _T("") : lpszExpandLabel;
	m_strCollapseLabel = lszCollapseLabel == NULL ? _T("") : lszCollapseLabel;

	m_lstCtrlsInCollapseArea.RemoveAll();
	m_nExpandedAreaHeight = -1;
	m_nExpandedCheckBoxHeight = -1;

	m_hwndExpandCheckBoxCtrl = hwndExpandCheckBoxCtrl;

	if (hwndExpandCheckBoxCtrl == NULL)
	{
		return;
	}

	BOOL bIsCollapsedArea = FALSE;

	for (CWnd* pWndChild = GetWindow(GW_CHILD); pWndChild != NULL; pWndChild = pWndChild->GetWindow(GW_HWNDNEXT))
	{
		HWND hwndChild = pWndChild->GetSafeHwnd();

		if (bIsCollapsedArea)
		{
			m_lstCtrlsInCollapseArea.AddTail(hwndChild);
		}
		else if (hwndChild == hwndExpandCheckBoxCtrl)
		{
			CBCGPButton* pButton = DYNAMIC_DOWNCAST(CBCGPButton, FromHandlePermanent(hwndExpandCheckBoxCtrl));
			if (pButton->GetSafeHwnd() == NULL)
			{
				m_btnExpandCheckBox.SubclassWindow(hwndExpandCheckBoxCtrl);
				pButton = &m_btnExpandCheckBox;
			}

			if (pButton->GetSafeHwnd() != NULL)
			{
				if (!m_strCollapseLabel.IsEmpty())
				{
					pButton->SetWindowText(m_strCollapseLabel);
				}

				pButton->m_bCheckBoxExpandCollapseMode = TRUE;
				pButton->SetMouseCursorHand();
				pButton->SizeToContent();
				pButton->SetCheck(1);

				CRect rectExpand;
				pButton->GetWindowRect(rectExpand);
				ScreenToClient(rectExpand);
				
				CRect rectClient;
				GetClientRect(rectClient);
				
				m_nExpandedAreaHeight = rectClient.bottom - rectExpand.bottom - EXPAND_MARGIN_VERT;
				m_nExpandedCheckBoxHeight = rectExpand.Height();
			}

			bIsCollapsedArea = TRUE;
		}
	}
}
//***************************************************************************
void CBCGPDialog::SetExpandAreaSpecialBackground(BOOL bSet, BOOL bRedraw)
{
	m_bExpandAreaSpecialBackground = bSet;

	if (bRedraw && GetSafeHwnd() != NULL)
	{
		RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
	}
}
//***************************************************************************
BOOL CBCGPDialog::Expand(BOOL bExpand)
{
	ASSERT(GetSafeHwnd() != NULL);

	if (m_lstCtrlsInCollapseArea.IsEmpty() || m_nExpandedAreaHeight <= 0)
	{
		return FALSE;
	}

	if (m_bIsExpanded == bExpand)
	{
		return TRUE;
	}

	OnBeforeExpand();

	CRect rectWindow;
	GetWindowRect(rectWindow);

	int dy = bExpand ? m_nExpandedAreaHeight : -m_nExpandedAreaHeight;

	SetWindowPos(NULL, -1, -1, rectWindow.Width(), rectWindow.Height() + dy, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

	for (POSITION pos = m_lstCtrlsInCollapseArea.GetHeadPosition(); pos != NULL;)
	{
		CWnd* pCtrl = CWnd::FromHandle(m_lstCtrlsInCollapseArea.GetNext(pos));

		if (!bExpand && GetFocus()->GetSafeHwnd() == pCtrl->GetSafeHwnd() && m_hwndExpandCheckBoxCtrl != NULL)
		{
			::SetFocus(m_hwndExpandCheckBoxCtrl);
		}

		pCtrl->ShowWindow(bExpand ? SW_SHOWNOACTIVATE : SW_HIDE);
	}

	m_bIsExpanded = bExpand;

	CString strLabel = m_bIsExpanded ? m_strCollapseLabel : m_strExpandLabel;

	CBCGPButton* pButton = DYNAMIC_DOWNCAST(CBCGPButton, CWnd::FromHandlePermanent(m_hwndExpandCheckBoxCtrl));
	ASSERT_VALID(pButton);

	pButton->SetCheck(m_bIsExpanded);

	if (!strLabel.IsEmpty())
	{
		pButton->SetWindowText(strLabel);
		pButton->SizeToContent();
	}

	RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);

	OnAfterExpand();
	return TRUE;
}

#if (!defined _BCGSUITE_) && (!defined BCGP_EXCLUDE_RIBBON)

void CBCGPDialog::OnDrawRibbonBackgroundImage(CDC* pDC, CRect rect)
{
	CFrameWnd* pTopFrame = GetTopLevelFrame();
	CBCGPRibbonBar* pRibbonBar = NULL;

	CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST (CBCGPMDIFrameWnd, pTopFrame);
	if (pMainFrame != NULL)
	{
		pRibbonBar = pMainFrame->GetRibbonBar();
	}
	else	// Maybe, SDI frame...
	{
		CBCGPFrameWnd* pFrame = DYNAMIC_DOWNCAST (CBCGPFrameWnd, pTopFrame);
		if (pFrame != NULL)
		{
			pRibbonBar = pFrame->GetRibbonBar();
		}
	}

	if (pRibbonBar != NULL)
	{
		ASSERT_VALID(pRibbonBar);

		const CBCGPToolBarImages& image = pRibbonBar->GetBackgroundImage();

		if (!globalData.IsHighContastMode() && CBCGPVisualManager::GetInstance ()->IsRibbonBackgroundImage() && image.GetCount() > 0)
		{
			CRect rectRibbon;
			pRibbonBar->GetClientRect(rectRibbon);
			pRibbonBar->MapWindowPoints(this, &rectRibbon);
			
			rectRibbon.bottom = rectRibbon.top + rect.Height();

			((CBCGPToolBarImages&)image).DrawEx(pDC, rectRibbon, 0, CBCGPToolBarImages::ImageAlignHorzRight, CBCGPToolBarImages::ImageAlignVertTop);
		}
	}
}

#endif

BOOL CBCGPDialog::OnNeedTipText(UINT id, NMHDR* pNMH, LRESULT* pResult)
{
	return m_Impl.OnNeedTipText(id, pNMH, pResult);
}
//**************************************************************************
LRESULT CBCGPDialog::OnBCGUpdateToolTips (WPARAM wp, LPARAM)
{
	UINT nTypes = (UINT) wp;

	if (nTypes & BCGP_TOOLTIP_TYPE_DEFAULT)
	{
		m_Impl.CreateTooltipInfo();
	}

	return 0;
}
//**************************************************************************
BOOL CBCGPDialog::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	if (m_Impl.OnSetCursor())
	{
		return TRUE;
	}
	
	return CDialog::OnSetCursor(pWnd, nHitTest, message);
}
//*****************************************************************************************
LRESULT CBCGPDialog::OnDPIChanged(WPARAM, LPARAM)
{
	LRESULT lRes = Default();
	
	OnChangeVisualManager(0, 0);
	return lRes;
}
//*****************************************************************************************
LRESULT CBCGPDialog::OnThemeChanged(WPARAM, LPARAM)
{
	LRESULT lRes = Default();
	
	CBCGPVisualManager::GetInstance()->OnUpdateSystemColors();
	OnChangeVisualManager(0, 0);
	return lRes;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPLightBoxDialog dialog

IMPLEMENT_DYNAMIC(CBCGPLightBoxDialog, CBCGPDialog)

void CBCGPLightBoxDialogButton::OnFillBackground(CDC* pDC, const CRect& rectClient)
{
	CRect rect = rectClient;
	globalData.DrawParentBackground(this, pDC, rect);
}
//*****************************************************************************************
void CBCGPLightBoxDialogButton::OnDrawBorder(CDC* /*pDC*/, CRect& /*rectClient*/, UINT /*uiState*/)
{
}

#ifdef _BCGSUITE_
	#define CBCGPDockManager	CDockingManager
	#define CBCGPMiniFrameWnd	CPaneFrameWnd
#endif

class CBCGPLightBoxShadow : public CBCGPWnd  
{
	friend class CBCGPLightBoxDialog;

public:
	CBCGPLightBoxShadow();
	virtual ~CBCGPLightBoxShadow();
	
	void DoRepos();
	CBCGPLightBoxDialog* GetLightBoxDialog();

	void HideFloatingPanes(CBCGPDockManager* pDM);
	void RestoreFloatingPanes();

	void DisableChildWindows(CWnd* pWnd);
	void DisableRibbonControls();
	
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBCGPLightBoxShadow)
	public:
	virtual BOOL Create(CBCGPLightBoxDialog* pDlg);
	protected:
	//}}AFX_VIRTUAL
	
	//{{AFX_MSG(CBCGPLightBoxShadow)
	afx_msg void OnPaint();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnNcDestroy();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
		
	CFrameWnd*			m_pParent;
	CList<HWND, HWND>	m_lstDisabled;
	CList<HWND, HWND&>	m_lstHiddenFloatingBars;
	CRect				m_rectDlg;
	HWND				m_hWndRibbon;
};

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPLightBoxShadow::CBCGPLightBoxShadow()
{
	m_pParent = NULL;
	m_rectDlg.SetRectEmpty();
	m_hWndRibbon = NULL;
}

CBCGPLightBoxShadow::~CBCGPLightBoxShadow()
{
}

BEGIN_MESSAGE_MAP(CBCGPLightBoxShadow, CBCGPWnd)
//{{AFX_MSG_MAP(CBCGPLightBoxShadow)
	ON_WM_PAINT()
	ON_WM_LBUTTONDOWN()
	ON_WM_NCDESTROY()
	ON_WM_ERASEBKGND()
	ON_WM_MOUSEACTIVATE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BOOL CBCGPLightBoxShadow::Create(CBCGPLightBoxDialog* pDlg) 
{
	ASSERT_VALID(pDlg);

	DWORD dwStyle = WS_POPUP | MFS_MOVEFRAME | MFS_SYNCACTIVE | MFS_BLOCKSYSMENU;

	m_pParent = DYNAMIC_DOWNCAST(CFrameWnd, AfxGetMainWnd());
	if (m_pParent == NULL)
	{
		ASSERT(FALSE);
		return FALSE;
	}

#ifndef _BCGSUITE_
	CFrameWnd* pTearOffFrame = DYNAMIC_DOWNCAST(CFrameWnd, CWnd::FromHandlePermanent(CBCGPMDIFrameWnd::m_hwndActiveDetachedMDIFrame));
	if (pTearOffFrame != NULL)
	{
		ASSERT_VALID(pTearOffFrame);
		m_pParent = pTearOffFrame;
	}
#endif

	CString strClassName = ::AfxRegisterWndClass (
		CS_SAVEBITS,
		::LoadCursor(NULL, IDC_ARROW),
		(HBRUSH)(COLOR_BTNFACE + 1), NULL);
	

	CRect rect(0, 0, 0, 0);
	if (!CBCGPWnd::CreateEx(WS_EX_TOOLWINDOW | WS_EX_LAYERED, strClassName, _T(""), dwStyle, rect, m_pParent, 0))
	{
		return FALSE;
	}

	globalData.SetLayeredAttrib (GetSafeHwnd (), 0, max(1, pDlg->m_LightBoxOptions.m_nAlpha), LWA_ALPHA);
	
	DisableChildWindows(m_pParent);

	BOOL bIsValidParentFrame = FALSE;

	CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST (CBCGPMDIFrameWnd, m_pParent);
	if (pMainFrame != NULL)
	{
		if (!pMainFrame->SetLightBoxDialog(pDlg))
		{
			return FALSE;
		}

#ifndef _BCGSUITE_
		HideFloatingPanes(pMainFrame->GetDockManager());
#else
		HideFloatingPanes(pMainFrame->GetDockingManager());
#endif
		bIsValidParentFrame = TRUE;
	}
	else	// Maybe, SDI frame...
	{
		CBCGPFrameWnd* pFrame = DYNAMIC_DOWNCAST (CBCGPFrameWnd, m_pParent);
		if (pFrame != NULL)
		{
			if (!pFrame->SetLightBoxDialog(pDlg))
			{
				return FALSE;
			}

#ifndef _BCGSUITE_
			HideFloatingPanes(pFrame->GetDockManager());
#else
			HideFloatingPanes(pFrame->GetDockingManager());
#endif
			bIsValidParentFrame = TRUE;
		}
	}
	
	if (!bIsValidParentFrame)
	{
		ASSERT(FALSE);
		return FALSE;
	}
	
	ShowWindow(SW_SHOWNOACTIVATE);
	return TRUE;
}
//***********************************************************************************************************
void CBCGPLightBoxShadow::DisableChildWindows(CWnd* pWnd)
{
	if (pWnd->GetSafeHwnd() == NULL)
	{
		return;
	}

	for (CWnd* pWndChild = pWnd->GetWindow(GW_CHILD); pWndChild != NULL; pWndChild = pWndChild->GetWindow(GW_HWNDNEXT))
	{
		if (pWndChild->IsWindowEnabled() && pWndChild->IsWindowVisible() && pWndChild != this)
		{
#ifndef BCGP_EXCLUDE_RIBBON
			CBCGPRibbonBar* pRibbonBar = DYNAMIC_DOWNCAST(CBCGPRibbonBar, pWndChild);
			if (pRibbonBar != NULL && pRibbonBar->IsReplaceFrameCaption() && !pRibbonBar->IsKindOf(RUNTIME_CLASS(CBCGPRibbonStatusBar)))
			{
				m_hWndRibbon = pRibbonBar->GetSafeHwnd();
				DisableChildWindows(pRibbonBar);
			}
			else
#endif
			{
				pWndChild->EnableWindow(FALSE);
				m_lstDisabled.AddTail(pWndChild->GetSafeHwnd());
			}
		}
	}
}
//***********************************************************************************************************
void CBCGPLightBoxShadow::DisableRibbonControls()
{
	if (m_hWndRibbon != NULL)
	{
		DisableChildWindows(CWnd::FromHandlePermanent(m_hWndRibbon));
	}
}
//***********************************************************************************************************
CBCGPLightBoxDialog* CBCGPLightBoxShadow::GetLightBoxDialog()
{
	CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST (CBCGPMDIFrameWnd, m_pParent);
	if (pMainFrame != NULL)
	{
		return pMainFrame->GetLightBoxDialog();
	}
	else	// Maybe, SDI frame...
	{
		CBCGPFrameWnd* pFrame = DYNAMIC_DOWNCAST (CBCGPFrameWnd, m_pParent);
		if (pFrame != NULL)
		{
			return pFrame->GetLightBoxDialog();
		}
	}

	return NULL;
}
//***********************************************************************************************************
void CBCGPLightBoxShadow::HideFloatingPanes(CBCGPDockManager* pDM)
{
	ASSERT_VALID(pDM);

	const CObList& lstMiniFrames = pDM->GetMiniFrames();
	for (POSITION pos = lstMiniFrames.GetHeadPosition (); pos != NULL;)
	{
		CWnd* pWndNext = (CWnd*) lstMiniFrames.GetNext (pos);

		HWND hWndNext = pWndNext->GetSafeHwnd ();
		if (::IsWindow (hWndNext) && ::IsWindowVisible(hWndNext))
		{
			::ShowWindow(hWndNext, SW_HIDE);
			if (m_lstHiddenFloatingBars.Find (hWndNext) == NULL)
			{
				m_lstHiddenFloatingBars.AddTail (hWndNext);
			}
		}
	}
}
//***********************************************************************************************************
void CBCGPLightBoxShadow::RestoreFloatingPanes()
{
	for (POSITION pos = m_lstHiddenFloatingBars.GetHeadPosition(); pos != NULL;)
	{
		HWND hWndNext = m_lstHiddenFloatingBars.GetNext(pos);
		if (::IsWindow(hWndNext))
		{
			CWnd* pWndNext = CWnd::FromHandle(hWndNext);
			
			CBCGPMiniFrameWnd* pMiniFrame = DYNAMIC_DOWNCAST(CBCGPMiniFrameWnd, pWndNext);
#ifndef _BCGSUITE_
			if (pMiniFrame != NULL && pMiniFrame->GetControlBarCount() > 0)
#else
			if (pMiniFrame != NULL && pMiniFrame->GetPaneCount() > 0)
#endif
			{
				::ShowWindow (hWndNext, SW_SHOWNOACTIVATE);
			}
		}
	}

	m_lstHiddenFloatingBars.RemoveAll ();
}
//***********************************************************************************************************
void CBCGPLightBoxShadow::DoRepos()
{
	if (m_pParent->GetSafeHwnd() == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	if (m_pParent->IsIconic())
	{
		return;
	}

	CRect rect;
	m_pParent->GetClientRect(rect);
	m_pParent->ClientToScreen(rect);

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_hWndRibbon != NULL && ::IsWindow(m_hWndRibbon))
	{
		CBCGPRibbonBar* pRibbonBar = DYNAMIC_DOWNCAST(CBCGPRibbonBar, CWnd::FromHandlePermanent(m_hWndRibbon));
		if (pRibbonBar != NULL 
#ifndef _BCGSUITE_
			&& !pRibbonBar->IsInAutoHideMode()
#endif
		)
		{
			rect.top += pRibbonBar->GetCaptionHeight();
		}
	}
#endif

	SetWindowPos(NULL, rect.left, rect.top, max(0, rect.Width()), max(0, rect.Height()), SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOCOPYBITS);

	DisableRibbonControls();
}
//***********************************************************************************************************
void CBCGPLightBoxShadow::OnPaint() 
{
	CPaintDC dcPaint(this); // device context for painting
	CBCGPMemDC memDC(dcPaint, this);
	CDC* pDC = &memDC.GetDC();

	CBCGPLightBoxDialogOptions options;
	int nShadowOffset = globalUtils.ScaleByDPI(8);

	CBCGPLightBoxDialog* pDlg = GetLightBoxDialog();
	if (pDlg != NULL)
	{
		ASSERT_VALID(pDlg);

		options = pDlg->m_LightBoxOptions;
		
		if (pDlg->IsAnimated())
		{
			nShadowOffset = (int)(pDlg->GetAnimatedValue() * nShadowOffset);
		}
	}

	CRect rectClient;
	GetClientRect(rectClient);

	pDC->FillSolidRect(rectClient, options.m_clrShading);

	if (options.m_bDrawFrameShadow && nShadowOffset > 2)
	{
		COLORREF clrShadow = CBCGPDrawManager::ColorMakeDarker(options.m_clrShading, .3);
		int nBrightness = CBCGPDrawManager::IsLightColor(options.m_clrShading) ? 80 : 50;

#ifndef _BCGSUITE_
		CBCGPShadowRenderer shadow;
		if (shadow.Create(nShadowOffset, clrShadow, 0, 100 - nBrightness, CBCGPToolBarImages::IsRTL()))
		{
			DWORD dwFlags = 0;
			CRect rect = m_rectDlg;

			switch (options.m_Location)
			{
			case BCGPLightBoxDialogLocation_Center:
				dwFlags = CBCGPShadowRenderer::BCGP_FRAME_SIDE_ALL;
				rect.InflateRect(nShadowOffset / 2, nShadowOffset / 2, nShadowOffset, nShadowOffset);
				break;

			case BCGPLightBoxDialogLocation_Left:
				dwFlags = CBCGPShadowRenderer::BCGP_FRAME_SIDE_RIGHT;
				rect.OffsetRect(nShadowOffset, 0);
				break;

			case BCGPLightBoxDialogLocation_Right:
				dwFlags = CBCGPShadowRenderer::BCGP_FRAME_SIDE_LEFT;
				rect.OffsetRect(-nShadowOffset, 0);
				break;

			case BCGPLightBoxDialogLocation_Top:
				dwFlags = CBCGPShadowRenderer::BCGP_FRAME_SIDE_BOTTOM;
				rect.OffsetRect(0, nShadowOffset);
				break;

			case BCGPLightBoxDialogLocation_Bottom:
				dwFlags = CBCGPShadowRenderer::BCGP_FRAME_SIDE_TOP;
				rect.OffsetRect(0, -nShadowOffset);
				break;
			}

			shadow.SetFrameSides(dwFlags);
			shadow.DrawFrame (pDC, rect);
		}
#else
		CBCGPDrawManager dm(*pDC);
		dm.DrawShadow(m_rectDlg, nShadowOffset, 100, nBrightness, NULL, NULL, clrShadow, options.m_Location != BCGPLightBoxDialogLocation_Right);
#endif
	}
}
//***********************************************************************************************************
void CBCGPLightBoxShadow::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CBCGPWnd::OnLButtonDown(nFlags, point);

	CBCGPLightBoxDialog* pDlg = GetLightBoxDialog();
	if (pDlg != NULL)
	{
		ASSERT_VALID(pDlg);
		
		if (!pDlg->IsAnimated() && pDlg->m_LightBoxOptions.m_bCloseLightBoxByClickingOutside)
		{
			pDlg->OnCancel();
		}

		return;
	}

	SendMessage(WM_CLOSE);
}
//***********************************************************************************************************
void CBCGPLightBoxShadow::OnNcDestroy() 
{
	if (m_pParent != NULL)
	{
		CBCGPLightBoxDialog* pDlg = GetLightBoxDialog();
		if (pDlg != NULL)
		{
			pDlg->m_pShadow = NULL;

			if (pDlg->m_bNotifyClose)
			{
				CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST (CBCGPMDIFrameWnd, m_pParent);
				if (pMainFrame != NULL)
				{
					pMainFrame->SetLightBoxDialog(NULL);
				}
				else	// Maybe, SDI frame...
				{
					CBCGPFrameWnd* pFrame = DYNAMIC_DOWNCAST (CBCGPFrameWnd, m_pParent);
					if (pFrame != NULL)
					{
						pFrame->SetLightBoxDialog(NULL);
					}
				}

				RestoreFloatingPanes();
			}
		}
	}

	while (!m_lstDisabled.IsEmpty())
	{
		HWND hwndChild = m_lstDisabled.RemoveHead();

		if (::IsWindow(hwndChild))
		{
			::EnableWindow(hwndChild, TRUE);
		}
	}

	CBCGPWnd::OnNcDestroy();
	delete this;
}
//***********************************************************************************************************
BOOL CBCGPLightBoxShadow::OnEraseBkgnd(CDC* /*pDC*/) 
{
	return TRUE;
}
//***********************************************************************************************************
int CBCGPLightBoxShadow::OnMouseActivate(CWnd* /*pDesktopWnd*/, UINT /*nHitTest*/, UINT /*message*/) 
{
	return MA_NOACTIVATE;
}

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPLightBoxDialog::CBCGPLightBoxDialog(UINT uiDlgResID, CWnd* pParent, CBCGPLightBoxDialogOptions* pOptions/* = NULL*/)
{
	UNREFERENCED_PARAMETER(pParent);

	CommonInit(pOptions);
	m_lpszDlgTemplateName = MAKEINTRESOURCE(uiDlgResID);
}
//***********************************************************************************************************
CBCGPLightBoxDialog::CBCGPLightBoxDialog(LPCTSTR lpszDlgResID, CWnd* pParent, CBCGPLightBoxDialogOptions* pOptions/* = NULL*/)
{
	UNREFERENCED_PARAMETER(pParent);

	CommonInit(pOptions);
	m_lpszDlgTemplateName = lpszDlgResID;
}
//***********************************************************************************************************
CBCGPLightBoxDialog::CBCGPLightBoxDialog(CSize sizeDlg, LPCTSTR lpszCaption, CBCGPLightBoxDialogOptions* pOptions /*= NULL*/)
{
	CommonInit(pOptions);
	m_sizeDlg = sizeDlg;

	if (lpszCaption != NULL && m_LightBoxOptions.m_CaptionStyle != BCGPLightBoxDialogCaptionStyle_None)
	{
		m_strCaption = lpszCaption;
	}
}
//***********************************************************************************************************
CBCGPLightBoxDialog::~CBCGPLightBoxDialog()
{
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::CommonInit(CBCGPLightBoxDialogOptions* pOptions)
{
	m_nCaptionHeight = 0;
	m_bActive = FALSE;
	m_bIsModalLoop = FALSE;
	m_bIsVertScrollBar = FALSE;
	m_bIsHorzScrollBar = FALSE;
	m_pShadow = NULL;
	m_sizeDlg = CSize(0, 0);
	m_sizeCurr = CSize(0, 0);
	m_sizeDlgAnim = CSize(-1, -1);
	m_ptOffsetDlgtAnim = CPoint(0, 0);
	m_lpszDlgTemplateName = NULL;
	m_bHasCloseButton = FALSE;
	m_rectClose.SetRectEmpty();
	m_bIsCloseButtonPressed = FALSE;
	m_bIsCloseButtonHighlighted = FALSE;
	m_bIsCloseButtonCaptured = FALSE;
	m_CurrTransitionType = BCGPLightBoxDialogTransitionType_None;
	m_bNotifyClose = TRUE;
	m_bIsShowAnimation = FALSE;
	m_bIsClosing = FALSE;

	if (pOptions != NULL)
	{
		m_LightBoxOptions = *pOptions;
	}
}

BEGIN_MESSAGE_MAP(CBCGPLightBoxDialog, CBCGPDialog)
//{{AFX_MSG_MAP(CBCGPLightBoxDialog)
	ON_WM_NCDESTROY()
	ON_WM_MOUSEACTIVATE()
	ON_WM_NCCREATE()
	ON_WM_NCACTIVATE()
	ON_WM_HSCROLL()
	ON_WM_VSCROLL()
	ON_WM_MOUSEWHEEL()
	ON_WM_NCCALCSIZE()
	ON_WM_NCHITTEST()
	ON_WM_NCPAINT()
	ON_WM_NCMOUSEMOVE()
	ON_WM_LBUTTONUP()
	ON_WM_MOUSEMOVE()
	ON_WM_CANCELMODE()
	ON_WM_LBUTTONDOWN()
	ON_WM_SIZE()
	//}}AFX_MSG_MAP
	ON_COMMAND(IDOK, OnOK)
	ON_COMMAND(IDCANCEL, OnCancel)
	ON_MESSAGE(WM_FLOATSTATUS, OnFloatStatus)
	ON_MESSAGE(WM_SETTEXT, OnSetText)
END_MESSAGE_MAP()

INT_PTR CBCGPLightBoxDialog::DoModal()
{
	// Create shadow window:
	m_pShadow = new CBCGPLightBoxShadow;
	
	if (!m_pShadow->Create(this))
	{
		return IDCANCEL;
	}
	
	BOOL bIsCreated = FALSE;
	m_bIsModalLoop = TRUE;

	if (m_lpszDlgTemplateName == NULL)
	{
		CBCGPLocalResource localRes;
		bIsCreated = CBCGPDialog::Create(MAKEINTRESOURCE(IDD_BCGBARRES_MSG_BOX));
	}
	else
	{
		bIsCreated = CBCGPDialog::Create(m_lpszDlgTemplateName);
	}

	if (!bIsCreated)
	{
		m_pShadow->DestroyWindow();
		m_pShadow = NULL;
		return IDCANCEL;
	}

	CommonCreate();
	INT_PTR nRes = (INT_PTR)RunModalLoop();

	DestroyWindow();
	return nRes;
}
//***********************************************************************************************************
BOOL CBCGPLightBoxDialog::PreTranslateMessage(MSG* pMsg)
{
	if (IsAnimated())
	{
		if ((pMsg->message >= WM_MOUSEFIRST && pMsg->message <= WM_MOUSELAST) ||
			(pMsg->message >= WM_KEYFIRST && pMsg->message <= WM_KEYLAST))
		{
			return TRUE;
		}
	}

	return CBCGPDialog::PreTranslateMessage(pMsg);
}
//***********************************************************************************************************
BOOL CBCGPLightBoxDialog::Create() 
{
	// Create shadow window:
	m_pShadow = new CBCGPLightBoxShadow;

	if (!m_pShadow->Create(this))
	{
		delete m_pShadow;
		m_pShadow = NULL;

		return FALSE;
	}

	if (m_lpszDlgTemplateName == NULL)
	{
		CBCGPLocalResource localRes;

		if (!CBCGPDialog::Create(MAKEINTRESOURCE(IDD_BCGBARRES_MSG_BOX)))
		{
			return FALSE;
		}
	}
	else
	{
		if (!CBCGPDialog::Create(m_lpszDlgTemplateName))
		{
			return FALSE;
		}
	}

	CommonCreate();
	return TRUE;
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::CommonCreate()
{
	if ((GetStyle() & WS_CAPTION) && m_lpszDlgTemplateName != NULL && m_LightBoxOptions.m_CaptionStyle != BCGPLightBoxDialogCaptionStyle_None)
	{
		GetWindowText(m_strCaption);
	}

	ModifyStyle(WS_BORDER | WS_CAPTION | DS_MODALFRAME | WS_THICKFRAME, MFS_SYNCACTIVE);
	ModifyStyleEx(WS_EX_DLGMODALFRAME, WS_EX_TOOLWINDOW);

	if (m_lpszDlgTemplateName != NULL && m_sizeDlg == CSize(0, 0))
	{
		CRect rectClient;
		GetClientRect(rectClient);
		
		m_sizeDlg = rectClient.Size();
	}

	if (GetStyle() & WS_VSCROLL)
	{
		m_bIsVertScrollBar = TRUE;

		SCROLLINFO si;
		
		ZeroMemory(&si, sizeof (SCROLLINFO));
		si.cbSize = sizeof(SCROLLINFO);

		si.fMask = SIF_RANGE | SIF_POS | SIF_PAGE;
		
		SetScrollInfo(SB_VERT, &si, TRUE);

		m_sizeDlg.cx += GetSystemMetrics(SM_CXVSCROLL);
	}

	if (GetStyle() & WS_HSCROLL)
	{
		m_bIsHorzScrollBar = TRUE;

		SCROLLINFO si;
		
		ZeroMemory(&si, sizeof (SCROLLINFO));
		si.cbSize = sizeof(SCROLLINFO);
		
		si.fMask = SIF_RANGE | SIF_POS | SIF_PAGE;
		
		SetScrollInfo(SB_HORZ, &si, TRUE);

		m_sizeDlg.cy += GetSystemMetrics(SM_CYHSCROLL);
	}

	if (!m_strCaption.IsEmpty())
	{
		m_nCaptionHeight = GetCaptionHeight();
		m_sizeDlg.cy += m_nCaptionHeight;

		CMenu* pSysMenu = GetSystemMenu (FALSE);
		if (pSysMenu->GetSafeHmenu() != NULL)
		{
			MENUITEMINFO menuInfo;
			ZeroMemory(&menuInfo,sizeof(MENUITEMINFO));
			menuInfo.cbSize = sizeof(MENUITEMINFO);
			menuInfo.fMask = MIIM_STATE;
			
			pSysMenu->GetMenuItemInfo(SC_CLOSE, &menuInfo);
			BOOL bCloseIsDisabled =	((menuInfo.fState & MFS_GRAYED) || (menuInfo.fState & MFS_DISABLED));

			if (!bCloseIsDisabled)
			{
				m_bHasCloseButton = TRUE;
			}
		}
	}
	else
	{
		m_nCaptionHeight = 1;
	}

	ModifyStyle(WS_SYSMENU, 0);

	if (IsTransitionAvailable(TRUE))
	{
		m_CurrTransitionType = m_LightBoxOptions.m_ShowTransitionType;
		if (m_CurrTransitionType == BCGPLightBoxDialogTransitionType_Default)
		{
			if (m_LightBoxOptions.m_Location == BCGPLightBoxDialogLocation_Center)
			{
				m_CurrTransitionType = BCGPLightBoxDialogTransitionType_Appear;
			}
			else
			{
				m_CurrTransitionType = BCGPLightBoxDialogTransitionType_Slide;
			}
		}

		if (m_CurrTransitionType == BCGPLightBoxDialogTransitionType_Slide)
		{
			m_sizeDlgAnim = CSize(0, 0);
		}

		m_bIsShowAnimation = TRUE;
		StartAnimation(0., 1., m_LightBoxOptions.m_dblShowTransitionSpeed, m_LightBoxOptions.m_ShowTransitionAnimation);
	}

	ShowWindow(SW_SHOWNOACTIVATE);
	DoRepos();

	if (!IsAnimated())
	{
		RecalcClosebutton();
	}
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::Dismiss()
{
	if (IsAnimated())
	{
		return;
	}

	if (IsTransitionAvailable(FALSE))
	{
		m_CurrTransitionType = m_LightBoxOptions.m_DismissTransitionType;
		if (m_CurrTransitionType == BCGPLightBoxDialogTransitionType_Default)
		{
			if (m_LightBoxOptions.m_Location == BCGPLightBoxDialogLocation_Center)
			{
				m_CurrTransitionType = BCGPLightBoxDialogTransitionType_Appear;	// TODO
			}
			else
			{
				m_CurrTransitionType = BCGPLightBoxDialogTransitionType_Slide;
			}
		}

		if (m_CurrTransitionType == BCGPLightBoxDialogTransitionType_Slide)
		{
			m_sizeDlgAnim = m_sizeDlg;
		}

		m_bIsShowAnimation = FALSE;
		StartAnimation(1., 0., m_LightBoxOptions.m_dblDismissTransitionSpeed, m_LightBoxOptions.m_DismissTransitionAnimation);

		CWinThread* pCurrThread = ::AfxGetThread ();
		if (pCurrThread == NULL)
		{
			ASSERT (FALSE);
			return;
		}
		
		MSG msg;
		
		BOOL bIdle = TRUE;
		LONG lIdleCount = 0;
		
		while (IsAnimated())
		{
			while (bIdle && !::PeekMessage(&msg, NULL, NULL, NULL, PM_NOREMOVE))
			{
				bIdle = pCurrThread->OnIdle(lIdleCount++);
			}
			
			while (::PeekMessage(&msg, NULL, NULL, NULL, PM_NOREMOVE))
			{
				if (!pCurrThread->PumpMessage())
				{
					AfxPostQuitMessage(0);
					return;
				}
				
				if (pCurrThread->IsIdleMessage(&msg))
				{
					bIdle = TRUE;
					lIdleCount = 0;
				}
			}
		}
	}
}
//***********************************************************************************************************
int CBCGPLightBoxDialog::GetCaptionHeight()
{
#ifndef _BCGSUITE_
	int nCaptionHeight = GetCaptionStyle() ==  BCGPLightBoxDialogCaptionStyle_Large ? globalData.GetCaptionTextHeight() : globalData.GetTextHeight();
#else
	int nCaptionHeight = globalData.GetTextHeight();
#endif
	return nCaptionHeight + globalUtils.ScaleByDPI(4);
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnDrawCaption(CDC* pDC, CRect rectCaption, CRect rectCloseButton, COLORREF clrText)
{
	ASSERT_VALID(pDC);

	CRect rectText = rectCaption;
	rectText.right -= rectCloseButton.Width();
	rectText.DeflateRect(globalUtils.ScaleByDPI(8), 0);
	
	COLORREF clrTextOld = pDC->SetTextColor(clrText);

	pDC->DrawText(m_strCaption, rectText, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);
	
	if (!rectCloseButton.IsRectEmpty())
	{
		OnDrawCloseButton(pDC, rectCloseButton, m_bIsCloseButtonHighlighted, m_bIsCloseButtonPressed);
	}

	pDC->SetTextColor(clrTextOld);
}
//***********************************************************************************************************
BCGPLightBoxDialogCaptionStyle CBCGPLightBoxDialog::GetCaptionStyle() const
{
	BCGPLightBoxDialogCaptionStyle captionStyle = m_LightBoxOptions.m_CaptionStyle;
	if (captionStyle == BCGPLightBoxDialogCaptionStyle_Default)
	{
#ifndef _BCGSUITE_
		captionStyle = CBCGPVisualManager::GetInstance()->UseLargeCaptionFontInDockingCaptions() ? BCGPLightBoxDialogCaptionStyle_Large : BCGPLightBoxDialogCaptionStyle_Regular;
#else
		captionStyle = BCGPLightBoxDialogCaptionStyle_Regular;
#endif
	}

	return captionStyle;
}
//***********************************************************************************************************
BOOL CBCGPLightBoxDialog::CloseMe(BOOL bNotify)
{
	if (GetSafeHwnd() != NULL && ::IsWindow(GetSafeHwnd()))
	{
		if (m_bIsModalLoop)
		{
			CBCGPRAII<double> raiiAnim(m_LightBoxOptions.m_dblDismissTransitionSpeed, 0.0);
			CBCGPRAII<BOOL> raiiNotify(m_bNotifyClose, bNotify);
			
			OnCancel();
			
			if (!IsClosing())
			{
				return FALSE;
			}

			if (m_pShadow->GetSafeHwnd() != NULL)
			{
				m_pShadow->DestroyWindow();
				m_pShadow = NULL;
			}
		}
		else
		{
			DestroyWindow();
		}
	}

	return TRUE;
}
//***********************************************************************************************************

#ifndef BCGP_EXCLUDE_RIBBON

class CBCGPRibbonBarUnderLightBox : public CBCGPRibbonBar
{
	friend class CBCGPLightBoxDialog;
};

#endif

BOOL CBCGPLightBoxDialog::ShouldCloseOnClick(CWnd* pWnd, CPoint ptScreen)
{
#ifndef BCGP_EXCLUDE_RIBBON
	CBCGPRibbonBar* pRibbonBar = DYNAMIC_DOWNCAST(CBCGPRibbonBar, pWnd);
	if (pRibbonBar != NULL && pRibbonBar->IsWindowVisible() && pRibbonBar->IsReplaceFrameCaption())
	{
		CRect rectRibbon;
		pRibbonBar->GetWindowRect(rectRibbon);
		
		if (rectRibbon.PtInRect(ptScreen))
		{
			CPoint pt = ptScreen;
			pRibbonBar->ScreenToClient(&pt);

			CBCGPBaseRibbonElement* pHit = pRibbonBar->HitTest(pt);
			if (pHit != NULL && DYNAMIC_DOWNCAST(CBCGPRibbonCaptionButton, pHit) == NULL)
			{
				return TRUE;
			}
		}
	}

#endif
	return FALSE;
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::RecalcClosebutton()
{
	if (m_bHasCloseButton && GetSafeHwnd() != NULL && ::IsWindow(GetSafeHwnd()))
	{
		CRect rect;
		GetWindowRect(rect);
		
		rect.bottom = rect.top + m_nCaptionHeight;

		m_rectClose = rect;
		
		if (GetExStyle() & WS_EX_LAYOUTRTL)
		{
			m_rectClose.right = m_rectClose.left + m_rectClose.Height();
		}
		else
		{
			m_rectClose.left = m_rectClose.right - m_rectClose.Height();
		}
	}
}
//***********************************************************************************************************
BOOL CBCGPLightBoxDialog::IsDrawBorder() const
{
	return m_LightBoxOptions.m_clrBorder != CLR_NONE;
}
//***********************************************************************************************************
BOOL CBCGPLightBoxDialog::IsTransitionAvailable(BOOL bShow) const
{
	if (!IsAnimationSupportedByOS())
	{
		return FALSE;
	}

	if (bShow)
	{
		return m_LightBoxOptions.m_dblShowTransitionSpeed > 0 && m_LightBoxOptions.m_ShowTransitionType != BCGPLightBoxDialogTransitionType_None;
	}
	else
	{
		return m_LightBoxOptions.m_dblDismissTransitionSpeed > 0 && m_LightBoxOptions.m_DismissTransitionType != BCGPLightBoxDialogTransitionType_None;
	}
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnDrawCloseButton(CDC* pDC, CRect rectClose, BOOL bIsHighlighted, BOOL bIsPressed)
{
	if (IsAnimated())
	{
		bIsHighlighted = bIsPressed = FALSE;
	}

	if (IsVisualManagerStyle())
	{
		CBCGPVisualManager::GetInstance()->OnDrawLightBoxDialogCloseButton(pDC, this, rectClose, bIsHighlighted, bIsPressed);
	}
	else
	{
		CBCGPVisualManager::GetInstance()->CBCGPVisualManager::OnDrawLightBoxDialogCloseButton(pDC, this, rectClose, bIsHighlighted, bIsPressed);
	}
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnAnimationValueChanged(double dblOldValue, double dblNewValue)
{
	CBCGPAnimationManager::OnAnimationValueChanged(dblOldValue, dblNewValue);

	if (m_pShadow->GetSafeHwnd() != NULL)
	{
		globalData.SetLayeredAttrib(m_pShadow->GetSafeHwnd(), 0, (BYTE)bcg_clamp((int)(dblNewValue * 1.5 * m_LightBoxOptions.m_nAlpha), 1, m_LightBoxOptions.m_nAlpha), LWA_ALPHA);
	}

	switch (m_CurrTransitionType)
	{
	case BCGPLightBoxDialogTransitionType_Slide:
		m_sizeDlgAnim.cx = (int)(0.5 + dblNewValue * m_sizeDlg.cx);	
		m_sizeDlgAnim.cy = (int)(0.5 + dblNewValue * m_sizeDlg.cy);	
		break;

	case BCGPLightBoxDialogTransitionType_Appear:
		{
			int delta = (int)((1.0 - dblNewValue) * globalUtils.ScaleByDPI(40.0));

			switch (m_LightBoxOptions.m_Location)
			{
			case BCGPLightBoxDialogLocation_Right:
			case BCGPLightBoxDialogLocation_Center:
				m_ptOffsetDlgtAnim.x = delta;
				break;

			case BCGPLightBoxDialogLocation_Left:
				m_ptOffsetDlgtAnim.x = -delta;
				break;

			case BCGPLightBoxDialogLocation_Top:
				m_ptOffsetDlgtAnim.y = -delta;
				break;

			case BCGPLightBoxDialogLocation_Bottom:
				m_ptOffsetDlgtAnim.y = delta;
				break;
			}
		}
		break;
	}

	DoRepos();
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnAnimationFinished()
{
	m_sizeDlgAnim = CSize(-1, -1);
	m_ptOffsetDlgtAnim = CPoint(0, 0);

	if (m_bIsShowAnimation)
	{
		DoRepos();
		SetWindowPos(NULL, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);

		RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_FRAME);

		if (m_pShadow->GetSafeHwnd() != NULL)
		{
			m_pShadow->RedrawWindow();
		}
	}
}
//***********************************************************************************************************
BOOL CBCGPLightBoxDialog::OnScroll(UINT nScrollCode, UINT nPos, BOOL bDoScroll /*= TRUE*/)
{
	int nStep = globalUtils.ScaleByDPI(10);

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
		x -= nStep;
		break;
	case SB_LINEDOWN:
		x += nStep;
		break;
	case SB_PAGEUP:
		x -= m_sizeCurr.cx;
		break;
	case SB_PAGEDOWN:
		x += m_sizeCurr.cx;
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
		y -= nStep;
		break;
	case SB_LINEDOWN:
		y += nStep;
		break;
	case SB_PAGEUP:
		y -= m_sizeCurr.cy;
		break;
	case SB_PAGEDOWN:
		y += m_sizeCurr.cy;
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
//***********************************************************************************************************
BOOL CBCGPLightBoxDialog::OnScrollBy(CSize sizeScroll, BOOL bDoScroll /*= TRUE*/)
{
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
		// do scroll and update scroll positions
		ScrollWindow(-(x-xOrig), -(y-yOrig));
		if (x != xOrig)
			SetScrollPos(SB_HORZ, x);
		if (y != yOrig)
			SetScrollPos(SB_VERT, y);
	}
	return TRUE;
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::DoRepos(BOOL bRedrawShadow)
{
	if (m_pShadow->GetSafeHwnd() == NULL || !::IsWindow(m_hWnd))
	{
		return;
	}

	CRect rectClientDlg;
	GetClientRect(rectClientDlg);

	int nPrevWidth = rectClientDlg.Width();

	m_pShadow->DoRepos();

	CRect rectClient;
	m_pShadow->GetClientRect(rectClient);

	CRect rectDlg = rectClient;

	CSize sizeDlg = m_sizeDlgAnim != CSize(-1, -1) ? m_sizeDlgAnim : m_sizeDlg;

#ifndef _BCGSUITE_
	const int nExtra = CBCGPVisualManager::GetInstance()->IsOwnerDrawCaption() ? 2 : 0;
#else
	const int nExtra = 0;
#endif

	switch (m_LightBoxOptions.m_Location)
	{
	case BCGPLightBoxDialogLocation_Center:
		rectDlg.left = max(rectClient.left, rectClient.CenterPoint().x - sizeDlg.cx / 2);
		rectDlg.right = min(rectClient.right, rectDlg.left + sizeDlg.cx);
		rectDlg.top = max(rectClient.top, rectClient.CenterPoint().y - sizeDlg.cy / 2);
		rectDlg.bottom = min(rectClient.bottom, rectDlg.top + sizeDlg.cy);
		rectDlg.OffsetRect(m_ptOffsetDlgtAnim);
		break;

	case BCGPLightBoxDialogLocation_Left:
		rectDlg.left -= nExtra;
		rectDlg.bottom += nExtra;

		rectDlg.right = min(rectClient.right, rectDlg.left + sizeDlg.cx + m_ptOffsetDlgtAnim.x);
		break;

	case BCGPLightBoxDialogLocation_Right:
		rectDlg.right += nExtra;
		rectDlg.bottom += nExtra;

		rectDlg.left = max(rectClient.left, rectDlg.right - sizeDlg.cx + m_ptOffsetDlgtAnim.x);
		break;

	case BCGPLightBoxDialogLocation_Top:
		rectDlg.left -= nExtra;
		rectDlg.right += nExtra;

		rectDlg.bottom = min(rectClient.bottom, rectDlg.top + sizeDlg.cy + m_ptOffsetDlgtAnim.y);
		break;

	case BCGPLightBoxDialogLocation_Bottom:
		rectDlg.left -= nExtra;
		rectDlg.right += nExtra;
		rectDlg.bottom += nExtra;

		rectDlg.top = max(rectClient.top, rectClient.bottom - sizeDlg.cy + m_ptOffsetDlgtAnim.y);
	}

	m_pShadow->m_rectDlg = rectDlg;
	m_pShadow->ClientToScreen(rectDlg);

	m_sizeCurr = rectDlg.Size();

	if (IsAnimated())
	{
		m_pShadow->RedrawWindow();
	}
	else
	{
		int dx = 0;
		int dy = 0;

		if (m_bIsHorzScrollBar)
		{
			SCROLLINFO si;
			
			ZeroMemory(&si, sizeof (SCROLLINFO));
			si.cbSize = sizeof(SCROLLINFO);
			
			si.fMask = SIF_RANGE | SIF_POS | SIF_PAGE;
			si.nMin = 0;

			int nPos = GetScrollPos(SB_HORZ);

			if (m_sizeDlg.cx > rectDlg.Width())
			{
				si.nMax = m_sizeDlg.cx;
				si.nPage = rectDlg.Width();
				si.nPos = bcg_clamp(nPos, 0, si.nMax - rectDlg.Width());
			}
			
			SetScrollInfo(SB_HORZ, &si, TRUE);
			SetScrollPos(SB_HORZ, si.nPos);

			dx = nPos - si.nPos;
		}

		if (m_bIsVertScrollBar)
		{
			SCROLLINFO si;

			ZeroMemory(&si, sizeof (SCROLLINFO));
			si.cbSize = sizeof(SCROLLINFO);

			si.fMask = SIF_RANGE | SIF_POS | SIF_PAGE;
			si.nMin = 0;

			int nPos = GetScrollPos(SB_VERT);

			if (m_sizeDlg.cy > rectDlg.Height())
			{
				si.nMax = m_sizeDlg.cy;
				si.nPage = rectDlg.Height();
				si.nPos = bcg_clamp(nPos, 0, si.nMax - rectDlg.Height());
			}
			
			SetScrollInfo(SB_VERT, &si, TRUE);
			SetScrollPos(SB_VERT, si.nPos);

			dy = nPos - si.nPos;
		}

		if (dx != 0 || dy != 0)
		{
			ScrollWindow(dx, dy);
		}
	}

	if (!IsAnimated() || m_LightBoxOptions.m_Location == BCGPLightBoxDialogLocation_Center)
	{
		RecalcClosebutton();
	}

	SetWindowPos(&wndTop, rectDlg.left, rectDlg.top, rectDlg.Width(), rectDlg.Height(), 0);

	GetClientRect(rectClientDlg);

	UINT nRedrawFlags = RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN;
	if (rectClientDlg.Width() != nPrevWidth)
	{
		nRedrawFlags |= RDW_FRAME;
	}

	RedrawWindow(NULL, NULL, nRedrawFlags);

	if (bRedrawShadow)
	{
		m_pShadow->RedrawWindow();
	}
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::Activate()
{
	SetWindowPos(&wndTop, -1, -1, -1, -1, SWP_NOMOVE | SWP_NOSIZE);
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnNcDestroy() 
{
	if (m_pShadow->GetSafeHwnd() != NULL)
	{
		m_pShadow->DestroyWindow();
		m_pShadow = NULL;
	}

	CBCGPDialog::OnNcDestroy();

	if (!m_bIsModalLoop)
	{
		delete this;
	}
}
//***********************************************************************************************************
int CBCGPLightBoxDialog::OnMouseActivate(CWnd* /*pDesktopWnd*/, UINT /*nHitTest*/, UINT /*message*/) 
{
	return MA_NOACTIVATE;
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnOK()
{
	if (IsAnimated())
	{
		return;
	}

	Dismiss();

	m_bIsClosing = TRUE;
	CBCGPDialog::OnOK();
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnCancel()
{
	if (IsAnimated())
	{
		return;
	}

	Dismiss();

	m_bIsClosing = TRUE;

	if (m_bIsModalLoop)
	{
		CBCGPDialog::OnCancel();
	}
	else
	{
		DestroyWindow();
	}
}
//***********************************************************************************************************
LRESULT CBCGPLightBoxDialog::OnFloatStatus(WPARAM wParam, LPARAM)
{
	// these asserts make sure no conflicting actions are requested
	ASSERT(!((wParam & FS_SHOW) && (wParam & FS_HIDE)));
	ASSERT(!((wParam & FS_ENABLE) && (wParam & FS_DISABLE)));
	ASSERT(!((wParam & FS_ACTIVATE) && (wParam & FS_DEACTIVATE)));
	
	// FS_SYNCACTIVE is used to detect MFS_SYNCACTIVE windows
	LRESULT lResult = 0;
	if ((GetStyle() & MFS_SYNCACTIVE) && (wParam & FS_SYNCACTIVE))
		lResult = 1;
	
	if (wParam & (FS_SHOW|FS_HIDE))
	{
		SetWindowPos(NULL, 0, 0, 0, 0,
			((wParam & FS_SHOW) ? SWP_SHOWWINDOW : SWP_HIDEWINDOW) | SWP_NOZORDER |
			SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
	}
	if (wParam & (FS_ENABLE|FS_DISABLE))
		EnableWindow((wParam & FS_ENABLE) != 0);
	
	if ((wParam & (FS_ACTIVATE|FS_DEACTIVATE)) &&
		GetStyle() & MFS_SYNCACTIVE)
	{
		ModifyStyle(MFS_SYNCACTIVE, 0);
		SendMessage(WM_NCACTIVATE, (wParam & FS_ACTIVATE) != 0);
		ModifyStyle(0, MFS_SYNCACTIVE);
	}
	
	return lResult;
}
//***********************************************************************************************************
BOOL CBCGPLightBoxDialog::OnNcCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (!CBCGPDialog::OnNcCreate(lpCreateStruct))
		return FALSE;

	ModifyStyle(0, MFS_SYNCACTIVE);
	
	if (GetStyle() & MFS_SYNCACTIVE)
	{
		// syncronize activation state with top level parent
		CWnd* pParentWnd = GetTopLevelParent();
		ASSERT(pParentWnd != NULL);
		CWnd* pActiveWnd = GetForegroundWindow();
		BOOL bActive = (pParentWnd == pActiveWnd) ||
			(pParentWnd->GetLastActivePopup() == pActiveWnd &&
			pActiveWnd->SendMessage(WM_FLOATSTATUS, FS_SYNCACTIVE) != 0);
		
		// the WM_FLOATSTATUS does the actual work
		SendMessage(WM_FLOATSTATUS, bActive ? FS_ACTIVATE : FS_DEACTIVATE);
	}
	
	return TRUE;
}
//***********************************************************************************************************
BOOL CBCGPLightBoxDialog::OnNcActivate(BOOL bActive) 
{
	if ((GetStyle() & MFS_SYNCACTIVE) == 0)
	{
		if (m_bActive != bActive)
		{
			m_bActive = bActive;
			SendMessage(WM_NCPAINT);
		}
	}
	else if (m_nFlags & WF_KEEPMINIACTIVE)
	{
		return FALSE;
	}

	return TRUE;
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	if (pScrollBar != NULL && pScrollBar->SendChildNotifyLastMsg())
		return;     // eat it
	
	// ignore scroll bar msgs from other controls
	if (pScrollBar != GetScrollBarCtrl(SB_HORZ))
		return;
	
	OnScroll(MAKEWORD(nSBCode, -1), nPos);
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	if (pScrollBar != NULL && pScrollBar->SendChildNotifyLastMsg())
		return;     // eat it
	
	// ignore scroll bar msgs from other controls
	if (pScrollBar != GetScrollBarCtrl(SB_VERT))
		return;
	
	OnScroll(MAKEWORD(-1, nSBCode), nPos);
}
//***********************************************************************************************************
BOOL CBCGPLightBoxDialog::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	if (CBCGPPopupMenu::GetActiveMenu () != NULL || IsAnimated())
	{
		return TRUE;
	}
	
	if (nFlags & (MK_SHIFT | MK_CONTROL))
	{
		return FALSE;
	}

	UNUSED_ALWAYS(pt);

	BOOL bIsVertScroll = (GetStyle() & WS_VSCROLL) != 0;
	if (!bIsVertScroll && (GetStyle() & WS_HSCROLL) == 0)
	{
		return FALSE;
	}

	const int nBar = bIsVertScroll ? SB_VERT : SB_HORZ;
	const UINT nTotal = (UINT)(bIsVertScroll ? m_sizeDlg.cy : m_sizeDlg.cx);
	
	SCROLLINFO scrollInfo;
	ZeroMemory(&scrollInfo, sizeof(SCROLLINFO));
	
    scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_ALL;
	
	GetScrollInfo (nBar, &scrollInfo);
	
	if (scrollInfo.nPage > nTotal || scrollInfo.nPos >= scrollInfo.nMax)
	{
		return FALSE;
	}
	
	int nSteps = abs(zDelta) / WHEEL_DELTA;
	if (nSteps > 1)
	{
		int nPos = GetScrollPos(nBar);

		memset (&scrollInfo, 0, sizeof (SCROLLINFO));
		scrollInfo.cbSize = sizeof (SCROLLINFO);
		scrollInfo.fMask = SIF_POS;
		scrollInfo.nPos = (zDelta < 0) ? (nPos + nSteps) : (nPos - nSteps);
		
		SetScrollInfo (nBar, &scrollInfo);

		if (bIsVertScroll)
		{
			OnVScroll(SB_THUMBTRACK, scrollInfo.nPos, NULL);
		}
		else
		{
			OnHScroll(SB_THUMBTRACK, scrollInfo.nPos, NULL);
		}
	}
	else
	{
		if (bIsVertScroll)
		{
			OnVScroll(zDelta < 0 ? SB_LINEDOWN : SB_LINEUP, 0, NULL);
		}
		else
		{
			OnHScroll(zDelta < 0 ? SB_LINERIGHT : SB_LINELEFT, 0, NULL);
		}
	}
	
	return TRUE;
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnNcCalcSize(BOOL bCalcValidRects, NCCALCSIZE_PARAMS FAR* lpncsp) 
{
	CBCGPDialog::OnNcCalcSize(bCalcValidRects, lpncsp);

	lpncsp->rgrc[0].top += m_nCaptionHeight;

	if (IsDrawBorder())
	{
		switch (m_LightBoxOptions.m_Location)
		{
		case BCGPLightBoxDialogLocation_Center:
			lpncsp->rgrc[0].top++;
			lpncsp->rgrc[0].bottom--;
			lpncsp->rgrc[0].left++;
			lpncsp->rgrc[0].right--;
			break;

		case BCGPLightBoxDialogLocation_Right:
			lpncsp->rgrc[0].left++;
			break;
			
		case BCGPLightBoxDialogLocation_Left:
			lpncsp->rgrc[0].right--;
			break;
			
		case BCGPLightBoxDialogLocation_Top:
			lpncsp->rgrc[0].bottom--;
			break;
			
		case BCGPLightBoxDialogLocation_Bottom:
			lpncsp->rgrc[0].top++;
			break;
		}
	}
}
//***********************************************************************************************************
BCGNcHitTestType CBCGPLightBoxDialog::OnNcHitTest(CPoint point) 
{
	BCGNcHitTestType nHitTtest = CBCGPDialog::OnNcHitTest(point);
	if (nHitTtest == HTCAPTION)
	{
		nHitTtest = HTNOWHERE;
	}

	if (!m_rectClose.IsRectEmpty() && m_rectClose.PtInRect(point))
	{
		nHitTtest = HTOBJECT;
	}

	return nHitTtest;
}
//***********************************************************************************************************
void CBCGPLightBoxDialog::OnNcPaint() 
{
	CBCGPDialog::OnNcPaint();

	CWindowDC dc(this);
	
	CRect rect;
	GetWindowRect(rect);

	rect.bottom -= rect.top;
	rect.right -= rect.left;
	rect.left = rect.top = 0;

	if (IsDrawBorder())
	{
		COLORREF clrBorder = m_LightBoxOptions.m_clrBorder;
		if (clrBorder == CLR_DEFAULT)
		{
			if (globalData.IsHighContastMode())
			{
				clrBorder = globalData.clrBtnText;
			}
			else
			{
				clrBorder = IsVisualManagerStyle() ? globalData.clrBarShadow : globalData.clrBtnShadow;
			}
		}

		if (m_LightBoxOptions.m_Location == BCGPLightBoxDialogLocation_Center)
		{
			dc.Draw3dRect(rect, clrBorder, clrBorder);
			rect.DeflateRect(1, 1);
		}
		else
		{
			CPoint pt1(0, 0);
			CPoint pt2(0, 0);

			switch (m_LightBoxOptions.m_Location)
			{
			case BCGPLightBoxDialogLocation_Right:
				pt1 = CPoint(rect.left, rect.top);
				pt2 = CPoint(rect.left, rect.bottom);
				rect.left++;
				break;
				
			case BCGPLightBoxDialogLocation_Left:
				rect.right--;
				pt1 = CPoint(rect.right, rect.top);
				pt2 = CPoint(rect.right, rect.bottom);
				break;
				
			case BCGPLightBoxDialogLocation_Top:
				rect.bottom--;
				pt1 = CPoint(rect.left, rect.bottom);
				pt2 = CPoint(rect.right, rect.bottom);
				break;
				
			case BCGPLightBoxDialogLocation_Bottom:
				pt1 = CPoint(rect.left, rect.top);
				pt2 = CPoint(rect.right, rect.top);
				rect.top++;
				break;
			}

			if (IsAnimated() && (m_LightBoxOptions.m_Location == BCGPLightBoxDialogLocation_Left || m_LightBoxOptions.m_Location == BCGPLightBoxDialogLocation_Top))
			{
				clrBorder = IsVisualManagerStyle() ? globalData.clrBarFace : globalData.clrBtnFace;
			}

			CBCGPPenSelector ps(dc, clrBorder);

			dc.MoveTo(pt1);
			dc.LineTo(pt2);
		}
	}

	if (m_nCaptionHeight == 0)
	{
		return;
	}

	rect.bottom = rect.top + m_nCaptionHeight;

	CRect rectClose;
	rectClose.SetRectEmpty();

	if (!m_rectClose.IsRectEmpty())
	{
		rectClose = rect;
		
		rectClose.left = rectClose.right - m_rectClose.Width();
		rectClose.bottom = rectClose.top + m_rectClose.Height();
	}

	COLORREF clrText = m_LightBoxOptions.m_clrCaption;
	
	if (!IsVisualManagerStyle())
	{
		dc.FillRect(rect, &globalData.brBtnFace);
		
		if (clrText == CLR_DEFAULT)
		{
			clrText = globalData.clrBtnText;
		}
	}
	else
	{
		CBCGPVisualManager::GetInstance()->OnFillDialog(&dc, this, rect);
		
		if (clrText == CLR_DEFAULT)
		{
#if !defined (_BCGSUITE_) && !defined (BCGP_EXCLUDE_RIBBON)
			clrText = CBCGPVisualManager::GetInstance()->GetRibbonBackstageInfoTextColor();
#else
			clrText = globalData.clrBarText;
#endif
		}
	}

	if (!m_strCaption.IsEmpty())
	{
		const BCGPLightBoxDialogCaptionStyle captionStyle = GetCaptionStyle();

		CBCGPFontSelector fs(dc, 
#ifndef _BCGSUITE_
			captionStyle == BCGPLightBoxDialogCaptionStyle_Large ? &globalData.fontCaption : 
#endif
			captionStyle == BCGPLightBoxDialogCaptionStyle_Bold ? &globalData.fontBold : &globalData.fontRegular);

		dc.SetBkMode(TRANSPARENT);

		OnDrawCaption(&dc, rect, rectClose, clrText);
	}
}
//**********************************************************************************
void CBCGPLightBoxDialog::OnNcMouseMove(UINT nHitTest, CPoint point)
{
	if (IsAnimated())
	{
		return;
	}

	if (!m_bIsCloseButtonCaptured)
	{
		if (m_rectClose.PtInRect (point) && m_bHasCloseButton)
		{
			m_bIsCloseButtonHighlighted = TRUE;

			SetCapture ();
			SendMessage(WM_NCPAINT);
		}
	}
	
	CBCGPDialog::OnNcMouseMove(nHitTest, point);
}
//**********************************************************************************
void CBCGPLightBoxDialog::OnLButtonUp(UINT nFlags, CPoint point) 
{
	if (IsAnimated())
	{
		return;
	}

	if (m_bIsCloseButtonCaptured)
	{
		ReleaseCapture ();

		m_bIsCloseButtonPressed = FALSE;
		m_bIsCloseButtonCaptured = FALSE;
		m_bIsCloseButtonHighlighted = FALSE;

		SendMessage(WM_NCPAINT);

		CPoint ptScreen = point;
		ClientToScreen(&ptScreen);

		if (m_rectClose.PtInRect(ptScreen))
		{
			OnCancel();
		}

		return;
	}
	
	CBCGPDialog::OnLButtonUp(nFlags, point);
}
//*************************************************************************************
void CBCGPLightBoxDialog::OnMouseMove(UINT nFlags, CPoint point) 
{
	if (IsAnimated())
	{
		return;
	}

	if (m_bIsCloseButtonCaptured)
	{
		CPoint ptScreen = point;
		ClientToScreen(&ptScreen);
		
		BOOL bIsButtonPressed = m_rectClose.PtInRect(ptScreen);
		if (bIsButtonPressed != m_bIsCloseButtonPressed)
		{
			m_bIsCloseButtonPressed = bIsButtonPressed;
			SendMessage(WM_NCPAINT);
		}

		return;
	}
	
	if (m_bIsCloseButtonHighlighted)
	{
		CPoint ptScreen = point;
		ClientToScreen(&ptScreen);

		if (!m_rectClose.PtInRect(ptScreen))
		{
			m_bIsCloseButtonHighlighted = FALSE;
			ReleaseCapture ();

			SendMessage(WM_NCPAINT);
		}
	}
	
	CBCGPDialog::OnMouseMove(nFlags, point);
}
//*********************************************************************************
void CBCGPLightBoxDialog::OnCancelMode() 
{
	CBCGPDialog::OnCancelMode();
	
	if (m_bHasCloseButton && GetCapture()->GetSafeHwnd() == GetSafeHwnd())
	{
		ReleaseCapture ();
	}

	m_bIsCloseButtonPressed = FALSE;
	m_bIsCloseButtonCaptured = FALSE;
	m_bIsCloseButtonHighlighted = FALSE;

	SendMessage(WM_NCPAINT);
}
//********************************************************************************
void CBCGPLightBoxDialog::OnLButtonDown(UINT nFlags, CPoint point) 
{
	if (IsAnimated())
	{
		return;
	}

	CPoint ptScreen = point;
	ClientToScreen(&ptScreen);

	if (m_rectClose.PtInRect(ptScreen) && m_bHasCloseButton)
	{
		SetFocus ();

		m_bIsCloseButtonPressed = TRUE;
		m_bIsCloseButtonCaptured = TRUE;

		SetCapture ();

		SendMessage(WM_NCPAINT);
		return;
	}

	CBCGPDialog::OnLButtonDown(nFlags, point);
}
//********************************************************************************
void CBCGPLightBoxDialog::OnSize(UINT nType, int cx, int cy) 
{
	CBCGPDialog::OnSize(nType, cx, cy);
	
	if (m_sizeDlg == CSize(0, 0))
	{
		m_sizeDlg.cx = cx;
		m_sizeDlg.cy = cy;
	}
}
//****************************************************************************
LRESULT CBCGPLightBoxDialog::OnSetText(WPARAM wParam, LPARAM lParam) 
{
	LRESULT	lRes = CBCGPDialog::OnSetText(wParam, lParam);

	if (m_LightBoxOptions.m_CaptionStyle != BCGPLightBoxDialogCaptionStyle_None)
	{
		GetWindowText(m_strCaption);
		SendMessage(WM_NCPAINT);
	}

	return lRes;
}
