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

// BCGPVisualManager.cpp: implementation of the CBCGPPVisualManager class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"

#include "BCGPGlobalUtils.h"
#include "BCGPVisualManager.h"
#include "BCGPToolbarButton.h"
#include "BCGPOutlookBarPane.h"
#include "BCGPOutlookButton.h"
#include "BCGGlobals.h"
#include "BCGPDockingControlBar.h"
#include "BCGPBaseControlBar.h"
#include "BCGPToolBar.h"
#include "BCGPTabWnd.h"
#include "BCGPDrawManager.h"
#include "BCGPShowAllButton.h"
#include "BCGPButton.h"
#include "BCGPMiniFrameWnd.h"
#include "BCGPCaptionBar.h"
#include "BCGPOutlookButton.h"
#include "BCGPTasksPane.h"
#include "BCGPSlider.h"
#include "BCGPCalendarBar.h"
#include "MenuImages.h"
#include "BCGPHeaderCtrl.h"
#include "BCGPSpinButtonCtrl.h"
#include "BCGPDockManager.h"
#include "BCGPTabbedControlBar.h"
#include "BCGPAutoHideButton.h"
#include "BCGPToolBox.h"
#include "BCGPPopupWindow.h"
#include "BCGPPlannerManagerCtrl.h"
#include "BCGPPropList.h"
#include "BCGPStatusBar.h"
#include "BCGPRibbonBar.h"
#include "BCGPRibbonPanel.h"
#include "BCGPRibbonCategory.h"
#include "BCGPRibbonButton.h"
#include "BCGPRibbonStatusBarPane.h"
#include "BCGPRibbonHyperlink.h"
#include "BCGPFrameWnd.h"
#include "BCGPMDIFrameWnd.h"
#include "BCGPRibbonEdit.h"
#include "BCGPRibbonLabel.h"
#include "BCGPRibbonSlider.h"
#include "BCGPGridCtrl.h"
#include "BCGPRibbonPaletteButton.h"
#include "BCGPRibbonProgressBar.h"
#include "BCGPToolTipCtrl.h"
#include "BCGPTooltipManager.h"
#include "BCGPRibbonColorButton.h"
#include "BCGPDockBar.h"
#include "BCGPGroup.h"
#include "BCGPSliderCtrl.h"
#include "bcgpstyle.h"
#include "BCGPGanttItem.h"
#include "BCGPGanttChart.h"
#include "BCGPDialogBar.h"
#include "BCGPBreadcrumb.h"
#include "BCGPCalculator.h"
#include "BCGPEdit.h"
#include "BCGPRibbonMainPanel.h"
#include "BCGPToolbarComboBoxButton.h"
#include "BCGPProgressCtrl.h"
#include "BCGPEditListBox.h"
#include "BCGPURLLinkButton.h"
#include "BCGPChartFormat.h"
#include "BCGPCircularGaugeImpl.h"
#include "BCGPLinearGaugeImpl.h"
#include "BCGPAppBarWnd.h"
#include "BCGPWinUITiles.h"
#include "BCGPSplitterWnd.h"
#include "BCGPDateTimeCtrl.h"
#include "BCGPSwitchImpl.h"

#ifndef BCGP_EXCLUDE_PLANNER
#include "BCGPPlannerViewDay.h"
#include "BCGPPlannerViewMonth.h"
#include "BCGPPlannerViewSchedule.h"
#endif

#ifndef BCGP_EXCLUDE_RIBBON
#include "BCGPRibbonProgressBar.h"
#endif

#include "BCGPSVGImage.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE (CBCGPVisualManager, CObject)

extern CObList	gAllToolbars;
extern CBCGPTooltipManager*	g_pTooltipManager;

CBCGPVisualManager*	CBCGPVisualManager::m_pVisManager = NULL;
CRuntimeClass*		CBCGPVisualManager::m_pRTIDefault = NULL;
BOOL				CBCGPVisualManager::m_bCanShowFrameShadows = TRUE;
BOOL				CBCGPVisualManager::m_bMDIAreaFullRedrawOnActivate = TRUE;
CString				CBCGPVisualManager::m_strStylePath;

UINT BCGM_CHANGEVISUALMANAGER = ::RegisterWindowMessage (_T("BCGM_CHANGEVISUALMANAGER"));

#ifndef HDS_FILTERBAR
	#define HDS_FILTERBAR  0x0100
#endif

#ifndef HDS_CHECKBOXES
	#define HDS_CHECKBOXES	0x0400
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPVisualManager::CBCGPVisualManager(BOOL bIsTemporary)
{
	m_bAutoDestroy = FALSE;
	m_bIsTemporary = bIsTemporary;

	if (!bIsTemporary)
	{
		if (m_pVisManager != NULL)
		{
			ASSERT (FALSE);
		}
		else
		{
			m_pVisManager = this;
		}
	}

	m_bLook2000 = FALSE;
	m_bMenuFlatLook = FALSE;
	m_nMenuShadowDepth = 6;
	m_bShadowHighlightedImage = FALSE;
	m_bEmbossDisabledImage = TRUE;
	m_bFadeInactiveImage = FALSE;
	m_bEnableToolbarButtonFill = TRUE;

	m_nVertMargin = globalUtils.ScaleByDPI(12);
	m_nHorzMargin = globalUtils.ScaleByDPI(12);
	m_nGroupVertOffset = globalUtils.ScaleByDPI(15);
	m_nGroupCaptionHeight = globalUtils.ScaleByDPI(25);
	m_nGroupCaptionHorzOffset = globalUtils.ScaleByDPI(13);
	m_nGroupCaptionVertOffset = globalUtils.ScaleByDPI(7);
	m_nTasksMinHeight = 0;
	m_nTasksVertMargin = 0;
	m_nTasksHorzOffset = globalUtils.ScaleByDPI(12);
	m_nTasksIconHorzOffset = globalUtils.ScaleByDPI(5);
	m_nTasksIconVertOffset = globalUtils.ScaleByDPI(4);
	m_bActiveCaptions = TRUE;
	m_bIsTransparentExpandingBox = FALSE;

	m_bOfficeXPStyleMenus = FALSE;
	m_nMenuBorderSize = 2;

	m_b3DTabWideBorder = TRUE;
	m_bAlwaysFillTab = FALSE;
	m_bFrameMenuCheckedItems = FALSE;
	m_clrMenuShadowBase = (COLORREF)-1;	// Used in derived classes
	m_clrFrameShadowBaseActive = (COLORREF)-1;
	m_clrFrameShadowBaseInactive = (COLORREF)-1;

	if (!bIsTemporary)
	{
		CBCGPDockManager::m_bSDParamsModified = TRUE;
		CBCGPDockManager::EnableDockBarMenu (FALSE);

		CBCGPAutoHideButton::m_bOverlappingTabs = TRUE;

		globalData.UpdateSysColors ();
	}

	m_bPlannerBackItemToday         = FALSE;
	m_bPlannerBackItemSelected      = FALSE;
	m_bPlannerCaptionBackItemHeader = FALSE;

	m_ptRibbonMainImageOffset = CPoint (0, 0);
	m_clrMainButton = (COLORREF)-1;
	m_bIsRectangulareRibbonTab = FALSE;
	m_bFillRibbonContextTab = TRUE;

	m_dblGrayScaleLumRatio = 1.0;

	m_bQuickAccessToolbarOnTop = TRUE;
	m_pImagesZoom = NULL;

	m_pImagesSwitchOn = m_pImagesSwitchOff = NULL;

	OnUpdateSystemColors ();
}
//*************************************************************************************
CBCGPVisualManager::~CBCGPVisualManager()
{
	if (!m_bIsTemporary)
	{
		m_pVisManager = NULL;
	}

	if (m_pImagesZoom != NULL)
	{
		delete m_pImagesZoom;
		m_pImagesZoom = NULL;
	}

	if (m_pImagesSwitchOn != NULL)
	{
		delete m_pImagesSwitchOn;
		m_pImagesSwitchOn = NULL;
	}

	if (m_pImagesSwitchOff != NULL)
	{
		delete m_pImagesSwitchOff;
		m_pImagesSwitchOff = NULL;
	}
}
//*************************************************************************************
BOOL CBCGPVisualManager::IsDWMCaptionSupported() const
{
	return globalData.DwmIsCompositionEnabled ();
}
//*************************************************************************************
BOOL CBCGPVisualManager::IsSmallSystemBorders() const
{
	return FALSE;
}
//*************************************************************************************
BOOL CBCGPVisualManager::CanShowFrameShadows() const
{
	return CBCGPVisualManager::m_bCanShowFrameShadows && IsSmallSystemBorders() && !globalData.bIsRemoteSession;
}
//*************************************************************************************
void CBCGPVisualManager::ShowFrameShadows(BOOL bShow)
{
	BOOL bCanShow = TRUE;

	CBCGPVisualManager* pThis = CBCGPVisualManager::m_pVisManager;
	if (pThis != NULL)
	{
		bCanShow = pThis->CanShowFrameShadows();
	}

	if (m_bCanShowFrameShadows != bShow)
	{
		m_bCanShowFrameShadows = bShow;

		if (pThis != NULL && bCanShow != pThis->CanShowFrameShadows())
		{
			AdjustFrames ();

			CWnd* pWndMain = AfxGetMainWnd();
			if (DYNAMIC_DOWNCAST(CDialog, pWndMain) != NULL || DYNAMIC_DOWNCAST(CPropertySheet, pWndMain) != NULL || DYNAMIC_DOWNCAST(CBCGPAppBarWnd, pWndMain) != NULL)
			{
				pWndMain->SendMessage(BCGM_CHANGEVISUALMANAGER);
			}
		}
	}
}
//*************************************************************************************
void CBCGPVisualManager::OnUpdateSystemColors ()
{
	m_clrPlannerWork		= RGB (255, 255, 0);
	m_clrPalennerLine		= RGB (128, 128, 128);
	m_clrReportGroupText	= globalData.clrHilite;

	m_clrEditCtrlSelectionBkActive = m_clrEditCtrlSelectionBkInactive = globalData.clrHilite;
	m_clrEditCtrlSelectionTextActive = m_clrEditCtrlSelectionTextInactive = globalData.clrTextHilite;

	if (globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode ())
	{
		double H;
		double S;
		double L;

		CBCGPDrawManager::RGBtoHSL (m_clrEditCtrlSelectionBkInactive, &H, &S, &L);
		m_clrEditCtrlSelectionBkInactive = CBCGPDrawManager::HLStoRGB_ONE (
			H,
			min (1., L * 1.1),
			min (1., S * 1.1));

		m_clrEditCtrlSelectionTextInactive	= CBCGPDrawManager::PixelAlpha (m_clrEditCtrlSelectionTextInactive, 110);
	}

	m_brDlgButtonsArea.DeleteObject();
	m_brDlgButtonsArea.CreateSolidBrush(globalData.clrBarShadow);

	m_clrSliderTic = globalData.clrBarShadow;
	m_clrSliderSelMark = globalData.clrBarText;
	
	if (m_pImagesZoom != NULL)
	{
		delete m_pImagesZoom;
		m_pImagesZoom = NULL;
	}

	m_ActivateFlag.RemoveAll();

	if (m_pImagesSwitchOn != NULL)
	{
		delete m_pImagesSwitchOn;
		m_pImagesSwitchOn = NULL;
	}
	
	if (m_pImagesSwitchOff != NULL)
	{
		delete m_pImagesSwitchOff;
		m_pImagesSwitchOff = NULL;
	}
}
//*************************************************************************************
BOOL CBCGPVisualManager::FillCheckImageList(CImageList& il, COLORREF clrBkgnd, const CArray<int, int>& states, BOOL bIsDisabled) const
{
	if (il.GetSafeHandle() == NULL)
	{
		return FALSE;
	}

	int cx = 0;
	int cy = 0;
	ImageList_GetIconSize(il.GetSafeHandle(), &cx, &cy);

	if (cx == 0 || cy == 0)
	{
		return FALSE;
	}

	const int count = (int)states.GetSize();
	if (count == 0)
	{
		return FALSE;
	}

	CDC* pDcCtrl = CDC::FromHandle(::GetDC(NULL));

	CDC memDC;
	memDC.CreateCompatibleDC(pDcCtrl);

	const CRect memRect(0, 0, cx * count, cy);

	CBitmap memBmp;
	memBmp.Attach(CBCGPDrawManager::CreateBitmap_32(memRect.Size(), NULL));

	::ReleaseDC(NULL, pDcCtrl->GetSafeHdc());

	CBitmap* pOldBitmap = (CBitmap*)memDC.SelectObject(&memBmp);

	if (clrBkgnd == CLR_DEFAULT)
	{
		clrBkgnd = globalData.clrWindow;
	}

	if (clrBkgnd != CLR_NONE)
	{
		memDC.FillSolidRect(memRect, clrBkgnd);
	}

	CRect rect(memRect);
	rect.right = cx;

	for (int i = 0; i < count; i++)
	{
		CBCGPVisualManager::GetInstance()->OnDrawCheckBoxEx(&memDC, rect, states[i], FALSE, FALSE, !bIsDisabled);
		rect.OffsetRect(cx, 0);
	}

	memDC.SelectObject(pOldBitmap);

	CBCGPDrawManager::FillAlpha(memRect, (HBITMAP)memBmp.GetSafeHandle (), 255);

	il.SetImageCount(0);
	il.Add(&memBmp, (CBitmap*)NULL);

	memDC.SelectObject(pOldBitmap);

	return TRUE;
}
//*************************************************************************************
void CBCGPVisualManager::SetDefaultManager (CRuntimeClass* pRTI)
{
	if (pRTI != NULL &&
		!pRTI->IsDerivedFrom (RUNTIME_CLASS (CBCGPVisualManager)))
	{
		ASSERT (FALSE);
		return;
	}

	m_pRTIDefault = pRTI;

	if (m_pVisManager != NULL)
	{
		ASSERT_VALID (m_pVisManager);

		delete m_pVisManager;
		m_pVisManager = NULL;
	}

	globalData.UpdateSysColors ();

	CBCGPDockManager::SetDockMode (BCGP_DT_STANDARD);
	CBCGPTabbedControlBar::ResetTabs ();

	AdjustFrames ();
	AdjustToolbars ();

	CWnd* pWndMain = AfxGetMainWnd();

	if (DYNAMIC_DOWNCAST(CDialog, pWndMain) != NULL || DYNAMIC_DOWNCAST(CPropertySheet, pWndMain) != NULL || DYNAMIC_DOWNCAST(CBCGPAppBarWnd, pWndMain) != NULL)
	{
		pWndMain->SendMessage(BCGM_CHANGEVISUALMANAGER);
	}

	RedrawAll ();

	if (g_pTooltipManager != NULL)
	{
		g_pTooltipManager->UpdateTooltips ();
	}
}
//*************************************************************************************
void CBCGPVisualManager::RedrawAll ()
{
	CWnd* pMainWnd = AfxGetMainWnd ();
	BOOL bIsMainWndRedrawn = FALSE;

	const CList<CFrameWnd*, CFrameWnd*>& lstFrames = 
		CBCGPFrameImpl::GetFrameList ();

	for (POSITION pos = lstFrames.GetHeadPosition (); pos != NULL;)
	{
		CFrameWnd* pFrame = lstFrames.GetNext (pos);

		if (CWnd::FromHandlePermanent (pFrame->m_hWnd) != NULL)
		{
			ASSERT_VALID (pFrame);

			pFrame->RedrawWindow (NULL, NULL,
				RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN | RDW_FRAME);

			if (pFrame->GetSafeHwnd () == pMainWnd->GetSafeHwnd ())
			{
				bIsMainWndRedrawn = TRUE;
			}
		}
	}

	if (!bIsMainWndRedrawn && pMainWnd->GetSafeHwnd () != NULL &&
		CWnd::FromHandlePermanent (pMainWnd->m_hWnd) != NULL)
	{
		pMainWnd->RedrawWindow (NULL, NULL,
			RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN | RDW_FRAME);
	}

	for (POSITION posTlb = gAllToolbars.GetHeadPosition (); posTlb != NULL;)
	{
		CBCGPControlBar* pToolBar = DYNAMIC_DOWNCAST (CBCGPControlBar, gAllToolbars.GetNext (posTlb));
		if (pToolBar != NULL)
		{
			if (CWnd::FromHandlePermanent (pToolBar->m_hWnd) != NULL)
			{
				ASSERT_VALID (pToolBar);
				
				pToolBar->RedrawWindow (NULL, NULL,
					RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
			}
		}
	}

	CBCGPMiniFrameWnd::RedrawAll ();
}
//*************************************************************************************
void CBCGPVisualManager::AdjustToolbars ()
{
	for (POSITION posTlb = gAllToolbars.GetHeadPosition (); posTlb != NULL;)
	{
		CBCGPToolBar* pToolBar = DYNAMIC_DOWNCAST (CBCGPToolBar, gAllToolbars.GetNext (posTlb));
		if (pToolBar != NULL)
		{
			if (CWnd::FromHandlePermanent (pToolBar->m_hWnd) != NULL)
			{
				ASSERT_VALID (pToolBar);
				pToolBar->OnChangeVisualManager ();
			}
		}
	}
}
//*************************************************************************************
void CBCGPVisualManager::AdjustFrames ()
{
	const CList<CFrameWnd*, CFrameWnd*>& lstFrames = 
		CBCGPFrameImpl::GetFrameList ();

	for (POSITION pos = lstFrames.GetHeadPosition (); pos != NULL;)
	{
		CFrameWnd* pFrame = lstFrames.GetNext (pos);

		if (CWnd::FromHandlePermanent (pFrame->m_hWnd) != NULL)
		{
			ASSERT_VALID (pFrame);
			pFrame->SendMessage (BCGM_CHANGEVISUALMANAGER);
		}
	}
}
//*************************************************************************************
void CBCGPVisualManager::OnDrawBarGripper (CDC* pDC, CRect rectGripper, BOOL bHorz,
									   CBCGPBaseControlBar* pBar)
{
	ASSERT_VALID (pDC);

	const COLORREF clrHilite = pBar != NULL && pBar->IsDialogControl () ?
		globalData.clrBtnHilite : globalData.clrBarHilite;
	const COLORREF clrShadow = pBar != NULL && pBar->IsDialogControl () ?
		globalData.clrBtnShadow : globalData.clrBarShadow;

	const BOOL bSingleGripper = m_bLook2000;

	const int iGripperSize = 3;
	const int iGripperOffset = bSingleGripper ? 0 : 1;
	const int iLinesNum = bSingleGripper ? 1 : 2;

	if (bHorz)
	{
		//-----------------
		// Gripper at left:
		//-----------------
		rectGripper.DeflateRect (0, bSingleGripper ? 3 : 2);

		rectGripper.left = iGripperOffset + rectGripper.CenterPoint().x - 
			( iLinesNum*iGripperSize + (iLinesNum-1)*iGripperOffset) / 2;

		rectGripper.right = rectGripper.left + iGripperSize;
 
		for (int i = 0; i < iLinesNum; i ++)
		{
			pDC->Draw3dRect (rectGripper, 
							clrHilite,
							clrShadow);

			if (! bSingleGripper ) 
			{
				//-----------------------------------
				// To look same as MS Office Gripper!
				//-----------------------------------
				pDC->SetPixel (CPoint (rectGripper.left, rectGripper.bottom - 1),
								clrHilite);
			}

			rectGripper.OffsetRect (iGripperSize+1, 0);
		}
	} 
	else 
	{
		//----------------
		// Gripper at top:
		//----------------
		rectGripper.top += iGripperOffset;
		rectGripper.DeflateRect (bSingleGripper ? 3 : 2, 0);

		rectGripper.top = iGripperOffset + rectGripper.CenterPoint().y - 
			( iLinesNum*iGripperSize + (iLinesNum-1)) / 2;

		rectGripper.bottom = rectGripper.top + iGripperSize;

		for (int i = 0; i < iLinesNum; i ++)
		{
			pDC->Draw3dRect (rectGripper,
							clrHilite,
							clrShadow);

			if (!bSingleGripper) 
			{
				//-----------------------------------
				// To look same as MS Office Gripper!
				//-----------------------------------
				pDC->SetPixel (CPoint (rectGripper.right - 1, rectGripper.top),
								clrHilite);
			}

			rectGripper.OffsetRect (0, iGripperSize+1);
		}
	}
}
//*************************************************************************************
void CBCGPVisualManager::SetLook2000 (BOOL bLook2000)
{
	if (!IsLook2000Allowed())
	{
		m_bLook2000 = TRUE;
		return;
	}

	if ((m_bLook2000 && bLook2000) || (!m_bLook2000 && !bLook2000))
	{
		return;
	}

	m_bLook2000 = bLook2000;

	if (AfxGetMainWnd () != NULL)
	{
		AfxGetMainWnd()->RedrawWindow (NULL, NULL,
			RDW_INVALIDATE | RDW_ERASENOW | RDW_ALLCHILDREN | RDW_UPDATENOW | RDW_FRAME);
	}

	RedrawAll ();
}
//*************************************************************************************
void CBCGPVisualManager::OnFillBarBackground (CDC* pDC, CBCGPBaseControlBar* pBar,
									CRect rectClient, CRect rectClip,
									BOOL /*bNCArea*/)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pBar);

	if (pBar->IsOnGlass ())
	{
		pDC->FillSolidRect (rectClient, RGB (0, 0, 0));
		return;
	}

	if (DYNAMIC_DOWNCAST (CReBar, pBar) != NULL ||
		DYNAMIC_DOWNCAST (CReBar, pBar->GetParent ()))
	{
		FillRebarPane (pDC, pBar, rectClient);
		return;
	}

	if (pBar->IsKindOf (RUNTIME_CLASS (CBCGPOutlookBarPane)) && !globalData.IsHighContastMode ())
	{
		CBrush br(globalData.clrBtnShadow);
		pDC->FillRect (rectClient, &br);
		return;
	}

	if (pBar->IsKindOf (RUNTIME_CLASS (CBCGPCaptionBar)))
	{
		CBCGPCaptionBar* pCaptionBar = (CBCGPCaptionBar*) pBar;

		if (pCaptionBar->IsMessageBarMode () || globalData.IsHighContastMode())
		{
			pDC->FillRect (rectClip, &globalData.brBarFace);
		}
		else
		{
			pDC->FillSolidRect	(rectClip, pCaptionBar->m_clrBarBackground == -1 ? 
				globalData.clrBarShadow : pCaptionBar->m_clrBarBackground);
		}
		return;
	}

	if (pBar->IsKindOf (RUNTIME_CLASS (CBCGPPopupMenuBar)))
	{
		CBCGPPopupMenuBar* pMenuBar = (CBCGPPopupMenuBar*) pBar;

		if (pMenuBar->IsDropDownListMode ())
		{
			pDC->FillRect (rectClip, &globalData.brWindow);
			return;
		}
	}

	// By default, control bar background is filled by 
	// the system 3d background color

	pDC->FillRect (rectClip.IsRectEmpty () ? rectClient : rectClip, 
				pBar->IsDialogControl () ?
					&globalData.brBtnFace : &globalData.brBarFace);
}
//*************************************************************************************
void CBCGPVisualManager::OnDrawBarBorder (CDC* pDC, CBCGPBaseControlBar* pBar, CRect& rect)
{
	ASSERT_VALID(pBar);
	ASSERT_VALID(pDC);

	if (pBar->IsFloating ())
	{
		return;
	}

	DWORD dwStyle = pBar->GetBarStyle  ();
	if (!(dwStyle & CBRS_BORDER_ANY))
		return;

	COLORREF clrBckOld = pDC->GetBkColor ();	// FillSolidRect changes it

	const COLORREF clrHilite = pBar->IsOnGlass () ? RGB (0, 0, 0) : pBar->IsDialogControl () ?
		globalData.clrBtnHilite : globalData.clrBarHilite;
	const COLORREF clrShadow = pBar->IsOnGlass () ? RGB (0, 0, 0) : pBar->IsDialogControl () ?
		globalData.clrBtnShadow : globalData.clrBarShadow;

	COLORREF clr = pBar->IsOnGlass () ? RGB (0, 0, 0) : clrHilite;

	if (dwStyle & CBRS_BORDER_LEFT)
		pDC->FillSolidRect(0, 0, 1, rect.Height() - 1, clr);
	if (dwStyle & CBRS_BORDER_TOP)
		pDC->FillSolidRect(0, 0, rect.Width()-1 , 1, clr);
	if (dwStyle & CBRS_BORDER_RIGHT)
		pDC->FillSolidRect(rect.right, 0/*RGL~:1*/, -1,
			rect.Height()/*RGL-: - 1*/, clrShadow);
	if (dwStyle & CBRS_BORDER_BOTTOM)
		pDC->FillSolidRect(0, rect.bottom, rect.Width()-1, -1, clrShadow);

	// if undockable toolbar at top of frame, apply special formatting to mesh
	// properly with frame menu
	if(!pBar->CanFloat ()) 
	{
		pDC->FillSolidRect(0,0,rect.Width(),1, clrShadow);
		pDC->FillSolidRect(0,1,rect.Width(),1, clrHilite);
	}

	if (dwStyle & CBRS_BORDER_LEFT)
		++rect.left;
	if (dwStyle & CBRS_BORDER_TOP)
		++rect.top;
	if (dwStyle & CBRS_BORDER_RIGHT)
		--rect.right;
	if (dwStyle & CBRS_BORDER_BOTTOM)
		--rect.bottom;

	// Restore Bk color:
	pDC->SetBkColor (clrBckOld);
}
//************************************************************************************
void CBCGPVisualManager::OnDrawMenuBorder (CDC* pDC, CBCGPPopupMenu* /*pMenu*/, CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->Draw3dRect (rect, globalData.clrBarLight, globalData.clrBarDkShadow);
	rect.DeflateRect (1, 1);
	pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarShadow);
}
//************************************************************************************
void CBCGPVisualManager::OnDrawMenuShadow (CDC* pPaintDC, const CRect& rectClient, const CRect& /*rectExclude*/,
								int nDepth,  int iMinBrightness,  int iMaxBrightness,  
								CBitmap* pBmpSaveBottom,  CBitmap* pBmpSaveRight,
								BOOL bRTL)
{
	ASSERT_VALID (pPaintDC);
	ASSERT_VALID (pBmpSaveBottom);
	ASSERT_VALID (pBmpSaveRight);

	//------------------------------------------------------
	// Simple draw the shadow, ignore rectExclude parameter:
	//------------------------------------------------------
	CBCGPDrawManager dm (*pPaintDC);
	dm.DrawShadow (rectClient, nDepth, iMinBrightness, iMaxBrightness,
				pBmpSaveBottom, pBmpSaveRight, (COLORREF)-1, !bRTL);
}
//************************************************************************************
void CBCGPVisualManager::OnFillButtonInterior (CDC* pDC,
				CBCGPToolbarButton* pButton, CRect rect,
				CBCGPVisualManager::BCGBUTTON_STATE state)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->IsKindOf (RUNTIME_CLASS (CBCGPShowAllButton)))
	{
		if (state == ButtonsIsHighlighted)
		{
			CBCGPDrawManager dm (*pDC);
			dm.HighlightRect (rect);
		}

		return;
	}

	if (!m_bEnableToolbarButtonFill)
	{
		BOOL bIsPopupMenu = FALSE;

		CBCGPToolbarMenuButton* pMenuButton = 
			DYNAMIC_DOWNCAST (CBCGPToolbarMenuButton, pButton);
		if (pMenuButton != NULL)
		{
			bIsPopupMenu = pMenuButton->GetParentWnd () != NULL &&
				pMenuButton->GetParentWnd ()->IsKindOf (RUNTIME_CLASS (CBCGPPopupMenuBar));
		}

		if (!bIsPopupMenu)
		{
			return;
		}
	}

	if (!pButton->IsKindOf (RUNTIME_CLASS (CBCGPOutlookButton)) &&
		!CBCGPToolBar::IsCustomizeMode () && state != ButtonsIsHighlighted &&
		(pButton->m_nStyle & (TBBS_CHECKED | TBBS_INDETERMINATE)))
	{
		CRect rectDither = rect;
		rectDither.InflateRect (-afxData.cxBorder2, -afxData.cyBorder2);

		CBCGPToolBarImages::FillDitheredRect (pDC, rectDither);
	}
}
//************************************************************************************
COLORREF CBCGPVisualManager::GetToolbarHighlightColor ()
{
	return globalData.clrBarFace;
}
//************************************************************************************
COLORREF CBCGPVisualManager::GetToolbarDisabledTextColor ()
{
	return globalData.clrGrayedText;
}
//************************************************************************************
COLORREF CBCGPVisualManager::GetToolbarEditPromptColor()
{
	return globalData.clrPrompt;
}
//************************************************************************************
COLORREF CBCGPVisualManager::GetHighlightedColor(UINT /*nType*/)
{
	return globalData.IsHighContastMode() ? globalData.clrHilite : globalData.clrActiveCaptionGradient;
}
//************************************************************************************
void CBCGPVisualManager::OnHighlightMenuItem (CDC*pDC, CBCGPToolbarMenuButton* /*pButton*/,
											CRect rect, COLORREF& /*clrText*/)
{
	ASSERT_VALID (pDC);
	pDC->FillRect (rect, &globalData.brHilite);
}
//************************************************************************************
COLORREF CBCGPVisualManager::GetHighlightedMenuItemTextColor (CBCGPToolbarMenuButton* pButton)
{
	ASSERT_VALID (pButton);

	if (pButton->m_nStyle & TBBS_DISABLED)
	{
		return globalData.clrGrayedText;
	}

	return globalData.clrTextHilite;
}
//************************************************************************************
void CBCGPVisualManager::OnHighlightRarelyUsedMenuItems (CDC* pDC, CRect rectRarelyUsed)
{
	ASSERT_VALID (pDC);

	CBCGPDrawManager dm (*pDC);
	dm.HighlightRect (rectRarelyUsed);

	pDC->Draw3dRect (rectRarelyUsed, globalData.clrBarShadow, globalData.clrBarHilite);
}
//************************************************************************************
void CBCGPVisualManager::OnDrawMenuCheck (CDC* pDC, CBCGPToolbarMenuButton* pButton,
		CRect rectCheck, BOOL /*bHighlight*/, BOOL bIsRadio)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	int iImage = bIsRadio ? CBCGPMenuImages::IdRadio : CBCGPMenuImages::IdCheck;
	CBCGPMenuImages::Draw (pDC, (CBCGPMenuImages::IMAGES_IDS) iImage, 
		rectCheck,
		(pButton->m_nStyle & TBBS_DISABLED) ? CBCGPMenuImages::ImageGray : CBCGPMenuImages::ImageBlack);
}
//************************************************************************************
void CBCGPVisualManager::OnDrawMenuItemButton (CDC* pDC, CBCGPToolbarMenuButton* /*pButton*/,
				CRect rectButton, BOOL bHighlight, BOOL /*bDisabled*/)
{
	ASSERT_VALID (pDC);

	CRect rect = rectButton;
	rect.right = rect.left + 1;
	rect.left--;
	rect.DeflateRect (0, bHighlight ? 1 : 4);

	pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarHilite);
}
//************************************************************************************
void CBCGPVisualManager::OnDrawButtonBorder (CDC* pDC,
		CBCGPToolbarButton* pButton, CRect rect, BCGBUTTON_STATE state)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	BOOL bIsOutlookButton = pButton->IsKindOf (RUNTIME_CLASS (CBCGPOutlookButton));
	COLORREF clrDark = bIsOutlookButton ? 
					   globalData.clrBarDkShadow : globalData.clrBarShadow;

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		if (state == ButtonsIsPressed || state == ButtonsIsHighlighted)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect (rect, (COLORREF)-1, clrDark);
		}
		return;
	}

	switch (state)
	{
	case ButtonsIsPressed:
		pDC->Draw3dRect (&rect, clrDark, globalData.clrBarHilite);
		return;

	case ButtonsIsHighlighted:
		pDC->Draw3dRect (&rect, globalData.clrBarHilite, clrDark);
		return;
	}
}
//*************************************************************************************
void CBCGPVisualManager::OnDrawButtonSeparator (CDC* pDC,
		CBCGPToolbarButton* pButton, CRect rect, CBCGPVisualManager::BCGBUTTON_STATE state,
		BOOL /*bHorz*/)
{
	ASSERT_VALID (pButton);

	if (!m_bMenuFlatLook || !pButton->IsDroppedDown ())
	{
		OnDrawButtonBorder (pDC, pButton, rect, state);
	}
}
//*************************************************************************************
COLORREF CBCGPVisualManager::GetSeparatorColor()
{
	return globalData.clrBarShadow;
}
//*************************************************************************************
void CBCGPVisualManager::OnDrawSeparator (CDC* pDC, CBCGPBaseControlBar* pBar,
										 CRect rect, BOOL bHorz)
{
	ASSERT_VALID (pBar);
	ASSERT_VALID (pDC);

	CRect rectSeparator = rect;

	if (bHorz)
	{
		rectSeparator.left += rectSeparator.Width () / 2 - 1;
		rectSeparator.right = rectSeparator.left + 2;
	}
	else
	{
		rectSeparator.top += rectSeparator.Height () / 2 - 1;
		rectSeparator.bottom = rectSeparator.top + 2;
	}

	const COLORREF clrHilite = pBar->IsDialogControl () ?
		globalData.clrBtnHilite : globalData.clrBarHilite;
	const COLORREF clrShadow = pBar->IsDialogControl () ?
		globalData.clrBtnShadow : globalData.clrBarShadow;

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		int x1, x2;
		int y1, y2;

		if (bHorz)
		{
			x1 = x2 = (rect.left + rect.right) / 2;
			y1 = rect.top;
			y2 = rect.bottom - 1;
		}
		else
		{
			y1 = y2 = (rect.top + rect.bottom) / 2;
			x1 = rect.left;
			x2 = rect.right;
		}

		CBCGPDrawManager dm (*pDC);
		dm.DrawLine (x1, y1, x2, y2, clrShadow);
	}
	else
	{
		pDC->Draw3dRect (rectSeparator, clrShadow, clrHilite);
	}
}
//***************************************************************************************
COLORREF CBCGPVisualManager::OnDrawMenuLabel (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBtnFace);

	CRect rectSeparator = rect;
	rectSeparator.top = rectSeparator.bottom - 2;

	pDC->Draw3dRect (rectSeparator, globalData.clrBtnShadow,
									globalData.clrBtnHilite);

	return globalData.clrBtnText;
}
//***************************************************************************************
COLORREF CBCGPVisualManager::OnDrawControlBarCaption (CDC* pDC, CBCGPDockingControlBar* /*pBar*/, 
			BOOL bActive, CRect rectCaption, CRect /*rectButtons*/)
{
	ASSERT_VALID (pDC);

	CBrush br (bActive ? globalData.clrActiveCaption : globalData.clrInactiveCaption);
	pDC->FillRect (rectCaption, &br);

    // get the text color
	return bActive ? globalData.clrCaptionText : globalData.clrInactiveCaptionText;
}
//****************************************************************************************
void CBCGPVisualManager::OnDrawControlBarCaptionText (CDC* pDC, CBCGPDockingControlBar* /*pBar*/, BOOL /*bActive*/, const CString& strTitle, CRect& rectCaption)
{
	ASSERT_VALID (pDC);
	pDC->DrawText (strTitle, rectCaption, DT_LEFT | DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
}
//****************************************************************************************
void CBCGPVisualManager::OnDrawCaptionButton (
						CDC* pDC, CBCGPCaptionButton* pButton, BOOL /*bActive*/,
						BOOL bHorz, BOOL bMaximized, BOOL bDisabled, 
						int nImageID /*= -1*/)
{
	ASSERT_VALID (pDC);
    CRect rc = pButton->GetRect ();

	CRect rectFill = rc;
	rectFill.DeflateRect(1, 1);

	pDC->FillRect(rectFill, &globalData.brBarFace);

	CBCGPMenuImages::IMAGES_IDS id = (CBCGPMenuImages::IMAGES_IDS)-1;
	
	if (nImageID != -1)
	{
		id = (CBCGPMenuImages::IMAGES_IDS)nImageID;
	}
	else
	{
		id = pButton->GetIconID (bHorz, bMaximized);
	}

	CRect rectImage = rc;

	if (pButton->m_bPushed && (pButton->m_bFocused || pButton->m_bDroppedDown))
	{
		rectImage.OffsetRect (1, 1);
	}

	CBCGPMenuImages::IMAGE_STATE imageState;
	
	if (bDisabled)
	{
		imageState = CBCGPMenuImages::ImageGray;
	}
	else
	{
		imageState = CBCGPMenuImages::GetStateByColor(globalData.clrBarFace);
	}

	CBCGPMenuImages::Draw(pDC, id, rectImage, imageState);

	if (!bDisabled)
	{
		if (pButton->m_bPushed && (pButton->m_bFocused || pButton->m_bDroppedDown))
		{
			pDC->Draw3dRect (rc, globalData.clrBarDkShadow, globalData.clrBarLight);
			rc.DeflateRect (1, 1);
			pDC->Draw3dRect (rc, globalData.clrBarDkShadow, globalData.clrBarHilite);
		}
		else if (!m_bLook2000)
		{
			pDC->Draw3dRect (rc, globalData.clrBarLight, globalData.clrBarDkShadow);
			rc.DeflateRect (1, 1);
			pDC->Draw3dRect (rc, globalData.clrBarHilite, globalData.clrBarShadow);
		}
		else if (pButton->m_bFocused || pButton->m_bPushed || pButton->m_bDroppedDown)
		{
			pDC->Draw3dRect (rc, globalData.clrBarHilite, globalData.clrBarShadow);
		}
	}
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawDockingBarScrollButton(CDC* pDC, CBCGPDockingBarScrollButton* pButton)
{
	OnDrawCaptionButton (pDC, pButton, FALSE, FALSE, FALSE, FALSE);
}
//***********************************************************************************
void CBCGPVisualManager::OnEraseTabsArea (CDC* pDC, CRect rect, 
										 const CBCGPBaseTabWnd* /*pTabWnd*/)
{
	ASSERT_VALID (pDC);
	pDC->FillRect (rect, &globalData.brBarFace);
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawTab (CDC* pDC, CRect rectTab,
						int iTab, BOOL bIsActive, const CBCGPBaseTabWnd* pTabWnd)
{
	ASSERT_VALID (pTabWnd);
	ASSERT_VALID (pDC);

	COLORREF clrTab = pTabWnd->GetTabBkColor (iTab);

	CRect rectClip;
	pDC->GetClipBox (rectClip);

	if (!pTabWnd->IsPureFlatTab())
	{
		if (pTabWnd->IsFlatTab ())
		{
			//----------------
			// Draw tab edges:
			//----------------
			#define FLAT_POINTS_NUM	4
			POINT pts [FLAT_POINTS_NUM];

			const int nHalfHeight = pTabWnd->GetTabsHeight () / 2;

			if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM)
			{
				rectTab.bottom --;

				pts [0].x = rectTab.left;
				pts [0].y = rectTab.top;

				pts [1].x = rectTab.left + nHalfHeight;
				pts [1].y = rectTab.bottom;

				pts [2].x = rectTab.right - nHalfHeight;
				pts [2].y = rectTab.bottom;

				pts [3].x = rectTab.right;
				pts [3].y = rectTab.top;
			}
			else
			{
				rectTab.top ++;

				pts [0].x = rectTab.left + nHalfHeight;
				pts [0].y = rectTab.top;

				pts [1].x = rectTab.left;
				pts [1].y = rectTab.bottom;

				pts [2].x = rectTab.right;
				pts [2].y = rectTab.bottom;

				pts [3].x = rectTab.right - nHalfHeight;
				pts [3].y = rectTab.top;

				rectTab.left += 2;
			}

			CBrush* pOldBrush = NULL;

			if (bIsActive && clrTab == (COLORREF)-1)
			{
				clrTab = GetActiveTabBackColor(pTabWnd);
			}

			CBrush br (clrTab);

			if (clrTab != (COLORREF)-1)
			{
				pOldBrush = pDC->SelectObject (&br);
			}

			pDC->Polygon (pts, FLAT_POINTS_NUM);

			if (pOldBrush != NULL)
			{
				pDC->SelectObject (pOldBrush);
			}
		}
		else if (pTabWnd->IsLeftRightRounded ())
		{
			CList<POINT, POINT> pts;

			POSITION posLeft = pts.AddHead (CPoint (rectTab.left, rectTab.top));
			posLeft = pts.InsertAfter (posLeft, CPoint (rectTab.left, rectTab.top + 2));

			POSITION posRight = pts.AddTail (CPoint (rectTab.right, rectTab.top));
			posRight = pts.InsertBefore (posRight, CPoint (rectTab.right, rectTab.top + 2));

			int xLeft = rectTab.left + 1;
			int xRight = rectTab.right - 1;

			int y = 0;

			for (y = rectTab.top + 2; y < rectTab.bottom - 4; y += 2)
			{
				posLeft = pts.InsertAfter (posLeft, CPoint (xLeft, y));
				posLeft = pts.InsertAfter (posLeft, CPoint (xLeft, y + 2));

				posRight = pts.InsertBefore (posRight, CPoint (xRight, y));
				posRight = pts.InsertBefore (posRight, CPoint (xRight, y + 2));

				xLeft++;
				xRight--;
			}

			if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_TOP)
			{
				xLeft--;
				xRight++;
			}

			const int nTabLeft = xLeft - 1;
			const int nTabRight = xRight + 1;

			for (;y < rectTab.bottom - 1; y++)
			{
				posLeft = pts.InsertAfter (posLeft, CPoint (xLeft, y));
				posLeft = pts.InsertAfter (posLeft, CPoint (xLeft + 1, y + 1));

				posRight = pts.InsertBefore (posRight, CPoint (xRight, y));
				posRight = pts.InsertBefore (posRight, CPoint (xRight - 1, y + 1));

				if (y == rectTab.bottom - 2)
				{
					posLeft = pts.InsertAfter (posLeft, CPoint (xLeft + 1, y + 1));
					posLeft = pts.InsertAfter (posLeft, CPoint (xLeft + 3, y + 1));

					posRight = pts.InsertBefore (posRight, CPoint (xRight, y + 1));
					posRight = pts.InsertBefore (posRight, CPoint (xRight - 2, y + 1));
				}

				xLeft++;
				xRight--;
			}

			posLeft = pts.InsertAfter (posLeft, CPoint (xLeft + 2, rectTab.bottom));
			posRight = pts.InsertBefore (posRight, CPoint (xRight - 2, rectTab.bottom));

			LPPOINT points = new POINT [pts.GetCount ()];

			int i = 0;

			for (POSITION pos = pts.GetHeadPosition (); pos != NULL; i++)
			{
				points [i] = pts.GetNext (pos);

				if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_TOP)
				{
					points [i].y = rectTab.bottom - (points [i].y - rectTab.top);
				}
			}

			CRgn rgnClip;
			rgnClip.CreatePolygonRgn (points, (int) pts.GetCount (), WINDING);

			pDC->SelectClipRgn (&rgnClip);

			CBrush br (clrTab == (COLORREF)-1 ? globalData.clrBtnFace : clrTab);
			OnFillTab (pDC, rectTab, &br, iTab, bIsActive, pTabWnd);

			pDC->SelectClipRgn (NULL);

			CPen pen (PS_SOLID, 1, globalData.clrBarShadow);
			CPen* pOldPen = pDC->SelectObject (&pen);

			for (i = 0; i < pts.GetCount (); i++)
			{
				if ((i % 2) != 0)
				{
					int x1 = points [i - 1].x;
					int y1 = points [i - 1].y;

					int x2 = points [i].x;
					int y2 = points [i].y;

					if (x1 > rectTab.CenterPoint ().x && x2 > rectTab.CenterPoint ().x)
					{
						x1--;
						x2--;
					}

					if (y2 >= y1)
					{
						pDC->MoveTo (x1, y1);
						pDC->LineTo (x2, y2);
					}
					else
					{
						pDC->MoveTo (x2, y2);
						pDC->LineTo (x1, y1);
					}
				}
			}

			delete [] points;
			pDC->SelectObject (pOldPen);

			rectTab.left = nTabLeft;
			rectTab.right = nTabRight;
		}
		else	// 3D Tab
		{
			CRgn rgnClip;

			CRect rectClipTab;
			pTabWnd->GetTabsRect (rectClipTab);

			BOOL bIsCutted = FALSE;

			const BOOL bIsOneNote = pTabWnd->IsOneNoteStyle () || pTabWnd->IsVS2005Style ();

			const int nExtra = bIsOneNote ?
				((pTabWnd->IsFirstTab(iTab) || bIsActive || pTabWnd->IsVS2005Style ()) ? 
				0 : rectTab.Height ()) : 0;

			if (rectTab.left + nExtra + 10 > rectClipTab.right ||
				rectTab.right - 10 <= rectClipTab.left)
			{
				return;
			}

			const int iVertOffset = 2;
			const int iHorzOffset = 2;
			const BOOL bIs2005 = pTabWnd->IsVS2005Style ();

			#define POINTS_NUM	8
			POINT pts [POINTS_NUM];

			if (!bIsActive || bIsOneNote || clrTab != (COLORREF)-1 || m_bAlwaysFillTab)
			{
				if (clrTab != (COLORREF)-1 || bIsOneNote || m_bAlwaysFillTab)
				{
					CRgn rgn;
					CBrush br (clrTab == (COLORREF)-1 ? globalData.clrBtnFace : clrTab);

					CRect rectFill = rectTab;

					if (bIsOneNote)
					{
						const int nHeight = rectTab.Height ();

						pts [0].x = rectTab.left;
						pts [0].y = rectTab.bottom;

						pts [1].x = rectTab.left;
						pts [1].y = rectTab.bottom;

						pts [2].x = rectTab.left + 2;
						pts [2].y = rectTab.bottom;
						
						pts [3].x = rectTab.left + nHeight;
						pts [3].y = rectTab.top + 2;
						
						pts [4].x = rectTab.left + nHeight + 4;
						pts [4].y = rectTab.top;
						
						pts [5].x = rectTab.right - 2;
						pts [5].y = rectTab.top;
						
						pts [6].x = rectTab.right;
						pts [6].y = rectTab.top + 2;

						pts [7].x = rectTab.right;
						pts [7].y = rectTab.bottom;

						for (int i = 0; i < POINTS_NUM; i++)
						{
							if (pts [i].x > rectClipTab.right)
							{
								pts [i].x = rectClipTab.right;
								bIsCutted = TRUE;
							}

							if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM)
							{
								pts [i].y = rectTab.bottom - pts [i].y + rectTab.top - 1;
							}
						}

						rgn.CreatePolygonRgn (pts, POINTS_NUM, WINDING);
						pDC->SelectClipRgn (&rgn);
					}
					else
					{
						rectFill.DeflateRect (1, 0);

						if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM)
						{
							rectFill.bottom--;
						}
						else
						{
							rectFill.top++;
						}

						rectFill.right = min (rectFill.right, rectClipTab.right);
					}

					OnFillTab (pDC, rectFill, &br, iTab, bIsActive, pTabWnd);
					pDC->SelectClipRgn (NULL);

					if (bIsOneNote)
					{
						CRect rectLeft;
						pTabWnd->GetClientRect (rectLeft);
						rectLeft.right = rectClipTab.left - 1;

						pDC->ExcludeClipRect (rectLeft);

						if (!pTabWnd->IsFirstTab(iTab) && !bIsActive && iTab != pTabWnd->GetFirstVisibleTabNum ())
						{
							CRect rectLeftTab = rectClipTab;
							rectLeftTab.right = rectFill.left + rectFill.Height () - 10;

							const int nVertOffset = bIs2005 ? 2 : 1;

							if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM)
							{
								rectLeftTab.top -= nVertOffset;
							}
							else
							{
								rectLeftTab.bottom += nVertOffset;
							}

							pDC->ExcludeClipRect (rectLeftTab);
						}

						pDC->Polyline (pts, POINTS_NUM);

						if (bIsCutted)
						{
							pDC->MoveTo (rectClipTab.right, rectTab.top);
							pDC->LineTo (rectClipTab.right, rectTab.bottom);
						}

						CRect rectRight = rectClipTab;
						rectRight.left = rectFill.right;

						pDC->ExcludeClipRect (rectRight);
					}
				}
			}

			COLORREF clrLight = IsDarkTheme() ? globalData.clrBarDkShadow : globalData.clrBarHilite;
			COLORREF clrShadow = IsDarkTheme() ? globalData.clrBarLight : globalData.clrBarShadow;
			COLORREF clrDkShadow = IsDarkTheme() ? globalData.clrBarHilite : globalData.clrBarDkShadow;

			if (!bIsActive && globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode () && clrTab != (COLORREF)-1)
			{
				clrLight = CBCGPDrawManager::PixelAlpha (clrTab, 120);
				clrShadow = clrTab;
				clrDkShadow = CBCGPDrawManager::PixelAlpha (clrTab, 95);
			}

			CPen penLight (PS_SOLID, 1, clrLight);
			CPen penShadow (PS_SOLID, 1, clrShadow);
			CPen penDark (PS_SOLID, 1, clrDkShadow);

			CPen* pOldPen = NULL;

			if (bIsOneNote)
			{
				pOldPen = (CPen*) pDC->SelectObject (&penLight);
				ASSERT(pOldPen != NULL);

				if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM)
				{
					if (!bIsCutted)
					{
						int yTop = bIsActive ? pts [7].y - 1 : pts [7].y;

						pDC->MoveTo (pts [6].x - 1, pts [6].y);
						pDC->LineTo (pts [7].x - 1, yTop);
					}
				}
				else
				{
					pDC->MoveTo (pts [2].x + 1, pts [2].y);
					pDC->LineTo (pts [3].x + 1, pts [3].y);

					pDC->MoveTo (pts [3].x + 1, pts [3].y);
					pDC->LineTo (pts [3].x + 2, pts [3].y);

					pDC->MoveTo (pts [3].x + 2, pts [3].y);
					pDC->LineTo (pts [3].x + 3, pts [3].y);

					pDC->MoveTo (pts [4].x - 1, pts [4].y + 1);
					pDC->LineTo (pts [5].x + 1, pts [5].y + 1);

					if (!bIsActive && !bIsCutted && m_b3DTabWideBorder)
					{
						pDC->SelectObject (&penShadow);

						pDC->MoveTo (pts [6].x - 2, pts [6].y - 1);
						pDC->LineTo (pts [6].x - 1, pts [6].y - 1);
					}

					pDC->MoveTo (pts [6].x - 1, pts [6].y);
					pDC->LineTo (pts [7].x - 1, pts [7].y);
				}
			}
			else
			{
				if (rectTab.right > rectClipTab.right)
				{
					CRect rectTabClip = rectTab;
					rectTabClip.right = rectClipTab.right;

					rgnClip.CreateRectRgnIndirect (&rectTabClip);
					pDC->SelectClipRgn (&rgnClip);
				}

				if (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM)
				{
					pOldPen = (CPen*) pDC->SelectObject (&penLight);
					ASSERT(pOldPen != NULL);

					if (!m_b3DTabWideBorder)
					{
						pDC->SelectObject(&penShadow);
					}

					pDC->MoveTo (rectTab.left, rectTab.top);
					pDC->LineTo (rectTab.left, rectTab.bottom - iVertOffset);

					if (m_b3DTabWideBorder)
					{
						pDC->SelectObject (&penDark);
					}

					pDC->LineTo (rectTab.left + iHorzOffset, rectTab.bottom);
					pDC->LineTo (rectTab.right - iHorzOffset, rectTab.bottom);
					pDC->LineTo (rectTab.right, rectTab.bottom - iVertOffset);
					pDC->LineTo (rectTab.right, rectTab.top - 1);

					pDC->SelectObject(&penShadow);

					if (m_b3DTabWideBorder)
					{
						pDC->MoveTo (rectTab.left + iHorzOffset + 1, rectTab.bottom - 1);
						pDC->LineTo (rectTab.right - iHorzOffset, rectTab.bottom - 1);
						pDC->LineTo (rectTab.right - 1, rectTab.bottom - iVertOffset);
						pDC->LineTo (rectTab.right - 1, rectTab.top - 1);
					}
				}
				else
				{
					pOldPen = pDC->SelectObject (
						m_b3DTabWideBorder ? &penDark : &penShadow);

					ASSERT(pOldPen != NULL);

					pDC->MoveTo (rectTab.right, bIsActive ? rectTab.bottom : rectTab.bottom - 1);
					pDC->LineTo (rectTab.right, rectTab.top + iVertOffset);
					pDC->LineTo (rectTab.right - iHorzOffset, rectTab.top);
					
					if (m_b3DTabWideBorder)
					{
						pDC->SelectObject (&penLight);
					}
					
					pDC->LineTo (rectTab.left + iHorzOffset, rectTab.top);
					pDC->LineTo (rectTab.left, rectTab.top + iVertOffset);

					pDC->LineTo (rectTab.left, rectTab.bottom);

					if (m_b3DTabWideBorder)
					{
						pDC->SelectObject (&penShadow);
						
						pDC->MoveTo (rectTab.right - 1, bIsActive ? rectTab.bottom : rectTab.bottom - 1);
						pDC->LineTo (rectTab.right - 1, rectTab.top + iVertOffset - 1);
					}
 				}
			}

			if (bIsActive)
			{
				const int iBarHeight = 1;
				const int y = (pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM) ? 
					(rectTab.top - iBarHeight - 1) : (rectTab.bottom);

				CRect rectFill (CPoint (rectTab.left, y), 
								CSize (rectTab.Width (), iBarHeight + 1));

				COLORREF clrActiveTab = pTabWnd->GetTabBkColor (iTab);

				if (bIsOneNote)
				{
					if (bIs2005)
					{
						rectFill.left += 3;
					}
					else
					{
						rectFill.OffsetRect (1, 0);
						rectFill.left++;
					}

					if (clrActiveTab == (COLORREF)-1)
					{
						clrActiveTab = IsDarkTheme() ? globalData.clrBarFace : globalData.clrWindow;
					}
				}

				if (clrActiveTab != (COLORREF)-1)
				{
					CBrush br (clrActiveTab);
					pDC->FillRect (rectFill, &br);
				}
				else
				{
					pDC->FillRect (rectFill, &globalData.brBarFace);
				}
			}

			pDC->SelectObject (pOldPen);

			if (bIsOneNote)
			{
				const int nLeftMargin = pTabWnd->IsVS2005Style () && bIsActive ?
					rectTab.Height () * 3 / 4 : rectTab.Height ();

				const int nTabImageMargin = globalUtils.ScaleByDPI(CBCGPBaseTabWnd::TAB_IMAGE_MARGIN);

				const int nRightMargin = pTabWnd->IsVS2005Style () && bIsActive ?
					nTabImageMargin * 3 / 4 : nTabImageMargin;

				rectTab.left += nLeftMargin;
				rectTab.right -= nRightMargin;

				if (pTabWnd->IsVS2005Style () && bIsActive && pTabWnd->HasImage (iTab))
				{
					rectTab.OffsetRect (nTabImageMargin, 0);
				}
			}

			pDC->SelectClipRgn (NULL);
		}
	}

	COLORREF clrText = pTabWnd->GetTabTextColor (iTab);
	COLORREF cltTextOld = (COLORREF)-1;
	if (clrText != (COLORREF)-1)
	{
		cltTextOld = pDC->SetTextColor (clrText);
	}

	if (pTabWnd->IsOneNoteStyle () || pTabWnd->IsVS2005Style ())
	{
		CRect rectTabs;
		pTabWnd->GetTabsRect (rectTabs);

		rectTab.right = min (rectTab.right, rectTabs.right - 2);
	}

	CRgn rgn;
	rgn.CreateRectRgnIndirect (rectClip);

	pDC->SelectClipRgn (&rgn);

	OnDrawTabContent (pDC, rectTab, iTab, bIsActive, pTabWnd, (COLORREF)-1);

	if (cltTextOld != (COLORREF)-1)
	{
		pDC->SetTextColor (cltTextOld);
	}

	pDC->SelectClipRgn (NULL);
}
//*********************************************************************************
void CBCGPVisualManager::OnFillTab (CDC* pDC, CRect rectFill, CBrush* pbrFill,
									 int iTab, BOOL bIsActive, 
									 const CBCGPBaseTabWnd* pTabWnd)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pbrFill);
	ASSERT_VALID (pTabWnd);

	if (pTabWnd->IsPureFlatTab())
	{
		return;
	}

	if (bIsActive && !globalData.IsHighContastMode () &&
		(pTabWnd->IsOneNoteStyle () || pTabWnd->IsVS2005Style () || pTabWnd->IsLeftRightRounded ()) &&
		pTabWnd->GetTabBkColor (iTab) == (COLORREF)-1)
	{
		pDC->FillRect (rectFill, &globalData.brWindow);
	}
	else
	{
		pDC->FillRect (rectFill, pbrFill);
	}
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawTabBorder(CDC* pDC, CBCGPTabWnd* pTabWnd, CRect rectBorder, COLORREF clrBorder,
										 CPen& penLine)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pTabWnd);

	pDC->Draw3dRect (&rectBorder, clrBorder, clrBorder);

	if (!pTabWnd->IsOneNoteStyle ())
	{
		int xRight = rectBorder.right - 1;

		if (!pTabWnd->IsDrawFrame())
		{
			xRight -= pTabWnd->GetTabBorderSize ();
		}

		pDC->SelectObject (&penLine);

		int y = pTabWnd->GetLocation () == CBCGPBaseTabWnd::LOCATION_BOTTOM ? rectBorder.bottom - 1 : rectBorder.top;

		pDC->MoveTo (rectBorder.left, y);
		pDC->LineTo (xRight, y);
	}
}
//*********************************************************************************
BOOL CBCGPVisualManager::OnEraseTabsFrame (CDC* pDC, CRect rect, 
										   const CBCGPBaseTabWnd* pTabWnd) 
{	
	ASSERT_VALID (pTabWnd);
	ASSERT_VALID (pDC);

	COLORREF clrActiveTab = pTabWnd->GetTabBkColor (pTabWnd->GetActiveTab ());

	if (clrActiveTab == (COLORREF)-1)
	{
		return FALSE;
	}

	pDC->FillSolidRect (rect, clrActiveTab);
	return TRUE;
}
//*********************************************************************************
BOOL CBCGPVisualManager::IsTabColorBar(CBCGPTabWnd* pTabWnd, int iTab)
{
	ASSERT_VALID (pTabWnd);
	return pTabWnd->IsPointerStyle() && pTabWnd->GetTabBkColor (iTab) != (COLORREF)-1;
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawTabContent (CDC* pDC, CRect rectTab,
						int iTab, BOOL bIsActive, const CBCGPBaseTabWnd* pTabWnd,
						COLORREF clrText)
{
	ASSERT_VALID (pTabWnd);
	ASSERT_VALID (pDC);

	m_nCurrDrawedTab = iTab;

	if (bIsActive && pTabWnd->GetActiveTabTextColor() != globalData.clrWindowText)
	{
		clrText = pTabWnd->GetActiveTabTextColor();
	}

	if (pTabWnd->GetTabBkColor(iTab) != (COLORREF)-1 && pTabWnd->IsPointerStyle())
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
		
		rectColor.DeflateRect(2, 0);
		
		CBrush br(pTabWnd->GetTabBkColor(iTab));
		pDC->FillRect(rectColor, &br);
	}

	if (pTabWnd->IsTabCloseButton())
	{
		CRect rectClose; rectClose.SetRectEmpty();
		BOOL bIsHighlighted = FALSE;
		BOOL bIsPressed = FALSE;

		if (pTabWnd->GetTabCloseButtonRect(iTab, rectClose, &bIsHighlighted, &bIsPressed) && !rectClose.IsRectEmpty())
		{
			rectTab.right = rectClose.left;

			OnDrawTabCloseButton (pDC, rectClose, pTabWnd, bIsHighlighted, bIsPressed, FALSE /* Disabled */);
		}
	}

	if (pTabWnd->IsFlatTab ())
	{
		//---------------
		// Draw tab text:
		//---------------
		UINT nFormat = DT_SINGLELINE | DT_CENTER | DT_VCENTER;
		if (pTabWnd->IsDrawNoPrefix ())
		{
			nFormat |= DT_NOPREFIX;
		}

		CString strText;
		pTabWnd->GetTabLabel (iTab, strText);

		if (pTabWnd->IsDrawNameInUpperCase())
		{
			strText.MakeUpper();
		}

		COLORREF clrTextOld = (COLORREF)-1;
		if (clrText != (COLORREF)-1)
		{
			clrTextOld = pDC->SetTextColor (clrText);
		}

		pDC->DrawText (strText, rectTab, nFormat);

		if (clrTextOld != (COLORREF)-1)
		{
			pDC->SetTextColor(clrTextOld);
		}
	}
	else
	{
		CSize sizeImage = pTabWnd->GetImageSize ();
		UINT uiIcon = pTabWnd->GetTabIcon (iTab);
		HICON hIcon = pTabWnd->GetTabHicon (iTab);

		if (uiIcon == (UINT)-1 && hIcon == NULL)
		{
			sizeImage.cx = 0;
		}

		const int nIconMargin = globalUtils.ScaleByDPI(CBCGPBaseTabWnd::TAB_IMAGE_MARGIN);
		const BOOL bResizeIcon = sizeImage.cx + 2 * nIconMargin > rectTab.Width ();

		if ((!bResizeIcon || CBCGPBaseTabWnd::m_bAlwaysDrawTabIcon) && rectTab.Width() - 2 * nIconMargin > 0)
		{
			if (hIcon != NULL)
			{
				//---------------------
				// Draw the tab's icon:
				//---------------------
				CRect rectImage = rectTab;

				rectImage.top += (rectTab.Height () - sizeImage.cy) / 2;
				rectImage.bottom = rectImage.top + sizeImage.cy;

				rectImage.left += globalUtils.ScaleByDPI(IMAGE_MARGIN);
				rectImage.right = bResizeIcon ? rectTab.right - globalUtils.ScaleByDPI(IMAGE_MARGIN) : rectImage.left + sizeImage.cx;
				
				if (!pTabWnd->DrawTabState(pDC, iTab, rectImage))
				{
					CSize sizeIcon(0, 0);

					if (pTabWnd->IsIconScaling())
					{
						sizeIcon = sizeImage;
					}

					if (bResizeIcon)
					{
						sizeIcon.cx = rectImage.Width();
					}

					::DrawIconEx (pDC->GetSafeHdc (),
						rectImage.left, rectImage.top, hIcon,
						sizeIcon.cx, sizeIcon.cy, 0, NULL, DI_NORMAL);
				}
			}
			else
			{
				const CImageList* pImageList = pTabWnd->GetImageList ();
				if (pImageList != NULL && uiIcon != (UINT)-1)
				{
					//----------------------
					// Draw the tab's image:
					//----------------------
					CRect rectImage = rectTab;

					rectImage.top += (rectTab.Height () - sizeImage.cy) / 2;
					rectImage.bottom = rectImage.top + sizeImage.cy;

					rectImage.left += globalUtils.ScaleByDPI(IMAGE_MARGIN);

					if (!bResizeIcon)
					{
						rectImage.right = rectImage.left + sizeImage.cx;
					}
					else
					{
						rectImage.right -= globalUtils.ScaleByDPI(IMAGE_MARGIN);
					}

					if (!pTabWnd->DrawTabState(pDC, iTab, rectImage))
					{
						ASSERT_VALID (pImageList);

						if (bResizeIcon)
						{
							HICON hIconTemp = ((CImageList*) pImageList)->ExtractIcon(uiIcon);
							
							::DrawIconEx (pDC->GetSafeHdc (),
								rectImage.left, rectImage.top, hIconTemp,
								rectImage.Width(), rectImage.Height(), 0, NULL, DI_NORMAL);

							::DestroyIcon(hIconTemp);
						}
						else
						{
							((CImageList*) pImageList)->Draw (pDC, uiIcon, rectImage.TopLeft (), ILD_TRANSPARENT);
						}
					}
				}
			}

			//------------------------------
			// Finally, draw the tab's text:
			//------------------------------
			CRect rcText = rectTab;
			rcText.left += sizeImage.cx + 2 * globalUtils.ScaleByDPI(TEXT_MARGIN);

			if (rcText.Width () < sizeImage.cx * 2 && 
				!pTabWnd->IsLeftRightRounded ())
			{
				rcText.right -= globalUtils.ScaleByDPI(TEXT_MARGIN);
			}
			
			if (clrText == (COLORREF)-1)
			{
				clrText = GetTabTextColor (pTabWnd, iTab, bIsActive);
			}

			if (clrText != (COLORREF)-1)
			{
				pDC->SetTextColor (clrText);
			}

			CString strText;
			pTabWnd->GetTabLabel (iTab, strText);

			if (pTabWnd->IsDrawNameInUpperCase())
			{
				strText.MakeUpper();
			}

			UINT nFormat = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
			if (pTabWnd->IsDrawNoPrefix ())
			{
				nFormat |= DT_NOPREFIX;
			}

			if (pTabWnd->IsOneNoteStyle () || pTabWnd->IsVS2005Style ())
			{
				nFormat |= DT_CENTER;
			}
			else
			{
				nFormat |= DT_LEFT;
			}

			pDC->DrawText (strText, rcText, nFormat);
		}
	}

	m_nCurrDrawedTab = -1;
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTabButton (CDC* pDC, CRect rect, 
											   const CBCGPBaseTabWnd* /*pTabWnd*/,
											   int nID,
											   BOOL bIsHighlighted,
											   BOOL bIsPressed,
											   BOOL /*bIsDisabled*/)
{
	if (bIsHighlighted)
	{
		pDC->FillRect (rect, &globalData.brBarFace);
	}

	CBCGPMenuImages::Draw (pDC, (CBCGPMenuImages::IMAGES_IDS)nID, rect, CBCGPMenuImages::ImageBlack);

	if (bIsHighlighted)
	{
		if (bIsPressed)
		{
			pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarHilite);
		}
		else
		{
			pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarDkShadow);
		}
	}
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTabCloseButton (CDC* pDC, CRect rect, 
											   const CBCGPBaseTabWnd* pTabWnd,
											   BOOL bIsHighlighted,
											   BOOL bIsPressed,
											   BOOL bIsDisabled)
{
	OnDrawTabButton(pDC, rect, pTabWnd, CBCGPMenuImages::IdClose, bIsHighlighted, bIsPressed, bIsDisabled);
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTabPinButton (CDC* pDC, CRect rect, 
											   const CBCGPBaseTabWnd* pTabWnd,
											   BOOL bIsPinned,
											   BOOL bIsHighlighted,
											   BOOL bIsPressed,
											   BOOL bIsDisabled)
{
	OnDrawTabButton(pDC, rect, pTabWnd, bIsPinned ? CBCGPMenuImages::IdPinVert : CBCGPMenuImages::IdPinHorz, bIsHighlighted, bIsPressed, bIsDisabled);
}
//**********************************************************************************
void CBCGPVisualManager::OnEraseTabsButton (CDC* pDC, CRect rect, 
											CBCGPButton* /*pButton*/,
											CBCGPBaseTabWnd* /*pWndTab*/)
{
	ASSERT_VALID (pDC);
	pDC->FillRect (rect, &globalData.brBarFace);
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTabsButtonBorder (CDC* pDC, CRect& rect, 
												 CBCGPButton* pButton, UINT uiState,
												 CBCGPBaseTabWnd* /*pWndTab*/)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->IsPressed () || (uiState & ODS_SELECTED))
	{
		pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarHilite);

		rect.left += 2;
		rect.top += 2;
	}
	else
	{
		pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarDkShadow);
	}

	rect.DeflateRect (2, 2);
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTabResizeBar (CDC* pDC, CBCGPBaseTabWnd* /*pWndTab*/, 
									BOOL bIsVert, CRect rect,
									CBrush* pbrFace, CPen* pPen)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pbrFace);
	ASSERT_VALID (pPen);

	pDC->FillRect (rect, pbrFace);

	CPen* pOldPen = pDC->SelectObject (pPen);
	ASSERT_VALID (pOldPen);

	if (bIsVert)
	{
		pDC->MoveTo (rect.left, rect.top);
		pDC->LineTo (rect.left, rect.bottom);
	}
	else
	{
		pDC->MoveTo (rect.left, rect.top);
		pDC->LineTo (rect.right, rect.top);
	}

	pDC->SelectObject (pOldPen);
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTabsPointer(CDC* pDC, CBCGPBaseTabWnd* pWndTab, const CRect& rectTabs, int nPointerAreaHeight, const CRect& rectActiveTab)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pWndTab);

	const BOOL bIsOnTop = pWndTab->GetLocation() == CBCGPBaseTabWnd::LOCATION_TOP;
	int yLine = bIsOnTop ? rectTabs.bottom - nPointerAreaHeight / 2 : rectTabs.top + nPointerAreaHeight / 2;
	
	CBCGPDrawManager dm(*pDC);

	COLORREF clrLine = GetTabPointerColor(pWndTab);

	if (!rectActiveTab.IsRectEmpty())
	{
		nPointerAreaHeight -= nPointerAreaHeight / 5;

		int x = rectActiveTab.CenterPoint().x;
		int x1 = x - nPointerAreaHeight / 2;
		int x2 = x1 + nPointerAreaHeight - 1;
		int y1 = bIsOnTop ? yLine - nPointerAreaHeight / 2 - 1 : yLine + nPointerAreaHeight / 2 + 1;

		dm.DrawLine(rectTabs.left, yLine, x1, yLine, clrLine);

		if (pWndTab->GetExStyle() & WS_EX_LAYOUTRTL)
		{
			dm.DrawLineA(x2, y1, x, yLine, clrLine);
			dm.DrawLineA(x, yLine, x1, y1, clrLine);
		}
		else
		{
			dm.DrawLineA(x1, yLine, x, y1, clrLine);
			dm.DrawLineA(x, y1, x2, yLine, clrLine);
		}

		dm.DrawLine(x2, yLine, rectTabs.right, yLine, clrLine);
	}
	else
	{
		dm.DrawLine(rectTabs.left, yLine, rectTabs.right, yLine, clrLine);
	}
}
//**********************************************************************************
COLORREF CBCGPVisualManager::GetTabPointerColor(CBCGPBaseTabWnd* /*pWndTab*/)
{
	return globalData.clrBarShadow;
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTabDot(CDC* pDC, CBCGPBaseTabWnd* pWndTab, int nShape, const CRect& recrIn, int iTab, BOOL bIsActive, BOOL bIsHighlighted)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);
	ASSERT_VALID(pWndTab);

	COLORREF clrShadow = pWndTab->IsDialogControl(TRUE) ? globalData.clrBtnShadow : globalData.clrBarShadow;
	COLORREF clrHighlight = pWndTab->IsDialogControl(TRUE) ? GetSysColor(COLOR_HIGHLIGHT) : globalData.clrHilite;
	
	CBCGPDrawManager dm (*pDC);

	CRect rect = recrIn;

	COLORREF clrFill = bIsHighlighted || bIsActive ? clrHighlight : clrShadow;
	COLORREF clrBorder = bIsHighlighted || bIsActive ? clrHighlight : clrShadow;

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
//**********************************************************************************
COLORREF CBCGPVisualManager::OnFillCommandsListBackground (CDC* pDC, CRect rect, BOOL bIsSelected)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	if (bIsSelected)
	{
		pDC->FillRect (rect, &globalData.brHilite);

		const int nFrameSize = 1;

		rect.DeflateRect (1, 1);
		rect.right--;
		rect.bottom--;

		pDC->PatBlt (rect.left, rect.top + nFrameSize, nFrameSize, rect.Height (), PATINVERT);
		pDC->PatBlt (rect.left, rect.top, rect.Width (), nFrameSize, PATINVERT);
		pDC->PatBlt (rect.right, rect.top, nFrameSize, rect.Height (), PATINVERT);
		pDC->PatBlt (rect.left + nFrameSize, rect.bottom, rect.Width (), nFrameSize, PATINVERT);

		return globalData.clrTextHilite;
	}

	pDC->FillRect (rect, &globalData.brBarFace);

	return globalData.clrBarText;
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawMenuArrowOnCustomizeList (CDC* pDC, 
	CRect rectCommand, BOOL bSelected)
{
	CRect rectTriangle = rectCommand;
	rectTriangle.left = rectTriangle.right - CBCGPMenuImages::Size ().cx;

	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdArowRightLarge, rectTriangle,
		bSelected ? CBCGPMenuImages::ImageWhite : CBCGPMenuImages::ImageBlack);

	CRect rectLine = rectCommand;
	rectLine.right = rectTriangle.left - 1;
	rectLine.left = rectLine.right - 2;
	rectLine.DeflateRect (0, 2);

	pDC->Draw3dRect (&rectLine, globalData.clrBtnShadow, globalData.clrBtnHilite);
}
//*********************************************************************************
CBCGPVisualManager* CBCGPVisualManager::CreateVisualManager (
					CRuntimeClass* pVisualManager)
{
	if (pVisualManager == NULL)
	{
		ASSERT (FALSE);
		return NULL;
	}

	CBCGPVisualManager* pVisManagerOld = m_pVisManager;
	
	CObject* pObj = pVisualManager->CreateObject ();
	if (pObj == NULL)
	{
		ASSERT (FALSE);
		return NULL;
	}

	ASSERT_VALID (pObj);
	
	if (pVisManagerOld != NULL)
	{
		ASSERT_VALID (pVisManagerOld);
		delete pVisManagerOld;
	}
	
	m_pVisManager = (CBCGPVisualManager*) pObj;
	m_pVisManager->m_bAutoDestroy = TRUE;

	return m_pVisManager;
}
//*************************************************************************************
void CBCGPVisualManager::DestroyInstance (BOOL bAutoDestroyOnly)
{
	if (m_pVisManager == NULL)
	{
		return;
	}

	ASSERT_VALID (m_pVisManager);

	if (bAutoDestroyOnly && !m_pVisManager->m_bAutoDestroy)
	{
		return;
	}

	delete m_pVisManager;
	m_pVisManager = NULL;
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawTearOffCaption (CDC* pDC, CRect rect, BOOL bIsActive)
{
	const int iBorderSize = 2;

	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBarFace);
	
	rect.DeflateRect (iBorderSize, globalUtils.ScaleByDPI(1));

	pDC->FillSolidRect (rect, 
		bIsActive ? 
		globalData.clrActiveCaption :
		globalData.clrInactiveCaption);
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawMenuResizeBar (CDC* pDC, CRect rect, int nResizeFlags)
{
	ASSERT_VALID (pDC);

	pDC->FillSolidRect (rect, globalData.clrInactiveCaption);
	OnDrawMenuResizeGipper(pDC, rect, nResizeFlags, globalData.clrInactiveCaptionText);
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawMenuResizeGipper(CDC* pDC, CRect rect, int nResizeFlags, COLORREF clrGripper)
{
	ASSERT_VALID (pDC);
	
	int nPadding = globalUtils.ScaleByDPI(2);

	rect.DeflateRect(nPadding, nPadding);

	CBCGPDrawManager dm(*pDC);
	
	if (nResizeFlags == (int) CBCGPPopupMenu::MENU_RESIZE_BOTTOM_RIGHT ||
		nResizeFlags == (int) CBCGPPopupMenu::MENU_RESIZE_TOP_RIGHT)
	{
		rect.left = rect.right - rect.Height();
		
		int dx = globalUtils.ScaleByDPI(rect.Width() / 2);
		int dy = globalUtils.ScaleByDPI(rect.Height() / 2);
		
		for (int x = rect.left; x < rect.right; x += dx)
		{
			for (int y = rect.top; y < rect.bottom; y += dy)
			{
				if (nResizeFlags == (int) CBCGPPopupMenu::MENU_RESIZE_BOTTOM_RIGHT)
				{
					if (x == rect.left && y == rect.top)
					{
						continue;
					}
				}
				else
				{
					if (x == rect.left && y == rect.top + dy)
					{
						continue;
					}
				}
				
				CRect rectBox(x + 1, y + 1, x + dx - 1, y + dy - 1);
				dm.DrawRect(rectBox, clrGripper, (COLORREF)-1);
			}
		}
	}
	else
	{
		int dx = rect.Height();
		
		for (int x = rect.CenterPoint().x - 3 * dx; x <= rect.CenterPoint().x + 3 * dx; x += dx)
		{
			CRect rectBox = rect;
			rectBox.left = x;
			rectBox.right = x + dx;
			
			rectBox.DeflateRect(1, 2, 2, 1);
			dm.DrawRect(rectBox, clrGripper, (COLORREF)-1);
		}
	}
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawMenuScrollButton (CDC* pDC, CRect rect, BOOL bIsScrollDown, 
												 BOOL bIsHighlited, BOOL /*bIsPressed*/,
												 BOOL /*bIsDisabled*/)
{
	ASSERT_VALID (pDC);

	CRect rectFill = rect;
	rectFill.top -= 2;

	pDC->FillRect (rectFill, &globalData.brBarFace);

	CBCGPMenuImages::Draw (pDC, bIsScrollDown ? CBCGPMenuImages::IdArowDown : CBCGPMenuImages::IdArowUp, rect);

	if (bIsHighlited)
	{
		pDC->Draw3dRect (rect,
			globalData.clrBarHilite,
			globalData.clrBarShadow);
	}
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawMenuSystemButton (CDC* pDC, CRect rect, 
												UINT uiSystemCommand, 
												UINT nStyle, BOOL /*bHighlight*/)
{
	ASSERT_VALID (pDC);

	UINT uiState = 0;

	switch (uiSystemCommand)
	{
	case SC_CLOSE:
		uiState |= DFCS_CAPTIONCLOSE;
		break;

	case SC_MINIMIZE:
		uiState |= DFCS_CAPTIONMIN;
		break;

	case SC_RESTORE:
		uiState |= DFCS_CAPTIONRESTORE;
		break;

	default:
		return;
	}

	if (nStyle & TBBS_PRESSED)
	{
		uiState |= DFCS_PUSHED;
	}
	
	if (nStyle & TBBS_DISABLED) // Jan Vasina: Add support for disabled buttons
	{
		uiState |= DFCS_INACTIVE;
	}

	pDC->DrawFrameControl (rect, DFC_CAPTION, uiState);
}
//**************************************************************************************
void CBCGPVisualManager::OnDrawComboDropButton (CDC* pDC, CRect rect,
											    BOOL bDisabled,
												BOOL bIsDropped,
												BOOL bIsHighlighted,
												CBCGPToolbarComboBoxButton* /*pButton*/)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID (this);

	COLORREF clrText = pDC->GetTextColor ();

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawRect (rect, globalData.clrBarFace, globalData.clrBarHilite);

		if (bIsDropped)
		{
			rect.OffsetRect (1, 1);
			dm.DrawRect (rect, (COLORREF)-1, globalData.clrBarShadow);
		}
		else if (bIsHighlighted)
		{
			dm.DrawRect (rect, (COLORREF)-1, globalData.clrBarShadow);
		}
	}
	else
	{
		pDC->FillRect (rect, &globalData.brBarFace);
		pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarHilite);

		if (bIsDropped)
		{
			rect.OffsetRect (1, 1);
			pDC->Draw3dRect (&rect, globalData.clrBarShadow, globalData.clrBarHilite);
		}
		else if (bIsHighlighted || globalData.IsHighContastMode())
		{
			pDC->Draw3dRect (&rect, globalData.clrBarHilite, globalData.clrBarShadow);
		}
	}

	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdArowDown, rect, bDisabled ? CBCGPMenuImages::ImageGray : CBCGPMenuImages::ImageBlack);

	pDC->SetTextColor (clrText);
}
//*************************************************************************************
void CBCGPVisualManager::OnDrawComboBorder (CDC* pDC, CRect rect,
												BOOL /*bDisabled*/,
												BOOL bIsDropped,
												BOOL bIsHighlighted,
												CBCGPToolbarComboBoxButton* /*pButton*/)
{
	ASSERT_VALID (pDC);

	if (bIsHighlighted || bIsDropped || globalData.IsHighContastMode() || CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		if (m_bMenuFlatLook)
		{
			rect.DeflateRect (1, 1);
		}

		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect (rect, (COLORREF)-1, globalData.clrBarDkShadow);
		}
		else
		{
			if (m_bMenuFlatLook)
			{
				pDC->Draw3dRect (&rect, globalData.clrBarDkShadow, globalData.clrBarDkShadow);
			}
			else
			{
				pDC->Draw3dRect (&rect, globalData.clrBarShadow, globalData.clrBarHilite);
			}
		}
	}
}
//*************************************************************************************
void CBCGPVisualManager::OnFillCombo (CDC* pDC, CRect rect,
										BOOL bDisabled,
										BOOL /*bIsDropped*/,
										BOOL /*bIsHighlighted*/,
										CBCGPToolbarComboBoxButton* /*pButton*/)
{
	ASSERT_VALID (pDC);

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);

		dm.DrawRect (rect, 
			bDisabled ? globalData.clrBtnFace : globalData.clrWindow, 
			bDisabled ? globalData.clrBarHilite : (COLORREF)-1);
	}
	else
	{
		pDC->FillRect (rect, bDisabled ? &globalData.brBtnFace : &globalData.brWindow);

		if (bDisabled)
		{
			pDC->Draw3dRect (&rect, globalData.clrBarHilite, globalData.clrBarHilite);
		}
	}
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetComboboxTextColor(CBCGPToolbarComboBoxButton* /*pButton*/, 
										BOOL bDisabled,
										BOOL /*bIsDropped*/,
										BOOL /*bIsHighlighted*/)
{
	return bDisabled ? globalData.clrGrayedText : globalData.clrWindowText;
}
//********************************************************************************
void CBCGPVisualManager::OnDrawStatusBarPaneBorder (CDC* pDC, CBCGPStatusBar* /*pBar*/,
					CRect rectPane, UINT /*uiID*/, UINT nStyle)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID (this);

	if (!(nStyle & SBPS_NOBORDERS))
	{
		// draw the borders
		COLORREF clrHilite;
		COLORREF clrShadow;

		if (nStyle & SBPS_POPOUT)
		{
			// reverse colors
			clrHilite = globalData.clrBarShadow;
			clrShadow = globalData.clrBarHilite;
		}
		else
		{
			// normal colors
			clrHilite = globalData.clrBarHilite;
			clrShadow = globalData.clrBarShadow;
		}

		pDC->Draw3dRect (rectPane, clrShadow, clrHilite);
	}
}
//*********************************************************************************
COLORREF CBCGPVisualManager::OnFillMiniFrameCaption (CDC* pDC, 
								CRect rectCaption, 
								CBCGPMiniFrameWnd* pFrameWnd, BOOL bActive)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pFrameWnd);

	if (DYNAMIC_DOWNCAST (CBCGPBaseToolBar, pFrameWnd->GetControlBar ()) != NULL)
	{
		bActive = TRUE;
	}

	CBrush br (bActive ? globalData.clrActiveCaption : globalData.clrInactiveCaption);
	pDC->FillRect (rectCaption, &br);

    // get the text color
	return bActive ? globalData.clrCaptionText : globalData.clrInactiveCaptionText;
}
//**************************************************************************************
void CBCGPVisualManager::OnDrawMiniFrameBorder (
										CDC* pDC, CBCGPMiniFrameWnd* pFrameWnd,
										CRect rectBorder, CRect rectBorderSize)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pFrameWnd);
	
	BCGP_PREDOCK_STATE preDockState = pFrameWnd->GetPreDockState ();

	if (preDockState == BCGP_PDS_DOCK_REGULAR)
	{
		// draw outer edge;
		pDC->Draw3dRect (rectBorder, RGB (127, 0, 0), globalData.clrBarDkShadow);
		rectBorder.DeflateRect (1, 1);
		pDC->Draw3dRect (rectBorder, globalData.clrBarHilite, RGB (127, 0, 0));
	}
	else if (preDockState == BCGP_PDS_DOCK_TO_TAB)
	{
		// draw outer edge;
		pDC->Draw3dRect (rectBorder, RGB (0, 0, 127), globalData.clrBarDkShadow);
		rectBorder.DeflateRect (1, 1);
		pDC->Draw3dRect (rectBorder, globalData.clrBarHilite, RGB (0, 0, 127));
	}
	else
	{
		// draw outer edge;
		pDC->Draw3dRect (rectBorder, globalData.clrBarFace, globalData.clrBarDkShadow);
		rectBorder.DeflateRect (1, 1);
		pDC->Draw3dRect (rectBorder, globalData.clrBarHilite, globalData.clrBarShadow);
	}	

	// draw the inner egde
	rectBorder.DeflateRect (rectBorderSize.right - 2, rectBorderSize.top - 2);
	pDC->Draw3dRect (rectBorder, globalData.clrBarFace, globalData.clrBarFace);
	rectBorder.InflateRect (1, 1);
	pDC->Draw3dRect (rectBorder, globalData.clrBarFace, globalData.clrBarFace);
}
//**************************************************************************************
void CBCGPVisualManager::OnDrawFloatingToolbarBorder (
												CDC* pDC, CBCGPBaseToolBar* /*pToolBar*/, 
												CRect rectBorder, CRect /*rectBorderSize*/)
{
	ASSERT_VALID (pDC);

	pDC->Draw3dRect (rectBorder, globalData.clrBarFace, globalData.clrBarDkShadow);
	rectBorder.DeflateRect (1, 1);
	pDC->Draw3dRect (rectBorder, globalData.clrBarHilite, globalData.clrBarShadow);
	rectBorder.DeflateRect (1, 1);
	pDC->Draw3dRect (rectBorder, globalData.clrBarFace, globalData.clrBarFace);
}
//**************************************************************************************
COLORREF CBCGPVisualManager::GetToolbarButtonTextColor (CBCGPToolbarButton* pButton,
												  CBCGPVisualManager::BCGBUTTON_STATE state)
{
	ASSERT_VALID (pButton);

	BOOL bDisabled = (CBCGPToolBar::IsCustomizeMode () && !pButton->IsEditable ()) ||
		(!CBCGPToolBar::IsCustomizeMode () && (pButton->m_nStyle & TBBS_DISABLED));

	if (pButton->IsKindOf (RUNTIME_CLASS (CBCGPOutlookButton)))
	{
		if (globalData.IsHighContastMode ())
		{
			return bDisabled ? globalData.clrGrayedText : globalData.clrWindowText;
		}

		return bDisabled ? globalData.clrBtnFace : globalData.clrWindow;
	}

	return	(bDisabled ? globalData.clrGrayedText : 
			(state == ButtonsIsHighlighted) ? 
				CBCGPToolBar::GetHotTextColor () : globalData.clrBarText);
}
//***************************************************************************************
BOOL CBCGPVisualManager::IsToolBarButtonDefaultBackground (CBCGPToolbarButton* pButton,
												CBCGPVisualManager::BCGBUTTON_STATE /*state*/)
{
	ASSERT_VALID (pButton);
	return !(pButton->m_nStyle & (TBBS_CHECKED | TBBS_INDETERMINATE));
}
//***************************************************************************************
void CBCGPVisualManager::OnFillOutlookPageButton (CDC* pDC, 
												  const CRect& rect,
												  BOOL /*bIsHighlighted*/, BOOL /*bIsPressed*/,
												  COLORREF& clrText)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBarFace);
	clrText = globalData.clrBarText;
}
//****************************************************************************************
void CBCGPVisualManager::OnDrawOutlookPageButtonBorder (
	CDC* pDC, CRect& rectBtn, BOOL bIsHighlighted, BOOL bIsPressed)
{
	ASSERT_VALID (pDC);

	if (bIsHighlighted && bIsPressed)
	{
		pDC->Draw3dRect (rectBtn, globalData.clrBarDkShadow, globalData.clrBarFace);
		rectBtn.DeflateRect (1, 1);
		pDC->Draw3dRect (rectBtn, globalData.clrBarShadow, globalData.clrBarHilite);
	}
	else
	{
		if (bIsHighlighted || bIsPressed)
		{
			pDC->Draw3dRect (rectBtn, globalData.clrBarFace, globalData.clrBarDkShadow);
			rectBtn.DeflateRect (1, 1);
		}

		pDC->Draw3dRect (rectBtn, globalData.clrBarHilite, globalData.clrBarShadow);
	}

	rectBtn.DeflateRect (1, 1);
}
//**********************************************************************************
COLORREF CBCGPVisualManager::GetCaptionBarTextColor (CBCGPCaptionBar* pBar)
{
	ASSERT_VALID (pBar);

	if (globalData.IsHighContastMode())
	{
		return globalData.clrBarText;
	}

	return pBar->IsMessageBarMode () ? ::GetSysColor (COLOR_INFOTEXT) : globalData.clrWindow;
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawCaptionBarBorder (CDC* pDC, 
	CBCGPCaptionBar* /*pBar*/, CRect rect, COLORREF clrBarBorder, BOOL bFlatBorder)
{
	ASSERT_VALID (pDC);

	if (clrBarBorder == (COLORREF) -1)
	{
		pDC->FillRect (rect, &globalData.brBarFace);
	}
	else
	{
		CBrush brBorder (clrBarBorder);
		pDC->FillRect (rect, &brBorder);
	}

	if (!bFlatBorder)
	{
		pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarShadow);
	}
}
//**************************************************************************************
void CBCGPVisualManager::OnDrawCaptionBarInfoArea (CDC* pDC, CBCGPCaptionBar* pBar, CRect rect)
{
	ASSERT_VALID (pDC);

	if (pBar->m_clrBarBackground != -1)
	{
		pDC->FillSolidRect(rect, pBar->m_clrBarBackground); 

		COLORREF clrBorder = pBar->m_clrBarBackground != -1 ? CBCGPDrawManager::PixelAlpha (pBar->m_clrBarBackground, 80) : globalData.clrBarShadow;
		pDC->Draw3dRect (rect, clrBorder, clrBorder);
	}
	else
	{
		::FillRect (pDC->GetSafeHdc (), rect, ::GetSysColorBrush (COLOR_INFOBK));

		pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarHilite);
		rect.DeflateRect (1, 1);
		pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarShadow);
	}
}
//**************************************************************************************
COLORREF CBCGPVisualManager::OnFillCaptionBarButton (CDC* pDC, CBCGPCaptionBar* pBar,
											CRect rect, BOOL bIsPressed, BOOL bIsHighlighted, 
											BOOL bIsDisabled, BOOL bHasDropDownArrow,
											BOOL bIsSysButton)
{
	UNREFERENCED_PARAMETER(bIsPressed);
	UNREFERENCED_PARAMETER(bIsHighlighted);
	UNREFERENCED_PARAMETER(bIsDisabled);
	UNREFERENCED_PARAMETER(bHasDropDownArrow);
	UNREFERENCED_PARAMETER(bIsSysButton);

	ASSERT_VALID (pBar);

	if (!pBar->IsMessageBarMode ())
	{
		return (COLORREF)-1;
	}

	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBarFace);
	return bIsDisabled ? globalData.clrGrayedText : globalData.clrBarText;
}
//**************************************************************************************
void CBCGPVisualManager::OnDrawCaptionBarButtonBorder (CDC* pDC, CBCGPCaptionBar* pBar,
											CRect rect, BOOL bIsPressed, BOOL bIsHighlighted, 
											BOOL bIsDisabled, BOOL bHasDropDownArrow,
											BOOL bIsSysButton)
{
	UNREFERENCED_PARAMETER(bIsDisabled);
	UNREFERENCED_PARAMETER(bHasDropDownArrow);
	UNREFERENCED_PARAMETER(bIsSysButton);

	ASSERT_VALID (pDC);
	ASSERT_VALID (pBar);

	if (bIsPressed)
	{
		pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarHilite);
	}
	else if (bIsHighlighted || pBar->IsMessageBarMode ())
	{
		pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarDkShadow);
	}
}
//**************************************************************************************
void CBCGPVisualManager::OnDrawCaptionBarCloseButton(CDC* pDC, CBCGPCaptionBar* /*pBar*/,
			CRect rectClose, BOOL /*bIsCloseBtnPressed*/, BOOL /*bIsCloseBtnHighlighted*/, COLORREF clrText)
{
	ASSERT_VALID (pDC);

	CBCGPMenuImages::DrawByColor(pDC, CBCGPMenuImages::IdClose, rectClose, clrText, FALSE);
}
//**************************************************************************************
void CBCGPVisualManager::OnDrawStatusBarProgress (CDC* pDC, CBCGPStatusBar* /*pStatusBar*/,
			CRect rectProgress, int nProgressTotal, int nProgressCurr,
			COLORREF clrBar, COLORREF clrProgressBarDest, COLORREF clrProgressText,
			BOOL bProgressText)
{
	ASSERT_VALID (pDC);

	if (nProgressTotal == 0)
	{
		return;
	}

	CRect rectComplete = rectProgress;
	rectComplete.right = rectComplete.left + 
		nProgressCurr * rectComplete.Width () / nProgressTotal;

	if (clrProgressBarDest == (COLORREF)-1)
	{
		// one-color bar
		CBrush br (clrBar);
		pDC->FillRect (rectComplete, &br);
	}
	else
	{
		// gradient bar:
		CBCGPDrawManager dm (*pDC);
		dm.FillGradient (rectComplete, clrBar, clrProgressBarDest, FALSE);
	}

	if (bProgressText)
	{
		CString strText;
		strText.Format (_T("%d%%"), nProgressCurr * 100 / nProgressTotal);

		COLORREF clrText = pDC->SetTextColor (globalData.clrBarText);

		pDC->DrawText (strText, rectProgress, DT_CENTER | DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX);

		CRgn rgn;
		rgn.CreateRectRgnIndirect (rectComplete);
		pDC->SelectClipRgn (&rgn);

		pDC->SetTextColor (clrProgressText == (COLORREF)-1 ?
			globalData.clrTextHilite : clrProgressText);
		
		pDC->DrawText (strText, rectProgress, DT_CENTER | DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX);
		pDC->SelectClipRgn (NULL);
		pDC->SetTextColor (clrText);
	}
}
//****************************************************************************************
void CBCGPVisualManager::OnFillHeaderCtrlBackground (CBCGPHeaderCtrl* pCtrl,
													 CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	BOOL bIsVMStyle = (pCtrl == NULL || pCtrl->m_bVisualManagerStyle);
	pDC->FillRect (rect, pCtrl->IsDialogControl () || !bIsVMStyle ? &globalData.brBtnFace : &globalData.brBarFace);
}
//****************************************************************************************
void CBCGPVisualManager::OnDrawHeaderCtrlBorder (CBCGPHeaderCtrl* pCtrl, CDC* pDC,
		CRect& rect, BOOL bIsPressed, BOOL /*bIsHighlighted*/)
{
	ASSERT_VALID (pDC);

	BOOL bHasFilterBar = pCtrl->GetSafeHwnd() != NULL && ((pCtrl->GetStyle() & HDS_FILTERBAR) == HDS_FILTERBAR);
	BOOL bHasCheckBoxes = pCtrl->GetSafeHwnd() != NULL && ((pCtrl->GetStyle() & HDS_CHECKBOXES) == HDS_CHECKBOXES);

	if (bHasFilterBar || bHasCheckBoxes)
	{
		pDC->FillRect(rect, &globalData.brBtnFace);
	}

	BOOL bIsVMStyle = (pCtrl == NULL || pCtrl->m_bVisualManagerStyle);

	if (bIsPressed)
	{
		if (pCtrl->IsDialogControl () || !bIsVMStyle)
		{
			pDC->Draw3dRect (rect, globalData.clrBtnShadow, globalData.clrBtnShadow);
		}
		else
		{
			pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarShadow);
		}

		rect.left++;
		rect.top++;
	}
	else
	{
		if (pCtrl->IsDialogControl () || !bIsVMStyle)
		{
			pDC->Draw3dRect (rect, globalData.clrBtnHilite, globalData.clrBtnShadow);
		}
		else
		{
			pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarShadow);
		}
	}
}
//*****************************************************************************************
void CBCGPVisualManager::OnDrawHeaderCtrlSortArrow (CBCGPHeaderCtrl* pCtrl, 
												   CDC* pDC,
												   CRect& rectArrow, BOOL bIsUp)
{
	DoDrawHeaderSortArrow (pDC, rectArrow, bIsUp, pCtrl != NULL && pCtrl->IsDialogControl ());
}
//*****************************************************************************************
void CBCGPVisualManager::OnDrawStatusBarSizeBox (CDC* pDC, CBCGPStatusBar* /*pStatBar*/,
			CRect rectSizeBox)
{
	ASSERT_VALID (pDC);

	CFont* pOldFont = pDC->SelectObject (&globalData.fontMarlett);
	ASSERT (pOldFont != NULL);

	// Char of the sizing box in "Marlett" font
	const CString strSizeBox (globalData.m_bIsRTL ? _T("x") : _T("o"));

	UINT nTextAlign = pDC->SetTextAlign (TA_RIGHT | TA_BOTTOM);
	COLORREF clrText = pDC->SetTextColor (globalData.clrBarShadow);

	pDC->ExtTextOut (rectSizeBox.right, rectSizeBox.bottom,
		ETO_CLIPPED, &rectSizeBox, strSizeBox, NULL);

	pDC->SelectObject (pOldFont);
	pDC->SetTextColor (clrText);
	pDC->SetTextAlign (nTextAlign);
}
//*************************************************************************************
void CBCGPVisualManager::OnDrawEditBorder (CDC* pDC, CRect rect,
												BOOL /*bDisabled*/,
												BOOL bIsHighlighted,
												CBCGPToolbarEditBoxButton* /*pButton*/)
{
	ASSERT_VALID (pDC);

	if (bIsHighlighted)
	{
		pDC->DrawEdge (rect, EDGE_SUNKEN, BF_RECT);
	}
}
//*************************************************************************************
HBRUSH CBCGPVisualManager::GetToolbarEditColors(CBCGPToolbarButton* /*pButton*/, COLORREF& clrText, COLORREF& clrBk)
{
	clrText = globalData.clrWindowText;
	clrBk = globalData.clrWindow;

	return (HBRUSH) globalData.brWindow.GetSafeHandle ();
}

#ifndef BCGP_EXCLUDE_TASK_PANE

void CBCGPVisualManager::OnFillTasksPaneBackground(CDC* pDC, CRect rectWorkArea)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rectWorkArea, &globalData.brWindow);
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTasksGroupCaption(CDC* pDC, CBCGPTasksGroup* pGroup, 
						BOOL bIsHighlighted, BOOL bIsSelected, BOOL bCanCollapse)
{
	ASSERT_VALID(pDC);
	ASSERT(pGroup != NULL);
	ASSERT_VALID (pGroup->m_pPage);

#ifndef BCGP_EXCLUDE_TOOLBOX
	BOOL bIsToolBox = pGroup->m_pPage->m_pTaskPane != NULL &&
		(pGroup->m_pPage->m_pTaskPane->IsKindOf (RUNTIME_CLASS (CBCGPToolBoxEx)));
#else
	BOOL bIsToolBox = FALSE;
#endif

	CRect rectGroup = pGroup->m_rect;

	if (bIsToolBox)
	{
		CRect rectFill = rectGroup;
		rectFill.DeflateRect (1, 0, 1, 1);

		CBrush brFill (globalData.clrBarShadow);
		pDC->FillRect (rectFill, &brFill);

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

			pDC->FillRect (rectButton, &globalData.brBarFace);

			OnDrawExpandingBox (pDC, rectButton, !pGroup->m_bIsCollapsed, 
				globalData.clrBarText);

			rectGroup.left = rectButton.right + nBoxOffset;
			bCanCollapse = FALSE;
		}
	}
	else
	{
		// ---------------------------------
		// Draw caption background (Windows)
		// ---------------------------------
		COLORREF clrBckOld = pDC->GetBkColor ();
		pDC->FillSolidRect(rectGroup, 
			(pGroup->m_bIsSpecial ? globalData.clrHilite : globalData.clrBarFace)); 
		pDC->SetBkColor (clrBckOld);
	}

	// ---------------------------
	// Draw an icon if it presents
	// ---------------------------
	BOOL bShowIcon = (pGroup->m_hIcon != NULL 
		&& pGroup->m_sizeIcon.cx < rectGroup.Width () - rectGroup.Height());
	if (bShowIcon)
	{
		OnDrawTasksGroupIcon(pDC, pGroup, 5, bIsHighlighted, bIsSelected, bCanCollapse);
	}

	// -----------------------
	// Draw group caption text
	// -----------------------
	CFont* pFontOld = pDC->SelectObject (&globalData.fontBold);
	COLORREF clrTextOld = pDC->GetTextColor();

	if (bIsToolBox)
	{
		pDC->SetTextColor (globalData.clrWindow);

	}
	else
	{
		if (bCanCollapse && bIsHighlighted)
		{
			clrTextOld = pDC->SetTextColor (pGroup->m_clrTextHot == (COLORREF)-1 ?
				(pGroup->m_bIsSpecial ? globalData.clrWindow : globalData.clrWindowText) :
				pGroup->m_clrTextHot);
		}
		else
		{
			clrTextOld = pDC->SetTextColor (pGroup->m_clrText == (COLORREF)-1 ?
				(pGroup->m_bIsSpecial ? globalData.clrWindow : globalData.clrWindowText) :
				pGroup->m_clrText);
		}
	}

	int nBkModeOld = pDC->SetBkMode(TRANSPARENT);
	
	int nTaskPaneHOffset = pGroup->m_pPage->m_pTaskPane->GetGroupCaptionHorzOffset();
	int nTaskPaneVOffset = pGroup->m_pPage->m_pTaskPane->GetGroupCaptionVertOffset();
	int nCaptionHOffset = (nTaskPaneHOffset != -1 ? nTaskPaneHOffset : m_nGroupCaptionHorzOffset);

	CRect rectText = rectGroup;
	rectText.left += (bShowIcon ? pGroup->m_sizeIcon.cx	+ 5: nCaptionHOffset);
	rectText.top += (nTaskPaneVOffset != -1 ? nTaskPaneVOffset : m_nGroupCaptionVertOffset);
	rectText.right = max(rectText.left, 
						rectText.right - (bCanCollapse ? rectGroup.Height() : nCaptionHOffset));

	pDC->DrawText (pGroup->m_strName, rectText, DT_SINGLELINE | DT_VCENTER);

	pDC->SetBkMode(nBkModeOld);
	pDC->SelectObject (pFontOld);
	pDC->SetTextColor (clrTextOld);

	// -------------------------
	// Draw group caption button
	// -------------------------
	if (bCanCollapse && !pGroup->m_strName.IsEmpty())
	{
		CSize sizeButton = CBCGPMenuImages::Size();
		CRect rectButton = rectGroup;
		rectButton.left = max(rectButton.left, 
			rectButton.right - (rectButton.Height()+1)/2 - (sizeButton.cx+1)/2);
		rectButton.top = max(rectButton.top, 
			rectButton.bottom - (rectButton.Height()+1)/2 - (sizeButton.cy+1)/2);
		rectButton.right = rectButton.left + sizeButton.cx;
		rectButton.bottom = rectButton.top + sizeButton.cy;

		if (rectButton.right <= rectGroup.right && rectButton.bottom <= rectGroup.bottom)
		{
			if (bIsHighlighted)
			{
				// Draw button frame
				CBrush* pBrushOld = (CBrush*) pDC->SelectObject (&globalData.brBarFace);
				COLORREF clrBckOld = pDC->GetBkColor ();

				pDC->Draw3dRect(&rectButton, globalData.clrWindow, globalData.clrBarShadow);

				pDC->SetBkColor (clrBckOld);
				pDC->SelectObject (pBrushOld);
			}

			if (!pGroup->m_bIsCollapsed)
			{
				CBCGPMenuImages::Draw(pDC, CBCGPMenuImages::IdArowUp, rectButton.TopLeft());
			}
			else
			{
				CBCGPMenuImages::Draw(pDC, CBCGPMenuImages::IdArowDown, rectButton.TopLeft());
			}
		}
	}
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTasksGroupIcon(CDC* pDC, CBCGPTasksGroup* pGroup,
		int nIconHOffset, BOOL /*bIsHighlighted*/, BOOL /*bIsSelected*/, BOOL /*bCanCollapse*/)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pGroup);

	if (pGroup->m_hIcon == NULL)
	{
		return;
	}

	int nTaskPaneVOffset = pGroup->m_pPage->m_pTaskPane->GetGroupCaptionVertOffset();

	CRect rectImage = pGroup->m_rect;
	rectImage.top += (nTaskPaneVOffset != -1 ? nTaskPaneVOffset : m_nGroupCaptionVertOffset);
	rectImage.left += nIconHOffset;
	rectImage.right = rectImage.left + pGroup->m_sizeIcon.cx;

	int x = max (0, (rectImage.Width () - pGroup->m_sizeIcon.cx) / 2);
	int y = max (0, (rectImage.Height () - pGroup->m_sizeIcon.cy) / 2);

	::DrawIconEx (pDC->GetSafeHdc (),
		rectImage.left + x, rectImage.bottom - y - pGroup->m_sizeIcon.cy,
		pGroup->m_hIcon, pGroup->m_sizeIcon.cx, pGroup->m_sizeIcon.cy,
		0, NULL, DI_NORMAL);
}
//**********************************************************************************
void CBCGPVisualManager::OnFillTasksGroupInterior(CDC* /*pDC*/, CRect /*rect*/, BOOL /*bSpecial*/)
{
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTasksGroupAreaBorder(CDC* pDC, CRect rect, BOOL bSpecial, 
												   BOOL bNoTitle)
{
	ASSERT_VALID(pDC);

	// Draw caption background:
	CPen* pPenOld = (CPen*) pDC->SelectObject (bSpecial ? &globalData.penHilite : &globalData.penBarFace);

	pDC->MoveTo (rect.left, rect.top);
	pDC->LineTo (rect.left, rect.bottom-1);
	pDC->LineTo (rect.right-1, rect.bottom-1);
	pDC->LineTo (rect.right-1, rect.top);
	if (bNoTitle)
	{
		pDC->LineTo (rect.left, rect.top);
	}
	else
	{
		pDC->LineTo (rect.right-1, rect.top-1);
	}
	pDC->SelectObject (pPenOld);
	
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTask(CDC* pDC, CBCGPTask* pTask, CImageList* pIcons, 
								   BOOL bIsHighlighted, BOOL /*bIsSelected*/)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pIcons);
	ASSERT(pTask != NULL);

	CRect rectText = pTask->m_rect;

	if (pTask->m_bIsSeparator)
	{
		CPen* pPenOld = (CPen*) pDC->SelectObject (&globalData.penBarFace);

		pDC->MoveTo (rectText.left, rectText.CenterPoint ().y);
		pDC->LineTo (rectText.right, rectText.CenterPoint ().y);

		pDC->SelectObject (pPenOld);
		return;
	}

	CBCGPTasksPane* pTaskPane = pTask->m_pGroup->m_pPage->m_pTaskPane;
	ASSERT_VALID (pTaskPane);

	// ---------
	// Draw icon
	// ---------
	CSize sizeIcon(0, 0);
	::ImageList_GetIconSize (pIcons->m_hImageList, (int*) &sizeIcon.cx, (int*) &sizeIcon.cy);
	if (pTask->m_nIcon >= 0 && sizeIcon.cx > 0)
	{
		CPoint ptIcon = rectText.TopLeft();
		if (!pTaskPane->IsWrapTasksEnabled())
		{
			ptIcon.y += max(0, (rectText.Height() - sizeIcon.cy) / 2);
		}

		pIcons->Draw (pDC, pTask->m_nIcon, ptIcon, ILD_TRANSPARENT);
	}
	int nTaskPaneOffset = pTask->m_pGroup->m_pPage->m_pTaskPane->GetTasksIconHorzOffset();
	rectText.left += sizeIcon.cx + (nTaskPaneOffset != -1 ? nTaskPaneOffset : m_nTasksIconHorzOffset);

	// ---------
	// Draw text
	// ---------
	BOOL bIsLabel = (pTask->m_uiCommandID == 0);

	CFont* pFontOld = NULL;
	COLORREF clrTextOld = pDC->GetTextColor ();
	if (bIsLabel)
	{
		pFontOld = pDC->SelectObject (
			pTask->m_bIsBold ? &globalData.fontBold : &globalData.fontRegular);
		pDC->SetTextColor (pTask->m_clrText == (COLORREF)-1 ?
			globalData.clrWindowText : pTask->m_clrText);
	}
	else if (!pTask->m_bEnabled)
	{
		pDC->SetTextColor (globalData.clrGrayedText);
		pFontOld = pDC->SelectObject (&globalData.fontRegular);
	}
	else if (bIsHighlighted)
	{
		pFontOld = pDC->SelectObject (&globalData.fontUnderline);
		pDC->SetTextColor (pTask->m_clrTextHot == (COLORREF)-1 ?
			globalData.clrWindowText : pTask->m_clrTextHot);
	}
	else
	{
		pFontOld = pDC->SelectObject (&globalData.fontRegular);
		pDC->SetTextColor (pTask->m_clrText == (COLORREF)-1 ?
			globalData.clrWindowText : pTask->m_clrText);
	}

	int nBkModeOld = pDC->SetBkMode(TRANSPARENT);

	BOOL bMultiline = bIsLabel ? 
		pTaskPane->IsWrapLabelsEnabled () : pTaskPane->IsWrapTasksEnabled ();

	if (bMultiline)
	{
		pDC->DrawText (pTask->m_strName, rectText, DT_WORDBREAK);
	}
	else
	{
		CString strText = pTask->m_strName;
		strText.Remove (_T('\n'));
		strText.Remove (_T('\r'));
		pDC->DrawText (strText, rectText, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
	}
	
	pDC->SetBkMode(nBkModeOld);
	pDC->SelectObject (pFontOld);
	pDC->SetTextColor (clrTextOld);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawScrollButtons(CDC* pDC, const CRect& rect, const int nBorderSize,
									int iImage, BOOL bHilited)
{
	ASSERT_VALID (pDC);

	CRect rectImage (CPoint (0, 0), CBCGPMenuImages::Size ());

	CRect rectFill = rect;
	rectFill.top -= nBorderSize;

	pDC->FillRect (rectFill, &globalData.brBarFace);

	if (bHilited)
	{
		CBCGPDrawManager dm (*pDC);
		dm.HighlightRect (rect);

		pDC->Draw3dRect (rect,
			globalData.clrBarHilite,
			globalData.clrBarDkShadow);
	}

	CBCGPMenuImages::Draw (pDC, (CBCGPMenuImages::IMAGES_IDS) iImage, rect);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawToolBoxFrame (CDC* pDC, const CRect& rect)
{
	ASSERT_VALID (pDC);
	pDC->Draw3dRect (rect, globalData.clrBarFace, globalData.clrBarFace);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawFocusRect (CDC* pDC, const CRect& rect)
{
	ASSERT_VALID (pDC);
	pDC->DrawFocusRect(rect);
}

#endif // BCGP_EXCLUDE_TASK_PANE

void CBCGPVisualManager::OnDrawSlider (CDC* pDC, CBCGPSlider* pSlider, CRect rect, BOOL bAutoHideMode)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pSlider);

	CRect rectScreen = globalData.m_rectVirtual;
	pSlider->ScreenToClient (&rectScreen);

	CRect rectFill = rect;
	rectFill.left = min (rectFill.left, rectScreen.left);

	OnFillBarBackground (pDC, pSlider, rectFill, rect);

	if (bAutoHideMode)
	{
		// draw outer edge;

		DWORD dwAlgn = pSlider->GetCurrentAlignment ();
		CRect rectBorder = rect;

		COLORREF clrBorder = globalData.clrBarDkShadow;
		
		if (dwAlgn & CBRS_ALIGN_LEFT)
		{
			rectBorder.left = rectBorder.right;
		}
		else if (dwAlgn & CBRS_ALIGN_RIGHT)
		{
			rectBorder.right = rectBorder.left;
			clrBorder = globalData.clrBarHilite;
		}
		else if (dwAlgn & CBRS_ALIGN_TOP)
		{
			rectBorder.top = rectBorder.bottom;
		}
		else if (dwAlgn & CBRS_ALIGN_BOTTOM)
		{
			rectBorder.bottom = rectBorder.top;
			clrBorder = globalData.clrBarHilite;
		}
		else
		{
			ASSERT(FALSE);
			return;
		}

		pDC->Draw3dRect (rectBorder, clrBorder, clrBorder);
	}
}
//********************************************************************************
void CBCGPVisualManager::OnDrawSplitterBorder (CDC* pDC, CBCGPSplitterWnd* pSplitterWnd, CRect rect)
{
	ASSERT_VALID(pDC);

	int nBorderSize = 2;
	
	if (pSplitterWnd != NULL)
	{
		ASSERT_VALID(pSplitterWnd);
		nBorderSize = pSplitterWnd->GetBorderSize();
	}

	if (nBorderSize > 0)
	{
		pDC->Draw3dRect(rect, globalData.clrBarShadow, globalData.clrBarHilite);

		if (nBorderSize > 1)
		{
			rect.DeflateRect(1, 1);
			pDC->Draw3dRect (rect, globalData.clrBarFace, globalData.clrBarFace);
		}
	}
}
//********************************************************************************
void CBCGPVisualManager::OnDrawSplitterBox (CDC* pDC, CBCGPSplitterWnd* /*pSplitterWnd*/, CRect& rect)
{
	ASSERT_VALID(pDC);
	pDC->Draw3dRect(rect, globalData.clrBarFace, globalData.clrBarShadow);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawSplitterBar(CDC* pDC, CBCGPSplitterWnd* pSplitterWnd, CRect& rect)
{
	ASSERT_VALID(pDC);

	OnFillSplitterBackground(pDC, pSplitterWnd, rect);
	rect.InflateRect(1, 0);
	pDC->Draw3dRect(rect, globalData.clrBarHilite, globalData.clrBarShadow);
}
//********************************************************************************
void CBCGPVisualManager::OnFillSplitterBackground (CDC* pDC, CBCGPSplitterWnd* /*pSplitterWnd*/, CRect rect)
{
	ASSERT_VALID(pDC);
	pDC->FillSolidRect (rect, globalData.clrBarFace);
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawCheckBox (CDC *pDC, CRect rect, 
										 BOOL bHighlighted, 
										 BOOL bChecked,
										 BOOL bEnabled)
{
	OnDrawCheckBoxEx (pDC, rect, bChecked ? 1 : 0, bHighlighted, FALSE, bEnabled);
}
//**********************************************************************************
CSize CBCGPVisualManager::GetCheckRadioDefaultSize ()
{
	if (m_pfGetThemePartSize != NULL && m_hThemeButton != NULL)
	{
		CSize size(0, 0);

		(*m_pfGetThemePartSize)(m_hThemeButton, NULL, BP_CHECKBOX, CBS_UNCHECKEDNORMAL, NULL, TS_TRUE, &size);

		if (size != CSize(0, 0))
		{
			return size;
		}
	}

	return CSize (::GetSystemMetrics (SM_CXMENUCHECK) + 1, ::GetSystemMetrics (SM_CYMENUCHECK) + 1);
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawCheckBoxEx (CDC *pDC, CRect rect, 
											int nState,
											BOOL bHighlighted, 
											BOOL /*bPressed*/,
											BOOL bEnabled)
{
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

	if (CBCGPToolBarImages::m_bIsDrawOnGlass || m_bScaleCheckRadio)
	{
		CBCGPDrawManager dm (*pDC);

		rect.DeflateRect (1, 1);

		dm.DrawRect (rect, 
			bEnabled ? globalData.clrWindow : globalData.clrBarFace, 
			globalData.clrBarDkShadow);

		if (nState == 1)
		{
			if (m_bScaleCheckRadio)
			{
				rect.DeflateRect(rect.Width() / 4, rect.Height() / 4);
				dm.DrawCheckMark(rect, globalData.clrBarDkShadow, (int)(0.5 + m_dblScaleCheckRadio));
			}
			else
			{
				CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdCheck, rect, CBCGPMenuImages::ImageBlack);
			}
		}
		else if (nState == 2)
		{
			rect.DeflateRect (1, 1);
			dm.DrawRect (rect, globalData.clrBarShadow,  (COLORREF)-1);
		}

		return;
	}

	if (bHighlighted)
	{
		pDC->DrawFocusRect (rect);
	}

	rect.DeflateRect (1, 1);
	pDC->FillSolidRect (&rect, bEnabled ? globalData.clrWindow :
		globalData.clrBarFace);

	pDC->Draw3dRect (&rect, 
					globalData.clrBarDkShadow,
					globalData.clrBarHilite);
	
	rect.DeflateRect (1, 1);
	pDC->Draw3dRect (&rect,
					globalData.clrBarShadow,
					globalData.clrBarLight);

	if (nState == 1)
	{
		CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdCheck, rect, CBCGPMenuImages::ImageBlack);
	}
	else if (nState == 2)
	{
		rect.DeflateRect (1, 1);
		
		WORD HatchBits [8] = { 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55 };

		CBitmap bmp;
		bmp.CreateBitmap (8, 8, 1, 1, HatchBits);

		CBrush br;
		br.CreatePatternBrush (&bmp);

		pDC->FillRect (rect, &br);
	}
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawRadioButton (CDC *pDC, CRect rect, 
											BOOL bOn,
											BOOL bHighlighted, 
											BOOL /*bPressed*/,
											BOOL bEnabled)
{
	CBCGPDrawManager dm (*pDC);

	rect.DeflateRect (1, 1);

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

	dm.DrawEllipse (rect,
		!globalData.IsHighContastMode() && bEnabled ? globalData.clrBarHilite : globalData.clrBarFace,
		(bHighlighted || globalData.IsHighContastMode()) && bEnabled ? globalData.clrBarDkShadow : globalData.clrBarShadow);
	
	if (bOn)
	{
		rect.DeflateRect (rect.Width () / 3, rect.Width () / 3);

		dm.DrawEllipse (rect,
			(bHighlighted || globalData.IsHighContastMode()) && bEnabled ? globalData.clrBarDkShadow : globalData.clrBarShadow,
			(COLORREF)-1);
	}
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawSwitch(CDC *pDC, CRect rect, BOOL bOn, BOOL bHighlighted, BOOL /*bPressed*/, BOOL bEnabled)
{
	CBCGPSwitchColors colors(TRUE, FALSE, TRUE);

	if (m_pImagesSwitchOn == NULL)
	{
		m_pImagesSwitchOn = new CBCGPToolBarImages;
		
		m_pImagesSwitchOn->Load(IDB_BCGBARRES_SWITCH_ON);
		m_pImagesSwitchOn->AdaptColors(RGB(0, 0, 0), colors.m_brFillOn.GetColor(), FALSE);

		globalUtils.ScaleByDPI(*m_pImagesSwitchOn);
	}

	if (m_pImagesSwitchOff == NULL)
	{
		m_pImagesSwitchOff = new CBCGPToolBarImages;
		
		m_pImagesSwitchOff->Load(IDB_BCGBARRES_SWITCH_OFF);

		if (!globalData.IsHighContastMode())
		{
			m_pImagesSwitchOff->AdaptColors(RGB(255, 255, 255), colors.m_brFillOff.GetColor(), FALSE);
			m_pImagesSwitchOff->AdaptColors(RGB(0, 0, 0), colors.m_brOutline.GetColor(), FALSE);
		}

		globalUtils.ScaleByDPI(*m_pImagesSwitchOff);
	}

	CBCGPToolBarImages* pImage = bOn ? m_pImagesSwitchOn : m_pImagesSwitchOff;

	BYTE bAlpha = 255;
	if (!bEnabled)
	{
		bAlpha = 150;
	}
	else if (bHighlighted)
	{
		bAlpha = 220;
	}

	pImage->DrawEx(pDC, rect, 0, CBCGPToolBarImages::ImageAlignHorzStretch, CBCGPToolBarImages::ImageAlignVertStretch, CRect(0, 0, 0, 0), bAlpha);
}
//**********************************************************************************
CSize CBCGPVisualManager::GetExpandButtonDefaultSize()
{
	if (m_pfGetThemePartSize != NULL && m_hThemeTaskDialog != NULL && !globalData.IsHighContastMode())
	{
		CSize size(0, 0);
		
		(*m_pfGetThemePartSize)(m_hThemeTaskDialog, NULL, TDLG_EXPANDOBUTTON, TDLGEBS_EXPANDEDNORMAL, NULL, TS_TRUE, &size);
		
		if (size != CSize(0, 0))
		{
			return size;
		}
	}
	
	CSize size = CBCGPMenuImages::Size();
	size.cx *= 2;
	size.cy *= 2;

	return size;
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawGroupExpandCollapse(CDC* pDC, CRect rect, 
									   BOOL bIsExpanded,
									   BOOL bHighlighted, 
									   BOOL bPressed,
									   BOOL bEnabled)
{
	if (m_hThemeTaskDialog == NULL || globalData.IsHighContastMode())
	{
		CBCGPDrawManager dm (*pDC);

		COLORREF clrFill = bEnabled ? (bHighlighted ? globalData.clrBarLight : globalData.clrBarHilite) : globalData.clrBarFace;
		COLORREF clrBorder = bHighlighted && bEnabled ? globalData.clrBarDkShadow : globalData.clrBarShadow;

		if (globalData.IsHighContastMode())
		{
			clrFill = (bEnabled && bHighlighted) ? globalData.clrHilite : globalData.clrWindow;
			clrBorder = bEnabled ? globalData.clrBarDkShadow : globalData.clrBarShadow;
		}

		dm.DrawEllipse (rect, clrFill, clrBorder);

		CBCGPMenuImages::DrawByColor(pDC, bIsExpanded ? CBCGPMenuImages::IdArowUpLarge : CBCGPMenuImages::IdArowDownLarge, rect, clrFill);

		return;
	}
	
	int nState = 0;
	
	if (bIsExpanded)
	{
		nState = TDLGEBS_EXPANDEDNORMAL;
		
		if (bPressed)
		{
			nState = TDLGEBS_EXPANDEDPRESSED;
		}
		else if (bHighlighted)
		{
			nState = TDLGEBS_EXPANDEDHOVER;
		}
	}
	else
	{
		nState = TDLGEBS_NORMAL;

		if (bPressed)
		{
			nState = TDLGEBS_PRESSED;
		}
		else if (bHighlighted)
		{
			nState = TDLGEBS_HOVER;
		}
	}

	(*m_pfDrawThemeBackground) (m_hThemeTaskDialog, pDC->GetSafeHdc(), TDLG_EXPANDOBUTTON, nState, &rect, 0);
}
//*************************************{********************************************
void CBCGPVisualManager::OnDrawSpinButtons (CDC* pDC, CRect rectSpin, 
	int nState, BOOL bOrientation, CBCGPSpinButtonCtrl* /*pSpinCtrl*/)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (this);

	rectSpin.DeflateRect (1, 1);

	CRect rect[2];

	rect[0] = rect[1] = rectSpin;

	if (!bOrientation)
	{
		rect[0].DeflateRect(0, 0, 0, rect[0].Height() / 2);
		rect[1].top = rect[0].bottom + 1;
	}
	else
	{
		rect[1].DeflateRect(0, 0, rect[0].Width() / 2, 0);
		rect[0].left = rect[1].right;
	}

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawRect (rectSpin, globalData.clrBarFace, globalData.clrBarHilite);
	}
	else
	{
		pDC->FillRect (rectSpin, &globalData.brBarFace);
		pDC->Draw3dRect (rectSpin, globalData.clrBarHilite, globalData.clrBarHilite);
	}

	CBCGPMenuImages::IMAGES_IDS id[2][2] = {{CBCGPMenuImages::IdArowUp, CBCGPMenuImages::IdArowDown}, {CBCGPMenuImages::IdArowRight, CBCGPMenuImages::IdArowLeft}};

	int idxPressed = (nState & (SPIN_PRESSEDUP | SPIN_PRESSEDDOWN)) - 1;
	BOOL bDisabled = nState & SPIN_DISABLED;

	for (int i = 0; i < 2; i ++)
	{
		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);

			if (idxPressed == i)
			{
				dm.DrawRect (rect[i], (COLORREF)-1, globalData.clrBarShadow);
			}
			else
			{
				dm.DrawRect (rect[i], (COLORREF)-1, globalData.clrBarHilite);
			}
		}
		else
		{
			if (idxPressed == i)
			{
				pDC->Draw3dRect (&rect[i], globalData.clrBarShadow, globalData.clrBarHilite);
			}
			else
			{
				pDC->Draw3dRect (&rect[i], globalData.clrBarHilite, globalData.clrBarShadow);
			}
		}

		CBCGPMenuImages::Draw (pDC, id[bOrientation ? 1 : 0][i], rect[i],
			bDisabled ? CBCGPMenuImages::ImageGray : CBCGPMenuImages::ImageBlack);
	}
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawDateTimeDropButton (CDC* pDC, CRect rect, 
	BOOL bDisabled, BOOL bPressed, BOOL bHighlighted, CBCGPDateTimeCtrl* pCtrl)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (this);

	if (m_hThemeButton != NULL)
	{
		int nState = PBS_NORMAL;
		
		if (pCtrl->GetSafeHwnd() != NULL && !pCtrl->IsWindowEnabled())
		{
			nState = PBS_DISABLED;
		}
		else if (bPressed)
		{
			nState = PBS_PRESSED;
		}
		else if (bHighlighted)
		{
			nState = PBS_HOT;
		}

		rect.bottom--;

		(*m_pfDrawThemeBackground) (m_hThemeButton, pDC->GetSafeHdc(), BP_PUSHBUTTON, nState, &rect, 0);
	}
	else
	{
		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect (rect, globalData.clrBtnFace, bPressed ? globalData.clrBtnShadow : globalData.clrBtnHilite);
		}
		else
		{
			pDC->FillRect (rect, &globalData.brBtnFace);

			if (bPressed)
			{
				pDC->Draw3dRect (rect, globalData.clrBtnShadow, globalData.clrBtnHilite);
			}
			else
			{
				pDC->Draw3dRect (rect, globalData.clrBtnHilite, globalData.clrBtnShadow);
			}
		}
	}

	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdArowDown, rect,
		bDisabled ? CBCGPMenuImages::ImageGray : CBCGPMenuImages::ImageDkGray);
}
//**************************************************************************************
void CBCGPVisualManager::OnDrawExpandingBox (CDC* pDC, CRect rect, BOOL bIsOpened,
											COLORREF colorBox)
{
	ASSERT_VALID(pDC);

	pDC->Draw3dRect (rect, colorBox, colorBox);

	rect.DeflateRect (2, 2);

	CPen penLine (PS_SOLID, 1, colorBox);
	CPen* pOldPen = pDC->SelectObject (&penLine);

	CPoint ptCenter = rect.CenterPoint ();

	pDC->MoveTo (rect.left, ptCenter.y);
	pDC->LineTo (rect.right, ptCenter.y);

	if (!bIsOpened)
	{
		pDC->MoveTo (ptCenter.x, rect.top);
		pDC->LineTo (ptCenter.x, rect.bottom);
	}

	pDC->SelectObject (pOldPen);
}

#ifndef BCGP_EXCLUDE_PROP_LIST

COLORREF CBCGPVisualManager::OnFillPropList(CDC* pDC, CBCGPPropList* /*pList*/, const CRect& rectClient, COLORREF& clrFill)
{
	pDC->FillRect (rectClient, &globalData.brWindow);
	
	clrFill = globalData.clrWindow;
	return globalData.clrWindowText;
}
//***********************************************************************************
void CBCGPVisualManager::OnFillPropListToolbarArea(CDC* pDC, CBCGPPropList* pList, const CRect& rectToolBar)
{
	if (pList == NULL || pList->DrawControlBarColors())
	{
		pDC->FillRect(rectToolBar, &globalData.brBarFace);
	}
	else
	{
		pDC->FillRect(rectToolBar, &globalData.brBtnFace);
	}
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawPropListComboButton(CDC* pDC, CRect rect, CBCGPProp* /*pProp*/, BOOL /*bControlBarColors*/,
												BOOL /*bFocused*/, BOOL bEnabled, BOOL bButtonIsDown, BOOL bButtonIsHighlighted)
{
	if (DrawComboDropButtonWinXP(pDC, rect, !bEnabled, bButtonIsDown, bButtonIsHighlighted))
	{
		return;
	}

	CBCGPToolbarComboBoxButton dummy;
	OnDrawComboDropButton (pDC, rect, !bEnabled, bButtonIsDown, bButtonIsHighlighted, &dummy);
}
//***********************************************************************************
COLORREF CBCGPVisualManager::OnDrawPropListPushButton(CDC* pDC, CRect rect, CBCGPProp* /*pProp*/, BOOL bControlBarColors,
												BOOL /*bFocused*/, BOOL bEnabled, BOOL bButtonIsDown, BOOL bButtonIsHighlighted)
{
	if (m_hThemeButton == NULL)
	{
		CBCGPToolbarComboBoxButton dummy;
		
		CBCGPVisualManager::BCGBUTTON_STATE state = bButtonIsDown ? 
			CBCGPVisualManager::ButtonsIsPressed : 
			bButtonIsHighlighted ? CBCGPVisualManager::ButtonsIsHighlighted : CBCGPVisualManager::ButtonsIsRegular;
		
		pDC->FillRect (rect, bControlBarColors ? &globalData.brBarFace : &globalData.brBtnFace);
		OnDrawButtonBorder (pDC, &dummy, rect, state);
		
		return bEnabled ? globalData.clrBarText : globalData.clrGrayedText;
	}

	int nState = PBS_NORMAL;

	if (!bEnabled)
	{
		nState = PBS_DISABLED;
	}
	else if (bButtonIsDown)
	{
		nState = PBS_PRESSED;
	}
	else if (bButtonIsHighlighted)
	{
		nState = PBS_HOT;
	}

	(*m_pfDrawThemeBackground) (m_hThemeButton, pDC->GetSafeHdc(), BP_PUSHBUTTON, nState, &rect, 0);

	COLORREF clrText = (COLORREF)-1;
	(*m_pfGetThemeColor) (m_hThemeButton, BP_PUSHBUTTON, nState, TMT_TEXTCOLOR, &clrText);

	return clrText;
}
//***********************************************************************************
COLORREF CBCGPVisualManager::OnFillPropertyListSelectedItem(CDC* pDC, CBCGPProp* pProp, CBCGPPropList* /*pWndList*/, const CRect& rectFill, BOOL bFocused)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID(pProp);

	if (!bFocused)
	{
		pDC->FillRect (rectFill, &globalData.brBarFace);
		return pProp->IsEnabled() ? globalData.clrBarText : globalData.clrGrayedText;
	}

	pDC->FillRect (rectFill, &globalData.brHilite);
	return pProp->IsEnabled() ? globalData.clrTextHilite : globalData.clrGrayedText;
}
//***********************************************************************************
COLORREF CBCGPVisualManager::OnFillPropListDescriptionArea(CDC* pDC, CBCGPPropList* pList, const CRect& rect)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pList);

	pDC->FillRect(rect, pList->DrawControlBarColors () ? &globalData.brBarFace : &globalData.brBtnFace);

	return pList->DrawControlBarColors () ? globalData.clrBarShadow : globalData.clrBtnShadow;
}
//***********************************************************************************
COLORREF CBCGPVisualManager::OnFillPropListCommandsArea(CDC* pDC, CBCGPPropList* pList, const CRect& rect)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pList);

	pDC->FillRect(rect, pList->DrawControlBarColors () ? &globalData.brBarFace : &globalData.brBtnFace);

	return pList->DrawControlBarColors () ? globalData.clrBarShadow : globalData.clrBtnShadow;
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawPropListMoreDescriptionIndicator(CDC* pDC, CBCGPPropList* pList, const CRect& rect)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pList);

	OnFillPropListDescriptionArea(pDC, pList, rect);

	CRect rectEllipse = rect;
	rectEllipse.DeflateRect(2, 2);

	CBCGPDrawManager dm(*pDC);

	dm.DrawEllipse(rectEllipse, pList->DrawControlBarColors () ? globalData.clrBarHilite : globalData.clrWindow, globalData.clrHilite);

	int x = rectEllipse.CenterPoint().x;
	int y = rectEllipse.top + 1;

	COLORREF clrText = GetPropListDisabledTextColor(pList);
	if (clrText == (COLORREF)-1)
	{
		clrText = globalData.clrGrayedText;
	}

	dm.DrawLine(x, y + 2, x, y + 3, clrText);

	if (rectEllipse.Height() > 10)
	{
		dm.DrawLine(x, y + 5, x, rectEllipse.bottom - 4, clrText);
	}
}

#endif

void CBCGPVisualManager::OnFillPropSheetHeaderArea(CDC* pDC, CBCGPPropertySheet* /*pPropSheet*/, CRect rect, BOOL& /*bDrawBottomLine*/)
{
	ASSERT_VALID (pDC);

	pDC->FillRect(rect, &globalData.brBarFace);
}
//***********************************************************************************
COLORREF CBCGPVisualManager::OnFillCalendarBarNavigationRow (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	CBrush br (globalData.clrInactiveCaption);
	pDC->FillRect (rect, &br);

	pDC->Draw3dRect (rect, globalData.clrInactiveBorder, globalData.clrInactiveBorder);
	return globalData.clrInactiveCaptionText;
}
//*************************************************************************************
void CBCGPVisualManager::GetCalendarColors (const CBCGPCalendar* pCalendar,
				   CBCGPCalendarColors& colors)
{
	colors.clrCaption = globalData.clrBtnFace;
	colors.clrSelected = globalData.IsHighContastMode () ? globalData.clrBtnText : globalData.clrBtnFace;
	colors.clrTodayBorder = RGB (187, 85, 3);

	if (pCalendar->GetSafeHwnd() != NULL && !pCalendar->IsWindowEnabled())
	{
		colors.clrCaptionText = globalData.clrGrayedText;
		colors.clrSelectedText = globalData.clrGrayedText;
		colors.clrText = globalData.clrGrayedText;
		colors.clrLine = globalData.clrBtnLight;
	}
	else
	{
		colors.clrCaptionText = globalData.clrBtnText;
		colors.clrSelectedText = globalData.IsHighContastMode () ? globalData.clrBtnFace : globalData.clrBtnText;
	}
}
//*************************************************************************************
void CBCGPVisualManager::OnDrawControlBorder (CWnd* pWndCtrl)
{
	ASSERT_VALID (pWndCtrl);

	CWindowDC dc (pWndCtrl);

	CRect rect;
	pWndCtrl->GetWindowRect (rect);

	rect.bottom -= rect.top;
	rect.right -= rect.left;
	rect.left = rect.top = 0;

	OnDrawControlBorder (&dc, rect, pWndCtrl, CBCGPToolBarImages::m_bIsDrawOnGlass);
}
//*************************************************************************************
void CBCGPVisualManager::OnDrawControlBorder (CDC* pDC, CRect rect, CWnd* pWndCtrl, BOOL bDrawOnGlass)
{
	ASSERT_VALID (pDC);

	if (bDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawRect (rect, (COLORREF)-1, globalData.clrBarShadow);

		rect.DeflateRect (1, 1);

		dm.DrawRect (rect, (COLORREF)-1, globalData.clrWindow);
	}
	else
	{
		if (pWndCtrl->GetSafeHwnd () != NULL &&
			(pWndCtrl->GetStyle () & WS_POPUP))
		{
			pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarShadow);
		}
		else
		{
			pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarLight);
		}

		rect.DeflateRect (1, 1);
		pDC->Draw3dRect (rect, globalData.clrWindow, globalData.clrWindow);
	}
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawControlBorderNoTheme (CWnd* pWndCtrl)
{
	ASSERT_VALID (pWndCtrl);

	CWindowDC dc (pWndCtrl);

	CRect rect;
	pWndCtrl->GetWindowRect (rect);

	rect.bottom -= rect.top;
	rect.right -= rect.left;
	rect.left = rect.top = 0;

	OnDrawControlBorderNoTheme (&dc, rect, pWndCtrl, CBCGPToolBarImages::m_bIsDrawOnGlass);
}
//*************************************************************************************
void CBCGPVisualManager::OnDrawControlBorderNoTheme (CDC* pDC, CRect rect, CWnd* pWndCtrl, BOOL bDrawOnGlass)
{
	UNREFERENCED_PARAMETER(pWndCtrl);

	ASSERT_VALID (pDC);

	COLORREF clrBorder = globalData.clrBtnShadow;
	COLORREF clrWinBorder = (COLORREF)-1;
	
	if (m_hThemeComboBox != NULL)
	{
		if ((*m_pfGetThemeColor) (m_hThemeComboBox, 5, 0, TMT_BORDERCOLOR, &clrWinBorder) == S_OK)
		{
			clrBorder = clrWinBorder;
		}
	}

	if (bDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawRect (rect, (COLORREF)-1, clrBorder);

		rect.DeflateRect (1, 1);
		dm.DrawRect (rect, (COLORREF)-1, globalData.clrWindow);
	}
	else
	{
		pDC->Draw3dRect (rect, clrBorder, clrBorder);

		rect.DeflateRect (1, 1);
		pDC->Draw3dRect (rect, globalData.clrWindow, globalData.clrWindow);
	}
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawShowAllMenuItems (CDC* pDC, CRect rect, 
												 CBCGPVisualManager::BCGBUTTON_STATE /*state*/)
{
	ASSERT_VALID (pDC);
	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdArowShowAll, rect);
}
//************************************************************************************
int CBCGPVisualManager::GetShowAllMenuItemsHeight (CDC* /*pDC*/, const CSize& /*sizeDefault*/)
{
	return CBCGPMenuImages::Size ().cy + 2 * globalUtils.ScaleByDPI(TEXT_MARGIN);
}
//**********************************************************************************
int CBCGPVisualManager::GetTabsMargin (const CBCGPTabWnd* /*pTabWnd*/) 
{ 
	return globalUtils.ScaleByDPI(2); 
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTabUnderline(CDC* pDC, CBCGPTabWnd* pTabWnd, CRect rect, int /*iTab*/)
{
	COLORREF clrFill = globalData.clrHilite;

	if (pTabWnd->IsMDITabGroup() && !(pTabWnd->IsMDIFocused() && pTabWnd->IsActiveInMDITabGroup()))
	{
		clrFill = globalData.clrBarShadow;
	}

	CBrush brFill(clrFill);
	pDC->FillRect(rect, &brFill);
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawTabSplitter(CDC* pDC, CBCGPTabWnd* pTabWnd, CRect rectTabSplitter)
{
	UNREFERENCED_PARAMETER(pTabWnd);

	pDC->FillRect(rectTabSplitter, &globalData.brBarFace);
	
	pDC->Draw3dRect (rectTabSplitter, globalData.clrBarDkShadow, globalData.clrBarDkShadow);
	rectTabSplitter.DeflateRect (1, 1);
	pDC->Draw3dRect (rectTabSplitter, globalData.clrBarHilite, globalData.clrBarShadow);
}
//**********************************************************************************
void CBCGPVisualManager::GetTabFrameColors (const CBCGPBaseTabWnd* pTabWnd,
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

	COLORREF clrActiveTab = pTabWnd->GetTabBkColor (pTabWnd->GetActiveTab ());

	if (pTabWnd->IsOneNoteStyle () && clrActiveTab != (COLORREF)-1)
	{
		clrFace = clrActiveTab;
	}
	else if (pTabWnd->IsDialogControl (TRUE))
	{
		clrFace = globalData.clrBtnFace;
	}
	else
	{
		clrFace = globalData.clrBarFace;
	}

	if (pTabWnd->IsDialogControl(TRUE))
	{
		clrDark = globalData.clrBtnShadow;
		clrBlack = globalData.clrBtnText;
		clrHighlight = pTabWnd->IsVS2005Style () ? globalData.clrBtnShadow : globalData.clrBtnHilite;
		clrDarkShadow = globalData.clrBtnDkShadow;
		clrLight = globalData.clrBtnLight;

		pbrFace = &globalData.brBtnFace;
	}
	else
	{
		clrDark = globalData.clrBarShadow;
		clrBlack = globalData.clrBarText;
		clrHighlight = pTabWnd->IsVS2005Style () ? globalData.clrBarShadow : globalData.clrBarHilite;
		clrDarkShadow = globalData.clrBarDkShadow;
		clrLight = globalData.clrBarLight;

		pbrFace = &globalData.brBarFace;
	}

	pbrBlack = &globalData.brBlack;
}
//*********************************************************************************
void CBCGPVisualManager::OnFillAutoHideButtonBackground (CDC* pDC, CRect rect, CBCGPAutoHideButton* /*pButton*/)
{
	ASSERT_VALID (pDC);
	pDC->FillRect (rect, &globalData.brBarFace);
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawAutoHideButtonBorder (CDC* pDC, CRect rectBounds, CRect rectBorderSize, CBCGPAutoHideButton* pButton)
{
	ASSERT_VALID (pDC);

	COLORREF clr = GetAutoHideButtonBorderColor(pButton);

	COLORREF clrText = pDC->GetTextColor ();

	if (rectBorderSize.left > 0)
	{
		pDC->FillSolidRect (rectBounds.left, rectBounds.top, 
							rectBounds.left + rectBorderSize.left, 
							rectBounds.bottom, clr);
	}
	if (rectBorderSize.top > 0)
	{
		pDC->FillSolidRect (rectBounds.left, rectBounds.top, 
							rectBounds.right, 
							rectBounds.top + rectBorderSize.top, clr);
	}
	if (rectBorderSize.right > 0)
	{
		pDC->FillSolidRect (rectBounds.right - rectBorderSize.right, rectBounds.top, 
							rectBounds.right, 
							rectBounds.bottom, clr);
	}
	if (rectBorderSize.bottom > 0)
	{
		pDC->FillSolidRect (rectBounds.left, rectBounds.bottom - rectBorderSize.bottom,  
							rectBounds.right, 
							rectBounds.bottom, clr);
	}

	pDC->SetTextColor (clrText);
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetAutoHideButtonTextColor (CBCGPAutoHideButton* /*pButton*/)
{
	return globalData.clrBarText;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetAutoHideButtonBorderColor(CBCGPAutoHideButton* /*pButton*/)
{
	return globalData.clrBarShadow;
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawOutlookBarSplitter (CDC* pDC, CRect rectSplitter)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rectSplitter, &globalData.brBarFace);
	pDC->Draw3dRect (rectSplitter, globalData.clrBarHilite, globalData.clrBarShadow);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawOutlookBarFrame(CDC* pDC, CRect rectFrame)
{
	ASSERT_VALID (pDC);
	pDC->Draw3dRect(rectFrame, globalData.clrBarShadow, globalData.clrBarShadow);
}
//********************************************************************************
void CBCGPVisualManager::OnFillOutlookBarCaption (CDC* pDC, CRect rectCaption, COLORREF& clrText)
{
	ASSERT_VALID (pDC);

	rectCaption.InflateRect(1, 0);

	pDC->FillRect(rectCaption, &globalData.brBarFace);
	pDC->Draw3dRect(rectCaption, globalData.clrBarShadow, globalData.clrBarShadow);

	clrText = globalData.clrBarText;
}
//********************************************************************************
COLORREF CBCGPVisualManager::OnDrawOutlookPopupButton(CDC* pDC, CRect& rectBtn, BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsOnCaption)
{
	UNREFERENCED_PARAMETER(bIsHighlighted);
	UNREFERENCED_PARAMETER(bIsOnCaption);

	ASSERT_VALID(pDC);
	
	pDC->FillRect(rectBtn, &globalData.brBarFace);

	if (bIsPressed)
	{
		pDC->Draw3dRect(rectBtn, globalData.clrBarShadow, globalData.clrBarHilite);
	}
	else
	{
		pDC->Draw3dRect(rectBtn, globalData.clrBarHilite, globalData.clrBarShadow);
	}
	
	return globalData.clrBarText;
}
//********************************************************************************
BOOL CBCGPVisualManager::OnDrawCalculatorButton (CDC* pDC, 
	CRect rect, CBCGPToolbarButton* /*pButton*/, 
	CBCGPVisualManager::BCGBUTTON_STATE state, 
	int /*cmd*/ /* CBCGPCalculator::CalculatorCommands */,
	CBCGPCalculator* /*pCalculator*/)
{
	ASSERT_VALID (pDC);

	switch (state)
	{
	case ButtonsIsPressed:
		pDC->FillRect (rect, &globalData.brLight);
		pDC->Draw3dRect (&rect, globalData.clrBarShadow, globalData.clrBarHilite);
		break;

	case ButtonsIsHighlighted:
		pDC->FillRect (rect, &globalData.brLight);

	default:
		pDC->Draw3dRect (&rect, globalData.clrBarHilite, globalData.clrBarShadow);
	}

	return TRUE;
}
//********************************************************************************
BOOL CBCGPVisualManager::OnDrawCalculatorDisplay (CDC* pDC, CRect rect, 
												  const CString& /*strText*/, BOOL /*bMem*/,
												  CBCGPCalculator* /*pCalculator*/)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brWindow);
	pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarHilite);

	return TRUE;
}
//*********************************************************************************
COLORREF CBCGPVisualManager::GetCalculatorButtonTextColor (BOOL bIsUserCommand, int nCmd /* CBCGPCalculator::CalculatorCommands */)
{
	if (globalData.m_bIsBlackHighContrast || globalData.m_bIsWhiteHighContrast)
	{
		return globalData.clrBarText;
	}

	CBCGPCalculator::CalculatorCommands cmd = (CBCGPCalculator::CalculatorCommands) nCmd;

	COLORREF clrText = RGB (0, 0, 255);

	if (bIsUserCommand)
	{
		clrText = RGB (255, 0, 255);
	}
	else if (cmd != CBCGPCalculator::idCommandNone && cmd != CBCGPCalculator::idCommandDot)
	{
		clrText = RGB (255, 0, 0);
	}

	return clrText;
}
//*********************************************************************************
BOOL CBCGPVisualManager::OnDrawBrowseButton (CDC* pDC, CRect rect, 
											 CBCGPEdit* pEdit, 
											 CBCGPVisualManager::BCGBUTTON_STATE state,
											 COLORREF& /*clrText*/)
{
	ASSERT_VALID (pDC);

	if (pEdit != NULL && pEdit->m_bOnGlass)
	{
		CBCGPDrawManager dm(*pDC);
		dm.DrawRect(rect, state == ButtonsIsPressed ? globalData.clrBtnShadow : globalData.clrBtnFace, globalData.clrBtnText);
		return TRUE;
	}

	pDC->FillRect (&rect, &globalData.brBtnFace);

	CRect rectFrame = rect;
	rectFrame.InflateRect (0, 1, 1, 1);

	pDC->Draw3dRect (rectFrame, globalData.clrBtnDkShadow, globalData.clrBtnDkShadow);

	rectFrame.DeflateRect (1, 1);
	pDC->DrawEdge (rectFrame, state == ButtonsIsPressed ? BDR_SUNKENINNER : BDR_RAISEDINNER, BF_RECT);

	return TRUE;
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawEditClearButton(CDC* pDC, CBCGPButton* pButton, CRect rect)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	pDC->FillRect (rect, &globalData.brBtnFace);

	rect.DeflateRect (1, 1);
	pDC->DrawEdge (rect, pButton->IsPressed() ? BDR_SUNKENINNER : BDR_RAISEDINNER, BF_RECT);
}
//*********************************************************************************
CBrush& CBCGPVisualManager::GetEditCtrlBackgroundBrush(CBCGPEdit* pEdit)
{
	if (pEdit->GetSafeHwnd() != NULL && ((pEdit->GetStyle() & ES_READONLY) == ES_READONLY || !pEdit->IsWindowEnabled()))
	{
		return GetDlgBackBrush (pEdit->GetParent());
	}

	return globalData.brWindow;
}
//*********************************************************************************
COLORREF CBCGPVisualManager::GetEditCtrlTextColor(CBCGPEdit* /*pEdit*/)
{
	return globalData.clrWindowText;
}
//*********************************************************************************
void CBCGPVisualManager::GetEditSimplifiedIconColors(COLORREF& clrSimple, COLORREF& clrHot, BOOL /*bIsCloseButton*/)
{
	clrSimple = CBCGPVisualManager::GetInstance ()->GetToolbarEditPromptColor();
	clrHot = globalData.clrHotText;
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawEditCtrlResizeBox (CDC* pDC, CBCGPEditCtrl* /*pEdit*/, CRect rect)
{
	CBrush brBox(globalData.clrBarLight);
	pDC->FillRect (rect, &brBox);
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawAppBarBorder (CDC* pDC, CBCGPAppBarWnd* /*pAppBarWnd*/,
									CRect rectBorder, CRect rectBorderSize)
{
	ASSERT_VALID (pDC);

	CBrush* pOldBrush = pDC->SelectObject (&globalData.brBtnFace);
	ASSERT (pOldBrush != NULL);

	pDC->PatBlt (rectBorder.left, rectBorder.top, rectBorderSize.left, rectBorder.Height (), PATCOPY);
	pDC->PatBlt (rectBorder.left, rectBorder.top, rectBorder.Width (), rectBorderSize.top, PATCOPY);
	pDC->PatBlt (rectBorder.right - rectBorderSize.right, rectBorder.top, rectBorderSize.right, rectBorder.Height (), PATCOPY);
	pDC->PatBlt (rectBorder.left, rectBorder.bottom - rectBorderSize.bottom, rectBorder.Width (), rectBorderSize.bottom, PATCOPY);

	rectBorderSize.DeflateRect (2, 2);
	rectBorder.DeflateRect (2, 2);

	pDC->SelectObject (&globalData.brLight);

	pDC->PatBlt (rectBorder.left, rectBorder.top + 1, rectBorderSize.left, rectBorder.Height () - 2, PATCOPY);
	pDC->PatBlt (rectBorder.left + 1, rectBorder.top, rectBorder.Width () - 2, rectBorderSize.top, PATCOPY);
	pDC->PatBlt (rectBorder.right - rectBorderSize.right, rectBorder.top + 1, rectBorderSize.right, rectBorder.Height () - 2, PATCOPY);
	pDC->PatBlt (rectBorder.left + 1, rectBorder.bottom - rectBorderSize.bottom, rectBorder.Width () - 2, rectBorderSize.bottom, PATCOPY);

	pDC->SelectObject (pOldBrush);
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawAppBarCaption (CDC* pDC, CBCGPAppBarWnd* /*pAppBarWnd*/, 
											CRect rectCaption, CString strCaption)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rectCaption, &globalData.brBarFace);

	// Paint caption text:
	int nOldMode = pDC->SetBkMode (TRANSPARENT);
	COLORREF clrOldText = pDC->SetTextColor (globalData.clrBarText);

	CFont* pOldFont = pDC->SelectObject(UseLargeCaptionFontInDockingCaptions() ? &globalData.fontCaption : &globalData.fontBold);
	ASSERT_VALID (pOldFont);

	CRect rectText = rectCaption;
	rectText.DeflateRect (globalUtils.ScaleByDPI(4), 0);

	pDC->DrawText (strCaption, rectText, DT_LEFT | DT_SINGLELINE | DT_VCENTER);

	pDC->SelectObject (pOldFont);
	pDC->SetBkMode (nOldMode);
	pDC->SetTextColor (clrOldText);
}
//******************************************************************************
void CBCGPVisualManager::GetSmartDockingBaseMarkerColors (
		COLORREF& clrBaseGroupBackground,
		COLORREF& clrBaseGroupBorder)
{
	clrBaseGroupBackground = globalData.clrBarFace;
	clrBaseGroupBorder = globalData.clrBarShadow;
}
//******************************************************************************
COLORREF CBCGPVisualManager::GetSmartDockingMarkerToneColor ()
{
	return globalData.clrActiveCaption;
}

#ifndef BCGP_EXCLUDE_TOOLBOX

BOOL CBCGPVisualManager::OnEraseToolBoxButton (CDC* pDC, CRect rect, 
											CBCGPToolBoxButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->GetCheck ())
	{
		rect.DeflateRect (1, 1);
		CBCGPToolBarImages::FillDitheredRect (pDC, rect);
	}

	return TRUE;
}
//**********************************************************************************
BOOL CBCGPVisualManager::OnDrawToolBoxButtonBorder (CDC* pDC, CRect& rect, 
												 CBCGPToolBoxButton* pButton, UINT /*uiState*/)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->GetCheck ())
	{
		pDC->Draw3dRect (&rect, globalData.clrBarShadow, globalData.clrBarHilite);
	}
	else if (pButton->IsHighlighted ())
	{
		pDC->Draw3dRect (&rect, globalData.clrBarHilite, globalData.clrBarShadow);
	}

	return TRUE;
}
//**********************************************************************************
COLORREF CBCGPVisualManager::GetToolBoxButtonTextColor (CBCGPToolBoxButton* /*pButton*/)
{
	return globalData.clrWindowText;
}

#endif // BCGP_EXCLUDE_TOOLBOX

#ifndef BCGP_EXCLUDE_POPUP_WINDOW

void CBCGPVisualManager::OnFillPopupWindowBackground (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);
	pDC->FillRect (rect, &globalData.brBarFace);
}
//**********************************************************************************
COLORREF CBCGPVisualManager::GetPopupWindowBorderBorderColor()
{
	return globalData.clrBarDkShadow;
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawPopupWindowBorder (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarDkShadow);
	rect.DeflateRect (1, 1);
	pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarShadow);
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawPopupWindowRoundedBorder (CDC* pDC, CRect rect, CBCGPPopupWindow* /*pPopupWnd*/, int nCornerRadius)
{
	ASSERT_VALID (pDC);

	CPen pen (PS_SOLID, 1, globalData.clrBarShadow);
	CPen* pOldPen = pDC->SelectObject (&pen);
	CBrush* pOldBrush = (CBrush*) pDC->SelectStockObject (NULL_BRUSH);
	
	pDC->RoundRect (rect.left, rect.top, rect.right, rect.bottom, nCornerRadius + 1, nCornerRadius + 1);

	rect.DeflateRect(1, 1);
	
	CPen pen1(PS_SOLID, 1, globalData.clrBarLight);
	pDC->SelectObject (&pen1);
	
	pDC->RoundRect (rect.left, rect.top, rect.right, rect.bottom, nCornerRadius, nCornerRadius);

	pDC->SelectObject (pOldPen);
	pDC->SelectObject (pOldBrush);
}
//**********************************************************************************
COLORREF CBCGPVisualManager::OnDrawPopupWindowCaption (CDC* pDC, CRect rectCaption, CBCGPPopupWindow* pPopupWnd)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID(pPopupWnd);

	switch (pPopupWnd->GetStemLocation())
	{
	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_TopCenter:
	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_TopLeft:
	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_TopRight:
		rectCaption.top -= pPopupWnd->GetStemSize();
		break;

	case CBCGPPopupWindow::BCGPPopupWindowStemLocation_Left:
		rectCaption.left -= pPopupWnd->GetStemSize();
		break;
	}

	if (!pPopupWnd->HasSmallCaption () || pPopupWnd->IsSmallCaptionGripper())
	{
		pDC->FillRect(rectCaption, &globalData.brActiveCaption);
		return globalData.clrCaptionText;
	}
	else
	{
		OnFillPopupWindowBackground(pDC, rectCaption);
		return globalData.clrBarText;
	}
}
//**********************************************************************************
COLORREF CBCGPVisualManager::GetPopupWindowCaptionTextColor(CBCGPPopupWindow* /*pPopupWnd*/, BOOL /*bButton*/)
{
	return globalData.clrCaptionText;
}
//**********************************************************************************
COLORREF CBCGPVisualManager::GetPopupWindowTextColor(CBCGPPopupWindow* /*pPopupWnd*/)
{
	return globalData.clrBarText;
}
//**********************************************************************************
COLORREF CBCGPVisualManager::GetPopupWindowLinkTextColor(CBCGPPopupWindow* /*pPopupWnd*/, BOOL bIsHover)
{
	return bIsHover ? globalData.clrHotLinkText : globalData.clrHotText;
}
//**********************************************************************************
void CBCGPVisualManager::OnErasePopupWindowButton (CDC* pDC, CRect /*rect*/, CBCGPPopupWndButton* pButton)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	globalData.DrawParentBackground(pButton, pDC);
}
//**********************************************************************************
void CBCGPVisualManager::OnDrawPopupWindowButtonBorder (CDC* pDC, CRect rect, CBCGPPopupWndButton* pButton)
{
	if (pButton->IsPressed ())
	{
		if (globalData.IsHighContastMode())
		{
			pDC->Draw3dRect (rect, globalData.clrHilite, globalData.clrHilite);
		}
		else
		{
			pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarLight);
			rect.DeflateRect (1, 1);
			pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarHilite);
		}
	}
	else
	{
		pDC->Draw3dRect (rect, globalData.clrBarLight, globalData.clrBarDkShadow);
		rect.DeflateRect (1, 1);
		pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarShadow);
	}
}

#endif // BCGP_EXCLUDE_POPUP_WINDOW

#ifndef BCGP_EXCLUDE_PLANNER

//*******************************************************************************
COLORREF CBCGPVisualManager::OnFillPlannerCaption (CDC* pDC,
		CBCGPPlannerView* pView, CRect rect, BOOL bIsToday, BOOL bIsSelected,
		BOOL bNoBorder/* = FALSE*/, BOOL /*bHorz = TRUE*/)
{
	ASSERT_VALID (pDC);

	COLORREF clrText = globalData.clrBtnText;

	const BOOL bMonth = DYNAMIC_DOWNCAST(CBCGPPlannerViewMonth, pView) != NULL;
	if (bMonth && !bIsToday && !bIsSelected)
	{
		return clrText;
	}

	rect.DeflateRect (1, 1);

	if (bIsToday)
	{
		pDC->FillRect (rect, &globalData.brBtnFace);
		CBCGPToolBarImages::FillDitheredRect (pDC, rect);
	}
	else if (bIsSelected)
	{
		pDC->FillRect (rect, &globalData.brHilite);
		clrText = globalData.clrTextHilite;
	}
	else
	{
		pDC->FillRect (rect, &globalData.brBtnFace);
	}

	if (bIsToday || !bNoBorder)
	{
		CPen pen (PS_SOLID, 1, bIsToday ? globalData.clrBarShadow : globalData.clrBtnShadow);
		CPen* pOldPen = pDC->SelectObject (&pen);

		pDC->MoveTo (rect.left, rect.bottom);
		pDC->LineTo (rect.right, rect.bottom);

		pDC->SelectObject (pOldPen);
	}

	return clrText;
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawPlannerCaptionText (CDC* pDC, 
		CBCGPPlannerView* /*pView*/, CRect rect, const CString& strText, 
		COLORREF clrText, int nAlign, BOOL bHighlight)
{
	const int nTextMargin = 2;

	if (bHighlight)
	{
		const int nTextLen = min (rect.Width (), 
			pDC->GetTextExtent (strText).cx + 2 * nTextMargin);

		CRect rectHighlight = rect;
		rectHighlight.DeflateRect (1, 1);

		switch (nAlign)
		{
		case DT_LEFT:
			rectHighlight.right = rectHighlight.left + nTextLen;
			break;

		case DT_RIGHT:
			rectHighlight.left = rectHighlight.right - nTextLen;
			break;

		case DT_CENTER:
			rectHighlight.left = rectHighlight.CenterPoint ().x - nTextLen / 2;
			rectHighlight.right = rectHighlight.left + nTextLen;
			break;
		}

		pDC->FillRect (rectHighlight, &globalData.brHilite);

		clrText = globalData.clrTextHilite;
	}

	rect.DeflateRect (nTextMargin, 0);

	COLORREF clrOld = pDC->SetTextColor (clrText);
	pDC->DrawText (strText, rect, DT_SINGLELINE | DT_VCENTER | nAlign);

	pDC->SetTextColor (clrOld);
}
//*******************************************************************************
void CBCGPVisualManager::GetPlannerAppointmentColors (CBCGPPlannerView* pView,
		BOOL bSelected, BOOL bSimple, DWORD dwDrawFlags, 
		COLORREF& clrBack1, COLORREF& clrBack2,
		COLORREF& clrFrame1, COLORREF& clrFrame2, COLORREF& clrText)
{
	ASSERT_VALID (pView);

	const BOOL bIsGradientFill = 
		dwDrawFlags & BCGP_PLANNER_DRAW_APP_GRADIENT_FILL &&
		globalData.m_nBitsPerPixel > 8 && 
		!(globalData.m_bIsBlackHighContrast || globalData.m_bIsWhiteHighContrast);
	const BOOL bIsOverrideSelection = (dwDrawFlags & BCGP_PLANNER_DRAW_APP_OVERRIDE_SELECTION) ==
		BCGP_PLANNER_DRAW_APP_OVERRIDE_SELECTION;
	const BOOL bIsDefaultBkgnd      = (dwDrawFlags & BCGP_PLANNER_DRAW_APP_DEFAULT_BKGND) ==
		BCGP_PLANNER_DRAW_APP_DEFAULT_BKGND;

	const BOOL bSelect = bSelected && !bIsOverrideSelection;

	const BOOL bDayView = pView->IsKindOf (RUNTIME_CLASS (CBCGPPlannerViewDay));

	if (bSelect)
	{
		if (bDayView)
		{
			if (!bSimple)
			{
				clrBack1 = globalData.clrBtnFace;
				clrText  = globalData.clrBtnText;
			}
		}
		else
		{
			clrBack1 = globalData.clrHilite;
			clrText  = globalData.clrTextHilite;
		}
	}

	BOOL bDefault = FALSE;

	if (clrBack1 == CLR_DEFAULT)
	{
		if (bIsGradientFill || bIsDefaultBkgnd)
		{
			clrBack1 = pView->GetPlanner ()->GetBackgroundColor ();
			if (clrBack1 == CLR_DEFAULT)
			{
				bDefault = TRUE;
				clrBack1 = m_clrPlannerWork;
			}
		}
		else
		{
			clrBack1 = (bDayView || !bSimple)
						? globalData.clrWindow
						: CLR_DEFAULT;
		}
	}

	clrBack2 = clrBack1;
	
	if (clrText == CLR_DEFAULT)
	{
		clrText = (bDefault && !bIsGradientFill) ? globalData.clrBarText : globalData.clrWindowText;
	}
	
	clrFrame1 = globalData.clrWindowFrame;
	clrFrame2 = clrFrame1;

	if (bIsGradientFill && clrBack1 != CLR_DEFAULT)
	{
		if (!bSelect)
		{
			clrBack2 = RGB (255, 255, 255);

			if (!bDefault)
			{
				clrBack2 = CBCGPDrawManager::PixelAlpha (clrBack1, clrBack2, 25);
			}
		}

		clrFrame1 = CBCGPDrawManager::PixelAlpha (clrBack1, RGB (0, 0, 0), 75);

		if (!bSelected)
		{
			clrFrame2 = clrFrame1;
		}
	}
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetPlannerAppointmentDurationColor (CBCGPPlannerView* /*pView*/, const CBCGPAppointment* pApp) const
{
	return pApp->GetDurationColor() == CLR_DEFAULT
			? globalData.clrWindow
			: pApp->GetDurationColor();
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetPlannerAppointmentTimeColor (CBCGPPlannerView* /*pView*/,
		BOOL /*bSelected*/, BOOL /*bSimple*/, DWORD /*dwDrawFlags*/)
{
	return CLR_DEFAULT;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetPlannerHourLineColor (CBCGPPlannerView* /*pView*/,
		BOOL /*bWorkingHours*/, BOOL /*bHour*/)
{
	return m_clrPalennerLine;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetPlannerSelectionColor (CBCGPPlannerView* /*pView*/)
{
	return globalData.clrHilite;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetPlannerSeparatorColor (CBCGPPlannerView* /*pView*/)
{
	return globalData.clrWindowFrame;
}
//*******************************************************************************
void CBCGPVisualManager::OnFillPlanner (CDC* pDC, CBCGPPlannerView* pView, 
		CRect rect, BOOL bWorkingArea)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pView);

	if (m_brPlanner.GetSafeHandle () == NULL)
	{
		WORD HatchBits [8] = { 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55 };

		CBitmap bmp;
		bmp.CreateBitmap (8, 8, 1, 1, HatchBits);

		CBrush brLight;
		m_brPlanner.CreatePatternBrush (&bmp);
	}

	COLORREF colorFill = globalData.IsHighContastMode () ?
		CLR_DEFAULT : pView->GetPlanner ()->GetBackgroundColor ();
	BOOL bIsDefaultColor = colorFill == CLR_DEFAULT;

	if (bIsDefaultColor)
	{
		colorFill = m_clrPlannerWork;	// Use default color
	}
		
	switch (pView->GetPlanner ()->GetType ())
	{
	case CBCGPPlannerManagerCtrl::BCGP_PLANNER_TYPE_DAY:
	case CBCGPPlannerManagerCtrl::BCGP_PLANNER_TYPE_WORK_WEEK:
	case CBCGPPlannerManagerCtrl::BCGP_PLANNER_TYPE_MULTI:
	case CBCGPPlannerManagerCtrl::BCGP_PLANNER_TYPE_SCHEDULE:
		{
			COLORREF clrTextOld = pDC->SetTextColor (colorFill);

			COLORREF clrBkOld = pDC->SetBkColor (bWorkingArea ? 
				RGB (255, 255, 255) : RGB (128, 128, 128));

			pDC->FillRect (rect, &m_brPlanner);

			pDC->SetTextColor (clrTextOld);
			pDC->SetBkColor (clrBkOld);
		}
		break;

	default:
		if (bIsDefaultColor || !bWorkingArea)
		{
			pDC->FillRect (rect, bWorkingArea ? 
				&globalData.brWindow : &globalData.brBtnFace);
		}
		else
		{
			CBrush br (colorFill);
			pDC->FillRect (rect, &br);
		}
	}
}
//*******************************************************************************
COLORREF CBCGPVisualManager::OnFillPlannerTimeBar (CDC* pDC, CBCGPPlannerView* /*pView*/, 
											   CRect rect, COLORREF& clrLine)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBtnFace);

	clrLine = globalData.clrBtnShadow;

	return globalData.clrBtnText;
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawPlannerTimeLine (CDC* pDC, CBCGPPlannerView* pView, CRect rect)
{
	ASSERT_VALID (pDC);

	if (DYNAMIC_DOWNCAST(CBCGPPlannerViewSchedule, pView) == NULL)
	{
		pDC->FillRect (rect, &globalData.brActiveCaption);

		CPen* pOldPen = pDC->SelectObject (&globalData.penBarShadow);

		pDC->MoveTo (rect.left, rect.bottom);
		pDC->LineTo (rect.right, rect.bottom);

		pDC->SelectObject (pOldPen);
	}
	else
	{
		pDC->FillRect (rect, &globalData.brActiveCaption);

		CPen* pOldPen = pDC->SelectObject (&globalData.penBarShadow);

		pDC->MoveTo (rect.right, rect.top);
		pDC->LineTo (rect.right, rect.bottom);

		pDC->SelectObject (pOldPen);
	}
}
//*******************************************************************************
void CBCGPVisualManager::OnFillPlannerWeekBar (CDC* pDC, CBCGPPlannerView* /*pView*/, CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBtnFace);
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawPlannerHeader (CDC* pDC, 
	CBCGPPlannerView* /*pView*/, CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBtnFace);
	pDC->Draw3dRect (rect, globalData.clrBtnHilite, globalData.clrBtnShadow);
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawPlannerHeaderPane (CDC* pDC, 
	CBCGPPlannerView* /*pView*/, CRect rect)
{
	ASSERT_VALID (pDC);
	pDC->Draw3dRect (rect, globalData.clrBtnShadow, globalData.clrBtnHilite);
}
//*******************************************************************************
void CBCGPVisualManager::OnFillPlannerHeaderAllDay (CDC* pDC, 
		CBCGPPlannerView* /*pView*/, CRect rect)
{
	ASSERT_VALID (pDC);
	
	CBrush br (globalData.clrBtnShadow);
	pDC->FillRect (rect, &br);
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawPlannerHeaderAllDayItem (CDC* pDC, 
		CBCGPPlannerView* /*pView*/, CRect rect, BOOL /*bIsToday*/, BOOL bIsSelected)
{
	if (bIsSelected)
	{
		pDC->FillRect (rect, &globalData.brWindow);
	}
}
//*******************************************************************************
void CBCGPVisualManager::PreparePlannerBackItem (BOOL /*bIsToday*/, BOOL /*bIsSelected*/)
{
	m_bPlannerBackItemToday    = FALSE;
	m_bPlannerBackItemSelected = FALSE;
}
//*******************************************************************************
void CBCGPVisualManager::PreparePlannerCaptionBackItem (BOOL bIsHeader)
{
	m_bPlannerCaptionBackItemHeader = bIsHeader;
}
//*******************************************************************************
UINT CBCGPVisualManager::GetPlannerWeekBarType() const
{
	return CBCGPPlannerView::BCGP_PLANNER_WEEKBAR_DAYS;
}
#endif // BCGP_EXCLUDE_PLANNER

//********************************************************************************
void CBCGPVisualManager::DoDrawHeaderSortArrow (CDC* pDC, CRect rectArrow, BOOL bIsUp, BOOL /*bDlgCtrl*/)
{
	ASSERT_VALID (pDC);

	if (m_hThemeHeader != NULL && globalData.bIsWindows7)
	{
		const int nState = bIsUp ? HSAS_SORTEDUP : HSAS_SORTEDDOWN;

		CRect rect = rectArrow;

		int nMinWidth = rect.Height() * 3 / 2;

		if (rect.Width() < nMinWidth)
		{
			rect.InflateRect(nMinWidth / 2, 0);
		}

		(*m_pfDrawThemeBackground) (m_hThemeHeader, pDC->GetSafeHdc(), HP_HEADERSORTARROW, nState, &rect, 0);
	}
	else
	{
		CBCGPMenuImages::Draw(pDC, bIsUp ? CBCGPMenuImages::IdArowUpLarge : CBCGPMenuImages::IdArowDownLarge, rectArrow);
	}
}

#ifndef BCGP_EXCLUDE_GRID_CTRL

void CBCGPVisualManager::OnDrawGridSortArrow (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC,
											  CRect rectArrow, BOOL bIsUp)
{
	DoDrawHeaderSortArrow (pDC, rectArrow, bIsUp, TRUE/*!pCtrl->IsControlBarColors ()*/);
}
//********************************************************************************
COLORREF CBCGPVisualManager::OnFillGridGroupByBoxBackground (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	CBrush br (globalData.clrBtnShadow);
	pDC->FillRect (rect, &br);

	return globalData.clrBarText;
}
//********************************************************************************
COLORREF CBCGPVisualManager::OnFillGridGroupByBoxTitleBackground (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBtnFace);
	return globalData.clrBtnShadow;
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetGridGroupByBoxLineColor () const
{
	return globalData.clrBarText;
}
//********************************************************************************
void CBCGPVisualManager::OnDrawGridGroupByBoxItemBorder (CBCGPGridCtrl* /*pCtrl*/,
														 CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);
	pDC->FillRect (rect, &globalData.brBtnFace);

	pDC->Draw3dRect (rect, globalData.clrBarWindow, globalData.clrBtnText);
	rect.DeflateRect (0, 0, 1, 1);
	pDC->Draw3dRect (rect, globalData.clrBarWindow, globalData.clrBtnShadow);
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetGridLeftOffsetColor (CBCGPGridCtrl* pCtrl)
{
	ASSERT_VALID (pCtrl);

	COLORREF clrGray = (COLORREF)-1;

	if (globalData.m_nBitsPerPixel <= 8)
	{
		clrGray = pCtrl->IsControlBarColors () ?
			globalData.clrBarShadow : globalData.clrBtnShadow;
	}
	else
	{
		clrGray = CBCGPDrawManager::PixelAlpha (
			pCtrl->IsControlBarColors () ? globalData.clrBarFace : globalData.clrBtnFace, 94);
	}

	return clrGray;
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetGridItemSortedColor (CBCGPGridCtrl* pCtrl)
{
	ASSERT_VALID (pCtrl);

	COLORREF clrSortedColumn = (COLORREF)-1;
	if (globalData.m_nBitsPerPixel <= 8)
	{
		clrSortedColumn = pCtrl->GetBkColor ();
	}
	else
	{
		clrSortedColumn = CBCGPDrawManager::PixelAlpha (
				pCtrl->GetBkColor (), .97, .97, .97);
	}

	return clrSortedColumn;
}
//********************************************************************************
void CBCGPVisualManager::OnFillGridGroupBackground (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, 
												   CRect rectFill)
{
	ASSERT_VALID (pDC);
	pDC->FillRect (rectFill, &globalData.brWindow);
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetGridGroupTextColor(const CBCGPGridCtrl* pCtrl, BOOL bSelected)
{
	ASSERT_VALID(pCtrl);
	
	if (!pCtrl->IsHighlightGroups() && bSelected)
	{
		return !pCtrl->IsFocused() ? globalData.clrHotText : globalData.clrTextHilite;
	}
	else // not selected
	{
		return (pCtrl->IsHighlightGroups() ? 
			(pCtrl->IsControlBarColors() ? globalData.clrBarShadow : globalData.clrBtnShadow) :
			globalData.clrWindowText);
	}
}
//********************************************************************************
void CBCGPVisualManager::OnDrawGridGroupUnderline (CBCGPGridCtrl* pCtrl, CDC* pDC, 
												   CRect rectFill)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pCtrl);

	COLORREF clrOld = pDC->GetBkColor ();
	pDC->FillSolidRect (rectFill, 
		pCtrl->IsControlBarColors () ? globalData.clrBarShadow : globalData.clrBtnShadow);
	pDC->SetBkColor (clrOld);
}
//********************************************************************************
COLORREF CBCGPVisualManager::OnFillGridRowBackground (CBCGPGridCtrl* pCtrl, 
												  CDC* pDC, CRect rectFill, BOOL bSelected)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pCtrl);

	// Fill area:
	if (!pCtrl->IsFocused ())
	{
		pDC->FillRect (rectFill, 
			pCtrl->IsControlBarColors () ? &globalData.brBarFace : &globalData.brBtnFace);
	}
	else
	{
		pDC->FillRect (rectFill, &globalData.brHilite);
	}

	// Return text color:
	if (!pCtrl->IsHighlightGroups () && bSelected)
	{
		return (!pCtrl->IsFocused ()) ? globalData.clrHotText : globalData.clrTextHilite;
	}

	return pCtrl->IsHighlightGroups () ? 
		(pCtrl->IsControlBarColors () ? globalData.clrBarShadow : globalData.clrBtnShadow) :
		globalData.clrWindowText;
}
//********************************************************************************
COLORREF CBCGPVisualManager::OnFillGridItem (CBCGPGridCtrl* pCtrl, 
											CDC* pDC, CRect rectFill,
											BOOL bSelected, BOOL bActiveItem, BOOL bSortedColumn)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pCtrl);

	// Fill area:
	if (bSelected && !bActiveItem)
	{
		if (!pCtrl->IsFocused() && globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode())
		{
			if (pCtrl->IsControlBarColors ())
			{
				pDC->FillRect (rectFill, &globalData.brBarFace);
				return globalData.clrBarText;
			}
			else
			{
				pDC->FillRect (rectFill, &globalData.brBtnFace);
				return globalData.clrBtnText;
			}
		}
		else
		{
			pDC->FillRect (rectFill, &globalData.brHilite);
			return globalData.clrTextHilite;
		}
	}
	else
	{
		if (bActiveItem)
		{
			pDC->FillRect (rectFill, &globalData.brWindow);
		}
		else if (bSortedColumn)
		{
			CBrush br (pCtrl->GetSortedColor ());
			pDC->FillRect (rectFill, &br);
		}
		else
		{
			// no painting
		}
	}

	return (COLORREF)-1;
}
//********************************************************************************
void CBCGPVisualManager::OnDrawGridHeaderMenuButton (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, CRect rect, 
		BOOL bHighlighted, BOOL /*bPressed*/, BOOL /*bDisabled*/)
{
	ASSERT_VALID (pDC);

	rect.DeflateRect (1, 1);

	if (bHighlighted)
	{
		pDC->Draw3dRect (rect, globalData.clrBtnHilite, globalData.clrBtnDkShadow);
	}
}
//********************************************************************************
void CBCGPVisualManager::OnDrawGridSelectionBorder (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->Draw3dRect (rect, globalData.clrBtnDkShadow, globalData.clrBtnDkShadow);
	rect.DeflateRect (1, 1);
	pDC->Draw3dRect (rect, globalData.clrBtnDkShadow, globalData.clrBtnDkShadow);
}
//********************************************************************************
COLORREF CBCGPVisualManager::OnFillGridCaptionRow (CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rectFill)
{
	ASSERT_VALID(pDC);
	
	pDC->FillRect(rectFill, &GetGridCaptionBrush(pCtrl));

	COLORREF clrBorderVert = !globalData.IsHighContastMode() ? globalData.clrBtnFace : globalData.clrBarShadow;
	COLORREF clrBorderHorz = !globalData.IsHighContastMode() ? globalData.clrBtnFace : globalData.clrBarShadow;
	
	if (pCtrl != NULL)
	{
		ASSERT_VALID(pCtrl);

		if (pCtrl->GetColorTheme().m_clrVertLine != (COLORREF)-1)
		{
			clrBorderVert = pCtrl->GetColorTheme().m_clrVertLine;
		}

		if (pCtrl->GetColorTheme().m_clrHorzLine != (COLORREF)-1)
		{
			clrBorderHorz = pCtrl->GetColorTheme().m_clrHorzLine;
		}
	}

	{
		CBCGPPenSelector penVert(*pDC, clrBorderVert);
		
		pDC->MoveTo(rectFill.right - 1, rectFill.top);
		pDC->LineTo(rectFill.right - 1, rectFill.bottom);
	}
	
	{
		CBCGPPenSelector penHorz(*pDC, clrBorderHorz);
		
		pDC->MoveTo(rectFill.left, rectFill.top);
		pDC->LineTo(rectFill.right - 1, rectFill.top);

		pDC->MoveTo(rectFill.left, rectFill.bottom - 1);
		pDC->LineTo(rectFill.right - 1, rectFill.bottom - 1);
	}

	return globalData.clrBarText;
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetGridTreeLineColor (CBCGPGridCtrl* /*pCtrl*/)
{
	return !globalData.IsHighContastMode() ? globalData.clrBarShadow : globalData.clrWindowText;
}
//********************************************************************************
BOOL CBCGPVisualManager::OnSetGridColorTheme (CBCGPGridCtrl* /*pCtrl*/, BCGP_GRID_COLOR_DATA& /*theme*/)
{
	return TRUE;
}
//********************************************************************************
CBrush& CBCGPVisualManager::GetGridCaptionBrush(CBCGPGridCtrl* /*pCtrl*/)
{
	return globalData.brBarFace;
}
//********************************************************************************
void CBCGPVisualManager::OnDrawGridDataBar (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, CRect rect)
{
	COLORREF clrBorder = (COLORREF)-1;
	COLORREF clrFill = (COLORREF)-1;

	if (CBCGPDrawManager::IsDarkColor(globalData.clrWindow) || IsDarkTheme())
	{
		clrFill = RGB (20, 75, 175);
		clrBorder = RGB (76, 131, 233);
	}
	else
	{
		clrFill = RGB (205, 221, 249);
		clrBorder = RGB (156, 187, 243);
	}

	CBrush brFill(clrFill);
	pDC->FillRect (rect, &brFill);

	pDC->Draw3dRect (rect, clrBorder, clrBorder);
}
//********************************************************************************
void CBCGPVisualManager::GetGridColorScaleBaseColors(CBCGPGridCtrl* /*pCtrl*/, COLORREF& clrLow, COLORREF& clrMid, COLORREF& clrHigh)
{
	clrLow = RGB(250, 149, 150);
	clrMid = RGB(255, 235, 132);
	clrHigh = RGB(99, 190, 123);
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetOutOffFilterTextColor(CBCGPGridCtrl* /*pCtrl*/)
{
	return globalData.clrBarDkShadow;
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetGridDragHeaderTextColor(CBCGPGridCtrl* /*pCtrl*/)
{
	return globalData.clrBarText;
}
//********************************************************************************
COLORREF CBCGPVisualManager::OnDrawGridColumnChooserItem(CBCGPGridCtrl* pCtrl, CDC* pDC, CRect rectItem, BOOL bIsSelected)
{
	ASSERT_VALID (pDC);

	OnDrawGridGroupByBoxItemBorder(pCtrl, pDC, rectItem);

	if (bIsSelected)
	{
		rectItem.DeflateRect(1, 1);
		pDC->FillRect (rectItem, &globalData.brHilite);

		return globalData.clrTextHilite;
	}

	return (COLORREF)-1;
}
//********************************************************************************
COLORREF CBCGPVisualManager::OnFillReportCtrlRowBackground (CBCGPGridCtrl* pCtrl, 
												  CDC* pDC, CRect rectFill,
												  BOOL bSelected, BOOL bGroup)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pCtrl);

	// Fill area:
	COLORREF clrText = (COLORREF)-1;

	if (bSelected)
	{
		if (!pCtrl->IsFocused ())
		{
			pDC->FillRect (rectFill, 
				pCtrl->IsControlBarColors () ? &globalData.brBarFace : &globalData.brBtnFace);

			clrText = m_clrReportGroupText;
		}
		else
		{
			pDC->FillRect (rectFill, &globalData.brHilite);
			clrText = globalData.clrTextHilite;
		}
	}
	else
	{
		if (bGroup)
		{
			// no painting
			clrText = m_clrReportGroupText;
		}
	}

	// Return text color:
	return clrText;
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetReportCtrlGroupBackgoundColor ()
{
	return globalData.clrBtnLight;
}
//********************************************************************************
void CBCGPVisualManager::OnFillGridBackground (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brWindow);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawGridExpandingBox (CDC* pDC, CRect rect, BOOL bIsOpened, COLORREF colorBox)
{
	OnDrawExpandingBox (pDC, rect, bIsOpened, colorBox);
}
//********************************************************************************
void CBCGPVisualManager::OnFillGridHeaderBackground (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBtnFace);
}
//********************************************************************************
BOOL CBCGPVisualManager::OnDrawGridHeaderItemBorder (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, CRect rect, BOOL /*bPressed*/)
{
	ASSERT_VALID (pDC);

	pDC->Draw3dRect (rect, globalData.clrBtnHilite, globalData.clrBtnDkShadow);
	return TRUE;
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetGridHeaderItemTextColor (CBCGPGridCtrl* /*pCtrl*/, BOOL /*bSelected*/, BOOL /*bGroupByBox*/)
{
	return (COLORREF)-1;
}
//********************************************************************************
void CBCGPVisualManager::OnFillGridRowHeaderBackground (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBtnFace);
}
//********************************************************************************
BOOL CBCGPVisualManager::OnDrawGridRowHeaderItemBorder (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, CRect rect, BOOL /*bPressed*/)
{
	ASSERT_VALID (pDC);

	pDC->Draw3dRect (rect, globalData.clrBtnHilite, globalData.clrBtnDkShadow);
	return TRUE;
}
//********************************************************************************
void CBCGPVisualManager::OnFillGridSelectAllAreaBackground (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, CRect rect, BOOL /*bPressed*/)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBtnFace);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawGridSelectAllMarker(CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, CRect rect, int /*nPadding*/, BOOL /*bPressed*/)
{
	ASSERT_VALID(pDC);
	
	POINT ptRgn [] =
	{
		{rect.right, rect.top},
		{rect.right, rect.bottom},
		{rect.left, rect.bottom}
	};
	
	CRgn rgn;
	rgn.CreatePolygonRgn (ptRgn, 3, WINDING);
	
	pDC->SelectClipRgn (&rgn, RGN_COPY);
	
	pDC->FillSolidRect(rect, globalData.clrBtnShadow);
	
	pDC->SelectClipRgn (NULL, RGN_COPY);
}
//********************************************************************************
BOOL CBCGPVisualManager::OnDrawGridSelectAllAreaBorder (CBCGPGridCtrl* /*pCtrl*/, CDC* pDC, CRect rect, BOOL /*bPressed*/)
{
	ASSERT_VALID (pDC);

	pDC->Draw3dRect (rect, globalData.clrBtnHilite, globalData.clrBtnDkShadow);
	return TRUE;
}

#endif // BCGP_EXCLUDE_GRID_CTRL

//********************************************************************************

#if !defined (BCGP_EXCLUDE_GRID_CTRL) && !defined (BCGP_EXCLUDE_GANTT)

//********************************************************************************
void CBCGPVisualManager::GetGanttColors (const CBCGPGanttChart* /*pChart*/, BCGP_GANTT_CHART_COLORS& colors, COLORREF clrBack) const
{
	if (clrBack == CLR_DEFAULT)
	{
		clrBack = globalData.clrWindow;
	}	

	CBCGPGanttChart::PrepareColorScheme (clrBack, colors);

    colors.clrBackground      = globalData.clrWindow;
	colors.clrShadows         = m_clrMenuShadowBase;

	colors.clrRowBackground   = colors.clrBackground;
    colors.clrGridLine0       = globalData.clrBarShadow;
    colors.clrGridLine1       = globalData.clrBarLight;
    colors.clrSelection       = globalData.clrHilite;
	colors.clrSelectionBorder = globalData.clrHilite;

	colors.clrBarFill         = RGB (0, 0, 255);
	colors.clrBarComplete     = RGB (0, 255, 0);
}
//********************************************************************************
void CBCGPVisualManager::DrawGanttChartBackground (const CBCGPGanttChart*, CDC& dc, const CRect& rectChart, COLORREF clrFill)
{
	dc.FillSolidRect (rectChart, clrFill);
}
//********************************************************************************
void CBCGPVisualManager::DrawGanttItemBackgroundCell (const CBCGPGanttChart*, CDC& dc, const CRect& /*rectItem*/, const CRect& rectClip, const BCGP_GANTT_CHART_COLORS& colors, BOOL bDayOff)
{
	dc.FillSolidRect (rectClip, (bDayOff) ? colors.clrRowDayOff : colors.clrRowBackground);
}
//********************************************************************************
void CBCGPVisualManager::DrawGanttHeaderCell (const CBCGPGanttChart*, CDC& dc, const BCGP_GANTT_CHART_HEADER_CELL_INFO& cellInfo, BOOL /*bHilite*/)
{
	dc.FillRect (cellInfo.rectCell, &globalData.brBarFace);
	dc.Draw3dRect (cellInfo.rectCell, globalData.clrBtnHilite, globalData.clrBtnDkShadow);
}
//********************************************************************************
void CBCGPVisualManager::DrawGanttHeaderText (const CBCGPGanttChart* pGanttChart, CDC& dc, const BCGP_GANTT_CHART_HEADER_CELL_INFO& cellInfo, const CString& sCellText, BOOL bHilite) const
{
	CRect rcCell = cellInfo.rectCell;
	rcCell.DeflateRect (1, 2, 2, 2); // text padding

	CRect rcVisible;
	rcVisible.IntersectRect (&cellInfo.rectClip, rcCell);

	if (rcVisible.IsRectEmpty ())
	{
		return;
	}

	CFont* pNewFont = NULL;
	if (pGanttChart != NULL)
	{
		pNewFont = pGanttChart->GetFont();
	}

	CFont* pOldFont = dc.SelectObject(pNewFont == NULL ? &globalData.fontRegular : pNewFont);

	dc.SetBkMode (TRANSPARENT);
	dc.SetTextColor (GetGanttHeaderTextColor (bHilite));

	ASSERT(cellInfo.pHeaderInfo != NULL);

	DWORD dwFlags = DT_VCENTER | DT_NOPREFIX | cellInfo.pHeaderInfo->dwAlignment;
	dc.DrawText (sCellText, rcCell, dwFlags);

	dc.SelectObject (pOldFont);
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetGanttHeaderTextColor (BOOL bHilite) const
{
	return bHilite ? globalData.clrHotText : globalData.clrBarText;
}
//********************************************************************************
void CBCGPVisualManager::FillGanttBar (const CBCGPGanttItem* /*pItem*/, CDC& dc, const CRect& rectFill, COLORREF color, double /*dGlowLine*/)
{
	dc.FillSolidRect (rectFill, color);
}
//********************************************************************************

#endif // !defined (BCGP_EXCLUDE_GRID_CTRL) && !defined (BCGP_EXCLUDE_GANTT)

//********************************************************************************

#ifndef BCGP_EXCLUDE_PROP_LIST

COLORREF CBCGPVisualManager::GetPropListGroupColor (CBCGPPropList* pPropList)
{
	ASSERT_VALID (pPropList);

	return pPropList->DrawControlBarColors () ?
		globalData.clrBarFace : globalData.clrBtnFace;
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetPropListGroupTextColor (CBCGPPropList* pPropList)
{
	ASSERT_VALID (pPropList);

	return pPropList->DrawControlBarColors () ?
		globalData.clrBarDkShadow : globalData.clrBtnDkShadow;
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetPropListCommandTextColor (CBCGPPropList* /*pPropList*/)
{	
	return globalData.clrHotText;
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetPropListDisabledTextColor(CBCGPPropList* /*pPropList*/)
{
	return globalData.clrGrayedText;
}

#endif // BCGP_EXCLUDE_PROP_LIST

COLORREF CBCGPVisualManager::GetMenuItemTextColor (
	CBCGPToolbarMenuButton* /*pButton*/, BOOL bHighlighted, BOOL bDisabled)
{
	if (bHighlighted)
	{
		return bDisabled ? globalData.clrBtnFace : globalData.clrTextHilite;
	}

	return bDisabled ? globalData.clrGrayedText : globalData.clrWindowText;
}
//*****************************************************************************
COLORREF CBCGPVisualManager::GetPopupMenuBackgroundColor() const
{
	return globalData.clrBarFace;
}
//*****************************************************************************
int CBCGPVisualManager::GetMenuDownArrowState (CBCGPToolbarMenuButton* /*pButton*/, BOOL /*bHightlight*/, BOOL /*bPressed*/, BOOL bDisabled)
{
	if (bDisabled)
	{
		return (int)CBCGPMenuImages::ImageGray;
	}

	return (int) CBCGPMenuImages::ImageBlack;
}
//*****************************************************************************
COLORREF CBCGPVisualManager::GetStatusBarPaneTextColor (CBCGPStatusBar* /*pStatusBar*/, 
									CBCGStatusBarPaneInfo* pPane)
{
	ASSERT (pPane != NULL);

	return (pPane->nStyle & SBPS_DISABLED) ? globalData.clrGrayedText : 
			pPane->clrText == (COLORREF)-1 ? globalData.clrBtnText : pPane->clrText;
}

#ifndef BCGP_EXCLUDE_RIBBON

BOOL CBCGPVisualManager::IsRibbonCaptionTextCentered(CBCGPRibbonBar* /*pBar*/)
{
	return TRUE;
}
//*****************************************************************************
COLORREF CBCGPVisualManager::GetRibbonCaptionTextColor(CBCGPRibbonBar* /*pBar*/, BOOL bActive)
{
	return GetNcCaptionTextColor(bActive, TRUE);
}
//*****************************************************************************
void CBCGPVisualManager::OnDrawRibbonCaption (CDC* pDC, CBCGPRibbonBar* pBar,
											  CRect rect, CRect rectText)
{
	ASSERT_VALID (pBar);

	CWnd* pWnd = pBar->GetParent ();
	ASSERT_VALID (pWnd);

	const BOOL bGlass	  = pBar->IsTransparentCaption ();
	const DWORD dwStyleEx = pWnd->GetExStyle ();
	const BOOL bIsRTL     = (dwStyleEx & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL;
	const BOOL bActive    = IsWindowActive (pWnd);

	BOOL bTextCenter = IsRibbonCaptionTextCentered(pBar);

	ASSERT_VALID (pDC);

	CRect rectQAT = pBar->GetQuickAccessToolbarLocation ();
	BOOL bHide  = (pBar->GetHideFlags () & BCGPRIBBONBAR_HIDE_ALL) != 0;
	BOOL bExtra = !bHide && pBar->IsQuickAccessToolbarOnTop () &&
				  rectQAT.left < rectQAT.right && !pBar->IsQATEmpty();

	if ((bHide && !bExtra) || pBar->IsDrawSystemIcon())
	{
		BOOL bDestroyIcon = FALSE;
		HICON hIcon = globalUtils.GetWndIcon (pWnd, &bDestroyIcon, FALSE);

		if (hIcon != NULL)
		{
			CSize szIcon (::GetSystemMetrics (SM_CXSMICON), ::GetSystemMetrics (SM_CYSMICON));

			long x = rect.left + 2 + pBar->GetControlsSpacing().cx / 2;
			long y = rect.top  + max (0, (rect.Height () - szIcon.cy) / 2);

			y -= globalUtils.ScaleByDPI(GetRibbonQATButtonHorzMargin().cy) / 2;

			if (bGlass)
			{
				globalData.DrawIconOnGlass (m_hThemeWindow, pDC, hIcon, CRect (x, y, x + szIcon.cx, y + szIcon.cy));
			}
			else
			{
				::DrawIconEx (pDC->GetSafeHdc (), x, y, hIcon, szIcon.cx, szIcon.cy,
					0, NULL, DI_NORMAL);
			}

			if (rectText.left < (x + szIcon.cx + 4))
			{
				rectText.left = x + szIcon.cx + 4;
			}

			if (bDestroyIcon)
			{
				::DestroyIcon (hIcon);
			}
		}

		bTextCenter = TRUE;
	}

	CFont* pOldFont = pDC->SelectObject (&GetNcCaptionTextFont());
	ASSERT (pOldFont != NULL);
	
	int nOldMode = pDC->SetBkMode(TRANSPARENT);

	CString strCaption;
	pWnd->GetWindowText (strCaption);

	DWORD dwTextStyle = DT_END_ELLIPSIS | DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX |
		(bIsRTL ? DT_RTLREADING | DT_RIGHT : 0);

	COLORREF clrText = GetRibbonCaptionTextColor(pBar, bActive);

	int widthFull = rectText.Width ();
	int width = pDC->GetTextExtent (strCaption).cx;

	if (bTextCenter && width < widthFull)
	{
		rectText.left += (widthFull - width) / 2;
	}

	rectText.right = min (rectText.left + width, rectText.right);

	if (rectText.right > rectText.left)
	{
		if (bGlass)
		{
			int nGlassGlowSize = 10;
			COLORREF clrTextGlass = globalData.GetDWMTextColor(bActive, pWnd->IsZoomed(), &nGlassGlowSize);

			DrawTextOnGlass(pDC, strCaption, rectText, dwTextStyle, nGlassGlowSize, clrTextGlass);
		}
		else
		{
			COLORREF clrOldText = pDC->SetTextColor (clrText);
			pDC->DrawText (strCaption, rectText, dwTextStyle);
			pDC->SetTextColor (clrOldText);
		}
	}

	pDC->SetBkMode    (nOldMode);
	pDC->SelectObject (pOldFont);
}
//*****************************************************************************
void CBCGPVisualManager::OnDrawRibbonCaptionButton (
					CDC* pDC, CBCGPRibbonCaptionButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	COLORREF clrText = OnFillRibbonButton (pDC, pButton);

	CBCGPMenuImages::IMAGES_IDS imageID;

	switch (pButton->GetID ())
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

	default:
		return;
	}

	const BOOL bActive = pButton->IsActive() || globalData.IsHighContastMode();

	CBCGPMenuImages::IMAGE_STATE state = pButton->IsHighlighted() ? CBCGPMenuImages::ImageBlack2 : (bActive ? CBCGPMenuImages::ImageBlack : CBCGPMenuImages::ImageLtGray);
	
	if (pButton->IsDisabled ())
	{
		state = CBCGPMenuImages::ImageGray;
	}
	else if (clrText != (COLORREF)-1)
	{
		if (GetRValue (clrText) <= 128 ||
			GetGValue (clrText) <= 128 ||
			GetBValue (clrText) <= 128)
		{
			state = pButton->IsHighlighted() ? CBCGPMenuImages::ImageBlack2 : (bActive ? CBCGPMenuImages::ImageBlack : CBCGPMenuImages::ImageLtGray);
		}
		else
		{
			state = bActive ? CBCGPMenuImages::ImageWhite : CBCGPMenuImages::ImageLtGray;
		}
	}

	CBCGPMenuImages::Draw (pDC, imageID, pButton->GetRect (), state);

	OnDrawRibbonButtonBorder (pDC, pButton);
}
//*******************************************************************************
COLORREF CBCGPVisualManager::OnDrawRibbonButtonsGroup (
					CDC* /*pDC*/, CBCGPRibbonButtonsGroup* /*pGroup*/,
					CRect /*rectGroup*/)
{
	return (COLORREF)-1;
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawDefaultRibbonImage (
					CDC* pDC, CRect rectImage,
					BOOL bIsDisabled,
					BOOL /*bIsPressed*/,
					BOOL /*bIsHighlighted*/)
{
	ASSERT_VALID (pDC);

	CRect rectBullet (rectImage.CenterPoint (), CSize (1, 1));
	
	int nRadius = globalUtils.ScaleByDPI(5);
	rectBullet.InflateRect(nRadius, nRadius);

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () ||
		globalData.bIsWindows9x)
	{
		CBrush br (bIsDisabled ? globalData.clrGrayedText : RGB (0, 127, 0));

		CBrush* pOldBrush = (CBrush*) pDC->SelectObject (&br);
		CPen* pOldPen = (CPen*) pDC->SelectStockObject (NULL_PEN);

		pDC->Ellipse (rectBullet);

		pDC->SelectObject (pOldBrush);
		pDC->SelectObject (pOldPen);
	}
	else
	{
		CBCGPDrawManager dm (*pDC);

		dm.DrawEllipse (rectBullet,
			bIsDisabled ? globalData.clrGrayedText : RGB (160, 208, 128),
			bIsDisabled ? globalData.clrBtnShadow : RGB (71, 117, 44));
	}
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawRibbonMainButton (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	const BOOL bIsHighlighted = pButton->IsHighlighted () || pButton->IsFocused ();
	const BOOL bIsPressed = pButton->IsPressed () || pButton->IsDroppedDown ();

	CRect rect = pButton->GetRect ();

	COLORREF clrFill = (bIsPressed && !globalData.IsHighContastMode ()) ? globalData.clrBarLight : globalData.clrBarFace;
	COLORREF clrLine = bIsHighlighted ? globalData.clrBarDkShadow : globalData.clrBarShadow;

	const BOOL bIsScenic = pButton->GetParentRibbonBar () != NULL && pButton->GetParentRibbonBar ()->IsScenicLook ();

	if (!bIsScenic)
	{
		rect.DeflateRect (2, 2);
	}

	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode () || globalData.bIsWindows9x)
	{
		CBrush br (clrFill);
		CBrush* pOldBrush = (CBrush*) pDC->SelectObject (&br);

		CPen pen (PS_SOLID, 1, clrLine);
		CPen* pOldPen = (CPen*) pDC->SelectObject (&pen);

		if (bIsScenic)
		{
			pDC->Rectangle (rect);
		}
		else
		{
			pDC->Ellipse (rect);
		}

		pDC->SelectObject (pOldBrush);
		pDC->SelectObject (pOldPen);
	}
	else
	{
		if (m_clrMainButton != (COLORREF)-1)
		{
			clrFill = m_clrMainButton;

			if (bIsHighlighted || bIsPressed)
			{
				clrFill = CBCGPDrawManager::ColorMakeLighter(clrFill);
			}
			else if (bIsPressed)
			{
				clrFill = CBCGPDrawManager::ColorMakeDarker(clrFill);
			}
		}

		CBCGPDrawManager dm (*pDC);

		if (bIsScenic)
		{
			dm.DrawRect (rect, clrFill, clrLine);
		}
		else
		{
			dm.DrawEllipse (rect, clrFill, clrLine);
		}
	}
}
//*****************************************************************************
void CBCGPVisualManager::OnDrawRibbonMainButtonLabel(CDC* pDC, CBCGPRibbonButton* pButton, CRect rectText, const CString strText)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	CString strLabel = strText;
	if (IsTopLevelMenuItemUpperCase())
	{
		strLabel.MakeUpper();
	}
	
	UINT nDTFlags = DT_SINGLELINE | DT_CENTER | DT_VCENTER;

	CRect rectShadow = rectText;
	rectShadow.OffsetRect(1, 1);

	if (globalData.IsHighContastMode())
	{
		pButton->DoDrawText(pDC, strLabel, rectText, nDTFlags, pButton->IsDisabled() ? globalData.clrGrayedText : globalData.clrWindowText);
	}
	else
	{
		pButton->DoDrawText(pDC, strLabel, rectShadow, nDTFlags, pButton->IsDisabled() ? RGB(128, 128, 128) : RGB(100, 100, 100));
		pButton->DoDrawText(pDC, strLabel, rectText, nDTFlags, RGB(255, 255, 255));
	}
}
//*****************************************************************************
COLORREF CBCGPVisualManager::OnDrawRibbonTabsFrame (
					CDC* pDC, 
					CBCGPRibbonBar* /*pWndRibbonBar*/, 
					CRect rectTab)
{
	ASSERT_VALID (pDC);

	CPen pen (PS_SOLID, 1, globalData.clrBarShadow);
	CPen* pOldPen = pDC->SelectObject (&pen);
	ASSERT (pOldPen != NULL);

	pDC->MoveTo (rectTab.left, rectTab.top);
	pDC->LineTo (rectTab.right, rectTab.top);

	pDC->SelectObject (pOldPen);

	return globalData.m_bIsBlackHighContrast ? globalData.clrBtnText : (COLORREF)-1;
}
//*****************************************************************************
void CBCGPVisualManager::OnDrawRibbonCategory (
		CDC* pDC, 
		CBCGPRibbonCategory* pCategory,
		CRect rectCategory)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pCategory);

	const int nShadowSize = 2;

	rectCategory.right -= nShadowSize;
	rectCategory.bottom -= nShadowSize;

	pDC->FillRect (rectCategory, &globalData.brBarFace);

	CRect rectActiveTab = pCategory->GetTabRect ();

	COLORREF clrCategory = m_bFillRibbonContextTab ? (COLORREF)-1 : RibbonCategoryColorToRGB(pCategory->GetTabColor());
	CPen penActive(PS_SOLID, 1, clrCategory);

	CPen pen (PS_SOLID, 1, globalData.clrBarShadow);

	CPen* pOldPen = pDC->SelectObject(clrCategory != (COLORREF)-1 ? &penActive : &pen);
	ASSERT (pOldPen != NULL);

	pDC->MoveTo (rectCategory.left, rectCategory.top);
	pDC->LineTo (rectActiveTab.left + 1, rectCategory.top);

	pDC->MoveTo (rectActiveTab.right - 2, rectCategory.top);
	pDC->LineTo (rectCategory.right, rectCategory.top);

	if (clrCategory != (COLORREF)-1)
	{
		pDC->SelectObject(&pen);
	}

	pDC->LineTo (rectCategory.right, rectCategory.bottom);
	pDC->LineTo (rectCategory.left, rectCategory.bottom);
	pDC->LineTo (rectCategory.left, rectCategory.top);

	pDC->SelectObject (pOldPen);

	if (!m_bIsRectangulareRibbonTab)
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawShadow (rectCategory, nShadowSize, 100, 75, NULL, NULL,
			m_clrMenuShadowBase);
	}
}
//*****************************************************************************
void CBCGPVisualManager::OnDrawRibbonDropDownPanel(CDC* pDC, CBCGPRibbonPanel* pPanel, CRect rectFill)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pPanel);

	CBCGPRibbonCategory* pCategory = pPanel->GetParentCategory();
	ASSERT_VALID (pCategory);
				
	CBCGPRibbonPanelMenuBar* pMenuBarSaved = pCategory->m_pParentMenuBar;
	pCategory->m_pParentMenuBar = pPanel->GetParentMenuBar();
				
	OnDrawRibbonCategory(pDC, pCategory, rectFill);
				
	pCategory->m_pParentMenuBar = pMenuBarSaved;
}
//*****************************************************************************
void CBCGPVisualManager::OnDrawRibbonCategoryScroll (
					CDC* pDC, 
					CBCGPRibbonCategoryScroll* pScroll)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pScroll);

	CRect rect = pScroll->GetRect ();
	rect.bottom--;

	pDC->FillRect (rect, &globalData.brBarFace);
	if (pScroll->IsHighlighted ())
	{
		CBCGPDrawManager dm (*pDC);
		dm.HighlightRect (rect);
	}

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

	pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarShadow);
}
//*****************************************************************************
COLORREF CBCGPVisualManager::OnDrawRibbonCategoryTab (
					CDC* pDC, 
					CBCGPRibbonTab* pTab,
					BOOL bIsActive)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pTab);

	CBCGPRibbonCategory* pCategory = pTab->GetParentCategory ();
	ASSERT_VALID (pCategory);
	CBCGPRibbonBar* pBar = pCategory->GetParentRibbonBar ();
	ASSERT_VALID (pBar);

	COLORREF clrCategory = RibbonCategoryColorToRGB(pCategory->GetTabColor());

	bIsActive = bIsActive && 
		((pBar->GetHideFlags () & BCGPRIBBONBAR_HIDE_ELEMENTS) == 0 || pTab->GetDroppedDown () != NULL);

	const BOOL bIsFocused	= pTab->IsFocused () && (pBar->GetHideFlags () & BCGPRIBBONBAR_HIDE_ELEMENTS);
	const BOOL bIsHighlighted = (pTab->IsHighlighted () || bIsFocused) && !pTab->IsDroppedDown ();

	CRect rectTab = pTab->GetRect();

	if (globalData.IsHighContastMode())
	{
		COLORREF clrText = globalData.clrBarText;

		if (bIsActive || bIsHighlighted || bIsFocused)
		{
			pDC->FillRect(rectTab, &globalData.brHilite);
			clrText = globalData.clrTextHilite;
		}
		
		if (clrCategory != (COLORREF)-1)
		{
			pDC->Draw3dRect(rectTab, clrCategory, clrCategory);
			rectTab.DeflateRect(1, 1);
			pDC->Draw3dRect(rectTab, clrCategory, clrCategory);
		}

		return clrText;
	}

	rectTab.top += 3;

	const int nTrancateRatio = pCategory->GetParentRibbonBar ()->GetTabTrancateRatio ();

	if (nTrancateRatio > 0)
	{
		const int nPercent = max (10, 100 - nTrancateRatio / 2);

		COLORREF color = CBCGPDrawManager::PixelAlpha (
			globalData.clrBarFace, nPercent);

		CPen pen (PS_SOLID, 1, color);
		CPen* pOldPen = pDC->SelectObject (&pen);

		pDC->MoveTo (rectTab.right - 1, rectTab.top);
		pDC->LineTo (rectTab.right - 1, rectTab.bottom - 2);

		pDC->SelectObject (pOldPen);
	}

	if (!bIsActive && !bIsHighlighted)
	{
		return globalData.clrBarText;
	}

	CPen pen(PS_SOLID, 1, globalData.clrBarShadow);
	CPen* pOldPen = pDC->SelectObject(&pen);
	ASSERT(pOldPen != NULL);

	rectTab.right -= 2;

	CRgn rgnClip;

	const int nPointsNum = 8;
	POINT pts [nPointsNum];

	if (!m_bIsRectangulareRibbonTab)
	{
		pts [0] = CPoint (rectTab.left, rectTab.bottom);
		pts [1] = CPoint (rectTab.left + 1, rectTab.bottom - 1);
		pts [2] = CPoint (rectTab.left + 1, rectTab.top + 2);
		pts [3] = CPoint (rectTab.left + 3, rectTab.top);
		pts [4] = CPoint (rectTab.right - 3, rectTab.top);
		pts [5] = CPoint (rectTab.right - 1, rectTab.top + 2);
		pts [6] = CPoint (rectTab.right - 1, rectTab.bottom - 1);
		pts [7] = CPoint (rectTab.right, rectTab.bottom);

		rgnClip.CreatePolygonRgn (pts, nPointsNum, WINDING);
	}

	CPen penActive(PS_SOLID, 2, clrCategory);

	if (bIsActive)
	{
		if (!m_bIsRectangulareRibbonTab)
		{
			pDC->SelectClipRgn (&rgnClip);
		}

		COLORREF clrFill = pTab->IsSelected () ? globalData.clrBarHilite : 
			m_bFillRibbonContextTab ? clrCategory : (COLORREF)-1;
		
		if (clrFill != (COLORREF)-1)
		{
			CBrush br (clrFill);
			pDC->FillRect (rectTab, &br);
		}
		else
		{
			pDC->FillRect (rectTab, 
				bIsHighlighted ? 
				&globalData.brWindow : &globalData.brBarFace);
		}

		if (!m_bIsRectangulareRibbonTab)
		{
			pDC->SelectClipRgn (NULL);
		}

		if (!m_bFillRibbonContextTab && clrCategory != (COLORREF)-1)
		{
			pDC->SelectObject(&penActive);
		}
	}

	if (!m_bIsRectangulareRibbonTab)
	{
		pDC->Polyline (pts, nPointsNum);
	}
	else
	{
		pDC->MoveTo(rectTab.left, rectTab.bottom);
		pDC->LineTo(rectTab.left, rectTab.top - 1);
		pDC->LineTo(rectTab.right, rectTab.top - 1);
		pDC->LineTo(rectTab.right, rectTab.bottom);
	}

	pDC->SelectObject (pOldPen);

	return globalData.clrBarText;
}
//****************************************************************************
COLORREF CBCGPVisualManager::OnFillRibbonBackstageLeftPane(CDC* pDC, CRect rectPanel)
{
	ASSERT_VALID(pDC);
	pDC->FillRect(rectPanel, &globalData.brBarFace);

	return OnGetRibbonBackstageLeftPaneTextColor();
}
//****************************************************************************
COLORREF CBCGPVisualManager::OnGetRibbonBackstageLeftPaneTextColor()
{
	return globalData.clrBarText;
}
//****************************************************************************
COLORREF CBCGPVisualManager::OnDrawRibbonPanel (
		CDC* pDC,
		CBCGPRibbonPanel* pPanel, 
		CRect rectPanel,
		CRect /*rectCaption*/)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pPanel);

	if (pPanel->IsBackstageView())
	{
		rectPanel = ((CBCGPRibbonMainPanel*)pPanel)->GetCommandsFrame ();
		return OnFillRibbonBackstageLeftPane(pDC, rectPanel);
	}

	COLORREF clrText = globalData.clrBarText;

	if (pPanel->IsCollapsed () && pPanel->GetDefaultButton ().IsFocused ())
	{
		pDC->FillRect (rectPanel, &globalData.brHilite);
		clrText = globalData.clrTextHilite;
	}
	else if (pPanel->IsHighlighted () && !globalData.IsHighContastMode())
	{
		CBCGPDrawManager dm (*pDC);
		dm.HighlightRect (rectPanel);
	}

	pDC->Draw3dRect (rectPanel, globalData.clrBarHilite, globalData.clrBarHilite);
	rectPanel.OffsetRect (-1, -1);
	pDC->Draw3dRect (rectPanel, globalData.clrBarShadow, globalData.clrBarShadow);

	return clrText;
}
//****************************************************************************
void CBCGPVisualManager::OnDrawRibbonPanelCaption (
		CDC* pDC,
		CBCGPRibbonPanel* pPanel,
		CRect rectCaption)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pPanel);

	rectCaption.DeflateRect (1, 1);
	rectCaption.right -= 2;

	COLORREF clrText = globalData.clrBarText;

	if (!globalData.IsHighContastMode())
	{
		clrText = OnFillRibbonPanelCaption (pDC, pPanel, rectCaption);
	}

	COLORREF clrTextOld = pDC->SetTextColor (clrText);

	CString str = pPanel->GetDisplayName ();

	if (pPanel->GetLaunchButton ().GetID () > 0)
	{
		rectCaption.right = pPanel->GetLaunchButton ().GetRect ().left;
	}

	pDC->DrawText (	str, rectCaption, 
		DT_SINGLELINE | DT_CENTER | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);

	pDC->SetTextColor (clrTextOld);
}
//****************************************************************************
COLORREF CBCGPVisualManager::OnFillRibbonPanelCaption (CDC* pDC, CBCGPRibbonPanel* pPanel, CRect rectCaption)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pPanel);

	CBrush br (pPanel->IsHighlighted () ? 
		globalData.clrActiveCaption : globalData.clrInactiveCaption);
	pDC->FillRect (rectCaption, &br);

	return pPanel->IsHighlighted () ? globalData.clrCaptionText : globalData.clrInactiveCaptionText;
}
//****************************************************************************
void CBCGPVisualManager::OnDrawRibbonLaunchButton (
					CDC* pDC,
					CBCGPRibbonLaunchButton* pButton,
					CBCGPRibbonPanel* pPanel)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);
	ASSERT_VALID (pPanel);

	OnFillRibbonButton (pDC, pButton);

	COLORREF clrText = pPanel->IsHighlighted () ? 
		globalData.clrCaptionText : globalData.clrInactiveCaptionText;

	if (globalData.IsHighContastMode())
	{
		clrText = globalData.clrBarText;
	}
	
	CBCGPMenuImages::IMAGE_STATE imageState = CBCGPMenuImages::ImageBlack;
	
	if (pButton->IsDisabled ())
	{
		imageState = CBCGPMenuImages::ImageGray;
	}
	else if (!pButton->IsHighlighted () || globalData.IsHighContastMode())
	{
		imageState = CBCGPMenuImages::GetStateByColor(clrText, FALSE);
	}

	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdLaunchArrow, 
		pButton->GetRect (), imageState);

	OnDrawRibbonButtonBorder (pDC, pButton);
}
//****************************************************************************
void CBCGPVisualManager::OnDrawRibbonDefaultPaneButton (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->IsQATMode () || pButton->IsSearchResultMode ())
	{
		OnFillRibbonButton (pDC, pButton);
		OnDrawRibbonDefaultPaneButtonContext (pDC, pButton);
		OnDrawRibbonButtonBorder (pDC, pButton);
	}
	else
	{
		OnDrawRibbonDefaultPaneButtonContext (pDC, pButton);
	}
}
//****************************************************************************
void CBCGPVisualManager::OnDrawRibbonDefaultPaneButtonContext (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	CRect rectImage = pButton->GetRect ();

	if (pButton->IsSearchResultMode ())
	{
		rectImage.left += 3;
		rectImage.right = rectImage.left + pButton->GetImageSize (CBCGPRibbonButton::RibbonImageSmall).cx;

		if (pButton->HasImage(CBCGPBaseRibbonElement::RibbonImageSmall))
		{
			pButton->DrawImage (pDC, CBCGPRibbonButton::RibbonImageSmall, rectImage);
		}

		CRect rectText = pButton->GetRect ();
		rectText.left = rectImage.right + 3;

		CString strText = pButton->GetText ();
		pDC->DrawText (strText, rectText, DT_SINGLELINE | DT_VCENTER);
		return;
	}

	if (pButton->IsQATMode ())
	{
		pButton->DrawImage (pDC, CBCGPRibbonButton::RibbonImageSmall, pButton->GetRect ());
		return;
	}

	rectImage.top += 10;
	rectImage.bottom = rectImage.top + pButton->GetImageSize (CBCGPRibbonButton::RibbonImageSmall).cy;

	pButton->DrawImage (pDC, CBCGPRibbonButton::RibbonImageSmall, rectImage);

	// Draw text:
	pButton->DrawBottomText (pDC, FALSE);
}
//****************************************************************************
void CBCGPVisualManager::OnDrawRibbonDefaultPaneButtonIndicator (
					CDC* pDC, 
					CBCGPRibbonButton* /*pButton*/,
					CRect rect, 
					BOOL /*bIsSelected*/, 
					BOOL /*bHighlighted*/)
{
	ASSERT_VALID (pDC);

	rect.left = rect.right - rect.Height ();
	rect.DeflateRect (1, 1);

	pDC->FillRect (rect, &globalData.brBarFace);
	pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarShadow);

	CRect rectWhite = rect;
	rectWhite.OffsetRect (0, 1);

	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdArowDown, rectWhite,
		CBCGPMenuImages::ImageWhite);

	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdArowDown, rect,
		CBCGPMenuImages::ImageBlack);
}
//****************************************************************************
COLORREF CBCGPVisualManager::OnFillRibbonButton (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->IsBackstageViewMode () && !pButton->IsQATMode())
	{
		if (pButton->IsHighlighted ())
		{
			pDC->FillRect (pButton->GetRect (), &globalData.brHilite);
			return globalData.clrTextHilite;
		}
		else if (!pButton->IsDisabled () && pButton->GetBackstageAttachedView() != NULL && pButton->IsChecked())
		{
			pDC->FillRect (pButton->GetRect (), &globalData.brLight);
			pDC->Draw3dRect(pButton->GetRect (), globalData.clrBarShadow, globalData.clrBarHilite);
			return (COLORREF)-1;
		}
		
		return pButton->IsDisabled() ? globalData.clrGrayedText : globalData.clrBarText;
	}

	if (pButton->IsKindOf (RUNTIME_CLASS (CBCGPRibbonEdit)))
	{
		BOOL bIsReadOnly = ((CBCGPRibbonEdit*)pButton)->IsReadOnly();

		COLORREF clrBorder = globalData.clrBarShadow;
		CRect rectCommand = pButton->GetCommandRect ();

		CRect rect = pButton->GetRect ();
		rect.left = rectCommand.left;

		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect (rect, globalData.clrWindow, clrBorder);
		}
		else
		{
			if ((pButton->IsDroppedDown () || pButton->IsHighlighted ()) && !bIsReadOnly)
			{
				pDC->FillRect (rectCommand, &globalData.brWindow);
			}
			else
			{
				CBCGPDrawManager dm (*pDC);
				dm.HighlightRect (rectCommand);
			}

			pDC->Draw3dRect (rect, clrBorder, clrBorder);
		}

		return (COLORREF)-1;
	}

	if (pButton->IsMenuMode () && !pButton->IsPaletteIcon ())
	{
		if (pButton->IsHighlighted ())
		{
			pDC->FillRect (pButton->GetRect (), &globalData.brHilite);
			return globalData.clrTextHilite;
		}
	}
	else 
	{
		if (pButton->IsChecked () && !pButton->IsHighlighted ())
		{
			if (CBCGPToolBarImages::m_bIsDrawOnGlass)
			{
				CBCGPDrawManager dm (*pDC);
				dm.DrawRect (pButton->GetRect (), globalData.clrWindow, (COLORREF)-1);
			}
			else
			{
				CBCGPToolBarImages::FillDitheredRect (pDC, pButton->GetRect ());
			}
		}
	}

	return (COLORREF)-1;
}
//****************************************************************************
COLORREF CBCGPVisualManager::OnFillRibbonPinnedButton (
					CDC* pDC, 
					CBCGPRibbonButton* pButton, BOOL& bIsDarkPin)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	bIsDarkPin = TRUE;

	COLORREF clrText = (COLORREF)-1;
	BOOL bPinIsHighlighted = FALSE;

	if (pButton->IsCommandAreaHighlighted() || pButton->IsFocused())
	{
		pDC->FillRect(pButton->GetRect(), &globalData.brHilite);
		clrText = globalData.clrTextHilite;
		bPinIsHighlighted = TRUE;
	}
	else if (pButton->IsMenuAreaHighlighted())
	{
		pDC->FillRect(pButton->GetMenuRect(), &globalData.brHilite);
		bPinIsHighlighted = TRUE;
	}

	if (bPinIsHighlighted)
	{
		if (GetRValue (globalData.clrTextHilite) > 192 &&
			GetGValue (globalData.clrTextHilite) > 192 &&
			GetBValue (globalData.clrTextHilite) > 192)
		{
			bIsDarkPin = FALSE;
		}
	}

	return clrText;
}
//****************************************************************************
COLORREF CBCGPVisualManager::OnFillRibbonMainPanelButton (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	return OnFillRibbonButton (pDC, pButton);
}
//****************************************************************************
void CBCGPVisualManager::OnDrawRibbonMainPanelButtonBorder (
		CDC* pDC, CBCGPRibbonButton* pButton)
{
	OnDrawRibbonButtonBorder (pDC, pButton);
}
//****************************************************************************
COLORREF CBCGPVisualManager::GetRibbonEditBackgroundColor (
					CBCGPRibbonEditCtrl* pEdit,
					BOOL bIsHighlighted,
					BOOL /*bIsPaneHighlighted*/,
					BOOL bIsDisabled)
{
	BOOL bIsReadOnly = pEdit->GetSafeHwnd() != NULL && (pEdit->GetStyle() & ES_READONLY) == ES_READONLY;
	return (bIsHighlighted && !bIsDisabled && !bIsReadOnly) ? globalData.clrWindow : globalData.clrBarFace;
}
//****************************************************************************
COLORREF CBCGPVisualManager::GetRibbonEditTextColor (
					CBCGPRibbonEditCtrl* /*pEdit*/,
					BOOL /*bIsHighlighted*/,
					BOOL bIsDisabled)
{
	return bIsDisabled ? globalData.clrGrayedText : (COLORREF)-1;
}
//****************************************************************************
COLORREF CBCGPVisualManager::GetRibbonEditPromptColor(
					CBCGPRibbonEditCtrl* /*pEdit*/,
					BOOL /*bIsHighlighted*/,
					BOOL /*bIsDisabled*/)
{
	return GetToolbarEditPromptColor();
}
//****************************************************************************
void CBCGPVisualManager::OnDrawRibbonButtonBorder (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	CRect rect = pButton->GetRect ();

	if (pButton->IsKindOf (RUNTIME_CLASS (CBCGPRibbonEdit)))
	{
		if (globalData.IsHighContastMode())
		{
			pDC->Draw3dRect(rect, globalData.clrWindowText, globalData.clrWindowText);
		}
		return;
	}

	if (pButton->IsMenuMode () &&
		pButton->IsChecked () && !pButton->IsHighlighted ())
	{
		return;
	}

	if (pButton->IsHighlighted () || pButton->IsChecked () ||
		pButton->IsDroppedDown () || pButton->IsFocused ())
	{
		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect (rect, (COLORREF)-1, globalData.clrBarShadow);
		}
		else
		{
			if (pButton->IsPressed () || pButton->IsChecked () ||
				pButton->IsDroppedDown ())
			{
				pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarHilite);
			}
			else 
			{
				pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarShadow);
			}
		}

		CRect rectMenu = pButton->GetMenuRect ();
		
		if (!rectMenu.IsRectEmpty ())
		{
			if (CBCGPToolBarImages::m_bIsDrawOnGlass)
			{
				CBCGPDrawManager dm (*pDC);
				
				if (pButton->IsMenuOnBottom ())
				{
					dm.DrawLine (rectMenu.left, rectMenu.top, rectMenu.right, rectMenu.top, globalData.clrBarShadow);
				}
				else
				{
					dm.DrawLine (rectMenu.left, rectMenu.top, rectMenu.left, rectMenu.bottom, globalData.clrBarShadow);
				}
			}
			else
			{
				CPen* pOldPen = pDC->SelectObject (&globalData.penBarShadow);
				ASSERT (pOldPen != NULL);

				if (pButton->IsMenuOnBottom ())
				{
					pDC->MoveTo (rectMenu.left, rectMenu.top);
					pDC->LineTo (rectMenu.right, rectMenu.top);
				}
				else
				{
					pDC->MoveTo (rectMenu.left, rectMenu.top);
					pDC->LineTo (rectMenu.left, rectMenu.bottom);
				}

				pDC->SelectObject (pOldPen);
			}
		}
	}
}
//****************************************************************************
void CBCGPVisualManager::OnDrawRibbonPinnedButtonBorder (
					CDC* pDC, 
					CBCGPRibbonButton* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	if (pButton->IsCommandAreaHighlighted() || pButton->IsFocused())
	{
		pDC->Draw3dRect (pButton->GetRect(), globalData.clrBarHilite, globalData.clrBarShadow);
	}
	else if (pButton->IsMenuAreaHighlighted())
	{
		pDC->Draw3dRect (pButton->GetMenuRect(), globalData.clrBarHilite, globalData.clrBarShadow);
	}
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawRibbonMenuCheckFrame (
					CDC* pDC,
					CBCGPRibbonButton* /*pButton*/, 
					CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBtnFace);

	pDC->Draw3dRect (rect, globalData.clrBtnShadow,
							globalData.clrBtnHilite);
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawRibbonMainPanelFrame (
					CDC* pDC, 
					CBCGPRibbonMainPanel* /*pPanel*/,
					CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarShadow);
	rect.InflateRect (1, 1);
	pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarHilite);
}
//***********************************************************************************
void CBCGPVisualManager::OnFillRibbonMenuFrame (
					CDC* pDC, 
					CBCGPRibbonMainPanel* /*pPanel*/,
					CRect rect)
{
	ASSERT_VALID (pDC);
	pDC->FillRect (rect, &globalData.brWindow);
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawRibbonRecentFilesFrame (
					CDC* pDC, 
					CBCGPRibbonMainPanel* /*pPanel*/,
					CRect rect)
{
	ASSERT_VALID (pDC);

	pDC->FillRect (rect, &globalData.brBtnFace);

	CRect rectSeparator = rect;
	rectSeparator.right = rectSeparator.left + 2;

	pDC->Draw3dRect (rectSeparator, globalData.clrBtnShadow,
									globalData.clrBtnHilite);
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawRibbonLabel (
					CDC* /*pDC*/, 
					CBCGPRibbonLabel* /*pLabel*/,
					CRect /*rect*/)
{
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawRibbonPaletteButton (
					CDC* pDC, 
					CBCGPRibbonPaletteIcon* pButton)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	OnFillRibbonButton (pDC, pButton);
	OnDrawRibbonButtonBorder (pDC, pButton);
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawRibbonPaletteButtonIcon (
					CDC* pDC, 
					CBCGPRibbonPaletteIcon* pButton,
					int nID,
					CRect rectImage)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	CBCGPMenuImages::IMAGES_IDS id = (CBCGPMenuImages::IMAGES_IDS)nID;

	CRect rectWhite = rectImage;
	rectWhite.OffsetRect (0, 1);
	
	CBCGPMenuImages::Draw (pDC, id, rectWhite, CBCGPMenuImages::ImageWhite);
	CBCGPMenuImages::Draw (pDC, id, rectImage,
		pButton->IsDisabled() ? CBCGPMenuImages::ImageGray : CBCGPMenuImages::ImageBlack);
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawRibbonPaletteBorder (
				CDC* pDC, 
				CBCGPRibbonPaletteButton* /*pButton*/, 
				CRect rectBorder)
{
	ASSERT_VALID (pDC);
	pDC->Draw3dRect (rectBorder, globalData.clrBarShadow, globalData.clrBarShadow);
}
//***********************************************************************************
COLORREF CBCGPVisualManager::RibbonCategoryColorToRGB (BCGPRibbonCategoryColor color)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		switch (color)
		{
		case BCGPCategoryColor_Red:
			return RGB (255, 0, 0);

		case BCGPCategoryColor_Orange:
			return RGB (255, 128, 0);

		case BCGPCategoryColor_Yellow:
			return RGB (255, 255, 0);

		case BCGPCategoryColor_Green:
			return RGB (0, 255, 0);

		case BCGPCategoryColor_Blue:
			return RGB (0, 0, 255);

		case BCGPCategoryColor_Indigo:
			return RGB (0, 0, 128);

		case BCGPCategoryColor_Violet:
			return RGB (255, 0, 255);
		}

		return (COLORREF)-1;
	}

	switch (color)
	{
	case BCGPCategoryColor_Red:
		return RGB (255, 160, 160);

	case BCGPCategoryColor_Orange:
		return RGB (239, 189, 55);

	case BCGPCategoryColor_Yellow:
		return RGB (253, 229, 27);

	case BCGPCategoryColor_Green:
		return RGB (113, 190, 89);

	case BCGPCategoryColor_Blue:
		return RGB (128, 181, 196);

	case BCGPCategoryColor_Indigo:
		return RGB (114, 163, 224);

	case BCGPCategoryColor_Violet:
		return RGB (214, 178, 209);
	}

	return (COLORREF)-1;
}
//***********************************************************************************
COLORREF CBCGPVisualManager::OnDrawRibbonCategoryCaption (
					CDC* pDC, 
					CBCGPRibbonContextCaption* pContextCaption)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pContextCaption);

	COLORREF clrFill = RibbonCategoryColorToRGB (pContextCaption->GetColor ());
	CRect rect = pContextCaption->GetRect ();
	
	if (clrFill != (COLORREF)-1)
	{
		if (globalData.IsHighContastMode())
		{
			pDC->Draw3dRect(rect, clrFill, clrFill);
			rect.DeflateRect(1, 1);
			pDC->Draw3dRect(rect, clrFill, clrFill);
			return globalData.clrBarText;
		}

		CBCGPDrawManager dm(*pDC);

		if (CBCGPToolBarImages::m_bIsDrawOnGlass)
		{
			rect.DeflateRect (0, 1);
			dm.DrawRect (rect, clrFill, (COLORREF)-1);
		}
		else
		{
			CBrush br (clrFill);
			pDC->FillRect (rect, &br);
		}

		int xTabRight = pContextCaption->GetRightTabX ();
		
		if (xTabRight > 0)
		{
			CRect rectTab (pContextCaption->GetParentRibbonBar ()->GetActiveCategory ()->GetTabRect ());
			
			CRect rectLeft = rectTab;
			rectLeft.left = rect.left;
			rectLeft.right = rect.left + 1;
			rectLeft.bottom = rectLeft.top + 2 * rectTab.Height() / 3;
			
			dm.FillGradient(rectLeft, globalData.clrBarFace, globalData.clrBarShadow, TRUE);
			
			CRect rectRight = rectTab;
			rectRight.left = xTabRight - 1;
			rectRight.right = xTabRight;
			rectRight.bottom = rectRight.top + 2 * rectTab.Height() / 3;
			
			dm.FillGradient(rectRight, globalData.clrBarFace, globalData.clrBarShadow, TRUE);
		}
	}

	return globalData.clrBarText;
}
//***********************************************************************************
void CBCGPVisualManager::OnDrawRibbonCategoryCaptionText (
					CDC* pDC, CBCGPRibbonContextCaption* /*pContextCaption*/, 
					CString& strTextIn, CRect rectText, BOOL bIsOnGlass, BOOL bIsZoomed)
{
	ASSERT_VALID (pDC);

	const UINT uiDTFlags = DT_CENTER | DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX;

	CString strText(strTextIn);
	
	if (IsTopLevelMenuItemUpperCase())
	{
		strText.MakeUpper();
	}

	if (bIsOnGlass)
	{
		DrawTextOnGlass (
			pDC, strText, rectText, uiDTFlags,
			bIsZoomed && !globalData.bIsWindows7 ? 0 : 10,
			bIsZoomed && !globalData.bIsWindows7 ? RGB (255, 255, 255) : (COLORREF)-1);
	}
	else
	{
		pDC->DrawText (strText, rectText, uiDTFlags);
	}
}
//********************************************************************************
COLORREF CBCGPVisualManager::OnDrawRibbonStatusBarPane (CDC* pDC, CBCGPRibbonStatusBar* /*pBar*/,
					CBCGPRibbonStatusBarPane* pPane)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pPane);

	CRect rect = pPane->GetRect ();

	if (pPane->IsHighlighted ())
	{
		CRect rectButton = rect;
		rectButton.DeflateRect (1, 1);

		pDC->Draw3dRect (rectButton, 
			pPane->IsPressed () ? globalData.clrBarShadow : globalData.clrBarHilite,
			pPane->IsPressed () ? globalData.clrBarHilite : globalData.clrBarShadow);
	}

	return (COLORREF)-1;
}
//********************************************************************************
void CBCGPVisualManager::GetRibbonSliderColors (CBCGPRibbonSlider* /*pSlider*/,
							BOOL bIsHighlighted, 
							BOOL bIsPressed,
							BOOL bIsDisabled,
							COLORREF& clrLine,
							COLORREF& clrFill)
{
	clrLine = bIsDisabled ? globalData.clrBarShadow : globalData.clrBarDkShadow;
	clrFill = bIsPressed && bIsHighlighted ?
		globalData.clrBarShadow :
		bIsHighlighted ? globalData.clrBarHilite : globalData.clrBarFace;
}
//********************************************************************************
void CBCGPVisualManager::OnDrawRibbonSliderZoomButton (
			CDC* pDC, CBCGPRibbonSlider* pSlider, 
			CRect rect, BOOL bIsZoomOut, 
			BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled)
{
	ASSERT_VALID (pDC);

	int nIconsCount = 6;
	int nRadius = globalUtils.ScaleByDPI(7);

	CPoint ptCenter = rect.CenterPoint ();
	CRect rectCircle (CPoint (ptCenter.x - nRadius, ptCenter.y - nRadius), CSize (nRadius * 2 + 1, nRadius * 2 + 1));

	if (m_pImagesZoom == NULL && !globalData.IsHighContastMode() && globalData.m_nBitsPerPixel > 8)
	{
		CBCGPGraphicsManagerParams params;
		params.bAlphaModePremultiplied = TRUE;

		CBCGPGraphicsManager* pGM = CBCGPGraphicsManager::CreateInstance(CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER_D2D, FALSE, &params);
		if (pGM != NULL)
		{
			pGM->EnableTransparentGradient();

			CSize sizeIcon = rectCircle.Size();

			CBCGPSize sizeImageList = sizeIcon;
			sizeImageList.cx *= nIconsCount;
			
			HBITMAP hbmp = CBCGPDrawManager::CreateBitmap_32(sizeImageList, NULL);
			if (hbmp != NULL)
			{
				CBCGPRect rectGM(CBCGPPoint(), sizeImageList);
				
				CDC dcMem;
				dcMem.CreateCompatibleDC(NULL);
				
				HBITMAP hbmpOld = (HBITMAP) dcMem.SelectObject(hbmp);
				
				pGM->BindDC(&dcMem, rectGM);
				pGM->BeginDraw();
				
				double xOffset = 0;

				for (int i = 0; i < nIconsCount; i++)
				{
					CBCGPRect rectIcon(CBCGPPoint(xOffset, 0), sizeIcon);
					rectIcon.DeflateRect(1, 1);

					CBCGPEllipse ellipse(rectIcon);

					COLORREF clrLine = (COLORREF)-1;
					COLORREF clrFill = (COLORREF)-1;

					BOOL bIconHighlighted = (i == 1 || i == 4);
					BOOL bIconPressed = (i == 2 || i == 5);

					GetRibbonSliderColors(pSlider, bIconHighlighted || bIconPressed, bIconPressed, FALSE, clrLine, clrFill);

					CBCGPColor clr = clrFill;
					CBCGPColor clrGradient = clrFill;

					clr.MakeDarker(.2);
					clrGradient.MakePale();

					pGM->FillEllipse(ellipse, CBCGPBrush(clr, clrGradient, bIconPressed ? CBCGPBrush::BCGP_GRADIENT_RADIAL_BOTTOM_RIGHT : CBCGPBrush::BCGP_GRADIENT_RADIAL_TOP_LEFT));
					
					CBCGPBrush brLine(clrLine);

					CBCGPPoint ptCenterIcon = rectIcon.CenterPoint();
					CBCGPRect rectSign(CBCGPPoint (ptCenterIcon.x - 0.5 * nRadius, ptCenterIcon.y - 0.5 * nRadius), CBCGPSize(nRadius, nRadius));

					pGM->DrawLine(rectSign.left, ptCenterIcon.y, rectSign.right, ptCenterIcon.y, brLine, 2.0);
					
					if (i >= nIconsCount / 2)
					{
						pGM->DrawLine(ptCenterIcon.x, rectSign.top, ptCenterIcon.x, rectSign.bottom, brLine, 2.0);
					}

					pGM->DrawEllipse(ellipse, brLine);

					xOffset += sizeIcon.cx;
				}

				pGM->EndDraw();
				
				dcMem.SelectObject (hbmpOld);
				
				pGM->BindDC(NULL);
	
				m_pImagesZoom = new CBCGPToolBarImages();
				m_pImagesZoom->SetImageSize(sizeIcon);
				m_pImagesZoom->AddImage(hbmp, TRUE);
				
				::DeleteObject(hbmp);
			}

			delete pGM;
		}
	}

	if (m_pImagesZoom != NULL)
	{
		int nIconIndex = bIsPressed ? 2 : bIsHighlighted ? 1 : 0;
		if (!bIsZoomOut)
		{
			nIconIndex += nIconsCount / 2;
		}

		if (m_pImagesZoom->DrawEx(pDC, rect, nIconIndex, CBCGPToolBarImages::ImageAlignHorzCenter, CBCGPToolBarImages::ImageAlignVertCenter, CRect(0, 0, 0, 0), bIsDisabled ? 127 : 255))
		{
			return;
		}
	}

	COLORREF clrLine = (COLORREF)-1;
	COLORREF clrFill = (COLORREF)-1;

	GetRibbonSliderColors(pSlider, bIsHighlighted, bIsPressed, bIsDisabled, clrLine, clrFill);
							
	CBCGPDrawManager dm (*pDC);

	dm.DrawEllipse (rectCircle, clrFill, clrLine);

	// Draw +/- sign:
	CRect rectSign (CPoint (ptCenter.x - nRadius / 2, ptCenter.y - nRadius / 2), CSize(nRadius, nRadius));

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		dm.DrawLine (rectSign.left, ptCenter.y, rectSign.right, ptCenter.y, clrLine);

		if (!bIsZoomOut)
		{
			dm.DrawLine (ptCenter.x, rectSign.top, ptCenter.x, rectSign.bottom, clrLine);
		}
	}
	else
	{
		CPen penLine (PS_SOLID, 2, clrLine);
		CPen* pOldPen = pDC->SelectObject (&penLine);

		pDC->MoveTo (rectSign.left - 1, ptCenter.y + 1);
		pDC->LineTo (rectSign.right, ptCenter.y + 1);

		if (!bIsZoomOut)
		{
			pDC->MoveTo (ptCenter.x + 1, rectSign.top);
			pDC->LineTo (ptCenter.x + 1, rectSign.bottom);
		}

		pDC->SelectObject (pOldPen);
	}
}
//********************************************************************************
void CBCGPVisualManager::OnDrawRibbonSliderChannel (
			CDC* pDC, CBCGPRibbonSlider* pSlider, CRect rect)
{
	BOOL bIsVert = FALSE;

	if (pSlider != NULL)
	{
		ASSERT_VALID (pSlider);
		bIsVert = pSlider->IsVert ();
	}

	OnDrawSliderChannel (pDC, NULL, bIsVert, rect, CBCGPToolBarImages::m_bIsDrawOnGlass);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawRibbonSliderThumb (
			CDC* pDC, CBCGPRibbonSlider* pSlider, 
			CRect rect, BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled)
{
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
//********************************************************************************
void CBCGPVisualManager::OnDrawRibbonProgressBar (CDC* pDC, 
												  CBCGPRibbonProgressBar* /*pProgress*/, 
												  CRect rectProgress, CRect rectChunk,
												  BOOL /*bInfiniteMode*/)
{
	ASSERT_VALID (pDC);

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);
		
		if (!rectChunk.IsRectEmpty ())
		{
			rectChunk.DeflateRect(1, 1);
			dm.DrawRect (rectChunk, globalData.clrHilite, (COLORREF)-1);
		}

		dm.DrawRect (rectProgress, (COLORREF)-1, globalData.clrBarShadow);
	}
	else
	{
		if (!rectChunk.IsRectEmpty ())
		{
			rectChunk.DeflateRect(1, 1);
			pDC->FillRect (rectChunk, &globalData.brHilite);
		}

		pDC->Draw3dRect (rectProgress, globalData.clrBarShadow, globalData.clrBarHilite);
	}
}
//********************************************************************************
void CBCGPVisualManager::OnFillRibbonQATPopup(CDC* pDC, CBCGPRibbonPanelMenuBar* /*pMenuBar*/, CRect rect)
{
	ASSERT_VALID(pDC);
	pDC->FillRect (rect, &globalData.brBarFace);
}
//********************************************************************************
void CBCGPVisualManager::OnFillRibbonPanelMenuBarInCtrlMode(CDC* pDC, CBCGPRibbonPanelMenuBar* /*pBar*/, CRect rectClient)
{
	ASSERT_VALID(pDC);
	pDC->FillRect(rectClient, &globalData.brWindow);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawRibbonQATSeparator (CDC* pDC, 
												   CBCGPRibbonSeparator* /*pSeparator*/, CRect rect)
{
	ASSERT_VALID (pDC);

	if (CBCGPToolBarImages::m_bIsDrawOnGlass)
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawRect (rect, (COLORREF)-1, globalData.clrBtnShadow);
	}
	else
	{
		pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarHilite);
	}
}
//********************************************************************************
void CBCGPVisualManager::OnDrawRibbonKeyTip (CDC* pDC, CBCGPBaseRibbonElement* pElement, 
											 CRect rect, CString str)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pElement);

	COLORREF clrText = ::GetSysColor (COLOR_INFOTEXT);
	COLORREF clrTextDisabled = globalData.clrBtnShadow;
	COLORREF clrBorder = clrText;

	if (m_pfDrawThemeBackground != NULL && m_hThemeToolTip != NULL)
	{
		CRect rectFill = rect;
		rectFill.InflateRect (2, 2);

		(*m_pfDrawThemeBackground) (m_hThemeToolTip, pDC->GetSafeHdc(), TTP_STANDARD, 0, &rectFill, 0);

		if (m_pfGetThemeColor != NULL)
		{
			(*m_pfGetThemeColor) (m_hThemeToolTip, TTP_STANDARD, 0, TMT_TEXTCOLOR, &clrText);
			(*m_pfGetThemeColor) (m_hThemeToolTip, TTP_STANDARD, 0, TMT_TEXTSHADOWCOLOR, &clrTextDisabled);
			(*m_pfGetThemeColor) (m_hThemeToolTip, TTP_STANDARD, 0, TMT_EDGELIGHTCOLOR, &clrBorder);
		}
	}
	else
	{
		::FillRect (pDC->GetSafeHdc (), rect, ::GetSysColorBrush (COLOR_INFOBK));
	}

	str.MakeUpper ();

	COLORREF clrTextOld = pDC->SetTextColor(pElement->IsDisabled () ? clrTextDisabled : clrText);
	
	pDC->DrawText (str, rect, DT_SINGLELINE | DT_VCENTER | DT_CENTER);

	pDC->SetTextColor (clrTextOld);

	pDC->Draw3dRect (rect, clrBorder, clrBorder);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawRibbonRadioButtonOnList (CDC* pDC, CBCGPRibbonButton* /*pButton*/, 
													 CRect rect, BOOL /*bIsSelected*/, BOOL /*bHighlighted*/)
{
	ASSERT_VALID (pDC);

	rect.OffsetRect (1, 1);
	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdRadio, rect, CBCGPMenuImages::ImageWhite);

	rect.OffsetRect (-1, -1);
	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdRadio, rect, CBCGPMenuImages::ImageBlack);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawRibbonCheckBoxOnList (CDC* pDC, CBCGPRibbonButton* /*pButton*/, 
													 CRect rect, BOOL /*bIsSelected*/, BOOL /*bHighlighted*/)
{
	ASSERT_VALID (pDC);

	rect.OffsetRect (1, 1);
	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdCheck, rect, CBCGPMenuImages::ImageWhite);

	rect.OffsetRect (-1, -1);
	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdCheck, rect, CBCGPMenuImages::ImageBlack);
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetRibbonHyperlinkTextColor (CBCGPRibbonHyperlink* pHyperLink)
{
	ASSERT_VALID (pHyperLink);

	if (pHyperLink->IsDisabled ())
	{
		return GetToolbarDisabledTextColor ();
	}

	return pHyperLink->IsHighlighted () ? globalData.clrHotLinkText : globalData.clrHotText;
}
//********************************************************************************
COLORREF CBCGPVisualManager::GetRibbonStatusBarTextColor (CBCGPRibbonStatusBar* /*pStatusBar*/)
{
	return globalData.clrBarText;
}
//********************************************************************************
int CBCGPVisualManager::GetRibbonPanelMargin(CBCGPRibbonCategory* pCategory)
{
	if (pCategory != NULL)
	{
		CBCGPRibbonMainPanel* pPanel = pCategory->GetMainPanel();
		if (pPanel != NULL)
		{
			if (pPanel->IsBackstageView ())
			{
				return 0;
			}
			else if (!pPanel->IsScenicLook ())
			{
				return 4;
			}
			else
			{
				return 2;
			}
		}
	}

	return 2;
}
//********************************************************************************
void CBCGPVisualManager::OnDrawRibbonColorPaletteBox (CDC* pDC, CBCGPRibbonColorButton* pColorButton,
		CBCGPRibbonPaletteIcon* pIcon,
		COLORREF color, CRect rect, BOOL bDrawTopEdge, BOOL bDrawBottomEdge,
		BOOL bIsHighlighted, BOOL bIsChecked, BOOL /*bIsDisabled*/)
{
	ASSERT_VALID (pDC);

	CRect rectFill = rect;
	rectFill.DeflateRect (1, 0);

	if (bIsHighlighted || bIsChecked)
	{
		CBCGPToolBarImages::FillDitheredRect (pDC, rect);
		rectFill.DeflateRect (1, 2);
	}

	if (color != (COLORREF)-1)
	{
		CBrush br (color);
		pDC->FillRect (rectFill, &br);
	}

	COLORREF clrBorder = globalData.clrBtnShadow;

	if (bDrawTopEdge && bDrawBottomEdge)
	{
		pDC->Draw3dRect (rect, clrBorder, clrBorder);
	}
	else
	{
		CPen penBorder (PS_SOLID, 1, clrBorder);

		CPen* pOldPen = pDC->SelectObject (&penBorder);
		ASSERT (pOldPen != NULL);

		pDC->MoveTo (rect.left, rect.top);
		pDC->LineTo (rect.left, rect.bottom);

		pDC->MoveTo (rect.right - 1, rect.top);
		pDC->LineTo (rect.right - 1, rect.bottom);

		if (bDrawTopEdge)
		{
			pDC->MoveTo (rect.left, rect.top);
			pDC->LineTo (rect.right, rect.top);
		}

		if (bDrawBottomEdge)
		{
			pDC->MoveTo (rect.left, rect.bottom - 1);
			pDC->LineTo (rect.right, rect.bottom - 1);
		}

		pDC->SelectObject (pOldPen);
	}

	OnDrawRibbonColorPaletteBoxHotBorder(pDC, pColorButton, pIcon, rect, bIsHighlighted, bIsChecked);
}
//********************************************************************************
void CBCGPVisualManager::OnDrawRibbonColorPaletteBoxHotBorder (CDC* pDC, CBCGPRibbonColorButton* /*pColorButton*/,
													  CBCGPRibbonPaletteIcon* /*pIcon*/,
													  CRect rect, BOOL bIsHighlighted, BOOL bIsChecked)
{
	ASSERT_VALID (pDC);
	
	if (bIsHighlighted)
	{
		pDC->Draw3dRect (&rect, globalData.clrBarHilite, globalData.clrBarShadow);
	}
	else if (bIsChecked)
	{
		pDC->Draw3dRect (&rect, globalData.clrBarShadow, globalData.clrBarHilite);
	}
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawRibbonMinimizeButtonImage(CDC* pDC, CBCGPRibbonMinimizeButton* pButton, BOOL bRibbonIsMinimized)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pButton);

	CBCGPMenuImages::IMAGE_STATE state = CBCGPMenuImages::ImageBlack;

	if (pButton->IsDisabled())
	{
		state = CBCGPMenuImages::ImageLtGray;
	}
	else if (!pButton->IsHighlighted() && IsRibbonTabsAreaDark())
	{
		state = CBCGPMenuImages::ImageWhite;
	}

	CBCGPMenuImages::Draw(pDC, 
		bRibbonIsMinimized ? CBCGPMenuImages::IdMinimizeRibbon : CBCGPMenuImages::IdMaximizeRibbon, pButton->GetRect(), state);
}
//*********************************************************************************
CSize CBCGPVisualManager::GetRibbonMinimizeButtonImageSize()
{
	return CBCGPMenuImages::Size();
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawRibbonBackstageTopLine(CDC* pDC, CRect rectLine)
{
	ASSERT_VALID (pDC);

	CBCGPDrawManager dm (*pDC);
	dm.DrawRect (rectLine, globalData.clrBarShadow, (COLORREF)-1);
}
//*****************************************************************************
COLORREF CBCGPVisualManager::GetRibbonBackstageTextColor()
{	
	return globalData.IsHighContastMode() ? globalData.clrBarText : RGB(0, 0, 0);
}
//*****************************************************************************
COLORREF CBCGPVisualManager::GetRibbonBackstageInfoTextColor()
{
	return GetRibbonBackstageTextColor();
}
//*****************************************************************************
void CBCGPVisualManager::SetMainButtonColor(COLORREF clr)
{
	m_clrMainButton = clr;
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawRibbonMenuArrow(CDC* pDC, CBCGPRibbonButton* pButton, UINT idIn, const CRect& rectMenuArrow)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pButton);

	CBCGPMenuImages::IMAGES_IDS id = (CBCGPMenuImages::IMAGES_IDS)idIn;

	CRect rectWhite = rectMenuArrow;
	rectWhite.OffsetRect (0, 1);
	
	CBCGPMenuImages::Draw (pDC, id, rectWhite, CBCGPMenuImages::ImageWhite);
	CBCGPMenuImages::Draw (pDC, id, rectMenuArrow, 
		pButton->IsDisabled() ? CBCGPMenuImages::ImageGray : CBCGPMenuImages::ImageBlack);
}
//*********************************************************************************
void CBCGPVisualManager::OnDrawRibbonAutoHideButton(CDC* pDC, const CRect& rect, BOOL bIsHighlighted, COLORREF clrFill, COLORREF clrText)
{
	ASSERT_VALID(pDC);

	if (bIsHighlighted)
	{
		if (globalData.IsHighContastMode())
		{
			pDC->FillRect(rect, &globalData.brHilite);
			clrText = globalData.clrTextHilite;
		}
		else
		{
			if (clrFill == CLR_DEFAULT)
			{
				clrFill = GetHighlightedColor(0);
			}

			CBrush br(clrFill);
			pDC->FillRect(rect, &br);

			if (clrText == CLR_DEFAULT)
			{
				clrText = CBCGPDrawManager::GetContrastColor(clrFill);
			}
		}
	}
	
	CBCGPDrawManager dm(*pDC);
	
	int ellipse_size = min(rect.Height(), globalUtils.ScaleByDPI(3));
	int x = rect.right - 3 * ellipse_size;

	if (clrText == CLR_DEFAULT)
	{
		clrText = globalData.clrBarText;
	}

	for (int i = 0; i < 3; i++)
	{
		CRect rectEllipse = rect;
		
		rectEllipse.left = x;
		rectEllipse.right = rectEllipse.left + ellipse_size;
		rectEllipse.top = rectEllipse.CenterPoint().y - ellipse_size / 2;
		rectEllipse.bottom = rectEllipse.top + ellipse_size;
		
		dm.DrawEllipse(rectEllipse, clrText, (COLORREF)-1);

		x -= 2 * ellipse_size;
	}
}

#endif // BCGP_EXCLUDE_RIBBON

BOOL CBCGPVisualManager::OnSetWindowRegion (CWnd* pWnd, CSize sizeWindow)
{
#ifdef BCGP_EXCLUDE_RIBBON
	UNREFERENCED_PARAMETER(pWnd);
	UNREFERENCED_PARAMETER(sizeWindow);
	return FALSE;
#else
	if (globalData.DwmIsCompositionEnabled () && IsDWMCaptionSupported())
	{
		return FALSE;
	}

	ASSERT_VALID (pWnd);

	CBCGPRibbonBar* pRibbonBar = NULL;
	BOOL IsFullScreen = FALSE;

	if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPFrameWnd)))
	{
		pRibbonBar = ((CBCGPFrameWnd*) pWnd)->GetRibbonBar ();
		IsFullScreen = ((CBCGPFrameWnd*) pWnd)->IsFullScreen ();
	}
	else if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPMDIFrameWnd)))
	{
		pRibbonBar = ((CBCGPMDIFrameWnd*) pWnd)->GetRibbonBar ();
		IsFullScreen = ((CBCGPMDIFrameWnd*) pWnd)->IsFullScreen ();
	}

	if (IsFullScreen)
	{
		pWnd->SetWindowRgn (NULL, FALSE);
		return FALSE;
	}

	if (pRibbonBar->GetSafeHwnd() == NULL || 
		!pRibbonBar->IsWindowVisible () ||
		!pRibbonBar->IsReplaceFrameCaption ())
	{
		return FALSE;
	}

	if (pWnd->IsZoomed())
	{
		pWnd->SetWindowRgn (NULL, FALSE);
		return TRUE;
	}

	const int nLeftRadius  = 11;
	const int nRightRadius = 11;

	CRgn rgnWnd;
	rgnWnd.CreateRectRgn (0, 0, sizeWindow.cx, sizeWindow.cy);

	CRgn rgnTemp;

	rgnTemp.CreateRectRgn (0, 0, nLeftRadius / 2, nLeftRadius / 2);
	rgnWnd.CombineRgn (&rgnTemp, &rgnWnd, RGN_XOR);

	rgnTemp.DeleteObject ();
	rgnTemp.CreateEllipticRgn (0, 0, nLeftRadius, nLeftRadius);
	rgnWnd.CombineRgn (&rgnTemp, &rgnWnd, RGN_OR);

	rgnTemp.DeleteObject ();
	rgnTemp.CreateRectRgn (sizeWindow.cx - nRightRadius / 2, 0, sizeWindow.cx, nRightRadius / 2);
	rgnWnd.CombineRgn (&rgnTemp, &rgnWnd, RGN_XOR);

	rgnTemp.DeleteObject ();
	rgnTemp.CreateEllipticRgn (sizeWindow.cx - nRightRadius + 1, 0, sizeWindow.cx + 1, nRightRadius);
	rgnWnd.CombineRgn (&rgnTemp, &rgnWnd, RGN_OR);
	
	pWnd->SetWindowRgn ((HRGN)rgnWnd.Detach (), TRUE);
	return TRUE;
#endif
}
//*****************************************************************************
BOOL CBCGPVisualManager::OnNcPaint (CWnd* /*pWnd*/, const CObList& /*lstSysButtons*/, CRect /*rectRedraw*/)
{
	return FALSE;
}
//*****************************************************************************
BOOL CBCGPVisualManager::CanSetActivateFlag(CWnd* pWnd) const
{
	ASSERT_VALID (pWnd);
	
	if (pWnd->GetSafeHwnd () == NULL)
	{
		return FALSE;
	}
	
	BOOL bIsRibbonCaption = FALSE;
	
#ifndef BCGP_EXCLUDE_RIBBON
	CBCGPRibbonBar* pBar = GetRibbonBar(pWnd);
	if (pBar->GetSafeHwnd() != NULL && pBar->IsWindowVisible() && pBar->IsReplaceFrameCaption())
	{
		bIsRibbonCaption = globalData.bIsWindows10;
	}
#endif

	return bIsRibbonCaption;
}
//*****************************************************************************
void CBCGPVisualManager::SetActivateFlag(CWnd* pWnd, BOOL bActive)
{
	m_ActivateFlag[(UINT_PTR)pWnd->GetSafeHwnd ()] = bActive;
	pWnd->SendMessage (WM_NCPAINT, 0, 0);
}
//*****************************************************************************
void CBCGPVisualManager::OnActivateApp (CWnd* pWnd, BOOL bActive)
{
	ASSERT_VALID (pWnd);
	
	if (!CanSetActivateFlag(pWnd))
	{
		return;
	}
	
	// but do not stay active if the window is disabled
	if (!pWnd->IsWindowEnabled())
	{
		bActive = FALSE;
	}

	SetActivateFlag(pWnd, bActive);
}
//*****************************************************************************
BOOL CBCGPVisualManager::OnNcActivate (CWnd* pWnd, BOOL bActive)
{
	ASSERT_VALID (pWnd);

	if (!CanSetActivateFlag(pWnd))
	{
		return FALSE;
	}

	// stay active if WF_STAYACTIVE bit is on
	if (pWnd->m_nFlags & WF_STAYACTIVE)
	{
		bActive = TRUE;
	}

	// but do not stay active if the window is disabled
	if (!pWnd->IsWindowEnabled())
	{
		bActive = FALSE;
	}

	SetActivateFlag(pWnd, bActive);

	return TRUE;
}
//*****************************************************************************
void CBCGPVisualManager::ResetActivateFlag(CWnd* pWnd)
{
	HWND hWnd = pWnd->GetSafeHwnd ();
	if (hWnd != NULL)
	{
		BOOL bActive = FALSE;

		if (m_ActivateFlag.Lookup((UINT_PTR)hWnd, bActive) && bActive)
		{
			m_ActivateFlag[(UINT_PTR)hWnd] = FALSE;
		}
	}
}
//*****************************************************************************
void CBCGPVisualManager::RemoveActivateFlag(CWnd* pWnd)
{
	HWND hWnd = pWnd->GetSafeHwnd ();
	if (hWnd != NULL)
	{
		m_ActivateFlag.RemoveKey((UINT_PTR)hWnd);
	}
}
//*****************************************************************************
CSize CBCGPVisualManager::GetNcBtnSize (BOOL /*bSmall*/) const
{
	return CSize (0, 0);
}
//*****************************************************************************
BOOL CBCGPVisualManager::OnUpdateNcButtons (CWnd* /*pWnd*/, const CObList& /*lstSysButtons*/, CRect /*rectCaption*/)
{
    return FALSE;
}
//*****************************************************************************
BOOL CBCGPVisualManager::IsWindowActive (CWnd* pWnd) const
{
	BOOL bActive = FALSE;

	HWND hWnd = pWnd->GetSafeHwnd ();

	if (hWnd != NULL)
	{
		if (!m_ActivateFlag.Lookup ((UINT_PTR)pWnd->GetSafeHwnd (), bActive))
		{
			bActive = TRUE;
		}
	}

	return bActive;
}
//*****************************************************************************
CSize CBCGPVisualManager::GetSystemBorders (BOOL bRibbonPresent) const
{
	CSize size(::GetSystemMetrics(SM_CXSIZEFRAME), ::GetSystemMetrics(SM_CYSIZEFRAME));
	
	if (bRibbonPresent)
	{
		size.cx--;
		size.cy--;
	}
	
	return size;
}
//*****************************************************************************
BOOL CBCGPVisualManager::OnEraseMDIClientArea (CDC* /*pDC*/, CRect /*rectClient*/)
{
	return FALSE;
}
//*******************************************************************************
BOOL CBCGPVisualManager::GetToolTipParams (CBCGPToolTipParams& params, 
										   UINT /*nType*/ /*= (UINT)(-1)*/)
{
	CBCGPToolTipParams dummy;
	params = dummy;

	return TRUE;
}
//*******************************************************************************
void CBCGPVisualManager::OnFillToolTip (CDC* pDC, CBCGPToolTipCtrl* /*pToolTip*/, CRect rect,
										COLORREF& clrText, COLORREF& clrLine)
{
	if (m_pfDrawThemeBackground != NULL && m_hThemeToolTip != NULL)
	{
		(*m_pfDrawThemeBackground) (m_hThemeToolTip, pDC->GetSafeHdc(), TTP_STANDARD, 0, &rect, 0);

		if (m_pfGetThemeColor != NULL)
		{
			(*m_pfGetThemeColor) (m_hThemeToolTip, TTP_STANDARD, 0, TMT_TEXTCOLOR, &clrText);
			(*m_pfGetThemeColor) (m_hThemeToolTip, TTP_STANDARD, 0, TMT_EDGEDKSHADOWCOLOR, &clrLine);
		}
	}
	else
	{
		::FillRect (pDC->GetSafeHdc (), rect, ::GetSysColorBrush (COLOR_INFOBK));
	}
}
//*******************************************************************************
CSize CBCGPVisualManager::GetSystemToolTipCornerRadius(CToolTipCtrl* pToolTip)
{
	if (m_pfGetWindowTheme != NULL && (*m_pfGetWindowTheme) (pToolTip->GetSafeHwnd ()) != NULL)
	{
		CSize size(3, 3);
		return size;
	}
	
	return CSize(0, 0);	// Tooltip is not themed
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawControlInfoTip(CDC* pDC, CRect rect, CWnd* pWndCtrl)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pWndCtrl);

	CWnd* pDlg = pWndCtrl->GetParent();
	ASSERT(pDlg->GetSafeHwnd() != NULL);

	CBCGPDrawManager dm(*pDC);
	
	COLORREF clrLine = GetDlgHeaderTextColor(pDlg);
	COLORREF clrText = GetDlgHeaderTextColor(pDlg);
	COLORREF clrFill = globalData.clrBtnHilite;
	
	if (!globalData.IsHighContastMode() && globalData.m_nBitsPerPixel > 8)
	{
		clrFill = RGB(255, 255, 255);
		clrLine = RGB(147, 175, 215);
		clrText = RGB(90, 127, 173);
	}
	
	dm.DrawEllipse(rect, clrFill, clrLine);
	
	int x = rect.CenterPoint().x;
	int y = rect.top + globalUtils.ScaleByDPI(1);
	
	dm.DrawLine(x, y + globalUtils.ScaleByDPI(2), x, y + globalUtils.ScaleByDPI(3), clrText);
	
	if (rect.Height() > globalUtils.ScaleByDPI(10))
	{
		dm.DrawLine(x, y + globalUtils.ScaleByDPI(5), x, rect.bottom - globalUtils.ScaleByDPI(2), clrText);
	}
}
//*******************************************************************************
BOOL CBCGPVisualManager::DrawTextOnGlass (CDC* pDC, CString strText, CRect rect, 
										  DWORD dwFlags, int nGlowSize, COLORREF clrText)
{
	ASSERT_VALID (pDC);

	COLORREF clrOldText = pDC->GetTextColor ();
	pDC->SetTextColor (RGB (0, 0, 0));	// TODO

	BOOL bRes = globalData.DrawTextOnGlass (m_hThemeButton, pDC, 0, 0, strText, rect, 
		dwFlags, nGlowSize, clrText);

	pDC->SetTextColor (clrOldText);

	return bRes;
}
//*******************************************************************************
BOOL CBCGPVisualManager::DrawIconOnGlass(CDC* pDC, HICON hIcon, CRect rect)
{
	ASSERT_VALID (pDC);

	if (hIcon == NULL)
	{
		return FALSE;
	}

	return globalData.DrawIconOnGlass(m_hThemeWindow, pDC, hIcon, rect);
}
//*******************************************************************************
COLORREF CBCGPVisualManager::OnDrawEditListCaption(CDC* pDC, CBCGPEditListBase* pList, CRect rect)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pList);

	pDC->FillRect(rect, &GetDlgBackBrush(pList->GetParent()));
	pDC->Draw3dRect(rect, globalData.clrBarShadow, globalData.clrBarHilite);

	return globalData.clrBarText;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::OnFillListBox(CDC* pDC, CBCGPListBox* /*pListBox*/, CRect rectClient)
{
	ASSERT_VALID (pDC);

	pDC->FillRect(rectClient, &globalData.brWindow);
	return globalData.clrWindowText;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::OnFillListBoxItem (CDC* pDC, CBCGPListBox* /*pListBox*/, int /*nItem*/, CRect rect, BOOL bIsHighlihted, BOOL bIsSelected)
{
	ASSERT_VALID (pDC);

	COLORREF clrText = (COLORREF)-1;

	if (bIsSelected)
	{
		pDC->FillRect (rect, &globalData.brHilite);
		clrText = globalData.clrTextHilite;
	}

	if (bIsHighlihted)
	{
		pDC->DrawFocusRect (rect);
	}

	return clrText;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::OnFillComboBoxItem(CDC* pDC, CBCGPComboBox* /*pComboBox*/, int /*nIndex*/, CRect rect, BOOL bIsHighlihted, BOOL bIsSelected)
{
	ASSERT_VALID (pDC);

	COLORREF clrText = (COLORREF)-1;

	if (bIsSelected)
	{
		pDC->FillRect (rect, &globalData.brHilite);
		clrText = globalData.clrTextHilite;
	}
	else
	{
		pDC->FillRect (rect, &globalData.brWindow);
		clrText = globalData.clrWindowText;
	}

	if (bIsHighlihted)
	{
		pDC->DrawFocusRect (rect);
	}

	return clrText;
}
//*******************************************************************************
HBRUSH CBCGPVisualManager::GetListControlFillBrush(CBCGPListCtrl* /*pListCtrl*/)
{
	return globalData.brWindow;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetListControlMarkedColor(CBCGPListCtrl* pListCtrl)
{
	CBrush* pBr = CBrush::FromHandle(GetListControlFillBrush(pListCtrl));
	
	LOGBRUSH lbr;
	pBr->GetLogBrush(&lbr);
	
	return CBCGPDrawManager::PixelAlpha(lbr.lbColor, .97, .97, .97);
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetListControlTextColor(CBCGPListCtrl* /*pListCtrl*/)
{
	return globalData.clrWindowText;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetTreeControlFillColor(CBCGPTreeCtrl* /*pTreeCtrl*/, BOOL /*bIsSelected*/, BOOL /*bIsFocused*/, BOOL /*bIsDisabled*/)
{
	// Use the -1 value to restore the system default colors.
	return (COLORREF)-1;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetTreeControlLineColor(CBCGPTreeCtrl* /*pTreeCtrl*/, BOOL /*bIsSelected*/, BOOL /*bIsFocused*/, BOOL /*bIsDisabled*/)
{
	// Use the CLR_DEFAULT value to restore the system default colors.
	return CLR_DEFAULT;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetTreeControlTextColor(CBCGPTreeCtrl* /*pTreeCtrl*/, BOOL /*bIsSelected*/, BOOL /*bIsFocused*/, BOOL /*bIsDisabled*/)
{
	return (COLORREF)-1;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetIntelliSenseFillColor(CBCGPBaseIntelliSenseLB* /*pCtrl*/, BOOL bIsSelected)
{
	return bIsSelected ? GetSysColor(COLOR_HIGHLIGHT) : GetSysColor(COLOR_WINDOW);
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetIntelliSenseTextColor(CBCGPBaseIntelliSenseLB* /*pTreeCtrl*/, BOOL bIsSelected)
{
	return bIsSelected ? GetSysColor(COLOR_HIGHLIGHTTEXT) : GetSysColor(COLOR_WINDOWTEXT);
}
//*******************************************************************************
void CBCGPVisualManager::OnScrollBarDrawThumb(CDC* pDC, CBCGPScrollBar* /*pScrollBar*/, CRect rect, BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bDisabled)
{
	ASSERT_VALID (pDC);
	
	if (m_hThemeScrollBar != NULL)
	{
		int nPartID = bHorz ? SBP_THUMBBTNHORZ : SBP_THUMBBTNVERT;
		
		int nStateID = SCRBS_NORMAL;
		if (bDisabled)
		{
			nStateID = SCRBS_DISABLED;
		}
		else if (bPressed)
		{
			nStateID = SCRBS_PRESSED;
		}
		else if (bHighlighted)
		{
			nStateID = SCRBS_HOT;
		}

		(*m_pfDrawThemeBackground)(m_hThemeScrollBar, pDC->GetSafeHdc(), nPartID, nStateID, &rect, 0);
	}
	else if (!bDisabled)
	{
		CBCGPDrawManager dm (*pDC);

		dm.DrawRect(rect, 
			bPressed ? globalData.clrBarShadow : globalData.clrBarFace, 
			bPressed || bHighlighted ? globalData.clrBarDkShadow : globalData.clrBarShadow);
	}
}
//*******************************************************************************
void CBCGPVisualManager::OnScrollBarDrawButton(CDC* pDC, CBCGPScrollBar* pScrollBar, CRect rect, BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bFirst, BOOL bDisabled)
{
	ASSERT_VALID (pDC);
	
	if (bHorz && pScrollBar->GetSafeHwnd() != NULL && (pScrollBar->GetExStyle() & WS_EX_LAYOUTRTL))
	{
		bFirst = !bFirst;
	}

	if (m_hThemeScrollBar != NULL)
	{
		int nPartID = SBP_ARROWBTN;

		int nStateID = ABS_UPNORMAL;
		if (bDisabled)
		{
			nStateID = ABS_UPDISABLED;
		}
		else if (bPressed)
		{
			nStateID = ABS_UPPRESSED;
		}
		else if (bHighlighted)
		{
			nStateID = ABS_UPHOT;
		}
		
		if (!bFirst)
		{
			nStateID += 4;
		}

		if (bHorz)
		{
			nStateID += 8;
		}
		
		(*m_pfDrawThemeBackground)(m_hThemeScrollBar, pDC->GetSafeHdc(), nPartID, nStateID, &rect, 0);
	}
	else
	{
		CBCGPDrawManager dm (*pDC);

		COLORREF clrFill = bPressed ? globalData.clrBarShadow : globalData.clrBarFace;
		dm.DrawRect(rect, clrFill, bPressed || bHighlighted ? globalData.clrBarDkShadow : globalData.clrBarShadow);

		if (!globalData.Is32BitIcons())
		{
			// Menu images are 16 colors - draw symbol instead:
			dm.DrawLine(rect.left, rect.top, rect.right, rect.bottom, globalData.clrBarText);
		}
		else
		{
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
				CBCGPMenuImages::Draw(pDC, ids, rect, CBCGPMenuImages::ImageLtGray);
			}
			else
			{
				CBCGPMenuImages::DrawByColor(pDC, ids, rect, clrFill);
			}
		}
	}
}
//*******************************************************************************
void CBCGPVisualManager::OnScrollBarFillBackground(CDC* pDC, CBCGPScrollBar* /*pScrollBar*/, CRect rect, BOOL bHorz, BOOL bHighlighted, BOOL bPressed, BOOL bFirst, BOOL bDisabled)
{
	ASSERT_VALID (pDC);

	if (m_hThemeScrollBar != NULL)
	{
		int nPartID = 0;
		int nStateID = SCRBS_NORMAL;
		if (bDisabled)
		{
			nStateID = SCRBS_DISABLED;
		}
		else if (bPressed)
		{
			nStateID = SCRBS_PRESSED;
		}
		else if (bHighlighted)
		{
			nStateID = SCRBS_HOT;
		}

		if (bHorz)
		{
			nPartID = bFirst ? SBP_LOWERTRACKHORZ : SBP_UPPERTRACKHORZ;
		}
		else
		{
			nPartID = bFirst ? SBP_LOWERTRACKVERT : SBP_UPPERTRACKVERT;
		}

		(*m_pfDrawThemeBackground)(m_hThemeScrollBar, pDC->GetSafeHdc(), nPartID, nStateID, &rect, 0);
	}
	else
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawRect(rect, 
			CBCGPDrawManager::IsPaleColor(globalData.clrBarFace) ?
				CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace) : CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace), 
			(COLORREF)-1);
	}
}
//*******************************************************************************
BOOL CBCGPVisualManager::OnDrawPushButton (CDC* pDC, CRect rect, CBCGPButton* pButton, COLORREF& clrText)
{
	ASSERT_VALID (pDC);
	
	BOOL bDisabled    = pButton->GetSafeHwnd() != NULL && !pButton->IsWindowEnabled ();
	BOOL bPressed     = pButton->GetSafeHwnd() != NULL && pButton->IsPressed ();
	BOOL bChecked     = pButton->GetSafeHwnd() != NULL && pButton->IsChecked ();
	BOOL bDefault	  = pButton->GetSafeHwnd() != NULL && pButton->IsDefaultButton();
	
	pDC->FillRect(rect, &globalData.brBarFace);

	if (bChecked)
	{
		CBCGPDrawManager dm (*pDC);
		dm.HighlightRect(rect);
	}

	if (bDefault)
	{
		pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarDkShadow);
		rect.DeflateRect(1, 1);
	}
	
	if (bPressed || bChecked)
	{
		pDC->Draw3dRect (rect, globalData.clrBarDkShadow, globalData.clrBarHilite);
		rect.DeflateRect(1, 1);
		pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarLight);
	}
	else
	{
		pDC->Draw3dRect (rect, globalData.clrBarLight, globalData.clrBarDkShadow);
		rect.DeflateRect(1, 1);
		pDC->Draw3dRect (rect, globalData.clrBarHilite, globalData.clrBarShadow);
	}

	clrText = bDisabled ? globalData.clrGrayedText : globalData.clrBarText;
	return TRUE;
}
//*******************************************************************************
int CBCGPVisualManager::GetPushButtonStdImageState(CBCGPButton* /*pButton*/, BOOL bIsDisabled)
{
	return (int)(bIsDisabled ? CBCGPMenuImages::ImageGray : CBCGPMenuImages::ImageBlack);
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawGroup (CDC* pDC, CBCGPGroup* pGroup, CRect rect, const CString& strName)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pGroup);

	CSize sizeText = pDC->GetTextExtent (strName);

	CRect rectFrame = rect;
	rectFrame.top += sizeText.cy / 2;

	if (sizeText == CSize (0, 0))
	{
		rectFrame.top += pDC->GetTextExtent (_T("A")).cy / 2;
	}

	int xMargin = sizeText.cy / 2;

	sizeText.cx = max(0, min(sizeText.cx, rect.Width() - 3 * xMargin));

	CRect rectText = rect;
	rectText.left += xMargin;
	rectText.right = rectText.left + sizeText.cx + xMargin;
	rectText.bottom = rectText.top + sizeText.cy;

	if (!strName.IsEmpty ())
	{
		pDC->ExcludeClipRect (rectText);
	}

	if (m_pfDrawThemeBackground != NULL && m_hThemeButton != NULL && !pGroup->m_bOnGlass)
	{
		(*m_pfDrawThemeBackground) (m_hThemeButton, pDC->GetSafeHdc(), BP_GROUPBOX, 
			pGroup->IsWindowEnabled () ? GBS_NORMAL : GBS_DISABLED, &rectFrame, 0);
	}
	else
	{
		CBCGPDrawManager dm (*pDC);
		dm.DrawRect (rectFrame, (COLORREF)-1, globalData.clrBarDkShadow);
	}

	if (strName.IsEmpty ())
	{
		return;
	}

	pDC->SelectClipRgn (NULL);

	DWORD dwTextStyle = DT_SINGLELINE | DT_VCENTER | DT_CENTER | DT_END_ELLIPSIS;

	if (pGroup->m_bOnGlass)
	{
		DrawTextOnGlass (pDC, strName, rectText, dwTextStyle, 10, globalData.clrBarText);
	}
	else
	{
		CString strCaption = strName;
		pDC->DrawText (strCaption, rectText, dwTextStyle);
	}
}
//*******************************************************************************
BOOL CBCGPVisualManager::OnFillDialog (CDC* pDC, CWnd* pDlg, CRect rect)
{
	ASSERT_VALID (pDC);
	pDC->FillRect (rect, &GetDlgBackBrush (pDlg));

	return TRUE;
}
//*******************************************************************************
void CBCGPVisualManager::OnFillRibbonBackstageForm(CDC* pDC, CWnd* /*pDlg*/, CRect rect)
{
	ASSERT_VALID (pDC);

	if (!globalData.IsHighContastMode())
	{
#ifndef BCGP_EXCLUDE_RIBBON
		if (IsRibbonBackstageWhiteBackground())
		{
			::FillRect(pDC->GetSafeHdc(), rect, (HBRUSH)::GetStockObject(WHITE_BRUSH));
		}
		else
#endif
		{
			pDC->FillRect(rect, &globalData.brBarFace);
		}
	}
	else
	{
		pDC->FillRect(rect, &globalData.brWindow);
	}
}
//*******************************************************************************
CBrush& CBCGPVisualManager::GetDlgBackBrush (CWnd* /*pDlg*/)
{
	return globalData.brBarFace;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetDlgTextColor(CWnd* /*pDlg*/)
{
	return globalData.clrBarText;
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawDlgSizeBox (CDC* pDC, CWnd* /*pDlg*/, CRect rectSizeBox)
{
	ASSERT_VALID(pDC);

	int nBKMode = pDC->SetBkMode(TRANSPARENT);
	OnDrawStatusBarSizeBox (pDC, NULL, rectSizeBox);
	pDC->SetBkMode(nBKMode);
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawButtonsArea(CDC* pDC, CWnd* pDlg, CRect rect)
{
	ASSERT_VALID (pDC);

	COLORREF clrLine;
	pDC->FillRect (rect, &GetDlgButtonsAreaBrush(pDlg, &clrLine));

	CPen pen (PS_SOLID, 1, clrLine);
	CPen* pOldPen = pDC->SelectObject (&pen);

	pDC->MoveTo(rect.left, rect.top);
	pDC->LineTo(rect.right, rect.top);

	pDC->SelectObject (pOldPen);
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawDlgExpandedArea(CDC* pDC, CWnd* pDlg, CRect rect)
{
	if (globalData.m_nBitsPerPixel <= 8 || globalData.IsHighContastMode ())
	{
		return;
	}

	ASSERT_VALID (pDC);

	COLORREF clrFill = (COLORREF)-1;
	COLORREF clrLine = (COLORREF)-1;

	CBrush& brFill = GetDlgBackBrush(pDlg);
	
	LOGBRUSH lbr;
	brFill.GetLogBrush(&lbr);

	if (IsDarkTheme())
	{
		clrFill = CBCGPDrawManager::ColorMakeLighter(lbr.lbColor);
		clrLine = CBCGPDrawManager::ColorMakeLighter(lbr.lbColor, .05);
	}
	else
	{
		clrLine = CBCGPDrawManager::ColorMakeDarker(lbr.lbColor);
		clrFill = CBCGPDrawManager::ColorMakeDarker(lbr.lbColor, .05);
	}

	pDC->FillSolidRect(rect, clrFill);

	CBCGPPenSelector ps(*pDC, clrLine);
	
	pDC->MoveTo(rect.left, rect.top);
	pDC->LineTo(rect.right, rect.top);
}
//*******************************************************************************
CBrush& CBCGPVisualManager::GetDlgButtonsAreaBrush(CWnd* /*pDlg*/, COLORREF* pclrLine)
{
	if (pclrLine != NULL)
	{
		*pclrLine = globalData.clrBarDkShadow;
	}

	return m_brDlgButtonsArea;
}
//*****************************************************************************
COLORREF CBCGPVisualManager::GetDlgHeaderTextColor(CWnd* /*pDlg*/)
{
	if (m_hThemeTextStyle != NULL && m_pfGetThemeColor != NULL)
	{
		COLORREF clrText = (COLORREF)-1;
		(*m_pfGetThemeColor)(m_hThemeTextStyle, TEXT_MAININSTRUCTION, TS_CONTROLLABEL_NORMAL, TMT_TEXTCOLOR, &clrText);

		if (clrText != (COLORREF)-1)
		{
			return clrText;
		}
	}

	return ::GetSysColor(COLOR_HIGHLIGHT);
}
//*****************************************************************************
CFont& CBCGPVisualManager::GetNcCaptionTextFont()
{
	return globalData.fontBold;
}
//*****************************************************************************
COLORREF CBCGPVisualManager::GetNcCaptionTextColor(BOOL /*bActive*/, BOOL /*bTitle*/) const
{
	return globalData.clrWindowText;
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawSliderChannel (CDC* pDC, CBCGPSliderCtrl* pSlider, BOOL bVert, CRect rect, BOOL bDrawOnGlass)
{
	ASSERT_VALID (pDC);

	if (bVert)
	{
		if (rect.Width () < 3)
		{
			rect.right++;
		}
	}
	else
	{
		if (rect.Height () < 3)
		{
			rect.bottom++;
		}
	}

	if (m_hThemeTrack != NULL && globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode ())
	{
		(*m_pfDrawThemeBackground) (m_hThemeTrack, pDC->GetSafeHdc(), 
			bVert ? TKP_TRACKVERT : TKP_TRACK,
			1, &rect, 0);
	}
	else
	{
		if (bDrawOnGlass)
		{
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect (rect, (COLORREF)-1, globalData.clrBarShadow);
		}
		else
		{
			pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarHilite);
		}
	}

	if (pSlider->GetSafeHwnd() != NULL && (pSlider->GetStyle() & TBS_ENABLESELRANGE))
	{
		CRect rectSel = pSlider->GetSelectionRect();
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

			rectSel.left = max(rectSel.left, rect.left + 1);
			rectSel.right = min(rectSel.right, rect.right - 1);

			COLORREF clrSelection = pSlider->IsWindowEnabled() ? pSlider->GetSelectionColor() : globalData.clrBarLight;
			if (clrSelection == (COLORREF)-1)
			{
				clrSelection = globalData.clrHilite;
			}
			
			CBCGPDrawManager dm (*pDC);
			dm.DrawRect(rectSel, clrSelection, (COLORREF)-1);
		}
	}
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawSliderThumb (CDC* pDC, CBCGPSliderCtrl* /*pSlider*/,
			CRect rect, BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled,
			BOOL bVert, BOOL bLeftTop, BOOL bRightBottom, BOOL /*bDrawOnGlass*/)
{
	ASSERT_VALID (pDC);

	if (m_hThemeTrack != NULL && globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode ())
	{
		int nPart = 0;
		int nState = 0;

		if (bLeftTop && bRightBottom)
		{
            nPart = bVert ? TKP_THUMBVERT : TKP_THUMB;
		}
  		else if (bLeftTop)
		{
			nPart = bVert ? TKP_THUMBLEFT : TKP_THUMBTOP;
		}
		else
		{
			nPart = bVert ? TKP_THUMBRIGHT : TKP_THUMBBOTTOM;
		}

		if (bIsDisabled)
		{
            nState = TUS_DISABLED;
		}
        else if (bIsPressed)
		{
            nState = TUS_PRESSED;
		}
        else if (bIsHighlighted)
		{
            nState = TUS_HOT;
		}
        else
		{
            nState = TUS_NORMAL;
		}

      	(*m_pfDrawThemeBackground) (m_hThemeTrack, pDC->GetSafeHdc(), nPart, nState, &rect, 0);
		return;
	}

	COLORREF clrLine = bIsDisabled ? globalData.clrBarShadow : globalData.clrBarDkShadow;
	COLORREF clrFill = !bIsDisabled && (bIsHighlighted || bIsPressed) ?
		globalData.clrBarHilite : globalData.clrBarFace;

	if (bVert)
	{
		rect.DeflateRect (2, 1);

		rect.left = rect.CenterPoint ().x - rect.Height ();
		rect.right = rect.left + 2 * rect.Height ();
	}
	else
	{
		rect.DeflateRect (1, 2);

		rect.top = rect.CenterPoint ().y - rect.Width ();
		rect.bottom = rect.top + 2 * rect.Width ();
	}

	CBCGPDrawManager dm (*pDC);
	dm.DrawRect (rect, clrFill, clrLine);
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawSliderTic(CDC* pDC, CBCGPSliderCtrl* pSlider, CRect rectTic, BOOL /*bVert*/, BOOL /*bLeftTop*/, BOOL bDrawOnGlass)
{
	ASSERT_VALID (pDC);

	BOOL bIsDisabled = pSlider->GetSafeHwnd() != NULL && !pSlider->IsWindowEnabled();
	
	COLORREF clrLine = 
		bDrawOnGlass ? (bIsDisabled ? globalData.clrBtnShadow : globalData.clrBtnDkShadow) : 
		bIsDisabled ? globalData.clrGrayedText : m_clrSliderTic;
	
	CBCGPDrawManager dm (*pDC);
	dm.DrawRect (rectTic, (COLORREF)-1, clrLine);
}
//*******************************************************************************
void CBCGPVisualManager::OnDrawSliderSelectionMarker(CDC* pDC, CBCGPSliderCtrl* /*pSlider*/, CRect rectMarker, BOOL bStart, BOOL /*bVert*/, BOOL bLeftTop, BOOL bDrawOnGlass)
{
	ASSERT_VALID (pDC);

	if (bDrawOnGlass)
	{
		CBCGPDrawManager dm(*pDC);
		dm.DrawRect (rectMarker, m_clrSliderSelMark, (COLORREF)-1);

		return;
	}

	POINT pt[3];

	if (bLeftTop)
	{
		if (bStart)
		{
			pt[0].x = rectMarker.left;
			pt[0].y = rectMarker.top;
			
			pt[1].x = rectMarker.right;
			pt[1].y = rectMarker.bottom;
			
			pt[2].x = rectMarker.right;
			pt[2].y = rectMarker.top;
		}
		else
		{
			pt[0].x = rectMarker.left;
			pt[0].y = rectMarker.top;
			
			pt[1].x = rectMarker.left;
			pt[1].y = rectMarker.bottom;
			
			pt[2].x = rectMarker.right;
			pt[2].y = rectMarker.top;
		}
	}
	else
	{
		if (bStart)
		{
			pt[0].x = rectMarker.right;
			pt[0].y = rectMarker.top - 1;
			
			pt[1].x = rectMarker.right;
			pt[1].y = rectMarker.bottom;

			pt[2].x = rectMarker.left - 1;
			pt[2].y = rectMarker.bottom;
		}
		else
		{
			pt[0].x = rectMarker.left;
			pt[0].y = rectMarker.top - 1;
			
			pt[1].x = rectMarker.right + 1;
			pt[1].y = rectMarker.bottom;

			pt[2].x = rectMarker.left;
			pt[2].y = rectMarker.bottom;
		}
	}

	CPen* pOldPen = (CPen*) pDC->SelectStockObject (NULL_PEN);
	CBCGPBrushSelector brush(*pDC, m_clrSliderSelMark);

	pDC->Polygon(pt, 3);

	pDC->SelectObject(pOldPen);
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetURLLinkColor (CBCGPURLLinkButton* pButton, BOOL bHover)
{
	if (pButton->GetSafeHwnd() != NULL && !pButton->IsWindowEnabled())
	{
		return globalData.clrGrayedText;
	}

	return bHover ? globalData.clrHotLinkText : globalData.clrHotText;
}
//****************************************************************************************
BOOL CBCGPVisualManager::OnFillParentBarBackground (CWnd* pWnd, CDC* pDC, LPRECT rectClip, BOOL bNCArea)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pWnd);

	CWnd* pParent = pWnd->GetParent ();
	ASSERT_VALID (pParent);

	CBCGPBaseControlBar* pParentBar = DYNAMIC_DOWNCAST(CBCGPBaseControlBar, pParent);
	if (pParentBar == NULL)
	{
		return FALSE;
	}

	CBCGPDialogBar* pDialogBar = DYNAMIC_DOWNCAST (CBCGPDialogBar, pParentBar);
	if (pDialogBar != NULL && pDialogBar->IsVisualManagerStyle ())
	{
		return FALSE;
	}

	CRgn rgn;
	if (rectClip != NULL)
	{
		rgn.CreateRectRgnIndirect (rectClip);
		pDC->SelectClipRgn (&rgn);
	}

	CPoint pt (0, 0);
	
	if (bNCArea)
	{
		CRect rectWnd;
		pWnd->GetWindowRect(rectWnd);
		pWnd->ScreenToClient(rectWnd);
		
		pt = rectWnd.TopLeft();
	}
	
	pWnd->MapWindowPoints (pParent, &pt, 1);

	pt = pDC->OffsetWindowOrg (pt.x, pt.y);

	CRect rectBar;
	pParentBar->GetClientRect (rectBar);

	OnFillBarBackground (pDC, pParentBar, rectBar, rectBar);

	pDC->SetWindowOrg (pt.x, pt.y);

	pDC->SelectClipRgn (NULL);

	return TRUE;
}

//****************************************************************************************
// Breadcrumb control default rendering
//****************************************************************************************

COLORREF CBCGPVisualManager::BreadcrumbFillBackground (CDC& dc, CBCGPBreadcrumb* pControl, CRect rectFill)
{
	COLORREF clrBackground = pControl->GetBackColor();
	if (clrBackground != CLR_INVALID)
	{
		CBrush br(clrBackground);
		dc.FillRect(&rectFill, &br);
	}
	else
	{
		dc.FillRect(&rectFill, &globalData.brWindow);
	}

	return pControl->IsWindowEnabled() ? globalData.clrWindowText : globalData.clrGrayedText;
}
//****************************************************************************************
void CBCGPVisualManager::BreadcrumbFillProgress (CDC& dc, CBCGPBreadcrumb* /*pControl*/, CRect rectProgress, CRect rectChunk, int /*nValue*/)
{
	if (m_hThemeProgress == NULL)
	{
		dc.FillRect (rectChunk, &globalData.brHilite);
		return;
	}

	UINT nPart = PP_CHUNK;
	UINT nState = 0;
	
	if (globalData.bIsWindowsVista)
	{
		nPart = PP_FILL;
		nState = PBFS_NORMAL;
	}
	else
	{
		rectChunk.DeflateRect(0, 1);
	}

	(*m_pfDrawThemeBackground) (m_hThemeProgress, dc.GetSafeHdc(), PP_CHUNK, nState, &rectChunk, &rectProgress);
}
//****************************************************************************************
void CBCGPVisualManager::BreadcrumbDrawItemBackground (CDC& dc, CBCGPBreadcrumb* /*pControl*/, BREADCRUMBITEMINFO* /*pItemInfo*/, CRect rectItem, UINT uState, COLORREF color)
{
	if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
	{	
		rectItem.right --;
		
		if (color == CLR_INVALID)
		{
			color = GetSysColor(COLOR_HIGHLIGHT);
		}

		dc.FillSolidRect(&rectItem, color);
	}
}
//****************************************************************************************
void CBCGPVisualManager::BreadcrumbDrawItem (CDC& dc, CBCGPBreadcrumb* pControl, BREADCRUMBITEMINFO* pItemInfo, CRect rectItem, UINT uState, COLORREF color)
{
	rectItem.InflateRect(-2, 0);

	if (uState == CDIS_SELECTED)
	{
		rectItem.left ++;
		rectItem.top += 2;
	}

	if (!pControl->IsWindowEnabled())
	{
		color = globalData.clrGrayedText;
	}

	if (color == CLR_INVALID)
	{
		if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
		{
			color = GetSysColor(COLOR_HIGHLIGHTTEXT);
		}
		else
		{
			color = GetSysColor(COLOR_WINDOWTEXT);
		}
	}

	dc.SetTextColor (color);
	dc.DrawText(pItemInfo->pszText, ::lstrlen(pItemInfo->pszText), &rectItem,
		DT_END_ELLIPSIS | DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
}
//****************************************************************************************
void CBCGPVisualManager::BreadcrumbDrawArrowBackground (CDC& dc, CBCGPBreadcrumb* /*pControl*/, BREADCRUMBITEMINFO* /*pItemInfo*/, CRect rectArrow, UINT uState, COLORREF color)
{
	if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
	{	
		if (color == CLR_INVALID)
		{
			color = GetSysColor(COLOR_HIGHLIGHT);
		}

		// Provide a dividing line between item parts
		rectArrow.left ++;

		dc.FillSolidRect(&rectArrow, color);
	}
}
//****************************************************************************************
void CBCGPVisualManager::BreadcrumbDrawArrow (CDC& dc, CBCGPBreadcrumb* /*pControl*/, BREADCRUMBITEMINFO* /*pItemInfo*/, CRect rect, UINT uState, COLORREF color)
{
	POINT ptC;
	ptC.x = (rect.left + rect.right) / 2 + 1;
	ptC.y = (rect.top + rect.bottom) / 2;

	int delta = globalUtils.ScaleByDPI(3);

	POINT pt[3];
	if (uState == CDIS_SELECTED)
	{
		pt[0].x = ptC.x - delta;
		pt[0].y = ptC.y;
		pt[1].x = ptC.x + delta;
		pt[1].y = ptC.y;
		pt[2].x = ptC.x;
		pt[2].y = ptC.y + delta;
	}
	else
	{
		pt[0].x = ptC.x - delta + 1;
		pt[0].y = ptC.y - delta;
		pt[1].x = ptC.x - delta + 1;
		pt[1].y = ptC.y + delta;
		pt[2].x = ptC.x + 1;
		pt[2].y = ptC.y;		
	}

	if (color == CLR_INVALID)
	{
		if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
		{
			color = GetSysColor(COLOR_HIGHLIGHTTEXT);
		}
		else
		{
			color = GetSysColor(COLOR_WINDOWTEXT);
		}
	}

	CBCGPPenSelector pen (dc, color);
	CBCGPBrushSelector brush (dc, color);
	dc.Polygon (pt, 3);
}
//****************************************************************************************
void CBCGPVisualManager::BreadcrumbDrawLeftArrowBackground (CDC& dc, CBCGPBreadcrumb* /*pControl*/, CRect rect, UINT uState, COLORREF color)
{
	if (uState == CDIS_HOT || uState == CDIS_SELECTED)
	{
		if (color == CLR_INVALID)
		{
			color = GetSysColor(COLOR_HIGHLIGHT);
		}

		dc.FillSolidRect (&rect, color);
	}
}
//****************************************************************************************
static void DrawBracketHelper (CDC& dc, int x, int y)
{
	dc.MoveTo (x, y - 2);
	dc.LineTo (x - 2, y);
	dc.LineTo (x + 1, y + 3);
}
//****************************************************************************************
void CBCGPVisualManager::BreadcrumbDrawLeftArrow (CDC& dc, CBCGPBreadcrumb* /*pControl*/, CRect rect, UINT uState, COLORREF color)
{
	POINT ptC;
	ptC.x = (rect.left + rect.right) / 2;
	ptC.y = (rect.top + rect.bottom) / 2;

	if (uState == CDIS_SELECTED)
	{
		ptC.x ++;
		ptC.y ++;
	}

	if (color == CLR_INVALID)
	{
		if (uState == CDIS_HOT || uState == CDIS_SELECTED || uState == CDIS_OTHERSIDEHOT)
		{
			color = GetSysColor(COLOR_HIGHLIGHTTEXT);
		}
		else
		{
			color = GetSysColor(COLOR_WINDOWTEXT);
		}
	}

	CBCGPPenSelector pen (dc, color);
	DrawBracketHelper (dc, ptC.x - 2, ptC.y);
	DrawBracketHelper (dc, ptC.x - 1, ptC.y);
	DrawBracketHelper (dc, ptC.x + 2, ptC.y);
	DrawBracketHelper (dc, ptC.x + 3, ptC.y);
}
//************************************************************************************
COLORREF CBCGPVisualManager::GetEditCtrlSelectionBkColor(CBCGPEditCtrl* /*pEdit*/, BOOL bIsFocused)
{
	return bIsFocused ? m_clrEditCtrlSelectionBkActive : m_clrEditCtrlSelectionBkInactive;
}
//************************************************************************************
COLORREF CBCGPVisualManager::GetEditCtrlSelectionTextColor(CBCGPEditCtrl* /*pEdit*/, BOOL bIsFocused)
{
	return bIsFocused ? m_clrEditCtrlSelectionTextActive : m_clrEditCtrlSelectionTextInactive;
}
//************************************************************************************
CSize CBCGPVisualManager::GetPinSize(BOOL /*bIsPinned*/)
{
	return globalUtils.ScaleByDPI(CBCGPMenuImages::Size());
}
//************************************************************************************
void CBCGPVisualManager::OnDrawPin(CDC* pDC, const CRect& rect, BOOL bIsPinned, BOOL bIsDark, BOOL /*bIsHighlighted*/, BOOL /*bIsPressed*/, BOOL bIsDisabled)
{
	CSize sizeDest(0, 0);

	if (globalData.GetRibbonImageScale() != 1.)
	{
		sizeDest = GetPinSize(bIsPinned);
	}

	CBCGPMenuImages::Draw (pDC, bIsPinned ? CBCGPMenuImages::IdPinVert : CBCGPMenuImages::IdPinHorz, 
		rect, bIsDisabled ? (IsDarkTheme() ? CBCGPMenuImages::ImageGray : CBCGPMenuImages::ImageLtGray) : (bIsDark ? CBCGPMenuImages::ImageBlack : CBCGPMenuImages::ImageWhite),
		sizeDest);
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetPrintPreviewBackgroundColor(CBCGPPrintPreviewView* /*pPrintPreview*/)
{
	return globalData.clrBarFace;
}
//*******************************************************************************
COLORREF CBCGPVisualManager::GetPrintPreviewFrameColor(BOOL bIsActive)
{
	return bIsActive ? globalData.clrHilite : globalData.clrBarDkShadow;
}

#ifndef BCGP_EXCLUDE_RIBBON

CBCGPRibbonBar*	CBCGPVisualManager::GetRibbonBar (CWnd* pWnd) const
{
	CBCGPRibbonBar* pBar = NULL;

	if (pWnd == NULL)
	{
		pWnd = AfxGetMainWnd ();
	}

	if (pWnd->GetSafeHwnd () == NULL)
	{
		return NULL;
	}

	if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPFrameWnd)))
	{
		pBar = ((CBCGPFrameWnd*) pWnd)->GetRibbonBar ();
	}
	else if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPMDIFrameWnd)))
	{
		pBar = ((CBCGPMDIFrameWnd*) pWnd)->GetRibbonBar ();
	}

	return pBar;
}
//*****************************************************************************
BOOL CBCGPVisualManager::IsRibbonPresent (CWnd* pWnd) const
{
	CBCGPRibbonBar* pBar = GetRibbonBar (pWnd);

	return pBar != NULL && pBar->IsWindowVisible ();
}

#endif

UINT CBCGPVisualManager::GetNavButtonsID(BOOL bIsLarge)
{
	return bIsLarge ? IDB_BCGBARRES_NAV_BUTTONS_120 : IDB_BCGBARRES_NAV_BUTTONS;
}
//*****************************************************************************
UINT CBCGPVisualManager::GetNavFrameID(BOOL bIsLarge)
{
	return bIsLarge ? IDB_BCGBARRES_NAV_FRAMES_120 : IDB_BCGBARRES_NAV_FRAMES;
}
//*****************************************************************************
int CBCGPVisualManager::GetNavImageState(BOOL bIsHighlighted, BOOL bIsPressed, BOOL bIsDisabled, BOOL bIsDrawOnGlass)
{
	UNREFERENCED_PARAMETER(bIsPressed);
	UNREFERENCED_PARAMETER(bIsDrawOnGlass);

	CBCGPMenuImages::IMAGE_STATE stateImage = CBCGPMenuImages::ImageBlack;
	
	if (bIsDisabled)
	{
		stateImage = CBCGPVisualManager::GetInstance()->IsDarkTheme() ? CBCGPMenuImages::ImageDkGray : CBCGPMenuImages::ImageGray;
	}
	else if (bIsHighlighted)
	{
		stateImage = CBCGPMenuImages::ImageBlack2;
	}
	
	return (int)stateImage;
}
//*****************************************************************************
void CBCGPVisualManager::OnDrawProgressMarqueeDot(CDC* pDC, CBCGPProgressCtrl* pProgressCtrl, CRect rect, int /*nIndex*/)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pProgressCtrl);

	COLORREF clr = pProgressCtrl->m_clrMarquee == (COLORREF)-1 ?
		(pProgressCtrl->m_bVisualManagerStyle ? globalData.clrHilite : GetSysColor(COLOR_HIGHLIGHT)) :
		pProgressCtrl->m_clrMarquee;

	if (pProgressCtrl->m_bOnGlass)
	{
		CBCGPDrawManager dm(*pDC);
		dm.DrawRect(rect, clr, (COLORREF)-1);
	}
	else
	{
		pDC->FillSolidRect(rect, clr);
	}
}
//*****************************************************************************
int CBCGPVisualManager::GetToolBarCustomizeButtonMargin() const
{
	return globalUtils.ScaleByDPI(2);
}
//*****************************************************************************
void CBCGPVisualManager::GetChartColors(CBCGPChartColors& colors)
{
	colors.m_clrFill = globalData.clrWindow;
	colors.m_clrOutline = globalData.clrWindowText;
}
//*****************************************************************************
void CBCGPVisualManager::GetCircularGaugeColors(CBCGPCircularGaugeColors& colors)
{
	COLORREF clrText = CBCGPDrawManager::ColorMakeLighter(globalData.clrBarText, .5);
	COLORREF clrOutline = CBCGPDrawManager::ColorMakeDarker(globalData.clrBarShadow, .2);

	colors.m_brFill.SetColors(CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace), CBCGPDrawManager::ColorMakePale(globalData.clrBarFace), CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT);
	colors.m_brText.SetColor(clrText);

	colors.m_brPointerOutline.SetColor(globalData.clrHilite);
	colors.m_brPointerFill.SetColor(globalData.clrHilite);

	colors.m_brFrameOutline.SetColor(clrOutline);
	colors.m_brFrameFill.SetColors(globalData.clrBarShadow, globalData.clrBarHilite, CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT);

	colors.m_brCapOutline.SetColor(clrOutline);
	colors.m_brCapFill.SetColors(CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace), CBCGPDrawManager::ColorMakePale(globalData.clrBarFace), CBCGPBrush::BCGP_GRADIENT_RADIAL_TOP);

	colors.m_brTickMarkFill.SetColors(globalData.clrBarFace, globalData.clrBarHilite, CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT);
	colors.m_brTickMarkOutline.SetColor(clrOutline);
}
//*****************************************************************************
void CBCGPVisualManager::GetSwitchColors(CBCGPSwitchColors& colors, BOOL bIsVertical)
{
	UNREFERENCED_PARAMETER(bIsVertical);

	if (colors.IsGDICtrl())
	{
		colors.m_brFillOn = ::GetSysColor(COLOR_HIGHLIGHT);
		colors.m_brFillOff = globalData.IsHighContastMode() ? globalData.clrWindow : globalData.clrBarHilite;
		colors.m_brOutline = globalData.IsHighContastMode() ? globalData.clrWindowText : globalData.clrBarDkShadow;

		return;
	}

	colors.m_brOutline = globalData.clrBarShadow;
	colors.m_brOutlineThumb = globalData.clrBarDkShadow;
	colors.m_brFillThumb = globalData.clrBarShadow;

	colors.m_brFillOn = globalData.clrHilite;
	colors.m_brFillOff = globalData.clrBarFace;

	colors.m_brLabelOn = CBCGPDrawManager::GetContrastColor((COLORREF)colors.m_brFillOn.GetColor());
	colors.m_brLabelOff = CBCGPDrawManager::GetContrastColor((COLORREF)colors.m_brFillOff.GetColor());
	
	colors.m_brFill = globalData.clrBarLight;
	colors.m_brFocus = globalData.clrBarLight;
}
//*****************************************************************************
void CBCGPVisualManager::GetLinearGaugeColors(CBCGPLinearGaugeColors& colors)
{
	COLORREF clrText = CBCGPDrawManager::ColorMakeLighter(globalData.clrBarText, .5);
	COLORREF clrOutline = CBCGPDrawManager::ColorMakeDarker(globalData.clrBarShadow, .2);

	colors.m_brFill.SetColors(CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace), CBCGPDrawManager::ColorMakePale(globalData.clrBarFace), CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT);
	colors.m_brText.SetColor(clrText);
	
	colors.m_brPointerOutline.SetColor(globalData.clrHilite);
	colors.m_brPointerFill.SetColor(globalData.clrHilite);
	
	colors.m_brFrameOutline.SetColor(clrOutline);
	colors.m_brFrameFill.SetColors(globalData.clrBarShadow, globalData.clrBarHilite, CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT);
	
	colors.m_brTickMarkFill.SetColors(globalData.clrBarFace, globalData.clrBarHilite, CBCGPBrush::BCGP_GRADIENT_DIAGONAL_LEFT);
	colors.m_brTickMarkOutline.SetColor(clrOutline);
}
//*****************************************************************************
void CBCGPVisualManager::GetWinUITilesColors(CBCGPWinUITilesColors& colors)
{
	colors.m_brFill = globalData.IsHighContastMode() ? globalData.clrWindow : globalData.clrBarLight;
	colors.m_brBorder = globalData.clrBarShadow;
	colors.m_colorTileBorder = globalData.clrBarShadow;
	colors.m_brTileFill = globalData.clrBarFace;
	colors.m_colorTileBorderHighlighted = globalData.clrHilite;
	colors.m_colorSelectedTileBorder = globalData.clrBarShadow;
	colors.m_brTileFillHighlighted = globalData.clrBarFace;
	colors.m_colorTileText = globalData.clrBarText;
#ifndef BCGP_EXCLUDE_RIBBON
	colors.m_colorCaptionText = GetRibbonBackstageInfoTextColor();
#else
	colors.m_colorCaptionText = globalData.clrBarText;
#endif
}

void CBCGPVisualManager::GetInfoBoxColors(CBCGPInfoBoxRenderer* pCtrl, COLORREF& clrBackground, COLORREF& clrForeground, COLORREF& clrFrame, COLORREF& clrLink)
{
	UNREFERENCED_PARAMETER(pCtrl);
	
	if (clrBackground == CLR_DEFAULT)
	{
		clrBackground = globalData.IsHighContastMode() ? ::GetSysColor(COLOR_INFOBK) : RGB(252, 247, 182);
	}

	if (clrFrame == CLR_DEFAULT)
	{
		clrFrame = globalData.IsHighContastMode() ? globalData.clrBarShadow : RGB(220, 220, 220);
	}

	if (clrForeground == CLR_DEFAULT)
	{
		clrForeground = globalData.IsHighContastMode() ? ::GetSysColor(COLOR_INFOTEXT) : RGB(0, 0, 0);
	}

	if (clrLink == CLR_DEFAULT)
	{
		clrLink = globalData.clrHotText;
	}
}
//*****************************************************************************
void CBCGPVisualManager::OnDrawLightBoxDialogCloseButton(CDC* pDC, CBCGPLightBoxDialog* /*pDlg*/, CRect rectClose, BOOL bIsHighlighted, BOOL bIsPressed)
{
	ASSERT_VALID(pDC);	
	
	rectClose.DeflateRect(2, 2);
	
	if (bIsPressed)
	{
		pDC->FillSolidRect(rectClose, globalData.clrBtnShadow);
	}
	else if (bIsHighlighted)
	{
		pDC->Draw3dRect(rectClose, globalData.clrBtnShadow, globalData.clrBtnShadow);
	}

	CBCGPMenuImages::Draw(pDC, CBCGPMenuImages::IdClose, rectClose, bIsHighlighted || bIsPressed ? CBCGPMenuImages::ImageBlack2 : CBCGPMenuImages::ImageBlack);
}
//*****************************************************************************
void CBCGPVisualManager::RedrawMDIArea(CMDIFrameWnd* pFrameWnd)
{
	ASSERT_VALID(pFrameWnd);

	CWnd* pMDIClient = CWnd::FromHandle(pFrameWnd->m_hWndMDIClient);
	ASSERT_VALID(pMDIClient);

	if (m_bMDIAreaFullRedrawOnActivate)
	{
		pMDIClient->RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_ALLCHILDREN);
	}
	else
	{
		for (CWnd* pWndChild = pMDIClient->GetWindow(GW_CHILD); pWndChild != NULL; pWndChild = pWndChild->GetNextWindow())
		{
			if (DYNAMIC_DOWNCAST(CMDIChildWnd, pWndChild) != NULL)
			{
				if (!pWndChild->IsZoomed() && pWndChild->IsWindowVisible())
				{
					pWndChild->RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOCHILDREN);
				}
			}
		}
	}
}

/////////////////////////////////////////////////////////////////////////////////////
// CBCGPWinXPThemeManager

CBCGPWinXPThemeManager::CBCGPWinXPThemeManager ()
{
	m_hThemeWindow = NULL;
	m_hThemeToolBar = NULL;
	m_hThemeButton = NULL;
	m_hThemeStatusBar = NULL;
	m_hThemeRebar = NULL;
	m_hThemeComboBox = NULL;
	m_hThemeProgress = NULL;
	m_hThemeHeader = NULL;
	m_hThemeScrollBar = NULL;
	m_hThemeExplorerBar = NULL;
	m_hThemeTree = NULL;
	m_hThemeStartPanel = NULL;
	m_hThemeTaskBand = NULL;
	m_hThemeTaskBar = NULL;
	m_hThemeTaskDialog = NULL;
	m_hThemeTextStyle = NULL;
	m_hThemeSpin = NULL;
	m_hThemeTab = NULL;
	m_hThemeToolTip = NULL;
	m_hThemeTrack = NULL;
	m_hThemeMenu = NULL;
	m_hThemeNavigation = NULL;

	m_pRrenderSysCaptionButton = NULL;
	m_bScaleCheckRadio = FALSE;
	m_dblScaleCheckRadio = 1.0;

	m_hinstUXDLL = LoadLibrary (_T("UxTheme.dll"));

	if (m_hinstUXDLL != NULL)
	{
		m_pfOpenThemeData = (OPENTHEMEDATA)::GetProcAddress (m_hinstUXDLL, "OpenThemeData");
		m_pfCloseThemeData = (CLOSETHEMEDATA)::GetProcAddress (m_hinstUXDLL, "CloseThemeData");
		m_pfDrawThemeBackground = (DRAWTHEMEBACKGROUND)::GetProcAddress (m_hinstUXDLL, "DrawThemeBackground");
		m_pfGetThemeColor = (GETTHEMECOLOR)::GetProcAddress (m_hinstUXDLL, "GetThemeColor");
		m_pfGetThemeSysColor = (GETTHEMESYSCOLOR)::GetProcAddress (m_hinstUXDLL, "GetThemeSysColor");
		m_pfGetCurrentThemeName = (GETCURRENTTHEMENAME)::GetProcAddress (m_hinstUXDLL, "GetCurrentThemeName");
		m_pfGetWindowTheme = (GETWINDOWTHEME)::GetProcAddress (m_hinstUXDLL, "GetWindowTheme");
		m_pfGetThemePartSize = (GETTHEMEPARTSIZE)::GetProcAddress (m_hinstUXDLL, "GetThemePartSize");

		UpdateSystemColors ();
	}
	else
	{
		m_pfOpenThemeData = NULL;
		m_pfCloseThemeData = NULL;
		m_pfDrawThemeBackground = NULL;
		m_pfGetThemeColor = NULL;
		m_pfGetThemeSysColor = NULL;
		m_pfGetCurrentThemeName = NULL;
		m_pfGetWindowTheme = NULL;
		m_pfGetThemePartSize = NULL;
	}
}
//*************************************************************************************
CBCGPWinXPThemeManager::~CBCGPWinXPThemeManager ()
{
	if (m_hinstUXDLL != NULL)
	{
		CleanUpThemes ();
		::FreeLibrary (m_hinstUXDLL);
	}

	if (m_pRrenderSysCaptionButton != NULL)
	{
		delete m_pRrenderSysCaptionButton;
		m_pRrenderSysCaptionButton = NULL;
	}
}
//*************************************************************************************
void CBCGPWinXPThemeManager::UpdateSystemColors ()
{
	if (m_hinstUXDLL != NULL)
	{
		CleanUpThemes ();

		if (m_pfOpenThemeData == NULL ||
			m_pfCloseThemeData == NULL ||
			m_pfDrawThemeBackground == NULL)
		{
			ASSERT (FALSE);
		}
		else
		{
			m_hThemeWindow = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"WINDOW");
			m_hThemeToolBar = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"TOOLBAR");
			m_hThemeButton = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"BUTTON");
			m_hThemeStatusBar = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"STATUS");
			m_hThemeRebar = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"REBAR");
			m_hThemeComboBox = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"COMBOBOX");
			m_hThemeProgress = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"PROGRESS");
			m_hThemeHeader = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"HEADER");
			m_hThemeScrollBar = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"SCROLLBAR");
			m_hThemeExplorerBar = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"EXPLORERBAR");
			m_hThemeTree = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"TREEVIEW");
			m_hThemeStartPanel = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"STARTPANEL");
			m_hThemeTaskBand = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"TASKBAND");
			m_hThemeTaskBar = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"TASKBAR");
			m_hThemeTaskDialog = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"TASKDIALOG");
			m_hThemeTextStyle = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"TEXTSTYLE");
			m_hThemeSpin = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"SPIN");
			m_hThemeTab = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"TAB");
			m_hThemeToolTip = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"TOOLTIP");
			m_hThemeTrack = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"TRACKBAR");
			m_hThemeMenu = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"MENU");
			m_hThemeNavigation = (*m_pfOpenThemeData)(AfxGetMainWnd ()->GetSafeHwnd (), L"NAVIGATION");
		}
	}

	if (m_pRrenderSysCaptionButton != NULL)
	{
		delete m_pRrenderSysCaptionButton;
		m_pRrenderSysCaptionButton = NULL;
	}
}
//************************************************************************************
void CBCGPWinXPThemeManager::CleanUpTheme (HTHEME& hTheme)
{
	if (hTheme != NULL)
	{
		(*m_pfCloseThemeData) (hTheme);
		hTheme = NULL;
	}
}
//************************************************************************************
void CBCGPWinXPThemeManager::CleanUpThemes ()
{
	if (m_pfCloseThemeData == NULL)
	{
		return;
	}

	CleanUpTheme (m_hThemeWindow);
	CleanUpTheme (m_hThemeToolBar);
	CleanUpTheme (m_hThemeRebar);
	CleanUpTheme (m_hThemeStatusBar);
	CleanUpTheme (m_hThemeButton);
	CleanUpTheme (m_hThemeComboBox);
	CleanUpTheme (m_hThemeProgress);
	CleanUpTheme (m_hThemeHeader);
	CleanUpTheme (m_hThemeScrollBar);
	CleanUpTheme (m_hThemeExplorerBar);
	CleanUpTheme (m_hThemeTree);
	CleanUpTheme (m_hThemeStartPanel);
	CleanUpTheme (m_hThemeTaskBand);
	CleanUpTheme (m_hThemeTaskBar);
	CleanUpTheme (m_hThemeTaskDialog);
	CleanUpTheme (m_hThemeTextStyle);
	CleanUpTheme (m_hThemeSpin);
	CleanUpTheme (m_hThemeTab);
	CleanUpTheme (m_hThemeToolTip);
	CleanUpTheme (m_hThemeTrack);
	CleanUpTheme (m_hThemeMenu);
	CleanUpTheme (m_hThemeNavigation);
}
//**********************************************************************************
BOOL CBCGPWinXPThemeManager::DrawThemeCaptionButton (CDC* pDC, CRect rect, BOOL bDisabled, BOOL bIsPressed, BOOL bIsHighlighted, BOOL bIsActiveFrame)
{
	if (m_pRrenderSysCaptionButton == NULL)
	{
		m_pRrenderSysCaptionButton = new CBCGPControlRenderer;

		if (m_hThemeWindow == NULL)
		{
			return FALSE;
		}

		HBITMAP hBmp = globalData.GetThemeBitmap(m_hThemeWindow, WP_MAXBUTTON, 0, TMT_DIBDATA);
		if (hBmp == NULL)
		{
			return FALSE;
		}

		BITMAP bitmap;
		::GetObject(hBmp, sizeof(BITMAP), (LPVOID)&bitmap);

		int cxPartWidth = bitmap.bmWidth;
		int cyPartHeight = bitmap.bmHeight / 8;

		if (cxPartWidth <= 0 || cyPartHeight <= 0)
		{
			return FALSE;
		}

		CBCGPControlRendererParams params(hBmp, (COLORREF)-1, CRect(0, 0, cxPartWidth, cyPartHeight), CRect(2, 2, 2, 2));
		if (!m_pRrenderSysCaptionButton->Create(params))
		{
			return FALSE;
		}
	}

	if (!m_pRrenderSysCaptionButton->IsValid())
	{
		return FALSE;
	}
	
	int nState = 0;
	BYTE alphaScr = 198;
	
	if (!bIsActiveFrame)
	{
		nState = 7;
		alphaScr = 120;
	}
	else if (bDisabled)
	{
		nState = 3;
	}
	else if (bIsPressed && bIsHighlighted)
	{
		nState = 2;
	}
	else if (bIsHighlighted)
	{
		nState = 1;
	}

	m_pRrenderSysCaptionButton->Draw(pDC, rect, nState, alphaScr);
	return TRUE;
}
//**********************************************************************************
BOOL CBCGPWinXPThemeManager::DrawPushButton (CDC* pDC, CRect rect, 
											 CBCGPButton* pButton, UINT /*uiState*/)
{
	if (m_hThemeButton == NULL)
	{
		return FALSE;
	}

	int nState = PBS_NORMAL;

	if (pButton->GetSafeHwnd() != NULL)
	{
		if (!pButton->IsWindowEnabled ())
		{
			nState = PBS_DISABLED;
		}
		else if (pButton->IsPressed () || pButton->GetCheck ())
		{
			nState = PBS_PRESSED;
		}
		else if (pButton->IsHighlighted ())
		{
			nState = PBS_HOT;
		}
		else if (pButton->IsDefaultButton ())
		{
			nState = PBS_DEFAULTED;
		}

		pButton->OnDrawParentBackground (pDC, rect);
	}

	(*m_pfDrawThemeBackground) (m_hThemeButton, pDC->GetSafeHdc(), BP_PUSHBUTTON, nState, &rect, 0);

	return TRUE;
}
//**************************************************************************************
BOOL CBCGPWinXPThemeManager::DrawStatusBarProgress (CDC* pDC, CBCGPStatusBar* /*pStatusBar*/,
			CRect rectProgress, int nProgressTotal, int nProgressCurr,
			COLORREF /*clrBar*/, COLORREF /*clrProgressBarDest*/, COLORREF /*clrProgressText*/,
			BOOL bProgressText)
{
	ASSERT_VALID (pDC);

	CRect rectChunk = rectProgress;
	
	if (nProgressTotal == 0)
	{
		rectChunk.SetRectEmpty();
	}
	else
	{
		rectChunk.right = rectChunk.left + nProgressCurr * rectChunk.Width () / nProgressTotal;
	}
	
	if (!DrawProgressBar(pDC, rectProgress, rectChunk, FALSE, FALSE, nProgressCurr, 0, nProgressTotal))
	{
		return FALSE;
	}

	if (bProgressText)
	{
		CString strText;
		strText.Format (_T("%d%%"), nProgressCurr * 100 / nProgressTotal);

		COLORREF clrText = pDC->SetTextColor (globalData.clrBtnText);
		pDC->DrawText (strText, rectProgress, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
		pDC->SetTextColor (clrText);
	}

	return TRUE;
}
//********************************************************************************
BOOL CBCGPWinXPThemeManager::DrawComboDropButton (CDC* pDC, CRect rect,
										BOOL bDisabled,
										BOOL bIsDropped,
										BOOL bIsHighlighted)
{
	if (m_hThemeComboBox == NULL)
	{
		return FALSE;
	}

	int nState = bDisabled ? CBXS_DISABLED : bIsDropped ? CBXS_PRESSED : bIsHighlighted ? CBXS_HOT : CBXS_NORMAL;

	(*m_pfDrawThemeBackground) (m_hThemeComboBox, pDC->GetSafeHdc(), CP_DROPDOWNBUTTON, 
		nState, &rect, 0);

	return TRUE;
}
//********************************************************************************
BOOL CBCGPWinXPThemeManager::DrawComboBorder (CDC* pDC, CRect rect,
										BOOL /*bDisabled*/,
										BOOL bIsDropped,
										BOOL bIsHighlighted)
{
	ASSERT_VALID (pDC);

	if (m_hThemeWindow == NULL)
	{
		return FALSE;
	}

	if (bIsHighlighted || bIsDropped)
	{
		rect.DeflateRect (1, 1);
		pDC->Draw3dRect (&rect,  globalData.clrHilite, globalData.clrHilite);
	}

	return TRUE;
}
//********************************************************************************
void CBCGPWinXPThemeManager::FillRebarPane (CDC* pDC, CBCGPBaseControlBar* pBar, CRect rectClient)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pBar);

	if (m_pfDrawThemeBackground == NULL || m_hThemeRebar == NULL)
	{
		pDC->FillRect (rectClient, &globalData.brBarFace);
		return;
	}

	CWnd* pWndParent = BCGPGetParentFrame (pBar);
	if (pWndParent->GetSafeHwnd () == NULL)
	{
		pWndParent = pBar->GetParent ();
	}

	ASSERT_VALID (pWndParent);

	CRect rectParent;
	pWndParent->GetWindowRect (rectParent);

	pBar->ScreenToClient (&rectParent);

	rectClient.right = max (rectClient.right, rectParent.right);
	rectClient.bottom = max (rectClient.bottom, rectParent.bottom);

	if (!pBar->IsFloating () && pBar->GetParentMiniFrame () == NULL)
	{
		rectClient.left = rectParent.left;
		rectClient.top = rectParent.top;

		if (!pBar->IsKindOf (RUNTIME_CLASS (CBCGPDockBar)))
		{
			CFrameWnd* pMainFrame = BCGCBProGetTopLevelFrame (pWndParent);
			if (pMainFrame->GetSafeHwnd () != NULL)
			{
				CRect rectMain;
				pMainFrame->GetClientRect (rectMain);
				pMainFrame->MapWindowPoints (pBar, &rectMain);

				rectClient.top = rectMain.top;
			}
		}
	}

	(*m_pfDrawThemeBackground) (m_hThemeRebar, pDC->GetSafeHdc(),
		0, 0, &rectClient, 0);
}
//*********************************************************************************************************
BOOL CBCGPWinXPThemeManager::DrawCheckBox (CDC *pDC, CRect rect, 
										 BOOL bHighlighted, 
										 int nState,
										 BOOL bEnabled,
										 BOOL bPressed)
{
	if (m_hThemeButton == NULL)
	{
		return FALSE;
	}

	nState = max (0, nState);
	nState = min (2, nState);

	ASSERT_VALID (pDC);

	BOOL bDrawCustomCheck = FALSE;

	if (m_bScaleCheckRadio && nState == 1)
	{
		nState = 0;
		bDrawCustomCheck = TRUE;
	}

	int nDrawState = nState == 1 ? CBS_CHECKEDNORMAL : nState == 2 ? CBS_MIXEDNORMAL : CBS_UNCHECKEDNORMAL;

	if (!bEnabled)
	{
		nDrawState = nState == 1 ? CBS_CHECKEDDISABLED : nState == 2 ? CBS_MIXEDDISABLED : CBS_UNCHECKEDDISABLED;
	}
	else if (bPressed)
	{
		nDrawState = nState == 1 ? CBS_CHECKEDPRESSED : nState == 2 ? CBS_MIXEDPRESSED : CBS_UNCHECKEDPRESSED;
	}
	else if (bHighlighted)
	{
		nDrawState = nState == 1 ? CBS_CHECKEDHOT : nState == 2 ? CBS_MIXEDHOT : CBS_UNCHECKEDHOT;
	}

	(*m_pfDrawThemeBackground) (m_hThemeButton, pDC->GetSafeHdc(), BP_CHECKBOX,
		nDrawState, &rect, 0);

	if (bDrawCustomCheck)
	{
		rect.DeflateRect(rect.Width() / 4 + 1, rect.Height() / 4 + 1);
		COLORREF clrLine = bEnabled ? globalData.clrBarText : globalData.clrGrayedText;
		
		CBCGPDrawManager dm(*pDC);
		dm.DrawCheckMark(rect, clrLine, (int)(0.5 + m_dblScaleCheckRadio));
	}

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPWinXPThemeManager::DrawRadioButton (CDC *pDC, CRect rect, 
										 BOOL bHighlighted, 
										 BOOL bChecked,
										 BOOL bEnabled,
										 BOOL bPressed)
{
	if (m_hThemeButton == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID (pDC);

	int nDrawState = bChecked ? RBS_CHECKEDNORMAL : RBS_UNCHECKEDNORMAL;

	if (!bEnabled)
	{
		nDrawState = bChecked ? RBS_CHECKEDDISABLED : RBS_UNCHECKEDDISABLED;
	}
	else if (bPressed)
	{
		nDrawState = bChecked ? RBS_CHECKEDPRESSED : RBS_UNCHECKEDPRESSED;
	}
	else if (bHighlighted)
	{
		nDrawState = bChecked ? RBS_CHECKEDHOT : RBS_UNCHECKEDHOT;
	}

	(*m_pfDrawThemeBackground) (m_hThemeButton, pDC->GetSafeHdc(), BP_RADIOBUTTON,
		nDrawState, &rect, 0);

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPWinXPThemeManager::DrawProgressBar (CDC *pDC, CRect rectProgress, 
											  CRect rectChunk, BOOL bInfiniteMode, BOOL bIsVertical,
											  int nPos, int nMin, int nMax)
{
	if (m_hThemeProgress == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID (pDC);

	UINT nPart = PP_BAR;
	UINT nState = 0;
	
	if (globalData.bIsWindowsVista)
	{
		nPart = bIsVertical ? PP_TRANSPARENTBARVERT : PP_TRANSPARENTBAR;
		nState = PBBS_NORMAL;
	}
	
	(*m_pfDrawThemeBackground) (m_hThemeProgress, pDC->GetSafeHdc(), nPart, nState, &rectProgress, 0);

	BOOL bDrawChunk = FALSE;
	
	if (!rectChunk.IsRectEmpty())
	{
		if (!bInfiniteMode)
		{
			if (nPos != nMin)
			{
				if (bIsVertical)
				{
					rectChunk.bottom--;
				}
				else
				{
					rectChunk.left++;
				}

				bDrawChunk = TRUE;
			}
		}
		else if (nMin != nMax)
		{
			bDrawChunk = TRUE;
		}
	}

	if (bDrawChunk)
	{
		nPart = bIsVertical ? PP_CHUNKVERT : PP_CHUNK;
		nState = 0;
		
		if (globalData.bIsWindowsVista)
		{
			nPart = bIsVertical ? PP_FILLVERT : PP_FILL;
			nState = PBFS_NORMAL;
		}
		else
		{
			if (bIsVertical)
			{
				rectChunk.DeflateRect(1, 0);
			}
			else
			{
				rectChunk.DeflateRect(0, 1);
			}
		}
		
		(*m_pfDrawThemeBackground) (m_hThemeProgress, pDC->GetSafeHdc(), nPart, nState, &rectChunk, &rectProgress);
	}

	return TRUE;
}
//*******************************************************************************
CBCGPWinXPThemeManager::WinXpTheme CBCGPWinXPThemeManager::GetStandardWinXPTheme ()
{
	WCHAR szName [256];
	WCHAR szColor [256];

	if (m_pfGetCurrentThemeName == NULL ||
		(*m_pfGetCurrentThemeName) (szName, 255, szColor, 255, NULL, 0) != S_OK)
	{
		return WinXpTheme_None;
	}

	CString strThemeName = szName;
	CString strWinXPThemeColor = szColor;

	TCHAR fname[_MAX_FNAME];   
#if _MSC_VER < 1400
	_tsplitpath (strThemeName, NULL, NULL, fname, NULL);
#else
	_tsplitpath_s (strThemeName, NULL, 0, NULL, 0, fname, _MAX_FNAME, NULL, 0);
#endif

	strThemeName = fname;

	if (strThemeName.CompareNoCase (_T("Luna")) != 0 && 
		strThemeName.CompareNoCase (_T("Aero")) != 0)
	{
		return WinXpTheme_NonStandard;
	}

	// Check for 3-d party visual managers:
	if (m_pfGetThemeColor != NULL && m_hThemeButton != NULL)
	{
		COLORREF clrTest = 0;
		if ((*m_pfGetThemeColor) (m_hThemeButton, BP_PUSHBUTTON, 0, TMT_ACCENTCOLORHINT, &clrTest) !=
			S_OK || clrTest == 1)
		{
			return WinXpTheme_NonStandard;
		}
	}

	if (strWinXPThemeColor.CompareNoCase (_T("normalcolor")) == 0)
	{
		return WinXpTheme_Blue;
	}

	if (strWinXPThemeColor.CompareNoCase (_T("homestead")) == 0)
	{
		return WinXpTheme_Olive;
	}

	if (strWinXPThemeColor.CompareNoCase (_T("metallic")) == 0)
	{
		// Check for Royale theme:
		CString strThemeLower = szName;
		strThemeLower.MakeLower ();

		if (strThemeLower.Find (_T("royale")) >= 0)
		{
			return WinXpTheme_NonStandard;
		}

		return WinXpTheme_Silver;
	}

	return WinXpTheme_NonStandard;
}
