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
// BCGPGlobalUtils.cpp: implementation of the CBCGPGlobalUtils class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BCGPGlobalUtils.h"

#ifndef _BCGSUITE_

#include "BCGPDockManager.h"
#include "BCGPBarContainerManager.h"
#include "BCGPDockingControlBar.h"
#include "BCGPMiniFrameWnd.h"
#include "BCGPMultiMiniFrameWnd.h"
#include "BCGPBaseTabbedBar.h"
#include "BCGPRibbonBar.h"

#include "BCGPFrameWnd.h"
#include "BCGPMDIFrameWnd.h"
#include "BCGPOleIPFrameWnd.h"
#include "BCGPOleDocIPFrameWnd.h"
#include "BCGPMDIChildWnd.h"
#include "BCGPOleCntrFrameWnd.h"
#include "BCGPURLLinkButton.h"
#include "BCGCBProVer.h"
#else
#include "BCGSuiteVer.h"
#endif

#include "BCGPDrawManager.h"
#include "BCGPDialog.h"
#include "BCGPPropertyPage.h"
#include "BCGPPropertySheet.h"
#include "BCGPFormView.h"
#include "BCGPDialogBar.h"
#include "BCGPCalendarBar.h"
#include "BCGPImageProcessing.h"
#include "bcgprores.h"

#ifndef SM_CXPADDEDBORDER
	#define SM_CXPADDEDBORDER 92
#endif

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

CBCGPGlobalUtils globalUtils;

#ifdef BCGP_PLATFORM_TOOLSET_DEF
	#define BCGP_PLATFORM_TOOLSET_TEMP(x) (#x)
	#define BCGP_PLATFORM_TOOLSET BCGP_PLATFORM_TOOLSET_TEMP(BCGP_PLATFORM_TOOLSET_DEF)
#endif

static BOOL s_bIsXPPlatformToolset = TRUE;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPGlobalUtils::CBCGPGlobalUtils()
{
	m_bDialogApp = FALSE;
	m_bIsDragging = FALSE;
	m_bUseMiniFrameParent = FALSE;
	m_bIsAdjustLayout = FALSE;
	m_bRemoveDoubleAmpersands = TRUE;

#ifdef _BCGSUITE_
	m_bUseDlgFontInControls = FALSE;
#endif

#if defined(BCGP_PLATFORM_TOOLSET)
	#if (BCGP_PLATFORM_TOOLSET_VERSION > 90)
		s_bIsXPPlatformToolset = strcmp(BCGP_PLATFORM_TOOLSET, "v110_xp") == 0 ||
								strcmp(BCGP_PLATFORM_TOOLSET, "v120_xp") == 0 ||
								strcmp(BCGP_PLATFORM_TOOLSET, "v140_xp") == 0 ||
								strcmp(BCGP_PLATFORM_TOOLSET, "v141_xp") == 0;
	#endif
#endif
}

CBCGPGlobalUtils::~CBCGPGlobalUtils()
{

}

#ifndef _BCGSUITE_

//------------------------------------------------------------------------//
BOOL CBCGPGlobalUtils::CheckAlignment (CPoint point, CBCGPBaseControlBar* pBar, int nSencitivity, 
                                       const CBCGPDockManager* pDockManager,
									   BOOL bOuterEdge, DWORD& dwAlignment, 
									   DWORD dwEnabledDockBars, LPCRECT lpRectBounds) const
{
    BOOL bSmartDocking = FALSE;
    CBCGPSmartDockingMarker::SDMarkerPlace nHilitedSide = CBCGPSmartDockingMarker::sdNONE;
	
	if (pDockManager == NULL && pBar != NULL)
    {
        pDockManager = globalUtils.GetDockManager (pBar->GetParent());
    }
	
	if (pDockManager != NULL)
	{
        CBCGPSmartDockingManager* pSDManager = pDockManager->GetSDManagerPermanent ();
        if (pSDManager != NULL && pSDManager->IsStarted ())
        {
            bSmartDocking = TRUE;
            nHilitedSide = pSDManager->GetHilitedMarkerNo ();
        }
	}
	
	CRect rectBounds;
	if (pBar != NULL)
	{
		pBar->GetWindowRect (rectBounds);
		
	}
	else if (lpRectBounds != NULL)
	{
		rectBounds = *lpRectBounds;
	}
	else
	{
		ASSERT(FALSE);
		return FALSE;
	}
	
	int nCaptionHeight = 0;
	int nTabAreaTopHeight = 0; 
	int nTabAreaBottomHeight = 0;
	
	CBCGPDockingControlBar* pDockingBar = 
		DYNAMIC_DOWNCAST (CBCGPDockingControlBar, pBar);
	
	if (pDockingBar != NULL)
	{
		nCaptionHeight = pDockingBar->GetCaptionHeight ();
		
		CRect rectTabAreaTop;
		CRect rectTabAreaBottom;
		pDockingBar->GetTabArea (rectTabAreaTop, rectTabAreaBottom);
		nTabAreaTopHeight = rectTabAreaTop.Height ();
		nTabAreaBottomHeight = rectTabAreaBottom.Height ();
	}
	
	// build rect for top area
	if (bOuterEdge)
	{
        if (bSmartDocking)
        {
            switch (nHilitedSide)
            {
            case CBCGPSmartDockingMarker::sdLEFT:
				dwAlignment = CBRS_ALIGN_LEFT;
				return TRUE;
            case CBCGPSmartDockingMarker::sdRIGHT:
				dwAlignment = CBRS_ALIGN_RIGHT;
				return TRUE;
            case CBCGPSmartDockingMarker::sdTOP:
				dwAlignment = CBRS_ALIGN_TOP;
				return TRUE;
            case CBCGPSmartDockingMarker::sdBOTTOM:
				dwAlignment = CBRS_ALIGN_BOTTOM;
				return TRUE;
            }
        }
		else
        {
			CRect rectToCheck (rectBounds.left - nSencitivity, rectBounds.top - nSencitivity, 
				rectBounds.right + nSencitivity, rectBounds.top);
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_TOP)
			{
				dwAlignment = CBRS_ALIGN_TOP;
				return TRUE;
			}
			
			// build rect for left area
			rectToCheck.right = rectBounds.left;
			rectToCheck.bottom = rectBounds.bottom + nSencitivity;
			
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_LEFT)
			{
				dwAlignment = CBRS_ALIGN_LEFT;
				return TRUE;
			}
			
			// build rect for bottom area
			rectToCheck.left = rectBounds.left - nSencitivity;
			rectToCheck.top = rectBounds.bottom;
			rectToCheck.right = rectBounds.right + nSencitivity;
			rectToCheck.bottom = rectBounds.bottom + nSencitivity;
			
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_BOTTOM)
			{
				dwAlignment = CBRS_ALIGN_BOTTOM;
				return TRUE;
			}
			
			// build rect for right area
			rectToCheck.left = rectBounds.right;
			rectToCheck.top = rectBounds.top - nSencitivity;
			
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_RIGHT)
			{
				dwAlignment = CBRS_ALIGN_RIGHT;
				return TRUE;
			}
        }
	}
	else
	{
        if (bSmartDocking)
        {
            switch (nHilitedSide)
            {
            case CBCGPSmartDockingMarker::sdCLEFT:
				dwAlignment = CBRS_ALIGN_LEFT;
				return TRUE;
            case CBCGPSmartDockingMarker::sdCRIGHT:
				dwAlignment = CBRS_ALIGN_RIGHT;
				return TRUE;
            case CBCGPSmartDockingMarker::sdCTOP:
				dwAlignment = CBRS_ALIGN_TOP;
				return TRUE;
            case CBCGPSmartDockingMarker::sdCBOTTOM:
				dwAlignment = CBRS_ALIGN_BOTTOM;
				return TRUE;
            }
        }
		else
        {
#ifdef __BOUNDS_FIX__
			CRect rectToCheck (rectBounds.left, rectBounds.top, 
				rectBounds.right, 
				rectBounds.top + nSencitivity + nCaptionHeight);
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_TOP)
			{
				dwAlignment = CBRS_ALIGN_TOP;
				return TRUE;
			}
			
			
			// build rect for left area
			rectToCheck.right = rectBounds.left + nSencitivity;
			rectToCheck.bottom = rectBounds.bottom + nSencitivity;
			
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_LEFT)
			{
				dwAlignment = CBRS_ALIGN_LEFT;
				return TRUE;
			}
			
			// build rect for bottom area
			rectToCheck.left = rectBounds.left;
			rectToCheck.top = rectBounds.bottom - nSencitivity - nTabAreaBottomHeight;
			rectToCheck.right = rectBounds.right;
			rectToCheck.bottom = rectBounds.bottom;
			
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_BOTTOM)
			{
				dwAlignment = CBRS_ALIGN_BOTTOM;
				return TRUE;
			}
			
			// build rect for right area
			rectToCheck.left = rectBounds.right - nSencitivity;
			rectToCheck.top = rectBounds.top - nSencitivity;
			
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_RIGHT)
			{
				dwAlignment = CBRS_ALIGN_RIGHT;
				return TRUE;
			}
#else
			
			// build rect for top area
			CRect rectToCheck (rectBounds.left - nSencitivity, rectBounds.top - nSencitivity, 
				rectBounds.right + nSencitivity, 
				rectBounds.top + nSencitivity + nCaptionHeight);
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_TOP)
			{
				dwAlignment = CBRS_ALIGN_TOP;
				return TRUE;
			}
			
			
			// build rect for left area
			rectToCheck.right = rectBounds.left + nSencitivity;
			rectToCheck.bottom = rectBounds.bottom + nSencitivity;
			
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_LEFT)
			{
				dwAlignment = CBRS_ALIGN_LEFT;
				return TRUE;
			}
			
			// build rect for bottom area
			rectToCheck.left = rectBounds.left - nSencitivity;
			rectToCheck.top = rectBounds.bottom - nSencitivity - nTabAreaBottomHeight;
			rectToCheck.right = rectBounds.right + nSencitivity;
			rectToCheck.bottom = rectBounds.bottom + nSencitivity;
			
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_BOTTOM)
			{
				dwAlignment = CBRS_ALIGN_BOTTOM;
				return TRUE;
			}
			
			// build rect for right area
			rectToCheck.left = rectBounds.right - nSencitivity;
			rectToCheck.top = rectBounds.top - nSencitivity;
			
			if (rectToCheck.PtInRect (point) && dwEnabledDockBars & CBRS_ALIGN_RIGHT)
			{
				dwAlignment = CBRS_ALIGN_RIGHT;
				return TRUE;
			}
#endif		
        }
	}
	
	return FALSE;
}
//------------------------------------------------------------------------//
CBCGPDockManager* CBCGPGlobalUtils::GetDockManager (CWnd* pWnd)
{
	if (pWnd == NULL)
	{
		return NULL;
	}

	ASSERT_VALID (pWnd);

	if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPFrameWnd)))
	{
		return ((CBCGPFrameWnd*) pWnd)->GetDockManager ();
	}
	else if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPMDIFrameWnd)))
	{
		return ((CBCGPMDIFrameWnd*) pWnd)->GetDockManager ();
	}
	else if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPOleIPFrameWnd)))
	{
		return ((CBCGPOleIPFrameWnd*) pWnd)->GetDockManager ();
	}
	else if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPOleDocIPFrameWnd)))
	{
		return ((CBCGPOleDocIPFrameWnd*) pWnd)->GetDockManager ();
	}
	else if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPMDIChildWnd)))
	{
		return ((CBCGPMDIChildWnd*) pWnd)->GetDockManager ();
	}
	else if (pWnd->IsKindOf (RUNTIME_CLASS (CDialog)) ||
		pWnd->IsKindOf (RUNTIME_CLASS (CPropertySheet)))
	{
		if (pWnd->GetSafeHwnd() == AfxGetMainWnd()->GetSafeHwnd())
		{
			m_bDialogApp = TRUE;
		}
	}
	else if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPOleCntrFrameWnd)))
	{
		return ((CBCGPOleCntrFrameWnd*) pWnd)->GetDockManager ();
	}
    else if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPMiniFrameWnd)))
    {
		CBCGPMiniFrameWnd* pMiniFrameWnd = DYNAMIC_DOWNCAST (CBCGPMiniFrameWnd, pWnd);
		ASSERT_VALID (pMiniFrameWnd);

		CBCGPDockManager* pManager = pMiniFrameWnd->GetDockManager ();
        return pManager != NULL ? pManager : GetDockManager (m_bUseMiniFrameParent ? pMiniFrameWnd->GetParent() : pWnd->GetParent ());
    }

	return NULL;
}
//------------------------------------------------------------------------//
void CBCGPGlobalUtils::FlipRect (CRect& rect, int nDegrees)
{
	CRect rectTmp = rect;
	switch (nDegrees)
	{
	case 90:
		rect.top = rectTmp.left;
		rect.right = rectTmp.top;
		rect.bottom = rectTmp.right;
		rect.left = rectTmp.bottom;
		break;
	case 180:
		rect.top = rectTmp.bottom;
		rect.bottom = rectTmp.top;
		break;
	case 275:
	case -90:
		rect.left = rectTmp.top;
		rect.top = rectTmp.right;
		rect.right = rectTmp.bottom;
		rect.bottom = rectTmp.left;
		break;
	}
}
//------------------------------------------------------------------------//
DWORD CBCGPGlobalUtils::GetOppositeAlignment (DWORD dwAlign)
{
	switch (dwAlign & CBRS_ALIGN_ANY)
	{
	case CBRS_ALIGN_LEFT:
		return CBRS_ALIGN_RIGHT;
	case CBRS_ALIGN_RIGHT:
		return CBRS_ALIGN_LEFT;
	case CBRS_ALIGN_TOP:
		return CBRS_ALIGN_BOTTOM;
	case CBRS_ALIGN_BOTTOM:
		return CBRS_ALIGN_TOP;
	}
	return 0;
}
//------------------------------------------------------------------------//
void CBCGPGlobalUtils::SetNewParent (CObList& lstControlBars, CWnd* pNewParent,
									 BOOL bCheckVisibility)
{
	ASSERT_VALID (pNewParent);
	for (POSITION pos = lstControlBars.GetHeadPosition (); pos != NULL;)
	{
		CBCGPBaseControlBar* pBar = (CBCGPBaseControlBar*) lstControlBars.GetNext (pos);

		if (bCheckVisibility && !pBar->IsBarVisible ())
		{
			continue;
		}
		if (!pBar->IsKindOf (RUNTIME_CLASS (CBCGPSlider)))
		{
			pBar->ShowWindow (SW_HIDE);
			pBar->SetParent (pNewParent);
			CRect rectWnd;
			pBar->GetWindowRect (rectWnd);
			pNewParent->ScreenToClient (rectWnd);

			pBar->SetWindowPos (NULL, -rectWnd.Width (), -rectWnd.Height (), 
									  100, 100, SWP_NOZORDER | SWP_NOSIZE  | SWP_NOACTIVATE);
			pBar->ShowWindow (SW_SHOW);
		}
		else
		{
			pBar->SetParent (pNewParent);
		}
	}
}
//------------------------------------------------------------------------//
void CBCGPGlobalUtils::CalcExpectedDockedRect (CBCGPBarContainerManager& barContainerManager, 
												CWnd* pWndToDock, CPoint ptMouse, 
												CRect& rectResult, BOOL& bDrawTab, 
												CBCGPDockingControlBar** ppTargetBar)
{
	ASSERT (ppTargetBar != NULL);

	DWORD dwAlignment = CBRS_ALIGN_ANY;
	BOOL bTabArea = FALSE;
	BOOL bCaption = FALSE;
	bDrawTab = FALSE;
	*ppTargetBar = NULL;

	rectResult.SetRectEmpty ();

	if (GetKeyState (VK_CONTROL) < 0)
	{
		return;
	}

	if (!GetCBAndAlignFromPoint (barContainerManager, ptMouse, 
								 ppTargetBar, dwAlignment, bTabArea, bCaption) || 
		*ppTargetBar == NULL)
	{
		return;
	}

	CBCGPControlBar* pBar = NULL;
	
	if (pWndToDock->IsKindOf (RUNTIME_CLASS (CBCGPMiniFrameWnd)))
	{
		CBCGPMiniFrameWnd* pMiniFrame = 
			DYNAMIC_DOWNCAST (CBCGPMiniFrameWnd, pWndToDock);
		ASSERT_VALID (pMiniFrame);
		pBar = DYNAMIC_DOWNCAST (CBCGPControlBar, pMiniFrame->GetFirstVisibleBar ());
	}
	else
	{
		pBar = DYNAMIC_DOWNCAST (CBCGPControlBar, pWndToDock);
	}

	if (*ppTargetBar != NULL)
	{
		DWORD dwTargetEnabledAlign = (*ppTargetBar)->GetEnabledAlignment ();
		DWORD dwTargetCurrentAlign = (*ppTargetBar)->GetCurrentAlignment ();
		BOOL bTargetBarIsFloating = ((*ppTargetBar)->GetParentMiniFrame () != NULL);

		if (pBar != NULL)
		{
			if (pBar->GetEnabledAlignment () != dwTargetEnabledAlign && bTargetBarIsFloating ||
				(pBar->GetEnabledAlignment () & dwTargetCurrentAlign) == 0 && !bTargetBarIsFloating)
			{
				return;
			}
		}
	}

	if (bTabArea || bCaption)
	{
		// can't make tab on miniframe
		bDrawTab = ((*ppTargetBar) != NULL);

		if (bDrawTab)
		{
			bDrawTab = (*ppTargetBar)->CanBeAttached () && CanBeAttached (pWndToDock) && 
				pBar != NULL && ((*ppTargetBar)->GetEnabledAlignment () == pBar->GetEnabledAlignment ());
		}

		if (!bDrawTab)
		{
			return;
		}

	}

	if ((*ppTargetBar) != NULL && (*ppTargetBar)->GetParentMiniFrame () != NULL && 
		!IsWndCanFloatMulti (pWndToDock))
	{
		bDrawTab = FALSE;
		return;
	}

	if ((*ppTargetBar) != NULL && 
		pWndToDock->IsKindOf (RUNTIME_CLASS (CBCGPBaseControlBar)) && 
		!(*ppTargetBar)->CanAcceptBar ((CBCGPBaseControlBar*) pWndToDock))
	{
		bDrawTab = FALSE;
		return;
	}

	CRect rectOriginal; 
	(*ppTargetBar)->GetWindowRect (rectOriginal);

	if ((*ppTargetBar) == pWndToDock ||
		pWndToDock->IsKindOf (RUNTIME_CLASS (CBCGPMiniFrameWnd)) && 
		(*ppTargetBar)->GetParentMiniFrame () == pWndToDock)
	{
		bDrawTab = FALSE;
		return;
	}
	
	CRect rectInserted;
	CRect rectSlider;
	DWORD dwSliderStyle;
	CSize sizeMinOriginal (0, 0);
	CSize sizeMinInserted (0, 0);

	
	pWndToDock->GetWindowRect (rectInserted);

	if ((dwAlignment & pBar->GetEnabledAlignment ()) != 0 ||
		CBCGPDockManager::m_bIgnoreEnabledAlignment)
	{
		barContainerManager.CalcRects (rectOriginal, rectInserted, rectSlider, dwSliderStyle, 
										 dwAlignment, sizeMinOriginal, sizeMinInserted);
		rectResult = rectInserted;
	}
	
}
//------------------------------------------------------------------------//
BOOL CBCGPGlobalUtils::CanBeAttached (CWnd* pWnd) const
{
	ASSERT_VALID (pWnd);

	if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPMiniFrameWnd)))
	{
		return ((CBCGPMiniFrameWnd*) pWnd)->CanBeAttached ();
	}

	if (pWnd->IsKindOf (RUNTIME_CLASS (CBCGPControlBar)))
	{
		return ((CBCGPControlBar*) pWnd)->CanBeAttached ();
	}

	return FALSE;

}
//------------------------------------------------------------------------//
BOOL CBCGPGlobalUtils::IsWndCanFloatMulti (CWnd* pWnd) const
{
	CBCGPControlBar* pBar = NULL;

	CBCGPMiniFrameWnd* pMiniFrame = 
		DYNAMIC_DOWNCAST (CBCGPMiniFrameWnd, pWnd);
	
	if (pMiniFrame != NULL)
	{
		pBar = DYNAMIC_DOWNCAST (CBCGPControlBar, pMiniFrame->GetControlBar ());
	}
	else
	{
		pBar = DYNAMIC_DOWNCAST (CBCGPControlBar, pWnd);

	}

	if (pBar == NULL)
	{
		return FALSE;
	}

	if (pBar->IsTabbed ())
	{
		CWnd* pParentMiniFrame = pBar->GetParentMiniFrame ();
		// tabbed bar that is floating in multi miniframe
		if (pParentMiniFrame != NULL && pParentMiniFrame->IsKindOf (RUNTIME_CLASS (CBCGPMultiMiniFrameWnd)))
		{
			return TRUE;
		}
	}

	
	return ((pBar->GetBarStyle () & CBRS_FLOAT_MULTI) != 0);
}
//------------------------------------------------------------------------//
BOOL CBCGPGlobalUtils::GetCBAndAlignFromPoint (CBCGPBarContainerManager& barContainerManager, 
													 CPoint pt, 
												     CBCGPDockingControlBar** ppTargetControlBar, 
												     DWORD& dwAlignment, 
													 BOOL& bTabArea, 
													 BOOL& bCaption)
{
	ASSERT (ppTargetControlBar != NULL);
	*ppTargetControlBar = NULL;

	BOOL bOuterEdge = FALSE;

	// if the mouse is over a miniframe's caption and this miniframe has only one
	// visible docking control bar, we need to draw a tab
	bCaption = barContainerManager.CheckForMiniFrameAndCaption (pt, ppTargetControlBar);
	if (bCaption)
	{
		return TRUE;
	}


	*ppTargetControlBar = 
		barContainerManager.ControlBarFromPoint (pt, CBCGPDockManager::m_nDockSencitivity, 
													TRUE, bTabArea, bCaption);

	if ((bCaption || bTabArea) && *ppTargetControlBar != NULL) 
	{
		return TRUE;
	}

	if (*ppTargetControlBar == NULL)
	{
		barContainerManager.ControlBarFromPoint (pt, CBCGPDockManager::m_nDockSencitivity, 
														FALSE, bTabArea, bCaption);
		// the exact bar was not found - it means the docked frame at the outer edge
		bOuterEdge = TRUE;
		return TRUE;
	}

	if (*ppTargetControlBar != NULL)
	{
		if (!globalUtils.CheckAlignment (pt, *ppTargetControlBar,
										CBCGPDockManager::m_nDockSencitivity, NULL,
										bOuterEdge, dwAlignment))
		{
			// unable for some reason to determine alignment
			*ppTargetControlBar = NULL;
		}
	}

	return TRUE;
}
//------------------------------------------------------------------------//
void CBCGPGlobalUtils::AdjustRectToWorkArea (CRect& rect, CRect* pRectDelta)
{
	CPoint ptStart;

	if (m_bIsDragging)
	{
		::GetCursorPos (&ptStart);
	}
	else
	{
		ptStart = rect.TopLeft ();
	}

	CRect rectScreen;
	MONITORINFO mi;
	mi.cbSize = sizeof (MONITORINFO);
	if (GetMonitorInfo (
		MonitorFromPoint (ptStart, MONITOR_DEFAULTTONEAREST),
		&mi))
	{
		rectScreen = mi.rcWork;
	}
	else
	{
		::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectScreen, 0);
	}

	int nDelta = pRectDelta != NULL ? pRectDelta->left : 0;

	if (rect.right <= rectScreen.left + nDelta)
	{
		rect.OffsetRect (rectScreen.left - rect.right + nDelta, 0);
	}

	nDelta = pRectDelta != NULL ? pRectDelta->right : 0;
	if (rect.left >= rectScreen.right - nDelta)
	{
		rect.OffsetRect (rectScreen.right - rect.left - nDelta, 0);
	}

	nDelta = pRectDelta != NULL ? pRectDelta->bottom : 0;
	if (rect.top >= rectScreen.bottom - nDelta)
	{
		rect.OffsetRect (0, rectScreen.bottom - rect.top - nDelta);
	}

	nDelta = pRectDelta != NULL ? pRectDelta->top : 0;
	if (rect.bottom < rectScreen.top + nDelta)
	{
		rect.OffsetRect (0, rectScreen.top - rect.bottom + nDelta);
	}
}
//------------------------------------------------------------------------//
void CBCGPGlobalUtils::ForceAdjustLayout (CBCGPDockManager* pDockManager, BOOL bForce,
										  BOOL bForceInvisible, BOOL bForceNcArea)
{
	if (pDockManager != NULL && 
		(CBCGPControlBar::m_bHandleMinSize || bForce))
	{
		CWnd* pDockSite = pDockManager->GetDockSite ();

		if (pDockSite == NULL)
		{
			return;
		}
			
		if (!pDockSite->IsWindowVisible () && !bForceInvisible)
		{
			return;
		}

		if (pDockSite->IsKindOf(RUNTIME_CLASS(CBCGPFrameWnd)) || pDockSite->IsKindOf(RUNTIME_CLASS(CBCGPMDIFrameWnd)))
		{
			m_bIsAdjustLayout = (pDockSite->GetStyle() & WS_MAXIMIZE) == WS_MAXIMIZE;
		}

		CRect rectWnd;
		pDockSite->SetRedraw (FALSE);
		pDockSite->GetWindowRect (rectWnd);
		pDockSite->SetWindowPos (NULL, -1, -1, 
			rectWnd.Width () + 1, rectWnd.Height () + 1, 
			SWP_NOZORDER |  SWP_NOMOVE | SWP_NOACTIVATE);
		pDockSite->SetWindowPos (NULL, -1, -1, 
			rectWnd.Width (), rectWnd.Height (), 
			SWP_NOZORDER |  SWP_NOMOVE  | SWP_NOACTIVATE);
		pDockSite->SetRedraw (TRUE);

		if (bForceNcArea)
		{
			pDockSite->SetWindowPos(NULL, 0, 0, 0, 0, 
				SWP_FRAMECHANGED | SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
		}

		m_bIsAdjustLayout = FALSE;

		pDockSite->RedrawWindow (NULL, NULL, 
			RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
	}
}

#endif

//------------------------------------------------------------------------//
BOOL CBCGPGlobalUtils::StringFromCy(CString& str, CY& cy)
{
	VARIANTARG varCy;
	VARIANTARG varBstr;
	AfxVariantInit(&varCy);
	AfxVariantInit(&varBstr);
	V_VT(&varCy) = VT_CY;
	V_CY(&varCy) = cy;
	if (FAILED(VariantChangeType(&varBstr, &varCy, 0, VT_BSTR)))
	{
		VariantClear(&varCy);
		VariantClear(&varBstr);
		return FALSE;
	}
	str = V_BSTR(&varBstr);
	VariantClear(&varCy);
	VariantClear(&varBstr);
	return TRUE;
}
//------------------------------------------------------------------------//
BOOL CBCGPGlobalUtils::CyFromString(CY& cy, LPCTSTR psz)
{
	USES_CONVERSION;

	if (psz == NULL || _tcslen (psz) == 0)
	{
		psz = _T("0");
	}

	VARIANTARG varBstr;
	VARIANTARG varCy;
	AfxVariantInit(&varBstr);
	AfxVariantInit(&varCy);
	V_VT(&varBstr) = VT_BSTR;
	V_BSTR(&varBstr) = SysAllocString(T2COLE(psz));
	if (FAILED(VariantChangeType(&varCy, &varBstr, 0, VT_CY)))
	{
		VariantClear(&varBstr);
		VariantClear(&varCy);
		return FALSE;
	}
	cy = V_CY(&varCy);
	VariantClear(&varBstr);
	VariantClear(&varCy);
	return TRUE;
}
//------------------------------------------------------------------------//
BOOL CBCGPGlobalUtils::StringFromDecimal(CString& str, DECIMAL& decimal)
{
	VARIANTARG varDecimal;
	VARIANTARG varBstr;
	AfxVariantInit(&varDecimal);
	AfxVariantInit(&varBstr);
	V_VT(&varDecimal) = VT_DECIMAL;
	V_DECIMAL(&varDecimal) = decimal;
	if (FAILED(VariantChangeType(&varBstr, &varDecimal, 0, VT_BSTR)))
	{
		VariantClear(&varDecimal);
		VariantClear(&varBstr);
		return FALSE;
	}
	str = V_BSTR(&varBstr);
	VariantClear(&varDecimal);
	VariantClear(&varBstr);
	return TRUE;
}
//------------------------------------------------------------------------//
BOOL CBCGPGlobalUtils::DecimalFromString(DECIMAL& decimal, LPCTSTR psz)
{
	USES_CONVERSION;

	if (psz == NULL || _tcslen (psz) == 0)
	{
		psz = _T("0");
	}

	VARIANTARG varBstr;
	VARIANTARG varDecimal;
	AfxVariantInit(&varBstr);
	AfxVariantInit(&varDecimal);
	V_VT(&varBstr) = VT_BSTR;
	V_BSTR(&varBstr) = SysAllocString(T2COLE(psz));
	if (FAILED(VariantChangeType(&varDecimal, &varBstr, 0, VT_DECIMAL)))
	{
		VariantClear(&varBstr);
		VariantClear(&varDecimal);
		return FALSE;
	}
	decimal = V_DECIMAL(&varDecimal);
	VariantClear(&varBstr);
	VariantClear(&varDecimal);
	return TRUE;
}
//------------------------------------------------------------------------//
void CBCGPGlobalUtils::RemoveSingleAmp(CString& strText, BOOL bReplace2AmpWith1/* = TRUE*/)
{
	if (!m_bRemoveDoubleAmpersands)
	{
		return;
	}

	const CString strDummyAmpSeq = _T("\001\001");
	
	strText.Replace(_T("&&"), strDummyAmpSeq);
	strText.Remove(_T('&'));
	strText.Replace(strDummyAmpSeq, bReplace2AmpWith1 ? _T("&") : _T("&&"));
}
//------------------------------------------------------------------------//
void CBCGPGlobalUtils::StringAddItemQuoted (CString& strOutput, LPCTSTR pszItem, BOOL bInsertBefore, TCHAR charDelimiter, TCHAR charQuote)
{
	CString strItem(pszItem);

	TCHAR special[] = {charDelimiter, charQuote, 0};
	if (strItem.IsEmpty () || strItem.FindOneOf (special) >= 0)
	{
		TCHAR singleQuote[] = {charQuote, 0};
		TCHAR doubleQuote[] = {charQuote, charQuote, 0};
		strItem.Replace (singleQuote, doubleQuote);

		CString strTemp;
		strTemp.Format (_T("%c%s%c"), charQuote, (LPCTSTR)strItem, charQuote);
		strItem = strTemp;
	}

	CString strTemp;
	if (strOutput.IsEmpty ())
	{
		if (bInsertBefore)
		{
			strTemp.Format(_T("%s%s"), (LPCTSTR)strItem, (LPCTSTR)strOutput);
		}
		else
		{
			strTemp.Format(_T("%s%s"), (LPCTSTR)strOutput, (LPCTSTR)strItem);
		}
	}
	else
	{
		if (bInsertBefore)
		{
			strTemp.Format(_T("%s%c%s"), (LPCTSTR)strItem, charDelimiter, (LPCTSTR)strOutput);
		}
		else
		{
			strTemp.Format(_T("%s%c%s"), (LPCTSTR)strOutput, charDelimiter, (LPCTSTR)strItem);
		}
	}

	strOutput = strTemp;
}

CString CBCGPGlobalUtils::StringExtractItemQuoted (const CString& strSource, LPCTSTR pszDelimiters, int& iStart, TCHAR charQuote)
{
	ASSERT (iStart >= 0);
	int iLen = strSource.GetLength ();
	if (iStart >= iLen)
	{
		return CString();
	}
		
	if (pszDelimiters == NULL || *pszDelimiters == 0)
	{
		return strSource.Mid (iStart);
	}

	int iPos = iStart;

	bool bQuoted = (strSource[iStart] == charQuote);
	if (bQuoted)
	{
		iPos ++;
	}

	CString strTokens(pszDelimiters);
	for (; iPos < iLen; ++ iPos)
	{
		if (bQuoted)
		{
			if (strSource[iPos] == charQuote)
			{
				iPos ++;
				if (iPos == iLen) break; // end quote found
				if (strTokens.Find (strSource[iPos]) >= 0) break; // closing quote found
			}
		}
		else
		{
			if (strTokens.Find (strSource[iPos]) >= 0)
			{
				break;
			}
		}
	}

	bool bEndsWithQuote = bQuoted && (strSource[iPos] == charQuote);
	CString strItem = strSource.Mid (iStart, (bEndsWithQuote ? iPos - 1 : iPos) - iStart);
	if (bQuoted)
	{
		TCHAR singleQuote[] = {charQuote, 0};
		TCHAR doubleQuote[] = {charQuote, charQuote, 0};
		strItem.Replace (doubleQuote, singleQuote);
	}

	iStart = min (iPos + 1, iLen);
	return strItem;
}
//------------------------------------------------------------------------//
HICON CBCGPGlobalUtils::GetWndIcon (CWnd* pWnd, BOOL* bDestroyIcon, BOOL bNoDefault)
{
#ifdef _BCGSUITE_
	UNREFERENCED_PARAMETER(bNoDefault);
#endif
	ASSERT_VALID (pWnd);
	
	if (pWnd->GetSafeHwnd () == NULL)
	{
		return NULL;
	}

	if (bDestroyIcon != NULL)
	{
		*bDestroyIcon = FALSE;
	}

	HICON hIcon = pWnd->GetIcon (FALSE);

	if (hIcon == NULL)
	{
		hIcon = pWnd->GetIcon (TRUE);

		if (hIcon != NULL)
		{
			CImageList il;
			il.Create (::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), ILC_COLOR32 | ILC_MASK, 0, 1);
			il.Add (hIcon);

			if (il.GetImageCount () == 1)
			{
				hIcon = il.ExtractIcon (0);
				if (bDestroyIcon != NULL)
				{
					*bDestroyIcon = hIcon != NULL;
				}
			}
		}
	}

	if (hIcon == NULL)
	{
		hIcon = (HICON)(LONG_PTR)::GetClassLongPtr(pWnd->GetSafeHwnd (), GCLP_HICONSM);
	}

	if (hIcon == NULL)
	{
		hIcon = (HICON)(LONG_PTR)::GetClassLongPtr(pWnd->GetSafeHwnd (), GCLP_HICON);
	}

#ifndef _BCGSUITE_
	if (hIcon == NULL && !bNoDefault && 
		!pWnd->IsKindOf (RUNTIME_CLASS(CDialog)) && !pWnd->IsKindOf (RUNTIME_CLASS(CPropertySheet)))
	{
		hIcon = globalData.m_hiconApp;
		if (bDestroyIcon != NULL)
		{
			*bDestroyIcon = FALSE;
		}
	}
#endif
	return hIcon;
}

HICON CBCGPGlobalUtils::GrayIcon(HICON hIcon)
{
	if (hIcon == NULL)
	{
		return NULL;
	}

#ifdef _BCGSUITE_
	return ::CopyIcon(hIcon);
#else
	HBITMAP hBmp = BCGPIconToBitmap32 (hIcon);

	BITMAP bmp;
	::GetObject (hBmp, sizeof (BITMAP), (LPVOID) &bmp);

	CBCGPToolBarImages images;
	images.SetImageSize (CSize (bmp.bmWidth, bmp.bmHeight));

	images.AddImage(hBmp, TRUE);
	images.ConvertToGrayScale();
	return images.ExtractIcon(0);

#endif
}

CSize CBCGPGlobalUtils::GetIconSize(HICON hIcon)
{
	if (hIcon == NULL)
	{
		return CSize(0, 0);
	}

	ICONINFO info;
	memset(&info, 0, sizeof (ICONINFO));

	::GetIconInfo(hIcon, &info);
	
	BITMAP bmp;
	::GetObject(info.hbmColor, sizeof(BITMAP), (LPVOID)&bmp);
	
	::DeleteObject(info.hbmColor);
	::DeleteObject(info.hbmMask);

	return CSize(bmp.bmWidth, bmp.bmHeight);
}

BOOL CBCGPGlobalUtils::GetParentDialogFont(HWND hWndCtrl, HFONT& hFont, BOOL bCheckParent/* = FALSE*/)
{
#ifndef _BCGSUITE_
	if (!globalData.m_bUseDlgFontInControls)
#else
	if (!m_bUseDlgFontInControls)
#endif
	{
		return FALSE;
	}

	if (hFont != NULL)
	{
		if (::GetObjectType(hFont) == OBJ_FONT)
		{
			return FALSE;	// Already set
		}

		hFont = NULL;
	}

	if (hWndCtrl == NULL || !::IsWindow(hWndCtrl))
	{
		return FALSE;
	}

	CWnd* pWnd = CWnd::FromHandle(hWndCtrl);
	if (pWnd == NULL)
	{
		return FALSE;
	}

	int nSteps = bCheckParent ? 2 : 1;

	for (int i = 0; i < nSteps; i++)
	{
		CWnd* pWndParent = pWnd->GetParent();
		if (pWndParent->GetSafeHwnd() == NULL)
		{
			return FALSE;
		}

		if (DYNAMIC_DOWNCAST(CBCGPDialog, pWndParent) != NULL ||
			DYNAMIC_DOWNCAST(CBCGPPropertyPage, pWndParent) != NULL ||
			DYNAMIC_DOWNCAST(CBCGPPropertySheet, pWndParent) != NULL ||
			DYNAMIC_DOWNCAST(CBCGPFormView, pWndParent) != NULL ||
			DYNAMIC_DOWNCAST(CBCGPDialogBar, pWndParent) != NULL)
		{
			HFONT hFontOld = hFont;

			hFont = (HFONT)::SendMessage(hWndCtrl, WM_GETFONT, 0, 0);
			return hFontOld != hFont;
		}

		pWnd = pWndParent;
	}

	return FALSE;
}

void CBCGPGlobalUtils::EnableWindowShadow(CWnd* pWnd, BOOL bEnable)
{
#ifdef _BCGSUITE_
	UNREFERENCED_PARAMETER(pWnd);
	UNREFERENCED_PARAMETER(bEnable);
#else

	if (!globalData.bIsWindowsVista)
	{
		return;
	}

	if (pWnd->GetSafeHwnd() == NULL)
	{
		return;
	}

	DWORD dwClassStyle = ::GetClassLong(pWnd->GetSafeHwnd(), GCL_STYLE);
	BOOL bHasShadow = (dwClassStyle & CS_DROPSHADOW) != 0;
	BOOL bWasChanged = FALSE;

	if (bEnable)
	{
		if (!bHasShadow)
		{
			::SetClassLong(pWnd->GetSafeHwnd(), GCL_STYLE, dwClassStyle | CS_DROPSHADOW);
			bWasChanged = TRUE;
		}
	}
	else
	{
		if (bHasShadow)
		{
			::SetClassLong(pWnd->GetSafeHwnd(), GCL_STYLE, dwClassStyle & (~CS_DROPSHADOW));
			bWasChanged = TRUE;
		}
	}

	if (bWasChanged && pWnd->IsWindowVisible())
	{
		BOOL bIsZoomed = pWnd->IsZoomed();
		BOOL bIsIconic = pWnd->IsIconic();

		pWnd->ShowWindow(SW_HIDE);
		pWnd->ShowWindow(bIsZoomed ? SW_SHOWMAXIMIZED : bIsIconic ? SW_SHOWMINIMIZED : SW_SHOW);
		pWnd->SetWindowPos(NULL, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

		if (AfxGetMainWnd()->GetSafeHwnd() == pWnd->GetSafeHwnd())
		{
			pWnd->BringWindowToTop();
		}
	}
#endif
}

BOOL CBCGPGlobalUtils::EnableEditCtrlAutoComplete (HWND hwndEdit, BOOL bDirsOnly)
{
#ifndef _BCGSUITE_
	if (globalData.bIsWindows9x || globalData.bIsWindowsNT4)
	{
		return FALSE;
	}
#endif

	typedef HRESULT (CALLBACK* LPFNAUTOCOMPLETE) (HWND ,DWORD);
	enum
	{
		shacf_usetab = 0x00000008,
		shacf_filesys_only = 0x00000010,
		shacf_filesys_dirs = 0x00000020,
		shacf_autosuggest_force_on = 0x10000000
	};

	BOOL bRes = FALSE;

	HINSTANCE hInst = ::LoadLibrary (_T("shlwapi.dll"));
	if(hInst != NULL)
	{
		LPFNAUTOCOMPLETE lpfnAutoComplete = (LPFNAUTOCOMPLETE)::GetProcAddress (hInst, "SHAutoComplete");
		if (lpfnAutoComplete != NULL)
		{
			DWORD dwFlags = shacf_usetab | shacf_autosuggest_force_on;
			dwFlags |= (bDirsOnly) ? shacf_filesys_dirs : shacf_filesys_only;
			
			if (SUCCEEDED(lpfnAutoComplete (hwndEdit, dwFlags)))
			{
				bRes = TRUE;
			}
		}

		::FreeLibrary (hInst);
	}

	return bRes;
}

CSize CBCGPGlobalUtils::GetSystemBorders(CWnd *pWnd)
{
	CSize size(0, 0);

	if (::IsWindow(pWnd->GetSafeHwnd()))
	{
		DWORD dwStyle = pWnd->GetStyle();

		if (pWnd->IsKindOf(RUNTIME_CLASS(CBCGPDialog)))
		{
			dwStyle |= ((CBCGPDialog*)pWnd)->m_Impl.m_bHasCaption ? WS_CAPTION : 0;
		}
		else if (pWnd->IsKindOf(RUNTIME_CLASS(CBCGPPropertySheet)))
		{
			dwStyle |= ((CBCGPPropertySheet*)pWnd)->m_Impl.m_bHasCaption ? WS_CAPTION : 0;
		}
#ifndef _BCGSUITE_
		else if (pWnd->IsKindOf(RUNTIME_CLASS(CBCGPMDIChildWnd)))
		{
			dwStyle &= ~WS_CHILD;
		}
#endif
		size = GetSystemBorders(dwStyle);
	}

	return size;
}

CSize CBCGPGlobalUtils::GetSystemBorders(DWORD dwStyle)
{
	CSize size(0, 0);

	BOOL bCaption = (dwStyle & (WS_CAPTION | WS_DLGFRAME)) != 0;

	if ((dwStyle & WS_THICKFRAME) == WS_THICKFRAME)
	{
		if ((dwStyle & WS_CHILD) == 0 && (dwStyle & WS_MINIMIZE) == 0)
		{
			size.cx = ::GetSystemMetrics(SM_CXSIZEFRAME);
			size.cy = ::GetSystemMetrics(SM_CYSIZEFRAME);

			if (!bCaption)
			{
				size.cx -= ::GetSystemMetrics(SM_CXBORDER);
				size.cy -= ::GetSystemMetrics(SM_CYBORDER);
			}

			if (size.cx != 0 && size.cy != 0 && (dwStyle & WS_MINIMIZE) == 0)
			{
				size.cx += ::GetSystemMetrics(SM_CXPADDEDBORDER);
				size.cy += ::GetSystemMetrics(SM_CXPADDEDBORDER);
			}
		}
		else
		{
			size.cx = ::GetSystemMetrics(SM_CXFIXEDFRAME);
			size.cy = ::GetSystemMetrics(SM_CYFIXEDFRAME);
		}
	}
	else if (bCaption || (dwStyle & (WS_BORDER | DS_MODALFRAME)) != 0)
	{
		size.cx = ::GetSystemMetrics(SM_CXFIXEDFRAME);
		size.cy = ::GetSystemMetrics(SM_CYFIXEDFRAME);

		if (bCaption && !IsXPPlatformToolset() && (dwStyle & WS_MINIMIZE) == 0)
		{
			size.cx += ::GetSystemMetrics(SM_CXPADDEDBORDER) + ::GetSystemMetrics(SM_CXBORDER);
			size.cy += ::GetSystemMetrics(SM_CXPADDEDBORDER) + ::GetSystemMetrics(SM_CYBORDER);
		}
	}

	return size;
}

CRuntimeClass* CBCGPGlobalUtils::RuntimeClassFromName(LPCSTR lpszClassName)
{
#if (_MSC_VER <= 1200)

	CRuntimeClass* pClass = NULL;

	// search app specific classes
	AFX_MODULE_STATE* pModuleState = AfxGetModuleState();
	AfxLockGlobals(CRIT_RUNTIMECLASSLIST);
	for (pClass = pModuleState->m_classList; pClass != NULL;
		pClass = pClass->m_pNextClass)
	{
		if (lstrcmpA(lpszClassName, pClass->m_lpszClassName) == 0)
		{
			AfxUnlockGlobals(CRIT_RUNTIMECLASSLIST);
			return pClass;
		}
	}
	AfxUnlockGlobals(CRIT_RUNTIMECLASSLIST);
#ifdef _AFXDLL
	// search classes in shared DLLs
	AfxLockGlobals(CRIT_DYNLINKLIST);
	for (CDynLinkLibrary* pDLL = pModuleState->m_libraryList; pDLL != NULL;
		pDLL = pDLL->m_pNextDLL)
	{
		for (pClass = pDLL->m_classList; pClass != NULL;
			pClass = pClass->m_pNextClass)
		{
			if (lstrcmpA(lpszClassName, pClass->m_lpszClassName) == 0)
			{
				AfxUnlockGlobals(CRIT_DYNLINKLIST);
				return pClass;
			}
		}
	}
	AfxUnlockGlobals(CRIT_DYNLINKLIST);
#endif

	return NULL;

#else

	return CRuntimeClass::FromName(lpszClassName);

#endif
}

CRuntimeClass* CBCGPGlobalUtils::RuntimeClassFromName(LPCWSTR lpszClassName)
{
#if (_MSC_VER <= 1200)

	if(lpszClassName == NULL)
	{
		return NULL;
	}

	int length = (int)wcslen(lpszClassName);
	if (length == 0)
	{
		return NULL;
	}

	LPSTR lpszClassNameA = new char[length + 1];
	AfxW2AHelper(lpszClassNameA, lpszClassName, length);

	CRuntimeClass* pClass = RuntimeClassFromName(lpszClassNameA);

	if (lpszClassNameA != NULL)
	{
		delete [] lpszClassNameA;
	}

	return pClass;

#else

	return CRuntimeClass::FromName(lpszClassName);

#endif
}

BOOL CBCGPGlobalUtils::ProcessMouseWheel(WPARAM wParam, LPARAM lParam)
{
	CPoint ptCursor;
	GetCursorPos(&ptCursor);
	
	HWND hwndCurr = ::WindowFromPoint(ptCursor);
	if (hwndCurr == NULL || !::IsWindow(hwndCurr))
	{
		return FALSE;
	}

	CWnd* pWndCurr = CWnd::FromHandlePermanent(hwndCurr);
	if (pWndCurr->GetSafeHwnd() == NULL)
	{
		return FALSE;
	}

	CWnd* pWndFocus = CWnd::GetFocus();
	if (pWndFocus->GetSafeHwnd() != NULL)
	{
		TCHAR szClass [256];
		::GetClassName(pWndFocus->GetSafeHwnd (), szClass, 255);
		
		CString strClass = szClass;
		
		if (strClass.CompareNoCase (_T("COMBOBOX")) == 0 || 
			strClass.CompareNoCase (WC_COMBOBOXEX) == 0)
		{
			if ((BOOL)pWndFocus->SendMessage(CB_GETDROPPEDSTATE))
			{
				// The focused combo box is dropped-down: pass the mose wheel event to this control:
				pWndFocus->SendMessage(WM_MOUSEWHEEL, wParam, lParam);
				return TRUE;
			}
		}
		else
		{
			CBCGPCalendar* pWndCalendar = DYNAMIC_DOWNCAST(CBCGPCalendar, pWndFocus);
			if (pWndCalendar->GetSafeHwnd() != NULL && pWndCalendar->IsPopup())
			{
				pWndFocus->SendMessage(WM_MOUSEWHEEL, wParam, lParam);
				return TRUE;
			}
		}
	}

#ifndef _BCGSUITE_
#ifndef BCGP_EXCLUDE_RIBBON
	if (DYNAMIC_DOWNCAST(CBCGPRibbonBar, pWndCurr->GetParent()) != NULL)
	{
		return FALSE;
	}
#endif
#endif
	
	pWndCurr->SendMessage(WM_MOUSEWHEEL, wParam, lParam);
	return TRUE;
}

static int inline _ScaleByDPI(int val)
{
	return (int)(.5 + globalData.GetRibbonImageScale() * val);
}

double CBCGPGlobalUtils::ScaleByDPI(double val)
{
	return globalData.GetRibbonImageScale() == 1.0 ? val : (globalData.GetRibbonImageScale() * val);
}

float CBCGPGlobalUtils::ScaleByDPI(float val)
{
	return globalData.GetRibbonImageScale() == 1.0 ? val : (float)(globalData.GetRibbonImageScale() * val);
}

int CBCGPGlobalUtils::ScaleByDPI(int val)
{
	return globalData.GetRibbonImageScale() == 1.0 ? val : _ScaleByDPI(val);
}

CSize CBCGPGlobalUtils::ScaleByDPI(CSize size)
{
	if (globalData.GetRibbonImageScale() != 1.0)
	{
		size.cx = _ScaleByDPI(size.cx);
		size.cy = _ScaleByDPI(size.cy);
	}
	
	return size;
}

CRect CBCGPGlobalUtils::ScaleByDPI(CRect rect)
{
	if (globalData.GetRibbonImageScale() != 1.0)
	{
		rect.left = _ScaleByDPI(rect.left);
		rect.right = _ScaleByDPI(rect.right);
		rect.top = _ScaleByDPI(rect.top);
		rect.bottom = _ScaleByDPI(rect.bottom);
	}
	
	return rect;
}

CSize CBCGPGlobalUtils::ScaleByDPI(CBCGPToolBarImages& images)
{
#if (!defined _BCGSUITE_) || (_MSC_VER >= 1600)
	if (globalData.GetRibbonImageScale() != 1.0 && !images.IsScaled())
	{
		images.SmoothResize(globalData.GetRibbonImageScale());
	}
#endif

	return images.GetImageSize();
}

HICON CBCGPGlobalUtils::ScaleByDPI(HICON hIcon, BOOL bDestroySourceIcon, BOOL bAlphaBlend)
{
#if (!defined _BCGSUITE_) || (_MSC_VER >= 1600)
	if (globalData.GetRibbonImageScale() != 1.0 && hIcon != NULL)
	{
		CBCGPToolBarImages images;
		images.SetImageSize(GetIconSize(hIcon));
		images.AddIcon(hIcon, bAlphaBlend);

		if (bDestroySourceIcon)
		{
			::DestroyIcon(hIcon);
		}
		
		ScaleByDPI(images);
		
		hIcon = images.ExtractIcon(0);
	}
#else
	UNREFERENCED_PARAMETER(bDestroySourceIcon);
	UNREFERENCED_PARAMETER(bAlphaBlend);
#endif
	return hIcon;
}

BOOL CBCGPGlobalUtils::IsPlacedOnActiveMenu(const CWnd* pWnd)
{
	if (pWnd->GetSafeHwnd() == NULL)
	{
		return FALSE;
	}

	if (CBCGPPopupMenu::GetActiveMenu() != NULL)
	{
		BOOL bIsParentMenu = FALSE;
		for (CWnd* pParent = pWnd->GetParent(); pParent != NULL && !bIsParentMenu; pParent = pParent->GetParent())
		{
			bIsParentMenu = pParent->GetSafeHwnd() == CBCGPPopupMenu::GetActiveMenu()->m_hWnd;
		}

		return bIsParentMenu;
	}
	
	return FALSE;
}

#ifndef _BCGSUITE_

#ifndef WM_GETTITLEBARINFOEX
#define WM_GETTITLEBARINFOEX            0x033F
#endif

#ifndef CCHILDREN_TITLEBAR
#define CCHILDREN_TITLEBAR              5
#endif

#if(WINVER < 0x0600)

typedef struct tagTITLEBARINFOEX
{
    DWORD cbSize;
    RECT rcTitleBar;
    DWORD rgstate[CCHILDREN_TITLEBAR + 1];
    RECT rgrect[CCHILDREN_TITLEBAR + 1];
} TITLEBARINFOEX, *PTITLEBARINFOEX, *LPTITLEBARINFOEX;

#endif

CSize CBCGPGlobalUtils::GetCaptionButtonSize(CWnd* pWnd, int nButton)
{
	NONCLIENTMETRICS ncm;
	globalData.GetNonClientMetrics(ncm);

	CSize sizeButton(ncm.iCaptionWidth, ncm.iCaptionHeight);
	
	if (globalData.bIsWindowsVista && pWnd->GetSafeHwnd() != NULL)
	{
		TITLEBARINFOEX info;
		memset(&info, 0, sizeof(info));
		info.cbSize = sizeof(info);
		
		pWnd->SendMessage(WM_GETTITLEBARINFOEX, 0, (LPARAM)&info);

		int nIndex = 5;

		switch (nButton)
		{
		case SC_CLOSE:
			nIndex = 5;
			break;

		case SC_MINIMIZE:
			nIndex = 2;
			break;
			
		case SC_MAXIMIZE:
			nIndex = 3;
			break;
		}		
			
		CRect rectButton = info.rgrect[nIndex];
		
		if (!rectButton.IsRectEmpty())
		{
			sizeButton = rectButton.Size();
		}
	}

	return sizeButton;
}

#endif

#ifndef PW_CLIENTONLY
#define PW_CLIENTONLY           0x00000001
#endif

#ifndef PW_RENDERFULLCONTENT
#define PW_RENDERFULLCONTENT    0x00000002
#endif

BOOL CBCGPGlobalUtils::IsXPPlatformToolset()
{
    return s_bIsXPPlatformToolset;
}
//***********************************************************************************
BOOL CBCGPGlobalUtils::CreateScreenshot(CBitmap& bmpScreenshot, CWnd* pWnd)
{
	if (bmpScreenshot.GetSafeHandle() != NULL)
	{
		bmpScreenshot.DeleteObject();
	}

	if (pWnd == NULL)
	{
		pWnd = AfxGetMainWnd();
	}

	if (pWnd->GetSafeHwnd() == NULL)
	{
		return FALSE;
	}

	CRect rectWnd;
	pWnd->GetWindowRect(rectWnd);

	CRect rectClient;
	pWnd->GetClientRect(rectClient);
	pWnd->ClientToScreen(rectClient);

	rectClient.OffsetRect(-rectWnd.left, -rectWnd.top);
	rectWnd.OffsetRect(-rectWnd.left, -rectWnd.top);

	BOOL bRes = FALSE;

#ifndef _BCGSUITE_
	const BOOL bIsChild = (pWnd->GetStyle() & WS_CHILD);
	BOOL bClearFrame = !bIsChild && globalData.bIsWindows8;
#else
	BOOL bClearFrame = TRUE;
#endif

	CDC dc;
	if (dc.CreateCompatibleDC(NULL))
	{
		LPBYTE pBitsWnd = NULL;
		HBITMAP hBitmap = CBCGPDrawManager::CreateBitmap_32 (rectWnd.Size(), (void**)&pBitsWnd);

		if (hBitmap != NULL)
		{
			HBITMAP hBitmapOld = (HBITMAP)dc.SelectObject(hBitmap);
			if (hBitmapOld != NULL)
			{
				dc.FillRect(rectWnd, &globalData.brWindow);

#ifndef _BCGSUITE_
				BOOL bSendWMPrint = bIsChild;

				if (!bSendWMPrint)
				{
					bSendWMPrint = !globalData.PrintWindow(pWnd->GetSafeHwnd(), dc.GetSafeHdc(), PW_RENDERFULLCONTENT);
				}
#else
				BOOL bSendWMPrint = TRUE;
#endif

				if (bSendWMPrint)
				{
					pWnd->SendMessage(WM_PRINT, (WPARAM)dc.GetSafeHdc(), (LPARAM)(PRF_CLIENT | PRF_CHILDREN | PRF_NONCLIENT | PRF_ERASEBKGND));
				}

				for (int i = 0; i < rectWnd.Width() * rectWnd.Height(); i++)
				{
					int x = i % rectWnd.Width();
					int y = rectWnd.Height() - (int)(i / rectWnd.Width()) - 1;

					if (!bSendWMPrint && bClearFrame && !rectClient.PtInRect(CPoint(x, y)) && pBitsWnd[0] == 0 && pBitsWnd[1] == 0 && pBitsWnd[2] == 0 && pBitsWnd[3] == 255)
					{
						pBitsWnd[0] = pBitsWnd[1] = pBitsWnd[2] = 255;
					}
					else
					{
						pBitsWnd[3] = 255;
					}

					pBitsWnd += 4;
				}
	
				dc.SelectObject (hBitmapOld);

				bRes = TRUE;
			}
	
			bmpScreenshot.Attach(hBitmap);
		}
	}

	return bRes;
}
//***********************************************************************************
int CBCGPGlobalUtils::DrawTextWithHighlightedArea(CDC* pDC, const CString& str, LPRECT lpRect, UINT nFormat, int nHightLightStart, int nHightlightLen, COLORREF clrHighlight, COLORREF clrHighlightText)
{
	ASSERT_VALID(pDC);

	int nRes = pDC->DrawText(str, lpRect, nFormat);
	
	if ((nFormat & DT_SINGLELINE) != 0 && nHightLightStart >= 0 && nHightlightLen > 0 && clrHighlight != CLR_NONE)
	{
		nHightLightStart = bcg_clamp(nHightLightStart, 0, str.GetLength() - 1);
		nHightlightLen = bcg_clamp(nHightlightLen, 0, str.GetLength() - nHightLightStart);
		
		if (nHightlightLen > 0)
		{
			CRect rectText(lpRect);
			CString strMarked = str.Mid(nHightLightStart, nHightlightLen);
			
			if ((nFormat & DT_RIGHT) == DT_RIGHT)
			{
				rectText.left = rectText.right - pDC->GetTextExtent(str).cx;
			}
			else if ((nFormat & DT_CENTER) == DT_CENTER)
			{
				int cxText = pDC->GetTextExtent(str).cx;
				
				rectText.left = rectText.CenterPoint().x - cxText / 2;
				rectText.right = rectText.left + cxText;
			}
			
			if (nHightLightStart > 0)
			{
				rectText.left += pDC->GetTextExtent(str.Left(nHightLightStart)).cx;
			}
			
			if (rectText.left < rectText.right)
			{
				rectText.right = min(rectText.right, rectText.left + pDC->GetTextExtent(strMarked).cx);
				
				CBrush br(clrHighlight);
				pDC->FillRect(rectText, &br);

				COLORREF clrTextOld = CLR_NONE;

				if (clrHighlightText != CLR_NONE)
				{
					clrTextOld = pDC->SetTextColor(clrHighlightText);
				}
				
				pDC->DrawText(strMarked, rectText, nFormat);

				if (clrTextOld != CLR_NONE)
				{
					pDC->SetTextColor(clrTextOld);
				}
			}
		}
	}

	return nRes;
}
//***********************************************************************************
BOOL CBCGPGlobalUtils::ProcessCtrlEditAccelerators(CWnd* pCtrl, UINT nChar, BOOL bIsEditable, BOOL* pbChanged/* = NULL*/)
{
	ASSERT_VALID(pCtrl);
	ASSERT(pCtrl->GetSafeHwnd() != NULL);

	if (pbChanged != NULL)
	{
		*pbChanged = FALSE;
	}

	if (!::IsWindow(pCtrl->GetSafeHwnd()))
	{
		return FALSE;
	}

	CRichEditCtrl* pRichEdit = DYNAMIC_DOWNCAST(CRichEditCtrl, pCtrl);

	BOOL bIsRAlt = (::GetAsyncKeyState (VK_RMENU) & 0x8000) != 0;
	BOOL bIsCtrl = (::GetAsyncKeyState (VK_CONTROL) & 0x8000) && !bIsRAlt;
	BOOL bIsShift = (::GetAsyncKeyState (VK_SHIFT) & 0x8000);
	
	if (bIsCtrl && nChar == _T('A'))
	{
		if (pRichEdit != NULL)
		{
			pRichEdit->SetSel(0, -1);
			return TRUE;
		}

		if ((pCtrl->SendMessage(WM_GETDLGCODE) & DLGC_HASSETSEL) == DLGC_HASSETSEL)
		{
			pCtrl->SendMessage(EM_SETSEL, 0, (LPARAM)-1);
			return TRUE;
		}
	}
	
	if (bIsCtrl && (nChar == _T('C') || nChar == VK_INSERT))
	{
		pCtrl->SendMessage(WM_COPY);
		return TRUE;
	}

	if (bIsCtrl || bIsShift || (pCtrl->SendMessage(WM_GETDLGCODE) & DLGC_WANTARROWS) == DLGC_WANTARROWS)
	{
		BOOL bIsMultilineEdit = (pCtrl->GetStyle() & ES_MULTILINE) == ES_MULTILINE;

		if (nChar == VK_LEFT || nChar == VK_RIGHT || nChar == VK_HOME || nChar == VK_END || (bIsMultilineEdit && (nChar == VK_DOWN || nChar == VK_UP)))
		{
			pCtrl->SendMessage(WM_KEYDOWN, nChar);
			return TRUE;
		}
	}
	
	if (!bIsEditable)
	{
		return bIsCtrl || bIsShift;
	}
	
	if (bIsCtrl && nChar == _T('V') || (bIsShift && nChar == VK_INSERT))
	{
		pCtrl->SendMessage(WM_PASTE);

		if (pbChanged != NULL)
		{
			*pbChanged = TRUE;
		}
		return TRUE;
	}
	
	if (bIsCtrl && nChar == _T('X') || (bIsShift && nChar == VK_DELETE))
	{
		pCtrl->SendMessage(WM_CUT);

		if (pbChanged != NULL)
		{
			*pbChanged = TRUE;
		}
		return TRUE;
	}
	
	if (nChar == VK_DELETE)
	{
		pCtrl->SendMessage(WM_KEYDOWN, VK_DELETE);

		if (pbChanged != NULL)
		{
			*pbChanged = TRUE;
		}
		return TRUE;
	}

	if (bIsCtrl && nChar == _T('Z'))
	{
		if (pRichEdit != NULL)
		{
			pRichEdit->Undo();
		}
		else
		{
			pCtrl->SendMessage (WM_UNDO);
		}

		if (pbChanged != NULL)
		{
			*pbChanged = TRUE;
		}

		return TRUE;
	}

	if (bIsCtrl && nChar == _T('Y') && pRichEdit != NULL)
	{
		pCtrl->SendMessage (EM_REDO);
		return TRUE;
	}

	return FALSE;
}
//***********************************************************************************
void CBCGPGlobalUtils::NotifyReleaseCapture(CWnd* pWnd)
{
	ASSERT_VALID(pWnd);

	CWnd* pWndParent = pWnd->GetParent();

	if (pWndParent->GetSafeHwnd() != NULL)
	{
		NMHDR nmhdr;
		memset(&nmhdr, 0, sizeof(NMHDR));
		
		nmhdr.hwndFrom = pWnd->GetSafeHwnd(); 
		nmhdr.idFrom = pWnd->GetDlgCtrlID(); 
		nmhdr.code = NM_RELEASEDCAPTURE; 
		
		pWndParent->SendMessage(WM_NOTIFY, pWnd->GetDlgCtrlID(), (LPARAM)&nmhdr);
	}
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPAboutDlg dialog

class CBCGPAboutDlg : public CBCGPDialog
{
// Construction
public:
	CBCGPAboutDlg(LPCTSTR lpszAppName, CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CBCGPAboutDlg)
	enum { IDD = IDD_BCGBARRES_ABOUT_DLG };
	CBCGPURLLinkButton	m_wndURL;
	CBCGPButton	m_wndPurchaseBtn;
	CStatic	m_wndAppName;
	CBCGPURLLinkButton	m_wndEMail;
	CString	m_strAppName;
	CString	m_strVersion;
	CString	m_strYear;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBCGPAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CBCGPAboutDlg)
	afx_msg void OnBcgbarresPurchase();
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	CFont				m_fontBold;
	CBCGPToolBarImages	m_Logo;
};

/////////////////////////////////////////////////////////////////////////////
// CBCGPAboutDlg dialog

CBCGPAboutDlg::CBCGPAboutDlg(LPCTSTR lpszAppName, CWnd* pParent /*=NULL*/)
	: CBCGPDialog(CBCGPAboutDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CBCGPAboutDlg)
	m_strAppName = _T("");
	m_strVersion = _T("");
	m_strYear = _T("");
	//}}AFX_DATA_INIT

	m_bIsLocal = TRUE;
	m_strVersion.Format (_T("%d.%d"), _BCGCBPRO_VERSION_MAJOR, _BCGCBPRO_VERSION_MINOR);

	CString strCurrDate = _T(__DATE__);
	m_strYear.Format (_T("1998-%s"), (LPCTSTR)strCurrDate.Right (4));

	m_strAppName = lpszAppName;

#ifndef _BCGSUITE_
	EnableVisualManagerStyle (globalData.m_bUseVisualManagerInBuiltInDialogs, TRUE);
#endif
}

void CBCGPAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CBCGPDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CBCGPAboutDlg)
	DDX_Control(pDX, IDC_BCGBARRES_URL, m_wndURL);
	DDX_Control(pDX, IDC_BCGBARRES_PURCHASE, m_wndPurchaseBtn);
	DDX_Control(pDX, IDC_BCGBARRES_NAME, m_wndAppName);
	DDX_Control(pDX, IDC_BCGBARRES_MAIL, m_wndEMail);
	DDX_Text(pDX, IDC_BCGBARRES_NAME, m_strAppName);
	DDX_Text(pDX, IDC_BCGBARRES_VERSION, m_strVersion);
	DDX_Text(pDX, IDC_BCGBARRES_YEAR, m_strYear);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CBCGPAboutDlg, CBCGPDialog)
	//{{AFX_MSG_MAP(CBCGPAboutDlg)
	ON_BN_CLICKED(IDC_BCGBARRES_PURCHASE, OnBcgbarresPurchase)
	ON_WM_PAINT()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPAboutDlg message handlers

void CBCGPAboutDlg::OnBcgbarresPurchase() 
{
#ifndef _BCGSUITE_
	CString strURL = _T("http://www.bcgsoft.com/register-bcgcbpe.htm");
#else
	CString strURL = _T("http://www.bcgsoft.com/register-bcgsuite.htm");
#endif

	::ShellExecute (NULL, _T("open"), strURL, NULL, NULL, SW_SHOWNORMAL);

	EndDialog (IDOK);
}

BOOL CBCGPAboutDlg::OnInitDialog() 
{
	CBCGPDialog::OnInitDialog();

	//------------------
	// Create bold font:
	//------------------
	CFont* pFont = m_wndAppName.GetFont ();
	ASSERT_VALID (pFont);

	LOGFONT lf;
	memset (&lf, 0, sizeof (LOGFONT));

	pFont->GetLogFont (&lf);

	lf.lfWeight = FW_BOLD;
	m_fontBold.CreateFontIndirect (&lf);

	m_wndAppName.SetFont (&m_fontBold);

	//-----------------------------
	// Setup URL and e-mail fields:
	//-----------------------------
	m_wndEMail.SetURLPrefix (_T("mailto:"));
	m_wndEMail.SetURL (_T("info@bcgsoft.com"));
	m_wndEMail.SizeToContent ();
	m_wndEMail.SetTooltip (_T("Send mail to us"));

	m_wndURL.SizeToContent();

#ifndef _BCGSUITE_
	m_wndURL.m_bDrawFocus = FALSE;
	m_wndURL.m_bAlwaysUnderlineText = FALSE;

	m_wndEMail.m_bDrawFocus = FALSE;
	m_wndEMail.m_bAlwaysUnderlineText = FALSE;
#endif

	//--------------------
	// Set dialog caption:
	//--------------------
	CString strCaption;
	strCaption.Format (_T("About %s"), (LPCTSTR)m_strAppName);

	SetWindowText (strCaption);

	//----------------------------
	// Hide Logo in High DPI mode:
	//----------------------------
	if (globalData.GetRibbonImageScale () > 1.)
	{
		CWnd* pWndLogo = GetDlgItem (IDC_BCGBARRES_DRAW_AREA);
		if (pWndLogo->GetSafeHwnd () != NULL)
		{
			m_Logo.Load (IDB_BCGBARRES_LOGO);

#ifndef _BCGSUITE_
			m_Logo.SetSingleImage (FALSE);
#else
			m_Logo.SetSingleImage();
#endif
			globalUtils.ScaleByDPI(m_Logo);
			pWndLogo->ShowWindow (SW_HIDE);
		}
	}

	//------------------------------------------
	// Hide "Purchase" button in retail version:
	//------------------------------------------
#ifndef _BCGCBPRO_EVAL_
	m_wndPurchaseBtn.EnableWindow (FALSE);
	m_wndPurchaseBtn.ShowWindow (SW_HIDE);
#endif

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void BCGPShowAboutDlg (LPCTSTR lpszAppName)
{
	CBCGPAboutDlg dlg (lpszAppName);
	dlg.DoModal ();
}

void CBCGPAboutDlg::OnPaint() 
{
	CPaintDC dc(this); // device context for painting
	
	if (!m_Logo.IsValid ())
	{
		return;
	}

	CRect rectClient;
	GetClientRect (&rectClient);

	m_Logo.DrawEx (&dc, rectClient, 0);
}

void BCGPShowAboutDlg (UINT uiAppNameResID)
{
	CString strAppName;
	strAppName.LoadString (uiAppNameResID);
	
	BCGPShowAboutDlg (strAppName);
}
