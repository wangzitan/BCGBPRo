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
// BCGPDragFrameImpl.cpp: implementation of the CBCGPDragFrameImpl class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "bcgcbpro.h"

#include "BCGPDragFrameImpl.h"

#include "BCGPTabbedControlBar.h"
#include "BCGPMiniFrameWnd.h"
#include "BCGPDockManager.h"
#include "BCGGlobals.h"
#include "BCGPGlobalUtils.h"
#include "BCGPMultiMiniFrameWnd.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

class CBCGPDummyDockingControlBar : public CBCGPDockingControlBar
{
	virtual void DoPaint(CDC* /*pDC*/) {}

protected:
	afx_msg BOOL OnEraseBkgnd(CDC* /*pDC*/) {return FALSE;}
	DECLARE_MESSAGE_MAP()
};


BEGIN_MESSAGE_MAP(CBCGPDummyDockingControlBar, CBCGPDockingControlBar)
	ON_WM_ERASEBKGND()
END_MESSAGE_MAP()

static UINT BCGP_DUMMY_WND_ID = (UINT) -2;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPDragFrameImpl::CBCGPDragFrameImpl()
{
	m_rectDrag.SetRectEmpty ();
	m_rectExpectedDocked.SetRectEmpty ();
	m_ptHot.x = m_ptHot.y = 0;
	m_pDraggedWnd = NULL;
	m_pDockManager = NULL;
	m_pTargetBar = NULL;
	m_nInsertedTabID = -1;
	m_bDockToTab = FALSE;
	m_pFinalTargetBar = NULL;
	m_bDragStarted = FALSE;
	m_bFrameTabDrawn = FALSE;
	m_pOldTargetBar = NULL;
	m_pWndDummy = NULL;
	m_pDragRectWnd = NULL;
}

CBCGPDragFrameImpl::~CBCGPDragFrameImpl()
{
	if (m_pWndDummy != NULL)
	{
		m_pWndDummy->DestroyWindow ();
		delete m_pWndDummy;
	}

	if (m_pDragRectWnd != NULL)
	{
		m_pDragRectWnd->DestroyWindow();
	}
}
//--------------------------------------------------------------------------------------//
void CBCGPDragFrameImpl::Init (CWnd* pDraggedWnd)
{
	ASSERT_VALID (pDraggedWnd);
	m_pDraggedWnd = pDraggedWnd;

	CWnd* pDockSite = NULL;
	if (m_pDraggedWnd->IsKindOf (RUNTIME_CLASS (CBCGPMiniFrameWnd)))
	{
		CBCGPMiniFrameWnd* pMiniFrame = DYNAMIC_DOWNCAST (CBCGPMiniFrameWnd, m_pDraggedWnd);
		pDockSite = pMiniFrame->GetParent ();
	}
	else if (m_pDraggedWnd->IsKindOf (RUNTIME_CLASS (CBCGPControlBar)))
	{
		CBCGPControlBar* pBar = 
				DYNAMIC_DOWNCAST (CBCGPControlBar, m_pDraggedWnd);
		ASSERT_VALID (pBar);

		CBCGPMiniFrameWnd* pParentMiniFrame = pBar->GetParentMiniFrame ();
		if (pParentMiniFrame != NULL)
		{
			pDockSite = pParentMiniFrame->GetParent ();
		}
		else
		{
			pDockSite = pBar->GetDockSite ();
		}
	}

	m_pDockManager = globalUtils.GetDockManager (pDockSite);
	if (globalUtils.m_bDialogApp)
	{
		return;
	}

	ASSERT(m_pDockManager != NULL);
}
//--------------------------------------------------------------------------------------//
void CBCGPDragFrameImpl::MoveDragFrame (BOOL bForceMove)
{
	ASSERT_VALID (m_pDraggedWnd);

	m_pFinalTargetBar = NULL;

	if (m_pDraggedWnd == NULL || m_pDockManager == NULL)
	{
		return;
	}

	if (m_pWndDummy == NULL)
	{
		m_pWndDummy = new CBCGPDummyDockingControlBar;
		m_pWndDummy->CreateEx (0, _T (""), BCGCBProGetTopLevelFrame (m_pDraggedWnd), CRect (0, 0, 0, 0), 
							FALSE, BCGP_DUMMY_WND_ID, WS_CHILD);
	}

	CSize szSencitivity = CBCGPDockingControlBar::GetDragSencitivity ();

	CPoint ptMouse;
	GetCursorPos (&ptMouse);

	CPoint ptOffset = ptMouse - m_ptHot;

	if (abs (ptOffset.x) < szSencitivity.cx && 
		abs (ptOffset.y) < szSencitivity.cy && 
		m_rectDrag.IsRectEmpty () && !bForceMove)
	{
		return;
	}

	m_bDragStarted = TRUE;
	
	m_pDockManager->LockUpdate (TRUE);
		
	CRect rectOld = m_rectExpectedDocked.IsRectEmpty () ? m_rectDrag : m_rectExpectedDocked;
	BOOL bFirstTime = FALSE;

	if (m_rectDrag.IsRectEmpty ())
	{
		if (m_pDraggedWnd->IsKindOf (RUNTIME_CLASS (CBCGPMiniFrameWnd)))
		{
			m_pDraggedWnd->GetWindowRect (m_rectDrag);
		}
		else if (m_pDraggedWnd->IsKindOf (RUNTIME_CLASS (CBCGPControlBar)))
		{
			CBCGPControlBar* pBar = 
				DYNAMIC_DOWNCAST (CBCGPControlBar, m_pDraggedWnd);
			ASSERT_VALID (pBar);
			m_pDraggedWnd->GetWindowRect (m_rectDrag);

			// if the bar is docked then the floating rect has to be set to recent floating rect
			if (pBar->GetParentMiniFrame () == NULL)
			{
				m_rectDrag.right = 
					m_rectDrag.left + pBar->m_recentDockInfo.m_rectRecentFloatingRect.Width ();
				m_rectDrag.bottom = 
					m_rectDrag.top + pBar->m_recentDockInfo.m_rectRecentFloatingRect.Height ();
			}

			if (!m_rectDrag.PtInRect (m_ptHot))
			{
				int nOffset = m_rectDrag.left - m_ptHot.x;
				m_rectDrag.OffsetRect (-nOffset - 5, 0); // offset of mouse pointer 
														 // from the drag rect bound
			}
		}
		bFirstTime = TRUE;
	}


	BOOL bDrawTab = FALSE;
	CBCGPDockingControlBar* pOldTargetBar = m_pTargetBar;
	CRect rectExpected; rectExpected.SetRectEmpty ();

    CBCGPSmartDockingManager* pSDManager = NULL;
    BOOL bSDockingIsOn = FALSE;

    if (m_pDockManager != NULL
        && (pSDManager = m_pDockManager->GetSDManagerPermanent()) != NULL
        && pSDManager->IsStarted())
    {
        bSDockingIsOn = TRUE;
    }

	
	m_pDockManager->CalcExpectedDockedRect (m_pDraggedWnd, ptMouse, 
							rectExpected, bDrawTab, &m_pTargetBar);

	if (pOldTargetBar != NULL && m_nInsertedTabID != -1 && 
		(pOldTargetBar != m_pTargetBar || !bDrawTab))
	{
        RemoveTabPreDocking (pOldTargetBar);
		bFirstTime = TRUE;
	}

	BOOL bCanBeAttached = TRUE;
	if (m_pDraggedWnd->IsKindOf (RUNTIME_CLASS (CBCGPMiniFrameWnd)))
	{
	}
	else if (m_pDraggedWnd->IsKindOf (RUNTIME_CLASS (CBCGPControlBar)))
	{
		CBCGPControlBar* pBar = 
			DYNAMIC_DOWNCAST (CBCGPControlBar, m_pDraggedWnd);
		bCanBeAttached = pBar->CanBeAttached ();
	}

	if (m_pTargetBar != NULL && bCanBeAttached)
	{
		CBCGPBaseTabbedBar* pTabbedBar = 
				DYNAMIC_DOWNCAST (CBCGPBaseTabbedBar, m_pTargetBar);
		if (pTabbedBar != NULL && bDrawTab && 
			 (pTabbedBar->GetVisibleTabsNum () > 1 && pTabbedBar->IsHideSingleTab () ||
			  pTabbedBar->GetVisibleTabsNum () > 0 && !pTabbedBar->IsHideSingleTab ()))
		{
			PlaceTabPreDocking (pTabbedBar, bFirstTime);
			return;
		}
		else if (bDrawTab)
		{
			if (m_nInsertedTabID != -1)
			{
				return;
			}
			if (!bFirstTime)
			{
				EndDrawDragFrame (FALSE);
			}
			DrawFrameTab (m_pTargetBar, FALSE);
			m_nInsertedTabID = 1;
			return;
		}
	}	
	
	m_rectDrag.OffsetRect (ptOffset);
	m_ptHot = ptMouse;

	m_rectExpectedDocked = rectExpected;

	CRect rectDocked;
	if (m_rectExpectedDocked.IsRectEmpty ())
	{
		if (!m_rectDrag.PtInRect (ptMouse))
		{
			CPoint ptMiddleRect (m_rectDrag.TopLeft ().x + m_rectDrag.Width () / 2, 
								 m_rectDrag.top + 5);

			ptOffset = ptMouse - ptMiddleRect;
			m_rectDrag.OffsetRect (ptOffset);
		}
		rectDocked = m_rectDrag;
	}
	else
	{
		rectDocked = m_rectExpectedDocked;
	}
	
	if (!bSDockingIsOn || !m_rectExpectedDocked.IsRectEmpty ())
	{
		DrawDragFrame (rectDocked, CBCGPSDPlaceMarkerWnd::BCGPSDPlaceMarkerWnd_Frame);
	}
}
//--------------------------------------------------------------------------------------//
void CBCGPDragFrameImpl::DrawFrameTab (CBCGPDockingControlBar* pTargetBar, BOOL bErase)
{
    CBCGPSmartDockingManager* pSDManager = NULL;
    BOOL bSDockingIsOn = FALSE;

    if (m_pDockManager != NULL
        && (pSDManager = m_pDockManager->GetSDManagerPermanent()) != NULL
        && pSDManager->IsStarted())
    {
        bSDockingIsOn = TRUE;
    }

	if (bErase)
	{
        if (bSDockingIsOn)
        {
            pSDManager->HidePlace ();
        } 
		else
        {
			if (m_pDragRectWnd->GetSafeHwnd() != NULL)
			{
				m_pDragRectWnd->DestroyWindow();
				m_pDragRectWnd = NULL;
			}

			m_bFrameTabDrawn = FALSE;
        }
	}
	else
	{
		CRect rectWnd;
		pTargetBar->GetWindowRect (rectWnd);
		
        if (bSDockingIsOn)
        {
            pSDManager->ShowTabbedPlaceAt (&rectWnd, 0, 0, 0);
        } 
		else
        {
			DrawDragFrame (rectWnd, CBCGPSDPlaceMarkerWnd::BCGPSDPlaceMarkerWnd_FrameWithTab);
			m_bFrameTabDrawn = TRUE;
        }
	}
}
//--------------------------------------------------------------------------------------//
void CBCGPDragFrameImpl::EndDrawDragFrame (BOOL bClearInternalRects)
{
	if (m_pDockManager == NULL)
	{
		return;
	}

    BOOL bSDockingIsOn = FALSE;
    CBCGPSmartDockingManager* pSDManager = NULL;

    if ((pSDManager = m_pDockManager->GetSDManagerPermanent()) != NULL && pSDManager->IsStarted())
    {
        bSDockingIsOn = TRUE;
        pSDManager->HidePlace ();
    }

	if (m_nInsertedTabID != -1)
	{
		m_bDockToTab = TRUE;
	}
	
	if (bClearInternalRects)
	{
        RemoveTabPreDocking ();

		m_rectExpectedDocked.SetRectEmpty ();
		m_rectDrag.SetRectEmpty ();

		m_pFinalTargetBar = m_pTargetBar;
		m_pTargetBar = NULL;
	}

	if (m_pDragRectWnd->GetSafeHwnd() != NULL)
	{
		m_pDragRectWnd->DestroyWindow();
		m_pDragRectWnd = NULL;
	}

	m_bDragStarted = FALSE;
	
	ASSERT (m_pDockManager != NULL);
	if (!bSDockingIsOn)
	{
		m_pDockManager->LockUpdate (FALSE);
	}
}
//--------------------------------------------------------------------------------------//
void CBCGPDragFrameImpl::DrawDragFrame(const CRect& rect, CBCGPSDPlaceMarkerWnd::BCGPSDPlaceMarkerWnd_STYLE style)
{
    CBCGPSmartDockingManager* pSDManager = NULL;

    if (m_pDockManager != NULL
        && (pSDManager = m_pDockManager->GetSDManagerPermanent()) != NULL
        && pSDManager->IsStarted())
    {
        pSDManager->ShowPlaceAt(rect);
		return;
    } 

	if (m_pDragRectWnd == NULL)
	{
		if (!rect.IsRectEmpty())
		{
			m_pDragRectWnd = new CBCGPSDPlaceMarkerWnd(style);

			if (!m_pDragRectWnd->Create(rect, m_pDraggedWnd))
			{
				ASSERT(FALSE);
			}
		}
	}
	else
	{
		if (rect.IsRectEmpty())
		{
			m_pDragRectWnd->DestroyWindow();
			m_pDragRectWnd = NULL;
		}
		else
		{
			m_pDragRectWnd->SetWindowPos(&CWnd::wndTopMost, rect.left, rect.top, rect.Width(), rect.Height(),
				SWP_NOACTIVATE | SWP_NOOWNERZORDER);
		}
	}
}
//******************************************************************************
void CBCGPDragFrameImpl::PlaceTabPreDocking (CBCGPBaseTabbedBar* pTabbedBar, BOOL bFirstTime)
{
	if (m_nInsertedTabID != -1)
	{
		return;
	}
	if (!bFirstTime)
	{
		EndDrawDragFrame (FALSE);
	}
	CString strLabel;
	if (m_pDraggedWnd->IsKindOf (RUNTIME_CLASS (CBCGPMultiMiniFrameWnd)))
	{
		CBCGPMultiMiniFrameWnd* pMultiMiniFrame = 
			DYNAMIC_DOWNCAST (CBCGPMultiMiniFrameWnd, m_pDraggedWnd);
		if (pMultiMiniFrame != NULL)
		{
			CWnd* pBar = pMultiMiniFrame->GetFirstVisibleBar ();
			ASSERT_VALID (pBar);

			if (pBar != NULL)
			{
				pBar->GetWindowText (strLabel);
			}
		}
	}
	else
	{
		m_pDraggedWnd->GetWindowText (strLabel);
	}	
	
	if (m_pWndDummy == NULL)
	{
		m_pWndDummy = new CBCGPDummyDockingControlBar;
		m_pWndDummy->CreateEx (0, _T (""), BCGCBProGetTopLevelFrame (m_pDraggedWnd), CRect (0, 0, 0, 0), 
							FALSE, BCGP_DUMMY_WND_ID, WS_CHILD);
	}

	pTabbedBar->GetUnderlinedWindow ()->AddTab (m_pWndDummy, strLabel);
	
	CBCGPSmartDockingManager* pSDManager = NULL;
	if ((pSDManager = m_pDockManager->GetSDManagerPermanent()) != NULL
		&& pSDManager->IsStarted())
	{
		m_pDraggedWnd->ShowWindow (SW_HIDE);
	}

	m_nInsertedTabID = pTabbedBar->GetUnderlinedWindow ()->GetTabFromHwnd (*m_pWndDummy);
	m_pOldTargetBar = pTabbedBar;
}
//******************************************************************************
void CBCGPDragFrameImpl::PlaceTabPreDocking (CWnd* pCBarToPlaceOn)
{
	CBCGPBaseTabbedBar* pTabbedBar = 
			DYNAMIC_DOWNCAST (CBCGPBaseTabbedBar, pCBarToPlaceOn);
	if (pTabbedBar != NULL &&
		 (pTabbedBar->GetVisibleTabsNum () > 1 && pTabbedBar->IsHideSingleTab () ||
		  pTabbedBar->GetVisibleTabsNum () > 0 && !pTabbedBar->IsHideSingleTab ()))
	{
		m_pTargetBar = pTabbedBar;
		PlaceTabPreDocking (pTabbedBar, TRUE);
		return;
	}
	else if (m_nInsertedTabID == -1)
	{
		CBCGPDockingControlBar* pDockingControlBar = DYNAMIC_DOWNCAST (CBCGPDockingControlBar, pCBarToPlaceOn);
		if (pDockingControlBar != NULL)
		{
			DrawFrameTab (pDockingControlBar, FALSE);
			m_pTargetBar = pDockingControlBar;
			m_pOldTargetBar = pDockingControlBar;
			m_nInsertedTabID = 1;
		}
	}
}
//******************************************************************************
void CBCGPDragFrameImpl::RemoveTabPreDocking (CBCGPDockingControlBar* pOldTargetBar)
{
	if (pOldTargetBar == NULL)
	{
		pOldTargetBar = m_pOldTargetBar;
	}

	if (pOldTargetBar != NULL && m_nInsertedTabID != -1)
	{
	    CBCGPBaseTabbedBar* pOldTabbedBar = 
			    DYNAMIC_DOWNCAST (CBCGPBaseTabbedBar, pOldTargetBar);
		if (pOldTabbedBar != NULL && !m_bFrameTabDrawn && m_pWndDummy != NULL && m_pWndDummy->GetSafeHwnd () != NULL)
		{
			CBCGPSmartDockingManager* pSDManager = NULL;
			BOOL bSDockingIsOn = FALSE;

			if (m_pDockManager != NULL
				&& (pSDManager = m_pDockManager->GetSDManagerPermanent()) != NULL
				&& pSDManager->IsStarted())
			{
				bSDockingIsOn = TRUE;
			}

			m_pWndDummy->ShowWindow (SW_HIDE);
			if (!bSDockingIsOn)
			{
				m_pDockManager->LockUpdate (FALSE);
			}
			CWnd* pWnd = pOldTabbedBar->GetUnderlinedWindow ()->GetTabWnd (m_nInsertedTabID);
			if (pWnd == m_pWndDummy)
			{
				pOldTabbedBar->GetUnderlinedWindow ()->RemoveTab (m_nInsertedTabID);
			}
			if (!bSDockingIsOn)
			{
				m_pDockManager->LockUpdate (TRUE);
			}
		}
		else
		{
			DrawFrameTab (pOldTargetBar, TRUE);
		}

		CBCGPSmartDockingManager* pSDManager = NULL;

		if ((pSDManager = m_pDockManager->GetSDManagerPermanent()) != NULL
			&& pSDManager->IsStarted())
		{
			m_pDraggedWnd->ShowWindow (SW_SHOW);
		}
	}

	m_nInsertedTabID = -1;
	m_pOldTargetBar = NULL;
}
//******************************************************************************
void CBCGPDragFrameImpl::ResetState ()
{
	m_ptHot = CPoint (-1, -1);
	m_rectDrag.SetRectEmpty ();
	m_rectExpectedDocked.SetRectEmpty ();

	m_pFinalTargetBar = NULL;
	m_pOldTargetBar   = NULL;
	m_bDockToTab	  = FALSE;
	m_bDragStarted    = FALSE;

	m_nInsertedTabID  = -1;
}