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
// BCGPRibbonDialogBar.cpp : implementation file
//

#include "stdafx.h"

#include "BCGPRibbonDialogBar.h"
#include "BCGPRibbonCategory.h"
#include "BCGPRibbonPanel.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#ifndef BCGP_EXCLUDE_RIBBON

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonDialogBar

IMPLEMENT_DYNCREATE(CBCGPRibbonDialogBar, CBCGPDockingControlBar)

CBCGPRibbonDialogBar::CBCGPRibbonDialogBar(CRuntimeClass* pRTI) :
	m_wndRibbon (FALSE)
{
	m_pCategory = NULL;
	m_nImagesLarge = 0;
	m_nImagesSmall = 0;

	if (pRTI != NULL)
	{
		ASSERT(pRTI->IsDerivedFrom(RUNTIME_CLASS(CBCGPRibbonBar)));
		m_pCustomRibbon = DYNAMIC_DOWNCAST(CBCGPRibbonBar, pRTI->CreateObject());
	}
	else
	{
		m_pCustomRibbon = NULL;
	}
}

CBCGPRibbonDialogBar::~CBCGPRibbonDialogBar()
{
	if (m_pCustomRibbon != NULL)
	{
		delete m_pCustomRibbon;
	}
}


BEGIN_MESSAGE_MAP(CBCGPRibbonDialogBar, CBCGPDockingControlBar)
	//{{AFX_MSG_MAP(CBCGPRibbonDialogBar)
	ON_WM_CREATE()
	ON_WM_ERASEBKGND()
	ON_WM_SIZE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


BOOL CBCGPRibbonDialogBar::Create(LPCTSTR lpszCaption, 
									CWnd* pParentWnd, 
									const RECT& rect, 
									UINT nImagesSmall,
									UINT nImagesLarge,
									BOOL bHasGripper, 
									UINT nID, 
									DWORD dwStyle, 
									DWORD dwTabbedStyle,
									DWORD dwBCGStyle,
									CCreateContext* pContext)
{
	ASSERT_VALID (this);

	m_nImagesSmall = nImagesSmall;
	m_nImagesLarge = nImagesLarge;

	if (!CBCGPDockingControlBar::CreateEx (0, lpszCaption, pParentWnd, rect, 
											 bHasGripper, nID, dwStyle, dwTabbedStyle, 
											 dwBCGStyle, pContext))
	{
		return FALSE;
	}

	SetOwner (BCGCBProGetTopLevelFrame (this));

	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonDialogBar message handlers



int CBCGPRibbonDialogBar::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CBCGPDockingControlBar::OnCreate(lpCreateStruct) == -1)
		return -1;

	CBCGPRibbonBar& wndRibbon = GetRibbon();
	
	if (!wndRibbon.Create (this, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0))
	{
		return -1;
	}

	m_pCategory = wndRibbon.AddCategory (_T("Main"), m_nImagesSmall, m_nImagesLarge);
	if (m_pCategory == NULL)
	{
		ASSERT(FALSE);
		return -1;
	}

	ASSERT_VALID (m_pCategory);

	m_pCategory->m_bOnDialogBar = TRUE;
	m_pCategory->m_ScrollLeft.m_bIsVertical = m_pCategory->m_ScrollRight.m_bIsVertical = TRUE;

	return 0;
}
//********************************************************************************
BOOL CBCGPRibbonDialogBar::OnEraseBkgnd(CDC* /*pDC*/) 
{
	return TRUE;
}
//********************************************************************************
void CBCGPRibbonDialogBar::OnSize(UINT nType, int cx, int cy) 
{
	CBCGPDockingControlBar::OnSize(nType, cx, cy);

	CBCGPRibbonBar& wndRibbon = GetRibbon();

	if (wndRibbon.GetSafeHwnd () != NULL)
	{
		wndRibbon.SetWindowPos (NULL, 0, 0, cx, cy, SWP_NOZORDER | SWP_NOACTIVATE);
		wndRibbon.RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
	}
}
//********************************************************************************
CBCGPRibbonPanel* CBCGPRibbonDialogBar::AddPanel (LPCTSTR lpszPanelName, HICON hIcon,
		int nRows, CRuntimeClass* pRTI, BOOL bAutoDestroyIcon)
{
	ASSERT_VALID (this);

	if (m_pCategory == NULL)
	{
		ASSERT (FALSE);
		return NULL;
	}

	ASSERT_VALID (m_pCategory);
	
	CBCGPRibbonPanel* pPanel = m_pCategory->AddPanel (lpszPanelName, hIcon, pRTI, bAutoDestroyIcon);
	if (pPanel == NULL)
	{
		ASSERT(FALSE);
		return NULL;

	}

	ASSERT_VALID (pPanel);
	
	pPanel->m_nMaxRows = nRows;
	pPanel->m_bOnDialogBar = TRUE;

	return pPanel;
}

#endif // BCGP_EXCLUDE_RIBBON
