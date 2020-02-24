// PropertySheetCtrl.cpp : implementation file
//
// This is a part of the BCGControlBar Library
// Copyright (C) 1998-2018 BCGSoft Ltd.
// All rights reserved.
//
// This source code can be used, distributed or modified
// only under terms and conditions 
// of the accompanying license agreement.

#include "stdafx.h"
#include "BCGCBPro.h"
#include "PropertySheetCtrl.h"
#include "BCGPPropertyPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPPropertySheetCtrl

IMPLEMENT_DYNAMIC(CBCGPPropertySheetCtrl, CBCGPPropertySheet)

CBCGPPropertySheetCtrl::CBCGPPropertySheetCtrl(UINT nIDCaption, CWnd* pParentWnd, UINT iSelectPage)
	: CBCGPPropertySheet(nIDCaption, pParentWnd, iSelectPage)
{
	m_hAccel = NULL;
	m_bIsAutoDestroy = TRUE;
	m_sizePadding = CSize(-1, -1);
}

CBCGPPropertySheetCtrl::CBCGPPropertySheetCtrl(LPCTSTR pszCaption, CWnd* pParentWnd, UINT iSelectPage)
	: CBCGPPropertySheet(pszCaption, pParentWnd, iSelectPage)
{
	m_hAccel = NULL;
	m_bIsAutoDestroy = TRUE;
	m_sizePadding = CSize(-1, -1);
}

CBCGPPropertySheetCtrl::CBCGPPropertySheetCtrl()
	: CBCGPPropertySheet(_T(""), NULL, 0)
{
	m_hAccel = NULL;
	m_bIsAutoDestroy = TRUE;
	m_sizePadding = CSize(-1, -1);
}

CBCGPPropertySheetCtrl::~CBCGPPropertySheetCtrl()
{
}

BEGIN_MESSAGE_MAP(CBCGPPropertySheetCtrl, CBCGPPropertySheet)
	//{{AFX_MSG_MAP(CBCGPPropertySheetCtrl)
	ON_WM_SIZE()
	ON_WM_NCLBUTTONDOWN()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPPropertySheetCtrl message handlers

void CBCGPPropertySheetCtrl::PostNcDestroy()
{
	// Call the base class routine first
	CBCGPPropertySheet::PostNcDestroy();
	
	if (m_bModeless && m_bIsAutoDestroy)
	{
		delete this;
	}
}

BOOL CBCGPPropertySheetCtrl::OnInitDialog()
{
	ASSERT_VALID(this);
	
	BOOL bRes = CBCGPPropertySheet::OnInitDialog();
	
	ModifyStyleEx(0, WS_EX_CONTROLPARENT);

	m_bDrawPageFrame = FALSE;

	ResizeControl();

	return bRes;
}

BOOL CBCGPPropertySheetCtrl::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	NMHDR* pNMHDR = (NMHDR*) lParam;
	ASSERT(pNMHDR != NULL);

	if (pNMHDR->code == TCN_SELCHANGE)
	{
		ResizeControl();
	}
	
	return CBCGPPropertySheet::OnNotify(wParam, lParam, pResult);
}

//////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////
void CBCGPPropertySheetCtrl::LoadAcceleratorTable(UINT nAccelTableID /*=0*/)
{
	if (nAccelTableID)
	{
		m_hAccel = ::LoadAccelerators(AfxGetInstanceHandle(), MAKEINTRESOURCE(nAccelTableID));
		ASSERT(m_hAccel);
	}
}

//////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////
BOOL CBCGPPropertySheetCtrl::PreTranslateMessage(MSG* pMsg)
{
	if (m_hAccel != NULL && ::TranslateAccelerator(m_hWnd, m_hAccel, pMsg))
	{
		return TRUE;
	}

	return CBCGPPropertySheet::PreTranslateMessage(pMsg);
}

CWnd* CBCGPPropertySheetCtrl::DoCreateOnPlaceHolder(CWnd* pParent, UINT nID, DWORD dwStyle, const CRect& rect, LPVOID /*lParam*/)
{
	if (!Create(pParent, dwStyle | WS_CLIPSIBLINGS, nID))
	{
		return NULL;
	}
	
	SetWindowPos(NULL, rect.left, rect.top, rect.Width(), rect.Height(), SWP_NOACTIVATE | SWP_NOZORDER);
	return this;
}

void CBCGPPropertySheetCtrl::OnSize(UINT nType, int cx, int cy) 
{
	SetRedraw(FALSE);

	CBCGPPropertySheet::OnSize(nType, cx, cy);

	ResizeControl();

	if (m_wndTab.GetSafeHwnd() != NULL)
	{
		CPropertyPage* pPage = GetActivePage();
		if (pPage != NULL)
		{
			OnActivatePage(pPage);
		}
	}

	CBCGPPropertyPage* pPage = DYNAMIC_DOWNCAST(CBCGPPropertyPage, GetActivePage());
	if (pPage != NULL)
	{
		pPage->AdjustControlsLayout();
	}

	SetRedraw(TRUE);
	RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
}

void CBCGPPropertySheetCtrl::ResizeControl()
{
	CWnd* pTabCtrl = GetTabControl();
	if (m_wndTab.GetSafeHwnd() != NULL)
	{
		pTabCtrl = &m_wndTab;
		m_wndTab.SetButtonsVisible(FALSE);
	}

	if (pTabCtrl->GetSafeHwnd() == NULL)
	{
		return;
	}

	const int nPageCount = GetPageCount();

	if (m_sizePadding == CSize(-1, -1))
	{
		if (m_look == PropSheetLook_Tabs && m_wndTab.GetSafeHwnd() == NULL && GetTabControl()->GetSafeHwnd() != NULL)
		{
			if (nPageCount > 0)
			{
				CRect rectTab;
				pTabCtrl->GetClientRect(rectTab);

				CRect rectTabs;
				if (GetTabControl()->GetItemRect(0, rectTabs))
				{
					rectTab.top = rectTabs.bottom + 2;

					CPropertyPage* pPage = GetPage(0);
					if (pPage->GetSafeHwnd() != NULL)
					{
						CRect rectPage;
						pPage->GetClientRect(rectPage);

						m_sizePadding.cx = rectTab.Width() - rectPage.Width();
						m_sizePadding.cy = rectTab.Height() - rectPage.Height();
					}
				}
			}
		}
		else
		{
			m_sizePadding = CSize(0, 0);
		}
	}

	CRect rectClient;
	GetClientRect(rectClient);

	pTabCtrl->SetWindowPos(NULL, 
		0, 0, rectClient.Width(), rectClient.Height(), 
		SWP_NOZORDER | SWP_NOACTIVATE);

	for (int nPage = 0; nPage <= nPageCount - 1; nPage++)
	{
		CPropertyPage* pPage = GetPage(nPage);
		if (pPage->GetSafeHwnd() != NULL)
		{
			CRect rectPage;

			pPage->GetWindowRect(&rectPage);
			pTabCtrl->ScreenToClient(rectPage);

			pPage->SetWindowPos(NULL, 
				rectPage.left, rectPage.top, 
				rectClient.Width() - m_sizePadding.cx, 
				rectClient.Height() - rectPage.top - m_sizePadding.cy,
				SWP_NOZORDER | SWP_NOACTIVATE);
		}
	}

	RedrawWindow();
}

void CBCGPPropertySheetCtrl::OnNcLButtonDown(UINT nHitTest, CPoint point) 
{
	if (IsDragClientAreaEnabled() && nHitTest == HTCAPTION)
	{
		CWnd* pParent = GetParent();
		if (pParent->GetSafeHwnd() != NULL)
		{
			pParent->SendMessage(WM_NCLBUTTONDOWN, (WPARAM) HTCAPTION, MAKELPARAM(point.x, point.y));
		}

		return;
	}

	CBCGPPropertySheet::OnNcLButtonDown(nHitTest, point);
}
