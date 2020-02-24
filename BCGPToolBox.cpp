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

// BCGPToolBox.cpp : implementation file
//

#include "stdafx.h"
#include "bcgcbpro.h"

#ifndef BCGP_EXCLUDE_TOOLBOX

#include "BCGPToolBox.h"
#include "BCGGlobals.h"
#include "BCGPVisualManager.h"
#include "BCGPLocalResource.h"
#include "BCGProRes.h"
#include "BCGPGlobalUtils.h"
#include "BCGPDlgImpl.h"

#ifndef _BCGSUITE_
#include "BCGPOutlookWnd.h"
#include "BCGPButton.h"
#include "BCGPPopupMenu.h"
#include "BCGPTooltipManager.h"
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define ID_FILTER	1

/////////////////////////////////////////////////////////////////////////////
// CBCGPToolBoxPage

#define TOOLBOX_IMAGE_MARGIN globalUtils.ScaleByDPI(4)

IMPLEMENT_DYNCREATE(CBCGPToolBoxPage, CBCGPControlBar)

CBCGPToolBoxPage::CBCGPToolBoxPage()
{
	m_nCheckedButton = -1;
	m_nHighlightedButton = -1;
	m_sizeImage = CSize (0, 0);
	m_pToolBox = NULL;
	m_pToolBoxEx = NULL;
	m_nVertScrollOffset = 0;
	m_nVertScrollSize = 0;
	m_nScrollValue = 0;
	m_bCheckFirstButton = TRUE;
	m_Mode = ToolBoxPageMode_Default;
	m_pButtonClass = NULL;
	m_pToolTip = NULL;
	m_uiBmpResID = 0;
}

CBCGPToolBoxPage::~CBCGPToolBoxPage()
{
	for (int i = 0; i < m_arButtons.GetSize (); i++)
	{
		delete m_arButtons [i];
	}
}

BEGIN_MESSAGE_MAP(CBCGPToolBoxPage, CBCGPControlBar)
	//{{AFX_MSG_MAP(CBCGPToolBoxPage)
	ON_WM_ERASEBKGND()
	ON_WM_SIZE()
	ON_WM_NCDESTROY()
	ON_WM_LBUTTONDBLCLK()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_MOUSEMOVE()
	ON_WM_DESTROY()
	ON_WM_CONTEXTMENU()
	ON_WM_SYSCOLORCHANGE()
	ON_WM_MOUSEWHEEL()
	//}}AFX_MSG_MAP
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnNeedTipText)
	ON_REGISTERED_MESSAGE(BCGM_UPDATETOOLTIPS, OnBCGUpdateToolTips)
END_MESSAGE_MAP()

BOOL CBCGPToolBoxPage::OnEraseBkgnd(CDC* /*pDC*/) 
{
	return  TRUE;
}
//*************************************************************************************
void CBCGPToolBoxPage::OnDestroy() 
{
	CBCGPTooltipManager::DeleteToolTip (m_pToolTip);
	CBCGPControlBar::OnDestroy();
}
//*************************************************************************************
void CBCGPToolBoxPage::DoPaint(CDC* pDCPaint)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDCPaint);

	CBCGPMemDC memDC (*pDCPaint, this);
	CDC* pDC = &memDC.GetDC ();

	CRect rectClient;
	GetClientRect (rectClient);

	CBCGPVisualManager::GetInstance ()->OnFillBarBackground (pDC, this,
		rectClient, rectClient);

#ifndef _BCGSUITE_
	if (m_Images.GetBitsPerPixel() < 32)
	{
		m_Images.SetTransparentColor (globalData.clrBtnFace);
	}

	CBCGPDrawState ds(CBCGPVisualManager::GetInstance()->IsAutoGrayscaleImages());
	m_Images.PrepareDrawImage (ds);
#else
	m_Images.SetTransparentColor (globalData.clrBtnFace);

	CSize sizeImageDest (-1, -1);
	if (globalData.GetRibbonImageScale () != 1.)
	{
		double dblImageScale = globalData.GetRibbonImageScale ();
		sizeImageDest = m_Images.GetImageSize ();
		sizeImageDest = CSize ((int)(.5 + sizeImageDest.cx * dblImageScale), (int)(.5 + sizeImageDest.cy * dblImageScale));
	}

	CBCGPDrawState ds;
	m_Images.PrepareDrawImage (ds, sizeImageDest);
#endif

	pDC->SetBkMode (TRANSPARENT);
	pDC->SetTextColor (globalData.clrWindowText);

	CFont* pOldFont = pDC->SelectObject (&globalData.fontRegular);
	ASSERT(pOldFont != NULL);

	for (int i = 0; i < m_arButtons.GetSize (); i++)
	{
		CBCGPToolBoxButton* pButton = m_arButtons [i];
		ASSERT_VALID (pButton);

		pButton->OnDraw (pDC);
	}

	m_Images.EndDrawImage (ds);
	pDC->SelectObject (pOldFont);
}
//****************************************************************************************
void CBCGPToolBoxPage::OnSize(UINT nType, int cx, int cy) 
{
	CBCGPControlBar::OnSize(nType, cx, cy);
	AdjustLayout ();
}
//****************************************************************************************
void CBCGPToolBoxPage::AdjustLayout ()
{
	CBCGPControlBar::AdjustLayout ();

	if (m_nScrollValue == 0)
	{
		m_nVertScrollOffset = 0;
	}

	m_nVertScrollSize = 0;

	ReposButtons ();

	if (m_pToolBox != NULL)
	{
		CBCGPOutlookWnd* pOlWnd = DYNAMIC_DOWNCAST (CBCGPOutlookWnd, m_pToolBox->GetTabWnd ());
		ASSERT_VALID (pOlWnd);

		pOlWnd->EnableScrollButtons (TRUE, m_nVertScrollOffset > 0, m_nVertScrollSize > 0);
	}

	RedrawWindow (NULL, NULL, RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
}
//****************************************************************************************
int CBCGPToolBoxPage::HitTest(CPoint point)
{
	CPoint ptScreen = point;
	ClientToScreen (&ptScreen);

	if (CWnd::WindowFromPoint (ptScreen)->GetSafeHwnd () != GetSafeHwnd ())
	{
		return -1;
	}

	for (int i = 0; i < m_arButtons.GetSize (); i++)
	{
		CBCGPToolBoxButton* pButton = m_arButtons [i];
		ASSERT_VALID (pButton);

		if (pButton->GetRect ().PtInRect (point))
		{
			return i;
		}
	}

	return -1;
}
//****************************************************************************************
void CBCGPToolBoxPage::ReposButtons()
{
	CRect rectClient;
	GetClientRect (rectClient);

	if (m_Mode == ToolBoxPageMode_Images)
	{
		rectClient.DeflateRect(TOOLBOX_IMAGE_MARGIN, 0);
	}

	m_nVertScrollSize = 0;

	const int nButtonWidth = m_Mode == ToolBoxPageMode_Images ? 
		m_sizeImage.cx + 3 * TOOLBOX_IMAGE_MARGIN :
		rectClient.Width ();

	const int nButtonHeight = m_Mode == ToolBoxPageMode_Images ? 
		m_sizeImage.cy + 3 * TOOLBOX_IMAGE_MARGIN :
		max (globalData.GetTextHeight (), m_sizeImage.cy) + TOOLBOX_IMAGE_MARGIN;

	int x = rectClient.left;
	int y = rectClient.top + TOOLBOX_IMAGE_MARGIN - 
		(m_nScrollValue > 0 ? m_nScrollValue : m_nVertScrollOffset * nButtonHeight);

	for (int i = 0; i < m_arButtons.GetSize (); i++)
	{
		CBCGPToolBoxButton* pButton = m_arButtons [i];
		ASSERT_VALID (pButton);

		CRect rectButton = CRect(CPoint(x, y), CSize(nButtonWidth, nButtonHeight));

		if ((m_nScrollValue > 0 ? rectButton.bottom : rectButton.top) >= rectClient.top && rectButton.top < rectClient.bottom && pButton->m_bInFilter)
		{
			pButton->SetRect (rectButton);

			if (m_pToolTip->GetSafeHwnd () != NULL)
			{
				m_pToolTip->SetToolRect (this, pButton->GetID (), rectButton);
			}
		}
		else
		{
			pButton->Hide ();

			if (m_pToolTip->GetSafeHwnd () != NULL)
			{
				m_pToolTip->SetToolRect (this, pButton->GetID (), CRect (0, 0, 0, 0));
			}
		}

		if (pButton->m_bInFilter)
		{
			x += nButtonWidth;

			if (x + nButtonWidth > rectClient.right)
			{
				x = rectClient.left;
				y += nButtonHeight;

				if (y >= rectClient.bottom)
				{
					m_nVertScrollSize++;
				}
			}
		}
	}

	RedrawWindow ();
}
//***************************************************************************************
BOOL CBCGPToolBoxPage::SetFilter(const CString& strFilter)
{
	ASSERT_VALID (this);

	BOOL bHasVisibleItems = FALSE;

	for (int i = 0; i < m_arButtons.GetSize (); i++)
	{
		CBCGPToolBoxButton* pButton = m_arButtons [i];
		ASSERT_VALID (pButton);

		if (pButton->SetFilter(strFilter))
		{
			bHasVisibleItems = TRUE;
		}
	}

	return bHasVisibleItems;
}
//***************************************************************************************
void CBCGPToolBoxPage::HighlightButton (int nButton)
{
	ASSERT_VALID (this);

	for (int i = 0; i < m_arButtons.GetSize (); i++)
	{
		CBCGPToolBoxButton* pButton = m_arButtons [i];
		ASSERT_VALID (pButton);

		if (pButton->IsHighlighted ())
		{
			if (i == nButton)
			{
				// Already highlighted
				return;
			}

			pButton->Highlight (FALSE);
			InvalidateRect (pButton->GetRect ());
		}
		
		if (i == nButton)
		{
			pButton->Highlight ();
			InvalidateRect (pButton->GetRect ());
		}
	}

	UpdateWindow ();
}
//***************************************************************************************
void CBCGPToolBoxPage::RedrawButton (int nButton)
{
	ASSERT_VALID (this);

	if (nButton < 0 || nButton >= m_arButtons.GetSize ())
	{
		return;
	}

	CBCGPToolBoxButton* pButton = m_arButtons [nButton];
	ASSERT_VALID (pButton);

	RedrawWindow (pButton->GetRect ());
}
//***************************************************************************************
CBCGPToolBoxButton* CBCGPToolBoxPage::GetButton (int nButton)
{
	ASSERT_VALID (this);

	if (nButton < 0 || nButton >= m_arButtons.GetSize ())
	{
		return NULL;
	}

	CBCGPToolBoxButton* pButton = m_arButtons [nButton];
	ASSERT_VALID (pButton);

	return pButton;
}
//***************************************************************************************
CBCGPToolBoxButton* CBCGPToolBoxPage::GetButtonByID (int nID)
{
	ASSERT_VALID (this);

	for (int i = 0; i < m_arButtons.GetSize (); i++)
	{
		CBCGPToolBoxButton* pButton = m_arButtons [i];
		ASSERT_VALID (pButton);

		if (pButton->GetID () == nID)
		{
			return pButton;
		}
	}

	return NULL;
}
//***************************************************************************************
BOOL CBCGPToolBoxPage::SetButtonUserData(int nItem, DWORD_PTR dwUserData)
{
	ASSERT_VALID (this);

	CBCGPToolBoxButton* pButton = GetButton(nItem);
	if (pButton == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID(pButton);

	pButton->m_dwUserData = dwUserData;
	return TRUE;
}
//***************************************************************************************
DWORD_PTR CBCGPToolBoxPage::GetButtonUserData(int nItem) const
{
	ASSERT_VALID (this);
	
	CBCGPToolBoxButton* pButton = ((CBCGPToolBoxPage*)this)->GetButton(nItem);
	if (pButton == NULL)
	{
		return 0;
	}
	
	ASSERT_VALID(pButton);
	return pButton->m_dwUserData;
}
//***************************************************************************************
int CBCGPToolBoxPage::GetMaxHeight()
{
	int nButtons = (int)m_arButtons.GetSize ();

#ifndef BCGP_EXCLUDE_TASK_PANE
	if (m_pToolBoxEx->GetSafeHwnd() != NULL
#if (!defined _BCGSUITE_)
		&& !m_pToolBoxEx->GetFilterText().IsEmpty()
#endif
		)
	{
		nButtons = 0;

		for (int i = 0; i < m_arButtons.GetSize(); i++)
		{
			CBCGPToolBoxButton* pButton = m_arButtons [i];
			ASSERT_VALID (pButton);
			
			if (pButton->m_bInFilter)
			{
				nButtons++;
			}
		}
	}
#endif

	if (nButtons == 0)
	{
		return 0;
	}

	int nRows = nButtons;
	int nButtonHeight = 0;
	
	if (m_Mode != ToolBoxPageMode_Images)
	{
		nButtonHeight = max (globalData.GetTextHeight (), m_sizeImage.cy) + TOOLBOX_IMAGE_MARGIN;
	}
	else
	{
		nButtonHeight = m_sizeImage.cy + 3 * TOOLBOX_IMAGE_MARGIN;
		const int nButtonWidth = m_sizeImage.cx + 3 * TOOLBOX_IMAGE_MARGIN;
		
		CRect rectClient;
		GetClientRect (rectClient);
		
		const int nButtonsInRow = rectClient.Width () / nButtonWidth;
		
		if (nButtonsInRow > 0)
		{
			nRows = nButtons / nButtonsInRow;
		}
		else
		{
			nRows = 0;
		}
		
		if (nButtonsInRow * nRows < nButtons)
		{
			nRows++;
		}
	}
	
	return nRows * nButtonHeight + 2 * TOOLBOX_IMAGE_MARGIN;
}
//***************************************************************************************
BOOL CBCGPToolBoxPage::InitPage (UINT uiBmpResID, int nImageWidth, 
								 const CStringList& lstLabels,
								 CRuntimeClass* pButtonClass)
{
	ASSERT_VALID (this);

	// Enable tooltips:
	CBCGPTooltipManager::CreateToolTip (m_pToolTip, this, BCGP_TOOLTIP_TYPE_TOOLBOX);

	m_pButtonClass = pButtonClass;
	m_uiBmpResID = uiBmpResID;

	if (uiBmpResID == 0)
	{
		return TRUE;	// Empty page
	}

	if (!m_Images.Load (uiBmpResID))
	{
		ASSERT (FALSE);
		return FALSE;
	}

	int cy = nImageWidth;
	
	HBITMAP hbmp = m_Images.GetImageWell();
	if (hbmp != NULL)
	{
		BITMAP bmp;
		GetObject(hbmp, sizeof (BITMAP), &bmp);
		
		cy = bmp.bmHeight;
	}

	m_sizeImage.cx = nImageWidth;
	m_sizeImage.cy = cy;

	m_Images.SetImageSize (m_sizeImage, TRUE);

	if (globalData.GetRibbonImageScale () != 1.)
	{
#ifdef _BCGSUITE_
		double dblImageScale = globalData.GetRibbonImageScale ();
		m_sizeImage = CSize ((int)(.5 + m_sizeImage.cx * dblImageScale), (int)(.5 + m_sizeImage.cy * dblImageScale));
#else
		m_Images.SetTransparentColor (globalData.clrBtnFace);
		m_Images.SmoothResize(globalData.GetRibbonImageScale());
		m_sizeImage = m_Images.GetImageSize();
#endif
	}

	for (int i = 0; i < m_Images.GetCount (); i++)
	{
		CBCGPToolBoxButton* pButton = DYNAMIC_DOWNCAST (CBCGPToolBoxButton, pButtonClass->CreateObject ());
		if (pButton == NULL)
		{
			ASSERT (FALSE);
			return FALSE;
		}

		ASSERT_VALID (pButton);

		pButton->SetParentPage (this);
		pButton->SetImageList (&m_Images);
		pButton->SetImageIndex (i);
		pButton->SetID (i + 1);

		POSITION posLabel = lstLabels.FindIndex (i);
		if (posLabel != NULL)
		{
			pButton->SetLabel (lstLabels.GetAt (posLabel));
		}

		if (i == 0 && m_bCheckFirstButton)
		{
			m_nCheckedButton = 0;
			pButton->SetCheck (TRUE);
			RedrawButton (m_nCheckedButton);
		}

		m_arButtons.Add (pButton);

		if (m_pToolTip->GetSafeHwnd () != NULL)
		{
			CRect rectDummy;
			rectDummy.SetRectEmpty ();

			m_pToolTip->AddTool (	this, LPSTR_TEXTCALLBACK, &rectDummy, 
									pButton->GetID ());
		}
	}

	return TRUE;
}
//*****************************************************************************************
BOOL CBCGPToolBoxPage::SetSelected (int nItem)
{
	ASSERT_VALID (this);

	if (nItem >= m_arButtons.GetSize ())
	{
		ASSERT (FALSE);
		return FALSE;
	}

	if (m_nCheckedButton >= 0)
	{
		if (nItem == m_nCheckedButton)
		{
			return TRUE;
		}

		m_arButtons [m_nCheckedButton]->SetCheck (FALSE);
		RedrawButton (m_nCheckedButton);
	}

	if (nItem >= 0)
	{
		m_arButtons [nItem]->SetCheck (TRUE);
		RedrawButton (nItem);
	}

	m_nCheckedButton = nItem;
	return TRUE;
}
//*****************************************************************************************
void CBCGPToolBoxPage::SetMode (ToolBoxPageMode mode, BOOL bRecalLayout)
{
	ASSERT_VALID (this);

	m_Mode = mode;

	m_nScrollValue = 0;
	m_nVertScrollOffset = 0;

	if (GetSafeHwnd () != NULL)
	{
		ReposButtons ();

		if (bRecalLayout)
		{
#ifndef BCGP_EXCLUDE_TASK_PANE
			if (m_pToolBoxEx->GetSafeHwnd() != NULL)
			{
				m_pToolBoxEx->AdjustLayout();
			}
#endif
		
			if (m_pToolBox->GetSafeHwnd() != NULL)
			{
				m_pToolBox->RecalcLayout();
			}
		}
	}
}
//*****************************************************************************************
int CBCGPToolBoxPage::AddButton(LPCTSTR lpszText, HICON hIcon, BOOL bAlphaBlend)
{
	ASSERT_VALID (this);
	ASSERT (lpszText != NULL);
	ASSERT (hIcon != NULL);

	// Add icon to the image list:
	int iIconIdx = -1;

	if (m_Images.GetCount() == 0)
	{
		m_Images.SetImageSize(globalUtils.GetIconSize(hIcon));
	}
	
#ifndef _BCGSUITE_
	if (m_Images.IsScaled())
	{
		CBCGPToolBarImages imageIcon;
		imageIcon.SetImageSize(globalUtils.GetIconSize(hIcon));

		imageIcon.AddIcon(hIcon, bAlphaBlend || m_Images.GetBitsPerPixel() == 32);

		iIconIdx = m_Images.AddImage(imageIcon, 0);
	}
	else
	{
		iIconIdx = m_Images.AddIcon(hIcon, bAlphaBlend || m_Images.GetBitsPerPixel() == 32);
	}
#else
	iIconIdx = m_Images.AddIcon(hIcon, bAlphaBlend);
#endif

	if (m_Images.GetCount() == 1)
	{
#ifndef _BCGSUITE_
		if (globalData.GetRibbonImageScale () != 1.0)
		{
			m_Images.SetTransparentColor (globalData.clrBtnFace);
			m_Images.SmoothResize(globalData.GetRibbonImageScale());
		}
#endif
	}

	if (m_sizeImage == CSize(0, 0))
	{
		m_sizeImage = m_Images.GetImageSize();
	}

	CBCGPToolBoxButton* pButton = DYNAMIC_DOWNCAST (CBCGPToolBoxButton, m_pButtonClass->CreateObject());
	if (pButton == NULL)
	{
		ASSERT (FALSE);
		return -1;
	}

	ASSERT_VALID (pButton);

	pButton->SetParentPage (this);
	pButton->SetImageList (&m_Images);
	pButton->SetImageIndex (iIconIdx);
	pButton->SetLabel (lpszText);
	pButton->SetID ((int) m_arButtons.GetSize () + 1);

	m_arButtons.Add (pButton);

	if (m_pToolTip->GetSafeHwnd () != NULL)
	{
		CRect rectDummy;
		rectDummy.SetRectEmpty ();

		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectDummy, pButton->GetID ());
	}

	if (m_pToolBox != NULL)
	{
		m_pToolBox->AdjustLayout ();
	}
	else if (m_pToolBoxEx != NULL)
	{
#if (!defined BCGP_EXCLUDE_TASK_PANE) && (!defined _BCGSUITE_)
		pButton->SetFilter(m_pToolBoxEx->GetFilterText());
		m_pToolBoxEx->AdjustLayout ();
#endif
	}

	AdjustLayout ();
	return iIconIdx;
}
//*****************************************************************************************
BOOL CBCGPToolBoxPage::DeleteButton (int nItem)
{
	ASSERT_VALID (this);

	if (nItem < 0 || nItem >= (int) m_arButtons.GetSize ())
	{
		return FALSE;
	}
	
	ASSERT_VALID (m_arButtons [nItem]);

	m_Images.DeleteImage (nItem);
	
	if (m_pToolTip->GetSafeHwnd () != NULL)
	{
		m_pToolTip->DelTool (this, m_arButtons [nItem]->GetID ());
	}
	
	delete m_arButtons [nItem];
	m_arButtons.RemoveAt (nItem);
	
	if (m_pToolBox != NULL)
	{
		m_pToolBox->AdjustLayout ();
	}
	else if (m_pToolBoxEx != NULL)
	{
#ifndef BCGP_EXCLUDE_TASK_PANE
		m_pToolBoxEx->AdjustLayout ();
#endif
	}
	
	m_nCheckedButton = m_nHighlightedButton = -1;

	AdjustLayout ();
	return TRUE;
}
//*****************************************************************************************
void CBCGPToolBoxPage::DeleteAllButons ()
{
	ASSERT_VALID (this);

	const int nSize = (int) m_arButtons.GetSize ();

	for (int i = 0; i < nSize; i++)
	{
		ASSERT_VALID (m_arButtons [i]);

		if (m_pToolTip->GetSafeHwnd () != NULL)
		{
			m_pToolTip->DelTool (this, m_arButtons [i]->GetID ());
		}

		delete m_arButtons [i];
	}
	
	m_arButtons.RemoveAll ();
	m_Images.Clear ();

	if (m_pToolBox != NULL)
	{
		m_pToolBox->AdjustLayout ();
	}
	else if (m_pToolBoxEx != NULL)
	{
#ifndef BCGP_EXCLUDE_TASK_PANE
		m_pToolBoxEx->AdjustLayout ();
#endif
	}
	
	m_nCheckedButton = m_nHighlightedButton = -1;
	AdjustLayout ();
}
//*****************************************************************************************
void CBCGPToolBoxPage::OnNcDestroy() 
{
	CBCGPControlBar::OnNcDestroy();
	delete this;
}
//*****************************************************************************************
void CBCGPToolBoxPage::SetScrollValue (int nScrollValue)
{
	ASSERT (nScrollValue >= 0);

	m_nScrollValue = nScrollValue;

	ReposButtons ();
}
//*********************************************************************************
BOOL CBCGPToolBoxPage::OnClickButton(int nButtonIndex)
{
	if (nButtonIndex < 0 || nButtonIndex >= (int)m_arButtons.GetSize())
	{
		return FALSE;
	}
	
	if (m_nCheckedButton >= 0)
	{
		m_arButtons [m_nCheckedButton]->SetCheck(FALSE);
		RedrawButton (m_nCheckedButton);
	}
	
	m_nCheckedButton = nButtonIndex;
	if (m_nCheckedButton < 0)
	{
		return FALSE;
	}
	
	m_arButtons [m_nCheckedButton]->SetCheck (TRUE);
	RedrawButton (m_nCheckedButton);
	
	if (m_pToolBox != NULL)
	{
		ASSERT_VALID (m_pToolBox);
		int nPage = m_pToolBox->GetPageNumber (this);
		
		m_pToolBox->OnClickTool (nPage, m_nCheckedButton);
	}
	
	if (m_pToolBoxEx != NULL)
	{
#ifndef BCGP_EXCLUDE_TASK_PANE
		ASSERT_VALID (m_pToolBoxEx);
		int nPage = m_pToolBoxEx->GetPageNumber (this);
		
		m_pToolBoxEx->OnClickTool (nPage, m_nCheckedButton);
#endif
	}

	return TRUE;
}
//*********************************************************************************
void CBCGPToolBoxPage::OnLButtonDblClk(UINT /*nFlags*/, CPoint /*point*/) 
{
}
//*********************************************************************************
void CBCGPToolBoxPage::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CWnd::OnLButtonDown(nFlags, point);

	SetFocus ();

	ReleaseCapture ();
	m_nHighlightedButton = -1;

	OnClickButton(HitTest (point));
}
//*********************************************************************************
void CBCGPToolBoxPage::OnLButtonUp(UINT nFlags, CPoint point) 
{
	CBCGPControlBar::OnLButtonUp(nFlags, point);
}
//*********************************************************************************
void CBCGPToolBoxPage::OnMouseMove(UINT nFlags, CPoint point) 
{
	CBCGPControlBar::OnMouseMove(nFlags, point);

	int nPrevHighlightedButton = m_nHighlightedButton;
	m_nHighlightedButton = HitTest (point);

	if (m_nHighlightedButton != nPrevHighlightedButton)
	{
		HighlightButton (m_nHighlightedButton);

		if (m_nHighlightedButton != -1)
		{
			if (nPrevHighlightedButton == -1)
			{
				SetCapture ();
			}
		}
		else
		{
			ReleaseCapture ();
		}
	}
}
//*********************************************************************************
BOOL CBCGPToolBoxPage::PreTranslateMessage(MSG* pMsg) 
{
   	switch (pMsg->message)
	{
	case WM_KEYDOWN:
	case WM_SYSKEYDOWN:
	case WM_LBUTTONDOWN:
	case WM_RBUTTONDOWN:
	case WM_MBUTTONDOWN:
	case WM_LBUTTONUP:
	case WM_RBUTTONUP:
	case WM_MBUTTONUP:
	case WM_NCLBUTTONDOWN:
	case WM_NCRBUTTONDOWN:
	case WM_NCMBUTTONDOWN:
	case WM_NCLBUTTONUP:
	case WM_NCRBUTTONUP:
	case WM_NCMBUTTONUP:
	case WM_MOUSEMOVE:
		if (m_pToolTip->GetSafeHwnd () != NULL)
		{
			m_pToolTip->RelayEvent(pMsg);
		}
		break;
	}

	return CBCGPControlBar::PreTranslateMessage(pMsg);
}
//**********************************************************************************
BOOL CBCGPToolBoxPage::OnNeedTipText(UINT /*id*/, NMHDR* pNMH, LRESULT* /*pResult*/)
{
	static CString strTipText;

	if (m_pToolTip->GetSafeHwnd () == NULL || 
		pNMH->hwndFrom != m_pToolTip->GetSafeHwnd ())
	{
		return FALSE;
	}

	if (CBCGPPopupMenu::GetActiveMenu () != NULL)
	{
		return FALSE;
	}

	LPNMTTDISPINFO	pTTDispInfo	= (LPNMTTDISPINFO) pNMH;
	ASSERT((pTTDispInfo->uFlags & TTF_IDISHWND) == 0);

	CBCGPToolBoxButton* pButton = GetButtonByID ((int) pNMH->idFrom);
	if (pButton == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID (pButton);

	strTipText = pButton->GetLabel ();
	pTTDispInfo->lpszText = const_cast<LPTSTR> ((LPCTSTR) strTipText);
	return TRUE;
}
//**************************************************************************
LRESULT CBCGPToolBoxPage::OnBCGUpdateToolTips (WPARAM wp, LPARAM)
{
	UINT nTypes = (UINT) wp;

	if (nTypes & BCGP_TOOLTIP_TYPE_TOOLBOX)
	{
		CBCGPTooltipManager::CreateToolTip (m_pToolTip, this,
			BCGP_TOOLTIP_TYPE_TOOLBOX);

		if (m_pToolTip->GetSafeHwnd () != NULL)
		{
			CRect rectDummy;
			rectDummy.SetRectEmpty ();

			for (int i = 0; i < m_arButtons.GetSize (); i++)
			{
				ASSERT_VALID (m_arButtons [i]);

				m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectDummy, 
					m_arButtons [i]->GetID ());
			}

			ReposButtons ();
		}
	}

	return 0;
}
//**************************************************************************
void CBCGPToolBoxPage::OnContextMenu(CWnd* pWnd, CPoint point) 
{
	ASSERT_VALID (this);

	CPoint ptClient = point;
	ScreenToClient (&ptClient);

	const int nHit = HitTest (ptClient);

	ReleaseCapture ();
	m_nHighlightedButton = -1;

	if (m_nCheckedButton >= 0)
	{
		m_arButtons [m_nCheckedButton]->SetCheck (FALSE);
		RedrawButton (m_nCheckedButton);
	}

	m_nCheckedButton = nHit;

	if (m_nCheckedButton >= 0)
	{
		m_arButtons [m_nCheckedButton]->SetCheck (TRUE);
		RedrawButton (m_nCheckedButton);
	}

	if (m_pToolBox != NULL)
	{
		ASSERT_VALID (m_pToolBox);

		if (m_pToolBox->OnShowToolboxMenu (point, this, nHit))
		{
			return;
		}
	}
	else if (m_pToolBoxEx != NULL)
	{
#ifndef BCGP_EXCLUDE_TASK_PANE
		ASSERT_VALID (m_pToolBoxEx);

		m_pToolBoxEx->OnActivatePage(m_pToolBoxEx->GetPageNumber(this));

		if (m_pToolBoxEx->OnShowToolboxMenu (point, this, nHit))
		{
			return;
		}
#endif
	}

	CBCGPControlBar::OnContextMenu (pWnd, point);
}
//****************************************************************************************
void CBCGPToolBoxPage::OnSysColorChange() 
{
	CBCGPControlBar::OnSysColorChange();
	UpdateImageList(m_uiBmpResID);
}
//*********************************************************************************
BOOL CBCGPToolBoxPage::UpdateImageList(UINT nResID)
{
	m_uiBmpResID = nResID;

#ifndef _BCGSUITE_
	CSize sizeImageOriginal = m_Images.GetOriginalImageSize();
	if (sizeImageOriginal != CSize(0, 0))
	{
		m_Images.SetImageSize(sizeImageOriginal);
	}

#endif
	if (!m_Images.Load (m_uiBmpResID))
	{
		return FALSE;
	}

#ifndef _BCGSUITE_
	if (globalData.GetRibbonImageScale () != 1.0)
	{
		m_Images.SetTransparentColor (globalData.clrBtnFace);
		m_Images.SmoothResize(globalData.GetRibbonImageScale());
		m_sizeImage = m_Images.GetImageSize();
	}
#endif

	RedrawWindow ();
	return TRUE;
}
//*********************************************************************************
BOOL CBCGPToolBoxPage::OnGestureEventPan(const CPoint& ptFrom, const CPoint& ptTo, CSize& sizeOverPan)
{
	if (m_pToolBoxEx != NULL)
	{
		return FALSE;
	}

	if (ptTo.y == ptFrom.y || m_pToolBox == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID(m_pToolBox);

	const int nDelta = ptTo.y - ptFrom.y;
	const int nButtonHeight = m_Mode == ToolBoxPageMode_Images ? 
		m_sizeImage.cy + 3 * TOOLBOX_IMAGE_MARGIN :
		max (globalData.GetTextHeight (), m_sizeImage.cy) + TOOLBOX_IMAGE_MARGIN;
	
	if (nButtonHeight == 0 || abs(nDelta) < nButtonHeight)
	{
		return FALSE;
	}

	const int nSteps = (int)(0.5 + abs(nDelta) / nButtonHeight);
	if (nSteps == 0)
	{
		return FALSE;
	}

	for (int i = 0; i < nSteps; i++)
	{
		if ((nDelta < 0 && m_nVertScrollSize == 0) || (nDelta > 0 && m_nVertScrollOffset == 0))
		{
			sizeOverPan.cy = nDelta;
			return TRUE;
		}
		
		m_pToolBox->OnScroll(nDelta < 0);
	}

	return TRUE;
}
//*********************************************************************************
BOOL CBCGPToolBoxPage::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	if (m_pToolBoxEx != NULL)
	{
		return CBCGPControlBar::OnMouseWheel(nFlags, zDelta, pt);
	}

	const int nButtonHeight = m_Mode == ToolBoxPageMode_Images ? 
		m_sizeImage.cy + 3 * TOOLBOX_IMAGE_MARGIN :
		max (globalData.GetTextHeight (), m_sizeImage.cy) + TOOLBOX_IMAGE_MARGIN;

	const int nSteps = (int)(0.5 + abs(zDelta) / nButtonHeight);
	if (nSteps == 0)
	{
		return FALSE;
	}
	
	for (int i = 0; i < nSteps; i++)
	{
		if (zDelta > 0)
		{
			if (m_nVertScrollOffset <= 0)
			{
				return TRUE;
			}
		}
		else
		{
			if (m_nVertScrollSize <= 0)
			{
				return TRUE;
			}
		}

		m_pToolBox->OnScroll(zDelta < 0);
	}

	return TRUE;
}
//*********************************************************************************
HRESULT CBCGPToolBoxPage::get_accParent(IDispatch **ppdispParent)
{
	if (ppdispParent == NULL)
	{
		return E_INVALIDARG;
	}

	*ppdispParent = NULL;

#ifndef BCGP_EXCLUDE_TASK_PANE
	if (m_pToolBoxEx->GetSafeHwnd() != NULL)
	{
		return AccessibleObjectFromWindow(m_pToolBoxEx->GetSafeHwnd(), (DWORD)OBJID_CLIENT, IID_IAccessible, (void**)ppdispParent);
	}
#endif

	if (m_pToolBox->GetSafeHwnd() != NULL)
	{
		return AccessibleObjectFromWindow(m_pToolBox->GetSafeHwnd(), (DWORD)OBJID_CLIENT, IID_IAccessible, (void**)ppdispParent);
	}

	return S_FALSE;
}
//*********************************************************************************
HRESULT CBCGPToolBoxPage::get_accChild(VARIANT /*varChild*/, IDispatch **ppdispChild)
{
	if (!(*ppdispChild))
	{
		return E_INVALIDARG;
	}

	return S_FALSE;
}
//*****************************************************************************
CBCGPToolBoxButton* CBCGPToolBoxPage::GetAccChild(int nIndex)
{
	ASSERT_VALID(this);

	nIndex--;

	if (nIndex < 0 || nIndex >= (int)m_arButtons.GetSize())
	{
		return NULL;
	}

	return m_arButtons[nIndex];
}
// *****************************************************************************
HRESULT CBCGPToolBoxPage::get_accChildCount(long *pcountChildren)
{
	if (!pcountChildren)
	{
		return E_INVALIDARG;
	}

	*pcountChildren = (long)m_arButtons.GetSize();
	return S_OK;
}
//****************************************************************************
HRESULT CBCGPToolBoxPage::get_accName(VARIANT varChild, BSTR *pszName)
{
	if (varChild.vt == VT_I4)
	{
		if (varChild.lVal == CHILDID_SELF)
		{
			CString strText = _T("Toolbox Page");

			if (GetSafeHwnd() != NULL)
			{
				GetWindowText(strText);
			}

			*pszName = strText.AllocSysString();
			return S_OK;
		}
		else
		{
			CBCGPToolBoxButton* pButton = GetAccChild(varChild.lVal);
			if (pButton != NULL)
			{
				CString strName = pButton->m_strLabel;
				*pszName = strName.AllocSysString();

				return S_OK;
			}
		}
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPToolBoxPage::get_accDescription(VARIANT varChild, BSTR *pszDescription)
{
	if (((varChild.vt != VT_I4) && (varChild.lVal != CHILDID_SELF)) || (NULL == pszDescription))
	{
		return E_INVALIDARG;
	}

	if (varChild.vt == VT_I4)
	{
		if (varChild.lVal == CHILDID_SELF)
		{
			*pszDescription = SysAllocString(L"Toolbox Page");
			return S_OK;
		}
		else
		{
			CBCGPToolBoxButton* pButton = GetAccChild(varChild.lVal);
			if (pButton != NULL)
			{
				*pszDescription = SysAllocString(L"Toolbox Item");
				return S_OK;
			}
		}
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPToolBoxPage::get_accRole(VARIANT varChild, VARIANT *pvarRole)
{

	if (!pvarRole || ((varChild.vt != VT_I4) && (varChild.lVal != CHILDID_SELF)))
	{
		return E_INVALIDARG;
	}

	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		pvarRole->vt = VT_I4;
		pvarRole->lVal = ROLE_SYSTEM_LIST;

		return S_OK;
	}

	CBCGPToolBoxButton* pButton = GetAccChild(varChild.lVal);
	if (pButton != NULL)
	{
		pvarRole->vt = VT_I4;
		pvarRole->lVal = ROLE_SYSTEM_LISTITEM;

		return S_OK;
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPToolBoxPage::get_accState(VARIANT varChild, VARIANT *pvarState)
{
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		pvarState->vt = VT_I4;
		pvarState->lVal = STATE_SYSTEM_NORMAL;
		return S_OK;
	}

	pvarState->vt = VT_I4;
	pvarState->lVal = STATE_SYSTEM_FOCUSABLE;
	pvarState->lVal |= STATE_SYSTEM_SELECTABLE;
	
	if (varChild.vt == VT_I4)
	{
		CBCGPToolBoxButton* pButton = GetAccChild(varChild.lVal);
		if (pButton != NULL)
		{
			ASSERT_VALID(pButton);

			if (pButton->m_bIsChecked)
			{
				pvarState->lVal |= STATE_SYSTEM_FOCUSED;
				pvarState->lVal |= STATE_SYSTEM_SELECTED;
			}

			if (!pButton->m_bInFilter)
			{
				pvarState->lVal |= STATE_SYSTEM_INVISIBLE;
			}
		}
	}

	return S_OK;
}
//****************************************************************************
HRESULT CBCGPToolBoxPage::get_accDefaultAction(VARIANT varChild, BSTR* pszDefaultAction)
{
	if (varChild.vt == VT_I4 && varChild.lVal > 0)
	{
		CBCGPToolBoxButton* pButton = GetAccChild(varChild.lVal);
		if (pButton != NULL)
		{
			*pszDefaultAction = SysAllocString(L"Double-click");
			return S_OK;
		}
	}
	
	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPToolBoxPage::accSelect(long flagsSelect, VARIANT varChild)
{
	if (varChild.vt != VT_I4)
	{
		return E_INVALIDARG;
	}

	if (flagsSelect == SELFLAG_NONE || varChild.lVal == CHILDID_SELF)
	{
		return S_FALSE;
	}

	if (flagsSelect == SELFLAG_TAKEFOCUS || flagsSelect == SELFLAG_TAKESELECTION)
	{
		CBCGPToolBoxButton* pButton = GetAccChild(varChild.lVal);
		if (pButton != NULL)
		{
			SetSelected(varChild.lVal - 1);
			return S_OK;
		}
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPToolBoxPage::accLocation(long *pxLeft, long *pyTop, long *pcxWidth, long *pcyHeight, VARIANT varChild)
{
	if (!pxLeft || !pyTop || !pcxWidth || !pcyHeight)
	{
		return E_INVALIDARG;
	}

	if (GetSafeHwnd() == NULL)
	{
		return S_FALSE;
	}

	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		CRect rc;
		GetWindowRect(rc);

		*pxLeft = rc.left;
		*pyTop = rc.top;
		*pcxWidth = rc.Width();
		*pcyHeight = rc.Height();

		return S_OK;
	}
	else if (varChild.vt == VT_I4)
	{
		CBCGPToolBoxButton* pButton = GetAccChild(varChild.lVal);
		if (pButton != NULL)
		{
			CRect rcProp = pButton->GetRect();
			ClientToScreen(&rcProp);

			*pxLeft = rcProp.left;
			*pyTop = rcProp.top;
			*pcxWidth = rcProp.Width();
			*pcyHeight = rcProp.Height();
		}
	}

	return S_OK;
}
//****************************************************************************
HRESULT CBCGPToolBoxPage::accHitTest(long  xLeft, long yTop, VARIANT *pvarChild)
{
	if (!pvarChild)
	{
		return E_INVALIDARG;
	}

	if (GetSafeHwnd() == NULL)
	{
        pvarChild->vt = VT_EMPTY;
		return S_FALSE;
	}

	CPoint pt(xLeft, yTop);
	ScreenToClient(&pt);

	pvarChild->vt = VT_I4;

	int nButtonIndex = HitTest(pt);
	if (nButtonIndex >= 0)
	{
		pvarChild->lVal = nButtonIndex + 1;
	}
	else
	{
		pvarChild->lVal = CHILDID_SELF;
	}

	return S_OK;
}
//******************************************************************************
HRESULT CBCGPToolBoxPage::accDoDefaultAction(VARIANT varChild)
{
    if (varChild.vt != VT_I4)
    {
        return E_INVALIDARG;
    }

	return OnClickButton((int)varChild.lVal - 1) ? S_OK : S_FALSE;
}
//******************************************************************************
int CBCGPToolBoxPage::GetOptimalWidth() const
{
	ASSERT_VALID(this);

	if (m_Mode == ToolBoxPageMode_Images)
	{
		return 0;
	}

	int nMaxWidth = 0;

	CClientDC dc((CWnd*)this);

	CFont* pOldFont = dc.SelectObject(&globalData.fontRegular);
	ASSERT(pOldFont != NULL);
	
	for (int i = 0; i < m_arButtons.GetSize(); i++)
	{
		CBCGPToolBoxButton* pButton = m_arButtons[i];
		ASSERT_VALID(pButton);

		nMaxWidth = max(nMaxWidth, pButton->GetMaxWidth(&dc));
	}

	dc.SelectObject(pOldFont);
	return nMaxWidth;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPToolBox

IMPLEMENT_DYNCREATE(CBCGPToolBox, CBCGPDockingControlBar)

CBCGPToolBox::CBCGPToolBox()
{
}

CBCGPToolBox::~CBCGPToolBox()
{
}

BEGIN_MESSAGE_MAP(CBCGPToolBox, CBCGPDockingControlBar)
	//{{AFX_MSG_MAP(CBCGPToolBox)
	ON_WM_CREATE()
	ON_WM_SIZE()
	//}}AFX_MSG_MAP
	ON_REGISTERED_MESSAGE(BCGM_ONGETGESTURECONFIG, OnGetGestureConfig)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPToolBox message handlers

int CBCGPToolBox::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CBCGPDockingControlBar::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	CRect rectDummy (0, 0, 0, 0);

	// Create ToolBox pane:
	DWORD dwStyle = WS_CHILD | WS_VISIBLE | CBRS_ALIGN_LEFT;
	DWORD dwBCGStyle = 0;

	if (!m_wndOutlook.Create (_T(""), this, rectDummy, 
		0, dwStyle, dwBCGStyle))
	{
		TRACE0("Failed to create ToolBox window\n");
		return -1;      // fail to create
	}

    m_wndOutlook.SetBarStyle (CBRS_ALIGN_LEFT | CBRS_TOOLTIPS | CBRS_FLYBY |
								CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC);
	m_wndOutlook.EnableSetCaptionTextToTabName (FALSE);

	return 0;
}
//****************************************************************************************
void CBCGPToolBox::OnSize(UINT nType, int cx, int cy) 
{
	CBCGPDockingControlBar::OnSize(nType, cx, cy);
	AdjustLayout ();
	RedrawWindow (NULL, NULL, RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
}
//****************************************************************************************
LRESULT CBCGPToolBox::OnGetGestureConfig(WPARAM, LPARAM)
{
	return 1L;	
}
//****************************************************************************************
BOOL CBCGPToolBox::AddToolsPage (LPCTSTR lpszPageName, UINT uiBmpResID, int nImageWidth,
								 LPCTSTR lpszLabels,
								 CRuntimeClass* pPageClass,
								 CRuntimeClass* pButtonClass)
{
	CStringList lstLabels;

	if (lpszLabels != NULL)
	{
		CString strLabels = lpszLabels;

		for (int i = 0; i < strLabels.GetLength ();)
		{
			CString strLabel;
				
			int iNextWord = strLabels.Find (_T('\n'), i);
			if (iNextWord == -1)
			{
				lstLabels.AddTail (strLabels.Mid (i));
				break;
			}

			lstLabels.AddTail (strLabels.Mid (i, iNextWord - i));
			i = iNextWord + 1;
		}
	}

	return AddToolsPage (lpszPageName, uiBmpResID, nImageWidth,
								 lstLabels, pPageClass, pButtonClass);
}
//****************************************************************************************
BOOL CBCGPToolBox::AddToolsPage (LPCTSTR lpszPageName, UINT uiBmpResID, int nImageWidth,
								 const CStringList& lstLabels, 
								 CRuntimeClass* pPageClass,
								 CRuntimeClass* pButtonClass)
{
	ASSERT_VALID (this);
	ASSERT (GetSafeHwnd () != NULL);
	ASSERT (pPageClass != NULL);
	ASSERT (pPageClass->IsDerivedFrom (RUNTIME_CLASS (CBCGPToolBoxPage)));
	ASSERT (pButtonClass != NULL);
	ASSERT (pButtonClass->IsDerivedFrom (RUNTIME_CLASS (CBCGPToolBoxButton)));
	ASSERT (lpszPageName != NULL);

	CBCGPToolBoxPage* pPage = DYNAMIC_DOWNCAST (
		CBCGPToolBoxPage, pPageClass->CreateObject ());
	if (pPage == NULL)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	CRect rectDummy (0, 0, 0, 0);
	pPage->Create (NULL, WS_VISIBLE | WS_CHILD, rectDummy, this, 0);

	pPage->SetWindowText (lpszPageName);
	if (!pPage->InitPage (uiBmpResID, nImageWidth, lstLabels, pButtonClass))
	{
		delete pPage;
		return FALSE;
	}

	m_wndOutlook.AddTab (pPage, TRUE, TRUE, FALSE);
	pPage->m_pToolBox = this;

	OnActivatePage (m_wndOutlook.GetTabsNum () - 1);

	CBCGPGestureConfig gestureConfig;
	gestureConfig.EnablePan(TRUE, BCGP_GC_PAN_WITH_SINGLE_FINGER_VERTICALLY | BCGP_GC_PAN_WITH_GUTTER | BCGP_GC_PAN_WITH_INERTIA);
	
	bcgpGestureManager.SetGestureConfig(pPage->GetSafeHwnd(), gestureConfig);
	return TRUE;
}
//*****************************************************************************************
int CBCGPToolBox::GetActivePage () const
{
	ASSERT_VALID (this);

	CBCGPBaseTabWnd* pTab = GetTabWnd ();
	if (pTab == NULL)
	{
		ASSERT (FALSE);
		return  -1;
	}

	ASSERT_VALID (pTab);

	return pTab->GetActiveTab ();
}
//*****************************************************************************************
int CBCGPToolBox::GetLastClickedTool (int nPage) const
{
	ASSERT_VALID (this);

	CBCGPBaseTabWnd* pTab = GetTabWnd ();
	if (pTab == NULL)
	{
		ASSERT (FALSE);
		return  -1;
	}

	ASSERT_VALID (pTab);

	if (nPage < 0 || nPage >= pTab->GetTabsNum ())
	{
		ASSERT (FALSE);
		return -1;
	}

	CBCGPToolBoxPage* pPage = DYNAMIC_DOWNCAST (CBCGPToolBoxPage, pTab->GetTabWnd (nPage));
	if (pPage == NULL)
	{
		ASSERT (FALSE);
		return -1;
	}

	ASSERT_VALID (pPage);
	return pPage->m_nCheckedButton;
}
//******************************************************************************************
int CBCGPToolBox::GetPageNumber (CBCGPToolBoxPage* pPage) const
{
	ASSERT_VALID (this);
	ASSERT_VALID (pPage);

	CBCGPBaseTabWnd* pTab = GetTabWnd ();
	if (pTab == NULL)
	{
		ASSERT (FALSE);
		return  -1;
	}

	ASSERT_VALID (pTab);

	for (int i = 0; i < pTab->GetTabsNum (); i++)
	{
		CBCGPToolBoxPage* pListPage = DYNAMIC_DOWNCAST (CBCGPToolBoxPage, pTab->GetTabWnd (i));
		if (pListPage->GetSafeHwnd () == pPage->GetSafeHwnd ())
		{
			return i;
		}
	}

	return -1;
}
//****************************************************************************************
CBCGPToolBoxPage* CBCGPToolBox::GetPage (int nPage) const
{
	ASSERT_VALID (this);

	CBCGPBaseTabWnd* pTab = GetTabWnd ();
	if (pTab == NULL)
	{
		ASSERT (FALSE);
		return  NULL;
	}

	ASSERT_VALID (pTab);

	return DYNAMIC_DOWNCAST (CBCGPToolBoxPage, pTab->GetTabWnd (nPage));
}
//****************************************************************************************
void CBCGPToolBox::OnClickTool (int /*nPage*/, int /*nIndex*/)
{
	ASSERT_VALID (this);

	CWnd* pOwner = GetOwner ();
	if (pOwner != NULL)
	{
		pOwner->PostMessage (	WM_COMMAND,
								MAKEWPARAM (GetDlgCtrlID (), BN_CLICKED),
								(LPARAM) m_hWnd);
	}
}
//*****************************************************************************************
BOOL CBCGPToolBox::OnShowToolboxMenu (CPoint /*point*/, 
									  CBCGPToolBoxPage* /*pPage*/, int /*nHit*/)
{
	return FALSE;
}
//*****************************************************************************************
CBCGPBaseTabWnd* CBCGPToolBox::GetTabWnd () const
{
	ASSERT_VALID (this);
#ifndef _BCGSUITE_
	return ((CBCGPToolBox*) this)->m_wndOutlook.GetUnderlinedWindow ();
#else
	return ((CBCGPToolBox*) this)->m_wndOutlook.GetUnderlyingWindow ();
#endif
}
//****************************************************************************
HRESULT CBCGPToolBox::get_accChildCount(long *pcountChildren)
{
	if (!pcountChildren)
	{
		return E_INVALIDARG;
	}

	*pcountChildren = (long)m_wndOutlook.GetTabsNum ();
	return S_OK;
}
//****************************************************************************
HRESULT CBCGPToolBox::get_accChild(VARIANT varChild, IDispatch** ppdispChild)
{
	if (ppdispChild == NULL)
	{
		return E_INVALIDARG;
	}

	if ((varChild.vt == VT_I4) && (varChild.lVal >= 1))
	{
		int nPage = varChild.lVal - 1;

		CBCGPToolBoxPage* pPage = GetPage(nPage);
		if (pPage->GetSafeHwnd() != NULL)
		{
			*ppdispChild = NULL;
			AccessibleObjectFromWindow(pPage->GetSafeHwnd(), (DWORD)OBJID_CLIENT, IID_IAccessible, (void**)ppdispChild);
		
			if (*ppdispChild != NULL)
			{
				return S_OK;
			}
		}
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPToolBox::accHitTest(long  xLeft, long yTop, VARIANT *pvarChild)
{
	if (!pvarChild)
	{
		return E_INVALIDARG;
	}

	CPoint pt(xLeft, yTop);
	CWnd* pWnd = CWnd::WindowFromPoint(pt);

	if (pWnd == this)
	{
		pvarChild->lVal = CHILDID_SELF;
		pvarChild->vt = VT_I4;
		return S_OK;
	}

	pvarChild->pdispVal = NULL;
	pvarChild->vt = VT_DISPATCH;
	
	AccessibleObjectFromWindow(pWnd->GetSafeHwnd(), (DWORD)OBJID_CLIENT, IID_IAccessible, (void**)&pvarChild->pdispVal);
	
	if (pvarChild->pdispVal != NULL)
	{
		return S_OK;
	}

	return S_FALSE;
}

BEGIN_MESSAGE_MAP(CBCGPToolBoxBar, CBCGPOutlookBar)
	ON_WM_CONTEXTMENU()
END_MESSAGE_MAP()

void CBCGPToolBoxBar::OnContextMenu(CWnd* pWnd, CPoint point)
{
	CBCGPToolBox* pParentBar = DYNAMIC_DOWNCAST (CBCGPToolBox, GetParent ());
	if (pParentBar == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	CPoint ptClient = point;
	ScreenToClient (&ptClient);

	CRect rectClient;
	GetClientRect (rectClient);

	if (!rectClient.PtInRect (ptClient) || 
		!pParentBar->OnShowToolboxMenu (point, NULL, -1))
	{
		CBCGPOutlookBar::OnContextMenu (pWnd, point);
	}
}
//***************************************************************************************
void CBCGPToolBoxBar::OnActivateTab (int iTabNum)
{
	ASSERT_VALID (this);

	CBCGPOutlookBar::OnActivateTab (iTabNum);

	CBCGPToolBox* pParentBar = DYNAMIC_DOWNCAST (CBCGPToolBox, GetParent ());
	if (pParentBar == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	pParentBar->OnActivatePage (iTabNum);
}
//****************************************************************************************
void CBCGPToolBoxBar::OnScroll (BOOL bDown)
{
	ASSERT_VALID (this);

	CBCGPOutlookBar::OnScroll (bDown);

	CBCGPToolBox* pParentBar = DYNAMIC_DOWNCAST (CBCGPToolBox, GetParent ());
	if (pParentBar == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	pParentBar->OnScroll (bDown);
}
//****************************************************************************************
void CBCGPToolBox::OnActivatePage (int /*nPage*/)
{
	ASSERT_VALID (this);

	CBCGPOutlookWnd* pOlWnd = DYNAMIC_DOWNCAST (CBCGPOutlookWnd, GetTabWnd ());
	ASSERT_VALID (pOlWnd);

	CBCGPToolBoxPage* pPage = GetPage (GetActivePage ());
	if (pPage == NULL)
	{
		ASSERT (FALSE);
		return;
	}

	pOlWnd->EnableScrollButtons (TRUE, pPage->m_nVertScrollOffset > 0, pPage->m_nVertScrollSize > 0);
}
//***************************************************************************************
void CBCGPToolBox::OnScroll (BOOL bDown)
{
	ASSERT_VALID (this);

	CBCGPOutlookWnd* pOlWnd = DYNAMIC_DOWNCAST (CBCGPOutlookWnd, GetTabWnd ());
	ASSERT_VALID (pOlWnd);

	CBCGPToolBoxPage* pPage = GetPage (GetActivePage ());
	if (pPage == NULL)
	{
		ASSERT (FALSE);
		return;
	}

	if (bDown)
	{
		pPage->m_nVertScrollOffset++;
	}
	else
	{
		pPage->m_nVertScrollOffset--;
	}

	pPage->ReposButtons ();
	pOlWnd->EnableScrollButtons (TRUE, pPage->m_nVertScrollOffset > 0, pPage->m_nVertScrollSize > 0);
}
//***************************************************************************************
void CBCGPToolBox::AdjustLayout ()
{
	CBCGPDockingControlBar::AdjustLayout ();

	CRect rectClient;
	GetClientRect (rectClient);

	m_wndOutlook.SetWindowPos (NULL,
            -1, -1,
            rectClient.Width (), rectClient.Height (),
            SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);

	int nActivePage = GetActivePage ();
	if (nActivePage >= 0)
	{
		OnActivatePage (nActivePage);
	}
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPToolBoxButton

IMPLEMENT_DYNCREATE(CBCGPToolBoxButton, CBCGPToolBoxButton)

CBCGPToolBoxButton::CBCGPToolBoxButton ()
{
	m_iImageIndex = -1;
	m_bIsHighlighted = FALSE;
	m_bIsChecked = FALSE;
	m_Rect.SetRectEmpty ();
	m_pImages = NULL;
	m_pPage = NULL;
	m_nID = -1;
	m_bInFilter = TRUE;
	m_dwUserData = 0;
}
//***************************************************************************************
void CBCGPToolBoxButton::OnFillBackground (CDC* pDC, const CRect& rectClient)
{
	CBCGPVisualManager::GetInstance ()->OnEraseToolBoxButton (pDC, rectClient, this);
}
//***************************************************************************************
void CBCGPToolBoxButton::OnDrawBorder (CDC* pDC, CRect& rectClient, UINT uiState)
{
	CBCGPVisualManager::GetInstance ()->OnDrawToolBoxButtonBorder (pDC, rectClient,
		this, uiState);
}
//***************************************************************************************
CSize CBCGPToolBoxButton::GetImageSize()
{
	CSize sizeImage(0, 0);

	if (m_pImages != NULL && m_iImageIndex >= 0)
	{
		ASSERT_VALID (m_pImages);
		
		sizeImage = m_pImages->GetImageSize ();
#ifdef _BCGSUITE_
		if (globalData.GetRibbonImageScale () != 1.)
		{
			double dblImageScale = globalData.GetRibbonImageScale ();
			sizeImage = CSize ((int)(.5 + sizeImage.cx * dblImageScale), (int)(.5 + sizeImage.cy * dblImageScale));
		}
#endif
	}
	
	return sizeImage;
}
//***************************************************************************************
BOOL CBCGPToolBoxButton::SetFilter(const CString& strFilter)
{
	ASSERT_VALID (this);

	if (strFilter.IsEmpty())
	{
		m_bInFilter = TRUE;
	}
	else
	{
		CString strName = m_strLabel;
		strName.MakeUpper();
		
		m_bInFilter = strName.Find(strFilter) >= 0;
	}

	return m_bInFilter;
}
//***************************************************************************************
void CBCGPToolBoxButton::OnDraw (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);
	ASSERT_VALID (m_pPage);

	if (m_Rect.IsRectEmpty ())
	{
		return;
	}

	OnFillBackground (pDC, m_Rect);
	OnDrawBorder (pDC, m_Rect, 0);

	CRect rectText = m_Rect;

	if (m_pPage->GetMode () != CBCGPToolBoxPage::ToolBoxPageMode_Images)
	{
		rectText.left += rectText.Height();
	}

	if (m_pImages != NULL && m_iImageIndex >= 0)
	{
		ASSERT_VALID (m_pImages);

		CSize sizeImage = GetImageSize();
		CRect rectImage = rectText;

		if (m_pPage->GetMode () != CBCGPToolBoxPage::ToolBoxPageMode_Images)
		{
			rectImage.right = rectImage.left + sizeImage.cx + TOOLBOX_IMAGE_MARGIN;
		}

		m_pImages->Draw (pDC,
				rectImage.left + (rectImage.Width () - sizeImage.cx) / 2,
				rectImage.top + (rectImage.Height () - sizeImage.cy) / 2,
				m_iImageIndex);

		rectText.left = rectImage.right;
	}

	if (m_pPage->GetMode () == CBCGPToolBoxPage::ToolBoxPageMode_Images)
	{
		return;
	}

	rectText.DeflateRect(TOOLBOX_IMAGE_MARGIN, 0);

	COLORREF clrTextOld = pDC->SetTextColor (CBCGPVisualManager::GetInstance ()->GetToolBoxButtonTextColor (this));
	pDC->DrawText (m_strLabel, rectText, DT_LEFT | DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
	pDC->SetTextColor(clrTextOld);
}
//***************************************************************************************
int CBCGPToolBoxButton::GetMaxWidth(CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);
	ASSERT_VALID (m_pPage);
	
	if (m_pPage->GetMode () == CBCGPToolBoxPage::ToolBoxPageMode_Images)
	{
		return 0;
	}

	const CSize sizeImage = GetImageSize();
	const int nButtonHeight = max (globalData.GetTextHeight (), sizeImage.cy) + TOOLBOX_IMAGE_MARGIN;
	
	int nWidth = nButtonHeight;

	if (sizeImage != CSize(0, 0))
	{
		nWidth += sizeImage.cx + TOOLBOX_IMAGE_MARGIN;
	}		

	nWidth += 2 * TOOLBOX_IMAGE_MARGIN + pDC->GetTextExtent(m_strLabel).cx;
	return nWidth;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPToolBoxEx

#ifndef BCGP_EXCLUDE_TASK_PANE

IMPLEMENT_DYNCREATE(CBCGPToolBoxEx, CBCGPTasksPane)

CBCGPToolBoxEx::CBCGPToolBoxEx()
{
	m_nVertMargin = 0;
	m_nHorzMargin = 0;
	m_nGroupVertOffset = 0;
	m_nGroupCaptionHorzOffset = 0;
	m_nGroupCaptionVertOffset = 0;

	m_bOffsetCustomControls = FALSE;
	m_bAnimationEnabled = TRUE;

#if (!defined _BCGSUITE_)
	m_bFilterBox = FALSE;
#endif
}

CBCGPToolBoxEx::~CBCGPToolBoxEx()
{
}

BEGIN_MESSAGE_MAP(CBCGPToolBoxEx, CBCGPTasksPane)
	//{{AFX_MSG_MAP(CBCGPToolBoxEx)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_CONTEXTMENU()
	//}}AFX_MSG_MAP
	ON_EN_CHANGE(ID_FILTER, OnFilter)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPToolBoxEx message handlers

int CBCGPToolBoxEx::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CBCGPTasksPane::OnCreate(lpCreateStruct) == -1)
		return -1;

	EnableScrollButtons (FALSE);
	return 0;
}
//****************************************************************************************
void CBCGPToolBoxEx::OnSize(UINT nType, int cx, int cy) 
{
	CBCGPTasksPane::OnSize(nType, cx, cy);
	AdjustLayout ();
}
//****************************************************************************************
BOOL CBCGPToolBoxEx::AddToolsPage (LPCTSTR lpszPageName, UINT uiBmpResID, int nImageWidth,
								 LPCTSTR lpszLabels,
								 CRuntimeClass* pPageClass,
								 CRuntimeClass* pButtonClass)
{

	CStringList lstLabels;

	if (lpszLabels != NULL)
	{
		CString strLabels = lpszLabels;

		for (int i = 0; i < strLabels.GetLength ();)
		{
			CString strLabel;
				
			int iNextWord = strLabels.Find (_T('\n'), i);
			if (iNextWord == -1)
			{
				lstLabels.AddTail (strLabels.Mid (i));
				break;
			}

			lstLabels.AddTail (strLabels.Mid (i, iNextWord - i));
			i = iNextWord + 1;
		}
	}

	return AddToolsPage (lpszPageName, uiBmpResID, nImageWidth,
								 lstLabels, pPageClass, pButtonClass);
}
//****************************************************************************************
BOOL CBCGPToolBoxEx::AddToolsPage (LPCTSTR lpszPageName, UINT uiBmpResID, int nImageWidth,
								 const CStringList& lstLabels, 
								 CRuntimeClass* pPageClass,
								 CRuntimeClass* pButtonClass)
{
	ASSERT_VALID (this);
	ASSERT (GetSafeHwnd () != NULL);
	ASSERT (pPageClass != NULL);
	ASSERT (pPageClass->IsDerivedFrom (RUNTIME_CLASS (CBCGPToolBoxPage)));
	ASSERT (pButtonClass != NULL);
	ASSERT (pButtonClass->IsDerivedFrom (RUNTIME_CLASS (CBCGPToolBoxButton)));
	ASSERT (lpszPageName != NULL);

	CBCGPToolBoxPage* pPage = DYNAMIC_DOWNCAST (
		CBCGPToolBoxPage, pPageClass->CreateObject ());
	if (pPage == NULL)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	pPage->m_bCheckFirstButton = FALSE;

	CRect rectDummy (0, 0, 0, 0);
	pPage->Create (NULL, WS_VISIBLE | WS_CHILD, rectDummy, this, 0);

	pPage->SetWindowText (lpszPageName);
	if (!pPage->InitPage (uiBmpResID, nImageWidth, lstLabels, pButtonClass))
	{
		delete pPage;
		return FALSE;
	}

	int nHeight = pPage->GetMaxHeight ();
	int nGroup = AddGroup (0, lpszPageName);
	AddWindow (nGroup, pPage->GetSafeHwnd (), nHeight);

	pPage->m_pToolBoxEx = this;

	return TRUE;
}
//*****************************************************************************************
BOOL CBCGPToolBoxEx::RemoveToolsPage (int nPage)
{
	ASSERT_VALID (this);

	CBCGPToolBoxPage* pPage = GetPage (nPage);
	if (pPage == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID (pPage);
	pPage->DestroyWindow ();

	RemoveGroup (nPage);
	return TRUE;
}
//*****************************************************************************************
void CBCGPToolBoxEx::RemoveAllToolsPages ()
{
	ASSERT_VALID (this);

	for (int i = 0; i < (int) m_lstTaskGroups.GetCount (); i++)
	{
		CBCGPToolBoxPage* pPage = GetPage (i);
		ASSERT_VALID (pPage);
	
		pPage->DestroyWindow ();
	}

	RemoveAllGroups ();
}
//****************************************************************************************
int CBCGPToolBoxEx::GetLastClickedTool (int& nClickedPage) const
{
	ASSERT_VALID (this);

	for (int nPage = 0; nPage < GetPageCount (); nPage++)
	{
		CBCGPToolBoxPage* pPage = GetPage (nPage);
		ASSERT_VALID (pPage);

		if (pPage->GetSelected () >= 0)
		{
			nClickedPage = nPage;
			return pPage->GetSelected ();
		}
	}

	nClickedPage = -1;
	return -1;
}
//******************************************************************************************
int CBCGPToolBoxEx::GetPageNumber (CBCGPToolBoxPage* pPage) const
{
	ASSERT_VALID (this);
	ASSERT_VALID (pPage);

	int nPage = 0;

	for (POSITION pos = m_lstTaskGroups.GetHeadPosition (); pos != NULL; nPage++)
	{
		CBCGPTasksGroup* pGroup = (CBCGPTasksGroup*) m_lstTaskGroups.GetNext (pos);
		ASSERT_VALID (pGroup);

		CBCGPTask* pTask = (CBCGPTask*) pGroup->m_lstTasks.GetHead ();
		if (pTask != NULL && pTask->m_hwndTask == pPage->GetSafeHwnd ())
		{
			return nPage;
		}
	}

	return -1;
}
//****************************************************************************************
int CBCGPToolBoxEx::GetPageCount(BOOL bInFilter/* = FALSE*/) const
{
	ASSERT_VALID (this);

#if (!defined _BCGSUITE_)
	if (!bInFilter || m_strFilter.IsEmpty())
#else
	UNREFERENCED_PARAMETER(bInFilter);
#endif
	{
		return (int)m_lstTaskGroups.GetCount();
	}

#if (!defined _BCGSUITE_)
	int nCount = 0;

	for (POSITION pos = m_lstTaskGroups.GetHeadPosition (); pos != NULL;)
	{
		CBCGPTasksGroup* pGroup = (CBCGPTasksGroup*) m_lstTaskGroups.GetNext (pos);
		ASSERT_VALID (pGroup);
		
		if (pGroup->m_bIsVisible)
		{
			nCount++;
		}
	}

	return nCount;
#endif
}
//****************************************************************************************
CBCGPToolBoxPage* CBCGPToolBoxEx::GetPage (int nPage, BOOL bInFilter/* = FALSE*/) const
{
	ASSERT_VALID (this);

	CBCGPTasksGroup* pGroup = NULL;

#if (!defined _BCGSUITE_)
	if (!bInFilter || m_strFilter.IsEmpty())
#else
	UNREFERENCED_PARAMETER(bInFilter);
#endif
	{
		POSITION posGroup = m_lstTaskGroups.FindIndex (nPage);
		if (posGroup == NULL)
		{
			return NULL;
		}

		pGroup = (CBCGPTasksGroup*)m_lstTaskGroups.GetAt(posGroup);
	}
#if (!defined _BCGSUITE_)
	else
	{
		int nCurIndex = 0;
		
		for (POSITION pos = m_lstTaskGroups.GetHeadPosition (); pos != NULL;)
		{
			CBCGPTasksGroup* pCurrGroup = (CBCGPTasksGroup*) m_lstTaskGroups.GetNext (pos);
			ASSERT_VALID (pCurrGroup);
			
			if (pCurrGroup->m_bIsVisible)
			{
				if (nCurIndex == nPage)
				{
					pGroup = pCurrGroup;
					break;
				}

				nCurIndex++;
			}
		}
	}
#endif

	if (pGroup == NULL)
	{
		return NULL;
	}

	ASSERT_VALID(pGroup);

	CBCGPTask* pTask = (CBCGPTask*) pGroup->m_lstTasks.GetHead ();
	if (pTask == NULL || pTask->m_hwndTask == NULL)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	return DYNAMIC_DOWNCAST (CBCGPToolBoxPage, CWnd::FromHandle (pTask->m_hwndTask));
}
//****************************************************************************************
void CBCGPToolBoxEx::OnClickTool (int nPage, int /*nIndex*/)
{
	ASSERT_VALID (this);

	CWnd* pOwner = GetOwner ();
	if (pOwner != NULL)
	{
		pOwner->PostMessage (	WM_COMMAND,
								MAKEWPARAM (GetDlgCtrlID (), BN_CLICKED),
								(LPARAM) m_hWnd);
	}

	OnActivatePage(nPage);
}
//*****************************************************************************************
void CBCGPToolBoxEx::OnActivatePage (int nPage)
{
	ASSERT_VALID (this);

	for (int i = 0; i < GetPageCount (); i++)
	{
		if (i != nPage)
		{
			CBCGPToolBoxPage* pPage = GetPage (i);
			ASSERT_VALID (pPage);

			pPage->SetSelected (-1);
			pPage->m_nHighlightedButton = -1;
			pPage->HighlightButton (-1);
		}
	}
}
//*****************************************************************************************
BOOL CBCGPToolBoxEx::OnShowToolboxMenu (CPoint /*point*/, 
									  CBCGPToolBoxPage* /*pPage*/, int /*nHit*/)
{
	return FALSE;
}
//*********************************************************************************
void CBCGPToolBoxEx::OnFillBackground (CDC* pDC, CRect rectFill)
{
	CBCGPVisualManager::GetInstance ()->OnFillBarBackground (pDC, this, rectFill, rectFill);

#if (!defined _BCGSUITE_)
	BOOL bIsEmpty = TRUE;

	for (POSITION pos = m_lstTaskGroups.GetHeadPosition (); pos != NULL;)
	{
		CBCGPTasksGroup* pGroup = (CBCGPTasksGroup*) m_lstTaskGroups.GetNext (pos);
		ASSERT_VALID (pGroup);

		if (pGroup->m_bIsVisible)
		{
			bIsEmpty = FALSE;
			break;
		}
	}

	if (bIsEmpty)
	{
		CRect rectText = rectFill;
		
		rectText.top += rectText.Height() / 3;
		rectText.DeflateRect(10, 0);
		
		pDC->SetBkMode (TRANSPARENT);
		pDC->SetTextColor (globalData.clrBarText);
		
		CFont* pOldFont = pDC->SelectObject (&globalData.fontRegular);
		ASSERT(pOldFont != NULL);

		pDC->DrawText(m_strOutOfFilter, rectText, DT_CENTER | DT_WORDBREAK);

		pDC->SelectObject(pOldFont);
	}
#endif
}
//*********************************************************************************
void CBCGPToolBoxEx::ScrollChild (HWND hwndChild, int nScrollValue)
{
	ASSERT_VALID (this);

	CWnd* pChildWnd = CWnd::FromHandle (hwndChild);
	CBCGPToolBoxPage* pPage = DYNAMIC_DOWNCAST (CBCGPToolBoxPage, pChildWnd);
	ASSERT_VALID (pPage);

	pPage->SetScrollValue (nScrollValue);
}
//*********************************************************************************
void CBCGPToolBoxEx::AdjustLayout ()
{
	ASSERT_VALID (this);

	for (POSITION pos = m_lstTaskGroups.GetHeadPosition (); pos != NULL;)
	{
		CBCGPTasksGroup* pGroup = (CBCGPTasksGroup*) m_lstTaskGroups.GetNext (pos);
		ASSERT_VALID (pGroup);

		CBCGPTask* pTask = (CBCGPTask*) pGroup->m_lstTasks.GetHead ();
		ASSERT_VALID (pTask);

		CBCGPToolBoxPage* pPage = (CBCGPToolBoxPage*) CWnd::FromHandle (pTask->m_hwndTask);
		ASSERT_VALID (pPage);

		pTask->m_nWindowHeight = pPage->GetMaxHeight ();

#if (!defined _BCGSUITE_)
		pGroup->m_bIsVisible = pTask->m_nWindowHeight > 0;
#endif
	}

	RecalcLayout ();
	RedrawWindow (NULL, NULL, RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
}
//*********************************************************************************
void CBCGPToolBoxEx::OnContextMenu(CWnd* pWnd, CPoint point) 
{
	CPoint ptClient = point;
	ScreenToClient (&ptClient);

	CRect rectClient;
	GetClientRect (rectClient);

	if (!rectClient.PtInRect (ptClient) || !OnShowToolboxMenu (point, NULL, -1))
	{
		CBCGPTasksPane::OnContextMenu (pWnd, point);
	}
}
#if (!defined _BCGSUITE_)

HRESULT CBCGPToolBoxEx::get_accChildCount(long *pcountChildren)
{
	if (!pcountChildren)
	{
		return E_INVALIDARG;
	}

	m_nAccHeaderControlsCount = 0;

	if (m_wndToolBar.IsVisible())
	{
		m_nAccHeaderControlsCount++;
	}

#if (!defined _BCGSUITE_)
	if (m_bFilterBox)
	{
		m_nAccHeaderControlsCount++;
	}
#endif
	
	*pcountChildren = AccGetChildCount() + m_nAccHeaderControlsCount;
	return S_OK;
}
//****************************************************************************
HRESULT CBCGPToolBoxEx::get_accChild(VARIANT varChild, IDispatch** ppdispChild)
{
	if (ppdispChild == NULL)
	{
		return E_INVALIDARG;
	}
	if (m_bFilterBox && varChild.lVal == m_nAccHeaderControlsCount)
	{
		return AccessibleObjectFromWindow(m_wndFilter.GetSafeHwnd(), (DWORD) OBJID_CLIENT, IID_IAccessible, (void**) ppdispChild);
	}

	return  CBCGPTasksPane::get_accChild(varChild, ppdispChild);
}
//*******************************************************************************************************************************************
void CBCGPToolBoxEx::EnableSearchBox(BOOL bEnable, LPCTSTR lpszPrompt, LPCTSTR lpszOutOfFilter, BOOL bSimplifiedIcon)
{
	ASSERT_VALID(this);
	ASSERT(m_hWnd != NULL);	// Should be already created

	m_bFilterBox = bEnable;

	if (bEnable)
	{
		if (m_wndFilter.GetSafeHwnd() == NULL)
		{
			CEdit editDummy;
			editDummy.Create(WS_CHILD, CRect (0, 0, 100, 20), this, (UINT)-1);

			if (m_fontFilter.GetSafeHandle() == NULL)
			{
				LOGFONT lf;
				memset(&lf, 0, sizeof(LOGFONT));

				globalData.fontRegular.GetLogFont(&lf);
				m_fontFilter.CreateFontIndirect(&lf);
			}
			
			editDummy.SetFont(&m_fontFilter);
			
			CRect rectEdit;
			editDummy.GetWindowRect(rectEdit);
			
			int nFilterBoxHeight = rectEdit.Height() + globalUtils.ScaleByDPI(6);

			m_wndFilter.m_bVisualManagerStyle = TRUE;
			m_wndFilter.Create(WS_CHILD | WS_VISIBLE, CRect(0, 0, 100, nFilterBoxHeight), this, ID_FILTER);
			m_wndFilter.ModifyStyleEx(0, WS_EX_CLIENTEDGE);

			m_wndFilter.SetFont(&m_fontFilter);

			editDummy.DestroyWindow ();
		}

		CString strFilterPrompt;

		if (lpszPrompt == NULL)
		{
			CBCGPLocalResource locaRes;
			strFilterPrompt.LoadString(IDS_BCGBARRES_SEARCH_PROMPT);
		}
		else
		{
			strFilterPrompt = lpszPrompt;
		}

		if (lpszOutOfFilter == NULL)
		{
			CBCGPLocalResource locaRes;
			m_strOutOfFilter.LoadString(IDS_BCGBARRES_PROPLIST_OUT_OF_FILTER);
		}
		else
		{
			m_strOutOfFilter = lpszOutOfFilter;
		}
		
		m_wndFilter.SetSimplifiedBrowseIcon(bSimplifiedIcon);
		m_wndFilter.EnableSearchMode(TRUE, strFilterPrompt);

		m_hwndHeader = m_wndFilter.GetSafeHwnd();
	}
	else
	{
		if (m_wndFilter.GetSafeHwnd() != NULL)
		{
			m_wndFilter.DestroyWindow();
		}

		m_hwndHeader = NULL;
		SetToolsFilter(NULL);
	}

	AdjustLayout();
}
//****************************************************************************
void CBCGPToolBoxEx::SetToolsFilter(LPCTSTR lpszFilter)
{
	m_strFilter = lpszFilter == NULL ? _T("") : lpszFilter;
	m_strFilter.MakeUpper();
	m_strFilter.TrimLeft();
	m_strFilter.TrimRight();

	for (POSITION pos = m_lstTaskGroups.GetHeadPosition (); pos != NULL;)
	{
		CBCGPTasksGroup* pGroup = (CBCGPTasksGroup*) m_lstTaskGroups.GetNext (pos);
		ASSERT_VALID (pGroup);
		
		CBCGPTask* pTask = (CBCGPTask*) pGroup->m_lstTasks.GetHead ();
		ASSERT_VALID (pTask);
		
		CBCGPToolBoxPage* pPage = (CBCGPToolBoxPage*) CWnd::FromHandle (pTask->m_hwndTask);
		ASSERT_VALID (pPage);

		if (pPage->SetFilter(m_strFilter) && !m_strFilter.IsEmpty() && pGroup->m_bIsCollapsed)
		{
			pGroup->m_bIsCollapsed = FALSE;
		}
	}

	AdjustLayout();
}

#endif

BOOL CBCGPToolBoxEx::PreTranslateMessage(MSG* pMsg)
{
#if (!defined _BCGSUITE_)
	if (m_bFilterBox && pMsg->message == WM_KEYDOWN && m_wndFilter.GetSafeHwnd() == ::GetFocus())
	{
		if (globalUtils.ProcessCtrlEditAccelerators(&m_wndFilter, (UINT)pMsg->wParam, TRUE))
		{
			return TRUE;
		}

		if (pMsg->wParam == VK_ESCAPE)
		{
			m_wndFilter.SetWindowText(_T(""));
			SetFocus();

			return TRUE;
		}
	}
#endif
	return CBCGPTasksPane::PreTranslateMessage(pMsg);
}
//****************************************************************************
void CBCGPToolBoxEx::OnFilter()
{
#if (!defined _BCGSUITE_)
	CString strFilter;
	m_wndFilter.GetWindowText(strFilter);

	SetToolsFilter(strFilter);
#endif
}

#endif // BCGP_EXCLUDE_TASK_PANE
#endif // BCGP_EXCLUDE_TOOLBOX
