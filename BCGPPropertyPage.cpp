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
// BCGPPropertyPage.cpp : implementation file
//

#include "stdafx.h"
#include "bcgcbpro.h"

#ifndef _BCGSUITE_
#include "BCGPPopupMenu.h"
#include "BCGPToolbarMenuButton.h"
#include "BCGPTooltipManager.h"
#include "BCGPPngImage.h"
#endif

#include "BCGPPropertySheet.h"
#include "BCGPPropertyPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define UM_REDRAWHEADER	(WM_USER + 1003)

/////////////////////////////////////////////////////////////////////////////
// CBCGPPropertyPage

IMPLEMENT_DYNCREATE(CBCGPPropertyPage, CPropertyPage)

#pragma warning (disable : 4355)

CBCGPPropertyPage::CBCGPPropertyPage() :
	m_Impl (*this)
{
	CommonInit ();
}

CBCGPPropertyPage::CBCGPPropertyPage(UINT nIDTemplate, UINT nIDCaption) :
	CPropertyPage (nIDTemplate, nIDCaption),
	m_Impl (*this)
{
	CommonInit ();
}

CBCGPPropertyPage::CBCGPPropertyPage(LPCTSTR lpszTemplateName, UINT nIDCaption) :
	CPropertyPage (lpszTemplateName, nIDCaption),
	m_Impl (*this)
{
	CommonInit ();
}

#pragma warning (default : 4355)

void CBCGPPropertyPage::CommonInit ()
{
	m_pCategory = NULL;
	m_nIcon = -1;
	m_nSelIconNum = -1;
	m_hIcon = NULL;
	m_hTreeNode = NULL;
	m_bIsAeroWizardPage = FALSE;
	m_rectHeader.SetRectEmpty();
	m_rectBackstageWatermark.SetRectEmpty();
	m_pBackstageWatermark = NULL;
}

CBCGPPropertyPage::~CBCGPPropertyPage()
{
	if (m_hIcon != NULL)
	{
		::DestroyIcon(m_hIcon);
	}
}

BEGIN_MESSAGE_MAP(CBCGPPropertyPage, CPropertyPage)
	//{{AFX_MSG_MAP(CBCGPPropertyPage)
	ON_WM_ACTIVATE()
	ON_WM_NCACTIVATE()
	ON_WM_ERASEBKGND()
	ON_WM_CTLCOLOR()
	ON_WM_DESTROY()
	ON_WM_NCPAINT()
	ON_WM_SIZE()
	ON_WM_CREATE()
	ON_WM_LBUTTONDOWN()
	ON_WM_SETCURSOR()
	ON_WM_LBUTTONUP()
	//}}AFX_MSG_MAP
	ON_MESSAGE(UM_REDRAWHEADER, OnRedrawHeader)
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)
	ON_REGISTERED_MESSAGE(BCGM_UPDATETOOLTIPS, OnBCGUpdateToolTips)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnNeedTipText)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPPropertyPage message handlers

void CBCGPPropertyPage::OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized) 
{
	m_Impl.OnActivate (nState, pWndOther);
	CPropertyPage::OnActivate (nState, pWndOther, bMinimized);
}
//****************************************************************************************
BOOL CBCGPPropertyPage::OnNcActivate(BOOL bActive) 
{
	m_Impl.OnNcActivate (bActive);

	//-----------------------------------------------------------
	// Do not call the base class because it will call Default()
	// and we may have changed bActive.
	//-----------------------------------------------------------
	return (BOOL) DefWindowProc (WM_NCACTIVATE, bActive, 0L);
}
//**************************************************************************************
void CBCGPPropertyPage::SetActiveMenu (CBCGPPopupMenu* pMenu)
{
	m_Impl.SetActiveMenu (pMenu);
}
//***************************************************************************************
BOOL CBCGPPropertyPage::PreTranslateMessage(MSG* pMsg) 
{
	if (m_Impl.PreTranslateMessage (pMsg))
	{
		return TRUE;
	}

	return CPropertyPage::PreTranslateMessage(pMsg);
}
//****************************************************************************************
BOOL CBCGPPropertyPage::OnCommand(WPARAM wParam, LPARAM lParam) 
{
	if (m_Impl.OnCommand (wParam, lParam))
	{
		return TRUE;
	}

	return CPropertyPage::OnCommand(wParam, lParam);
}
//****************************************************************************************
BOOL CBCGPPropertyPage::OnSetActive() 
{
	CBCGPPropertySheet* pParent = DYNAMIC_DOWNCAST(CBCGPPropertySheet, GetParent ());
	if (pParent != NULL)
	{
		pParent->OnActivatePage (this);
	}
	
	AdjustControlsLayout();
	return CPropertyPage::OnSetActive();
}
//****************************************************************************************
BOOL CBCGPPropertyPage::OnKillActive() 
{
	m_Impl.SetUnderlineKeyboardShortcuts(FALSE, FALSE);
	return CPropertyPage::OnKillActive();
}
//****************************************************************************************
BOOL CBCGPPropertyPage::OnInitDialog() 
{
	BOOL bRes = CPropertyPage::OnInitDialog();
	
	if (IsVisualManagerStyle ())
	{
		m_Impl.EnableVisualManagerStyle (TRUE);
	}

	if (IsBackstageMode())
	{
		m_Impl.EnableBackstageMode();
	}

	m_Impl.OnInit();

	CBCGPPropertySheet* pParent = DYNAMIC_DOWNCAST(CBCGPPropertySheet, GetParent ());
	if (pParent == NULL)
	{
		return bRes;
	}

	m_bIsAeroWizardPage = pParent->GetAeroHeight() > 0;

	if (m_bIsAeroWizardPage && globalData.IsHighContastMode())
	{
		m_Impl.EnableVisualManagerStyle (TRUE);
	}

	m_Impl.m_bIsWhiteBackground = m_bIsAeroWizardPage && !globalData.IsHighContastMode();

	int nHeaderHeight = pParent->GetHeaderHeight ();
	if (m_bIsAeroWizardPage && pParent->m_bDrawHeaderOnAeroCaption)
	{
		nHeaderHeight = 0;
	}

	if (GetLayout () != NULL)
	{
		CRect rectClient;
		GetClientRect (rectClient);

		if (pParent->GetLook() != CBCGPPropertySheet::PropSheetLook_Tabs)
		{
			rectClient.top += nHeaderHeight;
		}

		GetLayout ()->SetMinSize (rectClient.Size ());
	}

	if (nHeaderHeight <= 1)
	{
		return bRes;
	}

	CWnd* pWndChild = GetWindow (GW_CHILD);
	while (pWndChild != NULL)
	{
		CRect rectChild;
		pWndChild->GetWindowRect (rectChild);
		ScreenToClient (rectChild);

		rectChild.OffsetRect (0, nHeaderHeight);

		pWndChild->SetWindowPos (NULL, rectChild.left, rectChild.top,
			-1, -1, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);

		pWndChild = pWndChild->GetNextWindow ();
	}

	return bRes;
}
//*****************************************************************************************
void CBCGPPropertyPage::DoDrawBackground(CDC* pDC)
{
	ASSERT_VALID(pDC);

	CRect rectClient;
	GetClientRect (rectClient);

	m_Impl.DrawBackgroundImage(pDC, rectClient);
	
	if (!m_rectHeader.IsRectEmpty () && m_rectHeader.Height() > 1)
	{
		CBCGPPropertySheet* pParent = DYNAMIC_DOWNCAST(CBCGPPropertySheet, GetParent ());
		if (pParent != NULL)
		{
			pParent->OnDrawPageHeader (pDC, pParent->GetPageIndex (this), m_rectHeader);
		}
	}
	
	m_Impl.DrawGroupBoxes(pDC);
	m_Impl.DrawControlInfoTips(pDC);
}
//****************************************************************************************
BOOL CBCGPPropertyPage::OnEraseBkgnd(CDC* pDC) 
{
	BOOL bRes = TRUE;

	CRect rectClient;
	GetClientRect (rectClient);

	if (IsVisualManagerStyle ())
	{
#ifndef _BCGSUITE_
		if (IsBackstageMode())
		{
			CBCGPMemDC memDC(*pDC, this);
			CDC* pDCMem = &memDC.GetDC();

			CBCGPVisualManager::GetInstance ()->OnFillRibbonBackstageForm(pDCMem, this, rectClient);

			if (!m_rectBackstageWatermark.IsRectEmpty() && m_pBackstageWatermark != NULL && !globalData.IsHighContastMode())
			{
				ASSERT_VALID(m_pBackstageWatermark);
				m_pBackstageWatermark->DrawEx(pDCMem, m_rectBackstageWatermark, 0);
			}

			DoDrawBackground(pDCMem);
			return TRUE;
		}
		else
#endif
		{
			if (!CBCGPVisualManager::GetInstance ()->OnFillDialog (pDC, this, rectClient))
			{
				bRes = CPropertyPage::OnEraseBkgnd (pDC);
			}
		}
	}
	else if (m_bIsAeroWizardPage && !globalData.IsHighContastMode())
	{
		pDC->FillSolidRect(rectClient, RGB(255, 255, 255));
		bRes = TRUE;
	}
	else
	{
		bRes = CPropertyPage::OnEraseBkgnd(pDC);
	}

	DoDrawBackground(pDC);
	return bRes;
}
//*****************************************************************************************
void CBCGPPropertyPage::EnableVisualManagerStyle (BOOL bEnable, const CList<UINT,UINT>* plstNonSubclassedItems)
{
	ASSERT_VALID (this);

	m_Impl.EnableVisualManagerStyle (bEnable, FALSE, plstNonSubclassedItems);
}
//*****************************************************************************************
void CBCGPPropertyPage::SetGroupBoxesDrawByParent(BOOL bSet /*= TRUE*/)
{
	if (GetSafeHwnd() != NULL)
	{
		ASSERT(FALSE);
		return;
	}
	
	m_Impl.m_bGroupBoxesDrawByParent = bSet;
}
//*****************************************************************************************
HBRUSH CBCGPPropertyPage::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	if (IsVisualManagerStyle () || m_bIsAeroWizardPage)
	{
		HBRUSH hbr = m_Impl.OnCtlColor (pDC, pWnd, nCtlColor);
		if (hbr != NULL)
		{
			return hbr;
		}
	}	

	return CPropertyPage::OnCtlColor(pDC, pWnd, nCtlColor);
}
//*****************************************************************************************
void CBCGPPropertyPage::OnDestroy() 
{
	m_Impl.OnDestroy ();
	CPropertyPage::OnDestroy();
}
//*****************************************************************************************
void CBCGPPropertyPage::OnNcPaint() 
{
	if (!IsVisualManagerStyle ())
	{
		Default ();
	}
}
//*****************************************************************************************
BOOL CBCGPPropertyPage::OnApply() 
{
	BOOL bRes = CPropertyPage::OnApply();

	CBCGPPropertySheet* pParent = DYNAMIC_DOWNCAST(CBCGPPropertySheet, GetParent ());
	if (pParent != NULL)
	{
		ASSERT_VALID (pParent);
		pParent->OnApply ();
	}

	return bRes;
}
//*****************************************************************************************
void CBCGPPropertyPage::EnableLayout(BOOL bEnable, CRuntimeClass* pRTC)
{
	m_Impl.EnableLayout(bEnable, pRTC, FALSE);
}
//*****************************************************************************************
void CBCGPPropertyPage::OnSize(UINT nType, int cx, int cy) 
{
	CPropertyPage::OnSize(nType, cx, cy);

	CBCGPPropertySheet* pParent = DYNAMIC_DOWNCAST(CBCGPPropertySheet, GetParent ());
	if (pParent == NULL || pParent->m_bIsReady)
	{
		AdjustControlsLayout();
	}

	CRect rectHeaderOld = m_rectHeader;

	SetHeaderRect();

	m_Impl.UpdateToolTipsRect();

	if (!m_rectHeader.IsRectEmpty() && rectHeaderOld != m_rectHeader)
	{
		PostMessage(UM_REDRAWHEADER);
	}
}
//*****************************************************************************************
LRESULT CBCGPPropertyPage::OnRedrawHeader(WPARAM, LPARAM)
{
	InvalidateRect(m_rectHeader);
	UpdateWindow();

	return 0;
}
//*****************************************************************************************
void CBCGPPropertyPage::SetHeaderRect()
{
	m_rectHeader.SetRectEmpty();

	CRect rectClient;
	GetClientRect (rectClient);

	CBCGPPropertySheet* pParent = DYNAMIC_DOWNCAST(CBCGPPropertySheet, GetParent ());
	if (pParent != NULL && pParent->GetHeaderHeight () > 0 && !pParent->m_bDrawHeaderOnAeroCaption)
	{
		m_rectHeader = rectClient;
		m_rectHeader.bottom = m_rectHeader.top + pParent->GetHeaderHeight ();

		if (pParent->GetLook () == CBCGPPropertySheet::PropSheetLook_OutlookBar ||
			pParent->GetLook () == CBCGPPropertySheet::PropSheetLook_Tree ||
			pParent->GetLook () == CBCGPPropertySheet::PropSheetLook_List)
		{
			CRect rectParent;
			pParent->GetWindowRect (rectParent);

			ScreenToClient (rectParent);

			m_rectHeader.right = rectParent.right - 
				::GetSystemMetrics (pParent->IsLayoutEnabled () ? SM_CXFRAME : SM_CXDLGFRAME);
		}
	}
}
//*****************************************************************************************
int CBCGPPropertyPage::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CPropertyPage::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	CBCGPPropertySheet* pParent = DYNAMIC_DOWNCAST(CBCGPPropertySheet, GetParent ());
	if (pParent == NULL)
	{
		return 0;
	}

	if (pParent->IsLayoutEnabled())
	{
		m_Impl.EnableLayout(TRUE, pParent->GetLayout()->GetRuntimeClass(), FALSE);

		if (pParent->m_bFlickerLess)
		{
			ModifyStyle(0, WS_CLIPCHILDREN | WS_CLIPSIBLINGS);
		}
	}

	return m_Impl.OnCreate();
}
//**************************************************************************
LRESULT CBCGPPropertyPage::OnChangeVisualManager (WPARAM, LPARAM)
{
	if (IsVisualManagerStyle ())
	{
		m_Impl.OnChangeVisualManager ();
	}

	return 0;
}
//**************************************************************************
void CBCGPPropertyPage::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CBCGPPropertySheet* pParent = DYNAMIC_DOWNCAST(CBCGPPropertySheet, GetParent ());
	if (pParent->GetSafeHwnd() != NULL && pParent->IsDragClientAreaEnabled())
	{
		CString strInfo;
		CString strDescription;
		BOOL bIsClickable = FALSE;
		
		if (m_Impl.GetControlInfoTipFromPoint(point, strInfo, strDescription, bIsClickable) == NULL)
		{
			CPoint ptScreen = point;
			ClientToScreen(&ptScreen);

			pParent->SendMessage (WM_NCLBUTTONDOWN, (WPARAM) HTCAPTION, MAKELPARAM (ptScreen.x, ptScreen.y));
			return;
		}
	}
	
	CPropertyPage::OnLButtonDown(nFlags, point);
}
//**************************************************************************
void CBCGPPropertyPage::SetPageIcon(HICON hIcon)
{
	if (m_hIcon != NULL)
	{
		::DestroyIcon(m_hIcon);
	}
	
	if (hIcon != NULL)
	{
		m_hIcon = ::CopyIcon (hIcon);
	}
	else
	{
		m_hIcon = NULL;
	}
}
//*************************************************************************************
void CBCGPPropertyPage::SetControlInfoTip(UINT nCtrlID, LPCTSTR lpszInfoTip, DWORD dwVertAlign, BOOL bRedrawInfoTip, CBCGPControlInfoTip::BCGPControlInfoTipStyle style, BOOL bIsClickable, const CPoint& ptOffset)
{
	m_Impl.SetControlInfoTip(nCtrlID, lpszInfoTip, dwVertAlign, bRedrawInfoTip, style, bIsClickable, ptOffset);
}
//*************************************************************************************
void CBCGPPropertyPage::SetControlInfoTip(CWnd* pWndCtrl, LPCTSTR lpszInfoTip, DWORD dwVertAlign, BOOL bRedrawInfoTip, CBCGPControlInfoTip::BCGPControlInfoTipStyle style, BOOL bIsClickable, const CPoint& ptOffset)
{
	m_Impl.SetControlInfoTip(pWndCtrl, lpszInfoTip, dwVertAlign, bRedrawInfoTip, style, bIsClickable, ptOffset);
}
//*************************************************************************************
BOOL CBCGPPropertyPage::OnNeedTipText(UINT id, NMHDR* pNMH, LRESULT* pResult)
{
	return m_Impl.OnNeedTipText(id, pNMH, pResult);
}
//**************************************************************************
LRESULT CBCGPPropertyPage::OnBCGUpdateToolTips (WPARAM wp, LPARAM)
{
	UINT nTypes = (UINT) wp;

	if (nTypes & BCGP_TOOLTIP_TYPE_DEFAULT)
	{
		m_Impl.CreateTooltipInfo();
	}

	return 0;
}
//**************************************************************************
BOOL CBCGPPropertyPage::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	if (m_Impl.OnSetCursor())
	{
		return TRUE;
	}
	
	return CPropertyPage::OnSetCursor(pWnd, nHitTest, message);
}
//****************************************************************************
void CBCGPPropertyPage::OnLButtonUp(UINT nFlags, CPoint point) 
{
	if (m_Impl.ProcessInfoTipClick(point))
	{
		return;
	}
	
	CPropertyPage::OnLButtonUp(nFlags, point);
}
//*************************************************************************************
void CBCGPPropertyPage::SetBackgroundImage (HBITMAP hBitmap, 
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
BOOL CBCGPPropertyPage::SetBackgroundImage (UINT uiBmpResId,
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
//*************************************************************************************
BOOL CBCGPPropertyPage::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult) 
{
	ASSERT(lParam != 0);
	NMHDR* pNMHDR = (NMHDR*)lParam;

	if (pNMHDR->code == PSN_KILLACTIVE)
	{
		CBCGPPropertySheet* pParent = DYNAMIC_DOWNCAST(CBCGPPropertySheet, GetParent());
		if (pParent != NULL && !pParent->m_bIsReady && pParent->GetSafeHwnd() == pNMHDR->hwndFrom)
		{
			return FALSE;
		}
	}
	
	return CPropertyPage::OnNotify(wParam, lParam, pResult);
}
