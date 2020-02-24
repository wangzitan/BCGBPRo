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
// BCGPFormView.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPFormView.h"
#include "BCGPVisualManager.h"

#ifndef _BCGSUITE_
#include "BCGPTooltipManager.h"
#include "BCGPPopupMenu.h"
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

IMPLEMENT_DYNAMIC(CBCGPFormView, CFormView)
IMPLEMENT_DYNAMIC(CBCGPScrollView, CScrollView)

/////////////////////////////////////////////////////////////////////////////
// CBCGPFormView

#pragma warning (disable : 4355)

CBCGPFormView::CBCGPFormView(LPCTSTR lpszTemplateName)
	: TBCGPInternalScrollBarWrapperWnd<CFormView>(lpszTemplateName),
	m_Impl (*this)
{
	m_bInitDlgCompleted = FALSE;
}

CBCGPFormView::CBCGPFormView(UINT nIDTemplate)
	: TBCGPInternalScrollBarWrapperWnd<CFormView>(nIDTemplate),
	m_Impl (*this)
{
	m_bInitDlgCompleted = FALSE;
}

#pragma warning (default : 4355)

CBCGPFormView::~CBCGPFormView()
{
}

BEGIN_MESSAGE_MAP(CBCGPFormView, CFormView)
	//{{AFX_MSG_MAP(CBCGPFormView)
	ON_WM_CTLCOLOR()
	ON_WM_ERASEBKGND()
	ON_WM_DESTROY()
	ON_WM_SIZE()
	ON_WM_CREATE()
	ON_WM_MOUSEMOVE()
	ON_WM_SETCURSOR()
	ON_WM_LBUTTONUP()
	//}}AFX_MSG_MAP
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)	
	ON_REGISTERED_MESSAGE(BCGM_UPDATETOOLTIPS, OnBCGUpdateToolTips)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnNeedTipText)
	ON_MESSAGE(WM_GESTURE, OnGestureEvent)
	ON_MESSAGE(WM_TABLET_QUERYSYSTEMGESTURESTATUS, OnTabletQuerySystemGestureStatus)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPFormView diagnostics

#ifdef _DEBUG
void CBCGPFormView::AssertValid() const
{
	CFormView::AssertValid();
}

void CBCGPFormView::Dump(CDumpContext& dc) const
{
	CFormView::Dump(dc);
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CBCGPFormView message handlers

//*******************************************************************************************************************
LRESULT CBCGPFormView::OnGestureEvent(WPARAM wp, LPARAM lp)
{
	if (ProcessGestureEvent(this, wp, lp))
	{
		return Default();
	}

	return 0;
}
//******************************************************************************************************************
LRESULT CBCGPFormView::OnTabletQuerySystemGestureStatus(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	return (LRESULT)TABLET_DISABLE_PRESSANDHOLD;
}
//******************************************************************************************************************
HBRUSH CBCGPFormView::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	if (IsVisualManagerStyle () || (IsWhiteBackground() && !globalData.IsHighContastMode() && !CBCGPVisualManager::GetInstance ()->IsDarkTheme()))
	{
		HBRUSH hbr = m_Impl.OnCtlColor (pDC, pWnd, nCtlColor);
		if (hbr != NULL)
		{
			return hbr;
		}
	}	

	return CFormView::OnCtlColor(pDC, pWnd, nCtlColor);
}
//************************************************************************
BOOL CBCGPFormView::OnEraseBkgnd(CDC* pDC) 
{
	CRect rectClient;
	GetClientRect (rectClient);

	BOOL bRes = FALSE;

	if (IsWhiteBackground() && !globalData.IsHighContastMode() && !CBCGPVisualManager::GetInstance ()->IsDarkTheme())
	{
		pDC->FillSolidRect(rectClient, RGB(255, 255, 255));
		bRes = TRUE;
	}
	else if (IsVisualManagerStyle () && CBCGPVisualManager::GetInstance ()->OnFillDialog (pDC, this, rectClient))
	{
		bRes = TRUE;
	}
	else
	{
		bRes = CFormView::OnEraseBkgnd(pDC);
	}

	m_Impl.DrawGroupBoxes(pDC);
	m_Impl.DrawControlInfoTips(pDC);

	return bRes;
}
//************************************************************************
void CBCGPFormView::EnableVisualManagerStyle (BOOL bEnable, const CList<UINT,UINT>* plstNonSubclassedItems)
{
	ASSERT_VALID (this);

	m_Impl.EnableVisualManagerStyle (bEnable, FALSE, plstNonSubclassedItems);
	m_Impl.m_bTransparentStaticCtrls = FALSE;
}
//************************************************************************
void CBCGPFormView::SetWhiteBackground(BOOL bSet)
{
	ASSERT_VALID (this);

	m_Impl.m_bIsWhiteBackground = bSet && !globalData.IsHighContastMode();
}
//************************************************************************
void CBCGPFormView::SetGroupBoxesDrawByParent(BOOL bSet /*= TRUE*/)
{
	if (GetSafeHwnd() != NULL)
	{
		ASSERT(FALSE);
		return;
	}
	
	m_Impl.m_bGroupBoxesDrawByParent = bSet;
}
//************************************************************************
void CBCGPFormView::OnInitialUpdate() 
{
	if (!m_bInitDlgCompleted)
	{
		m_Impl.EnableGesturePan();
	}

	m_bInitDlgCompleted = TRUE;

	CFormView::OnInitialUpdate();
	
	if (IsVisualManagerStyle ())
	{
		m_Impl.EnableVisualManagerStyle (TRUE);
	}

	m_Impl.OnInit();

	if (!m_Impl.m_mapCtrlInfoTip.IsEmpty())
	{
		RedrawWindow();
	}
}
//************************************************************************
void CBCGPFormView::OnDestroy() 
{
	m_Impl.OnDestroy ();
	CFormView::OnDestroy();
}
//************************************************************************
void CBCGPFormView::OnSize(UINT nType, int cx, int cy) 
{
	CFormView::OnSize(nType, cx, cy);

	AdjustControlsLayout();
	m_Impl.UpdateToolTipsRect();
}
//************************************************************************
int CBCGPFormView::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CFormView::OnCreate(lpCreateStruct) == -1)
		return -1;
	
    return m_Impl.OnCreate();
}
//************************************************************************
LRESULT CBCGPFormView::OnChangeVisualManager (WPARAM, LPARAM)
{
	if (IsVisualManagerStyle ())
	{
		m_Impl.OnChangeVisualManager ();
	}

	return 0L;
}
//************************************************************************
BOOL CBCGPFormView::IsInternalScrollBarThemed() const
{
#ifndef _BCGSUITE_
	return (globalData.m_nThemedScrollBars & BCGP_THEMED_SCROLLBAR_FORMVIEW) != 0 && m_Impl.m_bVisualManagerStyle;
#else
	return FALSE;
#endif
}
//*************************************************************************************
void CBCGPFormView::SetControlInfoTip(UINT nCtrlID, LPCTSTR lpszInfoTip, DWORD dwVertAlign, BOOL bRedrawInfoTip, CBCGPControlInfoTip::BCGPControlInfoTipStyle style, BOOL bIsClickable, const CPoint& ptOffset)
{
	m_Impl.SetControlInfoTip(nCtrlID, lpszInfoTip, dwVertAlign, bRedrawInfoTip, style, bIsClickable, ptOffset);
}
//*************************************************************************************
void CBCGPFormView::SetControlInfoTip(CWnd* pWndCtrl, LPCTSTR lpszInfoTip, DWORD dwVertAlign, BOOL bRedrawInfoTip, CBCGPControlInfoTip::BCGPControlInfoTipStyle style, BOOL bIsClickable, const CPoint& ptOffset)
{
	m_Impl.SetControlInfoTip(pWndCtrl, lpszInfoTip, dwVertAlign, bRedrawInfoTip, style, bIsClickable, ptOffset);
}
//*************************************************************************************
BOOL CBCGPFormView::OnNeedTipText(UINT id, NMHDR* pNMH, LRESULT* pResult)
{
	return m_Impl.OnNeedTipText(id, pNMH, pResult);
}
//**************************************************************************
LRESULT CBCGPFormView::OnBCGUpdateToolTips (WPARAM wp, LPARAM)
{
	UINT nTypes = (UINT) wp;

	if (nTypes & BCGP_TOOLTIP_TYPE_DEFAULT)
	{
		m_Impl.CreateTooltipInfo();
	}

	return 0;
}
//**************************************************************************
BOOL CBCGPFormView::PreTranslateMessage(MSG* pMsg) 
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
	case WM_MOUSEMOVE:
		if (m_Impl.m_pToolTipInfo->GetSafeHwnd () != NULL)
		{
			m_Impl.m_pToolTipInfo->RelayEvent(pMsg);
		}
		break;
	}
	
	switch (pMsg->message)
	{
	case WM_SYSKEYDOWN:
		m_Impl.SetUnderlineKeyboardShortcuts(TRUE, TRUE);
		break;
		
	case WM_KEYDOWN:
		if (m_Impl.ProcessFocusEditAccelerators((UINT)pMsg->wParam))
		{
			return TRUE;
		}
	}

	return CFormView::PreTranslateMessage(pMsg);
}
//**************************************************************************
void CBCGPFormView::OnMouseMove(UINT nFlags, CPoint point) 
{
	CFormView::OnMouseMove(nFlags, point);

	if (m_Impl.m_pToolTipInfo->GetSafeHwnd () != NULL && !m_Impl.m_mapCtrlInfoTip.IsEmpty())
	{
		CString strDescription;
		CString strTipText;
		BOOL bIsClickable = FALSE;
		
		HWND hwnd = m_Impl.GetControlInfoTipFromPoint(point, strTipText, strDescription, bIsClickable);
		
		if (hwnd != m_Impl.m_hwndInfoTipCurr)
		{
			m_Impl.m_pToolTipInfo->Pop();
			m_Impl.m_hwndInfoTipCurr = hwnd;
		}
	}
}
//**************************************************************************
BOOL CBCGPFormView::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	if (m_Impl.OnSetCursor())
	{
		return TRUE;
	}
	
	return CFormView::OnSetCursor(pWnd, nHitTest, message);
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPScrollView view

BOOL CBCGPScrollView::IsInternalScrollBarThemed() const
{
#ifndef _BCGSUITE_
	return (globalData.m_nThemedScrollBars & BCGP_THEMED_SCROLLBAR_SCROLLVIEW) != NULL;
#else
	return FALSE;
#endif
}
void CBCGPFormView::OnLButtonUp(UINT nFlags, CPoint point) 
{
	if (m_Impl.ProcessInfoTipClick(point))
	{
		return;
	}
	
	CFormView::OnLButtonUp(nFlags, point);
}
//*********************************************************************************
void CBCGPFormView::OnActivateView(BOOL bActivate, CView* pActivateView, CView* pDeactiveView) 
{
	if (!bActivate)
	{
		m_Impl.SetUnderlineKeyboardShortcuts(FALSE, TRUE);
	}
	
	CFormView::OnActivateView(bActivate, pActivateView, pDeactiveView);
}
//*********************************************************************************
BOOL CBCGPFormView::OnGestureEventPan(const CPoint& ptFrom, const CPoint& ptTo, CSize& sizeOverPan)
{
	BOOL bResult = FALSE;

	if (CBCGPPopupMenu::GetActiveMenu() != NULL)
	{
		return bResult;
	}

	SendMessageToDescendants(WM_CANCELMODE);

	if (ptTo.x != ptFrom.x)
	{
		bResult = OnScrollBy(CSize(ptFrom.x - ptTo.x, 0), TRUE);
		if (!bResult)
		{
			sizeOverPan.cx = ptFrom.x < ptTo.x ? m_lineDev.cx : -m_lineDev.cx;
		}
	}
	else if (ptTo.y != ptFrom.y)
	{
		bResult = OnScrollBy(CSize(0, ptFrom.y - ptTo.y), TRUE);
		if (!bResult)
		{
			sizeOverPan.cy = ptFrom.y < ptTo.y ? m_lineDev.cy : -m_lineDev.cy;
		}
	}

	if (bResult)
	{
		UpdateWindow();
	}

	return bResult;
}
