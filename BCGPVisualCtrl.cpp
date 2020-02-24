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
// BCGPVisualCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPVisualCtrl.h"
#include "bcgglobals.h"
#include "trackmouse.h"
#include "BCGPDrawManager.h"
#include "BCGPMath.h"
#include "BCGPGraphicsManagerGDI.h"
#include "BCGPGraphicsManagerGdiPlus.h"
#include "BCGPGridCtrl.h"
#include "BCGPPopupWindow.h"
#include "BCGPGlobalUtils.h"

#if (!defined _BCGSUITE_) && (!defined _BCGPCHART_STANDALONE)
	#include "BCGPPopupMenu.h"
	#include "BCGPTooltipManager.h"
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPBaseVisualCtrl

IMPLEMENT_DYNAMIC(CBCGPBaseVisualCtrl, CStatic)
IMPLEMENT_DYNAMIC(CBCGPVisualCtrl, CBCGPBaseVisualCtrl)
IMPLEMENT_SERIAL(CBCGPVisualContainerCtrl, CBCGPBaseVisualCtrl, 1)

HHOOK CBCGPBaseVisualCtrl::m_hookMouse = NULL;
HWND CBCGPBaseVisualCtrl::m_hwndHookedPopup = NULL;
BOOL CBCGPBaseVisualCtrl::m_bInPaint = FALSE;

CBCGPBaseVisualCtrl::CBCGPBaseVisualCtrl()
{
	m_pGM = NULL;

	m_dblScale = 0.0;
	
	m_bIsCaptured = FALSE;
	m_bIsTracked = FALSE;

	m_bIsTooltip = FALSE;
	m_bToolTipCleared = FALSE;
	m_bTooltipTrackingMode = FALSE;
	m_bUsePopupWindowInTrackingMode = FALSE;
	m_pToolTip = NULL;
	m_pInfoTip = NULL;

	m_bIsFocused = FALSE;
	m_nDlgCode = 0;

	m_bIsPopup = FALSE;
	m_nPopupAlpha = 255;
	m_nPopupFlags = 0;
	m_nPopupAlphaAnimated = (BYTE)-1;
	
	m_bIgnoreMinSize = FALSE;
}

CBCGPBaseVisualCtrl::~CBCGPBaseVisualCtrl()
{
	if (m_pGM != NULL)
	{
		delete m_pGM;
	}

#if _MSC_VER >= 1300
	if (m_pProxy != NULL)
	{
		//force disconnect accessibility clients
		::CoDisconnectObject ((IAccessible*)m_pProxy, NULL);
	}
#endif
}

BEGIN_MESSAGE_MAP(CBCGPBaseVisualCtrl, CStatic)
	//{{AFX_MSG_MAP(CBCGPBaseVisualCtrl)
	ON_WM_ERASEBKGND()
	ON_WM_PAINT()
	ON_WM_SIZE()
	ON_WM_CREATE()
	ON_WM_SETCURSOR()
	ON_WM_ENABLE()
	ON_WM_DESTROY()
	ON_WM_GETDLGCODE()
	ON_WM_KILLFOCUS()
	ON_WM_SETFOCUS()
	ON_WM_KEYDOWN()
	ON_WM_KEYUP()
	ON_WM_NCDESTROY()
	ON_WM_NCCREATE()
	ON_WM_NCCALCSIZE()
	ON_WM_NCLBUTTONDOWN()
	ON_WM_NCLBUTTONUP()
	ON_WM_NCPAINT()
	ON_WM_NCHITTEST()
	//}}AFX_MSG_MAP
#if (!defined _BCGSUITE_) && (!defined _BCGPCHART_STANDALONE)
	ON_REGISTERED_MESSAGE(BCGM_UPDATETOOLTIPS, OnBCGUpdateToolTips)
	ON_REGISTERED_MESSAGE(BCGM_POSTREDRAW, OnBCGPostRedraw)
#endif
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnTTNeedTipText)
	ON_MESSAGE(WM_FLOATSTATUS, OnFloatStatus)
	ON_MESSAGE(WM_GESTURE, OnGestureEvent)
	ON_MESSAGE(WM_TABLET_QUERYSYSTEMGESTURESTATUS, OnTabletQuerySystemGestureStatus)
	ON_MESSAGE(WM_PRINTCLIENT, OnPrintClient)
	ON_MESSAGE(WM_PRINT, OnPrint)
	ON_MESSAGE(WM_GETOBJECT, OnGetObject)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPBaseVisualCtrl message handlers

BOOL CBCGPBaseVisualCtrl::Create(const RECT& rect, CWnd* pParentWnd, UINT nID, DWORD dwStyle, CCreateContext* pContext)
{
	CString strClassName = globalData.RegisterWindowClass (_T("BCGPVisualCtrl"));
	dwStyle |= (WS_CLIPCHILDREN | WS_CLIPSIBLINGS);

	return CWnd::Create (strClassName, _T (""),
							dwStyle, 
							rect,
							pParentWnd,
							nID,
							pContext);
}
//***********************************************************************************************************
BOOL CBCGPBaseVisualCtrl::CreatePopup(const CRect& rect, BYTE nTransparency, CWnd* pWndOwner, UINT nFlags)
{
#if !defined _BCGPCHART_STANDALONE
	CString strClassName = globalData.RegisterWindowClass (_T("BCGPVisualCtrl"));

	m_bIsPopup = TRUE;
	m_nPopupAlpha = nTransparency;
	m_nPopupFlags = nFlags;

#if _MSC_VER < 1700 || !defined(_BCGSUITE_)
    DWORD dwExStyle = (globalData.IsWindowsLayerSupportAvailable () && globalData.m_nBitsPerPixel > 8) ? 
		WS_EX_LAYERED : 0;
#else
    DWORD dwExStyle = globalData.m_nBitsPerPixel > 8 ? WS_EX_LAYERED : 0;
#endif

	if (!CWnd::CreateEx (dwExStyle | WS_EX_TOOLWINDOW, 
								strClassName, 
								_T (""),
								WS_POPUP | MFS_SYNCACTIVE, 
								rect,
								pWndOwner == NULL ? AfxGetMainWnd() : pWndOwner,
								0))
	{
		return FALSE;
	}

	if (pWndOwner != NULL)
	{
		SetOwner(pWndOwner);
	}

	if (CBCGPPopupMenu::GetActiveMenu() != NULL && ::IsWindow (CBCGPPopupMenu::GetActiveMenu()->m_hWnd))
	{
		CBCGPPopupMenu::GetActiveMenu()->SendMessage (WM_CLOSE);
	}

	ShowWindow(SW_SHOWNOACTIVATE);

	if ((nFlags & BCGP_VISUAL_POPUP_DONT_CLOSE_ON_MOUSE_CLICK) == 0)
	{
		if (m_hwndHookedPopup != NULL)
		{
			::SendMessage(m_hwndHookedPopup, WM_CLOSE, 0, 0);
		}

		ASSERT(m_hookMouse == NULL);
		m_hookMouse = ::SetWindowsHookEx(WH_MOUSE, MouseProc, 0, GetCurrentThreadId());

		m_hwndHookedPopup = GetSafeHwnd();
	}

	if (dwExStyle & WS_EX_LAYERED)
	{
		OnDrawLayeredPopup();
	}

	SetFocus();

	return TRUE;
#else
	return FALSE;
#endif
}
//***********************************************************************************************************
BOOL CBCGPBaseVisualCtrl::OnEraseBkgnd(CDC* /*pDC*/) 
{
	return TRUE;
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnPaint() 
{
	CPaintDC dc(this); // device context for painting

	if (m_pGM == NULL)
	{
		m_pGM = CBCGPGraphicsManager::CreateInstance();
	}

	if (m_pGM == NULL)
	{
		return;
	}

	double dblMemDCScale = m_pGM->IsSupported(BCGP_GRAPHICS_MANAGER_SCALING) ? 1.0 : m_dblScale;
	
	CBCGPMemDC memDC (dc, this, 0, dblMemDCScale);
	CDC* pDC = &memDC.GetDC ();

	CRect rectClip(0, 0, 0, 0);
	if (!IsDirty())
	{
		dc.GetClipBox(rectClip);
	}

	DoPaint(pDC, rectClip);
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::DoEraseBackground(CDC* pDC, const CRect& rectClient)
{
	ASSERT_VALID(pDC);

	pDC->FillRect(rectClient, &globalData.brWindow);
	globalData.DrawParentBackground (this, pDC);
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::DoPaint(CDC* pDC, CRect rectClip)
{
	CRect rectClient;
	GetClientRect(rectClient);

	m_bInPaint = TRUE;
	
	DoEraseBackground(pDC, rectClient);

	CBCGPRect rect = GetRect();

	if (m_pGM == NULL)
	{
		m_pGM = CBCGPGraphicsManager::CreateInstance();
	}

	if (m_pGM == NULL)
	{
		m_bInPaint = FALSE;
		return;
	}

	if (rect.IsRectEmpty())
	{
		SetRect(rectClient);
		rect = GetRect();
	}

	m_pGM->BindDC(pDC, rect);

	if (m_dblScale != 0.0 && m_pGM->IsSupported(BCGP_GRAPHICS_MANAGER_SCALING))
	{
		m_pGM->SetScale(m_dblScale);
	}

	if (!m_pGM->BeginDraw())
	{
		return;
	}

	BOOL bSetClipRect = FALSE;

	if (rectClip.IsRectEmpty())
	{
		rectClip = rect;
	}
	else
	{
		bSetClipRect = TRUE;
		m_pGM->SetClipRect(rectClip);
	}

	OnDraw(m_pGM, rectClip);

	if (bSetClipRect)
	{
		m_pGM->ReleaseClipArea();
	}

	SetDirty(FALSE);

	m_pGM->EndDraw();

	m_bInPaint = FALSE;
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnSize(UINT nType, int cx, int cy)
{
	CStatic::OnSize(nType, cx, cy);

	SetRect(CBCGPRect(0, 0, cx, cy), TRUE);

	CRect rect (0, 0, cx, cy);
	if (m_pToolTip->GetSafeHwnd () != NULL)
	{
		m_pToolTip->SetToolRect (this, 1, rect);
	}

	if (GetExStyle() & WS_EX_LAYERED)
	{
		OnDrawLayeredPopup();
	}
}
//***********************************************************************************************************
int CBCGPBaseVisualCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CStatic::OnCreate(lpCreateStruct) == -1)
		return -1;

	if (m_bIsPopup)
	{
		SendMessage(WM_IDLEUPDATECMDUI, (WPARAM)TRUE);
	}
	
	InitCtrl();
	InitTooltip();

	if (GetRect().IsRectEmpty())
	{
		CRect rect;
		GetClientRect(rect);

		CBCGPRAII<BOOL> raii(m_bIgnoreMinSize, TRUE);
		SetRect(rect);
	}

	CBCGPGestureConfig gestureConfig;
	if (GetGestureConfig(gestureConfig))
	{
		bcgpGestureManager.SetGestureConfig(GetSafeHwnd(), gestureConfig);
	}

	OnAfterCreateWnd();
	return 0;
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::PreSubclassWindow() 
{
	CStatic::PreSubclassWindow();

	if (!m_bIsPopup)
	{
		InitCtrl();
		InitTooltip();
	}

	ModifyStyle(0, SS_NOTIFY);

	if (GetRect().IsRectEmpty())
	{
		CRect rect;
		GetClientRect(rect);

		CBCGPRAII<BOOL> raii(m_bIgnoreMinSize, TRUE);
		SetRect(rect);
	}

	CBCGPGestureConfig gestureConfig;
	if (GetGestureConfig(gestureConfig))
	{
		bcgpGestureManager.SetGestureConfig(GetSafeHwnd(), gestureConfig);
	}

	_AFX_THREAD_STATE* pThreadState = AfxGetThreadState ();
	if (pThreadState->m_pWndInit == NULL)
	{
		OnAfterCreateWnd();
	}
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::InitTooltip()
{
	if (!m_bIsTooltip)
	{
		return;
	}

#if !defined _BCGPCHART_STANDALONE
	CBCGPTooltipManager::CreateToolTip (m_pToolTip, this, BCGP_TOOLTIP_TYPE_BUTTON);

	if (m_pToolTip->GetSafeHwnd () != NULL)
	{
		CRect rectClient;
		GetClientRect (&rectClient);

		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectClient, 1);
	}
#endif
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::SetGraphicsManager(CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER manager, CBCGPGraphicsManagerParams* pParams)
{
	if (m_pGM != NULL)
	{
		delete m_pGM;
	}

	m_pGM = CBCGPGraphicsManager::CreateInstance(manager, TRUE, pParams);
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::SetGraphicsManager(CRuntimeClass* pRTI)
{
	if (m_pGM != NULL)
	{
		delete m_pGM;
	}

	m_pGM = DYNAMIC_DOWNCAST(CBCGPGraphicsManager, pRTI->CreateObject());
	ASSERT_VALID(m_pGM);
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::EnableTooltip(BOOL bEnable, BOOL bTooltipTrackingMode)
{
#if !defined _BCGPCHART_STANDALONE
	m_bIsTooltip = bEnable;
	m_bTooltipTrackingMode = bTooltipTrackingMode;

	if (GetSafeHwnd() == NULL)
	{
		return;
	}

	if (m_bIsTooltip)
	{
		m_bUsePopupWindowInTrackingMode = FALSE;
		DeleteInfoTip();

		if (m_pToolTip->GetSafeHwnd() == NULL)
		{
			InitTooltip();

			if (m_pToolTip->GetSafeHwnd() != NULL)
			{
				m_pToolTip->Activate(TRUE);
			}
		}
	}
	else
	{
		CBCGPTooltipManager::DeleteToolTip(m_pToolTip);
	}
#endif
}

#ifndef BCGP_EXCLUDE_POPUP_WINDOW
void CBCGPBaseVisualCtrl::EnableInfoTip(BOOL bEnable/* = TRUE*/, CBCGPInfoTipOptions* pOptions/* = NULL*/)
{
	if (!bEnable)
	{
		DeleteInfoTip();
	}

	m_bUsePopupWindowInTrackingMode = bEnable;

	if (bEnable && pOptions != NULL)
	{
		m_InfoTipOptions = *pOptions;
	}
}
#endif

BOOL CBCGPBaseVisualCtrl::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	CPoint ptCursor;

	::GetCursorPos(&ptCursor);
	ScreenToClient(&ptCursor);

	if (OnSetMouseCursor(ptCursor))
	{
		return TRUE;
	}
	
	return CStatic::OnSetCursor(pWnd, nHitTest, message);
}
//***********************************************************************************************************
BOOL CBCGPBaseVisualCtrl::PreTranslateMessage(MSG* pMsg) 
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
		DeleteInfoTip();
		// Next case....

	case WM_MOUSEMOVE:

		if (m_pToolTip->GetSafeHwnd () != NULL)
		{
			m_pToolTip->RelayEvent(pMsg);
		}
	}

	switch (pMsg->message)
	{
	case WM_LBUTTONDOWN:
	case WM_RBUTTONDOWN:
	case WM_MBUTTONDOWN:
		if (pMsg->hwnd == GetSafeHwnd())
		{
			if (!m_bIsPopup)
			{
				SetFocus();
			}

			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));
			if (OnMouseDown(pMsg->message == WM_LBUTTONDOWN ? 0 : pMsg->message == WM_RBUTTONDOWN ? 1 : 2, pt))
			{
#if !defined _BCGPCHART_STANDALONE
				if (CBCGPPopupMenu::GetActiveMenu() != NULL && ::IsWindow (CBCGPPopupMenu::GetActiveMenu()->m_hWnd))
				{
					BOOL bIsParentMenu = FALSE;
					
					for (CWnd* pWndParent = GetParent(); pWndParent != NULL; pWndParent = pWndParent->GetParent())
					{
						if (pWndParent->GetSafeHwnd() == CBCGPPopupMenu::GetActiveMenu()->m_hWnd)
						{
							bIsParentMenu = TRUE;
							break;
						}
					}

					if (!bIsParentMenu)
					{
						CBCGPPopupMenu::GetActiveMenu()->SendMessage (WM_CLOSE);
					}
				}
#endif
				SetCapture();
				m_bIsCaptured = TRUE;

				StartClickAndHold();
				return TRUE;
			}
		}
		break;

	case WM_LBUTTONUP:
	case WM_RBUTTONUP:
	case WM_MBUTTONUP:
		if (pMsg->hwnd == GetSafeHwnd())
		{
			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));
			OnMouseUp(pMsg->message == WM_LBUTTONUP ? 0 : pMsg->message == WM_RBUTTONUP ? 1 : 2, pt);

			if (m_bIsCaptured)
			{
				ReleaseCapture();
				m_bIsCaptured = FALSE;
			}

			StopClickAndHold();

			if (pMsg->message == WM_LBUTTONUP)
			{
				return TRUE;
			}
		}
		break;

	case WM_LBUTTONDBLCLK:
	case WM_RBUTTONDBLCLK:
	case WM_MBUTTONDBLCLK:
		if (pMsg->hwnd == GetSafeHwnd())
		{
			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));
			OnMouseDblClick(pMsg->message == WM_LBUTTONDBLCLK ? 0 : pMsg->message == WM_RBUTTONDBLCLK ? 1 : 2, pt);
		}
		break;
	
	case WM_MOUSEWHEEL:
		if (CBCGPPopupMenu::GetActiveMenu() == NULL)
		{
			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));

			CWnd* pWnd = CWnd::FromHandle (pMsg->hwnd);

			if (pWnd->GetSafeHwnd() == GetSafeHwnd())
			{
				if (pWnd != NULL)
				{
					pWnd->ScreenToClient(&pt);
				}

				if (OnMouseWheel(pt, HIWORD(pMsg->wParam)))
				{
					if (m_pToolTip->GetSafeHwnd () != NULL)
					{
						m_pToolTip->Pop();
					}

					DeleteInfoTip();
					return TRUE;
				}
			}
		}
		break;

	case WM_NCMOUSEMOVE:
		{
			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));
			OnNcMouseMove(pt);
		}
		break;

	case WM_MOUSEMOVE:
		{
			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));

			OnMouseMove(pt);

			if (!m_bIsTracked)
			{
				m_bIsTracked = TRUE;
				
				TRACKMOUSEEVENT trackmouseevent;
				trackmouseevent.cbSize = sizeof(trackmouseevent);
				trackmouseevent.dwFlags = TME_LEAVE;
				trackmouseevent.hwndTrack = GetSafeHwnd();
				trackmouseevent.dwHoverTime = HOVER_DEFAULT;
				::BCGPTrackMouse (&trackmouseevent);	
			}

			if (m_pToolTip->GetSafeHwnd () != NULL || m_bUsePopupWindowInTrackingMode)
			{
				const int nDelta = 20;

				CString strToolTip;
				CString strDescr;

				OnGetToolTip(pt, strToolTip, strDescr);

				if (strToolTip != m_strLastDisplayedToolTip || strDescr != m_strLastDisplayedToolTipDescr)
				{
#ifndef BCGP_EXCLUDE_POPUP_WINDOW
					if (m_bUsePopupWindowInTrackingMode)
					{
						if (!strToolTip.IsEmpty())
						{
							CPoint ptScreen = pt;
							ClientToScreen(&ptScreen);
							ptScreen.x += nDelta;
							ptScreen.y += nDelta;
							
							CString strText = strToolTip;
							
							if (!strDescr.IsEmpty())
							{
								strText += _T("\r\n\r\n");
								strText += strDescr;
							}

							if (m_pInfoTip == NULL)
							{
								m_pInfoTip = new CBCGPPopupWindow;
								
								m_pInfoTip->SetRoundedCorners(m_InfoTipOptions.m_bRoundedCorners);
								m_pInfoTip->SetStemLocation(m_InfoTipOptions.m_StemLocation, m_InfoTipOptions.m_nStemSize, nDelta);
								m_pInfoTip->SetShadow(m_InfoTipOptions.m_bShadow);
								m_pInfoTip->SetTransparency(m_InfoTipOptions.m_nTransparency);

								m_pInfoTip->SetAnimationType(CBCGPPopupMenu::NO_ANIMATION);
								m_pInfoTip->SetSmallCaptionGripper(FALSE);
								m_pInfoTip->SetAutoCloseTime(0);
								m_pInfoTip->EnableCloseButton(FALSE);

								CBCGPPopupWindowColors colors;
								colors.clrFill = m_InfoTipOptions.m_clrFill;
								colors.clrText = m_InfoTipOptions.m_clrText;
								colors.clrBorder = m_InfoTipOptions.m_clrBorder;

								if (colors.clrFill == (COLORREF)-1)
								{
									colors.clrFill = GetInfoTipColor(pt, BCGP_INFOTIP_FILL_COLOR);
								}

								if (colors.clrText == (COLORREF)-1)
								{
									colors.clrText = GetInfoTipColor(pt, BCGP_INFOTIP_TEXT_COLOR);
								}

								if (colors.clrBorder == (COLORREF)-1)
								{
									colors.clrBorder = GetInfoTipColor(pt, BCGP_INFOTIP_BORDER_COLOR);

									if (colors.clrBorder != (COLORREF)-1 && m_InfoTipOptions.m_StemLocation != CBCGPPopupWindow::BCGPPopupWindowStemLocation_None)
									{
										if (CBCGPDrawManager::IsDarkColor(colors.clrBorder))
										{
											colors.clrBorder = CBCGPDrawManager::ColorMakeLighter(colors.clrBorder, .4);
										}
										else
										{
											colors.clrBorder = CBCGPDrawManager::ColorMakeDarker(colors.clrBorder, .4);
										}
									}
								}

								m_pInfoTip->SetCustomTheme(colors);

								CBCGPPopupWndParams params;

								params.m_strText = strText;
								params.m_nXPadding = m_InfoTipOptions.m_szPadding.cx;
								params.m_nYPadding = m_InfoTipOptions.m_szPadding.cy;

								m_pInfoTip->Create(this, params, NULL, ptScreen);
							}
							else
							{
								m_pInfoTip->UpdateContent(strText, ptScreen);
							}
						}
						else
						{
							DeleteInfoTip();
						}

						m_strLastDisplayedToolTip = strToolTip;
						m_strLastDisplayedToolTipDescr = strDescr;
					}
					else
#endif
					{
						if (m_bTooltipTrackingMode)
						{
							m_pToolTip->Update();
						}
						else
						{
							m_pToolTip->Pop ();
						}

						if (!strToolTip.IsEmpty() && !strDescr.IsEmpty())
						{
							m_bToolTipCleared = TRUE;
						}
					}
				}
#ifndef BCGP_EXCLUDE_POPUP_WINDOW
				else if (m_pInfoTip->GetSafeHwnd() != NULL)
				{
					CPoint ptScreen = pt;
					ClientToScreen(&ptScreen);

					if (m_pInfoTip->GetStemLocation() != CBCGPPopupWindow::BCGPPopupWindowStemLocation_None)
					{
						m_pInfoTip->AdjustLocationByStem(ptScreen, nDelta);
					}
					else
					{
						ptScreen.x += nDelta;
						ptScreen.y += nDelta;
					}

					CRect rectScreen = globalData.m_rectVirtual;

					CRect rectInfoTip;
					m_pInfoTip->GetWindowRect(rectInfoTip);

					if (ptScreen.x < rectScreen.left)
					{
						ptScreen.x = rectScreen.left;
					}
					else if (ptScreen.x + rectInfoTip.Width() > rectScreen.right)
					{
						ptScreen.x = rectScreen.right - rectInfoTip.Width();
					}
					
					if (ptScreen.y < rectScreen.top)
					{
						ptScreen.y = rectScreen.top;
					}
					else if (ptScreen.y + rectInfoTip.Height() > rectScreen.bottom)
					{
						ptScreen.y = rectScreen.bottom - rectInfoTip.Height();
					}

					m_pInfoTip->SetWindowPos(NULL, ptScreen.x, ptScreen.y, -1, -1, 
						SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSIZE);
				}
#endif
			}
		}
		break;

	case WM_CANCELMODE:
		if (m_bIsCaptured)
		{
			ReleaseCapture();
			m_bIsCaptured = FALSE;

			OnCancelMode();
		}

		StopClickAndHold();
		DeleteInfoTip();
		break;

	case WM_KEYDOWN:
		if (pMsg->wParam == VK_ESCAPE || (pMsg->wParam == VK_RETURN && IsReturnKeySupported()))
		{
			if (OnKeyboardDown((UINT)pMsg->wParam, (pMsg->lParam & 0xFFFF), 0))
			{
				return TRUE;
			}
		}
		break;

	case WM_TIMER:
		if (pMsg->hwnd == GetSafeHwnd() && pMsg->wParam == GetClickAndHoldID())
		{
			CPoint point;
			
			::GetCursorPos (&point);
			ScreenToClient (&point);

			if (GetClickAndHoldRect().PtInRect(point))
			{
				OnClickAndHoldEvent(GetClickAndHoldID(), point);
				return TRUE;
			}
		}
		break;
	}

	return CStatic::PreTranslateMessage(pMsg);
}
//********************************************************************************
LRESULT CALLBACK CBCGPBaseVisualCtrl::MouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	if (nCode == HC_ACTION && m_hwndHookedPopup != NULL)
	{
		MOUSEHOOKSTRUCT* lpMS = (MOUSEHOOKSTRUCT*)lParam;
		ASSERT (lpMS != NULL);

		switch (wParam)
		{
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
		case WM_MBUTTONDOWN:
		case WM_NCLBUTTONDOWN:
		case WM_NCRBUTTONDOWN:
		case WM_NCMBUTTONDOWN:
			{
				CBCGPBaseVisualCtrl* pWnd = DYNAMIC_DOWNCAST(CBCGPBaseVisualCtrl, CWnd::FromHandle(m_hwndHookedPopup));
				if (pWnd != NULL)
				{
					ASSERT_VALID(pWnd);
					
					if (pWnd->IsCloseOnClick(lpMS->pt))
					{
						pWnd->PostMessage(WM_CLOSE);
					}
				}
			}
			break;
		}
	}

	return CallNextHookEx(m_hookMouse, nCode, wParam, lParam);
}
//********************************************************************************
BOOL CBCGPBaseVisualCtrl::IsCloseOnClick(CPoint ptScreen)
{
	CRect rectWindow;
	GetWindowRect(rectWindow);

	return !rectWindow.PtInRect(ptScreen);
}
//********************************************************************************
void CBCGPBaseVisualCtrl::StartClickAndHold()
{
	if (!GetClickAndHoldRect().IsRectEmpty())
	{
		SetTimer(GetClickAndHoldID(), 200, NULL);
	}
}
//********************************************************************************
void CBCGPBaseVisualCtrl::StopClickAndHold()
{
	if (!GetClickAndHoldRect().IsRectEmpty())
	{
		KillTimer(GetClickAndHoldID());
	}
}
//********************************************************************************
void CBCGPBaseVisualCtrl::OnEnable(BOOL bEnable) 
{
	OnWndEnabled(bEnable);
	RedrawWindow();
}
//***********************************************************************************************************
LRESULT CBCGPBaseVisualCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam) 
{
	if (message == WM_MOUSELEAVE)
	{
		m_bIsTracked = FALSE;

		DeleteInfoTip();
		OnMouseLeave();
	}
	else if (message == WM_ACTIVATEAPP && m_bIsPopup && wParam == 0)
	{
		if ((m_nPopupFlags & BCGP_VISUAL_POPUP_DONT_CLOSE_ON_MOUSE_CLICK) == 0)
		{
			PostMessage(WM_CLOSE);
		}
	}
	
	return CStatic::WindowProc(message, wParam, lParam);
}
//***********************************************************************************************************
BOOL CBCGPBaseVisualCtrl::OnTTNeedTipText (UINT /*id*/, NMHDR* pNMH, LRESULT* /*pResult*/)
{
	static CString strTipText;

	CPoint point;
	
	::GetCursorPos (&point);
	ScreenToClient (&point);

	CString strDescr;
	if (!OnGetToolTip(point, strTipText, strDescr))
	{
		return FALSE;
	}

	LPNMTTDISPINFO	pTTDispInfo	= (LPNMTTDISPINFO) pNMH;
	ASSERT((pTTDispInfo->uFlags & TTF_IDISHWND) == 0);

	if (m_bToolTipCleared && strTipText == m_strLastDisplayedToolTip && strDescr == m_strLastDisplayedToolTipDescr)
	{
		return FALSE;
	}

#if !defined _BCGPCHART_STANDALONE
	if (!strDescr.IsEmpty())
	{
		CBCGPToolTipCtrl* pToolTip = DYNAMIC_DOWNCAST (
			CBCGPToolTipCtrl, m_pToolTip);

		if (pToolTip != NULL)
		{
			ASSERT_VALID (pToolTip);
			pToolTip->SetDescription(strDescr);
		}
	}
#endif

	pTTDispInfo->lpszText = const_cast<LPTSTR> ((LPCTSTR) strTipText);
	m_bToolTipCleared = FALSE;
	m_strLastDisplayedToolTip = strTipText;
	m_strLastDisplayedToolTipDescr = strDescr;

	return TRUE;
}
//***********************************************************************************************************
LRESULT CBCGPBaseVisualCtrl::OnBCGUpdateToolTips (WPARAM wp, LPARAM)
{
#if !defined _BCGPCHART_STANDALONE
	UINT nTypes = (UINT) wp;

	if (nTypes & BCGP_TOOLTIP_TYPE_DEFAULT)
	{
		CBCGPTooltipManager::CreateToolTip (m_pToolTip, this, BCGP_TOOLTIP_TYPE_BUTTON);

		if (m_pToolTip != NULL)
		{
			CRect rectClient;
			GetClientRect (&rectClient);

			m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectClient, 1);
		}
	}
#endif
	return 0;
}
//*************************************************************************************
void CBCGPBaseVisualCtrl::OnDestroy() 
{
	OnBeforeDestroyWnd();

	if (m_hwndHookedPopup == GetSafeHwnd() && m_hookMouse != NULL)
	{
		::UnhookWindowsHookEx(m_hookMouse);

		m_hookMouse = NULL;
		m_hwndHookedPopup = NULL;
	}

#if !defined _BCGPCHART_STANDALONE
	CBCGPTooltipManager::DeleteToolTip (m_pToolTip);
#endif
	
	DeleteInfoTip();

	CStatic::OnDestroy();
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::DeleteInfoTip()
{
#ifndef BCGP_EXCLUDE_POPUP_WINDOW
	if (m_pInfoTip != NULL)
	{
		if (m_pInfoTip->GetSafeHwnd() != NULL)
		{
			m_pInfoTip->DestroyWindow();
		}
		
		m_pInfoTip = NULL;
	}
#endif
}
//***********************************************************************************************************
UINT CBCGPBaseVisualCtrl::OnGetDlgCode() 
{
	return m_nDlgCode == 0 ? CStatic::OnGetDlgCode() : m_nDlgCode;
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnKillFocus(CWnd* pNewWnd) 
{
	CStatic::OnKillFocus(pNewWnd);
	
	m_bIsFocused = FALSE;
	RedrawWindow ();
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnSetFocus(CWnd* pOldWnd) 
{
	CStatic::OnSetFocus(pOldWnd);
	
	m_bIsFocused = TRUE;
	RedrawWindow ();
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (!OnKeyboardDown(nChar, nRepCnt, nFlags))
	{
		CStatic::OnKeyDown(nChar, nRepCnt, nFlags);
	}
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (!OnKeyboardUp(nChar, nRepCnt, nFlags))
	{
		CStatic::OnKeyUp(nChar, nRepCnt, nFlags);
	}
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnNcDestroy() 
{
	CStatic::OnNcDestroy();

	if (m_bIsPopup)
	{
		delete this;
	}
}
//***********************************************************************************************************
BOOL CBCGPBaseVisualCtrl::OnNcCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (!CStatic::OnNcCreate(lpCreateStruct))
		return FALSE;
	
	if (m_bIsPopup && (GetStyle() & MFS_SYNCACTIVE))
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
LRESULT CBCGPBaseVisualCtrl::OnFloatStatus(WPARAM wParam, LPARAM)
{
	if (!m_bIsPopup)
	{
		return Default();
	}

	// these asserts make sure no conflicting actions are requested
	ASSERT(!((wParam & FS_SHOW) && (wParam & FS_HIDE)));
	ASSERT(!((wParam & FS_ENABLE) && (wParam & FS_DISABLE)));
	ASSERT(!((wParam & FS_ACTIVATE) && (wParam & FS_DEACTIVATE)));

	// FS_SYNCACTIVE is used to detect MFS_SYNCACTIVE windows
	LRESULT lResult = 0;

	if ((GetStyle() & MFS_SYNCACTIVE) && (wParam & FS_SYNCACTIVE))
	{
		lResult = 1;
	}

	if (wParam & (FS_SHOW|FS_HIDE))
	{
		SetWindowPos(NULL, 0, 0, 0, 0,
			((wParam & FS_SHOW) ? SWP_SHOWWINDOW : SWP_HIDEWINDOW) | SWP_NOZORDER |
			SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
	}

	if (wParam & (FS_ENABLE|FS_DISABLE))
	{
		EnableWindow((wParam & FS_ENABLE) != 0);
	}

	if ((wParam & (FS_ACTIVATE|FS_DEACTIVATE)) && GetStyle() & MFS_SYNCACTIVE)
	{
		ModifyStyle(MFS_SYNCACTIVE, 0);
		SendMessage(WM_NCACTIVATE, (wParam & FS_ACTIVATE) != 0);
		ModifyStyle(0, MFS_SYNCACTIVE);
	}

	return lResult;
}
//***********************************************************************************************************
LRESULT CBCGPBaseVisualCtrl::OnBCGPostRedraw(WPARAM wp, LPARAM)
{
	BOOL bSetDirty = (BOOL)wp;

	if (bSetDirty)
	{
		SetDirty(bSetDirty);
	}

	RedrawWindow();
	return 0;
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnDrawLayeredPopup()
{
	CRect rect;
	GetClientRect (rect);

	if (rect.Width () == 0 || rect.Height () == 0)
	{
		return;
	}

	CPoint point (0, 0);
	CSize size (rect.Size ());

	LPBYTE pBits = NULL;
	HBITMAP hBitmap = CBCGPDrawManager::CreateBitmap_32 (size, (void**)&pBits);
	if (hBitmap == NULL)
	{
		return;
	}

	CBitmap bitmap;
	bitmap.Attach (hBitmap);

	CClientDC clientDC(this);
	CDC dc;
	dc.CreateCompatibleDC (&clientDC);

	CBitmap* pBitmapOld = (CBitmap*)dc.SelectObject (&bitmap);

	if (m_pGM == NULL)
	{
		CBCGPGraphicsManagerParams params;
		params.bAlphaModePremultiplied = TRUE;

		m_pGM = CBCGPGraphicsManager::CreateInstance(CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER_DEFAULT, TRUE, &params);
	
		if (m_pGM == NULL)
		{
			return;
		}

		m_pGM->EnableTransparentGradient();
	}

	SetRect(rect);

	m_pGM->BindDC(&dc, rect);

	if (!m_pGM->BeginDraw())
	{
		return;
	}

	OnDraw(m_pGM, rect);

	SetDirty(FALSE);

	m_pGM->EndDraw();

	m_pGM->BindDC(NULL);

	BLENDFUNCTION bf;
	bf.BlendOp             = AC_SRC_OVER;
	bf.BlendFlags          = 0;
	bf.SourceConstantAlpha = (BYTE)bcg_clamp(m_nPopupAlphaAnimated == (BYTE)-1 ? m_nPopupAlpha : m_nPopupAlphaAnimated, 0, 255);
	bf.AlphaFormat         = LWA_COLORKEY;

#ifndef _BCGSUITE_
	globalData.UpdateLayeredWindow (GetSafeHwnd (), NULL, 0, &size, dc.GetSafeHdc (), 
		&point, 0, &bf, 0x02);
#else
	UpdateLayeredWindow (NULL, 0, &size, &dc, &point, 0, &bf, 0x02);
#endif

	dc.SelectObject (pBitmapOld);
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::SetScale(double dblScale)
{
	m_dblScale = dblScale;
	SetDirty(TRUE);
}
//***********************************************************************************************************
BOOL CBCGPBaseVisualCtrl::DoPrint(CDC* pDCIn, CPrintInfo* pInfo)
{
	CObject* pBaseVisual = GetVisualBaseObject();
	if (pBaseVisual == NULL)
	{
		return FALSE;
	}

	CBCGPBaseVisualObject* pVisualObject = DYNAMIC_DOWNCAST(CBCGPBaseVisualObject, pBaseVisual);
	CBCGPVisualContainer* pVisualContainer = DYNAMIC_DOWNCAST(CBCGPVisualContainer, pBaseVisual);

	if (pVisualObject == NULL && pVisualContainer == NULL)
	{
		return FALSE;
	}

	CBCGPGraphicsManager* pGM = GetGraphicsManager();
	if (pGM == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID(pGM);

	CDC* pDC = pDCIn;
	CPrintDialog* pPrintDlg = NULL;

	if (pDC == NULL)
	{
		pPrintDlg = new CPrintDialog(FALSE);
		
		if (pPrintDlg->DoModal() == IDCANCEL)
		{
			delete pPrintDlg;
			return FALSE;
		}

		pDC = CDC::FromHandle(pPrintDlg->GetPrinterDC());
		pDC->m_bPrinting = TRUE;
		
		CString strTitle;
		strTitle.LoadString(AFX_IDS_APP_TITLE);
		
		DOCINFO di;
		ZeroMemory(&di, sizeof (DOCINFO));
		di.cbSize = sizeof (DOCINFO);
		di.lpszDocName = strTitle;
		
		if (!pDC->StartDoc(&di))
		{	
			pDC->AbortDoc();
			delete pPrintDlg;
			return FALSE;
		}

		pInfo = new CPrintInfo();
		pInfo->m_rectDraw.SetRect(0,0, pDC->GetDeviceCaps(HORZRES), pDC->GetDeviceCaps(VERTRES));
	}

	CRuntimeClass* pRTIGM = NULL;

	if (m_PrintOptions.m_bUseGDI && DYNAMIC_DOWNCAST(CBCGPGraphicsManagerGDI, pGM) == NULL)
	{
		pRTIGM = pGM->GetRuntimeClass();

		SetGraphicsManager(CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER_GDI);
		pGM = m_pGM;
	}

	if (pVisualContainer != NULL)
	{
		pVisualContainer->m_bIsPrinting = TRUE;
	}
	else if (pVisualObject != NULL)
	{
		pVisualObject->m_bIsPrinting = TRUE;
	}

	CSize size(pInfo->m_rectDraw.Size());
	CBCGPRect rect(CBCGPPoint(), size);
	
	CBCGPRect rectControl;
	if (pVisualContainer != NULL)
	{
		CBCGPSize sizeContainer = pVisualContainer->GetRect().Size();
		rectControl = CBCGPRect(0., 0., max(pVisualContainer->m_sizeScrollTotal.cx + 10, sizeContainer.cx), max(pVisualContainer->m_sizeScrollTotal.cy + 10, sizeContainer.cy));
	}
	else if (pVisualObject != NULL)
	{
		rectControl = pVisualObject->GetRect();
	}

	CBCGPSize sizeScaleRatio((double)size.cx / rectControl.Width(), (double)size.cy / rectControl.Height());

	double dAspect = min(sizeScaleRatio.cx, sizeScaleRatio.cy);
	sizeScaleRatio.cx = sizeScaleRatio.cy = dAspect;

	size.cx = bcg_clamp((int)(rectControl.Width() * dAspect), 0, size.cx);
	size.cy = bcg_clamp((int)(rectControl.Height() * dAspect), 0, size.cy);
	
	if (m_PrintOptions.m_bPrintScreenshot)
	{
		BOOL bRes = FALSE;

		CDC dcMem;
		if (dcMem.CreateCompatibleDC(pDC))
		{		
			HBITMAP hbmp = pVisualObject != NULL ? pVisualObject->ExportToBitmap(pGM) : pVisualContainer->ExportToBitmap(pGM, TRUE);
			if (hbmp != NULL)
			{
				BITMAP bmp;
				::GetObject(hbmp, sizeof(BITMAP), (LPVOID)&bmp);
				
				const int cx = bmp.bmWidth;
				const int cy = bmp.bmHeight;

				if (cx > 0 && cy > 0)
				{
					HBITMAP hbmpOld = (HBITMAP)dcMem.SelectObject(hbmp);
					if (hbmpOld != NULL)
					{
						CBCGPSize sizeScaleRatio1((double)size.cx / cx, (double)size.cy / cy);
						dAspect = min(sizeScaleRatio1.cx, sizeScaleRatio1.cy);
						
						if (dAspect > 0.0)
						{
							const int cxDest = bcg_clamp((int)(cx * dAspect), 0, size.cx);
							const int cyDest = bcg_clamp((int)(cy * dAspect), 0, size.cy);
							
							int nModeOld = pDC->SetStretchBltMode(HALFTONE);
							pDC->StretchBlt(pInfo->m_rectDraw.left, pInfo->m_rectDraw.top, cxDest, cyDest, &dcMem, 0, 0, cx, cy, SRCCOPY);
							pDC->SetStretchBltMode(nModeOld);
							
							bRes = TRUE;
						}

						dcMem.SelectObject(hbmpOld);
					}
				}
			}
		}

		if (pVisualContainer != NULL)
		{
			pVisualContainer->m_bIsPrinting = FALSE;
		}
		else if (pVisualObject != NULL)
		{
			pVisualObject->m_bIsPrinting = FALSE;
		}

		if (pRTIGM != NULL)
		{
			SetGraphicsManager(pRTIGM);
		}

		return bRes;
	}
	
	rect.SetSize(size);

	CDC	dcPrint;
	CBitmap	bmpPrint;
	CBitmap* pBmpPrintOld = NULL;
	CDC* pDCOut = NULL;

	CBrush brWhite(RGB(255, 255, 255));

	if (DYNAMIC_DOWNCAST(CBCGPGraphicsManagerGDI, pGM) != NULL ||
		DYNAMIC_DOWNCAST(CBCGPGraphicsManagerGdiPlus, pGM) != NULL)
	{
		dcPrint.CreateCompatibleDC(pDC);
		bmpPrint.Attach(CBCGPDrawManager::CreateBitmap_32(size, NULL));
		
		pBmpPrintOld = dcPrint.SelectObject (&bmpPrint);
		ASSERT(pBmpPrintOld != NULL);

		pDCOut = &dcPrint;
	}
	else
	{
		pGM->SetPrintInfo(pInfo);
		pDCOut = pDC;
	}

	pDCOut->FillRect((CRect)rect, &brWhite);
	pGM->BindDC(pDCOut);

	if (pGM->IsBindDCFailed())
	{
		pDCOut->FillRect((CRect)rect, &brWhite);
		pDCOut->TextOut(0, 0, _T("Out of memory"));
	}
	else
	{
		pGM->BeginDraw();

		CBCGPSize sizeScaleRatioOld;
		
		if (pVisualObject != NULL)
		{
			sizeScaleRatioOld = pVisualObject->GetScaleRatio();
			pVisualObject->SetScaleRatio(sizeScaleRatio);
			pVisualObject->SetRect(rect);
		}
		else
		{
			sizeScaleRatioOld = pVisualContainer->GetScaleRatio();
			pVisualContainer->SetScaleRatio(sizeScaleRatio);
			pVisualContainer->SetRect(rect);
		}

		pGM->Clear();

		if (pVisualObject != NULL)
		{
			pVisualObject->OnDraw (pGM, rect);
		}
		else
		{
			pVisualContainer->OnDraw (pGM, rect);

			for (int i = 0; i < pVisualContainer->m_arObjects.GetSize(); i++)
			{
				CBCGPWndHostVisualObject* pObject = DYNAMIC_DOWNCAST(CBCGPWndHostVisualObject, pVisualContainer->m_arObjects[i]);
				if (pObject == NULL)
				{
					continue;
				}
				
				ASSERT_VALID(pObject);
				
				if (pObject->GetWnd()->GetSafeHwnd() != NULL)
				{
					CBitmap bmpScreenshot;
					if (globalUtils.CreateScreenshot(bmpScreenshot, pObject->GetWnd()))
					{
						CBCGPImage image((HBITMAP)bmpScreenshot.GetSafeHandle(), TRUE);
						
						CBCGPSize sizeDest = pObject->GetRect().Size();

						BITMAP bmp;
						bmpScreenshot.GetObject(sizeof(BITMAP), (LPVOID)&bmp);
						
						CBCGPRect rectSrc(0, 0, bmp.bmWidth, bmp.bmHeight);
						pGM->DrawImage(image, pObject->GetRect().TopLeft(), sizeDest, 1.0, CBCGPImage::BCGP_IMAGE_INTERPOLATION_MODE_LINEAR, rectSrc);
					}
				}
			}
		}

		pGM->EndDraw();

		pGM->BindDC(NULL);

		if (pVisualObject != NULL)
		{
			pVisualObject->SetScaleRatio(sizeScaleRatioOld);
		}
		else
		{
			pVisualContainer->SetScaleRatio(sizeScaleRatioOld);
		}

		if (dcPrint.GetSafeHdc() != NULL)
		{
			CBCGPDrawManager::FillAlpha (rect, (HBITMAP)bmpPrint.GetSafeHandle (), 255);
		}
	}

	if (pVisualObject != NULL)
	{
		pVisualObject->SetRect(rectControl);
		pVisualObject->SetDirty ();
	}
	else
	{
		pVisualContainer->SetRect(rectControl);
		pVisualContainer->SetDirty ();
	}

	if (dcPrint.GetSafeHdc() != NULL)
	{
		BITMAP bmp = {0};
		bmpPrint.GetBitmap (&bmp);

		CRect rectOut(pInfo->m_rectDraw.TopLeft (), CSize(bmp.bmWidth, bmp.bmHeight));

		int mode = pDC->SetStretchBltMode(HALFTONE);

		pDC->BitBlt (rectOut.left, rectOut.top, rectOut.Width(), rectOut.Height(), &dcPrint, 0, 0, SRCCOPY);
		pDC->SetStretchBltMode(mode);

		dcPrint.SelectObject(pBmpPrintOld);
	}
	else
	{
		pGM->SetPrintInfo(NULL);
	}

	if (pPrintDlg != NULL)
	{
		pDC->EndDoc();
		delete pPrintDlg;
		delete pInfo;
	}

	if (pRTIGM != NULL)
	{
		SetGraphicsManager(pRTIGM);
	}

	if (pVisualContainer != NULL)
	{
		pVisualContainer->m_bIsPrinting = FALSE;
		SetWindowPos (NULL, -1, -1, -1, -1,  SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
	}
	else if (pVisualObject != NULL)
	{
		pVisualObject->m_bIsPrinting = FALSE;
		SetWindowPos (NULL, -1, -1, -1, -1,  SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
	}

	return TRUE;
}
//*******************************************************************************************************************
LRESULT CBCGPBaseVisualCtrl::OnGestureEvent(WPARAM wp, LPARAM lp)
{
	if (ProcessGestureEvent(this, wp, lp))
	{
		return Default();
	}
	
	return 0;
}
//******************************************************************************************************************
LRESULT CBCGPBaseVisualCtrl::OnTabletQuerySystemGestureStatus(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	return (LRESULT)TABLET_DISABLE_PRESSANDHOLD;
}
//******************************************************************************************************************
LRESULT CBCGPBaseVisualCtrl::OnPrintClient(WPARAM wp, LPARAM lp)
{
	if (m_bInPaint)
	{
		return Default();
	}

	if ((lp & PRF_CLIENT) == PRF_CLIENT)
	{
		CDC* pDC = CDC::FromHandle((HDC) wp);
		ASSERT_VALID(pDC);

		DoPaint(pDC, CRect(0, 0, 0, 0));
		return 0;
	}

	return Default();
}
//****************************************************************************
LRESULT CBCGPBaseVisualCtrl::OnPrint(WPARAM wp, LPARAM lp)
{
	if ((lp & PRF_NONCLIENT) == PRF_NONCLIENT)
	{
		CDC* pDC = CDC::FromHandle((HDC) wp);
		ASSERT_VALID(pDC);
		
		if (m_pGM == NULL)
		{
			if ((m_pGM = CBCGPGraphicsManager::CreateInstance()) == NULL)
			{
				return Default();
			}
		}
		
		CRect rect;
		GetWindowRect(rect);
		
		CRect rectClient;
		GetClientRect(rectClient);
		ClientToScreen(rectClient);
		
		rectClient.OffsetRect(-rect.TopLeft());
		rect.OffsetRect(-rect.TopLeft());
		
		pDC->ExcludeClipRect(rectClient);

		m_pGM->BindDC(pDC, rect);
		
		if (m_pGM->BeginDraw())
		{
			OnNcDraw(m_pGM, rect);
			m_pGM->EndDraw();
		}

		pDC->SelectClipRgn(NULL);
	}

	return Default();
}
//***********************************************************************************************************
LRESULT CBCGPBaseVisualCtrl::OnGetObject(WPARAM wParam, LPARAM lParam)
{
#ifndef _BCGPCHART_STANDALONE
	if (globalData.IsAccessibilitySupport ())
	{
		
        if ((DWORD)lParam == OBJID_CLIENT)
		{
			CBCGPBaseAccessibleObject* pVisualObject = DYNAMIC_DOWNCAST(CBCGPBaseAccessibleObject, GetVisualBaseObject());
			if (pVisualObject == NULL || !pVisualObject->IsAccessibilityCompatible())
			{
				return 0;
			}
			

			LPUNKNOWN pUnknown = pVisualObject->GetInterface(&IID_IAccessible);
			if (pUnknown == NULL)
			{
				return 0;
			}
			
			pVisualObject->OnSetIAccessible((IAccessible*)pUnknown);
			
			return LresultFromObject(IID_IAccessible, wParam, pUnknown);
		}
	}
#endif	
	return (LRESULT)0L;
}	
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnNcCalcSize(BOOL bCalcValidRects, NCCALCSIZE_PARAMS FAR* lpncsp) 
{
	CBCGPRect rectNCArea;
	OnCalcBorderSize(rectNCArea);
	
	lpncsp->rgrc[0].top += (int)rectNCArea.top;
	lpncsp->rgrc[0].bottom -= (int)rectNCArea.bottom;
	lpncsp->rgrc[0].left += (int)rectNCArea.left;
	lpncsp->rgrc[0].right -= (int)rectNCArea.right;
	
	CStatic::OnNcCalcSize(bCalcValidRects, lpncsp);
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnNcLButtonDown(UINT /*nHitTest*/, CPoint point) 
{
	CRect rect;
	GetWindowRect(rect);

	point.Offset(-rect.TopLeft());

	if (OnNcMouseDown(0, point))
	{
		SetCapture();
		m_bIsCaptured = TRUE;

		StartClickAndHold();
	}
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnNcLButtonUp(UINT /*nHitTest*/, CPoint point)
{
	CRect rect;
	GetWindowRect(rect);
	
	point.Offset(-rect.TopLeft());

	OnNcMouseUp(0, point);

	if (m_bIsCaptured)
	{
		ReleaseCapture();
		m_bIsCaptured = FALSE;
	}

	StopClickAndHold();
}
//***********************************************************************************************************
void CBCGPBaseVisualCtrl::OnNcPaint() 
{
	if (m_pGM == NULL)
	{
		if ((m_pGM = CBCGPGraphicsManager::CreateInstance()) == NULL)
		{
			return;
		}
	}
	
	CWindowDC dc(this);
	
	CRect rect;
	GetWindowRect(rect);

	CRect rectClient;
    GetClientRect(rectClient);
    ClientToScreen(rectClient);
	
    rectClient.OffsetRect(-rect.TopLeft());
	rect.OffsetRect(-rect.TopLeft());

    dc.ExcludeClipRect(rectClient);

	if (DYNAMIC_DOWNCAST(CBCGPGraphicsManagerGDI, m_pGM) != NULL)
	{
		m_pGM->BindDC(&dc, FALSE);
	}
	else
	{
		m_pGM->BindDC(&dc, rect);
	}
	
	if (m_pGM->BeginDraw())
	{
		OnNcDraw(m_pGM, rect);
		m_pGM->EndDraw();
	}
}
//***********************************************************************************************************
BCGNcHitTestType CBCGPBaseVisualCtrl::OnNcHitTest(CPoint point) 
{
	CRect rectClient;
	GetClientRect(rectClient);
	ClientToScreen(rectClient);

	return rectClient.PtInRect(point) ? HTCLIENT : HTBORDER;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPVisualContainerCtrl

BEGIN_MESSAGE_MAP(CBCGPVisualContainerCtrl, CBCGPBaseVisualCtrl)
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)
END_MESSAGE_MAP()

BOOL CBCGPVisualContainerCtrl::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult) 
{
#ifndef BCGP_EXCLUDE_GRID_CTRL

	CBCGPVisualContainer* pVisualContainer = GetVisualContainer();
	if (pVisualContainer != NULL)
	{
		NMHDR* pNMHDR = (NMHDR*)(lParam);
		if (pNMHDR != NULL && pNMHDR->code == BCGPGN_SELCHANGED)
		{
			pVisualContainer->OnGridSelChanged((UINT)pNMHDR->idFrom);

			*pResult = 0L;
			return TRUE;
		}
	}

#endif

	return CBCGPBaseVisualCtrl::OnNotify(wParam, lParam, pResult);
}
//***********************************************************************************************************
LRESULT CBCGPVisualContainerCtrl::OnChangeVisualManager(WPARAM, LPARAM)
{
	CBCGPVisualContainer* pVisualContainer = GetVisualContainer();
	if (pVisualContainer != NULL)
	{
		ASSERT_VALID(pVisualContainer);
		pVisualContainer->OnChangeVisualManager();
	}
	
	return 0;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPVisualCtrl

BEGIN_MESSAGE_MAP(CBCGPVisualCtrl, CBCGPBaseVisualCtrl)
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)
END_MESSAGE_MAP()

LRESULT CBCGPVisualCtrl::OnChangeVisualManager(WPARAM, LPARAM)
{
	CBCGPBaseVisualObject* pVisualObject = GetVisualObject();
	if (pVisualObject != NULL)
	{
		ASSERT_VALID(pVisualObject);
		pVisualObject->OnChangeVisualManager();
	}
	
	return 0;
}
