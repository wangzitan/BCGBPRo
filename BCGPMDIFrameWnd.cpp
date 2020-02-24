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

// BCGPMDIFrameWnd.cpp : implementation file
//

#include "stdafx.h"

#include "BCGPMDIFrameWnd.h"
#include "BCGPMDIChildWnd.h"
#include "BCGPToolbar.h"
#include "BCGPMenuBar.h"
#include "BCGPPopupMenu.h"
#include "BCGPPopupWindow.h"
#include "BCGPToolbarMenuButton.h"
#include "BCGPMiniFrameWnd.h"
#include "BCGPLocalResource.h"
#include "bcgprores.h"
#include "BCGPWindowsManagerDlg.h"
#include "BCGPUserToolsManager.h"
#include "BCGPContextMenuManager.h"
#include "BCGPSlider.h"
#include "BCGPWorkspace.h"
#include "BCGPRibbonBackstageView.h"
#include "BCGPIntelliSenseWnd.h"
#include "BCGPRegistry.h"

#if _MSC_VER >= 1300
	#include <..\atlmfc\src\mfc\oleimpl2.h>
#else
	#include <..\src\oleimpl2.h>
#endif

#include "BCGPDockingControlBar.h"
#include "BCGPTabbedControlBar.h"
#include "BCGPPrintPreviewView.h"
#include "BCGPRibbonBar.h"
#include "BCGPRibbonStatusBar.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


extern CBCGPWorkspace*	g_pWorkspace;
extern CObList			gAllToolbars;
extern UINT				BCGM_ON_CHANGE_BACKSTAGE_PROP_HIGHLIGHTING;

/////////////////////////////////////////////////////////////////////////////
// CBCGPMDIFrameWnd

IMPLEMENT_DYNCREATE(CBCGPMDIFrameWnd, CMDIFrameWnd)

BOOL CBCGPMDIFrameWnd::m_bDisableSetRedraw = TRUE;
HWND CBCGPMDIFrameWnd::m_hwndActiveDetachedMDIFrame = NULL;
HWND CBCGPMDIFrameWnd::m_hwndLastActiveDetachedMDIFrame = NULL;

CMap<UINT, UINT, HWND, HWND> g_mapWindowCommands;

#pragma warning (disable : 4355)

#ifndef MSGFLT_ADD
#define MSGFLT_ADD 1
#endif

#ifndef MSGFLT_REMOVE
#define MSGFLT_REMOVE 2
#endif

UINT BCGM_REACTIVATE_DETACHED_FRAME			= ::RegisterWindowMessage (_T("BCGM_REACTIVATE_DETACHED_FRAME"));
UINT BCGM_RESET_DETACHED_FRAME				= ::RegisterWindowMessage (_T("BCGM_RESET_DETACHED_FRAME"));
UINT BCGM_SYNC_FRAME_OPTIONS				= ::RegisterWindowMessage (_T("BCGM_SYNC_FRAME_OPTIONS"));
UINT BCGM_ON_ACTIVATE_MDI_TEAR_OFF_FRAME	= ::RegisterWindowMessage (_T("BCGM_ON_ACTIVATE_MDI_TEAR_OFF_FRAME"));
UINT BCGM_POST_ACTIVATE_MDI_TEAR_OFF_FRAME	= ::RegisterWindowMessage (_T("BCGM_POST_ACTIVATE_MDI_TEAR_OFF_FRAME"));

CBCGPMDIFrameWnd::CBCGPMDIFrameWnd() :
	m_Impl (this),
	m_hmenuWindow (NULL),
	m_bUseWindowsMenu(FALSE),
	m_bContextHelp (FALSE),
	m_bDoSubclass (TRUE),
	m_uiWindowsDlgMenuId (0),
	m_strMoreWindowsMenuItem(_T("&More Windows...")),
	m_bShowWindowsDlgAlways (FALSE),
	m_bShowWindowsDlgHelpButton (FALSE),
	m_bResizableWindowsDlg(FALSE),
	m_bWasMaximized (FALSE),
	m_bIsMinimized (FALSE),
	m_bClosing (FALSE),
	m_nFrameID (0),
	m_pPrintPreviewFrame (NULL),
	m_bCanCovertControlBarToMDIChild (FALSE),
	m_bAutoPaneActivation(TRUE),
	m_bIsMDIChildDetached(FALSE),
	m_bIsInternalEnable(FALSE),
	m_bActivatedBySysKey(FALSE),
	m_pMDITearOffFrameRTI(NULL)
{
	if (globalData.bIsWindows7)
	{
		globalData.ChangeWindowMessageFilter(WM_DWMSENDICONICTHUMBNAIL, MSGFLT_ADD);
		globalData.ChangeWindowMessageFilter(WM_DWMSENDICONICLIVEPREVIEWBITMAP, MSGFLT_ADD);
	}
}

#pragma warning (default : 4355)

CBCGPMDIFrameWnd::~CBCGPMDIFrameWnd()
{
}

//BEGIN_MESSAGE_MAP(CBCGPMDIFrameWnd, CMDIFrameWnd)
BEGIN_MESSAGE_MAP(CBCGPMDIFrameWnd, CMDIFrameWnd)
	//{{AFX_MSG_MAP(CBCGPMDIFrameWnd)
	ON_WM_MENUCHAR()
	ON_WM_WINDOWPOSCHANGED()
	ON_WM_ACTIVATE()
	ON_WM_CLOSE()
	ON_WM_DESTROY()
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_GETMINMAXINFO()
	ON_WM_CONTEXTMENU()
	ON_WM_NCPAINT()
	ON_WM_NCACTIVATE()
	ON_WM_NCMOUSEMOVE()
	ON_WM_NCHITTEST()
	ON_WM_NCCALCSIZE()
	ON_WM_LBUTTONUP()
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONDOWN()
	ON_WM_WINDOWPOSCHANGING()
	ON_WM_STYLECHANGING()
	ON_WM_SYSCOMMAND()
	ON_WM_ENABLE()
	ON_WM_SETTINGCHANGE()
	ON_WM_NCLBUTTONDOWN()
	//}}AFX_MSG_MAP
	ON_WM_ACTIVATEAPP()
	ON_MESSAGE(WM_IDLEUPDATECMDUI, OnIdleUpdateCmdUI)
	ON_REGISTERED_MESSAGE(BCGM_CREATETOOLBAR, OnToolbarCreateNew)
	ON_REGISTERED_MESSAGE(BCGM_DELETETOOLBAR, OnToolbarDelete)
	ON_COMMAND( ID_CONTEXT_HELP, OnContextHelp)
	ON_MESSAGE(WM_ENTERSIZEMOVE, OnEnterSizeMove)
	ON_MESSAGE(WM_EXITSIZEMOVE, OnExitSizeMove)
	ON_UPDATE_COMMAND_UI(ID_VIEW_STATUS_BAR, OnUpdateControlBarMenu)
	ON_COMMAND_EX(ID_VIEW_STATUS_BAR, OnBarCheck)
	ON_UPDATE_COMMAND_UI(ID_VIEW_TOOLBAR, OnUpdateControlBarMenu)
	ON_COMMAND_EX(ID_VIEW_TOOLBAR, OnBarCheck)
	ON_UPDATE_COMMAND_UI(ID_VIEW_REBAR, OnUpdateControlBarMenu)
	ON_COMMAND_EX(ID_VIEW_REBAR, OnBarCheck)
	ON_COMMAND(ID_WINDOW_NEW, OnWindowNew)
	ON_REGISTERED_MESSAGE(BCGM_TOOLBARMENU, OnToolbarContextMenu)
	ON_MESSAGE(WM_SETTEXT, OnSetText)
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)
	ON_REGISTERED_MESSAGE(BCGM_POSTSETPREVIEWFRAME, OnPostPreviewFrame)
	ON_REGISTERED_MESSAGE(BCGM_ON_AFTER_TASKBAR_ACTIVATE, OnAfterTaskbarActivate)
	ON_MESSAGE(WM_DWMCOMPOSITIONCHANGED, OnDWMCompositionChanged)
	ON_MESSAGE(WM_POWERBROADCAST, OnPowerBroadcast)
	ON_REGISTERED_MESSAGE(BCGM_REACTIVATE_DETACHED_FRAME, OnReactivateDetachedFrame)
	ON_REGISTERED_MESSAGE(BCGM_RESET_DETACHED_FRAME, OnCleanDetachedFrame)
	ON_REGISTERED_MESSAGE(BCGM_SYNC_FRAME_OPTIONS, OnSyncFrameOptions)
	ON_REGISTERED_MESSAGE(BCGM_ON_CHANGE_BACKSTAGE_PROP_HIGHLIGHTING, OnChangeBackstagePropHighlighting)
	ON_REGISTERED_MESSAGE(BCGM_POST_ACTIVATE_MDI_TEAR_OFF_FRAME, OnPostMDITearOffFrame)
	ON_MESSAGE(WM_DPICHANGED, OnDPIChanged)
	ON_MESSAGE(WM_THEMECHANGED, OnThemeChanged)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPMDIFrameWnd message handlers

BOOL CBCGPMDIFrameWnd::OnSetMenu (HMENU hmenu)
{
	HMENU hMenuctiveMDITearOff = GetActiveMDITearOffMenu();
	if (hMenuctiveMDITearOff != NULL)
	{
		hmenu = hMenuctiveMDITearOff;
	}

	COleClientItem*	pActiveItem = GetInPlaceActiveItem ();
	if (pActiveItem != NULL && 
		pActiveItem->GetInPlaceWindow () != NULL)
	{
		return FALSE;
	}

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_Impl.m_pRibbonBar != NULL && 
		((m_Impl.m_pRibbonBar->GetStyle () & WS_VISIBLE) == WS_VISIBLE ||
		m_Impl.IsFullScreen ()))
	{
		SetMenu (NULL);
		m_Impl.m_pRibbonBar->SetActiveMDIChild (MDIGetActive ());
		return TRUE;
	}
#endif

	if (m_Impl.m_pMenuBar != NULL)
	{
		SetMenu (NULL);
		m_Impl.m_pMenuBar->CreateFromMenu (hmenu == NULL ? m_Impl.m_hDefaultMenu : hmenu);
		return TRUE;
	}

	return FALSE;
}
//*******************************************************************************************
static void AddWindowMenuItem(CBCGPPopupMenu& menuPopup, CWnd* pMDIChild, int& nCount, BOOL bIsActive = FALSE)
{
	if (nCount == 0)
	{
		menuPopup.InsertSeparator();
	}

	CString strName;
	pMDIChild->GetWindowText(strName);
	
	CString strItem;
	strItem.Format(_T("&%d %s"), nCount + 1, (LPCTSTR)strName);
	
	const UINT nCmd = AFX_IDM_FIRST_MDICHILD + nCount;
	g_mapWindowCommands[nCmd] = pMDIChild->GetSafeHwnd();
	
	CBCGPToolbarMenuButton button(nCmd, NULL, -1, strItem);
	if (bIsActive)
	{
		button.m_nStyle |= TBBS_CHECKED;
	}
	
	menuPopup.InsertItem(button);
	nCount++;
}
//*******************************************************************************************
void CBCGPMDIFrameWnd::PrepareWindowsMenu(CBCGPPopupMenu& menuPopup)
{
	BOOL bAddWindowsDlg = FALSE;
	const int nMaxItems = 9;

	if (m_bUseWindowsMenu)
	{
		CMenu* pMenu = CMenu::FromHandle (m_hmenuWindow);
		if (pMenu != NULL)
		{
			int iCount = (int) pMenu->GetMenuItemCount ();
			BOOL bIsFirstWindowItem = TRUE;
			
			for (int i = 0; i < iCount; i ++)
			{
				UINT uiCmd = pMenu->GetMenuItemID (i);
				if (uiCmd < AFX_IDM_FIRST_MDICHILD || uiCmd == (UINT) -1)
				{
					continue;
				}
				
				if (m_uiWindowsDlgMenuId != 0 && uiCmd == AFX_IDM_FIRST_MDICHILD + nMaxItems)
				{
					bAddWindowsDlg = TRUE;
					continue;
				}
				
				if (bIsFirstWindowItem)
				{
					menuPopup.InsertSeparator ();
					bIsFirstWindowItem = FALSE;
					
					::SendMessage (m_hWndMDIClient, WM_MDIREFRESHMENU, 0, 0);
				}
				
				CString strText;
				pMenu->GetMenuString (i, strText, MF_BYPOSITION);
				
				CBCGPToolbarMenuButton button (uiCmd, NULL, -1, strText);
				
				UINT uiState = pMenu->GetMenuState (i, MF_BYPOSITION);
				if (uiState & MF_CHECKED)
				{
					button.m_nStyle |= TBBS_CHECKED;
				}
				
				menuPopup.InsertItem (button);
			}
		}
	}
	else
	{
		int nCount = 0;
		
		CBCGPMDIFrameWnd* pActiveMDIFrame = GetActiveTearOffFrame();
		if (pActiveMDIFrame == NULL)
		{
			pActiveMDIFrame = this;
		}
		
		CWnd* pWndActive = pActiveMDIFrame->MDIGetActive();

		if (pWndActive->GetSafeHwnd() != NULL)
		{
			AddWindowMenuItem(menuPopup, pWndActive, nCount, TRUE);
		}

		POSITION pos = NULL;

		const CList<CFrameWnd*, CFrameWnd*>& lstFrames = CBCGPFrameImpl::GetFrameList();

		CList<CBCGPMDIFrameWnd*, CBCGPMDIFrameWnd*> lstMDIFrames;
		lstMDIFrames.AddTail(pActiveMDIFrame);

		for (pos = lstFrames.GetHeadPosition(); pos != NULL;)
		{
			CBCGPMDIFrameWnd* pMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, lstFrames.GetNext(pos));
			if (pMDIFrame->GetSafeHwnd() != NULL && pMDIFrame != pActiveMDIFrame)
			{
				lstMDIFrames.AddTail(pMDIFrame);
			}
		}

		for (pos = lstMDIFrames.GetHeadPosition(); pos != NULL;)
		{
			if (nCount >= nMaxItems)
			{
				break;
			}

			CBCGPMDIFrameWnd* pMDIFrame = lstMDIFrames.GetNext(pos);
			if (pMDIFrame->GetSafeHwnd() != NULL)
			{
				for (HWND hwndMDIChild = ::GetWindow (pMDIFrame->m_hWndMDIClient, GW_CHILD); hwndMDIChild != NULL; 
					hwndMDIChild = ::GetWindow (hwndMDIChild, GW_HWNDNEXT))
				{
					if (nCount >= nMaxItems)
					{
						break;
					}

					CBCGPMDIChildWnd* pMDIChildFrame = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, CWnd::FromHandle (hwndMDIChild));
					if (pMDIChildFrame != NULL && pMDIChildFrame->IsWindowVisible() && pMDIChildFrame->CanShowOnWindowsList())
					{
						if (pMDIChildFrame->GetSafeHwnd() != pWndActive->GetSafeHwnd())
						{
							AddWindowMenuItem(menuPopup, pMDIChildFrame, nCount);
						}
					}
				}
			}
		}

		if (nCount >= nMaxItems)
		{
			if (m_uiWindowsDlgMenuId != 0)
			{
				bAddWindowsDlg = TRUE;
			}
			else
			{
				// Add "More windows..." item:
				menuPopup.InsertItem(CBCGPToolbarMenuButton(AFX_IDM_FIRST_MDICHILD + nMaxItems, NULL, -1, m_strMoreWindowsMenuItem));
			}
		}
	}
	
	if (m_uiWindowsDlgMenuId != 0 && (bAddWindowsDlg || m_bShowWindowsDlgAlways))
	{
		if (!CBCGPToolBar::GetBasicCommands ().IsEmpty())
		{
			CBCGPToolBar::AddBasicCommand(m_uiWindowsDlgMenuId);
		}
		
		//-----------------------------
		// Add our "Windows..." dialog:
		//-----------------------------
		menuPopup.InsertItem(CBCGPToolbarMenuButton(m_uiWindowsDlgMenuId, NULL, -1, m_strWindowsDlgMenuText));
	}
}
//*******************************************************************************************
BOOL CBCGPMDIFrameWnd::OnCreateClient(LPCREATESTRUCT lpcs, CCreateContext* pContext) 
{
	BOOL bRes = m_bIsMDIChildDetached ? CreateClient(lpcs, NULL) : CMDIFrameWnd::OnCreateClient(lpcs, pContext);
	if (!bRes)
	{
		return FALSE;
	}

	if (m_bDoSubclass)
	{
		m_wndClientArea.SubclassWindow (m_hWndMDIClient);
	}

	return TRUE;
}
//*******************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnMenuChar(UINT nChar, UINT nFlags, CMenu* pMenu) 
{
	if (m_Impl.OnMenuChar (nChar))
	{
		return MAKELPARAM (MNC_EXECUTE, -1);
	}
		
	return CMDIFrameWnd::OnMenuChar(nChar, nFlags, pMenu);
}
//*******************************************************************************************
void CBCGPMDIFrameWnd::OnWindowPosChanged(WINDOWPOS FAR* lpwndpos) 
{
	if ((lpwndpos->flags & SWP_FRAMECHANGED) == SWP_FRAMECHANGED)
	{
		m_Impl.OnWindowPosChanging (lpwndpos);
	}
	else if (IsZoomed() && (m_Impl.m_pRibbonBar != NULL || m_Impl.IsOwnerDrawCaption()) &&
		globalData.GetShellAutohideBars () != 0)
	{
		if ((lpwndpos->flags & SWP_NOMOVE) != SWP_NOMOVE)
		{
			ShowWindow(SW_RESTORE);
		}
	}

	CMDIFrameWnd::OnWindowPosChanged(lpwndpos);
	
	if (m_Impl.m_pShadow != NULL)
	{
		m_Impl.m_pShadow->Repos();
	}

	if (m_Impl.m_pMenuBar != NULL)
	{
		BOOL bMaximized;
		CMDIChildWnd* pChild = MDIGetActive (&bMaximized);

		if (pChild == NULL || !bMaximized)
		{
			m_Impl.m_pMenuBar->SetMaximizeMode (FALSE);
		}
		else
		{
			m_Impl.m_pMenuBar->SetMaximizeMode (TRUE, pChild);
		}
	}

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_Impl.m_pRibbonBar != NULL)
	{
		ASSERT_VALID (m_Impl.m_pRibbonBar);

		BOOL bMaximized;
		CMDIChildWnd* pChild = MDIGetActive (&bMaximized);

		if (pChild == NULL || !bMaximized)
		{
			m_Impl.m_pRibbonBar->SetMaximizeMode (FALSE);
		}
		else
		{
			m_Impl.m_pRibbonBar->SetMaximizeMode (TRUE, pChild);
		}
	}
#endif

	if ((lpwndpos->flags & SWP_NOSIZE) == 0)
	{
		m_Impl.AdjustRibbonBackstageViewRect();
	}

	if (m_Impl.m_pLightBoxDlg != NULL)
	{
		if ((lpwndpos->flags & SWP_NOMOVE) == 0 || (lpwndpos->flags & SWP_NOSIZE) == 0)
		{
			m_Impl.m_pLightBoxDlg->DoRepos();
		}
		
		m_Impl.m_pLightBoxDlg->RecalcClosebutton();
	}
}
//*******************************************************************************************
BOOL CBCGPMDIFrameWnd::PreTranslateMessage(MSG* pMsg) 
{
	BOOL bProcessAccel = TRUE;
	int nKey = 0;

	switch (pMsg->message)
	{
	case WM_SYSKEYDOWN:
		if (m_wndClientArea.IsDragMDIChild() || !::IsWindow(pMsg->hwnd))
		{
			return TRUE;
		}

		nKey = (int)pMsg->wParam;

#ifndef BCGP_EXCLUDE_RIBBON
		if (m_Impl.m_pRibbonBar != NULL && m_Impl.m_pRibbonBar->OnSysKeyDown (this, pMsg->wParam, pMsg->lParam))
		{
			return TRUE;
		}

		if (pMsg->wParam == VK_F4 && m_Impl.m_pRibbonBar != NULL && m_Impl.m_pRibbonBar->IsBackstageViewActive())
		{
			break;
		}
#endif
		if (CBCGPPopupMenu::GetSafeActivePopupMenu() == NULL)
		{
			m_Impl.CancelToolbarMode();
		}

	case WM_CONTEXTMENU:
		if (m_wndClientArea.IsDragMDIChild())
		{
			return TRUE;
		}

		if (pMsg->message == WM_CONTEXTMENU)
		{
			nKey = 0;
		}

		if (!globalData.m_bSysUnderlineKeyboardShortcuts && !globalData.m_bUnderlineKeyboardShortcuts)
		{
			globalData.m_bUnderlineKeyboardShortcuts = TRUE;
			CBCGPToolBar::RedrawUnderlines ();
		}

		if (CBCGPPopupMenu::GetSafeActivePopupMenu() != NULL && (nKey == VK_MENU || nKey == VK_F10))
		{
			if (!CBCGPPopupMenu::m_pActivePopupMenu->IsRibbonAutohidePopup())
			{
				CBCGPPopupMenu::m_pActivePopupMenu->SendMessage (WM_CLOSE);
				return TRUE;
			}
		}
		else if (m_Impl.ProcessKeyboard(nKey))
		{
			return TRUE;
		}
		break;

	case WM_SYSKEYUP:
		if (m_wndClientArea.IsDragMDIChild())
		{
			return TRUE;
		}

		m_bActivatedBySysKey = TRUE;

		if (m_Impl.ProcessSysKeyUp(pMsg->wParam, pMsg->lParam))
		{
			m_bActivatedBySysKey = FALSE;
			return TRUE;
		}

		m_bActivatedBySysKey = FALSE;
		break;

	case WM_KEYDOWN:
		if (m_wndClientArea.IsDragMDIChild())
		{
			if (pMsg->wParam == VK_ESCAPE)
			{
				CBCGPTabWnd* pCapturedTab = DYNAMIC_DOWNCAST(CBCGPTabWnd, GetCapture());
				if (pCapturedTab->GetSafeHwnd() != NULL)
				{
					pCapturedTab->SendMessage(WM_CANCELMODE);
				}
			}

			return TRUE;
		}

		//-----------------------------------------
		// Pass keyboard action to the active menu:
		//-----------------------------------------
		if (!CBCGPFrameImpl::IsHelpKey (pMsg) && 
			m_Impl.ProcessKeyboard ((int) pMsg->wParam))
		{
			return TRUE;
		}

		if (pMsg->wParam == VK_ESCAPE)
		{
			if (IsFullScreen())
			{
				if (!IsPrintPreview())
				{
					m_Impl.InvokeFullScreenCommand();
					return TRUE;
				}
			}

			CBCGPSmartDockingManager* pSDManager = NULL;
			if ((pSDManager = m_dockManager.GetSDManagerPermanent()) != NULL &&
				pSDManager->IsStarted())
			{
				pSDManager->CauseCancelMode ();
			}

			CBCGPSlider* pSlider = DYNAMIC_DOWNCAST (CBCGPSlider, GetCapture ());
			if (pSlider != NULL)
			{
				pSlider->SendMessage (WM_CANCELMODE);
				return TRUE;
			}
		}

		if (pMsg->wParam == VK_F4 && (0x8000 & GetKeyState(VK_CONTROL)) != 0)
		{
			CWnd* pWndCapture = GetCapture();
			if (pWndCapture->GetSafeHwnd() != NULL)
			{
				OnCancelWndCapture(pWndCapture);
			}
		}

		if (!bProcessAccel)
		{
			return FALSE;
		}
		break;

	case WM_LBUTTONDOWN:
	case WM_RBUTTONDOWN:
	case WM_RBUTTONUP:
	case WM_MBUTTONDOWN:
	case WM_MBUTTONUP:
		{
			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));
			CWnd* pWnd = CWnd::FromHandle(pMsg->hwnd);

			if (pWnd != NULL && ::IsWindow (pMsg->hwnd))
			{
				pWnd->ClientToScreen (&pt);
			}

			if (m_Impl.ProcessMouseClick (pMsg->message, pt, pMsg->hwnd))
			{
				return TRUE;
			}

			if (!::IsWindow (pMsg->hwnd))
			{
				return TRUE;
			}
		}
		break;

	case WM_NCLBUTTONDOWN:
	case WM_NCLBUTTONUP:
	case WM_NCRBUTTONDOWN:
	case WM_NCRBUTTONUP:
	case WM_NCMBUTTONDOWN:
	case WM_NCMBUTTONUP:
		if (m_Impl.ProcessMouseClick (pMsg->message,
			CPoint (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam)),
			pMsg->hwnd))
		{
			return TRUE;
		}
		break;

	case WM_MOUSEWHEEL:
		if (m_Impl.ProcessMouseWheel (pMsg->wParam, pMsg->lParam))
		{
			return TRUE;
		}
		break;

	case WM_COMMAND:
		if (m_hwndLastActiveDetachedMDIFrame != NULL)
		{
			PostMessage(BCGM_REACTIVATE_DETACHED_FRAME);
		}
		break;

	case WM_MOUSEMOVE:
		{
			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));
			CWnd* pWnd = CWnd::FromHandle(pMsg->hwnd);
			if (pWnd != NULL)
			{
				pWnd->ClientToScreen (&pt);
			}

			if (m_Impl.ProcessMouseMove (pt))
			{
				return TRUE;
			}
		}
	}

	return CMDIFrameWnd::PreTranslateMessage(pMsg);
}
//*******************************************************************************************
BOOL CBCGPMDIFrameWnd::OnCommand(WPARAM wParam, LPARAM lParam) 
{
	UINT uiCmd = LOWORD(wParam);

	if (uiCmd == ID_HELP && m_Impl.ProcessRibbonContextHelp())
	{
		return TRUE;
	}

	HWND hWndMDIChild = NULL;
	if (!g_mapWindowCommands.IsEmpty() && g_mapWindowCommands.Lookup(uiCmd, hWndMDIChild))
	{
		CWnd* pMDIChild = CWnd::FromHandlePermanent(hWndMDIChild);
		if (pMDIChild->GetSafeHwnd() != NULL)
		{
			CBCGPMDIFrameWnd* pMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, pMDIChild->GetParentFrame());
			if (pMDIFrame != NULL)
			{
				if (pMDIFrame->IsIconic())
				{
					pMDIFrame->ShowWindow(SW_RESTORE);
				}
				
				pMDIFrame->SetActiveWindow();
				
				WINDOWPLACEMENT	wndpl;
				wndpl.length = sizeof(WINDOWPLACEMENT);
				
				::GetWindowPlacement(hWndMDIChild, &wndpl);
				
				if (wndpl.showCmd == SW_SHOWMINIMIZED)
				{
					::ShowWindow(hWndMDIChild, SW_RESTORE);
				}
	
				if (pMDIFrame->GetSafeHwnd() == GetSafeHwnd())
				{
					m_hwndLastActiveDetachedMDIFrame = m_hwndActiveDetachedMDIFrame = NULL;
				}

				::SendMessage(pMDIFrame->m_hWndMDIClient, WM_MDIACTIVATE, (WPARAM)hWndMDIChild, 0);
				return TRUE;
			}
		}
	}

	if (HIWORD (wParam) == 1 && lParam == 0)
	{
		CBCGPToolBar::AddCommandUsage (uiCmd);

		if (IsWindowsNavigatorEnabled() && (uiCmd == m_Impl.m_nWindowsNextNavigationCmdID || uiCmd == m_Impl.m_nWindowsPrevNavigationCmdID))
		{
			m_Impl.ShowWindowsNavigator(uiCmd == m_Impl.m_nWindowsPrevNavigationCmdID);
			return TRUE;
		}

		if (m_bAutoPaneActivation)
		{
			if ((uiCmd == ID_PREV_PANE || uiCmd == ID_NEXT_PANE) && ActivateNextPane(uiCmd == ID_PREV_PANE))
			{
				return TRUE;
			}
		}

#ifndef BCGP_EXCLUDE_EDIT_CTRL
		if (CBCGPIntelliSenseWnd::m_hwndActive != NULL && ::IsWindow(CBCGPIntelliSenseWnd::m_hwndActive))
		{
			::SendMessage(CBCGPIntelliSenseWnd::m_hwndActive, WM_CLOSE, 0, 0);
		}
#endif

		//---------------------------
		// Simmulate ESC keystroke...
		//---------------------------
		if (m_Impl.ProcessKeyboard (VK_ESCAPE))
		{
			return TRUE;
		}

		if (g_pUserToolsManager != NULL &&
			g_pUserToolsManager->InvokeTool (uiCmd))
		{
			return TRUE;
		}
	}

	if (!CBCGPToolBar::IsCustomizeMode ())
	{
		return CMDIFrameWnd::OnCommand(wParam, lParam);
	}

	return FALSE;
}
//*******************************************************************************************
HMENU CBCGPMDIFrameWnd::GetWindowMenuPopup (HMENU hMenuBar)
{
	m_hmenuWindow = CMDIFrameWnd::GetWindowMenuPopup (hMenuBar);
	return m_hmenuWindow;
}
//********************************************************************************************
BOOL CBCGPMDIFrameWnd::LoadFrame(UINT nIDResource, DWORD dwDefaultStyle, CWnd* pParentWnd, CCreateContext* pContext) 
{
	m_Impl.m_bIsLoadingFrame = TRUE;
	m_Impl.m_nIDDefaultResource = nIDResource;
	m_Impl.LoadLargeIconsState ();
	
	if (!CMDIFrameWnd::LoadFrame(nIDResource, dwDefaultStyle, pParentWnd, pContext))
	{
		m_Impl.m_bIsLoadingFrame = FALSE;
		return FALSE;
	}

	m_Impl.OnLoadFrame ();

	if (GetMenuBar () != NULL)
	{
		m_hMenuDefault = m_Impl.m_hDefaultMenu;
	}

	m_Impl.m_bIsLoadingFrame = FALSE;
	return TRUE;
}
//*******************************************************************************************
void CBCGPMDIFrameWnd::OnClose() 
{
	if (m_Impl.m_pLightBoxDlg != NULL && m_Impl.m_pLightBoxDlg->IsAnimated())
	{
		return;
	}

	if (!m_Impl.CloseLightBox(FALSE))
	{
		return;
	}

	if (m_pPrintPreviewFrame != NULL)
	{
		m_pPrintPreviewFrame->SendMessage (WM_COMMAND, AFX_ID_PREVIEW_CLOSE);
		m_pPrintPreviewFrame = NULL;
		return;
	}

	if (m_bIsMDIChildDetached)
	{
		CPushRoutingFrame prf(this);

		HWND hwndMDIChild = ::GetWindow(m_hWndMDIClient, GW_CHILD);
		
		while (hwndMDIChild != NULL)
		{
			CBCGPMDIChildWnd* pMDIChildFrame = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, CWnd::FromHandle(hwndMDIChild));
			if (pMDIChildFrame != NULL)
			{
				ASSERT_VALID(pMDIChildFrame);
				
				CDocument* pDoc = pMDIChildFrame->GetActiveDocument();
				if (pDoc != NULL)
				{
					ASSERT_VALID(pDoc);

					if (!pDoc->SaveModified())
					{
						BringWindowToTop();
						SetActiveWindow();
						return;
					}
				}
			}
			
			hwndMDIChild = ::GetWindow(hwndMDIChild, GW_HWNDNEXT);
		}

		CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, AfxGetMainWnd());
		if (pMainFrame->GetSafeHwnd() != NULL)
		{
#ifndef BCGP_EXCLUDE_RIBBON
			if (pMainFrame->m_Impl.m_pBackstageView->GetSafeHwnd() != NULL)
			{
				pMainFrame->m_Impl.m_pBackstageView->SendMessage(WM_CLOSE);
				pMainFrame->BringWindowToTop();
				pMainFrame->SetActiveWindow();
			}
#endif
#if _MSC_VER < 1300
			if (pMainFrame->IsPrintPreview())
			{
				pMainFrame->SendMessage(WM_CLOSE);
			}
#endif
		}
	}

	if (!m_Impl.IsPrintPreview ())
	{
		m_bClosing = TRUE;

		// Deactivate OLE container first:
		COleClientItem*	pActiveItem = GetInPlaceActiveItem ();
		if (pActiveItem != NULL)
		{
			pActiveItem->Deactivate ();
		}

		m_Impl.OnCloseFrame();
	}

	HWND hwndThis = GetSafeHwnd();

	CMDIFrameWnd::OnClose();

	if (::IsWindow(hwndThis))
	{
		m_bClosing = FALSE;
	}
}
//***************************************************************************************
void CBCGPMDIFrameWnd::ClosePrintPreview()
{
#if _MSC_VER >= 1300
	CBCGPMDIChildWnd* pMDIChild = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, MDIGetActive ());
	if (pMDIChild != NULL && pMDIChild->IsPrintPreview())
	{
		pMDIChild->SendMessage (WM_COMMAND, AFX_ID_PREVIEW_CLOSE);
	}
#else
	if (IsPrintPreview())
	{
		SendMessage(WM_CLOSE);
	}
#endif
}
//*******************************************************************************************
BOOL CBCGPMDIFrameWnd::PreCreateWindow(CREATESTRUCT& cs) 
{
	m_Impl.SetDockManager (&m_dockManager);
	m_Impl.RestorePosition(cs);	
	return CMDIFrameWnd::PreCreateWindow(cs);
}
//*******************************************************************************************
BOOL CBCGPMDIFrameWnd::ShowPopupMenu (CBCGPPopupMenu* pMenuPopup)
{
	g_mapWindowCommands.RemoveAll();

	if (!m_Impl.OnShowPopupMenu (pMenuPopup, this))
	{
		return FALSE;
	}

	if (!CBCGPToolBar::IsCustomizeMode () && m_hmenuWindow != NULL &&
		pMenuPopup != NULL && pMenuPopup->GetHMenu () != NULL)
	{
		//-----------------------------------------------------------
		// Check the popup menu for the "Windows..." menu maching...:
		//-----------------------------------------------------------
		HMENU hMenuPop = pMenuPopup->GetHMenu ();
		BOOL bIsWindowMenu = FALSE;

		int iItemMax = ::GetMenuItemCount (hMenuPop);
		for (int iItemPop = 0; !bIsWindowMenu && iItemPop < iItemMax; iItemPop ++)
		{
			UINT nID = ::GetMenuItemID( hMenuPop, iItemPop);
			bIsWindowMenu =  (nID >= AFX_IDM_WINDOW_FIRST && nID <= AFX_IDM_WINDOW_LAST);
		}

		if (bIsWindowMenu)
		{
			PrepareWindowsMenu(*pMenuPopup);
		}
	}

	if (pMenuPopup != NULL && pMenuPopup->m_bShown)
	{
		return TRUE;
	}

	if (IsTearOff())
	{
		CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, AfxGetMainWnd());
		if (pMainFrame->GetSafeHwnd() != NULL)
		{
			return pMainFrame->OnShowPopupMenu(pMenuPopup);
		}
	}

	return OnShowPopupMenu(pMenuPopup); 
}
//**********************************************************************************
void CBCGPMDIFrameWnd::OnClosePopupMenu (CBCGPPopupMenu* pMenuPopup)
{
	if (globalData.IsAccessibilitySupport () && pMenuPopup != NULL)
	{
		CBCGPPopupMenu* pPopupParent = pMenuPopup->GetParentPopupMenu ();
		CBCGPToolbarMenuButton* pParentButton  = pMenuPopup->GetParentButton ();

		if (pMenuPopup->IsEscClose () || pPopupParent != NULL || pParentButton == NULL)
		{
			globalData.NotifyWinEvent (EVENT_SYSTEM_MENUPOPUPEND,
				pMenuPopup->GetSafeHwnd (), OBJID_WINDOW, CHILDID_SELF);
		}
		else
		{
			globalData.NotifyWinEvent (EVENT_SYSTEM_MENUEND,
				pMenuPopup->GetSafeHwnd (), OBJID_WINDOW, CHILDID_SELF);
		}
				

		CBCGPMenuBar* pMenuBar = (CBCGPMenuBar*)GetMenuBar ();
		
		if (pMenuBar->GetSafeHwnd () != NULL && pMenuBar == GetFocus ())
		{
			pMenuBar->AccNotifyObjectFocusEvent ();
		}
	}

	if (CBCGPPopupMenu::m_pActivePopupMenu == pMenuPopup)
	{
		CBCGPPopupMenu::m_pActivePopupMenu = NULL;
	}

	m_dockManager.OnClosePopupMenu ();
}
//*******************************************************************************************
BOOL CBCGPMDIFrameWnd::OnDrawMenuImage (CDC* pDC, 
									const CBCGPToolbarMenuButton* pMenuButton, 
									const CRect& rectImage)
{
	ASSERT_VALID (this);

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_bIsMDIChildDetached)
	{
		CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, AfxGetMainWnd());
		if (pMainFrame->GetSafeHwnd() != NULL)
		{
			return pMainFrame->OnDrawMenuImage(pDC, pMenuButton, rectImage);
		}
	}

	if (m_Impl.m_pRibbonBar != NULL)
	{
		ASSERT_VALID (m_Impl.m_pRibbonBar);
		return m_Impl.m_pRibbonBar->DrawMenuImage (pDC, pMenuButton, rectImage);
	}

#endif

	return FALSE;
}
//*******************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnToolbarCreateNew(WPARAM,LPARAM lp)
{
	ASSERT (lp != NULL);
	return (LRESULT) m_Impl.CreateNewToolBar ((LPCTSTR) lp);
}
//***************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnToolbarDelete(WPARAM,LPARAM lp)
{
	CBCGPToolBar* pToolbar = (CBCGPToolBar*) lp;
	ASSERT_VALID (pToolbar);

	return (LRESULT) m_Impl.DeleteToolBar (pToolbar);
}
//***************************************************************************************
#if _MSC_VER >= 1300
void CBCGPMDIFrameWnd::HtmlHelp(DWORD_PTR dwData, UINT nCmd)
{
	if (dwData > 0 || !m_bContextHelp)
	{
		CMDIFrameWnd::HtmlHelp(dwData, nCmd);
	}
	else
	{
		OnContextHelp ();
	}
}
#endif
//***************************************************************************************
void CBCGPMDIFrameWnd::WinHelp(DWORD_PTR dwData, UINT nCmd) 
{
	if (dwData > 0 || !m_bContextHelp)
	{
		CMDIFrameWnd::WinHelp(dwData, nCmd);
	}
	else
	{
		OnContextHelp ();
	}
}
//***************************************************************************************
void CBCGPMDIFrameWnd::OnContextHelp ()
{
	m_bContextHelp = TRUE;

	if (!m_bHelpMode && CanEnterHelpMode())
	{
		CBCGPToolBar::SetHelpMode ();
	}

	CMDIFrameWnd::OnContextHelp ();

	if (!m_bHelpMode)
	{
		CBCGPToolBar::SetHelpMode (FALSE);
	}

	m_bContextHelp = FALSE;
}
//***************************************************************************************
BOOL CBCGPMDIFrameWnd::DockControlBarLeftOf (CBCGPControlBar* pBar, CBCGPControlBar* pLeftOf)
{
	return m_dockManager.DockControlBarLeftOf (pBar, pLeftOf);
}
//***************************************************************************************
void CBCGPMDIFrameWnd::OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized) 
{
	if (!m_bIsMDIChildDetached)
	{
		m_hwndLastActiveDetachedMDIFrame = NULL;
	}

#ifndef BCGP_EXCLUDE_POPUP_WINDOW
	CBCGPPopupWindow* pPopup = DYNAMIC_DOWNCAST(CBCGPPopupWindow, pWndOther);
	if (pPopup != NULL && pPopup->m_bDontChangeActiveStatus)
	{
		return;
	}
#endif

	if (nState == WA_CLICKACTIVE && !m_bIsMDIChildDetached)
	{
		CPoint ptCursor;
		::GetCursorPos(&ptCursor);

		CWnd* pWndUnderCursor = CWnd::WindowFromPoint(ptCursor);

		if (DYNAMIC_DOWNCAST(CBCGPToolBar, pWndUnderCursor) != NULL
#ifndef BCGP_EXCLUDE_RIBBON
			|| DYNAMIC_DOWNCAST(CBCGPRibbonBar, pWndUnderCursor) != NULL
#endif
		)
		{
			CBCGPMDIFrameWnd* pOtherMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, pWndOther);
			if (pOtherMDIFrame->GetSafeHwnd() != NULL && pOtherMDIFrame->m_bIsMDIChildDetached)
			{
				m_hwndLastActiveDetachedMDIFrame = pOtherMDIFrame->GetSafeHwnd();
			}
		}
	}
	else if (nState == WA_ACTIVE && m_bActivatedBySysKey && !m_bIsMDIChildDetached)
	{
		CBCGPMDIFrameWnd* pOtherMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, pWndOther);
		if (pOtherMDIFrame->GetSafeHwnd() != NULL && pOtherMDIFrame->m_bIsMDIChildDetached)
		{
			m_hwndLastActiveDetachedMDIFrame = pOtherMDIFrame->GetSafeHwnd();
		}
	}
	
	CMDIFrameWnd::OnActivate(nState, pWndOther, bMinimized);

	switch (nState)
	{
	case WA_CLICKACTIVE:
		UpdateWindow ();
		break;

	case WA_INACTIVE:
		if (!CBCGPToolBar::IsCustomizeMode ())
		{
			m_Impl.DeactivateMenu ();
		}
		break;
	}

	m_Impl.AdjustShadow(nState != WA_INACTIVE && !bMinimized);

	if (nState != WA_INACTIVE)
	{
		m_Impl.ActivateRibbonBackstageView();

		if (AreMDIChildrenTearOff() && !CBCGPToolBar::IsCustomizeMode() && GetLightBoxDialog() == NULL)
		{
			HMENU hMenu = GetActiveMDITearOffMenu();

			if (!m_bIsMDIChildDetached)
			{
				OnUpdateFrameMenu(hMenu);
			}
			else
			{
				CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, AfxGetMainWnd());
				if (pMainFrame->GetSafeHwnd() != NULL)
				{
					pMainFrame->OnUpdateFrameMenu(hMenu);
				}
			}
		}
	}
}
//***************************************************************************************
#if _MSC_VER >= 1300
void CBCGPMDIFrameWnd::OnActivateApp(BOOL bActive, DWORD /*dwThreadID*/)
#else
void CBCGPMDIFrameWnd::OnActivateApp(BOOL bActive, HTASK /*hTask*/) 
#endif
{
	m_dockManager.OnActivateFrame (bActive);
	m_Impl.OnActivateApp (bActive);
}
//***************************************************************************************
void CBCGPMDIFrameWnd::EnableWindowsDialog (UINT uiMenuId, 
										   LPCTSTR lpszMenuText,
										   BOOL bShowAllways,
										   BOOL bShowHelpButton,
										   BOOL bResizable)
{
	ASSERT (lpszMenuText != NULL);
	ASSERT (uiMenuId != 0);

	m_uiWindowsDlgMenuId = uiMenuId;
	m_strWindowsDlgMenuText = lpszMenuText;
	m_bShowWindowsDlgAlways = bShowAllways;
	m_bShowWindowsDlgHelpButton = bShowHelpButton;
	m_bResizableWindowsDlg = bResizable;
}
//****************************************************************************
void CBCGPMDIFrameWnd::EnableWindowsDialog (UINT uiMenuId, 
										   UINT uiMenuTextResId,
										   BOOL bShowAllways,
										   BOOL bShowHelpButton,
										   BOOL bResizable)
{
	CString strMenuText;
	VERIFY (strMenuText.LoadString (uiMenuTextResId));

	EnableWindowsDialog (uiMenuId, strMenuText, bShowAllways, bShowHelpButton, bResizable);
}
//****************************************************************************
void CBCGPMDIFrameWnd::ShowWindowsDialog ()
{
	CBCGPLocalResource locaRes;

	CBCGPWindowsManagerDlg dlg (this, m_bShowWindowsDlgHelpButton, m_bResizableWindowsDlg);
	dlg.DoModal ();
}
//****************************************************************************
COleClientItem*	CBCGPMDIFrameWnd::GetInPlaceActiveItem ()
{
	CFrameWnd* pActiveFrame = GetActiveFrame ();
	if (pActiveFrame == NULL)
	{
		return NULL;
	}

	ASSERT_VALID (pActiveFrame);

	CView* pView = pActiveFrame->GetActiveView ();
	if (pView == NULL || pView->IsKindOf (RUNTIME_CLASS (CBCGPPrintPreviewView)))
	{
		return NULL;
	}

	ASSERT_VALID (pView);

	COleDocument* pDoc = DYNAMIC_DOWNCAST (COleDocument, pView->GetDocument ());
	if (pDoc == NULL)
	{
		return NULL;
	}

	ASSERT_VALID (pDoc);
	return pDoc->GetInPlaceActiveItem (pView);
}
//****************************************************************************
HMENU CBCGPMDIFrameWnd::GetActiveMDITearOffMenu() const
{
	CBCGPMDIFrameWnd* pActiveMDIFrame = GetActiveTearOffFrame();
	if (pActiveMDIFrame != NULL && pActiveMDIFrame->GetSafeHwnd() != GetSafeHwnd())
	{
		CMDIChildWnd* pActiveWnd = pActiveMDIFrame->MDIGetActive();
		if (pActiveWnd != NULL)
		{
			// attempt to get default menu from document
			CDocument* pDoc = pActiveWnd->GetActiveDocument();
			if (pDoc != NULL)
			{
				CMultiDocTemplate* pTemplate = DYNAMIC_DOWNCAST(CMultiDocTemplate, pDoc->GetDocTemplate());
				if (pTemplate != NULL)
				{
					return pTemplate->m_hMenuShared;
				}
			}
		}
	}

	return NULL;
}
//****************************************************************************
void CBCGPMDIFrameWnd::OnUpdateFrameMenu (HMENU hMenuAlt)
{
	CMDIFrameWnd::OnUpdateFrameMenu (hMenuAlt);

	BOOL bIsMenuBar = m_Impl.m_pMenuBar != NULL &&
		(m_Impl.m_pMenuBar->GetStyle () & WS_VISIBLE);

	BOOL bIsRibbon = FALSE;

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_Impl.m_pRibbonBar != NULL &&
		(m_Impl.m_pRibbonBar->GetStyle () & WS_VISIBLE))
	{
		bIsRibbon = TRUE;
	}
#endif

	if (bIsMenuBar || bIsRibbon)
	{
		COleClientItem*	pActiveItem = GetInPlaceActiveItem ();
		if (pActiveItem == NULL ||
			pActiveItem->GetInPlaceWindow () == NULL)
		{
			SetMenu (NULL);
		}
		else
		{
			SetMenu (CMenu::FromHandle (hMenuAlt));
		}
	}
}
//****************************************************************************************
void CBCGPMDIFrameWnd::OnDestroy() 
{
    if (m_hAccelTable != NULL)
    {
        ::DestroyAcceleratorTable (m_hAccelTable);
        m_hAccelTable = NULL;
    }

	m_dockManager.m_bEnableAdjustLayout = FALSE;

	CList<HWND, HWND> lstChildren;

	for (int i = 0; i < 2; i++)
	{
		CWnd* pNextWnd = GetTopWindow ();
		while (pNextWnd != NULL)
		{
			if (m_wndClientArea.m_hWnd != pNextWnd->m_hWnd)
			{
				const BOOL bIsSlider = pNextWnd->IsKindOf(RUNTIME_CLASS(CBCGPSlider));

				if ((i == 0 && !bIsSlider) || (i == 1 && bIsSlider))
				{
					lstChildren.AddTail (pNextWnd->m_hWnd);
				}
			}

			pNextWnd = pNextWnd->GetNextWindow ();
		}
	}

	for (POSITION pos = lstChildren.GetHeadPosition (); pos != NULL;)
	{
		HWND hwndNext = lstChildren.GetNext (pos);
		if (IsWindow (hwndNext) && ::GetParent (hwndNext) == m_hWnd)
		{
			::DestroyWindow (hwndNext);
		}
	}


	CBCGPFrameImpl::RemoveFrame (this);
	CMDIFrameWnd::OnDestroy();
}
//*************************************************************************************
int CBCGPMDIFrameWnd::ReattachTearOffWindows()
{
	if (GetSafeHwnd() == NULL)
	{
		return 0;
	}

	if (!AreMDIChildrenTearOff() || m_bIsMDIChildDetached)
	{
		return 0;
	}

	POSITION pos = NULL;
	CList<CBCGPMDIChildWnd*, CBCGPMDIChildWnd*> lstChildren;

	const CList<CFrameWnd*, CFrameWnd*>& lstFrames = CBCGPFrameImpl::GetFrameList();
	for (pos = lstFrames.GetHeadPosition(); pos != NULL;)
	{
		CBCGPMDIFrameWnd* pMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, lstFrames.GetNext(pos));
		
		if (pMDIFrame->GetSafeHwnd() != NULL && pMDIFrame->m_bIsMDIChildDetached)
		{
			for (HWND hwndMDIChild = ::GetWindow (pMDIFrame->m_hWndMDIClient, GW_CHILD); hwndMDIChild != NULL; hwndMDIChild = ::GetWindow (hwndMDIChild, GW_HWNDNEXT))
			{
				pMDIFrame->SetRedraw(FALSE);

				CBCGPMDIChildWnd* pMDIChild = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, CWnd::FromHandle(hwndMDIChild));
				if (pMDIChild != NULL)
				{
					lstChildren.AddTail(pMDIChild);
				}
			}
		}
	}

	SetRedraw(FALSE);

	for (pos = lstChildren.GetHeadPosition(); pos != NULL;)
	{
		AttachMDIChild(lstChildren.GetNext(pos));
	}

	SetRedraw();
	RedrawWindow();

	return (int)lstChildren.GetCount();
}
//*************************************************************************************
void CBCGPMDIFrameWnd::EnableMDITabbedGroups (BOOL bEnable, const CBCGPMDITabParams& params)
{
	if (!bEnable)
	{
		// Move all MDI children back to this frame:
		ReattachTearOffWindows();
	}

	m_wndClientArea.EnableMDITabbedGroups (bEnable, params);

	if (bEnable && AreMDIChildrenTearOff() && !m_bIsMDIChildDetached)
	{
		const CList<CFrameWnd*, CFrameWnd*>& lstFrames = CBCGPFrameImpl::GetFrameList();
		for (POSITION pos = lstFrames.GetHeadPosition(); pos != NULL;)
		{
			CBCGPMDIFrameWnd* pMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, lstFrames.GetNext(pos));
			
			if (pMDIFrame->GetSafeHwnd() != NULL && pMDIFrame->m_bIsMDIChildDetached)
			{
				pMDIFrame->PostMessage(BCGM_SYNC_FRAME_OPTIONS, 0, (LPARAM) GetSafeHwnd());
			}
		}
	}
}
//*************************************************************************************
void CBCGPMDIFrameWnd::EnableMDITabs (BOOL bEnable/* = TRUE*/,
				BOOL bIcons/* = TRUE*/,
				CBCGPTabWnd::Location tabLocation /* = CBCGPTabWnd::LOCATION_BOTTOM*/,
				BOOL bTabCloseButton/* = FALSE*/,
				CBCGPTabWnd::Style style/* = CBCGPTabWnd::STYLE_3D_SCROLLED*/,
				BOOL bTabCustomTooltips/* = FALSE*/,
				BOOL bActiveTabCloseButton/* = FALSE*/)
{
	ASSERT (style == CBCGPTabWnd::STYLE_3D_SCROLLED ||
			style == CBCGPTabWnd::STYLE_3D_ONENOTE ||
			style == CBCGPTabWnd::STYLE_3D_VS2005 ||
			style == CBCGPTabWnd::STYLE_3D_ROUNDED ||
			style == CBCGPTabWnd::STYLE_3D_ROUNDED_SCROLL ||
			style == CBCGPTabWnd::STYLE_POINTER ||
			style == CBCGPTabWnd::STYLE_UNDERLINE);

	if (bEnable)
	{
		// Move all MDI children back to this frame:
		ReattachTearOffWindows();
	}

	CBCGPMDITabParams params;
	params.m_style = style;
	params.m_tabLocation = tabLocation;
	params.m_bTabIcons = bIcons;
	params.m_bTabCloseButton = bTabCloseButton;
	params.m_bTabCustomTooltips = bTabCustomTooltips;
	params.m_bActiveTabCloseButton = bActiveTabCloseButton;

	m_wndClientArea.EnableMDITabs (bEnable, params);
}
//*************************************************************************************
void CBCGPMDIFrameWnd::OnSetPreviewMode(BOOL bPreview, CPrintPreviewState* pState)
{
	ASSERT_VALID (this);

	if (m_wndClientArea.DoesMDITabExist ())
	{
		m_wndClientArea.m_bTabIsVisible = !bPreview;
		((CWnd&)m_wndClientArea.GetMDITabs ()).ShowWindow (
			bPreview ? SW_HIDE : SW_SHOWNOACTIVATE);
	}

	m_dockManager.SetPrintPreviewMode (bPreview, pState);

	DWORD dwSavedState = pState->dwStates;
	CMDIFrameWnd::OnSetPreviewMode (bPreview, pState);
	pState->dwStates = dwSavedState;

	AdjustDockingLayout ();
	RecalcLayout ();

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_Impl.m_pRibbonBar != NULL &&
		m_Impl.m_pRibbonBar->IsReplaceFrameCaption ())
	{
		PostMessage (BCGM_POSTSETPREVIEWFRAME, bPreview);
	}
#endif

#if _MSC_VER < 1300
	if (g_pWorkspace != NULL && g_pWorkspace->IsTaskBarInteractionEnabled())
	{
		RegisterAllMDIChildrenWithTaskbar(!bPreview);
	}
#endif
}
//*************************************************************************************
BOOL CBCGPMDIFrameWnd::OnShowControlBars (BOOL bShow)
{
	ASSERT_VALID (this);
	
	BOOL bRibbonWasHidden = FALSE;

#ifndef BCGP_EXCLUDE_RIBBON
	if (bShow &&
		m_Impl.m_pRibbonBar != NULL &&
		(m_Impl.m_pRibbonBar->GetStyle () & WS_VISIBLE) == 0)
	{
		bRibbonWasHidden = TRUE;
	}
#endif

	BOOL bResult = m_dockManager.ShowControlBars (bShow);
	AdjustDockingLayout ();
	
#ifndef BCGP_EXCLUDE_RIBBON
	if (bShow && bRibbonWasHidden &&
		m_Impl.m_pRibbonBar != NULL &&
		(m_Impl.m_pRibbonBar->GetStyle () & WS_VISIBLE))
	{
		m_Impl.m_pRibbonBar->RedrawWindow(NULL, NULL, 
			RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
	}
#endif

	if (!bShow)
	{
		AdjustClientArea();
	}

	return bResult;
}

//*****************************************************************************
//******************* dockmanager layer ***************************************
//*****************************************************************************
void CBCGPMDIFrameWnd::AddDockBar ()
{
	ASSERT_VALID (this);
}
//*****************************************************************************
BOOL CBCGPMDIFrameWnd::AddControlBar (CBCGPBaseControlBar* pControlBar, BOOL bTail)
{
	ASSERT_VALID (this);

#ifndef BCGP_EXCLUDE_RIBBON
	CBCGPRibbonBar* pRibbonBar = DYNAMIC_DOWNCAST (CBCGPRibbonBar, pControlBar);
	if (pRibbonBar != NULL)
	{
		ASSERT_VALID (pRibbonBar);

		if (pRibbonBar->IsMainRibbonBar ())
		{
			m_Impl.m_pRibbonBar = pRibbonBar;
		}
	}

	CBCGPRibbonStatusBar* pRibbonStatusBar = DYNAMIC_DOWNCAST (CBCGPRibbonStatusBar, pControlBar);
	if (pRibbonStatusBar != NULL)
	{
		ASSERT_VALID (pRibbonStatusBar);
		m_Impl.m_pRibbonStatusBar = pRibbonStatusBar;
	}

#endif

	return m_dockManager.AddControlBar (pControlBar, bTail);
}
//*****************************************************************************
BOOL CBCGPMDIFrameWnd::InsertControlBar (CBCGPBaseControlBar* pControlBar, 
									  CBCGPBaseControlBar* pTarget, BOOL bAfter)
{
	ASSERT_VALID (this);
	return m_dockManager.InsertControlBar (pControlBar, pTarget, bAfter);
}
//*****************************************************************************
void CBCGPMDIFrameWnd::RemoveControlBarFromDockManager (CBCGPBaseControlBar* pControlBar, BOOL bDestroy,
										 BOOL bAdjustLayout, BOOL bAutoHide, CBCGPBaseControlBar* pBarReplacement)
{
	ASSERT_VALID (this);
	m_dockManager.RemoveControlBarFromDockManager (pControlBar, bDestroy, bAdjustLayout, bAutoHide, pBarReplacement);
}
//*****************************************************************************
void CBCGPMDIFrameWnd::DockControlBar (CBCGPBaseControlBar* pBar, UINT nDockBarID, 
									LPCRECT lpRect)
{
	ASSERT_VALID (this);
	m_dockManager.DockControlBar (pBar, nDockBarID, lpRect);
}
//*****************************************************************************
CBCGPBaseControlBar* CBCGPMDIFrameWnd::GetControlBar (UINT nID)
{
	ASSERT_VALID (this);
	
	CBCGPBaseControlBar* pBar = m_dockManager.FindBarByID (nID, TRUE);

	return pBar;
}
//*****************************************************************************
void CBCGPMDIFrameWnd::ShowControlBar (CBCGPBaseControlBar* pBar, BOOL bShow, 
									   BOOL bDelay, BOOL bActivate)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pBar);

	pBar->ShowControlBar (bShow, bDelay, bActivate);
}
//*****************************************************************************
CBCGPBaseControlBar* CBCGPMDIFrameWnd::ControlBarFromPoint (CPoint point, 
							int nSensitivity, bool bExactBar, 
							CRuntimeClass* pRTCBarType) const
{
	ASSERT_VALID (this);
	return m_dockManager.ControlBarFromPoint (point, nSensitivity, bExactBar, 
												pRTCBarType);
}
//*****************************************************************************
CBCGPBaseControlBar* CBCGPMDIFrameWnd::ControlBarFromPoint (CPoint point, 
								int nSensitivity, DWORD& dwAlignment, 
								CRuntimeClass* pRTCBarType) const
{
	ASSERT_VALID (this);
	return m_dockManager.ControlBarFromPoint (point, nSensitivity, dwAlignment, 
												pRTCBarType);
}
//*****************************************************************************
BOOL CBCGPMDIFrameWnd::IsPointNearDockBar (CPoint point, DWORD& dwBarAlignment, 
										   BOOL& bOuterEdge) const
{
	ASSERT_VALID (this);
	return m_dockManager.IsPointNearDockBar (point, dwBarAlignment, bOuterEdge);
}
//*****************************************************************************
void CBCGPMDIFrameWnd::AdjustDockingLayout (HDWP hdwp)
{
	ASSERT_VALID (this);

	if (m_dockManager.IsInAdjustLayout ())
	{
		return;
	}

	m_dockManager.AdjustDockingLayout (hdwp);
	AdjustClientArea ();
	
	if (m_dockManager.IsOLEContainerMode ())
	{
		RecalcLayout ();
	}

}
//*****************************************************************************
void CBCGPMDIFrameWnd::AdjustClientArea ()
{
	CRect rectClientAreaBounds = m_dockManager.GetClientAreaBounds ();

	rectClientAreaBounds.left += m_rectBorder.left;
	rectClientAreaBounds.top  += m_rectBorder.top;
	rectClientAreaBounds.right -= m_rectBorder.right;
	rectClientAreaBounds.bottom -= m_rectBorder.bottom;

	if (m_wndClientArea.GetSafeHwnd () != NULL)
	{
		m_wndClientArea.CalcWindowRect (rectClientAreaBounds, 0);
	}
}
//*****************************************************************************
BOOL CBCGPMDIFrameWnd::OnMoveMiniFrame	(CWnd* pFrame)
{
	ASSERT_VALID (this);
	return m_dockManager.OnMoveMiniFrame (pFrame);
}
//****************************************************************************************
BOOL CBCGPMDIFrameWnd::EnableDocking (DWORD dwDockStyle)
{
	return m_dockManager.EnableDocking (dwDockStyle);
}
//****************************************************************************************
BOOL CBCGPMDIFrameWnd::EnableAutoHideBars (DWORD dwDockStyle, BOOL bActivateOnMouseClick)
{
	return m_dockManager.EnableAutoHideBars (dwDockStyle, bActivateOnMouseClick);
}
//****************************************************************************************
void CBCGPMDIFrameWnd::EnableMaximizeFloatingBars(BOOL bEnable, BOOL bMaximizeByDblClick, BOOL bRestoreMaximizeFloatingBars)
{
	m_dockManager.EnableMaximizeFloatingBars(bEnable, bMaximizeByDblClick, bRestoreMaximizeFloatingBars);
}
//****************************************************************************************
BOOL CBCGPMDIFrameWnd::AreFloatingBarsCanBeMaximized() const
{
	return m_dockManager.AreFloatingBarsCanBeMaximized();
}
//****************************************************************************************
void CBCGPMDIFrameWnd::EnableFloatingBarsMenuButton(BOOL bEnable)
{
	m_dockManager.EnableFloatingBarsMenuButton(bEnable);
}
//****************************************************************************************
BOOL CBCGPMDIFrameWnd::AreFloatingBarsHaveMenuButton() const
{
	return m_dockManager.AreFloatingBarsHaveMenuButton();
}
//****************************************************************************************
void CBCGPMDIFrameWnd::RecalcLayout (BOOL bNotify)
{
	if (m_bInRecalcLayout)
		return;

	m_bInRecalcLayout = TRUE;

	BOOL bWasOleInPlaceActive = m_Impl.m_bIsOleInPlaceActive;
	m_Impl.m_bIsOleInPlaceActive = FALSE;

	COleClientItem*	pActiveItem = GetInPlaceActiveItem ();

	if (pActiveItem != NULL && pActiveItem->m_pInPlaceFrame != NULL &&
		pActiveItem->GetItemState () == COleClientItem::activeUIState)
	{
		m_Impl.m_bIsOleInPlaceActive = TRUE;
		m_Impl.m_bHadCaption = (GetStyle () & WS_CAPTION) != 0;
	}

	if (!m_bIsMinimized)
	{
		CView* pView = GetActiveView ();

		if (m_dockManager.IsPrintPreviewValid () ||
			m_dockManager.IsOLEContainerMode ())
		{
			if (pView != NULL && pView->IsKindOf (RUNTIME_CLASS (CBCGPPrintPreviewView)))
			{
	
				m_dockManager.RecalcLayout (bNotify);
				CRect rectClient = m_dockManager.GetClientAreaBounds ();
				pView->SetWindowPos (NULL, rectClient.left, rectClient.top, 
										rectClient.Width (), rectClient.Height (),
										SWP_NOZORDER  | SWP_NOACTIVATE);
			}
			else 
			{
				if (bNotify && m_dockManager.IsOLEContainerMode ())
				{
					ActiveItemRecalcLayout ();
				}
				else
				{
					m_bInRecalcLayout = FALSE;
					CMDIFrameWnd::RecalcLayout (bNotify);

					AdjustClientArea ();
				}
			}
		}
		else
		{
			m_dockManager.RecalcLayout (bNotify);
			AdjustClientArea ();
		}
	}

	if (bWasOleInPlaceActive != m_Impl.m_bIsOleInPlaceActive)
	{
		if (!m_Impl.m_bHadCaption)
		{
			if (m_Impl.m_bIsOleInPlaceActive)
			{
				ModifyStyle (0, WS_CAPTION);
			}
			else
			{
				ModifyStyle (WS_CAPTION, 0);
			}
		}

		PostMessage(BCGM_CHANGEVISUALMANAGER);
	}

	m_bInRecalcLayout = FALSE;
}
//****************************************************************************************
void CBCGPMDIFrameWnd::ActiveItemRecalcLayout ()
{
	COleClientItem*	pActiveItem = GetInPlaceActiveItem ();
	CView* pView = GetActiveView ();

	if (pActiveItem != NULL && (pActiveItem->m_pView == pView || pView == NULL))
	{
		if (pActiveItem->m_pInPlaceFrame != NULL)
		{
			CRect rectBounds = m_dockManager.GetClientAreaBounds ();
			pActiveItem->m_pInPlaceFrame->OnRecalcLayout ();
		}

		CView* pActiveView = pActiveItem->GetActiveView ();
		if (pActiveView != NULL)
		{
			CBCGPMDIChildWnd* pFrame = (CBCGPMDIChildWnd*) pActiveView->GetParentFrame ();
			
			if (pFrame != NULL && pFrame->m_bActivating)
			{
				pActiveItem->m_pInPlaceFrame->OnRecalcLayout ();		
			}
		}
	}

	AdjustClientArea ();
}
//****************************************************************************************
BOOL CBCGPMDIFrameWnd::NegotiateBorderSpace( UINT nBorderCmd, LPRECT lpRectBorder )
{
	CRect border, request;

	switch (nBorderCmd)
	{
	case borderGet:
	{
		CMDIFrameWnd::NegotiateBorderSpace (nBorderCmd, lpRectBorder);
		m_dockManager.AdjustDockingLayout ();
		CRect rectBounds = m_dockManager.GetClientAreaBounds ();
		ASSERT(lpRectBorder != NULL);

		*lpRectBorder = rectBounds;
		break;
	}
	case borderRequest:
		return TRUE;

	case borderSet:
		return CMDIFrameWnd::NegotiateBorderSpace (nBorderCmd, lpRectBorder);

	default:
		ASSERT(FALSE);  // invalid CFrameWnd::BorderCmd
	}

	return TRUE;
}
//****************************************************************************************
int CBCGPMDIFrameWnd::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CMDIFrameWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	m_dockManager.Create (this);	

	m_Impl.m_bHasBorder = (GetStyle () & WS_BORDER) != 0;
	m_Impl.m_bHasCaption = (GetStyle() & WS_CAPTION) != 0;

	CBCGPFrameImpl::AddFrame (this);

	OnChangeVisualManager (0, 0);
	return 0;
}
//*************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnEnterSizeMove(WPARAM, LPARAM)
{
	if (m_Impl.m_pLightBoxDlg->GetSafeHwnd() != NULL)
	{
		m_Impl.m_pLightBoxDlg->Activate();
	}
	
	return Default();
}
//****************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnExitSizeMove (WPARAM, LPARAM)
{
	if (m_Impl.m_pLightBoxDlg->GetSafeHwnd() != NULL)
	{
		m_Impl.m_pLightBoxDlg->Activate();
	}
	
	RecalcLayout ();
	m_dockManager.FixupVirtualRects ();
	return 0;
}
//****************************************************************************************
void CBCGPMDIFrameWnd::OnUpdateControlBarMenu(CCmdUI* pCmdUI)
{
	CBCGPBaseControlBar* pBar = GetControlBar(pCmdUI->m_nID);
	if (pBar != NULL)
	{
		pCmdUI->SetCheck(pBar->IsWindowVisible ());
		return;
	}

	pCmdUI->ContinueRouting();
}
//*************************************************************************************
BOOL CBCGPMDIFrameWnd::OnBarCheck(UINT nID)
{
	ASSERT_VALID (this);

	CBCGPBaseControlBar* pBar = GetControlBar(nID);
	if (pBar != NULL)
	{
		ShowControlBar(pBar, !pBar->IsWindowVisible (), FALSE, FALSE);
		return TRUE;
	}

	return FALSE;
}

//********************************************************************************
LRESULT CBCGPMDIFrameWnd::OnIdleUpdateCmdUI(WPARAM, LPARAM)
{
	if (m_bIsMDIChildDetached)
	{
		CWnd* pFocus = GetFocus();

		if (pFocus->GetSafeHwnd () != NULL && (IsChild (pFocus) || pFocus->GetSafeHwnd () == GetSafeHwnd ()))
		{
			CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, AfxGetMainWnd());
			if (pMainFrame->GetSafeHwnd() != NULL)
			{
				for (POSITION posTlb = gAllToolbars.GetHeadPosition (); posTlb != NULL;)
				{
					CBCGPToolBar* pToolBar = (CBCGPToolBar*) gAllToolbars.GetNext (posTlb);
					ASSERT (pToolBar != NULL);
					
					if (CWnd::FromHandlePermanent (pToolBar->m_hWnd) != NULL && pToolBar->GetOwner()->GetSafeHwnd() == pMainFrame->GetSafeHwnd() &&
						pToolBar->IsWindowVisible())
					{
						ASSERT_VALID(pToolBar);
						pToolBar->OnUpdateCmdUI(this, FALSE);
					}
				}

#ifndef BCGP_EXCLUDE_RIBBON
				if (pMainFrame->m_Impl.m_pRibbonBar->GetSafeHwnd() != NULL && pMainFrame->m_Impl.m_pRibbonBar->IsWindowVisible())
				{
					pMainFrame->m_Impl.m_pRibbonBar->OnUpdateCmdUI(this, FALSE);
				}
#endif
			}
		}
	}

	if (!m_bIsMDIChildDetached && m_hwndLastActiveDetachedMDIFrame != NULL && !m_Impl.IsMenuActive()
#ifndef BCGP_EXCLUDE_RIBBON
		&& m_Impl.m_pBackstageView->GetSafeHwnd() == NULL
#endif
		)
	{
		if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) == 0)
		{
			PostMessage(BCGM_RESET_DETACHED_FRAME);
		}
	}

	m_dockManager.SendMessageToMiniFrames (WM_IDLEUPDATECMDUI);
	return 0L;
}
//********************************************************************************
void CBCGPMDIFrameWnd::OnSize(UINT nType, int cx, int cy) 
{
	if (m_bClosing)
	{
		CMDIFrameWnd::OnSize(nType, cx, cy);
		return;
	}

	m_bIsMinimized = (nType == SIZE_MINIMIZED);

	if (m_Impl.m_pRibbonBar != NULL || m_Impl.IsOwnerDrawCaption ())
	{
		CRect rectWindow;
		GetWindowRect (rectWindow);

		WINDOWPOS wndpos;
		wndpos.flags = SWP_FRAMECHANGED;
		wndpos.x     = rectWindow.left;
		wndpos.y     = rectWindow.top;
		wndpos.cx    = rectWindow.Width ();
		wndpos.cy    = rectWindow.Height ();

		m_Impl.OnWindowPosChanging (&wndpos);
	}

	m_dockManager.OnActivateFrame (!m_bIsMinimized);

	if (!m_bIsMinimized && nType != SIZE_MAXIMIZED && !m_bWasMaximized)
	{
		m_dockManager.m_bSizeFrame = TRUE;
		AdjustDockingLayout ();			
		CMDIFrameWnd::OnSize(nType, cx, cy);
		m_dockManager.m_bSizeFrame = FALSE;
		BOOL bParam = FALSE;
		SystemParametersInfo (SPI_GETDRAGFULLWINDOWS, 0, &bParam, 0);
		if (!bParam)
		{
			RecalcLayout ();
		}

		m_Impl.UpdateCaption ();
		return;
	} 

	CMDIFrameWnd::OnSize(nType, cx, cy);

	if (nType == SIZE_MAXIMIZED || (nType == SIZE_RESTORED && m_bWasMaximized))
	{
		RecalcLayout ();
	}

	m_bWasMaximized = (nType == SIZE_MAXIMIZED);
	m_Impl.UpdateCaption ();
}
//**************************************************************************************
void CBCGPMDIFrameWnd::OnWindowNew()
{
	CMDIChildWnd* pActiveChild = MDIGetActive();
	if (pActiveChild == NULL)
	{
		ASSERT (FALSE);
		return;
	}

	ASSERT_VALID (pActiveChild);

	BOOL bIsZoomed = FALSE;

	if (pActiveChild->IsZoomed ())
	{
		pActiveChild->ShowWindow (SW_RESTORE);
		bIsZoomed = TRUE;
	}

	CMDIFrameWnd::OnWindowNew();

	pActiveChild->RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE);

	if (bIsZoomed)
	{
		pActiveChild = MDIGetActive();
		if (pActiveChild != NULL)
		{
			pActiveChild->ShowWindow (SW_MAXIMIZE);
		}
	}
}
//**************************************************************************************
void CBCGPMDIFrameWnd::OnGetMinMaxInfo(MINMAXINFO FAR* lpMMI) 
{
	if (IsFullScreen ())
	{
		m_Impl.m_FullScreenMgr.OnGetMinMaxInfo (lpMMI);
	}
	else
	{
		m_Impl.OnGetMinMaxInfo (lpMMI);
		CMDIFrameWnd::OnGetMinMaxInfo(lpMMI);
	}
}
//**************************************************************************************
CBCGPMDIChildWnd* CBCGPMDIFrameWnd::CreateDocumentWindow (LPCTSTR lpcszDocName, CObject* pObj)
{
	if (m_bIsMDIChildDetached)
	{
		CBCGPMDIChildWnd* pDocWindow = NULL;

		CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, AfxGetMainWnd());
		if (pMainFrame->GetSafeHwnd() != NULL)
		{
			pDocWindow = pMainFrame->CreateDocumentWindow(lpcszDocName, pObj);
			if (pDocWindow->GetSafeHwnd() != NULL)
			{
				AttachMDIChild(pDocWindow);
			}
		}

		return pDocWindow;
	}

	ASSERT (FALSE);
	TRACE0("If you use save/load state for MDI tabs, you must override this method in a derived class!\n");
	return NULL;
}
//**************************************************************************************
CBCGPMDIChildWnd* CBCGPMDIFrameWnd::CreateNewWindow (LPCTSTR lpcszDocName, CObject* /*pObj*/)
{
	TRACE0("If you use save/load state for MDI tabs, you should override this method in a derived class!\n");

	if (AreMDITabs ())
	{
		OnWindowNew ();
		return DYNAMIC_DOWNCAST (CBCGPMDIChildWnd, MDIGetActive ());
	}

	CDocument* pDoc = AfxGetApp()->OpenDocumentFile (lpcszDocName);
	if (pDoc == NULL)
	{
		return NULL;
	}

	POSITION pos = pDoc->GetFirstViewPosition ();
	if (pos == NULL)
	{
		ASSERT (FALSE);
		return NULL;
	}

	CView* pView = pDoc->GetNextView (pos);
	ASSERT_VALID (pView);

	return DYNAMIC_DOWNCAST (CBCGPMDIChildWnd, pView->GetParentFrame ());
}
//**************************************************************************************
BOOL CBCGPMDIFrameWnd::LoadMDIState (LPCTSTR lpszProfileName)
{
	if (!m_wndClientArea.LoadState (lpszProfileName, m_nFrameID))
	{
		return FALSE;
	}

	if (AreMDIChildrenTearOff())
	{
		for (int nIndex = 1;; nIndex++)
		{
			CString strKey;
			strKey.Format (CBCGPMainClientAreaWnd::m_strRegSectionFmt, lpszProfileName, nIndex);
			
			CBCGPRegistrySP regSP;
			CBCGPRegistry& reg = regSP.Create (FALSE, TRUE);

			if (!reg.VerifyKey(strKey))
			{
				break;
			}

			CBCGPMDIFrameWnd* pNewFrame = DetachMDIChild(NULL, CRect(0, 0, 0, 0));
			if (pNewFrame->GetSafeHwnd() != NULL)
			{
				pNewFrame->m_wndClientArea.LoadState(lpszProfileName, nIndex, TRUE);

				// If no windows were restored, destroy the frame:
				const CObList& lstTabGroups = pNewFrame->m_wndClientArea.GetMDITabGroups();
				BOOL bHasMDITab = FALSE;
				
				for (POSITION pos = lstTabGroups.GetHeadPosition(); pos != NULL;)
				{
					CBCGPTabWnd* pTab = DYNAMIC_DOWNCAST(CBCGPTabWnd, lstTabGroups.GetNext(pos));
					if (pTab != NULL && pTab->GetTabsNum() > 0)
					{
						bHasMDITab = TRUE;
						break;
					}
				}
				
				if (!bHasMDITab)
				{
					pNewFrame->SendMessage(WM_CLOSE);
				}
			}
		}
	}

	return TRUE;
}	
//**************************************************************************************
BOOL CBCGPMDIFrameWnd::SaveMDIState (LPCTSTR lpszProfileName)
{
	if (m_bIsMDIChildDetached)
	{
		return TRUE;
	}

	if (!m_wndClientArea.SaveState (lpszProfileName, m_nFrameID))
	{
		return FALSE;
	}

	int nIndex = 1;
	
	//------------------------------------------------
	// First, clear old detached MDIs in the registry:
	//------------------------------------------------
	for (nIndex = 1;; nIndex++)
	{
		CString strKey;
		strKey.Format (CBCGPMainClientAreaWnd::m_strRegSectionFmt, lpszProfileName, nIndex);
		
		CBCGPRegistrySP regSP;
		CBCGPRegistry& reg = regSP.Create (FALSE, FALSE);
		
		if (!reg.DeleteKey (strKey))
		{
			break;
		}
	}
	
	if (AreMDIChildrenTearOff())
	{
		nIndex = 1;
		
		const CList<CFrameWnd*, CFrameWnd*>& lstFrames = CBCGPFrameImpl::GetFrameList();
		for (POSITION pos = lstFrames.GetHeadPosition(); pos != NULL;)
		{
			CBCGPMDIFrameWnd* pMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, lstFrames.GetNext(pos));
			
			if (pMDIFrame->GetSafeHwnd() != NULL && pMDIFrame->m_bIsMDIChildDetached)
			{
				pMDIFrame->m_wndClientArea.SaveState(lpszProfileName, nIndex++, TRUE);
			}
		}
	}

	return TRUE;
}
//**************************************************************************************
void CBCGPMDIFrameWnd::OnContextMenu(CWnd* pWnd, CPoint point) 
{
	if (m_wndClientArea.GetMDITabs ().GetSafeHwnd () == NULL)
	{
		Default ();
		return;
	}

	if (CBCGPPopupMenu::GetActiveMenu () != NULL)
	{
		return;
	}

	if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)	// Left mouse button is pressed
	{
		return;
	}

	if (pWnd->GetSafeHwnd () == m_wndClientArea.GetSafeHwnd ())
	{
		CWnd* pWndCurr = WindowFromPoint (point);

		if (IsMemberOfMDITabGroup (pWndCurr))
		{
			CBCGPTabWnd* pWndTab = DYNAMIC_DOWNCAST (CBCGPTabWnd, pWndCurr);
			if (pWndTab != NULL)
			{
				CPoint ptTab = point;
				pWndTab->ScreenToClient (&ptTab);

				const int nTab = pWndTab->GetTabFromPoint (ptTab);

				if (nTab >= 0)
				{
					if (pWndTab->GetSelectedTabsCount() == 0)
					{
						pWndTab->SetActiveTab (nTab);
					}

					CRect rectClose;
					rectClose.SetRectEmpty();

					if (!pWndTab->GetTabCloseButtonRect(nTab, rectClose) || !rectClose.PtInRect(ptTab))
					{
						if (IsTearOff())
						{
							CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, AfxGetMainWnd());
							if (pMainFrame->GetSafeHwnd() != NULL)
							{
								pMainFrame->OnShowMDITabContextMenu (point, 0, FALSE);
							}
						}
						else
						{
							OnShowMDITabContextMenu (point, GetMDITabsContextMenuAllowedItems (), FALSE);
						}
					}
				}
			}
		}
		else if (CBCGPPopupMenu::GetActiveMenu () == NULL)
		{
			if (SendMessage (BCGM_TOOLBARMENU,
						(WPARAM) GetSafeHwnd (),
						MAKELPARAM (point.x, point.y)))
			{
				m_dockManager.OnControlBarContextMenu (point);
			}
		}
	}
	else if (pWnd->GetSafeHwnd () == m_wndClientArea.GetMDITabs ().GetSafeHwnd ())
	{
		CBCGPTabWnd& wndTab = (CBCGPTabWnd&) (*pWnd);
		
		CRect rectTabs;
		wndTab.GetTabsRect (rectTabs);

		CPoint ptTab = point;
		wndTab.ScreenToClient (&ptTab);

		const int nTab = wndTab.GetTabFromPoint (ptTab);

		if (nTab >= 0)
		{
			wndTab.SetActiveTab (nTab);

			CRect rectClose; rectClose.SetRectEmpty();

			if (!wndTab.GetTabCloseButtonRect(nTab, rectClose) || !rectClose.PtInRect(ptTab))
			{
				OnShowMDITabContextMenu (point, GetMDITabsContextMenuAllowedItems(), FALSE);
			}
		}
	}
	else
	{
		Default ();
	}
}
//**************************************************************************************
BOOL CBCGPMDIFrameWnd::OnShowPopupMenu (CBCGPPopupMenu* pMenuPopup)
{
	if (globalData.IsAccessibilitySupport () && pMenuPopup != NULL)
	{
		globalData.NotifyWinEvent(EVENT_SYSTEM_MENUPOPUPSTART,
			pMenuPopup->GetSafeHwnd (), OBJID_WINDOW , CHILDID_SELF);
	}

	return TRUE;
}
//**************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnToolbarContextMenu(WPARAM,LPARAM)
{
	return 1l;
}
//**************************************************************************************
BOOL CBCGPMDIFrameWnd::OnCmdMsg(UINT nID, int nCode, void* pExtra, AFX_CMDHANDLERINFO* pHandlerInfo) 
{
	BOOL bRes = FALSE;

	CBCGPMDIFrameWnd* pActiveMDIFrame = GetActiveTearOffFrame();
	if (pActiveMDIFrame != NULL && pActiveMDIFrame->GetSafeHwnd() != GetSafeHwnd())
	{
		bRes = pActiveMDIFrame->OnCmdMsg(nID, nCode, pExtra, pHandlerInfo);
	}
	else
	{
		bRes = CMDIFrameWnd::OnCmdMsg(nID, nCode, pExtra, pHandlerInfo);
	}

	if (bRes)
	{
		return TRUE;
	}

	return m_dockManager.ProcessControlBarContextMenuCommand (nID, nCode, pExtra, pHandlerInfo);
}
//**************************************************************************************
void CBCGPMDIFrameWnd::OnNcPaint() 
{
	if (!m_Impl.OnNcPaint ())
	{
		Default ();
	}
}
//*****************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnSetText(WPARAM, LPARAM lParam)
{
	LRESULT	lRes = Default();

	m_Impl.OnSetText ((LPCTSTR)lParam);
	return lRes;
}
//*****************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnPostMDITearOffFrame(WPARAM wp, LPARAM)
{
	CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, AfxGetMainWnd());
	if (pMainFrame->GetSafeHwnd() != NULL)
	{
		pMainFrame->SendMessage(BCGM_ON_ACTIVATE_MDI_TEAR_OFF_FRAME, wp, (LPARAM)this);
	}
	
	return 0;
}
//*****************************************************************************************
BOOL CBCGPMDIFrameWnd::OnNcActivate(BOOL bActive) 
{
	if (bActive)
	{
		HWND hwndActiveDetachedMDIFramePrev = m_hwndActiveDetachedMDIFrame;

		m_hwndActiveDetachedMDIFrame = m_bIsMDIChildDetached ? GetSafeHwnd() : NULL;

		if (hwndActiveDetachedMDIFramePrev != m_hwndActiveDetachedMDIFrame && m_hwndActiveDetachedMDIFrame != NULL)
		{
			PostMessage(BCGM_POST_ACTIVATE_MDI_TEAR_OFF_FRAME, TRUE);
		}

		CBCGPMDIFrameWnd* pPrev = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, CWnd::FromHandlePermanent(hwndActiveDetachedMDIFramePrev));
		if (pPrev != NULL && hwndActiveDetachedMDIFramePrev != GetSafeHwnd())
		{
			pPrev->m_Impl.CloseLightBox();
		}

	}
	else if (m_hwndActiveDetachedMDIFrame == GetSafeHwnd())
	{
		CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, AfxGetMainWnd());
		if (pMainFrame->GetSafeHwnd() != NULL)
		{
			pMainFrame->SendMessage(BCGM_ON_ACTIVATE_MDI_TEAR_OFF_FRAME, FALSE, (LPARAM)this);
		}
	}

	if (m_Impl.OnNcActivate (bActive))
	{
		return TRUE;
	}

	return CMDIFrameWnd::OnNcActivate(bActive);
}
//**************************************************************************
void CBCGPMDIFrameWnd::OnNcMouseMove(UINT nHitTest, CPoint point) 
{
	m_Impl.OnNcMouseMove (nHitTest, point);
	CMDIFrameWnd::OnNcMouseMove(nHitTest, point);
}
//**************************************************************************
BCGNcHitTestType CBCGPMDIFrameWnd::OnNcHitTest(CPoint point) 
{
	UINT nHit = m_Impl.OnNcHitTest (point);
	if (nHit != HTNOWHERE)
	{
		return nHit;
	}

	return CMDIFrameWnd::OnNcHitTest(point);
}
//**************************************************************************
LRESULT CBCGPMDIFrameWnd::OnChangeVisualManager (WPARAM, LPARAM)
{
	m_dockManager.SendMessageToMiniFrames(BCGM_CHANGEVISUALMANAGER);
	m_dockManager.SendMessageToControlBars(BCGM_CHANGEVISUALMANAGER);

	m_Impl.OnChangeVisualManager();
	return 0;
}
//***************************************************************************
LRESULT CBCGPMDIFrameWnd::OnAfterTaskbarActivate(WPARAM wp, LPARAM)
{
	if (globalData.bIsWindows7 && g_pWorkspace != NULL && g_pWorkspace->IsTaskBarInteractionEnabled())
	{
		BOOL bIsMinimized = (BOOL) wp;

		if (bIsMinimized)
		{
			AdjustDockingLayout ();
			RecalcLayout ();
		}

		SetWindowPos (NULL, -1, -1, -1, -1, 
				SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);

		RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN | RDW_ERASE);

		m_dockManager.RedrawAllMiniFrames();
	}

	return 0;
}
//***************************************************************************
void CBCGPMDIFrameWnd::OnNcCalcSize(BOOL bCalcValidRects, NCCALCSIZE_PARAMS FAR* lpncsp) 
{
	if (!m_Impl.OnNcCalcSize (bCalcValidRects, lpncsp))
	{
		CMDIFrameWnd::OnNcCalcSize(bCalcValidRects, lpncsp);
	}
}
//***************************************************************************
void CBCGPMDIFrameWnd::OnLButtonUp(UINT nFlags, CPoint point) 
{
	m_Impl.OnLButtonUp (point);
	CMDIFrameWnd::OnLButtonUp(nFlags, point);
}
//***************************************************************************
void CBCGPMDIFrameWnd::OnMouseMove(UINT nFlags, CPoint point) 
{
	m_Impl.OnMouseMove (point);
	CMDIFrameWnd::OnMouseMove(nFlags, point);
}
//***************************************************************************
void CBCGPMDIFrameWnd::OnLButtonDown(UINT nFlags, CPoint point) 
{
	m_Impl.OnLButtonDown (point);
	CMDIFrameWnd::OnLButtonDown(nFlags, point);
}
//***************************************************************************
LRESULT CBCGPMDIFrameWnd::OnPostPreviewFrame(WPARAM, LPARAM)
{
	return 0;
}
//***************************************************************************
CBCGPMDIChildWnd* CBCGPMDIFrameWnd::ControlBarToTabbedDocument (CBCGPDockingControlBar* pBar)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pBar);

	CBCGPMDIChildWnd* pFrame = pBar->OnNewMDITabbedChildWnd(this);
	ASSERT_VALID(pFrame);

	CString strName;
	pBar->GetWindowText (strName);

	DWORD dwStyle = WS_CHILD | WS_VISIBLE | FWS_ADDTOTITLE;
	if (!IsMDITabbedGroup())
	{
		dwStyle |= WS_OVERLAPPEDWINDOW;
	}

	if (!pFrame->Create(AfxRegisterWndClass(CS_DBLCLKS, 0, (HBRUSH)(COLOR_BTNFACE + 1), pBar->GetIcon (FALSE)), strName, dwStyle, rectDefault, this))
	{
		return NULL;
	}

	pFrame->SetTitle (strName);
	pFrame->SetWindowText (strName);
	pFrame->AddTabbedControlBar (pBar);

	return pFrame;
}
//***************************************************************************
BOOL CBCGPMDIFrameWnd::TabbedDocumentToControlBar (CBCGPMDIChildWnd* pMDIChildWnd)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pMDIChildWnd);

	if (!pMDIChildWnd->IsTabbedControlBar ())
	{
		ASSERT (FALSE);
		return FALSE;
	}

	CBCGPDockingControlBar* pBar = DYNAMIC_DOWNCAST (CBCGPDockingControlBar, pMDIChildWnd->GetTabbedControlBar ());
	if (pBar != NULL)
	{
		pBar->ShowWindow (SW_HIDE);
		pBar->SetParent (this);
		pBar->SetMDITabbed (FALSE);
		pBar->DockToResentPos ();
	}

	pMDIChildWnd->SendMessage (WM_CLOSE);
	

	return TRUE;
}
//***************************************************************************
void CBCGPMDIFrameWnd::UpdateMDITabbedBarsIcons ()
{
	ASSERT_VALID (this);

	//-----------------------------------
	// Set MDI tabbed control bars icons:
	//-----------------------------------
	HWND hwndMDIChild = ::GetWindow (m_hWndMDIClient, GW_CHILD);

	while (hwndMDIChild != NULL)
	{
		CBCGPMDIChildWnd* pMDIChildFrame = DYNAMIC_DOWNCAST (CBCGPMDIChildWnd, 
			CWnd::FromHandle (hwndMDIChild));

		if (pMDIChildFrame != NULL && pMDIChildFrame->IsTabbedControlBar ())
		{
			CBCGPDockingControlBar* pBar = pMDIChildFrame->GetTabbedControlBar ();
			ASSERT_VALID (pBar);

			SetClassLongPtr(hwndMDIChild, GCLP_HICONSM, (LONG_PTR)pBar->GetIcon (FALSE));
		}

		hwndMDIChild = ::GetWindow (hwndMDIChild, GW_HWNDNEXT);
	}
}
//***************************************************************************
BOOL CBCGPMDIFrameWnd::OnShowMDITabContextMenu (CPoint point, DWORD dwAllowedItems, BOOL /*bTabDrop*/)
{
	if ((dwAllowedItems & BCGP_MDI_CAN_BE_DOCKED) == 0)
	{
		return FALSE;
	}

	if (g_pContextMenuManager == NULL)
	{
		return FALSE;
	}

	const UINT idTabbed	= (UINT) -106;

	CMenu menu;
	menu.CreatePopupMenu ();

	{
		CBCGPLocalResource locaRes;

		CString strItem;

		strItem.LoadString (IDS_BCGBARRES_TABBED);
		menu.AppendMenu (MF_STRING, idTabbed, strItem);

		menu.CheckMenuItem (idTabbed, MF_CHECKED);
	}

	HWND hwndThis = GetSafeHwnd ();

	int nMenuResult = g_pContextMenuManager->TrackPopupMenu (
			menu, point.x, point.y, this);

	if (::IsWindow (hwndThis))
	{
		switch (nMenuResult)
		{
		case idTabbed:
			{
				CBCGPMDIChildWnd* pMDIChild = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, MDIGetActive ());
				if (pMDIChild != NULL)
				{
					TabbedDocumentToControlBar (pMDIChild);
				}
			}
		}
	}

	return TRUE;
}
//***************************************************************************
LRESULT CBCGPMDIFrameWnd::OnDWMCompositionChanged(WPARAM,LPARAM)
{
	m_Impl.OnDWMCompositionChanged ();
	return 0;
}
//***************************************************************************
LRESULT CBCGPMDIFrameWnd::OnPowerBroadcast(WPARAM wp, LPARAM)
{
	LRESULT lres = Default ();

	if (wp == PBT_APMRESUMESUSPEND)
	{
		globalData.Resume ();
	}

	return lres;
}
//***************************************************************************
void CBCGPMDIFrameWnd::RegisterAllMDIChildrenWithTaskbar(BOOL bRegister)
{
	ASSERT_VALID(this);

	HWND hwndMDIChild = ::GetWindow(m_hWndMDIClient, GW_CHILD);

	while (hwndMDIChild != NULL)
	{
		CBCGPMDIChildWnd* pMDIChildFrame = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, CWnd::FromHandle(hwndMDIChild));

		if (pMDIChildFrame != NULL)
		{
			if (bRegister)
			{
				// add at the end
				pMDIChildFrame->RegisterTaskbarTab(NULL);
			}
			else
			{
				pMDIChildFrame->UnregisterTaskbarTab(FALSE);
			}
		}

		hwndMDIChild = ::GetWindow(hwndMDIChild, GW_HWNDNEXT);
	}

	if (bRegister)
	{
		// if we are registering we need to set clip for the active child, which is done internally in InvalidateIconicBitmaps
		BOOL bMax = FALSE;
		CBCGPMDIChildWnd* pMDIChildFrame = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, MDIGetActive(&bMax));
		if (pMDIChildFrame != NULL)
		{
			pMDIChildFrame->InvalidateIconicBitmaps();
		}

		m_wndClientArea.UpdateTabs();
	}
	else
	{
		// if we're unregistering we need to reset clip rect on the main frame
		globalData.TaskBar_SetThumbnailClip(GetSafeHwnd(), NULL);
	}
}
//***************************************************************************
int CBCGPMDIFrameWnd::GetRegisteredWithTaskBarMDIChildCount()
{
	ASSERT_VALID(this);

	int nCount = 0;
	HWND hwndMDIChild = ::GetWindow(m_hWndMDIClient, GW_CHILD);

	while (hwndMDIChild != NULL)
	{
		CBCGPMDIChildWnd* pMDIChildFrame = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, CWnd::FromHandle(hwndMDIChild));

		if (pMDIChildFrame != NULL && pMDIChildFrame->IsRegisteredWithTaskbarTabs())
		{
			nCount++;	
		}

		hwndMDIChild = ::GetWindow(hwndMDIChild, GW_HWNDNEXT);
	}

	return nCount;
}
//************************************************************************************
BOOL CBCGPMDIFrameWnd::SetTaskBarProgressValue(int nCompleted, int nTotal)
{
	return globalData.TaskBar_SetProgressValue(GetSafeHwnd(), nCompleted, nTotal);
}
//************************************************************************************
BOOL CBCGPMDIFrameWnd::SetTaskBarProgressState(BCGP_TBPFLAG tbpFlags)
{
	return globalData.TaskBar_SetProgressState(GetSafeHwnd(), tbpFlags);
}
//************************************************************************************
BOOL CBCGPMDIFrameWnd::SetTaskBarOverlayIcon(UINT nIDResource, LPCTSTR lpcszDescr)
{
	return globalData.TaskBar_SetOverlayIcon(GetSafeHwnd(), nIDResource, lpcszDescr);
}
//************************************************************************************
BOOL CBCGPMDIFrameWnd::SetTaskBarOverlayIcon(HICON hIcon, LPCTSTR lpcszDescr)
{
	return globalData.TaskBar_SetOverlayIcon(GetSafeHwnd(), hIcon, lpcszDescr);
}
//************************************************************************************
void CBCGPMDIFrameWnd::ClearTaskBarOverlayIcon()
{
	globalData.TaskBar_SetOverlayIcon(GetSafeHwnd(), (HICON)NULL, NULL);
}
//************************************************************************************
BOOL CBCGPMDIFrameWnd::CloseRibbonBackstageView()
{
#ifndef BCGP_EXCLUDE_RIBBON
	if (m_Impl.m_pBackstageView->GetSafeHwnd() == NULL)
	{
		return FALSE;
	}

	m_Impl.m_pBackstageView->PostMessage(WM_CLOSE);
#endif
	return TRUE;
}
//************************************************************************************
void CBCGPMDIFrameWnd::OnCancelWndCapture(CWnd* pWndCapture)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pWndCapture);

	pWndCapture->SendMessage(WM_CANCELMODE);
}
//************************************************************************************
void CBCGPMDIFrameWnd::OnWindowPosChanging(WINDOWPOS FAR* lpwndpos)
{
	if (m_Impl.m_bInWindowPosChanging)
	{
		return;
	}

	m_Impl.m_bInWindowPosChanging = TRUE;

	CMDIFrameWnd::OnWindowPosChanging(lpwndpos);
	
	if (IsZoomed())
	{
		CRect rectMonitor(m_Impl.m_rectMonitor);

		if (m_Impl.AdjustMaximizedSize(lpwndpos->cx, lpwndpos->cy))
		{
			// Maybe the monitor changed
			if (lpwndpos->cx == 0 && lpwndpos->cy == 0 && 
				!rectMonitor.IsRectEmpty() && rectMonitor != m_Impl.m_rectMonitor)
			{
				SetWindowPos(NULL, 0, 0, m_Impl.m_rectMonitor.Width(), m_Impl.m_rectMonitor.Height(), SWP_NOZORDER | SWP_NOREPOSITION | SWP_NOOWNERZORDER | SWP_NOMOVE | SWP_FRAMECHANGED);
			}
		}
	}

	m_Impl.m_bInWindowPosChanging = FALSE;
}
//*******************************************************************************************
void CBCGPMDIFrameWnd::OnStyleChanging(int nStyleType, LPSTYLESTRUCT lpStyleStruct)
{
	BOOL bChanged = m_Impl.OnStyleChanging(nStyleType, lpStyleStruct);

	CMDIFrameWnd::OnStyleChanging(nStyleType, lpStyleStruct);

	if (bChanged)
	{
		PostMessage(BCGM_CHANGEVISUALMANAGER, 0, 0);
	}
}
//************************************************************************************
void CBCGPMDIFrameWnd::OnSysCommand(UINT nID, LPARAM lParam)
{
	m_Impl.OnSysCommand(nID, lParam);
	CMDIFrameWnd::OnSysCommand(nID, lParam);

	if ((nID & 0xFFF0) == SC_RESTORE && GetLightBoxDialog()->GetSafeHwnd() != NULL)
	{
		GetLightBoxDialog()->Activate();
	}
}
//************************************************************************************
void CBCGPMDIFrameWnd::EnableWindowsNavigator(BOOL bEnable)
{
	m_Impl.EnableWindowsNavigator(bEnable);
}
//************************************************************************************
void CBCGPMDIFrameWnd::EnableWindowsNavigator(UINT nNextCmdID, UINT nPrevCmdID)
{
	m_Impl.EnableWindowsNavigator(nNextCmdID, nPrevCmdID);
}
//************************************************************************************
CBCGPMDIFrameWnd* CBCGPMDIFrameWnd::DetachMDIChild(CBCGPMDIChildWnd* pMDIChildWnd, CRect rectNewFrame)
{
	ASSERT_VALID(this);

	if (!IsMDITabbedGroup())
	{
		return NULL;
	}

	CRect rectScreen;

	MONITORINFO mi;
	mi.cbSize = sizeof (MONITORINFO);
	if (GetMonitorInfo (MonitorFromRect(rectNewFrame, MONITOR_DEFAULTTONEAREST), &mi))
	{
		rectScreen = mi.rcWork;
	}
	else
	{
		::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectScreen, 0);
	}

	const int cyNewFrame = rectNewFrame.Height();

	if (rectNewFrame.top < rectScreen.top)
	{
		rectNewFrame.top = rectScreen.top;
		rectNewFrame.bottom = rectNewFrame.top + cyNewFrame;
	}
	else if (rectNewFrame.top > rectScreen.bottom - GetSystemMetrics(SM_CYCAPTION))
	{
		rectNewFrame.top -= GetSystemMetrics(SM_CYCAPTION);
		rectNewFrame.bottom = rectNewFrame.top + cyNewFrame;
	}
	
	CDocument* pDoc = pMDIChildWnd != NULL ? pMDIChildWnd->GetActiveDocument() : NULL;
	CView* pView = pMDIChildWnd != NULL ? pMDIChildWnd->GetActiveView() : NULL;
	
	CCreateContext context;
	context.m_pCurrentFrame = this;
	context.m_pCurrentDoc = pDoc;
	context.m_pNewViewClass = NULL;
	context.m_pLastView = pView;
	context.m_pNewDocTemplate = NULL;

	if (pMDIChildWnd->GetSafeHwnd() != NULL)
	{
		//----------------------------------
		// Detach MDI child from this frame:
		//----------------------------------
		m_wndClientArea.OnMDIDestroy((WPARAM)pMDIChildWnd->GetSafeHwnd(), 0);
		pMDIChildWnd->UnregisterTaskbarTab();
		pMDIChildWnd->m_bToBeDestroyed = FALSE;
	}

	//----------------------
	// Create new MDI frame:
	//----------------------
	CBCGPMDIFrameWnd* pNewMDIFrame = NULL;
	
	if (m_pMDITearOffFrameRTI != NULL)
	{
		pNewMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, m_pMDITearOffFrameRTI->CreateObject());
		if (pNewMDIFrame == NULL)
		{
			ASSERT(FALSE);
			return NULL;
		}
	}
	else
	{
		pNewMDIFrame = new CBCGPMDIFrameWnd;
	}

	pNewMDIFrame->m_bIsMDIChildDetached = TRUE;
	pNewMDIFrame->m_Impl.m_bIsTearOffFrame = TRUE;

	CString strTitle = GetTitle();

	if (rectNewFrame.IsRectEmpty() && pMDIChildWnd->GetSafeHwnd() != NULL)
	{
		pMDIChildWnd->GetWindowRect(rectNewFrame);
	}
	
	if (!pNewMDIFrame->Create(globalData.RegisterWindowClass(_T("BCGPDetachedMDIFrame")), strTitle, FWS_ADDTOTITLE | WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, rectNewFrame,
		NULL, 0, 0L, &context))
	{
		delete pNewMDIFrame;
		return NULL;
	}

	//----------------------------------------
	// Set the current icon to the new frame:
	//----------------------------------------
	HICON hIcon = GetIcon(FALSE);
	if (hIcon == NULL)
	{
		hIcon = (HICON)(LONG_PTR) GetClassLongPtr (GetSafeHwnd (), GCLP_HICONSM);
	}

	pNewMDIFrame->SetIcon(hIcon, FALSE);

	//-----------------------------------------
	// Copy default keybaord accelerator table:
	//-----------------------------------------
	int nAccelSize = ::CopyAcceleratorTable(m_hAccelTable, NULL, 0);
	if (nAccelSize > 0)
	{
		LPACCEL lpAccel = new ACCEL[nAccelSize];
		ASSERT(lpAccel != NULL);
	
		::CopyAcceleratorTable(m_hAccelTable, lpAccel, nAccelSize);
		pNewMDIFrame->m_hAccelTable = ::CreateAcceleratorTable(lpAccel, nAccelSize);

		delete[] lpAccel;
	}

	if (pMDIChildWnd->GetSafeHwnd() != NULL)
	{
		//-----------------------------------
		// Attach MDI child to the new frame:
		//-----------------------------------
		pMDIChildWnd->SetParent(&pNewMDIFrame->m_wndClientArea);
		pMDIChildWnd->m_pMDIFrame = pNewMDIFrame;

		pNewMDIFrame->m_wndClientArea.SendMessage(WM_MDIACTIVATE, (WPARAM)pMDIChildWnd->GetSafeHwnd(), 0);

		pMDIChildWnd->RegisterTaskbarTab();
	}

	pNewMDIFrame->SendMessage(BCGM_SYNC_FRAME_OPTIONS, 0, (LPARAM) GetSafeHwnd());

	pNewMDIFrame->ShowWindow(SW_SHOW);
	pNewMDIFrame->SetMenu(NULL);

	CleanUpLastMDIChild();
	return pNewMDIFrame;
}
//************************************************************************************
BOOL CBCGPMDIFrameWnd::AttachMDIChild(CBCGPMDIChildWnd* pMDIChildWnd)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pMDIChildWnd);

	if (!IsMDITabbedGroup())
	{
		return FALSE;
	}

	CBCGPMDIFrameWnd* pParentFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, pMDIChildWnd->GetParentFrame());
	if (pParentFrame->GetSafeHwnd() == NULL || pParentFrame->GetSafeHwnd() == GetSafeHwnd())
	{
		return FALSE;
	}

	//----------------------------------------
	// Detach MDI child from the parent frame:
	//----------------------------------------
	pParentFrame->m_wndClientArea.OnMDIDestroy((WPARAM)pMDIChildWnd->GetSafeHwnd(), 0);
	pMDIChildWnd->UnregisterTaskbarTab();
	pMDIChildWnd->m_bToBeDestroyed = FALSE;

	//---------------------------------
	// Attach MDI child to this frame:
	//---------------------------------
	pMDIChildWnd->m_pMDIFrame = this;
	pMDIChildWnd->SetParent(&m_wndClientArea);
	pMDIChildWnd->RegisterTaskbarTab();

	m_wndClientArea.UpdateMDITabbedGroups(TRUE);

	if (m_bIsMDIChildDetached)
	{
		SetMenu(NULL);
	}

	pParentFrame->CleanUpLastMDIChild();
	return TRUE;
}
//************************************************************************************
void CBCGPMDIFrameWnd::CleanUpLastMDIChild()
{
	const CObList& lstTabGroups = m_wndClientArea.GetMDITabGroups();
	BOOL bHasMDITab = FALSE;
	
	for (POSITION pos = lstTabGroups.GetHeadPosition(); pos != NULL;)
	{
		CBCGPTabWnd* pTab = DYNAMIC_DOWNCAST(CBCGPTabWnd, lstTabGroups.GetNext(pos));
		if (pTab != NULL && pTab->GetTabsNum() > 0)
		{
			bHasMDITab = TRUE;
			break;
		}
	}
	
	if (!bHasMDITab)
	{
		CBCGPMDIChildWnd* pDummyChild = new CBCGPMDIChildWnd;
		pDummyChild->Create(globalData.RegisterWindowClass(_T("BCGPMDIDummy")), _T(""), WS_CHILD | WS_VISIBLE, rectDefault, this, NULL);
		
		m_wndClientArea.SendMessage(WM_MDIACTIVATE, (WPARAM)pDummyChild->GetSafeHwnd(), 0);
		pDummyChild->SendMessage(WM_CLOSE);
		
		OnSetMenu(NULL);
	}
}
//************************************************************************************
void CBCGPMDIFrameWnd::RestoreDetachedFrameActivation()
{
	if (m_hwndLastActiveDetachedMDIFrame != NULL && ::IsWindow(m_hwndLastActiveDetachedMDIFrame))
	{
		CWnd* pWndDetachedMDIFrame = CWnd::FromHandlePermanent(m_hwndLastActiveDetachedMDIFrame);
		if (pWndDetachedMDIFrame->GetSafeHwnd() != NULL)
		{
			pWndDetachedMDIFrame->BringWindowToTop();
			pWndDetachedMDIFrame->SetActiveWindow();
		}
		
		m_hwndLastActiveDetachedMDIFrame = NULL;
	}
}
//************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnReactivateDetachedFrame(WPARAM, LPARAM)
{
	RestoreDetachedFrameActivation();
	return 0;
}
//************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnCleanDetachedFrame(WPARAM, LPARAM)
{
	m_hwndLastActiveDetachedMDIFrame = NULL;
	return 0;
}
//************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnSyncFrameOptions(WPARAM, LPARAM lParam)
{
	CBCGPMDIFrameWnd* pFrameSrc = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, CWnd::FromHandlePermanent((HWND)lParam));
	if (pFrameSrc->GetSafeHwnd() == NULL)
	{
		return 0L;
	}

	m_Impl.m_bIsWindowsNavigatorEnabled = pFrameSrc->m_Impl.m_bIsWindowsNavigatorEnabled;
	m_Impl.m_nWindowsNextNavigationCmdID = pFrameSrc->m_Impl.m_nWindowsNextNavigationCmdID;
	m_Impl.m_nWindowsPrevNavigationCmdID = pFrameSrc->m_Impl.m_nWindowsPrevNavigationCmdID;
	
	EnableTearOffMDIChildren(pFrameSrc->AreMDIChildrenTearOff(), pFrameSrc->GetMDITearOffFrameRTI());
	EnableMDITabbedGroups(TRUE, pFrameSrc->GetMDITabbedGroupsParams());
	
	m_wndClientArea.UpdateMDITabbedGroups(TRUE);

	return 0L;
}
//*****************************************************************************************
LRESULT CBCGPMDIFrameWnd::WindowProc(UINT message, WPARAM wParam, LPARAM lParam) 
{
	if (message == BCGM_ON_GET_TAB_TOOLTIP && IsTearOff())
	{
		CBCGPTabToolTipInfo* pInfo = (CBCGPTabToolTipInfo*)lParam;
		ASSERT(pInfo != NULL);
		ASSERT_VALID(pInfo->m_pTabWnd);
		
		if (pInfo->m_pTabWnd->IsMDITab ())
		{
			CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, AfxGetMainWnd());
			if (pMainFrame->GetSafeHwnd() != NULL)
			{
				return pMainFrame->SendMessage(BCGM_ON_GET_TAB_TOOLTIP, wParam, lParam);
			}
		}
	}
	
	return CMDIFrameWnd::WindowProc(message, wParam, lParam);
}
//*****************************************************************************************
void CBCGPMDIFrameWnd::OnEnable(BOOL bEnable) 
{
	CMDIFrameWnd::OnEnable(bEnable);
	
	if (!m_bIsInternalEnable)
	{
		// Enable/disable MDI tear-off frames
		const CList<CFrameWnd*, CFrameWnd*>& lstFrames = CBCGPFrameImpl::GetFrameList();
		for (POSITION pos = lstFrames.GetHeadPosition(); pos != NULL;)
		{
			CBCGPMDIFrameWnd* pMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, lstFrames.GetNext(pos));
			
			if (pMDIFrame->GetSafeHwnd() != NULL && pMDIFrame->GetSafeHwnd() != GetSafeHwnd())
			{
				pMDIFrame->EnableWindow(bEnable);
			}
		}
	}
}
//*****************************************************************************************
void CBCGPMDIFrameWnd::SetRibbonBackstageView (CBCGPRibbonBackstageView* pView)
{
	m_Impl.SetRibbonBackstageView(pView);

	if (AreMDIChildrenTearOff() && !m_bIsMDIChildDetached)
	{
		const CList<CFrameWnd*, CFrameWnd*>& lstFrames = CBCGPFrameImpl::GetFrameList();
		for (POSITION pos = lstFrames.GetHeadPosition(); pos != NULL;)
		{
			CBCGPMDIFrameWnd* pMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, lstFrames.GetNext(pos));
			
			if (pMDIFrame->GetSafeHwnd() != NULL && pMDIFrame->m_bIsMDIChildDetached)
			{
				pMDIFrame->m_bIsInternalEnable = TRUE;
				pMDIFrame->EnableWindow(pView == NULL);
				pMDIFrame->m_bIsInternalEnable = FALSE;
			}
		}
	}
}
//*****************************************************************************************
BOOL CBCGPMDIFrameWnd::LoadDockingLayout(LPCTSTR lpszProfileName/* = _T("DefaultLayout")*/)
{
	if (g_pWorkspace == NULL)
	{
		return FALSE;
	}

	if (IsFullScreen() && !IsPrintPreview())
	{
		m_Impl.InvokeFullScreenCommand();
	}

	return g_pWorkspace->LoadDockingLayout(lpszProfileName, &m_Impl);
}
//*****************************************************************************************
BOOL CBCGPMDIFrameWnd::SaveDockingLayout(LPCTSTR lpszProfileName/* = _T("DefaultLayout")*/)
{
	if (g_pWorkspace == NULL)
	{
		return FALSE;
	}

	return g_pWorkspace->SaveDockingLayout(lpszProfileName, &m_Impl);
}
//*****************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnChangeBackstagePropHighlighting(WPARAM, LPARAM)
{
#ifndef BCGP_EXCLUDE_RIBBON
	if (m_Impl.m_pBackstageView->GetSafeHwnd() != NULL)
	{
		m_Impl.m_pBackstageView->SendMessage(WM_CANCELMODE);
	}
#endif

	return 0;
}
//*****************************************************************************************
int CBCGPMDIFrameWnd::GetSelectedMDIChildren(CList<HWND, HWND>& lstSelected) const
{
	ASSERT_VALID(this);

	lstSelected.RemoveAll();

	HWND hwndActive = MDIGetActive()->GetSafeHwnd();
	if (hwndActive == NULL)
	{
		return 0;
	}

	if (m_wndClientArea.m_bIsMDITabbedGroup && m_wndClientArea.m_mdiTabParams.m_bTabMultipleSelection)
	{
		CBCGPTabWnd* pTabWnd = ((CBCGPMDIFrameWnd*)this)->m_wndClientArea.FindActiveTabWnd();
		if (pTabWnd != NULL)
		{
			ASSERT_VALID(pTabWnd);

			CList<int, int> lstSelectedTabs;
			if (pTabWnd->GetSelectedTabs(lstSelectedTabs) > 0)
			{
				CList<HWND, HWND> lstTabWnds;
				POSITION pos = NULL;
				
				for (pos = lstSelectedTabs.GetHeadPosition(); pos != NULL;)
				{
					int nIndex = lstSelectedTabs.GetNext(pos);
					HWND hWnd = pTabWnd->GetTabWnd(nIndex)->GetSafeHwnd();

					if (hWnd != NULL)
					{
						lstSelected.AddTail(hWnd);
					}
				}
			}
		}
	}

	if (lstSelected.Find(hwndActive) == NULL)
	{
		lstSelected.AddTail(hwndActive);
	}

	return (int)lstSelected.GetCount();
}
//*****************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnDPIChanged(WPARAM, LPARAM)
{
	LRESULT lRes = Default();
	
	OnChangeVisualManager(0, 0);
	return lRes;
}
//*****************************************************************************************
LRESULT CBCGPMDIFrameWnd::OnThemeChanged(WPARAM, LPARAM)
{
	LRESULT lRes = Default();
	
	CBCGPVisualManager::GetInstance()->OnUpdateSystemColors();
	OnChangeVisualManager(0, 0);
	return lRes;
}
//*****************************************************************************************
void CBCGPMDIFrameWnd::OnSettingChange(UINT uFlags, LPCTSTR lpszSection) 
{
	CMDIFrameWnd::OnSettingChange(uFlags, lpszSection);
	globalData.CheckForImmersiveColorSetChange(lpszSection);
}
//*****************************************************************************************
void CBCGPMDIFrameWnd::OnNcLButtonDown(UINT nHitTest, CPoint point) 
{
	HWND hWndThis = GetSafeHwnd();

	CMDIFrameWnd::OnNcLButtonDown(nHitTest, point);

	if (::IsWindow(hWndThis) && GetLightBoxDialog() != NULL)
	{
		GetLightBoxDialog()->Activate();
	}
}
