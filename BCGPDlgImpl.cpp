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
// BCGPDlgImpl.cpp: implementation of the CBCGPDlgImpl class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "bcgcbpro.h"

#ifndef _BCGSUITE_
#include "BCGPPopupMenu.h"
#include "BCGPToolbarMenuButton.h"
#include "BCGPDropDownList.h"
#include "BCGPRibbonButton.h"
#include "BCGPTooltipManager.h"
#include "BCGPSVGImage.h"
#else
#define FindDestBar	FindDestintationToolBar
#endif

#include "BCGPDialog.h"
#include "BCGPPropertySheet.h"
#include "BCGPButton.h"
#include "BCGPPropertyPage.h"
#include "BCGPDlgImpl.h"
#include "BCGPSliderCtrl.h"
#include "BCGPProgressCtrl.h"
#include "BCGPGroup.h"
#include "BCGPStatic.h"
#include "BCGPEdit.h"
#include "BCGPMenuButton.h"
#include "BCGPGridCtrl.h"
#include "BCGPComboBox.h"
#include "BCGPVisualManager.h"
#include "BCGPScrollBar.h"
#include "BCGPSpinButtonCtrl.h"
#include "BCGPFormView.h"
#include "BCGPGlobalUtils.h"
#include "BCGPLocalResource.h"
#include "BCGProRes.h"

#ifndef _BCGSUITE_
#include "BCGPRibbonElementHostCtrl.h"
#include "BCGPRibbonBackstageView.h"
#include "BCGPCaptionButton.h"
#include "BCGPFrameImpl.h"
#include "RegPath.h"
#include "BCGPRegistry.h"
#include "BCGPWorkspace.h"
#include "BCGPPopupDlg.h"

extern CBCGPWorkspace*	g_pWorkspace;
#else
CString AFXGetRegPath(LPCTSTR lpszPostFix, LPCTSTR lpszProfileName = NULL);
#endif

static const CString strDialogsProfile = _T("BCGDialogs");

#define REG_DLG_SECTION_FMT_STR		_T("%sBCGPDialog-%s")

#ifndef SM_CXPADDEDBORDER
	#define SM_CXPADDEDBORDER 92
#endif

#ifndef _BCGSUITE_
#define visualManagerMFC	CBCGPVisualManager::GetInstance ()
#define visualManager		CBCGPVisualManager::GetInstance ()
#else
#define visualManagerMFC	CMFCVisualManager::GetInstance ()
#define visualManager		CBCGPVisualManager::GetInstance ()
#endif

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

HHOOK CBCGPDlgImpl::m_hookMouse = NULL;
CBCGPDlgImpl* CBCGPDlgImpl::m_pMenuDlgImpl = NULL;
CMap<CBCGPDlgImpl::BCGPControlClass,CBCGPDlgImpl::BCGPControlClass,CRuntimeClass*,CRuntimeClass*> CBCGPDlgImpl::m_mapSubclassedCtrlsRTI;
BOOL CBCGPDlgImpl::m_bKeepBackstageLeftPanelHighlight = FALSE;

static const int nTooltipMaxWidth	= 640;
static const UINT nInfoToolTipID	= 101;

UINT BCGM_ONSETCONTROLAERO = ::RegisterWindowMessage (_T("BCGM_ONSETCONTROLAERO"));
UINT BCGM_ONSETCONTROLVMMODE = ::RegisterWindowMessage (_T("BCGM_ONSETCONTROLVMMODE"));
UINT BCGM_ONSETCONTROLBACKSTAGEMODE = ::RegisterWindowMessage (_T("BCGM_ONSETCONTROBACKSTAGEMMODE"));
UINT BCGM_ONCLICKINFOTIP = ::RegisterWindowMessage (_T("BCGM_ONCLICKINFOTIP"));
UINT BCGM_ONDRAWCUSTOMINFOTIP = ::RegisterWindowMessage (_T("BCGM_ONDRAWCUSTOMINFOTIP"));
UINT BCGM_ONCHANGEUNDERLINE = ::RegisterWindowMessage (_T("BCGM_ONCHANGEUNDERLINE"));
UINT BCGM_ONGETGESTURECONFIG = ::RegisterWindowMessage (_T("BCGM_ONGETGESTURECONFIG"));

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPDlgImpl::CBCGPDlgImpl(CWnd& dlg) :
	m_Dlg (dlg),
	m_bVisualManagerStyle (FALSE),
	m_bTransparentStaticCtrls (TRUE),
	m_bVisualManagerNCArea (FALSE),
	m_nHotSysButton (HTNOWHERE),
	m_nHitSysButton (HTNOWHERE),
	m_bWindowPosChanging (FALSE),
	m_bHasBorder (FALSE),
	m_bHasCaption (TRUE),
	m_bIsWindowRgn (FALSE),
	m_bIsWhiteBackground(FALSE),
	m_pLayout(NULL),
	m_bBackstageMode(FALSE),
	m_bResizeBox(FALSE),
	m_bLoadWindowPlacement(FALSE),
	m_bWindowPlacementIsSet(FALSE),
	m_bDragClientArea(FALSE),
	m_pShadow(NULL),
	m_pToolTipInfo(NULL),
	m_hwndInfoTipCurr(NULL),
	m_BkgrLocation(0),
	m_bDisableShadows(FALSE),
	m_bGroupBoxesDrawByParent(TRUE),
	m_bAutoResizeCtrls(FALSE),
	m_nEditBoxesVertAlignment(-1),
	m_bEditBoxesSimplifiedIcon(-1)
{
	m_AeroMargins.cxLeftWidth = 0;
	m_AeroMargins.cxRightWidth = 0;
	m_AeroMargins.cyTopHeight = 0;
	m_AeroMargins.cyBottomHeight = 0;

	m_rectRedraw.SetRectEmpty ();
	m_rectResizeBox.SetRectEmpty ();

	ZeroMemory(&m_LayoutMMI, sizeof(MINMAXINFO));
}
//*******************************************************************************************
CBCGPDlgImpl::~CBCGPDlgImpl()
{
	//------------------------------
	// Clear caption system buttons:
	//------------------------------
	while (!m_lstCaptionSysButtons.IsEmpty ())
	{
		delete m_lstCaptionSysButtons.RemoveHead ();
	}

	if (m_pLayout != NULL)
	{
		delete m_pLayout;
		m_pLayout = NULL;
	}

#ifndef _BCGSUITE_
	if (m_pShadow != NULL)
	{
		delete m_pShadow;
		m_pShadow = NULL;
	}
#endif
}
//*******************************************************************************************
void CBCGPDlgImpl::SetControlInfoTip(UINT nCtrlID, LPCTSTR lpszInfoTip, DWORD dwVertAlign, BOOL bRedrawInfoTip, CBCGPControlInfoTip::BCGPControlInfoTipStyle style, BOOL bIsClickable, const CPoint& ptOffset)
{
	if (nCtrlID == 0 || nCtrlID == IDC_STATIC)
	{
		ASSERT(FALSE);
		return;
	}

	if (m_Dlg.GetSafeHwnd() == NULL)
	{
		CString strInfoTip = (lpszInfoTip == NULL) ? _T("") : lpszInfoTip;
		if (strInfoTip.IsEmpty())
		{
			m_mapCtrlInfoTipDelayed.RemoveKey(nCtrlID);
		}
		else
		{
			m_mapCtrlInfoTipDelayed.SetAt(nCtrlID, CBCGPControlInfoTip(NULL, lpszInfoTip, dwVertAlign, style, bIsClickable));
		}

		return;
	}

	SetControlInfoTip(m_Dlg.GetDlgItem(nCtrlID), lpszInfoTip, dwVertAlign, bRedrawInfoTip, style, bIsClickable, ptOffset);
}
//*******************************************************************************************
void CBCGPDlgImpl::SetControlInfoTip(CWnd* pWndCtrl, LPCTSTR lpszInfoTip, DWORD dwVertAlign, BOOL bRedrawInfoTip, CBCGPControlInfoTip::BCGPControlInfoTipStyle style, BOOL bIsClickable, const CPoint& ptOffset)
{
	HWND hwndCtrl = pWndCtrl->GetSafeHwnd();
	if (hwndCtrl == NULL || !::IsWindow(hwndCtrl))
	{
		ASSERT(FALSE);
		return;
	}

	CString strInfoTip = (lpszInfoTip == NULL) ? _T("") : lpszInfoTip;
	if (strInfoTip.IsEmpty())
	{
		m_mapCtrlInfoTip.RemoveKey(hwndCtrl);
	}
	else
	{
		m_mapCtrlInfoTip.SetAt(hwndCtrl, CBCGPControlInfoTip(pWndCtrl, lpszInfoTip, dwVertAlign, style, bIsClickable, ptOffset));
	}

	if (bRedrawInfoTip && m_Dlg.GetSafeHwnd() != NULL)
	{
		CRect rect;
		GetControlInfoTipRect(pWndCtrl, rect, dwVertAlign, ptOffset);

		m_Dlg.RedrawWindow(rect);
	}
}
//*******************************************************************************************
BOOL CBCGPDlgImpl::ProcessMouseClick (POINT pt)
{
	if (!CBCGPToolBar::IsCustomizeMode () &&
		CBCGPPopupMenu::GetActiveMenu() != NULL &&
		::IsWindow (CBCGPPopupMenu::GetActiveMenu()->m_hWnd))
	{
		CBCGPPopupMenu::MENUAREA_TYPE clickArea = CBCGPPopupMenu::GetActiveMenu()->CheckArea (pt);

		if (clickArea == CBCGPPopupMenu::OUTSIDE)
		{
			// Click outside of menu

			//--------------------------------------------
			// Maybe secondary click on the parent button?
			//--------------------------------------------
			CRect rectParentBtn;
			CWnd* pWndParent = CBCGPPopupMenu::GetActiveMenu()->GetParentArea (rectParentBtn);

			if (pWndParent != NULL)
			{
				CBCGPPopupMenuBar* pWndParentPopupMenuBar = 
					DYNAMIC_DOWNCAST (CBCGPPopupMenuBar, pWndParent);

				CPoint ptClient = pt;
				pWndParent->ScreenToClient (&ptClient);

				if (rectParentBtn.PtInRect (ptClient))
				{
					//-------------------------------------------------------
					// If user clicks second time on the parent button,
					// we should close an active menu on the toolbar/menubar
					// and leave it on the popup menu:
					//-------------------------------------------------------
					if (pWndParentPopupMenuBar == NULL &&
						!CBCGPPopupMenu::GetActiveMenu()->InCommand ())
					{
						//----------------------------------------
						// Toolbar/menu bar: close an active menu!
						//----------------------------------------
						CBCGPPopupMenu::GetActiveMenu()->SendMessage (WM_CLOSE);
					}

					return TRUE;
				}

				if (pWndParentPopupMenuBar != NULL)
				{
					pWndParentPopupMenuBar->CloseDelayedSubMenu ();
					
					CBCGPPopupMenu* pWndParentPopupMenu = 
						DYNAMIC_DOWNCAST (CBCGPPopupMenu, 
						pWndParentPopupMenuBar->GetParent ());

					if (pWndParentPopupMenu != NULL)
					{
						CBCGPPopupMenu::MENUAREA_TYPE clickAreaParent = 
							pWndParentPopupMenu->CheckArea (pt);

						switch (clickAreaParent)
						{
						case CBCGPPopupMenu::MENU:
						case CBCGPPopupMenu::TEAROFF_CAPTION:
						case CBCGPPopupMenu::LOGO:
							return FALSE;

						case CBCGPPopupMenu::SHADOW_RIGHT:
						case CBCGPPopupMenu::SHADOW_BOTTOM:
							pWndParentPopupMenu->SendMessage (WM_CLOSE);
							m_Dlg.SetFocus ();

							return TRUE;
						}
					}
				}
			}

			if (!CBCGPPopupMenu::GetActiveMenu()->InCommand ())
			{
				CBCGPPopupMenu::GetActiveMenu()->SendMessage (WM_CLOSE);

				CWnd* pWndFocus = CWnd::GetFocus ();
				if (pWndFocus != NULL && pWndFocus->IsKindOf (RUNTIME_CLASS (CBCGPToolBar)))
				{
					m_Dlg.SetFocus ();
				}

				if (clickArea != CBCGPPopupMenu::OUTSIDE)	// Click on shadow
				{
					return TRUE;
				}
			}
		}
		else if (clickArea == CBCGPPopupMenu::SHADOW_RIGHT ||
				clickArea == CBCGPPopupMenu::SHADOW_BOTTOM)
		{
			CBCGPPopupMenu::GetActiveMenu()->SendMessage (WM_CLOSE);
			m_Dlg.SetFocus ();

			return TRUE;
		}
	}

	return FALSE;
}
//*******************************************************************************************
BOOL CBCGPDlgImpl::ProcessMouseMove (POINT pt)
{
	if (m_bBackstageMode && m_Dlg.GetParent() != NULL && !m_bKeepBackstageLeftPanelHighlight)
	{
		m_Dlg.GetParent()->SendMessage (WM_MOUSEMOVE, 0, MAKELPARAM (-1, -1));
	}

	if (!CBCGPToolBar::IsCustomizeMode ())
	{
#ifndef _BCGSUITE_
		CBCGPPopupMenu* pActivePopup = CBCGPPopupMenu::GetSafeActivePopupMenu();
#else
		CBCGPPopupMenu* pActivePopup = CBCGPPopupMenu::GetActiveMenu();
#endif
		if (pActivePopup != NULL)
		{
			CRect rectMenu;
			pActivePopup->GetWindowRect (rectMenu);

			if (rectMenu.PtInRect (pt) || pActivePopup->GetMenuBar ()->FindDestBar (pt) != NULL)
			{
				return FALSE;	// Default processing
			}

			return TRUE;		// Active menu "capturing"
		}
	}

	return FALSE;	// Default processing
}
//**************************************************************************************
BOOL CBCGPDlgImpl::PreTranslateMessage(MSG* pMsg) 
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
		if (m_pToolTipInfo->GetSafeHwnd () != NULL)
		{
			m_pToolTipInfo->RelayEvent(pMsg);
		}
		break;
	}
	
	switch (pMsg->message)
	{
	case WM_SYSKEYDOWN:
	case WM_CONTEXTMENU:
		if (CBCGPPopupMenu::GetActiveMenu() != NULL &&
			::IsWindow (CBCGPPopupMenu::GetActiveMenu()->m_hWnd) &&
			pMsg->wParam == VK_MENU)
		{
			CBCGPPopupMenu::GetActiveMenu()->SendMessage (WM_CLOSE);
			return TRUE;
		}

		SetUnderlineKeyboardShortcuts(TRUE, TRUE);
		break;

	case WM_SYSKEYUP:
		if (CBCGPPopupMenu::GetActiveMenu() != NULL &&
			::IsWindow (CBCGPPopupMenu::GetActiveMenu()->m_hWnd))
		{
			return TRUE;	// To prevent system menu opening
		}
		break;

	case WM_KEYDOWN:
		//-----------------------------------------
		// Pass keyboard action to the active menu:
		//-----------------------------------------
		if (!CBCGPToolBar::IsCustomizeMode () &&
			CBCGPPopupMenu::GetActiveMenu() != NULL && ::IsWindow (CBCGPPopupMenu::GetActiveMenu()->m_hWnd) &&
			CBCGPPopupMenu::GetActiveMenu()->IsWindowVisible())
		{
			//-------------------------------------------------------
			// Check if the dialog is not located on the active menu:
			//-------------------------------------------------------
			BOOL bIsParentMenu = m_Dlg.GetParent()->GetSafeHwnd() == CBCGPPopupMenu::GetActiveMenu()->m_hWnd;
			if (!bIsParentMenu)
			{
				BOOL bIsDropList = CBCGPPopupMenu::GetActiveMenu()->GetMenuBar ()->IsDropDownListMode ();

				CBCGPPopupMenu::GetActiveMenu()->SendMessage (WM_KEYDOWN, (int) pMsg->wParam);

				if (!bIsDropList || pMsg->wParam == VK_DOWN || pMsg->wParam == VK_UP)
				{
					return TRUE;
				}
				
#ifndef _BCGSUITE_
				CBCGPDropDownList* pDropDownList = DYNAMIC_DOWNCAST(
					CBCGPDropDownList, CBCGPPopupMenu::GetSafeActivePopupMenu());
#else
				CMFCDropDownListBox* pDropDownList = DYNAMIC_DOWNCAST(
					CMFCDropDownListBox, CMFCPopupMenu::GetActiveMenu());
#endif
				return pDropDownList == NULL || !pDropDownList->IsEditFocused ();
			}

			if (ProcessFocusEditAccelerators((UINT)pMsg->wParam))
			{
				return TRUE;
			}
		}

#ifndef BCGP_EXCLUDE_RIBBON
#ifndef _BCGSUITE_
		if (m_bBackstageMode && (pMsg->wParam == VK_HOME || pMsg->wParam == VK_ESCAPE))
		{
			for (CWnd* pWndParent = m_Dlg.GetParent(); pWndParent != NULL; pWndParent = pWndParent->GetParent())
			{
				CBCGPRibbonBackstageView* pView = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageView, pWndParent);
				if (pView != NULL)
				{
					if (pMsg->wParam == VK_HOME)
					{
						#define MAX_CLASS_NAME		255
						#define EDIT_CLASS			_T("Edit")

						TCHAR lpszClassName [MAX_CLASS_NAME + 1];

						if (CWnd::GetFocus()->GetSafeHwnd() != NULL)
						{
							::GetClassName (CWnd::GetFocus()->GetSafeHwnd (), lpszClassName, MAX_CLASS_NAME);
							CString strClass = lpszClassName;

							if (strClass == EDIT_CLASS)
							{
								return FALSE;
							}
						}

						pView->SetFocus();
					}
					else	// Escape
					{
						pView->SendMessage(WM_CLOSE);
					}
					return TRUE;
				}
			}
			return TRUE;
		}
#endif
#endif
		break;

	case WM_LBUTTONDOWN:
	case WM_RBUTTONDOWN:
	case WM_RBUTTONUP:
	case WM_MBUTTONDOWN:
	case WM_MBUTTONUP:
		{
			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));
			CWnd* pWnd = CWnd::FromHandle (pMsg->hwnd);

			if (pWnd != NULL && ::IsWindow (pMsg->hwnd))
			{
				pWnd->ClientToScreen (&pt);
			}

			if (ProcessMouseClick (pt))
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
		if (ProcessMouseClick (CPoint (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam))))
		{
			return TRUE;
		}

		if (pMsg->message == WM_NCRBUTTONUP && pMsg->hwnd == m_Dlg.GetSafeHwnd () && IsOwnerDrawCaption () && (m_Dlg.GetStyle() & WS_SYSMENU) == WS_SYSMENU)
		{
			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));

			UINT nHit = OnNcHitTest (pt);

			if (nHit == HTCAPTION || nHit == HTSYSMENU)
			{
				CMenu* pMenu = m_Dlg.GetSystemMenu (FALSE);
				if (pMenu->GetSafeHmenu () != NULL && ::IsMenu (pMenu->GetSafeHmenu ()))
				{
					UINT uiRes = ::TrackPopupMenu (pMenu->GetSafeHmenu(), TPM_LEFTBUTTON | TPM_RETURNCMD, 
						pt.x, pt.y, 0, m_Dlg.GetSafeHwnd (), NULL);

					if (uiRes != 0)
					{
						m_Dlg.SendMessage (WM_SYSCOMMAND, uiRes);
						return TRUE;
					}
				}

			}
		}
		break;

	case WM_MOUSEWHEEL:
		{
#ifdef _BCGSUITE_
			if (CBCGPPopupMenu::GetActiveMenu() != NULL &&
				::IsWindow (CBCGPPopupMenu::GetActiveMenu()->m_hWnd) &&
				CBCGPPopupMenu::GetActiveMenu()->IsScrollable ())
			{
				CBCGPPopupMenu::GetActiveMenu()->SendMessage (WM_MOUSEWHEEL,
					pMsg->wParam, pMsg->lParam);
			}
#else
			CBCGPPopupMenu* pActivePopupMenu = CBCGPPopupMenu::GetSafeActivePopupMenu();
			if (pActivePopupMenu != NULL)
			{
				ASSERT_VALID(pActivePopupMenu);

				if (DYNAMIC_DOWNCAST(CBCGPDlgPopupMenu, pActivePopupMenu))
				{
					CWnd* pWnd = CWnd::FromHandle(pMsg->hwnd);
					CPoint pt(BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));

					if (pWnd->GetSafeHwnd() != NULL && CWnd::WindowFromPoint(pt) == pWnd)
					{
						pWnd->SendMessage(WM_MOUSEWHEEL, pMsg->wParam, pMsg->lParam);
					}
				}
				else if (pActivePopupMenu->IsScrollable ())
				{
					pActivePopupMenu->SendMessage (WM_MOUSEWHEEL, pMsg->wParam, pMsg->lParam);
				}

				return TRUE;
			}
			else
			{
				if (!m_Dlg.IsWindowEnabled())
				{
					return FALSE;
				}

				if (g_pWorkspace != NULL && g_pWorkspace->IsMouseWheelInInactiveWindowEnabled() &&
					CBCGPGlobalUtils::ProcessMouseWheel(pMsg->wParam, pMsg->lParam))
				{
					return TRUE;
				}
			}
#endif
		}
		break;

	case WM_MOUSEMOVE:
		{
			CPoint pt (BCG_GET_X_LPARAM(pMsg->lParam), BCG_GET_Y_LPARAM(pMsg->lParam));

			if (m_pToolTipInfo->GetSafeHwnd () != NULL && pMsg->hwnd == m_Dlg.GetSafeHwnd() && !m_mapCtrlInfoTip.IsEmpty())
			{
				CString strDescription;
				CString strTipText;
				BOOL bIsClickable = FALSE;

				HWND hwnd = GetControlInfoTipFromPoint(pt, strTipText, strDescription, bIsClickable);
				
				if (hwnd != m_hwndInfoTipCurr)
				{
					m_pToolTipInfo->Pop();
					m_hwndInfoTipCurr = hwnd;
				}
			}

			CWnd* pWnd = CWnd::FromHandle (pMsg->hwnd);
			if (pWnd != NULL)
			{
				pWnd->ClientToScreen (&pt);
			}

			if (ProcessMouseMove (pt))
			{
				return TRUE;
			}
		}
	}

	return FALSE;
}
//**********************************************************************************
LRESULT CALLBACK CBCGPDlgImpl::BCGDlgMouseProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	if (m_pMenuDlgImpl != NULL)
	{
		switch (wParam)
		{
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
		case WM_MBUTTONDOWN:
		case WM_NCLBUTTONDOWN:
		case WM_NCRBUTTONDOWN:
		case WM_NCMBUTTONDOWN:
			{
				CPoint ptCursor;
				::GetCursorPos (&ptCursor);

				CRect rectWindow;
				m_pMenuDlgImpl->m_Dlg.GetWindowRect (rectWindow);

				if (!rectWindow.PtInRect (ptCursor))
				{
					m_pMenuDlgImpl->ProcessMouseClick (ptCursor);
				}
			}
		}
	}

	return CallNextHookEx (m_hookMouse, nCode, wParam, lParam);
}
//****************************************************************************************
void CBCGPDlgImpl::SetActiveMenu (CBCGPPopupMenu* pMenu)
{
#ifndef _BCGSUITE_
	CBCGPPopupMenu::m_pActivePopupMenu = pMenu;
#else
	class CBCGPPopupMenuDummy : public CBCGPPopupMenu
	{
		friend class CBCGPDlgImpl;
	};

	CBCGPPopupMenuDummy::m_pActivePopupMenu = pMenu;

#endif

	if (pMenu != NULL)
	{
		if (m_hookMouse == NULL)
		{
			m_hookMouse = ::SetWindowsHookEx (WH_MOUSE, BCGDlgMouseProc, 
				0, GetCurrentThreadId ());
		}

		m_pMenuDlgImpl = this;
	}
	else 
	{
		if (m_hookMouse != NULL)
		{
			::UnhookWindowsHookEx (m_hookMouse);
			m_hookMouse = NULL;
		}

		m_pMenuDlgImpl = NULL;
	}
}
//****************************************************************************************
void CBCGPDlgImpl::OnDestroy ()
{
	for (int i = 0; i < m_arSubclassedCtrls.GetSize (); i++)
	{
		delete m_arSubclassedCtrls [i];
	}

	m_arSubclassedCtrls.RemoveAll ();
	if (m_pMenuDlgImpl != NULL &&
		m_pMenuDlgImpl->m_Dlg.GetSafeHwnd () == m_Dlg.GetSafeHwnd ())
	{
		m_pMenuDlgImpl = NULL;
	}

#ifndef _BCGSUITE_
	visualManager->RemoveActivateFlag(&m_Dlg);
#endif
	SavePlacement();

	CBCGPTooltipManager::DeleteToolTip(m_pToolTipInfo);
}
//****************************************************************************************
BOOL CBCGPDlgImpl::OnCommand (WPARAM wParam, LPARAM lParam)
{
	if (HIWORD(wParam) == 1 && lParam == 0)
	{
		UINT uiCmd = LOWORD (wParam);

		CBCGPToolBar::AddCommandUsage (uiCmd);

		//---------------------------
		// Simulate ESC keystroke...
		//---------------------------
		if (CBCGPPopupMenu::GetActiveMenu() != NULL && ::IsWindow (CBCGPPopupMenu::GetActiveMenu()->m_hWnd))
		{
			//-------------------------------------------------------
			// Check if the dialog is not located on the active menu:
			//-------------------------------------------------------
			BOOL bIsParentMenu = m_Dlg.GetParent()->GetSafeHwnd() == CBCGPPopupMenu::GetActiveMenu()->m_hWnd;
			if (!bIsParentMenu)
			{
				CBCGPPopupMenu::GetActiveMenu()->SendMessage (WM_KEYDOWN, VK_ESCAPE);
				return TRUE;
			}
		}

		if (g_pUserToolsManager != NULL && g_pUserToolsManager->InvokeTool (uiCmd))
		{
			return TRUE;
		}
	}

	return FALSE;
}
//***************************************************************************************
void CBCGPDlgImpl::OnNcActivate (BOOL& bActive)
{
	//----------------------------------------
	// Stay active if WF_STAYACTIVE bit is on:
	//----------------------------------------

	// dialog does not have WF_STAYACTIVE flag (this is specific to CFrameWnd only)
	BOOL bStayActive = (m_Dlg.m_nFlags & WF_STAYACTIVE) == WF_STAYACTIVE;
	m_Dlg.m_nFlags &= ~WF_STAYACTIVE;

	//--------------------------------------------------
	// But do not stay active if the window is disabled:
	//--------------------------------------------------
	if (!m_Dlg.IsWindowEnabled ())
	{
		bActive = FALSE;
	}

	if (IsOwnerDrawCaption ())
	{
		visualManagerMFC->OnNcActivate (&m_Dlg, bActive);
		m_Dlg.RedrawWindow (NULL, NULL,
            RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW | (m_Dlg.IsZoomed() ? RDW_ERASE : RDW_NOCHILDREN));
	}

	if (bStayActive)
	{
		m_Dlg.m_nFlags |= WF_STAYACTIVE;
	}

	AdjustShadow(bActive);
}
//****************************************************************************************
void CBCGPDlgImpl::OnActivate(UINT nState, CWnd* pWndOther)
{
	m_Dlg.m_nFlags &= ~WF_STAYACTIVE;

	//--------------------------------------------------
	// Determine if this window should be active or not:
	//--------------------------------------------------
	CWnd* pWndActive = (nState == WA_INACTIVE) ? pWndOther : &m_Dlg;
	if (pWndActive != NULL)
	{
		BOOL bStayActive = (pWndActive->GetSafeHwnd () == m_Dlg.GetSafeHwnd () ||
			pWndActive->SendMessage (WM_FLOATSTATUS, FS_SYNCACTIVE));

		if (bStayActive)
		{
			m_Dlg.m_nFlags |= WF_STAYACTIVE;
		}
	}
	else 
	{
		//------------------------------------------
		// Force painting on our non-client area....
		//------------------------------------------
		m_Dlg.SendMessage (WM_NCPAINT, 1);
	}

	if (nState == WA_INACTIVE && IsOwnerDrawCaption ())
	{
		m_Dlg.RedrawWindow(NULL, NULL, RDW_FRAME | RDW_UPDATENOW);
	}

	AdjustShadow(nState != WA_INACTIVE);

	UpdateToolTipsRect();

	if (nState == WA_INACTIVE)
	{
		SetUnderlineKeyboardShortcuts(FALSE, TRUE);
	}
}
//*************************************************************************************
void CBCGPDlgImpl::SetSubclassControlRTI(BCGPControlClass idClass, CRuntimeClass* pRTI)
{
	if (idClass < BCGPCONTROLSCLASS_First || idClass > BCGPCONTROLSCLASS_Last)
	{
		ASSERT(FALSE);
		return;
	}

	if (pRTI == NULL)
	{
		m_mapSubclassedCtrlsRTI.RemoveKey(idClass);
	}
	else
	{
		m_mapSubclassedCtrlsRTI.SetAt(idClass, pRTI);
	}
}
//*************************************************************************************
CSize CBCGPDlgImpl::LoadInfoTipIcons(CBCGPToolBarImages& infotipIcons)
{
	if (infotipIcons.GetCount() == 0)
	{
		CSize sizeIcon(16, 16);

		CBCGPLocalResource locaRes;
		
#ifndef _BCGSUITE_
#ifndef BCGP_EXCLUDE_SVG
		// Try to load SVG icons:
		CBCGPSVGImageList svgList;
		if (svgList.Load(IDB_BCGBARRES_INFOTIP_ICONS))
		{
			infotipIcons.SetImageSize(globalUtils.ScaleByDPI(sizeIcon));
			infotipIcons.AddSVGList(svgList);
		}
#endif
#endif		
		if (infotipIcons.GetCount() == 0)
		{
			UINT nResID = IDB_BCGBARRES_INFOTIP_ICONS24;
			
			if (globalData.Is32BitIcons())
			{
#ifndef _BCGSUITE_
				const double dblScale = globalData.GetRibbonImageScale();
				
				if (dblScale > 1.51)
				{
					nResID = IDB_BCGBARRES_INFOTIP_ICONS200;
					sizeIcon = CSize(32, 32);
				}
				else if (dblScale > 1.26)
				{
					nResID = IDB_BCGBARRES_INFOTIP_ICONS150;
					sizeIcon = CSize(24, 24);
				}
				else if (dblScale > 1.1)
				{
					nResID = IDB_BCGBARRES_INFOTIP_ICONS125;
					sizeIcon = CSize(20, 20);
				}
				else
#endif
				{
					nResID = IDB_BCGBARRES_INFOTIP_ICONS100;
				}
			}
			
			infotipIcons.SetImageSize(sizeIcon);
			infotipIcons.SetTransparentColor(RGB(255, 0, 255));
			
			if (!infotipIcons.Load(nResID))
			{
				TRACE0("CBCGPDlgImpl::DrawControlInfoTips: cannot load info tip icons!\n");
				return CSize(0, 0);
			}
		}
	}

	return infotipIcons.GetImageSize();
}
//*************************************************************************************
CWnd* CBCGPDlgImpl::CreateSubclassedCtrl(BCGPControlClass idClass)
{
	CRuntimeClass* pRTI = m_mapSubclassedCtrlsRTI[idClass];
	if (pRTI == NULL)
	{
		switch (idClass)
		{
		case BCGPCONTROLSCLASS_Button:
			return new CBCGPButton;
			
		case BCGPCONTROLSCLASS_GroupBox:
			{
				CBCGPGroup* pGroup = new CBCGPGroup;
				pGroup->m_bDrawByParent = m_bGroupBoxesDrawByParent;

				return pGroup;
			}
			
		case BCGPCONTROLSCLASS_Progress:
			return new CBCGPProgressCtrl;
			
		case BCGPCONTROLSCLASS_TrackBar:
			return new CBCGPSliderCtrl;
			
		case BCGPCONTROLSCLASS_Edit:
			return new CBCGPEdit;
			
		case BCGPCONTROLSCLASS_ComboBox:
			return new CBCGPComboBox;
			
		case BCGPCONTROLSCLASS_ScrollBar:
			return new CBCGPScrollBar;
			
		case BCGPCONTROLSCLASS_UpDown:
			return new CBCGPSpinButtonCtrl;
			
		case BCGPCONTROLSCLASS_Static:
			return new CBCGPStatic;
		}

		ASSERT(FALSE);
		return NULL;
	}

	return (CWnd*)pRTI->CreateObject();
}
//*************************************************************************************
void CBCGPDlgImpl::EnableVisualManagerStyle (BOOL bEnable, BOOL bNCArea, const CList<UINT,UINT>* plstNonSubclassedItems)
{
	m_bVisualManagerStyle = bEnable;
	m_bVisualManagerNCArea = bNCArea;

	if (m_lstNonSubclassedItems.IsEmpty() && plstNonSubclassedItems != NULL)
	{
		m_lstNonSubclassedItems.AddTail((CList<UINT,UINT>*)plstNonSubclassedItems);
	}

	if (m_Dlg.GetSafeHwnd () == NULL)
	{
		return;
	}

	CWnd* pWndChild = m_Dlg.GetWindow (GW_CHILD);

	while (pWndChild != NULL)
	{
		ASSERT_VALID (pWndChild);

		if (plstNonSubclassedItems != NULL && plstNonSubclassedItems->Find (pWndChild->GetDlgCtrlID ()) != NULL)
		{
			pWndChild = pWndChild->GetNextWindow ();
			continue;
		}

		if (m_lstNonSubclassedItems.Find (pWndChild->GetDlgCtrlID ()) != NULL)
		{
			pWndChild = pWndChild->GetNextWindow ();
			continue;
		}

		CBCGPButton* pButton = DYNAMIC_DOWNCAST(CBCGPButton, pWndChild);
		if (pButton != NULL)
		{
			ASSERT_VALID (pButton);
			pButton->m_bVisualManagerStyle = m_bVisualManagerStyle;
		}

		if (m_bVisualManagerStyle &&
			CWnd::FromHandlePermanent (pWndChild->GetSafeHwnd ()) == NULL)
		{
			#define MAX_CLASS_NAME		255
			#define STATIC_CLASS		_T("Static")
			#define BUTTON_CLASS		_T("Button")
			#define EDIT_CLASS			_T("Edit")
			#define	COMBOBOX_CLASS		_T("ComboBox")
			#define SCROLLBAR_CLASS		_T("ScrollBar")

			TCHAR lpszClassName [MAX_CLASS_NAME + 1];

			::GetClassName (pWndChild->GetSafeHwnd (), lpszClassName, MAX_CLASS_NAME);
			CString strClass = lpszClassName;

			CWnd* pWndSubclassedCtrl = NULL;

			if (m_lstNonSubclassedWndClasses.Find(strClass) != NULL)
			{
				pWndChild = pWndChild->GetNextWindow ();
				continue;
			}

			if (strClass == STATIC_CLASS)
			{
				pWndSubclassedCtrl = CreateSubclassedCtrl(BCGPCONTROLSCLASS_Static);
			}
			else if (strClass == BUTTON_CLASS)
			{
				if ((pWndChild->GetStyle () & 0xF) == BS_GROUPBOX)
				{
					pWndSubclassedCtrl = CreateSubclassedCtrl(BCGPCONTROLSCLASS_GroupBox);
				}
				else
				{
					pWndSubclassedCtrl = CreateSubclassedCtrl(BCGPCONTROLSCLASS_Button);
				}
			}
			else if (strClass == PROGRESS_CLASS)
			{
				pWndSubclassedCtrl = CreateSubclassedCtrl(BCGPCONTROLSCLASS_Progress);
			}
			else if (strClass == TRACKBAR_CLASS)
			{
				pWndSubclassedCtrl = CreateSubclassedCtrl(BCGPCONTROLSCLASS_TrackBar);
			}
			else if (strClass == EDIT_CLASS)
			{
				pWndSubclassedCtrl = CreateSubclassedCtrl(BCGPCONTROLSCLASS_Edit);
			}
			else if (strClass == COMBOBOX_CLASS)
			{
				pWndSubclassedCtrl = CreateSubclassedCtrl(BCGPCONTROLSCLASS_ComboBox);
			}
			else if (strClass == SCROLLBAR_CLASS)
			{
				pWndSubclassedCtrl = CreateSubclassedCtrl(BCGPCONTROLSCLASS_ScrollBar);
			}
			else if (strClass == UPDOWN_CLASS)
			{
				pWndSubclassedCtrl = CreateSubclassedCtrl(BCGPCONTROLSCLASS_UpDown);
			}

			if (pWndSubclassedCtrl != NULL)
			{
				m_arSubclassedCtrls.Add (pWndSubclassedCtrl);
				pWndSubclassedCtrl->SubclassWindow (pWndChild->GetSafeHwnd ());
			}
		}

		pWndChild->SendMessage (BCGM_ONSETCONTROLVMMODE, (WPARAM) bEnable);
		pWndChild = pWndChild->GetNextWindow ();
	}

	OnChangeVisualManager ();

	if (m_Dlg.IsWindowVisible ())
	{
		m_Dlg.RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
	}
}
//********************************************************************************************************
BOOL CBCGPDlgImpl::EnableAero (BCGPMARGINS& margins)
{
	m_AeroMargins = margins;

	if (HasAeroMargins () && !m_bVisualManagerStyle)
	{
		EnableVisualManagerStyle (TRUE);
	}

	if (m_Dlg.GetSafeHwnd () == NULL)
	{
		return TRUE;
	}

	BOOL bRes = BCGCBProDwmExtendFrameIntoClientArea(m_Dlg.GetSafeHwnd (), &m_AeroMargins);
	BOOL bIsAeroEnabled = BCGCBProDwmIsCompositionEnabled();

	CRect rectClient;
	m_Dlg.GetClientRect (rectClient);

	CWnd* pWndChild = m_Dlg.GetWindow (GW_CHILD);

	while (pWndChild != NULL)
	{
		ASSERT_VALID (pWndChild);

		BOOL bGlass = FALSE;

		if (bIsAeroEnabled && HasAeroMargins ())
		{
			CRect rectChild;
			pWndChild->GetWindowRect (&rectChild);
			m_Dlg.ScreenToClient (&rectChild);

			CRect rectInter;

			if (m_AeroMargins.cxLeftWidth != 0)
			{
				CRect rectAero = rectClient;
				rectAero.right = rectAero.left + m_AeroMargins.cxLeftWidth;

				if (rectInter.IntersectRect (rectAero, rectChild))
				{
					bGlass = TRUE;
				}
			}

			if (!bGlass && m_AeroMargins.cxRightWidth != 0)
			{
				CRect rectAero = rectClient;
				rectAero.left = rectAero.right - m_AeroMargins.cxRightWidth;

				if (rectInter.IntersectRect (rectAero, rectChild))
				{
					bGlass = TRUE;
				}
			}

			if (!bGlass && m_AeroMargins.cyTopHeight != 0)
			{
				CRect rectAero = rectClient;
				rectAero.bottom = rectAero.top + m_AeroMargins.cyTopHeight;

				if (rectInter.IntersectRect (rectAero, rectChild))
				{
					bGlass = TRUE;
				}
			}

			if (!bGlass && m_AeroMargins.cyBottomHeight != 0)
			{
				CRect rectAero = rectClient;
				rectAero.top = rectAero.bottom - m_AeroMargins.cyBottomHeight;

				if (rectInter.IntersectRect (rectAero, rectChild))
				{
					bGlass = TRUE;
				}
			}
		}

		pWndChild->SendMessage (BCGM_ONSETCONTROLAERO, (WPARAM) bGlass);
		pWndChild = pWndChild->GetNextWindow ();
	}

	if (m_Dlg.IsWindowVisible ())
	{
		m_Dlg.RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
	}

	return bRes;
}
//********************************************************************************************************
void CBCGPDlgImpl::GetAeroMargins (BCGPMARGINS& margins) const
{
	margins = m_AeroMargins;
}
//********************************************************************************************************
BOOL CBCGPDlgImpl::HasAeroMargins () const
{
	return 	m_AeroMargins.cxLeftWidth != 0 ||
			m_AeroMargins.cxRightWidth != 0 ||
			m_AeroMargins.cyTopHeight != 0 ||
			m_AeroMargins.cyBottomHeight != 0;
}
//********************************************************************************************************
void CBCGPDlgImpl::ClearAeroAreas (CDC* pDC)
{
	BOOL bIsAeroEnabled = BCGCBProDwmIsCompositionEnabled();

	if (!HasAeroMargins () || m_Dlg.GetSafeHwnd () == NULL || !bIsAeroEnabled)
	{
		return;
	}

	CRect rectClient;
	m_Dlg.GetClientRect (rectClient);

	if (m_AeroMargins.cxLeftWidth != 0)
	{
		CRect rectAero = rectClient;
		rectAero.right = rectAero.left + m_AeroMargins.cxLeftWidth;

		pDC->FillSolidRect (rectAero, RGB (0, 0, 0));
	}

	if (m_AeroMargins.cxRightWidth != 0)
	{
		CRect rectAero = rectClient;
		rectAero.left = rectAero.right - m_AeroMargins.cxRightWidth;

		pDC->FillSolidRect (rectAero, RGB (0, 0, 0));
	}

	if (m_AeroMargins.cyTopHeight != 0)
	{
		CRect rectAero = rectClient;
		rectAero.bottom = rectAero.top + m_AeroMargins.cyTopHeight;

		pDC->FillSolidRect (rectAero, RGB (0, 0, 0));
	}

	if (m_AeroMargins.cyBottomHeight != 0)
	{
		CRect rectAero = rectClient;
		rectAero.top = rectAero.bottom - m_AeroMargins.cyBottomHeight;

		pDC->FillSolidRect (rectAero, RGB (0, 0, 0));
	}
}
//********************************************************************************************************
void CBCGPDlgImpl::OnDWMCompositionChanged ()
{
	if (m_Dlg.GetSafeHwnd () == NULL)
	{
		return;
	}

	OnChangeVisualManager ();

	if (HasAeroMargins ())
	{
		EnableAero (m_AeroMargins);
	}

	m_Dlg.RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
}
//********************************************************************************************************
HBRUSH CBCGPDlgImpl::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	#define MAX_CLASS_NAME	255
	#define STATIC_CLASS	_T("Static")
	#define BUTTON_CLASS	_T("Button")

	if (m_bVisualManagerStyle && !m_lstNonSubclassedItems.IsEmpty() && m_lstNonSubclassedItems.Find (pWnd->GetDlgCtrlID ()) != NULL)
	{
		return NULL;
	}

	if (nCtlColor == CTLCOLOR_STATIC)
	{
		TCHAR lpszClassName [MAX_CLASS_NAME + 1];

		::GetClassName (pWnd->GetSafeHwnd (), lpszClassName, MAX_CLASS_NAME);
		CString strClass = lpszClassName;

		if (m_lstNonSubclassedWndClasses.Find(strClass) != NULL)
		{
			return NULL;
		}

		if (strClass == STATIC_CLASS)
		{
			pDC->SetBkMode(TRANSPARENT);

			if (m_bVisualManagerStyle)
			{
				pDC->SetTextColor (visualManager->GetDlgTextColor(&m_Dlg));
			}

			if (m_bTransparentStaticCtrls && (pWnd->GetStyle () & SS_ICON) != SS_ICON)
			{
				return (HBRUSH) ::GetStockObject (HOLLOW_BRUSH);
			}

			if (m_bIsWhiteBackground && !globalData.IsHighContastMode() && !visualManager->IsDarkTheme())
			{
				return (HBRUSH) ::GetStockObject (WHITE_BRUSH);
			}

			return (HBRUSH) visualManager->GetDlgBackBrush (&m_Dlg).GetSafeHandle ();
		}

		if (strClass == BUTTON_CLASS)
		{
			const DWORD dwStyle = pWnd->GetStyle();

			if ((dwStyle & 0xF) == BS_GROUPBOX)
			{
				if (m_bVisualManagerStyle)
				{
					pDC->SetTextColor (globalData.clrBarText);
					pDC->SetBkMode(TRANSPARENT);
					return (HBRUSH) visualManager->GetDlgBackBrush (&m_Dlg).GetSafeHandle ();
				}
			}

			if (((dwStyle & 0xF) & BS_CHECKBOX) == 0)
			{
				pDC->SetBkMode(TRANSPARENT);
			}

			return (HBRUSH) ::GetStockObject (m_bIsWhiteBackground && !globalData.IsHighContastMode() && !visualManager->IsDarkTheme() ? WHITE_BRUSH : HOLLOW_BRUSH);
		}
	}

	return NULL;
}
//*********************************************************************************************************
BOOL CBCGPDlgImpl::OnNcPaint ()
{
#ifndef _BCGSUITE_
	if (globalData.m_bInSettingsChange || !IsOwnerDrawCaption ())
	{
		return FALSE;
	}
#else
	if (!IsOwnerDrawCaption ())
	{
		return FALSE;
	}
#endif

	return visualManagerMFC->OnNcPaint (&m_Dlg,
		m_lstCaptionSysButtons, m_rectRedraw);
}
//********************************************************************************
BOOL CBCGPDlgImpl::OnNcCalcSize(BOOL /*bCalcValidRects*/, NCCALCSIZE_PARAMS FAR* lpncsp)
{
	ASSERT (lpncsp != NULL);

	if (IsOwnerDrawCaption ())
	{
		if (!CBCGPGlobalUtils::IsXPPlatformToolset() && !m_Dlg.IsIconic() && 
			(m_Dlg.GetStyle() & WS_MAXIMIZE) == 0 && 
			(m_Dlg.GetStyle() & WS_THICKFRAME) == 0 &&
			!IsLayoutEnabled() && m_bHasCaption)
		{
			CSize size;
			size.cx = ::GetSystemMetrics(SM_CXPADDEDBORDER) + ::GetSystemMetrics(SM_CXBORDER);
			size.cy = ::GetSystemMetrics(SM_CXPADDEDBORDER) + ::GetSystemMetrics(SM_CXBORDER);
			::InflateRect(&lpncsp->rgrc[0], -size.cx, -size.cy);
		}

		lpncsp->rgrc[0].top += GetCaptionHeight ();
	}

	return (m_Dlg.GetStyle () & WS_MAXIMIZE) == WS_MAXIMIZE && IsOwnerDrawCaption ();
}
//********************************************************************************
CRect CBCGPDlgImpl::GetCaptionRect ()
{
	if ((m_Dlg.GetStyle () & WS_CAPTION) != WS_CAPTION && !m_bHasCaption)
	{
		return CRect(0, 0, 0, 0);
	}

	CSize szSystemBorder (globalUtils.GetSystemBorders (&m_Dlg));

	if (m_Dlg.IsIconic () || 
		(m_Dlg.GetStyle () & WS_MAXIMIZE) == WS_MAXIMIZE)
	{
		szSystemBorder = CSize (0, 0);
	}

	CRect rectWnd;
	m_Dlg.GetWindowRect (&rectWnd);
	m_Dlg.ScreenToClient (&rectWnd);

	int cyCaption = GetCaptionHeight();
	int cyOffset = szSystemBorder.cy;
	if (!m_Dlg.IsIconic () || (m_Dlg.IsIconic () && (m_Dlg.GetStyle() & WS_CHILD) == 0))
	{
		cyOffset += cyCaption;
	}

	rectWnd.OffsetRect (szSystemBorder.cx, cyOffset);

	CRect rectCaption (	rectWnd.left + szSystemBorder.cx, 
						rectWnd.top + szSystemBorder.cy, 
						rectWnd.right - szSystemBorder.cx, 
						rectWnd.top + szSystemBorder.cy + cyCaption);

	if (m_Dlg.IsIconic ())
	{
		rectCaption.left  += GetSystemMetrics(SM_CXFIXEDFRAME);
		rectCaption.right -= GetSystemMetrics(SM_CXFIXEDFRAME);
		rectCaption.OffsetRect (0, (rectWnd.Height () - rectCaption.Height ()) / 2);
	}

	return rectCaption;
}
//*************************************************************************
void CBCGPDlgImpl::UpdateCaption ()
{
	if (!IsOwnerDrawCaption ())
	{
		return;
	}

	if (m_lstCaptionSysButtons.IsEmpty ())
	{
		//------------------------
		// Create caption buttons:
		//------------------------
		const DWORD dwStyle = m_Dlg.GetStyle ();
		HMENU hSysMenu = NULL;
		CMenu* pSysMenu = m_Dlg.GetSystemMenu (FALSE);

		if (pSysMenu != NULL && ::IsMenu (pSysMenu->m_hMenu))
		{
			hSysMenu = pSysMenu->GetSafeHmenu ();
			if (!::IsMenu (hSysMenu) || (m_Dlg.GetStyle () & WS_SYSMENU) == 0)
			{
				hSysMenu = NULL;
			}
		}

		if (hSysMenu != NULL)
		{
			m_lstCaptionSysButtons.AddTail (new CBCGPFrameCaptionButton (HTCLOSE_BCG));

			if ((m_Dlg.GetExStyle() & WS_EX_TOOLWINDOW) == 0)
			{
				BOOL bHelp = TRUE;
				if ((dwStyle & WS_MAXIMIZEBOX) == WS_MAXIMIZEBOX)
				{
					m_lstCaptionSysButtons.AddTail (new CBCGPFrameCaptionButton (HTMAXBUTTON_BCG));
					bHelp = FALSE;
				}

				if ((dwStyle & WS_MINIMIZEBOX) == WS_MINIMIZEBOX)
				{
					m_lstCaptionSysButtons.AddTail (new CBCGPFrameCaptionButton (HTMINBUTTON_BCG));
					bHelp = FALSE;
				}

#ifndef _BCGSUITE_
				if (bHelp && ((m_Dlg.GetExStyle() & WS_EX_CONTEXTHELP) == WS_EX_CONTEXTHELP))
				{
					m_lstCaptionSysButtons.AddTail (new CBCGPFrameCaptionButton (HTHELPBUTTON_BCG));
				}
#endif
			}
		}
	}

	CRect rectCaption = GetCaptionRect ();

#ifndef _BCGSUITE_
    if (!visualManagerMFC->OnUpdateNcButtons(&m_Dlg, m_lstCaptionSysButtons, rectCaption))
#endif
    {
	    CSize sizeButton = visualManagerMFC->GetNcBtnSize ((m_Dlg.GetExStyle() & WS_EX_TOOLWINDOW) == WS_EX_TOOLWINDOW);
	    sizeButton.cy = min (sizeButton.cy, rectCaption.Height () - 2);

#ifndef _BCGSUITE_
		CSize szOffset = CBCGPVisualManager::GetInstance()->GetFrameCaptionButtonOffset(&m_Dlg);
#else
		CSize szOffset(0, 0);
#endif
	    int x = rectCaption.right - sizeButton.cx + szOffset.cx;
	    int y = rectCaption.top + max (0, (rectCaption.Height () - sizeButton.cy) / 2) + szOffset.cy;

	    for (POSITION pos = m_lstCaptionSysButtons.GetHeadPosition (); pos != NULL;)
	    {
		    CBCGPFrameCaptionButton* pButton = (CBCGPFrameCaptionButton*)
			    m_lstCaptionSysButtons.GetNext (pos);
		    ASSERT_VALID (pButton);

		    pButton->SetRect (CRect (CPoint (x, y), sizeButton));

		    x -= sizeButton.cx;
	    }
    }

#ifndef _BCGSUITE_
    m_Dlg.SendMessage (BCGM_ONAFTERUPDATECAPTION);
#endif

	m_Dlg.RedrawWindow (NULL, NULL,
		RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOCHILDREN);
}
//*************************************************************************
void CBCGPDlgImpl::UpdateCaptionButtons()
{
	if ((m_Dlg.GetStyle () & WS_SYSMENU) == 0)
	{
		while (!m_lstCaptionSysButtons.IsEmpty ())
		{
			delete m_lstCaptionSysButtons.RemoveHead ();
		}
	}
	else
	{
		CMenu* pSysMenu = m_Dlg.GetSystemMenu (FALSE);

		if (pSysMenu == NULL || !::IsMenu (pSysMenu->m_hMenu))
		{
			return;
		}

		for (POSITION pos = m_lstCaptionSysButtons.GetHeadPosition (); pos != NULL;)
		{
			CBCGPFrameCaptionButton* pButton = (CBCGPFrameCaptionButton*)m_lstCaptionSysButtons.GetNext (pos);
			ASSERT_VALID (pButton);

			if (pButton->GetHit () == HTCLOSE_BCG)
			{
				BOOL bGrayed = pSysMenu->GetMenuState (SC_CLOSE, MF_BYCOMMAND) & MF_GRAYED;
				pButton->m_bEnabled = bGrayed ? FALSE : TRUE;
			}

			if (pButton->GetHit () == HTMAXBUTTON_BCG)
			{
				BOOL bGrayed = pSysMenu->GetMenuState (SC_MAXIMIZE, MF_BYCOMMAND) & MF_GRAYED;
				pButton->m_bEnabled = bGrayed ? FALSE : TRUE;
			}

			if (pButton->GetHit () == HTMINBUTTON_BCG)
			{
				BOOL bGrayed = pSysMenu->GetMenuState (SC_MINIMIZE, MF_BYCOMMAND) & MF_GRAYED;
				pButton->m_bEnabled = bGrayed ? FALSE : TRUE;
			}

#ifndef _BCGSUITE_
			if (pButton->GetHit () == HTHELPBUTTON_BCG)
			{
				pButton->m_bEnabled = TRUE;
			}
#endif
		}
	}

#ifndef _BCGSUITE_
	m_Dlg.SendMessage (BCGM_ONAFTERUPDATECAPTION);
#endif

	m_Dlg.RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOCHILDREN);
}
//*************************************************************************
UINT CBCGPDlgImpl::OnNcHitTest (CPoint point)
{
	CPoint ptScreen = point;

	m_Dlg.ScreenToClient (&point);

	if (!m_rectResizeBox.IsRectEmpty ())
	{
		if (m_rectResizeBox.PtInRect(point))
		{
			BOOL bRTL = m_Dlg.GetExStyle() & WS_EX_LAYOUTRTL;
			return bRTL ? HTBOTTOMLEFT : HTBOTTOMRIGHT;
		}
	}

	if (IsDragClientAreaEnabled())
	{
		if ((m_Dlg.GetStyle() & WS_CHILD) == 0 && !m_Dlg.IsIconic() && !m_Dlg.IsZoomed())
		{
			CRect rectClient;
			m_Dlg.GetClientRect(&rectClient);

			if (rectClient.PtInRect(point) && CWnd::WindowFromPoint(ptScreen)->GetSafeHwnd() == m_Dlg.GetSafeHwnd())
			{
				CString strInfo;
				CString strDescription;
				BOOL bIsClickable = FALSE;

				if (GetControlInfoTipFromPoint(point, strInfo, strDescription, bIsClickable) == NULL)
				{
					return HTCAPTION;
				}
			}
		}
	}

	if (!IsOwnerDrawCaption ())
	{
		return HTNOWHERE;
	}

	CSize szSystemBorder(globalUtils.GetSystemBorders (&m_Dlg));

	if (m_Dlg.IsIconic ())
	{
		szSystemBorder = CSize (0, 0);
	}

	int cxOffset = szSystemBorder.cx;
	int cyOffset = szSystemBorder.cy;

	if (!m_Dlg.IsIconic () || (m_Dlg.IsIconic () && (m_Dlg.GetStyle() & WS_CHILD) == 0))
	{
		cyOffset += GetCaptionHeight ();
	}

	if (m_Dlg.IsZoomed ())
	{
		cxOffset -= szSystemBorder.cx;
		cyOffset -= szSystemBorder.cy;
	}

	point.Offset (cxOffset, cyOffset);

	for (POSITION pos = m_lstCaptionSysButtons.GetHeadPosition (); pos != NULL;)
	{
		CBCGPFrameCaptionButton* pButton = (CBCGPFrameCaptionButton*)
			m_lstCaptionSysButtons.GetNext (pos);
		ASSERT_VALID (pButton);

		if (pButton->GetRect ().PtInRect (point))
		{
			return pButton->m_nHit;
		}
	}

	CRect rectCaption = GetCaptionRect ();
	if (rectCaption.PtInRect (point))
	{
		if ((m_Dlg.GetStyle() & WS_SYSMENU) == WS_SYSMENU && (m_Dlg.GetExStyle () & WS_EX_TOOLWINDOW) == 0)
		{
			CRect rectSysMenu = rectCaption;
			rectSysMenu.right = rectSysMenu.left + ::GetSystemMetrics (SM_CXSMICON) + 2 +
				(m_Dlg.IsZoomed () ? szSystemBorder.cx : 0);

			return rectSysMenu.PtInRect (point) ? HTSYSMENU : HTCAPTION;
		}

		return HTCAPTION;
	}

	return HTNOWHERE;
}
//*************************************************************************
void CBCGPDlgImpl::OnNcMouseMove(UINT /*nHitTest*/, CPoint point)
{
	if (!IsOwnerDrawCaption ())
	{
		return;
	}

	OnTrackCaptionButtons (point);
}
//*************************************************************************
void CBCGPDlgImpl::OnLButtonDown(CPoint /*point*/)
{
	if (m_nHotSysButton == HTNOWHERE)
	{
		return;
	}

	CBCGPFrameCaptionButton* pBtn = GetSysButton (m_nHotSysButton);
	if (pBtn != NULL)
	{
		m_nHitSysButton = m_nHotSysButton;
		pBtn->m_bPushed = TRUE;
		RedrawCaptionButton (pBtn);
	}
}
//*************************************************************************
BOOL CBCGPDlgImpl::ProcessInfoTipClick(CPoint point)
{
	CString strTipText;
	CString strDescription;
	BOOL bIsClickable = FALSE;
	HWND hwndCtrl = GetControlInfoTipFromPoint(point, strTipText, strDescription, bIsClickable);
	
	if (hwndCtrl != NULL && bIsClickable)
	{
		m_Dlg.SendMessage(BCGM_ONCLICKINFOTIP, 0, (LPARAM)hwndCtrl);
		return TRUE;
	}

	return FALSE;
}
//*************************************************************************
void CBCGPDlgImpl::DrawGroupBoxes(CDC* pDC)
{
	ASSERT_VALID(pDC);

	if (!m_bGroupBoxesDrawByParent)
	{
		return;
	}

	for (CWnd* pWndChild = m_Dlg.GetWindow(GW_CHILD); pWndChild != NULL; pWndChild = pWndChild->GetNextWindow())
	{
		ASSERT_VALID(pWndChild);

		CBCGPGroup* pGroup = DYNAMIC_DOWNCAST(CBCGPGroup, pWndChild);
		if (pGroup->GetSafeHwnd() != NULL && pGroup->m_bDrawByParent)
		{
			if ((pGroup->GetStyle() & WS_VISIBLE) == WS_VISIBLE)
			{
				CRect rect;
				pGroup->GetWindowRect(rect);
				m_Dlg.ScreenToClient(rect);

				pGroup->DoDraw(pDC, rect);
			}
		}
	}
}
//*************************************************************************
void CBCGPDlgImpl::OnLButtonUp(CPoint point)
{
	if (ProcessInfoTipClick(point))
	{
		return;
	}

	if (!IsOwnerDrawCaption ())
	{
		return;
	}

	switch (m_nHitSysButton)
	{
	case HTCLOSE_BCG:
	case HTMAXBUTTON_BCG:
	case HTMINBUTTON_BCG:
#ifndef _BCGSUITE_
	case HTHELPBUTTON_BCG:
#endif
		{
			UINT nHot = m_nHotSysButton;
			UINT nHit = m_nHitSysButton;

			StopCaptionButtonsTracking ();

			if (nHot == nHit)
			{
				UINT nSysCmd = 0;

				switch (nHot)
				{
				case HTCLOSE_BCG:
					nSysCmd = SC_CLOSE;
					break;

				case HTMAXBUTTON_BCG:
					nSysCmd = 
						(m_Dlg.GetStyle () & WS_MAXIMIZE) == WS_MAXIMIZE ? SC_RESTORE : SC_MAXIMIZE;
					break;

				case HTMINBUTTON_BCG:
					nSysCmd = m_Dlg.IsIconic () ? SC_RESTORE : SC_MINIMIZE;
					break;

#ifndef _BCGSUITE_
				case HTHELPBUTTON_BCG:
					nSysCmd = SC_CONTEXTHELP;
					break;
#endif
				}

				m_Dlg.PostMessage (WM_SYSCOMMAND, nSysCmd);
			}
		}
	}
}
//*************************************************************************
void CBCGPDlgImpl::OnMouseMove(CPoint point)
{
	if (!IsOwnerDrawCaption ())
	{
		return;
	}

	CPoint ptScreen = point;
	m_Dlg.ClientToScreen (&ptScreen);

	OnTrackCaptionButtons (ptScreen);
}
//*************************************************************************
CBCGPFrameCaptionButton* CBCGPDlgImpl::GetSysButton (UINT nHit)
{
	for (POSITION pos = m_lstCaptionSysButtons.GetHeadPosition (); pos != NULL;)
	{
		CBCGPFrameCaptionButton* pButton = (CBCGPFrameCaptionButton*)
			m_lstCaptionSysButtons.GetNext (pos);
		ASSERT_VALID (pButton);

		if (pButton->m_nHit == nHit)
		{
			return pButton;
		}
	}

	return NULL;
}
//*************************************************************************
void CBCGPDlgImpl::SetHighlightedSysButton (UINT nHit)
{
	BOOL bRedraw = FALSE;

	for (POSITION pos = m_lstCaptionSysButtons.GetHeadPosition (); pos != NULL;)
	{
		CBCGPFrameCaptionButton* pButton = (CBCGPFrameCaptionButton*)
			m_lstCaptionSysButtons.GetNext (pos);
		ASSERT_VALID (pButton);

		if (pButton->m_nHit == nHit)
		{
			if (pButton->m_bFocused)
			{
				return;
			}

			pButton->m_bFocused = TRUE;
			bRedraw = TRUE;
		}
	}
}
//*************************************************************************************
void CBCGPDlgImpl::OnTrackCaptionButtons (CPoint point)
{
	if (CBCGPPopupMenu::GetActiveMenu () != NULL && !CBCGPToolBar::IsCustomizeMode())
	{
		return;
	}

	UINT nHot = m_nHotSysButton;
	CBCGPFrameCaptionButton* pBtn = GetSysButton (OnNcHitTest (point));

	if (pBtn != NULL && pBtn->m_bEnabled)
	{
		m_nHotSysButton = pBtn->GetHit ();
		pBtn->m_bFocused = TRUE;
	}
	else
	{
		m_nHotSysButton = HTNOWHERE;
	}

	if (m_nHotSysButton != nHot)
	{
		RedrawCaptionButton (pBtn);

		CBCGPFrameCaptionButton* pBtnOld = GetSysButton (nHot);
		if (pBtnOld != NULL)
		{
			pBtnOld->m_bFocused = FALSE;
			RedrawCaptionButton (pBtnOld);
		}
	}

	if (m_nHitSysButton == HTNOWHERE)
	{
		if (nHot != HTNOWHERE && m_nHotSysButton == HTNOWHERE)
		{
			::ReleaseCapture();
		}
		else if (nHot == HTNOWHERE && m_nHotSysButton != HTNOWHERE)
		{
			m_Dlg.SetCapture ();
		}
	}
}
//************************************************************************************
void CBCGPDlgImpl::StopCaptionButtonsTracking ()
{
	if (m_nHitSysButton != HTNOWHERE)
	{
		CBCGPFrameCaptionButton* pBtn = GetSysButton (m_nHitSysButton);
		m_nHitSysButton = HTNOWHERE;

		ReleaseCapture ();
		if (pBtn != NULL)
		{
			pBtn->m_bPushed = FALSE;
			RedrawCaptionButton (pBtn);
		}
	}

	if (m_nHotSysButton != HTNOWHERE)
	{
		CBCGPFrameCaptionButton* pBtn = GetSysButton (m_nHotSysButton);
		m_nHotSysButton = HTNOWHERE;

		ReleaseCapture ();
		if (pBtn != NULL)
		{
			pBtn->m_bFocused = FALSE;
			RedrawCaptionButton (pBtn);
		}
	}
}
//************************************************************************************
void CBCGPDlgImpl::RedrawCaptionButton (CBCGPFrameCaptionButton* pBtn)
{
	if (pBtn ==	NULL)
	{
		return;
	}

	ASSERT_VALID (pBtn);

	m_rectRedraw = pBtn->GetRect ();
	m_Dlg.SendMessage (WM_NCPAINT);
	m_rectRedraw.SetRectEmpty ();

	m_Dlg.UpdateWindow ();
}
//********************************************************************************
void CBCGPDlgImpl::OnWindowPosChanged(WINDOWPOS FAR* lpwndpos)
{
	if (m_bWindowPosChanging)
	{
		return;
	}

	if (((lpwndpos->flags & SWP_NOSIZE) == 0 || (lpwndpos->flags & SWP_FRAMECHANGED)) && 
		(IsOwnerDrawCaption ()))
	{
		m_bWindowPosChanging = TRUE;

		m_bIsWindowRgn = visualManagerMFC->OnSetWindowRegion (
			&m_Dlg, CSize (lpwndpos->cx, lpwndpos->cy));

		m_bWindowPosChanging = FALSE;
	}
}
//************************************************************************************
void CBCGPDlgImpl::OnChangeVisualManager ()
{
	if (m_Dlg.GetSafeHwnd () == NULL)
	{
		return;
	}

	if ((m_Dlg.GetStyle() & WS_CHILD) == 0)
	{
#ifndef _BCGSUITE_
		if (m_bVisualManagerNCArea)
		{
			globalData.EnableWindowAero(m_Dlg.GetSafeHwnd(), CBCGPVisualManager::GetInstance()->IsDWMCaptionSupported());
		}

		if (!m_bDisableShadows)
		{
			if (CBCGPVisualManager::GetInstance ()->CanShowFrameShadows() &&
				m_bVisualManagerNCArea && IsOwnerDrawCaption())
			{
				if (globalData.m_bShowFrameLayeredShadows)
				{
					if (m_pShadow == NULL)
					{
						BOOL bInteraction = (m_Dlg.GetStyle() & WS_THICKFRAME) != 0 || IsLayoutEnabled();

						m_pShadow = new CBCGPShadowManager(&m_Dlg, bInteraction);
						m_pShadow->Create(CSize(6, 6));
					}
					else
					{
						m_pShadow->UpdateBaseColor();
					}
				}
				else
				{
					globalUtils.EnableWindowShadow(&m_Dlg, TRUE);
				}
			}
			else
			{
				if (!globalData.m_bShowFrameLayeredShadows)
				{
					globalUtils.EnableWindowShadow(&m_Dlg, FALSE);
				}

				if (m_pShadow != NULL)
				{
					delete m_pShadow;
					m_pShadow = NULL;
				}
			}
		}

#endif
		CRect rectWindow;
		m_Dlg.GetWindowRect (rectWindow);

		BOOL bZoomed = m_Dlg.IsZoomed ();

		if (IsOwnerDrawCaption())
		{
			BOOL bChangeBorder = FALSE;

			if ((m_Dlg.GetStyle () & WS_BORDER) == WS_BORDER && m_bHasBorder)
			{
				bChangeBorder = TRUE;
				m_bWindowPosChanging = TRUE;
				m_Dlg.ModifyStyle (WS_BORDER, 0, SWP_FRAMECHANGED);
				m_bWindowPosChanging = FALSE;
			}

			m_bIsWindowRgn = visualManagerMFC->OnSetWindowRegion (
				&m_Dlg, rectWindow.Size ());

			if (bZoomed && bChangeBorder)
			{
#ifndef _BCGSUITE_
				m_Dlg.ShowWindow(CBCGPVisualManager::GetInstance ()->IsSmallSystemBorders() ? SW_RESTORE : SW_MINIMIZE);
#else
				m_Dlg.ShowWindow(SW_MINIMIZE);
#endif
				m_Dlg.ShowWindow(SW_MAXIMIZE);
			}
		}
		else
		{
			BOOL bChangeBorder = FALSE;

			if ((m_Dlg.GetStyle () & WS_BORDER) == 0 && m_bHasBorder)
			{
				bChangeBorder = TRUE;
				m_bWindowPosChanging = TRUE;
				m_Dlg.ModifyStyle (0, WS_BORDER, SWP_FRAMECHANGED);
				m_bWindowPosChanging = FALSE;
			}

			if (m_bIsWindowRgn)
			{
				m_bIsWindowRgn = FALSE;
				m_Dlg.SetWindowRgn (NULL, TRUE);
			}

			if (bZoomed && bChangeBorder)
			{
				NCCALCSIZE_PARAMS params;
				ZeroMemory(&params, sizeof (NCCALCSIZE_PARAMS));
				params.rgrc[0].left   = rectWindow.left;
				params.rgrc[0].top    = rectWindow.top;
				params.rgrc[0].right  = rectWindow.right;
				params.rgrc[0].bottom = rectWindow.bottom;

				m_Dlg.CalcWindowRect (&params.rgrc[0], CFrameWnd::adjustBorder);

				params.rgrc[0].top += GetCaptionHeight ();

				m_Dlg.SetWindowPos (NULL, params.rgrc[0].left, params.rgrc[0].top, 
					params.rgrc[0].right - params.rgrc[0].left, params.rgrc[0].bottom - params.rgrc[0].top,
					SWP_NOACTIVATE | SWP_NOZORDER);
			}
			else
			{
				m_Dlg.SetWindowPos (NULL, -1, -1, rectWindow.Width () + 1, rectWindow.Height () + 1,
					SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
				m_Dlg.SetWindowPos (NULL, -1, -1, rectWindow.Width (), rectWindow.Height (),
					SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
			}
		}

		UpdateCaption ();
		UpdateCaptionButtons();
	}
	
	m_Dlg.SendMessageToDescendants (BCGM_CHANGEVISUALMANAGER, 0, 0, FALSE, TRUE);
}
//**********************************************************************************************
int CBCGPDlgImpl::GetCaptionHeight ()
{
	if (m_Dlg.GetSafeHwnd () == NULL)
	{
		return 0;
	}

	if ((m_Dlg.GetStyle () & WS_CAPTION) != WS_CAPTION && !m_bHasCaption)
	{
		return 0;
	}

	return ::GetSystemMetrics ((m_Dlg.GetExStyle () & WS_EX_TOOLWINDOW) == WS_EX_TOOLWINDOW ? SM_CYSMCAPTION : SM_CYCAPTION);
}
//********************************************************************************
void CBCGPDlgImpl::OnGetMinMaxInfo (MINMAXINFO FAR* lpMMI)
{
	ASSERT (lpMMI != NULL);

	if ((m_Dlg.GetStyle () & WS_CAPTION) == 0 ||
		(m_Dlg.GetStyle () & WS_BORDER) == 0)
	{
		CRect rectWindow;
		m_Dlg.GetWindowRect (&rectWindow);

		if (m_Dlg.IsIconic () || m_Dlg.IsZoomed ())
		{
			WINDOWPLACEMENT wp;
			wp.length = sizeof (WINDOWPLACEMENT);

			m_Dlg.GetWindowPlacement (&wp);

			rectWindow = wp.rcNormalPosition;
		}

		CRect rect (0, 0, 0, 0);

		MONITORINFO mi;
		mi.cbSize = sizeof (MONITORINFO);

		if (GetMonitorInfo (MonitorFromPoint (rectWindow.CenterPoint (), MONITOR_DEFAULTTONEAREST),
			&mi))
		{
			CRect rectWork = mi.rcWork;
			CRect rectScreen = mi.rcMonitor;

			rect.left = rectWork.left - rectScreen.left;
			rect.top = rectWork.top - rectScreen.top;

			rect.right = rect.left + rectWork.Width ();
			rect.bottom = rect.top + rectWork.Height ();
		}
		else
		{
			::SystemParametersInfo (SPI_GETWORKAREA, 0, &rect, 0);
		}

#ifdef _BCGSUITE_
		int nShellAutohideBars = globalData.m_nShellAutohideBars;
#else
		int nShellAutohideBars = globalData.GetShellAutohideBars ();
#endif

		if (nShellAutohideBars & BCGP_AUTOHIDE_BOTTOM)
		{
			rect.bottom -= 2;
		}

		if (nShellAutohideBars & BCGP_AUTOHIDE_TOP)
		{
			rect.top += 2;
		}

		if (nShellAutohideBars & BCGP_AUTOHIDE_RIGHT)
		{
			rect.right -= 2;
		}

		if (nShellAutohideBars & BCGP_AUTOHIDE_LEFT)
		{
			rect.left += 2;
		}

		lpMMI->ptMaxPosition.x = rect.left;
		lpMMI->ptMaxPosition.y = rect.top;
		lpMMI->ptMaxSize.x = rect.Width ();
		lpMMI->ptMaxSize.y = rect.Height ();
	}

	if (m_pLayout != NULL && 
		m_LayoutMMI.ptMinTrackSize.x > 0 && m_LayoutMMI.ptMinTrackSize.y > 0)
	{
		lpMMI->ptMinTrackSize = m_LayoutMMI.ptMinTrackSize;
	}
}
//********************************************************************************
void CBCGPDlgImpl::EnableLayout(BOOL bEnable, CRuntimeClass* pRTC, BOOL bResizeBox)
{
	if (m_pLayout != NULL)
	{
		delete m_pLayout;
		m_pLayout = NULL;
	}

	if (!bEnable)
	{
		return;
	}

	if (pRTC == NULL)
	{
		pRTC = RUNTIME_CLASS (CBCGPStaticLayout);
	}
	else if (!pRTC->IsDerivedFrom (RUNTIME_CLASS (CBCGPControlsLayout)))
	{
		ASSERT(FALSE);
		return;
	}

	m_pLayout = (CBCGPControlsLayout*)pRTC->CreateObject ();
	m_bResizeBox = bResizeBox;
}
//********************************************************************************
void CBCGPDlgImpl::EnableDragClientArea(BOOL bEnable)
{
	m_bDragClientArea = bEnable;
}
//********************************************************************************
int CBCGPDlgImpl::OnCreate()
{
	if (globalData.m_bIsRTL)
	{
		m_Dlg.ModifyStyleEx (0, WS_EX_LAYOUTRTL);
	}

	const DWORD dwStyle = m_Dlg.GetStyle ();
	BOOL bIsChild = (dwStyle & WS_CHILD) == WS_CHILD;

	if (bIsChild)
	{
		m_bResizeBox = FALSE;
	}

	if (m_pLayout == NULL)
	{
		return 0;
	}

	ASSERT_VALID(m_pLayout);

	if (!m_pLayout->Create(&m_Dlg))
	{
		delete m_pLayout;
		m_pLayout = NULL;
		return -1;
	}

	CBCGPDialog* pDialog = DYNAMIC_DOWNCAST(CBCGPDialog, &m_Dlg);
	CBCGPPropertySheet* pPropSheet = DYNAMIC_DOWNCAST(CBCGPPropertySheet, &m_Dlg);

	if (!bIsChild && (pDialog != NULL || pPropSheet != NULL))
	{
		CRect rect;
		m_Dlg.GetClientRect(&rect);

		m_Dlg.ModifyStyle(DS_MODALFRAME, WS_POPUP | WS_THICKFRAME);

		::AdjustWindowRectEx(&rect, m_Dlg.GetStyle (), 
			::IsMenu(m_Dlg.GetMenu()->GetSafeHmenu()), m_Dlg.GetExStyle ());

		BOOL bIsCompositionEnabled = FALSE;

	#if !defined(_BCGSUITE_)
		bIsCompositionEnabled = CBCGPVisualManager::GetInstance()->IsDWMCaptionSupported();
	#else
		bIsCompositionEnabled = BCGCBProDwmIsCompositionEnabled();
	#endif

		if (m_bVisualManagerNCArea && IsOwnerDrawCaption() && !bIsCompositionEnabled)
		{
			int nHeight = ::GetSystemMetrics (SM_CYCAPTION);
			rect.top -= nHeight / 2;
			rect.bottom += nHeight - nHeight / 2;
		}

		m_Dlg.SetWindowPos(NULL, 0, 0, rect.Width(), rect.Height(), SWP_FRAMECHANGED | 
			SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);

		if (pPropSheet == NULL)
		{
			m_Dlg.GetWindowRect (rect);
			m_LayoutMMI.ptMinTrackSize.x = rect.Width ();
			m_LayoutMMI.ptMinTrackSize.y = rect.Height ();
		}
	}

	if (pPropSheet == NULL && DYNAMIC_DOWNCAST(CBCGPPropertyPage, &m_Dlg) == NULL)
	{
		CRect rect;
		m_Dlg.GetClientRect (rect);
		m_pLayout->SetMinSize (rect.Size ());
	}

	return 0;
}
//********************************************************************************
void CBCGPDlgImpl::OnInit()
{
	CreateTooltipInfo();
	DoAutoResizeControls();

	if (m_nEditBoxesVertAlignment >= 0)
	{
		SetEditBoxesVerticalAlignment(m_nEditBoxesVertAlignment);
	}

#ifndef _BCGSUITE_
	if (m_bEditBoxesSimplifiedIcon != -1)
	{
		SetEditBoxesSimplifiedBrowseIcon(m_bEditBoxesSimplifiedIcon);
	}
#endif
}
//********************************************************************************
void CBCGPDlgImpl::AdjustControlsLayout()
{
	if (m_pLayout != NULL)
	{
		BOOL bUpdate = FALSE;

		if (m_bResizeBox)
		{
			if (!m_rectResizeBox.IsRectEmpty ())
			{
				m_Dlg.InvalidateRect(m_rectResizeBox);
				bUpdate = TRUE;
			}

			CRect rectClient;
			m_Dlg.GetClientRect(rectClient);

			m_rectResizeBox = rectClient;
			m_rectResizeBox.left = m_rectResizeBox.right - ::GetSystemMetrics(SM_CXVSCROLL);
			m_rectResizeBox.top = m_rectResizeBox.bottom - ::GetSystemMetrics(SM_CYHSCROLL);

			if (!m_rectResizeBox.IsRectEmpty ())
			{
				m_Dlg.InvalidateRect(m_rectResizeBox);
				bUpdate = TRUE;
			}
		}
		else
		{
			if (!m_rectResizeBox.IsRectEmpty ())
			{
				m_rectResizeBox.SetRectEmpty();
				bUpdate = TRUE;
			}
		}
		
		if (!m_mapCtrlInfoTip.IsEmpty())
		{
			m_Dlg.Invalidate();
			bUpdate = TRUE;
		}

		m_pLayout->AdjustLayout();

		if (bUpdate)
		{
			m_Dlg.UpdateWindow();
		}
	}
}
//********************************************************************************
void CBCGPDlgImpl::DrawResizeBox(CDC* pDC)
{
	if (!m_rectResizeBox.IsRectEmpty())
	{
		if (globalData.IsHighContastMode())
		{
			pDC->SetBkMode(TRANSPARENT);
		}

#ifndef _BCGSUITE_
		CBCGPVisualManager::GetInstance ()->OnDrawDlgSizeBox (pDC, &m_Dlg, m_rectResizeBox);
#else
		CMFCVisualManager::GetInstance ()->OnDrawStatusBarSizeBox (pDC, NULL, m_rectResizeBox);
#endif
	}
}
//****************************************************************************
void CBCGPDlgImpl::GetControlInfoTipRect(CWnd* pWndCtrl, CRect& rect, DWORD dwVertAlign, const CPoint& ptOffset)
{
	ASSERT(pWndCtrl->GetSafeHwnd() != NULL);

	pWndCtrl->GetWindowRect(rect);
	m_Dlg.ScreenToClient(rect);
	
	const int nSize = globalUtils.ScaleByDPI(16);

	int nVertOffset = 0;

	switch (dwVertAlign)
	{
	case DT_VCENTER:
		nVertOffset = (rect.Height() - nSize) / 2;
		break;

	case DT_BOTTOM:
		nVertOffset = rect.Height() - nSize;
		break;
	}
	
	rect.OffsetRect(rect.Width() + nSize / 3 + ptOffset.x, nVertOffset + ptOffset.y);

	rect.right = rect.left + nSize;
	rect.bottom = rect.top + nSize;
}
//****************************************************************************
void CBCGPDlgImpl::DrawControlInfoTips(CDC* pDC)
{
	ASSERT_VALID(pDC);

	for (POSITION pos = m_mapCtrlInfoTip.GetStartPosition(); pos != NULL;)
	{
		HWND hwndCtrl = NULL;
		CBCGPControlInfoTip info;

		m_mapCtrlInfoTip.GetNextAssoc(pos, hwndCtrl, info);

		CWnd* pWnd = CWnd::FromHandle(hwndCtrl);

		if (pWnd->GetSafeHwnd() != NULL && (pWnd->GetStyle() & WS_VISIBLE) != 0)
		{
			CRect rectWnd;
			GetControlInfoTipRect(pWnd, rectWnd, info.m_dwVertAlign, info.m_ptOffset);

			info.m_rect = m_mapCtrlInfoTip[hwndCtrl].m_rect = rectWnd;

			if (info.m_style == CBCGPControlInfoTip::BCGPINFOTIP_Custom)
			{
				if (m_Dlg.SendMessage(BCGM_ONDRAWCUSTOMINFOTIP, (WPARAM)pDC->GetSafeHdc(), (LPARAM)&info) == 0)
				{				
					if (m_bVisualManagerStyle)
					{
						CBCGPVisualManager::GetInstance()->OnDrawControlInfoTip(pDC, rectWnd, pWnd);
					}
					else
					{
						CBCGPVisualManager::GetInstance ()->CBCGPVisualManager::OnDrawControlInfoTip(pDC, rectWnd, pWnd);
					}
				}
			}
			else
			{
				if (LoadInfoTipIcons(m_InfotipIcons) != CSize(0, 0))
				{
					int nIndex = (int)info.m_style;
					m_InfotipIcons.DrawEx(pDC, rectWnd, nIndex, CBCGPToolBarImages::ImageAlignHorzCenter, CBCGPToolBarImages::ImageAlignVertCenter);
				}
			}
		}
	}
}
//****************************************************************************
HWND CBCGPDlgImpl::GetControlInfoTipFromPoint(CPoint point, CString& strInfo, CString& strDescription, BOOL& bIsClickable)
{
	if (m_mapCtrlInfoTip.IsEmpty())
	{
		return NULL;
	}

	for (POSITION pos = m_mapCtrlInfoTip.GetStartPosition(); pos != NULL;)
	{
		HWND hwndCtrl = NULL;
		CBCGPControlInfoTip info;
		
		m_mapCtrlInfoTip.GetNextAssoc(pos, hwndCtrl, info);

		if (info.m_rect.PtInRect(point))
		{
			strInfo = info.m_strText;
			strDescription = info.m_strDescription;
			bIsClickable = info.m_bIsClickable;

			return hwndCtrl;
		}
	}

	return NULL;
}
//****************************************************************************
void CBCGPDlgImpl::EnableBackstageMode()
{
	m_bBackstageMode = TRUE;
	m_bTransparentStaticCtrls = FALSE;

#if (!defined _BCGSUITE_) && (!defined BCGP_EXCLUDE_RIBBON)
	m_bIsWhiteBackground = CBCGPVisualManager::GetInstance ()->IsRibbonBackstageWhiteBackground();
#else
	m_bIsWhiteBackground = TRUE;
#endif

	if (m_Dlg.GetSafeHwnd () == NULL)
	{
		return;
	}

	CWnd* pWndChild = m_Dlg.GetWindow (GW_CHILD);

	while (pWndChild != NULL)
	{
		ASSERT_VALID (pWndChild);

		pWndChild->SendMessage (BCGM_ONSETCONTROLBACKSTAGEMODE);
		pWndChild = pWndChild->GetNextWindow ();
	}
}
//****************************************************************************
BOOL CBCGPDlgImpl::GetPlacementSection(LPCTSTR lpszProfileName, CString& strSection)
{
	strSection.Empty();

	CString strClassName;

	CBCGPDialog* pDialog = DYNAMIC_DOWNCAST(CBCGPDialog, &m_Dlg);
	if (pDialog != NULL)
	{
		LPCTSTR lpID = pDialog->m_lpszTemplateName;
		if (lpID == NULL)
		{
			return FALSE;
		}

		if (IS_INTRESOURCE(lpID))
		{
			strClassName.Format(_T("%d"), lpID);
		}
		else
		{
			strClassName = lpID;
		}
	}
	else
	{
		CRuntimeClass* pClass = m_Dlg.GetRuntimeClass();
		if (pClass == NULL || pClass->m_lpszClassName == NULL)
		{
			return FALSE;
		}

		strClassName = pClass->m_lpszClassName;
		if (strClassName.IsEmpty() || strClassName == _T("CBCGPDialog") || strClassName == _T("CBCGPPropertySheet"))
		{
			return FALSE;
		}
	}

	CString strProfileName;
#ifndef _BCGSUITE_
	if (g_pWorkspace != NULL)
	{
		strProfileName = ::BCGPGetRegPath (g_pWorkspace->GetRegistryBase (), lpszProfileName);
	}
#else
	strProfileName = AFXGetRegPath (strDialogsProfile, lpszProfileName);
#endif

	strSection.Format(REG_DLG_SECTION_FMT_STR, (LPCTSTR)strProfileName, (LPCTSTR)strClassName);
	return TRUE;
}
//****************************************************************************
void CBCGPDlgImpl::EnableGesturePan(BOOL bEnable /*= TRUE*/)
{
	if (m_Dlg.GetSafeHwnd() == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	CBCGPGestureConfig gestureConfig;
		
	if (bEnable)
	{
		gestureConfig.EnablePan(TRUE, BCGP_GC_PAN_WITH_SINGLE_FINGER_HORIZONTALLY | BCGP_GC_PAN_WITH_SINGLE_FINGER_VERTICALLY | BCGP_GC_PAN_WITH_GUTTER | BCGP_GC_PAN_WITH_INERTIA);
	}
	else
	{
		gestureConfig.EnablePan(FALSE);
		gestureConfig.EnableZoom(FALSE);
		gestureConfig.EnableRotate(FALSE);
		gestureConfig.EnableTwoFingerTap(FALSE);
		gestureConfig.EnablePressAndTap(FALSE);
	}
		
	bcgpGestureManager.SetGestureConfig(m_Dlg.GetSafeHwnd(), gestureConfig);

	CBCGPGestureConfig gestureConfigChild;
	gestureConfigChild.EnablePan(bEnable);

	for (CWnd* pWndChild = m_Dlg.GetWindow(GW_CHILD); pWndChild != NULL; pWndChild = pWndChild->GetWindow(GW_HWNDNEXT))
	{
		if (pWndChild->SendMessage(BCGM_ONGETGESTURECONFIG, 0, (LPARAM)&gestureConfigChild) == 0)
		{
			bcgpGestureManager.SetGestureConfig(pWndChild->GetSafeHwnd(), gestureConfigChild);
		}
	}
}
//****************************************************************************
void CBCGPDlgImpl::SetEditBoxesVerticalAlignment(int nAlignment)
{
	if (m_Dlg.GetSafeHwnd () == NULL)
	{
		m_nEditBoxesVertAlignment = nAlignment;
		return;
	}

	m_nEditBoxesVertAlignment = -1;

	for (CWnd* pWndChild = m_Dlg.GetWindow(GW_CHILD); pWndChild != NULL; pWndChild = pWndChild->GetWindow(GW_HWNDNEXT))
	{
		CBCGPEdit* pWndEdit = DYNAMIC_DOWNCAST(CBCGPEdit, pWndChild);
		if (pWndEdit != NULL)
		{
			pWndEdit->SetVerticalAlignment(nAlignment);
		}
	}
}

#ifndef _BCGSUITE_

void CBCGPDlgImpl::SetEditBoxesSimplifiedBrowseIcon(BOOL bSet)
{
	if (m_Dlg.GetSafeHwnd () == NULL)
	{
		m_bEditBoxesSimplifiedIcon = bSet;
		return;
	}

	m_bEditBoxesSimplifiedIcon = -1;

	for (CWnd* pWndChild = m_Dlg.GetWindow(GW_CHILD); pWndChild != NULL; pWndChild = pWndChild->GetWindow(GW_HWNDNEXT))
	{
		CBCGPEdit* pWndEdit = DYNAMIC_DOWNCAST(CBCGPEdit, pWndChild);
		if (pWndEdit != NULL)
		{
			pWndEdit->SetSimplifiedBrowseIcon(bSet);
		}
	}
}

#endif

BOOL CBCGPDlgImpl::LoadPlacement(LPCTSTR lpszProfileName)
{
	if (!m_bLoadWindowPlacement || m_bWindowPlacementIsSet || m_Dlg.GetSafeHwnd () == NULL || (m_Dlg.GetStyle() & WS_CHILD) != 0)
	{
		return FALSE;
	}

	CString strSection(m_strWindowPlacementProfile);
	if (strSection.IsEmpty())
	{
		if (!GetPlacementSection(lpszProfileName, strSection))
		{
			return FALSE;
		}
	}

	WINDOWPLACEMENT wp;
	wp.length = sizeof (WINDOWPLACEMENT);

	if (!m_Dlg.GetWindowPlacement (&wp))
	{
		return FALSE;
	}

	CBCGPRegistrySP regSP;
	CBCGPRegistry& reg = regSP.Create (FALSE, TRUE);

	if (!reg.Open (strSection))
	{
		return FALSE;
	}

	CRect rectNormal;
	int nFlags = 0;
	int nShowCmd = SW_SHOWNORMAL;

	if (!reg.Read (_T("WindowRect"), rectNormal) ||
		!reg.Read (_T("Flags"), nFlags) ||
		!reg.Read (_T("ShowCmd"), nShowCmd))
	{
		return FALSE;
	}

	CRect rectDesktop;

	MONITORINFO mi;
	mi.cbSize = sizeof (MONITORINFO);
	if (GetMonitorInfo (MonitorFromPoint (rectNormal.TopLeft (), 
		MONITOR_DEFAULTTONEAREST), &mi))
	{
		rectDesktop = mi.rcWork;
	}
	else
	{
		::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectDesktop, 0);
	}

	CRect rectInter;
	if (rectInter.IntersectRect (&rectDesktop, &rectNormal))
	{
		wp.rcNormalPosition = rectNormal;
	}

	wp.showCmd = nShowCmd;

	if (m_Dlg.IsKindOf (RUNTIME_CLASS(CBCGPDialog)))
	{
		m_bWindowPlacementIsSet = ((CBCGPDialog&)m_Dlg).OnSetPlacement(wp);
	}
	else if (m_Dlg.IsKindOf (RUNTIME_CLASS(CBCGPPropertySheet)))
	{
		m_bWindowPlacementIsSet = ((CBCGPPropertySheet&)m_Dlg).OnSetPlacement(wp);
	}
	else
	{
		m_bWindowPlacementIsSet = SetPlacement (wp);
	}

	return TRUE;
}
//****************************************************************************
BOOL CBCGPDlgImpl::SavePlacement(LPCTSTR lpszProfileName)
{
	if (!m_bLoadWindowPlacement || m_Dlg.GetSafeHwnd () == NULL || (m_Dlg.GetStyle() & WS_CHILD) != 0)
	{
		return FALSE;
	}

	CString strSection(m_strWindowPlacementProfile);
	if (strSection.IsEmpty())
	{
		if (!GetPlacementSection(lpszProfileName, strSection))
		{
			return FALSE;
		}
	}

	WINDOWPLACEMENT wp;
	wp.length = sizeof (WINDOWPLACEMENT);

	if (!m_Dlg.GetWindowPlacement (&wp))
	{
		return FALSE;
	}

	CBCGPRegistrySP regSP;
	CBCGPRegistry& reg = regSP.Create (FALSE, FALSE);

	if (!reg.CreateKey (strSection))
	{
		return FALSE;
	}

	m_bWindowPlacementIsSet = FALSE;

	CRect rectWnd = wp.rcNormalPosition;

	CRect rectDesktop;

	MONITORINFO mi;
	mi.cbSize = sizeof (MONITORINFO);
	if (GetMonitorInfo (MonitorFromWindow (m_Dlg.GetSafeHwnd (), 
		MONITOR_DEFAULTTONEAREST), &mi))
	{
		CRect rectWork(mi.rcWork);
		rectDesktop = rectWork;
		
		rectDesktop.left = abs(rectDesktop.left - mi.rcMonitor.left);
		rectDesktop.top = abs(rectDesktop.top - mi.rcMonitor.top);
		rectDesktop.right = rectDesktop.left + rectWork.Width ();
		rectDesktop.bottom = rectDesktop.top + rectWork.Height ();
	}
	else
	{
		::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectDesktop, 0);
	}

	rectWnd.OffsetRect(rectDesktop.left, rectDesktop.top);

#ifndef _BCGSUITE_
	if (g_pWorkspace != NULL)
	{
		if (!m_Dlg.IsZoomed() && !m_Dlg.IsIconic() && g_pWorkspace->m_bIsAeroSnapPlacement)
		{
			// GetWindowPlacement doesn't work correctly in AeroSnap mode:
			m_Dlg.GetWindowRect(rectWnd);
		}
	}
#endif

	return	reg.Write (_T("WindowRect"), rectWnd) &&
			reg.Write (_T("Flags"), (int)wp.flags) &&
			reg.Write (_T("ShowCmd"), (int)wp.showCmd);
}
//****************************************************************************
BOOL CBCGPDlgImpl::SetPlacement(WINDOWPLACEMENT& wp)
{
	if (m_bWindowPlacementIsSet)
	{
		return TRUE;
	}

	if (m_Dlg.GetSafeHwnd () == NULL)
	{
		return FALSE;
	}

	if (wp.showCmd != SW_SHOWNORMAL)
	{
		m_Dlg.SetWindowPlacement (&wp);
	}
	else
	{
		CRect rect(wp.rcNormalPosition);

		if (IsLayoutEnabled())
		{
			m_Dlg.SetWindowPos (&CWnd::wndTop, rect.left, rect.top, rect.Width (), rect.Height (),
 				SWP_NOZORDER | SWP_NOOWNERZORDER);
		}
		else
		{
			m_Dlg.SetWindowPos (&CWnd::wndTop, rect.left, rect.top, -1, -1,
				SWP_NOZORDER | SWP_NOOWNERZORDER | SWP_NOSIZE);
		}
	}

	return TRUE;
}
//****************************************************************************
void CBCGPDlgImpl::AdjustShadow(BOOL bActive)
{
#ifndef _BCGSUITE_
	if (m_pShadow != NULL)
	{
		m_pShadow->SetVisible((bActive || (m_Dlg.m_nFlags & WF_STAYACTIVE) == WF_STAYACTIVE) && m_Dlg.IsWindowEnabled());
	}
#else
	UNREFERENCED_PARAMETER(bActive);
#endif
}
//********************************************************************************
void CBCGPDlgImpl::OnSysCommand(UINT nID, LPARAM /*lParam*/)
{
#ifndef _BCGSUITE_
	if (m_pShadow != NULL && ((nID & 0xFFF0) == SC_MINIMIZE))
	{
		m_pShadow->Show(FALSE);
	}
#else
	UNREFERENCED_PARAMETER(nID);
#endif
}
//********************************************************************************
void CBCGPDlgImpl::CreateTooltipInfo()
{
	CBCGPTooltipManager::CreateToolTip(m_pToolTipInfo, &m_Dlg, BCGP_TOOLTIP_TYPE_DEFAULT);

	if (m_pToolTipInfo->GetSafeHwnd() != NULL)
	{
		m_pToolTipInfo->SetMaxTipWidth(nTooltipMaxWidth);
		
		CRect rectDummy (0, 0, 0, 0);
		m_pToolTipInfo->AddTool(&m_Dlg, LPSTR_TEXTCALLBACK, &rectDummy, nInfoToolTipID);
		
		UpdateToolTipsRect();
	}

	if (!m_mapCtrlInfoTipDelayed.IsEmpty())
	{
		for (POSITION pos = m_mapCtrlInfoTipDelayed.GetStartPosition(); pos != NULL;)
		{
			UINT nCtrlID = 0;
			CBCGPControlInfoTip info;
			
			m_mapCtrlInfoTipDelayed.GetNextAssoc(pos, nCtrlID, info);
			
			if ((info.m_hWnd = m_Dlg.GetDlgItem(nCtrlID)->GetSafeHwnd()) != NULL)
			{
				m_mapCtrlInfoTip[info.m_hWnd] = info;
			}
		}

		m_mapCtrlInfoTipDelayed.RemoveAll();
	}
}
//********************************************************************************
BOOL CBCGPDlgImpl::OnNeedTipText(UINT /*id*/, NMHDR* pNMH, LRESULT* /*pResult*/)
{
	static CString strTipText;

	if (m_mapCtrlInfoTip.IsEmpty() || m_pToolTipInfo->GetSafeHwnd () == NULL || pNMH->hwndFrom != m_pToolTipInfo->GetSafeHwnd ())
	{
		return FALSE;
	}
	
	if (CBCGPPopupMenu::GetActiveMenu () != NULL)
	{
		return FALSE;
	}
	
	LPNMTTDISPINFO	pTTDispInfo	= (LPNMTTDISPINFO) pNMH;
	ASSERT((pTTDispInfo->uFlags & TTF_IDISHWND) == 0);
	
	CPoint point;
	
	::GetCursorPos (&point);
	m_Dlg.ScreenToClient (&point);

	CString strDescription;
	BOOL bIsClickable = FALSE;

	if (GetControlInfoTipFromPoint(point, strTipText, strDescription, bIsClickable) == NULL)
	{
		return FALSE;
	}
	
	if (strTipText.IsEmpty ())
	{
		return TRUE;
	}
	
	CBCGPToolTipCtrl* pToolTip = DYNAMIC_DOWNCAST(CBCGPToolTipCtrl, m_pToolTipInfo);
	if (pToolTip != NULL)
	{
		ASSERT_VALID(pToolTip);
		pToolTip->SetDescription(strDescription);
	}
	
	pTTDispInfo->lpszText = const_cast<LPTSTR> ((LPCTSTR) strTipText);
	return TRUE;
}
//***************************************************************************************************************
void CBCGPDlgImpl::UpdateToolTipsRect()
{
	if (m_pToolTipInfo->GetSafeHwnd () != NULL)
	{
		CRect rectToolTipClient;
		m_Dlg.GetClientRect (rectToolTipClient);

		m_pToolTipInfo->SetToolRect(&m_Dlg, nInfoToolTipID, rectToolTipClient);
	}
}
//***************************************************************************************************************
void CBCGPDlgImpl::SetUnderlineKeyboardShortcuts(BOOL bSet, BOOL bRedraw)
{
#ifndef _BCGSUITE_
	if (globalData.m_bSysUnderlineKeyboardShortcuts)
	{
		// Always underlined, do nothing
		return;
	}

	m_Dlg.SendMessageToDescendants(BCGM_ONCHANGEUNDERLINE, (WPARAM)bSet, (LPARAM)bRedraw, TRUE, TRUE);
#else
	UNREFERENCED_PARAMETER(bSet);
	UNREFERENCED_PARAMETER(bRedraw);
#endif
}
//***************************************************************************************************************
BOOL CBCGPDlgImpl::OnSetCursor()
{
	CPoint point;
	
	::GetCursorPos (&point);
	m_Dlg.ScreenToClient(&point);
	
	CString strTipText;
	CString strDescription;
	BOOL bIsClickable = FALSE;

	if (GetControlInfoTipFromPoint(point, strTipText, strDescription, bIsClickable) == NULL || !bIsClickable)
	{
		return FALSE;
	}

	::SetCursor (globalData.GetHandCursor ());
	return TRUE;
}
//**************************************************************************************************************
void CBCGPDlgImpl::DrawBackgroundImage(CDC* pDC, CRect rectClient)
{
	if (!m_Background.IsValid())
	{
		return;
	}

	CSize sizeBkgrBitmap = m_Background.GetImageSize();
	
	if (m_BkgrLocation != 0 /*BACKGR_TILE*/)
	{
		CPoint ptImage = rectClient.TopLeft ();
		
		switch (m_BkgrLocation)
		{
		case 1 /*BACKGR_TOPLEFT*/:
			break;
			
		case 2 /*BACKGR_TOPRIGHT*/:
			ptImage.x = rectClient.right - sizeBkgrBitmap.cx;
			break;
			
		case 3 /*BACKGR_BOTTOMLEFT*/:
			ptImage.y = rectClient.bottom - sizeBkgrBitmap.cy;
			break;
			
		case 4 /*BACKGR_BOTTOMRIGHT*/:
			ptImage.x = rectClient.right - sizeBkgrBitmap.cx;
			ptImage.y = rectClient.bottom - sizeBkgrBitmap.cy;
			break;
		}
		
		m_Background.DrawEx(pDC, CRect(ptImage, sizeBkgrBitmap), 0);
	}
	else
	{
		// Tile background image:
		for (int x = rectClient.left; x < rectClient.Width (); x += sizeBkgrBitmap.cx)
		{
			for (int y = rectClient.top; y < rectClient.Height (); y += sizeBkgrBitmap.cy)
			{
				m_Background.DrawEx(pDC, CRect(CPoint(x, y), sizeBkgrBitmap), 0);
			}
		}
	}
}
//**************************************************************************************************************
void CBCGPDlgImpl::AutoResizeControls(const CList<UINT,UINT>* plstNonAutoResizedItems/* = NULL*/)
{
	m_bAutoResizeCtrls = TRUE;

	m_lstNonAutoResizedItems.RemoveAll();
	if (plstNonAutoResizedItems != NULL)
	{
		m_lstNonAutoResizedItems.AddTail((CList<UINT,UINT>*)plstNonAutoResizedItems);
	}

	DoAutoResizeControls();
}
//**************************************************************************************************************
void CBCGPDlgImpl::DoAutoResizeControls()
{
	if (!m_bAutoResizeCtrls || m_Dlg.GetSafeHwnd() == NULL)
	{
		return;
	}

	for (CWnd* pWndChild = m_Dlg.GetWindow(GW_CHILD); pWndChild != NULL; pWndChild = pWndChild->GetWindow(GW_HWNDNEXT))
	{
		if (m_lstNonAutoResizedItems.Find(pWndChild->GetDlgCtrlID ()) != NULL)
		{
			continue;
		}
		
		CBCGPButton* pButton = DYNAMIC_DOWNCAST(CBCGPButton, pWndChild);
		if (pButton != NULL && (pButton->IsCheckBox() || pButton->IsRadioButton()))
		{
			pButton->SizeToContent();
		}
		
		CBCGPStatic* pStatic = DYNAMIC_DOWNCAST(CBCGPStatic, pWndChild);
		if (pStatic != NULL)
		{
			pStatic->SizeToContent();
		}
	}
}
//**************************************************************************************************************
BOOL CBCGPDlgImpl::ProcessFocusEditAccelerators(UINT nChar)
{
	if (m_Dlg.GetSafeHwnd() == NULL || (m_Dlg.GetStyle() & WS_CHILD) == NULL)
	{
		return FALSE;
	}

	CWnd* pWndFocus = CWnd::GetFocus();
	
	if (pWndFocus->GetSafeHwnd() == NULL || !m_Dlg.IsChild(pWndFocus))
	{
		return FALSE;
	}

	#define MAX_CLASS_NAME 255
	TCHAR lpszClassName[MAX_CLASS_NAME + 1];
	
	::GetClassName (pWndFocus->GetSafeHwnd(), lpszClassName, MAX_CLASS_NAME);

	if (_tcscmp(lpszClassName, _T("Edit")) == 0)
	{
		return globalUtils.ProcessCtrlEditAccelerators(pWndFocus, nChar, (pWndFocus->GetStyle() & ES_READONLY) == 0);
	}

	return FALSE;
}


/////////////////////////////////////////////////////////////////////////////
// CBCGPParentEditPtr

void CBCGPParentEditPtr::ClosePopupDlg(LPCTSTR lpszEditValue, BOOL bOK, DWORD_PTR dwUserData)
{
	if (m_pParentEdit->GetSafeHwnd() == NULL)
	{
		ASSERT(FALSE);
		return;
	}
	
	if (lpszEditValue != NULL && bOK)
	{
		m_pParentEdit->SetWindowText(lpszEditValue);
	}
	
	m_pParentEdit->ClosePopupDlg(bOK, dwUserData);
}
//********************************************************************************
void CBCGPParentEditPtr::OnDestroyPopupDlg()
{
	if (m_pParentEdit != NULL)
	{
		m_pParentEdit->m_pPopupDlg = NULL;
		m_pParentEdit->RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	}
}
//********************************************************************************
CWnd* CBCGPParentEditPtr::GetParentArea(CRect& rectParentBtn)
{
	if (m_pParentEdit->GetSafeHwnd() == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}
	
	m_pParentEdit->GetWindowRect(rectParentBtn);
	m_pParentEdit->ScreenToClient(rectParentBtn);
	
	return m_pParentEdit;
}
//********************************************************************************
void CBCGPParentEditPtr::GetParentText(CString& strEditValue)
{
	if (m_pParentEdit->GetSafeHwnd() != NULL)
	{
		m_pParentEdit->GetWindowText(strEditValue);
	}
}
//*****************************************************************************
void CBCGPParentEditPtr::SetParentText(const CString& strEditValue)
{
	if (m_pParentEdit->GetSafeHwnd() != NULL)
	{
		m_pParentEdit->SetWindowText(strEditValue);
	}
}

#ifndef BCGP_EXCLUDE_GRID_CTRL

/////////////////////////////////////////////////////////////////////////////
// CBCGPParentGridItemPtr

void CBCGPParentGridItemPtr::ClosePopupDlg(LPCTSTR lpszEditValue, BOOL bOK, DWORD_PTR dwUserData)
{
	if (m_pParentGridItem == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	m_pParentGridItem->ClosePopupDlg(lpszEditValue, bOK, dwUserData);
}
//*****************************************************************************
void CBCGPParentGridItemPtr::OnDestroyPopupDlg()
{
	if (m_pParentGridItem != NULL)
	{
		m_pParentGridItem->m_pPopupDlg = NULL;
	}
}
//*****************************************************************************
CWnd* CBCGPParentGridItemPtr::GetParentArea(CRect& rectParentBtn)
{
	if (m_pParentGridItem == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}
	
	rectParentBtn = m_pParentGridItem->GetButtonRect ();
	return m_pParentGridItem->GetOwnerList ();
}
//*****************************************************************************
void CBCGPParentGridItemPtr::GetParentText(CString& strEditValue)
{
	if (m_pParentGridItem != NULL)
	{
		ASSERT_VALID(m_pParentGridItem);

		strEditValue = m_pParentGridItem->FormatItem ();
	}
}
//*****************************************************************************
void CBCGPParentGridItemPtr::SetParentText(const CString& strEditValue)
{
	if (m_pParentGridItem != NULL)
	{
		ASSERT_VALID(m_pParentGridItem);
		m_pParentGridItem->SetValue((LPCTSTR)strEditValue);
	}
}

#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPParentMenuButtonPtr

void CBCGPParentMenuButtonPtr::ClosePopupDlg(LPCTSTR lpszEditValue, BOOL bOK, DWORD_PTR dwUserData)
{
	if (m_pParentMenuButton->GetSafeHwnd() == NULL)
	{
		ASSERT(FALSE);
		return;
	}
	
	if (lpszEditValue != NULL && bOK)
	{
		m_pParentMenuButton->SetWindowText(lpszEditValue);
	}
	
	m_pParentMenuButton->ClosePopupDlg(bOK, dwUserData);
}
//********************************************************************************
void CBCGPParentMenuButtonPtr::OnDestroyPopupDlg()
{
	if (m_pParentMenuButton != NULL)
	{
		m_pParentMenuButton->m_pPopupDlg = NULL;
		m_pParentMenuButton->RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	}
}
//********************************************************************************
CWnd* CBCGPParentMenuButtonPtr::GetParentArea(CRect& rectParentBtn)
{
	if (m_pParentMenuButton->GetSafeHwnd() == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}
	
	m_pParentMenuButton->GetWindowRect(rectParentBtn);
	m_pParentMenuButton->ScreenToClient(rectParentBtn);
	
	return m_pParentMenuButton;
}
//********************************************************************************
void CBCGPParentMenuButtonPtr::GetParentText(CString& strEditValue)
{
	if (m_pParentMenuButton->GetSafeHwnd() != NULL)
	{
		m_pParentMenuButton->GetWindowText(strEditValue);
	}
}
//*****************************************************************************
void CBCGPParentMenuButtonPtr::SetParentText(const CString& strEditValue)
{
	if (m_pParentMenuButton->GetSafeHwnd() != NULL)
	{
		m_pParentMenuButton->SetWindowText(strEditValue);
	}
}

#if (!defined _BCGSUITE_) && (!defined BCGP_EXCLUDE_RIBBON)

void CBCGPParentRibbonButtonPtr::ClosePopupDlg(LPCTSTR lpszEditValue, BOOL bOK, DWORD_PTR dwUserData)
{
	if (m_pParentRibbonButton == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	m_pParentRibbonButton->ClosePopupDlg(lpszEditValue, bOK, dwUserData);
}
//*****************************************************************************
void CBCGPParentRibbonButtonPtr::OnDestroyPopupDlg()
{
	if (m_pParentRibbonButton != NULL)
	{
		m_pParentRibbonButton->m_pPopupDlg = NULL;
		m_pParentRibbonButton->Redraw();
	}
}
//*****************************************************************************
CWnd* CBCGPParentRibbonButtonPtr::GetParentArea(CRect& rectParentBtn)
{
	if (m_pParentRibbonButton == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}

	CWnd* pWndParent = m_pParentRibbonButton->GetParentWnd();
	if (pWndParent->GetSafeHwnd() == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}
	
	rectParentBtn = m_pParentRibbonButton->GetMenuRect();
	
	if (rectParentBtn.IsRectEmpty())
	{
		rectParentBtn = m_pParentRibbonButton->GetRect();
	}

	return pWndParent;
}
//*****************************************************************************
void CBCGPParentRibbonButtonPtr::GetParentText(CString& strEditValue)
{
	if (m_pParentRibbonButton != NULL)
	{
		ASSERT_VALID(m_pParentRibbonButton);
		m_pParentRibbonButton->OnGetPopupDlgText(strEditValue);
	}
}
//*****************************************************************************
void CBCGPParentRibbonButtonPtr::SetParentText(const CString& strEditValue)
{
	if (m_pParentRibbonButton != NULL)
	{
		ASSERT_VALID(m_pParentRibbonButton);
		m_pParentRibbonButton->OnSetPopupDlgText(strEditValue);
	}
}

#endif
