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
// BCGPWindowsNavigator.cpp : implementation file
//

#include "stdafx.h"
#include "bcgcbpro.h"
#include "BCGPFrameImpl.h"
#include "BCGPPopupWindow.h"
#include "BCGPDockingControlBar.h"
#include "BCGPDrawManager.h"
#include "BCGPMDIFrameWnd.h"
#include "BCGPMDIChildWnd.h"
#include "BCGPGlobalUtils.h"
#include "MenuImages.h"
#include "BCGPToolbarMenuButton.h"
#include "BCGPWindowsNavigator.h"
#include "BCGPTabbedControlBar.h"
#include "BCGPKeyboardManager.h"
#include "BCGPVisualManager2007.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#ifndef BCGP_EXCLUDE_POPUP_WINDOW

#define ID_CHECK_KBD_STATE	1

class CBCGPWindowsNavigatorObject : public CObject
{
	friend CBCGPWindowsNavigator;
	
	CBCGPWindowsNavigatorObject(BOOL bIsPane, const CString& strName, HICON hIcon, HWND hwnd, LPCTSTR lpszPath = NULL, BOOL bAutoDestroyIcon = FALSE) :
		m_bIsPane(bIsPane),
		m_strName(strName),
		m_hIcon(hIcon),
		m_bAutoDestroyIcon(bAutoDestroyIcon),
		m_hWnd(hwnd),
		m_strPath(lpszPath == NULL ? _T("") : lpszPath)
	{
		m_rect.SetRectEmpty();
		m_nScroll = 0;
	}
	
	virtual ~CBCGPWindowsNavigatorObject()
	{
		if (m_bAutoDestroyIcon && m_hIcon != NULL)
		{
			::DestroyIcon(m_hIcon);
		}
	}
	
	void OnDraw(CDC* pDC, BOOL bIsSelected);
	
	const BOOL		m_bIsPane;
	const CString	m_strName;
	const HICON		m_hIcon;
	const BOOL		m_bAutoDestroyIcon;
	const HWND		m_hWnd;
	const CString	m_strPath;
	
	CRect	m_rect;
	int		m_nScroll;
};

void CBCGPWindowsNavigatorObject::OnDraw(CDC* pDC, BOOL bIsSelected)
{
	CRect rect = m_rect;
	if (rect.IsRectEmpty())
	{
		return;
	}

	ASSERT_VALID(pDC);

	if (m_nScroll != 0)
	{
		CBCGPMenuImages::DrawByColor(pDC, m_nScroll < 0 ? CBCGPMenuImages::IdArowUpLarge : CBCGPMenuImages::IdArowDownLarge, rect, pDC->GetTextColor(), FALSE);
		return;
	}

	COLORREF clrTextOld = (COLORREF)-1;

	if (bIsSelected)
	{
		CRgn rgn;

		if (CBCGPVisualManager::GetInstance()->IsKindOf (RUNTIME_CLASS(CBCGPVisualManager2007)) &&
			CBCGPVisualManager2007::GetStyle() == CBCGPVisualManager2007::VS2007_Aqua)
		{
			rgn.CreateRoundRectRgn (rect.left, rect.top, rect.right + 1, rect.bottom + 1, 4, 4);
			pDC->SelectClipRgn(&rgn);
		}

		CBCGPToolbarMenuButton dummy;

		COLORREF clrText = CBCGPVisualManager::GetInstance()->GetMenuItemTextColor(&dummy, TRUE, FALSE);

		CBCGPVisualManager::GetInstance()->OnHighlightMenuItem(pDC, &dummy, rect, clrText);

		if (clrText != (COLORREF)-1)
		{
			clrTextOld = pDC->SetTextColor(clrText);
		}

		if (rgn.GetSafeHandle() != NULL)
		{
			pDC->SelectClipRgn(NULL);
		}
	}

	int nHorzOffset = globalUtils.ScaleByDPI(5);

	if (m_hIcon != NULL)
	{
		::DrawIconEx (pDC->GetSafeHdc (), 
			rect.left + nHorzOffset,
			rect.top + max(0, (rect.Height() - globalData.m_sizeSmallIcon.cy) / 2),
			m_hIcon, globalData.m_sizeSmallIcon.cx, globalData.m_sizeSmallIcon.cy, 0, NULL,
			DI_NORMAL);
	}

	CRect rectText = rect;
	rectText.left += (globalData.m_sizeSmallIcon.cx + 2 * nHorzOffset);
	rectText.right -= nHorzOffset;

	pDC->DrawText(m_strName, rectText, DT_LEFT | DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);

	if (clrTextOld != (COLORREF)-1)
	{
		pDC->SetTextColor(clrTextOld);
	}
}

static BYTE GetVirt(BOOL bIgnoreShift)
{
	BYTE fVirt = 0;
	
	if (::GetAsyncKeyState (VK_CONTROL) & 0x8000)
	{
		fVirt |= FCONTROL;
	}
	
	if (::GetAsyncKeyState (VK_MENU) & 0x8000)
	{
		fVirt |= FALT;
	}

	if (!bIgnoreShift)
	{
		if (::GetAsyncKeyState (VK_SHIFT) & 0x8000)
		{
			fVirt |= FSHIFT;
		}
	}

	return fVirt;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPWindowsNavigator dialog

IMPLEMENT_DYNCREATE(CBCGPWindowsNavigator, CBCGPPopupDlg)

CBCGPWindowsNavigator::CBCGPWindowsNavigator()
{
	//{{AFX_DATA_INIT(CBCGPWindowsNavigator)
	//}}AFX_DATA_INIT

	m_nScrollOffset1 = m_nScrollOffset2 = 0;
	m_pSel = NULL;
	m_nMDIColumns = 1;
	m_sizeExtra = CSize(0, 0);
	m_bIsActivePane = FALSE;
	
	m_nInitialVirtKeys = GetVirt(TRUE);

	m_nNextChar = VK_TAB;
	m_nNextVirtKeys = FCONTROL;
	m_nPrevChar = VK_TAB;
	m_nPrevVirtKeys = FCONTROL | FSHIFT;
}
//*********************************************************************************************************
CBCGPWindowsNavigator::~CBCGPWindowsNavigator()
{
	int i = 0;

	for (i = 0; i < (int)m_arDockingPanes.GetSize(); i++)
	{
		CBCGPWindowsNavigatorObject* pPane = m_arDockingPanes[i];
		ASSERT_VALID(pPane);

		delete pPane;
	}
	
	for (i = 0; i < (int)m_arMDIChildren.GetSize(); i++)
	{
		CBCGPWindowsNavigatorObject* pDoc = m_arMDIChildren[i];
		ASSERT_VALID(pDoc);

		delete pDoc;
	}

	ReleaseCapture();

	CBCGPFrameImpl* pFrameImpl = (CBCGPFrameImpl*)m_lCustomParam;
	ASSERT(pFrameImpl != NULL);

	pFrameImpl->m_hwndWindowsNavigator = NULL;
}
//*********************************************************************************************************
void CBCGPWindowsNavigator::DoDataExchange(CDataExchange* pDX)
{
	CBCGPPopupDlg::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CBCGPWindowsNavigator)
	DDX_Control(pDX, IDD_BCGBAR_RES_LABEL3, m_wndPath);
	DDX_Control(pDX, IDC_BCGBARRES_LIST_VIEWS, m_wndViewListLocation);
	DDX_Control(pDX, IDC_BCGBARRES_LIST, m_wndPaneListLocation);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CBCGPWindowsNavigator, CBCGPPopupDlg)
	//{{AFX_MSG_MAP(CBCGPWindowsNavigator)
	ON_WM_TIMER()
	ON_WM_KEYDOWN()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPWindowsNavigator message handlers

BOOL CBCGPWindowsNavigator::OnInitDialog() 
{
	CBCGPPopupDlg::OnInitDialog();

	m_wndPath.SetFont(&globalData.fontRegular);
	m_wndPath.SetWindowText(_T(""));
	m_wndPath.ModifyStyle(0, SS_PATHELLIPSIS);

	GetDlgItem(IDC_BCGBARRES_LABEL1)->SetFont(&globalData.fontDefaultGUIBold);
	GetDlgItem(IDC_BCGBARRES_LABEL2)->SetFont(&globalData.fontDefaultGUIBold);
	
	CBCGPFrameImpl* pFrameImpl = (CBCGPFrameImpl*)m_lCustomParam;
	ASSERT(pFrameImpl != NULL);

	CWnd* pWndOwner = NULL;
	CBCGPPopupWindow* pWndParent = DYNAMIC_DOWNCAST(CBCGPPopupWindow, GetParent());
	if (pWndParent->GetSafeHwnd() != NULL)
	{
		pWndOwner = pWndParent->GetOwnerWnd();
	}

	//-------------------------
	// Fill docking panes list:
	//-------------------------
	CObList lstDockingPanes;
	
	if (pFrameImpl->m_pDockManager != NULL)
	{
		ASSERT_VALID(pFrameImpl->m_pDockManager);
		
		pFrameImpl->m_pDockManager->GetControlBarList(lstDockingPanes, TRUE, NULL, TRUE);
		
		for (POSITION pos = lstDockingPanes.GetHeadPosition(); pos != NULL;)
		{
			CBCGPDockingControlBar* pBar = DYNAMIC_DOWNCAST(CBCGPDockingControlBar, lstDockingPanes.GetNext(pos));
			if (pBar != NULL)
			{
				if ((pBar->IsVisible() || pBar->IsAutoHideMode()) && DYNAMIC_DOWNCAST(CBCGPTabbedControlBar, pBar) == NULL)
				{
					CString strName;
					CString strDetails;
					HICON hIcon = NULL;
					BOOL bAutoDestroyIcon = FALSE;
					
					pBar->GetPaneInfo(strName, strDetails, hIcon, bAutoDestroyIcon);
					
					CBCGPWindowsNavigatorObject* pObj = new CBCGPWindowsNavigatorObject(TRUE, strName, hIcon, pBar->GetSafeHwnd(), strDetails, bAutoDestroyIcon);
					m_arDockingPanes.Add(pObj);
					
					if (pBar->IsChild (GetFocus()))
					{
						SetSel(pObj);
						m_bIsActivePane = TRUE;
					}
				}
			}
		}
	}
	
	//------------------------
	// Fill MDI children list:
	//------------------------
	HWND hwndMDIActive = NULL;

	const CList<CFrameWnd*, CFrameWnd*>& lstFrames = CBCGPFrameImpl::GetFrameList();
	for (POSITION pos = lstFrames.GetHeadPosition(); pos != NULL;)
	{
		CBCGPMDIFrameWnd* pMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, lstFrames.GetNext(pos));
		if (pMDIFrame->GetSafeHwnd() != NULL)
		{
			ASSERT(pMDIFrame->m_hWndMDIClient != NULL);

			if (pMDIFrame->GetSafeHwnd() == pWndOwner->GetSafeHwnd())
			{
				hwndMDIActive = pMDIFrame->MDIGetActive()->GetSafeHwnd();
			}

			HWND hwndMDIChild = ::GetWindow (pMDIFrame->m_hWndMDIClient, GW_CHILD);
			
			while (hwndMDIChild != NULL)
			{
				CBCGPMDIChildWnd* pMDIChildFrame = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, CWnd::FromHandle (hwndMDIChild));
				
				if (pMDIChildFrame != NULL && pMDIChildFrame->IsWindowVisible() && pMDIChildFrame->CanShowOnWindowsList())
				{
					CString strName;
					CString strPath;
					HICON hIcon = NULL;
					BOOL bAutoDestroyIcon = FALSE;
					
					pMDIChildFrame->GetPaneInfo(strName, strPath, hIcon, bAutoDestroyIcon);

					CBCGPWindowsNavigatorObject* pObj = new CBCGPWindowsNavigatorObject(FALSE, strName, hIcon, pMDIChildFrame->GetSafeHwnd(), strPath, bAutoDestroyIcon);
					m_arMDIChildren.Add(pObj);
				}
				
				hwndMDIChild = ::GetWindow (hwndMDIChild, GW_HWNDNEXT);
			}
		}
	}

	if (IsEmptyDlg())
	{
		CWnd* pWndParentCurr = GetParent();
		if (pWndParentCurr->GetSafeHwnd() != NULL)
		{
			pWndParentCurr->PostMessage(WM_CLOSE);
		}

		return TRUE;
	}

	if (m_pSel == NULL && m_arMDIChildren.GetSize() > 0)
	{
		CBCGPWindowsNavigatorObject* pSel = m_arMDIChildren[m_arMDIChildren.GetSize() == 1 ? 0 : (pFrameImpl->m_bIsPrevNavigation ? m_arMDIChildren.GetSize() - 2 : 1)];
		ASSERT_VALID(pSel);

		if (pSel->m_hWnd == hwndMDIActive && m_arMDIChildren.GetSize() > 1)
		{
			pSel = m_arMDIChildren[pFrameImpl->m_bIsPrevNavigation ? m_arMDIChildren.GetSize() - 1 : 0];
			ASSERT_VALID(pSel);
		}

		SetSel(pSel);
		m_bIsActivePane = FALSE;
	}

	//---------------------------------
	// Resize the dialog (if required):
	//---------------------------------
	CRect rectPanes;
	m_wndPaneListLocation.GetWindowRect(rectPanes);
	ScreenToClient(rectPanes);

	int nColumnWidth = rectPanes.Width();

	CRect rectDocs;
	m_wndViewListLocation.GetWindowRect(rectDocs);
	ScreenToClient(rectDocs);

	CRect rectPath;
	m_wndPath.GetWindowRect(rectPath);
	ScreenToClient(rectPath);

	int nMDIColumns = 0;

	while (m_nMDIColumns < 5)
	{
		if (!ReposItems(&nMDIColumns) || m_nMDIColumns == 4)
		{
			break;
		}

		m_sizeExtra.cx += nColumnWidth;
		m_sizeExtra.cy += nColumnWidth / 3;

		rectPanes.bottom += nColumnWidth / 3;

		rectDocs.right += nColumnWidth;
		rectDocs.bottom += nColumnWidth / 3;

		rectPath.OffsetRect(0, nColumnWidth / 3);
		rectPath.right += nColumnWidth;

		m_wndPaneListLocation.SetWindowPos(NULL, -1, -1, rectPanes.Width(), rectPanes.Height(),
			SWP_NOZORDER | SWP_NOMOVE | SWP_NOACTIVATE);

		m_wndViewListLocation.SetWindowPos(NULL, -1, -1, rectDocs.Width(), rectDocs.Height(),
			SWP_NOZORDER | SWP_NOMOVE | SWP_NOACTIVATE);

		m_wndPath.SetWindowPos(NULL, rectPath.left, rectPath.top, rectPath.Width(), rectPath.Height(),
			SWP_NOZORDER | SWP_NOACTIVATE);

		m_nMDIColumns++;
	}

	if (nMDIColumns < m_nMDIColumns)
	{
		int nDX = (m_nMDIColumns - nMDIColumns) * nColumnWidth;

		m_sizeExtra.cx -= nDX;
		rectDocs.right -= nDX;
		rectPath.right -= nDX;
		
		m_wndPaneListLocation.SetWindowPos(NULL, -1, -1, rectPanes.Width(), rectPanes.Height(),
			SWP_NOZORDER | SWP_NOMOVE | SWP_NOACTIVATE);
		
		m_wndViewListLocation.SetWindowPos(NULL, -1, -1, rectDocs.Width(), rectDocs.Height(),
			SWP_NOZORDER | SWP_NOMOVE | SWP_NOACTIVATE);
		
		m_wndPath.SetWindowPos(NULL, -1, -1, rectPath.Width(), rectPath.Height(),
			SWP_NOZORDER | SWP_NOMOVE | SWP_NOACTIVATE);
	}

	//---------------------
	// Find next/prev keys:
	//---------------------
	UINT nCmdIDNext = pFrameImpl->m_nWindowsNextNavigationCmdID;
	UINT nCmdIDPrev = pFrameImpl->m_nWindowsPrevNavigationCmdID;

	if (nCmdIDNext != 0 && nCmdIDPrev != 0)
	{
		ACCEL accel = { 0,0,0 };
		CString strLabel;

		if (CBCGPKeyboardManager::FindDefaultAccelerator(nCmdIDNext, strLabel, pFrameImpl->m_pFrame, TRUE, &accel) ||
			CBCGPKeyboardManager::FindDefaultAccelerator(nCmdIDNext, strLabel, pFrameImpl->m_pFrame->GetActiveFrame (), FALSE, &accel))
		{
			m_nNextChar = accel.key;
			m_nNextVirtKeys = (BYTE)(accel.fVirt & (FSHIFT | FCONTROL | FALT));
		}

		if (CBCGPKeyboardManager::FindDefaultAccelerator(nCmdIDPrev, strLabel, pFrameImpl->m_pFrame, TRUE, &accel) ||
			CBCGPKeyboardManager::FindDefaultAccelerator(nCmdIDPrev, strLabel, pFrameImpl->m_pFrame->GetActiveFrame (), FALSE, &accel))
		{
			m_nPrevChar = accel.key;
			m_nPrevVirtKeys = (BYTE)(accel.fVirt & (FSHIFT | FCONTROL | FALT));
		}
	}

	SetCapture();
	SetTimer(ID_CHECK_KBD_STATE, 10, NULL);

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
//*********************************************************************************************************
void CBCGPWindowsNavigator::SetSel(CBCGPWindowsNavigatorObject* pSel)
{
	m_pSel = pSel;

	if (m_pSel != NULL)
	{
		if (m_pSel->m_bIsPane)
		{
			m_wndPath.ModifyStyle(SS_PATHELLIPSIS, SS_ENDELLIPSIS);
		}
		else
		{
			m_wndPath.ModifyStyle(SS_ENDELLIPSIS, SS_PATHELLIPSIS);
		}
	}

	CString strPath = m_pSel == NULL ? _T("") : m_pSel->m_strPath;
	
	CString strOldPath;
	m_wndPath.GetWindowText(strOldPath);

	if (strOldPath != strPath)
	{
		m_wndPath.SetWindowText(strPath);
	}

	if (m_pParentPopup != NULL)
	{
		ASSERT_VALID(m_pParentPopup);

		m_pParentPopup->SetWindowText(m_pSel == NULL ? _T("") : m_pSel->m_strName);
		m_pParentPopup->SetIcon(m_pSel == NULL ? NULL : m_pSel->m_hIcon, FALSE);

		m_pParentPopup->RedrawWindow(m_pParentPopup->GetCaptionRect(), NULL, RDW_NOFRAME | RDW_INVALIDATE | RDW_UPDATENOW);
	}
}
//*********************************************************************************************************
void CBCGPWindowsNavigator::OnDraw(CDC* pDC)
{
	CBCGPPopupDlg::OnDraw(pDC);

	CBCGPFontSelector fs(*pDC, &globalData.fontRegular);
	pDC->SetBkMode(TRANSPARENT);

	COLORREF clrText = CBCGPVisualManager::GetInstance ()->GetPopupWindowTextColor(m_pParentPopup);
	
	if (m_pParentPopup != NULL)
	{
		ASSERT_VALID(m_pParentPopup);
		
		COLORREF clrTxtParent = m_pParentPopup->GetTextColor(this, FALSE, FALSE);
		if (clrTxtParent != (COLORREF)-1)
		{
			clrText = clrTxtParent;
		}
	}
	
	pDC->SetTextColor(clrText);

	int i = 0;

	for (i = 0; i < (int)m_arDockingPanes.GetSize(); i++)
	{
		CBCGPWindowsNavigatorObject* pPane = m_arDockingPanes[i];
		ASSERT_VALID(pPane);

		pPane->OnDraw(pDC, pPane == m_pSel);
	}

	for (i = 0; i < (int)m_arMDIChildren.GetSize(); i++)
	{
		CBCGPWindowsNavigatorObject* pDoc = m_arMDIChildren[i];
		ASSERT_VALID(pDoc);

		pDoc->OnDraw(pDC, pDoc == m_pSel);
	}
}
//*********************************************************************************************************
BOOL CBCGPWindowsNavigator::ReposItems(LPINT lpMDIColumns/* = NULL*/)
{
	int i = 0;
	const int nItemHeight = max(globalData.GetTextHeight(), globalData.m_sizeSmallIcon.cy) + globalUtils.ScaleByDPI(6);

	// Repos docking panes:	
	CRect rectPanes;
	m_wndPaneListLocation.GetWindowRect(rectPanes);
	ScreenToClient(rectPanes);

	int nColumnWidth = rectPanes.Width();

	CRect rectItem = rectPanes;
	rectItem.bottom = rectItem.top + nItemHeight;

	for (i = 0; i < (int)m_arDockingPanes.GetSize(); i++)
	{
		CBCGPWindowsNavigatorObject* pPane = m_arDockingPanes[i];
		ASSERT_VALID(pPane);

		pPane->m_nScroll = 0;
		pPane->m_rect.SetRectEmpty();

		if (i < m_nScrollOffset1 - 1)
		{
			continue;
		}

		if (i == m_nScrollOffset1 - 1)
		{
			pPane->m_nScroll = -1;
		}

		if (rectItem.bottom <= rectPanes.bottom)
		{
			pPane->m_rect = rectItem;
			rectItem.OffsetRect(0, nItemHeight);
		}
		else if (i > 1)
		{
			m_arDockingPanes[i - 1]->m_nScroll = 1;
		}
	}

	// Repos MDI documents:
	m_wndViewListLocation.GetWindowRect(rectPanes);
	ScreenToClient(rectPanes);

	rectItem = rectPanes;
	rectItem.right = rectItem.left + nColumnWidth;
	rectItem.bottom = rectItem.top + nItemHeight;
	
	BOOL bIsMDIScrolling = FALSE;

	int nColumn = 0;

	for (i = 0; i < (int)m_arMDIChildren.GetSize(); i++)
	{
		CBCGPWindowsNavigatorObject* pPane = m_arMDIChildren[i];
		ASSERT_VALID(pPane);
		
		pPane->m_nScroll = 0;
		pPane->m_rect.SetRectEmpty();
		
		if (i < m_nScrollOffset2 - 1 || bIsMDIScrolling)
		{
			continue;
		}
		
		if (i == m_nScrollOffset2 - 1)
		{
			pPane->m_nScroll = -1;
		}
		
		if (i != (int)m_arMDIChildren.GetSize() - 1 && rectItem.bottom > rectPanes.bottom)
		{
			if (nColumn < m_nMDIColumns - 1)
			{
				rectItem.OffsetRect(nColumnWidth, 0);
				rectItem.top = rectPanes.top;
				rectItem.bottom = rectItem.top + nItemHeight;

				nColumn++;
			}
			else if (i > 1)
			{
				m_arMDIChildren[i - 1]->m_nScroll = 1;
				bIsMDIScrolling = TRUE;
			}
		}
		
		if (rectItem.bottom <= rectPanes.bottom)
		{
			pPane->m_rect = rectItem;
			rectItem.OffsetRect(0, nItemHeight);
		}
	}

	if (lpMDIColumns != NULL)
	{
		*lpMDIColumns = nColumn + 1;
	}

	return bIsMDIScrolling;
}
//*********************************************************************************************************
CBCGPWindowsNavigatorObject* CBCGPWindowsNavigator::HitTest(CPoint pt, int* pIndex)
{
	int i = 0;
	
	for (i = 0; i < (int)m_arDockingPanes.GetSize(); i++)
	{
		CBCGPWindowsNavigatorObject* pPane = m_arDockingPanes[i];
		ASSERT_VALID(pPane);

		if (pPane->m_rect.PtInRect(pt))
		{
			if (pIndex != NULL)
			{
				*pIndex = i;
			}

			return pPane;
		}
	}
	
	for (i = 0; i < (int)m_arMDIChildren.GetSize(); i++)
	{
		CBCGPWindowsNavigatorObject* pDoc = m_arMDIChildren[i];
		ASSERT_VALID(pDoc);
		
		if (pDoc->m_rect.PtInRect(pt))
		{
			if (pIndex != NULL)
			{
				*pIndex = i;
			}

			return pDoc;
		}
	}

	if (pIndex != NULL)
	{
		*pIndex = -1;
	}

	return NULL;
}
//*********************************************************************************************************
BOOL CBCGPWindowsNavigator::PreTranslateMessage(MSG* pMsg) 
{
	if (pMsg->message == WM_LBUTTONDOWN)
	{
		CPoint point;
		::GetCursorPos (&point);

		CWnd* pWndParent = GetParent();
		if (pWndParent->GetSafeHwnd() != NULL)
		{
			CRect rectWindow;
			pWndParent->GetWindowRect(&rectWindow);

			if (!rectWindow.PtInRect(point))
			{
				pWndParent->SendMessage(WM_CLOSE);
				return TRUE;
			}
		}

		ScreenToClient(&point);

		CBCGPWindowsNavigatorObject* pPane = HitTest(point);
		if (pPane != NULL)
		{
			ASSERT_VALID(pPane);
			
			if (pPane->m_nScroll != 0)
			{
				if (pPane->m_bIsPane)
				{
					m_nScrollOffset1 += pPane->m_nScroll;
				}
				else
				{
					m_nScrollOffset2 += pPane->m_nScroll;
				}

				ReposItems();
			}
			else
			{
				SetSel(pPane);
			}

			RedrawWindow();
		}

		return TRUE;
	}
	
	if (pMsg->message == WM_LBUTTONUP)
	{
		CPoint point;
		::GetCursorPos (&point);
		ScreenToClient(&point);
		
		if (HitTest(point) == m_pSel)
		{
			SetActive(m_pSel);
		}
		
		return TRUE;
	}

	if (pMsg->message == WM_MOUSEMOVE)
	{
		CPoint point;
		::GetCursorPos (&point);
		ScreenToClient(&point);

		SetCursor(HitTest(point) != NULL ? globalData.GetHandCursor() : ::LoadCursor (NULL, IDC_ARROW));
		return TRUE;
	}

	if (pMsg->message >= WM_MOUSEFIRST && pMsg->message <= WM_MOUSELAST)
	{
		return TRUE;
	}

	if (pMsg->message >= WM_KEYFIRST && pMsg->message <= WM_KEYLAST)
	{
		return TRUE;
	}

	if (pMsg->message == WM_MOUSEWHEEL)
	{
		return TRUE;
	}

	return CBCGPPopupDlg::PreTranslateMessage(pMsg);
}
//*********************************************************************************************************
void CBCGPWindowsNavigator::OnTimer(UINT_PTR nIDEvent) 
{
	if (nIDEvent == ID_CHECK_KBD_STATE)
	{
		if (GetVirt(TRUE) != m_nInitialVirtKeys)
		{
			SetActive(m_pSel);
		}

		return;
	}
	
	CBCGPPopupDlg::OnTimer(nIDEvent);
}
//*********************************************************************************************************
void CBCGPWindowsNavigator::SetActive(CBCGPWindowsNavigatorObject* pPane)
{
	if (pPane != NULL)
	{
		ASSERT_VALID(pPane);

		HWND hwnd = pPane->m_hWnd;

		if (pPane->m_bIsPane)
		{
			CBCGPDockingControlBar* pDockingPane = DYNAMIC_DOWNCAST(CBCGPDockingControlBar, CWnd::FromHandlePermanent(hwnd));
			if (pDockingPane != NULL)
			{
				ASSERT_VALID(pDockingPane);

				pDockingPane->ShowControlBar(TRUE, FALSE, TRUE);
				pDockingPane->SetFocus();
			}
		}
		else
		{
			CWnd* pMDIChild = CWnd::FromHandlePermanent(pPane->m_hWnd);
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
					
					::GetWindowPlacement(hwnd, &wndpl);
						
					if (wndpl.showCmd == SW_SHOWMINIMIZED)
					{
						::ShowWindow(hwnd, SW_RESTORE);
					}
						
					::SendMessage(pMDIFrame->m_hWndMDIClient, WM_MDIACTIVATE, (WPARAM)hwnd, 0);

					if (m_bIsActivePane && pMDIFrame->MDIGetActive()->GetSafeHwnd() == hwnd)
					{
						::SetFocus(hwnd);
					}
				}
			}
		}
	}

	CWnd* pWndParent = GetParent();
	if (pWndParent->GetSafeHwnd() != NULL)
	{
		pWndParent->SendMessage(WM_CLOSE);
	}
}
//*********************************************************************************************************
void CBCGPWindowsNavigator::OnKeyDown(UINT nChar, UINT /*nRepCnt*/, UINT /*nFlags*/) 
{
	BOOL bPrev = FALSE;
	
	if (nChar == m_nNextChar)
	{
		nChar = UINT_MAX;
		
		if (m_nNextChar == m_nPrevChar)
		{
			bPrev = (GetVirt(FALSE) == m_nPrevVirtKeys);
		}
	}
	else if (nChar == m_nPrevChar)
	{
		nChar = UINT_MAX - 1;
		bPrev = TRUE;
	}
	
	if (m_pSel != NULL)
	{
		ASSERT_VALID(m_pSel);
		
		CRect rectPanes;
		m_wndPaneListLocation.GetWindowRect(rectPanes);
		ScreenToClient(rectPanes);

		CRect rectDocs;
		m_wndViewListLocation.GetWindowRect(rectDocs);
		ScreenToClient(rectDocs);

		int nColumnWidth = rectPanes.Width();

		int i = 0;
		CBCGPWindowsNavigatorObject* pNewSel = NULL;
		
		const CArray<CBCGPWindowsNavigatorObject*, CBCGPWindowsNavigatorObject*>& arPanes = m_pSel->m_bIsPane ? m_arDockingPanes : m_arMDIChildren;
		const CArray<CBCGPWindowsNavigatorObject*, CBCGPWindowsNavigatorObject*>& arPanesOther = !m_pSel->m_bIsPane ? m_arDockingPanes : m_arMDIChildren;

		const int nSize = (int)arPanes.GetSize();

		switch (nChar)
		{
		case VK_SHIFT:
		case VK_CONTROL:
		case VK_MENU:
			return;

		case VK_RETURN:
			SetActive(m_pSel);
			break;

		case UINT_MAX:
		case UINT_MAX - 1:
		case VK_F6:
			if (nSize == 1)
			{
				break;
			}

		case VK_DOWN:
		case VK_UP:
			if (nSize > 0)
			{
				if (nChar == VK_DOWN)
				{
					bPrev = FALSE;
				}
				else if (nChar == VK_UP)
				{
					bPrev = TRUE;
				}

				for (i = 0; i < nSize; i++)
				{
					if (arPanes[i] == m_pSel)
					{
						if (bPrev)
						{
							if (nChar == VK_UP && arPanesOther.GetSize() > 0)
							{
								if (i > 0)
								{
									pNewSel = arPanes[i - 1];
								}
								else
								{
									pNewSel = arPanesOther[arPanesOther.GetSize() - 1];
								}
							}
							else
							{
								pNewSel = arPanes[(i == 0) ? (nSize - 1) : (i - 1)];
							}
						}
						else
						{
							if (nChar == VK_DOWN && arPanesOther.GetSize() > 0)
							{
								if (i < nSize - 1)
								{
									pNewSel = arPanes[i + 1];
								}
								else
								{
									pNewSel = arPanesOther[0];
								}
							}
							else
							{
								pNewSel = arPanes[(i == nSize - 1) ? 0 : (i + 1)];
							}
						}
					}
				}
			}
			break;

		case VK_LEFT:
		case VK_RIGHT:
			{
				CPoint ptHit = m_pSel->m_rect.CenterPoint();

				if (GetExStyle () & WS_EX_LAYOUTRTL)
				{
					nChar = (nChar == VK_LEFT) ? VK_RIGHT : VK_LEFT;
				}

				if (nChar == VK_RIGHT)
				{
					ptHit.x += nColumnWidth;

					if (ptHit.x > rectDocs.right)
					{
						ptHit.x = rectPanes.left + nColumnWidth / 2;
					}
				}
				else
				{
					ptHit.x -= nColumnWidth;

					if (ptHit.x < rectPanes.left)
					{
						ptHit.x = rectDocs.right - nColumnWidth / 2;
					}
				}

				pNewSel = HitTest(ptHit);
				if (pNewSel == NULL)
				{
					do 
					{
						ptHit.y -= m_pSel->m_rect.Height();
						pNewSel = HitTest(ptHit);
					} 
					while (pNewSel == NULL && ptHit.y >= rectDocs.top);
				}
			}
			break;

		case VK_HOME:
			if (nSize > 0)
			{
				pNewSel = arPanes[0];
			}
			break;

		case VK_END:
			if (nSize > 0)
			{
				pNewSel = arPanes[nSize - 1];
			}
			break;
			
		default:
			SetActive(NULL);
			return;
		}
		
		if (pNewSel != NULL)
		{
			CRect rectOldSel = m_pSel->m_rect;
			BOOL bFullRedraw = FALSE;

			SetSel(pNewSel);

			if (m_pSel->m_nScroll != 0)
			{
				if (m_pSel->m_bIsPane)
				{
					m_nScrollOffset1 += 2 * m_pSel->m_nScroll;
				}
				else
				{
					m_nScrollOffset2 += 2 * m_pSel->m_nScroll;
				}

				ReposItems();
				bFullRedraw = TRUE;
			}

			if (m_pSel->m_rect.IsRectEmpty())
			{
				if (m_pSel->m_bIsPane)
				{
					for (m_nScrollOffset1 = 0; m_nScrollOffset1 < m_arDockingPanes.GetSize(); m_nScrollOffset1++)
					{
						ReposItems();

						if (!m_pSel->m_rect.IsRectEmpty() && m_pSel->m_nScroll == 0)
						{
							break;
						}
					}
				}
				else
				{
					for (m_nScrollOffset2 = 0; m_nScrollOffset2 < m_arMDIChildren.GetSize(); m_nScrollOffset2++)
					{
						ReposItems();
						
						if (!m_pSel->m_rect.IsRectEmpty() && m_pSel->m_nScroll == 0)
						{
							break;
						}
					}
				}

				bFullRedraw = TRUE;
			}

			if (bFullRedraw)
			{
				RedrawWindow();
			}
			else
			{
				InvalidateRect(rectOldSel);
				InvalidateRect(m_pSel->m_rect);

				UpdateWindow();
			}
		}
	}
	else
	{
		switch (nChar)
		{
		case VK_SHIFT:
		case VK_CONTROL:
		case VK_MENU:
			return;
			
		case UINT_MAX:
		case UINT_MAX - 1:
		case VK_F6:
		case VK_DOWN:
		case VK_UP:
		case VK_LEFT:
		case VK_RIGHT:
		case VK_HOME:
		case VK_END:
			if (m_arDockingPanes.GetSize() > 0)
			{
				SetSel(m_arDockingPanes[0]);
				RedrawWindow();
			}
			return;

		default:
			SetActive(NULL);
			return;
		}
	}
}

#endif
