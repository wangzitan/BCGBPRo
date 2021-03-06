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
// BCGPDropDownList.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPDropDownList.h"
#include "BCGPToolbarMenuButton.h"
#include "BCGPDialog.h"
#include "BCGPPropertyPage.h"
#include "BCGPWorkspace.h"
#include "BCGPRibbonComboBox.h"
#include "BCGPRibbonBar.h"
#include "BCGPKeyboardManager.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern CBCGPWorkspace* g_pWorkspace;

const UINT idStart = (UINT) -200;

/////////////////////////////////////////////////////////////////////////////
// CBCGPDropDownList

IMPLEMENT_DYNAMIC(CBCGPDropDownList, CBCGPPopupMenu)

CBCGPDropDownList::CBCGPDropDownList()
{
	CommonInit ();
}
//***************************************************************************
CBCGPDropDownList::CBCGPDropDownList(CWnd* pEditCtrl)
{
	ASSERT_VALID (pEditCtrl);

	CommonInit ();
	m_pEditCtrl = pEditCtrl;
}

#ifndef BCGP_EXCLUDE_RIBBON

CBCGPDropDownList::CBCGPDropDownList(CBCGPRibbonComboBox* pRibbonCombo)
{
	ASSERT_VALID (pRibbonCombo);

	CommonInit ();

	m_pRibbonCombo = pRibbonCombo;
	m_pEditCtrl = pRibbonCombo->m_pWndEdit;
}

#endif

void CBCGPDropDownList::CommonInit ()
{
	m_pEditCtrl = NULL;
	m_pRibbonCombo = NULL;

	m_bShowScrollBar = TRUE;
	m_nMaxHeight = -1;

	m_Menu.CreatePopupMenu ();
	m_nCurSel = -1;

	m_nMinWidth = -1;
	m_bDisableAnimation = TRUE;

	m_nLastTypingTime = 0;

	SetAutoDestroy (FALSE);
}
//***************************************************************************
CBCGPDropDownList::~CBCGPDropDownList()
{
}

BEGIN_MESSAGE_MAP(CBCGPDropDownList, CBCGPPopupMenu)
	//{{AFX_MSG_MAP(CBCGPDropDownList)
	ON_WM_KEYDOWN()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

CFont* CBCGPDropDownList::GetMenuFont()
{
	CFont* pFont = CBCGPPopupMenu::GetMenuFont();

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_pRibbonCombo != NULL)
	{
		ASSERT_VALID(m_pRibbonCombo);
		
		CBCGPRibbonBar* pTopLevelRibbon = m_pRibbonCombo->GetTopLevelRibbonBar();
		if (pTopLevelRibbon != NULL)
		{
			ASSERT_VALID(pTopLevelRibbon);

			CFont* pRibbonFont = pTopLevelRibbon->GetFont();
			if (pRibbonFont->GetSafeHandle() != globalData.fontRegular.GetSafeHandle())
			{
				pFont = pTopLevelRibbon->GetFont();
			}
		}
	}
#endif
	return pFont;
}

void CBCGPDropDownList::Track (CPoint point, CWnd *pWndOwner)
{
	CBCGPPopupMenuBar* pMenuBar = GetMenuBar ();
	ASSERT_VALID (pMenuBar);
	
	pMenuBar->m_bDropDownListMode = TRUE;
	pMenuBar->m_iMinWidth = m_nMinWidth;
	
	if (!Create (pWndOwner, point.x, point.y, m_Menu, FALSE, TRUE))
	{
		return;
	}

	pMenuBar->AdjustLocations();
	pMenuBar->m_bDisableSideBarInXPMode = TRUE;

	HighlightItem (m_nCurSel);
	pMenuBar->RedrawWindow ();

	CRect rect;
	GetWindowRect (&rect);
	UpdateShadow (&rect);

	CBCGPDialog* pParentDlg = NULL;
	if (pWndOwner != NULL && pWndOwner->GetParent () != NULL)
	{
		pParentDlg = DYNAMIC_DOWNCAST (CBCGPDialog, pWndOwner->GetParent ());
		if (pParentDlg != NULL)
		{
			pParentDlg->SetActiveMenu (this);
		}
	}

	CBCGPPropertyPage* pParentPropPage = NULL;
	if (pWndOwner != NULL && pWndOwner->GetParent () != NULL)
	{
		pParentPropPage = DYNAMIC_DOWNCAST (CBCGPPropertyPage, pWndOwner->GetParent ());
		if (pParentPropPage != NULL)
		{
			pParentPropPage->SetActiveMenu (this);
		}
	}
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPDropDownList message handlers

int CBCGPDropDownList::GetCount() const
{
	ASSERT_VALID (this);
	return m_Menu.GetMenuItemCount ();
}
//***************************************************************************
int CBCGPDropDownList::GetCurSel()
{
	ASSERT_VALID (this);

	if (GetSafeHwnd () == NULL)
	{
		return m_nCurSel;
	}

	CBCGPPopupMenuBar* pMenuBar = ((CBCGPDropDownList*) this)->GetMenuBar ();
	ASSERT_VALID (pMenuBar);

	CBCGPToolbarButton* pSel = pMenuBar->GetHighlightedButton ();
	if (pSel == NULL)
	{
		return -1;
	}

	int nIndex = 0;

	for (int i = 0; i < pMenuBar->GetCount (); i++)
	{
		CBCGPToolbarButton* pItem = pMenuBar->GetButton (i);
		ASSERT_VALID (pItem);

		if (!(pItem->m_nStyle & TBBS_SEPARATOR))
		{
			if (pSel == pItem)
			{
				m_nCurSel = nIndex;
				return nIndex;
			}

			nIndex++;
		}
	}

	return -1;
}
//***************************************************************************
int CBCGPDropDownList::SetCurSel(int nSelect)
{
	ASSERT_VALID (this);

	const int nSelOld = GetCurSel ();

	if (GetSafeHwnd () == NULL)
	{
		m_nCurSel = nSelect;
		return nSelOld;
	}

	CBCGPPopupMenuBar* pMenuBar = ((CBCGPDropDownList*) this)->GetMenuBar ();
	ASSERT_VALID (pMenuBar);

	if (nSelect < 0)
	{
		pMenuBar->m_iHighlighted = -1;
		pMenuBar->RedrawWindow();

		return -1;
	}

	int nIndex = 0;

	for (int i = 0; i < pMenuBar->GetCount (); i++)
	{
		CBCGPToolbarButton* pItem = pMenuBar->GetButton (i);
		ASSERT_VALID (pItem);

		if (!(pItem->m_nStyle & TBBS_SEPARATOR))
		{
			if (nIndex == nSelect)
			{
				HighlightItem (i);
				return nSelOld;
			}

			nIndex++;
		}
	}

	return -1;
}
//***************************************************************************
void CBCGPDropDownList::GetText(int nIndex, CString& rString) const
{
	ASSERT_VALID (this);

	CBCGPToolbarButton* pItem = GetItem (nIndex);
	if (pItem == NULL)
	{
		return;
	}

	ASSERT_VALID (pItem);
	rString = pItem->m_strText;
}
//***************************************************************************
void CBCGPDropDownList::AddString (LPCTSTR lpszItem)
{
	ASSERT_VALID (this);
	ASSERT(lpszItem != NULL);
	ASSERT(GetSafeHwnd () == NULL);

	const UINT uiID = idStart - GetCount ();
	m_Menu.AppendMenu (MF_STRING, uiID, lpszItem);
}
//***************************************************************************
void CBCGPDropDownList::ResetContent()
{
	ASSERT_VALID (this);
	ASSERT(GetSafeHwnd () == NULL);

	m_Menu.DestroyMenu ();
	m_Menu.CreatePopupMenu ();
}
//***************************************************************************
void CBCGPDropDownList::HighlightItem (int nIndex)
{
	ASSERT_VALID (this);

	CBCGPPopupMenuBar* pMenuBar = GetMenuBar ();
	ASSERT_VALID (pMenuBar);

	if (nIndex < 0)
	{
		return;
	}

	pMenuBar->m_iHighlighted = nIndex;

	SCROLLINFO scrollInfo;
	ZeroMemory(&scrollInfo, sizeof(SCROLLINFO));

    scrollInfo.cbSize = sizeof(SCROLLINFO);
	scrollInfo.fMask = SIF_ALL;

	m_wndScrollBarVert.GetScrollInfo (&scrollInfo);

	int iOffset = nIndex;
	int nMaxOffset = scrollInfo.nMax;

	iOffset = min (max (0, iOffset), nMaxOffset);

	if (iOffset != pMenuBar->GetOffset ())
	{
		pMenuBar->SetOffset (iOffset);

		m_wndScrollBarVert.SetScrollPos (iOffset);
		AdjustScroll ();
	}
	else
	{
		pMenuBar->RedrawWindow();
	}
}
//****************************************************************************
int CBCGPDropDownList::FindString(int nStart, LPCTSTR lpszText)
{
	ASSERT_VALID (this);
	ASSERT(lpszText != NULL);

	CString strIn = lpszText;
	strIn.MakeUpper();
	
	CBCGPPopupMenuBar* pMenuBar = ((CBCGPDropDownList*) this)->GetMenuBar ();
	ASSERT_VALID (pMenuBar);
	
	int nCurrIndex = 0;
	
	for (int i = 0; i < pMenuBar->GetCount (); i++)
	{
		CBCGPToolbarButton* pItem = pMenuBar->GetButton (i);
		ASSERT_VALID (pItem);
		
		if (!(pItem->m_nStyle & TBBS_SEPARATOR))
		{
			if (nCurrIndex > nStart)
			{
				CString strCurr = pItem->m_strText;
				strCurr.MakeUpper();

				if (strCurr.Find(strIn) == 0)
				{
					return nCurrIndex;
				}
			}
			
			nCurrIndex++;
		}
	}
	
	return -1;
}
//****************************************************************************
CBCGPToolbarButton* CBCGPDropDownList::GetItem (int nIndex) const
{
	ASSERT_VALID (this);

	CBCGPPopupMenuBar* pMenuBar = ((CBCGPDropDownList*) this)->GetMenuBar ();
	ASSERT_VALID (pMenuBar);

	int nCurrIndex = 0;

	for (int i = 0; i < pMenuBar->GetCount (); i++)
	{
		CBCGPToolbarButton* pItem = pMenuBar->GetButton (i);
		ASSERT_VALID (pItem);

		if (!(pItem->m_nStyle & TBBS_SEPARATOR))
		{
			if (nIndex == nCurrIndex)
			{
				return pItem;
			}

			nCurrIndex++;
		}
	}

	return NULL;
}
//****************************************************************************
void CBCGPDropDownList::OnDrawItem (CDC* pDC, CBCGPToolbarMenuButton* pItem, 
									BOOL bHighlight)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);
	ASSERT_VALID (pItem);

	CRect rectText = pItem->Rect ();
	rectText.DeflateRect (2 * globalUtils.ScaleByDPI(TEXT_MARGIN), 0);

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_pRibbonCombo != NULL)
	{
		ASSERT_VALID(m_pRibbonCombo);

		int nIndex = (int) idStart - pItem->m_nID;

		if (m_pRibbonCombo->OnDrawDropListItem (pDC, nIndex, pItem, bHighlight))
		{
			return;
		}
	}
#else
	UNREFERENCED_PARAMETER(bHighlight);
#endif

	pDC->DrawText (pItem->m_strText, &rectText, DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX);
}
//****************************************************************************
CSize CBCGPDropDownList::OnGetItemSize (CDC* pDC, 
										CBCGPToolbarMenuButton* pItem, CSize sizeDefault)
{
	ASSERT_VALID (this);

	sizeDefault.cx += GetSystemMetrics(SM_CXVSCROLL);

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_pRibbonCombo != NULL)
	{
		ASSERT_VALID (m_pRibbonCombo);

		int nIndex = (int) idStart - pItem->m_nID;

		CSize size = m_pRibbonCombo->OnGetDropListItemSize (
			pDC, nIndex, pItem, sizeDefault);

		if (size != CSize (0, 0))
		{
			return size;
		}
	}
#else
	UNREFERENCED_PARAMETER(pDC);
	UNREFERENCED_PARAMETER(pItem);
#endif
	return sizeDefault;
}
//****************************************************************************
void CBCGPDropDownList::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	ASSERT_VALID (this);

	switch (nChar)
	{
	case VK_UP:
	case VK_DOWN:
	case VK_PRIOR:
	case VK_NEXT:
	case VK_ESCAPE:
	case VK_RETURN:
	case VK_BACK:
	case VK_DELETE:
		CBCGPPopupMenu::OnKeyDown(nChar, nRepCnt, nFlags);
		return;
	}
		
	if (m_pEditCtrl->GetSafeHwnd () != NULL)
	{
		m_pEditCtrl->SendMessage (WM_KEYDOWN, nChar, MAKELPARAM (nRepCnt, nFlags));
		return;
	}
	else
	{
		if (CBCGPKeyboardManager::IsKeyPrintable(nChar))
		{
			CString strText;

			BYTE lpKeyState [256];
			::GetKeyboardState (lpKeyState);
			
#ifndef _UNICODE
			WORD wChar = 0;
			if (::ToAsciiEx (nChar,
				MapVirtualKey (nChar, 0),
				lpKeyState,
				&wChar,
				0,
				::GetKeyboardLayout (AfxGetThread()->m_nThreadID)) > 0)
			{
				strText = (TCHAR)wChar;
			}
			
#else
			TCHAR szChar [2];
			memset (szChar, 0, sizeof (TCHAR) * 2);
			
			if (::ToUnicodeEx (nChar,
				MapVirtualKey (nChar, 0),
				lpKeyState,
				szChar,
				2,
				0,
				::GetKeyboardLayout (AfxGetThread()->m_nThreadID)) > 0)
			{
				strText = szChar;
			}

#endif // _UNICODE

			if (strText.IsEmpty())
			{
				CBCGPPopupMenu::OnKeyDown(nChar, nRepCnt, nFlags);
				return;
			}

			int nCurrSel = GetCurSel();
			BOOL bNewSearch = FALSE;
			
			clock_t nCurrTypingTime = clock();
			
			if (nCurrTypingTime - m_nLastTypingTime > 1000)
			{
				m_stTyping.Empty();
				bNewSearch = TRUE;
			}
			
			m_stTyping += strText;
			m_nLastTypingTime = nCurrTypingTime;
			
			int nNewSel = bNewSearch ? FindString(nCurrSel, m_stTyping) : -1;
			if (nNewSel < 0)
			{
				nNewSel = FindString(-1, m_stTyping);

				if (nNewSel < 0 && m_stTyping.GetLength() == 2 && m_stTyping[0] == m_stTyping[1])
				{
					m_stTyping = m_stTyping.Left(1);
					
					nNewSel = FindString(nCurrSel, m_stTyping);
					if (nNewSel < 0)
					{
						nNewSel = FindString(-1, m_stTyping);
					}
				}
			}
			
			if (nCurrSel != nNewSel && nNewSel >= 0)
			{
				HighlightItem(nNewSel);

				CBCGPPopupMenuBar* pMenuBar = GetMenuBar ();
				ASSERT_VALID (pMenuBar);

				pMenuBar->RedrawWindow ();
				return;
			}
		}
	}
	
	CBCGPPopupMenu::OnKeyDown(nChar, nRepCnt, nFlags);
}
//****************************************************************************
void CBCGPDropDownList::OnChooseItem (UINT uiCommandID)
{
	ASSERT_VALID (this);

	CBCGPPopupMenu::OnChooseItem (uiCommandID);

#ifndef BCGP_EXCLUDE_RIBBON
	int nIndex = (int) idStart - uiCommandID;

	if (m_pRibbonCombo != NULL)
	{
		ASSERT_VALID (m_pRibbonCombo);
		m_pRibbonCombo->OnSelectItem (nIndex);
	}
#else
	UNREFERENCED_PARAMETER(uiCommandID);
#endif
}
//****************************************************************************
void CBCGPDropDownList::OnChangeHot (int nHot)
{
	ASSERT_VALID (this);

	CBCGPPopupMenu::OnChangeHot (nHot);

#ifndef BCGP_EXCLUDE_RIBBON
	if (m_pRibbonCombo != NULL)
	{
		ASSERT_VALID (m_pRibbonCombo);
		m_pRibbonCombo->NotifyHighlightListItem (nHot);
	}
#else
	UNREFERENCED_PARAMETER(nHot);
#endif
}
//****************************************************************************
BOOL CBCGPDropDownList::IsParentEditFocused()
{
	return IsEditFocused();
}
