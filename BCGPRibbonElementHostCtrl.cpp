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
// BCGPRibbonElementHostCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "bcgcbpro.h"
#include "BCGPRibbonElementHostCtrl.h"
#include "BCGPRibbonCategory.h"
#include "BCGPRibbonBackstageView.h"
#include "BCGPVisualManager.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#ifndef BCGP_EXCLUDE_RIBBON

static const int nScrollBarID = 1;

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonGalleryCtrl

IMPLEMENT_DYNAMIC(CBCGPRibbonGalleryCtrl, CBCGPRibbonPanelMenuBar)

CBCGPRibbonGalleryCtrl::CBCGPRibbonGalleryCtrl()
{
	ASSERT_VALID(m_pPanel);

	m_PaletteButton.SetDrawDisabledItems(TRUE);

	m_bIsCtrlMode = TRUE;
	m_pPanel->m_bMenuMode = TRUE;

	m_pPanel->m_pPaletteButton = &m_PaletteButton;
	m_pPanel->m_pParentMenuBar = this;
	m_PaletteButton.m_pParentControl = this;

	m_bIsMouseClicked = FALSE;
	m_bNotifyBySingleClick = FALSE;
	m_pClicked = NULL;
}

CBCGPRibbonGalleryCtrl::~CBCGPRibbonGalleryCtrl()
{
	m_pPanel->m_arElements.RemoveAll ();
}

BEGIN_MESSAGE_MAP(CBCGPRibbonGalleryCtrl, CBCGPRibbonPanelMenuBar)
	//{{AFX_MSG_MAP(CBCGPRibbonGalleryCtrl)
	ON_WM_ERASEBKGND()
	ON_WM_CREATE()
	ON_WM_MOUSEWHEEL()
	ON_WM_DESTROY()
	ON_WM_LBUTTONUP()
	ON_WM_LBUTTONDBLCLK()
	ON_WM_LBUTTONDOWN()
	ON_WM_CANCELMODE()
	ON_WM_SETFOCUS()
	ON_WM_KILLFOCUS()
	ON_WM_ENABLE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonGalleryCtrl message handlers

void CBCGPRibbonGalleryCtrl::InitCtrl()
{
	ASSERT_VALID(m_pPanel);

	if (m_wndScrollBarVert.GetSafeHwnd() == NULL)
	{
		m_wndScrollBarVert.Create (WS_CHILD | WS_VISIBLE | SBS_VERT,
			CRect (0, 0, 0, 0), this, nScrollBarID);
		m_pPanel->m_pScrollBar = &m_wndScrollBarVert;
	}

	ModifyStyle(0, WS_CLIPCHILDREN);

	if (m_pPanel->IsEmpty())
	{
		RecalcLayout();
	}
}
//***********************************************************************************************	
void CBCGPRibbonGalleryCtrl::PreSubclassWindow() 
{
	CBCGPRibbonPanelMenuBar::PreSubclassWindow();
}
//***********************************************************************************************	
BOOL CBCGPRibbonGalleryCtrl::PreTranslateMessage(MSG* pMsg)
{
	if (pMsg->message == WM_KEYDOWN && IsWindowVisible())
	{
		switch (pMsg->wParam)
		{
		case VK_RETURN:
			NotifyCommand(FALSE);
			return TRUE;

		case VK_ESCAPE:
			return FALSE;
		}
	}

	return CBCGPRibbonPanelMenuBar::PreTranslateMessage(pMsg);
}
//***********************************************************************************************	
BOOL CBCGPRibbonGalleryCtrl::OnEraseBkgnd(CDC* /*pDC*/) 
{
	return TRUE;
}
//***********************************************************************************************	
int CBCGPRibbonGalleryCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CBCGPRibbonPanelMenuBar::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	InitCtrl();
	return 0;
}
//***********************************************************************************************	
BOOL CBCGPRibbonGalleryCtrl::Create(const RECT& rect, CWnd* pParentWnd, UINT nID, DWORD dwStyle) 
{
	return CWnd::Create (
		globalData.RegisterWindowClass (_T("BCGPRibbonGallery")), _T(""), dwStyle, rect, pParentWnd, nID, 0);
}
//***********************************************************************************************	
void CBCGPRibbonGalleryCtrl::RemoveAll(BOOL bRecalcLayout)
{
	ASSERT_VALID(m_pPanel);

	const int nScrollButtons = m_PaletteButton.m_bIsComboMode ? 0 : 3;

	for (int i = 0; i < nScrollButtons; i++)
	{
		delete m_PaletteButton.m_arIcons[m_PaletteButton.m_arIcons.GetSize() - i - 1];
	}

	m_pPanel->RemoveAll();
	m_PaletteButton.m_arIcons.RemoveAll ();

	if (bRecalcLayout)
	{
		AdjustLocations();
	}
}
//***********************************************************************************************	
void CBCGPRibbonGalleryCtrl::RecalcLayout()
{
	if (GetSafeHwnd() == NULL)
	{
		return;
	}

	ASSERT_VALID(m_pPanel);

	m_pPanel->m_arElements.RemoveAll ();
	m_PaletteButton.RemoveAll();

	m_PaletteButton.CreateIcons();

	CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arButtons;
	m_PaletteButton.GetMenuItems (arButtons);

	for (int i = 0; i < arButtons.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = arButtons [i];
		ASSERT_VALID (pButton);

		m_pPanel->Add (pButton);
	}

	AdjustLocations();
}
//***********************************************************************************************	
BOOL CBCGPRibbonGalleryCtrl::OnMouseWheel(UINT /*nFlags*/, short zDelta, CPoint /*pt*/) 
{
	ASSERT_VALID (this);

	const int nSteps = abs(zDelta) / WHEEL_DELTA;

	for (int i = 0; i < nSteps; i++)
	{
		if (m_iOffset > 0 || m_pPanel->m_bScrollDnAvailable)
		{
			int iOffset = GetOffset ();

			if (zDelta > 0)
			{
				if (m_iOffset > 0)
				{
					SetOffset (iOffset - 1);
				}
			}
			else
			{
				if (m_pPanel->m_bScrollDnAvailable)
				{
					SetOffset (iOffset + 1);
				}
			}
		}
		else
		{
			OnVScroll (zDelta < 0 ? SB_LINEDOWN : SB_LINEUP, 0, 
				&m_wndScrollBarVert);
		}
	}

	return TRUE;
}
//****************************************************************************************
void CBCGPRibbonGalleryCtrl::OnDestroy() 
{
	m_PaletteButton.Clear();
	m_pPanel->m_arElements.RemoveAll ();

	CBCGPRibbonPanelMenuBar::OnDestroy();
}
//****************************************************************************************
void CBCGPRibbonGalleryCtrl::OnLButtonUp(UINT nFlags, CPoint point) 
{
	int nOldSelected = GetSelectedItem();
	
	CBCGPRibbonPanelMenuBar::OnLButtonUp(nFlags, point);

	if (m_bNotifyBySingleClick && m_PaletteButton.GetHighlighted() == m_pClicked && m_pClicked != NULL)
	{
		m_pClicked = NULL;
		NotifyCommand(TRUE);
		return;
	}

	if (::IsWindow(GetSafeHwnd()) && nOldSelected != GetSelectedItem())
	{
		m_bIsMouseClicked = TRUE;

		CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST(CBCGPRibbonPaletteIcon, m_PaletteButton.GetHighlighted());
		if (pIcon != NULL)
		{
			GetParent()->SendMessage (BCGM_ON_HIGHLIGHT_RIBBON_LIST_ITEM, (WPARAM) pIcon->GetIndex(), (LPARAM) pIcon);
		}

		m_bIsMouseClicked = FALSE;
	}

	m_pClicked = NULL;
}
//****************************************************************************************
BOOL CBCGPRibbonGalleryCtrl::OnKey (UINT nChar)
{
	CBCGPBaseRibbonElement* pOldHighlighted = m_PaletteButton.GetHighlighted();
	
	BOOL bRes = CBCGPRibbonPanelMenuBar::OnKey(nChar);

	if (::IsWindow(GetSafeHwnd()))
	{
		if (pOldHighlighted != m_PaletteButton.GetHighlighted())
		{
			CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST(CBCGPRibbonPaletteIcon, m_PaletteButton.GetHighlighted());
			if (pIcon != NULL)
			{
				SelectItem(pIcon->GetIndex());
				RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
				GetParent()->SendMessage (BCGM_ON_HIGHLIGHT_RIBBON_LIST_ITEM, (WPARAM) pIcon->GetIndex(), (LPARAM) pIcon);
			}
		}
	}

	return bRes;
}
//****************************************************************************************
void CBCGPRibbonGalleryCtrl::OnLButtonDblClk(UINT /*nFlags*/, CPoint /*point*/) 
{
	if (!m_bNotifyBySingleClick)
	{
		NotifyCommand(TRUE);
	}
}
//**************************************************************************
void CBCGPRibbonGalleryCtrl::EnsureVisible(int nItemIndex)
{
	ASSERT_VALID(this);

	int nCurrIndex = 0;

	CRect rect;
	GetClientRect(rect);
	
	for (int i = 0; i < m_PaletteButton.m_arIcons.GetSize(); i++)
	{
		CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST(CBCGPRibbonPaletteIcon, m_PaletteButton.m_arIcons[i]);
		
		if (pIcon == NULL)
		{
			continue;
		}
		
		ASSERT_VALID (pIcon);
		
		if (nCurrIndex == nItemIndex)
		{
			if (pIcon->GetRect().bottom > rect.bottom)
			{
				while (TRUE)
				{
					int nScrollOffsetOld = m_wndScrollBarVert.GetScrollPos();

					OnVScroll(SB_LINEDOWN, 0, &m_wndScrollBarVert);

					if (nScrollOffsetOld == m_wndScrollBarVert.GetScrollPos() ||
						pIcon->GetRect().bottom <= rect.bottom)
					{
						break;
					}
				}
			}
			else if (pIcon->GetRect().top < rect.top)
			{
				while (TRUE)
				{
					int nScrollOffsetOld = m_wndScrollBarVert.GetScrollPos();
					
					OnVScroll(SB_LINEUP, 0, &m_wndScrollBarVert);
					
					if (nScrollOffsetOld == m_wndScrollBarVert.GetScrollPos() ||
						pIcon->GetRect().top >= rect.top)
					{
						break;
					}
				}
			}

			break;
		}
		
		nCurrIndex++;
	}
	
	RedrawWindow();
}
//**************************************************************************
void CBCGPRibbonGalleryCtrl::NotifyCommand(BOOL bByMouseClick)
{
	if (::IsWindow(GetSafeHwnd()) && GetOwner()->GetSafeHwnd() != NULL)
	{
		m_bIsMouseClicked = bByMouseClick;
		GetOwner()->SendMessage(WM_COMMAND, (WPARAM)GetDlgCtrlID(), (LPARAM)GetSafeHwnd());
		m_bIsMouseClicked = FALSE;
	}
}
//**************************************************************************
void CBCGPRibbonGalleryCtrl::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CBCGPRibbonPanelMenuBar::OnLButtonDown(nFlags, point);

	m_pClicked = m_PaletteButton.GetHighlighted();

	if (DYNAMIC_DOWNCAST(CBCGPRibbonPaletteIcon, m_pClicked) == NULL)
	{
		m_pClicked = NULL;
	}
}
//**************************************************************************
void CBCGPRibbonGalleryCtrl::OnCancelMode() 
{
	CBCGPRibbonPanelMenuBar::OnCancelMode();
	m_pClicked = NULL;
}
//**************************************************************************
void CBCGPRibbonGalleryCtrl::OnSetFocus(CWnd* pOldWnd) 
{
	CBCGPRibbonPanelMenuBar::OnSetFocus(pOldWnd);

	if (GetSelectedItem() < 0 && GetCount() > 0)
	{
		SelectItem(0);
	}

	m_PaletteButton.OnSetFocus(TRUE);
	m_PaletteButton.m_bIsFocused = TRUE;

	RedrawWindow();
}
//**************************************************************************
void CBCGPRibbonGalleryCtrl::OnKillFocus(CWnd* pNewWnd) 
{
	CBCGPRibbonPanelMenuBar::OnKillFocus(pNewWnd);
	m_PaletteButton.m_bIsFocused = FALSE;

	RedrawWindow();
}
//**************************************************************************
void CBCGPRibbonGalleryCtrl::OnEnable(BOOL bEnable) 
{
	CBCGPRibbonPanelMenuBar::OnEnable(bEnable);

	m_PaletteButton.m_bIsDisabled = !bEnable;

	if (m_wndScrollBarVert.GetSafeHwnd() != NULL)
	{
		m_wndScrollBarVert.EnableWindow(bEnable);
	}

	RedrawWindow();
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPNewDocumentGalleryCtrl control

IMPLEMENT_DYNAMIC(CBCGPNewDocumentGalleryCtrl, CBCGPRibbonGalleryCtrl)

CBCGPNewDocumentGalleryCtrl::CBCGPNewDocumentGalleryCtrl()
{
	m_nTemplatesGroupStartIndex = -1;
	m_sizeTemplatePadding = CSize(0, 0);
}
//**************************************************************************
int CBCGPNewDocumentGalleryCtrl::AddApplicationDocTemplates(const CString strGroupName, CSize sizeItem, CSize sizePadding)
{
	for (POSITION pos = AfxGetApp()->GetFirstDocTemplatePosition(); pos != NULL;)
	{
		CMultiDocTemplate* pTemplate = DYNAMIC_DOWNCAST(CMultiDocTemplate, AfxGetApp()->GetNextDocTemplate(pos));
		if (pTemplate != NULL)
		{
			ASSERT_VALID(pTemplate);
			m_arTemplates.Add((CBCGPMultiDocTemplate*)pTemplate);
		}
	}
	
	const int nSize = (int)m_arTemplates.GetSize();
	if (nSize == 0)
	{
		return 0;
	}

	m_nTemplatesGroupStartIndex = GetCount();

	AddGroup(strGroupName, globalUtils.ScaleByDPI(sizeItem), nSize);
	
	for (int i = 0; i < nSize; i++)
	{
		CBCGPMultiDocTemplate* pTemplate = m_arTemplates[i];
		ASSERT_VALID(pTemplate);

		int nIndex = m_nTemplatesGroupStartIndex + i;

		CString strTypeName;
		if (pTemplate->GetDocString(strTypeName, CDocTemplate::fileNewName) && !strTypeName.IsEmpty())
		{
			m_PaletteButton.SetItemToolTip(nIndex, strTypeName);
		}
	}

	m_sizeTemplatePadding = sizePadding;
	return nSize;
}
//**************************************************************************
BOOL CBCGPNewDocumentGalleryCtrl::CloseBackstageView()
{
	for (CWnd* pParent = GetParent(); pParent->GetSafeHwnd() != NULL; pParent = pParent->GetParent())
	{
		if (DYNAMIC_DOWNCAST(CBCGPRibbonBackstageView, pParent) != NULL)
		{
			pParent->SendMessage(WM_CLOSE);
			return TRUE;
		}
	}
	
	return FALSE;
}
//**************************************************************************
void CBCGPNewDocumentGalleryCtrl::NotifyCommand(BOOL bByMouseClick)
{
	if (m_arTemplates.GetSize() > 0)
	{
		int nSel = GetSelectedItem();
		int nTemplateIndex = nSel - m_nTemplatesGroupStartIndex;

		if (nTemplateIndex >= 0 && nTemplateIndex < (int)m_arTemplates.GetSize())
		{
			CBCGPMultiDocTemplate* pTemplate = m_arTemplates[nTemplateIndex];
			ASSERT_VALID(pTemplate);

			pTemplate->OpenDocumentFile(NULL);
			
			CloseBackstageView();
			return;
		}
	}

	CBCGPRibbonGalleryCtrl::NotifyCommand(bByMouseClick);
}
//**************************************************************************
void CBCGPNewDocumentGalleryCtrl::OnDrawGalleryItem (CDC* pDC, CRect rect, int nIconIndex, CBCGPRibbonPaletteIcon* pIcon, COLORREF clrText)
{
	int nTemplateIndex = nIconIndex - m_nTemplatesGroupStartIndex;
	
	int nImageHeight = 0;

	rect.DeflateRect(globalUtils.ScaleByDPI(m_sizeTemplatePadding.cx), globalUtils.ScaleByDPI(m_sizeTemplatePadding.cy));

	if (nTemplateIndex >= 0 && m_nTemplatesGroupStartIndex < (int)m_arTemplates.GetSize())
	{
		CBCGPMultiDocTemplate* pTemplate = m_arTemplates[nTemplateIndex];

		HICON hIcon = AfxGetApp()->LoadIcon (pTemplate->GetResId ());
		if (hIcon == NULL)
		{
			hIcon = ::LoadIcon(NULL, IDI_APPLICATION);
		}
	
		if (hIcon != NULL)
		{
			CSize size = globalUtils.GetIconSize(hIcon);

			::DrawIconEx(pDC->GetSafeHdc(), rect.left + max(0, (rect.Width() - size.cx) / 2), rect.top, hIcon, 
				size.cx, size.cy, NULL, (HBRUSH)NULL, DI_NORMAL);

			nImageHeight = size.cy + globalUtils.ScaleByDPI(m_sizeTemplatePadding.cy);
		}
	}

	CRect rectText = rect;
	rectText.top += nImageHeight;
	
	COLORREF clrTextOld = (COLORREF)-1;

	if (clrText != (COLORREF)-1)
	{
		clrTextOld = pDC->SetTextColor(clrText);
	}

	pDC->DrawText(m_PaletteButton.GetIconTextLabel(pIcon), rectText, DT_WORDBREAK | DT_END_ELLIPSIS | DT_CENTER);

	if (clrTextOld != (COLORREF)-1)
	{
		pDC->SetTextColor(clrTextOld);
	}
}

#endif // BCGP_EXCLUDE_RIBBON
