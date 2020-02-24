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
// BCGPTreeCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPTreeCtrl.h"
#include "BCGPDlgImpl.h"
#include "BCGPVisualManager.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPTreeCtrl

IMPLEMENT_DYNAMIC(CBCGPTreeCtrl, CTreeCtrl)

CBCGPTreeCtrl::CBCGPTreeCtrl()
{
	m_bVisualManagerStyle = FALSE;
	m_bImageListStateChanged = FALSE;
}

CBCGPTreeCtrl::~CBCGPTreeCtrl()
{
}

BEGIN_MESSAGE_MAP(CBCGPTreeCtrl, CTreeCtrl)
	//{{AFX_MSG_MAP(CBCGPTreeCtrl)
	ON_WM_ERASEBKGND()
	ON_WM_CREATE()
	ON_WM_NCPAINT()
	ON_WM_ENABLE()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_PRINT, OnPrint)
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, OnCustomDraw)
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLVMMODE, OnBCGSetControlVMMode)
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)
	ON_MESSAGE(TVM_SETIMAGELIST, OnSetImageList)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPTreeCtrl message handlers

void CBCGPTreeCtrl::OnCustomDraw(NMHDR* pNMHDR, LRESULT* pResult)
{
	if (m_bVisualManagerStyle && !m_bImageListStateChanged && (GetStyle() & TVS_CHECKBOXES) == TVS_CHECKBOXES)
	{
		CImageList* pImageList = GetImageList(TVSIL_STATE);
		if (m_ilState.GetSafeHandle() == NULL && pImageList != NULL)
		{
			int cx = 0;
			int cy = 0;
			if (ImageList_GetIconSize(pImageList->GetSafeHandle(), &cx, &cy))
			{
				if (cx != 0 && cy != 0)
				{
					m_ilState.Create(cx, cy, ILC_COLOR32, 2, 0);

					SetImageList(&m_ilState, TVSIL_STATE);

					UpdateImageListState();
				}
			}
		}
	}

	LPNMTVCUSTOMDRAW pNMTVCD = (LPNMTVCUSTOMDRAW) pNMHDR;
	if (pNMTVCD == NULL)
	{
		*pResult = -1;
		return;
	}

	*pResult = CDRF_DODEFAULT;

	if (!m_bVisualManagerStyle)
	{
		return;
	}

	if (pNMTVCD->nmcd.dwDrawStage == CDDS_PREPAINT)
	{
		*pResult = CDRF_NOTIFYPOSTPAINT | CDRF_NOTIFYITEMDRAW;
	}
	else if (pNMTVCD->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
	{
		BOOL bIsSelected = FALSE;
		BOOL bIsFocused = FALSE;
		BOOL bIsDisabled = FALSE;

		if (!IsWindowEnabled ())
		{
			bIsDisabled = TRUE;
		}
		else
		{
			UINT nState = GetItemState ((HTREEITEM) pNMTVCD->nmcd.dwItemSpec, TVIF_STATE);
			if ((nState & (TVIS_SELECTED | TVIS_DROPHILITED)) != 0)
			{
				if (GetFocus () != this && (nState & TVIS_DROPHILITED) == 0)
				{
					bIsSelected = (GetStyle () & TVS_SHOWSELALWAYS) == TVS_SHOWSELALWAYS;
				}
				else
				{
					bIsSelected = TRUE;
					bIsFocused = TRUE;
				}
			}
		}

		COLORREF clrBk = CBCGPVisualManager::GetInstance ()->GetTreeControlFillColor(this, bIsSelected, bIsFocused, bIsDisabled);
		COLORREF clrText = CBCGPVisualManager::GetInstance ()->GetTreeControlTextColor(this, bIsSelected, bIsFocused, bIsDisabled);

		if (clrBk != (COLORREF)-1)
		{
			pNMTVCD->clrTextBk = clrBk;
		}

		if (clrText != (COLORREF)-1)
		{
			pNMTVCD->clrText = clrText;
		}
	}
}
//**************************************************************************
LRESULT CBCGPTreeCtrl::OnBCGSetControlVMMode (WPARAM wp, LPARAM)
{
	m_bVisualManagerStyle = (BOOL) wp;
	OnChangeVisualManager (0, 0);

	return 0L;
}
//**************************************************************************
BOOL CBCGPTreeCtrl::OnEraseBkgnd(CDC* pDC) 
{
#ifndef _BCGSUITE_
	if (m_bVisualManagerStyle && (globalData.bIsWindows9x || globalData.bIsWindowsNT4))
	{
		COLORREF clr = CBCGPVisualManager::GetInstance ()->GetTreeControlFillColor(this);
		if (clr != (COLORREF)-1)
		{
			CRect rectClient;
			GetClientRect (&rectClient);

			CBrush br(clr);
			pDC->FillRect(rectClient, &br);

			return TRUE;
		}
	}
#endif
	return CTreeCtrl::OnEraseBkgnd(pDC);
}
//**************************************************************************
LRESULT CBCGPTreeCtrl::OnChangeVisualManager (WPARAM, LPARAM)
{
	HWND hWnd = GetSafeHwnd ();
	if (hWnd == NULL)
	{
		return 0L;
	}

	COLORREF clrBack = (COLORREF)-1;
	COLORREF clrText = (COLORREF)-1;
	COLORREF clrLine = CLR_DEFAULT;

	if (m_bVisualManagerStyle)
	{
		clrBack = CBCGPVisualManager::GetInstance ()->GetTreeControlFillColor(this);
		clrText = CBCGPVisualManager::GetInstance ()->GetTreeControlTextColor(this);
		clrLine = CBCGPVisualManager::GetInstance ()->GetTreeControlLineColor(this);

		UpdateImageListState();
	}

	TreeView_SetBkColor(hWnd, clrBack);
	TreeView_SetTextColor(hWnd, clrText);
	TreeView_SetLineColor(hWnd, clrLine);

	RedrawWindow(NULL, NULL, RDW_ERASE | RDW_INVALIDATE | RDW_UPDATENOW | RDW_FRAME);

	return 0L;
}
//**************************************************************************
int CBCGPTreeCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CTreeCtrl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	if (m_bVisualManagerStyle)
	{
		OnChangeVisualManager (0, 0);
	}
	
	return 0;
}
//**************************************************************************
void CBCGPTreeCtrl::PreSubclassWindow() 
{
	CTreeCtrl::PreSubclassWindow();

	if (m_bVisualManagerStyle)
	{
		OnChangeVisualManager (0, 0);
	}
}
//**************************************************************************
void CBCGPTreeCtrl::OnNcPaint()
{
	Default();

	if (!m_bVisualManagerStyle)
	{
		return;
	}

	if ((GetStyle () & WS_BORDER) != 0 || (GetExStyle () & WS_EX_CLIENTEDGE) != 0)
	{
		CBCGPVisualManager::GetInstance ()->OnDrawControlBorder(this);
	}
}
//****************************************************************************
LRESULT CBCGPTreeCtrl::OnPrint(WPARAM wp, LPARAM lp)
{
	LRESULT lRes = Default();

	if (!m_bVisualManagerStyle)
	{
		return lRes;
	}

	DWORD dwFlags = (DWORD)lp;

	if ((dwFlags & PRF_NONCLIENT) == PRF_NONCLIENT)
	{
		if ((GetStyle () & WS_BORDER) != 0 || (GetExStyle () & WS_EX_CLIENTEDGE) != 0)
		{
			CDC* pDC = CDC::FromHandle((HDC) wp);
			ASSERT_VALID(pDC);

			CRect rect;
			GetWindowRect(rect);
			
			rect.bottom -= rect.top;
			rect.right -= rect.left;
			rect.left = rect.top = 0;

			CBCGPVisualManager::GetInstance ()->OnDrawControlBorder(pDC, rect, this, FALSE);
			return 0;
		}
	}

	return lRes;
}
//************************************************************************
BOOL CBCGPTreeCtrl::IsInternalScrollBarThemed() const
{
#ifndef _BCGSUITE_
	return (globalData.m_nThemedScrollBars & BCGP_THEMED_SCROLLBAR_TREECTRL) != 0 && m_bVisualManagerStyle;
#else
	return FALSE;
#endif
}
//************************************************************************
void CBCGPTreeCtrl::OnEnable(BOOL bEnable)
{
	CTreeCtrl::OnEnable(bEnable);

	if (m_bVisualManagerStyle)
	{
		UpdateImageListState();
	}
}
//************************************************************************
LRESULT CBCGPTreeCtrl::OnSetImageList(WPARAM wp, LPARAM lp)
{
	LRESULT lRes = Default();

	if (wp == TVSIL_STATE)
	{
		if (m_ilState.GetSafeHandle() != (HIMAGELIST)lp)
		{
			m_bImageListStateChanged = TRUE;
		}
	}

	return lRes;
}
//************************************************************************
void CBCGPTreeCtrl::UpdateImageListState()
{
	if ((GetStyle() & TVS_CHECKBOXES) == 0)
	{
		return;
	}

	if (m_ilState.GetSafeHandle() != NULL && 
		m_ilState.GetSafeHandle() == GetImageList(TVSIL_STATE)->GetSafeHandle())
	{
		const BOOL bIsDisabled = !IsWindowEnabled();
		COLORREF clrBkgnd = CBCGPVisualManager::GetInstance ()->GetTreeControlFillColor(this, FALSE, FALSE, bIsDisabled);
		if (clrBkgnd == (COLORREF)-1)
		{
			clrBkgnd = CLR_DEFAULT;
		}

		CArray<int, int> states;
		states.SetSize(3);
		states[0] = 2;
		states[1] = 0;
		states[2] = 1;

		CBCGPVisualManager::GetInstance()->FillCheckImageList(m_ilState, clrBkgnd, states, bIsDisabled);
	}
}


/////////////////////////////////////////////////////////////////////////////
// CBCGPTreeCtrlEx

#ifndef BCGP_EXCLUDE_GRID_CTRL

IMPLEMENT_DYNAMIC(CBCGPTreeCtrlEx, CBCGPGridCtrl)

CBCGPTreeCtrlEx::CBCGPTreeCtrlEx()
{
	m_bHeader = FALSE;
	m_bSelectionBorder = FALSE;
	m_bEditFirstClick = FALSE;
	m_bHighlightActiveItem = TRUE;
	m_bRowMarker = FALSE;
	m_bFillLeftOffsetArea = FALSE;
	m_bGridLines = FALSE;
	m_bAlternateRows = FALSE;
	m_bWholeCellSel = FALSE;
	m_bNeedToRepos = FALSE;
	m_bTreeCtrlNotifications = TRUE;
}

CBCGPTreeCtrlEx::~CBCGPTreeCtrlEx()
{
}

BEGIN_MESSAGE_MAP(CBCGPTreeCtrlEx, CBCGPGridCtrl)
//{{AFX_MSG_MAP(CBCGPTreeCtrlEx)
	ON_WM_STYLECHANGED()
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

void CBCGPTreeCtrlEx::Init ()
{
	CBCGPGridCtrl::Init();

	InitStyle(GetStyle());
	InsertColumn(0, _T(""), max(OnGetColumnMinWidth(0), GetListRect().Width() - GetHierarchyOffset() - GetExtraHierarchyOffset()) );
	EnableColumnAutoSize(TRUE);
	AdjustLayout();
}
//****************************************************************************
void CBCGPTreeCtrlEx::OnStyleChanged (int nStyleType, LPSTYLESTRUCT lpStyleStruct)
{
	CBCGPGridCtrl::OnStyleChanged(nStyleType, lpStyleStruct);
	
	if (nStyleType & GWL_STYLE)
	{
		InitStyle(lpStyleStruct->styleNew);
	}
}
//****************************************************************************
void CBCGPTreeCtrlEx::InitStyle(DWORD style)
{
	// TVS_LINESATROOT is not supported
	EnableTreeLines((style & TVS_HASLINES) != 0);
	EnableTreeButtons((style & TVS_HASBUTTONS) != 0);

	BOOL bCheckBoxes = (style & TVS_CHECKBOXES) != 0;
	if (bCheckBoxes != m_bCheckBoxes)
	{
		m_bCheckBoxes = bCheckBoxes;
		GetColumnsInfo().EnableCheckBox(bCheckBoxes);
	}

	SetReadOnly((style & TVS_EDITLABELS) == 0);

 	BOOL bWholeRowSel = (style & TVS_FULLROWSELECT) != 0;
 	if (bWholeRowSel != m_bWholeRowSel)
 	{
		SetWholeRowSel((style & TVS_FULLROWSELECT) != 0);

		CBCGPGridItemID idSel = GetCurSelItemID();
		if (!idSel.IsNull())
		{
			SetCurSel(idSel, SM_SINGE_SEL_GROUP | SM_SINGLE_ITEM, FALSE);
		}
 	}

	AdjustLayout();
}
//****************************************************************************
CImageList* CBCGPTreeCtrlEx::GetImageList(UINT nImageList) const
{
	return (nImageList == TVSIL_NORMAL) ? CBCGPGridCtrl::GetImageList() : NULL;
}
//****************************************************************************
CImageList* CBCGPTreeCtrlEx::SetImageList(CImageList* pImageList, int nImageListType)
{
	if (nImageListType == TVSIL_NORMAL)
	{
		CImageList* pOld = CBCGPGridCtrl::GetImageList();
		CBCGPGridCtrl::SetImageList(pImageList);
		return pOld;
	}

	return NULL;
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::GetNextItem(HTREEITEM hItem, UINT nCode) const
{
	if (nCode == 0)
	{
		if (hItem == NULL || hItem == TVI_ROOT || hItem == TVI_FIRST || hItem == TVI_LAST || hItem == TVI_SORT)
		{
			nCode = TVGN_ROOT;
		}
	}

	if (hItem == NULL)
	{
		switch (nCode)
		{
		case TVGN_CHILD:
		case TVGN_NEXT:
		case TVGN_PREVIOUS:
		case TVGN_NEXTVISIBLE:
		case TVGN_PREVIOUSVISIBLE:
		case TVGN_PARENT:

			nCode = TVGN_ROOT;
			break;
		}
	}

	CBCGPGridRow* pRow = TreeItem(hItem);
	CBCGPGridRow* pResultRow = NULL;

	switch(nCode)
	{
	case TVGN_CARET:
		pResultRow = GetCurSel();
		break;

	case TVGN_CHILD:
		if (pRow == NULL)
		{
			return NULL;
		}

		ASSERT_VALID(pRow);
		pResultRow = pRow->GetSubItemsCount() > 0 ? pRow->GetSubItem(0) : NULL;
		break;

	case TVGN_DROPHILITE:	// not supported
		break;


	case TVGN_PARENT:
		if (pRow == NULL)
		{
			return NULL;
		}

		ASSERT_VALID(pRow);
		pResultRow = pRow->GetParent();
		break;

	case TVGN_ROOT:
		pResultRow = GetTotalItems() > 0 ? GetRow(0) : NULL;
		break;

	default:
		pResultRow = NextRow(pRow, nCode);
		break;
	}
	
	return TreeItem(pResultRow);
}
//****************************************************************************
CBCGPGridRow* CBCGPTreeCtrlEx::NextRow(CBCGPGridRow* pStartAfter, UINT nCode) const
{
	ASSERT_VALID(this);

	if (IsVirtualMode())
	{
		return NULL;
	}

	const CList<CBCGPGridRow*, CBCGPGridRow*>& lst = (!IsSortingMode() && !IsGrouping()) ? 
			m_lstItems : m_lstTerminalItems;

	switch (nCode)
	{
	case TVGN_NEXT:
	case TVGN_PREVIOUS:
	{
		BOOL bDirectionForward = (nCode == TVGN_NEXT);

		POSITION pos = (pStartAfter != NULL) ? lst.Find(pStartAfter) : NULL;
		if (pos != NULL)
		{
			bDirectionForward ? lst.GetNext(pos) : lst.GetPrev(pos);
		}

		CBCGPTreeCtrlEx* pThis = (CBCGPTreeCtrlEx*)this;
		
		while (pos != NULL)
		{
			CBCGPGridRow* pRow = bDirectionForward ? lst.GetNext(pos) : lst.GetPrev(pos);
			ASSERT_VALID(pRow);
			ASSERT_VALID(pStartAfter);

			if (pRow->GetParent() == pStartAfter->GetParent() && !pThis->IsRowFilteredByFilterBar(pRow)) // sibling
			{
				return pRow;
			}

			if (pStartAfter->GetParent() != NULL && !pRow->IsParent(pStartAfter->GetParent()))
			{
				return NULL; // row belongs to another parent
			}
		}

		return NULL;
	}
	
	case TVGN_NEXTVISIBLE:
	case TVGN_PREVIOUSVISIBLE:
	{
		if (m_rectList.IsRectEmpty())
		{
			return NULL;
		}
		
		BOOL bDirectionForward = (nCode == TVGN_NEXTVISIBLE);

		POSITION pos = (pStartAfter != NULL) ? lst.Find(pStartAfter) : NULL;
		if (pos != NULL)
		{
			bDirectionForward ? lst.GetNext(pos) : lst.GetPrev(pos);
		}
		
		while (pos != NULL)
		{
			CBCGPGridRow* pRow = bDirectionForward ? lst.GetNext(pos) : lst.GetPrev(pos);
			ASSERT_VALID(pRow);
	
			CRect rectRow = pRow->GetRect();

			if (!rectRow.IsRectEmpty() && IsRowVisible(pRow)) // only visible rows
			{
				// row is inside visible area
				if (rectRow.bottom >= m_rectList.top && rectRow.top <= m_rectList.bottom)
				{
					return pRow;
				}
			}
		}
	
		return NULL;
	}
	
	case TVGN_FIRSTVISIBLE:
	case TVGN_LASTVISIBLE:
	{
		if (m_rectList.IsRectEmpty())
		{
			return NULL;
		}
		
		CBCGPGridRow* pVisible = NULL;
	
		POSITION pos = lst.GetHeadPosition();
		while (pos != NULL)
		{
			CBCGPGridRow* pRow = lst.GetNext(pos);
			ASSERT_VALID(pRow);
	
			CRect rectRow = pRow->GetRect();

			if (!rectRow.IsRectEmpty() && IsRowVisible(pRow)) // only visible rows
			{
				// row is inside visible area
				if (rectRow.bottom >= m_rectList.top && rectRow.top <= m_rectList.bottom)
				{
					if (nCode == TVGN_FIRSTVISIBLE)
					{
						return pRow;
					}
		
					pVisible = pRow;
				}
				
				// row is after the last visible - stop search
				else if (rectRow.top > m_rectList.bottom)
				{
					break;
				}
			}
		}
		
		return pVisible;
	}

	//   TVGN_NEXTSELECTED is not supported
	default:
		return NULL;
	}
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::GetChildItem(HTREEITEM hItem) const
{
	return GetNextItem(hItem, TVGN_CHILD);
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::GetNextSiblingItem(HTREEITEM hItem) const
{
	return GetNextItem(hItem, TVGN_NEXT);
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::GetPrevSiblingItem(HTREEITEM hItem) const
{
	return GetNextItem(hItem, TVGN_PREVIOUS);
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::GetParentItem(HTREEITEM hItem) const
{
	return GetNextItem(hItem, TVGN_PARENT);
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::GetFirstVisibleItem() const
{
	return GetNextItem(NULL, TVGN_FIRSTVISIBLE);
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::GetNextVisibleItem(HTREEITEM hItem) const
{
	return GetNextItem(hItem, TVGN_NEXTVISIBLE);
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::GetPrevVisibleItem(HTREEITEM hItem) const
{
	return GetNextItem(hItem, TVGN_PREVIOUSVISIBLE);
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::GetSelectedItem() const
{
	return GetNextItem(NULL, TVGN_CARET);
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::GetRootItem() const
{
	return GetNextItem(NULL, TVGN_ROOT);
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::GetItem(TVITEM* pItem) const
{
	ASSERT_VALID(this);

	if (pItem == NULL)
	{
		return FALSE;
	}

	CBCGPGridRow* pRow = TreeItem(pItem->hItem);
	if (pRow == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID(pRow);

	//----------------------
	// Get row's properties:
	//----------------------
	if ((pItem->mask & TVIF_CHILDREN) != 0)
	{
		pItem->cChildren = pRow->GetSubItemsCount();
	}

	if ((pItem->mask & TVIF_IMAGE) != 0)
	{
		CBCGPGridItem* pGridItem = pRow->GetItem(0);
		pItem->iImage = (pGridItem != NULL) ? pGridItem->GetImage() : 0;
	}

	if ((pItem->mask & TVIF_SELECTEDIMAGE) != 0)
	{
		// TVIF_SELECTEDIMAGE is not supported
		pItem->iSelectedImage = 0;
	}

	if ((pItem->mask & TVIF_STATE) != 0)
	{
		pItem->state = 0;

		if ((pItem->stateMask & TVIS_SELECTED) != 0)
		{
			if (pRow->IsSelected())
			{
				pItem->state |= TVIS_SELECTED;
			}
		}

		if ((pItem->stateMask & TVIS_EXPANDED) != 0)
		{
			if (pRow->IsGroup() && pRow->IsExpanded())
			{
				pItem->state |= TVIS_EXPANDED;
			}
		}

		if ((pItem->stateMask & TVIS_BOLD) != 0)
		{
			if (pRow->IsBoldFont())
			{
				pItem->state |= TVIS_BOLD;
			}
		}

		if ((pItem->stateMask & TVIS_STATEIMAGEMASK) != 0)
		{
			if (IsCheckBoxesEnabled() && pRow->HasCheckBox())
			{
				int nStateImageIndex = (pRow->GetCheck() == BST_INDETERMINATE) ? 3 : (pRow->GetCheck() ? 2 : 1);
				pItem->state |= INDEXTOSTATEIMAGEMASK(nStateImageIndex);
			}
		}

		// TVIS_EXPANDPARTIAL is not supported
		// TVIS_CUT is not supported
		// TVIS_DROPHILITED is not supported
		// TVIS_EXPANDPARTIAL is not supported
		// TVIS_OVERLAYMASK is not supported
	}

	if ((pItem->mask & TVIF_PARAM) != 0)
	{
		pItem->lParam = (LPARAM) pRow->GetData();
	}

	if ((pItem->mask & TVIF_TEXT) != 0)
	{
		CString strText;

		CBCGPGridItem* pGridItem = pRow->GetItem(0);
		if (pGridItem != NULL)
		{
			strText = pGridItem->FormatItem();
		}

		if (pItem->cchTextMax < strText.GetLength() + 1)
		{
			return FALSE;
		}

#if _MSC_VER < 1400
		_tcsncpy(pItem->pszText, (LPCTSTR)strText, strText.GetLength());
#else
		_tcsncpy_s(pItem->pszText, pItem->cchTextMax, (LPCTSTR)strText, strText.GetLength());
#endif
	}

	return TRUE;
}
//****************************************************************************
CString CBCGPTreeCtrlEx::GetItemText(HTREEITEM hItem) const
{
	ASSERT_VALID(this);

	CBCGPGridRow* pRow = TreeItem(hItem);
	if (pRow != NULL)
	{
		ASSERT_VALID(pRow);
		
		CBCGPGridItem* pGridItem = pRow->GetItem(0);
		if (pGridItem != NULL)
		{
			return pGridItem->FormatItem();
		}
	}
	
	return CString();
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::GetItemImage(HTREEITEM hItem, int& nImage, int& nSelectedImage) const
{
	TVITEM item;
	memset(&item, 0, sizeof(item));
	item.mask = TVIF_HANDLE | TVIF_IMAGE | TVIF_SELECTEDIMAGE;
	item.hItem = hItem;
	BOOL bResult = GetItem(&item);

	nImage = item.iImage;
	nSelectedImage = item.iSelectedImage;

	return bResult;
}
//****************************************************************************
UINT CBCGPTreeCtrlEx::GetItemState(HTREEITEM hItem, UINT nStateMask) const
{
	TVITEM item;
	memset(&item, 0, sizeof(item));
	item.mask = TVIF_HANDLE | TVIF_STATE;
	item.hItem = hItem;
	item.stateMask = nStateMask;
	GetItem(&item);
	
	return (item.state & nStateMask);
}
//****************************************************************************
DWORD_PTR CBCGPTreeCtrlEx::GetItemData(HTREEITEM hItem) const
{
	TVITEM item;
	memset(&item, 0, sizeof(item));
	item.mask = TVIF_HANDLE | TVIF_PARAM;
	item.hItem = hItem;
	GetItem(&item);
	
	return (DWORD_PTR)item.lParam;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SetItem(TVITEM* pItem)
{
	if (pItem == NULL || (pItem->mask & TVIF_HANDLE) == 0)
	{
		return FALSE;
	}
	
	return SetItem(pItem->hItem, pItem->mask, pItem->pszText, pItem->iImage,
		pItem->iSelectedImage, pItem->state, pItem->stateMask, pItem->lParam);
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SetItem(HTREEITEM hItem, UINT nMask, LPCTSTR lpszItem, int nImage,
			 int nSelectedImage, UINT nState, UINT nStateMask, LPARAM lParam)
{
	ASSERT_VALID(this);

	if (hItem == NULL)
	{
		return FALSE;
	}

	CBCGPGridRow* pRow = TreeItem(hItem);
	if (pRow == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID(pRow);

	//----------------------
	// Set row's properties:
	//----------------------
	if ((nMask & TVIF_TEXT) != 0)
	{
		CBCGPGridItem* pItem = pRow->GetItem(0);
		if (pItem != NULL)
		{
			pItem->SetValue(_variant_t (lpszItem != NULL ? lpszItem : (LPCTSTR) _T("")), FALSE);
		}
	}

	if ((nMask & TVIF_IMAGE) != 0)
	{
		CBCGPGridItem* pItem = pRow->GetItem(0);
		if (pItem != NULL)
		{
			pItem->SetImage(nImage, FALSE);
		}
	}

	if ((nMask & TVIF_SELECTEDIMAGE) != 0)
	{
		// TVIF_SELECTEDIMAGE is not supported
		UNREFERENCED_PARAMETER(nSelectedImage);
	}

	if ((nMask & TVIF_STATE) != 0)
	{
		if ((nStateMask & TVIS_SELECTED) != 0)
		{
			BOOL bSelect = (nState & TVIS_SELECTED) != 0;
			if (bSelect != pRow->IsSelected())
			{
				pRow->Select(bSelect);
			}
		}

		if ((nStateMask & TVIS_EXPANDED) != 0)
		{
			BOOL bExpand = (nState & TVIS_EXPANDED) != 0;
			if (bExpand != pRow->IsExpanded())
			{
				if (!Expand(hItem, bExpand ? TVE_EXPAND : TVE_COLLAPSE))
				{
					return FALSE;
				}
			}
		}

		if ((nStateMask & TVIS_BOLD) != 0)
		{
			pRow->SetBoldFont((nState & TVIS_BOLD) != 0);
		}

		if ((nStateMask & TVIS_STATEIMAGEMASK) != 0)
		{
			int nStateImage = (BOOL)(nState >> 12) -1;
			pRow->SetCheck((nStateImage == BST_INDETERMINATE) ? 2 : (nStateImage == BST_CHECKED ? 1 : 0));
		}

		// TVIS_EXPANDPARTIAL is not supported
		// TVIS_CUT is not supported
		// TVIS_DROPHILITED is not supported
		// TVIS_EXPANDPARTIAL is not supported
		// TVIS_OVERLAYMASK is not supported
	}

	if ((nMask & TVIF_PARAM) != 0)
	{
		pRow->SetData(lParam);
	}

	return TRUE;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SetItemText(HTREEITEM hItem, LPCTSTR lpszItem)
{
	return SetItem(hItem, TVIF_TEXT, lpszItem, 0, 0, 0, 0, 0);
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SetItemImage(HTREEITEM hItem, int nImage, int nSelectedImage)
{
	return SetItem(hItem, TVIF_IMAGE | TVIF_SELECTEDIMAGE, NULL, nImage, nSelectedImage, 0, 0, 0);
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SetItemState(HTREEITEM hItem, UINT nState, UINT nStateMask)
{
	return SetItem(hItem, TVIF_STATE, NULL, 0, 0, nState, nStateMask, 0);
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SetItemData(HTREEITEM hItem, DWORD_PTR dwData)
{
	return SetItem(hItem, TVIF_PARAM, NULL, 0, 0, 0, 0, (LPARAM)dwData);
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::ItemHasChildren(HTREEITEM hItem) const
{
	CBCGPGridRow* pRow = TreeItem(hItem);
	if (pRow != NULL)
	{
		ASSERT_VALID(pRow);
		return pRow->IsGroup() && pRow->GetSubItemsCount() > 0;
	}

	return FALSE;
}
//****************************************************************************
CEdit* CBCGPTreeCtrlEx::GetEditControl() const
{
	ASSERT_VALID(this);
	
	CBCGPGridRow* pSel = GetCurSel();
	if (pSel == NULL)
	{
		return NULL;
	}
	
	return DYNAMIC_DOWNCAST(CEdit, pSel->GetInPlaceWnd());
}
//****************************************************************************
UINT CBCGPTreeCtrlEx::GetVisibleCount() const
{
	ASSERT_VALID(this);
	return (GetRowHeight() > 0) ? m_rectList.Height() / GetRowHeight() : 0;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::GetCheck(HTREEITEM hItem) const
{
	return ((BOOL)(GetItemState(hItem, TVIS_STATEIMAGEMASK) >> 12) -1);
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SetCheck(HTREEITEM hItem, BOOL fCheck)
{
	return SetItemState(hItem, INDEXTOSTATEIMAGEMASK((fCheck == BST_INDETERMINATE) ? 3 : (fCheck ? 2 : 1)), TVIS_STATEIMAGEMASK);
}
//****************************************************************************
void CBCGPTreeCtrlEx::EnableTreeCtrlNotifications(BOOL bEnable /*= TRUE*/)
{
	m_bTreeCtrlNotifications = bEnable;
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::InsertItem(LPTVINSERTSTRUCT lpInsertStruct)
{
	ASSERT(lpInsertStruct != NULL);
	
	return InsertItem(
		lpInsertStruct->item.mask, 
		(LPTSTR)lpInsertStruct->item.pszText,
		lpInsertStruct->item.iImage,
		lpInsertStruct->item.iSelectedImage,
		lpInsertStruct->item.state,
		lpInsertStruct->item.stateMask,
		lpInsertStruct->item.lParam,
		lpInsertStruct->hParent,
		lpInsertStruct->hInsertAfter);
}
//****************************************************************************
class CTreeRow : public CBCGPGridRow
{
	friend CBCGPTreeCtrlEx;
};

HTREEITEM CBCGPTreeCtrlEx::InsertItem(UINT nMask, 
									  LPCTSTR lpszItem, int nImage, int nSelectedImage, 
									  UINT nState, UINT nStateMask, LPARAM lParam, 
									  HTREEITEM hParent, HTREEITEM hInsertAfter)
{
	ASSERT_VALID(this);

	//--------------
	// Create a row:
	//--------------
	CBCGPGridRow* pRow = CreateRow ();
	ASSERT_VALID (pRow);
	
	pRow->SetOwnerList (this);

	const int nColumnsNum = GetColumnsInfo ().GetColumnCount ();
	for (int i = 0; i < nColumnsNum ; i++)
	{
		CBCGPGridItem* pItem = pRow->CreateItem (pRow->GetItemCount(), i);
		if (pItem != NULL)
		{
			pRow->AddItem (pItem);
		}
	}

	//----------------------
	// Set row's properties:
	//----------------------
	if ((nMask & TVIF_TEXT) != 0)
	{
		CBCGPGridItem* pItem = pRow->GetItem(0);
		if (pItem != NULL)
		{
			pItem->SetValue(_variant_t (lpszItem != NULL ? lpszItem : (LPCTSTR) _T("")), FALSE);
		}
	}
	
	if ((nMask & TVIF_IMAGE) != 0)
	{
		CBCGPGridItem* pItem = pRow->GetItem(0);
		if (pItem != NULL)
		{
			pItem->SetImage(nImage, FALSE);
		}
	}
	
	if ((nMask & TVIF_SELECTEDIMAGE) != 0)
	{
		// TVIF_SELECTEDIMAGE is not supported
		UNREFERENCED_PARAMETER(nSelectedImage);
	}
	
	if ((nMask & TVIF_STATE) != 0)
	{
		if ((nStateMask & TVIS_SELECTED) != 0)
		{
			BOOL bSelect = (nState & TVIS_SELECTED) != 0;
			if (bSelect != pRow->IsSelected())
			{
				pRow->Select(bSelect);
			}
		}
		
		if ((nStateMask & TVIS_EXPANDED) != 0)
		{
			((CTreeRow*)pRow)->m_bExpanded = (nState & TVIS_EXPANDED) != 0;
		}
		
		if ((nStateMask & TVIS_BOLD) != 0)
		{
			((CTreeRow*)pRow)->m_bBold = (nState & TVIS_BOLD) != 0;
		}

		// TVIS_EXPANDPARTIAL is not supported
		// TVIS_CUT is not supported
		// TVIS_DROPHILITED is not supported
		// TVIS_EXPANDPARTIAL is not supported
		// TVIS_OVERLAYMASK is not supported
		// TVIS_STATEIMAGEMASK is not supported
	}
	
	if ((nMask & TVIF_PARAM) != 0)
	{
		pRow->SetData(lParam);
	}

	//----------------
	// Insert the row:
	//----------------
	// 1) find insertion place
	POSITION posInsertAfter = NULL;
	POSITION posParent = NULL;
	POSITION posInsertAfterSubItem = NULL;
	ULONG_PTR nFlags = 0;

	if (hParent == TVI_ROOT || hParent == TVI_FIRST || hParent == TVI_LAST || hParent == TVI_SORT)
	{
		hParent = NULL;
	}
	
	if (hInsertAfter == TVI_ROOT || hInsertAfter == TVI_FIRST || hInsertAfter == TVI_LAST || hInsertAfter == TVI_SORT)
	{
		nFlags = (ULONG_PTR)hInsertAfter;
		hInsertAfter = NULL;
	}
	
	GetPositionInsertAfter(TreeItem(hParent), TreeItem(hInsertAfter), nFlags, posInsertAfter, posParent, posInsertAfterSubItem);

	// 2) insert to m_lstItems
	if (InsertAfterInternal(pRow, posInsertAfter, posParent, posInsertAfterSubItem) == NULL)
	{
		delete pRow;
		return NULL;
	}

	if (IsWindowVisible() || hInsertAfter != NULL)
	{
		// 3) update m_lstTerminalItems
		SetRebuildTerminalItems();
		ReposItems();

		m_bNeedToRepos = FALSE;
	}
	else
	{
		m_bNeedToRepos = TRUE;
	}

	return TreeItem(pRow);
}
//****************************************************************************
void CBCGPTreeCtrlEx::GetPositionInsertAfter(CBCGPGridRow* pParent, CBCGPGridRow* pInsertAfter, ULONG_PTR nFlags,
												POSITION& pos, POSITION& posParent, POSITION& posSub) const
{
	ASSERT_VALID(this);

	pos = NULL;			// position in m_lstItems
	posParent = NULL;	// position in m_lstItems
	posSub = NULL;		// position in pParent->m_lstSubItems

	//--------------
	// Check Parent:
	//--------------
	if (pParent != NULL)
	{
		ASSERT_VALID(pParent);
		posParent = m_lstItems.Find(pParent);
	}

	BOOL bInsertAsChild = (posParent != NULL); // Parent is found. pParent is a valid parent item. Use it.

	//-------------------
	// Check InsertAfter:
	//-------------------
	if (bInsertAsChild)
	{
		POSITION posSubItem = NULL;
		if (pInsertAfter != NULL && pParent != NULL)
		{
			ASSERT_VALID(pInsertAfter);
			posSubItem = ((CTreeRow*)pParent)->m_lstSubItems.Find(pInsertAfter);
		}

		POSITION posInsertAfterFound = NULL;
		if (posSubItem != NULL) 
		{
			posInsertAfterFound = m_lstItems.Find(pInsertAfter, posParent);
		}

		if (posInsertAfterFound != NULL) // pInsertAfter is a valid sub item.
		{
			// insert at the specified position
			pos = GetPositionOfLastSubItem(posInsertAfterFound, TRUE);
			posSub = posSubItem;
		}
		else // pInsertAfter is not valid. Must insert as last child by default.
		{
			if ((HTREEITEM)nFlags == TVI_FIRST) // insert as first child
			{
				pos = posParent;
				posSub = NULL;
			}
			else
			{
				pos = GetPositionOfLastSubItem(posParent, TRUE);
				posSub = ((CTreeRow*)pParent)->m_lstSubItems.GetTailPosition();
			}
		}
	}
	else // insert at root
	{
		if ((HTREEITEM)nFlags == TVI_FIRST)
		{
			pos = NULL; // insert as first
			posSub = NULL;
			return;
		}
		else
		{
			pos = GetPositionOfLastSubItem(m_lstItems.Find(pInsertAfter), TRUE);
			posSub = NULL;
		}
	}

	// By default must insert as last root item if pos == NULL
	if (pos == NULL)
	{
		pos = !m_lstItems.IsEmpty() ? m_lstItems.GetTailPosition() : NULL;
		posParent = NULL;
		posSub = NULL;
	}
}
//****************************************************************************
POSITION CBCGPTreeCtrlEx::GetPositionOfLastSubItem(POSITION posParent, BOOL bRecursive) const
{
	ASSERT_VALID(this);

	if (posParent == NULL)
	{
		return NULL;
	}

	CBCGPGridRow* pSubItem = NULL;
	if (posParent != NULL)
	{
		CBCGPGridRow* pParent = !m_lstItems.IsEmpty() ? m_lstItems.GetAt(posParent) : NULL;
		if (pParent != NULL)
		{
			ASSERT_VALID(pParent);

			if (!((CTreeRow*)pParent)->m_lstSubItems.IsEmpty())
			{
				pSubItem = ((CTreeRow*)pParent)->m_lstSubItems.GetTail();
			}
		}
	}

	POSITION posSubItem = NULL; // position in m_lstItems
	if (pSubItem != NULL)
	{
		ASSERT_VALID(pSubItem);
		posSubItem = m_lstItems.Find(pSubItem, posParent);
	}

	if (posSubItem == NULL)
	{
		return posParent;
	}

	return !bRecursive ? posSubItem : GetPositionOfLastSubItem(posSubItem, bRecursive);
}
//****************************************************************************
POSITION CBCGPTreeCtrlEx::InsertAfterInternal(CBCGPGridRow* pNew, POSITION pos, POSITION posParent, POSITION posSub)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pNew);
	
	if (m_bVirtualMode)
	{
		// Cannot insert row in the virtual mode
		ASSERT (FALSE);
		return NULL;
	}

#ifdef _DEBUG
	if (CBCGPGridCtrl::m_bEnableAssertValidInDebug)
	{
		for (POSITION posCurr = m_lstItems.GetHeadPosition (); posCurr != NULL; )
		{
			CBCGPGridRow* pListItem = m_lstItems.GetNext (posCurr);
			ASSERT_VALID (pListItem);
			
			if (pListItem == pNew || pListItem->IsSubItem (pNew))
			{
				// Can't ad the same item twice
				ASSERT (FALSE);
				return NULL;
			}
		}
	}
#endif // _DEBUG

	POSITION posResult = NULL;
	if (pos != NULL)
	{
		// insert after
		posResult = m_lstItems.InsertAfter(pos, pNew);

		CBCGPGridRow* pPrev = m_lstItems.GetAt(pos);
		if (pPrev != NULL && pPrev->GetRowId() != -1)
		{
			((CTreeRow*)pNew)->m_nIdRow = pPrev->GetRowId() + 1;
		}
	}
	else
	{
		// add first
		posResult = m_lstItems.AddHead(pNew);
		((CTreeRow*)pNew)->m_nIdRow = 0;
	}

	if (posResult == NULL)
	{
		return NULL;
	}

	pNew->SetOwnerList (this);

	if (posParent != NULL)
	{
		CBCGPGridRow* pParent = m_lstItems.GetAt(posParent);
		if (pParent != NULL)
		{
			POSITION posSubResult = NULL;
			if (posSub != NULL)
			{
				posSubResult = ((CTreeRow*)pParent)->m_lstSubItems.InsertAfter(posSub, pNew);
			}
			else
			{
				posSubResult = ((CTreeRow*)pParent)->m_lstSubItems.AddHead(pNew);
			}
			
			if (posSubResult != NULL)
			{
				pParent->AllowSubItems(TRUE);
				((CTreeRow*)pNew)->SetParent(pParent);
			}
		}
	}

	return posResult;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SendTreeNotification(UINT nCode, CBCGPGridRow* pOld, CBCGPGridRow* pNew, UINT nAction) const
{
	if (!m_bTreeCtrlNotifications)
	{
		return TRUE;
	}

	if (pOld == NULL && pNew == NULL)
	{
		return TRUE;
	}

	NMTREEVIEW tv;
	memset(&tv, 0, sizeof(NMTREEVIEW));
	
	tv.hdr.hwndFrom = GetSafeHwnd();
	tv.hdr.idFrom = GetDlgCtrlID();
	tv.hdr.code = nCode;
	
	tv.action = nAction;

	UINT nItemMask = TVIF_HANDLE | TVIF_PARAM;
	UINT nStateMask = 0;

	switch (nCode)
	{
	case TVN_SELCHANGED:
	case TVN_SELCHANGING:
		if (pOld == pNew)
		{
			return TRUE;
		}

		nItemMask |= TVIF_STATE;
		nStateMask |= TVIS_SELECTED;
		break;

	case TVN_ITEMEXPANDING:
	case TVN_ITEMEXPANDED:
		nItemMask |= (TVIF_STATE | TVIF_IMAGE);
		nStateMask |= (TVIS_EXPANDED | TVIS_SELECTED);
		break;
	}

	if (pOld != NULL)
	{
		tv.itemOld.hItem = TreeItem(pOld);
		tv.itemOld.mask = nItemMask;
		tv.itemOld.stateMask = nStateMask;
		GetItem(&tv.itemOld);
	}

	if (pNew != NULL)
	{
		tv.itemNew.hItem = TreeItem(pNew);
		tv.itemNew.mask = nItemMask;
		tv.itemNew.stateMask = nStateMask;
		GetItem(&tv.itemNew);
	}
		
	CWnd* pOwner = GetOwner();
	if (pOwner->GetSafeHwnd() != NULL && IsWindow(pOwner->GetSafeHwnd()))
	{
		return pOwner->SendMessage(WM_NOTIFY, tv.hdr.idFrom, (LPARAM)&tv) == 0;
	}
	
    return FALSE;
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::InsertItem(LPCTSTR lpszItem, HTREEITEM hParent, HTREEITEM hInsertAfter)
{
	return InsertItem(TVIF_TEXT, lpszItem, 0, 0, 0, 0, 0, hParent, hInsertAfter);
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::InsertItem(LPCTSTR lpszItem, int nImage, int nSelectedImage, HTREEITEM hParent, HTREEITEM hInsertAfter)
{
	return InsertItem(TVIF_TEXT|TVIF_IMAGE|TVIF_SELECTEDIMAGE, lpszItem, nImage, nSelectedImage, 0, 0, 0, hParent, hInsertAfter);
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SetColumnValue(HTREEITEM hItem, int nSubItem, const _variant_t& varValue)
{
	ASSERT_VALID(this);

	if (hItem == NULL || hItem == TVI_ROOT)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	CBCGPGridRow* pRow = TreeItem(hItem);
	if (pRow != NULL)
	{
		ASSERT_VALID(pRow);

		CBCGPGridItem* pItem = pRow->GetItem(nSubItem);
		if (pItem != NULL)
		{
			ASSERT_VALID(pItem);

			pItem->SetValue(varValue, FALSE);

			return TRUE;
		}
	}

	return FALSE;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SetColumnValue(HTREEITEM hItem, int nSubItem, const _variant_t& varValue, int nImage)
{
	ASSERT_VALID(this);
	
	if (hItem == NULL || hItem == TVI_ROOT)
	{
		ASSERT(FALSE);
		return FALSE;
	}
	
	CBCGPGridRow* pRow = TreeItem(hItem);
	if (pRow != NULL)
	{
		ASSERT_VALID(pRow);
		
		CBCGPGridItem* pItem = pRow->GetItem(nSubItem);
		if (pItem != NULL)
		{
			ASSERT_VALID(pItem);
			
			pItem->SetValue(varValue, FALSE);
			pItem->SetImage(nImage, FALSE);
			
			return TRUE;
		}
	}
	
	return FALSE;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::ReplaceColumnValue(HTREEITEM hItem, int nSubItem, CBCGPGridItem* pNewItem)
{
	ASSERT_VALID(this);
	
	if (hItem == NULL || hItem == TVI_ROOT || pNewItem == NULL)
	{
		ASSERT(FALSE);
		return FALSE;
	}
	
	CBCGPGridRow* pRow = TreeItem(hItem);
	if (pRow != NULL)
	{
		ASSERT_VALID(pRow);
		ASSERT_VALID(pNewItem);
		
		return pRow->ReplaceItem(nSubItem, pNewItem, FALSE);
	}
	
	return FALSE;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::DeleteItem(HTREEITEM hItem)
{
	ASSERT_VALID(this);
	
	if (hItem == NULL || hItem == TVI_ROOT)
	{
		return DeleteAllItems();
	}
	
	CBCGPGridRow* pRow = TreeItem(hItem);
	if (pRow != NULL)
	{
		ASSERT_VALID(pRow);

		int nDelCount = 0;

		POSITION posDel = m_lstItems.Find(pRow);
		if (posDel != NULL)
		{
			CList<CBCGPGridRow*, CBCGPGridRow*>	lstSubItems;
			pRow->GetSubItems(lstSubItems, TRUE);

			while (!lstSubItems.IsEmpty())
			{
				CBCGPGridRow* pSub = lstSubItems.RemoveHead();
				ASSERT_VALID(pSub);

				POSITION posSub = m_lstItems.Find(pSub, posDel);
				if (posSub != NULL)
				{
					SendTreeNotification(TVN_DELETEITEM, pSub, NULL);

					if (DoRemoveRow(posSub, FALSE))
					{
						nDelCount++;
					}
				}
			}

			SendTreeNotification(TVN_DELETEITEM, pRow, NULL);

			if (DoRemoveRow(posDel, TRUE))
			{
				nDelCount++;
			}

			return nDelCount > 0;
		}
	}

	return FALSE;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::DeleteAllItems()
{
	RemoveAll();
	return TRUE;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::Expand(HTREEITEM hItem, UINT nCode)
{
	CBCGPGridRow* pRow = TreeItem(hItem);
	if (pRow == NULL)
	{
		return FALSE;
	}

	ASSERT_VALID(pRow);

	if (!pRow->IsGroup() || pRow->GetSubItemsCount() == 0)
	{
		// Attempting to expand or collapse a node that has no children is considered a failure
		return FALSE;
	}

	if ((nCode & TVE_EXPAND) != 0 || (nCode & TVE_TOGGLE) != 0 && !pRow->IsExpanded())
	{
		if (pRow->IsExpanded())
		{
			// Expanding a node that is already expanded is considered a successful operation
			return TRUE;
		}

		// TVE_EXPANDPARTIAL is not supported

		if (!SendTreeNotification(TVN_ITEMEXPANDING, NULL, pRow, nCode))
		{
			return FALSE;
		}

		pRow->Expand(TRUE);
	}
	else if ((nCode & TVE_COLLAPSE) != 0 || (nCode & TVE_TOGGLE) != 0 && pRow->IsExpanded())
	{
		if (!pRow->IsExpanded())
		{
			// Collapsing a node returns zero if the node is already collapsed
			return FALSE;
		}

		if (!SendTreeNotification(TVN_ITEMEXPANDING, NULL, pRow, nCode))
		{
			return FALSE;
		}

		if ((nCode & TVE_COLLAPSERESET) != 0)
		{
			// remove all child items
			CList<CBCGPGridRow*, CBCGPGridRow*>	lstSubItems;
			pRow->GetSubItems(lstSubItems, TRUE);
			
			POSITION posParent = m_lstItems.Find(pRow);
			int nDelCount = 0;

			while (!lstSubItems.IsEmpty())
			{
				CBCGPGridRow* pSub = lstSubItems.RemoveHead();
				ASSERT_VALID(pSub);
				
				POSITION posSub = m_lstItems.Find(pSub, posParent);
				if (posSub != NULL)
				{
					SendTreeNotification(TVN_DELETEITEM, pSub, NULL);

					if (DoRemoveRow(posSub, FALSE))
					{
						nDelCount++;
					}
				}
			}
		}

		pRow->Expand(FALSE);
	}

	SendTreeNotification(TVN_ITEMEXPANDED, NULL, pRow, nCode);
	return TRUE;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::Select(HTREEITEM hItem, UINT nCode)
{
	ASSERT_VALID(this);

	CBCGPGridRow* pRow = TreeItem(hItem);

	if (pRow != NULL && IsRowFilteredByFilterBar(pRow))
	{
		return FALSE;
	}

	if (nCode == TVGN_CARET)
	{
		CBCGPGridRow* pOldSel = GetCurSel();

		if (!SendTreeNotification(TVN_SELCHANGING, pOldSel, pRow))
		{
			return FALSE;
		}

		SetCurSel(pRow, pRow == NULL);
		
		if (pOldSel != NULL && pOldSel != pRow)
		{
			pOldSel->Redraw();
		}

		if (pRow != NULL)
		{
			CBCGPGridCtrl::EnsureVisible(pRow, TRUE);
			pRow->Redraw();
		}

		SendTreeNotification(TVN_SELCHANGED, pOldSel, pRow);
		return TRUE;
	}
	else if (nCode == TVGN_FIRSTVISIBLE)
	{
		if (pRow != NULL)
		{
			ASSERT_VALID(pRow);

			m_nVertScrollOffset = 0;
			CBCGPGridCtrl::EnsureVisible(pRow, TRUE);

			pRow->Redraw();
			return TRUE;
		}
	}

	// TVGN_DROPHILITE code is not supported
	
	return FALSE;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SelectItem(HTREEITEM hItem)
{
	return Select(hItem, TVGN_CARET);
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SelectSetFirstVisible(HTREEITEM hItem)
{
	return Select(hItem, TVGN_FIRSTVISIBLE);
}
//****************************************************************************
CEdit* CBCGPTreeCtrlEx::EditLabel(HTREEITEM hItem)
{
	CBCGPGridRow* pRow = TreeItem(hItem);
	if (pRow != NULL)
	{
		ASSERT_VALID(pRow);

		if (!pRow->IsSelected())
		{
			SetCurSel(pRow, FALSE);
		}

		if (GetCurSelItem() == NULL)
		{
			SetCurSel(CBCGPGridItemID(pRow->GetRowId(), 0), SM_SET_ACTIVE_ITEM);
		}

		SendMessage(WM_KEYDOWN, VK_F2);
		// should send a TVN_BEGINLABELEDIT notification

		return (CEdit*) pRow->GetInPlaceWnd();
	}
	
	return NULL;
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::HitTest(CPoint pt, UINT* pFlags) const
{
	ASSERT_VALID(this);

	HTREEITEM hItem = NULL;
	UINT nFlags = TVHT_NOWHERE;

	if (pt.y < m_rectList.top)
	{
		nFlags = TVHT_ABOVE;
	}
	else if (pt.y > m_rectList.bottom)
	{
		nFlags = TVHT_BELOW;
	}
	else if (pt.x < m_rectList.left)
	{
		nFlags = TVHT_TOLEFT;
	}
	else if (pt.x > m_rectList.right)
	{
		nFlags = TVHT_TORIGHT;
	}
	else
	{
		CBCGPGridRow::ClickArea nArea;
		CBCGPGridRow* pRow = ((CBCGPGridCtrl*)this)->HitTest(pt, &nArea, TRUE);

		if (pRow != NULL)
		{
			if (nArea == CBCGPGridRow::ClickExpandBox)
			{
				nFlags = TVHT_ONITEMBUTTON;
			}
			else if (nArea == CBCGPGridRow::ClickCheckBox)
			{
				nFlags = TVHT_ONITEMSTATEICON;
			}
			else if (nArea == CBCGPGridRow::ClickName)
			{
				// TVHT_ONITEMINDENT, TVHT_ONITEMRIGHT
				nFlags = TVHT_ONITEMLABEL;
			}
			else if (nArea == CBCGPGridRow::ClickValue)
			{
				// TVHT_ONITEM, TVHT_ONITEMICON, TVHT_ONITEMLABEL
				nFlags = TVHT_ONITEMLABEL;
			}

			ASSERT_VALID(pRow);
			hItem = TreeItem(pRow);
		}
	}

	if (pFlags != NULL)
	{
		*pFlags = nFlags;
	}

	return hItem;
}
//****************************************************************************
HTREEITEM CBCGPTreeCtrlEx::HitTest(TVHITTESTINFO* pHitTestInfo) const
{
	ASSERT(pHitTestInfo != NULL);
	pHitTestInfo->hItem = HitTest(pHitTestInfo->pt, &pHitTestInfo->flags);
	return pHitTestInfo->hItem;
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::EnsureVisible(HTREEITEM hItem)
{
	CBCGPGridRow* pRow = TreeItem(hItem);
	if (pRow != NULL)
	{
		CBCGPGridCtrl::EnsureVisible(pRow, TRUE);
		return TRUE;
	}
	
	return FALSE;
}
//****************************************************************************
UINT CBCGPTreeCtrlEx::GetCount() const
{
	return (UINT)CBCGPGridCtrl::GetTotalRowCount();
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::GetItemRect(HTREEITEM hItem, LPRECT lpRect, BOOL bTextOnly) const
{
	ASSERT_VALID(this);

	BOOL bRes = FALSE;

	if (lpRect != NULL)
	{
		CBCGPGridRow* pRow = TreeItem(hItem);
		if (pRow != NULL)
		{
			ASSERT_VALID(pRow);

			if (bTextOnly)
			{
				CBCGPGridItem* pGridItem = pRow->GetItem(0);
				if (pGridItem != NULL)
				{
					*lpRect = pGridItem->GetRect();
					bRes = TRUE;
				}
			}
			else
			{
				*lpRect = pRow->GetRect();
				bRes = TRUE;
			}
		}
	}
	
	return bRes;
}
//****************************************************************************
void CBCGPTreeCtrlEx::OnDraw(CDC* pDCPaint)
{
	if (m_bNeedToRepos)
	{
		CBCGPGridItemID idSel = GetCurSelItemID();

		SetRebuildTerminalItems();
		ReposItems();

		m_bNeedToRepos = FALSE;

		if (!idSel.IsNull())
		{
			CBCGPGridCtrl::SetCurSel(idSel, SM_SINGE_SEL_GROUP | SM_SINGLE_ITEM, FALSE);
		}
	}

	CBCGPGridCtrl::OnDraw(pDCPaint);
}
//****************************************************************************
BOOL CBCGPTreeCtrlEx::SetCurSel(CBCGPGridItemID idItem, DWORD dwSelMode /*= SM_SINGE_SEL_GROUP | SM_SINGLE_ITEM*/, BOOL bRedraw /*= TRUE*/)
{
	CBCGPGridRow* pOldSel = GetCurSel();
	CBCGPGridRow* pNewSel = GetRow(idItem.m_nRow);

	BOOL bSendNotification = dwSelMode != 0 || idItem.IsNull();
	
	if (bSendNotification && !SendTreeNotification(TVN_SELCHANGING, pOldSel, pNewSel, m_nRecentSelectAction))
	{
		return FALSE;
	}

	if (!CBCGPGridCtrl::SetCurSel(idItem, dwSelMode, bRedraw))
	{
		return FALSE;
	}

	if (bSendNotification)
	{
		SendTreeNotification(TVN_SELCHANGED, pOldSel, pNewSel, m_nRecentSelectAction);
	}

	return TRUE;
}
//****************************************************************************
void CBCGPTreeCtrlEx::ExpandRow(CBCGPGridRow* pRow, BOOL bExpand/* = TRUE*/)
{
	if (pRow == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	const BOOL bNeedToNotify = pRow->IsExpanded() != bExpand;
	
	UINT nAction = bExpand ? TVE_EXPAND : TVE_COLLAPSE;

	if (bNeedToNotify && !SendTreeNotification(TVN_ITEMEXPANDING, NULL, pRow, nAction))
	{
		return;
	}

	CBCGPGridCtrl::ExpandRow(pRow, bExpand);

	if (bNeedToNotify)
	{
		SendTreeNotification(TVN_ITEMEXPANDED, NULL, pRow, nAction);
	}
}

#endif // BCGP_EXCLUDE_GRID_CTRL
