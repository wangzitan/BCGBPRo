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
// BCGPRibbonTreeCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "bcgcbpro.h"
#include "BCGPRibbonTreeCtrl.h"
#include "BCGPRibbonBar.h"
#include "BCGPRibbonCategory.h"
#include "BCGPRibbonPanel.h"
#include "BCGPRibbonCustomizePage.h"
#include "BCGPRibbonCustomizeQATPage.h"
#include "BCGPDrawManager.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#ifndef BCGP_EXCLUDE_GRID_CTRL
#ifndef BCGP_EXCLUDE_RIBBON

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonTreeCtrl

CBCGPRibbonTreeCtrl::CBCGPRibbonTreeCtrl(CBCGPRibbonCustomizationData* pCustomizationData, CBCGPRibbonBar* pRibbonBar,
										 BOOL bShowHiddenCategories) :
	m_pCustomizationData(pCustomizationData),
	m_pRibbonBar(pRibbonBar),
	m_bShowHiddenCategories(bShowHiddenCategories)
{
	m_bListMode = FALSE;
	m_bLockTooltipUpdate = FALSE;
}

CBCGPRibbonTreeCtrl::~CBCGPRibbonTreeCtrl()
{
}

BEGIN_MESSAGE_MAP(CBCGPRibbonTreeCtrl, CBCGPGridCtrl)
	//{{AFX_MSG_MAP(CBCGPRibbonTreeCtrl)
	ON_WM_LBUTTONUP()
	ON_WM_CHAR()
	ON_WM_RBUTTONUP()
	ON_WM_CONTEXTMENU()
	ON_WM_LBUTTONDOWN()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonTreeCtrl message handlers

void CBCGPRibbonTreeCtrl::RebuildItems(int nRibbonTabsDest, LPCTSTR lpszMainTabsCaption)
{
	ASSERT_VALID(m_pRibbonBar);

	int nTab = 0;
	m_bLockTooltipUpdate = TRUE;

	RemoveAll();
	RemoveSortColumn(0);

	m_bListMode = FALSE;

	CArray<CBCGPRibbonCategory*,CBCGPRibbonCategory*> arCategoriesOrdered;

	if (m_pCustomizationData != NULL)
	{
		m_pCustomizationData->PrepareTabsArray(m_pRibbonBar->m_arCategories, arCategoriesOrdered);
	}
	else
	{
		arCategoriesOrdered.Append(m_pRibbonBar->m_arCategories);
	}

	// Exclude non-relevant tabs:
	BOOL bShowMainTabs = nRibbonTabsDest == 0 /* All */ || nRibbonTabsDest == 1 /* Main */;
	BOOL bShowToolsTabs = nRibbonTabsDest == 0 /* All */ || nRibbonTabsDest == 2 /* Tools */;
	
	for (nTab = 0; nTab < (int)arCategoriesOrdered.GetSize(); nTab++)
	{
		CBCGPRibbonCategory* pTab = arCategoriesOrdered[nTab];
		ASSERT_VALID(pTab);
		
		UINT nContextID = GetCategoryContextID(pTab);
		
		if ((nContextID == 0 && !bShowMainTabs) || (nContextID != 0 && !bShowToolsTabs))
		{
			arCategoriesOrdered.RemoveAt(nTab--);
			continue;
		}
		
		if (pTab->IsCustom() && m_pCustomizationData == NULL)
		{
			arCategoriesOrdered.RemoveAt(nTab--);
			continue;
		}
		
		if (pTab->IsPrintPreview() || pTab->m_bToBeDeleted)
		{
			arCategoriesOrdered.RemoveAt(nTab--);
			continue;
		}
		
		if (pTab->m_bIsQATOnly || (!m_bShowHiddenCategories && !pTab->m_bIsVisible && pTab->GetContextID() == 0))
		{
			arCategoriesOrdered.RemoveAt(nTab--);
			continue;
		}
	}
	
	CBCGPRibbonCategory* pPrevTabTab = NULL;

	for (nTab = 0; nTab < (int)arCategoriesOrdered.GetSize(); nTab++)
	{
		CBCGPRibbonCategory* pTab = arCategoriesOrdered[nTab];
		ASSERT_VALID(pTab);

		if (pTab->IsHiddenInAppMode())
		{
			continue;
		}

		UINT nContextID = GetCategoryContextID(pTab);

		if (pPrevTabTab == NULL || (pPrevTabTab != NULL && GetCategoryContextID(pPrevTabTab) != nContextID))
		{
			if (nContextID == 0)
			{
				AddCaptionRow(lpszMainTabsCaption, FALSE);
			}
			else
			{
				CString strName;
				m_pRibbonBar->GetContextName(nContextID, strName);

				AddCaptionRow(strName, FALSE);
			}
		}

		AddCategoty(pTab);
		pPrevTabTab = pTab;
	}

	m_bLockTooltipUpdate = FALSE;
	AdjustLayout();
}
//*******************************************************************************************
void CBCGPRibbonTreeCtrl::RebuildItems(const CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arElements)
{
	ASSERT_VALID(m_pRibbonBar);

	m_bListMode = TRUE;
	m_bLockTooltipUpdate = TRUE;

	RemoveAll();
	RemoveSortColumn(0);

	for (int i = 0; i < arElements.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pElem = arElements [i];
		ASSERT_VALID (pElem);

		CBCGPGridRow* pRow = pElem->AddToTree(this, NULL);
		if (pRow != NULL && pRow->IsGroup())
		{
			pRow->Expand(FALSE);
		}
	}

	m_bLockTooltipUpdate = FALSE;

	SetSortColumn(0);
	AdjustLayout();
}
//*******************************************************************************************
void CBCGPRibbonTreeCtrl::GetVisibleItemsList(CBCGPGridRow* pRootRow, CList<CBCGPGridRow*, CBCGPGridRow*>& lst)
{
	if (pRootRow == NULL)
	{
		for (int nRow = 0; nRow < GetRowCount(); nRow++)
		{
			CBCGPGridRow* pRow = GetRow(nRow);
			ASSERT_VALID(pRow);

			GetVisibleItemsList(pRow, lst);
		}

		return;
	}

	ASSERT_VALID(pRootRow);

	CRect rect = pRootRow->GetRect();
	if (rect.IsRectEmpty())
	{
		return;
	}

	CBCGPBaseRibbonElement* pItem = DYNAMIC_DOWNCAST(CBCGPBaseRibbonElement, (CObject*) pRootRow->GetData());
	if (pItem != NULL)
	{
		lst.AddTail(pRootRow);
		return;
	}

	if (pRootRow->IsExpanded())
	{
		for (int i = 0; i < pRootRow->GetSubItemsCount(); i++)
		{
			CBCGPGridRow* pRow = pRootRow->GetSubItem(i);
			ASSERT_VALID(pRow);
			
			GetVisibleItemsList(pRow, lst);
		}
	}
}
//*******************************************************************************************
void CBCGPRibbonTreeCtrl::RebuildTooltips()
{
	if (m_pToolTip->GetSafeHwnd() == NULL || m_bLockTooltipUpdate)
	{
		return;
	}

	for (int i = m_pToolTip->GetToolCount() - 1 ; i > 0; i--)
	{
		m_pToolTip->DelTool(this, i + 1);
	}

	int nIndex = 2;

	CList<CBCGPGridRow*, CBCGPGridRow*> lstItems;
	GetVisibleItemsList(NULL, lstItems);

	for (POSITION pos = lstItems.GetHeadPosition(); pos != NULL;)
	{
		CBCGPGridRow* pRow = lstItems.GetNext(pos);
		ASSERT_VALID(pRow);

		CBCGPBaseRibbonElement* pItem = DYNAMIC_DOWNCAST(CBCGPBaseRibbonElement, (CObject*) pRow->GetData());
		if (pItem != NULL)
		{
			ASSERT_VALID(pItem);

			CString strTooltip;
			CString strName;

			if (!pItem->m_strToolTip.IsEmpty())
			{
				strName = pItem->m_strToolTip;
			}
			else
			{
				strName = pItem->m_strText;
			}

			if (!strName.IsEmpty())
			{
				if (m_pCustomizationData != NULL)
				{
					strTooltip = strName;

					CString strDescr = pItem->m_strDescription;
					if (!strDescr.IsEmpty() && strDescr != strName)
					{
						strTooltip += _T(" (");
						strTooltip += strDescr;
						strTooltip += _T(")");
					}
				}
				else
				{
					CBCGPRibbonCategory* pCategory = pItem->GetParentCategory();
					if (pCategory != NULL)
					{
						ASSERT_VALID(pCategory);

						UINT nContextID = pCategory->GetContextID();
						if (nContextID != 0)
						{
							CBCGPRibbonBar* pRibbonBar = pCategory->GetParentRibbonBar();
							if (pRibbonBar != NULL)
							{
								ASSERT_VALID(pRibbonBar);

								CString strContextName;
								pRibbonBar->GetContextName(nContextID, strContextName);

								if (!strContextName.IsEmpty())
								{
									strTooltip += strContextName;
									strTooltip += _T(" | ");
								}
							}
						}

						strTooltip += pCategory->GetName();
						strTooltip += _T(" | ");
					}
					
					CBCGPRibbonPanel* pPanel = pItem->GetParentPanel();
					if (pPanel != NULL)
					{
						ASSERT_VALID(pPanel);
						
						strTooltip += pPanel->GetName();
						strTooltip += _T(" | ");
					}

					strTooltip += strName;
				}

				globalUtils.RemoveSingleAmp(strTooltip);

				strTooltip.TrimLeft();
				strTooltip.TrimRight();

				m_pToolTip->AddTool(this, strTooltip, pRow->GetRect(), nIndex++);
			}
		}
	}
}
//*******************************************************************************************
UINT CBCGPRibbonTreeCtrl::GetCategoryContextID(CBCGPRibbonCategory* pCategory) const
{
	ASSERT_VALID(pCategory);
	return (m_pCustomizationData == NULL) ? pCategory->GetContextID() : m_pCustomizationData->GetTabContextID(pCategory);
}
//*******************************************************************************************
void CBCGPRibbonTreeCtrl::AddCategoty(CBCGPRibbonCategory* pCategory)
{
	ASSERT_VALID(pCategory);
	ASSERT_VALID(m_pRibbonBar);

	if (pCategory->IsHiddenInAppMode())
	{
		return;
	}

	BOOL bIsVisible = m_pCustomizationData == NULL || !m_pCustomizationData->IsTabHidden(pCategory);

	const int nColumns = GetColumnCount ();

	CBCGPGridRow* pTab = CreateRow (nColumns);
	ASSERT_VALID (pTab);

	pTab->AllowSubItems ();
	pTab->SetData((DWORD_PTR)pCategory);

	CString strCategoryName;
	if (m_pCustomizationData == NULL || !m_pCustomizationData->GetTabName(pCategory, strCategoryName))
	{
		strCategoryName = pCategory->GetName();
	}

	if (pCategory->IsCustom())
	{
		m_pRibbonBar->PrepareCustomLabel(strCategoryName);
	}
	else if (strCategoryName.IsEmpty())
	{
		strCategoryName = _T(" ");
	}

	CBCGPGridItem* pItem = NULL;
	
	if (m_pCustomizationData == NULL)
	{
		pItem = new CBCGPGridItem((LPCTSTR)strCategoryName);
	}
	else
	{
		CBCGPGridCheckItem* pCheckItem = new CBCGPGridCheckItem(bIsVisible == TRUE);
		pCheckItem->SetLabel(strCategoryName);

		pItem = pCheckItem;
	}

	pTab->ReplaceItem(0, pItem, FALSE);

	AddRow (pTab, FALSE);

	CArray<CBCGPRibbonPanel*,CBCGPRibbonPanel*> arPanelsOrdered;
	if (m_pCustomizationData != NULL)
	{
		m_pCustomizationData->PreparePanelsArray(pCategory->m_arPanels, arPanelsOrdered);
	}
	else
	{
		arPanelsOrdered.Append(pCategory->m_arPanels);
	}

	for (int nPanel = 0; nPanel < arPanelsOrdered.GetSize(); nPanel++)
	{
		CBCGPRibbonPanel* pPanel = arPanelsOrdered[nPanel];
		ASSERT_VALID(pPanel);

		if (pPanel->IsHiddenInAppMode())
		{
			continue;
		}

		if (pPanel->m_bToBeDeleted)
		{
			continue;
		}

		if (m_pCustomizationData != NULL && m_pCustomizationData->IsPanelHidden(pPanel))
		{
			continue;
		}

		if (m_pCustomizationData == NULL && pPanel->IsCustom())
		{
			continue;
		}

		CBCGPGridRow* pPanelItem = CreateRow (nColumns);
		ASSERT_VALID(pPanelItem);

		pPanelItem->AllowSubItems ();
		pPanelItem->SetData((DWORD_PTR)pPanel);

		CString strPanelName;
		if (m_pCustomizationData == NULL || !m_pCustomizationData->GetPanelName(pPanel, strPanelName))
		{
			strPanelName = pPanel->GetName();
		}

		if (pPanel->IsCustom() && pPanel->m_pOriginal == NULL)
		{
			m_pRibbonBar->PrepareCustomLabel(strPanelName);
		}
		else if (strPanelName.IsEmpty())
		{
			strPanelName = _T(" ");
		}

		pPanelItem->GetItem (0)->SetValue((LPCTSTR)strPanelName);

		pTab->AddSubItem (pPanelItem, FALSE);

		CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arElements;

		if (m_pCustomizationData != NULL && pPanel->IsCustom())
		{
			m_pCustomizationData->PrepareElementsArray(pPanel, arElements);
		}
		else
		{
			pPanel->GetElements(arElements);
		}

		for (int i = 0; i < arElements.GetSize(); i++)
		{
			CBCGPBaseRibbonElement* pElem = arElements[i];
			ASSERT_VALID(pElem);

			CBCGPGridRow* pRow = pElem->AddToTree(this, pPanelItem);
			if (pRow != NULL)
			{
				if (m_pCustomizationData != NULL)
				{
					LPCTSTR lpszName = m_pCustomizationData->GetElementName(pElem);
					if (lpszName != NULL)
					{
						pRow->GetItem (0)->SetValue(lpszName);
					}
				}

				if (pRow->IsGroup())
				{
					pRow->Expand(FALSE);
				}
			}
 		}

		if (!pPanel->IsCustom() || pPanel->m_pOriginal != NULL)
		{
			CBCGPGridRow* pRow = pPanel->m_btnLaunch.AddToTree(this, pPanelItem);
			if (pRow != NULL && m_pCustomizationData != NULL)
			{
				LPCTSTR lpszName = m_pCustomizationData->GetElementName(&pPanel->m_btnLaunch);
				if (lpszName != NULL)
				{
					pRow->GetItem (0)->SetValue(lpszName);
				}
			}
		}

		if (pPanelItem->IsGroup())
		{
			pPanelItem->Expand(FALSE);
		}
	}

	if (pTab->IsGroup())
	{
		pTab->Expand(FALSE);
	}
}
//*******************************************************************************************
void CBCGPRibbonTreeCtrl::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CBCGPGridRow* pHit = HitTest(point);
	if (pHit == NULL || DYNAMIC_DOWNCAST(CBCGPGridCaptionRow, pHit) != NULL)
	{
		return;
	}
	
	CBCGPGridCtrl::OnLButtonDown(nFlags, point);
}
//*******************************************************************************************
void CBCGPRibbonTreeCtrl::OnLButtonUp(UINT nFlags, CPoint point) 
{
	CBCGPGridRow* pSel = GetCurSel ();
	if (pSel != NULL)
	{
		CBCGPGridCheckItem* pItem = DYNAMIC_DOWNCAST(CBCGPGridCheckItem, pSel->GetItem (0));
		if (pItem != NULL)
		{
			CRect rectCheck = pItem->GetRect();
			rectCheck.right = rectCheck.left + rectCheck.Height();

			if (rectCheck.PtInRect(point))
			{
				BOOL bOldValue = (bool)pItem->GetValue ();
				pItem->SetValue (!bOldValue);

				pItem->SetItemChanged ();
				return;
			}
		}
	}
	
	CBCGPGridCtrl::OnLButtonUp(nFlags, point);
}
//*******************************************************************************************
void CBCGPRibbonTreeCtrl::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (nChar == VK_SPACE)
	{
		CBCGPGridRow* pSel = GetCurSel ();
		if (pSel != NULL)
		{
			CBCGPGridCheckItem* pItem = DYNAMIC_DOWNCAST(CBCGPGridCheckItem, pSel->GetItem (0));
			if (pItem != NULL)
			{
				pItem->PushChar(nChar);
				return;
			}
		}
	}

	CBCGPGridCtrl::OnChar(nChar, nRepCnt, nFlags);

	if (!m_bListMode)
	{
		return;
	}

	CBCGPGridItemID idStart;
	if (GetCurSel(idStart) != NULL)
	{
		CString str((TCHAR)nChar);
		CBCGPGridItemID idNext;

		BCGP_GRID_FINDREPLACE_PARAM params(str);
		
		params.bWholeCell = FALSE;
		params.bCase = FALSE;
		params.bWholeWord = FALSE;
		params.bStartWith = TRUE;

		for (int iStep = 0; iStep < 2; iStep++)
		{
			if (Search(idNext, idStart, params))
			{
				SetCurSel(idNext);
				EnsureVisible(GetCurSel());
				return;
			}

			idStart.m_nRow = 0;
		}
	}
}
//*******************************************************************************************
BOOL CBCGPRibbonTreeCtrl::OnDrawItem (CDC* pDC, CBCGPGridRow* pItem)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pItem);

	if (pItem->GetRect().IsRectEmpty() || pItem->GetItem(0) == NULL)
	{
		return CBCGPGridCtrl::OnDrawItem(pDC, pItem);
	}

	if (pItem->GetRect().top >= m_rectClip.bottom)
	{
		return FALSE;
	}

	if (pItem->GetRect().bottom < m_rectClip.top)
	{
		return TRUE;
	}

	CObject* lpData = (CObject*)pItem->GetData();
	if (lpData == NULL)
	{
		return CBCGPGridCtrl::OnDrawItem(pDC, pItem);
	}

	CBCGPBaseRibbonElement* pElem = DYNAMIC_DOWNCAST(CBCGPBaseRibbonElement, lpData);
	if (pElem == NULL)
	{
		return CBCGPGridCtrl::OnDrawItem(pDC, pItem);
	}

	if (pItem->HasExpandButton())
	{
		int dx = GetHierarchyOffset(pItem);
		CRect rectExpand = pItem->GetExpandBoxRect(dx);

		CRgn rgnClipExpand;
		CRect rectExpandClip = rectExpand;
		rectExpandClip.bottom = min (rectExpandClip.bottom, m_rectList.bottom);

		rgnClipExpand.CreateRectRgnIndirect (&rectExpandClip);
		pDC->SelectClipRgn (&rgnClipExpand);

		pItem->OnDrawExpandBox (pDC, rectExpand);

		pDC->SelectClipRgn (&m_rgnClip);
	}

	CString strText = (LPCTSTR)(_bstr_t)(pItem->GetItem(0)->GetValue());
	CRect rect = pItem->GetItem(0)->GetRect();

	COLORREF clrText = (COLORREF)-1;

	BOOL bDrawDisabled = (m_pCustomizationData != NULL && (!pElem->IsCustom() || pElem->GetParentPanel() == NULL));
	BOOL bIsFocused = GetFocus() == this;

	if (pItem->IsSelected())
	{
		if (IsVisualManagerStyle())
		{
			CBCGPGridColors colors = GetColorTheme();
			COLORREF clrTextNew;

			if (bIsFocused)
			{
				pDC->FillSolidRect(rect, colors.m_SelColors.m_clrBackground);
				clrTextNew = colors.m_SelColors.m_clrText;
			}
			else
			{
				pDC->FillSolidRect(rect, colors.m_SelColorsInactive.m_clrBackground == (COLORREF)-1 ? globalData.clrBarFace : colors.m_SelColorsInactive.m_clrBackground);
				clrTextNew = colors.m_SelColorsInactive.m_clrText == (COLORREF)-1 ? colors.m_SelColors.m_clrText : colors.m_SelColorsInactive.m_clrText;
			}

			if (bDrawDisabled)
			{
				clrTextNew = CBCGPDrawManager::IsDarkColor(clrTextNew) ? 
					CBCGPDrawManager::ColorMakeLighter(clrTextNew) : CBCGPDrawManager::ColorMakeDarker(clrTextNew);
			}

			clrText = pDC->SetTextColor(clrTextNew);
		}
		else
		{
			if (bIsFocused)
			{
				::FillRect (pDC->GetSafeHdc (), rect, GetSysColorBrush (COLOR_HIGHLIGHT));
				clrText = pDC->SetTextColor (!bDrawDisabled ? GetSysColor (COLOR_HIGHLIGHTTEXT) : globalData.clrBtnLight);
			}
			else
			{
				pDC->FillRect (rect, &globalData.brBtnFace);
				clrText = pDC->SetTextColor (!bDrawDisabled ? globalData.clrBtnText : globalData.clrGrayedText);
			}
		}
	}
	else if (bDrawDisabled)
	{
		COLORREF clrTextDisabled = globalData.clrGrayedText;

		if (IsVisualManagerStyle() && CBCGPVisualManager::GetInstance()->IsDarkTheme())
		{
			clrTextDisabled = CBCGPDrawManager::ColorMakeLighter(clrTextDisabled, .2);
		}

		clrText = pDC->SetTextColor(clrTextDisabled);
	}

	int nSmallImageSize = 16;

	if (globalData.IsRibbonImageScaleEnabled ())
	{
		nSmallImageSize = globalUtils.ScaleByDPI(nSmallImageSize);
		nSmallImageSize = max(nSmallImageSize, pElem->GetImageSize (CBCGPBaseRibbonElement::RibbonImageSmall).cx);
	}

	int nTextOffset = nSmallImageSize + 2;

	pElem->SetForceDrawDisabledOnList(bDrawDisabled);

	BOOL bDrawDefaultIconSaved = pElem->m_bDrawDefaultIcon;
	pElem->m_bDrawDefaultIcon = m_pCustomizationData != NULL && pElem->IsCustomIconAllowed();

	BOOL bIsRibbonImageScale = globalData.IsRibbonImageScaleEnabled ();
	globalData.EnableRibbonImageScale (FALSE);

	int nCustomImageIndexSaved = pElem->m_nCustomImageIndex;
	if (m_pCustomizationData != NULL && pElem->IsCustom())
	{
		pElem->m_nCustomImageIndex = m_pCustomizationData->GetElementImage(pElem);
	}

	pElem->m_bIsDrawingOnList = TRUE;
	pElem->OnDrawOnList(pDC, strText, nTextOffset, rect, pItem->IsSelected(), FALSE);
	
	pElem->SetForceDrawDisabledOnList(FALSE);

	if (clrText != (COLORREF)-1)
	{
		pDC->SetTextColor(clrText);
	}
	
	pElem->m_bDrawDefaultIcon = bDrawDefaultIconSaved;
	pElem->m_nCustomImageIndex = nCustomImageIndexSaved;
	pElem->m_bIsDrawingOnList = FALSE;

	globalData.EnableRibbonImageScale (bIsRibbonImageScale);

	return TRUE;
}
//*******************************************************************************************
void CBCGPRibbonTreeCtrl::ReposItems()
{
	CBCGPGridCtrl::ReposItems();
	RebuildTooltips();
}
//*******************************************************************************************
void CBCGPRibbonTreeCtrl::ShiftItems (int dx, int dy)
{
	CBCGPGridCtrl::ShiftItems(dx, dy);
	RebuildTooltips();
}
//*******************************************************************************************
BOOL CBCGPRibbonTreeCtrl::MoveSelection(BOOL bMoveNext)
{
	CBCGPGridRow* pSel = GetCurSel ();
	if (pSel == NULL)
	{
		return FALSE;
	}

	BOOL bSkipSubItems = FALSE;

	CBCGPGridRow* pParent = pSel->GetParent();
	if (pParent != NULL)
	{
		for (int i = 0; i < pParent->GetSubItemsCount(); i++)
		{
			CBCGPGridRow* pRow = pParent->GetSubItem (i);
			ASSERT_VALID(pRow);

			if (pRow == pSel || bSkipSubItems)
			{
				int nNewSelIndex = bMoveNext ? i + 1 : i - 1;
				if (nNewSelIndex < 0 || nNewSelIndex >= pParent->GetSubItemsCount())
				{
					return FALSE;
				}

				pRow = pParent->GetSubItem (nNewSelIndex);
				if (pRow == NULL)
				{
					return FALSE;
				}

				if (pSel->IsSubItem (pRow))
				{
					bSkipSubItems = TRUE;
					continue;
				}

				SetCurSel(pRow);
				EnsureVisible(GetCurSel());
				return TRUE;
			}
		}
	}
	else
	{
		for (int i = 0; i < GetRowCount(); i++)
		{
			CBCGPGridRow* pRow = GetRow (i);
			if (pRow == NULL)
			{
				continue;
			}

			ASSERT_VALID(pRow);

			if (pRow == pSel || bSkipSubItems)
			{
				int nNewSelIndex = bMoveNext ? i + 1 : i - 1;
				if (nNewSelIndex < 0 || nNewSelIndex >= GetRowCount())
				{
					return FALSE;
				}

				pRow = GetRow (nNewSelIndex);
				if (pRow == NULL)
				{
					return FALSE;
				}

				if (pSel->IsSubItem (pRow))
				{
					bSkipSubItems = TRUE;
					continue;
				}

				SetCurSel(pRow);
				EnsureVisible(GetCurSel());
				return TRUE;
			}
		}
	}

	return FALSE;
}
//***************************************************************************************************************
BOOL CBCGPRibbonTreeCtrl::IsSelFirstInGroup() const
{
	CBCGPGridRow* pSel = GetCurSel ();
	if (pSel == NULL)
	{
		return TRUE;
	}

	CBCGPGridRow* pParent = pSel->GetParent();
	if (pParent != NULL)
	{
		return pParent->GetSubItem(0) == pSel;
	}


	BOOL bIsFirst = TRUE;
	UINT nLastContextID = 0;
	UINT nSelContextID = GetContextID(pSel);

	CObject* lpData = (CObject*)pSel->GetData();
	if (lpData == NULL)
	{
		ASSERT(FALSE);
		return TRUE;
	}
	
	CBCGPRibbonCategory* pCategory = DYNAMIC_DOWNCAST(CBCGPRibbonCategory, lpData);
	if (pCategory == NULL)
	{
		ASSERT(FALSE);
		return TRUE;
	}

	for (int i = 0; i < GetRowCount(); i++)
	{
		CBCGPGridRow* pRow = GetRow (i);
		if (pRow == NULL)
		{
			continue;
		}
		
		ASSERT_VALID(pRow);

		if (DYNAMIC_DOWNCAST(CBCGPGridCaptionRow, pRow) != NULL || pRow->GetParent() != NULL)
		{
			continue;
		}

		if (pRow == pSel)
		{
			if (bIsFirst)
			{
				return TRUE;
			}

			if (nSelContextID == nLastContextID)
			{
				return FALSE;
			}

			return !pCategory->IsCustom();
		}

		nLastContextID = GetContextID(pRow);
		bIsFirst = FALSE;
	}
		
	return TRUE;
}
//***************************************************************************************************************
BOOL CBCGPRibbonTreeCtrl::IsSelLastInGroup() const
{
	CBCGPGridRow* pSel = GetCurSel ();
	if (pSel == NULL)
	{
		return TRUE;
	}
	
	CBCGPGridRow* pParent = pSel->GetParent();
	if (pParent != NULL)
	{
		return pParent->GetSubItem(pParent->GetSubItemsCount() - 1) == pSel;
	}

	BOOL bIsLast = TRUE;
	UINT nLastContextID = 0;
	UINT nSelContextID = GetContextID(pSel);
	
	CObject* lpData = (CObject*)pSel->GetData();
	if (lpData == NULL)
	{
		ASSERT(FALSE);
		return TRUE;
	}
	
	CBCGPRibbonCategory* pCategory = DYNAMIC_DOWNCAST(CBCGPRibbonCategory, lpData);
	if (pCategory == NULL)
	{
		ASSERT(FALSE);
		return TRUE;
	}
	
	for (int i = GetRowCount() - 1; i >= 0; i--)
	{
		CBCGPGridRow* pRow = GetRow (i);
		if (pRow == NULL)
		{
			continue;
		}
		
		ASSERT_VALID(pRow);
		
		if (DYNAMIC_DOWNCAST(CBCGPGridCaptionRow, pRow) != NULL || pRow->GetParent() != NULL)
		{
			continue;
		}
		
		if (pRow == pSel)
		{
			if (bIsLast)
			{
				return TRUE;
			}
			
			if (nSelContextID == nLastContextID)
			{
				return FALSE;
			}
			
			return !pCategory->IsCustom();
		}
		
		nLastContextID = GetContextID(pRow);
		bIsLast = FALSE;
	}
	
	return TRUE;
}
//***************************************************************************************************************
UINT CBCGPRibbonTreeCtrl::GetContextID(CBCGPGridRow* pRow, int nOffset) const
{
	if (pRow == NULL)
	{
		return 0;
	}

	while (pRow->GetParent() != NULL)
	{
		pRow = pRow->GetParent();
	}

	if (nOffset != 0)
	{
		int nCount = 0;
		BOOL bFound = FALSE;

		for (POSITION pos = m_lstItems.Find(pRow); pos != NULL;)
		{
			pRow = nOffset < 0 ? m_lstItems.GetPrev(pos) : m_lstItems.GetNext(pos);
			ASSERT_VALID(pRow);

			if (pRow->GetParent() == NULL && DYNAMIC_DOWNCAST(CBCGPGridCaptionRow, pRow) == NULL)
			{
				nCount++;

				if (nCount == abs(nOffset) + 1)
				{
					bFound = TRUE;
					break;
				}
			}
		}

		if (!bFound)
		{
			return 0;
		}
	}

	CObject* lpData = (CObject*)pRow->GetData();
	if (lpData == NULL)
	{
		ASSERT(FALSE);
		return 0;
	}

	CBCGPRibbonCategory* pCategory = DYNAMIC_DOWNCAST(CBCGPRibbonCategory, lpData);
	if (pCategory == NULL)
	{
		ASSERT(FALSE);
		return 0;
	}

	return GetCategoryContextID(pCategory);
}
//***************************************************************************************************************
void CBCGPRibbonTreeCtrl::OnRButtonUp(UINT nFlags, CPoint point) 
{
	CBCGPGridCtrl::OnRButtonUp(nFlags, point);

	CBCGPRibbonCustomizeRibbonPage* pParent = DYNAMIC_DOWNCAST(CBCGPRibbonCustomizeRibbonPage, GetParent());
	if (pParent != NULL)
	{
		pParent->OnShowTreeContextMenu(this, point);
	}
}
//***************************************************************************************************************
void CBCGPRibbonTreeCtrl::OnContextMenu(CWnd* /*pWnd*/, CPoint point) 
{
	if (point != CPoint(-1, -1))
	{
		return;
	}

	CBCGPRibbonCustomizeRibbonPage* pParent = DYNAMIC_DOWNCAST(CBCGPRibbonCustomizeRibbonPage, GetParent());
	if (pParent != NULL)
	{
		pParent->OnShowTreeContextMenu(this, point);
	}
}
//***************************************************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonTreeCtrl::FindElementInPanel(CBCGPRibbonPanel* pPanel, UINT nID) const
{
	CBCGPGridRow* pPanelRow = FindRowByData((DWORD_PTR)pPanel);
	if (pPanelRow == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}

	for (int i = 0; i < pPanelRow->GetSubItemsCount(); i++)
	{
		CBCGPGridRow* pRow = pPanelRow->GetSubItem(i);
		ASSERT_VALID(pRow);

		CObject* lpData = (CObject*)pRow->GetData();
		if (lpData == NULL)
		{
			ASSERT(FALSE);
			continue;
		}
		
		CBCGPBaseRibbonElement* pElem = DYNAMIC_DOWNCAST(CBCGPBaseRibbonElement, lpData);
		if (pElem != NULL)
		{
			CBCGPBaseRibbonElement* pCurrElem = pElem->FindByID(nID);
			if (pCurrElem != NULL)
			{
				return pCurrElem;
			}
		}
	}

	return FALSE;
}

#endif
#endif
