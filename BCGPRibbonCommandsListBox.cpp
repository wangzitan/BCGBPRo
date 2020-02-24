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
// BCGPRibbonCommandsListBox.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPRibbonCommandsListBox.h"
#include "BCGPRibbonBar.h"
#include "BCGPRibbonCategory.h"
#include "BCGPRibbonPanel.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#ifndef BCGP_EXCLUDE_RIBBON

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonCommandsListBox

CBCGPRibbonCommandsListBox::CBCGPRibbonCommandsListBox(CBCGPRibbonBar* pRibbonBar,
													   BOOL bIncludeSeparator/* = TRUE*/,
													   BOOL bDrawDefaultIcon/* = FALSE*/,
													   BOOL bCommandsOnly/* = FALSE*/)
{
	ASSERT_VALID (pRibbonBar);

	m_pRibbonBar = pRibbonBar;
	m_nTextOffset = 0;
	m_bDrawDefaultIcon = bDrawDefaultIcon;
	m_bCommandsOnly = bCommandsOnly;
	m_bSimpleDraw = FALSE;

	if (bIncludeSeparator)
	{
		m_pSeparator = new CBCGPRibbonSeparator (TRUE);
	}
	else
	{
		m_pSeparator = NULL;
	}
}
//******************************************************************************
CBCGPRibbonCommandsListBox::~CBCGPRibbonCommandsListBox()
{
	if (m_pSeparator != NULL)
	{
		delete m_pSeparator;
	}
}

BEGIN_MESSAGE_MAP(CBCGPRibbonCommandsListBox, CBCGPListBox)
	//{{AFX_MSG_MAP(CBCGPRibbonCommandsListBox)
	ON_WM_DRAWITEM_REFLECT()
	ON_WM_MEASUREITEM_REFLECT()
	ON_WM_HSCROLL()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonCommandsListBox message handlers

void CBCGPRibbonCommandsListBox::OnDrawItemContent(CDC* pDC, CRect rect, int nIndex)
{
	int cxExtent = GetHorizontalExtent();
	if (cxExtent > rect.Width())
	{
		rect.right = rect.left + cxExtent;
	}
	
	int cxScroll = GetScrollPos(SB_HORZ);
	rect.OffsetRect(-cxScroll, 0);
	
	pDC->SetBkMode (TRANSPARENT);

	CString strText;
	GetText(nIndex, strText);

	if (m_bSimpleDraw)
	{
		const int xMargin = 3;
		rect.DeflateRect(xMargin, 0);
		
		pDC->DrawText (strText, rect, DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
	}
	else
	{
		CBCGPBaseRibbonElement* pCommand = (CBCGPBaseRibbonElement*) GetItemData (nIndex);
		ASSERT_VALID (pCommand);

		pCommand->m_bIsDrawingOnList = TRUE;
		
		BOOL bDrawDefaultIconSaved = pCommand->m_bDrawDefaultIcon;
		pCommand->m_bDrawDefaultIcon = m_bDrawDefaultIcon;

		BOOL bIsSelected = GetCurSel() == nIndex;
		BOOL bIsHighlighted = bIsSelected && (GetFocus() == this);

		pCommand->OnDrawOnList (pDC, strText, m_nTextOffset, rect, bIsSelected, bIsHighlighted);

		pCommand->m_bDrawDefaultIcon = bDrawDefaultIconSaved;
		pCommand->m_bIsDrawingOnList = FALSE;
	}
}
//******************************************************************************
void CBCGPRibbonCommandsListBox::DrawItem(LPDRAWITEMSTRUCT /*lpDIS*/) 
{
}
//******************************************************************************
void CBCGPRibbonCommandsListBox::MeasureItem(LPMEASUREITEMSTRUCT lpMIS)
{
	ASSERT (lpMIS != NULL);

	CClientDC dc (this);
	
	CFont* pOldFont = NULL;
	CFont* pFont = GetFont();
	
	if (pFont->GetSafeHandle() == NULL)
	{
		pOldFont = (CFont*) dc.SelectStockObject (DEFAULT_GUI_FONT);
	}
	else
	{
		pOldFont = dc.SelectObject(pFont);
	}

	ASSERT_VALID (pOldFont);

	TEXTMETRIC tm;
	dc.GetTextMetrics (&tm);

	lpMIS->itemHeight = tm.tmHeight + globalUtils.ScaleByDPI(6);

	dc.SelectObject (pOldFont);
}
//******************************************************************************
void CBCGPRibbonCommandsListBox::FillFromCategory (CBCGPRibbonCategory* pCategory, BOOL bExcludeDefaultPanelButtons/* = FALSE*/)
{
	ASSERT_VALID (this);

	ResetContent ();

	m_nTextOffset = 0;

	if (pCategory == NULL)
	{
		return;
	}

	ASSERT_VALID (pCategory);

	CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arElements;
	pCategory->GetElements (arElements);

	FillFromArray (arElements, TRUE, TRUE, bExcludeDefaultPanelButtons);

	if (m_pSeparator != NULL)
	{
		ASSERT_VALID (m_pSeparator);
		m_pSeparator->AddToListBox (this, FALSE);
	}
}
//******************************************************************************
void CBCGPRibbonCommandsListBox::FillFromArray (
		const CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arElements, BOOL bDeep,
		BOOL bIgnoreSeparators, BOOL bExcludeDefaultPanelButtons)
{
	ASSERT_VALID (this);

	ResetContent ();
	m_nTextOffset = 0;

	for (int i = 0; i < arElements.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pElem = arElements [i];
		ASSERT_VALID (pElem);

		if (bIgnoreSeparators && pElem->IsKindOf (RUNTIME_CLASS (CBCGPRibbonSeparator)))
		{
			continue;
		}

		if (bExcludeDefaultPanelButtons && pElem->IsKindOf (RUNTIME_CLASS (CBCGPRibbonDefaultPanelButton)))
		{
			continue;
		}

		pElem->AddToListBox (this, bDeep);
		
		int nImageWidth = pElem->GetImageSize (CBCGPBaseRibbonElement::RibbonImageSmall).cx;

		m_nTextOffset = max(m_nTextOffset, nImageWidth);
	}

	if (m_nTextOffset == 0)
	{
		m_nTextOffset = globalUtils.ScaleByDPI(16);
	}

	m_nTextOffset += 2;

	if (GetCount () > 0)
	{
		SetCurSel (0);
	}

	RebuildTooltips();
}
//******************************************************************************
void CBCGPRibbonCommandsListBox::FillFromIDs (const CList<UINT,UINT>& lstCommands,
											  BOOL bDeep, BOOL bExcludeDefaultPanelButtons)
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pRibbonBar);

	CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arElements;

	for (POSITION pos = lstCommands.GetHeadPosition (); pos != NULL;)
	{
		UINT uiCmd = lstCommands.GetNext (pos);

		if (uiCmd == 0)
		{
			if (m_pSeparator != NULL)
			{
				arElements.Add (m_pSeparator);
			}

			continue;
		}

		CBCGPBaseRibbonElement* pElem = m_pRibbonBar->FindByID (uiCmd, FALSE);
		if (pElem == NULL)
		{
			continue;
		}

		ASSERT_VALID (pElem);

		if (bExcludeDefaultPanelButtons && pElem->IsKindOf (RUNTIME_CLASS (CBCGPRibbonDefaultPanelButton)))
		{
			continue;
		}

		arElements.Add (pElem);
	}

	FillFromArray (arElements, bDeep, FALSE);
}
//******************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonCommandsListBox::GetSelected () const
{
	ASSERT_VALID (this);

	int nIndex = GetCurSel ();

	if (nIndex < 0)
	{
		return NULL;
	}

	return GetCommand (nIndex);
}
//******************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonCommandsListBox::GetCommand (int nIndex) const
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement* pCommand = 
		(CBCGPBaseRibbonElement*) GetItemData (nIndex);
	ASSERT_VALID (pCommand);

	return pCommand;
}
//******************************************************************************
int CBCGPRibbonCommandsListBox::GetCommandIndex (UINT uiID) const
{
	ASSERT_VALID (this);

	for (int i = 0; i < GetCount (); i++)
	{
		CBCGPBaseRibbonElement* pCommand = (CBCGPBaseRibbonElement*) GetItemData (i);
		ASSERT_VALID (pCommand);

		if (pCommand->GetID () == uiID)
		{
			return i;
		}
	}

	return -1;
}
//******************************************************************************
BOOL CBCGPRibbonCommandsListBox::AddCommand (CBCGPBaseRibbonElement* pCmd, BOOL bSelect, BOOL bDeep)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pCmd);

	int nIndex = GetCommandIndex (pCmd->GetID ());
	if (nIndex >= 0 && pCmd->GetID () != 0)
	{
		return FALSE;
	}

	// Not found, add new:
	if (m_nTextOffset == 0)
	{
		m_nTextOffset = pCmd->GetImageSize (CBCGPBaseRibbonElement::RibbonImageSmall).cx + 2;
	}

	nIndex = pCmd->AddToListBox (this, bDeep);		

	RebuildTooltips();

	if (bSelect)
	{
		SetCurSel (nIndex);
	}

	return TRUE;
}
//******************************************************************************
void CBCGPRibbonCommandsListBox::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	CBCGPListBox::OnHScroll(nSBCode, nPos, pScrollBar);
	RedrawWindow();
}
//******************************************************************************
void CBCGPRibbonCommandsListBox::RebuildTooltips()
{
	for (int nIndex = 0; nIndex < GetCount(); nIndex++)
	{
		CBCGPBaseRibbonElement* pItem = (CBCGPBaseRibbonElement*) GetItemData(nIndex);
		if (pItem == NULL)
		{
			continue;
		}

		ASSERT_VALID(pItem);
		
		CString strTooltip;

		CString strName = pItem->m_strToolTip;
		strName.TrimLeft();
		strName.TrimRight();

		if (strName.IsEmpty())
		{
			strName = pItem->m_strText;
		}

		if (!strName.IsEmpty())
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

		SetItemToolTip(nIndex, strTooltip);
	}
}

#endif	// BCGP_EXCLUDE_RIBBON
