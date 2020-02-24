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
// BCGPRibbonButton.cpp: implementation of the CBCGPRibbonButton class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BCGPRibbonCategory.h"
#include "BCGPRibbonButton.h"
#include "BCGPRibbonBar.h"
#include "bcgglobals.h"
#include "BCGPVisualManager.h"
#include "MenuImages.h"
#include "BCGPRibbonButtonsGroup.h"
#include "BCGPRibbonPanelMenu.h"
#include "BCGPToolbarMenuButton.h"
#include "BCGPMDIFrameWnd.h"
#include "BCGPMDIChildWnd.h"
#include "BCGPRibbonCheckBox.h"
#include "BCGPRibbonFloaty.h"
#include "BCGPGridCtrl.h"
#include "BCGPPopupDlg.h"
#include "BCGPDialog.h"
#include "BCGPDrawManager.h"

#include "winuser.h"
#include "oleacc.h"

#ifndef BCGP_EXCLUDE_RIBBON

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

BOOL CBCGPRibbonButton::m_bUseMenuHandle = FALSE;

const int nLargeButtonMarginX = 5;
const int nLargeButtonMarginY = 1;

const int nSmallButtonMarginX = 3;
const int nSmallButtonMarginY = 3;
const int nBackstageViewMarginY = 5;

const int nDefaultPaneButtonMargin = 2;

const int nMaxMDIChildWindows = 9;

IMPLEMENT_DYNCREATE(CBCGPRibbonButton, CBCGPBaseRibbonElement)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPRibbonButton::CBCGPRibbonButton()
{
	CommonInit ();
}
//********************************************************************************
CBCGPRibbonButton::CBCGPRibbonButton (UINT nID, LPCTSTR lpszText, 
									  int nSmallImageIndex, 
									  int nLargeImageIndex,
									  BOOL bAlwaysShowDescription)
{
	CommonInit ();

	m_nID = nID;
	SetText (lpszText);

	m_nSmallImageIndex = nSmallImageIndex;
	m_nLargeImageIndex = nLargeImageIndex;

	m_bAlwaysShowDescription = bAlwaysShowDescription;
}
//********************************************************************************
CBCGPRibbonButton::CBCGPRibbonButton (
		UINT	nID, 
		LPCTSTR lpszText, 
		HICON	hIcon,
		BOOL	bAlwaysShowDescription,
		HICON	hIconSmall,
		BOOL	bAutoDestroyIcon,
		BOOL	bAlphaBlendIcon)
{
	CommonInit ();

	m_nID = nID;
	SetText (lpszText);
	m_hIcon = hIcon;
	m_hIconSmall = hIconSmall;
	m_bAlwaysShowDescription = bAlwaysShowDescription;
	m_bAutoDestroyIcon = bAutoDestroyIcon;
	m_bAlphaBlendIcon = bAlphaBlendIcon;
}
//********************************************************************************
CBCGPRibbonButton::CBCGPRibbonButton (
		UINT	nID, 
		HICON	hIconSmall,
		LPCTSTR lpszText, 
		BOOL	bAutoDestroyIcon,
		BOOL	bAlphaBlendIcon)
{
	CommonInit ();

	m_bSmallIconMode = TRUE;

	m_nID = nID;
	SetText (lpszText);
	m_hIcon = NULL;
	m_hIconSmall = hIconSmall;
	m_bAutoDestroyIcon = bAutoDestroyIcon;
	m_bAlphaBlendIcon = bAlphaBlendIcon;
}
//********************************************************************************
void CBCGPRibbonButton::CommonInit ()
{
	m_hMenu = NULL;
	m_bRightAlignMenu = FALSE;
	m_bIsDefaultCommand = TRUE;
	m_nMenuArrowMargin = 2;
	m_bAutodestroyMenu = FALSE;

	m_nSmallImageIndex = -1;
	m_nLargeImageIndex = -1;

	m_sizeTextRight = CSize (0, 0);
	m_sizeTextBottom = CSize (0, 0);

	m_szMargin = IsBackstageViewMode() ? CSize(GetBackstageViewMarginX(), nBackstageViewMarginY) : CSize (nSmallButtonMarginX, nSmallButtonMarginY);

	m_rectMenu.SetRectEmpty ();
	m_rectCommand.SetRectEmpty ();
	m_bMenuOnBottom = FALSE;
	m_bIsLargeImage = FALSE;

	m_bIsMenuHighlighted = FALSE;
	m_bIsCommandHighlighted = FALSE;

	m_hIconSimplified = NULL;
	m_hIcon = NULL;
	m_hIconSmall = NULL;
	m_bForceDrawBorder = FALSE;

	m_bToBeClosed = FALSE;
	m_bAlwaysShowDescription = FALSE;

	m_bCreatedFromMenu = FALSE;
	m_bIsWindowsMenu = FALSE;
	m_nWindowsMenuItems = 0;
	m_nWrapIndex = -1;

	m_bAutoDestroyIcon = FALSE;
	m_bAlphaBlendIcon = FALSE;
	m_bSmallIconMode = FALSE;

	m_QATType = BCGPRibbonButton_Show_Default;

	m_pRTIPopupDlg = NULL;
	m_lpszPopupDlgTemplateName = NULL;
	m_pPopupDlg = NULL;
	m_bIsResizablePopup = FALSE;
	m_bIsRightAlignedPopup = FALSE;

	m_bDontCloseParentPopupOnClick = FALSE;
	m_bForceSingleLineText = FALSE;
}
//********************************************************************************
CBCGPRibbonButton::~CBCGPRibbonButton()
{
	RemoveAllSubItems ();

	if (m_bAutodestroyMenu && m_hMenu != NULL)
	{
		::DestroyMenu (m_hMenu);
	}

	if (m_bAutoDestroyIcon && m_hIcon != NULL)
	{
		::DestroyIcon (m_hIcon);
	}

	if (m_bAutoDestroyIcon && m_hIconSmall != NULL)
	{
		::DestroyIcon (m_hIconSmall);
	}

	if (m_hIconSimplified != NULL)
	{
		::DestroyIcon(m_hIconSimplified);
	}
}
//********************************************************************************
int CBCGPRibbonButton::GetBackstageViewMarginX() const
{
	return IsDisplayIconInBackstageViewMode() ? 15 : 30;
}
//********************************************************************************
void CBCGPRibbonButton::SetText (LPCTSTR lpszText)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::SetText (lpszText);

	m_sizeTextRight = CSize (0, 0);
	m_sizeTextBottom = CSize (0, 0);

	m_arWordIndexes.RemoveAll ();

	int nStart = 0;

	if (!globalData.m_bAutoTrimRibbonLabels)
	{
		nStart = -1;

		for (int i = 0; i < m_strText.GetLength(); i++)
		{
			if (m_strText[i] != _T(' '))
			{
				nStart = i;
				break;
			}
		}

		if (nStart < 0)
		{
			return;
		}
	}

	for (int nOffset = nStart;;)
	{
		int nIndex = m_strText.Find (_T(' '), nOffset);
		if (nIndex >= 0)
		{
			ASSERT (nIndex != 0);
			m_arWordIndexes.Add (nIndex);
			nOffset = nIndex + 1;
		}
		else
		{
			break;
		}
	}
}
//********************************************************************************
void CBCGPRibbonButton::SetDescription (LPCTSTR lpszText)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::SetDescription (lpszText);

	if (m_bAlwaysShowDescription)
	{
		m_sizeTextRight = CSize (0, 0);
		m_sizeTextBottom = CSize (0, 0);
	}
}
//********************************************************************************
void CBCGPRibbonButton::SetMenu (HMENU hMenu, BOOL bIsDefaultCommand, BOOL bRightAlign)
{
	ASSERT_VALID (this);
	
	m_bIsWindowsMenu = FALSE;
	m_nWindowsMenuItems = 0;

	if (m_bAutodestroyMenu && m_hMenu != NULL)
	{
		::DestroyMenu (m_hMenu);
	}

	m_bAutodestroyMenu = FALSE;

	if (m_bUseMenuHandle)
	{
		m_hMenu = hMenu;
	}
	else if (hMenu != NULL)
	{
		CMenu* pMenu = CMenu::FromHandle (hMenu);

		for (int i = 0; i < (int) pMenu->GetMenuItemCount (); i++)
		{
			UINT uiID = pMenu->GetMenuItemID (i);

			switch (uiID)
			{
			case 0:
				{
					CBCGPRibbonSeparator* pSeparator = new CBCGPRibbonSeparator (TRUE);
					pSeparator->SetDefaultMenuLook ();

					AddSubItem (pSeparator);
					break;
				}

			default:
				{
					CString str;
					pMenu->GetMenuString (i, str, MF_BYPOSITION);

					//-----------------------------------
					// Remove standard aceleration label:
					//-----------------------------------
					int iTabOffset = str.Find (_T('\t'));
					if (iTabOffset >= 0)
					{
						str = str.Left(iTabOffset);
					}

					CBCGPRibbonButton* pItem = new CBCGPRibbonButton (uiID, str);
					pItem->SetDefaultMenuLook ();
					pItem->m_pRibbonBar = m_pRibbonBar;
					pItem->m_bDontCloseParentPopupOnClick = m_bDontCloseParentPopupOnClick;

					if (uiID == -1)
					{
						pItem->SetMenu (pMenu->GetSubMenu (i)->GetSafeHmenu (), FALSE);
					}

					AddSubItem (pItem);

					if (uiID >= AFX_IDM_WINDOW_FIRST && uiID <= AFX_IDM_WINDOW_LAST)
					{
						m_bIsWindowsMenu = TRUE;
					}
				}
			}
		}
	}

	m_bIsDefaultCommand = bIsDefaultCommand;

	if (m_nID == 0 || m_nID == (UINT)-1 || hMenu == NULL)
	{
		m_bIsDefaultCommand = FALSE;
	}

	if (!m_bIsDefaultCommand)
	{
		m_rectCommand.SetRectEmpty();
		m_rectMenu.SetRectEmpty();
	}

	m_bRightAlignMenu = bRightAlign;

	m_sizeTextRight = CSize (0, 0);
	m_sizeTextBottom = CSize (0, 0);

	m_bCreatedFromMenu = TRUE;
}
//********************************************************************************
void CBCGPRibbonButton::SetMenu (UINT uiMenuResID, BOOL bIsDefaultCommand, BOOL bRightAlign)
{
	ASSERT_VALID (this);

	SetMenu (
		uiMenuResID == 0 ? (HMENU)NULL : ::LoadMenu (::AfxGetResourceHandle (), MAKEINTRESOURCE (uiMenuResID)),
		bIsDefaultCommand, bRightAlign);

	m_bAutodestroyMenu = TRUE;
}
//********************************************************************************
void CBCGPRibbonButton::OnCalcTextSize (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	if (m_strText.IsEmpty () || IsMainRibbonButton ())
	{
		m_sizeTextRight = CSize (0, 0);
		m_sizeTextBottom = CSize (0, 0);
		return;
	}

	if (m_sizeTextRight != CSize (0, 0) && m_sizeTextBottom != CSize (0, 0))
	{
		// Already calculated
		return;
	}

	// Text placed on right will be always single line:

	CString strText = m_strText;

	if (m_bIsTabElement && CBCGPVisualManager::GetInstance()->IsTopLevelMenuItemUpperCase())
	{
		strText.MakeUpper();
	}

	globalUtils.RemoveSingleAmp(strText);

	CFont* pOldFont = NULL;
	if (GetBackstageAttachedView() != NULL)
	{
		CFont* pFont = CBCGPVisualManager::GetInstance()->GetBackstageViewEntryFont();
		if (pFont != NULL)
		{
			pOldFont = pDC->SelectObject(pFont);
		}
	}

	if (m_bAlwaysShowDescription && !m_strDescription.IsEmpty())
	{
		CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar();

		CFont* pOldFontBold = pDC->SelectObject((pTopLevelRibbon->GetSafeHwnd() != NULL) ? pTopLevelRibbon->GetBoldFont() : &globalData.fontBold);
		ASSERT (pOldFontBold != NULL);

		m_sizeTextRight = pDC->GetTextExtent (strText);

		pDC->SelectObject (pOldFontBold);

		// Desciption will be draw below the text (in case of text on right only)
		int nTextHeight = 0;
		int nTextWidth = 0;

		CString strDescr = m_strDescription;

		for (int dx = m_sizeTextRight.cx; dx < m_sizeTextRight.cx * 10; dx += 10)
		{
			CRect rectText (0, 0, dx, 10000);

			nTextHeight = pDC->DrawText (strDescr, rectText, DT_WORDBREAK | DT_CALCRECT);

			nTextWidth = rectText.Width ();
			
			if (nTextHeight <= 2 * m_sizeTextRight.cy)
			{
				break;
			}
		}
	
		m_sizeTextRight.cx = max (m_sizeTextRight.cx, nTextWidth);
		m_sizeTextRight.cy += min (2 * m_sizeTextRight.cy, nTextHeight) + 2 * m_szMargin.cy;
	}
	else
	{
		// Text placed on right will be always single line:
		m_sizeTextRight = pDC->GetTextExtent (strText);
	}

	CSize sizeImageLarge = GetImageSize (RibbonImageLarge);

	if (sizeImageLarge == CSize (0, 0))
	{
		m_sizeTextBottom = CSize (0, 0);
	}
	else
	{
		// Text placed on bottom will occupy large image size and 1-2 text rows:
		m_sizeTextBottom = DrawBottomText (pDC, TRUE /*bCalcOnly*/);
	}

	if (pOldFont != NULL)
	{
		pDC->SelectObject(pOldFont);
	}
}
//********************************************************************************
void CBCGPRibbonButton::OnDraw (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	if (m_rect.IsRectEmpty ())
	{
		return;
	}

	if (IsDefaultMenuLook () && !IsQATMode () && !m_bIsLargeImage && (!IsSearchResultMode () || IsCommandSearchMenu()))
	{
		CBCGPToolbarMenuButton dummy;

		dummy.m_strText = m_strText;
		dummy.m_nID = m_nID;

		if (GetImageSize(RibbonImageSmall) != CSize(0, 0))
		{
			dummy.m_bRibbonImage = TRUE;
		}

		dummy.m_bMenuMode = TRUE;
		dummy.m_pWndParent = GetParentWnd ();
		dummy.m_bIsRadio = m_bIsRadio;

		if (IsChecked ())
		{
			dummy.m_nStyle |= TBBS_CHECKED;
		}

		if (HasMenu ())
		{
			dummy.m_bDrawDownArrow = TRUE;
		}

		BOOL bIsHighlighted = m_bIsHighlighted;

		if (IsDisabled ())
		{
			dummy.m_nStyle |= TBBS_DISABLED;

			bIsHighlighted = IsFocused ();
		}

		if (m_bIsTemporaryPopupItem)
		{
			dummy.m_pRibbonElement = this;
		}

		dummy.OnDraw (pDC, m_rect, NULL, TRUE, FALSE, bIsHighlighted || m_bIsFocused);
		return;
	}

	BOOL bIsDisabled = m_bIsDisabled;
	BOOL bIsDroppedDown = m_bIsDroppedDown;
	BOOL bIsHighlighted = m_bIsHighlighted;
	BOOL bMenuHighlighted = m_bIsMenuHighlighted;
	BOOL bCommandHighlighted = m_bIsCommandHighlighted;

	const int cxDropDown = GetDropDownImageWidth ();

	if (m_bIsDisabled && (HasMenu () || HasPin()))
	{
		if (m_bIsDefaultCommand || 
			(!m_bIsDefaultCommand && !(m_nID == 0 || m_nID == (UINT)-1)))
		{
			m_bIsHighlighted = FALSE;
		}
		else if (!m_bTemporaryDisabled)
		{
			m_bIsDisabled = FALSE;
		}
	}

	if (m_bToBeClosed)
	{
		m_bIsDroppedDown = FALSE;
	}

	if (m_bIsFocused)
	{
		m_bIsHighlighted = TRUE;
		m_bIsMenuHighlighted = TRUE;
		m_bIsCommandHighlighted = TRUE;
	}

	CRect rectMenuArrow;
	rectMenuArrow.SetRectEmpty ();

	if (HasMenu () || HasPin())
	{
		rectMenuArrow = m_rect;

		rectMenuArrow.left = rectMenuArrow.right - cxDropDown - m_nMenuArrowMargin;
		if (m_sizeTextRight.cx == 0 && !m_bQuickAccessMode)
		{
			rectMenuArrow.left -= 2;
		}

		rectMenuArrow.bottom -= m_nMenuArrowMargin;

		if (m_bIsDefaultCommand)
		{
			m_rectMenu = m_rect;

			m_rectMenu.left = m_rectMenu.right - cxDropDown - m_nMenuArrowMargin - 1;
			
			m_rectCommand = m_rect;
			m_rectCommand.right = m_rectMenu.left;

			m_bMenuOnBottom = FALSE;
		}
	}

	CSize sizeImageLarge = GetImageSize (RibbonImageLarge);
	CSize sizeImageSmall = GetImageSize (RibbonImageSmall);

	CRect rectText = m_rect;
	BOOL bDrawText = !IsMainRibbonButton () && !m_bQuickAccessMode && !m_bFloatyMode;

	if (m_bQuickAccessMode || m_bFloatyMode)
	{
		bDrawText = FALSE;
	}
	else if (m_bCompactMode)
	{
		bDrawText = FALSE;
	}
	else if (sizeImageLarge != CSize (0, 0) && !m_bMenuOnBottom && m_bIsLargeImage)
	{
		if (!m_rectMenu.IsRectEmpty ())
		{
			m_rectMenu.left -= cxDropDown;
			m_rectCommand.right = m_rectMenu.left;
		}

		rectMenuArrow.OffsetRect (-cxDropDown / 2, 0);
	}

	const RibbonImageType imageType = 
		m_bIsLargeImage ? RibbonImageLarge : RibbonImageSmall;

	CSize sizeImage = GetImageSize (imageType);
	BOOL bDrawDefaultImage = FALSE;

	if ((m_bQuickAccessMode || m_bFloatyMode || m_bSearchResultMode) && sizeImage == CSize (0, 0))
	{
		// Use default image:
		sizeImage = globalUtils.ScaleByDPI(CSize (16, 16));
		bDrawDefaultImage = TRUE;
	}
	
	CRect rectImage = m_rect;
	rectImage.DeflateRect (m_szMargin);

	if (IsMainRibbonButton ())
	{
		if (globalData.GetRibbonImageScale () != 1.)
		{
			sizeImage.cx = (int) (.8 * globalData.GetRibbonImageScale () * sizeImage.cx);
			sizeImage.cy = (int) (.8 * globalData.GetRibbonImageScale () * sizeImage.cy);
		}

		rectImage.left += (rectImage.Width () - sizeImage.cx) / 2;
		rectImage.top  += (rectImage.Height () - sizeImage.cy) / 2;

		rectImage.OffsetRect (CBCGPVisualManager::GetInstance ()->GetRibbonMainImageOffset ());
	}
	else if (m_bIsLargeImage && !m_bTextAlwaysOnRight)
	{
		rectImage.left = rectImage.CenterPoint ().x - sizeImage.cx / 2;
		rectImage.top += m_szMargin.cy + 1;

		if (!bDrawText)
		{
			rectImage.top = rectImage.CenterPoint ().y - sizeImage.cy / 2;
		}
	}
	else
	{
		rectImage.left += m_sizePadding.cx / 2;
		rectImage.top = rectImage.CenterPoint ().y - sizeImage.cy / 2;
	}

	rectImage.right = rectImage.left + sizeImage.cx;
	rectImage.bottom = rectImage.top + sizeImage.cy;

	if (m_bIsLargeImage && !m_bTextAlwaysOnRight && (HasMenu () || HasPin()) && m_bIsDefaultCommand)
	{
		m_rectMenu = m_rect;
		m_rectMenu.top = rectImage.bottom + 3;
		
		m_rectCommand = m_rect;
		m_rectCommand.bottom = m_rectMenu.top;

		m_bMenuOnBottom = TRUE;
	}

	COLORREF clrText = (COLORREF)-1;

	if (!IsMainRibbonButton ())
	{
		clrText = OnFillBackground (pDC);
	}

	if (IsMenuMode () && IsChecked () && sizeImage != CSize (0, 0))
	{
		CBCGPVisualManager::GetInstance ()->OnDrawRibbonMenuCheckFrame (pDC,
			this, rectImage);

		if (bDrawDefaultImage && IsSearchResultMode ())
		{
			CBCGPToolbarMenuButton dummy;
			CBCGPVisualManager::GetInstance ()->OnDrawMenuCheck (pDC, &dummy, 
				rectImage, m_bIsHighlighted, m_bIsRadio);
		}
	}

	//------------
	// Draw image:
	//------------
	if (bDrawDefaultImage)
	{
		if (!IsSearchResultMode ())
		{
			CBCGPVisualManager::GetInstance ()->OnDrawDefaultRibbonImage (
				pDC, rectImage, m_bIsDisabled, m_bIsPressed, m_bIsHighlighted);
		}
	}
	else
	{
		BOOL bIsRibbonImageScale = globalData.IsRibbonImageScaleEnabled ();

		if (IsMenuMode () && !m_bIsLargeImage)
		{
			if (m_pParentMenu == NULL || !m_pParentMenu->IsMainPanel ())
			{
				globalData.EnableRibbonImageScale (FALSE);
			}
		}

		DrawImage (pDC, imageType, rectImage);
		globalData.EnableRibbonImageScale (bIsRibbonImageScale);
	}

	//-----------
	// Draw text:
	//-----------
	if (bDrawText)
	{
		CFont* pOldFont = NULL;

		CRect rectText1 = m_rect;
		UINT uiDTFlags = 0;

		COLORREF clrTextOld = (COLORREF)-1;

		if (bIsDisabled && 
			(m_bIsDefaultCommand ||
			(!m_bIsDefaultCommand && !(m_nID == 0 || m_nID == (UINT)-1))))
		{
			if (m_bQuickAccessMode)
			{
				clrText = CBCGPVisualManager::GetInstance ()->GetRibbonQATTextColor (TRUE);
			}
			else
			{
				clrTextOld = pDC->SetTextColor (
					clrText == (COLORREF)-1 ? 
						CBCGPVisualManager::GetInstance ()->GetRibbonBarTextColor (TRUE) : clrText);
			}
		}
		else if (clrText != (COLORREF)-1)
		{
			clrTextOld = pDC->SetTextColor (clrText);
		}

		if (m_bIsLargeImage && !m_bTextAlwaysOnRight)
		{
			DrawBottomText (pDC, FALSE);
			rectMenuArrow.SetRectEmpty ();
		}
		else
		{
			rectText1.left = rectImage.right;

			if (m_nImageOffset > 0)
			{
				rectText1.left = m_rect.left + m_nImageOffset + 3 * m_szMargin.cx;
			}
			else if (sizeImage.cx != 0)
			{
				rectText1.left += GetTextOffset ();
			}

			uiDTFlags = DT_SINGLELINE | DT_END_ELLIPSIS;
			
			if (!m_bAlwaysShowDescription || m_strDescription.IsEmpty ())
			{
				if (GetBackstageAttachedView() != NULL)
				{
					CFont* pFont = CBCGPVisualManager::GetInstance()->GetBackstageViewEntryFont();
					if (pFont != NULL)
					{
						pOldFont = pDC->SelectObject(pFont);
						ASSERT (pOldFont != NULL);
					}
				}

				uiDTFlags |= DT_VCENTER;
			}
			else
			{
				CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar();

				pOldFont = pDC->SelectObject((pTopLevelRibbon->GetSafeHwnd() != NULL) ? pTopLevelRibbon->GetBoldFont() : &globalData.fontBold);
				ASSERT (pOldFont != NULL);

				rectText1.top += max (0, (m_rect.Height () - m_sizeTextRight.cy) / 2);
			}

			if (m_bIsTabElement)
			{
				rectText1.OffsetRect (0, 2);
			}

			if (!rectMenuArrow.IsRectEmpty())
			{
				rectText1.right = min(rectText1.right, rectMenuArrow.left);
			}

			CString strText = m_strText;

			if (IsBackstageViewMode())
			{
				globalUtils.RemoveSingleAmp(strText);
			}

			int nTextHeight = DoDrawText (pDC, strText, rectText1, uiDTFlags,
				(bIsDisabled && CBCGPToolBarImages::m_bIsDrawOnGlass)
					? (m_bQuickAccessMode
						? CBCGPVisualManager::GetInstance ()->GetRibbonQATTextColor (TRUE)
						: CBCGPVisualManager::GetInstance ()->GetToolbarDisabledTextColor ())
					: -1);

			if (pOldFont != NULL)
			{
				pDC->SelectObject (pOldFont);
				pOldFont = NULL;
			}

			if (m_bAlwaysShowDescription && !m_strDescription.IsEmpty ())
			{
				rectText1.top += nTextHeight + m_szMargin.cy;
				rectText1.right = m_rect.right - m_szMargin.cx;

				pDC->DrawText (m_strDescription, rectText1, DT_WORDBREAK | DT_END_ELLIPSIS);
			}

			if (nTextHeight == m_sizeTextRight.cy &&
				m_bIsLargeImage && (HasMenu () || HasPin()))
			{
				rectMenuArrow = m_rect;
				rectMenuArrow.DeflateRect (m_nMenuArrowMargin, m_nMenuArrowMargin * 2);
				rectMenuArrow.right -= 2;

				int cyMenu = CBCGPMenuImages::Size ().cy;

				rectMenuArrow.top = rectMenuArrow.bottom - cyMenu;
				rectMenuArrow.bottom = rectMenuArrow.top + CBCGPMenuImages::Size ().cy;
			}
		}

		if (clrTextOld != (COLORREF)-1)
		{
			pDC->SetTextColor (clrTextOld);
		}
	}

	if (!IsMainRibbonButton ())
	{
		if (!rectMenuArrow.IsRectEmpty ())
		{
			OnDrawMenuArrow(pDC, rectMenuArrow);
		}

		OnDrawBorder (pDC);
	}

	m_bIsDisabled = bIsDisabled;
	m_bIsDroppedDown = bIsDroppedDown;
	m_bIsHighlighted = bIsHighlighted;
	m_bIsMenuHighlighted = bMenuHighlighted;
	m_bIsCommandHighlighted = bCommandHighlighted;
}
//********************************************************************************
void CBCGPRibbonButton::OnDrawMenuArrow(CDC* pDC, const CRect& rectMenuArrow)
{
	CBCGPMenuImages::IMAGES_IDS id = CBCGPMenuImages::IdArowDown;

	if (IsMenuMode ())
	{
		BOOL bIsRTL = FALSE;

		CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar ();
		if (pTopLevelRibbon->GetSafeHwnd () != NULL)
		{
			bIsRTL = (pTopLevelRibbon->GetExStyle () & WS_EX_LAYOUTRTL);
		}

		id = bIsRTL ? CBCGPMenuImages::IdArowLeftLarge : CBCGPMenuImages::IdArowRightLarge;
	}

	CBCGPVisualManager::GetInstance()->OnDrawRibbonMenuArrow(pDC, this, (UINT)id, rectMenuArrow);
}
//********************************************************************************
void CBCGPRibbonButton::OnDrawOnList (CDC* pDC, CString strText,
									  int nTextOffset, CRect rect,
									  BOOL bIsSelected,
									  BOOL bHighlighted)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	m_bIsDontDrawSimplifiedImage = TRUE;

	BOOL bIsDisabled = m_bIsDisabled;
	m_bIsDisabled = m_bForceDrawDisabledOnList;

	CRect rectImage = rect;
	rectImage.right = rect.left + nTextOffset;

	switch (m_QATType)
	{
	case BCGPRibbonButton_Show_As_CheckBox:
		CBCGPVisualManager::GetInstance ()->OnDrawRibbonCheckBoxOnList(
			pDC, this, rectImage, bIsSelected, bHighlighted);
		break;

	case BCGPRibbonButton_Show_As_RadioButton:
		CBCGPVisualManager::GetInstance ()->OnDrawRibbonRadioButtonOnList(
			pDC, this, rectImage, bIsSelected, bHighlighted);
		break;

	default:
		{
			CSize sizeImageSmall = GetImageSize (RibbonImageSmall);
			if (sizeImageSmall != CSize (0, 0))
			{
				rectImage.DeflateRect (1, 0);
				rectImage.top += max (0, (rectImage.Height () - sizeImageSmall.cy) / 2);
				rectImage.bottom = rectImage.top + sizeImageSmall.cy;

				DrawImage (pDC, RibbonImageSmall, rectImage);
			}
			else if (m_bDrawDefaultIcon)
			{
				CBCGPVisualManager::GetInstance ()->OnDrawDefaultRibbonImage (
					pDC, rectImage, m_bIsDisabled);
			}
		}
	}

	CRect rectText = rect;

	if (HasMenu () || HasPin())
	{
		CRect rectMenuArrow = rect;
		rectMenuArrow.left = rectMenuArrow.right - rectMenuArrow.Height ();

		BOOL bIsDarkMenu = TRUE;

		if (bHighlighted)
		{
			COLORREF clrText = GetSysColor (COLOR_HIGHLIGHTTEXT);
			
			if (GetRValue (clrText) > 100 &&
				GetGValue (clrText) > 100 &&
				GetBValue (clrText) > 100)
			{
				bIsDarkMenu = FALSE;
			}
		}

		if (HasPin())
		{
			CBCGPVisualManager::GetInstance()->OnDrawPin(pDC, rectMenuArrow, FALSE, bIsDarkMenu, bHighlighted, FALSE, FALSE);
		}
		else
		{
			CBCGPMenuImages::IMAGES_IDS id = 
				globalData.GetRibbonImageScale () > 1. ? 
				CBCGPMenuImages::IdArowRightLarge : CBCGPMenuImages::IdArowRight;

			if (globalData.m_bIsRTL)
			{
				id = globalData.GetRibbonImageScale () > 1. ? 
					CBCGPMenuImages::IdArowLeftLarge : CBCGPMenuImages::IdArowLeft;
			}

			CBCGPVisualManager::GetInstance()->OnDrawRibbonMenuArrow(pDC, this, (UINT)id, rectMenuArrow);
		}

		rectText.right = rectMenuArrow.left;
	}

	rectText.left += nTextOffset;

	const int xMargin = 3;
	rectText.DeflateRect (xMargin, 0);

	pDC->DrawText (strText, rectText, DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);

	m_bIsDisabled = bIsDisabled;
	m_bIsDontDrawSimplifiedImage = FALSE;
}
//********************************************************************************
CSize CBCGPRibbonButton::GetRegularSize (CDC* pDC)
{
	ASSERT_VALID (this);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		m_arSubItems [i]->SetParentCategory (m_pParent);
	}

	if (m_bQuickAccessMode || m_bFloatyMode)
	{
		return GetCompactSize (pDC);
	}

	if (!HasLargeMode () || m_bSearchResultMode)
	{
		return GetIntermediateSize (pDC);
	}

	CSize sizeImageLarge = GetImageSize (RibbonImageLarge);
	CSize sizeImageSmall = GetImageSize (RibbonImageSmall);

	if (IsMainRibbonButton ())
	{
		return sizeImageLarge + m_sizePadding;
	}

	const int cxExtra = GetGroupButtonExtraWidth ();

	if (sizeImageLarge == CSize (0, 0) || m_bTextAlwaysOnRight)
	{
		if (m_bTextAlwaysOnRight && sizeImageLarge != CSize (0, 0))
		{
			sizeImageSmall = CSize (sizeImageLarge.cx + 2, sizeImageLarge.cy + 2);
			m_szMargin.cy = 5;
		}

		int cx = sizeImageSmall.cx + 2 * m_szMargin.cx;
		
		if (m_sizeTextRight.cx > 0)
		{
			cx += m_szMargin.cx + m_sizeTextRight.cx;

			if (sizeImageLarge != CSize (0, 0) && m_bTextAlwaysOnRight)
			{
				cx += m_szMargin.cx;
			}
		}

		int cy = max (sizeImageSmall.cy, m_sizeTextRight.cy) + 2 * m_szMargin.cy;

		if (sizeImageSmall.cy == 0)
		{
			cy += 2 * m_szMargin.cy;
		}

		if (HasMenu () || HasPin())
		{
			cx += GetDropDownImageWidth ();

			if (m_bIsDefaultCommand && m_nID != -1 && m_nID != 0 &&
				m_sizeTextRight.cx > 0)
			{
				cx += m_nMenuArrowMargin;
			}
		}

		if (IsDefaultMenuLook () && !IsQATMode ())
		{
			cx += 2 * globalUtils.ScaleByDPI(TEXT_MARGIN);
		}

		return CSize (cx + cxExtra, cy) + m_sizePadding;
	}

	SetMargin (CSize (nLargeButtonMarginX, nLargeButtonMarginY));

	if (IsDefaultPanelButton ())
	{
		sizeImageLarge.cx += 2 * (m_szMargin.cx + 1);
	}

	int cx = max (sizeImageLarge.cx + 2 * m_szMargin.cx, m_sizeTextBottom.cx + 5);

	if (IsDefaultPanelButton ())
	{
		cx += nDefaultPaneButtonMargin;
	}

	if (IsDefaultMenuLook ())
	{
		cx += 2 * globalUtils.ScaleByDPI(TEXT_MARGIN);
	}

	TEXTMETRIC tm;
	pDC->GetTextMetrics (&tm);

	int cyExtra = (tm.tmHeight < 15 || globalData.GetRibbonImageScale () != 1) ? 1 : 5;

	const int cyText = max (m_sizeTextBottom.cy, sizeImageLarge.cy + 1);
	const int cy = sizeImageLarge.cy + cyText + cyExtra;

	return CSize (cx + cxExtra, cy) + m_sizePadding;
}
//********************************************************************************
CSize CBCGPRibbonButton::GetCompactSize (CDC* /*pDC*/)
{
	ASSERT_VALID (this);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		m_arSubItems [i]->SetParentCategory (m_pParent);
	}

	CSize sizeImageSmall = GetImageSize (RibbonImageSmall);

	if (IsMainRibbonButton ())
	{
		return sizeImageSmall + m_sizePadding;
	}

	const int cxDropDown = GetDropDownImageWidth ();

	int cxExtra = 0;

	if (m_bQuickAccessMode || m_bFloatyMode)
	{
		SetMargin (globalData.GetRibbonImageScale () != 1. ? 
			CSize (nSmallButtonMarginX, nSmallButtonMarginY - 1) : CSize (nSmallButtonMarginX, nSmallButtonMarginY));
	
		if (sizeImageSmall == CSize (0, 0))
		{
			sizeImageSmall = globalUtils.ScaleByDPI(CSize (16, 16));
		}
	}
	else
	{
		SetMargin (IsBackstageViewMode() ? CSize(GetBackstageViewMarginX() + m_nExtraMaginX, nBackstageViewMarginY) : CSize (nSmallButtonMarginX, nSmallButtonMarginY));
		cxExtra = GetGroupButtonExtraWidth ();

		if (IsDefaultMenuLook ())
		{
			cxExtra += 2 * globalUtils.ScaleByDPI(TEXT_MARGIN);
		}
	}

	int nMenuArrowWidth = 0;
	
	if (HasMenu () || HasPin())
	{
		if (m_bIsDefaultCommand)
		{
			nMenuArrowWidth = cxDropDown + m_szMargin.cx / 2 + 1;
		}
		else
		{
			nMenuArrowWidth = cxDropDown - m_szMargin.cx / 2 - 1;
		}
	}

	int cx = sizeImageSmall.cx + 2 * m_szMargin.cx + nMenuArrowWidth + cxExtra;
	int cy = sizeImageSmall.cy + 2 * m_szMargin.cy;

	return CSize (cx, cy) + m_sizePadding;
}
//********************************************************************************
CSize CBCGPRibbonButton::GetIntermediateSize (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		m_arSubItems [i]->SetParentCategory (m_pParent);
	}

	if (m_bQuickAccessMode || m_bFloatyMode)
	{
		return GetCompactSize (pDC);
	}

	SetMargin (IsBackstageViewMode() ? CSize(GetBackstageViewMarginX() + m_nExtraMaginX, nBackstageViewMarginY) : CSize (nSmallButtonMarginX, nSmallButtonMarginY));

	const int nMenuArrowWidth = (HasMenu() || HasPin()) ? GetDropDownImageWidth () : 0;

	CSize sizeImageSmall = GetImageSize (RibbonImageSmall);

	sizeImageSmall.cy = max (16, sizeImageSmall.cy);

	int cy = max (sizeImageSmall.cy, m_sizeTextRight.cy) + 2 * m_szMargin.cy;
	int cx = sizeImageSmall.cx + 2 * m_szMargin.cx + nMenuArrowWidth + m_sizeTextRight.cx + GetTextOffset () + 
		GetGroupButtonExtraWidth () + 1;

	if (IsDefaultMenuLook ())
	{
		cx += 2 * globalUtils.ScaleByDPI(TEXT_MARGIN);
		cy += globalUtils.ScaleByDPI(TEXT_MARGIN) / 2 - 1;
	}

	if (GetBackstageAttachedView() != NULL || (IsBackstageViewMode() && CBCGPVisualManager::GetInstance ()->IsRibbonBackstageHighCommands()))
	{
		cy += 2 * nBackstageViewMarginY;
	}

	return CSize (cx, cy) + m_sizePadding;
}
//********************************************************************************
void CBCGPRibbonButton::OnLButtonDown (CPoint point)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::OnLButtonDown (point);

	if (!HasMenu () || IsMenuMode ())
	{
		return;
	}

	if (!m_rectMenu.IsRectEmpty () && !m_rectMenu.PtInRect (point))
	{
		return;
	}

	if (m_bIsDefaultCommand && m_bIsDisabled)
	{
		return;
	}

	if (m_bIsDisabled && m_rectCommand.IsRectEmpty ())
	{
		return;
	}

	OnShowPopupMenu ();
}
//********************************************************************************
void CBCGPRibbonButton::OnLButtonUp (CPoint point)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::OnLButtonUp (point);

	BOOL bIsPressed = m_bIsPressed || (IsMenuMode () && !IsBackstageViewMode());

	if (m_bIsDisabled || !bIsPressed || !m_bIsHighlighted)
	{
		return;
	}

	if (m_bIsDroppedDown)
	{
		if (!m_rectCommand.IsRectEmpty () && m_rectCommand.PtInRect (point) && IsMenuMode ())
		{
			OnClick (point);
		}

		return;
	}

	if (!m_rectCommand.IsRectEmpty () && !m_rectCommand.PtInRect (point))
	{
		return;
	}

	OnClick (point);
}
//*****************************************************************************
void CBCGPRibbonButton::OnLButtonDblClk(CPoint point)
{
	ASSERT_VALID (this);
	
	CBCGPBaseRibbonElement::OnLButtonDblClk(point);
	
	if (m_bIsDisabled || !m_bIsHighlighted || m_bIsDroppedDown)
	{
		return;
	}
	
	if (!m_rectCommand.IsRectEmpty () && !m_rectCommand.PtInRect (point))
	{
		return;
	}
	
	OnClick(point);
}
//*****************************************************************************
void CBCGPRibbonButton::OnMouseMove (CPoint point)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::OnMouseMove (point);

	if (!HasPin())
	{
		if (!HasMenu () || m_nID == -1 || m_nID == 0)
		{
			return;
		}
	}

	BOOL bMenuWasHighlighted = m_bIsMenuHighlighted;
	BOOL bCommandWasHighlighted = m_bIsCommandHighlighted;

	m_bIsMenuHighlighted = m_rectMenu.PtInRect (point);
	m_bIsCommandHighlighted = m_rectCommand.PtInRect (point);

	if (bMenuWasHighlighted != m_bIsMenuHighlighted ||
		bCommandWasHighlighted != m_bIsCommandHighlighted)
	{
		Redraw ();

		if (m_pParentMenu != NULL)
		{
			ASSERT_VALID (m_pParentMenu);
			m_pParentMenu->OnChangeHighlighted (this);
		}
	}
}
//*****************************************************************************
void CBCGPRibbonButton::OnClick (CPoint point)
{
	ASSERT_VALID (this);

	if (IsMenuMode () && HasMenu () && m_rectCommand.IsRectEmpty ())
	{
		return;
	}

	if (m_pParentMenu != NULL)
	{
		ASSERT_VALID (m_pParentMenu);
		m_pParentMenu->OnClickButton (this, point);

		return;
	}

	NotifyCommand ();
}
//*****************************************************************************
void CBCGPRibbonButton::OnAccDefaultAction()
{
	ASSERT_VALID (this);

	if (NotifyCommand(TRUE))
	{
		return;
	}

	if (HasMenu ())
	{
		OnShowPopupMenu ();
	}
}
//******************************************************************************
void CBCGPRibbonButton::OnShowPopupMenu ()
{
	ASSERT_VALID (this);

	CWnd* pWndParent = GetParentWnd ();
	if (pWndParent->GetSafeHwnd () == NULL)
	{
		ASSERT (FALSE);
		return;
	}

	CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar ();
	if (pTopLevelRibbon->GetSafeHwnd () == NULL)
	{
		ASSERT (FALSE);
		return;
	}

	CBCGPBaseRibbonElement::OnShowPopupMenu ();

	if (DisplayPopupDialog(pWndParent))
	{
		return;
	}
	
	const BOOL bIsRTL = (pTopLevelRibbon->GetExStyle () & WS_EX_LAYOUTRTL);

	CWnd* pWndOwner = pTopLevelRibbon->GetTopLevelOwner();

	if (pTopLevelRibbon->GetParentFrame()->GetSafeHwnd() != AfxGetMainWnd()->GetSafeHwnd())
	{
		pWndOwner = pTopLevelRibbon->GetParentFrame();
		if (pWndOwner->GetSafeHwnd() == NULL)
		{
			pWndOwner = pTopLevelRibbon->GetOwner();
			if (pWndOwner->GetSafeHwnd() == NULL)
			{
				ASSERT(FALSE);
				return;
			}
		}
	}

	if (m_arSubItems.GetSize () > 0)
	{
		int i = 0;

		if (m_bIsWindowsMenu)
		{
			FillWindowList();
		}
		else
		{
			BOOL bHasVisibleSubItems = FALSE;

			for (i = 0; i < m_arSubItems.GetSize (); i++)
			{
				ASSERT_VALID (m_arSubItems [i]);
				if (!m_arSubItems[i]->IsHiddenInAppMode())
				{
					bHasVisibleSubItems = TRUE;
					break;
				}
			}

			if (!bHasVisibleSubItems)
			{
				return;
			}
		}

		//--------------------------------
		// Build popup menu from subitems:
		//--------------------------------
		CBCGPRibbonPanelMenu* pMenu = new CBCGPRibbonPanelMenu
			(pTopLevelRibbon, m_arSubItems);

		pMenu->SetParentRibbonElement (this);

		pMenu->SetMenuMode ();

		BOOL bIsPopupDefaultMenuLook = IsPopupDefaultMenuLook ();

		for (i = 0; bIsPopupDefaultMenuLook && i < m_arSubItems.GetSize (); i++)
		{
			ASSERT_VALID (m_arSubItems [i]);
			
			if (!m_arSubItems [i]->IsDefaultMenuLook ())
			{
				bIsPopupDefaultMenuLook = FALSE;
			}
		}

		pMenu->SetDefaultMenuLook (bIsPopupDefaultMenuLook);

		if (m_pOriginal != NULL && m_pOriginal->GetParentPanel () != NULL &&
			m_pOriginal->GetParentPanel ()->IsMainPanel ())
		{
			pMenu->SetDefaultMenuLook (FALSE);
		}
		
		CRect rectBtn = GetRect ();
		pWndParent->ClientToScreen (&rectBtn);

		int x = m_bRightAlignMenu || bIsRTL ? rectBtn.right : rectBtn.left;

		int y = rectBtn.bottom;

		if (m_bCreatedFromMenu && m_bRightAlignMenu && !bIsRTL)
		{
			pMenu->SetRightAlign ();
		}

		if (IsMenuMode ())
		{
			x = bIsRTL ? rectBtn.left : rectBtn.right;
			y = rectBtn.top;
		}

		CRect rectMenuLocation;
		rectMenuLocation.SetRectEmpty ();

		CBCGPRibbonPanel* pPanel = GetParentPanel ();

		if (pPanel != NULL && 
			pPanel->GetPreferedMenuLocation (rectMenuLocation))
		{
			pWndParent->ClientToScreen (&rectMenuLocation);

			x = bIsRTL ? rectMenuLocation.right : rectMenuLocation.left;
			y = rectMenuLocation.top;

			pMenu->SetPreferedSize (rectMenuLocation.Size ());
			pMenu->SetDefaultMenuLook (FALSE);
		}

		pMenu->Create (pWndOwner, x, y, (HMENU) NULL);

		SetDroppedDown (pMenu);
		return;
	}

	HMENU hPopupMenu = GetMenu ();
	if (hPopupMenu == NULL)
	{
		return;
	}

	CRect rectBtn = GetRect ();
	pWndParent->ClientToScreen (&rectBtn);

	CBCGPPopupMenu* pPopupMenu = new CBCGPPopupMenu;

	pPopupMenu->SetAutoDestroy (FALSE);
	pPopupMenu->SetRightAlign (m_bRightAlignMenu && !bIsRTL);

	pPopupMenu->SetParentRibbonElement (this);

	CBCGPPopupMenu* pMenuActive = CBCGPPopupMenu::GetActiveMenu ();
	if (pMenuActive->GetSafeHwnd() != NULL &&
		pMenuActive->GetSafeHwnd () != pWndParent->GetParent ()->GetSafeHwnd ())
	{
		pMenuActive->SendMessage (WM_CLOSE);
	}

	int x = m_bRightAlignMenu || bIsRTL ? rectBtn.right : rectBtn.left;
	int y = rectBtn.bottom;

	pPopupMenu->Create (pWndOwner, x, y, hPopupMenu, FALSE);

	SetDroppedDown (pPopupMenu);
}
//******************************************************************************
void CBCGPRibbonButton::SetParentCategory (CBCGPRibbonCategory* pParent)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::SetParentCategory (pParent);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pSubItem = m_arSubItems [i];
		ASSERT_VALID (pSubItem);

		pSubItem->SetParentCategory (m_pParent);
		pSubItem->SetDefaultMenuLook (!m_bUseMenuHandle && !pSubItem->HasLargeMode ());
	}
}
//******************************************************************************
void CBCGPRibbonButton::SetParentMenu (CBCGPRibbonPanelMenuBar* pMenuBar)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::SetParentMenu (pMenuBar);

	if (IsBackstageViewMode() && m_arSubItems.GetSize() > 0)
	{
		m_arSubItems[0]->SetParentMenu(pMenuBar);
	}
}
//******************************************************************************
void CBCGPRibbonButton::CopyFrom (const CBCGPBaseRibbonElement& s)
{
	ASSERT_VALID (this);

	if (m_bAutodestroyMenu && m_hMenu != NULL)
	{
		::DestroyMenu (m_hMenu);
	}

	if (m_bAutoDestroyIcon && m_hIcon != NULL)
	{
		::DestroyIcon (m_hIcon);
	}

	if (m_bAutoDestroyIcon && m_hIconSmall != NULL)
	{
		::DestroyIcon (m_hIconSmall);
	}

	RemoveAllSubItems ();

	CBCGPBaseRibbonElement::CopyFrom (s);

	CBCGPRibbonButton& src = (CBCGPRibbonButton&) s;

	m_nSmallImageIndex = src.m_nSmallImageIndex;
	m_nLargeImageIndex = src.m_nLargeImageIndex;

	if (m_bSearchResultMode)
	{
		return;
	}

	m_hMenu = src.m_hMenu;
	m_bAutodestroyMenu = FALSE;
	m_bRightAlignMenu = src.m_bRightAlignMenu;
	m_bIsDefaultCommand = src.m_bIsDefaultCommand;
	m_szMargin = src.m_szMargin;
	m_hIcon = src.m_hIcon;
	m_hIconSmall = src.m_hIconSmall;
	m_bAutoDestroyIcon = FALSE;
	m_bAlphaBlendIcon = src.m_bAlphaBlendIcon;
	m_bSmallIconMode = src.m_bSmallIconMode;
	m_bForceDrawBorder = src.m_bForceDrawBorder;
	m_bCreatedFromMenu = src.m_bCreatedFromMenu;
	m_bIsWindowsMenu = src.m_bIsWindowsMenu;
	m_nWindowsMenuItems = src.m_nWindowsMenuItems;

	int i = 0;

	for (i = 0; i < src.m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pSrcElem = src.m_arSubItems [i];
		ASSERT_VALID (pSrcElem);

		CBCGPBaseRibbonElement* pElem =
			(CBCGPBaseRibbonElement*) pSrcElem->GetRuntimeClass ()->CreateObject ();
		ASSERT_VALID (pElem);

		pElem->CopyFrom (*pSrcElem);
		m_arSubItems.Add (pElem);
	}

	m_nWrapIndex = src.m_nWrapIndex;

	m_arWordIndexes.RemoveAll ();

	for (i = 0; i < src.m_arWordIndexes.GetSize (); i++)
	{
		m_arWordIndexes.Add (src.m_arWordIndexes [i]);
	}

	m_QATType = src.m_QATType;

	m_pRTIPopupDlg = src.m_pRTIPopupDlg;
	m_lpszPopupDlgTemplateName = src.m_lpszPopupDlgTemplateName;
	m_bIsResizablePopup = src.m_bIsResizablePopup;
	m_bIsRightAlignedPopup = src.m_bIsRightAlignedPopup;

	m_bDontCloseParentPopupOnClick = src.m_bDontCloseParentPopupOnClick;
	m_bForceSingleLineText = src.m_bForceSingleLineText;
}
//******************************************************************************
void CBCGPRibbonButton::CopyBaseFrom (const CBCGPBaseRibbonElement& s)
{
	ASSERT_VALID (this);

	if (m_bAutoDestroyIcon && m_hIcon != NULL)
	{
		::DestroyIcon (m_hIcon);
	}

	if (m_bAutoDestroyIcon && m_hIconSmall != NULL)
	{
		::DestroyIcon (m_hIconSmall);
	}

	CBCGPBaseRibbonElement::CopyBaseFrom (s);

	CBCGPRibbonButton& src = (CBCGPRibbonButton&) s;

	if (m_nSmallImageIndex == -1 && m_nLargeImageIndex == -1)
	{
		m_nSmallImageIndex = src.m_nSmallImageIndex;
		m_nLargeImageIndex = src.m_nLargeImageIndex;
	}
}
//*****************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonButton::CreateCustomCopy()
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement* pElem = CBCGPBaseRibbonElement::CreateCustomCopy();
	
	CBCGPRibbonButton* pButton = DYNAMIC_DOWNCAST(CBCGPRibbonButton, pElem);
	if (pButton == NULL)
	{
		return pElem;
	}

	pButton->m_bSmallIconMode = TRUE;

	CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar ();
	if (pRibbonBar != NULL)
	{
		ASSERT_VALID(pRibbonBar);

		if (!pRibbonBar->GetCustomizationOptions().m_bAlwaysShowSmallIcons)
		{
			pButton->m_bSmallIconMode = FALSE;
		}
	}

	return pElem;
}
//*****************************************************************************
void CBCGPRibbonButton::SetOriginal (CBCGPBaseRibbonElement* pOriginal)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::SetOriginal (pOriginal);

	CBCGPRibbonButton* pOriginalButton = DYNAMIC_DOWNCAST (CBCGPRibbonButton, m_pOriginal);

	if (pOriginalButton == NULL)
	{
		return;
	}

	ASSERT_VALID (pOriginalButton);

	if (pOriginalButton->m_arSubItems.GetSize () != m_arSubItems.GetSize ())
	{
		return;
	}

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arSubItems [i];
		ASSERT_VALID (pButton);

		pButton->SetOriginal (pOriginalButton->m_arSubItems [i]);
	}
}
//*****************************************************************************
void CBCGPRibbonButton::DrawImage (CDC* pDC, RibbonImageType type, 
								   CRect rectImage)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	if (IsSearchResultMode () && !HasImage (RibbonImageSmall))
	{
		return;
	}

	HICON hiconSimplified = NULL;
	CBCGPRibbonButton* pOrigButton = DYNAMIC_DOWNCAST(CBCGPRibbonButton, m_pOriginal);

	if (IsDrawSimplifiedIcon() && !IsDrawingOnList())
	{
		if (m_hIconSimplified == NULL)
		{
			if (pOrigButton != NULL && pOrigButton->m_hIconSimplified != NULL)
			{
				hiconSimplified = pOrigButton->m_hIconSimplified;
			}
			else
			{
				m_hIconSimplified = CreateSimplifiedIcon();
				hiconSimplified = m_hIconSimplified;
			}
		}
		else
		{
			hiconSimplified = m_hIconSimplified;
		}
	}

	if (hiconSimplified == NULL && pOrigButton != NULL)
	{
		ASSERT_VALID (pOrigButton);

		BOOL bIsDisabled = pOrigButton->m_bIsDisabled;
		BOOL bQuickAccessMode = pOrigButton->m_bQuickAccessMode;

		pOrigButton->m_bIsDisabled = m_bIsDisabled;
		pOrigButton->m_bQuickAccessMode = m_bQuickAccessMode;

		CRect rect = pOrigButton->m_rect;
		pOrigButton->m_rect = m_rect;

		BOOL bIsDontDrawSimplifiedImage = pOrigButton->m_bIsDontDrawSimplifiedImage;
		pOrigButton->m_bIsDontDrawSimplifiedImage = m_bIsDontDrawSimplifiedImage;

		pOrigButton->DrawImage (pDC, type, rectImage);

		pOrigButton->m_bIsDisabled = bIsDisabled;
		pOrigButton->m_bQuickAccessMode = bQuickAccessMode;
		pOrigButton->m_rect = rect;
		pOrigButton->m_bIsDontDrawSimplifiedImage = bIsDontDrawSimplifiedImage;
		return;
	}

	if (hiconSimplified == NULL && m_bIsCustom)
	{
		CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar ();
		if (pRibbonBar != NULL)
		{
			pRibbonBar->DrawCustomElementImage(pDC, this, type, rectImage);
		}

		return;
	}

	if (hiconSimplified != NULL || m_hIcon != NULL || (m_bSmallIconMode && m_hIconSmall != NULL))
	{
		HICON hIcon = type == RibbonImageLarge || m_hIconSmall == NULL ? m_hIcon : m_hIconSmall;
		if (hiconSimplified != NULL)
		{
			hIcon = hiconSimplified;
		}

		CSize sizeIcon = globalUtils.ScaleByDPI(type == RibbonImageLarge ? CSize (32, 32) : CSize (16, 16));

		if (m_bIsDisabled)
		{
			CBCGPToolBarImages icon;

			CSize sz = globalUtils.GetIconSize(hIcon);

			BOOL bScale = sz.cx != sizeIcon.cx || sz.cy != sizeIcon.cy;

			if (bScale)
			{
				icon.SetImageSize (type == RibbonImageLarge ? CSize (32, 32) : CSize (16, 16));
			}
			else
			{
				icon.SetImageSize (sizeIcon);
			}

			icon.AddIcon (hIcon, m_bAlphaBlendIcon || CBCGPToolBarImages::m_bIsDrawOnGlass || hiconSimplified != NULL);
			
			CBCGPDrawState ds;
			icon.PrepareDrawImage (ds, bScale ? sizeIcon : CSize(0, 0));
			icon.Draw (pDC, rectImage.left, rectImage.top, 0, FALSE, TRUE);
			icon.EndDrawImage (ds);
		}
		else
		{
			UINT diFlags = DI_NORMAL;
			
			CWnd* pWndParent = GetParentWnd ();
			if (pWndParent != NULL && (pWndParent->GetExStyle () & WS_EX_LAYOUTRTL))
			{
				diFlags |= 0x0010 /*DI_NOMIRROR*/;
			}
			
			if (m_bIsDrawingOnList && !globalData.IsRibbonImageScaleEnabled())
			{
				// Fix for CBCGPRibbonTreeCtrl::OnDrawItem:
				sizeIcon = CSize(::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON));
				rectImage.top -= max(0, (sizeIcon.cy - 16) / 2);
			}

			::DrawIconEx (pDC->GetSafeHdc (), 
				rectImage.left, 
				rectImage.top,
				hIcon, sizeIcon.cx, sizeIcon.cy, 0, NULL,
				diFlags);
		}
		return;
	}

	if (m_pParentGroup != NULL)
	{
		ASSERT_VALID (m_pParentGroup);

		if (m_pParentGroup->HasImages ())
		{
			m_pParentGroup->OnDrawImage (pDC, rectImage, this, m_nSmallImageIndex);
			return;
		}
	}

	if (m_pParent == NULL || rectImage.Width () == 0 || rectImage.Height () == 0)
	{
		return;
	}

	ASSERT_VALID (m_pParent);

	m_pParent->OnDrawImage (
		pDC, rectImage, this, 
		type == RibbonImageLarge, 
		type == RibbonImageLarge ? m_nLargeImageIndex : m_nSmallImageIndex,
		FALSE /* no center */);
}
//*****************************************************************************
HICON CBCGPRibbonButton::ExportImageToIcon(BOOL bIsLargeIcon)
{
	ASSERT_VALID (this);

	const RibbonImageType imageType = bIsLargeIcon ? RibbonImageLarge : RibbonImageSmall;

	if (!HasImage(imageType))
	{
		return NULL;
	}

	CBCGPRibbonButton* pOrigButton = DYNAMIC_DOWNCAST(CBCGPRibbonButton, m_pOriginal);
	if (pOrigButton != NULL)
	{
		ASSERT_VALID(pOrigButton);
		return pOrigButton->ExportImageToIcon(bIsLargeIcon);
	}

	if (m_hIcon != NULL)
	{
		return ::CopyIcon(m_hIcon);
	}

	if (m_hIconSmall != NULL)
	{
		return ::CopyIcon(m_hIconSmall);
	}

	if (m_pParentGroup != NULL)
	{
		ASSERT_VALID (m_pParentGroup);

		if (m_pParentGroup->HasImages() && !bIsLargeIcon)
		{
			return m_pParentGroup->GetImages().ExtractIcon(m_nSmallImageIndex);
		}
	}

	if (m_pParent == NULL)
	{
		return NULL;
	}

	ASSERT_VALID (m_pParent);

	if (bIsLargeIcon)
	{
		return (m_nLargeImageIndex < 0) ? NULL : m_pParent->GetLargeImages().ExtractIcon(m_nLargeImageIndex);
	}
	else
	{
		return (m_nSmallImageIndex < 0) ? NULL : m_pParent->GetSmallImages().ExtractIcon(m_nSmallImageIndex);
	}
}
//*****************************************************************************
BOOL CBCGPRibbonButton::HasImage (RibbonImageType type) const
{
	ASSERT_VALID (this);

	CBCGPRibbonButton* pOrigButton = DYNAMIC_DOWNCAST (
		CBCGPRibbonButton, m_pOriginal);

	if (pOrigButton != NULL)
	{
		ASSERT_VALID (pOrigButton);
		return pOrigButton->HasImage (type);
	}

	if (m_bSmallIconMode)
	{
		if (type == RibbonImageLarge)
		{
			return FALSE;
		}

		return m_hIconSmall != NULL;
	}

	if (m_hIcon != NULL)
	{
		HICON hIcon = type == RibbonImageLarge || m_hIconSmall == NULL ? m_hIcon : m_hIconSmall;
		return hIcon != NULL;
	}

	if (m_bIsCustom)
	{
		CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar ();
		if (pRibbonBar != NULL)
		{
			return pRibbonBar->GetCustomElementImageSize(this, type) != CSize(0, 0);
		}

		return TRUE;
	}

	int nIndex = type == RibbonImageLarge ? m_nLargeImageIndex : m_nSmallImageIndex;
	return nIndex >= 0;
}
//*****************************************************************************
CSize CBCGPRibbonButton::GetImageSize (RibbonImageType type) const
{
	ASSERT_VALID (this);

	CBCGPRibbonButton* pOrigButton = DYNAMIC_DOWNCAST (
		CBCGPRibbonButton, m_pOriginal);

	if (pOrigButton != NULL)
	{
		ASSERT_VALID (pOrigButton);
		return pOrigButton->GetImageSize (type);
	}

	if (IsBackstageViewMode() && !IsDisplayIconInBackstageViewMode())
	{
		return CSize(0, 0);
	}

	if (m_hIcon != NULL || (m_bSmallIconMode && m_hIconSmall != NULL))
	{
		if (m_bSmallIconMode && type == RibbonImageLarge)
		{
			return CSize(0, 0);
		}

		return globalUtils.ScaleByDPI(type == RibbonImageLarge ? CSize (32, 32) : CSize (16, 16));
	}

	if (m_bIsCustom)
	{
		CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar ();
		if (pRibbonBar != NULL)
		{
			return pRibbonBar->GetCustomElementImageSize(this, type);
		}

		return CSize (0, 0);
	}

	const int nImageIndex = type == RibbonImageLarge  ? 
		m_nLargeImageIndex : m_nSmallImageIndex;

	if (nImageIndex < 0)
	{
		return CSize (0, 0);
	}

	if (m_pParentGroup != NULL)
	{
		ASSERT_VALID (m_pParentGroup);

		if (m_pParentGroup->HasImages ())
		{
			return m_pParentGroup->GetImageSize ();
		}
	}

	if (m_pParent == NULL)
	{
		return CSize (0, 0);
	}

	ASSERT_VALID (m_pParent);

	const int nImageCount = m_pParent->GetImageCount (type == RibbonImageLarge);

	if (nImageIndex >= nImageCount)
	{
		return CSize (0, 0);
	}

	return m_pParent->GetImageSize (type == RibbonImageLarge);
}
//*****************************************************************************
BOOL CBCGPRibbonButton::CanBeStretched ()
{
	ASSERT_VALID (this);
	return GetImageSize (RibbonImageLarge) != CSize (0, 0);
}
//*****************************************************************************
void CBCGPRibbonButton::AddSubItem (CBCGPBaseRibbonElement* pSubItem, int nIndex)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pSubItem);

	pSubItem->SetParentCategory (m_pParent);

	if (m_pParent == NULL)
	{
		pSubItem->SetParentRibbonBar(m_pRibbonBar);
	}

	pSubItem->SetDefaultMenuLook (!m_bUseMenuHandle && !pSubItem->HasLargeMode ());
	pSubItem->SetPadding(m_sizePadding);

	if (nIndex == -1)
	{
		m_arSubItems.Add (pSubItem);
	}
	else
	{
		m_arSubItems.InsertAt (nIndex, pSubItem);
	}

	if (IsBackstageViewMode())
	{
		pSubItem->SetBackstageViewMode();
	}

	pSubItem->OnAfterAddToParent (this);
}
//*****************************************************************************
int CBCGPRibbonButton::FindSubItemIndexByID (UINT uiID) const
{
	ASSERT_VALID (this);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		
		if (m_arSubItems [i]->GetID () == uiID)
		{
			return i;
		}
	}

	return -1;
}
//*****************************************************************************
BOOL CBCGPRibbonButton::RemoveSubItem (int nIndex)
{
	ASSERT_VALID (this);

	if (nIndex < 0 || nIndex >= m_arSubItems.GetSize ())
	{
		return FALSE;
	}

	ASSERT_VALID (m_arSubItems [nIndex]);
	delete m_arSubItems [nIndex];

	m_arSubItems.RemoveAt (nIndex);

	return TRUE;
}
//*****************************************************************************
void CBCGPRibbonButton::RemoveAllSubItems ()
{
	ASSERT_VALID (this);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		delete m_arSubItems [i];
	}

	m_arSubItems.RemoveAll ();
}
//*****************************************************************************
COLORREF CBCGPRibbonButton::OnFillBackground (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	COLORREF clrText = CBCGPVisualManager::GetInstance ()->
		OnFillRibbonButton (pDC, this);

	return clrText;
}
//*****************************************************************************
void CBCGPRibbonButton::OnDrawBorder (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	const BOOL bIsDisabled = m_bIsDisabled;

	if (m_bIsDisabled && HasMenu ())
	{
		m_bIsDisabled = FALSE;
	}

	CBCGPVisualManager::GetInstance ()->OnDrawRibbonButtonBorder (pDC, this);

	m_bIsDisabled = bIsDisabled;
}
//*****************************************************************************
int CBCGPRibbonButton::AddToListBox (CBCGPRibbonCommandsListBox* pWndListBox, BOOL bDeep)
{
	ASSERT_VALID (this);

	int nIndex = CBCGPBaseRibbonElement::AddToListBox (pWndListBox, bDeep);

	if (bDeep && !m_bCreatedFromMenu)
	{
		for (int i = 0; i < m_arSubItems.GetSize (); i++)
		{
			ASSERT_VALID (m_arSubItems [i]);

			if (m_arSubItems [i]->GetID () != 0)	// Don't add separators
			{
				nIndex = m_arSubItems [i]->AddToListBox (pWndListBox, TRUE);
			}
		}
	}

	return nIndex;
}
//*****************************************************************************
void CBCGPRibbonButton::SetCustom()
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::SetCustom();

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		m_arSubItems [i]->SetCustom();
	}
}
//*****************************************************************************
CBCGPGridRow* CBCGPRibbonButton::AddToTree (CBCGPGridCtrl* pGrid, CBCGPGridRow* pParent)
{
#ifndef BCGP_EXCLUDE_GRID_CTRL
	ASSERT_VALID (this);

	CBCGPGridRow* pRow = CBCGPBaseRibbonElement::AddToTree(pGrid, pParent);

	if (!m_bCreatedFromMenu && pRow != NULL)
	{
		if (m_arSubItems.GetSize () > 0 && pParent != NULL)
		{
			pRow->AllowSubItems();
		}

		for (int i = 0; i < m_arSubItems.GetSize (); i++)
		{
			CBCGPBaseRibbonElement* pSubItem = m_arSubItems[i];
			ASSERT_VALID(pSubItem);

			const BOOL bIsSeparator = pSubItem->GetID() == 0;
			const BOOL bIsWindow = pSubItem->GetID() >= AFX_IDM_WINDOW_FIRST && pSubItem->GetID() <= AFX_IDM_WINDOW_LAST;

			if (!bIsSeparator && !bIsWindow)	// Don't add separators and windows
			{
				pSubItem->AddToTree(pGrid, pParent != NULL ? pRow : NULL);
			}
		}
	}

	if (pRow != NULL && pRow->IsGroup())
	{
		pRow->Expand(FALSE);
	}

	return pRow;
#else
	return NULL;
#endif
}
//*****************************************************************************
void CBCGPRibbonButton::ClosePopupMenu ()
{
	ASSERT_VALID (this);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		m_arSubItems [i]->ClosePopupMenu ();
	}

	CBCGPBaseRibbonElement::ClosePopupMenu ();
}
//******************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonButton::FindByID (UINT uiCmdID)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement* pElem = CBCGPBaseRibbonElement::FindByID (uiCmdID);
	if (pElem != NULL)
	{
		ASSERT_VALID (pElem);
		return pElem;
	}

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arSubItems [i];
		ASSERT_VALID (pButton);

		pElem = pButton->FindByID (uiCmdID);
		if (pElem != NULL)
		{
			ASSERT_VALID (pElem);
			return pElem;
		}
	}
	
	return NULL;
}
//******************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonButton::FindByIDNonCustom (UINT uiCmdID)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement* pElem = CBCGPBaseRibbonElement::FindByIDNonCustom (uiCmdID);
	if (pElem != NULL)
	{
		ASSERT_VALID (pElem);
		return pElem;
	}

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arSubItems [i];
		ASSERT_VALID (pButton);

		pElem = pButton->FindByIDNonCustom (uiCmdID);
		if (pElem != NULL)
		{
			ASSERT_VALID (pElem);
			return pElem;
		}
	}
	
	return NULL;
}
//******************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonButton::FindByData (DWORD_PTR dwData)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement* pElem = CBCGPBaseRibbonElement::FindByData (dwData);
	if (pElem != NULL)
	{
		ASSERT_VALID (pElem);
		return pElem;
	}

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arSubItems [i];
		ASSERT_VALID (pButton);

		pElem = pButton->FindByData (dwData);
		if (pElem != NULL)
		{
			ASSERT_VALID (pElem);
			return pElem;
		}
	}
	
	return NULL;
}
//******************************************************************************
CString CBCGPRibbonButton::GetToolTipText () const
{
	ASSERT_VALID (this);

	if (!IsCustom() && !m_bQuickAccessMode && m_bAlwaysShowDescription && !m_strDescription.IsEmpty ())
	{
		return _T("");
	}

	return CBCGPBaseRibbonElement::GetToolTipText ();
}
//******************************************************************************
void CBCGPRibbonButton::SetParentRibbonBar (CBCGPRibbonBar* pRibbonBar)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::SetParentRibbonBar (pRibbonBar);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arSubItems [i];
		ASSERT_VALID (pButton);

		pButton->SetParentRibbonBar (pRibbonBar);
	}
}
//*************************************************************************************
CRect CBCGPRibbonButton::GetKeyTipRect (CDC* pDC, BOOL bIsMenu)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	CSize sizeKeyTip = GetKeyTipSize (pDC);
	CRect rectKeyTip (0, 0, 0, 0);

	if (sizeKeyTip == CSize (0, 0) || m_rect.IsRectEmpty ())
	{
		return rectKeyTip;
	}

	rectKeyTip.left = bIsMenu ? m_rect.right - sizeKeyTip.cx / 2 : m_rect.left + 10;
	rectKeyTip.top = m_rect.bottom - sizeKeyTip.cy / 2;

	CRect rectPanel;
	rectPanel.SetRectEmpty ();

	CBCGPRibbonPanel* pPanel = GetParentPanel ();
	if (pPanel != NULL && !IsMenuMode () && !(m_bQuickAccessMode && m_bFloatyMode) && !IsDefaultPanelButton ())
	{
		ASSERT_VALID (pPanel);

		rectPanel = pPanel->GetRect ();

		if (!rectPanel.IsRectEmpty ())
		{
			rectPanel.bottom -= pPanel->GetCaptionHeight ();
			rectKeyTip.top = rectPanel.bottom - sizeKeyTip.cy / 2;
		}
	}

	if (IsBackstageViewMode())
	{
		if (HasSubitems())
		{
			rectKeyTip.top = m_rect.top - 1;
			rectKeyTip.left = m_rect.left + m_szMargin.cx - sizeKeyTip.cx;
		}
		else
		{
			rectKeyTip.top = m_rect.CenterPoint().y - sizeKeyTip.cy;
			rectKeyTip.left = m_rect.left + m_szMargin.cx + sizeKeyTip.cx / 2;
		}
	}
	else if (IsDefaultPanelButton () && !m_bQuickAccessMode && !m_bFloatyMode)
	{
		rectKeyTip.top = m_rect.bottom;
		rectKeyTip.left = m_rect.CenterPoint ().x - sizeKeyTip.cx / 2;
	}
	else if (IsMainRibbonButton ())
	{
		// Center key tip:
		rectKeyTip.top = m_rect.CenterPoint ().y - sizeKeyTip.cy / 2;
		rectKeyTip.left = m_rect.CenterPoint ().x - sizeKeyTip.cx / 2;
	}
	else if (m_bIsLargeImage || m_bFloatyMode)
	{
		if (m_bTextAlwaysOnRight)
		{
			if (!bIsMenu)
			{
				rectKeyTip.left = m_rect.left + GetImageSize (RibbonImageLarge).cx - sizeKeyTip.cx + 4;
			}

			rectKeyTip.top = m_rect.bottom - sizeKeyTip.cy - 4;
		}
		else if (!bIsMenu)
		{
			rectKeyTip.left = m_rect.CenterPoint ().x - sizeKeyTip.cx / 2;
		}
	}
	else if (IsMenuMode ())
	{
		rectKeyTip.top = m_rect.CenterPoint ().y;
	}
	else
	{
		if (m_bQuickAccessMode)
		{
			rectKeyTip.left = m_rect.CenterPoint ().x - sizeKeyTip.cx / 2;
			rectKeyTip.top = m_rect.CenterPoint ().y;
		}

		if (!rectPanel.IsRectEmpty ())
		{
			if (m_rect.top < rectPanel.CenterPoint ().y &&
				m_rect.bottom > rectPanel.CenterPoint ().y)
			{
				rectKeyTip.top = m_rect.CenterPoint ().y - sizeKeyTip.cy / 2;
			}
			else if (m_rect.top < rectPanel.CenterPoint ().y)
			{
				rectKeyTip.top = m_rect.top - sizeKeyTip.cy / 2;
			}
		}
	}

	if (m_bIsTabElement)
	{
		rectKeyTip.top += 2;
	}

	rectKeyTip.right = rectKeyTip.left + sizeKeyTip.cx;
	rectKeyTip.bottom = rectKeyTip.top + sizeKeyTip.cy;

	return rectKeyTip;
}
//*************************************************************************************
BOOL CBCGPRibbonButton::OnKey (BOOL bIsMenuKey)
{
	ASSERT_VALID (this);

	if (IsDisabled ())
	{
		return FALSE;
	}

	if (m_rect.IsRectEmpty ())
	{
		return CBCGPBaseRibbonElement::OnKey (bIsMenuKey);
	}

	if (IsBackstageViewMode() && HasSubitems())
	{
		CBCGPRibbonPanel* pPanel = GetParentPanel ();
		if (pPanel != NULL)
		{
			ASSERT_VALID (pPanel);
			pPanel->SelectView(this);

			return TRUE;
		}
	}

	CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar ();

	if (HasMenu () && (bIsMenuKey || m_strMenuKeys.IsEmpty ()))
	{
		if (IsDroppedDown ())
		{
			return TRUE;
		}

		if (pTopLevelRibbon != NULL)
		{
			pTopLevelRibbon->HideKeyTips ();
		}

		CBCGPRibbonPanel* pPanel = GetParentPanel ();
		if (pPanel != NULL)
		{
			ASSERT_VALID (pPanel);
			pPanel->SetFocused (this);
		}

		OnShowPopupMenu ();

		if (m_pPopupMenu != NULL)
		{
			ASSERT_VALID (m_pPopupMenu);
			m_pPopupMenu->SendMessage (WM_KEYDOWN, VK_HOME);
		}

		return m_hMenu != NULL;
	}

	if (pTopLevelRibbon != NULL && pTopLevelRibbon->GetTopLevelFrame () != NULL)
	{
		pTopLevelRibbon->GetTopLevelFrame ()->SetFocus ();
	}

	OnClick (m_rect.TopLeft ());
	return TRUE;
}
//*************************************************************************************
void CBCGPRibbonButton::GetElementsByID (UINT uiCmdID,
		CArray <CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arElements)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::GetElementsByID (uiCmdID, arElements);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arSubItems [i];
		ASSERT_VALID (pButton);

		pButton->GetElementsByID (uiCmdID, arElements);
	}
}
//*************************************************************************************
void CBCGPRibbonButton::GetElementsByName (LPCTSTR lpszName,
		CArray <CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arElements, DWORD dwFlags)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::GetElementsByName (lpszName, arElements, dwFlags);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arSubItems [i];
		ASSERT_VALID (pButton);

		pButton->GetElementsByName (lpszName, arElements, dwFlags);
	}
}
//*************************************************************************************
BOOL CBCGPRibbonButton::QueryElements(const CStringArray& arWords, CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arButtons, int nMaxResults, BOOL bDescription, BOOL bAll)
{
	ASSERT_VALID (this);

	if (IsPopupDialogEnabled())
	{
		return FALSE;
	}

	int nCount = (int) arButtons.GetSize();

	BOOL bRes = CBCGPBaseRibbonElement::QueryElements(arWords, arButtons, nMaxResults, bDescription, bAll);

	if (nCount < (int) arButtons.GetSize())
	{
		return bRes;
	}

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arSubItems [i];
		ASSERT_VALID (pButton);
		
		CArray <CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arSubitems;

		pButton->QueryElements(arWords, arSubitems, nMaxResults, bDescription, bAll);

		if (arSubitems.GetSize() > 0)
		{
			arButtons.Add(this);
			return arButtons.GetSize() >= nMaxResults;
		}
	}

	return FALSE;
}
//*************************************************************************************
void CBCGPRibbonButton::GetElements (
		CArray <CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arElements)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::GetElements (arElements);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arSubItems [i];
		ASSERT_VALID (pButton);

		pButton->GetElements (arElements);
	}
}
//******************************************************************************
void CBCGPRibbonButton::OnAfterChangeRect (CDC* pDC)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::OnAfterChangeRect (pDC);

	if (IsMainRibbonButton ())
	{
		m_bIsLargeImage = TRUE;
		return;
	}

	m_bIsLargeImage = FALSE;

	if (m_bQuickAccessMode || m_bFloatyMode || m_bSearchResultMode)
	{
		return;
	}

	CSize sizeImageLarge = GetImageSize (RibbonImageLarge);
	CSize sizeImageSmall = GetImageSize (RibbonImageSmall);

	if (m_bCompactMode || m_bIntermediateMode)
	{
		m_bIsLargeImage = FALSE;

		if (sizeImageLarge != CSize (0, 0) && sizeImageSmall == CSize (0, 0) && !m_bIsCustom)
		{
			m_bIsLargeImage = TRUE;
		}
	}
	else
	{
		BOOL bIsSmallIcon = FALSE;

		if (m_hIcon != NULL || (m_bSmallIconMode && m_hIconSmall != NULL))
		{
			CSize sizeIcon = globalUtils.ScaleByDPI(m_bSmallIconMode ? CSize(16, 16) : CSize (32, 32));

			bIsSmallIcon =	sizeIcon.cx > m_rect.Width () ||
							sizeIcon.cy > m_rect.Height ();
		}

		if (sizeImageLarge != CSize (0, 0) && !bIsSmallIcon)
		{
			m_bIsLargeImage = TRUE;
		}
	}

	if (m_bIsLargeImage)
	{
		SetMargin (CSize (nLargeButtonMarginX, nLargeButtonMarginY));
	}
	else if (m_szMargin == CSize (nLargeButtonMarginX, nLargeButtonMarginY))
	{
		SetMargin (IsBackstageViewMode() ? CSize(GetBackstageViewMarginX() + m_nExtraMaginX, nBackstageViewMarginY) : CSize (nSmallButtonMarginX, nSmallButtonMarginY));
	}
}
//******************************************************************************
BOOL CBCGPRibbonButton::AddFrameMDIChildren(CBCGPMDIFrameWnd* pMDIFrameWnd)
{
	ASSERT_VALID(pMDIFrameWnd);

	CWnd* pMDIClient = CWnd::FromHandle(pMDIFrameWnd->m_hWndMDIClient);
	ASSERT_VALID(pMDIClient);

	CMDIChildWnd* pActiveChild = NULL;
	if (CBCGPMDIFrameWnd::GetActiveTearOffFrame() == NULL && !pMDIFrameWnd->IsTearOff())
	{
		pActiveChild = pMDIFrameWnd->MDIGetActive();
	}

	for (CWnd* pWndChild = pMDIClient->GetWindow(GW_CHILD); pWndChild != NULL; pWndChild = pWndChild->GetWindow(GW_HWNDNEXT))
	{
		CBCGPMDIChildWnd* pFrame = DYNAMIC_DOWNCAST (CBCGPMDIChildWnd, pWndChild);
		if (pFrame == NULL || !pFrame->CanShowOnWindowsList())
		{
			continue;
		}
		
		if (m_nWindowsMenuItems == 0)
		{
			const BOOL bPrevIsSeparator = m_arSubItems.GetSize() > 0 && DYNAMIC_DOWNCAST(CBCGPRibbonSeparator, m_arSubItems[m_arSubItems.GetSize() - 1]) != NULL;
			if (!bPrevIsSeparator)
			{
				CBCGPRibbonSeparator* pSeparator = new CBCGPRibbonSeparator (TRUE);
				pSeparator->SetDefaultMenuLook ();
				
				AddSubItem (pSeparator);
			}

			m_nWindowsMenuItems = 1;
		}

		CString strTitle;
		pFrame->GetWindowText(strTitle);
		
		CString strItem;
		strItem.Format (_T("&%d %s"), m_nWindowsMenuItems, (LPCTSTR)strTitle);
		
		CBCGPRibbonButton* pItem = new CBCGPRibbonButton (AFX_IDM_FIRST_MDICHILD, strItem);
		pItem->SetData ((DWORD_PTR)pFrame->GetSafeHwnd());
		pItem->SetDefaultMenuLook ();
		pItem->m_pRibbonBar = m_pRibbonBar;
		
		if (pActiveChild->GetSafeHwnd() == pFrame->GetSafeHwnd())
		{
			pItem->m_bIsChecked = TRUE;
		}
		
		AddSubItem (pItem);
		pItem->m_bIsTemporary = TRUE;
		
		if (m_nWindowsMenuItems == nMaxMDIChildWindows)
		{
			return FALSE;
		}

		m_nWindowsMenuItems++;
	}

	return TRUE;
}
//******************************************************************************
void CBCGPRibbonButton::FillWindowList ()
{
	if (m_nWindowsMenuItems > 0)
	{
		for (int i = 0; i < m_nWindowsMenuItems; i++)
		{
			int nIndex = (int) m_arSubItems.GetSize () - 1;

			delete m_arSubItems [nIndex];
			m_arSubItems.RemoveAt (nIndex);
		}
	}

	m_nWindowsMenuItems = 0;

	CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar ();
	if (pTopLevelRibbon == NULL)
	{
		return;
	}

	CBCGPMDIFrameWnd* pMDIFrameWnd = DYNAMIC_DOWNCAST (
		CBCGPMDIFrameWnd, pTopLevelRibbon->GetTopLevelFrame ());
	if (pMDIFrameWnd == NULL)
	{
		return;
	}

	if (pMDIFrameWnd->AreMDIChildrenTearOff())
	{
		const CList<CFrameWnd*, CFrameWnd*>& lstFrames = CBCGPFrameImpl::GetFrameList();

		for (POSITION pos = lstFrames.GetHeadPosition(); pos != NULL;)
		{
			CBCGPMDIFrameWnd* pMDIFrameCurr = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, lstFrames.GetNext(pos));
			if (pMDIFrameCurr->GetSafeHwnd() != NULL)
			{
				if (!AddFrameMDIChildren(pMDIFrameCurr))
				{
					break;
				}
			}
		}
	}
	else
	{
		AddFrameMDIChildren(pMDIFrameWnd);
	}

	if (pMDIFrameWnd->m_uiWindowsDlgMenuId != 0 &&
		(m_nWindowsMenuItems >= nMaxMDIChildWindows || pMDIFrameWnd->m_bShowWindowsDlgAlways))
	{
		//-------------------------
		// Add "Windows..." dialog:
		//-------------------------
		CBCGPRibbonButton* pItem = new CBCGPRibbonButton (
			pMDIFrameWnd->m_uiWindowsDlgMenuId, pMDIFrameWnd->m_strWindowsDlgMenuText);
		pItem->SetDefaultMenuLook ();
		pItem->m_pRibbonBar = m_pRibbonBar;

		AddSubItem (pItem);
		pItem->m_bIsTemporary = TRUE;

		m_nWindowsMenuItems++;
	}
}
//******************************************************************************
int CBCGPRibbonButton::GetGroupButtonExtraWidth ()
{
	if (m_pParentGroup == NULL)
	{
		return 0;
	}

	ASSERT_VALID (m_pParentGroup);

	switch (m_pParentGroup->GetCount ())
	{
	case 1:
		return 2;

	case 2:
		if (m_Location != RibbonElementFirstInGroup)
		{
			return 0;
		}
		break;
	}

	return m_Location == RibbonElementFirstInGroup || m_Location == RibbonElementLastInGroup ? 1 : 2;
}
//******************************************************************************
CSize CBCGPRibbonButton::DrawBottomText (CDC* pDC, BOOL bCalcOnly)
{
	ASSERT_VALID (this);

	if (m_pParent == NULL)
	{
		return CSize (0, 0);
	}

	if (m_strText.IsEmpty ())
	{
		return CSize (0, 0);
	}

	ASSERT_VALID (m_pParent);

	CSize sizeImageLarge = CSize (0, 0);

	if (m_pParent->IsCustom() && (m_nID == 0 || m_nID == (UINT)-1))
	{
		sizeImageLarge = globalUtils.ScaleByDPI(CSize(32, 32));
	}
	else
	{
		if (m_bIsCustom || m_pParent->IsCustom())
		{
			if (IsDefaultPanelButton ())
			{
				sizeImageLarge = globalUtils.ScaleByDPI(CSize(32, 32));
			}
			else
			{
				CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar();
				if (pRibbonBar != NULL)
				{
					sizeImageLarge = pRibbonBar->GetCustomElementImageSize(this, RibbonImageLarge);
				}
			}
		}
		else
		{
			sizeImageLarge = m_pParent->GetImageSize(TRUE);
		}
	}

	if (sizeImageLarge == CSize (0, 0))
	{
		return CSize (0, 0);
	}

	CSize sizeText = pDC->GetTextExtent (m_strText);

	const int nTextLineHeight = sizeText.cy;
	const int nMenuArrowWidth = (HasMenu () || IsDefaultPanelButton ()) ? (CBCGPMenuImages::Size ().cx) : 0;

	if (bCalcOnly)
	{
		m_nWrapIndex = -1;
		int nTextWidth = 0;

		if (m_arWordIndexes.GetSize () == 0 || m_bForceSingleLineText)
		{
			nTextWidth = sizeText.cx;
		}
		else
		{
			nTextWidth = 32767;

			for (int i = 0; i < m_arWordIndexes.GetSize (); i++)
			{
				int nIndex = m_arWordIndexes [i];

				CString strLineOne = m_strText.Left (nIndex);

				if (!IsDefaultPanelButton ())
				{
					globalUtils.RemoveSingleAmp(strLineOne);
				}

				const int cx1 = pDC->GetTextExtent (strLineOne).cx;

				CString strLineTwo = m_strText.Mid (nIndex + 1);

				if (!IsDefaultPanelButton ())
				{
					globalUtils.RemoveSingleAmp(strLineTwo);
				}

				const int cx2 = pDC->GetTextExtent (strLineTwo).cx + nMenuArrowWidth;

				int nWidth = max (cx1, cx2);

				if (nWidth < nTextWidth)
				{
					nTextWidth = nWidth;
					m_nWrapIndex = nIndex;
				}
			}
		}

		if (nTextWidth % 2)
		{
			nTextWidth--;
		}

		CSize size (nTextWidth, nTextLineHeight * 2);
		return size;
	}

	int y = m_rect.top + nLargeButtonMarginY + sizeImageLarge.cy + 5;
	CRect rectMenuArrow (0, 0, 0, 0);

	if (IsDefaultPanelButton ())
	{
		y += nDefaultPaneButtonMargin;
	}

	CRect rectText = m_rect;
	rectText.top = y;

	UINT uiDTFlags = DT_SINGLELINE | DT_CENTER;
	if (IsDefaultPanelButton ())
	{
		uiDTFlags |= DT_NOPREFIX;
	}

	if (m_nWrapIndex == -1)
	{
		if (m_bForceSingleLineText && pDC->GetTextExtent(m_strText).cx > rectText.Width())
		{
			uiDTFlags = DT_SINGLELINE | DT_LEFT | DT_END_ELLIPSIS;
		}

		// Single line text
		pDC->DrawText (m_strText, rectText, uiDTFlags);

		if (HasMenu () || IsDefaultPanelButton ())
		{
			rectMenuArrow = m_rect;

			rectMenuArrow.top = y + nTextLineHeight + 2;
			rectMenuArrow.left = m_rect.CenterPoint ().x - CBCGPMenuImages::Size ().cx / 2 - 1;
		}
	}
	else
	{
		CString strLineOne = m_strText.Left (m_nWrapIndex);
		pDC->DrawText (strLineOne, rectText, uiDTFlags);

		rectText.top = y + nTextLineHeight;
		rectText.right -= nMenuArrowWidth;

		if (globalData.GetRibbonImageScale () > 1.)
		{
			rectText.top -= 4;
		}

		CString strLineTwo = m_strText.Mid (m_nWrapIndex + 1);
		pDC->DrawText (strLineTwo, rectText, uiDTFlags);

		if (HasMenu () || IsDefaultPanelButton ())
		{
			rectMenuArrow = rectText;

			if (!IsDefaultPanelButton ())
			{
				globalUtils.RemoveSingleAmp(strLineTwo);
			}

			rectMenuArrow.top += globalUtils.ScaleByDPI(2);

			rectMenuArrow.left = rectText.right - (rectText.Width () - pDC->GetTextExtent (strLineTwo).cx) / 2 + globalUtils.ScaleByDPI(1);
			if (rectMenuArrow.left + nMenuArrowWidth > m_rect.right - 2)
			{
				rectMenuArrow.left = m_rect.right - nMenuArrowWidth - 2;
			}
		}
	}

	if (!rectMenuArrow.IsRectEmpty ())
	{
		int nMenuArrowHeight = CBCGPMenuImages::Size ().cy;

		rectMenuArrow.bottom = rectMenuArrow.top + nMenuArrowHeight;
		rectMenuArrow.right = rectMenuArrow.left + nMenuArrowWidth;

		CBCGPMenuImages::IMAGES_IDS id = 
			globalData.GetRibbonImageScale () > 1. ? 
			CBCGPMenuImages::IdArowDownLarge : CBCGPMenuImages::IdArowDown;

		CBCGPVisualManager::GetInstance()->OnDrawRibbonMenuArrow(pDC, this, (UINT)id, rectMenuArrow);
	}

	return CSize (0, 0);
}
//******************************************************************************
BOOL CBCGPRibbonButton::SetACCData (CWnd* pParent, CBCGPAccessibilityData& data)
{
	ASSERT_VALID (this);

	if (!CBCGPBaseRibbonElement::SetACCData (pParent, data))
	{
		return FALSE;
	}

	if (HasMenu ())
	{
		data.m_nAccRole = ROLE_SYSTEM_BUTTONDROPDOWN;
	
		if (!IsCommandAreaHighlighted ())
		{
			data.m_bAccState |= STATE_SYSTEM_HASPOPUP;
			data.m_strAccDefAction = _T("Open");

			if (IsDroppedDown ())
			{
				data.m_bAccState |= STATE_SYSTEM_PRESSED;
				data.m_strAccDefAction = _T("Close");
			}
		}
	}

	return TRUE;
}
//*************************************************************************************
void CBCGPRibbonButton::GetItemIDsList (CList<UINT,UINT>& lstItems) const
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::GetItemIDsList (lstItems);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		m_arSubItems [i]->GetItemIDsList (lstItems);
	}
}
//*************************************************************************************
void CBCGPRibbonButton::OnChangeVisualManager()
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::OnChangeVisualManager();

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		m_arSubItems [i]->OnChangeVisualManager();
	}
}
//*************************************************************************************
BOOL CBCGPRibbonButton::ReplaceSubitemByID (UINT uiCmdID, CBCGPBaseRibbonElement* pElem, BOOL bCopyContent)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pElem);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pSubItem = m_arSubItems[i];
		ASSERT_VALID (pSubItem);

		if (pSubItem->GetID () == uiCmdID)
		{
			if (bCopyContent)
			{
				pElem->CopyFrom(*pSubItem);
			}
			else
			{
				pElem->CopyBaseFrom(*pSubItem);
			}

			m_arSubItems[i] = pElem;

			delete pSubItem;
			return TRUE;
		}

		if (pSubItem->ReplaceSubitemByID (uiCmdID, pElem, bCopyContent))
		{
			return TRUE;
		}
	}
	
	return FALSE;
}
//******************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonButton::CreateQATCopy()
{
	ASSERT_VALID (this);

	if (m_QATType == BCGPRibbonButton_Show_Default)
	{
		return CBCGPBaseRibbonElement::CreateQATCopy();
	}

	CBCGPRibbonCheckBox* pButton = NULL;

	switch (m_QATType)
	{
	case BCGPRibbonButton_Show_As_CheckBox:
		pButton = new CBCGPRibbonCheckBox;
		break;

	case BCGPRibbonButton_Show_As_RadioButton:
		pButton = new CBCGPRibbonRadioButton;
		break;

	default:
		ASSERT(FALSE);
		return NULL;
	}

	ASSERT_VALID (pButton);

	pButton->CopyFrom(*this);

	if (m_bIsTemporary)
	{
		pButton->SetOriginal(NULL);
	}
	else
	{
		if (m_pOriginal != NULL)
		{
			pButton->SetOriginal(m_pOriginal);
		}
		else
		{
			pButton->SetOriginal(this);
		}
	}

	pButton->m_bQuickAccessMode = TRUE;

	return pButton;
}
//******************************************************************************
CSize CBCGPRibbonButton::GetToolTipImageSize(int& nRibbonImageType) const
{
	ASSERT_VALID(this);

	if (!IsDrawTooltipImage ())
	{
		return CSize (0, 0);
	}
	
	if (GetIcon () != NULL)
	{
		nRibbonImageType = IsLargeImage () ? CBCGPRibbonButton::RibbonImageLarge : CBCGPRibbonButton::RibbonImageSmall;
		return GetImageSize((CBCGPBaseRibbonElement::RibbonImageType)nRibbonImageType);
	}
	
	if (GetIcon (FALSE) != NULL)
	{
		nRibbonImageType = CBCGPRibbonButton::RibbonImageSmall;
		return GetImageSize ((CBCGPBaseRibbonElement::RibbonImageType)nRibbonImageType);
	}
	
	CSize sizeLarge (0, 0);
	CSize sizeSmall (0, 0);
	
	if (IsLargeImage () && GetImageIndex (TRUE) >= 0)
	{
		sizeLarge = GetImageSize (CBCGPRibbonButton::RibbonImageLarge);
	}
	
	if (sizeLarge != CSize (0, 0))
	{
		nRibbonImageType = CBCGPRibbonButton::RibbonImageLarge;
		return sizeLarge;
	}
	
	if (GetImageIndex (FALSE) >= 0)
	{
		sizeSmall = GetImageSize (CBCGPRibbonButton::RibbonImageSmall);
	}
	
	nRibbonImageType = CBCGPRibbonButton::RibbonImageSmall;
	
	return sizeSmall;
}
//******************************************************************************
void CBCGPRibbonButton::OnDrawTooltipImage(CDC* pDC, RibbonImageType type, const CRect& rectImage)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	BOOL bIsDisabled = m_bIsDisabled;
	m_bIsDisabled = FALSE;
	m_bIsDontDrawSimplifiedImage = TRUE;
	
	DrawImage (pDC, type, rectImage);
	
	m_bIsDisabled = bIsDisabled;
	m_bIsDontDrawSimplifiedImage = FALSE;
}
//*******************************************************************************
void CBCGPRibbonButton::SetPadding(const CSize& sizePadding)
{
	ASSERT_VALID(this);

	CBCGPBaseRibbonElement::SetPadding(sizePadding);
	
	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		m_arSubItems[i]->SetPadding(sizePadding);
	}
}
//*******************************************************************************
void CBCGPRibbonButton::SetIcon (HICON hIcon, BOOL bLargeIcon/* = TRUE*/, BOOL	bAutoDestroyIcon/* = FALSE*/, BOOL bAlphaBlendIcon/* = FALSE*/)
{
	HICON& hIconThis = bLargeIcon ? m_hIcon : m_hIconSmall;

	if (m_bAutoDestroyIcon && hIconThis != NULL)
	{
		::DestroyIcon(hIconThis);
	}
	
	hIconThis = hIcon;
	m_bAutoDestroyIcon = bAutoDestroyIcon;
	m_bAlphaBlendIcon = bAlphaBlendIcon;

	if (bLargeIcon)
	{
		m_bSmallIconMode = m_hIcon == NULL;
	}
	else
	{
		m_bSmallIconMode = m_hIconSmall != NULL && m_hIcon == NULL;
	}
}
//*******************************************************************************
void CBCGPRibbonButton::SetApplicationModes(UINT nAppModes, BOOL bIncludeSubItems)
{
	ASSERT_VALID(this);
	
	CBCGPBaseRibbonElement::SetApplicationModes(nAppModes);

	if (bIncludeSubItems)
	{
		for (int i = 0; i < m_arSubItems.GetSize (); i++)
		{
			ASSERT_VALID (m_arSubItems [i]);
			m_arSubItems[i]->SetApplicationModes(nAppModes);
		}
	}
}
//*******************************************************************************
BOOL CBCGPRibbonButton::DisplayPopupDialog(CWnd* pWndParent)
{
	if (!IsPopupDialogEnabled())
	{
		return FALSE;
	}

	if (m_pPopupDlg->GetSafeHwnd() != NULL)
	{
		ClosePopupDlg(_T(""), FALSE);
	}
	else
	{
		ASSERT_VALID(pWndParent);
		ASSERT(m_lpszPopupDlgTemplateName != NULL);
		ASSERT(m_pRTIPopupDlg != NULL);
		
		CBCGPDlgPopupMenu* pPopupMenu = new CBCGPDlgPopupMenu(this, m_pRTIPopupDlg, m_lpszPopupDlgTemplateName);
		if (pPopupMenu != NULL)
		{
			pPopupMenu->m_bIsResizable = m_bIsResizablePopup;
			pPopupMenu->SetRightAlign(m_bIsRightAlignedPopup);
			
			m_pPopupDlg = pPopupMenu->m_pDlg;
			if (m_pPopupDlg != NULL)
			{
				m_pPopupDlg->m_pParentEdit = new CBCGPParentRibbonButtonPtr(this);
			}
			
			CRect rect = (m_rectCommand.IsRectEmpty() || m_bIsRightAlignedPopup) ? m_rect : m_rectCommand;
			pWndParent->ClientToScreen (&rect);

			OnBeforeShowPopupDlg(m_pPopupDlg);

			CWnd* pWndEmbedded = GetEmbeddedWnd();
			
			pPopupMenu->Create(pWndEmbedded->GetSafeHwnd() == NULL ? pWndParent : pWndEmbedded, 
				m_bIsRightAlignedPopup ? rect.right - 1 : rect.left, rect.bottom, NULL, FALSE, FALSE);
			
			if (pWndParent->GetExStyle() & WS_EX_LAYOUTRTL)
			{
				pPopupMenu->ModifyStyleEx(0, WS_EX_LAYOUTRTL);
			}
		}
	}

	return TRUE;
}
//********************************************************************************
void CBCGPRibbonButton::EnablePopupDialog(CRuntimeClass* pRTI, UINT nIDTemplate, BOOL bIsResizable, BOOL bIsDefaultCommand, BOOL bIsRightAligned)
{
	EnablePopupDialog(pRTI, MAKEINTRESOURCE(nIDTemplate), bIsResizable, bIsDefaultCommand, bIsRightAligned);
}
//********************************************************************************
void CBCGPRibbonButton::EnablePopupDialog(CRuntimeClass* pRTI, LPCTSTR lpszTemplateName, BOOL bIsResizable, BOOL bIsDefaultCommand, BOOL bIsRightAligned)
{
	ASSERT(pRTI != NULL);

	if (pRTI != NULL && pRTI->m_pfnCreateObject == NULL)
	{
		TRACE(_T("CBCGPRibbonButton::EnablePopupDialog: you've to add DECLARE_DYNCREATE to your popup dialog class\n"));
		ASSERT(FALSE);
		
		pRTI = NULL;
	}

	m_pRTIPopupDlg = pRTI;
	m_lpszPopupDlgTemplateName = lpszTemplateName;
	m_bIsResizablePopup = bIsResizable;
	m_bIsRightAlignedPopup = bIsRightAligned;
	m_bIsDefaultCommand = bIsDefaultCommand;
}
//*****************************************************************************
void CBCGPRibbonButton::ClosePopupDlg(LPCTSTR lpszEditValue, BOOL bOK, DWORD_PTR /*dwUserData*/)
{
	if (m_pPopupDlg->GetSafeHwnd() != NULL)
	{
		if (m_pPopupDlg->GetParent() != NULL)
		{
			m_pPopupDlg->GetParent()->PostMessage(WM_CLOSE);
		}
		
		m_pPopupDlg = NULL;
	}
	
	if (bOK)
	{
		if (lpszEditValue != NULL)
		{
			OnSetPopupDlgText(lpszEditValue);
		}
	}

	OnAfterClosePopupDlg(bOK);
	Redraw();
}
//*******************************************************************************
BOOL CBCGPRibbonButton::IsHiddenInAppMode() const
{
	if (CBCGPBaseRibbonElement::IsHiddenInAppMode())
	{
		return TRUE;
	}

	if (m_arSubItems.GetSize() > 0 && !m_bIsDefaultCommand)
	{
		for (int i = 0; i < m_arSubItems.GetSize (); i++)
		{
			CBCGPBaseRibbonElement* pSubItem = m_arSubItems[i];
			ASSERT_VALID(pSubItem);

			if (!pSubItem->IsHiddenInAppMode() && !pSubItem->IsSeparator() /* Ignore separators */)
			{
				// At least one subitem is visible
				return FALSE;
			}
		}

		// All subitems are invisible
		return TRUE;
	}

	return FALSE;
}
//*******************************************************************************
void CBCGPRibbonButton::OnAfterClosePopupDlg(BOOL bOK)
{
	ASSERT_VALID (this);

	if (bOK)
	{
		NotifyCommand(TRUE);
	}
	
	if (m_pParentMenu != NULL)
	{
		ASSERT_VALID (m_pParentMenu);
		
		if (m_pParentMenu->IsFloaty ())
		{
			CBCGPRibbonFloaty* pFloaty = DYNAMIC_DOWNCAST(CBCGPRibbonFloaty, m_pParentMenu->GetParent());
			if (pFloaty != NULL && (!pFloaty->IsContextMenuMode () || !bOK))
			{
				m_bIsHighlighted = FALSE;
				Redraw ();
				return;
			}
		}
		
		CFrameWnd* pParentFrame = BCGPGetParentFrame(m_pParentMenu);
		ASSERT_VALID(pParentFrame);
		
		pParentFrame->PostMessage(WM_CLOSE);
	}
	else
	{
		Redraw();
	}
}
//*******************************************************************************
void CBCGPRibbonButton::SetDontCloseParentPopupOnClick(BOOL bSet/* = TRUE*/, BOOL bIncludeSubItems/* = TRUE*/)
{
	m_bDontCloseParentPopupOnClick = bSet;

	if (bIncludeSubItems)
	{
		for (int i = 0; i < m_arSubItems.GetSize (); i++)
		{
			CBCGPRibbonButton* pSubButton = DYNAMIC_DOWNCAST(CBCGPRibbonButton, m_arSubItems [i]);
			if (pSubButton != NULL)
			{
				ASSERT_VALID(pSubButton);
				pSubButton->SetDontCloseParentPopupOnClick(bSet);
			}
		}
	}
}
//*******************************************************************************
HICON CBCGPRibbonButton::CreateSimplifiedIcon()
{
	CBCGPRibbonButton* pOrigButton = DYNAMIC_DOWNCAST(CBCGPRibbonButton, m_pOriginal);
	
	if (pOrigButton != NULL)
	{
		return pOrigButton->CreateSimplifiedIcon();
	}
	
	HICON hIcon = m_hIconSmall == NULL ? m_hIcon : m_hIconSmall;
	if (hIcon != NULL)
	{
		return CBCGPToolBarImages::SimplifyIcon(hIcon, FALSE);
	}

	if (m_nSmallImageIndex < 0)
	{
		return NULL;
	}

	if (m_pParentGroup != NULL)
	{
		ASSERT_VALID(m_pParentGroup);
		
		if (m_pParentGroup->HasImages())
		{
			return m_pParentGroup->GetImages().ExtractSimplifiedIcon(m_nSmallImageIndex);
		}
	}

	if (m_pParent != NULL)
	{
		ASSERT_VALID(m_pParent);
		return m_pParent->GetSmallImages().ExtractSimplifiedIcon(m_nSmallImageIndex);
	}

	return NULL;
}
//*******************************************************************************
BOOL CBCGPRibbonButton::IsDrawSimplifiedIcon()
{
	if (m_bIsDontDrawSimplifiedImage)
	{
		return FALSE;
	}

	if (m_bQuickAccessMode && (m_pParentMenu == NULL || m_pParentMenu->IsRibbonPopupMode()))
	{
		CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar();
		if (pTopLevelRibbon == NULL)
		{
			return FALSE;
		}

		switch (pTopLevelRibbon->GetQuickAccessIconsStyle())
		{
		case BCGPRibbonIconStyle_Default:
			return CBCGPVisualManager::GetInstance()->IsSimplifiedRibbonQATIcon(this);

		case BCGPRibbonIconStyle_Simplified:
			return TRUE;
		}

		return FALSE;
	}
	
	if (m_bIsTabElement)
	{
		CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar();
		if (pTopLevelRibbon == NULL)
		{
			return FALSE;
		}

		switch (pTopLevelRibbon->GetTabIconsStyle())
		{
		case BCGPRibbonIconStyle_Default:
			return CBCGPVisualManager::GetInstance()->IsSimplifiedRibbonTabControlIcon(this);
			
		case BCGPRibbonIconStyle_Simplified:
			return TRUE;
		}
		
		return FALSE;
	}

	if (IsCaptionButton())
	{
		CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar();
		if (pTopLevelRibbon == NULL)
		{
			return FALSE;
		}
	
		BOOL bIsSimplified = FALSE;

		switch (pTopLevelRibbon->GetCaptionButtonIconStyle())
		{
		case BCGPRibbonIconStyle_Default:
			bIsSimplified = CBCGPVisualManager::GetInstance()->IsSimplifiedRibbonCaptionButtonIcon(this);
			break;
			
		case BCGPRibbonIconStyle_Simplified:
			bIsSimplified = TRUE;
			break;
		}

		if (bIsSimplified && pTopLevelRibbon->IsBackstageViewActive())
		{
			if (CBCGPVisualManager::GetInstance()->IsRibbonBackstageWhiteBackground() || CBCGPDrawManager::IsLightColor(globalData.clrBarFace))
			{
				bIsSimplified = FALSE;
			}
		}
		
		return bIsSimplified;
	}

	if (m_bIsBackstageViewMode)
	{
		CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar();
		if (pTopLevelRibbon == NULL || !pTopLevelRibbon->IsBackstageViewActive() || !pTopLevelRibbon->IsDisplayBackstageCommandIcons() ||
			globalData.IsHighContastMode())
		{
			return FALSE;
		}
		
		switch (pTopLevelRibbon->GetBackstageViewIconsStyle())
		{
		case BCGPRibbonIconStyle_Default:
			return CBCGPVisualManager::GetInstance()->IsSimplifiedBackstageViewIcon(this);
			
		case BCGPRibbonIconStyle_Simplified:
			return TRUE;
		}
		
		return FALSE;
	}


	return FALSE;
}
//*************************************************************************************************************************
void CBCGPRibbonButton::SetCommandSearchMenu(BOOL bSet /*= TRUE*/ )
{
	CBCGPBaseRibbonElement::SetCommandSearchMenu(bSet);

	for (int i = 0; i < m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_arSubItems [i]);
		m_arSubItems[i]->SetCommandSearchMenu(bSet);
	}
}

#endif // BCGP_EXCLUDE_RIBBON
