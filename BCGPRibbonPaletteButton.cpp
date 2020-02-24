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
// BCGPRibbonPaletteButton.cpp: implementation of the CBCGPRibbonPaletteButton class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "bcgpmath.h"
#include "BCGPRibbonPaletteButton.h"
#include "BCGPRibbonCategory.h"
#include "BCGGlobals.h"
#include "MenuImages.h"
#include "BCGPVisualManager.h"
#include "BCGPRibbonPanelMenu.h"
#include "BCGPRibbonBar.h"
#include "BCGPRibbonLabel.h"
#include "BCGPToolbarMenuButton.h"
#include "BCGPFrameWnd.h"
#include "BCGPMDIFrameWnd.h"
#include "BCGPLocalResource.h"
#include "bcgprores.h"
#include "BCGPRibbonElementHostCtrl.h"

#ifndef BCGP_EXCLUDE_RIBBON

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

const int nScrollUpID		= -1;
const int nScrollDownID		= -2;
const int nMenuID			= -3;
const int nImageMargin		= 4;
const int nBorderMarginX	= 1;
const int nBorderMarginY	= 3;

////////////////////////////////////////////
// CBCGPRibbonPaletteIcon

IMPLEMENT_DYNCREATE(CBCGPRibbonPaletteIcon, CBCGPRibbonButton)

CBCGPRibbonPaletteIcon::CBCGPRibbonPaletteIcon (
		CBCGPRibbonPaletteButton* pOwner, int nIndex) :
		m_pOwner (pOwner),
		m_nIndex (nIndex)
{
	if (m_pOwner != NULL)
	{
		m_pParent = m_pOwner->m_pParent;
	}

	m_bIsFirstInRow = FALSE;
	m_bIsLastInRow = FALSE;
	m_bIsFirstInColumn = FALSE;
	m_bIsLastInColumn = FALSE;
	m_nImageIndex = m_nIndex;
	m_nCheckState = FALSE;

	m_rectCheckBox.SetRectEmpty();
}
//**************************************************************************
void CBCGPRibbonPaletteIcon::SetCheckRect()
{
	m_rectCheckBox.SetRectEmpty();

	if (m_pOwner == NULL)
	{
		return;
	}

	ASSERT_VALID(m_pOwner);

	if (!m_pOwner->IsItemCheckBoxesEnabled())
	{
		return;
	}

	const CSize sizeCheckBox = CBCGPVisualManager::GetInstance ()->GetCheckRadioDefaultSize();

	if (2 * sizeCheckBox.cx + nImageMargin > m_rect.Width() || 2 * sizeCheckBox.cy + nImageMargin > m_rect.Height())
	{
		return;
	}

	m_rectCheckBox = m_rect;
	m_rectCheckBox.DeflateRect(nImageMargin + 1, nImageMargin + 1);

	switch (m_pOwner->GetCheckBoxLocation())
	{
	default:
	case CBCGPRibbonPaletteButton::RibbonPalleteCheckbox_BottomRight:
		m_rectCheckBox.left = m_rectCheckBox.right - sizeCheckBox.cx;
		m_rectCheckBox.top = m_rectCheckBox.bottom - sizeCheckBox.cy;
		m_rectCheckBox.OffsetRect(2, 0);
		break;

	case CBCGPRibbonPaletteButton::RibbonPalleteCheckbox_TopLeft:
		m_rectCheckBox.right = m_rectCheckBox.left + sizeCheckBox.cx;
		m_rectCheckBox.bottom = m_rectCheckBox.top + sizeCheckBox.cy;
		m_rectCheckBox.OffsetRect(-2, 0);
		break;

	case CBCGPRibbonPaletteButton::RibbonPalleteCheckbox_BottomLeft:
		m_rectCheckBox.right = m_rectCheckBox.left + sizeCheckBox.cx;
		m_rectCheckBox.top = m_rectCheckBox.bottom - sizeCheckBox.cy;
		m_rectCheckBox.OffsetRect(-2, 0);
		break;

	case CBCGPRibbonPaletteButton::RibbonPalleteCheckbox_TopRight:
		m_rectCheckBox.left = m_rectCheckBox.right - sizeCheckBox.cx;
		m_rectCheckBox.bottom = m_rectCheckBox.top + sizeCheckBox.cy;
		m_rectCheckBox.OffsetRect(2, 0);
		break;
	}
}
//**************************************************************************
void CBCGPRibbonPaletteIcon::OnDraw (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);
	ASSERT_VALID (m_pOwner);

	if (m_rect.IsRectEmpty ())
	{
		return;
	}

	BOOL bSkipBorder = FALSE;

	if (m_nIndex >= 0)
	{
		BOOL bIsMenuSideBar = m_pOwner->IsMenuSideBar() && m_pOwner->IsComboMode() && m_pParentMenu != NULL;
		int cxGutter = 0;

		if (bIsMenuSideBar)
		{
			BOOL bDisableSideBarInXPMode = m_pParentMenu->m_bDisableSideBarInXPMode;
			m_pParentMenu->m_bDisableSideBarInXPMode = FALSE;

			cxGutter = m_pParentMenu->GetGutterWidth();
			if (cxGutter < 1)
			{
				bIsMenuSideBar = FALSE;
			}

			m_pParentMenu->m_bDisableSideBarInXPMode = bDisableSideBarInXPMode;
		}

		if (!m_pOwner->IsDisabled () || m_pOwner->IsDrawDisabledItems())
		{
			COLORREF clrText = (COLORREF)-1;

			if (m_pOwner->m_bDefaultButtonStyle || !m_pOwner->m_bIsOwnerDraw)
			{
				BOOL bIsChecked = m_bIsChecked;
				if (m_bIsDisabled || (!m_pOwner->IsFocused() && !m_pOwner->m_bDrawSelectionAlways) || bIsMenuSideBar)
				{
					m_bIsChecked = FALSE;
				}

				if (!bIsMenuSideBar && m_bIsChecked && m_pOwner->IsComboMode() && m_pOwner->IsDrawCheckFrameAroundImage())
				{
					m_bIsChecked = FALSE;
					
					clrText = OnFillBackground(pDC);

					m_bIsChecked = TRUE;

					CRect rect = m_rect;
					m_rect.right = m_rect.left + m_pOwner->GetItemSize().cx + 2 * globalUtils.ScaleByDPI(m_szMargin.cx) + 2;

					OnFillBackground(pDC);

					if (!m_bIsHighlighted)
					{
						OnDrawBorder(pDC);
					}

					m_rect = rect;
					bSkipBorder = TRUE;
				}
				else
				{
					clrText = OnFillBackground (pDC);
				}

				m_bIsChecked = bIsChecked;
			}

			CRect rectIcon = m_rect;
			if (m_pOwner->m_bItemCheckBoxes && !m_pOwner->m_bCheckBoxOverlapsIcon)
			{
				int cxCheckBox = m_rectCheckBox.Width();

				switch (m_pOwner->GetCheckBoxLocation())
				{
				default:
				case CBCGPRibbonPaletteButton::RibbonPalleteCheckbox_BottomRight:
				case CBCGPRibbonPaletteButton::RibbonPalleteCheckbox_TopRight:
					rectIcon.right -= cxCheckBox;
					break;
					
				case CBCGPRibbonPaletteButton::RibbonPalleteCheckbox_TopLeft:
				case CBCGPRibbonPaletteButton::RibbonPalleteCheckbox_BottomLeft:
					rectIcon.OffsetRect(cxCheckBox, 0);
					break;
				}
			}

			if (bIsMenuSideBar)
			{
				if (m_bIsChecked)
				{
					CRect rectCheckBox = rectIcon;
					rectCheckBox.right = rectCheckBox.left + cxGutter;

					rectCheckBox.DeflateRect(globalUtils.ScaleByDPI(2), globalUtils.ScaleByDPI(2));

					CBCGPVisualManager::GetInstance ()->OnDrawRibbonMenuCheckFrame (pDC, this, rectCheckBox);

					CBCGPToolbarMenuButton dummy;
					CBCGPVisualManager::GetInstance ()->OnDrawMenuCheck (pDC, &dummy, rectCheckBox, m_bIsHighlighted, FALSE);
				}

				rectIcon.left += cxGutter;
			}

			m_pOwner->OnDrawPaletteIcon (pDC, rectIcon, m_nImageIndex, this, clrText);

			if (!m_rectCheckBox.IsRectEmpty())
			{
				CBCGPVisualManager::GetInstance ()->OnDrawCheckBoxEx (pDC,
					m_rectCheckBox, m_nCheckState, FALSE, FALSE, !m_pOwner->IsDisabled ());
			}

			if (m_pOwner->m_bItemTextLabels && !m_pOwner->m_bIsOwnerDraw)
			{
				CString strLabel = GetTextLabel();
				if (!strLabel.IsEmpty())
				{
					CRect rectText = rectIcon;
					rectText.right = min(m_rect.right, rectText.right);

					rectText.DeflateRect(globalUtils.ScaleByDPI(nImageMargin), globalUtils.ScaleByDPI(nImageMargin));

					UINT dtFlags = DT_SINGLELINE | DT_CENTER | DT_END_ELLIPSIS | DT_NOPREFIX;

					switch (m_pOwner->m_TextLabelLocation)
					{
					case CBCGPRibbonPaletteButton::RibbonPalleteTextLabel_Top:
						break;

					case CBCGPRibbonPaletteButton::RibbonPalleteTextLabel_Center:
						dtFlags |= DT_VCENTER;
						break;

					case CBCGPRibbonPaletteButton::RibbonPalleteTextLabel_Bottom:
						rectText.top = rectText.bottom - pDC->GetTextExtent(strLabel).cy - 1;
						break;
					}

					COLORREF clrTextOld = (COLORREF)-1;
					if (m_pOwner->m_clrTextLabel != (COLORREF)-1)
					{
						clrText = m_pOwner->m_clrTextLabel;
					}

					if (clrText != (COLORREF)-1)
					{
						clrTextOld = pDC->SetTextColor(clrText);
					}

					pDC->DrawText(strLabel, rectText, dtFlags);

					if (clrTextOld != (COLORREF)-1)
					{
						pDC->SetTextColor(clrTextOld);
					}
				}
			}

			if (!bSkipBorder && (m_pOwner->m_bDefaultButtonStyle || !m_pOwner->m_bIsOwnerDraw))
			{
				OnDrawBorder (pDC);
			}
		}
	}
	else
	{
		CBCGPVisualManager::GetInstance ()->OnDrawRibbonPaletteButton(pDC, this);

		// Draw scroll/menu button:
		CBCGPMenuImages::IMAGES_IDS id = 
			m_nIndex == nScrollUpID ? CBCGPMenuImages::IdArowUp :
			m_nIndex == nScrollDownID ? CBCGPMenuImages::IdArowDown : CBCGPMenuImages::IdCustomizeArowDown;

		CRect rectImage = m_rect;

		if (m_nIndex == nMenuID && rectImage.Height () > rectImage.Width () + 2)
		{
			rectImage.bottom = rectImage.top + rectImage.Width () + 2;
		}

		CBCGPVisualManager::GetInstance ()->OnDrawRibbonPaletteButtonIcon(pDC, this, (int)id, rectImage);
	}
}
//***************************************************************************
void CBCGPRibbonPaletteIcon::OnClick (CPoint point)
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pOwner);

	if (!m_pOwner->PreClickPaletteIcon(this, point))
	{
		return;
	}

	if (m_rectCheckBox.PtInRect(point))
	{
		SetCheck(!m_nCheckState);
		m_pOwner->OnClickPaletteIconCheckbox(m_pOriginal == NULL ? this : (CBCGPRibbonPaletteIcon*) m_pOriginal);
		return;
	}

	m_pOwner->OnClickPaletteIcon (m_pOriginal == NULL ? this : (CBCGPRibbonPaletteIcon*) m_pOriginal);

	if (m_nIndex < 0)
	{
		return;
	}

	CBCGPRibbonPanelMenuBar* pParentMenu = m_pParentMenu;
	if (pParentMenu == NULL && m_nIndex >= 0)
	{
		pParentMenu = m_pOwner->m_pParentMenu;
	}

	if (pParentMenu != NULL)
	{
		ASSERT_VALID (pParentMenu);

		if (m_pOwner->m_nPaletteID != 0)
		{
			m_pOwner->SetNotifyParentID (TRUE);
		}

		m_pOwner->m_bIsFocused = FALSE;
		m_pOwner->OnSetFocus (FALSE);

		pParentMenu->OnClickButton (m_pOwner, point);
	}
	else if (m_nIndex >= 0)
	{
		m_pOwner->NotifyCommand ();
	}
}
//*****************************************************************************
void CBCGPRibbonPaletteIcon::OnLButtonDown (CPoint point)
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pOwner);

	if (m_nIndex != nMenuID)
	{
		CBCGPRibbonButton::OnLButtonDown (point);
		return;
	}

	m_bIsHighlighted = m_bIsPressed = FALSE;
	Redraw ();

	m_pOwner->OnShowPopupMenu ();
}
//*****************************************************************************
void CBCGPRibbonPaletteIcon::CopyFrom (const CBCGPBaseRibbonElement& s)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::CopyFrom (s);

	CBCGPRibbonPaletteIcon& src = (CBCGPRibbonPaletteIcon&) s;

	m_nIndex = src.m_nIndex;
	m_nImageIndex = src.m_nImageIndex;
	m_pOwner = src.m_pOwner;
	m_bIsChecked = src.m_bIsChecked;
	m_nCheckState = src.m_nCheckState;
}
//***************************************************************************
CSize CBCGPRibbonPaletteIcon::GetCompactSize (CDC* pDC)
{
	return GetRegularSize (pDC);
}
//***************************************************************************
CSize CBCGPRibbonPaletteIcon::GetIntermediateSize (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pOwner);

	return GetRegularSize (pDC);
}
//***************************************************************************
CSize CBCGPRibbonPaletteIcon::GetRegularSize (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pOwner);

	CSize sizeItem = m_pOwner->GetItemSize();

	if (!m_pOwner->m_bSmallIcons)
	{
		sizeItem.cx += 2 * nImageMargin;
		sizeItem.cy += 2 * nImageMargin;
	}

	if (m_pOwner->m_bIsComboMode)
	{
		CSize sizeText = m_pOwner->GetComboItemTextSize(pDC, this);

		sizeItem.cx += sizeText.cx + 2 * m_szMargin.cx;
		sizeItem.cy = max(sizeItem.cy, sizeText.cy + 2 * m_szMargin.cy);
		
		if (m_pOwner->IsMenuSideBar())
		{
			sizeItem.cx += CBCGPToolBar::GetMenuImageSize ().cx + 2 * CBCGPVisualManager::GetInstance ()->GetMenuImageMargin() + 2;
		}
	}

	return sizeItem;
}
//***************************************************************************
BOOL CBCGPRibbonPaletteIcon::IsFirst () const
{
	ASSERT_VALID (this);
	return m_nIndex == nScrollUpID;
}
//***************************************************************************
BOOL CBCGPRibbonPaletteIcon::IsLast () const
{
	ASSERT_VALID (this);
	return m_nIndex == nMenuID;
}
//***************************************************************************
BOOL CBCGPRibbonPaletteIcon::IsAutoRepeatMode (int& /*nDelay*/) const
{
	ASSERT_VALID (this);
	return m_nIndex == nScrollUpID || m_nIndex == nScrollDownID;
}
//***************************************************************************
BOOL CBCGPRibbonPaletteIcon::OnAutoRepeat ()
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pOwner);

	if (m_bIsDisabled)
	{
		return FALSE;
	}

	m_pOwner->OnClickPaletteIcon (this);
	return TRUE;
}
//*************************************************************************************
BOOL CBCGPRibbonPaletteIcon::OnAddToQAToolbar (CBCGPRibbonQuickAccessToolbar& qat)
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pOwner);

	if (m_pOwner->IsCustom())
	{
		CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar();
		if (pRibbonBar != NULL)
		{
			CBCGPBaseRibbonElement* pOriginalElem = pRibbonBar->FindByID(m_pOwner->GetID (), FALSE);
			if (pOriginalElem != NULL)
			{
				ASSERT_VALID(pOriginalElem);
				pOriginalElem->OnAddToQAToolbar (qat);
				return TRUE;
			}
		}

		return FALSE;
	}

	m_pOwner->OnAddToQAToolbar (qat);
	return TRUE;
}
//*************************************************************************************
void CBCGPRibbonPaletteIcon::OnHighlight (BOOL bHighlight)
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pOwner);

	CBCGPRibbonButton::OnHighlight (bHighlight);

	if (!bHighlight)
	{
		CPoint ptCursor;
		::GetCursorPos (&ptCursor);

		CBCGPRibbonPaletteIcon* pCurrIcon = NULL;

		if (m_pParentMenu != NULL)
		{
			m_pParentMenu->ScreenToClient (&ptCursor);

			CBCGPRibbonPanel* pPanel = GetParentPanel ();

			if (pPanel != NULL)
			{
				pCurrIcon = DYNAMIC_DOWNCAST (
					CBCGPRibbonPaletteIcon, pPanel->HitTest (ptCursor));
			}
		}
		else
		{
			m_pOwner->GetParentWnd ()->ScreenToClient (&ptCursor);

			pCurrIcon = DYNAMIC_DOWNCAST (
				CBCGPRibbonPaletteIcon, m_pOwner->HitTest (ptCursor));
		}

		if (pCurrIcon != NULL && pCurrIcon->m_nIndex >= 0)
		{
			return;
		}
	}

	if (m_nIndex >= 0)
	{
		m_pOwner->NotifyHighlightListItem (bHighlight ? m_nIndex : -1);
	}
}
//*************************************************************************************
void CBCGPRibbonPaletteIcon::SetCheck(int nState, BOOL bRedraw)
{
	m_nCheckState = nState;

	CBCGPRibbonPaletteIcon* pOriginalIcon = DYNAMIC_DOWNCAST(CBCGPRibbonPaletteIcon, m_pOriginal);
	if (pOriginalIcon != NULL)
	{
		pOriginalIcon->SetCheck(nState, bRedraw);
	}

	if (bRedraw)
	{
		Redraw();
	}

	if (m_pOwner == NULL)
	{
		return;
	}

	CBCGPRibbonBar* pRibbonBar = m_pOwner->GetTopLevelRibbonBar();
	if (!m_pOwner->m_bDontNotify && pRibbonBar != NULL)
	{
		ASSERT_VALID(pRibbonBar);
		
		CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arButtons;
		pRibbonBar->GetElementsByID(m_pOwner->m_nID, arButtons, TRUE);
		
		for (int i = 0; i < arButtons.GetSize(); i++)
		{
			CBCGPRibbonPaletteButton* pOther = DYNAMIC_DOWNCAST(CBCGPRibbonPaletteButton, arButtons[i]);
			
			if (pOther != NULL && pOther != m_pOwner)
			{
				ASSERT_VALID (pOther);
				
				pOther->m_bDontNotify = TRUE;
				pOther->SetCheckState(m_nIndex, nState);
				pOther->m_bDontNotify = FALSE;
			}
		}
	}
}
//*************************************************************************************
BOOL CBCGPRibbonPaletteIcon::OnProcessKey (UINT nChar)
{
	ASSERT_VALID (this);

	if (nChar == VK_SPACE && !m_rectCheckBox.IsRectEmpty())
	{
		SetCheck(!m_nCheckState);
		return TRUE;
	}

	return CBCGPRibbonButton::OnProcessKey(nChar);
}
//*************************************************************************************
CRect CBCGPRibbonPaletteIcon::GetKeyTipRect(CDC* pDC, BOOL /*bIsMenu*/)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);
	ASSERT_VALID (m_pOwner);
	
	CSize sizeKeyTip = GetKeyTipSize (pDC);
	CRect rectKeyTip (0, 0, 0, 0);
	
	if (sizeKeyTip == CSize (0, 0) || m_rect.IsRectEmpty ())
	{
		return rectKeyTip;
	}
	
	rectKeyTip.left = m_rect.left + sizeKeyTip.cy / 2;
	rectKeyTip.top = m_rect.CenterPoint().y;
	rectKeyTip.right = rectKeyTip.left + sizeKeyTip.cx;
	rectKeyTip.bottom = rectKeyTip.top + sizeKeyTip.cy;

	CWnd* pWndParent = GetParentWnd();

	if (m_pParentMenu->GetSafeHwnd() == NULL || pWndParent->GetSafeHwnd() == NULL)
	{
		CRect rectOwner = m_pOwner->GetRect();
		rectKeyTip.OffsetRect(rectOwner.TopLeft());

		if (m_pOwner->IsButtonLook())
		{
			rectKeyTip.OffsetRect(0, rectOwner.Height());
		}
	}
	else
	{
		ASSERT_VALID(m_pParentMenu);
		ASSERT_VALID(pWndParent);

		CRect rectMenu;
		m_pParentMenu->GetWindowRect(rectMenu);

		pWndParent->ScreenToClient(rectMenu);
		rectKeyTip.OffsetRect(rectMenu.TopLeft());
	}

	return rectKeyTip;
}
//*************************************************************************************
CString CBCGPRibbonPaletteIcon::GetTextLabel() const
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pOwner);

	return m_pOwner->GetIconTextLabel(this);
}
//*************************************************************************************
CString CBCGPRibbonPaletteIcon::GetToolTipText () const
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pOwner);

	return m_pOwner->GetIconToolTip (this);
}
//*************************************************************************************
CString CBCGPRibbonPaletteIcon::GetDescription () const
{
	ASSERT_VALID (this);
	ASSERT_VALID (m_pOwner);

	return m_pOwner->GetIconDescription (this);
}
//*************************************************************************************
BOOL CBCGPRibbonPaletteIcon::SetACCData(CWnd* pParent, CBCGPAccessibilityData& data)
{
	CBCGPRibbonButton::SetACCData(pParent, data);

	switch (m_nIndex)
	{
	case nMenuID:
		data.m_nAccRole = ROLE_SYSTEM_BUTTONDROPDOWNGRID;

		data.m_bAccState |= STATE_SYSTEM_HASPOPUP;
		data.m_strAccDefAction = _T("Open");
		data.m_strAccName = GetToolTipText();

		if (IsDroppedDown())
		{
			data.m_bAccState |= STATE_SYSTEM_PRESSED;
			data.m_strAccDefAction = _T("Close");
		}

	case nScrollUpID:
	case nScrollDownID:
		if (m_nIndex == nMenuID)
		{
			if (m_pOwner != NULL)
			{
				ASSERT_VALID(m_pOwner);
				data.m_strAccName = m_pOwner->GetText();
			}
		}
		else
		{
			data.m_strAccName.LoadString(m_nIndex == nScrollUpID ? IDS_BCGBARRES_ROW_UP : IDS_BCGBARRES_ROW_DOWN);
		}

		data.m_strAccValue = GetToolTipText();
		break;

	default:
		{
			data.m_bAccState = STATE_SYSTEM_FOCUSABLE | STATE_SYSTEM_SELECTABLE;

			if (IsHighlighted())
			{
				 data.m_bAccState |= STATE_SYSTEM_SELECTED | STATE_SYSTEM_FOCUSED;
			}

			if (IsChecked() || m_nCheckState != 0)
			{
				data.m_bAccState |= STATE_SYSTEM_CHECKED;
			}
		
			data.m_strAccName = GetToolTipText();
			data.m_nAccRole = ROLE_SYSTEM_LISTITEM;
			data.m_strAccDefAction = _T("DoubleClick");

		}
		break;
	}

	return TRUE;
}

/////////////////////////////////////////////////////////////////////
// CBCGPRibbonPaletteButton

IMPLEMENT_DYNCREATE(CBCGPRibbonPaletteButton, CBCGPRibbonButton)

CMap<UINT,UINT,int,int>	CBCGPRibbonPaletteButton::m_mapSelectedItems;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPRibbonPaletteButton::CBCGPRibbonPaletteButton()
{
	CommonInit ();
}
//***************************************************************************
CBCGPRibbonPaletteButton::CBCGPRibbonPaletteButton(
		UINT				nID,
		LPCTSTR				lpszText, 
		int					nSmallImageIndex,
		int					nLargeImageIndex,
		CBCGPToolBarImages& imagesPalette) :
	CBCGPRibbonButton (nID, lpszText, nSmallImageIndex, nLargeImageIndex)
{
	CommonInit ();
	imagesPalette.CopyTo (m_imagesPalette);

	m_nIcons = m_imagesPalette.GetCount ();

	CreateIcons ();
}
//***************************************************************************
CBCGPRibbonPaletteButton::CBCGPRibbonPaletteButton(
		UINT				nID,
		LPCTSTR				lpszText, 
		int					nSmallImageIndex,
		int					nLargeImageIndex,
		UINT				uiImagesPaletteResID,
		int					cxPaletteImage) :
	CBCGPRibbonButton (nID, lpszText, nSmallImageIndex, nLargeImageIndex)
{
	CommonInit ();

	if (uiImagesPaletteResID != 0)
	{
		ASSERT (cxPaletteImage != 0);

		m_imagesPalette.Load(uiImagesPaletteResID);

		int cy = cxPaletteImage;

		HBITMAP hbmp = m_imagesPalette.GetImageWell();
		if (hbmp != NULL)
		{
			BITMAP bmp;
			GetObject(hbmp, sizeof (BITMAP), &bmp);

			cy = bmp.bmHeight;
		}

		m_imagesPalette.SetImageSize(CSize (cxPaletteImage, cy), TRUE);

		const double dblScale = globalData.GetRibbonImageScale ();
		if (dblScale != 1.0)
		{
			m_imagesPalette.SetTransparentColor (globalData.clrBtnFace);
			m_imagesPalette.SmoothResize (dblScale);
		}

		m_nIcons = m_imagesPalette.GetCount ();

		CreateIcons ();
	}
}
//***************************************************************************
CBCGPRibbonPaletteButton::CBCGPRibbonPaletteButton (
		UINT				nID,
		LPCTSTR				lpszText, 
		int					nSmallImageIndex,
		int					nLargeImageIndex,
		CSize				sizeIcon,
		int					nIconsNum,
		BOOL				bDefaultButtonStyle) :
	CBCGPRibbonButton (nID, lpszText, nSmallImageIndex, nLargeImageIndex)
{
	CommonInit ();

	m_bIsOwnerDraw = TRUE;
	m_bDefaultButtonStyle = bDefaultButtonStyle;

	m_imagesPalette.SetImageSize (sizeIcon);
	m_nIcons = nIconsNum;
}
//***************************************************************************
CBCGPRibbonPaletteButton::~CBCGPRibbonPaletteButton()
{
	RemoveAll ();
}
//***************************************************************************
void CBCGPRibbonPaletteButton::AddGroup (
		LPCTSTR lpszGroupName,
		UINT	uiImagesPaletteResID,
		int		cxPaletteImage)
{
	ASSERT_VALID (this);
	ASSERT (lpszGroupName != NULL);

	if (m_bIsOwnerDraw)
	{
		ASSERT (FALSE);
		return;
	}

	CBCGPToolBarImages imagesGroup;
	imagesGroup.Load (uiImagesPaletteResID);

	int cy = cxPaletteImage;
	
	HBITMAP hbmp = imagesGroup.GetImageWell();
	if (hbmp != NULL)
	{
		BITMAP bmp;
		GetObject(hbmp, sizeof (BITMAP), &bmp);
		
		cy = bmp.bmHeight;
	}

	imagesGroup.SetImageSize (CSize (cxPaletteImage, cy), TRUE);

	const double dblScale = globalData.GetRibbonImageScale ();
	if (dblScale != 1.0)
	{
		imagesGroup.SetTransparentColor (globalData.clrBtnFace);
		imagesGroup.SmoothResize (dblScale);
	}

	AddGroup (lpszGroupName, imagesGroup);
}
//***************************************************************************
void CBCGPRibbonPaletteButton::AddGroup (
		LPCTSTR lpszGroupName,
		CBCGPToolBarImages& imagesGroup)
{
	ASSERT_VALID (this);

	if (m_bIsOwnerDraw)
	{
		ASSERT (FALSE);
		return;
	}

	m_arGroupNames.Add (lpszGroupName);
	m_arGroupLen.Add (m_imagesPalette.GetCount ());

	if (m_imagesPalette.GetCount () == 0)
	{
		imagesGroup.CopyTo (m_imagesPalette);
	}
	else
	{
		ASSERT (CSize (imagesGroup.GetImageSize ()) == m_imagesPalette.GetImageSize ());
		m_imagesPalette.AddImageList(imagesGroup);
	}

	m_nIcons = m_imagesPalette.GetCount ();
	RemoveAll ();
}
//***************************************************************************
void CBCGPRibbonPaletteButton::AddGroup (
		LPCTSTR lpszGroupName,
		int	nIconsNum)
{
	ASSERT_VALID (this);

	if (!m_bIsOwnerDraw)
	{
		ASSERT (FALSE);
		return;
	}

	m_arGroupNames.Add (lpszGroupName);
	m_arGroupLen.Add (m_nIcons);

	m_nIcons += nIconsNum;
	RemoveAll ();
}
//***************************************************************************
void CBCGPRibbonPaletteButton::SetGroupName (int nGroupIndex, LPCTSTR lpszGroupName)
{
	ASSERT_VALID (this);

	m_arGroupNames [nGroupIndex] = lpszGroupName;

	if (m_arIcons.GetSize () == 0)
	{
		return;
	}

	CBCGPRibbonLabel* pLabel = DYNAMIC_DOWNCAST (CBCGPRibbonLabel,
		m_arIcons [m_arGroupLen [nGroupIndex]]);
	if (pLabel == NULL)
	{
		return;
	}

	ASSERT_VALID (pLabel);

	pLabel->SetText (lpszGroupName);

	CBCGPRibbonPanelMenu* pPanelMenu = DYNAMIC_DOWNCAST (CBCGPRibbonPanelMenu, m_pPopupMenu);
	if (pPanelMenu != NULL)
	{
		ASSERT_VALID (pPanelMenu);

		if (pPanelMenu->GetPanel () != NULL)
		{
			CBCGPBaseRibbonElement* pMenuElem = pPanelMenu->GetPanel ()->FindByData ((DWORD_PTR) pLabel);
			
			if (pMenuElem != NULL)
			{
				pMenuElem->SetText (lpszGroupName);
				pMenuElem->Redraw ();
			}
		}
	}
}
//***************************************************************************
LPCTSTR CBCGPRibbonPaletteButton::GetGroupName (int nGroupIndex) const
{
	ASSERT_VALID (this);
	return m_arGroupNames [nGroupIndex];
}
//***************************************************************************
void CBCGPRibbonPaletteButton::SetPalette (CBCGPToolBarImages& imagesPalette)
{
	ASSERT_VALID (this);

	if (m_bIsOwnerDraw)
	{
		ASSERT (FALSE);
		return;
	}

	Clear ();
	imagesPalette.CopyTo (m_imagesPalette);

	m_nIcons = m_imagesPalette.GetCount ();

	CreateIcons ();
}
//***************************************************************************
void CBCGPRibbonPaletteButton::SetPalette (UINT uiImagesPaletteResID, int cxPaletteImage)
{
	ASSERT_VALID (this);
	ASSERT (uiImagesPaletteResID != 0);
	ASSERT (cxPaletteImage != 0);

	if (m_bIsOwnerDraw)
	{
		ASSERT (FALSE);
		return;
	}

	Clear ();

	m_imagesPalette.Load (uiImagesPaletteResID);

	int cy = cxPaletteImage;
	
	HBITMAP hbmp = m_imagesPalette.GetImageWell();
	if (hbmp != NULL)
	{
		BITMAP bmp;
		GetObject(hbmp, sizeof (BITMAP), &bmp);
		
		cy = bmp.bmHeight;
	}

	m_imagesPalette.SetImageSize(CSize (cxPaletteImage, cy), TRUE);
	m_nIcons = m_imagesPalette.GetCount ();

	CreateIcons ();
}
//***************************************************************************
void CBCGPRibbonPaletteButton::Clear ()
{
	ASSERT_VALID (this);
	
	m_mapSelectedItems.RemoveKey (m_nPaletteID == 0 ? m_nID : m_nPaletteID);

	RemoveAll ();

	m_arGroupNames.RemoveAll ();
	m_arGroupLen.RemoveAll ();
	m_arToolTips.RemoveAll ();
	m_arUserData.RemoveAll();
	m_arKeys.RemoveAll();
	m_imagesPalette.Clear ();

	m_nScrollOffset = 0;
	m_nScrollTotal = 0;
	m_nIcons = 0;
}
//***************************************************************************
void CBCGPRibbonPaletteButton::RedrawIcons ()
{
	ASSERT_VALID (this);

	if (m_pPopupMenu != NULL &&
		m_pPopupMenu->GetMenuBar () != NULL)
	{
		m_pPopupMenu->GetMenuBar ()->RedrawWindow ();
		return;
	}

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		m_arIcons [i]->Redraw ();
	}
}
//***************************************************************************
void CBCGPRibbonPaletteButton::RemoveAll ()
{
	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		delete m_arIcons [i];
	}

	m_arIcons.RemoveAll ();
}
//*****************************************************************************
void CBCGPRibbonPaletteButton::AddSubItem (CBCGPBaseRibbonElement* pSubItem, int nIndex,
										   BOOL bOnTop)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pSubItem);

	pSubItem->m_bIsOnPaletteTop = bOnTop;

	CBCGPRibbonButton::AddSubItem (pSubItem, nIndex);
}
//***************************************************************************
void CBCGPRibbonPaletteButton::CommonInit ()
{
	m_bIsDefaultCommand = FALSE;
	m_bIsButtonMode = FALSE;
	m_bShowSelectedImage = FALSE;
	m_bShowSelectedLabel = FALSE;
	m_nImagesInRow = 0;
	m_nImagesInColumn = 0;
	m_bSmallIcons = FALSE;
	m_nScrollOffset = 0;
	m_nScrollTotal = 0;
	m_nSelected = 0;
	m_bDrawSelectionAlways = TRUE;
	m_bEnableMenuResize = FALSE;
	m_bMenuResizeVertical = FALSE;
	m_nIconsInRow = -1;
	m_nPaletteID = 0;
	m_bNotifyPaletteID = FALSE;
	m_nPanelColumns = 6;
	m_nInitialColumns = 0;
	m_bIsOwnerDraw = FALSE;
	m_bDefaultButtonStyle = TRUE;
	m_bMenuSideBar = FALSE;
	m_bIsCollapsed = FALSE;
	m_nIcons = 0;
	m_bResetColumns = FALSE;
	m_pParentControl = NULL;
	m_bIsComboMode = FALSE;
	m_bDrawCheckFrameAroundImage = FALSE;
	m_bDrawDisabledItems = FALSE;
	m_bItemCheckBoxes = FALSE;
	m_bCheckBoxOverlapsIcon = TRUE;
	m_CheckBoxLocation = RibbonPalleteCheckbox_BottomRight;
	m_bItemTextLabels = FALSE;
	m_TextLabelLocation = RibbonPalleteTextLabel_Bottom;
	m_clrTextLabel = (COLORREF)-1;
}
//***************************************************************************
CSize CBCGPRibbonPaletteButton::GetCompactSize (CDC* pDC)
{
	ASSERT_VALID (this);

	if (IsButtonLook ())
	{
		return CBCGPRibbonButton::GetCompactSize (pDC);
	}

	return CBCGPRibbonButton::GetRegularSize (pDC);
}
//***************************************************************************
CSize CBCGPRibbonPaletteButton::GetRegularSize (CDC* pDC)
{
	ASSERT_VALID (this);

	const CSize sizeImage = GetItemSize ();
	CSize sizePanelSmallImage (16, 16);

	if (m_pParent != NULL)
	{
		ASSERT_VALID (m_pParent);
		sizePanelSmallImage = m_pParent->GetImageSize (FALSE);
	}

	m_bSmallIcons = (sizeImage.cy <= sizePanelSmallImage.cy * 3 / 2) || IsAlwaysSmallIcons();

	if (m_bResetColumns && !m_bSmallIcons)
	{
		m_nPanelColumns = m_nInitialColumns > 0 ? m_nInitialColumns : 6;

		if (m_pParentMenu != NULL && m_pParentMenu->GetCategory () == NULL && !m_pParentMenu->IsRibbonPopupMode())
		{
			// From the default panel button
			m_nPanelColumns = 3;
		}
	}

	m_bResetColumns = FALSE;

	if (IsButtonLook ())
	{
		return CBCGPRibbonButton::GetRegularSize (pDC);
	}

	if (m_arIcons.GetSize () == 0)
	{
		CreateIcons ();
	}

	CSize sizePanelLargeImage (32, 32);

	if (m_pParent != NULL)
	{
		ASSERT_VALID (m_pParent);
		sizePanelLargeImage = m_pParent->GetImageSize (TRUE);
	}

	CSize size (0, 0);

	if (m_bSmallIcons)
	{
		size.cx = sizeImage.cx * m_nPanelColumns;

		int nRows = 3;

		if (sizePanelLargeImage != CSize (0, 0) && sizeImage.cy != 0)
		{
			nRows = max (nRows, sizePanelLargeImage.cy * 2 / sizeImage.cy);
		}

		size.cy = nRows * sizeImage.cy + 2 * nBorderMarginY;
	}
	else
	{
		size.cx = (sizeImage.cx + 2 * nImageMargin) * m_nPanelColumns;
		size.cy = sizeImage.cy + 3 * nImageMargin + 2 * nBorderMarginY;
	}

	//---------------------------------------
	// Add space for menu and scroll buttons:
	//---------------------------------------
	size.cx += GetDropDownImageWidth () + 3 * nImageMargin;

	return size;
}
//***************************************************************************
void CBCGPRibbonPaletteButton::OnDraw (CDC* pDC)
{
	ASSERT_VALID (this);

	if (IsButtonLook ())
	{
		CString strText = m_strText;

		if (m_bShowSelectedLabel && m_nSelected >= 0)
		{
			for (int i = 0; i < m_arIcons.GetSize (); i++)
			{
				CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST(CBCGPRibbonPaletteIcon, m_arIcons[i]);
				if (pIcon != NULL)
				{
					ASSERT_VALID (pIcon);

					if (pIcon->GetIndex() == m_nSelected)
					{
						m_strText = pIcon->GetText();
						if (m_strText.IsEmpty())
						{
							m_strText = GetIconTextLabel(pIcon);

							if (m_strText.IsEmpty())
							{
								m_strText = _T(" ");
							}
						}

						break;
					}
				}
			}
		}

		CBCGPRibbonButton::OnDraw (pDC);
		m_strText = strText;
		return;
	}

	CRect rectBorder = m_rect;
	rectBorder.DeflateRect (nBorderMarginX, nBorderMarginY);
	rectBorder.right -= 2 * nBorderMarginX;

	CBCGPVisualManager::GetInstance ()->OnDrawRibbonPaletteBorder (
		pDC, this, rectBorder);

	CRect rectImages = m_rect;
	const CSize sizeImage = GetIconSize ();

	CBCGPDrawState ds;

	if (m_imagesPalette.GetCount () > 0)
	{
		m_imagesPalette.SetTransparentColor (globalData.clrBtnFace);
		m_imagesPalette.PrepareDrawImage (ds, sizeImage);
	}

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		m_arIcons [i]->OnDraw (pDC);
	}

	if (m_imagesPalette.GetCount () > 0)
	{
		m_imagesPalette.EndDrawImage (ds);
	}
}
//***************************************************************************
void CBCGPRibbonPaletteButton::OnAfterChangeRect (CDC* pDC)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::OnAfterChangeRect (pDC);

	m_nScrollTotal = 0;
	m_nScrollOffset = 0;

	const CSize sizeItem = GetItemSize ();

	if (sizeItem.cx == 0 || sizeItem.cy == 0 || IsButtonLook ())
	{
		m_nImagesInRow = 0;
		m_nImagesInColumn = 0;

		RebuildIconLocations ();

		return;
	}

	const int cxMenu = GetDropDownImageWidth () + globalUtils.ScaleByDPI(6);

	CRect rectImages = m_rect;

	int nMargin = m_bSmallIcons ? 0 : nImageMargin;
	rectImages.DeflateRect (0, nMargin);

	rectImages.right -= cxMenu;

	m_nImagesInRow = rectImages.Width () / (sizeItem.cx + 2 * nMargin);
	m_nImagesInColumn = rectImages.Height () / (sizeItem.cy + 2 * nMargin);

	if (m_nImagesInRow == 0)
	{
		m_nScrollTotal = 0;
	}
	else
	{
		m_nScrollTotal = m_nIcons / m_nImagesInRow - m_nImagesInColumn;

		if (m_nIcons % m_nImagesInRow)
		{
			m_nScrollTotal++;
		}
	}

	RebuildIconLocations ();

	CRect rectBorder = m_rect;
	rectBorder.DeflateRect (nBorderMarginX, nBorderMarginY);
	rectBorder.right -= 2 * nBorderMarginX;

	const int cyMenu = rectBorder.Height () / 3;

	int yButton = rectBorder.top;

	CRect rectButtons = rectBorder;
	rectButtons.left = rectButtons.right - cxMenu;

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST (
			CBCGPRibbonPaletteIcon, m_arIcons [i]);
		if (pIcon == NULL)
		{
			continue;
		}

		ASSERT_VALID (pIcon);

		if (pIcon->m_nIndex < 0)	// Scroll button
		{
			int yBottom = yButton + cyMenu;

			if (i == m_arIcons.GetSize () - 1)
			{
				yBottom = rectBorder.bottom;
			}

			pIcon->m_rect = CRect (
				rectButtons.left,
				yButton,
				rectButtons.right,
				yBottom);

			yButton = yBottom;
		}
	}
}
//***************************************************************************
void CBCGPRibbonPaletteButton::CleanUpSizes()
{
	CBCGPRibbonButton::CleanUpSizes();

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		ASSERT_VALID(m_arIcons[i]);
		m_arIcons[i]->CleanUpSizes();
	}
}
//***************************************************************************
void CBCGPRibbonPaletteButton::OnDrawPaletteIcon (CDC* pDC, CRect rectIcon, 
												  int nIconIndex, CBCGPRibbonPaletteIcon* pIcon,
												  COLORREF clrText)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	if (m_bIsOwnerDraw)
	{
		CBCGPRibbonGalleryCtrl* pCtrl = DYNAMIC_DOWNCAST(CBCGPRibbonGalleryCtrl, m_pParentControl);
		if (pCtrl != NULL)
		{
			ASSERT_VALID(pCtrl);
			pCtrl->OnDrawGalleryItem(pDC, rectIcon, nIconIndex, pIcon, clrText);
			return;
		}

		if (!m_bIsComboMode)
		{
			// You should implement OnDrawPaletteIcon in your
			// CBCGPRibbonPaletteButton-derived class!
			ASSERT (FALSE);
			return;
		}
	}

	if (!m_bSmallIcons)
	{
		rectIcon.DeflateRect (nImageMargin, nImageMargin);
	}

	if (nIconIndex >= 0 && nIconIndex < m_imagesPalette.GetCount())
	{
		int xOffset = 0;
		int yOffset = 0;

		if (m_bIsComboMode)
		{
			xOffset = globalUtils.ScaleByDPI(pIcon->m_szMargin.cx / 2);
			yOffset = max(0, (rectIcon.Height() - m_imagesPalette.GetImageSize().cy) / 2);
		}

		m_imagesPalette.Draw (pDC, rectIcon.left + xOffset, rectIcon.top + yOffset, nIconIndex, FALSE, IsDisabled ());
	}

	if (m_bIsComboMode)
	{
		CRect rectText = rectIcon;
		rectText.left += GetIconSize().cx + globalUtils.ScaleByDPI(3 * pIcon->m_szMargin.cx);

		COLORREF clrTextOld = (COLORREF)-1;
		if (clrText != (COLORREF)-1)
		{
			clrTextOld = pDC->SetTextColor (clrText);
		}
		
		OnDrawComboItemText(pDC, pIcon, rectText);

		if (clrTextOld != (COLORREF)-1)
		{
			pDC->SetTextColor(clrTextOld);
		}
	}
}
//***************************************************************************
void CBCGPRibbonPaletteButton::OnDrawComboItemText(CDC* pDC, CBCGPRibbonPaletteIcon* pIcon, CRect rectText)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);
	ASSERT_VALID(pIcon);

	CString strDescr = pIcon->m_strDescription;
	
	UINT dtFlags = DT_LEFT | DT_SINGLELINE;
	
	CString strLabel = pIcon->GetText();
	if (strLabel.IsEmpty())
	{
		strLabel = GetIconTextLabel(pIcon);
	}
	
	if (strDescr.IsEmpty())
	{
		pDC->DrawText(strLabel, rectText, DT_LEFT | DT_SINGLELINE | DT_VCENTER);
	}
	else
	{
		CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar();
		
		CFont* pOldFont = pDC->SelectObject((pTopLevelRibbon->GetSafeHwnd() != NULL && m_pParentControl == NULL) ? pTopLevelRibbon->GetBoldFont() : &globalData.fontBold);
		ASSERT(pOldFont != NULL);
		
		int nTextHeight = pDC->DrawText(strLabel, rectText, dtFlags);
		
		pDC->SelectObject(pOldFont);
		pOldFont = NULL;
		
		rectText.top += nTextHeight + m_szMargin.cy;
		
		if (m_pParentControl != NULL)
		{
			pOldFont = pDC->SelectObject(&globalData.fontRegular);
		}
		
		pDC->DrawText(strDescr, rectText, DT_LEFT | DT_SINGLELINE);
		
		if (pOldFont != NULL)
		{
			pDC->SelectObject(pOldFont);
		}
	}
}
//***************************************************************************
CSize CBCGPRibbonPaletteButton::GetComboItemTextSize(CDC* pDC, CBCGPRibbonPaletteIcon* pIcon)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);
	ASSERT_VALID(pIcon);

	int cxText = 0;
	int cyText = 0;
	
	CString strLabel = pIcon->GetText();
	if (strLabel.IsEmpty())
	{
		strLabel = GetIconTextLabel(pIcon);
	}
	
	if (!pIcon->m_strDescription.IsEmpty())
	{
		CBCGPRibbonBar* pTopLevelRibbon = pIcon->GetTopLevelRibbonBar();
		
		CFont* pOldFont = pDC->SelectObject((pTopLevelRibbon->GetSafeHwnd() != NULL && m_pParentControl == NULL) ? pTopLevelRibbon->GetBoldFont() : &globalData.fontBold);
		
		CSize sizeLabel = pDC->GetTextExtent(strLabel);
		pDC->SelectObject(pOldFont);
		
		CSize sizeDescr = pDC->GetTextExtent(pIcon->m_strDescription);
		
		cxText = max(sizeLabel.cx, sizeDescr.cx);
		cyText = sizeLabel.cy + sizeDescr.cy + 2 * m_szMargin.cy;
	}
	else
	{
		CSize sizeLabel = pDC->GetTextExtent(strLabel);
		
		cxText = sizeLabel.cx;
		cyText = sizeLabel.cy + 2 * pIcon->m_szMargin.cy;
	}
	
	return CSize(cxText, cyText);
}
//***************************************************************************
void CBCGPRibbonPaletteButton::DrawImage (CDC* pDC, RibbonImageType type, CRect rectImage)
{
	ASSERT_VALID(this);

	if (!m_bShowSelectedImage)
	{
		CBCGPRibbonButton::DrawImage(pDC, type, rectImage);
		return;
	}

	if (m_bIsOwnerDraw)
	{
		// You should implement DrawImage in your
		// CBCGPRibbonPaletteButton-derived class!
		ASSERT (FALSE);
		return;
	}

	if (m_nSelected >= 0 && m_nSelected < m_imagesPalette.GetCount())
	{
		if (IsDrawSimplifiedIcon())
		{
			if (m_hIconSimplified == NULL)
			{
				CBCGPToolBarImages images;
				images.SetImageSize(m_imagesPalette.GetImageSize());
				
				HICON hIcon = m_imagesPalette.ExtractIcon(m_nSelected);
				if (hIcon != NULL)
				{
					images.AddIcon(hIcon, TRUE);
					images.Simplify();

					m_hIconSimplified = images.ExtractIcon(0);

					::DestroyIcon(hIcon);
				}
			}

			if (m_hIconSimplified != NULL)
			{
				CBCGPRibbonButton::DrawImage(pDC, type, rectImage);
				return;
			}
		}

		CSize sizeImage = m_imagesPalette.GetImageSize();
		if (sizeImage.cx != 0 && sizeImage.cy != 0)
		{
			double dAspect = min((double)rectImage.Width() / (double)sizeImage.cx, (double)rectImage.Height() / (double)sizeImage.cy);
			
			sizeImage.cx = bcg_clamp((long)((double)sizeImage.cx * dAspect), 0, rectImage.Width());
			sizeImage.cy = bcg_clamp((long)((double)sizeImage.cy * dAspect), 0, rectImage.Height());
			
			rectImage.OffsetRect((rectImage.Width() - sizeImage.cx) / 2, (rectImage.Height() - sizeImage.cy) / 2);
			rectImage.right = rectImage.left + sizeImage.cx;
			rectImage.bottom = rectImage.top + sizeImage.cy;
		}

		m_imagesPalette.DrawEx(pDC, rectImage, m_nSelected, CBCGPToolBarImages::ImageAlignHorzStretch, CBCGPToolBarImages::ImageAlignVertStretch, CRect(0, 0, 0, 0),
			IsDisabled() ? 127 : 255);
	}
}
//***************************************************************************
void CBCGPRibbonPaletteButton::AddScrollButtons()
{
	if (!m_bIsComboMode)
	{
		m_arIcons.Add (new CBCGPRibbonPaletteIcon (this, nScrollUpID));
		m_arIcons.Add (new CBCGPRibbonPaletteIcon (this, nScrollDownID));
		m_arIcons.Add (new CBCGPRibbonPaletteIcon (this, nMenuID));
	}
}
//***************************************************************************
void CBCGPRibbonPaletteButton::CreateIcons ()
{
	ASSERT_VALID (this);

	int nGroupIndex = 0;
	int i = 0;
	
	for (i = 0; i < m_nIcons; i++)
	{
		if (nGroupIndex < m_arGroupLen.GetSize () &&
			i == m_arGroupLen [nGroupIndex])
		{
			CString strLabel = m_arGroupNames [nGroupIndex++];
			if (!strLabel.IsEmpty ())
			{
				strLabel = _T("   ") + strLabel;
			}

			CBCGPRibbonLabel* pLabel = new CBCGPRibbonLabel (strLabel);

			pLabel->SetData ((DWORD_PTR) pLabel);
			pLabel->m_bIsPaletteGroup = TRUE;

			m_arIcons.Add (pLabel);
		}

		CBCGPRibbonPaletteIcon* pIcon = new CBCGPRibbonPaletteIcon (this, i);

		if (i == m_nSelected)
		{
			pIcon->m_bIsChecked = TRUE;
		}

		if (i >= 0 && i < (int)m_arKeys.GetSize())
		{
			pIcon->SetKeys(m_arKeys[i]);
		}
		
		m_arIcons.Add (pIcon);
	}

	AddScrollButtons();
}
//***************************************************************************
void CBCGPRibbonPaletteButton::RebuildIconLocations ()
{
	ASSERT_VALID (this);

	CRect rectImages = m_rect;

	const CSize sizeItem = GetItemSize ();

	int nMargin = m_bSmallIcons ? 0 : nImageMargin;
	rectImages.DeflateRect (0, nMargin);

	int yOffset = max (0, 
		(rectImages.Height () - (sizeItem.cy + 2 * nMargin) * m_nImagesInColumn) / 2);

	int nRow = 0;
	int nColumn = 0;

	CSize sizeIcon (sizeItem.cx + 2 * nMargin, sizeItem.cy + 2 * nMargin);

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST (
			CBCGPRibbonPaletteIcon, m_arIcons [i]);
		if (pIcon == NULL)
		{
			continue;
		}

		ASSERT_VALID (pIcon);

		pIcon->m_bIsFirstInRow = FALSE;
		pIcon->m_bIsLastInRow = FALSE;
		pIcon->m_bIsFirstInColumn = FALSE;
		pIcon->m_bIsLastInColumn = FALSE;

		pIcon->m_pParentMenu = m_pParentMenu;

		if (pIcon->m_nIndex < 0)	// Scroll button
		{
			if (pIcon->m_nIndex == nScrollUpID)
			{
				pIcon->m_bIsDisabled = (m_nScrollOffset == 0);
			}
			else if (pIcon->m_nIndex == nScrollDownID)
			{
				pIcon->m_bIsDisabled = (m_nScrollOffset >= m_nScrollTotal);
			}

			continue;
		}

		if (nRow - m_nScrollOffset >= m_nImagesInColumn || 
			nRow < m_nScrollOffset)
		{
			pIcon->m_rect.SetRectEmpty ();
		}
		else
		{
			CRect rectIcon (
				CPoint (
					rectImages.left + sizeIcon.cx * nColumn + nImageMargin / 2,
					rectImages.top + sizeIcon.cy * (nRow - m_nScrollOffset) + yOffset),
				sizeIcon);

			pIcon->m_rect = rectIcon;

			pIcon->m_bIsFirstInRow = (nColumn == 0);
			pIcon->m_bIsLastInRow = (nColumn == m_nImagesInRow - 1);
			pIcon->m_bIsFirstInColumn = (nRow - m_nScrollOffset == 0);
			pIcon->m_bIsLastInColumn = (nRow - m_nScrollOffset == m_nImagesInColumn - 1);
		}

		pIcon->SetCheckRect();

		nColumn++;
		
		if (nColumn == m_nImagesInRow)
		{
			nColumn = 0;
			nRow++;
		}
	}

	CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar ();
	if (pTopLevelRibbon != NULL)
	{
		pTopLevelRibbon->HideKeyTips ();
	}
}
//***************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonPaletteButton::HitTest (CPoint point)
{
	ASSERT_VALID (this);

	if (IsButtonLook ())
	{
		return CBCGPRibbonButton::HitTest (point);
	}

	if (IsDisabled ())
	{
		return NULL;
	}

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		if (m_arIcons [i]->GetRect ().PtInRect (point))
		{
			return m_arIcons [i];
		}
	}

	return NULL;
}
//***************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonPaletteButton::GetPressed ()
{
	ASSERT_VALID (this);

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arIcons [i];
		ASSERT_VALID (pButton);

		CBCGPBaseRibbonElement* pElem = pButton->GetPressed ();
		if (pElem != NULL)
		{
			ASSERT_VALID (pElem);
			return pElem;
		}
	}
	
	return NULL;
}
//***************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonPaletteButton::GetHighlighted ()
{
	ASSERT_VALID (this);

	if (IsButtonLook ())
	{
		return CBCGPRibbonButton::GetHighlighted ();
	}

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arIcons [i];
		ASSERT_VALID (pButton);

		CBCGPBaseRibbonElement* pElem = pButton->GetHighlighted ();
		if (pElem != NULL)
		{
			ASSERT_VALID (pElem);
			return pElem;
		}
	}
	
	return NULL;
}
//***************************************************************************
void CBCGPRibbonPaletteButton::OnClickPaletteIcon (CBCGPRibbonPaletteIcon* pIcon)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pIcon);

	switch (pIcon->m_nIndex)
	{
	case nScrollUpID:
		m_nScrollOffset = max (0, m_nScrollOffset - 1);
		RebuildIconLocations ();
		Redraw ();
		break;

	case nScrollDownID:
		m_nScrollOffset = min (m_nScrollTotal, m_nScrollOffset + 1);
		RebuildIconLocations ();
		Redraw ();
		break;

	case nMenuID:
		// Already shown in CBCGPRibbonPaletteIcon::OnLButtonDown
		break;

	default:
		{
			int nIconIndex = 0;

			for (int i = 0; i < m_arIcons.GetSize (); i++)
			{
				CBCGPRibbonPaletteIcon* pListIcon = DYNAMIC_DOWNCAST (
					CBCGPRibbonPaletteIcon, m_arIcons [i]);
				if (pListIcon == NULL)
				{
					continue;
				}

				ASSERT_VALID (pListIcon);

				if (pListIcon->m_bIsChecked)
				{
					pListIcon->m_bIsChecked = FALSE;
				}

				if (pListIcon == pIcon)
				{
					if (m_bShowSelectedImage && m_hIconSimplified != NULL)
					{
						::DestroyIcon(m_hIconSimplified);
						m_hIconSimplified = NULL;
					}

					m_nSelected = nIconIndex;

					pIcon->m_bIsChecked = TRUE;

					if (pIcon->m_rect.IsRectEmpty () && m_nImagesInRow > 0)
					{
						m_nScrollOffset = nIconIndex / m_nImagesInRow;
						m_nScrollOffset = min (m_nScrollTotal, m_nScrollOffset);
						RebuildIconLocations ();
					}

					if (m_bShowSelectedImage)
					{
						if (m_bShowSelectedLabel)
						{
							m_strToolTip = pIcon->GetToolTipText();
							m_strDescription = pIcon->GetDescription();
						}

						CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar ();
						if (pRibbonBar != NULL)
						{
							ASSERT_VALID (pRibbonBar);
							
							CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arButtons;
							pRibbonBar->GetElementsByID (m_nID, arButtons, TRUE);
							
							for (int j = 0; j < arButtons.GetSize (); j++)
							{
								CBCGPRibbonPaletteButton* pOther = DYNAMIC_DOWNCAST(CBCGPRibbonPaletteButton, arButtons [j]);
								
								if (pOther != NULL && pOther != this)
								{
									ASSERT_VALID (pOther);
									
									pOther->SelectItem(m_nSelected);
								}
							}
						}
					}
				}

				nIconIndex++;
			}
		}

		Redraw ();

		m_mapSelectedItems.SetAt (m_nPaletteID == 0 ? m_nID : m_nPaletteID, pIcon->m_nIndex);

		if (m_pParentControl->GetSafeHwnd() != NULL)
		{
			m_pParentControl->RedrawWindow();
		}
	}
}
//*****************************************************************************
void CBCGPRibbonPaletteButton::CopyFrom (const CBCGPBaseRibbonElement& s)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::CopyFrom (s);

	if (!s.IsKindOf (RUNTIME_CLASS (CBCGPRibbonPaletteButton)))
	{
		return;
	}

	CBCGPRibbonPaletteButton& src = (CBCGPRibbonPaletteButton&) s;

	RemoveAll ();

	src.m_imagesPalette.CopyTo (m_imagesPalette);
	m_bSmallIcons = src.m_bSmallIcons;
	m_nSelected = src.m_nSelected;
	m_bEnableMenuResize = src.m_bEnableMenuResize;
	m_bMenuResizeVertical = src.m_bMenuResizeVertical;
	m_nPaletteID = src.m_nPaletteID;
	m_bIsButtonMode = src.m_bIsButtonMode;
	m_bShowSelectedImage = src.m_bShowSelectedImage;
	m_bShowSelectedLabel = src.m_bShowSelectedLabel;
	m_nIconsInRow = src.m_nIconsInRow;
	m_nPanelColumns = src.m_nPanelColumns;
	m_nInitialColumns = src.m_nInitialColumns;
	m_bIsOwnerDraw = src.m_bIsOwnerDraw;
	m_bDefaultButtonStyle = src.m_bDefaultButtonStyle;
	m_nIcons = src.m_nIcons;
	m_bMenuSideBar = src.m_bMenuSideBar;
	m_bIsComboMode = src.m_bIsComboMode;
	m_bDrawCheckFrameAroundImage = src.m_bDrawCheckFrameAroundImage;
	m_bDrawDisabledItems = src.m_bDrawDisabledItems;
	m_bItemCheckBoxes = src.m_bItemCheckBoxes;
	m_bCheckBoxOverlapsIcon = src.m_bCheckBoxOverlapsIcon;
	m_CheckBoxLocation = src.m_CheckBoxLocation;
	m_bItemTextLabels = src.m_bItemTextLabels;
	m_TextLabelLocation = src.m_TextLabelLocation;
	m_clrTextLabel = src.m_clrTextLabel;

	ASSERT (src.m_arGroupNames.GetSize () == src.m_arGroupLen.GetSize ());

	m_arGroupNames.RemoveAll ();
	m_arGroupLen.RemoveAll ();

	int i = 0;

	for (i = 0; i < src.m_arGroupNames.GetSize (); i++)
	{
		m_arGroupNames.Add (src.m_arGroupNames [i]);
		m_arGroupLen.Add (src.m_arGroupLen [i]);
	}

	m_arToolTips.Copy(src.m_arToolTips);
	m_arUserData.Copy(src.m_arUserData);
	m_arKeys.Copy(src.m_arKeys);

	CreateIcons ();

	if (m_bItemCheckBoxes && m_arIcons.GetSize() == src.m_arIcons.GetSize())
	{
		for (i = 0; i < (int)m_arIcons.GetSize(); i++)
		{
			CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST (CBCGPRibbonPaletteIcon, m_arIcons [i]);
			CBCGPRibbonPaletteIcon* pSrcIcon = DYNAMIC_DOWNCAST (CBCGPRibbonPaletteIcon, src.m_arIcons [i]);

			if (pIcon != NULL && pSrcIcon != NULL)
			{
				pIcon->m_nCheckState = pSrcIcon->m_nCheckState;
			}
		}
	}
}
//*****************************************************************************
void CBCGPRibbonPaletteButton::SetParentCategory (CBCGPRibbonCategory* pParent)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::SetParentCategory (pParent);

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		ASSERT_VALID (m_arIcons [i]);
		m_arIcons [i]->SetParentCategory (pParent);
	}
}
//*****************************************************************************
void CBCGPRibbonPaletteButton::OnShowPopupMenu ()
{
	ASSERT_VALID (this);

	if (m_pParentControl != NULL && m_pRibbonBar == NULL)
	{
		CWnd* pWndFrame = AfxGetMainWnd ();		
		ASSERT_VALID (pWndFrame);

		if (pWndFrame->IsKindOf (RUNTIME_CLASS (CBCGPFrameWnd)))
		{
			m_pRibbonBar = ((CBCGPFrameWnd*) pWndFrame)->GetRibbonBar ();
		}
		else if (pWndFrame->IsKindOf (RUNTIME_CLASS (CBCGPMDIFrameWnd)))
		{
			m_pRibbonBar = ((CBCGPMDIFrameWnd*) pWndFrame)->GetRibbonBar ();
		}

		if (m_pRibbonBar == NULL)
		{
			ASSERT(FALSE);
			return;
		}
	}

	CWnd* pWndParent = m_pParentControl == NULL ? GetParentWnd () : m_pParentControl;
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

	const BOOL bIsRTL = (pTopLevelRibbon->GetExStyle () & WS_EX_LAYOUTRTL);

	if (m_arIcons.GetSize () == 0)
	{
		CreateIcons ();
	}

	int nSel = GetLastSelectedItem (m_nPaletteID == 0 ? m_nID : m_nPaletteID);
	if (nSel >= 0)
	{
		SelectItem (nSel);
	}

	CBCGPBaseRibbonElement* pMenuButton = IsButtonLook () ? this : m_arIcons [m_arIcons.GetSize () - 1];

	CWnd* pWndOwner = pTopLevelRibbon->GetTopLevelOwner();

	CBCGPRibbonPanelMenu* pMenu = new CBCGPRibbonPanelMenu (this);

	pMenu->SetParentRibbonElement (pMenuButton);
	pMenu->SetMenuMode ();
	
	CRect rectBtn = GetRect ();
	if (m_pParentControl != NULL)
	{
		m_pParentControl->GetClientRect (rectBtn);
	}

	pWndParent->ClientToScreen (&rectBtn);

	int nMargin = m_bSmallIcons ? 0 : nImageMargin;
	const CSize sizeItem = GetItemSize ();

	CSize sizeIcon (sizeItem.cx + 2 * nMargin, sizeItem.cy + 2 * nMargin);

	int x = bIsRTL ? rectBtn.right : rectBtn.left;
	int y = rectBtn.bottom;

	if (IsMenuMode ())
	{
		x = bIsRTL ? rectBtn.left : rectBtn.right;
		y = rectBtn.top;
	}

	if (!IsButtonLook ())
	{
		x = bIsRTL ? rectBtn.right : rectBtn.left;
		y = rectBtn.top + nBorderMarginY;
	}

	if (m_bIsComboMode)
	{
		CClientDC dc (pTopLevelRibbon);

		CFont* pOldFont = dc.SelectObject(m_pParentControl == NULL ? pTopLevelRibbon->GetFont() : &globalData.fontRegular);
		ASSERT (pOldFont != NULL);

		int cxMax = 0;

		for (int i = 0; i < m_arIcons.GetSize (); i++)
		{
			CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST (CBCGPRibbonPaletteIcon, m_arIcons [i]);
			if (pIcon != NULL)
			{
				cxMax = max(cxMax, pIcon->GetIntermediateSize(&dc).cx);
			}
		}
		
		dc.SelectObject (pOldFont);

		cxMax += ::GetSystemMetrics (SM_CXVSCROLL);

		if (m_pParentControl->GetSafeHwnd() != NULL)
		{
			CRect rectParent;
			m_pParentControl->GetClientRect (rectParent);

			cxMax = max(cxMax, rectParent.Width());
		}

		pMenu->SetPreferedSize (CSize (cxMax, 0));
	}
	else
	{
		int cxPreferred = 0;

		if (m_nIconsInRow > 0)
		{
			cxPreferred = m_nIconsInRow * sizeIcon.cx;
		}
		else
		{
			const int nPanelColumns = pMenuButton == this ? 4 : m_nPanelColumns;
			const int nIconsInRow = m_bSmallIcons ? (m_nInitialColumns > 0 ? m_nInitialColumns : 10) : max (nPanelColumns, 4);

			cxPreferred = nIconsInRow * sizeIcon.cx;
		}

		pMenu->SetPreferedSize (CSize (cxPreferred, 0));
	}

	pMenu->Create (pWndOwner, x, y, (HMENU) NULL);
	pMenuButton->SetDroppedDown (pMenu);

	m_bOnBeforeShowItemMenuIsSent = FALSE;

	if (pMenu->HasBeenResized ())
	{
		pMenu->TriggerResize ();
	}
}
//**************************************************************************
void CBCGPRibbonPaletteButton::SetComboMode(BOOL bSet/* = TRUE*/, BOOL bDrawCheckFrameAroundImage/* = FALSE*/)
{
	if (!IsButtonMode())
	{
		bSet = FALSE;
	}

	m_bIsComboMode = bSet;
	m_bDrawCheckFrameAroundImage = bDrawCheckFrameAroundImage;

	if (bSet)
	{
		for (int i = 0; i < m_arIcons.GetSize ();)
		{
			CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST(CBCGPRibbonPaletteIcon, m_arIcons[i]);
			if (pIcon == NULL)
			{
				i++;
				continue;
			}
			
			ASSERT_VALID(pIcon);

			if (pIcon->GetIndex() == nScrollUpID || pIcon->GetIndex() == nScrollDownID || pIcon->GetIndex() == nMenuID)
			{
				delete pIcon;
				m_arIcons.RemoveAt(i);

			}
			else
			{
				i++;
			}
		}
	}
}
//**************************************************************************
void CBCGPRibbonPaletteButton::SelectItem (int nItemIndex)
{
	ASSERT_VALID (this);

	if (m_bShowSelectedImage && m_hIconSimplified != NULL)
	{
		::DestroyIcon(m_hIconSimplified);
		m_hIconSimplified = NULL;
	}

	m_nSelected = nItemIndex;

	int nCurrIndex = 0;

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST (
			CBCGPRibbonPaletteIcon, m_arIcons [i]);

		if (pIcon == NULL)
		{
			continue;
		}

		ASSERT_VALID (pIcon);

		if (pIcon->m_bIsChecked)
		{
			pIcon->m_bIsChecked = FALSE;
		}

		if (nCurrIndex == nItemIndex)
		{
			pIcon->m_bIsChecked = TRUE;

			if (m_bShowSelectedLabel)
			{
				m_strToolTip = pIcon->GetToolTipText();
				m_strDescription = pIcon->GetDescription();
			}
		}

		nCurrIndex++;
	}

	m_mapSelectedItems.SetAt (m_nPaletteID == 0 ? m_nID : m_nPaletteID, m_nSelected);

	Redraw ();

	if (m_pParentControl->GetSafeHwnd() != NULL)
	{
		m_pParentControl->RedrawWindow();
	}
}
//**************************************************************************
void CBCGPRibbonPaletteButton::SetCheckState(int nItemIndex, int nState)
{
	ASSERT_VALID (this);

	int nCurrIndex = 0;

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST (
			CBCGPRibbonPaletteIcon, m_arIcons [i]);

		if (pIcon == NULL)
		{
			continue;
		}

		ASSERT_VALID (pIcon);

		if (nCurrIndex == nItemIndex)
		{
			pIcon->m_nCheckState = nState;
			break;
		}

		nCurrIndex++;
	}

	Redraw ();

	if (m_pParentControl->GetSafeHwnd() != NULL)
	{
		m_pParentControl->RedrawWindow();
	}
}
//**************************************************************************
int CBCGPRibbonPaletteButton::GetCheckState(int nItemIndex) const
{
	ASSERT_VALID (this);

	int nCurrIndex = 0;

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST (
			CBCGPRibbonPaletteIcon, m_arIcons [i]);

		if (pIcon == NULL)
		{
			continue;
		}

		ASSERT_VALID (pIcon);

		if (nCurrIndex == nItemIndex)
		{
			return pIcon->m_nCheckState;
		}

		nCurrIndex++;
	}

	return -1;
}
//**************************************************************************
void CBCGPRibbonPaletteButton::SetItemToolTip (int nItemIndex, LPCTSTR lpszToolTip)
{
	ASSERT_VALID (this);

	if (nItemIndex < 0)
	{
		ASSERT (FALSE);
		return;
	}

	if (nItemIndex >= m_arToolTips.GetSize ())
	{
		m_arToolTips.SetSize (nItemIndex + 1);
	}

	m_arToolTips [nItemIndex] = lpszToolTip == NULL ? _T("") : lpszToolTip;
}
//**************************************************************************
LPCTSTR CBCGPRibbonPaletteButton::GetItemToolTip (int nItemIndex) const
{
	ASSERT_VALID (this);

	if (nItemIndex < 0 || nItemIndex >= m_arToolTips.GetSize ())
	{
		ASSERT (FALSE);
		return NULL;
	}

	return m_arToolTips [nItemIndex];
}
//**************************************************************************
void CBCGPRibbonPaletteButton::SetItemUserData(int nItemIndex, DWORD_PTR dwUserData)
{
	ASSERT_VALID (this);

	if (nItemIndex < 0)
	{
		ASSERT (FALSE);
		return;
	}

	if (nItemIndex >= m_arUserData.GetSize())
	{
		int nStartIndex = (int)m_arUserData.GetSize();
		m_arUserData.SetSize (nItemIndex + 1);

		for (int i = nStartIndex; i < nItemIndex; i++)
		{
			m_arUserData[i] = 0;
		}
	}

	m_arUserData[nItemIndex] = dwUserData;
}
//**************************************************************************
DWORD_PTR CBCGPRibbonPaletteButton::GetItemUserData(int nItemIndex) const
{
	ASSERT_VALID (this);

	if (nItemIndex < 0 || nItemIndex >= m_arUserData.GetSize())
	{
		return 0;
	}

	return m_arUserData[nItemIndex];
}
//**************************************************************************
void CBCGPRibbonPaletteButton::SetItemKeys(int nItemIndex, LPCTSTR lpszKeys)
{
	ASSERT_VALID (this);

	if (nItemIndex < 0)
	{
		ASSERT (FALSE);
		return;
	}

	if (nItemIndex >= (int)m_arKeys.GetSize())
	{
		m_arKeys.SetSize(nItemIndex + 1);
	}

	m_arKeys[nItemIndex] = lpszKeys == NULL ? _T("") : lpszKeys;

	int nCurrIndex = 0;
	
	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST (
			CBCGPRibbonPaletteIcon, m_arIcons [i]);
		
		if (pIcon == NULL)
		{
			continue;
		}
		
		ASSERT_VALID (pIcon);
		
		if (nCurrIndex == nItemIndex)
		{
			pIcon->SetKeys(lpszKeys);
			break;
		}
		
		nCurrIndex++;
	}
}
//**************************************************************************
LPCTSTR CBCGPRibbonPaletteButton::GetItemKeys(int nItemIndex) const
{
	ASSERT_VALID (this);

	if (nItemIndex < 0 || nItemIndex >= (int)m_arKeys.GetSize())
	{
		ASSERT (FALSE);
		return NULL;
	}

	return m_arKeys[nItemIndex];
}
//**************************************************************************
void CBCGPRibbonPaletteButton::RemoveItemToolTips ()
{
	ASSERT_VALID (this);
	m_arToolTips.RemoveAll ();
}
//**************************************************************************
CString CBCGPRibbonPaletteButton::GetIconToolTip (const CBCGPRibbonPaletteIcon* pIcon) const
{
	ASSERT_VALID (this);
	ASSERT_VALID (pIcon);

	int nIndex = pIcon->m_nIndex;

	CString strTip;

	switch (nIndex)
	{
	case nScrollUpID:
	case nScrollDownID:
		{
			CBCGPLocalResource locaRes;

			if (m_nImagesInColumn == 1)
			{
				strTip.Format (IDS_BCGBARRES_GALLERY_ROW1_FMT, m_nScrollOffset + 1, m_nScrollTotal + m_nImagesInColumn);
			}
			else
			{
				strTip.Format (IDS_BCGBARRES_GALLERY_ROW2_FMT, m_nScrollOffset + 1, m_nScrollOffset + m_nImagesInColumn, m_nScrollTotal + m_nImagesInColumn);
			}
			return strTip;
		}

	case nMenuID:
		{
			CBCGPLocalResource locaRes;
			strTip.LoadString (IDS_BCGBARRES_MORE);
		}
		return strTip;
	}

	if (nIndex < 0 || nIndex >= m_arToolTips.GetSize ())
	{
		return _T("");
	}

	AfxExtractSubString(strTip, m_arToolTips [nIndex], 0, TCHAR('\n'));

	return strTip;
}
//**************************************************************************
CString CBCGPRibbonPaletteButton::GetIconDescription (const CBCGPRibbonPaletteIcon* pIcon) const
{
	ASSERT_VALID (this);
	ASSERT_VALID (pIcon);

	int nIndex = pIcon->m_nIndex;

	switch (nIndex)
	{
	case nScrollUpID:
	case nScrollDownID:
	case nMenuID:
		return m_strDescription;
	}

	CString strDesc;

	if (0 <= nIndex && nIndex < m_arToolTips.GetSize ())
	{
		AfxExtractSubString(strDesc, m_arToolTips [nIndex], 1, TCHAR('\n'));
	}

	return strDesc;
}
//**************************************************************************
CString CBCGPRibbonPaletteButton::GetIconTextLabel(const CBCGPRibbonPaletteIcon* pIcon) const
{
	int nIndex = pIcon->m_nIndex;

	CString strLabel;

	if (0 <= nIndex && nIndex < m_arToolTips.GetSize ())
	{
		AfxExtractSubString(strLabel, m_arToolTips [nIndex], 2, TCHAR('\n'));

		if (strLabel.IsEmpty())
		{
			strLabel = GetIconToolTip(pIcon);
		}
	}

	return strLabel;
}
//**************************************************************************
int CBCGPRibbonPaletteButton::GetLastSelectedItem (UINT uiCmdID)
{
	int nIndex = -1;

	m_mapSelectedItems.Lookup (uiCmdID, nIndex);
	return nIndex;
}
//**************************************************************************
void CBCGPRibbonPaletteButton::ClearLastSelectedItem(UINT uiCmdID)
{
	m_mapSelectedItems.SetAt (uiCmdID, -1);
}
//**************************************************************************
void CBCGPRibbonPaletteButton::GetMenuItems (
	CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arButtons)
{
	ASSERT_VALID (this);

	int i = 0;

	for (i = 0; i < m_arSubItems.GetSize (); i++)
	{
		arButtons.Add (m_arSubItems [i]);
	}

	const int nScrollButtons = m_bIsComboMode ? 0 : 3;

	for (i = 0; i < m_arIcons.GetSize () - nScrollButtons; i++)
	{
		arButtons.Add (m_arIcons [i]);
	}
}
//***************************************************************************
int CBCGPRibbonPaletteButton::GetMenuRowHeight () const
{
	ASSERT_VALID (this);

	int nMargin = m_bSmallIcons ? 0 : nImageMargin;
	const CSize sizeItem = GetItemSize();
	
	return sizeItem.cy + 2 * nMargin;
}
//***************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonPaletteButton::GetDroppedDown ()
{
	ASSERT_VALID (this);

	if (m_arIcons.GetSize () > 0)
	{
		CBCGPBaseRibbonElement* pMenuButton = m_arIcons [m_arIcons.GetSize () - 1];
		ASSERT_VALID (pMenuButton);

		if (pMenuButton->IsDroppedDown ())
		{
			return pMenuButton;
		}
	}

	return CBCGPRibbonButton::GetDroppedDown ();
}
//***************************************************************************
void CBCGPRibbonPaletteButton::OnEnable (BOOL bEnable)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::OnEnable (bEnable);

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		ASSERT_VALID (m_arIcons [i]);
		m_arIcons [i]->m_bIsDisabled = !bEnable;
	}
}
//***************************************************************************
void CBCGPRibbonPaletteButton::SetNotifyParentID (BOOL bSet)
{
	m_bNotifyPaletteID = bSet;

	if (m_pOriginal != NULL)
	{
		ASSERT_VALID (m_pOriginal);

		CBCGPRibbonPaletteButton* pOriginal = DYNAMIC_DOWNCAST (
			CBCGPRibbonPaletteButton, m_pOriginal);
		if (pOriginal != NULL)
		{
			pOriginal->m_bNotifyPaletteID = bSet;
		}
	}
}
//***************************************************************************
BOOL CBCGPRibbonPaletteButton::OnKey (BOOL /*bIsMenuKey*/)
{
	ASSERT_VALID (this);

	return CBCGPRibbonButton::OnKey (TRUE);
}
//***************************************************************************
CRect CBCGPRibbonPaletteButton::GetKeyTipRect (CDC* pDC, BOOL bIsMenu)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	if (IsButtonLook ())
	{
		return CBCGPRibbonButton::GetKeyTipRect (pDC, bIsMenu);
	}

	CSize sizeKeyTip = GetKeyTipSize (pDC);
	CRect rectKeyTip (0, 0, 0, 0);

	if (sizeKeyTip == CSize (0, 0) || m_rect.IsRectEmpty ())
	{
		return rectKeyTip;
	}

	rectKeyTip.left = m_rect.right - sizeKeyTip.cx / 2;
	rectKeyTip.top = m_rect.bottom - sizeKeyTip.cy / 2;

	// Find 'nMenuID' button:
	if (m_arIcons.GetSize () > 0)
	{
		CBCGPRibbonPaletteIcon* pMenuButton = DYNAMIC_DOWNCAST (CBCGPRibbonPaletteIcon, m_arIcons [m_arIcons.GetSize () - 1]);
		if (pMenuButton != NULL)
		{
			ASSERT_VALID (pMenuButton);

			const CRect rectMenu = pMenuButton->GetRect ();

			if (pMenuButton->m_nIndex == nMenuID && !rectMenu.IsRectEmpty ())
			{
				rectKeyTip.left = rectMenu.CenterPoint ().x;
				rectKeyTip.top = rectMenu.bottom - 3;
			}
		}
	}

	rectKeyTip.right = rectKeyTip.left + sizeKeyTip.cx;
	rectKeyTip.bottom = rectKeyTip.top + sizeKeyTip.cy;

	return rectKeyTip;
}
//***************************************************************************
CSize CBCGPRibbonPaletteButton::GetIconSize () const
{
	ASSERT_VALID (this);

	return (m_bIsOwnerDraw || m_imagesPalette.GetCount() > 0) ? (m_imagesPalette.GetImageSize ()) : CSize(0, 0);
}
//**************************************************************************************
CSize CBCGPRibbonPaletteButton::GetItemSize () const 
{ 
	CSize size = GetIconSize();
	if (size != CSize(0, 0) && m_bItemCheckBoxes && !m_bCheckBoxOverlapsIcon)
	{
		size.cx += CBCGPVisualManager::GetInstance ()->GetCheckRadioDefaultSize().cx;
	}

	return size; 
}
//**************************************************************************************
void CBCGPRibbonPaletteButton::OnRTLChanged (BOOL bIsRTL)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::OnRTLChanged (bIsRTL);

	m_imagesPalette.Mirror ();
}
//**************************************************************************************
void CBCGPRibbonPaletteButton::OnSetFocus (BOOL bSet)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::OnSetFocus (bSet);

	for (int i = (int)m_arIcons.GetSize () - 1; i >= 0; i--)
	{
		CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST (
			CBCGPRibbonPaletteIcon, m_arIcons [i]);
		if (pIcon != NULL)
		{
			ASSERT_VALID (pIcon);

			if (pIcon->m_nIndex == nMenuID)
			{
				pIcon->m_bIsFocused = bSet;
				pIcon->Redraw ();
				return;
			}
		}
	}
}
//**************************************************************************************
BOOL CBCGPRibbonPaletteButton::SetACCData (CWnd* pParent, CBCGPAccessibilityData& data)
{
	CBCGPRibbonButton::SetACCData (pParent, data);
	data.m_nAccRole = ROLE_SYSTEM_BUTTONDROPDOWNGRID;

	return TRUE;
}
//*************************************************************************************
void CBCGPRibbonPaletteButton::GetVisibleElements (CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arElements)
{
	ASSERT_VALID (this);

	if (IsButtonLook ())
	{
		CBCGPRibbonButton::GetVisibleElements(arElements);
		return;
	}

	CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar();
	if (pRibbonBar != NULL && pRibbonBar->IsInKeyboardNavigation())
	{
		CBCGPRibbonButton::GetVisibleElements(arElements);
		return;
	}

	for (int i = 0; i < m_arIcons.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pButton = m_arIcons [i];
		ASSERT_VALID (pButton);

		pButton->GetVisibleElements (arElements);
	}
}
//*************************************************************************************
CBCGPRibbonPaletteIcon* CBCGPRibbonPaletteButton::GetIconButton(int nIndex)
{
	ASSERT_VALID(this);
	
	if (nIndex < 0 || nIndex >= (int)m_arIcons.GetSize())
	{
		ASSERT(FALSE);
		return NULL;
	}

	return DYNAMIC_DOWNCAST(CBCGPRibbonPaletteIcon, m_arIcons[nIndex]);
}

////////////////////////////////////////////////
// CBCGPRibbonPaletteMenuButton

IMPLEMENT_DYNCREATE(CBCGPRibbonPaletteMenuButton, CBCGPToolbarMenuButton)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPRibbonPaletteMenuButton::CBCGPRibbonPaletteMenuButton()
{
	CommonInit ();
}
//**************************************************************************************
CBCGPRibbonPaletteMenuButton::CBCGPRibbonPaletteMenuButton(UINT uiID, int iImage, LPCTSTR lpszText,
		CBCGPToolBarImages& imagesPalette) :
	CBCGPToolbarMenuButton (uiID, NULL, iImage, lpszText),
	m_paletteButton (0, _T(""), -1, -1, imagesPalette)
{
	CommonInit ();
}
//**************************************************************************************
CBCGPRibbonPaletteMenuButton::CBCGPRibbonPaletteMenuButton (UINT uiID, int iImage, LPCTSTR lpszText,
		UINT uiImagesPaletteResID, int cxPaletteImage) :
	CBCGPToolbarMenuButton (uiID, NULL, iImage, lpszText),
	m_paletteButton (0, _T(""), -1, -1, uiImagesPaletteResID, cxPaletteImage)
{
	CommonInit ();
}
//**************************************************************************************
CBCGPRibbonPaletteMenuButton::~CBCGPRibbonPaletteMenuButton ()
{
}
//**************************************************************************************
void CBCGPRibbonPaletteMenuButton::CommonInit ()
{
	CBCGPRibbonBar* pRibbon = NULL;

	CFrameWnd* pParentFrame = m_pWndParent == NULL ?
		DYNAMIC_DOWNCAST (CFrameWnd, AfxGetMainWnd ()) :
		BCGCBProGetTopLevelFrame (m_pWndParent);

	CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST (CBCGPMDIFrameWnd, pParentFrame);
	if (pMainFrame != NULL)
	{
		pRibbon = pMainFrame->GetRibbonBar ();
	}
	else	// Maybe, SDI frame...
	{
		CBCGPFrameWnd* pFrame = DYNAMIC_DOWNCAST (CBCGPFrameWnd, pParentFrame);
		if (pFrame != NULL)
		{
			pRibbon = pFrame->GetRibbonBar ();
		}
	}

	if (pRibbon != NULL)
	{
		ASSERT_VALID (pRibbon);
		m_paletteButton.SetParentRibbonBar (pRibbon);
	}
	else
	{
		ASSERT (FALSE);	// Main farme should have the ribbon bar!
	}
}
//**************************************************************************************
CBCGPPopupMenu* CBCGPRibbonPaletteMenuButton::CreatePopupMenu ()
{
	ASSERT_VALID (this);

	m_paletteButton.SetID (m_nID);

	m_paletteButton.CBCGPBaseRibbonElement::OnShowPopupMenu ();

	if (m_paletteButton.m_nIcons == 0)
	{
		TRACE(_T("The palette is not initialized! You should add palette icons first.\n"));
		ASSERT (FALSE);
		return NULL;
	}

	if (m_paletteButton.m_arIcons.GetSize () == 0)
	{
		m_paletteButton.CreateIcons ();
	}

	m_paletteButton.SelectItem (CBCGPRibbonPaletteButton::GetLastSelectedItem (
		m_paletteButton.m_nPaletteID == 0 ? m_paletteButton.m_nID : m_paletteButton.m_nPaletteID));

	for (int i = 0; i < m_paletteButton.m_arSubItems.GetSize (); i++)
	{
		ASSERT_VALID (m_paletteButton.m_arSubItems [i]);
		m_paletteButton.m_arSubItems [i]->SetParentRibbonBar (m_paletteButton.m_pRibbonBar);
	}

	CBCGPRibbonPanelMenu* pMenu = new CBCGPRibbonPanelMenu (&m_paletteButton);

	pMenu->SetMenuMode ();
	
	int nMargin = m_paletteButton.m_bSmallIcons ? 0 : nImageMargin;
	const CSize sizeImage = m_paletteButton.GetIconSize ();
	CSize sizeIcon (sizeImage.cx + 2 * nMargin, sizeImage.cy + 2 * nMargin);

	int nIconsInRow = m_paletteButton.m_nIconsInRow > 0 ? 
		m_paletteButton.m_nIconsInRow : m_paletteButton.m_bSmallIcons ? 10 : 4;

	pMenu->SetPreferedSize (CSize (nIconsInRow * sizeIcon.cx, 0));
	pMenu->EnableCustomizeMenu (FALSE);

	return pMenu;
}
//**************************************************************************************
void CBCGPRibbonPaletteMenuButton::CopyFrom (const CBCGPToolbarButton& s)
{
	CBCGPToolbarMenuButton::CopyFrom (s);

	const CBCGPRibbonPaletteMenuButton& src = (const CBCGPRibbonPaletteMenuButton&) s;

	m_paletteButton.CopyFrom (src.m_paletteButton);
}

////////////////////////////////////////////////
// CBCGPRibbonComboGalleryCtrl

IMPLEMENT_DYNCREATE(CBCGPRibbonComboGalleryCtrl, CBCGPMenuButton)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPRibbonComboGalleryCtrl::CBCGPRibbonComboGalleryCtrl()
{
	m_paletteButton.m_nIconsInRow = 1;
	m_paletteButton.m_bIsComboMode = TRUE;
	m_bHasDescriptions = FALSE;
	m_bDrawFocus = FALSE;
}
//**************************************************************************************
CBCGPRibbonComboGalleryCtrl::~CBCGPRibbonComboGalleryCtrl ()
{
}

BEGIN_MESSAGE_MAP(CBCGPRibbonComboGalleryCtrl, CBCGPMenuButton)
	//{{AFX_MSG_MAP(CBCGPRibbonComboGalleryCtrl)
	ON_WM_LBUTTONDOWN()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

void CBCGPRibbonComboGalleryCtrl::OnDraw (CDC* pDC, const CRect& rect, UINT /*uiState*/)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	CSize sizeArrow = CBCGPMenuImages::Size ();

	CRect rectItem = rect;
	rectItem.right -= sizeArrow.cx + 2 * nImageMargin;

	int nCurSel = m_paletteButton.GetSelectedItem();

	if (nCurSel >= 0)
	{
		for (int i = 0; i <= nCurSel && i < m_paletteButton.m_arIcons.GetSize(); i++)
		{
			if (DYNAMIC_DOWNCAST (CBCGPRibbonLabel, m_paletteButton.m_arIcons [i]) != NULL)
			{
				nCurSel++;
			}
		}

		if (nCurSel < m_paletteButton.m_arIcons.GetSize())
		{
			CBCGPRibbonPaletteIcon* pIcon = DYNAMIC_DOWNCAST (CBCGPRibbonPaletteIcon, m_paletteButton.m_arIcons [nCurSel]);
			if (pIcon != NULL)
			{
				ASSERT_VALID(pIcon);

				CFont* pOldFont = pDC->SelectObject(&globalData.fontRegular);
				pDC->SetBkMode (TRANSPARENT);

				const CSize sizeImage = m_paletteButton.GetIconSize ();
				const CSize sizeItem = pIcon->GetSize(pDC);

				rectItem.top = rectItem.CenterPoint().y - sizeItem.cy / 2;
				rectItem.bottom = rectItem.top + sizeItem.cy;

				CBCGPDrawState ds;

				if (m_paletteButton.m_imagesPalette.GetCount () > 0)
				{
					m_paletteButton.m_imagesPalette.SetTransparentColor (globalData.clrBtnFace);
					m_paletteButton.m_imagesPalette.PrepareDrawImage (ds, sizeImage);
				}

				COLORREF clrText = m_clrText;

				BOOL bIsDisabled = m_paletteButton.m_bIsDisabled;
				if (!IsWindowEnabled())
				{
					clrText = CBCGPVisualManager::GetInstance ()->GetRibbonBarTextColor(TRUE);
					m_paletteButton.m_bIsDisabled = TRUE;
				}

				if (clrText == (COLORREF)-1)
				{
					clrText = m_bVisualManagerStyle ? CBCGPVisualManager::GetInstance ()->GetRibbonBarTextColor () : globalData.clrBtnText;
				}

				m_paletteButton.OnDrawPaletteIcon (pDC, rectItem, pIcon->m_nImageIndex, pIcon, clrText);

				m_paletteButton.m_bIsDisabled = bIsDisabled;
				pDC->SelectObject(pOldFont);

				if (m_paletteButton.m_imagesPalette.GetCount () > 0)
				{
					m_paletteButton.m_imagesPalette.EndDrawImage (ds);
				}
			}
		}
	}

	CRect rectArrow = rect;
	rectArrow.left = rectItem.right;

	CBCGPMenuImages::Draw (pDC, 
		m_bRightArrow ? CBCGPMenuImages::IdArowRightLarge : CBCGPMenuImages::IdArowDownLarge, 
		rectArrow,
		(!IsWindowEnabled()) ? CBCGPMenuImages::ImageGray : CBCGPMenuImages::ImageBlack);
}
//**************************************************************************************
void CBCGPRibbonComboGalleryCtrl::OnShowMenu ()
{
	ASSERT_VALID (this);

	if (m_paletteButton.m_arIcons.GetSize () == 0)
	{
		return;
	}

	if (m_paletteButton.m_pPopupMenu != NULL || m_paletteButton.IsDroppedDown())
	{
		m_paletteButton.ClosePopupMenu ();
		return;
	}

	m_paletteButton.m_nPaletteID = GetDlgCtrlID();
	m_paletteButton.SetNotifyParentID(TRUE);
	m_paletteButton.m_pParentControl = this;
	m_paletteButton.m_bIsButtonMode = TRUE;
	m_paletteButton.m_pWndTargetCmd = GetParent();

	m_paletteButton.OnShowPopupMenu ();
}
//**************************************************************************************
void CBCGPRibbonComboGalleryCtrl::SetIcons(UINT uiImagesResID, int cxImage)
{
	ASSERT_VALID (this);
	ASSERT (cxImage != 0);

	m_paletteButton.Clear();

	if (uiImagesResID == 0)
	{
		return;
	}

	m_paletteButton.m_imagesPalette.Load (uiImagesResID);

	int cy = cxImage;
	
	BOOL bIsSVGList = m_paletteButton.m_imagesPalette.IsSVGList();

	if (!bIsSVGList)
	{
		HBITMAP hbmp = m_paletteButton.m_imagesPalette.GetImageWell();
		if (hbmp != NULL)
		{
			BITMAP bmp;
			GetObject(hbmp, sizeof (BITMAP), &bmp);
			
			cy = bmp.bmHeight;
		}
	}

	m_paletteButton.m_imagesPalette.SetImageSize (CSize (cxImage, cy), !bIsSVGList);

	const double dblScale = globalData.GetRibbonImageScale ();
	if (dblScale != 1.0)
	{
		m_paletteButton.m_imagesPalette.SetTransparentColor (globalData.clrBtnFace);
		m_paletteButton.m_imagesPalette.SmoothResize (dblScale);
	}
}
//**************************************************************************************
int CBCGPRibbonComboGalleryCtrl::AddString(LPCTSTR lpszString, int nIconIndex)
{
	ASSERT_VALID (this);
	ASSERT (lpszString != NULL);

	CString str = lpszString;
	CString strDescr;

	int nIndex = str.Find(_T('\n'));
	if (nIndex >= 0)
	{
		strDescr = str.Mid(nIndex + 1);
		str = str.Left(nIndex);
	}

	CBCGPRibbonPaletteIcon* pIcon = new CBCGPRibbonPaletteIcon (&m_paletteButton, m_paletteButton.m_nIcons++);

	pIcon->SetText(str);
	pIcon->m_nImageIndex = nIconIndex;

	if (!strDescr.IsEmpty() || m_bHasDescriptions)
	{
		pIcon->SetDescription(strDescr.IsEmpty() ? _T(" ") : strDescr);

		if (!m_bHasDescriptions)
		{
			for (int i = 0; i < m_paletteButton.m_arIcons.GetSize (); i++)
			{
				m_paletteButton.m_arIcons [i]->SetDescription(_T(" "));
			}

			m_bHasDescriptions = TRUE;
		}
	}

	m_paletteButton.m_arIcons.Add (pIcon);

	return m_paletteButton.m_nIcons - 1;
}
//**************************************************************************************
void CBCGPRibbonComboGalleryCtrl::AddLabel(LPCTSTR lpszString)
{
	ASSERT_VALID (this);
	ASSERT (lpszString != NULL);

	CBCGPRibbonLabel* pLabel = new CBCGPRibbonLabel (lpszString);

	pLabel->SetData ((DWORD_PTR) pLabel);
	pLabel->m_bIsPaletteGroup = TRUE;

	m_paletteButton.m_arIcons.Add (pLabel);
}
//*****************************************************************************************
void CBCGPRibbonComboGalleryCtrl::AddCommand(UINT uiCmdID, LPCTSTR lpszString, BOOL bOnTop)
{
	ASSERT_VALID (this);
	ASSERT (lpszString != NULL);

	CBCGPRibbonButton* pCommand = new CBCGPRibbonButton (uiCmdID, lpszString);
	ASSERT_VALID(pCommand);

	pCommand->m_pWndTargetCmd = GetParent();

	m_paletteButton.AddSubItem(pCommand, -1, bOnTop);
}
//*****************************************************************************************
void CBCGPRibbonComboGalleryCtrl::ResetContent()
{
	ASSERT_VALID (this);

	m_paletteButton.Clear ();
	m_paletteButton.RemoveAllSubItems();
	m_paletteButton.m_nSelected = -1;

	m_bHasDescriptions = FALSE;

	if (GetSafeHwnd() != NULL)
	{
		RedrawWindow();
	}
}
//*****************************************************************************************
void CBCGPRibbonComboGalleryCtrl::OnLButtonDown(UINT /*nFlags*/, CPoint /*point*/) 
{
	SetFocus ();
	OnShowMenu ();
}

#endif // BCGP_EXCLUDE_RIBBON
