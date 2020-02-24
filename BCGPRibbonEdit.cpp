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
// BCGPRibbonEdit.cpp: implementation of the CBCGPRibbonEdit class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BCGPRibbonEdit.h"
#include "BCGPVisualManager.h"
#include "bcgglobals.h"
#include "BCGPRibbonBar.h"
#include "BCGPRibbonPanel.h"
#include "BCGPToolbarComboBoxButton.h"
#include "trackmouse.h"
#include "BCGPPopupMenu.h"
#include "BCGPSpinButtonCtrl.h"
#include "BCGPDrawManager.h"
#include "BCGPRibbonCategory.h"
#include "BCGPRibbonPanelMenu.h"
#include "BCGPKeyboardManager.h"
#include "BCGPLocalResource.h"
#include "BCGPContextMenuManager.h"
#include "bcgprores.h"

#ifndef BCGP_EXCLUDE_RIBBON

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#if _MSC_VER < 1300

/////////////////////////////////////////////////////////////////////////////
// _BCGP_RICHEDIT_STATE

class _BCGP_RICHEDIT_STATE : public CNoTrackObject
{
public:
	_BCGP_RICHEDIT_STATE()
	{
		m_hInstRichEdit = NULL;
	}

	virtual ~_BCGP_RICHEDIT_STATE()
	{
		if (m_hInstRichEdit != NULL)
		{
			::FreeLibrary(m_hInstRichEdit);
		}
	}

	HINSTANCE m_hInstRichEdit;      // handle to RICHED20.DLL
};

PROCESS_LOCAL(_BCGP_RICHEDIT_STATE, _bcgpRichEditState)

BOOL PASCAL BCGPInitRichEdit20 ()
{
	_BCGP_RICHEDIT_STATE* pState = _bcgpRichEditState;

	if (pState->m_hInstRichEdit == NULL)
	{
		pState->m_hInstRichEdit = LoadLibraryA("RICHED20.DLL");
	}

	return pState->m_hInstRichEdit != NULL;
}

#else

BOOL PASCAL BCGPInitRichEdit20 ()
{
	return FALSE;
}

#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPCommandsMenu

class CBCGPCommandsMenu : public CBCGPRibbonPanelMenu
{
	DECLARE_DYNAMIC(CBCGPCommandsMenu)
		
public:
	CBCGPCommandsMenu(
		CBCGPRibbonBar* pRibbonBar,
		const CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>&	arButtons) :
	CBCGPRibbonPanelMenu(pRibbonBar, arButtons)
	{
		m_bAutoDestroyParent = FALSE;
	}
	
	virtual ~CBCGPCommandsMenu()
	{
		CBCGPRibbonCommandsComboBox* pParent = DYNAMIC_DOWNCAST(CBCGPRibbonCommandsComboBox, GetParentRibbonElement());
		if (pParent != NULL)
		{
			pParent->CleanUp();
		}
	}
	
	virtual BOOL IsDropListMode()
	{
		return TRUE;
	}
	
	virtual BOOL IsParentEditFocused()
	{
		CBCGPRibbonCommandsComboBox* pParent = DYNAMIC_DOWNCAST(CBCGPRibbonCommandsComboBox, GetParentRibbonElement());
		return pParent != NULL && pParent->HasFocus();
	}
	
	BOOL OnClickSelection()
	{
		CBCGPBaseRibbonElement* pFocused = m_wndRibbonBar.GetFocused();
		if (pFocused != NULL)
		{
			OnKeyDown(VK_RETURN, 0, 0);
			return TRUE;
		}
		
		return FALSE;
	}
};

IMPLEMENT_DYNAMIC(CBCGPCommandsMenu, CBCGPRibbonPanelMenu)

/////////////////////////////////////////////////////////////////////////////
// CBCGPClearButton

class CBCGPClearButton : public CBCGPButton
{
	friend class CBCGPRibbonEdit;

	CBCGPClearButton (CBCGPRibbonEdit* pEdit = NULL)
	{
		m_pEdit = pEdit;
		m_bDrawFocus = FALSE;

		SetMouseCursorHand();
		SetImageAutoScale();

		CBCGPLocalResource locaRes;
		SetImage(globalData.Is32BitIcons () ? IDB_BCGBARRES_CLEAR32 : IDB_BCGBARRES_CLEAR);
	}

	afx_msg void OnClicked();
	DECLARE_MESSAGE_MAP()

	virtual void OnFillBackground (CDC* pDC, const CRect& rectClient)
	{
		BOOL bIsHighlighted = m_bHighlighted;

		if (CWnd::GetFocus() == m_pEdit->m_pWndEdit)
		{
			m_bHighlighted = TRUE;
		}

		CBCGPVisualManager::GetInstance ()->OnDrawEditClearButton(pDC, this, rectClient);

		m_bHighlighted = bIsHighlighted;
	}

	virtual void OnDrawBorder (CDC*, CRect&, UINT)
	{
	}

	CBCGPRibbonEdit*		m_pEdit;
};

BEGIN_MESSAGE_MAP(CBCGPClearButton, CBCGPButton)
	ON_CONTROL_REFLECT(BN_CLICKED, OnClicked)
END_MESSAGE_MAP()

void CBCGPClearButton::OnClicked() 
{
	ASSERT_VALID (this);
	ASSERT_VALID(m_pEdit->m_pWndEdit);

	m_pEdit->m_pWndEdit->SetFocus();
	m_pEdit->SetEditText(_T(""));
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonSpinButtonCtrl

class CBCGPRibbonSpinButtonCtrl : public CBCGPSpinButtonCtrl
{
	friend class CBCGPRibbonEdit;

	CBCGPRibbonSpinButtonCtrl (CBCGPRibbonEdit*	pEdit = NULL)
	{
		m_bQuickAccessMode = FALSE;
		m_pEdit = pEdit;
	}

	virtual BOOL IsRibbonControl () const
	{
		return TRUE;
	}

	virtual void OnDraw (CDC* pDC)
	{
		BOOL bIsDrawOnGlass = CBCGPToolBarImages::m_bIsDrawOnGlass;

		if (m_bQuickAccessMode)
		{
			CBCGPRibbonBar* pRibbonBar = DYNAMIC_DOWNCAST (CBCGPRibbonBar, GetParent ());
			if (pRibbonBar != NULL)
			{
				ASSERT_VALID (pRibbonBar);

				if (pRibbonBar->IsQuickAccessToolbarOnTop () &&
					pRibbonBar->IsTransparentCaption ())
				{
					CBCGPToolBarImages::m_bIsDrawOnGlass = TRUE;
				}
			}
		}

		CBCGPSpinButtonCtrl::OnDraw (pDC);

		CBCGPToolBarImages::m_bIsDrawOnGlass = bIsDrawOnGlass;
	}

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	//{{AFX_MSG(CBCGPRibbonSpinButtonCtrl)
	afx_msg void OnDeltapos(NMHDR* pNMHDR, LRESULT* pResult);
	//}}AFX_MSG
	afx_msg LRESULT OnAfterDeltaPos(WPARAM, LPARAM);
	DECLARE_MESSAGE_MAP()

	BOOL				m_bQuickAccessMode;
	CBCGPRibbonEdit*	m_pEdit;
};

#define UM_AFTERDELTAPOS	(WM_USER + 1001)

BEGIN_MESSAGE_MAP(CBCGPRibbonSpinButtonCtrl, CBCGPSpinButtonCtrl)
	//{{AFX_MSG_MAP(CBCGPRibbonSpinButtonCtrl)
	ON_NOTIFY_REFLECT(UDN_DELTAPOS, OnDeltapos)
	//}}AFX_MSG_MAP
	ON_MESSAGE(UM_AFTERDELTAPOS, OnAfterDeltaPos)
END_MESSAGE_MAP()

void CBCGPRibbonSpinButtonCtrl::OnDeltapos(NMHDR* pNMHDR, LRESULT* pResult) 
{
	if (m_pEdit != NULL && GetBuddy ()->GetSafeHwnd () != NULL)
	{
		m_pEdit->OnDeltaPos((LPNMUPDOWN)pNMHDR);
		PostMessage(UM_AFTERDELTAPOS);
	}

	*pResult = 0;
}
//**************************************************************************
LRESULT CBCGPRibbonSpinButtonCtrl::OnAfterDeltaPos(WPARAM, LPARAM)
{
	if (m_pEdit != NULL && GetBuddy ()->GetSafeHwnd () != NULL)
	{
		CString str;
		GetBuddy ()->GetWindowText (str);
		
		m_pEdit->SetEditText (str);
		
		m_pEdit->m_RecentChangeEvt = CBCGPRibbonEdit::BCGPRIBBON_EDIT_CHANGED_BY_SPIN_BUTTON;
		m_pEdit->NotifyCommand (TRUE);
	}

	return 0;
}
//**************************************************************************
BOOL CBCGPRibbonSpinButtonCtrl::PreTranslateMessage(MSG* pMsg) 
{
	if (pMsg->message == WM_LBUTTONDOWN)
	{
		CBCGPRibbonEditCtrl* pWndCurrEdit = DYNAMIC_DOWNCAST(CBCGPRibbonEditCtrl, CWnd::GetFocus());
		
		if (pWndCurrEdit->GetSafeHwnd() != NULL)
		{
			pWndCurrEdit->OnKillFocus(GetBuddy());
		}
	}

	return CBCGPSpinButtonCtrl::PreTranslateMessage(pMsg);
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonEdit

IMPLEMENT_DYNCREATE(CBCGPRibbonEdit, CBCGPRibbonButton)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPRibbonEdit::CBCGPRibbonEdit(
	UINT	uiID,
	int		nWidth,
	LPCTSTR lpszLabel, 
	int		nImage)
{
	CommonInit ();

	m_nID = uiID;
	m_nWidth = nWidth;
	m_nWidthFloaty = nWidth;
	m_nSmallImageIndex = nImage;
	
	SetText (lpszLabel);
}
//**************************************************************************
CBCGPRibbonEdit::CBCGPRibbonEdit()
{
	CommonInit ();
}
//**************************************************************************
void CBCGPRibbonEdit::CommonInit ()
{
	m_bIsEditFocused = FALSE;
	m_nWidth = 0;
	m_nWidthFloaty = 0;
	m_pWndEdit = NULL;
	m_pWndSpin = NULL;
	m_pBtnClear = NULL;
	m_bForceDrawBorder = TRUE;
	m_bHasDropDownList = FALSE;
	m_bHasSpinButtons = FALSE;
	m_nMin = INT_MAX;
	m_nMax = INT_MAX;
	m_nAlign = ES_LEFT;
	m_szMargin = CSize (2, 3);
	m_nLabelImageWidth = 0;
	m_bNotifyCommand = TRUE;
	m_bSearchMode = FALSE;
	m_bDontScaleInHighDPI = FALSE;
	m_bIsAutoComplete = FALSE;
	m_clrCustomText = (COLORREF)-1;
	m_RecentChangeEvt = BCGPRIBBON_EDIT_NOT_CHANGED;
	m_clrLight = (COLORREF)-1;
	m_nMaxWidth = -1;
	m_nLimitText = 0;
	m_bIsReadOnly = FALSE;
}
//**************************************************************************
CBCGPRibbonEdit::~CBCGPRibbonEdit()
{
	DestroyCtrl ();
}
//**************************************************************************
CSize CBCGPRibbonEdit::GetIntermediateSize (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	int cx = m_bFloatyMode ? m_nWidthFloaty : m_nWidth;
	if (!m_bDontScaleInHighDPI && !m_bSearchResultMode && globalData.GetRibbonImageScale () > 1.)
	{
		cx = globalUtils.ScaleByDPI(cx);
	}

	if (m_nMaxWidth > 0)
	{
		cx = min(cx, m_nMaxWidth);
	}

	TEXTMETRIC tm;
	pDC->GetTextMetrics (&tm);

	int nTextHeight = tm.tmHeight;
	if ((nTextHeight % 2) != 0)
	{
		nTextHeight++;
	}

	if (m_pParent != NULL)
	{
		ASSERT_VALID (m_pParent);
		
		int nSmallButtonHeight = m_pParent->GetImageSize (FALSE).cy + 6;

		m_szMargin.cy = max (2, (nSmallButtonHeight - nTextHeight) / 2);
	}

	int cy = nTextHeight + 2 * m_szMargin.cy;
	
	m_nLabelImageWidth = 0;

	if (!m_bQuickAccessMode && !m_bFloatyMode)
	{
		CSize sizeImageSmall = GetImageSize (RibbonImageSmall);

		if (sizeImageSmall != CSize (0, 0))
		{
			m_nLabelImageWidth += sizeImageSmall.cx + 2 * m_szMargin.cx;
			cy = max (cy, sizeImageSmall.cy);
		}

		if (m_sizeTextRight.cx > 0)
		{
			m_nLabelImageWidth += m_sizeTextRight.cx + 2 * m_szMargin.cx;
			cy = max (cy, m_sizeTextRight.cy);
		}

		cx += m_nLabelImageWidth;
	}

	return CSize (cx, cy + m_sizePadding.cy - 1);
}
//**************************************************************************
CSize CBCGPRibbonEdit::GetCompactSize (CDC* pDC)
{
	ASSERT_VALID (this);

	int nLabelWidth = m_sizeTextRight.cx;
	m_sizeTextRight.cx = 0;

	CSize size = GetIntermediateSize (pDC);

	m_sizeTextRight.cx = nLabelWidth;
	return size;
}
//**************************************************************************
void CBCGPRibbonEdit::SetEditText (CString strText)
{
	ASSERT_VALID (this);

	if (m_strEdit != strText)
	{
		m_strEdit = strText;

		if (m_pWndEdit->GetSafeHwnd () != NULL)
		{
			m_pWndEdit->SetWindowText (m_strEdit);
		}

		Redraw ();
	}

	if (!m_bDontNotify)
	{
		CBCGPRibbonBar* pRibbonBar = GetTopLevelRibbonBar ();
		if (pRibbonBar != NULL)
		{
			ASSERT_VALID (pRibbonBar);
			
			CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arButtons;
			pRibbonBar->GetElementsByID (m_nID, arButtons, TRUE);
			
			for (int i = 0; i < arButtons.GetSize (); i++)
			{
				CBCGPRibbonEdit* pOther =
					DYNAMIC_DOWNCAST (CBCGPRibbonEdit, arButtons [i]);
				
				if (pOther != NULL && pOther != this)
				{
					ASSERT_VALID (pOther);
					
					pOther->m_bDontNotify = TRUE;
					pOther->SetEditText (strText);
					pOther->m_bDontNotify = FALSE;
					pOther->m_RecentChangeEvt = m_RecentChangeEvt;
				}
			}
		}
	}
}
//**************************************************************************
void CBCGPRibbonEdit::EnableSearchMode(BOOL bEnable, LPCTSTR lpszPrompt)
{
	ASSERT_VALID (this);

	m_bSearchMode = bEnable;

	if (m_bSearchMode)
	{
		ASSERT(lpszPrompt != NULL);
		m_strSearchPrompt = lpszPrompt;

		if (!m_ImageSearch.IsValid())
		{
			CBCGPLocalResource locaRes;
			m_ImageSearch.Load(globalData.Is32BitIcons () ?
				IDB_BCGBARRES_SEARCH32 : IDB_BCGBARRES_SEARCH);
			m_ImageSearch.SetSingleImage(FALSE);
			m_ImageSearch.SetTransparentColor(globalData.clrBtnFace);
			
			globalUtils.ScaleByDPI(m_ImageSearch);
		}
	}
	else
	{
		if (m_pBtnClear != NULL)
		{
			m_pBtnClear->DestroyWindow ();
			delete m_pBtnClear;
			m_pBtnClear = NULL;
		}
	}
}
//**************************************************************************
void CBCGPRibbonEdit::SetPrompt(LPCTSTR lpszPrompt)
{
	ASSERT_VALID (this);
	m_strSearchPrompt = lpszPrompt == NULL ? _T("") : lpszPrompt;
}
//**************************************************************************
void CBCGPRibbonEdit::EnableSpinButtons (int nMin, int nMax)
{
	ASSERT_VALID (this);

	m_nMin = nMin;
	m_nMax = nMax;

	m_bHasSpinButtons = TRUE;

	if (m_pWndSpin->GetSafeHwnd () != NULL)
	{
	   m_pWndSpin->SetRange32 (m_nMin, m_nMax);
	}
}
//**************************************************************************
void CBCGPRibbonEdit::OnDeltaPos(LPNMUPDOWN pUpDown)
{
	UNREFERENCED_PARAMETER(pUpDown);
}
//**************************************************************************
void CBCGPRibbonEdit::SetTextAlign (int nAlign)
{
	ASSERT_VALID (this);
	ASSERT (nAlign == ES_LEFT || nAlign == ES_CENTER || nAlign == ES_RIGHT);

	m_nAlign = nAlign;
}
//**************************************************************************
void CBCGPRibbonEdit::SetWidth (int nWidth, BOOL bInFloatyMode)
{
	ASSERT_VALID (this);

	if (bInFloatyMode)
	{
		m_nWidthFloaty = nWidth;
	}
	else
	{
		m_nWidth = nWidth;
	}
}
//**************************************************************************
void CBCGPRibbonEdit::SetMaxWidth(int nMaxWidth)
{
	ASSERT_VALID (this);
	m_nMaxWidth = nMaxWidth;
}
//**************************************************************************
void CBCGPRibbonEdit::SetLimitText(int nChars)
{
	ASSERT_VALID(this);
	ASSERT(nChars >= 0);

	m_nLimitText = nChars;

	if (m_pWndEdit->GetSafeHwnd() != 0)
	{
		m_pWndEdit->LimitText(m_nLimitText);
	}
}
//**************************************************************************
void CBCGPRibbonEdit::SetReadOnly(BOOL bSet/* = TRUE*/)
{
	ASSERT_VALID(this);

	m_bIsReadOnly = bSet;
	
	if (m_pWndEdit->GetSafeHwnd() != NULL)
	{
		if (globalData.bIsWindowsVista)
		{
			m_pWndEdit->SendMessage(EM_SETREADONLY, (WPARAM)bSet);
		}
		else
		{
			m_pWndEdit->ModifyStyle(m_bIsReadOnly ? 0 : ES_READONLY, m_bIsReadOnly ? ES_READONLY : 0);
		}
	}
}
//**************************************************************************
void CBCGPRibbonEdit::OnDraw (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	if (m_rect.IsRectEmpty())
	{
		return;
	}

	OnDrawLabelAndImage (pDC);

	BOOL bIsHighlighted = m_bIsHighlighted;

	if (m_bIsFocused)
	{
		m_bIsHighlighted = TRUE;
	}

	if (IsDisabled ())
	{
		m_bIsHighlighted = FALSE;
	}

	CRect rectSaved = m_rect;
	CRect rectCommandSaved = m_rectCommand;

	m_rect.bottom--;
	m_rectCommand.bottom--;

	int cx = m_bFloatyMode ? m_nWidthFloaty : m_nWidth;
	if (!m_bDontScaleInHighDPI && !m_bSearchResultMode && globalData.GetRibbonImageScale () > 1.)
	{
		cx = globalUtils.ScaleByDPI(cx);
	}

	if (m_nMaxWidth > 0)
	{
		cx = min(cx, m_nMaxWidth);
	}

	if (m_bIsFixedWidth && m_nLabelImageWidth == 0)
	{
		cx = m_rect.Width();
	}

	m_rectCommand.left = m_rect.left = m_rect.right - cx;

	CBCGPVisualManager::GetInstance ()->OnFillRibbonButton (pDC, this);

	if (m_pWndEdit->GetSafeHwnd () == NULL)
	{
		CRect rectText = m_rectCommand;
		rectText.DeflateRect (m_szMargin);

		UINT uiDTFlags = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX;

		if (!m_bSearchResultMode)
		{
			if (m_nAlign == ES_CENTER)
			{
				uiDTFlags |= DT_CENTER;
			}
			else if (m_nAlign == ES_RIGHT)
			{
				uiDTFlags |= DT_RIGHT;
			}
		}

		DoDrawText (pDC, m_strEdit, rectText, uiDTFlags);
	}

	if (m_bSearchMode && m_strEdit.IsEmpty() && m_ImageSearch.IsValid())
	{
		CRect rectIcon = m_rect;
		rectIcon.left = rectIcon.right - rectIcon.Height();

		m_ImageSearch.DrawEx(pDC, rectIcon, 0, CBCGPToolBarImages::ImageAlignHorzCenter, CBCGPToolBarImages::ImageAlignVertCenter,
			CRect(0, 0, 0, 0), IsDisabled() ? (BYTE)127 : (BYTE)255);
	}

	if (IsCommandsCombo())
	{
		CRect rectIcon = m_rect;
		rectIcon.right = rectIcon.left + rectIcon.Height();

		COLORREF clr = CBCGPVisualManager::GetInstance ()->GetRibbonEditPromptColor(m_pWndEdit, FALSE, FALSE);
		if (clr != m_clrLight || !m_ImageLight.IsValid())
		{
			m_clrLight = clr;

			CBCGPLocalResource locaRes;

			m_ImageLight.Clear();

			m_ImageLight.Load(IDB_BCGBARRES_LIGHT32);
			m_ImageLight.SetSingleImage(FALSE);
			m_ImageLight.AdaptColors(RGB(0, 0, 0), m_clrLight, TRUE);
			
			globalUtils.ScaleByDPI(m_ImageLight);
		}

		m_ImageLight.DrawEx(pDC, rectIcon, 0, CBCGPToolBarImages::ImageAlignHorzCenter, CBCGPToolBarImages::ImageAlignVertCenter);
	}

	CBCGPVisualManager::GetInstance ()->OnDrawRibbonButtonBorder 
			(pDC, this);

	m_bIsHighlighted = bIsHighlighted;
	m_rect = rectSaved;
	m_rectCommand = rectCommandSaved;
}
//****************************************************************************
void CBCGPRibbonEdit::OnDrawLabelAndImage (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	if (m_bQuickAccessMode || m_bFloatyMode)
	{
		return;
	}

	CSize sizeImageSmall = GetImageSize (RibbonImageSmall);

	int x = m_rect.left;

	if (sizeImageSmall != CSize (0, 0))
	{
		CRect rectImage = m_rect;
		rectImage.left += m_szMargin.cx;
		rectImage.right = rectImage.left + sizeImageSmall.cx;
		rectImage.OffsetRect (0, max (0, (rectImage.Height () - sizeImageSmall.cy) / 2));

		DrawImage (pDC, RibbonImageSmall, rectImage);

		x = rectImage.right;
	}

	if (m_sizeTextRight.cx > 0 && !m_bCompactMode)
	{
		COLORREF clrTextOld = (COLORREF)-1;

		if (IsDisabled ())
		{
			clrTextOld = pDC->SetTextColor (
				CBCGPVisualManager::GetInstance ()->GetToolbarDisabledTextColor ());
		}

		CRect rectText = m_rect;
		rectText.left = x + m_szMargin.cx;

		UINT uiDTFlags = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX;

		if (!m_bSearchResultMode)
		{
			int cx = (m_bFloatyMode ? m_nWidthFloaty : m_nWidth) + 2 * m_szMargin.cx;
			if (!m_bDontScaleInHighDPI && globalData.GetRibbonImageScale () > 1.)
			{
				cx = globalUtils.ScaleByDPI(cx);
			}

			if (m_nMaxWidth > 0)
			{
				cx = min(cx, m_nMaxWidth);
			}

			cx += m_rectMenu.Width();

			if (m_nAlign == ES_CENTER)
			{
				uiDTFlags |= DT_CENTER;
				rectText.right -= cx;

				if (sizeImageSmall == CSize (0, 0))
				{
					rectText.left -= m_szMargin.cx / 2;
				}
			}
			else if (m_nAlign == ES_RIGHT)
			{
				uiDTFlags |= DT_RIGHT;
				rectText.right -= cx;

				if (sizeImageSmall == CSize (0, 0))
				{
					rectText.left -= m_szMargin.cx;
				}
			}
		}

		DoDrawText (pDC, m_strText, rectText, uiDTFlags);

		if (clrTextOld != (COLORREF)-1)
		{
			pDC->SetTextColor (clrTextOld);
		}
	}
}
//****************************************************************************
void CBCGPRibbonEdit::OnDrawOnList (CDC* pDC, CString strText,
									  int nTextOffset, CRect rect,
									  BOOL /*bIsSelected*/,
									  BOOL /*bHighlighted*/)
{
	const int nEditWidth = rect.Height () * 2;

	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	BOOL bIsDisabled = m_bIsDisabled;
	m_bIsDisabled = FALSE;

	CRect rectImage = rect;
	rectImage.right = rect.left + nTextOffset;

	CSize sizeImageSmall = GetImageSize (RibbonImageSmall);
	if (sizeImageSmall != CSize (0, 0))
	{
		rectImage.DeflateRect (1, 0);
		rectImage.top += max (0, (rectImage.Height () - sizeImageSmall.cy) / 2);
		rectImage.bottom = rectImage.top + sizeImageSmall.cy;

		DrawImage (pDC, RibbonImageSmall, rectImage);
	}

	CRect rectEdit = rect;
	rectEdit.left = rectEdit.right - nEditWidth;
	rectEdit.DeflateRect (1, 1);

	CRect rectText = rect;
	rectText.left += nTextOffset;
	rectText.right = rectEdit.left;
	const int xMargin = globalUtils.ScaleByDPI(3);
	rectText.DeflateRect (xMargin, 0);

	pDC->DrawText (strText, rectText, DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);

	CRect rectSaved = m_rect;
	CRect rectCommandSaved = m_rectCommand;
	
	m_rect = rectEdit;
	m_rectCommand = m_rect;
	m_rectCommand.right -= globalUtils.ScaleByDPI(15);

	CBCGPVisualManager::GetInstance ()->OnFillRibbonButton (pDC, this);
	CBCGPVisualManager::GetInstance ()->OnDrawRibbonButtonBorder 
		(pDC, this);

	if (m_bHasDropDownList)
	{
		CBCGPToolbarComboBoxButton buttonDummy;
		buttonDummy.m_bIsRibbon = TRUE;
		CRect rectDropButton = rectEdit;
		rectDropButton.left = m_rectCommand.right;
		rectDropButton.DeflateRect (1, 1);

		CBCGPVisualManager::GetInstance ()->OnDrawComboDropButton (
			pDC, rectDropButton, FALSE, FALSE, FALSE, &buttonDummy);
	}
	else
	{
		CRect rectCaret = rectEdit;
		rectCaret.DeflateRect (3, 3);
		rectCaret.bottom--;

		rectCaret.right = rectCaret.left + 7;

		CPen* pOldPen = (CPen*) pDC->SelectStockObject (BLACK_PEN);

		pDC->MoveTo (rectCaret.left, rectCaret.top);
		pDC->LineTo (rectCaret.right, rectCaret.top);

		pDC->MoveTo (rectCaret.CenterPoint ().x, rectCaret.top);
		pDC->LineTo (rectCaret.CenterPoint ().x, rectCaret.bottom);

		pDC->MoveTo (rectCaret.left, rectCaret.bottom);
		pDC->LineTo (rectCaret.right, rectCaret.bottom);

		pDC->SelectObject (pOldPen);
	}

	m_rect = rectSaved;
	m_rectCommand = rectCommandSaved;
	m_bIsDisabled = bIsDisabled;
}
//**************************************************************************
void CBCGPRibbonEdit::OnLButtonDown (CPoint point)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::OnLButtonDown (point);
}
//**************************************************************************
void CBCGPRibbonEdit::OnLButtonUp (CPoint /*point*/)
{
	ASSERT_VALID (this);
}
//**************************************************************************
void CBCGPRibbonEdit::CopyFrom (const CBCGPBaseRibbonElement& s)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::CopyFrom (s);

	CBCGPRibbonEdit& src = (CBCGPRibbonEdit&) s;

	m_strEdit = src.m_strEdit;

	DestroyCtrl();

	m_nWidth = src.m_nWidth;
	m_nWidthFloaty = src.m_nWidthFloaty;
	m_bHasSpinButtons = src.m_bHasSpinButtons;
	m_bHasDropDownList = src.m_bHasDropDownList;
	m_nMin = src.m_nMin;
	m_nMax = src.m_nMax;
	m_nAlign = src.m_nAlign;
	m_nLabelImageWidth = src.m_nLabelImageWidth;
	m_bSearchMode = src.m_bSearchMode;
	m_strSearchPrompt = src.m_strSearchPrompt;
	src.m_ImageSearch.CopyTo(m_ImageSearch);
	m_bDontScaleInHighDPI = src.m_bDontScaleInHighDPI;
	m_bIsAutoComplete = src.m_bIsAutoComplete;
	m_nLimitText = src.m_nLimitText;
	m_bIsReadOnly = src.m_bIsReadOnly;
}
//**************************************************************************
void CBCGPRibbonEdit::OnAfterChangeRect (CDC* pDC)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pDC);

	CBCGPRibbonButton::OnAfterChangeRect (pDC);

	ReposEditCtrl ();
}
//**************************************************************************
void CBCGPRibbonEdit::ReposEditCtrl ()
{
	if (m_rect.IsRectEmpty ())
	{
		if (m_pWndEdit->GetSafeHwnd () != NULL)
		{
			m_pWndEdit->ShowWindow (SW_HIDE);
		}

		if (m_pWndSpin->GetSafeHwnd () != NULL)
		{
			m_pWndSpin->ShowWindow (SW_HIDE);
		}

		if (m_pBtnClear->GetSafeHwnd () != NULL)
		{
			m_pBtnClear->ShowWindow (SW_HIDE);
		}

		return;
	}

	CRect rectCommandOld = m_rectCommand;

	m_Location = RibbonElementSingleInGroup;
	m_rectCommand = m_rect;

	if (m_pWndEdit == NULL)
	{
		DWORD dwEditStyle = WS_CHILD | ES_WANTRETURN | 
							ES_AUTOHSCROLL | WS_TABSTOP;

		dwEditStyle |= m_nAlign;

		CWnd* pWndParent = GetParentWnd ();
		ASSERT_VALID (pWndParent);

		if ((m_pWndEdit = CreateEdit (pWndParent, dwEditStyle)) == NULL)
		{
			return;
		}

		m_pWndEdit->SendMessage (EM_SETTEXTMODE, TM_PLAINTEXT);
		m_pWndEdit->SetEventMask (m_pWndEdit->GetEventMask () | ENM_CHANGE);
		m_pWndEdit->SetFont (GetTopLevelRibbonBar ()->GetFont ());
		m_pWndEdit->SetWindowText (m_strEdit);
	}
	
	if (rectCommandOld != m_rectCommand || !m_pWndEdit->IsWindowVisible ())
	{
		CRect rectEdit = m_rectCommand;
		rectEdit.DeflateRect(0, m_sizePadding.cy / 2);

		int cx = m_bFloatyMode ? m_nWidthFloaty : m_nWidth;
		if (!m_bDontScaleInHighDPI && !m_bSearchResultMode && globalData.GetRibbonImageScale () > 1.)
		{
			cx = globalUtils.ScaleByDPI(cx);
		}

		if (m_nMaxWidth > 0)
		{
			cx = min(cx, m_nMaxWidth);
		}

		if (m_bIsFixedWidth && m_nLabelImageWidth == 0)
		{
			cx = m_rect.Width();
		}

		rectEdit.left = rectEdit.right - cx;

		if (m_bHasSpinButtons)
		{
			rectEdit.DeflateRect (m_szMargin.cx, m_szMargin.cy, 2, m_szMargin.cy);
		}
		else
		{
			rectEdit.DeflateRect (m_szMargin.cx, m_szMargin.cy);
		}

		if (m_bSearchMode)
		{
			rectEdit.right -= m_rect.Height();
		}
		else if (IsCommandsCombo())
		{
			rectEdit.left += m_rect.Height();
		}

		m_pWndEdit->SetWindowPos (NULL, 
			rectEdit.left, rectEdit.top,
			rectEdit.Width (), rectEdit.Height (),
			SWP_NOZORDER | SWP_NOACTIVATE);

		m_pWndEdit->ShowWindow (SW_SHOWNOACTIVATE);

		if (m_pWndSpin->GetSafeHwnd () != NULL)
		{
			m_pWndSpin->m_bQuickAccessMode = m_bQuickAccessMode;
			m_pWndSpin->SetBuddy (m_pWndEdit);
			m_pWndSpin->ShowWindow (SW_SHOWNOACTIVATE);
		}

		if (m_bSearchMode && m_pBtnClear->GetSafeHwnd () != NULL)
		{
			const int nButtonSize = m_rect.Height () - 3;

			m_pBtnClear->SetWindowPos (NULL, 
				m_rect.right - m_rect.Height() + 2, m_rect.top + 1,
				nButtonSize, nButtonSize,
				SWP_NOZORDER | SWP_NOACTIVATE);

			m_pBtnClear->ShowWindow (m_strEdit.IsEmpty() ? SW_HIDE : SW_SHOWNOACTIVATE);
		}
	}
}
//****************************************************************************************
CBCGPRibbonEditCtrl* CBCGPRibbonEdit::CreateEdit (CWnd* pWndParent, DWORD dwEditStyle)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pWndParent);

	BOOL bIsDefaultCreate = FALSE;
	CBCGPRibbonEditCtrl* pWndEdit = NULL;

	if (!BCGPInitRichEdit20 ())
	{
		bIsDefaultCreate = TRUE;
	}

	if (m_bIsReadOnly)
	{
		dwEditStyle |= ES_READONLY;
	}

	if (!bIsDefaultCreate)
	{
		CWnd* pWnd = new CBCGPRibbonEditCtrl (*this);

	#ifdef UNICODE
		CString strClassName = L"RichEdit20W";
	#else
		CString strClassName = "RichEdit20A";
	#endif

		if (!pWnd->Create (strClassName, NULL, dwEditStyle, CRect (0, 0, 0, 0), 
								pWndParent, m_nID))
		{
			delete pWnd;
			return NULL;
		}

		pWndEdit = (CBCGPRibbonEditCtrl*)pWnd;
	}
	else
	{
		pWndEdit = new CBCGPRibbonEditCtrl (*this);

		if (!pWndEdit->Create (dwEditStyle, CRect (0, 0, 0, 0), 
							pWndParent, m_nID))
		{
			delete pWndEdit;
			return NULL;
		}
	}

	if (m_bHasSpinButtons)
	{
		CreateSpinButton (pWndEdit, pWndParent);
	}

	if (m_bSearchMode && m_pBtnClear == NULL)
	{
		m_pBtnClear = new CBCGPClearButton (this);
		m_pBtnClear->Create(_T(""), WS_CHILD, CRect (0, 0, 0, 0), pWndParent, (UINT)-1);

		if (m_bQuickAccessMode)
		{
			CBCGPRibbonBar* pRibbonBar = DYNAMIC_DOWNCAST (CBCGPRibbonBar, pWndParent);
			if (pRibbonBar != NULL)
			{
				ASSERT_VALID (pRibbonBar);

				if (pRibbonBar->IsQuickAccessToolbarOnTop () &&
					pRibbonBar->IsTransparentCaption ())
				{
					m_pBtnClear->m_bOnGlass = TRUE;
				}
			}
		}
	}

	if (m_nLimitText > 0)
	{
		pWndEdit->LimitText(m_nLimitText);
	}

	AddTooltip(pWndEdit);
	return pWndEdit;
}
//********************************************************************************
void CBCGPRibbonEdit::OnUpdateToolTips()
{
	ASSERT_VALID(this);
	AddTooltip(m_pWndEdit);
}
//********************************************************************************
BOOL CBCGPRibbonEdit::CreateSpinButton (CBCGPRibbonEditCtrl* pWndEdit, CWnd* pWndParent)
{
	ASSERT_VALID (this);
	ASSERT (m_pWndSpin == NULL);

	if (!m_bHasSpinButtons)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	m_pWndSpin = new CBCGPRibbonSpinButtonCtrl (this);

	if (!m_pWndSpin->Create(
		WS_CHILD | WS_VISIBLE | UDS_ALIGNRIGHT | UDS_ARROWKEYS | UDS_SETBUDDYINT,
		CRect (0, 0, 0, 0), pWndParent, m_nID))
	{
		delete m_pWndSpin;
		return FALSE;
	}

   m_pWndSpin->SetBuddy (pWndEdit);
   m_pWndSpin->SetRange32 (m_nMin, m_nMax);
   m_pWndSpin->m_bVisualManagerStyle = TRUE;

   return TRUE;
}
//********************************************************************************
void CBCGPRibbonEdit::OnShow (BOOL bShow)
{
	ASSERT_VALID (this);

	if (m_pWndEdit->GetSafeHwnd () != NULL)
	{
		m_pWndEdit->ShowWindow (bShow ? SW_SHOWNOACTIVATE : SW_HIDE);
	}

	if (m_pWndSpin->GetSafeHwnd () != NULL)
	{
		m_pWndSpin->ShowWindow (bShow ? SW_SHOWNOACTIVATE : SW_HIDE);
	}

	if (m_pBtnClear->GetSafeHwnd () != NULL)
	{
		m_pBtnClear->ShowWindow (bShow && !m_strEdit.IsEmpty() ? SW_SHOWNOACTIVATE : SW_HIDE);
	}
}
//********************************************************************************
void CBCGPRibbonEdit::Redraw ()
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::Redraw ();

	if (m_pWndEdit->GetSafeHwnd () != NULL &&
		m_pWndEdit->IsWindowVisible ())
	{
		m_pWndEdit->RedrawWindow ();
	}

	if (m_pWndSpin->GetSafeHwnd () != NULL &&
		m_pWndSpin->IsWindowVisible ())
	{
		m_pWndSpin->RedrawWindow ();
	}

	if (m_pBtnClear->GetSafeHwnd () != NULL &&
		m_pBtnClear->IsWindowVisible ())
	{
		m_pBtnClear->RedrawWindow ();
	}
}
//********************************************************************************
void CBCGPRibbonEdit::OnHighlight (BOOL bHighlight)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::OnHighlight (bHighlight);

	if (m_pWndEdit->GetSafeHwnd () != NULL &&
		m_pWndEdit->IsWindowVisible ())
	{
		m_pWndEdit->m_bIsHighlighted = bHighlight;
		m_pWndEdit->RedrawWindow ();
	}
}
//********************************************************************************
void CBCGPRibbonEdit::OnSetFocus (BOOL bSet)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::OnSetFocus (bSet);

	if (m_pWndEdit->GetSafeHwnd () != NULL &&
		m_pWndEdit->IsWindowVisible ())
	{
		if (bSet)
		{
			m_pWndEdit->SetFocus ();
			m_pWndEdit->SetSel (0, -1);
		}
		else
		{
			CBCGPRibbonBar* pRibbonBar = DYNAMIC_DOWNCAST (CBCGPRibbonBar, m_pWndEdit->GetParent ());
			if (pRibbonBar != NULL)
			{
				ASSERT_VALID (pRibbonBar);
				
				pRibbonBar->m_bDontSetKeyTips = TRUE;
				pRibbonBar->SetFocus ();
			}
			else
			{
				CBCGPRibbonPanelMenuBar* pMenuBar = DYNAMIC_DOWNCAST (CBCGPRibbonPanelMenuBar, m_pWndEdit->GetParent ());
				if (pMenuBar != NULL)
				{
					ASSERT_VALID (pMenuBar);

					if (pMenuBar->GetParent () != NULL)
					{
						pMenuBar->GetParent ()->SetFocus ();
					}
				}

				m_bIsEditFocused = FALSE;
				m_pWndEdit->SetSel (0, 0);
				Redraw ();
			}
		}

		m_pWndEdit->RedrawWindow ();
	}
}
//********************************************************************************
BOOL CBCGPRibbonEdit::PreLMouseDown (CPoint point)
{
	ASSERT_VALID (this);

	if (m_rect.PtInRect (point) || !m_bIsEditFocused || m_bIsFocused)
	{
		return FALSE;
	}

	OnSetFocus (FALSE);
	return TRUE;
}
//********************************************************************************
void CBCGPRibbonEdit::OnEnable (BOOL bEnable)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::OnEnable (bEnable);

	if (m_pWndEdit->GetSafeHwnd () != NULL)
	{
		m_pWndEdit->EnableWindow (bEnable);
	}

	if (m_pWndSpin->GetSafeHwnd () != NULL)
	{
		m_pWndSpin->EnableWindow (bEnable);
	}

	if (m_pBtnClear->GetSafeHwnd () != NULL)
	{
		m_pBtnClear->EnableWindow (bEnable);
	}
}
//*************************************************************************************
BOOL CBCGPRibbonEdit::OnKey (BOOL bIsMenuKey)
{
	ASSERT_VALID (this);

	if (m_bIsDisabled)
	{
		return FALSE;
	}

	if (m_rect.IsRectEmpty ())
	{
		return CBCGPBaseRibbonElement::OnKey (bIsMenuKey);
	}

	CBCGPRibbonBar* pTopLevelRibbon = GetTopLevelRibbonBar ();
	if (pTopLevelRibbon != NULL)
	{
		pTopLevelRibbon->HideKeyTips ();
	}

	if (bIsMenuKey)
	{
		DropDownList ();
		return TRUE;
	}

	if (m_pWndEdit->GetSafeHwnd () != NULL)
	{
		m_pWndEdit->SetFocus ();

		if (!m_strEdit.IsEmpty ())
		{
			m_pWndEdit->SetSel (0, -1);
		}

		return TRUE;
	}

	DropDownList ();
	return TRUE;
}
//*************************************************************************************
void CBCGPRibbonEdit::OnRTLChanged (BOOL bIsRTL)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::OnRTLChanged (bIsRTL);

	m_rectCommand.SetRectEmpty ();
	m_rectMenu.SetRectEmpty ();

	if (!m_rect.IsRectEmpty () && m_pWndEdit->GetSafeHwnd () != NULL && m_pWndEdit->IsWindowVisible ())
	{
		CClientDC dc (GetParentWnd ());
		OnAfterChangeRect (&dc);
	}
}
//*************************************************************************************
void CBCGPRibbonEdit::AddTooltip(CBCGPRibbonEditCtrl* pWndEdit)
{
	ASSERT_VALID(this);

	if (pWndEdit->GetSafeHwnd() == NULL)
	{
		return;
	}

	CWnd* pWndParent = pWndEdit->GetParent();
	CToolTipCtrl* pWndToolTip = NULL;

	CBCGPRibbonBar* pRibbonBar = DYNAMIC_DOWNCAST (CBCGPRibbonBar, pWndParent);
	if (pRibbonBar != NULL)
	{
		ASSERT_VALID (pRibbonBar);
		pWndToolTip = pRibbonBar->m_pToolTip;
	}
	else
	{
		CBCGPRibbonPanelMenuBar* pMenuBar = DYNAMIC_DOWNCAST (CBCGPRibbonPanelMenuBar, pWndParent);
		if (pMenuBar != NULL)
		{
			ASSERT_VALID(pMenuBar);
			pWndToolTip = pMenuBar->m_pToolTip;
		}
	}

	if (pWndToolTip->GetSafeHwnd() != NULL && ::IsWindow(pWndToolTip->GetSafeHwnd()))
	{
		pWndToolTip->AddTool(pWndEdit);
	}
}
//*************************************************************************************
void CBCGPRibbonEdit::DelTooltip()
{
	ASSERT_VALID(this);

	CToolTipCtrl* pWndToolTip = NULL;

	if (m_pWndEdit->GetSafeHwnd() != NULL)
	{
		CBCGPRibbonBar* pRibbonBar = DYNAMIC_DOWNCAST (CBCGPRibbonBar, m_pWndEdit->GetParent());
		if (pRibbonBar != NULL)
		{
			ASSERT_VALID (pRibbonBar);
			pWndToolTip = pRibbonBar->m_pToolTip;
		}
		else
		{
			CBCGPRibbonPanelMenuBar* pMenuBar = DYNAMIC_DOWNCAST (CBCGPRibbonPanelMenuBar, m_pWndEdit->GetParent());
			if (pMenuBar != NULL)
			{
				ASSERT_VALID(pMenuBar);
				pWndToolTip = pMenuBar->m_pToolTip;
			}
		}
	}

	if (pWndToolTip->GetSafeHwnd() != NULL && ::IsWindow(pWndToolTip->GetSafeHwnd()))
	{
		pWndToolTip->DelTool(m_pWndEdit);
	}
}
//*************************************************************************************
void CBCGPRibbonEdit::DestroyCtrl ()
{
	ASSERT_VALID (this);

	if (m_pWndEdit != NULL)
	{
		m_pWndEdit->DestroyWindow ();
		delete m_pWndEdit;
		m_pWndEdit = NULL;
	}

	if (m_pWndSpin != NULL)
	{
		m_pWndSpin->DestroyWindow ();
		delete m_pWndSpin;
		m_pWndSpin = NULL;
	}

	if (m_pBtnClear != NULL)
	{
		m_pBtnClear->DestroyWindow ();
		delete m_pBtnClear;
		m_pBtnClear = NULL;
	}

	m_clrCustomText = (COLORREF)-1;
}
//*************************************************************************************
void CBCGPRibbonEdit::OnChangeRibbonFont()
{
	ASSERT_VALID (this);
	
	if (m_pWndEdit->GetSafeHwnd() != NULL && GetTopLevelRibbonBar() != NULL)
	{
		m_pWndEdit->SetFont(GetTopLevelRibbonBar()->GetFont ());
	}
}
//*************************************************************************************
BOOL CBCGPRibbonEdit::SetACCData (CWnd* pParent, CBCGPAccessibilityData& data)
{
	ASSERT_VALID (this);

	CBCGPRibbonButton::SetACCData (pParent, data);

	data.m_strAccValue = m_strEdit;

	if (!IsMenuAreaHighlighted ())
	{
		data.m_nAccRole = ROLE_SYSTEM_TEXT;
	}
	else
	{
		 data.m_bAccState = 0;
		 data.m_nAccRole = ROLE_SYSTEM_PUSHBUTTON;
	}

	if (IsFocused ())
	{
		 data.m_bAccState |= STATE_SYSTEM_FOCUSED;
	}

	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonEditCtrl

IMPLEMENT_DYNAMIC(CBCGPRibbonEditCtrl, CRichEditCtrl)

CBCGPRibbonEditCtrl::CBCGPRibbonEditCtrl(CBCGPRibbonEdit& edit) :
	m_edit (edit)
{
	m_bTracked			= FALSE;
	m_bIsHighlighted	= FALSE;
	m_bIsContextMenu	= FALSE;

#if _MSC_VER >= 1300
	EnableActiveAccessibility();
#endif
}

CBCGPRibbonEditCtrl::~CBCGPRibbonEditCtrl()
{
}

BEGIN_MESSAGE_MAP(CBCGPRibbonEditCtrl, CRichEditCtrl)
	//{{AFX_MSG_MAP(CBCGPRibbonEditCtrl)
	ON_WM_SETFOCUS()
	ON_WM_KILLFOCUS()
	ON_WM_PAINT()
	ON_WM_CONTEXTMENU()
	ON_CONTROL_REFLECT_EX(EN_CHANGE, OnChange)
	ON_WM_ENABLE()
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_MOUSELEAVE, OnMouseLeave)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonEditCtrl message handlers

BOOL CBCGPRibbonEditCtrl::PreTranslateMessage(MSG* pMsg) 
{
	if (pMsg->message == WM_ERASEBKGND)
	{
		return TRUE;
	}

	if (m_bIsContextMenu)
	{
		return CRichEditCtrl::PreTranslateMessage(pMsg);
	}

	if (pMsg->message == WM_LBUTTONDOWN)
	{
		if (m_edit.m_pParentMenu != NULL && g_pContextMenuManager != NULL && g_pContextMenuManager->IsTracking())
		{
			g_pContextMenuManager->CancelTracking();
			return TRUE;
		}

		if (!m_edit.m_bIsEditFocused)
		{
			m_edit.OnSetFocus (TRUE);

			if (CBCGPPopupMenu::GetActiveMenu ()->GetSafeHwnd () != NULL)
			{
				if (m_edit.m_pParentMenu == NULL ||
					m_edit.m_pParentMenu->GetParent ()->GetSafeHwnd () != CBCGPPopupMenu::GetActiveMenu ()->GetSafeHwnd ())
				{
					return FALSE;
				}
			}

			return !m_edit.m_bFloatyMode;
		}
		else
		{
			if (!m_edit.m_bIsFocused && pMsg->hwnd != GetSafeHwnd ())
			{
				m_edit.OnSetFocus (FALSE);
			}
		}
	}

	if (pMsg->message == WM_MOUSEMOVE && !m_edit.IsDisabled ())
	{
		if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0 && GetFocus() != this)
		{
			return TRUE;
		}

		if (!m_bTracked)
		{
			m_bTracked = TRUE;
			
			TRACKMOUSEEVENT trackmouseevent;
			trackmouseevent.cbSize = sizeof(trackmouseevent);
			trackmouseevent.dwFlags = TME_LEAVE;
			trackmouseevent.hwndTrack = GetSafeHwnd ();
			trackmouseevent.dwHoverTime = HOVER_DEFAULT;

			::BCGPTrackMouse (&trackmouseevent);

			RedrawWindow ();
		}

		if (!m_bIsHighlighted && m_edit.GetParentWnd () != NULL)
		{
			CPoint point;
			::GetCursorPos (&point);

			m_bIsHighlighted = TRUE;
			RedrawWindow ();

			CBCGPRibbonPanel* pPanel = m_edit.GetParentPanel ();
			
			if (pPanel == NULL || !pPanel->IsMainPanel())
			{
				m_edit.GetParentWnd ()->ScreenToClient (&point);
				m_edit.GetParentWnd ()->SendMessage (WM_MOUSEMOVE, 0, MAKELPARAM (point.x, point.y));
			}
		}
	}

	if (pMsg->message == WM_KEYDOWN && !m_edit.IsDisabled ())
	{
		if (m_edit.OnProcessKey ((UINT)pMsg->wParam))
		{
			return TRUE;
		}

		if (ProcessClipboardAccelerators ((UINT) pMsg->wParam))
		{
			return TRUE;
		}

		switch (pMsg->wParam)
		{
		case VK_TAB:
			if (!m_edit.IsFocused ())
			{
				return TRUE;
			}
			break;

		case VK_DOWN:
			if (m_edit.m_bHasDropDownList && !m_edit.IsDroppedDown ())
			{
				m_edit.DropDownList ();
				return TRUE;
			}

		case VK_UP:
		case VK_PRIOR:
		case VK_NEXT:
			if (m_edit.IsDroppedDown () || m_edit.IsCommandsCombo())
			{
				::SendMessage (	CBCGPPopupMenu::GetActiveMenu ()->GetSafeHwnd (),
								WM_KEYDOWN, pMsg->wParam, pMsg->lParam);
				return TRUE;
			}
			break;

		case VK_LEFT:
		case VK_RIGHT:
		case VK_HOME:
		case VK_END:
			if (m_edit.IsCommandsCombo() && CBCGPPopupMenu::GetSafeActivePopupMenu() != NULL)
			{
				::SendMessage (	CBCGPPopupMenu::GetActiveMenu ()->GetSafeHwnd (),
					WM_KEYDOWN, pMsg->wParam, pMsg->lParam);
				return TRUE;
			}
			break;

		case VK_RETURN:
			if (m_edit.IsCommandsCombo())
			{
				if (DYNAMIC_DOWNCAST(CBCGPCommandsMenu, CBCGPPopupMenu::GetSafeActivePopupMenu()) != NULL)
				{
					return TRUE;
				}
			}
			else
			{
				BOOL bUseEditText = TRUE;

				if (m_edit.IsDroppedDown())
				{
					CBCGPDropDownList* pList = DYNAMIC_DOWNCAST(CBCGPDropDownList, m_edit.m_pPopupMenu);
					if (pList != NULL)
					{
						ASSERT_VALID(pList);
						if (pList->GetCurSel() >= 0)
						{
							bUseEditText = FALSE;
						}
					}
				}

				if (bUseEditText)
				{
					if (m_edit.IsDroppedDown())
					{
						m_edit.ClosePopupMenu();
					}

					CString str;
					GetWindowText (str);

					m_edit.SetEditText (str);
					m_edit.m_RecentChangeEvt = CBCGPRibbonEdit::BCGPRIBBON_EDIT_CHANGED_BY_ENTER_KEY;
					m_edit.NotifyCommand (TRUE);

					if (m_edit.m_pParentMenu != NULL &&
						m_edit.m_pParentMenu->GetParent ()->GetSafeHwnd () == CBCGPPopupMenu::GetActiveMenu ()->GetSafeHwnd ())
					{
						ASSERT_VALID (m_edit.m_pParentMenu);

						CFrameWnd* pParentFrame = BCGPGetParentFrame (m_edit.m_pParentMenu);
						ASSERT_VALID (pParentFrame);

						m_edit.m_bNotifyCommand = FALSE;
						pParentFrame->DestroyWindow ();
						return TRUE;
					}

					if (GetTopLevelFrame () != NULL)
					{
						m_edit.m_bNotifyCommand = FALSE;
						GetTopLevelFrame ()->SetFocus ();
						return TRUE;
					}
				}
			}
			break;

		case VK_F4:
			if (m_edit.m_bHasDropDownList && !m_edit.IsDroppedDown ())
			{
				m_edit.DropDownList ();
				return TRUE;
			}
			else if (m_edit.IsDroppedDown () && CBCGPPopupMenu::GetActiveMenu () != NULL)
			{
				CBCGPPopupMenu::GetActiveMenu ()->SendMessage (WM_CLOSE);
				return TRUE;
			}
			break;

		case VK_ESCAPE:
			if ((m_edit.IsDroppedDown () || m_edit.IsCommandsCombo()) && CBCGPPopupMenu::GetActiveMenu () != NULL)
			{
				CBCGPPopupMenu::GetActiveMenu ()->SendMessage (WM_CLOSE);
				return TRUE;
			}

			if (!m_edit.IsDroppedDown ())
			{
				SetWindowText (m_strOldText);
				m_edit.SetEditText (m_strOldText);
			}

			if (GetTopLevelFrame () != NULL && !m_edit.IsDroppedDown ())
			{
				GetTopLevelFrame ()->SetFocus ();
				return TRUE;
			}
			break;

		default:
			if (m_edit.m_bIsAutoComplete && CBCGPKeyboardManager::IsKeyPrintable ((UINT)pMsg->wParam))
			{
				long lStart = -1;
				long lEnd = -1;

				GetSel(lStart, lEnd);

				if (lStart >= 0)
				{
					CString strText = m_edit.m_strEdit.Left(lStart);
					strText += (TCHAR)CBCGPKeyboardManager::TranslateCharToUpper((UINT)pMsg->wParam);

					CString str;
					
					if (m_edit.FindItemByPrefix(strText, str) >= 0)
					{
						m_edit.m_bDontNotify = TRUE;

						SetWindowText(str);
						m_edit.SetEditText (str);
						SetSel(lStart + 1, -1);

						m_edit.m_bDontNotify = FALSE;

						return TRUE;
					}
				}
			}
		}
	}

	return CRichEditCtrl::PreTranslateMessage(pMsg);
}
//*********************************************************************************
BOOL CBCGPRibbonEditCtrl::ProcessClipboardAccelerators (UINT nChar)
{
	if (m_edit.IsDroppedDown() && (nChar == VK_LEFT || nChar == VK_RIGHT || nChar == VK_HOME || nChar == VK_END))
	{
		return FALSE;
	}

	BOOL bChanged = FALSE;

	if (!globalUtils.ProcessCtrlEditAccelerators(this, nChar, IsWindowEnabled() && !m_edit.m_bIsReadOnly, &bChanged))
	{
		return FALSE;
	}

	if (bChanged && m_edit.m_bSearchMode)
	{
		OnChange();
	}

	return TRUE;
}
//*************************************************************************************
BOOL CBCGPRibbonEditCtrl::OnChange()
{
	CString strText;
	GetWindowText (strText);

	m_edit.m_strEdit = strText;

	if (m_edit.m_bSearchMode && m_edit.m_pBtnClear->GetSafeHwnd () != NULL && IsWindowVisible())
	{
		m_edit.m_pBtnClear->ShowWindow (strText.IsEmpty() ? SW_HIDE : SW_SHOWNOACTIVATE);
	}

	return m_edit.OnEditChange ();
}
//*************************************************************************************
void CBCGPRibbonEditCtrl::OnSetFocus(CWnd* pOldWnd) 
{
	CRichEditCtrl::OnSetFocus(pOldWnd);
	
	m_edit.m_bIsEditFocused = TRUE;
	m_edit.m_bNotifyCommand = TRUE;
	m_edit.Redraw ();

	GetWindowText (m_strOldText);
}
//*************************************************************************************
void CBCGPRibbonEditCtrl::OnKillFocus(CWnd* pNewWnd) 
{
	CRichEditCtrl::OnKillFocus(pNewWnd);

	if (m_edit.IsCommandsCombo())
	{
		if (((CBCGPRibbonCommandsComboBox&)m_edit).CloseDropDownMenu())
		{
			SetWindowText(_T(""));
		}
	}

	m_edit.m_bIsEditFocused = FALSE;
	m_edit.Redraw ();

	SetSel (0, 0);
	RedrawWindow ();

	GetWindowText (m_strOldText);
	m_edit.m_strEdit = m_strOldText;

	if (m_edit.m_bIsAutoComplete)
	{
		m_edit.SetEditText(m_strOldText);
	}
	else if (pNewWnd->GetSafeHwnd() != NULL && m_edit.m_pParentMenu != NULL &&
		m_edit.m_pParentMenu->GetParent ()->GetSafeHwnd () == CBCGPPopupMenu::GetActiveMenu ()->GetSafeHwnd () &&
		::IsChild(pNewWnd->GetSafeHwnd(), GetSafeHwnd()))
	{
		m_edit.SetEditText(m_strOldText);
	}

	HWND hwndClear = m_edit.m_pBtnClear->GetSafeHwnd();

	if (m_edit.m_bNotifyCommand && (hwndClear == NULL || pNewWnd->GetSafeHwnd() != hwndClear))
	{
		m_edit.m_RecentChangeEvt = CBCGPRibbonEdit::BCGPRIBBON_EDIT_CHANGED_BY_FOCUS;
		m_edit.NotifyCommand(TRUE); 
		m_edit.m_bNotifyCommand = TRUE;
	}
}
//*************************************************************************************
void CBCGPRibbonEditCtrl::OnPaint() 
{
	CRect rect;
	GetClientRect (rect);

	CBCGPRibbonPanel* pPanel = m_edit.GetParentPanel ();

	BOOL bIsHighlighted = m_edit.IsHighlighted () || m_edit.IsDroppedDown () || m_edit.IsFocused () || m_bIsHighlighted;
	BOOL bIsDisabled = !IsWindowEnabled();

	COLORREF clrBackground = CBCGPVisualManager::GetInstance ()->GetRibbonEditBackgroundColor (
		this, bIsHighlighted, pPanel != NULL && pPanel->IsHighlighted (), bIsDisabled);

	COLORREF clrText = CBCGPVisualManager::GetInstance ()->GetRibbonEditTextColor(this, bIsHighlighted, bIsDisabled);
	if (clrText != m_edit.m_clrCustomText)
	{
		CHARFORMAT cf;
		memset(&cf, 0, sizeof(CHARFORMAT));

		cf.cbSize = sizeof(CHARFORMAT);
		cf.dwMask = CFM_COLOR;
		cf.crTextColor = (clrText == (COLORREF)-1) ? globalData.clrWindowText : clrText;

		SetDefaultCharFormat(cf);

		m_edit.m_clrCustomText = clrText;
	}

	if (!m_edit.m_strSearchPrompt.IsEmpty() && m_edit.m_strEdit.IsEmpty() && !m_edit.m_bIsEditFocused)
	{
		if (IsWindowEnabled())
		{
			CPaintDC dc (this);
			
			dc.SetTextColor (CBCGPVisualManager::GetInstance()->GetRibbonEditPromptColor(this, bIsHighlighted, bIsDisabled));
			dc.SetBkMode(TRANSPARENT);
			
			CBCGPRibbonBar* pTopLevelRibbon = m_edit.GetTopLevelRibbonBar();

			CFont* pOldFont = dc.SelectObject(pTopLevelRibbon != NULL ? pTopLevelRibbon->GetFont() : &globalData.fontRegular);

			CBCGPDrawManager dm(dc);
			dm.DrawRect(rect, clrBackground, (COLORREF)-1);

			CRect rectText = rect;
			rectText.DeflateRect(m_edit.m_szMargin.cx + 1, 0);

			UINT uiDTFlags = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;

			if (m_edit.m_nAlign == ES_LEFT)
			{
				uiDTFlags |= DT_LEFT;
			}
			else if (m_edit.m_nAlign == ES_CENTER)
			{
				uiDTFlags |= DT_CENTER;
			}
			else if (m_edit.m_nAlign == ES_RIGHT)
			{
				uiDTFlags |= DT_RIGHT;
			}

			m_edit.DoDrawText(&dc, m_edit.m_strSearchPrompt, rectText, uiDTFlags);

			dc.SelectObject (pOldFont);
		}
		else
		{
			SetBackgroundColor (FALSE, clrBackground);
			Default ();
		}
	}
	else
	{
		SetBackgroundColor (FALSE, clrBackground);
		Default ();
	}

	if (m_edit.m_pRibbonBar != NULL)
	{
		ASSERT_VALID (m_edit.m_pRibbonBar);

		if (m_edit.m_pRibbonBar->IsTransparentCaption ())
		{
			if ((m_edit.m_bQuickAccessMode && m_edit.m_pRibbonBar->IsQuickAccessToolbarOnTop ()) ||
				(m_edit.m_bIsTabElement && m_edit.m_pRibbonBar->AreTransparentTabs ()))
			{
				CClientDC dc (this);

				CBCGPDrawManager dm (dc);
				dm.FillAlpha (rect);
			}
		}
	}
}
//*************************************************************************************
LRESULT CBCGPRibbonEditCtrl::OnMouseLeave(WPARAM,LPARAM)
{
	if (m_edit.GetParentWnd () != NULL)
	{
		m_edit.GetParentWnd ()->SendMessage (WM_MOUSELEAVE);
	}

	if (m_bIsHighlighted)
	{
		m_bIsHighlighted = FALSE;
		RedrawWindow ();
	}

	m_bTracked = FALSE;

	return 0;
}
//*****************************************************************************************
void CBCGPRibbonEditCtrl::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	CBCGPPopupMenu* pActivePopupMenu = CBCGPPopupMenu::GetActiveMenu();
	HWND hwndActivePopup = pActivePopupMenu->GetSafeHwnd();

	if (pActivePopupMenu != NULL)
	{
		if (m_edit.m_pParentMenu == NULL ||
			m_edit.m_pParentMenu->GetParent ()->GetSafeHwnd () != CBCGPPopupMenu::GetActiveMenu ()->GetSafeHwnd ())
		{
			return;
		}

		if (g_pContextMenuManager != NULL)
		{
			g_pContextMenuManager->SetDontCloseActiveMenu();
		}
	}

	CBCGPRibbonBar* pRibbonBar = m_edit.GetTopLevelRibbonBar ();
	if (pRibbonBar != NULL)
	{
		ASSERT_VALID (pRibbonBar);

		m_bIsContextMenu = TRUE;

		pRibbonBar->OnEditContextMenu (this, point);

		m_bIsContextMenu = FALSE;
	}

	if (hwndActivePopup != NULL && pActivePopupMenu != NULL)
	{
		if (::IsWindow(hwndActivePopup))
		{
			CBCGPPopupMenu::m_pActivePopupMenu = pActivePopupMenu;
		}

		if (g_pContextMenuManager != NULL)
		{
			g_pContextMenuManager->SetDontCloseActiveMenu(FALSE);
		}
	}
}
//*************************************************************************************
void CBCGPRibbonEditCtrl::OnEnable(BOOL bEnable) 
{
	if (!bEnable)
	{
		SetBackgroundColor (FALSE, 
			CBCGPVisualManager::GetInstance ()->GetRibbonEditBackgroundColor(this, FALSE, FALSE, TRUE));
		return;
	}

	CRichEditCtrl::OnEnable(bEnable);
}
//*************************************************************************************
HRESULT CBCGPRibbonEditCtrl::get_accName(VARIANT varChild, BSTR *pszName)
{
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		return m_edit.get_accName(varChild, pszName);
	}

	return S_FALSE;
}
//*************************************************************************************
HRESULT CBCGPRibbonEditCtrl::get_accRole(VARIANT varChild, VARIANT *pvarRole)
{
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		pvarRole->vt = VT_I4;
		pvarRole->lVal = ROLE_SYSTEM_TEXT;

		return S_OK;
	}

	return S_FALSE;
}
//*************************************************************************************
HRESULT CBCGPRibbonEditCtrl::get_accValue(VARIANT varChild, BSTR *pszValue)
{
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		return  m_edit.get_accValue(varChild, pszValue);
	}

	return S_FALSE;
}
//*************************************************************************************
HRESULT CBCGPRibbonEditCtrl::get_accKeyboardShortcut(VARIANT varChild, BSTR* pszKeyboardShortcut)
{
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		return  m_edit.get_accKeyboardShortcut(varChild, pszKeyboardShortcut);
	}

	return S_FALSE;
}
//*************************************************************************************
HRESULT CBCGPRibbonEditCtrl::get_accDescription(VARIANT varChild, BSTR FAR* pszDescription)
{
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		return  m_edit.get_accDescription(varChild, pszDescription);
	}

	return S_FALSE;
}
//*************************************************************************************
void CBCGPRibbonEditCtrl::OnDestroy() 
{
	m_edit.DelTooltip();
	CRichEditCtrl::OnDestroy();
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonCommandsComboBox

IMPLEMENT_DYNCREATE(CBCGPRibbonCommandsComboBox, CBCGPRibbonEdit)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPRibbonCommandsComboBox::CBCGPRibbonCommandsComboBox(
	UINT	uiID,
	LPCTSTR lpszPrompt,
	int		nWidth) :
	CBCGPRibbonEdit(uiID, nWidth)
{
	if (lpszPrompt != NULL)
	{
		m_strSearchPrompt = lpszPrompt;
	}

	m_hWndCommandsMenu = NULL;
}
//**************************************************************************
CBCGPRibbonCommandsComboBox::CBCGPRibbonCommandsComboBox()
{
}
//**************************************************************************
void CBCGPRibbonCommandsComboBox::CleanUp()
{
	ASSERT_VALID(this);

	for (int i = 0; i < m_arSearchResultsFiltered.GetSize(); i++)
	{
		ASSERT_VALID(m_arSearchResultsFiltered[i]);
		delete m_arSearchResultsFiltered[i];
	}

	m_arSearchResultsFiltered.RemoveAll();
	m_arSearchResultsRecent.RemoveAll();

	CloseDropDownMenu();
	m_bIsDroppedDown = FALSE;
}
//**************************************************************************
BOOL CBCGPRibbonCommandsComboBox::CloseDropDownMenu()
{
	BOOL bRes = FALSE;

	if (m_hWndCommandsMenu != NULL)
	{
		if (::IsWindow(m_hWndCommandsMenu))
		{
			CBCGPPopupMenu* pPopupMenu = DYNAMIC_DOWNCAST(CBCGPPopupMenu, CWnd::FromHandlePermanent(m_hWndCommandsMenu));
			if (pPopupMenu != NULL)
			{
				pPopupMenu->SendMessage(WM_CLOSE);
				bRes = TRUE;
			}
		}
		
		m_hWndCommandsMenu = NULL;
	}
	else
	{
		bRes = TRUE;
	}

	return bRes;
}
//**************************************************************************
CBCGPRibbonCommandsComboBox::~CBCGPRibbonCommandsComboBox()
{
	CleanUp();
}
//**************************************************************************
BOOL CBCGPRibbonCommandsComboBox::OnEditChange()
{
	ASSERT_VALID(m_pRibbonBar);

	m_pRibbonBar->PopTooltip();

	CBCGPCommandsMenu* pCommandsMenu = NULL;
	if (m_hWndCommandsMenu != NULL && ::IsWindow(m_hWndCommandsMenu))
	{
		pCommandsMenu = DYNAMIC_DOWNCAST(CBCGPCommandsMenu, CWnd::FromHandlePermanent(m_hWndCommandsMenu));
	}

	CStringArray arWords;

	int nIndex = 0;
	CString strWord;

	while (AfxExtractSubString(strWord, m_strEdit, nIndex, _T(' ')))
	{
		if (!strWord.IsEmpty())
		{
			strWord.MakeLower();
			arWords.Add(strWord);
		}

		++nIndex;
	}

	CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*> arSearchResults;

	for (int nStep = 0; nStep < 3; nStep++)
	{
		if (m_pRibbonBar->QueryElements(arWords, arSearchResults, m_pRibbonBar->GetCommandSearchOptions().m_nMaxResults, nStep == 2, nStep == 0))
		{
			break;
		}
	}

	int i = 0;
	BOOL bChanged = FALSE;

	if (m_arSearchResultsRecent.GetSize() == arSearchResults.GetSize())
	{
		for (i = 0; i < arSearchResults.GetSize(); i++)
		{
			if (m_arSearchResultsRecent[i] != arSearchResults[i])
			{
				bChanged = TRUE;
				break;
			}
		}
	}
	else
	{
		bChanged = TRUE;
	}

	if (!bChanged && pCommandsMenu->GetSafeHwnd() != NULL && !m_strEdit.IsEmpty())
	{
		return FALSE;
	}

	CleanUp();

	if (arWords.GetSize() == 0)
	{
		return FALSE;
	}

	for (i = 0; i < arSearchResults.GetSize (); i++)
	{
		CBCGPBaseRibbonElement* pResElement = arSearchResults[i];
		ASSERT_VALID(pResElement);

		m_arSearchResultsRecent.Add(pResElement);
		
		if (pResElement->IsHiddenInAppMode() || !m_pRibbonBar->OnFilterSearchResult(pResElement))
		{
			continue;
		}

		BOOL bAlreadyAdded = FALSE;
		for (int j = 0; j < m_arSearchResultsFiltered.GetSize(); j++)
		{
			if (m_arSearchResultsFiltered[j]->GetID() == pResElement->GetID())
			{
				bAlreadyAdded = TRUE;
				break;
			}
		}
		
		if (bAlreadyAdded)
		{
			continue;
		}

		CString strLabel = pResElement->m_strText;
		if (strLabel.IsEmpty ())
		{
			strLabel = pResElement->m_strToolTip;
		}

		CBCGPBaseRibbonElement* pElement = pResElement->CreateCustomCopy();
		if (pElement != NULL)
		{
			ASSERT_VALID(pElement);

			pElement->SetDefaultMenuLook(TRUE);
			pElement->m_pRibbonBar = m_pRibbonBar;
			pElement->m_pOriginal = pResElement;

			// Remove accelerators:
			globalUtils.RemoveSingleAmp(pElement->m_strText, FALSE);
		
			m_arSearchResultsFiltered.Add (pElement);
		}
	}

	CWnd* pWndParent = m_pRibbonBar->GetParent();
	if (pWndParent != NULL)
	{
		CBCGPRibbonCommandsMenuCustomItems items;
		items.m_strInput = m_strEdit;

		pWndParent->SendMessage(BCGM_ON_GET_RIBBON_COMMANDS_MENU_CUSTOM_ITEMS, 0, (LPARAM)&items);

		if (items.m_arCustomItems.GetSize() > 0)
		{
			if (m_arSearchResultsFiltered.GetSize() > 0)
			{
				m_arSearchResultsFiltered.Add(new CBCGPRibbonSeparator(TRUE));
			}

			for (i = 0; i < items.m_arCustomItems.GetSize(); i++)
			{
				CBCGPBaseRibbonElement* pElement = items.m_arCustomItems[i];
				ASSERT_VALID(pElement);

				pElement->m_pRibbonBar = m_pRibbonBar;
				pElement->m_bIsTemporaryPopupItem = TRUE;
				m_arSearchResultsFiltered.Add(pElement);
			}
		}
	}

	if (m_arSearchResultsFiltered.GetSize() == 0)
	{
		return FALSE;
	}

	pCommandsMenu = new CBCGPCommandsMenu(m_pRibbonBar, m_arSearchResultsFiltered);

	pCommandsMenu->SetParentRibbonElement (this);
	pCommandsMenu->SetMenuMode ();
	pCommandsMenu->SetDefaultMenuLook ();

	CBCGPRibbonPanel* pPanel = pCommandsMenu->GetPanel();
	if (pPanel != NULL)
	{
		for (i = 0; i < pPanel->GetCount(); i++)
		{
			CBCGPBaseRibbonElement* pElem = pPanel->GetElement(i);
			ASSERT_VALID(pElem);

			pElem->SetDefaultMenuLook();
			pElem->m_bSearchResultMode = TRUE;
			pElem->SetCommandSearchMenu();

			pPanel->SetInvokeCommandBySpace(FALSE);
		}
	}

	pCommandsMenu->EnableCustomizeMenu (FALSE);

	CPoint point = CPoint(m_rect.left, m_rect.bottom + 1);
	m_pRibbonBar->ClientToScreen(&point);

	if (pCommandsMenu->Create(GetParentWnd(), point.x, point.y, (HMENU) NULL))
	{
		pCommandsMenu->SendMessage (WM_KEYDOWN, VK_DOWN);
		m_bIsDroppedDown = TRUE;

		m_hWndCommandsMenu = pCommandsMenu->GetSafeHwnd();
	}

	return FALSE;
}
//****************************************************************************************************
BOOL CBCGPRibbonCommandsComboBox::OnProcessKey(UINT nChar)
{
	if (nChar == VK_RETURN)
	{
		CBCGPCommandsMenu* pCommandsMenu = DYNAMIC_DOWNCAST(CBCGPCommandsMenu, CBCGPPopupMenu::GetSafeActivePopupMenu());
		if (pCommandsMenu != NULL)
		{
			ASSERT_VALID(pCommandsMenu);

			if (pCommandsMenu->OnClickSelection())
			{
				return TRUE;
			}
		}
	}

	return CBCGPRibbonEdit::OnProcessKey(nChar);
}

#endif // BCGP_EXCLUDE_RIBBON
