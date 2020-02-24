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
// BCGPEdit.cpp : implementation file
//

#include "stdafx.h"
#include "bcgcbpro.h"
#include "bcgglobals.h"

#ifndef _BCGSUITE_
#include "BCGPToolBarImages.h"
#include "BCGPShellManager.h"
#include "BCGPTooltipManager.h"
#include "BCGPDropDownList.h"
#include "BCGPToolbarComboBoxButton.h"
#include "BCGPPopupMenu.h"
#include "BCGPPopupMenuBar.h"
#endif

#include "BCGPVisualManager.h"
#include "BCGPDialog.h"
#include "BCGPPopupDlg.h"
#include "BCGPDlgImpl.h"
#include "BCGProRes.h"
#include "BCGPLocalResource.h"
#include "BCGPCalculator.h"
#include "BCGPEdit.h"
#include "BCGPDrawManager.h"
#include "BCGPGlobalUtils.h"
#include "BCGPComboBox.h"
#include "BCGPMessageBox.h"

#ifndef _BCGSUITE_
	#define visualManagerMFC	CBCGPVisualManager::GetInstance ()
#else
	#define visualManagerMFC	CMFCVisualManager::GetInstance ()
	#define CBCGPDropDownList	CMFCDropDownListBox
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define UM_REDRAWFRAME		(WM_USER + 1001)
#define ID_TOOLTIP_TEXT		1
#define ID_TOOLTIP_BUTTON	2

#ifndef _BCGSUITE_
UINT BCGM_EDIT_ON_FILL_AUTOCOMPLETE_LIST = ::RegisterWindowMessage (_T("BCGM_EDIT_ON_FILL_AUTOCOMPLETE_LIST"));

/////////////////////////////////////////////////////////////////////////////
// CBCGPEditDropDownList

IMPLEMENT_DYNAMIC(CBCGPEditDropDownList, CBCGPDropDownList)

const UINT idStart = (UINT) -200;

CBCGPEditDropDownList::CBCGPEditDropDownList(CBCGPEdit* pEdit) :
	CBCGPDropDownList(pEdit)
{
	m_pEdit = pEdit;
}
//*****************************************************************************
void CBCGPEditDropDownList::OnChooseItem (UINT uidCmdID)
{
	CBCGPDropDownList::OnChooseItem (uidCmdID);

	if (m_pEdit->GetSafeHwnd() != NULL)
	{
		ASSERT_VALID (m_pEdit);

		int nIndex = (int)idStart - uidCmdID;

		CString strItemText;
		GetText(nIndex, strItemText);

		m_pEdit->m_bInAutoComplete = TRUE;
		m_pEdit->SetWindowText(strItemText);
		m_pEdit->SendMessage(WM_KEYDOWN, VK_END);
		m_pEdit->m_bInAutoComplete = FALSE;
	}
}
//*****************************************************************************
BOOL CBCGPEditDropDownList::Compare(const CStringList& lstStrings) const
{
	if (GetCount() != lstStrings.GetCount())
	{
		return FALSE;
	}

	int i = 0;

	for (POSITION pos = lstStrings.GetHeadPosition (); pos != NULL; i++)
	{
		CString strItemText;
		GetText(i, strItemText);

		if (lstStrings.GetNext(pos) != strItemText)
		{
			return FALSE;
		}
	}

	return TRUE;
}
//*****************************************************************************
CFont* CBCGPEditDropDownList::GetMenuFont()
{
	CFont* pFont = CBCGPDropDownList::GetMenuFont();
	
	if (m_pEdit != NULL)
	{
		ASSERT_VALID(m_pEdit);
		pFont = m_pEdit->GetFont();
	}

	return pFont;
}

#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPEdit

IMPLEMENT_DYNAMIC(CBCGPEdit, CEdit)

CBCGPEdit::CBCGPEdit()
{
	m_nVerticalAlignment = TA_TOP;
	m_nNCTop = m_nNCBottom = 0;
	m_sizeBorder = CSize(-1, -1);

	m_rectBtn.SetRectEmpty ();
	m_bIsButtonPressed = FALSE;
	m_bIsButtonHighlighted = FALSE;
	m_bIsButtonCaptured = FALSE;
	m_Mode = BrowseMode_None;
	m_dwFileDialogFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
	m_sizeImage = CSize (0, 0);
	m_pCalcPopup = NULL;
	m_pToolTip = NULL;
	m_bShowToolTip = FALSE;
	m_bDefaultPrintClient = FALSE;
	m_bDisableBrowseButtonInReadOnlyMode = FALSE;
	
	m_nBrowseButtonWidth = globalUtils.ScaleByDPI(20);

	m_bDefaultImage = TRUE;
	m_bVisualManagerStyle = FALSE;
	m_bOnGlass = FALSE;
	m_bSearchMode = FALSE;
	m_bHasPrompt = FALSE;
	m_bTextIsEmpty = TRUE;
	m_clrPrompt = (COLORREF)-1;
	m_clrErrorText = (COLORREF)-1;

#ifndef _BCGSUITE_
	m_pDropDownPopup = NULL;
	m_bInAutoComplete = FALSE;
#endif

	m_pRTIPopupDlg = NULL;
	m_lpszPopupDlgTemplateName = NULL;
	m_pPopupDlg = NULL;
	m_bIsResizablePopup = FALSE;
	m_bIsRightAlignedPopup = FALSE;
	m_bComboBoxMode = FALSE;
	m_nFolderBrowseFlags = BIF_RETURNONLYFSDIRS;
	m_rectClick.SetRectEmpty();

	m_pAutoCompleteCombo = NULL;
	m_bIsLastPrintableCharacter = FALSE;
	m_bIsDelayedLayout = FALSE;

	m_bHasSpinButton = FALSE;
	
	m_bAllowEditingInPasswordPreview = FALSE;
	m_cPassword = _T('*');

#ifndef _BCGSUITE_
	m_bSimplifiedBrowseIcon = FALSE;
#endif
}

CBCGPEdit::~CBCGPEdit()
{
}

BEGIN_MESSAGE_MAP(CBCGPEdit, CEdit)
	//{{AFX_MSG_MAP(CBCGPEdit)
	ON_WM_LBUTTONUP()
	ON_WM_MOUSEMOVE()
	ON_WM_NCCALCSIZE()
	ON_WM_NCPAINT()
	ON_WM_NCLBUTTONDBLCLK()
	ON_WM_NCMOUSEMOVE()
	ON_WM_CANCELMODE()
	ON_WM_LBUTTONDOWN()
	ON_WM_PAINT()
	ON_CONTROL_REFLECT_EX(EN_CHANGE, OnChange)
	ON_WM_CTLCOLOR_REFLECT()
	ON_WM_DESTROY()
	ON_WM_KILLFOCUS()
	ON_WM_KEYDOWN()
	ON_WM_SETCURSOR()
	ON_WM_SETFOCUS()
	ON_WM_CREATE()
	ON_WM_NCHITTEST()
	//}}AFX_MSG_MAP
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLVMMODE, OnBCGSetControlVMMode)
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLAERO, OnBCGSetControlAero)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnNeedTipText)
	ON_MESSAGE(WM_PRINTCLIENT, OnPrintClient)
	ON_MESSAGE(WM_PRINT, OnPrint)
	ON_MESSAGE(UM_REDRAWFRAME, OnRedrawFrame)
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPEdit message handlers

void CBCGPEdit::OnLButtonUp(UINT nFlags, CPoint point) 
{
	if (m_bIsButtonCaptured)
	{
		ReleaseCapture ();

		m_bIsButtonPressed = FALSE;
		m_bIsButtonCaptured = FALSE;
		m_bIsButtonHighlighted = FALSE;

		if (m_Mode == BrowseMode_PasswordPreview)
		{
			if (m_bAllowEditingInPasswordPreview)
			{
				SetPasswordChar(m_cPassword);
			}
			else
			{
				ShowCaret();
			}
		}

		RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE);

		if (m_rectClick.PtInRect (point))
		{
			OnBrowse ();
		}

		return;
	}
	
	CEdit::OnLButtonUp(nFlags, point);
}
//*************************************************************************************
void CBCGPEdit::OnMouseMove(UINT nFlags, CPoint point) 
{
	if (m_bIsButtonCaptured)
	{
		BOOL bIsButtonPressed = m_rectClick.PtInRect (point);
		if (bIsButtonPressed != m_bIsButtonPressed)
		{
			m_bIsButtonPressed = bIsButtonPressed;
			RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE);
		}

		return;
	}
	
	if (m_bIsButtonHighlighted)
	{
		if (!m_rectClick.PtInRect (point))
		{
			m_bIsButtonHighlighted = FALSE;
			ReleaseCapture ();

			RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE);
		}
	}
	else if (m_bComboBoxMode)
	{
		if (m_rectClick.PtInRect(point) && IsBrowseButtonEnabled())
		{
			m_bIsButtonHighlighted = TRUE;
			SetCapture();

			RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE);
		}
	}
	
	CEdit::OnMouseMove(nFlags, point);
}
//*************************************************************************************
void CBCGPEdit::OnNcCalcSize(BOOL bCalcValidRects, NCCALCSIZE_PARAMS FAR* lpncsp) 
{
	CEdit::OnNcCalcSize(bCalcValidRects, lpncsp);

	CRect rectWindow;
	GetWindowRect(rectWindow);

	CRect rectClient;
	GetClientRect(rectClient);
	ClientToScreen(&rectClient);

	m_sizeBorder.cx = max(0, rectClient.left - rectWindow.left);
	m_sizeBorder.cy = max(0, rectClient.top - rectWindow.top - m_nNCTop);

	if (m_Mode != BrowseMode_None)
	{
		lpncsp->rgrc[0].right -= m_nBrowseButtonWidth;
	}
	else if (m_bHasSpinButton)
	{
		lpncsp->rgrc[0].right -= 2;
	}

	m_nNCBottom = m_nNCTop = 0;

	if (m_nVerticalAlignment == TA_TOP || (GetStyle() & ES_MULTILINE) != 0 || m_bComboBoxMode)
	{
		return;
	}

	CClientDC dc(this);
	CBCGPFontSelector fs(dc, GetFont());
	
	TEXTMETRIC tm;
	dc.GetTextMetrics(&tm);
	
	int nTextHeight = tm.tmHeight + tm.tmInternalLeading;

	int cy = max(0, rectWindow.Height() - nTextHeight - 2 * m_sizeBorder.cy);
	
	m_nNCTop = m_nVerticalAlignment == TA_CENTER ? cy / 2 : cy;
	m_nNCBottom = m_nVerticalAlignment == TA_CENTER ? cy / 2 : 0;
	
	lpncsp->rgrc[0].top += m_nNCTop;
	lpncsp->rgrc[0].bottom -= m_nNCBottom;
}
//*************************************************************************************
void CBCGPEdit::OnNcPaint() 
{
	const BOOL bHasScrollBars = ((GetStyle () & WS_VSCROLL) == WS_VSCROLL) || ((GetStyle () & WS_HSCROLL) == WS_HSCROLL);
	const BOOL bHasBorder = (GetExStyle () & WS_EX_CLIENTEDGE) || (GetStyle () & WS_BORDER);

	const COLORREF clrBorderCustom = (GetFocus() == this && m_ColorTheme.m_clrBorderFocused != (COLORREF)-1) ? 
		m_ColorTheme.m_clrBorderFocused : m_ColorTheme.m_clrBorder;

	if (bHasScrollBars || (!m_bVisualManagerStyle && !m_bOnGlass && clrBorderCustom == (COLORREF)-1))
	{
		CEdit::OnNcPaint ();
	}

	if (bHasBorder && (m_bVisualManagerStyle || m_bOnGlass))
	{
		BOOL bIsRTL = (GetExStyle () & WS_EX_RTLREADING);

		CWindowDC dc(this);

		CRect rect;
		GetWindowRect(rect);

		rect.OffsetRect(-rect.TopLeft());

		if (m_bHasSpinButton && !bIsRTL)
		{
			rect.right -= 2;
		}

		if (clrBorderCustom != (COLORREF)-1)
		{
			CBCGPDrawManager dm(dc);
			dm.DrawRect(rect, (COLORREF)-1, clrBorderCustom);
			
			rect.DeflateRect (1, 1);
			
			COLORREF clrInner = m_ColorTheme.m_clrBackground == (COLORREF)-1 ? globalData.clrWindow : m_ColorTheme.m_clrBackground;
			dm.DrawRect(rect, (COLORREF)-1, clrInner);
		}
		else
		{
#ifndef _BCGSUITE_
			visualManagerMFC->OnDrawControlBorder(&dc, rect, this, m_bOnGlass);
#else
			visualManagerMFC->OnDrawControlBorder(this);
#endif
		}
	}

	if (m_Mode == BrowseMode_None && m_nNCTop <= 0 && m_nNCBottom <= 0)
	{
		return;
	}

	CWindowDC dc(this);
	DoNcPaint(&dc, FALSE);
}
//********************************************************************************
void CBCGPEdit::DoNcPaint(CDC* pDC, BOOL bIsPrint)
{
	CRect rectWindow;
	GetWindowRect (rectWindow);

	CRect rectClient;
	GetClientRect (rectClient);
	ClientToScreen (&rectClient);

	if (m_nNCTop > 0)
	{
		CRect rectTopNC = rectWindow;

		rectTopNC.right -= rectTopNC.left + m_sizeBorder.cx;
		rectTopNC.left = m_sizeBorder.cx;
		rectTopNC.top = m_sizeBorder.cy;
		rectTopNC.bottom = rectTopNC.top + m_nNCTop;

		DoFillBackground(pDC, rectTopNC);
	}

	if (m_nNCBottom > 0)
	{
		CRect rectBottomNC = rectWindow;
		
		rectBottomNC.right -= rectBottomNC.left + m_sizeBorder.cx;
		rectBottomNC.bottom -= rectBottomNC.top + m_sizeBorder.cy;
		rectBottomNC.left = m_sizeBorder.cx;
		rectBottomNC.top = rectBottomNC.bottom - m_nNCBottom;
		
		DoFillBackground(pDC, rectBottomNC);
	}

	if (m_Mode == BrowseMode_None)
	{
		return;
	}

	m_rectBtn = rectWindow;
	m_rectBtn.left = m_rectBtn.right - m_nBrowseButtonWidth;
	
	m_rectBtn.OffsetRect (rectClient.right + m_nBrowseButtonWidth - rectWindow.right, 0);
	
	m_rectBtn.top += rectClient.top - rectWindow.top;
	if (m_nNCTop >= 0)
	{
		m_rectBtn.top -= m_nNCTop;
	}
	
	m_rectBtn.bottom -= rectWindow.bottom - rectClient.bottom;
	if (m_nNCBottom >= 0)
	{
		m_rectBtn.bottom += m_nNCBottom;
	}
	
	if ((GetStyle() & ES_MULTILINE) != 0)
	{
		m_rectBtn.bottom = min(m_rectBtn.bottom, m_rectBtn.top + m_nBrowseButtonWidth);

		if ((GetStyle() & WS_VSCROLL) != 0 && (!IsInternalScrollBarThemed() || !CBCGPVisualManager::GetInstance()->IsOwnerDrawScrollBar()))
		{
			m_rectBtn.OffsetRect(GetSystemMetrics(SM_CXVSCROLL), 0);
		}
	}

	CRect rect = m_rectBtn;
	rect.OffsetRect (-rectWindow.left, -rectWindow.top);

	CRgn rgnClip;

	if (!bIsPrint)
	{
		rgnClip.CreateRectRgnIndirect (&rect);
		pDC->SelectClipRgn (&rgnClip);
	}

	OnDrawBrowseButton(pDC, rect, m_bIsButtonPressed || m_pPopupDlg != NULL, m_bIsButtonHighlighted);

	if (!bIsPrint)
	{
		pDC->SelectClipRgn (NULL);
	}

	ScreenToClient (&m_rectBtn);

	m_rectClick = m_rectBtn;
	
	if (m_bComboBoxMode)
	{
		GetWindowRect(m_rectClick);
		ScreenToClient(m_rectClick);
	}
}
//********************************************************************************
BCGNcHitTestType CBCGPEdit::OnNcHitTest(CPoint point) 
{
	CPoint ptClient = point;
	ScreenToClient (&ptClient);

	if (m_Mode != BrowseMode_None && m_rectBtn.PtInRect (ptClient))
	{
		return HTBORDER;
	}
	
	return CEdit::OnNcHitTest(point);
}
//********************************************************************************
void CBCGPEdit::DoEraseBrowseButton(CDC* pDC, CRect rect)
{
	const BOOL bEnabled = IsBrowseButtonEnabled();

#ifndef _BCGSUITE_
	const BOOL bSimplifiedPressed = (m_bSimplifiedBrowseIcon && m_bIsButtonPressed && bEnabled) && (!m_bSearchMode || !m_bTextIsEmpty) && m_Mode != BrowseMode_PasswordPreview;
#else
	const BOOL bSimplifiedPressed = FALSE;
#endif

	if (m_bOnGlass)
	{
		CBCGPDrawManager dm(*pDC);
		dm.DrawRect(rect, bSimplifiedPressed ? globalData.clrBtnShadow : bEnabled ? globalData.clrWindow : globalData.clrBtnFace, (COLORREF)-1);
	}
	else if (m_brBackground.GetSafeHandle() != NULL)
	{
		pDC->FillRect(rect, bSimplifiedPressed ? &m_brBackgroundPressed : &m_brBackground);
	}
	else if (m_bVisualManagerStyle)
	{
		if (bSimplifiedPressed)
		{
			LOGBRUSH lbr;
			CBCGPVisualManager::GetInstance ()->GetEditCtrlBackgroundBrush(this).GetLogBrush(&lbr);

			COLORREF clrFill = lbr.lbColor; 
			
			if (CBCGPDrawManager::IsDarkColor(clrFill))
			{
				clrFill = CBCGPDrawManager::ColorMakeLighter(clrFill, .5);
			}
			else
			{
				clrFill = CBCGPDrawManager::ColorMakeDarker(clrFill);
			}

			pDC->FillSolidRect(rect, clrFill);
		}
		else
		{
			pDC->FillRect(rect, &CBCGPVisualManager::GetInstance ()->GetEditCtrlBackgroundBrush(this));
		}
	}
	else
	{
#ifndef _BCGSUITE_
		if (bSimplifiedPressed)
		{
			pDC->FillSolidRect(rect, CBCGPVisualManager::GetInstance()->CBCGPVisualManager::GetHighlightedColor(1));
		}
		else
#endif
		{
			pDC->FillRect(rect, bEnabled ? &globalData.brWindow : &globalData.brBtnFace);
		}
	}
}
//********************************************************************************
void CBCGPEdit::OnDrawBrowseButton (CDC* pDC, CRect rect, BOOL bIsButtonPressed, BOOL bHighlight)
{
	ASSERT (m_Mode != BrowseMode_None);
	ASSERT_VALID (pDC);

	BOOL bEnabled = IsBrowseButtonEnabled();

	if (m_bSearchMode && m_ImageSearch.IsValid() && m_bTextIsEmpty)
	{
		DoEraseBrowseButton(pDC, rect);

		if (bEnabled)
		{
			m_ImageSearch.DrawEx(pDC, rect, 0, CBCGPToolBarImages::ImageAlignHorzCenter, CBCGPToolBarImages::ImageAlignVertCenter);
		}
		else
		{
			CBCGPDrawState ds;

			m_ImageSearch.PrepareDrawImage (ds);

			CSize sizeImage = m_ImageSearch.GetImageSize();

			m_ImageSearch.Draw (pDC, 
				rect.left + max(0, (rect.Width() - sizeImage.cx) / 2), 
				rect.top + max(0, (rect.Height() - sizeImage.cy) / 2), 
				0, FALSE, TRUE);
			m_ImageSearch.EndDrawImage (ds);
		}

		return;
	}

	CBCGPVisualManager::BCGBUTTON_STATE state = CBCGPVisualManager::ButtonsIsRegular;

	if (bIsButtonPressed)
	{
		state = CBCGPVisualManager::ButtonsIsPressed;
	}
	else if (bHighlight)
	{
		state = CBCGPVisualManager::ButtonsIsHighlighted;
	}

	COLORREF clrText = m_bVisualManagerStyle ? globalData.clrBarText : globalData.clrBtnText;

	if (!bEnabled && !globalData.IsHighContastMode())
	{
		clrText = globalData.clrGrayedText;
	}
	
#ifndef _BCGSUITE_
	BOOL bSimplifiedBrowseIcon = m_bSimplifiedBrowseIcon;
	if (!m_bSearchMode && !m_ImageBrowseSimplified.IsValid())
	{
		bSimplifiedBrowseIcon = FALSE;
	}
#else
	BOOL bSimplifiedBrowseIcon = FALSE;
#endif
	if ((m_Mode != BrowseMode_PasswordPreview && !bSimplifiedBrowseIcon) || globalData.IsHighContastMode())
	{
		if (!CBCGPVisualManager::GetInstance ()->OnDrawBrowseButton (pDC, rect, this, state, clrText))
		{
			return;
		}
	}
	else
	{
		DoEraseBrowseButton(pDC, rect);
	}

	if (m_bSearchMode && !m_bTextIsEmpty)
	{
		CRect rectImage = rect;

		if (bIsButtonPressed && visualManagerMFC->IsOffsetPressedButton ())
		{
			rectImage.left++;
			rectImage.top++;
		}

		PrepareClearIcon(FALSE);

#ifndef _BCGSUITE_
		if (m_bSimplifiedBrowseIcon)
		{
			if (bHighlight)
			{
				m_ImageClearSimplifiedHot.DrawEx(pDC, rectImage, 0, CBCGPToolBarImages::ImageAlignHorzCenter, CBCGPToolBarImages::ImageAlignVertCenter);
			}
			else
			{
				m_ImageClearSimplified.DrawEx(pDC, rectImage, 0, CBCGPToolBarImages::ImageAlignHorzCenter, CBCGPToolBarImages::ImageAlignVertCenter);
			}
		}
		else
#endif
		{
			m_ImageClear.DrawEx(pDC, rectImage, 0, CBCGPToolBarImages::ImageAlignHorzCenter, CBCGPToolBarImages::ImageAlignVertCenter);
		}

		return;
	}

	if (m_Mode == BrowseMode_PopupDialog)
	{
		CBCGPToolbarComboBoxButton dummy;

#ifndef _BCGSUITE_
		if (m_bVisualManagerStyle)
		{
			visualManagerMFC->OnDrawComboDropButton(pDC, rect, !bEnabled, bIsButtonPressed, bHighlight, &dummy);
		}
		else
		{
			if (!visualManagerMFC->DrawComboDropButtonWinXP(pDC, rect, !bEnabled, bIsButtonPressed, bHighlight))
			{
				visualManagerMFC->CBCGPVisualManager::OnDrawComboDropButton(pDC, rect, !bEnabled, bIsButtonPressed, bHighlight, &dummy);
			}
		}
#else
		visualManagerMFC->OnDrawComboDropButton(pDC, rect, !bEnabled, bIsButtonPressed, bHighlight, &dummy);
#endif
		return;
	}

	int iImage = 0;

	CBCGPToolBarImages* pImageBrowse = &m_ImageBrowse;
	
#ifndef _BCGSUITE_
	if (m_bSimplifiedBrowseIcon)
	{
		PrepareSimplifiedIcons(FALSE);

		pImageBrowse = (m_bIsButtonHighlighted || m_bIsButtonPressed) ? &m_ImageBrowseSimplifiedHot : &m_ImageBrowseSimplified;
	}
#endif

	if (pImageBrowse->IsValid())
	{
		if (m_bDefaultImage)
		{
			switch (m_Mode)
			{
			case BrowseMode_Folder:
				iImage = 0;
				break;

			case BrowseMode_File:
				iImage = 1;
				break;

			case BrowseMode_Calculator:
				iImage = 2;
				break;

			case BrowseMode_PasswordPreview:
				iImage = 3;
				break;
			}
		}

		CRect rectImage = rect;

		if (bIsButtonPressed && visualManagerMFC->IsOffsetPressedButton ())
		{
			rectImage.left++;
			rectImage.top++;
		}

#ifndef _BCGSUITE_
		if (m_bOnGlass)
		{
			pImageBrowse->ConvertTo32Bits();
		}
#endif
		if (bEnabled)
		{
			pImageBrowse->DrawEx (pDC, rectImage, iImage, CBCGPToolBarImages::ImageAlignHorzCenter, CBCGPToolBarImages::ImageAlignVertCenter);
		}
		else
		{
			CBCGPDrawState ds;

			pImageBrowse->PrepareDrawImage (ds);

			CSize sizeImage = pImageBrowse->GetImageSize();

			pImageBrowse->Draw (pDC, 
				rectImage.left + max(0, (rectImage.Width() - sizeImage.cx) / 2), 
				rectImage.top + max(0, (rectImage.Height() - sizeImage.cy) / 2), 
				iImage, FALSE, TRUE);
			pImageBrowse->EndDrawImage (ds);
		}
	}
	else if (!m_strLabel.IsEmpty ())
	{
		CFont* pFont = (CFont*) pDC->SelectStockObject (DEFAULT_GUI_FONT);

		CRect rectText = rect;
		rectText.DeflateRect (1, 2);
		rectText.OffsetRect (0, -2);

		if (bIsButtonPressed)
		{
			rectText.OffsetRect (1, 1);
		}

		DWORD dwDTFlags = DT_SINGLELINE | DT_CENTER | DT_VCENTER;

		if (m_bOnGlass)
		{
			CBCGPVisualManager::GetInstance ()->DrawTextOnGlass(pDC, m_strLabel, rectText, dwDTFlags, 0, clrText);
		}
		else
		{
			COLORREF clrTextOld = pDC->SetTextColor (clrText);
			int nTextMode = pDC->SetBkMode (TRANSPARENT);

			pDC->DrawText (m_strLabel, rectText, dwDTFlags);

			pDC->SetTextColor (clrTextOld);
			pDC->SetBkMode (nTextMode);
		}

		pDC->SelectObject (pFont);
	}
}
//********************************************************************************
void CBCGPEdit::EnableBrowseButton (BOOL bEnable/* = TRUE*/,
									LPCTSTR szLabel/* = _T("...")*/)
{
	ASSERT_VALID (this);
	ASSERT (szLabel != NULL);

	m_Mode = bEnable ? BrowseMode_Default : BrowseMode_None;
	m_strLabel = szLabel;

	m_ImageBrowse.Clear ();

#ifndef _BCGSUITE_
	m_ImageBrowseSimplified.Clear ();
	m_ImageBrowseSimplifiedHot.Clear();
#endif

	m_sizeImage = CSize (0, 0);

	OnChangeLayout ();
}
//********************************************************************************
void CBCGPEdit::OnChangeLayout ()
{
	ASSERT_VALID (this);

	if (GetSafeHwnd () == NULL)
	{
		m_bIsDelayedLayout = TRUE;
		return;
	}

	m_bIsDelayedLayout = FALSE;

	m_nBrowseButtonWidth = globalUtils.ScaleByDPI(20);
	int nImageMargin = globalUtils.ScaleByDPI(8);

	m_nBrowseButtonWidth = max(m_nBrowseButtonWidth, m_sizeImage.cx + nImageMargin);

	SetWindowPos(NULL, -1, -1, -1, -1, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOZORDER | SWP_FRAMECHANGED);

	if (m_Mode != BrowseMode_None)
	{
		GetWindowRect (m_rectBtn);
		m_rectBtn.DeflateRect(m_sizeBorder.cx, m_sizeBorder.cy);
		m_rectBtn.left = m_rectBtn.right - m_nBrowseButtonWidth;

		ScreenToClient (&m_rectBtn);

		if (m_strBrowseButtonToolTip.IsEmpty())
		{
			if (m_pToolTip->GetSafeHwnd() != NULL)
			{
				m_pToolTip->DelTool(this, ID_TOOLTIP_BUTTON);
			}
		}
		else
		{
			if (m_pToolTip->GetSafeHwnd() == NULL)
			{
				CBCGPTooltipManager::CreateToolTip(m_pToolTip, this, BCGP_TOOLTIP_TYPE_DEFAULT);
			}

			if (m_pToolTip->GetSafeHwnd() != NULL)
			{
				m_pToolTip->AddTool(this, LPSTR_TEXTCALLBACK, &m_rectBtn, ID_TOOLTIP_BUTTON);
			}
		}
	}
	else
	{
		m_rectBtn.SetRectEmpty ();
	}
}
//********************************************************************************
void CBCGPEdit::OnNcLButtonDblClk(UINT /*nHitTest*/, CPoint /*point*/)
{
}
//********************************************************************************
void CBCGPEdit::OnBrowse ()
{
	ASSERT_VALID (this);
	ASSERT (GetSafeHwnd () != NULL);

	CRect rectWindow;
	GetWindowRect(rectWindow);

	if (m_Mode == BrowseMode_PopupDialog)
	{
		if (m_pPopupDlg->GetSafeHwnd() != NULL)
		{
			ClosePopupDlg(FALSE);
		}
		else
		{
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
					m_pPopupDlg->m_pParentEdit = new CBCGPParentEditPtr (this);
				}
				
				CBCGPPopupMenu* pMenuActive = CBCGPPopupMenu::GetActiveMenu ();
				if (pMenuActive->GetSafeHwnd() != NULL)
				{
					pMenuActive->SendMessage(WM_CLOSE);
				}
				
				OnBeforeShowPopupDlg(m_pPopupDlg);

				pPopupMenu->Create(this, m_bIsRightAlignedPopup ? rectWindow.right : rectWindow.left - 1, rectWindow.bottom, NULL, FALSE, FALSE);
			}
		}
		return;
	}

	if (m_bSearchMode && !m_bTextIsEmpty)
	{
		SetWindowText(_T(""));
		m_bIsButtonHighlighted = FALSE;
		RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE);
		return;
	}

	switch (m_Mode)
	{
	case BrowseMode_Folder:
#ifndef BCGP_EXCLUDE_SHELL
		{
#ifndef _BCGSUITE_
			CBCGPShellManager* pShellManager = g_pShellManager;
#else
			CShellManager* pShellManager = NULL;
			CWinAppEx* pApp = DYNAMIC_DOWNCAST(CWinAppEx, AfxGetApp ());
			if (pApp != NULL)
			{
				pShellManager = pApp->GetShellManager ();
			}
#endif
			if (pShellManager != NULL)
			{
				CString strFolder;
				GetWindowText (strFolder);

				if (strFolder.IsEmpty() && !m_strInitialFolder.IsEmpty())
                {
                    strFolder = m_strInitialFolder;
                }

				CString strResult;
				if (pShellManager->BrowseForFolder (strResult, this, strFolder, m_strFolderBrowseTitle.IsEmpty() ? NULL : (LPCTSTR)m_strFolderBrowseTitle, m_nFolderBrowseFlags) &&
					(strResult != strFolder))
				{
					SetWindowText (strResult);
					SetModify (TRUE);
					OnAfterUpdate ();
				}
			}
			else
			{
				ASSERT (FALSE);
			}
		}
#endif
		break;

	case BrowseMode_File:
		{
			CString strEditorText;
			GetWindowText (strEditorText);

			if (strEditorText.FindOneOf (_T("*?<>|")) >= 0)
			{
				if (!OnIllegalFileName (strEditorText))
				{
					SetFocus ();
					return;
				}
			}

			CString strFileName;
			CString strInitialDir;

			if (!strEditorText.IsEmpty ())
			{
				DWORD dwAttrs = ::GetFileAttributes (strEditorText); // Check if strEditorText is a directory
				if ((dwAttrs != DWORD(-1)) && (0 != (dwAttrs & FILE_ATTRIBUTE_DIRECTORY)))
				{
					strInitialDir = strEditorText;
				}
				else
				{
					int iBackSlashPos = strEditorText.ReverseFind (_T('\\'));
					if (iBackSlashPos > 0)
					{
						strInitialDir = strEditorText.Left (iBackSlashPos);
						strFileName = strEditorText.Mid (iBackSlashPos + 1);
					}
					else // no backslash found
					{
						// use current directory
						strFileName = strEditorText;
					}
				}
			}

			CFileDialog dlg (TRUE, m_strDefFileExt, strFileName, m_dwFileDialogFlags, m_strFileFilter, GetParent());

			if (strInitialDir.IsEmpty() && !m_strInitialFolder.IsEmpty())
			{
				strInitialDir = m_strInitialFolder;
			}
			
			// Setup initial directory if possible
			if (!strInitialDir.IsEmpty())
			{
				dlg.m_ofn.lpstrInitialDir = strInitialDir;
			}

			if (dlg.DoModal () == IDOK &&
				strEditorText != dlg.GetPathName ())
			{
				SetWindowText (dlg.GetPathName ());
				SetModify (TRUE);
				OnAfterUpdate ();
			}

			if (GetParent () != NULL)
			{
				GetParent ()->RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ALLCHILDREN);
			}
		}
		break;

	case BrowseMode_Calculator:
		{
			if (m_pCalcPopup != NULL)
			{
				m_pCalcPopup->SendMessage (WM_CLOSE);
				m_pCalcPopup = NULL;
				return;
			}

			CString strValue;
			GetWindowText (strValue);

			double dblValue = 0.;
			if (!strValue.IsEmpty ())
			{
				strValue.Replace (_T(','), _T('.'));
#if _MSC_VER < 1400
				_stscanf (strValue, _T("%lf"), &dblValue);
#else
				_stscanf_s (strValue, _T("%lf"), &dblValue);
#endif
			}

			m_pCalcPopup = new CBCGPCalculatorPopup (dblValue, 0, this);

			CBCGPCalculator* pCalc = DYNAMIC_DOWNCAST (CBCGPCalculator, m_pCalcPopup->GetMenuBar());
			if (pCalc != NULL)
			{
				ASSERT_VALID (pCalc);

				if (!m_lstCalcAdditionalCommands.IsEmpty ())
				{
					pCalc->SetAdditionalCommands (m_lstCalcAdditionalCommands);
				}

				if (!m_lstCalcExtCommands.IsEmpty ())
				{
					pCalc->SetExtendedCommands (m_lstCalcExtCommands);
				}

				if (!m_strCalcDisplayFormat.IsEmpty())
				{
					pCalc->SetDisplayFormat(m_strCalcDisplayFormat);
				}
			}

			BOOL bIsRTL = (GetExStyle () & WS_EX_RTLREADING);

			if (!m_pCalcPopup->Create (this, bIsRTL ? rectWindow.right : rectWindow.left, rectWindow.bottom, NULL, TRUE))
			{
				ASSERT (FALSE);
				m_pCalcPopup = NULL;
			}
			else
			{
				m_pCalcPopup->GetMenuBar()->SetFocus ();
				
				CRect rect;
				m_pCalcPopup->GetWindowRect (&rect);
				m_pCalcPopup->UpdateShadow (&rect);

				return;
			}
		}
		break;
	}

	SetFocus ();
}
//********************************************************************************
BOOL CBCGPEdit::OnIllegalFileName (CString& strFileName)
{
	CString strError;
	strError.LoadString (AFX_IDP_INVALID_FILENAME);

	BSTR bsError = strError.AllocSysString (); // Convert to unicode string
	BSTR bsFileName = strFileName.AllocSysString ();

	struct EDITBALLOONTIP_
	{
	    DWORD cbStruct;
		LPCWSTR pszTitle;
	    LPCWSTR pszText;
		INT ttiIcon;
	} tooltip;

	tooltip.cbStruct = sizeof (tooltip);
	tooltip.pszTitle = bsError;
	tooltip.pszText = bsFileName;
	tooltip.ttiIcon = 3; // TTI_ERROR constant

	if (!SendMessage(0x1503, 0, (LPARAM)&tooltip)) // Sending EM_SHOWBALLOONTIP message to edit control
	{
		CString strMessage;
		strMessage.Format (_T("%s\r\n%s"), (LPCTSTR)strFileName, (LPCTSTR)strError);

		BCGPShowMessageBox(m_bVisualManagerStyle, this, strMessage, NULL, MB_OK | MB_ICONEXCLAMATION);
	}

	::SysFreeString (bsFileName);
	::SysFreeString (bsError);
	return FALSE;
}
//********************************************************************************
void CBCGPEdit::SetBrowseButtonImage (HICON hIcon, BOOL bAutoDestroy, BOOL bAlphaBlend)
{
	m_ImageBrowse.Clear ();
#ifndef _BCGSUITE_
	m_ImageBrowseSimplified.Clear ();
	m_ImageBrowseSimplifiedHot.Clear();
#endif

	if (hIcon == NULL)
	{
		m_sizeImage = CSize (0, 0);
		return;
	}

	m_sizeImage = globalUtils.GetIconSize(hIcon);

	m_ImageBrowse.SetImageSize(m_sizeImage);
	m_ImageBrowse.AddIcon(hIcon, bAlphaBlend);

	m_sizeImage = globalUtils.ScaleByDPI(m_ImageBrowse);

	m_bDefaultImage = FALSE;

	if (bAutoDestroy)
	{
		::DestroyIcon (hIcon);
	}
}
//********************************************************************************
void CBCGPEdit::SetBrowseButtonImage (HBITMAP hBitmap, BOOL bAutoDestroy)
{
	m_ImageBrowse.Clear ();

#ifndef _BCGSUITE_
	m_ImageBrowseSimplified.Clear ();
	m_ImageBrowseSimplifiedHot.Clear();
#endif

	if (hBitmap == NULL)
	{
		m_sizeImage = CSize (0, 0);
		return;
	}

	BITMAP bmp;
	::GetObject (hBitmap, sizeof (BITMAP), (LPVOID) &bmp);

	m_sizeImage.cx = bmp.bmWidth;
	m_sizeImage.cy = bmp.bmHeight;

	m_ImageBrowse.SetImageSize(m_sizeImage);
	m_ImageBrowse.AddImage(hBitmap, TRUE);

	m_sizeImage = globalUtils.ScaleByDPI(m_ImageBrowse);

	m_bDefaultImage = FALSE;

	if (bAutoDestroy)
	{
		::DeleteObject (hBitmap);
	}
}
//********************************************************************************
void CBCGPEdit::SetBrowseButtonImage (UINT uiBmpResId)
{
	m_ImageBrowse.Clear ();

#ifndef _BCGSUITE_
	m_ImageBrowseSimplified.Clear ();
	m_ImageBrowseSimplifiedHot.Clear();
#endif

	if (uiBmpResId == 0)
	{
		m_sizeImage = CSize (0, 0);
		return;
	}

	m_ImageBrowse.SetTransparentColor(globalData.clrBtnFace);

	if (!m_ImageBrowse.Load (uiBmpResId))
	{
		ASSERT (FALSE);
		return;
	}

#ifndef _BCGSUITE_
	m_ImageBrowse.SetSingleImage(FALSE);
#else
	m_ImageBrowse.SetSingleImage();
#endif
	m_bDefaultImage = FALSE;
}
//*********************************************************************************
void CBCGPEdit::SetBrowseButtonToolTip(LPCTSTR lpszToolTip, LPCTSTR lpszDescription /*= NULL*/)
{
	ASSERT_VALID(this);

	m_strBrowseButtonToolTip = lpszToolTip == NULL ? _T("") : lpszToolTip;
	m_strBrowseButtonDescr = lpszDescription == NULL ? _T("") : lpszDescription;

	OnChangeLayout();
}
//*********************************************************************************
void CBCGPEdit::EnableFileBrowseButton (LPCTSTR lpszDefExt/* = NULL*/, LPCTSTR lpszFilter/* = NULL*/, LPCTSTR lpszInitialFolder/* = NULL*/, DWORD dwFlags/* = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT*/)
{
	ASSERT_VALID (this);

	m_strDefFileExt = lpszDefExt == NULL ? _T("") : lpszDefExt;
	m_strFileFilter = lpszFilter == NULL ? _T("") : lpszFilter;
	m_strInitialFolder = lpszInitialFolder == NULL ? _T("") : lpszInitialFolder;
	m_dwFileDialogFlags = dwFlags;

	m_Mode = BrowseMode_File;
	SetIntenalImage ();

	globalUtils.EnableEditCtrlAutoComplete (GetSafeHwnd (), FALSE);

	OnChangeLayout ();
}
//*********************************************************************************
void CBCGPEdit::EnableFolderBrowseButton (LPCTSTR lpszTitle, UINT ulFlags, LPCTSTR lpszInitialFolder)
{
#ifdef BCGP_EXCLUDE_SHELL
	ASSERT (FALSE);
#else
	ASSERT_VALID (this);
#ifndef _BCGSUITE_
	ASSERT (g_pShellManager != NULL);	// You need to call CBCGPWorkspace::InitShellManager () first!
#endif

	m_strFolderBrowseTitle = lpszTitle == NULL ? _T("") : lpszTitle;
	m_nFolderBrowseFlags = ulFlags;
	m_strInitialFolder = lpszInitialFolder == NULL ? _T("") : lpszInitialFolder;

	m_Mode = BrowseMode_Folder;
	SetIntenalImage ();

	globalUtils.EnableEditCtrlAutoComplete (GetSafeHwnd (), TRUE);

	OnChangeLayout ();
#endif
}
//*********************************************************************************
void CBCGPEdit::EnableCalculatorButton (const CStringList* plstAdditionalCommands,
										const CList<UINT, UINT>* plstExtCommands,
										LPCTSTR lpszDisplayFormat)
{
	m_Mode = BrowseMode_Calculator;
	m_lstCalcAdditionalCommands.RemoveAll ();
	m_lstCalcExtCommands.RemoveAll();

	if (plstAdditionalCommands != NULL)
	{
		m_lstCalcAdditionalCommands.AddTail ((CStringList*) plstAdditionalCommands);
	}

	if (plstExtCommands != NULL)
	{
		m_lstCalcExtCommands.AddTail ((CList<UINT, UINT>*) plstExtCommands);
	}

	if (lpszDisplayFormat != NULL)
	{
		m_strCalcDisplayFormat = lpszDisplayFormat;
	}

	SetIntenalImage ();
	OnChangeLayout ();
}
//********************************************************************************
void CBCGPEdit::EnablePasswordPreview(BOOL bEnable, BOOL bAllowEditingInPreview)
{
	m_Mode = bEnable ? BrowseMode_PasswordPreview : BrowseMode_None;

	m_bAllowEditingInPasswordPreview = bAllowEditingInPreview;

	SetupPasswordChar();
	SetIntenalImage();
	OnChangeLayout();
}
//********************************************************************************
void CBCGPEdit::EnablePopupDialog(CRuntimeClass* pRTI, UINT nIDTemplate, BOOL bIsResizable, BOOL bComboBoxMode, BOOL bIsRightAligned)
{
	EnablePopupDialog(pRTI, MAKEINTRESOURCE(nIDTemplate), bIsResizable, bComboBoxMode, bIsRightAligned);
}
//********************************************************************************
void CBCGPEdit::EnablePopupDialog(CRuntimeClass* pRTI, LPCTSTR lpszTemplateName, BOOL bIsResizable, BOOL bComboBoxMode, BOOL bIsRightAligned)
{
	if (pRTI != NULL && pRTI->m_pfnCreateObject == NULL)
	{
		TRACE(_T("CBCGPEdit::EnablePopupDialog: you've to add DECLARE_DYNCREATE to your popup dialog class\n"));
		ASSERT(FALSE);
		
		pRTI = NULL;
	}

	m_Mode = BrowseMode_PopupDialog;
	m_pRTIPopupDlg = pRTI;
	m_lpszPopupDlgTemplateName = lpszTemplateName;
	m_bIsResizablePopup = bIsResizable;
	m_bComboBoxMode = bComboBoxMode;
	m_bIsRightAlignedPopup = bIsRightAligned;

	if (GetSafeHwnd() != NULL)
	{
		SendMessage(EM_SETREADONLY, (WPARAM)m_bComboBoxMode);
	}

	OnChangeLayout ();
}
//********************************************************************************
void CBCGPEdit::SetIntenalImage ()
{
	m_ImageBrowse.Clear ();

#ifndef _BCGSUITE_
	m_ImageBrowseSimplified.Clear ();
	m_ImageBrowseSimplifiedHot.Clear();
#endif

	CBCGPLocalResource	lr;

	UINT uiImageListResID = globalData.Is32BitIcons () ? IDB_BCGBARRES_BROWSE32 : IDB_BCGBARRES_BROWSE;

	m_sizeImage = CSize (16, 15);
	m_ImageBrowse.SetImageSize(m_sizeImage);
	m_ImageBrowse.SetTransparentColor(RGB (255, 0, 255));
	m_ImageBrowse.Load(uiImageListResID);

	m_sizeImage = globalUtils.ScaleByDPI(m_ImageBrowse);

#ifndef _BCGSUITE_
	PrepareSimplifiedIcons(TRUE);
#endif

	m_bDefaultImage = TRUE;
}
//********************************************************************************
void CBCGPEdit::SetupPasswordChar()
{
#ifndef _UNICODE
	if (m_Mode == BrowseMode_PasswordPreview && m_bAllowEditingInPasswordPreview && GetSafeHwnd() != NULL)
	{
		SetPasswordChar(_T('*'));
		RedrawWindow();
	}
#endif
}

#ifndef _BCGSUITE_

BOOL CBCGPEdit::GetSimplifiedIconColors(COLORREF& clrSimple, COLORREF& clrHot, COLORREF& clrBase, BOOL bIsCloseButton)
{
	clrSimple = globalData.clrPrompt;
	clrHot = CBCGPDrawManager::ColorMakeDarker(clrSimple, .2);
	clrBase = RGB(0, 0, 0);
	
	if (m_bVisualManagerStyle)
	{
		CBCGPVisualManager::GetInstance()->GetEditSimplifiedIconColors(clrSimple, clrHot, bIsCloseButton);
		
		LOGBRUSH lbr;
		CBCGPVisualManager::GetInstance ()->GetEditCtrlBackgroundBrush(this).GetLogBrush(&lbr);
		
		COLORREF clrFill = lbr.lbColor; 
		
		if (CBCGPDrawManager::IsDarkColor(clrFill))
		{
			clrBase = RGB(255, 255, 255);
			return TRUE;
		}
	}

	return FALSE;
}
//********************************************************************************
void CBCGPEdit::PrepareSimplifiedIcons(BOOL bAlwaysRebuild)
{
	if (m_ImageBrowseSimplified.GetCount() > 0)
	{
		if (!bAlwaysRebuild && m_bSimplifiedBrowseIcon)
		{
			return;
		}

		m_ImageBrowseSimplified.Clear();
		m_ImageBrowseSimplifiedHot.Clear();
	}
	
	if (!m_bSimplifiedBrowseIcon || m_ImageBrowse.GetCount() == 0)
	{
		return;
	}
	
	if (!m_ImageBrowseSimplified.Load(IDB_BCGBARRES_BROWSE_SIMPLIFIED))
	{
		return;
	}

	globalUtils.ScaleByDPI(m_ImageBrowseSimplified);
	m_ImageBrowseSimplified.SetImageSize(m_sizeImage);

	COLORREF clrSimple;
	COLORREF clrHot;
	COLORREF clrBase;

	if (GetSimplifiedIconColors(clrSimple, clrHot, clrBase, FALSE))
	{
		m_ImageBrowseSimplified.InvertColors();
	}

	m_ImageBrowseSimplified.CopyTo(m_ImageBrowseSimplifiedHot);

	m_ImageBrowseSimplified.AdaptColors(clrBase, clrSimple, FALSE);
	m_ImageBrowseSimplifiedHot.AdaptColors(clrBase, clrHot, FALSE);
}

#endif

void CBCGPEdit::PrepareClearIcon(BOOL bAlwaysRebuild)
{
	if (m_ImageClear.GetCount() > 0)
	{
		if (!bAlwaysRebuild)
		{
			return;
		}
		
		m_ImageClear.Clear();
#ifndef _BCGSUITE_
		m_ImageClearSimplified.Clear();
		m_ImageClearSimplifiedHot.Clear();
#endif
	}

	CBCGPLocalResource locaRes;
	m_ImageClear.Load(globalData.Is32BitIcons() ? IDB_BCGBARRES_CLEAR32 : IDB_BCGBARRES_CLEAR);
	
#ifndef _BCGSUITE_
	m_ImageClear.SetSingleImage(FALSE);
#else
	m_ImageClear.SetSingleImage();
#endif
	m_ImageClear.SetTransparentColor(globalData.clrBtnFace);
	
	globalUtils.ScaleByDPI(m_ImageClear);

#ifndef _BCGSUITE_
	m_ImageClear.CopyTo(m_ImageClearSimplified);
	m_ImageClear.CopyTo(m_ImageClearSimplifiedHot);

	const COLORREF clrBaseOrigin = RGB(82, 111, 153);

	if (m_bVisualManagerStyle && CBCGPDrawManager::IsLightColor(globalData.clrBarText))
	{
		m_ImageClear.AdaptColors(clrBaseOrigin, CBCGPVisualManager::GetInstance ()->GetToolbarEditPromptColor(), FALSE);
	}
	
	COLORREF clrSimple;
	COLORREF clrHot;
	COLORREF clrBase;
	
	GetSimplifiedIconColors(clrSimple, clrHot, clrBase, TRUE);

	m_ImageClearSimplified.AdaptColors(clrBaseOrigin, clrSimple, FALSE);
	m_ImageClearSimplifiedHot.AdaptColors(clrBaseOrigin, clrHot, FALSE);
#endif
}
//********************************************************************************
void CBCGPEdit::DoFillBackground(CDC* pDC, const CRect& rect)
{
	ASSERT_VALID(pDC);

	if (rect.IsRectEmpty())
	{
		return;
	}

	const BOOL bIsReadOnly = (GetStyle() & ES_READONLY) == ES_READONLY;
	
	if (m_brBackground.GetSafeHandle() != NULL)
	{
		pDC->FillRect(rect, &m_brBackground);
	}
	else if (m_bVisualManagerStyle && IsWindowEnabled())
	{
		CBrush& br = CBCGPVisualManager::GetInstance ()->GetEditCtrlBackgroundBrush(this);
		pDC->FillRect(rect, &br);
	}
	else
	{
		if (IsWindowEnabled() && !bIsReadOnly)
		{
			pDC->FillRect(rect, &globalData.brWindow);
		}
		else if (m_bVisualManagerStyle)
		{
			CBrush& br = CBCGPVisualManager::GetInstance ()->GetEditCtrlBackgroundBrush(this);
			pDC->FillRect(rect, &br);
		}
		else
		{
			pDC->FillRect(rect, &globalData.brBtnFace);
		}
	}
}
//********************************************************************************
void CBCGPEdit::OnAfterUpdate ()
{
	if (GetOwner()->GetSafeHwnd() == NULL)
	{
		return;
	}

	GetOwner()->PostMessage (EN_CHANGE, GetDlgCtrlID (), (LPARAM) GetSafeHwnd ());
	GetOwner()->PostMessage (EN_UPDATE, GetDlgCtrlID (), (LPARAM) GetSafeHwnd ());
}
//**********************************************************************************
void CBCGPEdit::OnNcMouseMove(UINT nHitTest, CPoint point) 
{
	if (!m_bIsButtonCaptured)
	{
		CPoint ptClient = point;
		ScreenToClient (&ptClient);

		if (m_rectClick.PtInRect (ptClient) && IsBrowseButtonEnabled())
		{
			SetCapture ();

			m_bIsButtonHighlighted = TRUE;

			RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE);
		}
	}
	
	CEdit::OnNcMouseMove(nHitTest, point);
}
//*********************************************************************************
void CBCGPEdit::OnCancelMode() 
{
	CEdit::OnCancelMode();
	
	CloseAutocompleteList();

	if (IsBrowseButtonEnabled() && GetCapture()->GetSafeHwnd() == GetSafeHwnd())
	{
		ReleaseCapture ();
	}

	m_bIsButtonPressed = FALSE;
	m_bIsButtonCaptured = FALSE;
	m_bIsButtonHighlighted = FALSE;

	if (m_Mode == BrowseMode_PasswordPreview)
	{
		if (m_bAllowEditingInPasswordPreview)
		{
			SetPasswordChar(m_cPassword);
		}
		else
		{
			ShowCaret();
		}
	}

	RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE);
}
//********************************************************************************
void CBCGPEdit::OnLButtonDown(UINT nFlags, CPoint point) 
{
	if (m_Mode != BrowseMode_None && m_rectClick.PtInRect(point) && IsBrowseButtonEnabled())
	{
		SetFocus ();

		m_bIsButtonPressed = TRUE;
		m_bIsButtonCaptured = TRUE;

		if (m_Mode == BrowseMode_PasswordPreview)
		{
			if (m_bAllowEditingInPasswordPreview)
			{
				m_cPassword = GetPasswordChar();
				SetPasswordChar(0);
			}
			else
			{
				HideCaret();
			}
		}

		SetCapture ();

		RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE);
		return;
	}

	CEdit::OnLButtonDown(nFlags, point);
}
//********************************************************************************
BOOL CBCGPEdit::PreTranslateMessage(MSG* pMsg) 
{
	switch (pMsg->message)
	{
	case WM_SYSKEYDOWN:
		if (m_Mode != BrowseMode_None && m_Mode != BrowseMode_PasswordPreview &&
			(pMsg->wParam == VK_DOWN || pMsg->wParam == VK_RIGHT))
		{
			OnBrowse ();
			return TRUE;
		}

	case WM_LBUTTONDOWN:
	case WM_RBUTTONDOWN:
	case WM_MBUTTONDOWN:
	case WM_LBUTTONUP:
	case WM_RBUTTONUP:
	case WM_MBUTTONUP:
	case WM_NCLBUTTONDOWN:
	case WM_NCRBUTTONDOWN:
	case WM_NCMBUTTONDOWN:
	case WM_NCLBUTTONUP:
	case WM_NCRBUTTONUP:
	case WM_NCMBUTTONUP:
	case WM_MOUSEMOVE:
		if (m_pToolTip->GetSafeHwnd () != NULL)
		{
			m_pToolTip->RelayEvent(pMsg);
		}
		break;

	case WM_CHAR:
		if (m_Mode == BrowseMode_Calculator && !FilterCalcKey ((int) pMsg->wParam))
		{
			return TRUE;
		}
		break;
	}

	return CEdit::PreTranslateMessage(pMsg);
}
//*********************************************************************************
BOOL CBCGPEdit::FilterCalcKey (int nChar)
{
	CString str;
	GetWindowText (str);

	if (isdigit (nChar))
	{
		return TRUE;
	}

	switch (nChar)
	{
	case _T('-'):
		{
			int nStartPos, nEndPos;
			GetSel (nStartPos, nEndPos);

			if (nStartPos == 0 && nEndPos == str.GetLength())
			{
				return TRUE;
			}

			return str.Find (_T('-')) == -1 && nStartPos == 0;
		}

	case _T('.'):
	case _T(','):
		return str.FindOneOf (_T(".,")) == -1;
	}

	return !isprint (nChar);
}
//***********************************************************************************
void CBCGPEdit::OnCalculatorUserCommand (CBCGPCalculator* /*pCalculator*/, 
										 UINT /*uiCmd*/)
{
	ASSERT (FALSE);	// Must be implemented in derived class
}
//************************************************************************************
LRESULT CBCGPEdit::OnBCGSetControlVMMode (WPARAM wp, LPARAM)
{
	m_bVisualManagerStyle = (BOOL) wp;

#ifndef _BCGSUITE_
	PrepareSimplifiedIcons(TRUE);
#endif

	PrepareClearIcon(TRUE);

	return 0;
}
//*************************************************************************************
LRESULT CBCGPEdit::OnBCGSetControlAero (WPARAM wp, LPARAM)
{
	m_bOnGlass = (BOOL) wp;
	return 0;
}
//*************************************************************************************
BOOL CBCGPEdit::IsDrawPrompt()
{
	ASSERT_VALID(this);
	return (m_bSearchMode || m_bHasPrompt) && (m_bTextIsEmpty && !m_strSearchPrompt.IsEmpty() || !m_strErrorMessage.IsEmpty()) && (GetFocus() != this);
}
//*************************************************************************************
void CBCGPEdit::OnPaint() 
{
	m_bShowToolTip = FALSE;

	BOOL bDrawPrompt = IsDrawPrompt();
	BOOL bDrawPassword = (m_Mode == BrowseMode_PasswordPreview && m_bIsButtonPressed && m_bIsButtonCaptured && !m_bAllowEditingInPasswordPreview);

	if (!m_bOnGlass && !bDrawPrompt && !bDrawPassword)
	{
		Default ();
		return;
	}

	CPaintDC dc(this); // device context for painting

	CBCGPMemDC memDC (dc, this, 255 /* Opaque */);
	CDC* pDC = &memDC.GetDC ();

	DoPaint(pDC, bDrawPrompt, FALSE);
}
//*************************************************************************************
void CBCGPEdit::DoPaint(CDC* pDC, BOOL bDrawPrompt, BOOL /*bIsPrint*/)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);

	BOOL bDrawPassword = (m_Mode == BrowseMode_PasswordPreview && m_bIsButtonPressed && m_bIsButtonCaptured && !m_bAllowEditingInPasswordPreview);

	const DWORD dwStyle = GetStyle();
	const DWORD dwExStyle = GetExStyle();

	if (bDrawPrompt || bDrawPassword)
	{
		BOOL bIsMultiLine = (dwStyle & ES_MULTILINE) != 0;

		// Fill control background:
		if (GetWindowTextLength() > 0)
		{
			CRect rectClient;
			GetClientRect(rectClient);

			DoFillBackground(pDC, rectClient);
		}
		else
		{
			m_bDefaultPrintClient = TRUE;

			if (bIsMultiLine)
			{
				SendMessage (WM_PRINT, (WPARAM) pDC->GetSafeHdc (), (LPARAM)(PRF_CLIENT | PRF_ERASEBKGND));
			}

			SendMessage (WM_PRINTCLIENT, (WPARAM) pDC->GetSafeHdc (), (LPARAM)PRF_CLIENT);
			m_bDefaultPrintClient = FALSE;
		}

		COLORREF clrText = m_strErrorMessage.IsEmpty() ? m_clrPrompt : m_clrErrorText;

		if (clrText == (COLORREF)-1)
		{
			if (m_ColorTheme.m_clrPrompt != (COLORREF)-1)
			{
				clrText = m_ColorTheme.m_clrPrompt;
			}
			else
			{
#ifndef _BCGSUITE_
				clrText = m_bVisualManagerStyle ? CBCGPVisualManager::GetInstance ()->GetToolbarEditPromptColor() : globalData.clrPrompt;
#else
				clrText = globalData.clrGrayedText;
#endif
			}
		}

		pDC->SetTextColor(clrText);
		pDC->SetBkMode(TRANSPARENT);

		CFont* pOldFont = pDC->SelectObject (GetFont());

		CRect rectText;
		GetClientRect(rectText);

		DWORD lRes = GetMargins ();
		rectText.left += LOWORD(lRes);

		if ((dwStyle & WS_BORDER) != 0 || (dwExStyle & WS_EX_CLIENTEDGE) != 0)
		{
			rectText.DeflateRect (1, 1);
		}

		rectText.left += globalUtils.ScaleByDPI(2);

		UINT nFormat = bIsMultiLine ? (DT_WORDBREAK | DT_NOPREFIX) : (DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);
		if (dwExStyle & WS_EX_RIGHT)
		{
			nFormat |= DT_RIGHT;
		}
		else
		{
			if (dwStyle & ES_RIGHT)
			{
				nFormat |= DT_RIGHT;
			}
			else if (dwStyle & ES_CENTER)
			{
				nFormat |= DT_CENTER;
			}
			else
			{
				nFormat |= DT_LEFT;
			}
		}

		CString str = m_strErrorMessage.IsEmpty() ? m_strSearchPrompt : m_strErrorMessage;
		if (bDrawPassword)
		{
			if (m_bVisualManagerStyle || m_ColorTheme.m_clrText != (COLORREF)-1)
			{
				pDC->SetTextColor(
					m_ColorTheme.m_clrText != (COLORREF)-1 ?
					m_ColorTheme.m_clrText : CBCGPVisualManager::GetInstance ()->GetEditCtrlTextColor(this));
			}

			GetWindowText(str);
		}

		if (pDC->GetTextExtent(str).cx > rectText.Width())
		{
			m_bShowToolTip = TRUE;

			if (m_pToolTip->GetSafeHwnd() == NULL)
			{
				CBCGPTooltipManager::CreateToolTip (m_pToolTip, this, BCGP_TOOLTIP_TYPE_DEFAULT);
			}

			if (m_pToolTip->GetSafeHwnd () != NULL)
			{
				m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectText, ID_TOOLTIP_TEXT);
			}
		}

		if (m_bOnGlass)
		{
			CBCGPVisualManager::GetInstance ()->DrawTextOnGlass(pDC, str, rectText, nFormat, 0, clrText);
		}
		else
		{
			pDC->DrawText(str, rectText, nFormat);
		}

		pDC->SelectObject (pOldFont);
	}
	else
	{
		m_bDefaultPrintClient = TRUE;
		SendMessage (WM_PRINTCLIENT, (WPARAM) pDC->GetSafeHdc (), (LPARAM) PRF_CLIENT);
		m_bDefaultPrintClient = FALSE;
	}
}
//**********************************************************************************************************
BOOL CBCGPEdit::OnChange() 
{
	if (m_bOnGlass)
	{
		InvalidateRect (NULL, FALSE);
		UpdateWindow ();
	}

	if (m_pAutoCompleteCombo->GetSafeHwnd() != NULL && m_bIsLastPrintableCharacter)
	{
		CString strText;
		GetWindowText(strText);

		CString strOut;
		if (m_pAutoCompleteCombo->OnAutoComplete(strText, strOut) >= 0)
		{
			SetWindowText(strOut);
			SetSel(strText.GetLength(), -1, TRUE);
		}
	}

	m_bIsLastPrintableCharacter = FALSE;

#ifndef _BCGSUITE_
	BOOL bIsAutocompleteAvailable = (m_Mode == BrowseMode_None || m_Mode == BrowseMode_Default || (m_Mode == BrowseMode_PopupDialog && !m_bComboBoxMode));

	if (bIsAutocompleteAvailable && !m_bInAutoComplete)
	{
		BOOL bDestroyDropDown = TRUE;
		
		CString strText;
		GetWindowText(strText);

		if (!strText.IsEmpty() && (GetStyle() & ES_MULTILINE) == 0)
		{
			CStringList	lstAutocomplete;
			if (OnGetAutoCompleteList(strText, lstAutocomplete) && !lstAutocomplete.IsEmpty())
			{
				bDestroyDropDown = FALSE;

				if (::IsWindow(m_pDropDownPopup->GetSafeHwnd()) && m_pDropDownPopup->Compare(lstAutocomplete))
				{
					// Keep existing list
				}
				else
				{
					CreateAutocompleteList(lstAutocomplete);
				}
			}
		}

		if (bDestroyDropDown)
		{
			CloseAutocompleteList();
		}
	}
#endif

	if (m_bSearchMode || m_bHasPrompt)
	{
		BOOL bTextIsEmpty = m_bTextIsEmpty;

		CString str;
		GetWindowText (str);

		m_bTextIsEmpty = str.IsEmpty();

		if (!m_strErrorMessage.IsEmpty())
		{
			SetErrorMessage(NULL, m_clrErrorText);
		}
		else if (bTextIsEmpty != m_bTextIsEmpty)
		{
			RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
		}
	}

	return FALSE;
}
//**********************************************************************************************************
LRESULT CBCGPEdit::WindowProc(UINT message, WPARAM wParam, LPARAM lParam) 
{
	CPoint ptCaret (-1, -1);

	BOOL bCheckSel = FALSE;

	if (m_bOnGlass)
	{
		if ((message >= WM_MOUSEFIRST && message <= WM_MOUSELAST) ||
			(message >= WM_KEYFIRST && message <= WM_KEYLAST))
		{
			ptCaret = GetCaretPos ();
			bCheckSel = TRUE;
		}
	}

	LRESULT lres = TBCGPInternalScrollBarWrapperWnd<CEdit>::WindowProc(message, wParam, lParam);

	if (bCheckSel)
	{
		if (GetSel () != 0 || ptCaret != GetCaretPos ())
		{
			InvalidateRect (NULL, FALSE);
			UpdateWindow ();
		}
	}

	if (message == EM_SETSEL && m_bOnGlass)
	{
		InvalidateRect (NULL, FALSE);
		UpdateWindow ();
	}

	if (message == WM_SETTEXT)
	{
		if ((GetStyle() & ES_MULTILINE) != 0 && GetOwner()->GetSafeHwnd() != NULL && (m_bSearchMode || m_bHasPrompt))
		{
			// The EN_CHANGE notification code is not sent when the ES_MULTILINE style is used and the text is sent through WM_SETTEXT
			GetOwner ()->PostMessage(WM_COMMAND, MAKEWPARAM(GetDlgCtrlID(), EN_CHANGE), (LPARAM)GetSafeHwnd());
		}
	}

	return lres;
}
//**********************************************************************************************************
HBRUSH CBCGPEdit::CtlColor(CDC* pDC, UINT /*nCtlColor*/) 
{
	if (m_bVisualManagerStyle || m_ColorTheme.m_clrText != (COLORREF)-1 || m_brBackground.GetSafeHandle() != NULL)
	{
		CBrush& br = m_brBackground.GetSafeHandle() != NULL ?
			m_brBackground : CBCGPVisualManager::GetInstance ()->GetEditCtrlBackgroundBrush(this);

		if ((GetStyle() & ES_READONLY) == ES_READONLY || !IsWindowEnabled())
		{
			LOGBRUSH lbr;
			br.GetLogBrush(&lbr);
			
			pDC->SetBkColor(lbr.lbColor);

			if (!IsWindowEnabled())
			{
				pDC->SetTextColor (CBCGPVisualManager::GetInstance ()->GetEditCtrlTextColor(this));
			}
			else
			{
				pDC->SetTextColor (globalData.clrBarText);
			}
			
			return (HBRUSH)br.GetSafeHandle();
		}
		else
		{
			LOGBRUSH lbr;
			br.GetLogBrush(&lbr);
			
			pDC->SetBkColor(lbr.lbColor);

			pDC->SetTextColor(
				m_ColorTheme.m_clrText != (COLORREF)-1 ?
				m_ColorTheme.m_clrText : CBCGPVisualManager::GetInstance ()->GetEditCtrlTextColor(this));

			return (HBRUSH)br.GetSafeHandle();
		}
	}

	return NULL;
}
//**************************************************************************
void CBCGPEdit::SetPrompt(LPCTSTR lpszPrompt, COLORREF clrText, BOOL bRedraw)
{
	ASSERT_VALID (this);

	CString strOldPrompt = m_strSearchPrompt;
	BOOL bColorWasChanged = m_clrPrompt != clrText;

	m_strSearchPrompt = (lpszPrompt == NULL) ? _T("") : lpszPrompt;
	m_clrPrompt = clrText;
	m_bHasPrompt = !m_strSearchPrompt.IsEmpty() || !m_strErrorMessage.IsEmpty();
	
	if (GetSafeHwnd() != NULL)
	{
		BOOL bTextWasEmpty = m_bTextIsEmpty;

		CString str;
		GetWindowText (str);
		
		m_bTextIsEmpty = str.IsEmpty();

		if (bRedraw && (bColorWasChanged || m_strSearchPrompt != strOldPrompt || bTextWasEmpty != m_bTextIsEmpty))
		{
			RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
		}
	}
}
//**************************************************************************
void CBCGPEdit::SetErrorMessage(LPCTSTR lpszMessage, COLORREF clrText, BOOL bRedraw)
{
	ASSERT_VALID (this);
	
	CString strOldMessage = m_strErrorMessage;
	BOOL bColorWasChanged = m_clrErrorText != clrText;
	
	m_strErrorMessage = (lpszMessage == NULL) ? _T("") : lpszMessage;
	m_clrErrorText = clrText;
	m_bHasPrompt = !m_strSearchPrompt.IsEmpty() || !m_strErrorMessage.IsEmpty();
	
	if (GetSafeHwnd() != NULL)
	{
		CString str;
		GetWindowText (str);
		
		m_bTextIsEmpty = str.IsEmpty();
		
		if (bRedraw && (bColorWasChanged || strOldMessage != m_strErrorMessage))
		{
			RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
		}
	}
}
//**************************************************************************
void CBCGPEdit::EnableSearchMode(BOOL bEnable, LPCTSTR lpszPrompt, COLORREF clrText, BOOL bRedraw)
{
	ASSERT_VALID (this);

	m_bSearchMode = bEnable;

	if (m_bSearchMode)
	{
		ASSERT(lpszPrompt != NULL);

		m_strSearchPrompt = lpszPrompt;
		m_clrPrompt = clrText;

		if (!m_ImageSearch.IsValid())
		{
			CBCGPLocalResource locaRes;
			m_ImageSearch.Load(globalData.Is32BitIcons () ?
				IDB_BCGBARRES_SEARCH32 : IDB_BCGBARRES_SEARCH);

#ifndef _BCGSUITE_
			m_ImageSearch.SetSingleImage(FALSE);
#else
			m_ImageSearch.SetSingleImage();
#endif
			m_ImageSearch.SetTransparentColor(globalData.clrBtnFace);

			globalUtils.ScaleByDPI(m_ImageSearch);
		}

		PrepareClearIcon(TRUE);
	}
	else
	{
		if (m_ImageClear.IsValid())
		{
			m_ImageClear.Clear();
		}

#ifndef _BCGSUITE_
		if (m_ImageClearSimplified.IsValid())
		{
			m_ImageClear.Clear();
		}

		if (m_ImageClearSimplifiedHot.IsValid())
		{
			m_ImageClearSimplifiedHot.Clear();
		}
#endif

		if (m_ImageSearch.IsValid())
		{
			m_ImageSearch.Clear();
		}
	}

	EnableBrowseButton(bEnable);

	if (bRedraw && GetSafeHwnd() != NULL)
	{
		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	}
}
//**************************************************************************
void CBCGPEdit::OnDestroy() 
{
	CloseAutocompleteList();
	CBCGPTooltipManager::DeleteToolTip (m_pToolTip);
	CEdit::OnDestroy();
}
//*****************************************************************************************
BOOL CBCGPEdit::OnNeedTipText(UINT /*id*/, NMHDR* pNMH, LRESULT* /*pResult*/)
{
	LPNMTTDISPINFO	pTTDispInfo	= (LPNMTTDISPINFO) pNMH;
	ASSERT((pTTDispInfo->uFlags & TTF_IDISHWND) == 0);

	if (pNMH->hwndFrom == m_pToolTip->GetSafeHwnd ())
	{
		if (m_bShowToolTip && pNMH->idFrom == ID_TOOLTIP_TEXT)
		{
			const CString& str = m_strErrorMessage.IsEmpty() ? m_strSearchPrompt : m_strErrorMessage;
			pTTDispInfo->lpszText = const_cast<LPTSTR> ((LPCTSTR) str);
			return TRUE;
		}
		else if (pNMH->idFrom == ID_TOOLTIP_BUTTON)
		{
			pTTDispInfo->lpszText = const_cast<LPTSTR> ((LPCTSTR) m_strBrowseButtonToolTip);

			CBCGPToolTipCtrl* pToolTip = DYNAMIC_DOWNCAST(CBCGPToolTipCtrl, m_pToolTip);
			if (pToolTip != NULL)
			{
				ASSERT_VALID(pToolTip);
				pToolTip->SetDescription(m_strBrowseButtonDescr);
			}
		}
	}

	return FALSE;
}
//*****************************************************************************
LRESULT CBCGPEdit::OnPrintClient(WPARAM wp, LPARAM lp)
{
	DWORD dwFlags = (DWORD)lp;
	
	BOOL bDefaultPrint = TRUE;
	
	if ((dwFlags & PRF_CLIENT) == PRF_CLIENT)
	{
		BOOL bDrawPrompt = IsDrawPrompt();

		if ((m_bOnGlass || bDrawPrompt) && !m_bDefaultPrintClient)
		{
			CDC* pDC = CDC::FromHandle((HDC) wp);
			ASSERT_VALID(pDC);

			DoPaint(pDC, bDrawPrompt, TRUE);
			bDefaultPrint = FALSE;
		}
	}

	if (bDefaultPrint)
	{
		return Default();
	}

	return 0;
}
//*****************************************************************************
LRESULT CBCGPEdit::OnPrint(WPARAM wp, LPARAM lp)
{
	DWORD dwFlags = (DWORD)lp;
	
	if ((dwFlags & PRF_NONCLIENT) == PRF_NONCLIENT)
	{
		CDC* pDC = CDC::FromHandle((HDC) wp);
		ASSERT_VALID(pDC);

		const COLORREF clrBorderCustom = (GetFocus() == this && m_ColorTheme.m_clrBorderFocused != (COLORREF)-1) ? 
			m_ColorTheme.m_clrBorderFocused : m_ColorTheme.m_clrBorder;

		const BOOL bHasBorder = (GetExStyle () & WS_EX_CLIENTEDGE) || (GetStyle () & WS_BORDER);
		
		if (bHasBorder && (m_bVisualManagerStyle || m_bOnGlass || clrBorderCustom != (COLORREF)-1))
		{
			CRect rect;
			GetWindowRect(rect);

			rect.bottom -= rect.top;
			rect.right -= rect.left;
			rect.left = rect.top = 0;

			if (clrBorderCustom != (COLORREF)-1)
			{
				CBCGPDrawManager dm(*pDC);
				dm.DrawRect(rect, (COLORREF)-1, clrBorderCustom);
				
				rect.DeflateRect (1, 1);
				
				COLORREF clrInner = m_ColorTheme.m_clrBackground == (COLORREF)-1 ? globalData.clrWindow : m_ColorTheme.m_clrBackground;
				dm.DrawRect(rect, (COLORREF)-1, clrInner);
			}
			else
			{
#ifndef _BCGSUITE_
				visualManagerMFC->OnDrawControlBorder(pDC, rect, this, m_bOnGlass);
#else
				visualManagerMFC->OnDrawControlBorder(this);
#endif
			}
		}
		
		if (m_Mode != BrowseMode_None || m_nVerticalAlignment != TA_TOP)
		{
			DoNcPaint(pDC, TRUE);
		}
	}

	return Default();
}
//*****************************************************************************
void CBCGPEdit::OnKillFocus(CWnd* pNewWnd) 
{
	HWND hwndThis = GetSafeHwnd();

#ifndef _BCGSUITE_
	if (::IsWindow(m_pDropDownPopup->GetSafeHwnd()) && DYNAMIC_DOWNCAST(CBCGPEditDropDownList, m_pDropDownPopup) != NULL)
	{
		int nIndex = m_pDropDownPopup->GetCurSel();
		if (nIndex >= 0)
		{
			CString strItemText;
			m_pDropDownPopup->GetText(nIndex, strItemText);
			
			m_bInAutoComplete = TRUE;
			SetWindowText(strItemText);
			m_bInAutoComplete = FALSE;
		}

		m_pDropDownPopup->DestroyWindow();
	}

	m_pDropDownPopup = NULL;
#endif

	CEdit::OnKillFocus(pNewWnd);

	if (m_bVisualManagerStyle && ::IsWindow(hwndThis))
	{
		SetWindowPos(NULL, -1, -1, -1, -1, SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
	}
}

#ifndef _BCGSUITE_
BOOL CBCGPEdit::OnGetAutoCompleteList(const CString& strEditText, CStringList& lstAutocomplete)
{
	UNREFERENCED_PARAMETER(strEditText);

	CWnd* pWndOwner = GetOwner();
	if (pWndOwner->GetSafeHwnd() == NULL)
	{
		return FALSE;
	}

	pWndOwner->SendMessage(BCGM_EDIT_ON_FILL_AUTOCOMPLETE_LIST, (WPARAM)GetDlgCtrlID(), (LPARAM)&lstAutocomplete);

	return !lstAutocomplete.IsEmpty();
}
#endif

void CBCGPEdit::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
#ifndef _BCGSUITE_
	if (::IsWindow(m_pDropDownPopup->GetSafeHwnd()))
	{
		if (nChar == VK_UP || nChar == VK_DOWN)
		{
			return;
		}
	}
	else if (nChar == VK_DOWN && (GetStyle() & ES_MULTILINE) == 0)
	{
		CString strText;
		GetWindowText(strText);

		CStringList	lstAutocomplete;
		if (OnGetAutoCompleteList(strText, lstAutocomplete) && !lstAutocomplete.IsEmpty())
		{
			CreateAutocompleteList(lstAutocomplete);
		}
	}

#endif
	
	CEdit::OnKeyDown(nChar, nRepCnt, nFlags);

#ifdef _UNICODE
    m_bIsLastPrintableCharacter =  (nChar != VK_TAB && iswprint ((wint_t) nChar));
#else
	m_bIsLastPrintableCharacter = (isprint (nChar) || nChar > 127);
#endif
	
}
//*****************************************************************************
void CBCGPEdit::CreateAutocompleteList(const CStringList& lstAutocomplete)
{
#ifndef _BCGSUITE_
	CloseAutocompleteList();

	m_pDropDownPopup = new CBCGPEditDropDownList(this);
	
	for (POSITION pos = lstAutocomplete.GetHeadPosition (); pos != NULL;)
	{
		m_pDropDownPopup->AddString(lstAutocomplete.GetNext(pos));
	}
	
	BOOL bIsRTL = (GetExStyle () & WS_EX_RTLREADING);
	
	CRect rect;
	GetWindowRect(&rect);
	
	m_pDropDownPopup->SetMaxHeight(15 * globalData.GetTextHeight());
	m_pDropDownPopup->Track(CPoint(bIsRTL ? rect.right : rect.left, rect.bottom), GetOwner());

	if (m_pDropDownPopup->m_wndScrollBarVert.GetSafeHwnd() != NULL)
	{
		m_pDropDownPopup->EnableVertResize(2 * globalData.GetTextHeight());
		m_pDropDownPopup->RecalcLayout();
		m_pDropDownPopup->UpdateShadow();
	}
#else
	UNREFERENCED_PARAMETER(lstAutocomplete);
#endif
}
//*****************************************************************************
void CBCGPEdit::CloseAutocompleteList()
{
#ifndef _BCGSUITE_
	if (::IsWindow(m_pDropDownPopup->GetSafeHwnd()))
	{
		m_pDropDownPopup->DestroyWindow();
	}

	m_pDropDownPopup = NULL;
#endif
}
//*****************************************************************************
void CBCGPEdit::SetColorTheme(const CBCGPEditColors& colors, BOOL bRedraw)
{
	m_ColorTheme = colors;

	m_brBackground.DeleteObject();
	m_brBackgroundPressed.DeleteObject();

	if (m_ColorTheme.m_clrBackground != (COLORREF)-1)
	{
		m_brBackground.CreateSolidBrush(m_ColorTheme.m_clrBackground);

		COLORREF clrPressed = m_ColorTheme.m_clrBackground;

		if (CBCGPDrawManager::IsDarkColor(clrPressed))
		{
			clrPressed = CBCGPDrawManager::ColorMakeLighter(clrPressed);
		}
		else
		{
			clrPressed = CBCGPDrawManager::ColorMakeDarker(clrPressed);
		}

		m_brBackgroundPressed.CreateSolidBrush(clrPressed);
	}

	if (bRedraw && GetSafeHwnd() != NULL)
	{
		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	}
}
//*****************************************************************************
void CBCGPEdit::SetDisableBrowseButtonInReadOnlyMode(BOOL bDisable)
{
	m_bDisableBrowseButtonInReadOnlyMode = bDisable;

	if (GetSafeHwnd() != NULL)
	{
		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	}
}
//*****************************************************************************
void CBCGPEdit::SetVerticalAlignment(int nAlignment)
{
	if (m_nVerticalAlignment == nAlignment)
	{
		return;
	}

	m_nVerticalAlignment = nAlignment;
	OnChangeLayout();
}

#ifndef _BCGSUITE_

void CBCGPEdit::SetSimplifiedBrowseIcon(BOOL bSet /*= TRUE*/)
{
	m_bSimplifiedBrowseIcon = bSet;	

	if (m_bSimplifiedBrowseIcon)
	{
		PrepareSimplifiedIcons(FALSE);
	}
	else
	{
		m_ImageBrowseSimplified.Clear();
		m_ImageBrowseSimplifiedHot.Clear();
	}
	
	if (GetSafeHwnd() != NULL)
	{
		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	}
}

#endif

void CBCGPEdit::ClosePopupDlg(BOOL /*bOK*/, DWORD_PTR /*dwUserData*/)
{
	if (m_pPopupDlg->GetSafeHwnd() != NULL)
	{
		m_pPopupDlg->GetParent()->PostMessage(WM_CLOSE);
		m_pPopupDlg = NULL;
	}

	RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
}
//*****************************************************************************
BOOL CBCGPEdit::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	if (m_bComboBoxMode)
	{
		SetCursor(::LoadCursor(NULL, IDC_ARROW));
		return TRUE;
	}
	
	return CEdit::OnSetCursor(pWnd, nHitTest, message);
}
//************************************************************************
BOOL CBCGPEdit::IsInternalScrollBarThemed() const
{
	if (GetSafeHwnd() != NULL && (GetStyle() & ES_MULTILINE) == 0)
	{
		return FALSE;
	}

#ifndef _BCGSUITE_
	return (globalData.m_nThemedScrollBars & BCGP_THEMED_SCROLLBAR_EDITBOX) != 0 && m_bVisualManagerStyle;
#else
	return FALSE;
#endif
}
//************************************************************************
void CBCGPEdit::OnSetFocus(CWnd* pOldWnd) 
{
	CEdit::OnSetFocus(pOldWnd);
	
	if (m_bComboBoxMode)
	{
		::DestroyCaret();
	}

	if (m_bVisualManagerStyle)
	{
		SetWindowPos(NULL, -1, -1, -1, -1, SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
	}
}
//************************************************************************
BOOL CBCGPEdit::IsBrowseButtonEnabled() const
{
	if (GetSafeHwnd() == NULL)
	{
		return TRUE;
	}

	if (!IsWindowEnabled())
	{
		return FALSE;
	}

	if ((GetStyle() & ES_READONLY) == ES_READONLY && m_bDisableBrowseButtonInReadOnlyMode)
	{
		return FALSE;
	}

	return TRUE;
}
//************************************************************************
int CBCGPEdit::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CEdit::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	if (m_bIsDelayedLayout)
	{
		OnChangeLayout();
		PostMessage(UM_REDRAWFRAME);
	}
	
	SetupPasswordChar();
	return 0;
}
//************************************************************************
void CBCGPEdit::PreSubclassWindow() 
{
	CEdit::PreSubclassWindow();

	if (m_bIsDelayedLayout)
	{
		OnChangeLayout();
		PostMessage(UM_REDRAWFRAME);
	}

	SetupPasswordChar();
}
//**************************************************************************
LRESULT CBCGPEdit::OnRedrawFrame(WPARAM, LPARAM)
{
	SetWindowPos(NULL, -1, -1, -1, -1, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOZORDER | SWP_FRAMECHANGED);
	return 0;
}
//**************************************************************************
LRESULT CBCGPEdit::OnChangeVisualManager (WPARAM, LPARAM)
{
#ifndef _BCGSUITE_
	PrepareSimplifiedIcons(TRUE);
	PrepareClearIcon(TRUE);

	if (m_bSimplifiedBrowseIcon || (m_bSearchMode && !m_bTextIsEmpty))
	{
		RedrawWindow (NULL, NULL, RDW_FRAME | RDW_INVALIDATE);
	}
#endif

	return 0;
}
