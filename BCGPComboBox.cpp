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
// BCGPComboBox.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPComboBox.h"
#include "BCGPDlgImpl.h"
#include "BCGPGlobalUtils.h"
#ifndef _BCGSUITE_
#include "BCGPToolBarImages.h"
#include "BCGPToolbarComboBoxButton.h"
#endif
#include "BCGPVisualManager.h"
#include "BCGPDrawManager.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define UM_REDRAWFRAME	(WM_USER + 1001)
#define IMAGE_PADDING	globalUtils.ScaleByDPI(2)

/////////////////////////////////////////////////////////////////////////////
// CBCGPDropDownListBox

IMPLEMENT_DYNAMIC(CBCGPDropDownListBox, CListBox)

BEGIN_MESSAGE_MAP(CBCGPDropDownListBox, CListBox)
	ON_MESSAGE(WM_PRINTCLIENT, OnPrintClient)
	ON_WM_ERASEBKGND()
	ON_WM_NCPAINT()
	ON_WM_WINDOWPOSCHANGING()
END_MESSAGE_MAP()

CBCGPDropDownListBox::CBCGPDropDownListBox()
{
	m_pWndCombo = NULL;
	m_bIsReady = FALSE;
}
//****************************************************************************************
void CBCGPDropDownListBox::DrawItem(LPDRAWITEMSTRUCT lpDIS)
{
	if (m_pWndCombo->GetSafeHwnd() != NULL)
	{
		CBCGPRAII<BOOL> raii(m_pWndCombo->m_bDrawItemOnList, TRUE);

		lpDIS->CtlType = ODT_COMBOBOX;
		m_pWndCombo->DrawItem(lpDIS);
	}
}
//****************************************************************************************
void CBCGPDropDownListBox::MeasureItem(LPMEASUREITEMSTRUCT lpMIS)
{
	if (m_pWndCombo->GetSafeHwnd() != NULL)
	{
		lpMIS->CtlType = ODT_COMBOBOX;
		m_pWndCombo->MeasureItem(lpMIS);
	}
}
//*******************************************************************************
LRESULT CBCGPDropDownListBox::OnPrintClient(WPARAM wp, LPARAM lp)
{
	if ((lp & PRF_ERASEBKGND) == PRF_ERASEBKGND && GetCount() == 0 && m_pWndCombo != NULL && m_pWndCombo->m_bVisualManagerStyle)
	{
		CDC* pDC = CDC::FromHandle((HDC) wp);
		ASSERT_VALID(pDC);

		CRect rect;
		GetClientRect(rect);

		CBCGPVisualManager::GetInstance ()->OnFillComboBoxItem(pDC, m_pWndCombo, -1, rect, FALSE, FALSE);
		return 0;
	}

	return Default();
}
//*******************************************************************************
BOOL CBCGPDropDownListBox::OnEraseBkgnd(CDC* pDC)
{
	if ((GetStyle() & (CBS_OWNERDRAWVARIABLE | CBS_OWNERDRAWFIXED)) == 0)
	{
		return (BOOL)Default();
	}

	CRect rectClient;
	GetClientRect(rectClient);

	if (m_pWndCombo->m_bVisualManagerStyle || m_pWndCombo->m_bOnGlass)
	{
		CBCGPVisualManager::GetInstance ()->OnFillComboBoxItem(pDC, m_pWndCombo, -1, rectClient, FALSE, FALSE);
	}
	else
	{
		pDC->FillRect(rectClient, &globalData.brWindow);
	}

	return TRUE;
}
//*******************************************************************************
void CBCGPDropDownListBox::OnNcPaint()
{
	CListBox::OnNcPaint();

	const BOOL bHasBorder = (GetExStyle() & WS_EX_CLIENTEDGE) || (GetStyle() & WS_BORDER);
	const BOOL bIsSimple = (m_pWndCombo->GetStyle() & 0x0003L) == CBS_SIMPLE;

	if (bHasBorder && (m_pWndCombo->m_bVisualManagerStyle || m_pWndCombo->m_bOnGlass))
	{
		CRect rect;
		GetWindowRect (rect);
		
		rect.bottom -= rect.top;
		rect.right -= rect.left;
		rect.left = rect.top = 0;

		CWindowDC dc(this);
		
		if (!bIsSimple)
		{
			CRect rectInner = rect;
			rectInner.DeflateRect(1, 1);

			dc.ExcludeClipRect(rectInner);
		}

		CBCGPVisualManager::GetInstance()->OnDrawControlBorder(&dc, rect, this, m_pWndCombo->m_bOnGlass);

		dc.SelectClipRgn(NULL);
	}
}
//*******************************************************************************
void CBCGPDropDownListBox::OnWindowPosChanging(WINDOWPOS FAR* lpwndpos) 
{
	CListBox::OnWindowPosChanging(lpwndpos);

	if (m_pWndCombo->GetSafeHwnd() != NULL && m_pWndCombo->GetVerticalPadding() > 0)
	{
		CRect rectScreen;
		
		MONITORINFO mi;
		mi.cbSize = sizeof (MONITORINFO);
		if (GetMonitorInfo (MonitorFromWindow(GetSafeHwnd (), MONITOR_DEFAULTTONEAREST), &mi))
		{
			rectScreen = mi.rcWork;
		}
		else
		{
			::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectScreen, 0);
		}

		lpwndpos->cy = min(lpwndpos->cy, rectScreen.Height());
	}
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPComboBoxItemData

struct CBCGPComboBoxItemData
{
public:
	int		m_nImageIndex;
	int		m_iIndent;
	BOOL	m_bGroupStart;
	
	CBCGPComboBoxItemData()
		: m_iIndent(0)
		, m_nImageIndex(-1)
		, m_bGroupStart(FALSE)
	{
	}
};

/////////////////////////////////////////////////////////////////////////////
// CBCGPComboBox

IMPLEMENT_DYNAMIC(CBCGPComboBox, CComboBox)

CBCGPComboBox::CBCGPComboBox()
{
	m_bVisualManagerStyle = FALSE;
	m_bOnGlass = FALSE;
	m_bIsDroppedDown = FALSE;
	m_bIsButtonHighlighted = FALSE;
	m_rectBtn.SetRectEmpty ();
	m_bTracked = FALSE;
	m_clrPrompt = (COLORREF)-1;
	m_clrErrorText = (COLORREF)-1;
	m_bDefaultPrintClient = FALSE;
	m_bAutoComplete = FALSE;
	m_nRowHeight = -1;
	m_nVertPadding = 0;
	m_nHorzPadding = globalUtils.ScaleByDPI(2);
	m_hImageList = NULL;
	m_sizeImage = CSize(0, 0);
	m_nImageVertMargin = 0;
	m_bDrawItemOnList = FALSE;
}

CBCGPComboBox::~CBCGPComboBox()
{
	CleanUpItemsData();
}

void CBCGPComboBox::CleanUpItemsData()
{
	for (int i = 0; i < (int)m_arItemData.GetSize(); i++)
	{
		CBCGPComboBoxItemData* pItemData = m_arItemData[i];
		if (pItemData != NULL)
		{
			delete pItemData;
		}
	}

	m_arItemData.RemoveAll();
}

BEGIN_MESSAGE_MAP(CBCGPComboBox, CComboBox)
	//{{AFX_MSG_MAP(CBCGPComboBox)
	ON_WM_NCPAINT()
	ON_WM_PAINT()
	ON_WM_MOUSEMOVE()
	ON_WM_CANCELMODE()
	ON_WM_LBUTTONDOWN()
	ON_WM_KILLFOCUS()
	ON_WM_CREATE()
	ON_WM_SETFOCUS()
	ON_WM_ERASEBKGND()
	//}}AFX_MSG_MAP
	ON_CONTROL_REFLECT_EX(CBN_EDITUPDATE, OnEditupdate)
	ON_CONTROL_REFLECT_EX(CBN_SELCHANGE, OnSelchange)
	ON_CONTROL_REFLECT_EX(CBN_CLOSEUP, OnCloseup)
	ON_CONTROL_REFLECT_EX(CBN_DROPDOWN, OnDropdown)
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLVMMODE, OnBCGSetControlVMMode)
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLAERO, OnBCGSetControlAero)
	ON_MESSAGE(WM_SETTEXT, OnSetText)
	ON_MESSAGE(WM_PRINTCLIENT, OnPrintClient)
	ON_MESSAGE(UM_REDRAWFRAME, OnRedrawFrame)
	ON_REGISTERED_MESSAGE(BCGM_CHANGEVISUALMANAGER, OnChangeVisualManager)
	ON_MESSAGE(WM_SETFONT, OnSetFont)
	ON_MESSAGE(CB_INSERTSTRING, OnCBInsertString)
	ON_MESSAGE(CB_DELETESTRING, OnCBDeleteString)
	ON_MESSAGE(CB_RESETCONTENT , OnCBResetContent)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPComboBox message handlers

LRESULT CBCGPComboBox::OnBCGSetControlVMMode (WPARAM wp, LPARAM)
{
	m_bVisualManagerStyle = (BOOL) wp;

	SubclassEditBox();
	SubclassListBox();

	return 0;
}
//**************************************************************************
void CBCGPComboBox::SubclassEditBox()
{
	if (GetSafeHwnd() == NULL)
	{
		return;
	}

	if (!m_bVisualManagerStyle && !m_bOnGlass && m_strPrompt.IsEmpty() && m_strErrorMessage.IsEmpty())
	{
		return;
	}

	if (GetEditCtrl().GetSafeHwnd () == NULL && ((GetStyle () & CBS_DROPDOWN) || (GetStyle () & CBS_SIMPLE)))
	{
		CWnd* pWndChild = GetWindow (GW_CHILD);
		
		while (pWndChild != NULL)
		{
			ASSERT_VALID (pWndChild);
			
			if (CWnd::FromHandlePermanent (pWndChild->GetSafeHwnd ()) == NULL)
			{
				#define MAX_CLASS_NAME		255
				#define EDIT_CLASS			_T("Edit")
				
				TCHAR lpszClassName [MAX_CLASS_NAME + 1];
				
				::GetClassName (pWndChild->GetSafeHwnd (), lpszClassName, MAX_CLASS_NAME);
				CString strClass = lpszClassName;
				
				if (strClass == EDIT_CLASS)
				{
					GetEditCtrl().SubclassWindow (pWndChild->GetSafeHwnd ());

					GetEditCtrl().m_bOnGlass = m_bOnGlass;
					GetEditCtrl().m_bVisualManagerStyle = m_bVisualManagerStyle;
					
					if (m_bAutoComplete && (GetStyle () & 0x0003L) == CBS_DROPDOWN)
					{
						GetEditCtrl().m_pAutoCompleteCombo = this;
					}
					break;
				}
			}
			
			pWndChild = pWndChild->GetNextWindow ();
		}
	}
}
//**************************************************************************
void CBCGPComboBox::AdjustItemHeight()
{
	if ((GetStyle() & CBS_OWNERDRAWFIXED) == CBS_OWNERDRAWFIXED)
	{
		MEASUREITEMSTRUCT mis;
		memset(&mis, 0, sizeof(MEASUREITEMSTRUCT));
		
		mis.CtlType = ODT_COMBOBOX;
		mis.CtlID = GetDlgCtrlID();
		
		SendMessage(WM_MEASUREITEM, (WPARAM)mis.CtlID, (LPARAM)&mis);

		const int nHeight = max((int)mis.itemHeight, GetDropDownItemMinHeight());

		m_wndList.SetItemHeight(-1, nHeight);

		if (DYNAMIC_DOWNCAST(CBCGPToolBar, GetParent()) == NULL)
		{
			SetItemHeight(-1, nHeight);
		}
	}
}
//**************************************************************************
void CBCGPComboBox::SubclassListBox()
{
	CalcRowHeight();

	if (GetSafeHwnd() == NULL || !IsThemedDropDownList() || m_wndList.GetSafeHwnd() != NULL)
	{
		return;
	}

#ifndef _BCGSUITE_
	BCGP_COMBOBOXINFO info;
	memset(&info, 0, sizeof(info));
	info.cbSize = sizeof(info);
	
	if (!globalData.GetComboBoxInfo(GetSafeHwnd(), &info))
	{
		return;
	}
#else
	COMBOBOXINFO info;
	memset(&info, 0, sizeof(info));
	info.cbSize = sizeof(info);

	if (!GetComboBoxInfo(&info))
	{
		return;
	}
#endif

	m_wndList.m_pWndCombo = this;
	m_wndList.SubclassWindow(info.hwndList);

	AdjustItemHeight();

	const BOOL bIsSimple = (GetStyle() & 0x0003L) == CBS_SIMPLE;
	if (bIsSimple)
	{
		PostMessage(UM_REDRAWFRAME);
	}
}
//**************************************************************************
LRESULT CBCGPComboBox::OnBCGSetControlAero (WPARAM wp, LPARAM)
{
	m_bOnGlass = (BOOL) wp;
	GetEditCtrl().m_bOnGlass = m_bOnGlass;

	SubclassEditBox();
	return 0;
}
//**************************************************************************
void CBCGPComboBox::OnNcPaint() 
{
#ifndef _BCGSUITE_
	if (globalData.bIsWindows9x)
	{
		Default();
	}
#endif
}
//**************************************************************************
void CBCGPComboBox::OnPaint() 
{
#ifndef _BCGSUITE_
	if (globalData.bIsWindows9x)
	{
		Default();
		return;
	}
#endif

	BOOL bDrawPrompt = FALSE;

	if (!m_strPrompt.IsEmpty() || !m_strErrorMessage.IsEmpty())
	{
		BOOL bTextIsEmpty = GetWindowTextLength() == 0;

		if (GetEditCtrl().GetSafeHwnd () != NULL)
		{
			if (!m_strErrorMessage.IsEmpty())
			{
				GetEditCtrl().SetErrorMessage(m_strErrorMessage, m_clrErrorText);
			}
			else
			{
				GetEditCtrl().SetPrompt(bTextIsEmpty ? m_strPrompt : _T(""), m_clrPrompt);
			}
		}
		else
		{
			bDrawPrompt = bTextIsEmpty || !m_strErrorMessage.IsEmpty();
		}
	}

	if (!m_bVisualManagerStyle && !m_bOnGlass && !bDrawPrompt)
	{
		Default ();
		return;
	}

	CPaintDC dc(this); // device context for painting
	OnDraw(&dc, bDrawPrompt);
}
//**************************************************************************
void CBCGPComboBox::OnDraw(CDC* pDCIn, BOOL bDrawPrompt)
{
	ASSERT_VALID(pDCIn);

	const int nPadding = globalUtils.ScaleByDPI(4);
	const BOOL bIsSimple = (GetStyle() & 0x0003L) == CBS_SIMPLE;
	const BOOL bRight = GetExStyle() & WS_EX_RIGHT;
	const BOOL bLeftSB = GetExStyle() & WS_EX_LEFTSCROLLBAR;

	BYTE alpha = 0;
	if (m_bOnGlass)
	{
		alpha = 255;
	}

	CBCGPMemDC memDC(*pDCIn, this, alpha);
	CDC* pDC = &memDC.GetDC ();
	
	CRect rectClient;
	GetClientRect (rectClient);

	CBCGPDrawManager dm (*pDC);

	if (bIsSimple)
	{
		CBrush& br2 = CBCGPVisualManager::GetInstance()->GetEditCtrlBackgroundBrush(&GetEditCtrl());
		CRect rectFill = rectClient;

		if (m_wndList.GetSafeHwnd() != NULL) 
		{
			CRect rectLB;
			m_wndList.GetWindowRect(&rectLB);
			ScreenToClient(&rectLB);
			
			rectFill.bottom = rectLB.bottom;

			globalData.DrawParentBackground(this, pDC);
		}

		pDC->FillRect(rectFill, &br2);
	}
	else 
	{
		dm.DrawRect(rectClient, m_bVisualManagerStyle ? globalData.clrBarHilite : globalData.clrWindow, (COLORREF)-1);
	}

	BOOL bDefaultDraw = TRUE;

	if (bDrawPrompt)
	{
		COLORREF clrText = !m_strErrorMessage.IsEmpty() ? m_clrErrorText : m_clrPrompt;
		
		if (clrText == (COLORREF)-1)
		{
#ifndef _BCGSUITE_
			clrText = m_bVisualManagerStyle ? CBCGPVisualManager::GetInstance ()->GetToolbarEditPromptColor() : globalData.clrPrompt;
#else
			clrText = globalData.clrGrayedText;
#endif
		}
		
		pDC->SetTextColor(clrText);
		pDC->SetBkMode(TRANSPARENT);
		
		CFont* pOldFont = pDC->SelectObject (GetFont());
		
		CRect rectText;
		GetClientRect(rectText);

		UINT nFormat = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX;

		if (bRight || bLeftSB)
		{
			rectText.right -= nPadding;
		}
		else
		{
			rectText.left += nPadding;
		}

		nFormat |= bRight ? DT_RIGHT : DT_LEFT;

		if ((GetStyle () & WS_BORDER) != 0 || (GetExStyle () & WS_EX_CLIENTEDGE) != 0)
		{
			rectText.DeflateRect (1, 1);
		}
		
		const CString& str = !m_strErrorMessage.IsEmpty() ? m_strErrorMessage : m_strPrompt;
		
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
	#ifndef _BCGSUITE_
		if ((GetStyle() & CBS_OWNERDRAWFIXED) == 0 && (GetStyle() & CBS_OWNERDRAWVARIABLE) == 0)
		{
			bDefaultDraw = !CBCGPVisualManager::GetInstance ()->OnDrawComboBoxText(pDC, this);

			if (m_bVisualManagerStyle && bDefaultDraw)
			{
				CString strText;
				GetWindowText(strText);
				
				CRect rect;
				GetClientRect(rect);
				
				BOOL bIsFocused = GetSafeHwnd() == CWnd::GetFocus()->GetSafeHwnd();
				COLORREF clrText = CBCGPVisualManager::GetInstance ()->OnFillComboBoxItem(pDC, this, GetCurSel(), rect, bIsFocused, bIsFocused);
				
				if (!IsWindowEnabled())
				{
					clrText = CBCGPVisualManager::GetInstance ()->GetToolbarDisabledTextColor();
				}
				
				CFont* pOldFont = pDC->SelectObject(GetFont());
				ASSERT_VALID(pOldFont);
				
				int nOldBkMode = pDC->SetBkMode(TRANSPARENT);
				COLORREF clrTextOld = pDC->SetTextColor(clrText);
				
				const int cxDropDown = ::GetSystemMetrics (SM_CXVSCROLL);
				UINT nFormat = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX;

				if (bRight || bLeftSB)
				{
					rect.left += cxDropDown;
					rect.right -= nPadding;
				}
				else
				{
					rect.left += nPadding;
					rect.right -= cxDropDown;
				}

				nFormat |= bRight ? DT_RIGHT : DT_LEFT;

				if (m_bOnGlass)
				{
					CBCGPVisualManager::GetInstance ()->DrawTextOnGlass(pDC, strText, rect, nFormat, 0, clrText);
				}
				else
				{
					pDC->DrawText(strText, rect, nFormat);
				}

				pDC->SelectObject(pOldFont);
				pDC->SetBkMode(nOldBkMode);
				pDC->SetTextColor(clrTextOld);

				bDefaultDraw = FALSE;
			}
		}

	#endif
		if (!bIsSimple)
		{
			if (bDefaultDraw)
			{
				m_bDefaultPrintClient = TRUE;
				SendMessage (WM_PRINTCLIENT, (WPARAM) pDC->GetSafeHdc (), (LPARAM) PRF_CLIENT);
				m_bDefaultPrintClient = FALSE;
			}
		}
	}

	if ((GetStyle() & CBS_OWNERDRAWFIXED) != 0 || (GetStyle() & CBS_OWNERDRAWVARIABLE) != 0)
	{
		pDC->SelectClipRgn (NULL);
	}

	if (!bIsSimple) 
	{
		const int cxDropDown = ::GetSystemMetrics (SM_CXVSCROLL) + nPadding;

		m_rectBtn = rectClient;

		if (bRight || bLeftSB)
		{
			m_rectBtn.right = m_rectBtn.left + cxDropDown;
		}
		else
		{
			m_rectBtn.left = m_rectBtn.right - cxDropDown;
		}

		m_rectBtn.DeflateRect (2, 2);

		CBCGPDrawOnGlass dog (m_bOnGlass);

		CBCGPToolbarComboBoxButton buttonDummy;
#ifndef _BCGSUITE_
		buttonDummy.m_bIsCtrl = TRUE;

		CBCGPVisualManager::GetInstance ()->OnDrawComboDropButton (
			pDC, m_rectBtn, !IsWindowEnabled (), m_bIsDroppedDown,
			m_bIsButtonHighlighted,
			&buttonDummy);

		if (bDefaultDraw)
		{
			dm.DrawRect (rectClient, (COLORREF)-1, globalData.clrBarShadow);
		}
		else
		{
			CBCGPVisualManager::GetInstance ()->OnDrawControlBorder (pDC, rectClient, this, m_bOnGlass);
		}

#else
		CMFCVisualManager::GetInstance ()->OnDrawComboDropButton (
			pDC, m_rectBtn, !IsWindowEnabled (), m_bIsDroppedDown,
			m_bIsButtonHighlighted,
			&buttonDummy);

		dm.DrawRect (rectClient, (COLORREF)-1, globalData.clrBarShadow);

#endif
	} 
	else
	{
		if (m_wndList.GetSafeHwnd() != NULL) 
		{
			CRect rectLB;
			m_wndList.GetWindowRect(&rectLB);
			ScreenToClient(&rectLB);
			
			rectClient.bottom = rectLB.bottom;
		}

#ifndef _BCGSUITE_
		if (bDefaultDraw)
		{
			dm.DrawRect(rectClient, (COLORREF)-1, globalData.clrBarShadow);
		}
		else
		{
			CBCGPVisualManager::GetInstance()->OnDrawControlBorder(pDC, rectClient, this, m_bOnGlass);
		}
#else
		dm.DrawRect(rectClient, (COLORREF)-1, globalData.clrBarShadow);
#endif
	}

	rectClient.DeflateRect (1, 1);
	dm.DrawRect (rectClient, (COLORREF)-1, m_bVisualManagerStyle ? globalData.clrBarHilite : globalData.clrWindow);
}
//**************************************************************************
BOOL CBCGPComboBox::OnCloseup() 
{
	m_bIsDroppedDown = FALSE;
	m_bIsButtonHighlighted = FALSE;

	RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);

	if (m_wndList.GetSafeHwnd() != NULL)
	{
		m_wndList.m_bIsReady = FALSE;
	}

	return FALSE;
}
//**************************************************************************
BOOL CBCGPComboBox::OnDropdown() 
{
	if (m_bTracked)
	{
		ReleaseCapture ();
		m_bTracked = FALSE;
	}

	m_bIsDroppedDown = TRUE;
	m_bIsButtonHighlighted = FALSE;
	RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
	
	return FALSE;
}
//**************************************************************************
void CBCGPComboBox::OnMouseMove(UINT nFlags, CPoint point) 
{
	if ((nFlags & MK_LBUTTON) == 0 && !m_bIsDroppedDown)
	{
		BOOL bIsButtonHighlighted = m_bIsButtonHighlighted;
		m_bIsButtonHighlighted = m_rectBtn.PtInRect (point);

		if (bIsButtonHighlighted != m_bIsButtonHighlighted)
		{
			if (!m_bTracked)
			{
				if (m_bIsButtonHighlighted)
				{
					SetCapture ();
					m_bTracked = TRUE;
				}
			}
			else
			{
				if (!m_bIsButtonHighlighted)
				{
					ReleaseCapture ();
					m_bTracked = FALSE;
				}
			}
		
			RedrawWindow(m_rectBtn, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
		}
	}
	
	CComboBox::OnMouseMove(nFlags, point);
}
//*****************************************************************************************
void CBCGPComboBox::OnCancelMode() 
{
	CComboBox::OnCancelMode();
	
	if (m_bTracked)
	{
		ReleaseCapture ();
		m_bIsButtonHighlighted = FALSE;
		m_bTracked = FALSE;

		RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
	}
}
//**************************************************************************
void CBCGPComboBox::OnLButtonDown(UINT nFlags, CPoint point) 
{
	if (m_bTracked)
	{
		ReleaseCapture ();
		m_bIsButtonHighlighted = FALSE;
		m_bTracked = FALSE;

		RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
	}

	CComboBox::OnLButtonDown(nFlags, point);
}
//**************************************************************************
void CBCGPComboBox::OnKillFocus(CWnd* pNewWnd) 
{
	HWND hwndThis = GetSafeHwnd();

	CComboBox::OnKillFocus(pNewWnd);

	if (!::IsWindow(hwndThis))
	{
		return;
	}

	BOOL bDrawPrompt = (!m_strPrompt.IsEmpty() && GetWindowTextLength() == 0) || !m_strErrorMessage.IsEmpty();
	if (bDrawPrompt)
	{
		m_bIsButtonHighlighted = FALSE;
	}

	if (m_bVisualManagerStyle || bDrawPrompt)
	{
		RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
	}
}
//**************************************************************************
void CBCGPComboBox::SetPrompt(LPCTSTR lpszPrompt, COLORREF clrText, BOOL bRedraw)
{
	ASSERT_VALID (this);

	CString strOldPrompt = m_strPrompt;
	BOOL bColorWasChanged = m_clrPrompt != clrText;

	m_strPrompt = (lpszPrompt == NULL) ? _T("") : lpszPrompt;
	m_clrPrompt = clrText;
	
	SubclassEditBox();

	if (GetEditCtrl().GetSafeHwnd () != NULL)
	{
		GetEditCtrl().SetPrompt(lpszPrompt, clrText, bRedraw);
	}

	if (bRedraw && GetSafeHwnd() != NULL && (bColorWasChanged || m_strPrompt != strOldPrompt))
	{
		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	}
}
//**************************************************************************
void CBCGPComboBox::SetErrorMessage(LPCTSTR lpszPrompt, COLORREF clrText, BOOL bRedraw)
{
	ASSERT_VALID (this);

	CString strOldPrompt = m_strErrorMessage;
	BOOL bColorWasChanged = m_clrErrorText != clrText;

	m_strErrorMessage = (lpszPrompt == NULL) ? _T("") : lpszPrompt;
	m_clrErrorText = clrText;
	
	SubclassEditBox();

	if (GetEditCtrl().GetSafeHwnd () != NULL)
	{
		GetEditCtrl().SetErrorMessage(lpszPrompt, clrText, bRedraw);
	}

	if (bRedraw && GetSafeHwnd() != NULL && (bColorWasChanged || m_strErrorMessage != strOldPrompt))
	{
		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
	}
}
//**************************************************************************
BOOL CBCGPComboBox::IsThemedDropDownList() const
{
#ifndef _BCGSUITE_
	return (globalData.m_nThemedScrollBars & BCGP_THEMED_SCROLLBAR_COMBOBOX) != 0 && m_bVisualManagerStyle;
#else
	return FALSE;
#endif
}
//**************************************************************************
void CBCGPComboBox::CalcRowHeight()
{
	if (GetFont()->GetSafeHandle() == globalData.fontRegular.GetSafeHandle() || GetFont()->GetSafeHandle() == NULL)
	{
		m_nRowHeight = globalData.GetTextHeight();
	}
	else
	{
		CClientDC dc((CWnd*)this);
		
		CFont* pOldFont = dc.SelectObject(GetFont());
		ASSERT(pOldFont != NULL);
		
		TEXTMETRIC tm;
		dc.GetTextMetrics (&tm);
		
		int nTextMargin = tm.tmHeight < 15 ? 2 : 5;
		
		m_nRowHeight = tm.tmHeight + nTextMargin;
		
		dc.SelectObject(pOldFont);
	}

	if (m_sizeImage.cy > 0)
	{
		m_nRowHeight = max(m_nRowHeight, m_sizeImage.cy + 2 * m_nImageVertMargin);
	}

	m_nRowHeight += 2 * m_nVertPadding;
}
//**************************************************************************
int CBCGPComboBox::GetDropDownItemMinHeight() const
{
	int nHeight = m_nRowHeight;

#ifndef _BCGSUITE_
	if (DYNAMIC_DOWNCAST(CBCGPToolBar, GetParent()) != NULL)
#endif
	{
		nHeight -= globalUtils.ScaleByDPI(1);
	}

	return nHeight;
}
//**************************************************************************
void CBCGPComboBox::PreSubclassWindow() 
{
	CComboBox::PreSubclassWindow();

	SubclassEditBox();

	if (!m_strPrompt.IsEmpty() && GetEditCtrl().GetSafeHwnd () != NULL)
	{
		GetEditCtrl().SetPrompt(m_strPrompt, m_clrPrompt, FALSE);
	}

	SubclassListBox();

	if ((GetExStyle() & WS_EX_RIGHT) || (GetExStyle() & WS_EX_LEFTSCROLLBAR))
	{
		PostMessage(UM_REDRAWFRAME, 1);
	}
}
//**************************************************************************
int CBCGPComboBox::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CComboBox::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	SubclassEditBox();

	if (!m_strPrompt.IsEmpty() && GetEditCtrl().GetSafeHwnd () != NULL)
	{
		GetEditCtrl().SetPrompt(m_strPrompt, m_clrPrompt, FALSE);
	}
	
	SubclassListBox();

	if ((GetExStyle() & WS_EX_RIGHT) || (GetExStyle() & WS_EX_LEFTSCROLLBAR))
	{
		PostMessage(UM_REDRAWFRAME, 1);
	}

	return 0;
}
//**************************************************************************
void CBCGPComboBox::OnSetFocus(CWnd* pOldWnd) 
{
	CComboBox::OnSetFocus(pOldWnd);
	
	BOOL bDrawPrompt = (!m_strPrompt.IsEmpty() && GetWindowTextLength() == 0) || !m_strErrorMessage.IsEmpty();
	if (bDrawPrompt)
	{
		m_bIsButtonHighlighted = TRUE;
	}
	
	RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
}
//**************************************************************************
BOOL CBCGPComboBox::OnEraseBkgnd(CDC* pDC)
{
	if (!m_bVisualManagerStyle)
	{
		return CComboBox::OnEraseBkgnd(pDC);
	}

	return TRUE;
}
//**************************************************************************
BOOL CBCGPComboBox::OnEditupdate() 
{
	if (!m_strErrorMessage.IsEmpty())
	{
		SetErrorMessage(NULL, m_clrErrorText);
	}

	return FALSE;
}
//**************************************************************************
BOOL CBCGPComboBox::OnSelchange() 
{
	if (!m_strErrorMessage.IsEmpty())
	{
		SetErrorMessage(NULL, m_clrErrorText);
	}

	if (m_bVisualManagerStyle)
	{
		RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
	}

	return FALSE;
}
//**************************************************************************
LRESULT CBCGPComboBox::OnSetText (WPARAM, LPARAM)
{
	LRESULT lRes = Default();

	if (m_bVisualManagerStyle)
	{
		RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
	}

	return lRes;
}
//**************************************************************************
LRESULT CBCGPComboBox::WindowProc(UINT message, WPARAM wParam, LPARAM lParam) 
{
	LRESULT lRes = CComboBox::WindowProc(message, wParam, lParam);

	if (message == CB_SETCURSEL && m_bVisualManagerStyle)
	{
		RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
	}

	return lRes;
}
//*******************************************************************************
LRESULT CBCGPComboBox::OnPrintClient(WPARAM wp, LPARAM lp)
{
	if ((lp & PRF_CLIENT) == PRF_CLIENT)
	{
		CDC* pDC = CDC::FromHandle((HDC) wp);
		ASSERT_VALID(pDC);

		const BOOL bIsSimple = (GetStyle() & 0x0003L) == CBS_SIMPLE;

		if (m_bDefaultPrintClient)
		{
			if ((GetStyle() & CBS_OWNERDRAWFIXED) || (GetStyle() & CBS_OWNERDRAWVARIABLE))
			{
				CRect rect;
				GetClientRect(rect);

				rect.DeflateRect(2, 2);

				if (!bIsSimple)
				{
					const int cxDropDown = ::GetSystemMetrics (SM_CXVSCROLL);

					if (GetExStyle() & (WS_EX_LEFTSCROLLBAR | WS_EX_RIGHT))
					{
						rect.left += cxDropDown;
					}
					else
					{
						rect.right -= cxDropDown;
					}
				}

				DRAWITEMSTRUCT dis;
				memset (&dis, 0, sizeof (DRAWITEMSTRUCT));
				
				dis.CtlType = ODT_COMBOBOX;
				dis.hDC = pDC->GetSafeHdc();
				dis.rcItem = rect;
				dis.itemAction = ODA_DRAWENTIRE;
				dis.CtlID = GetDlgCtrlID();

				int nCurSel = GetCurSel();

				dis.itemID = nCurSel;

				if (nCurSel >= 0)
				{
					dis.itemData = GetItemData(nCurSel);
				}
				
				dis.itemState = (GetSafeHwnd() == CWnd::GetFocus()->GetSafeHwnd()) ? ODS_SELECTED : 0;
				dis.itemState |= ODS_COMBOBOXEDIT;
				
				if (!IsWindowEnabled())
				{
					dis.itemState |= ODS_DISABLED;
				}

				CFont* pOldFont = pDC->SelectObject (GetFont());

				DrawItem(&dis);

				pDC->SelectObject (pOldFont);
				return 0L;
			}
			else
			{
				return Default();
			}
		}

		if (bIsSimple)
		{
			return Default ();
		}
		
		BOOL bDrawPrompt = FALSE;
		
		if (!m_strPrompt.IsEmpty() || !m_strErrorMessage.IsEmpty())
		{
			BOOL bTextIsEmpty = GetWindowTextLength() == 0;
			
			if (GetEditCtrl().GetSafeHwnd () == NULL)
			{
				bDrawPrompt = bTextIsEmpty || !m_strErrorMessage.IsEmpty();
			}
		}

		if (!m_bVisualManagerStyle && !bDrawPrompt)
		{
			return Default();
		}
		
		OnDraw(pDC, bDrawPrompt);
	}

	return 0;
}
//*******************************************************************************
COLORREF CBCGPComboBox::OnFillLbItem(CDC* pDC, int nIndex, CRect rect, BOOL bIsHighlihted, BOOL bIsSelected)
{
	ASSERT_VALID (pDC);

	COLORREF clrText = globalData.clrBarText;

	if (m_bOnGlass)
	{
		CBCGPDrawManager dm(*pDC);
		dm.DrawRect (rect, globalData.clrWindow, (COLORREF)-1);

		return IsWindowEnabled() ? globalData.clrWindowText : globalData.clrGrayedText;
	}

	if (m_bVisualManagerStyle)
	{
		clrText = CBCGPVisualManager::GetInstance ()->OnFillComboBoxItem(pDC, this, nIndex, rect, bIsHighlihted, bIsSelected);

		if (!IsWindowEnabled())
		{
			clrText = globalData.clrGrayedText;
		}

		return clrText;
	}

	CBrush* pBrush = NULL;

	if (bIsSelected)
	{
		pBrush = CBrush::FromHandle(GetSysColorBrush(COLOR_HIGHLIGHT));
		clrText = GetSysColor(COLOR_HIGHLIGHTTEXT);
	}
	else
	{
		pBrush = CBrush::FromHandle(GetSysColorBrush(COLOR_WINDOW));
		clrText = GetSysColor(COLOR_WINDOWTEXT);
	}

	if (pBrush->GetSafeHandle() != NULL)
	{
		pDC->FillRect(rect, pBrush);
	}
	
	if (!IsWindowEnabled())
	{
		clrText = globalData.clrGrayedText;
	}

	return clrText;
}
//*******************************************************************************
void CBCGPComboBox::DrawItem(LPDRAWITEMSTRUCT lpDIS)
{
	ASSERT (lpDIS->CtlType == ODT_COMBOBOX);

	CDC* pDC = CDC::FromHandle(lpDIS->hDC);
	ASSERT_VALID (pDC);

	CRect rc = lpDIS->rcItem;
	int nIndexDC = pDC->SaveDC ();

	BOOL bSelected = (lpDIS->itemState & ODS_SELECTED) == ODS_SELECTED;

	if (m_bDefaultPrintClient && GetDroppedState())
	{
		bSelected = TRUE;
	}

	BOOL bInGlass = m_bOnGlass;
	if ((lpDIS->itemState & ODS_COMBOBOXEDIT) == 0)
	{
		m_bOnGlass = FALSE;
	}

	COLORREF clrText = OnFillLbItem(pDC, (int)lpDIS->itemID, rc, FALSE, bSelected);
	if (clrText != (COLORREF)-1)
	{
		pDC->SetTextColor (clrText);
	}

	m_bOnGlass = bInGlass;

	pDC->SetBkMode(TRANSPARENT);

	int id = (int)lpDIS->itemID;
	if (id >= 0)
	{
		if (!m_bDefaultPrintClient)
		{
			int nIndent = GetItemIndent(id);
			if (GetExStyle() & WS_EX_RIGHT)
			{
				rc.right -= nIndent;
			}
			else
			{
				rc.left += nIndent;
			}
		}

		CSize sizeImage = GetImageSize();
		if (sizeImage != CSize(0, 0))
		{
			CRect rectImage = rc;
			rectImage.top += max(0, (rc.Height () - sizeImage.cy) / 2);
			rectImage.bottom = rectImage.top + sizeImage.cy;
			
			if (GetExStyle() & WS_EX_RIGHT)
			{
				rectImage.right -= m_nHorzPadding;
				rectImage.left = rectImage.right - sizeImage.cx;
			}
			else
			{
				rectImage.left += m_nHorzPadding;
				rectImage.right = rectImage.left + sizeImage.cx;
			}
			
			int nIcon =  GetItemImage(id);
			if (nIcon >= 0)
			{
				const BOOL bIsDisabled = (lpDIS->itemState & ODS_DISABLED) == ODS_DISABLED;

				if (m_hImageList != NULL)
				{
					CImageList* pImageList = CImageList::FromHandle(m_hImageList);
					pImageList->Draw (pDC, nIcon, rectImage.TopLeft (), ILD_TRANSPARENT);
				}
				else if (m_ImageList.GetCount() > 0)
				{
					m_ImageList.DrawEx(pDC, rectImage, nIcon, CBCGPToolBarImages::ImageAlignHorzLeft, CBCGPToolBarImages::ImageAlignVertTop, CRect(0, 0, 0, 0), bIsDisabled ? 127 : 255);
				}
			}
			
			if (GetExStyle() & WS_EX_RIGHT)
			{
				rc.right -= sizeImage.cx + IMAGE_PADDING;
			}
			else
			{
				rc.left += sizeImage.cx + IMAGE_PADDING;
			}
		}

		CString strText;
		GetLBText (id, strText);

		rc.DeflateRect(m_nHorzPadding, 0);

		UINT nFormat = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX;
		
		if (GetExStyle() & WS_EX_RIGHT)
		{
			nFormat |= DT_RIGHT;
		}
		else
		{
			nFormat |= DT_LEFT;
		}

		if (m_bOnGlass)
		{
			CBCGPVisualManager::GetInstance ()->DrawTextOnGlass(pDC, strText, rc, nFormat, 0, clrText);
		}
		else
		{
			pDC->DrawText(strText, rc, nFormat);
		}
		
		DoDrawGroupStart(pDC, lpDIS->rcItem, id);
	}

	pDC->RestoreDC (nIndexDC);
}
//*******************************************************************************
void CBCGPComboBox::MeasureItem(LPMEASUREITEMSTRUCT lpMIS) 
{
	ASSERT(lpMIS->CtlType == ODT_COMBOBOX);
	
	CRect rc;
	GetWindowRect(&rc);
	lpMIS->itemWidth = rc.Width();
	
	lpMIS->itemHeight = max(lpMIS->itemHeight, (UINT)GetDropDownItemMinHeight());
}
//**************************************************************************
LRESULT CBCGPComboBox::OnRedrawFrame(WPARAM wp, LPARAM)
{
	CWnd* pParent = GetParent();
	if (pParent->GetSafeHwnd() != NULL)
	{
		CRect rect;
		GetWindowRect(rect);

		pParent->ScreenToClient(rect);
		pParent->RedrawWindow(rect);

		if (wp != 0)
		{
			RedrawWindow();
		}
	}

	return 0;
}
//**************************************************************************
LRESULT CBCGPComboBox::OnChangeVisualManager (WPARAM, LPARAM)
{
	if (GetEditCtrl().GetSafeHwnd() != NULL)
	{
		GetEditCtrl().SendMessage(BCGM_CHANGEVISUALMANAGER, 0, 0);
	}

	if (m_wndList.GetSafeHwnd() != NULL)
	{
		m_wndList.SendMessage(BCGM_CHANGEVISUALMANAGER, 0, 0);
	}

	return 0L;
}
//**************************************************************************
void CBCGPComboBox::EnableAutoComplete(BOOL bEnable/* = TRUE*/)	// For CBS_DROPDOWN style only
{
	m_bAutoComplete = bEnable;

	if (GetEditCtrl().GetSafeHwnd() != NULL && (GetStyle () & 0x0003L) == CBS_DROPDOWN)
	{
		GetEditCtrl().m_pAutoCompleteCombo = m_bAutoComplete ? this : NULL;
	}
}
//**************************************************************************
int CBCGPComboBox::OnAutoComplete(const CString& strText, CString& strOut)
{
	int nIndex = FindString(-1, strText);
	if (nIndex >= 0)
	{
		GetLBText(nIndex, strOut);
		SetCurSel(nIndex);
	}
	else
	{
		strOut.Empty();
	}

	return nIndex;
}
//**************************************************************************
CSize CBCGPComboBox::GetItemSize(CDC* pDC, int nIndex)
{
	ASSERT_VALID(pDC);

	CSize sizeImage = GetImageSize();
	if (sizeImage.cx > 0)
	{
		sizeImage.cx += IMAGE_PADDING;
	}

	CString str;
	GetLBText(nIndex, str);
	
	CSize sizeItem = pDC->GetTextExtent(str);
	
	TEXTMETRIC tm;
	pDC->GetTextMetrics (&tm);
	
	sizeItem.cx += tm.tmAveCharWidth + 2 * m_nHorzPadding + GetItemIndent(nIndex) + sizeImage.cx;
	sizeItem.cy = max(sizeItem.cy, sizeImage.cy);

	return sizeItem;
}
//**************************************************************************
LRESULT CBCGPComboBox::OnSetFont (WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	LRESULT lres = Default();

	CalcRowHeight();

	if (m_wndList.GetSafeHwnd() != NULL)
	{
		AdjustItemHeight();
	}

	return lres;
}
//***************************************************************************************
BOOL CBCGPComboBox::SetImageList (HIMAGELIST hImageList, int nVertMargin)
{
	m_hImageList = NULL;
	m_ImageList.Clear();
	m_sizeImage = CSize(0, 0);
	m_nImageVertMargin = 0;

	if (hImageList != NULL)
	{
		CImageList* pImageList = CImageList::FromHandle (hImageList);
		if (pImageList != NULL)
		{
			IMAGEINFO info;
			pImageList->GetImageInfo (0, &info);

			CRect rectImage = info.rcImage;
			m_sizeImage = rectImage.Size ();

			m_hImageList = hImageList;
			m_nImageVertMargin = nVertMargin;
		}
	}

	CalcRowHeight();

	if (m_wndList.GetSafeHwnd() != NULL)
	{
		AdjustItemHeight();
	}

	return TRUE;
}
//***************************************************************************************
BOOL CBCGPComboBox::SetImageList(UINT nImageListResID, int cxIcon, int nVertMargin, BOOL bAutoScale)
{
	m_hImageList = NULL;
	m_ImageList.Clear();
	m_sizeImage = CSize(0, 0);
	m_nImageVertMargin = 0;

	if (nImageListResID == 0)
	{
		CalcRowHeight();
		
		if (m_wndList.GetSafeHwnd() != NULL)
		{
			AdjustItemHeight();
		}

		return TRUE;
	}

	if (!m_ImageList.Load(nImageListResID))
	{
		ASSERT (FALSE);
		return FALSE;
	}

#ifndef _BCGSUITE_
	m_ImageList.SetSingleImage(FALSE);
#else
	m_ImageList.SetSingleImage();
#endif

	int cyIcon = m_ImageList.GetImageSize().cy;
	m_sizeImage = CSize(cxIcon, cyIcon);

	m_ImageList.SetImageSize(m_sizeImage, TRUE);

	if (bAutoScale)
	{
		m_sizeImage = globalUtils.ScaleByDPI(m_ImageList);
	}

	m_nImageVertMargin = nVertMargin;

	CalcRowHeight();

	if (m_wndList.GetSafeHwnd() != NULL)
	{
		AdjustItemHeight();
	}

	return TRUE;
}
//***********************************************************************************************************
void CBCGPComboBox::InitItemData(int nIndex)
{
	int nCurSize = (int)m_arItemData.GetSize();
	
	if (nIndex >= nCurSize)
	{
		m_arItemData.SetSize(nIndex + 1);
		
		for (int i = nCurSize; i < nIndex; i++)
		{
			m_arItemData[i] = NULL;
		}
	}
	
	if (m_arItemData[nIndex] == NULL)
	{
		m_arItemData[nIndex] = new CBCGPComboBoxItemData;
	}
}
//***********************************************************************************************************
void CBCGPComboBox::SetItemImage(int nIndex, int nImageIndex)
{
	InitItemData(nIndex);
	m_arItemData[nIndex]->m_nImageIndex = nImageIndex;
}
//***********************************************************************************************************
int CBCGPComboBox::GetItemImage(int nIndex) const
{
	if (nIndex < 0 || nIndex >= (int)m_arItemData.GetSize())
	{
		return -1;
	}

	if (m_arItemData[nIndex] == NULL)
	{
		return -1;
	}

	return m_arItemData[nIndex]->m_nImageIndex;
}
//*********************************************************************************************
LRESULT CBCGPComboBox::OnCBDeleteString(WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = DefWindowProc(CB_DELETESTRING, wParam, lParam);

	int nIndex = (int)wParam;

	if (nIndex >= 0 && nIndex < (int)m_arItemData.GetSize())
	{
		CBCGPComboBoxItemData* pItemData = m_arItemData[nIndex];
		if (pItemData != NULL)
		{
			delete pItemData;
		}

		m_arItemData.RemoveAt(nIndex, 1);
	}

	return lResult;
}
//*********************************************************************************************
LRESULT CBCGPComboBox::OnCBInsertString(WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = DefWindowProc(CB_INSERTSTRING, wParam, lParam);

	if (lResult != CB_ERR)
	{
		int nIndex = (int)lResult;
		if (nIndex >= 0 && nIndex <= (int)m_arItemData.GetSize())
		{
			m_arItemData.InsertAt(nIndex, (CBCGPComboBoxItemData*)NULL);
		}
	}

	return lResult;
}
//*********************************************************************************************
LRESULT CBCGPComboBox::OnCBResetContent(WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = DefWindowProc(CB_RESETCONTENT, wParam, lParam);
	
	CleanUpItemsData();
	return lResult;
}
//*********************************************************************************************
void CBCGPComboBox::SetItemIndent(int nIndex, int nIndent)
{
	InitItemData(nIndex);
	m_arItemData[nIndex]->m_iIndent = nIndent;
}
//*********************************************************************************************
int CBCGPComboBox::GetItemIndent(int nIndex) const
{
	if (nIndex < 0 || nIndex >= (int)m_arItemData.GetSize())
	{
		return -1;
	}
	
	if (m_arItemData[nIndex] == NULL)
	{
		return 0;
	}
	
	return m_arItemData[nIndex]->m_iIndent;
}
//*********************************************************************************************
void CBCGPComboBox::SetItemGroupStart(int nIndex, BOOL bSet /*= TRUE*/)
{
	InitItemData(nIndex);
	m_arItemData[nIndex]->m_bGroupStart = bSet;
}
//*********************************************************************************************
BOOL CBCGPComboBox::IsItemGroupStart(int nIndex) const
{
	if (nIndex < 0 || nIndex >= (int)m_arItemData.GetSize())
	{
		return FALSE;
	}
	
	if (m_arItemData[nIndex] == NULL)
	{
		return FALSE;
	}
	
	return m_arItemData[nIndex]->m_bGroupStart;
	
}
//*********************************************************************************************
void CBCGPComboBox::SetVerticalPadding(int nPadding)
{
	m_nVertPadding = max(0, nPadding);

	if (m_wndList.GetSafeHwnd() != NULL)
	{
		CalcRowHeight();
		AdjustItemHeight();
	}
}
//*********************************************************************************************
void CBCGPComboBox::SetHorizontalPadding(int nPadding)
{
	m_nHorzPadding = max(0, nPadding);
	
	if (m_wndList.GetSafeHwnd() != NULL)
	{
		RedrawWindow();
	}
}
//*********************************************************************************************
int CBCGPComboBox::GetHorizontalPadding() const
{
	return m_nHorzPadding;	
}
//*********************************************************************************************
void CBCGPComboBox::AdjustDroppedWidth()
{
	if (GetSafeHwnd() == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	CRect rectScreen;
	
	MONITORINFO mi;
	mi.cbSize = sizeof (MONITORINFO);
	if (GetMonitorInfo (MonitorFromWindow(GetSafeHwnd (), MONITOR_DEFAULTTONEAREST), &mi))
	{
		rectScreen = mi.rcWork;
	}
	else
	{
		::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectScreen, 0);
	}

	int dx = 0;
	const int cxExtra = ::GetSystemMetrics(SM_CXVSCROLL) + 2 * ::GetSystemMetrics(SM_CXEDGE);

	CClientDC dc(this);
	CBCGPFontSelector fs(dc, GetFont());
	
	for (int i = 0; i < GetCount(); i++)
	{
		dx = max(dx, GetItemSize(&dc, i).cx);

		if (dx >= rectScreen.Width() - cxExtra)
		{
			dx = rectScreen.Width() - cxExtra;
			break;
		}
	}
	
	SetDroppedWidth(dx + cxExtra);
}
//*********************************************************************************************
int CBCGPComboBox::GetVerticalPadding() const
{
	return m_nVertPadding;	
}
//*********************************************************************************************
BOOL CBCGPComboBox::PreTranslateMessage(MSG* pMsg) 
{
	if (m_bIsDroppedDown && pMsg->message == WM_CONTEXTMENU)
	{
		return TRUE;
	}
	
	return CComboBox::PreTranslateMessage(pMsg);
}
//*********************************************************************************************
void CBCGPComboBox::DoDrawGroupStart(CDC* pDC, CRect rect, int nIndex)
{
	ASSERT_VALID(pDC);

	if ((m_bDrawItemOnList || !m_bVisualManagerStyle) && IsItemGroupStart(nIndex))
	{
		CBCGPPenSelector ps(*pDC, m_bVisualManagerStyle ? globalData.clrBarShadow : globalData.clrBtnShadow);
		
		pDC->MoveTo(rect.left, rect.top);
		pDC->LineTo(rect.right, rect.top);
	}
}
//*********************************************************************************************
CSize CBCGPComboBox::GetImageSize() const
{
	if ((m_hImageList != NULL || m_ImageList.GetCount() > 0) && m_sizeImage != CSize(0, 0))
	{
		return m_sizeImage;
	}

	return CSize(0, 0);
}
