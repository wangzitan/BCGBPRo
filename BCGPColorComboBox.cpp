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
// BCGPColorComboBox.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPColorComboBox.h"
#include "BCGPGlobalUtils.h"
#include "BCGPGraphicsManager.h"

#ifndef _BCGSUITE_
#include "BCGPColorBar.h"
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPColorComboBox

CMap<UINT, UINT, CString, const CString&> CBCGPColorComboBox::m_mapSystemColorNames;

#ifndef _BCGSUITE_
	CMap<int, int, CString, const CString&> CBCGPColorComboBox::m_mapAccentColorNames;
#endif

CBCGPColorComboBox::CBCGPColorComboBox()
{
	m_bIsNoneColor = FALSE;
	m_clrDefault = RGB(255, 255, 255);

	if (m_mapSystemColorNames.IsEmpty())
	{
		m_mapSystemColorNames.SetAt(COLOR_SCROLLBAR, _T("Scrollbar"));
		m_mapSystemColorNames.SetAt(COLOR_BACKGROUND, _T("Background"));
		m_mapSystemColorNames.SetAt(COLOR_ACTIVECAPTION, _T("Active caption"));
		m_mapSystemColorNames.SetAt(COLOR_INACTIVECAPTION, _T("Inactive caption"));
		m_mapSystemColorNames.SetAt(COLOR_MENU, _T("Menu"));
		m_mapSystemColorNames.SetAt(COLOR_WINDOW, _T("Window"));
		m_mapSystemColorNames.SetAt(COLOR_WINDOWFRAME, _T("Window frame"));
		m_mapSystemColorNames.SetAt(COLOR_MENUTEXT, _T("Menu text"));
		m_mapSystemColorNames.SetAt(COLOR_WINDOWTEXT, _T("Window text"));
		m_mapSystemColorNames.SetAt(COLOR_CAPTIONTEXT, _T("Caption text"));
		m_mapSystemColorNames.SetAt(COLOR_ACTIVEBORDER, _T("Active border"));
		m_mapSystemColorNames.SetAt(COLOR_INACTIVEBORDER, _T("Inactive border"));
		m_mapSystemColorNames.SetAt(COLOR_APPWORKSPACE, _T("Application workspace"));
		m_mapSystemColorNames.SetAt(COLOR_HIGHLIGHT, _T("Highlight"));
		m_mapSystemColorNames.SetAt(COLOR_HIGHLIGHTTEXT, _T("Highlight text"));
		m_mapSystemColorNames.SetAt(COLOR_BTNFACE, _T("Button face"));
		m_mapSystemColorNames.SetAt(COLOR_BTNSHADOW, _T("Button shadow"));
		m_mapSystemColorNames.SetAt(COLOR_GRAYTEXT, _T("Grayed text"));
		m_mapSystemColorNames.SetAt(COLOR_BTNTEXT, _T("Button text"));
		m_mapSystemColorNames.SetAt(COLOR_INACTIVECAPTIONTEXT, _T("Inactive caption text"));
		m_mapSystemColorNames.SetAt(COLOR_BTNHIGHLIGHT, _T("Button highlight"));
		m_mapSystemColorNames.SetAt(COLOR_3DDKSHADOW, _T("3D dark shadow"));
		m_mapSystemColorNames.SetAt(COLOR_3DLIGHT, _T("3D light"));
		m_mapSystemColorNames.SetAt(COLOR_INFOTEXT, _T("Info text"));
		m_mapSystemColorNames.SetAt(COLOR_INFOBK, _T("Info background"));
		m_mapSystemColorNames.SetAt(26/*COLOR_HOTLIGHT*/, _T("Hot light"));
		m_mapSystemColorNames.SetAt(27/*COLOR_GRADIENTACTIVECAPTION*/, _T("Gradient inactive caption"));
		m_mapSystemColorNames.SetAt(28/*COLOR_GRADIENTINACTIVECAPTION*/, _T("Gradient active caption"));
		m_mapSystemColorNames.SetAt(29/*COLOR_MENUHILIGHT*/, _T("Menu light"));
		m_mapSystemColorNames.SetAt(30/*COLOR_MENUBAR*/, _T("Menu bar"));
	}

#ifndef _BCGSUITE_
	if (m_mapAccentColorNames.IsEmpty())
	{
		m_mapAccentColorNames.SetAt(-1, _T("Default"));
		m_mapAccentColorNames.SetAt(0, _T("Blue"));
		m_mapAccentColorNames.SetAt(1, _T("Brown"));
		m_mapAccentColorNames.SetAt(2, _T("Green"));
		m_mapAccentColorNames.SetAt(3, _T("Lime"));
		m_mapAccentColorNames.SetAt(4, _T("Magenta"));
		m_mapAccentColorNames.SetAt(5, _T("Orange"));
		m_mapAccentColorNames.SetAt(6, _T("Pink"));
		m_mapAccentColorNames.SetAt(7, _T("Purple"));
		m_mapAccentColorNames.SetAt(8, _T("Red"));
		m_mapAccentColorNames.SetAt(9, _T("Teal"));
	}
#endif
}

CBCGPColorComboBox::~CBCGPColorComboBox()
{
}

BEGIN_MESSAGE_MAP(CBCGPColorComboBox, CBCGPComboBox)
	//{{AFX_MSG_MAP(CBCGPColorComboBox)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPColorComboBox message handlers

int CBCGPColorComboBox::AddColor(COLORREF color, const CString& strColorName)
{
	ASSERT_VALID(this);
	ASSERT(GetSafeHwnd() != NULL);

	int nIndex = 0;

	if (GetCount() == 0 && m_bIsNoneColor)
	{
		nIndex = AddString(m_strNoneColor);
		if (nIndex >= 0)
		{
			SetItemData(nIndex, (DWORD_PTR)-1);
		}
	}

	nIndex = AddString(strColorName);
	if (nIndex >= 0)
	{
		SetItemData(nIndex, (DWORD_PTR)color);
	}

	return nIndex;
}

#ifdef _BCGSUITE_

class CBCGPColorBarAccess : public CBCGPColorBar
{
public:
	static BOOL GetColorName (COLORREF color, CString& strName)
	{
		return m_ColorNames.Lookup(color, strName);
	}
};
#endif

int CBCGPColorComboBox::SetStandardColors()
{
	ASSERT_VALID(this);
	ASSERT(GetSafeHwnd() != NULL);

	ResetContent();

	const CArray<COLORREF, COLORREF>& arColors = CBCGPColor::GetRGBArray();

	for (int i = 0; i < (int)arColors.GetSize(); i++)
	{
		CString strName;
#ifndef _BCGSUITE_
		CBCGPColorBar::GetColorName(arColors[i], strName);
#else
		CBCGPColorBarAccess::GetColorName(arColors[i], strName);
#endif

		AddColor(arColors[i], strName);
	}

	return (int)arColors.GetSize();
}
//********************************************************************************************
int CBCGPColorComboBox::SetSystemColors()
{
	ASSERT_VALID(this);
	ASSERT(GetSafeHwnd() != NULL);
	
	ResetContent();

	for (int i = 0; i <= (int)m_mapSystemColorNames.GetCount(); i++)
	{
		AddColor(GetSysColor(i), GetSystemColorName(i));

		if (i == 24)
		{
			i++;
		}
	}
	
	return (int)GetCount();
}

#ifndef _BCGSUITE_
int CBCGPColorComboBox::SetAccentColors(BOOL bIncludeDefault /*= TRUE*/)
{
	ASSERT_VALID(this);
	ASSERT(GetSafeHwnd() != NULL);
	
	ResetContent();
	
	for (int i = (bIncludeDefault ? - 1 : 0); i < (int)m_mapAccentColorNames.GetCount() - 1; i++)
	{
		CBCGPVisualManagerVS2012::AccentColor color = (CBCGPVisualManagerVS2012::AccentColor)i;

		AddColor(CBCGPVisualManagerVS2012::AccentColorToRGB(color), GetAccentColorName(color));
	}

	return (int)GetCount();
}
#endif

COLORREF CBCGPColorComboBox::GetSelectedColor() const
{
	ASSERT_VALID(this);
	ASSERT(GetSafeHwnd() != NULL);
	
	int nCurSel = GetCurSel();
	if (nCurSel < 0)
	{
		return CLR_NONE;
	}

	return GetItemColor(nCurSel);
}
//********************************************************************************************
BOOL CBCGPColorComboBox::DoDrawItem(CDC* pDC, CRect rect, CFont* pFont, COLORREF color, const CString& strText, BOOL bDisabled, COLORREF clrDefault)
{
	ASSERT_VALID(pDC);

	const int nPadding = globalUtils.ScaleByDPI(2);
	
	CRect rectColor = rect;
	rectColor.right = rectColor.left + rectColor.Height();
	
	rectColor.DeflateRect(nPadding, nPadding);
	
	COLORREF clrBorder = !bDisabled ? globalData.clrBarText : globalData.clrGrayedText;
	
	if (!bDisabled)
	{
		if (color == (COLORREF)-1)
		{
			CPen pen (PS_SOLID, 0, RGB (255, 0, 0));
			CPen* pOldPen = pDC->SelectObject (&pen);
			
			CRect rectInt = rectColor;
			
			pDC->MoveTo (rectInt.right - 1, rectInt.top);
			pDC->LineTo (rectInt.left - 1, rectInt.bottom);
			
			pDC->SelectObject (pOldPen);
		}
		else
		{
			CBrush br(color == CLR_DEFAULT ? clrDefault : color);
			pDC->FillRect(rectColor, &br);
		}
	}
	
	pDC->Draw3dRect(rectColor, clrBorder, clrBorder);
	
	CFont* pOldFont = pDC->SelectObject(pFont == NULL ? &globalData.fontRegular : pFont);
	
	CRect rectText = rect;
	rectText.left = rectColor.right + 2 * nPadding;

	BOOL bTextTrancated = pDC->GetTextExtent(strText).cx > rectText.Width();
	
	pDC->DrawText (strText, rectText, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);

	pDC->SelectObject (pOldFont);

	return !bTextTrancated;
}
//********************************************************************************************
void CBCGPColorComboBox::DrawItem(LPDRAWITEMSTRUCT lpDIS) 
{
	CDC* pDC = CDC::FromHandle (lpDIS->hDC);
	ASSERT_VALID (pDC);
	
	CRect rect = lpDIS->rcItem;
	int nIndex = lpDIS->itemID;

	BOOL bSelected = (lpDIS->itemState & ODS_SELECTED) == ODS_SELECTED;
	
	COLORREF clrText = OnFillLbItem(pDC, nIndex, rect, FALSE, bSelected);
	
	if (nIndex < 0 || ((GetStyle() & 0x0003L) == CBS_DROPDOWN) && m_bDefaultPrintClient)
	{
		return;
	}

	pDC->SetBkMode (TRANSPARENT);
	pDC->SetTextColor (clrText);

	CString str;
	GetLBText(nIndex, str);

	DoDrawItem(pDC, rect, GetFont(), (COLORREF)lpDIS->itemData, str, !IsWindowEnabled(), m_clrDefault);
	DoDrawGroupStart(pDC, rect, nIndex);
}
//********************************************************************************************
CSize CBCGPColorComboBox::GetImageSize() const
{
	ASSERT_VALID(this);

	int cy = GetItemHeight(-1);
	return CSize(cy, cy);
}
//********************************************************************************************
COLORREF CBCGPColorComboBox::GetItemColor(int nIndex) const
{
	ASSERT_VALID(this);
	ASSERT(GetSafeHwnd() != NULL);

	return (COLORREF)GetItemData(nIndex);
}
//********************************************************************************************
int CBCGPColorComboBox::FindIndexByColor(COLORREF color) const
{
	ASSERT_VALID(this);
	ASSERT(GetSafeHwnd() != NULL);

	for (int i = 0; i < GetCount(); i++)
	{
		if (GetItemColor(i) == color)
		{
			return i;
		}
	}
	
	return -1;
}
//********************************************************************************************
void CBCGPColorComboBox::SetSystemColorName(int nIndex, const CString& strName)
{
	m_mapSystemColorNames.SetAt(nIndex, strName);
}
//********************************************************************************************
const CString& CBCGPColorComboBox::GetSystemColorName(int nIndex)
{
	return m_mapSystemColorNames[nIndex];
}

#ifndef _BCGSUITE_

void CBCGPColorComboBox::SetAccentColorName(CBCGPVisualManagerVS2012::AccentColor color, const CString& strName)
{
	m_mapAccentColorNames.SetAt((int)color, strName);
}
//********************************************************************************************
const CString& CBCGPColorComboBox::GetAccentColorName(CBCGPVisualManagerVS2012::AccentColor color)
{
	return m_mapAccentColorNames[(int)color];
}

#endif

void CBCGPColorComboBox::EnableNoneColor(BOOL bEnable, const CString& strName)
{
	ASSERT_VALID(this);

	if (GetSafeHwnd() != NULL && GetCount() != 0)
	{
		ASSERT(FALSE);
		return;
	}

	m_bIsNoneColor = bEnable;

	if (m_bIsNoneColor)
	{
		m_strNoneColor = strName;
	}
}
//********************************************************************************************
void CBCGPColorComboBox::SetDefaultColor(COLORREF clrDefault)
{
	ASSERT_VALID(this);
	m_clrDefault = clrDefault;
}
