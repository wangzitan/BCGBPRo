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

// BCGFontComboBox.cpp : implementation file
//

#include "stdafx.h"
#include "BCGPFontComboBox.h"
#include "BCGPToolBar.h"
#include "BCGPToolbarFontCombo.h"
#include "BCGPGlobalUtils.h"
#include "BCGPLocalResource.h"
#include "bcgprores.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

const int nImageHeight = 15;
const int nImageWidth = 16;
const int nDrawUsingFontExtraHeight = 4;

BOOL CBCGPFontComboBox::m_bDrawUsingFont = FALSE;

#define IMAGE_PADDING	globalUtils.ScaleByDPI(4)

/////////////////////////////////////////////////////////////////////////////
// CBCGPFontComboBox

IMPLEMENT_DYNAMIC(CBCGPFontComboBox, CBCGPComboBox)

CBCGPFontComboBox::CBCGPFontComboBox() :
	m_bToolBarMode (FALSE)
{
}

CBCGPFontComboBox::~CBCGPFontComboBox()
{
}

BEGIN_MESSAGE_MAP(CBCGPFontComboBox, CBCGPComboBox)
	//{{AFX_MSG_MAP(CBCGPFontComboBox)
	ON_WM_CREATE()
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPFontComboBox message handlers

BOOL CBCGPFontComboBox::PreTranslateMessage(MSG* pMsg)
{
	if (m_bToolBarMode &&
		pMsg->message == WM_KEYDOWN &&
		!CBCGPToolbarFontCombo::IsFlatMode ())
	{
		CBCGPToolBar* pBar = (CBCGPToolBar*) GetParent ();

		switch (pMsg->wParam)
		{
		case VK_ESCAPE:
			if (BCGCBProGetTopLevelFrame (this) != NULL)
			{
				BCGCBProGetTopLevelFrame (this)->SetFocus ();
			}
			return TRUE;

		case VK_TAB:
			if (pBar != NULL)
			{
				pBar->GetNextDlgTabItem (this)->SetFocus();
			}
			return TRUE;

		case VK_UP:
		case VK_DOWN:
			if ((GetKeyState(VK_MENU) >= 0) && (GetKeyState(VK_CONTROL) >=0) &&
				!GetDroppedState())
			{
				ShowDropDown();
				return TRUE;
			}
		}
	}

	return CBCGPComboBox::PreTranslateMessage(pMsg);
}
//*****************************************************************************************
int CBCGPFontComboBox::CompareItem(LPCOMPAREITEMSTRUCT lpCIS)
{
	ASSERT(lpCIS->CtlType == ODT_COMBOBOX);

	int id1 = (int)(WORD)lpCIS->itemID1;
	if (id1 == -1)
	{
		return -1;
	}

	CString str1;
	GetLBText (id1, str1);

	int id2 = (int)(WORD)lpCIS->itemID2;
	if (id2 == -1)
	{
		return 1;
	}

	CString str2;
	GetLBText (id2, str2);

	return str1.Collate (str2);
}
//******************************************************************************************
void CBCGPFontComboBox::DrawItem(LPDRAWITEMSTRUCT lpDIS)
{
	if (m_Images.GetCount() == 0)
	{
		CBCGPLocalResource locaRes;
		
		m_Images.SetImageSize(CSize(nImageWidth, nImageHeight));
		
		if (globalData.Is32BitIcons())
		{
			m_Images.Load(IDB_BCGBARRES_FONT32);
			globalUtils.ScaleByDPI(m_Images);
		}
		else
		{
			m_Images.SetTransparentColor(RGB (255, 255, 255));
			m_Images.Load(IDB_BCGBARRES_FONT);
		}
	}

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
	
	COLORREF clrText = OnFillLbItem(pDC, (int)lpDIS->itemID, rc, FALSE, bSelected);
	if (clrText != (COLORREF)-1)
	{
		pDC->SetTextColor (clrText);
	}

	pDC->SetBkMode(TRANSPARENT);

	int id = (int)lpDIS->itemID;
	if (id >= 0)
	{
		CFont fontSelected;
		CFont* pOldFont = NULL;

		CBCGPFontDesc* pDesc = (CBCGPFontDesc*)lpDIS->itemData;
		if (pDesc != NULL)
		{
			if (pDesc->m_nType & (DEVICE_FONTTYPE | TRUETYPE_FONTTYPE))
			{
				const BOOL bIsDisabled = (lpDIS->itemState & ODS_DISABLED) == ODS_DISABLED;

				m_Images.DrawEx(pDC, rc, pDesc->GetImageIndex (), CBCGPToolBarImages::ImageAlignHorzLeft, CBCGPToolBarImages::ImageAlignVertCenter,
					CRect(0, 0, 0, 0), bIsDisabled ? 127 : 255);
			}

			rc.left += globalUtils.ScaleByDPI(nImageWidth) + IMAGE_PADDING;
			
			if (m_bDrawUsingFont && pDesc->m_nCharSet != SYMBOL_CHARSET)
			{
				CreateItemFont(fontSelected, pDesc);
				pOldFont = pDC->SelectObject (&fontSelected);
			}
		}

		CString strText;
		GetLBText (id, strText);

		pDC->DrawText (strText, rc, DT_SINGLELINE | DT_VCENTER);

		if (pOldFont != NULL)
		{
			pDC->SelectObject (pOldFont);
		}

		DoDrawGroupStart(pDC, lpDIS->rcItem, id);
	}

	pDC->RestoreDC (nIndexDC);
}
//****************************************************************************************
void CBCGPFontComboBox::MeasureItem(LPMEASUREITEMSTRUCT lpMIS)
{
	ASSERT(lpMIS->CtlType == ODT_COMBOBOX);

	CRect rc;
	GetWindowRect (&rc);
	lpMIS->itemWidth = rc.Width();

	int nTextHeight = globalData.GetTextHeight();
	if (m_bDrawUsingFont)
	{
		nTextHeight += nDrawUsingFontExtraHeight;
	}

	lpMIS->itemHeight = max(globalUtils.ScaleByDPI(nImageHeight), nTextHeight);
}
//****************************************************************************************
void CBCGPFontComboBox::PreSubclassWindow() 
{
	CBCGPComboBox::PreSubclassWindow();

	_AFX_THREAD_STATE* pThreadState = AfxGetThreadState ();
	if (pThreadState->m_pWndInit == NULL)
	{
		Init ();
	}
}
//****************************************************************************************
int CBCGPFontComboBox::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CBCGPComboBox::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	Init ();
	return 0;
}
//***************************************************************************************
void CBCGPFontComboBox::Init ()
{
	CWnd* pWndParent = GetParent ();
	ASSERT_VALID (pWndParent);

	m_bToolBarMode = pWndParent->IsKindOf (RUNTIME_CLASS (CBCGPToolBar));
	if (!m_bToolBarMode)
	{
		Setup ();
	}
}
//***************************************************************************************
void CBCGPFontComboBox::CleanUp ()
{
	ASSERT_VALID (this);
	ASSERT (::IsWindow (m_hWnd));
	
	if (m_bToolBarMode)
	{
		// Font data will be destroyed by CBCGPToolbarFontCombo object
		return;
	}

	for (int i = 0; i < GetCount (); i++)
	{
		delete (CBCGPFontDesc*) GetItemData (i);
	}

	ResetContent ();
}
//***************************************************************************************
BOOL CBCGPFontComboBox::Setup (int nFontType, BYTE nCharSet, BYTE nPitchAndFamily)
{
	ASSERT_VALID (this);
	ASSERT (::IsWindow (m_hWnd));
	
	if (m_bToolBarMode)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	CleanUp ();

	CBCGPToolbarFontCombo combo (0, (UINT)-1, nFontType, nCharSet,
								CBS_DROPDOWN, 0, nPitchAndFamily);

	for (int i = 0; i < combo.GetCount (); i++)
	{
		CString strFont = combo.GetItem (i);

		CBCGPFontDesc* pFontDescrSrc = (CBCGPFontDesc*) combo.GetItemData (i);
		ASSERT_VALID (pFontDescrSrc);

		if (FindStringExact (-1, strFont) <= 0)
		{
			CBCGPFontDesc* pFontDescr = new CBCGPFontDesc (*pFontDescrSrc);
			int iIndex = AddString (strFont);
			SetItemData (iIndex, (DWORD_PTR) pFontDescr);
		}
	}

	return TRUE;
}
//***************************************************************************************
void CBCGPFontComboBox::OnDestroy() 
{
	CleanUp ();
	CBCGPComboBox::OnDestroy();
}
//***************************************************************************************
BOOL CBCGPFontComboBox::SelectFont (CBCGPFontDesc* pDesc)
{
	ASSERT_VALID (this);
	ASSERT (::IsWindow (m_hWnd));
	ASSERT_VALID (pDesc);

	for (int i = 0; i < GetCount (); i++)
	{
		CBCGPFontDesc* pFontDescr = (CBCGPFontDesc*) GetItemData (i);
		ASSERT_VALID (pFontDescr);

		if (*pDesc == *pFontDescr)
		{
			SetCurSel (i);
			return TRUE;
		}
	}

	return FALSE;
}
//***************************************************************************************
BOOL CBCGPFontComboBox::SelectFont (LPCTSTR lpszName, BYTE nCharSet/* = DEFAULT_CHARSET*/)
{
	ASSERT_VALID (this);
	ASSERT (::IsWindow (m_hWnd));
	ASSERT (lpszName != NULL);

	for (int i = 0; i < GetCount (); i++)
	{
		CBCGPFontDesc* pFontDescr = (CBCGPFontDesc*) GetItemData (i);
		ASSERT_VALID (pFontDescr);

		if (pFontDescr->m_strName == lpszName)
		{
			if (nCharSet == DEFAULT_CHARSET ||
				nCharSet == pFontDescr->m_nCharSet)
			{
				SetCurSel (i);
				return TRUE;
			}
		}
	}

	return FALSE;
}
//***************************************************************************************
CBCGPFontDesc* CBCGPFontComboBox::GetSelFont () const
{
	ASSERT_VALID (this);
	ASSERT (::IsWindow (m_hWnd));

	int iIndex = GetCurSel ();
	if (iIndex < 0)
	{
		return NULL;
	}

	return (CBCGPFontDesc*) GetItemData (iIndex);
}
//**************************************************************************
BOOL CBCGPFontComboBox::CreateItemFont(CFont& font, CBCGPFontDesc* pDesc) const
{
	ASSERT(pDesc != NULL);

	LOGFONT lf;
	globalData.fontRegular.GetLogFont (&lf);
	
	lstrcpyn(lf.lfFaceName, pDesc->m_strName, LF_FACESIZE);
	
	if (pDesc->m_nCharSet != DEFAULT_CHARSET)
	{
		lf.lfCharSet = pDesc->m_nCharSet;
	}
	
	if (lf.lfHeight < 0)
	{
		lf.lfHeight -= nDrawUsingFontExtraHeight;
	}
	else
	{
		lf.lfHeight += nDrawUsingFontExtraHeight;
	}
	
	return font.CreateFontIndirect (&lf);
}
//**************************************************************************
CSize CBCGPFontComboBox::GetItemSize(CDC* pDC, int nIndex)
{
	ASSERT_VALID(pDC);

	CSize sizeImage = m_Images.GetImageSize();
	if (sizeImage.cx != 0)
	{
		sizeImage.cx += IMAGE_PADDING;
	}

	CString str;
	GetLBText(nIndex, str);
	
	CFont fontSelected;
	CFont* pOldFont = NULL;
	
	CBCGPFontDesc* pDesc = (CBCGPFontDesc*)GetItemData(nIndex);
	if (m_bDrawUsingFont && pDesc != NULL && pDesc->m_nCharSet != SYMBOL_CHARSET)
	{
		CreateItemFont(fontSelected, pDesc);
		pOldFont = pDC->SelectObject (&fontSelected);
	}

	CSize sizeItem = pDC->GetTextExtent(str);
	
	if (pOldFont != NULL)
	{
		pDC->SelectObject (pOldFont);
	}

	TEXTMETRIC tm;
	pDC->GetTextMetrics (&tm);
	
	sizeItem.cx += tm.tmAveCharWidth + 2 * m_nHorzPadding + GetItemIndent(nIndex) + sizeImage.cx;
	sizeItem.cy = max(sizeItem.cy, sizeImage.cy);

	return sizeItem;
}
