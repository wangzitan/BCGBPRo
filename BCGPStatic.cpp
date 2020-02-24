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
// BCGPStatic.cpp : implementation file
//

#include "stdafx.h"
#include "Bcgglobals.h"
#include "BCGPVisualManager.h"
#include "BCGPStatic.h"
#include "BCGPDlgImpl.h"
#include "BCGPGlobalUtils.h"
#include "BCGPDrawManager.h"
#include "MenuImages.h"

#ifndef _BCGSUITE_
	#include "BCGPTooltipManager.h"
	#include "BCGPToolTipCtrl.h"
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

IMPLEMENT_DYNAMIC(CBCGPStatic, CStatic)

/////////////////////////////////////////////////////////////////////////////
// CBCGPStatic

CBCGPStatic::CBCGPStatic()
{
	m_bOnGlass = FALSE;
	m_bVisualManagerStyle = FALSE;
	m_clrText = (COLORREF)-1;
	m_bBackstageMode = FALSE;
	m_hFont	= NULL;
#ifndef _BCGSUITE_
	m_bDrawTextUnderline = globalData.m_bSysUnderlineKeyboardShortcuts;
#else
	m_bDrawTextUnderline = TRUE;
#endif
	m_bPreview = FALSE;
}

CBCGPStatic::~CBCGPStatic()
{
}

BEGIN_MESSAGE_MAP(CBCGPStatic, CStatic)
	//{{AFX_MSG_MAP(CBCGPStatic)
	ON_WM_ERASEBKGND()
	ON_WM_PAINT()
	ON_WM_ENABLE()
	ON_WM_NCPAINT()
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLAERO, OnBCGSetControlAero)
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLVMMODE, OnBCGSetControlVMMode)
	ON_REGISTERED_MESSAGE(BCGM_ONSETCONTROLBACKSTAGEMODE, OnBCGSetControlBackStageMode)
	ON_MESSAGE(WM_SETTEXT, OnSetText)
	ON_MESSAGE(WM_SETFONT, OnSetFont)
	ON_MESSAGE(WM_PRINT, OnPrint)
	ON_REGISTERED_MESSAGE(BCGM_ONCHANGEUNDERLINE, OnBCGOnChangeUnderline)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPStatic message handlers

BOOL CBCGPStatic::OnEraseBkgnd(CDC* pDC) 
{
	if (!m_bOnGlass)
	{
		return CStatic::OnEraseBkgnd (pDC);
	}

	return TRUE;
}
//*****************************************************************************
void CBCGPStatic::SizeToContent(BCGP_SIZE_TO_CONTENT_LINES lines/* = BCGP_SIZE_TO_CONTENT_LINES_AUTO*/)
{
	ASSERT (GetSafeHwnd () != NULL);

	if (m_Picture.GetCount() > 0)
	{
		CSize sizeImage = GetPictureSize();
		SetWindowPos(NULL, -1, -1, sizeImage.cx, sizeImage.cy, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
		return;
	}

	const DWORD dwStyle = GetStyle ();
	
	if ((dwStyle & SS_BLACKRECT) == SS_BLACKRECT ||
		(dwStyle & SS_GRAYRECT) == SS_GRAYRECT ||
		(dwStyle & SS_WHITERECT) == SS_WHITERECT ||
		(dwStyle & SS_BLACKFRAME) == SS_BLACKFRAME ||
		(dwStyle & SS_GRAYFRAME) == SS_GRAYFRAME ||
		(dwStyle & SS_WHITEFRAME) == SS_WHITEFRAME ||
		(dwStyle & SS_ICON) == SS_ICON ||
		(dwStyle & SS_BITMAP) == SS_BITMAP ||
		(dwStyle & SS_USERITEM) == SS_USERITEM ||
		(dwStyle & SS_ENHMETAFILE) == SS_ENHMETAFILE ||
		(dwStyle & SS_ETCHEDFRAME) == SS_ETCHEDFRAME ||
		(dwStyle & SS_ETCHEDHORZ) == SS_ETCHEDHORZ ||
		(dwStyle & SS_ETCHEDVERT) == SS_ETCHEDVERT)
	{
		return;
	}

	CString strText;
	GetWindowText(strText);
	
	if (strText.IsEmpty())
	{
		return;
	}
	
	CRect rectText;
	GetClientRect(rectText);

	CClientDC dc(this);

	CFont* pOldFont = m_hFont == NULL ?
		(CFont*) dc.SelectStockObject (DEFAULT_GUI_FONT) :
		dc.SelectObject (CFont::FromHandle (m_hFont));
	ASSERT(pOldFont != NULL);

	BOOL bIsSingleLine = TRUE;

	if (lines == BCGP_SIZE_TO_CONTENT_LINES_AUTO)
	{
		TEXTMETRIC tm;
		dc.GetTextMetrics(&tm);

		bIsSingleLine = rectText.Height() < 2 * tm.tmHeight;
	}
	else if (lines == BCGP_SIZE_TO_CONTENT_LINES_MULTI)
	{
		bIsSingleLine = FALSE;
	}

	BOOL bHasEllipsis = ((dwStyle & SS_ENDELLIPSIS) == SS_ENDELLIPSIS) || ((dwStyle & SS_WORDELLIPSIS) == SS_WORDELLIPSIS) || ((dwStyle & SS_PATHELLIPSIS) == SS_PATHELLIPSIS);
	BOOL bNoWordWrap = (dwStyle & SS_LEFTNOWORDWRAP) == SS_LEFTNOWORDWRAP;
	UINT uiDTFlags = (bIsSingleLine || bHasEllipsis || bNoWordWrap) ? 0 : DT_WORDBREAK;

	if (!m_bDrawTextUnderline)
	{
		globalUtils.RemoveSingleAmp(strText, FALSE);
	}
	
	if (strText.Find(_T('\t')) >= 0)
	{
		uiDTFlags |= (DT_TABSTOP | 0x40);
	}

	if (bIsSingleLine)
	{
		rectText.right = INT_MAX;
	}
	
	dc.DrawText(strText, rectText, uiDTFlags | DT_CALCRECT);
	CSize sizeText = rectText.Size();

	SetWindowPos(NULL, -1, -1, sizeText.cx, sizeText.cy, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);

	dc.SelectObject (pOldFont);
}
//*****************************************************************************
BOOL CBCGPStatic::SetPicture(UINT nResID, BOOL bDPIAutoScale/* = TRUE*/, BOOL bAutoResize/* = FALSE*/, HINSTANCE hResInstance/* = NULL*/)
{
	BOOL bRes = TRUE;

	if (nResID == 0)
	{
		m_Picture.Clear();
	}
	else
	{
		
		bRes = m_Picture.Load(nResID, hResInstance);
		if (bRes)
		{
#ifndef _BCGSUITE_
			m_Picture.SetSingleImage(FALSE);
#else
			m_Picture.SetSingleImage();
#endif

			if (bDPIAutoScale)
			{
				globalUtils.ScaleByDPI(m_Picture);
			}
		}
	}
	
	if (GetSafeHwnd() != NULL)
	{
		if (bRes && bAutoResize)
		{
			CSize sizeImage = GetPictureSize();
			SetWindowPos(NULL, -1, -1, sizeImage.cx, sizeImage.cy, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
		}

		RedrawWindow();
	}
	
	return bRes;
}
//*****************************************************************************
BOOL CBCGPStatic::SetPicture(const CBCGPToolBarImages& src)
{
	BOOL bRes = ((CBCGPToolBarImages&)src).CopyTo(m_Picture);
	if (bRes)
	{
#ifndef _BCGSUITE_
		m_Picture.SetSingleImage(FALSE);
#else
		m_Picture.SetSingleImage();
#endif
	}

	if (GetSafeHwnd() != NULL)
	{
		RedrawWindow();
	}

	return bRes;
}
//*****************************************************************************
BOOL CBCGPStatic::SetPicture(const CString& strImagePath, BOOL bDPIAutoScale /*= TRUE*/, BOOL bAutoResize /*= FALSE*/)
{
	BOOL bRes = m_Picture.Load(strImagePath);
	if (bRes)
	{
#ifndef _BCGSUITE_
		m_Picture.SetSingleImage(FALSE);
#else
		m_Picture.SetSingleImage();
#endif
		
		if (bDPIAutoScale)
		{
			globalUtils.ScaleByDPI(m_Picture);
		}
	}

	if (GetSafeHwnd() != NULL)
	{
		if (bRes && bAutoResize)
		{
			CSize sizeImage = GetPictureSize();
			SetWindowPos(NULL, -1, -1, sizeImage.cx, sizeImage.cy, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
		}
		
		RedrawWindow();
	}

	return bRes;
}
//*****************************************************************************
BOOL CBCGPStatic::IsOwnerDraw() const
{
	return (m_bVisualManagerStyle || m_bOnGlass || m_clrText != (COLORREF)-1) || m_Picture.GetCount() > 0;
}
//*****************************************************************************
void CBCGPStatic::OnPaint() 
{
	if (!IsOwnerDraw())
	{
		Default ();
		return;
	}

	const DWORD dwStyle = GetStyle ();

	if ((dwStyle & SS_BLACKRECT) == SS_BLACKRECT ||
		(dwStyle & SS_GRAYRECT) == SS_GRAYRECT ||
		(dwStyle & SS_WHITERECT) == SS_WHITERECT ||
		(dwStyle & SS_BLACKFRAME) == SS_BLACKFRAME ||
		(dwStyle & SS_GRAYFRAME) == SS_GRAYFRAME ||
		(dwStyle & SS_WHITEFRAME) == SS_WHITEFRAME ||
		(dwStyle & SS_BITMAP) == SS_BITMAP ||
		(dwStyle & SS_USERITEM) == SS_USERITEM ||
		(dwStyle & SS_ENHMETAFILE) == SS_ENHMETAFILE ||
		(dwStyle & SS_ETCHEDFRAME) == SS_ETCHEDFRAME ||
		(dwStyle & SS_ETCHEDHORZ) == SS_ETCHEDHORZ ||
		(dwStyle & SS_ETCHEDVERT) == SS_ETCHEDVERT)
	{
		if (m_Picture.GetCount() == 0)
		{
			Default ();
			return;
		}
	}

	CPaintDC dc(this); // device context for painting

	CBCGPMemDC memDC (dc, this);
	CDC* pDC = &memDC.GetDC ();

	globalData.DrawParentBackground (this, pDC);
	DoPaint(pDC);
}
//*******************************************************************************************************
COLORREF CBCGPStatic::GetForegroundColor()
{
#if (!defined _BCGSUITE_)
	COLORREF clrText = (m_clrText == (COLORREF)-1) ? CBCGPVisualManager::GetInstance()->GetDlgTextColor(GetParent()) : m_clrText;
	
#if (!defined BCGP_EXCLUDE_RIBBON)
	if (m_bBackstageMode && m_clrText == (COLORREF)-1)
	{
		clrText = CBCGPVisualManager::GetInstance ()->GetRibbonBackstageTextColor();
	}
#endif
	
#else
	COLORREF clrText = (m_clrText == (COLORREF)-1) ? globalData.clrBarText : m_clrText;
#endif
	
	if (!IsWindowEnabled ())
	{
#ifndef _BCGSUITE_
		clrText = CBCGPVisualManager::GetInstance ()->GetToolbarDisabledTextColor ();
#else
		clrText = globalData.clrGrayedText;
#endif
	}

	return clrText;
}
//*******************************************************************************************************
void CBCGPStatic::DoPaint(CDC* pDC)
{
	const DWORD dwStyle = GetStyle ();

	CRect rectClient;
	GetClientRect (rectClient);

	if (m_Picture.GetCount() > 0)
	{
		CSize sizeImage = m_Picture.GetImageSize();
		
		double dblXScale = min(1.0, (double)rectClient.Width() / sizeImage.cx);
		double dblYScale = min(1.0, (double)rectClient.Height() / sizeImage.cy);
		
		double dblScale = min(dblXScale, dblYScale);
		if (dblScale == 1.0)
		{
			if ((dwStyle & SS_CENTERIMAGE) == SS_CENTERIMAGE)
			{
				m_Picture.DrawEx(pDC, rectClient, 0, CBCGPToolBarImages::ImageAlignHorzCenter, CBCGPToolBarImages::ImageAlignVertCenter);
			}
			else
			{
				m_Picture.DrawEx(pDC, rectClient, 0);
			}
		}
		else
		{
			CSize sizeDest(
				(int)(dblScale * sizeImage.cx),
				(int)(dblScale * sizeImage.cy));
			
			if (sizeDest.cx > 1 && sizeDest.cy > 1)
			{
				CPoint pt(rectClient.CenterPoint().x - sizeDest.cx / 2, rectClient.CenterPoint().y - sizeDest.cy / 2);
				
				pDC->SetStretchBltMode(HALFTONE);
				m_Picture.DrawEx(pDC, CRect(pt, sizeDest), 0, CBCGPToolBarImages::ImageAlignHorzStretch, CBCGPToolBarImages::ImageAlignVertStretch);
			}
		}

		return;
	}

	if ((dwStyle & SS_ICON) == SS_ICON)
	{
		HICON hIcon = GetIcon();
		if (hIcon != NULL)
		{
			CSize sizeIcon(0, 0);

			if ((dwStyle & SS_REALSIZEIMAGE) == SS_REALSIZEIMAGE)
			{
				sizeIcon = globalUtils.GetIconSize(hIcon);

				if ((dwStyle & SS_CENTERIMAGE) == SS_CENTERIMAGE)
				{
					rectClient.left += (rectClient.Width() - sizeIcon.cx) / 2;
					rectClient.top += (rectClient.Height() - sizeIcon.cy) / 2;
				}

				rectClient.right = rectClient.left + sizeIcon.cx;
				rectClient.bottom = rectClient.top + sizeIcon.cy;
			}

#ifndef _BCGSUITE_
			if (m_bOnGlass)
			{
				CBCGPVisualManager::GetInstance ()->DrawIconOnGlass(pDC, hIcon, rectClient);
			}
			else
#endif
			{
				::DrawIconEx(pDC->GetSafeHdc(), rectClient.left, rectClient.top, hIcon, rectClient.Width(), rectClient.Height(), 0, NULL, DI_NORMAL);
			}
		}

		return;
	}

	if (m_hFont != NULL && ::GetObjectType (m_hFont) != OBJ_FONT)
	{
		m_hFont = NULL;
	}

	CFont* pOldFont = m_hFont == NULL ?
		(CFont*) pDC->SelectStockObject (DEFAULT_GUI_FONT) :
		pDC->SelectObject (CFont::FromHandle (m_hFont));

	ASSERT(pOldFont != NULL);

	BOOL bHasEllipsis = FALSE;

	UINT uiDTFlags = 0;

	if (dwStyle & SS_CENTER)
	{
		uiDTFlags |= DT_CENTER;
	}
	else if (dwStyle & SS_RIGHT)
	{
		uiDTFlags |= DT_RIGHT;
	}

	if (dwStyle & SS_NOPREFIX)
	{
		uiDTFlags |= DT_NOPREFIX;
	}

	if ((dwStyle & SS_CENTERIMAGE) == SS_CENTERIMAGE)
	{
		uiDTFlags |= DT_SINGLELINE | DT_VCENTER;
	}

	if ((dwStyle & SS_ELLIPSISMASK) == SS_ENDELLIPSIS)
	{
		uiDTFlags |= DT_END_ELLIPSIS;
		bHasEllipsis = TRUE;
	}

	if ((dwStyle & SS_ELLIPSISMASK) == SS_PATHELLIPSIS)
	{
		uiDTFlags |= DT_PATH_ELLIPSIS;
		bHasEllipsis = TRUE;
	}

	if ((dwStyle & SS_ELLIPSISMASK) == SS_WORDELLIPSIS)
	{
		uiDTFlags |= DT_WORD_ELLIPSIS;
		bHasEllipsis = TRUE;
	}

	BOOL bNoWordWrap = (dwStyle & SS_LEFTNOWORDWRAP) == SS_LEFTNOWORDWRAP;

	if (!bHasEllipsis && !bNoWordWrap)
	{
		uiDTFlags |= DT_WORDBREAK;
	}

	COLORREF clrText = GetForegroundColor();

	CString strText;
	GetWindowText (strText);

	if (!m_bDrawTextUnderline)
	{
		globalUtils.RemoveSingleAmp(strText, FALSE);
	}

	if (strText.Find(_T('\t')) >= 0)
	{
		uiDTFlags |= (DT_TABSTOP | 0x40);
	}

	if (!m_bOnGlass)
	{
		COLORREF clrTextOld = pDC->SetTextColor (clrText);
		pDC->SetBkMode (TRANSPARENT);
		pDC->DrawText (strText, rectClient, uiDTFlags);
		pDC->SetTextColor (clrTextOld);
	}
	else
	{
		CBCGPVisualManager::GetInstance ()->DrawTextOnGlass (pDC, strText, rectClient, uiDTFlags, 6, IsWindowEnabled () ? m_clrText : globalData.clrGrayedText);
	}

	pDC->SelectObject (pOldFont);
}
//**************************************************************************
LRESULT CBCGPStatic::OnBCGOnChangeUnderline(WPARAM wp, LPARAM lp)
{
	BOOL bDrawTextUnderline = (wp != 0);
	
	if (m_bDrawTextUnderline != bDrawTextUnderline)
	{
		m_bDrawTextUnderline = bDrawTextUnderline;
		
		BOOL bRedraw = (BOOL)lp;
		if (bRedraw)
		{
			RedrawWindow();
		}
	}

	return 0;	
}
//*****************************************************************************
LRESULT CBCGPStatic::OnBCGSetControlAero (WPARAM wp, LPARAM)
{
	m_bOnGlass = (BOOL) wp;
	return 0;
}
//**************************************************************************
LRESULT CBCGPStatic::OnBCGSetControlVMMode (WPARAM wp, LPARAM)
{
	m_bVisualManagerStyle = (BOOL) wp;
	return 0;
}
//*****************************************************************************
LRESULT CBCGPStatic::OnSetText (WPARAM, LPARAM)
{
	if (m_bBackstageMode && !globalData.IsHighContastMode())
	{
		SetRedraw(FALSE);
	}

	LRESULT lr = Default ();

	if (m_bBackstageMode && !globalData.IsHighContastMode())
	{
		SetRedraw(TRUE);
		RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE);
	}

	if (GetParent () != NULL)
	{
		CRect rect;
		GetWindowRect (rect);

		GetParent ()->ScreenToClient (&rect);
		GetParent()->RedrawWindow(rect, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
	}

	return lr;
}
//*****************************************************************************
void CBCGPStatic::OnEnable(BOOL bEnable) 
{
	CStatic::OnEnable(bEnable);

	if (GetParent () != NULL)
	{
		CRect rect;
		GetWindowRect (rect);

		GetParent ()->ScreenToClient (&rect);
		GetParent ()->RedrawWindow (rect, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
	}
}
//*****************************************************************************
BOOL CBCGPStatic::IsOwnerDrawSeparator(BOOL& bIsHorz)
{
#ifndef _BCGSUITE_
	const DWORD dwStyle = GetStyle ();
	BOOL bIsSeparator = (dwStyle & SS_ETCHEDHORZ) == SS_ETCHEDHORZ || (dwStyle & SS_ETCHEDVERT) == SS_ETCHEDVERT;
	bIsHorz = (dwStyle & SS_ETCHEDVERT) != SS_ETCHEDVERT;
	
	return (bIsSeparator && CBCGPVisualManager::GetInstance ()->IsOwnerDrawDlgSeparator(this));
#else
	UNREFERENCED_PARAMETER(bIsHorz);
	return FALSE;
#endif
}
//*****************************************************************************
void CBCGPStatic::OnNcPaint() 
{
#ifndef _BCGSUITE_
	BOOL bIsHorz = TRUE;
	if (!IsOwnerDrawSeparator(bIsHorz))
	{
		Default();
		return;
	}

	CWindowDC dc(this);

	CRect rect;
	GetWindowRect (rect);

	rect.bottom -= rect.top;
	rect.right -= rect.left;
	rect.left = rect.top = 0;

	CBCGPVisualManager::GetInstance ()->OnDrawDlgSeparator(&dc, this, rect, bIsHorz);
#else
	Default();
#endif
}
//**************************************************************************
LRESULT CBCGPStatic::OnBCGSetControlBackStageMode (WPARAM, LPARAM)
{
	m_bBackstageMode = TRUE;
	return 0;
}
//*****************************************************************************
LRESULT CBCGPStatic::OnSetFont (WPARAM wParam, LPARAM)
{
	m_hFont = (HFONT) wParam;
	return Default();
}
//*****************************************************************************
LRESULT CBCGPStatic::OnPrint(WPARAM wp, LPARAM lp)
{
	const DWORD dwStyle = GetStyle ();
	
	if ((dwStyle & SS_ICON) == SS_ICON ||
		(dwStyle & SS_BLACKRECT) == SS_BLACKRECT ||
		(dwStyle & SS_GRAYRECT) == SS_GRAYRECT ||
		(dwStyle & SS_WHITERECT) == SS_WHITERECT ||
		(dwStyle & SS_BLACKFRAME) == SS_BLACKFRAME ||
		(dwStyle & SS_GRAYFRAME) == SS_GRAYFRAME ||
		(dwStyle & SS_WHITEFRAME) == SS_WHITEFRAME ||
		(dwStyle & SS_USERITEM) == SS_USERITEM)
	{
		if (m_Picture.GetCount() == 0)
		{
			return Default();
		}
	}

	DWORD dwFlags = (DWORD)lp;

	CBCGPRAII<BOOL> raiPreview(m_bPreview, TRUE);
	
#ifndef _BCGSUITE_
	BOOL bIsHorz = TRUE;

	if ((dwFlags & PRF_NONCLIENT) == PRF_NONCLIENT)
	{
		CDC* pDC = CDC::FromHandle((HDC) wp);
		ASSERT_VALID(pDC);
		
		BOOL bIsSeparator = (dwStyle & SS_ETCHEDHORZ) == SS_ETCHEDHORZ || (dwStyle & SS_ETCHEDVERT) == SS_ETCHEDVERT;
		if (bIsSeparator)
		{
			CRect rect;
			GetWindowRect (rect);
			
			rect.bottom -= rect.top;
			rect.right -= rect.left;
			rect.left = rect.top = 0;

			if (IsOwnerDrawSeparator(bIsHorz))
			{
				CBCGPVisualManager::GetInstance ()->OnDrawDlgSeparator(pDC, this, rect, bIsHorz);
			}
			else
			{
				if (bIsHorz)
				{
					rect.bottom = rect.top + 2;
				}
				else
				{
					rect.right = rect.left + 2;
				}

				pDC->Draw3dRect (rect, globalData.clrBarShadow, globalData.clrBarHilite);
			}

			return 0L;
		}
	}

	if (m_bBackstageMode && !IsOwnerDrawSeparator(bIsHorz))
	{
		if ((dwFlags & PRF_CLIENT) == PRF_CLIENT)
		{
			CDC* pDC = CDC::FromHandle((HDC) wp);
			ASSERT_VALID(pDC);

			DoPaint(pDC);
		}

		return 0;
	}
#endif

	if (IsOwnerDraw())
	{
		if ((dwFlags & PRF_CLIENT) == PRF_CLIENT)
		{
			CDC* pDC = CDC::FromHandle((HDC) wp);
			ASSERT_VALID(pDC);
			
			DoPaint(pDC);
			return 0;
		}
	}

	return Default();
}
//**************************************************************************
void CBCGPStatic::PreSubclassWindow()
{
	globalUtils.GetParentDialogFont(m_hWnd, m_hFont);
	CStatic::PreSubclassWindow();
}
//**************************************************************************
void CBCGPStatic::OnDestroy() 
{
	m_hFont = NULL;
	CStatic::OnDestroy();
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPInfoBoxRenderer

static void AdjustPadding(long& nValue)
{
	if (nValue < 0)
	{
		nValue = 10;
	}
}

CBCGPInfoBoxRenderer::CBCGPInfoBoxRenderer()
{
	m_clrBackground = CLR_DEFAULT;
	m_clrForeground = CLR_DEFAULT;
	m_clrCaption = CLR_DEFAULT;
	m_clrFrame = CLR_DEFAULT;
	m_clrLink = CLR_DEFAULT;
	m_rectPadding = CRect(-1, -1, -1, -1);
	m_nHorzAlign = -1;
	m_nVertAlign = -1;
	m_clrColorBar = CLR_NONE;
	m_nColorBarWidth = 5;
	m_bIsColorBarOnLeft = TRUE;
	m_bDrawShadow = FALSE;
	m_nIconIndex = -1;
	m_hInfoBoxFont = NULL;
	m_nRowHeight = 0;
	m_bTextIsTruncated = FALSE;
	m_bDrawMoreTooltipMarker = FALSE;
	m_CaptionStyle = BCGP_INFOBOX_CAPTION_BOLD;
	m_nTextHorzOffset = 0;
	m_bFixedFrameWidth = FALSE;
	m_bFixedFrameHeight = FALSE;
	m_bRoundedCorners = FALSE;

	m_rectMore.SetRectEmpty();
	m_rectLink.SetRectEmpty();

	m_bIsFocused = FALSE;
}
//*******************************************************************************************************
void CBCGPInfoBoxRenderer::SetInfoBoxFont(HFONT hFont)
{
	m_hInfoBoxFont = hFont;
	m_fontBold.DeleteObject();
	m_fontUnderline.DeleteObject();

	m_nRowHeight = 0;
}
//*******************************************************************************************************
CRect CBCGPInfoBoxRenderer::DoDraw(CDC* pDC, const CRect& rectBoundsIn, const CString& strInfo, const CString& strCaptionIn /*= _T("")*/, HICON hIcon /*= NULL*/, BOOL bCalcRectOnly/* = FALSE*/, BOOL bPreview/* = FALSE*/)
{
	ASSERT_VALID(pDC);

	CString strCaption = strCaptionIn;
	strCaption.TrimLeft();
	strCaption.TrimRight();

	m_bTextIsTruncated = FALSE;
	m_rectMore.SetRectEmpty();

	m_rectLink.SetRectEmpty();

	CRgn rgn;

	if (!bCalcRectOnly && !bPreview)
	{
		rgn.CreateRectRgnIndirect(rectBoundsIn);
		pDC->SelectClipRgn(&rgn);
	}

	COLORREF clrBackground = m_clrBackground;
	COLORREF clrFrame = m_clrFrame;
	COLORREF clrText = m_clrForeground;
	COLORREF clrLink = m_clrLink;

	CBCGPVisualManager::GetInstance()->CBCGPVisualManager::GetInfoBoxColors(this, clrBackground, clrText, clrFrame, clrLink);

	CRect rectPadding = m_rectPadding;
	
	AdjustPadding(rectPadding.left);
	AdjustPadding(rectPadding.right);
	AdjustPadding(rectPadding.top);
	AdjustPadding(rectPadding.bottom);

	CRect rectPaddingScaled = globalUtils.ScaleByDPI(rectPadding);

	rectPaddingScaled.left = max(0, rectPaddingScaled.left);
	rectPaddingScaled.right = max(0, rectPaddingScaled.right);
	rectPaddingScaled.top = max(0, rectPaddingScaled.top);
	rectPaddingScaled.bottom = max(0, rectPaddingScaled.bottom);

	const CSize sizeIcon = m_nIconIndex >= 0 ? CBCGPDlgImpl::LoadInfoTipIcons(m_InfotipIcons) : globalUtils.GetIconSize(hIcon);

	const int cxIconArea = sizeIcon.cx > 0 ? sizeIcon.cx + globalUtils.ScaleByDPI(10) : 0;
	const int cxColorBar = (m_clrColorBar != CLR_NONE && m_nColorBarWidth > 0) ? globalUtils.ScaleByDPI(m_nColorBarWidth) : 0;

	if (m_hInfoBoxFont != NULL && ::GetObjectType(m_hInfoBoxFont) != OBJ_FONT)
	{
		m_hInfoBoxFont = NULL;
		m_fontBold.DeleteObject();
		m_fontUnderline.DeleteObject();
	}

	CFont* pFont = CFont::FromHandle(m_hInfoBoxFont == NULL ? (HFONT)::GetStockObject(DEFAULT_GUI_FONT) : m_hInfoBoxFont);
	ASSERT(pFont != NULL);

	CFont* pOldFont = (CFont*)pDC->SelectObject(pFont);
	ASSERT(pOldFont != NULL);

	if (m_nRowHeight <= 0)
	{
		TEXTMETRIC tm;
		pDC->GetTextMetrics(&tm);
		
		m_nRowHeight = tm.tmHeight + tm.tmInternalLeading + 1;
	}
	
	const int nShadowSize = m_bDrawShadow ? globalUtils.ScaleByDPI(3) : 0;

	const UINT nHorzAlign = m_nHorzAlign < 0 ? TA_CENTER : (UINT)m_nHorzAlign;
	const UINT nVertAlign = m_nVertAlign < 0 ? TA_CENTER : (UINT)m_nVertAlign;

	const UINT nHorzAlignText = m_nHorzAlign < 0 ? TA_LEFT : (UINT)m_nHorzAlign;

	CRect rectBounds = rectBoundsIn;
	rectBounds.right -= nShadowSize;
	rectBounds.bottom -= nShadowSize;

	CRect rectText = rectBounds;
	rectText.DeflateRect(rectPaddingScaled);

	rectText.left += cxIconArea;

	if (m_bIsColorBarOnLeft)
	{
		rectText.left += cxColorBar;
	}
	else
	{
		rectText.right -= cxColorBar;
	}

	UINT nDTFlags = DT_CALCRECT | DT_WORDBREAK | DT_NOPREFIX;
	
	switch (nHorzAlignText)
	{
	case TA_CENTER:
		nDTFlags |= DT_CENTER;
		break;
		
	case TA_RIGHT:
		nDTFlags |= DT_RIGHT;
		break;
	}
	
	pDC->DrawText(strInfo, rectText, nDTFlags);
	
	if (!strCaption.IsEmpty())
	{
		CFont* pOldCaptionFont = SelectCaptionFont(pDC, pFont);

		int nCaptionWidth = pDC->GetTextExtent(strCaption).cx;

		if (pOldCaptionFont != NULL)
		{
			pDC->SelectObject(pOldCaptionFont);
		}

		if (nCaptionWidth > rectText.Width())
		{
			rectText.right = rectText.left + nCaptionWidth;
		}
	}

	m_nTextHorzOffset = rectText.left;

	CRect rectLink = rectText;

#ifndef _BCGSUITE_
	const int nCaptionRowsHeight = strCaption.IsEmpty() ? 0 : (m_CaptionStyle == BCGP_INFOBOX_CAPTION_LARGE ? globalData.GetCaptionTextHeight() : m_nRowHeight);
#else
	const int nCaptionRowsHeight = strCaption.IsEmpty() ? 0 : m_nRowHeight;
#endif
	const int nCaptionHeight = !strCaption.IsEmpty() ? (nCaptionRowsHeight + rectPaddingScaled.top / 2) : 0;
	const int nLinkHeight = !m_strLink.IsEmpty() ? pDC->DrawText(m_strLink, rectLink, DT_WORDBREAK | DT_NOPREFIX | DT_CALCRECT) + rectPaddingScaled.top : 0;

	rectText.bottom += nCaptionHeight + nLinkHeight;
	rectText.left -= cxIconArea;

	if (rectText.Height() < sizeIcon.cy)
	{
		rectText.bottom = rectText.top + sizeIcon.cy;
	}

	if (m_bIsColorBarOnLeft)
	{
		rectText.left -= cxColorBar;
	}
	else
	{
		rectText.right += cxColorBar;
	}

	int cxOffset =	nHorzAlign == TA_CENTER ? max(0, (rectBounds.Width() - rectText.Width() - rectPaddingScaled.left - rectPaddingScaled.right) / 2) :
					nHorzAlign == TA_LEFT ? 0 : max(0, (rectBounds.Width() - rectText.Width() - rectPaddingScaled.left - rectPaddingScaled.right));

	int cyOffset =	nVertAlign == TA_CENTER ? max(0, (rectBounds.Height() - rectText.Height() - rectPaddingScaled.top - rectPaddingScaled.bottom) / 2) :
					nVertAlign == TA_TOP ? 0 : max(0, (rectBounds.Height() - rectText.Height() - rectPaddingScaled.top - rectPaddingScaled.bottom));
	
	rectText.OffsetRect(cxOffset, cyOffset);
	
	CRect rectFrame = rectText;
	rectFrame.InflateRect(rectPaddingScaled);

	if (!bCalcRectOnly)
	{
		rectFrame.bottom = min(rectFrame.bottom, rectBounds.bottom);
	}
	
	if (m_bFixedFrameWidth)
	{
		rectFrame.right = max(rectBounds.right, rectFrame.right);
	}

	if (m_bFixedFrameHeight)
	{
		rectFrame.bottom = max(rectBounds.bottom, rectFrame.bottom);
	}

	const int nCornerRadius = m_bRoundedCorners ? globalUtils.ScaleByDPI(10) : 0;

	if (m_bDrawShadow && !bCalcRectOnly)
	{
		int nShadowSizeCurr = nShadowSize;
		CRect rectShadow = rectFrame;
		
		if (nCornerRadius > 0)
		{
			if (clrBackground == CLR_NONE)
			{
				nShadowSizeCurr = 0;
			}
			else
			{
				nShadowSizeCurr += 5 * nCornerRadius / 4;

				rectShadow.right -= nCornerRadius;
				rectShadow.bottom -= nCornerRadius;
			}
		}

		if (nShadowSizeCurr > 0)
		{
			CBCGPDrawManager dm(*pDC);
			dm.DrawShadow (rectShadow, nShadowSizeCurr, 100, 75);
		}
	}

	CRgn rgnFrame;
	if (nCornerRadius > 0 && !bCalcRectOnly)
	{
		rgnFrame.CreateRoundRectRgn(rectFrame.left, rectFrame.top, rectFrame.right + 1, rectFrame.bottom + 1, nCornerRadius, nCornerRadius);
		pDC->SelectClipRgn(&rgnFrame);
	}
	
	if (clrBackground != CLR_NONE && !bCalcRectOnly)
	{
		CBrush brFill(clrBackground);
		pDC->FillRect(rectFrame, &brFill);
	}

	COLORREF clrTextOld = pDC->SetTextColor(clrText);
	int nBkModeOld = pDC->SetBkMode(TRANSPARENT);

	if (cxColorBar > 0)
	{
		CRect rectColorBar = rectFrame;

		if (m_bIsColorBarOnLeft)
		{
			rectText.left += cxColorBar;
			rectColorBar.right = rectColorBar.left + cxColorBar;
		}
		else
		{
			rectText.right -= cxColorBar;
			rectColorBar.left = rectColorBar.right - cxColorBar;
		}

		if (!bCalcRectOnly)
		{
			CBrush br(m_clrColorBar);
			pDC->FillRect(rectColorBar, &br);
		}
	}

	if (sizeIcon.cx > 0)
	{
		if (!bCalcRectOnly)
		{
			CRect rectIcon = rectText;

			int nIconPadding = max(0, (nCaptionRowsHeight - sizeIcon.cy) / 2);
			rectIcon.DeflateRect(0, nIconPadding);
			
			if (m_nIconIndex >= 0)
			{
				m_InfotipIcons.DrawEx(pDC, rectIcon, m_nIconIndex);
			}
			else
			{
				pDC->DrawIcon(rectIcon.TopLeft(), hIcon);
			}
		}

		rectText.left += cxIconArea;
	}
	
	if (clrFrame != CLR_NONE && !bCalcRectOnly)
	{
		if (nCornerRadius > 0)
		{
			CBCGPBrushSelector bs(*pDC, CBrush::FromHandle((HBRUSH)::GetStockObject (HOLLOW_BRUSH)));
			CBCGPPenSelector ps(*pDC, clrFrame);

			pDC->RoundRect(rectFrame, CPoint(nCornerRadius, nCornerRadius));
		}
		else
		{
			pDC->Draw3dRect(rectFrame, clrFrame, clrFrame);
		}
	}
	
	CRect rectCaption = rectText;
	rectCaption.bottom = rectCaption.top + nCaptionHeight;

	rectText.top += nCaptionHeight;
	rectText.bottom = min(rectText.bottom, rectBounds.bottom - rectPaddingScaled.bottom);

	if (nLinkHeight > 0)
	{
		m_rectLink = rectText;
		m_rectLink.top = m_rectLink.bottom - nLinkHeight + rectPaddingScaled.top;

		rectText.bottom -= nLinkHeight;

		const int nLinkWidth = min(rectText.Width(), rectLink.Width());

		switch (nHorzAlign)
		{
		default:
			m_rectLink.right = m_rectLink.left + nLinkWidth;
			break;

		case TA_CENTER:
			m_rectLink.left = m_rectLink.CenterPoint().x - nLinkWidth / 2;
			m_rectLink.right = m_rectLink.left + nLinkWidth;
			break;

		case TA_RIGHT:
			m_rectLink.left = m_rectLink.right - nLinkWidth;
			break;
		}
	}

	if (!bCalcRectOnly)
	{
		nDTFlags = DT_WORDBREAK | DT_END_ELLIPSIS | DT_NOPREFIX;

		switch (nHorzAlignText)
		{
		case TA_CENTER:
			nDTFlags |= DT_CENTER;
			break;

		case TA_RIGHT:
			nDTFlags |= DT_RIGHT;
			break;
		}

		int nFullHeight = rectText.Height();

		int nRealHeight = pDC->DrawText(strInfo, rectText, nDTFlags);

		if (nRealHeight > nFullHeight)
		{
			m_bTextIsTruncated = TRUE;
		}

		m_nTextHorzOffset = rectText.left;

		if (!strCaption.IsEmpty())
		{
			SelectCaptionFont(pDC, pFont);

			COLORREF clrTextCaptionOld = CLR_NONE;
			if (m_clrCaption != CLR_DEFAULT)
			{
				clrTextCaptionOld = pDC->SetTextColor(m_clrCaption);
			}

			nDTFlags = DT_SINGLELINE | DT_LEFT | DT_END_ELLIPSIS;
			
			switch (nHorzAlignText)
			{
			case TA_CENTER:
				nDTFlags |= DT_CENTER;
				break;
				
			case TA_RIGHT:
				nDTFlags |= DT_RIGHT;
				break;
			}

			pDC->DrawText(strCaption, rectCaption, nDTFlags);

			if (clrTextCaptionOld != CLR_NONE)
			{
				pDC->SetTextColor(clrTextCaptionOld);
			}
		}

		if (m_bDrawMoreTooltipMarker && m_bTextIsTruncated)
		{
			m_rectMore = rectFrame;

			if (!m_bIsColorBarOnLeft)
			{
				m_rectMore.right -= cxColorBar;
			}

			const int nPadding = globalUtils.ScaleByDPI(3);
			
			CSize sizeMore = CBCGPMenuImages::Size();
			sizeMore.cx += 2 * nPadding;
			sizeMore.cy += nPadding;

			m_rectMore.DeflateRect(nPadding, nPadding);
			m_rectMore.left = m_rectMore.right - sizeMore.cx;
			m_rectMore.top = m_rectMore.bottom - sizeMore.cy;

			CBCGPMenuImages::IMAGE_STATE state = CBCGPMenuImages::ImageBlack;

			if (clrBackground != CLR_NONE)
			{
				CBrush brFill(clrBackground);
				pDC->FillRect(m_rectMore, &brFill);

				if (CBCGPDrawManager::IsDarkColor(clrBackground))
				{
					state = CBCGPMenuImages::ImageWhite;
				}
				else if (CBCGPVisualManager::GetInstance()->IsDarkTheme())
				{
					state = CBCGPMenuImages::ImageBlack2;
				}
			}

			CBCGPMenuImages::Draw(pDC, CBCGPMenuImages::IdArowDownLarge, m_rectMore, state);

			if (nHorzAlign == TA_RIGHT && !m_rectLink.IsRectEmpty())
			{
				m_rectLink.OffsetRect(-sizeMore.cx, 0);
			}
		}

		if (!m_strLink.IsEmpty())
		{
			if (m_fontUnderline.GetSafeHandle() == NULL)
			{
				LOGFONT lf;
				memset(&lf, 0, sizeof (LOGFONT));
				
				pFont->GetLogFont (&lf);
				lf.lfUnderline = TRUE;
				
				m_fontUnderline.CreateFontIndirect(&lf);
			}
			
			pDC->SelectObject(&m_fontUnderline);
			
			COLORREF clrLinkOld = pDC->SetTextColor(clrLink);
			
			CRect rectLinkDraw = m_rectLink;
			pDC->DrawText(m_strLink, rectLinkDraw, DT_WORDBREAK | DT_NOPREFIX);
			
			if (m_bIsFocused)
			{
				rectLinkDraw.InflateRect(1, 1);
				pDC->DrawFocusRect(rectLinkDraw);
			}
			
			pDC->SetTextColor(clrLinkOld);
		}
	}

	pDC->SetTextColor(clrTextOld);
	pDC->SetBkMode(nBkModeOld);
	pDC->SelectObject(pOldFont);

	if (rgn.GetSafeHandle() != NULL || rgnFrame.GetSafeHandle() != NULL)
	{
		pDC->SelectClipRgn(NULL);
	}

	rectFrame.right += nShadowSize;
	rectFrame.bottom += nShadowSize;

	return rectFrame;
}
//*******************************************************************************************************
CFont* CBCGPInfoBoxRenderer::SelectCaptionFont(CDC* pDC, CFont* pFont)
{
	ASSERT_VALID(pDC);
	ASSERT_VALID(pFont);

	if (m_CaptionStyle == BCGP_INFOBOX_CAPTION_BOLD)
	{
		if (m_fontBold.GetSafeHandle() == NULL)
		{
			LOGFONT lf;
			memset(&lf, 0, sizeof (LOGFONT));
			
			pFont->GetLogFont (&lf);
			lf.lfWeight = FW_BOLD;
			
			m_fontBold.CreateFontIndirect(&lf);
		}
		
		return pDC->SelectObject(&m_fontBold);
	}
#ifndef _BCGSUITE_
	else if (m_CaptionStyle == BCGP_INFOBOX_CAPTION_LARGE)
	{
		return pDC->SelectObject(&globalData.fontCaption);
	}
#endif
	return NULL;	// Use current font
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPInfoBoxCtrl

BCGCBPRODLLEXPORT UINT BCGM_INFOBOX_LINK_CLICKED = ::RegisterWindowMessage(_T("BCGM_INFOBOX_LINK_CLICKED"));

IMPLEMENT_DYNAMIC(CBCGPInfoBoxCtrl, CBCGPStatic)

CBCGPInfoBoxCtrl::CBCGPInfoBoxCtrl()
{
	m_hIcon = NULL;
	m_pToolTip = NULL;
	m_bDrawMoreTooltipMarker = TRUE;
}

BEGIN_MESSAGE_MAP(CBCGPInfoBoxCtrl, CBCGPStatic)
	ON_WM_DESTROY()
	ON_REGISTERED_MESSAGE(BCGM_UPDATETOOLTIPS, OnBCGUpdateToolTips)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnNeedTipText)
	ON_WM_SETCURSOR()
	ON_WM_LBUTTONUP()
	ON_WM_KILLFOCUS()
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CBCGPInfoBoxCtrl::SetCaption(const CString& strCaption)
{
	m_strCaption = strCaption;
}
//*******************************************************************************************************
void CBCGPInfoBoxCtrl::SetIcon(HICON hIcon)
{
	m_hIcon = hIcon;	
}
//*******************************************************************************************************
void CBCGPInfoBoxCtrl::SetLink(const CString& strLink)
{
	m_strLink = strLink;	
}
//*******************************************************************************************************
void CBCGPInfoBoxCtrl::SizeToContent(BCGP_SIZE_TO_CONTENT_LINES lines/* = BCGP_SIZE_TO_CONTENT_LINES_AUTO*/)
{
	UNREFERENCED_PARAMETER(lines);

	if (GetSafeHwnd() == NULL)
	{
		return;
	}

	CClientDC dc(this);
	CRect rect = DoDrawCtrl(&dc, TRUE);

	CRect rectWndOld;
	GetWindowRect(rectWndOld);

	SetWindowPos(NULL, -1, -1, rectWndOld.Width(), rect.Height(), SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

	CRect rectWnd;
	GetWindowRect(rectWnd);

	rectWnd.bottom = max(rectWnd.bottom, rectWndOld.bottom);

	CWnd* pParent = GetParent();

	pParent->ScreenToClient(rectWnd);

	pParent->RedrawWindow(rectWnd, NULL, RDW_ALLCHILDREN | RDW_ERASE | RDW_UPDATENOW | RDW_INVALIDATE);
}
//*******************************************************************************************************
CRect CBCGPInfoBoxCtrl::DoDrawCtrl(CDC* pDC, BOOL bCalcOnly)
{
	const DWORD dwStyle = GetStyle();
	
	m_nHorzAlign = TA_LEFT;
	
	if (dwStyle & SS_CENTER)
	{
		m_nHorzAlign = TA_CENTER;
	}
	else if (dwStyle & SS_RIGHT)
	{
		m_nHorzAlign = TA_RIGHT;
	}
	
	m_nVertAlign = TA_TOP;
	
	if ((dwStyle & SS_CENTERIMAGE) == SS_CENTERIMAGE)
	{
		m_nVertAlign = TA_CENTER;
	}

	CRect rectClient;
	GetClientRect(rectClient);
	
	CString strText;
	GetWindowText(strText);

	SetInfoBoxFont((HFONT)GetFont()->GetSafeHandle());
	
	return DoDraw(pDC, rectClient, strText, m_strCaption, m_hIcon, bCalcOnly, m_bPreview);
}
//*******************************************************************************************************
void CBCGPInfoBoxCtrl::DoPaint(CDC* pDC)
{
	COLORREF clrBackgroundSaved = m_clrBackground;
	COLORREF clrForegroundSaved = m_clrForeground;
	COLORREF clrFrameSaved = m_clrFrame;
	COLORREF clrLinkSaved = m_clrLink;

	if (m_clrBackground == CLR_NONE)
	{
		if (m_clrForeground == CLR_DEFAULT)
		{
			m_clrForeground = GetForegroundColor();
		}

		if (m_clrFrame == CLR_DEFAULT)
		{
			m_clrFrame = globalData.clrBarShadow;
		}

		if (m_clrLink == CLR_DEFAULT)
		{
#ifndef _BCGSUITE_
			m_clrLink = globalData.clrHotText;
#else
			m_clrLink = globalData.clrHotLinkNormalText;
#endif
		}
	}
	else if (m_clrBackground == CLR_DEFAULT && m_bVisualManagerStyle)
	{
		CBCGPVisualManager::GetInstance()->GetInfoBoxColors(this, m_clrBackground, m_clrForeground, m_clrFrame, m_clrLink);
	}

	DoDrawCtrl(pDC, FALSE);

	if (!m_rectLink.IsRectEmpty())
	{
		ModifyStyle(0, SS_NOTIFY);
	}

	if (!m_rectMore.IsRectEmpty() && m_pToolTip->GetSafeHwnd() == NULL)
	{
		ModifyStyle(0, SS_NOTIFY);

		CBCGPTooltipManager::CreateToolTip(m_pToolTip, this, BCGP_TOOLTIP_TYPE_DEFAULT);
		
		if (m_pToolTip->GetSafeHwnd() != NULL)
		{
			m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &m_rectMore, 1);
		}
	}

	m_clrBackground = clrBackgroundSaved;
	m_clrForeground = clrForegroundSaved;
	m_clrFrame = clrFrameSaved;
	m_clrLink = clrLinkSaved;

	UpdateTooltips();
}
//**************************************************************************
void CBCGPInfoBoxCtrl::UpdateTooltips()
{
	if (m_pToolTip->GetSafeHwnd () != NULL)
	{
		m_pToolTip->SetToolRect(this, 1, m_rectMore);
	}
}
//*********************************************************************************
void CBCGPInfoBoxCtrl::OnDestroy() 
{
	CBCGPTooltipManager::DeleteToolTip(m_pToolTip);
	CBCGPStatic::OnDestroy();
}
//**************************************************************************
LRESULT CBCGPInfoBoxCtrl::OnBCGUpdateToolTips(WPARAM wp, LPARAM)
{
	UINT nTypes = (UINT) wp;

	if ((nTypes & BCGP_TOOLTIP_TYPE_DEFAULT) && !m_rectMore.IsRectEmpty())
	{
		CBCGPTooltipManager::CreateToolTip(m_pToolTip, this, BCGP_TOOLTIP_TYPE_DEFAULT);

		m_pToolTip->AddTool(this, LPSTR_TEXTCALLBACK, &m_rectMore, 1);
	}

	return 0;
}
//**********************************************************************************
BOOL CBCGPInfoBoxCtrl::OnNeedTipText(UINT /*id*/, NMHDR* pNMH, LRESULT* /*pResult*/)
{
	static CString strTipText;
	
	if (m_pToolTip->GetSafeHwnd () == NULL || pNMH->hwndFrom != m_pToolTip->GetSafeHwnd ())
	{
		return FALSE;
	}
	
	GetWindowText(strTipText);
	
	LPNMTTDISPINFO	pTTDispInfo	= (LPNMTTDISPINFO) pNMH;
	ASSERT((pTTDispInfo->uFlags & TTF_IDISHWND) == 0);

	pTTDispInfo->lpszText = const_cast<LPTSTR> ((LPCTSTR) strTipText);
	return TRUE;
}
//**********************************************************************************
BOOL CBCGPInfoBoxCtrl::PreTranslateMessage(MSG* pMsg)
{
   	switch (pMsg->message)
	{
	case WM_KEYDOWN:
	case WM_SYSKEYDOWN:
	case WM_LBUTTONDOWN:
	case WM_RBUTTONDOWN:
	case WM_MBUTTONDOWN:
	case WM_LBUTTONUP:
	case WM_RBUTTONUP:
	case WM_MBUTTONUP:
	case WM_MOUSEMOVE:
		if (m_pToolTip->GetSafeHwnd () != NULL)
		{
			m_pToolTip->RelayEvent(pMsg);
		}
		break;
	}
	
	if (pMsg->message == WM_KEYDOWN && m_bIsFocused && !m_rectLink.IsRectEmpty())
	{
		if ((pMsg->wParam == VK_RETURN || pMsg->wParam == VK_SPACE) && GetOwner()->GetSafeHwnd() != NULL)
		{
			GetOwner()->SendMessage(BCGM_INFOBOX_LINK_CLICKED, GetDlgCtrlID());
			return TRUE;
		}
	}

	return CBCGPStatic::PreTranslateMessage(pMsg);
}
//**********************************************************************************
BOOL CBCGPInfoBoxCtrl::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message)
{
	CPoint point;
	
	::GetCursorPos (&point);
	ScreenToClient (&point);

	if (m_rectLink.PtInRect(point))
	{
		::SetCursor (globalData.GetHandCursor ());
		return TRUE;
	}
	
	return CWnd::OnSetCursor (pWnd, nHitTest, message);
}
//**********************************************************************************
void CBCGPInfoBoxCtrl::OnLButtonUp(UINT nFlags, CPoint point)
{
	if (m_rectLink.PtInRect(point) && GetOwner()->GetSafeHwnd() != NULL)
	{
		GetOwner()->SendMessage(BCGM_INFOBOX_LINK_CLICKED, GetDlgCtrlID());
		return;
	}

	CBCGPStatic::OnLButtonUp(nFlags, point);
}
//**********************************************************************************
void CBCGPInfoBoxCtrl::OnKillFocus(CWnd* pNewWnd)
{
	CBCGPStatic::OnKillFocus(pNewWnd);

	m_bIsFocused = FALSE;
	RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
}
//**********************************************************************************
void CBCGPInfoBoxCtrl::OnSetFocus(CWnd* pOldWnd) 
{
	CBCGPStatic::OnSetFocus(pOldWnd);
	
	m_bIsFocused = TRUE;
	RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
}
//*******************************************************************************************************
BOOL CBCGPInfoBoxCtrl::CopyToClipboard(BOOL bWithCaption/* = TRUE*/, BOOL bWithLink/* = FALSE*/)
{
#ifdef _UNICODE
	#define _TCF_TEXT	CF_UNICODETEXT
#else
	#define _TCF_TEXT	CF_TEXT
#endif

	BOOL bRes = FALSE;

	if (OpenClipboard())
	{
		EmptyClipboard();
		
		CString strText;
		GetWindowText(strText);

		if (bWithCaption && !m_strCaption.IsEmpty())
		{
			strText = m_strCaption + _T("\r\n\r\n") + strText;
		}

		if (bWithLink && !m_strLink.IsEmpty())
		{
			strText += _T("\r\n\r\n") + m_strLink;
		}

		HGLOBAL hClipbuffer = ::GlobalAlloc(GMEM_DDESHARE, (strText.GetLength() + 1) * sizeof(TCHAR));
		if (hClipbuffer != NULL)
		{
			LPTSTR lpszBuffer = (LPTSTR)GlobalLock(hClipbuffer);
			
			lstrcpy(lpszBuffer, (LPCTSTR)strText);
			
			::GlobalUnlock(hClipbuffer);
			::SetClipboardData(_TCF_TEXT, hClipbuffer);

			bRes = TRUE;
		}
		
		CloseClipboard();
	}

	return bRes;
}
