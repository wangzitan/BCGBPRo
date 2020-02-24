// BCGPCaptionBar.cpp: implementation of the CBCGPCaptionBar class.
//
// This is a part of BCGControlBar Library Professional Edition
// Copyright (C) 1998-2018 BCGSoft Ltd.
// All rights reserved.
//
// This source code can be used, distributed or modified
// only under terms and conditions 
// of the accompanying license agreement.
 //////////////////////////////////////////////////////////////////////

#include "stdafx.h"

#include "BCGPCaptionBar.h"
#include "BCGPVisualManager.h"
#include "BCGPToolBar.h"
#include "trackmouse.h"
#include "BCGPFrameWnd.h"
#include "BCGPMDIFrameWnd.h"
#include "BCGPOleIPFrameWnd.h"
#include "BCGPOleDocIPFrameWnd.h"
#include "BCGPMDIChildWnd.h"
#include "BCGPGlobalUtils.h"
#include "BCGPTooltipManager.h"
#include "BCGPToolTipCtrl.h"
#include "BCGPLocalResource.h"
#include "BCGProRes.h"

#define MENU_ARROW_WIDTH	globalUtils.ScaleByDPI(10)
#define BUTTON_HORZ_MARGIN	globalUtils.ScaleByDPI(10)
#define CLOSE_BUTTON_MARGIN	globalUtils.ScaleByDPI(4)

const int nIdToolTipClose = 1;
const int nIdToolTipText = 2;
const int nIdToolTipImage = 3;
const int nIdToolTipButton = 4;

UINT BCGM_ON_CLICK_CAPTIONBAR_HYPERLINK = ::RegisterWindowMessage (_T("BCGM_ON_CLICK_CAPTIONBAR_HYPERLINK"));

IMPLEMENT_DYNCREATE(CBCGPCaptionBar, CBCGPControlBar)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPCaptionBar::CBCGPCaptionBar() 
{
	m_pToolTip				= NULL;

	m_clrBarText			= (COLORREF)-1;
	m_clrBarBackground		= (COLORREF)-1;
	m_clrBarBorder			= (COLORREF)-1;

	m_nBorderSize			= globalUtils.ScaleByDPI(4);
	m_nMargin				= globalUtils.ScaleByDPI(4);
	m_nHorzElementOffset	= globalUtils.ScaleByDPI(4);

	m_hIcon					= NULL;
	m_hFont					= NULL;
	m_bUseCaptionFont		= FALSE;

	m_nDefaultHeight		= -1;
	m_nCurrentHeight		= 0;

	m_btnAlignnment			= ALIGN_LEFT;
	m_iconAlignment			= ALIGN_LEFT;
	m_textAlignment			= ALIGN_LEFT;

	m_bStretchImage			= FALSE;

	m_bFlatBorder			= FALSE;
	m_uiBtnID				= 0;

	m_bIsBtnPressed			= FALSE;
	m_bIsBtnForcePressed	= FALSE;
	m_bIsBtnHighlighted		= FALSE;

	m_bTracked				= FALSE;
	m_bBtnEnabled			= TRUE;
	m_bBtnHasDropDownArrow	= TRUE;

	m_rectImage.SetRectEmpty ();
	m_rectText.SetRectEmpty ();
	m_rectDrawText.SetRectEmpty ();
	m_rectButton.SetRectEmpty ();
	m_bTextIsTruncated		= FALSE;

	m_bIsMessageBarMode = FALSE;

	m_bIsCloseBtnPressed	= FALSE;
	m_bIsCloseBtnHighlighted= FALSE;
	m_bCloseTracked			= FALSE;
	
	m_rectClose.SetRectEmpty ();

	m_nHyperlinkHighlighted	= -1;
	m_nHyperlinkPressed		= -1;
}
//*****************************************************************************
CBCGPCaptionBar::~CBCGPCaptionBar()
{
}
//*****************************************************************************
BOOL CBCGPCaptionBar::Create (DWORD dwStyle, CWnd* pParentWnd, UINT uID, int nHeight,
							  BOOL bIsMessageBarMode)
{
	SetBarStyle (CBRS_ALIGN_TOP);
	m_nDefaultHeight = nHeight;

	if (bIsMessageBarMode)
	{
		m_dwStyle |= CBRS_HIDE_INPLACE;
	}
	
	if (!CBCGPControlBar::Create (NULL, dwStyle, CRect (0, 0, 0, 0),
		pParentWnd, uID, 0))
	{
		return FALSE;
	}

	// register with parent frames' dock manager!!!
	if (pParentWnd->IsKindOf (RUNTIME_CLASS (CBCGPFrameWnd)))
	{
		((CBCGPFrameWnd*) pParentWnd)->AddControlBar (this);
	}
	else if (pParentWnd->IsKindOf (RUNTIME_CLASS (CBCGPMDIFrameWnd)))
	{
		((CBCGPMDIFrameWnd*) pParentWnd)->AddControlBar (this);
	}
	else if (pParentWnd->IsKindOf (RUNTIME_CLASS (CBCGPOleIPFrameWnd)))
	{
		((CBCGPOleIPFrameWnd*) pParentWnd)->AddControlBar (this);
	}
	else if (pParentWnd->IsKindOf (RUNTIME_CLASS (CBCGPOleDocIPFrameWnd)))
	{
		((CBCGPOleIPFrameWnd*) pParentWnd)->AddControlBar (this);
	}
	else if (pParentWnd->IsKindOf (RUNTIME_CLASS (CBCGPMDIChildWnd)))
	{
		((CBCGPMDIChildWnd*) pParentWnd)->AddControlBar (this);
	}
	else if (pParentWnd->IsKindOf (RUNTIME_CLASS (CDialog)))
	{
		if (pParentWnd->GetSafeHwnd() == AfxGetMainWnd()->GetSafeHwnd())
		{
			globalUtils.m_bDialogApp = TRUE;
		}
	}
	else
	{
		ASSERT (FALSE);
	}
	
	m_bIsMessageBarMode = bIsMessageBarMode;
	return TRUE;
}

BEGIN_MESSAGE_MAP(CBCGPCaptionBar, CBCGPControlBar)
	//{{AFX_MSG_MAP(CBCGPCaptionBar)
	ON_WM_CREATE()
	ON_WM_ERASEBKGND()
	ON_WM_SIZE()
	ON_WM_SYSCOLORCHANGE()
	ON_WM_NCCALCSIZE()
	ON_WM_NCPAINT()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_MOUSEMOVE()
	ON_WM_RBUTTONUP()
	ON_WM_DESTROY()
	ON_WM_SETCURSOR()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_SETFONT, OnSetFont)
	ON_MESSAGE(WM_GETFONT, OnGetFont)
	ON_MESSAGE(WM_MOUSELEAVE, OnMouseLeave)
	ON_REGISTERED_MESSAGE(BCGM_UPDATETOOLTIPS, OnBCGUpdateToolTips)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXT, 0, 0xFFFF, OnNeedTipText)
END_MESSAGE_MAP()

void CBCGPCaptionBar::OnUpdateCmdUI(CFrameWnd* /*pTarget*/, BOOL /*bDisableIfNoHndler*/)
{
}
//*****************************************************************************
int CBCGPCaptionBar::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CBCGPControlBar::OnCreate (lpCreateStruct) == -1)
		return -1;

	CBCGPTooltipManager::CreateToolTip (m_pToolTip, this,
		BCGP_TOOLTIP_TYPE_CAPTIONBAR);

	if (m_pToolTip->GetSafeHwnd () != NULL)
	{
		CRect rectDummy (0, 0, 0, 0);

		m_pToolTip->SetMaxTipWidth (640);

		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectDummy, nIdToolTipClose);
		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectDummy, nIdToolTipText);
		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectDummy, nIdToolTipImage);
		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectDummy, nIdToolTipButton);
	}

	SetWindowText (_T("Caption Bar"));
	return 0;
}
//*****************************************************************************
void CBCGPCaptionBar::OnSize(UINT nType, int cx, int cy)
{
	CBCGPControlBar::OnSize(nType, cx, cy);
	RecalcLayout ();
	InvalidateRect (NULL);
}
//*****************************************************************************
void CBCGPCaptionBar::OnNcCalcSize(BOOL /*bCalcValidRects*/, NCCALCSIZE_PARAMS FAR* lpncsp)
{
	if (!m_bIsMessageBarMode)
	{
		lpncsp->rgrc[0].bottom -= m_nBorderSize;
		lpncsp->rgrc[0].top  += m_nBorderSize;
	}
}
//*****************************************************************************
void CBCGPCaptionBar::DoPaint(CDC* pDCIn)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDCIn);

	CBCGPMemDC memDC(*pDCIn, this);
	CDC* pDC = &memDC.GetDC();

	CRect rectClient;
	GetClientRect(rectClient);

	OnDrawBackground (pDC, rectClient);

	int nOldBkMode = pDC->SetBkMode (TRANSPARENT);
	COLORREF clrOldText = pDC->SetTextColor(m_clrBarText != (COLORREF) -1 ? m_clrBarText : CBCGPVisualManager::GetInstance ()->GetCaptionBarTextColor (this));

	CFont* pOldFont = SelectFont(pDC);

	CRect rectButton = m_rectButton;
	rectButton.top = max(rectClient.top + 1, rectButton.top);
	rectButton.bottom = min(rectClient.bottom - 1, rectButton.bottom);

	OnDrawButton (pDC, rectButton, m_strBtnText, m_bBtnEnabled);
	OnDrawText (pDC, m_rectDrawText, m_strText);

	if (!m_rectImage.IsRectEmpty ())
	{
		OnDrawImage (pDC, m_rectImage);
	}

	if (!m_rectClose.IsRectEmpty ())
	{
		COLORREF clrText = CBCGPVisualManager::GetInstance ()->OnFillCaptionBarButton (pDC, this,
			m_rectClose, m_bIsCloseBtnPressed, m_bIsCloseBtnHighlighted, FALSE,
			FALSE, TRUE);

		if (m_bIsMessageBarMode && !CBCGPVisualManager::GetInstance()->IsCaptionBarCloseButtonOutsideMessageBar() && !m_bIsCloseBtnPressed && !m_bIsCloseBtnHighlighted)
		{
			clrText = pDC->GetTextColor();
		}

		CBCGPVisualManager::GetInstance ()->OnDrawCaptionBarCloseButton(pDC, this, 
			m_rectClose, m_bIsCloseBtnPressed, m_bIsCloseBtnHighlighted, clrText);

		CBCGPVisualManager::GetInstance ()->OnDrawCaptionBarButtonBorder (pDC, this,
			m_rectClose, m_bIsCloseBtnPressed, m_bIsCloseBtnHighlighted, FALSE,
			FALSE, TRUE);
	}

	pDC->SelectObject (pOldFont);
	pDC->SetTextColor (clrOldText);
	pDC->SetBkMode (nOldBkMode);
}
//*****************************************************************************
void CBCGPCaptionBar::OnNcPaint() 
{
	CWindowDC dcWin(this);
	DoNcPaint(&dcWin);
}
//*****************************************************************************
void CBCGPCaptionBar::DoNcPaint(CDC* pDC)
{
	ASSERT_VALID(pDC);

	CRect rectClient;
	GetClientRect(rectClient);

	CRect rectWindow;
	GetWindowRect(rectWindow);

	CRect rectBorder = rectWindow;

	ScreenToClient(rectWindow);

	rectClient.OffsetRect(-rectWindow.left, -rectWindow.top);
	pDC->ExcludeClipRect (rectClient);
	
	rectBorder.OffsetRect(-rectBorder.left, -rectBorder.top);

	int nBorderSize = m_bIsMessageBarMode ? 0 : m_nBorderSize;

	int nTop = rectBorder.top;
	rectBorder.top = rectBorder.bottom - nBorderSize;
	OnDrawBorder(pDC, rectBorder);

	rectBorder.top = nTop;
	rectBorder.bottom = rectBorder.top + nBorderSize;

	OnDrawBorder(pDC, rectBorder);
	
	pDC->SelectClipRgn (NULL);
}
//*****************************************************************************
void CBCGPCaptionBar::OnDrawBackground (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	CBCGPVisualManager::GetInstance ()->OnFillBarBackground (pDC, this,
		rect, rect);

	if (m_bIsMessageBarMode)
	{
		int nMessageBarMargin = CBCGPVisualManager::GetInstance()->GetMessageBarMargin();

		CRect rectInfo = rect;
		rectInfo.DeflateRect (nMessageBarMargin, nMessageBarMargin);

		if (CBCGPVisualManager::GetInstance()->IsCaptionBarCloseButtonOutsideMessageBar())
		{
			rectInfo.right = m_rectClose.left - CLOSE_BUTTON_MARGIN;
		}

		CBCGPVisualManager::GetInstance ()->OnDrawCaptionBarInfoArea (pDC, this, rectInfo);
	}
}
//*****************************************************************************
void CBCGPCaptionBar::OnDrawBorder (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	rect.InflateRect (2, 0);

	CBCGPVisualManager::GetInstance ()->OnDrawCaptionBarBorder (
		pDC, this, rect, m_clrBarBorder, m_bFlatBorder);
}
//*****************************************************************************
void CBCGPCaptionBar::OnDrawButton (CDC* pDC, CRect rect, 
								   const CString& strButton, BOOL bEnabled)
{
	ASSERT_VALID (pDC);

	COLORREF clrText = CBCGPVisualManager::GetInstance ()->OnFillCaptionBarButton (pDC, this,
		rect, m_bIsBtnPressed || m_bIsBtnForcePressed, m_bIsBtnHighlighted, !m_bBtnEnabled,
		m_bBtnHasDropDownArrow, FALSE);

	CRect rectText = rect;
	rectText.DeflateRect (m_nHorzElementOffset, 0);

	if (m_bIsMessageBarMode)
	{
		rectText.DeflateRect (BUTTON_HORZ_MARGIN, 0);
	}

	if (m_uiBtnID != 0 && bEnabled && m_bBtnHasDropDownArrow)
	{
		rectText.right -= MENU_ARROW_WIDTH;
	}

	COLORREF clrTextOld = (COLORREF)-1;
	if (clrText != (COLORREF)-1)
	{
		clrTextOld = pDC->SetTextColor (clrText);
	}

	pDC->DrawText (strButton, rectText,
					DT_END_ELLIPSIS | DT_SINGLELINE | DT_VCENTER);

	if (clrTextOld != (COLORREF)-1)
	{
		pDC->SetTextColor (clrTextOld);
	}

	if (m_uiBtnID != 0 && bEnabled)
	{
		if (m_uiBtnID != 0 && m_bBtnHasDropDownArrow)
		{
			// Draw menu triangle:
			CRect rectArrow = rect;
			rectArrow.top = rectArrow.CenterPoint().y - MENU_ARROW_WIDTH / 2;
			rectArrow.bottom = rectArrow.top + MENU_ARROW_WIDTH;
			rectArrow.left = rectText.right;

			int iXMiddle = rectArrow.left + rectArrow.Width () / 2;

			rectArrow.DeflateRect (0, rectArrow.Height () / 3);
			rectArrow.DeflateRect (rectArrow.Height () / 3, rectArrow.Height () / 3);
			rectArrow.left = iXMiddle - rectArrow.Height () - 1;
			rectArrow.right = iXMiddle + rectArrow.Height () + 1;

			int iHalfWidth =	(rectArrow.Width () % 2 != 0) ?
								(rectArrow.Width () - 1) / 2 :
								rectArrow.Width () / 2;

			CPoint pts [3];
			pts[0].x = rectArrow.left;
			pts[0].y = rectArrow.top;
			pts[1].x = rectArrow.right;
			pts[1].y = rectArrow.top;
			pts[2].x = rectArrow.left + iHalfWidth;
			pts[2].y = rectArrow.bottom + 1;

			CBrush brArrow (globalData.IsHighContastMode() ? globalData.clrWindowText : (clrText != (COLORREF)-1 ? clrText : pDC->GetTextColor()));

			CPen* pOldPen = (CPen*) pDC->SelectStockObject (NULL_PEN);
			CBrush* pOldBrush = (CBrush*) pDC->SelectObject(&brArrow);

			pDC->SetPolyFillMode (WINDING);
			pDC->Polygon (pts, 3);

			pDC->SelectObject (pOldBrush);
			pDC->SelectObject (pOldPen);
		}

		CBCGPVisualManager::GetInstance ()->OnDrawCaptionBarButtonBorder (pDC, this,
			rect, m_bIsBtnPressed || m_bIsBtnForcePressed, m_bIsBtnHighlighted, !m_bBtnEnabled,
			m_bBtnHasDropDownArrow, FALSE);
	}
}	
//*****************************************************************************
void CBCGPCaptionBar::OnDrawText (CDC* pDC, CRect rect, const CString& strText)
{
	ASSERT_VALID (pDC);

	m_arHyperlinksRects.RemoveAll();

	if (m_arTextParts.GetSize () == 1)
	{
		pDC->DrawText (strText, rect, 
						DT_END_ELLIPSIS | DT_SINGLELINE | DT_VCENTER);
		return;
	}

	int x = rect.left;
	int nHyperLink = 0;

	for (int i = 0; i < m_arTextParts.GetSize (); i++)
	{
		if (!m_arTextParts [i].IsEmpty ())
		{
			CFont* pOldFont = NULL;

			if (m_arTextPartTypes[i] == 1 && nHyperLink != m_nHyperlinkHighlighted)
			{
				pOldFont = pDC->SelectObject (&globalData.fontUnderline);
			}
			else if (m_arTextPartTypes[i] == 2)
			{
				pOldFont = pDC->SelectObject (&globalData.fontBold);
			}

			CRect rectPart = rect;
			rectPart.left = x;

			CSize sizePart = pDC->GetTextExtent (m_arTextParts [i]);

			pDC->DrawText (m_arTextParts [i], rectPart, 
							DT_END_ELLIPSIS | DT_SINGLELINE | DT_VCENTER);

			if (m_arTextPartTypes[i] == 1)
			{
				rectPart.right = rectPart.left + pDC->GetTextExtent (m_arTextParts [i]).cx;
				m_arHyperlinksRects.Add(rectPart);
				nHyperLink++;
			}

			if (pOldFont != NULL)
			{
				pDC->SelectObject (pOldFont);
			}

			x += sizePart.cx;
		}
	}
}	
//*****************************************************************************
void CBCGPCaptionBar::OnDrawImage (CDC* pDC, CRect rect)
{
	ASSERT_VALID (pDC);

	if (m_hIcon != NULL)
	{
		DrawIconEx (pDC->GetSafeHdc(), rect.left, rect.top, m_hIcon, 
					rect.Width (), rect.Height (), NULL, (HBRUSH)NULL, 
					DI_NORMAL) ;
	}
	else if (m_Bitmap.GetCount () > 0)
	{
		CSize sizeDest;
		if (m_bStretchImage)
		{
			sizeDest = rect.Size ();
		}
		else
		{
			sizeDest = m_rectImage.Size ();
		}

		CBCGPDrawState ds;
		m_Bitmap.PrepareDrawImage (ds, sizeDest);
		m_Bitmap.Draw (pDC, rect.left, rect.top, 0);
		m_Bitmap.EndDrawImage (ds);
	}
}
//*****************************************************************************
void CBCGPCaptionBar::OnSysColorChange() 
{
}
//*****************************************************************************
CSize CBCGPCaptionBar::CalcFixedLayout(BOOL /*bStretch*/, BOOL /*bHorz*/)
{
	RecalcLayout ();
	return CSize (32767, m_nCurrentHeight);
}
//*****************************************************************************
void CBCGPCaptionBar::SetButton (LPCTSTR lpszLabel, UINT uiCmdUI, BarElementAlignment btnAlignmnet,
								 BOOL bHasDropDownArrow)
{
	ASSERT (lpszLabel != NULL);

	m_strBtnText = lpszLabel;
	m_uiBtnID = uiCmdUI;
	m_btnAlignnment = btnAlignmnet;
	m_bBtnHasDropDownArrow = bHasDropDownArrow;

	AdjustLayout ();
}
//*****************************************************************************
void CBCGPCaptionBar::EnableButton (BOOL bEnable)
{
	m_bBtnEnabled = bEnable;

	if (GetSafeHwnd () != NULL)
	{
		CRect rectButton = m_rectButton;

		RecalcLayout ();

		InvalidateRect (rectButton);
		InvalidateRect (m_rectButton);

		UpdateWindow ();
	}
}
//*****************************************************************************
void CBCGPCaptionBar::SetButtonPressed (BOOL bPresed)
{
	m_bIsBtnForcePressed = bPresed;

	if (GetSafeHwnd () != NULL)
	{
		InvalidateRect (m_rectButton);
		UpdateWindow ();
	}
}
//*****************************************************************************
void CBCGPCaptionBar::SetButtonToolTip (LPCTSTR lpszToolTip, LPCTSTR lpszDescription)
{
	m_strButtonToolTip = lpszToolTip == NULL ? _T("") : lpszToolTip;
	m_strButtonDescription = lpszDescription == NULL ? _T("") : lpszDescription;

	UpdateTooltips ();
}
//*****************************************************************************
void CBCGPCaptionBar::RemoveButton ()
{
	m_strBtnText.Empty ();
	AdjustLayout ();
}
//*****************************************************************************
void CBCGPCaptionBar::SetIcon (HICON hIcon, BarElementAlignment iconAlignment)
{
	m_Bitmap.Clear ();

	m_hIcon = hIcon;
	m_iconAlignment = iconAlignment;

	AdjustLayout ();
}
//*****************************************************************************
void CBCGPCaptionBar::RemoveIcon ()
{
	m_hIcon = NULL;
	AdjustLayout ();
}
//*****************************************************************************
void CBCGPCaptionBar::SetBitmap (HBITMAP hBitmap, COLORREF clrTransparent, 
								BOOL bStretch, BarElementAlignment bmpAlignment,
								BOOL bAutoScale)
{
	ASSERT (hBitmap != NULL);

	m_hIcon = NULL;
	m_Bitmap.Clear ();

	BITMAP bmp;
	::GetObject (hBitmap, sizeof (BITMAP), (LPVOID) &bmp);

	m_Bitmap.SetImageSize (CSize (bmp.bmWidth, bmp.bmHeight));
	m_Bitmap.SetTransparentColor (clrTransparent);
	m_Bitmap.AddImage (hBitmap, clrTransparent == (COLORREF)-1);

	if (bAutoScale)
	{
		globalUtils.ScaleByDPI(m_Bitmap);
	}
	
	m_bStretchImage = bStretch;

	m_iconAlignment = bmpAlignment;

	AdjustLayout ();
}
//*****************************************************************************
void CBCGPCaptionBar::SetBitmap (UINT uiBmpResID, COLORREF clrTransparent, 
								BOOL bStretch, BarElementAlignment bmpAlignment,
								BOOL bAutoScale)
{
	m_hIcon = NULL;
	m_Bitmap.Clear ();

	m_Bitmap.SetTransparentColor (clrTransparent);
	m_Bitmap.Load (uiBmpResID);
	m_Bitmap.SetSingleImage (FALSE);
	
	if (bAutoScale)
	{
		globalUtils.ScaleByDPI(m_Bitmap);
	}

	m_bStretchImage = bStretch;

	m_iconAlignment = bmpAlignment;

	AdjustLayout ();
}
//*****************************************************************************
void CBCGPCaptionBar::RemoveBitmap ()
{
	m_Bitmap.Clear ();

	AdjustLayout ();
}
//*****************************************************************************
void CBCGPCaptionBar::SetImageToolTip (LPCTSTR lpszToolTip, LPCTSTR lpszDescription)
{
	m_strImageToolTip = lpszToolTip == NULL ? _T("") : lpszToolTip;
	m_strImageDescription = lpszDescription == NULL ? _T("") : lpszDescription;

	UpdateTooltips ();
}
//*****************************************************************************
void CBCGPCaptionBar::SetText (const CString& strText, BarElementAlignment textAlignment)
{
	BOOL bWasEmptyText = m_strText.IsEmpty ();

	m_arTextParts.RemoveAll ();
	m_arTextPartTypes.RemoveAll();

	int iStart = 0;
	int nType = 0;

	for (int i = 0; i < strText.GetLength (); i++)
	{
		if (strText [i] == _T('\a') || strText [i] == _T('\b'))
		{
			m_arTextParts.Add (strText.Mid (iStart, i - iStart));
			m_arTextPartTypes.Add(nType);

			nType = (nType != 0) ? 0 : (strText [i] == _T('\a') ? 1 : 2);

			iStart = i + 1;
		}
	}

	m_arTextParts.Add (strText.Mid (iStart));
	m_arTextPartTypes.Add(nType);

	m_strText = strText;
	m_textAlignment = textAlignment;

	if (m_nCurrentHeight == 0 || m_strText.IsEmpty () || bWasEmptyText)
	{
		AdjustLayout ();
	}
	else
	{
		RecalcLayout ();
		RedrawWindow ();
	}
}
//*****************************************************************************
void CBCGPCaptionBar::RemoveText ()
{
	m_strText.Empty ();
	AdjustLayout ();
}
//*****************************************************************************
void CBCGPCaptionBar::SetCaptionFont(BOOL bSet)
{
	m_bUseCaptionFont = bSet;
	AdjustLayout();
}
//*****************************************************************************
LRESULT CBCGPCaptionBar::OnSetFont (WPARAM wParam, LPARAM /*lParam*/)
{
	m_hFont = (HFONT) wParam;

	AdjustLayout ();
	return 0;
}
//*****************************************************************************
LRESULT CBCGPCaptionBar::OnGetFont (WPARAM, LPARAM)
{
	return (LRESULT) m_hFont;
}
//*****************************************************************************
CBCGPCaptionBar::BarElementAlignment CBCGPCaptionBar::GetAlignment (BarElement elem)
{
	switch (elem)
	{
	case ELEM_BUTTON:
		return m_btnAlignnment;

	case ELEM_TEXT:
		return m_textAlignment;

	case ELEM_ICON:
		return m_iconAlignment;
	}

	ASSERT (FALSE);
	return ALIGN_INVALID;
}
//*****************************************************************************
void CBCGPCaptionBar::RecalcLayout ()
{
	CClientDC dc (NULL);

	CFont* pOldFont = SelectFont(&dc);
	ASSERT (pOldFont != NULL);

	TEXTMETRIC tm;
	dc.GetTextMetrics (&tm);

	int nTextHeight = tm.tmHeight + globalUtils.ScaleByDPI(2);
	CSize sizeImage = GetImageSize ();

	const int nButtonVertMargin = CBCGPVisualManager::GetInstance()->GetMessageBarMargin();
	const int nBorderSize = m_bIsMessageBarMode ? 0 : m_nBorderSize;

	//-------------------------------------------------------------------
	// the height is set to the default (provided by the user in Create)
	// or calculated if it is -1
	//-------------------------------------------------------------------
	if (m_nDefaultHeight != -1)
	{
		m_nCurrentHeight = m_nDefaultHeight;
	}
	else
	{
		if (!m_strBtnText.IsEmpty () && m_bIsMessageBarMode)
		{
			nTextHeight += 2 * nButtonVertMargin;
		}

		m_nCurrentHeight = max (nTextHeight, sizeImage.cy) + 
			m_nMargin * 2 + nBorderSize;
	}

	if (m_bIsMessageBarMode)
	{
		m_nCurrentHeight += CBCGPVisualManager::GetInstance()->GetMessageBarMargin() * 2;
	}

	// for left and center alignment: icon, button, text
	// for right alignment: text, button, icon

	CRect rectClient;
	GetClientRect (rectClient);

	if (rectClient.IsRectEmpty ())
	{
		dc.SelectObject (pOldFont);
		return;
	}

	if (m_bIsMessageBarMode)
	{
		CSize sizeImageMsg = CBCGPMenuImages::Size ();
		sizeImageMsg.cx += 2 * CLOSE_BUTTON_MARGIN;
		sizeImageMsg.cy += 2 * CLOSE_BUTTON_MARGIN;

		const int nHorzOffset = CBCGPVisualManager::GetInstance()->IsCaptionBarCloseButtonOutsideMessageBar() ? 0 : CLOSE_BUTTON_MARGIN;
		const int nVertOffset = CBCGPVisualManager::GetInstance()->IsCaptionBarCloseButtonOutsideMessageBar() ? CLOSE_BUTTON_MARGIN : max(0, (rectClient.Height() - sizeImageMsg.cy) / 2);

		m_rectClose = CRect (
			CPoint (rectClient.right - sizeImageMsg.cx - nHorzOffset - CLOSE_BUTTON_MARGIN, rectClient.top + nVertOffset),
			sizeImageMsg);

		rectClient.right = m_rectClose.left;
		rectClient.DeflateRect (CLOSE_BUTTON_MARGIN, CLOSE_BUTTON_MARGIN);
	}
	else
	{
		m_rectClose.SetRectEmpty();
	}

	BOOL bButtonLeftOfIcon = FALSE;
	BOOL bTextLeftOfButton = FALSE;
	BOOL bTextLeftOfIcon = FALSE;

	BOOL bIconCenter = FALSE;
	BOOL bButtonCenter = FALSE;
	BOOL bButtonAfterText = FALSE;
	BOOL bTextCenter = FALSE;

	// differs from the current height, because the border size is non-client area
	int nBaseLine = rectClient.CenterPoint ().y;
	int nCenterOffset = rectClient.CenterPoint ().x;

	int nNextXOffsetLeft  = rectClient.left + m_nMargin;
	int nNextXOffsetRight = rectClient.right - m_nMargin;
	int nNextXOffsetCenter = nCenterOffset;

	if (IsImageSet ())
	{
		if (sizeImage.cy < rectClient.Height () || m_bIsMessageBarMode)
		{
			// center the icon if its height lesser than client area height
			m_rectImage.top = nBaseLine - sizeImage.cy / 2;
		}
		else
		{
			// otherwise, clip it from the bottom
			m_rectImage.top = rectClient.top + m_nMargin;
		}

		if (!m_bStretchImage)
		{
			m_rectImage.bottom = m_rectImage.top + sizeImage.cy;
		}
		else
		{
			m_rectImage.bottom = rectClient.bottom - m_nMargin;
		}

		switch (m_iconAlignment)
		{
		case ALIGN_LEFT:
			m_rectImage.left = nNextXOffsetLeft;
			m_rectImage.right = m_rectImage.left + sizeImage.cx;
			nNextXOffsetLeft = m_rectImage.right + m_nHorzElementOffset;
			break;

		case ALIGN_RIGHT:
			m_rectImage.left = nNextXOffsetRight - sizeImage.cx;
			m_rectImage.right = m_rectImage.left + sizeImage.cx;
			nNextXOffsetRight = m_rectImage.left - m_nHorzElementOffset;
			// only in this case button and text is at the left side of the icon
			bButtonLeftOfIcon = TRUE; 
			bTextLeftOfIcon = TRUE;
			break;

		case ALIGN_CENTER:
			bIconCenter = TRUE;
			nNextXOffsetCenter -= sizeImage.cx / 2;

			if (m_btnAlignnment == ALIGN_LEFT)
			{
				bButtonLeftOfIcon = TRUE;
			}

			if (m_textAlignment == ALIGN_LEFT)
			{
				bTextLeftOfIcon = TRUE;
			}
			break;

		default:
			ASSERT (FALSE);
		}
	}
	else
	{
		m_rectImage.SetRectEmpty();
	}

	int nButtonWidth = 0;

	if (!m_strBtnText.IsEmpty ())
	{
		nButtonWidth = dc.GetTextExtent (m_strBtnText).cx + 2 * m_nHorzElementOffset;

		if (m_bIsMessageBarMode)
		{
			nButtonWidth += 2 * BUTTON_HORZ_MARGIN;
		}

		if (m_uiBtnID != 0 && m_bBtnEnabled && m_bBtnHasDropDownArrow)
		{
			nButtonWidth += MENU_ARROW_WIDTH;
		}

		// the button always has a height equivalent to the bar's height
		m_rectButton.top = rectClient.top;
		m_rectButton.bottom = rectClient.bottom;

		if (m_bIsMessageBarMode)
		{
			m_rectButton.DeflateRect (0, nButtonVertMargin);
		}

		switch (m_btnAlignnment)
		{
		case ALIGN_LEFT:
			if (!m_bIsMessageBarMode || m_textAlignment != ALIGN_LEFT)
			{
				m_rectButton.left = nNextXOffsetLeft;

				if (nNextXOffsetLeft == rectClient.left + m_nMargin)
				{
					m_rectButton.left = rectClient.left + m_nMargin;
				}

				m_rectButton.right = m_rectButton.left + nButtonWidth;
				nNextXOffsetLeft = m_rectButton.right + m_nHorzElementOffset;
			}
			else
			{
				bButtonAfterText = TRUE;
			}
			break;

		case ALIGN_RIGHT:
			m_rectButton.left = nNextXOffsetRight - nButtonWidth;

			if (nNextXOffsetRight == rectClient.right - m_nMargin)
			{
				m_rectButton.left = rectClient.right - nButtonWidth - m_nMargin;
			}

			m_rectButton.right = m_rectButton.left + nButtonWidth;
			nNextXOffsetRight = m_rectButton.left - m_nHorzElementOffset;
			// only in this case text at the left side of the button
			bTextLeftOfButton = TRUE;
			break;

		case ALIGN_CENTER:
			bButtonCenter = TRUE;
			nNextXOffsetCenter -= nButtonWidth / 2;

			if (m_textAlignment == ALIGN_LEFT)
			{
				bTextLeftOfButton = TRUE;
			}
			break;

		default:
			ASSERT (FALSE);
			return;
		}
	}
	else
	{
		m_rectButton.SetRectEmpty();
	}

	CSize sizeText (0, 0);

	if (!m_strText.IsEmpty ())
	{
		sizeText = GetTextSize (&dc, m_strText);

		m_rectText.top = nBaseLine - sizeText.cy / 2;
		m_rectText.bottom = m_rectText.top + sizeText.cy;

		switch (m_textAlignment)
		{
		case ALIGN_LEFT:
			m_rectText.left = nNextXOffsetLeft;
			nNextXOffsetLeft += sizeText.cx + 2 * m_nMargin;
			break;

		case ALIGN_RIGHT:
			m_rectText.left = nNextXOffsetRight - sizeText.cx;
			break;

		case ALIGN_CENTER:
			bTextCenter = TRUE;
			nNextXOffsetCenter -= sizeText.cx / 2;
			break;

		default:
			ASSERT (FALSE);
			return;
		}

		m_rectText.right = m_rectText.left + sizeText.cx;
		AdjustRectToMargin (m_rectText, rectClient, m_nMargin);
		m_rectDrawText = m_rectText;
	}

	if (bIconCenter)
	{
		m_rectImage.left = nNextXOffsetCenter;
		m_rectImage.right = m_rectImage.left + sizeImage.cx;
		nNextXOffsetCenter = m_rectImage.right + m_nHorzElementOffset;
	}

	if (bButtonAfterText)
	{
		m_rectButton.left = nNextXOffsetLeft;
		m_rectButton.right = m_rectButton.left + nButtonWidth;

		if (m_rectButton.right + m_nMargin > rectClient.right)
		{
			m_rectButton.right = rectClient.right - m_nMargin;
			m_rectButton.left = m_rectButton.right - nButtonWidth;
		}
	}
	else if (bButtonCenter)
	{
		m_rectButton.left = nNextXOffsetCenter;
		m_rectButton.right = m_rectButton.left + nButtonWidth;
		nNextXOffsetCenter = m_rectButton.right + m_nHorzElementOffset;
	}

	if (bTextCenter)
	{
		m_rectText.left = nNextXOffsetCenter;
		m_rectText.right = m_rectText.left + sizeText.cx; 
		AdjustRectToMargin (m_rectText, rectClient, m_nMargin);
		m_rectDrawText = m_rectText;
	}

	if (IsImageSet ())
	{
		// do not retain image size if it should be stretched
		AdjustRectToMargin (m_rectImage, rectClient, m_nMargin, !m_bStretchImage);

		if (m_rectImage.left < rectClient.left || 
			m_rectImage.right > rectClient.right)
		{
			m_rectImage.SetRectEmpty ();
		}
	}

	CRect rectButtonTemp = m_rectButton;
	if (!m_strBtnText.IsEmpty () && IsImageSet ())
	{
		CheckRectangle (rectButtonTemp, m_rectImage, bButtonLeftOfIcon);
	}

	if (!m_strBtnText.IsEmpty ())
	{
		AdjustRectToMargin (rectButtonTemp, rectClient, m_nMargin);

		if (m_rectButton.Width () + m_rectImage.Width () + 2 * m_nMargin > rectClient.Width ())
		{
			m_rectButton.SetRectEmpty ();
		}
	}

	if (!m_strText.IsEmpty ())
	{
		CheckRectangle (m_rectDrawText, m_rectImage, bTextLeftOfIcon); 
		CheckRectangle (m_rectDrawText, rectButtonTemp, bTextLeftOfButton || bButtonAfterText);
	}

	dc.SelectObject (pOldFont);

	m_bTextIsTruncated = m_rectDrawText.Width () < sizeText.cx;

	UpdateTooltips ();
}
//*****************************************************************************
BOOL CBCGPCaptionBar::CheckRectangle (CRect& rectSrc, const CRect& rectOther, 
									 BOOL bLeftOf)
{
	if (rectSrc.IsRectEmpty () || rectOther.IsRectEmpty ())
	{
		return FALSE;
	}

	CRect rectOtherWithOffset = rectOther;
	rectOtherWithOffset.InflateRect (m_nHorzElementOffset, m_nHorzElementOffset);

	if (rectSrc.left <= rectOtherWithOffset.right && 
		rectSrc.left >= rectOtherWithOffset.left)
	{
		rectSrc.left = rectOtherWithOffset.right;
	}

	if (rectSrc.right >= rectOtherWithOffset.left &&
		rectSrc.right <= rectOtherWithOffset.right)
	{
		rectSrc.right = rectOtherWithOffset.left;
	}

	if (rectSrc.left >= rectOtherWithOffset.left && 
		rectSrc.right <= rectOtherWithOffset.right)
	{
		rectSrc.right = rectSrc.left;
	}

	if (rectSrc.left <= rectOtherWithOffset.left && 
		rectSrc.right >= rectOtherWithOffset.right)
	{
		if (bLeftOf)
		{
			rectSrc.right = rectOtherWithOffset.left;
		}
		else
		{
			rectSrc.left = rectOtherWithOffset.right;
		}
	}

	if (bLeftOf && rectSrc.left >= rectOtherWithOffset.right ||
		!bLeftOf && rectSrc.right <= rectOtherWithOffset.left)
	{
		rectSrc.left = rectSrc.right;
	}

	return FALSE;
}
//*****************************************************************************
void CBCGPCaptionBar::AdjustRectToMargin (CRect& rectSrc, const CRect& rectClient, 
										 int nMargin, BOOL bRetainSize)
{
	BOOL bLeftChanged = FALSE;
	BOOL bRightChanged = FALSE;
	int nWidth = rectSrc.Width ();
	if (rectSrc.left < rectClient.left + nMargin)
	{
		rectSrc.left = rectClient.left + nMargin;
		bLeftChanged = TRUE;
	}

	if (rectSrc.right > rectClient.right - nMargin)
	{
		rectSrc.right = rectClient.right - nMargin;
		bRightChanged = TRUE;
	}

	if (bRetainSize)
	{
		if (bLeftChanged)   
		{
			rectSrc.right = rectSrc.left + nWidth;
		}
		else if (bRightChanged)
		{
			rectSrc.left = 	rectSrc.right - nWidth;
		}
	}
}
//*****************************************************************************
CSize CBCGPCaptionBar::GetImageSize () const
{
	if (m_Bitmap.GetCount () > 0)
	{
		ASSERT (m_hIcon == NULL);
		return m_Bitmap.GetImageSize ();
	}

	return globalUtils.GetIconSize(m_hIcon);
}
//*****************************************************************************
BOOL CBCGPCaptionBar::OnEraseBkgnd(CDC* /*pDC*/)
{
	return TRUE;
}
//*************************************************************************************
void CBCGPCaptionBar::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CBCGPControlBar::OnLButtonDown(nFlags, point);

	if (m_uiBtnID != 0 && m_bBtnEnabled && m_bIsBtnHighlighted)
	{
		m_bIsBtnPressed = TRUE;
		InvalidateRect (m_rectButton);
		UpdateWindow ();

		if (m_bBtnHasDropDownArrow)
		{
			ASSERT_VALID (GetOwner ());
			GetOwner()->SendMessage (WM_COMMAND, m_uiBtnID);
		}
	}

	if (m_bIsCloseBtnHighlighted)
	{
		m_bIsCloseBtnPressed= TRUE;
		InvalidateRect (m_rectClose);
		UpdateWindow ();
	}

	m_nHyperlinkPressed = HitTestHyperlink(point);
}
//*************************************************************************************
void CBCGPCaptionBar::OnLButtonUp(UINT nFlags, CPoint point) 
{
	CBCGPControlBar::OnLButtonUp(nFlags, point);

	if (m_bIsBtnPressed)
	{
		m_bIsBtnPressed = FALSE;
		InvalidateRect (m_rectButton);
		UpdateWindow ();

		if (!m_bBtnHasDropDownArrow && m_uiBtnID != 0)
		{
			ASSERT_VALID (GetOwner ());
			GetOwner()->SendMessage (WM_COMMAND, m_uiBtnID);
		}
	}
	else if (m_bIsCloseBtnPressed)
	{
		m_bIsCloseBtnPressed = FALSE;

		InvalidateRect (m_rectClose);
		UpdateWindow ();

		if (OnClickCloseButton())
		{
			ShowControlBar (FALSE, FALSE, FALSE);
		}
	}
	else if (m_nHyperlinkPressed >= 0)
	{
		if (m_nHyperlinkPressed == m_nHyperlinkHighlighted)
		{
			ASSERT_VALID (GetOwner ());
			GetOwner()->SendMessage(BCGM_ON_CLICK_CAPTIONBAR_HYPERLINK, m_nHyperlinkPressed, (LPARAM)this);
		}

		m_nHyperlinkPressed = -1;
	}
}
//*************************************************************************************
void CBCGPCaptionBar::OnMouseMove(UINT nFlags, CPoint point) 
{
	CBCGPControlBar::OnMouseMove(nFlags, point);

	BOOL bTrack = FALSE;

	if (m_uiBtnID != 0 && m_bBtnEnabled)
	{
		BOOL bIsBtnHighlighted = m_rectButton.PtInRect (point);

		if (m_bIsBtnHighlighted != bIsBtnHighlighted)
		{
			m_bIsBtnHighlighted = bIsBtnHighlighted;
			m_bIsBtnPressed = (nFlags & MK_LBUTTON) && m_bIsBtnHighlighted;

			InvalidateRect (m_rectButton);
			UpdateWindow ();

			bTrack = bIsBtnHighlighted;
		}
	}

	if (!m_rectClose.IsRectEmpty ())
	{
		BOOL bIsBtnHighlighted = m_rectClose.PtInRect (point);

		if (m_bIsCloseBtnHighlighted != bIsBtnHighlighted)
		{
			m_bIsCloseBtnHighlighted = bIsBtnHighlighted;
			m_bIsCloseBtnPressed = (nFlags & MK_LBUTTON) && m_bIsCloseBtnHighlighted;

			InvalidateRect (m_rectClose);
			UpdateWindow ();

			bTrack = bIsBtnHighlighted;
		}
	}

	int nHyperlinkHighlighted = HitTestHyperlink(point);
	if (m_nHyperlinkHighlighted != nHyperlinkHighlighted)
	{
		if (m_nHyperlinkHighlighted != -1)
		{
			InvalidateRect(&m_arHyperlinksRects[m_nHyperlinkHighlighted]);
		}

		m_nHyperlinkHighlighted = nHyperlinkHighlighted;

		if (m_nHyperlinkHighlighted != -1)
		{
			InvalidateRect(&m_arHyperlinksRects[m_nHyperlinkHighlighted]);
		}

		UpdateWindow();
		bTrack = TRUE;
	}

	if (!m_bTracked && bTrack)
	{
		m_bTracked = TRUE;
		
		TRACKMOUSEEVENT trackmouseevent;
		trackmouseevent.cbSize = sizeof(trackmouseevent);
		trackmouseevent.dwFlags = TME_LEAVE;
		trackmouseevent.hwndTrack = GetSafeHwnd();
		trackmouseevent.dwHoverTime = HOVER_DEFAULT;
		::BCGPTrackMouse (&trackmouseevent);	
	}
}
//*****************************************************************************************
LRESULT CBCGPCaptionBar::OnMouseLeave(WPARAM,LPARAM)
{
	m_bTracked = FALSE;

	if (m_bIsBtnPressed || m_bIsBtnHighlighted)
	{
		m_bIsBtnPressed = FALSE;
		m_bIsBtnHighlighted = FALSE;

		InvalidateRect (m_rectButton);
		UpdateWindow ();
	}

	if (m_bIsCloseBtnPressed || m_bIsCloseBtnHighlighted)
	{
		m_bIsCloseBtnPressed = FALSE;
		m_bIsCloseBtnHighlighted = FALSE;

		InvalidateRect (m_rectClose);
		UpdateWindow ();
	}

	if (m_nHyperlinkHighlighted != -1)
	{
		if (m_nHyperlinkHighlighted != -1)
		{
			InvalidateRect(&m_arHyperlinksRects[m_nHyperlinkHighlighted]);
		}

		m_nHyperlinkHighlighted = -1;

		UpdateWindow();
	}

	return 0;
}
//***************************************************************************************
void CBCGPCaptionBar::AdjustLayout ()
{
	if (GetSafeHwnd () == NULL)
	{
		return;
	}

	CFrameWnd* pParent = BCGPGetParentFrame (this);
	if (pParent != NULL && pParent->GetSafeHwnd () != NULL)
	{
		pParent->RecalcLayout ();
	}

	RecalcLayout ();
}
//*********************************************************************************
void CBCGPCaptionBar::OnRButtonUp(UINT nFlags, CPoint point) 
{
	if (!CBCGPToolBar::IsCustomizeMode ())
	{
		ASSERT_VALID (GetOwner ());

		ClientToScreen (&point);
		OnControlBarContextMenu (GetOwner (), point);
		return;
	}
	
	CBCGPControlBar::OnRButtonUp(nFlags, point);
}
//*********************************************************************************
void CBCGPCaptionBar::OnDestroy() 
{
	CBCGPTooltipManager::DeleteToolTip (m_pToolTip);
	CBCGPControlBar::OnDestroy();
}
//**********************************************************************************
BOOL CBCGPCaptionBar::OnNeedTipText(UINT /*id*/, NMHDR* pNMH, LRESULT* /*pResult*/)
{
	static CString strTipText;

	if (m_pToolTip->GetSafeHwnd () == NULL || 
		pNMH->hwndFrom != m_pToolTip->GetSafeHwnd ())
	{
		return FALSE;
	}

	if (CBCGPPopupMenu::GetActiveMenu () != NULL)
	{
		return FALSE;
	}

	LPNMTTDISPINFO	pTTDispInfo	= (LPNMTTDISPINFO) pNMH;
	ASSERT((pTTDispInfo->uFlags & TTF_IDISHWND) == 0);

	CString strDescr;

	switch (pNMH->idFrom)
	{
	case nIdToolTipClose:
		{
			CBCGPLocalResource locaRes;
			strTipText.LoadString (IDS_BCGBARRES_CLOSEBAR);
		}
		break;

	case nIdToolTipText:
		strTipText = m_strText;
		strTipText.Remove (_T('\a'));
		strTipText.Remove (_T('\b'));
		break;

	case nIdToolTipImage:
		strTipText = m_strImageToolTip;
		strDescr = m_strImageDescription;
		break;

	case nIdToolTipButton:
		strTipText = m_strButtonToolTip;
		strDescr = m_strButtonDescription;
		break;
	}

	if (strTipText.IsEmpty ())
	{
		return TRUE;
	}

	CBCGPToolTipCtrl* pToolTip = DYNAMIC_DOWNCAST (
		CBCGPToolTipCtrl, m_pToolTip);

	if (pToolTip != NULL && !strDescr.IsEmpty ())
	{
		ASSERT_VALID (pToolTip);
		pToolTip->SetDescription (strDescr);
	}

	pTTDispInfo->lpszText = const_cast<LPTSTR> ((LPCTSTR) strTipText);
	return TRUE;
}
//*****************************************************************************
BOOL CBCGPCaptionBar::PreTranslateMessage(MSG* pMsg) 
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

	return CBCGPControlBar::PreTranslateMessage(pMsg);
}
//**************************************************************************
LRESULT CBCGPCaptionBar::OnBCGUpdateToolTips (WPARAM wp, LPARAM)
{
	UINT nTypes = (UINT) wp;

	if (nTypes & BCGP_TOOLTIP_TYPE_CAPTIONBAR)
	{
		CBCGPTooltipManager::CreateToolTip (m_pToolTip, this,
			BCGP_TOOLTIP_TYPE_CAPTIONBAR);

		CRect rectDummy (0, 0, 0, 0);

		m_pToolTip->SetMaxTipWidth (640);

		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectDummy, nIdToolTipClose);
		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectDummy, nIdToolTipText);
		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectDummy, nIdToolTipImage);
		m_pToolTip->AddTool (this, LPSTR_TEXTCALLBACK, &rectDummy, nIdToolTipButton);
	}

	return 0;
}
//**************************************************************************
void CBCGPCaptionBar::UpdateTooltips ()
{
	if (m_pToolTip->GetSafeHwnd () != NULL)
	{
		m_pToolTip->SetToolRect (this, nIdToolTipClose, m_rectClose);

		if (m_bTextIsTruncated)
		{
			m_pToolTip->SetToolRect (this, nIdToolTipText, m_rectDrawText);
		}
		else
		{
			m_pToolTip->SetToolRect (this, nIdToolTipText, CRect (0, 0, 0, 0));
		}

		if (!m_strImageToolTip.IsEmpty ())
		{
			m_pToolTip->SetToolRect (this, nIdToolTipImage, m_rectImage);
		}
		else
		{
			m_pToolTip->SetToolRect (this, nIdToolTipImage, CRect (0, 0, 0, 0));
		}

		if (!m_strButtonToolTip.IsEmpty ())
		{
			m_pToolTip->SetToolRect (this, nIdToolTipButton, m_rectButton);
		}
		else
		{
			m_pToolTip->SetToolRect (this, nIdToolTipButton, CRect (0, 0, 0, 0));
		}
	}
}
//**************************************************************************
CSize CBCGPCaptionBar::GetTextSize (CDC* pDC, const CString& strText)
{
	ASSERT_VALID (pDC);

	if (m_arTextParts.GetSize () == 1)
	{
		return pDC->GetTextExtent (strText);
	}

	CSize sizeText (0, 0);

	for (int i = 0; i < m_arTextParts.GetSize (); i++)
	{
		if (!m_arTextParts [i].IsEmpty ())
		{
			CFont* pOldFont = NULL;

			if (m_arTextPartTypes[i] == 1)
			{
				pOldFont = pDC->SelectObject (&globalData.fontUnderline);
			}
			else if (m_arTextPartTypes[i] == 2)
			{
				pOldFont = pDC->SelectObject (&globalData.fontBold);
			}

			CSize sizePart = pDC->GetTextExtent (m_arTextParts [i]);

			sizeText.cx += sizePart.cx;
			sizeText.cy = max (sizeText.cy, sizePart.cy);

			if (pOldFont != NULL)
			{
				pDC->SelectObject (pOldFont);
			}
		}
	}

	return sizeText;
}
//**************************************************************************
BOOL CBCGPCaptionBar::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	CPoint ptCursor;
	::GetCursorPos (&ptCursor);

	ScreenToClient (&ptCursor);

	if (HitTestHyperlink(ptCursor) >= 0)
	{
		::SetCursor (globalData.GetHandCursor ());
		return TRUE;
	}
	
	return CBCGPControlBar::OnSetCursor(pWnd, nHitTest, message);
}
//***************************************************************************
int CBCGPCaptionBar::HitTestHyperlink(CPoint point)
{
	for (int i = 0; i < m_arHyperlinksRects.GetSize(); i++)
	{
		CRect rect = m_arHyperlinksRects[i];
		if (rect.PtInRect(point))
		{
			return i;
		}
	}

	return -1;
}
//***************************************************************************
void CBCGPCaptionBar::SetMessageBarMode(BOOL bIsMessageBarMode)
{
	if (m_bIsMessageBarMode == bIsMessageBarMode)
	{
		return;
	}

	m_bIsMessageBarMode = bIsMessageBarMode;

	AdjustLayout();
	RedrawWindow();
}
//***************************************************************************
CString CBCGPCaptionBar::GetLinkText (int nIndex)
{
	int count = -1;
	for (int i=0; i<m_arTextParts.GetSize (); i++)
	{
		if (m_arTextPartTypes [i] == 1)
		{
			count++;
			if (count == nIndex)
			{
				return m_arTextParts [i];
			}
		}
	}

	return _T("");
}
//***************************************************************************
BOOL CBCGPCaptionBar::OnSetAccData (long lVal)
{
	ASSERT_VALID (this);

	m_AccData.Clear ();

	int nObjectId = AccGetObjectByChildId(lVal);
	if (nObjectId != -1)
	{
		if (nObjectId == ACC_BUTTON)	
		{
			m_AccData.m_strAccName = m_strBtnText;
			m_AccData.m_strAccDefAction = m_bBtnHasDropDownArrow ? _T("Open") : _T("Press");
			m_AccData.m_strDescription = m_strButtonDescription;
			m_AccData.m_nAccHit = 1;
			m_AccData.m_nAccRole = m_bBtnHasDropDownArrow ? ROLE_SYSTEM_BUTTONDROPDOWN : ROLE_SYSTEM_PUSHBUTTON;
			m_AccData.m_bAccState = STATE_SYSTEM_FOCUSABLE;

			if (!m_bBtnEnabled)
			{
				m_AccData.m_bAccState |= STATE_SYSTEM_UNAVAILABLE;
			}

			if (m_bIsBtnPressed || m_bIsBtnForcePressed)
			{
			  m_AccData.m_bAccState |= STATE_SYSTEM_FOCUSED | STATE_SYSTEM_PRESSED;
			}
			
			if (m_bIsBtnHighlighted)
			{
				m_AccData.m_bAccState |= STATE_SYSTEM_HOTTRACKED;
			}
			
			m_AccData.m_rectAccLocation = m_rectButton;
			ClientToScreen (&m_AccData.m_rectAccLocation);
			return TRUE;
		}

		if (nObjectId >= ACC_LINK)	
		{
			int nHyperLink = nObjectId - ACC_LINK;

			m_AccData.m_strAccName = GetLinkText (nHyperLink);
			m_AccData.m_strAccDefAction = _T("Jump");
			m_AccData.m_nAccHit = 1;
			m_AccData.m_nAccRole = ROLE_SYSTEM_LINK;
			m_AccData.m_bAccState = STATE_SYSTEM_FOCUSABLE;

			if (m_nHyperlinkPressed == nHyperLink)
			{
			  m_AccData.m_bAccState |= STATE_SYSTEM_FOCUSED;
			}
			
			if (m_nHyperlinkHighlighted == nHyperLink)
			{
				m_AccData.m_bAccState |= STATE_SYSTEM_HOTTRACKED;
			}
			
			m_AccData.m_rectAccLocation = m_arHyperlinksRects [nHyperLink];
			ClientToScreen (&m_AccData.m_rectAccLocation);
			return TRUE;
		}

		if (nObjectId == ACC_CLOSE_BUTTON)
		{
			m_AccData.m_strAccName = _T("Close");
			m_AccData.m_strAccDefAction = _T("Press");
			m_AccData.m_strDescription = _T("Close CaptionBar button");
			m_AccData.m_nAccHit = 1;
			m_AccData.m_nAccRole = ROLE_SYSTEM_PUSHBUTTON;
			m_AccData.m_bAccState = STATE_SYSTEM_FOCUSABLE;

			if (m_bIsCloseBtnPressed)
			{
				m_AccData.m_bAccState |= STATE_SYSTEM_FOCUSED | STATE_SYSTEM_PRESSED;
			}
			
			if (m_bIsCloseBtnHighlighted || m_bCloseTracked)
			{
				m_AccData.m_bAccState |= STATE_SYSTEM_HOTTRACKED;
			}
			
			m_AccData.m_rectAccLocation = m_rectClose;
			ClientToScreen (&m_AccData.m_rectAccLocation);
			return TRUE;
		}
	}

	return FALSE;
}
//*****************************************************************************
HRESULT CBCGPCaptionBar::get_accChildCount(long *pcountChildren)
{
	if(!pcountChildren )
    {
        return E_INVALIDARG;
    }

	int count = 0;

	m_mapObjectToIndex.RemoveAll ();
	
	if (m_uiBtnID != 0 &&  !m_rectButton.IsRectEmpty ())
	{
		count++;
		m_mapObjectToIndex.SetAt (ACC_BUTTON, count);
	}

	if (!m_rectClose.IsRectEmpty ())
	{
		count++;
		m_mapObjectToIndex.SetAt (ACC_CLOSE_BUTTON, count);
	}

	for (int i=0; i<m_arHyperlinksRects.GetSize (); i++)
	{
		count++;
		int nIndex =ACC_LINK+i;
		m_mapObjectToIndex.SetAt (nIndex, count);
	}

	*pcountChildren = count;
	return S_OK;
}
//***************************************************************************
int CBCGPCaptionBar::AccGetObjectByChildId(long lVal)
{
	POSITION pos = 0;

	pos = m_mapObjectToIndex.GetStartPosition ();
	while (pos != NULL)
	{
		int nObjectId = 0;
		int nIndex = 0;

		m_mapObjectToIndex.GetNextAssoc (pos, nObjectId, nIndex);
		if (lVal == nIndex)
		{
			return nObjectId;
		}
	}

	return -1;
}
//***************************************************************************
int CBCGPCaptionBar::AccGetChildIdByObjectId(int nObjectId)
{
	int nIndex = -1;
	m_mapObjectToIndex.Lookup (nObjectId, nIndex);
	return nIndex;
}
//***************************************************************************
HRESULT CBCGPCaptionBar::get_accChild(VARIANT /*varChild*/, IDispatch **ppdispChild)
{
	if(!ppdispChild )
    {
        return E_INVALIDARG;
    }

	*ppdispChild = NULL;

	return S_FALSE;
}
//***************************************************************************
HRESULT CBCGPCaptionBar::get_accName(VARIANT varChild, BSTR *pszName)
{
	if (varChild.vt == VT_I4 && varChild.lVal == CHILDID_SELF)
	{
		CString strText = _T("Caption Bar");
		*pszName = strText.AllocSysString();
		return S_OK;
	}

	return CBCGPControlBar::get_accName(varChild, pszName);
}
//***************************************************************************
HRESULT CBCGPCaptionBar::get_accValue(VARIANT varChild, BSTR *pszValue)
{
	if (varChild.vt == VT_I4 && varChild.lVal == CHILDID_SELF)
	{
		CString strText = m_strText;
		strText.Replace(_T("\a"), _T(""));
		strText.Replace(_T("\b"), _T(""));
		*pszValue = strText.AllocSysString();
		return S_OK;
	}

	return CBCGPControlBar::get_accValue(varChild, pszValue);
}
//***************************************************************************
HRESULT CBCGPCaptionBar::get_accRole(VARIANT varChild, VARIANT *pvarRole)
{
	if (pvarRole == NULL)
    {
        return E_INVALIDARG;
    }

	if (varChild.vt == VT_I4 && varChild.lVal == CHILDID_SELF)
	{
		pvarRole->vt = VT_I4;
		pvarRole->lVal = ROLE_SYSTEM_CLIENT;
		return S_OK;
	}

	return CBCGPControlBar::get_accRole(varChild, pvarRole);
}
//***************************************************************************
HRESULT CBCGPCaptionBar::get_accState(VARIANT varChild, VARIANT *pvarState)
{
	if (pvarState == NULL)
    {
        return E_INVALIDARG;
    }

	if (varChild.vt == VT_I4 && varChild.lVal == CHILDID_SELF)
	{
		pvarState->vt = VT_I4;
		pvarState->lVal = STATE_SYSTEM_NORMAL;
		return S_OK;
	}

	return CBCGPControlBar::get_accState(varChild, pvarState);

}
//***************************************************************************
HRESULT CBCGPCaptionBar::accHitTest(long xLeft, long yTop, VARIANT *pvarChild)
{
	if( !pvarChild )
    {
        return E_INVALIDARG;
    }

	pvarChild->vt = VT_I4;
	pvarChild->lVal = CHILDID_SELF;

	CPoint pt (xLeft, yTop);
	ScreenToClient (&pt);

	if (m_uiBtnID != 0 && m_rectButton.PtInRect (pt))
	{
		int nIndex = AccGetChildIdByObjectId(ACC_BUTTON);
		pvarChild->lVal = nIndex;
		OnSetAccData (nIndex);
		return S_OK;
	}	

	if (m_rectClose.PtInRect (pt))
	{
		int nIndex = AccGetChildIdByObjectId(ACC_CLOSE_BUTTON);
		pvarChild->lVal = nIndex;
		OnSetAccData (nIndex);
		return S_OK;
	}

	int nLink = HitTestHyperlink(pt);
	if (nLink != -1)
	{
		int nIndex = AccGetChildIdByObjectId(ACC_LINK + nLink);
		pvarChild->lVal = nIndex;
		OnSetAccData (nIndex);
		return S_OK;
	}

	return S_OK;
}
//***************************************************************************
HRESULT CBCGPCaptionBar::accDoDefaultAction(VARIANT varChild)
{
	if (varChild.vt != VT_I4)
    {
        return E_INVALIDARG;
    }

	int nObjectId = AccGetObjectByChildId(varChild.lVal);
	if (nObjectId != -1)
	{
		if (nObjectId == ACC_BUTTON)
		{
			if (m_uiBtnID != 0)
			{
				ASSERT_VALID (GetOwner ());
				GetOwner()->SendMessage (WM_COMMAND, m_uiBtnID);
			}
		}

		if (nObjectId == ACC_CLOSE_BUTTON)
		{
			if (OnClickCloseButton())
			{
				ShowControlBar (FALSE, FALSE, FALSE);
			}
		}

		if (nObjectId >= ACC_LINK)
		{
			ASSERT_VALID (GetOwner ());

			int nHyperLink = nObjectId - ACC_LINK;
			GetOwner()->SendMessage(BCGM_ON_CLICK_CAPTIONBAR_HYPERLINK, nHyperLink, (LPARAM)this);
		}
	}

    return S_OK;
}
//***************************************************************************
CFont* CBCGPCaptionBar::SelectFont(CDC* pDC)
{
	return pDC->SelectObject(m_hFont == NULL ? (m_bUseCaptionFont ? &globalData.fontCaption : &globalData.fontRegular) : CFont::FromHandle (m_hFont));
}

