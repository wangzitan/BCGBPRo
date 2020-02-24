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
// BCGPSDPlaceMarkerWnd.cpp : implementation file
//

#include "stdafx.h"
#include "bcgcbpro.h"
#include "BCGPSDPlaceMarkerWnd.h"
#include "BCGPDrawManager.h"
#include "bcgglobals.h"
#include "BCGPTabbedControlBar.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPSDPlaceMarkerWnd

CString CBCGPSDPlaceMarkerWnd::m_strClassName;

CBCGPSDPlaceMarkerWnd::CBCGPSDPlaceMarkerWnd(BCGPSDPlaceMarkerWnd_STYLE style, COLORREF color/* = CLR_DEFAULT*/, BYTE nAlpha/* = 255*/) :
	m_Style(style),
	m_nAlpha(nAlpha)
{
	if (color == CLR_DEFAULT)
	{
		color = globalData.clrBtnShadow;
	}

	if (m_nAlpha != 255)
	{
		if (globalData.IsWindowsLayerSupportAvailable () && globalData.m_nBitsPerPixel > 8)
		{
			color = CBCGPDrawManager::PixelAlpha(color, 105);
		}
		else
		{
			color = CBCGPDrawManager::ColorMakePale(color, .9);
		}
	}

	m_Brush.CreateSolidBrush(color);
	m_rectLast.SetRectEmpty();
	m_bAutoDestroy = TRUE;
}
//***********************************************************************************************************
BOOL CBCGPSDPlaceMarkerWnd::Create(const CRect& rect, CWnd* pWndOwner, BOOL bAutoDestroy)
{
	m_bAutoDestroy = bAutoDestroy;

	if (m_strClassName.IsEmpty())
	{
		m_strClassName = AfxRegisterWndClass(CS_SAVEBITS | CS_OWNDC);
	}
	
	DWORD dwExStyle = WS_EX_TOOLWINDOW | WS_EX_TOPMOST;
	if (globalData.IsWindowsLayerSupportAvailable())
	{
		dwExStyle |= WS_EX_LAYERED;
	}

	if (!CWnd::CreateEx(dwExStyle, m_strClassName, _T(""), WS_POPUP | MFS_SYNCACTIVE, 
		rect, pWndOwner == NULL ? AfxGetMainWnd() : pWndOwner, 0))
	{
		return FALSE;
	}
	
	if (pWndOwner != NULL)
	{
		SetOwner(pWndOwner);
	}
	
	if (globalData.IsWindowsLayerSupportAvailable())
	{
		globalData.SetLayeredAttrib(GetSafeHwnd (), 0, m_nAlpha, LWA_ALPHA);
	}

	if (!rect.IsRectEmpty())
	{
		UpdateRgn();
		ShowWindow(SW_SHOWNOACTIVATE);
	}

	return TRUE;
}

BEGIN_MESSAGE_MAP(CBCGPSDPlaceMarkerWnd, CWnd)
//{{AFX_MSG_MAP(CBCGPSDPlaceMarkerWnd)
	ON_WM_ERASEBKGND()
	ON_WM_PAINT()
	ON_WM_SIZE()
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BOOL CBCGPSDPlaceMarkerWnd::OnEraseBkgnd(CDC* /*pDC*/) 
{
	return TRUE;
}
//***********************************************************************************************************
void CBCGPSDPlaceMarkerWnd::OnPaint() 
{
	CPaintDC dc(this); // device context for painting

	CRect rectClient;
	GetClientRect(&rectClient);

	dc.FillRect(rectClient, &m_Brush);
}
//***********************************************************************************************************
void CBCGPSDPlaceMarkerWnd::PostNcDestroy()
{
	CWnd::PostNcDestroy();

	if (m_bAutoDestroy)
	{
		delete this;
	}
}
//***********************************************************************************************************
void CBCGPSDPlaceMarkerWnd::OnSize(UINT nType, int cx, int cy) 
{
	CWnd::OnSize(nType, cx, cy);
	UpdateRgn();
}
//***********************************************************************************************************
void CBCGPSDPlaceMarkerWnd::UpdateRgn()
{
	const int nFrameSize = globalUtils.ScaleByDPI(4);

	CRect rectClient;
	GetClientRect(rectClient);

	if (m_Style == BCGPSDPlaceMarkerWnd_Frame)
	{
		CRgn rgn;
		rgn.CreateRectRgnIndirect(rectClient);
		
		rectClient.DeflateRect(nFrameSize, nFrameSize);
		
		CRgn rgnInner;
		rgnInner.CreateRectRgnIndirect(rectClient);
		
		rgn.CombineRgn(&rgn, &rgnInner, RGN_XOR);
		
		SetWindowRgn(rgn, TRUE);
	}
	else if (m_Style == BCGPSDPlaceMarkerWnd_FrameWithTab || m_Style == BCGPSDPlaceMarkerWnd_RectangleWithTab)
	{
		const int nTabHeight = globalData.GetTextHeight();

		CRgn rgn[2];

		const int nSteps = m_Style == BCGPSDPlaceMarkerWnd_FrameWithTab ? 2 : 1;

		for (int i = 0; i < nSteps; i++)
		{
			CRect rectWnd = rectClient;
			CRect rectTab = rectClient;

			if (CBCGPTabbedControlBar::m_bTabsAlwaysTop)
			{
				rectWnd.top += nTabHeight;
				rectTab.bottom = rectWnd.top;
			}
			else
			{
				rectWnd.bottom -= nTabHeight;
				rectTab.top = rectWnd.bottom;
			}

			rgn[i].CreateRectRgnIndirect(rectWnd);

			rectTab.left += globalUtils.ScaleByDPI(10);
			rectTab.right = rectTab.left + globalUtils.ScaleByDPI(40) - i * nFrameSize * 2;

			if (rectTab.right >= rectWnd.right)
			{
				rectTab.right = rectWnd.right - nFrameSize - 4;
			}

			CRgn rgnTab;
			rgnTab.CreateRectRgnIndirect(rectTab);
			
			rgn[i].CombineRgn(&rgn[i], &rgnTab, RGN_OR);

			rectClient.DeflateRect(nFrameSize, nFrameSize);
		}

		if (nSteps == 2)
		{
			rgn[0].CombineRgn(&rgn[0], &rgn[1], RGN_XOR);
		}

		SetWindowRgn(rgn[0], TRUE);
	}
	else
	{
		SetWindowRgn(NULL, TRUE);
	}
}
//***********************************************************************************************************
void CBCGPSDPlaceMarkerWnd::ShowAt(CRect rect, BCGPSDPlaceMarkerWnd_STYLE style)
{
    if (m_Style != style || m_rectLast != rect)
	{
		m_Style = style;

		SetWindowPos (&CWnd::wndTop, rect.left, rect.top, rect.Width(), rect.Height(), SWP_NOACTIVATE | SWP_SHOWWINDOW);
		m_rectLast = rect;
	}
}
//***********************************************************************************************************
void CBCGPSDPlaceMarkerWnd::Hide()
{
    if (IsWindowVisible())
    {
        ShowWindow(SW_HIDE);
		
		if (GetOwner()->GetSafeHwnd() != NULL)
		{
			GetOwner()->UpdateWindow();
		}
		
		if (m_pDockingWnd->GetSafeHwnd() != NULL)
		{
			m_pDockingWnd->UpdateWindow();
		}
		
		m_rectLast.SetRectEmpty();
    }
}
