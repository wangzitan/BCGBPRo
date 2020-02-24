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
// BCGPPageTransitionManager.cpp: implementation of the CBCGPPageTransitionManager class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "math.h"
#include "BCGPDrawManager.h"
#include "BCGPPageTransitionManager.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPPageTransitionManager::CBCGPPageTransitionManager()
{
	m_PageTransitionEffect = BCGPPageTransitionNone;
	m_animationType = CBCGPAnimationManager::BCGPANIMATION_Legacy;
	m_nPageTransitionTime = 300;
	m_nPageTransitionOffset = 0;
	m_nPageTransitionStep = 0;
	m_nPageTransitionTotal = 0;
	m_nPageScreenshot1 = 0;
	m_nPageScreenshot2 = 0;
	m_rectPageTransition.SetRectEmpty();
	m_hwndHost = NULL;
	m_clrFillFrame = (COLORREF)-1;
}
//*****************************************************************************************
CBCGPPageTransitionManager::~CBCGPPageTransitionManager()
{
	StopPageTransition();
}
//*****************************************************************************************
void CBCGPPageTransitionManager::StopPageTransition(BOOL bNotify)
{
	StopAnimation(bNotify);

	m_nPageTransitionOffset = 0;
	m_nPageTransitionStep = 0;
	m_nPageTransitionTotal = 0;
	m_rectPageTransition.SetRectEmpty();
	m_Panorama.Clear();
}
//*****************************************************************************************
BOOL CBCGPPageTransitionManager::StartPageTransition(HWND hwdHost, const CArray<HWND, HWND>& arPages, BOOL bReverseOrder, 
													 const CSize& szPageOffset, const CSize& szPageMax)
{
	StopPageTransition(FALSE);

	if (hwdHost == NULL || arPages.GetSize() < 2 || m_nPageTransitionTime < 1 || m_PageTransitionEffect == BCGPPageTransitionNone)
	{
		return FALSE;
	}

	CWnd* pWndCurrPage = CWnd::FromHandle(arPages[0]);

	m_hwndHost = hwdHost;
	CWnd* pWndHost = CWnd::FromHandle(m_hwndHost);

	const BOOL bIsRTL = (pWndHost->GetExStyle() & WS_EX_LAYOUTRTL);

	pWndCurrPage->GetWindowRect(m_rectPageTransition);

	if (szPageMax.cx != 0)
	{
		m_rectPageTransition.right = m_rectPageTransition.left + szPageMax.cx;
	}

	if (szPageMax.cy != 0)
	{
		m_rectPageTransition.bottom = m_rectPageTransition.top + szPageMax.cy;
	}

	int cxOffset = bIsRTL ? 0 : szPageOffset.cx;
	int cx = m_rectPageTransition.Width() - cxOffset;
	int cy = m_rectPageTransition.Height() - szPageOffset.cy;

	pWndHost->ScreenToClient(&m_rectPageTransition);

	// Create panorama bitmap:
	m_Panorama.SetImageSize(CSize(cx, cy));

	const BOOL bIsAppWnd = AfxGetMainWnd()->GetSafeHwnd() == m_hwndHost;

	if (bIsAppWnd)
	{
		pWndHost->LockWindowUpdate();
	}
	else
	{
		pWndHost->SetRedraw(FALSE);
	}

	BOOL bOnePageTransition = m_PageTransitionEffect == BCGPPageTransitionAppear;

	BOOL bTwoPagesTransition = 
		m_PageTransitionEffect != BCGPPageTransitionSlide && 
		m_PageTransitionEffect != BCGPPageTransitionSlideVertical;

	m_nPageScreenshot1 = bReverseOrder ? 1 : 0;
	m_nPageScreenshot2 = bReverseOrder ? 0 : 1;

	HBITMAP hbmp = CBCGPDrawManager::CreateBitmap_32(CSize(cx, cy), NULL);
	if (hbmp == NULL)
	{
		StopPageTransition();
		return FALSE;
	}

	CDC dcMem;
	dcMem.CreateCompatibleDC(NULL);

	int nFinishPage = bReverseOrder ? 0 : (int)arPages.GetSize() - 1;
	
	for (int i = 0; i < arPages.GetSize(); i++)
	{
		if (bOnePageTransition && i != nFinishPage)
		{
			continue;
		}

		if (bTwoPagesTransition && (i != 0 && i != arPages.GetSize() - 1))
		{
			continue;
		}

		CWnd* pPage = CWnd::FromHandle(arPages[i]);
		ASSERT(pPage->GetSafeHwnd() != NULL);

		HBITMAP hbmpOld = (HBITMAP) dcMem.SelectObject(hbmp);

		if (m_clrFillFrame != (COLORREF)-1)
		{
			dcMem.FillSolidRect(CRect(0, 0, cx, cy), m_clrFillFrame);
		}

		pPage->SendMessage(WM_PRINT, (WPARAM)dcMem.GetSafeHdc(), (LPARAM)(PRF_CLIENT | PRF_CHILDREN | PRF_NONCLIENT | PRF_ERASEBKGND));
		dcMem.SelectObject (hbmpOld);

		m_Panorama.AddImage(hbmp, TRUE);
	}

	if (bIsRTL)
	{
		m_Panorama.Mirror();
	}

	::DeleteObject(hbmp);

	pWndCurrPage->ShowWindow(SW_HIDE);

	if (bIsAppWnd)
	{
		pWndHost->UnlockWindowUpdate();
	}
	else
	{
		pWndHost->SetRedraw(TRUE);
	}
	
	return StartInternal(bReverseOrder);
}
//*****************************************************************************************
BOOL CBCGPPageTransitionManager::StartInternal(BOOL bReverseOrder)
{
	int nFinishValue = 0;

	if (m_PageTransitionEffect == BCGPPageTransitionFade)
	{
		nFinishValue = m_nPageTransitionTotal = 255;
		m_nPageTransitionOffset = 0;
	}
	else if (m_PageTransitionEffect == BCGPPageTransitionPop)
	{
		nFinishValue = m_nPageTransitionTotal = min(m_rectPageTransition.Width(), m_rectPageTransition.Height()); 
		m_nPageTransitionOffset = 0;
	}
	else if (m_PageTransitionEffect == BCGPPageTransitionAppear)
	{
		m_nPageTransitionTotal = m_rectPageTransition.Width() / 10;
		m_nPageTransitionOffset = bReverseOrder ? m_nPageTransitionTotal : -m_nPageTransitionTotal;

		nFinishValue = 0;
	}
	else
	{
		BOOL bIsVerticalSliding =
			m_PageTransitionEffect == BCGPPageTransitionSlideVertical ||
			m_PageTransitionEffect == BCGPPageTransitionSimpleSlideVertical;

		if (bIsVerticalSliding)
		{
			m_nPageTransitionTotal = (m_Panorama.GetCount() - 1) * m_Panorama.GetImageSize().cy;
		}
		else
		{
			m_nPageTransitionTotal = (m_Panorama.GetCount() - 1) * m_Panorama.GetImageSize().cx;
		}

		if (!bReverseOrder)
		{
			m_nPageTransitionOffset = 0;
			nFinishValue = m_nPageTransitionTotal;
		}
		else
		{
			m_nPageTransitionOffset = m_nPageTransitionTotal;
			nFinishValue = 0;
		}
	}

	if (m_nPageTransitionTotal == 0)
	{
		StopPageTransition();
		return FALSE;
	}

	if (!StartAnimation(m_nPageTransitionOffset, nFinishValue, .001 * m_nPageTransitionTime,
		m_animationType, &m_animationOptions))
	{
		StopPageTransition();
		return FALSE;
	}

	::RedrawWindow(m_hwndHost, m_rectPageTransition, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
	return TRUE;
}
//*****************************************************************************************
BOOL CBCGPPageTransitionManager::StartPageTransition(HWND hwdHost, HWND hwndPageFrom, HWND hwndPageTo, BOOL bReverseOrder, 
													 const CSize& szPageOffset, const CSize& szPageMax)
{
	CArray<HWND, HWND> arPages;

	if (bReverseOrder)
	{
		arPages.Add(hwndPageTo);
		arPages.Add(hwndPageFrom);
	}
	else
	{
		arPages.Add(hwndPageFrom);
		arPages.Add(hwndPageTo);
	}

	return StartPageTransition(hwdHost, arPages, bReverseOrder, szPageOffset, szPageMax);
}
//*****************************************************************************************
BOOL CBCGPPageTransitionManager::StartBitmapTransition(HWND hwdHost, const CArray<HBITMAP, HBITMAP>& arPages, const CRect& rectPageTransition, BOOL bReverseOrder)
{
	StopPageTransition(FALSE);

	if (hwdHost == NULL || arPages.GetSize() < 2 || m_nPageTransitionTime < 1 || m_PageTransitionEffect == BCGPPageTransitionNone)
	{
		return FALSE;
	}

	m_hwndHost = hwdHost;
	m_rectPageTransition = rectPageTransition;

	int cx = m_rectPageTransition.Width();
	int cy = m_rectPageTransition.Height();

	// Create panorama bitmap:
	m_Panorama.SetImageSize(CSize(cx, cy));

	BOOL bTwoPagesTransition = 
		m_PageTransitionEffect != BCGPPageTransitionSlide && 
		m_PageTransitionEffect != BCGPPageTransitionSlideVertical;

	m_nPageScreenshot1 = bReverseOrder ? 1 : 0;
	m_nPageScreenshot2 = bReverseOrder ? 0 : 1;

	HBITMAP hbmp = CBCGPDrawManager::CreateBitmap_32(CSize(cx, cy), NULL);
	if (hbmp == NULL)
	{
		StopPageTransition();
		return FALSE;
	}

	CDC dcMem;
	dcMem.CreateCompatibleDC(NULL);
	
	for (int i = 0; i < arPages.GetSize(); i++)
	{
		if (bTwoPagesTransition && (i != 0 && i != arPages.GetSize() - 1))
		{
			continue;
		}

		HBITMAP hbmpOld = (HBITMAP) dcMem.SelectObject(hbmp);

		if (m_clrFillFrame != (COLORREF)-1)
		{
			dcMem.FillSolidRect(CRect(0, 0, cx, cy), m_clrFillFrame);
		}

		BITMAP bmp;
		::GetObject(arPages[i], sizeof(BITMAP), (LPVOID)&bmp);

		CSize sizeDraw;
		sizeDraw.cx = min(cx, bmp.bmWidth);
		sizeDraw.cy = min(cy, bmp.bmHeight);

		dcMem.DrawState(CPoint(0, 0), sizeDraw, arPages[i], DSS_NORMAL);

		dcMem.SelectObject (hbmpOld);

		m_Panorama.AddImage(hbmp, TRUE);
	}

	::DeleteObject(hbmp);

	return StartInternal(bReverseOrder);
}
//*****************************************************************************************
BOOL CBCGPPageTransitionManager::StartBitmapTransition(HWND hwdHost, HBITMAP hbmpFrom, HBITMAP hbmpTo, const CRect& rectPageTransition, BOOL bReverseOrder)
{
	CArray<HBITMAP, HBITMAP> arPages;
	
	if (bReverseOrder)
	{
		arPages.Add(hbmpTo);
		arPages.Add(hbmpFrom);
	}
	else
	{
		arPages.Add(hbmpFrom);
		arPages.Add(hbmpTo);
	}
	
	return StartBitmapTransition(hwdHost, arPages, rectPageTransition, bReverseOrder);
}
//*****************************************************************************************
void CBCGPPageTransitionManager::OnAnimationValueChanged(double /*dblOldValue*/, double dblNewValue)
{
	m_nPageTransitionOffset = (int)dblNewValue;
	
	::RedrawWindow(m_hwndHost, m_rectPageTransition, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
}
//*****************************************************************************************
void CBCGPPageTransitionManager::OnAnimationFinished() 
{ 
	OnPageTransitionFinished(); 

	m_nPageTransitionOffset = 0;
	m_nPageTransitionStep = 0;
	m_nPageTransitionTotal = 0;
	m_rectPageTransition.SetRectEmpty();
	m_Panorama.Clear();
}
//*****************************************************************************************
void CBCGPPageTransitionManager::DoDrawTransition(CDC* pDC, BOOL bIsMemDC)
{
	if (m_rectPageTransition.IsRectEmpty())
	{
		return;
	}
	
	CRgn rgnClip;
	rgnClip.CreateRectRgnIndirect(bIsMemDC ? CRect(CPoint(0, 0), m_rectPageTransition.Size()) : m_rectPageTransition);
	
	pDC->SelectClipRgn(&rgnClip);
	
	if (m_PageTransitionEffect == BCGPPageTransitionFade)
	{
		int nPageTransitionOffset = min(255, max(0, m_nPageTransitionOffset));

		m_Panorama.DrawEx(pDC, m_rectPageTransition, m_nPageScreenshot1, CBCGPToolBarImages::ImageAlignHorzLeft, CBCGPToolBarImages::ImageAlignVertTop, 
			CRect(0, 0, 0, 0), (BYTE)(255 - nPageTransitionOffset), TRUE);
		
		m_Panorama.DrawEx(pDC, m_rectPageTransition, m_nPageScreenshot2, CBCGPToolBarImages::ImageAlignHorzLeft, CBCGPToolBarImages::ImageAlignVertTop, 
			CRect(0, 0, 0, 0), (BYTE)nPageTransitionOffset, TRUE);
	}
	else if (m_PageTransitionEffect == BCGPPageTransitionPop)
	{
		m_Panorama.DrawEx(pDC, m_rectPageTransition, m_nPageScreenshot1, CBCGPToolBarImages::ImageAlignHorzLeft, CBCGPToolBarImages::ImageAlignVertTop, 
			CRect(0, 0, 0, 0), 255, TRUE);
		
		int nWidth = m_rectPageTransition.Width() * m_nPageTransitionOffset / m_nPageTransitionTotal;
		int nHeight = m_rectPageTransition.Height() * m_nPageTransitionOffset / m_nPageTransitionTotal;
		
		CRect rectNewPage(CPoint(m_rectPageTransition.CenterPoint().x - nWidth / 2, m_rectPageTransition.CenterPoint().y - nHeight / 2), CSize(nWidth, nHeight));
		
		m_Panorama.DrawEx(pDC, rectNewPage, m_nPageScreenshot2, CBCGPToolBarImages::ImageAlignHorzStretch, CBCGPToolBarImages::ImageAlignVertStretch, 
			CRect(0, 0, 0, 0), 255, TRUE);
	}
	else
	{
		const CSize sizeSlide = m_Panorama.GetImageSize();
		const BOOL bIsVeritcal = m_PageTransitionEffect == BCGPPageTransitionSlideVertical || m_PageTransitionEffect == BCGPPageTransitionSimpleSlideVertical;
		
		for (int i = 0; i < m_Panorama.GetCount(); i++)
		{
			CRect rectFrame = m_rectPageTransition;

			if (bIsVeritcal)
			{
				rectFrame.OffsetRect(0, sizeSlide.cy * i - m_nPageTransitionOffset);
			}
			else
			{
				rectFrame.OffsetRect(sizeSlide.cx * i - m_nPageTransitionOffset, 0);
			}
			
			CRect rectInter;
			if (rectInter.IntersectRect(rectFrame, m_rectPageTransition))
			{
				BYTE bAlpha = m_PageTransitionEffect == BCGPPageTransitionAppear ?
					(BYTE)min(255, max(0, (255 - (int)fabs(255.0 * (double)m_nPageTransitionOffset / m_nPageTransitionTotal)))) : (BYTE)255;

				Sleep(1);
				m_Panorama.DrawEx(pDC, rectFrame, i, CBCGPToolBarImages::ImageAlignHorzLeft, CBCGPToolBarImages::ImageAlignVertTop, CRect(0, 0, 0, 0), bAlpha, TRUE);
			}
		}
	}
	
	pDC->SelectClipRgn (NULL);
}
