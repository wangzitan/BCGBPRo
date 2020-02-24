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
// BCGPTextGaugeImpl.cpp: implementation of the CBCGPTextGaugeImpl class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "bcgcbpro.h"
#include "BCGPTextGaugeImpl.h"
#include "BCGPDrawManager.h"
#include "BCGPVisualCtrl.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CBCGPTextGaugeImpl, CBCGPStaticGaugeImpl)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPTextGaugeImpl::CBCGPTextGaugeImpl(LPCTSTR lpszText, const CBCGPColor& color, CBCGPVisualContainer* pContainer) :
	CBCGPStaticGaugeImpl(pContainer)
{
	m_DefaultDrawFlags = BCGP_DRAW_DYNAMIC;

	SetText(lpszText == NULL ? _T("") : lpszText, color);
}
//*******************************************************************************
CBCGPTextGaugeImpl::CBCGPTextGaugeImpl(LPCTSTR lpszText, const CBCGPBrush& brush, CBCGPVisualContainer* pContainer) :
	CBCGPStaticGaugeImpl(pContainer)
{
	m_DefaultDrawFlags = BCGP_DRAW_DYNAMIC;

	SetText(lpszText == NULL ? _T("") : lpszText, brush);
}
//*******************************************************************************
CBCGPTextGaugeImpl::~CBCGPTextGaugeImpl()
{
}
//*******************************************************************************
void CBCGPTextGaugeImpl::OnDraw(CBCGPGraphicsManager* pGM, const CBCGPRect& /*rectClip*/, DWORD dwFlags)
{
	ASSERT_VALID(pGM);

	if (m_rect.IsRectEmpty() || !m_bIsVisible || m_strText.IsEmpty())
	{
		return;
	}

	if ((dwFlags & m_DefaultDrawFlags) == 0)
	{
		return;
	}

	CBCGPRect rect = m_rect;
	if (m_bIsSubGauge)
	{
		rect.OffsetRect(-m_ptScrollOffset);
	}

	pGM->FillRectangle(rect, GetFillBrush ());

	const CBCGPBrush& br = m_bOff ? m_brTextLight : m_brText;

	CreateResources(CBCGPBrush(), FALSE);

	CBCGPRect rectText = rect;
	BOOL bClipText = m_textFormat.IsClipText();

	if (m_dblScrollOffset != 0.0)
	{
		if (m_bIsHorizontalScroll)
		{
			rectText.OffsetRect(-m_dblScrollOffset, 0.0);
		}
		else
		{
			rectText.OffsetRect(0.0, -m_dblScrollOffset);
		}

		if (bClipText)
		{
			m_textFormat.SetClipText(FALSE);
		}

		pGM->SetClipRect(rect);
	}

	pGM->DrawText(m_strText, rectText, m_textFormat, br);

	if (m_dblScrollOffset != 0.0)
	{
		pGM->ReleaseClipArea();

		if (bClipText)
		{
			m_textFormat.SetClipText();
		}
	}

	pGM->DrawRectangle(rect, GetOutlineBrush (), GetScaleRatioMid());
	
	SetDirty(FALSE);
}
//*******************************************************************************
void CBCGPTextGaugeImpl::SetText(const CString& str, const CBCGPColor& clrText, BOOL bRedraw)
{
	m_strText = str;
	CreateResources(clrText == CBCGPColor() ? CBCGPBrush() : CBCGPBrush(clrText), FALSE);

	SetDirty();

	if (bRedraw)
	{
		Redraw();
	}
}
//*******************************************************************************
void CBCGPTextGaugeImpl::SetText(const CString& str, const CBCGPBrush& brText, BOOL bRedraw)
{
	m_strText = str;
	CreateResources(brText, TRUE);

	SetDirty();

	if (bRedraw)
	{
		Redraw();
	}
}
//*******************************************************************************
void CBCGPTextGaugeImpl::CreateResources(const CBCGPBrush& brush, BOOL bAllowEmptyBrush)
{
	if (!brush.IsEmpty() || bAllowEmptyBrush)
	{
		m_brTextLight = m_brText = brush;

		if (!brush.IsEmpty())
		{
			COLORREF clr = m_brTextLight.GetColor();

			double H;
			double S;
			double L;

			CBCGPDrawManager::RGBtoHSL (clr, &H, &S, &L);

			m_brTextLight.MakeLighter(S < .99 ? .3 : .6);
		}
	}

	if (m_textFormat.GetFontFamily().IsEmpty())
	{
		LOGFONT lf;
		globalData.fontBold.GetLogFont(&lf);

		m_textFormat.CreateFromLogFont(lf);

		m_textFormat.SetTextAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_CENTER);
		m_textFormat.SetTextVerticalAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_CENTER);
	}
}
//*******************************************************************************
CBCGPSize CBCGPTextGaugeImpl::GetDefaultSize(CBCGPGraphicsManager* pGM, const CBCGPBaseVisualObject* /*pParentGauge*/)
{
	ASSERT_VALID(pGM);

	CreateResources(CBCGPBrush(), FALSE);
	return pGM->GetTextSize(m_strText, m_textFormat);
}
//*******************************************************************************
void CBCGPTextGaugeImpl::SetTextFormat(const CBCGPTextFormat& textFormat)
{
	m_textFormat = textFormat;
	SetDirty();
}
//*******************************************************************************
void CBCGPTextGaugeImpl::SetTextColor(const CBCGPColor& color)
{
	CreateResources(CBCGPBrush(color), FALSE);
	SetDirty();
}
//*******************************************************************************
void CBCGPTextGaugeImpl::SetTextBrush(const CBCGPBrush& brush)
{
	CreateResources(brush, TRUE);
	SetDirty();
}
//*******************************************************************************
void CBCGPTextGaugeImpl::OnScaleRatioChanged(const CBCGPSize& sizeScaleRatioOld)
{
	CBCGPStaticGaugeImpl::OnScaleRatioChanged(sizeScaleRatioOld);

	m_textFormat.Scale(GetScaleRatioMid());
	SetDirty();
}
//*******************************************************************************
void CBCGPTextGaugeImpl::CopyFrom(const CBCGPBaseVisualObject& srcObj)
{
	CBCGPStaticGaugeImpl::CopyFrom(srcObj);

	const CBCGPTextGaugeImpl& src = (const CBCGPTextGaugeImpl&)srcObj;

	m_strText = src.m_strText;
	m_brText = src.m_brText;
	m_brTextLight = src.m_brTextLight;
	m_textFormat = src.m_textFormat;
}
//*******************************************************************************
CBCGPSize CBCGPTextGaugeImpl::GetContentTotalSize(CBCGPGraphicsManager* pGM)
{
	ASSERT_VALID(pGM);

	CBCGPSize size(0.0, 0.0);
	double dblWidth = 0.0;

	if (m_textFormat.IsWordWrap())
	{
		dblWidth = m_rect.Width();
		if (dblWidth <= 0.0)
		{
			return size;
		}
	}

	CreateResources(CBCGPBrush(), FALSE);

	if (pGM->GetType() == CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER_GDI)
	{
		CClientDC dc(NULL);
		pGM->BindDC(&dc, FALSE);

		size = pGM->GetTextSize(m_strText, m_textFormat, dblWidth);

		pGM->BindDC(NULL);
	}
	else
	{
		size = pGM->GetTextSize(m_strText, m_textFormat, dblWidth);
	}

	return size;
}
//*******************************************************************************
void CBCGPTextGaugeImpl::ResumeScrolling(const CString& strAddedText, BOOL bCleanOldInvisible)
{
	if (strAddedText.IsEmpty())
	{
		return;
	}

	int nOldTextLen = m_strText.GetLength();

	CBCGPGraphicsManager* pGM = NULL;

	CBCGPBaseVisualCtrl* pWnd = DYNAMIC_DOWNCAST(CBCGPBaseVisualCtrl, GetParentWnd());
	if (pWnd->GetSafeHwnd() != NULL)
	{
		pGM = pWnd->GetGraphicsManager();
	}
	
	if (bCleanOldInvisible && m_dblScrollOffset > 0.0 && pGM != NULL && m_bIsHorizontalScroll)
	{
		CreateResources(CBCGPBrush(), FALSE);

		double dblWidth = m_textFormat.IsWordWrap() ? m_rect.Width() : 0.0;

		if (!m_textFormat.IsWordWrap() || dblWidth > 0.0)
		{
			for (int i = 0; i < m_strText.GetLength(); i++)
			{
				CString strLeft = m_strText.Left(i + 1);
				
				double cx = 0.0;

				if (pGM->GetType() == CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER_GDI)
				{
					CClientDC dc(NULL);
					pGM->BindDC(&dc, FALSE);
					
					cx = pGM->GetTextSize(strLeft, m_textFormat, dblWidth).cx;
					
					pGM->BindDC(NULL);
				}
				else
				{
					cx = pGM->GetTextSize(strLeft, m_textFormat, dblWidth).cx;
				}

				if (cx >= m_dblScrollOffset)
				{
					m_dblScrollOffset -= cx;
					m_strText.Delete(0, i + 1);
					break;
				}
			}
		}
	}

	m_strText += strAddedText;

	if (nOldTextLen != 0)
	{
		// Adjust scroll time:
		int nNewTextLen = m_strText.GetLength();

		m_dblScrollTime = m_dblScrollTime * nNewTextLen / nOldTextLen;
	}
	
	m_bScrollInternal = TRUE;

	StartContentScrolling(m_dblScrollTime, m_dblScrollDelay, m_bIsHorizontalScroll, m_nScrollFlags,
		m_ScrollAnimationType, &m_ScrollAnimationOptions);
}
