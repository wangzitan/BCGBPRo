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
// BCGPRibbonBackstageViewPanel.cpp: implementation of the CBCGPRibbonBackstageViewPanel class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "bcgcbpro.h"
#include "BCGPRibbonBackstageViewPanel.h"
#include "BCGPRibbonBackstageView.h"
#include "BCGPRibbonButton.h"
#include "BCGPRibbonBar.h"
#include "PropertySheetCtrl.h"
#include "BCGPPropertyPage.h"
#include "BCGPRibbonBackstagePagePrint.h"
#include "BCGPRibbonBackstagePageRecent.h"
#include "BCGPLocalResource.h"
#include "BCGPWorkspace.h"
#include "BCGPFrameWnd.h"
#include "BCGPMDIFrameWnd.h"

extern CBCGPWorkspace*	g_pWorkspace;

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#ifndef BCGP_EXCLUDE_RIBBON

/////////////////////////////////////////////////////////////////////////
// CBCGPRibbonBackstageViewItemForm

IMPLEMENT_DYNCREATE(CBCGPRibbonBackstageViewItemForm, CBCGPBaseRibbonElement)

CBCGPRibbonBackstageViewItemForm::CBCGPRibbonBackstageViewItemForm()
{
	CommonConstruct();
}
//**************************************************************************
CBCGPRibbonBackstageViewItemForm::CBCGPRibbonBackstageViewItemForm(UINT nDlgTemplateID, CRuntimeClass* pDlgClass, CSize sizeWaterMark)
{
	ASSERT(pDlgClass != NULL);
	ASSERT(pDlgClass->IsDerivedFrom(RUNTIME_CLASS(CDialog)));

	CommonConstruct();

	m_nDlgTemplateID = nDlgTemplateID;
	m_pDlgClass = pDlgClass;
	m_sizeWaterMark = sizeWaterMark;
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemForm::CommonConstruct()
{
	m_pWndForm = NULL;
	m_nDlgTemplateID = 0;
	m_pDlgClass = NULL;
	m_sizeDlg = CSize(0, 0);
	m_sizeWaterMark = CSize(0, 0);
	m_clrWatermarkBaseColor = (COLORREF)-1;
	m_bImageMirror = FALSE;
	m_nRecentFlags = 0xFFFF;
	m_nPrintMaxZoomLevel = 100;
	m_rectCaption.SetRectEmpty();
}
//**************************************************************************
CBCGPRibbonBackstageViewItemForm::~CBCGPRibbonBackstageViewItemForm()
{
	if (m_pWndForm != NULL)
	{
		m_pWndForm->DestroyWindow ();
		delete m_pWndForm;
		m_pWndForm = NULL;
	}
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemForm::SetWaterMarkImage(UINT uiWaterMarkResID, COLORREF clrBase)
{
	if (m_Watermark.GetImageWell() != NULL)
	{
		m_Watermark.Clear();
	}

	if (m_WatermarkColorized.GetImageWell() != NULL)
	{
		m_WatermarkColorized.Clear();
	}

	m_clrWatermarkBaseColor = clrBase;
	m_bImageMirror = FALSE;

	if (uiWaterMarkResID != 0)
	{
		m_Watermark.Load(uiWaterMarkResID);
		m_Watermark.SetSingleImage(FALSE);

		m_sizeWaterMark = m_Watermark.GetImageSize();

		OnChangeVisualManager();
	}
	else
	{
		m_sizeWaterMark = CSize(0, 0);
	}
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemForm::OnDraw (CDC* pDC)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);

	if (m_rectCaption.IsRectEmpty())
	{
		return;
	}

	CBCGPVisualManager::GetInstance()->OnFillRibbonBackstageForm(pDC, m_pWndForm, m_rectCaption);

	CFont* pOldFont = pDC->SelectObject(&globalData.fontHeader);
	pDC->SetTextColor(CBCGPVisualManager::GetInstance()->GetRibbonBackstageTextColor());
	pDC->SetBkMode(TRANSPARENT);
	
	CRect rectText = m_rectCaption;
	rectText.DeflateRect(globalUtils.ScaleByDPI(10), 0);

	CString strText = m_strTransitionText.IsEmpty() ? m_strText : m_strTransitionText;

	pDC->DrawText(strText, rectText, DT_SINGLELINE | DT_END_ELLIPSIS | DT_VCENTER);
	pDC->SelectObject(pOldFont);
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemForm::OnAfterChangeRect (CDC* pDC)
{
	SetLayoutReady(FALSE);

	CBCGPBaseRibbonElement::OnAfterChangeRect(pDC);

	m_rectCaption.SetRectEmpty();

	if (m_rect.IsRectEmpty ())
	{
		if (m_pWndForm->GetSafeHwnd () != NULL)
		{
			m_pWndForm->ShowWindow (SW_HIDE);
		}

		m_strTransitionText.Empty();
		return;
	}

	BOOL bForceAdjustLocations = FALSE;

	if (m_pWndForm == NULL)
	{
		if ((m_pWndForm = OnCreateFormWnd()) == NULL)
		{
			return;
		}

		SetLayoutReady(FALSE);
		bForceAdjustLocations = TRUE;

		if (m_Watermark.GetImageWell() != NULL)
		{
			BOOL bMirror = FALSE;

			if (GetParentWnd()->GetExStyle () & WS_EX_LAYOUTRTL)
			{
				if (!m_bImageMirror)
				{
					m_bImageMirror = TRUE;
					bMirror = TRUE;
				}
			}
			else if (m_bImageMirror)
			{
				m_bImageMirror = FALSE;
				bMirror = TRUE;
			}

			if (bMirror)
			{
				m_Watermark.Mirror();
				OnChangeVisualManager();
			}
		}

		CBCGPDialog* pDialog = DYNAMIC_DOWNCAST(CBCGPDialog, m_pWndForm);
		if (pDialog != NULL)
		{
			pDialog->m_Impl.EnableGesturePan();
		}
	}

	CRect rectWindow = m_rect;
	
	CBCGPRibbonBackstageViewPanel* pParentPanel = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewPanel, GetParentPanel());
	if (pParentPanel != NULL && pDC != NULL)
	{
		ASSERT_VALID(pDC);
		ASSERT_VALID(pParentPanel);

		BOOL bReserveSpaceForCaption = FALSE;

		if (pParentPanel->m_PageTransitionEffect == CBCGPPageTransitionManager::BCGPPageTransitionNone)
		{
			bReserveSpaceForCaption = !m_strText.IsEmpty();
		}
		else
		{
			bReserveSpaceForCaption = pParentPanel->m_bHasCaptions;
		}

		if (bReserveSpaceForCaption)
		{
			m_rectCaption = m_rect;
			m_rectCaption.bottom = m_rectCaption.top + GetCaptionHeight();
			
			rectWindow.top = m_rectCaption.bottom;
		}
	}

	CRect rectWatermark(0, 0, 0, 0);

	if (m_sizeWaterMark != CSize(0, 0))
	{
		int xWatermark = max(rectWindow.Width(), m_sizeDlg.cx) - m_sizeWaterMark.cx;
		int yWatermark = max(rectWindow.Height(), m_sizeDlg.cy) - m_sizeWaterMark.cy;

		rectWatermark = CRect(CPoint(xWatermark, yWatermark), m_sizeWaterMark);
	}

	OnSetBackstageWatermarkRect(rectWatermark);

	m_pWndForm->SetWindowPos (NULL, 
		rectWindow.left, rectWindow.top,
		rectWindow.Width (), rectWindow.Height () + 1,
		SWP_NOZORDER | SWP_NOACTIVATE);

	if (!m_pWndForm->IsWindowVisible())
	{
		m_pWndForm->ShowWindow (SW_SHOWNOACTIVATE);
	}

	SetLayoutReady();

	if (bForceAdjustLocations)
	{
		CBCGPRibbonPanel* pPanel = GetParentPanel();
		if (pPanel != NULL && pPanel->GetParentMenuBar()->GetSafeHwnd() != NULL)
		{
			CRect rectPanel;
			pPanel->GetParentMenuBar()->GetClientRect(rectPanel);

			pPanel->GetParentMenuBar()->PostMessage(WM_SIZE, MAKEWPARAM(rectPanel.Width(), rectPanel.Height()));

			pPanel->GetParentMenuBar()->RedrawWindow(m_rectCaption);
		}
	}

	m_pWndForm->RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemForm::OnSetBackstageWatermarkRect(CRect rectWatermark)
{
	ASSERT_VALID (this);

	CBCGPDialog* pDialog = DYNAMIC_DOWNCAST(CBCGPDialog, m_pWndForm);
	if (pDialog != NULL)
	{
		pDialog->m_rectBackstageWatermark = rectWatermark;
	}
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemForm::CopyFrom (const CBCGPBaseRibbonElement& s)
{
	ASSERT_VALID (this);

	CBCGPBaseRibbonElement::CopyFrom (s);

	CBCGPRibbonBackstageViewItemForm& src = (CBCGPRibbonBackstageViewItemForm&) s;

	m_nDlgTemplateID = src.m_nDlgTemplateID;
	m_pDlgClass = src.m_pDlgClass;
	m_sizeWaterMark = src.m_sizeWaterMark;
	src.m_Watermark.CopyTo (m_Watermark);
	src.m_WatermarkColorized.CopyTo(m_WatermarkColorized);
	m_clrWatermarkBaseColor = src.m_clrWatermarkBaseColor;
	m_bImageMirror = src.m_bImageMirror;
	m_nRecentFlags = src.m_nRecentFlags;
	m_nPrintMaxZoomLevel = src.m_nPrintMaxZoomLevel;
}
//**************************************************************************
CWnd* CBCGPRibbonBackstageViewItemForm::OnCreateFormWnd()
{
	ASSERT_VALID (this);
	ASSERT(m_pDlgClass != NULL);

	CBCGPDialog* pForm = NULL;

	if (m_pDlgClass == RUNTIME_CLASS(CBCGPDialog))
	{
		pForm = new CBCGPDialog;
	}
	else
	{
		pForm = DYNAMIC_DOWNCAST(CBCGPDialog, m_pDlgClass->CreateObject());
	}

	if (pForm == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}

	CBCGPRibbonBackstagePageRecent* pRecentFrom = DYNAMIC_DOWNCAST(CBCGPRibbonBackstagePageRecent, pForm);
	if (pRecentFrom != NULL)
	{
		pRecentFrom->SetFlags(m_nRecentFlags);
	}

	CBCGPRibbonBackstagePagePrint* pPrintFrom = DYNAMIC_DOWNCAST(CBCGPRibbonBackstagePagePrint, pForm);
	if (pPrintFrom != NULL)
	{
		pPrintFrom->SetMaxZoomLevel(m_nPrintMaxZoomLevel);
	}

	if (!pForm->Create(m_nDlgTemplateID, GetParentWnd()))
	{
		ASSERT(FALSE);
		delete pForm;
		return NULL;
	}

	pForm->EnableVisualManagerStyle();
	pForm->EnableBackstageMode();
	pForm->m_pBackstageWatermark = m_WatermarkColorized.GetImageWell() != NULL ? &m_WatermarkColorized : &m_Watermark;

	CRect rectClient;
	pForm->GetWindowRect(&rectClient);

	m_sizeDlg = rectClient.Size();
	m_sizeDlg.cy += GetCaptionHeight();

	return pForm;
}
//**************************************************************************
int CBCGPRibbonBackstageViewItemForm::GetCaptionHeight()
{
	CClientDC dc(NULL);
	CFont* pOldFont = dc.SelectObject(&globalData.fontHeader);
	
	TEXTMETRIC tm;
	dc.GetTextMetrics (&tm);
	
	int nHeight = tm.tmHeight + globalUtils.ScaleByDPI(10);
	
	dc.SelectObject(pOldFont);
	return nHeight;
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemForm::OnChangeVisualManager()
{
	ASSERT_VALID(this);

	if (m_clrWatermarkBaseColor == (COLORREF)-1 || m_Watermark.GetImageWell() == NULL)
	{
		return;
	}

	if (m_WatermarkColorized.GetImageWell() != NULL)
	{
		m_WatermarkColorized.Clear();
	}

	COLORREF clrMainButton = CBCGPVisualManager::GetInstance()->GetMainButtonColor();

	if (clrMainButton != (COLORREF)-1)
	{
		m_Watermark.CopyTo(m_WatermarkColorized);
		m_WatermarkColorized.AdaptColors(m_clrWatermarkBaseColor, clrMainButton);
	}
}
//**************************************************************************
BOOL CBCGPRibbonBackstageViewItemForm::SetACCData(CWnd* pParent, CBCGPAccessibilityData& data)
{
	ASSERT_VALID (this);
	
	data.Clear ();
	
	data.m_strAccName = m_strText.IsEmpty() ? _T("Ribbon Backstage View") : m_strText;
	data.m_nAccRole = ROLE_SYSTEM_PANE;
	data.m_bAccState = STATE_SYSTEM_READONLY;
	
	data.m_rectAccLocation = m_rect;
	pParent->ClientToScreen (&data.m_rectAccLocation);
	
	return TRUE;
}

class CBCGPBackstagePropertySheetCtrl : public CBCGPPropertySheetCtrl
{
protected:
	virtual void OnDrawPageHeader(CDC* pDC, int nPage, CRect rectHeader)
	{
		ASSERT_VALID(pDC);

		CPropertyPage* pPage = GetPage(nPage);
		if (pPage->GetSafeHwnd() != NULL)
		{
			CString strName;
			pPage->GetWindowText(strName);

			CFont* pOldFont = pDC->SelectObject(&globalData.fontCaption);
			ASSERT_VALID(pOldFont);

			COLORREF clrTextOld = pDC->SetTextColor(CBCGPVisualManager::GetInstance()->GetRibbonBackstageInfoTextColor());
			pDC->SetBkMode(TRANSPARENT);

			int xMargin = 0;
			if (!m_PageXMargins.Lookup(pPage->GetSafeHwnd(), xMargin))
			{
				xMargin = globalUtils.ScaleByDPI(20);

				for (CWnd* pWnd = pPage->GetWindow(GW_CHILD); pWnd != NULL; pWnd = pWnd->GetWindow(GW_HWNDNEXT))
				{
					if (pWnd->GetStyle() & WS_VISIBLE)
					{
						CRect rectCtrl;
						pWnd->GetWindowRect(rectCtrl);

						pPage->ScreenToClient(rectCtrl);

						xMargin = min(xMargin, rectCtrl.left);
					}
				}

				m_PageXMargins.SetAt(pPage->GetSafeHwnd(), xMargin);
			}

			rectHeader.left += xMargin;
			rectHeader.OffsetRect(0, 1);

			int nTextHeight = pDC->DrawText(strName, rectHeader, DT_SINGLELINE | DT_VCENTER | DT_NOCLIP | DT_NOPREFIX);

			pDC->SetTextColor(clrTextOld);
			pDC->SelectObject(pOldFont);

			if (!globalData.IsHighContastMode () && CBCGPVisualManager::GetInstance ()->IsUnderlineListBoxCaption(&m_wndList))
			{
				CRect rectSeparator = rectHeader;
				rectSeparator.top = rectHeader.CenterPoint().y + nTextHeight / 2;
				rectSeparator.bottom = rectSeparator.top + 1;

				CBCGPStatic ctrl;
				ctrl.m_bBackstageMode = TRUE;

				if (CBCGPVisualManager::GetInstance ()->IsOwnerDrawDlgSeparator(&ctrl))
				{
					CBCGPVisualManager::GetInstance ()->OnDrawDlgSeparator(pDC, &ctrl, rectSeparator, TRUE);
				}
				else
				{
					CPen pen (PS_SOLID, 1, globalData.clrBtnShadow);
					CPen* pOldPen = (CPen*)pDC->SelectObject (&pen);
					
					pDC->MoveTo(rectSeparator.left, rectSeparator.top);
					pDC->LineTo(rectSeparator.right, rectSeparator.top);
					
					pDC->SelectObject(pOldPen);
				}
			}
		}
	}

	CMap<HWND, HWND, int, int&>	m_PageXMargins;
};

/////////////////////////////////////////////////////////////////////////
// CBCGPRibbonBackstageViewItemPropertySheet

IMPLEMENT_DYNCREATE(CBCGPRibbonBackstageViewItemPropertySheet, CBCGPRibbonBackstageViewItemForm)

CBCGPRibbonBackstageViewItemPropertySheet::CBCGPRibbonBackstageViewItemPropertySheet()
{
	m_nIconsListResID = 0;
	m_bIconsAutoScale = FALSE;
	m_nIconWidth = 0;
	m_PageTransitionEffect = CBCGPPageTransitionManager::BCGPPageTransitionNone;
	m_nPageTransitionTime = 300;
	m_bUseDefaultPageTransitionEffect = TRUE;
	m_bDefaultPageHeader = FALSE;
}
//**************************************************************************
CBCGPRibbonBackstageViewItemPropertySheet::CBCGPRibbonBackstageViewItemPropertySheet(UINT nIconsListResID, int nIconWidth, BOOL bIconsAutoScale, BOOL bDefaultPageHeader)
{
	m_nIconsListResID = nIconsListResID;
	m_nIconWidth = nIconWidth;
	m_bIconsAutoScale = bIconsAutoScale;
	m_PageTransitionEffect = CBCGPPageTransitionManager::BCGPPageTransitionNone;
	m_nPageTransitionTime = 300;
	m_bUseDefaultPageTransitionEffect = TRUE;
	m_bDefaultPageHeader = bDefaultPageHeader;
}
//**************************************************************************
CBCGPRibbonBackstageViewItemPropertySheet::~CBCGPRibbonBackstageViewItemPropertySheet()
{
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemPropertySheet::AddPage(CBCGPPropertyPage* pPage)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pPage);
	ASSERT(m_pWndForm->GetSafeHwnd() == NULL);

	m_arPages.Add(pPage);
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemPropertySheet::AddGroup(LPCTSTR lpszCaption)
{
	ASSERT_VALID(this);
	ASSERT(lpszCaption != NULL);

	m_arPages.Add(NULL);
	m_arCaptions.Add(lpszCaption);
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemPropertySheet::RemoveAll()
{
	m_arPages.RemoveAll();
	m_arCaptions.RemoveAll();
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemPropertySheet::CopyFrom (const CBCGPBaseRibbonElement& s)
{
	ASSERT_VALID (this);

	CBCGPRibbonBackstageViewItemForm::CopyFrom (s);

	CBCGPRibbonBackstageViewItemPropertySheet& src = (CBCGPRibbonBackstageViewItemPropertySheet&) s;

	m_arPages.Copy(src.m_arPages);
	m_arCaptions.Copy(src.m_arCaptions);
	m_nIconsListResID = src.m_nIconsListResID;
	m_nIconWidth = src.m_nIconWidth;
	m_bIconsAutoScale = src.m_bIconsAutoScale;
	m_PageTransitionEffect = src.m_PageTransitionEffect;
	m_bUseDefaultPageTransitionEffect = src.m_bUseDefaultPageTransitionEffect;
	m_nPageTransitionTime = src.m_nPageTransitionTime;
	m_bDefaultPageHeader = src.m_bDefaultPageHeader;
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemPropertySheet::OnSetBackstageWatermarkRect(CRect rectWatermark)
{
	ASSERT_VALID (this);

	if (m_pWndForm->GetSafeHwnd() == NULL)
	{
		return;
	}

	CRect rectWatermarkePage(0, 0, 0, 0);

	for (int i = 0; i < (int)m_arPages.GetSize(); i++)
	{
		CBCGPPropertyPage* pPage = m_arPages[i];
		if (pPage != NULL)
		{
			ASSERT_VALID(pPage);

			if (pPage->GetSafeHwnd() != NULL && rectWatermarkePage.IsRectEmpty())
			{
				rectWatermarkePage = rectWatermark;
				m_pWndForm->MapWindowPoints(pPage, &rectWatermarkePage); 
			}
				
			pPage->m_rectBackstageWatermark = rectWatermarkePage;
		}
	}
}
//**************************************************************************
void CBCGPRibbonBackstageViewItemPropertySheet::SetLayoutReady(BOOL bReady)
{
	ASSERT_VALID (this);

	CBCGPPropertySheetCtrl* pPropSheet = DYNAMIC_DOWNCAST(CBCGPPropertySheetCtrl, m_pWndForm);
	if (pPropSheet == NULL)
	{
		return;
	}

	ASSERT_VALID(pPropSheet);

	pPropSheet->m_bIsReady = bReady;

	if (m_bUseDefaultPageTransitionEffect && GetTopLevelRibbonBar() != NULL)
	{
		pPropSheet->EnablePageTransitionEffect(
			GetTopLevelRibbonBar()->GetBackstagePageTransitionEffect(),
			GetTopLevelRibbonBar()->GetBackstagePageTransitionTime());
	}
	else
	{
		pPropSheet->EnablePageTransitionEffect(m_PageTransitionEffect, m_nPageTransitionTime);
	}

	if (bReady)
	{
		pPropSheet->AdjustControlsLayout();
	
		CBCGPPropertyPage* pPage = DYNAMIC_DOWNCAST(CBCGPPropertyPage, pPropSheet->GetActivePage());
		if (pPage != NULL)
		{
			pPage->AdjustControlsLayout();
		}
	}
}
//**************************************************************************
int CBCGPRibbonBackstageViewItemPropertySheet::GetHeaderHeight()
{
	if (!m_bDefaultPageHeader)
	{
		return 0;
	}

	CBCGPPropertySheetCtrl* pPropSheet = DYNAMIC_DOWNCAST(CBCGPPropertySheetCtrl, m_pWndForm);
	if (pPropSheet->GetSafeHwnd() != NULL)
	{
		return pPropSheet->GetHeaderHeight();
	}

	return 0;
}
//**************************************************************************
CWnd* CBCGPRibbonBackstageViewItemPropertySheet::OnCreateFormWnd()
{
	ASSERT_VALID (this);

	CBCGPBackstagePropertySheetCtrl* pForm = new CBCGPBackstagePropertySheetCtrl;
	if (pForm == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}

	pForm->EnableVisualManagerStyle();
	pForm->EnableLayout();
	pForm->SetLook(CBCGPPropertySheet::PropSheetLook_List, -1);

	if (m_bDefaultPageHeader)
	{
		pForm->m_bSyncHeaderHeightWithListItemHeight = TRUE;
	}

	if (m_nIconsListResID != 0)
	{
		pForm->SetIconsList (m_nIconsListResID, m_nIconWidth, m_bIconsAutoScale);
	}

	pForm->m_bBackstageMode = TRUE;
	pForm->m_bIsAutoDestroy = FALSE;

	int nCaptionIndex = 0;

	for (int i = 0; i < (int)m_arPages.GetSize(); i++)
	{
		CBCGPPropertyPage* pPage = m_arPages[i];
		if (pPage == NULL)
		{
			pForm->AddGroup(m_arCaptions[nCaptionIndex++]);
		}
		else
		{
			ASSERT_VALID(pPage);

			pPage->m_pBackstageWatermark = m_WatermarkColorized.GetImageWell() != NULL ? &m_WatermarkColorized : &m_Watermark;
			pForm->AddPage(pPage);
		}
	}

	CWnd* pParent = GetParentWnd();
	pParent->LockWindowUpdate();

	BOOL bRes = pForm->Create(pParent, WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0);

	pParent->UnlockWindowUpdate();

	if (!bRes)
	{
		ASSERT(FALSE);
		delete pForm;
		return NULL;
	}

	CRect rectClient;
	pForm->GetWindowRect(&rectClient);

	m_sizeDlg = rectClient.Size();
	m_sizeDlg.cx += pForm->GetNavBarWidth();
	m_sizeDlg.cy += GetCaptionHeight();

	return pForm;
}

//////////////////////////////////////////////////////////////////////
// CBCGPRibbonBackstageViewPanel

IMPLEMENT_DYNCREATE(CBCGPRibbonBackstageViewPanel, CBCGPRibbonMainPanel)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPRibbonBackstageViewPanel::CBCGPRibbonBackstageViewPanel()
{
	m_bIsStartPage = FALSE;
	m_pSelected = NULL;
	m_pNewSelected = NULL;
	m_bSelectedByMouseClick = FALSE;
	m_rectRight.SetRectEmpty();
	m_sizeRightView = CSize(0, 0);
	m_bInAdjustScrollBars = FALSE;
	m_bHasCaptions = FALSE;
}
//********************************************************************************
CBCGPRibbonBackstageViewPanel::~CBCGPRibbonBackstageViewPanel()
{
}
//********************************************************************************
void CBCGPRibbonBackstageViewPanel::Repos (CDC* pDC, const CRect& rect)
{
	CBCGPRibbonMainPanel::Repos(pDC, rect);

	if (m_pSelected == NULL && !m_bIsCalcWidth)
	{
		CBCGPRibbonPanelMenuBar* pMenuBar = DYNAMIC_DOWNCAST(CBCGPRibbonPanelMenuBar, GetParentWnd());
		if (pMenuBar != NULL)
		{
			CFrameWnd* pTarget = (CFrameWnd*) pMenuBar->GetCommandTarget ();
			if (pTarget == NULL || !pTarget->IsFrameWnd())
			{
				pTarget = BCGPGetParentFrame(pMenuBar);
			}
			
			if (pTarget != NULL)
			{
				pMenuBar->OnUpdateCmdUI(pTarget, TRUE);
			}
		}

		if (m_pMainButton != NULL)
		{
			ASSERT_VALID(m_pMainButton);
			
			if (m_pMainButton->GetParentRibbonBar () != NULL)
			{
				ASSERT_VALID (m_pMainButton->GetParentRibbonBar ());

				int nInitialPage = m_pMainButton->GetParentRibbonBar ()->GetInitialBackstagePage();
				int nDefaultPage = m_pMainButton->GetParentRibbonBar ()->GetDefaultBackstagePage();

				int nStartPage = (nInitialPage == -1) ? nDefaultPage : nInitialPage;
				if (nStartPage >= 0)
				{
					int nPage = 0;

					for (int i = 0; i < m_arElements.GetSize (); i++)
					{
						CBCGPBaseRibbonElement* pElem = m_arElements [i];
						ASSERT_VALID (pElem);

						if (pElem->GetBackstageAttachedView() != NULL)
						{
							if (nPage == nStartPage)
							{
								if (!pElem->IsDisabled())
								{
									m_pSelected = pElem;
									m_pSelected->m_bIsChecked = TRUE;
								}

								break;
							}

							nPage++;
						}
					}
				}
			}
		}

		if (m_pSelected == NULL)
		{
			for (int i = 0; i < m_arElements.GetSize (); i++)
			{
				CBCGPBaseRibbonElement* pElem = m_arElements [i];
				ASSERT_VALID (pElem);

				if (!pElem->IsDisabled() && pElem->GetBackstageAttachedView() != NULL && !pElem->IsHiddenInAppMode())
				{
					m_pSelected = pElem;
					m_pSelected->m_bIsChecked = TRUE;
					break;
				}
			}
		}

		if (m_pHighlighted == NULL)
		{
			m_pHighlighted = m_pSelected;
		}
	}

	if (m_bIsStartPage)
	{
		m_rectMenuElements.right = m_rectMenuElements.left;
	}

	m_rectRight = rect;
	m_rectRight.left = m_rectMenuElements.right;
	m_rectRight.top -= m_nScrollOffset;

	if (m_bIsCalcWidth)
	{
		return;
	}

	AdjustScrollBars();
	ReposActiveForm();
}
//********************************************************************************
CBCGPBaseRibbonElement* CBCGPRibbonBackstageViewPanel::MouseButtonDown (CPoint point)
{
	CBCGPBaseRibbonElement* pHit = CBCGPRibbonMainPanel::MouseButtonDown (point);
	if (pHit == NULL || pHit->IsDisabled())
	{
		return NULL;
	}

	if (!m_rectPageTransition.IsRectEmpty())
	{
		StopPageTransition();
		OnPageTransitionFinished();

		if (pHit->GetRect().PtInRect(point))
		{
			SelectView(pHit);
			return pHit;
		}

		return NULL;
	}

	m_bSelectedByMouseClick = TRUE;
	SelectView(pHit);

	return pHit;
}
//********************************************************************************
CBCGPRibbonButton* CBCGPRibbonBackstageViewPanel::AddPrintPreview(UINT uiCommandID, LPCTSTR lpszLabel,
																  UINT uiWaterMarkResID, COLORREF clrBase,
																  int nPrintMaxZoomLevel)
{
	ASSERT_VALID(this);
	ASSERT(lpszLabel != NULL);

	CBCGPRibbonBackstageViewItemForm* pFormPrint = NULL;

	{
		CBCGPLocalResource locaRes;
		pFormPrint = new CBCGPRibbonBackstageViewItemForm (IDD_BCGBARRES_PRINT_FORM, RUNTIME_CLASS(CBCGPRibbonBackstagePagePrint));
	}

	ASSERT_VALID(pFormPrint);

	pFormPrint->m_nPrintMaxZoomLevel = nPrintMaxZoomLevel;
	pFormPrint->SetWaterMarkImage(uiWaterMarkResID, clrBase);

	return AddView (uiCommandID, lpszLabel, pFormPrint);
}
//********************************************************************************
CBCGPRibbonButton* CBCGPRibbonBackstageViewPanel::AddRecentView(UINT uiCommandID, LPCTSTR lpszLabel, 
																UINT nFlags, UINT uiWaterMarkResID, COLORREF clrBase)
{
	ASSERT_VALID(this);
	ASSERT(lpszLabel != NULL);

	CBCGPRibbonBackstageViewItemForm* pFormRecent = NULL;

	{
		CBCGPLocalResource locaRes;
		pFormRecent = new CBCGPRibbonBackstageViewItemForm(IDD_BCGBARRES_RECENT_FORM, RUNTIME_CLASS(CBCGPRibbonBackstagePageRecent));
	}

	ASSERT_VALID(pFormRecent);

	pFormRecent->m_nRecentFlags = nFlags;
	pFormRecent->SetWaterMarkImage(uiWaterMarkResID, clrBase);

	return AddView (uiCommandID, lpszLabel, pFormRecent);
}
//********************************************************************************
BOOL CBCGPRibbonBackstageViewPanel::AttachViewToItem(UINT uiID, CBCGPRibbonBackstageViewItemForm* pView, BOOL bByCommand/* = TRUE*/)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pView);

	CBCGPRibbonButton* pItem = NULL;

	if (bByCommand)
	{
		for (int i = 0; i < m_arElements.GetSize (); i++)
		{
			CBCGPBaseRibbonElement* pElem = m_arElements [i];
			ASSERT_VALID (pElem);

			if (pElem->GetID() == uiID)
			{
				pItem = DYNAMIC_DOWNCAST(CBCGPRibbonButton, pElem);
				break;
			}
		}
	}
	else
	{
		if (uiID < (UINT)m_arElements.GetSize ())
		{
			pItem = DYNAMIC_DOWNCAST(CBCGPRibbonButton, m_arElements [uiID]);
		}
	}

	if (pItem == NULL)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	if (pItem->GetSubItems().GetSize() > 0)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	pItem->AddSubItem(pView);

	return TRUE;
}
//********************************************************************************
BOOL CBCGPRibbonBackstageViewPanel::AttachPrintPreviewToItem(UINT uiID, BOOL bByCommand,
															 UINT uiWaterMarkResID, COLORREF clrBase,
															 int nPrintMaxZoomLevel)
{
	CBCGPRibbonBackstageViewItemForm* pFormPrint = NULL;

	{
		CBCGPLocalResource locaRes;
		pFormPrint = new CBCGPRibbonBackstageViewItemForm (IDD_BCGBARRES_PRINT_FORM, RUNTIME_CLASS(CBCGPRibbonBackstagePagePrint));
	}

	ASSERT_VALID(pFormPrint);

	pFormPrint->SetWaterMarkImage(uiWaterMarkResID, clrBase);
	pFormPrint->m_nPrintMaxZoomLevel = nPrintMaxZoomLevel;

	return AttachViewToItem(uiID, pFormPrint, bByCommand);
}
//********************************************************************************
BOOL CBCGPRibbonBackstageViewPanel::AttachRecentViewToItem(UINT uiID, UINT nFlags, BOOL bByCommand,
															 UINT uiWaterMarkResID, COLORREF clrBase)
{
	CBCGPRibbonBackstageViewItemForm* pFormRecent = NULL;

	{
		CBCGPLocalResource locaRes;
		pFormRecent = new CBCGPRibbonBackstageViewItemForm (IDD_BCGBARRES_RECENT_FORM, RUNTIME_CLASS(CBCGPRibbonBackstagePageRecent));
	}

	ASSERT_VALID(pFormRecent);

	pFormRecent->SetWaterMarkImage(uiWaterMarkResID, clrBase);
	pFormRecent->m_nRecentFlags = nFlags;

	return AttachViewToItem(uiID, pFormRecent, bByCommand);
}
//********************************************************************************
int CBCGPRibbonBackstageViewPanel::GetFormsCount() const
{
	ASSERT_VALID(this);

	int nCount = 0;

	for (int i = 0; i < m_arElements.GetSize (); i++)
	{
		CBCGPRibbonButton* pElem = DYNAMIC_DOWNCAST(CBCGPRibbonButton, m_arElements[i]);

		if (pElem != NULL && pElem->GetSubItems ().GetSize () == 1)
		{
			CBCGPRibbonBackstageViewItemForm* pSubItem = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewItemForm, pElem->GetSubItems ()[0]);
			if (pSubItem != NULL)
			{
				nCount++;
			}
		}
	}

	return nCount;
}
//********************************************************************************
CBCGPRibbonBackstageViewItemForm* CBCGPRibbonBackstageViewPanel::GetForm(int nIndex) const
{
	ASSERT_VALID(this);
	
	int nCount = 0;
	
	for (int i = 0; i < m_arElements.GetSize (); i++)
	{
		CBCGPRibbonButton* pElem = DYNAMIC_DOWNCAST(CBCGPRibbonButton, m_arElements[i]);

		if (pElem != NULL && pElem->GetSubItems ().GetSize () == 1)
		{
			CBCGPRibbonBackstageViewItemForm* pSubItem = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewItemForm, pElem->GetSubItems ()[0]);
			if (pSubItem != NULL)
			{
				if (nIndex == nCount)
				{
					return pSubItem;
				}

				nCount++;
			}
		}
	}
	
	return NULL;
}
//********************************************************************************
void CBCGPRibbonBackstageViewPanel::ReposActiveForm()
{
	ASSERT_VALID(this);

	CBCGPBaseRibbonElement* pSelected = m_pNewSelected != NULL ? m_pNewSelected : m_pSelected;

	if (pSelected != NULL)
	{
		ASSERT_VALID(pSelected);

		CBCGPBaseRibbonElement* pView = pSelected->GetBackstageAttachedView();
		if (pView != NULL)
		{
			ASSERT_VALID(pView);

			if (m_pMainButton != NULL)
			{
				ASSERT_VALID (m_pMainButton);
				ASSERT_VALID (m_pMainButton->GetParentRibbonBar ());

				m_bHasCaptions = m_pMainButton->GetParentRibbonBar()->HasBackstagePageCaptions();

				pView->SetRect(m_rectRight);

				if (m_bHasCaptions)
				{
					if (pView->m_strText.IsEmpty())
					{
						pView->m_strText = pSelected->m_strText;

						if (!pView->m_strText.IsEmpty())
						{
							globalUtils.RemoveSingleAmp(pView->m_strText);
						}
					}
				}
				else
				{
					pView->m_strText.Empty();
				}

				CClientDC dc(m_pMainButton->GetParentRibbonBar());

				CFont* pOldFont = dc.SelectObject (m_pMainButton->GetParentRibbonBar()->GetFont());
				ASSERT (pOldFont != NULL);

				pView->OnAfterChangeRect(&dc);
				m_sizeRightView = pView->GetSize(&dc);

				CBCGPRibbonBackstageViewItemPropertySheet* pPropSheetView = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewItemPropertySheet, pView);
				if (pPropSheetView != NULL)
				{
					ASSERT_VALID(pPropSheetView);
					m_sizeRightView.cy += pPropSheetView->GetHeaderHeight();
				}
				
				if (m_bIsStartPage)
				{
					CBCGPRibbonBackstageViewItemForm* pForm = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewItemForm, pView);
					if (pForm != NULL)
					{
						CBCGPDialog* pDlg = DYNAMIC_DOWNCAST(CBCGPDialog, pForm->m_pWndForm);
						if (pDlg != NULL)
						{
							pDlg->m_bIsRibbonStartPage = TRUE;
							m_pMainButton->GetParentRibbonBar()->m_nStartPageLeftPaneWidth = pDlg->GetRibbonStartPageLeftPaneWidth();
						}
					}
				}

				dc.SelectObject (pOldFont);

				if (m_pNewSelected != NULL)
				{
					RedrawElement(pView);
				}
			}
		}
	}
}
//**********************************************************************************
CBCGPRibbonButton* CBCGPRibbonBackstageViewPanel::FindFormParent (CRuntimeClass* pClass, BOOL bDerived)
{
	if (pClass == NULL)
	{
		return NULL;
	}

	CBCGPRibbonButton* pParent = NULL;
	for (int i = 0; i < m_arElements.GetSize (); i++)
	{
		CBCGPRibbonButton* pButton = DYNAMIC_DOWNCAST(CBCGPRibbonButton, m_arElements[i]);
		if (pButton == NULL || pButton->GetSubItems ().GetSize () != 1)
		{
			continue;
		}

		CBCGPRibbonBackstageViewItemForm* pSubItem = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewItemForm, pButton->GetSubItems ()[0]);
		if (pSubItem != NULL && pSubItem->m_pDlgClass != NULL)
		{
			if ((!bDerived && pSubItem->m_pDlgClass == pClass) ||
				(bDerived && pSubItem->m_pDlgClass->IsDerivedFrom (pClass)))
			{
				pParent = pButton;
				break;
			}
		}
	}

	return pParent;
}
//**********************************************************************************
int CBCGPRibbonBackstageViewPanel::FindFormPageIndex(CRuntimeClass* pClass, BOOL bDerived)
{
	if (pClass == NULL)
	{
		return -1;
	}

	int nIndex = 0;

	for (int i = 0; i < m_arElements.GetSize (); i++)
	{
		CBCGPRibbonButton* pButton = DYNAMIC_DOWNCAST(CBCGPRibbonButton, m_arElements[i]);
		if (pButton == NULL || pButton->GetSubItems ().GetSize () != 1)
		{
			continue;
		}

		CBCGPRibbonBackstageViewItemForm* pSubItem = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewItemForm, pButton->GetSubItems ()[0]);
		if (pSubItem != NULL)
		{
			if (pSubItem->m_pDlgClass != NULL)
			{
				if ((!bDerived && pSubItem->m_pDlgClass == pClass) ||
					(bDerived && pSubItem->m_pDlgClass->IsDerivedFrom (pClass)))
				{
					return nIndex;
				}
			}

			nIndex++;
		}
	}

	return -1;
}
//**********************************************************************************
void CBCGPRibbonBackstageViewPanel::CopyFrom (CBCGPRibbonPanel& s)
{
	CBCGPRibbonMainPanel::CopyFrom(s);

	CBCGPRibbonBackstageViewPanel& src = (CBCGPRibbonBackstageViewPanel&)s;
	
	m_bIsStartPage = src.m_bIsStartPage;
	m_bHasCaptions = src.m_bHasCaptions;
}
//**********************************************************************************
BOOL CBCGPRibbonBackstageViewPanel::OnKey (UINT nChar)
{
	if (!m_rectPageTransition.IsRectEmpty())
	{
		return TRUE;
	}

	if ((nChar == VK_RIGHT || nChar == VK_TAB) && m_pHighlighted != NULL &&
		m_pHighlighted->GetBackstageAttachedView() != NULL)
	{
		CBCGPRibbonBackstageViewItemForm* pForm = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewItemForm, m_pHighlighted->GetBackstageAttachedView());
		if (pForm != NULL && pForm->m_pWndForm->GetSafeHwnd() != NULL)
		{
			pForm->m_pWndForm->SetFocus();
			return FALSE;
		}
	}

	if (!CBCGPRibbonMainPanel::OnKey(nChar))
	{
		return FALSE;
	}

	switch (nChar)
	{
	case VK_HOME:
		GetParentWnd()->SetFocus();
		return TRUE;

	case VK_UP:
	case VK_DOWN:
		break;

	default:
		SelectView(m_pHighlighted);
		break;
	}

	GetParentWnd()->SetFocus();
	return TRUE;
}
//**********************************************************************************
void CBCGPRibbonBackstageViewPanel::SelectView(CBCGPBaseRibbonElement* pElem)
{
	ASSERT_VALID(this);

	m_pNewSelected = NULL;

	if (pElem == NULL)
	{
		return;
	}

	ASSERT_VALID(pElem);

	if (pElem->IsDisabled())
	{
		return;
	}

	BOOL bRedrawCaption = FALSE;

	if (pElem != m_pSelected && pElem->GetBackstageAttachedView() != NULL)
	{
		if (m_pSelected != NULL)
		{
			ASSERT_VALID(m_pSelected);

			if (m_pMainButton->GetParentRibbonBar () != NULL)
			{
				ASSERT_VALID (m_pMainButton->GetParentRibbonBar ());

				m_PageTransitionEffect = m_pMainButton->GetParentRibbonBar ()->GetBackstagePageTransitionEffect();
				m_nPageTransitionTime = m_pMainButton->GetParentRibbonBar ()->GetBackstagePageTransitionTime();
			}

			if (m_bSelectedByMouseClick && m_PageTransitionEffect != BCGPPageTransitionNone && !globalData.IsHighContastMode())
			{
				m_bSelectedByMouseClick = FALSE;

				CBCGPRibbonBackstageViewItemForm* pForm1 = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewItemForm, m_pSelected->GetBackstageAttachedView());
				if (pForm1 != NULL && pForm1->m_pWndForm->GetSafeHwnd() != NULL)
				{
					HWND hwndPanel = GetParentMenuBar()->GetSafeHwnd();
					GetParentMenuBar()->SetRedraw(FALSE);

					CBCGPRibbonBackstageViewItemForm* pForm2 = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewItemForm, pElem->GetBackstageAttachedView());
					if (pForm2 != NULL && hwndPanel != NULL)
					{
						if (pForm2->m_pWndForm == NULL)
						{
							pForm2->m_pWndForm = pForm2->OnCreateFormWnd();
						}

						if (pForm2->m_pWndForm->GetSafeHwnd() != NULL)
						{
							m_pNewSelected = pElem;

							ReposActiveForm();

							pForm1->m_pWndForm->ShowWindow(SW_HIDE);
							pForm2->m_pWndForm->ShowWindow(SW_HIDE);

							CSize szPageMax(m_rect.Width(), 0);

							m_clrFillFrame = CBCGPVisualManager::GetInstance()->IsRibbonBackstageWhiteBackground() ?
								RGB(255, 255, 255) : globalData.clrBarFace;

							m_pNewSelected->m_bIsChecked = TRUE;
							m_pSelected->m_bIsChecked = FALSE;

							if (StartPageTransition(hwndPanel,
								pForm1->m_pWndForm->GetSafeHwnd(), pForm2->m_pWndForm->GetSafeHwnd(),
								m_pSelected->GetRect().top > pElem->GetRect().top, CSize(0, 0), szPageMax))
							{
								RedrawElement(m_pSelected);

								if (m_bHasCaptions)
								{
									pForm1->m_strTransitionText = pForm2->m_strText;
									GetParentMenuBar()->RedrawWindow(pForm1->m_rectCaption);
								}

								return;
							}

							GetParentMenuBar()->SetRedraw(TRUE);
						}
					}
				}
			}
			else
			{
				bRedrawCaption = m_bHasCaptions;
			}

			CBCGPBaseRibbonElement* pView = m_pSelected->GetBackstageAttachedView();
			if (pView != NULL)
			{
				ASSERT_VALID(pView);
				pView->SetRect(CRect(0, 0, 0, 0));
				pView->OnAfterChangeRect(NULL);
			}

			m_pSelected->m_bIsChecked = FALSE;
			RedrawElement(m_pSelected);
		}

		m_pSelected = pElem;
		m_pSelected->m_bIsChecked = TRUE;

		CWnd* pParentWnd = GetParentWnd();
		if (pParentWnd->GetSafeHwnd() != NULL)
		{
			// Set Redraw back off to prevent flashing.
			ASSERT_VALID (pParentWnd);
			pParentWnd->SetRedraw(FALSE);
		}

		ReposActiveForm();
		AdjustScrollBars();
		ReposActiveForm();

		if (pParentWnd->GetSafeHwnd () != NULL)
		{
			// Set Redraw back on and force a redraw of everything.
			pParentWnd->SetRedraw ();
			pParentWnd->RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
		}

		RedrawElement(m_pSelected);
	}

	if (bRedrawCaption && m_pSelected != NULL && GetParentMenuBar()->GetSafeHwnd() != NULL)
	{
		CBCGPRibbonBackstageViewItemForm* pView = DYNAMIC_DOWNCAST(CBCGPRibbonBackstageViewItemForm, m_pSelected->GetBackstageAttachedView());
		if (pView != NULL)
		{
			GetParentMenuBar()->RedrawWindow(pView->m_rectCaption);
		}
	}
}
//**********************************************************************************
void CBCGPRibbonBackstageViewPanel::DoPaint(CDC* pDC)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pDC);

	CBCGPRibbonMainPanel::DoPaint(pDC);

	if (!m_rectPageTransition.IsRectEmpty())
	{
		CBCGPVisualManager::GetInstance()->OnFillRibbonBackstageForm(pDC, NULL, m_rectPageTransition);
		DoDrawTransition(pDC, FALSE);
	}

	if (!m_rectScrollCorner.IsRectEmpty())
	{
		CBrush br(globalData.clrBarLight);
		pDC->FillRect(m_rectScrollCorner, &br);
	}
}
//**********************************************************************************
void CBCGPRibbonBackstageViewPanel::OnPageTransitionFinished()
{
	ASSERT_VALID(this);

	if (m_pNewSelected == NULL)
	{
		return;
	}

	ASSERT_VALID(m_pNewSelected);

	if (m_pSelected != NULL)
	{
		ASSERT_VALID(m_pSelected);

		CBCGPBaseRibbonElement* pView = m_pSelected->GetBackstageAttachedView();
		if (pView != NULL)
		{
			ASSERT_VALID(pView);
			pView->SetRect(CRect(0, 0, 0, 0));
			pView->OnAfterChangeRect(NULL);
		}
		
		m_pSelected->m_bIsChecked = FALSE;
		RedrawElement(m_pSelected);
	}

	m_pSelected = m_pNewSelected;

	ReposActiveForm();
	AdjustScrollBars();
	ReposActiveForm();

	RedrawElement(m_pSelected);

	m_pNewSelected = NULL;
}
//**********************************************************************************
void CBCGPRibbonBackstageViewPanel::AdjustScrollBars()
{
	ASSERT_VALID(this);

	if (m_bInAdjustScrollBars)
	{
		return;
	}

	m_bInAdjustScrollBars = TRUE;

	int nScrollOffsetOld = m_nScrollOffset;
	int nScrollOffsetHorzOld = m_nScrollOffsetHorz;

	m_rectScrollCorner.SetRectEmpty();

	m_rectRight.right = m_rect.right;
	m_rectRight.bottom = m_rect.bottom;

	const int cxScroll = ::GetSystemMetrics (SM_CXVSCROLL);
	const int cyScroll = ::GetSystemMetrics (SM_CYHSCROLL);

	int nTotalWidth = m_sizeRightView.cx;
	if (!m_bIsStartPage)
	{
		nTotalWidth += m_rectMenuElements.Width();
	}

	BOOL bHasVertScrollBar = FALSE;
	BOOL bHasHorzScrollBar = m_pScrollBarHorz->GetSafeHwnd() != NULL && nTotalWidth > m_rect.Width();

	// Adjust vertical scroll bar:
	if (m_pScrollBar->GetSafeHwnd() == NULL)
	{
		m_nScrollOffset = 0;
	}
	else
	{
		int nTotalHeight = max(m_nTotalHeight, m_sizeRightView.cy);

		if (nTotalHeight <= m_rect.Height())
		{
			m_pScrollBar->ShowWindow(SW_HIDE);
			m_nScrollOffset = 0;
		}
		else
		{
			int nVertScrollHeight = m_rect.Height();
			if (bHasHorzScrollBar)
			{
				nVertScrollHeight -= cyScroll;
			}

			m_pScrollBar->SetWindowPos (NULL,
				m_rect.right - cxScroll, m_rect.top,
				cxScroll, nVertScrollHeight,
				SWP_NOZORDER | SWP_NOACTIVATE);

			SCROLLINFO si;

			ZeroMemory (&si, sizeof (SCROLLINFO));
			si.cbSize = sizeof (SCROLLINFO);
			si.fMask = SIF_RANGE | SIF_PAGE;

			si.nMin = 0;
			si.nMax = nTotalHeight;
			si.nPage = m_rect.Height ();

			m_pScrollBar->SetScrollInfo (&si, TRUE);
			m_pScrollBar->ShowWindow(SW_SHOWNOACTIVATE);
			m_pScrollBar->RedrawWindow();

			m_rectRight.right -= cxScroll;

			if (m_nScrollOffset > nTotalHeight)
			{
				m_nScrollOffset = nTotalHeight;
			}

			bHasVertScrollBar = TRUE;
		}
	}

	// Adjust horizontal scroll bar:
	if (m_pScrollBarHorz->GetSafeHwnd() == NULL)
	{
		m_nScrollOffsetHorz = 0;
	}
	else
	{
		if (nTotalWidth <= m_rect.Width())
		{
			m_pScrollBarHorz->ShowWindow(SW_HIDE);
			m_nScrollOffsetHorz = 0;
		}
		else
		{
			int nHorzScrollWidth = m_rect.Width();
			if (bHasVertScrollBar)
			{
				nHorzScrollWidth -= cxScroll;
			}

			m_pScrollBarHorz->SetWindowPos (NULL,
				m_rect.left, m_rect.bottom - cyScroll,
				nHorzScrollWidth, cyScroll,
				SWP_NOZORDER | SWP_NOACTIVATE);

			SCROLLINFO si;

			ZeroMemory (&si, sizeof (SCROLLINFO));
			si.cbSize = sizeof (SCROLLINFO);
			si.fMask = SIF_RANGE | SIF_PAGE;

			si.nMin = 0;
			si.nMax = nTotalWidth;
			si.nPage = m_rect.Width ();

			m_pScrollBarHorz->SetScrollInfo (&si, TRUE);
			m_pScrollBarHorz->ShowWindow(SW_SHOWNOACTIVATE);
			m_pScrollBarHorz->RedrawWindow();

			m_rectRight.bottom -= cyScroll;

			if (m_nScrollOffsetHorz > nTotalWidth)
			{
				m_nScrollOffsetHorz = nTotalWidth;
			}
		}
	}

	if (bHasVertScrollBar && bHasHorzScrollBar)
	{
		m_rectScrollCorner = m_rect;
		m_rectScrollCorner.left = m_rectScrollCorner.right - cxScroll - 1;
		m_rectScrollCorner.top = m_rectScrollCorner.bottom - cyScroll - 1;
	}

	if ((nScrollOffsetOld != m_nScrollOffset) || (nScrollOffsetHorzOld != m_nScrollOffsetHorz) && m_pMainButton != NULL)
	{
		ASSERT_VALID (m_pMainButton);
		ASSERT_VALID (m_pMainButton->GetParentRibbonBar ());

		CClientDC dc(m_pMainButton->GetParentRibbonBar());

		CFont* pOldFont = dc.SelectObject (m_pMainButton->GetParentRibbonBar()->GetFont());
		ASSERT (pOldFont != NULL);

		Repos(&dc, m_rect);

		dc.SelectObject (pOldFont);

		GetParentWnd()->RedrawWindow();

		m_pMainButton->GetParentRibbonBar()->SetBackstageHorzOffset(m_nScrollOffsetHorz);
	}

	if (m_pMainButton != NULL)
	{
		ASSERT_VALID (m_pMainButton);
		ASSERT_VALID (m_pMainButton->GetParentRibbonBar ());

		m_pMainButton->GetParentRibbonBar()->SetBackstageHorzScroll(bHasHorzScrollBar);
	}

	m_bInAdjustScrollBars = FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPRecentFilesListBox

IMPLEMENT_DYNAMIC(CBCGPRecentFilesListBox, CBCGPListBox)

CBCGPRecentFilesListBox::CBCGPRecentFilesListBox()
{
	m_bVisualManagerStyle = TRUE;
	m_bFoldersMode = FALSE;
	m_strCaption = "";
	m_bShowCaption = FALSE;
	m_arCustomizeInfo.RemoveAll();
}

CBCGPRecentFilesListBox::~CBCGPRecentFilesListBox()
{
	for (int i = 0; i < m_arIcons.GetSize(); i++)
	{
		HICON hIcon = m_arIcons[i];
		if (hIcon != NULL)
		{
			::DestroyIcon(hIcon);
		}
	}
}


BEGIN_MESSAGE_MAP(CBCGPRecentFilesListBox, CBCGPListBox)
	//{{AFX_MSG_MAP(CBCGPRecentFilesListBox)
	ON_WM_MOUSEMOVE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPRecentFilesListBox message handlers

void CBCGPRecentFilesListBox::OnDrawItemContent(CDC* pDC, CRect rect, int nIndex)
{
	if (IsSeparatorItem(nIndex)||IsCaptionItem(nIndex))
	{
		CBCGPListBox::OnDrawItemContent(pDC, rect, nIndex);
		return;
	}


	if (nIndex >= 0 && nIndex < m_arIcons.GetSize())
	{
		HICON hIcon = m_arIcons[nIndex];
		if (hIcon != NULL)
		{
			CSize sizeIcon = globalUtils.ScaleByDPI(CSize (32, 32));

			::DrawIconEx (pDC->GetSafeHdc (), 
				rect.left + 3, 
				rect.top + max(0, (rect.Height() - sizeIcon.cy) / 2),
				hIcon, sizeIcon.cx, sizeIcon.cy, 0, NULL,
				DI_NORMAL);

			rect.left += sizeIcon.cx + 1;
		}
	}

	CRect rectText = rect;
	rectText.DeflateRect (10, 0);

	CString strText;
	GetText (nIndex, strText);
	CString strName;
	CString strPath;

	if (!AfxExtractSubString(strName, (LPCTSTR)strText, 0, '\n'))
	{
		pDC->DrawText (strText, rectText, DT_LEFT);
	}
	else
	{
		AfxExtractSubString(strPath, (LPCTSTR)strText, 1, '\n');
		int nMargin = max(0, (rectText.Height() - pDC->GetTextExtent(strName).cy - pDC->GetTextExtent(strPath).cy) / 3);

		UINT uiDTFlags = DT_SINGLELINE | DT_END_ELLIPSIS;

		CFont* pOldFont = pDC->SelectObject(&globalData.fontBold);
		ASSERT(pOldFont != NULL);

		rectText.top += nMargin;

		const int nTabStops = (int)m_arTabStops.GetSize();
		if ((GetStyle() & LBS_USETABSTOPS) != 0 || nTabStops > 0)
		{
			rectText.right = rectText.left + m_arTabStops[0];
		}		

		int nTextHeight = pDC->DrawText(strName, rectText, uiDTFlags);

		pDC->SelectObject(pOldFont);
		rectText.top += nTextHeight + nMargin;
		pDC->DrawText(strPath, rectText, uiDTFlags);


		//draw CustomizeInfo
		rectText = rect;
		nMargin = max(0, (rectText.Height() - pDC->GetTextExtent(strName).cy) / 2);
		rectText.top = rect.top+ nMargin;
		rectText.right = rectText.left;
		if ((GetStyle() & LBS_USETABSTOPS) != 0)
		{
			int i;
			CString strTemp;
			for (i = 0; i < nTabStops-1; i++)
			{
				AfxExtractSubString(strTemp, (LPCTSTR)strText, i+2, '\n');	//1st=filename 2nd=path,3th = customize info
				rectText.left = rect.left + m_arTabStops[i];
				rectText.right = rect.left + m_arTabStops[i+1];
				pDC->DrawText(strTemp, rectText, uiDTFlags);
			}
		}
	}
}
//**********************************************************************************************
void CBCGPRecentFilesListBox::SetFoldersMode(BOOL bSet)
{
	if (m_bFoldersMode != bSet)
	{
		m_bFoldersMode = bSet;

		if (GetSafeHwnd() != NULL)
		{
			FillList();
		}
	}
}
//**********************************************************************************************
int CBCGPRecentFilesListBox::AddItem(const CString& strFilePath, UINT nCmd, BOOL bPin)
{
	if (strFilePath.IsEmpty())
	{
		return -1;
	}

	TCHAR path[_MAX_PATH];   
	TCHAR name[_MAX_PATH];   
	TCHAR drive[_MAX_DRIVE];   
	TCHAR dir[_MAX_DIR];
	TCHAR fname[_MAX_FNAME];   
	TCHAR ext[_MAX_EXT];

#if _MSC_VER < 1400
	_tsplitpath (strFilePath, drive, dir, fname, ext);
	_tmakepath (path, drive, dir, NULL, NULL);
	_tmakepath (name, NULL, NULL, fname, ext);
#else
	_tsplitpath_s (strFilePath, drive, _MAX_DRIVE, dir, _MAX_DIR, fname, _MAX_FNAME, ext, _MAX_EXT);
	_tmakepath_s (path, drive, dir, NULL, NULL);
	_tmakepath_s (name, NULL, NULL, fname, ext);
#endif
	
	CString strItem;
	CString strToolTip = strFilePath;

	if (m_bFoldersMode)
	{
		CString strPath = path;
		if (strPath.GetLength() > 0 && 
			(strPath[strPath.GetLength() - 1] == _T('\\') || strPath[strPath.GetLength() - 1] == _T('/')))
		{
			strPath = strPath.Left(strPath.GetLength() - 1);
		}

		int nIndex = max(strPath.ReverseFind(_T('\\')), strPath.ReverseFind(_T('/')));
		if (nIndex >= 0)
		{
			strItem = strPath.Right(strPath.GetLength() - nIndex - 1);
		}

		strToolTip = strPath;
	}
	else
	{
		strItem = name;
	}

	strItem += _T("\n");
	strItem += path;

	int nIndex = -1;

	if (FindStringExact(-1, strItem) < 0)
	{
		nIndex = AddString(strItem);
		SetItemData(nIndex, (DWORD_PTR)nCmd);

		SetItemToolTip(nIndex, strToolTip);

		if (bPin)
		{
			SetItemPinned(nIndex);
		}

		HICON hIcon = NULL;
		CString strFile = m_bFoldersMode ? path : strFilePath;

		CFrameWnd* pParentFrame = BCGCBProGetTopLevelFrame (this);

		CBCGPMDIFrameWnd* pMainFrame = DYNAMIC_DOWNCAST (CBCGPMDIFrameWnd, pParentFrame);
		if (pMainFrame != NULL)
		{
			hIcon = pMainFrame->OnGetRecentFileIcon(this, strFile, nIndex);
		}
		else	// Maybe, SDI frame...
		{
			CBCGPFrameWnd* pFrame = DYNAMIC_DOWNCAST (CBCGPFrameWnd, pParentFrame);
			if (pFrame != NULL)
			{
				hIcon = pFrame->OnGetRecentFileIcon(this, strFile, nIndex);
			}
		}

		m_arIcons.Add(hIcon);
	}

	return nIndex;
}
//***********************************************************************************************	
void CBCGPRecentFilesListBox::FillList(LPCTSTR lpszSelectedPath/* = NULL*/)
{
	CStringArray arRecent;
	arRecent.SetSize(CBCGPWinApp::_GetRecentFilesCount());

	for (int i = 0; i < CBCGPWinApp::_GetRecentFilesCount(); i++)
	{
		arRecent[i] = CBCGPWinApp::_GetRecentFilePath(i);
	}

	CStringArray arPinned;
	if (g_pWorkspace != NULL) 
	{
		arPinned.Copy(g_pWorkspace->GetPinnedPaths(!m_bFoldersMode));
	}

	FillList(arRecent, arPinned, lpszSelectedPath);
}
//***********************************************************************************************	
void CBCGPRecentFilesListBox::FillList(const CStringArray& arRecent, const CStringArray& arPinned, LPCTSTR lpszSelectedPath/* = NULL*/)
{
	ASSERT(GetSafeHwnd() != NULL);

	ResetContent();
	ResetPins();
	
	m_lstCaptionIndexes.RemoveAll();

	m_nHighlightedItem = -1;
	m_bIsPinHighlighted = FALSE;

	int cyIcon = globalUtils.ScaleByDPI(32);
	int nItemHeight = max(cyIcon + 4, globalData.GetTextHeight () * 5 / 2);
	
	SetItemHeight(-1, nItemHeight);

	int i = 0;
	for (i = 0; i < m_arIcons.GetSize(); i++)
	{
		HICON hIcon = m_arIcons[i];
		if (hIcon != NULL)
		{
			::DestroyIcon(hIcon);
		}
	}

	m_arIcons.RemoveAll();

	BOOL bHasPinnedItems = FALSE;

	if (m_bShowCaption)
	{
		AddCaption(m_strCaption);
		m_arIcons.Add(NULL);
	}

	// Add "pinned" items first
	for (i = 0; i < (int)arPinned.GetSize(); i++)
	{
		int nIndex = AddItem(arPinned[i], 0, TRUE);
		if (nIndex >= 0 && lpszSelectedPath != NULL && arPinned[i] == lpszSelectedPath)
		{
			SetCurSel(nIndex);
		}

		if (nIndex >= 0)
		{
			bHasPinnedItems = TRUE;
		}
	}

	if (bHasPinnedItems)
	{
		AddSeparator();
		m_arIcons.Add(NULL);
	}

	// Add MRU files:
	for (i = 0; i < (int)arRecent.GetSize(); i++)
	{
		CString strPath = arRecent[i];

		int nIndex = AddItem(strPath, (ID_FILE_MRU_FILE1 + i));
		if (nIndex >= 0 && lpszSelectedPath != NULL && strPath == lpszSelectedPath)
		{
			SetCurSel(nIndex);
		}
	}

	int nLastIndex = GetCount() - 1;
	if (nLastIndex >= 0 && IsSeparatorItem(nLastIndex))
	{
		DeleteString(nLastIndex);
	}
}
//***********************************************************************************************	
void CBCGPRecentFilesListBox::OnClickPin(int nClickedItem)
{
	BOOL bIsPinned = !IsItemPinned(nClickedItem);

	if (g_pWorkspace != NULL)
	{
		CString strPath = GetItemPath(nClickedItem);

		g_pWorkspace->PinPath(!m_bFoldersMode, strPath, bIsPinned);

		FillList(strPath);
	}
	else
	{
		SetItemPinned(nClickedItem, bIsPinned);
	}
}
//***********************************************************************************************	
void CBCGPRecentFilesListBox::OnClickItem(int nClickedItem)
{
	if (m_bFoldersMode)
	{
		OnChooseRecentFolder(GetItemPath(nClickedItem));
	}
	else
	{
		UINT uiCmd = (UINT)GetItemData(nClickedItem);
		if (uiCmd == 0)
		{
			OnChoosePinnedFile(GetItemPath(nClickedItem));
		}
		else
		{
			OnChooseRecentFile(uiCmd);
		}
	}
}
//***********************************************************************************************	
void CBCGPRecentFilesListBox::OnMouseMove(UINT nFlags, CPoint point)
{
	if (m_nClickedItem < 0)
	{
		CBCGPListBox::OnMouseMove(nFlags, point);
	}
}
//***********************************************************************************************
BOOL CBCGPRecentFilesListBox::CloseBackstageView()
{
	for (CWnd* pParent = GetParent(); pParent->GetSafeHwnd() != NULL; pParent = pParent->GetParent())
	{
		if (DYNAMIC_DOWNCAST(CBCGPRibbonBackstageView, pParent) != NULL)
		{
			pParent->SendMessage(WM_CLOSE);
			return TRUE;
		}
	}

	return FALSE;
}

//***********************************************************************************************
void CBCGPRecentFilesListBox::OnChooseRecentFile(UINT uiCmd)
{
	CFrameWnd* pParentFrame = BCGCBProGetTopLevelFrame (this);
	if (pParentFrame != NULL)
	{
		pParentFrame->PostMessage(WM_COMMAND, uiCmd);
	}

	CloseBackstageView();
}
//***********************************************************************************************	
static void _BCGPAppendFilterSuffix(CString& filter, OPENFILENAME& ofn,
	CDocTemplate* pTemplate, CString* pstrDefaultExt)
{
	ASSERT_VALID(pTemplate);
	ASSERT_KINDOF(CDocTemplate, pTemplate);

	CString strFilterExt, strFilterName;
	if (pTemplate->GetDocString(strFilterExt, CDocTemplate::filterExt) &&
	 !strFilterExt.IsEmpty() &&
	 pTemplate->GetDocString(strFilterName, CDocTemplate::filterName) &&
	 !strFilterName.IsEmpty())
	{
		// a file based document template - add to filter list
		ASSERT(strFilterExt[0] == '.');
		if (pstrDefaultExt != NULL)
		{
			// set the default extension
			*pstrDefaultExt = ((LPCTSTR)strFilterExt) + 1;  // skip the '.'
			ofn.lpstrDefExt = (LPTSTR)(LPCTSTR)(*pstrDefaultExt);
			ofn.nFilterIndex = ofn.nMaxCustFilter + 1;  // 1 based number
		}

		// add to filter
		filter += strFilterName;
		ASSERT(!filter.IsEmpty());  // must have a file type name
		filter += (TCHAR)'\0';  // next string please
		filter += (TCHAR)'*';
		filter += strFilterExt;
		filter += (TCHAR)'\0';  // next string please
		ofn.nMaxCustFilter++;
	}
}
//***********************************************************************************************	
void CBCGPRecentFilesListBox::OnChoosePinnedFile(LPCTSTR lpszFile)
{
	ASSERT(lpszFile != NULL);

	CloseBackstageView();

	if (AfxGetApp() == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	AfxGetApp()->OpenDocumentFile(lpszFile);
}
//***********************************************************************************************	
void CBCGPRecentFilesListBox::OnChooseRecentFolder(LPCTSTR lpszFolder)
{
	ASSERT(lpszFolder != NULL);

	CFileDialog dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_FILEMUSTEXIST, NULL, BCGCBProGetTopLevelFrame(this));

	CloseBackstageView();

	if (AfxGetApp() == NULL)
	{
		ASSERT(FALSE);
		return;
	}

	CString title;
	VERIFY(title.LoadString(AFX_IDS_OPENFILE));

	CString strFilter;
	CString strDefault;

	CDocManager* pDocManager = AfxGetApp()->m_pDocManager;
	if (pDocManager != NULL)
	{
		BOOL bFirst = TRUE;
		for (POSITION pos = pDocManager->GetFirstDocTemplatePosition (); pos != NULL;)
		{
			CDocTemplate* pTemplate = pDocManager->GetNextDocTemplate (pos);
			ASSERT_VALID (pTemplate);

			_BCGPAppendFilterSuffix(strFilter, dlgFile.m_ofn, pTemplate,
				bFirst ? &strDefault : NULL);
			bFirst = FALSE;
		}
	}

	CString fileName;

	// append the "*.*" all files filter
	CString allFilter;
	VERIFY(allFilter.LoadString(AFX_IDS_ALLFILTER));
	strFilter += allFilter;
	strFilter += (TCHAR)'\0';   // next string please
	strFilter += _T("*.*");
	strFilter += (TCHAR)'\0';   // last string
	dlgFile.m_ofn.nMaxCustFilter++;

	dlgFile.m_ofn.lpstrFilter = strFilter;
	dlgFile.m_ofn.lpstrTitle = title;
	dlgFile.m_ofn.lpstrFile = fileName.GetBuffer(_MAX_PATH);
	dlgFile.m_ofn.lpstrInitialDir = lpszFolder;

	if (dlgFile.DoModal() == IDOK)
	{
		AfxGetApp()->OpenDocumentFile(fileName);
	}

	fileName.ReleaseBuffer();
}
//***********************************************************************************************	
BOOL CBCGPRecentFilesListBox::OnReturnKey()
{
	int nCurSel = GetCurSel();
	if (nCurSel >= 0)
	{
		if (m_bFoldersMode)
		{
			OnChooseRecentFolder(GetItemPath(nCurSel));
		}
		else
		{
			UINT uiCmd = (UINT)GetItemData(nCurSel);
			if (uiCmd == 0)
			{
				OnChoosePinnedFile(GetItemPath(nCurSel));
			}
			else
			{
				OnChooseRecentFile(uiCmd);
			}
		}
	}

	return TRUE;
}
//***********************************************************************************************	
CString CBCGPRecentFilesListBox::GetItemPath(int nItem)
{
	CString strText;
	GetText(nItem, strText);

	CString strPath;

	int nIndex = strText.Find(_T('\n'));
	if (nIndex < 0)
	{
		return strPath;
	}

	strPath = strText.Right(strText.GetLength() - nIndex - 1);

	if (!m_bFoldersMode)
	{
		CString strFile = strText.Left(nIndex);
		strPath += strFile;
	}
	
	return strPath;
}

#endif // BCGP_EXCLUDE_RIBBON
