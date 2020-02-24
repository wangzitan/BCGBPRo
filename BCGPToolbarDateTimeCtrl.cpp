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
// BCGPToolbarDateTimeCtrl.cpp: implementation of the CBCGPToolbarDateTimeCtrl class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "bcgglobals.h"
#include "BCGPToolBar.h"
#include "BCGPToolbarMenuButton.h"
#include "MenuImages.h"
#include "BCGPToolbarDateTimeCtrl.h"
#include "BCGPGlobalUtils.h"
#include "BCGPDrawManager.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

BOOL CBCGPToolbarDateTimeCtrl::m_bThemedDateTimeControl = FALSE;

IMPLEMENT_SERIAL(CBCGPToolbarDateTimeCtrl, CBCGPToolbarButton, 1)

#define DEFAULT_SIZE	globalUtils.ScaleByDPI(100)
#define HORZ_MARGIN		globalUtils.ScaleByDPI(3)
#define VERT_MARGIN		globalUtils.ScaleByDPI(3)

BEGIN_MESSAGE_MAP(CBCGPDateTimeCtrlWin, CDateTimeCtrl)
	//{{AFX_MSG_MAP(CBCGPDateTimeCtrlWin)
	ON_NOTIFY_REFLECT(DTN_DATETIMECHANGE, OnDateTimeChange)
	ON_NOTIFY_REFLECT(DTN_DROPDOWN, OnDateTimeDropDown)
	ON_NOTIFY_REFLECT(DTN_CLOSEUP, OnDateTimeCloseUp)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

#pragma warning (disable : 4310)

void CBCGPDateTimeCtrlWin::OnDateTimeChange (NMHDR * pNotifyStruct, LRESULT* pResult)
{
	if (!m_bMonthCtrlDisplayed)
	{
		LPNMDATETIMECHANGE pNotify = (LPNMDATETIMECHANGE) pNotifyStruct;
		GetOwner()->PostMessage (WM_COMMAND, MAKEWPARAM(pNotify->nmhdr.idFrom, BN_CLICKED), (LPARAM) m_hWnd);
	}

	*pResult = 0;
}
//**************************************************************************************
void CBCGPDateTimeCtrlWin::OnDateTimeDropDown (NMHDR* /* pNotifyStruct */, LRESULT* pResult)
{
	m_bMonthCtrlDisplayed = true;

	*pResult = 0;
}
//**************************************************************************************
void CBCGPDateTimeCtrlWin::OnDateTimeCloseUp (NMHDR* pNotifyStruct, LRESULT* pResult)
{
	m_bMonthCtrlDisplayed = false;

	LPNMDATETIMECHANGE pNotify = (LPNMDATETIMECHANGE) pNotifyStruct;

	GetOwner()->PostMessage (WM_COMMAND, MAKEWPARAM(pNotify->nmhdr.idFrom, BN_CLICKED), (LPARAM) m_hWnd);
	*pResult = 0;
}

#pragma warning (default : 4310)

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPToolbarDateTimeCtrl::CBCGPToolbarDateTimeCtrl()
{
	m_dwStyle = WS_CHILD | WS_VISIBLE;
	m_iWidth = globalUtils.ScaleByDPI(DEFAULT_SIZE);

	Initialize ();
}
//**************************************************************************************
CBCGPToolbarDateTimeCtrl::CBCGPToolbarDateTimeCtrl (UINT uiId,
			int iImage,
			DWORD dwStyle,
			int iWidth) :
			CBCGPToolbarButton (uiId, iImage)
{
	m_dwStyle = dwStyle | WS_CHILD | WS_VISIBLE;
	m_iWidth = (iWidth == 0) ? globalUtils.ScaleByDPI(DEFAULT_SIZE) : iWidth;

	Initialize ();
}
//**************************************************************************************
void CBCGPToolbarDateTimeCtrl::Initialize ()
{
	m_pWndDateTime = NULL;
	m_pWndDateTimeThemed = NULL;
	m_bHorz = TRUE;
	m_dwTimeStatus = GDT_VALID;
	m_time = CTime::GetCurrentTime ();
}
//**************************************************************************************
CBCGPToolbarDateTimeCtrl::~CBCGPToolbarDateTimeCtrl ()
{
	if (m_pWndDateTime != NULL)
	{
		m_pWndDateTime->DestroyWindow ();
		delete m_pWndDateTime;
	}

	if (m_pWndDateTimeThemed != NULL)
	{
		m_pWndDateTimeThemed->DestroyWindow ();
		delete m_pWndDateTimeThemed;
	}
}
//**************************************************************************************
void CBCGPToolbarDateTimeCtrl::CopyFrom (const CBCGPToolbarButton& s)
{
	CBCGPToolbarButton::CopyFrom (s);

	DuplicateData ();

	const CBCGPToolbarDateTimeCtrl& src = (const CBCGPToolbarDateTimeCtrl&) s;
	m_dwStyle = src.m_dwStyle;
	m_iWidth = src.m_iWidth;
}
//**************************************************************************************
void CBCGPToolbarDateTimeCtrl::Serialize (CArchive& ar)
{
	CBCGPToolbarButton::Serialize (ar);

	if (ar.IsLoading ())
	{
		ar >> m_iWidth;
		m_rect.right = m_rect.left + m_iWidth;
		ar >> m_dwStyle;
		ar >> m_dwTimeStatus;
		ar >> m_time;

		if (m_pWndDateTimeThemed->GetSafeHwnd() != NULL)
		{
			COleDateTime dt = m_time.GetTime();
			m_pWndDateTimeThemed->SetDate(dt);
		}
		else if (m_pWndDateTime->GetSafeHwnd() != NULL)
		{
			m_pWndDateTime->SetTime (m_dwTimeStatus == GDT_VALID? &m_time : NULL);
		}

		DuplicateData ();
	}
	else
	{
		ReadData();

		ar << m_iWidth;
		ar << m_dwStyle;
		ar << m_dwTimeStatus;
		ar << m_time;
	}
}
//**************************************************************************************
SIZE CBCGPToolbarDateTimeCtrl::OnCalculateSize (CDC* pDC, const CSize& sizeDefault, BOOL bHorz)
{
	if (!IsVisible())
	{
		return CSize(0,0);
	}

	m_bHorz = bHorz;
	m_sizeText = CSize (0, 0);

	CWnd* pWndCtrl = GetCtrl();

	int cy = max (sizeDefault.cy, GetPreferedCtrlHeight() + VERT_MARGIN);

	if (bHorz)
	{
		if (pWndCtrl->GetSafeHwnd () != NULL && !m_bIsHidden)
		{
			pWndCtrl->ShowWindow (SW_SHOWNOACTIVATE);
		}

		if (m_bTextBelow && !m_strText.IsEmpty())
		{
			CRect rectText (0, 0, m_iWidth, cy);
			pDC->DrawText (m_strText, rectText, DT_CENTER | DT_CALCRECT | DT_WORDBREAK);
			m_sizeText = rectText.Size ();
		}

		return CSize (m_iWidth, cy + m_sizeText.cy);
	}
	else
	{
		if (pWndCtrl->GetSafeHwnd () != NULL &&
			(pWndCtrl->GetStyle () & WS_VISIBLE))
		{
			pWndCtrl->ShowWindow (SW_HIDE);
		}

		return CBCGPToolbarButton::OnCalculateSize (pDC, sizeDefault, bHorz);
	}
}
//**************************************************************************************
void CBCGPToolbarDateTimeCtrl::OnMove ()
{
	AdjustRect();
}
//**************************************************************************************
void CBCGPToolbarDateTimeCtrl::OnSize (int iSize)
{
	m_iWidth = iSize;
	m_rect.right = m_rect.left + m_iWidth;

	if (m_bTextBelow && !m_strText.IsEmpty() && GetParentWnd()->GetSafeHwnd() != NULL)
	{
		CRect rectText(0, 0, m_iWidth, 32767);

		CClientDC dc(GetParentWnd());
		CBCGPFontSelector fs(dc, &globalData.fontRegular);

		dc.DrawText(m_strText, rectText, DT_CENTER | DT_CALCRECT | DT_WORDBREAK);

		m_sizeText = rectText.Size ();
	}

	AdjustRect();
}
//**************************************************************************************
void CBCGPToolbarDateTimeCtrl::OnChangeParentWnd (CWnd* pWndParent)
{
	CBCGPToolbarButton::OnChangeParentWnd (pWndParent);

	CWnd* pWndCtrl = GetCtrl();

	if (pWndCtrl->GetSafeHwnd () != NULL)
	{
		CWnd* pWndParentCurr = pWndCtrl->GetParent ();
		ASSERT (pWndParentCurr != NULL);

		if (pWndParent != NULL &&
			pWndParentCurr->GetSafeHwnd () == pWndParent->GetSafeHwnd ())
		{
			return;
		}

		pWndCtrl->DestroyWindow ();
		delete pWndCtrl;
		
		m_pWndDateTime = NULL;
		m_pWndDateTimeThemed = NULL;
	}

	if (pWndParent == NULL || pWndParent->GetSafeHwnd () == NULL)
	{
		return;
	}

	CRect rect = m_rect;
	rect.InflateRect (-2, 0);
	rect.bottom = rect.top + GetPreferedCtrlHeight();

	if (m_bThemedDateTimeControl)
	{
		if ((m_pWndDateTimeThemed = CreateThemedDateTimeCtrl (pWndParent, rect)) == NULL)
		{
			ASSERT (FALSE);
			return;
		}
	}
	else
	{
		if ((m_pWndDateTime = CreateDateTimeCtrl (pWndParent, rect)) == NULL)
		{
			ASSERT (FALSE);
			return;
		}
	}

	pWndCtrl = GetCtrl();

	pWndCtrl->SetFont (&globalData.fontRegular);

	if (m_bThemedDateTimeControl)
	{
		COleDateTime dt = m_time.GetTime();
		m_pWndDateTimeThemed->SetDate(dt);
	}
	else
	{
		m_pWndDateTime->SetTime (m_dwTimeStatus == GDT_VALID? &m_time : NULL);
	}

	if (CBCGPToolBar::IsCustomizeMode())
	{
		pWndCtrl->EnableWindow(FALSE);
	}
}
//**************************************************************************************
void CBCGPToolbarDateTimeCtrl::AdjustRect ()
{
	CWnd* pWndCtrl = GetCtrl();
	if (pWndCtrl->GetSafeHwnd () == NULL ||(pWndCtrl->GetStyle () & WS_VISIBLE) == 0)
	{
		return;
	}
	
	pWndCtrl->SetWindowPos (NULL, m_rect.left + HORZ_MARGIN, m_rect.top + VERT_MARGIN / 2,
		m_rect.Width () - 2 * HORZ_MARGIN, GetPreferedCtrlHeight(), SWP_NOZORDER | SWP_NOACTIVATE);
}
//**************************************************************************************
void CBCGPToolbarDateTimeCtrl::ReadData()
{
	if (m_pWndDateTime->GetSafeHwnd() != NULL)
	{
		m_dwTimeStatus = m_pWndDateTime->GetTime(m_time);
		return;
	}

	if (m_pWndDateTimeThemed->GetSafeHwnd() == NULL)
	{
		return;
	}

	COleDateTime date = m_pWndDateTimeThemed->GetDate();
	
	switch (date.GetStatus())
	{
	case COleDateTime::valid:
		m_dwTimeStatus = (DWORD)GDT_VALID;
		m_time = CTime(date.GetYear(), date.GetMonth(), date.GetDay(), date.GetHour(), date.GetMinute(), date.GetSecond());
		break;
		
	case COleDateTime::invalid:
		m_dwTimeStatus = (DWORD)GDT_ERROR;
		m_time = CTime();
		break;

	case COleDateTime::null:
		m_dwTimeStatus = (DWORD)GDT_NONE;
		m_time = CTime();
		break;
	}
}
//**************************************************************************************
int CBCGPToolbarDateTimeCtrl::GetPreferedCtrlHeight()
{
	return globalData.GetTextHeight() + VERT_MARGIN;
}

#pragma warning (disable : 4310)

BOOL CBCGPToolbarDateTimeCtrl::NotifyCommand (int iNotifyCode)
{
	if (GetCtrl()->GetSafeHwnd () == NULL)
	{
		return FALSE;
	}

	if (iNotifyCode == LOWORD(DTN_DATETIMECHANGE) || iNotifyCode == DTN_DATETIMECHANGE)
	{
		ReadData();

		//------------------------------------------------------
		// Try set selection in ALL DateTimeCtrl's with the same ID:
		//------------------------------------------------------
		CObList listButtons;
		if (CBCGPToolBar::GetCommandButtons (m_nID, listButtons) > 0)
		{
			for (POSITION posCombo = listButtons.GetHeadPosition (); posCombo != NULL;)
			{
				CBCGPToolbarDateTimeCtrl* pDateTime = DYNAMIC_DOWNCAST (CBCGPToolbarDateTimeCtrl, listButtons.GetNext (posCombo));
				if (pDateTime != NULL && pDateTime != this)
				{
					pDateTime->SetTime (m_dwTimeStatus == GDT_VALID? &m_time : NULL);
				}
			}
		}
	}

	return TRUE;
}

#pragma warning (default : 4310)

void CBCGPToolbarDateTimeCtrl::OnAddToCustomizePage ()
{
	CObList listButtons;	// Existing buttons with the same command ID

	if (CBCGPToolBar::GetCommandButtons (m_nID, listButtons) == 0)
	{
		return;
	}

	CBCGPToolbarDateTimeCtrl* pOther = 
		(CBCGPToolbarDateTimeCtrl*) listButtons.GetHead ();
	ASSERT_VALID (pOther);
	ASSERT_KINDOF (CBCGPToolbarDateTimeCtrl, pOther);

	CopyFrom (*pOther);
}
//**************************************************************************************
HBRUSH CBCGPToolbarDateTimeCtrl::OnCtlColor (CDC* pDC, UINT /*nCtlColor*/)
{
	pDC->SetTextColor (globalData.clrWindowText);
	pDC->SetBkColor (globalData.clrWindow);

	return (HBRUSH) globalData.brWindow.GetSafeHandle ();
}
//**************************************************************************************
void CBCGPToolbarDateTimeCtrl::OnDraw (CDC* pDC, const CRect& rect, CBCGPToolBarImages* pImages,
						BOOL bHorz, BOOL bCustomizeMode,
						BOOL bHighlight,
						BOOL bDrawBorder, BOOL bGrayDisabledButtons)
{
	CWnd* pWndCtrl = GetCtrl();

	if (pWndCtrl->GetSafeHwnd () == NULL || (pWndCtrl->GetStyle () & WS_VISIBLE) == 0)
	{
		CBCGPToolbarButton::OnDraw (pDC, rect, pImages,
							bHorz, bCustomizeMode,
							bHighlight, bDrawBorder, bGrayDisabledButtons);
	}
	else if ((m_bTextBelow && bHorz) && !m_strText.IsEmpty())
	{
		//--------------------
		// Draw button's text:
		//--------------------
		CBCGPVisualManager::BCGBUTTON_STATE state = CBCGPVisualManager::ButtonsIsRegular;
		if (bHighlight)
		{
			state = CBCGPVisualManager::ButtonsIsHighlighted;
		}
		else if (m_nStyle & (TBBS_PRESSED | TBBS_CHECKED))
		{
			state = CBCGPVisualManager::ButtonsIsPressed;
		}
		
		COLORREF clrText = CBCGPVisualManager::GetInstance ()->GetToolbarButtonTextColor(this, state);
		
		COLORREF cltTextOld = pDC->SetTextColor (clrText);

		CRect rectBorder;
		pWndCtrl->GetWindowRect (rectBorder);
		pWndCtrl->GetParent ()->ScreenToClient (rectBorder);
		rectBorder.InflateRect (1, 1);

		CRect rectText = rect;
		rectText.top = (rectBorder.bottom + rect.bottom - m_sizeText.cy) / 2;
		pDC->DrawText (m_strText, &rectText, DT_CENTER | DT_WORDBREAK);

		pDC->SetTextColor(cltTextOld);
	}
}
//**************************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::OnClick (CWnd* /*pWnd*/, BOOL /*bDelay*/)
{	
	CWnd* pWndCtrl = GetCtrl();
	return pWndCtrl->GetSafeHwnd () != NULL && (pWndCtrl->GetStyle () & WS_VISIBLE);
}
//******************************************************************************************
int CBCGPToolbarDateTimeCtrl::OnDrawOnCustomizeList (
	CDC* pDC, const CRect& rect, BOOL bSelected)
{
	int iWidth = CBCGPToolbarButton::OnDrawOnCustomizeList (pDC, rect, bSelected) + 10;

	//------------------------------
	// Simulate DateTimeCtrl appearance:
	//------------------------------
	CRect rectDateTime = rect;
	int iDateTimeWidth = rect.Width () - iWidth;

	if (iDateTimeWidth < 20)
	{
		iDateTimeWidth = 20;
	}

	rectDateTime.left = rectDateTime.right - iDateTimeWidth;
	rectDateTime.DeflateRect (2, 3);

	pDC->FillSolidRect (rectDateTime, globalData.clrBarHilite);
	pDC->Draw3dRect (&rectDateTime,
		globalData.clrBarDkShadow,
		globalData.clrBarHilite);

	rectDateTime.DeflateRect (1, 1);

	pDC->Draw3dRect (&rectDateTime,
		globalData.clrBarShadow,
		globalData.clrBarLight);

	CRect rectBtn = rectDateTime;
	rectBtn.left = rectBtn.right - rectBtn.Height ();
	rectBtn.DeflateRect (1, 1);

	pDC->FillSolidRect (rectBtn, globalData.clrBarFace);
	pDC->Draw3dRect (&rectBtn,
		globalData.clrBarHilite,
		globalData.clrBarDkShadow);

	CBCGPMenuImages::Draw (pDC, CBCGPMenuImages::IdArowDown, rectBtn);

	return rect.Width ();
}
//********************************************************************************************
CBCGPDateTimeCtrlWin* CBCGPToolbarDateTimeCtrl::CreateDateTimeCtrl (CWnd* pWndParent, const CRect& rect)
{
	CBCGPDateTimeCtrlWin* pWndDateTime = new CBCGPDateTimeCtrlWin;
	if (!pWndDateTime->Create (m_dwStyle, rect, pWndParent, m_nID))
	{
		delete pWndDateTime;
		return NULL;
	}

	return pWndDateTime;
}
//****************************************************************************************
CBCGPDateTimeCtrl* CBCGPToolbarDateTimeCtrl::CreateThemedDateTimeCtrl(CWnd* pWndParent, const CRect& rect)
{
	CBCGPDateTimeCtrl* pWndDateTime = new CBCGPDateTimeCtrl;
	pWndDateTime->SetAutoResize(FALSE);

	if (!pWndDateTime->Create(_T(""), WS_CHILD | WS_VISIBLE | WS_BORDER, rect, pWndParent, m_nID))
	{
		delete pWndDateTime;
		return NULL;
	}

	UINT stateFlags = 0;

	if ((m_dwStyle & DTS_TIMEFORMAT) == DTS_TIMEFORMAT)
	{
		stateFlags |= CBCGPDateTimeCtrl::DTM_TIME | CBCGPDateTimeCtrl::DTM_SECONDS;
	}
	else
	{
		stateFlags |= (CBCGPDateTimeCtrl::DTM_DATE | CBCGPDateTimeCtrl::DTM_DROPCALENDAR);

		if ((m_dwStyle & DTS_LONGDATEFORMAT) == DTS_LONGDATEFORMAT)
		{
			pWndDateTime->m_monthFormat = 1;
		}
		else
		{
			pWndDateTime->m_monthFormat = -1;
		}
	}

	if ((m_dwStyle & DTS_UPDOWN) == DTS_UPDOWN)
	{
		stateFlags |= CBCGPDateTimeCtrl::DTM_SPIN;
	}

	if ((m_dwStyle & DTS_SHOWNONE) == DTS_SHOWNONE)
	{
		stateFlags |= CBCGPDateTimeCtrl::DTM_CHECKBOX;
	}

	const UINT stateMask = 
		CBCGPDateTimeCtrl::DTM_SPIN |
		CBCGPDateTimeCtrl::DTM_DROPCALENDAR | 
		CBCGPDateTimeCtrl::DTM_DATE |
		CBCGPDateTimeCtrl::DTM_TIME24H |
		CBCGPDateTimeCtrl::DTM_CHECKBOX |
		CBCGPDateTimeCtrl::DTM_TIME | 
		CBCGPDateTimeCtrl::DTM_SECONDS |
		CBCGPDateTimeCtrl::DTM_TIME24HBYLOCALE;

	pWndDateTime->SetState (stateFlags, stateMask);

	pWndDateTime->EnableVisualManagerStyle();
	return pWndDateTime;
}
//****************************************************************************************
void CBCGPToolbarDateTimeCtrl::OnShow (BOOL bShow)
{
	CWnd* pWndCtrl = GetCtrl();

	if (pWndCtrl->GetSafeHwnd () != NULL)
	{
		if (bShow && m_bHorz)
		{
			pWndCtrl->ShowWindow (SW_SHOWNOACTIVATE);
			OnMove ();
		}
		else
		{
			pWndCtrl->ShowWindow (SW_HIDE);
		}
	}
}
//*************************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::ExportToMenuButton (CBCGPToolbarMenuButton& menuButton) const
{
	CString strMessage;
	int iOffset;

	if (strMessage.LoadString (m_nID) &&
		(iOffset = strMessage.Find (_T('\n'))) != -1)
	{
		menuButton.m_strText = strMessage.Mid (iOffset + 1);
	}

	return TRUE;
}
//*************************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::SetTime (LPSYSTEMTIME pTimeNew /* = NULL */)
{
	BOOL bResult = FALSE;
	
	if (m_pWndDateTimeThemed->GetSafeHwnd() != NULL)
	{
		m_pWndDateTimeThemed->SetDate(pTimeNew != NULL ? *pTimeNew : COleDateTime());
		bResult = TRUE;
	}
	else if (m_pWndDateTime->GetSafeHwnd() != NULL)
	{
		bResult = m_pWndDateTime->SetTime(pTimeNew);
	}

	if (bResult)
	{
		NotifyCommand (DTN_DATETIMECHANGE);
	}

	return bResult;
}
//*************************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::SetTime (const COleDateTime& timeNew)
{
	BOOL bResult = FALSE;
	
	if (m_pWndDateTimeThemed->GetSafeHwnd() != NULL)
	{
		m_pWndDateTimeThemed->SetDate(timeNew);
		bResult = TRUE;
	}
	else if (m_pWndDateTime->GetSafeHwnd() != NULL)
	{
		bResult = m_pWndDateTime->SetTime(timeNew);
	}
	
	if (bResult)
	{
		NotifyCommand (DTN_DATETIMECHANGE);
	}
	
	return bResult;
}
//*************************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::SetTime (const CTime* pTimeNew)
{
	BOOL bResult = FALSE;
	
	if (m_pWndDateTimeThemed->GetSafeHwnd() != NULL)
	{
		COleDateTime date;
		if (pTimeNew != NULL)
		{
			date = pTimeNew->GetTime();
		}

		m_pWndDateTimeThemed->SetDate(date);
		bResult = TRUE;
	}
	else if (m_pWndDateTime->GetSafeHwnd() != NULL)
	{
		bResult = m_pWndDateTime->SetTime(pTimeNew);
	}
	
	if (bResult)
	{
		NotifyCommand (DTN_DATETIMECHANGE);
	}
	
	return bResult;
}
//*********************************************************************************
DWORD CBCGPToolbarDateTimeCtrl::GetTime(LPSYSTEMTIME pTimeDest)
{
	if (m_pWndDateTime->GetSafeHwnd() != NULL)
	{
		return m_pWndDateTime->GetTime(pTimeDest);
	}

	if (m_pWndDateTimeThemed->GetSafeHwnd() != NULL)
	{
		COleDateTime date = m_pWndDateTimeThemed->GetDate();

		if (date.GetStatus() == COleDateTime::valid && date.GetAsSystemTime(*pTimeDest))
		{
			return GDT_VALID;
		}
	}

	return GDT_NONE;
}
//*********************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::GetTime(COleDateTime& timeDest)
{
	if (m_pWndDateTime->GetSafeHwnd() != NULL)
	{
		return m_pWndDateTime->GetTime(timeDest);
	}
	
	if (m_pWndDateTimeThemed->GetSafeHwnd() != NULL)
	{
		timeDest = m_pWndDateTimeThemed->GetDate();
		return TRUE;
	}
	
	return FALSE;
}
//*********************************************************************************
DWORD CBCGPToolbarDateTimeCtrl::GetTime(CTime& timeDest) const
{
	if (m_pWndDateTime->GetSafeHwnd() != NULL)
	{
		return m_pWndDateTime->GetTime(timeDest);;
	}

	if (m_pWndDateTimeThemed->GetSafeHwnd() != NULL)
	{
		COleDateTime date = m_pWndDateTimeThemed->GetDate();
		
		if (date.GetStatus() == COleDateTime::valid)
		{
			timeDest = CTime(date.GetYear(), date.GetMonth(), date.GetDay(), date.GetHour(), date.GetMinute(), date.GetSecond());
			return GDT_VALID;
		}
	}
	
	return GDT_NONE;
}
//*********************************************************************************
CBCGPToolbarDateTimeCtrl* CBCGPToolbarDateTimeCtrl::GetByCmd (UINT uiCmd)
{
	CBCGPToolbarDateTimeCtrl* pSrcDateTime = NULL;

	CObList listButtons;
	if (CBCGPToolBar::GetCommandButtons(uiCmd, listButtons) > 0)
	{
		for (POSITION posDateTime= listButtons.GetHeadPosition (); pSrcDateTime == NULL && posDateTime != NULL;)
		{
			CBCGPToolbarDateTimeCtrl* pDateTime= DYNAMIC_DOWNCAST(CBCGPToolbarDateTimeCtrl, listButtons.GetNext(posDateTime));
			ASSERT (pDateTime != NULL);

			pSrcDateTime = pDateTime;
		}
	}

	return pSrcDateTime;
}
//*********************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::SetTimeAll (UINT uiCmd, LPSYSTEMTIME pTimeNew /* = NULL */)
{
	CBCGPToolbarDateTimeCtrl* pSrcDateTime = GetByCmd (uiCmd);

	if (pSrcDateTime)
	{
		pSrcDateTime->SetTime (pTimeNew);
	}

	return pSrcDateTime != NULL;
}
//*********************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::SetTimeAll (UINT uiCmd, const COleDateTime& timeNew)
{
	CBCGPToolbarDateTimeCtrl* pSrcDateTime = GetByCmd (uiCmd);

	if (pSrcDateTime)
	{
		pSrcDateTime->SetTime (timeNew);
	}

	return pSrcDateTime != NULL;
}
//*********************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::SetTimeAll (UINT uiCmd, const CTime* pTimeNew)
{
	CBCGPToolbarDateTimeCtrl* pSrcDateTime = GetByCmd (uiCmd);

	if (pSrcDateTime)
	{
		pSrcDateTime->SetTime (pTimeNew);
	}

	return pSrcDateTime != NULL;
}
//*********************************************************************************
DWORD CBCGPToolbarDateTimeCtrl::GetTimeAll (UINT uiCmd, LPSYSTEMTIME pTimeDest)
{
	CBCGPToolbarDateTimeCtrl* pSrcDateTime = GetByCmd (uiCmd);

	if (pSrcDateTime)
	{
		return pSrcDateTime->GetTime (pTimeDest);
	}
	else
		return GDT_NONE;
}
//*********************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::GetTimeAll (UINT uiCmd, COleDateTime& timeDest)
{
	CBCGPToolbarDateTimeCtrl* pSrcDateTime = GetByCmd (uiCmd);

	if (pSrcDateTime)
	{
		return pSrcDateTime->GetTime (timeDest);
	}
	else
		return FALSE;
}
//*********************************************************************************
DWORD CBCGPToolbarDateTimeCtrl::GetTimeAll (UINT uiCmd, CTime& timeDest)
{
	CBCGPToolbarDateTimeCtrl* pSrcDateTime = GetByCmd (uiCmd);

	if (pSrcDateTime)
	{
		return pSrcDateTime->GetTime (timeDest);
	}
	else
		return GDT_NONE;
}
//*********************************************************************************
void CBCGPToolbarDateTimeCtrl::SetStyle (UINT nStyle)
{
	CBCGPToolbarButton::SetStyle (nStyle);

	CWnd* pWndCtrl = GetCtrl();

	if (pWndCtrl != NULL && pWndCtrl->GetSafeHwnd () != NULL)
	{
		BOOL bDisabled = (CBCGPToolBar::IsCustomizeMode () && !IsEditable ()) ||
			(!CBCGPToolBar::IsCustomizeMode () && (m_nStyle & TBBS_DISABLED));

		pWndCtrl->EnableWindow (!bDisabled);
	}
}
//********************************************************************************
BOOL CBCGPToolbarDateTimeCtrl::OnUpdateToolTip (CWnd* /*pWndParent*/, int /*iButtonIndex*/,
												CToolTipCtrl& wndToolTip, CString& strTipText)
{
	if (!m_bHorz)
	{
		return FALSE;
	}

	CString strTips;
	if (OnGetCustomToolTipText (strTips))
	{
		strTipText = strTips;
	}

	CDateTimeCtrl* pWndDate = GetDateTimeCtrl ();
	if (pWndDate != NULL)
	{
		wndToolTip.AddTool (pWndDate, strTipText, NULL, 0);
		return TRUE;
	}

	return FALSE;
}
//*****************************************************************************************
void CBCGPToolbarDateTimeCtrl::OnGlobalFontsChanged()
{
	CBCGPToolbarButton::OnGlobalFontsChanged ();

	CWnd* pWndCtrl = GetCtrl();

	if (pWndCtrl->GetSafeHwnd () != NULL)
	{
		pWndCtrl->SetFont (&globalData.fontRegular);
	}
}
