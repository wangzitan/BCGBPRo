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
// BCGPPlannerViewSchedule.cpp: implementation of the CBCGPPlannerViewSchedule class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BCGPPlannerViewSchedule.h"

#ifndef BCGP_EXCLUDE_PLANNER

#include "BCGPPlannerManagerCtrl.h"
#include "BCGPDrawManager.h"
#include "BCGPAppointmentStorage.h"

#ifndef _BCGPCALENDAR_STANDALONE
	#include "BCGPVisualManager.h"
	#define visualManager	CBCGPVisualManager::GetInstance ()
#else
	#include "BCGPCalendarVisualManager.h"
	#define visualManager	CBCGPCalendarVisualManager::GetInstance ()
#endif

#include "BCGPMath.h"
#include "BCGPGlobalUtils.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif


static const int s_MinResourceRows = 4;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

IMPLEMENT_DYNCREATE(CBCGPPlannerViewSchedule, CBCGPPlannerViewDay)

CBCGPPlannerViewSchedule::CBCGPPlannerViewSchedule()
	: CBCGPPlannerViewDay()
	, m_rectResourcesBar (0, 0, 0, 0)
	, m_nColWidth        (0)
	, m_bMultiResources  (TRUE)
{
}

CBCGPPlannerViewSchedule::~CBCGPPlannerViewSchedule()
{
}

#ifdef _DEBUG
void CBCGPPlannerViewSchedule::AssertValid() const
{
	CBCGPPlannerViewDay::AssertValid();
}

void CBCGPPlannerViewSchedule::Dump(CDumpContext& dc) const
{
	CBCGPPlannerViewDay::Dump(dc);
}
#endif

BOOL CBCGPPlannerViewSchedule::IsScheduleTimeDeltaDay() const
{
	return GetPlanner()->IsScheduleTimeDeltaDay();
}

BOOL CBCGPPlannerViewSchedule::IsCurrentTimeVisible() const
{
	const int nDays = GetViewDuration ();
	if (m_ViewRects.GetSize() != nDays)
	{
		return FALSE;
	}

	COleDateTime curTime (COleDateTime::GetCurrentTime ());
	if (curTime.GetHour() < GetFirstViewHour() || GetLastViewHour() < curTime.GetHour())
	{
		return FALSE;
	}

	COleDateTime dtStart(GetDateStart ());
	dtStart = COleDateTime (dtStart.GetYear (), dtStart.GetMonth (), dtStart.GetDay (), 0, 0, 0);

	COleDateTime dtCur(curTime.GetYear (), curTime.GetMonth (), curTime.GetDay (), 0, 0, 0);

	int nIndex = (int)((dtCur - dtStart).GetDays());
	if (nIndex < 0 || nDays <= nIndex)
	{
		return FALSE;
	}

	int nPos = bcg_round(((curTime.GetHour () - GetFirstViewHour()) * 60 + curTime.GetMinute ()) * m_nColWidth / 
		(double)CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ())) + m_ViewRects[nIndex].left - m_nScrollHorzOffset + 1;

	return m_rectApps.left <= nPos && nPos <= m_rectApps.right;
}

void CBCGPPlannerViewSchedule::OnDrawAppointmentsDuration (CDC* pDC)
{
/*
	if ((GetDrawFlags () & BCGP_PLANNER_DRAW_VIEW_NO_DURATION) == 
			BCGP_PLANNER_DRAW_VIEW_NO_DURATION)
	{
		return;
	}
*/
	XBCGPAppointmentArray& arQueryApps = GetQueryedAppointments ();
	XBCGPAppointmentArray& arDragApps  = GetDragedAppointments ();

	if (arQueryApps.GetSize () == 0 && arDragApps.GetSize () == 0)
	{
		return;
	}

	BOOL bDragDrop        = IsDragDrop ();
	DROPEFFECT dragEffect = GetDragEffect ();
	BOOL bDragMatch       = IsCaptureMatched ();

	COleDateTime dtS = GetDateStart ();
	COleDateTime dtE = GetDateEnd ();

	CRect rtFill;
	rtFill.bottom = m_rectTimeBar.bottom - (m_rectTimeBar.Height() / 2 - m_nRowHeight) / 2;
	rtFill.top = rtFill.bottom - m_nRowHeight;

	if ((rtFill.top - m_rectTimeBar.top) != (m_rectTimeBar.bottom - rtFill.bottom))
	{
		rtFill.top++;
	}

	for (int nApp = 0; nApp < 2; nApp++)
	{
		XBCGPAppointmentArray& arApps = nApp == 1 ? arDragApps : arQueryApps;

		if (nApp == 1)
		{
			bDragDrop = bDragDrop && arDragApps.GetSize () > 0;
		}

		if (arApps.GetSize () == 0)
		{
			continue;
		}

		for (int i = 0; i < (int)arApps.GetSize (); i++)
		{
			CBCGPAppointment* pApp = arApps [i];
			if (pApp == NULL/* || !(pApp->IsAllDay () || pApp->IsMultiDay ())*/ || 
				pApp->GetDurationColor () == CLR_DEFAULT)
			{
				continue;
			}

			BOOL bDraw = FALSE;

			if (bDragDrop && dragEffect != DROPEFFECT_NONE && 
				pApp->IsSelected () && nApp == 0)
			{
				if ((dragEffect & DROPEFFECT_COPY) == DROPEFFECT_COPY || bDragMatch)
				{
					bDraw = TRUE;
				}
			}
			else
			{
				bDraw = TRUE;
			}

			if(!bDraw)
			{
				continue;
			}

			CBrush br (visualManager->GetPlannerAppointmentDurationColor(this, pApp));

			CRect rt(pApp->GetRectDraw());

			rtFill.left = rt.left;
			rtFill.right = rt.right;

			pDC->FillRect (rtFill, &br);
		}
	}
}

void CBCGPPlannerViewSchedule::GetCaptionFormatStrings (CStringArray& sa)
{
	sa.RemoveAll ();

	const BOOL bCompact = (GetDrawFlags () & (BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_COMPACT | BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_WITH_HEADER)) != 0;
	const BOOL bCaptionWithHeader = (GetDrawFlags() & BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_WITH_HEADER) == 
			BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_WITH_HEADER;

	if (!bCompact || bCaptionWithHeader)
	{
		CBCGPPlannerViewDay::GetCaptionFormatStrings(sa);
	}
	else
	{
		sa.Add (_T("d dddd"));
		sa.Add (_T("d ddd"));
		sa.Add (_T("d"));
	}
}

void CBCGPPlannerViewSchedule::AdjustScrollSizes ()
{
	const int nDays = GetViewDuration ();
	const int nCount = GetViewHours() * 60 / CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ());

	const int dxColumn = nDays * nCount * m_nColWidth;

	m_nScrollHorzPage  = m_rectApps.Width ();
	m_nScrollHorzTotal = dxColumn - 1;
	m_nScrollHorzOffset = min (m_nScrollHorzOffset, m_nScrollHorzTotal - m_nScrollHorzPage + 1);

	if ((m_nScrollHorzTotal - 1) < m_nScrollHorzPage)
	{
		m_nScrollHorzPage = 0;
		m_nScrollHorzTotal = 0;
		m_nScrollHorzOffset = 0;
	}

	const int nResCount = (int)m_Resources.GetSize();

	BOOL bScroll = FALSE;

	if (nResCount <= 1)
	{
		XBCGPAppointmentArray& arQueryApps = GetQueryedAppointments ();
		XBCGPAppointmentArray& arDragApps = GetDragedAppointments ();

		if (arQueryApps.GetSize () != 0 || arDragApps.GetSize () != 0)
		{
			XResource* pResource = nResCount == 1 ? &m_Resources[0] : NULL;

			m_nScrollTotal = 0;

			BOOL bFirst = TRUE;
			BOOL bDragDrop        = IsDragDrop ();
			DROPEFFECT dragEffect = GetDragEffect ();
			BOOL bDragMatch       = IsCaptureMatched ();

			bDragDrop = !bDragDrop || 
				(bDragDrop && ((dragEffect & DROPEFFECT_COPY) == DROPEFFECT_COPY && bDragMatch) || 
				!bDragMatch);
			bDragDrop = bDragDrop && arDragApps.GetSize () > 0;

			int nTop = 0;
			int nBottom = 0;

			for (int nApp = 0; nApp < 2; nApp++)
			{
				if (!bDragDrop && nApp == 1)
				{
					continue;
				}

				XBCGPAppointmentArray& arApps = nApp == 1 ? arDragApps : arQueryApps;

				XBCGPAppointmentArray arByInterval;

				int i;

				int nStartIndex = 0;
				if (pResource != NULL)
				{
					for (i = 0; i < (int)arApps.GetSize (); i++)
					{
						const CBCGPAppointment* pApp = arApps[i];
						if (pApp != NULL && pApp->GetResourceID () == pResource->m_ResourceID)
						{
							nStartIndex = i;
							break;
						}
					}
				}

				for (i = nStartIndex; i < (int)arApps.GetSize (); i++)
				{
					CBCGPAppointment* pApp = arApps[i];
					ASSERT_VALID (pApp);

					if (pResource != NULL && pApp->GetResourceID () != pResource->m_ResourceID)
					{
						break;
					}

					const CRect& rectDraw = pApp->GetRectDraw();
					if (rectDraw.Height() > 0)
					{
						if (bFirst)
						{
							nTop = rectDraw.top;
							nBottom = rectDraw.bottom;
							bFirst = FALSE;
						}
						else
						{
							nTop = min(nTop, rectDraw.top);
							nBottom = max(nBottom, rectDraw.bottom);
						}
					}
				}
			}

			nBottom = (int)ceil((nBottom - nTop) / (double)m_nRowHeight) * m_nRowHeight;
			if (nBottom > m_rectApps.Height())
			{
				m_nScrollTotal = nBottom - 1;
				bScroll = TRUE;
			}
		}
	}
	else
	{
		XResource& resource = m_Resources[m_Resources.GetSize() - 1];

		if (resource.m_Rects.GetSize() > 0 && resource.m_Rects[0].bottom != m_rectApps.bottom)
		{
			m_nScrollTotal = resource.m_Rects[0].bottom - m_rectApps.top - 1;
			bScroll = TRUE;
		}
	}

	if (bScroll)
	{
		m_nScrollPage = m_rectApps.Height ();
		m_nScrollOffset = min (m_nScrollOffset, m_nScrollTotal - m_nScrollPage + 1);
	}

	if (!bScroll || ((m_nScrollTotal - 1) < m_nScrollPage))
	{
		m_nScrollPage = 0;
		m_nScrollTotal = 0;
		m_nScrollOffset = 0;
	}

	CBCGPPlannerView::AdjustScrollSizes ();
}

void CBCGPPlannerViewSchedule::AdjustLayout (CDC* /*pDC*/, const CRect& rectClient)
{
	if (IsCurrentTimeVisible ())
	{
		StartTimer (FALSE);
	}
	else
	{
		StopTimer (FALSE);
	}

	const DWORD dwDrawFlags = GetDrawFlags ();

	m_rectResourcesBar.SetRect(0, 0, 0, 0);

	if (m_Resources.GetSize() > 0)
	{
		m_rectResourcesBar = rectClient;
		m_rectResourcesBar.right = m_rectResourcesBar.left + rectClient.Width() / 7;
	}

	m_rectApps.left = m_rectResourcesBar.right;

	const BOOL bCaptionWithHeader = (dwDrawFlags & BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_WITH_HEADER) == 
		BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_WITH_HEADER;
	const BOOL bHeaderEx = (dwDrawFlags & BCGP_PLANNER_DRAW_VIEW_HEADER_EXTENDED) == 
		BCGP_PLANNER_DRAW_VIEW_HEADER_EXTENDED;
	const BOOL bCaptionEx = (dwDrawFlags & BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_EXTENDED) == 
		BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_EXTENDED;
/*
	const BOOL bTimeBarSimple = (dwDrawFlags & BCGP_PLANNER_DRAW_VIEW_TIME_BAR_NO_MINUTS) == 
		BCGP_PLANNER_DRAW_VIEW_TIME_BAR_NO_MINUTS;
*/
	const int nDays = GetViewDuration ();
	const int nCount = GetViewHours() * 60 / CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ());
	const int nColWidthMin = 3 * m_nRowHeight;

	m_nColWidth = (m_rectApps.Width() - ::GetSystemMetrics(SM_CXVSCROLL)) / (nDays * nCount);
	m_nColWidth = max(m_nColWidth, IsScheduleTimeDeltaDay() ? m_nRowHeight / 2 : nColWidthMin);

	m_nHeaderHeight = bCaptionWithHeader ? 2 : 1;

	if (bCaptionWithHeader || bCaptionEx)
	{
		m_nHeaderHeight = m_nRowHeight;

		if (bCaptionEx)
		{
			m_nHeaderHeight = (m_nHeaderHeight * 5) / 4;
		}

		if (bCaptionWithHeader)
		{
			m_nHeaderHeight += bHeaderEx ? (m_nRowHeight * 5) / 4 : m_nRowHeight;
		}
	}
	
	if (!bCaptionWithHeader && !bCaptionEx)
	{
		m_nHeaderHeight *= m_nRowHeight;
	}

	m_rectApps.top = m_nHeaderHeight;

	m_rectTimeBar = m_rectApps;
	m_rectTimeBar.bottom = m_rectTimeBar.top + nColWidthMin;

	m_rectApps.top = m_rectTimeBar.bottom;
	
	// correct selection
	COleDateTime sel1 (GetSelectionStart ());
	COleDateTime sel2 (GetSelectionEnd ());

	SetSelection (sel1, sel2, FALSE);
}

void CBCGPPlannerViewSchedule::AdjustRects ()
{
	m_rectTimeBar.right = m_rectApps.right;

	CBCGPPlannerViewDay::AdjustRects();

	const int nDays = GetViewDuration ();
	const int nMinuts = CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ());
	const int nCount = GetViewHours() * 60 / nMinuts;

	const int dxColumn = nCount * m_nColWidth;

	CRect rect (m_rectApps);
	rect.right = rect.left + dxColumn;

	for (int nDay = 0; nDay < nDays; nDay++)
	{
		m_ViewRects[nDay] = rect;

		rect.OffsetRect (dxColumn, 0);
	}

	if (m_bMultiResources && m_Resources.GetSize () == 0)
	{
		OnUpdateStorage ();
	}

	if (m_rectTimeBar.right > m_ViewRects[m_ViewRects.GetSize() - 1].right)
	{
		m_rectTimeBar.right = m_ViewRects[m_ViewRects.GetSize() - 1].right;
	}

	const int count = (int) m_Resources.GetSize ();
	if (count > 0)
	{
		int i = 0;

		for (i = 0; i < count; i++)
		{
			XResource& resource = m_Resources[i];
			resource.m_bToolTipNeeded = FALSE;
			resource.m_Rects.SetSize (nDays);
		}

		for (i = 0; i < nDays; i++)
		{
			rect = m_ViewRects[i];
			int dyColumn = CBCGPPlannerView::round (rect.Height () / (double)count);

			BOOL bAdjust = TRUE;
			if ((dyColumn / m_nRowHeight) < s_MinResourceRows)
			{
				dyColumn = s_MinResourceRows * m_nRowHeight;

				if ((dyColumn * count) < rect.Height ())
				{
					dyColumn += m_nRowHeight;
				}

				bAdjust = FALSE;
			}

			rect.bottom = rect.top + dyColumn;

			for (int j = 0; j < count; j++)
			{
				m_Resources[j].m_Rects[i] = rect;

				rect.OffsetRect (0, dyColumn);

				if (bAdjust && j == (count - 2))
				{
					rect.bottom = m_ViewRects[i].bottom;
				}
			}
		}
	}

	AdjustScrollSizes();
}

void CBCGPPlannerViewSchedule::AdjustAppointments ()
{
	XBCGPAppointmentArray& arQueryApps = GetQueryedAppointments ();
	XBCGPAppointmentArray& arDragApps = GetDragedAppointments ();

	const int nDays = GetViewDuration ();

	if ((arQueryApps.GetSize () == 0 && arDragApps.GetSize () == 0) || 
		m_ViewRects.GetSize () != nDays)
	{
		ClearVisibleUpDownIcons ();
		return;
	}

	BOOL bDragDrop        = IsDragDrop ();
	DROPEFFECT dragEffect = GetDragEffect ();
	BOOL bDragMatch       = IsCaptureMatched ();

	bDragDrop = !bDragDrop || 
		(bDragDrop && ((dragEffect & DROPEFFECT_COPY) == DROPEFFECT_COPY && bDragMatch) || 
		!bDragMatch);
	bDragDrop = bDragDrop && arDragApps.GetSize () > 0;

	const int nTimeDelta = CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ());
	const int xOffset = m_nScrollHorzOffset;
	const int yOffset = m_nScrollOffset;

	for (int nApp = 0; nApp < 2; nApp++)
	{
		if (!bDragDrop && nApp == 0)
		{
			continue;
		}

		XBCGPAppointmentArray& arApps = nApp == 0 ? arDragApps : arQueryApps;

		for (int i = 0; i < (int)arApps.GetSize (); i++)
		{
			CBCGPAppointment* pApp = arApps[i];
			ASSERT_VALID (pApp);

			pApp->ResetDraw ();
		}
	}

	COleDateTimeSpan spanDay (1, 0, 0, 0);

	const int nViewHourFirst = GetFirstViewHour();

	const int nAppointmentSpace = globalUtils.ScaleByDPI(BCGP_PLANNER_APPOINTMENT_SPACE);

	const BOOL bResources = m_Resources.GetSize() > 0;
	const int nResCount = bResources ? (int)m_Resources.GetSize() : 1;

	XResource res;
	if (!bResources)
	{
		res.m_ResourceID = GetPlanner()->GetCurrentResourceID();
		res.m_Rects.SetSize(m_ViewRects.GetSize());

		for (int i = 0; i < m_ViewRects.GetSize(); i++)
		{
			res.m_Rects[i] = m_ViewRects[i];
		}
	}

	const BOOL bRealTimeOffset = TRUE;

	for (int nRes = 0; nRes < nResCount; nRes++)
	{
		XResource* pResource = bResources ? &m_Resources[nRes] : &res;

		for (int nApp = 0; nApp < 2; nApp++)
		{
			if (!bDragDrop && nApp == 1)
			{
				continue;
			}

			XBCGPAppointmentArray& arApps = nApp == 1 ? arDragApps : arQueryApps;

			XBCGPAppointmentArray arByInterval;

			int i;

			int nStartIndex = 0;
			if (bResources)
			{
				for (i = 0; i < (int)arApps.GetSize (); i++)
				{
					const CBCGPAppointment* pApp = arApps[i];
					if (pApp != NULL && pApp->GetResourceID () == pResource->m_ResourceID)
					{
						nStartIndex = i;
						break;
					}
				}
			}

			for (i = nStartIndex; i < (int)arApps.GetSize (); i++)
			{
				CBCGPAppointment* pApp = arApps[i];
				ASSERT_VALID (pApp);

				if (bResources && pApp->GetResourceID () != pResource->m_ResourceID)
				{
					break;
				}

				COleDateTime dtStart(pApp->GetStart());
				COleDateTime dtFinish(pApp->GetFinish());

				if ((dtStart < m_DateStart && dtFinish < m_DateStart) ||
					(m_DateEnd < dtStart && m_DateEnd < dtFinish))
				{
					continue;
				}

				if (dtStart == dtFinish)
				{
					if (pApp->IsAllDay())
					{
						dtFinish += COleDateTimeSpan(0, 23, 59, 59);
					}
					else if (!bRealTimeOffset)
					{
						dtStart = COleDateTime(dtStart.GetYear(), dtStart.GetMonth(), dtStart.GetDay(), dtStart.GetHour(), 0, 0);
						dtFinish = dtStart + COleDateTimeSpan(0, 0, nTimeDelta, 0);
					}
				}

				int nDay1 = bcg_clamp((dtStart - m_DateStart).GetDays(), 0, nDays - 1);
				int nDay2 = bcg_clamp(nDay1 + (dtFinish - dtStart).GetDays() + 1, 0, nDays - 1);

				CRect rtApp(0, 0, 0, 0);

				for (int nDay = nDay1; nDay <= nDay2; nDay++)
				{
					COleDateTime date(m_DateStart + COleDateTimeSpan(nDay, 0, 0, 0));

					if (!CBCGPPlannerView::IsAppointmentInDate (*pApp, date))
					{
						continue;
					}

					CRect rtView(pResource->m_Rects[nDay]);

					CRect rtFill (rtView);
					rtFill.top += 1;
					rtFill.bottom = rtFill.top + m_nRowHeight - 1;

					double dMinutes = 0.0;
					if (date < dtStart)
					{
						dMinutes = dtStart.GetHour () * 60 + dtStart.GetMinute ();
					}

					if (bRealTimeOffset)
					{
						rtFill.left += (long)(m_nColWidth * dMinutes / (double)nTimeDelta);
					}
					else
					{
						rtFill.left += m_nColWidth * (long)(dMinutes / (double)nTimeDelta);
					}

					dMinutes = 24.0 * 60.0;
					if (IsOneDay (date, dtFinish) && dtFinish < (date + spanDay))
					{
						dMinutes = dtFinish.GetHour () * 60.0 + dtFinish.GetMinute ();
					}

					if (bRealTimeOffset)
					{
						rtFill.right = rtView.left + (long)(m_nColWidth * dMinutes / (double)nTimeDelta);
					}
					else
					{
						rtFill.right = rtView.left + m_nColWidth * (long)ceil(dMinutes / (double)nTimeDelta);
					}

					if (bRealTimeOffset && (rtFill.right - rtFill.left) < nAppointmentSpace)
					{
						rtFill.right += nAppointmentSpace;
					}

					rtFill.OffsetRect(-m_nColWidth * long((nViewHourFirst * 60) / nTimeDelta), 0);
					rtFill.left = bcgp_max(rtFill.left, rtView.left);
					rtFill.right = bcgp_min(rtFill.right, rtView.right);

					if (rtApp.IsRectEmpty())
					{
						rtApp = rtFill;
					}
					else
					{
						rtApp.UnionRect(rtApp, rtFill);
					}
				}

				rtApp.OffsetRect(-xOffset, -yOffset);
				pApp->SetRectDraw(rtApp);

				arByInterval.Add(pApp);
			}

			// resort appointments in the view, if count of collection is great than 1
			const int c_Count = (int)arByInterval.GetSize ();
			if (c_Count == 0)
			{
				continue;
			}

			for (i = 0; i < c_Count; i++)
			{
				CBCGPAppointment* pApp1 = arByInterval[i];

				CRect rtApp1(pApp1->GetRectDraw());

				for (int j = 0; j < i; j++)
				{
					CBCGPAppointment* pApp2 = arByInterval[j];

					CRect rtApp2(pApp2->GetRectDraw());

					CRect rtInter;
					if (rtInter.IntersectRect (rtApp1, rtApp2))
					{
						rtApp1.top    = rtApp2.top;
						rtApp1.bottom = rtApp2.bottom;
						rtApp1.OffsetRect (0, m_nRowHeight);

						j = 0;
					}
				}

				pApp1->SetRectDraw(rtApp1);
			}

			for (i = 0; i < c_Count; i++)
			{
				CBCGPAppointment* pApp1 = arByInterval[i];
				CRect rtApp1(pApp1->GetRectDraw());

				rtApp1.left = bcgp_max(rtApp1.left, m_rectApps.left);
				if (rtApp1.left < rtApp1.right)
				{
					pApp1->SetRectDraw(rtApp1);
				}
			}
		}
	}

	COleDateTime date(m_DateStart);

	int nDay = 0;
	for (nDay = 0; nDay < nDays; nDay ++)
	{
		CheckVisibleAppointments (date, m_rectApps, nResCount > 1);

		date += spanDay;
	}

	CheckVisibleUpDownIcons(nResCount > 1);

	m_bUpdateToolTipInfo = TRUE;
}

void CBCGPPlannerViewSchedule::AddUpDownRect(BYTE nType, const CRect& rect)
{
	DWORD dwDrawFlags = GetDrawFlags ();
/*
	BOOL bDaysUpDown  = (dwDrawFlags & BCGP_PLANNER_DRAW_VIEW_DAYS_UPDOWN) == 
			BCGP_PLANNER_DRAW_VIEW_DAYS_UPDOWN;
*/
	CSize sz (GetPlanner ()->GetUpDownIconSize ());

	int offsetY = 0;
	if ((dwDrawFlags & BCGP_PLANNER_DRAW_VIEW_DAYS_UPDOWN_VCENTER) == 
			BCGP_PLANNER_DRAW_VIEW_DAYS_UPDOWN_VCENTER)
	{
		offsetY = (m_nRowHeight - sz.cy) / 2;
	}

	if ((nType & 0x01) == 0x01)
	{
		CRect rt (CPoint(0, 0), sz);

		rt.OffsetRect (rect.right - sz.cx, rect.top + offsetY);

		BOOL bAdd = TRUE;
		for (int i = 0; i < (int)m_UpRects.GetSize(); i++)
		{
			if (m_UpRects[i] == rt)
			{
				bAdd = FALSE;
				break;
			}
		}

		if (bAdd)
		{
 			m_UpRects.Add (rt);
		}
	}

	if ((nType & 0x02) == 0x02)
	{
		CRect rt (CPoint(0, 0), sz);

		rt.OffsetRect (rect.right - sz.cx, rect.bottom - sz.cy - offsetY - 1);

		BOOL bAdd = TRUE;
		for (int i = 0; i < (int)m_DownRects.GetSize(); i++)
		{
			if (m_DownRects[i] == rt)
			{
				bAdd = FALSE;
				break;
			}
		}

		if (bAdd)
		{
 			m_DownRects.Add (rt);
		}
	}
}

void CBCGPPlannerViewSchedule::CheckVisibleAppointments(const COleDateTime& date, const CRect& rect, 
	BOOL bFullVisible)
{
	ASSERT (date != COleDateTime ());

	XBCGPAppointmentArray& arQueryApps = GetQueryedAppointments ();
	XBCGPAppointmentArray& arDragApps = GetDragedAppointments ();

	if (arQueryApps.GetSize () == 0 && arDragApps.GetSize () == 0)
	{
		return;
	}

	BOOL bDragDrop        = IsDragDrop ();
	DROPEFFECT dragEffect = GetDragEffect ();
	BOOL bDragMatch       = IsCaptureMatched ();	

	bDragDrop = !bDragDrop || 
		(bDragDrop && ((dragEffect & DROPEFFECT_COPY) == DROPEFFECT_COPY && bDragMatch) || 
		!bDragMatch);
	bDragDrop = bDragDrop && arDragApps.GetSize () > 0;

	const BOOL bSelect = date != COleDateTime ();

	int nResCount = (int)m_Resources.GetSize();
	if (nResCount == 0)
	{
		nResCount = 1;
	}

	for (int nApp = 0; nApp < 2; nApp++)
	{
		if (!bDragDrop && nApp == 0)
		{
			continue;
		}

		XBCGPAppointmentArray& arApps = nApp == 0 ? arDragApps : arQueryApps;

		for (int nRes = 0; nRes < nResCount; nRes++)
		{
			XResource* pResource = NULL;
			CRect rectRes(rect);

			if (nResCount > 1)
			{
				pResource = &m_Resources[nRes];

				rectRes.top = pResource->m_Rects[0].top - m_nScrollOffset;
				rectRes.bottom = pResource->m_Rects[0].bottom - m_nScrollOffset;

				rectRes.IntersectRect (rectRes, rect);
			}

			for (int i = 0; i < (int)arApps.GetSize (); i++)
			{
				CBCGPAppointment* pApp = arApps [i];

				if (pResource != NULL && pResource->m_ResourceID != pApp->GetResourceID())
				{
					continue;
				}

				if (bSelect)
				{
					if (!IsAppointmentInDate (*pApp, date))
					{
						continue;
					}
				}


				CRect rtDraw (pApp->GetRectDraw (date));

				CRect rtInter;
				rtInter.IntersectRect (rtDraw, rectRes);

				pApp->SetVisibleDraw ((!bFullVisible && rtInter.Height () >= 2) || 
					(bFullVisible && rtInter.Height () == rtDraw.Height () && 
					 rtInter.bottom < rect.bottom), date);
			}
		}
	}
}

void CBCGPPlannerViewSchedule::CheckVisibleUpDownIcons(BOOL bFullVisible)
{
	ClearVisibleUpDownIcons ();

	CSize sz (GetPlanner ()->GetUpDownIconSize ());

	if (sz == CSize (0, 0))
	{
		return;
	}

	XBCGPAppointmentArray& arQueryApps = GetQueryedAppointments ();
	XBCGPAppointmentArray& arDragApps  = GetDragedAppointments ();

	if (arQueryApps.GetSize () == 0 && arDragApps.GetSize () == 0)
	{
		return;
	}

	BOOL bDragDrop        = IsDragDrop ();
	DROPEFFECT dragEffect = GetDragEffect ();
	BOOL bDragMatch       = IsCaptureMatched ();

	bDragDrop = !bDragDrop || 
		(bDragDrop && ((dragEffect & DROPEFFECT_COPY) == DROPEFFECT_COPY && bDragMatch) || 
		!bDragMatch);
	bDragDrop = bDragDrop && arDragApps.GetSize () > 0;

	const int nDays = GetViewDuration ();

	COleDateTime date (m_DateStart);

	int nResCount = (int)m_Resources.GetSize();
	if (nResCount == 0)
	{
		nResCount = 1;
	}

	for (int nDay = 0; nDay < nDays; nDay++)
	{
		CRect rect (m_ViewRects [nDay]);

		rect.OffsetRect(-m_nScrollHorzOffset, 0);

		if (rect.IntersectRect(m_rectApps, rect))
		{
			for (int nApp = 0; nApp < 2; nApp++)
			{
				XBCGPAppointmentArray& arApps = nApp == 1 ? arDragApps : arQueryApps;

				if (nApp == 1)
				{
					bDragDrop = bDragDrop && arDragApps.GetSize () > 0;
				}

				if (arApps.GetSize () == 0)
				{
					continue;
				}

				for (int i = 0; i < (int)arApps.GetSize (); i++)
				{
					CBCGPAppointment* pApp = arApps [i];
					if (pApp == NULL)
					{
						continue;
					}

					if (!IsAppointmentInDate (*pApp, date))
					{
						continue;
					}

					if (!pApp->IsVisibleDraw (date))
					{
						BYTE res = 0;

						if (nResCount > 1)
						{
							int nRes = FindResourceIndexByID(pApp->GetResourceID());
							if (nRes != -1)
							{
								rect = m_Resources[nRes].m_Rects[nDay];
								rect.OffsetRect(-m_nScrollHorzOffset, -m_nScrollOffset);

								if (!rect.IntersectRect(rect, m_rectApps))
								{
									continue;
								}
							}
						}

						BOOL bAdd = FALSE;

						if (bDragDrop && dragEffect != DROPEFFECT_NONE && pApp->IsSelected () &&
							nApp == 0)
						{
							if ((dragEffect & DROPEFFECT_COPY) == DROPEFFECT_COPY || bDragMatch)
							{
								bAdd = TRUE;
							}
						}
						else
						{
							bAdd = TRUE;
						}

						if (bAdd)
						{
							if (!rect.IsRectNull ())
							{
								CRect rectDraw (pApp->GetRectDraw (date));

								if (rectDraw.bottom <= rect.top)
								{
									res |= 0x01;
								}
								else if (rect.bottom <= (rectDraw.top + 1))
								{
									res |= 0x02;
								}

								if (bFullVisible)
								{
									if (rectDraw.top <= rect.top)
									{
										res |= 0x01;
									}
									else if (rect.bottom <= (rectDraw.bottom + 1))
									{
										res |= 0x02;
									}
								}
							}
						}

						if (res != 0)
						{
							AddUpDownRect (res, rect);
						}
					}
				}
			}
		}

		date += COleDateTimeSpan (1, 0, 0, 0);
	}
}

void CBCGPPlannerViewSchedule::OnDrawUpDownIcons (CDC* pDC)
{
	CBCGPPlannerView::OnDrawUpDownIcons (pDC);
}

void CBCGPPlannerViewSchedule::OnPaint (CDC* pDC, const CRect& rectClient)
{
	ASSERT_VALID (pDC);

	if (!m_rectResourcesBar.IsRectEmpty())
	{
		OnDrawResourcesBar(pDC, m_rectResourcesBar);
	}

	{
		CRgn rgn;
		rgn.CreateRectRgn (m_rectApps.left, rectClient.top, m_rectApps.right, m_rectApps.bottom);

		pDC->SelectClipRgn (&rgn);

		OnDrawClient (pDC, m_rectApps);

		int right = m_ViewRects.GetSize() > 0
						? m_ViewRects[m_ViewRects.GetSize() - 1].right
						: m_rectApps.right;

		if (m_nHeaderHeight != 0)
		{
			CRect rtHeader (rectClient);
			rtHeader.left   = m_rectApps.left;
			rtHeader.bottom = rtHeader.top + m_nHeaderHeight;

			if (right < m_rectApps.right)
			{
				rtHeader.right = right + 1;
			}

			OnDrawHeader (pDC, rtHeader);
		}

		if (!m_rectTimeBar.IsRectEmpty ())
		{
			OnDrawTimeBar (pDC, m_rectTimeBar, IsCurrentTimeVisible ());
		}

		if (right < m_rectApps.right)
		{
			pDC->SelectClipRgn (NULL);

			CRect rtDuration(m_rectApps);
			rgn.DeleteObject();
			rgn.CreateRectRgn (rtDuration.left, rectClient.top, right, rtDuration.bottom);

			pDC->SelectClipRgn (&rgn);
		}

		OnDrawAppointmentsDuration (pDC);

		pDC->SelectClipRgn (NULL);
	}

	{
		CRgn rgn;
		rgn.CreateRectRgn (m_rectApps.left, m_rectApps.top, m_rectApps.right, m_rectApps.bottom);

		pDC->SelectClipRgn (&rgn);
		
		OnDrawAppointments (pDC, m_rectApps);

		pDC->SelectClipRgn (NULL);
	}

	OnDrawUpDownIcons (pDC);

	InitToolTipInfo ();
}

void CBCGPPlannerViewSchedule::OnDrawClient (CDC* pDC, const CRect& rect)
{
	ASSERT_VALID (pDC);

	CRect rectFill (rect);

	const int xOffset = m_nScrollHorzOffset;
	const int yOffset = m_nScrollOffset;

	const BOOL bResources = m_Resources.GetSize() > 0;
	const int nDays = GetViewDuration ();
	const int nResCount = bResources ? (int)m_Resources.GetSize() : 1;

	const UINT nResourceID = GetPlanner()->GetCurrentResourceID();

	int nFirstWorkingHour   = GetFirstWorkingHour (bResources);
	int nFirstWorkingMinute = GetFirstWorkingMinute (bResources);
	int nLastWorkingHour    = GetLastWorkingHour (bResources);
	int nLastWorkingMinute  = GetLastWorkingMinute (bResources);

	const BOOL bDaysView = IsScheduleTimeDeltaDay();

	const int nMinuts = CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ());
	const int nCount = 60 / nMinuts;

	CPen     penHour[2];
	CPen     penHalfHour[2];

	for (int i = 0; i < 2; i++)
	{
		penHour[i].CreatePen (PS_SOLID, 0,
			GetHourLineColor (i == 0 /* Working */, TRUE));

		penHalfHour[i].CreatePen (PS_SOLID, 0, 
			GetHourLineColor (i == 0 /* Working */, FALSE));
	}

	XBCGPPlannerWorkingParameters WorkingParameters(this);
	COLORREF DefaultWorkingColorFill = visualManager->GetPlannerWorkColor ();

	CRect rectPlanner(rect);
	rectPlanner.top -= (m_nHeaderHeight + m_rectTimeBar.Height());

	OnFillPlanner (pDC, rectPlanner, FALSE /* Non-working */);

	CBrush brHilite (visualManager->GetPlannerSelectionColor (this));

	CPen penBlack (PS_SOLID, 0, visualManager->GetPlannerSeparatorColor (this));
	CPen* pOldPen = pDC->SelectObject (&penBlack);

	COleDateTime dtStart (GetDateStart ());

	BOOL bShowSelection = TRUE;
/*
	BOOL bShowSelection = !((m_Selection[0].GetHour ()   == 0 &&
							 m_Selection[0].GetMinute () == 0 &&
							 m_Selection[0].GetSecond () == 0 &&
							 m_Selection[1].GetHour ()   == 23 &&
							 m_Selection[1].GetMinute () == 59 &&
							 m_Selection[1].GetSecond () == 59) ||
							(m_Selection[1].GetHour ()   == 0 &&
							 m_Selection[1].GetMinute () == 0 &&
							 m_Selection[1].GetSecond () == 0 &&
							 m_Selection[0].GetHour ()   == 23 &&
							 m_Selection[0].GetMinute () == 59 &&
							 m_Selection[0].GetSecond () == 59));

	BOOL bIsDrawDuration = 
		(GetDrawFlags () & BCGP_PLANNER_DRAW_VIEW_NO_DURATION) == 0;
	const int nDurationBarWidth = globalUtils.ScaleByDPI(BCGP_PLANNER_DURATION_BAR_WIDTH);
	const int nDurationWidth = bIsDrawDuration ? nDurationBarWidth + 1 : 0;
*/
	XResource res;
	if (!bResources)
	{
		res.m_ResourceID = nResourceID;
		res.m_Rects.SetSize(m_ViewRects.GetSize());

		for (int i = 0; i < (int)m_ViewRects.GetSize(); i++)
		{
			res.m_Rects[i] = m_ViewRects[i];
		}
	}

	visualManager->PreparePlannerBackItem (FALSE, FALSE);

	int nDay = 0;
	for (nDay = 0; nDay < nDays; nDay++)
	{
		int nWD = dtStart.GetDayOfWeek ();
		BOOL bWeekEnd = nWD == 1 || nWD == 7;

		for (int nRes = 0; nRes < nResCount; nRes++)
		{
			XResource* pResource = bResources ? &m_Resources[nRes] : &res;

			int nFirstHour   = nFirstWorkingHour;
			int nFirstMinute = nFirstWorkingMinute;
			int nLastHour    = nLastWorkingHour;
			int nLastMinute  = nLastWorkingMinute;

			if (bResources && pResource->m_WorkStart < pResource->m_WorkEnd)
			{
				nFirstHour   = pResource->m_WorkStart.GetHour ();
				nFirstMinute = pResource->m_WorkStart.GetMinute ();
				nLastHour    = pResource->m_WorkEnd.GetHour ();
				nLastMinute  = pResource->m_WorkEnd.GetMinute ();
			}

			const int iStart = GetFirstViewHour() * nCount;
			const int iEnd   = iStart + GetViewHours() * nCount;

			const int iWorkStart = nFirstHour * nCount + (int)(nFirstMinute / nMinuts);
			const int iWorkEnd   = nLastHour * nCount + (int)(nLastMinute / nMinuts);

			rectFill = pResource->m_Rects[nDay];
			rectFill.OffsetRect(-xOffset, nResCount == 1 ? 0 : -yOffset);
			
			rectFill.right = rectFill.left + m_nColWidth;

			if (!bDaysView)
			{
				rectFill.left += 1;
			}

			BCGP_PLANNER_WORKING_STATUS AllPeriodWorkingStatus = 
				GetWorkingPeriodParameters (pResource->m_ResourceID, dtStart + COleDateTimeSpan (0, 0, iStart * nMinuts, 0), dtStart + COleDateTimeSpan (0, 0, (iEnd * nMinuts) - 1, 59), WorkingParameters); 
			BCGP_PLANNER_WORKING_STATUS SpecificPeriodWorkingStatus = AllPeriodWorkingStatus; 

			for (int iStep = iStart; iStep < iEnd; iStep++)
			{
				if (rectFill.right < m_rectApps.left)
				{
					rectFill.OffsetRect (m_nColWidth, 0);
					continue;
				}
				else if (rectFill.left > m_rectApps.right)
				{
					continue;
				}

				BOOL bIsWork = TRUE;
				if (AllPeriodWorkingStatus == BCGP_PLANNER_WORKING_STATUS_UNKNOWN)
				{ // We don't know for the day -> we should see for the period
					COleDateTime CurrentPeriodStart = dtStart + COleDateTimeSpan (0, 0, iStep * nMinuts, 0);
					COleDateTime CurrentPeriodEnd = CurrentPeriodStart + COleDateTimeSpan (0, 0, nMinuts - 1, 59);
					SpecificPeriodWorkingStatus = GetWorkingPeriodParameters (pResource->m_ResourceID, CurrentPeriodStart, CurrentPeriodEnd, WorkingParameters); 
				}

				switch (SpecificPeriodWorkingStatus)
				{
				case BCGP_PLANNER_WORKING_STATUS_ISNOTWORKING: // not a working period
					bIsWork = FALSE;
					break;
				case BCGP_PLANNER_WORKING_STATUS_ISWORKING: // forced to be a working period (we do not control working hours)
					bIsWork = TRUE;
					break;
				case BCGP_PLANNER_WORKING_STATUS_ISNORMALWORKINGDAY: // regular working day without control of week-end (so it may be a week-end day !)
					bIsWork = (iWorkStart <= iStep && iStep < iWorkEnd);
					break;
				case BCGP_PLANNER_WORKING_STATUS_ISNORMALWORKINGDAYINWEEK: // regular working day in a week (the week end is not a working day)
				default: // could not determine if period is working or not so we calculate as "standard"
					bIsWork = !bWeekEnd && (iWorkStart <= iStep && iStep < iWorkEnd);
				}

				if (pResource->m_ResourceID != nResourceID ||
					!IsDateInSelection (dtStart + COleDateTimeSpan (0, (iStep * nMinuts) / 60, (iStep * nMinuts) % 60, 0)) ||
					!bShowSelection)
				{
					if (bIsWork)
					{
						if (WorkingParameters.m_clrWorking != CLR_DEFAULT)
						{
							CBrush brush(WorkingParameters.m_clrWorking);
							pDC->FillRect (rectFill, &brush);
						}
						else
						{
							OnFillPlanner (pDC, rectFill, TRUE /* Working */);
						}
					}
					else
					{ // IF non working color is different from default non working color -> we should draw it with new color..
						if ((WorkingParameters.m_clrNonWorking != CLR_DEFAULT) && 
							(WorkingParameters.m_clrNonWorking != DefaultWorkingColorFill))
						{
							CBrush brush(WorkingParameters.m_clrNonWorking);
							pDC->FillRect (rectFill, &brush);
						}
					}
				}
				else
				{
					pDC->FillRect (rectFill, &brHilite);
				}

				if (!WorkingParameters.m_strText.IsEmpty())
				{
					COLORREF clrTextOld = CLR_DEFAULT;
					if (WorkingParameters.m_clrText != CLR_DEFAULT)
					{
						clrTextOld = pDC->SetTextColor (WorkingParameters.m_clrText);
					}

					CRect rectText=rectFill;
					rectText.DeflateRect (WorkingParameters.m_Margins);
					pDC->DrawText (WorkingParameters.m_strText, rectText, DT_CENTER | DT_VCENTER | DT_END_ELLIPSIS);

					if (clrTextOld != CLR_DEFAULT)
					{
						pDC->SetTextColor (clrTextOld);
					}
				}

				int nPenIndex = bIsWork ? 0 : 1;

				if (rectFill.right == (m_ViewRects[nDay].right - m_nScrollHorzOffset))
				{
					pDC->SelectObject (&penBlack);

					pDC->MoveTo (rectFill.right - 1, rectFill.top);
					pDC->LineTo (rectFill.right - 1, rectFill.bottom);

					pDC->MoveTo (rectFill.right, rectFill.top);
					pDC->LineTo (rectFill.right, rectFill.bottom);
				}
				else if (!bDaysView)
				{
					pDC->SelectObject (((iStep + 1) % nCount == 0) ? 
						&penHour [nPenIndex] : &penHalfHour [nPenIndex]);
				}

				if (!bDaysView && rectFill.right > m_rectApps.left)
				{
					pDC->MoveTo (rectFill.right, rectFill.top);
					pDC->LineTo (rectFill.right, rectFill.bottom);
				}

				rectFill.OffsetRect (m_nColWidth, 0);
			}

			if (nResCount > 1)
			{
				pDC->SelectObject (&penBlack);

				CRect rectRes(m_rectApps);
				rectRes.top = pResource->m_Rects[0].top;
				rectRes.OffsetRect(0, -m_nScrollOffset);

				if ((pResource->m_Rects[nDays - 1].right - m_nScrollHorzOffset) < m_rectApps.right)
				{
					rectRes.right = (pResource->m_Rects[nDays - 1].right - m_nScrollHorzOffset);
				}

				if (rectRes.top >= m_rectApps.top)
				{
					pDC->MoveTo(rectRes.left, rectRes.top);
					pDC->LineTo(rectRes.right, rectRes.top);
				}
			}
		}

		dtStart += COleDateTimeSpan (1, 0, 0, 0);
	}

	pDC->SelectObject (pOldPen);
}

void CBCGPPlannerViewSchedule::DrawHeader (CDC* pDC, const CRect& rect, int dxColumn)
{
	ASSERT_VALID (pDC);

	const int right = m_ViewRects[m_ViewRects.GetSize() - 1].right;

	CRect rectHeader(rect);
	rectHeader.left -= 1;
	if (right < m_rectApps.right)
	{
		rectHeader.right = right + 1;
	}

	visualManager->OnDrawPlannerHeader (pDC, this, rectHeader);

	rectHeader.left = m_ViewRects[0].left;
	rectHeader.right = right;
	rectHeader.OffsetRect(-m_nScrollHorzOffset, 0);

	CRect rt (rectHeader);

	rt.right = rt.left + dxColumn;
	rt.DeflateRect (1, 0);

	while (rt.right < (rectHeader.right - 4))
	{
		CRect rectPane (
			CPoint (rt.right, rt.top + 2), CSize (2, rt.Height () - 4));

		visualManager->OnDrawPlannerHeaderPane (pDC, this, rectPane);

		rt.OffsetRect (dxColumn, 0);
	}
}

void CBCGPPlannerViewSchedule::OnDrawHeader (CDC* pDC, const CRect& rectHeader)
{
	ASSERT_VALID (pDC);

	const int nDays = GetViewDuration ();

	DrawHeader (pDC, rectHeader, m_ViewRects[0].Width ());

	CRect rectDayCaption (rectHeader);

	COleDateTime day (GetDateStart ());

	COleDateTime dayCurrent = COleDateTime::GetCurrentTime ();
	dayCurrent.SetDateTime (dayCurrent.GetYear (), dayCurrent.GetMonth (), 
		dayCurrent.GetDay (), 0, 0, 0);

	const DWORD dwFlags = GetDrawFlags ();
	const BOOL bCaptionBold = (dwFlags & BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_BOLD) ==
			BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_BOLD;
	const BOOL bCaptionBoldFC = !bCaptionBold && (dwFlags & BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_BOLD_FC) ==
			BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_BOLD_FC;
	const BOOL bCaptionEx = (dwFlags & BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_EXTENDED) == 
		BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_EXTENDED;
	const BOOL bHeaderBold = (dwFlags & BCGP_PLANNER_DRAW_VIEW_HEADER_BOLD) ==
			BCGP_PLANNER_DRAW_VIEW_HEADER_BOLD;
	const BOOL bHeaderBoldCur = !bHeaderBold && (dwFlags & BCGP_PLANNER_DRAW_VIEW_HEADER_BOLD_CUR) ==
			BCGP_PLANNER_DRAW_VIEW_HEADER_BOLD_CUR;
	const BOOL bHeaderEx = (dwFlags & BCGP_PLANNER_DRAW_VIEW_HEADER_EXTENDED) == 
		BCGP_PLANNER_DRAW_VIEW_HEADER_EXTENDED;	
	const BOOL bCompact = (dwFlags & BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_COMPACT) ==
			BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_COMPACT;
	const BOOL bCaptionWithHeader = (dwFlags & BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_WITH_HEADER) == 
		BCGP_PLANNER_DRAW_VIEW_CAPTION_DAY_WITH_HEADER;

	const BOOL bCanDrawCaptionDayWithHeader = visualManager->CanDrawCaptionDayWithHeader();

	HFONT hOldFont = SetCurrFont (pDC, bCaptionBold, bCaptionEx);

	for (int nDay = 0; nDay < nDays; nDay++)
	{
		rectDayCaption.left = m_ViewRects [nDay].left - m_nScrollHorzOffset;
		rectDayCaption.right = m_ViewRects [nDay].right - m_nScrollHorzOffset;

		if (rectDayCaption.right <= m_rectApps.left || m_rectApps.right <= rectDayCaption.left)
		{
			day += COleDateTimeSpan (1, 0, 0, 0);
			continue;
		}

		rectDayCaption.left = max(rectDayCaption.left, m_rectApps.left);
		rectDayCaption.right = min(rectDayCaption.right, m_rectApps.right);

		BOOL bToday = day == dayCurrent;
		BOOL bSelected = FALSE;//m_SelectionAllDay[nDay];

		HFONT hOldFontFC = NULL;
		if (bCaptionBoldFC && (bToday || day.GetDay() == 1))
		{
			hOldFontFC = SetCurrFont (pDC, TRUE, bCaptionEx);
		}

		if (!bCompact && !bCaptionWithHeader)
		{
			visualManager->PreparePlannerCaptionBackItem (TRUE);
			COLORREF clrText = OnFillPlannerCaption (
				pDC, rectDayCaption, bToday, bSelected, FALSE);

			DrawCaptionText (pDC, rectDayCaption, day, clrText);
		}
		else
		{
			CString strText;
			ConstructCaptionText (day, strText, GetCaptionFormat ());

			if (strText.Find (TCHAR('\n')) != -1)
			{
				CString strDate;

				if (bCaptionWithHeader)
				{
					COLORREF clrText = CLR_DEFAULT;
					if (!bCanDrawCaptionDayWithHeader)
					{
						visualManager->PreparePlannerCaptionBackItem (TRUE);
						clrText = OnFillPlannerCaption (
							pDC, rectDayCaption, bToday, bSelected, FALSE);
					}

					CRect rectDayCaptionOld(rectDayCaption);

					if ((bHeaderEx && bCaptionEx) || (!bHeaderEx && !bCaptionEx))
					{
						rectDayCaption.top += rectDayCaption.Height() / 2;
					}
					else if (bHeaderEx)
					{
						rectDayCaption.top += (rectDayCaption.Height() * 20) / 36;
					}
					else
					{
						rectDayCaption.top += (rectDayCaption.Height() * 4) / 9;
					}

					HFONT hOldFontPart = SetCurrFont (pDC, bCaptionBold || (bCaptionBoldFC && (bToday || day.GetDay() == 1)), bCaptionEx);

					if (bCanDrawCaptionDayWithHeader)
					{
						visualManager->PreparePlannerCaptionBackItem (FALSE);
						clrText = OnFillPlannerCaption (
							pDC, rectDayCaption, bToday, bSelected, FALSE);
					}

					AfxExtractSubString (strDate, strText, 0, TCHAR('\n'));
					DrawCaptionText (pDC, rectDayCaption, strDate, clrText, DT_LEFT);

					rectDayCaption.bottom = rectDayCaption.top;
					rectDayCaption.top = rectDayCaptionOld.top;

					SetCurrFont (pDC, bHeaderBold || (bHeaderBoldCur && bToday), bHeaderEx);

					if (bCanDrawCaptionDayWithHeader)
					{
						visualManager->PreparePlannerCaptionBackItem (TRUE);
						clrText = OnFillPlannerCaption (
							pDC, rectDayCaption, bToday, bSelected, FALSE);
					}

					AfxExtractSubString (strDate, strText, 1, TCHAR('\n'));
					DrawCaptionText (pDC, rectDayCaption, strDate, clrText, DT_LEFT);

					if (hOldFontPart != NULL)
					{
						::SelectObject (pDC->GetSafeHdc (), hOldFontPart);
					}

					rectDayCaption.bottom = rectDayCaptionOld.bottom;
				}
				else
				{
					visualManager->PreparePlannerCaptionBackItem (TRUE);
					COLORREF clrText = OnFillPlannerCaption (
						pDC, rectDayCaption, bToday, bSelected, FALSE);

					AfxExtractSubString (strDate, strText, 0, TCHAR('\n'));
					DrawCaptionText (pDC, rectDayCaption, strDate, clrText, DT_LEFT);

					HFONT hOldFontPart = NULL;
					if (bCaptionBoldFC && (day != dayCurrent && day.GetDay() == 1))
					{
						hOldFontPart = SetCurrFont (pDC, FALSE, bCaptionEx);
					}

					AfxExtractSubString (strDate, strText, 1, TCHAR('\n'));
					if (!strDate.IsEmpty())
					{
						DrawCaptionText (pDC, rectDayCaption, strDate, clrText, DT_CENTER);
					}

					if (hOldFontPart != NULL)
					{
						::SelectObject (pDC->GetSafeHdc (), hOldFontPart);
					}
				}
			}
			else
			{
				visualManager->PreparePlannerCaptionBackItem (TRUE);
				COLORREF clrText = OnFillPlannerCaption (
					pDC, rectDayCaption, bToday, bSelected, FALSE);

				DrawCaptionText (pDC, rectDayCaption, strText, clrText, DT_LEFT);
			}
		}

		if (hOldFontFC != NULL)
		{
			::SelectObject (pDC->GetSafeHdc (), hOldFontFC);
		}

		day += COleDateTimeSpan (1, 0, 0, 0);
	}

	if (hOldFont != NULL)
	{
		::SelectObject (pDC->GetSafeHdc (), hOldFont);
	}
}

void CBCGPPlannerViewSchedule::OnDrawResourcesBar (CDC* pDC, const CRect& rectBar)
{
	if (rectBar.IsRectEmpty ())
	{
		return;
	}

	ASSERT_VALID (pDC);

	visualManager->PreparePlannerCaptionBackItem (FALSE);
	visualManager->OnFillPlannerWeekBar (pDC, this, rectBar);

	CFont* pFont = CFont::FromHandle(GetFont ());
	ASSERT (pFont != NULL);

	CFont* pOldFont = pDC->SelectObject (pFont);

	COLORREF clrLine = visualManager->GetPlannerSeparatorColor (this);

	CPen penBlack (PS_SOLID, 0, clrLine);
	CPen* pOldPen = pDC->SelectObject (&penBlack);

	COLORREF clrOldText = pDC->GetTextColor();

	pDC->MoveTo(rectBar.right - 1, rectBar.top);
	pDC->LineTo(rectBar.right - 1, rectBar.bottom);

	pDC->MoveTo(rectBar.left, m_rectApps.top);
	pDC->LineTo(rectBar.right, m_rectApps.top);

	const int c_Count = (int)m_Resources.GetSize();
	for (int i = 0; i < c_Count; i++)
	{
		CRect rect(rectBar);
		rect.top = m_Resources[i].m_Rects[0].top;
		rect.bottom = m_Resources[i].m_Rects[0].bottom;

		if (c_Count > 1)
		{
			rect.OffsetRect(0, -m_nScrollOffset);
		}

		CRect rtFill(rect);
		rtFill.top = m_rectApps.top;
		rtFill.bottom = m_rectApps.bottom;
		rtFill.IntersectRect(rect, rtFill);

		if (rtFill.Height() <= 0)
		{
			continue;
		}

		COLORREF clrText = visualManager->OnFillPlannerCaption (pDC, this, CRect(0, 0, 0, 0)/*rtFill*/, FALSE, FALSE, TRUE, FALSE);

		if (m_rectApps.top < rtFill.top && rtFill.top < m_rectApps.bottom)
		{
			pDC->MoveTo(rtFill.left, rtFill.top);
			pDC->LineTo(rtFill.right, rtFill.top);
		}

		if ((rtFill.top + m_nRowHeight) < rtFill.bottom)
		{
			rtFill.bottom = rtFill.top + m_nRowHeight;
		}

		if (rtFill.Height() > 2)
		{
			pDC->SetTextColor (clrText);

			rtFill.left += 5;
			pDC->DrawText(m_Resources[i].m_Description, rtFill, DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
		}
	}

	pDC->SetTextColor (clrOldText);

	pDC->SelectObject(pOldPen);

	pDC->SelectObject(pOldFont);
}

void CBCGPPlannerViewSchedule::OnDrawTimeLine (CDC* pDC, const CRect& rect)
{
	visualManager->OnDrawPlannerTimeLine (pDC, this, rect);
}

void CBCGPPlannerViewSchedule::OnDrawTimeBar (CDC* pDC, const CRect& rectBar, 
										 BOOL bDrawTimeLine)
{
	ASSERT_VALID (pDC);

	BOOL b24Hours = CBCGPPlannerView::Is24Hours ();

	CString strAM;
	CString strPM;

	if (!b24Hours)
	{
		strAM = CBCGPPlannerView::GetTimeDesignator (TRUE);
		strAM.MakeLower ();
		strPM = CBCGPPlannerView::GetTimeDesignator (FALSE);
		strPM.MakeLower ();
	}

/*
	const BOOL bTimeLineFullWidth = (GetDrawFlags() & BCGP_PLANNER_DRAW_VIEW_TIMELINE_FULL_WIDTH) == 
					BCGP_PLANNER_DRAW_VIEW_TIMELINE_FULL_WIDTH;
*/
	const BOOL bTimeBarSimple = (GetDrawFlags() & BCGP_PLANNER_DRAW_VIEW_TIME_BAR_NO_MINUTS) == 
					BCGP_PLANNER_DRAW_VIEW_TIME_BAR_NO_MINUTS;

	COLORREF clrLine = globalData.clrBtnShadow;

	COLORREF clrText = visualManager->OnFillPlannerTimeBar (pDC, this, rectBar, clrLine);

	CRect rtBarHalf(rectBar);
	rtBarHalf.top = (rtBarHalf.top + rtBarHalf.bottom) / 2;

	OnFillPlanner (pDC, rtBarHalf, TRUE /* Working */);//pDC->FillSolidRect (rtBarHalf, visualManager->GetPlannerVi());

	COLORREF clrTextOld = pDC->SetTextColor (clrText);

	int x = rectBar.left - m_nScrollHorzOffset/* - 1*/; 

	CPen pen (PS_SOLID, 0, clrLine);
	CPen* pOldPen = pDC->SelectObject (&pen);

	const long nCount = 60 / CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ());

	const BOOL bDaysView = IsScheduleTimeDeltaDay();

	UINT nFormat = DT_SINGLELINE | DT_LEFT | (bDaysView ? DT_NOCLIP : 0);

	CFont* pFont = CFont::FromHandle(GetFont ());
	ASSERT (pFont != NULL);

	CFont* pOldFont = pDC->SelectObject (pFont);

	CFont fontBold;
	LOGFONT lf;

	if (!bTimeBarSimple)
	{
		pFont->GetLogFont (&lf);
		lf.lfHeight *= 2;
	}
	else
	{
		CFont* pFontBold = CFont::FromHandle(GetFont (FALSE, TRUE));
		pFontBold->GetLogFont (&lf);
	}

	fontBold.CreateFontIndirect (&lf);

	if (bDrawTimeLine)
	{
		CRect rectTimeLine(rectBar);

		COleDateTime dtCur(m_CurrentTime.GetYear (), m_CurrentTime.GetMonth (), m_CurrentTime.GetDay (), 0, 0, 0);

		COleDateTime dtStart(GetDateStart ());
		dtStart = COleDateTime (dtStart.GetYear (), dtStart.GetMonth (), dtStart.GetDay (), 0, 0, 0);

		int nIndex = (int)((dtCur - dtStart).GetDays());
		if (0 <= nIndex && nIndex < (int)m_ViewRects.GetSize())
		{
			rectTimeLine.right  = bcg_round(((m_CurrentTime.GetHour () - GetFirstViewHour()) * 60 + m_CurrentTime.GetMinute ()) * m_nColWidth / 
				(double)CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ())) + m_ViewRects[nIndex].left - m_nScrollHorzOffset + 1;

			rectTimeLine.bottom = rectTimeLine.top + rectBar.Height() / 2;
			rectTimeLine.left   = rectTimeLine.right - m_nRowHeight / (nCount > 1 ? 2 : 4);

			OnDrawTimeLine (pDC, rectTimeLine);
		}
	}

	const int cxDelta = globalUtils.ScaleByDPI(5);
	const int cyDelta = globalUtils.ScaleByDPI(5);

	int nMinDelta = max (pDC->GetTextExtent (_T("000")).cx, 18);
	CString strSeparator = CBCGPPlannerView::GetTimeSeparator ();

	int nStartDay = 0;
	int nEndDay = GetViewDuration() - 1;

	for (int day = nStartDay; day <= nEndDay; day++)
	{
		int bDrawFirstAM_PM = 0;

		int nStartHour = GetFirstViewHour();
		int nEndHour = GetLastViewHour();

		for (int i = nStartHour; i <= nEndHour; i++)
		{
			CRect rectHour(rectBar);
			rectHour.left   = x;
			rectHour.right  = x + m_nColWidth * nCount;
			rectHour.bottom = (rectHour.top + rectHour.bottom) / 2;

			if (rectHour.right < rectBar.left)
			{
				x = rectHour.right;
				continue;
			}

			if (nCount > 1/* && !bTimeBarSimple*/ && !bDaysView)
			{
				long nd = x + m_nColWidth;

				for (int j = 0; j < nCount - 1; j++)
				{
					if (nd > rectBar.left)
					{
						pDC->MoveTo (nd, rectHour.top + nMinDelta);
						pDC->LineTo (nd, rectHour.bottom);
					}

					nd += m_nColWidth;
				}
			}

			if (rectHour.right >= rectBar.left)
			{
				if (x > rectBar.left && !bDaysView)
				{
					pDC->MoveTo (x, rectHour.top + cyDelta);
					pDC->LineTo (x, rectHour.bottom);
				}

				x += m_nColWidth * nCount;

				if (i == nEndHour)
				{
					pDC->MoveTo (x - 1, rectHour.top);
					pDC->LineTo (x - 1, rectBar.bottom);

					pDC->MoveTo (x, rectHour.top);
					pDC->LineTo (x, rectBar.bottom);
				}
			}

			if (rectHour.left >= rectBar.left || rectHour.right > rectBar.left)
			{
				if (rectHour.left >= rectBar.left - 1)
				{
					bDrawFirstAM_PM++;
				}

				rectHour.left += cxDelta;

				CString str (_T("00"));

				int nHour = (i < 24) ? i : i - 24;

				BOOL bAM = nHour < 12; 

				if (!b24Hours)
				{
					if (nHour == 0 || nHour == 12)
					{
						nHour = 12;
					}
					else if (nHour > 12)
					{
						nHour -= 12;
					}
				}

				if (!bDaysView || ((i == nStartHour || i == nEndHour) && bDaysView))
				{
					if (nCount == 1 && !bTimeBarSimple)
					{
						str.Format (_T("%2d%s00"), nHour, (LPCTSTR)strSeparator);

						if (!b24Hours)
						{
							if (nHour == 12 || bDaysView)
							{
								str.Format (_T("%d %s"), nHour, bAM ? (LPCTSTR)strAM : (LPCTSTR)strPM);
							}
						}

						BOOL bRight = (i == nEndHour) && bDaysView;
						if (bRight)
						{
							rectHour.OffsetRect(-cxDelta, 0);
						}

						pDC->DrawText (str, rectHour, nFormat | (bRight ? (DT_RIGHT | DT_NOCLIP) : DT_LEFT));

						if (bRight)
						{
							rectHour.OffsetRect(cxDelta, 0);
						}
					}
					else
					{
						x = rectHour.right;

						rectHour.right = rectHour.left + m_nColWidth;

						BOOL bDrawMinuts = !bTimeBarSimple;
						if (!b24Hours)
						{
							if (nHour == 12 || bDrawFirstAM_PM == 1 || bDaysView)
							{
								bDrawMinuts = TRUE;
								str = bAM ? strAM : strPM;
							}
						}

						pDC->SelectObject (&fontBold);

						CString strHour;
						strHour.Format (_T("%d"), nHour);

						CRect rtText(rectHour);
						if (bDrawMinuts)
						{
							pDC->DrawText (strHour, rtText, nFormat | DT_CALCRECT);
						}

						if ((i == nEndHour) && bDaysView)
						{
							CRect rtText2(rectHour);

							if (bDrawMinuts)
							{
								rtText2.OffsetRect(-rtText.Width() - 3 * cxDelta, 0);
							}
							else
							{
								rtText2.OffsetRect(-2 * cxDelta, 0);
							}

							pDC->DrawText (strHour, rtText2, nFormat | DT_RIGHT);
						}
						else
						{
							pDC->DrawText (strHour, rectHour, nFormat);
						}

						pDC->SelectObject (pFont);

						if (bDrawMinuts)
						{
							CRect rectMin (rectHour);
							if ((i == nEndHour) && bDaysView)
							{
								rectMin.right -= 2 * cxDelta;
							}
							else
							{
								rectMin.left += rtText.Width() + cxDelta;
							}

							pDC->DrawText (str, rectMin, DT_SINGLELINE | ((i == nEndHour && bDaysView) ? DT_RIGHT : DT_LEFT) | (bDaysView ? DT_NOCLIP : 0));
						}

						rectHour.right = x;
					}
				}
			}

			x = rectHour.right;
		}
	}

	pDC->SelectObject (pOldFont);

	pDC->MoveTo (rtBarHalf.left, rtBarHalf.top);
	pDC->LineTo (rtBarHalf.right, rtBarHalf.top);

	pDC->MoveTo (rtBarHalf.left, rectBar.bottom);
	pDC->LineTo (rtBarHalf.right, rectBar.bottom);

	pDC->SelectObject (pOldPen);

	pDC->SetTextColor (clrTextOld);
}

void CBCGPPlannerViewSchedule::DrawAppointment (CDC* pDC, CBCGPAppointment* pApp, CBCGPAppointmentDrawStructEx* pDS)
{
	ASSERT_VALID (pDC);
	ASSERT_VALID (pApp);

	if (pDC == NULL || pDC->GetSafeHdc () == NULL || pApp == NULL || pDS == NULL)
	{
		return;
	}

	const DWORD dwDrawFlags = GetDrawFlags ();
	const BOOL bIsGradientFill      = (dwDrawFlags & BCGP_PLANNER_DRAW_APP_GRADIENT_FILL) ==
		BCGP_PLANNER_DRAW_APP_GRADIENT_FILL;
	const BOOL bIsRoundedCorners    = (dwDrawFlags & BCGP_PLANNER_DRAW_APP_ROUNDED_CORNERS) ==
		BCGP_PLANNER_DRAW_APP_ROUNDED_CORNERS;
	const BOOL bDrawDuration = (dwDrawFlags & BCGP_PLANNER_DRAW_APP_NO_DURATION) == 0;
	const BOOL bDrawDurationShape = ((dwDrawFlags & BCGP_PLANNER_DRAW_APP_DURATION_SHAPE) == 
		BCGP_PLANNER_DRAW_APP_DURATION_SHAPE) && bDrawDuration;
	const BOOL bDrawDurationAlternative = ((dwDrawFlags & BCGP_PLANNER_DRAW_APP_DURATION_MULTIDAY) == 
		BCGP_PLANNER_DRAW_APP_DURATION_MULTIDAY) && bDrawDuration;
	const BOOL bIsOverrideSelection = ((dwDrawFlags & BCGP_PLANNER_DRAW_APP_OVERRIDE_SELECTION) ==
		BCGP_PLANNER_DRAW_APP_OVERRIDE_SELECTION) || bDrawDurationShape;
	const BOOL bDrawBorder = ((dwDrawFlags & BCGP_PLANNER_DRAW_APP_NO_BORDER) == 0) ||
		!bIsOverrideSelection;
	const BOOL bFrameRgn = bIsRoundedCorners || bDrawDurationShape || !bDrawBorder;

	CRect rect (pDS->GetRect ());

	COleDateTime dtStart  (pApp->GetStart ());
	COleDateTime dtFinish (pApp->GetFinish ());

	CString      strStart  (pApp->m_strStart);
	CString      strFinish (pApp->m_strFinish);

	const BOOL bDrawShadow  = FALSE;//IsDrawAppsShadow () && !pApp->IsSelected () && (dwDrawFlags & BCGP_PLANNER_DRAW_APP_NO_SHADOW) == 0;
	const BOOL bAlternative = pApp->IsAllDay () || pApp->IsMultiDay ();
	const BOOL bEmpty       = (dtStart == dtFinish) && !bAlternative;
	const BOOL bSelected    = pApp->IsSelected ();
/*
	BOOL bAllBoders = TRUE;

	if (bAlternative)
	{
		bAllBoders = pDS->GetBorder () == CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_ALL;
	}
*/
	const int nDurationBarWidth = globalUtils.ScaleByDPI(CBCGPPlannerViewDay::BCGP_PLANNER_DURATION_BAR_WIDTH);
	const int nSpacingLeft = 0;//bDrawDuration ? 0 : 3;
	const int nSpacingRight = 1;//3;

	const DWORD dwDrawTimeFlags = GetDrawTimeFlags ();
	const BOOL bDrawTime = dwDrawTimeFlags & BCGP_PLANNER_DRAW_TIME;
	const BOOL bDrawTimeMulti = dwDrawTimeFlags & BCGP_PLANNER_DRAW_TIME_MULTI;
	const BOOL bDrawTimeSmart = dwDrawTimeFlags & BCGP_PLANNER_DRAW_TIME_SMART;
/*
	if (nSpacingLeft == 0)
	{
		rect.left--;
	}
*/
/*
	if (!bAlternative && !IsOneDay (dtStart, dtFinish))
	{
		if (pDS->GetBorder () == CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_START)
		{
			dtFinish  = COleDateTime (dtFinish.GetYear (), dtFinish.GetMonth (), dtFinish.GetDay (),
				0, 0, 0);
			strFinish = CBCGPAppointment::GetFormattedTimeString (dtFinish, TIME_NOSECONDS, FALSE);
			strFinish.MakeLower ();
		}
		else if (pDS->GetBorder () == CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_FINISH)
		{
			dtStart  = COleDateTime (dtFinish.GetYear (), dtFinish.GetMonth (), dtFinish.GetDay (),
				0, 0, 0);
			strStart = CBCGPAppointment::GetFormattedTimeString (dtStart, TIME_NOSECONDS, FALSE);
			strStart.MakeLower ();
		}
	}
*/
	CRgn rgn;
	CRect rectClip (0, 0, 0, 0);
	pDC->GetClipBox (rectClip);

	if (bFrameRgn)
	{
		int left  = nSpacingLeft;
		int right = nSpacingRight;
/*
		if (bAlternative)
		{
			if (!bAllBoders)
			{
				if ((pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_START) == 0)
				{
					left = 0;
				}

				if ((pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_FINISH) == 0)
				{
					right = 0;
				}
			}

			if (bSelected)
			{
				rect.InflateRect (0, 1);
			}
		}
		else
*/
		{
			if (bSelected)
			{
				if (!bIsOverrideSelection)
				{
					left  = 0;
					right = 0;
				}
				else
				{
					rect.InflateRect (0, 1);
				}
			}
		}

		rect.DeflateRect (left, 0, right, 0);

		if (bIsRoundedCorners)
		{
			rgn.CreateRoundRectRgn (rect.left, rect.top, rect.right + 1, rect.bottom + 1, 5, 5);
		}
		else
		{
			rgn.CreateRectRgn (rect.left, rect.top, rect.right, rect.bottom);
		}
/*
		if (bAlternative && bIsRoundedCorners)
		{
			if (!bAllBoders)
			{
				if ((pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_START) == 0)
				{
					CRgn rgnTmp;
					rgnTmp.CreateRectRgn (rect.left - 1, rect.top, rect.left + 3, rect.bottom);
					rgn.CombineRgn (&rgn, &rgnTmp, RGN_OR);
				}

				if ((pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_FINISH) == 0)
				{
					CRgn rgnTmp;
					rgnTmp.CreateRectRgn (rect.right - 4, rect.top, rect.right + 1, rect.bottom);
					rgn.CombineRgn (&rgn, &rgnTmp, RGN_OR);
				}
			}
		}
*/
		pDC->SelectClipRgn (&rgn, RGN_AND);

		if (bSelected)
		{
			if (/*bAlternative || */(/*!bAlternative && */bIsOverrideSelection))
			{
				rect.DeflateRect (0, 1);
			}
		}
	}
	else
	{
/*
		if (bAlternative)
		{
			rect.DeflateRect (3, 0);
		}
		else */if (!bSelected || bIsOverrideSelection)
		{
			if (bIsOverrideSelection)
			{
				rect.DeflateRect (3, 0);
			}
			else if (!bSelected)
			{
				rect.DeflateRect (5, 0);
			}
		}
	}

	CRect rectDuration (0, 0, 0, 0);
	if ((bDrawDuration && !bAlternative) || (bDrawDurationAlternative && bAlternative))
	{
		rectDuration = rect;
		rectDuration.right = rectDuration.left + nDurationBarWidth;
	}

	CBCGPDrawManager dm (*pDC);

	COLORREF clrBack1   = pApp->GetBackgroundColor ();
	COLORREF clrBack2   = clrBack1;
	COLORREF clrText    = pApp->GetForegroundColor ();
	COLORREF clrFrame1  = globalData.clrWindowFrame;
	COLORREF clrFrame2  = clrFrame1;
	COLORREF clrTextOld = CLR_DEFAULT;

	visualManager->GetPlannerAppointmentColors (this, bSelected, !bAlternative, 
		dwDrawFlags, clrBack1, clrBack2, clrFrame1, clrFrame2, clrText);

	if (bSelected && !bIsOverrideSelection)
	{
		CBrush br (clrBack1);
		pDC->FillRect (rect, &br);
/*
		if (!bAlternative)
		{
			rect.DeflateRect (bFrameRgn ? 3 : 5, 0);
		}
*/
	}
	else
	{
		if (bIsGradientFill)
		{
			dm.FillGradient (rect, clrBack1, clrBack2);
		}
		else if (clrBack1 != CLR_DEFAULT || bAlternative)
		{
			CBrush br (clrBack1);
			pDC->FillRect (rect, &br);
		}
	}

	if (bDrawShadow)
	{
		CRect rt (rect);
		rt.bottom++;

		dm.DrawShadow (rt, CBCGPAppointment::BCGP_PLANNER_SHADOW_DEPTH, 100, 75);
	}

	if (!rectDuration.IsRectNull ())
	{
		CRect rt (rect);
		rt.right = rectDuration.right;

		if (bDrawDurationShape)
		{
			COLORREF clr = bIsRoundedCorners
							? pApp->GetDurationColor ()
							: visualManager->GetPlannerAppointmentDurationColor(this, pApp);
			if (clr != CLR_DEFAULT)
			{
				CBrush br (clr);

				CRect rt1 (rectDuration);
				if (bDrawBorder)
				{
					rt1.right++;
				}
				pDC->FillRect (rt1, &br);
			}

			rect.left = rt.right + 1;
		}
		else
		{
			pDC->FillRect (rt, &globalData.brWindow);

			CBrush br (visualManager->GetPlannerAppointmentDurationColor(this, pApp));

			CPen penBlack (PS_SOLID, 0, clrFrame1);
			CPen* pOldPen = pDC->SelectObject (&penBlack);

			CRect rt1 (rectDuration);
			pDC->FillRect (rt1, &br);

			rt1.right++;
			pDC->Draw3dRect (rt1, clrFrame1, clrFrame1);

			if (bDrawBorder)
			{
				pDC->MoveTo (rt.right, rt.top);
				pDC->LineTo (rt.right, rt.bottom);
			}
			else
			{
				rt1 = rt;
				rt1.right++;

				if (!bFrameRgn)
				{
					pDC->Draw3dRect (rt1, clrFrame1, clrFrame1);
				}
				else
				{
					CRgn rgnD;
					rgnD.CreateRectRgnIndirect(rt1);
					rgnD.CombineRgn(&rgn, &rgnD, RGN_AND);

					CBrush brFrame(clrFrame1);
					pDC->FrameRgn (&rgnD, &brFrame, 1, 1);
				}
			}

			pDC->SelectObject (pOldPen);

			rect.left = rt.right + 1;
		}
	}

	clrTextOld = pDC->SetTextColor (clrText);

	BOOL bCancelDraw = FALSE;
	BOOL bToolTipNeeded = FALSE;

	CSize szImage (CBCGPPlannerManagerCtrl::GetImageSize ());

	if (bDrawBorder || bSelected)
	{
		if (bFrameRgn)
		{
			int nWidth = 1;

			if (bSelected)
			{
				if (bIsOverrideSelection || bAlternative)
				{
					nWidth = 2;
				}
			}

			CBrush br (clrFrame2);
			pDC->FrameRgn (&rgn, &br, nWidth, nWidth);
		}
		else
		{
			if (bSelected && bIsOverrideSelection)
			{
				pDC->Draw3dRect (rect, clrFrame2, clrFrame2);
			}
			else if (bDrawBorder)
			{
				pDC->Draw3dRect (rect, clrFrame1, clrFrame1);
			}

			if (bSelected)
			{
				if (!bFrameRgn)
				{
					CRect rt (rect);
					if (bIsOverrideSelection)
					{
						rt.DeflateRect (1, 1);
					}

					pDC->Draw3dRect (rt, clrFrame2, clrFrame2);
				}
			}
/*
			if (bAlternative)
			{
				CPen pen (PS_SOLID, 0, clrFrame2);
				CPen* pOldPen = pDC->SelectObject (&pen);

				if (bAllBoders)
				{
					pDC->Draw3dRect (rect, clrFrame2, clrFrame2);
				}
				else
				{
					pDC->MoveTo (rect.left , rect.top);
					pDC->LineTo (rect.right, rect.top);

					pDC->MoveTo (rect.left , rect.bottom - 1);
					pDC->LineTo (rect.right, rect.bottom - 1);

					if (pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_START)
					{
						pDC->MoveTo (rect.left, rect.top);
						pDC->LineTo (rect.left, rect.bottom);
					}

					if (pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_FINISH)
					{
						pDC->MoveTo (rect.right - 1, rect.top);
						pDC->LineTo (rect.right - 1, rect.bottom);
					}
				}

				if (bSelected)
				{
					CPoint pt (rect.left, rect.right);

					if ((pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_START) == 0)
					{
						pt.x--;
					}

					if ((pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_FINISH) == 0)
					{
						pt.y++;
					}

					pDC->MoveTo (pt.x, rect.top - 1);
					pDC->LineTo (pt.y, rect.top - 1);

					pDC->MoveTo (pt.x, rect.bottom);
					pDC->LineTo (pt.y, rect.bottom);
				}

				pDC->SelectObject (pOldPen);
			}
			else
			{
				if (bSelected && bIsOverrideSelection)
				{
					CRect rt (rect);

					rt.DeflateRect (1, 0);
					pDC->Draw3dRect (rt, clrFrame2, clrFrame2);

					rt.InflateRect  (1, 1);
					pDC->Draw3dRect (rt, clrFrame2, clrFrame2);
				}
			}
*/
		}
	}

	if (bFrameRgn || bIsOverrideSelection)
	{
		rect.DeflateRect (/*bAlternative
							? 1
							: */bFrameRgn
								? 3
								: 2, 0);
	}

	CSize szClock (GetClockIconSize ());

	if (!pApp->IsAllDay () && (bDrawTime || bDrawTimeMulti))
	{
		if ((IsDrawTimeAsIcons () || 
			(bAlternative && (dwDrawFlags & BCGP_PLANNER_DRAW_APP_NO_MULTIDAY_CLOCKS) == 0)) && 
			szClock != CSize (0, 0))
		{
			int top = (rect.Height () - szClock.cy) / 2;

			if (bAlternative && bDrawTimeMulti)
			{
				if (pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_START)
				{
					if (dtStart.GetHour () != 0 || dtStart.GetMinute () != 0 ||
						dtStart.GetSecond () != 0 || !bDrawTimeSmart)
					{
						if (nSpacingLeft == 0)
						{
							rect.left++;
						}

						DrawClockIcon (pDC, CPoint (rect.left + 2, rect.top + top), 
							dtStart);

						rect.left += szClock.cx + 4;
					}
				}

				if (pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_FINISH)
				{
					if (dtFinish.GetHour () != 0 || dtFinish.GetMinute () != 0 ||
						dtFinish.GetSecond () != 0 || !bDrawTimeSmart)
					{
						DrawClockIcon (pDC, CPoint (rect.right - szClock.cx - 2, rect.top + top), 
							dtFinish);

						rect.right -= (szClock.cx + 4);
					}
				}
			}
			else if (!bAlternative && bDrawTime)
			{
				if (rect.Width () >= szClock.cx * 3)
				{
					DrawClockIcon (pDC, CPoint (rect.left + 1, rect.top + top), 
						dtStart);

					rect.left += szClock.cx + 3;

					if (rect.Width () >= szClock.cx * 4 && IsDrawTimeFinish ())
					{
						DrawClockIcon (pDC, CPoint (rect.left + 1, rect.top + top), 
							dtFinish);

						rect.left += szClock.cx + 3;
					}
				}

				bToolTipNeeded = TRUE;
			}
		}
		else if (!strStart.IsEmpty ())
		{
			COLORREF clrTime = visualManager->GetPlannerAppointmentTimeColor(this, 
				bSelected, !bAlternative, dwDrawFlags);
			if (clrTime != CLR_DEFAULT)
			{
				pDC->SetTextColor (clrTime);
			}

			COleDateTime dt (2005, 1, 1, 23, 59, 0);

			CString strFormat (_T("%H"));
			strFormat += GetTimeSeparator () + _T("%M");
			if (!Is24Hours ())
			{
				strFormat += _T("%p");
			}
			CString strDT (dt.Format (strFormat));
			strDT.MakeLower ();

			CSize szSpace (pDC->GetTextExtent (_T(" ")));
			CSize szDT (pDC->GetTextExtent (strDT));

			if ((!bAlternative && rect.Width () >= (szDT.cx + szSpace.cx) * 2) ||
				( bAlternative && rect.Width () >= (szDT.cx + szSpace.cx) * 3))
			{
				if (bAlternative && bDrawTimeMulti)
				{
					if (pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_START)
					{
						if (dtStart.GetHour () != 0 || dtStart.GetMinute () != 0 ||
							dtStart.GetSecond () != 0 || !bDrawTimeSmart)
						{
							if (nSpacingLeft == 0)
							{
								rect.left++;
							}

							CRect rectText (rect);
							rectText.DeflateRect (1, 0, 1, 1);
							rectText.right = rectText.left + szDT.cx;

							pDC->DrawText (strStart, rectText, 
								DT_NOPREFIX | DT_RIGHT | DT_VCENTER | DT_SINGLELINE);
							rect.left = rectText.right + szSpace.cx;
						}
					}

					if (pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_FINISH)
					{
						if (dtFinish.GetHour () != 0 || dtFinish.GetMinute () != 0 ||
							dtFinish.GetSecond () != 0 || !bDrawTimeSmart)
						{
							CRect rectText (rect);
							rectText.DeflateRect (1, 0, 1, 1);
							rectText.left = rectText.right - szDT.cx;

							pDC->DrawText (strFinish, rectText, 
								DT_NOPREFIX | DT_RIGHT | DT_VCENTER | DT_SINGLELINE);
							rect.right = rectText.left - szSpace.cx;
						}
					}
				}
				else if (!bAlternative && bDrawTime)
				{
					CRect rectText (rect);
					rectText.DeflateRect (0, 0, 1, 1);
					rectText.right = rectText.left + szDT.cx;

					pDC->DrawText (strStart, rectText, 
						DT_NOPREFIX | DT_RIGHT | DT_VCENTER | DT_SINGLELINE);
					rect.left = rectText.right + szSpace.cx;

					if (IsDrawTimeFinish ())
					{
						if (rect.Width () >= szDT.cx * 3)
						{
							rectText = rect;
							rectText.DeflateRect (0, 0, 1, 1);
							rectText.right = rectText.left + szDT.cx;

							if (!bEmpty)
							{
								pDC->DrawText (strFinish, rectText, 
									DT_NOPREFIX | DT_RIGHT | DT_VCENTER | DT_SINGLELINE);
							}

							rect.left = rectText.right + szSpace.cx;
						}
						else
						{
							bToolTipNeeded = TRUE;
						}
					}
					else
					{
						bToolTipNeeded = TRUE;
					}
				}
			}
			else
			{
				bToolTipNeeded = TRUE;
			}

			if (clrTime != CLR_DEFAULT)
			{
				pDC->SetTextColor (clrText);
			}
		}
	}

	CRect rectEdit (rect);
	rectEdit.right = pDS->GetRect ().right;
/*
	if (!bAlternative)
	{
		if (!bDrawTime && (bFrameRgn || bIsOverrideSelection))
		{
			rectEdit.left -= bFrameRgn ? 2 : 1;
		}

		if (bIsOverrideSelection)
		{
			rectEdit.right -= nSpacingRight + 1;
		}
	}
	else
*/
	{
		if (!bFrameRgn)
		{
			rectEdit.right -= nSpacingRight;
		}
	}

	rectEdit.IntersectRect(rectEdit, m_rectApps);

	pDS->SetRectEditHitTest (rectEdit);

//	if (bAlternative)
	{
		if (bFrameRgn)
		{
			rectEdit.left = pDS->GetRect ().left;

			if (pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_START)
			{
				rectEdit.left += nSpacingLeft + 1;
			}

			if (pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_FINISH)
			{
				rectEdit.right -= nSpacingRight + 1;
			}
		}
		else
		{
			rectEdit.left = pDS->GetRect ().left + nSpacingLeft;
			rectEdit.InflateRect (0, 1);

			if ((pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_START) == 0)
			{
				rectEdit.left--;
			}

			if ((pDS->GetBorder () & CBCGPAppointmentDrawStruct::BCGP_APPOINTMENT_DS_BORDER_FINISH) == 0)
			{
				rectEdit.right++;
			}
		}

		rectEdit.left += rectDuration.Width();
	}

	rectEdit.DeflateRect (1, 1);
	pDS->SetRectEdit (rectEdit);

	rectEdit = pDS->GetRectEditHitTest ();
	rectEdit.DeflateRect (1, 1);
	pDS->SetRectEditHitTest (rectEdit);

	CDWordArray dwIcons;

	CRect rtIcons (rect);
	rtIcons.left++;

	if (szImage != CSize (0, 0))
	{
		if (pApp->m_RecurrenceClone &&
			(dwDrawFlags & BCGP_PLANNER_DRAW_APP_NO_RECURRENCE_IMAGE) == 0)
		{
			if (rtIcons.Width () > szImage.cx * 1.5)
			{
				dwIcons.Add (pApp->m_RecurrenceEcp ? 1 : 0);
				rtIcons.left += szImage.cx;
			}
		}

		if (!pApp->m_lstImages.IsEmpty () &&
			(dwDrawFlags & BCGP_PLANNER_DRAW_APP_NO_IMAGES) == 0)
		{
			for (POSITION pos = pApp->m_lstImages.GetHeadPosition (); pos != NULL;)
			{
				if (rtIcons.Width () <= szImage.cx)
				{
					bCancelDraw = TRUE;
					break;
				}

				rtIcons.left += szImage.cx;

				dwIcons.Add (pApp->m_lstImages.GetNext (pos));
			}
		}
	}

	BOOL bDrawText = !(pApp->m_strDescription.IsEmpty () || bCancelDraw);

	int nTextWidth = pDC->GetTextExtent (pApp->m_strDescription).cx;

	if (dwIcons.GetSize () > 0)
	{
//		long nIconsWidth = rtIcons.left - rect.left;
		
		CPoint pt (rect.left + 1, rect.top + (rect.Height () - szImage.cy) / 2);
/*
		if (bAlternative)
		{
			nIconsWidth += nTextWidth;

			if (nIconsWidth < rect.Width ())
			{
				pt.x += (rect.Width () - nIconsWidth) / 2;
			}
		}
*/
		for (int i = 0; i < (int)dwIcons.GetSize (); i++)
		{
			CBCGPPlannerManagerCtrl::DrawImageIcon (pDC, pt, dwIcons[i]);

			pt.x += szImage.cx;
		}

		rect.left = pt.x + 1;
	}

	if (bDrawText)
	{
		CRect rectText (rect);

		rectText.DeflateRect (4, 0, 1, 1);
/*
		if (bAlternative && dwIcons.GetSize () == 0 && nTextWidth < rectText.Width ())
		{
			rectText.left += (rectText.Width () - nTextWidth) / 2;
		}
*/
		pDC->DrawText (pApp->m_strDescription, rectText, 
			DT_NOPREFIX | DT_VCENTER | DT_SINGLELINE);

		if (rectText.Width () < nTextWidth)
		{
			bToolTipNeeded = TRUE;
		}
	}
	else if (!pApp->m_strDescription.IsEmpty ())
	{
		bToolTipNeeded = TRUE;
	}

	pDC->SetTextColor (clrTextOld);

	if (bFrameRgn)
	{
		CRgn rgnClip;
		rgnClip.CreateRectRgn (rectClip.left, rectClip.top, rectClip.right, rectClip.bottom);

		pDC->SelectClipRgn (&rgnClip, RGN_COPY);
	}

	pDS->SetToolTipNeeded (bToolTipNeeded);
}

BYTE CBCGPPlannerViewSchedule::OnDrawAppointments (CDC* pDC, const CRect& rect, const COleDateTime& date)
{
	return CBCGPPlannerView::OnDrawAppointments (pDC, rect, date);
}

DROPEFFECT CBCGPPlannerViewSchedule::HitTestDrag( DWORD dwKeyState, const CPoint& point ) const
{
	BCGP_PLANNER_HITTEST hit = HitTestArea (point);

	if (hit == BCGP_PLANNER_HITTEST_NOWHERE || hit == BCGP_PLANNER_HITTEST_HEADER)
	{
		return DROPEFFECT_NONE;
	}

	return (dwKeyState & MK_CONTROL) ? DROPEFFECT_COPY : DROPEFFECT_MOVE;	
}

DROPEFFECT CBCGPPlannerViewSchedule::OnDragOver(COleDataObject* pDataObject, DWORD dwKeyState, CPoint point)
{
	ASSERT_VALID (m_pPlanner);

	ASSERT (pDataObject != NULL);
	ASSERT (pDataObject->IsDataAvailable (CBCGPPlannerManagerCtrl::GetClipboardFormat ()));

	DROPEFFECT dragEffect    = HitTestDrag (dwKeyState, point);
	DROPEFFECT dragEffectOld = GetDragEffect ();

	BOOL bDateChanged       = FALSE;
	BOOL bAreaChanged       = FALSE;
	BOOL bResourceChanged   = FALSE;
	BOOL bDragEffectChanged = dragEffectOld != dragEffect;
	BOOL bDragMatchOld      = IsCaptureMatched ();

//	BCGP_PLANNER_HITTEST htAreaOld = m_htCaptureAreaCurrent;

	m_AdjustAction = BCGP_PLANNER_ADJUST_ACTION_NONE;

	if (dragEffect != DROPEFFECT_NONE)
	{
		COleDateTime dtCur  = GetDateFromPoint (point);
		BCGP_PLANNER_HITTEST htArea = HitTestArea (point);
		UINT htResource = GetResourceFromPoint (point);

		bDateChanged = dtCur != COleDateTime () && dtCur != m_dtCaptureCurrent;
		bAreaChanged     = m_htCaptureAreaCurrent != htArea;
		bResourceChanged = m_htCaptureResourceCurrent != htResource;

		if (bDateChanged)
		{
			m_dtCaptureCurrent = dtCur;
		}

		if (bAreaChanged)
		{
/*
			if ((htArea == BCGP_PLANNER_HITTEST_HEADER ||
				 htArea == BCGP_PLANNER_HITTEST_HEADER_RESOURCE) && 
				m_htCaptureAreaCurrent == BCGP_PLANNER_HITTEST_HEADER_ALLDAY)
			{
				htArea = BCGP_PLANNER_HITTEST_HEADER_ALLDAY;
				bAreaChanged = FALSE;
			}
*/
			m_htCaptureAreaCurrent = htArea;
		}

		if (bResourceChanged)
		{
			m_htCaptureResourceCurrent = htResource;
		}
	}

//	BOOL bAllDay = m_htCaptureAreaCurrent == BCGP_PLANNER_HITTEST_HEADER_ALLDAY;

	BOOL bChanged = bDragEffectChanged || bDateChanged || bAreaChanged || bResourceChanged;

	if (bChanged)
	{
		if (m_pDragAppointments != NULL)
		{
			m_pDragAppointments->RemoveAll ();
			m_arDragAppointments.RemoveAll ();
		}
		else
		{
			m_pDragAppointments = 
				(CBCGPAppointmentStorage*)RUNTIME_CLASS(CBCGPAppointmentStorage)->CreateObject ();
		}

		if (dragEffect == DROPEFFECT_NONE)
		{
/*
			if (GetDragEffect () != dragEffect && 
				htAreaOld == BCGP_PLANNER_HITTEST_HEADER_ALLDAY)
			{
				m_AdjustAction = BCGP_PLANNER_ADJUST_ACTION_LAYOUT;
			}
			else
*/
			{
				m_AdjustAction = BCGP_PLANNER_ADJUST_ACTION_APPOINTMENTS;
			}

			return dragEffect;
		}

		CFile* pFile = pDataObject->GetFileData (CBCGPPlannerManagerCtrl::GetClipboardFormat ());
		if (pFile == NULL)
		{
			return FALSE;
		}

		XBCGPAppointmentArray ar;

		BOOL bRes = CBCGPPlannerManagerCtrl::SerializeFrom (*pFile, ar, 
			m_pPlanner->GetType (), 
			IsDragDrop () ? COleDateTime () : m_dtCaptureCurrent);

		delete pFile;

		if (bRes)
		{
			int i = 0;

			if (m_htCaptureAreaCurrent != BCGP_PLANNER_HITTEST_HEADER_RESOURCE)
			{
				for (i = 0; i < (int)ar.GetSize (); i++)
				{
					CBCGPAppointment* pApp = ar[i];

					if (pApp->IsAllDay () && 
						(m_dtCaptureCurrent.GetHour() != 0 || m_dtCaptureCurrent.GetMinute() != 0 || m_dtCaptureCurrent.GetSecond() != 0))
					{
						pApp->SetAllDay (FALSE);
						pApp->SetInterval (pApp->GetStart(), pApp->GetFinish() + COleDateTimeSpan(1, 0, 0, 0));
					}
				}

				if (IsDragDrop ())
				{
					COleDateTimeSpan spanTo (0, 0, 0, 0);

					BOOL bAdd = m_dtCaptureCurrent > m_dtCaptureStart;

					if (bAdd)
					{
						spanTo = m_dtCaptureCurrent - m_dtCaptureStart;
					}
					else
					{
						spanTo = m_dtCaptureStart - m_dtCaptureCurrent;
					}

					CBCGPPlannerManagerCtrl::MoveAppointments (ar, spanTo, bAdd);
				}
			}

			for (i = 0; i < (int)ar.GetSize (); i++)
			{
				CBCGPAppointment* pApp = ar[i];

				pApp->SetSelected (TRUE);

				pApp->SetResourceID (m_htCaptureResourceCurrent);
				m_pDragAppointments->Add (pApp);
			}

			m_pDragAppointments->Query (m_arDragAppointments, m_DateStart, m_DateEnd);

			DROPEFFECT dragEffectNew = CanDropAppointmets(dwKeyState, point);
			if (dragEffectNew != dragEffect)
			{
				dragEffect = dragEffectNew;

				if (dragEffect == DROPEFFECT_NONE)
				{
					m_pDragAppointments->RemoveAll();
					m_arDragAppointments.RemoveAll();
/*
					if (GetDragEffect () != dragEffect && 
						htAreaOld == BCGP_PLANNER_HITTEST_HEADER_ALLDAY)
					{
						m_AdjustAction = BCGP_PLANNER_ADJUST_ACTION_LAYOUT;
					}
					else
*/
					{
						m_AdjustAction = BCGP_PLANNER_ADJUST_ACTION_APPOINTMENTS;
					}

					return dragEffect;
				}
			}
		}
	}

	if (bChanged || bDragMatchOld != IsCaptureMatched ())
	{
/*
		if (
			(
				bAreaChanged && 
				(htAreaOld == BCGP_PLANNER_HITTEST_HEADER_ALLDAY ||
				 m_htCaptureAreaCurrent == BCGP_PLANNER_HITTEST_HEADER_ALLDAY)
			) ||
			(
				bDragEffectChanged && 
				((m_htCaptureAreaStart == BCGP_PLANNER_HITTEST_HEADER_ALLDAY &&
				  m_htCaptureAreaCurrent == BCGP_PLANNER_HITTEST_HEADER_ALLDAY) ||
				 (dragEffectOld == DROPEFFECT_NONE && 
				  m_htCaptureAreaCurrent == BCGP_PLANNER_HITTEST_HEADER_ALLDAY))
			) ||
			(
				(bDateChanged || bResourceChanged) && 
				m_htCaptureAreaCurrent == BCGP_PLANNER_HITTEST_HEADER_ALLDAY
			)
		   )
		{
			TRACE0("BCGP_PLANNER_ADJUST_ACTION_LAYOUT\n");
			m_AdjustAction = BCGP_PLANNER_ADJUST_ACTION_LAYOUT;
		}
		else
*/
		{
			TRACE0("BCGP_PLANNER_ADJUST_ACTION_APPOINTMENTS\n");
			m_AdjustAction = BCGP_PLANNER_ADJUST_ACTION_APPOINTMENTS;
		}
	}

	return dragEffect;
}

BOOL CBCGPPlannerViewSchedule::CanCaptureAppointment (BCGP_PLANNER_HITTEST hitCapturedAppointment) const
{
	if (IsReadOnly ())
	{
		return FALSE;
	}

	return hitCapturedAppointment == BCGP_PLANNER_HITTEST_APPOINTMENT_LEFT ||
		   hitCapturedAppointment == BCGP_PLANNER_HITTEST_APPOINTMENT_RIGHT;
}

CBCGPPlannerView::BCGP_PLANNER_HITTEST
CBCGPPlannerViewSchedule::HitTest (const CPoint& point) const
{
	BCGP_PLANNER_HITTEST hit = HitTestArea (point);

	if (hit == BCGP_PLANNER_HITTEST_CLIENT)
	{
		BCGP_PLANNER_HITTEST hitA = HitTestAppointment (point);

		if (hitA != BCGP_PLANNER_HITTEST_NOWHERE)
		{
			hit = hitA;
		}
	}

	return hit;
}

CBCGPPlannerView::BCGP_PLANNER_HITTEST
CBCGPPlannerViewSchedule::HitTestArea (const CPoint& point) const
{
	BCGP_PLANNER_HITTEST hit = BCGP_PLANNER_HITTEST_NOWHERE;

	BOOL bInTimeBar = FALSE;

	CRect rectApps(m_rectApps);
	if (rectApps.right > m_ViewRects[m_ViewRects.GetSize() - 1].right)
	{
		rectApps.right = m_ViewRects[m_ViewRects.GetSize() - 1].right;
	}

	if (!m_rectResourcesBar.IsRectEmpty() && m_rectResourcesBar.PtInRect (point) && (m_nHeaderHeight + m_rectTimeBar.Height()) <= point.y)
	{
		hit = BCGP_PLANNER_HITTEST_HEADER_RESOURCE;
	}
	else if (!m_rectTimeBar.IsRectEmpty() && m_rectTimeBar.PtInRect (point))
	{
		hit = BCGP_PLANNER_HITTEST_TIMEBAR;
		bInTimeBar = TRUE;
	}
	else if (rectApps.left <= point.x && point.x <= rectApps.right)
	{
		if (0 <= point.y && point.y < m_nHeaderHeight)
		{
			hit = BCGP_PLANNER_HITTEST_HEADER;
		}
		else if (rectApps.PtInRect (point))
		{
			hit = BCGP_PLANNER_HITTEST_CLIENT;
		}
	}

	if ((GetDrawFlags () & BCGP_PLANNER_DRAW_VIEW_NO_UPDOWN) == 0 &&
		(bInTimeBar || hit == BCGP_PLANNER_HITTEST_CLIENT))
	{
		if (hit == BCGP_PLANNER_HITTEST_TIMEBAR || hit == BCGP_PLANNER_HITTEST_CLIENT)
		{
			for (int i = 0; i < (int)m_UpRects.GetSize (); i++)
			{
				if (m_UpRects[i].PtInRect (point))
				{
					hit = BCGP_PLANNER_HITTEST_ICON_UP;
					break;
				}
			}

			if (hit != BCGP_PLANNER_HITTEST_ICON_UP)
			{
				for (int i = 0; i < (int)m_DownRects.GetSize (); i++)
				{
					if (m_DownRects[i].PtInRect (point))
					{
						hit = BCGP_PLANNER_HITTEST_ICON_DOWN;
						break;
					}
				}
			}
		}
		else if (bInTimeBar || hit == BCGP_PLANNER_HITTEST_HEADER_ALLDAY)
		{
			for (int i = 0; i < (int)m_HeaderUpRects.GetSize (); i++)
			{
				if (m_HeaderUpRects[i].PtInRect (point))
				{
					hit = BCGP_PLANNER_HITTEST_ICON_UP;
					break;
				}
			}

			if (hit != BCGP_PLANNER_HITTEST_ICON_UP)
			{
				for (int i = 0; i < (int)m_HeaderDownRects.GetSize (); i++)
				{
					if (m_HeaderDownRects[i].PtInRect (point))
					{
						hit = BCGP_PLANNER_HITTEST_ICON_DOWN;
						break;
					}
				}
			}
		}
	}

	return hit;
}

CBCGPPlannerView::BCGP_PLANNER_HITTEST
CBCGPPlannerViewSchedule::HitTestAppointment (const CPoint& point) const
{
	BCGP_PLANNER_HITTEST hit = BCGP_PLANNER_HITTEST_NOWHERE;

	CBCGPAppointment* pApp = ((CBCGPPlannerView*) this)->GetAppointmentFromPoint (point);

	if (pApp != NULL)
	{
		hit = BCGP_PLANNER_HITTEST_APPOINTMENT;

		switch(pApp->HitTestEx (point, CBCGPAppointment::BCGP_APPOINTMENT_HITTEST_TYPE_HORZ))
		{
		case CBCGPAppointment::BCGP_APPOINTMENT_HITTEST_INSIDE:
			hit = BCGP_PLANNER_HITTEST_APPOINTMENT;
			break;
		case CBCGPAppointment::BCGP_APPOINTMENT_HITTEST_MOVE:
			hit = BCGP_PLANNER_HITTEST_APPOINTMENT_MOVE;
			break;
		case CBCGPAppointment::BCGP_APPOINTMENT_HITTEST_LEFT:
			hit = BCGP_PLANNER_HITTEST_APPOINTMENT_LEFT;
			break;
		case CBCGPAppointment::BCGP_APPOINTMENT_HITTEST_RIGHT:
			hit = BCGP_PLANNER_HITTEST_APPOINTMENT_RIGHT;
			break;
		default:
			ASSERT(FALSE);
		}
	}

	return hit;
}

BOOL CBCGPPlannerViewSchedule::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	if (pScrollBar == GetPlanner()->GetHeaderScrollBarCtrl(SB_VERT))
	{
		return FALSE;
	}

	int nPrevOffset = m_nScrollOffset;

	if (pScrollBar->GetSafeHwnd() != NULL && (nSBCode == SB_THUMBPOSITION || nSBCode == SB_THUMBTRACK))
	{
		SCROLLINFO si;
		ZeroMemory(&si, sizeof(SCROLLINFO));
		si.cbSize = sizeof(SCROLLINFO);
		si.fMask  = SIF_TRACKPOS;
		pScrollBar->GetScrollInfo(&si, si.fMask);

		nPos = si.nTrackPos;
	}

	switch (nSBCode)
	{
	case SB_LINEUP:
		m_nScrollOffset -= m_nRowHeight;
		break;

	case SB_LINEDOWN:
		m_nScrollOffset += m_nRowHeight;
		break;

	case SB_TOP:
		m_nScrollOffset = 0;
		break;

	case SB_BOTTOM:
		m_nScrollOffset = m_nScrollTotal;
		break;

	case SB_PAGEUP:
		m_nScrollOffset -= m_nScrollPage;
		break;

	case SB_PAGEDOWN:
		m_nScrollOffset += m_nScrollPage;
		break;

	case SB_THUMBPOSITION:
	case SB_THUMBTRACK:
		if ((int)nPos != (m_nScrollTotal - m_nScrollPage + 1) && pScrollBar->GetSafeHwnd() != NULL)
		{
			m_nScrollOffset = (int)(nPos / m_nRowHeight) * m_nRowHeight;
		}
		else
		{
			m_nScrollOffset = nPos;
		}
		break;

	default:
		return FALSE;
	}

	m_nScrollOffset = max (0, m_nScrollOffset);

	if (m_nScrollTotal == m_nScrollPage)
	{
		if (m_nScrollOffset > 0)
		{
			m_nScrollOffset = 1;
		}
	}
	else
	{
		m_nScrollOffset = min (m_nScrollOffset, m_nScrollTotal - m_nScrollPage + 1);
	}

	if (m_nScrollOffset == nPrevOffset)
	{
		return FALSE;
	}

	return CBCGPPlannerView::OnVScroll (nSBCode, nPos, pScrollBar);
}

BOOL CBCGPPlannerViewSchedule::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	int nPrevOffset = m_nScrollHorzOffset;

	if (pScrollBar->GetSafeHwnd() != NULL && (nSBCode == SB_THUMBPOSITION || nSBCode == SB_THUMBTRACK))
	{
		SCROLLINFO si;
		ZeroMemory(&si, sizeof(SCROLLINFO));
		si.cbSize = sizeof(SCROLLINFO);
		si.fMask  = SIF_TRACKPOS;
		pScrollBar->GetScrollInfo(&si, si.fMask);

		nPos = si.nTrackPos;
	}

	switch (nSBCode)
	{
	case SB_LINEUP:
		m_nScrollHorzOffset -= m_nColWidth;
		break;

	case SB_LINEDOWN:
		m_nScrollHorzOffset += m_nColWidth;
		break;

	case SB_TOP:
		m_nScrollHorzOffset = 0;
		break;

	case SB_BOTTOM:
		m_nScrollHorzOffset = m_nScrollHorzTotal;
		break;

	case SB_PAGEUP:
		m_nScrollHorzOffset -= m_nScrollHorzPage;
		break;

	case SB_PAGEDOWN:
		m_nScrollHorzOffset += m_nScrollHorzPage;
		break;

	case SB_THUMBPOSITION:
	case SB_THUMBTRACK:
		if ((int)nPos != (m_nScrollHorzTotal - m_nScrollHorzPage + 1) && pScrollBar->GetSafeHwnd() != NULL)
		{
			m_nScrollHorzOffset = (int)(nPos / m_nColWidth) * m_nColWidth;
		}
		else
		{
			m_nScrollHorzOffset = nPos;
		}
		break;

	default:
		return FALSE;
	}

	m_nScrollHorzOffset = max (0, m_nScrollHorzOffset);

	if (m_nScrollHorzTotal == m_nScrollHorzPage)
	{
		if (m_nScrollHorzOffset > 0)
		{
			m_nScrollHorzOffset = 1;
		}
	}
	else
	{
		m_nScrollHorzOffset = min (m_nScrollHorzOffset, m_nScrollHorzTotal - m_nScrollHorzPage + 1);
	}

	if (m_nScrollHorzOffset == nPrevOffset)
	{
		return FALSE;
	}

	return CBCGPPlannerView::OnHScroll (nSBCode, nPos, pScrollBar);
}

BOOL CBCGPPlannerViewSchedule::OnKeyDown(UINT nChar, UINT /*nRepCnt*/, UINT /*nFlags*/)
{
	if (m_pCapturedAppointment != NULL)
	{
		GetPlanner ()->SendMessage (WM_CANCELMODE, 0, 0);
		return TRUE;
	}

	BOOL isShiftPressed = (0x8000 & GetKeyState(VK_SHIFT)) != 0;

	const int nMinuts = CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ());

	COleDateTime sel1 = GetSelectionStart ();
	COleDateTime sel2 = GetSelectionEnd ();

	if (sel1.GetSecond () == 59)
	{
		sel1 = COleDateTime (sel1.GetYear (), sel1.GetMonth (), sel1.GetDay (),
							 sel1.GetHour (), sel1.GetMinute () - nMinuts + 1, 0);
	}

	if (sel2.GetSecond () == 59)
	{
		sel2 = COleDateTime (sel2.GetYear (), sel2.GetMonth (), sel2.GetDay (),
							 sel2.GetHour (), sel2.GetMinute () - nMinuts + 1, 0);
	}

	BOOL bHandled = FALSE;
/*
	if (nChar == VK_PRIOR || nChar == VK_NEXT) // Page Up, Page Down
	{
		int nPrevScroll = m_nScrollOffset;

		OnVScroll (nChar == VK_PRIOR ? SB_PAGEUP : SB_PAGEDOWN, 0, NULL);

		if (nPrevScroll == m_nScrollOffset)
		{
			sel2 = COleDateTime (
				m_Date.GetYear (), m_Date.GetMonth (), m_Date.GetDay (), 
				nChar == VK_PRIOR ? 0 : 23, nChar == VK_PRIOR ? 0 : 60 - nMinuts, 0);
		}
		else
		{
			int n = m_nScrollPage * nMinuts;

			COleDateTimeSpan span (0, n / 60, n % 60, 0);

			if (nChar == VK_PRIOR)
			{
				sel2 -= span;
			}
			else
			{
				sel2 += span;
			}
		}

		if (!isShiftPressed)
		{
			sel1 = sel2;
		}

		SetSelection (sel1, sel2);

		bHandled = TRUE;
	}
	else */if (nChar == VK_HOME || nChar == VK_END)
	{
		if (nChar == VK_HOME)
		{
			sel2 = COleDateTime (
				m_Date.GetYear (), m_Date.GetMonth (), m_Date.GetDay (), 
				GetFirstSelectionHour (), (int)(GetFirstSelectionMinute () / nMinuts) * nMinuts,
				0);
		}
		else
		{
			int nLastHour  = GetLastSelectionHour ();
			int nLastMinut = (int)(GetLastSelectionMinute () / nMinuts) * nMinuts;
			if (nLastMinut < nMinuts)
			{
				if (nLastHour > 0)
				{
					nLastHour--;
				}
				nLastMinut = 60;
			}

			sel2 = COleDateTime (
				m_Date.GetYear (), m_Date.GetMonth (), m_Date.GetDay (), 
				nLastHour, nLastMinut - nMinuts, 0);
		}

		if (!isShiftPressed)
		{
			sel1 = sel2;
		}

		BOOL bScroll = FALSE;
		int nPos = 0;

		if (m_nScrollHorzTotal != 0)
		{
			const int nViewOffset = (sel2 - m_DateStart).GetDays() * GetViewHours() * 60;
			nPos = ((60 * (sel2.GetHour () - GetFirstViewHour()) + nViewOffset + sel2.GetMinute ()) / nMinuts + (nChar == VK_HOME ? 0 : 1)) * m_nColWidth;
			bScroll = nPos < m_nScrollHorzOffset || (m_nScrollHorzOffset + m_nScrollHorzPage) < nPos;
		}

		SetSelection (sel1, sel2, !bScroll);

		if (bScroll)
		{
			OnHScroll (SB_THUMBPOSITION, nChar == VK_HOME ? nPos : nPos - m_nScrollHorzPage, NULL);
		}

		bHandled = TRUE;
	}
	else if (nChar == VK_LEFT || nChar == VK_RIGHT)
	{
		COleDateTimeSpan span (0, nMinuts / 60, nMinuts == 60 ? 0 : nMinuts, 0);

		COleDateTime dtStart(m_DateStart.GetYear(), m_DateStart.GetMonth(), m_DateStart.GetDay(), GetFirstViewHour (), 0, 0);
		COleDateTime dtEnd(m_DateEnd.GetYear(), m_DateEnd.GetMonth(), m_DateEnd.GetDay(), GetLastViewHour (), 60 - nMinuts, 0);

		if (nChar == VK_LEFT)
		{
			sel2 -= span;
			if (sel2.GetHour() < GetFirstViewHour ())
			{
				sel2.SetDateTime(sel2.GetYear(), sel2.GetMonth(), sel2.GetDay(), GetLastViewHour(), 60 - nMinuts, 0);
				sel2 -= COleDateTimeSpan(1, 0, 0, 0);
			}
		}
		else
		{
			sel2 += span;
			if (GetLastViewHour () < sel2.GetHour())
			{
				sel2.SetDateTime(sel2.GetYear(), sel2.GetMonth(), sel2.GetDay(), GetFirstViewHour(), 0, 0);
				sel2 += COleDateTimeSpan(1, 0, 0, 0);
			}
		}

		if (sel2 < dtStart)
		{
			sel2 = dtStart;
		}

		if (sel2 > dtEnd)
		{
			sel2 = dtEnd;
		}

		if (!isShiftPressed)
		{
			sel1 = sel2;
		}

		if (!IsOneDay (m_Date, sel2))
		{
			m_Date = COleDateTime (sel2.GetYear (), sel2.GetMonth (), sel2.GetDay (), 0, 0, 0);
		}

		BOOL bScroll = FALSE;
		int nPos = 0;

		if (m_nScrollHorzTotal != 0)
		{
			const int nViewOffset = (sel2 - m_DateStart).GetDays() * GetViewHours() * 60;
			nPos = ((60 * (sel2.GetHour () - GetFirstViewHour()) + nViewOffset + sel2.GetMinute ()) / nMinuts + (nChar == VK_LEFT ? 0 : 1)) * m_nColWidth;
			bScroll = nPos < m_nScrollHorzOffset || (m_nScrollHorzOffset + m_nScrollHorzPage) < nPos;
		}

		SetSelection (sel1, sel2, !bScroll);

		if (bScroll)
		{
			OnHScroll (SB_THUMBPOSITION, nChar == VK_LEFT ? nPos : nPos - m_nScrollHorzPage, NULL);
		}

		bHandled = TRUE;
	}
	else if (nChar == VK_UP || nChar == VK_DOWN || nChar == VK_PRIOR || nChar == VK_NEXT)
	{
		if (nChar == VK_PRIOR)
		{
			nChar = VK_UP;
		}
		else if (nChar == VK_NEXT)
		{
			nChar = VK_DOWN;
		}

		if (m_Resources.GetSize() <= 1)
		{
			OnVScroll(nChar == VK_UP
						? SB_LINEUP
						: nChar == VK_DOWN
							? SB_LINEDOWN
							: nChar == VK_PRIOR
								? SB_PAGEUP
								: SB_PAGEDOWN,
						0, NULL);
		}
		else
		{
			int nRes = FindResourceIndexByID(GetCurrentResourceID());
			int nNewRes = bcg_clamp(nRes + (nChar == VK_UP ? -1 : 1), 0, (int)m_Resources.GetSize() - 1);

			if (nRes != nNewRes)
			{
				BOOL bScroll = FALSE;
				int nPos = 0;

				if (m_nScrollTotal != 0)
				{
					nPos = m_Resources[nNewRes].m_Rects[0].top - m_rectApps.top + (nChar == VK_UP ? 0 : m_Resources[nNewRes].m_Rects[0].Height());
					bScroll = nPos < m_nScrollOffset || (m_nScrollOffset + m_nScrollPage) < nPos;
				}

				SetCurrentResourceID(m_Resources[nNewRes].m_ResourceID, !bScroll, TRUE);

				if (bScroll)
				{
					OnVScroll (SB_THUMBPOSITION, nChar == VK_UP ? nPos : nPos - m_nScrollPage, NULL);
				}
			}
		}

		bHandled = TRUE;
	}

	return bHandled;
}

void CBCGPPlannerViewSchedule::GetDragScrollRect (CRect& rect)
{
	CBCGPPlannerView::GetDragScrollRect (rect);

	rect.left = m_rectResourcesBar.left;
}

COleDateTime CBCGPPlannerViewSchedule::GetDateFromPoint (const CPoint& point) const
{
	COleDateTime date;

	if (!m_rectResourcesBar.IsRectEmpty ())
	{
		if (m_rectResourcesBar.PtInRect(point))
		{
			return date;
		}
	}

	CRect rectClient (m_rectApps);

	if (!m_rectResourcesBar.IsRectEmpty ())
	{
		rectClient.left = m_rectResourcesBar.right;
	}

	CPoint pt (point - rectClient.TopLeft ());
	pt.x += m_nScrollHorzOffset;

	date = m_Date;

	const int nCount = GetViewHours() * 60 / CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ());
	const int dxColumn = nCount * m_nColWidth;

	int nDay = 0;

	{
		nDay = pt.x / dxColumn;
		date = GetDateStart () + COleDateTimeSpan (nDay, 0, 0, 0);

		if(date > m_DateEnd)
		{
			date = m_DateEnd;
		}
	}

	if ((!m_rectTimeBar.IsRectEmpty() && point.y < m_rectTimeBar.top) ||
		(m_rectTimeBar.IsRectEmpty() && point.y <= rectClient.top))
	{
		date.SetDate (date.GetYear (), date.GetMonth (), date.GetDay ());
		return date;
	}

	const int nMinuts = (pt.x - dxColumn * nDay) / m_nColWidth * CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta()) + GetFirstViewHour() * 60;

	date += COleDateTimeSpan (0, nMinuts / 60, nMinuts % 60, 0);

	return date;
}

CRect CBCGPPlannerViewSchedule::GetRectFromDate(const COleDateTime& date) const
{
	CRect rect(0, 0, 0, 0);

	const int nDays = GetViewDuration ();
	if (m_ViewRects.GetSize() != nDays)
	{
		return rect;
	}

	if (date.GetHour() < GetFirstViewHour() || GetLastViewHour() < date.GetHour())
	{
		return rect;
	}

	COleDateTime dt(date.GetYear(), date.GetMonth(), date.GetDay(), 0, 0, 0);
	COleDateTime dtCur(GetDateStart());

	int nIndex = (int)((dt - dtCur).GetDays());
	if (nIndex < 0 || nDays <= nIndex)
	{
		return rect;
	}

	rect = m_ViewRects[nIndex];
	rect.left += ((date.GetHour() - GetFirstViewHour()) * 60 + date.GetMinute()) * m_nColWidth / 
		CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ()) - m_nScrollHorzOffset;
	rect.right = rect.left + m_nColWidth;

	return rect;
}

BOOL CBCGPPlannerViewSchedule::OnTimer(UINT_PTR nIDEvent)
{
	CBCGPPlannerManagerCtrl* pPlanner = GetPlanner ();
	ASSERT_VALID (pPlanner);

	if (pPlanner->GetSafeHwnd () == NULL)
	{
		return FALSE;
	}

	if (nIDEvent == BCGP_PLANNER_TIMER_EVENT)
	{
		BOOL bDrawTimeLine = IsCurrentTimeVisible ();

		if (!bDrawTimeLine)
		{
			StopTimer (FALSE);
		}

		m_CurrentTime = COleDateTime::GetCurrentTime ();

		CDC* pDC = pPlanner->GetDC ();

		if (pDC != NULL && pDC->GetSafeHdc () != NULL)
		{
/*
			if (GetDrawFlags() & BCGP_PLANNER_DRAW_VIEW_TIMELINE_FULL_WIDTH)
			{
				pPlanner->DoPaint(pDC);
			}
			else
*/
			{
				int nOldBack = pDC->SetBkMode (TRANSPARENT);

				CRgn rgn;
				rgn.CreateRectRgn(m_rectTimeBar.left, m_rectTimeBar.top, m_rectTimeBar.right, m_rectTimeBar.top + m_rectTimeBar.Height() / 2);
				pDC->SelectClipRgn(&rgn);

				OnDrawTimeBar (pDC, m_rectTimeBar, bDrawTimeLine);

				pDC->SelectClipRgn(NULL);

				pDC->SetBkMode (nOldBack);
			}
		}

		pPlanner->ReleaseDC (pDC);

		return TRUE;
	}

	return CBCGPPlannerView::OnTimer(nIDEvent);
}

int CBCGPPlannerViewSchedule::GetViewHourOffset () const
{
	return m_nScrollHorzOffset + GetFirstViewHour () * 60 / CBCGPPlannerView::GetTimeDeltaInMinuts (GetTimeDelta ()) * m_nColWidth;
}

BOOL CBCGPPlannerViewSchedule::OnLButtonDown(UINT nFlags, CPoint point)
{
	BCGP_PLANNER_HITTEST hit = HitTest (point);

	if (hit == BCGP_PLANNER_HITTEST_HEADER_RESOURCE)
	{
		BOOL bRepaint = FALSE;

		if (GetSelectedAppointments ().GetCount () > 0)
		{
			bRepaint = TRUE;
			ClearAppointmentSelection (FALSE);
		}

		const UINT nResourceID = GetResourceFromPoint (point);

		if (nResourceID != GetPlanner()->GetCurrentResourceID ())
		{
			bRepaint = bRepaint || SetCurrentResourceID (nResourceID, FALSE, TRUE);
		}

		if (bRepaint)
		{
			GetPlanner()->RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
		}

		return TRUE;
	}

	return CBCGPPlannerViewDay::OnLButtonDown (nFlags, point);
}

BOOL CBCGPPlannerViewSchedule::OnMouseMove(UINT nFlags, CPoint point)
{
	if (IsCaptured ())
	{
		BCGP_PLANNER_HITTEST hit = HitTestArea (point);

		if (hit == BCGP_PLANNER_HITTEST_HEADER_RESOURCE)
		{
			return TRUE;
		}

		if (IsDragDrop () ||
			(hit != BCGP_PLANNER_HITTEST_HEADER &&
			 hit != BCGP_PLANNER_HITTEST_HEADER_RESOURCE &&
			 hit != BCGP_PLANNER_HITTEST_HEADER_ALLDAY) ||
			m_pCapturedAppointment != NULL)
		{
			CBCGPPlannerManagerCtrl* pPlanner = GetPlanner ();

			if (hit != BCGP_PLANNER_HITTEST_NOWHERE && hit != BCGP_PLANNER_HITTEST_HEADER)
			{
				COleDateTime date = GetDateFromPoint (point);

				m_htCaptureAreaCurrent = hit;

				if (date != COleDateTime () && date != m_dtCaptureCurrent)
				{
					m_dtCaptureCurrent = date;

					if (!IsDragDrop () && m_hitCapturedAppointment == BCGP_PLANNER_HITTEST_NOWHERE)
					{
						SetSelection (m_Selection [0], m_dtCaptureCurrent);
					}
					else if (m_pCapturedAppointment != NULL)
					{
						BCGP_PLANNER_HITTEST hitCapturedAppointment = m_hitCapturedAppointment;					

						COleDateTime dtS (m_pCapturedAppointment->GetStart ());
						COleDateTime dtF (m_pCapturedAppointment->GetFinish ());

						if (m_pCapturedAppointment->IsAllDay())
						{
							dtF += COleDateTimeSpan(1, 0, 0, 0);
						}

						if (hitCapturedAppointment == BCGP_PLANNER_HITTEST_APPOINTMENT_LEFT)
						{
							if (dtS != date)
							{
								dtS = COleDateTime (date.GetYear (), date.GetMonth (), date.GetDay (),
									date.GetHour (), date.GetMinute (), 0);

								if (dtF < dtS)
								{
									dtS = dtF;
								}
							}
						}
						else if (hitCapturedAppointment == BCGP_PLANNER_HITTEST_APPOINTMENT_RIGHT)
						{
							COleDateTime dt (date);
							dt += GetMinimumSpan ();

							if (dtF != dt)
							{
								dtF = COleDateTime (dt.GetYear (), dt.GetMonth (), dt.GetDay (),
									dt.GetHour (), dt.GetMinute (), 0);

								if (dtF < dtS)
								{
									dtF = dtS;
								}
							}
						}

						if (m_pCapturedAppointment->GetStart () != dtS ||
							m_pCapturedAppointment->GetFinish () != dtF)
						{
							m_bCapturedAppointmentChanged = TRUE;

							BOOL bWasMulti = m_pCapturedAppointment->IsMultiDay ();

							m_pCapturedAppointment->SetAllDay(FALSE);
							m_pCapturedAppointment->SetInterval (dtS, dtF);

							{
								// need resort appointments
								XBCGPAppointmentArray& ar = GetQueryedAppointments ();

								SortAppointments (ar, (int) ar.GetSize (), GetSortAppointmentsType());
							}

							if (bWasMulti != m_pCapturedAppointment->IsMultiDay () || bWasMulti)
							{
								pPlanner->AdjustLayout (FALSE);
								date = GetDateFromPoint (point);

								if (date != COleDateTime())
								{
									m_dtCaptureCurrent = date;
								}
							}
							else
							{
								AdjustAppointments ();
							}

							pPlanner->RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
						}
					}
				}
			}
		}
	}

	return CBCGPPlannerViewDay::OnMouseMove (nFlags, point);
}

BOOL CBCGPPlannerViewSchedule::EnsureVisible(const CBCGPAppointment* pApp, BOOL bPartialOK)
{
	ASSERT_VALID (pApp);
	if (pApp == NULL)
	{
		return FALSE;
	}

	CRect rect;
	BOOL bVisible = FALSE;

	const CBCGPAppointmentDSExMap& ds = pApp->GetDSDraw ();
	if (ds.IsEmpty ())
	{
		rect     = pApp->GetRectDraw ();
		bVisible = pApp->IsVisibleDraw ();
	}
	else
	{
		for (int i = 0; i < ds.GetCount (); i++)
		{
			const CBCGPAppointmentDrawStructEx* pDS = 
				(CBCGPAppointmentDrawStructEx*)ds.GetByIndex (i);

			if (pDS != NULL)
			{
				rect     = pDS->GetRect ();
				bVisible = pDS->IsVisible ();
				break;
			}
		}

		if (bPartialOK && !bVisible)
		{
			BOOL bPartialVisible = FALSE;
			for (int i = 0; i < ds.GetCount (); i++)
			{
				const CBCGPAppointmentDrawStructEx* pDS = 
					(CBCGPAppointmentDrawStructEx*)ds.GetByIndex (i);

				if (pDS != NULL)
				{
					if (pDS->IsVisible ())
					{
						bPartialVisible = TRUE;
						break;
					}
				}
			}

			if (bPartialVisible)
			{
				return TRUE;
			}
		}
	}

	if (rect.IsRectEmpty ())
	{
		return FALSE;
	}

	rect.OffsetRect (0, 1);

	if (bVisible)
	{
		if (bPartialOK || (rect.left >= m_rectApps.left && rect.right <= m_rectApps.right && 
			rect.top >= m_rectApps.top && rect.bottom <= m_rectApps.bottom))
		{
			return TRUE;
		}
	}

	if ((rect.left < m_rectApps.left && rect.right < m_rectApps.left) ||
		(m_rectApps.right < rect.left && m_rectApps.right < rect.right))
	{
		rect.OffsetRect(-m_rectApps.left + m_nScrollHorzOffset, 0);

		BOOL bLeft = TRUE;
		int nPos = rect.left;

		if (rect.left >= m_nScrollHorzOffset && rect.Width () <= m_rectApps.Width())
		{
			nPos = rect.right;
			bLeft = FALSE;
		}

		if (nPos < m_nScrollHorzOffset || (m_nScrollHorzOffset + m_nScrollHorzPage) < nPos)
		{
			OnHScroll (SB_THUMBPOSITION, bLeft ? nPos : nPos - m_nScrollHorzPage, NULL);
		}
	}
/*
	if (m_rectApps.top < rect.top || m_rectApps.bottom < rect.bottom)
	{
		int nPos = 0;
		if (rect.top < m_rectApps.top)
		{
			//scroll up
			nPos = (rect.top - m_rectApps.top);
		}
		else
		{
			//scroll down
			if (rect.Height () > m_rectApps.Height())
			{
				nPos = (rect.top - m_rectApps.top);
			}
			else
			{
				nPos = (rect.bottom - m_rectApps.top) - m_nScrollPage;
			}
		}

		OnVScroll (SB_THUMBPOSITION, nPos, NULL);
	}
*/
	return TRUE;
}

UINT CBCGPPlannerViewSchedule::GetCurrentResourceID () const
{
	if (m_Resources.GetSize () == 0)
	{
		return CBCGPPlannerViewDay::GetCurrentResourceID();
	}

	return m_nResourceID;
}

BOOL CBCGPPlannerViewSchedule::SetCurrentResourceID (UINT nResourceID, 
												  BOOL bRedraw/* = TRUE*/, BOOL bNotify/* = FALSE*/)
{
	if (m_bMultiResources && m_Resources.GetSize () == 0 && GetPlanner () != NULL)
	{
		CBCGPAppointmentBaseMultiStorage* pStorage = DYNAMIC_DOWNCAST(CBCGPAppointmentBaseMultiStorage, GetPlanner ()->GetStorage ());

		if (pStorage != NULL && m_bMultiResources)
		{
			CBCGPAppointmentBaseMultiStorage::XResourceIDArray ar;
			pStorage->GetResourceIDs (ar);

			const int count = (int) ar.GetSize ();
			if (count > 0 && count != (int)m_Resources.GetSize())
			{
				m_Resources.SetSize(count);
			}
		}
	}

	if (m_Resources.GetSize () == 0)
	{
		return CBCGPPlannerViewDay::SetCurrentResourceID(nResourceID, bRedraw, bNotify);
	}

	if (m_nResourceID != nResourceID)
	{
		m_nResourceID = nResourceID;

		if (bRedraw && GetPlanner()->GetSafeHwnd () != NULL)
		{
			GetPlanner()->RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
		}

		if (bNotify)
		{
			CBCGPPlannerManagerCtrl* pPlanner = GetPlanner ();

			if (pPlanner->IsNotifyParent () && pPlanner->GetSafeHwnd () != NULL)
			{
				pPlanner->GetParent ()->SendMessage (BCGP_PLANNER_RESOURCEID_CHANGED, 0, 0);
			}
		}
	}

	return TRUE;
}

UINT CBCGPPlannerViewSchedule::GetResourceFromPoint (const CPoint& point) const
{
	const int count = (int)m_Resources.GetSize ();
	if (count == 0)
	{
		return CBCGPPlannerViewDay::GetResourceFromPoint(point);
	}

	CPoint pt(point);
	pt.y += m_nScrollOffset;

	UINT nResourceID = GetPlanner()->GetCurrentResourceID ();

	XResourceCollection& resources = const_cast<XResourceCollection&>(m_Resources);

	for (int i = 0; i < count; i++)
	{
		const CRect& rect = resources[i].m_Rects[0];
		if (rect.top <= pt.y && pt.y < rect.bottom)
		{
			nResourceID = resources[i].m_ResourceID;
			break;
		}
	}

	return nResourceID;
}

int CBCGPPlannerViewSchedule::FindResourceIndexByID (UINT nResourceID) const
{
	int nIndex = -1;

	XResourceCollection& resources = const_cast<XResourceCollection&>(m_Resources);

	for (int i = 0; i < (int)resources.GetSize (); i++)
	{
		if (nResourceID == resources[i].m_ResourceID)
		{
			nIndex = i;
			break;
		}
	}

	return nIndex;
}

BOOL CBCGPPlannerViewSchedule::OnUpdateStorage ()
{
	if (GetPlanner () == NULL)
	{
		return FALSE;
	}

	BOOL bRes = FALSE;

	CBCGPAppointmentBaseMultiStorage* pStorage = DYNAMIC_DOWNCAST(CBCGPAppointmentBaseMultiStorage, GetPlanner ()->GetStorage ());

	if (pStorage != NULL && m_bMultiResources)
	{
		CBCGPAppointmentBaseMultiStorage::XResourceIDArray ar;
		pStorage->GetResourceIDs (ar);

		const int count = (int) ar.GetSize ();
		ASSERT (count > 0);

		if (count != m_Resources.GetSize ())
		{
			m_Resources.SetSize (count);
			bRes= TRUE;
		}

		for (int i = 0; i < count; i++)
		{
			XResource& resource = m_Resources[i];

			if (resource.m_ResourceID != ar[i])
			{
				resource.m_ResourceID = ar[i];
				bRes = TRUE;
			}

			const CBCGPAppointmentBaseResourceInfo* pInfo = pStorage->GetResourceInfo (ar[i]);

			if (pInfo != NULL)
			{
				if (resource.m_Description != pInfo->GetDescription ())
				{
					resource.m_Description = pInfo->GetDescription ();
					bRes = TRUE;
				}
				if (resource.m_WorkStart != pInfo->GetWorkStart ())
				{
					resource.m_WorkStart = pInfo->GetWorkStart ();
					bRes = TRUE;
				}
				if (resource.m_WorkEnd != pInfo->GetWorkEnd ())
				{
					resource.m_WorkEnd = pInfo->GetWorkEnd ();
					bRes = TRUE;
				}
			}
		}
	}
	else if (m_Resources.GetSize() > 0)
	{
		m_Resources.RemoveAll();
		bRes = TRUE;
	}

	return bRes;
}

CString CBCGPPlannerViewSchedule::GetAccName() const
{
	CString strValue(_T("Schedule View"));

	const CBCGPAppointmentBaseMultiStorage* pStorage = DYNAMIC_DOWNCAST(CBCGPAppointmentBaseMultiStorage, GetPlanner ()->GetStorage ());
	if (pStorage != NULL)
	{
		const CBCGPAppointmentBaseResourceInfo* pInfo = pStorage->GetResourceInfo (GetPlanner()->GetCurrentResourceID ());
		if (pInfo != NULL && !pInfo->GetDescription ().IsEmpty ())
		{
			strValue = pInfo->GetDescription () + _T(". ") + strValue;
		}
	}

	return strValue;
}

#endif // BCGP_EXCLUDE_PLANNER
