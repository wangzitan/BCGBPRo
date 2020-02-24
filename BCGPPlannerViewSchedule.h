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
// BCGPPlannerViewSchedule.h: interface for the CBCGPPlannerViewSchedule class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BCGPPLANNERVIEWSCHEDULE_H__99CF9A8D_73D3_4D83_836D_0E35487383E9__INCLUDED_)
#define AFX_BCGPPLANNERVIEWSCHEDULE_H__99CF9A8D_73D3_4D83_836D_0E35487383E9__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "BCGCBPro.h"

#ifndef BCGP_EXCLUDE_PLANNER

#include "BCGPPlannerViewDay.h"

class BCGCBPRODLLEXPORT CBCGPPlannerViewSchedule : public CBCGPPlannerViewDay
{
	friend class CBCGPPlannerManagerCtrl;

	DECLARE_DYNCREATE(CBCGPPlannerViewSchedule)

public:
	virtual ~CBCGPPlannerViewSchedule();

#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

	virtual COleDateTime GetDateFromPoint (const CPoint& point) const;
	virtual CRect GetRectFromDate(const COleDateTime& date) const;

	virtual int GetViewHourOffset () const;

	virtual UINT GetCurrentResourceID () const;
	virtual BOOL SetCurrentResourceID (UINT nResourceID, BOOL bRedraw = TRUE, BOOL bNotify = FALSE);
	virtual BOOL OnUpdateStorage ();

	BOOL IsScheduleTimeDeltaDay() const;

protected:
	CBCGPPlannerViewSchedule();

protected:
	virtual BOOL IsCurrentTimeVisible () const;

	virtual void GetCaptionFormatStrings (CStringArray& sa);
	virtual void AdjustScrollSizes ();
	virtual void AdjustLayout (CDC* pDC, const CRect& rectClient);
	virtual void AdjustRects ();
	virtual void AdjustAppointments ();

	virtual BCGP_PLANNER_HITTEST HitTest (const CPoint& point) const;
	virtual BCGP_PLANNER_HITTEST HitTestArea (const CPoint& point) const;
	virtual BCGP_PLANNER_HITTEST HitTestAppointment (const CPoint& point) const;

	virtual void CheckVisibleAppointments(const COleDateTime& date, const CRect& rect, 
		BOOL bFullVisible);
	virtual void CheckVisibleUpDownIcons(BOOL bFullVisible);

	virtual void OnDrawUpDownIcons (CDC* pDC);

	virtual void OnPaint (CDC* pDC, const CRect& rectClient);
	virtual void OnDrawClient (CDC* pDC, const CRect& rect);
	virtual void OnDrawHeader (CDC* pDC, const CRect& rectHeader);
	virtual void OnDrawResourcesBar (CDC* pDC, const CRect& rectBar);
	virtual void OnDrawTimeLine (CDC* pDC, const CRect& rectTime);
	virtual void OnDrawTimeBar (CDC* pDC, const CRect& rectBar, BOOL bDrawTimeLine);
	virtual void OnDrawAppointmentsDuration (CDC* pDC);

	virtual BYTE OnDrawAppointments (CDC* pDC, const CRect& rect = CRect (0, 0, 0, 0), 
		const COleDateTime& date = COleDateTime ());

	virtual void DrawHeader (CDC* pDC, const CRect& rect, int dxColumn);

	virtual void DrawAppointment (CDC* pDC, CBCGPAppointment* pApp, CBCGPAppointmentDrawStructEx* pDS);

	virtual void AddUpDownRect(BYTE nType, const CRect& rect);

	virtual BOOL OnLButtonDown(UINT nFlags, CPoint point);
	virtual BOOL OnMouseMove(UINT nFlags, CPoint point);
	virtual BOOL OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	virtual BOOL OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	virtual BOOL OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	virtual BOOL OnTimer(UINT_PTR nIDEvent);

	virtual void GetDragScrollRect (CRect& rect);

	virtual DROPEFFECT OnDragOver(COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);

	virtual BOOL EnsureVisible(const CBCGPAppointment* pApp, BOOL bPartialOK);

	virtual BOOL CanUseHeaderScrolling() const
	{
		return FALSE;
	}

	virtual BOOL CanUseScrollBar(int /*nBar*/ = SB_VERT)
	{
		return TRUE;
	}

	virtual DROPEFFECT HitTestDrag (DWORD dwKeyState, const CPoint& point) const;
	virtual BOOL CanCaptureAppointment (BCGP_PLANNER_HITTEST hitCapturedAppointment) const;

	virtual BCGP_SORTAPPOINTMENTS_TYPE GetSortAppointmentsType() const
	{
		return BCGP_SORTAPPOINTMENTS_BY_DATE;
	}

	virtual CString GetAccName() const;

protected:
	struct XResource
	{
		UINT m_ResourceID;
		CString m_Description;
		COleDateTime m_WorkStart;
		COleDateTime m_WorkEnd;
		BOOL m_bToolTipNeeded;
		CArray<CRect, CRect&> m_Rects;

		XResource()
			: m_bToolTipNeeded (FALSE)
		{
		}

		XResource(const XResource& src)
		{
			*this = src;
		}

		XResource& operator = (const XResource& rSrc)
		{
			m_ResourceID     = rSrc.m_ResourceID;
			m_Description    = rSrc.m_Description;
			m_WorkStart      = rSrc.m_WorkStart;
			m_WorkEnd        = rSrc.m_WorkEnd;
			m_bToolTipNeeded = rSrc.m_bToolTipNeeded;
			m_Rects.Copy (rSrc.m_Rects);
			
			return *this;
		}
	};

	typedef CArray<XResource, XResource&> XResourceCollection;

	virtual UINT GetResourceFromPoint (const CPoint& point) const;
	int FindResourceIndexByID (UINT nResourceID) const;

protected:
	BOOL				m_bMultiResources;

	int					m_nColWidth;
	CRect				m_rectResourcesBar;

	UINT				m_nResourceID;
	XResourceCollection	m_Resources;
};

#endif // BCGP_EXCLUDE_PLANNER

#endif // !defined(AFX_BCGPPLANNERVIEWSCHEDULE_H__99CF9A8D_73D3_4D83_836D_0E35487383E9__INCLUDED_)
