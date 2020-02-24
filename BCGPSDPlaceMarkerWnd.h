#if !defined(AFX_BCGPSDPLACEMARKERWND_H__FA84F558_73E5_40AA_9C70_4D69E9FF496C__INCLUDED_)
#define AFX_BCGPSDPLACEMARKERWND_H__FA84F558_73E5_40AA_9C70_4D69E9FF496C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

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
// BCGPSDPlaceMarkerWnd.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CBCGPSDPlaceMarkerWnd window

class CBCGPSDPlaceMarkerWnd : public CWnd
{
public:
	enum BCGPSDPlaceMarkerWnd_STYLE 
	{
		BCGPSDPlaceMarkerWnd_Rectangle,
		BCGPSDPlaceMarkerWnd_Frame,
		BCGPSDPlaceMarkerWnd_RectangleWithTab,
		BCGPSDPlaceMarkerWnd_FrameWithTab,
	};
	
	CBCGPSDPlaceMarkerWnd(BCGPSDPlaceMarkerWnd_STYLE style, COLORREF color = CLR_DEFAULT, BYTE nAlpha = 255);
	
	BOOL Create(const CRect& rect, CWnd* pWndOwner, BOOL bAutoDestroy = TRUE);
	
	void SetDockingWnd (CWnd* pDockingWnd)
	{
		m_pDockingWnd = pDockingWnd;
	}
	
    void ShowAt(CRect rect, BCGPSDPlaceMarkerWnd_STYLE style);
    void Hide ();

	BCGPSDPlaceMarkerWnd_STYLE GetStyle() const
	{
		return m_Style;
	}

protected:
	//{{AFX_MSG(CBCGPSDPlaceMarkerWnd)
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnPaint();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
		
	virtual void PostNcDestroy();
	
	void UpdateRgn();
	
	CWnd*						m_pDockingWnd;
	BCGPSDPlaceMarkerWnd_STYLE	m_Style;
	const BYTE					m_nAlpha;
	CBrush						m_Brush;
	CRect						m_rectLast;
	BOOL						m_bAutoDestroy;
	static CString				m_strClassName;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BCGPSDPLACEMARKERWND_H__FA84F558_73E5_40AA_9C70_4D69E9FF496C__INCLUDED_)
