#if !defined(AFX_BCGPTABWND_H__2ADBF81D_1F51_45D8_9FDB_FF3B13F30CAE__INCLUDED_)
#define AFX_BCGPTABWND_H__2ADBF81D_1F51_45D8_9FDB_FF3B13F30CAE__INCLUDED_

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
// BCGPTabWnd.h : header file
//

#ifndef BCGP_NO_TABCTRL

#ifndef __AFXTEMPL_H__
	#include "afxtempl.h"
#endif

#include "BCGCBPro.h"
#include "BCGPBaseControlBar.h"
#include "BCGPBaseTabWnd.h"
#include "BCGPButton.h"
#include "BCGPScrollBar.h"
#include "BCGPAnimationManager.h"

/////////////////////////////////////////////////////////////////////////////
// CBCGPTabWnd window

class CBCGPTabInfo;
class CBCGPSDPlaceMarkerWnd;

//////////////////////////////////////////////////////////////////////////////
// CBCGPTabWnd notification messages:

extern BCGCBPRODLLEXPORT UINT BCGM_GETDRAGBOUNDS;
extern BCGCBPRODLLEXPORT UINT BCGM_ON_DRAGCOMPLETE;
extern BCGCBPRODLLEXPORT UINT BCGM_ON_TABGROUPMOUSEMOVE;
extern BCGCBPRODLLEXPORT UINT BCGM_ON_MOVETABCOMPLETE;
extern BCGCBPRODLLEXPORT UINT BCGM_ON_CANCELTABMOVE;

extern BCGCBPRODLLEXPORT UINT BCGM_ON_HSCROLL;

class CBCGTabButton : public CBCGPButton
{
	virtual void OnFillBackground (CDC* pDC, const CRect& rectClient);
	virtual void OnDrawBorder (CDC* pDC, CRect& rectClient, UINT uiState);

public:
	void SetPressed (BOOL bOn)
	{
		m_bPushed = bOn;
		m_bHighlighted = bOn;

		RedrawWindow ();
	}
};

class BCGCBPRODLLEXPORT CBCGPTabWnd :	public CBCGPBaseTabWnd,
										public CBCGPAnimationManager,
										public CBCGPCtrlPlaceHolderHelper
{
	friend class CBCGPMainClientAreaWnd;
	friend class CBCGPTabbedControlBar;
	friend class CBCGTabButton;
	friend class CBCGPMDIChildDragWnd;

	DECLARE_DYNCREATE(CBCGPTabWnd)

// Construction
public:
	CBCGPTabWnd();

	enum Style
	{
		STYLE_3D						= 0,
		STYLE_FLAT						= 1,
		STYLE_FLAT_SHARED_HORZ_SCROLL	= 2,
		STYLE_3D_SCROLLED				= 3,
		STYLE_3D_ONENOTE				= 4,
		STYLE_3D_VS2005					= 5,
		STYLE_3D_ROUNDED				= 6,
		STYLE_3D_ROUNDED_SCROLL			= 7,
		STYLE_POINTER					= 8,
		STYLE_DOTS						= 9,
		STYLE_UNDERLINE					= 10,
	};

	enum ResizeMode
	{
		RESIZE_NO						= 0,
		RESIZE_VERT						= 1,
		RESIZE_HORIZ					= 2,
	};

	enum TabCloseButtonMode
	{
		TAB_CLOSE_BUTTON_NONE				= 0,
		TAB_CLOSE_BUTTON_ACTIVE				= 1,
		TAB_CLOSE_BUTTON_HIGHLIGHTED_COMPACT= 2,
		TAB_CLOSE_BUTTON_HIGHLIGHTED		= 3,
		TAB_CLOSE_BUTTON_ALL				= 4,
	};

	enum TabDotShape
	{
		TAB_DOT_ELLIPSE	= 0,
		TAB_DOT_SQUARE	= 1,
	};

	enum TabsDocumentsMenuStyle
	{
		TAB_DOCUMENTS_MENU_DOWN_ARROW	= 0,
		TAB_DOCUMENTS_MENU_ELLIPSIS		= 1,
		TAB_DOCUMENTS_MENU_HAMBURGER	= 2,
	};

// Operations
public:
	// Create methods:
	BOOL Create (Style style, const RECT& rect, CWnd* pParentWnd, UINT nID,
				Location location = LOCATION_BOTTOM,
				BOOL bCloseBtn = FALSE);

	virtual BOOL SetImageList (UINT uiID, int cx = 15, COLORREF clrTransp = RGB (255, 0, 255));
	virtual BOOL SetImageList (HIMAGELIST hImageList);


	BOOL ModifyTabStyle (Style style, BOOL bUpdateScrollControls = FALSE);

	virtual void RecalcLayout ();

	// Tab animation ("busy" indicator) support:
	BOOL SetTabStateAnimationImages(UINT uiAnimImagesResID, int cx = 16, BOOL bDPIAutoScale = TRUE, COLORREF clrTransp = RGB (255, 0, 255));
	BOOL SetTabStateAnimationImages(const CBCGPToolBarImages& animImages);
	
	void SetTabStateAnimationFrameDelay(UINT nDelay /* ms */)
	{
		m_nTabStateAnimationDelay = nDelay;
	}

	UINT GetTabStateAnimationFrameDelay() const
	{
		return m_nTabStateAnimationDelay;
	}

	BOOL StartTabStateAnimation(int iTab, BOOL bRedraw = TRUE);
	virtual void StopTabStateAnimation(int iTab, BOOL bRedraw = TRUE);

	BOOL IsTabStateAnimation(int iTab) const;
	
	void ResetTabStateAnimationStep(int iTab);

	void SetTabCloseButtonMode(TabCloseButtonMode mode);
	TabCloseButtonMode GetTabCloseButtonMode() const
	{
		return m_TabCloseButtonMode;
	}

	void EnableActiveTabCloseButton (BOOL bEnable = TRUE);
	virtual BOOL IsActiveTabCloseButton() const
	{
		return m_TabCloseButtonMode == TAB_CLOSE_BUTTON_ACTIVE && !IsDotsStyle();
	}
	
	virtual BOOL IsTabCloseButton() const 
	{ 
		return m_TabCloseButtonMode != TAB_CLOSE_BUTTON_NONE && !IsDotsStyle();
	}

	void EnableTabDocumentsMenu (BOOL bEnable = TRUE, TabsDocumentsMenuStyle style = TAB_DOCUMENTS_MENU_DOWN_ARROW);
	
	BOOL IsTabDocumentsMenu () const
	{
		return m_bTabDocumentsMenu;
	}
	
	TabsDocumentsMenuStyle GetTabDocumentsMenuStyle() const
	{
		return m_TabDocumentsMenuStyle;
	}

	virtual void EnableInPlaceEdit (BOOL bEnable);
	virtual void CalcRectEdit (CRect& rectEdit);

	virtual void EnableNewTab(BOOL bEnable = TRUE);
	virtual BOOL IsNewTabEnabled() const
	{
		return m_bNewTab;
	}

	void SetNewTabLabel(const CString strLabel);
	LPCTSTR GetNewTabLabel() const
	{
		return m_strNewTabLabel;
	}

	void EnableTabAnimation(BOOL bEnable = TRUE, double dblAnimationTime = 0.3 /* Seconnds */);	// STYLE_POINTER or STYLE_UNDERLINE only!
	BOOL IsTabAnimationEnabled() const
	{
		return m_bTabAnimationSupport;
	}
	double GetTabAnimationTime() const
	{
		return m_dblTabAnimationTime;
	}

	void EnableTabsHandCursor(BOOL bEnable = TRUE)
	{
		m_bTabsHandCursor = bEnable;
	}

	BOOL IsTabsHandCursor() const
	{
		return m_bTabsHandCursor;
	}

	void SetDotShape(TabDotShape shape)
	{
		m_DotShape = shape;
	}

	TabDotShape GetDotShape() const
	{
		return m_DotShape;
	}

	virtual BOOL SetActiveTab (int iTab);
	virtual BOOL EnsureVisible (int iTab);

	// Scrolling methods:
	BOOL SynchronizeScrollBar (SCROLLINFO* pScrollInfo = NULL);
	CScrollBar* GetScrollBar ()
	{
		return m_bSharedScroll ? &m_wndScrollWnd : NULL;
	}

	virtual void HideSingleTab (BOOL bHide = TRUE);
	void HideActiveWindowHorzScrollBar ();

	virtual DROPEFFECT OnDragEnter(COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
	virtual DROPEFFECT OnDragOver(COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);

	static HICON GetDocumentIcon (UINT nCmdID);

	virtual void OnDraw(CDC* pDC);

	virtual void CleanUp ();

protected:
	virtual CWnd* FindTargetWnd (const CPoint& pt);

	virtual void AdjustTabs ();
	virtual void AdjustTabsScroll ();
	virtual void AdjustWndScroll ();

	void RelayEvent(UINT message, WPARAM wParam, LPARAM lParam);
	void SetTabsHeight ();
	void UpdateScrollControls();
	void ReposButtons (CPoint pt, CSize sizeButton, BOOL bHide, int nButtonMargin);

	void SetupTabCloseButton(CBCGPTabInfo* pTab, int nIndex);
	void AdjustTooltipRect(CBCGPTabInfo* pTab, BOOL bShowTooltip);

	void UpdateScrollButtonsState ();
	void SetScrollButtons ();

	virtual int GetButtonState (CBCGTabButton* pButton);

// Attributes
public:
	virtual BOOL IsPtInTabArea (CPoint point) const;
	virtual BOOL AllowDestroyEmptyTabbedBar () const {return TRUE;}
	virtual int GetTabFromPoint (CPoint& pt) const;

	void SetFlatFrame (BOOL bFlat = TRUE, BOOL bRepaint = TRUE)
	{
		if (m_bFlatFrame != bFlat)
		{
			m_bFlatFrame = bFlat;

			if (bRepaint && GetSafeHwnd () != NULL)
			{
				Invalidate ();
				UpdateWindow ();
			}
		}
	}

	BOOL IsFlatFrame () const
	{
		return m_bFlatFrame;
	}

	void AutoSizeWindow (BOOL bAutoSize = TRUE)
	{
		m_bAutoSizeWindow = bAutoSize;
	}

	void GetWndArea (CRect& rect) const
	{
		rect = m_rectWndArea;
	}

	void GetTabArea (CRect& rectTabAreaTop, CRect& rectTabAreaBottom) const
	{
		rectTabAreaTop.SetRectEmpty ();
		rectTabAreaBottom.SetRectEmpty ();
		(m_location == LOCATION_TOP) ? rectTabAreaTop = m_rectTabsArea : 
										rectTabAreaBottom = m_rectTabsArea;
	}

	virtual void GetTabsRect (CRect& rect) const
	{
		rect = m_rectTabsArea;
	}

	virtual int GetTabsHeight () const
	{
		if (!m_bSharedScroll)
		{
			int nVisiableTabs = GetVisibleTabsNum ();
			if ((m_bHideSingleTab && nVisiableTabs <= 1) ||
				(m_bHideNoTabs && nVisiableTabs == 0))
			{
				return 0;
			}
		}

		return m_nTabsHeight;
	}

	virtual BOOL IsFlatTab () const
	{
		return m_bFlat;
	}

	BOOL IsSharedScroll () const
	{
		return m_bSharedScroll;
	}

	void SetScrollButtonFullSize(BOOL bSet = TRUE)
	{
		m_bScrollButtonFullSize = bSet;
	}

	BOOL IsScrollButtonFullSize() const
	{
		return m_bScrollButtonFullSize;
	}

	virtual BOOL Is3DStyle () const
	{
		return !m_bFlat && !m_bIsOneNoteStyle && !m_bIsVS2005Style && !m_bLeftRightRounded && !m_bIsPointerStyle && !m_bIsDotsStyle && !m_bIsUnderlineStyle;
	}

	virtual BOOL IsOneNoteStyle () const
	{
		return m_bIsOneNoteStyle;
	}

	virtual BOOL IsVS2005Style () const
	{
		return m_bIsVS2005Style;
	}

	virtual BOOL IsLeftRightRounded () const
	{
		return m_bLeftRightRounded;
	}

	virtual BOOL IsPointerStyle() const
	{
		return m_bIsPointerStyle;
	}

	virtual BOOL IsDotsStyle() const
	{
		return m_bIsDotsStyle;
	}

	virtual BOOL IsUnderlineStyle() const
	{
		return m_bIsUnderlineStyle;
	}

	void HideNoTabs (BOOL bHide = TRUE);

	void SetActiveTabBoldFont (BOOL bIsBold = TRUE);
	BOOL IsActiveTabBoldFont () const
	{
		return m_bIsActiveTabBold;
	}

	void SetDrawFrame (BOOL bDraw = TRUE);
	BOOL IsDrawFrame () const
	{
		return m_bDrawFrame;
	}

	static BOOL m_bEnableActivate;

	virtual int	GetFirstVisibleTabNum () const
	{
		// For m_bIsOneNoteStyle only!
		if (m_arTabIndexs.GetSize() != m_arTabs.GetSize() || m_arTabIndexs.GetSize() == 0)
		{
			return m_nFirstVisibleTab;
		}

		return m_arTabIndexs[0];
	}

	virtual void SwapTabs	(int nFisrtTabID, int nSecondTabID);
	virtual void MoveTab	(int nSource, int nDest);

	virtual void OnShowTabDocumentsMenu (CPoint point);

	void SetTabMaxWidth (int nTabMaxWidth);
	int GetTabMaxWidth () const
	{
		return m_nTabMaxWidth;
	}

	void SetResizeMode (ResizeMode resizeMode);
	ResizeMode GetResizeMode () const
	{
		return m_ResizeMode;
	}
	void StopResize (BOOL bCancel);
	virtual BOOL IsMDITabGroup () const;
	void ActivateMDITab (int nTab = -1);
	void SetActiveInMDITabGroup (BOOL bActive)  {m_bIsActiveInMDITabGroup = bActive;}

	virtual BOOL IsActiveInMDITabGroup () const {return m_bIsActiveInMDITabGroup;}

	void SetButtonsVisible (BOOL bVisible = TRUE);
	BOOL AreButtonsVisible () const
	{
		return m_bButtonsVisible;
	}

	void SetMDIFocused(BOOL bIsFocused = TRUE);
	virtual BOOL IsMDIFocused() const { return m_bIsMDIFocused; }

	virtual void OnChangeHighlightedTab(int iPrevHighlighted);

	int GetPointerAreaHeight() const;

	void SetAlowPartialDrawing(BOOL bSet = TRUE)
	{
		m_bAlowPartialDrawing = bSet;
	}

	BOOL IsPartialDrawingAllowed() const
	{
		return m_bAlowPartialDrawing;
	}

protected:
	BOOL			m_bFlat;			// Is flat (Excell-like) mode
	BOOL			m_bIsOneNoteStyle;	// Is OneNote-like mode
	BOOL			m_bIsVS2005Style;	// Is VS 2005 MDI-like mode
	BOOL			m_bLeftRightRounded;// Is VS 2005 two-side rounded tab
	BOOL			m_bIsPointerStyle;	// Office 2013 pointer style
	BOOL			m_bIsDotsStyle;		// "Dots" style: each tab looks like a dot
	BOOL			m_bIsUnderlineStyle;
	BOOL			m_bSharedScroll;	// Have a scrollbar shared with active window
	BOOL			m_bScroll;			// Scroll buttons
	BOOL			m_bCloseBtn;		// Close button
	BOOL			m_bScrollButtonFullSize;
	BOOL			m_bAlowPartialDrawing;
	
	int				m_nScrollBarRight;	// Scrollbar right side
	CRect			m_rectTabsArea;		// Tabs area
	CRect			m_rectWndArea;		// Child window area
	
	CBCGPScrollBar	m_wndScrollWnd;		// Active window horizontal scroll bar

	CBCGTabButton	m_btnScrollLeft;
	CBCGTabButton	m_btnScrollRight;
	CBCGTabButton	m_btnScrollFirst;
	CBCGTabButton	m_btnScrollLast;
	CBCGTabButton	m_btnClose;

	CList<HWND, HWND>	m_lstButtons;

	BOOL			m_bButtonsVisible;

	int				m_nTabsHorzOffset;
	int				m_nTabsHorzOffsetMax;
	int				m_nTabsTotalWidth;
	int				m_nFirstVisibleTab;

	int				m_nHorzScrollWidth;	// Shared scrollbar width
	CRect			m_rectTabSplitter;	// Splitter between tabs and scrollbar
	BOOL			m_bTrackSplitter;	// Is splitter tracked?


	BOOL			m_bFlatFrame;		// Is frame flat?
	
	BOOL			m_bAutoSizeWindow;	// Auto-resize tab windows
	

	BOOL			m_bTransparent;
	BOOL			m_bTopEdge;
	BOOL			m_bDrawFrame;		// Draw frame around window area
	
	BOOL			m_bHideNoTabs;		// Hide tabs are when no tabs are available

	BOOL				m_bIsActiveTabBold;				// Active tab text is drawing bold
	TabCloseButtonMode	m_TabCloseButtonMode;

	CRect					m_rectCurrClip;

	BOOL					m_bTabDocumentsMenu;
	TabsDocumentsMenuStyle	m_TabDocumentsMenuStyle;

	BOOL			m_bHiddenDocuments;
	BOOL			m_bNewTab;
	CString			m_strNewTabLabel;
	int				m_nTabMaxWidth;

	ResizeMode		m_ResizeMode;
	CRect			m_rectResize;
	CRect			m_rectResizeDrag;
	CRect			m_rectResizeBounds;
	BOOL			m_bResize;

	BOOL			m_bIsActiveInMDITabGroup;
	BOOL			m_bIsMDIFocused;

	BOOL			m_bTabAnimationSupport;
	BOOL			m_bProcessAnimation;
	double			m_dblTabAnimationTime;
	int				m_nTabAnimationVal;

	BOOL			m_bTabsHandCursor;
	TabDotShape		m_DotShape;

	int				m_nTabStateAnimationCount;
	UINT			m_nTabStateAnimationDelay;

	static CMap<UINT,UINT,HICON,HICON>	m_mapDocIcons;

	CBCGPSDPlaceMarkerWnd*	m_pDragRectWnd;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBCGPTabWnd)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual BOOL OnCommand(WPARAM wParam, LPARAM lParam);
	//}}AFX_VIRTUAL
	virtual BOOL OnNotify(WPARAM, LPARAM lParam, LRESULT* pResult);
	
	virtual void Draw3DTab (CDC* pDC, CBCGPTabInfo* pTab, BOOL bActive);
	virtual void DrawFlatTab (CDC* pDC, CBCGPTabInfo* pTab, BOOL bActive);
	virtual void DrawTabDot(CDC* pDC, CBCGPTabInfo* pTab, BOOL bActive);

	virtual void DrawResizeDragRect (CRect& rectNew, CRect& rectOld);
	virtual void OnRTLChanged (BOOL bIsRTL);

	virtual CFont* GetTabFont();

	virtual void OnAnimationValueChanged(double dblOldValue, double dblNewValue);
	virtual void OnAnimationFinished();

	virtual CWnd* DoCreateOnPlaceHolder(CWnd* pParent, UINT nID, DWORD dwStyle, const CRect& rect, LPVOID lParam);

// Implementation
public:
	virtual ~CBCGPTabWnd();

	// Generated message map functions
protected:
	//{{AFX_MSG(CBCGPTabWnd)
	afx_msg void OnDestroy();
	afx_msg void OnPaint();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnCancelMode();
	afx_msg void OnSysColorChange();
	afx_msg void OnSettingChange(UINT uFlags, LPCTSTR lpszSection);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnWindowPosChanged(WINDOWPOS FAR* lpwndpos);
	afx_msg void OnWindowPosChanging(WINDOWPOS FAR* lpwndpos);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	//}}AFX_MSG
	afx_msg LRESULT OnBCGUpdateToolTips(WPARAM, LPARAM);
	afx_msg LRESULT OnPrintClient(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnGestureEvent(WPARAM wp, LPARAM lp);
	DECLARE_MESSAGE_MAP()
};

#endif // BCGP_NO_TABCTRL

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BCGPTABWND_H__2ADBF81D_1F51_45D8_9FDB_FF3B13F30CAE__INCLUDED_)
