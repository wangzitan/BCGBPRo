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


#if !defined(AFX_BCGPDIALOG_H__09800B11_4F60_11D4_AB0C_000000000000__INCLUDED_)
#define AFX_BCGPDIALOG_H__09800B11_4F60_11D4_AB0C_000000000000__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "BCGCBPro.h"
#include "bcgpdlgimpl.h"
#include "BCGPButton.h"
#include "BCGPStatic.h"
#include "BCGPAnimationManager.h"
#include "BCGPScrollBar.h"

class CBCGPLocalResource;
class CBCGPToolBarImages;
class CBCGPClosePopupDialogImpl;

/////////////////////////////////////////////////////////////////////////////
// CBCGPDialog dialog

/// <summary>
/// The CBCGPDialog class extends the functionality of the CDialog class. It allows to define a background image or color for the dialog box window and handles all the palette an transparency issues.
/// To take advantage of CBCGPDialog features you just need to derive your dialog class from the CBCGPDialog class instead of CDialog class.
/// When you call SetBackgroundImage method and pass HBITMAP as the image handle, the CBCGPDialog object will take ownership over the image and delete it only if the m_bAutoDestroyBmp flag is TRUE. The current background image will be deleted if you call the SetBackgroundImage again or by OnDestroy message handler. Images loaded from the resource file  are always destroyed automatically.</summary>
/// <remarks></remarks>

class BCGCBPRODLLEXPORT CBCGPDialog : public CDialog
{
	friend class CBCGPPopupMenu;
	friend class CBCGPDropDownList;
	friend class CBCGPContextMenuManager;
	friend class CBCGPRibbonBackstageViewItemForm;
	friend class CBCGPDlgImpl;
	friend class CBCGPVisualManager2007;
	friend class CBCGPVisualManagerCarbon;
	friend class CBCGPEdit;
	friend class CBCGPGridPopupDlgItem;
	friend class CBCGPRibbonButton;
	friend class CBCGPRibbonBackstageViewPanel;
	friend class CBCGPMenuButton;
	friend class CBCGPGlobalUtils;

	DECLARE_DYNAMIC(CBCGPDialog)

// Construction
public:
	/// <summary>
	/// Constructs the object</summary>
	CBCGPDialog();

	/// <summary>
	/// Constructs the object</summary>
	/// <param name="nIDTemplate">Specifies the dialog template resource ID.</param>
	/// <param name="pParentWnd">A pointer to the parent window</param>
	CBCGPDialog(UINT nIDTemplate, CWnd* pParentWnd = NULL);

	/// <summary>
	/// Constructs the object</summary>
	/// <param name="lpszTemplateName">Specifies a dialog template resource name .</param>
	/// <param name="pParentWnd">A pointer to the parent window</param>
	CBCGPDialog(LPCTSTR lpszTemplateName, CWnd* pParentWnd = NULL);

protected:
	void CommonConstruct ();

// Attributes:
public:
	// Visualization:

	/// <summary>
	/// Automatically subclasses all hosted controls in order to take advantage of Visual Manager themes and the "Aero Glass" technology.</summary>
	/// <param name="bEnable">TRUE - enables Visual Manager theme support for hosted controls. FALSE - disables Visual Manager theme support for hosted controls.</param>
	/// <param name="bNCArea">If TRUE, the non-client area of dialog will be displayed using the current Visual Style</param>
	/// <param name="plstNonSubclassedItems">List of controls that shouldn't be subclassed.</param>
	/// <remarks>
	/// Call this function with bEnable TRUE to automatically subclass all hosted controls in order to take advantage of Visual Manager themes and the "Aero Glass" technology.</remarks>
	void EnableVisualManagerStyle (BOOL bEnable = TRUE, BOOL bNCArea = FALSE, const CList<UINT,UINT>* plstNonSubclassedItems = NULL);

	BOOL IsVisualManagerStyle () const
	{
		return m_Impl.m_bVisualManagerStyle;
	}

	BOOL IsVisualManagerNCArea () const
	{
		return m_Impl.m_bVisualManagerNCArea;
	}

	BOOL IsBackstageMode() const
	{
		return m_Impl.m_bBackstageMode;
	}

	enum BackgroundLocation
	{
		BACKGR_TILE,
		BACKGR_TOPLEFT,
		BACKGR_TOPRIGHT,
		BACKGR_BOTTOMLEFT,
		BACKGR_BOTTOMRIGHT,
	};

	/// <summary>
	/// This method has effect only under Windows Vista with the Aero theme enabled.
	/// Call EnableAero to display "Aero Glass" either on the whole dialog, or on some edges. 
	/// To enable glass for the whole dialog, just set _BCGPMARGINS::cyTopHeight to the height of the client area.See the SkinnedDialog sample for more info.</summary>
	/// <param name="margins">Specifies the margins for the "glass" areas.</param>
	BOOL EnableAero (BCGPMARGINS& margins);

	/// <summary></summary>
	/// <param name="margins"></param>
	void GetAeroMargins (BCGPMARGINS& margins) const;

	/// <summary></summary>
	BOOL HasAeroMargins () const
	{
		return m_Impl.HasAeroMargins ();
	}

	// Allow to drag the dialog box by clicking inside the client area
	void EnableDragClientArea(BOOL bEnable = TRUE);
	BOOL IsDragClientAreaEnabled() const
	{
		return m_Impl.IsDragClientAreaEnabled();
	}

	// Layout:
	
	/// <summary>
	/// Call this method to enable or disable layout management for a dialog box.
	/// If you specify a custom layout manager (pRTC), it must be derived from CBCGPControlsLayout or one of its descendants.</summary>
	/// <param name="bEnable">TRUE - enable layout management, FALSE - disable layout management.</param>
	/// <param name="pRTC">A pointer to a runtime class information of custom layout manager.</param>
	/// <param name="bResizeBox">TRUE - display resize box; FALSE - don't display resize box.</param>
	void EnableLayout(BOOL bEnable = TRUE, CRuntimeClass* pRTC = NULL, BOOL bResizeBox = TRUE);
	BOOL IsLayoutEnabled() const
	{
		return m_Impl.IsLayoutEnabled();
	}

	CBCGPControlsLayout* GetLayout()
	{
		return m_Impl.m_pLayout;
	}

	virtual void AdjustControlsLayout()
	{
		m_Impl.AdjustControlsLayout();
	}

	void UpdateCaptionButtons ()
	{
		m_Impl.UpdateCaptionButtons();
	}

	void EnableLoadWindowPlacement(BOOL bEnable = TRUE, LPCTSTR szWindowPlacementProfile = NULL)
	{
		m_Impl.m_bLoadWindowPlacement = bEnable;
		m_Impl.m_strWindowPlacementProfile = szWindowPlacementProfile;
	}

	BOOL IsWindowPlacementEnabled() const
	{
		return m_Impl.m_bLoadWindowPlacement;
	}

	void SetWhiteBackground(BOOL bSet = TRUE);
	BOOL IsWhiteBackground() const
	{
		return m_Impl.m_bIsWhiteBackground;
	}
	
	void SetGroupBoxesDrawByParent(BOOL bSet = TRUE);
	BOOL IsGroupBoxesDrawByParent() const
	{
		return m_Impl.m_bGroupBoxesDrawByParent;
	}

	void AutoResizeControls(const CList<UINT,UINT>* plstNonAutoResizedItems = NULL)
	{
		m_Impl.AutoResizeControls(plstNonAutoResizedItems);
	}

	// Collapse/expand controls support:
	void EnableExpand(UINT nExpandCheckBoxCtrlID, LPCTSTR lpszExpandLabel = NULL, LPCTSTR lszCollapseLabel = NULL);
	void EnableExpand(HWND hwndExpandCheckBoxCtrl, LPCTSTR lpszExpandLabel = NULL, LPCTSTR lszCollapseLabel = NULL);
	
	BOOL Expand(BOOL bExpand = TRUE);
	BOOL IsExpanded() const
	{
		return m_bIsExpanded;
	}

	void SetExpandAreaSpecialBackground(BOOL bSet, BOOL bRedraw = TRUE);
	BOOL IsExpandAreaSpecialBackground() const
	{
		return m_bExpandAreaSpecialBackground;
	}

	virtual void OnBeforeExpand() {}
	virtual void OnAfterExpand() {}

	virtual void OnRTLChanged (BOOL bIsRTL);
	
	void SetEditBoxesVerticalAlignment(int nAlignment)	// TA_TOP, TA_CENTER or TA_BOTTOM, single line only
	{
		m_Impl.SetEditBoxesVerticalAlignment(nAlignment);
	}

#if (!defined _BCGSUITE_) && (!defined _BCGSUITE_INC_)
	void SetEditBoxesSimplifiedBrowseIcon(BOOL bSet = TRUE)
	{
		m_Impl.SetEditBoxesSimplifiedBrowseIcon(bSet);
	}
#endif

protected:
	CBrush						m_brBkgr;
	BOOL						m_bWasMaximized;
	CBCGPDlgImpl				m_Impl;
	BOOL						m_bIsLocal;
	CBCGPLocalResource*			m_pLocaRes;
	CRect						m_rectBackstageWatermark;
	CBCGPToolBarImages*			m_pBackstageWatermark;
	CBCGPClosePopupDialogImpl*	m_pParentEdit;
	BOOL						m_bParentClosePopupDlgNotified;
	BOOL						m_bIsExpanded;
	CList<HWND, HWND>			m_lstCtrlsInCollapseArea;
	HWND						m_hwndExpandCheckBoxCtrl;
	CBCGPButton					m_btnExpandCheckBox;
	int							m_nExpandedAreaHeight;
	int							m_nExpandedCheckBoxHeight;
	CString						m_strExpandLabel;
	CString						m_strCollapseLabel;
	BOOL						m_bExpandAreaSpecialBackground;
	BOOL						m_bIsRibbonStartPage;
	BOOL						m_bCancelModeCapturedWindow;
	BOOL						m_bDisableShadows;

// Operations:
public:
	void SetBackgroundColor (COLORREF color, BOOL bRepaint = TRUE);
	void SetBackgroundImage (HBITMAP hBitmap, 
							BackgroundLocation location = BACKGR_TILE,
							BOOL bAutoDestroy = TRUE,
							BOOL bRepaint = TRUE);
	BOOL SetBackgroundImage (UINT uiBmpResId,
							BackgroundLocation location = BACKGR_TILE,
							BOOL bRepaint = TRUE);

	void ClosePopupDlg(LPCTSTR lpszEditValue = NULL, BOOL bOK = TRUE, DWORD_PTR dwUserData = 0);

	void SetControlInfoTip(UINT nCtrlID, LPCTSTR lpszInfoTip, DWORD dwVertAlign = DT_TOP, BOOL bRedrawInfoTip = FALSE, CBCGPControlInfoTip::BCGPControlInfoTipStyle style = CBCGPControlInfoTip::BCGPINFOTIP_Info, BOOL bIsClickable = FALSE, const CPoint& ptOffset = CPoint(0, 0));
	void SetControlInfoTip(CWnd* pWndCtrl, LPCTSTR lpszInfoTip, DWORD dwVertAlign = DT_TOP, BOOL bRedrawInfoTip = FALSE, CBCGPControlInfoTip::BCGPControlInfoTipStyle style = CBCGPControlInfoTip::BCGPINFOTIP_Info, BOOL bIsClickable = FALSE, const CPoint& ptOffset = CPoint(0, 0));

	CWnd* GetInfoTipControl() const
	{
		return CWnd::FromHandle(m_Impl.m_hwndInfoTipCurr);
	}

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBCGPDialog)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual INT_PTR DoModal();
	protected:
	virtual BOOL OnCommand(WPARAM wParam, LPARAM lParam);
	//}}AFX_VIRTUAL
	virtual void PreInitDialog();

	virtual BOOL IsLightBox() const
	{
		return FALSE;
	}

public:
	BOOL Create(UINT nIDTemplate, CWnd* pParentWnd = NULL);
	BOOL Create(LPCTSTR lpszTemplateName, CWnd* pParentWnd = NULL);

	virtual void OnOK();
	virtual void OnCancel();

	virtual BOOL OnSetPlacement(WINDOWPLACEMENT& wp);

public:
	virtual void OnDrawBackstageWatermark(CDC* /*pDC*/, CRect /*rect*/)	{}
	virtual int GetRibbonStartPageLeftPaneWidth() { return 0; }

#if (!defined _BCGSUITE_) && (!defined _BCGSUITE_INC_) && (!defined BCGP_EXCLUDE_RIBBON)
	virtual void OnDrawRibbonBackgroundImage(CDC* pDC, CRect rect);
#endif

// Implementation
protected:

	//{{AFX_MSG(CBCGPDialog)
	afx_msg void OnActivate(UINT nState, CWnd *pWndOther, BOOL bMinimized);
	afx_msg BOOL OnNcActivate(BOOL bActive);
	afx_msg void OnEnable(BOOL bEnable);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnDestroy();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnSysColorChange();
	afx_msg void OnSettingChange(UINT uFlags, LPCTSTR lpszSection);
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnNcPaint();
	afx_msg void OnNcMouseMove(UINT nHitTest, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg BCGNcHitTestType OnNcHitTest(CPoint point);
	afx_msg void OnNcCalcSize(BOOL bCalcValidRects, NCCALCSIZE_PARAMS FAR* lpncsp);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnWindowPosChanged(WINDOWPOS FAR* lpwndpos);
	afx_msg void OnGetMinMaxInfo(MINMAXINFO FAR* lpMMI);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	//}}AFX_MSG
	afx_msg void OnStyleChanged(int nStyleType, LPSTYLESTRUCT lpStyleStruct);
	afx_msg LRESULT OnDWMCompositionChanged(WPARAM,LPARAM);
	afx_msg LRESULT OnChangeVisualManager (WPARAM, LPARAM);
	afx_msg LRESULT OnSetText(WPARAM, LPARAM);
	afx_msg LRESULT OnPowerBroadcast(WPARAM wp, LPARAM);
	afx_msg LRESULT OnBCGUpdateToolTips (WPARAM, LPARAM);
	afx_msg BOOL OnNeedTipText(UINT id, NMHDR* pNMH, LRESULT* pResult);
	afx_msg LRESULT OnDPIChanged(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnThemeChanged(WPARAM wp, LPARAM lp);
	DECLARE_MESSAGE_MAP()

	virtual void SetActiveMenu (CBCGPPopupMenu* pMenu);
	
	void EnableBackstageMode()
	{
		m_Impl.EnableBackstageMode();
	}
};

/////////////////////////////////////////////////////////////////////////////
// CBCGPLightBoxDialog dialog

class CBCGPLightBoxShadow;

enum BCGPLightBoxDialogLocation
{
	BCGPLightBoxDialogLocation_Center = 0,
	BCGPLightBoxDialogLocation_Left,
	BCGPLightBoxDialogLocation_Right,
	BCGPLightBoxDialogLocation_Top,
	BCGPLightBoxDialogLocation_Bottom,
};

enum BCGPLightBoxDialogTransitionType
{
	BCGPLightBoxDialogTransitionType_None = 0,
	BCGPLightBoxDialogTransitionType_Default,	// BCGPLightBoxDialogTransitionType_Slide for the side-located, BCGPLightBoxDialogTransitionType_Appear - center
	BCGPLightBoxDialogTransitionType_Slide,
	BCGPLightBoxDialogTransitionType_Appear,
};

enum BCGPLightBoxDialogCaptionStyle
{
	BCGPLightBoxDialogCaptionStyle_None = 0,
	BCGPLightBoxDialogCaptionStyle_Default,	// Specified by currently selected visual theme
	BCGPLightBoxDialogCaptionStyle_Regular,
	BCGPLightBoxDialogCaptionStyle_Bold,
#if (!defined _BCGSUITE_) && (!defined _BCGSUITE_INC_)
	BCGPLightBoxDialogCaptionStyle_Large,
#endif
};


struct BCGCBPRODLLEXPORT CBCGPLightBoxDialogOptions
{
	CBCGPLightBoxDialogOptions()
	{
		m_CaptionStyle = BCGPLightBoxDialogCaptionStyle_Default;
		m_clrCaption = CLR_DEFAULT;
		m_clrBorder = CLR_DEFAULT;
		m_clrShading = RGB(50, 50, 50);
		m_nAlpha = 100;
		m_bDrawFrameShadow = TRUE;
		m_Location = BCGPLightBoxDialogLocation_Center;
		m_bCloseLightBoxByClickingOutside = TRUE;

		m_dblShowTransitionSpeed = m_dblDismissTransitionSpeed = 0.2;
		m_ShowTransitionAnimation = m_DismissTransitionAnimation = CBCGPAnimationManager::BCGPANIMATION_AccelerateDecelerate;
		m_ShowTransitionType = m_DismissTransitionType = BCGPLightBoxDialogTransitionType_Default;
	}

	BCGPLightBoxDialogCaptionStyle				m_CaptionStyle;
	COLORREF									m_clrCaption;
	COLORREF									m_clrBorder;
	COLORREF									m_clrShading;
	BYTE										m_nAlpha;
	BOOL										m_bDrawFrameShadow;
	BCGPLightBoxDialogLocation					m_Location;
	BOOL										m_bCloseLightBoxByClickingOutside;

	BCGPLightBoxDialogTransitionType			m_ShowTransitionType;
	CBCGPAnimationManager::BCGPAnimationType	m_ShowTransitionAnimation;
	double										m_dblShowTransitionSpeed;	// in seconds

	BCGPLightBoxDialogTransitionType			m_DismissTransitionType;
	CBCGPAnimationManager::BCGPAnimationType	m_DismissTransitionAnimation;
	double										m_dblDismissTransitionSpeed;	// in seconds
};

class CBCGPLightBoxDialogButton : public CBCGPButton
{
	virtual void OnFillBackground (CDC* pDC, const CRect& rectClient);
	virtual void OnDrawBorder (CDC* pDC, CRect& rectClient, UINT uiState);
};

#pragma warning (push) 
#pragma warning (disable : 4355)

template<>
inline TBCGPInternalScrollBarWrapperWnd<CBCGPDialog>::TBCGPInternalScrollBarWrapperWnd(LPCTSTR lpszParam)
	: CBCGPDialog(lpszParam)
	, CBCGPInternalScrollBarWrapper(this)
{
}

template<>
inline TBCGPInternalScrollBarWrapperWnd<CBCGPDialog>::TBCGPInternalScrollBarWrapperWnd(UINT nIDParam)
	: CBCGPDialog(nIDParam)
	, CBCGPInternalScrollBarWrapper(this)
{
}

#pragma warning (pop)

class BCGCBPRODLLEXPORT CBCGPLightBoxDialog :	public TBCGPInternalScrollBarWrapperWnd<CBCGPDialog>,
												public CBCGPAnimationManager
{
	DECLARE_DYNAMIC(CBCGPLightBoxDialog)

	friend class CBCGPLightBoxShadow;
	friend class CBCGPFrameImpl;
	friend class CBCGPFrameWnd;
	friend class CBCGPMDIFrameWnd;

public:
	CBCGPLightBoxDialog(UINT uiDlgResID, CWnd* pParent = NULL, CBCGPLightBoxDialogOptions* pOptions = NULL);
	CBCGPLightBoxDialog(LPCTSTR lpszDlgResID, CWnd* pParent = NULL, CBCGPLightBoxDialogOptions* pOptions = NULL);
	CBCGPLightBoxDialog(CSize sizeDlg, LPCTSTR lpszCaption, CBCGPLightBoxDialogOptions* pOptions = NULL);

	virtual ~CBCGPLightBoxDialog();

	void DoRepos(BOOL bRedrawShadow = FALSE);
	void Activate();
	void RecalcClosebutton();
	
	BOOL IsDrawBorder() const;
	BOOL IsTransitionAvailable(BOOL bShow) const;
	
	BOOL IsClosing() const
	{
		return m_bIsClosing;
	}

protected:
	void CommonInit(CBCGPLightBoxDialogOptions* pOptions);
	void CommonCreate();

	BCGPLightBoxDialogCaptionStyle GetCaptionStyle() const;
	
	BOOL CloseMe(BOOL bNotify);
	BOOL ShouldCloseOnClick(CWnd* pWnd, CPoint ptScreen);

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBCGPLightBoxDialog)
	public:
	virtual BOOL Create();
	virtual INT_PTR DoModal();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	//}}AFX_VIRTUAL

	virtual void OnAnimationValueChanged(double dblOldValue, double dblNewValue);
	virtual void OnAnimationFinished();

	virtual BOOL OnScroll(UINT nScrollCode, UINT nPos, BOOL bDoScroll = TRUE);
	virtual BOOL OnScrollBy(CSize sizeScroll, BOOL bDoScroll = TRUE);

	virtual BOOL IsLightBox() const
	{
		return TRUE;
	}

	virtual void Dismiss();
	virtual BOOL CanBeClosed()
	{
		return !IsAnimated();
	}
	
	virtual int GetCaptionHeight();
	virtual void OnDrawCaption(CDC* pDC, CRect rectCaption, CRect rectCloseButton, COLORREF clrText);
	virtual void OnDrawCloseButton(CDC* pDC, CRect rectClose, BOOL bIsHighlighted, BOOL bIsPressed);

	//{{AFX_MSG(CBCGPLightBoxDialog)
	afx_msg void OnNcDestroy();
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	afx_msg BOOL OnNcCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg BOOL OnNcActivate(BOOL bActive);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnNcCalcSize(BOOL bCalcValidRects, NCCALCSIZE_PARAMS FAR* lpncsp);
	afx_msg BCGNcHitTestType OnNcHitTest(CPoint point);
	afx_msg void OnNcPaint();
	afx_msg void OnNcMouseMove(UINT nHitTest, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnCancelMode();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	//}}AFX_MSG
	afx_msg void OnOK();
	afx_msg void OnCancel();
	afx_msg LRESULT OnFloatStatus(WPARAM wParam, LPARAM);
	afx_msg LRESULT OnSetText(WPARAM wParam, LPARAM lParam);
	DECLARE_MESSAGE_MAP()

	CBCGPLightBoxShadow*				m_pShadow;
	LPCTSTR								m_lpszDlgTemplateName;
	CSize								m_sizeDlg;
	CPoint								m_ptOffsetDlgtAnim;
	CSize								m_sizeDlgAnim;
	CSize								m_sizeCurr;
	BOOL								m_bActive;
	CBCGPLightBoxDialogOptions			m_LightBoxOptions;
	BOOL								m_bIsHorzScrollBar;
	BOOL								m_bIsVertScrollBar;
	BOOL								m_bIsModalLoop;
	CString								m_strCaption;
	int									m_nCaptionHeight;
	BOOL								m_bHasCloseButton;
	CRect								m_rectClose;
	BOOL								m_bIsCloseButtonPressed;
	BOOL								m_bIsCloseButtonHighlighted;
	BOOL								m_bIsCloseButtonCaptured;
	BCGPLightBoxDialogTransitionType	m_CurrTransitionType;
	BOOL								m_bNotifyClose;
	BOOL								m_bIsShowAnimation;
	BOOL								m_bIsClosing;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BCGPDIALOG_H__09800B11_4F60_11D4_AB0C_000000000000__INCLUDED_)
