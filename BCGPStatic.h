#if !defined(AFX_BCGPSTATIC_H__A34CC2CA_E7ED_4D03_93F4_E2513841C22C__INCLUDED_)
#define AFX_BCGPSTATIC_H__A34CC2CA_E7ED_4D03_93F4_E2513841C22C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

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
// BCGPStatic.h : header file
//

#include "BCGCBPro.h"

#if (!defined _BCGSUITE_) && (!defined _BCGSUITE_INC_)
#include "BCGPToolBarImages.h"
#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPStatic window

class BCGCBPRODLLEXPORT CBCGPStatic : public CStatic
{
	DECLARE_DYNAMIC(CBCGPStatic)

// Construction
public:
	CBCGPStatic();

// Attributes
public:
	BOOL		m_bOnGlass;
	BOOL		m_bVisualManagerStyle;
	BOOL		m_bBackstageMode;
	COLORREF	m_clrText;
	HFONT		m_hFont;

	CSize GetPictureSize() const
	{
		return m_Picture.GetCount() > 0 ? m_Picture.GetImageSize() : CSize(0, 0);
	}

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBCGPStatic)
	protected:
	virtual void PreSubclassWindow();
	//}}AFX_VIRTUAL

public:
	enum BCGP_SIZE_TO_CONTENT_LINES
	{
		BCGP_SIZE_TO_CONTENT_LINES_AUTO = 0,
		BCGP_SIZE_TO_CONTENT_LINES_SINGLE,
		BCGP_SIZE_TO_CONTENT_LINES_MULTI,
	};

	virtual void SizeToContent(BCGP_SIZE_TO_CONTENT_LINES lines = BCGP_SIZE_TO_CONTENT_LINES_AUTO);

	BOOL SetPicture(UINT nResID, BOOL bDPIAutoScale = TRUE, BOOL bAutoResize = FALSE, HINSTANCE hResInstance = NULL);
	BOOL SetPicture(const CBCGPToolBarImages& src);
	BOOL SetPicture(const CString& strImagePath, BOOL bDPIAutoScale = TRUE, BOOL bAutoResize = FALSE);

protected:
	virtual void DoPaint(CDC* pDC);
	virtual BOOL IsOwnerDraw() const;

// Implementation
public:
	virtual ~CBCGPStatic();

	// Generated message map functions
protected:
	//{{AFX_MSG(CBCGPStatic)
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnPaint();
	afx_msg void OnEnable(BOOL bEnable);
	afx_msg void OnNcPaint();
	afx_msg void OnDestroy();
	//}}AFX_MSG
	afx_msg LRESULT OnBCGSetControlAero (WPARAM, LPARAM);
	afx_msg LRESULT OnBCGSetControlVMMode (WPARAM, LPARAM);
	afx_msg LRESULT OnBCGSetControlBackStageMode (WPARAM, LPARAM);
	afx_msg LRESULT OnSetText (WPARAM, LPARAM lp);
	afx_msg LRESULT OnSetFont (WPARAM wParam, LPARAM lp);
	afx_msg LRESULT OnPrint(WPARAM wParam, LPARAM lp);
	afx_msg LRESULT OnBCGOnChangeUnderline(WPARAM wp, LPARAM);
	DECLARE_MESSAGE_MAP()

	BOOL IsOwnerDrawSeparator(BOOL& bIsHorz);
	COLORREF GetForegroundColor();
	
protected:
	BOOL				m_bDrawTextUnderline;
	BOOL				m_bPreview;
	CBCGPToolBarImages	m_Picture;
};

class BCGCBPRODLLEXPORT CBCGPInfoBoxRenderer
{
public:
	enum BCGP_INFOBOX_CAPTION_STYLE
	{
		BCGP_INFOBOX_CAPTION_REGULAR = 0,
		BCGP_INFOBOX_CAPTION_BOLD,
#if (!defined _BCGSUITE_) && (!defined _BCGSUITE_INC_)
		BCGP_INFOBOX_CAPTION_LARGE,
#endif	
	};

	CBCGPInfoBoxRenderer();

	void SetInfoBoxFont(HFONT hFont);

	BOOL IsTextTruncated() const
	{
		return m_bTextIsTruncated;
	}

	CRect DoDraw(CDC* pDC, const CRect& rectBounds, const CString& strInfo, const CString& strCaption = _T(""), HICON hIcon = NULL, BOOL bCalcRectOnly = FALSE, BOOL bPreview = FALSE);

	int GetTextHorizontalOffset() const
	{
		return m_nTextHorzOffset;
	}
	
	COLORREF					m_clrBackground;
	COLORREF					m_clrForeground;
	COLORREF					m_clrCaption;
	COLORREF					m_clrFrame;
	COLORREF					m_clrLink;
	CRect						m_rectPadding;
	int							m_nHorzAlign;
	int							m_nVertAlign;
	COLORREF					m_clrColorBar;
	int							m_nColorBarWidth;
	BOOL						m_bIsColorBarOnLeft;
	BOOL						m_bDrawShadow;
	CBCGPToolBarImages			m_InfotipIcons;
	int							m_nIconIndex;
	BOOL						m_bDrawMoreTooltipMarker;
	BCGP_INFOBOX_CAPTION_STYLE	m_CaptionStyle;
	BOOL						m_bFixedFrameWidth;
	BOOL						m_bFixedFrameHeight;
	BOOL						m_bRoundedCorners;
	
protected:
	
	CFont* SelectCaptionFont(CDC* pDC, CFont* pFont);

	HFONT						m_hInfoBoxFont;
	CFont						m_fontBold;
	CFont						m_fontUnderline;
	int							m_nRowHeight;
	BOOL						m_bTextIsTruncated;
	CRect						m_rectMore;
	int							m_nTextHorzOffset;
	CRect						m_rectLink;
	CString						m_strLink;
	BOOL						m_bIsFocused;
};

class BCGCBPRODLLEXPORT CBCGPInfoBoxCtrl :	public CBCGPStatic,
											public CBCGPInfoBoxRenderer
{
	DECLARE_DYNAMIC(CBCGPInfoBoxCtrl)

// Construction
public:
	CBCGPInfoBoxCtrl();

// Attributes
public:
	void SetCaption(const CString& strCaption);
	const CString& GetCaption() const
	{
		return m_strCaption;
	}

	void SetIcon(HICON hIcon);
	HICON GetIcon() const
	{
		return m_hIcon;
	}

	void SetLink(const CString& strLink);
	const CString& GetLink() const
	{
		return m_strLink;
	}

// Operations
public:
	BOOL CopyToClipboard(BOOL bWithCaption = TRUE, BOOL bWithLink = FALSE);

protected:
	CRect DoDrawCtrl(CDC* pDC, BOOL bCalcOnly);

// Overrides
public:
	virtual void SizeToContent(BCGP_SIZE_TO_CONTENT_LINES lines = BCGP_SIZE_TO_CONTENT_LINES_AUTO);

protected:
	virtual void DoPaint(CDC* pDC);
	virtual BOOL IsOwnerDraw() const 
	{ 
		return TRUE; 
	}

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	void UpdateTooltips();

	afx_msg void OnDestroy();
	afx_msg LRESULT OnBCGUpdateToolTips (WPARAM, LPARAM);
	afx_msg BOOL OnNeedTipText(UINT id, NMHDR* pNMH, LRESULT* pResult);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	DECLARE_MESSAGE_MAP()

// Attributes
protected:
	CString			m_strCaption;
	HICON			m_hIcon;
	CToolTipCtrl*	m_pToolTip;
};

BCGCBPRODLLEXPORT extern UINT BCGM_INFOBOX_LINK_CLICKED;

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BCGPSTATIC_H__A34CC2CA_E7ED_4D03_93F4_E2513841C22C__INCLUDED_)
