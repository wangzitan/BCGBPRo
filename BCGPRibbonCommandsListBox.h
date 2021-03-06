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
// BCGPRibbonCommandsListBox.h : header file
//

#if !defined(AFX_BCGPRIBBONCOMMANDSLISTBOX_H__86D859ED_7417_446D_9177_83DDA4245C08__INCLUDED_)
#define AFX_BCGPRIBBONCOMMANDSLISTBOX_H__86D859ED_7417_446D_9177_83DDA4245C08__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "BCGCBPro.h"
#include "BCGPListBox.h"

#ifndef BCGP_EXCLUDE_RIBBON

class CBCGPRibbonBar;
class CBCGPBaseRibbonElement;
class CBCGPRibbonCategory;
class CBCGPRibbonSeparator;

/////////////////////////////////////////////////////////////////////////////
// CBCGPRibbonCommandsListBox window

class BCGCBPRODLLEXPORT CBCGPRibbonCommandsListBox : public CBCGPListBox
{
// Construction
public:
	CBCGPRibbonCommandsListBox(	CBCGPRibbonBar* pRibbonBar, 
								BOOL bIncludeSeparator = TRUE,
								BOOL bDrawDefaultIcon = FALSE,
								BOOL bCommandsOnly = FALSE);

// Attributes
public:
	CBCGPBaseRibbonElement* GetSelected () const;
	CBCGPBaseRibbonElement* GetCommand (int nIndex) const;
	int GetCommandIndex (UINT uiID) const;

	BOOL CommandsOnly() const
	{
		return m_bCommandsOnly;
	}

	const CBCGPRibbonBar* GetRibbonBar() const
	{
		return m_pRibbonBar;
	}

	void SetSimpleDraw(BOOL bSet)
	{
		m_bSimpleDraw = bSet;
	}

	BOOL IsSimpleDraw() const
	{
		return m_bSimpleDraw;
	}

protected:
	CBCGPRibbonBar*			m_pRibbonBar;
	int						m_nTextOffset;
	CBCGPRibbonSeparator*	m_pSeparator;
	BOOL					m_bDrawDefaultIcon;
	BOOL					m_bCommandsOnly;
	BOOL					m_bSimpleDraw;

// Operations
public:
	void FillFromCategory (CBCGPRibbonCategory* pCategory, BOOL bExcludeDefaultPanelButtons = FALSE);
	void FillFromIDs (const CList<UINT,UINT>& lstCommands, BOOL bDeep, BOOL bExcludeDefaultPanelButtons = FALSE);
	void FillFromArray (
		const CArray<CBCGPBaseRibbonElement*, CBCGPBaseRibbonElement*>& arElements,
		BOOL bDeep, BOOL bIgnoreSeparators, BOOL bExcludeDefaultPanelButtons = FALSE);
	void FillAll ();

	BOOL AddCommand (CBCGPBaseRibbonElement* pCmd, BOOL bSelect = TRUE, BOOL bDeep = TRUE);
	void RebuildTooltips();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBCGPRibbonCommandsListBox)
	//}}AFX_VIRTUAL

	virtual void OnDrawItemContent(CDC* pDC, CRect rect, int nIndex);

// Implementation
public:
	virtual ~CBCGPRibbonCommandsListBox();

	// Generated message map functions
protected:
	//{{AFX_MSG(CBCGPRibbonCommandsListBox)
	afx_msg void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);
	afx_msg void MeasureItem(LPMEASUREITEMSTRUCT lpMeasureItemStruct);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

#endif // BCGP_EXCLUDE_RIBBON

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BCGPRIBBONCOMMANDSLISTBOX_H__86D859ED_7417_446D_9177_83DDA4245C08__INCLUDED_)
