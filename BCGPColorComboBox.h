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
// BCGPColorComboBox.h : header file
//

#if !defined(AFX_BCGPCOLORCOMBOBOX_H__CE32E477_B93F_4001_A2A1_470A350B67BE__INCLUDED_)
#define AFX_BCGPCOLORCOMBOBOX_H__CE32E477_B93F_4001_A2A1_470A350B67BE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "BCGCBPro.h"
#include "BCGPComboBox.h"

#if (!defined _BCGSUITE_) && (!defined _BCGSUITE_INC_)
	#include "BCGPVisualManagerVS2012.h"
#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPColorComboBox window

class BCGCBPRODLLEXPORT CBCGPColorComboBox : public CBCGPComboBox
{
// Construction
public:
	CBCGPColorComboBox();

// Attributes
public:
	void EnableNoneColor(BOOL bEnable = TRUE, const CString& strName = _T("No color"));
	BOOL IsNoneColorEnabled() const
	{
		return m_bIsNoneColor;
	}

	void SetDefaultColor(COLORREF clrDefault);
	COLORREF GetDefaultColor() const
	{
		return m_clrDefault;
	}

// Operations
public:
	int AddColor(COLORREF color, const CString& strColorName);
	int SetStandardColors();
	int SetSystemColors();
#if (!defined _BCGSUITE_) && (!defined _BCGSUITE_INC_)
	int SetAccentColors(BOOL bIncludeDefault = TRUE);
#endif
	
	COLORREF GetSelectedColor() const;

	COLORREF GetItemColor(int nIndex) const;
	int FindIndexByColor(COLORREF color) const;

	static void SetSystemColorName(int nIndex, const CString& strName);
	static const CString& GetSystemColorName(int nIndex);

#if (!defined _BCGSUITE_) && (!defined _BCGSUITE_INC_)
	static void SetAccentColorName(CBCGPVisualManagerVS2012::AccentColor color, const CString& strName);
	static const CString& GetAccentColorName(CBCGPVisualManagerVS2012::AccentColor color);
#endif

	static BOOL DoDrawItem(CDC* pDC, CRect rect, CFont* pFont, COLORREF color, const CString& strText, BOOL bDisabled, COLORREF clrDefault = RGB(255, 255, 255));

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBCGPColorComboBox)
	virtual void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);
	//}}AFX_VIRTUAL

	virtual CSize GetImageSize() const;

// Implementation
public:
	virtual ~CBCGPColorComboBox();

	// Generated message map functions
protected:
	//{{AFX_MSG(CBCGPColorComboBox)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	static CMap<UINT, UINT, CString, const CString&>	m_mapSystemColorNames;
#if (!defined _BCGSUITE_) && (!defined _BCGSUITE_INC_)
	static CMap<int, int, CString, const CString&>		m_mapAccentColorNames;
#endif
	BOOL												m_bIsNoneColor;
	CString												m_strNoneColor;
	COLORREF											m_clrDefault;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BCGPCOLORCOMBOBOX_H__CE32E477_B93F_4001_A2A1_470A350B67BE__INCLUDED_)
