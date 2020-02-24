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
// BCGPCircularProgressIndicator.h: interface for the CBCGPCircularProgressIndicatorImpl class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BCGPCIRCULARPROGRESSINDICATOR_H__CA3BCAE6_CA45_4214_A278_06D3F1DFD185__INCLUDED_)
#define AFX_BCGPCIRCULARPROGRESSINDICATOR_H__CA3BCAE6_CA45_4214_A278_06D3F1DFD185__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "BCGPCircularGaugeImpl.h"

class CBCGPCircularProgressIndicatorRenderer;

struct BCGCBPRODLLEXPORT CBCGPCircularProgressIndicatorOptions
{
	enum Shape
	{
		BCGPCircularProgressIndicator_First = 0,
		BCGPCircularProgressIndicator_Arc = BCGPCircularProgressIndicator_First,
		BCGPCircularProgressIndicator_Circle,
		BCGPCircularProgressIndicator_Line,
		BCGPCircularProgressIndicator_Pie,
		BCGPCircularProgressIndicator_Gradient,
		BCGPCircularProgressIndicator_Last = BCGPCircularProgressIndicator_Gradient,
	};

	CBCGPCircularProgressIndicatorOptions()
	{
		m_bMarqueeStyle = FALSE;
		m_dblProgressWidth = 0.0;
		m_Shape = BCGPCircularProgressIndicator_Arc;
		m_bClockwiseRotation = TRUE;
		m_bFadeColor = TRUE;
		m_bFadeSize = FALSE;
		m_strPercentageFormat = _T("%1.0f%%");
	}
	
	BOOL CompareWith(const CBCGPCircularProgressIndicatorOptions& other)
	{
		return 
			m_bMarqueeStyle == other.m_bMarqueeStyle &&
			m_dblProgressWidth == other.m_dblProgressWidth &&
			m_Shape == other.m_Shape &&
			m_bFadeColor == other.m_bFadeColor &&
			m_bFadeSize == other.m_bFadeSize &&
			m_bClockwiseRotation == other.m_bClockwiseRotation &&
			m_strPercentageFormat == other.m_strPercentageFormat;
	}

	BOOL		m_bMarqueeStyle;
	double		m_dblProgressWidth;
	Shape		m_Shape;
	BOOL		m_bFadeColor;
	BOOL		m_bFadeSize;
	BOOL		m_bClockwiseRotation;
	CString		m_strPercentageFormat;	// Non-marquee style only
};

struct BCGCBPRODLLEXPORT CBCGPCircularProgressIndicatorColors
{
	CBCGPCircularProgressIndicatorColors()
	{
		CBCGPCircularGaugeColors colors;
		colors.SetTheme(CBCGPCircularGaugeColors::BCGP_CIRCULAR_GAUGE_VISUAL_MANAGER);

		ConvertFrom(colors);
	}

	CBCGPBrush	m_brFill;
	CBCGPBrush	m_brFrameOutline;
	CBCGPBrush	m_brProgressLabel;
	CBCGPBrush	m_brTextLabel;

	CBCGPBrush	m_brProgressFill;
	CBCGPBrush	m_brProgressOutline;
	CBCGPBrush	m_brProgressFillGradient;
	CBCGPBrush	m_brProgressOutlineGradient;
	CBCGPBrush	m_brProgressFillNonComplete;
	CBCGPBrush	m_brProgressOutlineNonComplete;

	void ConvertFrom(const CBCGPCircularGaugeColors& colors)
	{
		m_brFill = colors.m_brFill;
		m_brFrameOutline = colors.m_brFrameOutline;
		m_brProgressFill = colors.m_brPointerOutline;
		
		m_brProgressLabel = colors.m_brPointerOutline;
		m_brTextLabel = colors.m_brText;

		m_brProgressOutline.Empty();
		m_brProgressFillGradient.Empty();
		m_brProgressOutlineGradient.Empty();
		m_brProgressFillNonComplete.Empty();
		m_brProgressOutlineNonComplete.Empty();
	}
};

class BCGCBPRODLLEXPORT CBCGPCircularProgressIndicatorImpl : public CBCGPCircularGaugeImpl
{
	friend class CBCGPCircularProgressIndicatorRenderer;

	DECLARE_DYNCREATE(CBCGPCircularProgressIndicatorImpl)

public:
	CBCGPCircularProgressIndicatorImpl(CBCGPVisualContainer* pContainer = NULL);
	virtual ~CBCGPCircularProgressIndicatorImpl();
	
	void StartMarquee(BOOL bStart = TRUE, double dblSpeed = 0.0 /* 0.0 - 1.0 */);	// SetOptions with m_bMarqueeStyle = TRUE should be called first
	BOOL IsMarqueeStarted() const
	{
		return m_bMarqueeStarted;
	}
	
	void SetMarqueeValue(double dblValue);	// < 0 - clean-up marquee

	double SetPos(double nPos);
	double GetPos() const;

	void SetRange(double nLower, double nUpper);
	void GetRange(double& nLower, double& nUpper) const;

	double OffsetPos(double nPos);

	double SetStep(double nStep);
	double GetStep(BOOL bCalcIfZero = FALSE) const;

	double StepIt();

	virtual void CopyFrom(const CBCGPBaseVisualObject& src);

	void SetColors(CBCGPCircularGaugeColors::BCGP_CIRCULAR_GAUGE_COLOR_THEME theme);
	void SetColors(const CBCGPCircularProgressIndicatorColors& colors);

	const CBCGPCircularProgressIndicatorColors& GetColors() const
	{
		return m_ProgressColors;
	}

	BOOL IsMarqueeStyle() const
	{
		return m_Options.m_bMarqueeStyle;
	}

	const CBCGPCircularProgressIndicatorOptions& GetOptions() const
	{
		return m_Options;
	}

	void SetOptions(const CBCGPCircularProgressIndicatorOptions& options);

	const CString& GetLabel() const
	{
		return m_strLabel;
	}

	void SetLabel(const CString& strLabel);

	virtual void SetRect(const CBCGPRect& rect, BOOL bRedraw = FALSE);
	virtual void OnDraw(CBCGPGraphicsManager* pGM, const CBCGPRect& rectClip, DWORD dwFlags = BCGP_DRAW_STATIC | BCGP_DRAW_DYNAMIC);

protected:
	virtual void OnChangeVisualManager()
	{
		if (m_Colors.IsVisualManagerTheme())
		{
			SetColors(CBCGPCircularGaugeColors::BCGP_CIRCULAR_GAUGE_VISUAL_MANAGER);
			Redraw();
		}
	}

	virtual CBCGPSize GetMinSize() const
	{
		return CBCGPSize(10, 10);
	}

	virtual BOOL GetLevelBarShape(const CBCGPPoint& center, double radius, BOOL bForValue, CBCGPComplexGeometry& shape,
		double dblWidth, double dblOffsetFromFrame, int nScale, double dblValue);

	virtual void OnDrawLevelBar(CBCGPGraphicsManager* pGM, const CBCGPComplexGeometry& shape, const CBCGPBrush& brFill, const CBCGPBrush& brOutline, double lineWidth, BOOL bIsDynamic);

	virtual void OnAnimation(CBCGPVisualDataObject* pDataObject);
	virtual void OnAnimationFinished(CBCGPVisualDataObject* pDataObject);

	virtual BOOL ReposAllSubGauges(CBCGPGraphicsManager* pGM);

	virtual BOOL OnMouseDown(int nButton, const CBCGPPoint& pt);
	virtual void OnMouseUp(int nButton, const CBCGPPoint& pt);
	virtual void OnMouseMove(const CBCGPPoint& pt);
	virtual void OnMouseLeave();
	virtual void OnCancelMode();

	virtual BOOL OnKeyboardDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	virtual BOOL IsReturnKeySupported()
	{
		return m_bIsInteractiveMode;
	}

	void AdjustLayout(const CBCGPRect& rect);
	
	virtual CString OnPreparePercentageLabel(double nPos) const;

	virtual HRESULT get_accRole(VARIANT varChild, VARIANT *pvarRole);
	virtual HRESULT get_accValue(VARIANT varChild, BSTR *pszValue);
	virtual HRESULT get_accState(VARIANT varChild, VARIANT *pvarState);

	CString									m_strLabel;
	double									m_dblMarquee;
	BOOL									m_bMarqueeStarted;
	CBCGPCircularProgressIndicatorOptions	m_Options;
	CBCGPCircularProgressIndicatorColors	m_ProgressColors;
	BOOL									m_bIsPressed;
	CBCGPBrush								m_brFocus;
	CBCGPBrush								m_brPressed;
	BOOL									m_bIsSmall;
	double									m_dblStep;
	CBCGPCircularProgressIndicatorRenderer*	m_pRenderer;
};

class BCGCBPRODLLEXPORT CBCGPCircularProgressIndicatorCtrl : public CBCGPVisualCtrl
{
	DECLARE_DYNAMIC(CBCGPCircularProgressIndicatorCtrl)
		
public:
	CBCGPCircularProgressIndicatorCtrl()
	{
		m_pCircularProgressIndicator = NULL;
		m_nDlgCode = DLGC_WANTARROWS | DLGC_WANTCHARS;
		EnableTooltip();
	}
	
	virtual ~CBCGPCircularProgressIndicatorCtrl()
	{
		if (m_pCircularProgressIndicator != NULL)
		{
			delete m_pCircularProgressIndicator;
		}
	}
	
	virtual CBCGPCircularProgressIndicatorImpl* GetCircularProgressIndicator()
	{
		if (m_pCircularProgressIndicator == NULL)
		{
			m_pCircularProgressIndicator = new CBCGPCircularProgressIndicatorImpl();
		}
		
		return m_pCircularProgressIndicator;
	}

	double SetPos(double nPos)
	{
		CBCGPCircularProgressIndicatorImpl* pProgress = GetCircularProgressIndicator();
		ASSERT_VALID(pProgress);

		return pProgress->SetPos(nPos);
	}
	
	double GetPos() const
	{
		const CBCGPCircularProgressIndicatorImpl* pProgress = 
			const_cast<CBCGPCircularProgressIndicatorCtrl*>(this)->GetCircularProgressIndicator();
		ASSERT_VALID(pProgress);

		return pProgress->GetPos();
	}

	void SetRange(double nLower, double nUpper)
	{
		CBCGPCircularProgressIndicatorImpl* pProgress = GetCircularProgressIndicator();
		ASSERT_VALID(pProgress);

		pProgress->SetRange(nLower, nUpper);
	}

	void SetRange32(double nLower, double nUpper)
	{
		SetRange(nLower, nUpper);
	}

	void GetRange(double& nLower, double& nUpper) const
	{
		const CBCGPCircularProgressIndicatorImpl* pProgress = 
			const_cast<CBCGPCircularProgressIndicatorCtrl*>(this)->GetCircularProgressIndicator();
		ASSERT_VALID(pProgress);

		pProgress->GetRange(nLower, nUpper);
	}

	void GetRange(int& nLower, int& nUpper) const
	{
		double dblLower = 0.0;
		double dblUpper = 0.0;

		GetRange(dblLower, dblUpper);

		nLower = (int)dblLower;
		nUpper = (int)dblUpper;
	}

	double OffsetPos(double nPos)
	{
		CBCGPCircularProgressIndicatorImpl* pProgress = GetCircularProgressIndicator();
		ASSERT_VALID(pProgress);
		
		return pProgress->OffsetPos(nPos);
	}
	
	double SetStep(double nStep)
	{
		CBCGPCircularProgressIndicatorImpl* pProgress = GetCircularProgressIndicator();
		ASSERT_VALID(pProgress);
		
		return pProgress->SetStep(nStep);
	}
	
	double StepIt()
	{
		CBCGPCircularProgressIndicatorImpl* pProgress = GetCircularProgressIndicator();
		ASSERT_VALID(pProgress);
		
		return pProgress->StepIt();
	}

protected:
	virtual CBCGPBaseVisualObject* GetVisualObject()
	{
		return GetCircularProgressIndicator();
	}
	
	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	virtual void OnAfterCreateWnd();

protected:
	CBCGPCircularProgressIndicatorImpl*	m_pCircularProgressIndicator;
};

class BCGCBPRODLLEXPORT CBCGPCircularProgressIndicatorRenderer
{
	friend class CBCGPCircularProgressIndicatorImpl;

public:
	CBCGPCircularProgressIndicatorRenderer(CWnd* pWndOwner)
	{
		ASSERT_VALID(pWndOwner);

		m_pCircularProgressWndOwner = pWndOwner;
		m_pCircularProgressGM = NULL;
		m_CircularProgress.m_pRenderer = this;

		m_rectCircularProgress.SetRectEmpty();

		m_CircularProgress.m_ProgressColors.m_brFrameOutline = m_CircularProgress.m_ProgressColors.m_brFill = CBCGPBrush();
		m_CircularProgress.SetColors(m_CircularProgress.m_ProgressColors);

		m_CircularProgress.m_Options.m_strPercentageFormat.Empty();
	}

	virtual ~CBCGPCircularProgressIndicatorRenderer()
	{
		if (m_pCircularProgressGM != NULL)
		{
			delete m_pCircularProgressGM;
		}
	}

	void DoDrawProgress(CDC* pDC, const CRect& rectProgress)
	{
		if (!EnsureCircularProgressGraphicsManager())
		{
			return;
		}

		m_rectCircularProgress = rectProgress;
	
		CRect rectClient;
		m_pCircularProgressWndOwner->GetClientRect(rectClient);

		m_pCircularProgressGM->BindDC(pDC, rectClient);
		m_pCircularProgressGM->BeginDraw();
			
		m_CircularProgress.SetRect(rectProgress);
		m_CircularProgress.OnDraw(m_pCircularProgressGM, CBCGPRect());
			
		m_pCircularProgressGM->EndDraw();
	}

	virtual BOOL EnsureCircularProgressGraphicsManager()
	{
		if (m_pCircularProgressGM == NULL)
		{
			m_pCircularProgressGM = CBCGPGraphicsManager::CreateInstance();
		}

		return m_pCircularProgressGM != NULL;
	}

protected:
	CWnd*								m_pCircularProgressWndOwner;
	CBCGPCircularProgressIndicatorImpl	m_CircularProgress;
	CBCGPGraphicsManager*				m_pCircularProgressGM;
	CRect								m_rectCircularProgress;
};

#endif // !defined(AFX_BCGPCIRCULARPROGRESSINDICATOR_H__CA3BCAE6_CA45_4214_A278_06D3F1DFD185__INCLUDED_)
