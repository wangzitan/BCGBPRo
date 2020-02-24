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
// BCGPStaticGaugeImpl.h: interface for the CBCGPStaticGaugeImpl class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BCGPSTATICGAUGEIMPL_H__2D040C84_CD25_430F_8680_C4131A778363__INCLUDED_)
#define AFX_BCGPSTATICGAUGEIMPL_H__2D040C84_CD25_430F_8680_C4131A778363__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "BCGPGaugeImpl.h"

extern BCGCBPRODLLEXPORT UINT BCGM_ON_GAUGE_SCROLL_FINISHED;

// Content scrolling flags:
#define BCGPSTATICGAUGE_CONTENT_SCROLL_CYCLIC			0x0001
#define BCGPSTATICGAUGE_CONTENT_SCROLL_PAUSE_ON_MOUSE	0x0002

#define BCGPSTATICGAUGE_ANIMATION_NONE		0
#define BCGPSTATICGAUGE_ANIMATION_FLASH		1
#define BCGPSTATICGAUGE_ANIMATION_SCROLL	2

class BCGCBPRODLLEXPORT CBCGPStaticGaugeImpl : public CBCGPGaugeImpl  
{
	DECLARE_DYNAMIC(CBCGPStaticGaugeImpl)

public:
	CBCGPStaticGaugeImpl(CBCGPVisualContainer* pContainer = NULL);
	virtual ~CBCGPStaticGaugeImpl();

	virtual void CopyFrom(const CBCGPBaseVisualObject& src);

// Operations:
public:
	// Flashing operations:
	void StartFlashing(UINT nTime = 100 /* ms */);
	void StopFlashing();

	BOOL IsFlashing() const
	{
		return m_arData[0]->IsAnimated() && m_nAnimationType == BCGPSTATICGAUGE_ANIMATION_FLASH;
	}

	// Ticker operations:
	BOOL StartContentScrolling(double dblScrollTime = 2.0 /* sec */, double dblScrollDelay = 0.0 /* sec */,
		BOOL bIsHorizontal = TRUE,
		UINT nFlags = BCGPSTATICGAUGE_CONTENT_SCROLL_CYCLIC,
		CBCGPAnimationManager::BCGPAnimationType type = CBCGPAnimationManager::BCGPANIMATION_Linear, 
		CBCGPAnimationManagerOptions* pOptions = NULL);

	void StopContentScrolling(BOOL bResetOffset = FALSE);

	BOOL IsContentScrolling() const
	{
		return m_arData[0]->IsAnimated() && m_nAnimationType == BCGPSTATICGAUGE_ANIMATION_SCROLL;
	}

	virtual CWnd* SetOwner(CWnd* pWndOwner, BOOL bRedraw = FALSE);

	virtual void OnAnimation(CBCGPVisualDataObject* pDataObject);
	virtual void OnAnimationFinished(CBCGPVisualDataObject* pDataObject);

	void SetFillBrush(const CBCGPBrush& brush);
	const CBCGPBrush& GetFillBrush() const { return m_brFill; }

	void SetOutlineBrush(const CBCGPBrush& brush);
	const CBCGPBrush& GetOutlineBrush() const { return m_brOutline; }

	double GetOpacity() const { return m_dblOpacity; }
	virtual void SetOpacity(double opacity, BOOL bRedraw = TRUE);

	DWORD GetDefaultDrawFlags() const { return m_DefaultDrawFlags; }
	void SetDefaultDrawFlags(DWORD dwDrawFlags, BOOL bRedraw = TRUE);

// Overrides:
protected:
	virtual BOOL OnMouseDown(int nButton, const CBCGPPoint& pt);
	virtual void OnMouseUp(int nButton, const CBCGPPoint& pt);
	virtual void OnMouseMove(const CBCGPPoint& pt);
	virtual void OnMouseLeave();
	virtual void OnCancelMode();

	virtual BOOL OnSetMouseCursor(const CBCGPPoint& pt);

	virtual CBCGPSize GetContentTotalSize(CBCGPGraphicsManager* /*pGM*/) { return CBCGPSize(0.0, 0.0); }

// Attributes:
protected:
	int			m_nAnimationType;
	UINT		m_nFlashTime;
	BOOL		m_bOff;
	BOOL		m_bIsPressed;
	double		m_dblOpacity;
	
	CBCGPBrush	m_brFill;
	CBCGPBrush	m_brOutline;

	BOOL		m_bIsHorizontalScroll;
	UINT		m_nScrollFlags;
	double		m_dblScrollOffset;
	double		m_dblScrollTime;
	double		m_dblScrollDelay;
	BOOL		m_bScrollInternal;
	BOOL		m_bScrollStopped;
	BOOL		m_bScrollPaused;

	CBCGPAnimationManager::BCGPAnimationType	m_ScrollAnimationType;
	CBCGPAnimationManagerOptions				m_ScrollAnimationOptions;

	DWORD		m_DefaultDrawFlags;
};

class BCGCBPRODLLEXPORT CBCGPStaticFrameImpl : public CBCGPStaticGaugeImpl  
{
	DECLARE_DYNCREATE(CBCGPStaticFrameImpl)
		
public:
	CBCGPStaticFrameImpl(CBCGPVisualContainer* pContainer = NULL);
	virtual ~CBCGPStaticFrameImpl();
	
// Overrides:
public:
	virtual void OnDraw(CBCGPGraphicsManager* pGM, const CBCGPRect& rectClip, DWORD dwFlags = BCGP_DRAW_STATIC | BCGP_DRAW_DYNAMIC);
	virtual CBCGPSize GetDefaultSize(CBCGPGraphicsManager* /*pGM*/, const CBCGPBaseVisualObject* /*pParentGauge*/) { return CBCGPSize(100.0, 100.0); }
	virtual void CopyFrom(const CBCGPBaseVisualObject& src);

// Attributes:
public:
	double GetFrameSize() const { return m_dblFrameSize; }
	void SetFrameSize(double dblFrameSize, BOOL bRedraw = TRUE);

	const CBCGPStrokeStyle& GetStrokeStyle() const { return m_strokeStyle; }
	void SetStrokeStyle(const CBCGPStrokeStyle& strokeStyle, BOOL bRedraw = TRUE);

	double GetCornerRadius() const { return m_dblCornerRadius; }
	void SetCornerRadius(double dblCornerRadius, BOOL bRedraw = TRUE);

protected:
	CBCGPStrokeStyle	m_strokeStyle;
	double				m_dblFrameSize;
	double				m_dblCornerRadius;
};

#endif // !defined(AFX_BCGPSTATICGAUGEIMPL_H__2D040C84_CD25_430F_8680_C4131A778363__INCLUDED_)
