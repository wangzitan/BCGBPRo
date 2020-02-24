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
// BCGPSVGImage.h: interface for the CBCGPSVGImage class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BCGPSVGIMAGE_H__84AC78B3_2F82_4D27_B5EF_EB5E49165E54__INCLUDED_)
#define AFX_BCGPSVGIMAGE_H__84AC78B3_2F82_4D27_B5EF_EB5E49165E54__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "bcgcbpro.h"
#include "BCGPGraphicsManager.h"
#include "BCGPXml.h"

#ifndef BCGP_EXCLUDE_SVG

#if (!defined _BCGSUITE_) && (!defined _BCGSUITE_INC_)
#include "BCGPToolBarImages.h"
#endif

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGBrush

class CBCGPSVGBrush : public CBCGPBrush
{
public:
	CBCGPSVGBrush()
	{
		m_bIsNone = FALSE;
		m_bIsCurrent = FALSE;
		m_dblOpacitySaved = -1.0;
	}

	const BOOL IsEmpty() const				{ return CBCGPBrush::IsEmpty() && m_strPath.IsEmpty(); }

	void SetPath(const CString& strPath)	{ m_strPath = strPath; }
	const CString& GetPath() const			{ return m_strPath; }

	void SetNone(BOOL bSet = TRUE)			{ m_bIsNone = bSet; }
	BOOL IsNone() const						{ return m_bIsNone; }

	void SetCurrent(BOOL bSet = TRUE)		{ m_bIsCurrent = bSet; }
	BOOL IsCurrent() const					{ return m_bIsCurrent; }

	void CopyFrom(const CBCGPSVGBrush& src)
	{
		m_strPath = src.m_strPath;
		m_bIsNone = src.m_bIsNone;
		m_bIsCurrent = src.m_bIsCurrent;

		CBCGPBrush::CopyFrom(src);

		Destroy();
	}

	void SetOpacity(double opacity)
	{
		if (m_dblOpacitySaved != m_opacity)
		{
			m_dblOpacitySaved = m_opacity;
			CBCGPBrush::SetOpacity(opacity);
		}
	}

	void RestoreOpacity()
	{
		if (m_dblOpacitySaved >= 0.0)
		{
			CBCGPBrush::SetOpacity(m_dblOpacitySaved);
			m_dblOpacitySaved = -1.0;
		}
	}

protected:
	CString	m_strPath;
	BOOL	m_bIsNone;
	BOOL	m_bIsCurrent;
	double	m_dblOpacitySaved;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGStyle

struct CBCGPSVGStyle
{
	CString	m_strFill;
	CString	m_strFillRule;
	CString	m_strOpacity;
	CString	m_strFillOpacity;
	CString	m_strStroke;
	CString	m_strStrokeWidth;
	CString	m_strStrokeDashes;
	CString	m_strStrokeDashOffset;
	CString	m_strStrokeOpacity;
	CString	m_strStrokeLineCap;
	CString	m_strStrokeLineJoin;
	CString	m_strClipPath;
	CString	m_strMask;
	CString	m_strFontFamily;
	CString	m_strFontSize;
	CString	m_strFontWeight;
	CString	m_strFontStyle;
	CString	m_strFontDecoration;
	CString	m_strTextAnchor;
	CString	m_strVisibility;
	CString m_strMarkerStart;
	CString m_strMarkerEnd;

	void Read(const CString& str);
	void Add(const CBCGPSVGStyle& style);
};

class CBCGPSVG;

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGBase

class CBCGPSVGBase : public CObject
{
	DECLARE_DYNAMIC(CBCGPSVGBase)

	friend class CBCGPSVGGroup;
	friend class CBCGPSVGClipPath;
	friend class CBCGPSVGElemTransform;
	friend class CBCGPSVGAttributesStorage;
	friend class CBCGPSVG;
	friend class CBCGPSVGViewBox;
	friend class CBCGPSVGGradient;
	
protected:
// Construction:
	CBCGPSVGBase(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);
	CBCGPSVGBase();
	void CommonInit();

	virtual ~CBCGPSVGBase();

// Attributes:
public:
	CBCGPSVG& GetRoot();

	double GetStrokeOpacity();
	const CBCGPSVGBrush& GetStrokeBrush()		{ return GetBrush(FALSE); }
	double GetStrokeWidth(BOOL bAllowDefault = TRUE);
	CBCGPStrokeStyle* GetStrokeStyle();
	double GetFillOpacity();
	const CBCGPSVGBrush& GetFillBrush()			{ return GetBrush(TRUE); }
	CString	GetFillRule();
	const CBCGPTextFormat& GetTextFormat();

	virtual CStringArray* GetErrors()			{ return (m_pParent == NULL) ? NULL : m_pParent->GetErrors(); }
	virtual BOOL IsAlphaDrawingMode() const		{ return (m_pParent == NULL) ? FALSE : m_pParent->IsAlphaDrawingMode(); }
	virtual BOOL IsPreserveAspectRatio() const	{ return FALSE; }

	BOOL IsVisible() const						{ return m_bIsVisible; }

	const CBCGPPoint& GetLocation() const		{ return m_ptLocation; }

	CBCGPSVGBase* CreateCopy(CBCGPSVGBase* pParent);

// Overrides:
public:
	void ReadAttributes(CBCGPXmlElement& elem, BOOL bAddToExisting = FALSE);
	virtual BOOL OnDraw(CBCGPGraphicsManager* pGM) = 0;
	virtual void OnApplyStyle()
	{
		ApplyStyle();
	}

	virtual void AdaptColors(COLORREF clrBase, COLORREF clrTone, BOOL bClampHue = FALSE);
	virtual void InvertColors();
	virtual void ConvertToGrayScale(double dblLumRatio = 1.0);
	virtual void MakeLighter(double dblRatio = .25);
	virtual void MakeDarker(double dblRatio = .25);
	virtual void Simplify();

	virtual CBCGPRect GetBounds() = 0;
	
	virtual void OnBeforeDraw(CBCGPGraphicsManager* pGM)
	{
		UNREFERENCED_PARAMETER(pGM);
	}
	
	virtual CBCGPSVGBase* FindByID(const CString& strID, BOOL bDefsOnly = FALSE)
	{
		UNREFERENCED_PARAMETER(bDefsOnly);
		return m_strID == strID && !strID.IsEmpty() ? this : NULL;
	}

	virtual BOOL FindStyle(const CString& strClassName, CBCGPSVGStyle& style)
	{
		UNREFERENCED_PARAMETER(strClassName);
		UNREFERENCED_PARAMETER(style);

		return FALSE;
	}

	virtual CBCGPRect GetViewBox() const
	{
		return CBCGPRect();
	}

protected:
	virtual void CopyFrom(const CBCGPSVGBase& src);
	virtual BOOL TranslateCoordinate(double& dblVal, BOOL bIsHorz, BOOL bIsSize);
	virtual CBCGPGeometry* CreateGeometry() { return NULL; }
	virtual CBCGPGeometry* GetGeometry() { return NULL; }

	const CBCGPSVGBrush& GetBrush(BOOL bIsFill, BOOL bAllowDefault = TRUE);
	
	BOOL HasTransform() const;
	BOOL SetTransfrom(CBCGPGraphicsManager* pGM);
	void CleanUpTransforms();

	BOOL DrawMarkers(CBCGPGraphicsManager* pGM);

	BOOL ReadStyleAttributes(CBCGPXmlElement& elem, const CString& strAttribute = _T("style"));
	BOOL ReadStyleAttributes(const CString& str);

	void ApplyStyle();
	void ApplyStyle(const CBCGPSVGStyle style);

	BOOL ReadStrokeStyleAttributes(CBCGPXmlElement& elem);
	BOOL ReadTransformAttributes(CBCGPXmlElement& elem, const CString& strAttribute = _T("transform"), BOOL bAddToExisting = FALSE);
	
	double ReadCoordinate(CBCGPXmlElement& elem, const CString& strAttribute, BOOL bIsHorz, BOOL bIsSize, double dblDefaultValue = 0.0);
	
	static double ReadNumericAttribute(CBCGPXmlElement& elem, const CString& strAttribute, double dblDefaultValue = 0.0, BOOL* pbIsAbsolute = NULL, BOOL bCheckFor01 = TRUE);

	void ReadClipPath(const CString& strClipPath);
	void ReadMask(const CString& strMask);
	BOOL ReadTextFormatAttributes(CBCGPXmlElement& elem);
	void TextAnchorToAlignment(const CString& strTextAnchor);
	void SetTextFormatWeight(const CString& strWeight);
	void SetTextFormatStyle(const CString& strStyle);
	void SetTextDecoration(const CString& strTextDecoration);
	void VeryfyTextFormat();

	CBCGPStrokeStyle* PrepareStrokeStyle(double dblLineWidth, BOOL& bIsTemp);
	CBCGPGeometry::BCGP_FILL_MODE PrepareFillMode();
	
	void GetGeometryBounds(CBCGPRect& rectBounds);

	CBCGPSVGBase*			m_pParent;
	CString					m_strID;
	CString					m_strClass;
	CString					m_strDefaultClass;
	double					m_dblOpacity;
	double					m_dblStrokeOpacity;
	CBCGPSVGBrush			m_brStroke;
	CBCGPSVGBrush			m_brStrokeAlpha;
	double					m_dblStrokeWidth;
	CBCGPStrokeStyle		m_StrokeStyle;
	double					m_dblFillOpacity;
	CBCGPSVGBrush			m_brFill;
	CString					m_strFillRule;
	CBCGPSVGBrush			m_brFillAlpha;
	CBCGPSVGBrush			m_brCurrentColor;
	CBCGPPoint				m_ptLocation;
	CBCGPTextFormat			m_textFormat;
	BOOL					m_bHasTextAttributes;
	BOOL					m_bIsVisible;
	CBCGPTransformsArray	m_arTransforms;
	CString					m_strClipPathID;
	CString					m_strMaskID;
	
	CString					m_strMarkerStart;
	CBCGPPoint				m_ptMarkerStart;
	double					m_dblMarkerStartAngle;

	CString					m_strMarkerEnd;
	CBCGPPoint				m_ptMarkerEnd;
	double					m_dblMarkerEndAngle;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGRect

class CBCGPSVGRect : public CBCGPSVGBase
{
	DECLARE_DYNCREATE(CBCGPSVGRect)

public:
	CBCGPSVGRect(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);
	
protected:
	virtual BOOL OnDraw(CBCGPGraphicsManager* pGM);
	virtual CBCGPRect GetBounds();
	virtual CBCGPGeometry* CreateGeometry();
	virtual void CopyFrom(const CBCGPSVGBase& src);

	CBCGPSVGRect()
	{
	}

	CBCGPSize	m_size;
	CBCGPSize	m_sizeCornerRadius;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGEllipse

class CBCGPSVGEllipse : public CBCGPSVGBase
{
	DECLARE_DYNCREATE(CBCGPSVGEllipse)

public:
	CBCGPSVGEllipse(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);

protected:
	CBCGPSVGEllipse()
	{
		m_dblRadiusX = m_dblRadiusY = 0.0;
	}

	virtual BOOL OnDraw(CBCGPGraphicsManager* pGM);
	virtual CBCGPRect GetBounds();
	virtual CBCGPGeometry* CreateGeometry();
	virtual void CopyFrom(const CBCGPSVGBase& src);

	double	m_dblRadiusX;
	double	m_dblRadiusY;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGLine

class CBCGPSVGLine : public CBCGPSVGBase
{
	DECLARE_DYNCREATE(CBCGPSVGLine)
		
public:
	CBCGPSVGLine(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);
	
protected:
	CBCGPSVGLine()
	{
	}

	virtual BOOL OnDraw(CBCGPGraphicsManager* pGM);
	virtual void CopyFrom(const CBCGPSVGBase& src);
	virtual CBCGPRect GetBounds();

	CBCGPPoint m_ptTo;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGText

class CBCGPSVGText : public CBCGPSVGBase
{
	DECLARE_DYNCREATE(CBCGPSVGText)
		
public:
	CBCGPSVGText(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);
	virtual ~CBCGPSVGText();
	
protected:
	CBCGPSVGText()
	{
		CommonInit();
	}

	void CommonInit();
	void CleanUp();

	virtual BOOL OnDraw(CBCGPGraphicsManager* pGM);
	virtual void CopyFrom(const CBCGPSVGBase& src);
	virtual CBCGPRect GetBounds();
	
	CString									m_strText;
	CArray<CBCGPSVGText*, CBCGPSVGText*>	m_arSpans;
	BOOL									m_bIsAbsoluteX;
	BOOL									m_bIsAbsoluteY;
	double									m_dblDY;
	BOOL									m_bYIsEM;
	BOOL									m_bDYIsEM;
	CBCGPPoint								m_ptTotalOffset;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGPolygon

class CBCGPSVGPolygon : public CBCGPSVGBase
{
	DECLARE_DYNCREATE(CBCGPSVGPolygon)
		
public:
	CBCGPSVGPolygon(CBCGPSVGBase* pParent, CBCGPXmlElement& elem, BOOL bIsClosed);
	
protected:
	CBCGPSVGPolygon()
	{
	}

	virtual BOOL OnDraw(CBCGPGraphicsManager* pGM);
	virtual CBCGPRect GetBounds();
	virtual CBCGPGeometry* CreateGeometry();
	virtual CBCGPGeometry* GetGeometry();
	virtual void CopyFrom(const CBCGPSVGBase& src);

	CBCGPPolygonGeometry	m_Geometry;
	CBCGPRect				m_rectBounds;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGPath

class CBCGPSVGPath : public CBCGPSVGBase
{
	DECLARE_DYNCREATE(CBCGPSVGPath)
		
public:
	CBCGPSVGPath(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);
	virtual ~CBCGPSVGPath();
	
protected:
	CBCGPSVGPath()
	{
		m_pGeometry = NULL;
	}

	BOOL ReadPathAttribute(CBCGPXmlElement& elem, const CString& strAttribute, CArray<CBCGPComplexGeometry*, CBCGPComplexGeometry*>& arGeometry, CStringArray* pErrors);
	void CalcBounds(const CBCGPPoint& pt);
	void CleanUp();

	virtual BOOL OnDraw(CBCGPGraphicsManager* pGM);
	virtual CBCGPRect GetBounds();
	virtual CBCGPGeometry* CreateGeometry();
	virtual CBCGPGeometry* GetGeometry();
	virtual void CopyFrom(const CBCGPSVGBase& src);
	
	CArray<CBCGPComplexGeometry*, CBCGPComplexGeometry*>	m_arGeometry;
	CBCGPGeometry*											m_pGeometry;
	CBCGPRect												m_rectBounds;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGGroup

class CBCGPSVGGroup : public CBCGPSVGBase
{
	DECLARE_DYNCREATE(CBCGPSVGGroup)
		
public:
	virtual CBCGPSVGBase* FindByID(const CString& strID, BOOL bDefsOnly = FALSE);

protected:
	CBCGPSVGGroup(CBCGPSVGBase* pParent, CBCGPXmlElement& elem, BOOL bIsDefs);
	CBCGPSVGGroup();

	virtual ~CBCGPSVGGroup();

	BOOL AddChildren(CBCGPXmlElement& root);
	CBCGPSVGBase* CreateUse(CBCGPSVGBase* pParent, CBCGPXmlElement& elem, CString& strUsedPath);

	virtual BOOL OnDraw(CBCGPGraphicsManager* pGM);
	virtual void CopyFrom(const CBCGPSVGBase& src);
	virtual CBCGPRect GetBounds();

	virtual void AdaptColors(COLORREF clrBase, COLORREF clrTone, BOOL bClampHue = FALSE);
	virtual void InvertColors();
	virtual void ConvertToGrayScale(double dblLumRatio = 1.0);
	virtual void MakeLighter(double dblRatio = .25);
	virtual void MakeDarker(double dblRatio = .25);
	virtual void Simplify();
	
	virtual BOOL FindStyle(const CString& strClassName, CBCGPSVGStyle& style);
	virtual void OnBeforeDraw(CBCGPGraphicsManager* pGM);

	void ReadStyles(const CString& str);

	virtual void OnApplyStyle();

	void CleanUp();

	CArray<CBCGPSVGBase*, CBCGPSVGBase*>	m_arElements;
	BOOL									m_bIsDefs;
	CString									m_strTitle;

	CMap<CString, LPCTSTR, CBCGPSVGStyle, const CBCGPSVGStyle&>	m_mapStyles;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGClipPath

class CBCGPSVGClipPath : public CBCGPSVGGroup
{
	DECLARE_DYNCREATE(CBCGPSVGClipPath)
		
public:
	CBCGPSVGClipPath(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);
	virtual ~CBCGPSVGClipPath();
	
protected:
	CBCGPSVGClipPath()
	{
	}

	virtual BOOL OnDraw(CBCGPGraphicsManager* /*pGM*/) { return TRUE; }
	virtual CBCGPGeometry* CreateGeometry();
	virtual void CopyFrom(const CBCGPSVGBase& src);

	void CleanUp();

	CArray<CBCGPGeometry*, CBCGPGeometry*> m_arGeometry;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGGradient

class CBCGPSVGGradient : public CBCGPSVGBase
{
	DECLARE_DYNAMIC(CBCGPSVGGradient)
		
public:
	CBCGPSVGGradient(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);

	const CBCGPSVGBrush& GetBrush() const
	{
		return m_Brush;
	}
	
protected:
	CBCGPSVGGradient()
	{
		m_pBase = NULL;
		m_bUserSpaceOnUse = FALSE;
		m_extendMode = CBCGPBrush::BCGP_GRADIENT_EXTEND_MODE_CLAMP;
	}

	virtual BOOL OnDraw(CBCGPGraphicsManager* /*pGM*/) { return TRUE; }
	virtual CBCGPRect GetBounds() { return CBCGPRect(); }
	virtual void CopyFrom(const CBCGPSVGBase& src);

	virtual void AdaptColors(COLORREF clrBase, COLORREF clrTone, BOOL bClampHue = FALSE);
	virtual void InvertColors();
	virtual void ConvertToGrayScale(double dblLumRatio = 1.0);
	virtual void MakeLighter(double dblRatio = .25);
	virtual void MakeDarker(double dblRatio = .25);
	virtual void Simplify();
	
	int ReadGradientStops(CBCGPXmlNodeList& elements, CBCGPBrushGradientStops& stops);
	
	CBCGPSVGBrush							m_Brush;
	CBCGPBrushGradientStops					m_stops;
	CBCGPSVGGradient*						m_pBase;
	BOOL									m_bUserSpaceOnUse;
	CBCGPBrush::BCGP_GRADIENT_EXTEND_MODE	m_extendMode;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGLinearGradient

class CBCGPSVGLinearGradient : public CBCGPSVGGradient
{
	DECLARE_DYNCREATE(CBCGPSVGLinearGradient)
		
public:
	CBCGPSVGLinearGradient(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);

protected:
	CBCGPSVGLinearGradient()
	{
	}
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGRadialGradient

class CBCGPSVGRadialGradient : public CBCGPSVGGradient
{
	DECLARE_DYNCREATE(CBCGPSVGRadialGradient)
		
public:
	CBCGPSVGRadialGradient(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);

protected:
	CBCGPSVGRadialGradient()
	{
	}
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGMask

class CBCGPSVGMask : public CBCGPSVGGroup
{
	DECLARE_DYNCREATE(CBCGPSVGMask)
		
public:
	CBCGPSVGMask(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);
	CBCGPSVGMask();
	
	virtual BOOL OnDraw(CBCGPGraphicsManager* /*pGM*/) { return TRUE; }
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGViewBox

class CBCGPSVGViewBox : public CBCGPSVGGroup
{
	DECLARE_DYNCREATE(CBCGPSVGViewBox)
		
public:
	CBCGPSVGViewBox(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);
	CBCGPSVGViewBox();
	
	virtual BOOL OnDraw(CBCGPGraphicsManager* pGM);
	virtual void CopyFrom(const CBCGPSVGBase& src);
	virtual BOOL TranslateCoordinate(double& dblVal, BOOL bIsHorz, BOOL bIsSize);
	
	const CBCGPSize& GetSize() const
	{
		return m_size;
	}
	
	virtual CBCGPRect GetBounds();
	virtual BOOL IsPreserveAspectRatio() const
	{
		return m_bPreserveAspectRatio;
	}
	
	virtual double GetExportScaleRatio() const
	{
		return 1.0;
	}
	
	virtual CBCGPRect GetViewBox() const
	{
		return m_rectViewBox;
	}
	
	const CString& GetXML() const
	{
		return m_strXML;
	}

	void Mirror()
	{
		m_bIsMirror = !m_bIsMirror;
	}

	BOOL IsMirror() const
	{
		return m_bIsMirror;
	}
	
protected:
	BOOL CreateData(CBCGPXmlElement& root, BOOL bApplyStyle = FALSE);
	
	CBCGPRect	m_rectViewBox;
	BOOL		m_bPreserveAspectRatio;
	CBCGPSize	m_size;
	CBCGPSize	m_sizeScale;
	CString		m_strXML;
	BOOL		m_bIsMirror;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVG

class CBCGPSVG : public CBCGPSVGViewBox
{
	DECLARE_DYNCREATE(CBCGPSVG)
		
	friend class CBCGPSVGBase;
	friend class CBCGPSVGImage;
	friend class CBCGPSVGImageList;

public:
	CBCGPSVG(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);
	CBCGPSVG();
	
	virtual BOOL IsAlphaDrawingMode() const	{ return m_pParent == NULL ? m_bIsAlphaDrawingMode : m_pParent->IsAlphaDrawingMode(); }
	virtual CStringArray* GetErrors() { return m_pParent == NULL ? &m_arErrors : m_pParent->GetErrors(); }
	virtual void AdaptColors(COLORREF clrBase, COLORREF clrTone, BOOL bClampHue = FALSE);

	const CBCGPTextFormat& GetDefaultTextFormat();
	const CBCGPColor& GetDefaultColor() const { return m_clrDefault; }

	virtual double GetExportScaleRatio() const
	{
		return m_dblExportScale;
	}

protected:
	CBCGPTextFormat	m_tfDefault;
	double			m_dblExportScale;
	BOOL			m_bIsAlphaDrawingMode;
	CStringArray	m_arErrors;
	CBCGPColor		m_clrDefault;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGSymbol

class CBCGPSVGSymbol : public CBCGPSVGViewBox
{
	DECLARE_DYNCREATE(CBCGPSVGSymbol)
		
public:
	CBCGPSVGSymbol(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) :
		CBCGPSVGViewBox(pParent, elem)
	{
	}
	
protected:
	CBCGPSVGSymbol()
	{
	}
	
	virtual BOOL OnDraw(CBCGPGraphicsManager* /*pGM*/) {	return TRUE; }
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGMarker

class CBCGPSVGMarker : public CBCGPSVGViewBox
{
	friend class CBCGPSVGBase;

	DECLARE_DYNCREATE(CBCGPSVGMarker)
		
public:
	CBCGPSVGMarker(CBCGPSVGBase* pParent, CBCGPXmlElement& elem);
	
protected:
	CBCGPSVGMarker()
	{
	}

	void DoDraw(CBCGPGraphicsManager* pGM, const CBCGPPoint& ptMarker, double dblMarkerAngle);
	
	virtual void CopyFrom(const CBCGPSVGBase& src);
	
	CBCGPSize	m_size;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGImage

class BCGCBPRODLLEXPORT CBCGPSVGImage
{
	friend class CBCGPSVGBase;
	friend class CBCGPSVGImageList;
	
public:
	static CString m_strSVGResType;	// "SVG" by default

	CBCGPSVGImage();
	CBCGPSVGImage(const CBCGPSVGImage& src);

	virtual ~CBCGPSVGImage();

	BOOL IsEmpty() const
	{
		return m_SVG.m_arElements.GetSize() == 0;
	}
	
	BOOL LoadFromFile(const CString& strFileName);
    BOOL LoadFromBuffer(LPBYTE lpBuffer, UINT length);
	BOOL LoadFromBuffer(const CString& strSVGBuffer);
	
	BOOL Load(UINT uiResID, HINSTANCE hinstRes = NULL);
	BOOL LoadStr(LPCTSTR lpszResourceName, HINSTANCE hinstRes = NULL);
	
	void CleanUp();
	
	void DoDraw(CBCGPGraphicsManager* pGM, const CBCGPPoint& pt = CBCGPPoint(), const CBCGPSize& size = CBCGPSize());
	void DoDraw(CDC* pDC, const CRect& rect, BOOL bDisabled = FALSE, BYTE alphaSrc = 255, BOOL bCashBitmap = FALSE);
	
	HBITMAP ExportToBitmap(double dblScale = 1.0, COLORREF clrBackground = CLR_NONE);
	HBITMAP ExportToBitmap(const CSize& sizeDest, COLORREF clrBackground = CLR_NONE);
	HICON ExportToIcon(const CSize& sizeIcon);
	
	BOOL ExportToFile(const CString& strFilePath, double dblScale = 1.0);
	
	void TraceProblems();
	
	const CBCGPSize& GetSize() const
	{
		return m_SVG.m_size;
	}
	
	const CString& GetName() const
	{
		return m_SVG.m_strTitle;
	}
	
	void AdaptColors(COLORREF clrBase, COLORREF clrTone, BOOL bClampHue = FALSE);
	void InvertColors();
	void ConvertToGrayScale(double dblLumRatio = 1.0);
	void MakeLighter(double dblRatio = .25);
	void MakeDarker(double dblRatio = .25);
	void Simplify();
	void Mirror();

	const CString& GetXML() const
	{
		return m_SVG.GetXML();
	}
	
	void SetLocation(const CBCGPPoint& pt);
	const CBCGPPoint& GetLocation() const
	{
		return m_SVG.GetLocation();
	}

protected:
	HBITMAP ExportToBitmap(const CSize& size, const CBCGPPoint& ptOffset, double dblScale, COLORREF clrBackground = CLR_NONE);
	
protected:
	BOOL					m_bTraceProblems;
	CBCGPGraphicsManager*	m_pGM;
	CBCGPSVG				m_SVG;
	CBCGPToolBarImages		m_Image;
	BOOL					m_bIsSimplified;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGImageList

class BCGCBPRODLLEXPORT CBCGPSVGImageList : public CArray<CBCGPSVGImage*, CBCGPSVGImage*>
{
public:
	static CString m_strSVGListResType;	// "SVGLIST" by default
	static BOOL m_bDesignMode;
	
	CBCGPSVGImageList(UINT uiResID = 0, HINSTANCE hinstRes = NULL);
	virtual ~CBCGPSVGImageList();

	BOOL Load(UINT uiResID, HINSTANCE hinstRes = NULL);
	BOOL LoadStr(LPCTSTR lpszResourceName, HINSTANCE hinstRes = NULL);

	BOOL LoadFromFile(const CString& strFileName);
	BOOL SaveToFile(const CString& strFileName);	// Design mode only!

	BOOL LoadFromSVGSprite(const CBCGPSVGImage& sprite);
    BOOL LoadFromSVGSpriteBuffer(LPBYTE lpBuffer, UINT length);

	void RemoveAll();

	HBITMAP ExportToBitmap(double dblScale = 1.0, COLORREF clrBackground = CLR_NONE);

protected:
	BOOL LoadSVG(const CString& strResID, HINSTANCE hinstRes);
};

#ifndef AddaptColors
#define AddaptColors	AdaptColors
#endif


#endif // BCGP_EXCLUDE_SVG

#endif // !defined(AFX_BCGPSVGIMAGE_H__84AC78B3_2F82_4D27_B5EF_EB5E49165E54__INCLUDED_)
