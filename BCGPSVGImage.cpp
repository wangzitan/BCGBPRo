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
// BCGPSVGImage.cpp: implementation of the CBCGPSVGImage class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BCGPSVGImage.h"
#include "BCGPImageProcessing.h"
#include "BCGPDrawManager.h"
#include "BCGPGlobalUtils.h"
#if (!defined _BCGSUITE_)
#include "BCGPPngImage.h"
#endif

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#ifndef BCGP_EXCLUDE_SVG

IMPLEMENT_DYNAMIC(CBCGPSVGBase, CObject)
IMPLEMENT_DYNCREATE(CBCGPSVGRect, CBCGPSVGBase)
IMPLEMENT_DYNCREATE(CBCGPSVGEllipse, CBCGPSVGBase)
IMPLEMENT_DYNCREATE(CBCGPSVGLine, CBCGPSVGBase)
IMPLEMENT_DYNCREATE(CBCGPSVGText, CBCGPSVGBase)
IMPLEMENT_DYNCREATE(CBCGPSVGPolygon, CBCGPSVGBase)
IMPLEMENT_DYNCREATE(CBCGPSVGPath, CBCGPSVGBase)
IMPLEMENT_DYNCREATE(CBCGPSVGGroup, CBCGPSVGBase)
IMPLEMENT_DYNCREATE(CBCGPSVGClipPath, CBCGPSVGGroup)
IMPLEMENT_DYNAMIC(CBCGPSVGGradient, CBCGPSVGBase)
IMPLEMENT_DYNCREATE(CBCGPSVGLinearGradient, CBCGPSVGGradient)
IMPLEMENT_DYNCREATE(CBCGPSVGRadialGradient, CBCGPSVGGradient)
IMPLEMENT_DYNCREATE(CBCGPSVGMask, CBCGPSVGGroup)
IMPLEMENT_DYNCREATE(CBCGPSVGViewBox, CBCGPSVGGroup)
IMPLEMENT_DYNCREATE(CBCGPSVGSymbol, CBCGPSVGViewBox)
IMPLEMENT_DYNCREATE(CBCGPSVGMarker, CBCGPSVGViewBox)
IMPLEMENT_DYNCREATE(CBCGPSVG, CBCGPSVGViewBox)

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGAttributesStorage

class CBCGPSVGAttributesStorage
{
public:
	CBCGPSVGAttributesStorage(CBCGPSVGBase* pElem) :
		m_FillBrush(pElem->GetFillBrush()),
		m_StrokeBrush(pElem->GetStrokeBrush())
	{
		m_bIsTemp = FALSE;
		m_dblStrokeWidth = pElem->GetStrokeWidth();
		m_pStrokeStyle = pElem->PrepareStrokeStyle(m_dblStrokeWidth, m_bIsTemp);
	}
	
	~CBCGPSVGAttributesStorage()
	{
		if (m_bIsTemp && m_pStrokeStyle != NULL)
		{
			delete m_pStrokeStyle;
		}

		((CBCGPSVGBrush&)m_FillBrush).RestoreOpacity();
		((CBCGPSVGBrush&)m_StrokeBrush).RestoreOpacity();
	}
	
	CBCGPStrokeStyle* GetStrokeStyle() { return m_pStrokeStyle; }
	double GetStrokeWidth() { return m_dblStrokeWidth; }
	
	const CBCGPSVGBrush& GetFillBrush() const
	{
		return m_FillBrush;
	}

	const CBCGPSVGBrush& GetStrokeBrush() const
	{
		return m_StrokeBrush;
	}

protected:
	CBCGPStrokeStyle*		m_pStrokeStyle;
	double					m_dblStrokeWidth;
	BOOL					m_bIsTemp;
	const CBCGPSVGBrush&	m_FillBrush;
	const CBCGPSVGBrush&	m_StrokeBrush;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGElemTransform

class CBCGPSVGElemTransform
{
public:
	CBCGPSVGElemTransform(CBCGPGraphicsManager* pGM, CBCGPSVGBase* pElem)
	{
		m_pGM = pGM;
		m_bHasTransform = pElem->SetTransfrom(pGM);
	}
	
	~CBCGPSVGElemTransform()
	{
		if (m_bHasTransform)
		{
			m_pGM->PopTransform();
		}
	}
	
protected:
	CBCGPGraphicsManager*	m_pGM;
	BOOL					m_bHasTransform;
};

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGEViewBoxTransform

class CBCGPSVGEViewBoxTransform
{
public:
	CBCGPSVGEViewBoxTransform(CBCGPGraphicsManager* pGM, const CBCGPRect& rectViewBox, const CBCGPSize& szScale, const CBCGPPoint& ptOffset, BOOL bIsMirror)
	{
		m_pGM = pGM;

		if (bIsMirror && rectViewBox.Width() <= 0.0)
		{
			bIsMirror = FALSE;
		}

		if (rectViewBox.TopLeft() != CBCGPPoint(0., 0.) || szScale != CBCGPSize(1., 1.) || ptOffset != CBCGPPoint() || bIsMirror)
		{
			pGM->PushTransform();

			if (ptOffset != CBCGPPoint())
			{
				pGM->SetTranslateTransform(CBCGPSize(ptOffset.x, ptOffset.y));
			}
		
			if (szScale != CBCGPSize(1., 1.))
			{
				pGM->SetScaleTransform(szScale);
			}
			
			if (rectViewBox.TopLeft() != CBCGPPoint(0., 0.))
			{
				pGM->SetTranslateTransform(CBCGPSize(-rectViewBox.left, -rectViewBox.top));
			}

			if (bIsMirror)
			{
				pGM->SetTranslateTransform(CBCGPSize(rectViewBox.CenterPoint().x * 2, 0.0));
				pGM->SetScaleTransform(CBCGPSize(-1., 1.));
			}

			m_bHasTransform = TRUE;
		}
		else
		{
			m_bHasTransform = FALSE;
		}
	}
	
	~CBCGPSVGEViewBoxTransform()
	{
		if (m_bHasTransform)
		{
			m_pGM->PopTransform();
		}
	}
	
protected:
	CBCGPGraphicsManager*	m_pGM;
	BOOL					m_bHasTransform;
};

static CMap<CString, LPCTSTR, int, int> g_mapSystemColors;

static CBCGPPoint ApplyBearing(const CBCGPPoint& ptPrev, CBCGPPoint& ptOrig, double dblBearing)
{
	if (dblBearing == 0.0)
	{
		return ptOrig;
	}
	
	double dblAngle = bcg_deg2rad(360 - dblBearing);
	double dblDist = bcg_distance(ptPrev, ptOrig);
	
	return CBCGPPoint(
		ptPrev.x + dblDist * cos(dblAngle),
		ptPrev.y - dblDist * sin(dblAngle));
}
//*******************************************************************************************************
static BOOL IsSpaceChar(TCHAR c)
{
	return (c == 0x20 || c == 0x9 || c == 0xD || c == 0xA);
}
//*******************************************************************************************************
static int FindStyleAttribute(const CString& str, const CString& strName)
{
	const int nLen = str.GetLength();
	const int nNameLen = strName.GetLength();

	for (int nStart = 0; nStart < nLen;)
	{
		int nIndex = str.Find(strName, nStart);
		if (nIndex < 0)
		{
			return -1;
		}

		nStart = nIndex + nNameLen;
		if (nStart >= nLen)
		{
			return -1;
		}
		
		if (nIndex > 0)
		{
			TCHAR tPrev = str[nIndex - 1];
			if (!IsSpaceChar(tPrev) && tPrev != _T(';'))
			{
				continue;
			}
		}

		for (int i = nStart; i < nLen; i++)
		{
			if (str[i] == _T(':'))
			{
				return i + 1;
			}

			if (!IsSpaceChar(str[i]))
			{
				break;
			}
		}
	}

	return -1;
}
//*******************************************************************************************************
static void BCGPStrCat(TCHAR* szCurr, TCHAR c)
{
	const int nLen = (int)_tcsclen(szCurr);
	szCurr[nLen] = c;
	szCurr[nLen + 1] = _T('\0');
}
//*******************************************************************************************************
static int PreparePathData(const CString& str, CStringArray& arData, const CString& strSkip = _T(","))
{
	const int nSize = (int)str.GetLength();

	arData.RemoveAll();
	arData.SetSize(0, nSize / 4);

	TCHAR* szCurr = new TCHAR[nSize + 1];
	szCurr[0] = _T('\0');

	BOOL bPrevIsE = FALSE;
	
	for (int i = 0; i < nSize; i++)
	{
		TCHAR c = str[i];
		
		if (IsSpaceChar(c) || strSkip.Find(c) >= 0)
		{
			if (_tcsclen(szCurr) > 0)
			{
				arData.Add(szCurr);
				szCurr[0] = _T('\0');
			}
		}
		else if (_istalpha(c))
		{
			if (c == _T('e') || c == _T('E'))
			{
				BCGPStrCat(szCurr, c);
				bPrevIsE = TRUE;
				continue;
			}
			else
			{
				if (_tcsclen(szCurr) > 0)
				{
					arData.Add(szCurr);
					szCurr[0] = _T('\0');
				}
				
				BCGPStrCat(szCurr, c);

				arData.Add(szCurr);
				szCurr[0] = _T('\0');
			}
		}
		else if (c == _T('-') && _tcsclen(szCurr) > 0 && !bPrevIsE)
		{
			arData.Add(szCurr);
			szCurr[0] = _T('\0');

			BCGPStrCat(szCurr, c);
		}
		else if (c == _T('.') && _tcschr(szCurr, c) != NULL)
		{
			arData.Add(szCurr);
			szCurr[0] = _T('\0');

			BCGPStrCat(szCurr, c);
		}
		else
		{
			BCGPStrCat(szCurr, c);
		}

		bPrevIsE = FALSE;
	}
	
	if (_tcsclen(szCurr) > 0)
	{
		arData.Add(szCurr);
	}
	
	delete [] szCurr;
	return (int)arData.GetSize();
}
//*******************************************************************************************************
static int PrepareTransformData(const CString& str, CStringArray& arData, const CString& strSkip = _T(",()"))
{
	arData.RemoveAll();
	
	CString strCurr;
	
	for (int i = 0; i < str.GetLength(); i++)
	{
		TCHAR c = str[i];
		
		if (IsSpaceChar(c) || strSkip.Find(c) >= 0)
		{
			if (!strCurr.IsEmpty())
			{
				arData.Add(strCurr);
				strCurr.Empty();
			}
		}
		else
		{
			strCurr += c;
		}
	}
	
	if (!strCurr.IsEmpty())
	{
		arData.Add(strCurr);
	}
	
	return (int)arData.GetSize();
}
//*******************************************************************************************************
static CString GetURLPath(const CString& str)
{
	CString strPref = _T("url(#");
	TCHAR cEnd = _T(')');

	int nStart = str.Find(strPref);
	if (nStart < 0)
	{
		strPref = _T("url(\'#");
		cEnd = _T('\'');
		nStart = str.Find(strPref);
	}

	CString strPath;

	if (nStart >= 0)
	{
		for (int i = nStart + strPref.GetLength(); i < str.GetLength(); i++)
		{
			TCHAR c = str[i];
			if (c == cEnd)
			{
				break;
			}
			
			strPath += c;
		}
	}

	return strPath;
}
//*******************************************************************************************************
static double StringToNumber(const CString& strIn, double dblDefault = 0.0)
{
	CString str = strIn;
	
	str.TrimLeft();
	str.TrimRight();

	double dblValue = dblDefault;
	int nLen = str.GetLength();
	
	if (nLen == 0)
	{
		return dblValue;
	}

	float fVal = 0;
	float fValDec = 0;
	float fValExp = 0;
	int nDecCount = 0;

	BOOL bAfterDot = FALSE;
	BOOL bExp = FALSE;
	BOOL bIsNegative = FALSE;
	BOOL bIsNegativeExp = FALSE;

	for (int i = 0; i < nLen; i++)
	{
		TCHAR c = str[i];

		if (c == _T('-'))
		{
			if (fVal == 0.0)
			{
				bIsNegative = TRUE;
			}
			else if (bExp)
			{
				bIsNegativeExp = TRUE;
			}
			
			continue;
		}

		if (c == _T('+'))
		{
			continue;
		}

		if (c == _T('.') || c == _T(','))
		{
			if (bAfterDot)
			{
				break;
			}

			bAfterDot = TRUE;
			continue;
		}

		if (c == _T('e') || c == _T('E'))
		{
			bExp = TRUE;
			continue;
		}

		if (c < _T('0') || c > _T('9'))
		{
			break;
		}

		const int nDigit = c - _T('0');

		if (bExp)
		{
			fValExp = fValExp * 10.0f + nDigit;
		}
		else if (bAfterDot)
		{
			fValDec = fValDec * 10.0f + nDigit;
			nDecCount++;
		}
		else
		{
			fVal = fVal * 10.0f + nDigit;
		}
	}

	if (nDecCount > 0)
	{
		fVal += fValDec / (float)pow(10.0f, nDecCount);
	}

	if (fValExp != 0.0)
	{
		if (bIsNegativeExp)
		{
			fVal /= (float)pow(10.0f, fValExp);
		}
		else
		{
			fVal *= (float)pow(10.0f, fValExp);
		}
	}

	dblValue = (double)(bIsNegative ? -fVal : fVal);

	const double DPI = 96.0;

	if (str.GetLength() > 2)
	{
		CString strR = str.Right(2);
		strR.MakeLower();

		if (strR == _T("cm"))
		{
			dblValue *= DPI / 2.54;
		}
		else if (strR == _T("mm"))
		{
			dblValue *= DPI / 25.4;
		}
		else if (strR == _T("em"))
		{
			dblValue *= 16;	// TODO
		}
		else if (strR == _T("ex"))
		{
			dblValue *= 7;	// TODO
		}
		else if (strR == _T("in"))
		{
			dblValue *= DPI;
		}
		else if (strR == _T("pc"))
		{
			dblValue *= DPI / 6;
		}
		else if (strR == _T("pt"))
		{
			dblValue *= DPI / 72;
		}
	}

	return dblValue;
}
//*******************************************************************************************************
static int StringToRGBValue(const CString& str)
{
	int nVal = _ttol(str);
	if (str.Find(_T("%")) >= 0)
	{
		nVal = bcg_clamp((int)(2.55 * nVal), 0, 255);
	}

	return nVal;
}
//*******************************************************************************************************
static CBCGPColor StringToColor(const CString& strIn)
{
	CString str = strIn;
	
	str.TrimLeft();
	str.TrimRight();

	str.MakeLower();
	
	if (str.IsEmpty() || str == _T("none"))
	{
		return CBCGPColor();
	}
	
	double a = 1.;

	if (str[0] == _T('#'))
	{
		str = str.Mid(1);
		
		if (str.GetLength() == 3)
		{
			str.Insert(2, str[2]);
			str.Insert(1, str[1]);
			str.Insert(0, str[0]);
		}
		
		if (str.GetLength() != 6)
		{
			TRACE(_T("CBCGPSVGImage: incorrect color format %s\n"), (LPCTSTR)str);
			return FALSE;
		}
		
		int nR = 0;
		int nG = 0;
		int nB = 0;

#if _MSC_VER < 1400
		_stscanf(str, _T("%2x%2x%2x"), &nR, &nG, &nB);
#else
		_stscanf_s(str, _T("%2x%2x%2x"), &nR, &nG, &nB);
#endif
		return CBCGPColor(RGB(nR, nG, nB));
	}

	CString strType = str.GetLength() > 3 ? str.Left(3) : CString();
	if (str.GetLength() > 4)
	{
		CString str4 = str.Left(4);
		if (str4 == _T("rgba") || str4 == _T("hsla"))
		{
			strType = str4;
		}
	}

	if (strType == _T("rgb") || strType == _T("rgba"))
	{
		CStringArray arRGB;
		int nCount = strType.GetLength();

		if (PreparePathData(str.Mid(nCount), arRGB, _T(",()")) == nCount)
		{
			int nR = StringToRGBValue(arRGB[0]);
			int nG = StringToRGBValue(arRGB[1]);
			int nB = StringToRGBValue(arRGB[2]);

			if (nCount == 4)
			{
				a = StringToNumber(arRGB[3], a);
			}

			return CBCGPColor(RGB(nR, nG, nB), a);
		}
	}
	else if (strType == _T("hsl") || strType == _T("hsla"))
	{
		CStringArray arHSL;
		int nCount = strType.GetLength();
		
		if (PreparePathData(str.Mid(nCount), arHSL, _T(",()")) == nCount)
		{
			double H = StringToNumber(arHSL[0]);
			if (H < 1.0)
			{
				H *= 100;
			}

			double S = StringToNumber(arHSL[1]);
			if (S > 1.0)
			{
				S *= 0.01;
			}

			double L = StringToNumber(arHSL[2]);
			if (L > 1.0)
			{
				L *= 0.01;
			}

			if (nCount == 4)
			{
				a = StringToNumber(arHSL[3], a);
			}

			return CBCGPColor(CBCGPDrawManager::HLStoRGB_TWO(H, L, S), a);
		}
	}
	
	str.Replace(_T("grey"), _T("gray"));
	
	const CArray<COLORREF, COLORREF>& arColors = CBCGPColor::GetRGBArray();
	for (int i = 0; i < (int)arColors.GetSize(); i++)
	{
		CBCGPColor color(arColors[i]);
		CString strName = color.ToString();
		strName.Remove(_T(' '));
		
		if (str.CompareNoCase(strName) == 0)
		{
			return color;
		}
	}
	
	// Try Windows system colors:
	if (g_mapSystemColors.IsEmpty())
	{
		g_mapSystemColors.SetAt(_T("ACTIVEBORDER"), COLOR_ACTIVEBORDER);
		g_mapSystemColors.SetAt(_T("ACTIVECAPTION"), COLOR_ACTIVECAPTION);
		g_mapSystemColors.SetAt(_T("APPWORKSPACE"), COLOR_APPWORKSPACE);
		g_mapSystemColors.SetAt(_T("BACKGROUND"), COLOR_BACKGROUND);
		g_mapSystemColors.SetAt(_T("BUTTONFACE"), COLOR_BTNFACE);
		g_mapSystemColors.SetAt(_T("BUTTONHIGHLIGHT"), COLOR_BTNHIGHLIGHT);
		g_mapSystemColors.SetAt(_T("BUTTONSHADOW"), COLOR_BTNSHADOW);
		g_mapSystemColors.SetAt(_T("BUTTONTEXT"), COLOR_BTNTEXT);
		g_mapSystemColors.SetAt(_T("CAPTIONTEXT"), COLOR_CAPTIONTEXT);
		g_mapSystemColors.SetAt(_T("GRAYTEXT"), COLOR_GRAYTEXT);
		g_mapSystemColors.SetAt(_T("HIGHLIGHT"), COLOR_HIGHLIGHT);
		g_mapSystemColors.SetAt(_T("HIGHLIGHTTEXT"), COLOR_HIGHLIGHTTEXT);
		g_mapSystemColors.SetAt(_T("INACTIVEBORDER"), COLOR_INACTIVEBORDER);
		g_mapSystemColors.SetAt(_T("INACTIVECAPTION"), COLOR_INACTIVECAPTION);
		g_mapSystemColors.SetAt(_T("INACTIVECAPTIONTEXT"), COLOR_INACTIVECAPTIONTEXT);
		g_mapSystemColors.SetAt(_T("INFOBACKGROUND"), COLOR_INFOBK);
		g_mapSystemColors.SetAt(_T("INFOTEXT"), COLOR_INFOTEXT);
		g_mapSystemColors.SetAt(_T("MENU"), COLOR_MENU);
		g_mapSystemColors.SetAt(_T("MENUTEXT"), COLOR_MENUTEXT);
		g_mapSystemColors.SetAt(_T("SCROLLBAR"), COLOR_SCROLLBAR);
		g_mapSystemColors.SetAt(_T("THREEDDARKSHADOW"), COLOR_3DDKSHADOW);
		g_mapSystemColors.SetAt(_T("THREEDFACE"), COLOR_3DFACE);
		g_mapSystemColors.SetAt(_T("THREEDHIGHLIGHT"), COLOR_3DHIGHLIGHT);
		g_mapSystemColors.SetAt(_T("THREEDLIGHTSHADOW"), COLOR_3DLIGHT);
		g_mapSystemColors.SetAt(_T("THREEDSHADOW"), COLOR_3DSHADOW);
		g_mapSystemColors.SetAt(_T("WINDOW"), COLOR_WINDOW);
		g_mapSystemColors.SetAt(_T("WINDOWTEXT"), COLOR_WINDOWTEXT);
	}

	int nSysColor = 0;
	str.MakeUpper();

	if (g_mapSystemColors.Lookup(str, nSysColor))
	{
		return CBCGPColor(::GetSysColor(nSysColor));
	}

	return CBCGPColor();
}
//*******************************************************************************************************
static BOOL StringToBrush(const CString& strIn, CBCGPSVGBrush& br)
{
	br.Destroy();

	CString str = strIn;

	str.TrimLeft();
	str.TrimRight();
	
	if (str.IsEmpty())
	{
		return TRUE;
	}

	if (str == _T("none"))
	{
		br.SetNone();
		return TRUE;
	}
	
	if (str == _T("currentColor"))
	{
		br.SetCurrent();
		return TRUE;
	}

	if (str == _T("transparent"))
	{
		br.SetColor(CBCGPColor::Black, 0.0);
		return TRUE;
	}

	if (str == _T("inherit"))
	{
		return TRUE;
	}

	br.SetPath(GetURLPath(str));
	if (!br.GetPath().IsEmpty())
	{
		return TRUE;
	}

	CBCGPColor color = StringToColor(str);
	if (!color.IsNull())
	{
		br.SetColor(color);
	}
	
	return TRUE;	
}
//*******************************************************************************************************
static BOOL ExtractStyle(const CString& str, const CString& strName, CString& strStyle)
{
	strStyle.Empty();

	int nIndex = FindStyleAttribute(str, strName);
	if (nIndex < 0)
	{
		return FALSE;
	}

	for (int i = nIndex; i < str.GetLength(); i++)
	{
		if (str[i] == _T(';'))
		{
			break;
		}
		
		strStyle += str[i];
	}
	
	strStyle.TrimLeft();
	strStyle.TrimRight();

	return !strStyle.IsEmpty();
}
//*******************************************************************************************************
static BOOL SetStokeCapStyle(const CString& strStyle, CBCGPStrokeStyle& style)
{
	if (!strStyle.IsEmpty())
	{
		CBCGPStrokeStyle::BCGP_CAP_STYLE capStyle = CBCGPStrokeStyle::BCGP_CAP_STYLE_FLAT;

		if (strStyle == _T("round"))
		{
			capStyle = CBCGPStrokeStyle::BCGP_CAP_STYLE_ROUND;
		}
		else if (strStyle == _T("square"))
		{
			capStyle = CBCGPStrokeStyle::BCGP_CAP_STYLE_SQUARE;
		}

		style.SetStartCap(capStyle);
		style.SetEndCap(capStyle);
		style.SetDashCap(capStyle);

		return !style.IsEmpty();
	}

	return FALSE;
}
//*******************************************************************************************************
static BOOL SetStokeJoinStyle(const CString& strStyle, CBCGPStrokeStyle& style)
{
	if (!strStyle.IsEmpty())
	{
		CBCGPStrokeStyle::BCGP_LINE_JOIN joinStyle = CBCGPStrokeStyle::BCGP_LINE_JOIN_MITER;
		
		if (strStyle == _T("round"))
		{
			joinStyle = CBCGPStrokeStyle::BCGP_LINE_JOIN_ROUND;
		}
		else if (strStyle == _T("bevel"))
		{
			joinStyle = CBCGPStrokeStyle::BCGP_LINE_JOIN_BEVEL;
		}
		else if (strStyle == _T("miter-clip"))
		{
			joinStyle = CBCGPStrokeStyle::BCGP_LINE_JOIN_MITER_OR_BEVEL;
		}
		
		style.SetLineJoin(joinStyle);
		return !style.IsEmpty();
	}
	
	return FALSE;
}
//*******************************************************************************************************
static BOOL ReadNumber(const CStringArray& arData, int& i, double& dblValue, BOOL bOneCharCmd = TRUE)
{
	if (i >= arData.GetSize())
	{
		return FALSE;
	}

	CString strNum = arData[i];

	if ((strNum.GetLength() == 1 || !bOneCharCmd) && _istalpha(strNum[0]))
	{
		return FALSE;
	}

	dblValue = StringToNumber(strNum);

	i++;
	return TRUE;
}
//*******************************************************************************************************
static BOOL ReadPoint(const CStringArray& arData, int& i, CBCGPPoint& pt, BOOL bIsRelative = FALSE)
{
	double x = 0;
	double y = 0;

	if (!ReadNumber(arData, i, x) || !ReadNumber(arData, i, y))
	{
		return FALSE;
	}

	if (bIsRelative)
	{
		pt.x += x;
		pt.y += y;
	}
	else
	{
		pt.x = x;
		pt.y = y;
	}

	return TRUE;
}
//*******************************************************************************************************
static BOOL ReadSize(const CStringArray& arData, int& i, CBCGPSize& size)
{
	double cx = 0;
	double cy = 0;
	
	if (!ReadNumber(arData, i, cx) || !ReadNumber(arData, i, cy))
	{
		return FALSE;
	}
	
	size.cx = cx;
	size.cy = cy;
	
	return TRUE;
}
//*******************************************************************************************************
static BOOL CreateBrushFromAttribute(CBCGPXmlElement& elem, const CString& strAttribute, CBCGPSVGBrush& br)
{
	return StringToBrush(elem.GetStringAttribute(strAttribute), br);
}
//*******************************************************************************************************
static BOOL StringToDashes(const CString& strIn, CBCGPStrokeStyle& style)
{
	CString str = strIn;

	str.TrimLeft();
	str.TrimRight();
	
	if (str.IsEmpty())
	{
		return TRUE;
	}

	CStringArray arData;
	PreparePathData(str, arData);
	
	CArray<FLOAT, FLOAT> arDashes;
	
	double dblDash = 0.0;
	int i = 0;
	
	while (ReadNumber(arData, i, dblDash))
	{
		arDashes.Add((FLOAT)dblDash);
	}
	
	style.SetDashes(arDashes);
	return TRUE;	
}
//*******************************************************************************************************
static BOOL StringToDashOffset(const CString& strIn, CBCGPStrokeStyle& style)
{
	CString str = strIn;

	str.TrimLeft();
	str.TrimRight();
	
	if (str.IsEmpty())
	{
		return TRUE;
	}

	style.SetDashOffset((FLOAT)StringToNumber(str));
	return TRUE;	
}

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGStyle

void CBCGPSVGStyle::Read(const CString& str)
{
	*this = CBCGPSVGStyle();

	if (str.IsEmpty())
	{
		return;
	}

	ExtractStyle(str, _T("fill"), m_strFill);
	ExtractStyle(str, _T("fill-rule"), m_strFillRule);
	ExtractStyle(str, _T("opacity"), m_strOpacity);
	ExtractStyle(str, _T("fill-opacity"), m_strFillOpacity);
	ExtractStyle(str, _T("stroke"), m_strStroke);
	ExtractStyle(str, _T("stroke-width"), m_strStrokeWidth);
	ExtractStyle(str, _T("stroke-dasharray"), m_strStrokeDashes);
	ExtractStyle(str, _T("stroke-dashoffset"), m_strStrokeDashOffset);
	ExtractStyle(str, _T("stroke-opacity"), m_strStrokeOpacity);
	ExtractStyle(str, _T("stroke-linecap"), m_strStrokeLineCap);
	ExtractStyle(str, _T("stroke-linejoin"), m_strStrokeLineJoin);
	ExtractStyle(str, _T("clip-path"), m_strClipPath);
	ExtractStyle(str, _T("mask"), m_strMask);
	ExtractStyle(str, _T("font-family"), m_strFontFamily);
	ExtractStyle(str, _T("font-size"), m_strFontSize);
	ExtractStyle(str, _T("font-weight"), m_strFontWeight);
	ExtractStyle(str, _T("font-style"), m_strFontStyle);
	ExtractStyle(str, _T("text-decoration"), m_strFontDecoration);
	ExtractStyle(str, _T("text-anchor"), m_strTextAnchor);
	ExtractStyle(str, _T("visibility"), m_strVisibility);
	ExtractStyle(str, _T("marker-start"), m_strMarkerStart);
	ExtractStyle(str, _T("marker-end"), m_strMarkerEnd);
}
//*******************************************************************************************************
void CBCGPSVGStyle::Add(const CBCGPSVGStyle& style)
{
	if (!style.m_strFill.IsEmpty())
	{
		m_strFill = style.m_strFill;
	}

	if (!style.m_strFillRule.IsEmpty())
	{
		m_strFillRule = style.m_strFillRule;
	}

	if (!style.m_strOpacity.IsEmpty())
	{
		m_strOpacity = style.m_strOpacity;
	}
	
	if (!style.m_strFillOpacity.IsEmpty())
	{
		m_strFillOpacity = style.m_strFillOpacity;
	}
	
	if (!style.m_strStroke.IsEmpty())
	{
		m_strStroke = style.m_strStroke;
	}

	if (!style.m_strStrokeWidth.IsEmpty())
	{
		m_strStrokeWidth = style.m_strStrokeWidth;
	}

	if (!style.m_strStrokeDashes.IsEmpty())
	{
		m_strStrokeDashes = style.m_strStrokeDashes;
	}

	if (!style.m_strStrokeDashOffset.IsEmpty())
	{
		m_strStrokeDashOffset = style.m_strStrokeDashOffset;
	}

	if (!style.m_strStrokeOpacity.IsEmpty())
	{
		m_strStrokeOpacity = style.m_strStrokeOpacity;
	}

	if (!style.m_strStrokeLineCap.IsEmpty())
	{
		m_strStrokeLineCap = style.m_strStrokeLineCap;
	}

	if (!style.m_strStrokeLineJoin.IsEmpty())
	{
		m_strStrokeLineJoin = style.m_strStrokeLineJoin;
	}

	if (!style.m_strClipPath.IsEmpty())
	{
		m_strClipPath = style.m_strClipPath;
	}

	if (!style.m_strMask.IsEmpty())
	{
		m_strMask = style.m_strMask;
	}

	if (!style.m_strFontFamily.IsEmpty())
	{
		m_strFontFamily = style.m_strFontFamily;
	}
	
	if (!style.m_strFontSize.IsEmpty())
	{
		m_strFontSize = style.m_strFontSize;
	}
	
	if (!style.m_strFontWeight.IsEmpty())
	{
		m_strFontWeight = style.m_strFontWeight;
	}

	if (!style.m_strFontStyle.IsEmpty())
	{
		m_strFontStyle = style.m_strFontStyle;
	}

	if (!style.m_strFontDecoration.IsEmpty())
	{
		m_strFontDecoration = style.m_strFontDecoration;
	}

	if (!style.m_strTextAnchor.IsEmpty())
	{
		m_strTextAnchor = style.m_strTextAnchor;
	}

	if (!style.m_strVisibility.IsEmpty())
	{
		m_strVisibility = style.m_strVisibility;
	}

	if (!style.m_strMarkerStart.IsEmpty())
	{
		m_strMarkerStart = style.m_strMarkerStart;
	}

	if (!style.m_strMarkerEnd.IsEmpty())
	{
		m_strMarkerEnd = style.m_strMarkerEnd;
	}
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGBase

static BOOL g_bNoAlphaBrush = FALSE;

CBCGPSVGBase::CBCGPSVGBase(CBCGPSVGBase* pParent, CBCGPXmlElement& elem)
{
	CommonInit();
	
	m_pParent = pParent;
	ReadAttributes(elem);
}
//*******************************************************************************************************
CBCGPSVGBase::CBCGPSVGBase()
{
	CommonInit();
}
//*******************************************************************************************************
CBCGPSVGBase::~CBCGPSVGBase()
{
	CleanUpTransforms();
}
//*******************************************************************************************************
CBCGPSVG& CBCGPSVGBase::GetRoot()
{
	CBCGPSVGBase* pRoot = this;
	while (pRoot->m_pParent != NULL)
	{
		pRoot = pRoot->m_pParent;
	}

	return (CBCGPSVG&)*pRoot;
}
//*******************************************************************************************************
void CBCGPSVGBase::CommonInit()
{
	m_dblStrokeWidth = -1.0;
	m_dblStrokeOpacity = -1.0;
	m_dblFillOpacity = -1.0;
	m_dblOpacity = -1.0;
	m_pParent = NULL;
	m_bIsVisible = TRUE;
	m_bHasTextAttributes = FALSE;
	m_dblMarkerStartAngle = 0.0;
	m_dblMarkerEndAngle = 0.0;
}
//*******************************************************************************************************
CBCGPSVGBase* CBCGPSVGBase::CreateCopy(CBCGPSVGBase* pParent)
{
	CRuntimeClass* pRTI = GetRuntimeClass();
	if (pRTI == NULL)
	{
		return NULL;
	}
	
	if (pRTI == RUNTIME_CLASS(CBCGPSVGSymbol))
	{
		pRTI = RUNTIME_CLASS(CBCGPSVGViewBox);
	}

	CObject* pObject = pRTI->CreateObject();
	if (pObject == NULL)
	{
		return NULL;
	}

	CBCGPSVGBase* pCopy = DYNAMIC_DOWNCAST(CBCGPSVGBase, pObject);
	if (pCopy == NULL)
	{
		delete pObject;
		return NULL;
	}

	pCopy->m_pParent = pParent;
	pCopy->CopyFrom(*this);

	pCopy->m_pParent = pParent;	// To be sure, that parent wasn't copied
	return pCopy;
}
//*******************************************************************************************************
void CBCGPSVGBase::ReadAttributes(CBCGPXmlElement& elem, BOOL bAddToExisting)
{
	m_strID = elem.GetStringAttribute(_T("id"));
	
	CString strClass = elem.GetStringAttribute(_T("class"));
	if (!strClass.IsEmpty())
	{
		m_strClass = _T(".") + strClass;
	}
	else if (!m_strID.IsEmpty())
	{
		m_strClass = _T("#") + m_strID;
	}
	
	if (!m_strClass.IsEmpty())
	{
		CBCGPSVGStyle style;
		if (GetRoot().FindStyle(m_strClass, style))
		{
			ApplyStyle(style);

			m_strClass.Empty();
			m_strDefaultClass.Empty();
		}
	}

	CreateBrushFromAttribute(elem, _T("stroke"), m_brStroke);
	CreateBrushFromAttribute(elem, _T("fill"), m_brFill);
	CreateBrushFromAttribute(elem, _T("color"), m_brCurrentColor);

	m_strFillRule = elem.GetStringAttribute(_T("fill-rule"));
	if (m_strFillRule == _T("inherit"))
	{
		m_strFillRule.Empty();
	}

	m_strMarkerStart = GetURLPath(elem.GetStringAttribute(_T("marker-start")));
	m_strMarkerEnd = GetURLPath(elem.GetStringAttribute(_T("marker-end")));
	
	m_dblStrokeWidth = ReadNumericAttribute(elem, _T("stroke-width"), m_dblStrokeWidth);
	m_dblOpacity = ReadNumericAttribute(elem, _T("opacity"), m_dblOpacity);
	m_dblFillOpacity = ReadNumericAttribute(elem, _T("fill-opacity"), m_dblFillOpacity);
	m_dblStrokeOpacity = ReadNumericAttribute(elem, _T("stroke-opacity"), m_dblStrokeOpacity);
	m_bIsVisible = elem.GetStringAttribute(_T("visibility")) != _T("hidden");
	
	ReadStyleAttributes(elem);
	ReadTransformAttributes(elem, _T("transform"), bAddToExisting);
	
	ReadStrokeStyleAttributes(elem);
	ReadClipPath(elem.GetStringAttribute(_T("clip-path")));
	ReadMask(elem.GetStringAttribute(_T("mask")));

	ReadTextFormatAttributes(elem);
}
//*******************************************************************************************************
void CBCGPSVGBase::AdaptColors(COLORREF clrBase, COLORREF clrTone, BOOL bClampHue /*= TRUE*/)
{
	m_brStroke.AdaptColors(clrBase, clrTone, bClampHue);

	if (m_brFill.IsEmpty() && !m_brFill.IsNone() && m_brFill.GetPath().IsEmpty() && clrBase == RGB(0, 0, 0))
	{
		m_brFill.SetColor(clrBase);
	}

	m_brFill.AdaptColors(clrBase, clrTone, bClampHue);
	m_brCurrentColor.AdaptColors(clrBase, clrTone, bClampHue);
}
//*******************************************************************************************************
void CBCGPSVGBase::InvertColors()
{
	m_brStroke.InvertColors();
	m_brFill.InvertColors();
	m_brCurrentColor.InvertColors();
}
//*******************************************************************************************************
void CBCGPSVGBase::ConvertToGrayScale(double dblLumRatio /*= 1.0*/)
{
	m_brStroke.ConvertToGrayScale(dblLumRatio);
	m_brFill.ConvertToGrayScale(dblLumRatio);
	m_brCurrentColor.ConvertToGrayScale(dblLumRatio);
}
//*******************************************************************************************************
void CBCGPSVGBase::MakeLighter(double dblRatio /*= .25*/)
{
	m_brStroke.MakeLighter(dblRatio);
	m_brFill.MakeLighter(dblRatio);
	m_brCurrentColor.MakeLighter(dblRatio);
}
//*******************************************************************************************************
void CBCGPSVGBase::MakeDarker(double dblRatio /*= .25*/)
{
	m_brStroke.MakeDarker(dblRatio);
	m_brFill.MakeDarker(dblRatio);
	m_brCurrentColor.MakeDarker(dblRatio);
}
//*******************************************************************************************************
void CBCGPSVGBase::Simplify()
{
	m_brStroke.Simplify();
	m_brFill.Simplify();
	m_brCurrentColor.Simplify();
}
//*******************************************************************************************************
void CBCGPSVGBase::CleanUpTransforms()
{
	m_arTransforms.RemoveAll();
}
//*******************************************************************************************************
BOOL CBCGPSVGBase::DrawMarkers(CBCGPGraphicsManager* pGM)
{
	BOOL bRes = FALSE;

	if (!m_strMarkerStart.IsEmpty())
	{
		CBCGPSVGMarker* pMarker = DYNAMIC_DOWNCAST(CBCGPSVGMarker, GetRoot().FindByID(m_strMarkerStart));
		if (pMarker != NULL)
		{
			ASSERT_VALID(pMarker);
			pMarker->DoDraw(pGM, m_ptMarkerStart, m_dblMarkerStartAngle);
			
			bRes = TRUE;
		}
	}

	if (!m_strMarkerEnd.IsEmpty())
	{
		CBCGPSVGMarker* pMarker = DYNAMIC_DOWNCAST(CBCGPSVGMarker, GetRoot().FindByID(m_strMarkerEnd));
		if (pMarker != NULL)
		{
			ASSERT_VALID(pMarker);
			pMarker->DoDraw(pGM, m_ptMarkerEnd, m_dblMarkerEndAngle);

			bRes = TRUE;
		}
	}

	return bRes;
}
//*******************************************************************************************************
void CBCGPSVGBase::CopyFrom(const CBCGPSVGBase& src)
{
	m_strID = src.m_strID;
	m_strDefaultClass = src.m_strDefaultClass;
	m_dblOpacity = src.m_dblOpacity;
	m_dblStrokeOpacity = src.m_dblStrokeOpacity;
	m_brStroke.CopyFrom(src.m_brStroke);
	m_brStrokeAlpha.CopyFrom(src.m_brStrokeAlpha);
	m_dblStrokeWidth = src.m_dblStrokeWidth;
	m_StrokeStyle.CopyFrom(src.m_StrokeStyle);
	m_dblFillOpacity = src.m_dblFillOpacity;
	m_brFill.CopyFrom(src.m_brFill);
	m_brCurrentColor.CopyFrom(src.m_brCurrentColor);
	m_strFillRule = src.m_strFillRule;
	m_brFillAlpha.CopyFrom(src.m_brFillAlpha);
	m_ptLocation = src.m_ptLocation;
	m_textFormat = src.m_textFormat;
	m_bHasTextAttributes = src.m_bHasTextAttributes;
	m_bIsVisible = src.m_bIsVisible;
	m_strClipPathID = src.m_strClipPathID;
	m_strMaskID = src.m_strMaskID;
	m_strMarkerStart = src.m_strMarkerStart;
	m_ptMarkerStart = src.m_ptMarkerStart;
	m_dblMarkerStartAngle = src.m_dblMarkerStartAngle;
	m_strMarkerEnd = src.m_strMarkerEnd;
	m_ptMarkerEnd = src.m_ptMarkerEnd;
	m_dblMarkerEndAngle = src.m_dblMarkerEndAngle;
	m_strClass = src.m_strClass;

	m_arTransforms.RemoveAll();
	m_arTransforms.Append(src.m_arTransforms);
}
//*******************************************************************************************************
BOOL CBCGPSVGBase::TranslateCoordinate(double& dblVal, BOOL bIsHorz, BOOL bIsSize)
{
	if (m_pParent != NULL)
	{
		return m_pParent->TranslateCoordinate(dblVal, bIsHorz, bIsSize);
	}

	return FALSE;
}
//*******************************************************************************************************
const CBCGPSVGBrush& CBCGPSVGBase::GetBrush(BOOL bIsFill, BOOL bAllowDefault)
{
	const CBCGPSVGBrush* pBrush = (bIsFill ? &m_brFill : &m_brStroke);

	if (pBrush->IsEmpty())
	{
		if (pBrush->IsNone())
		{
			return *pBrush;
		}

		const BOOL bIsCurrent = pBrush->IsCurrent();

		for (CBCGPSVGBase* pParent = bIsCurrent ? this : m_pParent; pParent != NULL; pParent = pParent->m_pParent)
		{
			if (bIsCurrent)
			{
				if (!pParent->m_brCurrentColor.IsEmpty())
				{
					pBrush = &pParent->m_brCurrentColor;
					break;
				}
			}
			else
			{
				pBrush = &pParent->GetBrush(bIsFill, FALSE);
				if (!pBrush->IsEmpty() || pBrush->IsNone())
				{
					break;
				}
			}
		}

		if (pBrush->IsEmpty() && !pBrush->IsNone() && !pBrush->IsCurrent() && bIsFill && bAllowDefault)
		{
			m_brFill.SetColor(GetRoot().GetDefaultColor());
			pBrush = &m_brFill;
		}
	}
	
	if (!pBrush->IsEmpty())
	{
		CString strPath = pBrush->GetPath();
		if (!strPath.IsEmpty())
		{
			CBCGPSVGBase* pElem = GetRoot().FindByID(strPath, TRUE);
			if (pElem != NULL)
			{
				const CBCGPSVGGradient* pGradientBrush = DYNAMIC_DOWNCAST(CBCGPSVGGradient, pElem);
				if (pGradientBrush != NULL)
				{
					pBrush = &pGradientBrush->GetBrush();
				}
			}
		}
		
		double dblOpacity = bIsFill ? GetFillOpacity() : GetStrokeOpacity();
		
		if (IsAlphaDrawingMode() && !g_bNoAlphaBrush)
		{
			CBCGPSVGBrush& brAlpha = bIsFill ? m_brFillAlpha : m_brStrokeAlpha;

			brAlpha.CopyFrom(*pBrush);
			brAlpha.PrepareAlphaMode(dblOpacity);

			return brAlpha;
		}
		else
		{
			if (dblOpacity >= 0.0)
			{
				((CBCGPSVGBrush*)pBrush)->SetOpacity(dblOpacity);
			}
		}
	}
	
	return *pBrush;
}
//*******************************************************************************************************
double CBCGPSVGBase::GetStrokeWidth(BOOL bAllowDefault)
{
	double dblStrokeWidth = m_dblStrokeWidth;
	
	if (dblStrokeWidth < 0.0)
	{
		if (m_pParent != NULL)
		{
			dblStrokeWidth = m_pParent->GetStrokeWidth(bAllowDefault);
		}
		else if (bAllowDefault)
		{
			dblStrokeWidth = 1.0;
		}
	}
	
	return dblStrokeWidth;
}
//*******************************************************************************************************
CBCGPStrokeStyle* CBCGPSVGBase::GetStrokeStyle()
{
	CBCGPStrokeStyle* pStyle = m_StrokeStyle.IsEmpty() ? NULL : &m_StrokeStyle;
	
	if (pStyle == NULL && m_pParent != NULL)
	{
		pStyle = m_pParent->GetStrokeStyle();
	}
	
	return pStyle;
}
//*******************************************************************************************************
double CBCGPSVGBase::GetFillOpacity()
{
	double dblOpacity = m_dblFillOpacity > 0 && m_dblOpacity > 0 ? (m_dblFillOpacity * m_dblOpacity) :
						m_dblFillOpacity < 0. ? m_dblOpacity : m_dblFillOpacity;
	
	if (dblOpacity < 0.0 && m_pParent != NULL)
	{
		dblOpacity = m_pParent->GetFillOpacity();
	}
	
	return dblOpacity;
}
//*******************************************************************************************************
double CBCGPSVGBase::GetStrokeOpacity()
{
	double dblOpacity = m_dblStrokeOpacity > 0 && m_dblOpacity > 0 ? (m_dblStrokeOpacity * m_dblOpacity) :
						m_dblStrokeOpacity < 0. ? m_dblOpacity : m_dblStrokeOpacity;
	
	if (dblOpacity < 0.0 && m_pParent != NULL)
	{
		dblOpacity = m_pParent->GetStrokeOpacity();
	}
	
	return dblOpacity;
}
//*******************************************************************************************************
CString CBCGPSVGBase::GetFillRule()
{
	CString strFillRule = m_strFillRule;

	if (strFillRule.IsEmpty() && m_pParent != NULL)
	{
		strFillRule = m_pParent->GetFillRule();
	}
	
	return strFillRule;
}
//*******************************************************************************************************
const CBCGPTextFormat& CBCGPSVGBase::GetTextFormat()
{
	const CBCGPTextFormat* tf = &m_textFormat;
	
	if (tf->IsEmpty())
	{
		if (m_pParent != NULL)
		{
			tf = &m_pParent->GetTextFormat();
		}
		else
		{
			tf = &GetRoot().GetDefaultTextFormat();
		}
	}
	
	return *tf;
}
//*******************************************************************************************************
BOOL CBCGPSVGBase::HasTransform() const
{
	return m_arTransforms.GetSize() > 0;
}
//*******************************************************************************************************
BOOL CBCGPSVGBase::SetTransfrom(CBCGPGraphicsManager* pGM)
{
	if (!HasTransform())
	{
		return FALSE;
	}

	pGM->PushTransform();

	for (int i = 0; i < (int)m_arTransforms.GetSize(); i++)
	{
		CBCGPTransformMatrix& transform = m_arTransforms[i];
		CBCGPTransformMatrix translate;
		CBCGPPoint ptCenter = transform.m_ptRotateCenter;

		if (ptCenter != CBCGPPoint(-DBL_MAX, -DBL_MAX) &&
			ptCenter != CBCGPPoint(DBL_MAX, DBL_MAX))
		{
			translate.CreateTranslate(ptCenter.x, ptCenter.y);
			pGM->SetMatrixTransform(translate.m_arMatrix);
		}

		pGM->SetMatrixTransform(m_arTransforms[i].m_arMatrix);

		if (!translate.IsEmpty())
		{
			translate.CreateTranslate(-ptCenter.x, -ptCenter.y);
			pGM->SetMatrixTransform(translate.m_arMatrix);
		}
	}
	
	return TRUE;
}
//*******************************************************************************************************
BOOL CBCGPSVGBase::ReadStyleAttributes(CBCGPXmlElement& elem, const CString& strAttribute)
{
	return ReadStyleAttributes(elem.GetStringAttribute(strAttribute));
}
//*******************************************************************************************************
BOOL CBCGPSVGBase::ReadStyleAttributes(const CString& str)
{
	if (str.IsEmpty())
	{
		return FALSE;
	}

	CBCGPSVGStyle style;
	style.Read(str);

	ApplyStyle(style);
	return TRUE;
}
//*******************************************************************************************************
void CBCGPSVGBase::ApplyStyle(const CBCGPSVGStyle style)
{
	if (!style.m_strFill.IsEmpty())
	{
		StringToBrush(style.m_strFill, m_brFill);
	}

	if (!style.m_strFillRule.IsEmpty())
	{
		m_strFillRule = style.m_strFillRule;
	}

	if (!style.m_strOpacity.IsEmpty())
	{
		m_dblOpacity = StringToNumber(style.m_strOpacity);
	}
	
	if (!style.m_strFillOpacity.IsEmpty())
	{
		m_dblFillOpacity = StringToNumber(style.m_strFillOpacity);
	}
	
	if (!style.m_strStroke.IsEmpty())
	{
		StringToBrush(style.m_strStroke, m_brStroke);
	}

	if (!style.m_strStrokeWidth.IsEmpty())
	{
		m_dblStrokeWidth = StringToNumber(style.m_strStrokeWidth);
	}

	if (!style.m_strStrokeDashes.IsEmpty())
	{
		StringToDashes(style.m_strStrokeDashes, m_StrokeStyle);
	}

	if (!style.m_strStrokeDashOffset.IsEmpty())
	{
		StringToDashOffset(style.m_strStrokeDashOffset, m_StrokeStyle);
	}

	if (!style.m_strStrokeOpacity.IsEmpty())
	{
		m_dblStrokeOpacity = StringToNumber(style.m_strStrokeOpacity);
	}

	if (!style.m_strStrokeLineCap.IsEmpty())
	{
		SetStokeCapStyle(style.m_strStrokeLineCap, m_StrokeStyle);
	}

	if (!style.m_strStrokeLineJoin.IsEmpty())
	{
		SetStokeJoinStyle(style.m_strStrokeLineJoin, m_StrokeStyle);
	}

	if (!style.m_strClipPath.IsEmpty())
	{
		ReadClipPath(style.m_strClipPath);
	}

	if (!style.m_strMask.IsEmpty())
	{
		ReadMask(style.m_strMask);
	}

	if (!style.m_strFontFamily.IsEmpty())
	{
		VeryfyTextFormat();
		m_textFormat.SetFontFamily(style.m_strFontFamily);
	}
	
	if (!style.m_strFontSize.IsEmpty())
	{
		VeryfyTextFormat();
		m_textFormat.SetFontSize((float)StringToNumber(style.m_strFontSize));
	}
	
	if (!style.m_strFontWeight.IsEmpty())
	{
		VeryfyTextFormat();
		SetTextFormatWeight(style.m_strFontWeight);
	}

	if (!style.m_strFontStyle.IsEmpty())
	{
		VeryfyTextFormat();
		SetTextFormatStyle(style.m_strFontStyle);
	}

	if (!style.m_strFontDecoration.IsEmpty())
	{
		VeryfyTextFormat();
		SetTextDecoration(style.m_strFontDecoration);
	}

	if (!style.m_strTextAnchor.IsEmpty())
	{
		VeryfyTextFormat();
		TextAnchorToAlignment(style.m_strTextAnchor);
	}

	if (!style.m_strVisibility.IsEmpty())
	{
		m_bIsVisible = style.m_strVisibility != _T("hidden");
	}

	if (!style.m_strMarkerStart.IsEmpty())
	{
		m_strMarkerStart = GetURLPath(style.m_strMarkerStart);
	}

	if (!style.m_strMarkerEnd.IsEmpty())
	{
		m_strMarkerEnd = GetURLPath(style.m_strMarkerEnd);
	}
}
//*******************************************************************************************************
void CBCGPSVGBase::ApplyStyle()
{
	if (m_pParent != NULL)
	{
		m_pParent->ApplyStyle();
	}

	CString strStyles = m_strClass;
	strStyles.TrimLeft();
	strStyles.TrimRight();

	BOOL bIsReady = FALSE;

	if (!strStyles.IsEmpty())
	{
		CString strStyle;

		for (int i = 0; i < strStyles.GetLength(); i++)
		{
			TCHAR c = strStyles[i];
			
			if (IsSpaceChar(c))
			{
				if (!strStyle.IsEmpty())
				{
					CBCGPSVGStyle style;
					if (GetRoot().FindStyle(strStyle, style))
					{
						ApplyStyle(style);
						bIsReady = TRUE;
					}

					strStyle = _T(".");
				}
			}
			else
			{
				strStyle += c;
			}
		}

		if (!strStyle.IsEmpty())
		{
			CBCGPSVGStyle style;
			if (GetRoot().FindStyle(strStyle, style))
			{
				ApplyStyle(style);
				bIsReady = TRUE;
			}
		}
	}

	if (!bIsReady && !m_strDefaultClass.IsEmpty())
	{
		CBCGPSVGStyle style;
		if (GetRoot().FindStyle(m_strDefaultClass, style))
		{
			ApplyStyle(style);
		}
	}
}
//*******************************************************************************************************
BOOL CBCGPSVGBase::ReadStrokeStyleAttributes(CBCGPXmlElement& elem)
{
	m_StrokeStyle.Destroy();

	SetStokeCapStyle(elem.GetStringAttribute(_T("stroke-linecap")), m_StrokeStyle);
	SetStokeJoinStyle(elem.GetStringAttribute(_T("stroke-linejoin")), m_StrokeStyle);

	StringToDashes(elem.GetStringAttribute(_T("stroke-dasharray")), m_StrokeStyle);
	StringToDashOffset(elem.GetStringAttribute(_T("stroke-dashoffset")), m_StrokeStyle);

	return !m_StrokeStyle.IsEmpty();
}
//*******************************************************************************************************
BOOL CBCGPSVGBase::ReadTransformAttributes(CBCGPXmlElement& elem, const CString& strAttribute, BOOL bAddToExisting)
{
	if (!bAddToExisting)
	{
		CleanUpTransforms();
	}

	CString str = elem.GetStringAttribute(strAttribute);
	if (str.IsEmpty())
	{
		return FALSE;
	}

	CStringArray arData;
	PrepareTransformData(str, arData);

	for (int i = 0; i < (int)arData.GetSize();)
	{
		CString strPart = arData[i];

		if (strPart == _T("translate"))
		{
			i++;

			double x = 0.;
			double y = 0.;

			if (ReadNumber(arData, i, x, FALSE))
			{
				ReadNumber(arData, i, y, FALSE);

				CBCGPTransformMatrix transform;
				transform.CreateTranslate(x, y);

				m_arTransforms.Add(transform);
			}
		}
		else if (strPart == _T("scale"))
		{
			i++;

			double x = 0.;
			double y = 0.;
			
			if (ReadNumber(arData, i, x, FALSE))
			{
				ReadNumber(arData, i, y, FALSE);

				CBCGPTransformMatrix transform;
				transform.CreateScale(x, y);				

				m_arTransforms.Add(transform);
			}
		}
		else if (strPart == _T("matrix"))
		{
			i++;
			
			CBCGPTransformMatrix transform;
			double val = 0.0;

			while (ReadNumber(arData, i, val, FALSE))
			{
				transform.m_arMatrix.Add(val);
			}

			m_arTransforms.Add(transform);
		}
		else if (strPart == _T("rotate"))
		{
			i++;
			
			double a = 0.;
			double x = DBL_MAX;
			double y = DBL_MAX;
			
			if (ReadNumber(arData, i, a, FALSE))
			{
				ReadNumber(arData, i, x, FALSE);
				ReadNumber(arData, i, y, FALSE);
				
				CBCGPTransformMatrix transform;
				transform.CreateRotate(a, x, y);

				m_arTransforms.Add(transform);
			}
		}
		else if (strPart == _T("skewX"))
		{
			i++;
			
			double a = 0.;
			if (ReadNumber(arData, i, a, FALSE))
			{
				CBCGPTransformMatrix transform;
				transform.CreateSkewX(a);
								
				m_arTransforms.Add(transform);
			}
		}
		else if (strPart == _T("skewY"))
		{
			i++;
			
			double a = 0.;
			if (ReadNumber(arData, i, a, FALSE))
			{
				CBCGPTransformMatrix transform;
				transform.CreateSkewY(a);

				m_arTransforms.Add(transform);
			}
		}
		else
		{
			i++;
		}
	}

	return TRUE;
}
//*******************************************************************************************************
double CBCGPSVGBase::ReadCoordinate(CBCGPXmlElement& elem, const CString& strAttribute, BOOL bIsHorz, BOOL bIsSize, double dblDefaultValue/* = 0.0*/)
{
	const CString str = elem.GetStringAttribute(strAttribute);
	double dblVal = StringToNumber(str, dblDefaultValue);

	if (str.Find(_T("%")) >= 0 && dblVal != 0.0)
	{
		TranslateCoordinate(dblVal, bIsHorz, bIsSize);
	}

	return dblVal;
}
//*******************************************************************************************************
double CBCGPSVGBase::ReadNumericAttribute(CBCGPXmlElement& elem, const CString& strAttribute, double dblDefaultValue, BOOL* pbIsAbsolute, BOOL bCheckFor01)
{
	CString str = elem.GetStringAttribute(strAttribute);
	double dblVal = StringToNumber(str, dblDefaultValue);

	if (pbIsAbsolute != NULL && !str.IsEmpty())
	{
		*pbIsAbsolute = str.Find(_T("%")) < 0;

		if (bCheckFor01 && *pbIsAbsolute && dblVal > 0 && dblVal <= 1)
		{
			*pbIsAbsolute = FALSE;
			dblVal *= 100;
		}
	}

	return dblVal;
}
//*******************************************************************************************************
void CBCGPSVGBase::ReadClipPath(const CString& strClipPath)
{
	if (!strClipPath.IsEmpty())
	{
		CString strTrimmed;
		
		for (int i = 0; i < strClipPath.GetLength(); i++)
		{
			TCHAR c = strClipPath[i];
			
			if (!IsSpaceChar(c))
			{
				strTrimmed += c;
			}
		}
		
		m_strClipPathID = GetURLPath(strTrimmed);
	}
}
//*******************************************************************************************************
void CBCGPSVGBase::ReadMask(const CString& strMask)
{
	if (!strMask.IsEmpty())
	{
		CString strTrimmed;
		
		for (int i = 0; i < strMask.GetLength(); i++)
		{
			TCHAR c = strMask[i];
			
			if (!IsSpaceChar(c))
			{
				strTrimmed += c;
			}
		}
		
		m_strMaskID = GetURLPath(strTrimmed);
	}
}
//*******************************************************************************************************
void CBCGPSVGBase::TextAnchorToAlignment(const CString& strTextAnchor)
{
	if (strTextAnchor == _T("start"))
	{
		m_textFormat.SetTextAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_LEADING);
	}
	else if (strTextAnchor == _T("end"))
	{
		m_textFormat.SetTextAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_TRAILING);
	}
	else  if (strTextAnchor == _T("middle"))
	{
		m_textFormat.SetTextAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_CENTER);
	}
}
//*******************************************************************************************************
void CBCGPSVGBase::SetTextFormatWeight(const CString& strWeight)
{
	long lFontWeight = FW_REGULAR;

	if (strWeight == _T("bold"))
	{
		lFontWeight = 700;
	}
	else if (strWeight == _T("normal"))
	{
		lFontWeight = 400;
	}
	else if (strWeight == _T("bolder"))
	{
		
	}
	else if (strWeight == _T("lighter"))
	{
		
	}
	else if (strWeight == _T("inherit"))
	{
		return;	
	}
	else
	{
		long lNum = _ttol(strWeight);
		if (lNum != 0)
		{
			lFontWeight = lNum;
		}
	}

	m_textFormat.SetFontWeight(lFontWeight);
}
//*******************************************************************************************************
void CBCGPSVGBase::SetTextFormatStyle(const CString& strStyle)
{
	CBCGPTextFormat::BCGP_FONT_STYLE style = CBCGPTextFormat::BCGP_FONT_STYLE_NORMAL;

	if (strStyle == _T("italic"))
	{
		style = CBCGPTextFormat::BCGP_FONT_STYLE_ITALIC;
	}
	else if (strStyle == _T("oblique"))
	{
		style = CBCGPTextFormat::BCGP_FONT_STYLE_OBLIQUE;
	}
	else if (strStyle == _T("inherit"))
	{
		return;
	}

	m_textFormat.SetFontStyle(style);
}
//*******************************************************************************************************
void CBCGPSVGBase::SetTextDecoration(const CString& strTextDecoration)
{
	if (strTextDecoration == _T("underline"))
	{
		m_textFormat.SetUnderline();
	}
	else if (strTextDecoration == _T("line-through"))
	{
		m_textFormat.SetStrikethrough();
	}
}
//*******************************************************************************************************
void CBCGPSVGBase::VeryfyTextFormat()
{
	if (!m_bHasTextAttributes)
	{
		m_textFormat = GetTextFormat();
		m_bHasTextAttributes = TRUE;
	}
}
//*******************************************************************************************************
CBCGPStrokeStyle* CBCGPSVGBase::PrepareStrokeStyle(double dblLineWidth, BOOL& bIsTemp)
{
	bIsTemp = FALSE;

	CBCGPStrokeStyle* pStrokeStyle = GetStrokeStyle();
	if (pStrokeStyle == NULL || dblLineWidth <= 1.0 || (pStrokeStyle->GetDashes().GetSize() == 0 && pStrokeStyle->GetDashOffset() == 0.0))
	{
		return pStrokeStyle;
	}

	bIsTemp = TRUE;

	CBCGPStrokeStyle* pCurrStrokeStyle = new CBCGPStrokeStyle;

	pCurrStrokeStyle->CopyFrom(*pStrokeStyle);

	CArray<FLOAT, FLOAT> arDashes;
	arDashes.Append(pCurrStrokeStyle->GetDashes());

	for (int i = 0; i < (int)arDashes.GetSize(); i++)
	{
		arDashes[i] /= (FLOAT)dblLineWidth;
	}

	pCurrStrokeStyle->SetDashes(arDashes);

	pCurrStrokeStyle->SetDashOffset(pCurrStrokeStyle->GetDashOffset() / (FLOAT)dblLineWidth);

	return pCurrStrokeStyle;
}
//*******************************************************************************************************
CBCGPGeometry::BCGP_FILL_MODE CBCGPSVGBase::PrepareFillMode()
{
	CString strFillRule = GetFillRule();
	return strFillRule == _T("evenodd") ? CBCGPGeometry::BCGP_FILL_MODE_ALTERNATE : CBCGPGeometry::BCGP_FILL_MODE_WINDING;
}
//*******************************************************************************************************
BOOL CBCGPSVGBase::ReadTextFormatAttributes(CBCGPXmlElement& elem)
{
	CString strFontFamily = elem.GetStringAttribute(_T("font-family"));
	if (!strFontFamily.IsEmpty())
	{
		VeryfyTextFormat();
		m_textFormat.SetFontFamily(strFontFamily);
	}

	CString strFontSize = elem.GetStringAttribute(_T("font-size"));
	if (!strFontSize.IsEmpty())
	{
		VeryfyTextFormat();
		m_textFormat.SetFontSize((float)StringToNumber(strFontSize));
	}
	
	CString strTextAnchor = elem.GetStringAttribute(_T("text-anchor"));
	if (!strTextAnchor.IsEmpty())
	{
		VeryfyTextFormat();
		TextAnchorToAlignment(strTextAnchor);
	}

	CString strWeight = elem.GetStringAttribute(_T("font-weight"));
	if (!strWeight.IsEmpty())
	{
		VeryfyTextFormat();
		SetTextFormatWeight(strWeight);
	}

	CString strStyle = elem.GetStringAttribute(_T("font-style"));
	if (!strStyle.IsEmpty())
	{
		VeryfyTextFormat();
		SetTextFormatStyle(strStyle);
	}

	CString strTextDecoration = elem.GetStringAttribute(_T("text-decoration"));
	if (!strTextDecoration.IsEmpty())
	{
		VeryfyTextFormat();
		SetTextDecoration(strTextDecoration);
	}

	return !m_textFormat.IsEmpty();
}
//*******************************************************************************************************
void CBCGPSVGBase::GetGeometryBounds(CBCGPRect& rectBounds)
{
	rectBounds.SetRectEmpty();

	CBCGPGeometry* pGeometry = GetGeometry();
	if (pGeometry == NULL)
	{
		return;
	}

	CArray<POINT, POINT> arPointsGDI;
	CArray<INT, INT> arPointsCount;

	pGeometry->GeometryToPolygon(arPointsGDI, arPointsCount, FALSE, NULL);

	if (arPointsGDI.GetSize() > 0)
	{
		rectBounds.left = rectBounds.right = arPointsGDI[0].x;
		rectBounds.top = rectBounds.bottom = arPointsGDI[0].y;

		for (int i = 1; i < (int)arPointsGDI.GetSize(); i++)
		{
			rectBounds.left = min(rectBounds.left, arPointsGDI[i].x);
			rectBounds.right = max(rectBounds.right, arPointsGDI[i].x);
			
			rectBounds.top = min(rectBounds.top, arPointsGDI[i].y);
			rectBounds.bottom = max(rectBounds.bottom, arPointsGDI[i].y);
		}
	}
}

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGRect

CBCGPSVGRect::CBCGPSVGRect(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) : CBCGPSVGBase(pParent, elem)
{
	m_strDefaultClass = _T("rect");

	m_ptLocation.x = ReadCoordinate(elem, _T("x"), TRUE, FALSE);
	m_ptLocation.y = ReadCoordinate(elem, _T("y"), FALSE, FALSE);
	
	m_size.cx = ReadCoordinate(elem, _T("width"), TRUE, TRUE);
	m_size.cy = ReadCoordinate(elem, _T("height"), FALSE, TRUE);
	
	m_sizeCornerRadius.cx = max(0., ReadNumericAttribute(elem, _T("rx")));
	m_sizeCornerRadius.cy = max(0., ReadNumericAttribute(elem, _T("ry")));
	
	if (m_sizeCornerRadius.cx > 0.0 && m_sizeCornerRadius.cy == 0.0)
	{
		m_sizeCornerRadius.cy = m_sizeCornerRadius.cx;
	}
	else if (m_sizeCornerRadius.cy > 0.0 && m_sizeCornerRadius.cx == 0.0)
	{
		m_sizeCornerRadius.cx = m_sizeCornerRadius.cy;
	}
}
//*******************************************************************************************************
BOOL CBCGPSVGRect::OnDraw(CBCGPGraphicsManager* pGM)
{
	ASSERT_VALID(pGM);

	CBCGPRect rect(m_ptLocation, m_size);

	CBCGPSVGAttributesStorage st(this);
	
	if (!m_sizeCornerRadius.IsNull())
	{
		CBCGPRoundedRect rectRounded(rect, m_sizeCornerRadius.cx, m_sizeCornerRadius.cy);
		
		pGM->FillRoundedRectangle(rectRounded, st.GetFillBrush());
		pGM->DrawRoundedRectangle(rectRounded, st.GetStrokeBrush(), st.GetStrokeWidth(), st.GetStrokeStyle());
	}
	else
	{
		pGM->FillRectangle(rect, st.GetFillBrush());
		pGM->DrawRectangle(rect, st.GetStrokeBrush(), st.GetStrokeWidth(), st.GetStrokeStyle());
	}
	
	return TRUE;
}
//*******************************************************************************************************
CBCGPRect CBCGPSVGRect::GetBounds()
{
	return CBCGPRect(m_ptLocation, m_size);
}
//*******************************************************************************************************
CBCGPGeometry* CBCGPSVGRect::CreateGeometry()
{
	CBCGPRect rect(m_ptLocation, m_size);
	
	if (!m_sizeCornerRadius.IsNull())
	{
		return new CBCGPRoundedRectangleGeometry(CBCGPRoundedRect(rect, m_sizeCornerRadius.cx, m_sizeCornerRadius.cy));
	}
	else
	{
		return new CBCGPRectangleGeometry(rect);
	}
}
//*******************************************************************************************************
void CBCGPSVGRect::CopyFrom(const CBCGPSVGBase& s)
{
	CBCGPSVGBase::CopyFrom(s);
	
	const CBCGPSVGRect& src = (const CBCGPSVGRect&)s;

	m_size = src.m_size;
	m_sizeCornerRadius = src.m_sizeCornerRadius;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGEllipse

CBCGPSVGEllipse::CBCGPSVGEllipse(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) : CBCGPSVGBase(pParent, elem)
{
	m_ptLocation.x = ReadCoordinate(elem, _T("cx"), TRUE, FALSE);
	m_ptLocation.y = ReadCoordinate(elem, _T("cy"), FALSE, FALSE);
	
	m_dblRadiusX = ReadCoordinate(elem, _T("r"), TRUE, TRUE);
	
	if (m_dblRadiusX == 0)
	{
		m_dblRadiusX = ReadCoordinate(elem, _T("rx"), TRUE, TRUE);
		m_dblRadiusY = ReadCoordinate(elem, _T("ry"), FALSE, TRUE);
	}
	else
	{
		m_dblRadiusY = m_dblRadiusX;
	}
}
//*******************************************************************************************************
BOOL CBCGPSVGEllipse::OnDraw(CBCGPGraphicsManager* pGM)
{
	ASSERT_VALID(pGM);
	
	CBCGPEllipse ellipse(m_ptLocation, m_dblRadiusX, m_dblRadiusY);
	CBCGPSVGAttributesStorage st(this);
	
	pGM->FillEllipse(ellipse, st.GetFillBrush());
	pGM->DrawEllipse(ellipse, st.GetStrokeBrush(), st.GetStrokeWidth(), st.GetStrokeStyle());
	
	return TRUE;
}
//*******************************************************************************************************
CBCGPRect CBCGPSVGEllipse::GetBounds()
{
	return CBCGPRect(
		m_ptLocation.x - m_dblRadiusX, m_ptLocation.y - m_dblRadiusY,
		m_ptLocation.x + m_dblRadiusX, m_ptLocation.y + m_dblRadiusY);
}
//*******************************************************************************************************
CBCGPGeometry* CBCGPSVGEllipse::CreateGeometry()
{
	return new CBCGPEllipseGeometry(CBCGPEllipse(m_ptLocation, m_dblRadiusX, m_dblRadiusY));
}
//*******************************************************************************************************
void CBCGPSVGEllipse::CopyFrom(const CBCGPSVGBase& s)
{
	CBCGPSVGBase::CopyFrom(s);
	
	const CBCGPSVGEllipse& src = (const CBCGPSVGEllipse&)s;

	m_dblRadiusX = src.m_dblRadiusX;
	m_dblRadiusY = src.m_dblRadiusY;
}

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGLine

CBCGPSVGLine::CBCGPSVGLine(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) : CBCGPSVGBase(pParent, elem)
{
	m_strDefaultClass = _T("line");

	m_ptLocation.x = ReadCoordinate(elem, _T("x1"), TRUE, FALSE);
	m_ptLocation.y = ReadCoordinate(elem, _T("y1"), FALSE, FALSE);
	
	m_ptTo.x = ReadCoordinate(elem, _T("x2"), TRUE, FALSE);
	m_ptTo.y = ReadCoordinate(elem, _T("y2"), FALSE, FALSE);

	if (!m_strMarkerStart.IsEmpty())
	{
		m_ptMarkerStart = m_ptLocation;
		m_dblMarkerStartAngle = bcg_rad2deg(bcg_angle(m_ptMarkerStart, m_ptTo));
	}
	
	if (!m_strMarkerEnd.IsEmpty())
	{
		m_ptMarkerEnd = m_ptTo;
		m_dblMarkerEndAngle = bcg_rad2deg(bcg_angle(m_ptLocation, m_ptMarkerEnd));
	}
}
//*******************************************************************************************************
BOOL CBCGPSVGLine::OnDraw(CBCGPGraphicsManager* pGM)
{
	ASSERT_VALID(pGM);
	
	CBCGPSVGAttributesStorage st(this);
	pGM->DrawLine(m_ptLocation.x, m_ptLocation.y, m_ptTo.x, m_ptTo.y, st.GetStrokeBrush(), st.GetStrokeWidth(), st.GetStrokeStyle());

	DrawMarkers(pGM);
	
	return TRUE;
}
//*******************************************************************************************************
void CBCGPSVGLine::CopyFrom(const CBCGPSVGBase& s)
{
	CBCGPSVGBase::CopyFrom(s);
	
	const CBCGPSVGLine& src = (const CBCGPSVGLine&)s;
	
	m_ptTo = src.m_ptTo;
}
//*******************************************************************************************************
CBCGPRect CBCGPSVGLine::GetBounds()
{
	return CBCGPRect(m_ptLocation, m_ptTo);	
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGText

static void CleanUpText(CString& strText)
{
	strText.Remove(_T('\r'));
	
	while (TRUE)
	{
		int iEOL = strText.Find(_T('\n'));
		if (iEOL < 0)
		{
			break;
		}
		
		int i = iEOL;
		
		for (; i < strText.GetLength() && IsSpaceChar(strText[i]); i++)
		{
		}
		
		strText.Delete(iEOL, i - iEOL);
		strText.Insert(iEOL, _T(' '));
	}
	
	strText.TrimLeft();
	strText.TrimRight();
}

CBCGPSVGText::CBCGPSVGText(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) : CBCGPSVGBase(pParent, elem)
{
	CommonInit();

	m_ptLocation.x = ReadCoordinate(elem, _T("x"), TRUE, FALSE, DBL_MAX);
	if (m_ptLocation.x == DBL_MAX)
	{
		m_ptLocation.x = 0.0;
		m_bIsAbsoluteX = FALSE;
	}

	m_ptLocation.x += ReadCoordinate(elem, _T("dx"), TRUE, TRUE);

	CString strY = elem.GetStringAttribute(_T("y"));
	if (!strY.IsEmpty())
	{
		m_bYIsEM = strY.GetLength() > 2 && strY.Right(2).CompareNoCase(_T("em")) == 0;
		if (m_bYIsEM)
		{
			strY = strY.Left(strY.GetLength() - 2);
		}
		
		double dblY = StringToNumber(strY);
		if (strY.Find(_T("%")) >= 0 && dblY != 0.0)
		{
			TranslateCoordinate(dblY , FALSE, TRUE);
		}

		m_ptLocation.y = dblY ;
	}
	else
	{
		m_bIsAbsoluteY = FALSE;
	}

	CString strDY = elem.GetStringAttribute(_T("dy"));
	if (!strDY.IsEmpty())
	{
		m_bDYIsEM = strDY.GetLength() > 2 && strDY.Right(2).CompareNoCase(_T("em")) == 0;
		if (m_bDYIsEM)
		{
			strDY = strDY.Left(strDY.GetLength() - 2);
		}
	
		m_dblDY = StringToNumber(strDY);
	}

	CBCGPXmlNodeList elements(elem.GetChildren());

	for (int i = 0; i < elements.GetLength(); i++)
	{
		CBCGPXmlNode node(elements.GetItem(i));
		if (!node.IsValid())
		{
			continue;
		}

		CString strName = node.GetNodeName();
		if (strName == _T("tspan"))
		{
			CBCGPXmlElement elemSpan(node);
			if (elemSpan.IsValid())
			{
				m_arSpans.Add(new CBCGPSVGText(this, elemSpan));
			}
		}
		else if (strName == _T("#text"))
		{
			CString str = node.GetText();
			str.TrimLeft();
			str.TrimRight();

			if (!str.IsEmpty())
			{
				if (m_strText.IsEmpty() && m_arSpans.GetSize() == 0)
				{
					m_strText = str;
				}
				else
				{
					CleanUpText(str);

					CBCGPSVGText* pText = new CBCGPSVGText();
					pText->m_strText = str;
					pText->m_pParent = this;

					pText->m_ptLocation.x = 0.0;
					pText->m_bIsAbsoluteX = FALSE;

					pText->m_ptLocation.y = 0.0;
					pText->m_bIsAbsoluteY = FALSE;

					m_arSpans.Add(pText);
				}
			}
		}
	}

	if (m_strText.IsEmpty() && elements.GetLength() <= 1)
	{
		m_strText = elem.GetText();
	}

	CleanUpText(m_strText);
}
//*******************************************************************************************************
void CBCGPSVGText::CommonInit()
{
	m_strDefaultClass = _T("text");

	m_bIsAbsoluteX = TRUE;
	m_bIsAbsoluteY = TRUE;
	m_bDYIsEM = FALSE;
	m_bYIsEM = FALSE;
	m_dblDY = 0.0;
}
//*******************************************************************************************************
void CBCGPSVGText::CleanUp()
{
	for (int i = 0; i < (int)m_arSpans.GetSize(); i++)
	{
		delete m_arSpans[i];
	}

	m_arSpans.RemoveAll();
}
//*******************************************************************************************************
CBCGPSVGText::~CBCGPSVGText()
{
	CleanUp();
}
//*******************************************************************************************************
BOOL CBCGPSVGText::OnDraw(CBCGPGraphicsManager* pGM)
{
	ASSERT_VALID(pGM);

	CBCGPSVGText* pTopText = this;
	while (TRUE)
	{
		CBCGPSVGText* pParent = DYNAMIC_DOWNCAST(CBCGPSVGText, pTopText->m_pParent);
		if (pParent == NULL)
		{
			break;
		}

		pTopText = pParent;
	}

	if (pTopText == this)
	{
		m_ptTotalOffset = m_ptLocation;
	}

	const CBCGPTextFormat& tf = GetTextFormat();
	
	if (!tf.IsEmpty() && !m_strText.IsEmpty())
	{
		CBCGPSize sizeText = pGM->GetTextSize(m_strText, tf);

		CBCGPPoint pt = m_ptLocation;
		if (m_bYIsEM)
		{
			pt.y *= fabs(tf.GetFontSize());
		}

		double dy = m_dblDY;
		if (m_bDYIsEM)
		{
			dy *= fabs(tf.GetFontSize());
		}

		pt.y += dy;

		if (!m_bIsAbsoluteX)
		{
			pt.x = pTopText->m_ptTotalOffset.x;
		}
		
		if (!m_bIsAbsoluteY)
		{
			pt.y += pTopText->m_ptTotalOffset.y;
		}

		CBCGPRect rectText(pt, sizeText);

		double cx = 0.0;

		switch (tf.GetTextAlignment())
		{
		case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_CENTER:
			cx = -sizeText.cx / 2;
			break;

		case CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_TRAILING:
			cx = -sizeText.cx;
			break;
		}

		rectText.OffsetRect(cx, -sizeText.cy);
		
		CBCGPSVGAttributesStorage st(this);
		pGM->DrawText(m_strText, rectText, tf, st.GetFillBrush());
		
		pTopText->m_ptTotalOffset = rectText.BottomRight();

		CString strExtra(_T("-"));
		pTopText->m_ptTotalOffset.x += pGM->GetTextSize(strExtra, tf).cx;
	}

	for (int i = 0; i < (int)m_arSpans.GetSize(); i++)
	{
		m_arSpans[i]->ApplyStyle();
		m_arSpans[i]->OnDraw(pGM);
	}

	return TRUE;
}
//*******************************************************************************************************
void CBCGPSVGText::CopyFrom(const CBCGPSVGBase& s)
{
	CBCGPSVGBase::CopyFrom(s);
	
	const CBCGPSVGText& src = (const CBCGPSVGText&)s;

	CleanUp();

	m_strText = src.m_strText;
	m_bIsAbsoluteX = src.m_bIsAbsoluteX;
	m_bIsAbsoluteY = src.m_bIsAbsoluteY;
	m_dblDY = src.m_dblDY;
	m_bYIsEM = src.m_bYIsEM;
	m_bDYIsEM = src.m_bDYIsEM;

	for (int i = 0; i < (int)src.m_arSpans.GetSize(); i++)
	{
		CBCGPSVGText* pSpan = DYNAMIC_DOWNCAST(CBCGPSVGText, src.m_arSpans[i]->CreateCopy(this));
		if (pSpan != NULL)
		{
			m_arSpans.Add(pSpan);
		}
	}
}
//*******************************************************************************************************
CBCGPRect CBCGPSVGText::GetBounds()
{
	return CBCGPRect();	// TODO
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGPolygon

CBCGPSVGPolygon::CBCGPSVGPolygon(CBCGPSVGBase* pParent, CBCGPXmlElement& elem, BOOL bIsClosed) : CBCGPSVGBase(pParent, elem)
{
	CBCGPPointsArray arPoints;
	
	CString str = elem.GetStringAttribute(_T("points"));
	if (!str.IsEmpty())
	{
		CStringArray arData;
		if (PreparePathData(str, arData) > 0)
		{
			arPoints.SetSize(0, arData.GetSize() / 2);
			
			for (int i = 0;;)
			{
				CBCGPPoint pt;
				if (!ReadPoint(arData, i, pt))
				{
					break;
				}
				
				arPoints.Add(pt);
			}
			
			m_Geometry.SetPoints(arPoints, bIsClosed);

			if (arPoints.GetSize() > 1)
			{
				if (!m_strMarkerStart.IsEmpty())
				{
					m_ptMarkerStart = arPoints[0];
					m_dblMarkerStartAngle = bcg_rad2deg(bcg_angle(m_ptMarkerStart, arPoints[1]));
				}

				if (!m_strMarkerEnd.IsEmpty())
				{
					m_ptMarkerEnd = arPoints[arPoints.GetSize() - 1];
					m_dblMarkerEndAngle = bcg_rad2deg(bcg_angle(arPoints[arPoints.GetSize() - 2], m_ptMarkerEnd));
				}
			}
		}
	}
}
//*******************************************************************************************************
BOOL CBCGPSVGPolygon::OnDraw(CBCGPGraphicsManager* pGM)
{
	ASSERT_VALID(pGM);
	
	m_Geometry.SetFillMode(PrepareFillMode());

	CBCGPSVGAttributesStorage st(this);

	pGM->FillGeometry(m_Geometry, st.GetFillBrush());
	pGM->DrawGeometry(m_Geometry, st.GetStrokeBrush(), st.GetStrokeWidth(), st.GetStrokeStyle());
	
	DrawMarkers(pGM);
	return TRUE;
}
//*******************************************************************************************************
CBCGPRect CBCGPSVGPolygon::GetBounds()
{
	if (m_rectBounds.IsRectNull())
	{
		GetGeometryBounds(m_rectBounds);
	}

	return m_rectBounds;
}
//*******************************************************************************************************
CBCGPGeometry* CBCGPSVGPolygon::GetGeometry() 
{ 
	return &m_Geometry;
}
//*******************************************************************************************************
CBCGPGeometry* CBCGPSVGPolygon::CreateGeometry()
{
	m_Geometry.SetFillMode(PrepareFillMode());

	CArray<CBCGPGeometry*, CBCGPGeometry*> arGeometry;
	arGeometry.Add(&m_Geometry);
	
	CBCGPGeometry* pGeometry = new CBCGPGeometryGroup(arGeometry);

	pGeometry->SetFillMode(m_Geometry.GetFillMode());
	return pGeometry;
}
//*******************************************************************************************************
void CBCGPSVGPolygon::CopyFrom(const CBCGPSVGBase& s)
{
	CBCGPSVGBase::CopyFrom(s);
	
	const CBCGPSVGPolygon& src = (const CBCGPSVGPolygon&)s;

	m_Geometry.CopyFrom(src.m_Geometry);
	m_rectBounds = src.m_rectBounds;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGPath

CBCGPSVGPath::CBCGPSVGPath(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) : CBCGPSVGBase(pParent, elem)
{
	m_strDefaultClass = _T("path");
	m_pGeometry = NULL;
	ReadPathAttribute(elem, _T("d"), m_arGeometry, GetErrors());
}
//*******************************************************************************************************
CBCGPSVGPath::~CBCGPSVGPath()
{
	CleanUp();
}
//*******************************************************************************************************
BOOL CBCGPSVGPath::OnDraw(CBCGPGraphicsManager* pGM)
{
	ASSERT_VALID(pGM);

	CBCGPGeometry* pGeometry = GetGeometry();
	if (pGeometry == NULL)
	{
		return FALSE;
	}

	CBCGPSVGAttributesStorage st(this);

	pGM->FillGeometry(*pGeometry, st.GetFillBrush());
	pGM->DrawGeometry(*pGeometry, st.GetStrokeBrush(), st.GetStrokeWidth(), st.GetStrokeStyle());
	
	DrawMarkers(pGM);
	return TRUE;
}
//*******************************************************************************************************
CBCGPRect CBCGPSVGPath::GetBounds()
{
	if (m_rectBounds.IsRectNull())
	{
		GetGeometryBounds(m_rectBounds);
	}

	return m_rectBounds;
}
//*******************************************************************************************************
BOOL CBCGPSVGPath::ReadPathAttribute(CBCGPXmlElement& elem, const CString& strAttribute, 
									 CArray<CBCGPComplexGeometry*, CBCGPComplexGeometry*>& arGeometry, CStringArray* pErrors)
{
	const CString strErrFmt = _T("CBCGPSVGPath::ReadPathAttribute[%d]: '%c' - cannot read point %d");

	int i = 0;

	for (i = 0; i < (int)arGeometry.GetSize(); i++)
	{
		delete arGeometry[i];
	}

	arGeometry.RemoveAll();

	CString strPath = elem.GetStringAttribute(strAttribute);
	if (strPath.IsEmpty())
	{
		return FALSE;
	}

	CStringArray arData;
	if (PreparePathData(strPath, arData) > 0)
	{
		arGeometry.SetSize(0, arData.GetSize() / 2);
	}

	CBCGPComplexGeometry* pGeometry = new CBCGPComplexGeometry(CBCGPPoint(-1., -1.), FALSE);
	arGeometry.Add(pGeometry);

	CBCGPPoint pt;
	CBCGPPoint ptPrev;
	CBCGPPoint ptPrevEnd;

	CBCGPPoint ptStart(DBL_MAX, DBL_MAX);
	CBCGPPoint ptFirst(DBL_MAX, DBL_MAX);

	double dblValue = 0.0;
	double dblCurrBearing = 0.0;
	BOOL bIsRelative = FALSE;
	TCHAR cPrev = 0;

	for (i = 0; i < (int)arData.GetSize();)
	{
		CString str = arData[i];

		if (str.GetLength() == 1 && _istalpha(str[0]))
		{
			i++;

			TCHAR c = str[0];
			bIsRelative = islower(c);

			CString strError;

			switch (toupper(c))
			{
			case _T('Z'):
				pGeometry->SetClosed();
				pt = pGeometry->GetStartPoint();

				if (i < (int)arData.GetSize() - 1)
				{
					pGeometry = new CBCGPComplexGeometry(CBCGPPoint(-1., -1.), FALSE);
					arGeometry.Add(pGeometry);
				}
				break;

			case _T('M'):

				if (i > 1 && toupper(cPrev) != _T('Z'))
				{
					pGeometry = new CBCGPComplexGeometry(CBCGPPoint(-1., -1.), FALSE);
					arGeometry.Add(pGeometry);
				}

				if (ReadPoint(arData, i, pt, bIsRelative))
				{
					pGeometry->SetStart(pt);
					ptPrev = pt;

					if (ptStart == CBCGPPoint(DBL_MAX, DBL_MAX))
					{
						ptStart = pt;
					}
				}
				break;

			case _T('L'):
				while (ReadPoint(arData, i, pt, bIsRelative))
				{
					pGeometry->AddLine(pt);
					ptPrevEnd = ptPrev;
					ptPrev = pt;

					if (ptFirst == CBCGPPoint(DBL_MAX, DBL_MAX))
					{
						ptFirst = ptPrev;
					}
				}
				break;

			case _T('H'):
				while (ReadNumber(arData, i, dblValue))
				{
					pt = CBCGPPoint(bIsRelative ? pt.x + dblValue : dblValue, pt.y);
					pt = ApplyBearing(ptPrev, pt, dblCurrBearing);

					pGeometry->AddLine(pt);

					ptPrevEnd = ptPrev;
					ptPrev = pt;

					if (ptFirst == CBCGPPoint(DBL_MAX, DBL_MAX))
					{
						ptFirst = ptPrev;
					}
				}
				break;

			case _T('V'):
				while (ReadNumber(arData, i, dblValue))
				{
					pt = CBCGPPoint(pt.x, bIsRelative ? pt.y + dblValue : dblValue);
					pt = ApplyBearing(ptPrev, pt, dblCurrBearing);

					pGeometry->AddLine(pt);

					ptPrevEnd = ptPrev;
					ptPrev = pt;

					if (ptFirst == CBCGPPoint(DBL_MAX, DBL_MAX))
					{
						ptFirst = ptPrev;
					}
				}
				break;

			case _T('C'):
				while (TRUE)
				{
					CBCGPPoint pt1 = pt;
					if (!ReadPoint(arData, i, pt1, bIsRelative))
					{
						break;
					}

					CBCGPPoint pt2 = pt;
					if (!ReadPoint(arData, i, pt2, bIsRelative))
					{
						strError.Format(strErrFmt, i, c, 2);
						break;
					}

					if (!ReadPoint(arData, i, pt, bIsRelative))
					{
						strError.Format(strErrFmt, i, c, 3);
						break;
					}
					
					pGeometry->AddBezier(pt1, pt2, pt);

					ptPrevEnd = pt;
					ptPrev = pt2;
					
					if (ptFirst == CBCGPPoint(DBL_MAX, DBL_MAX))
					{
						ptFirst = ptPrev;
					}
				}
				break;

			case _T('S'):
				while (TRUE)
				{
					CBCGPPoint pt1 = ptPrev;
					pt1.Reflection(pt);

					CBCGPPoint pt2 = pt;
					if (!ReadPoint(arData, i, pt2, bIsRelative))
					{
						break;
					}
					
					if (!ReadPoint(arData, i, pt, bIsRelative))
					{
						strError.Format(strErrFmt, i, c, 2);
						break;
					}
					
					pGeometry->AddBezier(pt1, pt2, pt);

					ptPrevEnd = pt;
					ptPrev = pt2;
					
					if (ptFirst == CBCGPPoint(DBL_MAX, DBL_MAX))
					{
						ptFirst = ptPrev;
					}
				}
				break;

			case _T('Q'):
				while (TRUE)
				{
					CBCGPPoint pt0 = pt;
					if (!ReadPoint(arData, i, pt0, bIsRelative))
					{
						break;
					}

					CBCGPPoint pt1(pt.x + 2.0 * (pt0.x - pt.x) / 3.0, pt.y + 2.0 * (pt0.y - pt.y) / 3.0);

					if (!ReadPoint(arData, i, pt, bIsRelative))
					{
						strError.Format(strErrFmt, i, c, 2);
						break;
					}

					CBCGPPoint pt2(pt.x + 2.0 * (pt0.x - pt.x) / 3.0, pt.y + 2.0 * (pt0.y - pt.y) / 3.0);

					pGeometry->AddBezier(pt1, pt2, pt);
					
					ptPrevEnd = pt;
					ptPrev = pt0;
					
					if (ptFirst == CBCGPPoint(DBL_MAX, DBL_MAX))
					{
						ptFirst = ptPrev;
					}
				}
				break;
				
			case _T('T'):
				while (TRUE)
				{
					CBCGPPoint pt0 = ptPrev;
					pt0.Reflection(pt);

					CBCGPPoint pt1(pt.x + 2.0 * (pt0.x - pt.x) / 3.0, pt.y + 2.0 * (pt0.y - pt.y) / 3.0);

					if (!ReadPoint(arData, i, pt, bIsRelative))
					{
						break;
					}

					CBCGPPoint pt2(pt.x + 2.0 * (pt0.x - pt.x) / 3.0, pt.y + 2.0 * (pt0.y - pt.y) / 3.0);

					pGeometry->AddBezier(pt1, pt2, pt);

					ptPrevEnd = pt;
					ptPrev = pt0;
					
					if (ptFirst == CBCGPPoint(DBL_MAX, DBL_MAX))
					{
						ptFirst = ptPrev;
					}
				}
				break;
				
			case _T('A'):
				while (TRUE)
				{
					CBCGPSize szRadius;
					if (!ReadSize(arData, i, szRadius))
					{
						break;
					}

					double dblRotationAngle = 0;
					if (!ReadNumber(arData, i, dblRotationAngle))
					{
						strError.Format(strErrFmt, i, c, 2);
						break;
					}

					double dblXAxisRotation  = 0;
					if (!ReadNumber(arData, i, dblXAxisRotation))
					{
						strError.Format(strErrFmt, i, c, 3);
						break;
					}
					
					double dblLargeArc = 0;
					if (!ReadNumber(arData, i, dblLargeArc))
					{
						strError.Format(strErrFmt, i, c, 4);
						break;
					}

					if (!ReadPoint(arData, i, pt, bIsRelative))
					{
						strError.Format(strErrFmt, i, c, 5);
						break;
					}

					pGeometry->AddArc(pt, szRadius, dblLargeArc != 0, dblXAxisRotation != 0, dblRotationAngle);

					ptPrevEnd = ptPrev;
					ptPrev = pt;
					
					if (ptFirst == CBCGPPoint(DBL_MAX, DBL_MAX))
					{
						ptFirst = ptPrev;
					}
				}
				break;
				
			case _T('B'):
				{
					double dblValueB = 0.0;

					if (ReadNumber(arData, i, dblValueB))
					{
						dblCurrBearing = bIsRelative ? dblCurrBearing + dblValueB : dblValueB;
					}
					else
					{
						dblCurrBearing = 0.0;
					}

				}
				break;

			default:
				if (pErrors != NULL)
				{
					CString strErr;
					strErr.Format(_T("CBCGPSVGImage: Unsupported command: %c"), c);

					pErrors->Add(strErr);
				}
				break;
			}

			if (pErrors != NULL && !strError.IsEmpty())
			{
				pErrors->Add(strError);
			}

			cPrev = c;
		}
		else
		{
			// Assume 'L':
			while (ReadPoint(arData, i, pt, bIsRelative))
			{
				pGeometry->AddLine(pt);
				ptPrevEnd = ptPrev;
				ptPrev = pt;
				
				if (ptFirst == CBCGPPoint(DBL_MAX, DBL_MAX))
				{
					ptFirst = ptPrev;
				}
			}
		}
	}

	if (!m_strMarkerStart.IsEmpty())
	{
		m_ptMarkerStart = ptStart;
		m_dblMarkerStartAngle = bcg_rad2deg(bcg_angle(ptStart, ptFirst));
	}

	if (!m_strMarkerEnd.IsEmpty())
	{
		m_ptMarkerEnd = ptPrev;
		m_dblMarkerEndAngle = bcg_rad2deg(bcg_angle(ptPrev, ptPrevEnd));
	}

	return TRUE;
}
//*******************************************************************************************************
void CBCGPSVGPath::CleanUp()
{
	if (m_pGeometry != NULL)
	{
		delete m_pGeometry;
		m_pGeometry = NULL;
	}
	
	for (int i = 0; i < (int)m_arGeometry.GetSize(); i++)
	{
		delete m_arGeometry[i];
	}
	
	m_arGeometry.RemoveAll();
}
//*******************************************************************************************************
CBCGPGeometry* CBCGPSVGPath::CreateGeometry()
{
	if (m_arGeometry.GetSize() == 0)
	{
		return NULL;
	}
	
	CArray<CBCGPGeometry*, CBCGPGeometry*> arGeometry;
	arGeometry.SetSize(m_arGeometry.GetSize());

	for (int i = 0; i < (int)m_arGeometry.GetSize(); i++)
	{
		arGeometry[i] = m_arGeometry[i];
	}

	CBCGPGeometry* pGeometry = new CBCGPGeometryGroup(arGeometry);
	
	pGeometry->SetFillMode(PrepareFillMode());
	return pGeometry;
}
//*******************************************************************************************************
CBCGPGeometry* CBCGPSVGPath::GetGeometry()
{
	if (m_pGeometry == NULL)
	{
		m_pGeometry = CreateGeometry();
	}

	return m_pGeometry;
}
//*******************************************************************************************************
void CBCGPSVGPath::CopyFrom(const CBCGPSVGBase& s)
{
	CBCGPSVGBase::CopyFrom(s);
	
	const CBCGPSVGPath& src = (const CBCGPSVGPath&)s;

	CleanUp();

	m_arGeometry.SetSize(src.m_arGeometry.GetSize());

	for (int i = 0; i < (int)src.m_arGeometry.GetSize(); i++)
	{
		CBCGPComplexGeometry* pSrcGeometry = src.m_arGeometry[i];

		CBCGPComplexGeometry* pGeometry = new CBCGPComplexGeometry;
		pGeometry->CopyFrom(*pSrcGeometry);

		m_arGeometry[i] = pGeometry;
	}

	m_rectBounds = src.m_rectBounds;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGGroup

CBCGPSVGGroup::CBCGPSVGGroup()
{
	m_bIsDefs = FALSE;
}
//*******************************************************************************************************
CBCGPSVGGroup::CBCGPSVGGroup(CBCGPSVGBase* pParent, CBCGPXmlElement& elem, BOOL bIsDefs) : CBCGPSVGBase(pParent, elem)
{
	m_bIsDefs = bIsDefs;
	AddChildren(elem);
}
//*******************************************************************************************************
CBCGPSVGGroup::~CBCGPSVGGroup()
{
	CleanUp();
}
//*******************************************************************************************************
BOOL CBCGPSVGGroup::AddChildren(CBCGPXmlElement& root)
{
	CleanUp();

	CStringArray* pErrors = GetErrors();
	CBCGPXmlNodeList elements(root.GetChildren());

	for (int i = 0; i < elements.GetLength(); i++)
	{
		CBCGPXmlNode node(elements.GetItem(i));
		if (!node.IsValid())
		{
			continue;
		}

		CBCGPXmlElement elem(node);
		if (!elem.IsValid())
		{
			continue;
		}

		CString strName = node.GetNodeName();

		if (strName == _T("a") || strName == _T("g"))
		{
			m_arElements.Add(new CBCGPSVGGroup(this, elem, FALSE));
		}
		else if (strName == _T("defs"))
		{
			m_arElements.Add(new CBCGPSVGGroup(this, elem, TRUE));
		}
		else if (strName == _T("svg"))
		{
			m_arElements.Add(new CBCGPSVG(this, elem));
		}
		else if (strName == _T("symbol"))
		{
			m_arElements.Add(new CBCGPSVGSymbol(this, elem));
		}
		else if (strName == _T("circle") || strName == _T("ellipse"))
		{
			m_arElements.Add(new CBCGPSVGEllipse(this, elem));
		}
		else if (strName == _T("rect"))
		{
			m_arElements.Add(new CBCGPSVGRect(this, elem));
		}
		else if (strName == _T("line"))
		{
			m_arElements.Add(new CBCGPSVGLine(this, elem));
		}
		else if (strName == _T("polyline"))
		{
			m_arElements.Add(new CBCGPSVGPolygon(this, elem, FALSE));
		}
		else if (strName == _T("polygon"))
		{
			m_arElements.Add(new CBCGPSVGPolygon(this, elem, TRUE));
		}
		else if (strName == _T("path"))
		{
			m_arElements.Add(new CBCGPSVGPath(this, elem));
		}
		else if (strName == _T("clipPath"))
		{
			m_arElements.Add(new CBCGPSVGClipPath(this, elem));
		}
		else if (strName == _T("mask"))
		{
			m_arElements.Add(new CBCGPSVGMask(this, elem));
		}
		else if (strName == _T("text"))
		{
			m_arElements.Add(new CBCGPSVGText(this, elem));
		}
		else if (strName == _T("marker"))
		{
			m_arElements.Add(new CBCGPSVGMarker(this, elem));
		}
		else if (strName == _T("linearGradient"))
		{
			m_arElements.Add(new CBCGPSVGLinearGradient(this, elem));
		}
		else if (strName == _T("radialGradient"))
		{
			m_arElements.Add(new CBCGPSVGRadialGradient(this, elem));
		}
		else if (strName == _T("use"))
		{
			CString strUsedPath;
			CBCGPSVGBase* pCopy = CreateUse(this, elem, strUsedPath);

			if (pCopy != NULL)
			{
				m_arElements.Add(pCopy);
			}
			else if (pErrors != NULL)
			{
				CString strErr;
				strErr.Format(_T("cannot create used element: %s"), (LPCTSTR)strUsedPath);
				
				pErrors->Add(strErr);
			}
		}
		else if (strName == _T("title"))
		{
			m_strTitle = node.GetText();
		}
		else if (strName == _T("style"))
		{
			ReadStyles(node.GetText());
		}
		else
		{
			if (pErrors != NULL)
			{
				CString strErr;
				strErr.Format(_T("CBCGPSVGImage: Unsupported element: %s"), (LPCTSTR)strName);
				
				pErrors->Add(strErr);
			}
		}
	}

	return TRUE;
}
//*******************************************************************************************************
CBCGPSVGBase* CBCGPSVGGroup::CreateUse(CBCGPSVGBase* pParent, CBCGPXmlElement& elem, CString& strUsedPath)
{
	const CString strXRef(_T("xlink:href"));
	const CString strHRef(_T("href"));
	
	strUsedPath = elem.GetStringAttribute(strXRef);
	if (strUsedPath.IsEmpty())
	{
		strUsedPath = elem.GetStringAttribute(strHRef);
		if (strUsedPath.IsEmpty())
		{
			return NULL;
		}
	}

	strUsedPath = strUsedPath.Mid(1);
	
	CBCGPSVGBase* pOrigin = GetRoot().FindByID(strUsedPath, FALSE);
	if (pOrigin == NULL && pParent != NULL)
	{
		for (CBCGPSVGBase* p = pParent; p != NULL && pOrigin == NULL; p = p->m_pParent)
		{
			pOrigin = p->FindByID(strUsedPath, FALSE);
		}
	}
	
	if (pOrigin == NULL)
	{
		return NULL;
	}

	ASSERT_VALID(pOrigin);

	CBCGPSVGBase* pCopy = pOrigin->CreateCopy(pParent);
	if (pCopy == NULL)
	{
		return NULL;
	}

	pCopy->ReadAttributes(elem, TRUE);
	pCopy->ReadStyleAttributes(elem);

	BOOL bIsAbsolute = TRUE;
	double x = pCopy->ReadNumericAttribute(elem, _T("x"), 0.0, &bIsAbsolute, FALSE);
	if (!bIsAbsolute)
	{
		if (x > 1.0)
		{
			x *= 0.01;
		}
		
		x *= GetRoot().GetSize().cx;
	}
	
	bIsAbsolute = TRUE;
	double y = pCopy->ReadNumericAttribute(elem, _T("y"), 0.0, &bIsAbsolute, FALSE);
	if (!bIsAbsolute)
	{
		if (y > 1.0)
		{
			y *= 0.01;
		}
		
		y *= GetRoot().GetSize().cy;
	}
	
	if (x != 0. || y != 0.)
	{
		CBCGPTransformMatrix transform;
		transform.CreateTranslate(x, y);
		
		pCopy->m_arTransforms.Add(transform);
	}
	
	double width = pCopy->ReadNumericAttribute(elem, _T("width"), 0.0, &bIsAbsolute);
	if (!bIsAbsolute && width > 1.0)
	{
		width *= 0.01;
	}
	
	bIsAbsolute = TRUE;
	double height = pCopy->ReadNumericAttribute(elem, _T("height"), 0.0, &bIsAbsolute);
	if (!bIsAbsolute && height > 1.0)
	{
		height *= 0.01;
	}
	
	if (width != 0. || height != 0.)
	{
		CBCGPTransformMatrix transform;
		double dblScaleX = 0.0;
		double dblScaleY = 0.0;
		
		if (bIsAbsolute)
		{
			CBCGPSize size = pOrigin->GetViewBox().Size();
			if (!size.IsEmpty())
			{
				dblScaleX = width == 0. ? 1.0 : width / size.cx;
				dblScaleY = height == 0. ? 1.0 : height / size.cy;
			}
		}
		else
		{
			dblScaleX = width == 0. ? 1.0 : width;
			dblScaleY = height == 0. ? 1.0 : height;
		}

		if (pOrigin->IsPreserveAspectRatio())
		{
			if (width != 0. && height != 0.)
			{
				dblScaleX = min(dblScaleX, dblScaleY);
				dblScaleY = dblScaleX;
			}
			else if (width != 0.)
			{
				dblScaleY = dblScaleX;
			}
			else
			{
				dblScaleX = dblScaleY;
			}
		}
		
		if (dblScaleX != 0. && dblScaleY != 0.)
		{
			transform.CreateScale(dblScaleX, dblScaleY);
			pCopy->m_arTransforms.Add(transform);
		}
	}

	return pCopy;
}
//*******************************************************************************************************
BOOL CBCGPSVGGroup::OnDraw(CBCGPGraphicsManager* pGM)
{
	if (m_bIsDefs)
	{
		return TRUE;
	}

	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];

		if (pElem->IsVisible())
		{
			CBCGPSVGElemTransform transform(pGM, pElem);
			
			CBCGPGeometry* pClipGeometry = NULL;
			CBCGPGeometry* pMaskGeometry = NULL;
			
			if (!pElem->m_strClipPathID.IsEmpty())
			{
				CBCGPSVGBase* pClipPath = GetRoot().FindByID(pElem->m_strClipPathID);
				if (pClipPath != NULL)
				{
					pClipGeometry = pClipPath->CreateGeometry();
					if (pClipGeometry != NULL)
					{
						pGM->SetClipArea(*pClipGeometry);
					}
				}
			}
			
			if (!pElem->m_strMaskID.IsEmpty())
			{
				CBCGPSVGMask* pMask = DYNAMIC_DOWNCAST(CBCGPSVGMask, GetRoot().FindByID(pElem->m_strMaskID));
				if (pMask != NULL && pMask->m_arElements.GetSize() > 0)
				{
					// Only 1-st elemnt is used!
					pMaskGeometry = pMask->m_arElements[0]->CreateGeometry();
					if (pMaskGeometry != NULL)
					{
						g_bNoAlphaBrush = TRUE;
						pGM->SetMask(*pMaskGeometry, pMask->m_arElements[0]->GetFillBrush());
						g_bNoAlphaBrush = FALSE;
					}
				}
			}

			pElem->OnDraw(pGM);

			if (pMaskGeometry != NULL)
			{
				pGM->ReleaseMask();
				delete pMaskGeometry;
			}

			if (pClipGeometry != NULL)
			{
				pGM->ReleaseClipArea();
				delete pClipGeometry;
			}
		}
	}

	return TRUE;
}
//*******************************************************************************************************
void CBCGPSVGGroup::CopyFrom(const CBCGPSVGBase& s)
{
	CStringArray* pErrors = GetErrors();

	CBCGPSVGBase::CopyFrom(s);
	
	const CBCGPSVGGroup& src = (const CBCGPSVGGroup&)s;

	CleanUp();

	m_arElements.SetSize(0, src.m_arElements.GetSize());

	for (int i = 0; i < (int)src.m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = src.m_arElements[i];
		ASSERT_VALID(pElem);

		CBCGPSVGBase* pElemCopy = pElem->CreateCopy(this);
		if (pElemCopy == NULL)
		{
			if (pErrors != NULL)
			{
				CString strErr;
				strErr.Format(_T("CBCGPSVGGroup: cannot create copy of element: %s"), (LPCTSTR)pElem->m_strID);
				
				pErrors->Add(strErr);
			}
		}
		else
		{
			m_arElements.Add(pElemCopy);
		}
	}

	m_strTitle = src.m_strTitle;
	m_bIsDefs = src.m_bIsDefs;

	m_mapStyles.RemoveAll();

	for (POSITION pos = src.m_mapStyles.GetStartPosition(); pos != NULL;)
	{
		CString strName;
		CBCGPSVGStyle style;
		
		src.m_mapStyles.GetNextAssoc(pos, strName, style);
		m_mapStyles[strName] = style;
	}
}
//*******************************************************************************************************
CBCGPRect CBCGPSVGGroup::GetBounds()
{
	CBCGPRect rect;
	
	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];
		ASSERT_VALID(pElem);

		CBCGPRect rectElem = pElem->GetBounds();
		if (!rectElem.IsRectNull())
		{
			rect.UnionRect(rect, rectElem);
		}
	}

	return rect;
}
//*******************************************************************************************************
void CBCGPSVGGroup::AdaptColors(COLORREF clrBase, COLORREF clrTone, BOOL bClampHue /*= TRUE*/)
{
	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];
		ASSERT_VALID(pElem);

		pElem->AdaptColors(clrBase, clrTone, bClampHue);
	}
}
//*******************************************************************************************************
void CBCGPSVGGroup::InvertColors()
{
	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];
		ASSERT_VALID(pElem);
		
		pElem->InvertColors();
	}
}
//*******************************************************************************************************
void CBCGPSVGGroup::ConvertToGrayScale(double dblLumRatio /*= 1.0*/)
{
	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];
		ASSERT_VALID(pElem);
		
		pElem->ConvertToGrayScale(dblLumRatio);
	}
}
//*******************************************************************************************************
void CBCGPSVGGroup::MakeLighter(double dblRatio /*= .25*/)
{
	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];
		ASSERT_VALID(pElem);
		
		pElem->MakeLighter(dblRatio);
	}
}
//*******************************************************************************************************
void CBCGPSVGGroup::MakeDarker(double dblRatio /*= .25*/)
{
	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];
		ASSERT_VALID(pElem);
		
		pElem->MakeDarker(dblRatio);
	}
}
//*******************************************************************************************************
void CBCGPSVGGroup::Simplify()
{
	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];
		ASSERT_VALID(pElem);
		
		pElem->Simplify();
	}
}
//*******************************************************************************************************
CBCGPSVGBase* CBCGPSVGGroup::FindByID(const CString& strID, BOOL bIsDefsOnly)
{
	if (bIsDefsOnly && !m_bIsDefs && m_pParent != NULL)
	{
		return NULL;
	}

	if (m_strID == strID && !strID.IsEmpty())
	{
		return this;
	}

	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i]->FindByID(strID);
		if (pElem != NULL)
		{
			return pElem;
		}
	}

	return NULL;
}
//*******************************************************************************************************
void CBCGPSVGGroup::CleanUp()
{
	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		delete m_arElements[i];
	}
	
	m_arElements.RemoveAll();
}
//*******************************************************************************************************
BOOL CBCGPSVGGroup::FindStyle(const CString& strClassName, CBCGPSVGStyle& style)
{
	if (strClassName.IsEmpty())
	{
		return FALSE;
	}

	if (m_mapStyles.Lookup(strClassName, style))
	{
		return TRUE;
	}

	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];
		ASSERT_VALID(pElem);

		if (pElem->FindStyle(strClassName, style))
		{
			return TRUE;
		}
	}

	return FALSE;
}
//*******************************************************************************************************
void CBCGPSVGGroup::OnBeforeDraw(CBCGPGraphicsManager* pGM)
{
	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];
		ASSERT_VALID(pElem);

		pElem->OnBeforeDraw(pGM);
	}
}
//*******************************************************************************************************
void CBCGPSVGGroup::ReadStyles(const CString& str)
{
	m_mapStyles.RemoveAll();

	CString strCurrName;
	CString strCurrStyle;
	int nCount = 0;

	for (int i = 0; i < str.GetLength(); i++)
	{
		TCHAR c = str[i];
		if (IsSpaceChar(c))
		{
			continue;
		}

		if (c == _T('{'))
		{
			nCount++;
			continue;
		}

		if (c == _T('}'))
		{
			nCount--;
			nCount = max(0, nCount);

			if (nCount == 0)
			{
				if (!strCurrName.IsEmpty() && !strCurrStyle.IsEmpty())
				{
					if (strCurrName == _T("svg"))
					{
						ReadStyleAttributes(strCurrStyle);
					}
					else
					{
						CBCGPSVGStyle style;
						style.Read(strCurrStyle);
						
						for (int iStart = 0; iStart < strCurrName.GetLength();)
						{
							int iNextStart = strCurrName.Find(_T(","), iStart);

							CString strCurr = iNextStart < 0 ? strCurrName.Mid(iStart) : strCurrName.Mid(iStart, iNextStart - iStart);
							strCurr.TrimLeft();
							strCurr.TrimRight();
							
							if (!strCurr.IsEmpty())
							{
								CBCGPSVGStyle styleOld;
								if (m_mapStyles.Lookup(strCurr, styleOld))
								{
									styleOld.Add(style);
									m_mapStyles.SetAt(strCurr, styleOld);
								}
								else
								{
									m_mapStyles.SetAt(strCurr, style);
								}
							}

							if (iNextStart < 0)
							{
								break;
							}

							iStart = iNextStart + 1;
						}
					}
				}

				strCurrName.Empty();
				strCurrStyle.Empty();
			}

			continue;
		}

		if (nCount > 0)
		{
			strCurrStyle += c;
		}
		else
		{
			strCurrName += c;
		}
	}

	if (m_mapStyles.IsEmpty())
	{
		ReadStyleAttributes(str);
	}
}
//*******************************************************************************************************
void CBCGPSVGGroup::OnApplyStyle()
{
	for (int i = 0; i < (int)m_arElements.GetSize(); i++)
	{
		CBCGPSVGBase* pElem = m_arElements[i];
		ASSERT_VALID(pElem);
		
		pElem->OnApplyStyle();
	}
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGClipPath

CBCGPSVGClipPath::CBCGPSVGClipPath(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) 
	: CBCGPSVGGroup(pParent, elem, FALSE)
{
}
//*******************************************************************************************************
CBCGPSVGClipPath::~CBCGPSVGClipPath()
{
	CleanUp();
}
//*******************************************************************************************************
CBCGPGeometry* CBCGPSVGClipPath::CreateGeometry()
{
	if (m_arGeometry.GetSize() == 0)
	{
		for (int i = 0; i < (int)m_arElements.GetSize(); i++)
		{
			CBCGPGeometry* pGeometry = m_arElements[i]->CreateGeometry();
			
			if (pGeometry != NULL)
			{
				m_arGeometry.Add(pGeometry);
			}
		}
	}

	if (m_arGeometry.GetSize() > 0)
	{
		return new CBCGPGeometryGroup(m_arGeometry);
	}

	return NULL;
}
//*******************************************************************************************************
void CBCGPSVGClipPath::CopyFrom(const CBCGPSVGBase& s)
{
	CBCGPSVGGroup::CopyFrom(s);
	CleanUp();
}
//*******************************************************************************************************
void CBCGPSVGClipPath::CleanUp()
{
	for (int i = 0; i < (int)m_arGeometry.GetSize(); i++)
	{
		delete m_arGeometry[i];
	}
	
	m_arGeometry.RemoveAll();
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGGradient

CBCGPSVGGradient::CBCGPSVGGradient(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) :
	CBCGPSVGBase(pParent, elem)
{
	m_pBase = NULL;

	CString strLink = elem.GetStringAttribute(_T("xlink:href"));
	if (strLink.IsEmpty())
	{
		strLink = elem.GetStringAttribute(_T("href"));
	}

	if (!strLink.IsEmpty())
	{
		strLink = strLink.Mid(1);	// Remove 1-st '#'
		CBCGPSVGBase* pOrigin = GetRoot().FindByID(strLink, FALSE);
		if (pOrigin == NULL && pParent != NULL)
		{
			for (CBCGPSVGBase* p = pParent; p != NULL && pOrigin == NULL; p = p->m_pParent)
			{
				pOrigin = p->FindByID(strLink, FALSE);
			}
		}

		if (pOrigin != NULL)
		{
			m_pBase = DYNAMIC_DOWNCAST(CBCGPSVGGradient, pOrigin);
		}
	}

	ReadTransformAttributes(elem, _T("gradientTransform"));

	m_bUserSpaceOnUse = elem.GetStringAttribute(_T("gradientUnits")) == _T("userSpaceOnUse");

	CString strSpreadMethod = elem.GetStringAttribute(_T("spreadMethod"));
	if (strSpreadMethod == _T("repeat"))
	{
		m_extendMode = CBCGPBrush::BCGP_GRADIENT_EXTEND_MODE_WRAP;
	}
	else if (strSpreadMethod == _T("reflect"))
	{
		m_extendMode = CBCGPBrush::BCGP_GRADIENT_EXTEND_MODE_MIRROR;
	}
	else
	{
		m_extendMode = CBCGPBrush::BCGP_GRADIENT_EXTEND_MODE_CLAMP;
	}

	CBCGPXmlNodeList elements(elem.GetChildren());
	ReadGradientStops(elements, m_stops);

	for (CBCGPSVGGradient* pBase = m_pBase; pBase != NULL && m_stops.GetSize() == 0; pBase = pBase->m_pBase)
	{
		if (pBase->m_stops.GetSize() > 0)
		{
			m_stops.Append(pBase->m_stops);
		}
	}
}
//*******************************************************************************************************	
void CBCGPSVGGradient::CopyFrom(const CBCGPSVGBase& s)
{
	CBCGPSVGBase::CopyFrom(s);
	
	const CBCGPSVGGradient& src = (const CBCGPSVGGradient&)s;

	m_Brush.CopyFrom(src.m_Brush);
	
	m_stops.RemoveAll();
	m_stops.Append(src.m_stops);
	m_pBase = NULL;
	m_bUserSpaceOnUse = src.m_bUserSpaceOnUse;
	m_extendMode = src.m_extendMode;
}
//*******************************************************************************************************	
void CBCGPSVGGradient::AdaptColors(COLORREF clrBase, COLORREF clrTone, BOOL bClampHue /*= TRUE*/)
{
	m_Brush.AdaptColors(clrBase, clrTone, bClampHue);

	for (int i = 0; i < m_stops.GetSize(); i++)
	{
		m_stops[i].m_Color.AdaptColors(clrBase, clrTone, bClampHue);
	}
}
//*******************************************************************************************************	
void CBCGPSVGGradient::InvertColors()
{
	m_Brush.InvertColors();
	
	for (int i = 0; i < m_stops.GetSize(); i++)
	{
		m_stops[i].m_Color.InvertColors();
	}
}
//*******************************************************************************************************	
void CBCGPSVGGradient::ConvertToGrayScale(double dblLumRatio /*= 1.0*/)
{
	m_Brush.ConvertToGrayScale(dblLumRatio);
	
	for (int i = 0; i < m_stops.GetSize(); i++)
	{
		m_stops[i].m_Color.ConvertToGrayScale(dblLumRatio);
	}
}
//*******************************************************************************************************	
void CBCGPSVGGradient::MakeLighter(double dblRatio /*= .25*/)
{
	m_Brush.MakeLighter(dblRatio);
	
	for (int i = 0; i < m_stops.GetSize(); i++)
	{
		m_stops[i].m_Color.MakeLighter(dblRatio);
	}
}
//*******************************************************************************************************	
void CBCGPSVGGradient::MakeDarker(double dblRatio /*= .25*/)
{
	m_Brush.MakeDarker(dblRatio);
	
	for (int i = 0; i < m_stops.GetSize(); i++)
	{
		m_stops[i].m_Color.MakeDarker(dblRatio);
	}
}
//*******************************************************************************************************	
void CBCGPSVGGradient::Simplify()
{
	m_Brush.Simplify();
	
	for (int i = 0; i < m_stops.GetSize(); i++)
	{
		m_stops[i].m_Color.Simplify();
	}
}
//*******************************************************************************************************
int CBCGPSVGGradient::ReadGradientStops(CBCGPXmlNodeList& elements, CBCGPBrushGradientStops& stops)
{
	stops.RemoveAll();

	for (int i = 0; i < elements.GetLength(); i++)
	{
		CBCGPXmlNode node(elements.GetItem(i));
		if (!node.IsValid())
		{
			continue;
		}
		
		CString strName = node.GetNodeName();

		CBCGPXmlElement elem(node);
		if (!elem.IsValid())
		{
			continue;
		}
		
		if (strName == _T("stop"))
		{
			CString strStopColor;
			CString strStopOpacity;
			
			CString strStyle = elem.GetStringAttribute(_T("style"));
			if (!strStyle.IsEmpty())
			{
				ExtractStyle(strStyle, _T("stop-color"), strStopColor);
				ExtractStyle(strStyle, _T("stop-opacity"), strStopOpacity);
			}

			if (strStopColor.IsEmpty())
			{
				strStopColor = elem.GetStringAttribute(_T("stop-color"));
			}

			if (strStopOpacity.IsEmpty())
			{
				strStopOpacity = elem.GetStringAttribute(_T("stop-opacity"));
			}

			CBCGPBrushGradientStop stop;
			stop.m_Color = CBCGPColor(StringToColor(strStopColor), StringToNumber(strStopOpacity, 1.0));

			CString strOffset = elem.GetStringAttribute(_T("offset"));

			stop.m_dblOffset = StringToNumber(strOffset, 0.0);
			if (strOffset.Find(_T("%")) >= 0)
			{
				stop.m_dblOffset *= 0.01;
			}

			stops.Add(stop);
		}
	}

	return (int)stops.GetSize();
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGLinearGradient

CBCGPSVGLinearGradient::CBCGPSVGLinearGradient(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) :
	CBCGPSVGGradient(pParent, elem)
{
	CBCGPGradientPoint pt1;
	BOOL bIsAbsoluteX = FALSE;
	BOOL bIsAbsoluteY = FALSE;

	pt1.x = ReadNumericAttribute(elem, _T("x1"), 0.0, &bIsAbsoluteX);
	pt1.y = ReadNumericAttribute(elem, _T("y1"), 0.0, &bIsAbsoluteY);

	pt1.bIsAbsolute = bIsAbsoluteX && bIsAbsoluteY && (pt1.x != 0 || pt1.y != 0);
	
	CBCGPGradientPoint pt2;
	bIsAbsoluteX = FALSE;
	bIsAbsoluteY = FALSE;

	pt2.x = ReadNumericAttribute(elem, _T("x2"), 100.0, &bIsAbsoluteX);
	pt2.y = ReadNumericAttribute(elem, _T("y2"), 0.0, &bIsAbsoluteY);

	pt2.bIsAbsolute = bIsAbsoluteX && bIsAbsoluteY && (pt2.x != 0 || pt2.y != 0);

	m_Brush.SetLinearGradientStops(m_stops, pt1, pt2, m_arTransforms, m_extendMode);
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGRadialGradient
	
CBCGPSVGRadialGradient::CBCGPSVGRadialGradient(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) :
	CBCGPSVGGradient(pParent, elem)
{
	CBCGPGradientPoint center;
	BOOL bIsAbsoluteX = FALSE;
	BOOL bIsAbsoluteY = FALSE;

	center.x = ReadNumericAttribute(elem, _T("cx"), 50.0, &bIsAbsoluteX);
	center.y = ReadNumericAttribute(elem, _T("cy"), 50.0, &bIsAbsoluteY);
	
	center.bIsAbsolute = bIsAbsoluteX && bIsAbsoluteY && (center.x != 0 || center.y != 0);

	CBCGPGradientPoint f;
	bIsAbsoluteX = center.bIsAbsolute;
	bIsAbsoluteY = center.bIsAbsolute;

	f.x = ReadNumericAttribute(elem, _T("fx"), center.x, &bIsAbsoluteX);
	f.y = ReadNumericAttribute(elem, _T("fy"), center.y, &bIsAbsoluteY);

	f.bIsAbsolute = bIsAbsoluteX && bIsAbsoluteY;

	CBCGPGradientPoint r;
	r.bIsAbsolute = FALSE;
	r.x = ReadNumericAttribute(elem, _T("r"), 50.0, &r.bIsAbsolute);

	if ((!center.bIsAbsolute || !f.bIsAbsolute) && r.bIsAbsolute && !m_bUserSpaceOnUse)
	{
		r.bIsAbsolute = FALSE;
		r.x *= 100.0;
	}

	r.y = r.x;
	
	m_Brush.SetRadialGradientStops(m_stops, center, r, f, m_arTransforms, m_extendMode);
}

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGMask

CBCGPSVGMask::CBCGPSVGMask()
{
}
//*******************************************************************************************************
CBCGPSVGMask::CBCGPSVGMask(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) : CBCGPSVGGroup(pParent, elem, FALSE)
{
	m_ptLocation.x = ReadCoordinate(elem, _T("x"), TRUE, FALSE);
	m_ptLocation.y = ReadCoordinate(elem, _T("y"), TRUE, FALSE);
}

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGViewBox

CBCGPSVGViewBox::CBCGPSVGViewBox()
{
	m_sizeScale = CBCGPSize(1., 1.);
	m_bPreserveAspectRatio = FALSE;
	m_bIsMirror = FALSE;
}
//*******************************************************************************************************
CBCGPSVGViewBox::CBCGPSVGViewBox(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) : CBCGPSVGGroup(pParent, elem, FALSE)
{
	m_sizeScale = CBCGPSize(1., 1.);
	m_bPreserveAspectRatio = FALSE;
	m_bIsMirror = FALSE;

	CBCGPSVGBase* pRoot = this;
	while (pRoot->m_pParent != NULL)
	{
		pRoot = pRoot->m_pParent;
	}
	
	CreateData(elem);
}
//*******************************************************************************************************
BOOL CBCGPSVGViewBox::OnDraw(CBCGPGraphicsManager* pGM)
{
	CBCGPSVGEViewBoxTransform viewTransform(pGM, m_rectViewBox, 
		CBCGPSize(
			m_sizeScale.cx * GetExportScaleRatio(), 
			m_sizeScale.cy * GetExportScaleRatio()), m_ptLocation, m_bIsMirror);

	CBCGPSVGElemTransform* pTransform = NULL;

	if (m_pParent == NULL)
	{
		pTransform = new CBCGPSVGElemTransform(pGM, this);
	}
	
	BOOL bRes = CBCGPSVGGroup::OnDraw(pGM);

	if (pTransform != NULL)
	{
		delete pTransform;
	}

	return bRes;
}
//*******************************************************************************************************
void CBCGPSVGViewBox::CopyFrom(const CBCGPSVGBase& s)
{
	CBCGPSVGGroup::CopyFrom(s);
	
	const CBCGPSVGViewBox& src = (const CBCGPSVGViewBox&)s;
	
	m_rectViewBox = src.m_rectViewBox;
	m_bPreserveAspectRatio = src.m_bPreserveAspectRatio;
	m_size = src.m_size;
	m_sizeScale = src.m_sizeScale;
	m_strXML = src.m_strXML;
	m_bIsMirror = src.m_bIsMirror;
}
//*******************************************************************************************************
CBCGPRect CBCGPSVGViewBox::GetBounds()
{
	if (m_rectViewBox.IsRectEmpty())
	{
		return CBCGPSVGGroup::GetBounds();
	}

	return m_rectViewBox;
}
//*******************************************************************************************************
BOOL CBCGPSVGViewBox::TranslateCoordinate(double& dblVal, BOOL bIsHorz, BOOL bIsSize)
{
	if (!m_rectViewBox.IsRectEmpty())
	{
		if (bIsHorz)
		{
			dblVal *= m_rectViewBox.Width() / 100;
			if (!bIsSize)
			{
				dblVal += m_rectViewBox.left;
			}
		}
		else
		{
			dblVal *= m_rectViewBox.Height() / 100;
			if (!bIsSize)
			{
				dblVal += m_rectViewBox.top;
			}
		}

		return TRUE;
	}

	if (!m_size.IsEmpty())
	{
		dblVal *= (bIsHorz ? m_size.cx : m_size.cy) / 100;
		return TRUE;
	}

	if (m_pParent != NULL)
	{
		return m_pParent->TranslateCoordinate(dblVal, bIsHorz, bIsSize);
	}
	
	dblVal = 0.0;
	return TRUE;
}
//*******************************************************************************************************
BOOL CBCGPSVGViewBox::CreateData(CBCGPXmlElement& root, BOOL bApplyStyle)
{
	if (CBCGPSVGImageList::m_bDesignMode)
	{
		root.GetXML(m_strXML);
	}

	CString strViewBox = root.GetStringAttribute(_T("viewBox"));
	
	CStringArray arPoints;
	if (PreparePathData(strViewBox, arPoints) == 4)
	{
		double x = StringToNumber(arPoints[0]);
		double y = StringToNumber(arPoints[1]);

		m_rectViewBox = CBCGPRect(
			x,
			y,
			x + StringToNumber(arPoints[2]),
			y + StringToNumber(arPoints[3]));
		
		m_size = m_rectViewBox.Size();
		
		m_bPreserveAspectRatio = TRUE;

		if (root.GetStringAttribute(_T("preserveAspectRatio")) == _T("none"))
		{
			m_bPreserveAspectRatio = FALSE;
		}
	}
	
	m_ptLocation.x = ReadCoordinate(root, _T("x"), TRUE, FALSE);
	m_ptLocation.y = ReadCoordinate(root, _T("y"), TRUE, FALSE);

	m_size.cx = ReadCoordinate(root, _T("width"), TRUE, TRUE, m_size.cx);
	m_size.cy = ReadCoordinate(root, _T("height"), FALSE, TRUE, m_size.cy);

	if (!m_size.IsEmpty() && !m_rectViewBox.IsRectEmpty())
	{
		m_sizeScale.cx = m_size.cx / m_rectViewBox.Width();
		m_sizeScale.cy = m_size.cy / m_rectViewBox.Height();

		if (m_bPreserveAspectRatio)
		{
			double scale = min(m_sizeScale.cx, m_sizeScale.cy);
			m_sizeScale.cx = m_sizeScale.cy = scale;
		}
	}
	
	ReadAttributes(root);
	ReadStyleAttributes(root);

	if (!AddChildren(root))
	{
		return FALSE;
	}

	if (bApplyStyle)
	{
		OnApplyStyle();
	}

	return TRUE;
}

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVG

CBCGPSVG::CBCGPSVG()
{
	m_strDefaultClass = _T("svg");
	m_dblExportScale = 1.0;
	m_bIsAlphaDrawingMode = FALSE;
	m_clrDefault = CBCGPColor::Black;
}
//*******************************************************************************************************
CBCGPSVG::CBCGPSVG(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) : CBCGPSVGViewBox(pParent, elem)
{
	m_dblExportScale = 1.0;
	m_bIsAlphaDrawingMode = FALSE;
	m_clrDefault = CBCGPColor::Black;
}
//*******************************************************************************************************
void CBCGPSVG::AdaptColors(COLORREF clrBase, COLORREF clrTone, BOOL bClampHue /*= FALSE*/)
{
	CBCGPSVGViewBox::AdaptColors(clrBase, clrTone, bClampHue);	
	m_clrDefault.AdaptColors(clrBase, clrTone, bClampHue);
}
//*******************************************************************************************************
const CBCGPTextFormat& CBCGPSVG::GetDefaultTextFormat()
{
	if (m_tfDefault.IsEmpty())
	{
		LOGFONT lf;
		memset(&lf, 0, sizeof(lf));
		
		globalData.fontRegular.GetLogFont(&lf);
		m_tfDefault.CreateFromLogFont(lf);

		m_tfDefault.SetTextAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_LEADING);
		m_tfDefault.SetTextVerticalAlignment(CBCGPTextFormat::BCGP_TEXT_ALIGNMENT_TRAILING);

		m_tfDefault.SetAdjustLineSpacing();
	}

	return m_tfDefault;
}

////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGMarker

CBCGPSVGMarker::CBCGPSVGMarker(CBCGPSVGBase* pParent, CBCGPXmlElement& elem) : CBCGPSVGViewBox(pParent, elem)
{
	m_strDefaultClass = _T("marker");

	m_ptLocation.x = ReadCoordinate(elem, _T("refX"), TRUE, FALSE);
	m_ptLocation.y = ReadCoordinate(elem, _T("refY"), FALSE, FALSE);

	m_size.cx = ReadCoordinate(elem, _T("markerWidth"), TRUE, TRUE);
	m_size.cy = ReadCoordinate(elem, _T("markerHeight"), FALSE, TRUE);
}
//*******************************************************************************************************
void CBCGPSVGMarker::DoDraw(CBCGPGraphicsManager* pGM, const CBCGPPoint& ptMarker, double dblMarkerAngle)
{
	CBCGPPoint pt = ptMarker - m_ptLocation;
	
	CBCGPTransformMatrix transformTranslate;
	transformTranslate.CreateTranslate(pt.x, pt.y);
	m_arTransforms.Add(transformTranslate);
	
	CBCGPTransformMatrix transformRotate;
	transformRotate.CreateRotate(dblMarkerAngle, -DBL_MAX, -DBL_MAX);
	m_arTransforms.Add(transformRotate);
	
	CBCGPTransformMatrix transformTranslate2;
	transformTranslate2.CreateTranslate(-m_size.cx, -m_size.cy);
	m_arTransforms.Add(transformTranslate2);
	
	{
		CBCGPSVGElemTransform transform(pGM, this);
		OnDraw(pGM);
	}
	
	CleanUpTransforms();
}
//*******************************************************************************************************
void CBCGPSVGMarker::CopyFrom(const CBCGPSVGBase& s)
{
	CBCGPSVGViewBox::CopyFrom(s);
	
	const CBCGPSVGMarker& src = (const CBCGPSVGMarker&)s;
	
	m_size = src.m_size;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGImage

CString CBCGPSVGImage::m_strSVGResType = _T("SVG");

CBCGPSVGImage::CBCGPSVGImage()
{
	m_bTraceProblems = TRUE;
	m_bIsSimplified = FALSE;
	m_pGM = NULL;
}
//*******************************************************************************************************
CBCGPSVGImage::CBCGPSVGImage(const CBCGPSVGImage& src)
{
	m_bTraceProblems = TRUE;
	m_bIsSimplified = src.m_bIsSimplified;
	m_pGM = NULL;

	m_SVG.CopyFrom(src.m_SVG);
}
//*******************************************************************************************************
CBCGPSVGImage::~CBCGPSVGImage()
{
}
//*******************************************************************************************************
BOOL CBCGPSVGImage::LoadFromFile(const CString& strFileName)
{
	m_SVG.m_arErrors.RemoveAll();

    CFileException ex;

	BOOL bRes = FALSE;

    CFile file;
    if (!file.Open(strFileName, CFile::modeRead | CFile::typeBinary, &ex))
    {
		CString strErr;
        strErr.Format(_T("CBCGPSVGImage: Cannot open %s file."), (LPCTSTR)strFileName);

	    TCHAR   szErrorMessage[512];
	    UINT    nHelpContext;
        
		if (ex.GetErrorMessage(szErrorMessage, sizeof(szErrorMessage) / sizeof(TCHAR), &nHelpContext))
        {
            CString strReason;
            strReason.Format(_T(" Reason: %s"), szErrorMessage);

            strErr += strReason;
        }

		m_SVG.m_arErrors.Add(strErr);
    }
    else
    {
        const UINT length = (UINT)file.GetLength();
        LPBYTE lpBuffer = new BYTE[length];
        if (lpBuffer == NULL)
        {
		    m_SVG.m_arErrors.Add(_T("General memory error."));
        }
        else
        {
            if (file.Read(lpBuffer, length) == length)
            {
		        m_bTraceProblems = FALSE;
        		
		        bRes = LoadFromBuffer(lpBuffer, length);

		        m_bTraceProblems = TRUE;
            }

            delete [] lpBuffer;
        }
    }

	if (m_SVG.m_arErrors.GetSize() > 0)
	{
		TRACE1("CBCGPSVGImage: file %s has the following problems:\n", (LPCTSTR)strFileName);
		TraceProblems();
	}

	return bRes;
}
//*******************************************************************************************************
BOOL CBCGPSVGImage::LoadFromBuffer(const CString& strSVGBuffer)
{
	if (m_bTraceProblems)
	{
		m_SVG.m_arErrors.RemoveAll();
	}

	BOOL bRes = FALSE;

	CleanUp();
	
	CBCGPXmlDocument doc;
	doc.Create();
	doc.SetProperty(_T("ProhibitDTD"), _variant_t(false));
	doc.SetValidateOnParse(FALSE);
	doc.SetResolveExternals(FALSE);
	
	if (!doc.LoadXML(strSVGBuffer))
	{
		CBCGPXmlParseError err = doc.GetParseError();
		m_SVG.m_arErrors.Add(err.GetReason());
	}
	else
	{
		CBCGPXmlElement root(doc.GetDocumentElement());
		bRes = m_SVG.CreateData(root, TRUE);
	}

	TraceProblems();
	return bRes;
}
//*******************************************************************************************************
BOOL CBCGPSVGImage::LoadFromBuffer(LPBYTE lpBuffer, UINT length)
{
	if (m_bTraceProblems)
	{
		m_SVG.m_arErrors.RemoveAll();
	}

	BOOL bRes = FALSE;

	CleanUp();
	
	CBCGPXmlDocument doc;
	doc.Create();
	doc.SetProperty(_T("ProhibitDTD"), _variant_t(false));
	doc.SetValidateOnParse(FALSE);
	doc.SetResolveExternals(FALSE);
	
	if (!doc.Load(lpBuffer, length))
	{
		CBCGPXmlParseError err = doc.GetParseError();
		m_SVG.m_arErrors.Add(err.GetReason());
	}
	else
	{
		CBCGPXmlElement root(doc.GetDocumentElement());
		bRes = m_SVG.CreateData(root, TRUE);
	}

	TraceProblems();
	return bRes;
}
//*******************************************************************************************************
BOOL CBCGPSVGImage::Load(UINT uiResID, HINSTANCE hinstRes /*= NULL*/)
{
	return LoadStr(MAKEINTRESOURCE(uiResID), hinstRes);
}
//*******************************************************************************************************
BOOL CBCGPSVGImage::LoadStr(LPCTSTR lpszResourceName, HINSTANCE hinstRes /*= NULL*/)
{
	if (hinstRes == NULL)
	{
		hinstRes = AfxFindResourceHandle(lpszResourceName, m_strSVGResType);
	}
	
	HRSRC hRsrc = ::FindResource(hinstRes, lpszResourceName, m_strSVGResType);
	if (hRsrc == NULL)
	{
		return FALSE;
	}
	
	HGLOBAL hGlobal = LoadResource (hinstRes, hRsrc);
	if (hGlobal == NULL)
	{
		return FALSE;
	}
	
	LPVOID lpBuffer = ::LockResource(hGlobal);
	if (lpBuffer == NULL)
	{
		FreeResource(hGlobal);
		return FALSE;
	}
	
	BOOL bRes = LoadFromBuffer((LPBYTE) lpBuffer, (UINT) ::SizeofResource (hinstRes, hRsrc));
	
	UnlockResource(hGlobal);
	FreeResource(hGlobal);
	
	return bRes;
}
//*******************************************************************************************************
void CBCGPSVGImage::CleanUp()
{
	m_Image.Clear();
	m_SVG.CleanUp();
}
//*******************************************************************************************************
void CBCGPSVGImage::DoDraw(CBCGPGraphicsManager* pGM, const CBCGPPoint& pt/* = CBCGPPoint()*/, const CBCGPSize& size/* = CBCGPSize()*/)
{
	m_pGM = pGM;
	
	int nTransformsToRemove = 0;
	
	if (pt != CBCGPPoint())
	{
		CBCGPTransformMatrix transform;
		transform.CreateTranslate(pt.x, pt.y);
		
		m_SVG.m_arTransforms.Add(transform);
		nTransformsToRemove++;
	}
	
	if (!size.IsEmpty())
	{
		const CBCGPSize sizeSVG = GetSize();
		if (!sizeSVG.IsEmpty())
		{
			CBCGPTransformMatrix transform;
			transform.CreateScale(size.cx / sizeSVG.cx, size.cy / sizeSVG.cy);
			
			m_SVG.m_arTransforms.Add(transform);
			nTransformsToRemove++;
		}
	}

	m_SVG.OnBeforeDraw(pGM);
	m_SVG.OnDraw(pGM);

	while (nTransformsToRemove > 0)
	{
		m_SVG.m_arTransforms.RemoveAt((int)m_SVG.m_arTransforms.GetSize() - 1);
		nTransformsToRemove--;
	}
}
//*******************************************************************************************************
void CBCGPSVGImage::DoDraw(CDC* pDC, const CRect& rect, BOOL bDisabled/* = FALSE*/, BYTE alphaSrc/* = 255*/, BOOL bCashBitmap/* = FALSE*/)
{
	ASSERT_VALID(pDC);

	const CSize sizeImage = rect.Size();

	if (m_Image.GetCount() == 0 || sizeImage != m_Image.GetImageSize())
	{
		m_Image.Clear();
		m_Image.SetImageSize(sizeImage);

		HBITMAP hbmp = ExportToBitmap(sizeImage);
		if (hbmp == NULL)
		{
			return;
		}

		m_Image.AddImage(hbmp, TRUE);
#if (!defined _BCGSUITE_)
		m_Image.ConvertToGrayScale(1.0, TRUE);
#endif
		m_Image.AddImage(hbmp, TRUE);
		::DeleteObject(hbmp);
	}

	m_Image.DrawEx(pDC, rect, bDisabled ? 0 : 1, CBCGPToolBarImages::ImageAlignHorzLeft, CBCGPToolBarImages::ImageAlignVertTop, CRect(0, 0, 0, 0), alphaSrc);

	if (!bCashBitmap)
	{
		m_Image.Clear();
	}
}
//*******************************************************************************************************
HBITMAP CBCGPSVGImage::ExportToBitmap(double dblScale /*= 1.0*/, COLORREF clrBackground/* = CLR_NONE*/)
{
	CBCGPSize size = GetSize();
	size.Scale(dblScale);

	size.cx = ceil(size.cx);
	size.cy = ceil(size.cy);
	
	if (size.IsEmpty())
	{
		m_SVG.m_arErrors.Add(_T("CBCGPSVGImage::ExportToBitmap: SVG size is not specified."));
		return NULL;
	}
	
	return ExportToBitmap(size, CBCGPPoint(0, 0), dblScale, clrBackground);
}
//*******************************************************************************************************
HBITMAP CBCGPSVGImage::ExportToBitmap(const CSize& size, COLORREF clrBackground/* = CLR_NONE*/)
{
	CBCGPSize sizeSVG = GetSize();
	if (sizeSVG.IsEmpty())
	{
		m_SVG.m_arErrors.Add(_T("CBCGPSVGImage::ExportToBitmap: SVG size is not specified."));
		return NULL;
	}

	const double dblScale = min((double)size.cx / sizeSVG.cx, (double)size.cy / sizeSVG.cy);

	CBCGPPoint ptOffset(
		max(0, 0.5 * (size.cx - dblScale * sizeSVG.cx)),
		max(0, 0.5 * (size.cy - dblScale * sizeSVG.cy)));

	return ExportToBitmap(size, ptOffset, dblScale, clrBackground);
}
//*******************************************************************************************************
HBITMAP CBCGPSVGImage::ExportToBitmap(const CSize& size, const CBCGPPoint& ptOffset, double dblScale, COLORREF clrBackground/* = CLR_NONE*/)
{
	CBCGPRect rect(ptOffset, size);
	
	HBITMAP hmbpDib = clrBackground == CLR_NONE ? CBCGPDrawManager::CreateBitmap_32(size, NULL) : CBCGPDrawManager::CreateBitmap_24(size, NULL);
	if (hmbpDib == NULL)
	{
		return NULL;
	}
	
	HBITMAP hmbpDibAlpha = NULL;
	
	if (clrBackground == CLR_NONE)
	{
		hmbpDibAlpha = CBCGPDrawManager::CreateBitmap_32(size, NULL);
		if (hmbpDibAlpha == NULL)
		{
			::DeleteObject(hmbpDib);
			return NULL;
		}
	}
	
	CBCGPGraphicsManager* pGM = m_pGM != NULL ? m_pGM : CBCGPGraphicsManager::CreateInstance();
	if (pGM == NULL)
	{
		::DeleteObject(hmbpDib);

		if (hmbpDibAlpha != NULL)
		{
			::DeleteObject(hmbpDibAlpha);
		}

		return NULL;
	}

	m_SVG.OnBeforeDraw(pGM);
	
	CDC dcMem;
	dcMem.CreateCompatibleDC (NULL);

	const int nSteps = clrBackground == CLR_NONE ? 2 : 1;

	for (int step = 0; step < nSteps; step++)
	{
		m_SVG.m_bIsAlphaDrawingMode = (step == 1);
		
		HBITMAP hbmpOld = (HBITMAP)dcMem.SelectObject(m_SVG.m_bIsAlphaDrawingMode ? hmbpDibAlpha : hmbpDib);
		
		pGM->BindDC(&dcMem, rect);
		pGM->BeginDraw();

		if (clrBackground != CLR_NONE)
		{
			pGM->FillRectangle(rect, CBCGPBrush(clrBackground));
		}

		{
			CBCGPRAII<double> scale(m_SVG.m_dblExportScale, dblScale);
			m_SVG.OnDraw(pGM);
		}
		
		pGM->EndDraw();
		
		dcMem.SelectObject (hbmpOld);
	}
	
	m_SVG.m_bIsAlphaDrawingMode = FALSE;
	
	pGM->BindDC(NULL);
	
	if (m_pGM == NULL)
	{
		delete pGM;
	}
	
	if (clrBackground == CLR_NONE)
	{
		CBCGPScanlinerBitmap scan;
		scan.Attach(hmbpDib);
		
		CBCGPScanlinerBitmap scanAlpha;
		scanAlpha.Attach(hmbpDibAlpha);
		
		for (LONG y = 0; y < size.cy; y++)
		{
			LPBYTE pRow = scan.Get();
			LPBYTE pRowAlpha = scanAlpha.Get();
			
			for (LONG x = 0; x < size.cx; x++)
			{
				BYTE a = pRowAlpha[0];

				if (m_bIsSimplified)
				{
					a = pRow[0] == 0 ? 1 : pRow[0];
					pRow[0] = pRow[1] = pRow[2] = a;
				}

				pRow[3] = a;
				
				pRow += 4;
				pRowAlpha += 4;
			}
			
			scan++;
			scanAlpha++;
		}
	}
	
	if (hmbpDibAlpha != NULL)
	{
		::DeleteObject(hmbpDibAlpha);
	}

	return hmbpDib;
}
//*******************************************************************************************************
HICON CBCGPSVGImage::ExportToIcon(const CSize& sizeIcon)
{
	HBITMAP hbmp = ExportToBitmap(sizeIcon);
	if (hbmp == NULL)
	{
		return NULL;
	}

#ifndef _BCGSUITE_
	CBCGPToolBarImages::MultiplyAlpha(hbmp);
#endif		

	CImageList imageList;
	if (!imageList.Create(sizeIcon.cx, sizeIcon.cy, ILC_COLOR32, 1, 0))
	{
		return NULL;
	}
		
	imageList.Add(CBitmap::FromHandle(hbmp), (CBitmap*)NULL);
	::DeleteObject(hbmp);
		
	return imageList.ExtractIcon(0);
}
//*******************************************************************************************************
BOOL CBCGPSVGImage::ExportToFile(const CString& strFilePath, double dblScale /*= 1.0*/)
{
	HBITMAP hbmp = ExportToBitmap(dblScale);
	if (hbmp == NULL)
	{
		return FALSE;
	}
	
	TCHAR ext[_MAX_EXT];
	
#if _MSC_VER < 1400
	_tsplitpath (strFilePath, NULL, NULL, NULL, ext);
#else
	_tsplitpath_s (strFilePath, NULL, 0, NULL, 0, NULL, 0, ext, _MAX_EXT);
#endif
	
	CString strExt = ext;
	strExt.MakeUpper();

	BOOL bRes = FALSE;
	
#if (!defined _BCGSUITE_)
	if (strExt == _T(".PNG"))
	{
		CBCGPPngImage pngImage;
		pngImage.Attach(hbmp);
		
		bRes = pngImage.SaveToFile(strFilePath);
	}
	else
#endif
	{
#ifdef _BCGSUITE_
		class CBCGPToolBarImagesForSave : public CBCGPToolBarImages
		{
		public:
			CBCGPToolBarImagesForSave()
			{
				m_bUserImagesList = TRUE;
			}
		};

		CBCGPToolBarImagesForSave img;
		img.AddImage(hbmp, TRUE);
		img.SetSingleImage();
#else
		CBCGPToolBarImages img;
		img.AddImage(hbmp, TRUE);
		img.SetSingleImage(FALSE);
#endif
		bRes = img.Save(strFilePath);
	}

	::DeleteObject(hbmp);
	return bRes;
}
//*******************************************************************************************************
void CBCGPSVGImage::TraceProblems()
{
	if (!m_bTraceProblems)
	{
		return;
	}

	for (int i = 0; i < m_SVG.m_arErrors.GetSize(); i++)
	{
		TRACE(_T("CBCGPSVGImage: %s\n"), (LPCTSTR)m_SVG.m_arErrors[i]);
	}
}
//*******************************************************************************************************
void CBCGPSVGImage::AdaptColors(COLORREF clrBase, COLORREF clrTone, BOOL bClampHue /*= TRUE*/)
{
	m_SVG.AdaptColors(clrBase, clrTone, bClampHue);
	m_Image.Clear();
}
//*******************************************************************************************************
void CBCGPSVGImage::InvertColors()
{
	m_SVG.InvertColors();
	m_Image.Clear();
}
//*******************************************************************************************************
void CBCGPSVGImage::ConvertToGrayScale(double dblLumRatio /*= 1.0*/)
{
	m_SVG.ConvertToGrayScale(dblLumRatio);
	m_Image.Clear();
}
//*******************************************************************************************************
void CBCGPSVGImage::MakeLighter(double dblRatio /*= .25*/)
{
	m_SVG.MakeLighter(dblRatio);
	m_Image.Clear();
}
//*******************************************************************************************************
void CBCGPSVGImage::MakeDarker(double dblRatio /*= .25*/)
{
	m_SVG.MakeDarker(dblRatio);
	m_Image.Clear();
}
//*******************************************************************************************************
void CBCGPSVGImage::Simplify()
{
	m_SVG.Simplify();
	m_Image.Clear();

	m_bIsSimplified = TRUE;
}
//*******************************************************************************************************
void CBCGPSVGImage::Mirror()
{
	m_SVG.Mirror();
	m_Image.Clear();
}
//*******************************************************************************************************
void CBCGPSVGImage::SetLocation(const CBCGPPoint& pt)
{
	m_SVG.m_ptLocation = pt;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////
// CBCGPSVGImageList

CString CBCGPSVGImageList::m_strSVGListResType = _T("SVGLIST");
BOOL CBCGPSVGImageList::m_bDesignMode = FALSE;

CBCGPSVGImageList::CBCGPSVGImageList(UINT uiResID /*= 0*/, HINSTANCE hinstRes /*= NULL*/)
{
	if (uiResID != 0)
	{
		Load(uiResID, hinstRes);
	}
}
//*******************************************************************************************************
CBCGPSVGImageList::~CBCGPSVGImageList()
{
	RemoveAll();	
}
//*******************************************************************************************************
BOOL CBCGPSVGImageList::Load(UINT uiResID, HINSTANCE hinstRes /*= NULL*/)
{
	return LoadStr(MAKEINTRESOURCE(uiResID), hinstRes);
}
//*******************************************************************************************************
BOOL CBCGPSVGImageList::LoadStr(LPCTSTR lpszResourceName, HINSTANCE hinstRes /*= NULL*/)
{
	RemoveAll();

	BOOL bRes = FALSE;

	// First, try to load SVG sprite:
	HINSTANCE hinstResSVG = hinstRes;
	if (hinstResSVG == NULL)
	{
		hinstResSVG = AfxFindResourceHandle(lpszResourceName, CBCGPSVGImage::m_strSVGResType);
	}
	
	HRSRC hRsrc = ::FindResource(hinstResSVG, lpszResourceName, CBCGPSVGImage::m_strSVGResType);
	if (hRsrc != NULL)
	{
		HGLOBAL hGlobal = LoadResource (hinstResSVG, hRsrc);
		if (hGlobal != NULL)
		{
			LPVOID lpBuffer = ::LockResource(hGlobal);
			if (lpBuffer != NULL)
			{
				bRes = LoadFromSVGSpriteBuffer((LPBYTE)lpBuffer, (UINT) ::SizeofResource (hinstResSVG, hRsrc));
				UnlockResource(hGlobal);
			}

			FreeResource(hGlobal);
		}
	}

	if (bRes)
	{
		return TRUE;
	}
	
	// Load from SVGLIST:
	if (hinstRes == NULL)
	{
		hinstRes = AfxFindResourceHandle(lpszResourceName, m_strSVGListResType);
	}
	
	hRsrc = ::FindResource(hinstRes, lpszResourceName, m_strSVGListResType);
	if (hRsrc == NULL)
	{
		return FALSE;
	}
	
	HGLOBAL hGlobal = LoadResource (hinstRes, hRsrc);
	if (hGlobal == NULL)
	{
		return FALSE;
	}
	
	LPVOID lpBuffer = ::LockResource(hGlobal);
	if (lpBuffer == NULL)
	{
		FreeResource(hGlobal);
		return FALSE;
	}
	
	const int nCount = (int)::SizeofResource (hinstRes, hRsrc);
	CString strCurrRes;

	for (int i = 0; i < nCount; i++)
	{
		TCHAR c = (TCHAR)((LPBYTE)lpBuffer)[i];

		if (IsSpaceChar(c))
		{
			if (!strCurrRes.IsEmpty())
			{
				LoadSVG(strCurrRes, hinstRes);
				strCurrRes.Empty();
			}
		}
		else
		{
			strCurrRes += c;
		}
	}
	
	if (!strCurrRes.IsEmpty())
	{
		LoadSVG(strCurrRes, hinstRes);
	}

	UnlockResource(hGlobal);
	FreeResource(hGlobal);
	
	return GetSize() > 0;
}
//*******************************************************************************************************
BOOL CBCGPSVGImageList::LoadFromFile(const CString& strFileName)
{
	RemoveAll();

    CFileException ex;
	
	BOOL bRes = FALSE;
	
    CFile file;
    if (file.Open(strFileName, CFile::modeRead | CFile::typeBinary, &ex))
    {
        const UINT length = (UINT)file.GetLength();
        LPBYTE lpBuffer = new BYTE[length];
        if (lpBuffer != NULL)
        {
            if (file.Read(lpBuffer, length) == length)
            {
				bRes = LoadFromSVGSpriteBuffer(lpBuffer, length);
            }
			
            delete [] lpBuffer;
        }
    }
	
	return bRes;
}
//*******************************************************************************************************
BOOL CBCGPSVGImageList::LoadFromSVGSprite(const CBCGPSVGImage& sprite)
{
	RemoveAll();

	for (int iSvgIndex = 0; iSvgIndex < (int)sprite.m_SVG.m_arElements.GetSize(); iSvgIndex++)
	{
		CBCGPSVG* pIcon = DYNAMIC_DOWNCAST(CBCGPSVG, sprite.m_SVG.m_arElements[iSvgIndex]);
		if (pIcon != NULL)
		{
			ASSERT_VALID(pIcon);

			CBCGPSVGImage* pSVGImage = new CBCGPSVGImage;
			pSVGImage->m_SVG.CopyFrom(*pIcon);

			CBCGPPoint pt = pSVGImage->m_SVG.m_ptLocation;

			BOOL bAdded = FALSE;

			for (int i = 0; i < (int)GetSize(); i++)
			{
				CBCGPSVGImage* pSVGImagePrev = GetAt(i);
				CBCGPPoint ptPrev = pSVGImagePrev->m_SVG.m_ptLocation;

				if (pt.y < ptPrev.y)
				{
					InsertAt(i, pSVGImage);
					bAdded = TRUE;
					break;
				}

				if (pt.y == ptPrev.y && pt.x < ptPrev.x)
				{
					InsertAt(i, pSVGImage);
					bAdded = TRUE;
					break;
				}
			}

			if (!bAdded)
			{
				Add(pSVGImage);
			}
		}
	}

	int nCount = (int)GetSize();
	if (nCount == 0)
	{
		Add(new CBCGPSVGImage(sprite));
		nCount = 1;
	}
	else
	{
		for (int i = 0; i < nCount; i++)
		{
			CBCGPSVGImage* pSVGImage = GetAt(i);
			pSVGImage->m_SVG.m_ptLocation = CBCGPPoint();
		}
	}

	return nCount > 0;
}
//*******************************************************************************************************
BOOL CBCGPSVGImageList::LoadFromSVGSpriteBuffer(LPBYTE lpBuffer, UINT length)
{
	CBCGPXmlDocument doc;
	doc.Create();
	doc.SetProperty(_T("ProhibitDTD"), _variant_t(false));
	doc.SetValidateOnParse(FALSE);
	doc.SetResolveExternals(FALSE);
	
	if (!doc.Load(lpBuffer, length))
	{
		return FALSE;
	}

	CBCGPXmlElement root(doc.GetDocumentElement());
	CBCGPXmlNodeList elements(root.GetChildren());
	
	for (int i = 0; i < elements.GetLength(); i++)
	{
		CBCGPXmlNode node(elements.GetItem(i));
		if (!node.IsValid())
		{
			continue;
		}
		
		CBCGPXmlElement elem(node);
		if (!elem.IsValid())
		{
			continue;
		}
		
		CString strName = node.GetNodeName();
		if (strName == _T("svg"))
		{
			CBCGPSVGImage* pSVGImage = new CBCGPSVGImage;

			if (pSVGImage->m_SVG.CreateData(elem, TRUE))
			{
				CBCGPPoint pt = pSVGImage->m_SVG.m_ptLocation;
				BOOL bAdded = FALSE;
				
				for (int j = 0; j < (int)GetSize(); j++)
				{
					CBCGPSVGImage* pSVGImagePrev = GetAt(j);
					CBCGPPoint ptPrev = pSVGImagePrev->m_SVG.m_ptLocation;
					
					if (pt.y < ptPrev.y)
					{
						InsertAt(i, pSVGImage);
						bAdded = TRUE;
						break;
					}
					
					if (pt.y == ptPrev.y && pt.x < ptPrev.x)
					{
						InsertAt(i, pSVGImage);
						bAdded = TRUE;
						break;
					}
				}
				
				if (!bAdded)
				{
					Add(pSVGImage);
				}
			}
			else
			{
				delete pSVGImage;
			}
		}
	}
	
	int nCount = (int)GetSize();
	if (nCount == 0)
	{
		CBCGPSVGImage* pSVGImage = new CBCGPSVGImage;
		if (pSVGImage->m_SVG.CreateData(root, TRUE))
		{
			Add(pSVGImage);
			nCount = 1;
		}
		else
		{
			delete pSVGImage;
		}
	}
	else
	{
		for (int i = 0; i < nCount; i++)
		{
			CBCGPSVGImage* pSVGImage = GetAt(i);
			pSVGImage->m_SVG.m_ptLocation = CBCGPPoint();
		}
	}
	
	return GetSize() > 0;
}
//*******************************************************************************************************
void CBCGPSVGImageList::RemoveAll()
{
	for (int i = 0; i < GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = GetAt(i);
		if (pSVG != NULL)
		{
			delete pSVG;
		}
	}

	CArray<CBCGPSVGImage*, CBCGPSVGImage*>::RemoveAll();
}
//*******************************************************************************************************
HBITMAP CBCGPSVGImageList::ExportToBitmap(double dblScale /*= 1.0*/, COLORREF clrBackground/* = CLR_NONE*/)
{
	if (GetSize() == 0 || dblScale <= 0.0)
	{
		return NULL;
	}

	int i = 0;

	// Calculate total size:
	CBCGPSize sizeTotal;

	for (i = 0; i < GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = GetAt(i);
		if (pSVG == NULL)
		{
			ASSERT(FALSE);
			return NULL;
		}

		CBCGPSize sizeCurr = pSVG->GetSize();
		sizeCurr *= dblScale;
		
		sizeTotal.cx += sizeCurr.cx;
		sizeTotal.cy = max(sizeTotal.cy, sizeCurr.cy);
	}

	if (sizeTotal.IsEmpty())
	{
		return NULL;
	}

	CSize sizeTotalGDI;
	sizeTotalGDI.cx = (int)(0.5 + sizeTotal.cx);
	sizeTotalGDI.cy = (int)(0.5 + sizeTotal.cy);

	HBITMAP hBmp = (clrBackground == CLR_NONE) ? CBCGPDrawManager::CreateBitmap_32(sizeTotalGDI, NULL) : CBCGPDrawManager::CreateBitmap_24(sizeTotalGDI, NULL);
	if (hBmp == NULL)
	{
		ASSERT (FALSE);
		return NULL;
	}

	CWindowDC	dc(NULL);
	CDC 		memDCSrc;
	CDC 		memDCDst;
	
	memDCSrc.CreateCompatibleDC (&dc);
	memDCDst.CreateCompatibleDC (&dc);
	
	HBITMAP hOldBitmapDst = (HBITMAP) memDCDst.SelectObject (hBmp);

	int x = 0;

	for (i = 0; i < GetSize(); i++)
	{
		CBCGPSVGImage* pSVG = GetAt(i);

		CBCGPSize sizeCurrSVG = pSVG->GetSize();
		sizeCurrSVG *= dblScale;

		CSize sizeCurr;
		sizeCurr.cx = (int)(0.5 + sizeCurrSVG.cx);
		sizeCurr.cy = (int)(0.5 + sizeCurrSVG.cy);

		HBITMAP hBmpSVG = pSVG->ExportToBitmap(dblScale, clrBackground);
		if (hBmpSVG != NULL)
		{
			HBITMAP hOldBitmapSrc = (HBITMAP) memDCSrc.SelectObject(hBmpSVG);
			memDCDst.BitBlt (x, 0, sizeCurr.cx, sizeCurr.cy, &memDCSrc, 0, 0, SRCCOPY);
			memDCSrc.SelectObject (hOldBitmapSrc);
			
			::DeleteObject(hBmpSVG);
		}

		x += sizeCurr.cx;
	}
	
	memDCDst.SelectObject (hOldBitmapDst);
	return hBmp;
}
//*******************************************************************************************************
BOOL CBCGPSVGImageList::LoadSVG(const CString& strResID, HINSTANCE hinstRes)
{
	CBCGPSVGImage* pSVG = new CBCGPSVGImage();

	UINT uiSVGResID = 0;
	
	for (int i = 0; i < strResID.GetLength(); i++)
	{
		if (!isdigit(strResID[i]))
		{
			break;
		}
		
		if (i == strResID.GetLength() - 1)
		{
			uiSVGResID = _ttol(strResID);
			break;
		}
	}
	
	BOOL bRes = (uiSVGResID != 0) ? pSVG->Load(uiSVGResID, hinstRes) : pSVG->LoadStr(strResID, hinstRes);
	
	if (!bRes)
	{
		ASSERT(FALSE);
		
		delete pSVG;
		pSVG = NULL;
	}
	
	Add(pSVG);
	return bRes;
}
//*******************************************************************************************************
BOOL CBCGPSVGImageList::SaveToFile(const CString& strFileName)
{
	if (!CBCGPSVGImageList::m_bDesignMode)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	int nCount = (int)GetSize();
	if (nCount == 0)
	{
		return FALSE;
	}

	double dblWidth = 0;
	double dblHeight = 0;
	int i = 0;

	for (i = 0; i < nCount; i++)
	{
		CBCGPSVGImage* pSVG = GetAt(i);
		if (pSVG != NULL)
		{
			CBCGPSize size = pSVG->GetSize();

			dblWidth = max(dblWidth, size.cx);
			dblHeight = max(dblHeight, size.cy);
		}
	}

	const int nIconsInRow = nCount < 9 ? nCount : (int)(0.5 + sqrt((double)nCount));
	const int nRows = (int)((double)nCount / nIconsInRow) + 1;

	try
	{
		CFile file;
		file.Open(strFileName, CFile::modeCreate | CFile::modeWrite);

		CBCGPXmlDocument doc;
		CBCGPXmlDocument docSvg;

		if (doc.CreateDocument(CBCGPXmlProcessingInstruction::c_szEncoding_Utf8, _T("svg"), _T("1.0"), FALSE))
		{
			CBCGPXmlElement svg(doc.GetDocumentElement());
			if (svg.IsValid())
			{
				svg.SetAttribute(_T("width"), _variant_t(dblWidth * nIconsInRow));
				svg.SetAttribute(_T("height"), _variant_t(dblHeight * nRows));
				svg.SetAttribute(_T("xmlns"), _variant_t(_T("http://www.w3.org/2000/svg")));

				CBCGPXmlComment comment(doc.CreateComment(_T("This file was automatically generated by BCGSoft ribbon/toolbar designer")));
				if (comment.IsValid())
				{
					svg.AppendChild(comment);
				}

				doc.Format(docSvg);
			}
		}

		if (docSvg.IsValid())
		{
			CBCGPXmlElement svg(docSvg.GetDocumentElement());
			if (svg.IsValid())
			{
				double x = 0.0;
				double y = 0.0;

				int nColumn = 0;
				double dblRowHeight = 0.0;

				for (i = 0; i < nCount; i++)
				{
					CBCGPSVGImage* pSVG = GetAt(i);
					if (pSVG == NULL)
					{
						continue;
					}

					CString strXML(pSVG->GetXML());
					if (strXML.IsEmpty())
					{
						continue;
					}

					CBCGPXmlDocument docElem;
					if (!docElem.LoadXML(strXML))
					{
						continue;
					}

					CBCGPXmlElement elem(docElem.GetDocumentElement());
					if (!elem.IsValid())
					{
						continue;
					}

					elem.SetAttribute(_T("x"), _variant_t(x));
					elem.SetAttribute(_T("y"), _variant_t(y));

					svg.AppendChild(elem);

					nColumn++;

					if (nColumn == nIconsInRow)
					{
						x = 0.0;
						y += dblRowHeight;

						nColumn = 0;
						dblRowHeight = 0.0;
					}
					else
					{
						CBCGPSize size = pSVG->GetSize();

						x += size.cx;
						dblRowHeight = max(dblRowHeight, size.cy);
					}
				}

				docSvg.Save(file);
			}
		}

		file.Close();
	}
	catch (CMemoryException* e)
	{
		e->ReportError();
		return FALSE;
	}
	catch (CFileException* e)
	{
		e->ReportError();
		return FALSE;
	}
	catch (CException* e)
	{
		e->ReportError();
		return FALSE;
	}

	return TRUE;
}

#endif // BCGP_EXCLUDE_SVG
