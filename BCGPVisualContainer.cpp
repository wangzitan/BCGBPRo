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
// BCGPVisualContainer.cpp: implementation of the CBCGPVisualContainer class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BCGPGridCtrl.h"
#include "BCGPMath.h"
#include "BCGPVisualContainer.h"
#include "BCGPVisualCtrl.h"

#ifndef _BCGPCHART_STANDALONE
#include "BCGPVisualCollector.h"
#include "BCGPVisualConstructor.h"
#endif

#if (!defined _BCGSUITE_) && (!defined _BCGPCHART_STANDALONE)
#include "BCGPPngImage.h"
#endif

#include "BCGPDrawManager.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

UINT BCGM_POSTREDRAW = ::RegisterWindowMessage (_T("BCGM_POSTREDRAW"));
UINT BCGM_CONTAINER_SELCHANGED = ::RegisterWindowMessage (_T("BCGM_CONTAINER_SELCHANGED"));
UINT BCGM_CONTAINER_SCALE_CHANGED = ::RegisterWindowMessage (_T("BCGM_CONTAINER_SCALE_CHANGED"));
UINT BCGM_CONTAINER_OBJECT_MOVED = ::RegisterWindowMessage (_T("BCGM_CONTAINER_OBJECT_MOVED"));
UINT BCGM_CONTAINER_BEFORE_OBJECT_MOVE = ::RegisterWindowMessage (_T("BCGM_CONTAINER_BEFORE_OBJECT_MOVE"));

IMPLEMENT_DYNCREATE(CBCGPVisualContainer, CBCGPBaseAccessibleObject)
IMPLEMENT_DYNAMIC(CBCGPVisualDataObject, CObject)

IMPLEMENT_DYNAMIC(CBCGPBaseVisualObject, CBCGPBaseAccessibleObject)

#define SEL_MARKER_SIZE 4

#ifdef _BCGSUITE_
class CBCGPToolBarImagesForSave : public CBCGPToolBarImages
{
public:
	CBCGPToolBarImagesForSave()
	{
		m_bUserImagesList = TRUE;
	}
};
#endif

static BOOL SaveBitmapToFile(HBITMAP hbmp, const CString& strFilePath)
{
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

#if (!defined _BCGSUITE_) && (!defined _BCGPCHART_STANDALONE)
	if (strExt == _T(".PNG"))
	{
		CBCGPPngImage pngImage;
		pngImage.Attach(hbmp);

		return pngImage.SaveToFile(strFilePath);
	}
	else
#endif
	{
#ifdef _BCGSUITE_
		CBCGPToolBarImagesForSave img;
#else
		CBCGPToolBarImages img;
#endif
		img.AddImage(hbmp, TRUE);

#ifndef _BCGSUITE_
		img.SetSingleImage(FALSE);
#else
		img.SetSingleImage();
#endif

		return img.Save(strFilePath);
	}
}

//////////////////////////////////////////////////////////////////////
// CBCGPVisualDataObject

void CBCGPVisualDataObject::OnAnimationValueChanged(double /*dblOldValue*/, double dblNewValue)
{
	m_dblAnimatedValue = dblNewValue;
	
	if (m_pParentVisual != NULL)
	{
		m_pParentVisual->OnAnimation(this);
		m_pParentVisual->Redraw();
	}
}

void CBCGPVisualDataObject::OnAnimationFinished()
{
	if (m_pParentVisual != NULL)
	{
		m_pParentVisual->OnAnimationFinished(this);
		m_pParentVisual->Redraw();
	}
}


//////////////////////////////////////////////////////////////////////
// CBCGPBaseVisualObject

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPBaseVisualObject::CBCGPBaseVisualObject(CBCGPVisualContainer* pContainer)
{
	m_pParentContainer = NULL;

	m_nID = 0;
	m_dwUserData = 0;
	m_DataAnimationType = CBCGPAnimationManager::BCGPANIMATION_Cubic;
	m_bIsVisible = TRUE;
	m_rect.SetRectEmpty();
	m_bIsAutoDestroy = TRUE;
	m_bIsDirty = TRUE;
	m_bIsRedrawEnabled = TRUE;
	m_bCacheImage = FALSE;
	m_bIsSelected = FALSE;
	m_uiEditFlags = 0;
	m_bIsTracked = FALSE;
	m_sizeScaleRatio = CBCGPSize(1., 1.);
	m_rectOriginal = CBCGPRect(-1., -1., 0., 0.);
	m_nClickAndHoldID = 0;
	m_bDontResetRect = FALSE;
	m_bIsInteractiveMode = FALSE;
	m_bIsEnabled = TRUE;
	m_bIsPrinting = FALSE;

	if (pContainer != NULL)
	{
		ASSERT_VALID(pContainer);
		pContainer->Add(this);
	}
	else
	{
		m_pWndOwner = NULL;
	}

#if _MSC_VER < 1300
	m_pStdObject = NULL;
#endif
}
//*******************************************************************************
CBCGPBaseVisualObject::CBCGPBaseVisualObject(const CBCGPBaseVisualObject& src)
{
	m_pParentContainer = NULL;
	m_pWndOwner = NULL;

#if _MSC_VER < 1300
	m_pStdObject = NULL;
#endif

	CopyFrom(src);
}
//*******************************************************************************
CBCGPBaseVisualObject::~CBCGPBaseVisualObject()
{
	RemoveAllData();

#if _MSC_VER < 1300
	if (m_pStdObject != NULL)
	{
		//force disconnect accessibility clients
		::CoDisconnectObject ((IAccessible*)m_pStdObject, NULL);
		m_pStdObject = NULL;
	}
#endif
}
//*******************************************************************************
HRESULT CBCGPBaseVisualObject::get_accName(VARIANT varChild, BSTR *pszName)
{
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		CString strText = GetName();
		*pszName = strText.AllocSysString();

		return S_OK;
	}

	return S_FALSE;
}
//*******************************************************************************
HRESULT CBCGPBaseVisualObject::get_accValue(VARIANT varChild, BSTR *pszValue)
{
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		if (m_arData.GetSize() == 0)
		{
			return S_FALSE;
		}
		
		double dblVal = GetValue();
		if (dblVal == 0.0)
		{
			return S_FALSE;
		}
		
		CString str;
		str.Format(_T("%f"), dblVal);
		*pszValue = str.AllocSysString();
		
		return S_OK;
	}
	
	return S_FALSE;
}
//*******************************************************************************
HRESULT CBCGPBaseVisualObject::get_accDescription(VARIANT /*varChild*/, BSTR* /*pszDescription*/)
{
	return S_FALSE;
}
//*******************************************************************************
HRESULT CBCGPBaseVisualObject::get_accRole(VARIANT varChild, VARIANT *pvarRole)
{
	if (pvarRole == NULL)
	{
		return E_INVALIDARG;
	}
	
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		pvarRole->vt = VT_I4;
		pvarRole->lVal = m_bIsInteractiveMode ? ROLE_SYSTEM_PUSHBUTTON : ROLE_SYSTEM_PANE;
		
		
		return S_OK;
	}
	
	return S_FALSE;
}
//*******************************************************************************
HRESULT CBCGPBaseVisualObject::get_accState(VARIANT varChild, VARIANT *pvarState)
{
	if (pvarState == NULL)
	{
		return E_INVALIDARG;
	}

	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		UINT nState = STATE_SYSTEM_FOCUSABLE | STATE_SYSTEM_SELECTABLE;

		if (!m_bIsInteractiveMode)
		{
			nState |= STATE_SYSTEM_READONLY;
		}

		if (IsSelected())
		{
			nState |= (STATE_SYSTEM_FOCUSED | STATE_SYSTEM_SELECTED);
		}
		
		pvarState->vt = VT_I4;
		pvarState->lVal = nState;

		return S_OK;
	}
	
	return S_FALSE;
}
//*******************************************************************************
HRESULT CBCGPBaseVisualObject::get_accDefaultAction(VARIANT varChild, BSTR* pszDefaultAction)
{
	if (pszDefaultAction == NULL)
	{
		return E_INVALIDARG;
	}
	
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF) && m_bIsInteractiveMode)
	{
		CString str = _T("Open");
		*pszDefaultAction = str.AllocSysString();
		return S_OK;
	}
	
	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPBaseVisualObject::accHitTest(long  /*xLeft*/, long /*yTop*/, VARIANT *pvarChild)
{
	if (!pvarChild)
	{
		return E_INVALIDARG;
	}
	
	pvarChild->vt = VT_I4;
	pvarChild->lVal = CHILDID_SELF;
	
	return S_OK;
}
//****************************************************************************
HRESULT CBCGPBaseVisualObject::accLocation(long *pxLeft, long *pyTop, long *pcxWidth, long *pcyHeight, VARIANT varChild)
{
	if (!pxLeft || !pyTop || !pcxWidth || !pcyHeight || varChild.vt != VT_I4)
	{
		return E_INVALIDARG;
	}
	
	if (m_pWndOwner->GetSafeHwnd() == NULL)
	{
		return S_FALSE;
	}
	
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		CRect rcProp = GetRect();
		m_pWndOwner->ClientToScreen(&rcProp);
		
		*pxLeft = rcProp.left;
		*pyTop = rcProp.top;
		*pcxWidth = rcProp.Width();
		*pcyHeight = rcProp.Height();
		
		return S_OK;
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPBaseVisualObject::get_accParent(IDispatch **ppdispParent)
{
    if (ppdispParent == NULL)
	{
		return E_INVALIDARG;
	}
	
	*ppdispParent = NULL;
	
	if (m_pWndOwner->GetSafeHwnd() != NULL)
	{
		CWnd* pWndParent = m_pWndOwner->GetParent();
		if (pWndParent->GetSafeHwnd() != NULL)
		{
			return AccessibleObjectFromWindow(pWndParent->GetSafeHwnd(), (DWORD)OBJID_CLIENT,
				IID_IAccessible, (void**)ppdispParent);
		}
	}
	
	return S_FALSE;
}

#if _MSC_VER < 1300
IMPLEMENT_OLECREATE(CBCGPBaseVisualObject, "BCGPOleAcc.CBCGPBaseVisualObject", 0xf0345314, 0x6b7d, 0x4f03, 0xbf, 0xb8, 0xb5, 0xaa, 0x36, 0x81, 0x33, 0x8d);
#endif

void CBCGPBaseVisualObject::CopyFrom(const CBCGPBaseVisualObject& src)
{
	ASSERT_VALID(this);

	m_bIsDirty = TRUE;
	m_bIsSelected = FALSE;
	m_bIsTracked = FALSE;
	
	m_nID = src.m_nID;
	m_bIsInteractiveMode = src.m_bIsInteractiveMode;
	m_dwUserData = src.m_dwUserData;
	m_DataAnimationType = src.m_DataAnimationType;
	m_bIsVisible = src.m_bIsVisible;
	m_rect = src.m_rect;
	m_bIsAutoDestroy = src.m_bIsAutoDestroy;
	m_bCacheImage = src.m_bCacheImage;
	m_uiEditFlags = src.m_uiEditFlags;
	m_sizeScaleRatio = src.m_sizeScaleRatio;
	m_rectOriginal = src.m_rectOriginal;
	m_bDontResetRect = FALSE;
	
	RemoveAllData();

	for (int i = 0; i < src.m_arData.GetSize(); i++)
	{
		CBCGPVisualDataObject* pData = src.m_arData[i];
		ASSERT_VALID (pData);
		
		CBCGPVisualDataObject* pCopyData = pData->CreateCopy();
		ASSERT_VALID(pCopyData);
		
		AddData (pCopyData);
	}
}
//*******************************************************************************
void CBCGPBaseVisualObject::OnWndEnabled(BOOL bEnable)
{
	m_bIsEnabled = bEnable;
}
//*******************************************************************************
int CBCGPBaseVisualObject::AddData(CBCGPVisualDataObject* pObject)
{
	if (pObject == NULL)
	{
		return -1;
	}

	pObject->SetParentVisual (this);
	return (int)m_arData.Add(pObject);
}
//*******************************************************************************
void CBCGPBaseVisualObject::RemoveAllData()
{
	for (int i = 0; i < m_arData.GetSize(); i++)
	{
		CBCGPVisualDataObject* pData = m_arData[i];
		ASSERT_VALID(pData);
		
		delete pData;
	}

	m_arData.RemoveAll();
}
//*******************************************************************************
CWnd* CBCGPBaseVisualObject::SetOwner(CWnd* pWndOwner, BOOL bRedraw/* = TRUE*/)
{
	CWnd* pOldOwner = m_pWndOwner;

	SetDirty();

	if (m_pWndOwner != pWndOwner)
	{
		if (pWndOwner->GetSafeHwnd () == NULL)
		{
			bRedraw = FALSE;
		}

		m_pWndOwner = pWndOwner;

		if (bRedraw)
		{
			Redraw ();
		}
	}

	return pOldOwner;
}
//*******************************************************************************
void CBCGPBaseVisualObject::SetRect(const CBCGPRect& rect, BOOL bRedraw)
{
	CBCGPRect rectOld;

	if (m_rect != rect)
	{
		rectOld = m_rect;
		m_rect = rect;

		if (!m_bDontResetRect && m_sizeScaleRatio != CBCGPSize(1., 1.))
		{
			m_rectOriginal = CBCGPRect(-1., -1., 0., 0.);
		}

		OnRectChanged(rectOld);
		SetDirty();
	}

	if (m_bIsSelected)
	{
		SetSelected();
	}

	if (bRedraw)
	{
		if (!rectOld.IsRectEmpty())
		{
			RedrawRect(rectOld);
		}

		Redraw();
	}
}
//*******************************************************************************
BOOL CBCGPBaseVisualObject::ResetRect()
{
	if (m_rectOriginal != CBCGPRect(-1., -1., 0., 0.))
	{
		SetRect(m_rectOriginal);
		return TRUE;
	}

	return FALSE;
}
//*******************************************************************************
BOOL CBCGPBaseVisualObject::ResetOriginalRect()
{
	if (m_sizeScaleRatio == CBCGPSize(1., 1.))
	{
		return FALSE;
	}

	m_rectOriginal = CBCGPRect(-1., -1., 0., 0.);
	return TRUE;
}
//*******************************************************************************
BOOL CBCGPBaseVisualObject::SetValue(double dblVal, int nIndex, UINT uiAnimationTime, BOOL bRedraw)
{
	if (nIndex < 0 || nIndex >= m_arData.GetSize())
	{
		ASSERT(FALSE);
		return FALSE;
	}

	CBCGPVisualDataObject* pData = m_arData[nIndex];
	ASSERT_VALID(pData);

	const double dblRange = GetAnimationRange(nIndex);

	SetDirty();

	if (uiAnimationTime == 0 || dblRange == 0.)
	{
		pData->SetValue(dblVal);

		if (bRedraw)
		{
			Redraw();
		}

		return TRUE;
	}

	double dblOldValue = pData->GetValue();

	pData->SetValue(dblVal);

	if (!pData->StartAnimation(dblOldValue, dblVal, 0.0001 * uiAnimationTime * dblRange, GetDataAnimationType()))
	{
		if (bRedraw)
		{
			Redraw();
		}
	}

	return TRUE;
}
//*******************************************************************************
double CBCGPBaseVisualObject::GetValue(int nIndex) const
{
	if (nIndex < 0 || nIndex >= m_arData.GetSize())
	{
		ASSERT(FALSE);
		return 0.;
	}

	return m_arData[nIndex]->GetValue();
}
// *******************************************************************************
void CBCGPBaseVisualObject::Redraw()
{
	ASSERT_VALID (this);

	if (!m_bIsRedrawEnabled)
	{
		return;
	}

	if (m_pWndOwner->GetSafeHwnd() != NULL)
	{
		if (m_pWndOwner->GetExStyle() & WS_EX_LAYERED)
		{
			CBCGPBaseVisualCtrl* pCtrl = DYNAMIC_DOWNCAST(CBCGPBaseVisualCtrl, m_pWndOwner);
			if (pCtrl != NULL)
			{
				pCtrl->OnDrawLayeredPopup();
				return;
			}
		}

		CBCGPRect rect = m_rect;
		rect.OffsetRect(-m_ptScrollOffset);

		m_pWndOwner->RedrawWindow((CRect)rect, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
	}
}
//*******************************************************************************
void CBCGPBaseVisualObject::RedrawRect(const CBCGPRect& rect)
{
	ASSERT_VALID (this);

	if (rect.IsRectEmpty() || m_pWndOwner->GetSafeHwnd() == NULL)
	{
		return;
	}

	if (m_pWndOwner->GetExStyle() & WS_EX_LAYERED)
	{
		CBCGPBaseVisualCtrl* pCtrl = DYNAMIC_DOWNCAST(CBCGPBaseVisualCtrl, m_pWndOwner);
		if (pCtrl != NULL)
		{
			pCtrl->OnDrawLayeredPopup();
			return;
		}
	}

	CRect rectGDI = rect;
	m_pWndOwner->RedrawWindow(rectGDI, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
}
//*******************************************************************************
void CBCGPBaseVisualObject::Invalidate()
{
	ASSERT_VALID (this);
	InvalidateRect(m_rect);
}
//*******************************************************************************
void CBCGPBaseVisualObject::InvalidateRect(const CBCGPRect& rect)
{
	ASSERT_VALID (this);

	if (rect.IsRectEmpty() || m_pWndOwner->GetSafeHwnd() == NULL)
	{
		return;
	}

	m_pWndOwner->InvalidateRect((CRect&)rect);
}
//*******************************************************************************
void CBCGPBaseVisualObject::UpdateWindow()
{
	ASSERT_VALID (this);
	
	if (m_pWndOwner->GetSafeHwnd() != NULL)
	{
		m_pWndOwner->UpdateWindow();
	}
}
//*******************************************************************************
void CBCGPBaseVisualObject::PostRedraw(BOOL bSetDirty)
{
	ASSERT_VALID (this);

	if (m_pWndOwner->GetSafeHwnd() != NULL && !m_rect.IsRectEmpty())
	{
		m_pWndOwner->PostMessage(BCGM_POSTREDRAW, (WPARAM)bSetDirty);
	}
}
//*******************************************************************************
void CBCGPBaseVisualObject::ResetGestureOptions()
{
	ASSERT_VALID (this);

	if (m_pWndOwner->GetSafeHwnd() != NULL)
	{
		CBCGPGestureConfig gestureConfig;
		if (GetGestureConfig(gestureConfig))
		{
			bcgpGestureManager.SetGestureConfig(m_pWndOwner->GetSafeHwnd(), gestureConfig);
		}
	}
}
//*******************************************************************************
BOOL CBCGPBaseVisualObject::OnSetMouseCursor(const CBCGPPoint& pt)
{
	int nHitTest = HitTest(pt);

	switch (nHitTest)
	{
	case HTTOPLEFT:
	case HTBOTTOMRIGHT:
		SetCursor (AfxGetApp ()->LoadStandardCursor (IDC_SIZENWSE));
		return TRUE;

	case HTTOP:
	case HTBOTTOM:
		SetCursor (AfxGetApp ()->LoadStandardCursor (IDC_SIZENS));
		return TRUE;

	case HTRIGHT:
	case HTLEFT:
		SetCursor (AfxGetApp ()->LoadStandardCursor (IDC_SIZEWE));
		return TRUE;

	case HTTOPRIGHT:
	case HTBOTTOMLEFT:
		SetCursor (AfxGetApp ()->LoadStandardCursor (IDC_SIZENESW));
		return TRUE;

	case HTCAPTION:
		SetCursor (AfxGetApp ()->LoadStandardCursor (IDC_SIZEALL));
		return TRUE;
	}

	return FALSE;
}
//*******************************************************************************
HBITMAP CBCGPBaseVisualObject::ExportToBitmap(CBCGPGraphicsManager* pGM)
{
	ASSERT_VALID(pGM);

	const CBCGPSize size = m_rect.Size();

	HBITMAP hmbpDib = CBCGPDrawManager::CreateBitmap_32(size, NULL);
	if (hmbpDib == NULL)
	{
		ASSERT (FALSE);
		return NULL;
	}

	CBCGPRect rect(CBCGPPoint(), size);

	CDC dcMem;
	dcMem.CreateCompatibleDC (NULL);

	HBITMAP hbmpOld = (HBITMAP) dcMem.SelectObject (hmbpDib);

	pGM->BindDC(&dcMem, rect);
	pGM->BeginDraw();

	CBCGPRect rectSaved = m_rect;

	SetRect(m_rect - rectSaved.TopLeft());
	SetDirty();

	BOOL bCacheImage = m_bCacheImage;
	m_bCacheImage = FALSE;

	pGM->FillRectangle(m_rect, CBCGPBrush(CBCGPColor::White));

	OnDraw(pGM, m_rect);

	pGM->EndDraw();

	dcMem.SelectObject (hbmpOld);

	CBCGPDrawManager::FillAlpha (rect, hmbpDib, 255);

	pGM->BindDC(NULL);

	SetRect(rectSaved);
	m_bCacheImage = bCacheImage;

	SetDirty();

	return hmbpDib;
}
//*******************************************************************************
BOOL CBCGPBaseVisualObject::ExportToImage(CBCGPGraphicsManager* pGMSrc, CBCGPImage& image, const CBCGPBrush& brFill)
{
	ASSERT_VALID(pGMSrc);

	CBCGPGraphicsManager* pGM = pGMSrc->CreateOffScreenManager(m_rect, &image);

	if (pGM == NULL)
	{
		return FALSE;
	}

	CBCGPRect rectSaved = m_rect;
	m_rect = m_rect - rectSaved.TopLeft();

	if (!brFill.IsEmpty())
	{
		pGM->FillRectangle(m_rect, brFill);
	}

	SetDirty();
	OnDraw(pGM, CBCGPRect());

	m_rect = rectSaved;
	delete pGM;

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPBaseVisualObject::CopyToClipboard(CBCGPGraphicsManager* pGM)
{
	if (pGM == NULL)
	{
		return FALSE;
	}

	HBITMAP hbmp = ExportToBitmap(pGM);

	CBCGPToolBarImages img;
	img.AddImage(hbmp, TRUE);

#ifndef _BCGSUITE_
	img.SetSingleImage(FALSE);
#else
	img.SetSingleImage();
#endif

	return img.CopyImageToClipboard(0);
}
//*******************************************************************************
BOOL CBCGPBaseVisualObject::ExportToFile(const CString& strFilePath, CBCGPGraphicsManager* pGM)
{
	if (pGM == NULL)
	{
		return FALSE;
	}

	return SaveBitmapToFile(ExportToBitmap(pGM), strFilePath);
}
//*******************************************************************************
void CBCGPBaseVisualObject::EnableImageCache(BOOL bEnable)
{
	m_bCacheImage = bEnable;
}
//*******************************************************************************
CBCGPRect CBCGPBaseVisualObject::MakeTrackMarker(double x, double y) const
{
	CBCGPRect rect(x, y, x, y);
	rect.InflateRect(SEL_MARKER_SIZE, SEL_MARKER_SIZE);

	return rect;
}
//*******************************************************************************
CBCGPRect CBCGPBaseVisualObject::MakeTrackMarker(const CBCGPPoint& pt) const
{
	return MakeTrackMarker(pt.x, pt.y);
}
//*******************************************************************************
void CBCGPBaseVisualObject::SetSelected(BOOL bSet)
{
	if (!IsEditMode())
	{
		ASSERT(FALSE);
		return;
	}

	if (!bSet)
	{
		m_rectTrack.SetRectEmpty();
		m_bIsSelected = FALSE;
		m_mapTrackRects.RemoveAll();
		return;
	}

	SetTrackingRect(m_rect);
	m_bIsSelected = TRUE;
}
//*******************************************************************************
void CBCGPBaseVisualObject::SetTrackingRect(const CBCGPRect& rect)
{
	const CBCGPSize sizeMin = GetMinSize();

	m_mapTrackRects.RemoveAll();

	if (!sizeMin.IsEmpty())
	{
		if (fabs(rect.Width()) < sizeMin.cx || fabs(rect.Height()) < sizeMin.cy)
		{
			return;
		}
	}

	m_rectTrack = rect;

	if ((m_uiEditFlags & BCGP_EDIT_NOSIZE) == BCGP_EDIT_NOSIZE)
	{
		return;
	}

	if ((m_uiEditFlags & (BCGP_EDIT_SIZE_NO_LEFT | BCGP_EDIT_SIZE_NO_TOP)) == 0)
	{
		m_mapTrackRects[HTTOPLEFT] = MakeTrackMarker(m_rectTrack.TopLeft());
	}

	if ((m_uiEditFlags & (BCGP_EDIT_SIZE_NO_RIGHT | BCGP_EDIT_SIZE_NO_BOTTOM)) == 0)
	{
		m_mapTrackRects[HTBOTTOMRIGHT] = MakeTrackMarker(m_rectTrack.BottomRight());
	}

	if ((m_uiEditFlags & (BCGP_EDIT_SIZE_NO_RIGHT | BCGP_EDIT_SIZE_NO_TOP)) == 0)
	{
		m_mapTrackRects[HTTOPRIGHT] = MakeTrackMarker(m_rectTrack.right, m_rectTrack.top);
	}

	if ((m_uiEditFlags & (BCGP_EDIT_SIZE_NO_LEFT | BCGP_EDIT_SIZE_NO_BOTTOM)) == 0)
	{
		m_mapTrackRects[HTBOTTOMLEFT] = MakeTrackMarker(m_rectTrack.left, m_rectTrack.bottom);
	}

	if ((m_uiEditFlags & BCGP_EDIT_SIZE_NO_LEFT) == 0)
	{
		m_mapTrackRects[HTLEFT] = MakeTrackMarker(m_rectTrack.left, m_rectTrack.CenterPoint().y);
	}

	if ((m_uiEditFlags & BCGP_EDIT_SIZE_NO_RIGHT) == 0)
	{
		m_mapTrackRects[HTRIGHT] = MakeTrackMarker(m_rectTrack.right, m_rectTrack.CenterPoint().y);
	}

	if ((m_uiEditFlags & BCGP_EDIT_SIZE_NO_TOP) == 0)
	{
		m_mapTrackRects[HTTOP] = MakeTrackMarker(m_rectTrack.CenterPoint().x, m_rectTrack.top);
	}

	if ((m_uiEditFlags & BCGP_EDIT_SIZE_NO_BOTTOM) == 0)
	{
		m_mapTrackRects[HTBOTTOM] = MakeTrackMarker(m_rectTrack.CenterPoint().x, m_rectTrack.bottom);
	}
}
//*******************************************************************************
BOOL CBCGPBaseVisualObject::IsEditMode() const
{
	return m_pParentContainer != NULL && m_pParentContainer->IsEditMode();
}
//*******************************************************************************
int CBCGPBaseVisualObject::HitTest(const CBCGPPoint& pt) const
{
	if (m_bIsSelected)
	{
		for (POSITION pos = m_mapTrackRects.GetStartPosition (); pos != NULL;)
		{
			UINT uiHitTest = 0;
			CBCGPRect rect;

			m_mapTrackRects.GetNextAssoc (pos, uiHitTest, rect);
			rect.OffsetRect(-m_ptScrollOffset);

			if (rect.PtInRect(pt))
			{
				return uiHitTest;
			}
		}
	}

	CBCGPRect rect = m_rect;
	rect.Normalize();

	rect.OffsetRect(-m_ptScrollOffset);

	if (!rect.PtInRect(pt))
	{
		return HTNOWHERE;
	}

	if (IsEditMode() && (m_uiEditFlags & BCGP_EDIT_NOMOVE) == 0)
	{
		return HTCAPTION;
	}

	return  HTCLIENT;
}
//*******************************************************************************
void CBCGPBaseVisualObject::DrawTrackingRect(CBCGPGraphicsManager* pGM,
											 const CBCGPBrush& brOutline, const CBCGPBrush& brFill)
{
	if (m_rectTrack.IsRectEmpty())
	{
		return;
	}

	CBCGPRect rectTrack = m_rectTrack;
	rectTrack.OffsetRect(-m_ptScrollOffset);

	pGM->DrawRectangle(rectTrack, brOutline);

	for (POSITION pos = m_mapTrackRects.GetStartPosition (); pos != NULL;)
	{
		UINT uiHitTest = 0;
		CBCGPRect rect;

		m_mapTrackRects.GetNextAssoc (pos, uiHitTest, rect);

		rect.OffsetRect(-m_ptScrollOffset);

		switch (uiHitTest)
		{
		case HTTOPLEFT:
		case HTBOTTOMRIGHT:
		case HTTOPRIGHT:
		case HTBOTTOMLEFT:
			pGM->FillEllipse(rect, brFill);
			pGM->DrawEllipse(rect, brOutline);
			break;

		default:
			pGM->FillRectangle(rect, brFill);
			pGM->DrawRectangle(rect, brOutline);
		}
	}
}
//*******************************************************************************
void CBCGPBaseVisualObject::OnBeforeDrawScrolled (const CBCGPPoint& ptDrawOffset)
{
	m_ptDrawOffset = ptDrawOffset;
}
//*******************************************************************************
void CBCGPBaseVisualObject::OnAfterDrawScrolled ()
{
	m_ptDrawOffset = CBCGPPoint(0, 0);
}
//*******************************************************************************
CBCGPBaseVisualObject* CBCGPBaseVisualObject::CreateCopy () const
{
	return NULL;
}
//*******************************************************************************
void CBCGPBaseVisualObject::SetClickAndHoldEvent(const CBCGPRect& rectBounds, UINT nEventID)
{
	m_rectClickAndHold = rectBounds;
	m_nClickAndHoldID = nEventID;
}

#ifndef _BCGPCHART_STANDALONE

//////////////////////////////////////////////////////////////////////
// CBCGPVisualLayout

void CBCGPVisualLayout::AdjustLayout ()
{
	const int count = (int)m_listWnd.GetCount ();
	if (count == 0)
	{
		return;
	}

	CSize szMin(GetMinSize());
	if (szMin != CSize(0, 0))
	{
		CRect rtCur;
		GetHostWndRect(rtCur);
		if (rtCur.Width() < szMin.cx || rtCur.Height() < szMin.cy)
		{
			return;
		}
	}

	POSITION pos = m_listWnd.GetHeadPosition ();
	while (pos != NULL)
	{
		XWndItem& item = m_listWnd.GetNext (pos);

		int nIndex = m_container.FindIndex((const CBCGPBaseVisualObject*)item.m_Handle);
		if (nIndex >= 0)
		{
			CRect rectItem;
			UINT uiFlags = CalculateItem (item, rectItem);

			if ((uiFlags & (SWP_NOMOVE | SWP_NOSIZE)) != (SWP_NOMOVE | SWP_NOSIZE))
			{
				CBCGPBaseVisualObject* pObject = (CBCGPBaseVisualObject*)m_container[nIndex];
				pObject->SetRect(rectItem);
			}
		}
	}
}
//*******************************************************************************
CRect CBCGPVisualLayout::GetItemRect(XWndItem& item) const
{
	int nIndex = m_container.FindIndex((const CBCGPBaseVisualObject*)item.m_Handle);
	if (nIndex >= 0)
	{
		return m_container[nIndex]->GetRect();
	}

	return CRect (0, 0, 0, 0);
}
//********************************************************************************
BOOL CBCGPVisualLayout::AddAnchor(CBCGPBaseVisualObject* pObject, XMoveType typeMove, XSizeType typeSize, 
								  const CPoint& percMove, const CPoint& percSize)
{
	if (pObject == NULL || m_container.FindIndex(pObject) < 0)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	return CBCGPStaticLayout::AddAnchor ((LPVOID)pObject, typeMove, typeSize, percMove, percSize);
}
//********************************************************************************
void CBCGPVisualLayout::GetHostWndRect(CRect& rect) const
{
	rect = m_container.GetRect();

	CSize szMin(GetMinSize());
	if (szMin != CSize(0, 0))
	{
		if (rect.Width() < szMin.cx)
		{
			rect.right = rect.left + szMin.cx;
		}

		if (rect.Height() < szMin.cy)
		{
			rect.bottom = rect.top + szMin.cy;
		}
	}
}

#endif

//////////////////////////////////////////////////////////////////////
// CBCGPVisualContainer

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBCGPVisualContainer::CBCGPVisualContainer(CWnd* pWndOwner)
{
#ifndef _BCGPCHART_STANDALONE
	m_pLayout = NULL;
#endif

	m_bDrawDynamicObjectsOnTop = TRUE;
	m_pWndOwner = pWndOwner;
	m_rect.SetRectEmpty();
	m_bCacheImage = FALSE;
	m_bIsEditMode = FALSE;
	m_bScrollBars = FALSE;
	m_dwEditFlags = 0;
	m_nDragMode = HTNOWHERE;
	m_bDrawScrollBars = TRUE;
	m_bNCScrollBars = FALSE;
	m_bIsPrinting = FALSE;

	m_bHitTestSelected = FALSE;
	m_ptClick = m_ptDragStart = m_ptDragFinish = CBCGPPoint(-1, -1);
	m_szDrag = CBCGPSize(::GetSystemMetrics(SM_CXDRAG), ::GetSystemMetrics(SM_CYDRAG));

	m_brSelFill = CBCGPBrush(CBCGPColor::LightBlue, .5);
	m_brSelOutline = CBCGPBrush(CBCGPColor::SteelBlue);
	m_brGrid = CBCGPBrush(CBCGPColor::WhiteSmoke);
	m_brGridDots = CBCGPBrush(CBCGPColor::CornflowerBlue);

	m_sizeGrid = CBCGPSize(10., 10.);

	m_pNewObject = NULL;
	m_bAddNewObjectMode = FALSE;

	m_GMType = (CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER)-1;

	m_sizeScaleRatio = CBCGPSize(1., 1.);

	m_nClickAndHoldID = 0;

	m_ScrollBarVert.SetHorizontal(FALSE);
	m_ScrollBarHorz.SetHorizontal(TRUE);

	m_bScaleByMouseWheel = FALSE;
	m_dblMinScaleRatio = 0.01;
	m_dblMaxScaleRatio = 4.0;

	m_styleGrid = BCGP_VISUAL_CONTAINER_GRID_LINES;

#if _MSC_VER < 1300
	m_pStdObject = NULL;
#endif
}
//*******************************************************************************
CBCGPVisualContainer::~CBCGPVisualContainer()
{
#ifndef _BCGPCHART_STANDALONE
	if (m_pLayout != NULL)
	{
		delete m_pLayout;
		m_pLayout = NULL;
	}
#endif

	RemoveAll();

#if _MSC_VER < 1300
	if (m_pStdObject != NULL)
	{
		//force disconnect accessibility clients
		::CoDisconnectObject ((IAccessible*)m_pStdObject, NULL);
		m_pStdObject = NULL;
	}
#endif
}

#if _MSC_VER < 1300
IMPLEMENT_OLECREATE(CBCGPVisualContainer, "BCGPOleAcc.CBCGPVisualContainer", 0xea74083, 0x4c72, 0x4a41, 0x86, 0xcc, 0x8b, 0x5a, 0xa7, 0xb8, 0x89, 0x8e);
#endif

CWnd* CBCGPVisualContainer::SetOwner(CWnd* pWndOwner, BOOL bRedraw/* = TRUE*/)
{
	CWnd* pOldOwner = m_pWndOwner;

	if (m_pWndOwner != pWndOwner)
	{
		m_pWndOwner = pWndOwner;

		for (int i = 0; i < m_arObjects.GetSize(); i++)
		{
			CBCGPBaseVisualObject* pObject = m_arObjects[i];
			if (pObject != NULL)
			{
				ASSERT_VALID(pObject);		
				pObject->SetOwner(m_pWndOwner, FALSE);
			}
		}

		if (bRedraw && m_pWndOwner->GetSafeHwnd () != NULL)
		{
			Redraw ();
		}
	}

	return pOldOwner;
}
//*******************************************************************************
void CBCGPVisualContainer::EnableImageCache(BOOL bEnable)
{
	m_bCacheImage = bEnable;

	if (!m_bCacheImage)
	{
		m_ImageCache.Destroy();
	}
}
//*******************************************************************************
void CBCGPVisualContainer::EnableScrollBars(BOOL bEnable, CBCGPVisualScrollBarColorTheme* pColorTheme, CBCGPVisualScrollBar::BCGP_VISUAL_SCROLLBAR_STYLE style)
{
	m_bScrollBars = bEnable;

	if (bEnable)
	{
		if (pColorTheme != NULL)
		{
			m_ScrollBarVert.SetColorTheme(*pColorTheme);
			m_ScrollBarHorz.SetColorTheme(*pColorTheme);
		}

		m_ScrollBarVert.SetStyle(style);
		m_ScrollBarHorz.SetStyle(style);
	}

	AdjustLayout();
}
//*******************************************************************************
BOOL CBCGPVisualContainer::Add(CBCGPBaseVisualObject* pObject, BOOL bAutoDestroy)
{
	ASSERT_VALID(pObject);
	return InsertAt(pObject, -1, bAutoDestroy);
}
//*******************************************************************************
BOOL CBCGPVisualContainer::InsertAt(CBCGPBaseVisualObject* pObject, int nIndex, BOOL bAutoDestroy)
{
	ASSERT_VALID(pObject);
	ASSERT(pObject->m_pParentContainer == NULL);

	if (nIndex == -1)
	{
#ifndef _BCGPCHART_STANDALONE
		if (pObject->IsKindOf (RUNTIME_CLASS(CBCGPDiagramConnector)))
		{
			nIndex = 0;
		}
		else
#endif
		{
			nIndex = (int)m_arObjects.GetSize();
		}
	}

	if (nIndex < 0 || nIndex > m_arObjects.GetSize())
	{
		return FALSE;
	}

	if (FindIndex(pObject) < 0)
	{
		ASSERT(pObject->m_pParentContainer == NULL);
		m_arObjects.InsertAt(nIndex, pObject);
	}

	pObject->m_pParentContainer = this;
	pObject->m_pWndOwner = m_pWndOwner;
	pObject->m_bIsAutoDestroy = bAutoDestroy;
	pObject->SetScrollOffset(m_ptScrollOffset);

	OnAdd (pObject);

	pObject->OnAddToContainer(this);

	AdjustLayout();
	return TRUE;
}
//*******************************************************************************
void CBCGPVisualContainer::OnAdd (CBCGPBaseVisualObject* pObject)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pObject);

	CBCGPSize sizeScaleRatioOld = pObject->GetScaleRatio();

	pObject->SetScaleRatio(GetScaleRatio());

	if ((m_sizeScaleRatio != CBCGPSize(1., 1.) || !pObject->ResetRect()))
	{
		const double dblAspectRatioX = m_sizeScaleRatio.cx / sizeScaleRatioOld.cx;
		const double dblAspectRatioY = m_sizeScaleRatio.cy / sizeScaleRatioOld.cy;
	
		BOOL bDontResetRectSaved = pObject->m_bDontResetRect;
		pObject->m_bDontResetRect = TRUE;
			
		CBCGPRect rect = pObject->GetRect();

		if (rect.IsRectEmpty())
		{
			return;
		}

		rect.Scale (dblAspectRatioX, dblAspectRatioY, rect.TopLeft());
			
		pObject->SetRect(rect);
			
		pObject->m_bDontResetRect = bDontResetRectSaved;
	}
}
//*******************************************************************************
BOOL CBCGPVisualContainer::Remove(int nIndex)
{
	if (nIndex < 0 || nIndex >= m_arObjects.GetSize())
	{
		return FALSE;
	}

	CBCGPBaseVisualObject* pObject = m_arObjects[nIndex];
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		
		OnRemove (pObject);
		
		if (pObject->m_bIsAutoDestroy)
		{
			delete pObject;
		}
	}

	m_arObjects.RemoveAt(nIndex);
	AdjustLayout();

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnCopyObject(CBCGPBaseVisualObject* pObject, const CBCGPRect& rectNew)
{
	ASSERT_VALID(pObject);

	CBCGPBaseVisualObject* pNewVisualObject = pObject->CreateCopy();
	if (pNewVisualObject == NULL)
	{
		pNewVisualObject = DYNAMIC_DOWNCAST(CBCGPBaseVisualObject, pObject->GetRuntimeClass()->CreateObject());
		
		if (pNewVisualObject == NULL)
		{
			ASSERT(FALSE);
			return FALSE;
		}
	}

	pNewVisualObject->SetRect(rectNew);

	Add(pNewVisualObject, TRUE);
	return TRUE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::Remove(CBCGPBaseVisualObject* pObject)
{
	int nIndex = FindIndex(pObject);
	if (nIndex == -1)
	{
		return FALSE;
	}

	return Remove(nIndex);
}
//*******************************************************************************
void CBCGPVisualContainer::OnRemove (CBCGPBaseVisualObject* pObject)
{
	ASSERT_VALID(pObject);
	
	POSITION pos = m_lstSel.Find(pObject);
	if (pos != NULL)
	{
		m_lstSel.RemoveAt(pos);
	}
}
//*******************************************************************************
void CBCGPVisualContainer::RemoveAll()
{
	m_lstSel.RemoveAll();

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);
			
			if (pObject->m_bIsAutoDestroy)
			{
				delete pObject;
			}
			else
			{
				pObject->m_pParentContainer = NULL;
			}

			m_arObjects[i] = NULL;
		}
	}

	m_arObjects.RemoveAll();
	AdjustLayout();
}
//*******************************************************************************
void CBCGPVisualContainer::RemoveSelected()
{
	for (POSITION pos = m_lstSel.GetHeadPosition(); pos != NULL;)
	{
		CBCGPBaseVisualObject* pObject = m_lstSel.GetNext(pos);
		ASSERT_VALID(pObject);

		int nIndex = FindIndex(pObject);
		if (nIndex == -1)
		{
			ASSERT(FALSE);
			continue;
		}

		m_arObjects.RemoveAt(nIndex);

		if (pObject->m_bIsAutoDestroy)
		{
			delete pObject;
		}
	}

	m_lstSel.RemoveAll();
	AdjustLayout();
}
//*******************************************************************************
int CBCGPVisualContainer::GetCount() const
{
	return (int)m_arObjects.GetSize();
}
//*******************************************************************************
CBCGPBaseVisualObject* CBCGPVisualContainer::GetAt(int nIndex)
{
	if (nIndex < 0 || nIndex >= m_arObjects.GetSize())
	{
		return NULL;
	}

	CBCGPBaseVisualObject* pObject = m_arObjects.GetAt(nIndex);

	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
	}

	return pObject;
}
//*******************************************************************************
CBCGPBaseVisualObject* CBCGPVisualContainer::GetByID(UINT nID)
{
	for (int i = 0; i < (int)m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects.GetAt(i);
		ASSERT_VALID(pObject);

		if (pObject->m_nID == nID)
		{
			return pObject;
		}
	}

	return NULL;
}
//*******************************************************************************
const CBCGPBaseVisualObject* CBCGPVisualContainer::operator[](int nIndex) const
{
	if (nIndex < 0 || nIndex >= m_arObjects.GetSize())
	{
		return NULL;
	}

	const CBCGPBaseVisualObject* pObject = m_arObjects[nIndex];

	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
	}

	return pObject;
}
//*******************************************************************************
CBCGPBaseVisualObject* CBCGPVisualContainer::GetFromPoint(const CPoint& point)
{
	if (!m_bNCScrollBars)
	{
		if (m_ScrollBarVert.GetRect().PtInRect(point) ||
			m_ScrollBarHorz.GetRect().PtInRect(point) ||
			m_rectBottomRight.PtInRect(point))
		{
			return NULL;
		}
	}

	if (m_bHitTestSelected)
	{
		for (POSITION pos = m_lstSel.GetHeadPosition(); pos != NULL;)
		{
			CBCGPBaseVisualObject* pObject = m_lstSel.GetNext(pos);
			ASSERT_VALID(pObject);

			if (pObject->HitTest(point) != HTNOWHERE)
			{
				return pObject;
			}
		}
	}

	int count = (int)m_arObjects.GetSize ();

	for (int i = 0; i < count; i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[count - i - 1];
		if (pObject == NULL)
		{
			continue;
		}

		ASSERT_VALID(pObject);

		if (pObject->IsVisible())
		{
			CBCGPRect rect = pObject->m_rect;
			rect.Normalize();

			rect.OffsetRect(-m_ptScrollOffset);

			BOOL bCheckForObjectRect = pObject->m_mapTrackRects.IsEmpty();

			if ((!bCheckForObjectRect || rect.PtInRect(point)) && pObject->HitTest (point) != HTNOWHERE)
			{
				return pObject;
			}
		}
	}

	return NULL;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::MoveForward(CBCGPBaseVisualObject* pObject)
{
	ASSERT_VALID(pObject);

	int nIndex = FindIndex(pObject);
	if (nIndex < 0)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	if (nIndex == m_arObjects.GetSize() - 1)
	{
		return FALSE;
	}

	CBCGPBaseVisualObject* pObjectNext = m_arObjects[nIndex + 1];

	m_arObjects[nIndex + 1] = pObject;
	m_arObjects[nIndex] = pObjectNext;

	m_ImageCache.Destroy();
	return TRUE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::MoveBackward(CBCGPBaseVisualObject* pObject)
{
	ASSERT_VALID(pObject);

	int nIndex = FindIndex(pObject);
	if (nIndex < 0)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	if (nIndex == 0)
	{
		return FALSE;
	}

	CBCGPBaseVisualObject* pObjectPrev = m_arObjects[nIndex - 1];

	m_arObjects[nIndex - 1] = pObject;
	m_arObjects[nIndex] = pObjectPrev;

	m_ImageCache.Destroy();
	return TRUE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::MoveToFront(CBCGPBaseVisualObject* pObject)
{
	ASSERT_VALID(pObject);

	int nIndex = FindIndex(pObject);
	if (nIndex < 0)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	if (nIndex == m_arObjects.GetSize() - 1)
	{
		return FALSE;
	}

	for (int i = nIndex; i < m_arObjects.GetSize() - 1; i++)
	{
		m_arObjects[i] = m_arObjects[i + 1];
	}

	m_arObjects[m_arObjects.GetSize() - 1] = pObject;

	m_ImageCache.Destroy();
	return TRUE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::MoveToBack(CBCGPBaseVisualObject* pObject)
{
	ASSERT_VALID(pObject);

	int nIndex = FindIndex(pObject);
	if (nIndex < 0)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	if (nIndex == 0)
	{
		return FALSE;
	}

	for (int i = nIndex; i > 0; i--)
	{
		m_arObjects[i] = m_arObjects[i - 1];
	}

	m_arObjects[0] = pObject;

	m_ImageCache.Destroy();
	return TRUE;
}
//*******************************************************************************
int CBCGPVisualContainer::FindIndex(const CBCGPBaseVisualObject* pObject) const
{
	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pArObject = m_arObjects[i];
		if (pArObject == NULL)
		{
			continue;
		}

		ASSERT_VALID(pArObject);
		
		if (pArObject == pObject)
		{
			return i;
		}
	}

	return -1;
}
//*******************************************************************************
CBCGPRect CBCGPVisualContainer::GetBoundsRect(BOOL bOnlyVisible) const
{
	CBCGPRect rect(0, 0, 0, 0);

	if (bOnlyVisible)
	{
		for (int i = 0; i < m_arObjects.GetSize(); i++)
		{
			CBCGPBaseVisualObject* pObject = m_arObjects[i];
			if (pObject != NULL)
			{
				ASSERT_VALID(pObject);
				
				if (pObject->IsVisible ())
				{
					rect.UnionRect(rect, pObject->m_rect);
				}
			}
		}
	}
	else
	{
		for (int i = 0; i < m_arObjects.GetSize(); i++)
		{
			CBCGPBaseVisualObject* pObject = m_arObjects[i];
			if (pObject != NULL)
			{
				ASSERT_VALID(pObject);
				
				rect.UnionRect(rect, pObject->m_rect);
			}
		}
	}

	return rect;
}
//*******************************************************************************
void CBCGPVisualContainer::SetRect(const CBCGPRect& rect, BOOL bReposObjects, BOOL bRedraw)
{
	CBCGPRect rectOld = m_rect;

	m_rect = rect;

	if (m_rect == rectOld)
	{
		return;
	}

	if (bReposObjects)
	{
		double dx = m_rect.left;
		double dy = m_rect.top;

		if (!rectOld.IsRectEmpty())
		{
			dx -= rectOld.left;
			dy -= rectOld.top;
		}

		for (int i = 0; i < m_arObjects.GetSize(); i++)
		{
			CBCGPBaseVisualObject* pObject = m_arObjects[i];
			if (pObject != NULL)
			{
				ASSERT_VALID(pObject);
				
				pObject->m_rect.OffsetRect(dx, dy);
			}
		}
	}

	AdjustLayout();

	if (bRedraw)
	{
		Redraw();
	}
}
//*******************************************************************************
void CBCGPVisualContainer::EnableEditMode(BOOL bEnable, DWORD dwFlags)
{
	m_bIsEditMode = bEnable;
	m_dwEditFlags = dwFlags;

	if ((m_dwEditFlags & BCGP_CONTAINER_GRID) == BCGP_CONTAINER_GRID &&
		(m_dwEditFlags & BCGP_CONTAINER_SNAP_TO_GRID) == BCGP_CONTAINER_SNAP_TO_GRID)
	{
		CBCGPSize sizeGrid(m_sizeGrid.cx * m_sizeScaleRatio.cx, m_sizeGrid.cy * m_sizeScaleRatio.cy);
		m_szDrag = CBCGPSize(sizeGrid.cx / 2, sizeGrid.cy / 2);
	}
	else
	{
		m_szDrag = CBCGPSize(::GetSystemMetrics(SM_CXDRAG), ::GetSystemMetrics(SM_CYDRAG));
	}
}
//*******************************************************************************
void CBCGPVisualContainer::SetDirty(BOOL bSet) const
{
	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);
			
			pObject->SetDirty(bSet);
		}
	}
}
//*******************************************************************************
BOOL CBCGPVisualContainer::IsDirty() const
{
	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);
			
			if (pObject->IsDirty())
			{
				return TRUE;
			}
		}
	}

	return FALSE;
}
//*******************************************************************************
void CBCGPVisualContainer::Redraw()
{
	ASSERT_VALID (this);

	if (m_pWndOwner->GetSafeHwnd() == NULL || m_rect.IsRectEmpty())
	{
		return;
	}

	if (m_pWndOwner->GetExStyle() & WS_EX_LAYERED)
	{
		CBCGPBaseVisualCtrl* pCtrl = DYNAMIC_DOWNCAST(CBCGPBaseVisualCtrl, m_pWndOwner);
		if (pCtrl != NULL)
		{
			pCtrl->OnDrawLayeredPopup();
			return;
		}
	}

	m_pWndOwner->RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
}
//*******************************************************************************
void CBCGPVisualContainer::RedrawRect(const CBCGPRect& rect)
{
	ASSERT_VALID (this);

	if (m_pWndOwner->GetSafeHwnd() == NULL || m_rect.IsRectEmpty())
	{
		return;
	}

	if (m_pWndOwner->GetExStyle() & WS_EX_LAYERED)
	{
		CBCGPBaseVisualCtrl* pCtrl = DYNAMIC_DOWNCAST(CBCGPBaseVisualCtrl, m_pWndOwner);
		if (pCtrl != NULL)
		{
			pCtrl->OnDrawLayeredPopup();
			return;
		}
	}

	CRect rectGDI = rect;
	m_pWndOwner->RedrawWindow(rectGDI, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
}
//*******************************************************************************
void CBCGPVisualContainer::DrawObjects(CBCGPGraphicsManager* pGM, CBCGPRect rectClip, DWORD dwFlags, const CBCGPPoint& ptOffset)
{
	CBCGPRect rectInter;

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject == NULL)
		{
			continue;
		}

		ASSERT_VALID(pObject);

		BOOL bCacheImage = pObject->m_bCacheImage;
		if (m_bCacheImage || !m_bDrawDynamicObjectsOnTop)
		{
			pObject->m_bCacheImage = FALSE;
		}

		CBCGPRect rectSaved = pObject->GetRect();

		if (ptOffset != CBCGPPoint(0, 0))
		{
			pObject->m_rect.OffsetRect(-ptOffset);
		}

		CBCGPRect rectObj = pObject->GetRect().NormalizedRect();

		if (!rectObj.IsRectEmpty() && pObject->IsVisible() &&
			rectInter.IntersectRect (rectObj, rectClip))
		{
			pObject->OnBeforeDrawScrolled (-ptOffset);
			pObject->OnDraw(pGM, rectClip, dwFlags);
			pObject->OnAfterDrawScrolled ();
			pObject->SetDirty(FALSE);
		}

		pObject->m_bCacheImage = bCacheImage;

		if (ptOffset != CBCGPPoint(0, 0))
		{
			pObject->m_rect = rectSaved;
		}
	}
}
//*******************************************************************************
void CBCGPVisualContainer::DrawTrackingRects(CBCGPGraphicsManager* pGM)
{
	if (!m_bIsEditMode)
	{
		return;
	}

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);
			pObject->DrawTrackingRect(pGM, m_brSelOutline, m_brSelFill);
		}
	}
}
//*******************************************************************************
void CBCGPVisualContainer::OnDraw(CBCGPGraphicsManager* pGM, const CBCGPRect& rectClip)
{
	ASSERT_VALID (this);
	ASSERT_VALID (pGM);

	if (m_bIsEditMode)
	{
		m_ImageCache.Destroy();
	}

	DWORD dwDrawFlags = m_bDrawDynamicObjectsOnTop ? BCGP_DRAW_STATIC : (BCGP_DRAW_STATIC | BCGP_DRAW_DYNAMIC);

	if (m_bCacheImage && !pGM->IsOffscreen() && !m_bIsEditMode)
	{	
		if (m_ImageCache.GetHandle() == NULL || IsDirty())
		{
			m_ImageCache.Destroy();

			CBCGPRect rectBounds = GetBoundsRect(TRUE);
			
			rectBounds.InflateRect(1., 1.);
			rectBounds.left = m_rect.left;
			rectBounds.top = m_rect.top;
			rectBounds.right = max(m_rect.right, rectBounds.right);
			rectBounds.bottom = max(m_rect.bottom, rectBounds.bottom);

			CBCGPGraphicsManager* pGMMem = pGM->CreateOffScreenManager(rectBounds, &m_ImageCache);
			if (pGMMem != NULL)
			{
				CBCGPRect rectSaved = m_rect;
				m_rect = m_rect - m_rect.TopLeft();

				OnFillBackground(pGMMem);
				DrawObjects(pGMMem, m_rect, dwDrawFlags, rectSaved.TopLeft());
				OnDrawBorder(pGMMem);

				delete pGMMem;

				m_rect = rectSaved;
			}
		}
	}

	BOOL bCached = FALSE;
	CBCGPRect rectSel;
	BOOL bIsSemitransp = pGM->IsSupported(BCGP_GRAPHICS_MANAGER_COLOR_OPACITY);

	if (m_ImageCache.GetHandle() != NULL)
	{
		pGM->DrawImage(m_ImageCache, m_rect.TopLeft());
		bCached = TRUE;
	}
	else
	{
		if (m_nDragMode == HTNOWHERE && m_ptDragStart != CBCGPPoint(-1, -1) && m_pNewObject == NULL)
		{
			rectSel = CBCGPRect(m_ptDragStart, CBCGPSize(m_ptDragFinish.x - m_ptDragStart.x, m_ptDragFinish.y - m_ptDragStart.y));
		}

		OnFillBackground(pGM);

		if (!bIsSemitransp && !rectSel.IsRectEmpty())
		{
			DrawSelectedArea(pGM, rectSel);
		}

		DrawObjects(pGM, rectClip, dwDrawFlags, m_ptScrollOffset);
	}

	if (m_bDrawDynamicObjectsOnTop)
	{
		DrawObjects(pGM, rectClip, BCGP_DRAW_DYNAMIC, m_ptScrollOffset);
	}

	if (m_bDrawScrollBars)
	{
		if (!bCached)
		{
			DrawTrackingRects(pGM);

			if (bIsSemitransp && !rectSel.IsRectEmpty())
			{
				DrawSelectedArea(pGM, rectSel);
			}

			if (m_pNewObject != NULL)
			{
				ASSERT_VALID(m_pNewObject);
				m_pNewObject->OnDraw(pGM, rectClip);
			}

			OnDrawBorder(pGM);
		}

		if (!m_bNCScrollBars)
		{
			CBCGPRect rectInter;

			if (!m_rectBottomRight.IsRectEmpty() && !rectClip.IsRectEmpty() && rectInter.IntersectRect (m_rectBottomRight, rectClip))
			{
				pGM->FillRectangle(m_rectBottomRight, m_ScrollBarVert.GetColorTheme().m_brFace);
			}

			if (!rectClip.IsRectEmpty() && rectInter.IntersectRect (m_ScrollBarVert.GetRect(), rectClip))
			{
				m_ScrollBarVert.DoDraw(pGM);
			}

			if (!rectClip.IsRectEmpty() && rectInter.IntersectRect (m_ScrollBarHorz.GetRect(), rectClip))
			{
				m_ScrollBarHorz.DoDraw(pGM);
			}
		}
	}
}
//*******************************************************************************
void CBCGPVisualContainer::OnFillBackground(CBCGPGraphicsManager* pGM)
{
	if (!m_brFill.IsEmpty())
	{
		CBCGPRect rect = m_rect;
		rect.Scale(GetScaleRatioMid());

		pGM->FillRectangle(rect, m_brFill);
	}

	if (m_bIsEditMode && (m_dwEditFlags & BCGP_CONTAINER_GRID) == BCGP_CONTAINER_GRID)
	{
		OnDrawGrid(pGM);
	}
}
//*******************************************************************************
void CBCGPVisualContainer::OnDrawBorder(CBCGPGraphicsManager* pGM)
{
	if (m_bNCScrollBars && m_bScrollBars)
	{
		return;
	}

	CBCGPRect rect(m_rect);
	if (pGM->GetType() != CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER_GDI)
	{
		rect.right -= 1;
		rect.bottom -= 1;
	}

	if (!m_brOutline.IsEmpty())
	{
		pGM->DrawRectangle(rect, m_brOutline, 1);
		
		if (m_brOutline.GetOpacity() == 1.0)
		{
			return;
		}
	}

	if (!m_ScrollBarHorz.GetRect().IsRectEmpty() || !m_ScrollBarVert.GetRect().IsRectEmpty())
	{
		CBCGPColor color = CBCGPColor::White;

		if (!m_brFill.IsEmpty())
		{
			color = m_brFill.GetColor();
		}
		else if (!m_brOutline.IsEmpty())
		{
			color = m_brOutline.GetColor();
		}

		color.a = 1.0;

		pGM->DrawRectangle(rect, CBCGPBrush(color, 1.0), 1);
	}
}
//*******************************************************************************
void CBCGPVisualContainer::DrawSelectedArea(CBCGPGraphicsManager* pGM, const CBCGPRect& rectSel)
{
	pGM->FillRectangle(rectSel, m_brSelFill);
	pGM->DrawRectangle(rectSel, m_brSelOutline, 1);
}
//*******************************************************************************
void CBCGPVisualContainer::OnDrawGrid(CBCGPGraphicsManager* pGM)
{
	const CBCGPBrush& br = GetGridBrush();
	if (br.IsEmpty())
	{
		return;
	}

	double scaleRatio = GetScaleRatioMid();
	CBCGPSize sizeGrid(m_sizeGrid.cx * m_sizeScaleRatio.cx, m_sizeGrid.cy * m_sizeScaleRatio.cy);

	CBCGPRect rect = m_rect;
	rect.Scale(scaleRatio);

	if (m_styleGrid == BCGP_VISUAL_CONTAINER_GRID_LINES)
	{
		for (double x = rect.left + sizeGrid.cx; x < rect.right; x += sizeGrid.cx)
		{
			pGM->DrawLine(x, rect.top, x, rect.bottom, br, scaleRatio);
		}

		for (double y = rect.top + sizeGrid.cy; y < rect.bottom; y += sizeGrid.cy)
		{
			pGM->DrawLine(rect.left, y, rect.right, y, br, scaleRatio);
		}
	}
	else	// BCGP_VISUAL_CONTAINER_GRID_DOTS
	{
		for (double x = rect.left; x <= rect.right; x += sizeGrid.cx)
		{
			for (double y = rect.top; y <= rect.bottom; y += sizeGrid.cy)
			{
				CBCGPEllipse dot(CBCGPPoint(x, y), CBCGPSize(2.0 * scaleRatio, 2.0 * scaleRatio));
				pGM->FillEllipse(dot, br);
			}
		}
	}
}
//*******************************************************************************
#ifndef _BCGPCHART_STANDALONE

BOOL CBCGPVisualContainer::LoadFromXML (UINT uiXMLResID)
{
	ASSERT_VALID (this);
	return LoadFromXML (MAKEINTRESOURCE (uiXMLResID));
}
//*******************************************************************************
BOOL CBCGPVisualContainer::LoadFromXML (LPCTSTR lpszXMLResID)
{
	CBCGPVisualInfo info;
	CBCGPVisualInfoLoader loader (info);

	if (!loader.Load (lpszXMLResID))
	{
		TRACE0("Cannot load dashboard from buffer\n");
		return FALSE;
	}

	CBCGPVisualConstructor constr (info);
	constr.Construct (*this);

	m_ImageCache.Destroy();

	SetDirty (TRUE);
	AdjustLayout ();
	Redraw ();

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::LoadFromFile (LPCTSTR lpszFileName)
{
	CBCGPVisualInfo info;
	CBCGPVisualInfoLoader loader (info);

	if (!loader.LoadFromFile (lpszFileName))
	{
		TRACE0("Cannot load dashboard from buffer\n");
		return FALSE;
	}

	CBCGPVisualConstructor constr (info);
	constr.Construct (*this);

	m_ImageCache.Destroy();

	SetDirty (TRUE);
	AdjustLayout ();
	Redraw ();

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::LoadFromBuffer (LPCTSTR lpszXMLBuffer)
{
	ASSERT_VALID (this);
	ASSERT (lpszXMLBuffer != NULL);

	CBCGPVisualInfo info;
	CBCGPVisualInfoLoader loader (info);

	if (!loader.LoadFromBuffer (lpszXMLBuffer))
	{
		TRACE0("Cannot load ribbon from buffer\n");
		return FALSE;
	}

	CBCGPVisualConstructor constr (info);
	constr.Construct (*this);

	m_ImageCache.Destroy();

	SetDirty (TRUE);
	AdjustLayout ();
	Redraw ();

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::SaveToXML (const CString& strFileName)
{
	CBCGPVisualInfo info;

	CBCGPVisualCollector collector (info);
	collector.Collect (*this);

	CBCGPVisualInfoWriter writer (info);
	return writer.Save (strFileName);
}
//*********************************************************************************
void CBCGPVisualContainer::EnableLayout(BOOL bEnable)
{
	if (m_pLayout != NULL)
	{
		delete m_pLayout;
		m_pLayout = NULL;
	}

	if (!bEnable)
	{
		return;
	}

	m_pLayout = new CBCGPVisualLayout(*this);
}

#endif

void CBCGPVisualContainer::AdjustLayout()
{
#ifndef _BCGPCHART_STANDALONE
	m_ImageCache.Destroy();

	if (m_bIsPrinting)
	{
		return;
	}

	if (m_pLayout == NULL)
	{
		if (GetOwner()->GetSafeHwnd() != NULL && m_bNCScrollBars)
		{
			GetOwner()->SetWindowPos (NULL, -1, -1, -1, -1, 
				SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
		}

		AdjustScrollBars();
		return;
	}

	ASSERT_VALID(m_pLayout);
	m_pLayout->AdjustLayout();
#endif
	
	AdjustScrollBars();
}
//*******************************************************************************
void CBCGPVisualContainer::AdjustScrollBars(BOOL bRedraw)
{
	m_ScrollBarVert.SetParentVisualContainer(this);
	m_ScrollBarHorz.SetParentVisualContainer(this);

	double xOffset = m_ScrollBarHorz.GetOffset();
	double yOffset = m_ScrollBarVert.GetOffset();

	m_rectBottomRight.SetRectEmpty();

	m_ScrollBarVert.Reset();
	m_ScrollBarHorz.Reset();

	m_sizeScrollTotal = CBCGPSize(0, 0);

	CBCGPPoint ptScrollOffsetOld = m_ptScrollOffset;

	double dblMaxX = 0.;
	double dblMaxY = 0.;

	CBCGPRect rectScrollVert;
	CBCGPRect rectScrollHorz;

	int i = 0;

	CBCGPRect rectWindow;
	if (m_bNCScrollBars)
	{
		CWnd* pWndOwner = GetOwner();
		if (pWndOwner->GetSafeHwnd() != NULL)
		{
			CRect rectOwner;

			pWndOwner->GetWindowRect(rectOwner);
			rectOwner.OffsetRect(-rectOwner.TopLeft());

			rectWindow = rectOwner;
		}
	}

	if (m_bScrollBars)
	{
		for (i = 0; i < m_arObjects.GetSize(); i++)
		{
			CBCGPBaseVisualObject* pObject = m_arObjects[i];
			if (pObject == NULL)
			{
				continue;
			}

			ASSERT_VALID(pObject);

			dblMaxX = max(dblMaxX, pObject->GetRect().right);
			dblMaxY = max(dblMaxY, pObject->GetRect().bottom);
		}
	}

	const int cxScrollBar = ::GetSystemMetrics(SM_CXVSCROLL);
	const int cyScrollBar = ::GetSystemMetrics(SM_CYHSCROLL);

	if (dblMaxX >= m_rect.right - cxScrollBar)
	{
		m_sizeScrollTotal.cx = dblMaxX;

		rectScrollHorz = m_bNCScrollBars ? rectWindow : m_rect;
		rectScrollHorz.DeflateRect(1, 1);

		rectScrollHorz.top = rectScrollHorz.bottom - cyScrollBar;

		if (dblMaxY >= m_rect.bottom - cyScrollBar)
		{
			rectScrollHorz.right -= cxScrollBar;
		}
	}

	if (dblMaxY >= m_rect.bottom - cyScrollBar)
	{
		m_sizeScrollTotal.cy = dblMaxY;

		rectScrollVert = m_bNCScrollBars ? rectWindow : m_rect;
		rectScrollVert.DeflateRect(1, 1);

		rectScrollVert.left = rectScrollVert.right - cxScrollBar;

		if (dblMaxX >= m_rect.right - cxScrollBar)
		{
			rectScrollVert.bottom -= cyScrollBar;
		}
	}

	m_ScrollBarHorz.SetTotal(m_sizeScrollTotal.cx);
	m_ScrollBarVert.SetTotal(m_sizeScrollTotal.cy);

	m_ptScrollOffset.x = max(0, min(m_ptScrollOffset.x, m_sizeScrollTotal.cx));
	m_ptScrollOffset.y = max(0, min(m_ptScrollOffset.y, m_sizeScrollTotal.cy));

	if (ptScrollOffsetOld != m_ptScrollOffset)
	{
		for (i = 0; i < m_arObjects.GetSize(); i++)
		{
			CBCGPBaseVisualObject* pObject = m_arObjects[i];
			if (pObject != NULL)
			{
				ASSERT_VALID(pObject);
				pObject->SetScrollOffset(m_ptScrollOffset);
			}
		}
	}

	if (m_bScrollBars)
	{
		m_ScrollBarHorz.SetOffset(bcg_clamp(xOffset, 0.0, m_ScrollBarHorz.GetTotal()));
		m_ScrollBarVert.SetOffset(bcg_clamp(yOffset, 0.0, m_ScrollBarVert.GetTotal()));

		m_ScrollBarHorz.SetRect(rectScrollHorz);
		m_ScrollBarVert.SetRect(rectScrollVert);

		if (!m_ScrollBarVert.GetRect().IsRectEmpty() && !m_ScrollBarHorz.GetRect().IsRectEmpty())
		{
			m_rectBottomRight = m_bNCScrollBars ? rectWindow : m_rect;
			m_rectBottomRight.DeflateRect(1, 1);

			m_rectBottomRight.top = m_rectBottomRight.bottom - m_ScrollBarHorz.GetRect().Height() - 1;
			m_rectBottomRight.left = m_rectBottomRight.right - m_ScrollBarVert.GetRect().Width() - 1;
		}

		SetDirty();

		if (bRedraw)
		{
			Redraw();
		}
	}
}
//*******************************************************************************
void CBCGPVisualContainer::SetOutlineBrush(const CBCGPBrush& br)
{
	m_brOutline = br;
}
//*******************************************************************************
void CBCGPVisualContainer::SetFillBrush(const CBCGPBrush& br)
{
	m_brFill = br;
}
//*******************************************************************************
void CBCGPVisualContainer::SetGridBrush(const CBCGPBrush& br, BCGP_VISUAL_CONTAINER_GRID_STYLE style)
{
	if (style == BCGP_VISUAL_CONTAINER_GRID_STYLE_CURRENT)
	{
		style = m_styleGrid;
	}

	if (style == BCGP_VISUAL_CONTAINER_GRID_LINES)
	{
		m_brGrid = br;
	}
	else
	{
		m_brGridDots = br;
	}
}
//*******************************************************************************
void CBCGPVisualContainer::SetGridSize(const CBCGPSize& size)
{
	if (size.cx < 2.0 || size.cy < 2.0)
	{
		ASSERT(FALSE);
		return;
	}

	m_sizeGrid = size;

	if ((m_dwEditFlags & BCGP_CONTAINER_GRID) == BCGP_CONTAINER_GRID &&
		(m_dwEditFlags & BCGP_CONTAINER_SNAP_TO_GRID) == BCGP_CONTAINER_SNAP_TO_GRID)
	{
		CBCGPSize sizeGrid(m_sizeGrid.cx * m_sizeScaleRatio.cx, m_sizeGrid.cy * m_sizeScaleRatio.cy);
		m_szDrag = CBCGPSize(sizeGrid.cx / 2, sizeGrid.cy / 2);
	}
}
//*******************************************************************************
void CBCGPVisualContainer::SetGridStyle(BCGP_VISUAL_CONTAINER_GRID_STYLE style)
{
	m_styleGrid = style;
}
//*******************************************************************************
void CBCGPVisualContainer::SetDrawDynamicObjectsOnTop(BOOL bSet)
{
	m_bDrawDynamicObjectsOnTop = bSet;
	SetDirty();
}
//*******************************************************************************
void CBCGPVisualContainer::OnChangeVisualManager()
{
	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);		
			pObject->OnChangeVisualManager();
		}
	}

	m_ScrollBarHorz.OnChangeVisualManager();
	m_ScrollBarVert.OnChangeVisualManager();
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnMouseDown(int nButton, const CBCGPPoint& pt)
{
	m_ptClick = pt;

	CBCGPBaseVisualObject* pObject = GetFromPoint(pt);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		if (pObject->OnMouseDown(nButton, pt))
		{
			return TRUE;
		}
	}

	if (!m_bNCScrollBars)
	{
		if (m_ScrollBarVert.OnMouseDown(pt) || m_ScrollBarHorz.OnMouseDown(pt))
		{
			return TRUE;
		}
	}

	BOOL bSelChanged = FALSE;

	if (m_bIsEditMode)
	{
		if (nButton == 0)
		{
			m_nDragMode = HTNOWHERE;

			if (pObject != NULL && !m_bAddNewObjectMode)
			{
				const BOOL bCtrl = ::GetAsyncKeyState (VK_CONTROL) & 0x8000;

				BOOL bWasSelected = pObject->IsSelected ();

				if ((m_dwEditFlags & BCGP_CONTAINER_ENABLE_COPY) == BCGP_CONTAINER_ENABLE_COPY && bWasSelected)
				{
					// Don't unselect selected object
				}
				else
				{
					Select(pObject, bCtrl);
					bSelChanged = bWasSelected != pObject->IsSelected ();

					pObject->Redraw();
				}

				m_nDragMode = pObject->HitTest(pt);

			}
			else
			{
				const int nSelCount = GetSelCount();

				ClearSelection();

				bSelChanged = (nSelCount != GetSelCount());

				if (IsSingleSel() && !m_bAddNewObjectMode)
				{
					Redraw();
					FireSelectionChangedEvent();
					return FALSE;
				}
			}

			m_ptDragStart = m_ptDragFinish = pt;
			
			if (m_bAddNewObjectMode)
			{
				m_pNewObject = OnStartAddNewObject(pt);

				if (m_pNewObject != NULL)
				{
					m_pNewObject->SetRect(CBCGPRect(pt, CBCGPSize(1., 1.)));
				}
			}

			Redraw();

			if (bSelChanged)
			{
				FireSelectionChangedEvent();
			}

			return TRUE;
		}
		else if (nButton == 1 && pObject == NULL)
		{
			ClearSelection();
			FireSelectionChangedEvent();
		}
	}

	return FALSE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnGetToolTip(const CBCGPPoint& pt, CString& strToolTip, CString& strDescr)
{
	CBCGPBaseVisualObject* pObject = GetFromPoint(pt);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);

		if (pObject->OnGetToolTip(pt, strToolTip, strDescr))
		{
			return TRUE;
		}
	}

	return FALSE;
}
//*******************************************************************************
COLORREF CBCGPVisualContainer::GetInfoTipColor(const CBCGPPoint& pt, int nColor)
{
	CBCGPBaseVisualObject* pObject = GetFromPoint(pt);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		return pObject->GetInfoTipColor(pt, nColor);
	}
	
	return (COLORREF)-1;
}
//*******************************************************************************
void CBCGPVisualContainer::Select(CBCGPBaseVisualObject* pObject, BOOL bAddToSel)
{
	ASSERT_VALID(pObject);

	if (!pObject->IsSelected())
	{
		if (!bAddToSel || IsSingleSel())
		{
			ClearSelection();
		}

		pObject->SetSelected();
		m_lstSel.AddTail(pObject);
	}
	else if (bAddToSel)
	{
		pObject->SetSelected(FALSE);

		POSITION pos = m_lstSel.Find(pObject);
		if (pos != NULL)
		{
			m_lstSel.RemoveAt(pos);
		}
	}
}
//*******************************************************************************
void CBCGPVisualContainer::SelectAll()
{
	if (!m_bIsEditMode || IsSingleSel ())
	{
		TRACE(_T("Container is not in edit mode or multi selection disabled.\n"));
		ASSERT(FALSE);
		return;
	}

	BOOL bSelChanged = FALSE;
	const int count = GetCount ();
	for (int i = 0; i < count; i++)
	{
		CBCGPBaseVisualObject* pObject = GetAt(i);
		if (!pObject->IsSelected ())
		{
			Select (pObject, TRUE);
			bSelChanged = TRUE;
		}
	}

	if (bSelChanged)
	{
		Redraw ();
		FireSelectionChangedEvent();
	}
}
//*******************************************************************************
void CBCGPVisualContainer::ClearSelection(BOOL bRedraw)
{
	if (!m_bIsEditMode)
	{
		return;
	}

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);

			BOOL bRedrawObject = pObject->IsSelected();
			
			pObject->SetSelected(FALSE);

			if (bRedrawObject)
			{
				CBCGPRect rectRedraw = pObject->GetRect();
				if (!rectRedraw.IsRectEmpty())
				{
					rectRedraw.InflateRect(SEL_MARKER_SIZE, SEL_MARKER_SIZE);
					pObject->InvalidateRect(rectRedraw);
				}
			}
		}
	}

	m_lstSel.RemoveAll();

	if (bRedraw && m_pWndOwner->GetSafeHwnd() != NULL)
	{
		m_pWndOwner->UpdateWindow();
	}
}
//*******************************************************************************
void CBCGPVisualContainer::SetHitTestSelected(BOOL bSet)
{
	m_bHitTestSelected = bSet;
}
//*******************************************************************************
CBCGPBaseVisualObject* CBCGPVisualContainer::GetSel(int nIndex) const
{
	if (m_lstSel.IsEmpty())
	{
		return NULL;
	}

	POSITION pos = m_lstSel.FindIndex(nIndex);
	if (pos == NULL)
	{
		return NULL;
	}

	return m_lstSel.GetAt(pos);
}
//*******************************************************************************
static void _ClampRect(const CBCGPRect& rectOld, CBCGPRect& rectNew, CBCGPSize sizeMin, BOOL bEmpty)
{
	if (sizeMin.IsEmpty() && bEmpty)
	{
		sizeMin.SetSize (1.0, 1.0);
	}

	if (fabs(rectNew.Width()) < sizeMin.cx || fabs(rectNew.Height()) < sizeMin.cy)
	{
		rectNew = rectOld;
	}
}
//*******************************************************************************
void CBCGPVisualContainer::OnMouseUp(int nButton, const CBCGPPoint& pt)
{
	m_bHitTestSelected = FALSE;

	m_ScrollBarVert.OnCancelMode();
	m_ScrollBarHorz.OnCancelMode();

	int i = 0;

	for (i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);
			
			if (pObject->IsCaptured())
			{
				pObject->OnMouseUp(nButton, pt);
				return;
			}
		}
	}

	CBCGPBaseVisualObject* pObject = GetFromPoint(pt);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		pObject->OnMouseUp(nButton, pt);
	}

	if (m_ptDragStart == CBCGPPoint(-1, -1))
	{
		m_ptClick = CBCGPPoint(-1, -1);
		return;	
	}

	BOOL bSelChanged = FALSE;

	if (m_nDragMode == HTNOWHERE)
	{
		CBCGPRect rectSel(m_ptDragStart, CBCGPSize(m_ptDragFinish.x - m_ptDragStart.x, m_ptDragFinish.y - m_ptDragStart.y));

		rectSel.OffsetRect(m_ptScrollOffset);

		if (m_pNewObject != NULL)
		{
			if ((m_pNewObject->GetEditFlags() & BCGP_EDIT_NO_NORMALIZE_RECT) == 0)
			{
				rectSel.Normalize();
			}

			if (OnFinishAddNewObject(m_pNewObject, rectSel))
			{
				m_pNewObject->SetRect(rectSel);
				Add(m_pNewObject);

				Select(m_pNewObject);
				bSelChanged = TRUE;
			}
			else
			{
				delete m_pNewObject;
			}

			m_pNewObject = NULL;
			m_bAddNewObjectMode = FALSE;
		}
		else
		{
			rectSel.Normalize();

			for (i = 0; i < m_arObjects.GetSize(); i++)
			{
				CBCGPBaseVisualObject* pObjectCurr = m_arObjects[i];
				if (pObjectCurr != NULL)
				{
					ASSERT_VALID(pObjectCurr);
					
					CBCGPRect rectInter;
					
					if (rectInter.IntersectRect(pObjectCurr->GetRect().NormalizedRect(), rectSel))
					{
						Select(pObjectCurr, TRUE);
						bSelChanged = TRUE;
					}
				}
			}
		}
	}
	else
	{
		if (fabs(m_ptClick.x - pt.x) > m_szDrag.cx || fabs(m_ptClick.y - pt.y) > m_szDrag.cy || m_ptClick != m_ptDragStart)
		{
			BOOL bObjectLocationWasChanged = FALSE;

			if (!m_lstSel.IsEmpty())
			{
				FireObjectMoveEvent(FALSE, FALSE);
			}

 			for (POSITION pos = m_lstSel.GetHeadPosition(); pos != NULL;)
			{
				CBCGPBaseVisualObject* pObjectCurr = m_lstSel.GetNext(pos);
				ASSERT_VALID(pObjectCurr);

				if (!pObjectCurr->m_rectTrack.IsRectEmpty())
				{
					CBCGPRect rect = pObjectCurr->m_rectTrack;

					if ((pObjectCurr->GetEditFlags() & BCGP_EDIT_NO_NORMALIZE_RECT) == 0)
					{
						rect.Normalize();
					}

					CBCGPRect rectSave(rect);

					BOOL bSnapToGrid = pObjectCurr->IsSnapTpGridAvailable() ? SnapToGrid(rect) : FALSE;

					if (bSnapToGrid)
					{
						_ClampRect(rectSave, rect, pObjectCurr->GetMinSize(), TRUE);
					}

					if (pObjectCurr->GetRect() != rect)
					{
						CBCGPRect rectArea = m_rect;
						rectArea.Scale (m_sizeScaleRatio);

						if (rect.bottom <= rectArea.top)
						{
							rect.OffsetRect(0, rectArea.top - rect.bottom + 2);
						}

						if (rect.right <= rectArea.left)
						{
							rect.OffsetRect(rectArea.left - rect.right + 2, 0);
						}

						BOOL bCtrlIsPressed = GetAsyncKeyState (VK_CONTROL);

						if (bCtrlIsPressed && (m_dwEditFlags & BCGP_CONTAINER_ENABLE_COPY) == BCGP_CONTAINER_ENABLE_COPY)
						{
							if (OnCopyObject(pObjectCurr, rect))
							{
								if (bSnapToGrid)
								{
									pObjectCurr->SetTrackingRect(rect);
								}
							}
						}
						else
						{
							CBCGPRect rectOld = pObjectCurr->GetRect();

							pObjectCurr->SetRect(rect);

							if (!bObjectLocationWasChanged)
							{
								bObjectLocationWasChanged = rectOld != pObjectCurr->GetRect();
							}
						}
					}
					else if (bSnapToGrid && pObjectCurr->m_rectTrack != rect)
					{
						pObjectCurr->SetTrackingRect(rect);
					}
				}
			}

			if (!m_lstSel.IsEmpty())
			{
				FireObjectMoveEvent(TRUE, bObjectLocationWasChanged);
			}
		}

		if (GetOwner()->GetSafeHwnd() != NULL && m_bNCScrollBars)
		{
			GetOwner()->SetWindowPos (NULL, -1, -1, -1, -1, 
				SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
		}

		AdjustScrollBars();
	}

	m_ptClick = m_ptDragStart = m_ptDragFinish = CBCGPPoint(-1, -1);

	Redraw();

	if (bSelChanged)
	{
		FireSelectionChangedEvent();
	}
}
//*******************************************************************************
void CBCGPVisualContainer::OnMouseMove(const CBCGPPoint& pt)
{
	if (m_ScrollBarHorz.OnDragThumb(pt) || m_ScrollBarVert.OnDragThumb(pt))
	{
		return;
	}

	if (m_bScrollBars)
	{
		if (m_ScrollBarVert.OnMouseMove(pt) || m_ScrollBarHorz.OnMouseMove(pt))
		{
			return;
		}
	}

	if (m_bAddNewObjectMode && m_ptDragStart == CBCGPPoint(-1, -1))
	{
		SetCursor (AfxGetApp ()->LoadStandardCursor (IDC_CROSS));
		return;
	}

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pArObject = m_arObjects[i];
		if (pArObject == NULL)
		{
			continue;
		}

		ASSERT_VALID(pArObject);

		if (pArObject->IsCaptured())
		{
			pArObject->OnMouseMove(pt);
			return;
		}

		if (pArObject->m_bIsTracked)
		{
			CBCGPRect rect = pArObject->m_rect;
			rect.Normalize();

			if (!rect.PtInRect(pt))
			{
				pArObject->m_bIsTracked = FALSE;
				pArObject->OnMouseLeave();
			}
		}
	}

	CBCGPBaseVisualObject* pObject = GetFromPoint(pt);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);

		pObject->OnMouseMove(pt);
		pObject->m_bIsTracked = TRUE;
	}

	if (m_ptDragStart == CBCGPPoint(-1, -1))
	{
		return;	
	}

	if (m_ptDragStart == m_ptClick && fabs(m_ptClick.x - pt.x) <= m_szDrag.cx && fabs(m_ptClick.y - pt.y) <= m_szDrag.cy)
	{
		return;
	}

	if (m_nDragMode != HTNOWHERE)
	{
		MoveTrackingRects (pt);

		if (m_bScrollBars)
		{
			if ((pt.y > m_rect.bottom || pt.y < m_rect.top) && !m_ScrollBarVert.GetRect().IsRectEmpty())
			{
				m_ScrollBarVert.OnScrollStep(pt.y > m_rect.bottom);
			}

			if ((pt.x > m_rect.right || pt.x < m_rect.left) && !m_ScrollBarHorz.GetRect().IsRectEmpty())
			{
				m_ScrollBarHorz.OnScrollStep(pt.x > m_rect.right);
			}
		}

		m_ptDragStart = pt;
	}
	else if (m_pNewObject != NULL)
	{
		ASSERT_VALID(m_pNewObject);

		CBCGPRect rect(m_ptDragStart, CBCGPSize(m_ptDragFinish.x - m_ptDragStart.x, m_ptDragFinish.y - m_ptDragStart.y));
		if ((m_pNewObject->GetEditFlags() & BCGP_EDIT_NO_NORMALIZE_RECT) == 0)
		{
			rect.Normalize();
		}

		m_pNewObject->SetRect(rect);
	}

	m_ptDragFinish = pt;

	if (m_nDragMode == HTNOWHERE)
	{
		SetCursor (AfxGetApp ()->LoadStandardCursor (IDC_CROSS));
	}

	Redraw();
}
//*******************************************************************************
void CBCGPVisualContainer::OnNcMouseMove(const CBCGPPoint& pt)
{
	if (m_bScrollBars && m_bNCScrollBars && m_pWndOwner->GetSafeHwnd() != NULL)
	{
		CPoint ptClient = pt;
		m_pWndOwner->ScreenToClient(&ptClient);

		m_ScrollBarVert.OnMouseMove(ptClient);
		m_ScrollBarHorz.OnMouseMove(ptClient);
	}
}
//*******************************************************************************
void CBCGPVisualContainer::MoveTrackingRects (CBCGPPoint pt)
{
	if (m_nDragMode == HTNOWHERE)
	{
		return;
	}

	if (m_ptDragStart == CBCGPPoint(-1, -1))
	{
		return;	
	}

	CBCGPPoint ptOffset = pt - m_ptDragStart;

	for (POSITION pos = m_lstSel.GetHeadPosition(); pos != NULL;)
	{
		CBCGPBaseVisualObject* pObject = m_lstSel.GetNext(pos);
		ASSERT_VALID(pObject);

		CBCGPRect rect = pObject->m_rectTrack;

		BOOL bLockAspRatio = (pObject->GetEditFlags() & BCGP_EDIT_SIZE_LOCK_ASPECT_RATIO) == BCGP_EDIT_SIZE_LOCK_ASPECT_RATIO;

		if (bLockAspRatio && m_nDragMode != HTCAPTION)
		{
			if (m_nDragMode == HTTOPLEFT || m_nDragMode == HTBOTTOMRIGHT)
			{
				if (fabs(ptOffset.x) > fabs(ptOffset.y))
				{
					ptOffset.y = ptOffset.x;
				}
				else
				{
					ptOffset.x = ptOffset.y;
				}
			}
			else if (m_nDragMode == HTTOPRIGHT || m_nDragMode == HTBOTTOMLEFT)
			{
				if (fabs(ptOffset.x) < fabs(ptOffset.y))
				{
					ptOffset.y = -ptOffset.x;
				}
				else
				{
					ptOffset.x = -ptOffset.y;
				}
			}
		}

		switch (m_nDragMode)
		{
		case HTCAPTION:
			rect += ptOffset;
			break;

		case HTLEFT:
			rect.left += ptOffset.x;
			if (bLockAspRatio)
			{
				rect.bottom -= ptOffset.x;
			}
			break;

		case HTRIGHT:
			rect.right += ptOffset.x;
			if (bLockAspRatio)
			{
				rect.bottom += ptOffset.x;
			}
			break;

		case HTTOP:
			rect.top += ptOffset.y;
			if (bLockAspRatio)
			{
				rect.right -= ptOffset.y;
			}
			break;

		case HTBOTTOM:
			rect.bottom += ptOffset.y;
			if (bLockAspRatio)
			{
				rect.right += ptOffset.y;
			}
			break;

		case HTTOPLEFT:
			rect.left += ptOffset.x;
			rect.top += ptOffset.y;
			break;

		case HTBOTTOMLEFT:
			rect.left += ptOffset.x;
			rect.bottom += ptOffset.y;
			break;

		case HTTOPRIGHT:
			rect.right += ptOffset.x;
			rect.top += ptOffset.y;
			break;

		case HTBOTTOMRIGHT:
			rect.right += ptOffset.x;
			rect.bottom += ptOffset.y;
			break;
		}

		_ClampRect(pObject->m_rectTrack, rect, pObject->GetMinSize(), TRUE);

		pObject->SetTrackingRect(rect);
	}
}
//*******************************************************************************
static int _SnapToGrid(double dbVal, int nGridStep)
{
	int val = ((int)dbVal) / nGridStep * nGridStep;
	if (dbVal - val > val + nGridStep - dbVal)
	{
		val += nGridStep;
	}

	return val;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::SnapToGrid(CBCGPRect& rect)
{
	if ((m_dwEditFlags & BCGP_CONTAINER_GRID) == BCGP_CONTAINER_GRID &&
		(m_dwEditFlags & BCGP_CONTAINER_SNAP_TO_GRID) == BCGP_CONTAINER_SNAP_TO_GRID)
	{
		CBCGPSize sizeGrid(m_sizeGrid.cx * m_sizeScaleRatio.cx, m_sizeGrid.cy * m_sizeScaleRatio.cy);

		CBCGPRect rectOld = rect;

		BOOL bKeepX = FALSE;
		int nX1Action = 1;	// -1 keep left, 0 - snap to grid, 1 - keep width
		int nX2Action = 1;	// -1 keep right, 0 - snap to grid, 1 - keep width

		BOOL bKeepY = FALSE;
		int nY1Action = 1;	// -1 keep top, 0 - snap to grid, 1 - keep height
		int nY2Action = 1;	// -1 keep bottom, 0 - snap to grid, 1 - keep height

		switch (m_nDragMode)
		{
		case HTLEFT:
			nX1Action = 0;
			nX2Action = -1;

			bKeepY = TRUE;
			break;

		case HTRIGHT:
			nX1Action = -1;
			nX2Action = 0;

			bKeepY = TRUE;
			break;

		case HTTOP:
			bKeepX = TRUE;

			nY1Action = 0;
			nY2Action = -1;
			break;
				
		case HTTOPLEFT:
			nX1Action = 0;
			nX2Action = -1;

			nY1Action = 0;
			nY2Action = -1;
			break;

		case HTTOPRIGHT:
			nX1Action = -1;
			nX2Action = 0;

			nY1Action = 0;
			nY2Action = -1;
			break;

		case HTBOTTOM:
			bKeepX = TRUE;

			nY1Action = -1;
			nY2Action = 0;
			break;

		case HTBOTTOMLEFT:
			nX1Action = 0;
			nX2Action = -1;

			nY1Action = -1;
			nY2Action = 0;
			break;

		case HTBOTTOMRIGHT:
			nX1Action = -1;
			nX2Action = 0;

			nY1Action = -1;
			nY2Action = 0;
			break;
		}

		int x1 = (int)rectOld.left;
		int x2 = (int)rectOld.right;

		if (!bKeepX)
		{
			x1 = nX1Action < 0 ? (int)rectOld.left : _SnapToGrid(rect.left, (int) sizeGrid.cx);
			x2 = nX2Action < 0 ? (int)rectOld.right : nX2Action == 0 ? _SnapToGrid(rect.right, (int) sizeGrid.cx) : (int)rectOld.Width() + x1;

			if (nX1Action > 0)
			{
				x1 = (int)(x2 - rectOld.Width());
			}
		}

		int y1 = (int)rectOld.top;
		int y2 = (int)rectOld.bottom;

		if (!bKeepY)
		{
			y1 = nY1Action < 0 ? (int)rectOld.top : _SnapToGrid(rect.top, (int) sizeGrid.cy);
			y2 = nY2Action < 0 ? (int)rectOld.bottom : nY2Action == 0 ? _SnapToGrid(rect.bottom, (int) sizeGrid.cy) : (int)rectOld.Height() + y1;

			if (nY1Action > 0)
			{
				y1 = (int)(y2 - rectOld.Height());
			}
		}

		rect = CBCGPRect(x1, y1, x2, y2);
		return rect != rectOld;
	}

	return FALSE;
}
//*******************************************************************************
void CBCGPVisualContainer::OnMouseLeave()
{
	m_ScrollBarVert.OnMouseLeave();
	m_ScrollBarHorz.OnMouseLeave();

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);
			
			if (pObject->m_bIsTracked)
			{
				pObject->m_bIsTracked = FALSE;
				pObject->OnMouseLeave();
			}
		}
	}
}
//*******************************************************************************
void CBCGPVisualContainer::OnCancelMode()
{
	m_bHitTestSelected = FALSE;

	m_ScrollBarVert.OnCancelMode();
	m_ScrollBarHorz.OnCancelMode();

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);
			
			pObject->OnCancelMode();
		}
	}
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnSetMouseCursor(const CBCGPPoint& pt)
{
	if (m_bAddNewObjectMode)
	{
		SetCursor (AfxGetApp ()->LoadStandardCursor (IDC_CROSS));
		return TRUE;
	}

	CBCGPBaseVisualObject* pObject = GetFromPoint(pt);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		return pObject->OnSetMouseCursor(pt);
	}

	return FALSE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnMouseWheel(const CBCGPPoint& pt, short zDelta)
{
	if (m_bScaleByMouseWheel && (::GetAsyncKeyState (VK_CONTROL) & 0x8000))	// Ctrl is pressed
	{
		const double dblDelta = zDelta > 0 ? 0.1 : -0.1;
		double dblScale = bcg_clamp(GetScaleRatioMid() + dblDelta, m_dblMinScaleRatio, m_dblMaxScaleRatio);

		CBCGPSize szScaleNew = CBCGPSize(dblScale, dblScale);

		if (m_sizeScaleRatio != szScaleNew)
		{
			SetScaleRatio(szScaleNew);
			Redraw();

			CWnd* pWndOwner = m_pWndOwner->GetOwner();
			if (pWndOwner->GetSafeHwnd() != NULL)
			{
				pWndOwner->SendMessage(BCGM_CONTAINER_SCALE_CHANGED, 0, (LPARAM)this);
			}
		}

		return TRUE;
	}

	if (m_bScrollBars)
	{
		BOOL bShiftIsPressed = (::GetAsyncKeyState (VK_SHIFT) & 0x8000);

		if (!m_ScrollBarVert.GetRect().IsRectEmpty() && (!bShiftIsPressed || m_ScrollBarHorz.GetRect().IsRectEmpty()))
		{
			const int nSteps = abs(zDelta) / WHEEL_DELTA;
			for (int i = 0; i < nSteps; i++)
			{
				m_ScrollBarVert.OnScrollStep(zDelta < 0);
			}

			return TRUE;
		}

		if (!m_ScrollBarHorz.GetRect().IsRectEmpty())
		{
			const int nSteps = abs(zDelta) / WHEEL_DELTA;
			for (int i = 0; i < nSteps; i++)
			{
				m_ScrollBarHorz.OnScrollStep(zDelta < 0);
			}

			return TRUE;
		}
	}

	CBCGPBaseVisualObject* pObject = GetFromPoint(pt);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		return pObject->OnMouseWheel(pt, zDelta);
	}

	return FALSE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnKeyboardDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (IsEditMode ())
	{
		if (m_ptDragStart != CBCGPPoint(-1, -1) && m_nDragMode != HTNOWHERE)
		{
			return FALSE;
		}

		if ((m_dwEditFlags & BCGP_CONTAINER_ENABLE_KEYBOARD) == BCGP_CONTAINER_ENABLE_KEYBOARD)
		{
			BOOL bResize = ::GetAsyncKeyState(VK_SHIFT) & 0x8000;
			const double delta = 1.0;
			CBCGPPoint ptOffset;

			switch (nChar)
			{
			case VK_LEFT:
				ptOffset.x -= delta;
				break;
			case VK_RIGHT:
				ptOffset.x += delta;
				break;
			case VK_UP:
				ptOffset.y -= delta;
				break;
			case VK_DOWN:
				ptOffset.y += delta;
				break;
			}

			BOOL bChanged = FALSE;

			if (ptOffset != CBCGPPoint())
			{
				if (!m_lstSel.IsEmpty())
				{
					FireObjectMoveEvent(FALSE, FALSE);
				}

				POSITION pos = m_lstSel.GetHeadPosition ();
				while (pos != NULL)
				{
					CBCGPBaseVisualObject* pObject = m_lstSel.GetNext (pos);
					if (pObject == NULL || !pObject->IsVisible ())
					{
						continue;
					}

					UINT uiEditFlags = pObject->GetEditFlags();

					if ((uiEditFlags & BCGP_EDIT_NO_KEYBOARD) == BCGP_EDIT_NO_KEYBOARD)
					{
						continue;
					}

					CBCGPRect rect (pObject->GetRect ());
					if (bResize)
					{
						if ((uiEditFlags & BCGP_EDIT_NOSIZE) == BCGP_EDIT_NOSIZE)
						{
							continue;
						}

						CBCGPPoint ptSize (ptOffset);
						if ((uiEditFlags & BCGP_EDIT_SIZE_LOCK_ASPECT_RATIO) == BCGP_EDIT_SIZE_LOCK_ASPECT_RATIO)
						{
							if (ptSize.x != 0.0)
							{
								ptSize.y = ptSize.x;
							}
							else if (ptSize.y != 0.0)
							{
								ptSize.x = ptSize.y;
							}
						}

						rect.SetSize (rect.Width () + ptSize.x, rect.Height () + ptSize.y);

						if ((uiEditFlags & BCGP_EDIT_NO_NORMALIZE_RECT) == 0)
						{
							rect.Normalize();
						}

						CBCGPSize sizeMin (pObject->GetMinSize());
						if (sizeMin.IsEmpty())
						{
							sizeMin.SetSize (1.0, 1.0);
						}

						if (fabs(rect.Width ()) < sizeMin.cx || fabs(rect.Height ()) < sizeMin.cy)
						{
							continue;
						}
					}
					else if ((uiEditFlags & BCGP_EDIT_NOMOVE) == 0)
					{
						rect.OffsetRect (ptOffset);
					}

					if (rect != pObject->GetRect())
					{
						pObject->SetTrackingRect (rect);
						pObject->SetRect (rect);
						bChanged = TRUE;
					}
				}

				if (!m_lstSel.IsEmpty())
				{
					FireObjectMoveEvent(TRUE, bChanged);
				}
			}

			if (bChanged)
			{
				Redraw();
			}
		}

		if ((m_dwEditFlags & BCGP_CONTAINER_DISABLE_KEYBOARD_OBJECT) == BCGP_CONTAINER_DISABLE_KEYBOARD_OBJECT)
		{
			return TRUE;
		}
	}

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);
			
			if (pObject->OnKeyboardDown(nChar, nRepCnt, nFlags))
			{
				return TRUE;
			}
		}
	}

	return FALSE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnKeyboardUp(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (IsEditMode ())
	{
		if (m_ptDragStart != CBCGPPoint(-1, -1) && m_nDragMode != HTNOWHERE)
		{
			return FALSE;
		}

		if ((m_dwEditFlags & BCGP_CONTAINER_DISABLE_KEYBOARD_OBJECT) == BCGP_CONTAINER_DISABLE_KEYBOARD_OBJECT)
		{
			return TRUE;
		}
	}

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);
			
			if (pObject->OnKeyboardUp(nChar, nRepCnt, nFlags))
			{
				return TRUE;
			}
		}
	}

	return FALSE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::ExportToImage(CBCGPGraphicsManager* pGMSrc, CBCGPImage& image, BOOL bFullImage)
{
	ASSERT_VALID(pGMSrc);

	CBCGPGraphicsManager* pGM = pGMSrc->CreateOffScreenManager(m_rect, &image);

	if (pGM == NULL)
	{
		return FALSE;
	}

	CBCGPRect rectSaved = m_rect;

	if (bFullImage)
	{
		CBCGPSize size = m_rect.Size();
		m_rect = CBCGPRect(0., 0., max(m_sizeScrollTotal.cx + 10, size.cx), max(m_sizeScrollTotal.cy + 10, size.cy));
	}
	else
	{
		m_rect = m_rect - rectSaved.TopLeft();
	}

	pGM->FillRectangle(m_rect, CBCGPBrush(CBCGPColor::White));

	SetDirty();
	OnDraw(pGM, CBCGPRect());

	m_rect = rectSaved;
	delete pGM;

	return TRUE;
}
//*******************************************************************************
HBITMAP CBCGPVisualContainer::ExportToBitmap(CBCGPGraphicsManager* pGM, BOOL bFullImage)
{
	ASSERT_VALID(pGM);

	CBCGPSize size = m_rect.Size();

	if (bFullImage)
	{
		size.cx = max(m_sizeScrollTotal.cx + 10, size.cx);
		size.cy = max(m_sizeScrollTotal.cy + 10, size.cy);
	}

	HBITMAP hmbpDib = CBCGPDrawManager::CreateBitmap_32(size, NULL);
	if (hmbpDib == NULL)
	{
		ASSERT (FALSE);
		return NULL;
	}

	CBCGPPoint ptScrollOffset = m_ptScrollOffset;

	if (bFullImage)
	{
		m_bDrawScrollBars = FALSE;
		m_ptScrollOffset.SetPoint(0., 0.);
	}

	CBCGPRect rect(CBCGPPoint(), size);

	CDC dcMem;
	dcMem.CreateCompatibleDC (NULL);

	HBITMAP hbmpOld = (HBITMAP) dcMem.SelectObject (hmbpDib);

	pGM->BindDC(&dcMem, rect);
	pGM->BeginDraw();

	CBCGPRect rectSaved = m_rect;

	if (bFullImage)
	{
		SetRect(CBCGPRect(CBCGPPoint(0., 0.), size));
	}
	else
	{
		SetRect(m_rect - rectSaved.TopLeft());
	}

	SetDirty();

	BOOL bCacheImage = m_bCacheImage;
	m_bCacheImage = FALSE;

	if (m_brFill.IsEmpty())
	{
		pGM->FillRectangle(m_rect, CBCGPBrush(CBCGPColor::White));
	}

	OnDraw(pGM, m_rect);

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPWndHostVisualObject* pObject = DYNAMIC_DOWNCAST(CBCGPWndHostVisualObject, m_arObjects[i]);
		if (pObject == NULL)
		{
			continue;
		}
		
		ASSERT_VALID(pObject);
		
		if (pObject->GetWnd()->GetSafeHwnd() != NULL)
		{
			CBitmap bmpScreenshot;
			if (globalUtils.CreateScreenshot(bmpScreenshot, pObject->GetWnd()))
			{
				CBCGPImage image((HBITMAP)bmpScreenshot.GetSafeHandle(), TRUE);
				pGM->DrawImage(image, pObject->GetRect().TopLeft());
			}
		}
	}

	pGM->EndDraw();

	dcMem.SelectObject (hbmpOld);

	CBCGPDrawManager::FillAlpha (rect, hmbpDib, 255);

	pGM->BindDC(NULL);

	SetRect(rectSaved);
	m_bCacheImage = bCacheImage;

	if (bFullImage)
	{
		m_bDrawScrollBars = TRUE;
		m_ptScrollOffset = ptScrollOffset;
	}

	SetDirty(TRUE);

	if (bFullImage)
	{
		Redraw();
	}

	return hmbpDib;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::CopyToClipboard(CBCGPGraphicsManager* pGM, BOOL bFullImage)
{
	if (pGM == NULL)
	{
		return FALSE;
	}

	HBITMAP hbmp = ExportToBitmap(pGM, bFullImage);

	CBCGPToolBarImages img;
	img.AddImage(hbmp, TRUE);

#ifndef _BCGSUITE_
	img.SetSingleImage(FALSE);
#else
	img.SetSingleImage();
#endif

	return img.CopyImageToClipboard(0);
}
//*******************************************************************************
BOOL CBCGPVisualContainer::ExportToFile(const CString& strFilePath, CBCGPGraphicsManager* pGM, BOOL bFullImage)
{
	if (pGM == NULL)
	{
		return FALSE;
	}

	return SaveBitmapToFile(ExportToBitmap(pGM, bFullImage), strFilePath);
}
//*******************************************************************************
void CBCGPVisualContainer::SetScaleRatio(const CBCGPSize& sizeScaleRatio)
{
	if (m_sizeScaleRatio == sizeScaleRatio)
	{
		return;
	}

	CBCGPSize sizeScaleRatioOld = m_sizeScaleRatio;
	m_sizeScaleRatio = sizeScaleRatio;

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject == NULL)
		{
			continue;
		}

		ASSERT_VALID(pObject);
		pObject->SetScaleRatio(m_sizeScaleRatio);
	}

	OnScaleRatioChanged(sizeScaleRatioOld);

	if (m_bScrollBars)
	{
		if (GetOwner()->GetSafeHwnd() != NULL && m_bNCScrollBars)
		{
			GetOwner()->SetWindowPos (NULL, -1, -1, -1, -1, 
				SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
		}

		AdjustScrollBars(FALSE);
	}
}
//*******************************************************************************
void CBCGPVisualContainer::EnableScalingByMouseWheel(BOOL bEnable, double dblMinScaleRatio/* = 0.1*/, double dblMaxScaleRatio/* = 4.0*/)
{
	m_bScaleByMouseWheel = bEnable;

	m_dblMinScaleRatio = max(0.01, dblMinScaleRatio);
	m_dblMaxScaleRatio = max(0.01, dblMaxScaleRatio);;
}
//*******************************************************************************
void CBCGPVisualContainer::OnScaleRatioChanged(const CBCGPSize& sizeScaleRatioOld)
{
	if (sizeScaleRatioOld.cx == 0.)
	{
		ASSERT(FALSE);
		return;
	}

	if (sizeScaleRatioOld.cy == 0.)
	{
		ASSERT(FALSE);
		return;
	}

	const double dblAspectRatioX = m_sizeScaleRatio.cx / sizeScaleRatioOld.cx;
	const double dblAspectRatioY = m_sizeScaleRatio.cy / sizeScaleRatioOld.cy;

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject == NULL)
		{
			continue;
		}

		ASSERT_VALID(pObject);

		if (m_sizeScaleRatio != CBCGPSize(1., 1.) || !pObject->ResetRect())
		{
			BOOL bDontResetRectSaved = pObject->m_bDontResetRect;
			pObject->m_bDontResetRect = TRUE;

			CBCGPRect rect = pObject->GetRect();
			rect.Scale (dblAspectRatioX, dblAspectRatioY);

			pObject->SetRect(rect);
			
			pObject->m_bDontResetRect = bDontResetRectSaved;
		}
	}
}
//*******************************************************************************
void CBCGPVisualContainer::FireSelectionChangedEvent()
{
	if (m_pWndOwner->GetSafeHwnd() == NULL)
	{
		return;
	}

#ifndef _BCGPCHART_STANDALONE
	if (globalData.IsAccessibilitySupport () && GetSelCount() > 0)
	{
		CBCGPBaseVisualObject* pSel = GetSel();
		if (pSel != NULL)
		{
#ifdef _BCGSUITE_
			::NotifyWinEvent(EVENT_OBJECT_FOCUS, m_pWndOwner->GetSafeHwnd(), OBJID_CLIENT, GetAccChildIndex(pSel));
#else
			globalData.NotifyWinEvent(EVENT_OBJECT_FOCUS, m_pWndOwner->GetSafeHwnd(), OBJID_CLIENT, GetAccChildIndex(pSel));
#endif
		}
	}
#endif

	CWnd* pWndOwner = m_pWndOwner->GetOwner();
	if (pWndOwner->GetSafeHwnd() == NULL)
	{
		return;
	}

	pWndOwner->SendMessage(BCGM_CONTAINER_SELCHANGED, 0, (LPARAM)this);
}
//*******************************************************************************
void CBCGPVisualContainer::FireObjectMoveEvent(BOOL bAfter, BOOL bHasMoved)
{
	if (m_pWndOwner->GetSafeHwnd() == NULL)
	{
		return;
	}

	CWnd* pWndOwner = m_pWndOwner->GetOwner();
	if (pWndOwner->GetSafeHwnd() == NULL)
	{
		return;
	}

	WPARAM wp = bAfter ? (WPARAM)bHasMoved : 0;

	pWndOwner->SendMessage(bAfter ? BCGM_CONTAINER_OBJECT_MOVED : BCGM_CONTAINER_BEFORE_OBJECT_MOVE, wp, (LPARAM)this);
}
//*******************************************************************************
UINT CBCGPVisualContainer::GetClickAndHoldID() const
{
	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject == NULL)
		{
			continue;
		}

		ASSERT_VALID(pObject);

		if (!pObject->GetClickAndHoldRect().IsRectEmpty())
		{
			return pObject->GetClickAndHoldID();
		}
	}

	return m_nClickAndHoldID;
}
//*******************************************************************************
CBCGPRect CBCGPVisualContainer::GetClickAndHoldRect()
{
	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject == NULL)
		{
			continue;
		}

		ASSERT_VALID(pObject);

		if (!pObject->GetClickAndHoldRect().IsRectEmpty())
		{
			return pObject->GetClickAndHoldRect();
		}
	}

	return m_rectClickAndHold;
}
//*******************************************************************************
void CBCGPVisualContainer::OnClickAndHoldEvent(UINT nID, const CBCGPPoint& point)
{
	if (m_ScrollBarVert.GetPrevButton().PtInRect(point))
	{
		m_ScrollBarVert.OnScrollStep(FALSE);
		return;
	}
	else if (m_ScrollBarVert.GetNextButton().PtInRect(point))
	{
		m_ScrollBarVert.OnScrollStep(TRUE);
		return;
	}
	else if (m_ScrollBarHorz.GetPrevButton().PtInRect(point))
	{
		m_ScrollBarHorz.OnScrollStep(FALSE);
		return;
	}
	else if (m_ScrollBarHorz.GetNextButton().PtInRect(point))
	{
		m_ScrollBarHorz.OnScrollStep(TRUE);
		return;
	}

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject == NULL)
		{
			continue;
		}

		ASSERT_VALID(pObject);

		if (!pObject->GetClickAndHoldRect().IsRectEmpty())
		{
			pObject->OnClickAndHoldEvent(nID, point);
		}
	}
}
//*******************************************************************************
void CBCGPVisualContainer::OnScroll(CBCGPVisualScrollBar* pScrollBar, double dblDelta)
{
	if (pScrollBar == &m_ScrollBarVert)
	{
		m_ptScrollOffset.y += dblDelta;
	}
	else if (pScrollBar == &m_ScrollBarHorz)
	{
		m_ptScrollOffset.x += dblDelta;
	}

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject != NULL)
		{
			ASSERT_VALID(pObject);
			pObject->SetScrollOffset(m_ptScrollOffset);
		}
	}

	SetDirty();
	Redraw();
}
//*******************************************************************************
void CBCGPVisualContainer::SetClickAndHoldEvent(const CBCGPRect& rectBounds, UINT nEventID)
{
	m_rectClickAndHold = rectBounds;
	m_nClickAndHoldID = nEventID;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnGestureEventZoom(const CBCGPPoint& ptCenter, double dblZoomFactor)
{
	CBCGPBaseVisualObject* pObject = GetFromPoint(ptCenter);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		if (pObject->OnGestureEventZoom(ptCenter, dblZoomFactor))
		{
			return TRUE;
		}
	}

	return FALSE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnGestureEventPan(const CBCGPPoint& ptFrom, const CBCGPPoint& ptTo, CBCGPSize& sizeOverPan)
{
	CBCGPBaseVisualObject* pObject = GetFromPoint(ptFrom);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		if (pObject->OnGestureEventPan(ptFrom, ptTo, sizeOverPan))
		{
			return TRUE;
		}
	}

	if (!m_ScrollBarVert.GetRect().IsRectEmpty() || !m_ScrollBarHorz.GetRect().IsRectEmpty())
	{
		if (m_ScrollBarVert.DoGestureEventPan(ptFrom, ptTo))
		{
			return TRUE;
		}

		if (m_ScrollBarHorz.DoGestureEventPan(ptFrom, ptTo))
		{
			return TRUE;
		}

		const BOOL bScrollVert = !m_ScrollBarVert.GetRect().IsRectEmpty();
		CBCGPVisualScrollBar& scrollBar = bScrollVert ? m_ScrollBarVert : m_ScrollBarHorz;

		double dblDelta = bScrollVert ? ptTo.y - ptFrom.y : ptTo.x - ptFrom.x;
		double dblOffsetOld = scrollBar.GetOffset();

		double dblOffsetNew = dblOffsetOld - dblDelta;
		double dblOffsetMax = scrollBar.GetTotal() - scrollBar.GetScrollSize();
		double dblOverPan = 0.;

		if (dblOffsetNew < 0.)
		{
			dblOverPan = -dblOffsetNew;
			dblOffsetNew = 0.;
			dblDelta = dblOffsetOld;
		}
		else if (dblOffsetNew > dblOffsetMax)
		{
			dblOverPan = dblOffsetMax - dblOffsetNew;
			dblOffsetNew = dblOffsetMax;
			dblDelta = dblOffsetOld - dblOffsetNew;
		}

		if (dblOverPan != 0.)
		{
			if (bScrollVert)
			{
				sizeOverPan.cy = dblOverPan;
			}
			else
			{
				sizeOverPan.cx = dblOverPan;
			}
		}

		if (dblOffsetNew != scrollBar.GetOffset())
		{
			scrollBar.SetOffset(dblOffsetNew);
			scrollBar.ReposThumb();

			OnScroll(&scrollBar, -dblDelta);
		}

		return TRUE;
	}

	return FALSE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnGestureEventRotate(const CBCGPPoint& ptCenter, double dblAngle)
{
	CBCGPBaseVisualObject* pObject = GetFromPoint(ptCenter);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		if (pObject->OnGestureEventRotate(ptCenter, dblAngle))
		{
			return TRUE;
		}
	}
	
	return FALSE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnGestureEventTwoFingerTap(const CBCGPPoint& ptCenter)
{
	CBCGPBaseVisualObject* pObject = GetFromPoint(ptCenter);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		if (pObject->OnGestureEventTwoFingerTap(ptCenter))
		{
			return TRUE;
		}
	}
	
	return FALSE;
}
//*******************************************************************************
BOOL CBCGPVisualContainer::OnGestureEventPressAndTap(const CBCGPPoint& ptPress, long lDelta)
{
	CBCGPBaseVisualObject* pObject = GetFromPoint(ptPress);
	if (pObject != NULL)
	{
		ASSERT_VALID(pObject);
		if (pObject->OnGestureEventPressAndTap(ptPress, lDelta))
		{
			return TRUE;
		}
	}
	
	return FALSE;
}
//*******************************************************************************
void CBCGPVisualContainer::ResetGestureOptions()
{
	ASSERT_VALID (this);

	if (m_pWndOwner->GetSafeHwnd() != NULL)
	{
		CBCGPGestureConfig gestureConfig;
		if (GetGestureConfig(gestureConfig))
		{
			bcgpGestureManager.SetGestureConfig(m_pWndOwner->GetSafeHwnd(), gestureConfig);
		}
	}
}
//*****************************************************************************
CBCGPBaseVisualObject* CBCGPVisualContainer::GetAccChild(int nIndex)
{
	ASSERT_VALID(this);

	nIndex--;

	if (nIndex < 0 || nIndex >= (int)m_arObjects.GetSize())
	{
		return NULL;
	}

	return m_arObjects[nIndex];
}
//*****************************************************************************
long CBCGPVisualContainer::GetAccChildIndex(CBCGPBaseVisualObject* pObject)
{
	ASSERT_VALID(this);
	ASSERT_VALID(pObject);

	if (!pObject->IsVisible())
	{
		return 0;
	}

	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pCurrObject = m_arObjects[i];
		ASSERT_VALID(pCurrObject);

		if (pCurrObject == pObject)
		{
			return i + 1;
		}
	}

	return 0;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::get_accChildCount(long *pcountChildren)
{
	if (!pcountChildren)
	{
		return E_INVALIDARG;
	}

	*pcountChildren = (long)m_arObjects.GetSize();
	return S_OK;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::get_accChild(VARIANT /*varChild*/, IDispatch** /*ppdispChild*/)
{
	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::get_accName(VARIANT varChild, BSTR *pszName)
{
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		CString strText;

		if (m_pWndOwner->GetSafeHwnd() != NULL)
		{
			m_pWndOwner->GetWindowText(strText);

			if (strText.IsEmpty())
			{
				strText = _T("Visual Container");
			}
		}

		*pszName = strText.AllocSysString();
		return S_OK;
	}

	if (varChild.vt == VT_I4)
	{
		CBCGPBaseVisualObject* pAccObject = GetAccChild(varChild.lVal);
		if (pAccObject != NULL)
		{
			ASSERT_VALID(pAccObject);

			varChild.lVal = CHILDID_SELF;
			return pAccObject->get_accName(varChild, pszName);
		}
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::get_accValue(VARIANT varChild, BSTR *pszValue)
{
	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		return S_FALSE;
	}

	if (varChild.vt == VT_I4)
	{
		CBCGPBaseVisualObject* pAccObject = GetAccChild(varChild.lVal);
		if (pAccObject != NULL)
		{
			ASSERT_VALID(pAccObject);
			
			varChild.lVal = CHILDID_SELF;
			return pAccObject->get_accValue(varChild, pszValue);
		}
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::get_accDescription(VARIANT varChild, BSTR *pszDescription)
{
	if (((varChild.vt != VT_I4) && (varChild.lVal != CHILDID_SELF)) || (NULL == pszDescription))
	{
		return E_INVALIDARG;
	}

	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		*pszDescription = SysAllocString(L"WinUITiles Table");
		return S_OK;
	}

	CBCGPBaseVisualObject* pAccObject = GetAccChild(varChild.lVal);
	if (pAccObject != NULL)
	{
		ASSERT_VALID(pAccObject);
		
		varChild.lVal = CHILDID_SELF;
		return pAccObject->get_accDescription(varChild, pszDescription);
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::get_accRole(VARIANT varChild, VARIANT *pvarRole)
{
	if (!pvarRole || ((varChild.vt != VT_I4) && (varChild.lVal != CHILDID_SELF)))
	{
		return E_INVALIDARG;
	}

	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		pvarRole->vt = VT_I4;
		pvarRole->lVal = ROLE_SYSTEM_GROUPING;

		return S_OK;
	}

	CBCGPBaseVisualObject* pAccObject = GetAccChild(varChild.lVal);
	if (pAccObject != NULL)
	{
		ASSERT_VALID(pAccObject);
		
		varChild.lVal = CHILDID_SELF;
		return pAccObject->get_accRole(varChild, pvarRole);
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::get_accState(VARIANT varChild, VARIANT *pvarState)
{
	if (pvarState == NULL)
	{
		return E_INVALIDARG;
	}

	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		pvarState->vt = VT_I4;
		pvarState->lVal = STATE_SYSTEM_NORMAL;
		return S_OK;
	}

	if (varChild.vt == VT_I4)
	{
		CBCGPBaseVisualObject* pAccObject = GetAccChild(varChild.lVal);
		if (pAccObject != NULL)
		{
			ASSERT_VALID(pAccObject);
			
			varChild.lVal = CHILDID_SELF;
			return pAccObject->get_accState(varChild, pvarState);
		}
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::get_accDefaultAction(VARIANT varChild, BSTR* pszDefaultAction)
{
	if (varChild.vt == VT_I4 && varChild.lVal > 0)
	{
		CBCGPBaseVisualObject* pAccObject = GetAccChild(varChild.lVal);
		if (pAccObject != NULL)
		{
			ASSERT_VALID(pAccObject);
			
			varChild.lVal = CHILDID_SELF;
			return pAccObject->get_accDefaultAction(varChild, pszDefaultAction);
		}
	}
	
	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::accSelect(long flagsSelect, VARIANT varChild)
{
	if (varChild.vt != VT_I4)
	{
		return E_INVALIDARG;
	}

	if (flagsSelect == SELFLAG_NONE || varChild.lVal == CHILDID_SELF)
	{
		return S_FALSE;
	}

	if (flagsSelect == SELFLAG_TAKEFOCUS || flagsSelect == SELFLAG_TAKESELECTION)
	{
		CBCGPBaseVisualObject* pAccObject = GetAccChild(varChild.lVal);
		if (pAccObject != NULL)
		{
			Select(pAccObject);
			return S_OK;
		}
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::accLocation(long *pxLeft, long *pyTop, long *pcxWidth, long *pcyHeight, VARIANT varChild)
{
	if (!pxLeft || !pyTop || !pcxWidth || !pcyHeight || varChild.vt != VT_I4)
	{
		return E_INVALIDARG;
	}

	if (m_pWndOwner->GetSafeHwnd() == NULL)
	{
		return S_FALSE;
	}

	if ((varChild.vt == VT_I4) && (varChild.lVal == CHILDID_SELF))
	{
		CRect rc;
		m_pWndOwner->GetWindowRect(rc);

		*pxLeft = rc.left;
		*pyTop = rc.top;
		*pcxWidth = rc.Width();
		*pcyHeight = rc.Height();

		return S_OK;
	}

	CBCGPBaseVisualObject* pAccObject = GetAccChild(varChild.lVal);
	if (pAccObject != NULL)
	{
		ASSERT_VALID(pAccObject);
		
		varChild.lVal = CHILDID_SELF;
		return pAccObject->accLocation(pxLeft, pyTop, pcxWidth, pcyHeight, varChild);
	}

	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::get_accParent(IDispatch **ppdispParent)
{
    if (ppdispParent == NULL)
	{
		return E_INVALIDARG;
	}
	
	*ppdispParent = NULL;
	
	if (m_pWndOwner->GetSafeHwnd() != NULL)
	{
		CWnd* pWndParent = m_pWndOwner->GetParent();
		if (pWndParent->GetSafeHwnd() != NULL)
		{
			return AccessibleObjectFromWindow(pWndParent->GetSafeHwnd(), (DWORD)OBJID_CLIENT,
				IID_IAccessible, (void**)ppdispParent);
		}
	}
	
	return S_FALSE;
}
//****************************************************************************
HRESULT CBCGPVisualContainer::accHitTest(long  xLeft, long yTop, VARIANT *pvarChild)
{
	if (!pvarChild)
	{
		return E_INVALIDARG;
	}

	if (m_pWndOwner->GetSafeHwnd() == NULL)
	{
		return S_FALSE;
	}

	pvarChild->vt = VT_I4;

	CPoint pt(xLeft, yTop);
	m_pWndOwner->ScreenToClient(&pt);

	CBCGPBaseVisualObject* pObject = GetFromPoint(pt);
	if (pObject != NULL)
	{
		pvarChild->lVal = GetAccChildIndex(pObject);
	}
	else
	{
		pvarChild->lVal = CHILDID_SELF;
	}

	return S_OK;
}
//******************************************************************************
HRESULT CBCGPVisualContainer::accDoDefaultAction(VARIANT varChild)
{
    if (varChild.vt != VT_I4)
    {
        return E_INVALIDARG;
    }

	CBCGPBaseVisualObject* pAccObject = GetAccChild(varChild.lVal);
	if (pAccObject != NULL)
	{
/*		if (OnClick(pAccObject))
		{
			return S_OK;
		}*/
    }

    return S_FALSE;
}
//******************************************************************************
void CBCGPVisualContainer::OnCalcBorderSize(CBCGPRect& rectNCArea)
{
	if (!m_bNCScrollBars || !m_bScrollBars)
	{
		return;
	}

	AdjustScrollBars(FALSE);

	double cxScroll = m_ScrollBarVert.GetRect().Width();
	double cyScroll = m_ScrollBarHorz.GetRect().Height();

	rectNCArea.left = 1;
	rectNCArea.top = 1;
	rectNCArea.right = cxScroll == 0 ? 1 : cxScroll + 2;
	rectNCArea.bottom = cyScroll == 0 ? 1 : cyScroll + 2;
}
//******************************************************************************
void CBCGPVisualContainer::OnNcDraw(CBCGPGraphicsManager* pGM, const CBCGPRect& rectIn)
{
	if (!m_bNCScrollBars || !m_bScrollBars)
	{
		return;
	}

	ASSERT_VALID(pGM);

	if (!m_rectBottomRight.IsRectEmpty())
	{
		pGM->FillRectangle(m_rectBottomRight, m_ScrollBarVert.GetColorTheme().m_brFace);
	}
	
	m_ScrollBarVert.DoDraw(pGM);
	m_ScrollBarHorz.DoDraw(pGM);

	CBCGPRect rect(rectIn);
	if (pGM->GetType() != CBCGPGraphicsManager::BCGP_GRAPHICS_MANAGER_GDI)
	{
		rect.right -= 1;
		rect.bottom -= 1;
	}

	if (!m_brOutline.IsEmpty() && m_brOutline.GetOpacity() == 1.0)
	{
		pGM->DrawRectangle(rect, m_brOutline, 1);
	}
	else
	{
		CBCGPColor color = CBCGPColor::White;
		
		if (!m_brOutline.IsEmpty())
		{
			color = m_brOutline.GetColor();
		}
		else if (!m_brFill.IsEmpty())
		{
			color = m_brFill.GetColor();
		}
		
		color.a = 1.0;
		
		pGM->DrawRectangle(rect, CBCGPBrush(color, 1.0), 1);
	}
}
//******************************************************************************
BOOL CBCGPVisualContainer::OnNcMouseDown(int nButton, const CBCGPPoint& pt)
{
	if (nButton == 0 && m_bNCScrollBars)
	{
		if (m_ScrollBarVert.OnMouseDown(pt) || m_ScrollBarHorz.OnMouseDown(pt))
		{
			return TRUE;
		}
	}

	return FALSE;
}
//******************************************************************************
void CBCGPVisualContainer::OnNcMouseUp(int /*nButton*/, const CBCGPPoint& /*pt*/)
{
	BOOL bNeedToRedraw = m_ScrollBarVert.IsInAction() || m_ScrollBarHorz.IsInAction();

	m_ScrollBarVert.OnCancelMode();
	m_ScrollBarHorz.OnCancelMode();

	if (bNeedToRedraw)
	{
		Redraw();
	}
}
//*******************************************************************************
void CBCGPVisualContainer::OnWndEnabled(BOOL bEnable)
{
	for (int i = 0; i < m_arObjects.GetSize(); i++)
	{
		CBCGPBaseVisualObject* pObject = m_arObjects[i];
		if (pObject == NULL)
		{
			continue;
		}
		
		ASSERT_VALID(pObject);
		pObject->m_bIsEnabled = bEnable;
	}
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPVisualScrollBar

static const UINT nIDScrollPrev = 1001;
static const UINT nIDScrollNext = 1002;

CBCGPVisualScrollBarColorTheme::CBCGPVisualScrollBarColorTheme(const CBCGPColor& color)
{
	if (color == CBCGPColor())
	{
		// Set default colors:
		m_brFace = CBCGPBrush(CBCGPColor::WhiteSmoke);
		m_brBorder = CBCGPBrush(CBCGPColor::Gray);
		m_brButton = CBCGPBrush(CBCGPColor::Gainsboro);
		m_brButtonPressed = CBCGPBrush(CBCGPColor::DarkGray);
		m_brButtonHighlighted = CBCGPBrush(CBCGPColor::Silver);
	}
	else
	{
		CBCGPColor colorFace = color;
		CBCGPColor colorBorder = color;
		CBCGPColor colorButton = color;
		CBCGPColor colorPressed = color;
		CBCGPColor colorHighlighted = color;

		if (color.IsPale())
		{
			colorBorder.MakeDarker(.3);
			colorButton.MakeDarker(.05);
			colorPressed = colorBorder;
			colorHighlighted.MakeDarker(.02);
		}
		else if (color.IsDark())
		{
			colorBorder.MakeLighter(.4);
			colorButton.MakeLighter(.2);
			colorPressed.MakeLighter(.3);
			colorHighlighted.MakeLighter(.25);
		}
		else
		{
			colorFace.MakeLighter(.15);
			colorBorder.MakeDarker(.2);
			colorButton.MakeLighter(.03);
			colorPressed.MakeDarker();
			colorHighlighted.MakeLighter(.1);
		}

		m_brFace = CBCGPBrush(colorFace);
		m_brBorder = CBCGPBrush(colorBorder);
		m_brButton = CBCGPBrush(colorButton);
		m_brButtonPressed = CBCGPBrush(colorPressed);
		m_brButtonHighlighted = CBCGPBrush(colorHighlighted);
	}
}
//*******************************************************************************
CBCGPVisualScrollBar::CBCGPVisualScrollBar()
{
	m_pOwner = NULL;
	m_pOwnerContainer = NULL;

	m_bIsHorizontal = TRUE;
	m_dblStep = 0.0;

	m_Style = BCGP_VISUAL_SCROLLBAR_FLAT;

	Reset();
}
//*******************************************************************************
CBCGPVisualScrollBar::~CBCGPVisualScrollBar()
{
	if (m_hBitmap != NULL)
	{
		::DeleteObject(m_hBitmap);
		m_hBitmap = NULL;
	}
}
//*******************************************************************************
void CBCGPVisualScrollBar::SetParentVisualObject(CBCGPBaseVisualObject* pOwner)
{
	m_pOwner = pOwner;
}
//*******************************************************************************
void CBCGPVisualScrollBar::SetParentVisualContainer(CBCGPVisualContainer* pOwnerContainer)
{
	m_pOwnerContainer = pOwnerContainer;
}
//*******************************************************************************
void CBCGPVisualScrollBar::Redraw(CBCGPRect* pRect)
{
	if (pRect == NULL)
	{
		if (m_rect.IsRectEmpty())
		{
			return;
		}

		pRect = &m_rect;
	}

	if (m_pOwner != NULL)
	{
		m_pOwner->RedrawRect(*pRect);
	}
	else if (m_pOwnerContainer != NULL)
	{
		m_pOwnerContainer->RedrawRect(*pRect);
	}
}
//*******************************************************************************
BOOL CBCGPVisualScrollBar::OnMouseDown(const CBCGPPoint& pt)
{
	if (m_dblTotal <= 0. || m_rect.IsRectEmpty())
	{
		return FALSE;
	}

	if (m_rectThumb.PtInRect(pt))
	{
		m_bIsDraggingThumb = TRUE;
		m_ptDragThumbLast = pt;
		Redraw();

		return TRUE;
	}

	double size = m_bIsHorizontal ? m_rectScroll.Width() : m_rectScroll.Height();
	double dblOffsetOld = m_dblOffset;
	double dblOffset = m_dblOffset;

	if (m_rectPrev.PtInRect(pt))
	{
		OnScrollStep(FALSE);

		if (m_pOwner != NULL)
		{
			m_pOwner->SetClickAndHoldEvent(m_rectPrev, nIDScrollPrev);
		}
		else if (m_pOwnerContainer != NULL)
		{
			m_pOwnerContainer->SetClickAndHoldEvent(m_rectPrev, nIDScrollPrev);
		}

		m_bPrevIsClicked = TRUE;
		Redraw();

		return TRUE;
	}
			
	if (m_rectNext.PtInRect(pt))
	{
		OnScrollStep(TRUE);

		if (m_pOwner != NULL)
		{
			m_pOwner->SetClickAndHoldEvent(m_rectNext, nIDScrollNext);
		}
		else if (m_pOwnerContainer != NULL)
		{
			m_pOwnerContainer->SetClickAndHoldEvent(m_rectNext, nIDScrollNext);
		}

		m_bNextIsClicked = TRUE;
		Redraw();

		return TRUE;
	}
	
	if (!m_rectScroll.PtInRect(pt))
	{
		return FALSE;
	}

	if (m_bIsHorizontal)
	{
		if (pt.x < m_rectThumb.left)
		{
			dblOffset -= size;
		}
		else
		{
			dblOffset += size;
		}
	}
	else
	{
		if (pt.y < m_rectThumb.top)
		{
			dblOffset -= size;
		}
		else
		{
			dblOffset += size;
		}
	}

	m_dblOffset = min(max(0, dblOffset), m_dblTotal - size); 
	ReposThumb();

	if (m_pOwner != NULL)
	{
		m_pOwner->OnScroll(this, m_dblOffset - dblOffsetOld);
	}
	else if (m_pOwnerContainer != NULL)
	{
		m_pOwnerContainer->OnScroll(this, m_dblOffset - dblOffsetOld);
	}

	return TRUE;
}
//*******************************************************************************
BOOL CBCGPVisualScrollBar::OnMouseMove(const CBCGPPoint& pt)
{
	if (m_rect.IsRectEmpty())
	{
		return FALSE;
	}

	BOOL bIsThumbHighlighted = m_rectThumb.PtInRect(pt);

	if (m_bIsThumbHighlighted != bIsThumbHighlighted)
	{
		m_bIsThumbHighlighted = bIsThumbHighlighted;
		Redraw(&m_rectThumb);

		return TRUE;
	}

	BOOL bIsPrevHighlighted = m_rectPrev.PtInRect(pt);
	
	if (m_bIsPrevHighlighted != bIsPrevHighlighted)
	{
		m_bIsPrevHighlighted = bIsPrevHighlighted;
		Redraw(&m_rectPrev);

		return TRUE;
	}

	BOOL bIsNextHighlighted = m_rectNext.PtInRect(pt);
	
	if (m_bIsNextHighlighted != bIsNextHighlighted)
	{
		m_bIsNextHighlighted = bIsNextHighlighted;
		Redraw(&m_rectNext);

		return TRUE;
	}

	return FALSE;
}
//*******************************************************************************
void CBCGPVisualScrollBar::OnCancelMode()
{
	m_bIsDraggingThumb = FALSE;
	m_bPrevIsClicked = FALSE;
	m_bNextIsClicked = FALSE;

	m_bIsThumbHighlighted = FALSE;
	m_bIsPrevHighlighted = FALSE;
	m_bIsNextHighlighted = FALSE;

	Redraw();
}
//*******************************************************************************
void CBCGPVisualScrollBar::OnMouseLeave()
{
	if (m_bIsThumbHighlighted)
	{
		m_bIsThumbHighlighted = FALSE;
		Redraw(&m_rect);
	}

	if (m_bIsPrevHighlighted)
	{
		m_bIsPrevHighlighted = FALSE;
		Redraw(&m_rectPrev);
	}

	if (m_bIsNextHighlighted)
	{
		m_bIsNextHighlighted = FALSE;
		Redraw(&m_rectNext);
	}
}
//*******************************************************************************
void CBCGPVisualScrollBar::ReposThumb()
{
	if (m_rect.IsRectEmpty())
	{
		Reset();
		return;
	}

	m_rectThumb = m_rectScroll;

	double size = m_bIsHorizontal ? m_rectScroll.Width() : m_rectScroll.Height();

	double dblScrollRatio = size / m_dblTotal;
	double dblThumbSize = max(globalUtils.ScaleByDPI(7.0), dblScrollRatio * size);

	if (m_bIsHorizontal)
	{
		m_rectThumb.left += m_dblOffset * dblScrollRatio;
		m_rectThumb.right = m_rectThumb.left + dblThumbSize;

		if (m_rectThumb.right > m_rectScroll.right)
		{
			m_rectThumb.right = m_rectScroll.right;
			m_rectThumb.left = m_rectThumb.right - dblThumbSize;
		}
	}
	else
	{
		m_rectThumb.top += m_dblOffset * dblScrollRatio;
		m_rectThumb.bottom = m_rectThumb.top + dblThumbSize;

		if (m_rectThumb.bottom > m_rectScroll.bottom)
		{
			m_rectThumb.bottom = m_rectScroll.bottom;
			m_rectThumb.top = m_rectThumb.bottom - dblThumbSize;
		}
	}
}
//*******************************************************************************
BOOL CBCGPVisualScrollBar::DoGestureEventPan(const CBCGPPoint& ptFrom, const CBCGPPoint& ptTo)
{
	if (!m_rect.PtInRect(ptFrom))
	{
		return FALSE;
	}

	if (m_rectThumb.PtInRect(ptFrom))
	{
		m_bIsDraggingThumb = TRUE;
		m_ptDragThumbLast = ptFrom;

		OnDragThumb(ptTo);

		m_bIsDraggingThumb = FALSE;
	}

	return TRUE;
}
//*******************************************************************************
void CBCGPVisualScrollBar::SetHorizontal(BOOL bSet)
{
	m_bIsHorizontal = bSet;
}
//*******************************************************************************
void CBCGPVisualScrollBar::SetRect(const CBCGPRect& rectIn)
{
	if (rectIn.IsRectEmpty())
	{
		Reset();
		return;
	}

	if (m_hBitmap != NULL)
	{
		::DeleteObject(m_hBitmap);
		m_hBitmap = NULL;
	}

	CBCGPRect rect = rectIn;
	rect.DeflateRect(0.5, 0.5);

	m_rect = m_rectScroll = m_rectNext = m_rectPrev = rect;

	if (m_bIsHorizontal)
	{
		m_rectPrev.right = m_rectPrev.left + m_rectPrev.Height();
		m_rectNext.left = m_rectNext.right - m_rectNext.Height();

		m_rectScroll.DeflateRect(m_rectScroll.Height(), 0);
	}
	else
	{
		m_rectPrev.bottom = m_rectPrev.top + m_rectPrev.Width();
		m_rectNext.top = m_rectNext.bottom - m_rectNext.Width();

		m_rectScroll.DeflateRect(0, m_rectScroll.Width());
	}

	ReposThumb();
}
//*******************************************************************************
void CBCGPVisualScrollBar::Reset()
{
	m_bIsDraggingThumb = FALSE;
	m_bPrevIsClicked = FALSE;
	m_bNextIsClicked = FALSE;

	m_bIsThumbHighlighted = FALSE;
	m_bIsPrevHighlighted = FALSE;
	m_bIsNextHighlighted = FALSE;

	m_dblTotal = 0.;
	m_dblOffset = 0.;

	m_rect.SetRectEmpty();
	m_rectScroll.SetRectEmpty();
	m_rectNext.SetRectEmpty();
	m_rectPrev.SetRectEmpty();
	m_rectThumb.SetRectEmpty();

	if (m_hBitmap != NULL)
	{
		::DeleteObject(m_hBitmap);
		m_hBitmap = NULL;
	}
}
//*******************************************************************************
void CBCGPVisualScrollBar::DoDraw(CBCGPGraphicsManager* pGM)
{
	if (m_rect.IsRectEmpty())
	{
		return;
	}

	if (m_Style == BCGP_VISUAL_SCROLLBAR_VISUAL_MANAGER && globalData.Is32BitIcons()
#ifdef _BCGSUITE_
		&& CBCGPVisualManager::GetInstance()->IsOwnerDrawScrollBar()
#endif
	)
	{
		CBCGPScrollBar dummy;

		CRect rect = m_rect;
		CPoint ptOffset = -rect.TopLeft();
		
		rect.OffsetRect(ptOffset);

		CRect rectButton1 = m_rectPrev;
		rectButton1.OffsetRect(ptOffset);

		CRect rectButton2 = m_rectNext;
		rectButton2.OffsetRect(ptOffset);

		CRect rectThumb = m_rectThumb;
		rectThumb.OffsetRect(ptOffset);

		CDC dc;
		if (dc.CreateCompatibleDC(NULL))
		{
			if (m_hBitmap == NULL)
			{
				m_hBitmap = CBCGPDrawManager::CreateBitmap_32(rect.Size(), NULL);
			}
			
			if (m_hBitmap != NULL)
			{
				HBITMAP hBitmapOld = (HBITMAP)dc.SelectObject(m_hBitmap);
				if (hBitmapOld != NULL)
				{
					CBCGPVisualManager::GetInstance ()->OnScrollBarFillBackground(&dc, &dummy, rect, m_bIsHorizontal, FALSE, FALSE, TRUE, FALSE);
					CBCGPVisualManager::GetInstance ()->OnScrollBarDrawButton(&dc, &dummy, rectButton1, m_bIsHorizontal, m_bIsPrevHighlighted, m_bIsPrevHighlighted && m_bPrevIsClicked, TRUE, FALSE);
					CBCGPVisualManager::GetInstance ()->OnScrollBarDrawButton(&dc, &dummy, rectButton2, m_bIsHorizontal, m_bIsNextHighlighted, m_bIsNextHighlighted && m_bNextIsClicked, FALSE, FALSE);
					CBCGPVisualManager::GetInstance ()->OnScrollBarDrawThumb(&dc, &dummy, rectThumb, m_bIsHorizontal, m_bIsThumbHighlighted, m_bIsDraggingThumb, FALSE);

					dc.SelectObject(hBitmapOld);

					pGM->DrawImage(CBCGPImage(m_hBitmap), m_rect.TopLeft());
					return;
				}					
			}
		}
	}

	double dblLineWidth = 2.0;
	
	BOOL bRoundedCorners = m_Style == BCGP_VISUAL_SCROLLBAR_FLAT_ROUNDED || m_Style == BCGP_VISUAL_SCROLLBAR_3D_ROUNDED;
	BOOL bIs3D = m_Style == BCGP_VISUAL_SCROLLBAR_3D || m_Style == BCGP_VISUAL_SCROLLBAR_3D_ROUNDED;

	if (bIs3D)
	{
		if (m_ColorTheme.m_brFace3D.IsEmpty())
		{
			CBCGPColor color1 = m_ColorTheme.m_brFace.GetColor();
			CBCGPColor color2 = color1;
			
			color2.IsDark() ? color2.MakeLighter() : color2.MakeDarker();

			if (m_bIsHorizontal)
			{
				color1 = color2;
				color2 = m_ColorTheme.m_brFace.GetColor();
			}

			m_ColorTheme.m_brFace3D = CBCGPBrush(color1, color2, m_bIsHorizontal ? CBCGPBrush::BCGP_GRADIENT_HORIZONTAL : CBCGPBrush::BCGP_GRADIENT_VERTICAL);
		}

		if (m_ColorTheme.m_brButton3D.IsEmpty())
		{
			CBCGPColor color1 = m_ColorTheme.m_brButton.GetColor();
			CBCGPColor color2 = color1;
			color1.MakeDarker();
			color2.MakePale();
			
			m_ColorTheme.m_brButton3D = CBCGPBrush(color1, color2, m_bIsHorizontal ? CBCGPBrush::BCGP_GRADIENT_PIPE_HORIZONTAL : CBCGPBrush::BCGP_GRADIENT_PIPE_VERTICAL);
		}

		if (m_ColorTheme.m_brButtonPressed3D.IsEmpty())
		{
			m_ColorTheme.m_brButtonPressed3D = m_ColorTheme.m_brButton3D;

			if (m_ColorTheme.m_brButtonPressed3D.GetColor().IsDark())
			{
				m_ColorTheme.m_brButtonPressed3D.MakeLighter();
			}
			else
			{
				m_ColorTheme.m_brButtonPressed3D.MakeDarker(.07);
			}
		}

		if (m_ColorTheme.m_brButtonHighlighted3D.IsEmpty())
		{
			m_ColorTheme.m_brButtonHighlighted3D = m_ColorTheme.m_brButton3D;
			
			if (m_ColorTheme.m_brButtonHighlighted3D.GetColor().IsDark())
			{
				m_ColorTheme.m_brButtonHighlighted3D.MakeLighter(.07);
			}
			else
			{
				m_ColorTheme.m_brButtonHighlighted3D.MakeDarker(.05);
			}
		}
	}

	pGM->FillRectangle(m_rect, bIs3D ? m_ColorTheme.m_brFace3D : m_ColorTheme.m_brFace);
	pGM->DrawRectangle(m_rect, m_ColorTheme.m_brBorder);

	const CBCGPBrush& brThumb = bIs3D ? 
		(m_bIsDraggingThumb ? m_ColorTheme.m_brButtonPressed3D : (m_bIsThumbHighlighted ? m_ColorTheme.m_brButtonHighlighted3D : m_ColorTheme.m_brButton3D)) : 
		(m_bIsDraggingThumb ? m_ColorTheme.m_brButtonPressed : (m_bIsThumbHighlighted ? m_ColorTheme.m_brButtonHighlighted : m_ColorTheme.m_brButton));

	if (bRoundedCorners)
	{
		double dblCornerRadius = min(m_rect.Width(), m_rect.Height()) / 3;

		CBCGPRect rectTumb = m_rectThumb;
		rectTumb.DeflateRect(0.5 * dblCornerRadius, 0.5 * dblCornerRadius);

		CBCGPRoundedRect rrThumb(rectTumb, dblCornerRadius, dblCornerRadius);

		pGM->FillRoundedRectangle(rrThumb, brThumb);
		pGM->DrawRoundedRectangle(rrThumb, m_ColorTheme.m_brBorder);
	}
	else
	{
		pGM->FillRectangle(m_rectThumb, brThumb);
		pGM->DrawRectangle(m_rectThumb, m_ColorTheme.m_brBorder);
	}

	const CBCGPBrush& brPrev = bIs3D ? 
		(m_bIsPrevHighlighted && m_bPrevIsClicked ? m_ColorTheme.m_brButtonPressed3D : (m_bIsPrevHighlighted ? m_ColorTheme.m_brButtonHighlighted3D : m_ColorTheme.m_brButton3D)) : 
		(m_bIsPrevHighlighted && m_bPrevIsClicked ? m_ColorTheme.m_brButtonPressed : (m_bIsPrevHighlighted ? m_ColorTheme.m_brButtonHighlighted : m_ColorTheme.m_brButton));

	pGM->FillRectangle(m_rectPrev, brPrev);
	pGM->DrawRectangle(m_rectPrev, m_ColorTheme.m_brBorder);

	CBCGPRect rectArrow = m_rectPrev;
	CBCGPPointsArray arPrevPoints;

	if (m_bIsHorizontal)
	{
		rectArrow.DeflateRect(5, 3);

		arPrevPoints.Add(CBCGPPoint(rectArrow.right, rectArrow.top));
		arPrevPoints.Add(CBCGPPoint(rectArrow.left, rectArrow.CenterPoint().y));
		arPrevPoints.Add(CBCGPPoint(rectArrow.right, rectArrow.bottom));
	}
	else
	{
		rectArrow.DeflateRect(3, 5);

		arPrevPoints.Add(CBCGPPoint(rectArrow.left, rectArrow.bottom));
		arPrevPoints.Add(CBCGPPoint(rectArrow.CenterPoint().x, rectArrow.top));
		arPrevPoints.Add(CBCGPPoint(rectArrow.right, rectArrow.bottom));
	}

	const CBCGPBrush& brArrowPrev = m_bIsPrevHighlighted && m_bPrevIsClicked && !bIs3D ? m_ColorTheme.m_brFace : m_ColorTheme.m_brBorder;

	pGM->DrawGeometry(CBCGPPolygonGeometry(arPrevPoints, FALSE), brArrowPrev, dblLineWidth);

	const CBCGPBrush& brNext = bIs3D ? 
		(m_bIsNextHighlighted && m_bNextIsClicked ? m_ColorTheme.m_brButtonPressed3D : (m_bIsNextHighlighted ? m_ColorTheme.m_brButtonHighlighted3D : m_ColorTheme.m_brButton3D)) : 
		(m_bIsNextHighlighted && m_bNextIsClicked ? m_ColorTheme.m_brButtonPressed : (m_bIsNextHighlighted ? m_ColorTheme.m_brButtonHighlighted : m_ColorTheme.m_brButton));

	pGM->FillRectangle(m_rectNext, brNext);
	pGM->DrawRectangle(m_rectNext, m_ColorTheme.m_brBorder);

	rectArrow = m_rectNext;
	CBCGPPointsArray arNextPoints;

	if (m_bIsHorizontal)
	{
		rectArrow.DeflateRect(5, 3);

		arNextPoints.Add(CBCGPPoint(rectArrow.left, rectArrow.top));
		arNextPoints.Add(CBCGPPoint(rectArrow.right, rectArrow.CenterPoint().y));
		arNextPoints.Add(CBCGPPoint(rectArrow.left, rectArrow.bottom));
	}
	else
	{
		rectArrow.DeflateRect(3, 5);

		arNextPoints.Add(CBCGPPoint(rectArrow.left, rectArrow.top));
		arNextPoints.Add(CBCGPPoint(rectArrow.CenterPoint().x, rectArrow.bottom));
		arNextPoints.Add(CBCGPPoint(rectArrow.right, rectArrow.top));
	}

	const CBCGPBrush& brArrowNext = m_bIsNextHighlighted && m_bNextIsClicked && !bIs3D ? m_ColorTheme.m_brFace : m_ColorTheme.m_brBorder;

	pGM->DrawGeometry(CBCGPPolygonGeometry(arNextPoints, FALSE), brArrowNext, dblLineWidth);
}
//*******************************************************************************
BOOL CBCGPVisualScrollBar::OnDragThumb(const CBCGPPoint& pt)
{
	if (!m_bIsDraggingThumb || m_dblTotal <= 0.)
	{
		return FALSE;
	}

	double size = m_bIsHorizontal ? m_rectScroll.Width() : m_rectScroll.Height();
	double dblScrollRatio = size / m_dblTotal;
	double dblOffsetOld = m_dblOffset;
	double delta = m_bIsHorizontal ? pt.x - m_ptDragThumbLast.x : pt.y - m_ptDragThumbLast.y;

	m_dblOffset = min(max(0, m_dblOffset + delta / dblScrollRatio), m_dblTotal - size); 

	m_ptDragThumbLast = pt;

	ReposThumb();
	
	if (m_pOwner != NULL)
	{
		m_pOwner->OnScroll(this, m_dblOffset - dblOffsetOld);
	}
	else if (m_pOwnerContainer != NULL)
	{
		m_pOwnerContainer->OnScroll(this, m_dblOffset - dblOffsetOld);
	}
	return TRUE;
}
//*******************************************************************************
void CBCGPVisualScrollBar::OnScrollStep(BOOL bNext)
{
	if (m_dblTotal <= 0.)
	{
		return;
	}

	double size = m_bIsHorizontal ? m_rectScroll.Width() : m_rectScroll.Height();
	double dblScrollRatio = size / m_dblTotal;
	double dblOffsetOld = m_dblOffset;
	double dblOffset = m_dblOffset;
	double delta = m_dblStep == 0. ? (m_rectPrev.Width() * dblScrollRatio) : m_dblStep;

	if (bNext)
	{
		dblOffset += delta;
	}
	else
	{
		dblOffset -= delta;
	}

	m_dblOffset = min(max(0, dblOffset), m_dblTotal - size); 

	ReposThumb();

	if (m_pOwner != NULL)
	{
		m_pOwner->OnScroll(this, m_dblOffset - dblOffsetOld);
	}
	else if (m_pOwnerContainer != NULL)
	{
		m_pOwnerContainer->OnScroll(this, m_dblOffset - dblOffsetOld);
	}
}
//*******************************************************************************
void CBCGPVisualScrollBar::OnChangeVisualManager()
{
	if (m_Style == BCGP_VISUAL_SCROLLBAR_VISUAL_MANAGER)
	{
		SetStyle(m_Style);
	}
}
//*******************************************************************************
void CBCGPVisualScrollBar::SetStyle(BCGP_VISUAL_SCROLLBAR_STYLE style)
{
	m_Style = style;

	if (style == BCGP_VISUAL_SCROLLBAR_VISUAL_MANAGER)
	{
		m_ColorTheme.m_brFace = CBCGPBrush(CBCGPDrawManager::IsPaleColor(globalData.clrBarFace) ?
			CBCGPDrawManager::ColorMakeDarker(globalData.clrBarFace) : CBCGPDrawManager::ColorMakeLighter(globalData.clrBarFace));

		m_ColorTheme.m_brBorder = CBCGPBrush(globalData.clrBarShadow);
		m_ColorTheme.m_brButton = CBCGPBrush(globalData.clrBarFace);
		m_ColorTheme.m_brButtonPressed = CBCGPBrush(globalData.clrBarShadow);
		m_ColorTheme.m_brButtonHighlighted = m_ColorTheme.m_brFace;
	}

	Redraw();
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPWndHostVisualObject

IMPLEMENT_DYNAMIC(CBCGPWndHostVisualObject, CBCGPBaseVisualObject)

CBCGPWndHostVisualObject::CBCGPWndHostVisualObject(CBCGPVisualContainer* pContainer, CRuntimeClass* pRTI)
	: CBCGPBaseVisualObject(pContainer)
	, m_pWnd(NULL)
	, m_pRTI(pRTI)
{
}
//*******************************************************************************
CBCGPWndHostVisualObject::~CBCGPWndHostVisualObject()
{
	if (m_pWnd != NULL)
	{
		if (m_pWnd->GetSafeHwnd() != NULL)
		{
			m_pWnd->DestroyWindow();
		}
		
		delete m_pWnd;
		m_pWnd = NULL;
	}
}
//*******************************************************************************
CWnd* CBCGPWndHostVisualObject::GetWnd()
{
	if (m_pWnd->GetSafeHwnd() == NULL && m_pWndOwner->GetSafeHwnd() != NULL)
	{
		m_pWnd = CreateWnd(m_rect, m_nID, m_pWndOwner);
		if (m_textFormat.IsEmpty())
		{
			LOGFONT lf;
			globalData.fontRegular.GetLogFont(&lf);

			SetTextFormat(CBCGPTextFormat(lf), FALSE);
		}

		if (m_pWnd->GetSafeHwnd() != NULL)
		{
			SetDirty(TRUE, TRUE);
		}
	}
	
	return m_pWnd;
}
//*******************************************************************************
void CBCGPWndHostVisualObject::ReposWnd()
{
	ASSERT_VALID(this);
	
	if (m_pParentContainer != NULL && m_pParentContainer->IsPrinting())
	{
		return;
	}

	CWnd* pWnd = GetWnd();
	if (pWnd->GetSafeHwnd() == NULL)
	{
		return;
	}

	ASSERT_VALID(pWnd);
		
	CRect rectWnd = m_rect;
	rectWnd.OffsetRect(-m_ptScrollOffset);

	pWnd->SetWindowPos(NULL, rectWnd.left, rectWnd.top, rectWnd.Width(), rectWnd.Height(),
		SWP_NOZORDER | SWP_NOACTIVATE);

	SetDirty(FALSE);
}
//*******************************************************************************
void CBCGPWndHostVisualObject::OnDraw(CBCGPGraphicsManager* pGM, const CBCGPRect& /*rectClip*/, DWORD dwFlags)
{
	if (m_pWnd->GetSafeHwnd() == NULL)
	{
		return;
	}

	if (m_Font.GetSafeHandle() == NULL)
	{
		LOGFONT lf;

		if (m_textFormat.GetFontFamily().IsEmpty())
		{
			globalData.fontRegular.GetLogFont(&lf);
		}
		else
		{
			m_textFormat.ExportToLogFont(lf);
		}

		m_Font.CreateFontIndirect(&lf);
		m_pWnd->SetFont(&m_Font);

		BOOL bDirty = IsDirty();

		ReposWnd();

		SetDirty(bDirty, FALSE);
	}

	if (IsEditMode())
	{
		if (m_pWnd->IsWindowVisible())
		{
			m_pWnd->ShowWindow(SW_HIDE);
		}
	}
	else
	{
		if (!m_pWnd->IsWindowVisible())
		{
			m_pWnd->ShowWindow(SW_SHOWNOACTIVATE);
		}

		return;
	}

	if ((dwFlags & BCGP_DRAW_STATIC) != BCGP_DRAW_STATIC)
	{
		return;
	}

	ASSERT_VALID (pGM);

	if (m_ImageCache.GetHandle() == NULL || IsDirty())
	{
		m_ImageCache.Destroy();

		CRect rectWnd;
		m_pWnd->GetWindowRect(rectWnd);
		rectWnd.OffsetRect(-rectWnd.left, -rectWnd.top);

		CDC dc;
		if (dc.CreateCompatibleDC(NULL))
		{
			LPBYTE pBitsWnd = NULL;
			HBITMAP hBitmap = CBCGPDrawManager::CreateBitmap_32 (rectWnd.Size(), (void**)&pBitsWnd);
			
			if (hBitmap != NULL)
			{
				HBITMAP hBitmapOld = (HBITMAP)dc.SelectObject(hBitmap);
				if (hBitmapOld != NULL)
				{
					dc.FillRect(rectWnd, &globalData.brWindow);

					m_pWnd->SendMessage(WM_PRINT, (WPARAM)dc.GetSafeHdc(), (LPARAM)(PRF_CLIENT | PRF_CHILDREN | PRF_NONCLIENT | PRF_ERASEBKGND));
					
					for (int i = 0; i < rectWnd.Width() * rectWnd.Height(); i++)
					{
						pBitsWnd[3] = 255;
						pBitsWnd += 4;
					}
					
					dc.SelectObject (hBitmapOld);
				}					

				m_ImageCache = CBCGPImage(hBitmap);
			}
		}
	}

	if (m_ImageCache.GetBitmap() != NULL)
	{
		pGM->DrawImage(m_ImageCache, m_rect.TopLeft());
	}
}
//*******************************************************************************
void CBCGPWndHostVisualObject::SetTextFormat(const CBCGPTextFormat& textFormat, BOOL bRedraw)
{
	m_textFormat = textFormat;
	
	if (m_Font.GetSafeHandle() != NULL)
	{
		m_Font.DeleteObject();
	}

	m_ImageCache.Destroy();

	SetDirty(TRUE, bRedraw);
}
//*******************************************************************************
void CBCGPWndHostVisualObject::SetColorTheme(const CBCGPWndHostColors& colors)
{
	m_Colors = colors;
	m_ImageCache.Destroy();

	SetDirty(TRUE, TRUE);
}
//*******************************************************************************
CWnd* CBCGPWndHostVisualObject::CreateWnd(const CRect& rect, UINT nID, CWnd* pParent)
{
	CWnd* pWnd = NULL;
	
	if (m_pParentContainer != NULL && m_pRTI != NULL && m_pRTI->IsDerivedFrom(RUNTIME_CLASS(CWnd)))
	{
		pWnd = DYNAMIC_DOWNCAST(CWnd, m_pRTI->CreateObject());
	}

	if (pWnd == NULL)
	{
		ASSERT(FALSE);
		return NULL;
	}
	
	ASSERT_VALID(pWnd);

	if (!m_pParentContainer->OnCreateCustomControl(this, pWnd, rect, nID, pParent))
	{
		delete pWnd;
		return NULL;
	}
	
	return pWnd;
}

#ifndef BCGP_EXCLUDE_GRID_CTRL

/////////////////////////////////////////////////////////////////////////////
// CBCGPGridVisualObject

IMPLEMENT_DYNCREATE(CBCGPGridVisualObject, CBCGPWndHostVisualObject)

CBCGPGridVisualObject::CBCGPGridVisualObject(CBCGPVisualContainer* pContainer, CRuntimeClass* pRTI)
	: CBCGPWndHostVisualObject(pContainer, pRTI)
{
	m_pColorTheme = new CBCGPGridColors;
}
//*******************************************************************************
CBCGPGridVisualObject::~CBCGPGridVisualObject()
{
	delete m_pColorTheme;
}
//*******************************************************************************
CBCGPGridCtrl* CBCGPGridVisualObject::GetGridCtrl()
{
	return DYNAMIC_DOWNCAST(CBCGPGridCtrl, GetWnd());
}
//*******************************************************************************
void CBCGPGridVisualObject::SetColorTheme(const CBCGPGridColors& colors)
{
	m_pColorTheme->FullCopy(colors);
	
	CBCGPGridCtrl* pWndGrid = DYNAMIC_DOWNCAST(CBCGPGridCtrl, m_pWnd);
	if (pWndGrid->GetSafeHwnd() != NULL)
	{
		pWndGrid->SetColorTheme(colors, TRUE, TRUE);
	}
}
//*******************************************************************************
CWnd* CBCGPGridVisualObject::CreateWnd(const CRect& rect, UINT nID, CWnd* pParent)
{
	CBCGPGridCtrl* pWndGrid = NULL;

	if (m_pRTI != NULL)
	{
		if (m_pRTI->IsDerivedFrom(RUNTIME_CLASS(CBCGPGridCtrl)))
		{
			pWndGrid = DYNAMIC_DOWNCAST(CBCGPGridCtrl, m_pRTI->CreateObject());
		}
		else
		{
			ASSERT(FALSE);
			return NULL;
		}
	}
	else
	{
		pWndGrid = new CBCGPGridCtrl;
	}

	ASSERT_VALID(pWndGrid);

	pWndGrid->Create(WS_CHILD | WS_VISIBLE | WS_BORDER, rect, pParent, nID);

	pWndGrid->SetScrollBarsStyle(CBCGPScrollBar::BCGP_SBSTYLE_VISUAL_MANAGER);
	pWndGrid->SetColorTheme(*m_pColorTheme, TRUE, TRUE);
	
	pWndGrid->SetScale(max(m_sizeScaleRatio.cx, m_sizeScaleRatio.cy));

	return pWndGrid;
}
//*******************************************************************************
void CBCGPGridVisualObject::OnScaleRatioChanged(const CBCGPSize& sizeScaleRatioOld)
{
	CBCGPWndHostVisualObject::OnScaleRatioChanged(sizeScaleRatioOld);

	CBCGPGridCtrl* pWndGrid = GetGridCtrl();
	if (pWndGrid->GetSafeHwnd() != NULL)
	{
		pWndGrid->SetScale(max(m_sizeScaleRatio.cx, m_sizeScaleRatio.cy));
	}
}

#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPEditVisualObject

IMPLEMENT_DYNCREATE(CBCGPEditVisualObject, CBCGPWndHostVisualObject)

CBCGPEditVisualObject::CBCGPEditVisualObject(CBCGPVisualContainer* pContainer, CRuntimeClass* pRTI)
	: CBCGPWndHostVisualObject(pContainer, pRTI)
{
}
//*******************************************************************************
CBCGPEditVisualObject::~CBCGPEditVisualObject()
{
}
//*******************************************************************************
CBCGPEdit* CBCGPEditVisualObject::GetEditCtrl()
{
	return DYNAMIC_DOWNCAST(CBCGPEdit, GetWnd());
}
//*******************************************************************************	
CWnd* CBCGPEditVisualObject::CreateWnd(const CRect& rect, UINT nID, CWnd* pParent)
{
	CBCGPEdit* pEdit = NULL;

	if (m_pRTI != NULL)
	{
		if (m_pRTI->IsDerivedFrom(RUNTIME_CLASS(CBCGPEdit)))
		{
			pEdit = DYNAMIC_DOWNCAST(CBCGPEdit, m_pRTI->CreateObject());
		}
		else
		{
			ASSERT(FALSE);
			return NULL;
		}
	}
	else
	{
		pEdit = new CBCGPEdit;
	}

	ASSERT_VALID(pEdit);

	pEdit->SetColorTheme(m_ColorTheme);
	pEdit->m_bVisualManagerStyle = TRUE;
	
	pEdit->Create(WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP, rect, pParent, nID);
	pEdit->ModifyStyleEx(0, WS_EX_CLIENTEDGE);
	pEdit->SetPrompt(m_strName);
	pEdit->SetWindowText(m_strInitialValue);

	return pEdit;
}
//*******************************************************************************	
void CBCGPEditVisualObject::ReposWnd()
{
	CBCGPEdit* pEdit = GetEditCtrl();
	if (pEdit->GetSafeHwnd() == NULL)
	{
		CBCGPWndHostVisualObject::ReposWnd();
		return;
	}

	if (m_pParentContainer != NULL && m_pParentContainer->IsPrinting())
	{
		return;
	}

	CBCGPRect rectSaved = m_rect;

	if ((pEdit->GetStyle() & ES_MULTILINE) == 0)
	{
		CClientDC dc(pEdit);
		CBCGPFontSelector fs(dc, pEdit->GetFont());
		
		TEXTMETRIC tm;
		dc.GetTextMetrics(&tm);
		
		int nHeight = min((int)m_rect.Height(), tm.tmHeight + 8);

		m_rect.bottom = m_rect.top + nHeight;
	}

	CBCGPWndHostVisualObject::ReposWnd();

	m_rect = rectSaved;
}
//*******************************************************************************	
void CBCGPEditVisualObject::SetInitialValue(const CString& strVal)
{
	m_strInitialValue = strVal;

	if (m_pWnd->GetSafeHwnd() != NULL)
	{
		m_pWnd->SetWindowText(m_strInitialValue);
	}
}
//*******************************************************************************	
void CBCGPEditVisualObject::SetColorTheme(const CBCGPEditColors& colors)
{
	m_ColorTheme = colors;

	CBCGPEdit* pEdit = DYNAMIC_DOWNCAST(CBCGPEdit, m_pWnd);
	if (pEdit->GetSafeHwnd() != NULL)
	{
		pEdit->SetColorTheme(colors);
	}
}
