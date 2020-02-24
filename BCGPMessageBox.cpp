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
// BCGPMessageBox.cpp : implementation file
//

#include "stdafx.h"
#include "bcgprores.h"
#include "BCGPMessageBox.h"
#include "BCGPVisualManager.h"
#include "BCGPDrawManager.h"
#include "BCGPStatic.h"
#include "BCGPGlobalUtils.h"

#ifndef _BCGSUITE_
	#include "BCGPPopupMenu.h"
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

BOOL CBCGPMessageBox::m_bDontUseDefaultIcon = FALSE;
BOOL CBCGPMessageBox::m_bUseNativeControls = FALSE;
BOOL CBCGPMessageBox::m_bUseNativeCaption = FALSE;
BOOL CBCGPMessageBox::m_bShowImmediately = FALSE;
BOOL CBCGPMessageBox::m_bUseNativeIcons = FALSE;

static CSize GetWindowTextExtent (CWnd* pControl, const CString& strText, CFont* pFont = NULL)
{
	if (pFont->GetSafeHandle() == NULL && pControl != NULL)
	{
		ASSERT (pControl != NULL);
		pFont = pControl->GetFont ();
	}

	CWindowDC dc (pControl);

	CFont* pOldFont = NULL;
	if (pFont->GetSafeHandle() != NULL)
	{
		pOldFont = dc.SelectObject (pFont);
	}

	CSize szText = dc.GetTextExtent (strText);

	if (pOldFont->GetSafeHandle() != NULL)
	{
		dc.SelectObject (pOldFont);
	}

	return szText;
}

static CSize GetDlgBaseUnits(CDialog* pDialog, CFont* pFont = NULL)
{
	if (pFont->GetSafeHandle() == NULL && pDialog->GetSafeHwnd() == NULL)
	{
		const long nBaseUnits = GetDialogBaseUnits();
		return CSize(LOWORD(nBaseUnits), HIWORD(nBaseUnits));
	}

	if (pFont->GetSafeHandle() == NULL)
	{
		ASSERT (pDialog != NULL);
		pFont = pDialog->GetFont ();
	}

	CWindowDC dc(pDialog);

	CFont* pOldFont = NULL;
	if (pFont->GetSafeHandle() != NULL)
	{
		pOldFont = dc.SelectObject (pFont);
	}

	TEXTMETRIC tm;
	GetTextMetrics(dc.GetSafeHdc(), &tm);

	CSize size;
	GetTextExtentPoint32(dc.GetSafeHdc(), _T("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"), 52, &size);

	if (pFont->GetSafeHandle() != NULL)
	{
		dc.SelectObject (pOldFont);
	}

	return CSize((size.cx / 26 + 1) / 2, tm.tmHeight);
}

// Adjust controls horizontally and vertically.
// If rcBounds.top and rcBounds.bottom are equal, rcBounds.top is used as Y-coord
// for all specified controls.
// Use TA_... constants for dwAdjustment
static void AdjustControls (CWnd** pControlsArray, UINT nControlsCount, const RECT& rcBounds, DWORD dwAdjustment, BOOL bRTL)
{
	if (nControlsCount == 0 || pControlsArray == NULL)
	{
		return;
	}

	RECT rect;
	pControlsArray[bRTL ? nControlsCount - 1 : 0]->GetWindowRect(&rect);
	int nLeft = rect.left;
	pControlsArray[bRTL ? 0 : nControlsCount - 1]->GetWindowRect(&rect);
	int nRight = rect.right;

	int xOffset = 0;

	const DWORD TA_HorzMask = TA_LEFT | TA_CENTER | TA_RIGHT;

	switch (dwAdjustment & TA_HorzMask)
	{
	case TA_LEFT:
		xOffset = rcBounds.left - nLeft;
		break;
	case TA_CENTER:
		xOffset = ((rcBounds.right + rcBounds.left) - (nRight + nLeft)) / 2;
		break;
	case TA_RIGHT:
		xOffset = rcBounds.right - nRight;
		break;
	}

	for (UINT i = 0; i < nControlsCount; ++i)
	{
		pControlsArray[i]->GetWindowRect (&rect);

		int y = (rcBounds.top == rcBounds.bottom) ? rcBounds.top : rect.top;
		pControlsArray[i]->MoveWindow (rect.left + xOffset, y, rect.right - rect.left, rect.bottom - rect.top, TRUE);
	}
}

static int LoadLocalizedString (HINSTANCE hInst, UINT uID, LPWSTR lpBuffer, int nBufferMax, WORD wLang)
{
	ASSERT (nBufferMax > 0);

	HRSRC hrsrc = ::FindResourceEx (hInst, RT_STRING, (LPCTSTR)(LONG_PTR)(1 + (USHORT)uID / 16), wLang);

	if (hrsrc == NULL)
	{
		hrsrc = ::FindResourceEx (hInst, RT_STRING, (LPCTSTR)(LONG_PTR)(1 + (USHORT)uID / 16), MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL));
	}

	int nCharCount = 0;

	if (hrsrc != NULL)
	{
		PWCHAR pwch = (PWCHAR)::LoadResource(hInst, hrsrc);
		if (pwch)
		{
			for (uID %= 16; uID; uID--)
			{
				pwch += *pwch + 1;
			}

			nCharCount = min (*pwch, nBufferMax - 1);
			memcpy (lpBuffer, pwch+1, nCharCount * sizeof(WCHAR));
		}
	}

	lpBuffer [nCharCount] = 0;
	return nCharCount;
}

#ifdef BCGP_USE_STANDARD_ICONS

class CBCGPStandardIcons
{
public:
    struct XIcons
    {
        HICON m_hSmall;
        HICON m_hLarge;

        XIcons()
            : m_hSmall(NULL)
            , m_hLarge(NULL)
        {
        }
        XIcons(HICON hSmall, HICON hLarge)
            : m_hSmall(hSmall)
            , m_hLarge(hLarge)
        {
        }
        ~XIcons()
        {
            if (m_hSmall != NULL)
            {
                ::DestroyIcon(m_hSmall);
            }

            if (m_hLarge != NULL)
            {
                ::DestroyIcon(m_hLarge);
            }
        }
    };

public:
    CBCGPStandardIcons()
    {
        _Initialize();
    }
    ~CBCGPStandardIcons()
    {
        POSITION pos = m_Map.GetStartPosition();
        while (pos != NULL)
        {
            UINT key = 0;
            XIcons* value = NULL;
            m_Map.GetNextAssoc(pos, key, value);

            if (value != NULL)
            {
                delete value;
            }
        }
    }

    static HICON GetIcon(LPCTSTR icon, BOOL bLarge = FALSE)
    {
        return s_Icons._GetIcon(icon, bLarge);
    }

protected:
    HICON _GetIcon(LPCTSTR icon, BOOL bLarge = FALSE) const
    {
        XIcons* value = NULL;
        if (!m_Map.Lookup((UINT)(WORD)icon, value))
        {
            return NULL;
        }

        if (value == NULL)
        {
            return NULL;
        }

        return bLarge ? value->m_hLarge : value->m_hSmall;
    }

    int _GetIconID(LPCTSTR icon) const
    {
        int nIcon = 0;

        switch ((WORD)icon)
        {
        case (WORD)IDI_APPLICATION:
            nIcon = -100;
            break;
        case (WORD)IDI_WARNING:
            nIcon = -101;
            break;
        case (WORD)IDI_QUESTION:
            nIcon = -102;
            break;
        case (WORD)IDI_ERROR:
            nIcon = -103;
            break;
        case (WORD)IDI_INFORMATION:
            nIcon = -104;
            break;
        case (WORD)IDI_WINLOGO:
            nIcon = -105; // default application icon (Windows XP)
            break;
#ifdef IDI_SHIELD
        case (WORD)IDI_SHIELD:
#else
		case (WORD)MAKEINTRESOURCE(32518):
#endif
            nIcon = -106;
            break;
        }

        return nIcon;
    }

    void _Initialize()
    {
        LPCTSTR szID[] =
        {
            IDI_APPLICATION, 
            IDI_ERROR, 
            IDI_QUESTION, 
            IDI_WARNING, 
            IDI_INFORMATION, 
            IDI_WINLOGO, 
#ifdef IDI_SHIELD
			IDI_SHIELD
#else
			MAKEINTRESOURCE(32518)
#endif
        };

        for (int i = 0; i < sizeof(szID) / sizeof(szID[0]); i++)
        {
            HICON hSmall = NULL;
            HICON hLarge = NULL;

            int nIcon = _GetIconID(szID[i]);
            if (nIcon == 0)
            {
                continue;
            }

            if (::ExtractIconEx(_T("user32.dll"), nIcon, &hLarge, &hSmall, 1) > 0)
            {
                m_Map[(UINT)(WORD)szID[i]] = new XIcons(hSmall, hLarge);
            }
        }
    }

protected:
    CMap<UINT, UINT, XIcons*, XIcons*>    m_Map;

    static CBCGPStandardIcons s_Icons;
};

CBCGPStandardIcons CBCGPStandardIcons::s_Icons;

#endif

class CBCGPButtonMB: public CBCGPButton
{
	friend class CBCGPMessageBox;

	DECLARE_DYNCREATE(CBCGPButtonMB)

public:
	CBCGPButtonMB();

protected:
	afx_msg void OnSetFocus(CWnd*);
	
	DECLARE_MESSAGE_MAP()
};

IMPLEMENT_DYNCREATE(CBCGPButtonMB, CBCGPButton)

CBCGPButtonMB::CBCGPButtonMB()
{
}

BEGIN_MESSAGE_MAP(CBCGPButtonMB, CBCGPButton)
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CBCGPButtonMB::OnSetFocus(CWnd* pOldWnd)
{
	CBCGPButton::OnSetFocus(pOldWnd);

	CWnd* pParent = GetParent();
	if (pParent->GetSafeHwnd() != NULL)
	{
		pParent->SendMessage(DM_SETDEFID, (WPARAM)GetDlgCtrlID(), (LPARAM)0);
	}
}


CBCGPMessageBox::CBCGPMessageBox (const MSGBOXPARAMS* pParams)
: CBCGPDialog(CBCGPMessageBox::IDD)
{
	if (pParams != NULL)
	{
		CopyMemory (&m_Params, pParams, sizeof(MSGBOXPARAMS));
	}
	Initialize ();
}

CBCGPMessageBox::CBCGPMessageBox (const BCGP_MSGBOXPARAMS* pParams)
	: CBCGPDialog(CBCGPMessageBox::IDD)
{
	if (pParams != NULL)
	{
		CopyMemory (&m_Params, pParams, sizeof(BCGP_MSGBOXPARAMS));
	}
	Initialize ();
}


CString CBCGPMessageBox::GetString (LPCTSTR lpszText, LPCTSTR lpszDefault /*= NULL*/) const
{
	CString strRet (lpszDefault);

	if (lpszText != NULL)
	{
		if (HIWORD (lpszText) == NULL)
		{
			PWCHAR pwszBuffer = new WCHAR[4096];
			LoadLocalizedString (m_Params.hInstance, LOWORD (lpszText), pwszBuffer, 4095, (WORD)m_Params.dwLanguageId);
			strRet = pwszBuffer;
			delete [] pwszBuffer;
		}
		else
		{
			strRet = lpszText;
		}
	}

	return strRet;
}

void CBCGPMessageBox::Initialize ()
{
	m_hwndOwner = m_Params.hwndOwner;
	m_bCancelModeCapturedWindow = TRUE;

	CWnd* pParent = CWnd::FromHandle (m_Params.hwndOwner);

	if (pParent == NULL)
	{
		pParent = CWnd::GetActiveWindow();

		if (pParent != NULL)
		{
			CWnd* pPopup = pParent->GetLastActivePopup();

			if (pPopup != NULL)
			{
				pParent = pPopup;
			}
		}

		m_hwndOwner = pParent->GetSafeHwnd();
	}

	// Loading an icon
	DWORD dwStyle = m_Params.dwStyle;
	if (globalData.m_bIsRTL)
	{
		dwStyle |= MB_RTLREADING;
	}

	// determine icon based on type specified
	if ((dwStyle & MB_ICONMASK) == 0 && !m_bDontUseDefaultIcon)
	{
		switch (dwStyle & MB_TYPEMASK)
		{
		case MB_OK:
		case MB_OKCANCEL:
			dwStyle |= MB_ICONWARNING;
			break;

		case MB_YESNO:
		case MB_YESNOCANCEL:
			dwStyle |= MB_ICONQUESTION;
			break;

		case MB_ABORTRETRYIGNORE:
		case MB_RETRYCANCEL:
			// No default icon for these types, since they are rarely used.
			// The caller should specify the icon.
			break;
		}
	}

	HINSTANCE iconInstance = NULL;
	LPCTSTR iconID = NULL;
	UINT uiLoad = LR_CREATEDIBSECTION | LR_DEFAULTSIZE;
	HICON hMessageIcon = NULL;
	BOOL bDestroyIcon = FALSE;

	if (dwStyle & MB_USERICON)
	{
		iconInstance = m_Params.hInstance;
		iconID = m_Params.lpszIcon;
	}
	else
	{
		switch (dwStyle & MB_ICONMASK)
		{
		case MB_ICONINFORMATION:
			iconID = IDI_INFORMATION;
			break;

		case MB_ICONWARNING:
			iconID = IDI_WARNING;
			break;

		case MB_ICONERROR:
			iconID = IDI_ERROR;
			break;

		case MB_ICONQUESTION:
			iconID = IDI_QUESTION;
			break;
		}

		uiLoad |= LR_SHARED;

#ifdef BCGP_USE_STANDARD_ICONS
		hMessageIcon = CBCGPStandardIcons::GetIcon(iconID, TRUE);
#endif
	}

	if (hMessageIcon == NULL)
	{
#ifndef _BCGSUITE_
		if (!m_bUseNativeIcons)
		{
			hMessageIcon = globalData.GetCommonIcon(iconID);
		}
#endif
		if (hMessageIcon == NULL)
		{
			hMessageIcon = (HICON)::LoadImage(iconInstance, iconID, IMAGE_ICON, 0, 0, uiLoad);
			bDestroyIcon = TRUE;
		}
	}

	m_rectIcon.SetRectEmpty ();

	if (hMessageIcon != NULL)
	{
		ICONINFO info;
		::GetIconInfo(hMessageIcon, &info);
		
		BITMAP bmp;
		::GetObject (info.hbmColor, sizeof (BITMAP), (LPVOID) &bmp);
		
		CBCGPToolBarImages image;
		m_Icon.SetImageSize (CSize (bmp.bmWidth, bmp.bmHeight));
		
		m_Icon.AddIcon (hMessageIcon, bmp.bmBitsPixel >= 32);

		const int nCXIcon = ::GetSystemMetrics(SM_CXICON);
		const int nCYIcon = ::GetSystemMetrics(SM_CYICON);

#ifndef _BCGSUITE_
		if (nCXIcon != bmp.bmWidth && bmp.bmWidth != 0)
		{
			double dblRatio = ((double)nCXIcon) / bmp.bmWidth;
			m_Icon.SmoothResize(dblRatio);
		}
#endif

		::DeleteObject (info.hbmColor);
		::DeleteObject (info.hbmMask);

        m_rectIcon.right = nCXIcon;
		m_rectIcon.bottom = nCYIcon;

		if (bDestroyIcon)
		{
			::DestroyIcon(hMessageIcon);
		}
	}

	m_strMessageCaption = GetString (m_Params.lpszCaption, AfxGetAppName());
	m_strMessageHeader = GetString(m_Params.lpszHeader);
	m_strMessageText = GetString(m_Params.lpszText);
	m_strCheckBox = GetString (m_Params.lpszCheckBoxText);

	switch (dwStyle & MB_DEFMASK)
	{
	case MB_DEFBUTTON1:
		m_nDefaultButtonIndex = 0;
		break;
	case MB_DEFBUTTON2:
		m_nDefaultButtonIndex = 1;
		break;
	case MB_DEFBUTTON3:
		m_nDefaultButtonIndex = 2;
		break;
	case MB_DEFBUTTON4:
		m_nDefaultButtonIndex = 3;
		break;
	default:
		m_nDefaultButtonIndex = -1;
		break;
	}

	m_bRightAlignment = (m_Params.dwStyle & MB_RIGHT) != 0;

	// Calculating margins, sizes and spacings
	const int cBaseUnit = 7;
	m_rectClientMargins.SetRect (cBaseUnit, cBaseUnit, cBaseUnit, cBaseUnit);
	m_rectButtonsMargins.SetRect (cBaseUnit, cBaseUnit, 4, 4);
	m_szButton.cx = 50;
	m_szButton.cy = 14;
	m_cxIconSpacing = cBaseUnit; // spacing between an icon and message text
	m_cyCheckBoxHeight = 10;
	m_cyCheckBoxSpacing = 9;
	m_cxButtonSpacing = globalUtils.ScaleByDPI(4); // gap between buttons
	m_bIsLocal = TRUE;

	EnableVisualManagerStyle (!m_Params.bUseNativeControls, !m_Params.bUseNativeCaption);
	EnableDragClientArea();

	m_nAccTextIndex = -1;
	m_nAccHeaderIndex = -1;

#if _MSC_VER >= 1300
	EnableActiveAccessibility();
#endif
}

BEGIN_MESSAGE_MAP(CBCGPMessageBox, CBCGPDialog)
	//{{AFX_MSG_MAP(CBCGPMessageBox)
	ON_WM_PAINT()
	ON_WM_ERASEBKGND()
	ON_MESSAGE(WM_HELP, OnHelp)
	ON_MESSAGE(DM_GETDEFID, OnGetDefID)
	ON_MESSAGE(DM_SETDEFID, OnSetDefID)
	ON_COMMAND_RANGE(1, 0xFFFE, OnButton)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPMessageBox message handlers

BOOL CBCGPMessageBox::OnEraseBkgnd(CDC* pDC)
{
	ASSERT_VALID(pDC);

	if (!IsVisualManagerStyle())
	{
		CRect rectClient;
		GetClientRect (rectClient);

		pDC->FillRect(rectClient, &globalData.brWindow);

		if (!m_rectButtons.IsRectEmpty())
		{
			pDC->FillRect(m_rectButtons, &globalData.brBtnFace);
		}

		return TRUE;
	}

	if (!CBCGPDialog::OnEraseBkgnd(pDC))
	{
		return FALSE;
	}

	if (globalData.m_nBitsPerPixel > 8 && !globalData.IsHighContastMode () && !m_rectButtons.IsRectEmpty())
	{
		CBCGPVisualManager::GetInstance()->OnDrawButtonsArea(pDC, this, m_rectButtons);
	}

	return TRUE;
}

void CBCGPMessageBox::OnPaint ()
{
	CPaintDC dc(this); // device context for painting

	if (IsVisualManagerStyle() && m_Impl.IsOwnerDrawCaption())
	{
		CRect rect(m_rectText);
		rect.UnionRect(rect, m_rectIcon);
		rect.UnionRect(rect, m_rectHeader);

		CRgn rgn;
		rgn.CreateRectRgnIndirect(rect);

		dc.SelectClipRgn(&rgn);

		SendMessage(WM_ERASEBKGND, (WPARAM)dc.GetSafeHdc(), 0);

		dc.SelectClipRgn(NULL);
	}

	if (m_Icon.IsValid())
	{
		m_Icon.DrawEx(&dc, m_rectIcon, 0, CBCGPToolBarImages::ImageAlignHorzStretch, CBCGPToolBarImages::ImageAlignVertStretch);
	}

	DrawHeader(&dc);
	DrawText(&dc);

	if (m_Params.bShowCheckBox && m_Params.bShowSeparator)
	{
		CRect rectClient;
		GetClientRect (&rectClient);
		rectClient.DeflateRect (m_rectClientMargins);
		rectClient.top  = m_cySeparatorLine - 1;
		rectClient.bottom = m_cySeparatorLine + 1;

		CBCGPStatic dummy;

#ifndef _BCGSUITE_
		if (IsVisualManagerStyle() && CBCGPVisualManager::GetInstance ()->IsOwnerDrawDlgSeparator(&dummy))
		{
			CBCGPVisualManager::GetInstance ()->OnDrawDlgSeparator(&dc, &dummy, rectClient, TRUE);
		}
		else
#endif
		{
			dc.Draw3dRect (rectClient, globalData.clrBarDkShadow, globalData.clrBarHilite);
		}
	}
}

BOOL CBCGPMessageBox::OnInitDialog()
{
	if (m_Params.m_bNoIdleMsg || CBCGPPopupMenu::GetActiveMenu() != NULL)
	{
		ModifyStyle(0, DS_NOIDLEMSG);
	}

	DoInitDialog();

	if (!CBCGPDialog::OnInitDialog ())
	{
		return FALSE;
	}

	CWnd* pWndOwner = NULL;

	if (m_Params.boxPosition == BCGP_MSGBOXPARAMS::BCGP_MSGBOX_POSITION_CENTER_OWNER)
	{
		pWndOwner = CWnd::FromHandle(m_hwndOwner);
	}

	if (m_Params.boxPosition != BCGP_MSGBOXPARAMS::BCGP_MSGBOX_POSITION_NO_CENTER)
	{
		CenterWindow (pWndOwner);
	}

	::MessageBeep (m_Params.dwStyle & MB_ICONMASK);

	BOOL bResult = TRUE;

	if (m_nDefaultButtonIndex >= 0 && m_nDefaultButtonIndex < (int)m_arrButtons.GetSize ())
	{
		m_arrButtons[m_nDefaultButtonIndex]->SetFocus ();
		bResult = FALSE; // Focus changed
	}

	if (CBCGPMessageBox::m_bShowImmediately)
	{
		ShowWindow(SW_SHOWNORMAL);
		UpdateWindow();
	}

	return bResult;
}

CBCGPMessageBox::~CBCGPMessageBox()
{
	for (int i = 0; i < (int)m_arrButtons.GetSize (); ++i)
	{
		delete m_arrButtons[i];
	}
}

void CBCGPMessageBox::DrawHeader(CDC* pDC)
{
	if (m_strMessageHeader.IsEmpty() || m_rectHeader.IsRectEmpty())
	{
		return;
	}

	CFont* pOldFont = pDC->SelectObject (&m_fntHeader);
	
	pDC->SetTextColor (IsVisualManagerStyle () ? 
		CBCGPVisualManager::GetInstance()->GetDlgHeaderTextColor(this) : 
		CBCGPVisualManager::GetInstance()->CBCGPVisualManager::GetDlgHeaderTextColor(this));

	int nBkMode = pDC->SetBkMode (TRANSPARENT);
	
	CRect rectTemp = m_rectHeader;

	UINT nFormat = DT_NOPREFIX | DT_NOCLIP | DT_END_ELLIPSIS | (m_Params.bHeaderWordBreak ? DT_WORDBREAK : 0);
	nFormat |= (m_bRightAlignment ? DT_RIGHT : DT_LEFT);

	pDC->DrawText (m_strMessageHeader, &rectTemp, nFormat);
	pDC->SelectObject (pOldFont);

	pDC->SetBkMode (nBkMode);
}

void CBCGPMessageBox::DrawText (CDC* pDC)
{
	CFont* pOldFont = pDC->SelectObject (&m_fntText);
	pDC->SetTextColor (IsVisualManagerStyle () ? CBCGPVisualManager::GetInstance()->GetDlgTextColor(this) : globalData.clrWindowText);

	int nBkMode = pDC->SetBkMode (TRANSPARENT);

	CRect rectTemp = m_rectText;

	pDC->DrawText (m_strMessageText, &rectTemp, m_dwDrawTextFlags | DT_NOCLIP);
	pDC->SelectObject (pOldFont);

	pDC->SetBkMode (nBkMode);
}

BOOL CBCGPMessageBox::ContinueModal()
{
	if (!CBCGPDialog::ContinueModal())
	{
		return FALSE;
	}

	return m_hWnd != NULL && ::IsWindow(m_hWnd);
}

#ifndef IDTRYAGAIN
#define IDTRYAGAIN  10
#endif

#ifndef IDCONTINUE
#define IDCONTINUE  11
#endif

#ifndef MB_CANCELTRYCONTINUE
#define MB_CANCELTRYCONTINUE    6
#endif

struct STD_BUTTON_TEXT 
{
	UINT    id;
	UINT    resId;
	LPCTSTR pszText;
};

static const STD_BUTTON_TEXT s_Buttons [] =
{
	{ IDOK,       800,  _T("OK")         },
	{ IDCANCEL,   801,  _T("Cancel")     },
	{ IDABORT,    802,  _T("&Abort")     },
	{ IDRETRY,    803,  _T("&Retry")     },
	{ IDIGNORE,   804,  _T("&Ignore")    },
	{ IDYES,      805,  _T("&Yes")       },
	{ IDNO,       806,  _T("&No")        },
	{ IDHELP,	  808,  _T("&Help")		 },
	{ IDTRYAGAIN, 809,  _T("&Try Again") },
	{ IDCONTINUE, 810,  _T("&Continue")  }
};

void CBCGPMessageBox::AddButton (UINT id, HINSTANCE hInst, LPCTSTR lpszCaption /*= NULL*/)
{
	CString strText = lpszCaption;

	if (strText.IsEmpty ())
	{
		const STD_BUTTON_TEXT* sbt = NULL;

		for (int i = 0; i < sizeof(s_Buttons) / sizeof(*s_Buttons); ++i)
		{
			if (s_Buttons[i].id == id)
			{
				sbt = &s_Buttons[i];
				break;
			}
		}

		if (sbt != NULL)
		{
			if (sbt->resId != 0)
			{
				if (hInst != NULL)
				{
					PWCHAR pwszBuffer = new WCHAR [256];
					LoadLocalizedString (hInst, sbt->resId, pwszBuffer, 255, (WORD)m_Params.dwLanguageId);
					strText = pwszBuffer;
					delete [] pwszBuffer;
				}
			}

			if (strText.IsEmpty ())
			{
				strText = sbt->pszText;
			}
		}
	}

	DWORD dwStyle = WS_VISIBLE | WS_CHILD | WS_TABSTOP;

	int nButtonIndex = (int)m_arrButtons.GetSize ();

	if (nButtonIndex == m_nDefaultButtonIndex)
	{
		dwStyle |= BS_DEFPUSHBUTTON;
	}
	else
	{
		dwStyle |= BS_PUSHBUTTON;
	}

	int nButtonWidth = m_szButton.cx;

	CRect rectButton (m_ptNextButtonPos.x, 0, m_ptNextButtonPos.x + nButtonWidth, m_szButton.cy);

	CBCGPButtonMB* pBtn = new CBCGPButtonMB ();
	pBtn->Create (strText, dwStyle, rectButton, this, id);

	if (pBtn->GetSafeHwnd () != NULL)
	{
		pBtn->SetFont(&m_fntButtons, FALSE);

		CSize szButtonText = GetWindowTextExtent (pBtn, strText, &m_fntButtons);
		const int cxButtonMargins = 5; // Minimal space between button text and button's left and right borders.

		if (szButtonText.cx > m_szButton.cx - cxButtonMargins * 2)
		{
			// Resize button to show whole text
			szButtonText.cx = min (szButtonText.cx, 512); // Preserve from too long button text
			nButtonWidth = szButtonText.cx + cxButtonMargins * 2;

			pBtn->SetWindowPos (NULL, 0, 0, nButtonWidth, m_szButton.cy, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS);
		}

		m_ptNextButtonPos.x += nButtonWidth + m_cxButtonSpacing;

		m_arrButtons.Add (pBtn);
		if (id == IDCANCEL)
		{
			m_bHasCancelButton = true; // enable close button
		}
	}
}

void CBCGPMessageBox::DoInitDialog()
{
	DWORD dwStyle = 0;
	DWORD dwStyleEx = 0;

	if (m_Params.dwStyle & MB_SYSTEMMODAL)
	{
		dwStyle |= DS_SYSMODAL;
	}
	if (m_Params.dwStyle & MB_SETFOREGROUND)
	{
		dwStyle |= DS_SETFOREGROUND;
	}

	if (m_Params.dwStyle & MB_RTLREADING)
	{
		dwStyleEx |= WS_EX_LAYOUTRTL;
	}
	if (m_Params.dwStyle & MB_TOPMOST)
	{
		dwStyleEx |= WS_EX_TOPMOST;
	}

	if (dwStyle != 0)
	{
		ModifyStyle(0, dwStyle);
	}
	if (dwStyleEx != 0)
	{
		ModifyStyleEx(0, dwStyleEx);
	}

	if (m_Params.dwStyle & (MB_TOPMOST | MB_SETFOREGROUND))
	{
		SetWindowPos(&wndTopMost, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOOWNERZORDER);
	}

	if (m_Params.dwStyle & MB_SETFOREGROUND)
	{
		if ((m_Params.dwStyle & MB_TOPMOST) == 0)
		{
			SetWindowPos(&wndNoTopMost, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOOWNERZORDER);
		}

		SetForegroundWindow();
	}

	SetWindowText (m_strMessageCaption);

	NONCLIENTMETRICS ncm;
	ZeroMemory (&ncm, sizeof (ncm));
	ncm.cbSize = sizeof (ncm);

	::SystemParametersInfo (SPI_GETNONCLIENTMETRICS, 0, (LPVOID)&ncm, 0);

	CFont fntDefault;
	fntDefault.CreateFontIndirect (&ncm.lfMessageFont);

	// Setting font
	if (m_Params.hfontText == NULL)
	{
		m_fntText.CreateFontIndirect (&ncm.lfMessageFont);
	}
	else
	{
		CFont* pFont = CFont::FromHandle(m_Params.hfontText);
		ASSERT_VALID(pFont);

		LOGFONT lf = {0};
		pFont->GetLogFont(&lf);

		m_fntText.CreateFontIndirect(&lf);
	}

	LOGFONT lfHeader = {0};
	m_fntText.GetLogFont(&lfHeader);
	lfHeader.lfHeight = 2 * lfHeader.lfHeight;

	m_fntHeader.CreateFontIndirect(&lfHeader);

	if (m_fntButtons.GetSafeHandle() != NULL)
	{
		m_fntButtons.DeleteObject();
	}

	if (m_Params.hfontButtons == NULL)
	{
		m_fntButtons.CreateFontIndirect (&ncm.lfMessageFont);
	}
	else
	{
		CFont* pFont = CFont::FromHandle(m_Params.hfontButtons);
		ASSERT_VALID(pFont);

		LOGFONT lf = {0};
		pFont->GetLogFont(&lf);

		m_fntButtons.CreateFontIndirect(&lf);
	}

	const CSize szBaseUnits(GetDlgBaseUnits(this, &fntDefault));

	CSize szTextUnits(szBaseUnits);
	if (m_fntText.GetSafeHandle() != NULL)
	{
		szTextUnits = GetDlgBaseUnits(this, &m_fntText);
	}

	CSize szButtonUnits(szBaseUnits);
	if (m_fntButtons.GetSafeHandle() != NULL)
	{
		szButtonUnits = GetDlgBaseUnits(this, &m_fntButtons);
	}

	// Calculating gaps, spaces and margins
	{
		m_rectClientMargins.left = ::MulDiv(m_rectClientMargins.left, szBaseUnits.cx, 4);
		m_rectClientMargins.top = ::MulDiv(m_rectClientMargins.top, szBaseUnits.cy, 4);
		m_rectClientMargins.right = ::MulDiv(m_rectClientMargins.right, szBaseUnits.cx, 4);
		m_rectClientMargins.bottom = ::MulDiv(m_rectClientMargins.bottom, szBaseUnits.cy, 4);

		m_rectButtonsMargins.left = ::MulDiv(m_rectButtonsMargins.left, szBaseUnits.cx, 4);
		m_rectButtonsMargins.top = ::MulDiv(m_rectButtonsMargins.top, szBaseUnits.cy, 8);
		m_rectButtonsMargins.right = ::MulDiv(m_rectButtonsMargins.right, szBaseUnits.cx, 4);
		m_rectButtonsMargins.bottom = ::MulDiv(m_rectButtonsMargins.bottom, szBaseUnits.cy, 8);

		m_cxIconSpacing = ::MulDiv(m_cxIconSpacing, szTextUnits.cx, 4);
		m_cyCheckBoxHeight = ::MulDiv(m_cyCheckBoxHeight, szTextUnits.cy, 8);
		m_cyCheckBoxSpacing = ::MulDiv(m_cyCheckBoxSpacing, szBaseUnits.cy, 8);

		m_cxButtonSpacing = ::MulDiv(m_cxButtonSpacing, szButtonUnits.cx, 4);
		m_szButton.cx = ::MulDiv(m_szButton.cx, szButtonUnits.cx, 4);
		m_szButton.cy = ::MulDiv(m_szButton.cy, szButtonUnits.cy, 8);
	}

	// Creating buttons

	m_bHasCancelButton = false;
	m_ptNextButtonPos.x = 0;
	m_ptNextButtonPos.y = 0;

	HINSTANCE hInst = ::LoadLibrary (_T("user32.dll"));

	if (!m_Params.bIgnoreStandardButtons)
	{
		switch (m_Params.dwStyle & MB_TYPEMASK)
		{
		case MB_OK:
			AddButton (IDOK, hInst);
			break;
		case MB_OKCANCEL:
			AddButton (IDOK, hInst);
			AddButton (IDCANCEL, hInst);
			break;
		case MB_ABORTRETRYIGNORE:
			AddButton (IDABORT, hInst);
			AddButton (IDRETRY, hInst);
			AddButton (IDIGNORE, hInst);
			break;
		case MB_YESNO:
			AddButton (IDYES, hInst);
			AddButton (IDNO, hInst);
			break;
		case MB_YESNOCANCEL:
			AddButton (IDYES, hInst);
			AddButton (IDNO, hInst);
			AddButton (IDCANCEL, hInst);
			break;
		case MB_RETRYCANCEL:
			AddButton (IDRETRY, hInst);
			AddButton (IDCANCEL, hInst);
			break;

		case MB_CANCELTRYCONTINUE:
			AddButton (IDCANCEL, hInst);
			AddButton (IDTRYAGAIN, hInst);
			AddButton (IDCONTINUE, hInst);
			break;
		}

		if (m_Params.dwStyle & MB_HELP)
		{
			AddButton (IDHELP, hInst);
		}
	}

	for (int iBtn = 0; iBtn < BCGP_MSGBOXPARAMS::UserButtonsCount; ++iBtn)
	{
		if (m_Params.uiUserButtonID [iBtn] != 0)
		{
			LPCTSTR pszCaption = m_Params.lpszUserButtonText[iBtn];
			AddButton (m_Params.uiUserButtonID [iBtn], pszCaption == NULL ? hInst : m_Params.hInstance, pszCaption);
		}
	}

	if (hInst != NULL)
	{
		::FreeLibrary (hInst);
	}

	int nBtnCount = (int)m_arrButtons.GetSize ();
	ASSERT (nBtnCount != 0);

	if (nBtnCount == 1 && (m_Params.dwStyle & MB_TYPEMASK) == MB_OK)
	{
		m_bHasCancelButton = true;
	}

	// Calculate window size

	CRect rectScreen;
	
	MONITORINFO mi;
	mi.cbSize = sizeof (MONITORINFO);
	
	if (GetMonitorInfo(MonitorFromWindow(GetSafeHwnd(), MONITOR_DEFAULTTONEAREST), &mi))
	{
		rectScreen = mi.rcWork;
	}
	else
	{
		::SystemParametersInfo (SPI_GETWORKAREA, 0, &rectScreen, 0);
	}

	int nMinClientWidth = max(m_ptNextButtonPos.x, rectScreen.Width() / 10 - m_rectClientMargins.left - m_rectClientMargins.right);

	int nMaxMessageBoxWidth = rectScreen.Width() - m_rectClientMargins.left - m_rectClientMargins.right;
	if (!m_rectIcon.IsRectEmpty ())
	{
		nMaxMessageBoxWidth -= (m_rectIcon.Width () + m_cxIconSpacing);
	}

	CClientDC dc(this);

	CFont* pOldFont =  dc.SelectObject(&m_fntHeader);
	ASSERT_VALID(pOldFont);

	int nHeaderHeight = 0;

	if (!m_strMessageHeader.IsEmpty())
	{
		CSize sizeHeader = dc.GetTextExtent(m_strMessageHeader);
		
		if (sizeHeader.cx <= min(dc.GetTextExtent(m_strMessageCaption).cx, nMaxMessageBoxWidth))
		{
			m_rectHeader.SetRect(0, 0, sizeHeader.cx, sizeHeader.cy);
		}
		else
		{
			for (int nTextWidth = min(dc.GetTextExtent(m_strMessageCaption).cx, nMaxMessageBoxWidth); nTextWidth < nMaxMessageBoxWidth; nTextWidth += 10)
			{
				m_rectHeader.SetRect (0, 0, nTextWidth, 32676);

				int nHeight = dc.DrawText(m_strMessageHeader, &m_rectHeader, DT_NOPREFIX | (m_Params.bHeaderWordBreak ? DT_WORDBREAK : 0) | DT_CALCRECT);
				
				if (m_rectHeader.Width() > 5 * nHeight || m_rectHeader.Width() >= nMaxMessageBoxWidth)
				{
					m_rectHeader.bottom = m_rectHeader.top + nHeight;
					break;
				}
			}
		}

		nHeaderHeight = m_rectHeader.Height() + dc.GetTextExtent(m_strMessageHeader).cy / 2;
	}
	else
	{
		m_rectHeader.SetRectEmpty();
	}

	dc.SelectObject(&m_fntText);

	m_dwDrawTextFlags = DT_NOPREFIX | (m_Params.bMessageWordBreak ? DT_WORDBREAK : 0) | DT_EXPANDTABS;

	CSize sizeText = dc.GetTextExtent(m_strMessageText);

	if (sizeText.cx <= m_rectHeader.Width() && m_strMessageText.Find(_T("\n")) < 0)
	{
		m_rectText.SetRect(0, nHeaderHeight, sizeText.cx, nHeaderHeight + sizeText.cy);
	}
	else
	{
		for (int nTextWidth = min(dc.GetTextExtent(m_strMessageCaption).cx, nMaxMessageBoxWidth); nTextWidth < nMaxMessageBoxWidth; nTextWidth += 10)
		{
			m_rectText.SetRect (0, nHeaderHeight, nTextWidth, 32676);
			int nHeight = dc.DrawText(m_strMessageText, &m_rectText, m_dwDrawTextFlags | DT_CALCRECT);

			if (m_rectText.Width() > nHeight * (m_Params.bShowCheckBox ? 6 : 5) ||
				m_rectText.Width() >= nMaxMessageBoxWidth)
			{
				m_rectText.bottom = m_rectText.top + nHeight;
				break;
			}
		}
	}

	m_dwDrawTextFlags |= (m_bRightAlignment ? DT_RIGHT : DT_LEFT);

	// Calculate minimal width for message box window

	if (m_Params.bShowCheckBox)
	{
		SIZE szCheckBox;
		::GetTextExtentPoint32 (dc.GetSafeHdc(), m_strCheckBox, m_strCheckBox.GetLength (), &szCheckBox);
		szCheckBox.cx += 3 * m_cyCheckBoxHeight;
		
		if (m_Icon.IsValid())
		{
			szCheckBox.cx += m_rectIcon.Width ();
		}

		nMinClientWidth = max (nMinClientWidth, szCheckBox.cx);
	}

	dc.SelectObject(pOldFont);

	CSize szTotal (max(m_rectHeader.Width(), m_rectText.Width()), nHeaderHeight + m_rectText.Height());
	
	if (m_Icon.IsValid())
	{
		CSize szTextOffset (m_rectIcon.Width () + m_cxIconSpacing, 0);

		if (m_rectIcon.Height () > m_rectText.Height ())
		{
			// Center text vertically against icon
			szTextOffset.cy = (m_rectIcon.Height () - m_rectText.Height()) / 2;
		}

		m_rectIcon.OffsetRect(0, nHeaderHeight);
		m_rectText.OffsetRect (szTextOffset);

		if (m_rectHeader.Width () < m_rectIcon.Width () + m_cxIconSpacing + m_rectText.Width())
		{
			szTotal.cx += m_rectIcon.Width () + m_cxIconSpacing;
		}

		szTotal.cy = max (szTotal.cy, m_rectIcon.Height () + nHeaderHeight);
	}

	if (szTotal.cx < nMinClientWidth)
	{
		szTotal.cx = nMinClientWidth;
	}

#ifndef _BCGSUITE_
	BOOL bDrawButtonsBanner = m_Params.bDrawButtonsBanner && !m_Params.bUseNativeControls && !globalData.IsHighContastMode() && CBCGPVisualManager::GetInstance()->IsDrawMessageBoxButtonsBanner();
#else
	BOOL bDrawButtonsBanner = TRUE;
#endif

	if (!IsVisualManagerStyle())
	{
		bDrawButtonsBanner = TRUE;
	}

	if (m_bRightAlignment)
	{
		m_rectText.OffsetRect (szTotal.cx - m_rectText.right, 0);

		if (!m_rectHeader.IsRectEmpty())
		{
			m_rectHeader.OffsetRect (szTotal.cx - m_rectHeader.right, 0);
		}
	}

	// Adjusting buttons

	int yButtons = szTotal.cy + m_rectClientMargins.top;

	// Add check box if necessary

	m_cySeparatorLine = -1;

	if (m_Params.bShowCheckBox)
	{
		int oldY = yButtons;
		yButtons += m_Params.bShowSeparator ? m_rectClientMargins.bottom : m_cyCheckBoxSpacing;

		int xCheckBox = m_rectClientMargins.left;
		if (m_Icon.IsValid())
		{
			xCheckBox += m_rectIcon.Width () + m_cxIconSpacing;
		}

		CRect rectCheckBox (xCheckBox, yButtons, m_rectClientMargins.left + szTotal.cx, yButtons + m_cyCheckBoxHeight);
		yButtons += m_cyCheckBoxHeight;

		if (m_Params.bShowSeparator)
		{
			m_cySeparatorLine = rectCheckBox.top;
			rectCheckBox.OffsetRect (0, m_cyCheckBoxSpacing);
			yButtons += 2 * m_cyCheckBoxSpacing - m_rectClientMargins.bottom;
			if (!bDrawButtonsBanner)
			{
				yButtons += m_rectButtonsMargins.top;
			}
		}

		m_wndDontAskCheckBox.Create (m_strCheckBox,
			WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX,
			rectCheckBox, this, (UINT)(WORD)IDC_STATIC);

		m_wndDontAskCheckBox.SetFont(&m_fntText);

		if (m_Params.bIsChecked)
		{
			m_wndDontAskCheckBox.SetCheck(TRUE);
		}

		szTotal.cy += yButtons - oldY;
	}

	yButtons += m_rectClientMargins.bottom;
	szTotal.cy += m_rectClientMargins.bottom;

	if (bDrawButtonsBanner)
	{
		yButtons += m_rectButtonsMargins.top;
		szTotal.cy += m_rectButtonsMargins.top;
	}

	m_rectIcon.OffsetRect (m_rectClientMargins.left, m_rectClientMargins.top);
	m_rectText.OffsetRect (m_rectClientMargins.left, m_rectClientMargins.top);

	if (!m_rectHeader.IsRectEmpty())
	{
		m_rectHeader.OffsetRect (m_rectClientMargins.left, m_rectClientMargins.top);
	}

	const BOOL bRTL = GetExStyle() & WS_EX_LAYOUTRTL;

	CRect rectButtons(m_rectButtonsMargins.left, yButtons, m_rectClientMargins.left + m_rectClientMargins.right - m_rectButtonsMargins.right + szTotal.cx, yButtons);
	AdjustControls (&m_arrButtons[0], (UINT)m_arrButtons.GetSize (), rectButtons, TA_RIGHT, bRTL);

	szTotal.cy += m_szButton.cy; // place for buttons

	CRect rectClient (0, 0, szTotal.cx, szTotal.cy);
	rectClient.InflateRect (m_rectClientMargins.left, m_rectClientMargins.top, m_rectClientMargins.right, m_rectButtonsMargins.bottom);

	AdjustWindowRectEx (&rectClient, 
		::GetWindowLong (GetSafeHwnd (), GWL_STYLE) | WS_CAPTION | WS_BORDER,
		false,
		::GetWindowLong (GetSafeHwnd (), GWL_EXSTYLE)
		);

	SetWindowPos(NULL, 0, 0, rectClient.Width (), rectClient.Height (), SWP_NOMOVE | SWP_NOOWNERZORDER | SWP_NOZORDER);

	if (!m_bHasCancelButton)
	{
		GetSystemMenu (FALSE)->EnableMenuItem (SC_CLOSE, MF_GRAYED);
		UpdateCaptionButtons();
	}

	if (bDrawButtonsBanner)
	{
		m_rectButtons = CRect(0, yButtons - m_rectButtonsMargins.top, rectClient.Width(), rectClient.Height());
	}
	else
	{
		m_rectButtons.SetRectEmpty();
	}
}

void CBCGPMessageBox::OnButton (UINT nID)
{
	if (nID == IDHELP)
	{
		HELPINFO hinfo;
		memset((void*)&hinfo, 0, sizeof(hinfo));
		hinfo.cbSize = sizeof(hinfo);

		OnHelp(0, (LPARAM)&hinfo);
		return;
	}

	BOOL bCheckBox = m_wndDontAskCheckBox.GetSafeHwnd() != NULL;
	if (bCheckBox && (UINT)m_wndDontAskCheckBox.GetDlgCtrlID() == nID)
	{
		return;
	}

	if (nID == IDCANCEL && !m_bHasCancelButton)
	{
		return;
	}

	if (m_Params.puiAutoRespond != NULL)
	{
		*m_Params.puiAutoRespond = 0;

		if (bCheckBox && m_wndDontAskCheckBox.IsChecked ())
		{
			*m_Params.puiAutoRespond = nID;
		}
	}

	if ((m_Params.dwStyle & MB_TYPEMASK) == MB_OK && m_arrButtons.GetSize () == 1)
	{
		// If 'OK' is the only button force modal result to IDOK
		nID = IDOK;
	}

	EndDialog (nID);
}

LRESULT CBCGPMessageBox::OnHelp (WPARAM, LPARAM lParam)
{
	if (m_Params.dwContextHelpId != 0)
	{
		SendMessage(WM_COMMANDHELP, 0, m_Params.dwContextHelpId);
		return 1L;
	}

	HELPINFO hinfo;
	memcpy (&hinfo, (void*)lParam, sizeof(hinfo));
	hinfo.dwContextId = GetWindowContextHelpId ();

	if (m_Params.lpfnMsgBoxCallback)
	{
		m_Params.lpfnMsgBoxCallback (&hinfo);
	}
	else if (m_hwndOwner != NULL)
	{
		::SendMessage(m_hwndOwner, WM_HELP, 0, (LPARAM)&hinfo);
	}

	return 1L;
}

LRESULT CBCGPMessageBox::OnGetDefID (WPARAM, LPARAM)
{
	WORD nID = 0;
	for (int i = 0; i < (int)m_arrButtons.GetSize(); i++)
	{
		CBCGPButtonMB* pBtn = DYNAMIC_DOWNCAST(CBCGPButtonMB, m_arrButtons[i]);
		if (pBtn->GetSafeHwnd() != NULL && ::IsWindow(pBtn->GetSafeHwnd()) && pBtn->IsWindowVisible())
		{
			if (pBtn->m_bDefaultButton)
			{
				nID = (WORD)pBtn->GetDlgCtrlID();
				break;
			}
		}
	}
	
	if (nID == 0)
	{
		return 0L;
	}
	
	return MAKELRESULT(nID, DC_HASDEFID);
}

LRESULT CBCGPMessageBox::OnSetDefID (WPARAM wParam, LPARAM)
{
	const int nCount = (int)m_arrButtons.GetSize();
	
	if (nCount > 1)
	{
		const int nID = (int)wParam;
		
		for (int i = 0; i < (int)m_arrButtons.GetSize(); i++)
		{
			CBCGPButtonMB* pBtn = DYNAMIC_DOWNCAST(CBCGPButtonMB, m_arrButtons[i]);
			if (pBtn->GetSafeHwnd() != NULL && ::IsWindow(pBtn->GetSafeHwnd()) && pBtn->IsWindowVisible())
			{
				BOOL bDefaultButton = pBtn->GetDlgCtrlID() == nID;
				
				if ((bDefaultButton && !pBtn->m_bDefaultButton) ||
					(!bDefaultButton && pBtn->m_bDefaultButton))
				{
					pBtn->m_bDefaultButton = bDefaultButton;
					pBtn->Invalidate();
				}
			}
		}
	}
	
	return 1L;
}

BOOL CBCGPMessageBox::PreTranslateMessage(MSG* pMsg) 
{
#ifdef _UNICODE
#define _TCF_TEXT	CF_UNICODETEXT
#else
#define _TCF_TEXT	CF_TEXT
#endif
	
	if (pMsg->message == WM_KEYDOWN)
	{
		UINT nChar = (UINT) pMsg->wParam;
		BOOL bIsCtrl = (::GetAsyncKeyState (VK_CONTROL) & 0x8000);
		
		if (bIsCtrl && (nChar == _T('C') || nChar == VK_INSERT))
		{
			if (OpenClipboard ())
			{
				::EmptyClipboard ();
				
				LPCTSTR szDelimiter = _T("---------------------------\r\n");
				LPCTSTR szEOL = _T("\r\n");
				CString strText;
				strText = szDelimiter;
				strText += m_strMessageCaption + szEOL;
				strText += szDelimiter;

				if (!m_strMessageHeader.IsEmpty())
				{
					strText += m_strMessageHeader + szEOL;
					strText += szEOL;
				}

				strText += m_strMessageText + szEOL;
				strText += szDelimiter;

				if (m_Params.bShowCheckBox)
				{
					CString strChecked = m_wndDontAskCheckBox.IsChecked() ? _T("[X] ") : _T("[ ] ");
					CString strCheckBox = m_strCheckBox;
					
					globalUtils.RemoveSingleAmp(strCheckBox);

					strText += strChecked + strCheckBox + szEOL;
					strText += szDelimiter;
				}

				for (int i = 0; i < (int)m_arrButtons.GetSize (); i++)
				{
					CString strButton;
					m_arrButtons[i]->GetWindowText (strButton);
					
					globalUtils.RemoveSingleAmp(strButton);

					strText += _T("[ ") + strButton + _T(" ] ");
				}

				strText += szEOL;
				strText += szDelimiter;
			
				HGLOBAL hClipbuffer = ::GlobalAlloc (GMEM_DDESHARE, (strText.GetLength () + 1) * sizeof (TCHAR));
				if (hClipbuffer != NULL)
				{
					LPTSTR lpszBuffer = (LPTSTR)::GlobalLock (hClipbuffer);

					lstrcpy (lpszBuffer, (LPCTSTR) strText);

					::GlobalUnlock (hClipbuffer);
					::SetClipboardData (_TCF_TEXT, hClipbuffer);
				}
				
				::CloseClipboard();
			}
		}
	}
	
	return CBCGPDialog::PreTranslateMessage(pMsg);
}

#if _MSC_VER >= 1300
HRESULT CBCGPMessageBox::get_accChildCount(long *pcountChildren)
{
	if (!pcountChildren)
	{
		return E_INVALIDARG;
	}

	if (m_rectText.IsRectEmpty () && m_rectHeader.IsRectEmpty())
	{
		return CBCGPDialog::get_accChildCount(pcountChildren);
	}

	long nCount = 0;	
	CBCGPDialog::get_accChildCount(&nCount);

	*pcountChildren = nCount + 1;
	m_nAccTextIndex = nCount + 1;

	if (!m_rectHeader.IsRectEmpty())
	{
		*pcountChildren = nCount + 2;
		m_nAccHeaderIndex = nCount + 2;
	}

	return S_OK;
}

HRESULT CBCGPMessageBox::get_accChild(VARIANT varChild, IDispatch **ppdispChild)
{
	if (m_rectText.IsRectEmpty ())
	{
		return GetAccessibleChild(varChild, ppdispChild);
	}

	if (varChild.lVal < m_nAccTextIndex)
	{
		return GetAccessibleChild(varChild, ppdispChild);
	}

	return S_FALSE;
}

HRESULT CBCGPMessageBox::get_accName(VARIANT varChild, BSTR *pszName)
{
	if (m_nAccTextIndex > 0 && varChild.vt == VT_I4 && varChild.lVal == m_nAccTextIndex)
	{
		*pszName = m_strMessageText.AllocSysString();
		return S_OK;
	}

	if (m_nAccHeaderIndex > 0 && varChild.vt == VT_I4 && varChild.lVal == m_nAccHeaderIndex)
	{
		*pszName = m_strMessageHeader.AllocSysString();
		return S_OK;
	}

	return CBCGPDialog::get_accName(varChild, pszName);
}

HRESULT CBCGPMessageBox::accLocation(long *pxLeft, long *pyTop, long *pcxWidth, long *pcyHeight, VARIANT varChild)
{
	if (!pxLeft || !pyTop || !pcxWidth || !pcyHeight)
    {
        return E_INVALIDARG;
    }

	if (m_nAccTextIndex > 0 && varChild.vt == VT_I4 && varChild.lVal == m_nAccTextIndex)
	{
		CRect rcTextRect = m_rectText;
		ClientToScreen(&rcTextRect);
		
		*pxLeft = rcTextRect.left;
		*pyTop = rcTextRect.top;
		*pcxWidth = rcTextRect.Width();
		*pcyHeight = rcTextRect.Height();

		return S_OK;
	}

	if (!m_rectHeader.IsRectEmpty () && varChild.vt == VT_I4 && varChild.lVal == m_nAccHeaderIndex)
	{
		CRect rcTextRect = m_rectHeader;
		ClientToScreen(&rcTextRect);
		
		*pxLeft = rcTextRect.left;
		*pyTop = rcTextRect.top;
		*pcxWidth = rcTextRect.Width();
		*pcyHeight = rcTextRect.Height();
		
		return S_OK;
	}

	return CBCGPDialog::accLocation(pxLeft, pyTop, pcxWidth, pcyHeight, varChild);
}

HRESULT CBCGPMessageBox::accHitTest(long xLeft, long yTop, VARIANT *pvarChild)
{
	CPoint pt (xLeft, yTop);
	ScreenToClient (&pt);

	if (m_rectText.PtInRect(pt))
	{
		pvarChild->vt = VT_I4;
		pvarChild->lVal = m_nAccTextIndex;
		return S_OK;
	}

	if (m_rectHeader.PtInRect(pt))
	{
		pvarChild->vt = VT_I4;
		pvarChild->lVal = m_nAccHeaderIndex;
		return S_OK;
	}

	return CBCGPDialog::accHitTest(xLeft, yTop, pvarChild);
}

HRESULT CBCGPMessageBox::get_accRole(VARIANT varChild, VARIANT *pvarRole)
{
	if (pvarRole == NULL)
    {
        return E_INVALIDARG;
    }

	if (m_nAccTextIndex > 0 && varChild.vt == VT_I4 && varChild.lVal >= m_nAccTextIndex)
	{
		pvarRole->vt = VT_I4;
		pvarRole->lVal = ROLE_SYSTEM_STATICTEXT;
		return S_OK;
	}

	return CBCGPDialog::get_accRole(varChild, pvarRole);
}

HRESULT CBCGPMessageBox::get_accState(VARIANT varChild, VARIANT *pvarState)
{
	if (m_nAccTextIndex > 0 && varChild.vt == VT_I4 && (varChild.lVal >= m_nAccTextIndex))
	{
		pvarState->vt = VT_I4;
		pvarState->lVal = STATE_SYSTEM_READONLY;
		return S_OK; 
	}

	return CBCGPDialog::get_accState(varChild, pvarState);
}

HRESULT CBCGPMessageBox::get_accDefaultAction(VARIANT varChild, BSTR *pszDefaultAction)
{
	if (m_nAccTextIndex > 0 && varChild.vt == VT_I4 && varChild.lVal >= m_nAccTextIndex)
	{
		return S_FALSE;
	}

	return CBCGPDialog::get_accDefaultAction(varChild, pszDefaultAction);
}

HRESULT CBCGPMessageBox::get_accDescription(VARIANT varChild, BSTR *pszDescription)
{
	if (m_nAccTextIndex > 0 && varChild.vt == VT_I4 && varChild.lVal >= m_nAccTextIndex)
	{
		return S_FALSE;
	}

	return CBCGPDialog::get_accDescription(varChild, pszDescription);
}

HRESULT CBCGPMessageBox::get_accHelp(VARIANT varChild, BSTR* pszHelp)
{
	if (m_nAccTextIndex > 0 && varChild.vt == VT_I4 && varChild.lVal >= m_nAccTextIndex)
	{
		return S_FALSE;
	}

	return CBCGPDialog::get_accHelp(varChild, pszHelp);
}

HRESULT CBCGPMessageBox::get_accHelpTopic(BSTR* pszHelpFile, VARIANT varChild, long* pidTopic)
{
	if (m_nAccTextIndex > 0 && varChild.vt == VT_I4 && varChild.lVal >= m_nAccTextIndex)
	{
		return S_FALSE;
	}

	return CBCGPDialog::get_accHelpTopic(pszHelpFile, varChild, pidTopic);
}

HRESULT CBCGPMessageBox::get_accKeyboardShortcut(VARIANT varChild, BSTR* pszKeyboardShortcut)
{
	if (m_nAccTextIndex > 0 && varChild.vt == VT_I4 && varChild.lVal >= m_nAccTextIndex)
	{
		return S_FALSE;
	}

	return CBCGPDialog::get_accKeyboardShortcut(varChild, pszKeyboardShortcut);
}

HRESULT CBCGPMessageBox::get_accValue(VARIANT varChild, BSTR *pszValue)
{
	if (m_nAccTextIndex > 0 && varChild.vt == VT_I4 && varChild.lVal >= m_nAccTextIndex)
	{
		return S_FALSE;
	}

	return CBCGPDialog::get_accValue(varChild, pszValue);
}
#endif

int BCGPMessageBox (HWND hWnd, LPCTSTR lpszText, LPCTSTR lpszCaption, UINT uType, LPCTSTR lpszCheckBoxLabel, BOOL* pCheckBoxResult, LPCTSTR lpszHeader)
{
	return BCGPMessageBoxEx (hWnd, lpszText, lpszCaption, uType, LANG_NEUTRAL, lpszCheckBoxLabel, pCheckBoxResult, lpszHeader);
}

int BCGPMessageBox (LPCTSTR lpszText, UINT nType, LPCTSTR lpszCheckBoxLabel, BOOL* pCheckBoxResult, LPCTSTR lpszHeader)
{
	return BCGPMessageBoxEx (::GetActiveWindow (), lpszText, AfxGetAppName (), nType, LANG_NEUTRAL, lpszCheckBoxLabel, pCheckBoxResult, lpszHeader);
}

int BCGPMessageBox (UINT nIDPrompt, UINT nType, UINT nIDCheckBoxPrompt, BOOL* pCheckBoxResult, LPCTSTR lpszHeader)
{
	CString str;
	str.LoadString(nIDPrompt);

	if (nIDCheckBoxPrompt != 0 && pCheckBoxResult != NULL)
	{
		CString strCheckBox;
		strCheckBox.LoadString(nIDCheckBoxPrompt);

		return BCGPMessageBoxEx (::GetActiveWindow (), str, AfxGetAppName (), nType, LANG_NEUTRAL, strCheckBox, pCheckBoxResult, lpszHeader);
	}

	return BCGPMessageBoxEx (::GetActiveWindow (), str, AfxGetAppName (), nType, LANG_NEUTRAL, NULL, NULL, lpszHeader);
}

int BCGPMessageBoxEx (HWND hWnd, LPCTSTR lpszText, LPCTSTR lpszCaption, UINT uType, WORD wLanguageId, LPCTSTR lpszCheckBoxLabel, BOOL* pCheckBoxResult, LPCTSTR lpszHeader)
{
	BCGP_MSGBOXPARAMS params;

	params.cbSize = sizeof(params);
	params.hwndOwner = hWnd;
	params.hInstance = NULL;
	params.lpszText = lpszText;
	params.lpszHeader = lpszHeader;
	params.lpszCaption = lpszCaption;
	params.dwStyle = uType;
	params.lpszIcon = NULL;
	params.dwContextHelpId = 0;
	params.lpfnMsgBoxCallback = NULL;
	params.dwLanguageId = wLanguageId;
	params.bUseNativeCaption = CBCGPMessageBox::m_bUseNativeCaption;
	params.bUseNativeControls = CBCGPMessageBox::m_bUseNativeControls;

	if (lpszCheckBoxLabel != NULL && pCheckBoxResult != NULL)
	{
		params.bShowCheckBox = TRUE;
		params.lpszCheckBoxText = lpszCheckBoxLabel;
		params.bIsChecked = *pCheckBoxResult;
		params.puiAutoRespond = (UINT*)pCheckBoxResult;
	}

	return BCGPMessageBoxIndirect (&params);
}

int BCGPMessageBoxIndirect (const MSGBOXPARAMS* pMsgBoxParams)
{
	CBCGPMessageBox mb (pMsgBoxParams);
	return (int)mb.DoModal ();
}

int BCGPMessageBoxIndirect (const BCGP_MSGBOXPARAMS* pMsgBoxParams)
{
	CBCGPMessageBox mb (pMsgBoxParams);
	return (int)mb.DoModal ();
}
