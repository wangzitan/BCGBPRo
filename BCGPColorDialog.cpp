// BCGColorDialog.cpp : implementation file
// This is a part of the BCGControlBar Library
// Copyright (C) 1998-2018 BCGSoft Ltd.
// All rights reserved.
//
// This source code can be used, distributed or modified
// only under terms and conditions 
// of the accompanying license agreement.
//

#include "stdafx.h"

#include "BCGCBPro.h"
#include "BCGPLocalResource.h"
#include "BCGPColorDialog.h"
#include "PropertySheetCtrl.h"
#include "BCGGlobals.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

BOOL CBCGPColorDialog::s_bProcessIdle = FALSE;
BOOL CBCGPColorDialog::s_bStayForeground = FALSE;

/////////////////////////////////////////////////////////////////////////////
// CBCGPScreenWnd window

class CBCGPScreenWnd : public CWnd
{
// Construction
public:
	CBCGPScreenWnd();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBCGPScreenWnd)
	public:
	virtual BOOL Create(CBCGPColorDialog* pColorDlg);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CBCGPScreenWnd();

	// Generated message map functions
protected:
	//{{AFX_MSG(CBCGPScreenWnd)
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	CBCGPColorDialog*	m_pColorDlg;
};


/////////////////////////////////////////////////////////////////////////////
// CBCGPColorDialog dialog

IMPLEMENT_DYNAMIC(CBCGPColorDialog, CBCGPDialog)

CBCGPColorDialog::CBCGPColorDialog (COLORREF clrInit, DWORD /*dwFlags - reserved */, 
					CWnd* pParentWnd, HPALETTE hPal)
	: CBCGPDialog(CBCGPColorDialog::IDD, pParentWnd)
{
	//{{AFX_DATA_INIT(CBCGPColorDialog)
	//}}AFX_DATA_INIT

	m_pColourSheetOne = NULL;
	m_pColourSheetTwo = NULL;

	m_CurrentColor = m_NewColor = clrInit;
	m_pPropSheet = NULL;
	m_bIsMyPalette = TRUE;
	m_pPalette = NULL;
	
	if (hPal != NULL)
	{
		m_pPalette = CPalette::FromHandle (hPal);
		m_bIsMyPalette = FALSE;
	}

	m_bPickerMode = FALSE;
	m_bIsLocal = TRUE;

	EnableVisualManagerStyle (globalData.m_bUseVisualManagerInBuiltInDialogs, TRUE);
}

CBCGPColorDialog::~CBCGPColorDialog ()
{
	if (m_pColourSheetOne != NULL)
	{
		delete m_pColourSheetOne;
	}

	if (m_pColourSheetTwo != NULL)
	{
		delete m_pColourSheetTwo;
	}
}

void CBCGPColorDialog::DoDataExchange(CDataExchange* pDX)
{
	CBCGPDialog::DoDataExchange(pDX);
	
	//{{AFX_DATA_MAP(CBCGPColorDialog)
	DDX_Control(pDX, IDC_BCGBARRES_COLOR_SELECT, m_btnColorSelect);
	DDX_Control(pDX, IDC_BCGBARRES_STATICPLACEHOLDER, m_wndStaticPlaceHolder);
	DDX_Control(pDX, IDC_BCGBARRES_COLOURPLACEHOLDER, m_wndColors);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CBCGPColorDialog, CBCGPDialog)
	//{{AFX_MSG_MAP(CBCGPColorDialog)
	ON_WM_DESTROY()
	ON_WM_SYSCOLORCHANGE()
	ON_BN_CLICKED(IDC_BCGBARRES_COLOR_SELECT, OnColorSelect)
	ON_WM_SETCURSOR()
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONDOWN()
	//}}AFX_MSG_MAP
	ON_WM_ACTIVATEAPP()
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPColorDialog message handlers

BOOL CBCGPColorDialog::OnInitDialog()
{
	CBCGPLocalResource locaRes;

	CBCGPDialog::OnInitDialog();

	if (AfxGetMainWnd () != NULL && 
		(AfxGetMainWnd ()->GetExStyle () & WS_EX_LAYOUTRTL))
	{
		ModifyStyleEx (0, WS_EX_LAYOUTRTL);
	}

	if (globalData.m_nBitsPerPixel < 8)	// 16 colors, call standard dialog
	{
		CColorDialog dlg (m_CurrentColor, CC_FULLOPEN | CC_ANYCOLOR);
		int nResult = (int) dlg.DoModal ();
		m_NewColor = dlg.GetColor ();
		EndDialog (nResult);

		return TRUE;
	}

	if (m_pPalette == NULL)
	{
		m_pPalette = new CPalette ();
		RebuildPalette ();
	}

	m_wndColors.SetType (CBCGPColorPickerCtrl::CURRENT);
	m_wndColors.SetPalette (m_pPalette);

	m_wndColors.SetOriginalColor (m_CurrentColor);
	m_wndColors.SetRGB (m_NewColor);

	// Create property sheet.
	m_pPropSheet = new CBCGPPropertySheetCtrl(_T(""), this);
	ASSERT(m_pPropSheet);

	m_pPropSheet->EnableVisualManagerStyle(globalData.m_bUseVisualManagerInBuiltInDialogs);

	if (globalData.m_bUseVisualManagerInBuiltInDialogs)
	{
		m_pPropSheet->SetLook (CBCGPPropertySheet::PropSheetLook_OneNoteTabs);
	}

	m_pColourSheetOne = new CBCGPColorPage1;
	m_pColourSheetTwo = new CBCGPColorPage2;

	// Set parent dialog.
	m_pColourSheetOne->m_pDialog = this;
	m_pColourSheetTwo->m_pDialog = this;

	m_pPropSheet->AddPage (m_pColourSheetOne);
	m_pPropSheet->AddPage (m_pColourSheetTwo);

	// Retrieve the location of the window
	CRect rectListWnd;
	m_wndStaticPlaceHolder.GetWindowRect(rectListWnd);
	ScreenToClient(rectListWnd);
	
	if (!m_pPropSheet->Create(this, WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE, 0))
	{
		TRACE0 ("CBCGPColorDialog::OnInitDialog(): Can't create the property sheet.....\n");
	}

	m_pPropSheet->SetActivePage(0);
	
	CRect rectPage1Initial;
	m_pPropSheet->GetWindowRect(rectPage1Initial);

	m_pPropSheet->SetWindowPos(NULL, rectListWnd.left, rectListWnd.top, rectListWnd.Width(),
				rectListWnd.Height(), SWP_NOZORDER | SWP_NOACTIVATE);

	SetPageOne (GetRValue (m_CurrentColor), 
				GetGValue (m_CurrentColor),
				GetBValue (m_CurrentColor));

	SetPageTwo (GetRValue (m_CurrentColor),
				GetGValue (m_CurrentColor),
				GetBValue (m_CurrentColor));

	m_btnColorSelect.SetImage (IDB_BCGBARRES_COLOR_PICKER);

	CRect rectPage1Current;
	m_pPropSheet->GetWindowRect(rectPage1Current);

	int cy = rectPage1Initial.Height() - rectPage1Current.Height();
	if (cy > 0)
	{
		// Resize the dialog:
		CRect rectThis;
		GetWindowRect(rectThis);

		SetWindowPos(NULL, -1, -1, rectThis.Width(), rectThis.Height() + cy, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
		m_pPropSheet->SetWindowPos(NULL, -1, -1, rectListWnd.Width(), rectListWnd.Height() + cy, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
	}

	m_hcurPicker = AfxGetApp()->LoadCursor (IDC_BCGBARRES_COLOR);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CBCGPColorDialog::SetCurrentColor(COLORREF rgb)
{
	m_CurrentColor = rgb;
}

void CBCGPColorDialog::SetPageOne(BYTE R, BYTE G, BYTE B)
{
	m_pColourSheetOne->m_hexpicker.SelectCellHexagon(R, G, B);
	m_pColourSheetOne->m_hexpicker_greyscale.SelectCellHexagon(R, G, B);
}

void CBCGPColorDialog::SetPageTwo(BYTE R, BYTE G, BYTE B)
{
	m_pColourSheetTwo->Setup (R, G, B);
}

void CBCGPColorDialog::OnDestroy() 
{
	if (m_bIsMyPalette && m_pPalette != NULL)
	{
		m_pPalette->DeleteObject ();
		delete m_pPalette;
		m_pPalette = NULL;
	}

	CBCGPDialog::OnDestroy();
}

void CBCGPColorDialog::SetNewColor (COLORREF rgb)
{
	COLORREF clrNewSaved = m_NewColor;

	m_NewColor = rgb;

	if (globalData.m_nBitsPerPixel == 8) // 256 colors
	{
		ASSERT (m_pPalette != NULL);

		UINT uiPalIndex = m_pPalette->GetNearestPaletteIndex (rgb);
		m_wndColors.SetRGB (PALETTEINDEX (uiPalIndex));
	}
	else
	{
		m_wndColors.SetRGB (rgb);
	}

	m_wndColors.Invalidate ();
	m_wndColors.UpdateWindow ();

	OnNewColorChanged(clrNewSaved, m_NewColor);
}

void CBCGPColorDialog::OnSysColorChange() 
{
	CBCGPDialog::OnSysColorChange();

	globalData.UpdateSysColors ();
	
	if (m_bIsMyPalette)
	{
		if (globalData.m_nBitsPerPixel < 8)	// 16 colors, call standard dialog
		{
			ShowWindow (SW_HIDE);

			CColorDialog dlg (m_CurrentColor, CC_FULLOPEN | CC_ANYCOLOR);
			int nResult = (int) dlg.DoModal ();
			m_NewColor = dlg.GetColor ();
			EndDialog (nResult);
		}
		else
		{
			::DeleteObject (m_pPalette->Detach ());
			RebuildPalette ();

			Invalidate ();
			UpdateWindow ();
		}
	}
}

void CBCGPColorDialog::RebuildPalette ()
{
	ASSERT (m_pPalette->GetSafeHandle () == NULL);

	// Create palette:
	CClientDC dc (this);

	int nColors = 256;	// Use 256 first entries
	UINT nSize = sizeof(LOGPALETTE) + (sizeof(PALETTEENTRY) * nColors);
	LOGPALETTE *pLP = (LOGPALETTE *) new BYTE[nSize];

	pLP->palVersion = 0x300;
	pLP->palNumEntries = (USHORT) nColors;

	::GetSystemPaletteEntries (dc.GetSafeHdc (), 0, nColors, pLP->palPalEntry);

	m_pPalette->CreatePalette (pLP);

	delete[] pLP;
}

void CBCGPColorDialog::OnColorSelect() 
{
	if (m_bPickerMode)
	{
		return;
	}

	CWinThread* pCurrThread = ::AfxGetThread ();
	if (pCurrThread == NULL)
	{
		ASSERT (FALSE);
		return;
	}

	MSG msg;
	m_bPickerMode = TRUE;;
	::SetCursor (m_hcurPicker);

	CBCGPScreenWnd* pScreenWnd = new CBCGPScreenWnd;
	if (!pScreenWnd->Create (this))
	{
		return;
	}

	SetForegroundWindow ();
	BringWindowToTop ();

	SetCapture ();

	COLORREF colorSaved = m_NewColor;

	const BOOL bProcessIdle = CBCGPColorDialog::s_bProcessIdle && (GetStyle() & DS_NOIDLEMSG) == 0;
	BOOL bIdle = TRUE;
	LONG lIdleCount = 0;

	while (m_bPickerMode)
	{
		if (bProcessIdle)
		{
			while (bIdle && !::PeekMessage(&msg, NULL, NULL, NULL, PM_NOREMOVE))
			{
				bIdle = pCurrThread->OnIdle(lIdleCount++);
			}
		}

		while (m_bPickerMode && ::PeekMessage(&msg, NULL, NULL, NULL, PM_NOREMOVE))
		{
			if (msg.message == WM_KEYDOWN)
			{
				switch (msg.wParam)
				{
				case VK_RETURN:
					m_bPickerMode = FALSE;
					break;

				case VK_ESCAPE:
					SetNewColor (colorSaved);
					m_bPickerMode = FALSE;
					break;
				}
			}
			else if (msg.message == WM_RBUTTONDOWN || msg.message == WM_MBUTTONDOWN)
			{
				m_bPickerMode = FALSE;
			}

			if (m_bPickerMode)
			{
				if (!pCurrThread->PumpMessage())
				{
					AfxPostQuitMessage(0);
					return;
				}

				if (bProcessIdle)
				{
					if (pCurrThread->IsIdleMessage(&msg))
					{
						bIdle = TRUE;
						lIdleCount = 0;
					}
				}
			}
			else
			{
				// remove last message
				::GetMessage(&msg, NULL, NULL, NULL);
			}
		}

		if (!m_bPickerMode)
		{
			WaitMessage();
		}
	}

	ReleaseCapture ();
	pScreenWnd->DestroyWindow();
	delete pScreenWnd;

	m_bPickerMode = FALSE;

	if (!CBCGPColorDialog::s_bStayForeground)
	{
		CWnd* pMainWnd = AfxGetMainWnd();
		if (pMainWnd != NULL)
		{
			pMainWnd->SetForegroundWindow();
		}
	}
}

BOOL CBCGPColorDialog::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	if (m_bPickerMode)
	{
		::SetCursor (m_hcurPicker);
		return TRUE;
	}
	
	return CBCGPDialog::OnSetCursor(pWnd, nHitTest, message);
}

void CBCGPColorDialog::OnMouseMove(UINT nFlags, CPoint point) 
{
	if (m_bPickerMode)
	{
		ClientToScreen (&point);

		CClientDC dc(NULL);
		SetNewColor (dc.GetPixel(point));
	}

	CBCGPDialog::OnMouseMove(nFlags, point);
}

void CBCGPColorDialog::OnLButtonDown(UINT nFlags, CPoint point) 
{
	m_bPickerMode = FALSE;

	SetPageOne (GetRValue (m_NewColor), 
				GetGValue (m_NewColor),
				GetBValue (m_NewColor));

	SetPageTwo (GetRValue (m_NewColor),
				GetGValue (m_NewColor),
				GetBValue (m_NewColor));

	CBCGPDialog::OnLButtonDown(nFlags, point);
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPScreenWnd

CBCGPScreenWnd::CBCGPScreenWnd()
{
}

CBCGPScreenWnd::~CBCGPScreenWnd()
{
}


BEGIN_MESSAGE_MAP(CBCGPScreenWnd, CWnd)
	//{{AFX_MSG_MAP(CBCGPScreenWnd)
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONDOWN()
	ON_WM_SETCURSOR()
	ON_WM_ERASEBKGND()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CBCGPScreenWnd message handlers

BOOL CBCGPScreenWnd::Create(CBCGPColorDialog* pColorDlg) 
{
	CWnd* pWndDesktop = GetDesktopWindow ();
	if (pWndDesktop == NULL)
	{
		ASSERT (FALSE);
		return FALSE;
	}

	m_pColorDlg = pColorDlg;
	
	CRect rectScreen;
	pWndDesktop->GetWindowRect (rectScreen);

	CString strClassName = ::AfxRegisterWndClass (
			CS_SAVEBITS,
			AfxGetApp()->LoadCursor (IDC_BCGBARRES_COLOR),
			(HBRUSH)(COLOR_BTNFACE + 1), NULL);

	return CWnd::CreateEx (WS_EX_TOOLWINDOW, strClassName, _T(""), WS_VISIBLE | WS_POPUP, rectScreen, NULL, 0);
}

void CBCGPScreenWnd::OnMouseMove(UINT nFlags, CPoint point) 
{
	MapWindowPoints (m_pColorDlg, &point, 1);
	m_pColorDlg->SendMessage (WM_MOUSEMOVE, nFlags, MAKELPARAM (point.x, point.y));

	CWnd::OnMouseMove(nFlags, point);
}

void CBCGPScreenWnd::OnLButtonDown(UINT nFlags, CPoint point) 
{
	MapWindowPoints (m_pColorDlg, &point, 1);
	m_pColorDlg->SendMessage (WM_LBUTTONDOWN, nFlags, MAKELPARAM (point.x, point.y));
}

BOOL CBCGPScreenWnd::OnEraseBkgnd(CDC* /*pDC*/) 
{
	return TRUE;
}

BOOL CBCGPColorDialog::PreTranslateMessage(MSG* pMsg) 
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
				EmptyClipboard ();

				CString strText;
				strText.Format (_T("RGB(%d, %d, %d)"),
					GetRValue (m_NewColor),
					GetGValue (m_NewColor),
					GetBValue (m_NewColor));

				HGLOBAL hClipbuffer = ::GlobalAlloc (GMEM_DDESHARE, (strText.GetLength () + 1) * sizeof (TCHAR));
				if (hClipbuffer != NULL)
				{
					LPTSTR lpszBuffer = (LPTSTR) GlobalLock (hClipbuffer);

					lstrcpy (lpszBuffer, (LPCTSTR) strText);

					::GlobalUnlock (hClipbuffer);
					::SetClipboardData (_TCF_TEXT, hClipbuffer);
				}

				CloseClipboard();
			}
		}
	}
	
	return CBCGPDialog::PreTranslateMessage(pMsg);
}

#if (_MSC_VER < 1300)
void CBCGPColorDialog::OnActivateApp(BOOL bActive, HTASK param)
#else
void CBCGPColorDialog::OnActivateApp(BOOL bActive, DWORD param)
#endif
{
	CBCGPDialog::OnActivateApp(bActive, param);
	
	if (m_bPickerMode)
	{
		m_bPickerMode = FALSE;
	}
}
