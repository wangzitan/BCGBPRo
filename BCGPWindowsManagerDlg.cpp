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
// BCGPWindowsManagerDlg.cpp : implementation file
//

#include "stdafx.h"
#include "BCGCBPro.h"
#include "BCGPMDIFrameWnd.h"
#include "BCGPWindowsManagerDlg.h"
#include "BCGPMDIChildWnd.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPWindowsManagerDlg dialog

UINT WM_BCGWINDOW_HELP = ::RegisterWindowMessage(_T("WINDOW_HELP"));

CBCGPWindowsManagerDlg::CBCGPWindowsManagerDlg(CBCGPMDIFrameWnd* pMDIFrame, BOOL bHelpButton, BOOL bResizable)
	: CBCGPDialog(CBCGPWindowsManagerDlg::IDD, pMDIFrame),
	m_pMDIFrame (pMDIFrame),
	m_bHelpButton (bHelpButton)
{
	//{{AFX_DATA_INIT(CBCGPWindowsManagerDlg)
	//}}AFX_DATA_INIT
	ASSERT_VALID (m_pMDIFrame);

	m_bMDIActions = TRUE;

	EnableVisualManagerStyle (globalData.m_bUseVisualManagerInBuiltInDialogs, TRUE);

	if (bResizable)
	{
		EnableLayout();
	}
}


void CBCGPWindowsManagerDlg::DoDataExchange(CDataExchange* pDX)
{
	CBCGPDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CBCGPWindowsManagerDlg)
	DDX_Control(pDX, IDC_BCGBARRES_LIST, m_wndList);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CBCGPWindowsManagerDlg, CBCGPDialog)
	//{{AFX_MSG_MAP(CBCGPWindowsManagerDlg)
	ON_BN_CLICKED(IDC_BCGBARRES_ACTIVATE, OnActivate)
	ON_BN_CLICKED(IDC_BCGBARRES_SAVE, OnSave)
	ON_BN_CLICKED(IDC_BCGBARRES_CLOSE, OnClose)
	ON_BN_CLICKED(IDC_BCGBARRES_CASCADE, OnCascade)
	ON_BN_CLICKED(IDC_BCGBARRES_TILEHORZ, OnTilehorz)
	ON_BN_CLICKED(IDC_BCGBARRES_TILEVERT, OnTilevert)
	ON_BN_CLICKED(IDC_BCGBARRES_MINIMIZE, OnMinimize)
	ON_LBN_SELCHANGE(IDC_BCGBARRES_LIST, OnSelchangeBcgbarresList)
	ON_LBN_DBLCLK(IDC_BCGBARRES_LIST, OnActivate)
	ON_BN_CLICKED(ID_HELP, OnBcgbarresWindowHelp)
	ON_WM_HELPINFO()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBCGPWindowsManagerDlg message handlers

void CBCGPWindowsManagerDlg::MDIMessage (UINT uMsg, WPARAM flag)
{
	ASSERT_VALID (m_pMDIFrame);

	CWaitCursor wait;

	int nItems = m_wndList.GetCount();
	if (nItems <= 0)
	{
		return;
	}

	HWND hMDIClient = m_pMDIFrame->m_hWndMDIClient;

	BOOL bLocked = ::LockWindowUpdate (hMDIClient);


	for(int i = nItems - 1; i >= 0; i--)
	{
		HWND hWnd= (HWND) m_wndList.GetItemData(i);
		::ShowWindow (hWnd, m_wndList.GetSel (i) > 0 ? SW_RESTORE : SW_MINIMIZE);
	}

	::SendMessage (hMDIClient, uMsg, flag, 0);
	if (bLocked)
	{
		::LockWindowUpdate (NULL);
	}
}
//**************************************************************************************
void CBCGPWindowsManagerDlg::OnActivate() 
{
	if (!CBCGPMDIFrameWnd::m_bDisableSetRedraw)
	{
		GetParent ()->SetRedraw (FALSE);
	}

	if (m_wndList.GetSelCount() == 1)
	{
		int	index;	
		if (m_wndList.GetSelItems (1, &index) == 1)
		{						 
			DWORD_PTR dw = m_wndList.GetItemData (index);
			if (dw != (DWORD_PTR) LB_ERR)
			{
				WINDOWPLACEMENT	wndpl;
				wndpl.length = sizeof(WINDOWPLACEMENT);
				::GetWindowPlacement ((HWND) dw,&wndpl);

				if (wndpl.showCmd == SW_SHOWMINIMIZED)
				{
					::ShowWindow((HWND) dw,SW_RESTORE);
				}

				CBCGPMDIFrameWnd* pMDIFrame = m_pMDIFrame;
				
				CBCGPMDIChildWnd* pMDIChild = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, CWnd::FromHandlePermanent((HWND)dw));
				if (pMDIChild != NULL)
				{
					ASSERT_VALID(pMDIChild);

					if (pMDIChild->m_pMDIFrame != NULL && pMDIChild->m_pMDIFrame != pMDIFrame)
					{
						pMDIFrame = pMDIChild->m_pMDIFrame;

						if (pMDIFrame->IsIconic())
						{
							pMDIFrame->ShowWindow(SW_RESTORE);
						}

						pMDIFrame->SetActiveWindow();
					}
				}

				::SendMessage(pMDIFrame->m_hWndMDIClient,WM_MDIACTIVATE,(WPARAM) dw,0);
				EndDialog(IDOK);
			}
		}
	}

	if (!CBCGPMDIFrameWnd::m_bDisableSetRedraw)
	{
		GetParent ()->SetRedraw (TRUE);
		GetParent ()->RedrawWindow (NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE | RDW_ALLCHILDREN);
	}
}
//**************************************************************************************
void CBCGPWindowsManagerDlg::OnSave() 
{
	int nItems = m_wndList.GetCount();
	if (nItems > 0)
	{
		for (int i = 0; i < nItems; i++)
		{
			if (m_wndList.GetSel (i) > 0)
			{
				HWND hWnd=(HWND) m_wndList.GetItemData (i);

				if (m_lstSaveDisabled.Find (hWnd) == NULL)
				{
					CWnd* pWnd = CWnd::FromHandle (hWnd);
					CFrameWnd* pFrame = DYNAMIC_DOWNCAST (CFrameWnd, pWnd);

					if (pFrame != NULL)
					{
						CDocument *pDoc = pFrame->GetActiveDocument();
						if (pDoc != NULL)
						{
							ASSERT_VALID (pDoc);
							pDoc->DoFileSave();
						}
					}
				}
			}
		}
	}

	FillWindowList ();
	SelActive ();
	UpdateButtons ();
}
//**************************************************************************************
void CBCGPWindowsManagerDlg::OnClose() 
{
	BOOL bIsTearOffFrame = m_pMDIFrame->IsTearOff();

	int nItems = m_wndList.GetCount();
	if (nItems > 0)
	{
		m_wndList.SetRedraw (FALSE);

		for (int i = nItems - 1; i>=0; i--)
		{
			if (m_wndList.GetSel(i) > 0)
			{					   
				HWND hWnd=(HWND) m_wndList.GetItemData (i);
				HWND hMDIClient = m_pMDIFrame->m_hWndMDIClient;
				
				CBCGPMDIChildWnd* pMDIChild = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, CWnd::FromHandlePermanent(hWnd));
				if (pMDIChild != NULL)
				{
					ASSERT_VALID(pMDIChild);
					
					if (pMDIChild->m_pMDIFrame != NULL && pMDIChild->m_pMDIFrame != m_pMDIFrame)
					{
						hMDIClient = pMDIChild->m_pMDIFrame->m_hWndMDIClient;
					}
				}

				::SendMessage(hWnd,WM_CLOSE,(WPARAM) 0, (LPARAM) 0);

				if (::GetParent(hWnd) == hMDIClient)
				{
					break;
				}
			}				  
		}

		m_wndList.SetRedraw(TRUE);
	}

	FillWindowList ();
	SelActive ();
	UpdateButtons ();

	if (m_wndList.GetCount() == 0 && bIsTearOffFrame)
	{
		EndDialog(IDCANCEL);
	}
}
//**************************************************************************************
void CBCGPWindowsManagerDlg::OnCascade() 
{
	MDIMessage (WM_MDICASCADE,0);
}
//**************************************************************************************
void CBCGPWindowsManagerDlg::OnTilehorz() 
{
	MDIMessage (WM_MDITILE, MDITILE_HORIZONTAL);
}
//**************************************************************************************
void CBCGPWindowsManagerDlg::OnTilevert() 
{
	MDIMessage (WM_MDITILE, MDITILE_VERTICAL);
}
//**************************************************************************************
void CBCGPWindowsManagerDlg::OnMinimize() 
{
	int nItems = m_wndList.GetCount();
	if (nItems > 0)
	{
		m_wndList.SetRedraw(FALSE);

		for (int i = nItems - 1; i >= 0; i--)
		{
			if (m_wndList.GetSel (i) > 0)
			{
				HWND hWnd= (HWND)m_wndList.GetItemData (i);
				::ShowWindow (hWnd,SW_MINIMIZE);
			}
		}

		m_wndList.SetRedraw(TRUE);
	}

	FillWindowList ();
	SelActive ();
	UpdateButtons ();
}
//**************************************************************************************
BOOL CBCGPWindowsManagerDlg::OnInitDialog() 
{
	CBCGPDialog::OnInitDialog();
	
	if (AfxGetMainWnd () != NULL && (AfxGetMainWnd ()->GetExStyle () & WS_EX_LAYOUTRTL))
	{
		ModifyStyleEx (0, WS_EX_LAYOUTRTL);
	}

	ModifyStyle(0, WS_CLIPCHILDREN | WS_CLIPSIBLINGS);

	FillWindowList();

	CBCGPStaticLayout* pLayout = (CBCGPStaticLayout*)GetLayout ();
	if (pLayout != NULL)
	{
		pLayout->AddAnchor(IDC_BCGBARRES_LIST, CBCGPStaticLayout::e_MoveTypeNone, CBCGPStaticLayout::e_SizeTypeBoth);
		pLayout->AddAnchor(IDC_BCGBARRES_ACTIVATE, CBCGPStaticLayout::e_MoveTypeHorz, CBCGPStaticLayout::e_SizeTypeNone);
		pLayout->AddAnchor(IDOK, CBCGPStaticLayout::e_MoveTypeHorz, CBCGPStaticLayout::e_SizeTypeNone);
		pLayout->AddAnchor(IDC_BCGBARRES_SAVE, CBCGPStaticLayout::e_MoveTypeHorz, CBCGPStaticLayout::e_SizeTypeNone);
		pLayout->AddAnchor(IDC_BCGBARRES_CLOSE, CBCGPStaticLayout::e_MoveTypeHorz, CBCGPStaticLayout::e_SizeTypeNone);
		pLayout->AddAnchor(IDC_BCGBARRES_CASCADE, CBCGPStaticLayout::e_MoveTypeHorz, CBCGPStaticLayout::e_SizeTypeNone);
		pLayout->AddAnchor(IDC_BCGBARRES_TILEHORZ, CBCGPStaticLayout::e_MoveTypeHorz, CBCGPStaticLayout::e_SizeTypeNone);
		pLayout->AddAnchor(IDC_BCGBARRES_TILEVERT, CBCGPStaticLayout::e_MoveTypeHorz, CBCGPStaticLayout::e_SizeTypeNone);
		pLayout->AddAnchor(IDC_BCGBARRES_MINIMIZE, CBCGPStaticLayout::e_MoveTypeHorz, CBCGPStaticLayout::e_SizeTypeNone);
		pLayout->AddAnchor(ID_HELP, CBCGPStaticLayout::e_MoveTypeBoth, CBCGPStaticLayout::e_SizeTypeNone);
	}

	// If no MDI actions are availible, hide all MDI-related buttons:

	if (!m_bMDIActions)
	{
		GetDlgItem(IDC_BCGBARRES_TILEHORZ)->ShowWindow (SW_HIDE);
		GetDlgItem(IDC_BCGBARRES_TILEVERT)->ShowWindow (SW_HIDE);
		GetDlgItem(IDC_BCGBARRES_CASCADE)->ShowWindow (SW_HIDE);
		GetDlgItem(IDC_BCGBARRES_MINIMIZE)->ShowWindow (SW_HIDE);
	}

	SelActive();
	UpdateButtons();

	CWnd* pBtnHelp = GetDlgItem (ID_HELP);
	if (pBtnHelp != NULL)
	{
		pBtnHelp->ShowWindow (m_bHelpButton ? SW_SHOW : SW_HIDE);
	}

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
//**************************************************************************************
void CBCGPWindowsManagerDlg::OnSelchangeBcgbarresList() 
{
	UpdateButtons();
}

// Enables/Disables states of buttons
void CBCGPWindowsManagerDlg::UpdateButtons()
{						   
	int	nSel = m_wndList.GetSelCount ();

	BOOL bClose = nSel > 0;
	BOOL bSave = FALSE;

	for (int i = 0; bClose && i < m_wndList.GetCount (); i++)
	{
		if (m_wndList.GetSel (i) > 0)
		{					   
			HWND hWnd= (HWND) m_wndList.GetItemData (i);

			if (m_lstCloseDisabled.Find (hWnd))
			{
				bClose = FALSE;
			}

			if (m_lstSaveDisabled.Find (hWnd) == NULL)
			{
				bSave = TRUE;
			}
		}
	}

	GetDlgItem(IDC_BCGBARRES_CLOSE)->EnableWindow (bClose);

	GetDlgItem(IDC_BCGBARRES_SAVE)->EnableWindow (bSave);
	GetDlgItem(IDC_BCGBARRES_TILEHORZ)->EnableWindow(m_bMDIActions && nSel>=2);
	GetDlgItem(IDC_BCGBARRES_TILEVERT)->EnableWindow(m_bMDIActions && nSel>=2);
	GetDlgItem(IDC_BCGBARRES_CASCADE)->EnableWindow(m_bMDIActions && nSel>=2);
	GetDlgItem(IDC_BCGBARRES_MINIMIZE)->EnableWindow(m_bMDIActions && nSel>0);

	GetDlgItem(IDC_BCGBARRES_ACTIVATE)->EnableWindow(nSel==1);
}

// Selects currently active window in listbox
void CBCGPWindowsManagerDlg::SelActive()
{
	ASSERT_VALID(m_pMDIFrame);

	int nItems = m_wndList.GetCount();
	if (nItems > 0)
	{
		m_wndList.SetRedraw(FALSE);
		m_wndList.SelItemRange (FALSE, 0, nItems - 1);

		CList<HWND, HWND> lstSelected;
		m_pMDIFrame->GetSelectedMDIChildren(lstSelected);

		if (lstSelected.IsEmpty())
		{
			return;
		}

		for (int i = 0; i < nItems; i++)
		{
			HWND hWnd = (HWND)m_wndList.GetItemData(i);

			if (lstSelected.Find(hWnd) != NULL)
			{
				m_wndList.SetSel(i);

				if (m_wndList.GetSelCount() == (int)lstSelected.GetCount())
				{
					break;
				}
			}
		}

		m_wndList.SetRedraw (TRUE);
	}
}

void CBCGPWindowsManagerDlg::AddMDIChildren(CBCGPMDIFrameWnd* pMDIFrame, CDC& dcList, int& cxExtent)
{
	ASSERT_VALID(pMDIFrame);
	ASSERT(pMDIFrame->GetSafeHwnd() != NULL);

	HWND hwndT = ::GetWindow(pMDIFrame->m_hWndMDIClient, GW_CHILD);
	while (hwndT != NULL)
	{
		CBCGPMDIChildWnd* pFrame = DYNAMIC_DOWNCAST (CBCGPMDIChildWnd, 
			CWnd::FromHandle (hwndT));
		if (pFrame == NULL)
		{
			hwndT = ::GetWindow(hwndT,GW_HWNDNEXT);
			continue;
		}
		
		if (!pFrame->IsWindowVisible())
		{
			hwndT = ::GetWindow(hwndT,GW_HWNDNEXT);
			continue;
		}
		
		if (!pFrame->CanShowOnWindowsList ())
		{
			hwndT = ::GetWindow(hwndT,GW_HWNDNEXT);
			continue;
		}
		
		TCHAR	szWndTitle[256];
		::GetWindowText(hwndT,szWndTitle,sizeof(szWndTitle)/sizeof(szWndTitle[0]));
		
		int index = m_wndList.AddString(szWndTitle);
		
		int cxCurr = dcList.GetTextExtent (szWndTitle).cx; 
		cxExtent = max (cxExtent, cxCurr);
		
		m_wndList.SetItemData(index,(DWORD_PTR) hwndT);
		
		if (pFrame != NULL && pFrame->IsReadOnly ())
		{
			m_lstSaveDisabled.AddTail (hwndT);
		}
		
		DWORD dwStyle = ::GetWindowLong (hwndT, GWL_STYLE);
		if ((dwStyle & WS_SYSMENU) == 0)
		{
			m_bMDIActions = FALSE;
		}
		else
		{
			HMENU hSysMenu = ::GetSystemMenu (hwndT, FALSE);
			if (hSysMenu == NULL)
			{
				m_bMDIActions = FALSE;
			}
			else
			{
				MENUITEMINFO menuInfo;
				ZeroMemory(&menuInfo,sizeof(MENUITEMINFO));
				menuInfo.cbSize = sizeof(MENUITEMINFO);
				menuInfo.fMask = MIIM_STATE;
				
				if (!::GetMenuItemInfo(hSysMenu, SC_CLOSE, FALSE, &menuInfo) ||
					(menuInfo.fState & MFS_GRAYED) || 
					(menuInfo.fState & MFS_DISABLED))
				{
					m_lstCloseDisabled.AddTail (hwndT);
				}
			}
		}
		
		hwndT = ::GetWindow(hwndT,GW_HWNDNEXT);
	}
}

// Refresh windows list
void CBCGPWindowsManagerDlg::FillWindowList(void)
{
	m_wndList.SetRedraw(FALSE);
	m_wndList.ResetContent();

	int cxExtent = 0;
	
	CClientDC dcList (&m_wndList);
	CFont* pOldFont = dcList.SelectObject (GetFont ());
	ASSERT_VALID (pOldFont);
 
	m_bMDIActions = TRUE;
	m_lstCloseDisabled.RemoveAll ();
	m_lstSaveDisabled.RemoveAll ();

	AddMDIChildren(m_pMDIFrame, dcList, cxExtent);

	if (m_pMDIFrame->AreMDIChildrenTearOff() && !m_pMDIFrame->IsTearOff())
	{
		const CList<CFrameWnd*, CFrameWnd*>& lstFrames = CBCGPFrameImpl::GetFrameList();
		for (POSITION pos = lstFrames.GetHeadPosition(); pos != NULL;)
		{
			CBCGPMDIFrameWnd* pMDIFrame = DYNAMIC_DOWNCAST(CBCGPMDIFrameWnd, lstFrames.GetNext(pos));
			
			if (pMDIFrame->GetSafeHwnd() != NULL && pMDIFrame != m_pMDIFrame && pMDIFrame->IsTearOff())
			{
				AddMDIChildren(pMDIFrame, dcList, cxExtent);
			}
		}
	}

	m_wndList.SetHorizontalExtent (cxExtent + ::GetSystemMetrics (SM_CXHSCROLL) + globalUtils.ScaleByDPI(30));
	dcList.SelectObject (pOldFont);

	m_Icons.DeleteImageList();
	m_Icons.Create(globalData.m_sizeSmallIcon.cx, globalData.m_sizeSmallIcon.cy, ILC_COLOR32 | ILC_MASK, 0, 10);

	for (int i = 0; i < m_wndList.GetCount(); i++)
	{
		HWND hwnd = (HWND)m_wndList.GetItemData(i);
		HICON hIcon = NULL;

		CBCGPMDIChildWnd* pMDIChild = DYNAMIC_DOWNCAST(CBCGPMDIChildWnd, CWnd::FromHandlePermanent(hwnd));
		if (pMDIChild->GetSafeHwnd() != NULL)
		{
			hIcon = pMDIChild->GetFrameIcon();
		}
		else
		{
			if ((hIcon = (HICON)::SendMessage(hwnd, WM_GETICON, FALSE, 0)) == NULL)
			{
				hIcon = (HICON)(LONG_PTR)GetClassLongPtr(hwnd, GCLP_HICONSM);
			}
		}

		if (hIcon != NULL)
		{
			m_Icons.Add(hIcon);
			m_wndList.SetItemImage(i, m_Icons.GetImageCount() - 1);
		}
	}

	m_wndList.SetImageList(m_Icons.GetSafeHandle());

	m_wndList.SetRedraw(TRUE);
}

void CBCGPWindowsManagerDlg::OnBcgbarresWindowHelp() 
{
	CWnd* pParentFrame = AfxGetMainWnd();
	pParentFrame->SendMessage(WM_BCGWINDOW_HELP, 0, (LPARAM) this);	
}

BOOL CBCGPWindowsManagerDlg::OnHelpInfo(HELPINFO* pHelpInfo) 
{
	UNREFERENCED_PARAMETER(pHelpInfo);
	CWnd* pParentFrame = AfxGetMainWnd();
	pParentFrame->SendMessage(WM_BCGWINDOW_HELP, 0, (LPARAM) this);	

	return FALSE;
}

