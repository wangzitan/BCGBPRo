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
// BCGPDHTMLDialog.cpp : implementation file
//

#include "stdafx.h"
#include <FLOAT.H>
#include "BCGPDHTMLDialog.h"
#include "bcgprores.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#if _MSC_VER < 1300

#pragma warning(push)
#pragma warning(disable: 4061) // enumerate XXX not explicitly handled by a case label
VARTYPE AtlGetUserDefinedType(
									 ITypeInfo *pTI, 
									 HREFTYPE hrt)
{
	ASSERT(pTI != NULL);
	
	CComPtr<ITypeInfo> spTypeInfo;
	VARTYPE vt = VT_USERDEFINED;
	HRESULT hr = E_FAIL;
	hr = pTI->GetRefTypeInfo(hrt, &spTypeInfo);
	if(FAILED(hr))
		return vt;
	TYPEATTR *pta = NULL;
	
	hr = spTypeInfo->GetTypeAttr(&pta);
	if(SUCCEEDED(hr) && pta && (pta->typekind == TKIND_ALIAS || pta->typekind == TKIND_ENUM))
	{
		if (pta->tdescAlias.vt == VT_USERDEFINED)
			vt = AtlGetUserDefinedType(spTypeInfo, pta->tdescAlias.hreftype);
		else
		{
			switch (pta->typekind)
			{
			case TKIND_ENUM :
				vt = VT_I4;
				break;
			case TKIND_INTERFACE :
				vt = VT_UNKNOWN;
				break;
			case TKIND_DISPATCH :
				vt = VT_DISPATCH;
				break;
			default:
				vt = pta->tdescAlias.vt;
			}
		}				
	}
	
	if(pta)
		spTypeInfo->ReleaseTypeAttr(pta);
	return vt;
	
}

#pragma warning(pop)


HRESULT AtlGetFuncInfoFromId(
									ITypeInfo* pTypeInfo, 
									const IID& /*iid*/, 
									DISPID dispidMember, 
									LCID /*lcid*/, 
									_ATL_FUNC_INFO& info)
{
	if (pTypeInfo == NULL)
		return E_INVALIDARG;
	
	HRESULT hr = S_OK;
	FUNCDESC* pFuncDesc = NULL;
	TYPEATTR* pAttr;
	hr = pTypeInfo->GetTypeAttr(&pAttr);
	if (FAILED(hr))
		return hr;
	int i;
	for (i=0;i<pAttr->cFuncs;i++)
	{
		hr = pTypeInfo->GetFuncDesc(i, &pFuncDesc);
		if (FAILED(hr))
			return hr;
		if (pFuncDesc->memid == dispidMember)
			break;
		pTypeInfo->ReleaseFuncDesc(pFuncDesc);
		pFuncDesc = NULL;
	}
	pTypeInfo->ReleaseTypeAttr(pAttr);
	if (pFuncDesc == NULL)
		return E_FAIL;
	
	// If this assert occurs, then add a #define _ATL_MAX_VARTYPES nnnn
	// before including atlcom.h
	ATLASSERT(pFuncDesc->cParams <= _ATL_MAX_VARTYPES);
	if (pFuncDesc->cParams > _ATL_MAX_VARTYPES)
		return E_FAIL;
	
	for (i = 0; i < pFuncDesc->cParams; i++)
	{
		info.pVarTypes[i] = pFuncDesc->lprgelemdescParam[i].tdesc.vt;
		if (info.pVarTypes[i] == VT_PTR)
			info.pVarTypes[i] = (VARTYPE)(pFuncDesc->lprgelemdescParam[i].tdesc.lptdesc->vt | VT_BYREF);
		if (info.pVarTypes[i] == VT_SAFEARRAY)
			info.pVarTypes[i] = (VARTYPE)(pFuncDesc->lprgelemdescParam[i].tdesc.lptdesc->vt | VT_ARRAY);
		if (info.pVarTypes[i] == VT_USERDEFINED)
			info.pVarTypes[i] = AtlGetUserDefinedType(pTypeInfo, pFuncDesc->lprgelemdescParam[i].tdesc.hreftype);
	}
	
	VARTYPE vtReturn = pFuncDesc->elemdescFunc.tdesc.vt;
	switch(vtReturn)
	{
	case VT_INT:
		vtReturn = VT_I4;
		break;
	case VT_UINT:
		vtReturn = VT_UI4;
		break;
	case VT_VOID:
		vtReturn = VT_EMPTY; // this is how DispCallFunc() represents void
		break;
	case VT_HRESULT:
		vtReturn = VT_ERROR;
		break;
	}
	info.vtReturn = vtReturn;
	info.cc = pFuncDesc->callconv;
	info.nParams = pFuncDesc->cParams;
	pTypeInfo->ReleaseFuncDesc(pFuncDesc);
	return S_OK;
}

#define ENSURE_ARG(cond)

#define ATL_VARIANT_TRUE VARIANT_BOOL( -1 )
#define ATL_VARIANT_FALSE VARIANT_BOOL( 0 )

/////////////////////////////////////////////////////////////////////////////
// CDHtmlEventSink

HRESULT CDHtmlEventSink::ConnectToConnectionPoint(IUnknown *punkObj, REFIID riid, DWORD *pdwCookie)
{
	return AtlAdvise(punkObj, (IDispatch *) this, riid, pdwCookie);
}

void CDHtmlEventSink::DisconnectFromConnectionPoint(IUnknown *punkObj, REFIID riid, DWORD& dwCookie)
{
	AtlUnadvise(punkObj, riid, dwCookie);
}

STDMETHODIMP CDHtmlEventSink::CDHtmlSinkHandlerQueryInterface(REFIID iid, LPVOID* ppvObj)
{
	if (!ppvObj)
		return E_POINTER;

	*ppvObj = NULL;
	if (IsEqualIID(iid, __uuidof(IDispatch)) || IsEqualIID(iid, __uuidof(IUnknown)))
	{
		*ppvObj = (IDispatch *) this;
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CDHtmlEventSink::AddRef()
{
	return 1;
}

STDMETHODIMP_(ULONG) CDHtmlEventSink::Release()
{
	return 1;
}

STDMETHODIMP CDHtmlEventSink::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	ATLTRACENOTIMPL(_T("CDHtmlEventSink::GetTypeInfoCount"));
}

STDMETHODIMP CDHtmlEventSink::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	iTInfo; // unused
	lcid; // unused

	*ppTInfo = NULL;
	ATLTRACENOTIMPL(_T("CDHtmlEventSink::GetTypeInfo"));
}

STDMETHODIMP CDHtmlEventSink::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	riid; // unused
	rgszNames; // unused
	cNames; // unused
	lcid; // unused
	rgDispId; // unused

	ATLTRACENOTIMPL(_T("CDHtmlEventSink::GetIDsOfNames"));
}

STDMETHODIMP CDHtmlEventSink::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
	DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	riid; // unused
	lcid; // unused
	wFlags; // unused

	IHTMLElement *psrcElement;
	HRESULT hr = S_OK;

	VariantInit(pVarResult);

	if (DHtmlEventHook(&hr, dispIdMember, pDispParams, pVarResult, pExcepInfo, puArgErr))
		return hr;

	const DHtmlEventMapEntry *pMap = GetDHtmlEventMap();

	int nIndex = FindDHtmlEventEntry(pMap, dispIdMember, &psrcElement);
	if (nIndex<0)
		return DISP_E_MEMBERNOTFOUND;

	// now call it
	if (pMap)
	{
		hr = (this->*((DHEVTFUNC) (GetDHtmlEventMap()[nIndex].pfnEventFunc)))(psrcElement);
		if (GetDHtmlEventMap()[nIndex].nType != DHTMLEVENTMAPENTRY_CONTROL && pVarResult)
		{
			pVarResult->vt = VT_BOOL;
			pVarResult->boolVal = (hr==S_OK) ? ATL_VARIANT_TRUE : ATL_VARIANT_FALSE;
		}
	}
	if (psrcElement)
		psrcElement->Release();

	return hr;
}

BOOL CDHtmlEventSink::DHtmlEventHook(HRESULT *phr, DISPID dispIdMember, DISPPARAMS *pDispParams,
	VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	phr; // unused
	dispIdMember; // unused
	pDispParams; // unused
	pVarResult; // unused
	pExcepInfo; // unused
	puArgErr; // unused

	// stub base implementation
	return FALSE;
}

int CDHtmlEventSink::FindDHtmlEventEntry(const DHtmlEventMapEntry *pEventMap, DISPID dispIdMember,
						   IHTMLElement **ppsrcElement)
{
	HRESULT hr = DISP_E_MEMBERNOTFOUND;
	CComPtr<IHTMLWindow2> sphtmlWnd;
	CComPtr<IHTMLEventObj> sphtmlEvent;
	CComPtr<IHTMLElement> sphtmlElement;
	CComPtr<IHTMLDocument2> sphtmlDoc;
	CComBSTR bstrName;
	CComBSTR bstrClass;
	CComBSTR bstrTagName;

	int i;
	int nIndexFound = -1;

	if(ppsrcElement == NULL)
		return E_POINTER;

	*ppsrcElement = NULL;

	if (!pEventMap)
		goto Error;

	// get the html document
	hr = GetDHtmlDocument(&sphtmlDoc);
	if (sphtmlDoc == NULL)
		goto Error;

	// get the element that generated the event
	sphtmlDoc->get_parentWindow(&sphtmlWnd);
	if ((sphtmlWnd==NULL) || FAILED(sphtmlWnd->get_event(&sphtmlEvent)) || (sphtmlEvent==NULL))
	{
		hr = DISP_E_MEMBERNOTFOUND;
		goto Error;
	}
	sphtmlEvent->get_srcElement(&sphtmlElement);
	*ppsrcElement = sphtmlElement;
	if (sphtmlElement)
		sphtmlElement.p->AddRef();
	// look for the dispid in the map
	for (i=0; pEventMap[i].nType != DHTMLEVENTMAPENTRY_END; i++)
	{
		if (pEventMap[i].dispId == dispIdMember)
		{
			if (pEventMap[i].nType == DHTMLEVENTMAPENTRY_NAME)
			{
				if (!bstrName && sphtmlElement)
					sphtmlElement->get_id(&bstrName);
				if (bstrName && pEventMap[i].szName && !wcscmp(bstrName, CComBSTR(pEventMap[i].szName)) ||
					(!bstrName && !sphtmlElement))
				{
					nIndexFound = i;
					break;
				}
			}
			else if (pEventMap[i].nType == DHTMLEVENTMAPENTRY_CLASS)
			{
				if (!bstrClass && sphtmlElement)
					sphtmlElement->get_className(&bstrClass);
				if (bstrClass && !wcscmp(bstrClass, CComBSTR(pEventMap[i].szName)))
				{
					nIndexFound = i;
					break;
				}
			}
			else if (pEventMap[i].nType == DHTMLEVENTMAPENTRY_TAG)
			{
				if (!bstrTagName && sphtmlElement)
					sphtmlElement->get_tagName(&bstrTagName);
				if (bstrTagName && !_wcsicmp(bstrTagName, CComBSTR(pEventMap[i].szName)))
				{
					nIndexFound = i;
					break;
				}
			}
		}
	}
Error:
	if (nIndexFound==-1 && *ppsrcElement)
	{
		(*ppsrcElement)->Release();
		*ppsrcElement = NULL;
	}
	return nIndexFound;
}

/////////////////////////////////////////////////////////////////////////////
// CDHtmlControlSink

CDHtmlControlSink::CDHtmlControlSink()
{
	m_dwCookie = 0;
	m_pHandler = NULL;
	m_dwThunkOffset = 0;
	memset(&m_iid, 0x00, sizeof(IID));
}

CDHtmlControlSink::CDHtmlControlSink(IUnknown *punkObj, CDHtmlSinkHandler *pHandler,
		LPCTSTR szControlId, DWORD_PTR dwThunkOffset /*= 0*/)
{
	m_dwCookie = 0;
	m_pHandler = pHandler;
	m_szControlId = szControlId;
	m_dwThunkOffset = dwThunkOffset;
	ConnectToControl(punkObj);
}

CDHtmlControlSink::~CDHtmlControlSink()
{
	if (m_dwCookie != 0)
		AtlUnadvise(m_spunkObj, m_iid, m_dwCookie);
}

HRESULT CDHtmlControlSink::ConnectToControl(IUnknown *punkObj)
{
	m_spunkObj = punkObj;
	HRESULT hr = AtlGetObjectSourceInterface(punkObj, &m_libid, &m_iid, &m_wMajor, &m_wMinor);
	if (FAILED(hr))
		return hr;

	CComPtr<ITypeLib> spTypeLib;

	hr = LoadRegTypeLib(m_libid, m_wMajor, m_wMinor, LOCALE_USER_DEFAULT, &spTypeLib);
	if (FAILED(hr))
		return hr;

	hr = spTypeLib->GetTypeInfoOfGuid(m_iid, &m_spTypeInfo);
	if (FAILED(hr))
		return hr;

	return AtlAdvise(punkObj, this, m_iid, &m_dwCookie);
}

STDMETHODIMP_(ULONG) CDHtmlControlSink::AddRef()
{
	return 1;
}

STDMETHODIMP_(ULONG) CDHtmlControlSink::Release()
{
	return 1;
}

STDMETHODIMP CDHtmlControlSink::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
	if (!ppvObj)
		return E_POINTER;

	*ppvObj = NULL;
	if (IsEqualIID(iid, __uuidof(IUnknown)) || 
		IsEqualIID(iid, __uuidof(IDispatch)) || 
		IsEqualIID(iid, m_iid))
	{
		*ppvObj = this;
		return S_OK;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP CDHtmlControlSink::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	ATLTRACENOTIMPL(_T("CDHtmlControlSink::GetTypeInfoCount"));
}

STDMETHODIMP CDHtmlControlSink::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	iTInfo; // unused
	lcid; // unused

	*ppTInfo = NULL;
	ATLTRACENOTIMPL(_T("CDHtmlControlSink::GetTypeInfo"));
}

STDMETHODIMP CDHtmlControlSink::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	riid; // unused
	rgszNames; // unused
	cNames; // unused
	lcid; // unused
	rgDispId; // unused

	ATLTRACENOTIMPL(_T("CDHtmlControlSink::GetIDsOfNames"));
}

STDMETHODIMP CDHtmlControlSink::Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags,
	DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	riid; // unused
	wFlags; // unused
	pExcepInfo; // unused
	puArgErr; // unused

	_ATL_FUNC_INFO info;

	BOOL fFound = FALSE;
	DHEVTFUNCCONTROL pEvent = NULL;

	const DHtmlEventMapEntry *pEventMap = m_pHandler->GetDHtmlEventMap();
	for (int i=0; pEventMap[i].nType != DHTMLEVENTMAPENTRY_END; i++)
	{
		if (pEventMap[i].nType==DHTMLEVENTMAPENTRY_CONTROL &&
			pEventMap[i].dispId == dispidMember &&
			!_tcscmp(pEventMap[i].szName, m_szControlId))
		{
			// found the entry
			pEvent = pEventMap[i].pfnEventFunc;
			fFound = TRUE;
			break;
		}
	}
	if (!fFound)
		return DISP_E_MEMBERNOTFOUND;

	HRESULT hr = GetFuncInfoFromId(m_iid, dispidMember, lcid, info);
	if (FAILED(hr))
	{
		return S_OK;
	}
	return InvokeFromFuncInfo(pEvent, info, pdispparams, pvarResult);
}

//Helper for invoking the event
HRESULT CDHtmlControlSink::InvokeFromFuncInfo(DHEVTFUNCCONTROL pEvent, _ATL_FUNC_INFO& info, DISPPARAMS* pdispparams, VARIANT* pvarResult)
{
#if _MSC_VER < 1300
	UNREFERENCED_PARAMETER(pEvent);
	UNREFERENCED_PARAMETER(info);
	UNREFERENCED_PARAMETER(pdispparams);
	UNREFERENCED_PARAMETER(pvarResult);

	return E_INVALIDARG;
#else
	USES_ATL_SAFE_ALLOCA;
	if (info.nParams < 0)
	{
		return E_INVALIDARG;
	}
	if (info.nParams > size_t(-1) / sizeof(VARIANTARG*))
	{
		return E_OUTOFMEMORY;
	}
	VARIANTARG** pVarArgs = info.nParams ? (VARIANTARG**)_ATL_SAFE_ALLOCA(sizeof(VARIANTARG*)*info.nParams,_ATL_SAFE_ALLOCA_DEF_THRESHOLD) : 0;
	if(!pVarArgs)
	{
		return E_OUTOFMEMORY;
	}

	for (int i=0; i<info.nParams; i++)
	{
		pVarArgs[i] = &pdispparams->rgvarg[info.nParams - i - 1];
	}

	CComStdCallThunk<CDHtmlSinkHandler> thunk;
	if (m_pHandler)
		thunk.Init(pEvent, reinterpret_cast< CDHtmlSinkHandler* >((DWORD_PTR) m_pHandler - m_dwThunkOffset));

	CComVariant tmpResult;
	if (pvarResult == NULL)
		pvarResult = &tmpResult;

	HRESULT hr = DispCallFunc(
		&thunk,
		0,
		info.cc,
		info.vtReturn,
		info.nParams,
		info.pVarTypes,
		pVarArgs,
		pvarResult);
	ATLASSERT(SUCCEEDED(hr));
	return hr;
#endif
}

HRESULT CDHtmlControlSink::GetFuncInfoFromId(const IID& iid, DISPID dispidMember, LCID lcid, _ATL_FUNC_INFO& info)
{
	if (!m_spTypeInfo)
		return E_FAIL;
	return AtlGetFuncInfoFromId(m_spTypeInfo, iid, dispidMember, lcid, info);
}

/*
VARTYPE CDHtmlControlSink::GetUserDefinedType(ITypeInfo *pTI, HREFTYPE hrt)
{
	return AtlGetUserDefinedType(pTI, hrt);
}
*/

/////////////////////////////////////////////////////////////////////////////
// CDHtmlElementEventSink

CDHtmlElementEventSink::CDHtmlElementEventSink(CDHtmlEventSink *pHandler, IDispatch *pdisp)
{
	m_pHandler = pHandler;
	pdisp->QueryInterface(__uuidof(IUnknown), (void **) &m_spunkElem);
	m_dwCookie = 0;
}

STDMETHODIMP_(ULONG) CDHtmlElementEventSink::AddRef()
{
	return 1;
}

STDMETHODIMP_(ULONG) CDHtmlElementEventSink::Release()
{
	return 1;
}

STDMETHODIMP CDHtmlElementEventSink::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
	if (!ppvObj)
		return E_POINTER;

	*ppvObj = NULL;
	if (IsEqualIID(iid, __uuidof(IUnknown)) || 
		IsEqualIID(iid, __uuidof(IDispatch)))
	{
		*ppvObj = this;
		return S_OK;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP CDHtmlElementEventSink::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	ATLTRACENOTIMPL(_T("CDHtmlElementEventSink::GetTypeInfoCount"));
}

STDMETHODIMP CDHtmlElementEventSink::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	iTInfo; // unused
	lcid; // unused

	*ppTInfo = NULL;
	ATLTRACENOTIMPL(_T("CDHtmlElementEventSink::GetTypeInfo"));
}

STDMETHODIMP CDHtmlElementEventSink::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	riid; // unused
	rgszNames; // unused
	cNames; // unused
	lcid; // unused
	rgDispId; // unused

	ATLTRACENOTIMPL(_T("CDHtmlElementEventSink::GetIDsOfNames"));
}

STDMETHODIMP CDHtmlElementEventSink::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
	DISPPARAMS *pdispparams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	riid; // unused
	lcid; // unused
	wFlags; // unused
	pdispparams; // unused
	pExcepInfo; // unused
	puArgErr; // unused

	HRESULT hr;
	CComPtr<IHTMLWindow2> sphtmlWnd;
	CComPtr<IHTMLDocument2> sphtmlDoc;
	CComPtr<IHTMLElement> sphtmlElem;
	CComPtr<IHTMLElement> spsrcElem;
	CComPtr<IHTMLEventObj> sphtmlEvent;

	CComBSTR bstrId;

	if (pVarResult)
		VariantInit(pVarResult);
	hr = m_spunkElem->QueryInterface(&sphtmlElem);
	if (!sphtmlElem)
		return hr;

	hr = sphtmlElem->get_id(&bstrId);
	if (FAILED(hr))
		return hr;

	hr = m_pHandler->GetDHtmlDocument(&sphtmlDoc);
	if (FAILED(hr))
		return hr;

	hr = sphtmlDoc->get_parentWindow(&sphtmlWnd);
	if (FAILED(hr))
		return hr;

	hr = sphtmlWnd->get_event(&sphtmlEvent);
	if (FAILED(hr))
		return hr;

	hr = sphtmlEvent->get_srcElement(&spsrcElem);
	if (FAILED(hr))
		return hr;

	const DHtmlEventMapEntry *pEventMap = m_pHandler->GetDHtmlEventMap();
	for (int i=0; pEventMap[i].nType != DHTMLEVENTMAPENTRY_END; i++)
	{
		if (pEventMap[i].nType==DHTMLEVENTMAPENTRY_ELEMENT &&
			pEventMap[i].dispId == dispIdMember)
		{
			if (pEventMap[i].szName && !wcscmp(CComBSTR(pEventMap[i].szName), bstrId))
			{
				// found the entry
				hr = (m_pHandler->*((DHEVTFUNC) (m_pHandler->GetDHtmlEventMap()[i].pfnEventFunc)))(spsrcElem);
				if (pVarResult)
				{
					pVarResult->vt = VT_BOOL;
					pVarResult->boolVal = (hr==S_OK) ? ATL_VARIANT_TRUE : ATL_VARIANT_FALSE;
				}
				return S_OK;
			}
		}
	}

	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CDHtmlElementEventSink::Advise(LPUNKNOWN pUnkObj, REFIID iid)
{
	return AtlAdvise((LPUNKNOWN)pUnkObj, (LPDISPATCH)this, iid, &m_dwCookie);

}

HRESULT CDHtmlElementEventSink::UnAdvise(LPUNKNOWN pUnkObj, REFIID iid)
{
	iid; // unused

	return AtlUnadvise((LPUNKNOWN)pUnkObj, __uuidof(HTMLElementEvents), m_dwCookie);
}

#endif

/////////////////////////////////////////////////////////////////////////////
// CBCGPBrowserControlSite

CBCGPBrowserControlSite::CBCGPBrowserControlSite(COleControlContainer* pCtrlCont, CBCGPDHtmlDialog *pHandler) :
	COleControlSite(pCtrlCont)
{
	ASSERT(pHandler);
	m_pHandler = pHandler;
}

LPUNKNOWN CBCGPBrowserControlSite::GetInterfaceHook(const void *iid)
{
	if (IsEqualIID((REFIID) (*(IID*)iid), IID_IDocHostUIHandler))
		return (IDocHostUIHandler *) this;
	return NULL;
}

// IUnknown methods
STDMETHODIMP CBCGPBrowserControlSite::QueryInterface(REFIID riid, void **ppvObject)
{
	return COleControlSite::ExternalQueryInterface(&riid, ppvObject);
}
STDMETHODIMP_(ULONG) CBCGPBrowserControlSite::AddRef()
{
	return ExternalAddRef();
}

STDMETHODIMP_(ULONG) CBCGPBrowserControlSite::Release()
{
	return ExternalRelease();
}

// IDocHostUIHandler methods
STDMETHODIMP CBCGPBrowserControlSite::ShowContextMenu(DWORD dwID, POINT *ppt, IUnknown *pcmdtReserved, IDispatch *pdispReserved)
{
	return m_pHandler->ShowContextMenu(dwID, ppt, pcmdtReserved, pdispReserved);
}
STDMETHODIMP CBCGPBrowserControlSite::GetHostInfo(DOCHOSTUIINFO *pInfo)
{
	return m_pHandler->GetHostInfo(pInfo);
}
STDMETHODIMP CBCGPBrowserControlSite::ShowUI(DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc)
{
	return m_pHandler->ShowUI(dwID, pActiveObject, pCommandTarget, pFrame, pDoc);
}
STDMETHODIMP CBCGPBrowserControlSite::HideUI(void)
{
	return m_pHandler->HideUI();
}
STDMETHODIMP CBCGPBrowserControlSite::UpdateUI(void)
{
	return m_pHandler->UpdateUI();
}
STDMETHODIMP CBCGPBrowserControlSite::EnableModeless(BOOL fEnable)
{
	return m_pHandler->EnableModeless(fEnable);
}
STDMETHODIMP CBCGPBrowserControlSite::OnDocWindowActivate(BOOL fActivate)
{
	return m_pHandler->OnDocWindowActivate(fActivate);
}
STDMETHODIMP CBCGPBrowserControlSite::OnFrameWindowActivate(BOOL fActivate)
{
	return m_pHandler->OnFrameWindowActivate(fActivate);
}
STDMETHODIMP CBCGPBrowserControlSite::ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fRameWindow)
{
	return m_pHandler->ResizeBorder(prcBorder, pUIWindow, fRameWindow);
}
STDMETHODIMP CBCGPBrowserControlSite::TranslateAccelerator(LPMSG lpMsg, const GUID *pguidCmdGroup, DWORD nCmdID)
{
	return m_pHandler->TranslateAccelerator(lpMsg, pguidCmdGroup, nCmdID);
}
STDMETHODIMP CBCGPBrowserControlSite::GetOptionKeyPath(LPOLESTR *pchKey, DWORD dw)
{
	return m_pHandler->GetOptionKeyPath(pchKey, dw);
}
STDMETHODIMP CBCGPBrowserControlSite::GetDropTarget(IDropTarget *pDropTarget, IDropTarget **ppDropTarget)
{
	return m_pHandler->GetDropTarget(pDropTarget, ppDropTarget);
}
STDMETHODIMP CBCGPBrowserControlSite::GetExternal(IDispatch **ppDispatch)
{
	return m_pHandler->GetExternal(ppDispatch);
}
STDMETHODIMP CBCGPBrowserControlSite::TranslateUrl(DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut)
{
	return m_pHandler->TranslateUrl(dwTranslate, pchURLIn, ppchURLOut);
}
STDMETHODIMP CBCGPBrowserControlSite::FilterDataObject(IDataObject *pDO, IDataObject **ppDORet)
{
	return m_pHandler->FilterDataObject(pDO, ppDORet);
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPDHtmlDialog dialog

BEGIN_MESSAGE_MAP(CBCGPDHtmlDialog, CBCGPDialog)
	ON_WM_DESTROY()
	ON_WM_SIZE()
	ON_REGISTERED_MESSAGE(CBCGPDHtmlDialog::WM_DESTROYMODELESS, OnDestroyModeless)
END_MESSAGE_MAP()

BEGIN_EVENTSINK_MAP(CBCGPDHtmlDialog, CBCGPDialog)
	ON_EVENT(CBCGPDHtmlDialog, IDC_BCGBARRES_BROWSER, 252 /* NavigateComplete2 */, _OnNavigateComplete2, VTS_DISPATCH VTS_PVARIANT)
	ON_EVENT(CBCGPDHtmlDialog, IDC_BCGBARRES_BROWSER, 250 /* BeforeNavigate2 */, _OnBeforeNavigate2, VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PBOOL)
	ON_EVENT(CBCGPDHtmlDialog, IDC_BCGBARRES_BROWSER, 259 /* DocumentComplete */, _OnDocumentComplete, VTS_DISPATCH VTS_PVARIANT)
END_EVENTSINK_MAP()

UINT CBCGPDHtmlDialog::WM_DESTROYMODELESS = 0;

IMPLEMENT_DYNAMIC(CBCGPDHtmlDialog, CBCGPDialog)

CBCGPDHtmlDialog::CBCGPDHtmlDialog() :
	CBCGPDialog()
{
	Initialize();
}

CBCGPDHtmlDialog::CBCGPDHtmlDialog(UINT nIDTemplate, UINT nHtmlResID /*= 0*/, CWnd *pParentWnd /*= NULL*/, UINT nIDBrowserPlaceHolder/* = 0*/) : 
	CBCGPDialog(nIDTemplate, pParentWnd)
{	
	Initialize();
	m_nHtmlResID = nHtmlResID;
	m_nIDBrowserPlaceHolder = nIDBrowserPlaceHolder;
}

CBCGPDHtmlDialog::CBCGPDHtmlDialog(LPCTSTR lpszTemplateName, LPCTSTR szHtmlResID /*= NULL*/, CWnd *pParentWnd /*= NULL*/, UINT nIDBrowserPlaceHolder/* = 0*/) :
	CBCGPDialog(lpszTemplateName, pParentWnd)
{
	Initialize();
	m_szHtmlResID = const_cast<LPTSTR>(szHtmlResID);
	m_nIDBrowserPlaceHolder = nIDBrowserPlaceHolder;
}

CBCGPDHtmlDialog::~CBCGPDHtmlDialog()
{
}

BOOL CBCGPDHtmlDialog::CreateControlSite(COleControlContainer* pContainer, 
		COleControlSite** ppSite, UINT /* nID */, REFCLSID /* clsid */)
{
	if(ppSite == NULL)
	{
		ASSERT(FALSE);
		return FALSE;
	}
		
	CBCGPBrowserControlSite *pBrowserSite = 
			new CBCGPBrowserControlSite(pContainer, this);
	if (!pBrowserSite)
		return FALSE;
		
	*ppSite = pBrowserSite;
	return TRUE;
}

HRESULT CBCGPDHtmlDialog::GetDHtmlDocument(IHTMLDocument2 **pphtmlDoc)
{
	if(pphtmlDoc == NULL)
	{
		ASSERT(FALSE);
		return E_POINTER;
	}
	
	*pphtmlDoc = NULL;

	if (m_spHtmlDoc)
	{
		*pphtmlDoc = m_spHtmlDoc;
		(*pphtmlDoc)->AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
}

void CBCGPDHtmlDialog::GetCurrentUrl(CString& szUrl)
{
	szUrl = m_strCurrentUrl;
}

BOOL CBCGPDHtmlDialog::CanAccessExternal()
{
	// if the dispatch we have is safe, 
	// we allow access
	if (IsExternalDispatchSafe())
		return TRUE;

	// the external dispatch is not safe, so we check
	// whether the current zone allows for scripting
	// of objects that are not safe for scripting
	if (m_spHtmlDoc == NULL)
		return FALSE;

	ATL::CComPtr<IInternetHostSecurityManager> spSecMan;
	m_spHtmlDoc->QueryInterface(IID_IInternetHostSecurityManager,
			(void **) &spSecMan);
	if (spSecMan == NULL)
		return FALSE;

	HRESULT hr = spSecMan->ProcessUrlAction(URLACTION_ACTIVEX_OVERRIDE_OBJECT_SAFETY,
		NULL, 0, NULL, 0, 0, PUAF_DEFAULT);
	if (hr == S_OK)
		return TRUE;
	return FALSE;
}

void CBCGPDHtmlDialog::OnBeforeNavigate(LPDISPATCH pDisp, LPCTSTR szUrl)
{
	szUrl; // unused

	if (pDisp != m_pBrowserApp)
		return;
	DisconnectDHtmlEvents();
	m_spHtmlDoc = NULL;
	m_strCurrentUrl.Empty();
}

void CBCGPDHtmlDialog::OnNavigateComplete(LPDISPATCH pDisp, LPCTSTR szUrl)
{
	if (pDisp != m_pBrowserApp)
		return;
	IDispatch *pdispDoc = NULL;
	m_pBrowserApp->get_Document(&pdispDoc);
	if (!pdispDoc)
		return;

	ASSERT(m_spHtmlDoc==NULL);

	pdispDoc->QueryInterface(IID_IHTMLDocument2, (void **) &m_spHtmlDoc);

	if (m_bUseHtmlTitle)
	{
		CComBSTR bstrTitle;
		m_spHtmlDoc->get_title(&bstrTitle);
		CString str = CString(bstrTitle);
		SetWindowText(str);
	}

	m_strCurrentUrl = szUrl;

	ConnectDHtmlEvents(pdispDoc);
	pdispDoc->Release();
}

void CBCGPDHtmlDialog::OnDocumentComplete(LPDISPATCH pDisp, LPCTSTR szUrl)
{
	szUrl; // unused

	if (pDisp != m_pBrowserApp)
		return;
	ConnectDHtmlElementEvents((((DWORD_PTR)static_cast< CDHtmlSinkHandler* >(this)) - (DWORD_PTR)this));
	UpdateData(FALSE);
}

BOOL CBCGPDHtmlDialog::LoadFromResource(LPCTSTR lpszResource)
{
	HINSTANCE hInstance = AfxGetResourceHandle();
	ASSERT(hInstance != NULL);

	CString strResourceURL;
	BOOL bRetVal = TRUE;
	LPTSTR lpszModule = new TCHAR[_MAX_PATH];

	int ret = GetModuleFileName(hInstance, lpszModule, _MAX_PATH);
	
	if (ret == 0 || ret == _MAX_PATH)
		bRetVal = FALSE;
	else
	{
		strResourceURL.Format(_T("res://%s/%s"), lpszModule, lpszResource);
		Navigate(strResourceURL, 0, 0, 0);
	}

	delete [] lpszModule;
	return bRetVal;
}

BOOL CBCGPDHtmlDialog::LoadFromResource(UINT nRes)
{
	HINSTANCE hInstance = AfxGetResourceHandle();
	ASSERT(hInstance != NULL);

	CString strResourceURL;
	BOOL bRetVal = TRUE;
	LPTSTR lpszModule = new TCHAR[_MAX_PATH];

	if (GetModuleFileName(hInstance, lpszModule, _MAX_PATH))
	{
		strResourceURL.Format(_T("res://%s/%d"), lpszModule, nRes);
		Navigate(strResourceURL, 0, 0, 0);
	}
	else
		bRetVal = FALSE;

	delete [] lpszModule;
	return bRetVal;
}

void CBCGPDHtmlDialog::Navigate(LPCTSTR lpszURL, DWORD dwFlags /*= 0*/, 
	LPCTSTR lpszTargetFrameName /*= NULL*/, LPCTSTR lpszHeaders /*= NULL*/,
	LPVOID lpvPostData /*= NULL*/, DWORD dwPostDataLen /*= 0*/)
{
	CComBSTR bstrURL = lpszURL;

	COleSafeArray vPostData;
	if (lpvPostData != NULL)
	{
		if (dwPostDataLen == 0)
			dwPostDataLen = static_cast<DWORD>(_tcslen((LPCTSTR) lpvPostData));

		vPostData.CreateOneDim(VT_UI1, dwPostDataLen, lpvPostData);
	}

	m_pBrowserApp->Navigate(bstrURL,
		COleVariant((long) dwFlags, VT_I4),
		COleVariant(lpszTargetFrameName, VT_BSTR),
		vPostData,
		COleVariant(lpszHeaders, VT_BSTR));
}

void CBCGPDHtmlDialog::DestroyModeless()
{
	if (!CBCGPDHtmlDialog::WM_DESTROYMODELESS)
		CBCGPDHtmlDialog::WM_DESTROYMODELESS = RegisterWindowMessage(_T("DHtmlDialogDestroy"));
	PostMessage(CBCGPDHtmlDialog::WM_DESTROYMODELESS, 0, 0);
}

void CBCGPDHtmlDialog::OnDestroy()
{
	DisconnectDHtmlEvents();
	m_spHtmlDoc = NULL;

	// now tell the browser control we're shutting down
	if (m_pBrowserApp)
	{
		ATL::CComPtr<IOleObject> spObject;
		m_pBrowserApp->QueryInterface(IID_IOleObject, (void **) &spObject);
		if (spObject != NULL)
		{
			spObject->Close(OLECLOSE_NOSAVE);
			spObject.Release();
		}
		m_pBrowserApp = NULL;
	}
	CBCGPDialog::OnDestroy();
}

void CBCGPDHtmlDialog::OnSize(UINT nType, int cx, int cy)
{
	CBCGPDialog::OnSize(nType, cx, cy);
	if (!m_bAttachedControl && m_nIDBrowserPlaceHolder == 0 && m_wndBrowser.m_hWnd)
		m_wndBrowser.MoveWindow(0, 0, cx, cy);	
}

BOOL CBCGPDHtmlDialog::OnInitDialog()
{
	AfxEnableControlContainer();

	CBCGPDialog::OnInitDialog();

	// if we've been created from the dynamic template
	// set the caption
	if (!m_lpszTemplateName)
		SetWindowText(m_strDlgCaption);

	// check if there is a browser control on the dialog
	// already
	CWnd *pCtrl = GetDlgItem(IDC_BCGBARRES_BROWSER);
	LPUNKNOWN lpUnk;
	if (pCtrl)
	{
		lpUnk = pCtrl->GetControlUnknown();
		if (lpUnk && SUCCEEDED(lpUnk->QueryInterface(IID_IWebBrowser2, (void **) &m_pBrowserApp)))
		{
			m_wndBrowser.Attach(pCtrl->m_hWnd);
			m_bAttachedControl = TRUE;
		}
	}
	if (m_pBrowserApp == NULL)
	{
		CRect rectBrowser;
		GetClientRect(rectBrowser);

		if (m_nIDBrowserPlaceHolder != 0)
		{
			CWnd* pWndPlaceHolder = GetDlgItem(m_nIDBrowserPlaceHolder);
			if (pWndPlaceHolder->GetSafeHwnd() != NULL)
			{
				pWndPlaceHolder->GetWindowRect(rectBrowser);
				ScreenToClient(rectBrowser);
			}
		}

		// create the control window
		m_wndBrowser.CreateControl(CLSID_WebBrowser, NULL,
					WS_VISIBLE | WS_CHILD, rectBrowser, this, IDC_BCGBARRES_BROWSER);
		lpUnk = m_wndBrowser.GetControlUnknown();
		if (FAILED(lpUnk->QueryInterface(IID_IWebBrowser2, (void**) &m_pBrowserApp)))
		{
			m_wndBrowser.DestroyWindow();
			DestroyWindow();
			return TRUE;
		}
	}

	if (m_nHtmlResID)
		LoadFromResource(m_nHtmlResID);
	else if (m_szHtmlResID)
		LoadFromResource(m_szHtmlResID);
	else if (m_strCurrentUrl)
		Navigate(m_strCurrentUrl);
	return TRUE;
}

BEGIN_DHTML_EVENT_MAP(CBCGPDHtmlDialog)
	DHTML_EVENT(DISPID_EVMETH_ONREADYSTATECHANGE, NULL, OnDocumentReadyStateChange)
END_DHTML_EVENT_MAP()

void CBCGPDHtmlDialog::Initialize()
{
	SetHostFlags(DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_SCROLL_NO);
	m_bUseHtmlTitle = FALSE;
	m_bAttachedControl = FALSE;
	m_pBrowserApp = NULL;
	m_dwDHtmlEventSinkCookie = 0;
	m_szHtmlResID = NULL;
	m_nHtmlResID = 0;
	m_nIDBrowserPlaceHolder = 0;
}

void CBCGPDHtmlDialog::_OnBeforeNavigate2(LPDISPATCH pDisp, VARIANT FAR* URL, VARIANT FAR* Flags, VARIANT FAR* TargetFrameName, VARIANT FAR* PostData, VARIANT FAR* Headers, BOOL FAR* Cancel)
{
	Flags; // unused
	TargetFrameName; // unused
	PostData; // unused
	Headers; // unused
	Cancel; // unused

	CString str(V_BSTR(URL));
	OnBeforeNavigate(pDisp, str);
}

void CBCGPDHtmlDialog::_OnNavigateComplete2(LPDISPATCH pDisp, VARIANT FAR* URL)
{
	ASSERT(V_VT(URL) == VT_BSTR);

	CString str(V_BSTR(URL));

	OnNavigateComplete(pDisp, str);
}

void CBCGPDHtmlDialog::_OnDocumentComplete(LPDISPATCH pDisp, VARIANT* URL)
{
	ASSERT(V_VT(URL) == VT_BSTR);

	CString str(V_BSTR(URL));
	OnDocumentComplete(pDisp, str);
}

HRESULT CBCGPDHtmlDialog::OnDocumentReadyStateChange(IHTMLElement *phtmlElem)
{
	phtmlElem; // unused

	ATL::CComPtr<IHTMLDocument2> sphtmlDoc;
	GetDHtmlDocument(&sphtmlDoc);
	if (sphtmlDoc)
	{
		CComBSTR bstrState;
		sphtmlDoc->get_readyState(&bstrState);
		if (bstrState)
		{
			if (bstrState==TEXT("complete"))
				ConnectDHtmlElementEvents((((DWORD_PTR)static_cast< CDHtmlSinkHandler* >(this)) - (DWORD_PTR) this));
			else if (bstrState==TEXT("loading"))
				DisconnectDHtmlElementEvents();
		}
	}

	return S_OK;
}

LRESULT CBCGPDHtmlDialog::OnDestroyModeless(WPARAM, LPARAM)
{
	DestroyWindow();
	return 0;
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(CDataExchange* pDX, LPCTSTR szId, DISPID dispid, CString& value)
{
	DDX_DHtml_ElementText(szId, dispid, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(CDataExchange* pDX, LPCTSTR szId, DISPID dispid, short& value)
{
	DDX_DHtml_ElementText(szId, dispid, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(CDataExchange* pDX, LPCTSTR szId, DISPID dispid, int& value)
{
	DDX_DHtml_ElementText(szId, dispid, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(CDataExchange* pDX, LPCTSTR szId, DISPID dispid, long& value)
{
	DDX_DHtml_ElementText(szId, dispid, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(CDataExchange* pDX, LPCTSTR szId, DISPID dispid, DWORD& value)
{
	DDX_DHtml_ElementText(szId, dispid, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(CDataExchange* pDX, LPCTSTR szId, DISPID dispid, float& value)
{
	DDX_DHtml_ElementText(szId, dispid, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(CDataExchange* pDX, LPCTSTR szId, DISPID dispid, double& value)
{
	DDX_DHtml_ElementText(szId, dispid, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_CheckBox(CDataExchange* pDX, LPCTSTR szId, int& value)
{
	DDX_DHtml_CheckBox(szId, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_Radio(CDataExchange* pDX, LPCTSTR szId, long& value)
{
	DDX_DHtml_Radio(szId, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_SelectValue(CDataExchange* pDX, LPCTSTR szId, CString& value)
{
	DDX_DHtml_SelectValue(szId, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_SelectString(CDataExchange* pDX, LPCTSTR szId, CString& value)
{
	DDX_DHtml_SelectString(szId, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_SelectIndex(CDataExchange* pDX, LPCTSTR szId, long& value)
{
	DDX_DHtml_SelectIndex(szId, value, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_AxControl(CDataExchange *pDX, LPCTSTR szId, DISPID dispid, VARIANT& var)
{
	DDX_DHtml_AxControl(szId, dispid, var, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_AxControl(CDataExchange *pDX, LPCTSTR szId, LPCTSTR szPropName, VARIANT& var)
{
	DDX_DHtml_AxControl(szId, szPropName, var, pDX->m_bSaveAndValidate);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(LPCTSTR szId, DISPID dispid, CString& value, BOOL bSave)
{
	ATL::CComPtr<IHTMLDocument2> sphtmlDoc;
	GetDHtmlDocument(&sphtmlDoc);
	if (sphtmlDoc == NULL)
		return;

	ATL::CComVariant var;
	if (bSave)
	{
		var = GetElementProperty(szId, dispid);
		var.ChangeType(VT_BSTR);
		value = var.bstrVal;
	}
	else
	{
		var = (LPCTSTR)value;
		SetElementProperty(szId, dispid, &var);
	}
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(LPCTSTR szId, DISPID dispid, short& value, BOOL bSave)
{
	if (bSave)
		DDX_DHtml_ElementTextWithFormat(szId, dispid, _T("%sd"), AFX_IDP_PARSE_INT, bSave, &value);
	else
		DDX_DHtml_ElementTextWithFormat(szId, dispid, _T("%hd"), AFX_IDP_PARSE_INT, bSave, value);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(LPCTSTR szId, DISPID dispid, int& value, BOOL bSave)
{
	if (bSave)
		DDX_DHtml_ElementTextWithFormat(szId, dispid, _T("%d"), AFX_IDP_PARSE_INT, bSave, &value);
	else
		DDX_DHtml_ElementTextWithFormat(szId, dispid, _T("%d"), AFX_IDP_PARSE_INT, bSave, value);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(LPCTSTR szId, DISPID dispid, long& value, BOOL bSave)
{
	if (bSave)
		DDX_DHtml_ElementTextWithFormat(szId, dispid, _T("%ld"), AFX_IDP_PARSE_INT, bSave, &value);
	else
		DDX_DHtml_ElementTextWithFormat(szId, dispid, _T("%ld"), AFX_IDP_PARSE_INT, bSave, value);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(LPCTSTR szId, DISPID dispid, DWORD& value, BOOL bSave)
{
	if (bSave)
		DDX_DHtml_ElementTextWithFormat(szId, dispid, _T("%lu"), AFX_IDP_PARSE_UINT, bSave, &value);
	else
		DDX_DHtml_ElementTextWithFormat(szId, dispid, _T("%lu"), AFX_IDP_PARSE_UINT, bSave, value);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(LPCTSTR szId, DISPID dispid, float& value, BOOL bSave)
{
	DDX_DHtml_ElementTextFloatFormat(szId, dispid, &value, value, FLT_DIG, bSave);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementText(LPCTSTR szId, DISPID dispid, double& value, BOOL bSave)
{
	DDX_DHtml_ElementTextFloatFormat(szId, dispid, &value, value, DBL_DIG, bSave);
}

/////////////////////////////////////////////////////////////////////////////
// Data exchange for special controls
void CBCGPDHtmlDialog::DDX_DHtml_CheckBox(LPCTSTR szId, int& value, BOOL bSave)
{
	COleVariant var;

	ATL::CComPtr<IHTMLDocument2> sphtmlDoc;
	GetDHtmlDocument(&sphtmlDoc);
	if (sphtmlDoc == NULL)
		return;

	ATL::CComPtr<IHTMLOptionButtonElement> spOptionButton;
	HRESULT hr = S_OK;

	hr = GetElementInterface(szId, __uuidof(IHTMLOptionButtonElement), (void **) &spOptionButton);
	if (spOptionButton == NULL)
		goto Error;


	if (bSave)
	{
		VARIANT_BOOL bIndeterminate;
		VARIANT_BOOL bChecked;
		hr = spOptionButton->get_checked(&bChecked);
		if (FAILED(hr))
			goto Error;
		hr = spOptionButton->get_indeterminate(&bIndeterminate);
		if (FAILED(hr))
			goto Error;
		value = ((bChecked != ATL_VARIANT_FALSE) ? BST_CHECKED : BST_UNCHECKED) | 
			((bIndeterminate != ATL_VARIANT_FALSE) ? BST_INDETERMINATE : 0);
	}
	else
	{
		hr = spOptionButton->put_checked((value & BST_CHECKED) ? ATL_VARIANT_TRUE : ATL_VARIANT_FALSE);
		if (FAILED(hr))
			goto Error;
		hr = spOptionButton->put_indeterminate((value & BST_INDETERMINATE) ? ATL_VARIANT_TRUE : ATL_VARIANT_FALSE);
		if (FAILED(hr))
			goto Error;
	}
Error:
	if (FAILED(hr))
	{
//		TRACE(traceHtml, 0, "Failed DDX_DHtml_CheckBox\n");
	}
}

void CBCGPDHtmlDialog::DDX_DHtml_Radio(LPCTSTR szId, long& value, BOOL bSave)
{
	COleVariant var;

	ATL::CComPtr<IHTMLDocument2> sphtmlDoc;
	GetDHtmlDocument(&sphtmlDoc);
	if (sphtmlDoc == NULL)
		return;

	ATL::CComPtr<IHTMLOptionButtonElement> spOptionButton;
	ATL::CComPtr<IDispatch> spdispCollection;
	ATL::CComPtr<IHTMLElementCollection> spElementColl;
	BOOL bIsCollection = FALSE;
	ATL::CComPtr<IDispatch> spdispElem;

	HRESULT hr = S_OK;
	COleVariant varIndex;
	COleVariant varEmpty;
	varIndex.vt = VT_I4;
	varIndex.lVal = 0;

	// get the radio buttons in the group
	hr = GetElement(szId, &spdispCollection, &bIsCollection);
	if (spdispCollection && bIsCollection)
	{
		hr = spdispCollection->QueryInterface(__uuidof(IHTMLElementCollection), (void **) &spElementColl);
		if (spElementColl == NULL)
		{
//			TRACE(traceHtml, 0, "Error: Collection didn't support IHTMLElementCollection!\n");
			ASSERT(FALSE);
			goto Error;
		}

		if (bSave)
			value = -1;     // value if none found

		long lCount = 0;
		spElementColl->get_length(&lCount);

		for (long lCntr = 0; lCntr < lCount; lCntr++)
		{				
			spdispElem = NULL;
			spOptionButton = NULL;
			varIndex.lVal = lCntr;
			hr = spElementColl->item(varIndex, varEmpty, &spdispElem);
			if (spdispElem == NULL)
				break;
			hr = spdispElem->QueryInterface(__uuidof(IHTMLOptionButtonElement), (void **) &spOptionButton);

			if (bSave)
			{
				VARIANT_BOOL bChecked;
				hr = spOptionButton->get_checked(&bChecked);
				if (FAILED(hr))
					goto Error;
				if (bChecked != ATL_VARIANT_FALSE)
				{
					value = varIndex.lVal;
					break;
				}
			}
			else
			{
				if (varIndex.lVal == value)
				{
					spOptionButton->put_checked(ATL_VARIANT_TRUE);
					break;
				}
			}
		};
	}
Error:
	if (FAILED(hr))
	{
//		TRACE(traceHtml, 0, "Warning: Failed DDX_DHtml_Radio\n");
	}
}

void CBCGPDHtmlDialog::DDX_DHtml_SelectValue(LPCTSTR szId, CString& value, BOOL bSave)
{
	ATL::CComPtr<IHTMLDocument2> sphtmlDoc;
	GetDHtmlDocument(&sphtmlDoc);
	if (sphtmlDoc == NULL)
		return;

	ATL::CComBSTR bstrText;
	ATL::CComPtr<IHTMLSelectElement> spSelect;
	HRESULT hr = GetElementInterface(szId, __uuidof(IHTMLSelectElement), (void **) &spSelect);
	if (FAILED(hr))
		goto Error;

	if (bSave)
	{
		spSelect->get_value(&bstrText);
		if (bstrText)
			value = bstrText;
	}
	else
	{
		bstrText.Attach(value.AllocSysString());
		spSelect->put_value(bstrText);
	}

Error:
	return;
}

void CBCGPDHtmlDialog::DDX_DHtml_SelectString(LPCTSTR szId, CString& value, BOOL bSave)
{
	ATL::CComPtr<IHTMLDocument2> sphtmlDoc;
	GetDHtmlDocument(&sphtmlDoc);
	if (sphtmlDoc == NULL)
		return;

	COleVariant varEmpty, varIndex;

	ATL::CComPtr<IHTMLSelectElement> spSelect;
	ATL::CComPtr<IDispatch> spdispOption;
	ATL::CComPtr<IHTMLOptionElement> spOption;
	CComBSTR bstrText;
	HRESULT hr = S_OK;
	long lIndex=-1;

	hr = GetElementInterface(szId, __uuidof(IHTMLSelectElement), (void **) &spSelect);
	if (spSelect == NULL)
		return;

	if (bSave)
	{
		// get the selected item
		value.Empty();
		spSelect->get_selectedIndex(&lIndex);
		if (lIndex >= 0)
		{
			varIndex = lIndex;

			spSelect->item(varIndex, varEmpty, &spdispOption);
			if (spdispOption)
			{
				spdispOption->QueryInterface(__uuidof(IHTMLOptionElement), (void **) &spOption);
				if (spOption)
				{
					spOption->get_text(&bstrText);
					if (bstrText)
						value = bstrText;
				}
			}
		}
	}
	else
	{
		bstrText.Attach(value.AllocSysString());
		lIndex = Select_FindString(spSelect, bstrText, FALSE);
		spSelect->put_selectedIndex(lIndex);
	}
}

void CBCGPDHtmlDialog::DDX_DHtml_SelectIndex(LPCTSTR szId, long& value, BOOL bSave)
{
	ATL::CComPtr<IHTMLDocument2> sphtmlDoc;
	GetDHtmlDocument(&sphtmlDoc);
	if (sphtmlDoc == NULL)
		return;

	ATL::CComPtr<IHTMLSelectElement> spSelect;
	HRESULT hr;

	hr = GetElementInterface(szId, __uuidof(IHTMLSelectElement), (void **) &spSelect);
	if (spSelect == NULL)
		return;

	if (bSave)
		spSelect->get_selectedIndex(&value);
	else
		spSelect->put_selectedIndex(value);
}

void CBCGPDHtmlDialog::DDX_DHtml_AxControl(LPCTSTR szId, DISPID dispid, VARIANT& var, BOOL bSave)
{
	if (bSave)
		var = GetControlProperty(szId, dispid);
	else
		SetControlProperty(szId, dispid, &var);
}

void CBCGPDHtmlDialog::DDX_DHtml_AxControl(LPCTSTR szId, LPCTSTR szPropName, VARIANT& var, BOOL bSave)
{
	if (bSave)
		var = GetControlProperty(szId, szPropName);
	else
		SetControlProperty(szId, szPropName, &var);
}

void CBCGPDHtmlDialog::OnDDXError(LPCTSTR szId, UINT nIDPrompt, BOOL /*bSave*/)
{
	AfxMessageBox(nIDPrompt);
	// default implementation just sets the
	// focus back to the offending element
	SetFocusToElement(szId);
	ATL::CComPtr<IHTMLEventObj> sphtmlEvent;
	GetEvent(&sphtmlEvent);
	if (sphtmlEvent)
		sphtmlEvent->put_cancelBubble(ATL_VARIANT_TRUE);
	AfxThrowUserException();
}

void __cdecl CBCGPDHtmlDialog::DDX_DHtml_ElementTextWithFormat(LPCTSTR szId,
	DISPID dispid, LPCTSTR lpszFormat, UINT nIDPrompt, BOOL bSave, ...)
	// only supports windows output formats - no floating point
{
	va_list pData;
	va_start(pData, bSave);

	CString value;

	if (bSave)
	{
		DDX_DHtml_ElementText(szId, dispid, value, bSave);
		// the following works for %d, %u, %ld, %lu
		if (!_AfxSimpleScanf(value, lpszFormat, pData))
		{
			// set the focus to the offending element
			SetFocusToElement(szId);

			OnDDXError(szId, nIDPrompt, bSave);
			va_end(pData);
			return;
		}
	}
	else
	{
		CString strTemp ;
		strTemp.FormatV(lpszFormat,pData);
		// does not support floating point numbers - see dlgfloat.cpp
		DDX_DHtml_ElementText(szId, dispid, strTemp, bSave);
	}
	va_end(pData);
}

void CBCGPDHtmlDialog::DDX_DHtml_ElementTextFloatFormat(LPCTSTR szId,
	DISPID dispid, void* pData, double value, int nSizeGcvt, BOOL bSave)
{
	ATLASSERT(pData != NULL);

	CString strValue;
	if (bSave)
	{
		DDX_DHtml_ElementText(szId, dispid, strValue, bSave);
		double d;
		if (!_AfxSimpleFloatParse(strValue, d))
		{
			OnDDXError(szId, AFX_IDP_PARSE_REAL, bSave);
		}
		if (nSizeGcvt == FLT_DIG)
			*((float*)pData) = (float)d;
		else
			*((double*)pData) = d;
	}
	else
	{
		CString szBuffer;
		szBuffer.Format(_T("%.*g"), nSizeGcvt, value);
		DDX_DHtml_ElementText(szId, dispid, szBuffer, bSave);
	}
}


BOOL CBCGPDHtmlDialog::_AfxSimpleScanf(LPCTSTR lpszText,
	LPCTSTR lpszFormat, va_list pData)
{
	ATLASSERT(lpszText != NULL);
	ATLASSERT(lpszFormat != NULL);

	ATLASSERT(*lpszFormat == '%');
	lpszFormat++;        // skip '%'

	BOOL bLong = FALSE;
	BOOL bShort = FALSE;
	if (*lpszFormat == 'l')
	{
		bLong = TRUE;
		lpszFormat++;
	}
	else if (*lpszFormat == 's')
	{
		bShort = TRUE;
		lpszFormat++;
	}

	ATLASSERT(*lpszFormat == 'd' || *lpszFormat == 'u');
	ATLASSERT(lpszFormat[1] == '\0');

	while (*lpszText == ' ' || *lpszText == '\t')
		lpszText++;
	TCHAR chFirst = lpszText[0];
	long l, l2;
	if (*lpszFormat == 'd')
	{
		// signed
		l = _tcstol(lpszText, (LPTSTR*)&lpszText, 10);
		l2 = (int)l;
	}
	else
	{
		// unsigned
		if (*lpszText == '-')
			return FALSE;
		l = (long)_tcstoul(lpszText, (LPTSTR*)&lpszText, 10);
		l2 = (unsigned int)l;
	}
	if (l == 0 && chFirst != '0')
		return FALSE;   // could not convert

	while (*lpszText == ' ' || *lpszText == '\t')
		lpszText++;
	if (*lpszText != '\0')
		return FALSE;   // not terminated properly

	if (bShort)
	{
		if ((short)l != l)
			return FALSE;   // too big for short
		*va_arg(pData, short*) = (short)l;
	}
	else
	{
		ATLASSERT(sizeof(long) == sizeof(int));
		ATLASSERT(l == l2);
		*va_arg(pData, long*) = l;
	}

	// all ok
	return TRUE;
}

BOOL CBCGPDHtmlDialog::_AfxSimpleFloatParse(LPCTSTR lpszText, double& d)
{
	ATLASSERT(lpszText != NULL);
	while (*lpszText == ' ' || *lpszText == '\t')
		lpszText++;

	TCHAR chFirst = lpszText[0];
	d = _tcstod(lpszText, (LPTSTR*)&lpszText);
	if (d == 0.0 && chFirst != '0')
		return FALSE;   // could not convert
	while (*lpszText == ' ' || *lpszText == '\t')
		lpszText++;

	if (*lpszText != '\0')
		return FALSE;   // not terminated properly

	return TRUE;
}

long CBCGPDHtmlDialog::Select_FindString(IHTMLSelectElement *pSelect, BSTR bstr, BOOL /*fExact*/)
{
	long lIndexFound = -1;
	COleVariant varIndex, varEmpty;
	ATL::CComPtr<IDispatch> spdispOption;
	ATL::CComPtr<IHTMLOptionElement> spOption;
	long lCount = 0;

	HRESULT hr = pSelect->get_length(&lCount);
	if (FAILED(hr))
		goto Error;
	varIndex.vt = VT_I4;
	// loop through the items searching for the string
	for (varIndex.lVal=0; varIndex.lVal<lCount; varIndex.lVal++)
	{
		pSelect->item(varIndex, varEmpty, &spdispOption);
		if (spdispOption)
		{
			spdispOption->QueryInterface(__uuidof(IHTMLOptionElement), (void **) &spOption);
			if (spOption)
			{
				CComBSTR bstrText;
				spOption->get_text(&bstrText);
				if (bstrText && !wcscmp(bstrText, bstr))
				{
					// we found it
					lIndexFound = varIndex.lVal;
					break;
				}
				spOption = NULL;
			}
			spdispOption = NULL;
		}
	}

Error:
	return lIndexFound;
}

void CBCGPDHtmlDialog::SetFocusToElement(LPCTSTR szId)
{
	// check if the element is a control element
	ATL::CComPtr<IHTMLControlElement> sphtmlCtrlElem;
	HRESULT hr = GetElementInterface(szId, &sphtmlCtrlElem);
	if (sphtmlCtrlElem)
	{
		sphtmlCtrlElem->focus();
		return;
	}

	// check if the element is an anchor element
	ATL::CComPtr<IHTMLAnchorElement> sphtmlAnchorElem;
	hr = GetElementInterface(szId, &sphtmlAnchorElem);
	if (sphtmlAnchorElem)
	{
		sphtmlAnchorElem->focus();
		return;
	}

	// otherwise all we can do is scroll the element into view
	ATL::CComPtr<IHTMLElement> sphtmlElem;
	hr = GetElementInterface(szId, &sphtmlElem);
	if (sphtmlElem)
	{
		ATL::CComVariant var;
		var.vt = VT_BOOL;
		var.boolVal = ATL_VARIANT_TRUE;
		sphtmlElem->scrollIntoView(var);
	}
}

HRESULT CBCGPDHtmlDialog::ConnectDHtmlEvents(IUnknown *punkDoc)
{
	return ConnectToConnectionPoint(punkDoc, __uuidof(HTMLDocumentEvents), &m_dwDHtmlEventSinkCookie);
}

void CBCGPDHtmlDialog::DisconnectDHtmlEvents()
{
	ATL::CComPtr<IHTMLDocument2> sphtmlDoc;
	GetDHtmlDocument(&sphtmlDoc);

	if (sphtmlDoc == NULL)
		return;
	DisconnectFromConnectionPoint(sphtmlDoc, __uuidof(HTMLDocumentEvents), m_dwDHtmlEventSinkCookie);
	DisconnectDHtmlElementEvents();
}

HRESULT CBCGPDHtmlDialog::ConnectDHtmlElementEvents(DWORD_PTR dwThunkOffset /*= 0*/)
{
	HRESULT hr = S_OK;
	const DHtmlEventMapEntry* pEventMap = GetDHtmlEventMap();
	if (!pEventMap)
		return hr;

	for (int i=0; pEventMap[i].nType != DHTMLEVENTMAPENTRY_END; i++)
	{
		if (pEventMap[i].nType==DHTMLEVENTMAPENTRY_ELEMENT)
		{
			// an element name must be specified when using element events
			ATLASSERT(pEventMap[i].szName);

			// connect to the element's event sink
			ATL::CComPtr<IDispatch> spdispElement;
			GetElement(pEventMap[i].szName, &spdispElement);
			if (spdispElement)
			{
				if (!IsSinkedElement(spdispElement))
				{
					CDHtmlElementEventSink *pSink = NULL;
					ATLTRY(pSink = new CDHtmlElementEventSink(this, spdispElement));
					if (pSink == NULL)
						return E_OUTOFMEMORY;
					hr = AtlAdvise(spdispElement, pSink, __uuidof(IDispatch), &pSink->m_dwCookie);
					if (SUCCEEDED(hr))
						m_SinkedElements.Add(pSink);
					else
						delete pSink;
#ifdef _DEBUG
					if (FAILED(hr)) {}
//						TRACE(traceHtml, 0, "Warning: Failed to connect to ConnectionPoint!\n");
#endif
				}
			}
		}
		else if (pEventMap[i].nType==DHTMLEVENTMAPENTRY_CONTROL)
		{
			// check if we already have a sink connected to this control
			if (!FindSinkForObject(pEventMap[i].szName))
			{
				// create a new sink and
				// connect it to the element's event sink
				ATL::CComPtr<IDispatch> spdispElement;
				GetElement(pEventMap[i].szName, &spdispElement);
				if (spdispElement)
				{
					ATL::CComPtr<IHTMLObjectElement> sphtmlObj;
					spdispElement->QueryInterface(__uuidof(IHTMLObjectElement), (void **) &sphtmlObj);
					if (sphtmlObj)
					{
						ATL::CComPtr<IDispatch> spdispControl;
						sphtmlObj->get_object(&spdispControl);
						if (spdispControl)
						{
							// create a new control sink to connect to the control's events
							CDHtmlControlSink *pSink = NULL; 
							ATLTRY(pSink = new CDHtmlControlSink(spdispControl, this, pEventMap[i].szName, dwThunkOffset));
							if (pSink == NULL)
								return E_OUTOFMEMORY;
							m_ControlSinks.Add(pSink);
						}
					}
				}
			}
		}
	}
	return hr;
}

BOOL CBCGPDHtmlDialog::FindSinkForObject(LPCTSTR szName)
{	
	int nLength = m_ControlSinks.GetSize();
	if (nLength > 0)
	{
		ENSURE_ARG(szName!=NULL);
	}
	for (int i=0; i<nLength; i++)
	{
		if (!_tcscmp(szName, m_ControlSinks[i]->m_szControlId))
			return TRUE;
	}
	return FALSE;
}

BOOL CBCGPDHtmlDialog::IsSinkedElement(IDispatch *pdispElem)
{
	ENSURE_ARG(pdispElem!=NULL);
	ATL::CComPtr<IUnknown> spunk;
	pdispElem->QueryInterface(__uuidof(IUnknown), (void **) &spunk);
	if (!spunk)
		return FALSE;
	for (int i=0; i<m_SinkedElements.GetSize(); i++)
	{
		if (spunk == m_SinkedElements[i]->m_spunkElem)
			return TRUE;
	}
	return FALSE;
}

void CBCGPDHtmlDialog::DisconnectDHtmlElementEvents()
{
	const DHtmlEventMapEntry* pEventMap = GetDHtmlEventMap();

	if (!pEventMap)
		return;

	int i;

	// disconnect from element events
	for (i=0; i<m_SinkedElements.GetSize(); i++)
	{
		CDHtmlElementEventSink *pSink = m_SinkedElements[i];
		AtlUnadvise(pSink->m_spunkElem, __uuidof(IDispatch), pSink->m_dwCookie);
		delete pSink;
	}
	m_SinkedElements.RemoveAll();

	// disconnect from control events
	for (i=0; i<m_ControlSinks.GetSize(); i++)
	{
		DisconnectFromConnectionPoint(m_ControlSinks[i]->m_spunkObj, 
				m_ControlSinks[i]->m_iid, m_ControlSinks[i]->m_dwCookie);
		delete m_ControlSinks[i];
	}
	m_ControlSinks.RemoveAll();
	return;
}

HRESULT CBCGPDHtmlDialog::GetElement(LPCTSTR szElementId, IDispatch **ppdisp, 
								 BOOL *pbCollection /*= NULL*/)
{
	ENSURE_ARG(ppdisp!=NULL);
	ATL::CComPtr<IHTMLElementCollection> sphtmlAll;
	ATL::CComPtr<IHTMLElementCollection> sphtmlColl;
	ATL::CComPtr<IDispatch> spdispElem;
	ATL::CComVariant varName;
	ATL::CComVariant varIndex;
	HRESULT hr = S_OK;
	ATL::CComPtr<IHTMLDocument2> sphtmlDoc;

	*ppdisp = NULL;

	if (pbCollection)
		*pbCollection = FALSE;

	hr = GetDHtmlDocument(&sphtmlDoc);
	if (sphtmlDoc == NULL)
		return hr;

	if(szElementId == NULL)
		return E_INVALIDARG;
		
	const CString strElementId(szElementId);
	varName.vt = VT_BSTR;
	varName.bstrVal = strElementId.AllocSysString();
	
#ifndef _UNICODE	
	if (!varName.bstrVal)
	{
		hr = E_OUTOFMEMORY;
		goto Error;
	}
#endif

	hr = sphtmlDoc->get_all(&sphtmlAll);
	if (sphtmlAll == NULL)
		goto Error;
	hr = sphtmlAll->item(varName, varIndex, &spdispElem);
	if (spdispElem == NULL)
	{
		hr = E_NOINTERFACE;
		goto Error;
	}

	spdispElem->QueryInterface(__uuidof(IHTMLElementCollection), (void **) &sphtmlColl);
	if (sphtmlColl)
	{
		if (pbCollection)
			*pbCollection = TRUE;
#ifdef _DEBUG
		else
		{
//			TRACE(traceHtml, 0, "Warning: duplicate IDs or NAMEs.\n");
			ATLASSERT(FALSE);
		}
#endif

	}
Error:
	if (SUCCEEDED(hr))
	{
		*ppdisp = spdispElem;
		if (spdispElem)
			(*ppdisp)->AddRef();
	}
	return hr;
}

HRESULT CBCGPDHtmlDialog::GetElement(LPCTSTR szElementId, IHTMLElement **pphtmlElement)
{
	return GetElementInterface(szElementId, __uuidof(IHTMLElement), (void **) pphtmlElement);
}

HRESULT CBCGPDHtmlDialog::GetElementInterface(LPCTSTR szElementId, REFIID riid, void **ppvObj)
{
	ENSURE_ARG(ppvObj!=NULL);
	HRESULT hr = E_NOINTERFACE;
	*ppvObj = NULL;
	ATL::CComPtr<IDispatch> spdispElem;

	hr = GetElement(szElementId, &spdispElem);

	if (spdispElem)
		hr = spdispElem->QueryInterface(riid, ppvObj);
	return hr;
}

BSTR CBCGPDHtmlDialog::GetElementText(LPCTSTR szElementId)
{
	BSTR bstrText = NULL;
	ATL::CComPtr<IHTMLElement> sphtmlElem;
	GetElement(szElementId, &sphtmlElem);
	if (sphtmlElem)
		sphtmlElem->get_innerText(&bstrText);
	return bstrText;
}

void CBCGPDHtmlDialog::SetElementText(LPCTSTR szElementId, BSTR bstrText)
{
	ATL::CComPtr<IHTMLElement> sphtmlElem;
	GetElement(szElementId, &sphtmlElem);
	if (sphtmlElem)
		sphtmlElem->put_innerText(bstrText);
}

void CBCGPDHtmlDialog::SetElementText(IUnknown *punkElem, BSTR bstrText)
{
	ENSURE_ARG(punkElem!=NULL);
	ATL::CComPtr<IHTMLElement> sphtmlElem;
	punkElem->QueryInterface(__uuidof(IHTMLElement), (void **) &sphtmlElem);
	if (sphtmlElem != NULL)
		sphtmlElem->put_innerText(bstrText);
}

BSTR CBCGPDHtmlDialog::GetElementHtml(LPCTSTR szElementId)
{
	BSTR bstrText = NULL;
	ATL::CComPtr<IHTMLElement> sphtmlElem;
	GetElement(szElementId, &sphtmlElem);
	if (sphtmlElem)
		sphtmlElem->get_innerHTML(&bstrText);
	return bstrText;
}

void CBCGPDHtmlDialog::SetElementHtml(LPCTSTR szElementId, BSTR bstrText)
{
	ATL::CComPtr<IHTMLElement> sphtmlElem;
	GetElement(szElementId, &sphtmlElem);
	if (sphtmlElem)
		sphtmlElem->put_innerHTML(bstrText);
}

void CBCGPDHtmlDialog::SetElementHtml(IUnknown *punkElem, BSTR bstrText)
{
	ENSURE_ARG(punkElem!=NULL);
	ATL::CComPtr<IHTMLElement> sphtmlElem;
	punkElem->QueryInterface(__uuidof(IHTMLElement), (void **) &sphtmlElem);
	if (sphtmlElem != NULL)
		sphtmlElem->put_innerHTML(bstrText);
}

VARIANT CBCGPDHtmlDialog::GetElementProperty(LPCTSTR szElementId, DISPID dispid)
{
	VARIANT varRet;
	ATL::CComPtr<IDispatch> spdispElem;
	varRet.vt = VT_EMPTY;
	GetElement(szElementId, &spdispElem);
	if (spdispElem)
	{
		DISPPARAMS dispparamsNoArgs = { NULL, NULL, 0, 0 };
		spdispElem->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, 
			DISPATCH_PROPERTYGET, &dispparamsNoArgs, &varRet, NULL, NULL);
	}
	return varRet;
}

void CBCGPDHtmlDialog::SetElementProperty(LPCTSTR szElementId, DISPID dispid, VARIANT *pVar)
{
	ATL::CComPtr<IDispatch> spdispElem;

	GetElement(szElementId, &spdispElem);
	if (spdispElem)
	{
		DISPPARAMS dispparams = {NULL, NULL, 1, 1};
		dispparams.rgvarg = pVar;
		DISPID dispidPut = DISPID_PROPERTYPUT;
		dispparams.rgdispidNamedArgs = &dispidPut;

		spdispElem->Invoke(dispid, IID_NULL,
				LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
				&dispparams, NULL, NULL, NULL);
	}
}

HRESULT CBCGPDHtmlDialog::GetControlDispatch(LPCTSTR szId, IDispatch **ppdisp)
{
	HRESULT hr = S_OK;
	ATL::CComPtr<IDispatch> spdispElem;

	hr = GetElement(szId, &spdispElem);

	if (spdispElem)
	{
		ATL::CComPtr<IHTMLObjectElement> sphtmlObj;

		hr = spdispElem.QueryInterface(&sphtmlObj);
		if (sphtmlObj)
		{
			spdispElem.Release();
			hr = sphtmlObj->get_object(ppdisp);
		}
	}
	return hr;
}

VARIANT CBCGPDHtmlDialog::GetControlProperty(IDispatch *pdispControl, DISPID dispid)
{
	VARIANT varRet;
	varRet.vt = VT_EMPTY;
	if (pdispControl)
	{
		DISPPARAMS dispparamsNoArgs = { NULL, NULL, 0, 0 };
		pdispControl->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, 
			DISPATCH_PROPERTYGET, &dispparamsNoArgs, &varRet, NULL, NULL);
	}
	return varRet;
}

VARIANT CBCGPDHtmlDialog::GetControlProperty(LPCTSTR szId, DISPID dispid)
{
	ATL::CComPtr<IDispatch> spdispElem;

	GetControlDispatch(szId, &spdispElem);
	return GetControlProperty(spdispElem, dispid);
}

VARIANT CBCGPDHtmlDialog::GetControlProperty(LPCTSTR szId, LPCTSTR szPropName)
{
	ATL::CComVariant varEmpty;
	ATL::CComPtr<IDispatch> spdispElem;

	GetControlDispatch(szId, &spdispElem);
	if (!spdispElem)
		return varEmpty;

	DISPID dispid;
	USES_CONVERSION;
	LPOLESTR pPropName = (LPOLESTR)T2COLE(szPropName);
	HRESULT hr = spdispElem->GetIDsOfNames(IID_NULL, &pPropName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (SUCCEEDED(hr))
		return GetControlProperty(spdispElem, dispid);
	return varEmpty;
}

void CBCGPDHtmlDialog::SetControlProperty(IDispatch *pdispControl, DISPID dispid, VARIANT *pVar)
{
	if (pdispControl != NULL)
	{
		DISPPARAMS dispparams = {NULL, NULL, 1, 1};
		dispparams.rgvarg = pVar;
		DISPID dispidPut = DISPID_PROPERTYPUT;
		dispparams.rgdispidNamedArgs = &dispidPut;

		pdispControl->Invoke(dispid, IID_NULL,
				LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
				&dispparams, NULL, NULL, NULL);
	}
}

void CBCGPDHtmlDialog::SetControlProperty(LPCTSTR szElementId, DISPID dispid, VARIANT *pVar)
{
	ATL::CComPtr<IDispatch> spdispElem;
	GetControlDispatch(szElementId, &spdispElem);

	SetControlProperty(spdispElem, dispid, pVar);
}

void CBCGPDHtmlDialog::SetControlProperty(LPCTSTR szElementId, LPCTSTR szPropName, VARIANT *pVar)
{
	ATL::CComPtr<IDispatch> spdispElem;
	GetControlDispatch(szElementId, &spdispElem);
	if (!spdispElem)
		return;
	DISPID dispid;
	USES_CONVERSION;
	LPOLESTR pPropName = (LPOLESTR)T2COLE(szPropName);
	HRESULT hr = spdispElem->GetIDsOfNames(IID_NULL, &pPropName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (SUCCEEDED(hr))
		SetControlProperty(spdispElem, dispid, pVar);
}

HRESULT CBCGPDHtmlDialog::GetEvent(IHTMLEventObj **ppEventObj)
{
	ENSURE_ARG(ppEventObj!=NULL);
	ATL::CComPtr<IHTMLWindow2> sphtmlWnd;
	ATL::CComPtr<IHTMLDocument2> sphtmlDoc;

	*ppEventObj = NULL;

	HRESULT hr = GetDHtmlDocument(&sphtmlDoc);
	if (sphtmlDoc == NULL)
		return hr;

	hr = sphtmlDoc->get_parentWindow(&sphtmlWnd);
	if (FAILED(hr))
		goto Error;
	hr = sphtmlWnd->get_event(ppEventObj);

Error:
	return hr;
}

void CBCGPDHtmlDialog::SetHostFlags(DWORD dwFlags)
{
	m_dwHostFlags = dwFlags;
}

void CBCGPDHtmlDialog::SetExternalDispatch(IDispatch *pdispExternal)
{
	m_spExternalDisp = pdispExternal;
}

BOOL CBCGPDHtmlDialog::IsExternalDispatchSafe()
{
	return FALSE;
}


STDMETHODIMP CBCGPDHtmlDialog::ShowContextMenu(DWORD dwID, POINT *ppt, IUnknown *pcmdtReserved, IDispatch *pdispReserved)
{
	dwID; // unused
	ppt; // unused
	pcmdtReserved; // unused
	pdispReserved; // unused

	return S_FALSE;
}
STDMETHODIMP CBCGPDHtmlDialog::GetHostInfo(DOCHOSTUIINFO *pInfo)
{
	pInfo->dwFlags = m_dwHostFlags;
	return S_OK;
}
STDMETHODIMP CBCGPDHtmlDialog::ShowUI(DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc)
{
	dwID; // unused
	pActiveObject; // unused
	pCommandTarget; // unused
	pFrame; // unused
	pDoc; // unused

	return S_FALSE;
}
STDMETHODIMP CBCGPDHtmlDialog::HideUI(void)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBCGPDHtmlDialog::UpdateUI(void)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBCGPDHtmlDialog::EnableModeless(BOOL fEnable)
{
	fEnable; // unused

	return E_NOTIMPL;
}
STDMETHODIMP CBCGPDHtmlDialog::OnDocWindowActivate(BOOL fActivate)
{
	fActivate; // unused

	return E_NOTIMPL;
}
STDMETHODIMP CBCGPDHtmlDialog::OnFrameWindowActivate(BOOL fActivate)
{
	fActivate; // unused

	return E_NOTIMPL;
}
STDMETHODIMP CBCGPDHtmlDialog::ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fRameWindow)
{
	prcBorder; // unused
	pUIWindow; // unused
	fRameWindow; // unused

	return E_NOTIMPL;
}
STDMETHODIMP CBCGPDHtmlDialog::TranslateAccelerator(LPMSG lpMsg, const GUID *pguidCmdGroup, DWORD nCmdID)
{
	lpMsg; // unused
	pguidCmdGroup; // unused
	nCmdID; // unused

	return S_FALSE;
}
STDMETHODIMP CBCGPDHtmlDialog::GetOptionKeyPath(LPOLESTR *pchKey, DWORD dw)
{
	pchKey; // unused
	dw; // unused

	return E_NOTIMPL;
}
STDMETHODIMP CBCGPDHtmlDialog::GetDropTarget(IDropTarget *pDropTarget, IDropTarget **ppDropTarget)
{
	pDropTarget; // unused

	*ppDropTarget = NULL;
	return E_NOTIMPL;
}
STDMETHODIMP CBCGPDHtmlDialog::GetExternal(IDispatch **ppDispatch)
{
	if(ppDispatch == NULL)
		return E_POINTER;
		
	*ppDispatch = NULL;
	if (m_spExternalDisp.p && CanAccessExternal())
	{
		m_spExternalDisp.p->AddRef();
		*ppDispatch = m_spExternalDisp.p;
		return S_OK;
	}
	return E_NOTIMPL;
}
STDMETHODIMP CBCGPDHtmlDialog::TranslateUrl(DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut)
{
	dwTranslate; // unused
	pchURLIn; // unused
	
	if(ppchURLOut == NULL)
		return E_POINTER;
		
	*ppchURLOut = NULL;
	return S_FALSE;
}
STDMETHODIMP CBCGPDHtmlDialog::FilterDataObject(IDataObject *pDO, IDataObject **ppDORet)
{
	pDO; // unused;
	
	if(ppDORet == NULL)
		return E_POINTER;
		
	*ppDORet = NULL;
	return S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// CBCGPMultiPageDHtmlDialog

IMPLEMENT_DYNAMIC(CBCGPMultiPageDHtmlDialog, CBCGPDHtmlDialog)

CBCGPMultiPageDHtmlDialog::CBCGPMultiPageDHtmlDialog() : CBCGPDHtmlDialog()
{
	m_pCurrentMap = NULL;
}

CBCGPMultiPageDHtmlDialog::CBCGPMultiPageDHtmlDialog(UINT nIDTemplate, UINT nHtmlResID /*= 0*/, CWnd *pParentWnd /*= NULL*/, UINT nIDBrowserPlaceHolder/* = 0*/) : 
	CBCGPDHtmlDialog(nIDTemplate, nHtmlResID, pParentWnd, nIDBrowserPlaceHolder)
{
	m_pCurrentMap = NULL;
}

CBCGPMultiPageDHtmlDialog::CBCGPMultiPageDHtmlDialog(LPCTSTR lpszTemplateName, LPCTSTR szHtmlResID /*= NULL*/, CWnd *pParentWnd /*= NULL*/, UINT nIDBrowserPlaceHolder/* = 0*/) :
	CBCGPDHtmlDialog(lpszTemplateName, szHtmlResID, pParentWnd, nIDBrowserPlaceHolder)
{
	m_pCurrentMap = NULL;
}

CBCGPMultiPageDHtmlDialog::~CBCGPMultiPageDHtmlDialog()
{
}

const DHtmlEventMapEntry* CBCGPMultiPageDHtmlDialog::GetEventMapForUrl(LPCTSTR szUrl)
{
	szUrl; // unused

	return NULL;
}

const DHtmlEventMapEntry* CBCGPMultiPageDHtmlDialog::GetDHtmlEventMap()
{
	return m_pCurrentMap;
}

void CBCGPMultiPageDHtmlDialog::OnNavigateComplete(LPDISPATCH pDisp, LPCTSTR szUrl)
{
	if (pDisp != m_pBrowserApp)
		return;
	CString strUrl = szUrl;
	if (!strUrl.Left(4).CompareNoCase(_T("res:")))
	{
		int nIndex = strUrl.ReverseFind('/');
		if (nIndex >= 0)
			strUrl = strUrl.Mid(nIndex+1);
	}
	m_pCurrentMap = GetEventMapForUrl(strUrl);
	CBCGPDHtmlDialog::OnNavigateComplete(pDisp, szUrl);
}

