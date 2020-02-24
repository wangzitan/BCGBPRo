// BCGPXml.cpp: implementation of the CBCGPXml class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "bcgcbpro.h"
#include "BCGPXml.h"
#include "bcgglobals.h"

#include "bcg_msxml6.tlh"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif


static const LPCTSTR c_szXml_Version	   = _T("version");
static const LPCTSTR c_szXml_Encoding	   = _T("encoding");
static const LPCTSTR c_szXml_Yes		   = _T("yes");
static const LPCTSTR c_szXml_No 		   = _T("no");


static const LPCTSTR c_szXml_FormatXSLT    = 
_T("<xsl:stylesheet version=\"1.0\" xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\">") \
_T("  <xsl:output method=\"xml\" indent=\"yes\"/>") \
_T("  <xsl:template match=\"@*|node()\">") \
_T("    <xsl:copy>") \
_T("      <xsl:apply-templates select=\"@*|node()\"/>") \
_T("    </xsl:copy>") \
_T("  </xsl:template>") \
_T("</xsl:stylesheet>");


const LPCTSTR CBCGPXmlProcessingInstruction::c_szEncoding_Utf8	  = _T("utf-8");
const LPCTSTR CBCGPXmlProcessingInstruction::c_szEncoding_Utf16   = _T("utf-16");


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

#define BCGP_HR_SUCCEEDED_OK(hr) (SUCCEEDED(hr) && hr == S_OK)

#if _MSC_VER < 1500
	#define BCGP_VARIANT_TRUE ((VARIANT_BOOL)-1)
	#define BCGP_VARIANT_FALSE ((VARIANT_BOOL)0)
#else
	#define BCGP_VARIANT_TRUE VARIANT_TRUE
	#define BCGP_VARIANT_FALSE VARIANT_FALSE
#endif

static inline BOOL BCGPIsStringEmpty(LPCTSTR str)
{
	return str == NULL || _tcslen(str) == 0;
}

#define BCG_MSXML_DOMNODE(value) ((BCG_MSXML::IXMLDOMNode*)((value).GetInterface()))
#define BCG_MSXML_DOMNODE_PTR(value) ((BCG_MSXML::IXMLDOMNode**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMCHARACTERDATA(value) ((BCG_MSXML::IXMLDOMCharacterData*)((value).GetInterface()))
#define BCG_MSXML_DOMCHARACTERDATA_PTR(value) ((BCG_MSXML::IXMLDOMCharacterData**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMTEXT(value) ((BCG_MSXML::IXMLDOMText*)((value).GetInterface()))
#define BCG_MSXML_DOMTEXT_PTR(value) ((BCG_MSXML::IXMLDOMText**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMATTRIBUTE(value) ((BCG_MSXML::IXMLDOMAttribute*)((value).GetInterface()))
#define BCG_MSXML_DOMATTRIBUTE_PTR(value) ((BCG_MSXML::IXMLDOMAttribute**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMELEMENT(value) ((BCG_MSXML::IXMLDOMElement*)((value).GetInterface()))
#define BCG_MSXML_DOMELEMENT_PTR(value) ((BCG_MSXML::IXMLDOMElement**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMENTITY(value) ((BCG_MSXML::IXMLDOMEntity*)((value).GetInterface()))
#define BCG_MSXML_DOMENTITY_PTR(value) ((BCG_MSXML::IXMLDOMEntity**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMENTITYREFERENCE(value) ((BCG_MSXML::IXMLDOMEntityReference*)((value).GetInterface()))
#define BCG_MSXML_DOMENTITYREFERENCE_PTR(value) ((BCG_MSXML::IXMLDOMEntityReference**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMCDATASECTION(value) ((BCG_MSXML::IXMLDOMCDATASection*)((value).GetInterface()))
#define BCG_MSXML_DOMCDATASECTION_PTR(value) ((BCG_MSXML::IXMLDOMCDATASection**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMCOMMENT(value) ((BCG_MSXML::IXMLDOMComment*)((value).GetInterface()))
#define BCG_MSXML_DOMCOMMENT_PTR(value) ((BCG_MSXML::IXMLDOMComment**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMPROCESSINGINSTRUCTION(value) ((BCG_MSXML::IXMLDOMProcessingInstruction*)((value).GetInterface()))
#define BCG_MSXML_DOMPROCESSINGINSTRUCTION_PTR(value) ((BCG_MSXML::IXMLDOMProcessingInstruction**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMIMPLEMENTATION(value) ((BCG_MSXML::IXMLDOMImplementation*)((value).GetInterface()))
#define BCG_MSXML_DOMIMPLEMENTATION_PTR(value) ((BCG_MSXML::IXMLDOMImplementation**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMNOTATION(value) ((BCG_MSXML::IXMLDOMNotation*)((value).GetInterface()))
#define BCG_MSXML_DOMNOTATION_PTR(value) ((BCG_MSXML::IXMLDOMNotation**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMDOCUMENTFRAGMENT(value) ((BCG_MSXML::IXMLDOMDocumentFragment*)((value).GetInterface()))
#define BCG_MSXML_DOMDOCUMENTFRAGMENT_PTR(value) ((BCG_MSXML::IXMLDOMDocumentFragment**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMDOCUMENTTYPE(value) ((BCG_MSXML::IXMLDOMDocumentType*)((value).GetInterface()))
#define BCG_MSXML_DOMDOCUMENTTYPE_PTR(value) ((BCG_MSXML::IXMLDOMDocumentType**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMNODELIST(value) ((BCG_MSXML::IXMLDOMNodeList*)((value).GetInterface()))
#define BCG_MSXML_DOMNODELIST_PTR(value) ((BCG_MSXML::IXMLDOMNodeList**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMNAMEDNODEMAP(value) ((BCG_MSXML::IXMLDOMNamedNodeMap*)((value).GetInterface()))
#define BCG_MSXML_DOMNAMEDNODEMAP_PTR(value) ((BCG_MSXML::IXMLDOMNamedNodeMap**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMDOCUMENT(value) ((BCG_MSXML::IXMLDOMDocument*)((value).GetInterface()))
#define BCG_MSXML_DOMDOCUMENT_PTR(value) ((BCG_MSXML::IXMLDOMDocument**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMPARSEERROR(value) ((BCG_MSXML::IXMLDOMParseError*)((value).GetInterface()))
#define BCG_MSXML_DOMPARSEERROR_PTR(value) ((BCG_MSXML::IXMLDOMParseError**)((value).GetInterfacePtr()))

#define BCG_MSXML_DOMPARSEERRORCOLLECTION(value) ((BCG_MSXML::IXMLDOMParseErrorCollection*)((value).GetInterface()))
#define BCG_MSXML_DOMPARSEERRORCOLLECTION_PTR(value) ((BCG_MSXML::IXMLDOMParseErrorCollection**)((value).GetInterfacePtr()))


#ifndef V_I8
	#define V_I8(X)          (*(LONGLONG*)&V_UNION(X, lVal))
#endif

#ifndef V_UI8
	#define V_UI8(X)         (*(ULONGLONG*)&V_UNION(X, ulVal))
#endif


template<class T>
struct _TBCGPVariantTypeInfo
{
	enum
	{
		value = VT_EMPTY
	};
};

template<class T>
struct _TBCGPVariantTypeBSTR
{
	enum
	{
		value = false
	};
};

template<>
struct _TBCGPVariantTypeBSTR<BSTR>
{
	enum
	{
		value = true
	};
};

template<> struct _TBCGPVariantTypeInfo<bool>             { enum{ value = VT_BOOL}; };
template<> struct _TBCGPVariantTypeInfo<char>             { enum{ value = VT_I1};   };
template<> struct _TBCGPVariantTypeInfo<signed char>      { enum{ value = VT_I1};   };
template<> struct _TBCGPVariantTypeInfo<unsigned char>    { enum{ value = VT_UI1};  };
template<> struct _TBCGPVariantTypeInfo<short>            { enum{ value = VT_I2};   };
template<> struct _TBCGPVariantTypeInfo<unsigned short>   { enum{ value = VT_UI2};  };
template<> struct _TBCGPVariantTypeInfo<long>             { enum{ value = VT_I4};   };
template<> struct _TBCGPVariantTypeInfo<unsigned long>    { enum{ value = VT_UI4};  };
template<> struct _TBCGPVariantTypeInfo<int>              { enum{ value = VT_INT};  };
template<> struct _TBCGPVariantTypeInfo<unsigned int>     { enum{ value = VT_UINT}; };
#if (_WIN32_WINNT >= 0x0501)
	template<> struct _TBCGPVariantTypeInfo<__int64>          { enum{ value = VT_I8};   };
	template<> struct _TBCGPVariantTypeInfo<unsigned __int64> { enum{ value = VT_UI8};  };
#endif
template<> struct _TBCGPVariantTypeInfo<float>            { enum{ value = VT_R4};   };
template<> struct _TBCGPVariantTypeInfo<double>           { enum{ value = VT_R8};   };
template<> struct _TBCGPVariantTypeInfo<long double>      { enum{ value = VT_R8};   };

BOOL
_BCGPVariantIsValid(VARTYPE vt)
{
	return vt != VT_NULL && vt != VT_EMPTY;
}

BOOL
_BCGPVariantIsValid(const VARIANT& value)
{
	return _BCGPVariantIsValid(value.vt);
}

BOOL
_BCGPVariantIsString(VARTYPE vt)
{
	return vt == VT_BSTR || vt == VT_LPSTR || vt == VT_LPWSTR;
}

BOOL
_BCGPVariantIsString(const VARIANT& value)
{
	return _BCGPVariantIsString(value.vt);
}

BOOL
_BCGPVariantIsDateTime(VARTYPE vt)
{
	return vt == VT_DATE || vt == VT_FILETIME;
}

BOOL
_BCGPVariantIsDateTime(const VARIANT& value)
{
	return _BCGPVariantIsDateTime(value.vt);
}

BOOL
_BCGPVariantIsLanguageSpecific(VARTYPE vt)
{
	return _BCGPVariantIsString(vt) || _BCGPVariantIsDateTime(vt) || vt == VT_DISPATCH;
}

BOOL
_BCGPVariantIsLanguageSpecific(const VARIANT& value)
{
	return _BCGPVariantIsLanguageSpecific(value.vt);
}

#ifndef LANG_NEUTRAL
	#define LANG_NEUTRAL 0x00
#endif

#ifndef LANG_INVARIANT
	#define LANG_INVARIANT 0x7f
#endif

#ifndef LOCALE_INVARIANT
	#define LOCALE_INVARIANT (MAKELCID(MAKELANGID(LANG_INVARIANT, SUBLANG_NEUTRAL), SORT_DEFAULT))
#endif

HRESULT
_BCGPVariantChangeTypeEx(VARTYPE vt, VARIANT& dst, const VARIANT* src, LCID lcid = LOCALE_INVARIANT)
{
    if (vt == VT_EMPTY)
    {
        vt = V_VT(&dst);
    }

    if (src == NULL)
    {
        src = &dst;
    }

    HRESULT hres = S_OK;

    if ((&dst != src) || (vt != V_VT(&dst)))
    {
        hres = ::VariantChangeTypeEx
                    (
                        &dst,
                        (VARIANTARG*)src,
                        lcid,
                        (vt == VT_BOOL ? VARIANT_ALPHABOOL : 0) | VARIANT_NOUSEROVERRIDE,
                        vt
                    );
    }

    return hres;
}

template<typename T>
void
_TBCGPVariantOperator(const _variant_t& var, T& val)
{
	val = (T)var;
}

template<> void _TBCGPVariantOperator<bool>(const _variant_t& var, bool& val)
{
	val = V_BOOL(&var) == BCGP_VARIANT_TRUE ? true : false;
}


#define BCGP_IMPLEMENT_VARIANTOPERATOR(type, op) \
template<> void _TBCGPVariantOperator<type>(const _variant_t& var, type& val) \
{ \
	val = op(&var); \
} \

BCGP_IMPLEMENT_VARIANTOPERATOR(char, V_I1)
BCGP_IMPLEMENT_VARIANTOPERATOR(signed char, V_I1)
BCGP_IMPLEMENT_VARIANTOPERATOR(unsigned char, V_UI1)
BCGP_IMPLEMENT_VARIANTOPERATOR(signed short, V_I2)
BCGP_IMPLEMENT_VARIANTOPERATOR(unsigned short, V_UI2)
BCGP_IMPLEMENT_VARIANTOPERATOR(signed long, V_I4)
BCGP_IMPLEMENT_VARIANTOPERATOR(unsigned long, V_UI4)
BCGP_IMPLEMENT_VARIANTOPERATOR(signed int, V_I4)
BCGP_IMPLEMENT_VARIANTOPERATOR(unsigned int, V_UI4)
#if (_WIN32_WINNT >= 0x0501)
	BCGP_IMPLEMENT_VARIANTOPERATOR(signed __int64, V_I8)
	BCGP_IMPLEMENT_VARIANTOPERATOR(unsigned __int64, V_UI8)
#endif
BCGP_IMPLEMENT_VARIANTOPERATOR(float, V_R4)
BCGP_IMPLEMENT_VARIANTOPERATOR(double, V_R8)
BCGP_IMPLEMENT_VARIANTOPERATOR(long double, V_R8)
BCGP_IMPLEMENT_VARIANTOPERATOR(BSTR, V_BSTR)


template<typename T>
HRESULT
_TBCGPVariantGetValue(const _variant_t& var, T& val)
{
    if (!_BCGPVariantIsValid(var))
    {
        return DISP_E_BADVARTYPE;
    }

    const bool text = _TBCGPVariantTypeBSTR<T>::value;

    const VARTYPE vtSrc = V_VT(&var);
    const bool languageSpecific = text || _BCGPVariantIsLanguageSpecific(vtSrc);
    const VARTYPE vtDst = text ? VT_BSTR : _TBCGPVariantTypeInfo<T>::value;

    HRESULT hresult = S_OK;

    if (vtSrc == vtDst || !languageSpecific)
    {
        _TBCGPVariantOperator<T>(var, val);
    }
    else
    {
        _variant_t varDst;

        hresult = _BCGPVariantChangeTypeEx(vtDst, varDst, &var);
        if (SUCCEEDED(hresult))
        {
            _TBCGPVariantOperator<T>(varDst, val);
        }
    }

    return hresult;
}

template<>
HRESULT
_TBCGPVariantGetValue<CString>(const _variant_t& var, CString& val)
{
    BSTR bstr = NULL;

    HRESULT hresult = _TBCGPVariantGetValue<BSTR>(var, bstr);
    if (SUCCEEDED(hresult))
    {
        val = (LPCTSTR)_bstr_t(bstr);
    }

    return hresult;
}

template<>
HRESULT
_TBCGPVariantGetValue<COleDateTime>(const _variant_t& var, COleDateTime& val)
{
    double dbl = 0.0;

    HRESULT hresult = _TBCGPVariantGetValue<double>(var, dbl);
    if (SUCCEEDED(hresult))
    {
        val = COleDateTime(dbl);
    }

    return hresult;
}


#define BCGP_IMPLEMENT_GETVARIANTVALUE(type) \
bool BCGPGetVariantValue(const _variant_t& var, type& value) \
{ \
	return SUCCEEDED(_TBCGPVariantGetValue<type>(var, value)); \
} \

BCGP_IMPLEMENT_GETVARIANTVALUE(bool);
BCGP_IMPLEMENT_GETVARIANTVALUE(char);
BCGP_IMPLEMENT_GETVARIANTVALUE(signed char);
BCGP_IMPLEMENT_GETVARIANTVALUE(unsigned char);
BCGP_IMPLEMENT_GETVARIANTVALUE(signed short);
BCGP_IMPLEMENT_GETVARIANTVALUE(unsigned short);
BCGP_IMPLEMENT_GETVARIANTVALUE(signed long);
BCGP_IMPLEMENT_GETVARIANTVALUE(unsigned long);
BCGP_IMPLEMENT_GETVARIANTVALUE(signed int);
BCGP_IMPLEMENT_GETVARIANTVALUE(unsigned int);
#if (_WIN32_WINNT >= 0x0501)
	BCGP_IMPLEMENT_GETVARIANTVALUE(signed __int64);
	BCGP_IMPLEMENT_GETVARIANTVALUE(unsigned __int64);
#endif
BCGP_IMPLEMENT_GETVARIANTVALUE(float);
BCGP_IMPLEMENT_GETVARIANTVALUE(double);
BCGP_IMPLEMENT_GETVARIANTVALUE(long double);
BCGP_IMPLEMENT_GETVARIANTVALUE(CString);
BCGP_IMPLEMENT_GETVARIANTVALUE(COleDateTime);


CBCGPXmlNode::CBCGPXmlNode(REFIID iid)
	: CBCGPXmlBase(iid)
{
}

CBCGPXmlNode::CBCGPXmlNode()
	: CBCGPXmlBase(__uuidof(BCG_MSXML::IXMLDOMNode))
{
}

CBCGPXmlNode::CBCGPXmlNode(const CBCGPXmlNode& other)
	: CBCGPXmlBase(other)
{
}

CBCGPXmlNode::~CBCGPXmlNode()
{
}

BCGPDOMNodeType CBCGPXmlNode::GetNodeType()
{
	BCG_MSXML::DOMNodeType value = (BCG_MSXML::DOMNodeType)BCG_MSXML::NODE_INVALID;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_nodeType(&value)
				: E_POINTER
		);

	return (BCGPDOMNodeType)value;
}

CString CBCGPXmlNode::GetNodeTypeString()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_nodeTypeString(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

CString CBCGPXmlNode::GetNodeName()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_nodeName(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

_variant_t CBCGPXmlNode::GetNodeTypedValue()
{
	_variant_t value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_nodeTypedValue(&value)
				: E_POINTER
		);

	return value;
}

BOOL CBCGPXmlNode::SetNodeTypedValue(const _variant_t& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->put_nodeTypedValue(value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CBCGPXmlNode CBCGPXmlNode::GetParentNode()
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_parentNode(BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

CString CBCGPXmlNode::GetBaseName()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_baseName(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

_variant_t CBCGPXmlNode::GetDataType()
{
	_variant_t value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_dataType(&value)
				: E_POINTER
		);

	return value;
}

BOOL CBCGPXmlNode::SetDataType(const CString& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->put_dataType(_bstr_t(value))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

_variant_t CBCGPXmlNode::GetNodeValue()
{
	_variant_t value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_nodeValue(&value)
				: E_POINTER
		);

	return value;
}

BOOL CBCGPXmlNode::SetNodeValue(const _variant_t& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->put_nodeValue(value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CString CBCGPXmlNode::GetNamespaceURI()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_namespaceURI(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

CString CBCGPXmlNode::GetPrefix()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_prefix(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

BOOL CBCGPXmlNode::IsParsed()
{
	VARIANT_BOOL value = BCGP_VARIANT_FALSE;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_parsed(&value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult()) && value == BCGP_VARIANT_TRUE;
}

BOOL CBCGPXmlNode::IsSpecified()
{
	VARIANT_BOOL value = BCGP_VARIANT_FALSE;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_specified(&value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult()) && value == BCGP_VARIANT_TRUE;
}

CString CBCGPXmlNode::GetText()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_text(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

BOOL CBCGPXmlNode::SetText(const CString& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->put_text(_bstr_t(value))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlNode::GetXML(CString& value)
{
	BSTR bstr = NULL;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_xml(&bstr)
				: E_POINTER
		);

	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CBCGPXmlNode CBCGPXmlNode::GetDefinition()
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_definition(BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

BOOL CBCGPXmlNode::HasChildren()
{
	VARIANT_BOOL value = BCGP_VARIANT_FALSE;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->hasChildNodes(&value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult()) && value == BCGP_VARIANT_TRUE;
}

BOOL CBCGPXmlNode::GetOwnerDocument(CBCGPXmlDocument& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_ownerDocument(BCG_MSXML_DOMDOCUMENT_PTR(value))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CBCGPXmlNode CBCGPXmlNode::CloneNode(BOOL deep)
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->cloneNode
							(
								deep ? BCGP_VARIANT_TRUE : BCGP_VARIANT_FALSE, 
								BCG_MSXML_DOMNODE_PTR(value)
							)
				: E_POINTER
		);

	return value;
}

CBCGPXmlNode CBCGPXmlNode::GetChildFirst()
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_firstChild(BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNode CBCGPXmlNode::GetChildLast()
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_lastChild(BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNode CBCGPXmlNode::GetSiblingNext()
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_nextSibling(BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNode CBCGPXmlNode::GetSiblingPrev()
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_previousSibling(BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

BOOL CBCGPXmlNode::AppendChild(const CBCGPXmlNode& newValue, CBCGPXmlNode* outValue)
{
	CBCGPXmlNode out;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->appendChild(BCG_MSXML_DOMNODE(const_cast<CBCGPXmlNode&>(newValue)), BCG_MSXML_DOMNODE_PTR(out))
				: E_POINTER
		);

	if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
	{
		if (outValue != NULL)
		{
			*outValue = out;
		}

		return TRUE;
	}

	return FALSE;
}

BOOL CBCGPXmlNode::RemoveChild(const CBCGPXmlNode& value, CBCGPXmlNode* outValue)
{
	CBCGPXmlNode out;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->removeChild(BCG_MSXML_DOMNODE(const_cast<CBCGPXmlNode&>(value)), BCG_MSXML_DOMNODE_PTR(out))
				: E_POINTER
		);

	if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
	{
		if (outValue != NULL)
		{
			*outValue = out;
		}

		return TRUE;
	}

	return FALSE;
}

BOOL CBCGPXmlNode::ReplaceChild(const CBCGPXmlNode& newValue, const CBCGPXmlNode& oldValue, CBCGPXmlNode* outValue)
{
	CBCGPXmlNode out;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->replaceChild
							(
								BCG_MSXML_DOMNODE(const_cast<CBCGPXmlNode&>(newValue)), 
								BCG_MSXML_DOMNODE(const_cast<CBCGPXmlNode&>(oldValue)), 
								BCG_MSXML_DOMNODE_PTR(out)
							)
				: E_POINTER
		);

	if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
	{
		if (outValue != NULL)
		{
			*outValue = out;
		}

		return TRUE;
	}

	return FALSE;
}

CBCGPXmlNodeList CBCGPXmlNode::GetChildren()
{
	CBCGPXmlNodeList value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_childNodes(BCG_MSXML_DOMNODELIST_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNamedNodeMap CBCGPXmlNode::GetAttributes()
{
	CBCGPXmlNamedNodeMap value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->get_attributes(BCG_MSXML_DOMNAMEDNODEMAP_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNode CBCGPXmlNode::SelectNode(const CString& query)
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->selectSingleNode(_bstr_t(query), BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNodeList CBCGPXmlNode::SelectNodes(const CString& query)
{
	CBCGPXmlNodeList value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODE(*this)->selectNodes(_bstr_t(query), BCG_MSXML_DOMNODELIST_PTR(value))
				: E_POINTER
		);

	return value;
}

CString CBCGPXmlNode::TransformNode(const CBCGPXmlNode& stylesheet)
{
	CString value;
	HRESULT result = E_INVALIDARG;

	if (stylesheet.IsValid())
	{
		BSTR bstr = NULL;
		result = IsValid()
					? BCG_MSXML_DOMNODE(*this)->transformNode(BCG_MSXML_DOMNODE(const_cast<CBCGPXmlNode&>(stylesheet)), &bstr)
					: E_POINTER;

		if (bstr != NULL)
		{
			if (BCGP_HR_SUCCEEDED_OK(result))
			{
				value = bstr;
			}

			::SysFreeString(bstr);
		}
	}

	SetLastResult(result);

	return value;
}

BOOL CBCGPXmlNode::TransformNodeToObject(const CBCGPXmlNode& stylesheet, _variant_t& value)
{
	HRESULT result = E_INVALIDARG;

	if (stylesheet.IsValid())
	{
		result = IsValid()
					? BCG_MSXML_DOMNODE(*this)->transformNodeToObject(BCG_MSXML_DOMNODE(const_cast<CBCGPXmlNode&>(stylesheet)), value)
					: E_POINTER;
	}

	SetLastResult(result);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlNode::TransformNodeToDocument(const CBCGPXmlNode& stylesheet, CBCGPXmlDocument& document)
{
	HRESULT result = E_INVALIDARG;

	if (document.IsValid())
	{
		ATL::CComPtr<IDispatch> dispatch;
		result = document.GetInterface()->QueryInterface(IID_IDispatch, (void**)&dispatch);

		if (BCGP_HR_SUCCEEDED_OK(result))
		{
			_variant_t var(dispatch, TRUE);

			TransformNodeToObject(stylesheet, var);
			result = GetLastResult();
		}
	}

	SetLastResult(result);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlNode::TransformNodeToStream(const CBCGPXmlNode& stylesheet, IStream* stream)
{
	_variant_t var(stream, TRUE);

	SetLastResult
		(
			stream != NULL
				? TransformNodeToObject(stylesheet, var)
				: E_INVALIDARG
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}


CBCGPXmlCharacterData::CBCGPXmlCharacterData(REFIID iid)
	: CBCGPXmlNode(iid)
{
}

CBCGPXmlCharacterData::CBCGPXmlCharacterData()
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMCharacterData))
{
}

CBCGPXmlCharacterData::CBCGPXmlCharacterData(const CBCGPXmlNode& other)
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMCharacterData))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlCharacterData::CBCGPXmlCharacterData(const CBCGPXmlCharacterData& other)
	: CBCGPXmlNode(other)
{
}

CBCGPXmlCharacterData::~CBCGPXmlCharacterData()
{
}

long CBCGPXmlCharacterData::GetLength()
{
	long value = 0;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMCHARACTERDATA(*this)->get_length(&value)
				: E_POINTER
		);

	return value;
}

CString CBCGPXmlCharacterData::GetData()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMCHARACTERDATA(*this)->get_data(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

BOOL CBCGPXmlCharacterData::SetData(const CString& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMCHARACTERDATA(*this)->put_data(_bstr_t(value))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlCharacterData::AppendData(const CString& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMCHARACTERDATA(*this)->appendData(_bstr_t(value))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlCharacterData::InsertData(long offset, const CString& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMCHARACTERDATA(*this)->insertData(offset, _bstr_t(value))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlCharacterData::DeleteData(long offset, long count)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMCHARACTERDATA(*this)->deleteData(offset, count)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlCharacterData::ReplaceData(long offset, long count, const CString& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMCHARACTERDATA(*this)->replaceData(offset, count, _bstr_t(value))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CString CBCGPXmlCharacterData::SubstringData(long offset, long count)
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMCHARACTERDATA(*this)->substringData(offset, count, &bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}


CBCGPXmlText::CBCGPXmlText(REFIID iid)
	: CBCGPXmlCharacterData(iid)
{
}

CBCGPXmlText::CBCGPXmlText(const CBCGPXmlNode& other)
	: CBCGPXmlCharacterData(__uuidof(BCG_MSXML::IXMLDOMText))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlText::CBCGPXmlText(const CBCGPXmlCharacterData& other)
	: CBCGPXmlCharacterData(__uuidof(BCG_MSXML::IXMLDOMText))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlText::CBCGPXmlText()
	: CBCGPXmlCharacterData(__uuidof(BCG_MSXML::IXMLDOMText))
{
}

CBCGPXmlText::CBCGPXmlText(const CBCGPXmlText& other)
	: CBCGPXmlCharacterData(other)
{
}

CBCGPXmlText::~CBCGPXmlText()
{
}

BOOL CBCGPXmlText::SplitText(long offset, CBCGPXmlText& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMTEXT(*this)->splitText(offset, BCG_MSXML_DOMTEXT_PTR(value))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}


CBCGPXmlComment::CBCGPXmlComment()
	: CBCGPXmlCharacterData(__uuidof(BCG_MSXML::IXMLDOMComment))
{
}

CBCGPXmlComment::CBCGPXmlComment(const CBCGPXmlNode& other)
	: CBCGPXmlCharacterData(__uuidof(BCG_MSXML::IXMLDOMComment))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlComment::CBCGPXmlComment(const CBCGPXmlCharacterData& other)
	: CBCGPXmlCharacterData(__uuidof(BCG_MSXML::IXMLDOMComment))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlComment::CBCGPXmlComment(const CBCGPXmlComment& other)
	: CBCGPXmlCharacterData(other)
{
}

CBCGPXmlComment::~CBCGPXmlComment()
{
}


CBCGPXmlCDATASection::CBCGPXmlCDATASection()
	: CBCGPXmlText(__uuidof(BCG_MSXML::IXMLDOMCDATASection))
{
}

CBCGPXmlCDATASection::CBCGPXmlCDATASection(const CBCGPXmlNode& other)
	: CBCGPXmlText(__uuidof(BCG_MSXML::IXMLDOMCDATASection))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlCDATASection::CBCGPXmlCDATASection(const CBCGPXmlCharacterData& other)
	: CBCGPXmlText(__uuidof(BCG_MSXML::IXMLDOMCDATASection))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlCDATASection::CBCGPXmlCDATASection(const CBCGPXmlText& other)
	: CBCGPXmlText(__uuidof(BCG_MSXML::IXMLDOMCDATASection))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlCDATASection::CBCGPXmlCDATASection(const CBCGPXmlCDATASection& other)
	: CBCGPXmlText(other)
{
}

CBCGPXmlCDATASection::~CBCGPXmlCDATASection()
{
}


CBCGPXmlAttribute::CBCGPXmlAttribute()
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMAttribute))
{
}

CBCGPXmlAttribute::CBCGPXmlAttribute(const CBCGPXmlNode& other)
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMAttribute))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlAttribute::CBCGPXmlAttribute(const CBCGPXmlAttribute& other)
	: CBCGPXmlNode(other)
{
}

CBCGPXmlAttribute::~CBCGPXmlAttribute()
{
}

CString CBCGPXmlAttribute::GetName()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMATTRIBUTE(*this)->get_name(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

_variant_t CBCGPXmlAttribute::GetValue()
{
	_variant_t value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMATTRIBUTE(*this)->get_value(&value)
				: E_POINTER
		);

	return value;
}

BOOL CBCGPXmlAttribute::SetValue(const _variant_t& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMATTRIBUTE(*this)->put_value(value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

#define BCGP_IMPLEMENT_XMLATTRIBUTE_OP(type) \
CBCGPXmlAttribute::operator type() \
{ \
	type value; \
	BCGPGetVariantValue(GetValue(), value); \
	return value; \
} \

BCGP_IMPLEMENT_XMLATTRIBUTE_OP(bool);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(char);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(signed char);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(unsigned char);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(signed short);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(unsigned short);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(signed long);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(unsigned long);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(signed int);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(unsigned int);
#if (_WIN32_WINNT >= 0x0501)
	BCGP_IMPLEMENT_XMLATTRIBUTE_OP(signed __int64);
	BCGP_IMPLEMENT_XMLATTRIBUTE_OP(unsigned __int64);
#endif
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(float);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(double);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(long double);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(CString);
BCGP_IMPLEMENT_XMLATTRIBUTE_OP(COleDateTime);



CBCGPXmlElement::CBCGPXmlElement()
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMElement))
{
}

CBCGPXmlElement::CBCGPXmlElement(const CBCGPXmlNode& other)
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMElement))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlElement::CBCGPXmlElement(const CBCGPXmlElement& other)
	: CBCGPXmlNode(other)
{
}

CBCGPXmlElement::~CBCGPXmlElement()
{
}

CString CBCGPXmlElement::GetTagName()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMELEMENT(*this)->get_tagName(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

CBCGPXmlNodeList CBCGPXmlElement::GetElementsByTagName(const CString& strName)
{
	CBCGPXmlNodeList value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMELEMENT(*this)->getElementsByTagName(_bstr_t(strName), BCG_MSXML_DOMNODELIST_PTR(value))
				: E_POINTER
		);

	return value;
}

_variant_t CBCGPXmlElement::GetAttribute(const CString& strName)
{
	_variant_t value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMELEMENT(*this)->getAttribute(_bstr_t(strName), &value)
				: E_POINTER
		);

	return value;
}


#define BCGP_IMPLEMENT_XMLELEMENT_VALUE(type) \
bool CBCGPXmlElement::GetAttributeValue(const CString& strName, type& value) \
{ \
	return BCGPGetVariantValue(GetAttribute(strName), value); \
} \

BCGP_IMPLEMENT_XMLELEMENT_VALUE(bool);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(char);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(signed char);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(unsigned char);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(signed short);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(unsigned short);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(signed long);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(unsigned long);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(signed int);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(unsigned int);
#if (_WIN32_WINNT >= 0x0501)
	BCGP_IMPLEMENT_XMLELEMENT_VALUE(signed __int64);
	BCGP_IMPLEMENT_XMLELEMENT_VALUE(unsigned __int64);
#endif
BCGP_IMPLEMENT_XMLELEMENT_VALUE(float);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(double);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(long double);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(CString);
BCGP_IMPLEMENT_XMLELEMENT_VALUE(COleDateTime);


BOOL CBCGPXmlElement::GetBoolAttribute(const CString& strName)
{
	bool value = false;
	GetAttributeValue(strName, value);

	return value;
}

long CBCGPXmlElement::GetIntAttribute(const CString& strName)
{
	long value = 0;
	GetAttributeValue(strName, value);

	return value;
}

CString CBCGPXmlElement::GetStringAttribute(const CString& strName)
{
	CString value;
	GetAttributeValue(strName, value);

	return value;
}

void CBCGPXmlElement::GetRectangleAttribute(CRect& rect)
{
	rect.top = GetIntAttribute(_T("top"));
	rect.bottom = GetIntAttribute(_T("bottom"));
	rect.left = GetIntAttribute(_T("left"));
	rect.right = GetIntAttribute(_T("right"));
}

void CBCGPXmlElement::SetRectangleAttribute(const CRect& rect)
{
	SetAttribute(_T("top"), rect.top);
	SetAttribute(_T("left"),rect.left);
	SetAttribute(_T("bottom"), rect.bottom);
	SetAttribute(_T("right"), rect.right);
}

BOOL CBCGPXmlElement::SetAttribute(const CString& strName, const _variant_t& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMELEMENT(*this)->setAttribute(_bstr_t(strName), value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlElement::RemoveAttribute(const CString& strName)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMELEMENT(*this)->removeAttribute(_bstr_t(strName))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CBCGPXmlAttribute CBCGPXmlElement::GetAttributeNode(const CString& strName)
{
	CBCGPXmlAttribute value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMELEMENT(*this)->getAttributeNode(_bstr_t(strName), BCG_MSXML_DOMATTRIBUTE_PTR(value))
				: E_POINTER
		);

	return value;
}

BOOL CBCGPXmlElement::SetAttributeNode(const CBCGPXmlAttribute& value, CBCGPXmlAttribute* outValue)
{
	CBCGPXmlAttribute out;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMELEMENT(*this)->setAttributeNode(BCG_MSXML_DOMATTRIBUTE(const_cast<CBCGPXmlAttribute&>(value)), BCG_MSXML_DOMATTRIBUTE_PTR(out))
				: E_POINTER
		);

	if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
	{
		if (outValue != NULL)
		{
			*outValue = out;
		}

		return TRUE;
	}

	return FALSE;
}

BOOL CBCGPXmlElement::RemoveAttributeNode(const CBCGPXmlAttribute& value, CBCGPXmlAttribute* outValue)
{
	CBCGPXmlAttribute out;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMELEMENT(*this)->removeAttributeNode(BCG_MSXML_DOMATTRIBUTE(const_cast<CBCGPXmlAttribute&>(value)), BCG_MSXML_DOMATTRIBUTE_PTR(out))
				: E_POINTER
		);

	if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
	{
		if (outValue != NULL)
		{
			*outValue = out;
		}

		return TRUE;
	}

	return FALSE;
}

BOOL CBCGPXmlElement::Normalize()
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMELEMENT(*this)->normalize()
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CBCGPXmlElement CBCGPXmlElement::FindChildElement(const CString& strName)
{
	CBCGPXmlNodeList subElementList(GetChildren());

	if (subElementList.IsValid())
	{
		const int nCount = subElementList.GetLength();
		for (int i = 0; i < nCount; i++)
		{
			CBCGPXmlElement element(subElementList.GetItem(i));
			if(element.IsValid() && element.GetTagName() == strName)
			{
				return element;
			}
		}
	}

	return CBCGPXmlElement();
}


CBCGPXmlEntity::CBCGPXmlEntity()
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMEntity))
{
}

CBCGPXmlEntity::CBCGPXmlEntity(const CBCGPXmlNode& other)
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMEntity))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlEntity::CBCGPXmlEntity(const CBCGPXmlEntity& other)
	: CBCGPXmlNode(other)
{
}

CBCGPXmlEntity::~CBCGPXmlEntity()
{
}

CString CBCGPXmlEntity::GetNotationName()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMENTITY(*this)->get_notationName(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

_variant_t CBCGPXmlEntity::GetPublicId()
{
	_variant_t value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMENTITY(*this)->get_publicId(&value)
				: E_POINTER
		);

	return value;
}

_variant_t CBCGPXmlEntity::GetSystemId()
{
	_variant_t value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMENTITY(*this)->get_systemId(&value)
				: E_POINTER
		);

	return value;
}


CBCGPXmlEntityReference::CBCGPXmlEntityReference()
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMEntityReference))
{
}

CBCGPXmlEntityReference::CBCGPXmlEntityReference(const CBCGPXmlNode& other)
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMEntityReference))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlEntityReference::CBCGPXmlEntityReference(const CBCGPXmlEntityReference& other)
	: CBCGPXmlNode(other)
{
}

CBCGPXmlEntityReference::~CBCGPXmlEntityReference()
{
}


CBCGPXmlProcessingInstruction::CBCGPXmlProcessingInstruction()
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMProcessingInstruction))
{
}

CBCGPXmlProcessingInstruction::CBCGPXmlProcessingInstruction(const CBCGPXmlNode& other)
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMProcessingInstruction))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlProcessingInstruction::CBCGPXmlProcessingInstruction(const CBCGPXmlProcessingInstruction& other)
	: CBCGPXmlNode(other)
{
}

CBCGPXmlProcessingInstruction::~CBCGPXmlProcessingInstruction()
{
}

CString CBCGPXmlProcessingInstruction::GetTarget()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMPROCESSINGINSTRUCTION(*this)->get_target(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

CString CBCGPXmlProcessingInstruction::GetData()
{
	BSTR bstr = NULL;
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMPROCESSINGINSTRUCTION(*this)->get_data(&bstr)
				: E_POINTER
		);

	CString value;
	if (bstr != NULL)
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			value = bstr;
		}

		::SysFreeString(bstr);
	}

	return value;
}

BOOL CBCGPXmlProcessingInstruction::SetData(const CString& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMPROCESSINGINSTRUCTION(*this)->put_data(_bstr_t(value))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CString CBCGPXmlProcessingInstruction::GetEncoding()
{
	CString encoding;
	HRESULT result = E_POINTER;

	if (IsValid())
	{
		CBCGPXmlNamedNodeMap attributes(GetAttributes());

		if (attributes.IsValid())
		{
			CBCGPXmlNode attribute(attributes.GetItem(c_szXml_Encoding));
			result = attributes.GetLastResult();

			if (attribute.IsValid())
			{
				encoding = attribute.GetText();
				result = attribute.GetLastResult();
			}
		}
	}

	SetLastResult(result);

	return encoding;
}

BOOL CBCGPXmlProcessingInstruction::SetEncoding(LPCTSTR encoding)
{
	if (BCGPIsStringEmpty(encoding))
	{
		SetLastResult(E_INVALIDARG);
		return FALSE;
	}

	BOOL bRes = FALSE;
	HRESULT result = E_POINTER;

	if (IsValid())
	{
		CBCGPXmlNamedNodeMap attributes(GetAttributes());

		if (attributes.IsValid())
		{
			CBCGPXmlNode attribute(attributes.GetItem(c_szXml_Encoding));
			result = attributes.GetLastResult();

			if (attribute.IsValid())
			{
				bRes = attribute.SetText(encoding);
				result = attribute.GetLastResult();
			}
		}
	}

	SetLastResult(result);

	return bRes;
}


CBCGPXmlNodeList::CBCGPXmlNodeList()
	: CBCGPXmlBase(__uuidof(BCG_MSXML::IXMLDOMNodeList))
{
}

CBCGPXmlNodeList::CBCGPXmlNodeList(const CBCGPXmlNodeList& other)
	: CBCGPXmlBase(other)
{
}

CBCGPXmlNodeList::~CBCGPXmlNodeList()
{
}

long CBCGPXmlNodeList::GetLength()
{
	long value = 0;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODELIST(*this)->get_length(&value)
				: E_POINTER
		);

	return value;
}

void CBCGPXmlNodeList::Reset()
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODELIST(*this)->reset()
				: E_POINTER
		);
}

CBCGPXmlNode CBCGPXmlNodeList::GetItem(long index)
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODELIST(*this)->get_item(index, BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNode CBCGPXmlNodeList::NextNode()
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNODELIST(*this)->nextNode(BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}


CBCGPXmlNamedNodeMap::CBCGPXmlNamedNodeMap()
	: CBCGPXmlBase(__uuidof(BCG_MSXML::IXMLDOMNamedNodeMap))
{
}

CBCGPXmlNamedNodeMap::CBCGPXmlNamedNodeMap(const CBCGPXmlNamedNodeMap& other)
	: CBCGPXmlBase(other)
{
}

CBCGPXmlNamedNodeMap::~CBCGPXmlNamedNodeMap()
{
}

long CBCGPXmlNamedNodeMap::GetLength()
{
	long value = 0;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNAMEDNODEMAP(*this)->get_length(&value)
				: E_POINTER
		);

	return value;
}

void CBCGPXmlNamedNodeMap::Reset()
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNAMEDNODEMAP(*this)->reset()
				: E_POINTER
		);
}

CBCGPXmlNode CBCGPXmlNamedNodeMap::GetItem(long index)
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNAMEDNODEMAP(*this)->get_item(index, BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNode CBCGPXmlNamedNodeMap::NextNode()
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMNAMEDNODEMAP(*this)->nextNode(BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNode CBCGPXmlNamedNodeMap::GetItem(const CString& strName, LPCTSTR namespaceURI)
{
	CBCGPXmlNode value;

	if (BCGPIsStringEmpty(namespaceURI))
	{
		SetLastResult
			(
				IsValid()
					? BCG_MSXML_DOMNAMEDNODEMAP(*this)->getNamedItem(_bstr_t(strName), BCG_MSXML_DOMNODE_PTR(value))
					: E_POINTER
			);
	}
	else
	{
		SetLastResult
			(
				IsValid()
					? BCG_MSXML_DOMNAMEDNODEMAP(*this)->getQualifiedItem(_bstr_t(strName), _bstr_t(namespaceURI), BCG_MSXML_DOMNODE_PTR(value))
					: E_POINTER
			);
	}

	return value;
}

BOOL CBCGPXmlNamedNodeMap::RemoveItem(LPCTSTR strName, LPCTSTR namespaceURI, CBCGPXmlNode* outValue)
{
	CBCGPXmlNode value;

	if (BCGPIsStringEmpty(namespaceURI))
	{
		SetLastResult
			(
				IsValid()
					? BCG_MSXML_DOMNAMEDNODEMAP(*this)->removeNamedItem(_bstr_t(strName), BCG_MSXML_DOMNODE_PTR(value))
					: E_POINTER
			);
	}
	else
	{
		SetLastResult
			(
				IsValid()
					? BCG_MSXML_DOMNAMEDNODEMAP(*this)->removeQualifiedItem(_bstr_t(strName), _bstr_t(namespaceURI), BCG_MSXML_DOMNODE_PTR(value))
					: E_POINTER
			);
	}

	if (outValue != NULL)
	{
		*outValue = value;
	}

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlNamedNodeMap::SetItem(const CBCGPXmlNode& newValue, CBCGPXmlNode* outValue)
{
    CBCGPXmlNode value;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMNAMEDNODEMAP(*this)->setNamedItem(BCG_MSXML_DOMNODE(const_cast<CBCGPXmlNode&>(newValue)), BCG_MSXML_DOMNODE_PTR(value))
                : E_POINTER
        );

    if (outValue != NULL)
    {
        *outValue = value;
    }

    return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}


CBCGPXmlParseErrorCollection::CBCGPXmlParseErrorCollection()
    : CBCGPXmlBase(__uuidof(BCG_MSXML::IXMLDOMParseErrorCollection))
{
}

CBCGPXmlParseErrorCollection::CBCGPXmlParseErrorCollection(const CBCGPXmlParseErrorCollection& other)
    : CBCGPXmlBase(other)
{
}

CBCGPXmlParseErrorCollection::~CBCGPXmlParseErrorCollection()
{
}

long CBCGPXmlParseErrorCollection::GetLength()
{
	long value = 0;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMPARSEERRORCOLLECTION(*this)->get_length(&value)
				: E_POINTER
		);

	return value;
}

void CBCGPXmlParseErrorCollection::Reset()
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMPARSEERRORCOLLECTION(*this)->reset()
				: E_POINTER
		);
}

CBCGPXmlParseError CBCGPXmlParseErrorCollection::GetItem(long index)
{
	CBCGPXmlBase error(__uuidof(BCG_MSXML::IXMLDOMParseError2));

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMPARSEERRORCOLLECTION(*this)->get_item(index, (BCG_MSXML::IXMLDOMParseError2**)error.GetInterfacePtr())
				: E_POINTER
		);

	CBCGPXmlParseError value;
	if (BCGP_HR_SUCCEEDED_OK(GetLastResult()) && error.IsValid())
	{
		(CBCGPXmlBase&)value = error;
	}

	return value;
}

CBCGPXmlParseError CBCGPXmlParseErrorCollection::Next()
{
	CBCGPXmlBase error(__uuidof(BCG_MSXML::IXMLDOMParseError2));

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMPARSEERRORCOLLECTION(*this)->get_next((BCG_MSXML::IXMLDOMParseError2**)error.GetInterfacePtr())
				: E_POINTER
		);

	CBCGPXmlParseError value;
	if (BCGP_HR_SUCCEEDED_OK(GetLastResult()) && error.IsValid())
	{
		(CBCGPXmlBase&)value = error;
	}

	return value;
}


CBCGPXmlImplemetation::CBCGPXmlImplemetation()
    : CBCGPXmlBase(__uuidof(BCG_MSXML::IXMLDOMImplementation))
{
}

CBCGPXmlImplemetation::CBCGPXmlImplemetation(const CBCGPXmlImplemetation& other)
    : CBCGPXmlBase(other)
{
}

CBCGPXmlImplemetation::~CBCGPXmlImplemetation()
{
}

BOOL CBCGPXmlImplemetation::HasFeature(const CString& feature, LPCTSTR version)
{
    VARIANT_BOOL value = BCGP_VARIANT_FALSE;

    SetLastResult
        (
            IsValid()
            ? BCG_MSXML_DOMIMPLEMENTATION(*this)->hasFeature(_bstr_t(feature), _bstr_t(version), &value)
                : E_POINTER
        );

    return BCGP_HR_SUCCEEDED_OK(GetLastResult()) && value == BCGP_VARIANT_TRUE;
}


CBCGPXmlNotation::CBCGPXmlNotation()
    : CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMNotation))
{
}

CBCGPXmlNotation::CBCGPXmlNotation(const CBCGPXmlNode& other)
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMNotation))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlNotation::CBCGPXmlNotation(const CBCGPXmlNotation& other)
    : CBCGPXmlNode(other)
{
}

CBCGPXmlNotation::~CBCGPXmlNotation()
{
}

_variant_t CBCGPXmlNotation::GetPublicId()
{
    _variant_t value;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMNOTATION(*this)->get_publicId(&value)
                : E_POINTER
        );

    return value;
}

_variant_t CBCGPXmlNotation::GetSystemId()
{
    _variant_t value;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMNOTATION(*this)->get_systemId(&value)
                : E_POINTER
        );

    return value;
}


CBCGPXmlDocumentFragment::CBCGPXmlDocumentFragment()
    : CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMDocumentFragment))
{
}

CBCGPXmlDocumentFragment::CBCGPXmlDocumentFragment(const CBCGPXmlNode& other)
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMDocumentFragment))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlDocumentFragment::CBCGPXmlDocumentFragment(const CBCGPXmlDocumentFragment& other)
    : CBCGPXmlNode(other)
{
}

CBCGPXmlDocumentFragment::~CBCGPXmlDocumentFragment()
{
}


CBCGPXmlDocumentType::CBCGPXmlDocumentType()
    : CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMDocumentType))
{
}

CBCGPXmlDocumentType::CBCGPXmlDocumentType(const CBCGPXmlNode& other)
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMDocumentType))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlDocumentType::CBCGPXmlDocumentType(const CBCGPXmlDocumentType& other)
    : CBCGPXmlNode(other)
{
}

CBCGPXmlDocumentType::~CBCGPXmlDocumentType()
{
}

CString CBCGPXmlDocumentType::GetName()
{
    BSTR bstr = NULL;
    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMDOCUMENTTYPE(*this)->get_name(&bstr)
                : E_POINTER
        );

    CString value;
    if (bstr != NULL)
    {
        if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
        {
            value = bstr;
        }

        ::SysFreeString(bstr);
    }

    return value;
}

CBCGPXmlNamedNodeMap CBCGPXmlDocumentType::GetEntities()
{
	CBCGPXmlNamedNodeMap value;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMDOCUMENTTYPE(*this)->get_entities(BCG_MSXML_DOMNAMEDNODEMAP_PTR(value))
                : E_POINTER
        );

    return value;
}

CBCGPXmlNamedNodeMap CBCGPXmlDocumentType::GetNotations()
{
	CBCGPXmlNamedNodeMap value;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMDOCUMENTTYPE(*this)->get_notations(BCG_MSXML_DOMNAMEDNODEMAP_PTR(value))
                : E_POINTER
        );

    return value;
}


CBCGPXmlParseError::CBCGPXmlParseError()
    : CBCGPXmlBase(__uuidof(BCG_MSXML::IXMLDOMParseError))
{
}

CBCGPXmlParseError::CBCGPXmlParseError(const CBCGPXmlParseError& other)
    : CBCGPXmlBase(other)
{
}

CBCGPXmlParseError::~CBCGPXmlParseError()
{
}

long CBCGPXmlParseError::GerErrorCode()
{
    long value = -1;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMPARSEERROR(*this)->get_errorCode(&value)
                : E_POINTER
        );

    return value;
}

long CBCGPXmlParseError::GetFilePos()
{
    long value = -1;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMPARSEERROR(*this)->get_filepos(&value)
                : E_POINTER
        );

    return value;
}

long CBCGPXmlParseError::GetLine()
{
    long value = -1;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMPARSEERROR(*this)->get_line(&value)
                : E_POINTER
        );

    return value;
}

long CBCGPXmlParseError::GetLinePos()
{
    long value = -1;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMPARSEERROR(*this)->get_linepos(&value)
                : E_POINTER
        );

    return value;
}

CString CBCGPXmlParseError::GetReason()
{
    BSTR bstr = NULL;
    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMPARSEERROR(*this)->get_reason(&bstr)
                : E_POINTER
        );

    CString value;
    if (bstr != NULL)
    {
        if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
        {
            value = bstr;
        }

        ::SysFreeString(bstr);
    }

    return value;
}

CString CBCGPXmlParseError::GetSrcText()
{
    BSTR bstr = NULL;
    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMPARSEERROR(*this)->get_srcText(&bstr)
                : E_POINTER
        );

    CString value;
    if (bstr != NULL)
    {
        if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
        {
            value = bstr;
        }

        ::SysFreeString(bstr);
    }

    return value;
}

CString CBCGPXmlParseError::GetURL()
{
    BSTR bstr = NULL;
    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMPARSEERROR(*this)->get_url(&bstr)
                : E_POINTER
        );

    CString value;
    if (bstr != NULL)
    {
        if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
        {
            value = bstr;
        }

        ::SysFreeString(bstr);
    }

    return value;
}

CBCGPXmlParseErrorCollection CBCGPXmlParseError::GetAllErrors()
{
	CBCGPXmlParseErrorCollection value;

    HRESULT result = E_POINTER;

    if (IsValid())
    {
		ATL::CComQIPtr<BCG_MSXML::IXMLDOMParseError2> error;
		GetInterface()->QueryInterface(__uuidof(BCG_MSXML::IXMLDOMParseError2), (void**)&error);

        result = error != NULL
                    ? error->get_allErrors(BCG_MSXML_DOMPARSEERRORCOLLECTION_PTR(value))
                    : E_NOINTERFACE;
    }

    SetLastResult(result);

    return value;
}

CString CBCGPXmlParseError::GetErrorXPath()
{
	CString value;
	HRESULT result = E_POINTER;

	if (IsValid())
	{
		ATL::CComQIPtr<BCG_MSXML::IXMLDOMParseError2> error;
		GetInterface()->QueryInterface(__uuidof(BCG_MSXML::IXMLDOMParseError2), (void**)&error);

        BSTR bstr = NULL;
        result = error != NULL
                    ? error->get_errorXPath(&bstr)
                    : E_NOINTERFACE;

        if (bstr != NULL)
        {
            if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
            {
                value = bstr;
            }

            ::SysFreeString(bstr);
        }
	}

    SetLastResult(result);

    return value;
}

long CBCGPXmlParseError::GetErrorParametersCount()
{
    long value = -1;
    HRESULT result = E_POINTER;

    if (IsValid())
    {
		ATL::CComQIPtr<BCG_MSXML::IXMLDOMParseError2> error;
		GetInterface()->QueryInterface(__uuidof(BCG_MSXML::IXMLDOMParseError2), (void**)&error);

        result = error != NULL
                    ? error->get_errorParametersCount(&value)
                    : E_NOINTERFACE;
    }

    SetLastResult(result);

    return value;
}

CString CBCGPXmlParseError::GetErrorParameters(long index)
{
    CString value;
    HRESULT result = E_POINTER;

    if (IsValid())
    {
		ATL::CComQIPtr<BCG_MSXML::IXMLDOMParseError2> error;
		GetInterface()->QueryInterface(__uuidof(BCG_MSXML::IXMLDOMParseError2), (void**)&error);

        BSTR bstr = NULL;
        result = error != NULL
                    ? error->errorParameters(index, &bstr)
                    : E_NOINTERFACE;

        if (bstr != NULL)
        {
            if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
            {
                value = bstr;
            }

            ::SysFreeString(bstr);
        }
    }

    SetLastResult(result);

    return value;

}


CBCGPXmlDocument::CBCGPXmlDocument()
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMDocument))
{
}

CBCGPXmlDocument::CBCGPXmlDocument(const CBCGPXmlNode& other)
	: CBCGPXmlNode(__uuidof(BCG_MSXML::IXMLDOMDocument))
{
	((CBCGPXmlBase&)*this) = (const CBCGPXmlBase&)other;
}

CBCGPXmlDocument::CBCGPXmlDocument(const CBCGPXmlDocument& other)
	: CBCGPXmlNode(other)
{
}

CBCGPXmlDocument::~CBCGPXmlDocument()
{
}

BOOL CBCGPXmlDocument::Create(XXmlDocument document)
{
	if (IsValid())
	{
		return TRUE;
	}

	if (!CreateInstance(document))
	{
		return FALSE;
	}

	return TRUE;
}

BOOL CBCGPXmlDocument::CreateDocument
(
	LPCTSTR encoding,
	LPCTSTR root,
	LPCTSTR version,
	BOOL standalone,
	XXmlDocument document
)
{
	if (!Create(document))
	{
		return FALSE;
	}

	CString data;
	data.Format
			(
				_T("version=\"%s\" encoding=\"%s\""), 
				BCGPIsStringEmpty(version)
					? _T("1.0")
					: version, 
				BCGPIsStringEmpty(encoding)
					? CBCGPXmlProcessingInstruction::c_szEncoding_Utf16
					: encoding
			);

	if (standalone)
	{
		CString str;
		str.Format(_T(" standalone=\"%s\""), c_szXml_Yes);
		data += str;
	}

	CBCGPXmlProcessingInstruction processing(CreateProcessingInstruction(_T("xml"), data));
	if (BCGP_HR_SUCCEEDED_OK(GetLastResult()) && processing.IsValid())
	{
		AppendChild(processing);
	}

	if (BCGP_HR_SUCCEEDED_OK(GetLastResult()) && !BCGPIsStringEmpty(root))
	{
		CBCGPXmlElement element(CreateElement(root));
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()) && element.IsValid())
		{
			AppendChild(element);
		}
	}

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlDocument::CreateInstance(XXmlDocument document)
{
#ifndef _BCGSUITE_
	if (!globalData.ComInit())
	{
		return FALSE;
	}
#endif

	if (IsValid())
	{
		return TRUE;
	}

	HRESULT result = S_FALSE;

	CLSID clsids[5];
	clsids[0] = BCG_MSXML::CLSID_DOMDocument60; // {0x88d96a05, 0xf192, 0x11d4, {0xa6, 0x5f, 0x00, 0x40, 0x96, 0x32, 0x51, 0xe5}}
	clsids[1] = BCG_MSXML::CLSID_DOMDocument40; // {0x88d969c0, 0xf192, 0x11d4, {0xa6, 0x5f, 0x00, 0x40, 0x96, 0x32, 0x51, 0xe5}}
	clsids[2] = BCG_MSXML::CLSID_DOMDocument30; // {0xf5078f32, 0xc551, 0x11d3, {0x89, 0xb9, 0x00, 0x00, 0xf8, 0x1f, 0xe2, 0x21}}
	clsids[3] = BCG_MSXML::CLSID_DOMDocument26; // {0xf5078f1b, 0xc551, 0x11d3, {0x89, 0xb9, 0x00, 0x00, 0xf8, 0x1f, 0xe2, 0x21}}
	clsids[4] = BCG_MSXML::CLSID_DOMDocument;   // {0xf6d90f11, 0x9c73, 0x11d3, {0xb3, 0x2e, 0x00, 0xc0, 0x4f, 0x99, 0x0b, 0xb4}}

	if (document == e_XmlDocumentDefault)
	{
		for (int i = 0; i < _countof(clsids); i++)
		{
			result = ::CoCreateInstance(clsids[i], NULL, CLSCTX_INPROC_SERVER, m_iid, (void**)BCG_MSXML_DOMDOCUMENT_PTR(*this));
			if (BCGP_HR_SUCCEEDED_OK(result) && IsValid())
			{
				document = (XXmlDocument)(_countof(clsids) - i);
				break;
			}
		}
	}
	else
	{
		result = ::CoCreateInstance(clsids[_countof(clsids) - document], NULL, CLSCTX_INPROC_SERVER, m_iid, (void**)BCG_MSXML_DOMDOCUMENT_PTR(*this));
	}

	if (BCGP_HR_SUCCEEDED_OK(result) && !IsValid())
	{
		result = S_FALSE;
	}

	SetLastResult(result);

	if (IsValid())
	{
		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			SetAsync(TRUE);
		}

		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			SetPreserveWhiteSpace(TRUE);
		}

		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			SetResolveExternals(FALSE);
		}

		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			SetValidateOnParse(FALSE);
		}

		if (BCGP_HR_SUCCEEDED_OK(GetLastResult()))
		{
			if (document > e_XmlDocument30)
			{
				SetProperty(_T("SelectionLanguage"), _T("XPath"));
			}
		}
	}

	if (!BCGP_HR_SUCCEEDED_OK(GetLastResult()) && IsValid())
	{
		Release();
	}

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CBCGPXmlAttribute CBCGPXmlDocument::CreateAttribute(const CString& strName, LPCTSTR namespaceURI)
{
	CBCGPXmlAttribute value;

	if (IsValid())
	{
		if (!BCGPIsStringEmpty(namespaceURI))
		{
			value = CreateNode(BCGP_NODE_ATTRIBUTE, strName, namespaceURI);
		}
		else
		{
			SetLastResult(BCG_MSXML_DOMDOCUMENT(*this)->createAttribute(_bstr_t(strName), BCG_MSXML_DOMATTRIBUTE_PTR(value)));
		}
	}

	return value;
}

CBCGPXmlCDATASection CBCGPXmlDocument::CreateCDATASection(const CString& data)
{
	CBCGPXmlCDATASection value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->createCDATASection(_bstr_t(data), BCG_MSXML_DOMCDATASECTION_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlDocumentFragment CBCGPXmlDocument::CreateDocumentFragment()
{
	CBCGPXmlDocumentFragment value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->createDocumentFragment(BCG_MSXML_DOMDOCUMENTFRAGMENT_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlComment CBCGPXmlDocument::CreateComment(const CString& comment)
{
	CBCGPXmlComment value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->createComment(_bstr_t(comment), BCG_MSXML_DOMCOMMENT_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlElement CBCGPXmlDocument::CreateElement(const CString& strName, LPCTSTR namespaceURI)
{
	CBCGPXmlElement value;

	if (IsValid())
	{
		if (!BCGPIsStringEmpty(namespaceURI))
		{
			value = CreateNode(BCGP_NODE_ELEMENT, strName, namespaceURI);
		}
		else
		{
			SetLastResult(BCG_MSXML_DOMDOCUMENT(*this)->createElement(_bstr_t(strName), BCG_MSXML_DOMELEMENT_PTR(value)));
		}
	}

	return value;
}

CBCGPXmlEntityReference CBCGPXmlDocument::CreateEntityReferrence(const CString& strName)
{
    CBCGPXmlEntityReference value;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMDOCUMENT(*this)->createEntityReference(_bstr_t(strName), BCG_MSXML_DOMENTITYREFERENCE_PTR(value))
                : E_POINTER
        );

    return value;
}

CBCGPXmlNode CBCGPXmlDocument::CreateNode(BCGPDOMNodeType type, const CString& strName, LPCTSTR namespaceURI)
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->createNode(variant_t((long)type), _bstr_t(strName), _bstr_t(namespaceURI), BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlProcessingInstruction CBCGPXmlDocument::CreateProcessingInstruction(const CString& target, const CString& data)
{
	CBCGPXmlProcessingInstruction value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->createProcessingInstruction(_bstr_t(target), _bstr_t(data), BCG_MSXML_DOMPROCESSINGINSTRUCTION_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlText CBCGPXmlDocument::CreateTextNode(const CString& text)
{
	CBCGPXmlText value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->createTextNode(_bstr_t(text), BCG_MSXML_DOMTEXT_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlDocument::XReadyState CBCGPXmlDocument::GetReadyState()
{
	long value = 0;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->get_readyState(&value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult())
			? XReadyState(value)
			: e_ReadyStateUnknown;
}

CBCGPXmlParseError CBCGPXmlDocument::GetParseError()
{
    CBCGPXmlParseError value;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMDOCUMENT(*this)->get_parseError(BCG_MSXML_DOMPARSEERROR_PTR(value))
                : E_POINTER
        );

    return value;
}

CBCGPXmlParseError CBCGPXmlDocument::Validate()
{
    CBCGPXmlParseError value;
	HRESULT result = E_POINTER;

	if (IsValid())
	{
		ATL::CComQIPtr<BCG_MSXML::IXMLDOMDocument2> doc;
		GetInterface()->QueryInterface(__uuidof(BCG_MSXML::IXMLDOMDocument2), (void**)&doc);

		result = doc != NULL
					? doc->validate(BCG_MSXML_DOMPARSEERROR_PTR(value))
					: E_NOINTERFACE;
	}

	SetLastResult(result);

    return value;
}

CBCGPXmlParseError CBCGPXmlDocument::ValidateNode(const CBCGPXmlNode& node)
{
    CBCGPXmlParseError value;
	HRESULT result = E_POINTER;

	if (IsValid())
	{
		ATL::CComQIPtr<BCG_MSXML::IXMLDOMDocument3> doc;
		GetInterface()->QueryInterface(__uuidof(BCG_MSXML::IXMLDOMDocument3), (void**)&doc);

		result = doc != NULL
					? doc->validateNode(BCG_MSXML_DOMNODE(const_cast<CBCGPXmlNode&>(node)), BCG_MSXML_DOMPARSEERROR_PTR(value))
					: E_NOINTERFACE;
	}

	SetLastResult(result);

    return value;
}

BOOL CBCGPXmlDocument::Abort()
{
    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMDOCUMENT(*this)->abort()
                : E_POINTER
        );

    return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CBCGPXmlDocumentType CBCGPXmlDocument::GetDocumentType()
{
	CBCGPXmlDocumentType value;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMDOCUMENT(*this)->get_doctype(BCG_MSXML_DOMDOCUMENTTYPE_PTR(value))
                : E_POINTER
        );

    return value;
}

CBCGPXmlImplemetation CBCGPXmlDocument::GetImplementation()
{
	CBCGPXmlImplemetation value;

    SetLastResult
        (
            IsValid()
                ? BCG_MSXML_DOMDOCUMENT(*this)->get_implementation(BCG_MSXML_DOMIMPLEMENTATION_PTR(value))
                : E_POINTER
        );

    return value;
}

CBCGPXmlProcessingInstruction CBCGPXmlDocument::GetProcessingInstruction()
{
	CBCGPXmlProcessingInstruction value;
	HRESULT result = E_POINTER;

	if (IsValid())
	{
		result = E_NOINTERFACE;

		CBCGPXmlNodeList children(GetChildren());

		if (children.IsValid())
		{
			for (long i = 0; i < children.GetLength(); i++)
			{
				CBCGPXmlNode child(children.GetItem(i));

				if (child.IsValid())
				{
					BCGPDOMNodeType type = child.GetNodeType();
					result = child.GetLastResult();

					if (BCGP_HR_SUCCEEDED_OK(result) && type == BCGP_NODE_PROCESSING_INSTRUCTION)
					{
						value = child;
						result = value.GetLastResult();
						break;
					}
				}
				else
				{
					result = children.GetLastResult();
					break;
				}
			}
		}
		else
		{
			result = GetLastResult();
		}
	}

	SetLastResult(result);

	return value;
}

CBCGPXmlElement CBCGPXmlDocument::GetDocumentElement()
{
	CBCGPXmlElement value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->get_documentElement(BCG_MSXML_DOMELEMENT_PTR(value))
				: E_POINTER
		);

	return value;
}

BOOL CBCGPXmlDocument::SetDocumentElement(const CBCGPXmlElement& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->putref_documentElement(BCG_MSXML_DOMELEMENT(const_cast<CBCGPXmlElement&>(value)))
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

_variant_t CBCGPXmlDocument::GetProperty(const CString& strName)
{
	_variant_t value;
	HRESULT result = E_POINTER;

	if (IsValid())
	{
		ATL::CComQIPtr<BCG_MSXML::IXMLDOMDocument2> doc;
		GetInterface()->QueryInterface(__uuidof(BCG_MSXML::IXMLDOMDocument2), (void**)&doc);

		result = doc != NULL
					? doc->getProperty(_bstr_t(strName), &value)
					: E_NOINTERFACE;
	}

	SetLastResult(result);

	return value;
}

BOOL CBCGPXmlDocument::SetProperty(const CString& strName, const _variant_t& value)
{
	HRESULT result = E_POINTER;

	if (IsValid())
	{
		ATL::CComQIPtr<BCG_MSXML::IXMLDOMDocument2> doc;
		GetInterface()->QueryInterface(__uuidof(BCG_MSXML::IXMLDOMDocument2), (void**)&doc);

		result = doc != NULL
					? doc->setProperty(_bstr_t(strName), value)
					: E_NOINTERFACE;
	}

	SetLastResult(result);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlDocument::GetAsync()
{
	VARIANT_BOOL value = BCGP_VARIANT_FALSE;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->get_async(&value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult()) && value == BCGP_VARIANT_TRUE;
}

BOOL CBCGPXmlDocument::SetAsync(BOOL value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->put_async(value ? BCGP_VARIANT_TRUE : BCGP_VARIANT_FALSE)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlDocument::GetPreserveWhiteSpace()
{
	VARIANT_BOOL value = BCGP_VARIANT_FALSE;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->get_preserveWhiteSpace(&value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult()) && value == BCGP_VARIANT_TRUE;
}

BOOL CBCGPXmlDocument::SetPreserveWhiteSpace(BOOL value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->put_preserveWhiteSpace(value ? BCGP_VARIANT_TRUE : BCGP_VARIANT_FALSE)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlDocument::GetResolveExternals()
{
	VARIANT_BOOL value = BCGP_VARIANT_FALSE;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->get_resolveExternals(&value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult()) && value == BCGP_VARIANT_TRUE;
}

BOOL CBCGPXmlDocument::SetResolveExternals(BOOL value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->put_resolveExternals(value ? BCGP_VARIANT_TRUE : BCGP_VARIANT_FALSE)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlDocument::GetValidateOnParse()
{
	VARIANT_BOOL value = BCGP_VARIANT_FALSE;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->get_validateOnParse(&value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult()) && value == BCGP_VARIANT_TRUE;
}

BOOL CBCGPXmlDocument::SetValidateOnParse(BOOL value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->put_validateOnParse(value ? BCGP_VARIANT_TRUE : BCGP_VARIANT_FALSE)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

CBCGPXmlNodeList CBCGPXmlDocument::GetElementsByTagName(const CString& strName)
{
	CBCGPXmlNodeList value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->getElementsByTagName(_bstr_t(strName), BCG_MSXML_DOMNODELIST_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNode CBCGPXmlDocument::GetNodeFromID(const CString& id)
{
	CBCGPXmlNode value;

	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->nodeFromID(_bstr_t(id), BCG_MSXML_DOMNODE_PTR(value))
				: E_POINTER
		);

	return value;
}

CBCGPXmlNode CBCGPXmlDocument::ImportNode(const CBCGPXmlNode& node, BOOL deep)
{
    CBCGPXmlNode value;
	HRESULT result = E_POINTER;

	if (IsValid())
	{
		ATL::CComQIPtr<BCG_MSXML::IXMLDOMDocument3> doc;
		GetInterface()->QueryInterface(__uuidof(BCG_MSXML::IXMLDOMDocument3), (void**)&doc);

		result = doc != NULL
					? doc->importNode(BCG_MSXML_DOMNODE(const_cast<CBCGPXmlNode&>(node)), deep ? BCGP_VARIANT_TRUE : BCGP_VARIANT_FALSE, BCG_MSXML_DOMNODE_PTR(value))
					: E_NOINTERFACE;
	}

	SetLastResult(result);

    return value;
}

BOOL CBCGPXmlDocument::Load(const CString& fileName)
{
	return _Load(_variant_t(_bstr_t(fileName)));
}

BOOL CBCGPXmlDocument::Load(CFile& file)
{
	HRESULT result = E_INVALIDARG;
	BOOL bRes = FALSE;

	const UINT size = (UINT)file.GetLength();
	if (size > 0)
	{
		BYTE* buffer = new BYTE[size];
		if (buffer != NULL)
		{
			if (file.Read(buffer, size) == size)
			{
				bRes = Load(buffer, (UINT)size);
				result = GetLastResult();
			}

			delete [] buffer;
		}
	}

	SetLastResult(result);

	return bRes;
}

BOOL CBCGPXmlDocument::Load(LPBYTE buffer, UINT size)
{
	HRESULT result = E_INVALIDARG;
	BOOL bRes = FALSE;

	if (size > 0)
	{
		HGLOBAL hGlobal = ::GlobalAlloc(GHND, size);
		if (hGlobal != NULL)
		{
			LPVOID lpVoid = ::GlobalLock(hGlobal);

			if (lpVoid != NULL)
			{
				memcpy(lpVoid, buffer, size);

				::GlobalUnlock(hGlobal);

				IStream* stream = NULL;
				result = ::CreateStreamOnHGlobal(hGlobal, FALSE, &stream);

				if (BCGP_HR_SUCCEEDED_OK(result) && stream != NULL)
				{
					bRes = Load(stream);
					GetLastResult();
				}

				if (stream != NULL)
				{
					stream->Release();
				}
			}

			::GlobalFree(hGlobal);
		}
	}

	SetLastResult(result);

	return bRes;
}

BOOL CBCGPXmlDocument::Load(IStream* stream)
{
	STATSTG stat = {0};
	stream->Stat(&stat, STATFLAG_NONAME);
	ULONGLONG size = stat.cbSize.QuadPart;
	if (size == 0)
	{
		SetLastResult(E_INVALIDARG);
		return FALSE;
	}

	return _Load(_variant_t(stream, TRUE));
}

BOOL CBCGPXmlDocument::_Load(const _variant_t& value)
{
	if (!Create())
	{
		return FALSE;
	}

	VARIANT_BOOL varBool = BCGP_VARIANT_FALSE;
	SetLastResult(BCG_MSXML_DOMDOCUMENT(*this)->load(value, &varBool));

	return BCGP_HR_SUCCEEDED_OK(GetLastResult()) && varBool == BCGP_VARIANT_TRUE;
}

BOOL CBCGPXmlDocument::LoadXML(const CString& xml)
{
	if (!Create())
	{
		return FALSE;
	}

	VARIANT_BOOL varBool = BCGP_VARIANT_FALSE;
	SetLastResult(BCG_MSXML_DOMDOCUMENT(*this)->loadXML(_bstr_t(xml), &varBool));

	return BCGP_HR_SUCCEEDED_OK(GetLastResult()) && varBool == BCGP_VARIANT_TRUE;
}

BOOL CBCGPXmlDocument::Save(const CString& fileName)
{
	return _Save(_variant_t(_bstr_t(fileName)));
}

BOOL CBCGPXmlDocument::Save(CFile& file)
{
	BOOL bRes = FALSE;

	LPBYTE buffer = NULL;
	UINT size = 0;

	if (Save(&buffer, size))
	{
		file.Write(buffer, size);
		bRes = TRUE;
	}

	if (buffer != NULL)
	{
		delete [] buffer;
	}

	return bRes;
}

BOOL CBCGPXmlDocument::Save(LPBYTE* buffer, UINT& size)
{
	HRESULT result = E_INVALIDARG;
	BOOL bRes = FALSE;

	size = 0;

	if (buffer != NULL)
	{
		HGLOBAL hGlobal = ::GlobalAlloc(GHND, 0);
		if (hGlobal != NULL)
		{
			IStream* stream = NULL;
			result = ::CreateStreamOnHGlobal(hGlobal, FALSE, &stream);

			if (BCGP_HR_SUCCEEDED_OK(result) && stream != NULL)
			{
				bRes = Save(stream);
				GetLastResult();

				if (bRes)
				{
					STATSTG stat = {0};
					stream->Stat(&stat, STATFLAG_NONAME);
					UINT ulSize = (UINT)stat.cbSize.QuadPart;

					if (ulSize > 0)
					{
						*buffer = new BYTE[ulSize];
						LARGE_INTEGER dlibMove = {0};
						stream->Seek(dlibMove, STREAM_SEEK_SET, NULL);
						stream->Read(*buffer, ulSize, NULL);

						size = ulSize;
					}
				}
			}

			if (stream != NULL)
			{
				stream->Release();
			}

			::GlobalFree(hGlobal);
		}
	}

	SetLastResult(result);

	return bRes;
}

BOOL CBCGPXmlDocument::Save(IStream* stream)
{
	return _Save(_variant_t(stream, TRUE));
}

BOOL CBCGPXmlDocument::_Save(const _variant_t& value)
{
	SetLastResult
		(
			IsValid()
				? BCG_MSXML_DOMDOCUMENT(*this)->save(value)
				: E_POINTER
		);

	return BCGP_HR_SUCCEEDED_OK(GetLastResult());
}

BOOL CBCGPXmlDocument::Format(CBCGPXmlDocument& document)
{
	HRESULT result = E_POINTER;
	BOOL bRes = FALSE;

	if (IsValid())
	{
		if (document.Create())
		{
			document.SetPreserveWhiteSpace(GetPreserveWhiteSpace());

			CBCGPXmlDocument stylesheet;
			bRes = stylesheet.LoadXML(c_szXml_FormatXSLT);
			result = stylesheet.GetLastResult();

			if (bRes)
			{
				bRes = TransformNodeToDocument(CBCGPXmlNode(stylesheet), document);
				if (bRes)
				{
					CBCGPXmlProcessingInstruction src(GetProcessingInstruction());
					CBCGPXmlProcessingInstruction dst(document.GetProcessingInstruction());

					if (src.IsValid() && dst.IsValid())
					{
						dst.SetEncoding(src.GetEncoding());
					}
				}
			}
		}
	}

	SetLastResult(result);

	return bRes;
}
