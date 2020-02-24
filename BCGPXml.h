// BCGPXml.h: interface for the CBCGPXml class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BCGPXML_H__90C51E93_0257_4B10_8AC2_EF654091EA98__INCLUDED_)
#define AFX_BCGPXML_H__90C51E93_0257_4B10_8AC2_EF654091EA98__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "BCGCBPro.h"

#include <atlbase.h>
#include "comdef.h"


class CBCGPXmlNode;
class CBCGPXmlText;
class CBCGPXmlDocument;
class CBCGPXmlNodeList;
class CBCGPXmlNamedNodeMap;
class CBCGPXmlParseErrorCollection;


bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, bool& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, char& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, signed char& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, unsigned char& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, signed short& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, unsigned short& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, signed long& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, unsigned long& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, signed int& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, unsigned int& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, signed __int64& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, unsigned __int64& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, float& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, double& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, long double& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, CString& value);
bool BCGCBPRODLLEXPORT BCGPGetVariantValue(const _variant_t& var, COleDateTime& value);


class CBCGPXmlBase
{
public:
	explicit CBCGPXmlBase(REFIID iid)
		: m_iid(iid)
		, m_p  (NULL)
		, m_result(S_OK)
	{
	}
	CBCGPXmlBase(const CBCGPXmlBase& other)
		: m_iid   (other.m_iid)
		, m_result(other.GetLastResult())
	{
		if ((m_p = other.m_p) != NULL)
		{
			m_p->AddRef();
		}
	}

	virtual ~CBCGPXmlBase()
	{
		if (IsValid())
		{
			m_p->Release();
		}
	}

public:
	BOOL IsValid() const 
	{
		return m_p != NULL; 
	}

	HRESULT GetLastResult() const
	{
		return m_result;
	}
	void ClearLastResult()
	{
		SetLastResult(S_OK);
	}

	void Release()
	{
		if (IsValid())
		{
			IUnknown* p = m_p;

			m_p = NULL;
			p->Release();
		}
	}

	ATL::_NoAddRefReleaseOnCComPtr<IUnknown>* GetInterface()
	{
		ATLASSERT(m_p != NULL);
		return (ATL::_NoAddRefReleaseOnCComPtr<IUnknown>*)m_p;
	}
	IUnknown** GetInterfacePtr()
	{
		ATLASSERT(m_p == NULL);
		return &m_p;
	}

	const CBCGPXmlBase& operator = (const CBCGPXmlBase& other)
	{
		m_result = other.GetLastResult();

		if (m_iid == other.m_iid)
		{
			ATL::AtlComPtrAssign((IUnknown**)&m_p, other.m_p);
		}
		else
		{
			ATL::AtlComQIPtrAssign((IUnknown**)&m_p, other.m_p, m_iid);
		}

		return *this;
	}

protected:
	void SetLastResult(HRESULT hr) { m_result = hr; }

protected:
	IUnknown*	m_p;
	GUID		m_iid;

	HRESULT 	m_result;
};

enum BCGPDOMNodeType
{
    BCGP_NODE_INVALID = 0,
    BCGP_NODE_ELEMENT = 1,
    BCGP_NODE_ATTRIBUTE = 2,
    BCGP_NODE_TEXT = 3,
    BCGP_NODE_CDATA_SECTION = 4,
    BCGP_NODE_ENTITY_REFERENCE = 5,
    BCGP_NODE_ENTITY = 6,
    BCGP_NODE_PROCESSING_INSTRUCTION = 7,
    BCGP_NODE_COMMENT = 8,
    BCGP_NODE_DOCUMENT = 9,
    BCGP_NODE_DOCUMENT_TYPE = 10,
    BCGP_NODE_DOCUMENT_FRAGMENT = 11,
    BCGP_NODE_NOTATION = 12
};

class BCGCBPRODLLEXPORT CBCGPXmlNode: public CBCGPXmlBase
{
protected:
	explicit CBCGPXmlNode(REFIID iid);

public:
	CBCGPXmlNode();
	CBCGPXmlNode(const CBCGPXmlNode& other);

	virtual ~CBCGPXmlNode();

public:
	BCGPDOMNodeType GetNodeType();
	CString GetNodeTypeString();
	CString GetNodeName();

	_variant_t GetNodeTypedValue();
	BOOL SetNodeTypedValue(const _variant_t& value);

	CBCGPXmlNode GetParentNode();

	CString GetBaseName();

	_variant_t GetDataType();
	BOOL SetDataType(const CString& value);

	_variant_t GetNodeValue();
	BOOL SetNodeValue(const _variant_t& value);

	CString GetNamespaceURI();

	CString GetPrefix();

	BOOL IsParsed();

	BOOL IsSpecified();

	CString GetText();
	BOOL SetText(const CString& value);

	BOOL GetXML(CString& value);

	CBCGPXmlNode GetDefinition();

	BOOL HasChildren();

	BOOL GetOwnerDocument(CBCGPXmlDocument& value);

	CBCGPXmlNode CloneNode(BOOL deep = TRUE);

	CBCGPXmlNode GetChildFirst();
	CBCGPXmlNode GetChildLast();

	CBCGPXmlNode GetSiblingNext();
	CBCGPXmlNode GetSiblingPrev();

	BOOL AppendChild(const CBCGPXmlNode& newValue, CBCGPXmlNode* outValue = NULL);
	BOOL RemoveChild(const CBCGPXmlNode& value, CBCGPXmlNode* outValue = NULL);
	BOOL ReplaceChild(const CBCGPXmlNode& newValue, const CBCGPXmlNode& oldValue, CBCGPXmlNode* outValue = NULL);

	CBCGPXmlNodeList GetChildren();

	CBCGPXmlNamedNodeMap GetAttributes();

	CBCGPXmlNode SelectNode(const CString& query);
	CBCGPXmlNodeList SelectNodes(const CString& query);

	CString TransformNode(const CBCGPXmlNode& stylesheet);
	BOOL TransformNodeToObject(const CBCGPXmlNode& stylesheet, _variant_t& value);
	BOOL TransformNodeToDocument(const CBCGPXmlNode& stylesheet, CBCGPXmlDocument& document);
	BOOL TransformNodeToStream(const CBCGPXmlNode& stylesheet, IStream* stream);
};


class BCGCBPRODLLEXPORT CBCGPXmlCharacterData: public CBCGPXmlNode
{
protected:
	explicit CBCGPXmlCharacterData(REFIID iid);
	
public:
	CBCGPXmlCharacterData();
	CBCGPXmlCharacterData(const CBCGPXmlNode& other);
	CBCGPXmlCharacterData(const CBCGPXmlCharacterData& other);

	virtual ~CBCGPXmlCharacterData();

public:
	long GetLength();

	CString GetData();
	BOOL SetData(const CString& value);

	BOOL AppendData(const CString& value);
	BOOL InsertData(long offset, const CString& value);
	BOOL DeleteData(long offset, long count);
	BOOL ReplaceData(long offset, long count, const CString& value);
	CString SubstringData(long offset, long count);
};


class BCGCBPRODLLEXPORT CBCGPXmlText: public CBCGPXmlCharacterData
{
protected:
	explicit CBCGPXmlText(REFIID iid);

public:
	CBCGPXmlText();
	CBCGPXmlText(const CBCGPXmlNode& other);
	CBCGPXmlText(const CBCGPXmlCharacterData& other);
	CBCGPXmlText(const CBCGPXmlText& other);

	virtual ~CBCGPXmlText();

public:
	BOOL SplitText(long offset, CBCGPXmlText& value);
};


class BCGCBPRODLLEXPORT CBCGPXmlComment: public CBCGPXmlCharacterData
{
public:
	CBCGPXmlComment();
	CBCGPXmlComment(const CBCGPXmlNode& other);
	CBCGPXmlComment(const CBCGPXmlCharacterData& other);
	CBCGPXmlComment(const CBCGPXmlComment& other);

	virtual ~CBCGPXmlComment();
};

class BCGCBPRODLLEXPORT CBCGPXmlCDATASection: public CBCGPXmlText
{
public:
	CBCGPXmlCDATASection();
	CBCGPXmlCDATASection(const CBCGPXmlNode& other);
	CBCGPXmlCDATASection(const CBCGPXmlCharacterData& other);
	CBCGPXmlCDATASection(const CBCGPXmlText& other);
	CBCGPXmlCDATASection(const CBCGPXmlCDATASection& other);

	virtual ~CBCGPXmlCDATASection();
};


class BCGCBPRODLLEXPORT CBCGPXmlAttribute: public CBCGPXmlNode
{
public:
	CBCGPXmlAttribute();
	CBCGPXmlAttribute(const CBCGPXmlNode& other);
	CBCGPXmlAttribute(const CBCGPXmlAttribute& other);

	virtual ~CBCGPXmlAttribute();

public:
	CString GetName();

	_variant_t GetValue();
	BOOL SetValue(const _variant_t& value);

    operator bool();
    operator char();
    operator signed char();
    operator unsigned char();
    operator signed short();
    operator unsigned short();
    operator long();
    operator unsigned long();
    operator int();
    operator unsigned int();
#if (_WIN32_WINNT >= 0x0501)
    operator __int64();
    operator unsigned __int64();
#endif
    operator float();
    operator double();
    operator long double();
    operator CString();
    operator COleDateTime();
};


class BCGCBPRODLLEXPORT CBCGPXmlElement: public CBCGPXmlNode
{
public:
	CBCGPXmlElement();
	CBCGPXmlElement(const CBCGPXmlNode& other);
	CBCGPXmlElement(const CBCGPXmlElement& other);

	virtual ~CBCGPXmlElement();

public:
	CString GetTagName();

	CBCGPXmlNodeList GetElementsByTagName(const CString& strName);

	_variant_t GetAttribute(const CString& strName);

    bool GetAttributeValue(const CString& strName, bool& value);
    bool GetAttributeValue(const CString& strName, char& value);
    bool GetAttributeValue(const CString& strName, signed char& value);
    bool GetAttributeValue(const CString& strName, unsigned char& value);
    bool GetAttributeValue(const CString& strName, signed short& value);
    bool GetAttributeValue(const CString& strName, unsigned short& value);
    bool GetAttributeValue(const CString& strName, long& value);
    bool GetAttributeValue(const CString& strName, unsigned long& value);
    bool GetAttributeValue(const CString& strName, int& value);
    bool GetAttributeValue(const CString& strName, unsigned int& value);
#if (_WIN32_WINNT >= 0x0501)
    bool GetAttributeValue(const CString& strName, __int64& value);
    bool GetAttributeValue(const CString& strName, unsigned __int64& value);
#endif
    bool GetAttributeValue(const CString& strName, float& value);
    bool GetAttributeValue(const CString& strName, double& value);
    bool GetAttributeValue(const CString& strName, long double& value);
    bool GetAttributeValue(const CString& strName, CString& value);
    bool GetAttributeValue(const CString& strName, COleDateTime& value);

	BOOL GetBoolAttribute(const CString& strName);
	long GetIntAttribute(const CString& strName);
	CString GetStringAttribute(const CString& strName);
	void GetRectangleAttribute(CRect& rect);

	BOOL SetAttribute(const CString& strName, const _variant_t& value);
	BOOL RemoveAttribute(const CString& strName);

	void SetRectangleAttribute(const CRect& rect);

	CBCGPXmlAttribute GetAttributeNode(const CString& strName);
	BOOL SetAttributeNode(const CBCGPXmlAttribute& value, CBCGPXmlAttribute* outValue = NULL);
	BOOL RemoveAttributeNode(const CBCGPXmlAttribute& value, CBCGPXmlAttribute* outValue = NULL);

	BOOL Normalize();

	CBCGPXmlElement FindChildElement(const CString& strName);
};


class BCGCBPRODLLEXPORT CBCGPXmlEntity: public CBCGPXmlNode
{
public:
	CBCGPXmlEntity();
	CBCGPXmlEntity(const CBCGPXmlNode& other);
	CBCGPXmlEntity(const CBCGPXmlEntity& other);

	virtual ~CBCGPXmlEntity();

public:
	CString GetNotationName();

	_variant_t GetPublicId();
	_variant_t GetSystemId();
};


class BCGCBPRODLLEXPORT CBCGPXmlEntityReference: public CBCGPXmlNode
{
public:
	CBCGPXmlEntityReference();
	CBCGPXmlEntityReference(const CBCGPXmlNode& other);
	CBCGPXmlEntityReference(const CBCGPXmlEntityReference& other);

	virtual ~CBCGPXmlEntityReference();
};


class BCGCBPRODLLEXPORT CBCGPXmlProcessingInstruction: public CBCGPXmlNode
{
public:
	static const LPCTSTR   c_szEncoding_Utf8;
	static const LPCTSTR   c_szEncoding_Utf16;

public:
	CBCGPXmlProcessingInstruction();
	CBCGPXmlProcessingInstruction(const CBCGPXmlNode& other);
	CBCGPXmlProcessingInstruction(const CBCGPXmlProcessingInstruction& other);

	virtual ~CBCGPXmlProcessingInstruction();

public:
	CString GetTarget();

	CString GetData();
	BOOL SetData(const CString& value);

	CString GetEncoding();
	BOOL SetEncoding(LPCTSTR encoding);
};


class BCGCBPRODLLEXPORT CBCGPXmlImplemetation: public CBCGPXmlBase
{
public:
    CBCGPXmlImplemetation();
    CBCGPXmlImplemetation(const CBCGPXmlImplemetation& other);

	virtual ~CBCGPXmlImplemetation();

public:
	BOOL HasFeature(const CString& value, LPCTSTR version = NULL);
};


class BCGCBPRODLLEXPORT CBCGPXmlNotation: public CBCGPXmlNode
{
public:
    CBCGPXmlNotation();
	CBCGPXmlNotation(const CBCGPXmlNode& other);
    CBCGPXmlNotation(const CBCGPXmlNotation& other);

	virtual ~CBCGPXmlNotation();

public:
	_variant_t GetPublicId();
	_variant_t GetSystemId();
};


class BCGCBPRODLLEXPORT CBCGPXmlDocumentFragment: public CBCGPXmlNode
{
public:
    CBCGPXmlDocumentFragment();
	CBCGPXmlDocumentFragment(const CBCGPXmlNode& other);
    CBCGPXmlDocumentFragment(const CBCGPXmlDocumentFragment& other);

	virtual ~CBCGPXmlDocumentFragment();
};


class BCGCBPRODLLEXPORT CBCGPXmlDocumentType: public CBCGPXmlNode
{
public:
    CBCGPXmlDocumentType();
	CBCGPXmlDocumentType(const CBCGPXmlNode& other);
    CBCGPXmlDocumentType(const CBCGPXmlDocumentType& other);

	virtual ~CBCGPXmlDocumentType();

public:
	CString GetName();

	CBCGPXmlNamedNodeMap GetEntities();
	CBCGPXmlNamedNodeMap GetNotations();
};


class BCGCBPRODLLEXPORT CBCGPXmlParseError: public CBCGPXmlBase
{
public:
	CBCGPXmlParseError();
	CBCGPXmlParseError(const CBCGPXmlParseError& other);

	virtual ~CBCGPXmlParseError();

public:
    long GerErrorCode();
    long GetFilePos();
    long GetLine();
    long GetLinePos();
    CString GetReason();
    CString GetSrcText();
    CString GetURL();

    CBCGPXmlParseErrorCollection GetAllErrors();
    CString GetErrorXPath();
    long GetErrorParametersCount();
    CString GetErrorParameters(long index);
};


class BCGCBPRODLLEXPORT CBCGPXmlDocument: public CBCGPXmlNode
{
public:
	enum XXmlDocument
	{
		e_XmlDocumentFirst   = 0,
		e_XmlDocumentDefault = e_XmlDocumentFirst,
		e_XmlDocument        = 1,
		e_XmlDocument26      = 2,
		e_XmlDocument30      = 3,
		e_XmlDocument40      = 4,
		e_XmlDocument60      = 5,
		e_XmlDocumentLast    = e_XmlDocument60
	};

	enum XReadyState
	{
		e_ReadyStateFirst		= 0,
		e_ReadyStateUnknown 	= e_ReadyStateFirst,
		e_ReadyStateLoading 	= 1,
		e_ReadyStateLoaded		= 2,
		e_ReadyStateInteractive = 3,
		e_ReadyStateCompleted	= 4,
		e_ReadyStateLast		= e_ReadyStateCompleted
	};

public:
	CBCGPXmlDocument();
	CBCGPXmlDocument(const CBCGPXmlNode& other);
	CBCGPXmlDocument(const CBCGPXmlDocument& other);

	virtual ~CBCGPXmlDocument();

public:
	BOOL Create(XXmlDocument document = e_XmlDocumentDefault);
	BOOL CreateDocument(LPCTSTR encoding, LPCTSTR root = NULL, LPCTSTR version = _T("1.0"), BOOL standalone = TRUE, XXmlDocument document = e_XmlDocumentDefault);

	CBCGPXmlAttribute CreateAttribute(const CString& strName, LPCTSTR namespaceURI = NULL);
	CBCGPXmlCDATASection CreateCDATASection(const CString& data);
	CBCGPXmlDocumentFragment CreateDocumentFragment();
	CBCGPXmlComment CreateComment(const CString& comment);
	CBCGPXmlElement CreateElement(const CString& strName, LPCTSTR namespaceURI = NULL);
	CBCGPXmlEntityReference CreateEntityReferrence(const CString& strName);
	CBCGPXmlNode CreateNode(BCGPDOMNodeType type, const CString& strName, LPCTSTR namespaceURI = NULL);
	CBCGPXmlProcessingInstruction CreateProcessingInstruction(const CString& target, const CString& data);
	CBCGPXmlText CreateTextNode(const CString& text);

	XReadyState GetReadyState();

	CBCGPXmlParseError GetParseError();

	CBCGPXmlParseError Validate();
	CBCGPXmlParseError ValidateNode(const CBCGPXmlNode& node);

	BOOL Abort();

	CBCGPXmlDocumentType GetDocumentType();

	CBCGPXmlImplemetation GetImplementation();

	CBCGPXmlProcessingInstruction GetProcessingInstruction();

	CBCGPXmlElement GetDocumentElement();
	BOOL SetDocumentElement(const CBCGPXmlElement& value);

	_variant_t GetProperty(const CString& strName);
	BOOL SetProperty(const CString& strName, const _variant_t& value);

	BOOL GetAsync();
	BOOL SetAsync(BOOL value);

	BOOL GetPreserveWhiteSpace();
	BOOL SetPreserveWhiteSpace(BOOL value);

	BOOL GetResolveExternals();
	BOOL SetResolveExternals(BOOL value);

	BOOL GetValidateOnParse();
	BOOL SetValidateOnParse(BOOL value);

	CBCGPXmlNodeList GetElementsByTagName(const CString& strName);

	CBCGPXmlNode GetNodeFromID(const CString& id);

	CBCGPXmlNode ImportNode(const CBCGPXmlNode& node, BOOL deep = TRUE);

	BOOL Load(const CString& fileName);
	BOOL Load(CFile& file);
	BOOL Load(LPBYTE buffer, UINT size);
	BOOL Load(IStream* stream);
	BOOL LoadXML(const CString& xml);

	BOOL Save(const CString& fileName);
	BOOL Save(CFile& file);
	BOOL Save(LPBYTE* buffer, UINT& size);
	BOOL Save(IStream* stream);

	BOOL Format(CBCGPXmlDocument& document);

protected:
	BOOL CreateInstance(XXmlDocument document);
	BOOL _Load(const _variant_t& value);
	BOOL _Save(const _variant_t& value);
};


class BCGCBPRODLLEXPORT CBCGPXmlNodeList: public CBCGPXmlBase
{
public:
	CBCGPXmlNodeList();
	CBCGPXmlNodeList(const CBCGPXmlNodeList& other);

	virtual ~CBCGPXmlNodeList();

public:
	long GetLength();
	void Reset();
	CBCGPXmlNode GetItem(long index);

	CBCGPXmlNode NextNode();
};


class BCGCBPRODLLEXPORT CBCGPXmlNamedNodeMap: public CBCGPXmlBase
{
public:
	CBCGPXmlNamedNodeMap();
	CBCGPXmlNamedNodeMap(const CBCGPXmlNamedNodeMap& other);

	virtual ~CBCGPXmlNamedNodeMap();

public:
	long GetLength();
	void Reset();
	CBCGPXmlNode GetItem(long index);

	CBCGPXmlNode NextNode();
    CBCGPXmlNode GetItem(const CString& strName, LPCTSTR namespaceURI = NULL);
    BOOL RemoveItem(LPCTSTR strName, LPCTSTR namespaceURI = NULL, CBCGPXmlNode* outValue = NULL);
    BOOL SetItem(const CBCGPXmlNode& newValue, CBCGPXmlNode* outValue = NULL);
};


class BCGCBPRODLLEXPORT CBCGPXmlParseErrorCollection: public CBCGPXmlBase
{
public:
	CBCGPXmlParseErrorCollection();
	CBCGPXmlParseErrorCollection(const CBCGPXmlParseErrorCollection& other);

    virtual ~CBCGPXmlParseErrorCollection();

	long GetLength();
	void Reset();
	CBCGPXmlParseError GetItem(long index);

	CBCGPXmlParseError Next();
};

#endif // !defined(AFX_BCGPXML_H__90C51E93_0257_4B10_8AC2_EF654091EA98__INCLUDED_)
