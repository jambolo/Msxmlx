#include "Msxmlx.h"

namespace Msxmlx
{
//! @param    pNode        The node in question
//! @param    type        The type to test for
//!
//! @return                true, if the node has the specified type

bool IsNodeType(IXMLDOMNode * pNode, DOMNodeType type)
{
    HRESULT     hr;
    DOMNodeType nodeType;

    hr = pNode->get_nodeType(&nodeType);

    return nodeType == type;
}

//! @param    pNode        The node in question
//!
//! @return                true, if the node has the type NODE_ELEMENT

bool IsElementNode(IXMLDOMNode * pNode)
{
    return IsNodeType(pNode, NODE_ELEMENT);
}

//! This function returns the named sub-element. If the sub-element does not exist or there is an error, NULL is
//! returned.
//! @param    pElement        Element node to query
//! @param    sName            Name of the sub-node
//! @param    ppResult        Location to put the element pointer or NULL.
//!
//! @return        The HRESULT, generally S_OK if everything is ok and S_FAIL if the named node was not found.

HRESULT GetSubElement(IXMLDOMElement * pElement, char const * sName, IXMLDOMElement ** ppResult)
{
    HRESULT hr;
    CComPtr<IXMLDOMNodeList> pNodeList;
    CComPtr<IXMLDOMNode>     pSubNode;

    *ppResult = nullptr;

    hr = pElement->get_childNodes(&pNodeList);

    for (pNodeList->nextNode(&pSubNode); pSubNode; pNodeList->nextNode(&pSubNode))
    {
        if (IsElementNode(pSubNode))
        {
            CComBSTR tag;
            CComQIPtr<IXMLDOMElement> pSubElement(pSubNode);

            pSubElement->get_tagName(&tag);

            if (tag == CComBSTR(sName))
            {
                pSubElement.CopyTo(ppResult);
                return S_OK;
            }
        }

        pSubNode.Release();
    }

    return S_FALSE;
}

//! @param    pElement        Element node to query
//! @param    sName            Name of the sub-element
//! @param    ppAttributes    Location to put the attributes pointer, or NULL if there is an error.
//!
//! @return        The HRESULT, generally S_OK if everything is ok and S_FAIL if the named element was not found.

HRESULT GetSubElementAttributes(IXMLDOMElement *       pElement,
                                char const *           sName,
                                IXMLDOMNamedNodeMap ** ppAttributes)
{
    HRESULT hr;
    CComPtr<IXMLDOMElement> pSubElement;

    hr = GetSubElement(pElement, sName, &pSubElement);

    if (pSubElement)
        hr = pSubElement->get_attributes(ppAttributes);
    else
        *ppAttributes = nullptr;

    return hr;
}

//! @param    pElement        Element node to query
//! @param    sName            Name of the sub-element
//! @param    pValue            Location to put the value.
//!
//! @return        The HRESULT, generally S_OK if everything is ok and S_FAIL if the named element was not found.

HRESULT GetSubElementValue(IXMLDOMElement * pElement, char const * sName, VARIANT * pValue)
{
    HRESULT hr;
    CComPtr<IXMLDOMElement> pSubElement;

    hr = GetSubElement(pElement, sName, &pSubElement);

    if (pSubElement)
    {
        CComPtr<IXMLDOMNodeList> pNodeList;
        CComPtr<IXMLDOMNode>     pSubNode;

        hr = pSubElement->get_childNodes(&pNodeList);

        for (pNodeList->nextNode(&pSubNode); pSubNode; pNodeList->nextNode(&pSubNode))
        {
            if (IsNodeType(pSubNode, NODE_TEXT))
            {
                pSubNode->get_nodeValue(pValue);
                return S_OK;
            }

            pSubNode.Release();
        }
    }

    return hr;
}

//! @param    pElement        Element to query
//! @param    sName            Name of the attribute to get
//! @param    sDefault        Value to return if the attribute is not present. The default default value is an empty
//!                            string.
//!
//! @return        The value of the attribute (coerced to a string, if necessary)

std::string GetStringAttribute(IXMLDOMElement * pElement, char const * sName, char const * sDefault /* = ""*/)
{
    HRESULT     hr;
    CComVariant value;

    hr = pElement->getAttribute(CComBSTR(sName), &value);
    if (hr == S_OK)
    {
        USES_CONVERSION;
        hr = value.ChangeType(VT_BSTR);

        return std::string(OLE2CA(value.bstrVal));
    }
    else
    {
        return std::string(sDefault);
    }
}

//! @param    pAttributes        Attribute list
//! @param    sName            Name of the attribute to get
//! @param    sDefault        Value to return if the attribute is not present. The default default value is an empty
//!                            string.
//!
//! @return        The value of the attribute (coerced to a string, if necessary)

std::string GetStringAttribute(IXMLDOMNamedNodeMap * pAttributes, char const * sName, char const * sDefault /* = ""*/)
{
    HRESULT hr;
    CComPtr<IXMLDOMNode> pAttribute;

    hr = pAttributes->getNamedItem(CComBSTR(sName), &pAttribute);
    if (hr == S_OK)
    {
        USES_CONVERSION;
        CComVariant value;
        hr = pAttribute->get_nodeTypedValue(&value);
        hr = value.ChangeType(VT_BSTR);

        return std::string(OLE2CA(value.bstrVal));
    }
    else
    {
        return std::string(sDefault);
    }
}

//! @param    pElement        Element to query
//! @param    sName            Name of the attribute to get
//! @param    fDefault        Value to return if the attribute is not present. The default default value is 0.
//!
//! @return        The value of the attribute (coerced to a float, if necessary)

float GetFloatAttribute(IXMLDOMElement * pElement, char const * sName, float fDefault /* = 0.f*/)
{
    HRESULT     hr;
    CComVariant value;

    hr = pElement->getAttribute(CComBSTR(sName), &value);
    if (hr == S_OK)
    {
        hr = value.ChangeType(VT_R4);

        return value.fltVal;
    }
    else
    {
        return fDefault;
    }
}

//! @param    pAttributes        Attribute list
//! @param    sName            Name of the attribute to get
//! @param    fDefault        Value to return if the attribute is not present. The default default value is 0.
//!
//! @return        The value of the attribute (coerced to a float, if necessary)

float GetFloatAttribute(IXMLDOMNamedNodeMap * pAttributes, char const * sName, float fDefault /* = 0.f*/)
{
    HRESULT hr;
    CComPtr<IXMLDOMNode> pAttribute;

    hr = pAttributes->getNamedItem(CComBSTR(sName), &pAttribute);
    if (hr == S_OK)
    {
        CComVariant value;
        hr = pAttribute->get_nodeTypedValue(&value);
        hr = value.ChangeType(VT_R4);

        return value.fltVal;
    }
    else
    {
        return fDefault;
    }
}

//! @param    pElement        Element to query
//! @param    sName            Name of the attribute to get
//! @param    iDefault        Value to return if the attribute is not present. The default default value is 0.
//!
//! @return        The value of the attribute (coerced to an int, if necessary), or the default value if the attribute
//!                is not present.

int GetIntAttribute(IXMLDOMElement * pElement, char const * sName, int iDefault /* = 0*/)
{
    HRESULT     hr;
    CComVariant value;

    hr = pElement->getAttribute(CComBSTR(sName), &value);
    if (hr == S_OK)
    {
        hr = value.ChangeType(VT_INT);

        return value.intVal;
    }
    else
    {
        return iDefault;
    }
}

//! @param    pAttributes        Attribute list
//! @param    sName            Name of the attribute to get
//! @param    iDefault        Value to return if the attribute is not present. The default default value is 0.
//!
//! @return        The value of the attribute (coerced to an int, if necessary), or the default value if the attribute
//!                is not present.

int GetIntAttribute(IXMLDOMNamedNodeMap * pAttributes, char const * sName, int iDefault /* = 0*/)
{
    HRESULT hr;
    CComPtr<IXMLDOMNode> pAttribute;

    hr = pAttributes->getNamedItem(CComBSTR(sName), &pAttribute);
    if (hr == S_OK)
    {
        CComVariant value;
        hr = pAttribute->get_nodeTypedValue(&value);
        hr = value.ChangeType(VT_INT);

        return value.intVal;
    }
    else
    {
        return iDefault;
    }
}

//! @param    pElement        Element to query
//! @param    sName            Name of the attribute to get
//! @param    iDefault        Value to return if the attribute is not present. The default default value is 0.
//!
//! @return        The value of the attribute (coerced to an unsigned int, if necessary), or the default value if the
//!                attribute is not present.

uint32_t GetHexAttribute(IXMLDOMElement * pElement, char const * sName, uint32_t iDefault /* = 0*/)
{
    HRESULT     hr;
    CComVariant value;

    hr = pElement->getAttribute(CComBSTR(sName), &value);
    if (hr == S_OK)
    {
        hr = value.ChangeType(VT_BSTR);
        return wcstoul(value.bstrVal, NULL, 16);
    }
    else
    {
        return iDefault;
    }
}

//! @param    pAttributes        Attribute list
//! @param    sName            Name of the attribute to get
//! @param    iDefault        Value to return if the attribute is not present. The default default value is 0.
//!
//! @return        The value of the attribute (coerced to an unsigned int, if necessary), or the default value if the
//!                attribute is not present.

uint32_t GetHexAttribute(IXMLDOMNamedNodeMap * pAttributes, char const * sName, uint32_t iDefault /* = 0*/)
{
    HRESULT hr;
    CComPtr<IXMLDOMNode> pAttribute;

    hr = pAttributes->getNamedItem(CComBSTR(sName), &pAttribute);
    if (hr == S_OK)
    {
        CComVariant value;
        hr = pAttribute->get_nodeTypedValue(&value);
        hr = value.ChangeType(VT_UINT);

        return value.uintVal;
    }
    else
    {
        return iDefault;
    }
}

//! @param    pElement        Element to query
//! @param    sName            Name of the attribute to get
//! @param    bDefault        Value to return if the attribute is not present. The default default value is false.
//!
//! @return        The value of the attribute (coerced to a bool, if necessary), or the default value if the
//!                attribute is not present.

bool GetBoolAttribute(IXMLDOMElement * pElement, char const * sName, bool bDefault /* = false*/)
{
    HRESULT     hr;
    CComVariant value;

    hr = pElement->getAttribute(CComBSTR(sName), &value);
    if (hr == S_OK)
    {
        hr = value.ChangeType(VT_BOOL);

        return value.boolVal != 0;
    }
    else
    {
        return bDefault;
    }
}

//! @param    pAttributes        Attribute list
//! @param    sName            Name of the attribute to get
//! @param    bDefault        Value to return if the attribute is not present. The default default value is false.
//!
//! @return        The value of the attribute (coerced to an bool, if necessary), or the default value if the
//!                attribute is not present.

bool GetBoolAttribute(IXMLDOMNamedNodeMap * pAttributes, char const * sName, bool bDefault /* = false*/)
{
    HRESULT hr;
    CComPtr<IXMLDOMNode> pAttribute;

    hr = pAttributes->getNamedItem(CComBSTR(sName), &pAttribute);
    if (hr == S_OK)
    {
        CComVariant value;
        hr = pAttribute->get_nodeTypedValue(&value);
        hr = value.ChangeType(VT_BOOL);

        return value.boolVal != 0;
    }
    else
    {
        return bDefault;
    }
}

//! @param    pElement        Element to query
//! @param    sName            Name of the sub-element to get
//! @param    sDefault        Value to return if the sub-element is not present. The default default value is an empty
//!                            string.
//!
//! @return        The value of the sub-element (coerced to a string, if necessary)

std::string GetStringSubElement(IXMLDOMElement * pElement, char const * sName, char const * sDefault /* = ""*/)
{
    HRESULT     hr;
    CComVariant value;

    hr = GetSubElementValue(pElement, sName, &value);
    if (hr == S_OK)
    {
        USES_CONVERSION;
        hr = value.ChangeType(VT_BSTR);

        return std::string(OLE2CA(value.bstrVal));
    }
    else
    {
        return std::string(sDefault);
    }
}

//! @param    pElement        Element to query
//! @param    sName            Name of the sub-element to get
//! @param    fDefault        Value to return if the sub-element is not present. The default default value is 0.
//!
//! @return        The value of the sub-element (coerced to a float, if necessary)

float GetFloatSubElement(IXMLDOMElement * pElement, char const * sName, float fDefault /* = 0.f*/)
{
    HRESULT     hr;
    CComVariant value;

    hr = GetSubElementValue(pElement, sName, &value);
    if (hr == S_OK)
    {
        hr = value.ChangeType(VT_R4);

        return value.fltVal;
    }
    else
    {
        return fDefault;
    }
}

//! @param    pElement        Element to query
//! @param    sName            Name of the sub-element to get
//! @param    iDefault        Value to return if the sub-element is not present. The default default value is 0.
//!
//! @return        The value of the sub-element (coerced to an int, if necessary), or the default value if the sub-element
//!                is not present.

int GetIntSubElement(IXMLDOMElement * pElement, char const * sName, int iDefault /* = 0*/)
{
    HRESULT     hr;
    CComVariant value;

    hr = GetSubElementValue(pElement, sName, &value);
    if (hr == S_OK)
    {
        hr = value.ChangeType(VT_INT);

        return value.intVal;
    }
    else
    {
        return iDefault;
    }
}

//! @param    pElement        Element to query
//! @param    sName            Name of the sub-element to get
//! @param    iDefault        Value to return if the sub-element is not present. The default default value is 0.
//!
//! @return        The value of the sub-element (coerced to an unsigned int, if necessary), or the default value if the
//!                sub-element is not present.

uint32_t GetHexSubElement(IXMLDOMElement * pElement, char const * sName, uint32_t iDefault /* = 0*/)
{
    HRESULT     hr;
    CComVariant value;

    hr = GetSubElementValue(pElement, sName, &value);
    if (hr == S_OK)
    {
        hr = value.ChangeType(VT_BSTR);
        return wcstoul(value.bstrVal, NULL, 16);
    }
    else
    {
        return iDefault;
    }
}

//! @param    pElement        Element to query
//! @param    sName            Name of the sub-element to get
//! @param    bDefault        Value to return if the sub-element is not present. The default default value is false.
//!
//! @return        The value of the sub-element (coerced to a bool, if necessary), or the default value if the
//!                sub-element is not present.

bool GetBoolSubElement(IXMLDOMElement * pElement, char const * sName, bool bDefault /* = false*/)
{
    HRESULT     hr;
    CComVariant value;

    hr = GetSubElementValue(pElement, sName, &value);
    if (hr == S_OK)
    {
        hr = value.ChangeType(VT_BOOL);

        return value.boolVal != 0;
    }
    else
    {
        return bDefault;
    }
}

//! This function calls the specified function for each node in the specified node list. If false is
//! returned, then the function aborted the enumeration process.
//!
//! @param    pNodeList   The node list to be enumerated.
//! @param    f           The function to call for each node. See Msxmlx::ForEachNodeCB.
//! @param    context     A value that is passed to the function.
//!
//! @return        false if the function aborted the enumeration.

bool ForEachNode(IXMLDOMNodeList * pNodeList, ForEachNodeCB f)
{
    bool bContinue = true;
    CComPtr<IXMLDOMNode> pNode;

    for (pNodeList->nextNode(&pNode); pNode && bContinue; pNodeList->nextNode(&pNode))
    {
        bContinue = f(pNode);
        pNode.Release();
    }

    return bContinue;
}

//! This function calls the specified function for each element node in the specified node list. If
//! false is returned, then the function aborted the enumeration process.
//!
//! @param    pNodeList    The node whose subnodes are to be enumerated.
//! @param    f    The function to call for each element. See Msxmlx::ForEachElementCB.
//! @param    context        A value that is passed to the function.
//!
//! @return        false if the function aborted the enumeration.

bool ForEachElement(IXMLDOMNodeList * pNodeList, ForEachElementCB f)
{
    bool bContinue = true;
    CComPtr<IXMLDOMNode> pNode;

    for (pNodeList->nextNode(&pNode); pNode && bContinue; pNodeList->nextNode(&pNode))
    {
        if (IsElementNode(pNode))
            bContinue = f(CComQIPtr<IXMLDOMElement>(pNode));

        pNode.Release();
    }

    return bContinue;
}

//! This function calls the specified function for each subnode of the specified node. If false is
//! returned, then the function aborted the enumeration process.
//!
//! @param    pNode        The node whose subnodes are to be enumerated.
//! @param    f    The function to call for each subnode. See Msxmlx::ForEachNodeCB.
//! @param    context        A value that is passed to the function.
//!
//! @return        false, if the function aborted the enumeration.

bool ForEachSubNode(IXMLDOMNode * pNode, ForEachNodeCB f)
{
    HRESULT hr;
    CComPtr<IXMLDOMNodeList> pSubNodeList;

    hr = pNode->get_childNodes(&pSubNodeList);

    return ForEachNode(pSubNodeList, f);
}

//! This function calls the specified function for each element subnode of the specified node. If false
//! is returned, then the function aborted the enumeration process.
//!
//! @param    pNode        The node whose subnodes are to be enumerated.
//! @param    f    The function to call for each subnode. See Msxmlx::ForEachElementCB.
//! @param    context        A value that is passed to the function.
//!
//! @return        false, if the function aborted the enumeration.

bool ForEachSubElement(IXMLDOMNode * pNode, ForEachElementCB f)
{
    HRESULT hr;
    CComPtr<IXMLDOMNodeList> pSubNodeList;

    hr = pNode->get_childNodes(&pSubNodeList);

    return ForEachElement(pSubNodeList, f);
}

//! This function creates an element with the given name containing a single text sub-node with the given value.
//!
//! @param    pDocument        Document containing the node
//! @param    sName            Name of the node
//! @param    value            Value of the text
//! @param    ppResult        Where to put the created node
//!
//! @return        The HRESULT, generally S_OK if everything is ok and S_FAIL if the node could not be created.

HRESULT CreateTextElement(IXMLDOMDocument2 * pDocument,
                          char const * sName, VARIANT const & value,
                          IXMLDOMElement ** ppResult)
{
    HRESULT hr;
    CComPtr<IXMLDOMElement> pElement;

    if (FAILED(hr = pDocument->createElement(CComBSTR(sName), &pElement)))
        return hr;

    CComPtr<IXMLDOMText> pText;
    CComVariant          variant(value);

    variant.ChangeType(VT_BSTR, NULL);

    if (FAILED(hr = pDocument->createTextNode(V_BSTR(&variant), &pText)))
        return hr;
    if (FAILED(hr = pElement->appendChild(pText, NULL)))
        return hr;

    pElement.CopyTo(ppResult);
    return S_OK;
}
} // namespace Msxmlx
