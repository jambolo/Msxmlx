#pragma once

#if !defined(MSXMLX_MSXMLX_H)
#define MSXMLX_MSXMLX_H

#include <atlbase.h>
#include <atlcomcli.h>
#include <cstdint>
#include <functional>
#include <msxml2.h>
#include <string>

//! Miscellaneous functions supporting MSXML.

namespace Msxmlx
{
/********************************************************************************************************************/
/*												N A V I G A T I O N													*/
/********************************************************************************************************************/

//! Returns true if the node is an element node.
bool IsElementNode(IXMLDOMNode * pNode);

//! Returns true if the node is the specified type.
bool IsNodeType(IXMLDOMNode * pNode, DOMNodeType type);

//! Returns the named sub-element, or NULL if error or not found.
HRESULT GetSubElement(IXMLDOMElement * pElement, char const * sName, IXMLDOMElement ** ppResult);

//! Returns attributes for a specific sub-element or NULL if error or not found.
HRESULT GetSubElementAttributes(IXMLDOMElement *       pElement,
                                char const *           sName,
                                IXMLDOMNamedNodeMap ** ppAttributes);

//! Creates a node with a text sub-node. Returns an HRESULT.
HRESULT CreateTextElement(IXMLDOMDocument2 * pDocument,
                          char const * sName, VARIANT const & value,
                          IXMLDOMElement ** ppResult);

/********************************************************************************************************************/
/*											A T T R I B U T E   V A L U E S											*/
/********************************************************************************************************************/

//! Returns the value of a string attribute (or a default value, if the attribute is not present).
std::string GetStringAttribute(IXMLDOMElement * pElement, char const * sName, char const * sDefault = "");

//! Returns the value of a string attribute (or a default value, if the attribute is not present).
std::string GetStringAttribute(IXMLDOMNamedNodeMap * pAttributes,
                               char const *          sName,
                               char const *          sDefault = "");

//! Returns the value of a float attribute (or a default value, if the attribute is not present).
float GetFloatAttribute(IXMLDOMElement * pElement, char const * sName, float fDefault = 0.f);

//! Returns the value of a float attribute (or a default value, if the attribute is not present).
float GetFloatAttribute(IXMLDOMNamedNodeMap * pAttributes, char const * sName, float fDefault = 0.f);

//! Returns the value of an integer attribute (or a default value, if the attribute is not present).
int GetIntAttribute(IXMLDOMElement * pElement, char const * sName, int iDefault = 0);

//! Returns the value of an integer attribute (or a default value, if the attribute is not present).
int GetIntAttribute(IXMLDOMNamedNodeMap * pAttributes, char const * sName, int iDefault = 0);

//! Returns the value of an integer attribute (or a default value, if the attribute is not present).
uint32_t GetHexAttribute(IXMLDOMElement * pElement, char const * sName, uint32_t iDefault = 0);

//! Returns the value of an integer attribute (or a default value, if the attribute is not present).
uint32_t GetHexAttribute(IXMLDOMNamedNodeMap * pAttributes, char const * sName, uint32_t iDefault = 0);

//! Returns the value of an integer attribute (or a default value, if the attribute is not present).
bool GetBoolAttribute(IXMLDOMElement * pElement, char const * sName, bool bDefault = false);

//! Returns the value of an integer attribute (or a default value, if the attribute is not present).
bool GetBoolAttribute(IXMLDOMNamedNodeMap * pAttributes, char const * sName, bool bDefault = false);

/********************************************************************************************************************/
/*											E L E M E N T   V A L U E S												*/
/********************************************************************************************************************/

//! Returns the first occurrence of text for a specific sub-element or nothing if error or not found.
HRESULT GetSubElementValue(IXMLDOMElement * pElement, char const * sName, VARIANT * pValue);

//! Returns the value of a string sub-element (or a default value, if the sub-element is not present).
std::string GetStringSubElement(IXMLDOMElement * pElement, char const * sName, char const * sDefault = "");

//! Returns the value of a float sub-element (or a default value, if the sub-element is not present).
float GetFloatSubElement(IXMLDOMElement * pElement, char const * sName, float fDefault = 0.f);

//! Returns the value of an integer sub-element (or a default value, if the sub-element is not present).
int GetIntSubElement(IXMLDOMElement * pElement, char const * sName, int iDefault = 0);

//! Returns the value of an integer sub-element (or a default value, if the sub-element is not present).
uint32_t GetHexSubElement(IXMLDOMElement * pElement, char const * sName, uint32_t iDefault = 0);

//! Returns the value of an integer sub-element (or a default value, if the sub-element is not present).
bool GetBoolSubElement(IXMLDOMElement * pElement, char const * sName, bool bDefault = false);

/********************************************************************************************************************/
/*												E N U M E R A T I O N												*/
/********************************************************************************************************************/

//! Callback function prototype for ForEachSubNode() and ForEachNode().
//!
//! This function is called by ForEachSubNode() for each subnode of a node or by ForEachNode() for each node in
//! the node list.
//!
//! @param	pNode		Node
//!
//! @return		@c false if the enumeration should be aborted

using ForEachNodeCB = std::function<bool(IXMLDOMNode * pElement)>;

//! Callback function prototype for ForEachSubElement().
//!
//! This function is called by ForEachSubElement() for each element subnode of a node.
//!
//! @param	pElement    Element
//!
//! @return		@c false if the enumeration should be aborted

using ForEachElementCB = std::function<bool(IXMLDOMElement * pElement)>;

//! Calls a function for each node in the list and returns false if the function was aborted.
bool ForEachNode(IXMLDOMNodeList * pNodeList, ForEachNodeCB f);

//! Calls a function for each element in the list and returns false if the function was aborted.
bool ForEachElement(IXMLDOMNodeList * pNodeList, ForEachElementCB f);

//! Calls a function for each subnode and returns false if the function was aborted.
bool ForEachSubNode(IXMLDOMNode * pNode, ForEachNodeCB f);

//! Calls a function for each element subnode and returns false if the function was aborted.
bool ForEachSubElement(IXMLDOMNode * pNode, ForEachElementCB f);
} // namespace Msxmlx

#endif // !defined(MSXMLX_MSXMLX_H)
