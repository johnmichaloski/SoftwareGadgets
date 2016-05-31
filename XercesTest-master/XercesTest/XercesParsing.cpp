//
// XercesParsing.cpp
//

// DISCLAIMER:
// This software was developed by U.S. Government employees as part of
// their official duties and is not subject to copyright. No warranty implied 
// or intended.


#include "StdAfx.h"
#include "XercesParsing.h"
#include <stdio.h>
#include <iostream>


CXercesParsing::CXercesParsing(void)
{
}


CXercesParsing::~CXercesParsing(void)
{
}

std::string CXercesParsing::GetAttribute(DOMNode* node, std::string attribute)
{
       XMLCh* xpathStr=XMLString::transcode(attribute.c_str()); 
	   std::string text;
	   ATLASSERT(node!=NULL);
	
	   std::cout  << XMLString::transcode( node->getNodeName() ) << std::endl;

	   DOMElement* currentElement = dynamic_cast< xercesc::DOMElement* >( node );
	   if(currentElement==NULL)
		   return text;
	   const XMLCh* xmlch_OptionA  = currentElement->getAttribute(xpathStr);
	   text = XMLString::transcode(xmlch_OptionA);
       return text;
}

std::map<std::string,std::string> CXercesParsing::GetMTConnectData(XERCES_CPP_NAMESPACE::DOMDocument*  p_DOMDocument)
{
	std::map<std::string,std::string> data;
	std::string items[3] = {std::string(".//Samples"), std::string(".//Events") , std::string(".//Condition") };
	for(int ii=0; ii<3 ; ii++)
	{
		std::vector<DOMNode*> samples = FindXPathMatches(p_DOMDocument, items[ii]);
		for(int j=0; j< samples.size(); j++)
		{
			DOMNode* pSampleHive = samples[j];                                  

			// Get each child
			DOMNodeList*      children = pSampleHive->getChildNodes();
			const  XMLSize_t nodeCount = children->getLength();
			for(XMLSize_t k=0; k< nodeCount; k++)
			{
				DOMNode* pSample = children->item(k);;
				if( pSample->getNodeType()==NULL &&  // true is not NULL
					pSample->getNodeType() != DOMNode::ELEMENT_NODE ) // is element
				{
					continue;
				}
				if(XMLString::transcode( pSample->getNodeName() )=="#text")
					continue;
				//ptime datetime;
				std::string name ;
				std::string value;
				std::string timestamp;
				std::string sequence;


				name =  GetAttribute(pSample, "name");
				if(name.empty())
					name =  GetAttribute(pSample, "dataItemId");
				if(name.empty())
					continue;

				value = XMLString::transcode(pSample->getTextContent());

				//if(items[ii]== bstr_t(".//Condition") )
				//	value =  std::string((LPCSTR) pSample->nodeName) + "."  + value  ;

				//if(_valuerenames.find(name+"."+value)!=_valuerenames.end())
				//	value=_valuerenames[name+"."+value];

				//timestamp = (LPCSTR)  GetAttribute(pSample, "timestamp");
				//sequence = (LPCSTR)  GetAttribute(pSample, "sequence");

				data[name]= value;
			}
		}

	}
	return data;
}

// XPATH  Sample
std::vector<DOMNode*>  CXercesParsing::FindXPathMatches(XERCES_CPP_NAMESPACE::DOMDocument*  p_DOMDocument, std::string element)
{
	XMLCh* xpathStr;
	std::vector<DOMNode*>  nodes ;
	try
	{
		xpathStr=XMLString::transcode(element.c_str()); // "//mstns:ConnectionMethod");
		//XERCES_CPP_NAMESPACE::DOMDocument * domdoc = (XERCES_CPP_NAMESPACE::DOMDocument *) doc.GetNode();
		XERCES_CPP_NAMESPACE::DOMElement* domroot = static_cast<XERCES_CPP_NAMESPACE::DOMElement*> (p_DOMDocument->getDocumentElement());
		XERCES_CPP_NAMESPACE::DOMXPathNSResolver* resolver=p_DOMDocument->createNSResolver(domroot);

		XERCES_CPP_NAMESPACE::DOMXPathResult* result=p_DOMDocument->evaluate(
			xpathStr,
			domroot,
			resolver,
			xercesc::DOMXPathResult::ORDERED_NODE_SNAPSHOT_TYPE,
			NULL);
		XMLSize_t nLength = result->getSnapshotLength();
		for(XMLSize_t i = 0; i < nLength; i++)
		{
			result->snapshotItem(i);
			DOMNode*  node  =  result->getNodeValue();
			std::cout  << XMLString::transcode( node->getTextContent() ) << std::endl;
			nodes.push_back( node );
		}

		result->release();
		resolver->release ();
	}
	catch(const DOMXPathException& e)
	{
		XERCES_STD_QUALIFIER cerr << "An error occurred during processing of the XPath expression. Msg is:"
			<< XERCES_STD_QUALIFIER endl
			<< XMLString::transcode(e.getMessage()) << XERCES_STD_QUALIFIER endl;
	}
	catch(const DOMException& e)
	{
		XERCES_STD_QUALIFIER cerr << "An error occurred during processing of the XPath expression. Msg is:"
			<< XERCES_STD_QUALIFIER endl
			<< XMLString::transcode(e.getMessage()) << XERCES_STD_QUALIFIER endl;
	}
	std::string str =  XMLString::transcode( xpathStr );
	XMLString::release(&xpathStr);
	return nodes;
}

void CXercesParsing::ParseTree (XERCES_CPP_NAMESPACE::DOMDocument*     xmlDoc)
{
	try {
		DOMElement* elementRoot = xmlDoc->getDocumentElement();
		if( !elementRoot ) throw(std::runtime_error( "empty XML document" ));

		// Parse XML file for tags of interest: "ComponentStream"
		// Look one level nested within "root". (child of root)

		DOMNodeList*      children = elementRoot->getChildNodes();
		const  XMLSize_t nodeCount = children->getLength();
		std::cout  << "Number nodes = " << nodeCount  << std::endl;
#if 1
		// For all nodes, children of "root" in the XML tree.

		for( XMLSize_t xx = 0; xx < nodeCount; ++xx )
		{
			DOMNode* currentNode = children->item(xx);
			if( currentNode->getNodeType() &&  // true is not NULL
				currentNode->getNodeType() == DOMNode::ELEMENT_NODE ) // is element
			{
				// Found node which is an Element. Re-cast node as element
				DOMElement* currentElement	= dynamic_cast< xercesc::DOMElement* >( currentNode );
				std::cout<< XMLString::transcode(currentElement->getTagName()) << std::endl;

				//if( XMLString::equals(currentElement->getTagName(), TAG_ApplicationSettings))
				//{
				//	// Already tested node as type element and of name "ApplicationSettings".
				//	// Read attributes of element "ApplicationSettings".
				//	const XMLCh* xmlch_OptionA	= currentElement->getAttribute(ATTR_OptionA);
				//	m_OptionA = XMLString::transcode(xmlch_OptionA);

				//	const XMLCh* xmlch_OptionB	= currentElement->getAttribute(ATTR_OptionB);
				//	m_OptionB = XMLString::transcode(xmlch_OptionB);

				//	break;  // Data found. No need to look at other elements in tree.
				//}
			}
		}
#endif
	}
	catch( xercesc::XMLException& e )
	{
		char* message = xercesc::XMLString::transcode( e.getMessage() );
		//ostringstream errBuf;
		//errBuf << "Error parsing file: " << message << flush;
		XMLString::release( &message );
	}
}