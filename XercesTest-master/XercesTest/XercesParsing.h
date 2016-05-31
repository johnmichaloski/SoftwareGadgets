//
// XercesParsing.h
// 

// DISCLAIMER:
// This software was developed by U.S. Government employees as part of
// their official duties and is not subject to copyright. No warranty implied 
// or intended.


#pragma once
#include <string>
#include <vector>
#include <map>

#define XERCES_STATIC_LIBRARY

#include <xercesc/util/PlatformUtils.hpp>
#include <xercesc/dom/DOM.hpp>
XERCES_CPP_NAMESPACE_USE;

class CXercesParsing
{
public:
	CXercesParsing(void);
	~CXercesParsing(void);
	std::vector<DOMNode*> FindXPathMatches(XERCES_CPP_NAMESPACE::DOMDocument*  p_DOMDocument, std::string element);
	static void ParseTree (XERCES_CPP_NAMESPACE::DOMDocument*     xmlDoc);
	std::string GetAttribute(DOMNode* node, std::string attribute);
	std::map<std::string,std::string> GetMTConnectData(XERCES_CPP_NAMESPACE::DOMDocument*  p_DOMDocument);
	/////////////////////////////////////////////////////////////////
	  char* m_OptionA;
	   char* m_OptionB;
	 

	   XMLCh* TAG_root;
	 
	   XMLCh* TAG_ApplicationSettings;
	   XMLCh* ATTR_OptionA;
	   XMLCh* ATTR_OptionB;
};

