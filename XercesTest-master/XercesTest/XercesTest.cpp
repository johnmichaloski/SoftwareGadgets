// XercesTest.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
// tell Xerces that you are not interested in linking to its .dll components at runtime 
//  https://code.google.com/p/uncwddas/wiki/Xercesc_Static_Linking
#define XERCES_STATIC_LIBRARY

#include <xercesc/util/PlatformUtils.hpp>

// DOM (if you want SAX, then that's a different include)
#include <xercesc/dom/DOM.hpp>
#include <xercesc/parsers/XercesDOMParser.hpp>

// Define namespace symbols (Otherwise we'd have to prefix Xerces code with 
// "XERCES_CPP_NAMESPACE::")
//namespace xerces = XERCES_CPP_NAMESPACE ;
using namespace xercesc;


#if defined(WIN64) && defined( _DEBUG)
#pragma message( "DEBUG x64" )
#pragma comment(lib, "xerces-c_static_3D.lib")
#pragma comment(lib, "xerces-c_3D.lib")
#elif !defined( _DEBUG) && defined(WIN64)
#pragma message( "RELEASE x64" )
#pragma comment(lib, "xerces-c_static_3.lib")
#elif defined(WIN32) && defined( _DEBUG)
#pragma message( "DEBUG x32" )
#pragma comment(lib, "xerces-c_static_3D.lib")
//#pragma comment(lib, "xerces-c_3D.lib")
#elif !defined( _DEBUG) && defined(WIN32)
#pragma message( "RELEASE x32" )
#pragma comment(lib, "xerces-c_static_3.lib")
#endif

#include "StdstringFcn.h"
#include "File.h"
#include "XercesParsing.h"

int _tmain(int argc, _TCHAR* argv[])
{

	// Initilize Xerces.
    XMLPlatformUtils::Initialize();

    // Pointer to our DOMImplementation.
    XERCES_CPP_NAMESPACE::DOMImplementation*    p_DOMImplementation = NULL;

    // Get the DOM Implementation (used for creating DOMDocuments).
    // Also see: http://www.w3.org/TR/2000/REC-DOM-Level-2-Core-20001113/core.html
    p_DOMImplementation = DOMImplementationRegistry::getDOMImplementation(
             XMLString::transcode("core"));


	std::string configFile = "C:\\Users\\michalos\\Documents\\Visual Studio 2010\\Projects\\XercesTest\\XercesTest\\Win32\\Debug\\TestData.xml";
	//std::string configFile = "C:\\Users\\michalos\\Documents\\Visual Studio 2010\\Projects\\XercesTest\\XercesTest\\Win32\\Debug\\Sample.xml";
	XercesDOMParser * m_ConfigFileParser = new XercesDOMParser;
	m_ConfigFileParser->setValidationScheme( XercesDOMParser::Val_Never );
	m_ConfigFileParser->setDoNamespaces( false );
	m_ConfigFileParser->setDoSchema( false );
	m_ConfigFileParser->setLoadExternalDTD( false );
	m_ConfigFileParser->parse( configFile.c_str() );

	// no need to free this pointer - owned by the parent parser object
	XERCES_CPP_NAMESPACE::DOMDocument* xmlDoc = m_ConfigFileParser->getDocument();
	CXercesParsing parser;
	//parser.FindXPathMatches(xmlDoc, "Header");
	std::map<std::string,std::string> values = parser.GetMTConnectData(xmlDoc);
//	CXercesParsing::ParseTree (xmlDoc);

    // Cleanup.
    delete m_ConfigFileParser;
    XMLPlatformUtils::Terminate();
	return 0;
}

