/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#include "stdafx.h"
#include "UIAWrappers.h"
#include "UIAUtils.h"
#include "UIASerializer.h"
#include "testUtils.h"

using namespace WEX::Logging;
using namespace WEX::Common;
using namespace WEX::TestExecution;

// TAEF test module

// The UIAutomation client instance all tests use
CComPtr<IUIAutomation> UIAClient;

// TAEF module set-up
MODULE_SETUP(moduleInit) {
	// Beep to indecate start of test run (As other AT should be turned off while running)
	Beep(220, 100);
	// Initialize COM and create a UIAutomation client for the tests.
	CoInitialize(NULL);
	UIAClient = createUIAClient();
	return true;
}

// A TAEF test class for testing UI Automation in Microsoft Edge
class EdgeDocument {
	TEST_CLASS(EdgeDocument);
	HWND edgeWindow{nullptr};

	// A setup method for all tests in this class which Launches Microsoft Edge
	// And loads a specific document for each test.
	// The documents should be named className_testName.html
	// And be located in the testFiles directory under the project directory.
	// These will be copied to the target directory to be picked up by TAEF.
	TEST_METHOD_SETUP(methodInit) {
		Beep(330,100);
		WEX::Common::String testName;
		RuntimeParameters::TryGetValue(L"TestName", testName);
		testName.Replace(L"::", L"_");
		WEX::Common::String testDir;
		RuntimeParameters::TryGetValue(L"TestDeploymentDir", testDir);
		std::wstringstream URL;
		URL << L"file:///" << static_cast<const wchar_t*>(testDir) << L"\\testFiles\\" << static_cast<const wchar_t*>(testName) << L".html";
		edgeWindow=launchEdgeWithURL(URL.str().c_str());
		return true;
	}

	// A basic test to time the fetching of a document's name.
	// This is really just to ensure that Edge launches and a test runs.
	TEST_METHOD(DocumentName)
	{
		BEGIN_TEST_METHOD_PROPERTIES()
			TEST_METHOD_PROPERTY(L"Description", L"Times fetching of document name")
		END_TEST_METHOD_PROPERTIES()
		CComPtr<IUIAutomationElement> document=locateEdgeDocumentUIAElement(UIAClient,edgeWindow,L"Edge Test Document");
		CComVariant nameVal;
		Sleep(2000);
		verifyTakesLessThan(L"IUIAutomationElement::get_currentPropertyValue", 0.005, [&] {
			nameVal = UIAElement_GetCurrentPropertyValue(document,UIA_NamePropertyId);
		});
		VERIFY_ARE_EQUAL(L"Edge Test Document", VariantToString(nameVal));
	}

	// A test to time the serialization of a horizontal navbar of links
	// The second line of this document contains a horizontal navbar made up of a list of 7 links. 
	TEST_METHOD(horizontalNavbar)
	{
		BEGIN_TEST_METHOD_PROPERTIES()
			TEST_METHOD_PROPERTY(L"Description", L"Times fetching a line of links")
			//TEST_METHOD_PROPERTY(L"Ignore", L"True")
		END_TEST_METHOD_PROPERTIES()
		// try and locate the loaded document 
		CComPtr<IUIAutomationElement> document = locateEdgeDocumentUIAElement(UIAClient, edgeWindow, L"Horizontal Navbar");
		// Get the document's textPattern, and the textRange spanning the entire document
		CComQIPtr<IUIAutomationTextPattern> textPattern = UIAElement_GetCurrentPattern(document, UIA_TextPatternId);
		if (!textPattern) VERIFY_FAIL(L"document element has no text pattern");
		// Fetch the range for the document 
		CComPtr<IUIAutomationTextRange> textRange = UIATextPattern_GetDocumentRange(textPattern);
		// collapse the range to the start
		UIATextRange_MoveEndpointByRange(textRange, TextPatternRangeEndpoint_End, textRange, TextPatternRangeEndpoint_Start);
		// Move to the second line
		UIATextRange_Move(textRange, TextUnit_Line, 1);
		// Expand to line (covering the horizontal navbar)
		UIATextRange_ExpandToEnclosingUnit(textRange, TextUnit_Line);
		//Prepare a UIASerializer, setting the wanted text attributes and element properties
		auto serializer = UIATextContentSerializer(UIAClient, textPattern);
		serializer.registerTextAttribute(L"fontName", UIA_FontNameAttributeId);
		serializer.registerTextAttribute(L"fontSize", UIA_FontSizeAttributeId);
		serializer.registerElementProperty(L"name", UIA_NamePropertyId);
		serializer.registerElementProperty(L"controlType", UIA_LocalizedControlTypePropertyId);
		std::wstring content;
		// sleep for a while to ensure the document is not using much CPU
		Sleep(2000); 
		// Time the serialization of the horizontal navbar 
		verifyTakesLessThan(L"UIATextContentSerializer::serializeTextcontent", 0.3, [&] {
			content = serializer.serializeTextcontent(textRange);
		});
		// Write the serialized content to file
		{
			std::wofstream outFile("edgeDocument_horizontalNavbar_output.xml");
			outFile << content;
		}
		// Double check the content with the expected result
		std::wifstream inFile("testFiles\\edgeDocument_horizontalNavbar_expected.xml");
		std::wstringstream s;
		s << inFile.rdbuf();
		VERIFY_ARE_EQUAL(0, content.compare(s.str()), L"Comparing serialized content");
	}

	TEST_METHOD_CLEANUP(methodCleanup) {
		// Close the Edge window
		SendMessage(edgeWindow,WM_CLOSE,0,0);
		Beep(660,100);
		return true;
	}

};

MODULE_CLEANUP(moduleCleanup) {
	CoUninitialize();
	// Beep to indicate the end of the tests. 
	Beep(880, 100);
	return true;
}
