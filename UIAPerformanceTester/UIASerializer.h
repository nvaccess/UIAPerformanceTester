/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0. 
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#pragma once

#include "stdafx.h"

// Allows the fetching of  content from a UIA text range, serialized as xml.
class UIATextContentSerializer {
public:
	// Constructor
	UIATextContentSerializer(IUIAutomation* client, IUIAutomationTextPattern* textPattern);
	// Requests the UIASerializer to include the given UIA text attribute in its serialized XML for each text range.
	void registerTextAttribute(const std::wstring& name, const int attrib);
	// Requests the UIASerializer to include the given UIA element property in its serialized xml.
	void registerElementProperty(const std::wstring& name, const int prop);
	// Walks the given text range and descendants, creating a serialized xml representation of the content.
	std::wstring serializeTextcontent(IUIAutomationTextRange* textRange);
private:
	CComPtr<IUIAutomation> client;
	CComPtr<IUIAutomationTextPattern> textPattern;
	CComPtr<IUIAutomationTreeWalker> treeWalker;
	CComPtr<IUIAutomationCacheRequest> cacheRequest;
	std::vector<int> textAttributes;
	std::map<int, std::wstring> textAttributeLabels;
	std::map<int, std::wstring> elementProperties;
	std::wstring _escapeText(const std::wstring& text);
	std::wstring _escapeAttributeValue(const std::wstring& value);
	void _serializeElementStart(std::wostream& log, IUIAutomationElement* element, bool clippedStart = false);
	void _serializeTextRange(std::wostream& log, IUIAutomationTextRange* textRange);
	void _serializeTextRangeByUnit(std::wostream& log, IUIAutomationTextRange* textRange, TextUnit unit);
	void _serializeElementEnd(std::wostream& log, IUIAutomationElement* element, bool clippedEnd = false);
	void _serializeTextContent(std::wostream& log, IUIAutomationTextRange* textRange, IUIAutomationElement* rootElement = nullptr);
};
