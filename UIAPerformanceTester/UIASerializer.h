/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0. 
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#pragma once

#include "stdafx.h"

class UIATextContentSerializer {
public:
	UIATextContentSerializer(IUIAutomation* client, IUIAutomationTextPattern* textPattern);
	void registerTextAttribute(const std::wstring& name, const int attrib);
	void registerElementProperty(const std::wstring& name, const int prop);
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
