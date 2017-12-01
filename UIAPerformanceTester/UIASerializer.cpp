/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#include "stdafx.h"
#include "UIAWrappers.h"
#include "UIAUtils.h"
#include "UIASerializer.h"

std::wstring VariantToString(VARIANT& val) {
	switch (val.vt) {
	case VT_BSTR:
		return (val.bstrVal ? val.bstrVal : L"");
	case VT_I4:
		std::wostringstream s;
		s << val.lVal;
		return s.str();
	}
	return L"";
}

UIATextContentSerializer::UIATextContentSerializer(IUIAutomation* client, IUIAutomationTextPattern* textPattern) :
	client(client), textPattern(textPattern),
	treeWalker(UIA_GetRawViewWalker(client)),
	cacheRequest(UIA_CreateCacheRequest(client))
{}

void UIATextContentSerializer::registerTextAttribute(const std::wstring& name, const int attrib) {
	textAttributes.push_back(attrib);
	textAttributeLabels[attrib] = name;
}

void UIATextContentSerializer::registerElementProperty(const std::wstring& name, const int prop) {
	UIACacheRequest_AddProperty(cacheRequest, prop);
	elementProperties[prop] = name;
}

std::wstring UIATextContentSerializer::_escapeText(const std::wstring& text) {
	std::wstring newText;
	for (auto c : text) {
		if (c == L'<') {
			newText += L"&lt;";
		}
		else if (c == L'>') {
			newText += L"&gt;";
		}
		else {
			newText += c;
		}
	}
	return newText;
}

std::wstring UIATextContentSerializer::_escapeAttributeValue(const std::wstring& value) {
	std::wstring newValue;
	for (auto c : value) {
		if (c == L'"') {
			newValue += L"\\\"";
		}
		else {
			newValue += c;
		}
	}
	return newValue;
}

std::wstring UIATextContentSerializer::serializeTextcontent(IUIAutomationTextRange* textRange) {
	std::wostringstream log;
	log << L"<root>" << std::endl;
	_serializeTextContent(log, textRange);
	log << L"</root>" << std::endl;
	return log.str();
}

void UIATextContentSerializer::_serializeElementStart(std::wostream& log, IUIAutomationElement* element, bool clippedStart) {
	log << L"<UIAElement ";
	for (auto prop : elementProperties) {
		std::wstring& name = prop.second;
		std::wstring val = VariantToString(UIAElement_GetCachedPropertyValue(element, prop.first));
		log << name << L"=\"" << _escapeAttributeValue(val) << L"\" ";
	}
	log << L">" << std::endl;
	if (clippedStart) log << L"..." << std::endl;
}

void UIATextContentSerializer::_serializeTextRangeByUnit(std::wostream& log, IUIAutomationTextRange* textRange, TextUnit unit) {
	for (auto textChunk : splitUIATextRangeByUnit(textRange, unit)) {
		_serializeTextRange(log, textChunk);
	}
}

void UIATextContentSerializer::_serializeTextRange(std::wostream& log, IUIAutomationTextRange* textRange) {
	std::wstring text = UIATextRange_GetText(textRange, -1);
	if (text.empty()) return;
	log << L"<textRange ";
	auto attribValues = UIATextRangeX_GetAttributeValues(textRange, textAttributes);
	size_t attribCount = attribValues.size();
	for (auto attribIndex = 0; attribIndex<attribCount; ++attribIndex) {
		std::wstring& name = textAttributeLabels[textAttributes[attribIndex]];
		std::wstring val = VariantToString(attribValues[attribIndex]);
		log << L"\"" << name << L"\":" << L"\"" << val << L"\" ";
	}
	log << L">" << std::endl;
	log << _escapeText(text) << std::endl;
	log << L"</textRange>" << std::endl;
}

void UIATextContentSerializer::_serializeElementEnd(std::wostream& log, IUIAutomationElement* element, bool clippedEnd) {
	if (clippedEnd) log << L"..." << std::endl;
	log << L"</UIAElement>" << std::endl;
}

void UIATextContentSerializer::_serializeTextContent(std::wostream& log, IUIAutomationTextRange* textRange, IUIAutomationElement* rootElement) {
	std::vector<std::pair<CComPtr<IUIAutomationElement>, CComPtr<IUIAutomationTextRange>>> parents;
	auto enclosingElement = UIATextRangeX_GetEnclosingElementBuildCache(textRange, cacheRequest);
	auto parent = enclosingElement;
	while (parent && (!rootElement || !UIA_CompareElements(client, parent, rootElement))) {
		auto parentRange = UIATextPattern_RangeFromChild(textPattern, parent);
		if (!parentRange) break;
		parents.emplace_back(std::make_pair(parent, parentRange));
		parent = UIATreeWalker_GetParentElementBuildCache(treeWalker, parent, cacheRequest);
	}
	bool clippedStart = true;
	for (auto i = parents.rbegin(); i != parents.rend(); ++i) {
		if (clippedStart) clippedStart = UIATextRange_CompareEndpoints(i->second, TextPatternRangeEndpoint_Start, textRange, TextPatternRangeEndpoint_Start)<0;
		_serializeElementStart(log, i->first, clippedStart);
	}
	auto children = UIATextRangeX_GetChildrenBuildCache(textRange, cacheRequest);
	size_t childCount = children.size();
	if (childCount == 0) {
		_serializeTextRangeByUnit(log, textRange, TextUnit_Format);
	}
	else {
		size_t lastChildIndex = childCount - 1;
		auto tempRange = UIATextRange_Clone(textRange);
		UIATextRange_MoveEndpointByRange(tempRange, TextPatternRangeEndpoint_End, tempRange, TextPatternRangeEndpoint_Start);
		for (int childIndex = 0; childIndex<childCount; ++childIndex) {
			bool clippedStart = false;
			bool clippedEnd = false;
			auto& childElement = children[childIndex];
			if (!childElement) continue;
			if (UIA_CompareElements(client, childElement, enclosingElement)) continue;
			auto childRange = UIATextPattern_RangeFromChild(textPattern, childElement);
			if (!childRange) continue;
			if (childIndex == lastChildIndex) {
				if (UIATextRange_CompareEndpoints(childRange, TextPatternRangeEndpoint_Start, textRange, TextPatternRangeEndpoint_End) >= 0) break;
				int lastChildEndDelta = UIATextRange_CompareEndpoints(childRange, TextPatternRangeEndpoint_End, textRange, TextPatternRangeEndpoint_End);
				if (lastChildEndDelta>0) {
					UIATextRange_MoveEndpointByRange(childRange, TextPatternRangeEndpoint_End, textRange, TextPatternRangeEndpoint_End);
					clippedEnd = true;
				}
			}
			int childStartDelta = UIATextRange_CompareEndpoints(childRange, TextPatternRangeEndpoint_Start, tempRange, TextPatternRangeEndpoint_End);
			if (childStartDelta>0) {
				UIATextRange_MoveEndpointByRange(tempRange, TextPatternRangeEndpoint_End, childRange, TextPatternRangeEndpoint_Start);
				_serializeTextRangeByUnit(log, tempRange, TextUnit_Format);
			}
			else if (childStartDelta<0) {
				UIATextRange_MoveEndpointByRange(childRange, TextPatternRangeEndpoint_Start, tempRange, TextPatternRangeEndpoint_End);
				clippedStart = true;
			}
			if (childIndex == 0 || childIndex == lastChildIndex) {
				if (UIATextRange_CompareEndpoints(childRange, TextPatternRangeEndpoint_Start, childRange, TextPatternRangeEndpoint_End) == 0) continue;
			}
			_serializeElementStart(log, childElement, clippedStart);
			_serializeTextContent(log, childRange, childElement);
			_serializeElementEnd(log, childElement, clippedEnd);
			UIATextRange_MoveEndpointByRange(tempRange, TextPatternRangeEndpoint_Start, childRange, TextPatternRangeEndpoint_End);
		}
		if (UIATextRange_CompareEndpoints(tempRange, TextPatternRangeEndpoint_Start, textRange, TextPatternRangeEndpoint_End)<0) {
			UIATextRange_MoveEndpointByRange(tempRange, TextPatternRangeEndpoint_End, textRange, TextPatternRangeEndpoint_End);
			_serializeTextRangeByUnit(log, tempRange, TextUnit_Format);
		}
	}
	bool clippedEnd = false;
	for (auto i = parents.begin(); i != parents.end(); ++i) {
		if (!clippedEnd) clippedEnd = UIATextRange_CompareEndpoints(i->second, TextPatternRangeEndpoint_End, textRange, TextPatternRangeEndpoint_End)>0;
		_serializeElementEnd(log, parent, clippedEnd);
	}
}
