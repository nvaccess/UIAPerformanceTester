/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#include "stdafx.h"
#include "uiaWrappers.h"
#include "UIAUtils.h"

using namespace WEX::Logging;
using namespace WEX::Common;

CComPtr<IUIAutomationElement> UIATextRangeX_GetEnclosingElementBuildCache(IUIAutomationTextRange* textRange, IUIAutomationCacheRequest* cacheRequest) {
	CComQIPtr<IUIAutomationTextRange3> textRange3=textRange;
	if(textRange3) {
		return UIATextRange3_GetEnclosingElementBuildCache(textRange3,cacheRequest);
	}
	Log::Warning(L"IUIAutomationTextRange3::GetEnclosingElementBuildCache not available. Emulating");
	auto enclosingElement=UIATextRange_GetEnclosingElement(textRange);
	return UIAElement_BuildUpdatedCache(enclosingElement,cacheRequest);
}

std::vector<CComPtr<IUIAutomationElement>> UIATextRangeX_GetChildrenBuildCache(IUIAutomationTextRange* textRange, IUIAutomationCacheRequest* cacheRequest) {
	std::vector<CComPtr<IUIAutomationElement>> children;
	CComQIPtr<IUIAutomationTextRange3> textRange3=textRange;
	auto childrenArray=textRange3?UIATextRange3_GetChildrenBuildCache(textRange3,cacheRequest):UIATextRange_GetChildren(textRange);
	if(!childrenArray) return children;
	int childCount=UIAElementArray_GetLength(childrenArray);
	if(!textRange3) Log::Warning(L"IUIAutomationTextRange3::GetChildrenBuildCache not available. Emulating");
	for(int index=0;index<childCount;++index) {
		if(textRange3) {
			children.emplace_back(UIAElementArray_GetElement(childrenArray,index));
		} else {
			children.emplace_back(UIAElement_BuildUpdatedCache(UIAElementArray_GetElement(childrenArray, index),cacheRequest));
		}
	}
	return children;
}

std::vector<CComVariant> UIATextRangeX_GetAttributeValues(IUIAutomationTextRange* textRange, const std::vector<int>& attribs) {
	CComQIPtr<IUIAutomationTextRange3> textRange3=textRange;
	if(textRange3) {
		return UIATextRange3_GetAttributeValues(textRange3,attribs);
	}
	Log::Warning(L"IUIAutomationTextRange3::GetAttributeValues not available. Emulating");
	std::vector<CComVariant> values;
	for (auto attrib : attribs) {
		values.emplace_back(UIATextRange_GetAttributeValue(textRange,attrib));
	}
	return values;
}

std::vector<CComPtr<IUIAutomationTextRange>> splitUIATextRangeByUnit(IUIAutomationTextRange* pTextRange, TextUnit unit) {
	std::vector<CComPtr<IUIAutomationTextRange>> textRanges;
	auto tempRange=UIATextRange_Clone(pTextRange);
	UIATextRange_MoveEndpointByRange(tempRange,TextPatternRangeEndpoint_End, pTextRange, TextPatternRangeEndpoint_Start);
	auto endRange = UIATextRange_Clone(tempRange);
	while(UIATextRange_Move(endRange,unit,1)>0) {
		UIATextRange_MoveEndpointByRange(tempRange,TextPatternRangeEndpoint_End, endRange, TextPatternRangeEndpoint_Start);
		bool passedEnd = UIATextRange_CompareEndpoints(tempRange,TextPatternRangeEndpoint_End, pTextRange, TextPatternRangeEndpoint_End)>0;
		if(passedEnd) {
			UIATextRange_MoveEndpointByRange(tempRange,TextPatternRangeEndpoint_End, pTextRange, TextPatternRangeEndpoint_End);
		}
		textRanges.emplace_back(UIATextRange_Clone(tempRange));
		if(passedEnd) {
			return textRanges;
		}
		UIATextRange_MoveEndpointByRange(tempRange,TextPatternRangeEndpoint_Start, tempRange, TextPatternRangeEndpoint_End);
	}
	if(UIATextRange_CompareEndpoints(tempRange,TextPatternRangeEndpoint_End, pTextRange, TextPatternRangeEndpoint_End)<0) {
		UIATextRange_MoveEndpointByRange(tempRange,TextPatternRangeEndpoint_End, pTextRange, TextPatternRangeEndpoint_End);
		textRanges.emplace_back(UIATextRange_Clone(tempRange));
	}
	return textRanges;
}
