/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#pragma once

#include "stdafx.h"

// All the following functions are simply wrappers for existing UIAutomation methods, made c++ friendly.
 // There is no extra functionality here.

std::wstring UIAElement_GetCurrentName(IUIAutomationElement* element);
CComPtr<IUIAutomationElement> UIA_GetFocusedElement(IUIAutomation* client);
void UIACacheRequest_AddProperty(IUIAutomationCacheRequest* cacheRequest, int prop);
CComPtr<IUIAutomationTreeWalker> UIA_GetRawViewWalker(IUIAutomation* client);
CComPtr<IUIAutomationCacheRequest> UIA_CreateCacheRequest(IUIAutomation* client);
CComPtr<IUIAutomationElement> UIAElement_BuildUpdatedCache(IUIAutomationElement* element, IUIAutomationCacheRequest* cacheRequest);
CComPtr<IUIAutomationElement> UIATreeWalker_GetParentElementBuildCache(IUIAutomationTreeWalker* treeWalker, IUIAutomationElement* child, IUIAutomationCacheRequest* cacheRequest);
CComPtr<IUIAutomationTextRange> UIATextPattern_RangeFromChild(IUIAutomationTextPattern* textPattern, IUIAutomationElement* element);
bool UIA_CompareElements(IUIAutomation* UIAClient, IUIAutomationElement* firstElement, IUIAutomationElement* secondElement);
CComVariant UIAElement_GetCachedPropertyValue(IUIAutomationElement* element, long prop);
CComVariant UIAElement_GetCurrentPropertyValue(IUIAutomationElement* element, long prop);
std::wstring UIATextRange_GetText(IUIAutomationTextRange* pTextRange, int length);
int UIATextRange_Move(IUIAutomationTextRange* pTextRange, TextUnit unit, long count);
CComVariant UIATextRange_GetAttributeValue(IUIAutomationTextRange* textRange, int attrib);
std::vector<CComVariant> UIATextRange3_GetAttributeValues(IUIAutomationTextRange3* textRange, const std::vector<int> attribs);
CComPtr<IUIAutomationTextRange> UIATextRange_Clone(IUIAutomationTextRange* pTextRange);
void UIATextRange_ExpandToEnclosingUnit(IUIAutomationTextRange* textRange, TextUnit unit);
int UIATextRange_CompareEndpoints(IUIAutomationTextRange* thisTextRange, enum TextPatternRangeEndpoint thisEndpoint, IUIAutomationTextRange* otherTextRange, enum TextPatternRangeEndpoint otherEndpoint);
void UIATextRange_MoveEndpointByRange(IUIAutomationTextRange* thisTextRange, enum TextPatternRangeEndpoint thisEndpoint, IUIAutomationTextRange* otherTextRange, enum TextPatternRangeEndpoint otherEndpoint);
CComPtr<IUIAutomationElement> UIATextRange_GetEnclosingElement(IUIAutomationTextRange* pTextRange);
CComPtr<IUIAutomationElement> UIATextRange3_GetEnclosingElementBuildCache(IUIAutomationTextRange3* pTextRange, IUIAutomationCacheRequest* cacheRequest);
CComPtr<IUIAutomationElementArray> UIATextRange_GetChildren(IUIAutomationTextRange* pTextRange);
CComPtr<IUIAutomationElementArray> UIATextRange3_GetChildrenBuildCache(IUIAutomationTextRange3* pTextRange, IUIAutomationCacheRequest* cacheRequest);
int UIAElementArray_GetLength(IUIAutomationElementArray* elementArray);
CComPtr<IUIAutomationElement> UIAElementArray_GetElement(IUIAutomationElementArray* elementArray, int index);
CComPtr<IUIAutomationTextRange> UIATextPattern_GetDocumentRange(IUIAutomationTextPattern* textPattern);
CComPtr<IUnknown> UIAElement_GetCurrentPattern(IUIAutomationElement* element, int patternID);
