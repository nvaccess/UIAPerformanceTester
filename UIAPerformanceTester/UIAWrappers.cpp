/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#include "stdafx.h"
#include "UIAWrappers.h"

using namespace WEX::Logging;
using namespace WEX::Common;

CComPtr<IUnknown> UIAElement_GetCurrentPattern(IUIAutomationElement* element, int patternID) {
	HRESULT res;
	CComPtr<IUnknown> pattern;
	res=element->GetCurrentPattern(patternID,&pattern);
	if (FAILED(res)||!pattern) VERIFY_FAIL(L"IUIAutomationElement::GetCurrentPattern");
	return pattern;
}

CComPtr<IUIAutomationTextRange> UIATextPattern_GetDocumentRange(IUIAutomationTextPattern* textPattern) {
	HRESULT res;
	CComPtr<IUIAutomationTextRange> documentRange;
	res=textPattern->get_DocumentRange(&documentRange);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextPattern::get_documentRange");
	return documentRange;
}

std::wstring UIAElement_GetCurrentName(IUIAutomationElement* element) {
	HRESULT res;
	CComBSTR name;
	res = element->get_CurrentName(&name);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationElement::get_CurrentName");
	return (name ? std::wstring(name) : L"");
}

CComPtr<IUIAutomationElement> UIA_GetFocusedElement(IUIAutomation* client) {
	HRESULT res;
	CComPtr<IUIAutomationElement> focus;
	res = client->GetFocusedElement(&focus);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomation::GetFocusedElement");
	return focus;
}

void UIACacheRequest_AddProperty(IUIAutomationCacheRequest* cacheRequest, int prop) {
	HRESULT res;
	res = cacheRequest->AddProperty(prop);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationCacheRequest::AddProperty");
}

CComPtr<IUIAutomationTreeWalker> UIA_GetRawViewWalker(IUIAutomation* client) {
	HRESULT res;
	CComPtr<IUIAutomationTreeWalker> treeWalker;
	res = client->get_RawViewWalker(&treeWalker);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomation::get_RawViewWalker");
	return treeWalker;
}

CComPtr<IUIAutomationCacheRequest> UIA_CreateCacheRequest(IUIAutomation* client) {
	HRESULT res;
	CComPtr<IUIAutomationCacheRequest> cacheRequest;
	res = client->CreateCacheRequest(&cacheRequest);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomation::CreateCacheRequest");
	return cacheRequest;
}

CComPtr<IUIAutomationElement> UIAElement_BuildUpdatedCache(IUIAutomationElement* element, IUIAutomationCacheRequest* cacheRequest) {
	HRESULT res;
	CComPtr<IUIAutomationElement> newElement;
	res = element->BuildUpdatedCache(cacheRequest, &newElement);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationElement::BuildUpdatedCache");
	return newElement;
}

CComPtr<IUIAutomationElement> UIATreeWalker_GetParentElementBuildCache(IUIAutomationTreeWalker* treeWalker, IUIAutomationElement* child, IUIAutomationCacheRequest* cacheRequest) {
	CComPtr<IUIAutomationElement> parent;
	treeWalker->GetParentElementBuildCache(child, cacheRequest, &parent);
	return parent;
}

CComPtr<IUIAutomationTextRange> UIATextPattern_RangeFromChild(IUIAutomationTextPattern* textPattern, IUIAutomationElement* element) {
	CComPtr<IUIAutomationTextRange> childRange;
	textPattern->RangeFromChild(element, &childRange);
	return childRange;
}

bool UIA_CompareElements(IUIAutomation* UIAClient, IUIAutomationElement* firstElement, IUIAutomationElement* secondElement) {
	HRESULT res;
	BOOL isSame = false;
	res = UIAClient->CompareElements(firstElement, secondElement, &isSame);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomation::CompareElements");
	return static_cast<bool>(isSame);
}

CComVariant UIAElement_GetCachedPropertyValue(IUIAutomationElement* element, long prop) {
	HRESULT res;
	CComVariant val;
	res = element->GetCachedPropertyValue(prop, &val);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationElement::GetCachedPropertyValue");
	return val;
}

CComVariant UIAElement_GetCurrentPropertyValue(IUIAutomationElement* element, long prop) {
	HRESULT res;
	CComVariant val;
	res = element->GetCurrentPropertyValue(prop, &val);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationElement::GetCurrentPropertyValue");
	return val;
}

std::wstring UIATextRange_GetText(IUIAutomationTextRange* pTextRange, int length) {
	HRESULT res;
	CComBSTR _text;
	res = pTextRange->GetText(length, &_text);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange::getText");
	return _text?std::wstring(_text):L"";
}

void UIATextRange_ExpandToEnclosingUnit(IUIAutomationTextRange* textRange, TextUnit unit) {
	HRESULT res;
	res=textRange->ExpandToEnclosingUnit(unit);
	if(FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange::ExpandToEnclosingUnit");
}

int UIATextRange_Move(IUIAutomationTextRange* pTextRange, TextUnit unit, long count) {
	HRESULT res;
	int unitsMoved = 0;
	res = pTextRange->Move(unit, count, &unitsMoved);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange::Move");
	return unitsMoved;
}

CComVariant UIATextRange_GetAttributeValue(IUIAutomationTextRange* textRange, int attrib) {
	CComVariant val;
	textRange->GetAttributeValue(attrib, &val);
	return val;
}

std::vector<CComVariant> UIATextRange3_GetAttributeValues(IUIAutomationTextRange3* textRange, const std::vector<int> attribs) {
	HRESULT res;
	CComSafeArray<VARIANT> psa;
	res = textRange->GetAttributeValues(attribs.data(), static_cast<int>(attribs.size()), &(psa.m_psa));
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange3::GetAttributeValues");
	long count = psa.GetCount();
	std::vector<CComVariant> values;
	for (long i = 0; i<count; ++i) {
		VARIANT& val = psa.GetAt(i);
		values.push_back(val);
	}
	return values;
}

CComPtr<IUIAutomationTextRange> UIATextRange_Clone(IUIAutomationTextRange* pTextRange) {
	HRESULT res;
	CComPtr<IUIAutomationTextRange> cloned;
	res = pTextRange->Clone(&cloned);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange::Clone");
	return cloned;
}

int UIATextRange_CompareEndpoints(IUIAutomationTextRange* thisTextRange, enum TextPatternRangeEndpoint thisEndpoint, IUIAutomationTextRange* otherTextRange, enum TextPatternRangeEndpoint otherEndpoint) {
	HRESULT res;
	int endComp = 0;
	res = thisTextRange->CompareEndpoints(thisEndpoint, otherTextRange, otherEndpoint, &endComp);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange::compareEndpoints");
	return endComp;
}

void UIATextRange_MoveEndpointByRange(IUIAutomationTextRange* thisTextRange, enum TextPatternRangeEndpoint thisEndpoint, IUIAutomationTextRange* otherTextRange, enum TextPatternRangeEndpoint otherEndpoint) {
	HRESULT res;
	res = thisTextRange->MoveEndpointByRange(thisEndpoint, otherTextRange, otherEndpoint);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange::MoveEndpointByRange");
}

CComPtr<IUIAutomationElement> UIATextRange_GetEnclosingElement(IUIAutomationTextRange* pTextRange) {
	HRESULT res;
	CComPtr<IUIAutomationElement> enclosingElement;
	res = pTextRange->GetEnclosingElement(&enclosingElement);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange::GetEnclosingElement");
	return enclosingElement;
}

CComPtr<IUIAutomationElement> UIATextRange3_GetEnclosingElementBuildCache(IUIAutomationTextRange3* pTextRange, IUIAutomationCacheRequest* cacheRequest) {
	HRESULT res;
	CComPtr<IUIAutomationElement> enclosingElement;
	res = pTextRange->GetEnclosingElementBuildCache(cacheRequest, &enclosingElement);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange3::GetEnclosingElementBuildCache");
	return enclosingElement;
}

CComPtr<IUIAutomationElementArray> UIATextRange_GetChildren(IUIAutomationTextRange* pTextRange) {
	HRESULT res;
	CComPtr<IUIAutomationElementArray> children;
	res = pTextRange->GetChildren(&children);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange::GetChildren");
	return children;
}

CComPtr<IUIAutomationElementArray> UIATextRange3_GetChildrenBuildCache(IUIAutomationTextRange3* pTextRange, IUIAutomationCacheRequest* cacheRequest) {
	HRESULT res;
	CComPtr<IUIAutomationElementArray> children;
	res = pTextRange->GetChildrenBuildCache(cacheRequest, &children);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationTextRange::GetChildren");
	return children;
}

int UIAElementArray_GetLength(IUIAutomationElementArray* elementArray) {
	HRESULT res;
	int length = 0;
	res = elementArray->get_Length(&length);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationElementArray::length");
	return length;
}

CComPtr<IUIAutomationElement> UIAElementArray_GetElement(IUIAutomationElementArray* elementArray, int index) {
	HRESULT res;
	CComPtr<IUIAutomationElement> child;
	res = elementArray->GetElement(index, &child);
	if (FAILED(res)) VERIFY_FAIL(L"IUIAutomationElementArray::GetElement");
	return child;
}
