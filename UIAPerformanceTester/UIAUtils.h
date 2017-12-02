/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#pragma once

#include "stdafx.h"

// Fetches the enclosing element of a UIA text range, caching requested properties. 
// Uses IUIAutomationTextRange3::GetEnclosingElementBuildCache if available,
// Or falls back to IUIAutomationTextRange::GetEnclosingElement, and calls BuildUpdatedCache on the resulting element. 
CComPtr<IUIAutomationElement> UIATextRangeX_GetEnclosingElementBuildCache(IUIAutomationTextRange* textRange, IUIAutomationCacheRequest* cacheRequest);

// Fetches the children of a UIA text range, caching requested properties.
// Uses IUIAutomationTextRange3::GetChildrenBuildCache if available,
// Or falls back to IUIAutomationTextRange::GetChildren, and calls BuildUpdatedcache on each child. 
std::vector<CComPtr<IUIAutomationElement>> UIATextRangeX_GetChildrenBuildCache(IUIAutomationTextRange* textRange, IUIAutomationCacheRequest* cacheRequest);

// Fetches 1 or more attributes from a UIA text range.
// Uses IUIAutomationTextRange3::GetAttributeValues if avilable,
// Or falls back to calling IUIAutomationTextRange::GetAttributeValue for each attribute.
std::vector<CComVariant> UIATextRangeX_GetAttributeValues(IUIAutomationTextRange* textRange, const std::vector<int>& attribs);

// Splits a UIA text range into smaller units.
std::vector<CComPtr<IUIAutomationTextRange>> splitUIATextRangeByUnit(IUIAutomationTextRange* pTextRange, TextUnit unit);
