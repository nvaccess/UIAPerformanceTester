/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#pragma once

#include "stdafx.h"

CComPtr<IUIAutomationElement> UIATextRangeX_GetEnclosingElementBuildCache(IUIAutomationTextRange* textRange, IUIAutomationCacheRequest* cacheRequest);
std::vector<CComPtr<IUIAutomationElement>> UIATextRangeX_GetChildrenBuildCache(IUIAutomationTextRange* textRange, IUIAutomationCacheRequest* cacheRequest);
std::vector<CComVariant> UIATextRangeX_GetAttributeValues(IUIAutomationTextRange* textRange, const std::vector<int>& attribs);
std::vector<CComPtr<IUIAutomationTextRange>> splitUIATextRangeByUnit(IUIAutomationTextRange* pTextRange, TextUnit unit);
