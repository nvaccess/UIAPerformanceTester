/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#pragma once

#include "stdafx.h"
#include "UIAWrappers.h"
#include "UIAUtils.h"

// Initializes a UIAutomation client instance, returning an IUIAutomation interface 
CComPtr<IUIAutomation> createUIAClient();
// Times a function, logging its time and asserting it does not take longer than a maximum time
void verifyTakesLessThan(const wchar_t* description, double expectedTime, std::function<void()> func);
// Launches Edge and returns Edge's root window
HWND launchEdgeWithURL(const wchar_t* URL);
// Locates a document UIAElement with the given name, within the given Edge window.
// This function keeps searching over and over until it finds the document or  a timeout exceeds.
CComPtr<IUIAutomationElement> locateEdgeDocumentUIAElement(IUIAutomation* UIAClient, HWND edgeWindow, const wchar_t* documentName);
