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

// launches a UWP app given an app ID and parameters, and closes it on distruction
class UWPApp {
	public:
	UWPApp(const wchar_t* appID, const wchar_t* params);
	~UWPApp();
	protected:
	HWND appWindow{nullptr};
	CHandle processHandle{0};
};

class UWPApp_Edge: public UWPApp {
	protected:
	static const wchar_t* appID;
	public:
	UWPApp_Edge(const wchar_t* URL);
	CComPtr<IUIAutomationElement> locateDocumentUIAElement(IUIAutomation* UIAClient, const wchar_t* documentName);
};
