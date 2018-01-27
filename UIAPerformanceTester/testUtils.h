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
void verifyTakesLessThan(std::wstring_view description, double expectedTime, std::function<void()> func);

// launches a UWP app given an app ID and parameters, and closes it on distruction
class UWPApp {
	public:
	UWPApp(std::wstring_view appID, std::wstring_view params, bool launchInWDAG);
	~UWPApp();
	protected:
	HWND appWindow{nullptr};
	CHandle processHandle{0};
	private:
	void initAsStandard(std::wstring_view appID, std::wstring_view params);
	void initInWDAG(std::wstring_view appID, std::wstring_view params);
};

class UWPApp_Edge: public UWPApp {
	protected:
	static const std::wstring appID;
	public:
	UWPApp_Edge(std::wstring_view URL, bool launchInWDAG=false);
	CComPtr<IUIAutomationElement> locateDocumentUIAElement(IUIAutomation* UIAClient, std::wstring_view documentName);
};

class UWPApp_Word {
protected:
	static const std::wstring appID;
public:
	UWPApp_Word(std::wstring_view URL);
	~UWPApp_Word();
	CComPtr<IUIAutomationElement> locateDocumentUIAElement(IUIAutomation* UIAClient);
protected:
	HWND appWindow{ nullptr };
	CHandle processHandle{ 0 };
};
