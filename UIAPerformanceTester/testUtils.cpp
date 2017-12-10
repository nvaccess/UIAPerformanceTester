/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#include "stdafx.h"
#include "UIAWrappers.h"
#include "testUtils.h"

using namespace WEX::Logging;
using namespace WEX::Common;

CComPtr<IUIAutomation> createUIAClient() {
	HRESULT res;
	CComPtr<IUIAutomation> client = nullptr;
	res=client.CoCreateInstance(CLSID_CUIAutomation8,nullptr, CLSCTX_INPROC_SERVER);
	if (FAILED(res)||!client) VERIFY_FAIL(L"CoCreateInstance of UIAutomation");
	return client;
}

void verifyTakesLessThan(const wchar_t* description, double expectedTime, std::function<void()> func) {
	LARGE_INTEGER frequency;
	QueryPerformanceFrequency(&frequency);
	LARGE_INTEGER counter;
	QueryPerformanceCounter(&counter);
	double startTime = (double)counter.QuadPart / frequency.QuadPart;
	func();
	QueryPerformanceCounter(&counter);
	double endTime = (double)counter.QuadPart / frequency.QuadPart;
	double realTime = endTime - startTime;
	std::wstringstream s;
	s << L"Timing "<<description << std::endl << L"Execution took " << realTime << L"seconds, expected less than " << expectedTime<<L" seconds.";
	Log::Comment(s.str().c_str());
	VERIFY_IS_LESS_THAN(realTime, expectedTime, description);
}

void closeForegroundApp() {
	HWND fg = GetForegroundWindow();
	SendMessage(fg, WM_CLOSE, 0, 0);
}

HWND launchEdgeWithURL(const wchar_t* URL) {
	CComPtr<IApplicationActivationManager> actMgr;
	HRESULT res;
	res = actMgr.CoCreateInstance(CLSID_ApplicationActivationManager, NULL, CLSCTX_INPROC_SERVER);
	if(FAILED(res)) VERIFY_FAIL(L"CoCreateInstance of ApplicationActivationManager");
	DWORD appProcessID = 0;
	Log::Comment(L"Launching Edge");
	res = actMgr->ActivateApplication(L"Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge", URL, AO_NONE, &appProcessID);
	if (FAILED(res)) VERIFY_FAIL(L"IApplicationActivationManager::ActivateApplication");
	HWND appWindow=nullptr;
	wchar_t* lastErrorMsg=L"Unknown";
	for(int tryCount=0;tryCount<50;++tryCount) {
		HWND tempWindow=GetForegroundWindow();
		wchar_t fgClassName[256]={0};
		GetClassName(tempWindow,fgClassName,ARRAYSIZE(fgClassName));
		if(wcscmp(fgClassName, L"ApplicationFrameWindow")==0) {
			tempWindow = FindWindowEx(tempWindow, nullptr, L"Windows.UI.Core.CoreWindow", nullptr);
			if(tempWindow) {
				DWORD tempProcessID=0;
				GetWindowThreadProcessId(tempWindow,&tempProcessID);
				if(tempProcessID==appProcessID) {
					appWindow=tempWindow;
					break;
				} else lastErrorMsg=L"Active ApplicationFrameHost is not hosting correct process";
			} else lastErrorMsg=L"ApplicationFrameHost does not yet have a core window";
		} else lastErrorMsg=L"No active ApplicationFrameHost window";
		Beep(3000, 30);
		Sleep(200);
	}
	if(!appWindow) VERIFY_FAIL(lastErrorMsg);
	return appWindow;
}

CComPtr<IUIAutomationElement> locateEdgeDocumentUIAElement(IUIAutomation* UIAClient, HWND edgeWindow, const wchar_t* documentName) { 
	CComPtr<IUIAutomationElement> edgeRoot = UIA_ElementFromHandle(UIAClient, edgeWindow);
	if(!edgeRoot) VERIFY_FAIL(L"Can't get Edge's root UIA element");
	CComPtr<IUIAutomationCondition> nameCondition = UIA_CreatePropertyCondition(UIAClient, UIA_NamePropertyId, documentName);
	CComPtr<IUIAutomationCondition> controlTypeCondition = UIA_CreatePropertyCondition(UIAClient, UIA_ControlTypePropertyId, UIA_PaneControlTypeId);
	CComPtr<IUIAutomationCondition> textPatternCondition = UIA_CreatePropertyCondition(UIAClient, UIA_IsTextPatternAvailablePropertyId, true);
	std::vector<IUIAutomationCondition*> propertyConditions = { nameCondition,controlTypeCondition,textPatternCondition };
	CComPtr<IUIAutomationCondition> andCondition = UIA_CreateAndConditionFromArray(UIAClient, propertyConditions);
	CComPtr<IUIAutomationElement> document;
	for(int tryCount=0;tryCount<=50;++tryCount) {
		document=UIAElement_FindFirst(edgeRoot, TreeScope_Subtree, andCondition);
		if(document) break;
		Beep(3000,30);
		Sleep(200);
	}
	if(!document) VERIFY_FAIL(L"Can't locate Edge document");
	return document;
}
