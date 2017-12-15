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

void verifyTakesLessThan(std::wstring_view description, double expectedTime, std::function<void()> func) {
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
	VERIFY_IS_LESS_THAN(realTime, expectedTime, description.data());
}

void UWPApp::initAsStandard(std::wstring_view appID, std::wstring_view params) {
	CComPtr<IApplicationActivationManager> actMgr;
	VERIFY_SUCCEEDED(actMgr.CoCreateInstance(CLSID_ApplicationActivationManager, NULL, CLSCTX_INPROC_SERVER));
	DWORD appProcessID = 0;
	VERIFY_SUCCEEDED(actMgr->ActivateApplication(appID.data(), params.data(), AO_NONE, &appProcessID));
	processHandle.Attach(OpenProcess(SYNCHRONIZE | PROCESS_TERMINATE, false, appProcessID));
	if(!processHandle) VERIFY_FAIL(L"Cannot open process for synchronization and termination");
	wchar_t* lastErrorMsg=L"Unknown";
	for(int tryCount=0;tryCount<20;++tryCount) {
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
		Sleep(1000);
	}
	if(!appWindow) VERIFY_FAIL(lastErrorMsg);
}

void UWPApp::initInWDAG(std::wstring_view appID, std::wstring_view params) {
	CComPtr<IIsolatedAppLauncher> appLauncher;
	VERIFY_SUCCEEDED(appLauncher.CoCreateInstance(CLSID_IsolatedAppLauncher, NULL, CLSCTX_LOCAL_SERVER));
	VERIFY_SUCCEEDED(CoAllowSetForegroundWindow(appLauncher, nullptr));
	IsolatedAppLauncherTelemetryParameters telemetryParameters={TRUE,{0xff}};
	VERIFY_SUCCEEDED(appLauncher->Launch(appID.data(), params.data(), &telemetryParameters));
	wchar_t* lastErrorMsg = L"Unknown";
	for (int tryCount = 0; tryCount<20; ++tryCount) {
		EnumWindows([](HWND topWindow, LPARAM lParam) -> BOOL {
			wchar_t tempBuf[256] = { 0 };
			if(GetClassName(topWindow, tempBuf, ARRAYSIZE(tempBuf))>0&&wcscmp(tempBuf, L"RAIL_WINDOW") == 0) {
				if(GetWindowText(topWindow,tempBuf,ARRAYSIZE(tempBuf))>0&&wcsstr(tempBuf,L"UIAPerformanceTester:")) {
					*((HWND*)lParam)=topWindow;
					return false;
				}
			}
			return true;
		},(LPARAM)&appWindow);
		if(appWindow) break;
Beep(3000, 30);
		Sleep(1000);
	}
	if (!appWindow) VERIFY_FAIL(L"Could not locate RAIL_WINDOW");
	DWORD appProcessID = 0;
	GetWindowThreadProcessId(appWindow, &appProcessID);
	if(!appProcessID) VERIFY_FAIL(L"Could not get process ID of RAIL_WINDOW");
	processHandle.Attach(OpenProcess(SYNCHRONIZE | PROCESS_TERMINATE, false, appProcessID));
	if (!processHandle) VERIFY_FAIL(L"Cannot open process for synchronization and termination");
}

UWPApp::UWPApp(std::wstring_view appID, std::wstring_view params, bool launchInWDAG) {
	if (launchInWDAG) {
		initInWDAG(appID, params);
	}
	else {
		initAsStandard(appID, params);
	}
}

UWPApp::~UWPApp() {
	VERIFY_WIN32_SUCCEEDED(SendMessage(appWindow,WM_CLOSE,0,0));
	if(WaitForSingleObject(processHandle,30000)==WAIT_TIMEOUT) {
		Log::Warning(L"Timed out waiting for process to die after WM_CLOSE, forcing termination");
		TerminateProcess(processHandle,0);
	}
}

const std::wstring UWPApp_Edge::appID{L"Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge"};

UWPApp_Edge::UWPApp_Edge(std::wstring_view URL, bool launchInWDAG): UWPApp(UWPApp_Edge::appID,URL,launchInWDAG) {
}

CComPtr<IUIAutomationElement> UWPApp_Edge::locateDocumentUIAElement(IUIAutomation* UIAClient, std::wstring_view documentName) { 
	CComPtr<IUIAutomationElement> edgeRoot = UIA_ElementFromHandle(UIAClient, appWindow);
	if(!edgeRoot) VERIFY_FAIL(L"Can't get Edge's root UIA element");
	CComPtr<IUIAutomationCondition> nameCondition = UIA_CreatePropertyCondition(UIAClient, UIA_NamePropertyId, documentName.data());
	CComPtr<IUIAutomationCondition> controlTypeCondition = UIA_CreatePropertyCondition(UIAClient, UIA_ControlTypePropertyId, UIA_PaneControlTypeId);
	CComPtr<IUIAutomationCondition> textPatternCondition = UIA_CreatePropertyCondition(UIAClient, UIA_IsTextPatternAvailablePropertyId, true);
	std::vector<IUIAutomationCondition*> propertyConditions = { nameCondition,controlTypeCondition,textPatternCondition };
	CComPtr<IUIAutomationCondition> andCondition = UIA_CreateAndConditionFromArray(UIAClient, propertyConditions);
	CComPtr<IUIAutomationElement> document;
	for(int tryCount=0;tryCount<=20;++tryCount) {
		document=UIAElement_FindFirst(edgeRoot, TreeScope_Subtree, andCondition);
		if(document) break;
		Beep(3000,30);
		Sleep(1000);
	}
	if(!document) VERIFY_FAIL(L"Can't locate Edge document");
	return document;
}
