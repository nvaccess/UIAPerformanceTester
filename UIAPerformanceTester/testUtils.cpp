/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#include "stdafx.h"
#include "testUtils.h"

using namespace WEX::Logging;
using namespace WEX::Common;

CComPtr<IUIAutomation> createUIAClient() {
	HRESULT res;
	CComPtr<IUIAutomation> client = nullptr;
	res=client.CoCreateInstance(CLSID_CUIAutomation);
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

void launchEdgeWithURL(const wchar_t* URL) {
	CComPtr<IApplicationActivationManager> actMgr;
	HRESULT res;
	res = actMgr.CoCreateInstance(CLSID_ApplicationActivationManager, NULL, CLSCTX_INPROC_SERVER);
	if(FAILED(res)) VERIFY_FAIL(L"CoCreateInstance of ApplicationActivationManager");
	DWORD processID = 0;
	res = actMgr->ActivateApplication(L"Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge", URL, AO_NONE, &processID);
	if (FAILED(res)) VERIFY_FAIL(L"IApplicationActivationManager::ActivateApplication");
	Sleep(3000);
}
