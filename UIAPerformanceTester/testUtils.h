/*
Copyright (C) 2017 NV Access Limited.
This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

#pragma once

#include "stdafx.h"

CComPtr<IUIAutomation> createUIAClient();
void verifyTakesLessThan(const wchar_t* description, double expectedTime, std::function<void()> func);
void closeForegroundApp();
void launchEdgeWithURL(const wchar_t* URL);
