#pragma once
#include "clr.h"

void StartPowerShell(LPCWSTR pwszCommand);
BOOL DisablePowerShellEtwProvider(mscorlib::_AppDomain* pAppDomain);
BOOL CreateInitialRunspaceConfiguration(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtRunspaceConfiguration);
BOOL StartConsoleShell(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtRunspaceConfiguration, LPCWSTR pwszBanner, LPCWSTR pwszHelp, LPCWSTR* ppwszArguments, DWORD dwArgumentCount);
BOOL ExecutePowerShellCommand(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtRunspaceConfiguration, LPCWSTR pwszCommand, BSTR* ppwszOutput);
BOOL ProcessPowerShellOutput(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtInvokeResult, BSTR* ppwszOutput);