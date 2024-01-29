// 
// Copyright (c) 2023 - 2025 by QUALCOMM Technologies, Incorporated.  All Rights Reserved.
//
// 這個程式碼使用WMI API來列出所有在裝置管理員中有黃色驚嘆號的裝置名稱。
// This code uses the WMI API to list all device names that have a yellow exclamation mark in Device Manager.
// 
#include <iostream>
#include <wbemidl.h>
#pragma comment(lib, "wbemuuid.lib")
#include <comdef.h>

using namespace std;

int main(int argc, char** argv) {

    printf("Qualcomm device lost list tool is powered by Taiwan CE team.\nCopyright (c) %s %s by QUALCOMM Technologies, Incorporated. All Rights Reserved.\n\n", __DATE__, __TIME__);

    HRESULT hres;

    // Initialize COM.
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        cout << "Failed to initialize COM library. Error code = 0x"
            << hex << hres << endl;
        return 1;
    }

    // Set general COM security levels.
    hres = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE,
        NULL);
    if (FAILED(hres)) {
        cout << "Failed to initialize security. Error code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1;
    }

    // Obtain the initial locator to WMI.
    IWbemLocator* pLoc = NULL;
    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator,
        (LPVOID*)&pLoc);
    if (FAILED(hres)) {
        cout << "Failed to create IWbemLocator object. "
            << "Error code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1;
    }

    // Connect to WMI through the IWbemLocator::ConnectServer method.
    IWbemServices* pSvc = NULL;
    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"),
        NULL,
        NULL,
        0,
        NULL,
        0,
        0,
        &pSvc);
    if (FAILED(hres)) {
        cout << "Could not connect. Error code = 0x"
            << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    // Set the IWbemServices proxy so that impersonation
    // of the user (client) occurs.
    hres = CoSetProxyBlanket(
        pSvc,
        RPC_C_AUTHN_WINNT,
        RPC_C_AUTHZ_NONE,
        NULL,
        RPC_C_AUTHN_LEVEL_CALL,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE);
    if (FAILED(hres)) {
        cout << "Could not set proxy blanket. Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    // Use the IWbemServices pointer to make requests of WMI.
    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode <> 0"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);
    if (FAILED(hres)) {
        cout << "Query for operating system name failed."
            << " Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    // Get the data from the query.
    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;
    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) {
            break;
        }

        VARIANT vtProp;
        //hr = pclsObj->Get(L"PNPDeviceID", 0, &vtProp, 0, 0);
        //wcout << "PNP Device ID : " << vtProp.bstrVal << endl;
        //VariantClear(&vtProp);

        hr = pclsObj->Get(L"DeviceID", 0, &vtProp, 0, 0);
        wcout << "Device ID : " << vtProp.bstrVal << endl;
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << "Device Name : " << vtProp.bstrVal << endl;
        VariantClear(&vtProp);

        pclsObj->Release();
    }

    // Cleanup.
    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();
    CoUninitialize();

    return 0;
}
