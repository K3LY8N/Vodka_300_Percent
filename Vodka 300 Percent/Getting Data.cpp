#define _WIN32_DCOM
#include <iostream>
using namespace std;
#include <comdef.h>
#include <Wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")

int main(int argc, char** argv)
{
    setlocale(LC_ALL, "RUS");
    HRESULT hres;

    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres))
    {
        cout << "Failed to initialize COM library. Error code = 0x"
            << hex << hres << endl;
        return 1;
    }

    hres = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE,
        NULL
    );

    if (FAILED(hres))
    {
        cout << "Failed to initialize security. Error code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1;
    }

    IWbemLocator* pLoc = NULL;

    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);

    if (FAILED(hres))
    {
        cout << "Failed to create IWbemLocator object."
            << " Err code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1;
    }

    IWbemServices* pSvc = NULL;

    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"),
        NULL,
        NULL,
        0,
        NULL,
        0,
        0,
        &pSvc
    );

    if (FAILED(hres))
    {
        cout << "Could not connect. Error code = 0x"
            << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

    hres = CoSetProxyBlanket(
        pSvc,
        RPC_C_AUTHN_WINNT,
        RPC_C_AUTHZ_NONE,
        NULL,
        RPC_C_AUTHN_LEVEL_CALL,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE
    );

    if (FAILED(hres))
    {
        cout << "Could not set proxy blanket. Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    /*Полученние данных о системе*/
    
    IEnumWbemClassObject* Getting_System_Data = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_OperatingSystem"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &Getting_System_Data);

    if (FAILED(hres))
    {
        cout << "ERROR" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }
    
    /*Вывод полученных данных*/
    
    IWbemClassObject* Operating_System_Object = NULL;
    ULONG Operating_System_Return = 0;

    while (Getting_System_Data)
    {
        HRESULT hr = Getting_System_Data->Next(WBEM_INFINITE, 1,
            &Operating_System_Object, &Operating_System_Return);

        if (0 == Operating_System_Return)
        {
            break;
        }

        VARIANT vtProp;

        VariantInit(&vtProp);
        hr = Operating_System_Object->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << "System Name: " << vtProp.bstrVal << endl;
        hr = Operating_System_Object->Get(L"Version", 0, &vtProp, 0, 0);
        wcout << "Version: " << vtProp.bstrVal << endl;
        VariantClear(&vtProp);

        Operating_System_Object->Release();
    }

    pSvc->Release();
    pLoc->Release();
    Getting_System_Data->Release();
    CoUninitialize();
    
    /*Полученние данных о памяти техники*/
    
    IEnumWbemClassObject* Getting_Memory_Data = NULL;
    hres = pSvc->ExecQuery( /*Здесь проблема*/
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_LogicalDisk"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &Getting_Memory_Data);

    if (FAILED(hres))
    {
        cout << "ERROR" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }
    
    /*Вывод полученных данных*/
    
    IWbemClassObject* Logical_Disk_Object = NULL;
    ULONG Logical_Disk_Return = 0;

    while (Getting_Memory_Data)
    {
        HRESULT hr = Getting_Memory_Data->Next(WBEM_INFINITE, 1,
            &Logical_Disk_Object, &Logical_Disk_Return);

        if (0 == Logical_Disk_Return)
        {
            break;
        }

        VARIANT vtProp;

        VariantInit(&vtProp);
        hr = Logical_Disk_Object->Get(L"Description", 0, &vtProp, 0, 0);
        wcout << "Disk Name: " << vtProp.bstrVal << endl;
        hr = Logical_Disk_Object->Get(L"Size", 0, &vtProp, 0, 0);
        wcout << "Size: " << vtProp.bstrVal << endl;
        hr = Logical_Disk_Object->Get(L"FreeSpace", 0, &vtProp, 0, 0);
        wcout << "FreeSpace: " << vtProp.bstrVal << endl;
        VariantClear(&vtProp);

        Logical_Disk_Object->Release();
    }

    pSvc->Release();
    pLoc->Release();
    Getting_Memory_Data->Release();
    CoUninitialize();
    
    /*Полученние данных о процессоре*/
    
    IEnumWbemClassObject* Getting_CPU_Data = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_Processor"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &Getting_CPU_Data);

    if (FAILED(hres))
    {
        cout << "ERROR" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }
    
    /*Вывод полученных данных*/
    
    IWbemClassObject* Processor_Object = NULL;
    ULONG Processor_Return = 0;

    while (Getting_CPU_Data)
    {
        HRESULT hr = Getting_CPU_Data->Next(WBEM_INFINITE, 1,
            &Processor_Object, &Processor_Return);

        if (0 == Processor_Return)
        {
            break;
        }

        VARIANT vtProp;

        VariantInit(&vtProp);
        hr = Processor_Object->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << "Processor Name: " << vtProp.bstrVal << endl;
        /*hr = Processor_Object->Get(L"NumberOfCores", 0, &vtProp, 0, 0);
        wcout << "Number Of Cores: " << vtProp.bstrVal << endl;*/ /*Программа выдаёт ошибку*/
        hr = Processor_Object->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        wcout << "Manufacturer: " << vtProp.bstrVal << endl;
        VariantClear(&vtProp);

        Processor_Object->Release();
    }

    pSvc->Release();
    pLoc->Release();
    Getting_CPU_Data->Release();
    CoUninitialize();
    

    /*Полученние данных о видеокарте*/
    
    IEnumWbemClassObject* Getting_VideoController_Data = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_VideoController"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &Getting_VideoController_Data);

    if (FAILED(hres))
    {
        cout << "ERROR" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }
    
    /*Вывод полученных данных*/
    
    IWbemClassObject* VideoController_Object = NULL;
    ULONG VideoController_Return = 0;

    while (Getting_VideoController_Data)
    {
        HRESULT hr = Getting_VideoController_Data->Next(WBEM_INFINITE, 1,
            &VideoController_Object, &VideoController_Return);

        if (0 == VideoController_Return)
        {
            break;
        }

        VARIANT vtProp;

        VariantInit(&vtProp);
        hr = VideoController_Object->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << "VideoController Name: " << vtProp.bstrVal << endl;
        hr = VideoController_Object->Get(L"DriverVersion", 0, &vtProp, 0, 0);
        wcout << "Driver Version: " << vtProp.bstrVal << endl;
        VariantClear(&vtProp);

        VideoController_Object->Release();
    }

    pSvc->Release();
    pLoc->Release();
    Getting_VideoController_Data->Release();
    CoUninitialize();

    /*Полученние данных о звуковых устройствах*/

    IEnumWbemClassObject* Getting_Sound_Device_Data = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_SoundDevice"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &Getting_Sound_Device_Data);

    if (FAILED(hres))
    {
        cout << "ERROR" << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }
    
    /*Вывод полученных данных*/
    
    IWbemClassObject* Sound_Device_Object = NULL;
    ULONG Sound_Device_Return = 0;

    while (Getting_Sound_Device_Data)
    {
        HRESULT hr = Getting_Sound_Device_Data->Next(WBEM_INFINITE, 1,
            &Sound_Device_Object, &Sound_Device_Return);

        if (0 == Sound_Device_Return)
        {
            break;
        }

        VARIANT vtProp;

        VariantInit(&vtProp);
        hr = Sound_Device_Object->Get(L"ProductName", 0, &vtProp, 0, 0);
        wcout << "Sound Device: " << vtProp.bstrVal << endl;
        VariantClear(&vtProp);

        Sound_Device_Object->Release();
    }

    pSvc->Release();
    pLoc->Release();
    Getting_Sound_Device_Data->Release();
    CoUninitialize();
    return 0;
}