// ApplicationPathRegKeyQueryTest.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <iostream>
#include <Shobjidl.h>
#include <shlobj.h>
#include <KnownFolders.h>
#include <wrl/client.h>

int main()
{
    HKEY hkey;
    //if (ERROR_SUCCESS != RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\winword.exe", 0, KEY_QUERY_VALUE, &hkey))
    if (ERROR_SUCCESS != RegOpenKeyEx(HKEY_CLASSES_ROOT, L"Word.Application\\CurVer", 0, KEY_QUERY_VALUE, &hkey))
    {
        std::cout << "Word Is Absent" << std::endl;
    }
    else
    {
        std::cout << "Word Is Present" << hkey << std::endl;
    }

    Microsoft::WRL::ComPtr<IShellItem> shellItem;

    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    HRESULT hr = SHCreateItemInKnownFolder(
        FOLDERID_AppsFolder,
        KF_FLAG_DONT_VERIFY,
        L"Microsoft.Office.WINWORD.EXE.15",
        IID_PPV_ARGS(shellItem.GetAddressOf()));

    if (hr == S_OK)
    {
        std::cout << "Found word using Shell Item win32 aumid" << std::endl;
    }

    system("pause");
    return 0;
}

