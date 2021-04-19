// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
//#include <shlwapi.h>
#include <HtmlHelp.h>
#include <shellapi.h>
#include <iostream>

HMODULE hMod;
std::string PathAndName;
std::string OnlyPath;

/***********************************************************************
 *              create_file_OF
 *
 * Wrapper for CreateFile that takes OF_* mode flags.
 */
static HANDLE create_file_OF(LPCSTR path, INT mode)
{
    DWORD access, sharing, creation;

    if (mode & OF_CREATE)
    {
        creation = CREATE_ALWAYS;
        access = GENERIC_READ | GENERIC_WRITE;
    }
    else
    {
        creation = OPEN_EXISTING;
        switch (mode & 0x03)
        {
        case OF_READ:      access = GENERIC_READ; break;
        case OF_WRITE:     access = GENERIC_WRITE; break;
        case OF_READWRITE: access = GENERIC_READ | GENERIC_WRITE; break;
        default:           access = 0; break;
        }
    }

    switch (mode & 0x70)
    {
    case OF_SHARE_EXCLUSIVE:  sharing = 0; break;
    case OF_SHARE_DENY_WRITE: sharing = FILE_SHARE_READ; break;
    case OF_SHARE_DENY_READ:  sharing = FILE_SHARE_WRITE; break;
    case OF_SHARE_DENY_NONE:
    case OF_SHARE_COMPAT:
    default:                  sharing = FILE_SHARE_READ | FILE_SHARE_WRITE; break;
    }
    return CreateFileA(path, access, sharing, NULL, creation, FILE_ATTRIBUTE_NORMAL, 0);
}

std::string redirect_path(LPCSTR path)
{
    /*std::string tmpPath = path;
    char newPath[MAX_PATH];
    if (tmpPath.rfind("{CD}", 0) == 0) {
        sprintf_s(newPath, "%s\\%s\\%s", OnlyPath.c_str(), "CD", tmpPath.substr(4).c_str());
    }
    else {
        sprintf_s(newPath, "%s", tmpPath.c_str());
    }
    */
    std::string tmpPath = path;
    if (tmpPath.rfind("{CD}", 0) == 0) {
        return OnlyPath + "\\CD\\" + tmpPath.substr(4);
    }
    else {
        return tmpPath;
    }
}

/***********************************************************************
 *           _hwrite   (KERNEL32.@)
 *
 *	experimentation yields that _lwrite:
 *		o truncates the file at the current position with
 *		  a 0 len write
 *		o returns 0 on a 0 length write
 *		o works with console handles
 *
 */
LONG WINAPI _hwrite(HFILE handle, LPCSTR buffer, LONG count)
{
    DWORD result;

    if (!count)
    {
        /* Expand or truncate at current position */
        if (!SetEndOfFile(LongToHandle(handle))) return HFILE_ERROR;
        return 0;
    }
    if (!WriteFile(LongToHandle(handle), buffer, count, &result, NULL))
        return HFILE_ERROR;
    return result;
}

/***********************************************************************
 *           _lclose   (KERNEL32.@)
 */
static HFILE WINAPI KERNEL32__lclose(HFILE hFile)
{
    return CloseHandle(LongToHandle(hFile)) ? 0 : HFILE_ERROR;
}

/***********************************************************************
 *           _lcreat   (KERNEL32.@)
 */
static HFILE WINAPI KERNEL32__lcreat(LPCSTR path, INT attr)
{
    HANDLE hfile;

    /* Mask off all flags not explicitly allowed by the doc */
    attr &= FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM;
    hfile = CreateFileA(redirect_path(path).c_str(), GENERIC_READ | GENERIC_WRITE,
        FILE_SHARE_READ | FILE_SHARE_WRITE, NULL,
        CREATE_ALWAYS, attr, 0);
    return HandleToLong(hfile);
}

/***********************************************************************
 *           _lopen   (KERNEL32.@)
 */
static HFILE WINAPI KERNEL32__lopen(LPCSTR path, INT mode)
{
    HANDLE hfile;

    //MessageBoxA(NULL, redirect_path(path).c_str(), NULL, MB_ICONINFORMATION | MB_OK);
    hfile = create_file_OF(redirect_path(path).c_str(), mode & ~OF_CREATE);
    return HandleToLong(hfile);
}

/***********************************************************************
 *           _lread   (KERNEL32.@)
 */
static UINT WINAPI KERNEL32__lread(HFILE handle, LPVOID buffer, UINT count)
{
    DWORD result;
    if (!ReadFile(LongToHandle(handle), buffer, count, &result, NULL))
        return HFILE_ERROR;
    return result;
}

/***********************************************************************
 *           _llseek   (KERNEL32.@)
 */
static LONG WINAPI KERNEL32__llseek(HFILE hFile, LONG lOffset, INT nOrigin)
{
    return SetFilePointer(LongToHandle(hFile), lOffset, NULL, nOrigin);
}

/***********************************************************************
 *           _lwrite   (KERNEL32.@)
 */
static UINT WINAPI KERNEL32__lwrite(HFILE hFile, LPCSTR buffer, UINT count)
{
    return (UINT)_hwrite(hFile, buffer, (LONG)count);
}

static bool WINAPI KERNEL32_SetCurrentDirectoryA(LPCSTR path)
{
    return SetCurrentDirectoryA(redirect_path(path).c_str());
}

static bool WINAPI USER32_WinHelpA(HWND hWnd, LPCSTR lpHelpFile, UINT wCommand, ULONG_PTR dwData)
{
    //MessageBoxA(NULL, "WinHelpA was hooked!", NULL, MB_ICONINFORMATION | MB_OK);
    std::string tmpHelpFile = lpHelpFile;
    char chmHelpFile[MAX_PATH];
    sprintf_s(chmHelpFile, "%s\\%s%s", OnlyPath.c_str(), tmpHelpFile.substr(0, tmpHelpFile.find_last_of('.')).c_str(), ".CHM");
    switch (wCommand)
    {
        case HELP_CONTEXT:
            return HtmlHelpA(hWnd, chmHelpFile, HH_HELP_CONTEXT, dwData);
        case HELP_QUIT:
            return HtmlHelpA(NULL, chmHelpFile, HH_CLOSE_ALL, NULL);
        case HELP_CONTENTS:
            return HtmlHelpA(hWnd, chmHelpFile, HH_DISPLAY_TOPIC, NULL);
        case HELP_HELPONHELP:
            //MessageBoxA(NULL, "Not implemented!", NULL, MB_ICONINFORMATION | MB_OK);
            return ShellExecuteA(0, 0, "https://docs.microsoft.com/en-us/previous-versions/windows/desktop/htmlhelp/about-the-html-help-viewer", 0, 0, SW_SHOW);
        default:
            return TRUE;
    }
}

static HMODULE app_module;
static PIMAGE_IMPORT_DESCRIPTOR app_imports;

static BOOL patch_app_import(ULONG_PTR API, ULONG_PTR Hook)
{
    BOOL result = FALSE;
    PIMAGE_IMPORT_DESCRIPTOR import = app_imports;
    for (; import->Characteristics; ++import) {
        PIMAGE_THUNK_DATA thunk = (PIMAGE_THUNK_DATA)((ULONG_PTR)app_module + import->OriginalFirstThunk);
        PIMAGE_THUNK_DATA table = (PIMAGE_THUNK_DATA)((ULONG_PTR)app_module + import->FirstThunk);
        for (; thunk->u1.Ordinal; ++thunk, ++table) {
            if (table->u1.Function == API) {
                DWORD Protect;
                if (VirtualProtect(&table->u1.Function, sizeof(IMAGE_THUNK_DATA), PAGE_EXECUTE_READWRITE, &Protect)) {
                    table->u1.Function = Hook;
                    VirtualProtect(&table->u1.Function, sizeof(IMAGE_THUNK_DATA), Protect, &Protect);
                    FlushInstructionCache(GetCurrentProcess(), &table->u1.Function, sizeof(IMAGE_THUNK_DATA));
                    result = TRUE;
                }
            }
        }
    }
    return result;
}

#define HOOK(m, p) patch_app_import((ULONG_PTR)p, (ULONG_PTR)m##_##p);
#define UNHOOK(m, p) patch_app_import((ULONG_PTR)m##_##p, (ULONG_PTR)p);

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        DisableThreadLibraryCalls(hModule);
        app_module = GetModuleHandleA(NULL);
        if (app_module) {
            PIMAGE_DOS_HEADER DosHeader = (PIMAGE_DOS_HEADER)app_module;
            if (DosHeader && (IMAGE_DOS_SIGNATURE == DosHeader->e_magic) && (DosHeader->e_lfanew > 0)) {
                PIMAGE_NT_HEADERS32 NtHeaders = (PIMAGE_NT_HEADERS32)((ULONG_PTR)app_module + (ULONG)(DosHeader->e_lfanew));
                if ((IMAGE_NT_SIGNATURE == NtHeaders->Signature) && (IMAGE_FILE_MACHINE_I386 == NtHeaders->FileHeader.Machine) &&
                    (NtHeaders->FileHeader.SizeOfOptionalHeader >= FIELD_OFFSET(IMAGE_OPTIONAL_HEADER32, DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT + 1])) &&
                    (NtHeaders->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT].VirtualAddress > 0) &&
                    (NtHeaders->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT].Size >= sizeof(IMAGE_IMPORT_DESCRIPTOR)))
                    app_imports = (PIMAGE_IMPORT_DESCRIPTOR)((ULONG_PTR)app_module + NtHeaders->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT].VirtualAddress);
            }
        }
        if (app_imports) {
            HOOK(KERNEL32, _lclose);
            HOOK(KERNEL32, _lcreat);
            HOOK(KERNEL32, _lopen);
            HOOK(KERNEL32, _lread);
            HOOK(KERNEL32, _llseek);
            HOOK(KERNEL32, _lwrite);
            HOOK(KERNEL32, SetCurrentDirectoryA);
            HOOK(USER32, WinHelpA);
        }
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        //if (app_imports) {
            //UNHOOK(USER32, WinHelpA);
            //MessageBoxA(NULL, "WinHelpA quit?", NULL, MB_ICONINFORMATION | MB_OK);
        //}
        break;
    }
    hMod = hModule;
    char buffer[MAX_PATH];
    if (::GetModuleFileNameA(hMod, buffer, MAX_PATH - 1) <= 0)
    {
        return TRUE;
    }

    PathAndName = buffer;

    size_t found = PathAndName.find_last_of("/\\");
    OnlyPath = PathAndName.substr(0, found);
    return TRUE;

}

