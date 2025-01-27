#include <iostream>

#include <winsdkver.h>
#include <windows.h>
#include <stdio.h>
#include <string>

#pragma comment(lib, "advapi32.lib")

void show_usage(const wchar_t* cmd)
{
    wprintf(L"\nInstall Tools 1.0 - Creates edge conditions for installer testing\n");
    wprintf(L"Copyright (C) 2025 Alnoor Allidina\n");
    wprintf(L"https://needleinathreadstack.wordpress.com/\n");
    wprintf(L"\n");
    wprintf(L"Usage: %ls [-s][-f][-w]\n", cmd);
    wprintf(L"  -s <servicename> - Open service for read and wait to close the handle\n");
    wprintf(L"  -f <filename> [sharing flags (1=read, 2=write, 4=delete)]- Open file for read and wait to close the handle\n");
    wprintf(L"  -e <exe> <filepath> Copy exe to filename and execute it. Wait to close the handle\n");
    wprintf(L"\n");
}

enum class operation
{
    INVALID,
    SERVICE_OPEN_WAIT,
    FILE_OPEN_WAIT,
    EXE_COPY_EXECUTE_WAIT,
};

std::wstring get_error_message(DWORD errorCode) {
    LPWSTR error_buffer = nullptr;

    // Format the error message
    DWORD size = FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        nullptr,
        errorCode,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        (LPWSTR)&error_buffer,
        0,
        nullptr);

    if (size == 0) {
        // If FormatMessage fails, return a generic error message
        return L"Unknown error.";
    }

    std::wstring error_message(error_buffer);
    LocalFree(error_buffer);

    // Trim trailing newline characters if any
    while (!error_message.empty() && (error_message.back() == L'\n' || error_message.back() == L'\r')) {
        error_message.pop_back();
    }

    return error_message;
}

void print_error(const wchar_t* msg, DWORD error)
{
  wprintf(L"%ls. Error: %u (%ls)", msg, error, get_error_message(error).c_str());
}

void wait_for_enter()
{
  int c;
  // Eat up other chars
  while ((c = getchar()) != '\n' && c != EOF) 
  {
  };
}

bool service_open_wait(const wchar_t* service_name) {
  SC_HANDLE hSCManager = NULL;
  SC_HANDLE hService = NULL;
  bool result = false;

  // Open a handle to the Service Control Manager
  hSCManager = OpenSCManager(NULL, NULL, SC_MANAGER_CONNECT);
  if (hSCManager == NULL) {
      print_error(L"Failed to open Service Control Manager", GetLastError());
      goto cleanup;
  }

  wprintf(L"Attempting to open service: %ls\n", service_name);
  hService = OpenServiceW(hSCManager, service_name, SERVICE_QUERY_STATUS);
  if (hService == NULL) {
      print_error(L"Failed to open service", GetLastError());
      CloseServiceHandle(hSCManager);
      goto cleanup;
  }

  wprintf(L"Successfully opened service: %ls\n", service_name);
  wprintf(L"Press Enter to close the service handle and exit.\n");
  wait_for_enter();

  result = true;

cleanup:

  if (hService)
  {
    wprintf(L"Closing service handle.\n");
    CloseServiceHandle(hService);
    hService = nullptr;
  }

  if (hSCManager)
  {
    CloseServiceHandle(hSCManager);
    hSCManager = nullptr;
  }

  return result;
}

bool file_open_wait(const wchar_t* file_name, int sharing_flags) 
{
  HANDLE hFile = INVALID_HANDLE_VALUE;
  bool result = false;

  wprintf(L"Opening file: %ls with sharing %u\n", file_name, sharing_flags);
  hFile = CreateFileW(file_name, GENERIC_READ, sharing_flags, nullptr, OPEN_EXISTING, 0, nullptr);
  if (INVALID_HANDLE_VALUE == hFile)
  {
    print_error(L"Failed to open file", GetLastError());
    goto cleanup;
  }

  wprintf(L"Successfully opened file: %ls\n", file_name);
  wprintf(L"Press Enter to close the service handle and exit.\n");
  wait_for_enter();

  result = true;

cleanup:

  if (INVALID_HANDLE_VALUE != hFile)
  {
    wprintf(L"Closing file handle.\n");
    CloseHandle(hFile);
    hFile = nullptr;
  }

  return result;

}

bool exe_copy_execute_wait(const wchar_t* exe_name, const wchar_t* file_name)
{
  wprintf(L"Copying %ls to %ls\n", exe_name, file_name);
  wprintf(L"This will overwrite %ls if it exists - Please confirm by pressing Enter\n", file_name);
  wait_for_enter();
  if (!CopyFileW(exe_name, file_name, false))
  {
    print_error(L"Copy failed", GetLastError());
    return false;
  }

  STARTUPINFOW si = {0};
  PROCESS_INFORMATION pi = {0};
  wprintf(L"Executing %ls\n", file_name);
  if (!CreateProcessW(nullptr, const_cast<wchar_t*>(file_name), nullptr, nullptr, false, CREATE_NO_WINDOW, nullptr, nullptr, &si, &pi))
  {
    print_error(L"Execute failed", GetLastError());
    return false;
  }
  CloseHandle(pi.hThread);
  wprintf(L"Successfully exected %ls. PID is %u\n", file_name, pi.dwProcessId);
  wprintf(L"Presumably it's a long-running process. Press Enter to kill it.\n");
  wait_for_enter();

  if (!TerminateProcess(pi.hProcess, 0))
  {
    print_error(L"Terminate failed", GetLastError());
  }
  CloseHandle(pi.hProcess);

  return true;
}

std::wstring normalize_name(const wchar_t* path)
{
  constexpr DWORD MAX_PATH_BUFFER = 32768; 
  std::wstring normalized_path(MAX_PATH_BUFFER, L'\0');

  HANDLE hFile = CreateFileW(path, 0, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, 0, nullptr);
  if (INVALID_HANDLE_VALUE != hFile)
  {
    DWORD len = GetFinalPathNameByHandleW(hFile, &normalized_path[0], MAX_PATH_BUFFER, FILE_NAME_NORMALIZED);
    if (len)
    {
      // Success
      normalized_path.resize(len);
    }
    else
    {
      normalized_path = path;  
    }
    CloseHandle(hFile);
    hFile = INVALID_HANDLE_VALUE;
  }
  else {
    normalized_path = path;
  }

  return normalized_path;
}

int wmain(int argc, const wchar_t* argv[]) 
{
    operation op = operation::INVALID;
    int result = EXIT_FAILURE;

    const wchar_t* service_name;

    std::wstring file_name;
    std::wstring exe_name;
    int sharing_flags = 0;

#ifdef _M_IX86
    PVOID old_value = nullptr;  
    Wow64DisableWow64FsRedirection(&old_value);
#endif    
    
    for (int i = 1; i < argc; i++)
    {
        std::wstring arg = argv[i];
        if (arg == L"-s" || arg == L"-S")
        {
            if (i + 1 < argc)
            {
                i++;
                service_name = argv[i];
                op = operation::SERVICE_OPEN_WAIT;
            }
        }

        if (arg == L"-f" || arg == L"-F")
        {
            if (i + 1 < argc)
            {
                i++;
                file_name = argv[i];
                op = operation::FILE_OPEN_WAIT;

                // optional flags
                if (i + 1 < argc) {
                  i++;
                  sharing_flags = _wtoi(argv[i]);
                }
            }
        }

        if (arg == L"-e" || arg == L"-E")
        {
          if (i + 2 < argc)
          {
            i++;
            exe_name = argv[i];
            i++;
            file_name = argv[i];
            op = operation::EXE_COPY_EXECUTE_WAIT;
          }
        }
    }

    if (file_name.length())
    {
      file_name = normalize_name(file_name.c_str());
    }
    if (exe_name.length())
    {
      exe_name = normalize_name(exe_name.c_str());
    }

    switch (op) {
      case operation::SERVICE_OPEN_WAIT:
        service_open_wait(service_name);
        break;
      case operation::FILE_OPEN_WAIT:
        file_open_wait(file_name.c_str(), sharing_flags);
        break;
      case operation::EXE_COPY_EXECUTE_WAIT:
        exe_copy_execute_wait(exe_name.c_str(), file_name.c_str());
        break;
      case operation::INVALID:
      default:
        show_usage(argv[0]);
        result = EXIT_FAILURE;
    }

    return result;
}

