//#include "pch.h"
#include "extensionutils.h"
#include <windows.h>
#include <stdio.h>
#include <shellapi.h>

// stores our response to the implant
Logger logger;

// argument parsing helper
std::vector<std::wstring> convertArgsBufferToArgv(const char* argsBuffer) {
    int bufferSize = MultiByteToWideChar(CP_UTF8, 0, argsBuffer, -1, nullptr, 0);
    WCHAR* wideBuffer = new WCHAR[bufferSize];
    MultiByteToWideChar(CP_UTF8, 0, argsBuffer, -1, wideBuffer, bufferSize);

    int argc;
    LPWSTR* argvW = CommandLineToArgvW(wideBuffer, &argc);

    std::vector<std::wstring> args(argvW, argvW + argc);

    // Cleanup
    LocalFree(argvW);
    delete[] wideBuffer;

    return args;
}

// to help return output to the implant
void Logger::log(const char* format, ...) {
    char buffer[512];
    va_list args;
    va_start(args, format);
    vsnprintf(buffer, sizeof(buffer), format, args);
    va_end(args);

    output += buffer;
}

void Logger::wlog(const wchar_t* format, ...) {
    wchar_t wbuffer[512];
    va_list args;
    va_start(args, format);
    _vsnwprintf_s(wbuffer, sizeof(wbuffer) / sizeof(wchar_t), format, args);
    va_end(args);

    // Convert to narrow for storage
    std::string buffer = wstringToString(wbuffer);
    output += buffer;
}

void Logger::sendOutput(goCallback callback) {
    callback(output.c_str(), output.length());
}

// string conversion
std::string wstringToString(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();

    // Get the required buffer size (-1 is not used to avoid issues with embedded nulls)
    int size = WideCharToMultiByte(CP_UTF8, 0, wstr.data(), wstr.size(), NULL, 0, NULL, NULL);

    // Allocate the buffer
    std::string result(size, 0);

    // Do the actual conversion
    WideCharToMultiByte(CP_UTF8, 0, wstr.data(), wstr.size(), &result[0], size, NULL, NULL);

    return result;
}

std::wstring ConvertToWideString(const std::string& narrowString) {
    // Determine the required size for the wide string
    size_t wideStringSize = 0;
    mbstowcs_s(&wideStringSize, nullptr, 0, narrowString.c_str(), 0);

    // Convert narrow string to wide string
    std::wstring wideString;
    wideString.resize(wideStringSize - 1); // excluding null terminator
    mbstowcs_s(nullptr, &wideString[0], wideStringSize, narrowString.c_str(), wideStringSize - 1);

    return wideString;
}