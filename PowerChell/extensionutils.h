#pragma once
#include <vector>
#include <string>

// provided for communication with the implant
typedef int (*goCallback)(const char*, int);

// argument parsing
std::vector<std::wstring> convertArgsBufferToArgv(const char* argsBuffer);

// output for callback
class Logger {
private:
    std::string output;
public:
    void log(const char* format, ...);
    void wlog(const wchar_t* format, ...);
    void sendOutput(goCallback callback);
    void clear() { output.clear(); }
};

// stores our response to the implant
extern Logger logger;

// string conversion
std::string wstringToString(const std::wstring& wstr);
std::wstring ConvertToWideString(const std::string& narrowString);