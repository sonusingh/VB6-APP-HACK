/*

NOTE THIS MUST BE COMPILED AS WIN32 (NOT x64) for the dll to be injected
As the DLL is also win32 because VB APPS are win32

*/
#define _WIN32_WINNT 0x0501
#include <Windows.h>
#include <stdio.h>
#include <iostream>
#include <shlobj.h>
#include <tlhelp32.h>
#include <tchar.h>

// required for cout and endl
using namespace std;

/* GLOBAL DECLARATIONS */

//store process ID of running VB APP
unsigned long vbAppId = 0;
LPCSTR szDllName = "extractor.dll";

//enumerate desktop windows
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam)
{
	char class_name[256];
	GetClassName(hwnd, class_name, sizeof(class_name));

	if (strcmp(class_name, "ThunderRT6FormDC") == 0)
	{
		// Found a window with the specified class name
		DWORD process_id;
		vbAppId = GetWindowThreadProcessId(hwnd, &process_id);

		// You can perform actions on the window here
		// For example, bring it to the foreground
		SetForegroundWindow(hwnd);

		// To stop enumerating windows, return FALSE
		return FALSE;
	}

	// Continue enumerating windows
	return TRUE;
}

// main function
// we will check for startup commanline args passed
int main(int argc, char* argv[]) {
	//get a handle of the console window for changing output text color
	HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);

	// Enumerate all top-level windows
	if (EnumWindows(EnumWindowsProc, 0))
	{
		//set output color to green
		SetConsoleTextAttribute(hConsole, FOREGROUND_RED);
		std::cout << "No windows with class name 'ThunderRT6FormDC' were found, i.e. no VB apps running!\n\n" << std::endl;
		// Prompt the user to press Enter key to exit
		std::cout << "Press Enter key to exit...";
		std::cin.get(); // Wait for the user to press Enter

		return 0;
	}

	//set output color to green
	SetConsoleTextAttribute(hConsole, 10);
	//print KME/AIMS app id
	printf("VB6 App Process ID: %i\n", vbAppId);

	// get id of this app and pass it to the dll
	unsigned long currentAppID = GetCurrentProcessId();
	printf("This (injector) App's ID: %i\n", currentAppID);

	//set console output color back to white
	SetConsoleTextAttribute(hConsole, 15);

	HINSTANCE hinst = LoadLibrary(szDllName);

	if (hinst) {
		typedef void (*Install)(unsigned long, unsigned long);
		typedef void (*Uninstall)();

		Install install = (Install)GetProcAddress(hinst, "install");
		Uninstall uninstall = (Uninstall)GetProcAddress(hinst, "uninstall");

		install(vbAppId, currentAppID);

		MSG msg = {};

		while (GetMessage(&msg, NULL, 0, 0)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}

		uninstall();
	}
	else {
		//set console output color to red
		SetConsoleTextAttribute(hConsole, 12);
		cout << "Unable to load library (" << szDllName << ")\n" << endl;
		//set console output color back to white
		SetConsoleTextAttribute(hConsole, 15);
	}

	return 0;
}
