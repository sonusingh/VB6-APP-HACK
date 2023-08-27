/*

NOTE THIS MUST BE COMPILED AS WIN32 (NOT x64) for the dll to be injected
As the DLL is also win32 because AIMS is win32

for MingW32 use:

C:\Users\ssingh1\Downloads\MinGW\mingw32\bin\i686-w64-mingw32-g++ -m32 -o C:\Users\ssingh1\Downloads\MinGW\code\test.exe C:\Users\ssingh1\Downloads\MinGW\code\app.cpp -s -static-libgcc -static-libstdc++

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

//AIMS processID
unsigned long AIMSprocessID = 0;

// depending on arguments supplied
// we declare the vars and set the correct data in vars
// in the main()
LPCSTR szClass = "ThunderRT6FormDC"; //is the same for both KME and AIMS
LPCSTR szCaption = "Form1"; //this is what changes depending on user supplied args
LPCSTR szDllName = "extractor.dll"; //default to this

// get the id of process from window handle
unsigned long GetTargetThreadIdFromWindow(const char* className, const char* windowName)
{
	HWND targetWnd;
	unsigned long processID = 0;

	targetWnd = FindWindow(className, windowName);
	return GetWindowThreadProcessId(targetWnd, &processID);
}


// main function
// we will check for startup commanline args passed
int main(int argc, char* argv[]) {

	//get a handle of the console window for changing output text color
	HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);

	// injection of DLL starts here
	unsigned long threadID = GetTargetThreadIdFromWindow(szClass, szCaption); // supply the class and window title
	//unsigned long threadID = 0;

	//set output color to green
	SetConsoleTextAttribute(hConsole, 10);
	//print KME/AIMS app id
	printf("VB6 App Process ID: %i\n", threadID);

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

		install(threadID, currentAppID);

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
