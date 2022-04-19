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
LPCSTR szClass 		= "ThunderRT6FormDC"; //is the same for both KME and AIMS
LPCSTR szCaption 	= ""; //this is what changes depending on user supplied args
LPCSTR szDllName 	= "extractor.dll"; //default to this

//set to false initially
bool processing = false;

/*
// get the main windows
HWND targetWnd = FindWindow("ThunderRT6FormDC", "Select Track Section");
HWND aimsReviewWindow = FindWindow("ThunderRT6FormDC","A.I.M.S  3.4.6 - Main");

// get controls on window
HWND cancelBtn = FindWindowExA(targetWnd, NULL, "ThunderRT6CommandButton", "&Cancel");
HWND backBtn = FindWindowExA(targetWnd, NULL, "ThunderRT6CommandButton", "<< &Back");
HWND browseBtn = FindWindowExA(targetWnd, NULL, "ThunderRT6CommandButton", "Browse &Archive");
HWND selectBtn = FindWindowExA(targetWnd, NULL, "ThunderRT6CommandButton", "&Select");
HWND importExceedences = FindWindowExA(targetWnd, NULL, "ThunderRT6CommandButton", "Import Exceedances");
HWND tabs = FindWindowExA(targetWnd, NULL, "SSTabCtlWndClass", NULL);
// note control is child of sub-control
HWND checkBox = FindWindowExA(tabs, NULL, "ThunderRT6CheckBox", NULL);
HWND flexGrid = FindWindowExA(tabs, NULL, "MSFlexGridWndClass", NULL);
// size for the tabbed area and msflexgrid
int width = 640;
int height = 1250;
*/

//check if a given process is running
bool IsProcessRunning(const TCHAR* const executableName) {
    PROCESSENTRY32 entry;
    entry.dwSize = sizeof(PROCESSENTRY32);

    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, NULL);

    if (!Process32First(snapshot, &entry)) {
        CloseHandle(snapshot);
        return false;
    }

    do {
        if (!_tcsicmp(entry.szExeFile, executableName)) {
            CloseHandle(snapshot);
            return true;
        }
    } while (Process32Next(snapshot, &entry));

    CloseHandle(snapshot);
    return false;
}

//start a process
bool startup(LPCTSTR lpApplicationName)
{
   // additional information
   STARTUPINFO si;     
   PROCESS_INFORMATION pi;

   // set the size of the structures
   ZeroMemory( &si, sizeof(si) );
   si.cb = sizeof(si);
   ZeroMemory( &pi, sizeof(pi) );

  // start the program up
  if (!CreateProcess( lpApplicationName,   // the path
    NULL,        	// Command line
    NULL,           // Process handle not inheritable
    NULL,           // Thread handle not inheritable
    FALSE,          // Set handle inheritance to FALSE
    0,              // No creation flags
    NULL,           // Use parent's environment block
    NULL,           // Use parent's starting directory 
    &si,            // Pointer to STARTUPINFO structure
    &pi             // Pointer to PROCESS_INFORMATION structure (removed extra parentheses)
    )){
		printf( "CreateProcess failed (%d).\n", GetLastError() );
        return false;
	}
	
	//this is the MAIN process ID for AIMS
	//it may help in closing the AIMS APP
	//AIMSprocessID = GetProcessId(pi.hProcess);
	//cout <<  AIMSprocessID << endl;
	
	return true;
}

// get the id of process from window handle
unsigned long GetTargetThreadIdFromWindow(const char* className, const char* windowName)
{
    HWND targetWnd;
    HANDLE hProcess;
    unsigned long processID = 0;

    targetWnd = FindWindow(className, windowName);
    return GetWindowThreadProcessId(targetWnd, &processID);
}

//get desktop directory for current user
string desktop_directory()
{
    static char path[MAX_PATH+1];
    if (SHGetSpecialFolderPathA(HWND_DESKTOP, path, CSIDL_DESKTOP, FALSE))
        return path;
    else
        return "ERROR";
}

// main function
// we will check for startup commanline args passed
int main(int argc, char* argv[]) {

	//get a handle of the console window for changing output text color
	HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);	

	//if we are injecting DLL into another app go straight to testing
	if (argc > 1 && argv[1] == std::string("processing")) {
		//set caption var
		szCaption = "Geomatic Technologies Kilometrage Engine";
		HWND targetWnd = FindWindow(szClass, szCaption);
		//if kme not running exit
		if (!targetWnd){
			//set console output color to red
			//see https://stackoverflow.com/a/4053879 for color codes
			SetConsoleTextAttribute(hConsole, 12);
			cout << "\n\nAIMS GT KME WINDOW NOT FOUND!\n\nQuitting!" << endl;
			SetConsoleTextAttribute(hConsole, 15);
			return 0;
		}
		
		//if KME window minimised exit
		if (IsIconic(targetWnd)){
			SetConsoleTextAttribute(hConsole, 12);
			cout << "\n\nKME window is not visible please show window on desktop and try again!\n\n" << endl;
			SetConsoleTextAttribute(hConsole, 15);
			return 0;
		}

		//set size of kme window
		SetWindowPos(targetWnd, 0, 0, 0, 2000, 1300, SWP_SHOWWINDOW);
		SetForegroundWindow(targetWnd);	
		
		//load processing dll
		szDllName = "processing_DLL.dll";
		
		//set to true
		processing = true;
		
		//ask user for run name
		int run = 0;
		cout << "\n\nPlease select the run number:\n" << endl;
		cout << "\t1 = North\n\t2 = South\n\t3 = West\n\t4 = Hornsby\n\t5 = Bankstown\n\t6 = East Hills\n\t7 = City\n\n" << endl;
		cout << "Enter number:";
		cin >> run;
		cout << "\n" << run << "\n\n" << endl;
		
		//inject the processing dll
		goto Testing;
	}
	
	// default message
	if (argc < 2){
		
		if (!IsProcessRunning("AIMS_Office.exe")){
			cout << "AIMS not running...\nStarting AIMS..." << endl;
			if (startup("C:\\Program Files (x86)\\GeomaticTechnologies AIMS Office v3.0\\AIMS_Office.exe")){
				HWND aimsThrottleCheck;
				while ((aimsThrottleCheck = FindWindow("#32770","Throttle device check ...")) == NULL) {
					cout << "waiting...";
					Sleep(50);
					cout << "\r";
				}
				cout << endl;
				HWND okBtn = FindWindowExA(aimsThrottleCheck, NULL, "Button", "OK");
				SendMessage(okBtn, BM_CLICK, 0, 0);
			}else{
				cout << "Unable to start AIMS. Please start AIMS manually and run the injector afterwards." << endl;
			}
		}

		/*
		cout << "OPTIONS:\n "
				"\tcheck_tracks resize-only [only resize the AIMS track selection window and exit]\n"
				"\tcheck_tracks desktop-path [used only by VBS script to get the current user's desktop path]\n"
				"\n"
				"no option specified - will run full app...\n\n" << endl;
		*/
	}
	
	// had an issue with using VBS to get the desktop path for current user
	// the %USERPROFILE% var returned C:\users\ssingh1\Desktop
	// but the true desktop path is more like \\sdcwcappc056.corp.trans.internal\TRAINS01\ssingh1
	// so we need to use C++ to get this true path and the VBS script will take the output and
	// use it to create the shortcut on the desktop
	if (argc > 1 && argv[1] == std::string("desktop-path")) {
		//MessageBox(0,desktop_directory().c_str(),"Desktop",0);
		cout << desktop_directory() << endl;
		return 0;
	}
	
	/*
	// exit if target AIMS select track or AIMS Review window not found
	if (!targetWnd && !aimsReviewWindow){
		//set console output color to red
		//see https://stackoverflow.com/a/4053879 for color codes
		SetConsoleTextAttribute(hConsole, 12);
		cout << "The track selection window or AIMS review window is not open.\nOpen the window and try again DOOFUS!\nexiting...\n\n" << endl;
		//set console output color back to white
		SetConsoleTextAttribute(hConsole, 15);

		system("PAUSE");
		return 0;
	}
	
	//if the review window is open set the caption
	if (aimsReviewWindow)
		szCaption = "A.I.M.S  3.4.6 - Main";

	
	//if the select tracks window is open the following code gets excuted only
	//the first time the app runs
	if (targetWnd)
	{
		// set size of main window
		SetWindowPos(targetWnd, 0, 200, 90, 1342, 1300, SWP_SHOWWINDOW);
		
		// hide unneeded controls
		ShowWindow(cancelBtn,1);
		ShowWindow(checkBox,0);
		ShowWindow(backBtn,0);
		ShowWindow(browseBtn,0);
		
		// resize controls
		SetWindowPos(tabs, 0, 680, 10, width, height, SWP_SHOWWINDOW);
		SetWindowPos(flexGrid, 0, 0, 0, width, height, SWP_SHOWWINDOW);
		SetWindowPos(selectBtn, 0, 4, 323, 657, 33, SWP_SHOWWINDOW);

		SetWindowPos(cancelBtn, 0, 4, 575, 657, 683, SWP_SHOWWINDOW);
		SendMessage(cancelBtn, WM_SETTEXT, 0, (LPARAM)"Singh Moded\n\n2022");
		EnableWindow(cancelBtn,0);

		UpdateWindow(targetWnd);
	}
	
	// if the user only wants to make the track selection window bigger then
	// check if an argument was passed otherwise continue injecting DLL
	// Check the number of parameters
    if (argc > 1 && argv[1] == std::string("resize-only")) {
		// testing output args
		//cout << argv[0] << " " << argv[1] <<endl;
		//set output color to green
		SetConsoleTextAttribute(hConsole, 10);
		cout << "Exiting app as only window resize requested\n\n" << endl;
		//set console output color back to white
		SetConsoleTextAttribute(hConsole, 15);
		
		system("PAUSE");
		return 0;
	}
	*/
	
	//jump to this block
	Testing:
	// injection of DLL starts here
	unsigned long threadID = GetTargetThreadIdFromWindow(szClass, szCaption); // supply the class and window title
	//unsigned long threadID = 0;
	
	//set output color to green
	SetConsoleTextAttribute(hConsole, 10);
	//print KME/AIMS app id
	(processing) ? printf("KME Process ID: %i\n", threadID) : printf("AIMS Review Process ID: %i\n", threadID);
	
	// get id of this app and pass it to the dll
	unsigned long currentAppID = GetCurrentProcessId();
	printf("This (injector) App's ID: %i\n", currentAppID);
    
	//print instructions as required
	(processing) ? printf("\n\nInstructions:\n\t1. Load Session\n\t2. Select Track\n\t3. Select Starting GPS\n\t4. HIT 'ESC' key on keyboard to go to end of selected track\n\n") : printf("\n\nIn Track Selection window to copy processed tracks highlight any track and press '`'\nPress the ESC key to toggle POIs\nDouble click POI Detail window to edit and close POIs\n\n");
	//set console output color back to white
	SetConsoleTextAttribute(hConsole, 15);

    HINSTANCE hinst = LoadLibrary(szDllName);

    if (hinst) {
        typedef void (*Install)(unsigned long,unsigned long);
        typedef void (*Uninstall)();

        Install install = (Install)GetProcAddress(hinst, "install");
        Uninstall uninstall = (Uninstall)GetProcAddress(hinst, "uninstall");
		
        install(threadID,currentAppID);

        MSG msg = {};

        while (GetMessage(&msg, NULL, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        uninstall();
    }else{
		//set console output color to red
		SetConsoleTextAttribute(hConsole, 12);
		cout << "Unable to load library (" << szDllName << ")\n" << endl;
		//set console output color back to white
		SetConsoleTextAttribute(hConsole, 15);
	}

	return 0;
}
