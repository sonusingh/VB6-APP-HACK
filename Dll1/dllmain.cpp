// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include "detours.h"
#include <tchar.h>
#include <atlstr.h>
#include <atlwin.h>
#include <string>
#include <fstream>
#include <vector>
#include <Windows.h>
#include "resource.h"
using namespace std;

//linker to detours.lib
#pragma comment(lib, "detours.lib")

//need to use this otherwise we get the following error:
//error C4996: 'strcat': This function or variable may be unsafe. Consider using strcat_s instead. To disable deprecation, use _CRT_SECURE_NO_WARNINGS
#pragma warning(disable:4996)
//ignore the "xyz could be 0"...blah blah
#pragma warning(disable:6387)

HINSTANCE hinst;
#pragma data_seg(".shared")
HHOOK g_hHook;
#pragma data_seg()

//since we are using MFC when linking it will complain of DLLMAIN
//already being linked from MFC so use this to tell the linker to use our DLLMAIN
extern "C" { int _afxForceUSRDLL; }

// Function pointer to the original (un-detoured) Text APIs
int (WINAPI* Real_DrawText)
    (HDC a0, LPCWSTR a1, int a2, LPRECT a3, UINT a4) = DrawTextW;

BOOL (WINAPI* Real_TextOut)
    (__in HDC hdc,__in int nXStart,__in int nYStart,__in LPCTSTR lpString,__in int cbString) = TextOut;

int (WINAPI* Real_DrawTextEx)
    (__in HDC hdc,__inout LPTSTR lpchText,__in int cchText,__inout LPRECT lprc,__in UINT dwDTFormat,__in LPDRAWTEXTPARAMS lpDTParams) = DrawTextEx;

BOOL(WINAPI* Real_ExtTextOut)
    (__in HDC hdc,__in int X,__in int Y,__in UINT fuOptions,__in const RECT* lprc,__in LPCTSTR lpString,__in UINT cbCount,__in const INT* lpDx) = ExtTextOut;

/*


Helper functions


*/
//debug message viewer
void showMessageBox(LPCSTR msg, LPCSTR title)
{
    MessageBox(HWND_DESKTOP, msg, title, MB_OK);
}

// returns the text value of HWND (WINDOW HANDLE)
std::string HWNDToString(HWND input)
{
    std::string output = "";
    size_t sizeTBuffer = GetWindowTextLength(input) + 1;

    if (sizeTBuffer > 0)
    {
        output.resize(sizeTBuffer);
        sizeTBuffer = GetWindowText(input, &output[0], sizeTBuffer);
        output.resize(sizeTBuffer);
    }

    return output;
}

//write output to text file
void WriteToFile(std::string text, HDC hDC = nullptr)
{
    //output to text file
    std::ofstream outfile;
    outfile.open("textout.txt", std::ios_base::app); // ,std::fstream::trunc -> truncate mode ,std::ios_base::app -> append mode
    //outfile << mainWindowName << " -> " << windowName << " -> " << text << "\n";
    outfile << text << "\n";
    outfile.close();
}

//convert LPCSTR to string
std::string convertToStr(LPCSTR str) {
    return std::string(str);
}


//get real parent window of control
HWND GetRealParent(HWND hWnd)
{
    HWND hWndOwner;

    // To obtain a window's owner window, instead of using GetParent,
    // use GetWindow with the GW_OWNER flag.

    if (NULL != (hWndOwner = GetWindow(hWnd, GW_OWNER)))
        return hWndOwner;

    // Obtain the parent window and not the owner
    return GetAncestor(hWnd, GA_PARENT);
}



/*


Detoured text functions


*/
int WINAPI Mine_DrawText(__in HDC hDC, __inout LPCWSTR lpchText, __in int nCount, __inout LPRECT lpRect, __in UINT uFormat)
{
    //convert to LPCSTR
    LPCSTR text = CW2A(lpchText);
    std::string stext = convertToStr(text);
    //showMessageBox(text, "DrawText");

    //ouput
    //WriteToFile(hDC, "DrawText -> " + stext);

    return Real_DrawText(hDC, lpchText, nCount, lpRect, uFormat);
}

BOOL WINAPI Mine_TextOut(__in HDC hdc,__in int nXStart,__in int nYStart,__in LPCTSTR lpString,__in int cbString)
{
    if (!convertToStr(lpString).empty()) {
        std::string stext = convertToStr(lpString);
        WriteToFile(stext, hdc);   
    }

    return Real_TextOut(hdc, nXStart, nYStart, lpString, cbString);
}

int WINAPI Mine_DrawTextEx(__in HDC hdc,__inout LPTSTR lpchText,__in int cchText,__inout LPRECT lprc,__in UINT dwDTFormat,__in LPDRAWTEXTPARAMS lpDTParams)
{
    return Real_DrawTextEx(hdc, lpchText, cchText, lprc, dwDTFormat, lpDTParams);
}

BOOL WINAPI Mine_ExtTextOut(__in HDC hdc,__in int X,__in int Y,__in UINT fuOptions,__in const RECT* lprc,__in LPCTSTR lpString,__in UINT cbCount,__in const INT* lpDx)
{
    return Real_ExtTextOut(hdc, X, Y, fuOptions, lprc, lpString, cbCount, lpDx);
}

//store the our injector app's process ID in text file for later access
void storeInjAppID(unsigned long processID) {
    //output to text file
    std::ofstream outfile;
    outfile.open("pid.txt", std::fstream::trunc); // truncate file (ie clear contents)
    outfile << processID;
    outfile.close();

}

//retrieve injector app's id from text file
unsigned long getID() {
    int sum = 0;
    int x;
    std::ifstream inFile;

    inFile.open("pid.txt");
    if (!inFile) {
        showMessageBox("Unable to open C:\\Users\\Public\\Documents\\pid.txt", "Unable to open PID file");
    }
    else {
        while (inFile >> x) {
            sum = sum + x;
        }
        inFile.close();
    }

    return sum;
}


/*
        This code is from Textify

        https://github.com/m417z/Textify

        TextDlg.cpp
        CTextDlg::OnInitDialog
*/
//this extracts the text from the control that is currently under the mouse pointer (cursor)
void GetAccessibleInfoFromPoint(POINT pt, CWindow& window, CString& outString, CRect& outRc, std::vector<int>& outIndexes)
{
    outString = L"(no text could be retrieved)";
    outRc = CRect{ pt, CSize{ 0, 0 } };
    outIndexes = std::vector<int>();

    HRESULT hr;

    CComPtr<IAccessible> pAcc;
    CComVariant vtChild;
    hr = AccessibleObjectFromPoint(pt, &pAcc, &vtChild);
    if (FAILED(hr))
    {
        return;
    }

    hr = WindowFromAccessibleObject(pAcc, &window.m_hWnd);
    if (FAILED(hr))
    {
        return;
    }

    DWORD processId = 0;
    GetWindowThreadProcessId(window.m_hWnd, &processId);

    while (true)
    {
        CString string;
        std::vector<int> indexes;

        CComBSTR bsName;
        hr = pAcc->get_accName(vtChild, &bsName);
        if (SUCCEEDED(hr) && bsName.Length() > 0)
        {
            string = bsName;
        }

        CComBSTR bsValue;
        hr = pAcc->get_accValue(vtChild, &bsValue);
        if (SUCCEEDED(hr) && bsValue.Length() > 0 &&
            bsValue != bsName)
        {
            if (!string.IsEmpty())
            {
                string += L"\r\n";
                indexes.push_back(string.GetLength());
            }

            string += bsValue;
        }

        CComVariant vtRole;
        hr = pAcc->get_accRole(CComVariant(CHILDID_SELF), &vtRole);
        if (FAILED(hr) || vtRole.lVal != ROLE_SYSTEM_TITLEBAR) // ignore description for the system title bar
        {
            CComBSTR bsDescription;
            hr = pAcc->get_accDescription(vtChild, &bsDescription);
            if (SUCCEEDED(hr) && bsDescription.Length() > 0 &&
                bsDescription != bsName && bsDescription != bsValue)
            {
                if (!string.IsEmpty())
                {
                    string += "\r\n";
                    indexes.push_back(string.GetLength());
                }

                string += bsDescription;
            }
        }

        if (!string.IsEmpty())
        {
            // Normalize newlines.
            string.Replace("\r\n", "\n");
            string.Replace("\r", "\n");
            string.Replace("\n", "\r\n");

            string.TrimRight();

            if (!string.IsEmpty())
            {
                outString = string;
                outIndexes = indexes;
                break;
            }
        }

        if (vtChild.lVal == CHILDID_SELF)
        {
            CComPtr<IDispatch> pDispParent;
            HRESULT hr = pAcc->get_accParent(&pDispParent);
            if (FAILED(hr) || !pDispParent)
            {
                break;
            }

            CComQIPtr<IAccessible> pAccParent(pDispParent);
            pAcc.Attach(pAccParent.Detach());

            HWND hWnd;
            hr = WindowFromAccessibleObject(pAcc, &hWnd);
            if (FAILED(hr))
            {
                break;
            }

            DWORD compareProcessId = 0;
            GetWindowThreadProcessId(hWnd, &compareProcessId);
            if (compareProcessId != processId)
            {
                break;
            }
        }
        else
        {
            vtChild.lVal = CHILDID_SELF;
        }
    }

    long pxLeft, pyTop, pcxWidth, pcyHeight;
    hr = pAcc->accLocation(&pxLeft, &pyTop, &pcxWidth, &pcyHeight, vtChild);
    if (SUCCEEDED(hr))
    {
        outRc = CRect{ CPoint{ pxLeft, pyTop }, CSize{ pcxWidth, pcyHeight } };
    }
}

//enable edit box and open/close in POI Detail window
//and if EAM defects window open and user double click on EAM ID open
//browser with SAP defect ID
void controlProperties() 
{
    //get location of mouse cursor
    CPoint pt;
    GetCursorPos(&pt);

    //this code is from Textify
    //TextDlg.cpp
    //CTextDlg::OnInitDialog
    CWindow wndAcc;
    CString strText;
    CRect rcAccObject;
    std::vector<int> m_editIndexes;
    int m_lastSelEnd = 0;

    //call the method to fill the strText from control currently under the mouse
    GetAccessibleInfoFromPoint(pt, wndAcc, strText, rcAccObject, m_editIndexes);

    //testing
    showMessageBox(strText, "Text");    
    
    return;
}

//get processed tracks method
void getProcessedTracks(HWND targetWnd, int sort = 0) {
    //class name placeholder
    char szClassName[128];
    memset(szClassName, 0, sizeof(szClassName));

    //child of main window
    HWND tabs = FindWindowExA(targetWnd, NULL, "SSTabCtlWndClass", NULL);
    //child of control
    //HWND flexGridHWND = FindWindowExA(tabs, NULL, "MSFlexGridWndClass", NULL);
    HWND flexGridHWND = FindWindowExA(targetWnd, NULL, "MSFlexGridWndClass", NULL);

    //if the flexgrid is not found - exit
    if (!flexGridHWND) {
        MessageBox(targetWnd, "MSFlexGridWndClass Not Found!", "Error", MB_OK);
        return;
    }

    //check we have the track selection grid in focus
    ::GetClassName(flexGridHWND, szClassName, sizeof(szClassName) - 1);
    if (strcmp(szClassName, "MSFlexGridWndClass") == 0)
    {
        DWORD* p = (DWORD*) ::GetWindowLong(flexGridHWND, GWL_USERDATA);
        p++; //add 8;
        p++; //add 8;
        LPDISPATCH pDisp = (LPDISPATCH)*p;
        long row, col;

        //ole control driver
        COleDispatchDriver iDriver;

        //release on exit
        iDriver.m_bAutoRelease = TRUE;
        
        //attach to pointer
        iDriver.AttachDispatch(pDisp);

        // getrows()
        iDriver.InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_I4, (void*)&row, NULL);
        // getcols() 
        iDriver.InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_I4, (void*)&col, NULL);
        
        //this is called when the track selection window is displayed upon loading up
        if (sort > 0) {
            static BYTE parmss[] = VTS_I4;
            //setCol() select the column we want to sort
            iDriver.InvokeHelper(0xb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parmss, sort);

            static BYTE parmsss[] = VTS_I2;
            //setSort() sort the column by Numeric Ascending
            int s;
            (sort == 4) ? s = 5 : s = 3;
            iDriver.InvokeHelper(0x2e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parmsss, s);

            //note the last param for setSort is a number based on:
            /*
            flexSortNone 0 flexSortGenericAscending 1 flexSortGenericDescending 2 flexSortNumericAscending 3
            flexSortNumericDescending 4 flexSortStringNoCaseAsending 5 flexSortNoCaseDescending 6 flexSortStringAscending 7 flexSortStringDescending 8
            flexSortCustom 9
            */

            //release driver
            iDriver.DetachDispatch();
            iDriver.ReleaseDispatch();
            //delete iDriver;
            iDriver = NULL;

            //delete p, pDisp;
            p, pDisp = NULL;

            //only wanted to sort the grid - exit
            return;
        }

        //store the returned text in the given row & col
        CString result;
        static BYTE parms[] = VTS_I4;

        //declare data dir var
        TCHAR appData[MAX_PATH];

        //output to csv file
        std::ofstream outfile;
        outfile.open("processed_tracks.csv", std::fstream::trunc); // truncate file (ie clear contents)
        //MessageBox(NULL, appData, "Output File", MB_OK);
        //outfile << text << "\n";

        //loop through the rows
        for (int i = 0; i < row; i++) {
            //if first row don't insert new line
            if (i == 0) {
                //do nothing
            }
            else {
                outfile << "\n";
            }
            //loop throught he columns
            for (int j = 0; j < col; j++)
            {
                iDriver.InvokeHelper(0x37, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, col * i + j);
                outfile << result << "\t";
                //showMessageBox(result, "Data");
            }
        }

        //release driver
        iDriver.DetachDispatch();
        iDriver.ReleaseDispatch();
        //delete iDriver;
        iDriver = NULL;

        //delete p, pDisp;
        p, pDisp = NULL;

        //close the csv file
        outfile.close();

        /*
        //USE THIS CODE TO CLOSE THE injecting APP
        //test what is being saved in PID.txt file
        //auto ss = std::to_string(getID());
        //MessageBox(pMsg->hwnd, ss.c_str(), "Process ID", MB_OK);

        // exit the injector app
        // see https://stackoverflow.com/a/1916668
        const auto inj_app = OpenProcess(PROCESS_TERMINATE, false, getID());
        TerminateProcess(inj_app, 1);
        CloseHandle(inj_app);

        */

        //open the folder containing the xls/csv file
        //and select/highlight the file named Get_Processed_Tracks
        //BrowseToFile("C:\\Users\\Public\\Documents\\", "C:\\Users\\Public\\Documents\\Get_Processed_Tracks.xlsm");
        
        //INSTEAD OF OPENING THE FOLDER ABOVE we will start the excel file
        ShellExecute(NULL, NULL, "processed_tracks.csv", NULL, NULL, SW_SHOWDEFAULT);
    }
}
/* --------------------------------------------------------------------------- 



            THE MAIN FUNCTION CONTROL CENTRE



------------------------------------------------------------------------------ */
//custom GetMessageProcedure
LRESULT CALLBACK GetMsgProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode < 0)
    {
        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }

    //message param
    PMSG pMsg = (PMSG)lParam;

    //freopen("output.txt", "a", stdout);
    //cout << (PMSG)wParam << endl;

    //someone double click?
    if (pMsg->message == WM_LBUTTONDBLCLK)
    {
        //use textify code to get text
        controlProperties();

        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }

    //capture keyboard input
    if (pMsg->message == WM_KEYUP)
    {
        //use ESC key to turn POIs on/off
        if (pMsg->wParam == VK_ESCAPE)
        {
            //check if select track window open
            HWND targetWnd = FindWindow("ThunderRT6FormDC", "Form1");
            if (targetWnd) {
                MessageBox(targetWnd, "yeah is injected", "success", MB_OK);
                getProcessedTracks(targetWnd);
            }
            return CallNextHookEx(NULL, nCode, wParam, lParam);
        }
        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}



// Install the DrawText detour whenever this DLL is loaded into any process...
BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        //reuired for creating hook
        //and closing the dll on un-hooking
        hinst = hModule;
        DetourTransactionBegin();
        DetourUpdateThread(GetCurrentThread());
        DetourAttach(&(PVOID&)Real_TextOut, Mine_TextOut); // <- magic
        DetourTransactionCommit();
        break;

    case DLL_PROCESS_DETACH:
        DetourTransactionBegin();
        DetourUpdateThread(GetCurrentThread());
        DetourDetach(&(PVOID&)Real_TextOut, Mine_TextOut);
        DetourTransactionCommit();
        break;
    }

    return TRUE;
}

//called by our injector app
extern "C" __declspec(dllexport) void install(unsigned long threadID, unsigned long injectorAppID) {
    //store the injector app's process id
    storeInjAppID(injectorAppID);
    g_hHook = SetWindowsHookEx(WH_GETMESSAGE, (HOOKPROC)GetMsgProc, hinst, threadID);
}

//called when the injector app is closed
extern "C" __declspec(dllexport) void uninstall() {
    UnhookWindowsHookEx(g_hHook);
}