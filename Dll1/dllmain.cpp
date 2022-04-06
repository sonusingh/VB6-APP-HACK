// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include "detours.h"
#include <tchar.h>
#include <atlstr.h>
#include <atlwin.h>
#include <string>
#include <fstream>
#include <vector>

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

//handle to the app that is to be injected into
HWND targetWnd = FindWindow("ThunderRT6FormDC", "Select Track Section");
//child of main window
HWND tabs = FindWindowExA(targetWnd, NULL, "SSTabCtlWndClass", NULL);
//chile of control
HWND flexGridHWND = FindWindowExA(tabs, NULL, "MSFlexGridWndClass", NULL);

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
void WriteToFile(HDC hDC, std::string text)//, std::string handle = "NO_HANDLE", std::string fnName = "NO_FUNCTION_NAME") 
{
    /*
    //this will get the parent window (control) where the text is being written
    HWND hWindow = WindowFromDC(hDC);

    //this will get the parent of the control (not required)
    HWND mainWindow = ::GetParent(hWindow);

    //get window and control name
    std::string windowName = HWNDToString(hWindow).c_str();
    std::string mainWindowName = HWNDToString(mainWindow).c_str();
    */

    //output to text file
    std::ofstream outfile;
    outfile.open("C:\\Users\\Public\\Documents\\kms.txt", std::fstream::trunc); // truncate mode
    //outfile << mainWindowName << " -> " << windowName << " -> " << text << "\n";
    outfile << text << "\n";
    outfile.close();
}


//check if string has numbers
bool has_any_digits(const std::string& s)
{
    if (std::string::npos != s.find_first_of("0123456789"))
    {
        //std::cout << "digit(s)found!" << std::endl;
        return true;
    }
    return false;
}

//convert LPCSTR to string
std::string convertToStr(LPCSTR str) {
    return std::string(str);
}



//open menu
HMENU findMenu(HMENU menu, LPARAM menuText) //TCHAR menuText[ 2048 ] = "File";
{
    int menuCount = GetMenuItemCount(menu);
    for (int i = 0; i < menuCount; i++)
    {
        char menuString[2048];
        GetMenuString(menu, i, menuString, 2048, MF_BYPOSITION);
        if (_tcsstr(menuString, LPCTSTR(menuText)) != NULL)
        {
            //cout << "FOUND:" << menuString << endl;
            return GetSubMenu(menu, i);
        }
    }
    return NULL;
}

//select a menu item
void selectMenuItem(HWND hwnd, HMENU menu, LPARAM menuText)
{
    int menuCount = GetMenuItemCount(menu);
    for (int i = 0; i < menuCount; i++)
    {
        char menuString[2048];
        GetMenuString(menu, i, menuString, 2048, MF_BYPOSITION);
        //cout << "MENUSTRING:" << menuString << endl;
        if (_tcsstr(menuString, LPCTSTR(menuText)) != NULL)
        {
            int menuItemID = GetMenuItemID(menu, i);
            PostMessage(hwnd, WM_COMMAND, menuItemID, 0);
            //cout << "MENUSELECTED:" << menuItemID << ":OF:" << menuCount << endl;
            return;
        }
    }
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
    COLORREF redcolor = RGB(255, 0, 0);//red
    COLORREF blackcolor = RGB(0, 0, 0);//black

    CString cszMyString = lpString;
    int strLength = cszMyString.GetLength();
    
    //this will get the parent window (control) where the text is being written
    HWND hWindow = WindowFromDC(hdc);
    //this will get the parent of the control (not required)
    //HWND mainWindow = ::GetParent(hWindow);
    
    //stub to see what the container window is where the text is being written
    //showMessageBox(HWNDToString(hWindow).c_str(), "Window");
    
    //filter out textout only for "Location" control window
    //and skip if textout is blank
    //and skip is textout does not contain numbers
    if (HWNDToString(hWindow).compare("Location") == 0 && !convertToStr(lpString).empty() && has_any_digits(lpString)) {
        //showMessageBox(HWNDToString(hWindow).c_str(), "Window");
        //showMessageBox(lpString, "TextOut");
        std::string stext = convertToStr(lpString);
        WriteToFile(hdc, stext);
        
        //here we check to see if POIs are turned off
        //if they are then show KMS in red colour
        //check if review window open
        HWND aimsAPP = FindWindow("ThunderRT6FormDC", "A.I.M.S  3.4.6 - Main");
        if (aimsAPP)
        {
            //get menu
            HMENU fileMenu = findMenu(GetMenu(aimsAPP), (LPARAM)TEXT("View"));

            //store menu info
            MENUITEMINFO info;
            info.cbSize = sizeof(info);
            info.fMask = MIIM_STATE;

            if (!GetMenuItemInfo(fileMenu, 4, TRUE, &info))
                showMessageBox("couldn't get info about view menu", "menu checked/un-checked");

            //get state of menu item ie checked/un-checked
            if (info.fState == MF_CHECKED)
            {
                SetTextColor(hdc, blackcolor);
            }
            else {
                SetTextColor(hdc, redcolor);
            }
        }
    }
    else {
        //std::string stext = convertToStr(lpString);
        //WriteToFile(hdc, "TextOut -> " + stext);
        //showMessageBox(lpString, "outtext");
    }

    //EAM defect IDs are 12 digits long
    //unfortunatley the EAM dailogs use WPF and there
    //text out is not detected
    //the following code only works on POI Detail window (VB6)
    if (strLength == 12) {
        //showMessageBox("12 chars long", "outtext");
        //EAM ID only has numbers no letters
        if (cszMyString.SpanIncluding("0123456789") == cszMyString) {
            //MessageBox(NULL, cszMyString, "ID", MB_OK);
            //showMessageBox(HWNDToString(hWindow).c_str(), "Window");
        }
    }

    return Real_TextOut(hdc, nXStart, nYStart, lpString, cbString);
}

int WINAPI Mine_DrawTextEx(__in HDC hdc,__inout LPTSTR lpchText,__in int cchText,__inout LPRECT lprc,__in UINT dwDTFormat,__in LPDRAWTEXTPARAMS lpDTParams)
{
    //convert to string
    std::string stext = convertToStr(lpchText);

    //ouput
    //WriteToFile(hdc, "DrawTextEx -> " + stext);

    //showMessageBox(lpchText, "DrawTextEx");
    return Real_DrawTextEx(hdc, lpchText, cchText, lprc, dwDTFormat, lpDTParams);
}

BOOL WINAPI Mine_ExtTextOut(__in HDC hdc,__in int X,__in int Y,__in UINT fuOptions,__in const RECT* lprc,__in LPCTSTR lpString,__in UINT cbCount,__in const INT* lpDx)
{
    //convert to string
    std::string stext = convertToStr(lpString);

    //ouput
    //WriteToFile(hdc, "ExtTextOut -> " + stext);

    //showMessageBox(lpString, "ExtTextOut");
    return Real_ExtTextOut(hdc, X, Y, fuOptions, lprc, lpString, cbCount, lpDx);
}

//store the our injector app's process ID in text file for later access
void storeInjAppID(unsigned long processID) {
    //output to text file
    std::ofstream outfile;
    outfile.open("C:\\Users\\Public\\Documents\\pid.txt", std::fstream::trunc); // truncate file (ie clear contents)
    outfile << processID;
    outfile.close();

}

//retrieve injector app's id from text file
unsigned long getID() {
    int sum = 0;
    int x;
    std::ifstream inFile;

    inFile.open("C:\\Users\\Public\\Documents\\pid.txt");
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

//open folder and select final output file
//see https://stackoverflow.com/a/47564530
void BrowseToFile(LPCTSTR path, LPCTSTR pathNfilename)
{

    HRESULT hr;
    hr = CoInitializeEx(0, COINIT_MULTITHREADED);

    ITEMIDLIST* folder = ILCreateFromPath("C:\\Users\\Public\\Documents\\");
    std::vector<LPITEMIDLIST> v;
    v.push_back(ILCreateFromPath("C:\\Users\\Public\\Documents\\Get_Processed_Tracks.xlsm"));

    SHOpenFolderAndSelectItems(folder, v.size(), (LPCITEMIDLIST*)v.data(), 0);

    for (auto idl : v)
    {
        ILFree(idl);
    }
    ILFree(folder);
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
    HWND targetWnd = FindWindow("ThunderRT6FormDC", "POI Detail");
    HWND frame = FindWindowExA(targetWnd, NULL, "ThunderRT6Frame", "POI Data");
    HWND poiDescriptionEditBox = GetDlgItem(frame, 8);
    HWND openCloseDropDown = GetDlgItem(frame, 9);

    //did we find the contorls?
    if (targetWnd && frame) {
        if (poiDescriptionEditBox || openCloseDropDown) {
            //if disabled then enable them
            if (!IsWindowEnabled(poiDescriptionEditBox)) {
                EnableWindow(poiDescriptionEditBox, 1);
            }
            if (!IsWindowEnabled(openCloseDropDown)) {
                EnableWindow(openCloseDropDown, 1);
            }
        }
    }

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
    //showMessageBox(strText, "Defect ID");
    
    //Defect can only contain numbers
    if (strText.SpanIncluding("0123456789") == strText) {}
    //return _T("string only numeric characters");
    else {
        //return _T("string contains non numeric characters");
        return;
    }

    //check that the copied text is 12 digits long
    //anything more or less then exit
    m_lastSelEnd = strText.GetLength();

    if (m_lastSelEnd != 12) {
        return;
    }

    //open the browser window with the SAP defect
    ShellExecute(NULL, "open", "https://ecc.transport.nsw.gov.au/sap/bc/gui/sap/its/webgui;~sysid=TPE;~service=3200?~transaction=*iw22 RIWO00-QMNUM=" + strText + ";OKCODE=SHOW", NULL, NULL, SW_SHOWNORMAL);
    //"https://ecc.transport.nsw.gov.au/sap/bc/gui/sap/its/webgui;~sysid=TPE;~service=3200?~transaction=*iw22 RIWO00-QMNUM=" + notificationID.Text + ";OKCODE=SHOW"

    
    return;

    
}

/* --------------------------------------------------------------------------- 



            THE MAIN FUNCTION CONTROL CENTRE



------------------------------------------------------------------------------ */
//custom cbtproc
LRESULT CALLBACK CBTProcedure(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode < 0)
    {
        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }

    //message param
    PMSG pMsg = (PMSG)lParam;

    //class name placeholder
    char szClassName[128];
    memset(szClassName, 0, sizeof(szClassName));

    //someone double click?
    if (pMsg->message == WM_LBUTTONDBLCLK)
    {
        controlProperties();
        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }

    //capture keyboard input
    if (pMsg->message == WM_KEYUP)
    {
        //use ESC key to turn POIs on/off
        if (pMsg->wParam == VK_ESCAPE)
        {
            // Select View->POI Markers
            HWND aimsAPP = FindWindow("ThunderRT6FormDC", "A.I.M.S  3.4.6 - Main");
            if (aimsAPP)
            {
                HMENU fileMenu = findMenu(GetMenu(aimsAPP), (LPARAM)TEXT("View"));
                selectMenuItem(aimsAPP, fileMenu, (LPARAM)TEXT("POI Markers"));
                Sleep(100);
                // Select View->Refresh
                HMENU refreshMenu = findMenu(GetMenu(aimsAPP), (LPARAM)TEXT("View"));
                selectMenuItem(aimsAPP, refreshMenu, (LPARAM)TEXT("Refresh"));

            }
            return CallNextHookEx(NULL, nCode, wParam, lParam);
        }

        //get tracks processed from track selection window
        if (pMsg->wParam == '`')
        {

            //check we have the track selection grid in focus
            ::GetClassName(pMsg->hwnd, szClassName, sizeof(szClassName) - 1);
            if (strcmp(szClassName, "MSFlexGridWndClass") == 0)
            {
                DWORD* p = (DWORD*) ::GetWindowLong(flexGridHWND, GWL_USERDATA);
                p++; //add 8;
                p++; //add 8;
                LPDISPATCH pDisp = (LPDISPATCH)*p;
                long row, col;
                COleDispatchDriver iDriver;
                iDriver.m_bAutoRelease = TRUE;

                iDriver.AttachDispatch(pDisp);
                // getrows() 
                iDriver.InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_I4, (void*)&row, NULL);
                // getcols() 
                iDriver.InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_I4, (void*)&col, NULL);
                CString result;
                static BYTE parms[] = VTS_I4;

                //declare data dir var
                TCHAR appData[MAX_PATH];

                //save to c drive
                strcpy_s(appData, "C:\\Users\\Public\\Documents");

                //output to csv file
                std::ofstream outfile;
                outfile.open(strcat(appData, "\\processed_tracks.csv"), std::fstream::trunc); // truncate file (ie clear contents)
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
                iDriver.ReleaseDispatch();
                delete iDriver;

                //close the csv file
                outfile.close();

                //let the user know we are done
                //it appears if the user clicks multiple times it crashes the app
                //so let them do it once and exit
                /*
                MessageBox(pMsg->hwnd,
                    "The data has been copied to C:\\Users\\Public\\Documents\\processed_tracks.csv\n"
                    "Run the Get_Processed_Tracks.xlsm to load the tracks!\n\n"
                    "The injected app will now exit!\n\n"
                    "-------------------------------\n\n"
                    "Source code @ sukhpalsingh OneDrive",
                    "Data Copy Finished",
                    MB_OK);
                */

                //open the folder containing the xls/csv file
                //and select/highlight the file named Get_Processed_Tracks
                BrowseToFile("C:\\Users\\Public\\Documents\\", "C:\\Users\\Public\\Documents\\Get_Processed_Tracks.xlsm");

                //test what is being saved in PID.txt file
                //auto ss = std::to_string(getID());
                //MessageBox(pMsg->hwnd, ss.c_str(), "Process ID", MB_OK);

                // exit the injector app
                // see https://stackoverflow.com/a/1916668
                //const auto inj_app = OpenProcess(PROCESS_TERMINATE, false, getID());
                //TerminateProcess(inj_app, 1);
                //CloseHandle(inj_app);

                return CallNextHookEx(NULL, nCode, wParam, lParam);
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
        //MessageBox(HWND_DESKTOP, "activated", "test", MB_OK);
        DetourTransactionBegin();
        DetourUpdateThread(GetCurrentThread());
        //the KMs are outputted by TextOut
        DetourAttach(&(PVOID&)Real_TextOut, Mine_TextOut); // <- magic
        //other info may use drawtext etc
        //use the following if needed
        //DetourAttach(&(PVOID&)Real_DrawText, Mine_DrawText); 
        //DetourAttach(&(PVOID&)Real_DrawTextEx, Mine_DrawTextEx);
        //DetourAttach(&(PVOID&)Real_ExtTextOut, Mine_ExtTextOut);
        DetourTransactionCommit();
        break;

    case DLL_PROCESS_DETACH:
        DetourTransactionBegin();
        DetourUpdateThread(GetCurrentThread());
        DetourDetach(&(PVOID&)Real_TextOut, Mine_TextOut);
        //use the following if needed
        //DetourDetach(&(PVOID&)Real_DrawText, Mine_DrawText);
        //DetourDetach(&(PVOID&)Real_DrawTextEx, Mine_DrawTextEx);
        //DetourDetach(&(PVOID&)Real_ExtTextOut, Mine_ExtTextOut);
        DetourTransactionCommit();
        break;
    }

    return TRUE;
}

//called by our injector app
extern "C" __declspec(dllexport) void install(unsigned long threadID, unsigned long injectorAppID) {
    //store the injector app's process id
    storeInjAppID(injectorAppID);
    g_hHook = SetWindowsHookEx(WH_GETMESSAGE, (HOOKPROC)CBTProcedure, hinst, threadID);
}

//called when the injector app is closed
extern "C" __declspec(dllexport) void uninstall() {
    UnhookWindowsHookEx(g_hHook);
}