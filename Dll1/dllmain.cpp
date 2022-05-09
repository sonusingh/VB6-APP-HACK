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
HHOOK cbt_hHook;
HHOOK WndProc_hHook;
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
    outfile.open("C:\\Users\\Public\\Documents\\kms.txt", std::fstream::trunc); // ,std::fstream::trunc -> truncate mode ,std::ios_base::app -> append mode
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
        WriteToFile(stext, hdc);
        
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
    
    return;
}

//set default font for all child controls in window
bool CALLBACK SetFont(HWND child, LPARAM font) {
    SendMessage(child, WM_SETFONT, font, true);
    return true;
}

//global window display control
//these are reuqired if creating a pop window manually
bool windowCreated = false;
HWND windowHandle;

//control IDs for checkboxes
//these for manually created controls
#define srs2t1 	1500
#define srs3t3 	1502
#define srs3t5 	1504


//just for the fun of it add an image to the track selection window
void LoadScreen(HWND hWnd) {
    //to load bmp from file
    //HBITMAP hBitmap = (HBITMAP)LoadImage(NULL, "C:\\Users\\ssingh1\\Downloads\\MinGW\\code\\fatherson.bmp", IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
    
    //load from resouce
    HBITMAP hBitmap = (HBITMAP)LoadImage(hinst, MAKEINTRESOURCE(IDB_BITMAP1), IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE);
    /*
    DWORD error;
    if (!hBitmap)
    {
        error = GetLastError();
        auto si = std::to_string(error);
        showMessageBox(si.c_str(), "ss");
        //report error
    }
    */
    SendMessage(hWnd, STM_SETIMAGE, (WPARAM)IMAGE_BITMAP, (LPARAM)hBitmap);
    InvalidateRect(hWnd, NULL, FALSE);
}

//open the browser and load SAP iw29 or iw65 screen for the given start and finish KMs and the Basecode
void open_iw29_iw65(HWND hwnd, CString screen) {
    //declare
    TCHAR start_km_text[128], end_km_text[128], basecode_text[128], defect_id[128], turnout_location[128];
    TCHAR turnout_type[128], territory[128], tout_kms[128], tout_points_id[128], tout_basecode[128];

    //get text from edits on dialogbox
    GetDlgItemText(hwnd, start_km, start_km_text, 128);GetDlgItemText(hwnd, end_km, end_km_text, 128);GetDlgItemText(hwnd, basecode, basecode_text, 128);
    GetDlgItemText(hwnd, partial_defect_id, defect_id, 128); GetDlgItemText(hwnd, turnout_functional_location, turnout_location, 128);
    GetDlgItemText(hwnd, turnout_type_drop_down, turnout_type, 128); GetDlgItemText(hwnd, territory_drop_down, territory, 128);
    GetDlgItemText(hwnd, turnout_km, tout_kms, 128); GetDlgItemText(hwnd, points_id, tout_points_id, 128);
    GetDlgItemText(hwnd, turnout_basecode, tout_basecode, 128);

    //need to conver to tchar to CString to use
    //need to change project settings Character Set: Use multi-byte character set
    //see https://stackoverflow.com/a/6006371
    CString s_km_text = start_km_text, e_km_text = end_km_text, b_text = basecode_text;
    CString d_id = defect_id, tout_loc = turnout_location;
    CString tout_type = turnout_type, tout_territory = territory, tout_km = tout_kms, tout_pnt_id = tout_points_id, tout_bcode = tout_basecode;

    //presses the get list of defects for given KMs and basecode either for iw29 or iw65
    if (screen == "iw29" || screen == "iw65"){
        if (s_km_text.IsEmpty() || e_km_text.IsEmpty() || b_text.IsEmpty()) {
            showMessageBox("Start, End and Basecode must be provided to list defects.", "Try Again...");
            return;
        }
        //open the browser window with the SAP defects
        ShellExecute(NULL, "open",
            "https://ecc.transport.nsw.gov.au/sap/bc/gui/sap/its/webgui;~sysid=TPE;~service=3200?~transaction=*" + screen + " S_START-LOW=" + s_km_text + ";S_START-HIGH=" + e_km_text + ";S_LRPID-LOW=*" + b_text + ";OKCODE=SHOW#",
            NULL, NULL, SW_SHOWNORMAL);
    }

    //search for a SAP defect using defect id/partial defect ID
    if (screen == "iw29_srch_defect") {
        if (d_id.IsEmpty()) {
            showMessageBox("Please enter a defect ID or partial defect ID.", "Try Again...");
            return;
        }
        //search iw65 (closed defects as well?)
        if (!IsDlgButtonChecked(hwnd, search_closed_checkbox)) {
            ShellExecute(NULL, "open",
                "https://ecc.transport.nsw.gov.au/sap/bc/gui/sap/its/webgui;~sysid=TPE;~service=3200?~transaction=*iw29 QMNUM-LOW=*" + d_id + "*;S_LRPID-LOW=*;OKCODE=SHOW#",
                NULL, NULL, SW_SHOWNORMAL);
        }
        else {
            ShellExecute(NULL, "open",
                "https://ecc.transport.nsw.gov.au/sap/bc/gui/sap/its/webgui;~sysid=TPE;~service=3200?~transaction=*iw65 QMNUM-LOW=*" + d_id + "*;S_LRPID-LOW=*;OKCODE=SHOW#",
                NULL, NULL, SW_SHOWNORMAL);
        }
    }

    //the user wants to edit sap defect in iw22
    if (screen == "iw22_edit") {
        //make sure defect id is at least 10 digits long
        //e.g. 1001225236
        if (int(d_id.GetLength()) < 10) {
            showMessageBox("Please enter complete SAP defect ID", "Try Again...");
            return;
        }
        else {
            ShellExecute(NULL, "open",
                "https://ecc.transport.nsw.gov.au/sap/bc/gui/sap/its/webgui;~sysid=TPE;~service=3200?~transaction=*iw22 RIWO00-QMNUM=" + d_id + ";OKCODE=SHOW#",
                NULL, NULL, SW_SHOWNORMAL);
        }
    }

    //list defects for a given turnout
    if (screen == "view_defects_for_turnout") {
        if (tout_loc.IsEmpty()) {
            showMessageBox("Please enter the turnout functional location (usually starts with T-TOUT)", "Try Again...");
            return;
        }
        ShellExecute(NULL, "open",
            "https://ecc.transport.nsw.gov.au/sap/bc/gui/sap/its/webgui;~sysid=TPE;~service=3200?~transaction=*iw29 STRNO-LOW=" + tout_loc + ";OKCODE=SHOW#",
            NULL, NULL, SW_SHOWNORMAL);
    }

    //search for a search for a turnout/diamond etc
    if (screen == "find_turnout") {
        //declarations
        CString Territory_code, StartKM, EndKM;
        
        //we need at least the points no. or KMs and Basecode
        if (tout_km.IsEmpty() && tout_pnt_id.IsEmpty() && tout_bcode.IsEmpty()) {
            showMessageBox("Please enter the basecode and KMs or just the points number to look for a turnout.", "Try Again...");
            return;
        }
        //if kms provided add and subtract 0.075 KMs
        if (!tout_km.IsEmpty()) {
            double startKMs = atof(tout_km) - 0.075;
            double endKMs   = atof(tout_km) + 0.075;

            StartKM.Format("%g", startKMs);
            EndKM.Format("%g", endKMs);
        }
        else {
            StartKM, EndKM = tout_kms;
        }

        //get the territory code
        Territory_code = tout_territory.Right(tout_territory.GetLength() - (tout_territory.Find('-') + 1));

        ShellExecute(NULL, "open",
            "https://ecc.transport.nsw.gov.au/sap/bc/gui/sap/its/webgui;~sysid=TPE;~service=3200?~transaction=*ih06 STRNO-LOW=" + tout_type + ";S_START-LOW=" + StartKM + ";S_START-HIGH=" + EndKM + ";S_UOM=ZNK;S_LRPID-LOW=*" + tout_bcode + "*;PLTXT-LOW=*" + tout_pnt_id + "*;GEWRK-LOW=" + Territory_code + ";OKCODE=SHOW#",
            NULL, NULL, SW_SHOWNORMAL);
    }
}

//the dialog from the resource file has a dialog proc to handle events
BOOL CALLBACK DialogProc(HWND hwnd, UINT Message, WPARAM wParam, LPARAM lParam)
{
    //need to manually add items to the drop downs
    HWND t_out, territories = NULL;
    const char* t_out_list[] = { "*T-TOUT*", "*T-DIA*","*T-CPNT*", "*" };
    const char* territory_list[] = { "ALL -*", 
        "Blacktown -NWSCVT1",
        "CBD -NCECVT1", 
        "Clyde -NCWCVT1",
        "Glenfield -NSWCVT1",
        "Gosford -NCCCVT1",
        "Hamilton -NCCCVT2",
        "Hornsby -NCNCVT1",
        "Lawson -NWSCVT2",
        "Sutherland -NCSCVT2",
        "Sydenham -NCSCVT1",
        "Wollongong -NSCCVT1"
    };

    switch (Message)
    {
    case WM_INITDIALOG:
        //add the data to the drop downs
        //firstly turnout type
        t_out = GetDlgItem(hwnd, turnout_type_drop_down);
        for (int Count = 0; Count < 4; Count++)
        {
            SendMessage(t_out, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>((LPCTSTR)t_out_list[Count]));
        }
        SendMessage(t_out, CB_SETCURSEL, (WPARAM)0, (LPARAM)0);
        //and territories
        territories = GetDlgItem(hwnd, territory_drop_down);
        for (int Count = 0; Count < 12; Count++)
        {
            SendMessage(territories, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>((LPCTSTR)territory_list[Count]));
        }
        SendMessage(territories, CB_SETCURSEL, (WPARAM)0, (LPARAM)0);        
        return TRUE;
    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case search_iw29:
            open_iw29_iw65(hwnd, "iw29");
            break;
        case search_iw65:
            open_iw29_iw65(hwnd, "iw65");
            break;
        case find_sap_defect:
            open_iw29_iw65(hwnd, "iw29_srch_defect");
            break; 
        case edit_sap_defect:
            open_iw29_iw65(hwnd, "iw22_edit");
            break;
        case show_defects_for_turnout:
            open_iw29_iw65(hwnd, "view_defects_for_turnout");
            break;
        case find_turnout_btn:
            open_iw29_iw65(hwnd, "find_turnout");
            break;
        case IDOK:
            EndDialog(hwnd, IDOK);
            break;
        }
        break;
    default:
        return FALSE;
    }
    return TRUE;
}

/*

If we a create a window manually it needs its own wndproc to handle window events

*/
/*
LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
    {
        switch (LOWORD(wParam)) // *** Tells us which control ID sent a message ***
        {
        case srs2t1: // *** user clicked box 1 ***
        {
            if (SendMessage(GetDlgItem(hwnd, srs2t1), BM_GETCHECK, 0, 0) == BST_CHECKED)
                showMessageBox("1","1");
        } break;
        case srs3t3:
        {
            //if (SendMessage(GetDlgItem(hwnd, srs3t3), BM_GETCHECK, 0, 0) == BST_CHECKED)
            DialogBox(hinst,MAKEINTRESOURCE(IDD_DIALOG1), hwnd, DialogProc);
        } break;
        case srs3t5:
        {
            if (SendMessage(GetDlgItem(hwnd, srs3t5), BM_GETCHECK, 0, 0) == BST_CHECKED)
                //showMessageBox("3", "3");
                //ShowWindow(windowHandle, SW_HIDE);
                DestroyWindow(hwnd);
        } break;
        }
    }
    break;
    case WM_CHAR: //this is just for a program exit besides window's borders/task bar
       //if (wParam == VK_ESCAPE)
            //DestroyWindow(hwnd);
    case WM_CLOSE:
        DestroyWindow(hwnd);
    case WM_DESTROY:
        windowCreated = false;
        PostQuitMessage(0);
        return 0;
    default:
        return DefWindowProc(hwnd, message, wParam, lParam);
    }
    return 0;
}
*/

//this is the window that allows user to enter a defect id and search SAP
//also look up t/o, diamond etc
int WINAPI showDefectWindow()
{
    //rather than creating a window manually just use the resource dialog box
    DialogBox(hinst, MAKEINTRESOURCE(IDD_DIALOG1), 0, DialogProc);
    return 0;

    /* use the following code to create a window progmatically*/
    /*
    //check if review window open
    HWND reviewWindow = FindWindow("ThunderRT6FormDC", "A.I.M.S  3.4.6 - Main");
    if (!reviewWindow) {
        return 0;
    }

    //if the window has already been created
    //just show it and exit
    if (windowCreated) {
        ShowWindow(windowHandle, SW_SHOW);
        return 0;
    }
    //start creating the GUI window
    WNDCLASS windowClass{};

    //if the window was closed the class remains registered
    //so check if class already exists or not
    if (!GetClassInfo(hinst, "AIMS-SAP-HELPER", &windowClass))
    {
        windowClass.hbrBackground = GetSysColorBrush(COLOR_3DFACE);//(HBRUSH)GetStockObject(COLOR_WINDOW + 1);
        windowClass.hCursor = LoadCursor(NULL, IDC_ARROW);
        windowClass.hInstance = hinst;
        windowClass.lpfnWndProc = WndProc;
        windowClass.lpszClassName = "AIMS-SAP-HELPER";
        windowClass.style = CS_HREDRAW | CS_VREDRAW;
        if (!RegisterClass(&windowClass)) {
            MessageBox(NULL, "Could not register class", "Error", MB_OK);
        }
    }

    //create the main window
    windowHandle = CreateWindow("AIMS-SAP-HELPER",
        "SAP Defect/Equipment Finder",
        WS_SYSMENU | WS_CAPTION,
        0, 0, 500, 300,
        NULL, NULL, NULL, NULL);
    
    //disable close button
    EnableMenuItem(GetSystemMenu(windowHandle, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);

    //group box
    CreateWindow("BUTTON", "South Run", WS_CHILD | WS_VISIBLE | BS_GROUPBOX,5,5,480,250,windowHandle, NULL, NULL, NULL);

    //create checkbox
    HWND SouthS2T1 = CreateWindowEx(WS_EX_TRANSPARENT, "BUTTON", "Session 2 SR - Main Up 10104",
        WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
        20, 20, 460, 15,
        windowHandle, NULL, NULL, NULL);
    HWND SouthS3T3 = CreateWindowEx(WS_EX_TRANSPARENT, "BUTTON", "Session 3 SR - Illawarra Dive Down 10140",
        WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
        20, 35, 460, 15,
        windowHandle, NULL, NULL, NULL);
    HWND SouthS3T5 = CreateWindowEx(WS_EX_TRANSPARENT, "BUTTON", "Session 3 SR - Illawarra Down Main 1 10113",
        WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
        20, 50, 460, 15,
        windowHandle, NULL, NULL, NULL);

    //give the control its ID
    SetWindowLong(SouthS2T1, GWL_ID, srs2t1);
    SetWindowLong(SouthS3T3, GWL_ID, srs3t3);
    SetWindowLong(SouthS3T5, GWL_ID, srs3t5);

    //for some reason the default font is very large and ugly
    //so set the system default font with this
    EnumChildWindows(windowHandle, (WNDENUMPROC)SetFont, (LPARAM)GetStockObject(DEFAULT_GUI_FONT));

    //show window
    ShowWindow(windowHandle, SW_RESTORE);

    //set windowcreated to true
    windowCreated = true;

    MSG messages;
    while (GetMessage(&messages, NULL, 0, 0) > 0)
    {
        TranslateMessage(&messages);
        DispatchMessage(&messages);
    }
    DeleteObject(windowHandle); //doing it just in case
    return messages.wParam;
    */
}

//show the track selection window as we want it to be
void resizeTrackSelectionWindow(bool notshowBtn = true)
{
    HWND trackSelectionWindow = FindWindow("ThunderRT6FormDC", "Select Track Section");
    if (trackSelectionWindow) {

        //get current position
        RECT rect;
        GetWindowRect(trackSelectionWindow, &rect);
        int x = rect.left;
        int y = rect.top;

        // get controls on window
        HWND cancelBtn = FindWindowExA(trackSelectionWindow, NULL, "ThunderRT6CommandButton", "&Cancel");
        HWND backBtn = FindWindowExA(trackSelectionWindow, NULL, "ThunderRT6CommandButton", "<< &Back");
        HWND browseBtn = FindWindowExA(trackSelectionWindow, NULL, "ThunderRT6CommandButton", "Browse &Archive");
        HWND selectBtn = FindWindowExA(trackSelectionWindow, NULL, "ThunderRT6CommandButton", "&Select");
        HWND importExceedences = FindWindowExA(trackSelectionWindow, NULL, "ThunderRT6CommandButton", "Import Exceedances");
        HWND tabs = FindWindowExA(trackSelectionWindow, NULL, "SSTabCtlWndClass", NULL);
        // note control is child of sub-control
        HWND checkBox = FindWindowExA(tabs, NULL, "ThunderRT6CheckBox", NULL);
        HWND flexGrid = FindWindowExA(tabs, NULL, "MSFlexGridWndClass", NULL);

        // size for the tabbed area and msflexgrid
        int width = 640;
        int height = 950;

        // set size of main window
        SetWindowPos(trackSelectionWindow, 0, x, y, 1342, 1000, SWP_SHOWWINDOW);

        // hide unneeded controls
        ShowWindow(cancelBtn, 0);
        ShowWindow(checkBox, 0);
        ShowWindow(backBtn, 0);
        ShowWindow(browseBtn, 0);
        if (notshowBtn)
        {
            ShowWindow(selectBtn, 0);
        }

        // resize controls
        SetWindowPos(tabs, 0, 680, 10, width, height, SWP_SHOWWINDOW);
        SetWindowPos(flexGrid, 0, 0, 0, width, height, SWP_SHOWWINDOW);
        //SetWindowPos(selectBtn, 0, 4, 323, 657, 33, SWP_SHOWWINDOW);

        //reload the image everytime
        if (FindWindowExA(trackSelectionWindow, NULL, "STATIC", NULL)) {
            //load the image 
            LoadScreen(FindWindowExA(trackSelectionWindow, NULL, "STATIC", NULL));
        }

        //if custom buttons not found create them
        if (!FindWindowExA(trackSelectionWindow, NULL, "BUTTON", "Sort by Track ID")) {
            //create the custom buttons
            CreateWindowEx(WS_EX_CONTROLPARENT, "BUTTON", "Sort by Track ID",
                WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                4, 360, 326, 33,
                trackSelectionWindow, NULL, NULL, NULL);
            CreateWindowEx(WS_EX_CONTROLPARENT, "BUTTON", "Sort by Sessions",
                WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                330, 360, 331, 33,
                trackSelectionWindow, NULL, NULL, NULL);
            CreateWindowEx(WS_EX_CONTROLPARENT, "BUTTON", "Check Processed Tracks in Excel",
                WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                4, 395, 326, 33,
                trackSelectionWindow, NULL, NULL, NULL);
            CreateWindowEx(WS_EX_CONTROLPARENT, "BUTTON", "Sort by Track Name",
                WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                330, 395, 331, 33,
                trackSelectionWindow, NULL, NULL, NULL);
            HWND staticControl = CreateWindowEx(WS_EX_DLGMODALFRAME, "STATIC", NULL,
                WS_VISIBLE | WS_CHILD | SS_BITMAP,
                73, 476, 504, 412,
                trackSelectionWindow, NULL, NULL, NULL);
            //load the image 
            LoadScreen(staticControl);
        }

        //update display
        UpdateWindow(trackSelectionWindow);
    }
}

//get processed tracks method
void getProcessedTracks(HWND targetWnd, int sort = 0) {
    //class name placeholder
    char szClassName[128];
    memset(szClassName, 0, sizeof(szClassName));

    //child of main window
    HWND tabs = FindWindowExA(targetWnd, NULL, "SSTabCtlWndClass", NULL);
    //child of control
    HWND flexGridHWND = FindWindowExA(tabs, NULL, "MSFlexGridWndClass", NULL);

    //if the flexgrid is not found - exit
    if (!tabs && !flexGridHWND) {
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

        /*none of the following work!
        //set the font size
        static BYTE parmst[] = VTS_R4;
        iDriver.InvokeHelper(0x4e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parmst, 200);
        
        //set font width
        static BYTE parmsC[] = VTS_R4;
        iDriver.InvokeHelper(0x54, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parmsC, 200);

        //refresh
        iDriver.InvokeHelper(DISPID_REFRESH, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);

        //get data source
        LPDISPATCH pDispatch;
        iDriver.InvokeHelper(0x4c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&pDispatch, NULL);
                
        //output to csv file
        std::ofstream Toutfile;
        Toutfile.open("C:\\Users\\Public\\Documents\\xtra.txt", std::ios_base::app); // append
        Toutfile << pDispatch << "\n";
        Toutfile.close();
        //END OF NOT WORKING SECTION
        */

        //GetMouseCol
        //used this to get col no.
        //long mouseCol;
        //iDriver.InvokeHelper(0x1c, DISPATCH_PROPERTYGET, VT_I4, (void*)&mouseCol, NULL);
        
        //std::string test = std::to_string(mouseCol);

        //Section = 0, Session = 1, Capture Date = 2, Track ID = 3 ... etc
        //col no. 3 is the track ID col
        //showMessageBox(test.c_str(), "col");
        
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
        ShellExecute(NULL, NULL, "C:\\Users\\Public\\Documents\\Get_Processed_Tracks.xlsm", NULL, "C:\\Users\\Public\\Documents", SW_SHOWDEFAULT);
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
        //local declarations
        char szClassName[128];
        memset(szClassName, 0, sizeof(szClassName));

        //get class of item the user clicked on
        ::GetClassName(pMsg->hwnd, szClassName, sizeof(szClassName) - 1);

        //check if it is msflexgrid control
        if (IsWindow(pMsg->hwnd) && (strcmp(szClassName, "MSFlexGridWndClass") == 0)) {
            HWND trackSelectionWindow = FindWindow("ThunderRT6FormDC", "Select Track Section");
            if (trackSelectionWindow) {
                //get select button
                HWND selectBtn = FindWindowExA(trackSelectionWindow, NULL, "ThunderRT6CommandButton", "&Select");
                //move and show button
                if (selectBtn) {
                    SetWindowPos(selectBtn, 0, 4, 323, 657, 33, SWP_SHOWWINDOW);
                }
            }
        }
        else 
        {
            //check if the user double clicked on the defect ID in EAM defect window
            //or if they want to enable disabled text editing in POI window
            controlProperties();
        }

        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }

    //someone left click?
    if (pMsg->message == WM_LBUTTONUP)
    {
        //check if select track window open
        HWND targetWnd = FindWindow("ThunderRT6FormDC", "Select Track Section");
        if (targetWnd && IsWindowVisible(targetWnd)) {
            //resizeTrackSelectionWindow();
            //local declarations
            char szClassName[128];
            memset(szClassName, 0, sizeof(szClassName));
            char szText[128];
            memset(szText, 0, sizeof(szText));

            //get class of item the user clicked on
            ::GetClassName(pMsg->hwnd, szClassName, sizeof(szClassName) - 1);

            //check if it is a button
            if (IsWindow(pMsg->hwnd) && (strcmp(szClassName, "Button") == 0)) {
                GetWindowText(pMsg->hwnd, szText, sizeof(szText) - 1);
                //check if it is the button we want
                if (strcmp(szText, "Sort by Track ID") == 0) {
                    //sort the tracks selection grid by track id
                    getProcessedTracks(targetWnd, 3);
                }
                if (strcmp(szText, "Sort by Sessions") == 0) {
                    //sort the tracks selection grid by SESSION number
                    getProcessedTracks(targetWnd, 1);
                }
                if (strcmp(szText, "Check Processed Tracks in Excel") == 0) {
                    //export track list to csv and open excel file
                    getProcessedTracks(targetWnd);
                }                
                if (strcmp(szText, "Sort by Track Name") == 0) {
                    //sort by track name
                    getProcessedTracks(targetWnd, 4);
                }
            }
        }
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
                //when review window is hidden and if we press escape key it will crash AIMS
                if (!IsWindowVisible(aimsAPP)) {
                    return 0;
                }
                //Select View->POI Markers
                HMENU fileMenu = findMenu(GetMenu(aimsAPP), (LPARAM)TEXT("View"));
                selectMenuItem(aimsAPP, fileMenu, (LPARAM)TEXT("POI Markers"));
                Sleep(100);
                // Select View->Refresh
                HMENU refreshMenu = findMenu(GetMenu(aimsAPP), (LPARAM)TEXT("View"));
                selectMenuItem(aimsAPP, refreshMenu, (LPARAM)TEXT("Refresh"));

            }

            return CallNextHookEx(NULL, nCode, wParam, lParam);
        }

        //show about aims window and add our custom defect lookup stuff
        if (pMsg->wParam == VK_F9)
        {
            // are v reviewing?
            HWND aimsAPP = FindWindow("ThunderRT6FormDC", "A.I.M.S  3.4.6 - Main");
            if (aimsAPP)
            {
                showDefectWindow();
            }

            return CallNextHookEx(NULL, nCode, wParam, lParam);
        }

        //get tracks processed from track selection window
        // VK_OEM_3 = ` and ~
        /*
        if (pMsg->wParam == VK_OEM_3)
        {
            //if no track selection window open - exit
            HWND targetWnd = FindWindow("ThunderRT6FormDC", "Select Track Section");
            if (!targetWnd) {
                return 0;
            }
            //call the get processed tracks method to export track list to csv
            getProcessedTracks(targetWnd);

            return CallNextHookEx(NULL, nCode, wParam, lParam);
        }
        //pressed the "s" key
        if (pMsg->wParam == 0x53)
        {
            //if no track selection window open - exit
            HWND targetWnd = FindWindow("ThunderRT6FormDC", "Select Track Section");
            if (!targetWnd) {
                return 0;
            }
            //sort the tracks selection grid by track id
            getProcessedTracks(targetWnd, true);

            return CallNextHookEx(NULL, nCode, wParam, lParam);
        }
        */
        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

//check if resize window function has run
bool windowResized = false;

//custom WNDPROC NOT WORKING 100% properly
/*
LRESULT CALLBACK NuWndProc(int nCode, WPARAM wParam, LPARAM lParam) {
    //message param
    PMSG pMsg = (PMSG)lParam;

    if (pMsg->message == BN_CLICKED)
    {
        if (pMsg->wParam == srs2t1)
        {
            showMessageBox("1", "1");
        }
    }

    if (nCode == HC_ACTION) {
        //lets extract the data
        CWPSTRUCT* data = (CWPSTRUCT*)lParam;
        if (data->message == WM_CLOSE) {
            //lets get the name of the program closed
            char name[260];
            GetWindowModuleFileNameA(data->hwnd, name, 260);
            //extract only the exe from the path
            char* extract = (char*)((DWORD)name + lstrlenA(name) - 1);
            while (*extract != '\\')
                extract--;
            extract++;
            MessageBoxA(0, "A program has been closed", extract, 0);
        }
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}
*/


//custom CBTProc
LRESULT CALLBACK CBTProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode < 0)
    {
        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }
        
    //freopen("output.txt", "a", stdout);
    //cout << "x: " << x << " y: " << y << endl;


    //when we the track selection window is ready 
    //the cancel button is the last to be created 
    //so check for it and its parent must be "select track section" window
    if (nCode == HCBT_SETFOCUS)
    {
        HWND hwnd = (HWND)wParam;
        if (!HWNDToString(hwnd).empty()) {
            std::string parentWindow = HWNDToString(GetRealParent(hwnd));
            std::string control = HWNDToString(hwnd);
            //WriteToFile(control + parentWindow);

            //check for it and resize the window to the way we want it to be
            if (control.compare("&Cancel") == 0 && parentWindow.compare("Select Track Section") == 0) {
                resizeTrackSelectionWindow();
                windowResized = true;
                return CallNextHookEx(NULL, nCode, wParam, lParam);
            }
            return CallNextHookEx(NULL, nCode, wParam, lParam);
        }

        //for some reason the resize of window works but
        //it has to be done again otherwise the "select" button is not
        //positioned correctly
        if (windowResized) {
            resizeTrackSelectionWindow(false);
        }
        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }

    if (nCode == HCBT_DESTROYWND)
    {
        /*
        HWND hTemp = (HWND)wParam;
        DWORD id;

        if (hTemp != NULL)
        {
            //MessageBox(GetDesktopWindow(), "hTemp e NULL", "a", MB_OK);
            id = GetWindowThreadProcessId(hTemp, NULL);
            if (id != NULL)
            {
                auto sid = std::to_string(id);
                //MessageBox(GetDesktopWindow(), "idu e 0", "a", MB_OK);
                WriteToFile(sid);
            }
        }
        */

        /*
        if (!HWNDToString(hwnd).empty()) {
            std::string parentWindow = " parent: " + HWNDToString(GetRealParent(hwnd));
            std::string control = "child: " + HWNDToString(hwnd);
            WriteToFile(control + parentWindow);
            if (control.compare("AIMS_Office") == 0) {
                if (windowCreated) {
                    showMessageBox("closing", "s");
                    DestroyWindow(windowHandle);
                }
            }
        }
        */
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
    //WndProc_hHook = SetWindowsHookEx(WH_CALLWNDPROC, (HOOKPROC)NuWndProc, hinst, threadID);
    g_hHook = SetWindowsHookEx(WH_GETMESSAGE, (HOOKPROC)GetMsgProc, hinst, threadID);
    cbt_hHook = SetWindowsHookEx(WH_CBT, (HOOKPROC)CBTProc, hinst, threadID);
}

//called when the injector app is closed
extern "C" __declspec(dllexport) void uninstall() {
    UnhookWindowsHookEx(g_hHook);
    UnhookWindowsHookEx(cbt_hHook);
    UnhookWindowsHookEx(WndProc_hHook);
}