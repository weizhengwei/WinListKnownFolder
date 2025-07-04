#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#ifndef UNICODE
#define UNICODE
#endif

#ifndef _UNICODE
#define _UNICODE
#endif

#include <Windows.h>
#include <windowsx.h>
#include <tchar.h>
#include <ShlObj.h>
#include <string>
#include <vector>
#include <map>

namespace std {
	using tstring = std::basic_string<TCHAR>;
}

#ifdef _MSC_VER
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")
#endif

const TCHAR* CLASSNAME = _T("CLSNAME_KNOWNFOLDER");
const TCHAR* WIN_TITLE = _T("WinListKnownFolder");
const TCHAR* VER = _T("V1.0.1");

std::vector<std::tstring> vecCols{
    _T("FOLDERID"),
    _T("GUID"),
    _T("Path"),
    _T("LocalizedModule"),
    _T("ResourceID"),
    _T("LocalizedName")
};

std::map<int, int> vecColWidth = {
    {0, 250},
    {1, 300},
    {2, 650},
    {3, 300},
    {4, 100},
    {5, 150}
};

#define ID_LISTVIEW 100
HWND hListView;

#define IDM_FILE_REFRESH 1001
#define IDM_FILE_EXIT 1002
#define IDM_HELP_ABOUT 2001

#define IDM_CTX_OPEN 101
#define IDM_CTX_COPY 102

LRESULT CALLBACK WinProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
void QuitApp();
void CreateListView(HWND hwnd);
void CreateMainMenu(HWND hwnd);
void FillPathInfoToListView();
std::tstring GetFolderGUIDStr(GUID folderGuid);
void CopySelectItemInfo(HWND hwnd);
void CopyTextToClipboard(HWND hwnd, const std::tstring& text);

std::vector<GUID> vecFolderGUID
{
    FOLDERID_NetworkFolder,
    FOLDERID_ComputerFolder,
    FOLDERID_InternetFolder,
    FOLDERID_ControlPanelFolder,
    FOLDERID_PrintersFolder,
    FOLDERID_SyncManagerFolder,
    FOLDERID_SyncSetupFolder,
    FOLDERID_ConflictFolder,
    FOLDERID_SyncResultsFolder,
    FOLDERID_RecycleBinFolder,
    FOLDERID_ConnectionsFolder,
    FOLDERID_Fonts,
    FOLDERID_Desktop,
    FOLDERID_Startup,
    FOLDERID_Programs,
    FOLDERID_StartMenu,
    FOLDERID_Recent,
    FOLDERID_SendTo,
    FOLDERID_Documents,
    FOLDERID_Favorites,
    FOLDERID_NetHood,
    FOLDERID_PrintHood,
    FOLDERID_Templates,
    FOLDERID_CommonStartup,
    FOLDERID_CommonPrograms,
    FOLDERID_CommonStartMenu,
    FOLDERID_PublicDesktop,
    FOLDERID_ProgramData,
    FOLDERID_CommonTemplates,
    FOLDERID_PublicDocuments,
    FOLDERID_RoamingAppData,
    FOLDERID_LocalAppData,
    FOLDERID_LocalAppDataLow,
    FOLDERID_InternetCache,
    FOLDERID_Cookies,
    FOLDERID_History,
    FOLDERID_System,
    FOLDERID_SystemX86,
    FOLDERID_Windows,
    FOLDERID_Profile,
    FOLDERID_Pictures,
    FOLDERID_ProgramFilesX86,
    FOLDERID_ProgramFilesCommonX86,
    FOLDERID_ProgramFilesX64,
    FOLDERID_ProgramFilesCommonX64,
    FOLDERID_ProgramFiles,
    FOLDERID_ProgramFilesCommon,
    FOLDERID_UserProgramFiles,
    FOLDERID_UserProgramFilesCommon,
    FOLDERID_AdminTools,
    FOLDERID_CommonAdminTools,
    FOLDERID_Music,
    FOLDERID_Videos,
    FOLDERID_Ringtones,
    FOLDERID_PublicPictures,
    FOLDERID_PublicMusic,
    FOLDERID_PublicVideos,
    FOLDERID_PublicRingtones,
    FOLDERID_ResourceDir,
    FOLDERID_LocalizedResourcesDir,
    FOLDERID_CommonOEMLinks,
    FOLDERID_CDBurning,
    FOLDERID_UserProfiles,
    FOLDERID_Playlists,
    FOLDERID_SamplePlaylists,
    FOLDERID_SampleMusic,
    FOLDERID_SamplePictures,
    FOLDERID_SampleVideos,
    FOLDERID_PhotoAlbums,
    FOLDERID_Public,
    FOLDERID_ChangeRemovePrograms,
    FOLDERID_AppUpdates,
    FOLDERID_AddNewPrograms,
    FOLDERID_Downloads,
    FOLDERID_PublicDownloads,
    FOLDERID_SavedSearches,
    FOLDERID_QuickLaunch,
    FOLDERID_Contacts,
    FOLDERID_SidebarParts,
    FOLDERID_SidebarDefaultParts,
    FOLDERID_PublicGameTasks,
    FOLDERID_GameTasks,
    FOLDERID_SavedGames,
    FOLDERID_Games,
    FOLDERID_SEARCH_MAPI,
    FOLDERID_SEARCH_CSC,
    FOLDERID_Links,
    FOLDERID_UsersFiles,
    FOLDERID_UsersLibraries,
    FOLDERID_SearchHome,
    FOLDERID_OriginalImages,
    FOLDERID_DocumentsLibrary,
    FOLDERID_MusicLibrary,
    FOLDERID_PicturesLibrary,
    FOLDERID_VideosLibrary,
    FOLDERID_RecordedTVLibrary,
    FOLDERID_HomeGroup,
    FOLDERID_HomeGroupCurrentUser,
    FOLDERID_DeviceMetadataStore,
    FOLDERID_Libraries,
    FOLDERID_PublicLibraries,
    FOLDERID_UserPinned,
    FOLDERID_ImplicitAppShortcuts,
    FOLDERID_AccountPictures,
    FOLDERID_PublicUserTiles,
    FOLDERID_AppsFolder,
    FOLDERID_StartMenuAllPrograms,
    FOLDERID_CommonStartMenuPlaces,
    FOLDERID_ApplicationShortcuts,
    FOLDERID_RoamingTiles,
    FOLDERID_RoamedTileImages,
    FOLDERID_Screenshots,
    FOLDERID_CameraRoll,
    FOLDERID_SkyDrive,
    FOLDERID_OneDrive,
    FOLDERID_SkyDriveDocuments,
    FOLDERID_SkyDrivePictures,
    FOLDERID_SkyDriveMusic,
    FOLDERID_SkyDriveCameraRoll,
    FOLDERID_SearchHistory,
    FOLDERID_SearchTemplates,
    FOLDERID_CameraRollLibrary,
    FOLDERID_SavedPictures,
    FOLDERID_SavedPicturesLibrary,
    FOLDERID_RetailDemo,
    FOLDERID_Device,
    FOLDERID_DevelopmentFiles,
    FOLDERID_Objects3D,
    FOLDERID_AppCaptures,
    FOLDERID_LocalDocuments,
    FOLDERID_LocalPictures,
    FOLDERID_LocalVideos,
    FOLDERID_LocalMusic,
    FOLDERID_LocalDownloads,
    FOLDERID_RecordedCalls,
    FOLDERID_AllAppMods,
    FOLDERID_CurrentAppMods,
    FOLDERID_AppDataDesktop,
    FOLDERID_AppDataDocuments,
    FOLDERID_AppDataFavorites,
    FOLDERID_AppDataProgramData,
    FOLDERID_LocalStorage
};

int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE hPreInstance, LPTSTR lpCmdLine, int nShowCmd)
{
	WNDCLASS wc;
	ZeroMemory(&wc, sizeof(wc));
	wc.style = CS_HREDRAW | CS_VREDRAW;
	wc.lpfnWndProc = WinProc;
	wc.hInstance = hInstance;
	wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wc.lpszClassName = CLASSNAME;

	if (RegisterClass(&wc) == 0)
	{
		return 1;
	}

	HWND hwnd = CreateWindowEx(WS_EX_APPWINDOW,
		CLASSNAME,
		WIN_TITLE,
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		NULL,
		NULL,
		hInstance,
		NULL);
	CreateMainMenu(hwnd);
	CreateListView(hwnd);
	ShowWindow(hwnd, SW_SHOWDEFAULT);
	UpdateWindow(hwnd);

	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}

LRESULT CALLBACK WinProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
    case WM_COMMAND:
    {
        if (LOWORD(wParam) == IDM_FILE_REFRESH)
        {
            ListView_DeleteAllItems(hListView);
            UpdateWindow(hListView);
            FillPathInfoToListView();
        }
        else if (LOWORD(wParam) == IDM_FILE_EXIT)
        {
            QuitApp();
        }
        else if (LOWORD(wParam) == IDM_HELP_ABOUT)
        {
            TCHAR szTitle[256] = { 0 };
            _stprintf_s(szTitle, ARRAYSIZE(szTitle), _T("%s %s\n\nCopyright(C) 2025 hostzhengwei@gmail.com"), WIN_TITLE, VER);
            MessageBox(hwnd, szTitle, WIN_TITLE, MB_OK);
        }
        else if (LOWORD(wParam) == IDM_CTX_OPEN)
        {
            UINT itemid = -1;
            itemid = ListView_GetNextItem(hListView, itemid, LVNI_SELECTED);
            if (itemid != -1)
            {
                TCHAR szText[32] = {};
                ListView_GetItemText(hListView, itemid, 2, szText, ARRAYSIZE(szText));
                SHELLEXECUTEINFO ShExecInfo;
                ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
                ShExecInfo.fMask = NULL;
                ShExecInfo.hwnd = NULL;
                ShExecInfo.lpVerb = NULL;
                ShExecInfo.lpFile = _T("explorer");
                ShExecInfo.lpParameters = szText;
                ShExecInfo.lpDirectory = NULL;
                ShExecInfo.nShow = SW_SHOWDEFAULT;
                ShExecInfo.hInstApp = NULL;

                ShellExecuteEx(&ShExecInfo);
            }
        }
        else if (LOWORD(wParam) == IDM_CTX_COPY)
        {
            CopySelectItemInfo(hwnd);
        }
        break;
    }
    case WM_SIZE:
    {
        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        InflateRect(&rcClient, -15, -15);
        SetWindowPos(hListView,
            NULL,
            rcClient.left,
            rcClient.top,
            rcClient.right - rcClient.left,
            rcClient.bottom - rcClient.top,
            SWP_NOACTIVATE | SWP_NOZORDER
        );
        break;
    }
    case WM_CONTEXTMENU:
    {
        POINT pt;
        pt.x = GET_X_LPARAM(lParam);
        pt.y = GET_Y_LPARAM(lParam);
        ScreenToClient(hListView, &pt);
        LVHITTESTINFO hitTestInfo;
        hitTestInfo.pt = pt;
        ListView_HitTest(hListView, &hitTestInfo);
        if (hitTestInfo.iItem != -1) {
            UINT uCount = ListView_GetSelectedCount(hListView);
            LANGID langID = GetUserDefaultUILanguage();
            if (langID != 0x0804 && langID != 0x0409)
            {
                langID = 0x0409;
            }
            HMENU hMenu = CreatePopupMenu();
            if (uCount == 1)
            {
                AppendMenu(hMenu, MF_STRING, IDM_CTX_OPEN, _T("Open"));
                AppendMenu(hMenu, MF_SEPARATOR, 0, L"");
                AppendMenu(hMenu, MF_STRING, IDM_CTX_COPY, _T("Copy"));
            }
            else
            {
                AppendMenu(hMenu, MF_STRING, IDM_CTX_COPY, _T("Copy"));
            }


            // Determine the menu position
            pt.x = GET_X_LPARAM(lParam);
            pt.y = GET_Y_LPARAM(lParam);
            UINT uFlags = TPM_RIGHTBUTTON; // Use this flag to handle right clicks
            TrackPopupMenuEx(hMenu, uFlags, pt.x, pt.y, hwnd, NULL);
            DestroyMenu(hMenu);
        }
        break;
    }
	case WM_DESTROY:
	{
		QuitApp();
		break;
	}
	default:
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}
	return 0;
}

void QuitApp()
{
	PostQuitMessage(0);
}

void CreateListView(HWND hwnd)
{
    RECT rc;
    GetClientRect(hwnd, &rc);
    InflateRect(&rc, -10, -10);
    hListView = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        WC_LISTVIEW,
        NULL,
        WS_VISIBLE | WS_CHILD | WS_CLIPCHILDREN | LVS_REPORT | LVS_NOSORTHEADER,
        rc.left,
        rc.top,
        rc.right - rc.left,
        rc.bottom - rc.top,
        hwnd,
        (HMENU)ID_LISTVIEW,
        GetModuleHandle(NULL),
        NULL);
    ListView_SetTextColor(hListView, RGB(10, 10, 160));
    ListView_SetExtendedListViewStyle(hListView, LVS_EX_FULLROWSELECT);

    TCHAR szText[MAX_PATH] = {};
    LVCOLUMN lvCol;
    lvCol.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvCol.fmt = LVCFMT_LEFT;
    lvCol.cxDefault = 150;
    lvCol.pszText = szText;

    for (int iCol = 0; iCol < vecCols.size(); iCol++)
    {
        lvCol.cx = vecColWidth.at(iCol);
        _stprintf_s(szText, MAX_PATH, _T("%s"), vecCols[iCol].c_str());
        ListView_InsertColumn(hListView, iCol, &lvCol);
    }

    FillPathInfoToListView();
}

void FillPathInfoToListView()
{
    int iCount = 0;
    for (auto item : vecFolderGUID)
    {
        PWSTR lpPath = NULL;
        HRESULT hr = SHGetKnownFolderPath(item, KF_FLAG_DEFAULT, NULL, &lpPath);

        
        WCHAR szGUID[MAX_PATH] = {};
        StringFromGUID2(item, szGUID, MAX_PATH);

        LVITEM lvItem = {};
        TCHAR szText[128] = {};

        lvItem.iItem = iCount++;

        lvItem.iSubItem = 0;
        _stprintf_s(szText, ARRAYSIZE(szText), _T("%s"), GetFolderGUIDStr(item).c_str());
        lvItem.mask = LVIF_TEXT;
        lvItem.pszText = szText;
        ListView_InsertItem(hListView, &lvItem);

        lvItem.iSubItem = 1;
        _stprintf_s(szText, ARRAYSIZE(szText), _T("%s"), szGUID);
        lvItem.mask = LVIF_TEXT;
        lvItem.pszText = szText;
        ListView_SetItem(hListView, &lvItem);

        if (SUCCEEDED(hr))
        {
            lvItem.iSubItem = 2;
            _stprintf_s(szText, ARRAYSIZE(szText), _T("%s"), lpPath);
            lvItem.mask = LVIF_TEXT;
            lvItem.pszText = szText;
            ListView_SetItem(hListView, &lvItem);
        }
        else
        {
            lvItem.iSubItem = 2;
            _stprintf_s(szText, ARRAYSIZE(szText), _T("%s"), L"");
            lvItem.mask = LVIF_TEXT;
            lvItem.pszText = szText;
            ListView_SetItem(hListView, &lvItem);
        }
        //---
        WCHAR szResMod[MAX_PATH] = {};
        int idsRes = 0;
        hr = SHGetLocalizedName(lpPath, szResMod, MAX_PATH, &idsRes);
        if (SUCCEEDED(hr))
        {
            TCHAR szFullName[MAX_PATH] = {};
            ExpandEnvironmentStrings(szResMod, szFullName, MAX_PATH);
            lvItem.iSubItem = 3;
            _stprintf_s(szText, ARRAYSIZE(szText), _T("%s"), szFullName);
            lvItem.mask = LVIF_TEXT;
            lvItem.pszText = szText;
            ListView_SetItem(hListView, &lvItem);

            lvItem.iSubItem = 4;
            _stprintf_s(szText, ARRAYSIZE(szText), _T("%d"), idsRes);
            lvItem.mask = LVIF_TEXT;
            lvItem.pszText = szText;
            ListView_SetItem(hListView, &lvItem);

            HMODULE hMod = GetModuleHandle(szFullName);
            TCHAR szResName[MAX_PATH] = {};
            int res = 0;
            if (hMod != NULL)
            {
                res = LoadString(hMod, idsRes, szResName, MAX_PATH);
            }
            else
            {
                hMod = LoadLibrary(szFullName);
                res = LoadString(hMod, idsRes, szResName, MAX_PATH);
            }
            if (res > 0)
            {
                lvItem.iSubItem = 5;
                _stprintf_s(szText, ARRAYSIZE(szText), _T("%s"), szResName);
                lvItem.mask = LVIF_TEXT;
                lvItem.pszText = szText;
                ListView_SetItem(hListView, &lvItem);

                //OutputDebugString(lpPath);
                //OutputDebugString(L"-----");
                //OutputDebugString(szResName);
                //OutputDebugString(L"\n");
            }
        }
        //---

        CoTaskMemFree(lpPath);
    }
}

void CreateMainMenu(HWND hwnd)
{
    HMENU hMainMenu = CreateMenu();
    HMENU hFileSubMenu = CreatePopupMenu();
    AppendMenu(hFileSubMenu, MF_STRING, IDM_FILE_REFRESH, _T("Refresh"));
    AppendMenu(hFileSubMenu, MF_SEPARATOR, 0, _T(""));
    AppendMenu(hFileSubMenu, MF_STRING, IDM_FILE_EXIT, _T("Exit"));

    HMENU hHelpSubMenu = CreatePopupMenu();
    AppendMenu(hHelpSubMenu, MF_STRING, IDM_HELP_ABOUT, _T("About"));

    AppendMenu(hMainMenu, MF_POPUP, (UINT_PTR)hFileSubMenu, _T("File"));
    AppendMenu(hMainMenu, MF_POPUP, (UINT_PTR)hHelpSubMenu, _T("Help"));
    SetMenu(hwnd, hMainMenu);
}

std::tstring GetFolderGUIDStr(GUID folderGuid) 
{
    if (folderGuid == FOLDERID_NetworkFolder) return _T("FOLDERID_NetworkFolder");
    if (folderGuid == FOLDERID_ComputerFolder) return _T("FOLDERID_ComputerFolder");
    if (folderGuid == FOLDERID_InternetFolder) return _T("FOLDERID_InternetFolder");
    if (folderGuid == FOLDERID_ControlPanelFolder) return _T("FOLDERID_ControlPanelFolder");
    if (folderGuid == FOLDERID_PrintersFolder) return _T("FOLDERID_PrintersFolder");
    if (folderGuid == FOLDERID_SyncManagerFolder) return _T("FOLDERID_SyncManagerFolder");
    if (folderGuid == FOLDERID_SyncSetupFolder) return _T("FOLDERID_SyncSetupFolder");
    if (folderGuid == FOLDERID_ConflictFolder) return _T("FOLDERID_ConflictFolder");
    if (folderGuid == FOLDERID_SyncResultsFolder) return _T("FOLDERID_SyncResultsFolder");
    if (folderGuid == FOLDERID_RecycleBinFolder) return _T("FOLDERID_RecycleBinFolder");
    if (folderGuid == FOLDERID_ConnectionsFolder) return _T("FOLDERID_ConnectionsFolder");
    if (folderGuid == FOLDERID_Fonts) return _T("FOLDERID_Fonts");
    if (folderGuid == FOLDERID_Desktop) return _T("FOLDERID_Desktop");
    if (folderGuid == FOLDERID_Startup) return _T("FOLDERID_Startup");
    if (folderGuid == FOLDERID_Programs) return _T("FOLDERID_Programs");
    if (folderGuid == FOLDERID_StartMenu) return _T("FOLDERID_StartMenu");
    if (folderGuid == FOLDERID_Recent) return _T("FOLDERID_Recent");
    if (folderGuid == FOLDERID_SendTo) return _T("FOLDERID_SendTo");
    if (folderGuid == FOLDERID_Documents) return _T("FOLDERID_Documents");
    if (folderGuid == FOLDERID_Favorites) return _T("FOLDERID_Favorites");
    if (folderGuid == FOLDERID_NetHood) return _T("FOLDERID_NetHood");
    if (folderGuid == FOLDERID_PrintHood) return _T("FOLDERID_PrintHood");
    if (folderGuid == FOLDERID_Templates) return _T("FOLDERID_Templates");
    if (folderGuid == FOLDERID_CommonStartup) return _T("FOLDERID_CommonStartup");
    if (folderGuid == FOLDERID_CommonPrograms) return _T("FOLDERID_CommonPrograms");
    if (folderGuid == FOLDERID_CommonStartMenu) return _T("FOLDERID_CommonStartMenu");
    if (folderGuid == FOLDERID_PublicDesktop) return _T("FOLDERID_PublicDesktop");
    if (folderGuid == FOLDERID_ProgramData) return _T("FOLDERID_ProgramData");
    if (folderGuid == FOLDERID_CommonTemplates) return _T("FOLDERID_CommonTemplates");
    if (folderGuid == FOLDERID_PublicDocuments) return _T("FOLDERID_PublicDocuments");
    if (folderGuid == FOLDERID_RoamingAppData) return _T("FOLDERID_RoamingAppData");
    if (folderGuid == FOLDERID_LocalAppData) return _T("FOLDERID_LocalAppData");
    if (folderGuid == FOLDERID_LocalAppDataLow) return _T("FOLDERID_LocalAppDataLow");
    if (folderGuid == FOLDERID_InternetCache) return _T("FOLDERID_InternetCache");
    if (folderGuid == FOLDERID_Cookies) return _T("FOLDERID_Cookies");
    if (folderGuid == FOLDERID_History) return _T("FOLDERID_History");
    if (folderGuid == FOLDERID_System) return _T("FOLDERID_System");
    if (folderGuid == FOLDERID_SystemX86) return _T("FOLDERID_SystemX86");
    if (folderGuid == FOLDERID_Windows) return _T("FOLDERID_Windows");
    if (folderGuid == FOLDERID_Profile) return _T("FOLDERID_Profile");
    if (folderGuid == FOLDERID_Pictures) return _T("FOLDERID_Pictures");
    if (folderGuid == FOLDERID_ProgramFilesX86) return _T("FOLDERID_ProgramFilesX86");
    if (folderGuid == FOLDERID_ProgramFilesCommonX86) return _T("FOLDERID_ProgramFilesCommonX86");
    if (folderGuid == FOLDERID_ProgramFilesX64) return _T("FOLDERID_ProgramFilesX64");
    if (folderGuid == FOLDERID_ProgramFilesCommonX64) return _T("FOLDERID_ProgramFilesCommonX64");
    if (folderGuid == FOLDERID_ProgramFiles) return _T("FOLDERID_ProgramFiles");
    if (folderGuid == FOLDERID_ProgramFilesCommon) return _T("FOLDERID_ProgramFilesCommon");
    if (folderGuid == FOLDERID_UserProgramFiles) return _T("FOLDERID_UserProgramFiles");
    if (folderGuid == FOLDERID_UserProgramFilesCommon) return _T("FOLDERID_UserProgramFilesCommon");
    if (folderGuid == FOLDERID_AdminTools) return _T("FOLDERID_AdminTools");
    if (folderGuid == FOLDERID_CommonAdminTools) return _T("FOLDERID_CommonAdminTools");
    if (folderGuid == FOLDERID_Music) return _T("FOLDERID_Music");
    if (folderGuid == FOLDERID_Videos) return _T("FOLDERID_Videos");
    if (folderGuid == FOLDERID_Ringtones) return _T("FOLDERID_Ringtones");
    if (folderGuid == FOLDERID_PublicPictures) return _T("FOLDERID_PublicPictures");
    if (folderGuid == FOLDERID_PublicMusic) return _T("FOLDERID_PublicMusic");
    if (folderGuid == FOLDERID_PublicVideos) return _T("FOLDERID_PublicVideos");
    if (folderGuid == FOLDERID_PublicRingtones) return _T("FOLDERID_PublicRingtones");
    if (folderGuid == FOLDERID_ResourceDir) return _T("FOLDERID_ResourceDir");
    if (folderGuid == FOLDERID_LocalizedResourcesDir) return _T("FOLDERID_LocalizedResourcesDir");
    if (folderGuid == FOLDERID_CommonOEMLinks) return _T("FOLDERID_CommonOEMLinks");
    if (folderGuid == FOLDERID_CDBurning) return _T("FOLDERID_CDBurning");
    if (folderGuid == FOLDERID_UserProfiles) return _T("FOLDERID_UserProfiles");
    if (folderGuid == FOLDERID_Playlists) return _T("FOLDERID_Playlists");
    if (folderGuid == FOLDERID_SamplePlaylists) return _T("FOLDERID_SamplePlaylists");
    if (folderGuid == FOLDERID_SampleMusic) return _T("FOLDERID_SampleMusic");
    if (folderGuid == FOLDERID_SamplePictures) return _T("FOLDERID_SamplePictures");
    if (folderGuid == FOLDERID_SampleVideos) return _T("FOLDERID_SampleVideos");
    if (folderGuid == FOLDERID_PhotoAlbums) return _T("FOLDERID_PhotoAlbums");
    if (folderGuid == FOLDERID_Public) return _T("FOLDERID_Public");
    if (folderGuid == FOLDERID_ChangeRemovePrograms) return _T("FOLDERID_ChangeRemovePrograms");
    if (folderGuid == FOLDERID_AppUpdates) return _T("FOLDERID_AppUpdates");
    if (folderGuid == FOLDERID_AddNewPrograms) return _T("FOLDERID_AddNewPrograms");
    if (folderGuid == FOLDERID_Downloads) return _T("FOLDERID_Downloads");
    if (folderGuid == FOLDERID_PublicDownloads) return _T("FOLDERID_PublicDownloads");
    if (folderGuid == FOLDERID_SavedSearches) return _T("FOLDERID_SavedSearches");
    if (folderGuid == FOLDERID_QuickLaunch) return _T("FOLDERID_QuickLaunch");
    if (folderGuid == FOLDERID_Contacts) return _T("FOLDERID_Contacts");
    if (folderGuid == FOLDERID_SidebarParts) return _T("FOLDERID_SidebarParts");
    if (folderGuid == FOLDERID_SidebarDefaultParts) return _T("FOLDERID_SidebarDefaultParts");
    if (folderGuid == FOLDERID_PublicGameTasks) return _T("FOLDERID_PublicGameTasks");
    if (folderGuid == FOLDERID_GameTasks) return _T("FOLDERID_GameTasks");
    if (folderGuid == FOLDERID_SavedGames) return _T("FOLDERID_SavedGames");
    if (folderGuid == FOLDERID_Games) return _T("FOLDERID_Games");
    if (folderGuid == FOLDERID_SEARCH_MAPI) return _T("FOLDERID_SEARCH_MAPI");
    if (folderGuid == FOLDERID_SEARCH_CSC) return _T("FOLDERID_SEARCH_CSC");
    if (folderGuid == FOLDERID_Links) return _T("FOLDERID_Links");
    if (folderGuid == FOLDERID_UsersFiles) return _T("FOLDERID_UsersFiles");
    if (folderGuid == FOLDERID_UsersLibraries) return _T("FOLDERID_UsersLibraries");
    if (folderGuid == FOLDERID_SearchHome) return _T("FOLDERID_SearchHome");
    if (folderGuid == FOLDERID_OriginalImages) return _T("FOLDERID_OriginalImages");
    if (folderGuid == FOLDERID_DocumentsLibrary) return _T("FOLDERID_DocumentsLibrary");
    if (folderGuid == FOLDERID_MusicLibrary) return _T("FOLDERID_MusicLibrary");
    if (folderGuid == FOLDERID_PicturesLibrary) return _T("FOLDERID_PicturesLibrary");
    if (folderGuid == FOLDERID_VideosLibrary) return _T("FOLDERID_VideosLibrary");
    if (folderGuid == FOLDERID_RecordedTVLibrary) return _T("FOLDERID_RecordedTVLibrary");
    if (folderGuid == FOLDERID_HomeGroup) return _T("FOLDERID_HomeGroup");
    if (folderGuid == FOLDERID_HomeGroupCurrentUser) return _T("FOLDERID_HomeGroupCurrentUser");
    if (folderGuid == FOLDERID_DeviceMetadataStore) return _T("FOLDERID_DeviceMetadataStore");
    if (folderGuid == FOLDERID_Libraries) return _T("FOLDERID_Libraries");
    if (folderGuid == FOLDERID_PublicLibraries) return _T("FOLDERID_PublicLibraries");
    if (folderGuid == FOLDERID_UserPinned) return _T("FOLDERID_UserPinned");
    if (folderGuid == FOLDERID_ImplicitAppShortcuts) return _T("FOLDERID_ImplicitAppShortcuts");
    if (folderGuid == FOLDERID_AccountPictures) return _T("FOLDERID_AccountPictures");
    if (folderGuid == FOLDERID_PublicUserTiles) return _T("FOLDERID_PublicUserTiles");
    if (folderGuid == FOLDERID_AppsFolder) return _T("FOLDERID_AppsFolder");
    if (folderGuid == FOLDERID_StartMenuAllPrograms) return _T("FOLDERID_StartMenuAllPrograms");
    if (folderGuid == FOLDERID_CommonStartMenuPlaces) return _T("FOLDERID_CommonStartMenuPlaces");
    if (folderGuid == FOLDERID_ApplicationShortcuts) return _T("FOLDERID_ApplicationShortcuts");
    if (folderGuid == FOLDERID_RoamingTiles) return _T("FOLDERID_RoamingTiles");
    if (folderGuid == FOLDERID_RoamedTileImages) return _T("FOLDERID_RoamedTileImages");
    if (folderGuid == FOLDERID_Screenshots) return _T("FOLDERID_Screenshots");
    if (folderGuid == FOLDERID_CameraRoll) return _T("FOLDERID_CameraRoll");
    if (folderGuid == FOLDERID_SkyDrive) return _T("FOLDERID_SkyDrive");
    if (folderGuid == FOLDERID_OneDrive) return _T("FOLDERID_OneDrive");
    if (folderGuid == FOLDERID_SkyDriveDocuments) return _T("FOLDERID_SkyDriveDocuments");
    if (folderGuid == FOLDERID_SkyDrivePictures) return _T("FOLDERID_SkyDrivePictures");
    if (folderGuid == FOLDERID_SkyDriveMusic) return _T("FOLDERID_SkyDriveMusic");
    if (folderGuid == FOLDERID_SkyDriveCameraRoll) return _T("FOLDERID_SkyDriveCameraRoll");
    if (folderGuid == FOLDERID_SearchHistory) return _T("FOLDERID_SearchHistory");
    if (folderGuid == FOLDERID_SearchTemplates) return _T("FOLDERID_SearchTemplates");
    if (folderGuid == FOLDERID_CameraRollLibrary) return _T("FOLDERID_CameraRollLibrary");
    if (folderGuid == FOLDERID_SavedPictures) return _T("FOLDERID_SavedPictures");
    if (folderGuid == FOLDERID_SavedPicturesLibrary) return _T("FOLDERID_SavedPicturesLibrary");
    if (folderGuid == FOLDERID_RetailDemo) return _T("FOLDERID_RetailDemo");
    if (folderGuid == FOLDERID_Device) return _T("FOLDERID_Device");
    if (folderGuid == FOLDERID_DevelopmentFiles) return _T("FOLDERID_DevelopmentFiles");
    if (folderGuid == FOLDERID_Objects3D) return _T("FOLDERID_Objects3D");
    if (folderGuid == FOLDERID_AppCaptures) return _T("FOLDERID_AppCaptures");
    if (folderGuid == FOLDERID_LocalDocuments) return _T("FOLDERID_LocalDocuments");
    if (folderGuid == FOLDERID_LocalPictures) return _T("FOLDERID_LocalPictures");
    if (folderGuid == FOLDERID_LocalVideos) return _T("FOLDERID_LocalVideos");
    if (folderGuid == FOLDERID_LocalMusic) return _T("FOLDERID_LocalMusic");
    if (folderGuid == FOLDERID_LocalDownloads) return _T("FOLDERID_LocalDownloads");
    if (folderGuid == FOLDERID_RecordedCalls) return _T("FOLDERID_RecordedCalls");
    if (folderGuid == FOLDERID_AllAppMods) return _T("FOLDERID_AllAppMods");
    if (folderGuid == FOLDERID_CurrentAppMods) return _T("FOLDERID_CurrentAppMods");
    if (folderGuid == FOLDERID_AppDataDesktop) return _T("FOLDERID_AppDataDesktop");
    if (folderGuid == FOLDERID_AppDataDocuments) return _T("FOLDERID_AppDataDocuments");
    if (folderGuid == FOLDERID_AppDataFavorites) return _T("FOLDERID_AppDataFavorites");
    if (folderGuid == FOLDERID_AppDataProgramData) return _T("FOLDERID_AppDataProgramData");
    if (folderGuid == FOLDERID_LocalStorage) return _T("FOLDERID_LocalStorage");

    return _T("UNKNOWN_FOLDERID"); // Î´ÖªµÄ GUID
}


void CopySelectItemInfo(HWND hwnd)
{
    UINT itemid = -1;
    itemid = ListView_GetNextItem(hListView, itemid, LVNI_SELECTED);
    std::tstring strInfo;
    while (itemid != -1)
    {
        for (int i = 0; i < vecCols.size(); i++)
        {
            TCHAR szText[128] = {};
            ListView_GetItemText(hListView, itemid, i, szText, ARRAYSIZE(szText));
            strInfo += vecCols[i];
            strInfo += _T(" : ");
            strInfo += szText;
            strInfo += _T("\n");
        }
        strInfo += _T("\n");
        itemid = ListView_GetNextItem(hListView, itemid, LVNI_SELECTED);
    }
    CopyTextToClipboard(hwnd, strInfo);
}

void CopyTextToClipboard(HWND hwnd, const std::tstring& text)
{
    BOOL bCopied = FALSE;
    if (OpenClipboard(NULL))
    {
        if (EmptyClipboard())
        {
            size_t len = text.length();
            HGLOBAL hData = GlobalAlloc(GMEM_MOVEABLE | GMEM_DDESHARE, (len + 1) * sizeof(TCHAR));
            if (hData != NULL)
            {
                TCHAR* pszData = static_cast<TCHAR*>(GlobalLock(hData));
                if (pszData != NULL)
                {
                    //CopyMemory(pszData, text.c_str(), len*sizeof(TCHAR));
                    _tcscpy_s(pszData, len + 1, text.c_str());
                    GlobalUnlock(hData);
                    if (SetClipboardData(CF_UNICODETEXT, hData) != NULL)
                    {
                        bCopied = TRUE;;
                    }
                }
            }
        }
        CloseClipboard();
    }
}