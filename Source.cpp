#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "shlwapi")
#pragma comment(lib, "imagehlp")
#pragma comment(lib, "comctl32")

#include <windows.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <imagehlp.h>
#include <string>

#define ID_COPYTOCLIPBOARD 1000
#define ID_SELECTALL 1001
#define ID_CHECK1 1002

TCHAR szClassName[] = TEXT("Window");

LPTSTR lpszStandardLibrary[] =
{
	TEXT("ADVAPI32.DLL"),
	TEXT("AVICAP32.DLL"),
	TEXT("CABINET.DLL"),
	TEXT("COMCTL32.DLL"),
	TEXT("COMDLG32.DLL"),
	TEXT("CRYPT32.DLL"),
	TEXT("D3D9.DLL"),
	TEXT("DBGHELP.DLL"),
	TEXT("DINPUT8.DLL"),
	TEXT("DNSAPI.DLL"),
	TEXT("DSOUND.DLL"),
	TEXT("DWMAPI.DLL"),
	TEXT("GDI32.DLL"),
	TEXT("GDIPLUS.DLL"),
	TEXT("GLU32.DLL"),
	TEXT("IMAGEHLP.DLL"),
	TEXT("IMM32.DLL"),
	TEXT("IPHLPAPI.DLL"),
	TEXT("KERNEL32.DLL"),
	TEXT("MMDEVAPI.DLL"),
	TEXT("MPR.DLL"),
	TEXT("MSACM32.DLL"),
	TEXT("MSI.DLL"),
	TEXT("MSIMG32.DLL"),
	TEXT("MSWSOCK.DLL"),
	TEXT("NETAPI32.DLL"),
	TEXT("NORMALIZ.DLL"),
	TEXT("NTDLL.DLL"),
	TEXT("ODBC32.DLL"),
	TEXT("OLE32.DLL"),
	TEXT("OLEACC.DLL"),
	TEXT("OLEAUT32.DLL"),
	TEXT("OLEDLG.DLL"),
	TEXT("OLEPRO32.DLL"),
	TEXT("OPENGL32.DLL"),
	TEXT("PDH.DLL"),
	TEXT("POWRPROF.DLL"),
	TEXT("PROPSYS.DLL"),
	TEXT("PSAPI.DLL"),
	TEXT("RASAPI32.DLL"),
	TEXT("RPCRT4.DLL"),
	TEXT("SECUR32.DLL"),
	TEXT("SENSAPI.DLL"),
	TEXT("SETUPAPI.DLL"),
	TEXT("SHELL32.DLL"),
	TEXT("SHFOLDER.DLL"),
	TEXT("SHLWAPI.DLL"),
	TEXT("URLMON.DLL"),
	TEXT("USER32.DLL"),
	TEXT("USERENV.DLL"),
	TEXT("USP10.DLL"),
	TEXT("UXTHEME.DLL"),
	TEXT("VERSION.DLL"),
	TEXT("WINDOWSCODECS.DLL"),
	TEXT("WINHTTP.DLL"),
	TEXT("WININET.DLL"),
	TEXT("WINMM.DLL"),
	TEXT("WINSPOOL.DRV"),
	TEXT("WINTRUST.DLL"),
	TEXT("WLDAP32.DLL"),
	TEXT("WS2_32.DLL"),
	TEXT("WSOCK32.DLL"),
	TEXT("WTSAPI32.DLL"),
};

template <class T> PIMAGE_SECTION_HEADER GetEnclosingSectionHeader(DWORD rva, T* pNTHeader) // 'T' == PIMAGE_NT_HEADERS 
{
	PIMAGE_SECTION_HEADER section = IMAGE_FIRST_SECTION(pNTHeader);
	unsigned i;

	for (i = 0; i < pNTHeader->FileHeader.NumberOfSections; ++i, ++section)
	{
		// This 3 line idiocy is because Watcom's linker actually sets the
		// Misc.VirtualSize field to 0.  (!!! - Retards....!!!)
		DWORD size = section->Misc.VirtualSize;
		if (0 == size)
			size = section->SizeOfRawData;

		// Is the RVA within this section?
		if ((rva >= section->VirtualAddress) &&
			(rva < (section->VirtualAddress + size)))
			return section;
	}

	return 0;
}

template <class T> LPVOID GetPtrFromRVA(DWORD rva, T* pNTHeader, PBYTE imageBase) // 'T' = PIMAGE_NT_HEADERS 
{
	PIMAGE_SECTION_HEADER pSectionHdr = GetEnclosingSectionHeader(rva, pNTHeader);
	if (!pSectionHdr)
		return 0;

	INT delta = (INT)(pSectionHdr->VirtualAddress - pSectionHdr->PointerToRawData);
	return (PVOID)(imageBase + rva - delta);
}

void DumpDllFromPath(HWND hList, HWND hBackgroundList2, HWND hBackgroundList3, wchar_t* pathW)
{
	DWORD len = WideCharToMultiByte(CP_ACP, 0, pathW, -1, 0, 0, 0, 0);
	LPSTR pathA = (LPSTR)GlobalAlloc(GPTR, len * sizeof(CHAR));
	WideCharToMultiByte(CP_ACP, 0, pathW, -1, pathA, len, 0, 0);

	std::string str;

	str += "[";
	str += PathFindFileNameA(pathA);
	str += "] ";

	PLOADED_IMAGE image = ImageLoad(pathA, 0);

	if (image && image->FileHeader->OptionalHeader.NumberOfRvaAndSizes >= 2)
	{
		PIMAGE_IMPORT_DESCRIPTOR importDesc =
			(PIMAGE_IMPORT_DESCRIPTOR)GetPtrFromRVA(
				image->FileHeader->OptionalHeader.DataDirectory[1].VirtualAddress,
				image->FileHeader, image->MappedAddress);

		while (1)
		{
			// See if we've reached an empty IMAGE_IMPORT_DESCRIPTOR
			if (importDesc == 0 || ((importDesc->TimeDateStamp == 0) && (importDesc->Name == 0)))
				break;
			LPCSTR lpszModuleName = (LPCSTR)GetPtrFromRVA(importDesc->Name,
				image->FileHeader,
				image->MappedAddress);
			if (SendMessageA(hBackgroundList3, LB_FINDSTRINGEXACT, -1, (LPARAM)lpszModuleName) == LB_ERR &&
				SendMessageA(hBackgroundList2, LB_FINDSTRINGEXACT, -1, (LPARAM)lpszModuleName) == LB_ERR)
			{
				SendMessageA(hBackgroundList2, LB_ADDSTRING, 0, (LPARAM)lpszModuleName);
			}
			importDesc++;
		}
	}
	ImageUnload(image);

	CHAR szModuleName[MAX_PATH];

	while (SendMessageA(hBackgroundList2, LB_GETCOUNT, 0, 0))
	{
		SendMessageA(hBackgroundList2, LB_GETTEXT, 0, (LPARAM)szModuleName);
		str += (LPSTR)szModuleName;
		str += ", ";
		SendMessageW(hBackgroundList2, LB_DELETESTRING, 0, 0);
	}

	while (*str.rbegin() == ' ' || *str.rbegin() == ',')
	{
		str.pop_back();
	}

	if (SendMessageA(hList, LB_FINDSTRINGEXACT, -1, (LPARAM)str.c_str()) == LB_ERR)
	{
		SendMessageA(hList, LB_ADDSTRING, 0, (LPARAM)str.c_str());
	}

	GlobalFree(pathA);
}

BOOL IsTargetFile(LPCWSTR lpszFilePath)
{
	WCHAR szExtList[] = L"*.exe;*.dll;";
	LPCWSTR seps = L";";
	WCHAR *next;
	LPWSTR token = wcstok_s(szExtList, seps, &next);
	while (token != NULL)
	{
		if (PathMatchSpecW(lpszFilePath, token))
		{
			return TRUE;
		}
		token = wcstok_s(0, seps, &next);
	}
	return FALSE;
}

VOID TargetFileCount(HWND hBackgroundList1, LPCWSTR lpInputPath)
{
	WCHAR szFullPattern[MAX_PATH];
	WIN32_FIND_DATAW FindFileData;
	HANDLE hFindFile;
	PathCombineW(szFullPattern, lpInputPath, L"*");
	hFindFile = FindFirstFileW(szFullPattern, &FindFileData);
	if (hFindFile != INVALID_HANDLE_VALUE)
	{
		do
		{
			if (FindFileData.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY)
			{
				if (lstrcmpW(FindFileData.cFileName, L"..") != 0 &&
					lstrcmpW(FindFileData.cFileName, L".") != 0)
				{
					PathCombineW(szFullPattern, lpInputPath, FindFileData.cFileName);
					TargetFileCount(hBackgroundList1, szFullPattern);
				}
			}
			else
			{
				PathCombineW(szFullPattern, lpInputPath, FindFileData.cFileName);
				if (IsTargetFile(szFullPattern))
				{
					SendMessageW(hBackgroundList1, LB_ADDSTRING, 0, (LPARAM)szFullPattern);
				}
			}
		} while (FindNextFileW(hFindFile, &FindFileData));
		FindClose(hFindFile);
	}
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hList, hProgress;
	static HWND hBackgroundList1;
	static HWND hBackgroundList2;
	static HWND hBackgroundList3;
	static HWND hCheck;
	static HBRUSH hbrBkgnd;
	switch (msg)
	{
	case WM_CREATE:
		InitCommonControls();
		hList = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("LISTBOX"), 0, WS_CHILD | WS_VSCROLL | LBS_NOINTEGRALHEIGHT | LBS_EXTENDEDSEL | LBS_MULTIPLESEL | LBS_SORT, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hBackgroundList1 = CreateWindow(TEXT("LISTBOX"), 0, 0, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hBackgroundList2 = CreateWindow(TEXT("LISTBOX"), 0, 0, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hBackgroundList3 = CreateWindow(TEXT("LISTBOX"), 0, 0, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hCheck = CreateWindow(TEXT("BUTTON"), TEXT("標準DLLを無視する"), WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 0, 0, 0, 0, hWnd, (HMENU)ID_CHECK1, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessage(hCheck, BM_SETCHECK, TRUE, 0);
		hProgress = CreateWindow(TEXT("msctls_progress32"), 0, WS_CHILD, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_SIZE:
		MoveWindow(hCheck, 0, 0, 256, 32, TRUE);
		MoveWindow(hList, 0, 32, LOWORD(lParam), HIWORD(lParam) - 32, TRUE);
		MoveWindow(hProgress, 10, HIWORD(lParam) / 2 - 64 / 2, LOWORD(lParam) - 20, 64, 1);
		break;
	case WM_DROPFILES:
		{
			SendMessage(hBackgroundList3, LB_RESETCONTENT, 0, 0);
			if (SendMessage(hCheck, BM_GETCHECK, 0, 0))
			{
				for (auto v : lpszStandardLibrary)
				{
					SendMessage(hBackgroundList3, LB_ADDSTRING, 0, (LPARAM)v);
				}
			}
			ShowWindow(hList, SW_HIDE);
			SendMessage(hList, LB_RESETCONTENT, 0, 0);
			SendMessage(hProgress, PBM_SETPOS, 0, 0);
			SendMessage(hProgress, PBM_SETMARQUEE, TRUE, 0);
			ShowWindow(hProgress, SW_SHOW);
			WCHAR szFilePath[MAX_PATH];
			const UINT iFileNum = DragQueryFileW((HDROP)wParam, -1, NULL, 0);
			for (UINT i = 0; i < iFileNum; ++i)
			{
				DragQueryFileW((HDROP)wParam, i, szFilePath, MAX_PATH);
				if (PathIsDirectoryW(szFilePath))
				{
					TargetFileCount(hBackgroundList1, szFilePath);
				}
				else if (IsTargetFile(szFilePath))
				{
					SendMessageW(hBackgroundList1, LB_ADDSTRING, 0, (LPARAM)szFilePath);
				}
			}
			DragFinish((HDROP)wParam);
			SendMessage(hProgress, PBM_SETMARQUEE, FALSE, 0);
			SendMessage(hProgress, PBM_SETRANGE32, 0, SendMessage(hBackgroundList1, LB_GETCOUNT, 0, 0));
			SendMessage(hProgress, PBM_SETSTEP, 1, 0);
			while (SendMessageW(hBackgroundList1, LB_GETCOUNT, 0, 0))
			{
				SendMessageW(hBackgroundList1, LB_GETTEXT, 0, (LPARAM)szFilePath);
				DumpDllFromPath(hList, hBackgroundList2, hBackgroundList3, szFilePath);
				SendMessageW(hBackgroundList1, LB_DELETESTRING, 0, 0);
				SendMessage(hProgress, PBM_STEPIT, 0, 0);
			}
			ShowWindow(hProgress, SW_HIDE);
			ShowWindow(hList, SW_SHOW);
		}
		break;
	case WM_CTLCOLORSTATIC:
		SetBkColor((HDC)wParam, GetSysColor(COLOR_WINDOW));
		if (hbrBkgnd == NULL)
		{
			hbrBkgnd = CreateSolidBrush(GetSysColor(COLOR_WINDOW));
		}
		return (INT_PTR)hbrBkgnd;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case ID_COPYTOCLIPBOARD:
			{
				const int nSelItems = (int)SendMessage(hList, LB_GETSELCOUNT, 0, 0);
				if (nSelItems > 0)
				{
					int* pBuffer = (int*)GlobalAlloc(0, sizeof(int) * nSelItems);
					SendMessage(hList, LB_GETSELITEMS, nSelItems, (LPARAM)pBuffer);
					INT nLen = 1; // NULL分
					for (int i = 0; i < nSelItems; ++i)
					{
						nLen += (int)SendMessage(hList, LB_GETTEXTLEN, pBuffer[i], 0);
						nLen += 2; // 改行文字分
					}
					HGLOBAL hMem = GlobalAlloc(GMEM_DDESHARE | GMEM_MOVEABLE, sizeof(TCHAR)*(nLen + 1));
					LPTSTR lpszBuflpszBuf = (LPTSTR)GlobalLock(hMem);
					for (int i = 0; i < nSelItems; i++)
					{
						const int nSize = (int)SendMessage(hList, LB_GETTEXT, pBuffer[i], (LPARAM)lpszBuflpszBuf);
						lstrcat(lpszBuflpszBuf, TEXT("\r\n"));
						lpszBuflpszBuf += nSize + 2;
					}
					GlobalFree(pBuffer);
					GlobalUnlock(hMem);
					OpenClipboard(NULL);
					EmptyClipboard();
					SetClipboardData(CF_UNICODETEXT, hMem);
					CloseClipboard();
				}
			}
			break;
		case ID_SELECTALL:
			SendMessage(hList, LB_SETSEL, 1, -1);
			break;
		}
		break;
	case WM_DESTROY:
		DeleteObject(hbrBkgnd);
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("ドロップされたEXE,DLLが依存するDLLを列挙"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	ACCEL Accel[] = {
		{ FVIRTKEY | FCONTROL, 'A', ID_SELECTALL },
		{ FVIRTKEY | FCONTROL, 'C', ID_COPYTOCLIPBOARD },
	};
	HACCEL hAccel = CreateAcceleratorTable(Accel, sizeof(Accel) / sizeof(ACCEL));
	while (GetMessage(&msg, 0, 0, 0))
	{
		if (!TranslateAccelerator(hWnd, hAccel, &msg) && !IsDialogMessage(hWnd, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	DestroyAcceleratorTable(hAccel);
	return (int)msg.wParam;
}
