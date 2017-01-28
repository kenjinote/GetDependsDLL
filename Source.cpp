#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "shlwapi")
#pragma comment(lib, "Imagehlp")
#include <windows.h>
#include <shlwapi.h>
#include "ImageHlp.h"

TCHAR szClassName[] = TEXT("Window");

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
	PIMAGE_SECTION_HEADER pSectionHdr;
	INT delta;

	pSectionHdr = GetEnclosingSectionHeader(rva, pNTHeader);
	if (!pSectionHdr)
		return 0;

	delta = (INT)(pSectionHdr->VirtualAddress - pSectionHdr->PointerToRawData);
	return (PVOID)(imageBase + rva - delta);
}

void DumpDllFromPath(HWND hList, wchar_t* path)
{
	char name[300];
	wcstombs(name, path, 300);

	PLOADED_IMAGE image = ImageLoad(name, 0);

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

			CHAR szText[1024];
			wsprintfA(szText, "%s", (LPSTR)GetPtrFromRVA(importDesc->Name,
				image->FileHeader,
				image->MappedAddress));
			SendMessageA(hList, LB_ADDSTRING, 0, (LPARAM)szText);
			importDesc++;
		}
	}
	ImageUnload(image);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hList;
	switch (msg)
	{
	case WM_CREATE:
		hList = CreateWindow(TEXT("LISTBOX"), TEXT("変換"), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_SIZE:
		MoveWindow(hList, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
		break;
	case WM_DROPFILES:
	{
		TCHAR szFilePath[MAX_PATH];
		const UINT iFileNum = DragQueryFile((HDROP)wParam, -1, NULL, 0);
		if (iFileNum == 1)
		{
			DragQueryFile((HDROP)wParam, 0, szFilePath, MAX_PATH);
			if (PathMatchSpec(szFilePath, TEXT("*.exe")) || PathMatchSpec(szFilePath, TEXT("*.dll")))
			{
				SendMessage(hList, LB_RESETCONTENT, 0, 0);
				DumpDllFromPath(hList, szFilePath);
			}
		}
		DragFinish((HDROP)wParam);
	}
	break;
	case WM_DESTROY:
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
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
