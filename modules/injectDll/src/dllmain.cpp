// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"

#include <sstream>
#include <string>

typedef HRESULT(WINAPI* TRegisterDragDrop)(HWND, LPDROPTARGET);
typedef HRESULT(WINAPI* TDrop)
(	void * point,
	IDataObject *pDataObj,
	DWORD       grfKeyState,
	POINTL      pt,
	DWORD       *pdwEffect
);

typedef INT_PTR(WINAPI* TDialogBoxIndirectParamW)
(	HINSTANCE       hInstance,
	LPCDLGTEMPLATEW hDialogTemplate,
	HWND            hWndParent,
	DLGPROC         lpDialogFunc,
	LPARAM          dwInitParam
);

typedef INT_PTR (CALLBACK * TDialogProc)
(
	HWND   hwndDlg,
	UINT   uMsg,
	WPARAM  wParam,
	LPARAM  lParam
);

blackbone::Detour<TRegisterDragDrop> detourRegisterDragDrop;
blackbone::Detour<TDrop> detourDrop;
blackbone::Detour<TDialogBoxIndirectParamW> detourDialogBoxIndirectParamW;
blackbone::Detour<TDialogProc> detourDialogProc;

IUIAutomation * pAutomation;

void PrintMsg(const std::wstring & msg)
{
	std::wstring fullMsg = L"TestTask: ";
	fullMsg += msg;
	OutputDebugString(fullMsg.c_str());
}

std::wstring GetFolder(IUIAutomationElement * element)
{
	std::wstring path;
	IUIAutomationCondition* pTypeControlCondition = NULL;
	IUIAutomationElementArray* pFound = NULL;
	
	CComVariant varTypeValue(UIA_ToolBarControlTypeId);
	pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varTypeValue, &pTypeControlCondition);
	if (pTypeControlCondition == NULL)
	{
		goto cleanup;
	}
	
	element->FindAll(TreeScope_Subtree, pTypeControlCondition, &pFound);
	if (pFound == NULL)
	{
		goto cleanup;
	}

	int len;
	pFound->get_Length(&len);

	IUIAutomationElement * foundElement;
	for (int i = 0; i < len; ++i)
	{
		foundElement = NULL;
		pFound->GetElement(i, &foundElement);
		if (foundElement != NULL)
		{
			CComBSTR value;
			foundElement->get_CurrentName(&value);
			if (value.m_str)
			{
				std::wstring name = value.m_str;
				std::size_t pos = name.find(L"Адрес:");
				if (pos != std::wstring::npos)
				{
					path = name.substr(pos + 6);
					foundElement->Release();
					return path;
				}
			}
			foundElement->Release();
		}
	}
	
cleanup:
	if (pTypeControlCondition != NULL)
	{
		pTypeControlCondition->Release();
	}
	
	if (pFound != NULL)
	{
		pFound->Release();
	}
	return path;
}

std::vector<std::wstring> GetSelectedFiles(IUIAutomationElement * element)
{
	std::vector<std::wstring> selectedFiles;
	
	IUIAutomationCondition* pTypeControlCondition = NULL;
	IUIAutomationElement* pFound = NULL;
	IUIAutomationElementArray* pselElements = NULL;

	CComVariant varTypeValue(UIA_ListControlTypeId);
	pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varTypeValue, &pTypeControlCondition);
	if (pTypeControlCondition == NULL)
	{
		goto cleanup;
	}

	element->FindFirst(TreeScope_Subtree, pTypeControlCondition, &pFound);
	if (pFound == NULL)
	{
		goto cleanup;
	}

	IUIAutomationSelectionPattern * selectionPattern = NULL;
	pFound->GetCurrentPatternAs(UIA_SelectionPatternId, IID_PPV_ARGS(&selectionPattern));
	if (selectionPattern == NULL)
	{
		goto cleanup;
	}

	selectionPattern->GetCurrentSelection(&pselElements);
	if (pselElements == NULL)
	{
		goto cleanup;
	}

	int len;
	pselElements->get_Length(&len);

	IUIAutomationElement * foundElement;
	for (int i = 0; i < len; ++i)
	{
		foundElement = NULL;
		pselElements->GetElement(i, &foundElement);
		if (foundElement != NULL)
		{
			CComBSTR value;
			foundElement->get_CurrentName(&value);
			if (value.m_str)
			{
				selectedFiles.push_back(value.m_str);
				PrintMsg(value.m_str);
			}
			foundElement->Release();
		}
	}

cleanup:
	if (pTypeControlCondition != NULL)
	{
		pTypeControlCondition->Release();
	}

	if (pFound != NULL)
	{
		pFound->Release();
	}

	if (selectionPattern != NULL)
	{
		selectionPattern->Release();
	}

	if (pselElements != NULL)
	{
		pselElements->Release();
	}

	return selectedFiles;
}

std::vector<std::wstring> GetFileLists(IUIAutomationElement * element)
{
	std::wstring path = GetFolder(element);
	std::vector<std::wstring> fileList = GetSelectedFiles(element);
	for (int i = 0; i < fileList.size(); ++i)
	{
		fileList[i] = path + L"\\" + fileList[i];
		PrintMsg(fileList[i]);
	}
	return fileList;
}

INT_PTR CALLBACK hkDialogProc(
	HWND &  hwndDlg,
	UINT & uMsg,
	WPARAM & wParam,
	LPARAM & lParam
)
{
	//PrintMsg(L"hkDialogProc");
	switch (uMsg)
	{
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDOK:

			HRESULT hr = CoCreateInstance(CLSID_CUIAutomation, NULL,
				CLSCTX_INPROC_SERVER, IID_IUIAutomation,
				reinterpret_cast<void**>(&pAutomation));

			if (FAILED(hr))
			{
				return 1;
			}

			IUIAutomationElement * element = NULL;
			pAutomation->ElementFromHandle(hwndDlg, &element);
			if (SUCCEEDED(hr) && element)
			{
				GetFileLists(element);
			}

			pAutomation->Release();
			if (element)
			{
				element->Release();
			}
		}
	}

	return 1;
}

HRESULT WINAPI  hkDrop
(	
	void * & point,
	IDataObject * & pDataObj,
	DWORD      & grfKeyState,
	POINTL     & pt,
	DWORD    * &pdwEffect
	)
{
	PrintMsg(L"hkDrop");
	
	FORMATETC fmte = { CF_HDROP, NULL, DVASPECT_CONTENT,
					-1, TYMED_HGLOBAL };
	STGMEDIUM stgm;

	if (!SUCCEEDED(pDataObj->GetData(&fmte, &stgm)))
	{
		return S_OK;
	}

	HDROP hdrop = reinterpret_cast<HDROP>(stgm.hGlobal);
	
	UINT cFiles = DragQueryFile(hdrop, 0xFFFFFFFF, NULL, 0);
	
	for (UINT i = 0; i < cFiles; i++) 
	{
		TCHAR szFile[MAX_PATH];
		UINT cch = DragQueryFile(hdrop, i, szFile, MAX_PATH);
		if (cch > 0 && cch < MAX_PATH) 
		{
			PrintMsg(szFile);
		}
	}
	ReleaseStgMedium(&stgm);
	
	return S_OK;
}

HRESULT WINAPI hkRegisterDragDrop(
	IN HWND      &   hwnd,
	IN LPDROPTARGET & pDropTarget
)
{
	PrintMsg(L"hkRegisterDragDrop");
	DWORD64 addrVirtTable = *((DWORD64*)(pDropTarget));
	addrVirtTable += 6 * 8;
	TDrop pDropFunc = reinterpret_cast<TDrop>(*((DWORD64*)addrVirtTable));

	detourDrop.Hook(pDropFunc, &hkDrop, blackbone::HookType::Inline, blackbone::CallOrder::HookLast);

	return ERROR_SUCCESS;
}

INT_PTR WINAPI hkDialogBoxIndirectParamW
(HINSTANCE       & hInstance,
	LPCDLGTEMPLATEW & hDialogTemplate,
	HWND            & hWndParent,
	DLGPROC         & lpDialogFunc,
	LPARAM         &  dwInitParam
	)
{
	PrintMsg(L"hkDialogBoxIndirectParamW");
	
	detourDialogProc.Hook(lpDialogFunc, &hkDialogProc, blackbone::HookType::Inline, blackbone::CallOrder::HookFirst);
	
	return 1;
}

std::string ProcessIdToName(DWORD processId)
{
	std::string ret;
	HANDLE handle = OpenProcess(
		PROCESS_QUERY_LIMITED_INFORMATION,
		FALSE,
		processId 
	);
	
	if (!handle)
	{
		return ret;
	}
	
	DWORD buffSize = 1024;
	CHAR buffer[1024];
	if (QueryFullProcessImageNameA(handle, 0, buffer, &buffSize))
	{
		ret = buffer;
	}
	CloseHandle(handle);

	return ret;
}


BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved

                     )
{
	HMODULE hOle32;
	HMODULE hUser32;
	std::vector<std::string> browserNames;
	std::string currProcName;
	bool isOurBrowser = false;
	TRegisterDragDrop pRegisterDragDrop = NULL;
	TDialogBoxIndirectParamW pDialogBoxIndirectParamW = NULL;
	bool res;
	
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		browserNames.push_back("chrome.exe");
		browserNames.push_back("firefox.exe");
		currProcName = ProcessIdToName(GetCurrentProcessId());
		
		for (auto browserName : browserNames)
		{
			if (currProcName.find(browserName) != std::string::npos)
			{
				isOurBrowser = true;
				break;
			}
		}
		if (!isOurBrowser)
		{
			break;
		}
		
		hOle32 = GetModuleHandle(TEXT("Ole32"));
		
		CoInitialize(NULL);
		if (hOle32)
		{
			pRegisterDragDrop = reinterpret_cast<TRegisterDragDrop>(::GetProcAddress(hOle32, "RegisterDragDrop"));
			res = detourRegisterDragDrop.Hook(pRegisterDragDrop, &hkRegisterDragDrop, blackbone::HookType::Inline, blackbone::CallOrder::HookLast);
		}
		
		hUser32 = GetModuleHandle(TEXT("User32"));
		if (hUser32)
		{
			pDialogBoxIndirectParamW = reinterpret_cast<TDialogBoxIndirectParamW>(::GetProcAddress(hUser32, "DialogBoxIndirectParamW"));
			res = detourDialogBoxIndirectParamW.Hook(pDialogBoxIndirectParamW, &hkDialogBoxIndirectParamW, blackbone::HookType::Inline, blackbone::CallOrder::HookFirst);
		}
	
		break;
	
	//case DLL_THREAD_ATTACH:
    //case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
		CoUninitialize();
	break;

	}
    return TRUE;
}

