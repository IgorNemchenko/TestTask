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
/*
typedef HRESULT (WINAPI* TCoCreateInstance)
(
	REFCLSID  rclsid,
	LPUNKNOWN pUnkOuter,
	DWORD     dwClsContext,
	REFIID    riid,
	LPVOID    *ppv
);
*/
typedef INT_PTR(WINAPI* TDialogBoxIndirectParamW)
(	HINSTANCE       hInstance,
	LPCDLGTEMPLATEW hDialogTemplate,
	HWND            hWndParent,
	DLGPROC         lpDialogFunc,
	LPARAM          dwInitParam
);

/*typedef HRESULT(WINAPI* TGetResult)
(
	void *point,
	IShellItem **ppsi
);

typedef HRESULT(WINAPI* TGetResults)
(	void *point,
	IShellItemArray **ppenum
);

typedef HRESULT(WINAPI* TShow)
(
	void *point,
	HWND hwndOwner
);
*/
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

//blackbone::Detour<TCoCreateInstance> detourCoCreateInstance;
//blackbone::Detour<TGetResult> detourGetResult;
//blackbone::Detour<TGetResults> detourGetResults;
//blackbone::Detour<TShow> detourShow;
//IFileDialog *pfd = NULL;
//IShellItem * psiResult = NULL;

void PrintMsg(const std::wstring & msg)
{
	std::wstring fullMsg = L"TestTask: ";
	fullMsg += msg;
	OutputDebugString(fullMsg.c_str());
}

std::wstring GetName(IAccessible *pAcc)
{
	CComBSTR bstrName;
	if (!pAcc || FAILED(pAcc->get_accName(CComVariant((int)CHILDID_SELF), &bstrName)) || !bstrName.m_str)
		return L"";

	return bstrName.m_str;
}

HRESULT FindSelectedFiles(CComPtr<IAccessible> pAcc)
{
	long childCount = 0;
	long returnCount = 0;

	HRESULT hr = pAcc->get_accChildCount(&childCount);

	if (childCount == 0)
		return S_OK;

	CComVariant* pArray = new CComVariant[childCount];
	hr = ::AccessibleChildren(pAcc, 0L, childCount, pArray, &returnCount);
	if (FAILED(hr))
		return hr;

	for (int x = 0; x < returnCount; x++)
	{
		CComVariant vtChild = pArray[x];
		if (vtChild.vt != VT_DISPATCH)
			continue;

		CComPtr<IDispatch> pDisp = vtChild.pdispVal;
		CComQIPtr<IAccessible> pAccChild = pDisp;
		if (!pAccChild)
			continue;

		std::wstring name = GetName(pAccChild).data();
				
		CComVariant roleName;
		HRESULT h = pAccChild->get_accRole(CComVariant((int)CHILDID_SELF), &roleName);

		if ((hr == S_OK) && 
			(roleName.vt == VT_I4) && 
			(roleName.lVal == 0x21) && 
			name.find(L"Просмотр элементов") != -1)
		{
			CComVariant SelElement;
			pAccChild->get_accSelection(&SelElement);
			if (SelElement.punkVal)
			{
				CComPtr<IUnknown> pUnk = SelElement.punkVal;
				CComPtr <IEnumVARIANT> pList;
				pUnk->QueryInterface(IID_IEnumVARIANT, (void**)&pList);
				
				CComVariant currElement;
				while (S_OK == pList->Next(1, &currElement, 0))
				{
					CComPtr<IDispatch> pDisp = currElement.pdispVal;
					CComQIPtr<IAccessible> pAccChild = pDisp;
					if (!pAccChild)
						continue;
					std::wstring name = GetName(pAccChild).data();
					PrintMsg(name);
				}
			}
		}
		if (FindSelectedFiles(pAccChild) == S_FALSE)
			return S_FALSE;
	}

	delete[] pArray;
	return S_OK;
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

			CComPtr<IAccessible> pAccMain1;
			
			HRESULT hr = ::AccessibleObjectFromWindow(hwndDlg, 1, IID_IAccessible, (void**)(&pAccMain1)); 
			CComPtr<IAccessible> pAccMain2;
			
			hr = ::AccessibleObjectFromWindow(hwndDlg, OBJID_CLIENT, IID_IAccessible, (void**)(&pAccMain2));
			FindSelectedFiles(pAccMain2);
		}
	}

	return 1;
}

/*HRESULT WINAPI hkGetResult
(
	void * & point,
	IShellItem ** & ppsi
)
{
	PrintMsg(L"hkGetResult");
	return S_OK;
}

HRESULT WINAPI hkGetResults
(
	void * &point,
	IShellItemArray ** & ppenum
)
{
	PrintMsg(L"hkGetResults"); 
	return S_FALSE;
}

HRESULT WINAPI hkShow
(
	void * & point,
	HWND & hwndOwner
)
{
	PrintMsg(L"hkShow");
	return S_OK;
}*/

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
/*
HRESULT WINAPI hkCoCreateInstance
(
	REFCLSID   rclsid,
	LPUNKNOWN & pUnkOuter,
	DWORD     & dwClsContext,
	REFIID    riid,
	LPVOID *  & ppv
)
{
	//PrintMsg(L"hkCoCreateInstance");
	LPOLESTR clsidStr;
	StringFromCLSID(
		rclsid,
		&clsidStr
	);
	//PrintMsg(clsidStr);
	std::wstring clsidCurr = clsidStr;
	CoTaskMemFree(clsidStr);
	
	if (clsidCurr != L"{725F645B-EAED-4FC5-B1C5-D9AD0ACCBA5E}")
	{
		return S_OK;
	}
	
	pfd = reinterpret_cast<IFileOpenDialog*>(*ppv);

	DWORD64 addrVirtTable = *((DWORD64*)(pfd));
	
	addrVirtTable += 27 * 8;
	
	TGetResults pGetResults = reinterpret_cast<TGetResults>(*((DWORD64*)addrVirtTable));
	detourGetResults.Hook(pGetResults, &hkGetResults, blackbone::HookType::Inline, blackbone::CallOrder::HookFirst);
	
	
	addrVirtTable = *((DWORD64*)(pfd));

	addrVirtTable += 3 * 8;
	
	TShow pShow = reinterpret_cast<TShow>(*((DWORD64*)addrVirtTable));
	detourShow.Hook(pShow, &hkShow, blackbone::HookType::Inline, blackbone::CallOrder::HookFirst);

	return S_OK;
}
*/
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
	//TCoCreateInstance pCoCreateInstance = NULL;
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

			//pCoCreateInstance = reinterpret_cast<TCoCreateInstance>(::GetProcAddress(hOle32, "CoCreateInstance"));
			//res = detourCoCreateInstance.Hook(pCoCreateInstance, &hkCoCreateInstance, blackbone::HookType::Inline, blackbone::CallOrder::HookLast);

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
	break;

	}
    return TRUE;
}

