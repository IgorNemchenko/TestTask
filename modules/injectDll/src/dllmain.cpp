// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"

#include <fstream>
#include <sstream>
#include <string>
#include "ole2.h"
HMODULE h = NULL;

typedef HMODULE(WINAPI* HOOKLoadLibraryW)(LPCWSTR);
typedef HMODULE(WINAPI* HOOKLoadLibraryA)(LPCSTR);
typedef HRESULT(WINAPI* HOOKRegisterDragDrop)(HWND, LPDROPTARGET);
typedef HRESULT(WINAPI*  HOOKDrop)
(	void * point,
	IDataObject *pDataObj,
	DWORD       grfKeyState,
	POINTL      pt,
	DWORD       *pdwEffect
);

HOOKLoadLibraryW addrW = NULL;
HOOKLoadLibraryA addrA = NULL;
HOOKRegisterDragDrop addrRegisterDragDrop = NULL;
blackbone::Detour<HOOKLoadLibraryW> detourLoadLibraryW;
blackbone::Detour<HOOKLoadLibraryA> detourLoadLibraryA;
blackbone::Detour<HOOKRegisterDragDrop> detourRegisterDragDrop;
blackbone::Detour<HOOKDrop> detourDrop;

HRESULT WINAPI  hkDrop
(	
	void * & point,
	IDataObject * & pDataObj,
	DWORD      & grfKeyState,
	POINTL     & pt,
	DWORD    * &pdwEffect
	)
{
	OutputDebugString(L"hkDrop");
	return S_OK;
}

HRESULT WINAPI hkRegisterDragDrop(
	IN HWND      &   hwnd,
	IN LPDROPTARGET & pDropTarget
)
{
	OutputDebugString(L"hkRegisterDragDrop");
	
	DWORD64 addrVirtTable = *((DWORD64*)(pDropTarget));
	addrVirtTable += 6 * 8;
	HOOKDrop DropFunc = (HOOKDrop)(*((DWORD64*)addrVirtTable));

	detourDrop.Hook(DropFunc, &hkDrop, blackbone::HookType::Inline, blackbone::CallOrder::HookFirst);

	return ERROR_SUCCESS;
}

HMODULE WINAPI hkLoadLibraryW(LPCWSTR & lpLibFileName)
{
	OutputDebugString(L"---LoadLibraryW");
	OutputDebugString(lpLibFileName);
	if (std::wstring(lpLibFileName) == L"Ole32")
	{
		OutputDebugString(L"---LoadLibraryW-Ole32.dll---");
	}
	
	return (HMODULE)1;
}

HMODULE WINAPI hkLoadLibraryA(LPCSTR & lpLibFileName)
{
	OutputDebugString(L"---LoadLibraryA");
	std::string temp(lpLibFileName);
	std::wstring tempW(temp.begin(), temp.end());
	OutputDebugString(tempW.c_str());
	

	return (HMODULE)1;
}


std::string ProcessIdToName(DWORD processId)
{
	std::string ret;
	HANDLE handle = OpenProcess(
		PROCESS_QUERY_LIMITED_INFORMATION,
		FALSE,
		processId 
	);
	if (handle)
	{
		DWORD buffSize = 1024;
		CHAR buffer[1024];
		if (QueryFullProcessImageNameA(handle, 0, buffer, &buffSize))
		{
			ret = buffer;
		}
		else
		{
			//printf("Error GetModuleBaseNameA : %lu", GetLastError());
		}
		CloseHandle(handle);
	}
	else
	{
		//printf("Error OpenProcess : %lu", GetLastError());
	}
	return ret;
}


BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved

                     )
{
	DWORD d;
	std::wstring text;
	std::ostringstream stream;
	std::string st;
	std::string path;
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		d = GetCurrentProcessId();
		path = ProcessIdToName(d);
		//if (path.find("chrome.exe") != std::string::npos)
		{

			h = GetModuleHandle(TEXT("Ole32"));
			if (h)
			{
				OutputDebugString(L"Ole32 has founded");

				addrRegisterDragDrop = (HOOKRegisterDragDrop)::GetProcAddress(h, "RegisterDragDrop");
				bool res = detourRegisterDragDrop.Hook(addrRegisterDragDrop, &hkRegisterDragDrop, blackbone::HookType::Inline, blackbone::CallOrder::HookLast);
			}
			else
			{
				h = GetModuleHandle(TEXT("Kernel32"));


				stream << " DLL_PROCESS_ATTACH PID = " << d;

				if (h)
				{
					stream << " Kernel32 was founded";
					addrW = (HOOKLoadLibraryW)::GetProcAddress(h, "LoadLibraryW");
					addrA = (HOOKLoadLibraryA)::GetProcAddress(h, "LoadLibraryA");
					bool res = detourLoadLibraryW.Hook(addrW, &hkLoadLibraryW, blackbone::HookType::Inline, blackbone::CallOrder::HookLast);
					//bool res1 = detourLoadLibraryA.Hook(addrA, &hkLoadLibraryA, blackbone::HookType::Inline, blackbone::CallOrder::HookLast);


				}
				else
				{
					stream << "  Kernel32 was founded";
				}
				st = stream.str();
				text.assign(st.begin(), st.end());
				OutputDebugString(text.c_str());
			}
		}
	
		break;
	
	//case DLL_THREAD_ATTACH:
    //case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
	break;

	}
    return TRUE;
}

