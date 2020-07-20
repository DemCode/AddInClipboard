// dllmain.cpp : Defines the entry point for the DLL application.

#pragma setlocale("ru-RU" )

#include <windows.h>

#include <ComponentBase.h>
#include <AddInDefBase.h>
#include <IMemoryManager.h>

#include "CAddInClipboard.h"

#define DLLEXPORT __declspec(dllexport)

BOOL APIENTRY DllMain(HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}

const WCHAR_T* GetClassNames()
{
	return L"AddInClipboard";
}

long GetClassObject(const WCHAR_T* clsName, IComponentBase** pIntf)
{
	if (wcscmp(clsName, L"AddInClipboard") == 0)
	{
		*pIntf = new CAddInClipboard();

		return (long)*pIntf;
	}

	return 0;
}

long DestroyObject(IComponentBase** pIntf)
{
	if (!*pIntf)
		return -1;

	delete *pIntf;
	*pIntf = 0;
	return 0;
}

AppCapabilities SetPlatformCapabilities(const AppCapabilities capabilities)
{
	return eAppCapabilitiesLast;
}