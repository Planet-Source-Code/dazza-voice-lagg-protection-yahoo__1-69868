#include <windows.h>
#include "Util.h"

int ForgeHook(DWORD pAddr, DWORD pAddrToJump, BYTE **Buffer, DWORD *pBufSize);
void UnforgeHook(DWORD pAddr, BYTE *Buffer, DWORD OrigSize);

void Hook(struct VCHOOK *hk, LPCTSTR pszModule, LPCTSTR pszName, LPVOID pvNew)
{
	if(!(hk->hMod = LoadLibrary(pszModule)))
		return;
	if(!(hk->pFunction = (BYTE *)GetProcAddress(hk->hMod, pszName)))
	{
		FreeLibrary(hk->hMod);
		return;
	}
	if(!ForgeHook((DWORD)hk->pFunction, (DWORD)pvNew, &hk->pTrampoline, &hk->cbOriginal))
	{
		FreeLibrary(hk->hMod);
		return;
	}
}

void Unhook(struct VCHOOK *hk)
{
	UnforgeHook((DWORD)hk->pFunction, hk->pTrampoline, hk->cbOriginal);
	FreeLibrary(hk->hMod);
}
