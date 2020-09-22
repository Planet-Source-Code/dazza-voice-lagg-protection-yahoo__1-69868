#include <winsock2.h>
#include "Util.h"

struct VCHOOK
	hookconnect;

typedef int (WSAAPI *Pconnect)
	(SOCKET s, struct sockaddr *name, int *namelen);

int WSAAPI myconnect(SOCKET s, struct sockaddr *name,  int *namelen)
#define SIN(name) ((struct sockaddr_in *)(name))
{
   struct sockaddr_in *paddr =(struct sockaddr_in *)name;
   char *ip =inet_ntoa(SIN(name)->sin_addr);
   int port =ntohs(paddr->sin_port);

if (port == 5001)
{
    SIN(name)->sin_addr.s_addr  = 0x0100007f;
	return ((Pconnect)hookconnect.pTrampoline)(s, name, namelen);
}

return ((Pconnect)hookconnect.pTrampoline)(s, name, namelen);
}

BOOL WINAPI _DllMainCRTStartup(HANDLE hDllHandle, DWORD dwReason, LPVOID lpReserved)
{
	switch(dwReason)
	{
	case DLL_PROCESS_ATTACH: {
		WSADATA wsaData;
    if(!WSAStartup(MAKEWORD( 2, 2 ) , &wsaData))
{
		Hook(&hookconnect, "WSOCK32.DLL", "connect", myconnect);
		}
	else
		{
	  //failed to load
		}
	return TRUE;
	}
	case DLL_PROCESS_DETACH:
		Unhook(&hookconnect);
		return TRUE;
	default:
		return TRUE;
	}
}
