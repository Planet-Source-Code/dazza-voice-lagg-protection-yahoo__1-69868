struct VCHOOK
{
HMODULE hMod;
BYTE *pFunction;
BYTE *pTrampoline;
DWORD cbOriginal;
};

void Hook(struct VCHOOK *hk, LPCTSTR pszModule, LPCTSTR pszName, LPVOID pvNew);
void Unhook(struct VCHOOK *hk);