Attribute VB_Name = "modInjProc"
Option Explicit

Private Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long
 th32DefaultHeapID As Long
 th32ModuleID As Long
 cntThreads As Long
 th32ParentProcessID As Long
 pcPriClassBase As Long
 dwFlags As Long
 szExeFile As String * 260
End Type

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal fAllocType As Long, FlProtect As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal ProcessHandle As Long, lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Any, ByVal lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public ProsH As Long

Public Function InjectDll(DllPath As String, ProsH As Long)
On Error Resume Next
Dim DLLVirtLoc As Long, DllLength, Inject As Long, LibAddress As Long
Dim CreateThread As Long, ThreadID As Long
DllLength = Len(DllPath)
DLLVirtLoc = VirtualAllocEx(ProsH, ByVal &H0, DllLength, ByVal &H1000, ByVal &H4)
If DLLVirtLoc = 0 Then Exit Function
Inject = WriteProcessMemory(ProsH, DLLVirtLoc, ByVal DllPath, DllLength, vbNull)
LibAddress = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryA")
If LibAddress = 0 Then Exit Function
CreateThread = CreateRemoteThread(ProsH, vbNull, 0, LibAddress, DLLVirtLoc, 0, ThreadID)
Call WaitForSingleObject(CreateThread, &HFFFF)
If DLLVirtLoc <> 0 Then Call VirtualFreeEx(ProsH, DLLVirtLoc, 0, &H8000)
If LibAddress <> 0 Then Call CloseHandle(LibAddress)
If CreateThread <> 0 Then Call CloseHandle(CreateThread)
If ProsH <> 0 Then Call CloseHandle(ProsH)
End Function

Public Function GetHProcExe(strExeName As String) As Long
On Error Resume Next
Dim hSnap As Long
hSnap = CreateToolhelpSnapshot(2, 0)
Dim peProcess As PROCESSENTRY32
peProcess.dwSize = LenB(peProcess)
Dim nProcess As Long
nProcess = Process32First(hSnap, peProcess)
Do While nProcess
If StrComp(Trim(peProcess.szExeFile), strExeName, vbTextCompare) = 0 Then
GetHProcExe = OpenProcess(&H1F0FFF, False, peProcess.th32ProcessID)
Exit Function
End If
peProcess.szExeFile = vbNullString
nProcess = Process32Next(hSnap, peProcess)
Loop
CloseHandle hSnap
End Function

Public Function OpenApplication(ByVal StrAppPath As String) As Long
On Error Resume Next
Dim ReturnVal As Long
ReturnVal = ShellExecute(0&, "open", StrAppPath, "", "", vbNormalFocus)
If ReturnVal = 5& Then
OpenApplication = ShellExecute(0&, "runas", StrAppPath, "", "", vbNormalFocus)
Else
OpenApplication = ReturnVal
End If
End Function
