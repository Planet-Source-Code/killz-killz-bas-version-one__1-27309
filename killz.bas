Attribute VB_Name = "killz"
'====================================================
'Sup, this is killz' bas file. It took me a long time to
'create and I got almost all the api functions from
'www.allapi.net's API Guide and API Toolshed. Its their
'for download so check it out. I use both aol95 and aol6.0
'Their aren't many aol functions here, but their are some
'useful ones. Im not a big aol programmer that much.
'I hope you find some use to this .bas file and dont
'be a lamer and rename this and say you coded it.
'
'
'                   Later,
'                         killz.
'====================================================


Option Explicit
Dim hMenu As Long
Dim PrevProc
Public Const MAX_PATH = 260
Public stopbust As Boolean
Public roombusted As Boolean


Public Enum KeyRoot
  [HKEY_CLASSES_ROOT] = &H80000000  'stores OLE class information and file associations
  [HKEY_CURRENT_CONFIG] = &H80000005 'stores computer configuration information
  [HKEY_CURRENT_USER] = &H80000001 'stores program information for the current user.
  [HKEY_LOCAL_MACHINE] = &H80000002 'stores program information for all users
  [HKEY_USERS] = &H80000003 'has all the information for any user (not just the one provided by HKEY_CURRENT_USER)
End Enum
Public Enum KeyType
  [REG_BINARY] = 3 'A non-text sequence of bytes
  [REG_DWORD] = 4 'A 32-bit integer...visual basic data type of Long
  [REG_SZ] = 1 'A string terminated by a null character
End Enum

Public Const KEY_ALL_ACCESS = &HF003F 'Permission for all types of access.
Public Const KEY_ENUMERATE_SUB_KEYS = &H8 'Permission to enumerate subkeys.
Public Const KEY_READ = &H20019 'Permission for general read access.
Public Const KEY_WRITE = &H20006 'Permission for general write access.
Public Const KEY_QUERY_VALUE = &H1 'Permission to query subkey data.
' used for import/export registry key
Public Const REG_FORCE_RESTORE As Long = 8& 'Permission to overwrite a registry key
Public Const TOKEN_QUERY As Long = &H8&
Public Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Public Const SE_PRIVILEGE_ENABLED As Long = &H2
Public Const SE_RESTORE_NAME = "SeRestorePrivilege" 'Important for what we're trying to accomplish
Public Const SE_BACKUP_NAME = "SeBackupPrivilege"
Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
' used for enumerating registrykeys
Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
' used for import/export registry key
Public Type LUID
  lowpart As Long
  highpart As Long
End Type
Public Type LUID_AND_ATTRIBUTES
  pLuid As LUID
  Attributes As Long
End Type
Public Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges As LUID_AND_ATTRIBUTES
End Type



Public Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    Style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type

Public Enum dwRop

    WHITENESS = &HFF0062
    BLACKNESS = &H42
    SRCAND = &H8800C6
    SRCCOPY = &HCC0020
    SRCINVERT = &H660046
    SRCERASE = &H440328
    SRCPAINT = &HEE0086
    
End Enum


Type WIN32_FIND_DATA ' 318 Bytes
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved_ As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
    End Type
    
    Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Const FLAG_ICC_FORCE_CONNECTION = &H1
Public Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptGetProvParam Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Public Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Public Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Public Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Declare Function RegisterDLL Lib "Regist10.dll" Alias "REGISTERDLL" (ByVal DllPath As String, bRegister As Boolean) As Boolean
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
Public Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function FindFirstFile& Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA)
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Public Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Public Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal Dest&, ByVal nCount&)
Public Declare Function dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal Dest As Long, ByVal nCount&)
Public Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Public Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Public Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Public Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function WriteFile Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function DiskSpaceFree Lib "STKIT432.DLL" () As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef source As Any, ByVal nBytes As Long)
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal length As Long)
Public Declare Function GetAllUsersProfileDirectory Lib "userenv.dll" Alias "GetAllUsersProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Public Declare Function GetDefaultUserProfileDirectory Lib "userenv.dll" Alias "GetDefaultUserProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Public Declare Function GetProfilesDirectory Lib "userenv.dll" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Public Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.



'Encryption Const
Public Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
Public Const PROV_RSA_FULL As Long = 1
Public Const PP_NAME As Long = 4
Public Const PP_CONTAINER As Long = 6
Public Const CRYPT_NEWKEYSET As Long = 8
Public Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Public Const ALG_CLASS_HASH As Long = 32768
Public Const ALG_TYPE_ANY As Long = 0
Public Const ALG_TYPE_STREAM As Long = 2048
Public Const ALG_SID_RC4 As Long = 1
Public Const ALG_SID_MD5 As Long = 3
Public Const CALG_MD5 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Public Const CALG_RC4 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Public Const ENCRYPT_ALGORITHM As Long = CALG_RC4
Public Const ENCRYPT_NUMBERKEY As String = "16006833"
Public lngCryptProvider As Long
Public avarSeedValues As Variant
Public lngSeedLevel As Long
Public lngDecryptPointer As Long
Public astrEncryptionKey(0 To 131) As String
Public Const lngALPKeyLength As Long = 8
Public strKeyContainer As String
'My Constants
Public Const WM_GETCHATTEXT = 14
Public Const PL_GETCERTAIN = 13
' Color constants
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
' ExWindowStyles
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&
' Window styles
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_DBLCLKS = &H8
Public Const CS_OWNDC = &H20
Public Const CS_CLASSDC = &H40
Public Const CS_PARENTDC = &H80
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOCLOSE = &H200
Public Const CS_SAVEBITS = &H800
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_PUBLICCLASS = &H4000
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const CB_DELETESTRING = &H144
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_RESETCONTENT = &H14B
Public Const GFSR_SYSTEMRESOURCES = 0
Public Const GFSR_GDIRESOURCES = 1
Public Const GFSR_USERRESOURCES = 2
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_SIZE = &H5
Public Const WM_PASTE = &H302
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const conMCIAppTitle = "MCI Control Application"
Public Const conMCIErrInvalidDeviceID = 30257
Public Const conMCIErrDeviceOpen = 30263
Public Const conMCIErrCannotLoadDriver = 30266
Public Const conMCIErrUnsupportedFunction = 30274
Public Const conMCIErrInvalidFile = 30304
Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
Public Const FADE_BLUE = &HFF0000
Public Const FADE_YELLOW = &HFFFF&
Public Const FADE_WHITE = &HFFFFFF
Public Const FADE_BLACK = &H0&
Public Const FADE_PURPLE = &HFF00FF
Public Const FADE_GREY = &HC0C0C0
Public Const FADE_PINK = &HFF80FF
Public Const FADE_TURQUOISE = &HC0C000
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &HF012
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
Public Const LB_MULTIPLEADDSTRING = &H1B1
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181
Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9
Public Const VK_UP = &H26
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNA = 8
Public Const SW_MAX = 10
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const WM_SYSCOMMAND = &H112
Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Const MB_DEFBUTTON1 = &H0&
Const MB_DEFBUTTON2 = &H100&
Const MB_DEFBUTTON3 = &H200&
Const MB_ICONASTERISK = &H40&
Const MB_ICONEXCLAMATION = &H30&
Const MB_ICONHAND = &H10&
Const MB_ICONINFORMATION = MB_ICONASTERISK
Const MB_ICONQUESTION = &H20&
Const MB_ICONSTOP = MB_ICONHAND
Const MB_OK = &H0&
Const MB_OKCANCEL = &H1&
Const MB_YESNO = &H4&
Const MB_YESNOCANCEL = &H3&
Const MB_ABORTRETRYIGNORE = &H2&
Const MB_RETRYCANCEL = &H5&
' Standard ID's of cursors
Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_SIZE = 32640&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_NO = 32648&
Public Const IDC_APPSTARTING = 32650&
Public Const GWL_WNDPROC = -4




Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer

    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer

    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Dim DevM As DEVMODE

Type COLORRGB
  red As Long
  Green As Long
  blue As Long
End Type

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   x As Long
   y As Long
End Type

Public DialogCaption As String


Function FileFound(strFileName As String) As Boolean
    'Code Created by Lucian
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim hFindFirst As Long
    hFindFirst = FindFirstFile(strFileName, lpFindFileData)


    If hFindFirst > 0 Then
        FindClose hFindFirst
        FileFound = True
    Else
        FileFound = False
    End If
End Function


Public Sub Form_Center(f As Form)
    f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
    f.Left = Screen.Width / 2 - f.Width / 2
End Sub


Public Function BlankString() As String
    BlankString$ = Chr(32) & Chr(160)
End Function

Function GetClassNameNow(Ret As String)
Dim winwnd As Long
Dim lpClassName As String
Dim retval As Long
    winwnd = FindWindow(vbNullString, UCase(Ret$))
    If winwnd = 0 Then MsgBox "Couldn't find the window ...": Exit Function
    lpClassName = Space(256)
    retval = GetClassName(winwnd, lpClassName, 256)
    GetClassNameNow = Left$(lpClassName, retval)
End Function

Public Function MakeIt3d(TheForm As Form, TheControl As Control)
Dim OldMode As Long
If TheForm.AutoRedraw = False Then
    OldMode = TheForm.ScaleMode
        TheForm.ScaleMode = 3
        TheForm.AutoRedraw = True
        TheForm.CurrentX = TheControl.Left - 1
        TheForm.CurrentY = TheControl.Top + TheControl.Height
        TheForm.Line -Step(0, -(TheControl.Height + 1)), RGB(90, 90, 90)
        TheForm.Line -Step(TheControl.Width + 1, 0), RGB(90, 90, 90)
        TheForm.Line -Step(0, TheControl.Height + 1), RGB(255, 255, 255)
        TheForm.Line -Step(-(TheControl.Width + 1), 0), RGB(255, 255, 255)
        TheForm.AutoRedraw = False
    TheForm.ScaleMode = OldMode
End If
If TheForm.AutoRedraw = True Then
    OldMode = TheForm.ScaleMode
        TheForm.ScaleMode = 3
        TheForm.CurrentX = TheControl.Left - 1
        TheForm.CurrentY = TheControl.Top + TheControl.Height
        TheForm.Line -Step(0, -(TheControl.Height + 1)), RGB(90, 90, 90)
        TheForm.Line -Step(TheControl.Width + 1, 0), RGB(90, 90, 90)
        TheForm.Line -Step(0, TheControl.Height + 1), RGB(255, 255, 255)
        TheForm.Line -Step(-(TheControl.Width + 1), 0), RGB(255, 255, 255)
    TheForm.ScaleMode = OldMode
End If
End Function


Public Sub Window_Enable(window)
    Call EnableWindow(window, 1)
End Sub





Public Sub RemoveItem_Combo(ComboWin As Long, TheString As String)
Dim FindIt As Long, DeleteIt As Long
FindIt = SendMessageByString(ComboWin, CB_FINDSTRINGEXACT, -1, TheString)
If FindIt <> -1 Then
    Call SendMessageByString(ComboWin, CB_DELETESTRING, FindIt, 0)
End If
End Sub
Public Sub RemoveItem_ListBoX(ListWin, TheString)
Dim FindIt As Long, DeleteIt As Long
FindIt = SendMessageByString(ListWin, LB_FINDSTRINGEXACT, -1, TheString)
If FindIt <> -1 Then
    Call SendMessageByString(ListWin, LB_DELETESTRING, FindIt, 0)
End If
End Sub
Public Sub Draw3DBorder(C As Control, iLook As Integer)
'Makes A Control Look 3D
Dim iOldScaleMode As Integer, iFirstColor As Integer
Dim iSecondColor As Integer, RAISED As Variant, PIXELS As Variant
    If iLook = RAISED Then
        iFirstColor = 15
        iSecondColor = 8
    Else
        iFirstColor = 8
        iSecondColor = 15
    End If
iOldScaleMode = C.Parent.ScaleMode
C.Parent.ScaleMode = PIXELS
C.Parent.Line (C.Left, C.Top - 1)-(C.Left + C.Width, C.Top - 1), QBColor(iFirstColor)
C.Parent.Line (C.Left - 1, C.Top)-(C.Left - 1, C.Top + C.Height), QBColor(iFirstColor)
C.Parent.Line (C.Left + C.Width, C.Top)-(C.Left + C.Width, C.Top + C.Height), QBColor(iSecondColor)
C.Parent.Line (C.Left, C.Top + C.Height)-(C.Left + C.Width, C.Top + C.Height), QBColor(iSecondColor)
C.Parent.ScaleMode = iOldScaleMode
End Sub
Public Sub WriteToLog(what As String, LoGPath As String)
Dim x As Long, sSTR As String
If LoGPath = "" Then Exit Sub
If InStr(LoGPath, ".") = 0 Then Exit Sub
x& = FreeFile
Open LoGPath For Binary Access Write As x&
    sSTR$ = what & Chr(10)
    Put #1, LOF(1) + 1, sSTR$
Close x&
End Sub
Public Function WindowSPYLabels(WinHdl, WinClass, WinTxT, WinStyle, WinIDNum, WinPHandle, WinPText, WinPClass, WinModule)
'Call This In A Timer
Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
Dim sClassName As String * 100, hWndOver As Long, hWndParent As Long
Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
Dim hInstance As Long, sParentWindowText As String * 100
Dim sModuleFileName As String * 100, r As Long
Static hWndLast As Long
    Call GetCursorPos(pt32)
    ptx = pt32.x
    pty = pt32.y
    hWndOver = WindowFromPointXY(ptx, pty)
    If hWndOver <> hWndLast Then
        hWndLast = hWndOver
        WinHdl.Caption = "Window Handle: " & hWndOver
        sWindowText = Space(100)
        r = GetWindowText(hWndOver, sWindowText, 100)
        WinTxT.Caption = "Window Text: " & Left(sWindowText, r)
        sClassName = Space(100)
        r = GetClassName(hWndOver, sClassName, 100)
        WinClass.Caption = "Window Class Name: " & Left(sClassName, r)
        lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)
        WinStyle.Caption = "Window Style: " & lWindowStyle
        hWndParent = GetParent(hWndOver)
            If hWndParent <> 0 Then
                wID = GetWindowWord(hWndOver, GWW_ID)
                WinIDNum.Caption = "Window ID Number: " & wID
                WinPHandle.Caption = "Parent Window Handle: " & hWndParent
                sParentWindowText = Space(100)
                r = GetWindowText(hWndParent, sParentWindowText, 100)
                WinPText.Caption = "Parent Window Text: " & Left(sParentWindowText, r)
                sParentClassName = Space(100)
                r = GetClassName(hWndParent, sParentClassName, 100)
                WinPClass.Caption = "Parent Window Class Name: " & Left(sParentClassName, r)
            Else
                WinIDNum.Caption = "Window ID Number: Not Available"
                WinPHandle.Caption = "Parent Window Handle: Not Available"
                WinPText.Caption = "Parent Window Text : Not Available"
                WinPClass.Caption = "Parent Window Class Name: Not Available"
            End If
                hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
                sModuleFileName = Space(100)
                r = GetModuleFileName(hInstance, sModuleFileName, 100)
        WinModule.Caption = "Module: " & Left(sModuleFileName, r)
    End If
End Function

Public Function Click_List(window, index)
    Call SendMessage(window, LB_SETCURSEL, ByVal CLng(index), ByVal 0&)
End Function
Public Function TileBitmap(TheForm As Form, theBitmap As PictureBox)
Dim Across As Integer, Down As Integer
theBitmap.AutoSize = True
    For Down = 0 To (TheForm.Width \ theBitmap.Width) + 1
        For Across = 0 To (TheForm.Height \ theBitmap.Height) + 1
            TheForm.PaintPicture theBitmap.Picture, Down * theBitmap.Width, Across * theBitmap.Height, theBitmap.Width, theBitmap.Height
    Next Across, Down
End Function
Public Sub Window_Maximize(window)
    Call ShowWindow(window, SW_MAXIMIZE)
End Sub
Public Sub Window_Minimize(window)
    Call ShowWindow(window, SW_MINIMIZE)
End Sub
Public Function MakeASCIIChart(List As ListBox)
Dim x As Long
For x = 33 To 255
    List.AddItem Chr(x)
Next x
End Function
Public Function WindowSPYTextBoxs(WinHdl As TextBox, WinClass As TextBox, WinTxT As TextBox, WinStyle As TextBox, WinIDNum As TextBox, WinPHandle As TextBox, WinPText As TextBox, WinPClass As TextBox, WinModule As TextBox)
'Call This In A Timer
Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
Dim sClassName As String * 100, hWndOver As Long, hWndParent As Long
Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
Dim hInstance As Long, sParentWindowText As String * 100
Dim sModuleFileName As String * 100, r As Long
Static hWndLast As Long
    Call GetCursorPos(pt32)
    ptx = pt32.x
    pty = pt32.y
    hWndOver = WindowFromPointXY(ptx, pty)
    If hWndOver <> hWndLast Then
        hWndLast = hWndOver
        WinHdl.text = "Window Handle: " & hWndOver
        r = GetWindowText(hWndOver, sWindowText, 100)
        WinTxT.text = "Window Text: " & Left(sWindowText, r)
        r = GetClassName(hWndOver, sClassName, 100)
        WinClass.text = "Window Class Name: " & Left(sClassName, r)
        lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)
        WinStyle.text = "Window Style: " & lWindowStyle
        hWndParent = GetParent(hWndOver)
            If hWndParent <> 0 Then
                wID = GetWindowWord(hWndOver, GWW_ID)
                WinIDNum.text = "Window ID Number: " & wID
                WinPHandle.text = "Parent Window Handle: " & hWndParent
                r = GetWindowText(hWndParent, sParentWindowText, 100)
                WinPText.text = "Parent Window Text: " & Left(sParentWindowText, r)
                r = GetClassName(hWndParent, sParentClassName, 100)
                WinPClass.text = "Parent Window Class Name: " & Left(sParentClassName, r)
            Else
                WinIDNum.text = "Window ID Number: N/A"
                WinPHandle.text = "Parent Window Handle: N/A"
                WinPText.text = "Parent Window Text : N/A"
                WinPClass.text = "Parent Window Class Name: N/A"
            End If
                hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
                r = GetModuleFileName(hInstance, sModuleFileName, 100)
        WinModule.text = "Module: " & Left(sModuleFileName, r)
    End If
End Function

Public Sub ExtractAnIcon(CmmDlg As Control)
Dim sSourcePgm As String, lIcon As Long

Dim a%
    On Error Resume Next
  With CmmDlg
    .FileName = sSourcePgm
    .CancelError = True
    .DialogTitle = "Select a DLL or EXE which includes Icons"
    .Filter = "Icon Resources (*.ico;*.exe;*.dll)|*.ico;*.exe;*.dll|All files|*.*"
    .Action = 1
    If Err Then
      Err.Clear
      Exit Sub
    End If
    sSourcePgm = .FileName
    DestroyIcon lIcon
    End With
    Do
      lIcon = ExtractIcon(App.hInstance, sSourcePgm, a)
      If lIcon = 0 Then Exit Do
      a = a + 1
      DestroyIcon lIcon
    Loop
    If a = 0 Then
      MsgBox "No Icons in this file!"
    End If
End Sub




Public Sub Click(icon)
    Call SendMessage(icon, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(icon, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MIDI_Play(Midi As String)
Dim FilE As String
FilE$ = Dir(Midi$)
If FilE$ <> "" Then
    Call mciSendString("play " & Midi$, 0&, 0, 0)
End If
End Sub
Public Sub MIDI_Stop(Midi As String)
Dim FilE As String
FilE$ = Dir(Midi$)
If FilE$ <> "" Then
    Call mciSendString("stop " & Midi$, 0&, 0, 0)
End If
End Sub

Sub Click_Double(icon&)
    Call SendMessageByNum(icon&, WM_LBUTTONDBLCLK, &HD, 0)
End Sub


Public Function FindChildByTitle(Parent As Long, child As String)
    FindChildByTitle = FindWindowEx(Parent, 0&, vbNullString, child)
End Function



Sub Click_StartButton()
Dim Windows As Long, StartButton As Long
Windows& = FindWindow("Shell_TrayWnd", vbNullString)
StartButton& = FindWindowEx(Windows&, 0&, "Button", vbNullString)
Click (StartButton&)
End Sub


Public Sub Window_Hide(window As Long)
    Call ShowWindow(window, 0)
End Sub

Public Sub Window_Show(window As Long)
    Call ShowWindow(window, 5)
End Sub
Public Sub StayOffTop(f As Form)
    Call SetWindowPos(f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Sub DecompileProtect(ExeLocation)
Dim ThaFile As String, Cat As String
On Error Resume Next
    If ExeLocation = "" Then MsgBox "Executable File Not Found", vbOKOnly
ThaFile = FreeFile
Open ExeLocation For Binary As #ThaFile
    Cat = "."
Seek #ThaFile, 25
Put #ThaFile, , Cat
Close #1
If Err Then MsgBox "Not A Visual Basic Made File!", vbOKOnly, "Error In File": Exit Sub
MsgBox "Youre File Has Been Protected", vbOKOnly
End Sub

Public Function ClearDocuments()
    Call SHAddToRecentDocs(0, 0)
End Function

Public Function FindChildByClass(Parent, child)
    FindChildByClass = FindWindowEx(Parent, 0&, child, vbNullString)
End Function




Public Sub File_Delete(FilE$)
Dim NoFreeze As Long
If Not File_Exists(FilE$) Then Exit Sub
Kill FilE$
NoFreeze& = DoEvents()
End Sub


Public Sub DeleteListItem(List As ListBox, item$)

    item$ = List.ListIndex
    List.RemoveItem (item$)
End Sub


Public Function DirExists(TheDir)
Dim Test As Integer
On Error Resume Next
    If Right(TheDir, 1) <> "/" Then TheDir = TheDir & "/"
Test = Len(Dir$(TheDir))
If Err Or Test = 0 Then DirExists = False: Exit Function
DirExists = True
End Function
Public Function File_Exists(ByVal FileName As String) As Integer
Dim Test As Integer
On Error Resume Next
    Test = Len(Dir$(FileName))
If Err Or Test = 0 Then File_Exists = False: Exit Function
File_Exists = True
End Function



Public Function File_GetAttributes(TheFile As String)
Dim FilE As String
    FilE = Dir(TheFile)
If FilE <> "" Then File_GetAttributes = GetAttr(TheFile)
End Function
Public Sub File_SetHidden(TheFile As String)
Dim FilE As String
    FilE = Dir(TheFile)
If FilE <> "" Then SetAttr TheFile, vbHidden
End Sub

Public Sub File_SetReadOnly(TheFile As String)
Dim FilE As String
    FilE = Dir(TheFile)
If FilE <> "" Then SetAttr TheFile, vbReadOnly
End Sub


Public Sub LoadFonts(List As Control)
Dim x As Long
List.Clear
For x = 1 To Screen.FontCount
    List.AddItem Screen.Fonts(x - 1)
Next
End Sub
Public Function GetClass(child&) As String
Dim sString As String, Plop As String
sString$ = String$(250, 0)
    GetClass = GetClassName(child, sString$, 250)
    GetClass = sString$
End Function
Public Function GetCaption(window)
Dim windowtitle As String, WindowText As String, WindowLength As Long
WindowLength& = GetWindowTextLength(window)
    windowtitle$ = String$(WindowLength&, 0)
    WindowText$ = GetWindowText(window, windowtitle$, (WindowLength& + 1))
    GetCaption = windowtitle$
End Function

Public Function GetText(child)
Dim TheTrimmer As Long, TrmSpace As String, GetStr As Long
TheTrimmer& = SendMessageByNum(child, WM_GETCHATTEXT, 0&, 0&)
    TrmSpace$ = Space$(TheTrimmer)
GetStr = SendMessageByString(child, PL_GETCERTAIN, TheTrimmer + 1, TrmSpace$)
    GetText = TrmSpace$
End Function



Public Function TaskBar_Hide()
Dim Bar As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Call ShowWindow(Bar&, 0)
End Function
Public Function TaskBar_Show()
Dim Bar As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
    Call ShowWindow(Bar&, 5)
End Function
Public Function StartButton_Hide()
Dim Bar As Long, Button As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(Bar&, 0&, "Button", vbNullString)
Call ShowWindow(Button&, 0)
End Function
Public Function StartButton_Show()
Dim Bar As Long, Button As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(Bar&, 0&, "Button", vbNullString)
    Call ShowWindow(Button&, 5)
End Function




Public Sub Window_Close(window)
    Call SendMessageByNum(window, WM_CLOSE, 0, 0)
End Sub

Public Sub CenterForm(f As Form)
    f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
    f.Left = Screen.Width / 2 - f.Width / 2
End Sub

Private Sub ListBox2Clipboard(List As ListBox)
Dim sn As Long, thelist As String
For sn = 0 To List.ListCount - 1
If sn = 0 Then
    thelist = List.List(sn)
Else
    thelist = thelist & "," & List.List(sn)
End If
Next
Clipboard.Clear
TimeOut 0.1
Clipboard.SetText thelist
End Sub



Public Sub RunMenuByString(window, StringSearch)
Dim FindWin As Long, CountMenu As Long, FindString As Long, MenuItem As Long
Dim FindWinSub As Long, MenuItemCount As Long, getstring As Long
Dim SubCount As Long, MenuString As String, GetStringMenu As Long
FindWin& = GetMenu(window)
CountMenu& = GetMenuItemCount(FindWin&)

For FindString = 0 To CountMenu& - 1
    FindWinSub& = GetSubMenu(FindWin&, FindString)
    MenuItemCount& = GetMenuItemCount(FindWinSub&)
For getstring = 0 To MenuItemCount& - 1
    SubCount& = GetMenuItemID(FindWinSub&, getstring)
    MenuString$ = String$(100, " ")
    GetStringMenu& = GetMenuString(FindWinSub&, SubCount&, MenuString$, 100, 1)
If InStr(UCase(MenuString$), UCase(StringSearch)) Then
    MenuItem& = SubCount&
    GoTo MatchString
End If
Next getstring
Next FindString

MatchString:
    Call SendMessage(window, WM_COMMAND, MenuItem&, 0)
End Sub


Public Sub MakeShortcut(ShortcutDir, ShortcutName, ShortcutPath)
Dim WinShortcutDir As String, WinShortcutName As String, WinShortcutExePath As String, retval As Long
    WinShortcutDir$ = ShortcutDir
    WinShortcutName$ = ShortcutName
    WinShortcutExePath$ = ShortcutPath
retval& = fCreateShellLink("", WinShortcutName$, WinShortcutExePath$, "")
    Name "C:\Windows\Start Menu\Programs\" & WinShortcutName$ & ".LNK" As WinShortcutDir$ & "\" & WinShortcutName$ & ".LNK"
End Sub


Public Sub ParentChange(frm As Form, window&)
    Call SetParent(frm.hwnd, window&)
End Sub


Public Function ReadINI(Header As String, Key As String, location As String) As String
Dim sString As String
    sString = String(750, Chr(0))
    Key$ = LCase$(Key$)
    ReadINI$ = Left(sString, GetPrivateProfileString(Header$, ByVal Key$, "", sString, Len(sString), location$))
End Function

Public Sub File_ReName(FilE$, NewName$)
Dim NoFreeze As Long
    Name FilE$ As NewName$
    NoFreeze& = DoEvents()
End Sub



Public Sub RunMenu(menu1 As Integer, menu2 As Integer)
Static Working As Integer
Dim Menus As Long, SubMenu As Long, ItemID As Long, Works As Long, MenuClick As Long
Menus& = GetMenu(FindWindow("AOL Frame25", vbNullString))
SubMenu& = GetSubMenu(Menus&, menu1)
ItemID = GetMenuItemID(SubMenu&, menu2)
Works = CLng(0) * &H10000 Or Working
MenuClick = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, ItemID, 0&)
End Sub

Public Sub Window_SetText(window, text)
    Call SendMessageByString(window, WM_SETTEXT, 0, text)
End Sub

Public Sub shutdownwindows()
Dim EWX_SHUTDOWN
    Dim MsgRes As Long
    MsgRes = MsgBox("Do you really want to Shut Down Windows 9x", vbYesNo Or vbQuestion)
    If MsgRes = vbNo Then Exit Sub
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End Sub


Public Sub StayOnTop(TheForm As Form)
    Call SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub



Public Function StringInList(thelist As ListBox, FindMe As String)
Dim a As Long
If thelist.ListCount = 0 Then GoTo ListEmpty
For a = 0 To thelist.ListCount - 1
thelist.ListIndex = a
    If UCase(thelist.text) = UCase(FindMe) Then
        StringInList = a
    Exit Function
    End If
Next a
ListEmpty:
StringInList = -1
End Function






Public Sub TimeOut(length)
    Dim begin As Long
    begin = Timer
Do While Timer - begin >= length
    DoEvents
Loop
End Sub
Public Sub Pause(length)
'Same As Timeout
    Dim begin As Long
    begin = Timer
Do While Timer - begin >= length
    DoEvents
Loop
End Sub




Public Sub waitforok()
Dim waitforok As Long, OK As Long, OKButton As Long
Do
    DoEvents
    OK = FindWindow("#32770", "America Online")
    DoEvents
Loop Until OK <> 0
OKButton = FindWindowEx(OK, 0&, vbNullString, "OK")
    Call SendMessageByNum(OKButton, WM_LBUTTONDOWN, 0, 0&)
    Call SendMessageByNum(OKButton, WM_LBUTTONUP, 0, 0&)
End Sub

Public Sub WriteToINI(Header As String, Key As String, KeyValue As String, location As String)
    Call WritePrivateProfileString(Header$, UCase$(Key$), KeyValue$, location$)
End Sub
Public Function Form_Drag(Form As Form)
'This Goes In Mouse Down Events Of A Label/Button
    Call ReleaseCapture
    Call SendMessage(Form.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Function


Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Long
    ZeroPos = InStr(1, sInput, Chr$(0))
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

Function GetWinDir()
    Dim sSave As String, Ret As Long
    sSave = Space(255)
    Ret = GetWindowsDirectory(sSave, 255)
    sSave = Left$(sSave, Ret)
    GetWinDir = sSave
End Function

Function GetProfilesDir(who)
Dim dirst
Dim ttt
dirst = GetWinDir()
ttt = InStr(4, dirst, "\")
If ttt <> 0 Then
If FileFound(GetWinDir() & "Profiles\" & who) = False Then GetProfilesDir = False: MsgBox "That profiles member does not exist": Exit Function
GetProfilesDir = dirst & "profiles\" & who
ElseIf ttt = 0 Then
GetProfilesDir = dirst & "\profiles\" & who
End If
End Function

Function GetShortPath(strng As String)
Dim txt$
Dim ttt&
txt$ = String(165, 0)
ttt& = GetShortPathName(strng$, txt$, 165)
GetShortPath = txt$
End Function

Function RandomWinPos(win As Long, x As String, y As String, wx2 As String, wy2 As String)
Randomize
x = SetWindowPos(win&, HWND_TOPMOST, x * Rnd, y * Rnd, wx2 * Rnd, wy2 * Rnd, &H40)
End Function

Function RandomCursorPos(x As String, y As String)
Randomize
x = SetCursorPos(x * Rnd, y * Rnd)
End Function

Function RunAOLToolbar(MenuNumber As String, letter As String)
Dim aolframe&
aolframe& = FindWindow("AOL Frame25", vbNullString)
Dim aoltoolbar&
aoltoolbar& = FindWindowEx(aolframe&, 0&, "AOL Toolbar", vbNullString)
Dim aoltoolbar2
aoltoolbar2 = FindWindowEx(aoltoolbar&, 0&, "_AOL_Toolbar", vbNullString)
Dim aolicon
aolicon = FindWindowEx(aoltoolbar2, 0&, "_AOL_Icon", vbNullString)
Dim Count
For Count = 1 To MenuNumber
aolicon = FindWindowEx(aoltoolbar2, aolicon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(aolicon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(aolicon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
Dim menu
menu = FindWindow("#32768", vbNullString)
Dim found
found = IsWindowVisible(menu)
Loop Until found <> 0
letter = Asc(letter)
Call PostMessage(menu, WM_CHAR, letter, 0&)
End Function



Public Function FindChatRoom() As Long
Dim Counter As Long
Dim AOLStatic5 As Long
Dim AOLIcon3 As Long
Dim AOLStatic4 As Long
Dim aollistbox As Long
Dim AOLStatic3 As Long
Dim aolimage As Long
Dim AOLIcon2 As Long
Dim RICHCNTL2 As Long
Dim AOLStatic2 As Long
Dim i As Long
Dim aolicon As Long
Dim AOLCombobox As Long
Dim richcntl As Long
Dim aolstatic As Long
Dim aolchild As Long
Dim mdiclient As Long
Dim aolframe As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "AOL Child", vbNullString)
aolstatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
richcntl& = FindWindowEx(aolchild&, 0&, "RICHCNTL", vbNullString)
AOLCombobox& = FindWindowEx(aolchild&, 0&, "_AOL_Combobox", vbNullString)
aolicon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    aolicon& = FindWindowEx(aolchild&, aolicon&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(aolchild&, aolstatic&, "_AOL_Static", vbNullString)
RICHCNTL2& = FindWindowEx(aolchild&, richcntl&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(aolchild&, aolicon&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon2& = FindWindowEx(aolchild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
aolimage& = FindWindowEx(aolchild&, 0&, "_AOL_Image", vbNullString)
AOLStatic3& = FindWindowEx(aolchild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLStatic3& = FindWindowEx(aolchild&, AOLStatic3&, "_AOL_Static", vbNullString)
aollistbox& = FindWindowEx(aolchild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic4& = FindWindowEx(aolchild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(aolchild&, AOLIcon2&, "_AOL_Icon", vbNullString)
For i& = 1& To 7&
    AOLIcon3& = FindWindowEx(aolchild&, AOLIcon3&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic5& = FindWindowEx(aolchild&, AOLStatic4&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (aolstatic& = 0& Or richcntl& = 0& Or AOLCombobox& = 0& Or aolicon& = 0& Or AOLStatic2& = 0& Or RICHCNTL2& = 0& Or AOLIcon2& = 0& Or aolimage& = 0& Or AOLStatic3& = 0& Or aollistbox& = 0& Or AOLStatic4& = 0& Or AOLIcon3& = 0& Or AOLStatic5& = 0&): DoEvents
    aolchild& = FindWindowEx(mdiclient&, aolchild&, "AOL Child", vbNullString)
    aolstatic& = FindWindowEx(aolchild&, 0&, "_AOL_Static", vbNullString)
    richcntl& = FindWindowEx(aolchild&, 0&, "RICHCNTL", vbNullString)
    AOLCombobox& = FindWindowEx(aolchild&, 0&, "_AOL_Combobox", vbNullString)
    aolicon& = FindWindowEx(aolchild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        aolicon& = FindWindowEx(aolchild&, aolicon&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(aolchild&, aolstatic&, "_AOL_Static", vbNullString)
    RICHCNTL2& = FindWindowEx(aolchild&, richcntl&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(aolchild&, aolicon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon2& = FindWindowEx(aolchild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    aolimage& = FindWindowEx(aolchild&, 0&, "_AOL_Image", vbNullString)
    AOLStatic3& = FindWindowEx(aolchild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLStatic3& = FindWindowEx(aolchild&, AOLStatic3&, "_AOL_Static", vbNullString)
    aollistbox& = FindWindowEx(aolchild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic4& = FindWindowEx(aolchild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(aolchild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    For i& = 1& To 7&
        AOLIcon3& = FindWindowEx(aolchild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic5& = FindWindowEx(aolchild&, AOLStatic4&, "_AOL_Static", vbNullString)
    If aolstatic& And richcntl& And AOLCombobox& And aolicon& And AOLStatic2& And RICHCNTL2& And AOLIcon2& And aolimage& And AOLStatic3& And aollistbox& And AOLStatic4& And AOLIcon3& And AOLStatic5& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindChatRoom& = aolchild&
    Exit Function
End If
End Function
Function SecsToMins(Secs As Integer)
    If Secs < 60 Then SecsToMins = "00:" & Format(Secs, "00") Else SecsToMins = Format(Secs / 60, "00") & ":" & Format(Secs - Format(Secs / 60, "00") * 60, "00")
End Function
Function FindToolbar2() As Long
Dim aol&, tool1&, tool2&
aol& = FindWindow("AOL Frame25", vbNullString)
tool1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
tool2& = FindWindowEx(tool1&, 0&, "_AOL_Toolbar", vbNullString)
FindToolbar2& = tool2&
End Function

Function FindAOLChild() As Long
Dim aol&, mdi&, child&
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindChildByClass(aol&, "MDIClient")
child& = FindChildByClass(mdi&, "AOL Child")
FindAOLChild& = child&
End Function

Function ClickToolbar(icon As Long)
Call SendMessage(icon, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(icon, WM_KEYUP, VK_SPACE, 0)
End Function


Function ClickReadMail()
Dim toolbar&, icon1&, icon2&
toolbar& = FindToolbar2()
icon1& = FindChildByClass(toolbar&, "_AOL_Icon")
icon2& = GetWindow(icon1&, 2)
Call ClickToolbar(icon2&)
End Function

Function GetSN() As String
Dim child&, txt$, sn$, scn$, x
child& = FindAOLChild()
Do
DoEvents
txt$ = GetText(child&)
If InStr(txt$, "Welcome, ") Then
x = InStr(txt$, " ")
sn$ = Mid(txt$, x + 1, Len(txt$))
scn$ = Mid(sn$, 1, Len(sn$) - 1)
Exit Do
End If
child& = GetWindow(child&, 2)
Loop
GetSN$ = scn$
End Function


Function Find30Chat() As Long
'_AOL_Static, _AOL_View, _AOL_Edit, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Image, _AOL_Static, _AOL_Static, _AOL_Listbox, _AOL_Static, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Icon, _AOL_Static
Dim child&, staticc&, what, i, view&, edit&, icon&, aolimage&, List&
child& = FindAOLChild()
what = GetChildrenNum()
For i = 1 To what
staticc& = FindChildByClass(child&, "_AOL_Static")
view& = FindChildByClass(child&, "_AOL_View")
edit& = FindChildByClass(child&, "_AOL_Edit")
icon& = FindChildByClass(child&, "_AOL_Icon")
aolimage& = FindChildByClass(child&, "_AOL_Image")
List& = FindChildByClass(child&, "_AOL_Listbox")
If staticc& <> 0 And view& <> 0 And edit& <> 0 And icon& <> 0 And aolimage& <> 0 And List& <> 0 Then
Find30Chat& = child&
Exit Function
Else
child& = GetWindow(child&, 2)
End If
Next i
End Function

Function GetChildrenNum()
Dim child&, num
child& = FindAOLChild()
If child& <> 0 Then num = num + 1
While child&
DoEvents
child& = GetWindow(child&, 2)
If child& <> 0 Then num = num + 1
Wend
GetChildrenNum = num
End Function

Function Add30Room(thelist As ListBox, adduser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    room& = Find30Chat&
    If room& = 0& Then Exit Function
    rList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
 
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If adduser = True Then
                thelist.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If

End Function



Public Function MoveSprite(ByRef Sprite As PictureBox, ByRef Mask As PictureBox, ByRef Background As PictureBox, ByVal Direction As String, ByVal Distance_Pixels As Long, ByVal startX As Single, startY As Single, ByVal Speed As Long, Optional ByVal NumberOfFrames As Long = 1) As String

Dim x As Single, y As Single

Select Case Direction

Case "Up"
    
    x = startX
    
    For y = startY To Distance_Pixels + startY
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, x, y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Mask, 0, 0, Mask.ScaleWidth / NumberOfFrames, Mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
    
    Next y

Case "Down"
    
    x = startX
    
    For y = Distance_Pixels + startY To startY Step -1
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, x, y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Mask, 0, 0, Mask.ScaleWidth / NumberOfFrames, Mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
        
    Next y

Case "Left"
    
    y = startY

    For x = Distance_Pixels + startX To startX Step -1
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, x, y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Mask, 0, 0, Mask.ScaleWidth / NumberOfFrames, Mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
        
    Next x

Case "Right"

    y = startY

    For x = startX To Distance_Pixels + startX
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, x, y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Mask, 0, 0, Mask.ScaleWidth / NumberOfFrames, Mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
    
    Next x

End Select

End Function

Public Function DoBitBlt(ByRef Destination As PictureBox, ByVal DestinationX As Long, ByVal DestinationY As Long, ByVal DestinationWidth As Long, ByVal DestinationHeight As Long, ByRef Sprite As PictureBox, ByVal SpriteX As Long, ByVal SpriteY As Long, ByVal SpriteWidth As Long, ByVal SpriteHeight As Long, ByRef Mask As PictureBox, ByVal MaskX As Long, ByVal MaskY As Long, ByVal MaskWidth As Long, ByVal MaskHeight As Long) As Long

If DestinationWidth = SpriteWidth And DestinationHeight = SpriteHeight Then
    
    DoBitBlt = BitBlt(Destination.hdc, DestinationX, DestinationY, DestinationWidth, DestinationHeight, Mask.hdc, MaskX, MaskY, dwRop.SRCAND)
    DoBitBlt = BitBlt(Destination.hdc, DestinationX, DestinationY, DestinationWidth, DestinationHeight, Sprite.hdc, SpriteX, SpriteY, dwRop.SRCPAINT)

ElseIf DestinationWidth <> SpriteWidth Or DestinationHeight <> SpriteHeight Then
    
    DoBitBlt = StretchBlt(Destination.hdc, DestinationX, DestinationY, DestinationWidth, DestinationHeight, Mask.hdc, MaskX, MaskY, MaskWidth, MaskHeight, dwRop.SRCAND)
    DoBitBlt = StretchBlt(Destination.hdc, DestinationX, DestinationY, DestinationWidth, DestinationHeight, Sprite.hdc, SpriteX, SpriteY, SpriteWidth, SpriteHeight, dwRop.SRCPAINT)

Else

    DoBitBlt = 0
    
End If

End Function


Function ReplaceOneString(FullString As String, ReplaceWhat As String, ReplaceWith As String)
'case sensitive
Dim searchfor$, leftstring$, rightstring$
searchfor$ = InStr(FullString$, ReplaceWhat$)
If searchfor$ = 0 Then MsgBox "String not found.": Exit Function
leftstring$ = Left(FullString$, searchfor$ - 1)
rightstring$ = Mid(FullString$, searchfor$ + 1, Len(FullString$))
ReplaceOneString = leftstring$ + ReplaceWith$ + rightstring$
End Function

Function ROSNCS(FullString As String, ReplaceWhat As String, ReplaceWith As String)
'not case sensitive
'ROSNCS = Replace One String Not Case Sensative
Dim searchfor$, leftstring$, rightstring$
searchfor$ = InStr(UCase(FullString$), UCase(ReplaceWhat$))
If searchfor$ = 0 Then MsgBox "String not found.": Exit Function
leftstring$ = Left(FullString$, searchfor$ - 1)
rightstring$ = Mid(FullString$, searchfor$ + 1, Len(FullString$))
ROSNCS = leftstring$ + ReplaceWith$ + rightstring$
End Function

Sub CreateNewStartButton()
    Dim r As RECT
    tWnd = FindWindow("Shell_TrayWnd", vbNullString)
    bWnd = FindWindowEx(tWnd, ByVal 0&, "BUTTON", vbNullString)
    GetWindowRect bWnd, r
    ncWnd = CreateWindowEx(ByVal 0&, "BUTTON", "Hello !", WS_CHILD, 0, 0, r.Right - r.Left, r.Bottom - r.Top, tWnd, ByVal 0&, App.hInstance, ByVal 0&)
    ShowWindow ncWnd, SW_NORMAL
    ShowWindow bWnd, SW_HIDE
End Sub

Sub DestroyNewSB()
    ShowWindow bWnd, SW_NORMAL
    DestroyWindow ncWnd
End Sub


Function StripSpace(txt As String) As String
If InStr(txt$, " ") = 0 Then StripSpace$ = txt$: Exit Function
While InStr(txt$, " ")
txt$ = ReplaceOneString(txt$, " ", "")
DoEvents
Wend
StripSpace$ = txt$
End Function

Public Function ScreenWipe(Form As Form, CutSpeed As Integer) As Boolean
    Dim OldWidth As Integer
    Dim OldHeight As Integer
Form.WindowState = 0
If CutSpeed <= 0 Then
MsgBox "You cannot use 0 as a speed value"
Exit Function
End If
Do
OldWidth = Form.Width
Form.Width = Form.Width - CutSpeed
DoEvents
If Form.Width <> OldWidth Then
Form.Left = Form.Left + CutSpeed / 2
DoEvents
End If
OldHeight = Form.Height
Form.Height = Form.Height - CutSpeed
DoEvents
If Form.Height <> OldHeight Then
Form.Top = Form.Top + CutSpeed / 2
DoEvents
End If
Loop While Form.Width <> OldWidth Or Form.Height <> OldHeight
End Function

Public Function LineCount(TheString As String) As Long
Dim charcount$
charcount$ = InStr(TheString$, Chr(13))
If charcount$ <> 0 Then LineCount& = 1
Do
DoEvents
charcount$ = InStr(charcount$ + 1, TheString$, Chr(13))
If charcount$ <> 0 Then LineCount& = LineCount& + 1
DoEvents
Loop Until charcount$ = 0
LineCount& = LineCount& + 1
End Function

Public Function GetChatName() As String
GetChatName$ = GetCaption(FindChatRoom())
End Function

Public Function StripChatName()
StripChatName = StripSpace(GetChatName())
End Function

Public Function RoomBuster(room As String) As Long
stopbust = False
Dim aol&, mdi&, keyword&, aoledit&, msgboxx&, chatroom&
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
redo:
If stopbust = True Then Exit Function
Call RunAOLToolbar("11", "G")
Do
DoEvents
keyword& = FindWindowEx(mdi&, 0&, "AOL Child", "Keyword")
aoledit& = FindWindowEx(keyword&, 0&, "_AOL_Edit", vbNullString)
Loop Until keyword& <> 0 And aoledit& <> 0
Call SendMessageByString(aoledit&, WM_SETTEXT, 0, "aol://2719:2-2-" & room$)
Call PostMessage(aoledit&, WM_CHAR, 13, 0)
Do
DoEvents
msgboxx& = FindWindow("#32770", "America Online")
chatroom& = FindChatRoom()
Loop Until msgboxx& <> 0 Or chatroom& <> 0 And UCase(StripChatName()) = UCase(room$)
If msgboxx& <> 0 Then
Call SendMessage(msgboxx&, WM_CLOSE, 0, 0)
stopbust = False
GoTo redo
Exit Function
End If
If chatroom& <> 0 Then
stopbust = True
Exit Function
End If
End Function

Public Function LoadListboxRooms(List As ListBox, Directory As String) As Long
Dim a$
Open Directory$ For Input As #1
Do While Not EOF(1)
Input #1, a$
List.AddItem a$
DoEvents
Loop
Close #1
End Function

Public Function SaveListboxRooms(List As ListBox, Directory As String) As Long
Open Directory$ For Output As #1
For i = 0 To List.ListCount - 1
Write #1, List.List(i)
Next i
Close #1
End Function
Sub BarFadeFrm(frm, Style)
Dim cx, cy, f, F1, F2, i
frm.AutoRedraw = True
frm.Cls
frm.ScaleMode = 3
cx = frm.ScaleWidth / 2
cy = frm.ScaleHeight / 2
drawwidth = 2
For i = 255 To 0 Step -2
f = i / 255
F1 = 1 - f: F2 = 1 + f
If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If Style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
frm.Line (cx * F1, 0)-(cx * F2, cy * 2), , BF
Next i
End Sub
Sub CFadeFrm(frm, Style)
frm.AutoRedraw = True
frm.Cls
Dim cx, cy, i
frm.ScaleMode = 3
cx = frm.ScaleWidth \ 2
cy = frm.ScaleHeight \ 2
frm.drawwidth = 2
For i = 0 To 255
If Style = 1 Then frm.Circle (cx, cy), i, RGB(i, i, i)  'Black to white
If Style = 2 Then frm.Circle (cx, cy), i, RGB(0, i, i)  'Black to Cyan
If Style = 3 Then frm.Circle (cx, cy), i, RGB(i, 0, i)  'Black to Purple
If Style = 4 Then frm.Circle (cx, cy), i, RGB(i, i, 0)  'Black to Yellow
If Style = 5 Then frm.Circle (cx, cy), i, RGB(0, 0, i)  'Black to Blue
If Style = 6 Then frm.Circle (cx, cy), i, RGB(i, 0, 0)  'Black to Red
If Style = 7 Then frm.Circle (cx, cy), i, RGB(0, i, 0)  'Black to Green
If Style = 8 Then frm.Circle (cx, cy), i, RGB(0, i, 255)  'Blue to Green
If Style = 9 Then frm.Circle (cx, cy), i, RGB(i, i, 255)  'Blue to White
If Style = 11 Then frm.Circle (cx, cy), i, RGB(i, 0, 255)  'Blue to Purple
If Style = 12 Then frm.Circle (cx, cy), i, RGB(0, 0, 255 - i)  'Blue to Black
If Style = 13 Then frm.Circle (cx, cy), i, RGB(255, 0, i)  'Red to Purple
If Style = 14 Then frm.Circle (cx, cy), i, RGB(255, i, i)  'Red to White
If Style = 15 Then frm.Circle (cx, cy), i, RGB(255, i, 0)  'Red to Yellow
If Style = 16 Then frm.Circle (cx, cy), i, RGB(255 - i, 0, 0)  'Red to Black
If Style = 17 Then frm.Circle (cx, cy), i, RGB(i, 255, i)  'Green to White
If Style = 18 Then frm.Circle (cx, cy), i, RGB(0, 255, i)  'Green to Blue
If Style = 19 Then frm.Circle (cx, cy), i, RGB(i, 255, 0)  'Green to Yellow
If Style = 20 Then frm.Circle (cx, cy), i, RGB(0, 255 - i, 0)  'Green to Black
If Style = 21 Then frm.Circle (cx, cy), i, RGB(255 - i, 255 - i, 255 - i)  'White to Black
If Style = 22 Then frm.Circle (cx, cy), i, RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.Circle (cx, cy), i, RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.Circle (cx, cy), i, RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.Circle (cx, cy), i, RGB(255, 255, i)  'Yellow to White
If Style = 26 Then frm.Circle (cx, cy), i, RGB(255, i, 255)  'Purple to White
If Style = 27 Then frm.Circle (cx, cy), i, RGB(i, 255, 255)  'Cyan to White
If Style = 28 Then frm.Circle (cx, cy), i, RGB(255 - i, 255 - i, 0)  'Yellow to Black
If Style = 29 Then frm.Circle (cx, cy), i, RGB(255 - i, 0, 255 - i)  'Purple to Black
If Style = 30 Then frm.Circle (cx, cy), i, RGB(0, 255 - i, 255 - i)  'Cyan to Black
Dim s1, s2, s3
If Style = 31 Then frm.Circle (cx, cy), i, RGB(s1 - i, s2 - i, s3 - i)  'Selected color to black
Next i
End Sub

Sub DoubleFade(frm, Style)
frm.AutoRedraw = True
frm.Cls
Dim cx, cy, f, F1, F2, i
frm.AutoRedraw = True
frm.Cls
frm.ScaleMode = 3
cx = frm.ScaleWidth / 2
cy = frm.ScaleHeight / 2
Dim drawwidth
drawwidth = 2
For i = 255 To 0 Step -2
f = i / 255
F1 = 1 - f: F2 = 1 + f
If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If Style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
frm.Line (cx * F1, cy * F1)-(cx * F2, cy * F2), , BF
Next i
frm.ScaleMode = 3   ' Set ScaleMode to pixels.
cx = frm.ScaleWidth / 2 ' Get horizontal center.
cy = frm.ScaleHeight / 2    ' Get vertical center.
frm.drawwidth = 2
For i = 0 To 255
If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If Style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
f = i / 255  ' Perform interim
F1 = 1 - f: F2 = 1 + f  ' calculations.
frm.Line (cx * F1, cy)-(cx, cy * F1)   ' Draw upper-left.
frm.Line -(cx * F2, cy) ' Draw upper-right.
frm.Line -(cx, cy * F2) ' Draw lower-right.
frm.Line -(cx * F1, cy) ' Draw lower-left.
Next i
End Sub
Sub ExplosiveFade(frm, Style)
frm.AutoRedraw = True
frm.Cls
Dim cx, cy, f, F1, F2, i
frm.ScaleMode = 3
cx = frm.ScaleWidth / 2
cy = frm.ScaleHeight / 2
frm.drawwidth = 2
For i = 0 To 255
If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If Style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
f = i / 255  ' Perform interim
F1 = 1 - f: F2 = 1 + f  ' calculations.
frm.Line (cx * F1, cy)-(cx, cy * F1)   ' Draw upper-left.
frm.Line -(cx * F2, cy) ' Draw upper-right.
frm.Line -(cx, cy * F2) ' Draw lower-right.
frm.Line -(cx * F1, cy) ' Draw lower-left.
Next i
End Sub
Sub FadeFrm(frm, Style)
frm.ScaleMode = vbPixels
frm.AutoRedraw = True
frm.DrawStyle = vbInsideSolid
frm.Cls
frm.drawwidth = 2
frm.DrawMode = 13
frm.ScaleHeight = 256
For i = 0 To 255
If Style = 1 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, i), BF ' Black to white
If Style = 2 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, i, i), BF ' Black to Cyan
If Style = 3 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 0, i), BF ' Black to Purple
If Style = 4 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, 0), BF ' Black to Yellow
If Style = 5 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 0, i), BF ' Black to Blue
If Style = 6 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 0, 0), BF ' Black to Red
If Style = 7 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, i, 0), BF ' Black to Green
If Style = 8 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, i, 255), BF ' Blue to Green
If Style = 9 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, 255), BF ' Blue to White
If Style = 11 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 0, 255), BF ' Blue to Purple
If Style = 12 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 0, 255 - i), BF ' Blue to Black
If Style = 13 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 0, i), BF ' Red to Purple
If Style = 14 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, i, i), BF ' Red to White
If Style = 15 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, i, 0), BF ' Red to Yellow
If Style = 16 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 0, 0), BF ' Red to Black
If Style = 17 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 255, i), BF ' Green to White
If Style = 18 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 255, i), BF ' Green to Blue
If Style = 19 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 255, 0), BF ' Green to Yellow
If Style = 20 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 255 - i, 0), BF ' Green to Black
If Style = 21 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 255 - i, 255 - i), BF ' White to Black
If Style = 22 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 255, 255 - i), BF 'White to Yellow
If Style = 23 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 255 - i, 255), BF 'White to Purple
If Style = 24 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 255, 255), BF 'White to Cyan
If Style = 25 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 255, i), BF ' Yellow to White
If Style = 26 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, i, 255), BF ' Purple to White
If Style = 27 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 255, 255), BF ' Cyan to White
If Style = 28 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 255 - i, 0), BF ' Yellow to Black
If Style = 29 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 0, 255 - i), BF ' Purple to Black
If Style = 30 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 255 - i, 255 - i), BF ' Cyan to Black
If Style = 31 Then If i = 193 Then Exit Sub: frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, i), BF ' black to Gray
Next i
End Sub
Sub RFadeFrm(frm, Style)
Dim cx, cy, f, F1, F2, i
frm.AutoRedraw = True
frm.Cls
frm.ScaleMode = 3
cx = frm.ScaleWidth / 2
cy = frm.ScaleHeight / 2
drawwidth = 2
For i = 255 To 0 Step -2
f = i / 255
F1 = 1 - f: F2 = 1 + f
If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If Style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
frm.Line (cx * F1, cy * F1)-(cx * F2, cy * F2), , BF
Next i
End Sub
Sub SideFade(frm, Style)
Dim drawwidth
Dim cx, cy, f, F1, F2, i
frm.AutoRedraw = True
frm.Cls
frm.ScaleMode = 3
cx = frm.ScaleWidth
cy = frm.ScaleHeight
drawwidth = 2
For i = 255 To 0 Step -2
f = i / 255
F1 = 1 - f: F2 = 1 + f
If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If Style = 31 Then frm.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
frm.Line (cx * F1, 0)-(cx * F2, cy * 2), , BF
Next i
End Sub
Sub Text3D(Ctrl As Control, text, bevel, Style, Font)
Ctrl.AutoRedraw = True
Ctrl.FontSize = bevel * 1.4
Ctrl.Font = Font
For i = 0 To bevel * 10
If Style = 1 Then Ctrl.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then Ctrl.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then Ctrl.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then Ctrl.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then Ctrl.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then Ctrl.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then Ctrl.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then Ctrl.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then Ctrl.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then Ctrl.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then Ctrl.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then Ctrl.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then Ctrl.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then Ctrl.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then Ctrl.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then Ctrl.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then Ctrl.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then Ctrl.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then Ctrl.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then Ctrl.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then Ctrl.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then Ctrl.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then Ctrl.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then Ctrl.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then Ctrl.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then Ctrl.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then Ctrl.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then Ctrl.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then Ctrl.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
Dim s1, s2, s3
If Style = 31 Then Ctrl.ForeColor = RGB(s1 - i, s2 - i, s3 - i) ' Selected color to black
Ctrl.CurrentY = i \ 2
Ctrl.CurrentX = i \ 2
Ctrl.Print text
Next i
End Sub

Function RGBtoHEX(RGB)
Dim a$, length
    a$ = Hex(RGB)
    length = Len(a$)
    If length = 5 Then a$ = "0" & a$
    If length = 4 Then a$ = "00" & a$
    If length = 3 Then a$ = "000" & a$
    If length = 2 Then a$ = "0000" & a$
    If length = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function


Public Function IMessage(who As String, message As String)
Call RunAOLToolbar("3", "I")
Do
DoEvents
Dim aol&, mdi&, im&, aoledit&, richcntl&, aolicon&
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindChildByClass(aol&, "MDIClient")
im& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
aoledit& = FindWindowEx(im&, 0&, "_AOL_Edit", vbNullString)
richcntl& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
Loop Until richcntl& <> 0
Call SendMessageByString(aoledit&, WM_SETTEXT, 0, who$)
Call SendMessageByString(richcntl&, WM_SETTEXT, 0, message$)
aolicon& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(im&, aolicon&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(im&, aolicon&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(im&, aolicon&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(im&, aolicon&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(im&, aolicon&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(im&, aolicon&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(im&, aolicon&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(im&, aolicon&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(im&, aolicon&, "_AOL_Icon", vbNullString)
Call SendMessage(aolicon&, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(aolicon&, WM_KEYUP, VK_SPACE, 0)
Dim mesbox&
Do
DoEvents
mesbox& = FindWindow("#32770", "America Online Error")
Loop Until mesbox& <> 0 Or im& = 0
If mesbox& <> 0 Then
Dim staticc&, txt$
staticc& = FindChildByClass(mesbox&, "Static")
txt$ = GetCaption(staticc&)
Call SendMessage(mesbox&, WM_CLOSE, 0, 0)
Call SendMessage(im&, WM_CLOSE, 0, 0)
Exit Function
End If
End Function

Public Function ExportRegKey(KeyRoot As KeyRoot, KeyPath As String, FileName As String) As Boolean
  ' routine to export registry keys
  On Error Resume Next
  Dim hKey As Long
  Dim ReturnValue As Long

  ' check to see if allowed to do this
  If EnablePrivilege(SE_BACKUP_NAME) = False Then
    ExportRegKey = False
    Exit Function
  End If
  ' open the registry key
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0&, KEY_ALL_ACCESS, hKey)
  If ReturnValue <> 0 Then
    ' error encountered
    ExportRegKey = False
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  ' check for a copy of the export and delete old one if applicable
  If Dir(FileName) <> "" Then Kill FileName
  ' export the registry key
  ReturnValue = RegSaveKey(hKey, FileName, ByVal 0&)
  If ReturnValue = 0 Then
    ' no error encountered
    ExportRegKey = True
  Else
    ' error encountered
    ExportRegKey = False
  End If
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'ExportRegKey(KeyRoot As KeyRoot, KeyPath As String, FileName As String) As Boolean
'=========================================================================================
Public Function ImportRegKey(KeyRoot As KeyRoot, KeyPath As String, FileName As String) As Boolean
  ' routine to import registry keys
  ' will overwrite current settings, but will not create keys
  On Error Resume Next
  Dim hKey As Long
  Dim ReturnValue As Long

  ' check to see if allowed to do this
  If EnablePrivilege(SE_RESTORE_NAME) = False Then
    ImportRegKey = False
    Exit Function
  End If
  ' open the registry key
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0&, KEY_ALL_ACCESS, hKey)
  If ReturnValue <> 0 Then
    ' error encountered
    ImportRegKey = False
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  ' import the registry key
  ReturnValue = RegRestoreKey(hKey, FileName, REG_FORCE_RESTORE)
  If ReturnValue = 0 Then
    ' no error encountered
    ImportRegKey = True
  Else
    ' error encountered
    ImportRegKey = False
  End If
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'ImportRegKey(KeyRoot As KeyRoot, KeyPath As String, FileName As String) As Boolean
'=========================================================================================
Public Function ReadRegKey(KeyRoot As KeyRoot, KeyPath As String, SubKey As String, Optional NoKeyFoundValue As String = "") As String
  ' routine to read entry from registry
  On Error Resume Next
  Dim hKey As Long  ' receives a handle to the opened registry key
  Dim ReturnValue As Long  ' return value

  ' open the registry key
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_READ, hKey)
  If ReturnValue <> 0 Then
    ' key doesn't exist so return default value
    ReadRegKey = NoKeyFoundValue
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  ' get the keys value
  ReadRegKey = GetSubKeyValue(hKey, SubKey)
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'ReadRegKey(KeyRoot As KeyRoot, KeyPath As String, SubKey As String, Optional NoKeyFoundValue As String = "") As String
'=========================================================================================
Public Function WriteRegKey(KeyType As KeyType, KeyRoot As KeyRoot, KeyPath As String, SubKey As String, SubKeyValue As String) As Boolean
  ' routine to write entry to registry
  On Error Resume Next
  Dim hKey As Long  ' receives handle to the newly created or opened registry key
  Dim SecurityAttribute As SECURITY_ATTRIBUTES  ' security settings of the key
  Dim NewKey As Long  ' receives 1 if new key was created or 2 if an existing key was opened
  Dim ReturnValue As Long  ' return value

  ' Set the name of the new key and the default security settings
  SecurityAttribute.nLength = Len(SecurityAttribute)  ' size of the structure
  SecurityAttribute.lpSecurityDescriptor = 0  ' default security level
  SecurityAttribute.bInheritHandle = True  ' the default value for this setting

  ' create or open the registry key
  ReturnValue = RegCreateKeyEx(KeyRoot, KeyPath, 0, "", 0, KEY_WRITE, SecurityAttribute, hKey, NewKey)
  If ReturnValue <> 0 Then
    ' error encountered
    WriteRegKey = False
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If

  ' determine type of key and write it to the registry
  Select Case KeyType
    Case REG_SZ
      ReturnValue = RegSetValueEx(hKey, SubKey, 0, KeyType, ByVal SubKeyValue, Len(SubKeyValue))
    Case REG_DWORD
      ReturnValue = RegSetValueEx(hKey, SubKey, 0, KeyType, CLng(SubKeyValue), 4)
    Case REG_BINARY
      ReturnValue = RegSetValueEx(hKey, SubKey, 0, KeyType, CByte(SubKeyValue), 4)
  End Select

  If ReturnValue = 0 Then
    ' no error encountered
    WriteRegKey = True
  Else
    ' error encountered
    WriteRegKey = False
  End If

  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'WriteRegKey(KeyType As KeyType, KeyRoot As KeyRoot, KeyPath As String, SubKey As String, SubKeyValue As String) As Boolean
'=========================================================================================
Public Function EnumerateRegKeys(KeyRoot As KeyRoot, KeyPath As String) As String
  ' routine to enumerate all subkeys under a registry key
  On Error Resume Next
  Dim hKey As Long  ' receives a handle to the opened registry key
  Dim ReturnValue As Long  ' return value
  Dim Counter As Long
  Dim MyBuffer As String
  Dim MyBufferSize As Long
  Dim ClassNameBuffer As String
  Dim ClassNameBufferSize As Long
  Dim LastWrite As FILETIME

  ' open the registry key
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
  If ReturnValue <> 0 Then
    ' key doesn't exist so return default value
    EnumerateRegKeys = ""
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  Counter = 0
  ' loop until no more registry keys
  Do Until ReturnValue <> 0
    MyBuffer = Space(255)
    ClassNameBuffer = Space(255)
    MyBufferSize = 255
    ClassNameBufferSize = 255
    ReturnValue = RegEnumKeyEx(hKey, Counter, MyBuffer, MyBufferSize, ByVal 0, ClassNameBuffer, ClassNameBufferSize, LastWrite)
    If ReturnValue = 0 Then
      MyBuffer = Left$(MyBuffer, MyBufferSize)
      ClassNameBuffer = Left$(ClassNameBuffer, ClassNameBufferSize)
      EnumerateRegKeys = EnumerateRegKeys & MyBuffer & ","
    End If
    Counter = Counter + 1
  Loop
  ' trim off the last delimiter
  If EnumerateRegKeys <> "" Then EnumerateRegKeys = Left$(EnumerateRegKeys, Len(EnumerateRegKeys) - 1)
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'EnumerateRegKeys(KeyRoot As KeyRoot, KeyPath As String) As String
'=========================================================================================
Public Function EnumerateRegKeyValues(KeyRoot As KeyRoot, KeyPath As String) As String
  ' routine to enumerate all the values under a key in the registry
  On Error Resume Next
  Dim hKey As Long  ' receives a handle to the opened registry key
  Dim ReturnValue As Long  ' return value
  Dim Counter As Long
  Dim MyBuffer As String
  Dim MyBufferSize As Long
  Dim KeyType As KeyType

  ' open the registry key to enumerate the values of.
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_QUERY_VALUE, hKey)
  ' check to see if an error occured.
  If ReturnValue <> 0 Then
    EnumerateRegKeyValues = ""
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  Counter = 0
  ' loop until no more registry keys value
  Do Until ReturnValue <> 0
    MyBuffer = Space(255)
    MyBufferSize = 255
    ReturnValue = RegEnumValue(hKey, Counter, MyBuffer, MyBufferSize, 0, KeyType, ByVal 0&, ByVal 0&) 'ByteData(0), ByteDataSize)
    If ReturnValue = 0 Then
      MyBuffer = Left$(MyBuffer, MyBufferSize)
      EnumerateRegKeyValues = EnumerateRegKeyValues & MyBuffer & "*"
      EnumerateRegKeyValues = EnumerateRegKeyValues & GetSubKeyValue(hKey, MyBuffer) & ","
    End If
    Counter = Counter + 1
  Loop
  ' trim off the last delimiter
  If EnumerateRegKeyValues <> "" Then EnumerateRegKeyValues = Left$(EnumerateRegKeyValues, Len(EnumerateRegKeyValues) - 1)
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'EnumerateRegKeyValues(KeyRoot As KeyRoot, KeyPath As String) As String
'=========================================================================================
Public Function DeleteRegKey(KeyRoot As KeyRoot, KeyPath As String, SubKey As String) As Boolean
  ' routine to delete a registry key
  ' under Win NT/2000 all subkeys must be deleted first
  ' under Win 9x all subkeys are deleted
  On Error Resume Next
  Dim ReturnValue As Long  ' return value

  ' Attempt to delete the desired registry key.
  ReturnValue = RegDeleteKey(KeyRoot, KeyPath & "\" & SubKey)
  If ReturnValue = 0 Then
    ' no error encountered
    DeleteRegKey = True
  Else
    ' error encountered
    DeleteRegKey = False
  End If
End Function 'DeleteRegKey(KeyRoot As KeyRoot, KeyPath As String, SubKey As String) As Boolean
'=========================================================================================
Public Function DeleteRegKeyValue(KeyRoot As KeyRoot, KeyPath As String, Optional SubKey As String = "") As Boolean
  ' routine to delete a value from a key (but not the key) in the registry
  On Error Resume Next
  Dim hKey As Long  ' handle to the open registry key
  Dim ReturnValue As Long  ' return value

  ' First, open up the registry key which holds the value to delete.
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_ALL_ACCESS, hKey)
  If ReturnValue <> 0 Then
    ' error encountered
    DeleteRegKeyValue = False
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  ' check to see if we are deleting a subkey or primary key
  If SubKey = "" Then SubKey = KeyPath
  ' successfully opened registry key so delete the desired value from the key.
  ReturnValue = RegDeleteValue(hKey, SubKey)
  If ReturnValue = 0 Then
    ' no error encountered
    DeleteRegKeyValue = True
  Else
    ' error encountered
    DeleteRegKeyValue = False
  End If
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'DeleteRegKeyValue(KeyRoot As KeyRoot, KeyPath As String, Optional SubKey As String = "") As Boolean
'=========================================================================================
Public Function GetSubKeyValue(ByVal hKey As Long, ByVal SubKey As String) As String
  ' routine to get the registry key value and convert to a string
  On Error Resume Next
  Dim ReturnValue As Long
  Dim KeyType As KeyType
  Dim MyBuffer As String
  Dim MyBufferSize As Long

  'get registry key information
  ReturnValue = RegQueryValueEx(hKey, SubKey, 0, KeyType, ByVal 0, MyBufferSize)
  If ReturnValue = 0 Then ' no error encountered
    ' determine what the KeyType is
    Select Case KeyType
      Case REG_SZ
        ' create a buffer
        MyBuffer = String(MyBufferSize, Chr$(0))
        ' retrieve the key's content
        ReturnValue = RegQueryValueEx(hKey, SubKey, 0, 0, ByVal MyBuffer, MyBufferSize)
        If ReturnValue = 0 Then
          ' remove the unnecessary chr$(0)'s
          GetSubKeyValue = Left$(MyBuffer, InStr(1, MyBuffer, Chr$(0)) - 1)
        End If
      Case Else 'REG_DWORD or REG_BINARY
        Dim MyNewBuffer As Long
        ' retrieve the key's value
        ReturnValue = RegQueryValueEx(hKey, SubKey, 0, 0, MyNewBuffer, MyBufferSize)
        If ReturnValue = 0 Then ' no error encountered
          GetSubKeyValue = MyNewBuffer
        End If
    End Select
  End If
End Function 'GetSubKeyValue(ByVal hKey As Long, ByVal SubKey As String) As String
'=========================================================================================
Public Function EnablePrivilege(seName As String) As Boolean
  ' routine to enable inport/export of registry settings
  On Error Resume Next
  Dim p_lngRtn As Long
  Dim p_lngToken As Long
  Dim p_lngBufferLen As Long
  Dim p_typLUID As LUID
  Dim p_typTokenPriv As TOKEN_PRIVILEGES
  Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES

  ' open the current process token
  p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
  If p_lngRtn = 0 Then
    ' error encountered
    EnablePrivilege = False
    Exit Function
  End If
  If Err.LastDllError <> 0 Then
    ' error encountered
    EnablePrivilege = False
    Exit Function
  End If
  ' look up the privileges LUID
  p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)
  If p_lngRtn = 0 Then
    ' error encountered
    EnablePrivilege = False
    Exit Function
  End If
  ' adjust the program's security privilege.
  p_typTokenPriv.PrivilegeCount = 1
  p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
  p_typTokenPriv.Privileges.pLuid = p_typLUID
  ' try to adjust privileges and return success or failure
  EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function 'EnablePrivilege(seName As String) As Boolean

Public Function DecryptWithALP(strData As String) As String
    Dim strALPKey As String
    Dim strALPKeyMask As String
    Dim lngIterator As Long
    Dim blnOscillator As Boolean
    Dim strOutput As String
    Dim lngHex As Long
    If Len(strData) = 0 Then
        Exit Function
    End If
    strALPKeyMask = Right$(String$(lngALPKeyLength, "0") + DoubleToBinary(CLng("&H" + Left$(strData, 2))), lngALPKeyLength)
    strData = Right$(strData, Len(strData) - 2)
    For lngIterator = lngALPKeyLength To 1 Step -1
        If Mid$(strALPKeyMask, lngIterator, 1) = "1" Then
            strALPKey = Left$(strData, 1) + strALPKey
            strData = Right$(strData, Len(strData) - 1)
        Else
            strALPKey = Right$(strData, 1) + strALPKey
            strData = Left$(strData, Len(strData) - 1)
        End If
    Next lngIterator
    lngIterator = 0
    Do Until Len(strData) = 0
        blnOscillator = Not blnOscillator
        lngIterator = lngIterator + 1
        If lngIterator > lngALPKeyLength Then
            lngIterator = 1
        End If
        lngHex = IIf(blnOscillator, CLng("&H" + Left$(strData, 2) - Asc(Mid$(strALPKey, lngIterator, 1))), CLng("&H" + Left$(strData, 2) + Asc(Mid$(strALPKey, lngIterator, 1))))
        If lngHex > 255 Then
            lngHex = lngHex - 255
        ElseIf lngHex < 0 Then
            lngHex = lngHex + 255
        End If
        strOutput = strOutput + Chr$(lngHex)
        strData = Right$(strData, Len(strData) - 2)
    Loop
    DecryptWithALP = strOutput
End Function
Public Function DecryptWithClipper(ByVal strData As String, ByVal strCryptKey As String) As String
    Dim strDecryptionChunk As String
    Dim strDecryptedText As String
    On Error Resume Next
    InitCrypt strCryptKey
    Do Until Len(strData) < 16
        strDecryptionChunk = ""
        strDecryptionChunk = Left$(strData, 16)
        strData = Right$(strData, Len(strData) - 16)
        If Len(strDecryptionChunk) > 0 Then
            strDecryptedText = strDecryptedText + PerformClipperDecryption(strDecryptionChunk)
        End If
    Loop
    DecryptWithClipper = strDecryptedText
End Function
Public Function DecryptWithCSP(ByVal strData As String, ByVal strCryptKey As String) As String
    Dim lngEncryptionCount As Long
    Dim strDecrypted As String
    Dim strCurrentCryptKey As String
    If EncryptionCSPConnect() Then
        lngEncryptionCount = DecryptNumber(Mid$(strData, 1, 8))
        strCurrentCryptKey = strCryptKey & lngEncryptionCount
        strDecrypted = EncryptDecrypt(Mid$(strData, 9), strCurrentCryptKey, False)
        DecryptWithCSP = strDecrypted
        EncryptionCSPDisconnect
    End If
End Function
Public Function EncryptWithALP(strData As String) As String
    Dim strALPKey As String
    Dim strALPKeyMask As String
    Dim lngIterator As Long
    Dim blnOscillator As Boolean
    Dim strOutput As String
    Dim lngHex As Long
    If Len(strData) = 0 Then
        Exit Function
    End If
    Randomize
    For lngIterator = 1 To lngALPKeyLength
        strALPKey = strALPKey + Trim$(Hex$(Int(16 * Rnd)))
        strALPKeyMask = strALPKeyMask + Trim$(Int(2 * Rnd))
    Next lngIterator
    lngIterator = 0
    Do Until Len(strData) = 0
        blnOscillator = Not blnOscillator
        lngIterator = lngIterator + 1
        If lngIterator > lngALPKeyLength Then
            lngIterator = 1
        End If
        lngHex = IIf(blnOscillator, CLng(Asc(Left$(strData, 1)) + Asc(Mid$(strALPKey, lngIterator, 1))), CLng(Asc(Left$(strData, 1)) - Asc(Mid$(strALPKey, lngIterator, 1))))
        If lngHex > 255 Then
            lngHex = lngHex - 255
        ElseIf lngHex < 0 Then
            lngHex = lngHex + 255
        End If
        strOutput = strOutput + Right$(String$(2, "0") + Hex$(lngHex), 2)
        strData = Right$(strData, Len(strData) - 1)
    Loop
    For lngIterator = 1 To lngALPKeyLength
        If Mid$(strALPKeyMask, lngIterator, 1) = "1" Then
            strOutput = Mid$(strALPKey, lngIterator, 1) + strOutput
        Else
            strOutput = strOutput + Mid$(strALPKey, lngIterator, 1)
        End If
    Next lngIterator
    EncryptWithALP = Right$(String$(2, "0") + Hex$(BinaryToDouble(strALPKeyMask)), 2) + strOutput
End Function
Public Function EncryptWithClipper(ByVal strData As String, ByVal strCryptKey As String) As String
    Dim strEncryptionChunk As String
    Dim strEncryptedText As String
    If Len(strData) > 0 Then
        InitCrypt strCryptKey
        Do Until Len(strData) = 0
            strEncryptionChunk = ""
            If Len(strData) > 6 Then
                strEncryptionChunk = Left$(strData, 6)
                strData = Right$(strData, Len(strData) - 6)
            Else
                strEncryptionChunk = Left$(strData + Space(6), 6)
                strData = ""
            End If
            If Len(strEncryptionChunk) > 0 Then
                strEncryptedText = strEncryptedText + PerformClipperEncryption(strEncryptionChunk)
            End If
        Loop
    End If
    EncryptWithClipper = strEncryptedText
End Function
Public Function EncryptWithCSP(ByVal strData As String, ByVal strCryptKey As String) As String
    Dim strEncrypted As String
    Dim lngEncryptionCount As Long
    Dim strCurrentCryptKey As String
    If EncryptionCSPConnect() Then
        lngEncryptionCount = 0
        strCurrentCryptKey = strCryptKey & lngEncryptionCount
        strEncrypted = EncryptDecrypt(strData, strCurrentCryptKey, True)
        Do While (InStr(1, strEncrypted, vbCr) > 0) Or (InStr(1, strEncrypted, vbLf) > 0) Or (InStr(1, strEncrypted, Chr$(0)) > 0) Or (InStr(1, strEncrypted, vbTab) > 0)
            lngEncryptionCount = lngEncryptionCount + 1
            strCurrentCryptKey = strCryptKey & lngEncryptionCount
            strEncrypted = EncryptDecrypt(strData, strCurrentCryptKey, True)
            If lngEncryptionCount = 99999999 Then
                Err.Raise vbObjectError + 999, "EncryptWithCSP", "This Data cannot be successfully encrypted"
                EncryptWithCSP = ""
                Exit Function
            End If
        Loop
        EncryptWithCSP = EncryptNumber(lngEncryptionCount) & strEncrypted
        EncryptionCSPDisconnect
    End If
End Function
Public Function GetCSPDetails() As String
    Dim lngDataLength As Long
    Dim bytContainer() As Byte
    If EncryptionCSPConnect Then
        If lngCryptProvider = 0 Then
            GetCSPDetails = "Not connected to CSP"
            Exit Function
        End If
        lngDataLength = 1000
        ReDim bytContainer(lngDataLength)
        If CryptGetProvParam(lngCryptProvider, PP_NAME, bytContainer(0), lngDataLength, 0) <> 0 Then
            GetCSPDetails = "Cryptographic Service Provider name: " & ByteToString(bytContainer, lngDataLength)
        End If
        lngDataLength = 1000
        ReDim bytContainer(lngDataLength)
        If CryptGetProvParam(lngCryptProvider, PP_CONTAINER, bytContainer(0), lngDataLength, 0) <> 0 Then
            GetCSPDetails = GetCSPDetails & vbCrLf & "Key Container name: " & ByteToString(bytContainer, lngDataLength)
        End If
        EncryptionCSPDisconnect
    Else
        GetCSPDetails = "Not connected to CSP"
    End If
End Function
Public Function DecryptNumber(ByVal strData As String) As Long
    Dim lngIterator As Long
    For lngIterator = 1 To 8
        DecryptNumber = (10 * DecryptNumber) + (Asc(Mid$(strData, lngIterator, 1)) - Asc(Mid$(ENCRYPT_NUMBERKEY, lngIterator, 1)))
    Next lngIterator
End Function
Public Function EncryptDecrypt(ByVal strData As String, ByVal strCryptKey As String, ByVal Encrypt As Boolean) As String
    Dim lngDataLength As Long
    Dim strTempData As String
    Dim lngHaslngCryptKey As Long
    Dim lngCryptKey As Long
    If lngCryptProvider = 0 Then
        'Err.Raise vbObjectError + 999, "EncryptDecrypt", "Not connected to CSP"
        Exit Function
    End If
    If CryptCreateHash(lngCryptProvider, CALG_MD5, 0, 0, lngHaslngCryptKey) = 0 Then
        Err.Raise vbObjectError + 999, "EncryptDecrypt", "Error during CryptCreateHash."
    End If
    If CryptHashData(lngHaslngCryptKey, strCryptKey, Len(strCryptKey), 0) = 0 Then
        Err.Raise vbObjectError + 999, "EncryptDecrypt", "Error during CryptHashData."
    End If
    If CryptDeriveKey(lngCryptProvider, ENCRYPT_ALGORITHM, lngHaslngCryptKey, 0, lngCryptKey) = 0 Then
        Err.Raise vbObjectError + 999, "EncryptDecrypt", "Error during CryptDeriveKey!"
    End If
    strTempData = strData
    lngDataLength = Len(strData)
    If Encrypt Then
        If CryptEncrypt(lngCryptKey, 0, 1, 0, strTempData, lngDataLength, lngDataLength) = 0 Then
            Err.Raise vbObjectError + 999, "EncryptDecrypt", "Error during CryptEncrypt."
        End If
    Else
        If CryptDecrypt(lngCryptKey, 0, 1, 0, strTempData, lngDataLength) = 0 Then
            Err.Raise vbObjectError + 999, "EncryptDecrypt", "Error during CryptDecrypt."
        End If
    End If
    EncryptDecrypt = Mid$(strTempData, 1, lngDataLength)
    If lngCryptKey <> 0 Then
        CryptDestroyKey lngCryptKey
    End If
    If lngHaslngCryptKey <> 0 Then
        CryptDestroyHash lngHaslngCryptKey
    End If
End Function
Public Function EncryptionCSPConnect() As Boolean
    If Len(strKeyContainer) = 0 Then
        strKeyContainer = "FastTrack"
    End If
    If CryptAcquireContext(lngCryptProvider, strKeyContainer, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0 Then
        If CryptAcquireContext(lngCryptProvider, strKeyContainer, SERVICE_PROVIDER, PROV_RSA_FULL, 0) = 0 Then
            Err.Raise vbObjectError + 999, "EncryptionCSPConnect", "Error during CryptAcquireContext for a new key container." & vbCrLf & "A container with this name probably already exists."
            EncryptionCSPConnect = False
            Exit Function
        End If
    End If
    EncryptionCSPConnect = True
End Function
Public Function EncryptNumber(ByVal lngData As Long) As String
    Dim lngIterator As Long
    Dim strData As String
    strData = Format$(lngData, "00000000")
    For lngIterator = 1 To 8
        EncryptNumber = EncryptNumber & Chr$(Asc(Mid$(ENCRYPT_NUMBERKEY, lngIterator, 1)) + Val(Mid$(strData, lngIterator, 1)))
    Next lngIterator
End Function
Public Sub EncryptionCSPDisconnect()
    If lngCryptProvider <> 0 Then
        CryptReleaseContext lngCryptProvider, 0
    End If
End Sub
Public Sub InitCrypt(ByRef strEncryptionKey As String)
    avarSeedValues = Array("A3", "D7", "09", "83", "F8", "48", "F6", "F4", "B3", "21", "15", "78", "99", "B1", "AF", _
    "F9", "E7", "2D", "4D", "8A", "CE", "4C", "CA", "2E", "52", "95", "D9", "1E", "4E", "38", "44", "28", "0A", "DF", _
    "02", "A0", "17", "F1", "60", "68", "12", "B7", "7A", "C3", "E9", "FA", "3D", "53", "96", "84", "6B", "BA", "F2", _
    "63", "9A", "19", "7C", "AE", "E5", "F5", "F7", "16", "6A", "A2", "39", "B6", "7B", "0F", "C1", "93", "81", "1B", _
    "EE", "B4", "1A", "EA", "D0", "91", "2F", "B8", "55", "B9", "DA", "85", "3F", "41", "BF", "E0", "5A", "58", "80", _
    "5F", "66", "0B", "D8", "90", "35", "D5", "C0", "A7", "33", "06", "65", "69", "45", "00", "94", "56", "6D", "98", _
    "9B", "76", "97", "FC", "B2", "C2", "B0", "FE", "DB", "20", "E1", "EB", "D6", "E4", "DD", "47", "4A", "1D", "42", _
    "ED", "9E", "6E", "49", "3C", "CD", "43", "27", "D2", "07", "D4", "DE", "C7", "67", "18", "89", "CB", "30", "1F", _
    "8D", "C6", "8F", "AA", "C8", "74", "DC", "C9", "5D", "5C", "31", "A4", "70", "88", "61", "2C", "9F", "0D", "2B", _
    "87", "50", "82", "54", "64", "26", "7D", "03", "40", "34", "4B", "1C", "73", "D1", "C4", "FD", "3B", "CC", "FB", _
    "7F", "AB", "E6", "3E", "5B", "A5", "AD", "04", "23", "9C", "14", "51", "22", "F0", "29", "79", "71", "7E", "FF", _
    "8C", "0E", "E2", "0C", "EF", "BC", "72", "75", "6F", "37", "A1", "EC", "D3", "8E", "62", "8B", "86", "10", "E8", _
    "08", "77", "11", "BE", "92", "4F", "24", "C5", "32", "36", "9D", "CF", "F3", "A6", "BB", "AC", "5E", "6C", "A9", _
    "13", "57", "25", "B5", "E3", "BD", "A8", "3A", "01", "05", "59", "2A", "46")
    SetKey strEncryptionKey
End Sub
Public Function PerformClipperDecryption(ByVal strData As String) As String
    Dim bytChunk(1 To 4, 0 To 32) As String
    Dim bytCounter(0 To 32) As Byte
    Dim lngIterator As Long
    Dim strDecryptedData As String
    On Error Resume Next
    bytChunk(1, 32) = Mid(strData, 1, 4)
    bytChunk(2, 32) = Mid(strData, 5, 4)
    bytChunk(3, 32) = Mid(strData, 9, 4)
    bytChunk(4, 32) = Mid(strData, 13, 4)
    lngSeedLevel = 32
    lngDecryptPointer = 31
    For lngIterator = 0 To 32
        bytCounter(lngIterator) = lngIterator + 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel - 1) = PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey())
        bytChunk(2, lngSeedLevel - 1) = PerformXOR(PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey()), PerformXOR(bytChunk(3, lngSeedLevel), Hex(bytCounter(lngSeedLevel - 1))))
        bytChunk(3, lngSeedLevel - 1) = bytChunk(4, lngSeedLevel)
        bytChunk(4, lngSeedLevel - 1) = bytChunk(1, lngSeedLevel)
        lngDecryptPointer = lngDecryptPointer - 1
        lngSeedLevel = lngSeedLevel - 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel - 1) = PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey())
        bytChunk(2, lngSeedLevel - 1) = bytChunk(3, lngSeedLevel)
        bytChunk(3, lngSeedLevel - 1) = bytChunk(4, lngSeedLevel)
        bytChunk(4, lngSeedLevel - 1) = PerformXOR(PerformXOR(bytChunk(1, lngSeedLevel), bytChunk(2, lngSeedLevel)), Hex(bytCounter(lngSeedLevel - 1)))
        lngDecryptPointer = lngDecryptPointer - 1
        lngSeedLevel = lngSeedLevel - 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel - 1) = PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey())
        bytChunk(2, lngSeedLevel - 1) = PerformXOR(PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey()), PerformXOR(bytChunk(3, lngSeedLevel), Hex(bytCounter(lngSeedLevel - 1))))
        bytChunk(3, lngSeedLevel - 1) = bytChunk(4, lngSeedLevel)
        bytChunk(4, lngSeedLevel - 1) = bytChunk(1, lngSeedLevel)
        lngDecryptPointer = lngDecryptPointer - 1
        lngSeedLevel = lngSeedLevel - 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel - 1) = PerformClipperDecryptionChunk(bytChunk(2, lngSeedLevel), astrEncryptionKey())
        bytChunk(2, lngSeedLevel - 1) = bytChunk(3, lngSeedLevel)
        bytChunk(3, lngSeedLevel - 1) = bytChunk(4, lngSeedLevel)
        bytChunk(4, lngSeedLevel - 1) = PerformXOR(PerformXOR(bytChunk(1, lngSeedLevel), bytChunk(2, lngSeedLevel)), Hex(bytCounter(lngSeedLevel - 1)))
        lngDecryptPointer = lngDecryptPointer - 1
        lngSeedLevel = lngSeedLevel - 1
    Next lngIterator
    strDecryptedData = HexToString(bytChunk(1, 0) & bytChunk(2, 0) & bytChunk(3, 0) & bytChunk(4, 0))
    If InStr(strDecryptedData, Chr$(0)) > 0 Then
        strDecryptedData = Left$(strDecryptedData, InStr(strDecryptedData, Chr$(0)) - 1)
    End If
    PerformClipperDecryption = strDecryptedData
End Function
Public Function PerformClipperDecryptionChunk(ByVal strData As String, ByRef strEncryptionKey() As String) As String
    Dim astrDecryptionLevel(1 To 6) As String
    Dim strDecryptedString As String
    astrDecryptionLevel(5) = Mid(strData, 1, 2)
    astrDecryptionLevel(6) = Mid(strData, 3, 2)
    strDecryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrDecryptionLevel(5), strEncryptionKey((4 * lngDecryptPointer) + 3)))))
    astrDecryptionLevel(4) = PerformXOR(strDecryptedString, astrDecryptionLevel(6))
    strDecryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrDecryptionLevel(4), strEncryptionKey((4 * lngDecryptPointer) + 2)))))
    astrDecryptionLevel(3) = PerformXOR(strDecryptedString, astrDecryptionLevel(5))
    strDecryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrDecryptionLevel(3), strEncryptionKey((4 * lngDecryptPointer) + 1)))))
    astrDecryptionLevel(2) = PerformXOR(strDecryptedString, astrDecryptionLevel(4))
    strDecryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrDecryptionLevel(2), strEncryptionKey(4 * lngDecryptPointer)))))
    astrDecryptionLevel(1) = PerformXOR(strDecryptedString, astrDecryptionLevel(3))
    strDecryptedString = astrDecryptionLevel(1) & astrDecryptionLevel(2)
    PerformClipperDecryptionChunk = strDecryptedString
End Function
Public Function PerformClipperEncryption(ByVal strData As String) As String
    Dim bytChunk(1 To 4, 0 To 32) As String
    Dim lngCounter As Long
    Dim lngIterator As Long
    On Error Resume Next
    strData = StringToHex(strData)
    bytChunk(1, 0) = Mid(strData, 1, 4)
    bytChunk(2, 0) = Mid(strData, 5, 4)
    bytChunk(3, 0) = Mid(strData, 9, 4)
    bytChunk(4, 0) = Mid(strData, 13, 4)
    lngSeedLevel = 0
    lngCounter = 1
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel + 1) = PerformXOR(PerformXOR(PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey()), bytChunk(4, lngSeedLevel)), Hex(lngCounter))
        bytChunk(2, lngSeedLevel + 1) = PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey())
        bytChunk(3, lngSeedLevel + 1) = bytChunk(2, lngSeedLevel)
        bytChunk(4, lngSeedLevel + 1) = bytChunk(3, lngSeedLevel)
        lngCounter = lngCounter + 1
        lngSeedLevel = lngSeedLevel + 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel + 1) = bytChunk(4, lngSeedLevel)
        bytChunk(2, lngSeedLevel + 1) = PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey())
        bytChunk(3, lngSeedLevel + 1) = PerformXOR(PerformXOR(bytChunk(1, lngSeedLevel), bytChunk(2, lngSeedLevel)), Hex(lngCounter))
        bytChunk(4, lngSeedLevel + 1) = bytChunk(3, lngSeedLevel)
        lngCounter = lngCounter + 1
        lngSeedLevel = lngSeedLevel + 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel + 1) = PerformXOR(PerformXOR(PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey()), bytChunk(4, lngSeedLevel)), Hex(lngCounter))
        bytChunk(2, lngSeedLevel + 1) = PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey())
        bytChunk(3, lngSeedLevel + 1) = bytChunk(2, lngSeedLevel)
        bytChunk(4, lngSeedLevel + 1) = bytChunk(3, lngSeedLevel)
        lngCounter = lngCounter + 1
        lngSeedLevel = lngSeedLevel + 1
    Next lngIterator
    For lngIterator = 1 To 8
        bytChunk(1, lngSeedLevel + 1) = bytChunk(4, lngSeedLevel)
        bytChunk(2, lngSeedLevel + 1) = PerformClipperEncryptionChunk(bytChunk(1, lngSeedLevel), astrEncryptionKey())
        bytChunk(3, lngSeedLevel + 1) = PerformXOR(PerformXOR(bytChunk(1, lngSeedLevel), bytChunk(2, lngSeedLevel)), Hex(lngCounter))
        bytChunk(4, lngSeedLevel + 1) = bytChunk(3, lngSeedLevel)
        lngCounter = lngCounter + 1
        lngSeedLevel = lngSeedLevel + 1
    Next lngIterator
    PerformClipperEncryption = bytChunk(1, 32) & bytChunk(2, 32) & bytChunk(3, 32) & bytChunk(4, 32)
End Function
Public Function PerformClipperEncryptionChunk(ByVal strData As String, ByRef strEncryptionKey() As String) As String
    Dim astrEncryptionLevel(1 To 6) As String
    Dim strEncryptedString As String
    astrEncryptionLevel(1) = Mid(strData, 1, 2)
    astrEncryptionLevel(2) = Mid(strData, 3, 2)
    strEncryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrEncryptionLevel(2), strEncryptionKey(4 * lngSeedLevel)))))
    astrEncryptionLevel(3) = PerformXOR(strEncryptedString, astrEncryptionLevel(1))
    strEncryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrEncryptionLevel(3), strEncryptionKey((4 * lngSeedLevel) + 1)))))
    astrEncryptionLevel(4) = PerformXOR(strEncryptedString, astrEncryptionLevel(2))
    strEncryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrEncryptionLevel(4), strEncryptionKey((4 * lngSeedLevel) + 2)))))
    astrEncryptionLevel(5) = PerformXOR(strEncryptedString, astrEncryptionLevel(3))
    strEncryptedString = avarSeedValues(CByte(PerformTranslation(PerformXOR(astrEncryptionLevel(5), strEncryptionKey((4 * lngSeedLevel) + 3)))))
    astrEncryptionLevel(6) = PerformXOR(strEncryptedString, astrEncryptionLevel(4))
    strEncryptedString = astrEncryptionLevel(5) & astrEncryptionLevel(6)
    PerformClipperEncryptionChunk = strEncryptedString
End Function
Public Function PerformTranslation(ByVal strData As String) As Double
    Dim strTranslationString As String
    Dim strTranslationChunk As String
    Dim lngTranslationIterator As Long
    Dim lngHexConversion As Long
    Dim lngHexConversionIterator As Long
    Dim dblTranslation As Double
    Dim lngTranslationMarker As Long
    Dim lngTranslationModifier As Long
    Dim lngTranslationLayerModifier As Long
    strTranslationString = strData
    strTranslationString = Right$(strTranslationString, 8)
    strTranslationChunk = String$(8 - Len(strTranslationString), "0") + strTranslationString
    strTranslationString = ""
    For lngTranslationIterator = 1 To 8
        lngHexConversion = Val("&H" + Mid$(strTranslationChunk, lngTranslationIterator, 1))
        For lngHexConversionIterator = 3 To 0 Step -1
            If lngHexConversion And 2 ^ lngHexConversionIterator Then
                strTranslationString = strTranslationString + "1"
            Else
                strTranslationString = strTranslationString + "0"
            End If
        Next lngHexConversionIterator
    Next lngTranslationIterator
    dblTranslation = 0
    For lngTranslationIterator = Len(strTranslationString) To 1 Step -1
        If Mid(strTranslationString, lngTranslationIterator, 1) = "1" Then
            lngTranslationLayerModifier = 1
            lngTranslationMarker = (Len(strTranslationString) - lngTranslationIterator)
            lngTranslationModifier = 2
            Do While lngTranslationMarker > 0
                Do While (lngTranslationMarker / 2) = (lngTranslationMarker \ 2)
                    lngTranslationModifier = (lngTranslationModifier * lngTranslationModifier) Mod 255
                    lngTranslationMarker = lngTranslationMarker / 2
                Loop
                lngTranslationLayerModifier = (lngTranslationModifier * lngTranslationLayerModifier) Mod 255
                lngTranslationMarker = lngTranslationMarker - 1
            Loop
            dblTranslation = dblTranslation + lngTranslationLayerModifier
        End If
    Next lngTranslationIterator
    PerformTranslation = dblTranslation
End Function
Public Function PerformXOR(ByVal strData As String, ByVal strMask As String) As String
    Dim strXOR As String
    Dim lngXORIterator As Long
    Dim lngXORMarker As Long
    lngXORMarker = Len(strData) - Len(strMask)
    If lngXORMarker < 0 Then
        strXOR = Left$(strMask, Abs(lngXORMarker))
        strMask = Mid$(strMask, Abs(lngXORMarker) + 1)
    ElseIf lngXORMarker > 0 Then
        strXOR = Left$(strData, Abs(lngXORMarker))
        strData = Mid$(strData, lngXORMarker + 1)
    End If
    For lngXORIterator = 1 To Len(strData)
        strXOR = strXOR + Hex$(Val("&H" + Mid$(strData, lngXORIterator, 1)) Xor Val("&H" + Mid$(strMask, lngXORIterator, 1)))
    Next lngXORIterator
    PerformXOR = Right(strXOR, 8)
End Function
Public Sub SetKey(ByVal strEncryptionKey As String)
    Dim intEncryptionKeyIterator As Integer
    For intEncryptionKeyIterator = 0 To 131 Step 10
        If intEncryptionKeyIterator = 130 Then
            astrEncryptionKey(intEncryptionKeyIterator + 0) = Mid(strEncryptionKey, 1, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 1) = Mid(strEncryptionKey, 3, 2)
        Else
            astrEncryptionKey(intEncryptionKeyIterator + 0) = Mid(strEncryptionKey, 1, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 1) = Mid(strEncryptionKey, 3, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 2) = Mid(strEncryptionKey, 5, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 3) = Mid(strEncryptionKey, 7, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 4) = Mid(strEncryptionKey, 9, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 5) = Mid(strEncryptionKey, 11, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 6) = Mid(strEncryptionKey, 13, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 7) = Mid(strEncryptionKey, 15, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 8) = Mid(strEncryptionKey, 17, 2)
            astrEncryptionKey(intEncryptionKeyIterator + 9) = Mid(strEncryptionKey, 19, 2)
        End If
    Next
End Sub


Public Function BinaryToDouble(ByVal strData As String) As Double
    Dim dblOutput As Double
    Dim lngIterator As Long
    Do Until Len(strData) = 0
        dblOutput = dblOutput + IIf(Right$(strData, 1) = "1", (2 ^ lngIterator), 0)
        strData = Left$(strData, Len(strData) - 1)
        lngIterator = lngIterator + 1
    Loop
    BinaryToDouble = dblOutput
End Function

Public Function DoubleToBinary(ByVal dblData As Double) As String
    Dim strOutput As String
    Dim lngIterator As Long
    Do Until (2 ^ lngIterator) > dblData
        strOutput = IIf(((2 ^ lngIterator) And dblData) > 0, "1", "0") + strOutput
        lngIterator = lngIterator + 1
    Loop
    DoubleToBinary = strOutput
End Function
Public Function HexToString(ByVal strData As String) As String
    Dim strOutput As String
    Do Until Len(strData) < 2
        strOutput$ = strOutput$ + Chr$(CLng("&H" + Left$(strData, 2)))
        strData = Right$(strData, Len(strData) - 2)
    Loop
    HexToString = strOutput
End Function

Public Function StringToHex(ByVal strData As String) As String
    Dim strOutput As String
    Do Until Len(strData) = 0
        strOutput = strOutput + Right$(String$(2, "0") + Hex$(Asc(Left$(strData, 1))), 2)
        strData = Right$(strData, Len(strData) - 1)
    Loop
    StringToHex = strOutput
End Function
Public Function ByteToString(ByRef bytData() As Byte, ByVal lngDataLength As Long) As String
    Dim lngIterator As Long
    For lngIterator = LBound(bytData) To (LBound(bytData) + lngDataLength)
        ByteToString = ByteToString & Chr$(bytData(lngIterator))
    Next lngIterator
End Function


