Attribute VB_Name = "Module1"
Option Explicit

'------------------ Privileges ------------------------
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_QUERY As Long = &H8
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const ANYSIZE_ARRAY As Long = 1
Private Type mudtLUID
    LowPart As Long
    HighPart As Long
End Type
Private Type mudtLUID_AND_ATTRIBUTES
    pLuid As mudtLUID
    Attributes As Long
End Type
Private Type mudtTOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As mudtLUID_AND_ATTRIBUTES
End Type
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As mudtLUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As mudtTOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As mudtTOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
'-------------------------------------------------------


Public Sub SetDebugPrivileges()
    Dim hPr As Long, hToken As Long, lr As Long
    Dim udtLUID As mudtLUID
    Dim udtPriv As mudtTOKEN_PRIVILEGES, udtNewPriv As mudtTOKEN_PRIVILEGES

    hPr = GetCurrentProcess()
    OpenProcessToken hPr, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeDebugPrivilege", udtLUID
    udtPriv.PrivilegeCount = 1
    udtPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    udtPriv.Privileges(0).pLuid = udtLUID
    lr = AdjustTokenPrivileges(hToken, False, udtPriv, 4 + (12 * udtPriv.PrivilegeCount), udtNewPriv, 4 + (12 * udtNewPriv.PrivilegeCount))
End Sub



