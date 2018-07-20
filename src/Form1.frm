VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Перехватчик вывода консольных окон"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4680
      TabIndex        =   23
      Top             =   4800
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   6135
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Text            =   "100"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Text            =   "20"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Отключиться от окна"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Искать консольные окна"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label18 
      Caption         =   "Путь к процессу:"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "0"
      Height          =   255
      Left            =   8760
      TabIndex        =   21
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "Codepage:"
      Height          =   255
      Left            =   7800
      TabIndex        =   20
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   255
      Left            =   7080
      TabIndex        =   19
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "0"
      Height          =   255
      Left            =   7080
      TabIndex        =   18
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "Позиция курсора Y:"
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Позиция курсора X:"
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Буфер экрана Y:"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Буфер экрана X:"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label6 
      Caption         =   "Заголовок:"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Частота перехвата текста                мс."
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Частота поиска окон                         мс."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "ProcessID:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Handle:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Статус"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH As Long = 260&

Private Type COORD
    X As Integer
    Y As Integer
End Type
Private Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type
Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function AttachConsole Lib "kernel32" (ByVal dwProcessId As Long) As Boolean
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function ReadConsoleOutputCharacter Lib "kernel32" Alias "ReadConsoleOutputCharacterW" (ByVal hConsoleOutput As Long, ByVal lpCharacter As Long, ByVal nLength As Long, ByVal dwReadCoord As Long, lpNumberOfCharsRead As Long) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Private Declare Function GetConsoleCP Lib "kernel32" () As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal revert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameW" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, lpFilePart As Long) As Long
Private Declare Function QueryFullProcessImageName Lib "kernel32.dll" Alias "QueryFullProcessImageNameW" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As Long, ByVal lpdwSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function QueryDosDevice Lib "kernel32.dll" Alias "QueryDosDeviceW" (ByVal lpDeviceName As Long, ByVal lpTargetPath As Long, ByVal ucchMax As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const STD_OUTPUT_HANDLE         As Long = -11

Private cH          As Long
Private hOut        As Long
Private isAttached  As Boolean

Private Sub PrintHandleAndPID(hwnd As Long, PID As Long)
    Label2.Caption = "Foreground Handle: " & hwnd
    Label3.Caption = "Foreground PID:      " & PID
End Sub

Function GetPIDbyWindowHandle(hwnd As Long)
    Dim hThread     As Long
    Dim PID         As Long
    
    hThread = GetWindowThreadProcessId(ByVal hwnd, PID)
    GetPIDbyWindowHandle = PID
End Function

Private Sub Form_Load()
    Dim PID     As Long
    Form1.Caption = "Console Window Interceptor - STOP"
    cH = GetForegroundWindow()
    PID = GetPIDbyWindowHandle(cH)
    Call PrintHandleAndPID(cH, PID)
End Sub

Private Sub Command1_Click()
    Form1.Caption = "Console Window Interceptor - WATCH..."
    Text1.Text = "" 'перехваченный текст
    Label7.Caption = "" 'заголовок
    Timer1.Interval = CLng(Text2.Text)
    Timer1.Enabled = True
    Command1.Enabled = False
End Sub

Private Sub Command2_Click()
    Dim lr      As Long
    Label1.Caption = ""
    Form1.Caption = "Console Window Interceptor - STOP"
    Timer2.Enabled = False
    Sleep (500) ' дать время на завершение перехвата текста консоли (Timer2)
    If isAttached Then
        lr = FreeConsole()
        isAttached = False
    End If
    Command1.Enabled = True
End Sub

Private Sub Timer1_Timer() ' отслеживание активного окна
    Dim hwnd    As Long
    Dim lr      As Long
    Dim PID     As Long
    
    hwnd = GetForegroundWindow()
    If cH <> hwnd Then ' если активное окно изменилось
        cH = hwnd
        PID = GetPIDbyWindowHandle(hwnd)
        Call PrintHandleAndPID(hwnd, PID)
        
        If GetClassNameByHandle(hwnd) = "ConsoleWindowClass" Then ' если окно консольное
            lr = AttachConsole(PID)
            If lr <> 0 Then
                hOut = GetStdHandle(STD_OUTPUT_HANDLE)
                Label1.Caption = "Attach Successful"
                Timer1.Enabled = False
                Timer2.Interval = CLng(Text3.Text)
                Timer2.Enabled = True
                isAttached = True
                RemoveCross hwnd                                    ' Делаю неактивной кнопку закрытия окна
                Label7.Caption = GetTitle()                         ' Получаю заголовок окна
                Text4.Text = GetProcessNameByPID(PID)               ' Получаю путь к процессу
            Else
                Label1.Caption = "Error: " & Err.LastDllError
            End If
        End If
    End If
End Sub

Private Sub Timer2_Timer() ' актуализация перехвата текста консоли
    Text1.Text = GetConsoleText()
    Label17.Caption = GetConsoleCP()
End Sub

Public Function GetConsoleText() As String
    Dim infBuf  As CONSOLE_SCREEN_BUFFER_INFO
    Dim count   As Long
    Dim buf     As String
    Dim i       As Long
    If Not isAttached Then Exit Function
    GetConsoleScreenBufferInfo hOut, infBuf
    Form1.Label10.Caption = infBuf.dwSize.X
    Form1.Label11.Caption = infBuf.dwSize.Y
    Form1.Label14.Caption = infBuf.dwCursorPosition.X
    Form1.Label15.Caption = infBuf.dwCursorPosition.Y
    count = infBuf.dwSize.X 'infBuf.dwSize And &HFFFF&
    buf = Space(count)
    For i = 0 To infBuf.dwCursorPosition.Y '(infBuf.dwSize \ &H10000) And &HFFFF&
        ReadConsoleOutputCharacter hOut, StrPtr(buf), count, i * &H10000, count
        GetConsoleText = GetConsoleText & RTrim(buf) & vbNewLine
    Next
End Function

Function GetClassNameByHandle(hwnd As Long) As String
        Dim nMaxCount As Long: nMaxCount = 256
        Dim lpClassName As String: lpClassName = Space(nMaxCount)
        Dim lresult As Long: lresult = GetClassName(hwnd, lpClassName, nMaxCount)
        If lresult <> 0 Then GetClassNameByHandle = Left$(lpClassName, lresult)
End Function

Sub RemoveCross(hwnd As Long)
    ' Делаю неактивной кнопку закрытия окна
    Const SC_CLOSE          As Long = &HF060
    Const MF_BYCOMMAND      As Long = &H0
    Dim hMenu               As Long
    hMenu = GetSystemMenu(hwnd, False)
    Call RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
End Sub

Function GetTitle() As String
    Dim nRet    As Long
    Dim Title   As String
    Title = Space$(1024)
    nRet = GetConsoleTitle(Title, Len(Title))
    If nRet Then GetTitle = Left$(Title, nRet)
End Function

'Function GetProcessNameByPID(PID As Long) As String
'    Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
'    Const PROCESS_QUERY_INFORMATION         As Long = &H400
'    Const MAX_PATH                          As Long = 260
'    Dim hProc               As Long
'    Dim Path                As String
'    Dim lStr                As Long
'    Dim inf(68)             As Long
'    Dim IsVistaAndLater     As Boolean
'    inf(0) = 276: GetVersionEx inf(0): IsVistaAndLater = inf(1) >= 6
'    hProc = OpenProcess(IIf(IsVistaAndLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), False, PID)
'    If hProc <> 0 Then 'INVALID_HANDLE_VALUE Then
'        lStr = MAX_PATH
'        Path = Space(lStr)
'        ' minimum Windows Vista !!!!!!!
'        If QueryFullProcessImageName(hProc, 0, StrPtr(Path), lStr) Then
'            GetProcessNameByPID = Left$(Path, lStr)
'        End If
'        CloseHandle hProc
'    End If
'End Function


Function GetProcessNameByPID(PID As Long) As String
    On Error GoTo ErrorHandler:

    Const MAX_PATH_W                        As Long = 32767&
    Const PROCESS_VM_READ                   As Long = 16&
    Const PROCESS_QUERY_INFORMATION         As Long = 1024&
    Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&
    Const ERROR_ACCESS_DENIED               As Long = 5&
    Const ERROR_PARTIAL_COPY                As Long = 299&
    
    Dim ProcPath    As String
    Dim hProc       As Long
    Dim cnt         As Long
    Dim pos         As Long
    Dim FullPath    As String
    Dim SizeOfPath  As Long
    Dim lpFilePart  As Long
    Dim IsVistaAndLater     As Boolean
    Dim inf(68)             As Long
    Dim sWinDir             As String
    Dim lr                  As Long
    
    sWinDir = Space$(MAX_PATH)
    lr = GetWindowsDirectory(StrPtr(sWinDir), MAX_PATH)
    If lr Then
        sWinDir = Left$(sWinDir, lr)
    Else
        sWinDir = Environ("SystemRoot")
    End If

    inf(0) = 276: GetVersionEx inf(0): IsVistaAndLater = inf(1) >= 6

    hProc = OpenProcess(IIf(IsVistaAndLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0&, PID)
    
    If hProc = 0 Then
        If Err.LastDllError = ERROR_ACCESS_DENIED Then
            hProc = OpenProcess(IIf(IsVistaAndLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0&, PID)
        End If
    End If
    
    If hProc <> 0 Then
    
        If IsVistaAndLater Then
            cnt = MAX_PATH_W + 1
            ProcPath = Space$(cnt)
            Call QueryFullProcessImageName(hProc, 0&, StrPtr(ProcPath), VarPtr(cnt))
        End If
        
        If 0 <> Err.LastDllError Or Not IsVistaAndLater Then     'Win 2008 Server (x64) can cause Error 128 if path contains space characters
        
            ProcPath = Space$(MAX_PATH)
            cnt = GetModuleFileNameEx(hProc, 0&, StrPtr(ProcPath), Len(ProcPath))
        
            If cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
                ProcPath = Space$(MAX_PATH_W)
                cnt = GetModuleFileNameEx(hProc, 0&, StrPtr(ProcPath), Len(ProcPath))
            End If
        End If
        
        If cnt <> 0 Then                          'clear path
            ProcPath = Left$(ProcPath, cnt)
            If StrComp("\SystemRoot\", Left$(ProcPath, 12), 1) = 0 Then ProcPath = sWinDir & Mid$(ProcPath, 12)
            If "\??\" = Left$(ProcPath, 4) Then ProcPath = Mid$(ProcPath, 5)
        End If
        
        If ERROR_PARTIAL_COPY = Err.LastDllError Or cnt = 0 Then     'because GetModuleFileNameEx cannot access to that information for 64-bit processes on WOW64
            ProcPath = Space$(MAX_PATH)
            cnt = GetProcessImageFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
            
            If cnt <> 0 Then
                ProcPath = Left$(ProcPath, cnt)
                
                ' Convert DosDevice format to Disk drive format
                If StrComp(Left$(ProcPath, 8), "\Device\", 1) = 0 Then
                    pos = InStr(9, ProcPath, "\")
                    If pos <> 0 Then
                        FullPath = ConvertDosDeviceToDriveName(Left$(ProcPath, pos - 1))
                        If Len(FullPath) <> 0 Then
                            ProcPath = FullPath & Mid$(ProcPath, pos + 1)
                        End If
                    End If
                End If
                
            End If
            
        End If
        
        If cnt <> 0 Then    'if process ran with 8.3 style, GetModuleFileNameEx will return 8.3 style on x64 and full pathname on x86
                            'so wee need to expand it ourself
        
            FullPath = Space$(MAX_PATH)
            SizeOfPath = GetFullPathName(StrPtr(ProcPath), MAX_PATH, StrPtr(FullPath), lpFilePart)
            If SizeOfPath <> 0& Then
                GetProcessNameByPID = Left$(FullPath, SizeOfPath)
            Else
                GetProcessNameByPID = ProcPath
            End If
            
        End If
        
        CloseHandle hProc
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print Err, "GetFilePathByPID"
End Function

Public Function ConvertDosDeviceToDriveName(inDosDeviceName As String) As String
    On Error GoTo ErrorHandler:

    Static DosDevices   As New Collection
    
    If DosDevices.count Then
        ConvertDosDeviceToDriveName = DosDevices(inDosDeviceName)
        Exit Function
    End If
    
    Dim aDrive()        As String
    Dim sDrives         As String
    Dim cnt             As Long
    Dim i               As Long
    Dim DosDeviceName   As String
    
    cnt = GetLogicalDriveStrings(0&, StrPtr(sDrives))
    
    sDrives = Space(cnt)
    
    cnt = GetLogicalDriveStrings(Len(sDrives), StrPtr(sDrives))

    If 0 = Err.LastDllError Then
    
        aDrive = Split(Left$(sDrives, cnt - 1), vbNullChar)
    
        For i = 0 To UBound(aDrive)
            
            DosDeviceName = Space(MAX_PATH)
            
            cnt = QueryDosDevice(StrPtr(Left$(aDrive(i), 2)), StrPtr(DosDeviceName), Len(DosDeviceName))
            
            If cnt <> 0 Then
            
                DosDeviceName = Left$(DosDeviceName, InStr(DosDeviceName, vbNullChar) - 1)

                DosDevices.Add aDrive(i), DosDeviceName

            End If
            
        Next
    
    End If
    
    ConvertDosDeviceToDriveName = DosDevices(inDosDeviceName)
    Exit Function
ErrorHandler:
    Debug.Print Err, "ConvertDosDeviceToDriveName"
End Function
