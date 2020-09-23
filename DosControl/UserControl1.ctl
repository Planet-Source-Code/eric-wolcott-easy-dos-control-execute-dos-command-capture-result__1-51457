VERSION 5.00
Begin VB.UserControl DOS 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   555
   ScaleHeight     =   570
   ScaleWidth      =   555
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.PictureBox Picture3 
      Height          =   600
      Left            =   0
      Picture         =   "UserControl1.ctx":0312
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   -15
      Width           =   555
   End
End
Attribute VB_Name = "DOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type


Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private mCommand As String
Private mOutputs As String

Public Event OnReceiveOutputs(CommandOutputs As String)
Public Event OnReceiveError(ErrorNumner As Integer, Description As String)

Public Property Let CommandLine(DOSCommand As String)
    mCommand = DOSCommand
End Property

Public Property Get CommandLine() As String
    CommandLine = mCommand
End Property

Public Property Get Outputs()
    Outputs = mOutputs
End Property

Public Function ExecuteCommand(Optional CommandLine As String) As String
    Dim proc As PROCESS_INFORMATION
    Dim ret As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long
    Dim hWritePipe As Long
    Dim lngBytesread As Long
    Dim strBuff As String * 256
    If Len(CommandLine) > 0 Then
        mCommand = CommandLine
    End If
    If Len(mCommand) = 0 Then
        RaiseEvent OnReceiveError(0, "Command Line empty")
        mOutputs = "Command Line empty ERROR # - 0"
        Exit Function
    End If
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If ret = 0 Then
        RaiseEvent OnReceiveError(Err.LastDllError, "CreatePipe failed")
         mOutputs = "CreatePipe failed ERROR # - " & Err.LastDllError
        Exit Function
    End If
    start.cb = Len(start)
    start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
    ret& = CreateProcessA(0&, mCommand, sa, sa, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    If ret <> 1 Then
        RaiseEvent OnReceiveError(1, "File Or Command Not Found!")
        mOutputs = "File Or Command Not Found! ERROR # - 1"
        Exit Function
    End If
    ret = CloseHandle(hWritePipe)
    mOutputs = ""
Dim Tmp As String
    Do
    ret = ReadFile(hReadPipe, strBuff, 256, lngBytesread, 0&)
    mOutputs = mOutputs & Left(strBuff, lngBytesread)
    Tmp = Tmp & Left$(strBuff, lngBytesread)
    Loop While ret <> 0
    RaiseEvent OnReceiveOutputs(Tmp)
    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)
    ExecuteCommand = mOutputs
End Function

Private Sub UserControl_Resize()
UserControl.Width = Picture3.Width
UserControl.Height = Picture3.Height
End Sub
