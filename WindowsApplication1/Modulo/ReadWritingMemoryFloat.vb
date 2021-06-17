Module Module1
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Integer, ByRef lpdwProcessId As Integer) As Integer
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
    Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddressAsAny As Integer, ByRef lpBufferAsAny As Integer, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
    Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddressAsAny As Integer, ByRef lpBufferAsAny As Integer, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParamAsAny As Integer) As Integer
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Private Declare Sub ReleaseCapture Lib "user32" ()
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Integer, ByVal lpAddress As Integer, ByVal dwSize As Integer, ByVal flAllocationType As Integer, ByVal flProtect As Integer) As Integer
    Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Integer, ByRef lpAddressAsAny As Integer, ByVal dwSize As Integer, ByVal dwFreeType As Integer) As Integer
    Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Integer, ByRef lpThreadAttributes As Integer, ByVal dwStackSize As Integer, ByRef lpStartAddress As Integer, ByRef lpParameterAsAny As Integer, ByVal dwCreationFlags As Integer, ByRef lpThreadId As Integer) As Integer
    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
    Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Integer, ByVal lProcessID As Integer) As Integer
    Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Integer, ByRef uProcess As MODULEENTRY32) As Integer
    Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Integer, ByRef uProcess As MODULEENTRY32) As Integer
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Integer
    Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Integer, ByVal lpProcName As String) As Integer
    Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Integer
    Public Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal Key As Integer) As Short
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer

    '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Some constants used through the code:
    Const PROCESS_ALL_ACCESS As Integer = &H1F0FFF
    Const MEM_COMMIT As Short = &H1000
    Const MEM_RELEASE As Integer = &H8000
    Const PAGE_READWRITE As Short = &H4
    Const HTCAPTION As Short = 2
    Const MOUSEEVENTF_LEFTUP As Short = &H4
    Const WM_KEYDOWN As Short = &H100
    Const WM_KEYUP As Short = &H101
    Const WM_NCLBUTTONDOWN As Short = &HA1
    Const MOUSEEVENTF_LEFTDOWN As Short = &H2


    Structure MODULEENTRY32
        Dim dwSize As Integer
        Dim th32ModuleID As Integer
        Dim th32ProcessID As Integer
        Dim GlblcntUsage As Integer
        Dim ProccntUsage As Integer
        Dim modBaseAddr As Integer
        Dim modBaseSize As Integer
        Dim hModule As Integer
        <VBFixedString(256)> Dim szModule As Char
        <VBFixedString(260)> Dim szExePath As Char
    End Structure

    Const TH32CS_SNAPPROCESS As Short = &H2
    Const TH32CS_SNAPMODULE As Short = &H8

    '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'FEW ASM OPCODES HEX VALUES, USEFUL FOR CODE INJECTION:
    '*remove if not wanted*
    Public Const PUSH_BYTE As Short = &H6A
    Public Const PUSH_DWORD As Short = &H68
    Public Const PUSHAD As Short = &H60
    Public Const PUSH_EAX As Short = &H50
    Public Const PUSH_ECX As Short = &H51
    Public Const PUSH_EDX As Short = &H52
    Public Const PUSH_EBX As Short = &H53
    Public Const PUSH_ESP As Short = &H54
    Public Const PUSH_EBP As Short = &H55
    Public Const PUSH_ESI As Short = &H56
    Public Const PUSH_EDI As Short = &H57
    Public Const POPAD As Short = &H61
    Public Const POP_EAX As Short = &H58
    Public Const POP_ECX As Short = &H59
    Public Const POP_EDX As Short = &H5A
    Public Const POP_EBX As Short = &H5B
    Public Const POP_ESP As Short = &H5C
    Public Const POP_EBP As Short = &H5D
    Public Const POP_ESI As Short = &H5E
    Public Const POP_EDI As Short = &H5F
    Public Const CALL_FUNC As Short = &HE8
    Public Const RETN As Short = &HC3
    Public Const ADD_ESP As Integer = &HC483
    Public Const JMP_LONG As Short = &HE9
    Public Const JMP_SHORT As Short = &HEB
    Public Const PUSH_DWORD_PTR As Short = &H35FF

    '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Setup before using module:
    'EX: CurrentProcess = "Delta Force 1.00.03.03P" 'DF1
    'EX: CurrentProcess = "Delta Force 2,  V1.06.15" 'DF2
    'EX: CurrentProcess = "Delta Force,  V1.5.0.5" 'BHD
    'EX: CurrentProcess = "Jedi Knight®: Jedi Academy (MP)" 'JKA
    'EX: CurrentProcess = "Delta Force Land Warrior, Demo V0.99.49"' LW_Demo
    'EX: CurrentProcess = "Call of Duty 4"
    Public CurrentProcess As String = "THE HOUSE OF THE DEAD 3"


    'Compatible with BHD Windowed.
    'Public Sub InjectForm(FormToInject As Form, strProgramCaption As String)
    '#Const defUse_InjectForm = True
    Public Sub InjectForm(ByVal FormToInject As Form)

        Dim hWnd As Integer
        hWnd = FindWindow(vbNullString, CurrentProcess)

        If hWnd <> 0 Then
            SetParent(FormToInject.Handle.ToInt32, hWnd)

        End If
    End Sub

    Public Sub WriteFloat(ByVal lngAddress As Integer, ByVal sngValue As Single, Optional ByRef NumberOfBytesWritten As Short = 0)
        Dim hWnd As Object, processHandle As Object, processId As Integer

        hWnd = FindWindow(vbNullString, CurrentProcess)
        If (hWnd = 0) Then Exit Sub

        GetWindowThreadProcessId(hWnd, processId)
        processHandle = OpenProcess(PROCESS_ALL_ACCESS, False, processId)
        WriteProcessMemory(processHandle, lngAddress, sngValue, 4, NumberOfBytesWritten)
        CloseHandle(processHandle)
    End Sub
    Public Function ReadFloat(ByVal lngAddress As Integer) As Single
        Dim hWnd As Object, processHandle As Object, processId As Integer

        hWnd = FindWindow(vbNullString, CurrentProcess)
        If (hWnd = 0) Then Exit Function

        GetWindowThreadProcessId(hWnd, processId)
        processHandle = OpenProcess(PROCESS_ALL_ACCESS, False, processId)
        ReadProcessMemory(processHandle, lngAddress, ReadFloat, 4, 0)
        CloseHandle(processHandle)
        Return 0
    End Function
End Module
