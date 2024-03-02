VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iwctl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("controllers")
Option Explicit

Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function PostMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function SendMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare PtrSafe Function SendInput Lib "user32.dll" (ByVal nInputs As Long, ByRef pInputs As Any, ByVal cbSize As Long) As Long
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)

Private Type KeyboardInput
    dwType As Long
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    dwTime As Long
    dwExtraInfo As LongPtr
    dwPadding As Currency
End Type

Private Const WM_ACTIVATE As Long = &H6
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const INPUT_KEYBOARD As Long = &H1
Private Const KEYEVENTF_KEYUP As Long = &H2
Private Const VK_CONTROL As Integer = &H11
Private Const VK_A As Integer = &H41
Private Const VK_DELETE As Integer = &H2E

Private Type TImmediateWindowManager
    hwndVBE As LongPtr
    hwndImmediate As LongPtr
    currentwd As String
    previouswd As String
    localFileSync As localSync
    fso As Scripting.FileSystemObject
End Type

Private this As TImmediateWindowManager

Private Sub Class_Initialize()
    With this
        .hwndVBE = FindWindowA("wndclass_desked_gsk", vbNullString)
        .hwndImmediate = GetHWNDImmediate(.hwndVBE)
        Set .localFileSync = New localSync
        Set .fso = New Scripting.FileSystemObject
        .currentwd = .fso.GetParentFolderName(.localFileSync.path(ThisWorkbook.fullname))
        .previouswd = .currentwd
    End With
End Sub

Private Function GetHWNDImmediate(ByVal hwndVBE As LongPtr) As LongPtr
    Select Case ThisWorkbook.Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    Case MsoLanguageID.msoLanguageIDEnglishUS
        GetHWNDImmediate = FindWindowExA(hwndVBE, ByVal 0&, "VbaWindow", "Immediate")
    Case MsoLanguageID.msoLanguageIDGerman
        GetHWNDImmediate = FindWindowExA(hwndVBE, ByVal 0&, "VbaWindow", "Direct")
    End Select
End Function

Private Sub Class_Terminate()
    this.hwndImmediate = 0
    this.hwndVBE = 0
End Sub

Public Sub clear()
    PostMessageA this.hwndImmediate, WM_ACTIVATE, 1, 0&
    Application.SendKeys "{F16}"
    SendCtrlA
    SendDelete
    Application.SendKeys "{F16}"
End Sub

Public Sub banner()
    Debug.Print "{I)mmediate Window (SH)ell" & vbNewLine & _
                "ProcessPioneer Solutions (c)" & vbNewLine & vbNewLine
End Sub

Public Sub headerbar()
    Debug.Print this.currentwd & ">" & vbNewLine & _
                "? "
    SendBackspace
End Sub

Private Sub SendCtrlA()

    keybd_event VK_CONTROL, 0, 0, 0
    keybd_event vbKeyA, 0, 0, 0
    keybd_event vbKeyA, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
   
End Sub

Private Sub SendDelete()
    keybd_event vbKeyDelete, 0, 0, 0
    keybd_event vbKeyDelete, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub SendBackspace()
    keybd_event vbKeyBack, 0, 0, 0
    keybd_event vbKeyBack, 0, KEYEVENTF_KEYUP, 0
End Sub

