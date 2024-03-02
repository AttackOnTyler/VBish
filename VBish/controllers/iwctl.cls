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
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)

Private Const WM_ACTIVATE As Long = &H6
Private Const KEYEVENTF_KEYUP As Long = &H2
Private Const VK_CONTROL As Integer = &H11
Private Const ERR_HANDLES_NOT_VALID As String = "Error: Invalid window handles." & vbNewLine & vbNewLine & _
                                                "Run `? ish.start()` to rebind the Immediate Window"

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
    End With
    
    If handlesNotValid Then
        Class_Terminate
        Err.Raise vbObjectError + 513, "iwctl::Class_Initialize", "Failed to bind to the Immediate Window." & vbNewLine & vbNewLine & "Is the Immediate Window open?"
    End If
    
    With this
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
    With this
        .hwndImmediate = 0
        .hwndVBE = 0
        Set .localFileSync = Nothing
        Set .fso = Nothing
    End With
End Sub

Private Function handlesNotValid() As Boolean
    handlesNotValid = this.hwndVBE = 0 Or this.hwndImmediate = 0
End Function

Public Function clear() As String
    If handlesNotValid Then clear = ERR_HANDLES_NOT_VALID: Exit Function
    PostMessageA this.hwndImmediate, WM_ACTIVATE, 1, 0&
    Application.SendKeys "{F16}"
    SendCtrlA
    SendDelete
    Application.SendKeys "{F16}"
End Function

Public Function banner() As String
    If handlesNotValid Then banner = ERR_HANDLES_NOT_VALID: Exit Function
    Debug.Print "{I)mmediate Window (SH)ell" & vbNewLine & _
                "ProcessPioneer Solutions (c)" & vbNewLine & vbNewLine
End Function

Public Function headerbar() As String
    If handlesNotValid Then headerbar = ERR_HANDLES_NOT_VALID: Exit Function
    Debug.Print this.currentwd & ">" & vbNewLine & _
                "? "
    SendBackspace
End Function

Private Function SendCtrlA() As String
    If handlesNotValid Then SendCtrlA = ERR_HANDLES_NOT_VALID: Exit Function
    keybd_event VK_CONTROL, 0, 0, 0
    keybd_event vbKeyA, 0, 0, 0
    keybd_event vbKeyA, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
End Function

Private Function SendDelete() As String
    If handlesNotValid Then SendDelete = ERR_HANDLES_NOT_VALID: Exit Function
    keybd_event vbKeyDelete, 0, 0, 0
    keybd_event vbKeyDelete, 0, KEYEVENTF_KEYUP, 0
End Function

Private Function SendBackspace() As String
    If handlesNotValid Then SendBackspace = ERR_HANDLES_NOT_VALID: Exit Function
    keybd_event vbKeyBack, 0, 0, 0
    keybd_event vbKeyBack, 0, KEYEVENTF_KEYUP, 0
End Function

