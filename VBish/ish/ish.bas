Attribute VB_Name = "ish"
'@Folder("ish")
Option Explicit

Private iwController As iwctl

Public Function start() As String
    On Error GoTo ErrorHandler
    
    If Not iwController Is Nothing Then Application.OnTime Now, "ish.headerbar": Exit Function
    
    Set iwController = New iwctl
    iwController.clear
    Application.OnTime Now, "ish.banner"
    
    On Error GoTo 0
    Exit Function
    
ErrorHandler:
    start = "Error initializing ish: " & Err.Description
    On Error GoTo 0
End Function

Private Sub banner()
    On Error Resume Next
    iwController.banner
    iwController.headerbar
    On Error GoTo 0
End Sub

Private Sub headerbar()
    On Error Resume Next
    iwController.headerbar
    On Error GoTo 0
End Sub

Public Function esc() As String
    If iwController Is Nothing Then GoTo ErrorHandler
    
    iwController.clear
    Set iwController = Nothing
    
    Exit Function
ErrorHandler:
    esc = "Failed to escape ish:" & vbNewLine & _
          "Have you ran `? ish.start()`?"
End Function

Public Function cls() As String
    If iwController Is Nothing Then GoTo ErrorHandler
    
    iwController.clear
    Application.OnTime Now, "ish.headerbar"
    
    Exit Function
ErrorHandler:
    cls = "Failed to clear screen:" & vbNewLine & _
          "Have you ran `? ish.start()`?"
End Function

Public Function clear() As String
    clear = cls()
End Function
