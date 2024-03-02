VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "localSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ish.helpers")
Option Explicit

Public Function path(ByVal fullname As String) As String
    If Not Left(fullname, 8) = "https://" Then path = fullname: Exit Function
    
    Dim i As Integer
    For i = 1 To Len(fullname)
        i = InStr(i, fullname, "/")
        Select Case True
            Case Mid(fullname, i - 4, 4) = ".net"
                i = i + 1
                i = InStr(i, fullname, "/")
                Exit For
        End Select
    Next i
    
    path = Environ("onedrive") & Replace(Mid(fullname, i), "/", "\")
End Function
