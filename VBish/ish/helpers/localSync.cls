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

Public Enum storeType
    localFS
    onedrive
    sharepoint
    unknown
End Enum

Public Function path(ByVal fullname As String) As String
    If Len(fullname) = 0 Then Exit Function 'value passed in matches the default return of path()
    
    Dim storage As storeType, startSlashPos As Integer
    storage = getStoreType(fullname, startSlashPos)
    
    Select Case storage
        Case storeType.localFS
            path = fullname
        Case storeType.onedrive
            path = Environ("onedrive") & Replace(Mid(fullname, startSlashPos), "/", "\")
        Case storeType.sharepoint
            
    End Select
    
End Function

Public Function getStoreType(ByVal fullname As String, ByRef startSlashPos As Integer) As storeType 'default return is localFS
    If Not Left(fullname, 8) = "https://" Then Exit Function 'startSlashPos is 0 by default and getStoreType is set to localFS by default
    
    Dim storage As storeType
    If InStr(1, fullname, "d.docs.live.net") <> 0 Then storage = onedrive
    If InStr(1, fullname, ".sharepoint.com") <> 0 Then storage = sharepoint
    If Not CBool(storage) Then Err.Raise 5, "localSync::getStoreType", "Failed to find host for file to map local path.": getStoreType = unknown: Exit Function
    
    getStoreType = storage
    
    Select Case storage
        Case storeType.onedrive
            startSlashPos = InStr(25, fullname, "/")
        Case storeType.sharepoint
            Err.Raise 5, "localSync::getStoreType", "Sharepoint not implemented"
    End Select
    
End Function
