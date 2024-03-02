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
'
''This Function will convert a OneDrive/SharePoint Url path, e.g. Url containing
''https://d.docs.live.net/; .sharepoint.com/sites; my.sharepoint.com/personal/...
''to the locally synchronized path on your current pc or mac, e.g. a path like
''C:\users\username\OneDrive\ on Windows; or /Users/username/OneDrive/ on MacOS,
''if you have the remote directory locally synchronized with the OneDrive app.
''If no local path can be found, the input value will be returned unmodified.
''Author: Guido Witt-Dörring
''Source: https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
''        https://stackoverflow.com/a/73577057/12287457
'Public Function path(ByVal fullname As String, Optional ByVal returnAll As Boolean = False, Optional ByVal preferredMountPointOwner As String = "", Optional ByVal rebuildCache As Boolean = False) As String
'    Const ps As String = "\"                     'conditional compilation is preferred here.
'    Const isMac As Boolean = False
'    Const methodName As String = "path"
'    Const vbErrFileNotFound                As Long = 53
'    Const vbErrOutOfMemory                 As Long = 7
'    Const vbErrKeyAlreadyExists            As Long = 457
'    Const vbErrInvalidFormatInResourceFile As Long = 325
'
'    Static locToWebColl As Collection, lastCacheUpdate As Date
'
'    If Not Left(fullname, 8) = "https://" Then path = fullname: Exit Function
'
'    Dim webRoot As String, locRoot As String, s As String, vItem As Variant
'    Dim pmpo As String: pmpo = LCase$(preferredMountPointOwner)
'    If Not locToWebColl Is Nothing And Not rebuildCache Then
'        Dim resColl As Collection: Set resColl = New Collection
'        'If the locToWebColl is initialized, this logic will find the local path
'        For Each vItem In locToWebColl
'            locRoot = vItem(0): webRoot = vItem(1)
'            If InStr(1, fullname, webRoot, vbTextCompare) = 1 Then _
'               resColl.Add Key:=vItem(2), _
'               Item:=Replace(Replace(fullname, webRoot, locRoot, , 1), "/", ps)
'        Next vItem
'        If resColl.Count > 0 Then
'            If returnAll Then
'                For Each vItem In resColl: s = s & "//" & vItem: Next vItem
'                path = Mid$(s, 3): Exit Function
'            End If
'            On Error Resume Next: path = resColl(pmpo): On Error GoTo 0
'            If path <> "" Then Exit Function
'            path = resColl(1): Exit Function
'        End If
'        'Local path was not found with cached mountpoints
'        path = fullname                          'No Exit Function here! Check if cache needs rebuild
'    End If
'
'    Dim settPaths As Collection: Set settPaths = New Collection
'    Dim settPath As Variant, clpPath As String
'    settPaths.Add Environ("LOCALAPPDATA") & "\Microsoft\OneDrive\settings\"
'    clpPath = Environ("LOCALAPPDATA") & "\Microsoft\Office\CLP\"
'
'    Dim i As Long
'    'Find all subdirectories in OneDrive settings folder:
'    Dim oneDriveSettDirs As Collection: Set oneDriveSettDirs = New Collection
'    For Each settPath In settPaths
'        Dim dirName As String: dirName = Dir(settPath, vbDirectory)
'        Do Until dirName = vbNullString
'            If dirName = "Personal" Or dirName Like "Business#" Then _
'               oneDriveSettDirs.Add Item:=settPath & dirName & ps
'            dirName = Dir(, vbDirectory)
'        Loop
'    Next settPath
'
'    If Not locToWebColl Is Nothing Or isMac Then
'        Dim requiredFiles As Collection: Set requiredFiles = New Collection
'        'Get collection of all required files
'        Dim vDir As Variant
'        For Each vDir In oneDriveSettDirs
'            Dim cID As String: cID = IIf(vDir Like "*" & ps & "Personal" & ps, _
'                                         "????????????*", _
'                                         "????????-????-????-????-????????????")
'            Dim fileName As String: fileName = Dir(vDir, vbNormal)
'            Do Until fileName = vbNullString
'                If fileName Like cID & ".ini" _
'                   Or fileName Like cID & ".dat" _
'                   Or fileName Like "ClientPolicy*.ini" _
'                   Or StrComp(fileName, "GroupFolders.ini", vbTextCompare) = 0 _
'                   Or StrComp(fileName, "global.ini", vbTextCompare) = 0 _
'                   Or StrComp(fileName, "SyncEngineDatabase.db", _
'                              vbTextCompare) = 0 Then _
'                   requiredFiles.Add Item:=vDir & fileName
'                fileName = Dir
'            Loop
'        Next vDir
'    End If
'
'    'This part should ensure perfect accuracy despite the mount point cache
'    'while sacrificing almost no performance at all by querying FileDateTimes.
'    If Not locToWebColl Is Nothing And Not rebuildCache Then
'        'Check if a settings file was modified since the last cache rebuild
'        Dim vFile As Variant
'        For Each vFile In requiredFiles
'            If FileDateTime(vFile) > lastCacheUpdate Then _
'               rebuildCache = True: Exit For     'full cache refresh is required!
'        Next vFile
'        If Not rebuildCache Then Exit Function
'    End If
'
'    'If execution reaches this point, the cache will be fully rebuilt... so delete everything above and just rebuild
'    Dim fileNum As Long, syncID As String, b() As Byte, j As Long, k As Long
'    'Variables for manual decoding of UTF-8, UTF-32 and ANSI
'    Dim m As Long, ansi() As Byte, sAnsi As String
'    Dim utf16() As Byte, sUtf16 As String, utf32() As Byte
'    Dim utf8() As Byte, sUtf8 As String, numBytesOfCodePoint As Long
'    Dim codepoint As Long, lowSurrogate As Long, highSurrogate As Long
'
'    lastCacheUpdate = Now()
'
'    'Declare some variables that will be used in the loop over OneDrive settings
'    Dim line As Variant, parts() As String, n As Long, libNr As String
'    Dim tag As String, mainMount As String, relPath As String, email As String
'    Dim parentID As String, folderID As String, folderName As String
'    Dim idPattern As String, folderType As String, keyExists As Boolean
'    Dim siteID As String, libID As String, webID As String, lnkID As String
'    Dim mainSyncID As String, syncFind As String, mainSyncFind As String
'    'The following are "constants" and needed for reading the .dat files:
'    Dim sig1 As String:       sig1 = ChrB$(2)
'    Dim sig2 As String * 4:   MidB$(sig2, 1) = ChrB$(1)
'    Dim vbNullByte As String: vbNullByte = ChrB$(0)
'    Const sig3 As String = vbNullChar
'
'    'Writing locToWebColl using .ini and .dat files in the OneDrive settings:
'    'Here, a Scripting.Dictionary would be nice but it is not available on Mac!
'    Dim lastAccountUpdates As Collection, lastAccountUpdate As Date
'    Set lastAccountUpdates = New Collection
'    Set locToWebColl = New Collection
'    For Each vDir In oneDriveSettDirs            'One folder per logged in OD account
'        dirName = Mid$(vDir, InStrRev(vDir, ps, Len(vDir) - 1, 0) + 1)
'        dirName = Left$(dirName, Len(dirName) - 1)
'
'        'Read global.ini to get cid
'        If Dir(vDir & "global.ini", vbNormal) = "" Then GoTo NextFolder
'        fileNum = FreeFile()
'        Open vDir & "global.ini" For Binary Access Read As #fileNum
'        ReDim b(0 To LOF(fileNum)): Get fileNum, , b
'        Close #fileNum: fileNum = 0
'
'        For Each line In Split(b, vbNewLine)
'            If line Like "cid = *" Then cID = Mid$(line, 7): Exit For
'        Next line
'
'        If cID = vbNullString Then GoTo NextFolder
'        If (Dir(vDir & cID & ".ini") = vbNullString Or _
'            (Dir(vDir & "SyncEngineDatabase.db") = vbNullString And _
'             Dir(vDir & cID & ".dat") = vbNullString)) Then GoTo NextFolder
'        If dirName Like "Business#" Then
'            idPattern = Replace(Space$(32), " ", "[a-f0-9]") & "*"
'        ElseIf dirName = "Personal" Then
'            idPattern = Replace(Space$(12), " ", "[A-F0-9]") & "*!###*"
'        End If
'        'Alternatively maybe a general pattern like this performs better:
'        'idPattern = Replace(Space$(12), " ", "[a-fA-F0-9]") & "*"
'
'        'Get email for business accounts
'        '(only necessary to let user choose preferredMountPointOwner)
'        fileName = Dir(clpPath, vbNormal)
'        Do Until fileName = vbNullString
'            i = InStrRev(fileName, cID, , vbTextCompare)
'            If i > 1 And cID <> vbNullString Then _
'               email = LCase$(Left$(fileName, i - 2)): Exit Do
'            fileName = Dir
'        Loop
'
'        'Read all the ClientPloicy*.ini files:
'        Dim cliPolColl As Collection: Set cliPolColl = New Collection
'        fileName = Dir(vDir, vbNormal)
'        Do Until fileName = vbNullString
'            If fileName Like "ClientPolicy*.ini" Then
'                fileNum = FreeFile()
'                Open vDir & fileName For Binary Access Read As #fileNum
'                ReDim b(0 To LOF(fileNum)): Get fileNum, , b
'                Close #fileNum: fileNum = 0
'                cliPolColl.Add Key:=fileName, Item:=New Collection
'                For Each line In Split(b, vbNewLine)
'                    If InStr(1, line, " = ", vbBinaryCompare) Then
'                        tag = Left$(line, InStr(1, line, " = ", 0) - 1)
'                        s = Mid$(line, InStr(1, line, " = ", 0) + 3)
'                        Select Case tag
'                        Case "DavUrlNamespace"
'                            cliPolColl(fileName).Add Key:=tag, Item:=s
'                        Case "SiteID", "IrmLibraryId", "WebID" 'Only used for
'                            s = Replace(LCase$(s), "-", "") 'backup method later
'                            If Len(s) > 3 Then s = Mid$(s, 2, Len(s) - 2)
'                            cliPolColl(fileName).Add Key:=tag, Item:=s
'                        End Select
'                    End If
'                Next line
'            End If
'            fileName = Dir
'        Loop
'
'        'If cid.dat file doesn't exist, skip this part:
'        Dim odFolders As Collection: Set odFolders = Nothing
'        If Dir(vDir & cID & ".dat") = vbNullString Then GoTo Continue
'
'        'Read cid.dat file if it exists:
'        Const chunkOverlap          As Long = 1000
'        Const maxDirName            As Long = 255
'        Dim buffSize As Long: buffSize = -1      'Buffer uninitialized
'Try:            On Error GoTo Catch
'        Set odFolders = New Collection
'        Dim lastChunkEndPos As Long: lastChunkEndPos = 1
'        Dim lastFileUpdate As Date:  lastFileUpdate = FileDateTime(vDir & _
'                                                                   cID & ".dat")
'        i = 0                                    'i = current reading pos.
'        Do
'            'Ensure file is not changed while reading it
'            If FileDateTime(vDir & cID & ".dat") > lastFileUpdate Then GoTo Try
'            fileNum = FreeFile
'            Open vDir & cID & ".dat" For Binary Access Read As #fileNum
'            Dim lenDatFile As Long: lenDatFile = LOF(fileNum)
'            If buffSize = -1 Then buffSize = lenDatFile 'Initialize buffer
'            'Overallocate a bit so read chunks overlap to recognize all dirs
'            ReDim b(0 To buffSize + chunkOverlap)
'            Get fileNum, lastChunkEndPos, b: s = b
'            Dim size As Long: size = LenB(s)
'            Close #fileNum: fileNum = 0
'            lastChunkEndPos = lastChunkEndPos + buffSize
'
'            For vItem = 16 To 8 Step -8
'                i = InStrB(vItem + 1, s, sig2, 0) 'Sarch pattern in cid.dat
'                Do While i > vItem And i < size - 168 'and confirm with another
'                    If StrComp(MidB$(s, i - vItem, 1), sig1, 0) = 0 Then 'one
'                        i = i + 8: n = InStrB(i, s, vbNullByte, 0) - i
'                        If n < 0 Then n = 0      'i:Start pos, n: Length
'                        If n > 39 Then n = 39
'                        folderID = StrConv(MidB$(s, i, n), vbUnicode)
'                        i = i + 39: n = InStrB(i, s, vbNullByte, 0) - i
'                        If n < 0 Then n = 0
'                        If n > 39 Then n = 39
'                        parentID = StrConv(MidB$(s, i, n), vbUnicode)
'                        i = i + 121
'                        n = InStr(-Int(-(i - 1) / 2) + 1, s, sig3, 0) * 2 - i - 1
'                        If n > maxDirName * 2 Then n = maxDirName * 2
'                        If n < 0 Then n = 0
'                        If folderID Like idPattern _
'                           And parentID Like idPattern Then
'                            folderName = MidB$(s, i, n)
'                            'VBA.Array() instead of just Array() is used in this
'                            'function because it ignores Option Base 1
'                            odFolders.Add VBA.Array(parentID, folderName), _
'                                          folderID
'                        End If
'                    End If
'                    i = InStrB(i + 1, s, sig2, 0) 'Find next sig2 in cid.dat
'                Loop
'                If odFolders.Count > 0 Then Exit For
'            Next vItem
'        Loop Until lastChunkEndPos >= lenDatFile _
'             Or buffSize >= lenDatFile
'        GoTo Continue
'Catch:
'        Select Case Err.Number
'        Case vbErrKeyAlreadyExists
'            'This can happen at chunk boundries, folder might get added twice:
'            odFolders.Remove folderID            'Make sure the folder gets added new again
'            Resume                               'to avoid folderNames truncated by chunk ends
'        Case Is <> vbErrOutOfMemory: Err.Raise Err, methodName
'        End Select
'        If buffSize > &HFFFFF Then buffSize = buffSize / 2: Resume Try
'        Err.Raise Err, methodName                'Raise error if less than 1 MB RAM available
'Continue:
'        On Error GoTo 0
'        'If .dat file didn't exist, read db file, otherwise skip this part
'        If Not odFolders Is Nothing Then GoTo SkipDbFile
'        'The following code for reading the .db file is an adaptation of the
'        'original code by Cristian Buse, see procedure 'GetODDirsFromDB' in the
'        'repository: https://github.com/cristianbuse/VBA-FileTools
'        fileNum = FreeFile()
'        Open vDir & "SyncEngineDatabase.db" For Binary Access Read As #fileNum
'        size = LOF(fileNum)
'        If size = 0 Then GoTo CloseFile
'        '                             __    ____
'        'Signature bytes: 0b0b0b0b0b0b080b0b08080b0b0b0b where b>=0, b <= 9
'        Dim sig88 As String: sig88 = ChrW$(&H808)
'        Const sig8 As Long = 8
'        Const sig8Offset As Long = -3
'        Const maxSigByte As Byte = 9
'        Const sig88ToDataOffset As Long = 6      'Data comes after the signature
'        Const headBytes6 As Long = &H16
'        Const headBytes5 As Long = &H15
'        Const headBytes6Offset As Long = -16     'Header comes before the signature
'        Const headBytes5Offset As Long = -15
'        Const chunkSize As Long = &H100000       '1MB
'
'        Dim lastRecord As Long, bytes As Long, nameSize As Long
'        Dim idSize(1 To 4) As Byte
'        Dim lastFolderID As String, lastParentID As String
'        Dim lastNameStart As Long
'        Dim lastNameSize As Long
'        Dim currDataEnd As Long, lastDataEnd As Long
'        Dim headByte As Byte, lastHeadByte As Byte
'        Dim has5HeadBytes As Boolean
'
'        lastFileUpdate = 0
'        ReDim b(1 To chunkSize)
'        Do
'            i = 0
'            If FileDateTime(vDir & "SyncEngineDatabase.db") > lastFileUpdate Then
'                Set odFolders = New Collection
'                Dim heads As Collection: Set heads = New Collection
'
'                lastFileUpdate = FileDateTime(vDir & "SyncEngineDatabase.db")
'                lastRecord = 1
'                lastFolderID = vbNullString
'            End If
'            If LenB(lastFolderID) > 0 Then
'                folderName = MidB$(s, lastNameStart, lastNameSize)
'            End If
'            Get fileNum, lastRecord, b
'            s = b
'            i = InStrB(1 - headBytes6Offset, s, sig88, vbBinaryCompare)
'            lastDataEnd = 0
'            Do While i > 0
'                If i + headBytes6Offset - 2 > lastDataEnd _
'                   And LenB(lastFolderID) > 0 Then
'                    If lastDataEnd > 0 Then
'                        folderName = MidB$(s, lastNameStart, lastNameSize)
'                    End If
'                    sUtf8 = folderName: GoSub DecodeUTF8
'                    folderName = sUtf16
'                    On Error Resume Next
'                    odFolders.Add VBA.Array(lastParentID, folderName), _
'                                  lastFolderID
'                    If Err.Number <> 0 Then
'                        If heads(lastFolderID) < lastHeadByte Then
'                            If odFolders(lastFolderID)(1) <> folderName _
'                               Or odFolders(lastFolderID)(0) <> lastParentID Then
'                                odFolders.Remove lastFolderID
'                                heads.Remove lastFolderID
'                                odFolders.Add VBA.Array(lastParentID, _
'                                                        folderName), _
'                                              lastFolderID
'                            End If
'                        End If
'                    End If
'                    heads.Add lastHeadByte, lastFolderID
'                    On Error GoTo 0
'                    lastFolderID = vbNullString
'                End If
'
'                If b(i + sig8Offset) <> sig8 Then GoTo NextSig
'                has5HeadBytes = True
'                If b(i + headBytes5Offset) = headBytes5 Then
'                    j = i + headBytes5Offset
'                ElseIf b(i + headBytes6Offset) = headBytes6 Then
'                    j = i + headBytes6Offset
'                    has5HeadBytes = False        'Has 6 bytes header
'                ElseIf b(i + headBytes5Offset) <= maxSigByte Then
'                    j = i + headBytes5Offset
'                Else
'                    GoTo NextSig
'                End If
'                headByte = b(j)
'
'                bytes = sig88ToDataOffset
'                For k = 1 To 4
'                    If k = 1 And headByte <= maxSigByte Then
'                        idSize(k) = b(j + 2)     'Ignore first header byte
'                    Else
'                        idSize(k) = b(j + k)
'                    End If
'                    If idSize(k) < 37 Or idSize(k) Mod 2 = 0 Then GoTo NextSig
'                    idSize(k) = (idSize(k) - 13) / 2
'                    bytes = bytes + idSize(k)
'                Next k
'                If has5HeadBytes Then
'                    nameSize = b(j + 5)
'                    If nameSize < 15 Or nameSize Mod 2 = 0 Then GoTo NextSig
'                    nameSize = (nameSize - 13) / 2
'                Else
'                    nameSize = (b(j + 5) - 128) * 64 + (b(j + 6) - 13) / 2
'                    If nameSize < 1 Or b(j + 6) Mod 2 = 0 Then GoTo NextSig
'                End If
'                bytes = bytes + nameSize
'
'                currDataEnd = i + bytes - 1
'                If currDataEnd > chunkSize Then  'Next chunk
'                    i = i - 1
'                    Exit Do
'                End If
'                j = i + sig88ToDataOffset
'                folderID = StrConv(MidB$(s, j, idSize(1)), vbUnicode)
'                j = j + idSize(1)
'                parentID = StrConv(MidB$(s, j, idSize(2)), vbUnicode)
'                parentID = StrConv(MidB$(s, j, idSize(2)), vbUnicode)
'
'                If folderID Like idPattern And parentID Like idPattern Then
'                    lastNameStart = j + idSize(2) + idSize(3) + idSize(4)
'                    lastNameSize = nameSize
'                    lastFolderID = Left(folderID, 32) 'Ignore the "+##.." in IDs
'                    lastParentID = Left(parentID, 32) 'of Business OneDrive
'                    lastHeadByte = headByte
'                    lastDataEnd = currDataEnd
'                End If
'NextSig:
'                i = InStrB(i + 1, s, sig88, vbBinaryCompare)
'            Loop
'            If i = 0 Then
'                lastRecord = lastRecord + chunkSize + headBytes6Offset
'            Else
'                lastRecord = lastRecord + i + headBytes6Offset
'            End If
'        Loop Until lastRecord > size
'CloseFile:
'        Close #fileNum
'SkipDbFile:
'
'        'Read cid.ini file
'        fileNum = FreeFile()
'        Open vDir & cID & ".ini" For Binary Access Read As #fileNum
'        ReDim b(0 To LOF(fileNum)): Get fileNum, , b
'        Close #fileNum: fileNum = 0
'        Select Case True
'        Case dirName Like "Business#"            'Settings files for a business OD account
'            'Max 9 Business OneDrive accounts can be signed in at a time.
'            Dim libNrToWebColl As Collection: Set libNrToWebColl = New Collection
'            mainMount = vbNullString
'            For Each line In Split(b, vbNewLine)
'                webRoot = "": locRoot = "": parts = Split(line, """")
'                Select Case Left$(line, InStr(1, line, " = ", 0) - 1)
'                Case "libraryScope"              'One line per synchronized library
'                    locRoot = parts(9)
'                    syncFind = locRoot: syncID = Split(parts(10), " ")(2)
'                    libNr = Split(line, " ")(2)
'                    folderType = parts(3): parts = Split(parts(8), " ")
'                    siteID = parts(1): webID = parts(2): libID = parts(3)
'                    If mainMount = vbNullString Or folderType = "ODB" Then
'                        mainMount = locRoot: fileName = "ClientPolicy.ini"
'                        mainSyncID = syncID: mainSyncFind = syncFind
'                    Else: fileName = "ClientPolicy_" & libID & siteID & ".ini"
'                    End If
'                    On Error Resume Next         'On error try backup method...
'                    webRoot = cliPolColl(fileName)("DavUrlNamespace")
'                    On Error GoTo 0
'                    If webRoot = "" Then         'Backup method to find webRoot:
'                        For Each vItem In cliPolColl
'                            If vItem("SiteID") = siteID _
'                               And vItem("WebID") = webID _
'                               And vItem("IrmLibraryId") = libID Then
'                                webRoot = vItem("DavUrlNamespace"): Exit For
'                            End If
'                        Next vItem
'                    End If
'                    If webRoot = vbNullString Then Err.Raise vbErrFileNotFound _
'                       , methodName
'                    libNrToWebColl.Add VBA.Array(libNr, webRoot), libNr
'                    If Not locRoot = vbNullString Then _
'                       locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
'                                                  syncID, syncFind, dirName), Key:=locRoot
'                Case "libraryFolder"             'One line per synchronized library folder
'                    libNr = Split(line, " ")(3)
'                    locRoot = parts(1): syncFind = locRoot
'                    syncID = Split(parts(4), " ")(1)
'                    s = vbNullString: parentID = Left$(Split(line, " ")(4), 32)
'                    Do                           'If not synced at the bottom dir of the library:
'                        '   -> add folders below mount point to webRoot
'                        On Error Resume Next: odFolders parentID
'                        keyExists = (Err.Number = 0): On Error GoTo 0
'                        If Not keyExists Then Exit Do
'                        s = odFolders(parentID)(1) & "/" & s
'                        parentID = odFolders(parentID)(0)
'                    Loop
'                    webRoot = libNrToWebColl(libNr)(1) & s
'                    locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
'                                               syncID, syncFind, dirName), locRoot
'                Case "AddedScope"                'One line per folder added as link to personal
'                    relPath = parts(5): If relPath = " " Then relPath = "" 'lib
'                    parts = Split(parts(4), " "): siteID = parts(1)
'                    webID = parts(2): libID = parts(3): lnkID = parts(4)
'                    fileName = "ClientPolicy_" & libID & siteID & lnkID & ".ini"
'                    On Error Resume Next         'On error try backup method...
'                    webRoot = cliPolColl(fileName)("DavUrlNamespace") & relPath
'                    On Error GoTo 0
'                    If webRoot = "" Then         'Backup method to find webRoot:
'                        For Each vItem In cliPolColl
'                            If vItem("SiteID") = siteID _
'                               And vItem("WebID") = webID _
'                               And vItem("IrmLibraryId") = libID Then
'                                webRoot = vItem("DavUrlNamespace") & relPath
'                                Exit For
'                            End If
'                        Next vItem
'                    End If
'                    If webRoot = vbNullString Then Err.Raise vbErrFileNotFound _
'                       , methodName
'                    s = vbNullString: parentID = Left$(Split(line, " ")(3), 32)
'                    Do                           'If link is not at the bottom of the personal library:
'                        On Error Resume Next: odFolders parentID
'                        keyExists = (Err.Number = 0): On Error GoTo 0
'                        If Not keyExists Then Exit Do 'add folders below
'                        s = odFolders(parentID)(1) & ps & s 'mount point to
'                        parentID = odFolders(parentID)(0) 'locRoot
'                    Loop
'                    locRoot = mainMount & ps & s
'                    locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
'                                               mainSyncID, mainSyncFind, dirName), locRoot
'                Case Else: Exit For
'                End Select
'            Next line
'        Case dirName = "Personal"                'Settings files for a personal OD account
'            'Only one Personal OneDrive account can be signed in at a time.
'            For Each line In Split(b, vbNewLine) 'Loop should exit at first line
'                If line Like "library = *" Then
'                    parts = Split(line, """"): locRoot = parts(3)
'                    syncFind = locRoot: syncID = Split(parts(4), " ")(2)
'                    Exit For
'                End If
'            Next line
'            On Error Resume Next                 'This file may be missing if the personal OD
'            webRoot = cliPolColl("ClientPolicy.ini")("DavUrlNamespace") 'account
'            On Error GoTo 0                      'was logged out of the OneDrive app
'            If locRoot = "" Or webRoot = "" Or cID = "" Then GoTo NextFolder
'            locToWebColl.Add VBA.Array(locRoot, webRoot & "/" & cID, email, _
'                                       syncID, syncFind, dirName), Key:=locRoot
'            If Dir(vDir & "GroupFolders.ini") = "" Then GoTo NextFolder
'            'Read GroupFolders.ini file
'            cID = vbNullString: fileNum = FreeFile()
'            Open vDir & "GroupFolders.ini" For Binary Access Read As #fileNum
'            ReDim b(0 To LOF(fileNum)): Get fileNum, , b
'            Close #fileNum: fileNum = 0       'Two lines per synced folder from other peoples personal ODs
'            For Each line In Split(b, vbNewLine)
'                If line Like "*_BaseUri = *" And cID = vbNullString Then
'                    cID = LCase$(Mid$(line, InStrRev(line, "/", , 0) + 1, _
'                                      InStrRev(line, "!", , 0) - InStrRev(line, "/", , 0) - 1))
'                    folderID = Left$(line, InStr(1, line, "_", 0) - 1)
'                ElseIf cID <> vbNullString Then
'                    locToWebColl.Add VBA.Array(locRoot & ps & odFolders( _
'                                               folderID)(1), webRoot & "/" & cID & "/" & _
'                                                            Mid$(line, Len(folderID) + 9), email, _
'                                               syncID, syncFind, dirName), _
'                                               Key:=locRoot & ps & odFolders(folderID)(1)
'                    cID = vbNullString: folderID = vbNullString
'                End If
'            Next line
'        End Select
'NextFolder:
'        cID = vbNullString: s = vbNullString: email = vbNullString
'    Next vDir
'
'    'Clean the finished "dictionary" up, remove trailing "\" and "/"
'    Dim tmpColl As Collection: Set tmpColl = New Collection
'    For Each vItem In locToWebColl
'        locRoot = vItem(0): webRoot = vItem(1): syncFind = vItem(4)
'        If Right$(webRoot, 1) = "/" Then _
'           webRoot = Left$(webRoot, Len(webRoot) - 1)
'        If Right$(locRoot, 1) = ps Then _
'           locRoot = Left$(locRoot, Len(locRoot) - 1)
'        If Right$(syncFind, 1) = ps Then _
'           syncFind = Left$(syncFind, Len(syncFind) - 1)
'        tmpColl.Add VBA.Array(locRoot, webRoot, vItem(2), _
'                              vItem(3), syncFind), locRoot
'    Next vItem
'    Set locToWebColl = tmpColl
'
'    path = path(fullname, returnAll, pmpo, False): Exit Function
'    Exit Function
'DecodeUTF8:                                      'UTF-8 must be transcoded to UTF-16 manually in VBA
'    Const raiseErrors As Boolean = False         'Raise error if invalid UTF-8 is found?
'    Dim o As Long, p As Long, q As Long
'    Static numBytesOfCodePoints(0 To 255) As Byte
'    Static mask(2 To 4) As Long
'    Static minCp(2 To 4) As Long
'
'    If numBytesOfCodePoints(0) = 0 Then
'        For o = &H0& To &H7F&: numBytesOfCodePoints(o) = 1: Next o '0xxxxxxx
'        '110xxxxx - C0 and C1 are invalid (overlong encoding)
'        For o = &HC2& To &HDF&: numBytesOfCodePoints(o) = 2: Next o
'        For o = &HE0& To &HEF&: numBytesOfCodePoints(o) = 3: Next o '1110xxxx
'        '11110xxx - 11110100, 11110101+ (= &HF5+) outside of valid Unicode range
'        For o = &HF0& To &HF4&: numBytesOfCodePoints(o) = 4: Next o
'        For o = 2 To 4: mask(o) = (2 ^ (7 - o) - 1): Next o
'        minCp(2) = &H80&: minCp(3) = &H800&: minCp(4) = &H10000
'    End If
'    Dim currByte As Byte
'    utf8 = sUtf8
'    ReDim utf16(0 To (UBound(utf8) - LBound(utf8) + 1) * 2)
'    p = 0
'    o = LBound(utf8)
'    Do While o <= UBound(utf8)
'        codepoint = utf8(o)
'        numBytesOfCodePoint = numBytesOfCodePoints(codepoint)
'        If numBytesOfCodePoint = 0 Then
'            If raiseErrors Then Err.Raise 5
'            GoTo insertErrChar
'        ElseIf numBytesOfCodePoint = 1 Then
'            utf16(p) = codepoint
'            p = p + 2
'        ElseIf o + numBytesOfCodePoint - 1 > UBound(utf8) Then
'            If raiseErrors Then Err.Raise 5
'            GoTo insertErrChar
'        Else
'            codepoint = utf8(o) And mask(numBytesOfCodePoint)
'            For q = 1 To numBytesOfCodePoint - 1
'                currByte = utf8(o + q)
'                If (currByte And &HC0&) = &H80& Then
'                    codepoint = (codepoint * &H40&) + (currByte And &H3F)
'                Else
'                    If raiseErrors Then _
'                       Err.Raise 5
'                    GoTo insertErrChar
'                End If
'            Next q
'            'Convert the Unicode codepoint to UTF-16LE bytes
'            If codepoint < minCp(numBytesOfCodePoint) Then
'                If raiseErrors Then Err.Raise 5
'                GoTo insertErrChar
'            ElseIf codepoint < &HD800& Then
'                utf16(p) = CByte(codepoint And &HFF&)
'                utf16(p + 1) = CByte(codepoint \ &H100&)
'                p = p + 2
'            ElseIf codepoint < &HE000& Then
'                If raiseErrors Then Err.Raise 5
'                GoTo insertErrChar
'            ElseIf codepoint < &H10000 Then
'                If codepoint = &HFEFF& Then GoTo nextCp '(BOM - will be ignored)
'                utf16(p) = codepoint And &HFF&
'                utf16(p + 1) = codepoint \ &H100&
'                p = p + 2
'            ElseIf codepoint < &H110000 Then     'Calculate surrogate pair
'                m = codepoint - &H10000
'                Dim loSurrogate As Long: loSurrogate = &HDC00& Or (m And &H3FF)
'                Dim hiSurrogate As Long: hiSurrogate = &HD800& Or (m \ &H400&)
'                utf16(p) = hiSurrogate And &HFF&
'                utf16(p + 1) = hiSurrogate \ &H100&
'                utf16(p + 2) = loSurrogate And &HFF&
'                utf16(p + 3) = loSurrogate \ &H100&
'                p = p + 4
'            Else
'                If raiseErrors Then Err.Raise 5
'insertErrChar:                  utf16(p) = &HFD
'                utf16(p + 1) = &HFF
'                p = p + 2
'                If numBytesOfCodePoint = 0 Then numBytesOfCodePoint = 1
'            End If
'        End If
'nextCp:         o = o + numBytesOfCodePoint      'Move to the next UTF-8 codepoint
'    Loop
'    sUtf16 = MidB$(utf16, 1, p)
'    Return
'
'DecodeANSI:                                      'Code for decoding ANSI string manually:
'    ansi = sAnsi
'    p = UBound(ansi) - LBound(ansi) + 1
'    If p > 0 Then
'        ReDim utf16(0 To p * 2 - 1): q = 0
'        For p = LBound(ansi) To UBound(ansi)
'            utf16(q) = ansi(p): q = q + 2
'        Next p
'        sUtf16 = utf16
'    Else
'        sUtf16 = vbNullString
'    End If
'    Return
'End Function


