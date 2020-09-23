Attribute VB_Name = "modMain"
Global Port As Long
Global Oncount As Integer
Global lstchange As Boolean
Global CliChange As Boolean
Global GrpChange As Boolean
Global CliGrpChange As Boolean
Global BinChange As Boolean
Global lstHidAdd As Boolean
Global lstHidRemove As Boolean
Global spy_client As Boolean
Global MaxClients As Integer
Global TransferBufferSize As Long
Global ClientCounter As Long
Global ConnectedClients As Long
Global ServerActive As Boolean
Global uUser As Integer
Global Pcnt As Integer
Global file_is_open As Boolean
Global create_dir As String
Global requested As Boolean
Global DummyS As String
Global turn As String
Global drivesvisi As Boolean
Global spot1 As Integer
Global spot2 As String
Global spot3 As String
Global hidDir_list As String
Global found As Boolean
Global use_bin As Boolean
Global del_path As String
Global halt_transfer As Boolean

Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
       Global Const HELP_CONTEXT = &H1
       Global Const HELP_QUIT = &H2
       Global Const HELP_INDEX = &H3
       Global Const HELP_HELPONHELP = &H4
       Global Const HELP_SETINDEX = &H5
       Global Const HELP_KEY = &H101
       Global Const HELP_MULTIKEY = &H201
       Global vHelp&

Type Priv
  Path As String
  Accs As String
                 
End Type
Global Privs(20) As Priv
Global aItem As Integer
Global uItem As Integer

Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal _
    lpFileName As String, lpFindFileData _
    As WIN32_FIND_DATA) _
    As Long

Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal _
    hFindFile As Long, lpFindFileData _
    As WIN32_FIND_DATA) _
    As Long

Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime _
    As FileTime, lpSystemTime _
    As SYSTEMTIME) _
    As Long

Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Declare Function FindClose Lib "kernel32" (ByVal _
    hFindFile As Long) _
    As Long

Global Const MAX_PATH = 260

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Long
End Type

Global Const MAX_IDLE_TIME = 900

Global Const MAX_CONNECTIONS = 50

Global client(MAX_CONNECTIONS) As ftpClient

Enum ClientStatus
    stat_IDLE = 0
    stat_LOGGING_IN = 1
    stat_GETTING_DIR_LIST = 2
    stat_UPLOADING = 3
    stat_DOWNLOADING = 4
End Enum

Enum ConnectModes
    cMode_NORMAL = 0
    cMode_PASV = 1
End Enum

Type ftpClient
    Current_Access As String
    Password As String
    inUse As Boolean                'Identifies if this slot is being used.
    user_ID As Integer
    group_ID As Integer
    Group_Name As String
    group_Access As String
    Priv(20) As Privtyp
    Pcnt As Integer
    id As Long                      'Unique number to identify a client.
    UserName As String              'User name client is is logged in as.
    IPAddress As String             'IP address of the client.
    DataPort As Long                'Port number open on the client for the server to connect to.
    ConnectedAt As String           'Time the client first connected.
    IdleSince As String             'Last recorded time the client sent a command to the server.
    TotalBytesUploaded As Long      'Total bytes uploaded by client from the current session.
    TotalBytesDownloaded As Long    'Total bytes downloaded by client from the current session.
    TotalFilesUploaded As Long      'Total files uploaded by client from the current session.
    TotalFilesDownloaded As Long    'Total files downloaded by client from the current session.
    CurrentFile As String           'Current file being transfer, if any.
    cFileTotalBytes As Long         'Total number of bytes of the file being transfered.
    cTotalBytesXfer As Long         'Total bytes of the current file that has been transfered.
    fFile As Long                   'Reference number to an open file on the server, if any.
    ConnectMode As ConnectModes     'If the client uses PASV mode or not.
    HomeDir As String               'Initial directory client starts in when they first connect.
    CurrentDir As String            'Current directory.
    full_file_name As String
    Status As ClientStatus          'What the client is currently doing.
End Type

Private Type BrowseInfo
       hWndOwner As Long
       pIDLRoot As Long
       pszDisplayName As Long
       lpszTitle As Long
       ulFlags As Long
       lpfnCallback As Long
       lParam As Long
       iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
       (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" _
       (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
       (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Function Privledges(Group_Name As String) As String
Dim Ctr As Integer
If grpnum > 0 Then
   For Ctr = 1 To grpnum
   If group_info.GrpType(Ctr).Name = Group_Name Then
   Privledges = group_info.GrpType(Ctr).Access
   End If
   Next
End If
End Function
Public Function Convert_to_Word(Letter As String) As String
Select Case Letter
  Case "T"
  Convert_to_Word = "Transfer"
  Case "W"
  Convert_to_Word = "Write"
  Case "D"
  Convert_to_Word = "Delete Files"
  Case "X"
  Convert_to_Word = "Execute"
  Case "L"
  Convert_to_Word = "List Files"
  Case "M"
  Convert_to_Word = "Make Directory"
  Case "H"
  Convert_to_Word = "Remove Directory"
  Case "S"
  Convert_to_Word = "Sub Directories"
  Case "R"
  Convert_to_Word = "Show Home Relative"
  Case "B"
  Convert_to_Word = "Show All Drives"
  Case "Q"
  Convert_to_Word = "Group Account Disabled"
  Case "P"
  Convert_to_Word = "No Restrictions"
  Case "A"
  Convert_to_Word = "Time Limit"
  Case Else
  MsgBox "     Forgot a letter !    " & Letter
  End Select
End Function
Public Function open_file(file_path As String, Socket As Integer)
Dim intt As Integer
intt = Socket
'Open file_path For Binary As #5
Open file_path For Binary As intt
client(Socket).fFile = intt
End Function
Public Function Replace(sIn As String, sFind As String, sReplace As String, Optional lStart As Long = 1, _
    Optional iCount As Long = -1, Optional bCompare As VbCompareMethod = vbBinaryCompare) As String
    Dim lC As Long, iPos As Integer, sOut As String
    sOut = sIn
    iPos = InStr(lStart, sOut, sFind, bCompare)
    If iPos = 0 Then GoTo EndFn:


    Do
        lC = lC + 1
        sOut = Left(sOut, iPos - 1) & sReplace & Mid(sOut, iPos + Len(sFind))
        If iCount <> -1 And nC >= iCount Then Exit Do
        iPos = InStr(lStart, sOut, sFind, bCompare)
    Loop While iPos > 0
EndFn:
    Replace = sOut
End Function
Function Parse2Array(ByVal strText As String, ByRef strArray() As String, ByVal strDelim As String) As Long
       Dim intPos As Long
       Dim intIndex As Long
       strText = Trim(strText)
       ReDim strArray(10) As String
       Do While strText <> ""
           If intIndex > UBound(strArray()) Then
               ReDim Preserve strArray(intIndex + 20)
           End If
           intPos = InStr(1, strText, strDelim)
           If intPos > 0 Then
               strArray(intIndex) = Left(strText, InStr(1, strText, strDelim) - 1)
               strText = Trim(Mid(strText, InStr(1, strText, strDelim) + 1))
           Else
               strArray(intIndex) = strText
               Exit Do
           End If
           intIndex = intIndex + 1
       Loop
       ReDim Preserve strArray(intIndex) As String
       Parse2Array = UBound(strArray())
   End Function
Public Function InStrRev(String1 As String, String2 As String) As Integer
    Dim pos As Integer
    Dim pos2 As Integer
    Let pos2 = Len(String1)


    Do
        Let pos = (InStr(pos2, String1, String2))
        Let pos2 = pos2 - 1
    Loop Until pos > 0 Or pos2 = 0
    Let InStrRev = pos
End Function
Sub Wait(WaitSeconds As Single)

Dim StartTime As Single

StartTime = Timer

Do While Timer < StartTime + WaitSeconds
DoEvents
Loop
End Sub
Public Function Perm(Num As Integer) As String
Select Case Num
  Case "16"
  Perm = "drwxr-xr-x "
  Case "3"
  Perm = "-r-xr-xr-x "
  Case "38"
  Perm = "srwxr-xr-x "
  Case "32"
  Perm = "-rwxr-xr-x "
  Case "7"
  Perm = "sr-xr-xr-x "
  Case "2"
  Perm = "-rwxr-xr-x "
  Case "6"
  Perm = "srwxr-xr-x "
  Case "22"
  Perm = "drwxr-xr-x "
  Case "34"
  Perm = "-rwxr-xr-x "
End Select
End Function
Function CountStr(ByVal parseStringx, Parser As String) As Variant
On Error Resume Next
Dim lastPos As Integer
Dim subPos As Integer
Dim argPos(1 To 2000) As Integer
Dim argContent(1 To 2000)
parsestring = parseStringx
parsestring = Trim(Right(parsestring, ((Len(parsestring)) - (InStr(parsestring, Parser)))))

parsestring = parsestring & Parser 'vbCrLf
argcount = 0
Do
    DoEvents
    lastPos = InStr((lastPos + 1), parsestring, Parser)
    If lastPos = 0 Then Exit Do
    argcount = argcount + 1
    argPos(argcount) = lastPos
Loop
If argcount = 0 Then Exit Function
CountStr = argcount
End Function
Function Parse(ByVal parseStringx, ByVal argNum As Integer) As Variant
On Error Resume Next
Dim lastPos As Integer
Dim subPos As Integer
Dim argPos(1 To 2000) As Integer
Dim argContent(1 To 2000)
parsestring = parseStringx
parsestring = Trim(Right(parsestring, ((Len(parsestring)) - (InStr(parsestring, " ")))))

parsestring = parsestring & " "
argcount = 0
Do
    DoEvents
    lastPos = InStr((lastPos + 1), parsestring, " ")
    If lastPos = 0 Then Exit Do
    argcount = argcount + 1
    argPos(argcount) = lastPos
Loop
If argcount = 0 Then Exit Function
For i = 1 To argcount
    Select Case i
        Case argcount
            If argcount <> 1 Then
                subPos = argPos(i - 1)
            Else
                subPos = 1
            End If
        Case 1
            subPos = 1
        Case Else
            subPos = argPos(i - 1)
    End Select
    DoEvents
    argContent(i) = Trim(Mid(parsestring, subPos, (argPos(i) - subPos)))
Next i
Parse = argContent(argNum)
End Function
Public Function CreateDefault()
Open (App.Path & "\ftp_srv.ini") For Output As #11
    Print #11, "[Settings]"
    Print #11, "Version=1.1.2"
    Print #11, "DefaultSetDate=" & Date
    Print #11, "DefaultSetTime=" & Time
    Print #11, ""
    Print #11, "[Common]"
    Print #11, "Port=21"
    Print #11, "Anonomous=Yes"
    Print #11, "DenyAll=No"
    Print #11, "Maximum=10"
    Print #11, "SeeALL=Yes"
    Print #11, "BeepAttempt=Yes"
    Print #11, "BeepDelete=Yes"
    Print #11, "UseHidden=Yes"
    Print #11, "Hidden=1"
    Print #11, "Hid1=windows"
    Print #11, ""
    Print #11, "[Users]"
    Print #11, "Groups=3"
    Print #11, "Users=3"
    Print #11, "Name1=boss"
    Print #11, "Pass1=itsme"
    Print #11, "Group1=Administrator"
    Print #11, "Group2=Guest"
    Print #11, "Group3=Visitor"
    Print #11, "DirCnt1=4"
    Print #11, "Home1=c:\"
    Print #11, "Access1_1=c:\,WDLSTMH"
    Print #11, "Access1_2=d:\,WDLSTMH"
    Print #11, "Access1_3=e:\,WDLSTMH"
    Print #11, "Access1_4=f:\,WDLSTMH"
    Print #11, "Name2=user"
    Print #11, "Pass2=testing"
    Print #11, "DirCnt2=2"
    Print #11, "Home2=d:\Temp\FileSrv"
    Print #11, "Access2_1=d:\Temp\FileSrv,WSL"
    Print #11, "Access2_2=d:\Temp\FileSrv\Shared,WSL"
    Print #11, "Name3=visitor"
    Print #11, "Pass3=umm"
    Print #11, "DirCnt3=1"
    Print #11, "Home3=d:\Temp\FileSrv\Shared\Doug"
    Print #11, "Access3_1=d:\Temp\FileSrv\Shared\Doug,LT"
    Print #11, "Group1Dis=No"
    Print #11, "Group2Dis=No"
    Print #11, "Group3Dis=No"
    Print #11, "GrAcc1=BP"
    Print #11, "GrAcc2=B"
    Print #11, "GrAcc3=R"
    Print #11, "GrpName1=Administrator"
    Print #11, "GrpName2=Guest"
    Print #11, "GrpName3=Visitor"
    Close #11
    MsgBox " - Default Settings Restored - " & vbCrLf & "You will need to restart the application" & vbCrLf & "for the changes to take effect."
End Function
Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String

       Dim iNull As Integer
       Dim lpIDList As Long
       Dim lResult As Long
       Dim sPath As String
       Dim udtBI As BrowseInfo
       With udtBI
       .hWndOwner = hWndOwner
       .lpszTitle = lstrcat(sPrompt, "")
       .ulFlags = BIF_RETURNONLYFSDIRS
End With

lpIDList = SHBrowseForFolder(udtBI)

If lpIDList Then
       sPath = String$(MAX_PATH, 0)
       lResult = SHGetPathFromIDList(lpIDList, sPath)
       Call CoTaskMemFree(lpIDList)
       iNull = InStr(sPath, vbNullChar)

              If iNull Then
                     sPath = Left$(sPath, iNull - 1)
              End If

End If

BrowseForFolder = sPath
End Function
Function PathParts(ByVal sPath As String, ByRef sParts() As String) As Long

       '     'Returns 0 if parts can be extracted, Err if not.
       'Assigns substrings of sPath to sParts() in the following ma
       '     nner:
       '     'sPart(1)Substring to the right of the last backslash
       '     '(filename or final directory)
       '     'sPart(2)Substring to the right of the last period
       '     '(extension)
       'sPart(3)Substring from beginning to last backslash, inclusi
       '     ve
       '     '(full directory path)
       'sPart(4)Substring between second to last backslash and last
       '      backslash
       '     '(final directory name)
       'sPart(4) is equal to sPart(3) if only one backslash exists
       '     '(drive root directory).
       '     'SAMPLE CALL
       '     'Sub Button1_Click()
       '     'Dim l as long
       '     'Dim Parts() as String
       '     'l = PathParts("c:\windows\system\comctl32.dll",Parts())
       '     'End Sub
       '     'SAMPLE CALL RETURNS
       '     'Parts(1) "comctl32.dll"
       '     'Parts(2) "dll"
       '     'Parts(3) "c:\windows\system"
       '     'Parts(4) "system"
       On Error GoTo PathParts_Err
       Dim i As Integer
       Dim iPeriod As Integer
       Dim iSlash(2) As Integer
       Dim iLen As Integer
       ReDim sParts(4)
       sPath = Trim(sPath)
       iLen = Len(sPath)

              For i = iLen To 1 Step -1

                            If Mid(sPath, i, 1) = "\" Then
                                   iSlash(IIf(Not iSlash(1), 1, 2)) = i
                            ElseIf Mid(sPath, i, 1) = "." And Not iPeriod Then
                                   iPeriod = i
                            End If

              Next i

       sParts(1) = Right(sPath, Len(sPath) - iSlash(1))

              If iPeriod Then sParts(2) = Right(sPath, Len(sPath) - iPeriod)
                     sParts(3) = Left(sPath, iSlash(1))
                     iSlash(2) = iSlash(2) + 1

                            If iSlash(2) > 1 Then
                                   sParts(4) = Mid(sPath, iSlash(2), iSlash(1) - iSlash(2))
                            Else
                                   sParts(4) = sParts(3)
                            End If

                     Exit Function
PathParts_Err:
                     PathParts = Err
              End Function


