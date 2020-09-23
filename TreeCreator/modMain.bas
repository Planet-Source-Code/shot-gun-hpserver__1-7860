Attribute VB_Name = "modMain"
Global selected As Boolean
Global tFile As String
Global Sel_dir As String
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
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
       (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" _
       (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
       (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Sub Wait(WaitSeconds As Single)

Dim StartTime As Single

StartTime = Timer

Do While Timer < StartTime + WaitSeconds
DoEvents
Loop
End Sub

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
Sel_dir = sPath
If Right(Sel_dir, 1) <> "\" Then
Sel_dir = Sel_dir & "\"
End If
End Function

Function Create(Path As String)
On Error GoTo air
If Right(Path, 1) <> "\" Then
Path = Path & "\"
End If
MkDir Path & "Homeplay\"
MkDir Path & "Homeplay\Shared\"
MkDir Path & "Homeplay\Shared\Paul\"
MkDir Path & "Homeplay\Shared\Paul\Transfer\"
MkDir Path & "Homeplay\Shared\Doug\"
MkDir Path & "Homeplay\Shared\Brian\"
MkDir Path & "Homeplay\Shared\Susan\"
MkDir Path & "Homeplay\Shared\Games\"
MkDir Path & "Homeplay\Shared\Games\Good\"
MkDir Path & "Homeplay\Shared\Games\Better\"
MkDir Path & "Homeplay\Shared\Apps\"
MkDir Path & "Homeplay\Shared\Apps\Free\"
MkDir Path & "Homeplay\Shared\Apps\Free\Pictures\"
MkDir Path & "Homeplay\Shared\Apps\Sample\"
MkDir Path & "Homeplay\Shared\Apps\Sample\Mine\"
MkDir Path & "Homeplay\Shared\Music\"
MkDir Path & "Homeplay\Shared\Music\Classical\"
MkDir Path & "Homeplay\Shared\Music\Rock\"
MkDir Path & "Homeplay\Shared\System\"
MkDir Path & "Homeplay\Shared\System\Drivers\"
frmMain.lblMessage.Caption = "Directories Created"
Call Wait(0.75)

Exit Function
air:
frmMain.lblMessage.Caption = " - Error - "
MsgBox "There was an Error, the directories may already exist. You should check."
End Function
Function Create_File(Path As String, tmpfile As String) As Boolean
If tmpfile = False Then
Open (App.Path & "\ftp_srv.ini") For Output As #11
Else
Open (tFile) For Output As #11
End If
If Right(Path, 1) <> "\" Then
Path = Path & "\"
End If

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
    Print #11, "Groups=4"
    Print #11, "Users=5"
    Print #11, "Name1=boss"
    Print #11, "Pass1=itsme"
    Print #11, "Group1=Administrator"
    Print #11, "Group2=Guest"
    Print #11, "Group3=Visitor"
    Print #11, "Group4=Brother"
    Print #11, "DirCnt1=4"
    Print #11, "Home1=c:\"
    Print #11, "Access1_1=c:\,WDLSTMH"
    Print #11, "Access1_2=d:\,WDLSTMH"
    Print #11, "Access1_3=e:\,WDLSTMH"
    Print #11, "Access1_4=f:\,WDLSTMH"
    Print #11, "Name2=Brian"
    Print #11, "Pass2=friend"
    Print #11, "DirCnt2=5"
    Print #11, "Home2=" & Path & "Homeplay"
    Print #11, "Access2_1=" & Path & "Homeplay\Shared" & ",WSL"
    Print #11, "Access2_2=" & Path & "Homeplay\Shared\Apps" & ",WSL"
    Print #11, "Access2_3=" & Path & "Homeplay\Shared\Apps\Free" & ",WSL"
    Print #11, "Access2_4=" & Path & "Homeplay\Shared\Apps\Free\Pictures" & ",WSLT"
    Print #11, "Access2_5=" & Path & "Homeplay\Shared\Brian" & ",WSLT"
    Print #11, "Name3=Paul"
    Print #11, "Pass3=bro"
    Print #11, "DirCnt3=8"
    Print #11, "Home3=" & Path & "Homeplay"
    Print #11, "Access3_1=" & Path & "Homeplay\Shared" & ",WSL"
    Print #11, "Access3_2=" & Path & "Homeplay\Shared\Paul" & ",WSLT"
    Print #11, "Access3_3=" & Path & "Homeplay\Shared\Games" & ",WSL"
    Print #11, "Access3_4=" & Path & "Homeplay\Shared\Games\Better" & ",WSLT"
    Print #11, "Access3_5=" & Path & "Homeplay\Shared\Paul\Transfer" & ",WSLT"
    Print #11, "Access3_6=" & Path & "Homeplay\Shared\Apps" & ",WSL"
    Print #11, "Access3_7=" & Path & "Homeplay\Shared\Apps\Free" & ",WSL"
    Print #11, "Access3_8=" & Path & "Homeplay\Shared\Apps\Free\Pictures" & ",WSLT"
    Print #11, "Name4=Susan"
    Print #11, "Pass4=sweet"
    Print #11, "DirCnt4=6"
    Print #11, "Home4=" & Path & "Homeplay"
    Print #11, "Access4_1=" & Path & "Homeplay\Shared" & ",WSL"
    Print #11, "Access4_2=" & Path & "Homeplay\Shared\Susan" & ",WSLT"
    Print #11, "Access4_3=" & Path & "Homeplay\Shared\Games" & ",WSL"
    Print #11, "Access4_4=" & Path & "Homeplay\Shared\Games\Good" & ",WSLT"
    Print #11, "Access4_5=" & Path & "Homeplay\Shared\Apps" & ",WSL"
    Print #11, "Access4_6=" & Path & "Homeplay\Shared\Apps\Free" & ",WSLT"
    Print #11, "Name5=visitor"
    Print #11, "Pass5=umm"
    Print #11, "DirCnt5=4"
    Print #11, "Home5=" & Path & "Homeplay"
    Print #11, "Access5_1=" & Path & "Homeplay\Shared" & ",WSL"
    Print #11, "Access5_2=" & Path & "Homeplay\Shared\Apps" & ",WSL"
    Print #11, "Access5_3=" & Path & "Homeplay\Shared\Games" & ",WSL"
    Print #11, "Access5_4=" & Path & "Homeplay\Shared\Games\Better" & ",WSLT"
    Print #11, "Group1Dis=No"
    Print #11, "Group2Dis=No"
    Print #11, "Group3Dis=No"
    Print #11, "GrAcc1=BP"
    Print #11, "GrAcc2=B"
    Print #11, "GrAcc3=B"
    Print #11, "GrAcc4=R"
    Print #11, "GrpName1=Administrator"
    Print #11, "GrpName2=Guest"
    Print #11, "GrpName3=Guest"
    Print #11, "GrpName4=Guest"
    Print #11, "GrpName5=Visitor"
    
    Close #11
    If tmpfile = False Then
frmMain.lblMessage.Caption = "Fake Files being Created"
Call Wait(0.75)
    End If
    If tmpfile = True Then
frmMain.lblMessage.Caption = "New ftp_srv.ini created in apps directory"
Call Wait(0.75)
    End If

End Function
Function FakeFile()
Dim a As String
Dim b As String, c As String, d As String, e As String, f As String, g As String, h As String
a = Sel_dir & "Homeplay\Shared\Apps\apps.txt"
b = Sel_dir & "Homeplay\Shared\Games\games.txt"
c = Sel_dir & "Homeplay\Shared\Games\Better\better.txt"
d = Sel_dir & "Homeplay\Shared\Games\Good\good.txt"
e = Sel_dir & "Homeplay\Shared\Paul\Transfer\PaulTrs.txt"
f = Sel_dir & "Homeplay\Shared\Apps\Free\Pictures\Pics.txt"
g = Sel_dir & "Homeplay\Shared\Susan\Susan.txt"
h = Sel_dir & "Homeplay\Shared\Brian\brian.txt"
FileCopy tFile, a
FileCopy tFile, b
FileCopy tFile, c
FileCopy tFile, d
FileCopy tFile, e
FileCopy tFile, f
FileCopy tFile, g
FileCopy tFile, h
frmMain.lblMessage.Caption = "Fake Files Moved to Directories"
Call Wait(0.75)
Kill tFile
frmMain.lblMessage.Caption = "Temporary File Deleted"
Call Wait(0.75)

End Function
