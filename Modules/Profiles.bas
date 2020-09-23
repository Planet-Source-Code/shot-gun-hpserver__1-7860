Attribute VB_Name = "Profiles"
Option Explicit

Global Const MAX_N_USERS = 25
Global Const N_RECOGNIZED_USERS = 8
Global Const DEFAULT_DRIVE = "D:\Temp\FileSrv"

Type Privtyp
  Path As String
  Accs As String
                 
End Type
Global Privtyp As Privtyp

Type Group
 Access As String
 Number As Integer
 Name As String
 Disabled As String
 Restrictions As Boolean
 Relative As Boolean
 BrwsDrives As Boolean
 Count As Integer
End Type
Global Group As Group

Type group_info
 GrpType(10) As Group
End Type
Global group_info As group_info

Type UserInfo
  Hide As String
  Name As String
  Pass As String
  Spy As Boolean
  id As Integer
  IP As String
  Pcnt As Integer
  Group As String
  Priv(20) As Privtyp
  Home As String
  Port As Integer
  local_file As String
  remote_file As String
  file_size As Long
  Disabled As String
End Type

Type User_IDs
  Hide As String
  Count As Integer
  No(0 To MAX_N_USERS) As UserInfo
End Type

Global UserIDs As User_IDs

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName _
    As String) As Integer

Declare Function WritePrivateProfileString% Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal _
    lpFileName$)
Global Version As Integer
Global CurrentProfile As String
Dim v As Integer, mgn As String
Global i As Integer, Number As Integer, grpnum As Integer
Public Function LoadProfile(ByVal Filename As String) As Boolean
  Dim tStr As String
  Dim Ctr As Integer, X As Integer, Pcnt As Integer
  Dim tst As String
  On Error Resume Next
  Ctr = FileLen(Filename)
  If Err.Number > 0 Then
    Err.Clear
    LoadProfile = False
    Exit Function
  End If
  On Error Resume Next
  LoadProfile = True
  If Ctr < 1 Then
    Exit Function
  End If
  
  Version = Val(GetFromIni("Settings", "Version", Filename))
  If Len(Version) < 1 Then
    LoadProfile = False
    Exit Function
  End If
  
  With frmMain
   .Grps.Clear
   .lstHidden.Clear
   .User2.Clear
   .UserList.Clear
   .BelongGrp.Text = ""
   .UsrName = ""
   .PWord = ""
  End With
  
  Number = Val(GetFromIni("Users", "Users", Filename))
  grpnum = Val(GetFromIni("Users", "Groups", Filename))
  
  Group.Number = grpnum
  UserIDs.Count = Number
  
  If grpnum > 0 Then
   For Ctr = 1 To grpnum
   group_info.GrpType(Ctr).Name = GetFromIni("Users", "Group" & Ctr, Filename)
   frmMain.Grps.AddItem GetFromIni("Users", "Group" & Ctr, Filename)
   group_info.GrpType(Ctr).Number = Ctr
   group_info.GrpType(Ctr).Access = GetFromIni("Users", "GrAcc" & Ctr, Filename)
   Next
  End If
  
  If Number > 0 Then
    For Ctr = 1 To Number
      UserIDs.No(Ctr).Name = GetFromIni("Users", "Name" & Ctr, Filename)
      UserIDs.No(Ctr).Pass = GetFromIni("Users", "Pass" & Ctr, Filename)
      UserIDs.No(Ctr).Group = GetFromIni("Users", "GrpName" & Ctr, Filename)
      UserIDs.No(Ctr).Disabled = GetFromIni("Users", "Group" & Ctr & "Dis", Filename)
      UserIDs.No(Ctr).Spy = False
      Pcnt = Val(GetFromIni("Users", "DirCnt" & Ctr, Filename))
      UserIDs.No(Ctr).Pcnt = Pcnt
      Debug.Print "User:" & Ctr & ", DirCnt=" & Pcnt
      
      For X = 1 To Pcnt
        tStr = GetFromIni("Users", "Access" & Ctr & "_" & X, Filename)
        i = InStr(tStr, ",")
        UserIDs.No(Ctr).Priv(X).Path = Left(tStr, i - 1)
        UserIDs.No(Ctr).Priv(X).Accs = Right(tStr, (Len(tStr) - i))
      Next
      
      UserIDs.No(Ctr).Home = GetFromIni("Users", "Home" & Ctr, Filename)
    Next
    
  End If
  
  Dim hStr As Integer
  frmMain.lstHidden.Clear
  hStr = GetFromIni("Common", "Hidden", Filename)
  For X = 1 To hStr
        tStr = GetFromIni("Common", "Hid" & X, Filename)
        frmMain.lstHidden.AddItem tStr
      Next
      
      hStr = GetFromIni("Common", "BeepAttempt", Filename)
      If hStr = "Yes" Then
      frmMain.chkAccess.Value = 1
      End If
      
      hStr = GetFromIni("Common", "BeepDelete", Filename)
      If hStr = "Yes" Then
      frmMain.chkDelete.Value = 1
      End If
      
      hStr = GetFromIni("Common", "UseHidden", Filename)
      If hStr = "Yes" Then
      frmMain.chkHidden.Value = 1
      End If
      
      hStr = GetFromIni("Common", "SeeALL", Filename)
      If hStr = "Yes" Then
      frmMain.chkAdmin.Value = 1
      End If
      
  
  Dim hws As String
  
  frmMain.LisPort = (GetFromIni("Common", "Port", Filename))
  frmMain.maxUnits = (GetFromIni("Common", "Maximum", Filename))
  hws = (GetFromIni("Common", "Anonomous", Filename))
  If hws = "Yes" Then
  frmMain.AllowAnon.Value = 1
  Else
  frmMain.AllowAnon.Value = 0
  End If
  hws = (GetFromIni("Common", "DenyAll", Filename))
  If hws = "Yes" Then
  frmMain.DenAll.Value = 1
  Else
  frmMain.DenAll.Value = 0
  End If
  
  CurrentProfile = Filename
  
End Function
Public Function SaveProfile(ByVal Filename As String, SaveSettings As Boolean) As Boolean
  Dim Terminal As String, Alias As String
  Dim Ctr As Integer, X As Integer, h As String
  SaveProfile = False
  If SaveSettings Then
    If WritePrivateProfileString("Settings", "Version", _
        App.Major & "." & App.Minor & "." & App.Revision, Filename) = 0 Then
      SaveProfile = False
      Exit Function
    End If

    WritePrivateProfileString "Users", "Users", CStr(UserIDs.Count), Filename
    For Ctr = 1 To UserIDs.Count
      WritePrivateProfileString "Users", "Name" & Ctr, CStr(UserIDs.No(Ctr).Name), Filename
      WritePrivateProfileString "Users", "Pass" & Ctr, UserIDs.No(Ctr).Pass, Filename
      WritePrivateProfileString "Users", "GrpName" & Ctr, UserIDs.No(Ctr).Group, Filename
      WritePrivateProfileString "Users", "DirCnt" & Ctr, CStr(UserIDs.No(Ctr).Pcnt), Filename
      
      If CliChange = True Then
      If UserIDs.No(Ctr).Disabled = "" Or UserIDs.No(Ctr).Disabled = "No" Then
      WritePrivateProfileString "Users", "Group" & Ctr & "Dis", "No", Filename
      Else
      WritePrivateProfileString "Users", "Group" & Ctr & "Dis", UserIDs.No(Ctr).Disabled, Filename
      End If
      End If
      
      If lstchange = True Then
      For X = 1 To UserIDs.No(Ctr).Pcnt
        WritePrivateProfileString "Users", "Access" & Ctr & "_" & X, _
          UserIDs.No(Ctr).Priv(X).Path & "," & UserIDs.No(Ctr).Priv(X).Accs, Filename
        WritePrivateProfileString "Users", "Home" & Ctr, CStr(UserIDs.No(Ctr).Home), Filename
      Next
      End If
    Next
    
  Dim Z As Integer, mgn As String
  Z = Group.Number
  
  If GrpChange = True Then
  For v = 1 To Z
  mgn = group_info.GrpType(v).Name
  WritePrivateProfileString "Users", "Group" & v, mgn, Filename
  WritePrivateProfileString "Users", "GrAcc" & v, group_info.GrpType(v).Access, Filename
  WritePrivateProfileString "Users", "MainGroup" & v & "Dis", group_info.GrpType(v).Disabled, Filename
  Next
  End If
  
  If frmMain.AllowAnon.Value = 1 Then
  WritePrivateProfileString "Common", "UseHidden", "Yes", Filename
  Z = frmMain.lstHidden.ListCount - 1
  For Ctr = 1 To Z
  frmMain.lstHidden.ListIndex = Ctr - 1
  h = frmMain.lstHidden.Text
  WritePrivateProfileString "Common", "Hid" & Ctr, h, Filename
  Next
  Else
  WritePrivateProfileString "Common", "UseHidden", "No", Filename
  End If
  
  WritePrivateProfileString "Common", "Port", frmMain.LisPort.Text, Filename
  WritePrivateProfileString "Common", "Maximum", frmMain.maxUnits.Text, Filename
  
  If frmMain.chkAdmin.Value = 1 Then
  WritePrivateProfileString "Common", "SeeALL", "Yes", Filename
  Else
  WritePrivateProfileString "Common", "SeeALL", "No", Filename
  End If
  If frmMain.chkAccess.Value = 1 Then
  WritePrivateProfileString "Common", "BeepAttempt", "Yes", Filename
  Else
  WritePrivateProfileString "Common", "BeepAttempt", "No", Filename
  End If
  If frmMain.chkDelete.Value = 1 Then
  WritePrivateProfileString "Common", "BeepDelete", "Yes", Filename
  Else
  WritePrivateProfileString "Common", "BeepDelete", "No", Filename
  End If
  If frmMain.AllowAnon.Value = 1 Then
  WritePrivateProfileString "Common", "Anonomous", "Yes", Filename
  Else
  WritePrivateProfileString "Common", "Anonomous", "No", Filename
  End If
  If frmMain.DenAll.Value = 1 Then
  WritePrivateProfileString "Common", "DenyAll", "Yes", Filename
  Else
  WritePrivateProfileString "Common", "DenyAll", "No", Filename
  End If

    CurrentProfile = Filename
    SaveProfile = True
  End If
  
  GrpChange = False
  CliChange = False
  CliGrpChange = False
  lstchange = False
End Function
Public Function GetFromIni(strSectionHeader As String, strVariableName As _
    String, strFileName As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, _
      GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", _
      strReturn, Len(strReturn), strFileName))
End Function




