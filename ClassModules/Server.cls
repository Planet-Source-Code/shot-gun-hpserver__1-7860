VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event ServerErrorOccurred(ByVal errNumber As Long)
Public Event ServerStarted()
Public Event ServerStopped()
Public Event NewClient(ByVal ClientID As Long)
Public Event ClientLoggedIn(ByVal ClientID As Long)
Public Event ClientLoggedOut(ByVal ClientID As Long)
Public Event ClientSentCommand(ByVal ClientID As Long, Command As String, Args As String)
Public Event ClientStatusChanged(ByVal ClientID As Long)
Public Property Get ListeningPort() As Long
    ListeningPort = Port
End Property

Public Property Let ListeningPort(NewPort As Long)
    If Port = 0 Then
        Port = NewPort
    End If
End Property

Public Property Get ServerMaxClients() As Integer
    ServerMaxClients = MaxClients
End Property

Public Property Let ServerMaxClients(Max As Integer)
    If Max >= 0 Then
        MaxClients = Max
    End If
End Property

Public Property Get TransBufferSize() As Long
    TransBufferSize = TransferBufferSize
End Property

Public Property Let TransBufferSize(BuffSize As Long)
    If BuffSize > 0 Then
        TransferBufferSize = BuffSize
    End If
End Property

Public Property Get CurrentConnectedClients() As Long
    CurrentConnectedClients = ConnectedClients
End Property

Public Property Get CurrentClientCounter() As Long
    CurrentClientCounter = ClientCounter
End Property
Public Property Get GetClientConnectedAt(ClientID As Long) As String
    GetClientConnectedAt = client(GetClientArrayLocByID(ClientID)).ConnectedAt
End Property

Public Property Get GetClientConnectMode(ClientID As Long) As String
    GetClientConnectMode = client(GetClientArrayLocByID(ClientID)).ConnectMode
End Property

Public Property Get GetClientcTotalBytesXfer(ClientID As Long) As Long
    GetClientcTotalBytesXfer = client(GetClientArrayLocByID(ClientID)).cTotalBytesXfer
End Property

Public Property Get GetClientcFileTotalBytes(ClientID As Long) As Long
    GetClientcFileTotalBytes = client(GetClientArrayLocByID(ClientID)).cFileTotalBytes
End Property

Public Property Get GetClientCurrentDir(ClientID As Long) As String
    GetClientCurrentDir = client(GetClientArrayLocByID(ClientID)).CurrentDir
End Property

Public Property Get GetClientCurrentFile(ClientID As Long) As String
    GetClientCurrentFile = client(GetClientArrayLocByID(ClientID)).CurrentFile
End Property

Public Property Get GetClientDataPort(ClientID As Long) As Long
    GetClientDataPort = client(GetClientArrayLocByID(ClientID)).DataPort
End Property

Public Property Get GetClientfFile(ClientID As Long) As Long
    GetClientfFile = client(GetClientArrayLocByID(ClientID)).fFile
End Property

Public Property Get GetClientHomeDir(ClientID As Long) As String
    GetClientHomeDir = client(GetClientArrayLocByID(ClientID)).HomeDir
End Property

Public Property Get GetClientIdleSince(ClientID As Long) As Long
    GetClientIdleSince = client(GetClientArrayLocByID(ClientID)).IdleSince
End Property

Public Property Get GetClientIPAddress(ClientID As Long) As String
    GetClientIPAddress = client(GetClientArrayLocByID(ClientID)).IPAddress
End Property

Public Property Get GetClientStatus(ClientID As Long) As String
    GetClientStatus = ServerGetClientStatusDescription(client(GetClientArrayLocByID(ClientID)).Status)
End Property

Public Property Get GetClientTotalBytesDownloaded(ClientID As Long) As Long
    GetClientTotalBytesDownloaded = client(GetClientArrayLocByID(ClientID)).TotalBytesDownloaded
End Property

Public Property Get GetClientTotalBytesUploaded(ClientID As Long) As Long
    GetClientTotalBytesUploaded = client(GetClientArrayLocByID(ClientID)).TotalBytesUploaded
End Property

Public Property Get GetClientTotalFilesDownloaded(ClientID As Long) As Long
    GetClientTotalFilesDownloaded = client(GetClientArrayLocByID(ClientID)).TotalFilesDownloaded
End Property

Public Property Get GetClientTotalFilesUploaded(ClientID As Long) As Long
    GetClientTotalFilesUploaded = client(GetClientArrayLocByID(ClientID)).TotalFilesUploaded
End Property

Public Property Get GetClientUserName(ClientID As Long) As String
    GetClientUserName = client(GetClientArrayLocByID(ClientID)).UserName
End Property
Public Function StartServer() As Long
On Error GoTo errhand
    'Make sure you dont start the server again.
    If ServerActive = True Then
        StartServer = 1001
        Exit Function
    End If

    'Make sure the server has been givin a port before doing anything.
    If Port < 1 Then
        StartServer = 1002
        Exit Function
    End If

    'If no TransferBufferSize was assigned then default to 4096 (4K)
    If TransferBufferSize < 1 Then TransferBufferSize = 4096

    'Start listening using the intial instance of the Winsock control.
    With frmWinsock.CommandSock(0)
        .LocalPort = Port
        .Listen
    End With

    ServerActive = True
    'Server started with out error and is ready to accept clients.
    RaiseEvent ServerStarted
    Exit Function
errhand:
    ServerActive = False
    RaiseEvent ServerStopped
    frmWinsock.CommandSock(0).Close
    frmMain.lblMessage.Caption = "Error Starting Server..."
    frmMain.txtCount.Text = ConnectedClients
End Function

Public Sub NewClient(requestID As Long)

    Dim tmpID As Long
    Dim i As Integer
    Dim str As String

    'Increment counters
    ConnectedClients = ConnectedClients + 1
    ClientCounter = ClientCounter + 1

    tmpID = ClientCounter
    Do
        i = i + 1
    Loop Until client(i).inUse = False
    Load frmWinsock.CommandSock(i)
    Load frmWinsock.DataSock(i)
    frmWinsock.CommandSock(i).Accept requestID

    With client(i)
        .inUse = True
        .ConnectedAt = Now
        .id = tmpID
        .Status = stat_LOGGING_IN
        .IdleSince = Now
        .IPAddress = frmWinsock.CommandSock(i).RemoteHostIP
    End With

    RaiseEvent NewClient(client(i).id)

    If ((ConnectedClients > MaxClients) And (MaxClients <> 0)) Or (ConnectedClients > MAX_CONNECTIONS) Then
        SendResponse i, "421 Too many users - try again later."
        LogoutClient i
        Exit Sub
    End If

   If frmMain.txtWel(0).Text = "" Then
   str = "HomePlay FTP version " & App.Major & ".0" & App.Minor & " build " & App.Revision
   Else
   str = frmMain.txtWel(0).Text & vbCrLf
   End If
   If frmMain.txtWel(1).Text <> "" Then
   str = str & frmMain.txtWel(1).Text & vbCrLf
   End If
   If frmMain.txtWel(2).Text <> "" Then
   str = str & frmMain.txtWel(2).Text
   End If
   SendResponse i, "220 " & str
End Sub

Public Sub SendResponse(sckArrayLoc As Integer, data As String)
On Error GoTo dunn
    frmWinsock.CommandSock(sckArrayLoc).SendData data & vbCrLf
    DoEvents
dunn:
End Sub

Private Sub SendData(sckArrayLoc As Integer, data As String)
On Error GoTo ber
    frmWinsock.DataSock(sckArrayLoc).SendData data
    
ber:
If Err.Number = 40006 Then
frmWinsock.DataSock(sckArrayLoc).Close
End If

End Sub

Public Sub SendComplete(sckArrayLoc As Integer)
    With client(sckArrayLoc)
        Select Case .Status
            Case stat_GETTING_DIR_LIST
            Call Wait(0.25)
                frmWinsock.DataSock(sckArrayLoc).Close
                SendResponse sckArrayLoc, "226 Transfer complete."
                .Status = stat_IDLE
                RaiseEvent ClientStatusChanged(.id)
    
            Case stat_DOWNLOADING
                If .cFileTotalBytes = .cTotalBytesXfer Then
                    Close #.fFile
                    Call Wait(0.5)
                    .TotalFilesDownloaded = .TotalFilesDownloaded + 1
                    frmWinsock.DataSock(sckArrayLoc).Close
                    .DataPort = 0
                    SendResponse sckArrayLoc, "226 Transfer complete."
                    .cFileTotalBytes = 0
                    .cTotalBytesXfer = 0
                    .Status = stat_IDLE
                    RaiseEvent ClientStatusChanged(.id)
                Else
                    SendFile sckArrayLoc
                End If
        End Select
    End With

End Sub

Private Sub LoginClient(cArrayLoc As Integer, Password As String, Nick As String)
Dim uji As Integer
Dim l As Integer

If Nick = "anonymous" Or Nick = "Anonymous" Then
For uji = 1 To Number
If Nick = UserIDs.No(uji).Name Then
   turn = Privledges(UserIDs.No(uji).Group)
   With client(cArrayLoc)
      l = 1
      Do Until l > UserIDs.No(uji).Pcnt
      .Priv(l).Path = UserIDs.No(uji).Priv(l).Path
      .Priv(l).Accs = UserIDs.No(uji).Priv(l).Accs
      l = l + 1
      Loop
   .Pcnt = UserIDs.No(uji).Pcnt
   .id = cArrayLoc
   .group_ID = uji
   .HomeDir = UserIDs.No(uji).Home
   .CurrentDir = UserIDs.No(uji).Home
   .Current_Access = UserIDs.No(uji).Priv(1).Accs
   .group_Access = turn
   .Group_Name = UserIDs.No(uji).Group
   SendResponse cArrayLoc, "230 User logged in, proceed."
    .Status = stat_IDLE
    frmMain.lstConned.AddItem " " & .id & " " & .UserName & " " & .IPAddress & " " & .Group_Name
   End With
 End If
Next
GoTo anon
End If

 For uji = 1 To Number
      If Password = UserIDs.No(uji).Pass And UserIDs.No(uji).Name = Nick Then
      turn = Privledges(UserIDs.No(uji).Group)
        If InStr(turn, "Q") Then
        SendResponse cArrayLoc, "530 Not Logged in, " & (UserIDs.No(uji).Group) & " account dis-abled."
        frmMain.txtSvrLog.SelStart = Len(frmMain.txtSvrLog)
        frmMain.txtSvrLog.Text = frmMain.txtSvrLog.Text & vbCrLf & "530  - Not Logged In - Account Dis-abled"
        Call Wait(0.1)
        LogoutClient cArrayLoc
        GoTo dunn
        End If

      If UserIDs.No(uji).Disabled = "" Or UserIDs.No(uji).Disabled = "No" Then
      UserIDs.No(uji).IP = frmWinsock.CommandSock(cArrayLoc).RemoteHostIP
      
      With client(cArrayLoc)
      l = 1
      Do Until l > UserIDs.No(uji).Pcnt
      .Priv(l).Path = UserIDs.No(uji).Priv(l).Path
      .Priv(l).Accs = UserIDs.No(uji).Priv(l).Accs
      l = l + 1
      Loop
      End With
      
      With client(cArrayLoc)
        .Pcnt = UserIDs.No(uji).Pcnt
        .id = cArrayLoc
        .HomeDir = UserIDs.No(uji).Home
        .group_ID = uji
        .Group_Name = UserIDs.No(uji).Group
        .group_Access = turn
        .Current_Access = UserIDs.No(uji).Priv(1).Accs
        .Password = Password
        .CurrentDir = .HomeDir
        SendResponse cArrayLoc, "230 User logged in, proceed."
        .Status = stat_IDLE
        frmMain.lstConned.AddItem " " & .id & " " & .UserName & " " & .IPAddress & " " & .Group_Name
      End With
      
anon:
      
         RaiseEvent ClientLoggedIn(ByVal client(cArrayLoc).id)
         RaiseEvent ClientStatusChanged(ByVal client(cArrayLoc).id)
         'MsgBox client(cArrayLoc).CurrentDir & "  " & client(cArrayLoc).Current_Access

        If InStr(turn, "B") Then
        drivesvisi = True
        Else
        drivesvisi = False
        End If
        GoTo dunn
       Else
      SendResponse cArrayLoc, "530 Not Logged in, password failed."
      frmMain.txtSvrLog.SelStart = Len(frmMain.txtSvrLog)
      frmMain.txtSvrLog.Text = frmMain.txtSvrLog.Text & vbCrLf & "530  - Password Incorrect for username"
      LogoutClient cArrayLoc
      GoTo dunn
       End If
      End If
      Next
      
      SendResponse cArrayLoc, "530 Not Logged in, password failed."
      frmMain.txtSvrLog.SelStart = Len(frmMain.txtSvrLog)
      frmMain.txtSvrLog.Text = frmMain.txtSvrLog.Text & vbCrLf & "530  - Password Incorrect for username"
      LogoutClient cArrayLoc
dunn:
frmMain.txtCount = ConnectedClients
End Sub

Public Sub LogoutClient(Optional ByVal id As Long, Optional cArrayLoc As Integer)
    On Error Resume Next

    If id = 0 And cArrayLoc = 0 Then Exit Sub
    
    Dim ArrayPos As Integer
    Dim tmp As Long
    
    If id = 0 Then
        ArrayPos = cArrayLoc
    Else
        ArrayPos = GetClientArrayLocByID(id)
    End If

    If ArrayPos < 1 Then Exit Sub
    Dim jim As String, Y As Integer, idd As Integer
    For Y = 0 To frmMain.lstConned.ListCount
    On Error GoTo gku
    
    If frmMain.lstConned.ListCount = 0 Then
    GoTo gku
    End If
    
    If frmMain.lstConned.ListCount = 0 Then
    Exit For
    End If
    
    frmMain.lstConned.ListIndex = Y
    jim = Parse(frmMain.lstConned.Text, 2)
    idd = Parse(frmMain.lstConned.Text, 1)
    If jim = client(ArrayPos).UserName And idd = client(ArrayPos).id Then
    frmMain.lstConned.RemoveItem Y
    End If
    Next
    
Call Wait(0.125)
gku:
    With client(ArrayPos)
    On Local Error GoTo hrr
        frmWinsock.CommandSock(ArrayPos).Close
        frmWinsock.DataSock(ArrayPos).Close
hrr:
        Unload frmWinsock.CommandSock(ArrayPos)
        Unload frmWinsock.DataSock(ArrayPos)

        If .fFile <> 0 Then Close #.fFile
        .Group_Name = ""
        .group_ID = "0"
        .Password = ""
        .ConnectedAt = ""
        .ConnectMode = 0
        .cTotalBytesXfer = 0
        .cFileTotalBytes = 0
        .CurrentDir = ""
        .CurrentFile = ""
        .DataPort = 0
        .fFile = 0
        .HomeDir = ""
        tmp = .id
        .id = 0
        .IdleSince = ""
        .IPAddress = ""
        .Status = stat_IDLE
        .TotalBytesDownloaded = 0
        .TotalBytesUploaded = 0
        .TotalFilesDownloaded = 0
        .TotalFilesUploaded = 0
        .UserName = ""
        .inUse = False
    End With

    ConnectedClients = ConnectedClients - 1
    RaiseEvent ClientLoggedOut(ByVal tmp)
    frmMain.txtCount = ConnectedClients
End Sub

Public Function GetClientArrayLocByID(ByVal id As Long) As Integer
    Dim i As Integer

    For i = 0 To UBound(client)
        If client(i).id = id Then
            GetClientArrayLocByID = i
            Exit Function
        End If
    Next
    GetClientArrayLocByID = -1

End Function

Public Sub ProcFTPCommand(ByVal sckArrayLoc As Integer, ByRef raw_data As String)
    Dim data
    Dim ftpCommand As String
    Dim ftpArgs As String
    Dim uj As Integer, findr As Boolean
    Dim tri
    data = Replace$(raw_data, vbCrLf, "")

    If InStr(data, " ") = 0 Then
        ftpCommand = data
    Else
        ftpCommand = Left$(data, (InStr(data, " ") - 1))
        ftpArgs = Right$(data, (Len(data) - InStr(data, " ")))
    End If

    RaiseEvent ClientSentCommand(client(sckArrayLoc).id, ftpCommand, ftpArgs)
    client(sckArrayLoc).IdleSince = Now
If spy_client = True Then
If UserIDs.No(sckArrayLoc).Spy = True Then
Select Case ftpCommand
Case "DELE"
frmMain.lstSpy.AddItem "* DELE - " & ftpArgs & " *"
Beep
frmMain.lblMessage.Caption = "*** Alert ***  " & UserIDs.No(sckArrayLoc).Name & " is Deleting " & client(sckArrayLoc).CurrentDir & ftpArgs
Case "CWD"
frmMain.lstSpy.AddItem "CWD - " & ftpArgs
Case "CDUP"
frmMain.lstSpy.AddItem "CDUP - " & ftpArgs
Case "STOR"
frmMain.lstSpy.AddItem "STOR - " & ftpArgs
Case "RETR"
frmMain.lstSpy.AddItem "RETR - " & ftpArgs
Case "QUIT"
frmMain.lstSpy.AddItem " *** " & UserIDs.No(sckArrayLoc).Name & " *** " & "    Quit - Logged Off    " & Time
End Select
End If
End If
    Select Case UCase$(ftpCommand)

        Case "USER"
        frmMain.txtCount.Text = ConnectedClients
          findr = False
      If ftpArgs = "anonymous " Or ftpArgs = "Anonymous " And frmMain.AllowAnon.Value = 0 Then
      SendResponse sckArrayLoc, "550 Anonymous Account Dis-Abled"
      LogoutClient sckArrayLoc
      Exit Sub
      End If
          
      For uj = 1 To Number
      If ftpArgs = UserIDs.No(uj).Name Then
      If UserIDs.No(uj).Disabled = "" Or UserIDs.No(uj).Disabled = "No" Then
        findr = True
      client(sckArrayLoc).UserName = ftpArgs
      UserIDs.No(uj).id = sckArrayLoc
      SendResponse sckArrayLoc, "331 User name ok, need password."
        GoTo dun
       Else
      SendResponse sckArrayLoc, "550 Not Logged in: Account Disabled."
      LogoutClient sckArrayLoc
      GoTo dun
       End If
      End If
      findr = False
      Next
       SendResponse sckArrayLoc, ("550 Not Logged in: No User Account for " & ftpArgs)
       LogoutClient sckArrayLoc
dun:
        Case "PASS"
        Dim doubleup As String
        doubleup = client(sckArrayLoc).UserName
        LoginClient sckArrayLoc, ftpArgs, doubleup
      
        Case "TYPE"
        If ftpArgs = "A" Then
            SendResponse sckArrayLoc, "200 Type set to ASCII"
        End If
        If ftpArgs = "I" Then
            SendResponse sckArrayLoc, "200 Type set to Binary"
        End If
        
        Case "REST"
            'SendResponse sckArrayLoc, "350 Restarting at " & ftpArgs & " - send STORE or RETRIEVE to initiate transfer."
            SendResponse sckArrayLoc, "550 Command not Implemented Yet"
        
        Case "PWD"
        turn = client(sckArrayLoc).group_Access
        If InStr(turn, "P") Then
        GoTo def
        End If
        
       If InStr(turn, "R") Then
        SendResponse sckArrayLoc, "257 " & Chr(34) _
            & ConvPathToRelative(client(sckArrayLoc).HomeDir, client(sckArrayLoc).CurrentDir) _
            & Chr(34) & " is current directory."
        Exit Sub
        End If
def:
        If InStr(turn, "B") Then
        drivesvisi = True
        Else
        drivesvisi = False
        End If

        tri = client(sckArrayLoc).CurrentDir
        SendResponse sckArrayLoc, "257 " & Chr(34) & tri & Chr(34) & " is working directory." '& vbCrLf
                
        Case "PORT"
            Dim ret As Boolean
            Dim tmpArray() As String
            ret = Parse2Array(ftpArgs, tmpArray, ",")
            client(sckArrayLoc).DataPort = tmpArray(4) * 256 Or tmpArray(5)
            SendResponse sckArrayLoc, "200 Port command successful."
            
        Case "HELP"
        Dim str As String
        str = "214 DELE USER PASS TYPE STOR RETR PWD LIST SYST MKD CWD RMD CDUP"
            SendResponse sckArrayLoc, str
         
        Case "LIST"
        turn = client(sckArrayLoc).group_Access
        If InStr(turn, "P") Then
        GoTo wbirf
        End If
        
        If client(sckArrayLoc).UserName = "anonymous" Or client(sckArrayLoc).UserName = "Anonymous " Then
        GoTo wbirf
        End If
    
        If InStr(client(sckArrayLoc).Current_Access, "L") Then
        GoTo wbirf
        Else
        SendResponse sckArrayLoc, ("550 LIST: Not Allowed")
        WriteToLogWindow "Client " & client(sckArrayLoc).id & " List Not Allowed", True
        Exit Sub
        End If
wbirf:
            
            client(sckArrayLoc).Status = stat_GETTING_DIR_LIST
            SendResponse sckArrayLoc, "150 Opening ASCII mode data connection for /bin/ls."
            
            client(sckArrayLoc).Status = stat_GETTING_DIR_LIST
            RaiseEvent ClientStatusChanged(client(sckArrayLoc).id)
            Dim min As String
            min = client(sckArrayLoc).CurrentDir
            
            If InStr(client(sckArrayLoc).Current_Access, "S") Then
            ret = GetDirectoryList(sckArrayLoc, min, "LIST")
            Else
            ret = GetDirectoryList(sckArrayLoc, min, "NLST")
            End If

        Case "RETR"
        turn = client(sckArrayLoc).group_Access
        If InStr(turn, "P") Then
        GoTo airf
        End If
        If InStr(client(sckArrayLoc).Current_Access, "T") Then
        GoTo airf
        Else
        SendResponse sckArrayLoc, ("550 RETR: Not Allowed")
        WriteToLogWindow "Client " & client(sckArrayLoc).id & " RETR Not Allowed", True
        Exit Sub
        End If
airf:
         client(sckArrayLoc).Status = stat_DOWNLOADING
         GetFileToSend sckArrayLoc, ftpArgs
            
        Case "STOR"
        turn = client(sckArrayLoc).group_Access
        If InStr(turn, "P") Then
        GoTo jihrf
        End If
        
        If InStr(client(sckArrayLoc).Current_Access, "T") Then
        GoTo jihrf
        Else
        SendResponse sckArrayLoc, ("550 STOR: Not Allowed")
        WriteToLogWindow "Client " & client(sckArrayLoc).id & " STOR Not Allowed", True
        Exit Sub
        End If
jihrf:
        If Right(client(sckArrayLoc).CurrentDir, 1) <> "\" Then
        client(sckArrayLoc).CurrentDir = client(sckArrayLoc).CurrentDir & "\"
        End If
        MakeDataConnection sckArrayLoc
        ret = open_file(client(sckArrayLoc).CurrentDir & ftpArgs, sckArrayLoc)
        client(sckArrayLoc).Status = stat_UPLOADING
        client(sckArrayLoc).full_file_name = (client(sckArrayLoc).CurrentDir & ftpArgs)
        SendResponse sckArrayLoc, "150 Openning BINARY mode connection. " & (client(sckArrayLoc).CurrentDir & ftpArgs)
       
        Case "CWD"
        turn = client(sckArrayLoc).group_Access
        If InStr(turn, "P") Then
        GoTo wif
        End If
        
        ret = CheckPath(ftpArgs, sckArrayLoc)
        
        If ret = False Then
        SendResponse sckArrayLoc, ("550 No Access to Directory - " & ftpArgs)
        Exit Sub
        End If
wif:
        If InStr(turn, "P") Then
        GoTo strt
        End If
        
        If isHiddenDir(sckArrayLoc, ftpArgs) = True Then
        GoTo nott
        End If
strt:
            ChangeDirectory sckArrayLoc, ftpArgs
            Exit Sub
nott:
        SendResponse sckArrayLoc, ("550 CWD to " & ftpArgs & " Not Allowed")
        If frmMain.chkAccess.Value = 1 Then
        Beep
        frmMain.lblMessage = " *** ALERT ***  " & "'" & client(sckArrayLoc).UserName & "'" & "  is trying to access " & ftpArgs
        End If
        WriteToLogWindow "Client " & client(sckArrayLoc).id & " CWD to " & ftpArgs & " Not Allowed", True
        SendResponse sckArrayLoc, "257 " & Chr(34) & tri & Chr(34) & " is working directory."
        
        Case "DELE"
        turn = client(sckArrayLoc).group_Access
        If InStr(turn, "P") Then
        GoTo wihrf
        End If
        
        If InStr(client(sckArrayLoc).Current_Access, "D") Then
        GoTo wihrf
        Else
        SendResponse sckArrayLoc, ("550 DELE: Not Allowed")
        WriteToLogWindow "Client " & client(sckArrayLoc).id & " DELE Not Allowed", True
        Exit Sub
        End If
wihrf:
            Beep
            frmMain.lblMessage.Caption = "*** Alert ***  " & client(sckArrayLoc).UserName & " is Deleting " & client(sckArrayLoc).CurrentDir & ftpArgs
            
        Dim ActionFlag As Long
            
        
       
doinit:
        ShellDelete (client(sckArrayLoc).CurrentDir & ftpArgs & Chr$(0)), del_path
    
            SendResponse sckArrayLoc, "250 " & "file " & ftpArgs & " was deleted successfully."
            
        Case "RMD"
        turn = client(sckArrayLoc).group_Access
        If InStr(turn, "P") Then
        GoTo dihrf
        End If
        
        If InStr(client(sckArrayLoc).Current_Access, "H") Then
        GoTo dihrf
        Else
        SendResponse sckArrayLoc, ("550 RMD: Not Allowed")
        WriteToLogWindow "Client " & client(sckArrayLoc).id & " RMD Not Allowed", True
        Exit Sub
        End If
dihrf:
        On Local Error GoTo air
        
           RmDir client(sckArrayLoc).CurrentDir & ftpArgs & "\"
           'Call DelTree32(client(sckArrayLoc).CurrentDir, ftpArgs)

            
            SendResponse sckArrayLoc, "250 " & "Directory " & ftpArgs & " was deleted successfully."
            frmMain.txtSvrLog.SelStart = Len(frmMain.txtSvrLog)
            frmMain.txtSvrLog.Text = frmMain.txtSvrLog & vbCrLf & ("[" & Date & " " & Time & "]" & " 250 directory " & ftpArgs & " removed")
            Exit Sub
air:
        If Err.Number = 75 Then
        SendResponse sckArrayLoc, "550 " & "directory " & ftpArgs & " was NOT removed."
        frmMain.txtSvrLog.SelStart = Len(frmMain.txtSvrLog)
        frmMain.txtSvrLog.Text = frmMain.txtSvrLog & vbCrLf & ("[" & Date & " " & Time & "]" & " *** 550 directory " & ftpArgs & " NOT removed ***")
        End If
        
        Case "CDUP"
        turn = client(sckArrayLoc).group_Access
        If InStr(turn, "P") Then
        GoTo wrf
        End If
        
        If InStr(client(sckArrayLoc).Current_Access, "W") Then
        GoTo wrf
        Else
        SendResponse sckArrayLoc, ("550 CDUP: Not Allowed")
        Exit Sub
        End If
wrf:
            Dim tmp As String

            tmp = client(sckArrayLoc).CurrentDir
            
            If isRootDir(sckArrayLoc, tmp) = False Then
                If Right$(tmp, 1) = "\" Then tmp = Left$(tmp, Len(tmp) - 1)
                tmp = Left$(tmp, InStrRev(tmp, "\"))
            End If
            
            ChangeDirectory sckArrayLoc, ConvPathToRelative(client(sckArrayLoc).HomeDir, tmp)
            
        Case "MKD"
        turn = client(sckArrayLoc).group_Access
        If InStr(turn, "P") Then
        GoTo bwrf
        End If
        If InStr(client(sckArrayLoc).Current_Access, "M") Then
        GoTo bwrf
        Else
        SendResponse sckArrayLoc, ("550 MKD: Not Allowed")
        Exit Sub
        End If
bwrf:
        create_dir = client(sckArrayLoc).CurrentDir & ftpArgs
        MkDir create_dir
        SendResponse sckArrayLoc, "257 " & create_dir & " created successfully."
        
        Case "PASV"
            client(sckArrayLoc).ConnectMode = cMode_PASV
            SendResponse sckArrayLoc, "227 Entering Passive Mode (" _
                & Replace(frmWinsock.CommandSock(0).LocalIP, ".", ",") & OpenLocalDataPort(sckArrayLoc) & ")"

        Case "NOOP"
            SendResponse sckArrayLoc, "200 NOOP command successful."
            
        Case "SYST"
            SendResponse sckArrayLoc, "215 HomePlay TCP/IP"
            
        Case "QUIT"
        SendResponse sckArrayLoc, "221 " & frmMain.Get_closing
        Call Wait(0.25)
        LogoutClient client(sckArrayLoc).id
        
        Case "ŸÔ"
        halt_transfer = True
        
        Case "ÒABOR"
        If client(sckArrayLoc).Status = stat_UPLOADING Then
        frmWinsock.DataSock(sckArrayLoc).Close
        Close #5
        SendResponse sckArrayLoc, "226 ABOR command executed."
        client(sckArrayLoc).TotalFilesUploaded = client(sckArrayLoc).TotalFilesUploaded - 1
        Kill client(sckArrayLoc).full_file_name
        client(sckArrayLoc).Status = stat_IDLE
        End If
        
        If client(sckArrayLoc).Status = stat_DOWNLOADING Then
        frmWinsock.DataSock(sckArrayLoc).Close
        SendResponse sckArrayLoc, "426 Data connection closed"
        Call Wait(0.1)
        SendResponse sckArrayLoc, "226 Aborted OK"
        client(sckArrayLoc).TotalFilesDownloaded = client(sckArrayLoc).TotalFilesDownloaded - 1
        client(sckArrayLoc).Status = stat_IDLE
        End If

        
        Case "ABOR"
        If client(sckArrayLoc).Status = stat_UPLOADING Then
        frmWinsock.DataSock(sckArrayLoc).Close
        Close #5
        SendResponse sckArrayLoc, "226 ABOR command executed."
        client(sckArrayLoc).TotalFilesUploaded = client(sckArrayLoc).TotalFilesUploaded - 1
        Kill client(sckArrayLoc).full_file_name
        client(sckArrayLoc).Status = stat_IDLE
        End If
        
        If client(sckArrayLoc).Status = stat_DOWNLOADING Then
        frmWinsock.DataSock(sckArrayLoc).Close
        SendResponse sckArrayLoc, "226 Aborted OK"
        client(sckArrayLoc).TotalFilesDownloaded = client(sckArrayLoc).TotalFilesDownloaded - 1
        client(sckArrayLoc).Status = stat_IDLE
        End If

        Case Else
        frmMain.Text1.Text = UCase$(ftpCommand)

            SendResponse sckArrayLoc, "502 Command not implemented."

    End Select

End Sub

Private Function GetDirectoryList(cArrayLoc As Integer, PathName As String, ByVal typ As String)
  Dim File_Name As String, name_ As String, exte_ As String
  Dim DummyS As String
  Dim SepN As Integer
  Dim Full_Name As String
  Dim Buffer As String
  Dim DIRList As String
If InStr(PathName, "*") Or InStr(PathName, "?") Then
      Full_Name = PathName
      File_Name = Dir$(Full_Name)
    ElseIf (GetAttr(PathName) And 16) <> 0 Then
      If Right$(PathName, 1) = "\" Then
        PathName = Left$(PathName, Len(PathName) - 1)
      End If
      File_Name = Dir$(PathName & "\*.*", 16)
      Full_Name = PathName & "\" & File_Name
    Else
      Full_Name = PathName
      File_Name = Dir$(Full_Name)
    End If
    If Err.Number = 0 Then
      Do
        If Not File_Name = "pagefile.sys" Then
        
        If File_Name = "." Or File_Name = ".." Then
          DummyS = Format$(File_Name, "@@@@@@@@@@@@!") & " <DIR>"
        ElseIf InStr(Full_Name, "*") Or InStr(Full_Name, "?") Then
        'If InStr(Full_Name, "*") Or InStr(Full_Name, "?") Then
          'file
          SepN = InStr(File_Name, ".")
          If SepN <> 0 Then
            name_ = Left$(File_Name, SepN - 1)
            exte_ = Mid$(File_Name, SepN + 1)
          Else
            name_ = File_Name
            exte_ = "   "
          End If
          DummyS = "-rwxr--r--   1 user    group  "
          If typ = "LIST" Then
            DummyS = DummyS & Format$(FileLen(Full_Name), " @@@@@@@@@") _
             & " " & Format$(FileDateTime(Full_Name), " mmm dd hh:nn ") & File_Name
          ElseIf typ = "NLST" Then
          End If
        ElseIf GetAttr(Full_Name) = 16 Then
          SepN = InStr(File_Name, ".")
          If SepN <> 0 Then
            name_ = Left$(File_Name, SepN - 1)
            exte_ = Mid$(File_Name, SepN + 1)
          Else
            name_ = File_Name
            exte_ = "   "
          End If
          DummyS = "drwxr-xr-x   1 user    group  "
          If typ = "LIST" Then
            DummyS = DummyS & Format$(FileLen(Full_Name), " @@@@@@@@@") _
             & " " & Format$(FileDateTime(Full_Name), " mmm dd hh:nn ") & File_Name
          ElseIf typ = "NLST" Then
            DummyS = ""
          End If
        Else
          SepN = InStr(File_Name, ".")
          If SepN <> 0 Then
            name_ = Left$(File_Name, SepN - 1)
            exte_ = Mid$(File_Name, SepN + 1)
          Else
            name_ = File_Name
            exte_ = "   "
          End If
          DummyS = "-rwxr--r--   1 user    group  "
          If typ = "LIST" Then
            DummyS = DummyS & Format$(FileLen(Full_Name), " @@@@@@@@@") _
             & " " & Format$(FileDateTime(Full_Name), " mmm dd hh:nn ") & File_Name
          ElseIf typ = "NLST" Then
          DummyS = "-rwxr--r--   1 user    group  "
          DummyS = DummyS & Format$(FileLen(Full_Name), " @@@@@@@@@") _
             & " " & Format$(FileDateTime(Full_Name), " mmm dd hh:nn ") & File_Name
          End If
        End If
        Buffer = Buffer & DummyS & vbCrLf
        File_Name = Dir$
        If Left(File_Name, 1) = "p" Then
          File_Name = Dir$
        End If
      Debug.Print "File Name = " & File_Name
              If File_Name = "" Then Exit Do
              Full_Name = PathName & "\" & File_Name
      Else
        File_Name = Dir$
      End If
      Loop
      DIRList = Buffer
      GoTo doit
    ElseIf (Err.Number > 51 And Err.Number < 77) Or (Err.Number > 707 And Err.Number < 732) Then
      SendResponse cArrayLoc, "450 Command Not Completed"
      frmMain.lblMessage.Caption = "Closeing Data connection for " & client(cArrayLoc).UserName & " Logging off"
      LogoutClient client(cArrayLoc).id
    Else
    End If
  Exit Function
doit:
    
    MakeDataConnection cArrayLoc
    
    If drivesvisi = True Then
    If dd = True Then
    DIRList = DIRList & "drwxr-xr-x   1 user    group          " & " 0 " & " Apr 12 " & "22:05 " & "D-DRIVE" & vbCrLf
                End If
    If ed = True Then
    DIRList = DIRList & "drwxr-xr-x   1 user    group          " & " 0 " & " Apr 12 " & "22:05 " & "E-DRIVE" & vbCrLf
                End If
    If fd = True Then
    DIRList = DIRList & "drwxr-xr-x   1 user    group          " & " 0 " & " Apr 12 " & "22:05 " & "F-DRIVE" & vbCrLf
                End If
                End If
                
      DIRList = "total " & CountStr(DIRList, vbCrLf) & vbCrLf & DIRList
    SendData cArrayLoc, DIRList

End Function

Private Function MakeDataConnection(sckArrayLoc As Integer) As Long
'On Error GoTo air
    If client(sckArrayLoc).ConnectMode = cMode_NORMAL Then
        frmWinsock.DataSock(sckArrayLoc).RemoteHost = client(sckArrayLoc).IPAddress
        frmWinsock.DataSock(sckArrayLoc).RemotePort = client(sckArrayLoc).DataPort
        frmWinsock.DataSock(sckArrayLoc).Connect
    End If
    Do
        DoEvents
    Loop Until frmWinsock.DataSock(sckArrayLoc).State = sckConnected
    Exit Function
air:
frmWinsock.DataSock(sckArrayLoc).Close
End Function

Private Function OpenLocalDataPort(sckArrayLoc As Integer) As String

    'In PASV mode the server creates a listing socket for the
    'client to connect to.  In Normal mode the client creates
    'the listening socket and the server will connect but if
    'the client is behind a firewall this can cause problems.

    Dim Nr1 As Integer
    Dim Nr2 As Integer

    Randomize Timer
    Nr1 = Int(Rnd * 12) + 5
    Nr2 = Int(Rnd * 254) + 1

    frmWinsock.DataSock(sckArrayLoc).LocalPort = (Nr1 * 256) Or Nr2
    frmWinsock.DataSock(sckArrayLoc).Listen

    'In PASV mode you will not make the connection here, you first need to
    'tell the client what the new open port is.

    OpenLocalDataPort = "," & Nr1 & "," & Nr2

End Function

Private Function isRootDir(cArrayLoc As Integer, strDir As String) As Boolean

    If client(cArrayLoc).HomeDir = strDir Then isRootDir = True

End Function
Private Function isHiddenDir(cArrayLoc As Integer, strDir As String) As Boolean
If InStr(hidDir_list, UCase(strDir)) Then
        If client(cArrayLoc).Group_Name = "Administrator" And frmMain.chkAdmin.Value = 1 Then
        isHiddenDir = False
        Else
        isHiddenDir = True
        End If
        End If
End Function
Private Sub ChangeDirectory(cArrayLoc As Integer, ChangeTo As String)

If Len(ChangeTo) = 1 And ChangeTo = "/" Then
ChDir client(cArrayLoc).HomeDir
client(cArrayLoc).CurrentDir = client(cArrayLoc).HomeDir
If frmMain.txtChgdir.Text <> "" Then
SendResponse cArrayLoc, "250 " & frmMain.txtChgdir.Text & " Directory changed to /"
Else
SendResponse cArrayLoc, "250 Directory changed to /"
End If
Exit Sub
End If
If Len(ChangeTo) = 1 And ChangeTo = "d" Then
ChDrive "d:"
ChDir "d:\"
client(cArrayLoc).CurrentDir = "d:\"
If frmMain.txtChgdir.Text <> "" Then
SendResponse cArrayLoc, "250 " & frmMain.txtChgdir.Text & " Directory changed to /"
Else
SendResponse cArrayLoc, "250 Directory changed to /"
End If
Exit Sub
End If
If Len(ChangeTo) = 1 And ChangeTo = "e" Then
ChDrive "e:"
ChDir "e:\"
client(cArrayLoc).CurrentDir = "e:\"
If frmMain.txtChgdir.Text <> "" Then
SendResponse cArrayLoc, "250 " & frmMain.txtChgdir.Text & " Directory changed to /"
Else
SendResponse cArrayLoc, "250 Directory changed to /"
End If
Exit Sub
End If

If Len(ChangeTo) = 1 And ChangeTo = "f" Then
ChDrive "f:"
ChDir "e:\"
client(cArrayLoc).CurrentDir = "f:\"
If frmMain.txtChgdir.Text <> "" Then
SendResponse cArrayLoc, "250 " & frmMain.txtChgdir.Text & " Directory changed to /"
Else
SendResponse cArrayLoc, "250 Directory changed to /"
End If
Exit Sub
End If

If ChangeTo = ".." Then
ChDrive "c:"
ChDir "c:\"
client(cArrayLoc).CurrentDir = "c:\"
If frmMain.txtChgdir.Text <> "" Then
SendResponse cArrayLoc, "250 " & frmMain.txtChgdir.Text & " Directory changed to /"
Else
SendResponse cArrayLoc, "250 Directory changed to /"
End If
Exit Sub
End If

If ChangeTo = "D-DRIVE" Then
ChDrive "d:"
ChDir "d:\"
client(cArrayLoc).CurrentDir = "d:\"
If frmMain.txtChgdir.Text <> "" Then
SendResponse cArrayLoc, "250 " & frmMain.txtChgdir.Text & " Directory changed to /"
Else
SendResponse cArrayLoc, "250 Directory changed to /"
End If
Exit Sub
End If

If ChangeTo = "E-DRIVE" Then
ChDrive "e:"
ChDir "e:\"
client(cArrayLoc).CurrentDir = "e:\"
If frmMain.txtChgdir.Text <> "" Then
SendResponse cArrayLoc, "250 " & frmMain.txtChgdir.Text & " Directory changed to /"
Else
SendResponse cArrayLoc, "250 Directory changed to /"
End If
Exit Sub
End If

If ChangeTo = "F-DRIVE" Then
ChDrive "f:"
ChDir "e:\"
client(cArrayLoc).CurrentDir = "f:\"
If frmMain.txtChgdir.Text <> "" Then
SendResponse cArrayLoc, "250 " & frmMain.txtChgdir.Text & " Directory changed to /"
Else
SendResponse cArrayLoc, "250 Directory changed to /"
End If
Exit Sub
End If

    If Left$(ChangeTo, 1) = "/" Then
        'Client specified an absolute path.
        If FileExists(ConvPathToLocal(client(cArrayLoc).HomeDir, ChangeTo)) = True Then
            client(cArrayLoc).CurrentDir = ConvPathToLocal(client(cArrayLoc).HomeDir, ChangeTo)
        Else
            SendResponse cArrayLoc, "550 " & ChangeTo & " : No such file or directory."
            Exit Sub
        End If
    Else
        'Client specified a path relative to their current path.
        If FileExists(ConvPathToLocal(client(cArrayLoc).CurrentDir, ChangeTo)) = True Then
            client(cArrayLoc).CurrentDir = ConvPathToLocal(client(cArrayLoc).CurrentDir, ChangeTo)
        Else
            SendResponse cArrayLoc, "550 " & ChangeTo & ": No such file or directory."
            Exit Sub
        End If
    End If

If frmMain.txtChgdir.Text <> "" Then
SendResponse cArrayLoc, "250 " & frmMain.txtChgdir.Text & "  Directory changed to " & ConvPathToRelative(client(cArrayLoc).HomeDir, client(cArrayLoc).CurrentDir)
Else
SendResponse cArrayLoc, "250 Directory changed to " & ConvPathToRelative(client(cArrayLoc).HomeDir, client(cArrayLoc).CurrentDir)
End If
    

End Sub

Private Sub GetFileToSend(cArrayLoc As Integer, File As String)

    With client(cArrayLoc)
        If FileExists(.CurrentDir & File) = False Then
            SendResponse cArrayLoc, "550 " & File & ": No such file or directory."
            Exit Sub
        End If
        .cFileTotalBytes = FileLen(.CurrentDir & File)
        .CurrentFile = .CurrentDir & File
        SendResponse cArrayLoc, "150 Opening BINARY mode data connection for " & File & " (" & .cFileTotalBytes & " bytes)"
        .fFile = FreeFile
        Open .CurrentDir & File For Binary Access Read As #.fFile

        .Status = stat_DOWNLOADING
        RaiseEvent ClientStatusChanged(.id)
    End With
    MakeDataConnection cArrayLoc
    SendFile cArrayLoc

End Sub

Private Sub SendFile(cArrayLoc As Integer)
    
    Dim BlockSize As Integer
    Dim DataToSend As String
    
    If halt_transfer = True Then
    Exit Sub
    End If
    
    BlockSize = TransferBufferSize

    With client(cArrayLoc)
        If BlockSize > (.cFileTotalBytes - .cTotalBytesXfer) Then
            BlockSize = (.cFileTotalBytes - .cTotalBytesXfer)
        End If

        DataToSend = Space$(BlockSize)
        Get .fFile, , DataToSend

        .cTotalBytesXfer = .cTotalBytesXfer + BlockSize
        .TotalBytesDownloaded = .TotalBytesDownloaded + BlockSize
    End With
    SendData cArrayLoc, DataToSend

End Sub

Public Function ShutdownServer() As Long
    frmWinsock.CommandSock(0).Close

    ServerActive = False
    RaiseEvent ServerStopped

End Function

Public Function ConvPathToLocal(ByVal StartPath As String, ByVal CurrentPath As String) As String

    'This function takes a relative path and converts to to a local one.
    'eg. home directory is "/", locally "/" equates to "C:\Files\".
    'Their current directory is "/Uploads/Stuff".
    'If passed through this function you will get
    '"C:\Files\Uploads\Stuff\"

    'StartPath must be a local path.  "C:\Files\" identifying the home directory.
    'CurrentPath is the clients current relative path.  "/Uploads/Stuff"

    Dim Result As String

    If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"

    If Left$(CurrentPath, 1) = "/" Then CurrentPath = Right$(CurrentPath, Len(CurrentPath) - 1)
    CurrentPath = Replace$(CurrentPath, "/", "\")

    Result = StartPath & CurrentPath

    If Right$(Result, 1) <> "\" Then Result = Result & "\"

    ConvPathToLocal = Result

End Function

Public Function ConvPathToRelative(ByVal StartPath As String, ByVal CurrentPath As String) As String

    'This function takes an absolute path and converts it to a relative unix style one.
    'eg. home directory is "C:\MyFiles\Music\" and the client is currently in
    '"C:\MyFiles\Music\Alternative\", pass these two paths into this function
    'will give you the result /Alternative

    'If Right$(tmp, 1) = "\" Then tmp = Left$(tmp, Len(tmp) - 1)
    'tmp = Left$(tmp, InStrRev(tmp, "\"))

    If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"
    If Right$(CurrentPath, 1) <> "\" Then CurrentPath = CurrentPath & "\"

    Dim strRelPath As String

    If StartPath = CurrentPath Then
        strRelPath = "/"
    Else
        strRelPath = Replace$(CurrentPath, StartPath, "/")
        strRelPath = Replace$(strRelPath, "\", "/")
    
        If Right$(strRelPath, 1) = "/" Then strRelPath = Left$(strRelPath, Len(strRelPath) - 1)
    End If

    ConvPathToRelative = strRelPath

End Function

Public Function ServerGetClientStatusDescription(ByVal stat As Integer) As String

    Select Case stat
        Case stat_IDLE: ServerGetClientStatusDescription = "Idle"
        Case stat_LOGGING_IN: ServerGetClientStatusDescription = "Connecting..."
        Case stat_GETTING_DIR_LIST: ServerGetClientStatusDescription = "Downloading list of files"
        Case stat_UPLOADING: ServerGetClientStatusDescription = "Uploading"
        Case stat_DOWNLOADING: ServerGetClientStatusDescription = "Downloading"
        Case Else: ServerGetClientStatusDescription = "Unknown status"
    End Select

End Function

Public Function ServerGetErrorDescription(ByVal errCode As Long) As String

    Select Case errCode

        Case 1001: ServerGetErrorDescription = "Server is already running."
        Case 1002: ServerGetErrorDescription = "Server failed to start becuase no port or invalid port was specified."
        
        Case Else: ServerGetErrorDescription = "Unknown error " & errCode

    End Select

End Function


'================================= MISC CODE =================================


Private Function get_date(FT As FileTime) As String
    Dim ST As SYSTEMTIME
    Dim r As Long
    Dim ds As String
    
    r = FileTimeToSystemTime(FT, ST)

    ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)

    If DateDiff("d", ds, Date) > 365 Then
        get_date = Format$(ds, " mmm dd yyyy ")
    Else
        get_date = Format$(ds & " " & ST.wHour & ":" & ST.wMinute, " mmm dd hh:mm ")
    End If

End Function

Private Function FileExists(Filename As String) As Boolean

    Dim hFindFile As Long
    Dim FileData As WIN32_FIND_DATA
    If Right(Filename, 1) = "\" Then
        Filename = Left(Filename, Len(Filename) - 1)
    End If

    hFindFile = FindFirstFile(Filename, FileData)
    If hFindFile = -1 Then
        FileExists = False
    Else
        FileExists = True
    End If

    FindClose hFindFile

End Function
