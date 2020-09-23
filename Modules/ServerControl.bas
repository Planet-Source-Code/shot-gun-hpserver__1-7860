Attribute VB_Name = "ServerControl"
Option Explicit

Public Sub StartServer()
On Error GoTo skip
    'Variable to store the result of functions
    Dim r As Long

    'Before you can actually start the server you must set
    'the proper settings first.
    With frmMain
        'Tell the server object which port to listen on.
        .FTPServer.ListeningPort = frmMain.LisPort.Text

        'Total max clients
        .FTPServer.ServerMaxClients = frmMain.maxUnits.Text
    
        'Start the FTP server.
        r = .FTPServer.StartServer()

        If r <> 0 Then  'Problem starting server
            MsgBox .FTPServer.ServerGetErrorDescription(r), vbCritical: DoEvents
        End If
    End With
skip:
End Sub

Public Sub StopServer()

    frmMain.FTPServer.ShutdownServer

End Sub
