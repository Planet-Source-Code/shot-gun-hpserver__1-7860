Attribute VB_Name = "PathAccess"
Public Function CheckPath(Path As String, Socket As Integer) As Boolean
Dim m As Integer
CheckPath = False
If Right(Path, 1) <> "\" Then
        Path = Path & "\"
        End If
        
If Right(UCase(client(Socket).CurrentDir), 1) <> "\" Then
        tri = UCase(client(Socket).CurrentDir & "\" & Path)
        Else
        tri = UCase(client(Socket).CurrentDir & Path)
        End If
        
        For m = 1 To client(Socket).Pcnt
        DummyS = UCase(client(Socket).Priv(m).Path)
        If Right(DummyS, 1) <> "\" Then
        DummyS = DummyS & "\"
        End If
        If DummyS = tri Then
        client(Socket).Current_Access = client(Socket).Priv(m).Accs
        CheckPath = True
        Exit For
        End If
        Next
        
        DummyS = (client(Socket).Priv(m).Accs)
End Function

Function Validate_Directory(ByVal strPath As String)
       Validate_Directory = True

              If (Right$(strPath, 1) = "\") Then
                     strPath = strPath & "\"
              End If

                            If Dir$(strPath, 16) = "" Then
                                   Validate_Directory = False
                            End If


End Function


