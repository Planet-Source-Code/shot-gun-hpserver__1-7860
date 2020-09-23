Attribute VB_Name = "Recycle"
Public Type SHFILEOPSTRUCT
 hWnd As Long
 wFunc As Long
 pFrom As String
 pTo As String
 fFlags As Integer
 fAborted As Boolean
 hNameMaps As Long
 sProgress As String
 End Type
 
 Public Const FO_MOVE = &H1
 Public Const FO_COPY = &H2
 Public Const FO_DELETE = &H3
 Public Const FOF_ALLOWUNDO = &H40
 Public Const FOF_NOCONFIRMATION = &H10
 Public Const FOF_SILENT = &H4
 
 Dim FO_FLAG As Long
 Dim FOF_FLAGS As Long
 Global fNames() As String
 
 Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
(lpFileOp As SHFILEOPSTRUCT) As Long
Public Function Build_of_Flags() As Long
        Dim flag As Long
        flag = 0&
        
        If frmMain.mnuMove.Checked Then
        FO_FLAG = FO_MOVE
        Else
        FO_FLAG = FO_DELETE
        End If
        
        If frmMain.mnuConfirm.Checked Or frmMain.mnuSilent.Checked Then flag = flag Or FOF_ALLOWUNDO
        If frmMain.mnuKill.Checked Then
        flag = flag Or FOF_SILENT & flag Or FOF_NOCONFIRMATION
        End If
        If frmMain.mnuSilent.Checked Then flag = flag Or FOF_NOCONFIRMATION
        flag = flag Or FOF_RENAMEONCOLLISION
        
        Build_of_Flags = flag
End Function
Public Function ShellDelete(sFile As String, sDestination As String)
        Dim i As Integer
        Dim r As Long
        Dim sFiles As String
        Dim SHFileOp As SHFILEOPSTRUCT
        sFiles = sFile & Chr$(0)
        FOF_FLAGS = Build_of_Flags()
    
        With SHFileOp
       .wFunc = FO_FLAG
       .pFrom = sFiles
       .pTo = sDestination
       .fFlags = FOF_FLAGS
        End With

        r = SHFileOperation(SHFileOp)
End Function
