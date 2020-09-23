Attribute VB_Name = "MiscFunctions"
Option Explicit

Public Sub WriteToLogWindow(strString As String, Optional TimeStamp As Boolean)

    Dim strTimeStamp As String

    If TimeStamp = True Then strTimeStamp = "[" & Now & "] "
    frmMain.txtSvrLog.SelStart = Len(frmMain.txtSvrLog.Text)  ' Moved was last line
    frmMain.txtSvrLog.Text = frmMain.txtSvrLog.Text & vbCrLf & strTimeStamp & strString

End Sub

Public Function StripNulls(strString As Variant) As String

    If InStr(strString, vbNullChar) Then
        StripNulls = Left(strString, InStr(strString, vbNullChar) - 1)
    Else
        StripNulls = strString
    End If

End Function
