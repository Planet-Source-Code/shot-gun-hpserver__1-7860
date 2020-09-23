VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "TreeCreate"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   2640
      Width           =   5295
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         Caption         =   "This program must run in same directory as HPFTP Server  Not Sub Dir"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      Begin VB.Image Butt3 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":0442
         ToolTipText     =   "Exit"
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Butt2 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":074C
         ToolTipText     =   "Delete Tree"
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image Butt1 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":0A56
         ToolTipText     =   "Create Tree"
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "Select Directory to Create In"
         Height          =   355
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label lblDescrip 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0E98
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to TreeCreator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, tmpS As String

Private Sub Butt1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And selected = True Then
Butt1.BorderStyle = 1
End If
End Sub

Private Sub Butt1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim retval As Boolean
Butt1.BorderStyle = 0
If selected = False Then Exit Sub
retval = Create_File(tmpS, True)
Create (tmpS)
FakeFile
retval = Create_File(tmpS, False)
End Sub

Private Sub Butt2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Butt2.BorderStyle = 1
End If
End Sub

Private Sub Butt2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Butt2.BorderStyle = 0
MsgBox "Click on the directory then hit 'Delete' on your keyboard"
End If
End Sub

Private Sub Butt3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Butt3.BorderStyle = 1
End If
End Sub

Private Sub Butt3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Butt3.BorderStyle = 0
Call Wait(0.25)
Unload Me
End Sub

Private Sub Command1_Click()
Dim ret As String
ret = BrowseForFolder(frmMain.hWnd, "Select a Directory")
If ret = "" Then Exit Sub
tmpS = ret
selected = True
End Sub

Private Sub Form_Load()
selected = False
tFile = "C:\Windows\Temp\hpftp.dat"
End Sub
