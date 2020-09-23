VERSION 5.00
Begin VB.Form frmAddUser 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New User"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtPass 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Name and Password for New Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim i As Integer
If txtName = "anonymous" Or txtName = "Anonymous" Then
If txtPass <> "" Then
MsgBox "  Password not Required for Anonymous Account  ": DoEvents
txtPass = ""
End If
GoTo anony
End If

If txtName = "" Or txtPass = "" Then
MsgBox "  More Information Needed to Complete New Account   ": DoEvents
Exit Sub
End If

anony:
frmMain.User2.ListIndex = -1
frmMain.UserList.ListIndex = -1
frmMain.User2.AddItem txtName
frmMain.UserList.AddItem txtName
frmMain.UsrName = txtName
frmMain.PWord = txtPass
i = UserIDs.Count + 1
UserIDs.No(i).Name = txtName
UserIDs.No(i).Pass = txtPass
UserIDs.Count = i
Me.Hide
End Sub
