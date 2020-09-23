VERSION 5.00
Begin VB.Form frmKick 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kick Client"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "frmKick.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdKick 
      Caption         =   "KICK"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox Check2 
         BackColor       =   &H00808080&
         Caption         =   "Dis-Able this user account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   2400
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "Dis-Able this IP from Connecting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   2100
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   600
         X2              =   3720
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label idleTime 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label lbl4 
         BackColor       =   &H00808080&
         Caption         =   "Idle Since:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblCon 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00808080&
         Caption         =   "Connected At:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblIP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lbl2 
         BackStyle       =   0  'Transparent
         Caption         =   "Current IP:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Do You want to Kick this User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   210
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   3015
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmKick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdKick_Click()
On Error GoTo air
Dim u As Integer, fet As Boolean
For u = 0 To ConnectedClients
If client(u).id = spot1 And spot2 = client(u).UserName Then
frmWinsock.CommandSock(u).SendData ("220 You have been kicked" & vbCrLf)
Call Wait(0.25)
frmMain.FTPServer.LogoutClient client(u).id
GoTo temp
End If
Next
Exit Sub
temp:
Call Wait(0.25)
Me.Hide
Call Wait(0.5)
Unload Me
Exit Sub
air:
If Err.Number = 40006 Then
frmMain.lblMessage.Caption = "Error - client might have got dis-connected"
End If
frmMain.FTPServer.LogoutClient client(u).id
End Sub
