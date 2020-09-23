VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   Caption         =   "FTP Server"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm 
      BackColor       =   &H00808080&
      Caption         =   "Server Log"
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
      Height          =   3495
      Left            =   120
      TabIndex        =   86
      Top             =   2640
      Width           =   5175
      Begin VB.CommandButton closeSpy 
         Caption         =   "Close"
         Height          =   255
         Left            =   4140
         TabIndex        =   87
         Top             =   2330
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Clear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   4140
         TabIndex        =   88
         Top             =   285
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtSvrLog 
         BackColor       =   &H00E0E0E0&
         Height          =   1995
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   90
         Text            =   "frmMain.frx":0442
         Top             =   240
         Width           =   4935
      End
      Begin VB.ListBox lstSpy 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         ItemData        =   "frmMain.frx":0467
         Left            =   120
         List            =   "frmMain.frx":0469
         TabIndex        =   89
         Top             =   2280
         Visible         =   0   'False
         Width           =   4935
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808080&
      Height          =   895
      Left            =   3840
      TabIndex        =   84
      Top             =   1640
      Width           =   1400
      Begin VB.TextBox txtCount 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Playbill"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   690
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "0"
         Top             =   160
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   615
         Index           =   2
         Left            =   120
         Picture         =   "frmMain.frx":046B
         Stretch         =   -1  'True
         Top             =   200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":08AD
         Stretch         =   -1  'True
         Top             =   200
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   615
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":0CEF
         Stretch         =   -1  'True
         Top             =   200
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Height          =   1455
      Left            =   3840
      TabIndex        =   80
      Top             =   120
      Width           =   1400
      Begin VB.CheckBox butChoice 
         Caption         =   "Messages"
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1010
         Width           =   1165
      End
      Begin VB.CheckBox butChoice 
         Caption         =   "Accounts"
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   600
         Width           =   1165
      End
      Begin VB.CheckBox butChoice 
         Caption         =   "Users"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   180
         Value           =   1  'Checked
         Width           =   1165
      End
   End
   Begin VB.Frame frm2 
      BackColor       =   &H00808080&
      Caption         =   "Online Users"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   78
      Top             =   120
      Width           =   3615
      Begin VB.ListBox lstConned 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         ItemData        =   "frmMain.frx":1131
         Left            =   120
         List            =   "frmMain.frx":1133
         TabIndex        =   79
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   11160
      Top             =   120
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Left            =   8760
      ScaleHeight     =   195
      ScaleWidth      =   1575
      TabIndex        =   15
      Top             =   6240
      Width           =   1635
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   6960
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   14
      Top             =   6240
      Width           =   1815
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   6900
      TabIndex        =   13
      Top             =   6240
      Width           =   6960
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Listening..."
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Accounts"
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
      Height          =   5955
      Index           =   1
      Left            =   5400
      TabIndex        =   1
      Tag             =   "SECOND"
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   91
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00808080&
         Caption         =   "Groups"
         ForeColor       =   &H00FFFFFF&
         Height          =   1755
         Left            =   120
         TabIndex        =   44
         Top             =   3810
         Width           =   4695
         Begin VB.CommandButton cmdTmSet 
            Caption         =   "Set Times"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   235
            Left            =   3720
            TabIndex        =   68
            Top             =   750
            Width           =   855
         End
         Begin VB.CheckBox tmLimit 
            BackColor       =   &H00808080&
            Caption         =   "Time Limit"
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2160
            TabIndex        =   67
            Top             =   750
            Width           =   1095
         End
         Begin VB.CheckBox Restrik 
            BackColor       =   &H00808080&
            Caption         =   "No Restrictions"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2160
            TabIndex        =   63
            Top             =   1350
            Width           =   1695
         End
         Begin VB.CheckBox AccDis 
            BackColor       =   &H00808080&
            Caption         =   "Account Dis-Abled"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2160
            TabIndex        =   50
            Top             =   200
            Width           =   2295
         End
         Begin VB.CheckBox Brws 
            BackColor       =   &H00808080&
            Caption         =   "Browse All Drives"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2160
            TabIndex        =   49
            Top             =   1050
            Width           =   2415
         End
         Begin VB.CheckBox rel 
            BackColor       =   &H00808080&
            Caption         =   "Show Path Relative     ""/"""
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2160
            TabIndex        =   48
            Top             =   480
            Width           =   2415
         End
         Begin VB.CommandButton Command6 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   47
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton CmdAddGrp 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   46
            Top             =   420
            Width           =   375
         End
         Begin VB.ListBox Grps 
            Height          =   1425
            ItemData        =   "frmMain.frx":1135
            Left            =   120
            List            =   "frmMain.frx":1137
            TabIndex        =   45
            Top             =   230
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00808080&
         Caption         =   "Access"
         ForeColor       =   &H00FFFFFF&
         Height          =   2505
         Left            =   120
         TabIndex        =   32
         Top             =   1230
         Width           =   4695
         Begin VB.CheckBox dRemove 
            BackColor       =   &H00808080&
            Caption         =   "Remove"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3600
            TabIndex        =   66
            Top             =   1800
            Width           =   1020
         End
         Begin VB.CheckBox dMake 
            BackColor       =   &H00808080&
            Caption         =   "Make"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3600
            TabIndex        =   65
            Top             =   1500
            Width           =   975
         End
         Begin VB.CheckBox fTrans 
            BackColor       =   &H00808080&
            Caption         =   "Transfer"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   64
            Top             =   1800
            Width           =   975
         End
         Begin VB.CommandButton RmDir 
            Caption         =   "Remove Directory"
            Height          =   295
            Left            =   3000
            TabIndex        =   43
            Top             =   2140
            Width           =   1455
         End
         Begin VB.CommandButton AddDir 
            Caption         =   "Add Directory"
            Height          =   295
            Left            =   240
            TabIndex        =   42
            Top             =   2140
            Width           =   1455
         End
         Begin VB.CheckBox DSub 
            BackColor       =   &H00808080&
            Caption         =   "Sub Dirs"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3600
            TabIndex        =   40
            Top             =   840
            Width           =   945
         End
         Begin VB.CheckBox FEx 
            BackColor       =   &H00808080&
            Caption         =   "Execute"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   39
            Top             =   1500
            Width           =   895
         End
         Begin VB.CheckBox FDelete 
            BackColor       =   &H00808080&
            Caption         =   "Delete"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   38
            Top             =   1170
            Width           =   855
         End
         Begin VB.CheckBox FWrite 
            BackColor       =   &H00808080&
            Caption         =   "Chng Dir"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3600
            TabIndex        =   37
            Top             =   1170
            Width           =   975
         End
         Begin VB.CheckBox DList 
            BackColor       =   &H00808080&
            Caption         =   "List"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   36
            Top             =   840
            Width           =   855
         End
         Begin VB.ListBox AccsList 
            Height          =   1230
            ItemData        =   "frmMain.frx":1139
            Left            =   120
            List            =   "frmMain.frx":113B
            TabIndex        =   35
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton cmdHome 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4260
            TabIndex        =   34
            Top             =   490
            Width           =   300
         End
         Begin VB.TextBox HomeDir 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label lblBelong 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1680
            TabIndex        =   77
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label3 
            BackColor       =   &H00808080&
            Caption         =   "Home Directory"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox PWord 
         Height          =   285
         Left            =   2880
         TabIndex        =   31
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox UsrName 
         Height          =   285
         Left            =   2880
         TabIndex        =   29
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton RmUser 
         Caption         =   "Remove User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   295
         Left            =   3120
         TabIndex        =   27
         Top             =   5610
         Width           =   1575
      End
      Begin VB.CommandButton AddUser 
         Caption         =   "Add User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   295
         Left            =   240
         TabIndex        =   26
         Top             =   5610
         Width           =   1575
      End
      Begin VB.ListBox User2 
         Height          =   840
         ItemData        =   "frmMain.frx":113D
         Left            =   120
         List            =   "frmMain.frx":113F
         TabIndex        =   25
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Pass"
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
         Left            =   2160
         TabIndex        =   30
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label7 
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
         Left            =   2160
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Users"
      ForeColor       =   &H00FFFFFF&
      Height          =   6015
      Index           =   0
      Left            =   5400
      TabIndex        =   0
      Tag             =   "FIRST"
      Top             =   120
      Width           =   4935
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Caption         =   "Hidden Directories"
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
         Height          =   1175
         Left            =   120
         TabIndex        =   69
         Top             =   4760
         Width           =   4695
         Begin VB.CheckBox chkHidden 
            BackColor       =   &H00808080&
            Caption         =   "Use Hidden"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3480
            TabIndex        =   76
            Top             =   10
            Width           =   1140
         End
         Begin VB.CheckBox chkDelete 
            BackColor       =   &H00808080&
            Caption         =   "Beep on File Delete"
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
            Left            =   2400
            TabIndex        =   75
            Top             =   810
            Width           =   2175
         End
         Begin VB.CheckBox chkAccess 
            BackColor       =   &H00808080&
            Caption         =   "Beep on Attemp"
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
            Left            =   2400
            TabIndex        =   74
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkAdmin 
            BackColor       =   &H00808080&
            Caption         =   "Allow Administrator"
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
            Left            =   2400
            TabIndex        =   73
            Top             =   260
            Width           =   2055
         End
         Begin VB.CommandButton Command4 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1920
            TabIndex        =   72
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmdHidDir 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1920
            TabIndex        =   71
            Top             =   270
            Width           =   375
         End
         Begin VB.ListBox lstHidden 
            BackColor       =   &H00E0E0E0&
            Height          =   840
            ItemData        =   "frmMain.frx":1141
            Left            =   120
            List            =   "frmMain.frx":1143
            TabIndex        =   70
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame frmServ 
         BackColor       =   &H00808080&
         Caption         =   "Server Access"
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
         Height          =   2055
         Left            =   120
         TabIndex        =   51
         Top             =   2640
         Width           =   4695
         Begin VB.CheckBox AllowAnon 
            BackColor       =   &H00808080&
            Caption         =   "Allow Anonymous Connections"
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
            Left            =   360
            TabIndex        =   62
            Top             =   510
            Width           =   4215
         End
         Begin VB.CheckBox DenAll 
            BackColor       =   &H00808080&
            Caption         =   "Deny ALL logins (except Administrator)"
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
            Left            =   360
            TabIndex        =   61
            Top             =   240
            Width           =   4215
         End
         Begin VB.PictureBox picPretty 
            Height          =   375
            Left            =   2880
            ScaleHeight     =   315
            ScaleWidth      =   1635
            TabIndex        =   57
            Top             =   840
            Width           =   1695
            Begin VB.Label lblLocalIP2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   30
               TabIndex        =   58
               Top             =   15
               Width           =   1575
            End
         End
         Begin VB.TextBox maxUnits 
            Height          =   285
            Left            =   2880
            TabIndex        =   55
            Text            =   "10"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox LisPort 
            Height          =   285
            Left            =   2880
            TabIndex        =   53
            Text            =   "21"
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblLocalIP1 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   360
            TabIndex        =   56
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label lblMax 
            BackStyle       =   0  'Transparent
            Caption         =   "MAXIMUM USERS:"
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
            Left            =   360
            TabIndex        =   54
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label lblPort 
            BackStyle       =   0  'Transparent
            Caption         =   "PORT TO LISTEN ON:"
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
            Left            =   360
            TabIndex        =   52
            Top             =   1320
            Width           =   2415
         End
      End
      Begin VB.Frame frmGrp 
         BackColor       =   &H00808080&
         Caption         =   "Belong to User Group"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   1960
         Width           =   4695
         Begin VB.CommandButton cmdGroup 
            Caption         =   "Edit"
            Height          =   305
            Left            =   3630
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox BelongGrp 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.ListBox UserList 
         BackColor       =   &H00E0E0E0&
         Height          =   1230
         ItemData        =   "frmMain.frx":1145
         Left            =   300
         List            =   "frmMain.frx":1147
         TabIndex        =   16
         Top             =   590
         Width           =   4335
      End
      Begin VB.Frame frmClient 
         BackColor       =   &H00808080&
         Height          =   1740
         Left            =   180
         TabIndex        =   17
         Top             =   190
         Width           =   4575
         Begin VB.CheckBox GrpDisabled 
            BackColor       =   &H00808080&
            Caption         =   "Disable Account"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Messages"
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
      Height          =   6015
      Index           =   2
      Left            =   5400
      TabIndex        =   2
      Tag             =   "THIRD"
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtChgdir 
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Text            =   "Where were you going?"
         Top             =   1920
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PM"
         Height          =   305
         Left            =   4320
         TabIndex        =   12
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox txtPriv 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "                           Private Message"
         Top             =   4800
         Width           =   4215
      End
      Begin VB.TextBox txtOff 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Text            =   "Do stop by again"
         Top             =   4440
         Width           =   4695
      End
      Begin VB.TextBox txtOff 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Text            =   "Have a great day from HomePlay"
         Top             =   4140
         Width           =   4695
      End
      Begin VB.TextBox txtOff 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Text            =   "Thanks for stopping by"
         Top             =   3840
         Width           =   4695
      End
      Begin VB.TextBox txtWel 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Text            =   "Or leave what you have..."
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox txtWel 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Text            =   "We hope you find what you need..."
         Top             =   900
         Width           =   4695
      End
      Begin VB.TextBox txtWel 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Text            =   "Welcome to HomePlay Entertainments FTP-Server"
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label lblChgDir 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Directory Message:"
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
         TabIndex        =   59
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label lblOff 
         BackStyle       =   0  'Transparent
         Caption         =   "Logoff Message:"
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
         Left            =   360
         TabIndex        =   4
         Top             =   3570
         Width           =   1695
      End
      Begin VB.Label lblWel 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome Message:"
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
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSv 
         Caption         =   "Save for Log"
         Begin VB.Menu mnuSvSvr 
            Caption         =   "Server Log"
         End
         Begin VB.Menu ln6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSvSpy 
            Caption         =   "Spy List"
         End
         Begin VB.Menu ln7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSvAll 
            Caption         =   "All"
         End
      End
      Begin VB.Menu ln5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu mnuDel 
         Caption         =   "Delete"
         Begin VB.Menu mnuRecycle 
            Caption         =   "Recycle"
            Begin VB.Menu mnuB 
               Caption         =   "To Recycle Bin"
               Begin VB.Menu mnuConfirm 
                  Caption         =   "With Confirmation"
               End
               Begin VB.Menu ln12 
                  Caption         =   "-"
               End
               Begin VB.Menu mnuSilent 
                  Caption         =   "Without Confirm"
                  Checked         =   -1  'True
               End
            End
            Begin VB.Menu ln14 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMove 
               Caption         =   "Move to Different Dir"
            End
         End
         Begin VB.Menu ln10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuKill 
            Caption         =   "Kill"
         End
      End
      Begin VB.Menu ln11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStartSvr 
         Caption         =   "&Start Server"
      End
      Begin VB.Menu mnuStopSvr 
         Caption         =   "S&top  Server"
         Enabled         =   0   'False
      End
      Begin VB.Menu ln4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update Record"
      End
      Begin VB.Menu ln9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResize 
         Caption         =   "Resize"
      End
   End
   Begin VB.Menu mnuHel 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "H&elp"
      End
      Begin VB.Menu ln3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Defaults"
      End
      Begin VB.Menu ln2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu xx 
      Caption         =   "PopMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuKick 
         Caption         =   "Kick"
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpy 
         Caption         =   "Spy"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
       Dim j As Integer, drivcount As Integer, qu As Integer
Public WithEvents FTPServer As Server
Attribute FTPServer.VB_VarHelpID = -1
Public Function Get_closing() As String
Dim str As String
If frmMain.txtOff(0).Text = "" Then
   str = "Closing Connection, Goodbye."
   Else
   str = frmMain.txtOff(0).Text
   End If
   If frmMain.txtOff(1).Text <> "" Then
   str = str & ",  " & frmMain.txtOff(1).Text
   End If
   If frmMain.txtOff(2).Text <> "" Then
   str = str & ",  " & frmMain.txtOff(2).Text
   End If
   Get_closing = str
End Function

Private Function Driver()
        allDrives$ = VBGetLogicalDriveStrings()
        Do Until allDrives$ = Chr$(0)
        j = j + 1
       currDrive$ = StripNulls$(allDrives$)
       If currDrive$ = "a:\" Then
       ad = True
       End If
       If currDrive$ = "b:\" Then
       bd = True
       End If
       If currDrive$ = "c:\" Then
       cd = True
       End If
       If currDrive$ = "d:\" Then
       dd = True
       End If
       If currDrive$ = "e:\" Then
       ed = True
       End If
       If currDrive$ = "f:\" Then
       fd = True
       End If
       If currDrive$ = "g:\" Then
       gd = True
       End If
       If currDrive$ = "h:\" Then
       hd = True
       End If
       If currDrive$ = "i:\" Then
       id = True
       End If
       If currDrive$ = "j:\" Then
       id = True
       End If
       If currDrive$ = "k:\" Then
       id = True
       End If
       '''     'get the drive type
       drvType$ = rgbGetDriveType(currDrive$)
       If drvType$ = "Floppy drive." Then
       GoTo skip
       End If
       If drvType$ = "CD-ROM drive." Then
       GoTo skip
       End If
       qu = qu + 1
skip:
        Loop
End Function
Private Function VBGetLogicalDriveStrings() As String

        Dim r As Long
        Dim i As Integer
        Dim tmp As String
        
        tmp$ = Space$(64)
        
        r& = GetLogicalDriveStrings(Len(tmp$), tmp$)
        
        VBGetLogicalDriveStrings = Trim$(tmp$)
End Function
Private Function rgbGetDriveType(RootPathName$) As String
        
        Dim r As Long
        
        r& = GetDriveType(RootPathName$)
        
        Select Case r&
       Case 0: rgbGetDriveType$ = "The drive type cannot be determined."
       Case 1: rgbGetDriveType$ = "The root directory does not exist."
       Case DRIVE_REMOVABLE:
        Select Case Left$(RootPathName$, 1)
        Case "a", "b": rgbGetDriveType$ = "Floppy drive."
        Case Else: rgbGetDriveType$ = "Removable drive."
        End Select
       Case DRIVE_FIXED: rgbGetDriveType$ = "Hard drive; can not be removed."
       Case DRIVE_REMOTE: rgbGetDriveType$ = "Remote (network) drive."
       Case DRIVE_CDROM: rgbGetDriveType$ = "CD-ROM drive."
       Case DRIVE_RAMDISK: rgbGetDriveType$ = "RAM disk."
        End Select
        
End Function
Private Function StripNulls(startStrg$) As String
        Dim c As Integer
        Dim Item As String
        
        c% = 1
        
        Do

              If Mid$(startStrg$, c%, 1) = Chr$(0) Then
                      
                      Item$ = Mid$(startStrg$, 1, c% - 1)
                      startStrg$ = Mid$(startStrg$, c% + 1, Len(startStrg$))
                      StripNulls$ = Item$
                      Exit Function
              End If

       c% = c% + 1
        Loop
End Function

Private Sub AccDis_Click()
GrpChange = True
End Sub

Private Sub AccsList_Click()
Dim X As Integer, Z As Integer
  aItem = AccsList.ListIndex
  Debug.Print "Access List Item = " & aItem
  ClearAccs
  Z = aItem + 1
  Debug.Print UserIDs.No(uUser).Priv(Z).Accs
  If InStr(Privs(Z).Accs, "W") Then
    FWrite.Value = 1
  End If
  If InStr(Privs(Z).Accs, "D") Then
    FDelete.Value = 1
  End If
  If InStr(Privs(Z).Accs, "X") Then
    FEx.Value = 1
  End If
  If InStr(Privs(Z).Accs, "L") Then
    DList.Value = 1
  End If
  If InStr(Privs(Z).Accs, "S") Then
    DSub.Value = 1
  End If
  If InStr(Privs(Z).Accs, "M") Then
    dMake.Value = 1
  End If
  If InStr(Privs(Z).Accs, "H") Then
    dRemove.Value = 1
  End If
  If InStr(Privs(Z).Accs, "T") Then
    fTrans.Value = 1
  End If
  
End Sub

Private Sub AccsList_DblClick()
MsgBox "Directory Name Is:" & vbCrLf & AccsList.Text
End Sub

Private Sub AddDir_Click()
Dim ret As String, cnt As Integer
ret = BrowseForFolder(Me.hWnd, "Add a Directory for Access")
If ret = "" Then Exit Sub
AccsList.AddItem ret
Pcnt = Pcnt + 1
UserIDs.No(uUser).Priv(Pcnt).Path = ret
lstchange = True
End Sub

Private Sub AddUser_Click()
frmAddUser.Show
End Sub

Private Sub BelongGrp_Change()
CliGrpChange = True
End Sub

Private Sub Brws_Click()
GrpChange = True
End Sub

Private Sub butChoice_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
GrpDisabled.Value = 0
Select Case Index
Case "0"
Frame1(0).Visible = True
Frame1(1).Visible = False
Frame1(2).Visible = False
butChoice(1).Value = 0
butChoice(2).Value = 0
Case "1"
Frame1(0).Visible = False
Frame1(1).Visible = True
Frame1(2).Visible = False
butChoice(0).Value = 0
butChoice(2).Value = 0
Case "2"
Frame1(0).Visible = False
Frame1(1).Visible = False
Frame1(2).Visible = True
butChoice(0).Value = 0
butChoice(1).Value = 0
End Select
End If
End Sub

Private Sub chkHome_Click()
GrpChange = True
End Sub

Private Sub Clear_Click()
txtSvrLog.Text = ""
End Sub

Private Sub closeSpy_Click()
txtSvrLog.Height = 3075
closeSpy.Visible = False
lstSpy.Clear
lstSpy.Visible = False
spy_client = False
End Sub

Private Sub CmdAddGrp_Click()
If UserList.SelCount = 0 Then Exit Sub
Load frmGroup
frmGroup.Show
End Sub

Private Sub cmdGroup_Click()
Load frmGroup
frmGroup.Show
End Sub

Private Sub cmdHidDir_Click()
Dim ret, vet, str As String, man As Integer
Dim Parts() As String
man = 0
ret = BrowseForFolder(Me.hWnd, "Select a Directory to Hide")
ret = ret & "\Explorer.exe"
man = CountStr(ret, "\")
vet = Parse2Array(ret, Parts(), "\")
str = Parts(man - 1)

If str <> "" Then
lstHidden.AddItem str
lstHidAdd = True
End If

End Sub

Private Sub cmdHome_Click()
Dim ret
ret = BrowseForFolder(Me.hWnd, "Select a Home Directory")
If ret = "" Then Exit Sub
HomeDir = ret
End Sub

Private Sub Command1_Click()
'"220 " & txtpriv.text
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
If lstHidden.ListIndex <> -1 Then
lstHidden.RemoveItem (lstHidden.ListIndex)
lstHidRemove = True
End If
End Sub

Private Sub Command5_Click()
frmNameGrp.Tag = ""
frmNameGrp.Show 1
If frmNameGrp.Tag = "" Then Exit Sub
GrpChange = True
MsgBox frmNameGrp.Tag: DoEvents
Unload frmNameGrp
End Sub

Private Sub Command6_Click()
If Grps.SelCount = 0 Then Exit Sub
'Grps.RemoveItem
MsgBox "Still got to Remove This  " & Grps.Text: DoEvents
End Sub

Private Sub DList_Click()
lstchange = True
End Sub

Private Sub dMake_Click()
lstchange = True
End Sub

Private Sub dRemove_Click()
lstchange = True
End Sub

Private Sub DSub_Click()
lstchange = True
End Sub

Private Sub FDelete_Click()
lstchange = True
End Sub

Private Sub FEx_Click()
lstchange = True
End Sub

Private Sub Form_Activate()
Driver
End Sub

Private Sub Form_Load()
App.HelpFile = (App.Path & "\" & "hpserver.hlp")
    Set FTPServer = New Server
    Set frmWinsock.FTPServer = FTPServer
lblDate.Caption = " " & Date
If LoadProfile(App.Path & "\ftp_srv.ini") Then
End If
Dim X As Integer, Y As Integer
  Y = UserIDs.Count
  If (Y > 0) Then
    For X = 1 To UserIDs.Count
      UserList.AddItem UserIDs.No(X).Name
      User2.AddItem UserIDs.No(X).Name
    Next
  End If
  aItem = -1
  uItem = -1
  lblLocalIP1.Caption = "Local IP for " & UCase(frmWinsock.CommandSock(0).LocalHostName)
  lblLocalIP2.Caption = frmWinsock.CommandSock(0).LocalIP
  CliChange = False
  GrpChange = False
  lstchange = False
  CliGrpChange = False
  lstHidAdd = False
  lstHidRemove = False
  found = False
  halt_transfer = False
  del_path = ""
  txtSvrLog.Height = 3075
  Oncount = 0
  txtCount.Text = 0
  spy_client = False
  file_is_open = False
  requested = False
End Sub

Public Sub Form_Resize()

    On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
    frmMain.Height = 7185
    frmMain.Width = 10515 Or frmMain.Width = 5475

End Sub

Public Sub Form_Unload(Cancel As Integer)
    StopServer

    Set FTPServer = Nothing
    Set frmWinsock.FTPServer = Nothing

    Unload frmWinsock
    Unload Me
    Set frmWinsock = Nothing
    Set frmMain = Nothing
    End

End Sub

Private Sub fTrans_Click()
lstchange = True
End Sub

Private Sub FWrite_Click()
lstchange = True
End Sub

Private Sub GrpDisabled_Click()
CliChange = True
End Sub

Private Sub Grps_Click()
Dim str As String, Ctr As Integer
str = Grps.Text
If grpnum > 0 Then
   For Ctr = 1 To grpnum
   If str = group_info.GrpType(Ctr).Name Then
   str = group_info.GrpType(Ctr).Access
   End If
   Next
End If

ClearGrps
Dim Z As Integer
aItem = Grps.ListIndex
  Debug.Print "Group List Item = " & aItem
  Z = aItem + 1
If InStr(str, "B") Then
    Brws.Value = 1
  End If
  If InStr(str, "R") Then
    rel.Value = 1
  End If
  If InStr(str, "Q") Then
    AccDis.Value = 1
  End If
  If InStr(str, "P") Then
    Restrik.Value = 1
  End If
  If InStr(str, "A") Then
    tmLimit.Value = 1
  End If
End Sub

Private Sub HomeDir_Change()
lstchange = True
End Sub

Private Sub lstConned_Click()
Dim str As String
str = lstConned.Text
spot1 = Parse(str, 1)
spot2 = Parse(str, 2)
spot3 = Parse(str, 3)
End Sub

Private Sub lstConned_DblClick()
Dim u As Integer, Z As Integer, tmp1 As String, tmp2 As String, tmp3 As String
'tmp3 = ",  "

For u = 0 To Number
If UserIDs.No(u).id = spot1 And UserIDs.No(u).Name = spot2 And UserIDs.No(u).IP = spot3 Then
tmp1 = client(UserIDs.No(u).id).Current_Access
GoTo getword
End If
Next
Exit Sub

getword:

For Z = 1 To Len(tmp1)
tmp2 = Mid(tmp1, Z, 1)
tmp2 = Convert_to_Word(tmp2)
tmp3 = tmp3 & "     " & tmp2 & vbCrLf
Next
tmp3 = "Access to Current Drive is:" & vbCrLf & tmp3
'MsgBox tmp3


MsgBox "Information For:  " & UCase(UserIDs.No(u).Name) & vbCrLf & "Current IP:  " & UserIDs.No(u).IP & vbCrLf & "Connected at:  " & client(UserIDs.No(u).id).ConnectedAt & vbCrLf & "Idle Since:  " & client(UserIDs.No(u).id).IdleSince & vbCrLf & "Uploaded Bytes:  " & client(UserIDs.No(u).id).TotalBytesUploaded & vbCrLf & "Downloaded Bytes:  " & client(UserIDs.No(u).id).TotalBytesDownloaded & vbCrLf & "Current Directory:  " & client(UserIDs.No(u).id).CurrentDir & vbCrLf & "Group Name: " & client(UserIDs.No(u).id).Group_Name & vbCrLf & "Group Access:  " & client((UserIDs.No(u).id)).group_Access & vbCrLf & "Intial Directory Access = " & client(UserIDs.No(u).id).Priv(1).Accs & vbCrLf & vbCrLf & tmp3    ' "Current Access:  " & client(UserIDs.No(u).id).Current_Access
End Sub

Private Sub lstConned_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstConned.SelCount >= 1 And Button = 2 Then
Me.PopupMenu xx
End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuConfirm_Click()
mnuSilent.Checked = False
mnuKill.Checked = False
mnuMove.Checked = False
mnuConfirm.Checked = True
End Sub

Private Sub mnuExit_Click()
On Error GoTo dunn
Dim Can As Boolean, decision As String, message As String, many As Integer
message = "220 Server is Closing..."
If ServerActive = True Then
If ConnectedClients > 0 Then
decision = MsgBox("There are " & ConnectedClients & " clients connected" & ", " & "do you still want to shut down the server ?", _
    vbOKCancel + vbQuestion, "Connections Detected"): DoEvents
If decision = vbOK Then
    Can = False
Else

    Can = True
    GoTo dunn
End If
End If

For many = 0 To ConnectedClients
If frmWinsock.CommandSock(many).State = sckConnected Then
frmWinsock.CommandSock(many).SendData message
End If
Call Wait(0.25)
FTPServer.LogoutClient client(many).id
Next

Call Wait(0.25)
Call mnuStopSvr_Click
End If
dunn:
Call Wait(0.5)
End
End Sub

Private Sub mnuHelp_Click()
vHelp& = WinHelp(frmMain.hWnd, App.HelpFile, HELP_INDEX, CLng(0))
End Sub

Private Sub mnuKick_Click()
Load frmKick
Dim u As Integer
For u = 0 To ConnectedClients
If spot2 = client(u).UserName And spot1 = client(u).id Then
frmKick.lblName.Caption = client(u).UserName
frmKick.lblIP.Caption = client(u).IPAddress
frmKick.lblCon.Caption = client(u).ConnectedAt
frmKick.idleTime.Caption = client(u).IdleSince
frmKick.Tag = client(u).id
GoTo try
End If
Next
Exit Sub
try:
frmKick.Show
End Sub

Private Sub mnuKill_Click()
mnuMove.Checked = False
mnuSilent.Checked = False
mnuConfirm.Checked = False
mnuKill.Checked = True
End Sub

Private Sub mnuMove_Click()
Dim ret As String
Dim retval As Boolean
mnuSilent.Checked = False
mnuConfirm.Checked = False
mnuKill.Checked = False
mnuMove.Checked = True
ret = BrowseForFolder(Me.hWnd, "Select a Deleted to Directory, Cancel for Default")

If ret = "" Then
ret = App.Path & "\Deleted"
End If

retval = Validate_Directory(ret)

If retval = False Then
MkDir ret
End If

del_path = ret & "\"
End Sub

Private Sub mnuResize_Click()
If Me.Width = 10515 Then
Me.Width = 5475
Exit Sub
End If
If Me.Width = 5475 Then
Me.Width = 10515
Exit Sub
End If
End Sub

Private Sub mnuRestore_Click()
Dim ans As Integer, Can As Boolean
ans = MsgBox("This will create a new initialization file with default settings" & vbCrLf & "Click on OK to proceed", _
    vbOKCancel + vbQuestion, "Replace Defaults")
If ans = vbOK Then
    Can = False
Else

    Can = True
    Exit Sub
End If

CreateDefault
End Sub

Private Sub mnuSilent_Click()
mnuConfirm.Checked = False
mnuKill.Checked = False
mnuMove.Checked = False
mnuSilent.Checked = True
'FOF_FLAGS = FOF_NOCONFIRMATION
End Sub

Private Sub mnuSpy_Click()
spy_client = True
txtSvrLog.Height = 1995
lstSpy.Visible = True
lstSpy.Clear
closeSpy.Visible = True
Dim u As Integer

u = 1
Do Until u > ConnectedClients
UserIDs.No(u).Spy = False
u = u + 1
Loop

For u = 0 To ConnectedClients
If spot2 = client(u).UserName And spot1 = client(u).id Then
UserIDs.No(u).Spy = True
lstSpy.AddItem " *** Now Tracking ***"
lstSpy.AddItem " Name =  " & client(u).UserName
lstSpy.AddItem " Password =  " & client(u).Password
lstSpy.AddItem " IP Address =  " & client(u).IPAddress
lstSpy.AddItem " Group =  " & client(u).Group_Name
lstSpy.AddItem " Connected At =  " & client(u).ConnectedAt
lstSpy.AddItem " Idle Since =  " & client(u).IdleSince
lstSpy.AddItem " Home Directory =  " & client(u).HomeDir
lstSpy.AddItem " Current Directory =  " & client(u).CurrentDir
lstSpy.AddItem " Downloaded Bytes =  " & client(u).TotalBytesDownloaded
lstSpy.AddItem " Uploaded Bytes =  " & client(u).TotalBytesUploaded
lstSpy.AddItem " Downloaded Files =  " & client(u).TotalFilesDownloaded
lstSpy.AddItem " Uploaded Files =  " & client(u).TotalFilesUploaded
End If
Next

End Sub

Private Sub mnuStartSvr_Click()
LoadProfile (App.Path & "\ftp_srv.ini")
Dim X As Integer, Y As Integer
  Y = UserIDs.Count
  If (Y > 0) Then
    For X = 1 To UserIDs.Count
      UserList.AddItem UserIDs.No(X).Name
      User2.AddItem UserIDs.No(X).Name
    Next
  End If
Dim hd As Integer
Oncount = 0
hidDir_list = " "
If ServerActive = True Then Exit Sub
txtSvrLog.Text = "Starting Server..."
lblMessage.Caption = "Starting Server..."
StartServer
If ServerActive = False Then
mnuStartSvr.Enabled = True
mnuStopSvr.Enabled = False
Exit Sub
End If
Image1(0).Visible = False
Image1(2).Visible = True
Call Wait(0.75)
lblMessage.Caption = "Service Started..."

If chkHidden.Value = 1 Then
For hd = 0 To lstHidden.ListCount - 1
lstHidden.ListIndex = hd
hidDir_list = hidDir_list & lstHidden.Text & " "
Next
hidDir_list = UCase(hidDir_list)
End If

Image1(2).Visible = False
Image1(1).Visible = True
Clear.Visible = True
mnuStartSvr.Enabled = False
mnuStopSvr.Enabled = True
End Sub

Private Sub mnuStopSvr_Click()
Image1(1).Visible = False
txtSvrLog.Text = "Shutting Down Server..."
lblMessage.Caption = "Shutting Down Server..."
Image1(2).Visible = True
Call Wait(0.75)
lblMessage.Caption = "Not Listening..."
Image1(2).Visible = False
Image1(0).Visible = True
Clear.Visible = False
    StopServer
mnuStopSvr.Enabled = False
mnuStartSvr.Enabled = True
End Sub
Private Sub FTPServer_ServerStarted()

    WriteToLogWindow "Server started!", True

End Sub

Private Sub FTPServer_ServerStopped()

    WriteToLogWindow "Server stopped!", True

End Sub

Private Sub FTPServer_ServerErrorOccurred(ByVal errNumber As Long)

    MsgBox FTPServer.ServerGetErrorDescription(errNumber), vbInformation, "Error occured!": DoEvents

End Sub

Private Sub FTPServer_NewClient(ByVal ClientID As Long)

    WriteToLogWindow "Client " & ClientID & " connected! (" & FTPServer.GetClientIPAddress(ClientID) & ")", True

End Sub

Private Sub FTPServer_ClientSentCommand(ByVal ClientID As Long, Command As String, Args As String)

    WriteToLogWindow "Client " & ClientID & " sent: " & Command & " " & Args, True

End Sub

Private Sub FTPServer_ClientStatusChanged(ByVal ClientID As Long)


    WriteToLogWindow "Client " & ClientID & " Status: " & FTPServer.GetClientStatus(ClientID), True

End Sub

Private Sub FTPServer_ClientLoggedOut(ByVal ClientID As Long)
    
    WriteToLogWindow "Client " & ClientID & " logged out!", True

End Sub

Private Sub mnuSvAll_Click()
Call mnuSvSvr_Click
Call mnuSvSpy_Click
MsgBox "Files have been saved in" & vbCrLf & "----------------------------------------" & vbCrLf & App.Path & vbCrLf & vbCrLf & "Named ServerLog.txt and SpyLog.txt": DoEvents
End Sub

Private Sub mnuSvSpy_Click()
Dim f As Integer, lst As String
Open (App.Path & "\SpyLog.txt") For Output As #6
Print #6, "HPFtp Server Spy Log"
Print #6, "Created - " & Time & "  " & Date
Print #6,
For f = 0 To lstSpy.ListCount - 1
lstSpy.ListIndex = f
lst = lstSpy.Text
Print #6, lst
Next
Close #6
End Sub

Private Sub mnuSvSvr_Click()
Open (App.Path & "\ServerLog.txt") For Output As #6
Print #6, "HpFtp Server Log"
Print #6, "Created - " & Time & "  " & Date
Print #6,
Print #6, txtSvrLog.Text
Close #6
End Sub

Private Sub mnuUpdate_Click()
Dim Z As Integer, Ctr As Integer, S As String, quim As String, h As String
Dim tru As Integer
Unload frmAddUser

'If BinChange = True Then
'If mnuBin.Checked = True Then
'WritePrivateProfileString "Settings", "UseBin", "Yes", (App.Path & "\ftp_srv.ini")
'Else
'WritePrivateProfileString "Settings", "UseBin", "No", (App.Path & "\ftp_srv.ini")
'End If
'Exit Sub
'End If

If lstHidAdd = True Or lstHidRemove = True Then
tru = GetFromIni("Common", "Hidden", (App.Path & "\ftp_srv.ini"))
  Z = frmMain.lstHidden.ListCount
  S = Z
  If Z = -1 Or Z = 0 Then Exit Sub
  For Ctr = 1 To Z
  frmMain.lstHidden.ListIndex = Ctr - 1
  h = frmMain.lstHidden.Text
  WritePrivateProfileString "Settings", "Date", h, (App.Path & "\ftp_srv.ini")
  WritePrivateProfileString "Settings", "Time", h, (App.Path & "\ftp_srv.ini")
  WritePrivateProfileString "Common", "Hid" & Ctr, h, (App.Path & "\ftp_srv.ini")
  WritePrivateProfileString "Common", "Hidden", S, (App.Path & "\ftp_srv.ini")
  Next
  
  lstHidAdd = False
  lstHidRemove = False
Exit Sub
End If

If CliGrpChange = True And CliChange = False And GrpChange = False Then
  UserIDs.No(uUser).Name = UsrName
  UserIDs.No(uUser).Pass = PWord
  UserIDs.No(uUser).Home = HomeDir
  UserIDs.No(uUser).Pcnt = Pcnt
  UserIDs.No(uUser).Group = BelongGrp.Text
  If SaveProfile(App.Path & "\ftp_srv.ini", True) Then
  End If
Exit Sub
End If

  UserIDs.No(uUser).Name = UsrName
  UserIDs.No(uUser).Pass = PWord
  UserIDs.No(uUser).Home = HomeDir
  UserIDs.No(uUser).Pcnt = Pcnt
  UserIDs.No(uUser).Group = BelongGrp.Text
  
  If lstchange = True Then
  S = ""
  Z = aItem + 1
  If FWrite.Value = 1 Then S = S & "W"
  If FDelete.Value = 1 Then S = S & "D"
  If FEx.Value = 1 Then S = S & "X"
  If DList.Value = 1 Then S = S & "L"
  If DSub.Value = 1 Then S = S & "S"
  If fTrans.Value = 1 Then S = S & "T"
  If dMake.Value = 1 Then S = S & "M"
  If dRemove.Value = 1 Then S = S & "H"
  Privs(Z).Accs = S
  End If
  
  If GrpChange = True Then
  If rel.Value = 1 Then h = h & "R"
  If Brws.Value = 1 Then h = h & "B"
  If AccDis.Value = 1 Then h = h & "Q"
  If Restrik.Value = 1 Then h = h & "P"
  If tmLimit.Value = 1 Then h = h & "A"
  End If

  If CliChange = True Then
  If GrpDisabled.Value = 1 Then
  UserIDs.No(Z).Disabled = "Yes"
  Else
  UserIDs.No(Z).Disabled = "No"
  End If
  End If
  
unch:

  
  If lstchange = True Then    ' added
  UserIDs.No(uUser).Priv(Z).Accs = S
  End If
  
  If GrpChange = True Then    ' added
  group_info.GrpType(Z).Access = h
  End If
  
gtr:
  If SaveProfile(App.Path & "\ftp_srv.ini", True) Then
  End If
  If User2.ListCount > 0 Then
  User2.ListIndex = 0
  End If
End Sub

Private Sub PWord_KeyPress(KeyAscii As Integer)
CliChange = True
End Sub

Private Sub rel_Click()
GrpChange = True
End Sub

Private Sub Restrik_Click()
GrpChange = True
End Sub

Private Sub RmDir_Click()
Dim Z As Integer
  For Z = (aItem + 1) To UserIDs.No(uUser).Pcnt
    UserIDs.No(uUser).Priv(Z).Path = UserIDs.No(uUser).Priv(Z + 1).Path
    UserIDs.No(uUser).Priv(Z).Accs = UserIDs.No(uUser).Priv(Z + 1).Accs
  Next
  UserIDs.No(uUser).Pcnt = UserIDs.No(uUser).Pcnt - 1
  AccsList.RemoveItem (aItem)
lstchange = True
End Sub

Private Sub Timer1_Timer()
lblTime.Caption = Time & " "
End Sub

Private Sub tmLimit_Click()
GrpChange = True
If tmLimit.Value = 1 Then
cmdTmSet.Enabled = True
Else
cmdTmSet.Enabled = False
End If
End Sub

Private Sub User2_Click()
Dim X As Integer, Z As Integer, fre As String
  uItem = User2.ListIndex
  Debug.Print "User List Item = " & uItem
  uUser = uItem + 1
  AccsList.Clear
  ClearAccs
  PWord = ""
  HomeDir = ""
  aItem = -1
  UsrName = UserIDs.No(uUser).Name
  PWord = UserIDs.No(uUser).Pass
  HomeDir = UserIDs.No(uUser).Home
  Pcnt = UserIDs.No(uUser).Pcnt
  fre = UserIDs.No(uUser).Group
  BelongGrp.Text = fre
  
  Dim c As Integer, B As String
  For c = 1 To Number
  B = UserIDs.No(c).Group
  If B = fre Then
  Grps.Text = B
  lblBelong.Enabled = True
  lblBelong.Caption = "Belongs to Group:  " & B
  End If
  Next
  
  For Z = 1 To Pcnt
    Privs(Z).Path = UserIDs.No(uUser).Priv(Z).Path
    Privs(Z).Accs = UserIDs.No(uUser).Priv(Z).Accs
    AccsList.AddItem Privs(Z).Path
  Next
  
  UserList.ListIndex = User2.ListIndex
End Sub
Private Sub ClearAccs()
  FWrite.Value = 0
  FDelete.Value = 0
  FEx.Value = 0
  DList.Value = 0
  DSub.Value = 0
  dMake.Value = 0
  dRemove.Value = 0
  fTrans.Value = 0
End Sub
Private Sub ClearGrps()
AccDis.Value = 0
rel.Value = 0
Brws.Value = 0
Restrik.Value = 0
tmLimit = 0
End Sub

Private Sub UserList_Click()
GrpDisabled.Value = 0
User2.ListIndex = UserList.ListIndex
End Sub

Private Sub UsrName_KeyPress(KeyAscii As Integer)
CliChange = True
End Sub
