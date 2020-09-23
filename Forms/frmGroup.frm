VERSION 5.00
Begin VB.Form frmGroup 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Double Click to Select"
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
      Height          =   3165
      Left            =   120
      TabIndex        =   0
      Tag             =   "SELECT"
      Top             =   0
      Width           =   2295
      Begin VB.ListBox Grps 
         Height          =   2790
         ItemData        =   "frmGroup.frx":0000
         Left            =   120
         List            =   "frmGroup.frx":0002
         TabIndex        =   1
         Tag             =   "SELECT"
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim r As String, w As Long
Grps.Clear
w = 0
Do Until w = frmMain.Grps.ListCount
frmMain.Grps.ListIndex = w
r = frmMain.Grps.Text
Grps.AddItem r
w = w + 1
Loop
End Sub

Private Sub Grps_DblClick()
CliGrpChange = True
frmMain.BelongGrp.Text = Grps.Text
Call Wait(0.25)
Unload Me
End Sub
