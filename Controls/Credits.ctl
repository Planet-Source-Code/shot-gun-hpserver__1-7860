VERSION 5.00
Begin VB.UserControl Credits 
   BackColor       =   &H00808080&
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   ScaleHeight     =   1530
   ScaleWidth      =   3570
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4080
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   4080
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   6480
         Left            =   0
         MouseIcon       =   "Credits.ctx":0000
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Credits.ctx":030A
         Top             =   10
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3615
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "Credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Staterup()
VScroll1.Max = Picture1.Height
VScroll1.min = 0 - Text1.Height
VScroll1.Value = VScroll1.Max
Text1.Top = VScroll1.Value
Text1.Visible = True
Timer1.Enabled = True
End Function
Public Function Killer()
Timer1.Enabled = False
Text1.Visible = False
End Function

Private Sub Timer1_Timer()
If VScroll1.Value >= VScroll1.min + 30 Then
  VScroll1.Value = VScroll1.Value - 20
Else
  VScroll1.Value = VScroll1.Max
  DoEvents
  Timer1.Enabled = False
  Timer2.Enabled = True
End If
Text1.Top = VScroll1.Value
Text1.Visible = True
DoEvents
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
VScroll1.Max = Picture1.Height
VScroll1.min = 0 - Text1.Height
VScroll1.Value = VScroll1.Max
Text1.Top = VScroll1.Value
Text1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub UserControl_Terminate()
Timer1.Enabled = False
End Sub
