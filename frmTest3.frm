VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmTest3 
   Caption         =   "Test"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000007&
         Caption         =   "Option1"
         ForeColor       =   &H8000000E&
         Height          =   1095
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   2280
         TabIndex        =   2
         Top             =   2400
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   1680
         Width           =   975
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1085
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmTest3.frx":0000
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         Height          =   975
         Left            =   2760
         Top             =   360
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   5880
         X2              =   3720
         Y1              =   0
         Y2              =   1440
      End
      Begin VB.Line Line2 
         X1              =   3840
         X2              =   5760
         Y1              =   360
         Y2              =   2760
      End
      Begin VB.Line Line3 
         X1              =   5880
         X2              =   3720
         Y1              =   1200
         Y2              =   3600
      End
      Begin VB.Line Line4 
         X1              =   4560
         X2              =   4320
         Y1              =   4320
         Y2              =   1800
      End
      Begin VB.Line Line5 
         X1              =   6240
         X2              =   3960
         Y1              =   4560
         Y2              =   2160
      End
      Begin VB.Line Line6 
         X1              =   4440
         X2              =   3960
         Y1              =   480
         Y2              =   4320
      End
   End
End
Attribute VB_Name = "frmTest3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
modResize.Resize Me
End Sub
