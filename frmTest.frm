VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmTest 
   Caption         =   "Test"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Show Form 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   855
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Width           =   1620
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"frmTest.frx":0000
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   480
      Y1              =   0
      Y2              =   840
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   3720
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
frmTest2.Show
End Sub

'for every form you want to resize put this in the resize
'event procedure.  Remember to mention me if you use this ;)
'resizes fonts too!!
Private Sub Form_Resize()
    modResize.Resize Me
End Sub

Private Sub Form_Load()
    Image1.Picture = LoadPicture(App.Path & "\shaung.jpg")
End Sub

