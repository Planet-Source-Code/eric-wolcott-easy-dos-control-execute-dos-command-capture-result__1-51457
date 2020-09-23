VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4230
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   720
      Width           =   7260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Execute"
      Height          =   315
      Left            =   6090
      TabIndex        =   2
      Top             =   360
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   660
      TabIndex        =   1
      Text            =   "nbtstat -a"
      Top             =   375
      Width           =   5385
   End
   Begin Project1.DOS DOS1 
      Height          =   600
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   1058
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Text2.Text = Text2.Text & DOS1.ExecuteCommand(Text1.Text)
End Sub
