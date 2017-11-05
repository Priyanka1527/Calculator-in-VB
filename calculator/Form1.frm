VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "backspace"
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   615
      Index           =   9
      Left            =   480
      TabIndex        =   10
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   615
      Index           =   8
      Left            =   2400
      TabIndex        =   9
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   615
      Index           =   7
      Left            =   1440
      TabIndex        =   8
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   615
      Index           =   6
      Left            =   480
      TabIndex        =   7
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   615
      Index           =   5
      Left            =   2400
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   615
      Index           =   4
      Left            =   1440
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   615
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   615
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   615
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Private Sub Command1_Click(Index As Integer)
Text1.Text = Text1.Text + Command1(Index).Caption
End Sub

Private Sub Command2_Click()
Dim length As Integer
length = Len(Text1.Text)
If (length <> 0) Then
Text1.Text = Left(Text1.Text, length - 1)
End If
End Sub
