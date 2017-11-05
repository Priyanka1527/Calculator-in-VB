VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Calculator"
   ClientHeight    =   5640
   ClientLeft      =   8130
   ClientTop       =   3075
   ClientWidth     =   4320
   Icon            =   "Calculator(Simple).frx":0000
   LinkTopic       =   "Form5"
   Picture         =   "Calculator(Simple).frx":030A
   ScaleHeight     =   5640
   ScaleWidth      =   4320
   Begin VB.CommandButton Command22 
      BackColor       =   &H8000000B&
      Caption         =   "OFF"
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
      Left            =   3000
      MaskColor       =   &H00808080&
      TabIndex        =   23
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command30 
      Caption         =   "MS"
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
      Left            =   3240
      TabIndex        =   22
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command29 
      Caption         =   "MR"
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
      Left            =   2280
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command28 
      Caption         =   "M+"
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
      Left            =   1320
      TabIndex        =   20
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command27 
      Caption         =   "M-"
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
      Left            =   360
      TabIndex        =   19
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   18
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   17
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   16
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   15
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   14
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "backspace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   1320
      TabIndex        =   10
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   2280
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   1320
      TabIndex        =   8
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   2280
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu abt 
      Caption         =   "&About"
   End
   Begin VB.Menu vi 
      Caption         =   "&View"
      Begin VB.Menu std 
         Caption         =   "Standard"
      End
      Begin VB.Menu scf 
         Caption         =   "Scientific"
      End
   End
   Begin VB.Menu hel 
      Caption         =   "&Help"
   End
   Begin VB.Menu pt 
      Caption         =   "&Partners"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op1, op2 As Double
Dim op As String
Dim result As Double
Dim save, mem As Double
Private Sub abt_Click()
Form3.Show
End Sub

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

Private Sub Command22_Click()
End
End Sub

Private Sub Command27_Click()
If (save = -909999) Then
MsgBox "NOTHING SAVED IN MEMORY"
Else
op = "M-"
op1 = Val(Text1.Text)
save = op1 - save
Text1.Text = save
End If
End Sub

Private Sub Command28_Click()
If (save = -909999) Then
MsgBox "NOTHING SAVED IN MEMORY"
Else
op = "M+"
op1 = Val(Text1.Text)
save = save + op1
Text1.Text = save
End If
End Sub

Private Sub Command29_Click()
op = "MR"
Text1.Text = ""
Text1.Text = mem
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Command30_Click()
op = "MS"
save = result
mem = result
Text1.Text = ""
MsgBox "NUMBER SAVED"
End Sub

Private Sub Command4_Click()
op1 = Val(Text1.Text)
Text1.Text = " "
Text1.SetFocus
op = "+"
End Sub

Private Sub Command5_Click()
op1 = Val(Text1.Text)
Text1.Text = " "
Text1.SetFocus
op = "-"
End Sub

Private Sub Command6_Click()
op1 = Val(Text1.Text)
Text1.Text = " "
Text1.SetFocus
op = "*"
End Sub

Private Sub Command7_Click()
op1 = Val(Text1.Text)
Text1.Text = " "
Text1.SetFocus
op = "/"
End Sub

Private Sub Command8_Click()
op2 = Val(Text1.Text)
If (op = "+") Then
result = op1 + op2
ElseIf (op = "-") Then
result = op1 - op2
ElseIf (op = "*") Then
result = op1 * op2
ElseIf (op = "/") Then
result = op1 / op2
End If
Text1.Text = result
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text + "."
End Sub

Private Sub hel_Click()
Form2.Show
End Sub

Private Sub pt_Click()
Form4.Show
End Sub

Private Sub scf_Click()
Unload Me
Form1.Show
End Sub

Private Sub std_Click()
Unload Me
Form5.Show
End Sub

