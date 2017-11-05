VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000011&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   5415
   ClientLeft      =   5805
   ClientTop       =   3060
   ClientWidth     =   8430
   Icon            =   "Calculator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Calculator.frx":0442
   ScaleHeight     =   5415
   ScaleWidth      =   8430
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
      Left            =   6960
      TabIndex        =   38
      Top             =   1080
      Width           =   975
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
      Left            =   5760
      TabIndex        =   37
      Top             =   1080
      Width           =   975
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
      Left            =   4560
      TabIndex        =   36
      Top             =   1080
      Width           =   975
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
      Left            =   3360
      TabIndex        =   35
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Pi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   34
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command25 
      Caption         =   "CosInv"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   33
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command24 
      Caption         =   "SinInv"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   32
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command23 
      Caption         =   "TanInv"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   31
      Top             =   1800
      Width           =   735
   End
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
      Height          =   615
      Left            =   7200
      MaskColor       =   &H00808080&
      TabIndex        =   30
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Int"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   29
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Exp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   28
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   27
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   26
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   25
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   24
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "n!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   23
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "x^y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   22
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "+/-"
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
      Left            =   4320
      TabIndex        =   21
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   20
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   19
      Top             =   1800
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
      Left            =   480
      TabIndex        =   18
      Top             =   4320
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
      Left            =   2400
      TabIndex        =   17
      Top             =   4320
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
      Left            =   3360
      TabIndex        =   16
      Top             =   4320
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
      Left            =   3360
      TabIndex        =   15
      Top             =   3480
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
      Left            =   3360
      TabIndex        =   14
      Top             =   2640
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
      Left            =   3360
      TabIndex        =   13
      Top             =   1800
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
      Left            =   2040
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
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
      Left            =   480
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
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
      Left            =   1440
      TabIndex        =   10
      Top             =   4320
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
      Left            =   2400
      TabIndex        =   9
      Top             =   3480
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
      Left            =   1440
      TabIndex        =   8
      Top             =   3480
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
      Left            =   480
      TabIndex        =   7
      Top             =   3480
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
      Left            =   2400
      TabIndex        =   6
      Top             =   2640
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
      Left            =   1440
      TabIndex        =   5
      Top             =   2640
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
      Left            =   480
      TabIndex        =   4
      Top             =   2640
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
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
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
      Left            =   480
      TabIndex        =   1
      Top             =   1800
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
      Width           =   7575
   End
   Begin VB.Menu ab 
      Caption         =   "&About"
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu stan 
         Caption         =   "Standard"
      End
      Begin VB.Menu sci 
         Caption         =   "Scientific"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
   End
   Begin VB.Menu part 
      Caption         =   "&Partners"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op1, op2 As Double
Dim op As String
Dim result As Double
Dim save, mem As Double

Private Sub ab_Click()
Form3.Show
End Sub

Private Sub Command1_Click(Index As Integer)
Text1.Text = Text1.Text + Command1(Index).Caption
End Sub

Private Sub Command10_Click()
op1 = Val(Text1.Text)
Text1.Text = op1 ^ 0.5
End Sub

Private Sub Command11_Click()
op1 = Val(Text1.Text)
Text1.Text = 1 / op1
End Sub

Private Sub Command12_Click()
op1 = Val(Text1.Text)
Text1.Text = (-1) * op1
End Sub

Private Sub Command13_Click()
op1 = Val(Text1.Text)
Text1.Text = " "
Text1.SetFocus
op = "x^y"
End Sub

Private Sub Command14_Click()
Dim f, i As Integer
f = 1
op1 = Val(Text1.Text)
For i = 1 To op1
f = f * i
Next i
Text1.Text = f
End Sub

Private Sub Command15_Click()
op1 = Val(Text1.Text)
op1 = op1 * ((22 / 7) / 180)
Text1.Text = Sin(op1)
End Sub

Private Sub Command16_Click()
op1 = Val(Text1.Text)
op1 = op1 * ((22 / 7) / 180)
Text1.Text = Cos(op1)
End Sub

Private Sub Command17_Click()
op1 = Val(Text1.Text)
op1 = op1 * ((22 / 7) / 180)
Text1.Text = Tan(op1)
End Sub

Private Sub Command18_Click()
op1 = Val(Text1.Text)
If (op1 > 0) Then
Text1.Text = Log(op1)
Else
Text1.Text = "invalid"
End If
End Sub

Private Sub Command19_Click()
op1 = Val(Text1.Text)
Text1.Text = Exp(op1)
End Sub

Private Sub Command2_Click()
Dim length As Integer
length = Len(Text1.Text)
If (length <> 0) Then
Text1.Text = Left(Text1.Text, length - 1)
End If
End Sub

Private Sub Command20_Click()
Dim v As Integer
op1 = Val(Text1.Text)
v = op1
Text1.Text = v
End Sub

Private Sub Command21_Click()
MsgBox "Made by Priyanka Saha"
End Sub

Private Sub Command22_Click()
End
End Sub

Private Sub Command23_Click()
Dim v As Double
op1 = Val(Text1.Text)
v = Atn(op1)
Text1.Text = v * (180 / (22 / 7))
End Sub

Private Sub Command24_Click()
Dim v, t As Double
op1 = Val(Text1.Text)
v = op1 / (Sqr(1 - (op1 * op1)))
t = Atn(v)
Text1.Text = t * (180 / (22 / 7))
End Sub

Private Sub Command25_Click()
Dim v, t As Double
op1 = Val(Text1.Text)
v = op1 / (Sqr(1 - (op1 * op1)))
t = ((22 / 7) / 2) - Atn(v)
Text1.Text = t * (180 / (22 / 7))
End Sub

Private Sub Command26_Click()
Text1.Text = 3.1415926
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
ElseIf (op = "x^y") Then
result = op1 ^ op2
End If
Text1.Text = result
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text + "."
End Sub


Private Sub help_Click()
Form2.Show
End Sub

Private Sub part_Click()
Form4.Show
End Sub

Private Sub sci_Click()
Unload Me
Form1.Show
End Sub

Private Sub stan_Click()
Unload Me
Form5.Show
End Sub
