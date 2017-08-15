VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "CURSO"
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PROFESOR"
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NOTAS"
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ALUMNOS"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COLEGIO NACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   10770
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   19500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Command4_Click()
Form5.Show
End Sub
