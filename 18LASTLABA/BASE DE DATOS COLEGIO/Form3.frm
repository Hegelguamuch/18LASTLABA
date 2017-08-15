VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FF8080&
   Caption         =   "Form3"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11505
   LinkTopic       =   "Form3"
   ScaleHeight     =   5145
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "IDnota"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   720
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      DataField       =   "IDAlumno"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   1320
      Width           =   5535
   End
   Begin VB.TextBox Text3 
      DataField       =   "IDcurso"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   1920
      Width           =   5535
   End
   Begin VB.TextBox Text4 
      DataField       =   "Unidad"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   2520
      Width           =   5535
   End
   Begin VB.TextBox Text5 
      DataField       =   "promedio"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   3120
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NUEVO"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ELIMINAR"
      Height          =   495
      Left            =   9480
      TabIndex        =   4
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MODIFICAR"
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GUARDAR"
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   9000
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form3.frx":0000
      OLEDBString     =   $"Form3.frx":009F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Notas"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTAS"
      BeginProperty Font 
         Name            =   "@Adobe Fan Heiti Std B"
         Size            =   30
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   16
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "IDNOTA"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "IDALUMNO"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "IDCURSO"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "UNIDAD"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "PROMEDIO"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   3120
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.CancelUpdate
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
End If
End Sub
