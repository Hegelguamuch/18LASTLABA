VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form5"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form5"
   ScaleHeight     =   6705
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GUARDAR"
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MODIFICAR"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ELIMINAR"
      Height          =   495
      Left            =   9120
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NUEVO"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "IDprofesor"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3000
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre del curso"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      DataField       =   "IDcurso"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   5535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   8880
      Top             =   2040
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
      Connect         =   $"Form5.frx":0000
      OLEDBString     =   $"Form5.frx":009F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Curso"
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
   Begin VB.Label Label2 
      Caption         =   "IDCURSO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "IDPROFESOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "NOMBRE DEL CURSO"
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
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CURSO"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "Form5"
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
