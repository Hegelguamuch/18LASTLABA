VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form4"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form4"
   ScaleHeight     =   8190
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   1200
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombres"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Top             =   1800
      Width           =   5535
   End
   Begin VB.TextBox Text3 
      DataField       =   "Apellidos"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   2400
      Width           =   5535
   End
   Begin VB.TextBox Text4 
      DataField       =   "Dni"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   3000
      Width           =   5535
   End
   Begin VB.TextBox Text5 
      DataField       =   "Direccion"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   3600
      Width           =   5535
   End
   Begin VB.TextBox Text6 
      DataField       =   "Telefono"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   4200
      Width           =   5535
   End
   Begin VB.TextBox Text7 
      DataField       =   "Email"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   4800
      Width           =   5535
   End
   Begin VB.TextBox Text8 
      DataField       =   "Password"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   5400
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NUEVO"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ELIMINAR"
      Height          =   495
      Left            =   9360
      TabIndex        =   4
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MODIFICAR"
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GUARDAR"
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   7440
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   9000
      Top             =   4680
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
      Connect         =   $"Form4.frx":0000
      OLEDBString     =   $"Form4.frx":009F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Profesor"
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
      Caption         =   "PROFESOR"
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
      TabIndex        =   23
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "IDPROFESOR"
      Height          =   495
      Left            =   480
      TabIndex        =   22
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "NOMBRES"
      Height          =   495
      Left            =   480
      TabIndex        =   21
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "APELLIDOS"
      Height          =   495
      Left            =   480
      TabIndex        =   20
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "DNI"
      Height          =   495
      Left            =   480
      TabIndex        =   19
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "DIRECCION"
      Height          =   495
      Left            =   480
      TabIndex        =   18
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "TELEFONO"
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "EMAIL"
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "PASSWORD"
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   8880
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      DataField       =   "imagen"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9240
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
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
X = App.Path
Image1.Picture = LoadPicture(X & "/profesor/" & Label12.Caption)
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
End If
X = App.Path
Image1.Picture = LoadPicture(X & "/profesor/" & Label12.Caption)
End Sub

Private Sub Form_Load()
X = App.Path
Image1.Picture = LoadPicture(X & "/profesor/" & Label12.Caption)
End Sub
