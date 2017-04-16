VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form4"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8190
   LinkTopic       =   "Form4"
   ScaleHeight     =   8160
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "codigo_cliente"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre_Cliente"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "Edad"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\stricker\Desktop\BaseDeDatos\Toyota.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   7560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguente"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Menu"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      DataField       =   "Fecha_Nacimiento"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      DataField       =   "Genero"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Edad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha de Nacimiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Genero"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF Then
Data1.Recordset.MoveLast
End If
End Sub

Private Sub Command2_Click()
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
Data1.Recordset.MoveFirst
End If
End Sub

Private Sub Command3_Click()
Form2.Hide
Form3.Show
End Sub

Private Sub Command4_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command5_Click()
Data1.Recordset.Update
Data1.Recordset.MoveNext
End Sub

Private Sub Command6_Click()
Data1.Recordset.Delete
Data1.Recordset.MovePrevious
End Sub

