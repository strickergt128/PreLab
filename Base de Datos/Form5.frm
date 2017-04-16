VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form5"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   LinkTopic       =   "Form5"
   ScaleHeight     =   7860
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "Serie"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "Placa"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "Especificaciones"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\stricker\Desktop\BaseDeDatos\Toyota.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Autos_Semi_Nuevos"
      Top             =   7200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguente"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Menu"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      DataField       =   "Modelo"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      DataField       =   "Años_Usado"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Autos SemiNuevos"
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
      Left            =   2400
      TabIndex        =   16
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Serie"
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
      Left            =   1920
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Placa"
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
      Left            =   1920
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Especificaciones"
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
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Modelo"
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
      Left            =   1920
      TabIndex        =   12
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Años Usado"
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
      Left            =   1560
      TabIndex        =   11
      Top             =   4560
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
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
Form4.Hide
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

