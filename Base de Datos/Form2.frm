VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7665
   LinkTopic       =   "Form2"
   ScaleHeight     =   5985
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Menu"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguente"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2040
      ScaleHeight     =   270
      ScaleWidth      =   3435
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   1935
      Left            =   360
      ScaleHeight     =   1875
      ScaleWidth      =   6915
      TabIndex        =   1
      Top             =   960
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Alquiler"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command3_Click()
Form2.Hide
Form3.Show
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Update
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Delete
Adodc1.Recordset.MovePrevious
End Sub
