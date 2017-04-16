VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   7575
   ClientTop       =   2670
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   7740
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5520
      Top             =   4440
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5280
      Top             =   5520
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4080
      Top             =   5400
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4560
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3840
      Top             =   4320
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
x = Shape1.Top
x = x + 50
Shape1.Top = x
x = Shape1.Left
x = x + 50
Shape1.Left = x
If Shape1.Top > 4800 Then
Shape1.Top = 4800
Timer2.Enabled = True
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
x = Shape1.Top
x = x - 50
Shape1.Top = x
If Shape1.Top < 0 Then
Shape1.Top = 0
x = Shape1.Left
x = x - 50
Shape1.Left = x
End If
If Shape1.Left < 0 Then
Shape1.Left = 0
Timer2.Enabled = False
Timer3.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
x = Shape1.Top
x = x + 50
Shape1.Top = x
If Shape1.Top > 5160 Then
Shape1.Top = 5160
Timer3.Enabled = False
Timer4.Enabled = True
End If
End Sub

Private Sub Timer4_Timer()
x = Shape1.Top
x = x - 50
Shape1.Top = x
x = Shape1.Left
x = x + 50
Shape1.Left = x
If Shape1.Left > 5160 Then
Shape1.Left = 5160
Timer5.Enabled = True
Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
x = Shape1.Left
x = x - 50
Shape1.Left = x
If Shape1.Left < 0 Then
Shape1.Left = 0
Timer5.Enabled = False
Timer1.Enabled = True
End If
End Sub
