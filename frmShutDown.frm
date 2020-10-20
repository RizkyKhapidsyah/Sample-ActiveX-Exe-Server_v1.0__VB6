VERSION 5.00
Begin VB.Form frmShutDown 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   120
   End
   Begin VB.Label Label1 
      Caption         =   "Shut Down in 10 Secs"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmShutDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ticks As Integer

Private Sub Form_Load()
ticks = 0
Me.Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
DoEvents
ticks = ticks + 1
If ticks = 10 Then
    End
Else
    Label1.Caption = "Shutdown in " & 10 - ticks & " Seconds"
End If
End Sub
