VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample Client Form"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetEmp 
      Caption         =   "Get Employee Record"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect to Server"
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdMan 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   12
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdMan 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   11
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdMan 
      Caption         =   "&Modify"
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   10
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdMan 
      Caption         =   "&New"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   ">>"
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   7
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   ">"
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   6
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "<"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "<<"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Server Status : Connected"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H005F5F5F&
      Caption         =   "Record Navigators"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Name"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim client As clsimpl
Dim rstEmployee As ADODB.Recordset

Private Sub cmdConnect_Click()
On Error Resume Next
If Not Me.cmdConnect.Caption = "Disconnect" Then
    Me.Label5.Caption = ""
    Me.Label5.Visible = True

    Me.Label5.Caption = " Server Status : Connecting"
    Set client = New clsimpl
    Set client.ShutDownForm = frmShutDown
    
    ' Note : if the server applcation is located on other machine
    ' use createObject("MyServer.clsServer","MachineName") instead.
    
    Set objServer = CreateObject("MyServer.clsServer")
    
    Sleep 0.5
    Me.Label5.Caption = " Server Status : Server Found"
    
    Sleep 0.5
    objServer.Login "Admin", "admin", client, "comp1"
    Me.Label5.Caption = " Server Status : Authenticating"
    
    Sleep 0.5
    Me.Label5.Caption = " Server Status : Connected"
    Me.cmdConnect.Caption = "Disconnect"
    Me.cmdGetEmp.Enabled = True
Else
    Set objServer = Nothing
    Me.cmdConnect.Caption = "Connect to Server"
    Me.cmdGetEmp.Enabled = False
End If
End Sub

Private Sub Sleep(secs!)
Dim start!
start = Timer

While (Timer < (start + secs))
    DoEvents
Wend
End Sub

Private Sub cmdGetEmp_Click()


Set rstEmployee = objServer.getEmployeeRecord
rstEmployee.MoveFirst

Me.txtName = rstEmployee![EmpName]
Me.txtAddress = rstEmployee![EmpAddress]
End Sub

Private Sub cmdMan_Click(Index As Integer)
Select Case Index
    Case 0 'NEw
        rstEmployee.AddNew
        
        Me.txtName = ""
        Me.txtAddress = ""
        Me.txtName.Locked = False
        Me.txtAddress.Locked = False
        
        Me.cmdMan(0).Enabled = False
        Me.cmdMan(1).Enabled = False
        Me.cmdMan(2).Enabled = True
        Me.cmdMan(2).SetFocus
        Me.cmdMan(3).Enabled = True
        
    Case 1 'Modify
        Me.txtName.Locked = False
        Me.txtAddress.Locked = False
        
        Me.cmdMan(0).Enabled = False
        Me.cmdMan(1).Enabled = False
        Me.cmdMan(2).Enabled = True
        Me.cmdMan(2).SetFocus
        Me.cmdMan(3).Enabled = True
        
    
    Case 2 'Save
        rstEmployee![EmpName] = Me.txtName
        rstEmployee![EmpAddress] = Me.txtAddress
        rstEmployee.Update
        objServer.SubmitData rstEmployee
        GoTo canceled
        
    Case 3 'Cancel
canceled:
        On Error Resume Next
            rstEmployee.CancelUpdate
        
        rstEmployee.MoveFirst
        Me.txtAddress = rstEmployee![EmpAddress]
        Me.txtName = rstEmployee![EmpName]
        On Error GoTo Handler
        Me.txtName.Locked = True
        Me.txtAddress.Locked = True
        
        Me.cmdMan(0).Enabled = True
        Me.cmdMan(0).SetFocus
        Me.cmdMan(1).Enabled = True
        Me.cmdMan(2).Enabled = False
        Me.cmdMan(3).Enabled = False
End Select
Exit Sub

Handler:
    MsgBox Err.Description
    Resume Next
End Sub

Private Sub cmdNav_Click(Index As Integer)
Select Case Index
    Case 0 'First Record
        rstEmployee.MoveFirst
        Me.cmdNav(1).Enabled = False
        Me.cmdNav(2).Enabled = True
        
    Case 1 'Previous
        rstEmployee.MovePrevious
        cmdNav(2).Enabled = True
        
        If rstEmployee.BOF Then
            rstEmployee.MoveFirst
            Me.cmdNav(0).SetFocus
            Me.cmdNav(1).Enabled = False
        End If
    Case 2 'Next
        rstEmployee.MoveNext
        Me.cmdNav(1).Enabled = True
        
        If rstEmployee.EOF Then
            rstEmployee.MoveLast
            
            Me.cmdNav(1).SetFocus
            Me.cmdNav(2).Enabled = False
        End If
    Case 3 'Last
        Me.cmdNav(1).Enabled = True
        Me.cmdNav(2).Enabled = False
        rstEmployee.MoveLast
End Select

Me.txtName = rstEmployee![EmpName]
Me.txtAddress = rstEmployee![EmpAddress]
End Sub
