VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Server Monitor"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBroadCast 
      Caption         =   "Broadcast Message"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdSendMsg 
      Caption         =   "Send a Message to Selected"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Disconnect Selected User"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtActLog 
      Appearance      =   0  'Flat
      Height          =   1815
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmServer.frx":0000
      Top             =   3600
      Width           =   6735
   End
   Begin MSComctlLib.ListView lvUsers 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   420
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Machine Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Login Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Logout Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H005F5F5F&
      Caption         =   " User Activity Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   6735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H005F5F5F&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Currently Logged In Users"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBroadCast_Click()
Dim yog As Integer

Dim msg As String
msg = InputBox("Enter Broadcast Message")
For yog = 1 To Users.Count
    Users(yog).client.sendMessage msg
Next
End Sub

Private Sub cmdSendMsg_Click()
On Error Resume Next
Dim msg As String
Dim obj As UserInfo

Set obj = Users.Item(CInt(Mid(Me.lvUsers.SelectedItem.Key, 2)))
msg = InputBox("Enter Message")
obj.client.sendMessage msg
End Sub

Private Sub Command1_Click()
'On Error Resume Next
Dim msg As String
Dim obj As UserInfo

Set obj = Users.Item(CInt(Mid(Me.lvUsers.SelectedItem.Key, 2)))
obj.client.Shutdown (0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
For yog = 1 To Users.Count
    Users(yog).client.Shutdown 0
Next

End Sub
