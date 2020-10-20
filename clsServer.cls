VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public User As UserInfo
Private isValidUser As Boolean

Private Sub Class_Initialize()
AddActivity "Server Object Created"
Set User = New UserInfo
End Sub

Private Sub Class_Terminate()
    AddActivity "Logout Request Received For User : " & User.UserName
    User.LogoutTime = Time
    User.status = False
    conServer.Execute "Update Users Set Status=0 WHERE UserName='" & User.UserName & "'"
    UpdateUserList User.UserName
    AddActivity "Disconnected User :  " & User.UserName
End Sub

Public Function Login(UserName As String, Password As String, clbkClient As Object, Optional Workstation As String, Optional ErrMsg As String) As Boolean
On Error GoTo Handler
Dim rstTmp As Recordset

AddActivity "Authenticating User : " & UserName
Set rstTmp = New Recordset
With rstTmp
    .ActiveConnection = conServer
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    'MsgBox "SELECT UserName,Password From Users WHERE UserName='" & UserName & "' And Status=0"
    .Open "SELECT UserName,Password From Users WHERE UserName='" & UserName & "' And Status=0"
End With

'rstTmp.MoveFirst
If getRowsOk(rstTmp) Then
    rstTmp.MoveFirst
    If StrComp(Password, rstTmp![Password], vbBinaryCompare) = 0 Then
        isValidUser = True
        AddActivity "Client authentication succesfull for user : " & UserName
        
        Set User = New UserInfo
        With User
            .UserName = UserName
            .LoginTime = Time
            .Workstation = Workstation
            .status = True
            Set .client = clbkClient
        End With
        conServer.Execute "Update Users Set Status=1 WHERE UserName='" & User.UserName & "'"
        AddActivity "Adding to User List user : " & UserName
        Users.Add User, User.UserName
        UpdateUserList
        
        Login = True
    Else
        AddActivity "Client authentication Failed for user : " & UserName
        ErrMsg = "Incorrect Password"
        Login = False
    End If
Else
    Login = False
    ErrMsg = "The Specified User Does not Exists" & vbCrLf & "Or Has already Logged In"
End If
Exit Function

Handler:
    MsgBox Err.Description
    Exit Function
End Function

Private Sub UpdateUserList(Optional UserName As String = "")
On Error GoTo Handler
Dim yog As Integer
Dim clObj As UserInfo

If Not UserName = "" Then Users.Remove UserName
frmServer.lvUsers.ListItems.Clear
    For yog = 1 To Users.Count
        Set clObj = Users.Item(yog)
        With frmServer.lvUsers
            .ListItems.Add , "K" & yog
                .ListItems(yog).SubItems(1) = clObj.UserName
                .ListItems(yog).SubItems(2) = clObj.Workstation
                .ListItems(yog).SubItems(3) = clObj.LoginTime
                .ListItems(yog).SubItems(4) = clObj.LogoutTime
                .ListItems(yog).SubItems(5) = IIf(clObj.status, "Logged In", "Logged Off")
        End With
    Next
Exit Sub

Handler:
    'MsgBox Err.Description
    Resume Next
End Sub

Public Function getEmployeeRecord() As Recordset
Dim rstEmployee As Recordset
AddActivity "Request for Employee Data received"
Set rstEmployee = New Recordset
With rstEmployee
    .ActiveConnection = conServer
    .CursorType = adOpenDynamic
    .CursorLocation = adUseClient
    .LockType = adLockBatchOptimistic
    .Open "SELECT * FROM Employee"
End With
Set rstEmployee.ActiveConnection = Nothing
Set getEmployeeRecord = rstEmployee
AddActivity "Employee data Sent to user : " & User.UserName
End Function

Public Sub SubmitData(rst As adodb.Recordset)
On Error Resume Next
AddActivity "Data Update Request received From User :" & User.UserName
Set rst.ActiveConnection = conServer
rst.UpdateBatch
AddActivity "Data Updated From User :" & User.UserName
Set rst.ActiveConnection = Nothing
End Sub
