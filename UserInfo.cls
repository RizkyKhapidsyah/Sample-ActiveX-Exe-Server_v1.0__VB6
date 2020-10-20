VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_UserName As String
Private m_LoginTime As String
Private m_LogoutTime As String
Private m_workStation As String
Private m_Status As Boolean
Public client As Object

Public Property Get Workstation() As String
    Workstation = m_workStation
End Property

Friend Property Let Workstation(newWork As String)
    m_workStation = newWork
End Property
Public Property Get UserName() As String
    UserName = m_UserName
End Property

Friend Property Let UserName(newName As String)
    m_UserName = newName
End Property

Public Property Get LoginTime() As String
    LoginTime = m_LoginTime
End Property

Friend Property Let LoginTime(newTime As String)
    m_LoginTime = newTime
End Property

Public Property Get LogoutTime() As String
    LogoutTime = m_LogoutTime
End Property

Public Property Let LogoutTime(newTime As String)
    m_LogoutTime = newTime
End Property

Public Property Get status() As Boolean
    status = m_Status
End Property

Public Property Let status(newState As Boolean)
    m_Status = newState
End Property
