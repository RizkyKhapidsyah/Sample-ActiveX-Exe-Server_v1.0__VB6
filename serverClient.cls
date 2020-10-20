VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "serverClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This is a callback module which allows communication between
' server and the client. Using this module you can shutdown
' clients , send messages etc.

Public Sub sendMessage(message As String)

End Sub

Public Sub Shutdown(secs As Integer)

End Sub


