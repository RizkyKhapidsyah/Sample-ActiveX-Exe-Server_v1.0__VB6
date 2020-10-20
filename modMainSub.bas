Attribute VB_Name = "modMainSub"
Public Users As Collection
Public conServer As Connection

Sub main()
On Error GoTo Handler
    If App.PrevInstance = True Then
    End
Else
    If App.StartMode = vbSModeStandalone Then
      Set Users = New Collection
    AddActivity "Attempting to Connect to Database"
      Set conServer = New Connection
      With conServer
        .Provider = "Microsoft.Jet.OLEDB.3.51"
        .ConnectionString = "Data Source=" & App.Path & "\" & "mydb.mdb"
        .Open
      End With
      AddActivity "Connection to Database opened"
      frmServer.Show
    Else
        End
    End If
End If
Exit Sub

Handler:
    MsgBox "Error Occured While Opening Server connection", vbInformation
    End
End Sub

Public Function getRowsOk(ByVal rst As Recordset) As Boolean
On Error GoTo Handler
    Dim vardata
       
    vardata = rst.GetRows(1)
    rst.MoveFirst
    getRowsOk = True
Exit Function

Handler:
    getRowsOk = False
    Exit Function
End Function

Public Sub AddActivity(msg As String)
    frmServer.txtActLog.Text = frmServer.txtActLog.Text & vbCrLf & msg
    frmServer.txtActLog.SelStart = Len(frmServer.txtActLog.Text) - 1
    frmServer.txtActLog.SelLength = 1
End Sub
