Attribute VB_Name = "Main"
Public server As InterruptHttpServer

Public Sub Main()
    Set server = New InterruptHttpServer
    
    On Error GoTo ERR_HNDL_BINDING
    server.Init 1018
    GoTo ERR_NO
ERR_HNDL_BINDING:
    MsgBox "Error in initializing server: " + Err.Description
    Exit Sub
ERR_NO:
    periodically
End Sub

Private Sub periodically()
    server.server
    If server.isStopped Then
        MsgBox "Server was stopped"
    ElseIf server.isWaiting Then
        Application.OnTime (Now + TimeValue("00:00:02")), "periodically"
    Else
        Application.OnTime (Now + TimeValue("00:00:00")), "periodically"
    End If
End Sub

