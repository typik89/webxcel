VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InterruptHttpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'HTTP SERVER
Private m_tcpServer As TcpServer
Private m_tcpClient As TcpClient

Private m_lockFile As FileInfo

Private m_workingController As RestController
Private m_controllersDictionary



Private Sub Class_Initialize()
    Set m_controllersDictionary = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    For i = 1 To Worksheets.Count
        Dim sheet
        Set sheet = Worksheets(i)
        
        Dim sheetName As String
        sheetName = LCase(Trim(sheet.name))
        
        Dim table As DynamicWorksheetTable
        Set table = New DynamicWorksheetTable
        table.Initialize sheet
        
        Dim controller As RestController
        Set controller = New RestController
        controller.InitiateProperties table
        
        m_controllersDictionary.Add sheetName, controller
    Next i
    
    Set m_tcpServer = New TcpServer
    
    isWaitingResponse = False
End Sub



Public Sub Init(ByVal port As Long)
    Set m_lockFile = New FileInfo
    m_lockFile.Initialize ActiveWorkbook.FullName & ".lock"
    m_lockFile.Create
    
    m_tcpServer.BindTo port, 100

End Sub
Public Sub server()
    If Not m_tcpClient Is Nothing Then
        If m_workingController.checkResponseReady Then
            m_tcpClient.SendString m_workingController.getResponse.ToString()
            m_tcpClient.Dispose
            Set m_tcpClient = Nothing
            Set m_workingController = Nothing
        End If
    ElseIf isStopped Then
        m_tcpServer.Dispose
        Exit Sub
    Else
        Set m_tcpClient = m_tcpServer.AcceptTcpClient(1000)
        If Not m_tcpClient Is Nothing Then
            Dim requestText As String
            requestText = m_tcpClient.ReceiveString()
            
            If requestText <> "" Then
                On Error GoTo ERR_HANDLER
                
                Dim request As HttpRequest
                Set request = New HttpRequest
                
                request.Parse requestText
                
                Dim keyFromUrl As String
                keyFromUrl = LCase(Replace(request.Url, "/", ""))
                
                If m_controllersDictionary.Exists(keyFromUrl) Then
                    Set m_workingController = m_controllersDictionary(keyFromUrl)
                End If
                
                If m_workingController Is Nothing Or request.RequestMethod <> MethodPost Then
                    m_tcpClient.SendString createErrRespBody(501, request.Url & " is unimplemented for such kind request")
                    Set m_tcpClient = Nothing
                Else
                    m_workingController.ProcessRequest request
                End If
                GoTo NO_ERR_HANDLER
ERR_HANDLER:
    m_tcpClient.SendString createErrRespBody(500, Err.Description)
    Set m_tcpClient = Nothing
    Set m_workingController = Nothing
NO_ERR_HANDLER:
            Else
                Set m_tcpClient = Nothing
            End If
        End If
    End If
End Sub
Public Property Get isStopped()
    isStopped = Not m_lockFile.Exists
    If Not m_lockFile.Exists And Not m_tcpServer Is Nothing Then
        m_tcpServer.Dispose
        Set m_tcpServer = Nothing
    End If
End Property

Public Property Get isWaiting()
    If Not m_workingController Is Nothing Then
        isWaiting = Not m_workingController.checkResponseReady
    Else
        isWaiting = False
    End If
End Property

Private Function createErrRespBody(code As Integer, msg As String) As String
    Dim resp As HttpResponse
    Set resp = New HttpResponse
    
    resp.StatusCode = code
    resp.Body = msg
    
    createErrRespBody = resp.ToString()
End Function




