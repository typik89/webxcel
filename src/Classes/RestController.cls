VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RestController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_jsonConverter As JsonConverter
Private m_jsonParser As JsonParser
Private m_table As DynamicWorksheetTable

Public Sub InitiateProperties(table As DynamicWorksheetTable)
    Set m_table = table
End Sub


Private Sub Class_Initialize()
    Set m_jsonParser = New JsonParser
    Set m_jsonConverter = New JsonConverter
End Sub


Public Function MatchesUrl(requestUrl As String) As Boolean
    IWebController_MatchesUrl = False
End Function
Public Function getResponse()
    Dim jsonArrayResult As JsonArray
    Set jsonArrayResult = New JsonArray

    For Each entry In m_table.entries
        jsonArrayResult.AddItem m_jsonConverter.convertToJsonObject(entry)
    Next entry
    
    Set getResponse = New HttpResponse
    getResponse.Headers.AddHeader "Content-Type", "application/json"
    getResponse.Headers.AddHeader "Access-Control-Allow-Origin", "*"
    getResponse.StatusCode = 201
    
    Dim val As IJson
    Set val = jsonArrayResult
    getResponse.Body = val.ToJson()
End Function
Public Sub ProcessRequest(request As HttpRequest)
    Dim postBody As JsonArray
    Set postBody = m_jsonParser.ParseArray(request.Body)
        
    m_table.Activate
    m_table.ClearAllValues
    
    Dim entries As Collection
    Set entries = New Collection
    
    Dim i As Integer
    For i = 0 To postBody.Count - 1
        entries.Add m_jsonConverter.convertToDict(postBody.GetItem(i))
    Next i
    
    m_table.SetEntries entries
End Sub
Public Function checkResponseReady() As Boolean
    For Each entry In m_table.entries
        Dim key As Variant
        For Each key In entry.keys
            If entry(key) = "#N/A Requesting Data..." Then
                checkResponseReady = False
                GoTo end_checkresponse
            End If
        Next key
    Next entry
    checkResponseReady = True
end_checkresponse:
End Function






