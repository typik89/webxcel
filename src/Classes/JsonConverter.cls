VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function convertToJsonObject(dict) As JsonObject
    Set convertToJsonObject = New JsonObject
    Dim key As Variant
    For Each key In dict.keys
        Dim JsonValue As JsonValue
        Set JsonValue = New JsonValue
        JsonValue.value = dict(key)
        convertToJsonObject.SetProperty key, JsonValue
    Next key
End Function

Public Function convertToDict(JsonObject As JsonObject)
    Set convertToDict = CreateObject("Scripting.Dictionary")
    Dim property As Variant
    For Each property In JsonObject.PropertyNames
        Dim JsonValue As JsonValue
        Set JsonValue = JsonObject.GetProperty(property)
        convertToDict.Add property, JsonValue.value
    Next property
End Function

