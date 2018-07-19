VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicWorksheetTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type ColumnHeader
    name As String
    index As Integer
End Type


Private m_sheet
Private m_columns As Collection

Public Property Get Columns() As Collection
    Set Columns = New Collection
    
    For Each column In m_columns
        Columns.Add column.name
    Next column
End Property
Public Property Get entries() As Collection
    Set entries = New Collection
    Dim i As Long
    For i = 2 To m_sheet.Rows.Count
        Dim row
        Set row = m_sheet.Rows(i)
        
        If Not isRowFilled(row) Then
            Exit For
        End If
        
        Dim entry
        Set entry = CreateObject("Scripting.Dictionary")
                
        For Each column In m_columns
            Dim val
            val = Trim(row.Cells(column.index).text)
            
            If val = "" Then
                entry(column.name) = Null
            Else
                entry(column.name) = val
            End If
        Next column
                
        entries.Add entry
    Next i
End Property

Public Sub SetEntries(entries)
    Dim index As Long
    index = 2
    For Each entry In entries
        For Each column In m_columns
            If column.formula = "" Then
                Cells(index, column.index).value = entry(column.name)
            Else
                Cells(index, column.index).formula = column.formula
            End If
        Next column
        index = index + 1
    Next entry
End Sub

Public Sub Activate()
    m_sheet.Activate
End Sub

Public Sub ClearAllValues()
    Dim i As Long
    For i = 2 To m_sheet.Rows.Count
        Dim row
        Set row = m_sheet.Rows(i)
        
        If isRowFilled(row) Then
            For Each column In m_columns
                If column.formula = "" Then
                    row.Cells(column.index).value = ""
                ElseIf i > 2 Then
                    row.Cells(column.index).value = ""
                    row.Cells(column.index).formula = ""
                End If
            Next column
        Else
            Exit Sub
        End If
    Next i
End Sub

Private Function isRowFilled(row) As Boolean
    isRowFilled = False
    For Each column In m_columns
        If column.formula = "" Then
            Dim val
            val = Trim(row.Cells(column.index).text)
            If val <> "" Then
                isRowFilled = True
                Exit For
            End If
        End If
    Next column
End Function

Public Sub Initialize(Worksheet)
    Set m_sheet = Worksheet
    Set m_columns = New Collection
    
    Dim firstRow
    Set firstRow = m_sheet.Rows(1)
    
    Dim secondRow
    Set secondRow = m_sheet.Rows(2)
    
    
    Dim i As Integer
    For i = 1 To firstRow.Cells.Count
        Dim column
        Set column = firstRow.Cells(i)
        Dim text As String
        text = Trim(column.text)
        If text = "" Then
            Exit For
        End If
        Dim header As WorksheetTableColumn
        Set header = New WorksheetTableColumn
        header.name = text
        header.index = i
        If secondRow.Cells(i).HasFormula Then
            header.formula = secondRow.Cells(i).formula
        End If
        m_columns.Add header
    Next i
End Sub

