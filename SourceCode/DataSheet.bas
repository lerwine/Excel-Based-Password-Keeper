Private lSuspendChangeEvents As Integer
Private lSelectedRowId As String
Private lSelectedRowIndex As Long
Private lSelectedName As String
Private lSelectedUrl As String
Private lSelectedLogin As String
Private lSelectedPassword As String
Private lSelectedPin As String
Private lSelectedNotes As String

Property Get SuspendChangeEvents() As Integer
    SuspendChangeEvents = lSuspendChangeEvents
End Property

Property Let SuspendChangeEvents(value As Integer)
    Dim prevActiveWks As Worksheet
    
    If lSuspendChangeEvents = 0 Then
        If value > 0 Then lSuspendChangeEvents = value
    Else
        If value > 0 Then
            lSuspendChangeEvents = value
        Else
            lSuspendChangeEvents = 0
            Set prevActiveWks = ActiveCell.Worksheet
            If prevActiveWks.name <> DataSheet.name Then DataSheet.Activate
            local_UpdateSelectedProperties
            prevActiveWks.Activate
            If prevActiveWks.name = ManageSheet.name Then ManageSheet.Update_SelectedItem_Cells
        End If
    End If
End Property

Property Get SelectedRowIndex() As Long
    SelectedRowIndex = lSelectedRowIndex
End Property

Property Let SelectedRowIndex(value As Long)
    Dim prevSelected As Long
    Dim prevActiveWks As Worksheet
    
    If SuspendChangeEvents > 0 Then
        lSelectedRowIndex = value
    Else
        prevSelected = lSelectedRowIndex
        lSelectedRowIndex = value

        Set prevActiveWks = ActiveCell.Worksheet
        
        If prevActiveWks.name = DataSheet.name Then
            If value <> prevSelected Then DataSheet.Cells(lSelectedRowIndex, ActiveCell.Column).Activate
            local_UpdateSelectedProperties
        Else
            DataSheet.Activate
            If value <> prevSelected Then DataSheet.Cells(lSelectedRowIndex, ActiveCell.Column).Activate
            local_UpdateSelectedProperties
            prevActiveWks.Activate
            If prevActiveWks.name = ManageSheet.name Then ManageSheet.Update_SelectedItem_Cells
        End If
    End If
End Property

Property Get SelectedName() As String
    SelectedName = lSelectedName
End Property

Property Get SelectedUrl() As String
    SelectedUrl = lSelectedUrl
End Property

Property Get SelectedLogin() As String
    SelectedLogin = lSelectedLogin
End Property

Property Get SelectedNotes() As String
    SelectedNotes = lSelectedNotes
End Property

Property Get SelectedPassword() As String
    SelectedPassword = lSelectedPassword
End Property

Property Get SelectedPin() As String
    SelectedPin = lSelectedPin
End Property

Sub InitializeSelectedProperties()
    Dim prevActiveWks As Worksheet
    Dim prevSelectedRowIndex As Long
    
    SuspendChangeEvents = 1
    
    Set prevActiveWks = ActiveCell.Worksheet
    
    If prevActiveWks.name <> DataSheet.name Then DataSheet.Activate
    
    prevSelectedRowIndex = SelectedRowIndex
    SelectedRowIndex = ActiveCell.Row
    
    If SelectedRowIndex < 2 Then SelectedRowIndex = 2
    
    SuspendChangeEvents = 0
    
    If prevActiveWks.name <> DataSheet.name Then prevActiveWks.Activate
End Sub

Private Sub local_UpdateSelectedProperties()
    If SelectedRowIndex > 1 Then
        lSelectedRowId = ActiveCell.EntireRow.ID
        lSelectedName = DataSheet.Cells(SelectedRowIndex, 1).value
        lSelectedUrl = DataSheet.Cells(SelectedRowIndex, 2).value
        lSelectedLogin = DataSheet.Cells(SelectedRowIndex, 3).value
        lSelectedHasPassword = DataSheet.Cells(SelectedRowIndex, 4).value
        lSelectedPin = DataSheet.Cells(SelectedRowIndex, 5).value
        lSelectedNotes = DataSheet.Cells(SelectedRowIndex, 6).value
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim r As Range
    Dim newRowNum As Long
    Dim startRowNum As Long
    Dim endRowNum As Long
    
    If lSelectedRowId = ActiveCell.EntireRow.ID Then
        If Me.SuspendChangeEvents = 0 Then local_UpdateSelectedProperties
    Else
        newRowNum = 0
        endRowNum = Target.Row + Target.Rows.Count
        If endRowNum < SelectedRowIndex Then endRowNum = SelectedRowIndex
        
        For i = 2 To Target.Row + endRowNum
            Set r = DataSheet.Rows(i, 1)
            If r.EntireRow.ID = lSelectedRowId Then
                newRowNum = i
                Exit For
            End If
        Next
        If newRowNum > 0 Then
            SelectedRowIndex = newRowNum
        Else
            startRowNum = endRowNum
            endRowNum = endRowNum + Target.Rows.Count + 1
            For i = startRowNum To endRowNum
                Set r = DataSheet.Rows(i, 1)
                If r.EntireRow.ID = lSelectedRowId Then
                    n = i
                    Exit For
                End If
            Next
            If newRowNum > 0 Then
                SelectedRowIndex = newRowNum
            Else
                SelectedRowIndex = Target.Row
            End If
        End If
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    SelectedRowIndex = ActiveCell.Row
End Sub

Sub GoToPreviousRecord()
    If SelectedRowIndex > 2 Then SelectedRowIndex = SelectedRowIndex - 1
End Sub

Sub GoToNextRecord()
    SelectedRowIndex = SelectedRowIndex + 1
End Sub

Sub OpenSelectedUrl()
    If Trim(SelectedUrl) = "" Then
        MsgBox "No URL defined."
    Else
        ThisWorkbook.FollowHyperlink SelectedUrl, , True
    End If
End Sub

Sub CopySelectedPassword()
    Dim MyDataObj As New DataObject
    
    If SelectedPassword = "" Then
        MsgBox "No password to copy."
    Else
       MyDataObj.SetText SelectedPassword
        MyDataObj.PutInClipboard
    End If
End Sub

Sub CopySelectedPin()
    Dim MyDataObj As New DataObject
    
    If SelectedPin = "" Then
        MsgBox "No pin to copy."
    Else
        MyDataObj.SetText SelectedPin
        MyDataObj.PutInClipboard
    End If
End Sub

Sub EditSelectedRowIndexData()
    Dim editForm As New EditEntryUserForm
    If SelectedRowIndex < 2 Then SelectedRowIndex = 2
    editForm.InitializeControls
    editForm.Show False
End Sub

Private Function local_UpdateCellValue(rowNum As Long, colNum As Integer, value As String) As Boolean
    Dim prevRow As Long
    Dim prevCol As Integer
    
    prevRow = SelectedRowIndex
    prevCol = ActiveCell.Column
    If prevRow <> rowNum Or prevCol <> colNum Then DataSheet.Cells(rowNum, colNum).Activate
    DataSheet.Cells(rowNum, colNum).value = value
    If prevRow <> rowNum Or prevCol <> colNum Then DataSheet.Cells(prevRow, prevCol).Activate
    If prevRow = rowNum Then local_UpdateCellValue = True Else local_UpdateCellValue = False
End Function

Function UpdateCellValue(rowNum As Long, colNum As Integer, value As String) As Boolean
    Dim prevActiveWks As Worksheet
    
    If rowNum > 1 Then
        Set prevActiveWks = ActiveCell.Worksheet
        
        If prevActiveWks.name = DataSheet.name Then
            UpdateCellValue = local_UpdateCellValue(rowNum, colNum, value)
        Else
            DataSheet.Activate
            UpdateCellValue = local_UpdateCellValue(rowNum, colNum, value)
            prevActiveWks.Activate
        End If
    Else
        UpdateCellValue = False
    End If
End Function

Private Function local_GetValueAt(rowNum As Long, colNum As Integer) As String
    Dim prevRow As Long
    Dim prevCol As Integer
    
    prevRow = SelectedRowIndex
    prevCol = ActiveCell.Column
    If prevRow <> rowNum Or prevCol <> colNum Then DataSheet.Cells(rowNum, colNum).Activate
    local_GetValueAt = DataSheet.Cells(rowNum, colNum).value
    If prevRow <> rowNum Or prevCol <> colNum Then DataSheet.Cells(prevRow, prevCol).Activate
End Function

Function GetValueAt(rowNum As Long, colNum As Integer) As String
    Dim prevActiveWks As Worksheet
    
    If rowNum > 1 Then
        Set prevActiveWks = ActiveCell.Worksheet
        
        If prevActiveWks.name = DataSheet.name Then
            GetValueAt = local_GetValueAt(rowNum, colNum)
        Else
            DataSheet.Activate
            GetValueAt = local_GetValueAt(rowNum, colNum)
            prevActiveWks.Activate
        End If
    Else
        GetValueAt = ""
    End If
End Function

Sub UpdateName(rowNum As Long, value As String)
    If UpdateCellValue(rowNum, 1, value) Then lSelectedName = value
End Sub

Function GetNameAt(rowNum As Long) As String
    GetNameAt = GetValueAt(rowNum, 1)
End Function

Sub UpdateUrl(rowNum As Long, value As String)
    If UpdateCellValue(rowNum, 2, value) Then lSelectedUrl = value
End Sub

Function GetUrlAt(rowNum As Long) As String
    GetUrlAt = GetValueAt(rowNum, 2)
End Function

Sub UpdateLogin(rowNum As Long, value As String)
    If UpdateCellValue(rowNum, 3, value) Then lSelectedLogin = value
End Sub

Function GetLoginAt(rowNum As Long) As String
    GetLoginAt = GetValueAt(rowNum, 3)
End Function

Sub UpdatePassword(rowNum As Long, value As String)
    If UpdateCellValue(rowNum, 4, value) Then lSelectedPassword = value
End Sub

Function GetPasswordAt(rowNum As Long) As String
    GetPasswordAt = GetValueAt(rowNum, 4)
End Function

Sub UpdatePin(rowNum As Long, value As String)
    If UpdateCellValue(rowNum, 5, value) Then lSelectedPin = value
End Sub

Function GetPinAt(rowNum As Long) As String
    GetPinAt = GetValueAt(rowNum, 5)
End Function

Sub UpdateNotes(rowNum As Long, value As String)
    If UpdateCellValue(rowNum, 6, value) Then lSelectedNotes = value
End Sub

Function GetNotesAt(rowNum As Long) As String
    GetNotesAt = GetValueAt(rowNum, 6)
End Function
