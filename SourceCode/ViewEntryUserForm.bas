Private Sub CloseCommandButton_Click()
    Unload Me
End Sub

Private Sub CopyLoginCommandButton_Click()
    Dim main_ws As Worksheet
    Dim MyDataObj As New DataObject
    
    EnsureNotOnHeaderRow
    
    Set main_ws = ActiveCell.Worksheet
    MyDataObj.SetText main_ws.Cells(ActiveCell.Row, 3).value
    MyDataObj.PutInClipboard
End Sub

Private Sub CopyPasswordCommandButton_Click()
    Dim main_ws As Worksheet
    Dim MyDataObj As New DataObject
    
    EnsureNotOnHeaderRow
    
    Set main_ws = ActiveCell.Worksheet
    MyDataObj.SetText main_ws.Cells(ActiveCell.Row, 4).value
    MyDataObj.PutInClipboard
End Sub

Private Sub CopyPinCommandButton_Click()
    Dim main_ws As Worksheet
    Dim MyDataObj As New DataObject
    
    EnsureNotOnHeaderRow
    
    Set main_ws = ActiveCell.Worksheet
    MyDataObj.SetText main_ws.Cells(ActiveCell.Row, 5).value
    MyDataObj.PutInClipboard
End Sub

Private Sub DeleteCommandButton_Click()
    Dim main_ws As Worksheet
    Dim result As VbMsgBoxResult
    Dim r As Long
    
    EnsureNotOnHeaderRow
    
    r = ActiveCell.Row
    result = MsgBox("Are you sure you want to delete this row?", VbMsgBoxStyle.vbYesNo Or VbMsgBoxStyle.vbQuestion, "Confirm")
    If result = vbYes Then
        Set main_ws = ActiveCell.Worksheet
        main_ws.Rows(r).EntireRow.Delete
        UserForm_Activate
    End If
End Sub

Private Sub EditCommandButton_Click()
    Dim editForm As New EditEntryUserForm
    
    EnsureNotOnHeaderRow
    
    editForm.RowNum = ActiveCell.Row
    editForm.Show False
    UserForm_Activate
End Sub

Private Sub NextCommandButton_Click()
    Dim main_ws As Worksheet
    
    EnsureNotOnHeaderRow
    
    Set main_ws = ActiveCell.Worksheet
    main_ws.Cells(ActiveCell.Row + 1, ActiveCell.Column).Activate
    UserForm_Activate
End Sub

Private Sub OpenUrlCommandButton1_Click()
    Dim main_ws As Worksheet
    
    EnsureNotOnHeaderRow
    
    Set main_ws = ActiveCell.Worksheet
    Dim MyDataObj As New DataObject
    ThisWorkbook.FollowHyperlink main_ws.Cells(ActiveCell.Row, 2).value, , True
End Sub

Private Sub PreviousCommandButton_Click()
    Dim main_ws As Worksheet
    
    EnsureNotOnHeaderRow
    
    If ActiveCell.Row > 1 Then
        Set main_ws = ActiveCell.Worksheet
        main_ws.Cells(ActiveCell.Row - 1, ActiveCell.Column).Activate
        UserForm_Activate
    End If
End Sub

Sub EnsureNotOnHeaderRow()
    Dim main_ws As Worksheet

    If ActiveCell.Row < 2 Then
        Set main_ws = ActiveCell.Worksheet
        main_ws.Cells(2, ActiveCell.Column).Activate
    End If
End Sub

Private Sub UserForm_Activate()
    Dim main_ws As Worksheet
    Dim txt As String
    
    EnsureNotOnHeaderRow
    
    Set main_ws = ActiveCell.Worksheet
    
    Me.NameTextBox.Text = main_ws.Cells(ActiveCell.Row, 1).value
    
    txt = main_ws.Cells(ActiveCell.Row, 2).value
    Me.UrlTextBox.Text = txt
    If Trim(txt) = "" Then
        Me.OpenUrlCommandButton1.Enabled = False
    Else
        Me.OpenUrlCommandButton1.Enabled = True
    End If
    
    txt = main_ws.Cells(ActiveCell.Row, 3).value
    Me.LoginTextBox.Text = txt
    If Trim(txt) = "" Then
        Me.CopyLoginCommandButton.Enabled = False
    Else
        Me.CopyLoginCommandButton.Enabled = True
    End If
    
    txt = main_ws.Cells(ActiveCell.Row, 4).value
    If Trim(txt) = "" Then
        Me.CopyPasswordCommandButton.Enabled = False
    Else
        Me.CopyPasswordCommandButton.Enabled = True
    End If
    
    txt = main_ws.Cells(ActiveCell.Row, 5).value
    If Trim(txt) = "" Then
        Me.CopyPinCommandButton.Enabled = False
    Else
        Me.CopyPinCommandButton.Enabled = True
    End If
    
    Me.NotesTextBox.Text = main_ws.Cells(ActiveCell.Row, 6).value
    Me.RowNumberTextBox.Text = ActiveCell.Row
    If ActiveCell.Row > 2 Then
        Me.PreviousCommandButton.Enabled = True
    Else
        Me.PreviousCommandButton.Enabled = False
    End If
End Sub
