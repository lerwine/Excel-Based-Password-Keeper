Private lRowNum As Long

Property Get RowNum() As Long
    RowNum = lRowNum
End Property

Property Let RowNum(value As Long)
    Dim main_ws As Worksheet
    
    If value > 1 Then
        If lRowNum > 0 Then
            MsgBox "Another row is already being edited."
        Else
            Set main_ws = ActiveCell.Worksheet
            lRowNum = value
            Me.NameTextBox.Text = main_ws.Cells(value, 1).value
            Me.UrlTextBox.Text = main_ws.Cells(value, 2).value
            Me.LoginTextBox.Text = main_ws.Cells(value, 3).value
            Me.PasswordTextBox.Text = main_ws.Cells(value, 4).value
            Me.PinTextBox.Text = main_ws.Cells(value, 5).value
            Me.NotesTextBox.Text = main_ws.Cells(value, 6).value
            Me.RowNumberTextBox.Text = value
        End If
    End If
End Property
 
Private Sub CancelCommandButton_Click()
    Unload Me
End Sub

Private Sub MaskPasswordCheckBox_Change()
    If Me.MaskPasswordCheckBox.value Then
        Me.PasswordTextBox.PasswordChar = "*"
    Else
        Me.PasswordTextBox.PasswordChar = ""
    End If
End Sub

Private Sub MaskPinCheckBox_Change()
    If Me.MaskPinCheckBox.value Then
        Me.PinTextBox.PasswordChar = "*"
    Else
        Me.PinTextBox.PasswordChar = ""
    End If
End Sub

Private Sub SaveCommandButton_Click()
    Dim main_ws As Worksheet
    Set main_ws = ActiveCell.Worksheet
    
    main_ws.Cells(lRowNum, 1).value = Me.NameTextBox.Text
    main_ws.Cells(lRowNum, 2).value = Me.UrlTextBox.Text
    main_ws.Cells(lRowNum, 3).value = Me.LoginTextBox.Text
    main_ws.Cells(lRowNum, 4).value = Me.PasswordTextBox.Text
    main_ws.Cells(lRowNum, 5).value = Me.PinTextBox.Text
    main_ws.Cells(lRowNum, 6).value = Me.NotesTextBox.Text
    Unload Me
End Sub

Private Sub UserForm_Activate()
    If lRowNum = 0 Then
        If ActiveCell.Row < 2 Then
            lRowNum = 2
        Else
            lRowNum = ActiveCell.Row
        End If
    End If
End Sub
