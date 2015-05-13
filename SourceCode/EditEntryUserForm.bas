Private lRowNum As Long
Private lOriginalName As String
Private lOriginalUrl As String
Private lOriginalLogin As String
Private lOriginalPassword As String
Private lOriginalPin As String
Private lOriginalNotes As String

Sub InitializeControls()
    lRowNum = DataSheet.SelectedRowIndex
    lOriginalName = DataSheet.SelectedName
    Me.NameTextBox.Text = lOriginalName
    lOriginalUrl = DataSheet.SelectedUrl
    Me.UrlTextBox.Text = lOriginalUrl
    lOriginalLogin = DataSheet.SelectedLogin
    Me.LoginTextBox.Text = lOriginalLogin
    lOriginalPassword = DataSheet.SelectedPassword
    Me.PasswordTextBox.Text = lOriginalPassword
    lOriginalPin = DataSheet.SelectedPin
    Me.PinTextBox.Text = lOriginalPin
    lOriginalNotes = DataSheet.SelectedNotes
    Me.NotesTextBox.Text = lOriginalNotes
    Me.RowNumberTextBox.Text = lRowNum
End Sub

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
    Dim result As VbMsgBoxResult
    
    DataSheet.SuspendChangeEvents = DataSheet.SuspendChangeEvents + 1
    
    If lOriginalName <> DataSheet.GetNameAt(lRowNum) Or lOriginalUrl <> DataSheet.GetUrlAt(lRowNum) _
            Or lOriginalLogin <> DataSheet.GetLoginAt(lRowNum) Or lOriginalPassword <> DataSheet.GetPasswordAt(lRowNum) _
            Or lOriginalPin <> DataSheet.GetPinAt(lRowNum) Or lOriginalNotes <> DataSheet.GetNotesAt(lRowNum) Then
        result = MsgBox("Values on the spreadsheet have changed while this form was open. Do you want to save?", vbYesNo, "Changes detected")
    Else
        result = vbYes
    End If
    
    If result = vbYes Then
        DataSheet.UpdateName lRowNum, Me.NameTextBox.Text
        DataSheet.UpdateUrl lRowNum, Me.UrlTextBox.Text
        DataSheet.UpdateLogin lRowNum, Me.LoginTextBox.Text
        DataSheet.UpdatePassword lRowNum, Me.PasswordTextBox.Text
        DataSheet.UpdatePin lRowNum, Me.PinTextBox.Text
        DataSheet.UpdateNotes lRowNum, Me.NotesTextBox.Text
        Unload Me
    End If
    
    DataSheet.SuspendChangeEvents = DataSheet.SuspendChangeEvents - 1
End Sub
