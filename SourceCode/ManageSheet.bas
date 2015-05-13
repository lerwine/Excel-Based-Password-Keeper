Private Sub Worksheet_Activate()
    local_Update_SelectedItem_Cells
End Sub

Sub Update_SelectedItem_Cells()
    Set prevActiveWks = ActiveCell.Worksheet
    
    If prevActiveWks.name = ManageSheet.name Then
        local_Update_SelectedItem_Cells
    Else
        ManageSheet.Activate
        local_Update_SelectedItem_Cells
        prevActiveWks.Activate
    End If
End Sub

Private Sub local_Update_SelectedItem_Cells()
    ManageSheet.Unprotect
    ManageSheet.Cells(2, 2).value = (DataSheet.SelectedRowIndex - 1)
    ManageSheet.Cells(3, 2).value = DataSheet.SelectedName
    ManageSheet.Cells(4, 2).value = DataSheet.SelectedUrl
    ManageSheet.Cells(5, 2).value = DataSheet.SelectedLogin
    ManageSheet.Cells(7, 2).value = DataSheet.SelectedNotes
    ManageSheet.Protect
End Sub
