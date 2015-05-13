Private Sub Workbook_Open()
    DataSheet.InitializeSelectedProperties
    If ActiveCell.Worksheet.name = ManageSheet.name Then ManageSheet.Update_SelectedItem_Cells
End Sub
