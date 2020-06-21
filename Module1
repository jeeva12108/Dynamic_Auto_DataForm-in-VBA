''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''DevBy:[AJ]'''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''Module Code'''''''''''''''''''''''''''''''''''''''''''''''
Sub Reset()
Dim iRow As Long
iRow = [Counta(Database!A:A)] 'identify the last row entry
With frmForm

    .txtid.Value = ""
    .txtname.Value = ""
    .txtname.Value = ""
    .txtcity.Value = ""
    .txtcountry.Value = ""
    .optmale.Value = False
    .optfemale.Value = False
    .cmbdepartment.Clear
    .cmbdepartment.AddItem "FCC"
    .cmbdepartment.AddItem "Hydro"
    .cmbdepartment.AddItem "Naphtha"
    .cmbdepartment.AddItem "Petro"
    .cmbdepartment.AddItem "Chemical"
    .cmbdepartment.AddItem "Tools"
    
    .lstdatabase.ColumnCount = 9
    .lstdatabase.ColumnHeads = True
    
    .lstdatabase.ColumnWidths = "30,60,75,40,60,45,55,70"
    
    If iRow > 1 Then
    .lstdatabase.RowSource = "Database!A2:I" & iRow
    Else
    .lstdatabase.RowSource = "Database!A2:I2"
    End If
   
End With
End Sub

Sub submit()

Dim sh As Worksheet
Dim iRow As Long

Set sh = ThisWorkbook.Sheets("Database")
iRow = [Counta(Database!A:A)] + 1 'gets the last row in the database and adds 1 to it
    With sh
    .Cells(iRow, 1) = iRow - 1 'to give S.No
    .Cells(iRow, 2) = frmForm.txtid.Value
    .Cells(iRow, 3) = frmForm.txtname.Value
    .Cells(iRow, 4) = IIf(frmForm.optfemale.Value = True, "Female", "Male") 'based on option selected by user
    .Cells(iRow, 5) = frmForm.cmbdepartment.Value
    .Cells(iRow, 6) = frmForm.txtcity.Value
    .Cells(iRow, 7) = frmForm.txtcountry.Value
    .Cells(iRow, 8) = Application.UserName
    .Cells(iRow, 9) = [Text(Now(), "DD-MM-YYYY HH:MM:SS")]
    End With
End Sub

Sub show_form()

   frmForm.Show


End Sub
