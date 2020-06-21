''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''DevBy:[AJ]''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''FormCode'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CommandButton1_Click()

End Sub

Private Sub cmdreset_Click()
  Dim msgvalue As VbMsgBoxResult
    msgvalue = MsgBox("Do you want to reset the form?", vbYesNo + vbInformation, "confirmation")
     
    If msgvalue = vbNo Then Exit Sub
    
    Call Reset

End Sub

Private Sub cmdsave_Click()
    Dim msgvalue As VbMsgBoxResult
    msgvalue = MsgBox("Do you want to save the data?", vbYesNo + vbInformation, "confirmation")
     
    If msgvalue = vbNo Then Exit Sub
    Call submit
    Call Reset


End Sub


Private Sub UserForm_Initialize()

Call Reset

End Sub
