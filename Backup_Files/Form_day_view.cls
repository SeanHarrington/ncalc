VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_day_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub add_button_Click()
    
    Me.shift_sel.Value = Null
    DoCmd.OpenForm "shift_edit"

End Sub

Private Sub button_edit_shift_Click()

    If IsNull(List0.Value) = False Then
        Me.shift_sel.Value = List0.Value
        DoCmd.OpenForm "shift_edit"
    End If
      
End Sub

Private Sub Command18_Click()

End Sub

Private Sub delete_button_Click()

On Error GoTo Err_delete_button_Click
    Dim Answer As Integer
    Answer = MsgBox("Are you sure you wish to delete this record?", vbYesNo + vbExclamation + vbDefaultButton2, "Delete Confirmation")
    If Answer = vbYes Then
        DoCmd.SetWarnings False
        Dim shift_number As Integer
        If IsNull(List0.Value) = False Then
            shiftNumber = CInt(List0.Value)
            CurrentDb.Execute "delete * from shift_default where ID = " & shiftNumber
            CurrentDb.Execute "delete * from shift_client where shift_id = " & shiftNumber
            Small_load
        Me.Refresh
        End If
        DoCmd.SetWarnings True
    Else
    End If

Exit_delete_button_Click:
    Exit Sub

Err_delete_button_Click:
    MsgBox Err.Description
    Resume Exit_delete_button_Click
    
End Sub

Private Sub Form_GotFocus()
    
    Small_load
    Me.Refresh

End Sub

Private Sub Form_Load()
    
    Small_load

End Sub

Private Sub Small_load()
    
    Me.house_id.Value = CInt([Forms]![calendar_view]![house_id].Value)
    Me.day_of_week.Value = CInt([Forms]![calendar_view]![day_of_week].Value)
    update_houses_list

End Sub


Private Sub update_houses_list()
    
    Dim prepString As String
    Dim rs2 As recordSet
    
    prepString = "" ' reset stringbuilder '
    Set rs2 = CurrentDb.OpenRecordset("SELECT first_name FROM clients WHERE house_id = " & Me.house_id.Value & ";") ' get the clients with this house # '
    If rs2.RecordCount > 0 Then ' if there are any clients then proceed '
        Do While Not rs2.EOF
            prepString = prepString + rs2![first_name] + " " ' build the string '
            rs2.MoveNext
        Loop
    End If
    Me.house_label.Caption = prepString
    rs2.Close

End Sub

Private Sub List0_AfterUpdate()

End Sub

Private Sub List0_BeforeUpdate(Cancel As Integer)

End Sub
