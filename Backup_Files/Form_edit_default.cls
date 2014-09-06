VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_edit_default"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cancel_Click()
    DoCmd.Close
End Sub

Private Sub Command35_Click()

 Dim startTime As Date
    Dim endTime As Date
    Dim employeeId As Integer
    Dim clientId As Integer
    Dim shiftNumber As Integer
    Dim ihss As String
    Dim overnight_value As String
    Dim dayOfWeek As Integer
    
    
    dayOfWeek = [Forms]![edit_defaults_select]![day_of_week].Value
    
    
    If IsNull([Forms]![edit_default]![cb_overnight].Value) = False Then
        overnight_value = [Forms]![edit_default]![cb_overnight].Value
    Else
        overnight_value = "No"
    End If

    
    startTime = [Forms]![edit_default]![textBox_startTime].Value
    endTime = [Forms]![edit_default]![textBox_endTime].Value


    If IsNull([Forms]![edit_default]![Combo9].Value) = False Then
        employeeId = [Forms]![edit_default]![Combo9].Value
    Else
        employeeId = 0
    End If

    If IsNull([Forms]![edit_default]![Combo11].Value) = False Then
        clientId = [Forms]![edit_default]![Combo11].Value
    End If

    If IsNull([Forms]![edit_default]![Combo33].Value) = False Then
        ihss = [Forms]![edit_default]![Combo33].Value
    Else
        ihss = "No"
    End If

    If Command35.Caption = "Add" Then
        Dim house_id As Integer
        house_id = CInt([Forms]![edit_defaults_select]![house_id].Value)
        
        If IsNull([Forms]![edit_default]![Combo9].Value) = True Then
            MsgBox ("Employee Cannot Be Left Blank")
        Else
            CurrentDb.Execute "INSERT INTO default_shift (day_of_week,start_time,end_time,employee_id,house_id,client_id,ihss,sleep_shift) VALUES (" & dayOfWeek & ",#" & startTime & "#,#" & endTime & "#," & employeeId & ", " & house_id & "," & clientId & ",'" & ihss & "','" & overnight_value & "');"
            MsgBox ("Shift Successfully Added")
            [Forms]![edit_defaults_select].Refresh
            DoCmd.Close
        End If
    Else 'This is an EDIT!
        shiftNumber = CInt([Forms]![edit_defaults_select]![shift_sel].Value)
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("UPDATE default_shift SET [day_of_week] = " & dayOfWeek & ",[sleep_shift] = '" & overnight_value & "', [ihss] = '" & ihss & "', [client_id] = " & clientId & ", [employee_id] = " & employeeId & ", [start_time] = #" & startTime & "#, [end_time] = #" & endTime & "# WHERE [ID] = " & shiftNumber)
        DoCmd.SetWarnings True
        [Forms]![edit_defaults_select].Refresh
        DoCmd.Close
    End If

End Sub

Private Sub Form_Load()
    
    If IsNull([Forms]![edit_defaults_select]![house_label].Value) = False Then
        house_label.Caption = [Forms]![edit_defaults_select]![house_label].Value
    Else
        DoCmd.Close
        Exit Sub
    End If
       
    Dim shiftNumber As Integer
    If IsNull([Forms]![edit_defaults_select]![shift_sel].Value) = True Then
        Dim house_id As Integer
        house_id = CInt([Forms]![edit_defaults_select]![house_id].Value)
        Command35.Caption = "Add"
        textBox_startTime.Value = "00:01"
        textBox_endTime.Value = "01:00"
    Else
        Command35.Caption = "Edit"
        shiftNumber = CInt([Forms]![edit_defaults_select]![shift_sel].Value)
        textBox_startTime.Value = DLookup("start_time", "default_shift", "id = " & shiftNumber)
        textBox_endTime.Value = DLookup("end_time", "default_shift", "id = " & shiftNumber)
        Combo9.Value = DLookup("employee_id", "default_shift", "id = " & shiftNumber)
        Combo11.Value = DLookup("client_id", "default_shift", "id = " & shiftNumber)
        Combo33.Value = DLookup("ihss", "default_shift", "id = " & shiftNumber)
        cb_overnight.Value = DLookup("sleep_shift", "default_shift", "id = " & shiftNumber)
    End If
        
End Sub

Private Sub house_label_Click()

End Sub
