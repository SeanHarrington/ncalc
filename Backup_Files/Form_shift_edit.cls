VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_shift_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub bt_add_client_Click()

    Dim clientId As Integer
    Dim ihssHours As Integer
    Dim ctrlListBox As ListBox
    Dim firstName As String
    Dim Continue As Boolean

    If IsNull([Forms]![shift_edit]![Combo11].Value) = False Then
        clientId = [Forms]![shift_edit]![Combo11].Value
        If clientId > 0 Then
            If IsNull(Me.text_field_ihss_hours.Value) = False Then
                ihssHours = CInt(Me.text_field_ihss_hours.Value)
            Else
                ihssHours = 0
            End If
            Set ctrlListBox = Me.list_box_shift_client
            Continue = True
            
            For i = 0 To Me.list_box_shift_client.ListCount - 1
                If CInt(Me.list_box_shift_client.Column(0, i)) = CInt(clientId) Then
                    Continue = False
                End If
            Next i
            
            If Continue = True Then
                firstName = DLookup("first_name", "clients", "id = " & clientId)
                Dim hours As Integer
                hours = getRemainingIHSS(clientId)
                ctrlListBox.AddItem Item:=CStr(clientId) & ";" & CStr(firstName) & ";" & CStr(ihssHours) & ";" & CStr(hours - ihssHours)
            End If
        End If
    End If

End Sub
Private Function getRemainingIHSS(ByVal clientId As Integer) As Integer

    Dim BegDate As Date
    Dim EndDate As Date
    Dim NowDate As Date
    Dim TotalIHSS As Integer
    Dim MonthlyIHSS As Integer
    Dim sqlstring As String
    Dim shiftNumber As Integer
    
    NowDate = CDate(Text12.Caption)
    EndDate = DateSerial(Year(NowDate), Month(NowDate) + 1, 1) ' get next months last day
    BegDate = DateSerial(Year(NowDate), Month(NowDate), 0) ' get last months last day
    
    If Command35.Caption = "Update" Then
        shiftNumber = CInt([Forms]![day_view]![shift_sel].Value)
    Else
        shiftNumber = 0
    End If
    
    sqlstring = "SELECT SUM (shift_client.ihss) as total FROM shift_client, shift_default WHERE shift_client.client_id = " & clientId & " AND shift_client.shift_id = shift_default.id AND shift_default.shift_date > #" & BegDate & "# AND shift_default.shift_date < #" & EndDate & "# AND shift_client.shift_id <> " & shiftNumber & ";"
    Set rs1 = CurrentDb.OpenRecordset(sqlstring)
    If rs1.RecordCount > 0 Then ' if there are any clients then proceed '
        Do While Not rs1.EOF
            If IsNull(rs1![total]) Then
                MonthlyIHSS = 0
                TotalIHSS = 0
            Else
                TotalIHSS = CInt(rs1![total])
            End If
            rs1.MoveNext
        Loop
    End If
    rs1.Close
            
    If IsNull(DLookup("monthly_ihss", "clients", "id = " & clientId)) Then
        MonthlyIHSS = 0
    Else
        MonthlyIHSS = CInt(DLookup("monthly_ihss", "clients", "id = " & clientId))
    End If
    getRemainingIHSS = MonthlyIHSS - TotalIHSS

End Function
Private Sub bt_remove_client_Click()

    Dim clientId As Integer

    If IsNull(list_box_shift_client.Value) = False Then
        clientId = CInt(list_box_shift_client.Value)
        list_box_shift_client.RemoveItem (list_box_shift_client.ListIndex)
    End If

End Sub

Private Sub cb_overnight_Change()

    Dim overnight_value As String
    overnight_value = cb_overnight.Value
    If overnight_value = "Yes" Then
        cb_awakened.Enabled = True
    Else
        cb_awakened.Enabled = False
    End If
    cb_awakened.Value = "No"

End Sub

Private Sub Command35_Click()

    Dim startTime As Date
    Dim endTime As Date
    Dim employeeId As Integer
    Dim comments As String
    Dim clientId As Integer
    Dim shiftNumber As Integer
    Dim ihss As String
    Dim overnight_value As String
    Dim awakened_value As String
    Dim exception_value As String

    If IsNull([Forms]![shift_edit]![cb_awakened].Value) = False Then
        awakened_value = [Forms]![shift_edit]![cb_awakened].Value
    Else
        awakened_value = "No"
    End If

    If IsNull([Forms]![shift_edit]![cb_overnight].Value) = False Then
        overnight_value = [Forms]![shift_edit]![cb_overnight].Value
    Else
        overnight_value = "No"
    End If

    If IsNull([Forms]![shift_edit]![Combo7].Value) = False Then
        exception_value = [Forms]![shift_edit]![Combo7].Value
    Else
        exception_value = "No"
    End If
 
    If IsNull([Forms]![shift_edit]![emp_add_text_comment].Value) = False Then
        comments = string_helpers.fStripIllegal([Forms]![shift_edit]![emp_add_text_comment].Value)
    End If

    startTime = [Forms]![shift_edit]![textBox_startTime].Value
    endTime = [Forms]![shift_edit]![textBox_endTime].Value

    If IsNull([Forms]![shift_edit]![Combo9].Value) = False Then
        employeeId = [Forms]![shift_edit]![Combo9].Value
    Else
        employeeId = 0
    End If

    If IsNull([Forms]![shift_edit]![Combo11].Value) = False Then
        clientId = [Forms]![shift_edit]![Combo11].Value
    End If
    
    If Command35.Caption = "Add" Then
        Dim house_id As Integer
        house_id = CInt([Forms]![day_view]![house_id].Value)
        Dim shift_date As Date
        shift_date = CDate(Text12.Caption)
        
        If IsNull([Forms]![shift_edit]![Combo9].Value) = True Then
            MsgBox ("Employee Cannot Be Left Blank")
        Else
            Dim db As DAO.Database
            Dim newRow As Long
            Set db = CurrentDb
            db.Execute "INSERT INTO shift_default (start_time,end_time,employee_id,house_id,shift_date,comments,sleep_shift,awake,exception_shift) VALUES (#" & startTime & "#,#" & endTime & "#," & employeeId & ", " & house_id & ",#" & shift_date & "#,'" & comments & "','" & overnight_value & "','" & awakened_value & "','" & exception_value & "');"
            newRow = db.OpenRecordset("SELECT @@IDENTITY")(0)
            Set db = Nothing
            Call UpdateClientValues(CInt(newRow))
            MsgBox ("Shift Successfully Added")
            [Forms]![day_view].Refresh
            DoCmd.Close
        End If
    Else 'This is an EDIT!
        shiftNumber = CInt([Forms]![day_view]![shift_sel].Value)
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("UPDATE shift_default SET [exception_shift] = '" & exception_value & "',[sleep_shift] = '" & overnight_value & "', [awake] = '" & awakened_value & "', [employee_id] = " & employeeId & ", [start_time] = #" & startTime & "#, [end_time] = #" & endTime & "#, [comments] = '" & comments & "' WHERE [ID] = " & shiftNumber)
        DoCmd.SetWarnings True
        Call UpdateClientValues(0)
        [Forms]![day_view].Refresh
        DoCmd.Close
    End If

End Sub
Private Sub UpdateClientValues(ByVal shiftNumber As Integer)

    Dim clientId As Integer
    Dim ihss As Integer
    Dim employeeId As Integer
    Dim i As Integer

    If shiftNumber = 0 Then
        shiftNumber = CInt([Forms]![day_view]![shift_sel].Value)
    End If

    DoCmd.SetWarnings False
    DoCmd.RunSQL ("DELETE * FROM shift_client WHERE shift_id = " & shiftNumber)
    DoCmd.SetWarnings True

    employeeId = CInt(Combo9.Value)
    For i = 0 To Me.list_box_shift_client.ListCount - 1
        clientId = CInt(Me.list_box_shift_client.Column(0, i))
        ihss = CInt(Me.list_box_shift_client.Column(2, i))
        CurrentDb.Execute "INSERT INTO shift_client (client_id,employee_id,shift_id,ihss) VALUES (" & clientId & ", " & employeeId & ", " & shiftNumber & ", " & ihss & ");"
    Next i

End Sub


Private Sub Form_Load()
    
    Dim shiftNumber As Integer
    
    If IsNull([Forms]![day_view]![house_label].Caption) = False Then
        house_label.Caption = [Forms]![day_view]![house_label].Caption
    Else
        DoCmd.Close
        Exit Sub
    End If
    
    FormType = "Edit"
    If IsNull([Forms]![day_view]![shift_sel].Value) = True Then
        Dim house_id As Integer
        house_id = CInt([Forms]![day_view]![house_id].Value)
        Command35.Caption = "Add"
        Text12.Caption = [Forms]![calendar_view].[dateStore].Value
        textBox_startTime.Value = "12:00"
        textBox_endTime.Value = "12:00"
    Else
        Command35.Caption = "Update"
        shiftNumber = CInt([Forms]![day_view]![shift_sel].Value)
        textBox_startTime.Value = DLookup("start_time", "shift_default", "id = " & shiftNumber)
        textBox_endTime.Value = DLookup("end_time", "shift_default", "id = " & shiftNumber)
        Combo9.Value = DLookup("employee_id", "shift_default", "id = " & shiftNumber)
        emp_add_text_comment.Value = DLookup("comments", "shift_default", "id = " & shiftNumber)
        Text12.Caption = DLookup("shift_date", "shift_default", "id = " & shiftNumber)
        cb_awakened.Value = DLookup("awake", "shift_default", "id = " & shiftNumber)
        cb_overnight.Value = DLookup("sleep_shift", "shift_default", "id = " & shiftNumber)
        Combo7.Value = DLookup("exception_shift", "shift_default", "id = " & shiftNumber)
        Call update_client_list(shiftNumber)
        If cb_overnight.Value = "Yes" Then
            cb_awakened.Enabled = True
        End If
    End If
    
End Sub

Private Sub update_client_list(ByVal shiftNum As Integer)
    
    Dim prepString As String
    Dim ctrlListBox As ListBox
    Dim rs2 As recordSet
    Dim clientId As Integer
    Dim BegDate As Date
    Dim EndDate As Date
    Dim NowDate As Date
    Dim TotalIHSS As Integer
    Dim MonthlyIHSS As Integer
    
    NowDate = CDate(DLookup("shift_date", "shift_default", "id = " & shiftNum))
    EndDate = DateSerial(Year(NowDate), Month(NowDate) + 1, 1) ' get next months last day
    BegDate = DateSerial(Year(NowDate), Month(NowDate), 0) ' get last months last day
    Set ctrlListBox = Me.list_box_shift_client
        
    Set rs2 = CurrentDb.OpenRecordset("SELECT shift_client.client_id, clients.first_name, shift_client.ihss FROM clients , shift_client WHERE shift_client.shift_id = " & shiftNum & " AND shift_client.client_id = clients.id;")
    prepString = "" ' reset stringbuilder '
    If rs2.RecordCount > 0 Then ' if there are any clients then proceed '
        Do While Not rs2.EOF
            clientId = rs2![client_id]
            Dim sqlstring As String
            sqlstring = "SELECT SUM (shift_client.ihss) as total FROM shift_client, shift_default WHERE shift_client.client_id = " & rs2![client_id] & " AND shift_client.shift_id = shift_default.id AND shift_default.shift_date > #" & BegDate & "# AND shift_default.shift_date < #" & EndDate & "#;"
            Set rs1 = CurrentDb.OpenRecordset(sqlstring)
            If rs1.RecordCount > 0 Then ' if there are any clients then proceed '
                Do While Not rs1.EOF
                    If IsNull(rs1![total]) Then
                        MonthlyIHSS = 0
                        TotalIHSS = 0
                    Else
                        TotalIHSS = CInt(rs1![total])
                        If IsNull(DLookup("monthly_ihss", "clients", "id = " & clientId)) Then
                            MonthlyIHSS = 0
                        Else
                            MonthlyIHSS = CInt(DLookup("monthly_ihss", "clients", "id = " & clientId))
                        End If
                    End If
                    rs1.MoveNext
                Loop
            End If
            rs1.Close
            ctrlListBox.AddItem Item:=CStr(rs2![client_id]) & ";" & CStr(rs2![first_name]) & ";" & CStr(rs2![ihss]) & ";" & CStr(MonthlyIHSS - TotalIHSS)
            rs2.MoveNext
        Loop
    End If
    rs2.Close

End Sub


Private Sub ihss_btn_cancel_Click()

    [Forms]![day_view].Refresh
    DoCmd.Close

End Sub

