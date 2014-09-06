VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_meeting_create"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Create_Update_Click()
    Call Submit_Meeting
End Sub

Private Sub Submit_Meeting()
    On Error GoTo ErrorHandler
    Dim house_id As Long
    Dim start_time As Date
    Dim end_time As Date
    Dim meeting_date As Date
    Dim title As String
    Dim clients_present As Boolean
    Dim comment As String
    Dim intCount As Long ' Used with Dlookup
    
    
    
    If IsNull(Me.mc_cb_house_select.Value) = False Then ' Valid selection
        house_id = Me.mc_cb_house_select.Column(0) ' Get the house id
    Else
        MsgBox ("Please select a house from the house selection list")
        Resume ExitHandler
    End If
    
    If IsNull(Me.mc_start_time_text) = False Then
        start_time = Me.mc_start_time_text.Value
    Else
        MsgBox ("Please enter a start time for the meeting")
        Resume ExitHandler
    End If
    
    If IsNull(Me.mc_end_time_text) = False Then
        end_time = Me.mc_end_time_text.Value
    Else
        MsgBox ("Please enter an end time for the meeting")
        Resume ExitHandler
    End If
        
    If IsNull(Me.mc_date_select) = False Then
        meeting_date = Me.mc_date_select.Value
    Else
        MsgBox ("Please select a date for this meeting")
        Resume ExitHandler
    End If
    
    If IsNull(Me.mc_title_text) = False Then
        title = string_helpers.fStripIllegal(Me.mc_title_text.Value)
    Else
        MsgBox ("Please enter a title for this meeting")
        Resume ExitHandler
    End If
    
    If IsNull(Me.mc_clients_present) = False Then
        clients_present = Me.mc_clients_present.Value
    Else
        clients_present = False ' Optional
    End If
    
    If IsNull(Me.mc_add_text_comment) = False Then
        comment = string_helpers.fStripIllegal(Me.mc_add_text_comment.Value)
    Else
        comment = " " ' Optional
    End If
    
    Dim attrs As Variant: attrs = Array("[house_id]", "[date]", "[start]", "[end]", "[title]", "[clients_present]", "[comments]")
    Dim insertAttrs As Variant: insertAttrs = Array("" & house_id, "'" & meeting_date & "'", "'" & start_time & "'", "'" & end_time & "'", _
    "'" & title & "'", "" & clients_present, "'" & comment & "'")
    Dim insertVals As String: insertVals = string_helpers.join_array_to_str(insertAttrs, ",")
    
    Dim query As String: query = string_helpers.assemble_query(InsertQuery, attrs, Array("Meetings"), , insertVals)

    Call helpers.exec_query(query, Application.CurrentDb)  ' Insert
    
    Dim meetingId As Integer
    Dim rs2 As recordSet
    Set rs2 = CurrentDb.OpenRecordset("SELECT ID FROM meetings;")
  '  MsgBox ("opened record set")
    If rs2.RecordCount > 0 Then ' if there are any clients then proceed '
        Do While Not rs2.EOF
            rs2.MoveLast
'            MsgBox ("moved to last")
            meetingId = rs2![ID]
 '           MsgBox ("got id")
            rs2.MoveNext
        Loop
    End If
    rs2.Close
    
    'now we have house_id and meetingId,start_time,end_time,meeting_date,comment
    'start_time, end_time, employee_id, house_id, shift_date, comments,ihss = no, sleep_shift = no,
    'awake = no, exception_shift = Meeting:House
    'shift_default
    'SELECT DISTINCT [employees].[ID] FROM employees, shift_default WHERE [employees].[ID] = [shift_default].[employee_id];
    Dim tempstring As String
    
    Dim rs1 As recordSet
    Set rs1 = CurrentDb.OpenRecordset("SELECT DISTINCT [employees].[ID] FROM employees, shift_default WHERE [employees].[ID] = [shift_default].[employee_id] AND [shift_default].[house_id] = " & house_id & ";")
    If rs1.RecordCount > 0 Then ' if there are any clients then proceed '
        Do While Not rs1.EOF
            Dim mySqlString As String
            
             mySqlString = "INSERT INTO shift_default(start_time, end_time, employee_id, house_id, shift_date, comments,ihss, sleep_shift,awake, exception_shift) " & _
                               "VALUES (#" & start_time & "#,#" & end_time & "#," & CInt(rs1![ID]) & ", " & house_id & ",#" & meeting_date & "#,'" & comment & "','No','No','No','Meeting:House');"
           ' MsgBox (mySqlString)
            
            CurrentDb.Execute (mySqlString)
            CurrentDb.Execute "INSERT INTO meetings_employees (employee_id, meeting_id) VALUES (" & CInt(rs1![ID]) & "," & meetingId & ");"
            'tempstring = tempstring & rs1![ID] & ","
            rs1.MoveNext
        Loop
    End If
    rs1.Close
    'MsgBox (tempstring)
    
    
    
    
   ' MsgBox (meetingId)
    
    'meetingId = DLookup("start_time", "meetings", "id = " & shiftNumber)
    'we need to search the shifts table for employees associated with this house
    ' return distinct on the employees
    ' add each employee + meeting number to DB
    
    MsgBox ("Meeting Successfully Created")
DoCmd.Close
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch this exception
        Case Else ' All other cases
            MsgBox ("Meeting Create Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
       


End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim x As Variant ' Will be used to hold the values of Me.OpenArgs as it is split
    Dim house_id As Long
    
    Call update_houses_list ' First update the houses list manually
    
    If IsNull(Me.OpenArgs) = False Then ' Check for passed in parameters via the Global Me.OpenArgs variable
    
    
    ' split based on | delimeter
    
        x = Split(Me.OpenArgs, "|")
        strCtrl = x(0) ' The control name
        house_id = x(1) ' The house ID
        Me(strCtrl).Value = house_id
    
    End If
End Sub
Private Sub update_houses_list()
Dim prepString As String
Dim valIsNull As Boolean: valIsNull = True ' By default this value is null
Dim rs1 As recordSet
Set rs1 = CurrentDb.OpenRecordset("SELECT ID FROM houses;") ' get all the houses '

If IsNull(Me.mc_cb_house_select.Value) = False Then ' Store the old value index
    valIsNull = False ' Not Null
End If

Call helpers.clear_cb_items(Me, "mc_cb_house_select") ' Clear/reset it


rs1.MoveFirst


Do While Not rs1.EOF

    Dim rs2 As recordSet
    prepString = "" ' reset stringbuilder '
    Set rs2 = CurrentDb.OpenRecordset("SELECT first_name FROM clients WHERE house_id = " & rs1![ID] & ";") ' get the clients with this house # '
    If rs2.RecordCount > 0 Then ' if there are any clients then proceed '
    
        Do While Not rs2.EOF
            prepString = prepString + rs2![first_name] + " " ' build the string '
            rs2.MoveNext
        Loop
        
    
    Dim strItem As String
    strItem = rs1![ID] & ";" & prepString
    Me.mc_cb_house_select.AddItem Item:=strItem ' Populate the list
     
    
    ' Column(0) = rs1![ID] ' set cb[0] to ID '
    ' Me.ha_cb_house_select.Column(1) = prepString ' set cb[1] to built string '
   
 End If
   
   rs1.MoveNext
Loop



If valIsNull = False Then
   Me.mc_cb_house_select.Value = Me.mc_cb_house_select.OldValue ' Better way of retrieving the oldvalue via Microsoft Access
End If

Me.mc_cb_house_select.SetFocus ' This lets us add a visual emphasis on the combo box list



End Sub
