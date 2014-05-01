VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ihss_eligibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database




Private Sub Command35_Click()

End Sub

Private Sub ClearBtn_Click()
    Call helpers.clear_form(Me)
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    Dim x As Variant
    Dim strControl As String
    Dim lngID As Long
    If Len(Me.OpenArgs) > 0 Then
    
    
    ' split based on | delimeter
    
    x = Split(Me.OpenArgs, "|")
    strControl = x(0)
    lngID = x(1)
    'MsgBox ("strControl = '" & strControl & "' and ID = '" & lngID & "'")
    Me(strControl).Value = lngID
    
    
    End If
    
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            msgbox ("Error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_Open(Cancel As Integer)


End Sub

Private Sub ihss_cb_client_select_Click()

End Sub

Private Sub ihss_cb_emp_select_Click()

End Sub

Private Sub ihss_chb_check_Click()

End Sub

Private Sub ihss_update_Click()
    Call submit_ihss
End Sub


Private Sub submit_ihss()
    On Error GoTo ErrorHandler
    Dim intCount As Integer: intCount = 0
    Dim emp_id As Integer: emp_id = 0
    Dim client_id As Integer: client_id = 0
    Dim eligible As Boolean

    
    If IsNull(ihss_cb_emp_select.Value) = False Then
        emp_id = ihss_cb_emp_select.Column(0)
    Else
        msgbox ("Please select an employee before continuing")
        Exit Sub
    End If
    
    If IsNull(ihss_cb_client_select.Value) = False Then
        client_id = ihss_cb_client_select.Column(0)
    Else
        msgbox ("Please select an employee before continuing")
        Exit Sub
    End If
    
    If IsNull(ihss_cb_eligibility_select.Value) = False Then
        If (ihss_cb_eligibility_select.Value = "Yes") Then
            eligible = True
        ElseIf (ihss_cb_eligibility_select.Value = "No") Then
            eligible = False
        End If
        
       
    Else
        msgbox ("You have not selected whether this employee is IHSS eligible with this client.")
        Exit Sub
    End If
    
    
    Dim whereQuery As String: whereQuery = string_helpers.join_array_to_str(Array("WHERE employee_id = " & emp_id, "and client_id = " & client_id))
    
    Dim query As String: query = string_helpers.assemble_query(SelectQuery, Array("ID", "employee_id", "client_id"), _
    Array("ihss_eligibility"), whereQuery)
    
    Dim rst As DAO.recordSet: Set rst = helpers.get_record(query, Application.CurrentDb)
    intCount = rst.RecordCount
    If intCount > 0 Then ' Found so update the record
        query = string_helpers.assemble_query(updateQuery, Array("ihss_eligible = " & eligible), Array("ihss_eligibility"), "WHERE ID = " & rst("ID"))
        Call helpers.exec_query(query, Application.CurrentDb)
        
         msgbox ("Successfully updated the pre-existing IHSS eligibility record")
        
    Else
        Dim answer As VbMsgBoxResult
        answer = msgbox("A Pre-existing IHSS Eligiblity record was not found, would you like to create one?", vbQuestion + vbYesNo)
        If answer = vbYes Then ' Then go ahead and create the record 5/1/2014
            Dim insertAttrs As Variant: insertAttrs = Array("" & emp_id, "" & client_id, "" & eligible)
            Dim insertVals As String: insertVals = string_helpers.join_array_to_str(insertAttrs, ",")
            query = string_helpers.assemble_query(InsertQuery, Array("employee_id", "client_id", "ihss_eligible"), Array("ihss_eligibility"), _
            , insertVals)
            
           Call helpers.exec_query(query, Application.CurrentDb)
           ' msgbox (query)
           msgbox ("Successfully created a new IHSS eligibility record")
            
            
            
        End If
        
    End If
    
    Set rst = Nothing ' Clear it out for the garbage collector

    
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            'Handle silently
        Case Else ' All other cases
            msgbox ("Submit_ihss Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
    
End Sub

