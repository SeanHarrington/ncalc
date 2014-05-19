VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_employee_blackout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    Dim x As Variant
    Dim strCtrl As String
    Dim strCtrl2 As String
    Dim name As String
    Dim empID As Long
    If IsNull(Me.OpenArgs) = False Then
    
    
    ' split based on | delimeter
    
        x = Split(Me.OpenArgs, "|")
        strCtrl = x(0) ' The control name
        name = x(1) ' The employee name
        strCtrl2 = x(2) ' Second control name
        empID = x(3) ' Employee record number
        Me(strCtrl).Value = name
        Me(strCtrl2).Value = empID
    
    
    End If
    
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("Error On Employee Blackout Load Form: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Dim x As Variant
    Dim strControl As String
    Dim name As String
    
    If IsNull(Me.OpenArgs) = True Then ' Then throw an error, we should not be coming in here with zero args
        MsgBox ("Please open the employee add form to select an existing employee to perform a blackout")
        DoCmd.Close acForm, "employee_blackout" ' Close the form
        Exit Sub ' Trigger case 20
    End If
    
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("Error On Employee Blackout Load Form: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub Post_Click()
   Call Submit_Post
End Sub

Private Sub Submit_Post()
     On Error GoTo ErrorHandler
     Dim employeeID As Long
     Dim last_name As String
     Dim first_name As String
     Dim blackout_type As String
     Dim start_date As Date
     Dim end_date As Date
     Dim x As Variant ' Array
     
     If IsNull(Me.empb_name_text) = False Then
        x = Split(Me.empb_name_text, ",") ' Split by comma
        last_name = Trim(x(0)) ' Trim the whitespace from this array value
        first_name = Trim(x(1)) ' Same as above for the second array value
     End If
     
     
    If IsNull(Me.empb_cb_select_type) = False Then
        blackout_type = Me.empb_cb_select_type.Value
    End If
     
    If IsNull(Me.empb_start_date_calendar) = False Then
        start_date = Me.empb_start_date_calendar
    Else
        MsgBox ("Please select a start date")
        Exit Sub
    End If
     
    If IsNull(Me.empb_end_date_calendar) = False Then
        end_date = Me.empb_end_date_calendar
    Else
        MsgBox ("Please select an end date")
        Exit Sub
    End If
    
    If helpers.valid_dates(start_date, end_date) = False Then
        MsgBox ("End date is greater than start date, please select proper dates.")
        Exit Sub
    End If
    
    employeeID = empb_box_id.Value ' Do not need to convert to long, already a general number
    Dim insertAttrs As Variant: insertAttrs = Array("" & employeeID, "'" & start_date & "'", "'" & end_date & "'", _
     "'" & blackout_type & "'")
    
    Dim attrs As Variant: attrs = Array("employee_id", "blackout_start", "blackout_end", "type")
    Dim insertVals As String: insertVals = string_helpers.join_array_to_str(insertAttrs, ",")
    Dim query As String: query = string_helpers.assemble_query(InsertQuery, attrs, Array("employee_blackouts"), , insertVals)
    
    Call helpers.exec_query(query, Application.CurrentDb)
    
    MsgBox ("Successfully posted blackout dates for " & first_name & ", " & last_name & "")
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("Error On Employee Blackout Post: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

