VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_employee_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' Enforce strict variable 'typing'

Private Sub Activate_Click()
    On Error GoTo ErrorHandler
    
    If Me.e_cb_name_select And IsNull(Me.e_cb_name_select.Column(0)) = False Then
        'Dim yes_no As Boolean
        Dim active As Boolean: active = Not (GetToggleVal()) ' Flip it
        Dim emp_id As Integer: emp_id = Me.e_cb_name_select.Column(0)
        Dim intCount As Integer: intCount = CLng(Nz(DLookup("ID", "employees", "ID = " & emp_id), 0))
        Dim query As String: query = "update employees set active = " & active
        query = query + " where ID = " & emp_id
        Call helpers.exec_query(query, Application.CurrentDb)
        Call ToggleActive(active)
    Else
        MsgBox ("Please select a valid employee to activate")
    End If
    
    
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            
        Case Else ' All other cases
            MsgBox ("Employee Activate Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub Blackout_Click()
On Error GoTo ErrorHandler
    
    If Me.e_cb_name_select And IsNull(Me.e_cb_name_select.Column(0)) = False Then
        Dim emp_id As Long: emp_id = Me.e_cb_name_select.Column(0) ' Get the idea to get the first and last name
        Dim emp_name As String: emp_name = Me.e_cb_name_select.Column(1) ' Get the employee full name
        DoCmd.Close acForm, "employee_blackout" ' Reload it, for now
        DoCmd.OpenForm "employee_blackout", , , , , , "empb_name_text|" & emp_name & "|empb_box_id|" & emp_id
       
    Else
        MsgBox ("Please select a valid employee before attempting to go to the employee black out form")
    End If
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            
        Case Else ' All other cases
            MsgBox ("Employee Blackout Error" + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub ClearBtn_Click()
    Call reset_form ' Reset the form to the default state
End Sub

Private Sub reset_form()

    Call helpers.clear_form(Me)
    Call helpers.change_control_caption("Create", Me.Create_Update)
    Call ToggleActive(2) ' Any value other than 0 or 1 will trigger total reset
    
End Sub



Private Sub Submit_Employee()
On Error GoTo ErrorHandler ' Error handling

Dim emp_id As Long ' changed to long
Dim first_name As String
Dim last_name As String
Dim middle_name As String
Dim pay_rate As Currency
Dim sick_leave_rate As Long ' Changed to long
Dim title As String
Dim phone As String
Dim regional_center As Long ' changed to long
Dim hire_date As Date
Dim comment As String
Dim yes_no As Boolean: yes_no = True
Dim fillColTab As String

Dim intCount As Integer: intCount = 0 ' Initialized to zero


If IsNull(Me.emp_add_text_first) = False Then
    first_name = string_helpers.fStripIllegal(Me.emp_add_text_first.Value)
Else
    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.emp_add_text_last) = False Then
    last_name = string_helpers.fStripIllegal(Me.emp_add_text_last.Value)
Else

    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.emp_add_text_middle) = False Then
    middle_name = string_helpers.fStripIllegal(Me.emp_add_text_middle.Value)
Else
    middle_name = " " ' Set optional field to empty value
End If

If IsNull(Me.e_cb_region_select) = False And Me.e_cb_region_select.ListIndex <> -1 Then
    regional_center = CLng(Me.e_cb_region_select.Column(0)) ' Grab the text value
Else
    regional_center = 0 ' Optional-default
End If

If IsNull(Me.emp_add_text_payrate) = False Then
    pay_rate = Format(val(string_helpers.fStripIllegal(Me.emp_add_text_payrate.Value)), "currency") ' val converts string to double
Else
    pay_rate = Format(0#, "currency")    ' Optional
End If

If IsNull(Me.emp_add_text_sickrate) = False Then
    sick_leave_rate = CLng(string_helpers.fStripIllegal(Me.emp_add_text_sickrate.Value))
Else
    sick_leave_rate = 0 ' Optional
End If

If IsNull(Me.emp_add_text_title) = False Then
    title = string_helpers.fStripIllegal(Me.emp_add_text_title.Value)
Else
    title = " " ' Optional
End If

If IsNull(Me.emp_add_text_phone) = False Then
    phone = string_helpers.fStripIllegal(Me.emp_add_text_phone.Value)
Else
    phone = " " ' Optional
End If

If IsNull(Me.emp_add_text_comment) = False Then
    comment = string_helpers.fStripIllegal(Me.emp_add_text_comment.Value)
Else
    comment = " " ' Optional
End If

If IsNull(Me.emp_select_date.Value) = False Then
    
    hire_date = Me.emp_select_date.Value  ' string_helpers.fStripIllegal would strip the '/' character from the data
Else
    hire_date = Date ' Call the MS access Date function
End If


' CLng(Nz(DLookup("ID", "Table1"),0))

    If IsNull(Me.e_cb_name_select.Value) = False Then
    
        emp_id = Me.e_cb_name_select.Column(0)
        intCount = CLng(Nz(DLookup("ID", "employees", "ID = " & emp_id), 0))
    End If
    
If intCount > 0 Then ' Update employee, since employee exists
    Dim lookupAttrs As Variant: lookupAttrs = Array("first_name = '" & first_name & "'", "last_name = '" & last_name & "'" _
    , "middle_initial = '" & middle_name & "'", "payrate = '" & pay_rate & "'", "sick_hour_rate = " & sick_leave_rate _
    , "title = '" & title & "'", "phone = '" & phone & "'", "regional_center = " & regional_center _
    , "comments = '" & comment & "'", "hiredate = '" & hire_date & "'")
    
    Dim query As String: query = string_helpers.assemble_query(updateQuery, lookupAttrs, Array("employees"), "WHERE ID = " & emp_id)
    
    Call helpers.exec_query(query, Application.CurrentDb)
     
    Me![e_cb_name_select].Requery ' Refresh
     
    Set lookupAttrs = Nothing ' Clear it out
Else ' Submit/Create new employee


Dim insertAttrs As Variant: insertAttrs = Array("'" & first_name & "'", "'" & last_name _
    & "'", "'" & middle_name & "'", "" & yes_no, "" & pay_rate _
    , "" & sick_leave_rate, "'" & title & "'", "'" & phone & "'" _
    , "" & regional_center, "'" & comment & "'", "'" & hire_date & "'")
    
Dim attrs As Variant: attrs = Array("first_name", "last_name", "middle_initial", "active", "payrate", "sick_hour_rate" _
, "title", "phone", "regional_center", "comments", "hiredate")

Dim insertVals As String: insertVals = string_helpers.join_array_to_str(insertAttrs, ",")

fillColTab = string_helpers.assemble_query(InsertQuery, attrs, Array("employees"), , insertVals)

'currentdb.Execute fillColTab, dbFailOnError
Call helpers.exec_query(fillColTab, Application.CurrentDb)


' Clean up section, garbage collector handles the real clean up
Set insertAttrs = Nothing
Set attrs = Nothing


Call reset_form ' Reset the form
Me![e_cb_name_select].Requery ' Refresh

Me.emp_add_text_last.SetFocus
MsgBox "New Employee Record Successfully Created"


End If


ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            MsgBox ("Please fill in required fields")
        Case Else ' All other cases
            MsgBox ("Employee Add Form Error" + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
    
End Sub

Private Sub Delete_Employee()
    On Error GoTo ErrorHandler
    
    Dim emp_id As Integer
    Dim intCount As Long
    Dim query As String
     If IsNull(Me.e_cb_name_select.Value) = False Then
        emp_id = Me.e_cb_name_select.Column(0)
        intCount = CLng(Nz(DLookup("ID", "employees", "ID = " & emp_id), 0))
    End If
    
    If intCount > 0 Then ' Valid delete since employee exists
        query = string_helpers.assemble_query(DeleteQuery, Nothing, Array("employees"), "WHERE ID = " & emp_id)
        Call helpers.exec_query(query, Application.CurrentDb)
    Else
        MsgBox ("Not a valid employee, please select an employee from the combobox")
        Exit Sub
    End If
    
    Call reset_form ' Reset the form
    Me![e_cb_name_select].Requery ' Refresh
    MsgBox "Employee Record Successfully Deleted"
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Silently exit
            ' Exit silent
        Case Else ' All other cases
            MsgBox ("Delete Employee Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
    
End Sub




Private Sub Create_Update_Click()
  
    Call Submit_Employee
    
End Sub

Private Sub Delete_Click()
    
    Call Delete_Employee

End Sub

Private Sub e_btn_ihss_Click()
On Error GoTo ErrorHandler
    
    If Me.e_cb_name_select And IsNull(Me.e_cb_name_select.Column(0)) = False Then
       Dim employee_id As Integer: employee_id = Me.e_cb_name_select.Column(0)
       
       DoCmd.Close acForm, "ihss_eligibility" ' Reload it, for now
       DoCmd.OpenForm "ihss_eligibility", , , , , , "ihss_cb_emp_select|" & Me.e_cb_name_select.Column(0)
       
    Else
        Dim Answer As VbMsgBoxResult
        Answer = MsgBox("Continue to the IHSS Eligibility Form without selecting an employee?", vbQuestion + vbYesNo)
        If Answer = vbYes Then
            DoCmd.OpenForm "ihss_eligibility" ' No options
        Else
            ' Else, stay here and don't open the form
            
        End If
        

    End If
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("Error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub e_cb_name_select_BeforeUpdate(Cancel As Integer)
On Error GoTo ErrorHandler

If Me.e_cb_name_select And IsNull(Me.e_cb_name_select.Column(0)) = True Then ' If ID doesn't exist then we have an invalid employee
    Resume ExitHandler ' Silently exit and let MS Access give us a dialog error and the dropdown list
End If


Dim objs2Ignore As Object: Set objs2Ignore = CreateObject("Scripting.Dictionary")
Call objs2Ignore.Add("e_cb_name_select", "e_cb_name_select") ' Key, Value Pair, key must always be unique.

Dim attrs As Variant: attrs = Array("last_name", "first_name", "middle_initial" _
, "regional_center", "phone", "sick_hour_rate", "payrate", "title", "comments", "hiredate", "active") ' Dictates the order to fill the boxes in



Dim employee_id As Integer: employee_id = Me.e_cb_name_select.Column(0) ' Retrieve the employee_id
Dim query As String: query = string_helpers.assemble_query(SelectQuery, attrs, Array("employees"), "WHERE ID = " & employee_id)
Dim rst As DAO.recordSet: Set rst = helpers.get_record(query, Application.CurrentDb)
'Set DBS = Application.CurrentDb
'Set rst = DBS.OpenRecordset("SELECT last_name, middle_initial, first_name FROM employees WHERE ID = " & employee_id) ' Explicit values to select
'Set rst = helpers.get_record(query, Application.CurrentDb)

If IsNull(rst) = False And helpers.populate(attrs, rst, Me, objs2Ignore) = True Then ' Fill in the fields based on record set for 'Me' form and the fieldKeys array
   Call helpers.change_control_caption("Update", Me.Create_Update)
End If

Call ToggleActive(rst("active")) ' toggle status of employee with the record set attribute 'active'



' Garbage handler does most of the work, we just need to clear it out to help move it along

Set rst = Nothing ' Clear it out
Set attrs = Nothing ' Clear it out
Set ignoreObjs = Nothing ' Clear it out
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("e_cb_name_select Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub e_cb_name_select_Change()
On Error GoTo ErrorHandler


If Me.e_cb_name_select.ListIndex = -1 Then  ' List is empty, so clear the rest of the form
    Call reset_form
    Resume ExitHandler ' Silently exit
End If

ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("Error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select




End Sub

Private Sub ToggleActive(ByVal onoff As Integer)
    If onoff = True Then ' Value must be 1
        Me.e_active_toggle.BackColor = ToggleGreen
        Call helpers.change_control_caption("De-activate", Me.Activate)
        
    ElseIf onoff = False Then ' Value must be 0
        Me.e_active_toggle.BackColor = ToggleRed
        Call helpers.change_control_caption("Activate", Me.Activate)
          
    Else ' Any other 'positive' integer
        Me.e_active_toggle.BackColor = RGB(191, 191, 191) ' Grey

    End If
End Sub

Private Function GetToggleVal() As Boolean
      If Me.e_active_toggle.BackColor = ToggleGreen Then ' Value must be 1
        GetToggleVal = True
    ElseIf Me.e_active_toggle.BackColor = ToggleRed Then ' Value must be 0
        GetToggleVal = False
    End If
End Function




