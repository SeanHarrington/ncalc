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



Private Sub ClearBtn_Click()
    Call helpers.clear_form(Me)
    Call helpers.change_control_caption("Create", Me.Create_Update)
End Sub



Private Sub Submit_Employee()
On Error GoTo ErrorHandler ' Error handling

Dim emp_id As Integer
Dim first_name As String
Dim last_name As String
Dim middle_name As String
Dim pay_rate As Currency
Dim sick_leave_rate As Integer
Dim title As String
Dim phone As String
Dim regional_center As Integer
Dim comment As String
Dim yes_no As Boolean: yes_no = True
Dim fillColTab As String

Dim intCount As Integer: intCount = 0 ' Initialized to zero


If IsNull(Me.emp_add_text_first) = False Then
    first_name = Me.emp_add_text_first.Value
Else
    'Call validation.Instance.Highlight(Me.emp_add_text_first)

    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.emp_add_text_last) = False Then
    last_name = Me.emp_add_text_last.Value
Else

    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.emp_add_text_middle) = False Then
    middle_name = Me.emp_add_text_middle.Value
Else
    middle_name = " " ' Set optional field to empty value
End If

If IsNull(Me.e_cb_region_select) = False And Me.e_cb_region_select.ListIndex <> -1 Then
    regional_center = Me.e_cb_region_select.Column(0) ' Grab the text value
Else
    regional_center = 0 ' Optional-default
End If

If IsNull(Me.emp_add_text_payrate) = False Then
    pay_rate = Format(val(Me.emp_add_text_payrate.Value), "currency")
    'MsgBox (pay_rate)
Else
    pay_rate = Format(0#, "currency")    ' Optional
    'MsgBox (pay_rate)
End If

If IsNull(Me.emp_add_text_sickrate) = False Then
    sick_leave_rate = Me.emp_add_text_sickrate.Value
Else
    sick_leave_rate = 0 ' Optional
End If

If IsNull(Me.emp_add_text_title) = False Then
    title = Me.emp_add_text_title.Value
Else
    title = " " ' Optional
End If

If IsNull(Me.emp_add_text_phone) = False Then
    phone = Me.emp_add_text_phone.Value
Else
    phone = " " ' Optional
End If

If IsNull(Me.emp_add_text_comment) = False Then
    comment = Me.emp_add_text_comment.Value
Else
    comment = " " ' Optional
End If
' CLng(Nz(DLookup("ID", "Table1"),0))

    If IsNull(Me.e_cb_name_select.Value) = False Then
        'Me.e_cb_name_select.Column(0) ' Employee_id
         emp_id = Me.e_cb_name_select.Column(0)
        'intCount = CLng(Nz(DLookup("ID", "employees", "first_name = '" & first_name & "' and last_name = '" & last_name & "' and middle_initial = '" & middle_name & "'"), 0))
        intCount = CLng(Nz(DLookup("ID", "employees", "ID = " & emp_id), 0))
    End If
    
  
    
If intCount > 0 Then ' Update employee
    Dim lookupAttrs As Variant: lookupAttrs = Array("first_name = '" & first_name & "'", "last_name = '" & last_name & "'" _
    , "middle_initial = '" & middle_name & "'", "payrate = '" & pay_rate & "'", "sick_hour_rate = " & sick_leave_rate _
    , "title = '" & title & "'", "phone = '" & phone & "'", "regional_center = " & regional_center _
    , "comments = '" & comment & "'")
    
    Dim query As String: query = string_helpers.assemble_query(updateQuery, lookupAttrs, Array("employees"), "WHERE ID = " & emp_id)
    'currentdb.Execute query, dbFailOnError
    Call helpers.exec_query(query, Application.currentdb)
        
    Set lookupAttrs = Nothing ' Clear it out
Else ' Submit/Create new employee

'SELECT [employees].[ID], [employees].[last_name] + ",  " + [employees].[first_name] + " " + [employees].[middle_initial] FROM employees ORDER BY [last_name], [first_name], [middle_initial];


'fillColTab = "INSERT INTO employees (first_name,last_name,middle_initial) VALUES ('" & first_name & "','" & last_name & "','" & middle_name & "');"

Dim insertAttrs As Variant: insertAttrs = Array("'" & first_name & "'", "'" & last_name _
    & "'", "'" & middle_name & "'", "" & yes_no, "" & pay_rate _
    , "" & sick_leave_rate, "'" & title & "'", "'" & phone & "'" _
    , "" & regional_center, "'" & comment & "'")
    
Dim attrs As Variant: attrs = Array("first_name", "last_name", "middle_initial", "active", "payrate", "sick_hour_rate" _
, "title", "phone", "regional_center", "comments")

Dim insertVals As String: insertVals = string_helpers.join_array_to_str(insertAttrs, ",")
MsgBox (insertVals)

fillColTab = string_helpers.assemble_query(InsertQuery, attrs, Array("employees"), , insertVals)
MsgBox (fillColTab)

currentdb.Execute fillColTab, dbFailOnError

'Me.emp_add_text_last.Value = ""
'Me.emp_add_text_middle.Value = ""
'Me.emp_add_text_first.Value = ""
Set insertAttrs = Nothing
Set attrs = Nothing


Call helpers.clear_form(Me)

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


Private Sub Command18_Click()

End Sub

Private Sub Create_Update_Click()
  
    Call Submit_Employee
    
End Sub

Private Sub e_btn_ihss_Click()
On Error GoTo ErrorHandler
    
    If Me.e_cb_name_select And IsNull(Me.e_cb_name_select.Column(0)) = False Then
       Dim employee_id As Integer: employee_id = Me.e_cb_name_select.Column(0)
       
       'MsgBox ("Sending id = '" & Me.e_cb_name_select.Column(0) & "'")
       DoCmd.Close acForm, "ihss_eligibility" ' Reload it, for now
       DoCmd.OpenForm "ihss_eligibility", , , , , , "ihss_cb_emp_select|" & Me.e_cb_name_select.Column(0)
       
    Else
        Dim answer As VbMsgBoxResult
        answer = MsgBox("Continue without a selected employee?", vbQuestion + vbYesNo)
        If answer = vbYes Then
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

Private Sub e_cb_name_select_AfterUpdate()
    


End Sub

Private Sub e_cb_name_select_BeforeUpdate(Cancel As Integer)
On Error GoTo ErrorHandler

If Me.e_cb_name_select And IsNull(Me.e_cb_name_select.Column(0)) = True Then ' If ID doesn't exist then we have an invalid employee
    Resume ExitHandler ' Silently exit and let MS Access give us a dialog error and the dropdown list
End If


Dim ignoreObjs As Object
Set ignoreObjs = CreateObject("Scripting.Dictionary")
Call ignoreObjs.Add("0", "e_cb_name_select")

Dim attrs As Variant: attrs = Array("last_name", "first_name", "middle_initial" _
, "regional_center", "phone", "sick_hour_rate", "payrate", "title", "comments") ' Dictates the order to fill the boxes in







Dim employee_id As Integer: employee_id = Me.e_cb_name_select.Column(0) ' Retrieve the employee_id
Dim query As String: query = string_helpers.assemble_query(SelectQuery, attrs, Array("employees"), "WHERE ID = " & employee_id)
Dim rst As DAO.recordSet: Set rst = helpers.get_record(query, Application.currentdb)
'Set DBS = Application.CurrentDb
'Set rst = DBS.OpenRecordset("SELECT last_name, middle_initial, first_name FROM employees WHERE ID = " & employee_id) ' Explicit values to select
'Set rst = helpers.get_record(query, Application.CurrentDb)

If IsNull(rst) = False And helpers.fill_fields(attrs, rst, Me, ignoreObjs) = True Then ' Fill in the fields based on record set for 'Me' form and the fieldKeys array
   Call helpers.change_control_caption("Update", Me.Create_Update)
End If


Set rst = Nothing ' Clear it out
Set attrs = Nothing ' Clear it out
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

Private Sub e_cb_name_select_Change()
On Error GoTo ErrorHandler


If (Me.e_cb_name_select) And Me.e_cb_name_select.ListIndex = -1 Then ' List is empty, so clear the rest of the form
    Call helpers.clear_form(Me)
    Call helpers.change_control_caption("Create", Me.Create_Update) ' Reset
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




