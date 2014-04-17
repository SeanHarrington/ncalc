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
' CLng(Nz(DLookup("ID", "Table1"),0))

    If IsNull(Me.e_cb_name_select.Value) = False Then
        'Me.e_cb_name_select.Column(0) ' Employee_id
         emp_id = Me.e_cb_name_select.Column(0)
        'intCount = CLng(Nz(DLookup("ID", "employees", "first_name = '" & first_name & "' and last_name = '" & last_name & "' and middle_initial = '" & middle_name & "'"), 0))
        intCount = CLng(Nz(DLookup("ID", "employees", "ID = " & emp_id), 0))
    End If
    
If intCount > 0 Then ' Update employee
    
    Dim lookupAttrs As Variant: lookupAttrs = Array("first_name ='" & first_name & "',last_name = '" & last_name & "', middle_initial = '" & middle_name & "'")
    Dim query As String: query = string_helpers.assemble_query(updateQuery, lookupAttrs, Array("employees"), "WHERE ID = " & emp_id)
    CurrentDb.Execute query, dbFailOnError
    Set lookupAttrs = Nothing ' Clear it out
Else ' Submit/Create new employee




'SELECT [employees].[ID], [employees].[last_name] + ",  " + [employees].[first_name] + " " + [employees].[middle_initial] FROM employees ORDER BY [last_name], [first_name], [middle_initial];


fillColTab = "INSERT INTO employees (first_name,last_name,middle_initial) VALUES ('" & first_name & "','" & last_name & "','" & middle_name & "');"
CurrentDb.Execute fillColTab, dbFailOnError

Me.emp_add_text_last.Value = ""
Me.emp_add_text_middle.Value = ""
Me.emp_add_text_first.Value = ""

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
            MsgBox ("Error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
    
End Sub


Private Sub Command18_Click()

End Sub

Private Sub Create_Update_Click()
  
    Call Submit_Employee
    
End Sub

Private Sub e_cb_name_select_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub e_cb_name_select_Change()
'this si the place I think
Dim textboxAttrs As Variant: textboxAttrs = Array("last_name", "first_name", "middle_initial") ' Dictates the order in which the values will be placed into the textboxes
Dim employee_id As Integer: employee_id = Me.e_cb_name_select.Column(0) ' Retrieve the employee_id
Dim query As String: query = string_helpers.assemble_query(SelectQuery, textboxAttrs, Array("employees"), "WHERE ID = " & employee_id)
Dim rst As DAO.recordSet: Set rst = helpers.get_record(query, Application.CurrentDb)
'Set DBS = Application.CurrentDb
'Set rst = DBS.OpenRecordset("SELECT last_name, middle_initial, first_name FROM employees WHERE ID = " & employee_id) ' Explicit values to select
'Set rst = helpers.get_record(query, Application.CurrentDb)

If helpers.fill_fields_4_textboxes(textboxAttrs, rst, Me) = True Then ' Fill in the fields based on record set for 'Me' form and the fieldKeys array
   Call helpers.change_control_caption("Update", Me.Create_Update)
End If


Set rst = Nothing ' Clear it out
Set textboxAttrs = Nothing ' Clear it out



End Sub




