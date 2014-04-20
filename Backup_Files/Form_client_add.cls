VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_client_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' Enforce strict variable 'typing'


Private Sub Clear_Btn_Click()
    Call helpers.clear_form(Me)
    Call helpers.change_control_caption("Create", Me.Create_Update)
End Sub



Private Sub c_cb_name_select_BeforeUpdate(Cancel As Integer)
On Error GoTo ErrorHandler ' Error handling



If Me.c_cb_name_select And IsNull(Me.c_cb_name_select.Column(0)) = True Then
    Resume ExitHandler ' Silently exit and let MS Access give us a dialog error and the dropdown list
End If

Dim textboxAttrs As Variant: textboxAttrs = Array("last_name", "first_name", "middle_initial") ' Dictates the order in which the values will be placed into the textboxes
Dim client_id As Integer: client_id = Me.c_cb_name_select.Column(0) ' Retrieve the client_id
Dim query As String: query = string_helpers.assemble_query(SelectQuery, textboxAttrs, Array("clients"), "WHERE ID = " & client_id)
Dim rst As DAO.recordSet: Set rst = helpers.get_record(query, Application.currentdb)


'Dim DBS As DAO.Database, rst As DAO.recordSet
'Set DBS = Application.CurrentDb
'Set rst = DBS.OpenRecordset("SELECT last_name, middle_initial, first_name FROM clients WHERE ID = " & client_id) ' Explicit values to select



'Call fill_fields_4_textboxes(fieldKeys, rst, Me) ' Fill in the fields based on record set for 'Me' form and the fieldKeys array
If helpers.fill_fields_4_textboxes(textboxAttrs, rst, Me) = True Then ' Fill in the fields based on record set for 'Me' form and the fieldKeys array
   Call helpers.change_control_caption("Update", Me.Create_Update)
End If

Set rst = Nothing ' Clear it out
Set textboxAttrs = Nothing ' Clear it out

ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("Client Add Error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select


End Sub

Private Sub c_cb_name_select_Change()
On Error GoTo ErrorHandler ' Error handling

If (Me.c_cb_name_select) And Me.c_cb_name_select.ListIndex = -1 Then ' List is empty, so clear the rest of the form
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
            MsgBox ("Client Add Error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select

'this si the place I think
'Dim fieldKeys As Variant: fieldKeys = Array("last_name", "first_name", "middle_initial") ' Dictates the order in which the values will be placed into the textboxes
'Dim client_id As Integer: client_id = Me.c_cb_name_select.Column(0) ' Retrieve the client_id
'Dim DBS As DAO.Database, rst As DAO.recordSet
'Set DBS = Application.CurrentDb
'Set rst = DBS.OpenRecordset("SELECT last_name, middle_initial, first_name FROM clients WHERE ID = " & client_id) ' Explicit values to select


'If helpers.fill_fields_4_textboxes(fieldKeys, rst, Me) = True Then ' Fill in the fields based on record set for 'Me' form and the fieldKeys array
 '   Call helpers.change_control_caption("Update", Me.Create_Update)
'End If


'Set rst = Nothing ' Clear it out
'Set fieldKeys = Nothing ' Clear it out


End Sub

Private Sub Command16_Click() ' Click the Create/Update button
    

End Sub


Private Sub clear_form()
On Error Resume Next ' Error handling, in case control doesn't have an error property

Dim vntControl As Variant ' vntControl can be any type of control
For Each vntControl In Me.Controls
    vntControl.Value = Null
Next vntControl

End Sub

Private Sub Submit_Client()
On Error GoTo ErrorHandler ' Error handling

Dim client_id As Integer
Dim first_name As String
Dim last_name As String
Dim middle_name As String
Dim fillColTab As String
Dim intCount As Integer: intCount = 0 ' Initialize to zero


If IsNull(Me.client_add_text_first) = False Then
    first_name = Me.client_add_text_first.Value
Else
    'Call validation.Instance.Highlight(Me.emp_add_text_first)

    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.client_add_text_last) = False Then
    last_name = Me.client_add_text_last.Value
Else

    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.client_add_text_middle) = False Then
    middle_name = Me.client_add_text_middle.Value
Else
    middle_name = " " ' Set optional field to empty value
End If
' CLng(Nz(DLookup("ID", "Table1"),0))

If IsNull(Me.c_cb_name_select.Value) = False Then
    client_id = Me.c_cb_name_select.Column(0)
    'intCount = CLng(Nz(DLookup("ID", "clients", "first_name = '" & first_name & "' and last_name = '" & last_name & "' and middle_initial = '" & middle_name & "'"), 0))
    intCount = CLng(Nz(DLookup("ID", "clients", "ID = " & client_id), 0))
End If

If intCount > 0 Then ' Update client
    Dim lookupAttrs As Variant: lookupAttrs = Array("first_name ='" & first_name & "',last_name = '" & last_name & "', middle_initial = '" & middle_name & "'")
    Dim query As String: query = string_helpers.assemble_query(updateQuery, lookupAttrs, Array("clients"), "WHERE ID = " & client_id)
    currentdb.Execute query, dbFailOnError
    Set lookupAttrs = Nothing ' Clear it out
Else

'SELECT [employees].[ID], [employees].[last_name] + ",  " + [employees].[first_name] + " " + [employees].[middle_initial] FROM employees ORDER BY [last_name], [first_name], [middle_initial];


fillColTab = "INSERT INTO clients (first_name,last_name,middle_initial) VALUES ('" & first_name & "','" & last_name & "','" & middle_name & "');"
currentdb.Execute fillColTab, dbFailOnError

Me.client_add_text_last.Value = ""
Me.client_add_text_middle.Value = ""
Me.client_add_text_first.Value = ""

Me.client_add_text_first.SetFocus
MsgBox "New Client Record Successfully Created"


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



Private Sub Command16_Enter()

End Sub


Private Sub Command17_Click()

End Sub

Private Sub Command17_Enter()

End Sub

Private Sub Command18_Click()

End Sub

Private Sub Create_Update_Click()
    Call Submit_Client
End Sub

Private Sub OLEUnbound32_Updated(Code As Integer)

End Sub
