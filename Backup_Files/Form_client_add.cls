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


Private Sub Command15_Click() ' Cancel clears the form

Call clear_form

End Sub



Private Sub c_cb_name_select_BeforeUpdate(Cancel As Integer)

Dim client_id As Integer: client_id = Me.c_cb_name_select.Column(0) ' Retrieve the client_id
Dim DBS As DAO.Database, rst As DAO.recordSet
Set DBS = Application.CurrentDb
Set rst = DBS.OpenRecordset("SELECT last_name, middle_initial, first_name FROM clients WHERE ID = " & client_id) ' Explicit values to select

Dim fieldKeys As Variant ' fieldKeys will be an array
fieldKeys = Array("last_name", "first_name", "middle_initial") ' Dictates the order in which the values will be placed into the textboxes

Call fill_fields_4_textboxes(fieldKeys, rst, Me) ' Fill in the fields based on record set for 'Me' form and the fieldKeys array


Set rst = Nothing ' Clear it out
Set fieldKeys = Nothing ' Clear it out

End Sub

Private Sub c_cb_name_select_Change()

'this si the place I think

Dim client_id As Integer: client_id = Me.c_cb_name_select.Column(0) ' Retrieve the client_id
Dim DBS As DAO.Database, rst As DAO.recordSet
Set DBS = Application.CurrentDb
Set rst = DBS.OpenRecordset("SELECT last_name, middle_initial, first_name FROM clients WHERE ID = " & client_id) ' Explicit values to select

Dim fieldKeys As Variant ' fieldKeys will be an array
fieldKeys = Array("last_name", "first_name", "middle_initial") ' Dictates the order in which the values will be placed into the textboxes

Call fill_fields_4_textboxes(fieldKeys, rst, Me) ' Fill in the fields based on record set for 'Me' form and the fieldKeys array


Set rst = Nothing ' Clear it out
Set fieldKeys = Nothing ' Clear it out



End Sub

Private Sub Command16_Click() ' Click the Create/Update button
Call Submit_Client

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


Dim first_name As String
Dim last_name As String
Dim middle_name As String
Dim fillColTab As String
Dim intCount As Integer


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

intCount = CLng(Nz(DLookup("ID", "clients", "first_name = '" & first_name & "' and last_name = '" & last_name & "' and middle_initial = '" & middle_name & "'"), 0))

If intCount > 0 Then
' Do Nothing
Else

'SELECT [employees].[ID], [employees].[last_name] + ",  " + [employees].[first_name] + " " + [employees].[middle_initial] FROM employees ORDER BY [last_name], [first_name], [middle_initial];


fillColTab = "INSERT INTO clients (first_name,last_name,middle_initial) VALUES ('" & first_name & "','" & last_name & "','" & middle_name & "');"
CurrentDb.Execute fillColTab, dbFailOnError

Me.emp_add_text_last.Value = ""
Me.emp_add_text_middle.Value = ""
Me.emp_add_text_first.Value = ""

Me.emp_add_text_last.SetFocus
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





' fill_fields_4_textboxes needs to be moved into a custom library for our VBA project
Private Sub fill_fields_4_textboxes(fieldKeys As Variant, recordSet As DAO.recordSet, myForm As Form) ' fieldKeys is an array, recordSet for a select query, myForm is a form
On Error GoTo ErrorHandler ' Error handling

 ' This function still needs tweaking for error checking in case we run into the situation where the number of textboxes is greater than the number of fieldKey array values

Dim vntControl As Variant ' vntControl can be any type of control in this sub procedure
Dim index As Integer: index = 0 ' Index for recordSet.Fields(index) iteration
Dim indexMax As Integer: indexMax = UBound(fieldKeys) ' Upper bounds index for the fieldKeys array variable


For Each vntControl In myForm.Controls ' Grab myForm's controls
    If vntControl.ControlType = acTextBox And index <= indexMax Then ' If the variant control type is equal to an access TextBox
        vntControl.Value = recordSet.Fields(fieldKeys(index)) ' Then search the key in the recordSet.Fields object
        index = index + 1 ' Increment index by 1
    End If
   
Next vntControl ' This still needs to exit when we've reached the indexMax but it is buggy if we try to use 'Exit For' as we get an error...

'Form.emp_add_text_last.Value = rst.Fields("last_name")
'Form.emp_add_text_middle.Value = rst.Fields("middle_initial")
'Form.emp_add_text_first.Value = rst.Fields("first_name")





ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case Else ' All other cases
            MsgBox ("Error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select

End Sub

Private Sub OLEUnbound32_Updated(Code As Integer)

End Sub
