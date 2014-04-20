Attribute VB_Name = "helpers"
Option Compare Database
Option Explicit ' Explicit 'typing' for variables

Public Function fill_fields_4_textboxes(ByRef fieldKeys As Variant, ByRef recordSet As DAO.recordSet, ByRef myForm As Form) _
As Boolean ' fieldKeys is an array, recordSet for a select query, myForm is a form
On Error GoTo ErrorHandler ' Error handling
On Error Resume Next ' Error handling, for the For Each loop Added by - James A. 4/16/2014

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
fill_fields_4_textboxes = True ' Successfully filled the textboxes



' Now exit
ExitHandler:
    Exit Function
ErrorHandler:
    Select Case Err
        Case Else ' All other cases
            MsgBox ("Error Received: " + Err.Description)
            fill_fields_4_textboxes = False ' Error received
            Resume ExitHandler ' Invoke Exit Handler
    End Select

End Function

Public Function fill_fields(ByRef fieldKeys As Variant, ByRef recordSet As DAO.recordSet, _
ByRef myForm As Form, Optional ByRef ignCol As Scripting.Dictionary = Null) As Boolean ' fieldKeys is an array, recordSet for a select query, myForm is a form

On Error GoTo ErrorHandler ' Error handling
On Error Resume Next ' Error handling, for the For Each loop Added by - James A. 4/16/2014

 ' This function still needs tweaking for error checking in case we run into the situation where the number of textboxes is greater than the number of fieldKey array values

Dim vntControl As Variant ' vntControl can be any type of control in this sub procedure
Dim index As Integer: index = 0 ' Index for recordSet.Fields(index) iteration
Dim indexMax As Integer: indexMax = UBound(fieldKeys) ' Upper bounds index for the fieldKeys array variable
Dim ignored As Boolean

For Each vntControl In myForm.Controls ' Grab myForm's controls
    
    If vntControl.ControlType = acTextBox And index <= indexMax Then ' If the variant control type is equal to an access TextBox
          ignored = sender_contains_key(vntControl.Name, CStr(index), ignCol)
        If ignored = False Then
            'MsgBox ("Setting acTextBox to: " & recordSet.Fields(fieldKeys(index)))
            vntControl.Value = recordSet.Fields(fieldKeys(index)) ' Then search the key in the recordSet.Fields object
            index = index + 1 ' Increment index by 1
        ElseIf ignored = True Then
            'vntControl.Value = recordSet.Fields(fieldKeys(index)) ' Then search the key in the recordSet.Fields object
            'index = index + 1 ' Increment index by 1
        End If
    ElseIf vntControl.ControlType = acComboBox And index <= indexMax Then
        ignored = sender_contains_key(vntControl.Name, CStr(index), ignCol)
        'MsgBox ("Sender is: " + vntControl.Name + " and ignored is" & ignored)
        If ignored = False Then
            
            
           'MsgBox ("Setting combobox to: " & val(recordSet.Fields(fieldKeys(index))))
            
            vntControl.Value = val(recordSet.Fields(fieldKeys(index)))
            
            
            index = index + 1
        ElseIf ignored = True Then
            'vntControl.Value = val(recordSet.Fields(fieldKeys(index)))
            'index = index + 1
        End If
    End If
   
Next vntControl ' This still needs to exit when we've reached the indexMax but it is buggy if we try to use 'Exit For' as we get an error...

'Form.emp_add_text_last.Value = rst.Fields("last_name")
'Form.emp_add_text_middle.Value = rst.Fields("middle_initial")
'Form.emp_add_text_first.Value = rst.Fields("first_name")
fill_fields = True ' Successfully filled the textboxes



' Now exit
ExitHandler:
    Exit Function
ErrorHandler:
    Select Case Err
        Case Else ' All other cases
            MsgBox ("Fill Fields Error: " + Err.Description)
            fill_fields = False ' Error received
            Resume ExitHandler ' Invoke Exit Handler
    End Select

End Function
Private Function sender_contains_key(ByVal sender As String, ByVal key As String, _
ByRef collect As Scripting.Dictionary) As Boolean
    On Error GoTo ErrorHandler
    Dim obj As Variant
    
    
        If CStr(collect(key)) = sender Then
            sender_contains_key = True ' Sender is in so return true
        Else
            sender_contains_key = False ' Sender is not
        End If
        
        
        
    'sender_contains_key = False
'Now exit
ExitHandler:
    Exit Function
ErrorHandler:
        MsgBox ("Sender_contains_key Error: " + Err.Description)
        sender_contains_key = True ' Error received
        Resume ExitHandler ' Invoke Exit Handler
End Function
Private Function contains_key(ByVal key As Variant, ByRef hash As Collection) As Boolean
    On Error GoTo ErrorHandler
    Dim obj As Variant
    
    obj = hash(key)
    contains_key = False
    
'Now exit
ExitHandler:
    Exit Function
ErrorHandler:
        contains_key = True ' Error received
        Resume ExitHandler ' Invoke Exit Handler
End Function


Public Function get_record(ByRef query As String, ByRef curr_db As DAO.Database) As DAO.recordSet ' Return recordset
    On Error GoTo ErrorHandler ' Error handling
    
    
    Set get_record = curr_db.OpenRecordset(query) ' Return record set

    
ExitHandler:
    Exit Function
ErrorHandler:
    Select Case Err
        Case Else ' All Error cases not accounted for
            MsgBox ("Error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Function

Public Sub change_control_caption(ByVal newCaption As String, ByRef ctlVariant As Variant) ' Support for any control but dangerous
    On Error GoTo ErrorHandler ' Error handling
    ctlVariant.Caption = newCaption
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case Else ' All Error cases not accounted for
            MsgBox ("Error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub


Public Sub clear_form(ByRef myForm As Form)
    On Error Resume Next ' Error handling, in case control doesn't have an error property

    Dim vntControl As Variant ' vntControl can be any type of control
    For Each vntControl In myForm.Controls
        vntControl.Value = Null
    Next vntControl

End Sub

'exec_query serves as an abstraction or a 'wrapper' to allow for better error handling
Public Sub exec_query(ByRef query As String, ByRef curr_db As DAO.Database)
    On Error GoTo ErrorHandler
    
    curr_db.Execute query, dbFailOnError
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case Else ' All Error cases not accounted for
            MsgBox ("Execute Query Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub
