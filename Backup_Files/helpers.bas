Attribute VB_Name = "helpers"
Option Compare Database
Option Explicit ' Explicit 'typing' for variables


Public Function fill_fields_4_textboxes(ByRef fieldKeys As Variant, ByRef recordSet As DAO.recordSet, ByRef myform As Form) _
As Boolean ' fieldKeys is an array, recordSet for a select query, myForm is a form
On Error GoTo ErrorHandler ' Error handling
On Error Resume Next ' Error handling, for the For Each loop Added by - James A. 4/16/2014

 ' This function still needs tweaking for error checking in case we run into the situation where the number of textboxes is greater than the number of fieldKey array values

Dim vntControl As Variant ' vntControl can be any type of control in this sub procedure
Dim index As Integer: index = 0 ' Index for recordSet.Fields(index) iteration
Dim indexMax As Integer: indexMax = UBound(fieldKeys) ' Upper bounds index for the fieldKeys array variable


For Each vntControl In myform.Controls ' Grab myForm's controls
    If vntControl.ControlType = acTextBox And index <= indexMax Then ' If the variant control type is equal to an access TextBox
        vntControl.Value = recordSet.Fields(fieldKeys(index)) ' Then search the key in the recordSet.Fields object
        index = index + 1 ' Increment index by 1
    End If
   
Next vntControl ' This still needs to exit when we've reached the indexMax but it is buggy if we try to use 'Exit For' as we get an error...

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

Public Function populate(ByRef fieldKeys As Variant, ByRef recordSet As DAO.recordSet, _
ByRef myform As Form, Optional ByRef ignDict As Scripting.Dictionary = Nothing) As Boolean ' fieldKeys is an array, recordSet for a select query, myForm is a form

On Error GoTo ErrorHandler ' Error handling
On Error Resume Next ' Error handling, for the For Each loop Added by - James A. 4/16/2014

 ' This function still needs tweaking for error checking in case we run into the situation where the number of textboxes is greater than the number of fieldKey array values

Dim vntControl As Variant ' vntControl can be any type of control in this sub procedure
Dim index As Integer: index = 0 ' Index for recordSet.Fields(index) iteration
Dim indexMax As Integer: indexMax = UBound(fieldKeys) ' Upper bounds index for the fieldKeys array variable
Dim ignored As Boolean: ignored = False ' By default, in case ignDict is null

For Each vntControl In myform.Controls ' Grab myForm's controls
    
    If vntControl.ControlType = acTextBox And index <= indexMax Then ' If the variant control type is equal to an access TextBox
          
          If Not (ignDict Is Nothing) Then
            ignored = sender_is_dkey(vntControl.name, ignDict)
          End If
          
        If ignored = False Then ' Then go ahead and set the value
        
            vntControl.Value = recordSet.Fields(fieldKeys(index)) ' Then search the key in the recordSet.Fields object
            index = index + 1 ' Increment index by 1
        
        ElseIf ignored = True Then ' Then ignore it
        
        End If
    ElseIf vntControl.ControlType = acComboBox And index <= indexMax Then
    
        If Not (ignDict Is Nothing) Then
            ignored = sender_is_dkey(vntControl.name, ignDict)
        End If
        
        If ignored = False Then ' Then go ahead and set the value
            
            vntControl.Value = val(recordSet.Fields(fieldKeys(index)))
            index = index + 1
        ElseIf ignored = True Then ' Then ignore it
            
        End If
    End If
   
Next vntControl ' This still needs to exit when we've reached the indexMax but it is buggy if we try to use 'Exit For' as we get an error...

'Form.emp_add_text_last.Value = rst.Fields("last_name")
'Form.emp_add_text_middle.Value = rst.Fields("middle_initial")
'Form.emp_add_text_first.Value = rst.Fields("first_name")
populate = True ' Successfully filled the textboxes



' Now exit
ExitHandler:
    Exit Function
ErrorHandler:
    Select Case Err
        Case Else ' All other cases
            MsgBox ("Populate Error: " + Err.Description)
            populate = False ' Error received
            Resume ExitHandler ' Invoke Exit Handler
    End Select

End Function
Private Function sender_is_dkey(ByVal sender As String, ByRef dict As Scripting.Dictionary) As Boolean
    On Error GoTo ErrorHandler
    
    
        If Not (dict Is Nothing) And CStr(dict(sender)) = sender Then
            sender_is_dkey = True ' Sender is in so return true
        Else
            sender_is_dkey = False ' Sender is not
        End If
        
'Now exit
ExitHandler:
    Exit Function
ErrorHandler:
        MsgBox ("Sender_is_dkey Error: " + Err.Description)
        sender_is_dkey = True ' Error received
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
            MsgBox ("change_control_caption Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub


Public Sub clear_form(ByRef myform As Form)
    On Error Resume Next ' Error handling, in case control doesn't have an error property

    Dim vntControl As Variant ' vntControl can be any type of control
    For Each vntControl In myform.Controls
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


Public Function valid_dates(ByRef start_date As Date, ByRef end_date As Date) As Boolean
    
    If end_date < start_date Then
        valid_dates = False
    Else
        valid_dates = True
    End If
    
End Function


Public Sub set_rowsource(ByRef newsource As String, ByRef myform As Form, ByVal ctrl As String)
On Error GoTo ErrorHandler
    myform(ctrl).RowSource = newsource
    
ExitHandler:
    Exit Sub
    
ErrorHandler:
    Select Case Err
    Case Else
        MsgBox ("Set Rowsource Error: " + Err.Description)
        Resume ExitHandler
    End Select
End Sub
' This function will clear the combo box items for both combo boxes and for value lists
Public Sub clear_cb_items(ByRef myform As Form, ByVal ctrl As String)
On Error GoTo ErrorHandler
    Dim i As Integer
    
    If IsNull(myform) = True Then
        MsgBox ("Exiting")
        Exit Sub
    End If
    
    If myform(ctrl).RowSourceType = "Value List" Then
        For i = 1 To myform(ctrl).ListCount
            myform(ctrl).RemoveItem 0 ' Pop the top item
        Next i
    ElseIf myform(ctrl).RowSourceType = "Table/Query" Then
        'MsgBox ("This is the table/query type")
    End If
    
ExitHandler:
    Exit Sub
    
ErrorHandler:

    Select Case Err
    Case Else
        MsgBox ("Set Clear CB Items Error: " + Err.Description)
        Resume ExitHandler
    End Select
    

End Sub
