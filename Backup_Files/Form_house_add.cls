VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_house_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' Enforce strict typing and declarations

Private Sub Create_Update_Click()
    Call submit_house
End Sub

Private Sub submit_house()
On Error GoTo ErrorHandler ' Error handling

Dim house_id As Integer
Dim house_name As String
Dim address As String
Dim city As String
Dim state As String
Dim zip As String
Dim phone As String: phone = "default"
Dim fillColTab As String

Dim intCount As Integer: intCount = 0 ' Initialized to zero


If IsNull(Me.house_add_text_name) = False Then
    house_name = Me.house_add_text_name.Value
Else
    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.house_add_text_addr) = False Then
    address = Me.house_add_text_addr.Value
Else
    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.house_add_text_city) = False Then
    city = Me.house_add_text_city.Value
Else
    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.house_add_text_state) = False Then
    state = Me.house_add_text_state.Value
Else
    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.house_add_text_zip) = False Then
    zip = Me.house_add_text_zip
Else
    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If


' CLng(Nz(DLookup("ID", "Table1"),0))

    If IsNull(Me.ha_cb_select) = False Then
       
         house_id = Me.ha_cb_select.Column(0)
        
        intCount = CLng(Nz(DLookup("ID", "houses", "ID = " & house_id), 0))
    End If
    
If intCount > 0 Then ' Update house
    Dim lookupAttrs As Variant: lookupAttrs = Array("house_name ='" & house_name & "',address = '" & address & "',city = '" & city & "',state = '" & state & "', zip = '" & zip & "'")
    Dim query As String: query = string_helpers.assemble_query(updateQuery, lookupAttrs, Array("houses"), "WHERE ID = " & house_id)
    CurrentDb.Execute query, dbFailOnError
    
    Set lookupAttrs = Nothing ' Clear it out
Else ' Submit/Create new house





fillColTab = "INSERT INTO houses (address,city, state, zip, phone, house_name) VALUES ('" & address & "','" & city & "','" & state & "', '" & zip & "', '" & phone & "', '" & house_name & "');"
CurrentDb.Execute fillColTab, dbFailOnError

Me.house_add_text_addr.Value = ""
Me.house_add_text_city.Value = ""
Me.house_add_text_state.Value = ""
Me.house_add_text_zip.Value = ""
Me.house_add_text_name.Value = ""

Me.house_add_text_name.SetFocus
MsgBox "New House Record Successfully Created"


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


Private Sub ha_cb_select_BeforeUpdate(Cancel As Integer) ' Renamed, see other sub procedure below


End Sub

Private Sub ha_cb_select_Change() ' Renamed, see other sub procedure below

End Sub

Private Sub ha_cb_house_select_BeforeUpdate(Cancel As Integer)
On Error GoTo ErrorHandler ' Error handling


If Me.ha_cb_house_select And IsNull(Me.ha_cb_house_select.Column(0)) = True Then ' Invalid primary key
    Resume ExitHandler ' Silently exit and let MS Access give us a dialog error and the dropdown list
End If

Dim textboxAttrs As Variant: textboxAttrs = Array("name", "address", "city", "state", "zip") ' Dictates the order in which the values will be placed into the textboxes
Dim house_id As Integer: house_id = Me.ha_cb_house_select.Column(0) ' Retrieve the house_id
Dim query As String: query = string_helpers.assemble_query(SelectQuery, textboxAttrs, Array("houses"), "WHERE ID = " & house_id)
Dim rst As DAO.recordSet: Set rst = helpers.get_record(query, Application.CurrentDb)

'Call fill_fields_4_textboxes(fieldKeys, rst, Me) ' Fill in the fields based on record set for 'Me' form and the fieldKeys array
If helpers.populate(textboxAttrs, rst, Me) = True Then ' Fill in the fields based on record set for 'Me' form and the fieldKeys array
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
            MsgBox ("Client House Select Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub ha_cb_house_select_Change()
On Error GoTo ErrorHandler ' Error handling
If (Me.ha_cb_house_select) And Me.ha_cb_house_select.ListIndex = -1 Then ' List is empty, so clear the rest of the form
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
            MsgBox ("Client House Select Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub
