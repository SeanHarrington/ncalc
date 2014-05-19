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


Private Sub ClearBtn_Click()
    Call reset_form
End Sub
Private Sub reset_form()
    Call helpers.clear_form(Me)
    Call helpers.change_control_caption("Post", Me.Create_Update)
End Sub
Private Sub Create_Update_Click()
    Call submit_house
End Sub

Private Sub submit_house()
On Error GoTo ErrorHandler ' Error handling

    Dim house_id As Integer
    Dim comment As String
    Dim fillColTab As String

    Dim intCount As Integer: intCount = 0 ' Initialized to zero

' CLng(Nz(DLookup("ID", "Table1"),0))

    If IsNull(Me.ha_cb_house_select) = False Then
        house_id = Me.ha_cb_house_select.Column(0)
        intCount = CLng(Nz(DLookup("ID", "houses", "ID = " & house_id), 0))
    End If
    
    If IsNull(Me.ha_add_text_comment) = False Then
        comment = Me.ha_add_text_comment.Value
    End If

    If intCount > 0 Then ' Update house
        Dim lookupAttrs As Variant: lookupAttrs = Array("comments = " + """" & comment & """")  ' To prevent injection and database errors, surround variable with double quotes
        Dim query As String: query = string_helpers.assemble_query(updateQuery, lookupAttrs, Array("houses"), "WHERE ID = " & house_id)
        Call helpers.exec_query(query, Application.CurrentDb)
    
        Set lookupAttrs = Nothing ' Clear it out
    Else ' Submit/Create new house
        Dim answer As VbMsgBoxResult: answer = vbYes ' By Default, yes
        Dim client_id As Integer
        Dim rst As DAO.recordSet
        
        If IsNull(Me.ha_cb_unassignedc_select) = False Then
            client_id = Me.ha_cb_unassignedc_select.Column(0)
        Else
            MsgBox ("Please select an unassigned client to start a new house")
            Exit Sub
        End If
        
        
        If IsNull(Me.ha_add_text_comment) = True Then
            answer = MsgBox("Go ahead and create a house with no comment?", vbQuestion + vbYesNo)
        End If
        
        If answer = vbNo Then
            Exit Sub
        End If ' Otherwise continue
        
        fillColTab = "INSERT INTO houses (comments) VALUES ('" & comment & "');" ' Create the new house
        
        Call helpers.exec_query(fillColTab, Application.CurrentDb) ' Execute
        'house_id = @@identity
        Set rst = helpers.get_record("SELECT @@Identity;", Application.CurrentDb)
        house_id = CLng(rst.Fields(0)) ' Get the last inserted row's ID
        fillColTab = "UPDATE [clients] SET [house_id] = " & house_id & " WHERE [ID] = " & client_id

        Call helpers.exec_query(fillColTab, Application.CurrentDb) ' Connect the client to the new house
        
        Call update_houses_list
        Call update_assigned_list
        Call update_unassigned_list
        'Call reset_form
        MsgBox "New House Record Successfully Created"
    End If


ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            
        Case Else ' All other cases
            MsgBox ("House Add Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
    
End Sub

Private Sub Disband_House_Click()
On Error GoTo ErrorHandler ' Error handling
    Dim house_id As Long
    Dim intCount As Long
    
    If IsNull(Me.ha_cb_house_select) = False Then
        house_id = Me.ha_cb_house_select.Column(0) ' Get primary key
        intCount = CLng(Nz(DLookup("ID", "houses", "ID = " & house_id), 0))
    Else
        MsgBox ("Please select a house to disband in the combobox list")
        Exit Sub
    End If
    
    ' We will delete the house and remove all relationships to this house
    
    If intCount > 0 Then ' Then we are sure this record exists
        Dim query As String: query = "UPDATE clients SET clients.house_id = " & 0 & " WHERE clients.house_id = " & house_id ' Set to default value
        
        Call helpers.exec_query(query, Application.CurrentDb) ' Remove all relationships to this house before deleting the house
        
        query = "DELETE FROM houses WHERE houses.id = " & house_id
        
        Call helpers.exec_query(query, Application.CurrentDb) ' Delete the house from the houses table
        
        
        MsgBox ("House #" & house_id & " successfully disbanded!")
        
        Call update_houses_list
        Call update_assigned_list
        Call update_unassigned_list
        Call reset_form ' Just reset the whole thing
       
        
    End If
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("Client House Disband Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_Open(Cancel As Integer)
    Call update_unassigned ' Set the unassigned rowsource
    Call update_houses_list ' Update the houses list
    Call update_unassigned_list ' Update both the unassigned clients list
End Sub
Private Sub update_unassigned_list()
     Me![ha_cb_unassignedc_select].Requery ' Update unassigned
End Sub

Private Sub update_assigned_list()
    Me![ha_cb_assignedc_select].Requery
End Sub

Private Sub update_houses_list()
Dim prepString As String
Dim valIsNull As Boolean: valIsNull = True ' By default this value is null
Dim rs1 As recordSet
Set rs1 = CurrentDb.OpenRecordset("SELECT ID FROM houses;") ' get all the houses '

If IsNull(Me.ha_cb_house_select.Value) = False Then ' Store the old value index
    valIsNull = False ' Not Null
End If

Call helpers.clear_cb_items(Me, "ha_cb_house_select") ' Clear/reset it


rs1.MoveFirst


Do While Not rs1.EOF

    Dim rs2 As recordSet
    prepString = "" ' reset stringbuilder '
    Set rs2 = CurrentDb.OpenRecordset("SELECT first_name FROM clients WHERE house_id = " & rs1![ID] & ";") ' get the clients with this house # '
    If rs2.RecordCount > 0 Then ' if there are any clients then proceed '
    
        Do While Not rs2.EOF
            prepString = prepString + rs2![first_name] + " " ' build the string '
            rs2.MoveNext
        Loop
        
    
    Dim strItem As String
    strItem = rs1![ID] & ";" & prepString
    Me.ha_cb_house_select.AddItem Item:=strItem ' Populate the list
     
    
    ' Column(0) = rs1![ID] ' set cb[0] to ID '
    ' Me.ha_cb_house_select.Column(1) = prepString ' set cb[1] to built string '
   
 End If
   
   rs1.MoveNext
Loop



If valIsNull = False Then
   Me.ha_cb_house_select.Value = Me.ha_cb_house_select.OldValue ' Better way of retrieving the oldvalue via Microsoft Access
End If

Me.ha_cb_house_select.SetFocus ' This lets us add a visual emphasis on the combo box list



End Sub

Private Sub ha_cb_house_select_BeforeUpdate(Cancel As Integer)
On Error GoTo ErrorHandler ' Error handling


If Me.ha_cb_house_select And IsNull(Me.ha_cb_house_select.Column(0)) = True Then ' Invalid primary key
    Resume ExitHandler ' Silently exit and let MS Access give us a dialog error and the dropdown list
End If

Dim textboxAttrs As Variant: textboxAttrs = Array("comments") ' Dictates the order in which the values will be placed into the textboxes
Dim house_id As Integer: house_id = Me.ha_cb_house_select.Column(0) ' Retrieve the house_id
Dim query As String: query = string_helpers.assemble_query(SelectQuery, textboxAttrs, Array("houses"), "WHERE ID = " & house_id)
Dim rst As DAO.recordSet: Set rst = helpers.get_record(query, Application.CurrentDb)
Dim ignoreObjs As Scripting.Dictionary: Set ignoreObjs = CreateObject("Scripting.Dictionary")



Call ignoreObjs.Add("ha_cb_assignedc_select", "ha_cb_assignedc_select") ' Ignore cb list
Call ignoreObjs.Add("ha_cb_house_select", "ha_cb_house_select") ' Ignore cb list
Call ignoreObjs.Add("ha_cb_unassignedc_select", "ha_cb_unassignedc_select") ' Ignore cb list



If helpers.populate(textboxAttrs, rst, Me, ignoreObjs) = True Then ' Fill in the fields based on record set for 'Me' form and the fieldKeys array
   Call helpers.change_control_caption("Update", Me.Create_Update)
End If

Call update_assigned_by(house_id) ' Update the combo boxes by this primary key in their SQL queries
Call update_unassigned_list ' Update the unassigned list


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
If Me.ha_cb_house_select.ListIndex = -1 Then ' List is empty, so clear the rest of the form
    Call helpers.clear_form(Me)
    Call helpers.set_rowsource("", Me, "ha_cb_assignedc_select")
    Call update_assigned_list
    Call helpers.change_control_caption("Post", Me.Create_Update) ' Reset
End If

ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("House Select Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub


Private Sub update_assigned_by(ByVal pkey As Integer)
    Dim query As String: query = "SELECT clients.ID, clients.last_name + ', ' + clients.first_name + ' ' +  clients.middle_initial FROM clients where clients.house_id = " & pkey & ";"
    Call helpers.set_rowsource(query, Me, "ha_cb_assignedc_select")
End Sub

Private Sub update_unassigned()
    Dim row_query As String: row_query = "SELECT clients.ID, clients.last_name + ', ' + clients.first_name + ' ' + clients.middle_initial FROM clients WHERE clients.house_id = 0"
    Call helpers.set_rowsource(row_query, Me, "ha_cb_unassignedc_select")

End Sub

Private Sub Meeting_Create_Click()
   
    If IsNull(Me.ha_cb_house_select.Column(0)) = False Then
        
        Dim house_id As Long: house_id = Me.ha_cb_house_select.Column(0) ' Get the house id so we can pass it
        DoCmd.Close acForm, "meeting_create" ' Reload it, it may already be open.  This is an expensive operation...
        DoCmd.OpenForm "meeting_create", , , , , , "mc_cb_house_select|" & house_id ' We need to pass in our first parameter (the target) and our value seperated by '|'
       
    Else
        MsgBox ("Please select a valid employee before attempting to go to the meeting create form")
    End If
End Sub

Private Sub Remove_From_House_Click()
    On Error GoTo ErrorHandler ' Error handling
    Dim house_id As Integer
    Dim intCount As Long
    Dim client_id As Integer
    
    
    If IsNull(Me.ha_cb_house_select) = False And IsNull(Me.ha_cb_house_select.Column(0)) = False Then
        house_id = Me.ha_cb_house_select.Column(0) ' Get primary key
        intCount = CLng(Nz(DLookup("ID", "houses", "ID = " & house_id), 0))
    Else
        MsgBox ("Please select a house # to remove a client from")
        Exit Sub
    End If
    
    If IsNull(Me.ha_cb_assignedc_select) = False And IsNull(Me.ha_cb_assignedc_select.Column(0)) = False Then
        client_id = Me.ha_cb_assignedc_select.Column(0)
    Else
        MsgBox ("Please select a valid assigned client before attempting to remove from a house")
        Exit Sub
    End If
    
    If intCount > 0 Then ' Go ahead and update the records
        Dim query As String: query = "UPDATE clients SET clients.house_id = " & 0 & " WHERE clients.ID = " & client_id
        
        Call helpers.exec_query(query, Application.CurrentDb)
        
        MsgBox (" " & Me.ha_cb_assignedc_select.Column(1) & " successfully removed from house #" & house_id) ' Confirmation message
        Call update_houses_list ' Update houses list
        Call update_assigned_list
        Call update_unassigned_list
        
    End If
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("House Remove Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub Assign_To_House_Click()
    On Error GoTo ErrorHandler ' Error handling
    Dim intCount As Long
    Dim house_id As Integer
    Dim client_id As Integer
    
    If IsNull(Me.ha_cb_house_select) = False And IsNull(Me.ha_cb_house_select.Column(0)) = False Then
        house_id = Me.ha_cb_house_select.Column(0) ' Get the house_id
        intCount = CLng(Nz(DLookup("ID", "houses", "ID = " & house_id), 0))
    Else
        MsgBox ("Please select a house # to assign a client to")
        Exit Sub
    End If
    
    If IsNull(Me.ha_cb_unassignedc_select) = False Then
        client_id = Me.ha_cb_unassignedc_select.Column(0) ' Get the client_id
    Else
        MsgBox ("Please select a valid unassigned client before assigning to a house")
        Exit Sub
    End If
    
    If intCount > 0 Then ' Go ahead and update the records
        Dim query As String: query = "UPDATE clients SET clients.house_id = " & house_id & " WHERE clients.ID = " & client_id
        Call helpers.exec_query(query, Application.CurrentDb)
        MsgBox (" " & Me.ha_cb_unassignedc_select.Column(1) & " successfully assigned to house #" & house_id) ' Confirmation message
        Call update_houses_list ' Update houses list
        Me![ha_cb_unassignedc_select].Requery ' Update the unassigned combo box list
        Me![ha_cb_assignedc_select].Requery ' Update the assigned combo box list, I think this might be unnecessary and it does it automatically when something is added...
        
    End If
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            ' Silently catch an exit without an error
        Case Else ' All other cases
            MsgBox ("House Assign Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub
