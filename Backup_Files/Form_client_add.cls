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


Private Sub Activate_Click()
      On Error GoTo ErrorHandler
    
    If Me.c_cb_name_select And IsNull(Me.c_cb_name_select.Column(0)) = False Then
        'Dim yes_no As Boolean
        Dim active As Boolean: active = Not (GetToggleVal()) ' Flip it
        Dim client_id As Integer: client_id = Me.c_cb_name_select.Column(0)
        Dim intCount As Integer: intCount = CLng(Nz(DLookup("ID", "clients", "ID = " & client_id), 0))
        Dim query As String: query = "update clients set active = " & active
        query = query + " where ID = " & client_id
        Call helpers.exec_query(query, Application.CurrentDb)
        Call ToggleActive(active)
    Else
        MsgBox ("Please select a valid client to activate")
    End If
    
    
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Without Error exit; our custom error handling
            
        Case Else ' All other cases
            MsgBox ("Client Activate Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
End Sub

Private Sub c_btn_ihss_Click()
On Error GoTo ErrorHandler
    
    If Me.c_cb_name_select And IsNull(Me.c_cb_name_select.Column(0)) = False Then
       Dim client_id As Integer: client_id = Me.c_cb_name_select.Column(0)
       
       DoCmd.Close acForm, "ihss_eligibility" ' Reload it, for now
       DoCmd.OpenForm "ihss_eligibility", , , , , , "ihss_cb_client_select|" & Me.c_cb_name_select.Column(0)
       
    Else
        Dim Answer As VbMsgBoxResult
        Answer = MsgBox("Continue without selecting a client?", vbQuestion + vbYesNo)
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

Private Sub Clear_Btn_Click()
   
   Call reset_form
   
End Sub

Private Sub reset_form()

    Call helpers.clear_form(Me)
    Call helpers.change_control_caption("Create", Me.Create_Update)
    Call ToggleActive(2) ' Any value other than 0 or 1 will trigger total reset
End Sub

Private Sub c_cb_name_select_BeforeUpdate(Cancel As Integer)
On Error GoTo ErrorHandler ' Error handling


If Me.c_cb_name_select And IsNull(Me.c_cb_name_select.Column(0)) = True Then
    Resume ExitHandler ' Silently exit and let MS Access give us a dialog error and the dropdown list
End If

'  Added '0' AS monthly_ihss_min to the query which we will not be derived from the database record but later transposed as a seperate value in the extracted recordset
' Easiest way to do it - James Alaniz 1/12/15

Dim sqlAttrs As Variant: sqlAttrs = Array("last_name", "first_name", "middle_initial", "authorized_regional_center_id" _
, "monthly_ihss", "regional_center", "comments", "billable_hours", "active")  ' Dictates the order in which the values will be placed into the textboxes
Dim txtBoxAttrs As Variant: txtBoxAttrs = Array("last_name", "first_name", "middle_initial", "authorized_regional_center_id" _
, "monthly_ihss", "regional_center", "comments", "billable_hours", "monthly_ihss_min", "active")  ' Dict

Dim client_id As Integer: client_id = Me.c_cb_name_select.Column(0) ' Retrieve the client_id
Dim query As String: query = string_helpers.assemble_query(SelectQuery, sqlAttrs, Array("clients"), "WHERE ID = " & client_id)
Dim objs2Ignore As Scripting.dictionary: Set objs2Ignore = CreateObject("Scripting.Dictionary")

Dim rst As DAO.recordSet: Set rst = helpers.get_record(query, Application.CurrentDb) ' Retrieve record from database


Call objs2Ignore.Add("c_cb_name_select", "c_cb_name_select") ' Ignore the object/element on the GUI by it's given name

Dim IHSS_HOURS As Double
Dim IHSS_MINS As Single
Dim IHSS_CONV As Single
IHSS_HOURS = rst("monthly_ihss")
IHSS_CONV = IHSS_HOURS - Int([IHSS_HOURS])
IHSS_MINS = Math.Round(IHSS_CONV * 60)
IHSS_HOURS = Int([IHSS_HOURS])

Dim dictionarySet As Scripting.dictionary: Set dictionarySet = CreateObject("Scripting.Dictionary")
Call dictionarySet.Add("last_name", rst("last_name"))
Call dictionarySet.Add("first_name", rst("first_name"))
Call dictionarySet.Add("middle_initial", rst("middle_initial"))
Call dictionarySet.Add("authorized_regional_center_id", rst("authorized_regional_center_id"))
Call dictionarySet.Add("monthly_ihss", IHSS_HOURS)
Call dictionarySet.Add("regional_center", rst("regional_center"))
Call dictionarySet.Add("comments", rst("comments"))
Call dictionarySet.Add("billable_hours", rst("billable_hours"))
Call dictionarySet.Add("monthly_ihss_min", IHSS_MINS)
Call dictionarySet.Add("active", rst("active"))



If helpers.populateWDictionary(txtBoxAttrs, dictionarySet, Me, objs2Ignore) = True Then ' Fill in the fields based on record set for 'Me' form and the fieldKeys array
   Call helpers.change_control_caption("Update", Me.Create_Update)
End If

Call ToggleActive(rst("active")) ' Toggle the green active status based on the record set's value for the 'active' attribute

Set rst = Nothing ' Clear it out
Set txtBoxAttrs = Nothing ' Clear it out
Set sqlAttrs = Nothing

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
            MsgBox ("Client Add Error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select



End Sub


Private Sub Submit_Client()
On Error GoTo ErrorHandler ' Error handling

Dim client_id As Integer
Dim first_name As String
Dim last_name As String
Dim middle_name As String
Dim regional_center As Integer
Dim authorized_num As Long
Dim monthly_ihss_hrs As Long
Dim monthly_ihss_hrs_s As Single
Dim monthly_ihss_mins_s As Single
Dim monthly_ihss_mins As Long
Dim auth_hrs As Long
Dim comments As String
Dim active_value As Integer
Dim house_id As Integer: house_id = 0 ' Default
Dim fillColTab As String
Dim intCount As Integer: intCount = 0 ' Initialize to zero


If IsNull(Me.client_add_text_first) = False Then
    first_name = string_helpers.fStripIllegal(Me.client_add_text_first.Value)
Else
    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.client_add_text_last) = False Then
    last_name = string_helpers.fStripIllegal(Me.client_add_text_last.Value)
Else
    Resume ExitHandler ' Invoke Exit Handler with Error Code 20
End If

If IsNull(Me.client_add_text_middle) = False Then
    middle_name = string_helpers.fStripIllegal(Me.client_add_text_middle.Value)
Else
    middle_name = " " ' Set optional field to empty value
End If

If IsNull(Me.c_cb_region_select) = False And Me.c_cb_region_select.ListIndex <> -1 Then
    regional_center = Me.c_cb_region_select.Column(0) ' Grab the text value
Else
    regional_center = 0
End If

If IsNull(Me.client_add_text_authn) = False Then
    authorized_num = CLng(string_helpers.fStripIllegal(Me.client_add_text_authn.Value))
Else
   authorized_num = 0
End If

If IsNull(Me.client_add_text_monthly_ihss) = False Then
    monthly_ihss_hrs = CLng(string_helpers.fStripIllegal(Me.client_add_text_monthly_ihss.Value))
Else
    monthly_ihss_hrs = 0
End If

If IsNull(Me.client_add_text_monthly_ihss_min) = False Then
    monthly_ihss_mins = CLng(string_helpers.fStripIllegal(Me.client_add_text_monthly_ihss_min.Value))
Else
    monthly_ihss_mins = 0
End If


If IsNull(Me.client_add_Text_auth_billable_hrs) = False Then
     auth_hrs = CLng(string_helpers.fStripIllegal(Me.client_add_Text_auth_billable_hrs.Value))
Else
    auth_hrs = 0
End If


                'Dim ihssMins As Integer
                'ihssMins = CInt(monthly_ihss_mins)
               ' MsgBox (" " & ihssMins)
                'If ihssMins > 59 Then
                    'Dim addhour As Integer
                    'Dim minleft As Integer
                    'addhour = ihssMins / 60
                    'minleft = (ihssMins Mod 60)
                    'MsgBox (addhour & ":" & minleft)
                    'monthly_ihss_hrs = monthly_ihss_hrs + addhour
                    'monthly_ihss_mins = minleft
                'End If
                Dim IHSS_HOURS As Single
                Dim IHSS_MINS As Single
                Dim IHSS_MONTHLY As Double
                IHSS_HOURS = monthly_ihss_hrs
                IHSS_MINS = monthly_ihss_mins
                IHSS_MONTHLY = (IHSS_HOURS + (IHSS_MINS / 60))
                


If IsNull(Me.client_add_text_comments) = False Then
    comments = string_helpers.fStripIllegal(Me.client_add_text_comments.Value)
Else
    comments = " "
End If

If Me.c_active_toggle.BackColor = ToggleGreen Then
    active_value = 1
Else
    active_value = 0
End If


' CLng(Nz(DLookup("ID", "Table1"),0))

If IsNull(Me.c_cb_name_select.Value) = False Then
    client_id = Me.c_cb_name_select.Column(0)
    intCount = CLng(Nz(DLookup("ID", "clients", "ID = " & client_id), 0))
End If

If intCount > 0 Then ' Update client
    Dim lookupAttrs As Variant: lookupAttrs = Array("first_name = '" & first_name & "'", "last_name = '" & last_name & "'" _
    , "middle_initial = '" & middle_name & "'", "regional_center = " & regional_center, "authorized_regional_center_id = " & authorized_num _
    , "monthly_ihss = " & IHSS_MONTHLY, "billable_hours = " & auth_hrs, "active = " & active_value, "comments = '" & comments & "'")
    
    Dim query As String: query = string_helpers.assemble_query(updateQuery, lookupAttrs, Array("clients"), "WHERE ID = " & client_id)

    Call helpers.exec_query(query, Application.CurrentDb)
    Me![c_cb_name_select].Requery ' Refresh
    Set lookupAttrs = Nothing ' Clear it out
Else

Dim insertAttrs As Variant: insertAttrs = Array("'" & first_name & "'", "'" & last_name & "'" _
    , "'" & middle_name & "'", "" & authorized_num _
    , "" & IHSS_MONTHLY, "" & regional_center, "'" & comments & "'", "" & auth_hrs)

Dim attrs As Variant: attrs = Array("first_name", "last_name", "middle_initial", "authorized_regional_center_id", "monthly_ihss" _
, "regional_center", "comments", "billable_hours")
    
Dim insertVals As String: insertVals = string_helpers.join_array_to_str(insertAttrs, ",")

fillColTab = string_helpers.assemble_query(InsertQuery, attrs, Array("clients"), , insertVals)


Call helpers.exec_query(fillColTab, Application.CurrentDb)

' Cleanup but let the garbage collector do all the actual clean up
Set insertAttrs = Nothing
Set attrs = Nothing
insertVals = " " ' Clear it




Call reset_form ' Reset it to default state
Me![c_cb_name_select].Requery ' Refresh



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
            MsgBox ("Client_error Received: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
    
End Sub

Private Sub Delete_Client()
     On Error GoTo ErrorHandler
    
    Dim client_id As Integer
    Dim intCount As Integer
    Dim query As String
     If IsNull(Me.c_cb_name_select.Value) = False Then
        client_id = string_helpers.fStripIllegal(Me.c_cb_name_select.Column(0))
        intCount = CLng(Nz(DLookup("ID", "clients", "ID = " & client_id), 0))
    End If
    
    If intCount > 0 Then ' Valid delete since employee exists
        query = string_helpers.assemble_query(DeleteQuery, Nothing, Array("clients"), "WHERE ID = " & client_id)
        Call helpers.exec_query(query, Application.CurrentDb)
    Else
        MsgBox ("Not a valid client, please select a client from the combobox list")
        Exit Sub
    End If
    
    Call reset_form ' Reset the form
    Me![c_cb_name_select].Requery ' Refresh
    MsgBox "Client Record Successfully Deleted"
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Select Case Err
        Case 20 ' Silently exit
            ' Exit silent
        Case Else ' All other cases
            MsgBox ("Delete Client Error: " + Err.Description)
            Resume ExitHandler ' Invoke Exit Handler
    End Select
    
End Sub


Private Sub Create_Update_Click()
    Call Submit_Client
End Sub



Private Sub Delete_Click()
    
    Call Delete_Client
    
End Sub


Private Sub ToggleActive(ByVal onoff As Integer)
    If onoff = True Then ' Value must be 1
        Me.c_active_toggle.BackColor = ToggleGreen
         Call helpers.change_control_caption("De-activate", Me.Activate)
    ElseIf onoff = False Then ' Value must be 0
        Me.c_active_toggle.BackColor = ToggleRed
         Call helpers.change_control_caption("Activate", Me.Activate)
    Else ' Any other 'positive' integer
        Me.c_active_toggle.BackColor = RGB(191, 191, 191) ' Grey
        

    End If
End Sub

Private Function GetToggleVal() As Boolean
      If Me.c_active_toggle.BackColor = ToggleGreen Then ' Value must be 1
        GetToggleVal = True
    ElseIf Me.c_active_toggle.BackColor = ToggleRed Then ' Value must be 0
        GetToggleVal = False
    End If
End Function
