VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_meeting_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
    Call update_houses_list
End Sub

Private Sub update_houses_list()
Dim prepString As String
Dim valIsNull As Boolean: valIsNull = True ' By default this value is null
Dim rs1 As recordSet
Set rs1 = CurrentDb.OpenRecordset("SELECT ID FROM houses;") ' get all the houses '

If IsNull(Me.me_cb_house_select.Value) = False Then ' Store the old value index
    valIsNull = False ' Not Null
End If

Call helpers.clear_cb_items(Me, "me_cb_house_select") ' Clear/reset it


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
    Me.me_cb_house_select.AddItem Item:=strItem ' Populate the list
     
    
    ' Column(0) = rs1![ID] ' set cb[0] to ID '
    ' Me.ha_cb_house_select.Column(1) = prepString ' set cb[1] to built string '
   
 End If
   
   rs1.MoveNext
Loop



If valIsNull = False Then
   Me.me_cb_house_select.Value = Me.me_cb_house_select.OldValue ' Better way of retrieving the oldvalue via Microsoft Access
End If

Me.me_cb_house_select.SetFocus ' This lets us add a visual emphasis on the combo box list



End Sub
