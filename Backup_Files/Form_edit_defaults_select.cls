VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_edit_defaults_select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub add_shift(ByVal what_day As Integer)
Me.shift_sel.Value = Null
Me.day_of_week.Value = what_day
DoCmd.OpenForm "edit_default"
End Sub

Private Sub Box123_Click()

End Sub

Private Sub bt_add_01_Click()
Call add_shift(1)
End Sub

Private Sub bt_add_02_Click()
Call add_shift(2)
End Sub

Private Sub bt_add_03_Click()
Call add_shift(3)
End Sub

Private Sub bt_add_04_Click()
Call add_shift(4)
End Sub

Private Sub bt_add_05_Click()
Call add_shift(5)
End Sub

Private Sub bt_add_06_Click()
Call add_shift(6)
End Sub

Private Sub bt_add_07_Click()
Call add_shift(7)
End Sub

Private Sub bt_delete_01_Click()
Call delete_data(list_box_01, 1)
End Sub


Private Sub delete_data(ByRef box As Object, ByVal week_day As Integer)

On Error GoTo Err_delete_data

    If IsNull(box.Value) = False Then
        Dim Answer As Integer
        Answer = MsgBox("Are you sure you wish to delete this record?", vbYesNo + vbExclamation + vbDefaultButton2, "Delete Confirmation")
        If Answer = vbYes Then
            DoCmd.SetWarnings False
            Dim shift_number As Integer
            shiftNumber = CInt(box.Value)
            CurrentDb.Execute "delete * from default_shift where ID = " & shiftNumber
            Call populate_window(box, week_day)
            Me.Refresh
        End If
        DoCmd.SetWarnings True
    Else
    
    End If

Exit_delete_data:
    Exit Sub

Err_delete_data:
    MsgBox Err.Description
    Resume Exit_delete_data

End Sub


Private Sub bt_delete_02_Click()
Call delete_data(list_box_02, 2)
End Sub

Private Sub bt_delete_03_Click()
Call delete_data(list_box_03, 3)
End Sub

Private Sub bt_delete_04_Click()
Call delete_data(list_box_04, 4)
End Sub

Private Sub bt_delete_05_Click()
Call delete_data(list_box_05, 5)
End Sub

Private Sub bt_delete_06_Click()
Call delete_data(list_box_06, 6)
End Sub

Private Sub bt_delete_07_Click()
Call delete_data(list_box_07, 7)
End Sub

Private Sub bt_edit_01_Click()
Call edit_data(list_box_01, 1)
End Sub

Private Sub edit_data(ByRef box As Object, ByVal what_day As Integer)
 If IsNull(box.Value) = False Then
      Me.day_of_week.Value = what_day
      Me.shift_sel.Value = box.Value
      DoCmd.OpenForm "edit_default"
      End If
End Sub

Private Sub bt_edit_02_Click()
Call edit_data(list_box_02, 2)
End Sub

Private Sub bt_edit_03_Click()
Call edit_data(list_box_03, 3)
End Sub

Private Sub bt_edit_04_Click()
Call edit_data(list_box_04, 4)
End Sub

Private Sub bt_edit_05_Click()
Call edit_data(list_box_05, 5)
End Sub

Private Sub bt_edit_06_Click()
Call edit_data(list_box_06, 6)
End Sub

Private Sub bt_edit_07_Click()
Call edit_data(list_box_07, 7)
End Sub

Private Sub Form_Activate()
' bt_delete_0
 If IsNull(mc_cb_house_select.Value) = False Then
 
 Call mc_cb_house_select_Click
End If


End Sub

Private Sub Form_Load()
    
    
    MsgBox ("!!WARNING!!" & Chr(13) & "Changes to these values can have large" & Chr(13) & "and sometimes unforeseen impact on the" & Chr(13) & "scheduling application's functionality")
    
    Dim x As Variant ' Will be used to hold the values of Me.OpenArgs as it is split
    Dim house_id As Long
    
    Call update_houses_list ' First update the houses list manually
    
    If IsNull(Me.OpenArgs) = False Then ' Check for passed in parameters via the Global Me.OpenArgs variable
    ' split based on | delimeter
        x = Split(Me.OpenArgs, "|")
        strCtrl = x(0) ' The control name
        house_id = x(1) ' The house ID
        Me(strCtrl).Value = house_id
    End If
    
    
    [Forms]![edit_defaults_select]![house_id].Value = 1
    
    
    
    
End Sub


Private Sub update_houses_list()
 '   MsgBox ("updating house")
    Dim prepString As String
    Dim valIsNull As Boolean: valIsNull = True ' By default this value is null
    Dim rs1 As recordSet
    Set rs1 = CurrentDb.OpenRecordset("SELECT ID FROM houses;") ' get all the houses '
    If IsNull(Me.mc_cb_house_select.Value) = False Then ' Store the old value index
        valIsNull = False ' Not Null
    End If
    Call helpers.clear_cb_items(Me, "mc_cb_house_select") ' Clear/reset it
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
            Me.mc_cb_house_select.AddItem Item:=strItem ' Populate the list
        End If
        rs1.MoveNext
    Loop
    If valIsNull = False Then
        Me.mc_cb_house_select.Value = Me.mc_cb_house_select.OldValue ' Better way of retrieving the oldvalue via Microsoft Access
    End If
    Me.mc_cb_house_select.SetFocus ' This lets us add a visual emphasis on the combo box list
End Sub


Private Sub mc_cb_house_select_Click()
    
    
    Dim stupid As Integer
    stupid = CInt(Me.mc_cb_house_select.Value)
    house_id.Value = stupid
    
    'MsgBox (mc_cb_house_select.Column(1))
    house_label.Value = mc_cb_house_select.Column(1)
    Call populate_window(Me.list_box_01, 1)
    Call populate_window(Me.list_box_02, 2)
    Call populate_window(Me.list_box_03, 3)
    Call populate_window(Me.list_box_04, 4)
    Call populate_window(Me.list_box_05, 5)
    Call populate_window(Me.list_box_06, 6)
    Call populate_window(Me.list_box_07, 7)
End Sub


Private Sub populate_window(ByRef box As Object, ByVal week_day As Integer)

 
                
                Call ClearListBox(box)
                Dim rs2 As recordSet
                Dim prepString As String
                prepString = "" ' reset stringbuilder '
                Set rs2 = CurrentDb.OpenRecordset("SELECT [default_shift].[ID], [default_shift].[start_time], [default_shift].[end_time], [employees].[last_name], [employees].[first_name], [employees].[middle_initial]  FROM default_shift, employees WHERE [default_shift].[employee_id] = [employees].[id] AND [default_shift].[house_id] = " & Me.house_id.Value & " AND [default_shift].[day_of_week] = " & week_day & ";")
                If rs2.RecordCount > 0 Then ' if there are any clients then proceed '
                    Do While Not rs2.EOF
                        prepString = Format(rs2![start_time], "H:mm") + " - " + Format(rs2![end_time], "H:mm")
                        iLength = Len(prepString)
                        iLength = 13 - iLength
                        If iLength = 2 Then
                            prepString = prepString + Space(iLength + 4) + GetFirstLetter(rs2![first_name]) + GetFirstLetter(rs2![last_name]) + GetFirstLetter(rs2![middle_initial])
                        ElseIf iLength > 0 Then
                            
                            prepString = prepString + Space(iLength + 3) + GetFirstLetter(rs2![first_name]) + GetFirstLetter(rs2![last_name]) + GetFirstLetter(rs2![middle_initial])
                        Else
                            prepString = prepString + Space(2) + GetFirstLetter(rs2![first_name]) + GetFirstLetter(rs2![last_name]) + GetFirstLetter(rs2![middle_initial])
                        End If
                        
                        
                        Call AddItemToEnd(box, prepString, CInt(rs2![ID]))
                        
                        rs2.MoveNext
                    Loop
                Else
                    Call AddItemToEnd(box, "EMPTY", CInt(0))

                End If
                rs2.Close
            
           
            


End Sub

      Function GetFirstLetter(ByVal word1 As String) As String
            Dim word2 As String
            
            word2 = Left(word1, 1)
            
            GetFirstLetter = word2
            
        End Function
        
        Function ClearListBox(ctrlListBox As ListBox)
            
            ctrlListBox.RowSource = ""
            
        End Function
        
        Function AddItemToEnd(ctrlListBox As ListBox, _
            ByVal strItem As String, rsID As Integer)
            
            
            ctrlListBox.AddItem Item:=strItem + ";" + CStr(rsID)
            
        End Function
