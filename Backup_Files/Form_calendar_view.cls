VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_calendar_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Function getDayStamp(ByVal days As Integer, ByVal modifier As Integer, ByRef box As Object) As Integer
    
    Dim dayval As Integer
    Dim todaysDate As Date
    Dim LastDay As Integer
    Dim PrevLastDay As Integer
    
    todaysDate = [Forms]![calendar_view]![date_box].Value 'selects the mid date for the calendar
    dayval = Day(todaysDate) 'this is the day of the week
    LastDay = Day(DateSerial(Year(todaysDate), Month(todaysDate) + 1, 0)) ' get this month's last day
    PrevLastDay = Day(DateSerial(Year(todaysDate), Month(todaysDate), 0)) ' get last months last day
    
    If dayval - days + modifier < 1 Then
        dayval = dayval + PrevLastDay
        box.ForeColor = RGB(46, 139, 87)
        box.FontUnderline = False
    ElseIf dayval - days + modifier > LastDay Then
        dayval = dayval - LastDay
        box.ForeColor = RGB(46, 139, 87)
        box.FontUnderline = False
    Else
        box.ForeColor = RGB(0, 0, 0)
        If dayval - days + modifier = dayval Then
            box.ForeColor = RGB(255, 0, 0)
            box.FontUnderline = True
        Else
            box.FontUnderline = False
        End If
    End If
    box.ControlTipText = todaysDate + modifier - days
    getDayStamp = dayval - days + modifier
    
End Function

Public Function button_day_stuff(ByVal date_value As Date)
    
    If IsNull([Forms]![calendar_view]![house_id].Value) = False Then
        Dim LWeekday As Integer
        LWeekday = Weekday(date_value, vbSunday)
        Forms!calendar_view.day_of_week = LWeekday
        Forms!calendar_view.dateStore = date_value
        DoCmd.OpenForm "day_view"
    End If
    
End Function

Private Sub bt_populate_week_1_Click()

  Call bt_populate_week(Me.button_day_01)
    Call bt_populate_week(Me.button_day_02)
    Call bt_populate_week(Me.button_day_03)
    Call bt_populate_week(Me.button_day_04)
    Call bt_populate_week(Me.button_day_05)
    Call bt_populate_week(Me.button_day_06)
    Call bt_populate_week(Me.button_day_07)
    If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub bt_populate_week_2_Click()

  Call bt_populate_week(Me.button_day_08)
    Call bt_populate_week(Me.button_day_09)
    Call bt_populate_week(Me.button_day_10)
    Call bt_populate_week(Me.button_day_11)
    Call bt_populate_week(Me.button_day_12)
    Call bt_populate_week(Me.button_day_13)
    Call bt_populate_week(Me.button_day_14)
    If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub bt_populate_week_3_Click()
    
    Call bt_populate_week(Me.button_day_15)
    Call bt_populate_week(Me.button_day_16)
    Call bt_populate_week(Me.button_day_17)
    Call bt_populate_week(Me.button_day_18)
    Call bt_populate_week(Me.button_day_19)
    Call bt_populate_week(Me.button_day_20)
    Call bt_populate_week(Me.button_day_21)
    If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub bt_populate_week_default(ByRef box As Object)
   Dim bt_name As String
    Dim day_string As String
    Dim day_integer As Integer
    Dim target_day_integer As Integer
    Dim target_day_string As String
    Dim do_q As String
    Dim date_value As Date
    Dim rs2 As recordSet
    
    bt_name = box.name
    date_value = CDate(box.ControlTipText)
    
    ' This tests to see if any data already exists in this spot
    Set rs2 = CurrentDb.OpenRecordset("SELECT [shift_default].[ID], [shift_default].[start_time], [shift_default].[end_time], [employees].[last_name], [employees].[first_name], [employees].[middle_initial]  FROM shift_default, employees WHERE [shift_default].[employee_id] = [employees].[id] AND [shift_default].[house_id] = " & Me.house_id.Value & " AND [shift_default].[shift_date] = #" & date_value & "#;")
    If rs2.RecordCount > 0 Then ' if there are any clients then proceed '
        do_q = False
    Else
        do_q = True
    
    End If
    rs2.Close

'
    
    If do_q = True Then

    Dim dayOfWeek As Integer
    dayOfWeek = Weekday(date_value)
 
 
 Dim NewSql As String
 NewSql = "SELECT *  FROM default_shift WHERE [default_shift].[house_id] = " & Me.house_id.Value & " AND [default_shift].[day_of_week] = " & dayOfWeek & ";"
 
 
       Set rs2 = CurrentDb.OpenRecordset(NewSql)
 
        
        
        If rs2.RecordCount > 0 Then ' if there are any clients then proceed '
            Do While Not rs2.EOF
                strSQL = ""
                Dim part1 As String
                Dim part2 As String
                part1 = "INSERT INTO shift_default("
                part2 = ") VALUES ("
                If IsNull(rs2![start_time]) = False Then
                    part1 = part1 & "start_time,"
                    part2 = part2 & "#" & rs2![start_time] & "#,"
                End If
                If IsNull(rs2![end_time]) = False Then
                    part1 = part1 & "end_time,"
                    part2 = part2 & "#" & rs2![end_time] & "#,"
                End If
                If IsNull(rs2![employee_id]) = False Then
                    part1 = part1 & "employee_id,"
                    part2 = part2 & rs2![employee_id] & ","
                End If
                If IsNull(rs2![house_id]) = False Then
                    part1 = part1 & "house_id,"
                    part2 = part2 & rs2![house_id] & ","
                End If
                
                part1 = part1 & "shift_date,"
                part2 = part2 & "#" & date_value & "#,"
                
                If IsNull(rs2![comments]) = False Then
                    If rs2![comments] = "" Then
                    Else
                        part1 = part1 & "comments,"
                        part2 = part2 & "'" & rs2![comments] & "',"
                    End If
                End If
                If IsNull(rs2![client_id]) = False Then
                    part1 = part1 & "client_id,"
                    part2 = part2 & rs2![client_id] & ","
                End If
                If IsNull(rs2![ihss]) = False Then
                    If rs2![ihss] = "Yes" Then
                        part1 = part1 & "ihss,"
                        part2 = part2 & "'Yes',"
                    Else
                        part1 = part1 & "ihss,"
                        part2 = part2 & "'No',"
                    End If
                End If
                If IsNull(rs2![sleep_shift]) = False Then
                    If rs2![sleep_shift] = "Yes" Then
                        part1 = part1 & "sleep_shift,"
                        part2 = part2 & "'Yes',"
                    Else
                        part1 = part1 & "sleep_shift,"
                        part2 = part2 & "'No',"
                    End If
                End If
                If IsNull(rs2![awake]) = False Then
                    If rs2![awake] = "Yes" Then
                        part1 = part1 & "awake,"
                        part2 = part2 & "'Yes',"
                    Else
                        part1 = part1 & "awake,"
                        part2 = part2 & "'No',"
                    End If
                End If
                
                part1 = Left(part1, Len(part1) - 1)
                part2 = Left(part2, Len(part2) - 1)
                strSQL = part1 & part2 & ");"
 
                CurrentDb.Execute strSQL
                rs2.MoveNext
            Loop
        Else
            'MsgBox ("has no values")
        End If
        rs2.Close





    End If

End Sub

Private Sub bt_populate_week(ByRef box As Object)
    
    Dim bt_name As String
    Dim day_string As String
    Dim day_integer As Integer
    Dim target_day_integer As Integer
    Dim target_day_string As String
    Dim do_q As String
    Dim date_value As Date
    Dim rs2 As recordSet
    
    bt_name = box.name
    date_value = CDate(box.ControlTipText)
    
    Set rs2 = CurrentDb.OpenRecordset("SELECT [shift_default].[ID], [shift_default].[start_time], [shift_default].[end_time], [employees].[last_name], [employees].[first_name], [employees].[middle_initial]  FROM shift_default, employees WHERE [shift_default].[employee_id] = [employees].[id] AND [shift_default].[house_id] = " & Me.house_id.Value & " AND [shift_default].[shift_date] = #" & date_value & "#;")
    
    If rs2.RecordCount > 0 Then ' if there are any clients then proceed '
        do_q = False
    Else
        do_q = True
    End If
    
    rs2.Close
    
    If do_q = True Then
        
        Dim dayval As Integer
        Dim todaysDate As Date
        Dim LastDay As Integer
        Dim PrevLastDay As Integer
        todaysDate = box.ControlTipText
        dayval = Day(todaysDate) 'this is the day of the week
        LastDay = Day(DateSerial(Year(todaysDate), Month(todaysDate) + 1, 0)) ' get this month's last day
        PrevLastDay = Day(DateSerial(Year(todaysDate), Month(todaysDate), 0)) ' get last months last day
        day_string = box.Caption
        day_integer = CInt(day_string)
        target_day_integer = day_integer - 7
        Dim targetDate As Date
        Dim targetString As String
        Dim targetMonth As Integer
        Dim targetYear As Integer
        targetYear = Year(date_value)
        targetMonth = Month(date_value)
        
        If target_day_integer < 1 Then
            target_day_integer = PrevLastDay + target_day_integer
            targetMonth = targetMonth - 1
            If targetMonth = 0 Then
                targetMonth = 12
                targetYear = targetYear - 1
            End If
        ElseIf target_day_integer > LastDay Then
            targetMonth = targetMonth + 1
            If targetMonth = 13 Then
                targetMonth = 1
                targetYear = targetYear + 1
            End If
            target_day_integer = target_day_integer - LastDay
        End If
        target_day_string = CStr(target_day_integer)
        targetString = CStr(targetMonth) + "/" + target_day_string + "/" + CStr(targetYear)
        targetDate = CDate(targetString)
        Dim strSQL As String
        
        Set rs2 = CurrentDb.OpenRecordset("SELECT *  FROM shift_default WHERE [shift_default].[house_id] = " & Me.house_id.Value & " AND [shift_default].[shift_date] = #" & targetDate & "#;")
        If rs2.RecordCount > 0 Then ' if there are any clients then proceed '
            Do While Not rs2.EOF
                strSQL = ""
                Dim part1 As String
                Dim part2 As String
                part1 = "INSERT INTO shift_default("
                part2 = ") VALUES ("
                If IsNull(rs2![start_time]) = False Then
                    part1 = part1 & "start_time,"
                    part2 = part2 & "#" & rs2![start_time] & "#,"
                End If
                If IsNull(rs2![end_time]) = False Then
                    part1 = part1 & "end_time,"
                    part2 = part2 & "#" & rs2![end_time] & "#,"
                End If
                If IsNull(rs2![employee_id]) = False Then
                    part1 = part1 & "employee_id,"
                    part2 = part2 & rs2![employee_id] & ","
                End If
                If IsNull(rs2![house_id]) = False Then
                    part1 = part1 & "house_id,"
                    part2 = part2 & rs2![house_id] & ","
                End If
                
                part1 = part1 & "shift_date,"
                part2 = part2 & "#" & date_value & "#,"
                
                If IsNull(rs2![comments]) = False Then
                    If rs2![comments] = "" Then
                    Else
                        part1 = part1 & "comments,"
                        part2 = part2 & "'" & rs2![comments] & "',"
                    End If
                End If
                If IsNull(rs2![client_id]) = False Then
                    part1 = part1 & "client_id,"
                    part2 = part2 & rs2![client_id] & ","
                End If
                If IsNull(rs2![ihss]) = False Then
                    If rs2![ihss] = "Yes" Then
                        part1 = part1 & "ihss,"
                        part2 = part2 & "'Yes',"
                    Else
                        part1 = part1 & "ihss,"
                        part2 = part2 & "'No',"
                    End If
                End If
                If IsNull(rs2![sleep_shift]) = False Then
                    If rs2![sleep_shift] = "Yes" Then
                        part1 = part1 & "sleep_shift,"
                        part2 = part2 & "'Yes',"
                    Else
                        part1 = part1 & "sleep_shift,"
                        part2 = part2 & "'No',"
                    End If
                End If
                If IsNull(rs2![awake]) = False Then
                    If rs2![awake] = "Yes" Then
                        part1 = part1 & "awake,"
                        part2 = part2 & "'Yes',"
                    Else
                        part1 = part1 & "awake,"
                        part2 = part2 & "'No',"
                    End If
                End If
                
                part1 = Left(part1, Len(part1) - 1)
                part2 = Left(part2, Len(part2) - 1)
                strSQL = part1 & part2 & ");"
              '  MsgBox (strSQL)
                Dim db As DAO.Database
                Dim newRow As Long
                Set db = CurrentDb
                db.Execute strSQL
                
                            
                newRow = db.OpenRecordset("SELECT @@IDENTITY")(0)
                Set db = Nothing
                                
                Dim old_shift_id As Integer
                old_shift_id = rs2![ID]
                Dim new_shift_id As Integer
                new_shift_id = CInt(newRow)
                
                
        Set rs3 = CurrentDb.OpenRecordset("SELECT *  FROM shift_client WHERE shift_id = " & old_shift_id & ";")
        If rs3.RecordCount > 0 Then ' if there are any clients then proceed '
            Do While Not rs3.EOF
                'MsgBox ("INSERT INTO shift_client (client_id,employee_id,shift_id,ihss) VALUES (" & CInt(rs3![client_id]) & ", " & CInt(rs3![employee_id]) & ", " & new_shift_id & ", " & CInt(rs3![ihss]) & ");")
                CurrentDb.Execute "INSERT INTO shift_client (client_id,employee_id,shift_id,ihss) VALUES (" & CInt(rs3![client_id]) & ", " & CInt(rs3![employee_id]) & ", " & new_shift_id & ", " & CInt(rs3![ihss]) & ");"
                rs3.MoveNext
            Loop
        End If
        rs3.Close
                                
                                
                                
                rs2.MoveNext
            Loop
        Else
            'MsgBox ("has no values")
        End If
        rs2.Close
    End If
     
End Sub


Private Sub bt_populate_week_4_Click()

 Call bt_populate_week(Me.button_day_22)
    Call bt_populate_week(Me.button_day_23)
    Call bt_populate_week(Me.button_day_24)
    Call bt_populate_week(Me.button_day_25)
    Call bt_populate_week(Me.button_day_26)
    Call bt_populate_week(Me.button_day_27)
    Call bt_populate_week(Me.button_day_28)
    If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub bt_populate_week_5_Click()

 Call bt_populate_week(Me.button_day_29)
    Call bt_populate_week(Me.button_day_30)
    Call bt_populate_week(Me.button_day_31)
    Call bt_populate_week(Me.button_day_32)
    Call bt_populate_week(Me.button_day_33)
    Call bt_populate_week(Me.button_day_34)
    Call bt_populate_week(Me.button_day_35)
    If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub bt_populate_week_6_Click()

 Call bt_populate_week(Me.button_day_36)
    Call bt_populate_week(Me.button_day_37)
    Call bt_populate_week(Me.button_day_38)
    Call bt_populate_week(Me.button_day_39)
    Call bt_populate_week(Me.button_day_40)
    Call bt_populate_week(Me.button_day_41)
    Call bt_populate_week(Me.button_day_42)
    If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub bt_populate_week_default_1_Click()

    Call bt_populate_week_default(Me.button_day_01)
    Call bt_populate_week_default(Me.button_day_02)
    Call bt_populate_week_default(Me.button_day_03)
    Call bt_populate_week_default(Me.button_day_04)
    Call bt_populate_week_default(Me.button_day_05)
    Call bt_populate_week_default(Me.button_day_06)
    Call bt_populate_week_default(Me.button_day_07)
    
    If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If

End Sub

Private Sub bt_populate_week_default_2_Click()

    Call bt_populate_week_default(Me.button_day_08)
    Call bt_populate_week_default(Me.button_day_09)
    Call bt_populate_week_default(Me.button_day_10)
    Call bt_populate_week_default(Me.button_day_11)
    Call bt_populate_week_default(Me.button_day_12)
    Call bt_populate_week_default(Me.button_day_13)
    Call bt_populate_week_default(Me.button_day_14)
    
    If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub bt_populate_week_default_3_Click()

    Call bt_populate_week_default(Me.button_day_15)
    Call bt_populate_week_default(Me.button_day_16)
    Call bt_populate_week_default(Me.button_day_17)
    Call bt_populate_week_default(Me.button_day_18)
    Call bt_populate_week_default(Me.button_day_19)
    Call bt_populate_week_default(Me.button_day_20)
    Call bt_populate_week_default(Me.button_day_21)
    
      If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub bt_populate_week_default_4_Click()

    Call bt_populate_week_default(Me.button_day_22)
    Call bt_populate_week_default(Me.button_day_23)
    Call bt_populate_week_default(Me.button_day_24)
    Call bt_populate_week_default(Me.button_day_25)
    Call bt_populate_week_default(Me.button_day_26)
    Call bt_populate_week_default(Me.button_day_27)
    Call bt_populate_week_default(Me.button_day_28)
    
      If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub bt_populate_week_default_5_Click()

    Call bt_populate_week_default(Me.button_day_29)
    Call bt_populate_week_default(Me.button_day_30)
    Call bt_populate_week_default(Me.button_day_31)
    Call bt_populate_week_default(Me.button_day_32)
    Call bt_populate_week_default(Me.button_day_33)
    Call bt_populate_week_default(Me.button_day_34)
    Call bt_populate_week_default(Me.button_day_35)
      If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub bt_populate_week_default_6_Click()

    Call bt_populate_week_default(Me.button_day_36)
    Call bt_populate_week_default(Me.button_day_37)
    Call bt_populate_week_default(Me.button_day_38)
    Call bt_populate_week_default(Me.button_day_39)
    Call bt_populate_week_default(Me.button_day_40)
    Call bt_populate_week_default(Me.button_day_41)
    Call bt_populate_week_default(Me.button_day_42)
    
      If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        Call populate_window
    End If
    
End Sub

Private Sub button_day_01_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_01.ControlTipText)
    Call button_day_stuff(date_value)
End Sub

Private Sub button_day_02_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_02.ControlTipText)
    Call button_day_stuff(date_value)
End Sub

Private Sub button_day_03_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_03.ControlTipText)
    Call button_day_stuff(date_value)
End Sub

Private Sub button_day_04_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_04.ControlTipText)
    Call button_day_stuff(date_value)
End Sub

Private Sub button_day_05_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_05.ControlTipText)
    Call button_day_stuff(date_value)
End Sub

Private Sub button_day_06_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_06.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_07_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_07.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_08_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_08.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_09_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_09.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_10_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_10.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_11_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_11.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_12_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_12.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_13_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_13.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_14_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_14.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_15_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_15.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_16_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_16.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_17_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_17.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_18_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_18.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_19_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_19.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_20_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_20.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_21_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_21.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_22_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_22.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_23_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_23.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_24_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_24.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_25_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_25.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_26_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_26.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_27_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_27.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_28_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_28.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_29_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_29.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_30_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_30.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_31_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_31.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_32_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_32.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_33_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_33.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_34_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_34.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_35_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_35.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_36_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_36.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_37_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_37.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_38_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_38.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_39_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_39.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_40_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_40.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_41_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_41.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub button_day_42_Click()
    Dim date_value As Date
    date_value = CDate(Me.button_day_42.ControlTipText)
    Call button_day_stuff(date_value)
    
End Sub

Private Sub DATE_PICK_Click()
    If IsNull(CInt([Forms]![calendar_view]![house_id].Value)) = False Then
        Call Calendar_Pop
        
        Call populate_window
    End If
    
End Sub

Private Sub Form_Activate()

DATE_PICK_Click


End Sub

Private Sub Form_GotFocus()
End Sub

Private Sub Form_Load()
    
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
    
    Call populate_date
    Call Calendar_Pop
    [Forms]![calendar_view]![house_id].Value = 1
    
End Sub

Private Sub Calendar_Pop()
    
    
    Dim todays_date As Date
    todays_date = [Forms]![calendar_view]![date_box].Value
    
    
    Dim LWeekday As Integer
    LWeekday = Weekday(todays_date, vbSunday)
    
    Dim dayval As Integer
    dayval = Day(todays_date)
    
    Dim Sun As Integer
    Dim Mon As Integer
    Dim Tue As Integer
    Dim Wed As Integer
    Dim Thu As Integer
    Dim Fri As Integer
    Dim Sat As Integer
    
    
    
    
    If LWeekday = 1 Then
        Sun = 0
        Mon = -1
        Tue = -2
        Wed = -3
        Thu = -4
        Fri = -5
        Sat = -6
    End If
    
    
    If LWeekday = 2 Then
        Sun = 1
        Mon = 0
        Tue = -1
        Wed = -2
        Thu = -3
        Fri = -4
        Sat = -5
    End If
    
    
    If LWeekday = 3 Then
        Sun = 2
        Mon = 1
        Tue = 0
        Wed = -1
        Thu = -2
        Fri = -3
        Sat = -4
    End If
    
    
    If LWeekday = 4 Then
        Sun = 3
        Mon = 2
        Tue = 1
        Wed = 0
        Thu = -1
        Fri = -2
        Sat = -3
    End If
    
    
    If LWeekday = 5 Then
        Sun = 4
        Mon = 3
        Tue = 2
        Wed = 1
        Thu = 0
        Fri = -1
        Sat = -2
    End If
    
    
    
    If LWeekday = 6 Then
        Sun = 5
        Mon = 4
        Tue = 3
        Wed = 2
        Thu = 1
        Fri = 0
        Sat = -1
    End If
    
    
    If LWeekday = 7 Then
        Sun = 6
        Mon = 5
        Tue = 4
        Wed = 3
        Thu = 2
        Fri = 1
        Sat = 0
    End If
    
    
    Dim wkMod As Integer
    
'getDayStamp(Sun,14)
    wkMod = -7
    Me.button_day_01.Caption = getDayStamp(Sun, wkMod, button_day_01)
    
    Me.button_day_02.Caption = getDayStamp(Mon, wkMod, button_day_02)
    Me.button_day_03.Caption = getDayStamp(Tue, wkMod, button_day_03)
    Me.button_day_04.Caption = getDayStamp(Wed, wkMod, button_day_04)
    Me.button_day_05.Caption = getDayStamp(Thu, wkMod, button_day_05)
    Me.button_day_06.Caption = getDayStamp(Fri, wkMod, button_day_06)
    Me.button_day_07.Caption = getDayStamp(Sat, wkMod, button_day_07)
    
    wkMod = 0
    Me.button_day_08.Caption = getDayStamp(Sun, wkMod, button_day_08)
    Me.button_day_09.Caption = getDayStamp(Mon, wkMod, button_day_09)
    Me.button_day_10.Caption = getDayStamp(Tue, wkMod, button_day_10)
    Me.button_day_11.Caption = getDayStamp(Wed, wkMod, button_day_11)
    Me.button_day_12.Caption = getDayStamp(Thu, wkMod, button_day_12)
    Me.button_day_13.Caption = getDayStamp(Fri, wkMod, button_day_13)
    Me.button_day_14.Caption = getDayStamp(Sat, wkMod, button_day_14)
    wkMod = 7
    Me.button_day_15.Caption = getDayStamp(Sun, wkMod, button_day_15)
    Me.button_day_16.Caption = getDayStamp(Mon, wkMod, button_day_16)
    Me.button_day_17.Caption = getDayStamp(Tue, wkMod, button_day_17)
    Me.button_day_18.Caption = getDayStamp(Wed, wkMod, button_day_18)
    Me.button_day_19.Caption = getDayStamp(Thu, wkMod, button_day_19)
    Me.button_day_20.Caption = getDayStamp(Fri, wkMod, button_day_20)
    Me.button_day_21.Caption = getDayStamp(Sat, wkMod, button_day_21)
    wkMod = 14
    Me.button_day_22.Caption = getDayStamp(Sun, wkMod, button_day_22)
    Me.button_day_23.Caption = getDayStamp(Mon, wkMod, button_day_23)
    Me.button_day_24.Caption = getDayStamp(Tue, wkMod, button_day_24)
    Me.button_day_25.Caption = getDayStamp(Wed, wkMod, button_day_25)
    Me.button_day_26.Caption = getDayStamp(Thu, wkMod, button_day_26)
    Me.button_day_27.Caption = getDayStamp(Fri, wkMod, button_day_27)
    Me.button_day_28.Caption = getDayStamp(Sat, wkMod, button_day_28)
    wkMod = 21
    Me.button_day_29.Caption = getDayStamp(Sun, wkMod, button_day_29)
    Me.button_day_30.Caption = getDayStamp(Mon, wkMod, button_day_30)
    Me.button_day_31.Caption = getDayStamp(Tue, wkMod, button_day_31)
    Me.button_day_32.Caption = getDayStamp(Wed, wkMod, button_day_32)
    Me.button_day_33.Caption = getDayStamp(Thu, wkMod, button_day_33)
    Me.button_day_34.Caption = getDayStamp(Fri, wkMod, button_day_34)
    Me.button_day_35.Caption = getDayStamp(Sat, wkMod, button_day_35)
    wkMod = 28
    Me.button_day_36.Caption = getDayStamp(Sun, wkMod, button_day_36)
    Me.button_day_37.Caption = getDayStamp(Mon, wkMod, button_day_37)
    Me.button_day_38.Caption = getDayStamp(Tue, wkMod, button_day_38)
    Me.button_day_39.Caption = getDayStamp(Wed, wkMod, button_day_39)
    Me.button_day_40.Caption = getDayStamp(Thu, wkMod, button_day_40)
    Me.button_day_41.Caption = getDayStamp(Fri, wkMod, button_day_41)
    Me.button_day_42.Caption = getDayStamp(Sat, wkMod, button_day_42)
End Sub

Private Sub update_houses_list()
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
    Call Calendar_Pop
    
    Dim stupid As Integer
    stupid = CInt(Me.mc_cb_house_select.Value)
    [Forms]![calendar_view]![house_id].Value = stupid
    Call populate_window
End Sub

Private Sub populate_window()
    Me.house_id.Value = CInt([Forms]![calendar_view]![house_id].Value)
    Dim e As Control
    Dim TestArray As Variant
    Dim iLength As Integer
    iLength = 0
    For Each e In Me.Controls
        If (TypeOf e Is ListBox) Then
            Call ClearListBox(e)
            Dim listBoxName As String
            listBoxName = e.name
            TestArray = Split(listBoxName, "_")
            listBoxName = TestArray(2)
            
            Dim day_of_week As Integer
            day_of_week = CInt(listBoxName)
            day_of_week = day_of_week Mod 7
            
            If day_of_week = 0 Then
                day_of_week = 7
            End If
            
            listBoxName = "button_day_" & listBoxName
            
            Dim date_value As Date
            For Each f In Me.Controls
                If (TypeOf f Is CommandButton) Then
                    If f.name = listBoxName Then
                        date_value = CDate(f.ControlTipText)
                        
                        
                    End If
                End If
                Next f
                
' MsgBox (date_value)
'   Dim date_value As Date
'    Dim button As Variant
'   button = Me.listBoxName
                
                
'  date_value = CDate(button.ControlTipText)
'MsgBox (date_value)
                
' SELECT [shift_default].[ID], [shift_default].[start_time], [shift_default].[end_time], [employees].[last_name] + ",  " + [employees].[first_name] + " " + [employees].[middle_initial] FROM shift_default, employees WHERE [shift_default].[employee_id] = [employees].[id] AND [shift_default].[house_id] =  [Forms]![calendar_view]![house_id].Value AND [shift_default].[shift_date] = [Forms]![calendar_view].[dateStore].Value
                Dim rs2 As recordSet
                Dim prepString As String
                prepString = "" ' reset stringbuilder '
                Set rs2 = CurrentDb.OpenRecordset("SELECT [shift_default].[ID], [shift_default].[exception_shift], [shift_default].[start_time], [shift_default].[end_time], [employees].[last_name], [employees].[first_name], [employees].[middle_initial]  FROM shift_default, employees WHERE [shift_default].[employee_id] = [employees].[id] AND [shift_default].[house_id] = " & Me.house_id.Value & " AND [shift_default].[shift_date] = #" & date_value & "#;")
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
                        
                        If rs2![exception_shift] = "Meeting:House" Then
                            prepString = prepString + "*"
                        End If
                        
                        Call AddItemToEnd(e, prepString)
                        
                        rs2.MoveNext
                    Loop
                Else
                    Call AddItemToEnd(e, "EMPTY")
'MsgBox ("found nothing")
                End If
                rs2.Close
            End If
            Next e
            
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
            ByVal strItem As String)
            
            
            ctrlListBox.AddItem Item:=strItem
            
        End Function
        
        Private Sub populate_date()
            [Forms]![calendar_view]![date_box].Value = Format(Date, "Short Date")
        End Sub
        


