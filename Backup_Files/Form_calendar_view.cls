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

Private Sub Form_Load()


'first step
'get todays date
Dim objApp As Object

Dim todays_date As Date
todays_date = Date


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

Me.button_day_01.Caption = (dayval - Sun - 14)
Me.button_day_02.Caption = (dayval - Mon - 14)
Me.button_day_03.Caption = (dayval - Tue - 14)
Me.button_day_04.Caption = (dayval - Wed - 14)
Me.button_day_05.Caption = (dayval - Thu - 14)
Me.button_day_06.Caption = (dayval - Fri - 14)
Me.button_day_07.Caption = (dayval - Sat - 14)
Me.button_day_08.Caption = (dayval - Sun - 7)
Me.button_day_09.Caption = (dayval - Mon - 7)
Me.button_day_10.Caption = (dayval - Tue - 7)
Me.button_day_11.Caption = (dayval - Wed - 7)
Me.button_day_12.Caption = (dayval - Thu - 7)
Me.button_day_13.Caption = (dayval - Fri - 7)
Me.button_day_14.Caption = (dayval - Sat - 7)
Me.button_day_15.Caption = (dayval - Sun)
Me.button_day_16.Caption = (dayval - Mon)
Me.button_day_17.Caption = (dayval - Tue)
Me.button_day_18.Caption = (dayval - Wed)
Me.button_day_19.Caption = (dayval - Thu)
Me.button_day_20.Caption = (dayval - Fri)
Me.button_day_21.Caption = (dayval - Sat)
Me.button_day_22.Caption = (dayval - Sun + 7)
Me.button_day_23.Caption = (dayval - Mon + 7)
Me.button_day_24.Caption = (dayval - Tue + 7)
Me.button_day_25.Caption = (dayval - Wed + 7)
Me.button_day_26.Caption = (dayval - Thu + 7)
Me.button_day_27.Caption = (dayval - Fri + 7)
Me.button_day_28.Caption = (dayval - Sat + 7)
Me.button_day_29.Caption = (dayval - Sun + 14)
Me.button_day_30.Caption = (dayval - Mon + 14)
Me.button_day_31.Caption = (dayval - Tue + 14)
Me.button_day_32.Caption = (dayval - Wed + 14)
Me.button_day_33.Caption = (dayval - Thu + 14)
Me.button_day_34.Caption = (dayval - Fri + 14)
Me.button_day_35.Caption = (dayval - Sat + 14)
Me.button_day_36.Caption = (dayval - Sun + 21)
Me.button_day_37.Caption = (dayval - Mon + 21)
Me.button_day_38.Caption = (dayval - Tue + 21)
Me.button_day_39.Caption = (dayval - Wed + 21)
Me.button_day_40.Caption = (dayval - Thu + 21)
Me.button_day_41.Caption = (dayval - Fri + 21)
Me.button_day_42.Caption = (dayval - Sat + 21)







'center the calendar





End Sub
