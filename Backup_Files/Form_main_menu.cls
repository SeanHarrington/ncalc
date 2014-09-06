VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_main_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub bt_open_calendar_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command7_Click()

End Sub

Private Sub Form_Load()
MsgBox ("NCALC Known Bugs To Be Fixed:" & Chr(13) & _
        "- Special characters are unsupported for input. (!,',^, etc)" & Chr(13) & _
        "- Some date fields are in 24 hour format, some are in 12" & Chr(13) & _
        "- Createed Meetings do not check for already existing shifts")
End Sub
