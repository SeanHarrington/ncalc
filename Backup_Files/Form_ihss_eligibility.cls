VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ihss_eligibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database




Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    Dim x As Variant
    Dim strControl As String
    Dim lngID As Long
    If Len(Me.OpenArgs) > 0 Then
    
    
    ' split based on | delimeter
    
    x = Split(Me.OpenArgs, "|")
    strControl = x(0)
    lngID = x(1)
    'MsgBox ("strControl = '" & strControl & "' and ID = '" & lngID & "'")
    Me(strControl).Value = lngID
    
    
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

Private Sub Form_Load()

End Sub

Private Sub Form_Open(Cancel As Integer)


End Sub

Private Sub ihss_cb_client_select_Click()

End Sub

Private Sub ihss_cb_emp_select_Click()

End Sub

Private Sub ihss_chb_check_Click()

End Sub
