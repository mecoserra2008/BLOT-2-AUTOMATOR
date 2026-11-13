Option Explicit

Private Sub Workbook_Open()
    ScheduleDaily
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    OnWorkbookClose
End Sub
