
  Sub Sendsms()
  On Error Resume Next
  Dim rows As Long
  rowsz = 1000
  rowsz = Wapp.Cells.End(xlDown).Row
  Dim z As Integer
  For z = 2 To rowsz
  Dim  ie As Object
  Set ie = CreateObject("InternetExplorer.Application")
  If Wapp.Cells(z, 2) = vbNullString Then
  Exit Sub
  Else
  ie.navigate "whatsapp://send?phone=" & Wapp.Cells(z, 2) & "&text=" & Wapp.Cells(z, 3)
  Application.Wait Now() + TimeSerial(0, 0, 3)
  SendKeys "~", True
  End If
  Next z
  End Sub


