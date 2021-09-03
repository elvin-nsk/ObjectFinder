Attribute VB_Name = "ObjectFinderMain"
Sub FindObjectBySize()
  If Documents.Count > 0 Then
    frmMain.Show vbModeless
  Else
    MsgBox "Нет открытых документов", vbCritical
  End If
End Sub

