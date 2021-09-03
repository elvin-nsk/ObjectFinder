VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Найти объекты по размеру"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   OleObjectBlob   =   "frmMain.frx":0000
   Tag             =   "CorelObjectFinder"
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Private Const txtGreater As String = "Больше чем"
Private Const txtSmaller As String = "Меньше чем"
Private Const txtNoShapesSelected As String = _
  "Выделите объекты, среди которых осуществляется поиск"
Private Const txtNoActiveDocument As String = "Нет открытых документов"

'===============================================================================

Private Sub UserForm_Activate()

  With cmbOpW
    .AddItem txtGreater
    .AddItem txtSmaller
  End With
    
  With cmbOpH
    .AddItem txtGreater
    .AddItem txtSmaller
  End With
    
  Me.Top = GetSetting(Me.Tag, "Settings", "Top", 100)
  Me.Left = GetSetting(Me.Tag, "Settings", "Left", 100)
  cmbOpW.ListIndex = GetSetting(Me.Tag, "Settings", "ListW", 1)
  cmbOpH.ListIndex = GetSetting(Me.Tag, "Settings", "Listh", 1)
  txtW = GetSetting(Me.Tag, "Settings", "SizeW", 1)
  txtH = GetSetting(Me.Tag, "Settings", "SizeH", 1)
  opLog = GetSetting(Me.Tag, "Settings", "OpLog", True)

End Sub

'===============================================================================

Private Sub cmdGetSize_Click()

  Dim Document As Document
  Dim Shape As Shape
    
  On Error GoTo ErrHandler
    
  Set Document = ActiveDocument
    
  If Not Document Is Nothing Then
    Document.Unit = cdrMillimeter
    Set Shape = Document.ActiveShape
    If Not Shape Is Nothing Then
      txtW = Round(Shape.SizeWidth, 3)
      txtH = Round(Shape.SizeHeight, 3)
    Else
      MsgBox txtNoShapesSelected, vbCritical
    End If
  Else
    MsgBox txtNoActiveDocument, vbCritical
  End If
    
  Exit Sub
    
ErrHandler:
  MsgBox "Error occured! " & Err.Description & " | " & Err.LastDllError & " | " & Err.Number
  
End Sub

Private Sub cmdOK_Click()

  Dim Shapes As Shapes
  Dim Shape As Shape
    
  Dim i As Long
  Dim shID() As Long
    
  Dim getW As Double
  Dim getH As Double
  Dim resW As Double
  Dim resH As Double
  Dim resB As Boolean
    
  getW = Val(txtW)
  getH = Val(txtH)
  If getH <= 0 Or getW <= 0 Then
    MsgBox "Неправильное значение", vbCritical
    Exit Sub
  End If
    
  ActiveDocument.Unit = cdrMillimeter
  Set Shapes = ActiveSelection.Shapes
  If Shapes.Count = 0 Then
    MsgBox txtNoShapesSelected, vbCritical
    Exit Sub
  End If
  i = 0
    
  For Each Shape In Shapes
    resW = IIf(cmbOpW.ListIndex, getW - Shape.SizeWidth, Shape.SizeWidth - getW)
    resH = IIf(cmbOpH.ListIndex, getH - Shape.SizeHeight, Shape.SizeHeight - getH)
    resB = IIf(opLog, (resW > 0) And (resH > 0), (resW > 0) Or (resH > 0))
        
    If resB Then
      i = i + 1
      ReDim Preserve shID(1 To i) As Long
      shID(i) = Shape.ObjectData("CDRStaticID").Value
    End If
  Next Shape
    
  If i > 0 Then
    
    ActivePage.Shapes.All.RemoveFromSelection
    Set Shape = Nothing
    
    For i = LBound(shID) To UBound(shID)
      Set Shape = ActivePage.Shapes.FindShape(, , shID(i))
      Shape.AddToSelection
    Next i
  Else
    MsgBox "По заданным условиям объектов не найдено", vbCritical
    
  End If
  
End Sub

Private Sub CommandButton1_Click()
  Unload Me
End Sub

Private Sub opLog_Click()
  If opLog Then
    opLog.Caption = "И"
  Else
    opLog.Caption = "ИЛИ"
  End If
End Sub

Private Sub txtH_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 46 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub txtW_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 46 Or KeyAscii = 8 Then Exit Sub
  If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  SaveSetting Me.Tag, "Settings", "Top", Me.Top
  SaveSetting Me.Tag, "Settings", "Left", Me.Left
  SaveSetting Me.Tag, "Settings", "SizeW", txtW
  SaveSetting Me.Tag, "Settings", "SizeH", txtH
  SaveSetting Me.Tag, "Settings", "OpLog", opLog
  SaveSetting Me.Tag, "Settings", "ListW", cmbOpW.ListIndex
  SaveSetting Me.Tag, "Settings", "ListH", cmbOpH.ListIndex
End Sub

'===============================================================================
