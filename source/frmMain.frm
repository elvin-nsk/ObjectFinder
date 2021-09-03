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
Private Const txtWrongValue As String = "Неправильное значение"
Private Const txtNoShapesSelected As String = _
  "Выделите объекты, среди которых осуществляется поиск"
Private Const txtNoActiveDocument As String = "Нет открытых документов"
Private Const txtNotFound As String = "По заданным условиям объектов не найдено"
Private Const txtAnd As String = "И"
Private Const txtOr As String = "ИЛИ"

'===============================================================================

Private Sub GetSize()
    
  On Error GoTo ErrHandler
    
  Dim Document As Document
  Set Document = ActiveDocument
  If Document Is Nothing Then
    MsgBox txtNoActiveDocument, vbCritical
    Exit Sub
  End If
  
  Dim Shape As Shape
  Set Shape = Document.ActiveShape
  If Shape Is Nothing Then
    MsgBox txtNoShapesSelected, vbCritical
    Exit Sub
  End If
  
  txtW = Round(Shape.SizeWidth, 3)
  txtH = Round(Shape.SizeHeight, 3)
    
  Exit Sub
    
ErrHandler:
  MsgBox "Ошибка! " & Err.Description & " | " & Err.LastDllError & " | " & Err.Number

End Sub

Private Sub Search()

  Dim Shapes As Shapes
  Dim ShapesFound As ShapeRange
    
  Dim getW As Double
  Dim getH As Double
  Dim resW As Double
  Dim resH As Double
    
  getW = Val(txtW)
  getH = Val(txtH)
  If getH <= 0 Or getW <= 0 Then
    MsgBox txtWrongValue, vbCritical
    Exit Sub
  End If
    
  ActiveDocument.Unit = cdrMillimeter
  Set Shapes = ActiveSelection.Shapes
  If Shapes.Count = 0 Then
    MsgBox txtNoShapesSelected, vbCritical
    Exit Sub
  End If
  
  Set ShapesFound = CreateShapeRange
  Dim Shape As Shape
  For Each Shape In Shapes
    resW = IIf(cmbOpW.ListIndex, getW - Shape.SizeWidth, Shape.SizeWidth - getW)
    resH = IIf(cmbOpH.ListIndex, getH - Shape.SizeHeight, Shape.SizeHeight - getH)
    If IIf(opLog, (resW > 0) And (resH > 0), (resW > 0) Or (resH > 0)) _
      Then ShapesFound.Add Shape
  Next Shape
    
  If ShapesFound.Count > 0 Then
    ShapesFound.CreateSelection
  Else
    MsgBox txtNotFound, vbInformation
  End If
  
End Sub

Private Sub LoadSettings()
  Me.Top = GetSetting(Me.Tag, "Settings", "Top", 100)
  Me.Left = GetSetting(Me.Tag, "Settings", "Left", 100)
  cmbOpW.ListIndex = GetSetting(Me.Tag, "Settings", "ListW", 1)
  cmbOpH.ListIndex = GetSetting(Me.Tag, "Settings", "Listh", 1)
  txtW = GetSetting(Me.Tag, "Settings", "SizeW", 1)
  txtH = GetSetting(Me.Tag, "Settings", "SizeH", 1)
  opLog = GetSetting(Me.Tag, "Settings", "OpLog", True)
End Sub

Private Sub SaveSettings()
  SaveSetting Me.Tag, "Settings", "Top", Me.Top
  SaveSetting Me.Tag, "Settings", "Left", Me.Left
  SaveSetting Me.Tag, "Settings", "SizeW", txtW
  SaveSetting Me.Tag, "Settings", "SizeH", txtH
  SaveSetting Me.Tag, "Settings", "OpLog", opLog
  SaveSetting Me.Tag, "Settings", "ListW", cmbOpW.ListIndex
  SaveSetting Me.Tag, "Settings", "ListH", cmbOpH.ListIndex
End Sub

Private Sub FormCancel()
  Me.Hide
End Sub

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
    
  LoadSettings

End Sub

'===============================================================================

Private Sub cmdGetSize_Click()
  GetSize
End Sub

Private Sub cmdOK_Click()
  Search
End Sub

Private Sub opLog_Click()
  If opLog Then
    opLog.Caption = txtAnd
  Else
    opLog.Caption = txtOr
  End If
End Sub

Private Sub txtH_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNum KeyAscii
End Sub

Private Sub txtW_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNum KeyAscii
End Sub

'===============================================================================

Private Sub OnlyNum(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc(",")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    SaveSettings
    Сancel = True
    FormCancel
  End If
End Sub
