VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+---------------------------------------------------------------
'| PVL_Resizer class
'| Copyright © 2003 by Tomasz Puwalski (tpuwalski@op.pl)
'|
'| Everyone has a full permission to use this code
'| in any kind of software as long as this copyright
'| notice and full information about contacting
'| the author is preserved.
'|
'| If you feel that this code is useful and/or if you want
'| to use this code in your own software - please appreciate
'| my work and leave your vote on:
'| http://www.pscode.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=50708
'+---------------------------------------------------------------
Option Explicit

Const AN_NONE = 0
Const AN_TOP = 1
Const AN_BOTTOM = 2
Const AN_LEFT = 4
Const AN_RIGHT = 8

Const AA_ANCHORS = 0
Const AA_LEFT = 1
Const AA_TOP = 2
Const AA_WIDTH = 3
Const AA_HEIGHT = 4

Enum AnchorValue
  avNone = AN_NONE
  avTop = AN_TOP
  avBottom = AN_BOTTOM
  avLeft = AN_LEFT
  avRight = AN_RIGHT
End Enum

Private objContainer As Object
Private dicControls As New Dictionary
Private intFormWidth As Integer, intFormHeight As Integer

Public Property Set Container(MainContainer As Object)
  ' In a valid container, controls collection named
  ' "Controls" must be available. Use standard form
  ' or PVL_Container control as a container.
  Dim objControl As Object
  
  On Error Resume Next
  Set objContainer = MainContainer
  If TypeOf objContainer Is Form Then
    intFormWidth = objContainer.ScaleWidth
    intFormHeight = objContainer.ScaleHeight
  Else
    intFormWidth = objContainer.Width
    intFormHeight = objContainer.Height
  End If
End Property

Public Sub SetAnchors(ByVal ControlName As String, ByVal Anchors As AnchorValue)
  Dim intControlIndex As Integer
  Dim objControl As Object
  
  If InStr(ControlName, "(") > 0 Then
    ' Control Array...
    intControlIndex = CInt(Replace(Mid(ControlName, InStr(ControlName, "(") + 1), ")", ""))
    For Each objControl In objContainer.Controls
      If objControl.Name = Left(ControlName, InStr(ControlName, "(") - 1) Then
        If objControl.Index = intControlIndex Then
          ' Above two IFs can't be merged together!
          Exit For
        End If
      End If
    Next
  Else
    Set objControl = objContainer.Controls(ControlName)
  End If
  With objControl
    If dicControls.Exists(ControlName) Then
      dicControls(ControlName) = Array(Anchors, .Left, .Top, .Width, .Height)
    Else
      dicControls.Add ControlName, Array(Anchors, .Left, .Top, .Width, .Height)
    End If
  End With
End Sub

Public Sub ReSet()
  ' This procedure should be called to update
  ' informations about all resized controls
  Dim varControl As Variant
  
  If TypeOf objContainer Is Form Then
    intFormWidth = objContainer.ScaleWidth
    intFormHeight = objContainer.ScaleHeight
  Else
    intFormWidth = objContainer.Width
    intFormHeight = objContainer.Height
  End If
  For Each varControl In dicControls
    With objContainer.Controls(varControl)
      dicControls(varControl) = Array(dicControls(varControl)(AA_ANCHORS), .Left, .Top, .Width, .Height)
    End With
  Next
End Sub

Public Sub Resize()
  Dim OldControlX As Integer, OldControlY As Integer
  Dim OldControlWidth As Integer, OldControlHeight As Integer
  Dim NewWidth As Integer, NewHeight As Integer
  Dim NewControlX As Integer, NewControlY As Integer
  Dim NewControlWidth As Integer, NewControlHeight As Integer
  Dim varControl As Variant
  Dim blnDoIt As Boolean
  Dim strControlName As String
  Dim intControlIndex As Integer
  
  On Error Resume Next
  If TypeOf objContainer Is Form Then
    blnDoIt = (objContainer.WindowState <> 1)
  Else
    blnDoIt = True
  End If
  If blnDoIt Then
    If TypeOf objContainer Is Form Then
      NewWidth = objContainer.ScaleWidth
      NewHeight = objContainer.ScaleHeight
    Else
      NewWidth = objContainer.Width
      NewHeight = objContainer.Height
    End If
    For Each varControl In dicControls
      OldControlX = dicControls(varControl)(AA_LEFT)
      OldControlY = dicControls(varControl)(AA_TOP)
      OldControlWidth = dicControls(varControl)(AA_WIDTH)
      OldControlHeight = dicControls(varControl)(AA_HEIGHT)
      Select Case dicControls(varControl)(AA_ANCHORS)
      Case AN_NONE
        NewControlX = (OldControlX / (intFormWidth - OldControlWidth)) * (NewWidth - OldControlWidth)
        NewControlY = (OldControlY / (intFormHeight - OldControlHeight)) * (NewHeight - OldControlHeight)
        NewControlWidth = OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_TOP
        NewControlX = (OldControlX / (intFormWidth - OldControlWidth)) * (NewWidth - OldControlWidth)
        NewControlY = OldControlY
        NewControlWidth = OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_BOTTOM
        NewControlX = (OldControlX / (intFormWidth - OldControlWidth)) * (NewWidth - OldControlWidth)
        NewControlY = NewHeight - intFormHeight + OldControlY
        NewControlWidth = OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_TOP + AN_BOTTOM
        NewControlX = (OldControlX / (intFormWidth - OldControlWidth)) * (NewWidth - OldControlWidth)
        NewControlY = OldControlY
        NewControlWidth = OldControlWidth
        NewControlHeight = NewHeight - intFormHeight + OldControlHeight
      Case AN_LEFT
        NewControlX = OldControlX
        NewControlY = (OldControlY / (intFormHeight - OldControlHeight)) * (NewHeight - OldControlHeight)
        NewControlWidth = OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_TOP + AN_LEFT
        NewControlX = OldControlX
        NewControlY = OldControlY
        NewControlWidth = OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_BOTTOM + AN_LEFT
        NewControlX = OldControlX
        NewControlY = NewHeight - intFormHeight + OldControlY
        NewControlWidth = OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_TOP + AN_BOTTOM + AN_LEFT
        NewControlX = OldControlX
        NewControlY = OldControlY
        NewControlWidth = OldControlWidth
        NewControlHeight = NewHeight - intFormHeight + OldControlHeight
      Case AN_RIGHT
        NewControlX = NewWidth - intFormWidth + OldControlX
        NewControlY = (OldControlY / (intFormHeight - OldControlHeight)) * (NewHeight - OldControlHeight)
        NewControlWidth = OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_TOP + AN_RIGHT
        NewControlX = NewWidth - intFormWidth + OldControlX
        NewControlY = OldControlY
        NewControlWidth = OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_BOTTOM + AN_RIGHT
        NewControlX = NewWidth - intFormWidth + OldControlX
        NewControlY = NewHeight - intFormHeight + OldControlY
        NewControlWidth = OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_TOP + AN_BOTTOM + AN_RIGHT
        NewControlX = NewWidth - intFormWidth + OldControlX
        NewControlY = OldControlY
        NewControlWidth = OldControlWidth
        NewControlHeight = NewHeight - intFormHeight + OldControlHeight
      Case AN_LEFT + AN_RIGHT
        NewControlX = OldControlX
        NewControlY = (OldControlY / (intFormHeight - OldControlHeight)) * (NewHeight - OldControlHeight)
        NewControlWidth = NewWidth - intFormWidth + OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_TOP + AN_LEFT + AN_RIGHT
        NewControlX = OldControlX
        NewControlY = OldControlY
        NewControlWidth = NewWidth - intFormWidth + OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_BOTTOM + AN_LEFT + AN_RIGHT
        NewControlX = OldControlX
        NewControlY = NewHeight - intFormHeight + OldControlY
        NewControlWidth = NewWidth - intFormWidth + OldControlWidth
        NewControlHeight = OldControlHeight
      Case AN_TOP + AN_BOTTOM + AN_LEFT + AN_RIGHT
        NewControlX = OldControlX
        NewControlY = OldControlY
        NewControlWidth = NewWidth - intFormWidth + OldControlWidth
        NewControlHeight = NewHeight - intFormHeight + OldControlHeight
      End Select
      If InStr(varControl, "(") > 0 Then
        strControlName = Left(varControl, InStr(varControl, "(") - 1)
        intControlIndex = CInt(Replace(Mid(varControl, InStr(varControl, "(") + 1), ")", ""))
        objContainer.Controls(strControlName)(intControlIndex).Move NewControlX, NewControlY, IIf(NewControlWidth > 0, NewControlWidth, 0), IIf(NewControlHeight > 0, NewControlHeight, 0)
      Else
        objContainer.Controls(varControl).Move NewControlX, NewControlY, IIf(NewControlWidth > 0, NewControlWidth, 0), IIf(NewControlHeight > 0, NewControlHeight, 0)
      End If
    Next
  End If
End Sub
