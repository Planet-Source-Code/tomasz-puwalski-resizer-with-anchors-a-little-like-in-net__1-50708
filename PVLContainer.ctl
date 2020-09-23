VERSION 5.00
Begin VB.UserControl PVLContainer 
   BackColor       =   &H00EAF3F3&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "PVLContainer.ctx":0000
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00B6D4DA&
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "PVLContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------
' PVL Container
' by Tomasz Puwalski (pvl@cps.pl)
'---------------------------------------------

Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."

Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
  UserControl.BackColor = New_Color
  PropertyChanged
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = shpBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal New_Color As OLE_COLOR)
  shpBorder.BorderColor = New_Color
  PropertyChanged
End Property

Public Property Get Controls() As Collection
  Dim objControl As Object
  Dim strControlName As String
  
  On Error Resume Next
  Set Controls = New Collection
  For Each objControl In UserControl.ContainedControls
    strControlName = objControl.Name
    strControlName = strControlName & CStr(objControl.Index)
    Controls.Add objControl, strControlName
  Next
End Property

Private Sub UserControl_Resize()
  shpBorder.Width = UserControl.Width
  shpBorder.Height = UserControl.Height
  RaiseEvent Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "BackColor", BackColor, &H8000000F
  PropBag.WriteProperty "BorderColor", BorderColor, &H8000000F
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  BorderColor = PropBag.ReadProperty("BorderColor", &H8000000F)
End Sub


