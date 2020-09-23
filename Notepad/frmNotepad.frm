VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNotepad 
   ClientHeight    =   5280
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8760
   Icon            =   "frmNotepad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtMain 
      Height          =   5295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save as..."
      End
      Begin VB.Menu mnuDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuDiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuTimeDate 
         Caption         =   "Time/Date"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Format"
      Begin VB.Menu mnuFont 
         Caption         =   "Font..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About PVL Notepad"
      End
   End
End
Attribute VB_Name = "frmNotepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+---------------------------------------------------------------
'| PVL Resizer Samples Collection
'|
'| PVL Notepad
'| by Tomasz Puwalski (tpuwalski@op.pl)
'|
'| A basic notepad sample.
'|
'| If you feel that this code is useful and/or if you want
'| to use this code in your own software - please appreciate
'| my work and leave your vote on:
'| http://www.pscode.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=50708
'+---------------------------------------------------------------
Option Explicit

Private strFile As String
Private intDocNo As Integer
Private objResizer As clsResizer

Private Sub Form_Activate()
  txtMain.SetFocus
End Sub

Private Sub Form_Load()
  strFile = "Untitled"
  SetCaption
  
  '------------------
  ' Resizing - setup
  '------------------
  Set objResizer = New clsResizer
  With objResizer
    Set .Container = Me
    .SetAnchors "txtMain", avTop + avBottom + avLeft + avRight
  End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = Not SaveFileEx()
End Sub

Private Sub Form_Resize()
  '----------------------
  ' Resizing - main call
  '----------------------
  objResizer.Resize
End Sub

Private Sub mnuAbout_Click()
  MsgBox App.ProductName & " (Resizer sample)" & vbNewLine & _
         "Written by Tomasz Puwalski" & vbNewLine & vbNewLine & _
         "Vote me! ;-)", vbInformation, App.ProductName
End Sub

Private Sub mnuCopy_Click()
  Clipboard.Clear
  Clipboard.SetText txtMain.SelText
End Sub

Private Sub mnuCut_Click()
  Clipboard.Clear
  Clipboard.SetText txtMain.SelText
  txtMain.SelText = ""
End Sub

Private Sub mnuDelete_Click()
  txtMain.SelText = ""
End Sub

Private Sub mnuEdit_Click()
  mnuCopy.Enabled = (txtMain.SelText <> vbNullString)
  mnuCut.Enabled = (txtMain.SelText <> vbNullString)
  mnuDelete.Enabled = (txtMain.SelText <> vbNullString)
  mnuPaste.Enabled = (Clipboard.GetText <> vbNullString)
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuFont_Click()
  On Error Resume Next
  With CommonDialog1
    .FontName = txtMain.Font.Name
    .FontSize = txtMain.Font.Size
    .Color = txtMain.ForeColor
    .FontBold = txtMain.Font.Bold
    .FontItalic = txtMain.Font.Italic
    .FontUnderline = txtMain.Font.Underline
    .FontStrikethru = txtMain.Font.Strikethrough
    .Flags = cdlCFEffects Or cdlCFBoth
    .ShowFont
    txtMain.Font.Name = .FontName
    txtMain.Font.Size = .FontSize
    txtMain.ForeColor = .Color
    txtMain.Font.Bold = .FontBold
    txtMain.Font.Italic = .FontItalic
    txtMain.Font.Underline = .FontUnderline
    txtMain.Font.Strikethrough = .FontStrikethru
  End With
End Sub

Private Sub mnuNew_Click()
  If SaveFileEx() Then
    txtMain.Text = ""
    intDocNo = intDocNo + 1
    strFile = "Untitled" & CStr(intDocNo)
    SetCaption
  End If
End Sub

Private Sub mnuOpen_Click()
  Dim intHandle As Integer
  Dim strTemp As String
  
  On Error Resume Next
  If SaveFileEx() Then
    With CommonDialog1
      .DialogTitle = "Open"
      .CancelError = False
      .FileName = ""
      .Filter = "Text Documents (*.txt)|*.txt|All files|*.*"
      .ShowOpen
      If Len(.FileName) = 0 Then
        Exit Sub
      End If
      strFile = .FileName
    End With
    intHandle = FreeFile
    Open strFile For Binary As #intHandle
    strTemp = Space(LOF(intHandle))
    Get #intHandle, , strTemp
    Close #intHandle
    txtMain.Text = strTemp
    mnuSave.Enabled = False
    SetCaption
  End If
End Sub

Private Sub mnuPaste_Click()
  txtMain.SelText = Clipboard.GetText
End Sub

Private Sub mnuPrint_Click()
  On Error Resume Next
  With CommonDialog1
    .DialogTitle = "Print"
    .CancelError = True
    .Flags = cdlPDReturnDC + cdlPDNoPageNums
    If txtMain.SelLength = 0 Then
      .Flags = .Flags + cdlPDAllPages
    Else
      .Flags = .Flags + cdlPDSelection
    End If
    .ShowPrinter
    If Err <> MSComDlg.cdlCancel Then
      Printer.Print txtMain.Text
    End If
  End With
End Sub

Private Sub mnuSave_Click()
  SaveFile
End Sub

Private Sub mnuSaveAs_Click()
  SaveFile True
End Sub

Private Sub mnuSelectAll_Click()
  txtMain.SelStart = 0
  txtMain.SelLength = Len(txtMain.Text)
End Sub

Private Sub mnuTimeDate_Click()
  txtMain.SelText = Left(CStr(Time), 5) & " " & CStr(Date)
End Sub

Private Function JustFileName(ByVal fname As String) As String
  JustFileName = Mid(fname, Len(fname) - InStr(StrReverse(fname), "\") + 2)
End Function

Private Sub RichTextBox1_Change()
  txtMain.DataChanged = True
End Sub

Private Function SaveFile(Optional ByVal AsNew As Boolean) As Boolean
  Dim intHandle As Integer
  
  On Error Resume Next
  If Left(strFile, 8) = "Untitled" Or AsNew Then
    With CommonDialog1
      .DialogTitle = "Save"
      .CancelError = False
      .Filter = "Text Documents (*.txt)|*.txt|All files|*.*"
      .ShowSave
      If Len(.FileName) = 0 Then
        Exit Function
      End If
      strFile = .FileName
    End With
  End If
  intHandle = FreeFile
  Open strFile For Output As #intHandle
  Print #intHandle, txtMain.Text
  Close intHandle
  SetCaption
  SaveFile = True
End Function

Private Function SaveFileEx(Optional ByVal AsNew As Boolean) As Boolean
  Dim Selection As Integer
  
  If txtMain.DataChanged Then
    Selection = MsgBox("Text in the " & JustFileName(strFile) & " file has changed." & vbNewLine & vbNewLine & "Do you want to save the changes?", vbYesNoCancel + vbExclamation, App.ProductName)
    Select Case Selection
    Case vbYes
      SaveFileEx = SaveFile(AsNew)
    Case vbNo
      SaveFileEx = True
    End Select
  Else
    SaveFileEx = True
  End If
End Function

Private Sub SetCaption()
  Caption = IIf(InStr(strFile, "\") > 0, JustFileName(strFile), strFile) & " - " & App.ProductName
  txtMain.DataChanged = False
End Sub

Private Sub txtMain_Change()
  mnuSave.Enabled = True
End Sub
