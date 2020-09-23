VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Contacts"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   Icon            =   "Contacts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAF3F3&
      Height          =   3075
      Left            =   2940
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1800
      Width           =   5535
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New"
      Height          =   345
      Left            =   5640
      TabIndex        =   7
      Top             =   4980
      Width           =   1335
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAF3F3&
      Height          =   5295
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2775
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   7080
      TabIndex        =   8
      Top             =   4980
      Width           =   1335
   End
   Begin Project1.PVLContainer Frame1 
      Height          =   1635
      Left            =   2940
      TabIndex        =   9
      Top             =   60
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2884
      BackColor       =   15397875
      BorderColor     =   -2147483642
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   645
         TabIndex        =   5
         Top             =   1200
         Width           =   4770
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         TabIndex        =   3
         Top             =   840
         Width           =   1995
      End
      Begin VB.TextBox txtMobile 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3420
         TabIndex        =   4
         Top             =   840
         Width           =   1995
      End
      Begin VB.TextBox txtLastname 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtFirstname 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label lblMobile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
         Height          =   195
         Left            =   2850
         TabIndex        =   13
         Top             =   885
         Width           =   510
      End
      Begin VB.Label lblPhone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   885
         Width           =   510
      End
      Begin VB.Label lblLastname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last name:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   525
         Width           =   780
      End
      Begin VB.Label lblFirstname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First name:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   165
         Width           =   765
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+---------------------------------------------------------------
'| PVL Resizer Samples Collection
'|
'| PVL Contacts
'| by Tomasz Puwalski (tpuwalski@op.pl)
'|
'| If you feel that this code is useful and/or if you want
'| to use this code in your own software - please appreciate
'| my work and leave your vote on:
'| http://www.pscode.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=50708
'+---------------------------------------------------------------
Option Explicit
  
Private objFormResizer As clsResizer
Private objFrameResizer As clsResizer
Private objDict As Collection
Private dbConn As New Connection
Private lngId As Long

Private Sub cmdAdd_Click()
  With Me
    .txtEmail.Text = ""
    .txtFirstname.Text = ""
    .txtLastname.Text = ""
    .txtMobile.Text = ""
    .txtNote.Text = ""
    .txtPhone.Text = ""
  End With
  lngId = 0
End Sub

Private Sub cmdSave_Click()
  Dim dbRS As Recordset

  If txtFirstname.Text = vbNullString Then
    MsgBox "Field ""First name"" is required", vbExclamation
    Exit Sub
  End If
  If txtLastname.Text = vbNullString Then
    MsgBox "Field ""Last name"" is required", vbExclamation
    Exit Sub
  End If
  Set dbRS = New ADODB.Recordset
  With dbRS
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    If lngId = 0 Then
      ' New
      .Open "main", dbConn, , , adCmdTable
      .AddNew
    Else
      ' Update
      .Open "SELECT * FROM main WHERE Id=" & lngId, dbConn, , , adCmdText
    End If
    .Fields("Email") = txtEmail.Text
    .Fields("Firstname") = txtFirstname.Text
    .Fields("Surname") = txtLastname.Text
    .Fields("Mobile") = txtMobile.Text
    .Fields("Note") = txtNote.Text
    .Fields("Phone") = txtPhone.Text
    .Update
    .Close
  End With
  Set dbRS = Nothing
  Call FillListBox
End Sub

Private Sub Form_Load()
  dbConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Contacts.mdb"
  Call FillListBox
  
  '------------------
  ' Resizing - setup
  '------------------
  Set objFormResizer = New clsResizer
  With objFormResizer
    Set .Container = Me
    .SetAnchors "lstNames", avTop + avBottom + avLeft
    .SetAnchors "Frame1", avTop + avLeft + avRight
    .SetAnchors "txtNote", avTop + avBottom + avLeft + avRight
    .SetAnchors "cmdAdd", avBottom + avRight
    .SetAnchors "cmdSave", avBottom + avRight
  End With
  
  Set objFrameResizer = New clsResizer
  With objFrameResizer
    Set .Container = Frame1
    .SetAnchors "txtFirstname", avTop + avLeft + avRight
    .SetAnchors "txtLastname", avTop + avLeft + avRight
    .SetAnchors "txtPhone", avTop + avLeft + avRight
    .SetAnchors "lblMobile", avTop + avRight
    .SetAnchors "txtMobile", avTop + avRight
    .SetAnchors "txtEmail", avTop + avLeft + avRight
  End With
End Sub

Private Sub Form_Resize()
  '----------------------
  ' Resizing - main call
  '----------------------
  objFormResizer.Resize
  objFrameResizer.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set dbConn = Nothing
End Sub

Private Sub lstNames_Click()
  Dim dbRS As Recordset
  
  On Error Resume Next
  Set dbRS = dbConn.Execute("SELECT * FROM main WHERE Id=" & objDict("k" & CStr(lstNames.ListIndex + 1)))
  With Me
    .txtEmail.Text = dbRS("Email")
    .txtFirstname.Text = dbRS("Firstname")
    .txtLastname.Text = dbRS("Surname")
    .txtMobile.Text = dbRS("Mobile")
    .txtNote.Text = dbRS("Note")
    .txtPhone.Text = dbRS("Phone")
  End With
  lngId = dbRS("Id")
  dbRS.Close
End Sub

Private Sub FillListBox()
  Dim dbRS As Recordset

  lstNames.Clear
  Set objDict = New Collection
  Set dbRS = dbConn.Execute("SELECT * FROM main ORDER BY Surname, Firstname")
  Do While Not dbRS.EOF
    lstNames.AddItem dbRS("Surname") & " " & dbRS("Firstname")
    objDict.Add (dbRS("Id")), "k" & CStr(lstNames.ListCount)
    dbRS.MoveNext
  Loop
  Set dbRS = Nothing
End Sub
