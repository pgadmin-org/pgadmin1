VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   3330
      TabIndex        =   4
      ToolTipText     =   "Accept the changes."
      Top             =   2835
      Width           =   1320
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   2715
      Left            =   0
      TabIndex        =   6
      Top             =   45
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   4789
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Logging"
      TabPicture(0)   =   "frmOptions.frx":128A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkEnableLogging"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkMaskPassword"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtLogfile"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBrowse"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CommonDialog1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Colours"
      TabPicture(1)   =   "frmOptions.frx":12A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lvWords"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkItalic"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkBold"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkCase"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdChangeCol"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdAddNew"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtWord"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdRemove"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   345
         Left            =   -71355
         TabIndex        =   9
         ToolTipText     =   "Remove the selected word."
         Top             =   420
         Width           =   900
      End
      Begin VB.TextBox txtWord 
         Height          =   285
         Left            =   -74415
         TabIndex        =   7
         ToolTipText     =   "Enter a word to highlight."
         Top             =   450
         Width           =   2070
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "Add"
         Height          =   345
         Left            =   -72300
         TabIndex        =   8
         ToolTipText     =   "Add the selected word."
         Top             =   420
         Width           =   900
      End
      Begin VB.CommandButton cmdChangeCol 
         Caption         =   "Change Colour"
         Height          =   330
         Left            =   -71895
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Select a colour for the word."
         Top             =   855
         Width           =   1440
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "Case Sensitive"
         Height          =   285
         Left            =   -74910
         TabIndex        =   10
         ToolTipText     =   "Should the word be treated case sensitively."
         Top             =   878
         Width           =   1410
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "Bold"
         Height          =   285
         Left            =   -73470
         TabIndex        =   11
         ToolTipText     =   "Should the word be made bold?"
         Top             =   878
         Width           =   690
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "Italic"
         Height          =   285
         Left            =   -72705
         TabIndex        =   12
         ToolTipText     =   "Should the word be made italic?"
         Top             =   878
         Width           =   675
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   45
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   330
         Left            =   3690
         TabIndex        =   3
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtLogfile 
         Height          =   285
         Left            =   495
         TabIndex        =   2
         Top             =   1845
         Width           =   3165
      End
      Begin VB.CheckBox chkMaskPassword 
         Caption         =   "&Mask password in logfile"
         Height          =   195
         Left            =   495
         TabIndex        =   1
         Top             =   1260
         Width           =   3435
      End
      Begin VB.CheckBox chkEnableLogging 
         Caption         =   "&Enable advanced logging"
         Height          =   195
         Left            =   495
         TabIndex        =   0
         Top             =   810
         Width           =   3435
      End
      Begin MSComctlLib.ListView lvWords 
         Height          =   1365
         Left            =   -74955
         TabIndex        =   15
         ToolTipText     =   "Displays the Text Formatting rules."
         Top             =   1260
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   2408
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Word"
         Height          =   255
         Left            =   -74910
         TabIndex        =   14
         Top             =   495
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Logfile"
         Height          =   195
         Left            =   495
         TabIndex        =   5
         Top             =   1620
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin - PostgreSQL db Administration/Management for Win32
' Copyright (C) 1998 - 2001, Dave Page

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Option Explicit

Private Sub cmdAddNew_Click()
On Error GoTo Err_Handler
Dim itmX As ListItem
  If txtWord.Text = "" Then
    MsgBox "You must enter a word to add!", vbExclamation, "Error"
    txtWord.SetFocus
    Exit Sub
  End If
  For Each itmX In lvWords.ListItems
    If itmX.Text = txtWord.Text Then
      MsgBox "That word is already in the list!", vbExclamation, "Error"
      txtWord.SetFocus
      Exit Sub
    End If
  Next itmX

  'Add the new listitem
  Set itmX = lvWords.ListItems.Add(, , txtWord.Text)
  itmX.SubItems(1) = txtWord.ForeColor
  If chkCase.Value = "1" Then
    itmX.SubItems(2) = "Y"
  Else
    itmX.SubItems(2) = "N"
  End If
  If chkBold = "1" Then
    itmX.SubItems(3) = "Y"
  Else
    itmX.SubItems(3) = "N"
  End If
  If chkItalic.Value = "1" Then
    itmX.SubItems(4) = "Y"
  Else
    itmX.SubItems(4) = "N"
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmOptions, cmdAdd_Click"
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo Err_Handler
  With CommonDialog1
    .FileName = txtLogfile.Text
    .CancelError = True
    .Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    .Filter = "Log Files (*.log)|*.log"
    .ShowSave
  End With
  If CommonDialog1.FileName = "" Then Exit Sub
  txtLogfile.Text = CommonDialog1.FileName
  Exit Sub
Err_Handler: If Err.Number <> 0 And Err.Number <> 32755 Then LogError Err, "frmOptions, cmdBrowse_click"
End Sub

Private Sub cmdChangeCol_Click()
On Error GoTo Err_Handler
  CommonDialog1.ShowColor
  txtWord.ForeColor = CommonDialog1.Color
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmOptions, cmdChangeCol_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
Dim itmX As ListItem
Dim szTextColours As String
  Logging = chkEnableLogging.Value
  MaskPassword = chkMaskPassword.Value
  LogFile = txtLogfile.Text
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Logging", ValString, CStr(Logging)
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Mask Password", ValString, CStr(MaskPassword)
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Log File", ValString, CStr(LogFile)
  
  'Rebuild the Text Colour String
  For Each itmX In lvWords.ListItems
    szTextColours = szTextColours & itmX.Text & "|"
    If itmX.SubItems(2) = "Y" Then
      szTextColours = szTextColours & "1|"
    Else
      szTextColours = szTextColours & "0|"
    End If
    If itmX.SubItems(3) = "Y" Then
      szTextColours = szTextColours & "1|"
    Else
      szTextColours = szTextColours & "0|"
    End If
    If itmX.SubItems(4) = "Y" Then
      szTextColours = szTextColours & "1|"
    Else
      szTextColours = szTextColours & "0|"
    End If
    szTextColours = szTextColours & itmX.SubItems(1) & ";"
  Next itmX
  TextColours = szTextColours
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Text Colours", ValString, CStr(TextColours)
  
  'Unload the form.
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmOptions, cmdOK_Click"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
  If MsgBox("Are you sure you wish to remove the selected word?", vbQuestion + vbYesNo, "Remove Word") = vbNo Then Exit Sub
  lvWords.ListItems.Remove lvWords.SelectedItem.Index
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmOptions, cmdRemove_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim X As Printer
Dim iLoop As Integer
Dim itmX As ListItem
Dim szStrings() As String
Dim szValues() As String
  Me.Width = 4770
  Me.Height = 3570
  txtWord.ForeColor = RGB(0, 0, 0)
  lvWords.ColumnHeaders.Add , , "Wordlist", (lvWords.Width / 11) * 4
  lvWords.ColumnHeaders.Add , , "Colour", (lvWords.Width / 11) * 3
  lvWords.ColumnHeaders.Add , , "C", (lvWords.Width / 11)
  lvWords.ColumnHeaders.Add , , "B", (lvWords.Width / 11)
  lvWords.ColumnHeaders.Add , , "I", (lvWords.Width / 11)
  chkEnableLogging.Value = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Logging", 0)
  chkMaskPassword.Value = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Mask Password", 1)
  txtLogfile.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Log File", "C:\pgAdmin.log")
  
  'Load the text colours into the grid.
  lvWords.ListItems.Clear
  szStrings = Split(TextColours, ";")
  For iLoop = 0 To UBound(szStrings) - 1
    szValues = Split(szStrings(iLoop), "|")
    Set itmX = lvWords.ListItems.Add(, , szValues(0))
    itmX.SubItems(1) = szValues(4)
    If szValues(2) = "1" Then
      itmX.SubItems(2) = "Y"
    Else
      itmX.SubItems(2) = "N"
    End If
    If szValues(3) = "1" Then
      itmX.SubItems(4) = "Y"
    Else
      itmX.SubItems(4) = "N"
    End If
    If szValues(1) = "1" Then
      itmX.SubItems(3) = "Y"
    Else
      itmX.SubItems(3) = "N"
    End If
  Next iLoop
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrinter, Form_Load"
End Sub
