VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   5805
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   4410
      TabIndex        =   7
      ToolTipText     =   "Accept the changes."
      Top             =   4995
      Width           =   1320
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   4785
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   8440
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Logging/Tracking"
      TabPicture(0)   =   "frmOptions.frx":128A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CommonDialog1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTracking"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Loggin"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Syntax Highlight"
      TabPicture(1)   =   "frmOptions.frx":12A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Project"
      TabPicture(2)   =   "frmOptions.frx":12C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraMode"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Colour"
         Height          =   4065
         Left            =   -74820
         TabIndex        =   1
         Top             =   495
         Width           =   5235
         Begin VB.CheckBox chkItalic 
            Caption         =   "Italic"
            Height          =   285
            Left            =   990
            TabIndex        =   15
            ToolTipText     =   "Should the word be made italic?"
            Top             =   840
            Width           =   675
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "Bold"
            Height          =   285
            Left            =   135
            TabIndex        =   14
            ToolTipText     =   "Should the word be made bold?"
            Top             =   840
            Width           =   690
         End
         Begin VB.CommandButton cmdChangeCol 
            Caption         =   "Change Colour"
            Height          =   330
            Left            =   2295
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Select a colour for the word."
            Top             =   795
            Width           =   1800
         End
         Begin VB.CommandButton cmdAddNew 
            Caption         =   "Add"
            Height          =   345
            Left            =   4185
            TabIndex        =   13
            ToolTipText     =   "Add the selected word."
            Top             =   360
            Width           =   900
         End
         Begin VB.TextBox txtWord 
            Height          =   285
            Left            =   630
            TabIndex        =   12
            ToolTipText     =   "Enter a word to highlight."
            Top             =   390
            Width           =   3450
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   345
            Left            =   4185
            TabIndex        =   17
            ToolTipText     =   "Remove the selected word."
            Top             =   795
            Width           =   900
         End
         Begin MSComctlLib.ListView lvWords 
            Height          =   2625
            Left            =   90
            TabIndex        =   18
            ToolTipText     =   "Displays the Text Formatting rules."
            Top             =   1260
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   4630
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
            Left            =   135
            TabIndex        =   2
            Top             =   435
            Width           =   465
         End
      End
      Begin VB.Frame Loggin 
         Caption         =   "Logging"
         Height          =   1815
         Left            =   180
         TabIndex        =   3
         Top             =   495
         Width           =   5235
         Begin VB.CheckBox chkEnableLogging 
            Caption         =   "Enable &advanced logging"
            Height          =   195
            Left            =   135
            TabIndex        =   8
            Top             =   360
            Width           =   3570
         End
         Begin VB.CheckBox chkMaskPassword 
            Caption         =   "&Mask password in logfile"
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   630
            Width           =   3435
         End
         Begin VB.TextBox txtLogfile 
            Height          =   285
            Left            =   135
            TabIndex        =   10
            Top             =   1215
            Width           =   3165
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   330
            Left            =   3375
            TabIndex        =   11
            Top             =   1215
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Logfile"
            Height          =   195
            Left            =   135
            TabIndex        =   0
            Top             =   990
            Width           =   465
         End
      End
      Begin VB.Frame fraTracking 
         Caption         =   "Revision Tracking"
         Height          =   1185
         Left            =   180
         TabIndex        =   4
         Top             =   2430
         Width           =   5235
         Begin VB.TextBox txtTrackVer 
            Height          =   285
            Left            =   2565
            TabIndex        =   20
            ToolTipText     =   "Enter the last Revision Version Number."
            Top             =   630
            Width           =   915
         End
         Begin VB.CheckBox chkTracking 
            Alignment       =   1  'Right Justify
            Caption         =   "Revision Tracking enabled?"
            Height          =   240
            Left            =   180
            TabIndex        =   19
            ToolTipText     =   "Check to enable the Revision Tracking System."
            Top             =   360
            Width           =   2580
         End
         Begin VB.Label Label3 
            Caption         =   "Last Release Version"
            Height          =   195
            Left            =   225
            TabIndex        =   5
            Top             =   675
            Width           =   1545
         End
      End
      Begin VB.Frame fraMode 
         Caption         =   "Mode"
         Height          =   2220
         Left            =   -74820
         TabIndex        =   21
         Top             =   495
         Width           =   5235
         Begin VB.OptionButton optMode 
            Caption         =   "Production"
            Height          =   240
            Index           =   1
            Left            =   135
            TabIndex        =   24
            ToolTipText     =   "Select to put the Database in Production Mode."
            Top             =   315
            Width           =   4965
         End
         Begin VB.OptionButton optMode 
            Caption         =   "Development"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   25
            ToolTipText     =   "Select to put the Database in Development mode."
            Top             =   1260
            Width           =   4965
         End
         Begin VB.Label Label5 
            Caption         =   $"frmOptions.frx":12DE
            Height          =   645
            Left            =   360
            TabIndex        =   23
            Top             =   1530
            Width           =   4785
         End
         Begin VB.Label Label4 
            Caption         =   $"frmOptions.frx":1390
            Height          =   645
            Left            =   360
            TabIndex        =   22
            Top             =   585
            Width           =   4650
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   45
         Top             =   4230
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgadmin - PostgreSQL db Administration/Management for Win32
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




Private Sub chkTracking_Click()
On Error GoTo Err_Handler
  If chkTracking.Value = 1 Then
    txtTrackVer.Enabled = True
  Else
    txtTrackVer.Enabled = False
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmOptions, chkTracking_Click"
End Sub

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
  If chkBold = "1" Then
    itmX.SubItems(2) = "Y"
  Else
    itmX.SubItems(2) = "N"
  End If
  If chkItalic.Value = "1" Then
    itmX.SubItems(3) = "Y"
  Else
    itmX.SubItems(3) = "N"
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
  cmdChangeCol.BackColor = CommonDialog1.Color
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
  RegWrite HKEY_CURRENT_USER, "Software\pgadmin", "Logging", ValString, CStr(Logging)
  RegWrite HKEY_CURRENT_USER, "Software\pgadmin", "Mask Password", ValString, CStr(MaskPassword)
  RegWrite HKEY_CURRENT_USER, "Software\pgadmin", "Log File", ValString, CStr(LogFile)
  
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
    szTextColours = szTextColours & itmX.SubItems(1) & ";"
  Next itmX
  TextColours = szTextColours
  RegWrite HKEY_CURRENT_USER, "Software\pgadmin", "Text Colours", ValString, CStr(TextColours)
  
  'Save the Developer options...
  If optMode(0).Value = True Then
    DevMode = True
    fMainForm.StatusBar1.Panels("Mode").Text = "Development Mode"
    LogMsg "Executing: UPDATE pgadmin_param SET param_value = 'Y' WHERE param_id = 4"
    gConnection.Execute "UPDATE pgadmin_param SET param_value = 'Y' WHERE param_id = 4"
  Else
    DevMode = False
    fMainForm.StatusBar1.Panels("Mode").Text = "Production Mode"
    LogMsg "Executing: UPDATE pgadmin_param SET param_value = 'N' WHERE param_id = 4"
    gConnection.Execute "UPDATE pgadmin_param SET param_value = 'N' WHERE param_id = 4"
  End If
  If chkTracking.Value = 1 Then
    Tracking = True
    LogMsg "Executing: UPDATE pgadmin_param SET param_value = 'Y' WHERE param_id = 2"
    gConnection.Execute "UPDATE pgadmin_param SET param_value = 'Y' WHERE param_id = 2"
  Else
    Tracking = False
    LogMsg "Executing: UPDATE pgadmin_param SET param_value = 'N' WHERE param_id = 2"
    gConnection.Execute "UPDATE pgadmin_param SET param_value = 'N' WHERE param_id = 2"
  End If
  TrackVer = Val(txtTrackVer.Text)
  LogMsg "Executing: UPDATE pgadmin_param SET param_value = '" & TrackVer & "' WHERE param_id = 3"
  gConnection.Execute "UPDATE pgadmin_param SET param_value = '" & TrackVer & "' WHERE param_id = 3"
  
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
Dim x As Printer
Dim iLoop As Integer
Dim itmX As ListItem
Dim szStrings() As String
Dim szValues() As String
  Me.Width = 5900
  Me.Height = 5800
  txtWord.ForeColor = RGB(0, 0, 0)
  lvWords.ColumnHeaders.Add , , "Wordlist", (lvWords.Width / 11) * 4
  lvWords.ColumnHeaders.Add , , "Colour", (lvWords.Width / 11) * 3
  lvWords.ColumnHeaders.Add , , "B", (lvWords.Width / 11)
  lvWords.ColumnHeaders.Add , , "I", (lvWords.Width / 11)
  chkEnableLogging.Value = RegRead(HKEY_CURRENT_USER, "Software\pgadmin", "Logging", 0)
  chkMaskPassword.Value = RegRead(HKEY_CURRENT_USER, "Software\pgadmin", "Mask Password", 1)
  txtLogfile.Text = RegRead(HKEY_CURRENT_USER, "Software\pgadmin", "Log File", "C:\pgadmin.log")
  
  'Load the text colours into the grid.
  lvWords.ListItems.Clear
  szStrings = Split(TextColours, ";")
  For iLoop = 0 To UBound(szStrings) - 1
    szValues = Split(szStrings(iLoop), "|")
    Set itmX = lvWords.ListItems.Add(, , szValues(0))
    itmX.SubItems(1) = szValues(3)
    If szValues(2) = "1" Then
      itmX.SubItems(3) = "Y"
    Else
      itmX.SubItems(3) = "N"
    End If
    If szValues(1) = "1" Then
      itmX.SubItems(2) = "Y"
    Else
      itmX.SubItems(2) = "N"
    End If
  Next iLoop
  
  'Developer Options
  If DevMode = True Then
    optMode(0).Value = True
  Else
    optMode(1).Value = True
  End If
  If Tracking = True Then
    chkTracking.Value = 1
    txtTrackVer.Enabled = True
  Else
    chkTracking.Value = 0
    txtTrackVer.Enabled = False
  End If
  txtTrackVer.Text = TrackVer
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmOptions, Form_Load"
End Sub

Private Sub lvWords_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err_Handler
  CommonDialog1.Color = Val(Item.SubItems(1))
  cmdChangeCol.BackColor = CommonDialog1.Color
  txtWord.ForeColor = CommonDialog1.Color
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmOptions, lvWords_ItemClick"
End Sub

