VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmExporters 
   Caption         =   "Exporter Manager"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   Icon            =   "frmExporters.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   4665
   Begin VB.CommandButton cmdUninstall 
      Caption         =   "&Uninstall Exporter"
      Height          =   330
      Left            =   3015
      TabIndex        =   3
      ToolTipText     =   "Install a new Exporter."
      Top             =   765
      Width           =   1590
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4140
      Top             =   2565
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "&Install Exporter"
      Height          =   330
      Left            =   3015
      TabIndex        =   2
      ToolTipText     =   "Install a new Exporter."
      Top             =   405
      Width           =   1590
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   3015
      TabIndex        =   1
      ToolTipText     =   "Refresh the list of installed Exporters."
      Top             =   45
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   1770
      Left            =   0
      TabIndex        =   6
      Top             =   3015
      Width           =   4650
      Begin HighlightBox.HBX txtAuthor 
         Height          =   825
         Left            =   90
         TabIndex        =   9
         Top             =   855
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   1455
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Caption         =   "Author"
      End
      Begin VB.TextBox txtVersion 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   540
         Width           =   3525
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   3525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   540
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.ListBox lstExporters 
      Height          =   2985
      ItemData        =   "frmExporters.frx":030A
      Left            =   0
      List            =   "frmExporters.frx":030C
      TabIndex        =   0
      Top             =   0
      Width           =   2940
   End
End
Attribute VB_Name = "frmExporters"
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

Private Sub lstExporters_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXExporters
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmExporters, lstExporters_MouseUp"
End Sub

Public Sub cmdInstall_Click()
On Error GoTo Err_Handler
Dim Hdc As Long
  CommonDialog1.Flags = cdlOFNHideReadOnly
  CommonDialog1.Filter = "Plugins (*.dll)|*.dll|All Files (*.*)|*.*"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName = "" Then
    MsgBox "No Exporter selected - operation aborted!", vbExclamation, "Error"
    Exit Sub
  Else
    StartMsg "Registering Exporter..."
    LogMsg "Registering: " & CommonDialog1.FileName & "..."
    Hdc = GetDesktopWindow()
    ShellExecute Hdc, "Open", "regsvr32", " /s " & QUOTE & CommonDialog1.FileName & QUOTE, "C:\", SW_SHOWNORMAL
    EndMsg
  End If
  ListExporters
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmPlugins, cmdRefresh_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
  ListExporters
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPlugins, cmdRefresh_Click"
End Sub

Public Sub cmdUninstall_Click()
On Error GoTo Err_Handler
Dim x As Integer
Dim i As Long
Dim Res As String
Dim objTemp As pgExporter
Dim Hdc As Long
Dim CLSID As String
Dim DllFile As String
  If lstExporters.Text = "" Then
    MsgBox "You must select an Exporter to Uninstall!", vbExclamation, "Error"
    Exit Sub
  End If
  If MsgBox("Are you sure you wish to uninstall: " & lstExporters.Text & "?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
    StartMsg "Uninstalling: " & lstExporters.Text & "..."
    Res = RegGetSubkey(HKEY_CLASSES_ROOT, "", i)
    Do Until Res = "Not Found"
      If InStr(1, Res, "pgAdmin_Exporter") Then
        Set objTemp = CreateObject(Res)
        If objTemp.Description = lstExporters.Text Then
          CLSID = RegRead(HKEY_CLASSES_ROOT, Res & "\Clsid", "")
          DllFile = RegRead(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\InProcServer32", "")
          Hdc = GetDesktopWindow()
          LogMsg "Uninstalling Exporter: " & DllFile & "..."
          ShellExecute Hdc, "Open", "regsvr32", " /s /u " & QUOTE & DllFile & QUOTE, "C:\", SW_SHOWNORMAL
        End If
      End If
      i = i + 1
      Res = RegGetSubkey(HKEY_CLASSES_ROOT, "", i)
    Loop
    EndMsg
  End If
  ListExporters
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmPlugins, cmdUninstall_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4785
  Me.Height = 5220
  ListExporters
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPlugins, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtAuthor.Minimise
  If Me.WindowState = 0 Then
    Me.Width = 4785
    Me.Height = 5220
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPlugins, Form_Resize"
End Sub

Private Sub ListExporters()
On Error GoTo Err_Handler
Dim x As Integer
Dim i As Long
Dim Res As String
  StartMsg "Searching for installed Exporters..."
  lstExporters.Clear
  ReDim Exporters(0)
  Res = RegGetSubkey(HKEY_CLASSES_ROOT, "", i)
  Do Until Res = "Not Found"
    If InStr(1, Res, "pgAdmin_Exporter") Then
      Set Exporters(UBound(Exporters)) = CreateObject(Res)
      LogMsg "Loading Exporter: " & Res & " (" & Exporters(UBound(Exporters)).Description & " v" & Exporters(UBound(Exporters)).Version & ")"
      ReDim Preserve Exporters(UBound(Exporters) + 1)
Continue:
    End If
    i = i + 1
    Res = RegGetSubkey(HKEY_CLASSES_ROOT, "", i)
  Loop
  If UBound(Exporters) > 0 Then
    ReDim Preserve Exporters(UBound(Exporters) - 1)
    For x = 0 To UBound(Exporters)
      lstExporters.AddItem Exporters(x).Description
    Next
  End If
  txtDescription.Text = ""
  txtVersion.Text = ""
  txtAuthor.Text = ""
  EndMsg
  Exit Sub
Err_Handler:
  If Err.Number = -2147024770 Then
    LogMsg "Exporter: " & Res & " is registered but could not be found!"
    GoTo Continue
  ElseIf Err.Number = 13 Or Err.Number = 429 Then
    LogMsg "Exporter: " & Res & " is corrupt or invalid!"
    GoTo Continue
  Else
    EndMsg
    If Err.Number <> 0 Then LogError Err, "frmPlugins, ListExporters"
  End If
End Sub

Private Sub lstExporters_Click()
On Error GoTo Err_Handler
  txtDescription.Text = Exporters(lstExporters.ListIndex).Description
  txtVersion.Text = Exporters(lstExporters.ListIndex).Version
  txtAuthor.Text = Exporters(lstExporters.ListIndex).Author
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPlugins, ListExporters"
End Sub
